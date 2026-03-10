using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Net;
using ContextBuilder.WinUI.Models;

namespace ContextBuilder.WinUI.Services.Providers;

public sealed class OsmOverpassProvider
{
    private readonly HttpClient _httpClient;
    private static readonly string[] Endpoints =
    [
        "https://overpass-api.de/api/interpreter",
        "https://overpass.kumi.systems/api/interpreter",
        "https://overpass.private.coffee/api/interpreter"
    ];

    public OsmOverpassProvider()
    {
        _httpClient = new HttpClient();
        _httpClient.DefaultRequestHeaders.UserAgent.ParseAdd("ContextBuilder.WinUI/0.1 (+local app)");
    }

    public async Task<List<LayerResult>> GetLayersAsync(
        GeoPoint center,
        double radiusMeters,
        IReadOnlyCollection<ContextLayer> layers,
        CancellationToken cancellationToken)
    {
        var targetLayers = layers
            .Where(l => l != ContextLayer.LidarTopoSurface)
            .Distinct()
            .ToList();
        if (targetLayers.Count == 0)
        {
            return [];
        }

        var query = BuildCombinedQuery(center, radiusMeters, targetLayers, includeRelations: true);
        var fallbackQuery = BuildCombinedQuery(center, radiusMeters, targetLayers, includeRelations: false);
        string? json = null;
        Exception? lastException = null;

        foreach (var endpoint in Endpoints)
        {
            try
            {
                json = await ExecuteOverpassAsync(endpoint, query, cancellationToken).ConfigureAwait(false);
                break;
            }
            catch (Exception ex)
            {
                lastException = ex;
            }

            try
            {
                json = await ExecuteOverpassAsync(endpoint, fallbackQuery, cancellationToken).ConfigureAwait(false);
                break;
            }
            catch (Exception ex)
            {
                lastException = ex;
            }
        }

        if (json is null)
        {
            throw lastException ?? new HttpRequestException("Overpass request failed.");
        }

        return ParseCombinedResults(targetLayers, json);
    }

    private async Task<string> ExecuteOverpassAsync(string endpoint, string query, CancellationToken cancellationToken)
    {
        using var response = await _httpClient.PostAsync(
            endpoint,
            new StringContent($"data={Uri.EscapeDataString(query)}", Encoding.UTF8, "application/x-www-form-urlencoded"),
            cancellationToken).ConfigureAwait(false);

        if (response.StatusCode == HttpStatusCode.TooManyRequests ||
            response.StatusCode == HttpStatusCode.ServiceUnavailable ||
            response.StatusCode == HttpStatusCode.GatewayTimeout)
        {
            throw new HttpRequestException($"HTTP {(int)response.StatusCode}");
        }

        if (!response.IsSuccessStatusCode)
        {
            var text = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
            var snippet = text.Length > 180 ? text[..180] : text;
            throw new HttpRequestException($"HTTP {(int)response.StatusCode}: {snippet}");
        }

        return await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
    }

    private static string BuildCombinedQuery(
        GeoPoint center,
        double radiusMeters,
        IReadOnlyCollection<ContextLayer> layers,
        bool includeRelations)
    {
        var halfSideMeters = Math.Max(50, radiusMeters);
        var (south, west, north, east) = CalculateSquareBounds(center, halfSideMeters);
        var southStr = south.ToString("0.######", CultureInfo.InvariantCulture);
        var westStr = west.ToString("0.######", CultureInfo.InvariantCulture);
        var northStr = north.ToString("0.######", CultureInfo.InvariantCulture);
        var eastStr = east.ToString("0.######", CultureInfo.InvariantCulture);

        var selectors = new List<string>();

        if (layers.Contains(ContextLayer.Buildings))
        {
            selectors.Add($"way({southStr},{westStr},{northStr},{eastStr})[building]");
            selectors.Add($"way({southStr},{westStr},{northStr},{eastStr})[\"building:part\"]");
            if (includeRelations)
            {
                selectors.Add($"relation({southStr},{westStr},{northStr},{eastStr})[building]");
                selectors.Add($"relation({southStr},{westStr},{northStr},{eastStr})[\"building:part\"]");
            }
        }

        if (layers.Contains(ContextLayer.Roads))
        {
            selectors.Add($"way({southStr},{westStr},{northStr},{eastStr})[highway]");
        }

        if (layers.Contains(ContextLayer.Water))
        {
            selectors.Add($"way({southStr},{westStr},{northStr},{eastStr})[natural=water]");
            selectors.Add($"way({southStr},{westStr},{northStr},{eastStr})[waterway]");
            if (includeRelations)
            {
                selectors.Add($"relation({southStr},{westStr},{northStr},{eastStr})[natural=water]");
            }
        }

        if (layers.Contains(ContextLayer.Parks))
        {
            selectors.Add($"way({southStr},{westStr},{northStr},{eastStr})[leisure=park]");
            selectors.Add($"way({southStr},{westStr},{northStr},{eastStr})[landuse=grass]");
            selectors.Add($"way({southStr},{westStr},{northStr},{eastStr})[landuse=recreation_ground]");
            if (includeRelations)
            {
                selectors.Add($"relation({southStr},{westStr},{northStr},{eastStr})[leisure=park]");
            }
        }

        if (layers.Contains(ContextLayer.Parcels))
        {
            selectors.Add($"way({southStr},{westStr},{northStr},{eastStr})[boundary=parcel]");
            selectors.Add($"way({southStr},{westStr},{northStr},{eastStr})[boundary=lot]");
            selectors.Add($"way({southStr},{westStr},{northStr},{eastStr})[boundary=plot]");
            selectors.Add($"way({southStr},{westStr},{northStr},{eastStr})[cadastre]");
            selectors.Add($"way({southStr},{westStr},{northStr},{eastStr})[cadastral]");
        }

        if (selectors.Count == 0)
        {
            return "[out:json][timeout:40];();out body geom;";
        }

        var body = string.Join(';', selectors);
        return $"[out:json][timeout:40];({body};);out body geom;";
    }

    private static List<LayerResult> ParseCombinedResults(IReadOnlyCollection<ContextLayer> selectedLayers, string json)
    {
        var output = selectedLayers.ToDictionary(k => k, v => new LayerResult { Layer = v });
        using var doc = JsonDocument.Parse(json);
        if (!doc.RootElement.TryGetProperty("elements", out var elements) || elements.ValueKind != JsonValueKind.Array)
        {
            return output.Values.ToList();
        }

        foreach (var element in elements.EnumerateArray())
        {
            if (!element.TryGetProperty("geometry", out var geometry) || geometry.ValueKind != JsonValueKind.Array)
            {
                continue;
            }

            var tags = element.TryGetProperty("tags", out var t) && t.ValueKind == JsonValueKind.Object ? t : default;
            var points = ReadPoints(geometry);
            if (points.Count < 2)
            {
                continue;
            }

            var isClosed = IsClosed(points);
            var isRoad = IsRoad(tags);
            var isBuilding = IsBuilding(tags);
            var isWater = IsWater(tags);
            var isPark = IsPark(tags);
            var isParcel = IsParcel(tags);

            if (isRoad && output.TryGetValue(ContextLayer.Roads, out var roads))
            {
                roads.LineStrings.Add(points);
            }

            if (isBuilding && output.TryGetValue(ContextLayer.Buildings, out var buildings))
            {
                AddPolyOrLine(buildings, points, isClosed);
                if (isClosed && points.Count >= 4)
                {
                    var (baseHeightMeters, heightMeters, heightSource) = ParseBuildingHeight(tags);
                    buildings.BuildingFootprints.Add(new BuildingFootprint
                    {
                        Ring = new List<GeoPoint>(points),
                        BaseHeightMeters = baseHeightMeters,
                        HeightMeters = heightMeters,
                        HeightSource = heightSource
                    });
                }
            }

            if (isWater && output.TryGetValue(ContextLayer.Water, out var waters))
            {
                AddPolyOrLine(waters, points, isClosed);
            }

            if (isPark && output.TryGetValue(ContextLayer.Parks, out var parks))
            {
                AddPolyOrLine(parks, points, isClosed);
            }

            if (isParcel && output.TryGetValue(ContextLayer.Parcels, out var parcels))
            {
                AddPolyOrLine(parcels, points, isClosed);
            }
        }

        return output.Values.ToList();
    }

    private static List<GeoPoint> ReadPoints(JsonElement geometry)
    {
        var points = new List<GeoPoint>();
        foreach (var node in geometry.EnumerateArray())
        {
            if (!node.TryGetProperty("lat", out var latElement) || !node.TryGetProperty("lon", out var lonElement))
            {
                continue;
            }

            points.Add(new GeoPoint(latElement.GetDouble(), lonElement.GetDouble()));
        }

        return points;
    }

    private static void AddPolyOrLine(LayerResult target, List<GeoPoint> points, bool isClosed)
    {
        if (isClosed && points.Count >= 4)
        {
            EnsureClosed(points);
            target.Polygons.Add(points);
        }
        else
        {
            target.LineStrings.Add(points);
        }
    }

    private static bool IsRoad(JsonElement tags) => HasTag(tags, "highway");
    private static bool IsBuilding(JsonElement tags) => HasTag(tags, "building") || HasTag(tags, "building:part");
    private static bool IsWater(JsonElement tags) =>
        TagEquals(tags, "natural", "water") ||
        HasTag(tags, "waterway") ||
        TagEquals(tags, "landuse", "reservoir");
    private static bool IsPark(JsonElement tags) =>
        TagEquals(tags, "leisure", "park") ||
        TagEquals(tags, "landuse", "grass") ||
        TagEquals(tags, "landuse", "recreation_ground");
    private static bool IsParcel(JsonElement tags)
    {
        var boundary = (GetTagString(tags, "boundary") ?? string.Empty).ToLowerInvariant();
        var landuse = (GetTagString(tags, "landuse") ?? string.Empty).ToLowerInvariant();

        if (boundary is "parcel" or "lot" or "plot" or "cadastral")
        {
            return true;
        }

        if (HasTag(tags, "cadastre") || HasTag(tags, "cadastral"))
        {
            return true;
        }

        return landuse is "residential" or "commercial" or "industrial" or "retail" or "allotments" or "garages";
    }

    private static (double BaseHeightMeters, double HeightMeters, string Source) ParseBuildingHeight(JsonElement tags)
    {
        var baseZ = ParseLengthMeters(GetTagString(tags, "min_height"));
        var minLevels = ParseLengthMeters(GetTagString(tags, "min_level"));
        if (baseZ <= 0 && minLevels > 0)
        {
            baseZ = minLevels * 3.2;
        }

        var topZ = ParseLengthMeters(GetTagString(tags, "height"));
        var levels = ParseLengthMeters(GetTagString(tags, "building:levels"));
        if (topZ <= 0 && levels > 0)
        {
            topZ = levels * 3.2;
        }

        var source = "unknown";
        if (topZ > 0)
        {
            source = GetTagString(tags, "height") is not null ? "osm_height" : "osm_levels";
        }

        if (topZ <= baseZ + 0.1)
        {
            topZ = baseZ + 3.0;
        }

        var height = topZ > 0 ? Math.Clamp(topZ - baseZ, 3.0, 320.0) : 0;
        return (Math.Max(0, baseZ), height, source);
    }

    private static string? GetTagString(JsonElement tags, string key)
    {
        if (tags.ValueKind != JsonValueKind.Object || !tags.TryGetProperty(key, out var v))
        {
            return null;
        }

        return v.GetString();
    }

    private static double ParseLengthMeters(string? input)
    {
        if (string.IsNullOrWhiteSpace(input))
        {
            return 0;
        }

        var text = input.Trim().ToLowerInvariant();
        if (text.EndsWith("ft"))
        {
            var feetValue = ParseLeadingNumber(text);
            return feetValue > 0 ? feetValue * 0.3048 : 0;
        }

        if (text.EndsWith("m"))
        {
            var mValue = ParseLeadingNumber(text);
            return mValue > 0 ? mValue : 0;
        }

        var direct = ParseLeadingNumber(text);
        return direct > 0 ? direct : 0;
    }

    private static double ParseLeadingNumber(string text)
    {
        var cleaned = new string(text
            .TakeWhile(c => char.IsDigit(c) || c == '.' || c == ',' || c == '-' || c == '+')
            .ToArray())
            .Replace(',', '.');

        if (double.TryParse(cleaned, NumberStyles.Float, CultureInfo.InvariantCulture, out var value))
        {
            return value;
        }

        return 0;
    }

    private static bool HasTag(JsonElement tags, string key)
    {
        return tags.ValueKind == JsonValueKind.Object && tags.TryGetProperty(key, out _);
    }

    private static bool TagEquals(JsonElement tags, string key, string value)
    {
        if (tags.ValueKind != JsonValueKind.Object || !tags.TryGetProperty(key, out var v))
        {
            return false;
        }

        return string.Equals(v.GetString(), value, StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsClosed(List<GeoPoint> points)
    {
        if (points.Count < 3)
        {
            return false;
        }

        var first = points[0];
        var last = points[^1];
        return Math.Abs(first.Latitude - last.Latitude) < 1e-7 &&
               Math.Abs(first.Longitude - last.Longitude) < 1e-7;
    }

    private static void EnsureClosed(List<GeoPoint> points)
    {
        if (points.Count < 3)
        {
            return;
        }

        if (!points[0].Equals(points[^1]))
        {
            points.Add(points[0]);
        }
    }

    private static (double South, double West, double North, double East) CalculateSquareBounds(GeoPoint center, double halfSideMeters)
    {
        var latRad = center.Latitude * Math.PI / 180.0;
        var metersPerDegreeLat = 111_320d;
        var metersPerDegreeLon = Math.Max(1e-9, metersPerDegreeLat * Math.Cos(latRad));

        var dLat = halfSideMeters / metersPerDegreeLat;
        var dLon = halfSideMeters / metersPerDegreeLon;

        return (
            center.Latitude - dLat,
            center.Longitude - dLon,
            center.Latitude + dLat,
            center.Longitude + dLon
        );
    }
}
