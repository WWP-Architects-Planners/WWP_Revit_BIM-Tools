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

        var query = BuildCombinedQuery(center, radiusMeters);
        var json = await ExecuteOverpassAsync(query, cancellationToken).ConfigureAwait(false);
        return ParseCombinedResults(targetLayers, json);
    }

    private async Task<string> ExecuteOverpassAsync(string query, CancellationToken cancellationToken)
    {
        var delayMs = 800;
        Exception? lastException = null;

        for (var attempt = 0; attempt < 6; attempt++)
        {
            var endpoint = Endpoints[attempt % Endpoints.Length];
            using var response = await _httpClient.PostAsync(
                endpoint,
                new StringContent(query, Encoding.UTF8, "application/x-www-form-urlencoded"),
                cancellationToken).ConfigureAwait(false);

            if (response.IsSuccessStatusCode)
            {
                return await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
            }

            if (response.StatusCode == HttpStatusCode.TooManyRequests || response.StatusCode == HttpStatusCode.ServiceUnavailable)
            {
                var retryAfter = response.Headers.RetryAfter?.Delta?.TotalMilliseconds;
                var waitMs = retryAfter.HasValue && retryAfter.Value > 0
                    ? (int)Math.Min(retryAfter.Value, 8000)
                    : delayMs;
                await Task.Delay(waitMs, cancellationToken).ConfigureAwait(false);
                delayMs = Math.Min(delayMs * 2, 8000);
                continue;
            }

            lastException = new HttpRequestException($"Overpass request failed ({(int)response.StatusCode} {response.ReasonPhrase})");
        }

        throw lastException ?? new HttpRequestException("Overpass request failed after retries (likely rate-limited).");
    }

    private static string BuildCombinedQuery(GeoPoint center, double radiusMeters)
    {
        var lat = center.Latitude.ToString("0.######", CultureInfo.InvariantCulture);
        var lon = center.Longitude.ToString("0.######", CultureInfo.InvariantCulture);
        var radius = Math.Max(50, radiusMeters).ToString("0", CultureInfo.InvariantCulture);
        var body =
            $"(" +
            $"way(around:{radius},{lat},{lon})[building];relation(around:{radius},{lat},{lon})[building];" +
            $"way(around:{radius},{lat},{lon})[highway];" +
            $"way(around:{radius},{lat},{lon})[natural=water];relation(around:{radius},{lat},{lon})[natural=water];way(around:{radius},{lat},{lon})[waterway];way(around:{radius},{lat},{lon})[landuse=reservoir];" +
            $"way(around:{radius},{lat},{lon})[leisure=park];relation(around:{radius},{lat},{lon})[leisure=park];way(around:{radius},{lat},{lon})[landuse=grass];way(around:{radius},{lat},{lon})[landuse=recreation_ground];" +
            $"way(around:{radius},{lat},{lon})[boundary=parcel];way(around:{radius},{lat},{lon})[cadastre];" +
            $");";
        return $"data=[out:json][timeout:40];{body}out body geom;";
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
                    var (heightMeters, heightSource) = ParseBuildingHeight(tags);
                    buildings.BuildingFootprints.Add(new BuildingFootprint
                    {
                        Ring = new List<GeoPoint>(points),
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
    private static bool IsBuilding(JsonElement tags) => HasTag(tags, "building");
    private static bool IsWater(JsonElement tags) =>
        TagEquals(tags, "natural", "water") ||
        HasTag(tags, "waterway") ||
        TagEquals(tags, "landuse", "reservoir");
    private static bool IsPark(JsonElement tags) =>
        TagEquals(tags, "leisure", "park") ||
        TagEquals(tags, "landuse", "grass") ||
        TagEquals(tags, "landuse", "recreation_ground");
    private static bool IsParcel(JsonElement tags) =>
        TagEquals(tags, "boundary", "parcel") ||
        HasTag(tags, "cadastre");

    private static (double HeightMeters, string Source) ParseBuildingHeight(JsonElement tags)
    {
        var fromHeight = ParseLengthMeters(GetTagString(tags, "height"));
        if (fromHeight > 0)
        {
            return (fromHeight, "osm_height");
        }

        var levelsText = GetTagString(tags, "building:levels");
        if (double.TryParse(levelsText, NumberStyles.Float, CultureInfo.InvariantCulture, out var levels) ||
            double.TryParse(levelsText, out levels))
        {
            return (Math.Clamp(levels * 3.2, 3.0, 180.0), "osm_levels");
        }

        return (0, "unknown");
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
}
