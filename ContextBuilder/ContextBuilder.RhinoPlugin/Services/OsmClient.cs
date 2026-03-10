using System.Globalization;
using System.Net;
using System.Text;
using System.Text.Json;
using ContextBuilder.RhinoPlugin.Models;

namespace ContextBuilder.RhinoPlugin.Services;

public sealed class OsmClient
{
    private static readonly HttpClient Http = new();
    private static readonly string[] OverpassEndpoints =
    [
        "https://overpass-api.de/api/interpreter",
        "https://overpass.kumi.systems/api/interpreter",
        "https://overpass.private.coffee/api/interpreter"
    ];

    public OsmClient()
    {
        if (!Http.DefaultRequestHeaders.UserAgent.Any())
        {
            Http.DefaultRequestHeaders.UserAgent.ParseAdd("ContextBuilder.RhinoPlugin/0.1");
        }
    }

    public async Task<GeoPoint?> GeocodeAsync(string address, CancellationToken ct)
    {
        var url = $"https://nominatim.openstreetmap.org/search?format=json&limit=1&q={Uri.EscapeDataString(address)}";
        using var response = await Http.GetAsync(url, ct).ConfigureAwait(false);
        response.EnsureSuccessStatusCode();
        var json = await response.Content.ReadAsStringAsync(ct).ConfigureAwait(false);
        using var doc = JsonDocument.Parse(json);
        if (doc.RootElement.ValueKind != JsonValueKind.Array || doc.RootElement.GetArrayLength() == 0)
        {
            return null;
        }

        var first = doc.RootElement[0];
        if (!first.TryGetProperty("lat", out var latEl) || !first.TryGetProperty("lon", out var lonEl))
        {
            return null;
        }

        if (!double.TryParse(latEl.GetString(), NumberStyles.Float, CultureInfo.InvariantCulture, out var lat))
        {
            return null;
        }
        if (!double.TryParse(lonEl.GetString(), NumberStyles.Float, CultureInfo.InvariantCulture, out var lon))
        {
            return null;
        }

        return new GeoPoint(lat, lon);
    }

    public async Task<(List<OsmElement> Elements, string Endpoint, string Mode)> FetchLayersAsync(
        GeoPoint center,
        double radiusMeters,
        IReadOnlyCollection<string> layers,
        Action<string> log,
        CancellationToken ct)
    {
        var primary = BuildOverpassQuery(center, radiusMeters, layers, includeRelations: true);
        var fallback = BuildOverpassQuery(center, radiusMeters, layers, includeRelations: false);

        Exception? last = null;
        foreach (var endpoint in OverpassEndpoints)
        {
            try
            {
                log($"Overpass request: {endpoint}");
                var primaryJson = await ExecuteOverpassAsync(endpoint, primary, ct).ConfigureAwait(false);
                return (ParseElements(primaryJson), endpoint, "primary");
            }
            catch (HttpRequestException ex)
            {
                log($"Primary failed at {endpoint}: {ex.Message}");
                last = ex;
            }

            try
            {
                log($"Retrying fallback query at {endpoint}");
                var fallbackJson = await ExecuteOverpassAsync(endpoint, fallback, ct).ConfigureAwait(false);
                return (ParseElements(fallbackJson), endpoint, "fallback");
            }
            catch (Exception ex)
            {
                log($"Fallback failed at {endpoint}: {ex.Message}");
                last = ex;
            }
        }

        throw last ?? new InvalidOperationException("Overpass request failed.");
    }

    private static async Task<string> ExecuteOverpassAsync(string endpoint, string query, CancellationToken ct)
    {
        using var content = new StringContent($"data={Uri.EscapeDataString(query)}", Encoding.UTF8, "application/x-www-form-urlencoded");
        using var response = await Http.PostAsync(endpoint, content, ct).ConfigureAwait(false);
        if (response.StatusCode == HttpStatusCode.TooManyRequests || response.StatusCode == HttpStatusCode.GatewayTimeout)
        {
            throw new HttpRequestException($"HTTP {(int)response.StatusCode}");
        }

        if (!response.IsSuccessStatusCode)
        {
            var text = await response.Content.ReadAsStringAsync(ct).ConfigureAwait(false);
            var snippet = text.Length > 180 ? text[..180] : text;
            throw new HttpRequestException($"HTTP {(int)response.StatusCode}: {snippet}");
        }

        return await response.Content.ReadAsStringAsync(ct).ConfigureAwait(false);
    }

    private static string BuildOverpassQuery(GeoPoint center, double radiusMeters, IReadOnlyCollection<string> layers, bool includeRelations)
    {
        var halfSideMeters = Math.Max(50, radiusMeters);
        var (south, west, north, east) = CalculateSquareBounds(center, halfSideMeters);
        var southStr = south.ToString("0.######", CultureInfo.InvariantCulture);
        var westStr = west.ToString("0.######", CultureInfo.InvariantCulture);
        var northStr = north.ToString("0.######", CultureInfo.InvariantCulture);
        var eastStr = east.ToString("0.######", CultureInfo.InvariantCulture);

        var selectors = new List<string>();

        if (layers.Contains("buildings"))
        {
            selectors.Add($"way({southStr},{westStr},{northStr},{eastStr})[building]");
            selectors.Add($"way({southStr},{westStr},{northStr},{eastStr})[\"building:part\"]");
            if (includeRelations)
            {
                selectors.Add($"relation({southStr},{westStr},{northStr},{eastStr})[building]");
                selectors.Add($"relation({southStr},{westStr},{northStr},{eastStr})[\"building:part\"]");
            }
        }

        if (layers.Contains("roads"))
        {
            selectors.Add($"way({southStr},{westStr},{northStr},{eastStr})[highway]");
        }

        if (layers.Contains("water"))
        {
            selectors.Add($"way({southStr},{westStr},{northStr},{eastStr})[natural=water]");
            selectors.Add($"way({southStr},{westStr},{northStr},{eastStr})[waterway]");
            if (includeRelations)
            {
                selectors.Add($"relation({southStr},{westStr},{northStr},{eastStr})[natural=water]");
            }
        }

        if (layers.Contains("parks"))
        {
            selectors.Add($"way({southStr},{westStr},{northStr},{eastStr})[leisure=park]");
            selectors.Add($"way({southStr},{westStr},{northStr},{eastStr})[landuse=grass]");
            if (includeRelations)
            {
                selectors.Add($"relation({southStr},{westStr},{northStr},{eastStr})[leisure=park]");
            }
        }

        if (layers.Contains("parcels"))
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

    private static List<OsmElement> ParseElements(string json)
    {
        var output = new List<OsmElement>();
        using var doc = JsonDocument.Parse(json);
        if (!doc.RootElement.TryGetProperty("elements", out var elements) || elements.ValueKind != JsonValueKind.Array)
        {
            return output;
        }

        foreach (var element in elements.EnumerateArray())
        {
            if (!element.TryGetProperty("geometry", out var geometry) || geometry.ValueKind != JsonValueKind.Array)
            {
                continue;
            }

            var points = new List<GeoPoint>();
            foreach (var node in geometry.EnumerateArray())
            {
                if (!node.TryGetProperty("lat", out var lat) || !node.TryGetProperty("lon", out var lon))
                {
                    continue;
                }

                points.Add(new GeoPoint(lat.GetDouble(), lon.GetDouble()));
            }

            if (points.Count < 2)
            {
                continue;
            }

            var tags = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (element.TryGetProperty("tags", out var tagsEl) && tagsEl.ValueKind == JsonValueKind.Object)
            {
                foreach (var p in tagsEl.EnumerateObject())
                {
                    tags[p.Name] = p.Value.GetString() ?? string.Empty;
                }
            }

            output.Add(new OsmElement { Geometry = points, Tags = tags });
        }

        return output;
    }
}
