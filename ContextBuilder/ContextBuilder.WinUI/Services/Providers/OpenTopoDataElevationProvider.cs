using System.Globalization;
using System.Text;
using System.Text.Json;
using ContextBuilder.WinUI.Models;

namespace ContextBuilder.WinUI.Services.Providers;

public sealed class OpenTopoDataElevationProvider
{
    private readonly HttpClient _httpClient;

    public OpenTopoDataElevationProvider()
    {
        _httpClient = new HttpClient();
        _httpClient.DefaultRequestHeaders.UserAgent.ParseAdd("ContextBuilder.WinUI/0.1 (+local app)");
    }

    public async Task<List<(GeoPoint Point, double Elevation)>> SampleSquareGridAsync(
        GeoPoint center,
        double radiusMeters,
        int cellsPerSide,
        CancellationToken cancellationToken)
    {
        var points = BuildGrid(center, radiusMeters, cellsPerSide);
        var output = new List<(GeoPoint Point, double Elevation)>(points.Count);

        // Limit URL size for GET requests.
        const int batchSize = 70;
        for (var i = 0; i < points.Count; i += batchSize)
        {
            var batch = points.Skip(i).Take(batchSize).ToList();
            var locations = string.Join("|", batch.Select(p =>
                $"{p.Latitude.ToString("0.######", CultureInfo.InvariantCulture)},{p.Longitude.ToString("0.######", CultureInfo.InvariantCulture)}"));
            var url = $"https://api.opentopodata.org/v1/aster30m?locations={Uri.EscapeDataString(locations)}";

            using var response = await _httpClient.GetAsync(url, cancellationToken).ConfigureAwait(false);
            if (!response.IsSuccessStatusCode)
            {
                // If elevation service is unavailable, return flat terrain.
                output.AddRange(batch.Select(p => (p, 0d)));
                continue;
            }

            var json = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
            using var doc = JsonDocument.Parse(json);
            if (!doc.RootElement.TryGetProperty("results", out var results) || results.ValueKind != JsonValueKind.Array)
            {
                output.AddRange(batch.Select(p => (p, 0d)));
                continue;
            }

            var index = 0;
            foreach (var result in results.EnumerateArray())
            {
                var elevation = result.TryGetProperty("elevation", out var e) && e.ValueKind == JsonValueKind.Number
                    ? e.GetDouble()
                    : 0d;
                if (index < batch.Count)
                {
                    output.Add((batch[index], elevation));
                }

                index++;
            }

            while (index < batch.Count)
            {
                output.Add((batch[index], 0d));
                index++;
            }
        }

        return output;
    }

    private static List<GeoPoint> BuildGrid(GeoPoint center, double radiusMeters, int cellsPerSide)
    {
        var clampedCells = Math.Clamp(cellsPerSide, 8, 80);
        var points = new List<GeoPoint>(clampedCells * clampedCells);
        var size = radiusMeters * 2.0;
        var step = size / (clampedCells - 1);
        var half = size / 2.0;

        for (var y = 0; y < clampedCells; y++)
        {
            for (var x = 0; x < clampedCells; x++)
            {
                var dx = -half + x * step;
                var dy = -half + y * step;
                points.Add(GeoProjection.OffsetByMeters(center, dx, dy));
            }
        }

        return points;
    }
}
