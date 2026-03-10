using System.Globalization;
using System.Text.Json;
using ContextBuilder.WinUI.Models;

namespace ContextBuilder.WinUI.Services;

public sealed class GeoJsonExporter
{
    public List<string> Export(
        string folderPath,
        string sourceEpsg,
        string targetEpsg,
        IReadOnlyCollection<LayerResult> layers,
        IReadOnlyList<(GeoPoint Point, double Elevation)>? elevationGrid)
    {
        Directory.CreateDirectory(folderPath);
        var files = new List<string>();

        foreach (var layer in layers)
        {
            var path = Path.Combine(folderPath, $"{layer.Layer.ToString().ToLowerInvariant()}.geojson");
            var fc = BuildLayerFeatureCollection(layer, sourceEpsg, targetEpsg);
            File.WriteAllText(path, JsonSerializer.Serialize(fc, PrettyJson), System.Text.Encoding.UTF8);
            files.Add(path);
        }

        if (elevationGrid is { Count: > 0 })
        {
            var path = Path.Combine(folderPath, "lidar_toposurface_points.geojson");
            var fc = BuildElevationFeatureCollection(elevationGrid, sourceEpsg, targetEpsg);
            File.WriteAllText(path, JsonSerializer.Serialize(fc, PrettyJson), System.Text.Encoding.UTF8);
            files.Add(path);
        }

        return files;
    }

    private static object BuildLayerFeatureCollection(LayerResult layer, string sourceEpsg, string targetEpsg)
    {
        var features = new List<object>();

        foreach (var line in layer.LineStrings)
        {
            var coords = line.Select(p => ToCoordinatePair(p, sourceEpsg, targetEpsg)).ToList();
            features.Add(new
            {
                type = "Feature",
                properties = new { layer = layer.Layer.ToString(), geomType = "LineString" },
                geometry = new { type = "LineString", coordinates = coords }
            });
        }

        for (var i = 0; i < layer.Polygons.Count; i++)
        {
            var polygon = layer.Polygons[i];
            var coords = polygon.Select(p => ToCoordinatePair(p, sourceEpsg, targetEpsg)).ToList();
            if (!CoordinatesMatch(coords.FirstOrDefault(), coords.LastOrDefault()) && coords.Count > 2)
            {
                coords.Add(coords[0]);
            }

            double? heightM = null;
            double? baseHeightM = null;
            string? heightSource = null;
            if (layer.Layer == ContextLayer.Buildings && i < layer.BuildingFootprints.Count)
            {
                baseHeightM = Math.Round(layer.BuildingFootprints[i].BaseHeightMeters, 3);
                heightM = Math.Round(layer.BuildingFootprints[i].HeightMeters, 3);
                heightSource = layer.BuildingFootprints[i].HeightSource;
            }

            features.Add(new
            {
                type = "Feature",
                properties = new
                {
                    layer = layer.Layer.ToString(),
                    geomType = "Polygon",
                    base_height_m = baseHeightM,
                    height_m = heightM,
                    height_source = heightSource
                },
                geometry = new { type = "Polygon", coordinates = new[] { coords } }
            });
        }

        return new
        {
            type = "FeatureCollection",
            name = layer.Layer.ToString(),
            crs = new { type = "name", properties = new { name = targetEpsg } },
            features
        };
    }

    private static object BuildElevationFeatureCollection(
        IReadOnlyList<(GeoPoint Point, double Elevation)> elevationGrid,
        string sourceEpsg,
        string targetEpsg)
    {
        var features = elevationGrid.Select(p => new
        {
            type = "Feature",
            properties = new { elevation = p.Elevation.ToString("0.###", CultureInfo.InvariantCulture) },
            geometry = new { type = "Point", coordinates = ToCoordinatePair(p.Point, sourceEpsg, targetEpsg) }
        }).ToList();

        return new
        {
            type = "FeatureCollection",
            name = "LiDARTopoPoints",
            crs = new { type = "name", properties = new { name = targetEpsg } },
            features
        };
    }

    private static double[] ToCoordinatePair(GeoPoint point, string sourceEpsg, string targetEpsg)
    {
        var xy = GeoProjection.ProjectForTarget(point, sourceEpsg, targetEpsg);
        return
        [
            Math.Round(xy.X, 6),
            Math.Round(xy.Y, 6)
        ];
    }

    private static bool CoordinatesMatch(double[]? a, double[]? b)
    {
        if (a is null || b is null || a.Length < 2 || b.Length < 2)
        {
            return false;
        }

        return Math.Abs(a[0] - b[0]) < 1e-7 && Math.Abs(a[1] - b[1]) < 1e-7;
    }

    private static readonly JsonSerializerOptions PrettyJson = new()
    {
        WriteIndented = true
    };
}
