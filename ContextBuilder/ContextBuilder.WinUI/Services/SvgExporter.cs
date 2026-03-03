using System.Globalization;
using System.Text;
using ContextBuilder.WinUI.Models;

namespace ContextBuilder.WinUI.Services;

public sealed class SvgExporter
{
    public List<string> Export(
        string folderPath,
        GeoPoint center,
        string sourceEpsg,
        string targetEpsg,
        IReadOnlyCollection<LayerResult> layers,
        IReadOnlyList<(GeoPoint Point, double Elevation)>? elevationGrid)
    {
        Directory.CreateDirectory(folderPath);
        var outputFiles = new List<string>();

        foreach (var layer in layers)
        {
            var path = Path.Combine(folderPath, $"{layer.Layer.ToString().ToLowerInvariant()}.svg");
            var svg = BuildLayerSvg(layer, sourceEpsg, targetEpsg);
            File.WriteAllText(path, svg, Encoding.UTF8);
            outputFiles.Add(path);
        }

        if (elevationGrid is { Count: > 0 })
        {
            var path = Path.Combine(folderPath, "lidar_toposurface.svg");
            var svg = BuildElevationSvg(sourceEpsg, targetEpsg, elevationGrid);
            File.WriteAllText(path, svg, Encoding.UTF8);
            outputFiles.Add(path);
        }

        return outputFiles;
    }

    private static string BuildLayerSvg(LayerResult layer, string sourceEpsg, string targetEpsg)
    {
        var width = 2000d;
        var height = 2000d;
        var margin = 60d;

        var allPoints = new List<(double X, double Y)>();
        allPoints.AddRange(layer.LineStrings.SelectMany(line => line.Select(p => ProjectPoint(p, sourceEpsg, targetEpsg))));
        allPoints.AddRange(layer.Polygons.SelectMany(poly => poly.Select(p => ProjectPoint(p, sourceEpsg, targetEpsg))));

        if (allPoints.Count == 0)
        {
            return EmptySvg(width, height, layer.Layer.ToString());
        }

        var minX = allPoints.Min(p => p.X);
        var maxX = allPoints.Max(p => p.X);
        var minY = allPoints.Min(p => p.Y);
        var maxY = allPoints.Max(p => p.Y);
        var spanX = Math.Max(1e-6, maxX - minX);
        var spanY = Math.Max(1e-6, maxY - minY);
        var scale = Math.Min((width - margin * 2) / spanX, (height - margin * 2) / spanY);

        var color = LayerColor(layer.Layer);
        var sb = new StringBuilder();
        sb.AppendLine($"<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"{width.ToString("0", CultureInfo.InvariantCulture)}\" height=\"{height.ToString("0", CultureInfo.InvariantCulture)}\" viewBox=\"0 0 {width.ToString("0", CultureInfo.InvariantCulture)} {height.ToString("0", CultureInfo.InvariantCulture)}\">");
        sb.AppendLine($"  <desc>Layer {layer.Layer} exported from {sourceEpsg} to {targetEpsg}</desc>");
        sb.AppendLine("  <rect x=\"0\" y=\"0\" width=\"100%\" height=\"100%\" fill=\"white\"/>");

        foreach (var line in layer.LineStrings)
        {
            var points = line.Select(p => ProjectPoint(p, sourceEpsg, targetEpsg)).Select(p => ToSvgPoint(p, minX, minY, scale, margin, height, spanY));
            sb.AppendLine($"  <polyline points=\"{string.Join(" ", points)}\" fill=\"none\" stroke=\"{color}\" stroke-width=\"2\" />");
        }

        foreach (var polygon in layer.Polygons)
        {
            var points = polygon.Select(p => ProjectPoint(p, sourceEpsg, targetEpsg)).Select(p => ToSvgPoint(p, minX, minY, scale, margin, height, spanY));
            sb.AppendLine($"  <polygon points=\"{string.Join(" ", points)}\" fill=\"{color}33\" stroke=\"{color}\" stroke-width=\"1.5\" />");
        }

        sb.AppendLine("</svg>");
        return sb.ToString();
    }

    private static string BuildElevationSvg(
        string sourceEpsg,
        string targetEpsg,
        IReadOnlyList<(GeoPoint Point, double Elevation)> grid)
    {
        var width = 2000d;
        var height = 2000d;
        var margin = 60d;
        var projected = grid.Select(x => (Projected: ProjectPoint(x.Point, sourceEpsg, targetEpsg), x.Elevation)).ToList();

        var minX = projected.Min(p => p.Projected.X);
        var maxX = projected.Max(p => p.Projected.X);
        var minY = projected.Min(p => p.Projected.Y);
        var maxY = projected.Max(p => p.Projected.Y);
        var spanX = Math.Max(1e-6, maxX - minX);
        var spanY = Math.Max(1e-6, maxY - minY);
        var scale = Math.Min((width - margin * 2) / spanX, (height - margin * 2) / spanY);

        var minZ = projected.Min(p => p.Elevation);
        var maxZ = projected.Max(p => p.Elevation);
        var spanZ = Math.Max(0.1, maxZ - minZ);

        var sb = new StringBuilder();
        sb.AppendLine($"<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"{width.ToString("0", CultureInfo.InvariantCulture)}\" height=\"{height.ToString("0", CultureInfo.InvariantCulture)}\" viewBox=\"0 0 {width.ToString("0", CultureInfo.InvariantCulture)} {height.ToString("0", CultureInfo.InvariantCulture)}\">");
        sb.AppendLine($"  <desc>LiDAR/DEM points exported from {sourceEpsg} to {targetEpsg}</desc>");
        sb.AppendLine("  <rect x=\"0\" y=\"0\" width=\"100%\" height=\"100%\" fill=\"white\"/>");

        foreach (var point in projected)
        {
            var xy = ToSvgPointTuple(point.Projected, minX, minY, scale, margin, height, spanY);
            var t = (point.Elevation - minZ) / spanZ;
            var color = t switch
            {
                < 0.2 => "#2B83BA",
                < 0.4 => "#80BFAB",
                < 0.6 => "#FFFFBF",
                < 0.8 => "#FDAE61",
                _ => "#D7191C"
            };
            sb.AppendLine($"  <circle cx=\"{xy.X.ToString("0.##", CultureInfo.InvariantCulture)}\" cy=\"{xy.Y.ToString("0.##", CultureInfo.InvariantCulture)}\" r=\"2.2\" fill=\"{color}\" />");
        }

        sb.AppendLine("</svg>");
        return sb.ToString();
    }

    private static (double X, double Y) ProjectPoint(GeoPoint point, string sourceEpsg, string targetEpsg)
    {
        return GeoProjection.ProjectForTarget(point, sourceEpsg, targetEpsg);
    }

    private static string ToSvgPoint((double X, double Y) value, double minX, double minY, double scale, double margin, double height, double spanY)
    {
        var t = ToSvgPointTuple(value, minX, minY, scale, margin, height, spanY);
        return $"{t.X.ToString("0.##", CultureInfo.InvariantCulture)},{t.Y.ToString("0.##", CultureInfo.InvariantCulture)}";
    }

    private static (double X, double Y) ToSvgPointTuple((double X, double Y) value, double minX, double minY, double scale, double margin, double height, double spanY)
    {
        var x = margin + (value.X - minX) * scale;
        var y = height - margin - (value.Y - minY) * scale;
        return (x, y);
    }

    private static string LayerColor(ContextLayer layer) => layer switch
    {
        ContextLayer.Buildings => "#333333",
        ContextLayer.Roads => "#C28D2C",
        ContextLayer.Water => "#2D7DD2",
        ContextLayer.Parks => "#3CA650",
        ContextLayer.Parcels => "#7E57C2",
        _ => "#111111"
    };

    private static string EmptySvg(double width, double height, string layerName)
    {
        return
            $"<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"{width}\" height=\"{height}\"><rect width=\"100%\" height=\"100%\" fill=\"white\"/><text x=\"20\" y=\"40\" font-size=\"20\">No {layerName} geometry found.</text></svg>";
    }
}
