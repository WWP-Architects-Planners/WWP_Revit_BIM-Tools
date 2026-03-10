using System.Globalization;
using System.Text;
using ContextBuilder.WinUI.Models;

namespace ContextBuilder.WinUI.Services;

public sealed class ObjExporter
{
    public string Export(
        string folderPath,
        GeoPoint center,
        IReadOnlyCollection<LayerResult> layers,
        IReadOnlyList<(GeoPoint Point, double Elevation)>? elevationGrid,
        bool useElevation)
    {
        Directory.CreateDirectory(folderPath);
        var outputPath = Path.Combine(folderPath, "context_map.obj");

        var sb = new StringBuilder();
        var vertexIndex = 1;

        sb.AppendLine("# ContextBuilder OBJ export");
        sb.AppendLine("# Units: meters");

        if (useElevation && elevationGrid is { Count: > 0 })
        {
            vertexIndex = AppendTerrainFromGrid(sb, center, elevationGrid, vertexIndex);
        }
        else
        {
            vertexIndex = AppendFlatTerrain(sb, 800, vertexIndex);
        }

        foreach (var layer in layers)
        {
            switch (layer.Layer)
            {
                case ContextLayer.Buildings:
                    vertexIndex = AppendBuildingMeshes(sb, center, layer, vertexIndex);
                    break;
                case ContextLayer.Roads:
                    vertexIndex = AppendRoadSurfaces(sb, center, layer, vertexIndex);
                    break;
                case ContextLayer.Water:
                case ContextLayer.Parks:
                case ContextLayer.Parcels:
                    vertexIndex = AppendSurfaceOrLineLayer(sb, center, layer, vertexIndex);
                    break;
            }
        }

        File.WriteAllText(outputPath, sb.ToString(), Encoding.UTF8);
        return outputPath;
    }

    private static int AppendFlatTerrain(StringBuilder sb, double halfSize, int vertexIndex)
    {
        sb.AppendLine("g terrain_flat");
        var v1 = vertexIndex++;
        var v2 = vertexIndex++;
        var v3 = vertexIndex++;
        var v4 = vertexIndex++;
        sb.AppendLine(Vertex(-halfSize, -halfSize, 0));
        sb.AppendLine(Vertex(halfSize, -halfSize, 0));
        sb.AppendLine(Vertex(halfSize, halfSize, 0));
        sb.AppendLine(Vertex(-halfSize, halfSize, 0));
        sb.AppendLine($"f {v1} {v2} {v3}");
        sb.AppendLine($"f {v1} {v3} {v4}");
        return vertexIndex;
    }

    private static int AppendTerrainFromGrid(
        StringBuilder sb,
        GeoPoint center,
        IReadOnlyList<(GeoPoint Point, double Elevation)> grid,
        int vertexIndex)
    {
        sb.AppendLine("g terrain_dem");
        var estimatedSide = (int)Math.Round(Math.Sqrt(grid.Count));
        if (estimatedSide * estimatedSide != grid.Count || estimatedSide < 2)
        {
            return AppendFlatTerrain(sb, 800, vertexIndex);
        }

        var sorted = grid.Select(x => x.Elevation).OrderBy(x => x).ToList();
        var median = sorted[sorted.Count / 2];
        var indexMap = new int[estimatedSide, estimatedSide];
        for (var i = 0; i < grid.Count; i++)
        {
            var row = i / estimatedSide;
            var col = i % estimatedSide;
            var xy = GeoProjection.ToLocalMeters(center, grid[i].Point);
            var normalized = grid[i].Elevation - median;
            // Clamp noisy outliers from free DEM APIs to avoid spike artifacts.
            var z = Math.Clamp(normalized, -60, 120);
            indexMap[row, col] = vertexIndex;
            sb.AppendLine(Vertex(xy.X, xy.Y, z));
            vertexIndex++;
        }

        for (var row = 0; row < estimatedSide - 1; row++)
        {
            for (var col = 0; col < estimatedSide - 1; col++)
            {
                var a = indexMap[row, col];
                var b = indexMap[row, col + 1];
                var c = indexMap[row + 1, col + 1];
                var d = indexMap[row + 1, col];
                sb.AppendLine($"f {a} {b} {c}");
                sb.AppendLine($"f {a} {c} {d}");
            }
        }

        return vertexIndex;
    }

    private static int AppendBuildingMeshes(StringBuilder sb, GeoPoint center, LayerResult buildingLayer, int vertexIndex)
    {
        sb.AppendLine("g buildings");
        var footprints = buildingLayer.BuildingFootprints.Count > 0
            ? buildingLayer.BuildingFootprints
            : buildingLayer.Polygons.Select(p => new BuildingFootprint { Ring = p, BaseHeightMeters = 0, HeightMeters = 0, HeightSource = "heuristic" }).ToList();

        foreach (var footprint in footprints)
        {
            var ring = footprint.Ring.Select(p => GeoProjection.ToLocalMeters(center, p)).ToList();
            ring = CleanRing(ring);
            if (ring.Count < 3)
            {
                continue;
            }

            var baseHeight = Math.Max(0, footprint.BaseHeightMeters);
            var height = footprint.HeightMeters > 0
                ? Math.Clamp(footprint.HeightMeters, 3, 320)
                : InferBuildingHeightFromArea(ring);
            var topHeight = baseHeight + height;

            var bottom = new List<int>(ring.Count);
            var top = new List<int>(ring.Count);

            foreach (var p in ring)
            {
                bottom.Add(vertexIndex);
                sb.AppendLine(Vertex(p.X, p.Y, baseHeight));
                vertexIndex++;
            }

            foreach (var p in ring)
            {
                top.Add(vertexIndex);
                sb.AppendLine(Vertex(p.X, p.Y, topHeight));
                vertexIndex++;
            }

            for (var i = 0; i < ring.Count; i++)
            {
                var next = (i + 1) % ring.Count;
                var b1 = bottom[i];
                var b2 = bottom[next];
                var t2 = top[next];
                var t1 = top[i];
                sb.AppendLine($"f {b1} {b2} {t2}");
                sb.AppendLine($"f {b1} {t2} {t1}");
            }

            var triangles = TriangulatePolygon(ring);
            foreach (var tri in triangles)
            {
                sb.AppendLine($"f {top[tri.A]} {top[tri.B]} {top[tri.C]}");
            }
        }

        return vertexIndex;
    }

    private static int AppendRoadSurfaces(StringBuilder sb, GeoPoint center, LayerResult layer, int vertexIndex)
    {
        sb.AppendLine("g roads_surface");
        const double z = 0.12;
        const double halfWidth = 4.0;

        foreach (var line in layer.LineStrings)
        {
            var points = line.Select(p => GeoProjection.ToLocalMeters(center, p)).ToList();
            for (var i = 0; i < points.Count - 1; i++)
            {
                var p0 = points[i];
                var p1 = points[i + 1];
                var dx = p1.X - p0.X;
                var dy = p1.Y - p0.Y;
                var len = Math.Sqrt(dx * dx + dy * dy);
                if (len < 0.2)
                {
                    continue;
                }

                var nx = -dy / len;
                var ny = dx / len;
                var a = (X: p0.X + nx * halfWidth, Y: p0.Y + ny * halfWidth);
                var b = (X: p0.X - nx * halfWidth, Y: p0.Y - ny * halfWidth);
                var c = (X: p1.X - nx * halfWidth, Y: p1.Y - ny * halfWidth);
                var d = (X: p1.X + nx * halfWidth, Y: p1.Y + ny * halfWidth);

                var ia = vertexIndex++;
                var ib = vertexIndex++;
                var ic = vertexIndex++;
                var id = vertexIndex++;
                sb.AppendLine(Vertex(a.X, a.Y, z));
                sb.AppendLine(Vertex(b.X, b.Y, z));
                sb.AppendLine(Vertex(c.X, c.Y, z));
                sb.AppendLine(Vertex(d.X, d.Y, z));
                sb.AppendLine($"f {ia} {ib} {ic}");
                sb.AppendLine($"f {ia} {ic} {id}");
            }
        }

        // Fill polygonal road-like data if present.
        foreach (var polygon in layer.Polygons)
        {
            var ring = CleanRing(polygon.Select(p => GeoProjection.ToLocalMeters(center, p)).ToList());
            if (ring.Count < 3)
            {
                continue;
            }

            var indices = new List<int>();
            foreach (var p in ring)
            {
                indices.Add(vertexIndex);
                sb.AppendLine(Vertex(p.X, p.Y, z));
                vertexIndex++;
            }

            foreach (var tri in TriangulatePolygon(ring))
            {
                sb.AppendLine($"f {indices[tri.A]} {indices[tri.B]} {indices[tri.C]}");
            }
        }

        return vertexIndex;
    }

    private static int AppendSurfaceOrLineLayer(StringBuilder sb, GeoPoint center, LayerResult layer, int vertexIndex)
    {
        sb.AppendLine($"g {layer.Layer.ToString().ToLowerInvariant()}");
        var z = layer.Layer switch
        {
            ContextLayer.Water => -0.1,
            ContextLayer.Parks => 0.05,
            ContextLayer.Parcels => 0.15,
            _ => 0
        };

        foreach (var polygon in layer.Polygons)
        {
            var ring = CleanRing(polygon.Select(p => GeoProjection.ToLocalMeters(center, p)).ToList());
            if (ring.Count < 3)
            {
                continue;
            }

            var idx = new List<int>();
            foreach (var p in ring)
            {
                idx.Add(vertexIndex);
                sb.AppendLine(Vertex(p.X, p.Y, z));
                vertexIndex++;
            }

            foreach (var tri in TriangulatePolygon(ring))
            {
                sb.AppendLine($"f {idx[tri.A]} {idx[tri.B]} {idx[tri.C]}");
            }
        }

        foreach (var line in layer.LineStrings)
        {
            var indices = new List<int>();
            foreach (var point in line)
            {
                var xy = GeoProjection.ToLocalMeters(center, point);
                indices.Add(vertexIndex);
                sb.AppendLine(Vertex(xy.X, xy.Y, z));
                vertexIndex++;
            }

            if (indices.Count > 1)
            {
                sb.AppendLine($"l {string.Join(' ', indices)}");
            }
        }

        return vertexIndex;
    }

    private static List<(double X, double Y)> CleanRing(List<(double X, double Y)> ring)
    {
        if (ring.Count == 0)
        {
            return ring;
        }

        if (DistanceSq(ring[0], ring[^1]) < 1e-6)
        {
            ring.RemoveAt(ring.Count - 1);
        }

        var cleaned = new List<(double X, double Y)>();
        foreach (var p in ring)
        {
            if (cleaned.Count == 0 || DistanceSq(cleaned[^1], p) > 1e-6)
            {
                cleaned.Add(p);
            }
        }

        return cleaned;
    }

    private static double DistanceSq((double X, double Y) a, (double X, double Y) b)
    {
        var dx = a.X - b.X;
        var dy = a.Y - b.Y;
        return dx * dx + dy * dy;
    }

    private static double InferBuildingHeightFromArea(List<(double X, double Y)> ring)
    {
        var area = Math.Abs(SignedArea(ring));
        if (area < 25)
        {
            return 6;
        }

        var rough = 6 + Math.Sqrt(area) * 0.55;
        return Math.Clamp(rough, 8, 80);
    }

    private static double SignedArea(List<(double X, double Y)> ring)
    {
        double a = 0;
        for (var i = 0; i < ring.Count; i++)
        {
            var j = (i + 1) % ring.Count;
            a += (ring[i].X * ring[j].Y) - (ring[j].X * ring[i].Y);
        }

        return a * 0.5;
    }

    private static List<(int A, int B, int C)> TriangulatePolygon(List<(double X, double Y)> ring)
    {
        var triangles = new List<(int A, int B, int C)>();
        if (ring.Count < 3)
        {
            return triangles;
        }

        var area = SignedArea(ring);
        var vertexOrder = area > 0
            ? Enumerable.Range(0, ring.Count).ToList()
            : Enumerable.Range(0, ring.Count).Reverse().ToList();

        var guard = 0;
        while (vertexOrder.Count > 3 && guard < 10000)
        {
            var earFound = false;
            for (var i = 0; i < vertexOrder.Count; i++)
            {
                var prev = vertexOrder[(i - 1 + vertexOrder.Count) % vertexOrder.Count];
                var curr = vertexOrder[i];
                var next = vertexOrder[(i + 1) % vertexOrder.Count];

                if (!IsConvex(ring[prev], ring[curr], ring[next]))
                {
                    continue;
                }

                var containsOther = false;
                for (var j = 0; j < vertexOrder.Count; j++)
                {
                    var candidate = vertexOrder[j];
                    if (candidate == prev || candidate == curr || candidate == next)
                    {
                        continue;
                    }

                    if (PointInTriangle(ring[candidate], ring[prev], ring[curr], ring[next]))
                    {
                        containsOther = true;
                        break;
                    }
                }

                if (containsOther)
                {
                    continue;
                }

                triangles.Add((prev, curr, next));
                vertexOrder.RemoveAt(i);
                earFound = true;
                break;
            }

            if (!earFound)
            {
                break;
            }

            guard++;
        }

        if (vertexOrder.Count == 3)
        {
            triangles.Add((vertexOrder[0], vertexOrder[1], vertexOrder[2]));
        }
        else if (triangles.Count == 0)
        {
            // Fallback for malformed polygons.
            for (var i = 1; i < ring.Count - 1; i++)
            {
                triangles.Add((0, i, i + 1));
            }
        }

        return triangles;
    }

    private static bool IsConvex((double X, double Y) a, (double X, double Y) b, (double X, double Y) c)
    {
        var cross = (b.X - a.X) * (c.Y - a.Y) - (b.Y - a.Y) * (c.X - a.X);
        return cross > 1e-9;
    }

    private static bool PointInTriangle(
        (double X, double Y) p,
        (double X, double Y) a,
        (double X, double Y) b,
        (double X, double Y) c)
    {
        var s1 = Sign(p, a, b);
        var s2 = Sign(p, b, c);
        var s3 = Sign(p, c, a);
        var hasNeg = s1 < -1e-9 || s2 < -1e-9 || s3 < -1e-9;
        var hasPos = s1 > 1e-9 || s2 > 1e-9 || s3 > 1e-9;
        return !(hasNeg && hasPos);
    }

    private static double Sign((double X, double Y) p1, (double X, double Y) p2, (double X, double Y) p3)
    {
        return (p1.X - p3.X) * (p2.Y - p3.Y) - (p2.X - p3.X) * (p1.Y - p3.Y);
    }

    private static string Vertex(double x, double y, double z)
    {
        return $"v {x.ToString("0.###", CultureInfo.InvariantCulture)} {y.ToString("0.###", CultureInfo.InvariantCulture)} {z.ToString("0.###", CultureInfo.InvariantCulture)}";
    }
}
