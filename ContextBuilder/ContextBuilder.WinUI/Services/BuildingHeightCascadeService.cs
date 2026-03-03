using ContextBuilder.WinUI.Models;

namespace ContextBuilder.WinUI.Services;

public sealed class BuildingHeightCascadeService
{
    public void ApplyCascade(
        IReadOnlyCollection<LayerResult> layers,
        IReadOnlyList<(GeoPoint Point, double Elevation)>? elevationGrid)
    {
        var buildingLayer = layers.FirstOrDefault(l => l.Layer == ContextLayer.Buildings);
        if (buildingLayer is null || buildingLayer.BuildingFootprints.Count == 0)
        {
            return;
        }

        foreach (var footprint in buildingLayer.BuildingFootprints)
        {
            if (footprint.HeightMeters > 0 &&
                (string.Equals(footprint.HeightSource, "osm_height", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(footprint.HeightSource, "osm_levels", StringComparison.OrdinalIgnoreCase)))
            {
                footprint.HeightMeters = Math.Clamp(footprint.HeightMeters, 3, 320);
                continue;
            }

            if (TryEstimateFromElevation(footprint.Ring, elevationGrid, out var lidarHeight))
            {
                footprint.HeightMeters = lidarHeight;
                footprint.HeightSource = "lidar_derived";
                continue;
            }

            var ringMeters = footprint.Ring.Select(p => GeoProjection.ToLocalMeters(footprint.Ring[0], p)).ToList();
            footprint.HeightMeters = InferHeuristicHeight(ringMeters);
            footprint.HeightSource = "heuristic";
        }
    }

    private static bool TryEstimateFromElevation(
        IReadOnlyList<GeoPoint> ring,
        IReadOnlyList<(GeoPoint Point, double Elevation)>? elevationGrid,
        out double heightMeters)
    {
        heightMeters = 0;
        if (elevationGrid is null || elevationGrid.Count < 4 || ring.Count < 4)
        {
            return false;
        }

        var cleanRing = ring;
        if (ring.Count > 1 && IsSamePoint(ring[0], ring[^1]))
        {
            cleanRing = ring.Take(ring.Count - 1).ToList();
        }

        if (cleanRing.Count < 3)
        {
            return false;
        }

        var centroid = new GeoPoint(
            cleanRing.Average(p => p.Latitude),
            cleanRing.Average(p => p.Longitude));

        var maxRadius = 0d;
        foreach (var p in cleanRing)
        {
            var xy = GeoProjection.ToLocalMeters(centroid, p);
            var radius = Math.Sqrt((xy.X * xy.X) + (xy.Y * xy.Y));
            maxRadius = Math.Max(maxRadius, radius);
        }

        if (maxRadius <= 0.5)
        {
            return false;
        }

        var sampleRadius = Math.Max(6, maxRadius * 1.5);
        var buildingElevations = new List<double>();
        foreach (var sample in elevationGrid)
        {
            var xy = GeoProjection.ToLocalMeters(centroid, sample.Point);
            if ((xy.X * xy.X) + (xy.Y * xy.Y) > (sampleRadius * sampleRadius))
            {
                continue;
            }

            buildingElevations.Add(sample.Elevation);
        }

        if (buildingElevations.Count < 4)
        {
            return false;
        }

        buildingElevations.Sort();
        var p10 = Percentile(buildingElevations, 0.10);
        var p90 = Percentile(buildingElevations, 0.90);
        var spread = p90 - p10;
        if (spread < 2.0)
        {
            return false;
        }

        heightMeters = Math.Clamp(spread, 3.0, 180.0);
        return true;
    }

    private static double InferHeuristicHeight(IReadOnlyList<(double X, double Y)> ring)
    {
        var area = Math.Abs(SignedArea(ring));
        if (area < 25)
        {
            return 6;
        }

        var rough = 6 + Math.Sqrt(area) * 0.55;
        return Math.Clamp(rough, 8, 80);
    }

    private static double SignedArea(IReadOnlyList<(double X, double Y)> ring)
    {
        double area = 0;
        for (var i = 0; i < ring.Count; i++)
        {
            var j = (i + 1) % ring.Count;
            area += (ring[i].X * ring[j].Y) - (ring[j].X * ring[i].Y);
        }

        return area * 0.5;
    }

    private static bool IsSamePoint(GeoPoint a, GeoPoint b)
    {
        return Math.Abs(a.Latitude - b.Latitude) < 1e-7 &&
               Math.Abs(a.Longitude - b.Longitude) < 1e-7;
    }

    private static double Percentile(IReadOnlyList<double> sortedValues, double p)
    {
        if (sortedValues.Count == 0)
        {
            return 0;
        }

        var index = p * (sortedValues.Count - 1);
        var lo = (int)Math.Floor(index);
        var hi = (int)Math.Ceiling(index);
        if (lo == hi)
        {
            return sortedValues[lo];
        }

        var t = index - lo;
        return sortedValues[lo] + ((sortedValues[hi] - sortedValues[lo]) * t);
    }
}
