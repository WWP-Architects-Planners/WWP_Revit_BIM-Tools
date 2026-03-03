using System.Globalization;
using ContextBuilder.RhinoPlugin.Models;
using Rhino;
using Rhino.DocObjects;
using Rhino.Geometry;

namespace ContextBuilder.RhinoPlugin.Services;

public sealed class GeometryBuilder
{
    private const double EarthRadiusM = 6378137.0;

    public (double X, double Y) ToLocal(GeoPoint center, GeoPoint p)
    {
        var latR = DegreesToRadians(p.Latitude);
        var lonR = DegreesToRadians(p.Longitude);
        var cLatR = DegreesToRadians(center.Latitude);
        var cLonR = DegreesToRadians(center.Longitude);
        return ((lonR - cLonR) * EarthRadiusM * Math.Cos(cLatR), (latR - cLatR) * EarthRadiusM);
    }

    public Dictionary<string, int> EnsureLayers(RhinoDoc doc)
    {
        return new Dictionary<string, int>
        {
            ["buildings"] = EnsureLayer(doc, "ContextBuilder_Buildings", System.Drawing.Color.Gainsboro),
            ["roads"] = EnsureLayer(doc, "ContextBuilder_Roads", System.Drawing.Color.Goldenrod),
            ["roads_center"] = EnsureLayer(doc, "ContextBuilder_RoadCenterlines", System.Drawing.Color.Yellow),
            ["water"] = EnsureLayer(doc, "ContextBuilder_Water", System.Drawing.Color.SteelBlue),
            ["parks"] = EnsureLayer(doc, "ContextBuilder_Parks", System.Drawing.Color.DarkSeaGreen),
            ["parcels"] = EnsureLayer(doc, "ContextBuilder_Parcels", System.Drawing.Color.DarkGray),
        };
    }

    public static double DefaultRoadWidth(string highway, IReadOnlyDictionary<string, double> settings)
    {
        var h = (highway ?? string.Empty).ToLowerInvariant();
        if (h is "motorway" or "motorway_link" or "trunk" or "trunk_link") return settings["motorway"];
        if (h is "primary" or "primary_link" or "secondary" or "secondary_link") return settings["primary"];
        if (h is "tertiary" or "tertiary_link" or "residential" or "unclassified" or "road" or "living_street") return settings["local"];
        if (h is "service" or "track") return settings["service"];
        if (h is "pedestrian" or "footway" or "path" or "cycleway" or "steps") return settings["pedestrian"];
        return settings["default"];
    }

    public List<Brep> BuildRoadSurface(Curve curve, double width, double z = 0.12)
    {
        var half = Math.Max(0.25, width * 0.5);
        var tol = RhinoDoc.ActiveDoc.ModelAbsoluteTolerance;

        var c = curve.DuplicateCurve();
        c.Transform(Transform.Translation(0, 0, z));

        var outBreps = new List<Brep>();
        var segments = c.DuplicateSegments();
        if (segments is null || segments.Length == 0)
        {
            var div = c.DivideByCount(30, true);
            if (div is { Length: > 1 })
            {
                var builtSegments = new List<Curve>();
                for (var i = 0; i < div.Length - 1; i++)
                {
                    builtSegments.Add(new LineCurve(c.PointAt(div[i]), c.PointAt(div[i + 1])));
                }
                segments = builtSegments.ToArray();
            }
        }

        foreach (var seg in segments ?? Array.Empty<Curve>())
        {
            var p0 = seg.PointAtStart;
            var p1 = seg.PointAtEnd;
            var dx = p1.X - p0.X;
            var dy = p1.Y - p0.Y;
            var len = Math.Sqrt((dx * dx) + (dy * dy));
            if (len < 0.1) continue;
            var nx = -dy / len;
            var ny = dx / len;

            var ring = new Polyline
            {
                new Point3d(p0.X + (nx * half), p0.Y + (ny * half), z),
                new Point3d(p0.X - (nx * half), p0.Y - (ny * half), z),
                new Point3d(p1.X - (nx * half), p1.Y - (ny * half), z),
                new Point3d(p1.X + (nx * half), p1.Y + (ny * half), z),
                new Point3d(p0.X + (nx * half), p0.Y + (ny * half), z)
            };
            if (!ring.IsValid) continue;
            var b = Brep.CreatePlanarBreps(ring.ToNurbsCurve(), tol);
            if (b is { Length: > 0 }) outBreps.AddRange(b);
        }

        return outBreps;
    }

    public Brep? BuildBuilding(IReadOnlyList<(double X, double Y)> ring, double baseZ, double topZ)
    {
        var footprint = EnsureClosed(ring);
        if (footprint.Count < 4) return null;

        var curve = new Polyline(footprint.Select(p => new Point3d(p.X, p.Y, 0))).ToNurbsCurve();
        if (curve is null || !curve.IsClosed) return null;

        var h = Math.Max(0.5, topZ - baseZ);
        var extr = Extrusion.Create(curve, h, true);
        if (extr is null) return null;
        var brep = extr.ToBrep();
        if (brep is null) return null;
        return NormalizeBrepToBase(brep, baseZ);
    }

    public static (double Base, double Top) ParseHeightRange(IReadOnlyDictionary<string, string> tags)
    {
        static double? ParseLen(string? input)
        {
            if (string.IsNullOrWhiteSpace(input)) return null;
            var t = input.Trim().ToLowerInvariant();
            if (t.EndsWith("ft") && double.TryParse(t[..^2], NumberStyles.Float, CultureInfo.InvariantCulture, out var f)) return f * 0.3048;
            if (t.EndsWith("m") && double.TryParse(t[..^1], NumberStyles.Float, CultureInfo.InvariantCulture, out var m)) return m;
            return double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out var n) ? n : null;
        }

        var baseZ = ParseLen(tags.TryGetValue("min_height", out var mh) ? mh : null)
            ?? ((ParseLen(tags.TryGetValue("min_level", out var ml) ? ml : null) ?? 0) * 3.2);

        var top = ParseLen(tags.TryGetValue("height", out var h) ? h : null)
            ?? ((ParseLen(tags.TryGetValue("building:levels", out var lv) ? lv : null) ?? 0) * 3.2);

        if (top <= baseZ + 0.1) top = baseZ + 3.0;
        if (top <= 0) top = 10;
        return (Math.Max(0, baseZ), Math.Min(500, top));
    }

    public static bool IsParcel(IReadOnlyDictionary<string, string> tags)
    {
        tags.TryGetValue("boundary", out var boundary);
        tags.TryGetValue("landuse", out var landuse);
        boundary = (boundary ?? string.Empty).ToLowerInvariant();
        landuse = (landuse ?? string.Empty).ToLowerInvariant();
        if (boundary is "parcel" or "lot" or "plot" or "cadastral") return true;
        if (tags.ContainsKey("cadastre") || tags.ContainsKey("cadastral")) return true;
        return landuse is "residential" or "commercial" or "industrial" or "retail" or "allotments" or "garages";
    }

    private static int EnsureLayer(RhinoDoc doc, string name, System.Drawing.Color color)
    {
        var idx = doc.Layers.FindByFullPath(name, -1);
        if (idx >= 0)
        {
            var layer = doc.Layers[idx];
            var changed = false;
            if (!layer.IsVisible) { layer.IsVisible = true; changed = true; }
            if (layer.IsLocked) { layer.IsLocked = false; changed = true; }
            if (changed) doc.Layers.Modify(layer, idx, false);
            return idx;
        }

        var newLayer = new Layer { Name = name, Color = color, IsVisible = true, IsLocked = false };
        return doc.Layers.Add(newLayer);
    }

    private static List<(double X, double Y)> EnsureClosed(IReadOnlyList<(double X, double Y)> ring)
    {
        var output = ring.ToList();
        if (output.Count < 3) return output;
        var first = output[0];
        var last = output[^1];
        if (Math.Abs(first.X - last.X) > 1e-9 || Math.Abs(first.Y - last.Y) > 1e-9)
        {
            output.Add(first);
        }
        return output;
    }

    private static Brep NormalizeBrepToBase(Brep brep, double baseZ)
    {
        var bbox = brep.GetBoundingBox(true);
        if (!bbox.IsValid) return brep;
        var dz = baseZ - bbox.Min.Z;
        if (Math.Abs(dz) <= 1e-9) return brep;
        var moved = brep.DuplicateBrep();
        moved.Transform(Transform.Translation(0, 0, dz));
        return moved;
    }

    private static double DegreesToRadians(double value) => value * Math.PI / 180.0;
}
