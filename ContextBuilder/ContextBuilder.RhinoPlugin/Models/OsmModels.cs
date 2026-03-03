namespace ContextBuilder.RhinoPlugin.Models;

public readonly record struct GeoPoint(double Latitude, double Longitude);

public sealed class OsmElement
{
    public required Dictionary<string, string> Tags { get; init; }
    public required List<GeoPoint> Geometry { get; init; }
}
