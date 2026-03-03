namespace ContextBuilder.WinUI.Models;

public sealed class LayerResult
{
    public required ContextLayer Layer { get; init; }
    public List<List<GeoPoint>> LineStrings { get; } = [];
    public List<List<GeoPoint>> Polygons { get; } = [];
    public List<BuildingFootprint> BuildingFootprints { get; } = [];
}
