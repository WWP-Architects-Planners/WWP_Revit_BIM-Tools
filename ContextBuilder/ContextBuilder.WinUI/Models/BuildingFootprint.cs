namespace ContextBuilder.WinUI.Models;

public sealed class BuildingFootprint
{
    public required List<GeoPoint> Ring { get; init; }
    public double HeightMeters { get; set; }
    public string HeightSource { get; set; } = "unknown";
}
