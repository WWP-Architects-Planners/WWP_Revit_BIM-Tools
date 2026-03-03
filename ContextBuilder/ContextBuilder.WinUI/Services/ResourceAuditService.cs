using ContextBuilder.WinUI.Models;

namespace ContextBuilder.WinUI.Services;

public sealed class ResourceAuditService
{
    public string BuildAuditSummary(IEnumerable<ContextLayer> layers, bool needsElevation)
    {
        var selected = layers.ToHashSet();
        var lines = new List<string>
        {
            "Resource Audit:",
            "- Vector data: OpenStreetMap + Overpass API (active)",
            "- Geocoding: OSM Nominatim (active)"
        };

        if (selected.Contains(ContextLayer.LidarTopoSurface) || needsElevation)
        {
            lines.Add("- Elevation: OpenTopoData (active), flat fallback (active)");
        }
        else
        {
            lines.Add("- Elevation: not requested");
        }

        lines.Add("- Overture Maps: planned adapter slot (not implemented)");
        lines.Add("- Autodesk Forma: planned connector slot (not implemented)");
        lines.Add("- Best source decision: OSM stack chosen for immediate open coverage and API simplicity.");

        return string.Join(Environment.NewLine, lines);
    }
}
