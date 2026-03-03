using System.Globalization;
using System.Text.Json;
using ContextBuilder.WinUI.Models;

namespace ContextBuilder.WinUI.Services;

public sealed class NominatimGeocoder
{
    private readonly HttpClient _httpClient;

    public NominatimGeocoder()
    {
        _httpClient = new HttpClient();
        _httpClient.DefaultRequestHeaders.UserAgent.ParseAdd("ContextBuilder.WinUI/0.1 (+local app)");
    }

    public async Task<GeoPoint?> GeocodeAsync(string address, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(address))
        {
            return null;
        }

        var encoded = Uri.EscapeDataString(address);
        var url = $"https://nominatim.openstreetmap.org/search?format=json&limit=1&q={encoded}";
        var json = await _httpClient.GetStringAsync(url, cancellationToken).ConfigureAwait(false);
        using var doc = JsonDocument.Parse(json);
        if (doc.RootElement.ValueKind != JsonValueKind.Array || doc.RootElement.GetArrayLength() == 0)
        {
            return null;
        }

        var first = doc.RootElement[0];
        if (!first.TryGetProperty("lat", out var latElement) || !first.TryGetProperty("lon", out var lonElement))
        {
            return null;
        }

        var lat = double.Parse(latElement.GetString()!, CultureInfo.InvariantCulture);
        var lon = double.Parse(lonElement.GetString()!, CultureInfo.InvariantCulture);
        return new GeoPoint(lat, lon);
    }
}
