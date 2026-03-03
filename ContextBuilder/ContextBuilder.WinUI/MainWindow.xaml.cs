using System.Globalization;
using System.Text.Json;
using ContextBuilder.WinUI.Models;
using ContextBuilder.WinUI.Services;
using ContextBuilder.WinUI.Services.Providers;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Windows.Storage.Pickers;
using WinRT.Interop;

namespace ContextBuilder.WinUI;

public sealed partial class MainWindow : Window
{
    private readonly NominatimGeocoder _geocoder = new();
    private readonly OsmOverpassProvider _osmProvider = new();
    private readonly OpenTopoDataElevationProvider _elevationProvider = new();
    private readonly SvgExporter _svgExporter = new();
    private readonly ObjExporter _objExporter = new();
    private readonly GeoJsonExporter _geoJsonExporter = new();
    private readonly GeoPackageExporter _geoPackageExporter = new();
    private readonly BuildingHeightCascadeService _buildingHeightCascade = new();
    private readonly ResourceAuditService _auditService = new();

    private GeoPoint? _selectedPoint;
    private bool _mapReady;

    public MainWindow()
    {
        InitializeComponent();
        _ = InitializeMapAsync();
    }

    private async Task InitializeMapAsync()
    {
        var mapPath = Path.Combine(AppContext.BaseDirectory, "Assets", "map.html");
        MapView.Source = new Uri(mapPath);
        await MapView.EnsureCoreWebView2Async();
        MapView.WebMessageReceived += MapView_WebMessageReceived;
        _mapReady = true;
    }

    private void MapView_WebMessageReceived(WebView2 sender, Microsoft.Web.WebView2.Core.CoreWebView2WebMessageReceivedEventArgs args)
    {
        using var doc = JsonDocument.Parse(args.WebMessageAsJson);
        if (!doc.RootElement.TryGetProperty("type", out var typeElement))
        {
            return;
        }

        if (!string.Equals(typeElement.GetString(), "picked", StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        var lat = doc.RootElement.GetProperty("lat").GetDouble();
        var lon = doc.RootElement.GetProperty("lon").GetDouble();
        _selectedPoint = new GeoPoint(lat, lon);
        StatusText.Text = $"Picked location: {lat:0.#####}, {lon:0.#####}";
        UpdateExtentPreview();
    }

    private async void SearchAddress_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(AddressBox.Text))
        {
            StatusText.Text = "Enter an address to search.";
            return;
        }

        try
        {
            StatusText.Text = "Searching address...";
            var point = await _geocoder.GeocodeAsync(AddressBox.Text, CancellationToken.None);
            if (point is null)
            {
                StatusText.Text = "No geocoding result found.";
                return;
            }

            _selectedPoint = point;
            var lat = point.Value.Latitude.ToString("0.######", CultureInfo.InvariantCulture);
            var lon = point.Value.Longitude.ToString("0.######", CultureInfo.InvariantCulture);
            await MapView.ExecuteScriptAsync($"setMapPin({lat}, {lon});");
            StatusText.Text = $"Address located at {lat}, {lon}";
            UpdateExtentPreview();
        }
        catch (Exception ex)
        {
            StatusText.Text = $"Address search failed: {ex.Message}";
        }
    }

    private void AuditSources_Click(object sender, RoutedEventArgs e)
    {
        var summary = _auditService.BuildAuditSummary(GetSelectedLayers(), UseElevationToggle.IsOn);
        StatusText.Text = summary;
    }

    private async void Export2d_Click(object sender, RoutedEventArgs e)
    {
        if (!TryBuildRequest(out var center, out var radiusMeters, out var layers, out var sourceEpsg, out var targetEpsg))
        {
            return;
        }

        var folder = await PickFolderAsync();
        if (folder is null)
        {
            return;
        }

        ToggleBusy(true);
        try
        {
            StatusText.Text = "Downloading vector layers...";
            var fetched = await _osmProvider.GetLayersAsync(center, radiusMeters, layers, CancellationToken.None);

            List<(GeoPoint Point, double Elevation)>? elevation = null;
            if (UseElevationToggle.IsOn || layers.Contains(ContextLayer.LidarTopoSurface))
            {
                StatusText.Text = "Sampling elevation points...";
                elevation = await _elevationProvider.SampleSquareGridAsync(center, radiusMeters, 30, CancellationToken.None);
            }

            var files = _svgExporter.Export(folder.Path, center, sourceEpsg, targetEpsg, fetched, elevation);
            StatusText.Text = $"2D export complete ({files.Count} SVG files): {folder.Path}";
        }
        catch (Exception ex)
        {
            StatusText.Text = $"2D export failed: {ex.Message}";
        }
        finally
        {
            ToggleBusy(false);
        }
    }

    private async void Export3d_Click(object sender, RoutedEventArgs e)
    {
        if (!TryBuildRequest(out var center, out var radiusMeters, out var layers, out var sourceEpsg, out var targetEpsg))
        {
            return;
        }

        var folder = await PickFolderAsync();
        if (folder is null)
        {
            return;
        }

        ToggleBusy(true);
        try
        {
            StatusText.Text = "Downloading vector layers...";
            var fetched = await _osmProvider.GetLayersAsync(center, radiusMeters, layers, CancellationToken.None);

            List<(GeoPoint Point, double Elevation)>? elevation = null;
            if (UseElevationToggle.IsOn || layers.Contains(ContextLayer.LidarTopoSurface))
            {
                StatusText.Text = "Sampling elevation grid for terrain...";
                elevation = await _elevationProvider.SampleSquareGridAsync(center, radiusMeters, 40, CancellationToken.None);
            }

            _buildingHeightCascade.ApplyCascade(fetched, elevation);
            StatusText.Text = "Building OBJ...";
            if (!string.Equals(sourceEpsg, "EPSG:4326", StringComparison.OrdinalIgnoreCase))
            {
                StatusText.Text = "Source EPSG metadata differs from OSM native EPSG:4326; proceeding with OSM coordinates.";
            }
            var path = _objExporter.Export(folder.Path, center, fetched, elevation, UseElevationToggle.IsOn);
            StatusText.Text = $"3D export complete: {path}";
        }
        catch (Exception ex)
        {
            StatusText.Text = $"3D export failed: {ex.Message}";
        }
        finally
        {
            ToggleBusy(false);
        }
    }

    private async void ExportGis_Click(object sender, RoutedEventArgs e)
    {
        if (!TryBuildRequest(out var center, out var radiusMeters, out var layers, out var sourceEpsg, out var targetEpsg))
        {
            return;
        }

        var folder = await PickFolderAsync();
        if (folder is null)
        {
            return;
        }

        ToggleBusy(true);
        try
        {
            StatusText.Text = "Downloading vector layers...";
            var fetched = await _osmProvider.GetLayersAsync(center, radiusMeters, layers, CancellationToken.None);

            List<(GeoPoint Point, double Elevation)>? elevation = null;
            if (UseElevationToggle.IsOn || layers.Contains(ContextLayer.LidarTopoSurface))
            {
                StatusText.Text = "Sampling elevation points...";
                elevation = await _elevationProvider.SampleSquareGridAsync(center, radiusMeters, 30, CancellationToken.None);
            }

            _buildingHeightCascade.ApplyCascade(fetched, elevation);
            if (UseGeoPackageToggle.IsOn)
            {
                var gpkgPath = _geoPackageExporter.Export(folder.Path, sourceEpsg, targetEpsg, fetched, elevation);
                StatusText.Text = $"GeoPackage export complete: {gpkgPath}";
            }
            else
            {
                var files = _geoJsonExporter.Export(folder.Path, sourceEpsg, targetEpsg, fetched, elevation);
                StatusText.Text = $"GeoJSON export complete ({files.Count} files): {folder.Path}";
            }
        }
        catch (Exception ex)
        {
            StatusText.Text = $"GIS export failed: {ex.Message}";
        }
        finally
        {
            ToggleBusy(false);
        }
    }

    private bool TryBuildRequest(out GeoPoint center, out double radiusMeters, out List<ContextLayer> layers, out string sourceEpsg, out string targetEpsg)
    {
        center = default;
        layers = [];
        sourceEpsg = "EPSG:4326";
        targetEpsg = "EPSG:3857";
        radiusMeters = 500;

        if (_selectedPoint is null)
        {
            StatusText.Text = "Pick a map point or search an address first.";
            return false;
        }

        center = _selectedPoint.Value;
        layers = GetSelectedLayers();
        if (layers.Count == 0)
        {
            StatusText.Text = "Select at least one layer.";
            return false;
        }

        if (!double.TryParse(RadiusMetersBox.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out radiusMeters))
        {
            if (!double.TryParse(RadiusMetersBox.Text, out radiusMeters))
            {
                StatusText.Text = "Radius is invalid.";
                return false;
            }
        }

        radiusMeters = Math.Clamp(radiusMeters, 100, 5000);
        var selectedSourceEpsg = (SourceEpsgCombo.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "EPSG:4326";
        sourceEpsg = selectedSourceEpsg.Split(' ')[0];
        targetEpsg = sourceEpsg;
        if (EnableTargetEpsgToggle.IsOn)
        {
            var parsedTarget = ParseTargetEpsgInput(TargetEpsgInputBox.Text);
            if (parsedTarget is null)
            {
                StatusText.Text = "Target EPSG is invalid. Enter numeric code like 3857.";
                return false;
            }

            targetEpsg = parsedTarget;
            if (!GeoProjection.IsTargetSupported(targetEpsg))
            {
                StatusText.Text = $"Target {targetEpsg} is accepted as metadata, but numeric transform is currently implemented for EPSG:4326 and EPSG:3857 only.";
            }
        }

        return true;
    }

    private List<ContextLayer> GetSelectedLayers()
    {
        var layers = new List<ContextLayer>();
        if (BuildingsCheck.IsChecked == true) layers.Add(ContextLayer.Buildings);
        if (RoadsCheck.IsChecked == true) layers.Add(ContextLayer.Roads);
        if (WaterCheck.IsChecked == true) layers.Add(ContextLayer.Water);
        if (ParksCheck.IsChecked == true) layers.Add(ContextLayer.Parks);
        if (ParcelsCheck.IsChecked == true) layers.Add(ContextLayer.Parcels);
        if (LidarTopoCheck.IsChecked == true) layers.Add(ContextLayer.LidarTopoSurface);
        return layers;
    }

    private async Task<Windows.Storage.StorageFolder?> PickFolderAsync()
    {
        var picker = new FolderPicker();
        picker.FileTypeFilter.Add("*");
        InitializeWithWindow.Initialize(picker, WindowNative.GetWindowHandle(this));
        return await picker.PickSingleFolderAsync();
    }

    private void ToggleBusy(bool isBusy)
    {
        Export2dButton.IsEnabled = !isBusy;
        Export3dButton.IsEnabled = !isBusy;
        ExportGisButton.IsEnabled = !isBusy;
    }

    private void RadiusMetersBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        UpdateExtentPreview();
    }

    private void UpdateExtentPreview()
    {
        if (_selectedPoint is null)
        {
            ExtentText.Text = "Pick a map point to show extent.";
            return;
        }

        var radius = ParseRadiusMetersOrDefault();
        var center = _selectedPoint.Value;
        var sw = GeoProjection.OffsetByMeters(center, -radius, -radius);
        var ne = GeoProjection.OffsetByMeters(center, radius, radius);

        ExtentText.Text =
            $"Lat: {sw.Latitude:0.######} to {ne.Latitude:0.######} | " +
            $"Lon: {sw.Longitude:0.######} to {ne.Longitude:0.######} | " +
            $"Size: {(radius * 2):0}m x {(radius * 2):0}m";

        if (_mapReady)
        {
            var swLat = sw.Latitude.ToString("0.######", CultureInfo.InvariantCulture);
            var swLon = sw.Longitude.ToString("0.######", CultureInfo.InvariantCulture);
            var neLat = ne.Latitude.ToString("0.######", CultureInfo.InvariantCulture);
            var neLon = ne.Longitude.ToString("0.######", CultureInfo.InvariantCulture);
            _ = MapView.ExecuteScriptAsync($"setExtentBox({swLat}, {swLon}, {neLat}, {neLon});");
        }
    }

    private double ParseRadiusMetersOrDefault()
    {
        if (!double.TryParse(RadiusMetersBox.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out var radius) &&
            !double.TryParse(RadiusMetersBox.Text, out radius))
        {
            radius = 500;
        }

        return Math.Clamp(radius, 100, 5000);
    }

    private static string? ParseTargetEpsgInput(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return null;
        }

        var digits = new string(text.Where(char.IsDigit).ToArray());
        if (string.IsNullOrEmpty(digits))
        {
            return null;
        }

        return $"EPSG:{digits}";
    }

    private void EnableTargetEpsgToggle_Toggled(object sender, RoutedEventArgs e)
    {
        TargetEpsgInputBox.IsEnabled = EnableTargetEpsgToggle.IsOn;
    }
}
