using System.Globalization;
using ContextBuilder.RhinoPlugin.Models;
using ContextBuilder.RhinoPlugin.Services;
using Eto.Drawing;
using Eto.Forms;
using Rhino;
using Rhino.DocObjects;
using Rhino.Geometry;
using System.IO;
using System.Text;

namespace ContextBuilder.RhinoPlugin.UI;

public sealed class ContextBuilderDialog : Dialog<bool>
{
    private readonly RhinoDoc _doc;
    private readonly OsmClient _osm = new();
    private readonly GeometryBuilder _geo = new();

    private readonly TextBox _address = new() { Text = "110 Adelaide street east, toronto", Width = 560 };
    private readonly TextBox _radius = new() { Text = "500", Width = 90 };
    private readonly TextBox _roadScale = new() { Text = "1.0", Width = 80 };

    private readonly CheckBox _buildings = new() { Text = "Buildings", Checked = true };
    private readonly CheckBox _roads = new() { Text = "Roads", Checked = true };
    private readonly CheckBox _water = new() { Text = "Water", Checked = true };
    private readonly CheckBox _parks = new() { Text = "Parks", Checked = true };
    private readonly CheckBox _parcels = new() { Text = "Parcel", Checked = false };
    private readonly CheckBox _unionRoads = new() { Text = "Boolean/Union Roads", Checked = false };

    private readonly TextBox _wMotorway = new() { Text = "24", Width = 60 };
    private readonly TextBox _wPrimary = new() { Text = "14", Width = 60 };
    private readonly TextBox _wLocal = new() { Text = "10", Width = 60 };
    private readonly TextBox _wService = new() { Text = "6", Width = 60 };
    private readonly TextBox _wPed = new() { Text = "4", Width = 60 };
    private readonly TextBox _wDefault = new() { Text = "8", Width = 60 };

    private readonly Label _status = new() { Text = "Ready." };
    private readonly TextArea _log = new() { ReadOnly = true, Wrap = false, Height = 200 };
    private readonly WebView _previewMap = new();

    private readonly Button _run = new() { Text = "Download + Build" };
    private readonly Button _preview = new() { Text = "Preview Area" };
    private int _previewHtmlVersion;

    public ContextBuilderDialog(RhinoDoc doc)
    {
        _doc = doc;
        Title = "ContextBuilder Rhino Plugin";
        ClientSize = new Size(1040, 760);
        Padding = new Padding(10);
        Resizable = true;

        _run.Click += async (_, _) => await RunAsync();
        _preview.Click += async (_, _) => await PreviewAsync();
        _address.KeyDown += async (_, e) =>
        {
            if (e.KeyData != Keys.Enter && e.Key != Keys.Enter)
            {
                return;
            }

            e.Handled = true;
            await PreviewAsync();
        };

        var close = new Button { Text = "Close" };
        close.Click += (_, _) => Close(true);

        var clearLog = new Button { Text = "Clear Log" };
        clearLog.Click += (_, _) => _log.Text = string.Empty;

        var labelW = 130;
        var widthBoxW = 72;
        _wMotorway.Width = widthBoxW;
        _wPrimary.Width = widthBoxW;
        _wLocal.Width = widthBoxW;
        _wService.Width = widthBoxW;
        _wPed.Width = widthBoxW;
        _wDefault.Width = widthBoxW;
        _log.Height = 110;

        Control Field(string text, Control input, int lblWidth = 120)
        {
            return new StackLayout
            {
                Orientation = Orientation.Horizontal,
                Spacing = 6,
                VerticalContentAlignment = VerticalAlignment.Center,
                Items =
                {
                    new Label { Text = text, Width = lblWidth },
                    input
                }
            };
        }

        var topRow = new StackLayout
        {
            Orientation = Orientation.Horizontal,
            Spacing = 12,
            VerticalContentAlignment = VerticalAlignment.Center,
            Items =
            {
                new Label { Text = "Address", Width = 90 },
                new StackLayoutItem(_address, true),
                new Label { Text = "Radius (m)", Width = 80 },
                _radius
            }
        };

        var scaleRow = new StackLayout
        {
            Orientation = Orientation.Horizontal,
            Spacing = 12,
            VerticalContentAlignment = VerticalAlignment.Center,
            Items =
            {
                new Label { Text = "Road Width Scale", Width = 130 },
                _roadScale
            }
        };

        var widthsRow1 = new StackLayout
        {
            Orientation = Orientation.Horizontal,
            Spacing = 14,
            VerticalContentAlignment = VerticalAlignment.Center,
            Items =
            {
                Field("Motorway/Trunk", _wMotorway, labelW),
                Field("Primary/Secondary", _wPrimary, labelW),
                Field("Local", _wLocal, 60)
            }
        };

        var widthsRow2 = new StackLayout
        {
            Orientation = Orientation.Horizontal,
            Spacing = 14,
            VerticalContentAlignment = VerticalAlignment.Center,
            Items =
            {
                Field("Service/Track", _wService, labelW),
                Field("Pedestrian", _wPed, labelW),
                Field("Default", _wDefault, 60)
            }
        };

        var layersRow = new StackLayout
        {
            Orientation = Orientation.Horizontal,
            Spacing = 16,
            VerticalContentAlignment = VerticalAlignment.Center,
            Items =
            {
                new Label { Text = "Layers", Width = 90 },
                _buildings, _roads, _water, _parks, _parcels, _unionRoads
            }
        };

        var buttonRow = new StackLayout
        {
            Orientation = Orientation.Horizontal,
            Spacing = 8,
            Items =
            {
                _preview, _run, close, clearLog
            }
        };

        _previewMap.Size = new Size(640, 520);
        LoadPreviewPlaceholder();

        var root = new TableLayout
        {
            Spacing = new Size(8, 8),
            Rows =
            {
                new TableRow(topRow),
                new TableRow(scaleRow),
                new TableRow(new Label { Text = "Road Widths (m)" }),
                new TableRow(widthsRow1),
                new TableRow(widthsRow2),
                new TableRow(layersRow),
                new TableRow(buttonRow),
                new TableRow(_status),
                new TableRow(new Label { Text = "Map Preview" }),
                new TableRow(new TableCell(_previewMap, true)) { ScaleHeight = true },
                new TableRow(new Label { Text = "Log" }),
                new TableRow(_log)
            }
        };

        Content = root;
    }

    private async Task RunAsync()
    {
        try
        {
            ToggleBusy(true);
            _log.Text = string.Empty;

            var address = (_address.Text ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(address))
            {
                SetStatus("Address is required.");
                return;
            }

            var layers = SelectedLayers();
            if (layers.Count == 0)
            {
                SetStatus("Select at least one layer.");
                return;
            }

            var radius = ParseOrDefault(_radius.Text, 500);
            var roadScale = Math.Max(0.1, ParseOrDefault(_roadScale.Text, 1.0));
            var widthSettings = new Dictionary<string, double>
            {
                ["motorway"] = Math.Max(0.5, ParseOrDefault(_wMotorway.Text, 24)),
                ["primary"] = Math.Max(0.5, ParseOrDefault(_wPrimary.Text, 14)),
                ["local"] = Math.Max(0.5, ParseOrDefault(_wLocal.Text, 10)),
                ["service"] = Math.Max(0.5, ParseOrDefault(_wService.Text, 6)),
                ["pedestrian"] = Math.Max(0.5, ParseOrDefault(_wPed.Text, 4)),
                ["default"] = Math.Max(0.5, ParseOrDefault(_wDefault.Text, 8)),
            };

            AppendLog($"Input: address='{address}', radius={radius}m, road_scale={roadScale}");
            AppendLog($"Layers: {string.Join(", ", layers)}");

            var ct = CancellationToken.None;
            var center = await _osm.GeocodeAsync(address, ct);
            if (center is null)
            {
                SetStatus("Address lookup failed.");
                return;
            }

            AppendLog($"Geocode result: lat={center.Value.Latitude:0.######}, lon={center.Value.Longitude:0.######}");
            LoadPreviewMap(center.Value, radius);
            SetStatus("Downloading OSM data...");
            var fetched = await _osm.FetchLayersAsync(center.Value, radius, layers, AppendLog, ct);
            AppendLog($"Overpass success: endpoint={fetched.Endpoint}, mode={fetched.Mode}, elements={fetched.Elements.Count}");

            SetStatus("Building geometry...");
            var layerMap = _geo.EnsureLayers(_doc);
            var addedIds = new List<Guid>();
            var seenBuildings = new HashSet<string>(StringComparer.Ordinal);
            var roadsToAdd = new List<Brep>();

            foreach (var element in fetched.Elements)
            {
                var tags = element.Tags;
                var pts = element.Geometry.Select(p => _geo.ToLocal(center.Value, p)).ToList();
                if (pts.Count < 2) continue;

                if (layers.Contains("buildings") && (tags.ContainsKey("building") || tags.ContainsKey("building:part")))
                {
                    var (baseZ, topZ) = GeometryBuilder.ParseHeightRange(tags);
                    var sig = BuildingSignature(pts, baseZ, topZ);
                    if (!seenBuildings.Add(sig)) continue;

                    var brep = _geo.BuildBuilding(pts, baseZ, topZ);
                    if (brep is not null)
                    {
                        var id = AddBrepOnLayer(_doc, brep, layerMap["buildings"]);
                        if (id != Guid.Empty) addedIds.Add(id);
                    }
                    continue;
                }

                if (layers.Contains("roads") && tags.ContainsKey("highway"))
                {
                    var centerline = BuildCurve(pts, 0.12, close: false);
                    if (centerline is not null)
                    {
                        var cId = AddCurveOnLayer(_doc, centerline, layerMap["roads_center"]);
                        if (cId != Guid.Empty) addedIds.Add(cId);

                        var width = GeometryBuilder.DefaultRoadWidth(tags.GetValueOrDefault("highway", "road"), widthSettings) * roadScale;
                        roadsToAdd.AddRange(_geo.BuildRoadSurface(centerline, width));
                    }
                    continue;
                }

                var isWaterway = layers.Contains("water") && tags.ContainsKey("waterway");
                var isWaterArea = layers.Contains("water") && tags.GetValueOrDefault("natural", "").Equals("water", StringComparison.OrdinalIgnoreCase);
                var isPark = layers.Contains("parks") && (tags.GetValueOrDefault("leisure", "").Equals("park", StringComparison.OrdinalIgnoreCase) || tags.GetValueOrDefault("landuse", "").Equals("grass", StringComparison.OrdinalIgnoreCase));
                var isParcel = layers.Contains("parcels") && GeometryBuilder.IsParcel(tags);

                if (isParcel)
                {
                    var c = BuildCurve(pts, 0, close: pts.Count >= 3);
                    if (c is not null)
                    {
                        var id = AddCurveOnLayer(_doc, c, layerMap["parcels"]);
                        if (id != Guid.Empty) addedIds.Add(id);
                    }
                    continue;
                }

                if (isWaterway)
                {
                    var c = BuildCurve(pts, 0, close: false);
                    if (c is not null)
                    {
                        var id = AddCurveOnLayer(_doc, c, layerMap["water"]);
                        if (id != Guid.Empty) addedIds.Add(id);
                    }
                    continue;
                }

                if (isWaterArea || isPark)
                {
                    var c = BuildCurve(pts, 0, close: true);
                    if (c is null) continue;
                    var planar = Brep.CreatePlanarBreps(c, _doc.ModelAbsoluteTolerance);
                    if (planar is not { Length: > 0 }) continue;
                    foreach (var b in planar)
                    {
                        var id = AddBrepOnLayer(_doc, b, layerMap[isWaterArea ? "water" : "parks"]);
                        if (id != Guid.Empty) addedIds.Add(id);
                    }
                }
            }

            if (roadsToAdd.Count > 0)
            {
                IReadOnlyList<Brep> roadsFinal = roadsToAdd;
                if (_unionRoads.Checked == true && roadsToAdd.Count > 1)
                {
                    var rawArea = roadsToAdd.Sum(Area);
                    var union = Brep.CreateBooleanUnion(roadsToAdd, _doc.ModelAbsoluteTolerance);
                    if (union is { Length: > 0 })
                    {
                        var unionArea = union.Sum(Area);
                        if (unionArea <= rawArea * 1.05)
                        {
                            roadsFinal = union;
                        }
                    }
                }

                foreach (var b in roadsFinal)
                {
                    var id = AddBrepOnLayer(_doc, b, layerMap["roads"]);
                    if (id != Guid.Empty) addedIds.Add(id);
                }
            }

            ZoomTo(addedIds);
            _doc.Views.Redraw();
            SetStatus($"Done. Added {addedIds.Count} objects.");
        }
        catch (Exception ex)
        {
            SetStatus($"Failed: {ex.Message}");
            AppendLog($"Failed with exception: {ex}");
        }
        finally
        {
            ToggleBusy(false);
        }
    }

    private async Task PreviewAsync()
    {
        try
        {
            ToggleBusy(true);
            var address = (_address.Text ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(address))
            {
                SetStatus("Address is required.");
                return;
            }

            var radius = Math.Max(50, ParseOrDefault(_radius.Text, 500));
            SetStatus("Geocoding preview location...");
            var center = await _osm.GeocodeAsync(address, CancellationToken.None);
            if (center is null)
            {
                SetStatus("Address lookup failed.");
                return;
            }

            LoadPreviewMap(center.Value, radius);
            AppendLog($"Preview: lat={center.Value.Latitude:0.######}, lon={center.Value.Longitude:0.######}, radius={radius:0}m");
            SetStatus("Preview updated.");
        }
        catch (Exception ex)
        {
            SetStatus($"Preview failed: {ex.Message}");
            AppendLog($"Preview failed with exception: {ex}");
        }
        finally
        {
            ToggleBusy(false);
        }
    }

    private void ToggleBusy(bool busy)
    {
        _run.Enabled = !busy;
        _preview.Enabled = !busy;
    }

    private void SetStatus(string text)
    {
        _status.Text = text;
        AppendLog(text);
    }

    private void AppendLog(string text)
    {
        var app = Application.Instance;
        if (app is null)
        {
            AppendLogCore(text);
            return;
        }

        app.AsyncInvoke(() => AppendLogCore(text));
    }

    private void AppendLogCore(string text)
    {
        _log.Text = string.IsNullOrEmpty(_log.Text) ? text : _log.Text + System.Environment.NewLine + text;
    }

    private void LoadPreviewPlaceholder()
    {
        var html = """
<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <style>
    body { margin: 0; font-family: Segoe UI, sans-serif; background: #0b1320; color: #e5e7eb; }
    .wrap { height: 100vh; display: grid; place-items: center; text-align: center; padding: 20px; }
    .card { border: 1px solid #334155; border-radius: 12px; padding: 16px; background: #111827; max-width: 480px; }
  </style>
</head>
<body>
  <div class="wrap">
    <div class="card">
      <h3>Map preview</h3>
      <p>Enter an address and click <b>Preview Area</b> to view the selected radius before download.</p>
    </div>
  </div>
</body>
</html>
""";
        WritePreviewHtml(html);
    }

    private void LoadPreviewMap(GeoPoint center, double radiusMeters)
    {
        var lat = center.Latitude.ToString("0.######", CultureInfo.InvariantCulture);
        var lon = center.Longitude.ToString("0.######", CultureInfo.InvariantCulture);
        var halfSideMeters = Math.Max(50, radiusMeters);
        var (south, west, north, east) = CalculateSquareBounds(center, halfSideMeters);
        var southStr = south.ToString("0.######", CultureInfo.InvariantCulture);
        var westStr = west.ToString("0.######", CultureInfo.InvariantCulture);
        var northStr = north.ToString("0.######", CultureInfo.InvariantCulture);
        var eastStr = east.ToString("0.######", CultureInfo.InvariantCulture);

        var html = string.Format(CultureInfo.InvariantCulture, """
<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
  <style>
    html, body, #map {{ height: 100%; margin: 0; }}
    .leaflet-container {{ font: 12px/1.2 Segoe UI, sans-serif; }}
  </style>
</head>
<body>
  <div id="map"></div>
  <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
  <script>
    const lat = {0};
    const lon = {1};
    const south = {2};
    const west = {3};
    const north = {4};
    const east = {5};
    const map = L.map("map", {{ zoomControl: true }}).setView([lat, lon], 16);
    L.tileLayer("https://{{s}}.tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png", {{
      maxZoom: 19,
      attribution: "&copy; OpenStreetMap contributors"
    }}).addTo(map);

    L.marker([lat, lon]).addTo(map).bindPopup("Selected center").openPopup();
    const area = L.rectangle([[south, west], [north, east]], {{
      color: "#0f766e",
      fillColor: "#14b8a6",
      fillOpacity: 0.2,
      weight: 2
    }}).addTo(map).bindTooltip("Selected square area", {{ sticky: true }});
    map.fitBounds(area.getBounds(), {{ padding: [24, 24] }});
  </script>
</body>
</html>
""", lat, lon, southStr, westStr, northStr, eastStr);

        WritePreviewHtml(html);
    }

    private static (double South, double West, double North, double East) CalculateSquareBounds(GeoPoint center, double halfSideMeters)
    {
        var latRad = center.Latitude * Math.PI / 180.0;
        var metersPerDegreeLat = 111_320d;
        var metersPerDegreeLon = Math.Max(1e-9, metersPerDegreeLat * Math.Cos(latRad));

        var dLat = halfSideMeters / metersPerDegreeLat;
        var dLon = halfSideMeters / metersPerDegreeLon;

        return (
            center.Latitude - dLat,
            center.Longitude - dLon,
            center.Latitude + dLat,
            center.Longitude + dLon
        );
    }

    private void WritePreviewHtml(string html)
    {
        var fileName = $"ContextBuilderRhinoPreview-{Interlocked.Increment(ref _previewHtmlVersion)}.html";
        var path = Path.Combine(Path.GetTempPath(), fileName);
        File.WriteAllText(path, html, Encoding.UTF8);
        _previewMap.Url = new Uri(path);
    }

    private List<string> SelectedLayers()
    {
        var layers = new List<string>();
        if (_buildings.Checked == true) layers.Add("buildings");
        if (_roads.Checked == true) layers.Add("roads");
        if (_water.Checked == true) layers.Add("water");
        if (_parks.Checked == true) layers.Add("parks");
        if (_parcels.Checked == true) layers.Add("parcels");
        return layers;
    }

    private static double ParseOrDefault(string? text, double fallback)
    {
        if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var v)) return v;
        if (double.TryParse(text, out v)) return v;
        return fallback;
    }

    private static Curve? BuildCurve(IReadOnlyList<(double X, double Y)> points, double z, bool close)
    {
        var list = points.ToList();
        if (close && list.Count >= 3)
        {
            var f = list[0];
            var l = list[^1];
            if (Math.Abs(f.X - l.X) > 1e-9 || Math.Abs(f.Y - l.Y) > 1e-9)
            {
                list.Add(f);
            }
        }

        if (list.Count < 2) return null;
        var poly = new Polyline(list.Select(p => new Point3d(p.X, p.Y, z)));
        return poly.IsValid ? poly.ToNurbsCurve() : null;
    }

    private static string BuildingSignature(IReadOnlyList<(double X, double Y)> points, double baseZ, double topZ)
    {
        var rounded = points.Select(p => $"{Math.Round(p.X, 3):0.###},{Math.Round(p.Y, 3):0.###}").ToArray();
        return string.Join(";", rounded) + $"|{Math.Round(baseZ, 2):0.##}|{Math.Round(topZ, 2):0.##}";
    }

    private static Guid AddBrepOnLayer(RhinoDoc doc, Brep brep, int layerIndex)
    {
        var attr = new ObjectAttributes { LayerIndex = layerIndex };
        return doc.Objects.AddBrep(brep, attr);
    }

    private static Guid AddCurveOnLayer(RhinoDoc doc, Curve curve, int layerIndex)
    {
        var attr = new ObjectAttributes { LayerIndex = layerIndex };
        return doc.Objects.AddCurve(curve, attr);
    }

    private static double Area(Brep brep)
    {
        using var amp = AreaMassProperties.Compute(brep);
        return amp?.Area ?? 0;
    }

    private void ZoomTo(IReadOnlyList<Guid> ids)
    {
        if (ids.Count == 0 || _doc.Views.ActiveView is null) return;
        BoundingBox? bbox = null;
        foreach (var id in ids)
        {
            var obj = _doc.Objects.FindId(id);
            if (obj?.Geometry is null) continue;
            var b = obj.Geometry.GetBoundingBox(true);
            if (!b.IsValid) continue;
            bbox = bbox.HasValue ? Union(bbox.Value, b) : b;
        }

        if (bbox.HasValue && bbox.Value.IsValid)
        {
            _doc.Views.ActiveView.ActiveViewport.ZoomBoundingBox(bbox.Value);
        }
    }

    private static BoundingBox Union(BoundingBox a, BoundingBox b)
    {
        a.Union(b);
        return a;
    }
}
