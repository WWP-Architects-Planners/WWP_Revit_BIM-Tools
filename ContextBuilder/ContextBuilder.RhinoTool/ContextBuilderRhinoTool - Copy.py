"""
ContextBuilder Rhino Tool (Rhino 8, Python)

Plugin-style script UI to download OSM context data and build polysurfaces in Rhino.
"""

import json
import math
import urllib.error
import urllib.parse
import urllib.request

import Rhino
import Rhino.Geometry as rg
import scriptcontext as sc
import System
import System.Drawing as sd

import Eto.Forms as forms
import Eto.Drawing as drawing


EARTH_RADIUS_M = 6378137.0
OVERPASS_ENDPOINTS = [
    "https://overpass-api.de/api/interpreter",
    "https://overpass.kumi.systems/api/interpreter",
    "https://overpass.private.coffee/api/interpreter",
]


def http_get(url, timeout=20):
    req = urllib.request.Request(url, headers={"User-Agent": "ContextBuilderRhinoTool/0.1"})
    with urllib.request.urlopen(req, timeout=timeout) as res:
        return res.read().decode("utf-8")


def http_post_form(url, form_data, timeout=45):
    encoded = urllib.parse.urlencode(form_data).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=encoded,
        headers={
            "User-Agent": "ContextBuilderRhinoTool/0.1",
            "Content-Type": "application/x-www-form-urlencoded",
        },
    )
    with urllib.request.urlopen(req, timeout=timeout) as res:
        return res.read().decode("utf-8")


def geocode(address):
    q = urllib.parse.quote(address)
    url = "https://nominatim.openstreetmap.org/search?format=json&limit=1&q={}".format(q)
    payload = json.loads(http_get(url))
    if not payload:
        return None
    return float(payload[0]["lat"]), float(payload[0]["lon"])


def overpass_query(lat, lon, radius_m, layers, include_relations=True):
    radius = int(max(50, radius_m))
    selectors = []

    if "buildings" in layers:
        selectors.append("way(around:{r},{lat},{lon})[building]".format(r=radius, lat=lat, lon=lon))
        selectors.append("way(around:{r},{lat},{lon})[\"building:part\"]".format(r=radius, lat=lat, lon=lon))
        if include_relations:
            selectors.append("relation(around:{r},{lat},{lon})[building]".format(r=radius, lat=lat, lon=lon))
            selectors.append("relation(around:{r},{lat},{lon})[\"building:part\"]".format(r=radius, lat=lat, lon=lon))

    if "roads" in layers:
        selectors.append("way(around:{r},{lat},{lon})[highway]".format(r=radius, lat=lat, lon=lon))

    if "water" in layers:
        selectors.append("way(around:{r},{lat},{lon})[natural=water]".format(r=radius, lat=lat, lon=lon))
        if include_relations:
            selectors.append("relation(around:{r},{lat},{lon})[natural=water]".format(r=radius, lat=lat, lon=lon))
        selectors.append("way(around:{r},{lat},{lon})[waterway]".format(r=radius, lat=lat, lon=lon))

    if "parks" in layers:
        selectors.append("way(around:{r},{lat},{lon})[leisure=park]".format(r=radius, lat=lat, lon=lon))
        if include_relations:
            selectors.append("relation(around:{r},{lat},{lon})[leisure=park]".format(r=radius, lat=lat, lon=lon))
        selectors.append("way(around:{r},{lat},{lon})[landuse=grass]".format(r=radius, lat=lat, lon=lon))

    if "parcels" in layers:
        selectors.append("way(around:{r},{lat},{lon})[boundary=parcel]".format(r=radius, lat=lat, lon=lon))
        if include_relations:
            selectors.append("relation(around:{r},{lat},{lon})[boundary=parcel]".format(r=radius, lat=lat, lon=lon))
        selectors.append("way(around:{r},{lat},{lon})[boundary=lot]".format(r=radius, lat=lat, lon=lon))
        selectors.append("way(around:{r},{lat},{lon})[boundary=plot]".format(r=radius, lat=lat, lon=lon))
        selectors.append("way(around:{r},{lat},{lon})[cadastre]".format(r=radius, lat=lat, lon=lon))
        selectors.append("way(around:{r},{lat},{lon})[cadastral]".format(r=radius, lat=lat, lon=lon))
        selectors.append("way(around:{r},{lat},{lon})[landuse~\"residential|commercial|industrial|retail|allotments|garages\"]".format(r=radius, lat=lat, lon=lon))

    if not selectors:
        return "[out:json][timeout:40];();out body geom;"

    selectors = [s for s in selectors if s]
    body = "(" + ";".join(selectors) + ";)"
    return "[out:json][timeout:40];{};out body geom;".format(body)


def overpass_fetch(query, fallback_query=None, logger=None):
    last_err = None
    def log(msg):
        if logger:
            logger(msg)
    for endpoint in OVERPASS_ENDPOINTS:
        try:
            log("Overpass request: {}".format(endpoint))
            return http_post_form(endpoint, {"data": query}, timeout=50), endpoint, "primary"
        except urllib.error.HTTPError as ex:
            log("HTTP {} at {}".format(ex.code, endpoint))
            # Some Overpass mirrors reject larger/mixed queries with 400.
            if ex.code == 400 and fallback_query:
                try:
                    log("Retrying fallback query (ways-only/lean selectors) at {}".format(endpoint))
                    return http_post_form(endpoint, {"data": fallback_query}, timeout=50), endpoint, "fallback_400"
                except Exception as fallback_ex:
                    last_err = fallback_ex
                    continue
            if ex.code in (429, 504) and fallback_query:
                try:
                    log("Retrying fallback query after {} at {}".format(ex.code, endpoint))
                    return http_post_form(endpoint, {"data": fallback_query}, timeout=60), endpoint, "fallback_busy"
                except Exception as fallback_ex:
                    last_err = fallback_ex
                    continue
            last_err = ex
            continue
        except Exception as ex:
            log("Request error at {}: {}".format(endpoint, str(ex)))
            last_err = ex
            continue
    if last_err:
        raise last_err
    raise Exception("Overpass request failed.")


def parse_length_m(value):
    if value is None:
        return None
    txt = str(value).strip().lower()
    if not txt:
        return None
    try:
        if txt.endswith("ft"):
            return float(txt[:-2].strip()) * 0.3048
        if txt.endswith("m"):
            return float(txt[:-1].strip())
        return float(txt)
    except Exception:
        return None


def parse_height_range(tags):
    if not tags:
        return 0.0, 10.0

    base = parse_length_m(tags.get("min_height"))
    if base is None:
        min_level = parse_length_m(tags.get("min_level"))
        if min_level is not None:
            base = min_level * 3.2
    if base is None:
        base = 0.0

    top = parse_length_m(tags.get("height"))
    if top is None:
        levels = parse_length_m(tags.get("building:levels"))
        if levels is not None:
            top = levels * 3.2
    if top is None:
        top = 10.0

    if top <= base + 0.1:
        top = base + 3.0

    return max(0.0, base), min(500.0, top)


def to_local_xy(center_lat, center_lon, lat, lon):
    lat_r = math.radians(lat)
    lon_r = math.radians(lon)
    c_lat_r = math.radians(center_lat)
    c_lon_r = math.radians(center_lon)
    x = (lon_r - c_lon_r) * EARTH_RADIUS_M * math.cos(c_lat_r)
    y = (lat_r - c_lat_r) * EARTH_RADIUS_M
    return x, y


def is_closed_xy(pts):
    if len(pts) < 3:
        return False
    a = pts[0]
    b = pts[-1]
    return abs(a[0] - b[0]) < 1e-9 and abs(a[1] - b[1]) < 1e-9


def ensure_closed_xy(pts):
    if pts and not is_closed_xy(pts):
        pts.append(pts[0])
    return pts


def open_ring_xy(pts):
    if not pts:
        return []
    ring = list(pts)
    if len(ring) > 1 and abs(ring[0][0] - ring[-1][0]) < 1e-9 and abs(ring[0][1] - ring[-1][1]) < 1e-9:
        ring = ring[:-1]
    return ring


def canonical_cycle(points):
    n = len(points)
    if n == 0:
        return tuple()
    best = None
    for i in range(n):
        cand = tuple(points[i:] + points[:i])
        if best is None or cand < best:
            best = cand
    return best


def footprint_signature_xy(pts_xy, decimals=3):
    ring = open_ring_xy(pts_xy)
    if len(ring) < 3:
        return None
    rounded = [(round(p[0], decimals), round(p[1], decimals)) for p in ring]
    forward = canonical_cycle(rounded)
    reverse = canonical_cycle(list(reversed(rounded)))
    return min(forward, reverse)


def building_signature_xy(pts_xy, base_z, top_z, decimals=3):
    sig = footprint_signature_xy(pts_xy, decimals=decimals)
    if sig is None:
        return None
    return (sig, round(float(base_z), 2), round(float(top_z), 2))


def polyline_curve_from_xy(pts, z=0.0):
    points = [rg.Point3d(p[0], p[1], z) for p in pts]
    if len(points) < 2:
        return None
    poly = rg.Polyline(points)
    if not poly.IsValid:
        return None
    return poly.ToNurbsCurve()


def build_building_brep(footprint_xy, base_z, top_z):
    footprint_xy = ensure_closed_xy(list(footprint_xy))
    if len(footprint_xy) < 4:
        return None
    crv = polyline_curve_from_xy(footprint_xy, 0.0)
    if crv is None or not crv.IsClosed:
        return None
    height = max(0.5, float(top_z) - float(base_z))
    # Rhino-safe closed-profile extrusion.
    extr = rg.Extrusion.Create(crv, float(height), True)
    if extr is not None:
        brep = extr.ToBrep()
        if brep is not None:
            return normalize_brep_to_base(brep, base_z)

    # Fallback path for profiles that fail Extrusion.Create.
    planar = rg.Brep.CreatePlanarBreps(crv, sc.doc.ModelAbsoluteTolerance)
    if not planar:
        return None
    face_brep = planar[0]
    breps = face_brep.Faces[0].CreateExtrusion(crv, True) if face_brep.Faces.Count > 0 else None
    if breps:
        return normalize_brep_to_base(breps, base_z)
    return None


def normalize_brep_to_base(brep, base_z):
    if brep is None:
        return None
    bbox = brep.GetBoundingBox(True)
    if not bbox.IsValid:
        return brep
    dz = float(base_z) - bbox.Min.Z
    if abs(dz) <= 1e-9:
        return brep
    moved = brep.DuplicateBrep()
    moved.Transform(rg.Transform.Translation(0, 0, dz))
    return moved


def build_road_surface(line_xy, width):
    if len(line_xy) < 2:
        return []
    crv = polyline_curve_from_xy(line_xy, 0.12)
    if crv is None:
        return []
    return build_road_surface_from_curve(crv, width)


def build_road_surface_from_curve(crv, width):
    if crv is None:
        return []
    tol = sc.doc.ModelAbsoluteTolerance
    half_w = max(0.25, float(width) * 0.5)
    z = 0.12

    base_curve = crv.DuplicateCurve()
    base_curve.Transform(rg.Transform.Translation(0, 0, z))

    out = []
    segments = list(base_curve.DuplicateSegments() or [])
    if not segments:
        t_vals = base_curve.DivideByCount(32, True)
        if t_vals and len(t_vals) > 1:
            pts = [base_curve.PointAt(t) for t in t_vals]
            for i in range(len(pts) - 1):
                seg = rg.LineCurve(pts[i], pts[i + 1])
                segments.append(seg)
    for seg in segments:
        p0 = seg.PointAtStart
        p1 = seg.PointAtEnd
        dx = p1.X - p0.X
        dy = p1.Y - p0.Y
        seg_len = math.sqrt(dx * dx + dy * dy)
        if seg_len <= 0.1:
            continue
        nx = -dy / seg_len
        ny = dx / seg_len
        a = (p0.X + nx * half_w, p0.Y + ny * half_w)
        b = (p0.X - nx * half_w, p0.Y - ny * half_w)
        c = (p1.X - nx * half_w, p1.Y - ny * half_w)
        d = (p1.X + nx * half_w, p1.Y + ny * half_w)
        ring = [
            rg.Point3d(a[0], a[1], z),
            rg.Point3d(b[0], b[1], z),
            rg.Point3d(c[0], c[1], z),
            rg.Point3d(d[0], d[1], z),
            rg.Point3d(a[0], a[1], z),
        ]
        poly = rg.Polyline(ring)
        if not poly.IsValid:
            continue
        c0 = poly.ToNurbsCurve()
        if c0 is None:
            continue
        planar = rg.Brep.CreatePlanarBreps(c0, tol)
        if planar:
            out.extend(planar)

    return out


def default_road_width(highway_tag, width_settings=None):
    settings = width_settings or {}
    h = (highway_tag or "").lower().strip()
    if h in ("motorway", "motorway_link", "trunk", "trunk_link"):
        return settings.get("motorway", 24.0)
    if h in ("primary", "primary_link", "secondary", "secondary_link"):
        return settings.get("primary", 14.0)
    if h in ("tertiary", "tertiary_link", "residential", "unclassified", "road", "living_street"):
        return settings.get("local", 10.0)
    if h in ("pedestrian", "footway", "path", "cycleway", "steps", "bridleway"):
        return settings.get("pedestrian", 4.0)
    if h in ("service", "track"):
        return settings.get("service", 6.0)
    return settings.get("default", 8.0)


def is_parcel_feature(tags):
    if not tags:
        return False
    boundary = str(tags.get("boundary", "")).lower()
    landuse = str(tags.get("landuse", "")).lower()
    if boundary in ("parcel", "lot", "plot", "cadastral"):
        return True
    if "cadastre" in tags or "cadastral" in tags:
        return True
    if landuse in ("residential", "commercial", "industrial", "retail", "allotments", "garages"):
        return True
    return False


def ensure_layer(name, color):
    index = sc.doc.Layers.FindByFullPath(name, -1)
    if index >= 0:
        layer = sc.doc.Layers[index]
        changed = False
        if layer.IsVisible is False:
            layer.IsVisible = True
            changed = True
        if layer.IsLocked:
            layer.IsLocked = False
            changed = True
        if changed:
            sc.doc.Layers.Modify(layer, index, False)
        return index
    layer = Rhino.DocObjects.Layer()
    layer.Name = name
    layer.Color = color
    layer.IsVisible = True
    layer.IsLocked = False
    return sc.doc.Layers.Add(layer)


def add_brep_on_layer(brep, layer_index):
    attr = Rhino.DocObjects.ObjectAttributes()
    attr.LayerIndex = layer_index
    obj_id = sc.doc.Objects.AddBrep(brep, attr)
    return obj_id if obj_id != System.Guid.Empty else None


def add_curve_on_layer(curve, layer_index):
    attr = Rhino.DocObjects.ObjectAttributes()
    attr.LayerIndex = layer_index
    obj_id = sc.doc.Objects.AddCurve(curve, attr)
    return obj_id if obj_id != System.Guid.Empty else None


class ContextBuilderDialog(forms.Dialog[bool]):
    def __init__(self):
        super(ContextBuilderDialog, self).__init__()
        self.Title = "ContextBuilder Rhino Tool"
        self.ClientSize = drawing.Size(900, 380)
        self.Padding = drawing.Padding(10)
        self.Resizable = True

        self.address_box = forms.TextBox()
        self.address_box.Text = "110 Adelaide street east, toronto"
        self.address_box.Width = 560
        self.radius_box = forms.TextBox()
        self.radius_box.Text = "500"
        self.radius_box.Width = 100
        self.roads_width_box = forms.TextBox()
        self.roads_width_box.Text = "1.0"
        self.roads_width_box.Width = 100
        self.width_motorway_box = forms.TextBox()
        self.width_motorway_box.Text = "24"
        self.width_motorway_box.Width = 60
        self.width_primary_box = forms.TextBox()
        self.width_primary_box.Text = "14"
        self.width_primary_box.Width = 60
        self.width_local_box = forms.TextBox()
        self.width_local_box.Text = "10"
        self.width_local_box.Width = 60
        self.width_service_box = forms.TextBox()
        self.width_service_box.Text = "6"
        self.width_service_box.Width = 60
        self.width_ped_box = forms.TextBox()
        self.width_ped_box.Text = "4"
        self.width_ped_box.Width = 60
        self.width_default_box = forms.TextBox()
        self.width_default_box.Text = "8"
        self.width_default_box.Width = 60
        self.status = forms.Label()
        self.status.Text = "Ready."
        self.log_box = forms.TextArea()
        self.log_box.ReadOnly = True
        self.log_box.Wrap = False
        self.log_box.Height = 180

        self.chk_buildings = forms.CheckBox()
        self.chk_buildings.Text = "Buildings"
        self.chk_buildings.Checked = True
        self.chk_roads = forms.CheckBox()
        self.chk_roads.Text = "Roads"
        self.chk_roads.Checked = True
        self.chk_water = forms.CheckBox()
        self.chk_water.Text = "Water"
        self.chk_water.Checked = True
        self.chk_parks = forms.CheckBox()
        self.chk_parks.Text = "Parks"
        self.chk_parks.Checked = True
        self.chk_parcels = forms.CheckBox()
        self.chk_parcels.Text = "Parcel"
        self.chk_parcels.Checked = False
        self.chk_union_roads = forms.CheckBox()
        self.chk_union_roads.Text = "Boolean/Union Roads"
        self.chk_union_roads.Checked = True

        self.run_button = forms.Button()
        self.run_button.Text = "Download + Build Polysurfaces"
        self.close_button = forms.Button()
        self.close_button.Text = "Close"
        self.clear_log_button = forms.Button()
        self.clear_log_button.Text = "Clear Log"
        self.run_button.Click += self.on_run
        self.close_button.Click += self.on_close
        self.clear_log_button.Click += self.on_clear_log

        layout = forms.DynamicLayout()
        layout.Spacing = drawing.Size(10, 10)
        lbl_address = forms.Label()
        lbl_address.Text = "Address"
        lbl_radius = forms.Label()
        lbl_radius.Text = "Radius (m)"
        lbl_roadw = forms.Label()
        lbl_roadw.Text = "Road Width Scale"
        lbl_widths = forms.Label()
        lbl_widths.Text = "Road Widths (m)"
        lbl_layers = forms.Label()
        lbl_layers.Text = "Layers"
        lbl_w_motorway = forms.Label()
        lbl_w_motorway.Text = "Motorway/Trunk"
        lbl_w_primary = forms.Label()
        lbl_w_primary.Text = "Primary/Secondary"
        lbl_w_local = forms.Label()
        lbl_w_local.Text = "Local"
        lbl_w_service = forms.Label()
        lbl_w_service.Text = "Service/Track"
        lbl_w_ped = forms.Label()
        lbl_w_ped.Text = "Pedestrian"
        lbl_w_default = forms.Label()
        lbl_w_default.Text = "Default"
        layers_layout = forms.DynamicLayout()
        layers_layout.Spacing = drawing.Size(8, 6)
        layers_layout.AddRow(self.chk_buildings, self.chk_roads, self.chk_water, None)
        layers_layout.AddRow(self.chk_parks, self.chk_parcels, self.chk_union_roads, None)

        top_row = forms.DynamicLayout()
        top_row.Spacing = drawing.Size(8, 6)
        top_row.AddRow(lbl_address, self.address_box, lbl_radius, self.radius_box)

        second_row = forms.DynamicLayout()
        second_row.Spacing = drawing.Size(8, 6)
        second_row.AddRow(lbl_roadw, self.roads_width_box, None)

        widths_row1 = forms.DynamicLayout()
        widths_row1.Spacing = drawing.Size(8, 6)
        widths_row1.AddRow(
            lbl_w_motorway, self.width_motorway_box,
            lbl_w_primary, self.width_primary_box,
            lbl_w_local, self.width_local_box,
            None
        )

        widths_row2 = forms.DynamicLayout()
        widths_row2.Spacing = drawing.Size(8, 6)
        widths_row2.AddRow(
            lbl_w_service, self.width_service_box,
            lbl_w_ped, self.width_ped_box,
            lbl_w_default, self.width_default_box,
            None
        )

        button_row = forms.DynamicLayout()
        button_row.Spacing = drawing.Size(8, 6)
        button_row.AddRow(self.run_button, self.close_button, self.clear_log_button, None)

        layout.BeginVertical()
        layout.AddRow(top_row)
        layout.AddRow(second_row)
        layout.AddRow(lbl_widths)
        layout.AddRow(widths_row1)
        layout.AddRow(widths_row2)
        layout.AddRow(lbl_layers)
        layout.AddRow(layers_layout)
        layout.AddRow(button_row)
        layout.AddRow(self.status)
        layout.AddRow(self.log_box)
        layout.EndVertical()
        self.Content = layout

    def append_log(self, message):
        line = str(message)
        existing = self.log_box.Text or ""
        if existing:
            self.log_box.Text = existing + "\n" + line
        else:
            self.log_box.Text = line

    def on_clear_log(self, sender, e):
        self.log_box.Text = ""

    def selected_layers(self):
        layers = []
        if self.chk_buildings.Checked:
            layers.append("buildings")
        if self.chk_roads.Checked:
            layers.append("roads")
        if self.chk_water.Checked:
            layers.append("water")
        if self.chk_parks.Checked:
            layers.append("parks")
        if self.chk_parcels.Checked:
            layers.append("parcels")
        return layers

    def on_run(self, sender, e):
        try:
            self.log_box.Text = ""
            address = self.address_box.Text.strip()
            if not address:
                self.status.Text = "Address is required."
                self.append_log("Address is required.")
                return
            layers = self.selected_layers()
            if not layers:
                self.status.Text = "Select at least one layer."
                self.append_log("No layers selected.")
                return
            radius = float(self.radius_box.Text)
            road_w_scale = float(self.roads_width_box.Text)
            if road_w_scale <= 0:
                road_w_scale = 1.0
            width_settings = {
                "motorway": max(0.5, float(self.width_motorway_box.Text)),
                "primary": max(0.5, float(self.width_primary_box.Text)),
                "local": max(0.5, float(self.width_local_box.Text)),
                "service": max(0.5, float(self.width_service_box.Text)),
                "pedestrian": max(0.5, float(self.width_ped_box.Text)),
                "default": max(0.5, float(self.width_default_box.Text)),
            }
            self.append_log("Input: address='{}', radius={}m, road_scale={}".format(address, radius, road_w_scale))
            self.append_log("Layers: {}".format(", ".join(layers)))
            self.append_log("Road widths: {}".format(width_settings))

            geo = geocode(address)
            if not geo:
                self.status.Text = "Address lookup failed."
                self.append_log("Geocoding failed.")
                return
            center_lat, center_lon = geo
            self.append_log("Geocode result: lat={}, lon={}".format(round(center_lat, 6), round(center_lon, 6)))
            self.status.Text = "Downloading OSM data..."
            query = overpass_query(center_lat, center_lon, radius, layers, include_relations=True)
            fallback_query = overpass_query(center_lat, center_lon, radius, layers, include_relations=False)
            payload_text, endpoint_used, mode_used = overpass_fetch(query, fallback_query, logger=self.append_log)
            self.append_log("Overpass success: endpoint={}, mode={}".format(endpoint_used, mode_used))
            payload = json.loads(payload_text)
            elements = payload.get("elements", [])
            self.append_log("Overpass elements: {}".format(len(elements)))

            added = 0
            added_ids = []
            self.status.Text = "Building polysurfaces in Rhino..."
            layer_map = {
                "buildings": ensure_layer("ContextBuilder_Buildings", sd.Color.FromArgb(222, 222, 222)),
                "roads": ensure_layer("ContextBuilder_Roads", sd.Color.FromArgb(240, 180, 50)),
                "roads_centerline": ensure_layer("ContextBuilder_RoadCenterlines", sd.Color.FromArgb(255, 220, 120)),
                "water": ensure_layer("ContextBuilder_Water", sd.Color.FromArgb(80, 140, 220)),
                "parks": ensure_layer("ContextBuilder_Parks", sd.Color.FromArgb(120, 190, 120)),
                "parcels": ensure_layer("ContextBuilder_Parcels", sd.Color.FromArgb(180, 180, 180)),
            }
            road_curves_by_type = {}
            road_breps_raw = []
            seen_building_signatures = set()
            building_duplicates_skipped = 0
            building_seen = 0
            water_seen = 0
            park_seen = 0
            parcel_seen = 0
            road_seen = 0

            for el in elements:
                geom = el.get("geometry")
                if not geom:
                    continue
                tags = el.get("tags", {})
                pts_xy = [to_local_xy(center_lat, center_lon, n["lat"], n["lon"]) for n in geom if "lat" in n and "lon" in n]
                if len(pts_xy) < 2:
                    continue

                # Buildings
                if "buildings" in layers and ("building" in tags or "building:part" in tags):
                    building_seen += 1
                    base_z, top_z = parse_height_range(tags)
                    sig = building_signature_xy(pts_xy, base_z, top_z, decimals=3)
                    if sig is not None:
                        if sig in seen_building_signatures:
                            building_duplicates_skipped += 1
                            continue
                        seen_building_signatures.add(sig)
                    brep = build_building_brep(pts_xy, base_z, top_z)
                    if brep:
                        obj_id = add_brep_on_layer(brep, layer_map["buildings"])
                        if obj_id:
                            added_ids.append(obj_id)
                            added += 1
                    continue

                # Roads: collect centerlines, then join for continuous ribbons.
                if "roads" in layers and "highway" in tags:
                    road_seen += 1
                    crv = polyline_curve_from_xy(pts_xy, 0.12)
                    if crv:
                        center_id = add_curve_on_layer(crv, layer_map["roads_centerline"])
                        if center_id:
                            added_ids.append(center_id)
                            added += 1
                        road_type = str(tags.get("highway", "road")).lower()
                        if road_type not in road_curves_by_type:
                            road_curves_by_type[road_type] = []
                        road_curves_by_type[road_type].append(crv)
                    continue

                # Polygonal layers as planar surfaces (except linear waterways).
                is_water_area = "water" in layers and tags.get("natural") == "water"
                is_waterway = "water" in layers and ("waterway" in tags)
                is_water = is_water_area or is_waterway
                is_park = "parks" in layers and (tags.get("leisure") == "park" or tags.get("landuse") == "grass")
                is_parcel = "parcels" in layers and is_parcel_feature(tags)
                if is_water:
                    water_seen += 1
                if is_park:
                    park_seen += 1
                if is_parcel:
                    parcel_seen += 1
                if is_water or is_park or is_parcel:
                    if is_parcel:
                        # Parcels are more reliable as boundary lines in OSM/cadastral datasets.
                        parcel_pts = ensure_closed_xy(list(pts_xy)) if len(pts_xy) >= 3 else list(pts_xy)
                        crv = polyline_curve_from_xy(parcel_pts, 0.0)
                        if crv:
                            obj_id = add_curve_on_layer(crv, layer_map["parcels"])
                            if obj_id:
                                added_ids.append(obj_id)
                                added += 1
                    elif is_waterway:
                        # Streams/rivers from waterway tags should remain linework.
                        crv = polyline_curve_from_xy(list(pts_xy), 0.0)
                        if crv:
                            obj_id = add_curve_on_layer(crv, layer_map["water"])
                            if obj_id:
                                added_ids.append(obj_id)
                                added += 1
                    else:
                        ring = ensure_closed_xy(list(pts_xy))
                        if len(ring) >= 4:
                            crv = polyline_curve_from_xy(ring, 0.0)
                            if crv:
                                planar = rg.Brep.CreatePlanarBreps(crv, sc.doc.ModelAbsoluteTolerance)
                                if planar:
                                    for p in planar:
                                        if is_water:
                                            obj_id = add_brep_on_layer(p, layer_map["water"])
                                        else:
                                            obj_id = add_brep_on_layer(p, layer_map["parks"])
                                        if obj_id:
                                            added_ids.append(obj_id)
                                            added += 1

            if "roads" in layers and road_curves_by_type:
                tol = max(sc.doc.ModelAbsoluteTolerance * 2.0, 0.5)
                for road_type, curves in road_curves_by_type.items():
                    joined = rg.Curve.JoinCurves(curves, tol)
                    if not joined:
                        joined = curves
                    width = default_road_width(road_type, width_settings) * road_w_scale
                    self.append_log("Road type '{}': centerlines={}, joined={}, width={}m".format(
                        road_type, len(curves), len(joined), round(width, 2)))
                    for c in joined:
                        for b in build_road_surface_from_curve(c, width):
                            road_breps_raw.append(b)

                if road_breps_raw:
                    roads_to_add = road_breps_raw
                    if self.chk_union_roads.Checked and len(road_breps_raw) > 1:
                        union = rg.Brep.CreateBooleanUnion(road_breps_raw, sc.doc.ModelAbsoluteTolerance)
                        if union and len(union) > 0:
                            roads_to_add = list(union)
                        else:
                            # Fallback: attempt join if boolean union cannot resolve all overlaps.
                            joined_breps = rg.Brep.JoinBreps(road_breps_raw, sc.doc.ModelAbsoluteTolerance)
                            if joined_breps and len(joined_breps) > 0:
                                roads_to_add = list(joined_breps)
                        self.append_log("Road surfaces raw={}, after-merge={}".format(len(road_breps_raw), len(roads_to_add)))

                    for rb in roads_to_add:
                        obj_id = add_brep_on_layer(rb, layer_map["roads"])
                        if obj_id:
                            added_ids.append(obj_id)
                            added += 1

            self.append_log("Detected features: buildings={}, roads={}, water={}, parks={}, parcels={}".format(
                building_seen, road_seen, water_seen, park_seen, parcel_seen))
            if building_duplicates_skipped > 0:
                self.append_log("Building duplicates skipped: {}".format(building_duplicates_skipped))

            if added_ids:
                bbox = None
                for obj_id in added_ids:
                    rh_obj = sc.doc.Objects.FindId(obj_id)
                    if rh_obj is None:
                        continue
                    g_bbox = rh_obj.Geometry.GetBoundingBox(True)
                    if not g_bbox.IsValid:
                        continue
                    if bbox is None:
                        bbox = g_bbox
                    else:
                        bbox.Union(g_bbox)
                if bbox is not None and bbox.IsValid and sc.doc.Views.ActiveView is not None:
                    sc.doc.Views.ActiveView.ActiveViewport.ZoomBoundingBox(bbox)

            sc.doc.Views.Redraw()
            if added > 0:
                self.status.Text = "Done. Added {} objects.".format(added)
                self.append_log("Added object count: {}".format(added))
            else:
                self.status.Text = "Done, but no objects were added. Check layer visibility/locks and selected layers."
                self.append_log("No geometry added. If parcel count is 0, source likely has no parcel-like features in that area.")
        except Exception as ex:
            self.status.Text = "Failed: {}".format(str(ex))
            self.append_log("Failed with exception: {}".format(str(ex)))

    def on_close(self, sender, e):
        self.Close(False)


def run():
    dlg = ContextBuilderDialog()
    Rhino.UI.EtoExtensions.ShowSemiModal(dlg, Rhino.RhinoDoc.ActiveDoc, Rhino.UI.RhinoEtoApp.MainWindow)


if __name__ == "__main__":
    run()
