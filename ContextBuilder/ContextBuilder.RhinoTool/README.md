# ContextBuilder Rhino Tool

Rhino-side context downloader/builder for polysurfaces.

## What it does
- Pick a center by address (Nominatim geocode) or by Rhino point mapped to lat/lon (manual input path).
- Download OSM context layers (buildings, roads, water, parks, parcels) using Overpass.
- Build Rhino geometry directly:
  - Buildings: closed footprints extruded to Breps (height from `height` or `building:levels` when available)
  - Roads: corridor surfaces (offset centerlines to width and create planar/loft surfaces)
  - Water / parks / parcels: planar surfaces where possible

## File
- `ContextBuilderRhinoTool.py`

## Run in Rhino 8
1. Open Rhino 8.
2. Open `ScriptEditor`.
3. Open `ContextBuilderRhinoTool.py`.
4. Run script.
5. A dialog appears with controls similar to ContextBuilder WinUI.

## Notes
- This tool currently uses online OSM endpoints.
- EPSG conversion in Rhino tool is metadata-guided (WGS84 native with optional WebMercator transform path).
- Treat this as an initial plugin-style tool in script form; if you want, it can be migrated into a compiled Rhino plugin next.
