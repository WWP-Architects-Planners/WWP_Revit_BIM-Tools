# ContextBuilder

WinUI3 desktop tool to pick a location and export:
- 2D layer SVGs (`buildings`, `roads`, `water`, `parks`, `parcel`, optional `lidar_toposurface`)
- 3D context OBJ (`context_map.obj`)

## Data sources
- Geocoding: OSM Nominatim
- Vector layers: OSM Overpass API
- Elevation (optional): OpenTopoData (ASTER30m), with flat fallback

## Run
```powershell
dotnet build .\ContextBuilder.WinUI\ContextBuilder.WinUI.csproj -c Debug -p:Platform=x64
dotnet run --project .\ContextBuilder.WinUI\ContextBuilder.WinUI.csproj -c Debug -p:Platform=x64
```

## Workflow
1. Enter address or click on map.
2. Choose layers.
3. Choose flat or elevation toggle (5 mm UI toggle style).
4. Pick EPSG output mode (`EPSG:4326` or `EPSG:3857`).
5. Click export buttons:
   - `Export 2D (SVGs)`
   - `Export 3D (OBJ)`

## Notes
- Overture and Autodesk Forma are exposed as planned provider slots in the audit summary and can be added as future adapters.
