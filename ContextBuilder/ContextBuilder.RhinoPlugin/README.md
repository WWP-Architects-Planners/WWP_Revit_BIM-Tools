# ContextBuilder.RhinoPlugin

Rhino `.rhp` plugin version of the ContextBuilder road/building downloader.

## Build

1. Install Rhino 8 (for `RhinoCommon.dll` at `C:\Program Files\Rhino 8\System\RhinoCommon.dll`).
2. Build from repo root:

```powershell
dotnet build ContextBuilder.RhinoPlugin\ContextBuilder.RhinoPlugin.csproj -c Debug
```

Output: `ContextBuilder.RhinoPlugin.rhp` in `bin\Debug\net7.0-windows\`.

## Load in Rhino

1. Rhino command: `PluginManager`
2. `Install...` and pick the generated `.rhp`
3. Run command: `ContextBuilderDownload`

## Notes

- UI is Eto (Rhino-native), not WinUI3.
- Roads generate centerlines + surface strips.
- Waterways are linework, water areas are surfaces.
- Parcels are boundary curves.
