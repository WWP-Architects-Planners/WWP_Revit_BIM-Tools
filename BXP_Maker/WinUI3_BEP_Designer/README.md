# WinUI3 BEP Designer (Python-backed)

This app is a WinUI 3 desktop UI with a Python generation engine for BIM Execution Plan sections.

## Implemented features

- Multi-tab UI:
  - `Project`
  - `ACC Workflow`
  - `Clash Sessions`
  - `Output & Export`
- ACC package method selector:
  - `1. Live Model`
  - `2. Shared Package`
  - `3. Consumed Model`
- Clash session controls:
  - `Delete From Beginning`
  - `Generate Back Missing Sessions`
  - `Start Fresh This Cycle`
- Generate BEP section text via Python (`Generate BEP Section`)
- Copy generated output
- Save current form/session state as preset (`.json`)
- Load preset (`.json`) and repopulate all fields/sessions
- Export generated output to Word (`.docx`)

## Run (recommended)

From `BXP_Maker/WinUI3_BEP_Designer`:

```powershell
.\run.ps1
```

This publishes self-contained output and launches:

- `bin/x64/Debug/net8.0-windows10.0.19041.0/win-x64/publish/BEPDesigner.WinUI.exe`

## Optional release run

```powershell
.\run.ps1 -Configuration Release
```

## Requirements

- Windows 10/11
- .NET 8 SDK
- Python in PATH (`python --version` works)

## Architecture

- WinUI front end: `MainWindow.xaml`
- UI behavior + export/presets: `MainWindow.xaml.cs`
- Payload model: `Models/BepPayload.cs`
- Python bridge: `Services/PythonBridge.cs`
- Python generator: `python/bep_engine.py`

`Generate BEP Section` serializes the current payload to JSON, sends it to Python, and displays generated markdown. `Export DOCX` writes that output into a Word file.
