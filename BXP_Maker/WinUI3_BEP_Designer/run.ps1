param(
    [string]$Configuration = "Debug"
)

$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectRoot

Write-Host "Publishing WinUI app ($Configuration, win-x64)..."
Get-Process -Name "BEPDesigner.WinUI" -ErrorAction SilentlyContinue | Stop-Process -Force

dotnet publish -c $Configuration -r win-x64 -p:Platform=x64 | Out-Host
if ($LASTEXITCODE -ne 0) {
    throw "dotnet publish failed with exit code $LASTEXITCODE"
}

$exe = Join-Path $projectRoot "bin\x64\$Configuration\net8.0-windows10.0.19041.0\win-x64\publish\BEPDesigner.WinUI.exe"

if (!(Test-Path $exe)) {
    throw "Expected exe not found: $exe"
}

Write-Host "Launching: $exe"
Start-Process -FilePath $exe
 