# Publishes Envoy as a multi-file self-contained Windows x64 app.
# Output: ..\publish\Envoy-win-x64  (also zipped alongside).
#
# Multi-file on purpose: avoids the first-run extraction delay that
# PublishSingleFile imposes on every new machine.

param(
    [switch]$NoZip
)

$ErrorActionPreference = 'Stop'
$repo  = Split-Path -Parent $PSScriptRoot
$proj  = Join-Path $repo 'Envoy\Envoy.csproj'
$out   = Join-Path $repo 'publish\Envoy-win-x64'
$zip   = "$out.zip"

if (Test-Path $out) { Remove-Item $out -Recurse -Force }
if (Test-Path $zip) { Remove-Item $zip -Force }

dotnet publish $proj `
    -c Release `
    -r win-x64 `
    --self-contained true `
    -p:PublishProfile=win-x64 `
    -p:PublishSingleFile=false `
    -p:PublishReadyToRun=true `
    -p:DebugType=none `
    -p:DebugSymbols=false

if ($LASTEXITCODE -ne 0) { throw "dotnet publish failed ($LASTEXITCODE)" }

if (-not $NoZip) {
    Compress-Archive -Path (Join-Path $out '*') -DestinationPath $zip -Force
    Write-Host "Zipped -> $zip"
}

Write-Host "Done. Folder: $out"
