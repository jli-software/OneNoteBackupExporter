#Requires -Version 5.1
<#
.SYNOPSIS
    Baut das OneNoteExporter-Projekt und kopiert das fertige Release in den Ordner "build".
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ProjectRoot = $PSScriptRoot
$BuildDir    = Join-Path $ProjectRoot 'build'
$ArtifactDir = Join-Path $ProjectRoot 'artifacts'
$PublishDir  = Join-Path $ProjectRoot 'bin\Release\net10.0-windows\win-x64\publish'
$ProjectFile = Join-Path $ProjectRoot 'OneNoteExporter.csproj'
[xml]$ProjectXml = Get-Content $ProjectFile
$Version = $ProjectXml.Project.PropertyGroup.Version | Select-Object -First 1

if (-not $Version) {
    throw 'Keine <Version> in OneNoteExporter.csproj gefunden.'
}

# ── 1. build-Ordner leeren / anlegen ─────────────────────────────────────────
Write-Host ">> Bereite build-Ordner vor..." -ForegroundColor Cyan
if (Test-Path $BuildDir) {
    Remove-Item -Path $BuildDir -Recurse -Force
}
New-Item -ItemType Directory -Path $BuildDir | Out-Null

# ── 2. dotnet publish (Release, self-contained, x64) ─────────────────────────
Write-Host ">> Starte dotnet publish..." -ForegroundColor Cyan
dotnet publish $ProjectFile `
    -c Release `
    -r win-x64 `
    --self-contained true `
    -p:PublishSingleFile=false `
    "-p:Version=$Version"

if ($LASTEXITCODE -ne 0) {
    Write-Host "FEHLER: dotnet publish fehlgeschlagen (Exit-Code $LASTEXITCODE)." -ForegroundColor Red
    exit $LASTEXITCODE
}

# ── 3. Dateien in build-Ordner kopieren ──────────────────────────────────────
Write-Host ">> Kopiere Dateien nach: $BuildDir" -ForegroundColor Cyan
Copy-Item -Path "$PublishDir\*" -Destination $BuildDir -Recurse -Force

# ── 4. Inno Setup – Installer bauen ──────────────────────────────────────────
Write-Host ">> Suche Inno Setup Compiler (ISCC.exe)..." -ForegroundColor Cyan

$IssFile  = Join-Path $ProjectRoot 'innosetup.iss'
$IsssPaths = @(
    'C:\Program Files (x86)\Inno Setup 6\ISCC.exe',
    'C:\Program Files\Inno Setup 6\ISCC.exe',
    'C:\Program Files (x86)\Inno Setup 5\ISCC.exe',
    'C:\Program Files\Inno Setup 5\ISCC.exe'
)

$Iscc = $IsssPaths | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $Iscc) {
    Write-Host "WARNUNG: ISCC.exe nicht gefunden - Installer wird übersprungen." -ForegroundColor Yellow
    Write-Host "  -> Inno Setup installieren: https://jrsoftware.org/isdl.php" -ForegroundColor Yellow
}
elseif (-not (Test-Path $IssFile)) {
    Write-Host "WARNUNG: innosetup.iss nicht gefunden - Installer wird übersprungen." -ForegroundColor Yellow
}
else {
    Write-Host ">> Starte Inno Setup: $Iscc" -ForegroundColor Cyan
    & $Iscc "/DMyAppVersion=$Version" $IssFile
    if ($LASTEXITCODE -ne 0) {
        Write-Host "FEHLER: Inno Setup fehlgeschlagen (Exit-Code $LASTEXITCODE)." -ForegroundColor Red
        exit $LASTEXITCODE
    }

    # Setup-EXE ausgeben
    $SetupExe = Get-ChildItem -Path $ArtifactDir -Filter 'OneNoteBackupExporter_Setup_*.exe' |
                Sort-Object LastWriteTime -Descending |
                Select-Object -First 1
    if ($SetupExe) {
        $size = [math]::Round($SetupExe.Length / 1MB, 1)
        Write-Host "  $($SetupExe.Name)  ($size MB)" -ForegroundColor Green
    }
}

# ── 5. Fertig ─────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "Fertig! Build liegt in: $BuildDir" -ForegroundColor Green
$exe = Join-Path $BuildDir 'OneNoteExporter.exe'
if (Test-Path $exe) {
    $size = [math]::Round((Get-Item $exe).Length / 1MB, 1)
    Write-Host "  OneNoteExporter.exe  ($size MB)" -ForegroundColor Green
}
