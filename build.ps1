#Requires -Version 5.1
<#
.SYNOPSIS
    Builds the OneNoteExporter project and copies the release output to the "build" folder.
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
    throw 'No <Version> was found in OneNoteExporter.csproj.'
}

# ── 1. Prepare build folder ───────────────────────────────────────────────────
Write-Host ">> Preparing build folder..." -ForegroundColor Cyan
if (Test-Path $BuildDir) {
    Remove-Item -Path $BuildDir -Recurse -Force
}
New-Item -ItemType Directory -Path $BuildDir | Out-Null

# ── 2. dotnet publish (Release, self-contained, x64) ─────────────────────────
Write-Host ">> Starting dotnet publish..." -ForegroundColor Cyan
dotnet publish $ProjectFile `
    -c Release `
    -r win-x64 `
    --self-contained true `
    -p:PublishSingleFile=false `
    "-p:Version=$Version"

if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: dotnet publish failed (exit code $LASTEXITCODE)." -ForegroundColor Red
    exit $LASTEXITCODE
}

# ── 3. Copy files to build folder ─────────────────────────────────────────────
Write-Host ">> Copying files to: $BuildDir" -ForegroundColor Cyan
Copy-Item -Path "$PublishDir\*" -Destination $BuildDir -Recurse -Force

# ── 4. Build Inno Setup installer ─────────────────────────────────────────────
Write-Host ">> Looking for Inno Setup Compiler (ISCC.exe)..." -ForegroundColor Cyan

$IssFile  = Join-Path $ProjectRoot 'innosetup.iss'
$IsssPaths = @(
    'C:\Program Files (x86)\Inno Setup 6\ISCC.exe',
    'C:\Program Files\Inno Setup 6\ISCC.exe',
    'C:\Program Files (x86)\Inno Setup 5\ISCC.exe',
    'C:\Program Files\Inno Setup 5\ISCC.exe'
)

$Iscc = $IsssPaths | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $Iscc) {
    Write-Host "WARNING: ISCC.exe was not found; installer creation will be skipped." -ForegroundColor Yellow
    Write-Host "  -> Install Inno Setup: https://jrsoftware.org/isdl.php" -ForegroundColor Yellow
}
elseif (-not (Test-Path $IssFile)) {
    Write-Host "WARNING: innosetup.iss was not found; installer creation will be skipped." -ForegroundColor Yellow
}
else {
    Write-Host ">> Starting Inno Setup: $Iscc" -ForegroundColor Cyan
    & $Iscc "/DMyAppVersion=$Version" $IssFile
    if ($LASTEXITCODE -ne 0) {
        Write-Host "ERROR: Inno Setup failed (exit code $LASTEXITCODE)." -ForegroundColor Red
        exit $LASTEXITCODE
    }

    # Print setup EXE details
    $SetupExe = Get-ChildItem -Path $ArtifactDir -Filter 'OneNoteBackupExporter_Setup_*.exe' |
                Sort-Object LastWriteTime -Descending |
                Select-Object -First 1
    if ($SetupExe) {
        $size = [math]::Round($SetupExe.Length / 1MB, 1)
        Write-Host "  $($SetupExe.Name)  ($size MB)" -ForegroundColor Green
    }
}

# ── 5. Complete ───────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "Complete! Build output: $BuildDir" -ForegroundColor Green
$exe = Join-Path $BuildDir 'OneNoteExporter.exe'
if (Test-Path $exe) {
    $size = [math]::Round((Get-Item $exe).Length / 1MB, 1)
    Write-Host "  OneNoteExporter.exe  ($size MB)" -ForegroundColor Green
}
