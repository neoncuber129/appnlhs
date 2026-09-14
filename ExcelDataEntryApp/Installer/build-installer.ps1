# Build Inno Setup installer for Excel Data Entry
$ErrorActionPreference = 'Stop'

$root = Split-Path $PSScriptRoot -Parent
$stage = Join-Path $PSScriptRoot 'payload'
$outDir = Join-Path $PSScriptRoot 'out'
$iscc = Join-Path ${env:ProgramFiles(x86)} 'Inno Setup 6\ISCC.exe'
if (-not (Test-Path $iscc)) {
  $iscc = Join-Path $env:ProgramFiles 'Inno Setup 6\ISCC.exe'
}
if (-not (Test-Path $iscc)) {
  throw 'Không tìm thấy Inno Setup 6 (ISCC.exe). Cài từ https://jrsoftware.org/isinfo.php'
}

Write-Host '==> Publishing app (self-contained, single-file)...'
dotnet publish (Join-Path $root 'ExcelDataEntryApp.csproj') `
  -c Release `
  -r win-x64 `
  --self-contained true `
  -p:PublishSingleFile=true `
  -p:IncludeNativeLibrariesForSelfExtract=true `
  -p:DebugType=None `
  -p:DebugSymbols=false `
  -o $stage

if (-not (Test-Path (Join-Path $stage 'ExcelDataEntryApp.exe'))) {
  throw 'Publish failed: ExcelDataEntryApp.exe not found.'
}

New-Item -ItemType Directory -Force -Path $outDir | Out-Null

Write-Host '==> Building Setup.exe with Inno Setup...'
& $iscc (Join-Path $PSScriptRoot 'Setup.iss')

$setup = Join-Path $outDir 'ExcelDataEntrySetup.exe'
if (-not (Test-Path $setup)) {
  throw 'Build failed: ExcelDataEntrySetup.exe not found.'
}

Write-Host ''
Write-Host "Done: $setup"
