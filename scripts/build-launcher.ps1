param(
    [string]$OutputPath = (Join-Path $PSScriptRoot "..\ScreenTimeTracker.exe")
)

$ErrorActionPreference = "Stop"

$projectRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
$sourcePath = Join-Path $projectRoot "launcher\ScreenTimeTrackerLauncher.cs"
$iconPath = Join-Path $projectRoot "Image\tracker_time.ico"
$resolvedOutputPath = [System.IO.Path]::GetFullPath($OutputPath)

if (-not (Test-Path -LiteralPath $sourcePath)) {
    throw "Launcher source was not found: $sourcePath"
}

$candidates = @(
    (Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\csc.exe"),
    (Join-Path $env:WINDIR "Microsoft.NET\Framework\v4.0.30319\csc.exe")
)

$cscPath = $null
foreach ($candidate in $candidates) {
    if (Test-Path -LiteralPath $candidate) {
        $cscPath = $candidate
        break
    }
}

if ([string]::IsNullOrWhiteSpace([string]$cscPath)) {
    throw "Could not find csc.exe. Install .NET Framework compiler tools or build on Windows with .NET Framework available."
}

$outputDirectory = Split-Path -Parent $resolvedOutputPath
if (-not (Test-Path -LiteralPath $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory | Out-Null
}

$arguments = @(
    "/nologo",
    "/target:winexe",
    "/optimize+",
    "/reference:System.Windows.Forms.dll",
    "/out:$resolvedOutputPath"
)

if (Test-Path -LiteralPath $iconPath) {
    $arguments += "/win32icon:$iconPath"
}

$arguments += $sourcePath

& $cscPath @arguments
if ($LASTEXITCODE -ne 0) {
    throw "Launcher build failed with exit code $LASTEXITCODE."
}

Write-Output "Built launcher: $resolvedOutputPath"
