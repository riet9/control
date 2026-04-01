param(
    [string]$Version = "v0.1.0"
)

$ErrorActionPreference = "Stop"

$projectRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
$releaseRoot = Join-Path $projectRoot "release"
$packageName = "ScreenTimeTracker-portable-$Version"
$packageDirectory = Join-Path $releaseRoot $packageName
$zipPath = Join-Path $releaseRoot ($packageName + ".zip")

& (Join-Path $PSScriptRoot "build-launcher.ps1") -OutputPath (Join-Path $projectRoot "ScreenTimeTracker.exe")

Remove-Item -LiteralPath $packageDirectory -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath $zipPath -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Path $packageDirectory -Force | Out-Null

$rootFiles = @(
    "ScreenTimeTracker.ps1",
    "ScreenTimeTracker.exe",
    "start-tracker.vbs",
    "start-tracker.bat",
    "settings.json",
    "rules.json",
    "README.md",
    "LICENSE"
)

foreach ($file in $rootFiles) {
    $sourcePath = Join-Path $projectRoot $file
    if (-not (Test-Path -LiteralPath $sourcePath)) {
        throw "Required release file is missing: $file"
    }

    Copy-Item -LiteralPath $sourcePath -Destination $packageDirectory -Force
}

$directoryItems = @(
    "browser-extension",
    "Image"
)

foreach ($directoryItem in $directoryItems) {
    $sourcePath = Join-Path $projectRoot $directoryItem
    if (-not (Test-Path -LiteralPath $sourcePath)) {
        throw "Required release directory is missing: $directoryItem"
    }

    Copy-Item -LiteralPath $sourcePath -Destination $packageDirectory -Recurse -Force
}

Compress-Archive -Path (Join-Path $packageDirectory "*") -DestinationPath $zipPath

Write-Output ("Portable release ready: {0}" -f $zipPath)
