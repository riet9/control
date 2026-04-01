param(
    [switch]$SelfTest,
    [switch]$StartMinimized,
    [switch]$MigrateStorage,
    [switch]$RebuildSummaryCache
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms.DataVisualization
[System.Windows.Forms.Application]::EnableVisualStyles()

if (-not ("ScreenTime.NativeMethods" -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Text;

namespace ScreenTime
{
    public static class NativeMethods
    {
        [StructLayout(LayoutKind.Sequential)]
        public struct LASTINPUTINFO
        {
            public uint cbSize;
            public uint dwTime;
        }

        [DllImport("user32.dll")]
        public static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        public static extern uint GetTickCount();
    }
}
"@
}

if (-not ("ScreenTime.ListViewItemComparer" -as [type])) {
    Add-Type -ReferencedAssemblies @("System.Windows.Forms", "System") -TypeDefinition @"
using System;
using System.Collections;
using System.Globalization;
using System.Windows.Forms;

namespace ScreenTime
{
    public class ListViewItemComparer : IComparer
    {
        private readonly int _column;
        private readonly SortOrder _order;

        public ListViewItemComparer(int column, SortOrder order)
        {
            _column = column;
            _order = order;
        }

        public int Compare(object x, object y)
        {
            ListViewItem itemX = x as ListViewItem;
            ListViewItem itemY = y as ListViewItem;

            string valueX = GetValue(itemX);
            string valueY = GetValue(itemY);
            string mode = GetMode(itemX);
            int result = CompareValues(valueX, valueY, mode);

            if (_order == SortOrder.Descending)
            {
                result *= -1;
            }

            return result;
        }

        private string GetValue(ListViewItem item)
        {
            if (item == null)
            {
                return string.Empty;
            }

            if (_column < item.SubItems.Count)
            {
                return item.SubItems[_column].Text ?? string.Empty;
            }

            return item.Text ?? string.Empty;
        }

        private string GetMode(ListViewItem item)
        {
            if (item == null || item.ListView == null || item.ListView.Tag == null)
            {
                return "string";
            }

            string raw = item.ListView.Tag.ToString();
            string[] parts = raw.Split('|');
            if (_column >= 0 && _column < parts.Length)
            {
                return parts[_column];
            }

            return "string";
        }

        private int CompareValues(string left, string right, string mode)
        {
            switch ((mode ?? "string").ToLowerInvariant())
            {
                case "time":
                    return ParseDuration(left).CompareTo(ParseDuration(right));
                case "number":
                    return ParseNumber(left).CompareTo(ParseNumber(right));
                case "percent":
                    return ParseNumber(left.Replace("%", "")).CompareTo(ParseNumber(right.Replace("%", "")));
                case "date":
                    return ParseDate(left).CompareTo(ParseDate(right));
                default:
                    return string.Compare(left, right, StringComparison.CurrentCultureIgnoreCase);
            }
        }

        private TimeSpan ParseDuration(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return TimeSpan.Zero;
            }

            TimeSpan parsed;
            if (TimeSpan.TryParse(value, CultureInfo.InvariantCulture, out parsed))
            {
                return parsed;
            }

            string[] parts = value.Split(':');
            if (parts.Length == 2)
            {
                int minutes;
                int seconds;
                if (int.TryParse(parts[0], out minutes) && int.TryParse(parts[1], out seconds))
                {
                    return new TimeSpan(0, minutes, seconds);
                }
            }

            return TimeSpan.Zero;
        }

        private double ParseNumber(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return 0;
            }

            double parsed;
            if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out parsed))
            {
                return parsed;
            }

            if (double.TryParse(value, NumberStyles.Any, CultureInfo.CurrentCulture, out parsed))
            {
                return parsed;
            }

            return 0;
        }

        private DateTime ParseDate(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return DateTime.MinValue;
            }

            DateTime parsed;
            if (DateTime.TryParseExact(value, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out parsed))
            {
                return parsed;
            }

            if (DateTime.TryParse(value, out parsed))
            {
                return parsed;
            }

            return DateTime.MinValue;
        }
    }
}
"@
}

$script:ScriptPath = $MyInvocation.MyCommand.Path
$script:AppRoot = Split-Path -Parent $script:ScriptPath
$script:DataDirectory = Join-Path $script:AppRoot "data"
$script:UsageDaysDirectory = Join-Path $script:DataDirectory "days"
$script:UsageSummaryPath = Join-Path $script:DataDirectory "usage-summary.json"
$script:ExportsDirectory = Join-Path $script:DataDirectory "exports"
$script:BackupsDirectory = Join-Path $script:DataDirectory "backups"
$script:RulesPath = Join-Path $script:AppRoot "rules.json"
$script:SettingsPath = Join-Path $script:AppRoot "settings.json"
$script:UsageDataPath = Join-Path $script:DataDirectory "usage-data.json"
$script:BrowserActivityPath = Join-Path $script:DataDirectory "browser-activity.json"
$script:BrowserExtensionDirectory = Join-Path $script:AppRoot "browser-extension"
$script:ImageDirectory = Join-Path $script:AppRoot "Image"
$script:AppIconPath = Join-Path $script:ImageDirectory "tracker_time.ico"
$script:VbsLauncherPath = Join-Path $script:AppRoot "start-tracker.vbs"
$script:StartupLogPath = Join-Path $script:DataDirectory "startup-error.log"
$script:ReopenSignalPath = Join-Path $script:DataDirectory "reopen.signal"
$script:WindowTitle = "Screen Time Tracker"
$script:BrowserBridgePort = 38945
$script:AutoStartRegPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
$script:AutoStartValueName = "ScreenTimeTracker"
$script:BrowserProcesses = @("chrome", "msedge", "brave", "opera", "firefox")
$script:AppMutex = $null
$script:AppMutexName = "Local\ScreenTimeTracker.SingleInstance"
$script:ClassificationCache = @{}
$script:UsageSchemaVersion = 2
$script:MaxBackupFilesPerPrefix = 14
$script:CategoryDefinitions = @{}
$script:CategoryChoices = @("study", "browser_fun", "socials", "other")
$script:DayStatsNormalizationCache = @{}
$script:UsageDateIndex = @{}
$script:UsageSummaryCache = @{}

function ConvertTo-PlainData {
    param(
        [Parameter(ValueFromPipeline = $true)]
        $InputObject
    )

    if ($null -eq $InputObject) {
        return $null
    }

    if ($InputObject -is [System.Collections.IDictionary]) {
        $result = @{}
        foreach ($key in $InputObject.Keys) {
            $result[$key] = ConvertTo-PlainData $InputObject[$key]
        }
        return $result
    }

    if (($InputObject -is [System.Collections.IEnumerable]) -and -not ($InputObject -is [string])) {
        $items = @()
        foreach ($item in $InputObject) {
            $items += ,(ConvertTo-PlainData $item)
        }
        return $items
    }

    if ($InputObject -is [pscustomobject] -or $InputObject -is [System.Management.Automation.PSObject]) {
        $result = @{}
        foreach ($property in $InputObject.PSObject.Properties) {
            $result[$property.Name] = ConvertTo-PlainData $property.Value
        }
        return $result
    }

    return $InputObject
}

function Get-DayStatsCacheKey {
    param(
        [hashtable]$UsageData,
        [string]$DateKey
    )

    if (-not $UsageData.ContainsKey($DateKey)) {
        return ""
    }

    $usageHash = [string][System.Runtime.CompilerServices.RuntimeHelpers]::GetHashCode($UsageData)
    $day = $UsageData[$DateKey]
    $dayHash = [string][System.Runtime.CompilerServices.RuntimeHelpers]::GetHashCode($day)
    $sessionsHash = "0"
    $activitiesHash = "0"
    $totalsHash = "0"
    if ($day -is [System.Collections.IDictionary]) {
        if ($day.ContainsKey("sessions") -and $null -ne $day["sessions"]) {
            $sessionsHash = [string][System.Runtime.CompilerServices.RuntimeHelpers]::GetHashCode($day["sessions"])
        }
        if ($day.ContainsKey("activities") -and $null -ne $day["activities"]) {
            $activitiesHash = [string][System.Runtime.CompilerServices.RuntimeHelpers]::GetHashCode($day["activities"])
        }
        if ($day.ContainsKey("totals") -and $null -ne $day["totals"]) {
            $totalsHash = [string][System.Runtime.CompilerServices.RuntimeHelpers]::GetHashCode($day["totals"])
        }
    }

    return "{0}|{1}|{2}|{3}|{4}|{5}" -f $usageHash, $DateKey, $dayHash, $sessionsHash, $activitiesHash, $totalsHash
}

function Get-ParentCategoryDefinitions {
    return @(
        [pscustomobject]@{ key = "study"; label = "Study"; parent = "study"; builtIn = $true },
        [pscustomobject]@{ key = "browser_fun"; label = "Browser fun / manga"; parent = "browser_fun"; builtIn = $true },
        [pscustomobject]@{ key = "socials"; label = "Social media"; parent = "socials"; builtIn = $true },
        [pscustomobject]@{ key = "other"; label = "Other"; parent = "other"; builtIn = $true }
    )
}

function Get-ParentCategoryKeys {
    return @((Get-ParentCategoryDefinitions) | ForEach-Object { [string]$_.key })
}

function Normalize-CategoryKey {
    param(
        [AllowNull()]
        [string]$Key
    )

    $value = Normalize-Text $Key
    $value = $value -replace "[\s\-]+", "_"
    $value = $value -replace "[^a-z0-9_]", ""
    $value = $value.Trim("_")
    return $value
}

function Convert-CategoryKeyToLabel {
    param(
        [string]$Key
    )

    if ([string]::IsNullOrWhiteSpace($Key)) {
        return "Category"
    }

    $words = @()
    foreach ($part in (($Key -replace "_", " ") -split "\s+")) {
        if ([string]::IsNullOrWhiteSpace($part)) {
            continue
        }

        $word = $part.Substring(0, 1).ToUpperInvariant()
        if ($part.Length -gt 1) {
            $word += $part.Substring(1)
        }
        $words += ,$word
    }

    if ($words.Count -eq 0) {
        return "Category"
    }

    return ($words -join " ")
}

function Normalize-CustomCategories {
    param(
        [System.Array]$Categories
    )

    $normalized = @()
    $seen = @{}
    $parentKeys = @(Get-ParentCategoryKeys)
    foreach ($category in @($Categories)) {
        $plain = ConvertTo-PlainData $category
        if ($plain -isnot [System.Collections.IDictionary]) {
            continue
        }

        $key = Normalize-CategoryKey ([string](Get-RuleField -Rule $plain -Name "key" -Default (Get-RuleField -Rule $plain -Name "Key" -Default "")))
        if ([string]::IsNullOrWhiteSpace($key) -or ($parentKeys -contains $key) -or $seen.ContainsKey($key)) {
            continue
        }

        $label = [string](Get-RuleField -Rule $plain -Name "label" -Default (Get-RuleField -Rule $plain -Name "Label" -Default ""))
        if ([string]::IsNullOrWhiteSpace($label)) {
            $label = Convert-CategoryKeyToLabel -Key $key
        }

        $parent = Normalize-CategoryKey ([string](Get-RuleField -Rule $plain -Name "parent" -Default (Get-RuleField -Rule $plain -Name "Parent" -Default "other")))
        if (-not ($parentKeys -contains $parent)) {
            $parent = "other"
        }

        $seen[$key] = $true
        $normalized += ,[pscustomobject]@{
            key = $key
            label = $label.Trim()
            parent = $parent
        }
    }

    return @($normalized)
}

function Sync-CategoryRegistry {
    param(
        [hashtable]$Settings
    )

    $definitions = @{}
    $choices = @()
    foreach ($parentCategory in (Get-ParentCategoryDefinitions)) {
        $definitions[[string]$parentCategory.key] = $parentCategory
        $choices += ,[string]$parentCategory.key
    }

    $customCategories = @()
    if ($null -ne $Settings -and $Settings.ContainsKey("categories")) {
        $customCategories = @(Normalize-CustomCategories -Categories $Settings.categories)
    }

    foreach ($category in $customCategories) {
        $definitions[[string]$category.key] = [pscustomobject]@{
            key = [string]$category.key
            label = [string]$category.label
            parent = [string]$category.parent
            builtIn = $false
        }
        $choices += ,[string]$category.key
    }

    $script:CategoryDefinitions = $definitions
    $script:CategoryChoices = @($choices)
}

function Test-CategoryExists {
    param(
        [string]$Category
    )

    return $script:CategoryDefinitions.ContainsKey((Normalize-CategoryKey $Category))
}

function Get-CategoryParentKey {
    param(
        [string]$Category
    )

    $key = Normalize-CategoryKey $Category
    if ($script:CategoryDefinitions.ContainsKey($key)) {
        return [string]$script:CategoryDefinitions[$key].parent
    }

    return "other"
}

function Get-TrackedCategoryKeysForTotals {
    param(
        [string]$Category
    )

    $categoryKey = Normalize-CategoryKey $Category
    if ([string]::IsNullOrWhiteSpace($categoryKey)) {
        $categoryKey = "other"
    }
    if (-not (Test-CategoryExists -Category $categoryKey)) {
        $categoryKey = "other"
    }

    $keys = @($categoryKey)
    $parentKey = Get-CategoryParentKey -Category $categoryKey
    if ($parentKey -ne $categoryKey) {
        $keys += ,$parentKey
    }

    return @($keys)
}

function Add-SecondsToCategoryTotals {
    param(
        $Totals,
        [string]$Category,
        [double]$Seconds
    )

    foreach ($key in (Get-TrackedCategoryKeysForTotals -Category $Category)) {
        if (-not $Totals.ContainsKey($key)) {
            $Totals[$key] = 0.0
        }
        $Totals[$key] += $Seconds
    }
}

function Remove-SecondsFromCategoryTotals {
    param(
        $Totals,
        [string]$Category,
        [double]$Seconds
    )

    foreach ($key in (Get-TrackedCategoryKeysForTotals -Category $Category)) {
        if (-not $Totals.ContainsKey($key)) {
            $Totals[$key] = 0.0
        }
        $Totals[$key] = [math]::Max(0, [double]$Totals[$key] - $Seconds)
    }
}

function Ensure-Storage {
    if (-not (Test-Path -LiteralPath $script:DataDirectory)) {
        New-Item -ItemType Directory -Path $script:DataDirectory | Out-Null
    }

    if (-not (Test-Path -LiteralPath $script:UsageDaysDirectory)) {
        New-Item -ItemType Directory -Path $script:UsageDaysDirectory | Out-Null
    }

    if (-not (Test-Path -LiteralPath $script:ExportsDirectory)) {
        New-Item -ItemType Directory -Path $script:ExportsDirectory | Out-Null
    }

    if (-not (Test-Path -LiteralPath $script:BackupsDirectory)) {
        New-Item -ItemType Directory -Path $script:BackupsDirectory | Out-Null
    }

    if (-not (Test-Path -LiteralPath $script:UsageDataPath)) {
        "{}" | Set-Content -LiteralPath $script:UsageDataPath -Encoding UTF8
    }

    if (-not (Test-Path -LiteralPath $script:UsageSummaryPath)) {
        "{}" | Set-Content -LiteralPath $script:UsageSummaryPath -Encoding UTF8
    }

    if (-not (Test-Path -LiteralPath $script:BrowserActivityPath)) {
        "{}" | Set-Content -LiteralPath $script:BrowserActivityPath -Encoding UTF8
    }
}

function Read-JsonFile {
    param(
        [string]$Path,
        $Fallback
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return $Fallback
    }

    $raw = Get-Content -LiteralPath $Path -Raw
    if ([string]::IsNullOrWhiteSpace($raw)) {
        return $Fallback
    }

    try {
        return ConvertTo-PlainData (ConvertFrom-Json -InputObject $raw)
    }
    catch {
        return $Fallback
    }
}

function Save-JsonFile {
    param(
        [string]$Path,
        $Data
    )

    $json = ConvertTo-Json -InputObject (ConvertTo-PlainData $Data) -Depth 10 -Compress
    Set-Content -LiteralPath $Path -Value $json -Encoding UTF8
}

function Trim-BackupFiles {
    param(
        [string]$Directory,
        [string]$Prefix,
        [int]$Keep = $script:MaxBackupFilesPerPrefix
    )

    if ($Keep -lt 1 -or -not (Test-Path -LiteralPath $Directory)) {
        return
    }

    $files = @(Get-ChildItem -LiteralPath $Directory -File -Filter ("{0}-*" -f $Prefix) | Sort-Object -Property LastWriteTime -Descending)
    if ($files.Count -le $Keep) {
        return
    }

    foreach ($file in ($files | Select-Object -Skip $Keep)) {
        Remove-Item -LiteralPath $file.FullName -Force -ErrorAction SilentlyContinue
    }
}

function Ensure-DailyBackup {
    param(
        [string]$SourcePath,
        [string]$Prefix
    )

    if (-not (Test-Path -LiteralPath $SourcePath)) {
        return
    }

    Ensure-Storage
    $dayStamp = Get-Date -Format "yyyy-MM-dd"
    $extension = [System.IO.Path]::GetExtension($SourcePath)
    $backupPath = Join-Path $script:BackupsDirectory ("{0}-{1}{2}" -f $Prefix, $dayStamp, $extension)
    if (Test-Path -LiteralPath $backupPath) {
        Trim-BackupFiles -Directory $script:BackupsDirectory -Prefix $Prefix
        return
    }

    Copy-Item -LiteralPath $SourcePath -Destination $backupPath -Force
    Trim-BackupFiles -Directory $script:BackupsDirectory -Prefix $Prefix
}

function New-UsageDataEnvelope {
    param(
        [hashtable]$UsageData
    )

    return @{
        schemaVersion = $script:UsageSchemaVersion
        savedAt = (Get-Date).ToString("o")
        days = $UsageData
    }
}

function New-UsageStorageManifest {
    param(
        [int]$DateCount = 0
    )

    return @{
        schemaVersion = $script:UsageSchemaVersion
        savedAt = (Get-Date).ToString("o")
        storageMode = "day-files"
        dateCount = [int]$DateCount
    }
}

function New-UsageSummaryEnvelope {
    param(
        [hashtable]$SummaryDays
    )

    return @{
        schemaVersion = $script:UsageSchemaVersion
        savedAt = (Get-Date).ToString("o")
        days = $SummaryDays
    }
}

function Get-UsageDayPath {
    param(
        [string]$DateKey
    )

    return (Join-Path $script:UsageDaysDirectory ("{0}.json" -f $DateKey))
}

function New-DaySummary {
    param(
        [string]$DateKey
    )

    return @{
        date = $DateKey
        total = 0.0
        study = 0.0
        browser_fun = 0.0
        socials = 0.0
        other = 0.0
        processes = @{}
        exactCategories = @{}
    }
}

function Build-DaySummaryFromDayStats {
    param(
        [string]$DateKey,
        [hashtable]$DayStats
    )

    $summary = New-DaySummary -DateKey $DateKey
    if ($null -eq $DayStats) {
        return $summary
    }

    $summary.total = [double]$DayStats.totals.total
    $summary.study = [double]$DayStats.totals.study
    $summary.browser_fun = [double]$DayStats.totals.browser_fun
    $summary.socials = [double]$DayStats.totals.socials
    $summary.other = [math]::Max(0.0, [double]$summary.total - [double]$summary.study - [double]$summary.browser_fun - [double]$summary.socials)

    foreach ($activity in $DayStats.activities.Values) {
        $process = [string]$activity.process
        if ([string]::IsNullOrWhiteSpace($process)) {
            $process = "Unknown"
        }

        if (-not $summary.processes.ContainsKey($process)) {
            $summary.processes[$process] = @{
                process = $process
                category = [string]$activity.category
                seconds = 0.0
            }
        }
        $summary.processes[$process].seconds += [double]$activity.seconds

        $category = Normalize-CategoryKey ([string]$activity.category)
        if ([string]::IsNullOrWhiteSpace($category) -or -not (Test-CategoryExists -Category $category)) {
            $category = "other"
        }

        if (-not $summary.exactCategories.ContainsKey($category)) {
            $summary.exactCategories[$category] = @{
                category = $category
                parent = Get-CategoryParentKey -Category $category
                seconds = 0.0
            }
        }
        $summary.exactCategories[$category].seconds += [double]$activity.seconds
    }

    return $summary
}

function Refresh-UsageDateIndex {
    Ensure-Storage
    $script:UsageDateIndex = @{}
    foreach ($file in @(Get-ChildItem -LiteralPath $script:UsageDaysDirectory -File -Filter "*.json" -ErrorAction SilentlyContinue)) {
        if ($file.BaseName -match '^\d{4}-\d{2}-\d{2}$') {
            $script:UsageDateIndex[$file.BaseName] = $true
        }
    }
}

function Load-UsageSummaryCache {
    Ensure-Storage
    $loaded = Read-JsonFile -Path $script:UsageSummaryPath -Fallback @{}
    $days = $loaded
    if ($loaded -is [System.Collections.IDictionary] -and $loaded.ContainsKey("days")) {
        $days = $loaded["days"]
    }

    if ($days -isnot [System.Collections.IDictionary]) {
        $script:UsageSummaryCache = @{}
        return $script:UsageSummaryCache
    }

    $normalized = @{}
    foreach ($dateKey in @($days.Keys)) {
        if ([string]$dateKey -notmatch '^\d{4}-\d{2}-\d{2}$') {
            continue
        }

        $entry = ConvertTo-PlainData $days[$dateKey]
        if ($entry -isnot [System.Collections.IDictionary]) {
            continue
        }

        if (-not $entry.ContainsKey("processes") -or $entry.processes -isnot [System.Collections.IDictionary]) {
            $entry["processes"] = @{}
        }
        if (-not $entry.ContainsKey("exactCategories") -or $entry.exactCategories -isnot [System.Collections.IDictionary]) {
            $entry["exactCategories"] = @{}
        }
        $normalized[[string]$dateKey] = $entry
    }

    $script:UsageSummaryCache = $normalized
    return $script:UsageSummaryCache
}

function Save-UsageSummaryCache {
    Ensure-Storage
    Save-JsonFile -Path $script:UsageSummaryPath -Data (New-UsageSummaryEnvelope -SummaryDays $script:UsageSummaryCache)
}

function Update-UsageSummaryForDay {
    param(
        [string]$DateKey,
        [hashtable]$DayStats
    )

    $script:UsageSummaryCache[$DateKey] = Build-DaySummaryFromDayStats -DateKey $DateKey -DayStats $DayStats
}

function Get-DaySummary {
    param(
        [hashtable]$UsageData,
        [string]$DateKey
    )

    if ($script:UsageSummaryCache.ContainsKey($DateKey)) {
        return $script:UsageSummaryCache[$DateKey]
    }

    if ($UsageData.ContainsKey($DateKey) -or (Test-UsageDayExists -DateKey $DateKey)) {
        $day = Get-DayStats -UsageData $UsageData -DateKey $DateKey
        $summary = Build-DaySummaryFromDayStats -DateKey $DateKey -DayStats $day
        $script:UsageSummaryCache[$DateKey] = $summary
        return $summary
    }

    return (New-DaySummary -DateKey $DateKey)
}

function Test-UsageDayExists {
    param(
        [string]$DateKey
    )

    if ([string]::IsNullOrWhiteSpace($DateKey)) {
        return $false
    }

    if ($script:UsageDateIndex.ContainsKey($DateKey)) {
        return $true
    }

    $path = Get-UsageDayPath -DateKey $DateKey
    if (Test-Path -LiteralPath $path) {
        $script:UsageDateIndex[$DateKey] = $true
        return $true
    }

    return $false
}

function Save-UsageStorageManifest {
    param(
        [int]$DateCount = -1
    )

    if ($DateCount -lt 0) {
        $DateCount = @($script:UsageDateIndex.Keys).Count
    }

    Save-JsonFile -Path $script:UsageDataPath -Data (New-UsageStorageManifest -DateCount $DateCount)
}

function Load-UsageDayFile {
    param(
        [string]$DateKey
    )

    $path = Get-UsageDayPath -DateKey $DateKey
    $loaded = Read-JsonFile -Path $path -Fallback $null
    if ($null -eq $loaded -or $loaded -isnot [System.Collections.IDictionary]) {
        return $null
    }

    $usage = @{ $DateKey = $loaded }
    return (Get-DayStats -UsageData $usage -DateKey $DateKey)
}

function Save-UsageDayFile {
    param(
        [string]$DateKey,
        [hashtable]$DayStats
    )

    if ([string]::IsNullOrWhiteSpace($DateKey) -or $null -eq $DayStats) {
        return
    }

    Ensure-Storage
    $path = Get-UsageDayPath -DateKey $DateKey
    Ensure-DailyBackup -SourcePath $path -Prefix ("usage-day-{0}" -f $DateKey)
    Save-JsonFile -Path $path -Data $DayStats
    $script:UsageDateIndex[$DateKey] = $true
}

function Get-LegacyUsageDaysFromFile {
    $loaded = Read-JsonFile -Path $script:UsageDataPath -Fallback @{}
    if ($loaded -isnot [System.Collections.IDictionary]) {
        return @{}
    }

    if ($loaded.ContainsKey("storageMode")) {
        return @{}
    }

    $days = $loaded
    if ($loaded.ContainsKey("days")) {
        $days = $loaded["days"]
    }

    if ($days -isnot [System.Collections.IDictionary]) {
        return @{}
    }

    $result = @{}
    foreach ($dateKey in @($days.Keys)) {
        if ([string]$dateKey -match '^\d{4}-\d{2}-\d{2}$') {
            $result[[string]$dateKey] = $days[$dateKey]
        }
    }

    return $result
}

function Migrate-LegacyUsageDataToDayFiles {
    Ensure-Storage

    $existingDayFiles = @(Get-ChildItem -LiteralPath $script:UsageDaysDirectory -File -Filter "*.json" -ErrorAction SilentlyContinue)
    if ($existingDayFiles.Count -gt 0) {
        Refresh-UsageDateIndex
        Save-UsageStorageManifest
        return $true
    }

    $days = Get-LegacyUsageDaysFromFile
    if (@($days.Keys).Count -eq 0) {
        Save-UsageStorageManifest -DateCount 0
        return $false
    }

    Ensure-DailyBackup -SourcePath $script:UsageDataPath -Prefix "usage-data-legacy"
    $script:UsageSummaryCache = @{}
    foreach ($dateKey in @($days.Keys)) {
        $usage = @{ ([string]$dateKey) = $days[$dateKey] }
        $normalizedDay = Get-DayStats -UsageData $usage -DateKey ([string]$dateKey)
        Save-UsageDayFile -DateKey ([string]$dateKey) -DayStats $normalizedDay
        Update-UsageSummaryForDay -DateKey ([string]$dateKey) -DayStats $normalizedDay
    }

    Refresh-UsageDateIndex
    Save-UsageStorageManifest
    Save-UsageSummaryCache
    return $true
}

function Get-AllStoredUsageData {
    param(
        [hashtable]$UsageData
    )

    Refresh-UsageDateIndex
    $snapshot = @{}
    foreach ($dateKey in @($script:UsageDateIndex.Keys | Sort-Object)) {
        $loadedDay = Load-UsageDayFile -DateKey ([string]$dateKey)
        if ($null -ne $loadedDay) {
            $snapshot[[string]$dateKey] = $loadedDay
        }
    }

    if ($null -ne $UsageData) {
        foreach ($dateKey in @($UsageData.Keys)) {
            if ([string]::IsNullOrWhiteSpace([string]$dateKey)) {
                continue
            }

            $snapshot[[string]$dateKey] = Get-DayStats -UsageData $UsageData -DateKey ([string]$dateKey)
        }
    }

    return $snapshot
}

function Rebuild-UsageSummaryCache {
    Ensure-Storage
    Refresh-UsageDateIndex
    $script:UsageSummaryCache = @{}

    foreach ($file in @(Get-ChildItem -LiteralPath $script:UsageDaysDirectory -File -Filter "*.json" -ErrorAction SilentlyContinue | Sort-Object -Property BaseName)) {
        $dateKey = [string]$file.BaseName
        if ($dateKey -notmatch '^\d{4}-\d{2}-\d{2}$') {
            continue
        }

        $loaded = Read-JsonFile -Path $file.FullName -Fallback $null
        if ($null -eq $loaded -or $loaded -isnot [System.Collections.IDictionary]) {
            continue
        }

        $usage = @{ $dateKey = $loaded }
        $day = Get-DayStats -UsageData $usage -DateKey $dateKey
        Update-UsageSummaryForDay -DateKey $dateKey -DayStats $day
    }

    Save-UsageSummaryCache
    return @($script:UsageSummaryCache.Keys).Count
}

function Load-Settings {
    $fallback = @{
        idleThresholdSeconds = 300
        sampleIntervalSeconds = 1
        browserBridge = @{
            port = $script:BrowserBridgePort
            titleMatchToleranceHours = 12
        }
        notifications = @{
            warningPercent = 80
            hardLimitEnabled = $true
            hardLimitSnoozeMinutes = 5
            closeDistractingWindows = $true
            blockCooldownSeconds = 12
        }
        focusMode = @{
            defaultMinutes = 50
            promptCooldownSeconds = 20
            closeDistractingWindows = $true
        }
        autostart = @{
            enabled = $true
        }
        limits = @{
            total = 10800
            studyMin = 7200
            studyMax = 9000
            browser_fun = 1800
            socials = 1080
        }
        categories = @()
    }

    $loaded = Read-JsonFile -Path $script:SettingsPath -Fallback $fallback
    if ($null -eq $loaded) {
        return $fallback
    }

    if (-not $loaded.ContainsKey("browserBridge")) {
        $loaded["browserBridge"] = @{ port = $script:BrowserBridgePort; titleMatchToleranceHours = 12 }
    }
    if (-not $loaded.browserBridge.ContainsKey("port")) {
        $loaded.browserBridge["port"] = $script:BrowserBridgePort
    }
    if (-not $loaded.browserBridge.ContainsKey("titleMatchToleranceHours")) {
        $loaded.browserBridge["titleMatchToleranceHours"] = 12
    }
    if (-not $loaded.ContainsKey("notifications")) {
        $loaded["notifications"] = @{ warningPercent = 80; hardLimitEnabled = $true; hardLimitSnoozeMinutes = 5; closeDistractingWindows = $true; blockCooldownSeconds = 12 }
    }
    if (-not $loaded.notifications.ContainsKey("warningPercent")) {
        $loaded.notifications["warningPercent"] = 80
    }
    if (-not $loaded.notifications.ContainsKey("hardLimitEnabled")) {
        $loaded.notifications["hardLimitEnabled"] = $true
    }
    if (-not $loaded.notifications.ContainsKey("hardLimitSnoozeMinutes")) {
        $loaded.notifications["hardLimitSnoozeMinutes"] = 5
    }
    if (-not $loaded.notifications.ContainsKey("closeDistractingWindows")) {
        $loaded.notifications["closeDistractingWindows"] = $true
    }
    if (-not $loaded.notifications.ContainsKey("blockCooldownSeconds")) {
        $loaded.notifications["blockCooldownSeconds"] = 12
    }
    if (-not $loaded.ContainsKey("focusMode")) {
        $loaded["focusMode"] = @{ defaultMinutes = 50; promptCooldownSeconds = 20; closeDistractingWindows = $true }
    }
    if (-not $loaded.focusMode.ContainsKey("defaultMinutes")) {
        $loaded.focusMode["defaultMinutes"] = 50
    }
    if (-not $loaded.focusMode.ContainsKey("promptCooldownSeconds")) {
        $loaded.focusMode["promptCooldownSeconds"] = 20
    }
    if (-not $loaded.focusMode.ContainsKey("closeDistractingWindows")) {
        $loaded.focusMode["closeDistractingWindows"] = $true
    }
    if (-not $loaded.ContainsKey("autostart")) {
        $loaded["autostart"] = @{ enabled = $true }
    }
    if (-not $loaded.autostart.ContainsKey("enabled")) {
        $loaded.autostart["enabled"] = $true
    }

    if (-not $loaded.ContainsKey("categories")) {
        $loaded["categories"] = @()
    }
    $loaded["categories"] = @(Normalize-CustomCategories -Categories $loaded.categories)
    Sync-CategoryRegistry -Settings $loaded

    return $loaded
}

function Load-Rules {
    $fallback = @()
    $loaded = Read-JsonFile -Path $script:RulesPath -Fallback $fallback
    if ($loaded -isnot [System.Array]) {
        $loaded = @($loaded)
    }

    Clear-ClassificationCache
    return @(Normalize-Rules -Rules $loaded)
}

function Load-UsageData {
    Ensure-Storage
    Refresh-UsageDateIndex
    [void](Load-UsageSummaryCache)
    if (@($script:UsageDateIndex.Keys).Count -eq 0) {
        [void](Migrate-LegacyUsageDataToDayFiles)
        Refresh-UsageDateIndex
        [void](Load-UsageSummaryCache)
    }

    $loadedDays = @{}
    $currentDateKey = Get-DateKey
    if (Test-UsageDayExists -DateKey $currentDateKey) {
        $currentDay = Load-UsageDayFile -DateKey $currentDateKey
        if ($null -ne $currentDay) {
            $loadedDays[$currentDateKey] = $currentDay
        }
    }

    return $loadedDays
}

function Save-UsageData {
    param(
        [hashtable]$UsageData,
        [string[]]$DateKeys = @()
    )

    $targetDateKeys = @($DateKeys | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Select-Object -Unique)
    if ($targetDateKeys.Count -eq 0) {
        $targetDateKeys = @($UsageData.Keys | Sort-Object)
    }

    foreach ($dateKey in $targetDateKeys) {
        if (-not $UsageData.ContainsKey($dateKey)) {
            continue
        }

        $day = Get-DayStats -UsageData $UsageData -DateKey ([string]$dateKey)
        Save-UsageDayFile -DateKey ([string]$dateKey) -DayStats $day
        Update-UsageSummaryForDay -DateKey ([string]$dateKey) -DayStats $day
    }

    Save-UsageStorageManifest
    Save-UsageSummaryCache
}

function Save-Settings {
    param(
        [hashtable]$Settings
    )

    if (-not $Settings.ContainsKey("categories")) {
        $Settings["categories"] = @()
    }
    $Settings["categories"] = @(Normalize-CustomCategories -Categories $Settings.categories)
    Sync-CategoryRegistry -Settings $Settings
    Save-JsonFile -Path $script:SettingsPath -Data $Settings
}

function Save-Rules {
    param(
        [System.Array]$Rules
    )

    $normalized = @(Normalize-Rules -Rules $Rules)
    Save-JsonFile -Path $script:RulesPath -Data $normalized
    Clear-ClassificationCache
}

function Clear-ClassificationCache {
    $script:ClassificationCache = @{}
}

function Mark-UsageDateDirty {
    param(
        [hashtable]$State,
        [string]$DateKey
    )

    if ($null -eq $State -or [string]::IsNullOrWhiteSpace($DateKey)) {
        return
    }

    if (-not $State.ContainsKey("DirtyDateKeys") -or $null -eq $State.DirtyDateKeys) {
        $State["DirtyDateKeys"] = @{}
    }

    $State.DirtyDateKeys[$DateKey] = $true
}

function Get-RuleCategoryChoices {
    return @($script:CategoryChoices)
}

function Get-RuleTargetChoices {
    return @("process", "window", "title", "url", "domain", "either")
}

function Get-RuleMatchModeChoices {
    return @("contains", "exact", "starts_with", "ends_with", "regex")
}

function Normalize-Rule {
    param(
        $Rule
    )

    if ($null -eq $Rule) {
        return $null
    }

    $patterns = @()
    foreach ($pattern in @(Get-RuleField -Rule $Rule -Name "Patterns" -Default @())) {
        $trimmed = [string]$pattern
        if (-not [string]::IsNullOrWhiteSpace($trimmed)) {
            $patterns += ,$trimmed.Trim()
        }
    }

    if ($patterns.Count -eq 0) {
        $trimmed = [string](Get-RuleField -Rule $Rule -Name "Pattern" -Default "")
        if (-not [string]::IsNullOrWhiteSpace($trimmed)) {
            $patterns += ,$trimmed.Trim()
        }
    }

    $categoryValue = Normalize-CategoryKey ([string](Get-RuleField -Rule $Rule -Name "Category" -Default "other"))
    if (-not (Test-CategoryExists -Category $categoryValue)) {
        $categoryValue = "other"
    }

    $targetValue = [string](Get-RuleField -Rule $Rule -Name "Target" -Default "either")
    if ([string]::IsNullOrWhiteSpace($targetValue)) {
        $targetValue = "either"
    }

    $matchModeValue = [string](Get-RuleField -Rule $Rule -Name "MatchMode" -Default "contains")
    if ([string]::IsNullOrWhiteSpace($matchModeValue)) {
        $matchModeValue = "contains"
    }

    return [pscustomobject]@{
        Name = [string](Get-RuleField -Rule $Rule -Name "Name" -Default "")
        Category = $categoryValue
        Target = $targetValue
        Patterns = $patterns
        Priority = [int](Get-RuleField -Rule $Rule -Name "Priority" -Default 100)
        MatchMode = $matchModeValue
        Enabled = [bool](Get-RuleField -Rule $Rule -Name "Enabled" -Default $true)
    }
}

function Normalize-Rules {
    param(
        [System.Array]$Rules
    )

    $normalized = @()
    foreach ($rule in @($Rules)) {
        $plainRule = Normalize-Rule -Rule $rule
        if ($null -eq $plainRule) {
            continue
        }

        if ([string]::IsNullOrWhiteSpace([string]$plainRule.Name) -or [string]::IsNullOrWhiteSpace([string]$plainRule.Category) -or [string]::IsNullOrWhiteSpace([string]$plainRule.Target) -or @($plainRule.Patterns).Count -eq 0) {
            continue
        }

        $normalized += ,$plainRule
    }

    return $normalized | Sort-Object -Property Priority, Name
}

function Get-AppIcon {
    if (Test-Path -LiteralPath $script:AppIconPath) {
        try {
            $stream = [System.IO.File]::OpenRead($script:AppIconPath)
            try {
                return New-Object System.Drawing.Icon($stream)
            }
            finally {
                $stream.Dispose()
            }
        }
        catch {
        }
    }

    return [System.Drawing.SystemIcons]::Information
}

function Show-StartupFailure {
    param(
        [System.Exception]$Exception
    )

    Ensure-Storage
    $message = $Exception.ToString()
    Set-Content -LiteralPath $script:StartupLogPath -Value $message -Encoding UTF8

    try {
        [System.Windows.Forms.MessageBox]::Show(
            "The tracker could not start.`n`n$($Exception.Message)`n`nDetails were saved to:`n$script:StartupLogPath",
            "Screen Time Tracker",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
    catch {
    }
}

function Acquire-AppMutex {
    $createdNew = $false
    $script:AppMutex = New-Object System.Threading.Mutex($true, $script:AppMutexName, [ref]$createdNew)
    if (-not $createdNew) {
        try {
            $script:AppMutex.Dispose()
        }
        catch {
        }
        $script:AppMutex = $null
        return $false
    }

    return $true
}

function Release-AppMutex {
    if ($null -eq $script:AppMutex) {
        return
    }

    try {
        $script:AppMutex.ReleaseMutex() | Out-Null
    }
    catch {
    }

    try {
        $script:AppMutex.Dispose()
    }
    catch {
    }

    $script:AppMutex = $null
}

function New-DayStats {
    param(
        [string]$DateKey
    )

    $totals = @{ total = 0.0 }
    foreach ($category in (Get-ParentCategoryKeys)) {
        $totals[$category] = 0.0
    }

    return @{
        date = $DateKey
        totals = $totals
        activities = @{}
        sessions = @()
    }
}

function Get-DayStats {
    param(
        [hashtable]$UsageData,
        [string]$DateKey
    )

    if (-not $UsageData.ContainsKey($DateKey)) {
        if (Test-UsageDayExists -DateKey $DateKey) {
            $loadedDay = Load-UsageDayFile -DateKey $DateKey
            if ($null -ne $loadedDay) {
                $UsageData[$DateKey] = $loadedDay
            }
            else {
                $UsageData[$DateKey] = New-DayStats -DateKey $DateKey
            }
        }
        else {
            $UsageData[$DateKey] = New-DayStats -DateKey $DateKey
        }
    }

    $cacheKey = Get-DayStatsCacheKey -UsageData $UsageData -DateKey $DateKey
    if (-not [string]::IsNullOrWhiteSpace($cacheKey) -and $script:DayStatsNormalizationCache.ContainsKey($cacheKey)) {
        return $UsageData[$DateKey]
    }

    $day = ConvertTo-PlainData $UsageData[$DateKey]

    if (-not $day.ContainsKey("totals")) {
        $day["totals"] = (New-DayStats -DateKey $DateKey).totals
    }

    if (-not $day.ContainsKey("activities")) {
        $day["activities"] = @{}
    }
    if (-not $day.ContainsKey("sessions") -or $null -eq $day["sessions"]) {
        $day["sessions"] = @()
    }
    elseif ($day["sessions"] -is [System.Collections.IDictionary] -or $day["sessions"] -is [pscustomobject] -or $day["sessions"] -is [System.Management.Automation.PSObject]) {
        $sessionObject = ConvertTo-PlainData $day["sessions"]
        if ($sessionObject -is [System.Collections.IDictionary] -and $sessionObject.Count -eq 0) {
            $day["sessions"] = @()
        }
        else {
            $day["sessions"] = @($sessionObject)
        }
    }
    elseif ($day["sessions"] -isnot [System.Array]) {
        $day["sessions"] = @($day["sessions"])
    }

    $normalizedSessions = @()
    foreach ($session in @($day["sessions"])) {
        $plainSession = ConvertTo-PlainData $session
        if ($plainSession -isnot [System.Collections.IDictionary] -or $plainSession.Count -eq 0) {
            continue
        }

        $endAt = ConvertTo-SafeDateTime -Value ([string](Get-ActivityField -Activity $plainSession -Name "end"))
        if ($null -eq $endAt) {
            continue
        }

        $startAt = ConvertTo-SafeDateTime -Value ([string](Get-ActivityField -Activity $plainSession -Name "start"))
        $startText = ""
        if ($null -ne $startAt) {
            $startText = $startAt.ToString("o")
        }
        $normalizedSessions += ,@{
            start = $startText
            end = $endAt.ToString("o")
            process = Get-ActivityField -Activity $plainSession -Name "process"
            title = Get-ActivityField -Activity $plainSession -Name "title"
            url = Get-ActivityField -Activity $plainSession -Name "url"
            domain = Get-ActivityField -Activity $plainSession -Name "domain"
            category = Get-ActivityField -Activity $plainSession -Name "category"
            seconds = [double](Get-ActivityField -Activity $plainSession -Name "seconds")
        }
    }
    $day["sessions"] = $normalizedSessions

    foreach ($category in @("total") + (Get-ParentCategoryKeys)) {
        if (-not $day["totals"].ContainsKey($category)) {
            $day["totals"][$category] = 0.0
        }
    }

    $UsageData[$DateKey] = $day
    $cacheKey = Get-DayStatsCacheKey -UsageData $UsageData -DateKey $DateKey
    if (-not [string]::IsNullOrWhiteSpace($cacheKey)) {
        $script:DayStatsNormalizationCache[$cacheKey] = $true
    }
    return $UsageData[$DateKey]
}

function Get-IdleSeconds {
    $info = New-Object ScreenTime.NativeMethods+LASTINPUTINFO
    $info.cbSize = [System.Runtime.InteropServices.Marshal]::SizeOf($info)

    [void][ScreenTime.NativeMethods]::GetLastInputInfo([ref]$info)
    $elapsed = [ScreenTime.NativeMethods]::GetTickCount() - $info.dwTime
    return [math]::Max(0, [math]::Round($elapsed / 1000.0, 1))
}

function Get-ActiveWindowInfo {
    $handle = [ScreenTime.NativeMethods]::GetForegroundWindow()
    if ($handle -eq [IntPtr]::Zero) {
        return $null
    }

    $builder = New-Object System.Text.StringBuilder 1024
    [void][ScreenTime.NativeMethods]::GetWindowText($handle, $builder, $builder.Capacity)
    $title = $builder.ToString().Trim()

    [uint32]$processId = 0
    [void][ScreenTime.NativeMethods]::GetWindowThreadProcessId($handle, [ref]$processId)
    if ($processId -eq 0) {
        return $null
    }

    try {
        $process = Get-Process -Id $processId -ErrorAction Stop
    }
    catch {
        return $null
    }

    return @{
        ProcessName = [string]$process.ProcessName
        WindowTitle = $title
        ProcessId = [int]$processId
        WindowHandle = $handle
    }
}

function Enter-ActionCooldown {
    param(
        [hashtable]$State,
        [string]$Key,
        [double]$CooldownSeconds
    )

    if ($null -eq $State -or [string]::IsNullOrWhiteSpace($Key)) {
        return $false
    }

    if (-not $State.ContainsKey("ActionCooldowns") -or $null -eq $State.ActionCooldowns) {
        $State["ActionCooldowns"] = @{}
    }

    $now = Get-Date
    if ($State.ActionCooldowns.ContainsKey($Key)) {
        $lastAt = $State.ActionCooldowns[$Key]
        if ($null -ne $lastAt -and ((Get-Date) - [datetime]$lastAt).TotalSeconds -lt $CooldownSeconds) {
            return $false
        }
    }

    $State.ActionCooldowns[$Key] = $now
    return $true
}

function Request-ActivityWindowClose {
    param(
        $Activity
    )

    if ($null -eq $Activity) {
        return $false
    }

    if (Should-IgnoreActivity $Activity) {
        return $false
    }

    $windowHandle = Get-ActivityField -Activity $Activity -Name "WindowHandle"
    if ($null -ne $windowHandle -and $windowHandle -ne [IntPtr]::Zero) {
        try {
            if ([ScreenTime.NativeMethods]::PostMessage([IntPtr]$windowHandle, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero)) {
                return $true
            }
        }
        catch {
        }
    }

    $processId = 0
    try {
        $processId = [int](Get-ActivityField -Activity $Activity -Name "ProcessId")
    }
    catch {
        $processId = 0
    }

    if ($processId -le 0) {
        return $false
    }

    try {
        $process = Get-Process -Id $processId -ErrorAction Stop
        if ($process.CloseMainWindow()) {
            return $true
        }
    }
    catch {
    }

    return $false
}

function Test-IsBrowserProcess {
    param(
        [string]$ProcessName
    )

    return $script:BrowserProcesses -contains (Normalize-Text $ProcessName)
}

function Start-BrowserBridgeJob {
    param(
        [int]$Port,
        [string]$OutputPath
    )

    return Start-Job -Name "ScreenTimeBrowserBridge-$Port" -ScriptBlock {
        param($JobPort, $JobOutputPath)

        $listener = New-Object System.Net.HttpListener
        $listener.Prefixes.Add("http://127.0.0.1:$JobPort/")
        $listener.Start()

        try {
            while ($listener.IsListening) {
                $context = $listener.GetContext()
                try {
                    $response = $context.Response
                    $response.ContentType = "application/json"
                    $response.Headers["Access-Control-Allow-Origin"] = "*"
                    $response.Headers["Access-Control-Allow-Headers"] = "Content-Type"
                    $response.Headers["Access-Control-Allow-Methods"] = "POST, GET, OPTIONS"

                    if ($context.Request.HttpMethod -eq "OPTIONS") {
                        $response.StatusCode = 204
                        $response.Close()
                        continue
                    }

                    if ($context.Request.HttpMethod -eq "GET") {
                        $payload = '{"status":"ok"}'
                        $bytes = [System.Text.Encoding]::UTF8.GetBytes($payload)
                        $response.OutputStream.Write($bytes, 0, $bytes.Length)
                        $response.Close()
                        continue
                    }

                    $reader = New-Object System.IO.StreamReader($context.Request.InputStream, $context.Request.ContentEncoding)
                    $body = $reader.ReadToEnd()
                    $reader.Dispose()

                    $data = ConvertFrom-Json -InputObject $body
                    $record = [ordered]@{
                        receivedAt = (Get-Date).ToString("o")
                        browser = [string]$data.browser
                        title = [string]$data.title
                        url = [string]$data.url
                        domain = [string]$data.domain
                    }

                    $json = ConvertTo-Json -InputObject $record -Depth 5
                    Set-Content -LiteralPath $JobOutputPath -Value $json -Encoding UTF8

                    $okBytes = [System.Text.Encoding]::UTF8.GetBytes('{"ok":true}')
                    $response.StatusCode = 200
                    $response.OutputStream.Write($okBytes, 0, $okBytes.Length)
                    $response.Close()
                }
                catch {
                    try {
                        $context.Response.StatusCode = 500
                        $context.Response.Close()
                    }
                    catch {
                    }
                }
            }
        }
        finally {
            $listener.Stop()
            $listener.Close()
        }
    } -ArgumentList @($Port, $OutputPath)
}

function Stop-BrowserBridgeJob {
    param(
        $Job
    )

    if ($null -eq $Job) {
        return
    }

    try {
        Stop-Job -Job $Job -ErrorAction SilentlyContinue | Out-Null
        Receive-Job -Job $Job -ErrorAction SilentlyContinue | Out-Null
        Remove-Job -Job $Job -Force -ErrorAction SilentlyContinue | Out-Null
    }
    catch {
    }
}

function Ensure-BackgroundServicesStarted {
    param(
        [hashtable]$State
    )

    if ($null -eq $State) {
        return
    }

    if ($null -eq $State.BrowserBridgeJob -or [string]$State.BrowserBridgeJob.State -in @("Completed", "Failed", "Stopped")) {
        $State.BrowserBridgeJob = Start-BrowserBridgeJob -Port ([int]$State.Settings.browserBridge.port) -OutputPath $script:BrowserActivityPath
    }
}

function Get-BrowserActivitySnapshot {
    $fallback = @{}
    $data = Read-JsonFile -Path $script:BrowserActivityPath -Fallback $fallback
    if ($data -isnot [System.Collections.IDictionary]) {
        return $fallback
    }

    return $data
}

function Get-ActivityTimestamp {
    param(
        [hashtable]$Activity
    )

    if ($null -eq $Activity -or -not $Activity.ContainsKey("receivedAt")) {
        return $null
    }

    return ConvertTo-SafeDateTime -Value ([string]$Activity.receivedAt)
}

function ConvertTo-SafeDateTime {
    param(
        [AllowNull()]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    try {
        return [datetime]::Parse($Value, [System.Globalization.CultureInfo]::InvariantCulture)
    }
    catch {
        try {
            return [datetime]::Parse($Value)
        }
        catch {
            return $null
        }
    }
}

function Get-EffectiveActivity {
    param(
        $Activity,
        [hashtable]$State
    )

    if ($null -eq $Activity) {
        return $null
    }

    $result = ConvertTo-PlainData $Activity
    $result["Url"] = ""
    $result["Domain"] = ""
    $result["BrowserSource"] = ""

    if (-not (Test-IsBrowserProcess -ProcessName $Activity.ProcessName)) {
        return $result
    }

    $browserData = Get-BrowserActivitySnapshot
    $browserTimestamp = Get-ActivityTimestamp -Activity $browserData
    if ($null -eq $browserTimestamp) {
        return $result
    }

    $maxAgeHours = [double]$State.Settings.browserBridge.titleMatchToleranceHours
    if (((Get-Date) - $browserTimestamp).TotalHours -gt $maxAgeHours) {
        return $result
    }

    $windowTitle = Normalize-Text $Activity.WindowTitle
    $tabTitle = Normalize-Text ([string]$browserData.title)
    $titleMatches = $false
    if ([string]::IsNullOrWhiteSpace($windowTitle) -or [string]::IsNullOrWhiteSpace($tabTitle)) {
        $titleMatches = $true
    }
    elseif ($windowTitle.Contains($tabTitle) -or $tabTitle.Contains($windowTitle)) {
        $titleMatches = $true
    }

    if (-not $titleMatches) {
        return $result
    }

    $result["Url"] = [string]$browserData.url
    $result["Domain"] = [string]$browserData.domain
    $result["BrowserSource"] = [string]$browserData.browser
    if (-not [string]::IsNullOrWhiteSpace([string]$browserData.title)) {
        $result["WindowTitle"] = [string]$browserData.title
    }

    return $result
}

function Get-AutoStartCommand {
    $quotedLauncher = '"' + $script:VbsLauncherPath + '"'
    return '"' + (Get-WscriptPath) + '" ' + $quotedLauncher + ' -StartMinimized'
}

function Get-WscriptPath {
    $candidate = Join-Path $env:WINDIR "System32\wscript.exe"
    if (Test-Path -LiteralPath $candidate) {
        return $candidate
    }

    return "wscript.exe"
}

function Get-StartupShortcutPath {
    $startupFolder = [Environment]::GetFolderPath([System.Environment+SpecialFolder]::Startup)
    return Join-Path $startupFolder "Screen Time Tracker.lnk"
}

function Get-DesiredAutoStartEnabled {
    param(
        [hashtable]$Settings
    )

    if ($null -eq $Settings) {
        return $true
    }

    if (-not $Settings.ContainsKey("autostart")) {
        $Settings["autostart"] = @{ enabled = $true }
    }
    if (-not $Settings.autostart.ContainsKey("enabled")) {
        $Settings.autostart["enabled"] = $true
    }

    return [bool]$Settings.autostart.enabled
}

function Test-StartupShortcutEnabled {
    $shortcutPath = Get-StartupShortcutPath
    if (-not (Test-Path -LiteralPath $shortcutPath)) {
        return $false
    }

    try {
        $shell = New-Object -ComObject WScript.Shell
        $shortcut = $shell.CreateShortcut($shortcutPath)
        return ([string]$shortcut.TargetPath).ToLowerInvariant().EndsWith("wscript.exe") -and ([string]$shortcut.Arguments).Contains("start-tracker.vbs")
    }
    catch {
        return $true
    }
}

function Set-StartupShortcutEnabled {
    param(
        [bool]$Enabled
    )

    $shortcutPath = Get-StartupShortcutPath
    if ($Enabled) {
        $shell = New-Object -ComObject WScript.Shell
        $shortcut = $shell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = Get-WscriptPath
        $shortcut.Arguments = '"' + $script:VbsLauncherPath + '" -StartMinimized'
        $shortcut.WorkingDirectory = $script:AppRoot
        if (Test-Path -LiteralPath $script:AppIconPath) {
            $shortcut.IconLocation = $script:AppIconPath
        }
        $shortcut.Save()
        return
    }

    Remove-Item -LiteralPath $shortcutPath -ErrorAction SilentlyContinue
}

function Test-AutoStartEnabled {
    $regEnabled = $false
    try {
        $value = (Get-ItemProperty -Path $script:AutoStartRegPath -Name $script:AutoStartValueName -ErrorAction Stop).$($script:AutoStartValueName)
        $regEnabled = -not [string]::IsNullOrWhiteSpace([string]$value)
    }
    catch {
    }

    return $regEnabled -or (Test-StartupShortcutEnabled)
}

function Set-AutoStartEnabled {
    param(
        [bool]$Enabled
    )

    if ($Enabled) {
        New-ItemProperty -Path $script:AutoStartRegPath -Name $script:AutoStartValueName -Value (Get-AutoStartCommand) -PropertyType String -Force | Out-Null
        Set-StartupShortcutEnabled -Enabled $true
        return
    }

    Remove-ItemProperty -Path $script:AutoStartRegPath -Name $script:AutoStartValueName -ErrorAction SilentlyContinue
    Set-StartupShortcutEnabled -Enabled $false
}

function Sync-AutoStartWithSettings {
    param(
        [hashtable]$Settings
    )

    $desiredEnabled = Get-DesiredAutoStartEnabled -Settings $Settings
    $currentEnabled = Test-AutoStartEnabled
    if ($desiredEnabled -eq $currentEnabled) {
        return
    }

    try {
        Set-AutoStartEnabled -Enabled $desiredEnabled
    }
    catch {
        try {
            Ensure-Storage
            Add-Content -LiteralPath $script:StartupLogPath -Value ("Autostart sync failed: " + $_.Exception.Message) -Encoding UTF8
        }
        catch {
        }
    }
}

function Show-TrackerWindow {
    param(
        $Form
    )

    if ($null -eq $Form) {
        return
    }

    $Form.ShowInTaskbar = $true
    $Form.Show()
    $Form.WindowState = [System.Windows.Forms.FormWindowState]::Normal
    $Form.BringToFront()
    $Form.Activate()
}

function Show-QuickGlanceWindow {
    param(
        [hashtable]$State
    )

    if ($null -eq $State -or -not $State.ContainsKey("Controls")) {
        return
    }

    $form = $State.Controls.QuickGlanceForm
    if ($null -eq $form) {
        return
    }

    $workingArea = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $desiredX = [math]::Max($workingArea.Left, $workingArea.Right - $form.Width - 24)
    $desiredY = [math]::Max($workingArea.Top, $workingArea.Top + 24)
    $form.Location = New-Object System.Drawing.Point($desiredX, $desiredY)
    $form.Show()
    $form.WindowState = [System.Windows.Forms.FormWindowState]::Normal
    $form.BringToFront()
    $form.Activate()
}

function Hide-QuickGlanceWindow {
    param(
        [hashtable]$State
    )

    if ($null -eq $State -or -not $State.ContainsKey("Controls")) {
        return
    }

    $form = $State.Controls.QuickGlanceForm
    if ($null -eq $form) {
        return
    }

    $form.Hide()
}

function Toggle-QuickGlanceWindow {
    param(
        [hashtable]$State
    )

    if ($null -eq $State -or -not $State.ContainsKey("Controls")) {
        return
    }

    $form = $State.Controls.QuickGlanceForm
    if ($null -eq $form) {
        return
    }

    if ([bool]$form.Visible) {
        Hide-QuickGlanceWindow -State $State
        return
    }

    Show-QuickGlanceWindow -State $State
}

function Request-ReopenRunningInstance {
    Ensure-Storage

    try {
        Set-Content -LiteralPath $script:ReopenSignalPath -Value (Get-Date).ToString("o") -Encoding UTF8
        return $true
    }
    catch {
        return $false
    }
}

function Consume-ReopenRunningInstanceRequest {
    if (-not (Test-Path -LiteralPath $script:ReopenSignalPath)) {
        return $false
    }

    try {
        Remove-Item -LiteralPath $script:ReopenSignalPath -Force -ErrorAction SilentlyContinue
    }
    catch {
    }

    return $true
}

function Hide-TrackerWindowToTray {
    param(
        $Form,
        [hashtable]$State,
        [string]$Message
    )

    if ($null -eq $Form) {
        return
    }

    $Form.WindowState = [System.Windows.Forms.FormWindowState]::Minimized
    $Form.ShowInTaskbar = $false
    $Form.Hide()

    if ($null -ne $State) {
        $dailyNotice = Ensure-NotificationState -State $State -DateKey $State.CurrentDateKey
        if (-not $dailyNotice.trayHint) {
            $dailyNotice.trayHint = $true
            Show-TrackerNotification -Title "Still tracking" -Message $Message
        }
    }
}

function Select-MainTab {
    param(
        [hashtable]$State,
        [string]$TabKey
    )

    if ($null -eq $State -or $null -eq $State.Controls.MainTabs) {
        return $false
    }

    $desiredKey = Normalize-Text $TabKey
    foreach ($tab in $State.Controls.MainTabs.TabPages) {
        if ((Normalize-Text ([string]$tab.Text)) -eq $desiredKey) {
            $State.Controls.MainTabs.SelectedTab = $tab
            return $true
        }
    }

    return $false
}

function Open-TrackerSection {
    param(
        [hashtable]$State,
        $Form,
        [string]$TabKey
    )

    if ($null -eq $State -or $null -eq $Form) {
        return
    }

    [void](Select-MainTab -State $State -TabKey $TabKey)
    Show-TrackerWindow -Form $Form
    Request-UiRefresh -State $State
    Update-Ui -State $State
}

function Normalize-Text {
    param(
        [AllowNull()]
        [string]$Value
    )

    if ($null -eq $Value) {
        return ""
    }

    return $Value.ToLowerInvariant()
}

function Get-ActivityField {
    param(
        $Activity,
        [string]$Name
    )

    if ($null -eq $Activity) {
        return ""
    }

    if ($Activity -is [System.Collections.IDictionary] -and $Activity.Contains($Name)) {
        return [string]$Activity[$Name]
    }

    $property = $Activity.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return ""
    }

    return [string]$property.Value
}

function Get-RuleField {
    param(
        $Rule,
        [string]$Name,
        $Default = $null
    )

    if ($null -eq $Rule) {
        return $Default
    }

    if ($Rule -is [System.Collections.IDictionary] -and $Rule.Contains($Name)) {
        return $Rule[$Name]
    }

    $property = $Rule.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $Default
    }

    return $property.Value
}

function Test-PatternMatch {
    param(
        [string]$Text,
        [string]$Pattern,
        [string]$MatchMode = "contains"
    )

    $normalizedText = Normalize-Text $Text
    $normalizedPattern = Normalize-Text $Pattern
    if ($normalizedPattern.EndsWith(".exe")) {
        $normalizedPattern = $normalizedPattern.Substring(0, $normalizedPattern.Length - 4)
    }

    if ([string]::IsNullOrWhiteSpace($normalizedPattern)) {
        return $false
    }

    switch (Normalize-Text $MatchMode) {
        "exact" { return $normalizedText -eq $normalizedPattern }
        "starts_with" { return $normalizedText.StartsWith($normalizedPattern) }
        "ends_with" { return $normalizedText.EndsWith($normalizedPattern) }
        "regex" {
            try {
                return [System.Text.RegularExpressions.Regex]::IsMatch($Text, $Pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            }
            catch {
                return $false
            }
        }
        default { return $normalizedText.Contains($normalizedPattern) }
    }
}

function Get-RulesCacheNamespace {
    param(
        [System.Array]$Rules
    )

    if ($null -eq $Rules) {
        return "default"
    }

    return [string][System.Runtime.CompilerServices.RuntimeHelpers]::GetHashCode($Rules)
}

function Get-ClassificationCacheKey {
    param(
        $Activity,
        [string]$RulesNamespace = "default"
    )

    return "{0}|{1}|{2}|{3}|{4}" -f
        $RulesNamespace,
        (Normalize-Text ([string]$Activity.ProcessName)),
        (Normalize-Text ([string]$Activity.WindowTitle)),
        (Normalize-Text (Get-ActivityField -Activity $Activity -Name "Url")),
        (Normalize-Text (Get-ActivityField -Activity $Activity -Name "Domain"))
}

function Get-RuleTextsForTarget {
    param(
        $Activity,
        [string]$Target
    )

    switch (Normalize-Text $Target) {
        "process" { return @([string]$Activity.ProcessName) }
        "window" { return @([string]$Activity.WindowTitle) }
        "title" { return @([string]$Activity.WindowTitle) }
        "url" { return @((Get-ActivityField -Activity $Activity -Name "Url"), (Get-ActivityField -Activity $Activity -Name "Domain")) }
        "domain" { return @((Get-ActivityField -Activity $Activity -Name "Domain")) }
        default {
            return @(
                [string]$Activity.ProcessName,
                [string]$Activity.WindowTitle,
                (Get-ActivityField -Activity $Activity -Name "Url"),
                (Get-ActivityField -Activity $Activity -Name "Domain")
            )
        }
    }
}

function Resolve-ActivityClassification {
    param(
        $Activity,
        [System.Array]$Rules
    )

    if ($null -eq $Activity) {
        return @{
            Category = "other"
            ParentCategory = "other"
            Matched = $false
            RuleName = ""
            RuleTarget = ""
            MatchMode = ""
            Pattern = ""
        }
    }

    $cacheKey = Get-ClassificationCacheKey -Activity $Activity -RulesNamespace (Get-RulesCacheNamespace -Rules $Rules)
    if ($script:ClassificationCache.ContainsKey($cacheKey)) {
        return $script:ClassificationCache[$cacheKey]
    }

    foreach ($rule in @($Rules)) {
        if (-not [bool](Get-RuleField -Rule $rule -Name "Enabled" -Default $true)) {
            continue
        }

        foreach ($pattern in @(Get-RuleField -Rule $rule -Name "Patterns" -Default @())) {
            foreach ($text in (Get-RuleTextsForTarget -Activity $Activity -Target ([string](Get-RuleField -Rule $rule -Name "Target" -Default "either")))) {
                if (Test-PatternMatch -Text $text -Pattern $pattern -MatchMode ([string](Get-RuleField -Rule $rule -Name "MatchMode" -Default "contains"))) {
                    $result = @{
                        Category = [string](Get-RuleField -Rule $rule -Name "Category" -Default "other")
                        ParentCategory = (Get-CategoryParentKey -Category ([string](Get-RuleField -Rule $rule -Name "Category" -Default "other")))
                        Matched = $true
                        RuleName = [string](Get-RuleField -Rule $rule -Name "Name" -Default "")
                        RuleTarget = [string](Get-RuleField -Rule $rule -Name "Target" -Default "either")
                        MatchMode = [string](Get-RuleField -Rule $rule -Name "MatchMode" -Default "contains")
                        Pattern = [string]$pattern
                    }
                    $script:ClassificationCache[$cacheKey] = $result
                    return $result
                }
            }
        }
    }

    $fallback = @{
        Category = "other"
        ParentCategory = "other"
        Matched = $false
        RuleName = ""
        RuleTarget = ""
        MatchMode = ""
        Pattern = ""
    }
    $script:ClassificationCache[$cacheKey] = $fallback
    return $fallback
}

function Should-IgnoreActivity {
    param(
        $Activity
    )

    if ($null -eq $Activity) {
        return $true
    }

    return (Normalize-Text $Activity.WindowTitle).Contains((Normalize-Text $script:WindowTitle))
}

function Get-CategoryForActivity {
    param(
        $Activity,
        [System.Array]$Rules
    )

    return [string](Resolve-ActivityClassification -Activity $Activity -Rules $Rules).Category
}

function Get-DateKey {
    return (Get-Date).ToString("yyyy-MM-dd")
}

function Get-ActivityKey {
    param(
        $Activity,
        [string]$Category
    )

    $title = [string]$Activity.WindowTitle
    if ($title.Length -gt 120) {
        $title = $title.Substring(0, 120)
    }

    $urlMarker = Get-ActivityField -Activity $Activity -Name "Domain"
    if ([string]::IsNullOrWhiteSpace($urlMarker)) {
        $urlMarker = Get-ActivityField -Activity $Activity -Name "Url"
    }
    if ($urlMarker.Length -gt 120) {
        $urlMarker = $urlMarker.Substring(0, 120)
    }

    return "{0}|{1}|{2}|{3}" -f $Category, $Activity.ProcessName, $title, $urlMarker
}

function Add-UsageSample {
    param(
        [hashtable]$UsageData,
        [string]$DateKey,
        $Activity,
        [string]$Category,
        [double]$Seconds,
        [hashtable]$State = $null
    )

    if ($Seconds -le 0) {
        return
    }

    $day = Get-DayStats -UsageData $UsageData -DateKey $DateKey
    $day.totals.total += $Seconds
    Add-SecondsToCategoryTotals -Totals $day.totals -Category $Category -Seconds $Seconds

    $normalizedCategory = Normalize-CategoryKey $Category
    if (-not (Test-CategoryExists -Category $normalizedCategory)) {
        $normalizedCategory = "other"
    }

    $key = Get-ActivityKey -Activity $Activity -Category $normalizedCategory
    if (-not $day.activities.ContainsKey($key)) {
        $day.activities[$key] = @{
            process = [string]$Activity.ProcessName
            title = [string]$Activity.WindowTitle
            url = (Get-ActivityField -Activity $Activity -Name "Url")
            domain = (Get-ActivityField -Activity $Activity -Name "Domain")
            category = $normalizedCategory
            seconds = 0.0
        }
    }

    $day.activities[$key].seconds += $Seconds
    Update-UsageSummaryForDay -DateKey $DateKey -DayStats $day
    Mark-UsageDateDirty -State $State -DateKey $DateKey
}

function Start-TrackedSession {
    param(
        [hashtable]$State,
        [string]$DateKey,
        $Activity,
        [string]$Category,
        [double]$Seconds
    )

    $day = Get-DayStats -UsageData $State.UsageData -DateKey $DateKey
    $session = @{
        start = (Get-Date).AddSeconds(-1 * [math]::Max(0.0, $Seconds)).ToString("o")
        end = (Get-Date).ToString("o")
        process = [string]$Activity.ProcessName
        title = [string]$Activity.WindowTitle
        url = (Get-ActivityField -Activity $Activity -Name "Url")
        domain = (Get-ActivityField -Activity $Activity -Name "Domain")
        category = $Category
        seconds = [double]$Seconds
    }
    $day.sessions += ,$session
    $State.CurrentSession = @{
        DateKey = $DateKey
        Key = Get-ActivityKey -Activity $Activity -Category $Category
        Session = $session
    }
}

function Update-TrackedSession {
    param(
        [hashtable]$State,
        [string]$DateKey,
        $Activity,
        [string]$Category,
        [double]$Seconds
    )

    if ($Seconds -le 0) {
        return
    }

    $sessionKey = Get-ActivityKey -Activity $Activity -Category $Category
    if ($null -eq $State.CurrentSession -or [string]$State.CurrentSession.DateKey -ne $DateKey -or [string]$State.CurrentSession.Key -ne $sessionKey) {
        Start-TrackedSession -State $State -DateKey $DateKey -Activity $Activity -Category $Category -Seconds $Seconds
        return
    }

    $State.CurrentSession.Session.end = (Get-Date).ToString("o")
    $State.CurrentSession.Session.seconds = [double]$State.CurrentSession.Session.seconds + $Seconds
}

function Stop-TrackedSession {
    param(
        [hashtable]$State
    )

    $State.CurrentSession = $null
}

function Format-Duration {
    param(
        [double]$Seconds
    )

    $timespan = [TimeSpan]::FromSeconds([math]::Max(0, [math]::Round($Seconds)))
    return "{0:00}:{1:00}:{2:00}" -f [math]::Floor($timespan.TotalHours), $timespan.Minutes, $timespan.Seconds
}

function Format-ShortDuration {
    param(
        [double]$Seconds
    )

    $timespan = [TimeSpan]::FromSeconds([math]::Max(0, [math]::Round($Seconds)))
    if ($timespan.TotalHours -ge 1) {
        return "{0}:{1:00}" -f [math]::Floor($timespan.TotalHours), $timespan.Minutes
    }

    return "{0}:{1:00}" -f $timespan.Minutes, $timespan.Seconds
}

function Get-CategoryLabel {
    param(
        [string]$Category
    )

    $key = Normalize-CategoryKey $Category
    if ($script:CategoryDefinitions.ContainsKey($key)) {
        return [string]$script:CategoryDefinitions[$key].label
    }

    return Convert-CategoryKeyToLabel -Key $key
}

function Format-ClassificationSummaryText {
    param(
        $Resolution
    )

    if ($null -eq $Resolution) {
        return "-"
    }

    $categoryText = Get-CategoryLabel ([string]$Resolution.Category)
    $parentCategory = [string]$Resolution.ParentCategory
    if (-not [string]::IsNullOrWhiteSpace($parentCategory) -and $parentCategory -ne [string]$Resolution.Category) {
        $categoryText = "{0} -> {1}" -f $categoryText, (Get-CategoryLabel $parentCategory)
    }
    if (-not [bool]$Resolution.Matched) {
        return "{0}`nNo rule matched yet" -f $categoryText
    }

    $ruleName = [string]$Resolution.RuleName
    if ([string]::IsNullOrWhiteSpace($ruleName)) {
        $ruleName = "Matched rule"
    }

    $pattern = [string]$Resolution.Pattern
    if ($pattern.Length -gt 28) {
        $pattern = $pattern.Substring(0, 28) + "..."
    }

    return "{0}`n{1} [{2} {3}: {4}]" -f $categoryText, $ruleName, ([string]$Resolution.RuleTarget), ([string]$Resolution.MatchMode), $pattern
}

function Get-TopActivities {
    param(
        [hashtable]$UsageData,
        [string]$DateKey,
        [int]$Top = 12
    )

    $day = Get-DayStats -UsageData $UsageData -DateKey $DateKey
    $activities = @()
    foreach ($value in $day.activities.Values) {
        $activities += ,[pscustomobject]@{
            process = [string]$value.process
            title = [string]$value.title
            url = [string](Get-ActivityField -Activity $value -Name "url")
            domain = [string](Get-ActivityField -Activity $value -Name "domain")
            category = [string]$value.category
            seconds = [double]$value.seconds
        }
    }

    return $activities |
        Sort-Object -Property seconds -Descending |
        Select-Object -First $Top
}

function Get-TopProcesses {
    param(
        [hashtable]$UsageData,
        [string]$DateKey,
        [int]$Top = 12
    )

    $day = Get-DayStats -UsageData $UsageData -DateKey $DateKey
    $groups = @{}
    foreach ($activity in $day.activities.Values) {
        $process = [string]$activity.process
        if ([string]::IsNullOrWhiteSpace($process)) {
            $process = "Unknown"
        }

        if (-not $groups.ContainsKey($process)) {
                $groups[$process] = [pscustomobject]@{
                    process = $process
                    category = [string]$activity.category
                    seconds = 0.0
                }
        }

        $groups[$process].seconds += [double]$activity.seconds
    }

    return $groups.Values |
        Sort-Object -Property seconds -Descending |
        Select-Object -First $Top
}

function Get-TopExactCategoriesFromActivityValues {
    param(
        $ActivityValues,
        [int]$Top = 12
    )

    $groups = @{}
    foreach ($activity in $ActivityValues) {
        $category = Normalize-CategoryKey ([string]$activity.category)
        if ([string]::IsNullOrWhiteSpace($category) -or -not (Test-CategoryExists -Category $category)) {
            $category = "other"
        }

        if (-not $groups.ContainsKey($category)) {
            $groups[$category] = [pscustomobject]@{
                category = $category
                parent = Get-CategoryParentKey -Category $category
                seconds = 0.0
            }
        }

        $groups[$category].seconds += [double]$activity.seconds
    }

    return $groups.Values |
        Sort-Object -Property seconds -Descending |
        Select-Object -First $Top
}

function Get-TopExactCategories {
    param(
        [hashtable]$UsageData,
        [string]$DateKey,
        [int]$Top = 12
    )

    $day = Get-DayStats -UsageData $UsageData -DateKey $DateKey
    return @(Get-TopExactCategoriesFromActivityValues -ActivityValues $day.activities.Values -Top $Top)
}

function Get-TopProcessSummaryText {
    param(
        [hashtable]$UsageData,
        [string]$DateKey
    )

    $top = @(Get-TopProcesses -UsageData $UsageData -DateKey $DateKey -Top 1)
    if ($top.Count -eq 0) {
        return "Most used app today: no data yet"
    }

    return "Most used app today: {0} ({1})" -f [string]$top[0].process, (Format-Duration $top[0].seconds)
}

function Set-ListViewSort {
    param(
        $ListView,
        [int]$Column,
        [System.Windows.Forms.SortOrder]$Order
    )

    $ListView.ListViewItemSorter = New-Object ScreenTime.ListViewItemComparer($Column, $Order)
    $ListView.Sort()
}

function Invoke-ListViewBatchUpdate {
    param(
        $ListView,
        [scriptblock]$Action
    )

    if ($null -eq $ListView -or $null -eq $Action) {
        return
    }

    $ListView.BeginUpdate()
    try {
        & $Action
    }
    finally {
        $ListView.EndUpdate()
    }
}

function Get-ToggledSortOrder {
    param(
        $CurrentOrder
    )

    if ($CurrentOrder -eq [System.Windows.Forms.SortOrder]::Ascending) {
        return [System.Windows.Forms.SortOrder]::Descending
    }

    return [System.Windows.Forms.SortOrder]::Ascending
}

function Get-ActivityDisplayTitle {
    param(
        $Activity
    )

    $domain = Get-ActivityField -Activity $Activity -Name "domain"
    if ([string]::IsNullOrWhiteSpace($domain)) {
        $domain = Get-ActivityField -Activity $Activity -Name "Domain"
    }

    $title = Get-ActivityField -Activity $Activity -Name "title"
    if ([string]::IsNullOrWhiteSpace($title)) {
        $title = Get-ActivityField -Activity $Activity -Name "WindowTitle"
    }

    if (-not [string]::IsNullOrWhiteSpace($domain)) {
        return "$domain | $title"
    }

    return $title
}

function Get-RecentDateKeys {
    param(
        [int]$Days = 7
    )

    $keys = @()
    for ($offset = $Days - 1; $offset -ge 0; $offset--) {
        $keys += ,((Get-Date).Date.AddDays(-$offset).ToString("yyyy-MM-dd"))
    }

    return $keys
}

function Get-RecentDaySummaries {
    param(
        [hashtable]$UsageData,
        [int]$Days = 7
    )

    $items = @()
    foreach ($dateKey in (Get-RecentDateKeys -Days $Days)) {
        $day = Get-DaySummary -UsageData $UsageData -DateKey $dateKey
        $items += ,[pscustomobject]@{
            date = $dateKey
            total = [double]$day.total
            study = [double]$day.study
            browser_fun = [double]$day.browser_fun
            socials = [double]$day.socials
        }
    }

    return $items
}

function Get-WeeklyInsightSummary {
    param(
        [hashtable]$UsageData,
        [int]$Days = 7
    )

    $summaries = Get-RecentDaySummaries -UsageData $UsageData -Days $Days
    $total = 0.0
    $study = 0.0
    $browserFun = 0.0
    $socials = 0.0
    foreach ($summary in $summaries) {
        $total += [double]$summary.total
        $study += [double]$summary.study
        $browserFun += [double]$summary.browser_fun
        $socials += [double]$summary.socials
    }
    $avg = 0
    if ($Days -gt 0) {
        $avg = $total / $Days
    }
    $bestDay = $summaries | Sort-Object -Property study -Descending | Select-Object -First 1

    $topWeeklyApps = @(Get-TopProcessesForRange -UsageData $UsageData -Days $Days -Top 10)
    $topApp = $topWeeklyApps | Select-Object -First 1
    $studyDays = @($summaries | Where-Object { $_.study -ge 7200 }).Count

    return @{
        total = [double]$total
        study = [double]$study
        browser_fun = [double]$browserFun
        socials = [double]$socials
        average = [double]$avg
        topApp = $topApp
        bestDay = $bestDay
        studyDays = [int]$studyDays
        weeklyApps = @($topWeeklyApps)
        weeklyCategories = @(Get-TopExactCategoriesForRange -UsageData $UsageData -Days $Days -Top 10)
    }
}

function Get-WeeklyReviewSummary {
    param(
        [hashtable]$UsageData,
        [hashtable]$Settings,
        [int]$Days = 7
    )

    $limits = $Settings.limits
    $summaries = @(Get-RecentDaySummaries -UsageData $UsageData -Days $Days)
    $slippedDays = @()
    $underLimitDays = 0
    $studyGoalDays = 0
    $studyWindowDays = 0
    $socialOverTotal = 0.0
    $funOverTotal = 0.0
    $studyGapTotal = 0.0

    foreach ($summary in $summaries) {
        $totalOver = [math]::Max(0.0, [double]$summary.total - [double]$limits.total)
        $studyGap = [math]::Max(0.0, [double]$limits.studyMin - [double]$summary.study)
        $funOver = [math]::Max(0.0, [double]$summary.browser_fun - [double]$limits.browser_fun)
        $socialOver = [math]::Max(0.0, [double]$summary.socials - [double]$limits.socials)

        if ([double]$summary.total -le [double]$limits.total) {
            $underLimitDays += 1
        }
        if ([double]$summary.study -ge [double]$limits.studyMin) {
            $studyGoalDays += 1
        }
        if ([double]$summary.study -ge [double]$limits.studyMin -and [double]$summary.study -le [double]$limits.studyMax) {
            $studyWindowDays += 1
        }

        $socialOverTotal += $socialOver
        $funOverTotal += $funOver
        $studyGapTotal += $studyGap

        $issueText = "Clean day"
        $score = 0.0
        if ($totalOver -gt 0) {
            $issueText = "Total over by $(Format-ShortDuration $totalOver)"
            $score = $totalOver + $socialOver + $funOver
        }
        elseif ($socialOver -gt 0) {
            $issueText = "Socials over by $(Format-ShortDuration $socialOver)"
            $score = $socialOver
        }
        elseif ($funOver -gt 0) {
            $issueText = "Fun over by $(Format-ShortDuration $funOver)"
            $score = $funOver
        }
        elseif ($studyGap -gt 0) {
            $issueText = "Study short by $(Format-ShortDuration $studyGap)"
            $score = $studyGap
        }

        if ($score -gt 0) {
            $slippedDays += ,[pscustomobject]@{
                date = [string]$summary.date
                issue = $issueText
                total = [double]$summary.total
                study = [double]$summary.study
                score = [double]$score
            }
        }
    }

    $topDistractingApps = @(Get-TopProcessesForRange -UsageData $UsageData -Days $Days -Top 15 | Where-Object {
        (Get-CategoryParentKey -Category ([string]$_.category)) -in @("socials", "browser_fun")
    } | Select-Object -First 10)

    $biggestSlip = @($slippedDays | Sort-Object -Property score -Descending | Select-Object -First 1)
    $bestWin = @($summaries | Where-Object {
        [double]$_.total -le [double]$limits.total -and [double]$_.study -ge [double]$limits.studyMin
    } | Sort-Object -Property study -Descending | Select-Object -First 1)

    $summaryText = "Week review: $underLimitDays / $Days days stayed under the 3h limit | $studyGoalDays / $Days days reached at least 2h study"
    $winText = "Best win: keep building clean study days."
    if ($bestWin.Count -gt 0) {
        $winText = "Best win: $([string]$bestWin[0].date) | $(Format-Duration $bestWin[0].study) study and still within limit"
    }

    $biggestSlipText = "Biggest slip: no major problems this week."
    if ($biggestSlip.Count -gt 0) {
        $biggestSlipText = "Biggest slip: $([string]$biggestSlip[0].date) | $([string]$biggestSlip[0].issue)"
    }

    $coachNote = "Coach note: this week looks steady. Keep the same rhythm."
    if ($socialOverTotal -gt $funOverTotal -and $socialOverTotal -gt 0) {
        $coachNote = "Coach note: socials were the main leak this week. Protect them first."
    }
    elseif ($funOverTotal -gt 0) {
        $coachNote = "Coach note: browser fun / manga took the bigger bite this week. Tighten that block first."
    }
    elseif ($studyGapTotal -gt 0) {
        $coachNote = "Coach note: the main gain is more study consistency. Add $(Format-ShortDuration $studyGapTotal) across the week."
    }

    $mainDistractionText = "Main distraction: none stood out."
    if ($topDistractingApps.Count -gt 0) {
        $mainDistractionText = "Main distraction: $([string]$topDistractingApps[0].process) | $(Format-Duration $topDistractingApps[0].seconds)"
    }

    return @{
        summaryText = $summaryText
        winText = $winText
        biggestSlipText = $biggestSlipText
        coachNote = $coachNote
        mainDistractionText = $mainDistractionText
        underLimitDays = [int]$underLimitDays
        studyGoalDays = [int]$studyGoalDays
        studyWindowDays = [int]$studyWindowDays
        slippedDays = @($slippedDays | Sort-Object -Property score -Descending)
        topDistractingApps = $topDistractingApps
    }
}

function Get-WeekStartDate {
    param(
        [datetime]$Date
    )

    $dayOffset = (([int]$Date.DayOfWeek + 6) % 7)
    return $Date.Date.AddDays(-$dayOffset)
}

function Get-HeatmapDayStatus {
    param(
        $Summary,
        [hashtable]$Settings,
        [datetime]$Date
    )

    $today = (Get-Date).Date
    if ($Date.Date -gt $today) {
        return [pscustomobject]@{
            key = "future"
            label = "Future"
            backColor = [System.Drawing.Color]::WhiteSmoke
            foreColor = [System.Drawing.Color]::DarkGray
        }
    }

    $total = [double]$Summary.total
    $study = [double]$Summary.study
    $fun = [double]$Summary.browser_fun
    $socials = [double]$Summary.socials
    $limits = $Settings.limits

    if ($total -le 0) {
        return [pscustomobject]@{
            key = "empty"
            label = "No tracked time"
            backColor = [System.Drawing.Color]::Gainsboro
            foreColor = [System.Drawing.Color]::Black
        }
    }

    if ($total -gt [double]$limits.total) {
        return [pscustomobject]@{
            key = "over_total"
            label = "Total over limit"
            backColor = [System.Drawing.Color]::IndianRed
            foreColor = [System.Drawing.Color]::White
        }
    }

    if ($socials -gt [double]$limits.socials -or $fun -gt [double]$limits.browser_fun) {
        return [pscustomobject]@{
            key = "over_fun_social"
            label = "Fun or socials over limit"
            backColor = [System.Drawing.Color]::DarkOrange
            foreColor = [System.Drawing.Color]::White
        }
    }

    if ($study -ge [double]$limits.studyMin -and $study -le [double]$limits.studyMax) {
        return [pscustomobject]@{
            key = "clean"
            label = "Within limit and study goal met"
            backColor = [System.Drawing.Color]::ForestGreen
            foreColor = [System.Drawing.Color]::White
        }
    }

    return [pscustomobject]@{
        key = "within_limit"
        label = "Within limit but study goal missed"
        backColor = [System.Drawing.Color]::Khaki
        foreColor = [System.Drawing.Color]::Black
    }
}

function Get-CalendarHeatmapSummary {
    param(
        [hashtable]$UsageData,
        [hashtable]$Settings,
        [int]$Weeks = 6
    )

    $today = (Get-Date).Date
    $startDate = (Get-WeekStartDate -Date $today).AddDays(-7 * ($Weeks - 1))
    $cells = @()
    $counts = @{
        clean = 0
        within_limit = 0
        over_fun_social = 0
        over_total = 0
        empty = 0
    }

    for ($index = 0; $index -lt ($Weeks * 7); $index += 1) {
        $date = $startDate.AddDays($index)
        $dateKey = $date.ToString("yyyy-MM-dd")
        $summary = Get-DaySummary -UsageData $UsageData -DateKey $dateKey
        $status = Get-HeatmapDayStatus -Summary $summary -Settings $Settings -Date $date
        if ($counts.ContainsKey([string]$status.key)) {
            $counts[[string]$status.key] += 1
        }

        $tooltip = "{0}`n{1}`nTotal {2} | Study {3} | Fun {4} | Socials {5}" -f $dateKey, [string]$status.label, (Format-ShortDuration ([double]$summary.total)), (Format-ShortDuration ([double]$summary.study)), (Format-ShortDuration ([double]$summary.browser_fun)), (Format-ShortDuration ([double]$summary.socials))
        $cells += ,[pscustomobject]@{
            date = $date
            dateKey = $dateKey
            day = $date.Day
            statusKey = [string]$status.key
            statusLabel = [string]$status.label
            backColor = $status.backColor
            foreColor = $status.foreColor
            tooltip = $tooltip
        }
    }

    $legend = "Legend: green = clean day | yellow = within limit but study short | orange = fun/social spill | red = total over | gray = no usage"
    $summaryText = "Heatmap: clean $($counts.clean) | short-study $($counts.within_limit) | fun/social spills $($counts.over_fun_social) | total overs $($counts.over_total)"
    return @{
        cells = $cells
        legend = $legend
        summaryText = $summaryText
    }
}

function Get-CategoryColor {
    param(
        [string]$Category
    )

    switch (Get-CategoryParentKey -Category $Category) {
        "study" { return [System.Drawing.Color]::FromArgb(76, 175, 80) }
        "browser_fun" { return [System.Drawing.Color]::FromArgb(255, 167, 38) }
        "socials" { return [System.Drawing.Color]::FromArgb(66, 165, 245) }
        default { return [System.Drawing.Color]::FromArgb(158, 158, 158) }
    }
}

function Get-CategoryKeysInDisplayOrder {
    return @(Get-ParentCategoryKeys)
}

function Get-TopProcessesForRange {
    param(
        [hashtable]$UsageData,
        [int]$Days = 30,
        [int]$Top = 15
    )

    $groups = @{}
    foreach ($dateKey in (Get-RecentDateKeys -Days $Days)) {
        $day = Get-DaySummary -UsageData $UsageData -DateKey $dateKey
        foreach ($processEntry in $day.processes.Values) {
            $process = [string]$processEntry.process
            if ([string]::IsNullOrWhiteSpace($process)) {
                $process = "Unknown"
            }

            if (-not $groups.ContainsKey($process)) {
                $groups[$process] = [pscustomobject]@{
                    process = $process
                    category = [string]$processEntry.category
                    seconds = 0.0
                }
            }

            $groups[$process].seconds += [double]$processEntry.seconds
        }
    }

    return $groups.Values |
        Sort-Object -Property seconds -Descending |
        Select-Object -First $Top
}

function Get-TopExactCategoriesForRange {
    param(
        [hashtable]$UsageData,
        [int]$Days = 30,
        [int]$Top = 12
    )

    $groups = @{}
    foreach ($dateKey in (Get-RecentDateKeys -Days $Days)) {
        $day = Get-DaySummary -UsageData $UsageData -DateKey $dateKey
        foreach ($categoryEntry in $day.exactCategories.Values) {
            $category = Normalize-CategoryKey ([string]$categoryEntry.category)
            if ([string]::IsNullOrWhiteSpace($category) -or -not (Test-CategoryExists -Category $category)) {
                $category = "other"
            }

            if (-not $groups.ContainsKey($category)) {
                $groups[$category] = [pscustomobject]@{
                    category = $category
                    parent = Get-CategoryParentKey -Category $category
                    seconds = 0.0
                }
            }

            $groups[$category].seconds += [double]$categoryEntry.seconds
        }
    }

    return @($groups.Values |
        Sort-Object -Property seconds -Descending |
        Select-Object -First $Top)
}

function Get-RecentSessions {
    param(
        [hashtable]$UsageData,
        [int]$Days = 1,
        [int]$Top = 20
    )

    $items = @()
    foreach ($dateKey in (Get-RecentDateKeys -Days $Days)) {
        $day = Get-DayStats -UsageData $UsageData -DateKey $dateKey
        foreach ($session in $day.sessions) {
            $items += ,[pscustomobject]@{
                date = $dateKey
                start = Get-ActivityField -Activity $session -Name "start"
                end = Get-ActivityField -Activity $session -Name "end"
                process = Get-ActivityField -Activity $session -Name "process"
                title = Get-ActivityField -Activity $session -Name "title"
                url = Get-ActivityField -Activity $session -Name "url"
                domain = Get-ActivityField -Activity $session -Name "domain"
                category = Get-ActivityField -Activity $session -Name "category"
                seconds = [double](Get-ActivityField -Activity $session -Name "seconds")
            }
        }
    }

    return $items |
        Sort-Object -Property @{ Expression = {
            $endAt = ConvertTo-SafeDateTime -Value ([string](Get-ActivityField -Activity $_ -Name "end"))
            if ($null -ne $endAt) { return $endAt }
            return [datetime]::MinValue
        }; Descending = $true } |
        Select-Object -First $Top
}

function Get-SessionStreakSummary {
    param(
        [hashtable]$UsageData,
        [hashtable]$Settings,
        [int]$Days = 30
    )

    $summaries = @(Get-RecentDaySummaries -UsageData $UsageData -Days $Days)
    $currentUnderLimit = 0
    $bestUnderLimit = 0
    $currentStudyGoal = 0
    $bestStudyGoal = 0
    $rollingUnder = 0
    $rollingStudy = 0

    foreach ($day in $summaries) {
        $hasUsage = [double]$day.total -gt 0
        $underLimit = $hasUsage -and [double]$day.total -le [double]$Settings.limits.total
        $studyGoal = $hasUsage -and [double]$day.study -ge [double]$Settings.limits.studyMin -and [double]$day.study -le [double]$Settings.limits.studyMax

        if ($underLimit) {
            $rollingUnder += 1
            if ($rollingUnder -gt $bestUnderLimit) {
                $bestUnderLimit = $rollingUnder
            }
        }
        else {
            $rollingUnder = 0
        }

        if ($studyGoal) {
            $rollingStudy += 1
            if ($rollingStudy -gt $bestStudyGoal) {
                $bestStudyGoal = $rollingStudy
            }
        }
        else {
            $rollingStudy = 0
        }
    }

    $recentDays = @($summaries | Sort-Object -Property date -Descending)
    foreach ($day in $recentDays) {
        $hasUsage = [double]$day.total -gt 0
        if ($hasUsage -and [double]$day.total -le [double]$Settings.limits.total) {
            $currentUnderLimit += 1
        }
        else {
            break
        }
    }

    foreach ($day in $recentDays) {
        $hasUsage = [double]$day.total -gt 0
        if ($hasUsage -and [double]$day.study -ge [double]$Settings.limits.studyMin -and [double]$day.study -le [double]$Settings.limits.studyMax) {
            $currentStudyGoal += 1
        }
        else {
            break
        }
    }

    return @{
        currentUnderLimit = [int]$currentUnderLimit
        bestUnderLimit = [int]$bestUnderLimit
        currentStudyGoal = [int]$currentStudyGoal
        bestStudyGoal = [int]$bestStudyGoal
    }
}

function Get-GoalsDashboardSummary {
    param(
        [hashtable]$UsageData,
        [string]$DateKey,
        [hashtable]$Settings
    )

    $day = Get-DayStats -UsageData $UsageData -DateKey $DateKey
    $totals = $day.totals
    $limits = $Settings.limits

    $met = 0
    if ([double]$totals.total -le [double]$limits.total) { $met += 1 }
    if ([double]$totals.study -ge [double]$limits.studyMin -and [double]$totals.study -le [double]$limits.studyMax) { $met += 1 }
    if ([double]$totals.browser_fun -le [double]$limits.browser_fun) { $met += 1 }
    if ([double]$totals.socials -le [double]$limits.socials) { $met += 1 }

    $summary = "Goals met today: $met / 4"
    if ($met -eq 4) {
        $summary = "Goals met today: 4 / 4"
    }

    $primary = ""
    if ([double]$totals.study -lt [double]$limits.studyMin) {
        $primary = "Next milestone: $(Format-ShortDuration ([double]$limits.studyMin - [double]$totals.study)) left to reach study minimum."
    }
    elseif ([double]$totals.total -gt [double]$limits.total) {
        $primary = "Main risk: total computer time is over by $(Format-ShortDuration ([double]$totals.total - [double]$limits.total))."
    }
    elseif ([double]$totals.socials -gt [double]$limits.socials) {
        $primary = "Main risk: social media is over by $(Format-ShortDuration ([double]$totals.socials - [double]$limits.socials))."
    }
    elseif ([double]$totals.browser_fun -gt [double]$limits.browser_fun) {
        $primary = "Main risk: browser fun / manga is over by $(Format-ShortDuration ([double]$totals.browser_fun - [double]$limits.browser_fun))."
    }
    elseif ([double]$totals.study -gt [double]$limits.studyMax) {
        $primary = "Study target passed the upper bound by $(Format-ShortDuration ([double]$totals.study - [double]$limits.studyMax))."
    }
    else {
        $primary = "You are inside the daily limits right now."
    }

    $unlockText = "Fun unlocked: not yet. Finish at least $(Format-ShortDuration ([double]$limits.studyMin - [double]$totals.study)) of study first."
    if ([double]$totals.study -ge [double]$limits.studyMin) {
        $unlockText = "Fun unlocked: yes. Study minimum is already done."
    }

    return @{
        goalsMet = [int]$met
        summary = $summary
        primary = $primary
        unlock = $unlockText
    }
}

function Get-TodayTimelineBuckets {
    param(
        [hashtable]$UsageData,
        [string]$DateKey
    )

    $buckets = @()
    for ($hour = 0; $hour -lt 24; $hour++) {
        $buckets += ,[pscustomobject]@{
            hour = $hour
            label = "{0:00}:00" -f $hour
            total = 0.0
            study = 0.0
            browser_fun = 0.0
            socials = 0.0
            other = 0.0
        }
    }

    $day = Get-DayStats -UsageData $UsageData -DateKey $DateKey
    foreach ($session in @($day.sessions)) {
        $endAt = ConvertTo-SafeDateTime -Value ([string](Get-ActivityField -Activity $session -Name "end"))
        if ($null -eq $endAt) {
            continue
        }

        $seconds = [double](Get-ActivityField -Activity $session -Name "seconds")
        if ($seconds -le 0) {
            continue
        }

        $startAt = ConvertTo-SafeDateTime -Value ([string](Get-ActivityField -Activity $session -Name "start"))
        if ($null -eq $startAt) {
            $startAt = $endAt.AddSeconds(-1 * $seconds)
        }
        if ($startAt -gt $endAt) {
            $startAt = $endAt.AddSeconds(-1 * $seconds)
        }

        $category = [string](Get-ActivityField -Activity $session -Name "category")
        if ([string]::IsNullOrWhiteSpace($category) -or -not (Test-CategoryExists -Category $category)) {
            $category = "other"
        }

        $cursor = $startAt
        while ($cursor -lt $endAt) {
            $hourStart = $cursor.Date.AddHours($cursor.Hour)
            $hourEnd = $hourStart.AddHours(1)
            $segmentEnd = $hourEnd
            if ($endAt -lt $hourEnd) {
                $segmentEnd = $endAt
            }
            $segmentSeconds = [math]::Max(0.0, ($segmentEnd - $cursor).TotalSeconds)
            if ($segmentSeconds -gt 0 -and $cursor.Hour -ge 0 -and $cursor.Hour -lt 24) {
                $buckets[$cursor.Hour].$category += $segmentSeconds
                $buckets[$cursor.Hour].total += $segmentSeconds
            }
            $cursor = $segmentEnd
        }
    }

    return $buckets
}

function Update-HourlyTimelineChart {
    param(
        $Chart,
        [System.Array]$Buckets
    )

    $Chart.Series.Clear()
    $Chart.Legends.Clear()
    $Chart.Titles.Clear()

    $hasData = $false
    foreach ($bucket in $Buckets) {
        if ([double]$bucket.total -gt 0) {
            $hasData = $true
            break
        }
    }
    if (-not $hasData) {
        [void]$Chart.Titles.Add("No sessions tracked yet today")
        return
    }

    $legend = New-Object System.Windows.Forms.DataVisualization.Charting.Legend
    $legend.Docking = [System.Windows.Forms.DataVisualization.Charting.Docking]::Top
    [void]$Chart.Legends.Add($legend)

    foreach ($category in (Get-CategoryKeysInDisplayOrder)) {
        $series = New-Object System.Windows.Forms.DataVisualization.Charting.Series((Get-CategoryLabel $category))
        $series.ChartArea = "Timeline"
        $series.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::StackedColumn
        $series.BorderWidth = 1
        $series.Color = Get-CategoryColor -Category $category
        [void]$Chart.Series.Add($series)
    }

    foreach ($bucket in $Buckets) {
        foreach ($category in (Get-CategoryKeysInDisplayOrder)) {
            $hours = [double]$bucket.$category / 3600.0
            [void]$Chart.Series[(Get-CategoryLabel $category)].Points.AddXY($bucket.label, [math]::Round($hours, 2))
        }
    }
}

function Get-AnalyticsSummary {
    param(
        [hashtable]$UsageData,
        [int]$Days = 30
    )

    $summaries = Get-RecentDaySummaries -UsageData $UsageData -Days $Days
    $totals = @{
        total = 0.0
        study = 0.0
        browser_fun = 0.0
        socials = 0.0
        other = 0.0
    }

    foreach ($summary in $summaries) {
        $totals.total += [double]$summary.total
        $totals.study += [double]$summary.study
        $totals.browser_fun += [double]$summary.browser_fun
        $totals.socials += [double]$summary.socials
        $totals.other += [double]$summary.total - [double]$summary.study - [double]$summary.browser_fun - [double]$summary.socials
    }

    $topApps = @(Get-TopProcessesForRange -UsageData $UsageData -Days $Days -Top 15)
    $topDistractingApp = @($topApps | Where-Object { $_.category -in @("browser_fun", "socials") } | Select-Object -First 1)
    if ($topDistractingApp.Count -eq 0) {
        $topDistractingApp = @($topApps | Where-Object { $_.category -ne "study" } | Select-Object -First 1)
    }

    $averageSeconds = 0.0
    if ($Days -gt 0) {
        $averageSeconds = [double]($totals.total / $Days)
    }

    return @{
        days = [int]$Days
        totals = $totals
        average = $averageSeconds
        bestStudyDay = ($summaries | Sort-Object -Property study -Descending | Select-Object -First 1)
        topApp = ($topApps | Select-Object -First 1)
        topDistractingApp = ($topDistractingApp | Select-Object -First 1)
        daily = @($summaries)
        topApps = $topApps
        topExactCategories = @(Get-TopExactCategoriesForRange -UsageData $UsageData -Days $Days -Top 12)
    }
}

function Get-AnalyticsExportRows {
    param(
        [hashtable]$UsageData,
        [int]$Days = 30
    )

    $summary = Get-AnalyticsSummary -UsageData $UsageData -Days $Days
    $rows = @()
    foreach ($day in $summary.daily) {
        $other = [double]$day.total - [double]$day.study - [double]$day.browser_fun - [double]$day.socials
        $rows += ,[pscustomobject]@{
            RowType = "day"
            Scope = "$Days-day"
            Date = [string]$day.date
            Process = ""
            Category = ""
            Seconds = ""
            Hours = ""
            SharePercent = ""
            TotalHours = [math]::Round(([double]$day.total / 3600.0), 2)
            StudyHours = [math]::Round(([double]$day.study / 3600.0), 2)
            BrowserFunHours = [math]::Round(([double]$day.browser_fun / 3600.0), 2)
            SocialsHours = [math]::Round(([double]$day.socials / 3600.0), 2)
            OtherHours = [math]::Round(($other / 3600.0), 2)
        }
    }

    $periodTotal = [double]$summary.totals.total
    foreach ($app in $summary.topApps) {
        $sharePercent = 0.0
        if ($periodTotal -gt 0) {
            $sharePercent = [math]::Round((([double]$app.seconds / $periodTotal) * 100), 1)
        }

        $rows += ,[pscustomobject]@{
            RowType = "app"
            Scope = "$Days-day"
            Date = ""
            Process = [string]$app.process
            Category = [string]$app.category
            Seconds = [math]::Round([double]$app.seconds)
            Hours = [math]::Round(([double]$app.seconds / 3600.0), 2)
            SharePercent = $sharePercent
            TotalHours = ""
            StudyHours = ""
            BrowserFunHours = ""
            SocialsHours = ""
            OtherHours = ""
        }
    }

    foreach ($category in $summary.topExactCategories) {
        $sharePercent = 0.0
        if ($periodTotal -gt 0) {
            $sharePercent = [math]::Round((([double]$category.seconds / $periodTotal) * 100), 1)
        }

        $rows += ,[pscustomobject]@{
            RowType = "exact_category"
            Scope = "$Days-day"
            Date = ""
            Process = ""
            Category = [string]$category.category
            Seconds = [math]::Round([double]$category.seconds)
            Hours = [math]::Round(([double]$category.seconds / 3600.0), 2)
            SharePercent = $sharePercent
            TotalHours = ""
            StudyHours = ""
            BrowserFunHours = ""
            SocialsHours = ""
            OtherHours = ""
        }
    }

    return $rows
}

function Export-AnalyticsCsv {
    param(
        [hashtable]$UsageData,
        [int]$Days = 30,
        [string]$Path
    )

    $rows = Get-AnalyticsExportRows -UsageData $UsageData -Days $Days
    $rows | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
}

function Export-UsageDataJson {
    param(
        [hashtable]$UsageData,
        [string]$Path
    )

    $snapshot = Get-AllStoredUsageData -UsageData $UsageData
    $json = ConvertTo-Json -InputObject (ConvertTo-PlainData $snapshot) -Depth 12
    Set-Content -LiteralPath $Path -Value $json -Encoding UTF8
}

function Show-SaveFileDialog {
    param(
        [string]$Title,
        [string]$Filter,
        [string]$DefaultFileName
    )

    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Title = $Title
    $dialog.Filter = $Filter
    $dialog.FileName = $DefaultFileName
    $dialog.InitialDirectory = $script:ExportsDirectory
    $dialog.OverwritePrompt = $true
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.FileName
    }

    return $null
}

function Update-StackedTrendChart {
    param(
        $Chart,
        [System.Array]$Summaries
    )

    $Chart.Series.Clear()
    $Chart.Legends.Clear()
    $Chart.Titles.Clear()

    $legend = New-Object System.Windows.Forms.DataVisualization.Charting.Legend
    $legend.Docking = [System.Windows.Forms.DataVisualization.Charting.Docking]::Top
    [void]$Chart.Legends.Add($legend)

    $hasData = $false
    foreach ($day in $Summaries) {
        if ([double]$day.total -gt 0) {
            $hasData = $true
            break
        }
    }
    if (-not $hasData) {
        [void]$Chart.Titles.Add("No tracked time yet")
        return
    }

    foreach ($category in (Get-CategoryKeysInDisplayOrder)) {
        $series = New-Object System.Windows.Forms.DataVisualization.Charting.Series((Get-CategoryLabel $category))
        $series.ChartArea = "Trend"
        $series.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::StackedColumn
        $series.BorderWidth = 1
        $series.Color = Get-CategoryColor -Category $category
        $series.IsValueShownAsLabel = $false
        [void]$Chart.Series.Add($series)
    }

    foreach ($day in $Summaries) {
        $label = ([datetime]::ParseExact([string]$day.date, "yyyy-MM-dd", [System.Globalization.CultureInfo]::InvariantCulture)).ToString("MM-dd")
        $other = [math]::Max(0.0, [double]$day.total - [double]$day.study - [double]$day.browser_fun - [double]$day.socials)
        $values = @{
            study = [double]$day.study / 3600.0
            browser_fun = [double]$day.browser_fun / 3600.0
            socials = [double]$day.socials / 3600.0
            other = $other / 3600.0
        }

        foreach ($category in (Get-CategoryKeysInDisplayOrder)) {
            [void]$Chart.Series[(Get-CategoryLabel $category)].Points.AddXY($label, [math]::Round([double]$values[$category], 2))
        }
    }
}

function Update-DoughnutChart {
    param(
        $Chart,
        [hashtable]$Totals
    )

    $Chart.Series.Clear()
    $Chart.Legends.Clear()
    $Chart.Titles.Clear()

    $legend = New-Object System.Windows.Forms.DataVisualization.Charting.Legend
    $legend.Docking = [System.Windows.Forms.DataVisualization.Charting.Docking]::Right
    [void]$Chart.Legends.Add($legend)

    $hasData = $false
    foreach ($category in (Get-CategoryKeysInDisplayOrder)) {
        if ([double]$Totals[$category] -gt 0) {
            $hasData = $true
            break
        }
    }
    if (-not $hasData) {
        [void]$Chart.Titles.Add("No tracked time yet")
        return
    }

    $series = New-Object System.Windows.Forms.DataVisualization.Charting.Series("Breakdown")
    $series.ChartArea = "Breakdown"
    $series.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Doughnut
    $series.IsValueShownAsLabel = $true
    $series.LabelForeColor = [System.Drawing.Color]::Black
    $series.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    [void]$Chart.Series.Add($series)

    foreach ($category in (Get-CategoryKeysInDisplayOrder)) {
        $seconds = [double]$Totals[$category]
        if ($seconds -le 0) {
            continue
        }

        $pointIndex = $series.Points.AddY([math]::Round(($seconds / 3600.0), 2))
        $point = $series.Points[$pointIndex]
        $point.AxisLabel = Get-CategoryLabel $category
        $point.LegendText = Get-CategoryLabel $category
        $point.Label = "{0}: {1:N1}h" -f (Get-CategoryLabel $category), ($seconds / 3600.0)
        $point.Color = Get-CategoryColor -Category $category
    }
}

function Move-ExistingActivityToCategory {
    param(
        [hashtable]$State,
        [string]$DateKey,
        $Activity,
        [string]$NewCategory
    )

    $day = Get-DayStats -UsageData $State.UsageData -DateKey $DateKey
    $keys = @($day.activities.Keys)
    foreach ($key in $keys) {
        $entry = $day.activities[$key]
        $sameProcess = ([string]$entry.process -eq [string]$Activity.ProcessName)
        $sameTitle = ([string]$entry.title -eq [string]$Activity.WindowTitle)
        $sameDomain = ((Get-ActivityField -Activity $entry -Name "domain") -eq (Get-ActivityField -Activity $Activity -Name "Domain"))

        if (-not $sameProcess) {
            continue
        }

        if (-not $sameTitle -and -not $sameDomain) {
            continue
        }

        $oldCategory = [string]$entry.category
        if ($oldCategory -eq $NewCategory) {
            continue
        }

        $seconds = [double]$entry.seconds
        Remove-SecondsFromCategoryTotals -Totals $day.totals -Category $oldCategory -Seconds $seconds
        Add-SecondsToCategoryTotals -Totals $day.totals -Category $NewCategory -Seconds $seconds

        $day.activities.Remove($key)
        $normalizedNewCategory = Normalize-CategoryKey $NewCategory
        if (-not (Test-CategoryExists -Category $normalizedNewCategory)) {
            $normalizedNewCategory = "other"
        }
        $entry.category = $normalizedNewCategory
        $newKey = "{0}|{1}|{2}|{3}" -f $normalizedNewCategory, $entry.process, $entry.title, (Get-ActivityField -Activity $entry -Name "domain")
        $day.activities[$newKey] = $entry
    }

    Update-UsageSummaryForDay -DateKey $DateKey -DayStats $day
    Clear-DerivedCaches -State $State
    Mark-UsageDateDirty -State $State -DateKey $DateKey
    $State.IsDirty = $true
}

function Move-ExistingActivitiesByRule {
    param(
        [hashtable]$State,
        [string]$DateKey,
        [string]$NewCategory,
        [string]$Target,
        [string]$Pattern,
        [string]$MatchMode = "contains"
    )

    if ([string]::IsNullOrWhiteSpace([string]$Pattern)) {
        return
    }

    $day = Get-DayStats -UsageData $State.UsageData -DateKey $DateKey
    $keys = @($day.activities.Keys)
    foreach ($key in $keys) {
        $entry = $day.activities[$key]
        $activity = @{
            ProcessName = [string]$entry.process
            WindowTitle = [string]$entry.title
            Url = Get-ActivityField -Activity $entry -Name "url"
            Domain = Get-ActivityField -Activity $entry -Name "domain"
        }

        $matchesRule = $false
        foreach ($text in (Get-RuleTextsForTarget -Activity $activity -Target $Target)) {
            if (Test-PatternMatch -Text $text -Pattern $Pattern -MatchMode $MatchMode) {
                $matchesRule = $true
                break
            }
        }

        if (-not $matchesRule) {
            continue
        }

        $oldCategory = [string]$entry.category
        if ($oldCategory -eq $NewCategory) {
            continue
        }

        $seconds = [double]$entry.seconds
        Remove-SecondsFromCategoryTotals -Totals $day.totals -Category $oldCategory -Seconds $seconds
        Add-SecondsToCategoryTotals -Totals $day.totals -Category $NewCategory -Seconds $seconds

        $day.activities.Remove($key)
        $normalizedNewCategory = Normalize-CategoryKey $NewCategory
        if (-not (Test-CategoryExists -Category $normalizedNewCategory)) {
            $normalizedNewCategory = "other"
        }
        $entry.category = $normalizedNewCategory
        $newKey = "{0}|{1}|{2}|{3}" -f $normalizedNewCategory, $entry.process, $entry.title, (Get-ActivityField -Activity $entry -Name "domain")
        $day.activities[$newKey] = $entry
    }

    Update-UsageSummaryForDay -DateKey $DateKey -DayStats $day
    Clear-DerivedCaches -State $State
    Mark-UsageDateDirty -State $State -DateKey $DateKey
    $State.IsDirty = $true
}

function Add-ManualRuleForActivity {
    param(
        [hashtable]$State,
        $Activity,
        [string]$Category,
        [string]$Target,
        [string]$Pattern,
        [string]$MatchMode = "contains",
        [int]$Priority = 10
    )

    $trimmed = $Pattern.Trim()
    if ([string]::IsNullOrWhiteSpace($trimmed)) {
        return $false
    }

    $rule = @{
        Name = "Manual: $trimmed"
        Category = $Category
        Target = $Target
        MatchMode = $MatchMode
        Priority = $Priority
        Enabled = $true
        Patterns = @($trimmed.ToLowerInvariant())
    }

    $rules = @($rule)
    foreach ($existing in $State.Rules) {
        $rules += ,$existing
    }

    Save-Rules -Rules $rules
    $State.Rules = @(Normalize-Rules -Rules $rules)
    Clear-DerivedCaches -State $State
    return $true
}

function Get-FocusModeText {
    param(
        [hashtable]$State
    )

    if (-not $State.FocusMode.Enabled) {
        return "Focus mode: off"
    }

    if ($null -eq $State.FocusMode.Until) {
        return "Focus mode: on"
    }

    $left = [math]::Max(0, [math]::Round(($State.FocusMode.Until - (Get-Date)).TotalSeconds))
    return "Focus mode: on ({0} left)" -f (Format-ShortDuration $left)
}

function Start-FocusMode {
    param(
        [hashtable]$State,
        [int]$Minutes
    )

    $State.FocusMode.Enabled = $true
    $State.FocusMode.Until = (Get-Date).AddMinutes($Minutes)
    $State.FocusMode.LastPromptAt = Get-Date "2000-01-01"
    Show-TrackerNotification -Title "Focus mode started" -Message "Focus mode is active for $Minutes minutes."
}

function Stop-FocusMode {
    param(
        [hashtable]$State
    )

    $State.FocusMode.Enabled = $false
    $State.FocusMode.Until = $null
    Show-TrackerNotification -Title "Focus mode stopped" -Message "Distracting apps are no longer blocked by focus mode."
}

function Show-FocusModeDialog {
    param(
        [hashtable]$State
    )

    $defaultMinutes = [int]$State.Settings.focusMode.defaultMinutes
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Start focus mode for $defaultMinutes minutes?`nDuring focus mode, socials and browser-fun windows will be interrupted.",
        "Focus mode",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        Start-FocusMode -State $State -Minutes $defaultMinutes
        return $true
    }

    return $false
}

function Handle-FocusMode {
    param(
        [hashtable]$State,
        $Activity,
        [string]$Category
    )

    if (-not $State.FocusMode.Enabled) {
        return $true
    }

    if ($null -ne $State.FocusMode.Until -and (Get-Date) -ge $State.FocusMode.Until) {
        Stop-FocusMode -State $State
        return $true
    }

    $parentCategory = Get-CategoryParentKey -Category $Category
    if ($parentCategory -notin @("socials", "browser_fun")) {
        return $true
    }

    if ([bool]$State.Settings.focusMode.closeDistractingWindows) {
        [void](Invoke-DistractingActivityBlock -State $State -Activity $Activity -Category $Category -Source "focus")
        return $false
    }

    $cooldown = [double]$State.Settings.focusMode.promptCooldownSeconds
    if (((Get-Date) - $State.FocusMode.LastPromptAt).TotalSeconds -lt $cooldown) {
        return $false
    }

    $State.FocusMode.LastPromptAt = Get-Date
    [System.Windows.Forms.MessageBox]::Show(
            "Focus mode is active.`n`nSwitch away from $(Get-CategoryLabel $Category) and return to your study task.",
        "Focus mode block",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    ) | Out-Null
    return $false
}

function Ensure-NotificationState {
    param(
        [hashtable]$State,
        [string]$DateKey
    )

    if (-not $State.Notifications.ContainsKey($DateKey)) {
        $State.Notifications[$DateKey] = @{
            total = $false
            totalWarn = $false
            browser_fun = $false
            browser_funWarn = $false
            socials = $false
            socialsWarn = $false
            study = $false
            studyWarn = $false
            trayHint = $false
        }
    }

    return $State.Notifications[$DateKey]
}

function Show-TrackerNotification {
    param(
        [string]$Title,
        [string]$Message
    )

    if ($null -eq $script:NotifyIcon) {
        return
    }

    $script:NotifyIcon.BalloonTipTitle = $Title
    $script:NotifyIcon.BalloonTipText = $Message
    $script:NotifyIcon.ShowBalloonTip(5000)
}

function Invoke-DistractingActivityBlock {
    param(
        [hashtable]$State,
        $Activity,
        [string]$Category,
        [string]$Source
    )

    if ($null -eq $Activity) {
        return $false
    }

    $parentCategory = Get-CategoryParentKey -Category $Category
    if ($parentCategory -notin @("socials", "browser_fun")) {
        return $false
    }

    $cooldownSeconds = 10.0
    if ([string]$Source -eq "focus") {
        $cooldownSeconds = [double]$State.Settings.focusMode.promptCooldownSeconds
    }
    elseif ($State.Settings.notifications.ContainsKey("blockCooldownSeconds")) {
        $cooldownSeconds = [double]$State.Settings.notifications.blockCooldownSeconds
    }

    $cooldownKey = "{0}|{1}" -f $Source, $parentCategory
    if (-not (Enter-ActionCooldown -State $State -Key $cooldownKey -CooldownSeconds $cooldownSeconds)) {
        return $false
    }

    $label = Get-CategoryLabel $Category
    $closed = Request-ActivityWindowClose -Activity $Activity
    if ([string]$Source -eq "focus") {
        if ($closed) {
            Show-TrackerNotification -Title "Focus mode blocked a distraction" -Message "$label was closed so you can get back to your task."
        }
        else {
            Show-TrackerNotification -Title "Focus mode block" -Message "Switch away from $label and return to your study task."
        }
    }
    else {
        if ($closed) {
            Show-TrackerNotification -Title "Limit enforced" -Message "$label was closed because its daily limit is already exhausted."
        }
        else {
            Show-TrackerNotification -Title "Limit enforced" -Message "$label is already over limit. Switch to something else."
        }
    }

    return $closed
}

function Check-Limits {
    param(
        [hashtable]$State
    )

    $notified = Ensure-NotificationState -State $State -DateKey $State.CurrentDateKey
    $day = Get-DayStats -UsageData $State.UsageData -DateKey $State.CurrentDateKey
    $limits = $State.Settings.limits
    $warningPercent = [int]$State.Settings.notifications.warningPercent
    $warningFactor = [double]$State.Settings.notifications.warningPercent / 100.0

    if (-not $notified.totalWarn -and $day.totals.total -ge ([double]$limits.total * $warningFactor)) {
        $notified.totalWarn = $true
        Show-TrackerNotification -Title "Approaching total limit" -Message "You used $warningPercent% of your daily total computer time."
    }

    if (-not $notified.total -and $day.totals.total -ge [double]$limits.total) {
        $notified.total = $true
        Show-TrackerNotification -Title "Daily screen time limit" -Message "You reached $(Format-ShortDuration $limits.total) of total computer time."
    }

    if (-not $notified.browser_funWarn -and $day.totals.browser_fun -ge ([double]$limits.browser_fun * $warningFactor)) {
        $notified.browser_funWarn = $true
        Show-TrackerNotification -Title "Approaching browser fun limit" -Message "You used $warningPercent% of your browser fun / manga time."
    }

    if (-not $notified.browser_fun -and $day.totals.browser_fun -ge [double]$limits.browser_fun) {
        $notified.browser_fun = $true
        Show-TrackerNotification -Title "Browser fun limit" -Message "Browser games / manga time reached $(Format-ShortDuration $limits.browser_fun)."
    }

    if (-not $notified.socialsWarn -and $day.totals.socials -ge ([double]$limits.socials * $warningFactor)) {
        $notified.socialsWarn = $true
        Show-TrackerNotification -Title "Approaching social limit" -Message "You used $warningPercent% of your social media time."
    }

    if (-not $notified.socials -and $day.totals.socials -ge [double]$limits.socials) {
        $notified.socials = $true
        Show-TrackerNotification -Title "Social media limit" -Message "Social media time reached $(Format-ShortDuration $limits.socials)."
    }

    if (-not $notified.studyWarn -and $day.totals.study -ge ([double]$limits.studyMax * $warningFactor)) {
        $notified.studyWarn = $true
        Show-TrackerNotification -Title "Approaching study cap" -Message "You used $warningPercent% of your study upper bound."
    }

    if (-not $notified.study -and $day.totals.study -ge [double]$limits.studyMax) {
        $notified.study = $true
        Show-TrackerNotification -Title "Study cap reached" -Message "Study time reached $(Format-ShortDuration $limits.studyMax)."
    }

    Handle-HardLimit -State $State -Category "total" -Current ([double]$day.totals.total) -Limit ([double]$limits.total)
    Handle-HardLimit -State $State -Category "browser_fun" -Current ([double]$day.totals.browser_fun) -Limit ([double]$limits.browser_fun)
    Handle-HardLimit -State $State -Category "socials" -Current ([double]$day.totals.socials) -Limit ([double]$limits.socials)
}

function Save-IfDirty {
    param(
        [hashtable]$State,
        [switch]$Force
    )

    $now = Get-Date
    if (-not $Force -and (-not $State.IsDirty)) {
        return
    }

    if (-not $Force -and (($now - $State.LastSavedAt).TotalSeconds -lt 30)) {
        return
    }

    $dirtyDateKeys = @()
    if ($State.ContainsKey("DirtyDateKeys") -and $null -ne $State.DirtyDateKeys) {
        $dirtyDateKeys = @($State.DirtyDateKeys.Keys)
    }

    Save-UsageData -UsageData $State.UsageData -DateKeys $dirtyDateKeys
    $State.LastSavedAt = $now
    $State.IsDirty = $false
    $State.DirtyDateKeys = @{}
}

function Get-StatusText {
    param(
        [double]$Current,
        [double]$Limit
    )

    if ($Current -le $Limit) {
        return "Within limit"
    }

    return "{0} over limit" -f (Format-ShortDuration ($Current - $Limit))
}

function Get-RemainingSeconds {
    param(
        [double]$Current,
        [double]$Limit
    )

    return [math]::Max(0, $Limit - $Current)
}

function Get-RemainingSummaryText {
    param(
        [hashtable]$State
    )

    $day = Get-DayStats -UsageData $State.UsageData -DateKey $State.CurrentDateKey
    $limits = $State.Settings.limits
    $remainingTotal = Get-RemainingSeconds -Current ([double]$day.totals.total) -Limit ([double]$limits.total)
    $remainingFun = Get-RemainingSeconds -Current ([double]$day.totals.browser_fun) -Limit ([double]$limits.browser_fun)
    $remainingSocials = Get-RemainingSeconds -Current ([double]$day.totals.socials) -Limit ([double]$limits.socials)

    return "Left today {0} | Fun {1} | Socials {2}" -f (Format-ShortDuration $remainingTotal), (Format-ShortDuration $remainingFun), (Format-ShortDuration $remainingSocials)
}

function Get-CompactProcessName {
    param(
        [string]$ProcessName
    )

    $value = [string]$ProcessName
    if ([string]::IsNullOrWhiteSpace($value)) {
        return "-"
    }

    if ($value.Length -gt 18) {
        return $value.Substring(0, 18) + "..."
    }

    return $value
}

function Get-CurrentTrayActivityText {
    param(
        [hashtable]$State
    )

    if ($null -eq $State) {
        return "Current: idle / no app"
    }

    if ([bool]$State.IsIdle) {
        return "Current: idle ({0})" -f (Format-ShortDuration $State.CurrentIdleSeconds)
    }

    if ($null -eq $State.LastActivity) {
        return "Current: idle / no app"
    }

    $process = Get-CompactProcessName -ProcessName ([string]$State.LastActivity.ProcessName)
    $category = "other"
    if (-not [string]::IsNullOrWhiteSpace([string]$State.LastActivityCategory)) {
        $category = [string]$State.LastActivityCategory
    }

    return "Current: {0} ({1})" -f $process, (Get-CategoryLabel $category)
}

function Get-TrayTooltipText {
    param(
        [hashtable]$State
    )

    $summary = Get-RemainingSummaryText -State $State
    $current = Get-CurrentTrayActivityText -State $State
    $tooltip = "{0}`n{1}" -f $summary, $current

    if ($tooltip.Length -gt 63) {
        $tooltip = $tooltip.Substring(0, 60) + "..."
    }

    return $tooltip
}

function Update-QuickGlanceUi {
    param(
        [hashtable]$State,
        [hashtable]$DayStats = $null
    )

    if ($null -eq $State -or -not $State.ContainsKey("Controls")) {
        return
    }

    $quickForm = $State.Controls.QuickGlanceForm
    if ($null -eq $quickForm -or -not [bool]$quickForm.Visible) {
        return
    }

    $day = $DayStats
    if ($null -eq $day) {
        $day = Get-DayStats -UsageData $State.UsageData -DateKey $State.CurrentDateKey
    }

    $limits = $State.Settings.limits
    $remainingTotal = Get-RemainingSeconds -Current ([double]$day.totals.total) -Limit ([double]$limits.total)
    $remainingStudyMin = [math]::Max(0, [double]$limits.studyMin - [double]$day.totals.study)
    $remainingFun = Get-RemainingSeconds -Current ([double]$day.totals.browser_fun) -Limit ([double]$limits.browser_fun)
    $remainingSocials = Get-RemainingSeconds -Current ([double]$day.totals.socials) -Limit ([double]$limits.socials)

    $State.Controls.QuickGlanceDateLabel.Text = "Today: $($State.CurrentDateKey)"
    $State.Controls.QuickGlanceTotalLabel.Text = "Left total: $(Format-ShortDuration $remainingTotal)"
    $State.Controls.QuickGlanceStudyLabel.Text = "To study goal: $(Format-ShortDuration $remainingStudyMin)"
    $State.Controls.QuickGlanceFunLabel.Text = "Fun left: $(Format-ShortDuration $remainingFun)"
    $State.Controls.QuickGlanceSocialsLabel.Text = "Socials left: $(Format-ShortDuration $remainingSocials)"
    $State.Controls.QuickGlanceCurrentLabel.Text = Get-CurrentTrayActivityText -State $State
    $State.Controls.QuickGlanceFocusLabel.Text = Get-FocusModeText -State $State

    $statusText = "Tracking paused"
    if ($State.TrackingEnabled) {
        $statusText = "Tracking active"
    }
    if ($State.IsIdle) {
        $statusText = "Idle for $(Format-ShortDuration $State.CurrentIdleSeconds)"
    }
    $State.Controls.QuickGlanceStatusLabel.Text = $statusText

    $pauseButtonText = "Resume"
    if ($State.TrackingEnabled) {
        $pauseButtonText = "Pause"
    }
    $State.Controls.QuickGlancePauseButton.Text = $pauseButtonText
}

function Update-TrayStatusUi {
    param(
        [hashtable]$State
    )

    if ($null -eq $script:NotifyIcon -or -not $State.ContainsKey("TrayItems")) {
        return
    }

    $summary = Get-RemainingSummaryText -State $State
    $text = Get-TrayTooltipText -State $State

    try {
        if (-not $State.ContainsKey("LastTrayTooltipText") -or [string]$State.LastTrayTooltipText -ne $text) {
            $script:NotifyIcon.Text = $text
            $State.LastTrayTooltipText = $text
        }
    }
    catch {
        $script:NotifyIcon.Text = "Screen Time Tracker"
        $State.LastTrayTooltipText = "Screen Time Tracker"
    }

    $State.TrayItems.StatusItem.Text = $summary
    $pauseText = "Resume tracking"
    if ($State.TrackingEnabled) {
        $pauseText = "Pause tracking"
    }
    $focusText = "Start focus mode"
    if ($State.FocusMode.Enabled) {
        $focusText = "Stop focus mode"
    }
    $State.TrayItems.PauseItem.Text = $pauseText
    $State.TrayItems.FocusItem.Text = $focusText

    if ($State.TrayItems.ContainsKey("TopAppItem")) {
        $State.TrayItems.TopAppItem.Text = Get-TopProcessSummaryText -UsageData $State.UsageData -DateKey $State.CurrentDateKey
    }

    if ($State.TrayItems.ContainsKey("CurrentItem")) {
        $State.TrayItems.CurrentItem.Text = Get-CurrentTrayActivityText -State $State
    }

    if ($State.TrayItems.ContainsKey("ReviewItem")) {
        $reviewSummary = Get-UncategorizedSnapshotForState -State $State -Top 20
        $reviewText = "Review uncategorized"
        if ($reviewSummary.count -gt 0) {
            $reviewText = "Review uncategorized ({0})" -f $reviewSummary.count
        }
        $State.TrayItems.ReviewItem.Text = $reviewText
        $State.TrayItems.ReviewItem.Enabled = $reviewSummary.count -gt 0
    }

    if ($State.TrayItems.ContainsKey("ClassifyItem")) {
        $State.TrayItems.ClassifyItem.Enabled = $null -ne $State.LastActivity
    }

    if ($State.TrayItems.ContainsKey("QuickGlanceItem") -and $State.Controls.ContainsKey("QuickGlanceForm")) {
        $quickGlanceText = "Show quick glance"
        if ([bool]$State.Controls.QuickGlanceForm.Visible) {
            $quickGlanceText = "Hide quick glance"
        }
        $State.TrayItems.QuickGlanceItem.Text = $quickGlanceText
    }
}

function Show-HardLimitDialog {
    param(
        [hashtable]$State,
        [string]$Category,
        [double]$LimitSeconds
    )

    $label = switch ($Category) {
        "total" { "Total computer time" }
        "browser_fun" { "Browser fun / manga" }
        "socials" { "Social media" }
        default { "Time limit" }
    }

    $result = [System.Windows.Forms.MessageBox]::Show(
        "$label reached $(Format-ShortDuration $LimitSeconds).`n`nYes = take a break now and pause tracking.`nNo = allow one short extra block.",
        "Hard limit reached",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        $State.TrackingEnabled = $false
        return
    }

    $minutes = [double]$State.Settings.notifications.hardLimitSnoozeMinutes
    $State.OverrideUntil = (Get-Date).AddMinutes($minutes)
    $State.HardLimitHandledFor[$Category] = $false
    Show-TrackerNotification -Title "Short extension granted" -Message "Tracking will allow another $(Format-ShortDuration ($minutes * 60)) before hard limit returns."
}

function Handle-HardLimit {
    param(
        [hashtable]$State,
        [string]$Category,
        [double]$Current,
        [double]$Limit
    )

    if (-not [bool]$State.Settings.notifications.hardLimitEnabled) {
        return
    }

    if ($Current -lt $Limit) {
        return
    }

    if ($Category -in @("browser_fun", "socials") -and [bool]$State.Settings.notifications.closeDistractingWindows) {
        $currentCategory = "other"
        if (-not [string]::IsNullOrWhiteSpace([string]$State.LastActivityCategory)) {
            $currentCategory = [string]$State.LastActivityCategory
        }

        if ((Get-CategoryParentKey -Category $currentCategory) -eq $Category) {
            [void](Invoke-DistractingActivityBlock -State $State -Activity $State.LastActivity -Category $currentCategory -Source "hard-limit")
        }
        return
    }

    if ($null -ne $State.OverrideUntil -and (Get-Date) -lt $State.OverrideUntil) {
        return
    }

    if ($State.HardLimitHandledFor.ContainsKey($Category) -and $State.HardLimitHandledFor[$Category]) {
        return
    }

    $State.HardLimitHandledFor[$Category] = $true
    Show-HardLimitDialog -State $State -Category $Category -LimitSeconds $Limit
}

function Get-StudyStatusText {
    param(
        [double]$Current,
        [double]$StudyMin,
        [double]$StudyMax
    )

    if ($Current -lt $StudyMin) {
        return "{0} left to reach goal" -f (Format-ShortDuration ($StudyMin - $Current))
    }

    if ($Current -le $StudyMax) {
        return "Inside target range"
    }

    return "{0} above target cap" -f (Format-ShortDuration ($Current - $StudyMax))
}

function Set-RowState {
    param(
        $Row,
        [double]$Current,
        [double]$Limit,
        [string]$Status,
        [bool]$Alert
    )

    $Row.ValueLabel.Text = "{0} / {1}" -f (Format-Duration $Current), (Format-Duration $Limit)
    $Row.ProgressBar.Maximum = [int][math]::Max(1, [math]::Round($Limit))
    $Row.ProgressBar.Value = [int][math]::Min($Row.ProgressBar.Maximum, [math]::Round($Current))
    $Row.StatusLabel.Text = $Status
    $statusColor = [System.Drawing.Color]::DarkGreen
    if ($Alert) {
        $statusColor = [System.Drawing.Color]::Firebrick
    }
    $Row.StatusLabel.ForeColor = $statusColor
}

function Update-RuleList {
    param(
        $RuleList,
        [System.Array]$Rules
    )

    $RuleList.Items.Clear()
    foreach ($rule in $Rules) {
        $item = New-Object System.Windows.Forms.ListViewItem($rule.Name)
        [void]$item.SubItems.Add((Get-CategoryLabel $rule.Category))
        [void]$item.SubItems.Add([string]$rule.Target)
        [void]$item.SubItems.Add([string]$rule.MatchMode)
        [void]$item.SubItems.Add([string]$rule.Priority)
        $enabledText = "No"
        if ([bool]$rule.Enabled) {
            $enabledText = "Yes"
        }
        [void]$item.SubItems.Add($enabledText)
        [void]$item.SubItems.Add(([string[]]$rule.Patterns -join ", "))
        [void]$RuleList.Items.Add($item)
    }
}

function Get-UncategorizedActivities {
    param(
        [hashtable]$UsageData,
        [string]$DateKey,
        [System.Array]$Rules,
        [int]$Top = 20
    )

    return @((Get-UncategorizedAnalysis -UsageData $UsageData -DateKey $DateKey -Rules $Rules -Top $Top).items)
}

function Get-UncategorizedActivitySummary {
    param(
        [hashtable]$UsageData,
        [string]$DateKey,
        [System.Array]$Rules
    )

    $analysis = Get-UncategorizedAnalysis -UsageData $UsageData -DateKey $DateKey -Rules $Rules -Top 20
    return [pscustomobject]@{
        count = $analysis.count
        seconds = $analysis.seconds
    }
}

function Get-UncategorizedAnalysis {
    param(
        [hashtable]$UsageData,
        [string]$DateKey,
        [System.Array]$Rules,
        [int]$Top = 20
    )

    $items = @()
    $count = 0
    $totalSeconds = 0.0
    $day = Get-DayStats -UsageData $UsageData -DateKey $DateKey
    foreach ($value in $day.activities.Values) {
        $activity = @{
            ProcessName = [string]$value.process
            WindowTitle = [string]$value.title
            Url = Get-ActivityField -Activity $value -Name "url"
            Domain = Get-ActivityField -Activity $value -Name "domain"
        }
        $resolution = Resolve-ActivityClassification -Activity $activity -Rules $Rules
        if ([bool]$resolution.Matched) {
            continue
        }

        $count += 1
        $totalSeconds += [double]$value.seconds
        $items += ,[pscustomobject]@{
            process = [string]$value.process
            title = Get-ActivityDisplayTitle $value
            seconds = [double]$value.seconds
            activity = $activity
        }
    }

    $topItems = @($items | Sort-Object -Property seconds -Descending | Select-Object -First $Top)
    return [pscustomobject]@{
        count = $count
        seconds = $totalSeconds
        items = $topItems
    }
}

function Test-ProcessRuleSuggestionCandidate {
    param(
        [string]$ProcessName
    )

    $normalized = Normalize-Text $ProcessName
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return $false
    }

    $blocked = @(
        "chrome",
        "msedge",
        "firefox",
        "brave",
        "opera",
        "vivaldi",
        "iexplore",
        "browser",
        "explorer",
        "applicationframehost",
        "searchhost",
        "widgets"
    )

    return $blocked -notcontains $normalized
}

function Get-RuleSuggestionForActivity {
    param(
        $Activity
    )

    if ($null -eq $Activity) {
        return $null
    }

    $domain = [string](Get-ActivityField -Activity $Activity -Name "Domain")
    if ([string]::IsNullOrWhiteSpace($domain)) {
        $url = [string](Get-ActivityField -Activity $Activity -Name "Url")
        if (-not [string]::IsNullOrWhiteSpace($url)) {
            try {
                $uri = [System.Uri]$url
                if ($null -ne $uri -and -not [string]::IsNullOrWhiteSpace([string]$uri.Host)) {
                    $domain = [string]$uri.Host
                }
            }
            catch {
            }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($domain)) {
        $pattern = $domain.Trim().ToLowerInvariant()
        if ($pattern.StartsWith("www.")) {
            $pattern = $pattern.Substring(4)
        }

        return [pscustomobject]@{
            target = "domain"
            matchMode = "contains"
            pattern = $pattern
            reason = "Stable site match"
        }
    }

    $process = [string]$Activity.ProcessName
    if (Test-ProcessRuleSuggestionCandidate -ProcessName $process) {
        return [pscustomobject]@{
            target = "process"
            matchMode = "exact"
            pattern = $process.Trim().ToLowerInvariant()
            reason = "Stable app match"
        }
    }

    return $null
}

function Get-RuleSuggestions {
    param(
        [hashtable]$UsageData,
        [string]$DateKey,
        [System.Array]$Rules,
        [int]$Top = 10
    )

    $groups = @{}
    $coveredHits = 0
    $seconds = 0.0
    $skipped = 0
    $day = Get-DayStats -UsageData $UsageData -DateKey $DateKey
    foreach ($value in $day.activities.Values) {
        $activity = @{
            ProcessName = [string]$value.process
            WindowTitle = [string]$value.title
            Url = Get-ActivityField -Activity $value -Name "url"
            Domain = Get-ActivityField -Activity $value -Name "domain"
        }
        $resolution = Resolve-ActivityClassification -Activity $activity -Rules $Rules
        if ([bool]$resolution.Matched) {
            continue
        }

        $suggestion = Get-RuleSuggestionForActivity -Activity $activity
        if ($null -eq $suggestion) {
            $skipped += 1
            continue
        }

        $coveredHits += 1
        $entrySeconds = [double]$value.seconds
        $seconds += $entrySeconds
        $groupKey = "{0}|{1}|{2}" -f [string]$suggestion.target, [string]$suggestion.matchMode, [string]$suggestion.pattern
        if (-not $groups.ContainsKey($groupKey)) {
            $groups[$groupKey] = [pscustomobject]@{
                target = [string]$suggestion.target
                matchMode = [string]$suggestion.matchMode
                pattern = [string]$suggestion.pattern
                reason = [string]$suggestion.reason
                seconds = 0.0
                hits = 0
                process = [string]$value.process
                example = Get-ActivityDisplayTitle $value
                activity = $activity
            }
        }

        $groups[$groupKey].seconds += $entrySeconds
        $groups[$groupKey].hits += 1
    }

    $topItems = @($groups.Values | Sort-Object -Property seconds -Descending | Select-Object -First $Top)
    return [pscustomobject]@{
        count = [int]$groups.Count
        coveredHits = $coveredHits
        seconds = $seconds
        skipped = $skipped
        items = $topItems
    }
}

function Clear-DerivedCaches {
    param(
        [hashtable]$State
    )

    if ($null -eq $State) {
        return
    }

    $State.DerivedCaches = @{}
}

function Get-UncategorizedSnapshotForState {
    param(
        [hashtable]$State,
        [int]$Top = 20
    )

    if (-not $State.ContainsKey("DerivedCaches")) {
        $State["DerivedCaches"] = @{}
    }

    $cacheKey = "uncategorized"
    $rulesNamespace = Get-RulesCacheNamespace -Rules $State.Rules
    if ($State.DerivedCaches.ContainsKey($cacheKey)) {
        $cached = $State.DerivedCaches[$cacheKey]
        if ($null -ne $cached -and
            [string]$cached.dateKey -eq $State.CurrentDateKey -and
            [string]$cached.rulesNamespace -eq $rulesNamespace -and
            ((Get-Date) - [datetime]$cached.createdAt).TotalSeconds -lt 15) {
            return $cached
        }
    }

    $analysis = Get-UncategorizedAnalysis -UsageData $State.UsageData -DateKey $State.CurrentDateKey -Rules $State.Rules -Top $Top
    $snapshot = @{
        dateKey = $State.CurrentDateKey
        rulesNamespace = $rulesNamespace
        createdAt = Get-Date
        count = [int]$analysis.count
        seconds = [double]$analysis.seconds
        items = @($analysis.items)
    }
    $State.DerivedCaches[$cacheKey] = $snapshot
    return $snapshot
}

function Get-ClassificationHealthSnapshot {
    param(
        [hashtable]$UsageData,
        [string]$DateKey,
        [System.Array]$Rules,
        [int]$Days = 7
    )

    $todayDay = Get-DayStats -UsageData $UsageData -DateKey $DateKey
    $todayTotal = [double]$todayDay.totals.total
    $todayMatched = 0.0
    $todayUnmatched = 0.0
    $todayUnmatchedCount = 0
    $uncategorizedGroups = @{}

    foreach ($value in $todayDay.activities.Values) {
        $activity = @{
            ProcessName = [string]$value.process
            WindowTitle = [string]$value.title
            Url = Get-ActivityField -Activity $value -Name "url"
            Domain = Get-ActivityField -Activity $value -Name "domain"
        }

        $resolution = Resolve-ActivityClassification -Activity $activity -Rules $Rules
        $seconds = [double]$value.seconds
        if ([bool]$resolution.Matched) {
            $todayMatched += $seconds
            continue
        }

        $todayUnmatched += $seconds
        $todayUnmatchedCount += 1
        $process = [string]$value.process
        if ([string]::IsNullOrWhiteSpace($process)) {
            $process = "Unknown"
        }

        if (-not $uncategorizedGroups.ContainsKey($process)) {
            $uncategorizedGroups[$process] = [pscustomobject]@{
                process = $process
                seconds = 0.0
                examples = 0
                title = Get-ActivityDisplayTitle $value
            }
        }

        $uncategorizedGroups[$process].seconds += $seconds
        $uncategorizedGroups[$process].examples += 1
    }

    $weekTotal = 0.0
    $weekMatched = 0.0
    $weekUnmatched = 0.0
    $ruleGroups = @{}
    foreach ($recentDateKey in (Get-RecentDateKeys -Days $Days)) {
        $day = Get-DayStats -UsageData $UsageData -DateKey $recentDateKey
        foreach ($value in $day.activities.Values) {
            $activity = @{
                ProcessName = [string]$value.process
                WindowTitle = [string]$value.title
                Url = Get-ActivityField -Activity $value -Name "url"
                Domain = Get-ActivityField -Activity $value -Name "domain"
            }

            $resolution = Resolve-ActivityClassification -Activity $activity -Rules $Rules
            $seconds = [double]$value.seconds
            $weekTotal += $seconds
            if (-not [bool]$resolution.Matched) {
                $weekUnmatched += $seconds
                continue
            }

            $weekMatched += $seconds
            $ruleName = [string]$resolution.RuleName
            if ([string]::IsNullOrWhiteSpace($ruleName)) {
                $ruleName = "Matched rule"
            }

            $category = [string]$resolution.Category
            $groupKey = "{0}|{1}" -f $ruleName, $category
            if (-not $ruleGroups.ContainsKey($groupKey)) {
                $ruleGroups[$groupKey] = [pscustomobject]@{
                    rule = $ruleName
                    category = $category
                    seconds = 0.0
                    matches = 0
                }
            }

            $ruleGroups[$groupKey].seconds += $seconds
            $ruleGroups[$groupKey].matches += 1
        }
    }

    $ruleSuggestions = Get-RuleSuggestions -UsageData $UsageData -DateKey $DateKey -Rules $Rules -Top 10

    $todayCoveragePercent = 0.0
    if ($todayTotal -gt 0) {
        $todayCoveragePercent = [math]::Round((($todayMatched / $todayTotal) * 100), 1)
    }

    $weekCoveragePercent = 0.0
    if ($weekTotal -gt 0) {
        $weekCoveragePercent = [math]::Round((($weekMatched / $weekTotal) * 100), 1)
    }

    $topUncategorized = @($uncategorizedGroups.Values | Sort-Object -Property seconds -Descending | Select-Object -First 12)
    $topRules = @($ruleGroups.Values | Sort-Object -Property seconds -Descending | Select-Object -First 12)

    $mainGapText = "Nothing urgent to classify."
    if ($topUncategorized.Count -gt 0) {
        $mainGapText = "Main gap today: {0} | {1}" -f [string]$topUncategorized[0].process, (Format-Duration $topUncategorized[0].seconds)
    }

    return [pscustomobject]@{
        todayCoveragePercent = $todayCoveragePercent
        weekCoveragePercent = $weekCoveragePercent
        todayMatchedSeconds = [double]$todayMatched
        todayUnmatchedSeconds = [double]$todayUnmatched
        todayUnmatchedCount = [int]$todayUnmatchedCount
        weekMatchedSeconds = [double]$weekMatched
        weekUnmatchedSeconds = [double]$weekUnmatched
        topUncategorized = $topUncategorized
        topRules = $topRules
        suggestedRuleCount = [int]$ruleSuggestions.count
        suggestedRuleSeconds = [double]$ruleSuggestions.seconds
        suggestionSkippedCount = [int]$ruleSuggestions.skipped
        topSuggestions = @($ruleSuggestions.items)
        mainGap = $mainGapText
    }
}

function Get-ClassificationHealthSnapshotForState {
    param(
        [hashtable]$State,
        [int]$Days = 7
    )

    if (-not $State.ContainsKey("DerivedCaches")) {
        $State["DerivedCaches"] = @{}
    }

    $cacheKey = "classification-health"
    $rulesNamespace = Get-RulesCacheNamespace -Rules $State.Rules
    if ($State.DerivedCaches.ContainsKey($cacheKey)) {
        $cached = $State.DerivedCaches[$cacheKey]
        if ($null -ne $cached -and
            [string]$cached.dateKey -eq $State.CurrentDateKey -and
            [string]$cached.rulesNamespace -eq $rulesNamespace -and
            ((Get-Date) - [datetime]$cached.createdAt).TotalSeconds -lt 20) {
            return $cached.snapshot
        }
    }

    $snapshot = Get-ClassificationHealthSnapshot -UsageData $State.UsageData -DateKey $State.CurrentDateKey -Rules $State.Rules -Days $Days
    $State.DerivedCaches[$cacheKey] = @{
        dateKey = $State.CurrentDateKey
        rulesNamespace = $rulesNamespace
        createdAt = Get-Date
        snapshot = $snapshot
    }

    return $snapshot
}

function Update-ClassificationReviewUi {
    param(
        [hashtable]$State
    )

    $analysis = Get-UncategorizedSnapshotForState -State $State -Top 20
    $items = @($analysis.items)
    $summary = $analysis

    $reviewSummaryText = "Uncategorized today: nothing important left to classify."
    if ($summary.count -gt 0) {
        $reviewSummaryText = "Uncategorized today: $($summary.count) items | $(Format-Duration $summary.seconds)"
    }
    $State.Controls.ReviewSummaryLabel.Text = $reviewSummaryText

    $list = $State.Controls.ReviewList
    Invoke-ListViewBatchUpdate -ListView $list -Action {
        $list.Items.Clear()
        foreach ($itemData in $items) {
            $item = New-Object System.Windows.Forms.ListViewItem([string]$itemData.process)
            [void]$item.SubItems.Add([string]$itemData.title)
            [void]$item.SubItems.Add((Format-Duration $itemData.seconds))
            $item.Tag = $itemData.activity
            [void]$list.Items.Add($item)
        }

        Set-ListViewSort -ListView $list -Column $State.ListSorts.Review.Column -Order $State.ListSorts.Review.Order
    }
    $State.Controls.ReviewClassifyButton.Enabled = $list.Items.Count -gt 0
}

function Update-ClassificationHealthUi {
    param(
        [hashtable]$State
    )

    $snapshot = Get-ClassificationHealthSnapshotForState -State $State -Days 7

    $State.Controls.HealthTodayCoverageLabel.Text = "Today classified: {0:N1}% ({1} matched / {2} needs review)" -f $snapshot.todayCoveragePercent, (Format-Duration $snapshot.todayMatchedSeconds), (Format-Duration $snapshot.todayUnmatchedSeconds)
    $State.Controls.HealthWeekCoverageLabel.Text = "7-day classified: {0:N1}% ({1} matched / {2} needs review)" -f $snapshot.weekCoveragePercent, (Format-Duration $snapshot.weekMatchedSeconds), (Format-Duration $snapshot.weekUnmatchedSeconds)
    $coverageColor = [System.Drawing.Color]::Firebrick
    if ($snapshot.todayCoveragePercent -ge 80) {
        $coverageColor = [System.Drawing.Color]::DarkGreen
    }
    elseif ($snapshot.todayCoveragePercent -ge 50) {
        $coverageColor = [System.Drawing.Color]::DarkOrange
    }
    $State.Controls.HealthTodayCoverageLabel.ForeColor = $coverageColor

    $reviewSummaryText = "Uncategorized today: fully reviewed."
    if ($snapshot.todayUnmatchedCount -gt 0) {
        $reviewSummaryText = "Uncategorized today: {0} grouped items still need rules" -f $snapshot.todayUnmatchedCount
    }
    $State.Controls.HealthReviewSummaryLabel.Text = $reviewSummaryText
    $State.Controls.HealthMainGapLabel.Text = $snapshot.mainGap

    $uncatList = $State.Controls.HealthUncategorizedList
    Invoke-ListViewBatchUpdate -ListView $uncatList -Action {
        $uncatList.Items.Clear()
        foreach ($itemData in @($snapshot.topUncategorized)) {
            $item = New-Object System.Windows.Forms.ListViewItem([string]$itemData.process)
            [void]$item.SubItems.Add((Format-Duration $itemData.seconds))
            [void]$item.SubItems.Add([string]$itemData.examples)
            [void]$item.SubItems.Add([string]$itemData.title)
            [void]$uncatList.Items.Add($item)
        }

        Set-ListViewSort -ListView $uncatList -Column $State.ListSorts.HealthUncategorized.Column -Order $State.ListSorts.HealthUncategorized.Order
    }

    $ruleList = $State.Controls.HealthRulesList
    Invoke-ListViewBatchUpdate -ListView $ruleList -Action {
        $ruleList.Items.Clear()
        foreach ($ruleData in @($snapshot.topRules)) {
            $item = New-Object System.Windows.Forms.ListViewItem([string]$ruleData.rule)
            [void]$item.SubItems.Add((Get-CategoryLabel $ruleData.category))
            [void]$item.SubItems.Add((Format-Duration $ruleData.seconds))
            [void]$item.SubItems.Add([string]$ruleData.matches)
            [void]$ruleList.Items.Add($item)
        }

        Set-ListViewSort -ListView $ruleList -Column $State.ListSorts.HealthRules.Column -Order $State.ListSorts.HealthRules.Order
    }

    $suggestionsSummaryText = "Suggested rules: no strong automatic suggestion yet."
    if ($snapshot.suggestedRuleCount -gt 0) {
        $suggestionsSummaryText = "Suggested rules today: {0} candidates covering {1}" -f $snapshot.suggestedRuleCount, (Format-Duration $snapshot.suggestedRuleSeconds)
        if ($snapshot.suggestionSkippedCount -gt 0) {
            $suggestionsSummaryText += " | $($snapshot.suggestionSkippedCount) items still need manual review"
        }
    }
    $State.Controls.HealthSuggestionsSummaryLabel.Text = $suggestionsSummaryText

    $suggestionsList = $State.Controls.HealthSuggestionsList
    Invoke-ListViewBatchUpdate -ListView $suggestionsList -Action {
        $suggestionsList.Items.Clear()
        foreach ($suggestionData in @($snapshot.topSuggestions)) {
            $item = New-Object System.Windows.Forms.ListViewItem([string]$suggestionData.target)
            [void]$item.SubItems.Add([string]$suggestionData.pattern)
            [void]$item.SubItems.Add((Format-Duration $suggestionData.seconds))
            [void]$item.SubItems.Add([string]$suggestionData.hits)
            [void]$item.SubItems.Add([string]$suggestionData.reason)
            [void]$item.SubItems.Add([string]$suggestionData.example)
            $item.Tag = $suggestionData
            [void]$suggestionsList.Items.Add($item)
        }

        Set-ListViewSort -ListView $suggestionsList -Column $State.ListSorts.HealthSuggestions.Column -Order $State.ListSorts.HealthSuggestions.Order
    }

    $State.Controls.HealthClassifyButton.Enabled = $null -ne $State.LastActivity
    $State.Controls.HealthUseSuggestionButton.Enabled = $suggestionsList.SelectedItems.Count -gt 0
}

function Update-ActivityList {
    param(
        $ActivityList,
        [hashtable]$State
    )

    Invoke-ListViewBatchUpdate -ListView $ActivityList -Action {
        $ActivityList.Items.Clear()
        foreach ($activity in (Get-TopActivities -UsageData $State.UsageData -DateKey $State.CurrentDateKey -Top 12)) {
            $item = New-Object System.Windows.Forms.ListViewItem((Get-CategoryLabel $activity.category))
            [void]$item.SubItems.Add([string]$activity.process)
            [void]$item.SubItems.Add((Get-ActivityDisplayTitle $activity))
            [void]$item.SubItems.Add((Format-Duration $activity.seconds))
            [void]$ActivityList.Items.Add($item)
        }

        Set-ListViewSort -ListView $ActivityList -Column $State.ListSorts.Activity.Column -Order $State.ListSorts.Activity.Order
    }
}

function Update-AppUsageList {
    param(
        $AppList,
        [hashtable]$State
    )

    Invoke-ListViewBatchUpdate -ListView $AppList -Action {
        $AppList.Items.Clear()
        $day = Get-DayStats -UsageData $State.UsageData -DateKey $State.CurrentDateKey
        $total = [double]$day.totals.total

        foreach ($app in (Get-TopProcesses -UsageData $State.UsageData -DateKey $State.CurrentDateKey -Top 15)) {
            $share = "0.0%"
            if ($total -gt 0) {
                $share = "{0:N1}%" -f (($app.seconds / $total) * 100)
            }
            $item = New-Object System.Windows.Forms.ListViewItem([string]$app.process)
            [void]$item.SubItems.Add((Get-CategoryLabel $app.category))
            [void]$item.SubItems.Add((Format-Duration $app.seconds))
            [void]$item.SubItems.Add($share)
            [void]$AppList.Items.Add($item)
        }

        Set-ListViewSort -ListView $AppList -Column $State.ListSorts.Apps.Column -Order $State.ListSorts.Apps.Order
    }
}

function Update-ExactCategoryList {
    param(
        $CategoryList,
        $Items,
        [hashtable]$SortState
    )

    if ($null -eq $CategoryList) {
        return
    }

    Invoke-ListViewBatchUpdate -ListView $CategoryList -Action {
        $CategoryList.Items.Clear()
        foreach ($entry in @($Items)) {
            $item = New-Object System.Windows.Forms.ListViewItem((Get-CategoryLabel $entry.category))
            [void]$item.SubItems.Add((Get-CategoryLabel $entry.parent))
            [void]$item.SubItems.Add((Format-Duration $entry.seconds))
            [void]$CategoryList.Items.Add($item)
        }

        if ($null -ne $SortState) {
            Set-ListViewSort -ListView $CategoryList -Column $SortState.Column -Order $SortState.Order
        }
    }
}

function Update-HistoryList {
    param(
        $HistoryList,
        [hashtable]$State
    )

    Invoke-ListViewBatchUpdate -ListView $HistoryList -Action {
        $HistoryList.Items.Clear()
        foreach ($day in (Get-RecentDaySummaries -UsageData $State.UsageData -Days 7)) {
            $item = New-Object System.Windows.Forms.ListViewItem([string]$day.date)
            [void]$item.SubItems.Add((Format-ShortDuration $day.total))
            [void]$item.SubItems.Add((Format-ShortDuration $day.study))
            [void]$item.SubItems.Add((Format-ShortDuration $day.browser_fun))
            [void]$item.SubItems.Add((Format-ShortDuration $day.socials))
            [void]$HistoryList.Items.Add($item)
        }

        Set-ListViewSort -ListView $HistoryList -Column $State.ListSorts.History.Column -Order $State.ListSorts.History.Order
    }
}

function Update-AutoStartUi {
    param(
        [hashtable]$State
    )

    $enabled = Test-AutoStartEnabled
    $desired = Get-DesiredAutoStartEnabled -Settings $State.Settings
    if ($desired -and $enabled) {
        $State.Controls.AutoStartLabel.Text = "Autostart: enabled"
        $State.Controls.AutoStartLabel.ForeColor = [System.Drawing.Color]::DarkGreen
    }
    elseif ($desired) {
        $State.Controls.AutoStartLabel.Text = "Autostart: repair pending"
        $State.Controls.AutoStartLabel.ForeColor = [System.Drawing.Color]::DarkOrange
    }
    else {
        $State.Controls.AutoStartLabel.Text = "Autostart: disabled"
        $State.Controls.AutoStartLabel.ForeColor = [System.Drawing.Color]::Firebrick
    }

    $autoStartButtonText = "Enable autostart"
    if ($desired) {
        $autoStartButtonText = "Disable autostart"
    }
    $State.Controls.AutoStartButton.Text = $autoStartButtonText
}

function Update-BrowserBridgeUi {
    param(
        [hashtable]$State
    )

    $snapshot = Get-BrowserActivitySnapshot
    $timestamp = Get-ActivityTimestamp -Activity $snapshot
    if ($null -eq $timestamp) {
        $State.Controls.BrowserBridgeLabel.Text = "Browser bridge: waiting for extension"
        $State.Controls.BrowserBridgeLabel.ForeColor = [System.Drawing.Color]::SaddleBrown
        return
    }

    $ageSeconds = [math]::Round(((Get-Date) - $timestamp).TotalSeconds)
    $domain = Get-ActivityField -Activity $snapshot -Name "domain"
    if ([string]::IsNullOrWhiteSpace($domain)) {
        $domain = Get-ActivityField -Activity $snapshot -Name "url"
    }

    $State.Controls.BrowserBridgeLabel.Text = "Browser bridge: {0} ({1}s ago)" -f $domain, $ageSeconds
    $State.Controls.BrowserBridgeLabel.ForeColor = [System.Drawing.Color]::DarkGreen
}

function Update-InsightsUi {
    param(
        [hashtable]$State
    )

    $insights = Get-WeeklyInsightSummary -UsageData $State.UsageData -Days 7
    $State.Controls.InsightTotalLabel.Text = "7-day total: $(Format-Duration $insights.total)"
    $State.Controls.InsightStudyLabel.Text = "Study: $(Format-Duration $insights.study)"
    $State.Controls.InsightAverageLabel.Text = "Daily average: $(Format-Duration $insights.average)"
    $State.Controls.InsightStudyDaysLabel.Text = "Days with 2h+ study: $($insights.studyDays) / 7"

    $bestDayText = "-"
    if ($null -ne $insights.bestDay) {
        $bestDayText = "$($insights.bestDay.date) | $(Format-Duration $insights.bestDay.study) study"
    }
    $State.Controls.InsightBestDayLabel.Text = "Best study day: $bestDayText"

    $topAppText = "-"
    if ($null -ne $insights.topApp) {
        $topAppText = "$([string]$insights.topApp.process) | $(Format-Duration $insights.topApp.seconds)"
    }
    $State.Controls.InsightTopAppLabel.Text = "Top app this week: $topAppText"

    $list = $State.Controls.InsightAppsList
    Invoke-ListViewBatchUpdate -ListView $list -Action {
        $list.Items.Clear()
        foreach ($app in $insights.weeklyApps) {
            $item = New-Object System.Windows.Forms.ListViewItem([string]$app.process)
            [void]$item.SubItems.Add((Get-CategoryLabel $app.category))
            [void]$item.SubItems.Add((Format-Duration $app.seconds))
            [void]$list.Items.Add($item)
        }
    }

    Update-ExactCategoryList -CategoryList $State.Controls.InsightCategoriesList -Items $insights.weeklyCategories -SortState $State.ListSorts.InsightCategories
}

function Update-WeeklyReviewUi {
    param(
        [hashtable]$State
    )

    $review = Get-WeeklyReviewSummary -UsageData $State.UsageData -Settings $State.Settings -Days 7
    $State.Controls.WeeklyReviewSummaryLabel.Text = $review.summaryText
    $State.Controls.WeeklyReviewWinLabel.Text = $review.winText
    $State.Controls.WeeklyReviewSlipLabel.Text = $review.biggestSlipText
    $State.Controls.WeeklyReviewDistractionLabel.Text = $review.mainDistractionText
    $State.Controls.WeeklyReviewCoachLabel.Text = $review.coachNote

    $slipList = $State.Controls.WeeklyReviewDaysList
    Invoke-ListViewBatchUpdate -ListView $slipList -Action {
        $slipList.Items.Clear()
        foreach ($day in @($review.slippedDays)) {
            $item = New-Object System.Windows.Forms.ListViewItem([string]$day.date)
            [void]$item.SubItems.Add([string]$day.issue)
            [void]$item.SubItems.Add((Format-Duration $day.total))
            [void]$item.SubItems.Add((Format-Duration $day.study))
            [void]$slipList.Items.Add($item)
        }

        Set-ListViewSort -ListView $slipList -Column $State.ListSorts.WeeklyReviewDays.Column -Order $State.ListSorts.WeeklyReviewDays.Order
    }

    $appsList = $State.Controls.WeeklyReviewAppsList
    Invoke-ListViewBatchUpdate -ListView $appsList -Action {
        $appsList.Items.Clear()
        foreach ($app in @($review.topDistractingApps)) {
            $item = New-Object System.Windows.Forms.ListViewItem([string]$app.process)
            [void]$item.SubItems.Add((Get-CategoryLabel $app.category))
            [void]$item.SubItems.Add((Format-Duration $app.seconds))
            [void]$appsList.Items.Add($item)
        }

        Set-ListViewSort -ListView $appsList -Column $State.ListSorts.WeeklyReviewApps.Column -Order $State.ListSorts.WeeklyReviewApps.Order
    }

    $heatmap = Get-CalendarHeatmapSummary -UsageData $State.UsageData -Settings $State.Settings -Weeks 6
    $State.Controls.WeeklyReviewHeatmapLegendLabel.Text = $heatmap.legend
    $State.Controls.WeeklyReviewHeatmapSummaryLabel.Text = $heatmap.summaryText
    $labels = @($State.Controls.WeeklyReviewHeatmapLabels)
    for ($index = 0; $index -lt $labels.Count; $index += 1) {
        $label = $labels[$index]
        if ($index -ge @($heatmap.cells).Count) {
            $label.Text = ""
            $label.BackColor = [System.Drawing.Color]::WhiteSmoke
            $label.ForeColor = [System.Drawing.Color]::DarkGray
            continue
        }

        $cell = $heatmap.cells[$index]
        $label.Text = [string]$cell.day
        $label.BackColor = $cell.backColor
        $label.ForeColor = $cell.foreColor
        $label.Tag = $cell
        if ($null -ne $State.Controls.WeeklyReviewHeatmapToolTip) {
            $State.Controls.WeeklyReviewHeatmapToolTip.SetToolTip($label, [string]$cell.tooltip)
        }
    }
}

function Update-GoalsUi {
    param(
        [hashtable]$State
    )

    $day = Get-DayStats -UsageData $State.UsageData -DateKey $State.CurrentDateKey
    $totals = $day.totals
    $limits = $State.Settings.limits
    $summary = Get-GoalsDashboardSummary -UsageData $State.UsageData -DateKey $State.CurrentDateKey -Settings $State.Settings

    $State.Controls.GoalsSummaryLabel.Text = $summary.summary
    $State.Controls.GoalsPrimaryLabel.Text = $summary.primary
    $State.Controls.GoalsUnlockLabel.Text = $summary.unlock
    $goalsSummaryColor = [System.Drawing.Color]::Firebrick
    if ($summary.goalsMet -ge 3) {
        $goalsSummaryColor = [System.Drawing.Color]::DarkGreen
    }
    $State.Controls.GoalsSummaryLabel.ForeColor = $goalsSummaryColor

    Set-RowState -Row $State.Controls.GoalsTotalRow -Current $totals.total -Limit $limits.total -Status (Get-StatusText -Current $totals.total -Limit $limits.total) -Alert ($totals.total -gt [double]$limits.total)
    Set-RowState -Row $State.Controls.GoalsStudyRow -Current $totals.study -Limit $limits.studyMax -Status (Get-StudyStatusText -Current $totals.study -StudyMin $limits.studyMin -StudyMax $limits.studyMax) -Alert ($totals.study -gt [double]$limits.studyMax)
    Set-RowState -Row $State.Controls.GoalsBrowserFunRow -Current $totals.browser_fun -Limit $limits.browser_fun -Status (Get-StatusText -Current $totals.browser_fun -Limit $limits.browser_fun) -Alert ($totals.browser_fun -gt [double]$limits.browser_fun)
    Set-RowState -Row $State.Controls.GoalsSocialsRow -Current $totals.socials -Limit $limits.socials -Status (Get-StatusText -Current $totals.socials -Limit $limits.socials) -Alert ($totals.socials -gt [double]$limits.socials)
}

function Update-AnalyticsUi {
    param(
        [hashtable]$State
    )

    $summary = Get-AnalyticsSummary -UsageData $State.UsageData -Days 30
    $streaks = Get-SessionStreakSummary -UsageData $State.UsageData -Settings $State.Settings -Days 30
    $State.Controls.AnalyticsTotalLabel.Text = "30-day total: $(Format-Duration $summary.totals.total)"
    $State.Controls.AnalyticsAverageLabel.Text = "Daily average: $(Format-Duration $summary.average)"
    $bestStudyText = "-"
    if ($null -ne $summary.bestStudyDay) {
        $bestStudyText = "$($summary.bestStudyDay.date) | $(Format-Duration $summary.bestStudyDay.study) study"
    }
    $State.Controls.AnalyticsBestDayLabel.Text = "Best study day: $bestStudyText"
    $topDistractingText = "-"
    if ($null -ne $summary.topDistractingApp) {
        $topDistractingText = "$([string]$summary.topDistractingApp.process) | $(Format-Duration $summary.topDistractingApp.seconds)"
    }
    $State.Controls.AnalyticsDistractionLabel.Text = "Top distraction: $topDistractingText"
    $State.Controls.AnalyticsLimitStreakLabel.Text = "Under-limit streak: $($streaks.currentUnderLimit) days (best $($streaks.bestUnderLimit))"
    $State.Controls.AnalyticsStudyStreakLabel.Text = "Study-goal streak: $($streaks.currentStudyGoal) days (best $($streaks.bestStudyGoal))"

    Update-StackedTrendChart -Chart $State.Controls.AnalyticsTrendChart -Summaries (Get-RecentDaySummaries -UsageData $State.UsageData -Days 7)
    Update-DoughnutChart -Chart $State.Controls.AnalyticsBreakdownChart -Totals $summary.totals

    $list = $State.Controls.AnalyticsAppsList
    Invoke-ListViewBatchUpdate -ListView $list -Action {
        $list.Items.Clear()
        $periodTotal = [double]$summary.totals.total
        foreach ($app in $summary.topApps) {
            $share = "0.0%"
            if ($periodTotal -gt 0) {
                $share = "{0:N1}%" -f (($app.seconds / $periodTotal) * 100)
            }
            $item = New-Object System.Windows.Forms.ListViewItem([string]$app.process)
            [void]$item.SubItems.Add((Get-CategoryLabel $app.category))
            [void]$item.SubItems.Add((Format-Duration $app.seconds))
            [void]$item.SubItems.Add($share)
            [void]$list.Items.Add($item)
        }

        Set-ListViewSort -ListView $list -Column $State.ListSorts.AnalyticsApps.Column -Order $State.ListSorts.AnalyticsApps.Order
    }
}

function Update-SessionsUi {
    param(
        [hashtable]$State
    )

    $todaySessions = @(Get-RecentSessions -UsageData $State.UsageData -Days 1 -Top 30)
    $weekSessions = @(Get-RecentSessions -UsageData $State.UsageData -Days 7 -Top 15 | Sort-Object -Property seconds -Descending)
    $timelineBuckets = @(Get-TodayTimelineBuckets -UsageData $State.UsageData -DateKey $State.CurrentDateKey)

    $State.Controls.SessionCountLabel.Text = "Sessions today: $($todaySessions.Count)"
    $longestToday = @($todaySessions | Sort-Object -Property seconds -Descending | Select-Object -First 1)
    $sessionLongestText = "Longest today: -"
    if ($longestToday.Count -gt 0) {
        $sessionLongestText = "Longest today: $([string]$longestToday[0].process) | $(Format-ShortDuration $longestToday[0].seconds)"
    }
    $State.Controls.SessionLongestLabel.Text = $sessionLongestText
    $mostActiveBucket = @($timelineBuckets | Sort-Object -Property total -Descending | Select-Object -First 1)
    $timelinePeakText = "Peak hour: -"
    if ($mostActiveBucket.Count -gt 0 -and [double]$mostActiveBucket[0].total -gt 0) {
        $timelinePeakText = "Peak hour: $([string]$mostActiveBucket[0].label) with $(Format-ShortDuration $mostActiveBucket[0].total)"
    }
    $State.Controls.TimelinePeakLabel.Text = $timelinePeakText
    Update-HourlyTimelineChart -Chart $State.Controls.TimelineChart -Buckets $timelineBuckets

    $todayList = $State.Controls.TodaySessionsList
    Invoke-ListViewBatchUpdate -ListView $todayList -Action {
        $todayList.Items.Clear()
        foreach ($session in $todaySessions) {
            $endAt = ConvertTo-SafeDateTime -Value ([string]$session.end)
            $endTime = "-"
            if ($null -ne $endAt) {
                $endTime = $endAt.ToString("HH:mm")
            }
            $item = New-Object System.Windows.Forms.ListViewItem($endTime)
            [void]$item.SubItems.Add((Format-ShortDuration $session.seconds))
            [void]$item.SubItems.Add((Get-CategoryLabel $session.category))
            [void]$item.SubItems.Add([string]$session.process)
            [void]$item.SubItems.Add((Get-ActivityDisplayTitle $session))
            [void]$todayList.Items.Add($item)
        }

        Set-ListViewSort -ListView $todayList -Column $State.ListSorts.TodaySessions.Column -Order $State.ListSorts.TodaySessions.Order
    }

    $weekList = $State.Controls.WeeklySessionsList
    Invoke-ListViewBatchUpdate -ListView $weekList -Action {
        $weekList.Items.Clear()
        foreach ($session in $weekSessions) {
            $dateText = "-"
            if (-not [string]::IsNullOrWhiteSpace([string]$session.date)) {
                $dateText = [string]$session.date
            }
            $item = New-Object System.Windows.Forms.ListViewItem($dateText)
            [void]$item.SubItems.Add((Format-ShortDuration $session.seconds))
            [void]$item.SubItems.Add((Get-CategoryLabel $session.category))
            [void]$item.SubItems.Add([string]$session.process)
            [void]$item.SubItems.Add((Get-ActivityDisplayTitle $session))
            [void]$weekList.Items.Add($item)
        }

        Set-ListViewSort -ListView $weekList -Column $State.ListSorts.WeeklySessions.Column -Order $State.ListSorts.WeeklySessions.Order
    }
}

function Get-ActiveTabKey {
    param(
        [hashtable]$State
    )

    if ($null -eq $State.Controls.MainTabs -or $null -eq $State.Controls.MainTabs.SelectedTab) {
        return "today"
    }

    return (Normalize-Text ([string]$State.Controls.MainTabs.SelectedTab.Text))
}

function Get-ActiveTabRefreshIntervalSeconds {
    param(
        [hashtable]$State
    )

    switch (Get-ActiveTabKey -State $State) {
        "today" { return 2 }
        "goals" { return 3 }
        "timeline" { return 4 }
        "review" { return 8 }
        "health" { return 12 }
        "week review" { return 12 }
        "history" { return 12 }
        "insights" { return 12 }
        "analytics" { return 15 }
        "rules" { return 30 }
        default { return 5 }
    }
}

function Request-UiRefresh {
    param(
        [hashtable]$State
    )

    if ($null -eq $State) {
        return
    }

    $State.LastActivityRefresh = Get-Date "2000-01-01"
    if ($State.ContainsKey("LastSidebarRefresh")) {
        $State.LastSidebarRefresh = Get-Date "2000-01-01"
    }
}

function Update-ActiveTabUi {
    param(
        [hashtable]$State
    )

    switch (Get-ActiveTabKey -State $State) {
        "today" {
            Update-ActivityList -ActivityList $State.Controls.ActivityList -State $State
            Update-AppUsageList -AppList $State.Controls.AppList -State $State
            Update-ExactCategoryList -CategoryList $State.Controls.ExactCategoryList -Items (Get-TopExactCategories -UsageData $State.UsageData -DateKey $State.CurrentDateKey -Top 12) -SortState $State.ListSorts.ExactCategories
        }
        "history" {
            Update-HistoryList -HistoryList $State.Controls.HistoryList -State $State
        }
        "insights" {
            Update-InsightsUi -State $State
        }
        "week review" {
            Update-WeeklyReviewUi -State $State
        }
        "goals" {
            Update-GoalsUi -State $State
        }
        "analytics" {
            Update-AnalyticsUi -State $State
        }
        "timeline" {
            Update-SessionsUi -State $State
        }
        "review" {
            Update-ClassificationReviewUi -State $State
        }
        "health" {
            Update-ClassificationHealthUi -State $State
        }
    }
}

function Update-Ui {
    param(
        [hashtable]$State
    )

    $now = Get-Date
    $day = Get-DayStats -UsageData $State.UsageData -DateKey $State.CurrentDateKey
    $limits = $State.Settings.limits
    $totals = $day.totals
    $quickGlanceVisible = $false
    if ($State.ContainsKey("Controls") -and $State.Controls.ContainsKey("QuickGlanceForm") -and $null -ne $State.Controls.QuickGlanceForm) {
        $quickGlanceVisible = [bool]$State.Controls.QuickGlanceForm.Visible
    }

    if ($null -ne $State.Form -and -not [bool]$State.Form.Visible) {
        Update-TrayStatusUi -State $State
        if ($quickGlanceVisible) {
            Update-QuickGlanceUi -State $State -DayStats $day
        }
        return
    }

    $State.Controls.HeaderLabel.Text = "Date: $($State.CurrentDateKey)"
    $modeText = "Tracking is paused"
    $modeColor = [System.Drawing.Color]::Firebrick
    $pauseButtonText = "Resume tracking"
    if ($State.TrackingEnabled) {
        $modeText = "Tracking is running"
        $modeColor = [System.Drawing.Color]::DarkGreen
        $pauseButtonText = "Pause tracking"
    }
    $State.Controls.ModeLabel.Text = $modeText
    $State.Controls.ModeLabel.ForeColor = $modeColor
    $State.Controls.PauseButton.Text = $pauseButtonText

    Set-RowState -Row $State.Controls.TotalRow -Current $totals.total -Limit $limits.total -Status (Get-StatusText -Current $totals.total -Limit $limits.total) -Alert ($totals.total -gt [double]$limits.total)
    Set-RowState -Row $State.Controls.StudyRow -Current $totals.study -Limit $limits.studyMax -Status (Get-StudyStatusText -Current $totals.study -StudyMin $limits.studyMin -StudyMax $limits.studyMax) -Alert ($totals.study -gt [double]$limits.studyMax)
    Set-RowState -Row $State.Controls.BrowserFunRow -Current $totals.browser_fun -Limit $limits.browser_fun -Status (Get-StatusText -Current $totals.browser_fun -Limit $limits.browser_fun) -Alert ($totals.browser_fun -gt [double]$limits.browser_fun)
    Set-RowState -Row $State.Controls.SocialsRow -Current $totals.socials -Limit $limits.socials -Status (Get-StatusText -Current $totals.socials -Limit $limits.socials) -Alert ($totals.socials -gt [double]$limits.socials)

    $uncategorizedSummary = Get-UncategorizedSnapshotForState -State $State -Top 20
    $otherLabelText = "Uncategorized: $(Format-Duration $totals.other) | fully reviewed today"
    if ($uncategorizedSummary.count -gt 0) {
        $otherLabelText = "Uncategorized: $(Format-Duration $totals.other) | $($uncategorizedSummary.count) items need review"
    }
    $State.Controls.OtherLabel.Text = $otherLabelText
    $State.Controls.TopAppLabel.Text = Get-TopProcessSummaryText -UsageData $State.UsageData -DateKey $State.CurrentDateKey
    $activeProcessText = "-"
    if ($null -ne $State.LastActivity) {
        $activeProcessText = [string]$State.LastActivity.ProcessName
    }
    $activeTitleText = "-"
    if ($null -ne $State.LastActivity -and -not [string]::IsNullOrWhiteSpace($State.LastActivity.WindowTitle)) {
        $activeTitleText = [string]$State.LastActivity.WindowTitle
    }
    $lastDomain = Get-ActivityField -Activity $State.LastActivity -Name "Domain"
    $lastUrl = Get-ActivityField -Activity $State.LastActivity -Name "Url"
    $activeUrlText = "-"
    if (-not [string]::IsNullOrWhiteSpace($lastDomain)) {
        $activeUrlText = $lastDomain
    }
    elseif (-not [string]::IsNullOrWhiteSpace($lastUrl)) {
        $activeUrlText = $lastUrl
    }
    $activeResolution = $null
    if ($null -ne $State.LastActivity) {
        $activeResolution = Resolve-ActivityClassification -Activity $State.LastActivity -Rules $State.Rules
    }
    $activeCategoryColor = [System.Drawing.Color]::SaddleBrown
    if ($null -ne $activeResolution -and [bool]$activeResolution.Matched) {
        $activeCategoryColor = [System.Drawing.Color]::DarkGreen
    }
    $idleText = "Active"
    if ($State.IsIdle) {
        $idleText = "Idle for $(Format-ShortDuration $State.CurrentIdleSeconds)"
    }
    $State.Controls.ActiveProcessValue.Text = $activeProcessText
    $State.Controls.ActiveTitleValue.Text = $activeTitleText
    $State.Controls.ActiveUrlValue.Text = $activeUrlText
    $State.Controls.ActiveCategoryValue.Text = Format-ClassificationSummaryText -Resolution $activeResolution
    $State.Controls.ActiveCategoryValue.ForeColor = $activeCategoryColor
    $State.Controls.IdleValue.Text = $idleText

    if (($now - $State.LastSidebarRefresh).TotalSeconds -ge 5) {
        Update-AutoStartUi -State $State
        Update-BrowserBridgeUi -State $State
        $State.LastSidebarRefresh = $now
    }

    if (($now - $State.LastActivityRefresh).TotalSeconds -ge (Get-ActiveTabRefreshIntervalSeconds -State $State)) {
        Update-ActiveTabUi -State $State
        $State.LastActivityRefresh = $now
    }

    $State.Controls.FocusModeLabel.Text = Get-FocusModeText -State $State
    $focusButtonText = "Start focus"
    if ($State.FocusMode.Enabled) {
        $focusButtonText = "Stop focus"
    }
    $State.Controls.FocusButton.Text = $focusButtonText
    if ($quickGlanceVisible) {
        Update-QuickGlanceUi -State $State -DayStats $day
    }
    Update-TrayStatusUi -State $State
}

function Reset-Today {
    param(
        [hashtable]$State
    )

    Stop-TrackedSession -State $State
    $State.UsageData[$State.CurrentDateKey] = New-DayStats -DateKey $State.CurrentDateKey
    $State.Notifications[$State.CurrentDateKey] = @{
        total = $false
        totalWarn = $false
        browser_fun = $false
        browser_funWarn = $false
        socials = $false
        socialsWarn = $false
        study = $false
        studyWarn = $false
        trayHint = $false
    }
    $State.OverrideUntil = $null
    $State.HardLimitHandledFor = @{}
    Update-UsageSummaryForDay -DateKey $State.CurrentDateKey -DayStats $State.UsageData[$State.CurrentDateKey]
    Clear-DerivedCaches -State $State
    Mark-UsageDateDirty -State $State -DateKey $State.CurrentDateKey
    $State.IsDirty = $true
    Save-IfDirty -State $State -Force
    Update-Ui -State $State
}

function New-SettingsValueLabel {
    param(
        [string]$Text,
        [int]$Left,
        [int]$Top,
        [int]$Width = 200
    )

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.Location = New-Object System.Drawing.Point($Left, $Top)
    $label.Size = New-Object System.Drawing.Size($Width, 22)
    return $label
}

function New-SettingsNumeric {
    param(
        [decimal]$Value,
        [int]$Left,
        [int]$Top,
        [int]$Maximum = 99999
    )

    $box = New-Object System.Windows.Forms.NumericUpDown
    $box.Location = New-Object System.Drawing.Point($Left, $Top)
    $box.Size = New-Object System.Drawing.Size(120, 24)
    $box.Minimum = 1
    $box.Maximum = $Maximum
    $box.Value = [decimal]$Value
    return $box
}

function Show-SettingsEditor {
    param(
        [hashtable]$State
    )

    $settings = ConvertTo-PlainData $State.Settings
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Tracker settings"
    $form.StartPosition = "CenterParent"
    $form.Size = New-Object System.Drawing.Size(560, 660)
    $form.MinimumSize = $form.Size
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false

    $controls = @()
    $controls += New-SettingsValueLabel -Text "Idle timeout (seconds)" -Left 20 -Top 24
    $idleBox = New-SettingsNumeric -Value $settings.idleThresholdSeconds -Left 300 -Top 20 -Maximum 3600
    $controls += $idleBox

    $controls += New-SettingsValueLabel -Text "Sample interval (seconds)" -Left 20 -Top 60
    $sampleBox = New-SettingsNumeric -Value $settings.sampleIntervalSeconds -Left 300 -Top 56 -Maximum 60
    $controls += $sampleBox

    $controls += New-SettingsValueLabel -Text "Total limit (minutes)" -Left 20 -Top 110
    $totalBox = New-SettingsNumeric -Value ([math]::Round($settings.limits.total / 60)) -Left 300 -Top 106 -Maximum 1440
    $controls += $totalBox

    $controls += New-SettingsValueLabel -Text "Study min (minutes)" -Left 20 -Top 146
    $studyMinBox = New-SettingsNumeric -Value ([math]::Round($settings.limits.studyMin / 60)) -Left 300 -Top 142 -Maximum 1440
    $controls += $studyMinBox

    $controls += New-SettingsValueLabel -Text "Study max (minutes)" -Left 20 -Top 182
    $studyMaxBox = New-SettingsNumeric -Value ([math]::Round($settings.limits.studyMax / 60)) -Left 300 -Top 178 -Maximum 1440
    $controls += $studyMaxBox

    $controls += New-SettingsValueLabel -Text "Browser fun limit (minutes)" -Left 20 -Top 218
    $browserFunBox = New-SettingsNumeric -Value ([math]::Round($settings.limits.browser_fun / 60)) -Left 300 -Top 214 -Maximum 1440
    $controls += $browserFunBox

    $controls += New-SettingsValueLabel -Text "Socials limit (minutes)" -Left 20 -Top 254
    $socialsBox = New-SettingsNumeric -Value ([math]::Round($settings.limits.socials / 60)) -Left 300 -Top 250 -Maximum 1440
    $controls += $socialsBox

    $controls += New-SettingsValueLabel -Text "Warning threshold (%)" -Left 20 -Top 304
    $warningBox = New-SettingsNumeric -Value $settings.notifications.warningPercent -Left 300 -Top 300 -Maximum 100
    $controls += $warningBox

    $autoStartCheck = New-Object System.Windows.Forms.CheckBox
    $autoStartCheck.Text = "Start tracker automatically with Windows"
    $autoStartCheck.Location = New-Object System.Drawing.Point(20, 336)
    $autoStartCheck.Size = New-Object System.Drawing.Size(360, 24)
    $autoStartCheck.Checked = [bool](Get-DesiredAutoStartEnabled -Settings $settings)
    $controls += $autoStartCheck

    $hardLimitCheck = New-Object System.Windows.Forms.CheckBox
    $hardLimitCheck.Text = "Enable hard limit dialog after limit is reached"
    $hardLimitCheck.Location = New-Object System.Drawing.Point(20, 368)
    $hardLimitCheck.Size = New-Object System.Drawing.Size(360, 24)
    $hardLimitCheck.Checked = [bool]$settings.notifications.hardLimitEnabled
    $controls += $hardLimitCheck

    $controls += New-SettingsValueLabel -Text "Extra time block after hard limit (minutes)" -Left 20 -Top 406 -Width 260
    $snoozeBox = New-SettingsNumeric -Value $settings.notifications.hardLimitSnoozeMinutes -Left 300 -Top 402 -Maximum 120
    $controls += $snoozeBox

    $closeOnLimitCheck = New-Object System.Windows.Forms.CheckBox
    $closeOnLimitCheck.Text = "Close distracting socials / browser-fun windows after limit"
    $closeOnLimitCheck.Location = New-Object System.Drawing.Point(20, 438)
    $closeOnLimitCheck.Size = New-Object System.Drawing.Size(470, 24)
    $closeOnLimitCheck.Checked = [bool]$settings.notifications.closeDistractingWindows
    $controls += $closeOnLimitCheck

    $controls += New-SettingsValueLabel -Text "Block cooldown (seconds)" -Left 20 -Top 476 -Width 260
    $blockCooldownBox = New-SettingsNumeric -Value $settings.notifications.blockCooldownSeconds -Left 300 -Top 472 -Maximum 600
    $controls += $blockCooldownBox

    $focusCloseCheck = New-Object System.Windows.Forms.CheckBox
    $focusCloseCheck.Text = "Focus mode closes distracting windows instead of only warning"
    $focusCloseCheck.Location = New-Object System.Drawing.Point(20, 508)
    $focusCloseCheck.Size = New-Object System.Drawing.Size(470, 24)
    $focusCloseCheck.Checked = [bool]$settings.focusMode.closeDistractingWindows
    $controls += $focusCloseCheck

    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Text = "Save"
    $saveButton.Location = New-Object System.Drawing.Point(300, 560)
    $saveButton.Size = New-Object System.Drawing.Size(100, 34)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = New-Object System.Drawing.Point(416, 560)
    $cancelButton.Size = New-Object System.Drawing.Size(100, 34)

    foreach ($control in $controls) {
        [void]$form.Controls.Add($control)
    }
    [void]$form.Controls.Add($saveButton)
    [void]$form.Controls.Add($cancelButton)

    $saved = $false
    $saveButton.Add_Click({
        if ($studyMinBox.Value -gt $studyMaxBox.Value) {
            [System.Windows.Forms.MessageBox]::Show("Study min cannot be larger than study max.", "Invalid settings", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }

        $newSettings = @{
            idleThresholdSeconds = [int]$idleBox.Value
            sampleIntervalSeconds = [int]$sampleBox.Value
            browserBridge = @{
                port = [int]$settings.browserBridge.port
                titleMatchToleranceHours = [int]$settings.browserBridge.titleMatchToleranceHours
            }
            autostart = @{
                enabled = [bool]$autoStartCheck.Checked
            }
            notifications = @{
                warningPercent = [int]$warningBox.Value
                hardLimitEnabled = [bool]$hardLimitCheck.Checked
                hardLimitSnoozeMinutes = [int]$snoozeBox.Value
                closeDistractingWindows = [bool]$closeOnLimitCheck.Checked
                blockCooldownSeconds = [int]$blockCooldownBox.Value
            }
            focusMode = @{
                defaultMinutes = [int]$settings.focusMode.defaultMinutes
                promptCooldownSeconds = [int]$settings.focusMode.promptCooldownSeconds
                closeDistractingWindows = [bool]$focusCloseCheck.Checked
            }
            limits = @{
                total = [int]$totalBox.Value * 60
                studyMin = [int]$studyMinBox.Value * 60
                studyMax = [int]$studyMaxBox.Value * 60
                browser_fun = [int]$browserFunBox.Value * 60
                socials = [int]$socialsBox.Value * 60
            }
            categories = @(Normalize-CustomCategories -Categories $settings.categories)
        }

        Save-Settings -Settings $newSettings
        $State.Settings = ConvertTo-PlainData $newSettings
        Sync-AutoStartWithSettings -Settings $State.Settings
        $State.Notifications[$State.CurrentDateKey] = Ensure-NotificationState -State $State -DateKey $State.CurrentDateKey
        $State.HardLimitHandledFor = @{}
        $State.OverrideUntil = $null
        $saved = $true
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })

    $cancelButton.Add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.Close()
    })

    [void]$form.ShowDialog()
    return $saved
}

function Show-CategoriesEditor {
    param(
        [hashtable]$State
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Category editor"
    $form.StartPosition = "CenterParent"
    $form.Size = New-Object System.Drawing.Size(840, 520)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(16, 16)
    $grid.Size = New-Object System.Drawing.Size(790, 390)
    $grid.AllowUserToAddRows = $true
    $grid.AllowUserToDeleteRows = $true
    $grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $grid.RowHeadersVisible = $false

    $keyColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $keyColumn.HeaderText = "Key"
    [void]$grid.Columns.Add($keyColumn)

    $labelColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $labelColumn.HeaderText = "Label"
    [void]$grid.Columns.Add($labelColumn)

    $parentColumn = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
    $parentColumn.HeaderText = "Rule parent"
    [void]$parentColumn.Items.AddRange((Get-ParentCategoryKeys))
    [void]$grid.Columns.Add($parentColumn)

    foreach ($category in @(Normalize-CustomCategories -Categories $State.Settings.categories)) {
        [void]$grid.Rows.Add(@(
            [string]$category.key,
            [string]$category.label,
            [string]$category.parent
        ))
    }

    $introLabel = New-SettingsValueLabel -Text "Base parents stay fixed: study, browser_fun, socials, other. Add your own categories and map each one to a parent." -Left 16 -Top 418 -Width 780

    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Text = "Save categories"
    $saveButton.Location = New-Object System.Drawing.Point(590, 438)
    $saveButton.Size = New-Object System.Drawing.Size(100, 34)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = New-Object System.Drawing.Point(706, 438)
    $cancelButton.Size = New-Object System.Drawing.Size(100, 34)

    [void]$form.Controls.Add($grid)
    [void]$form.Controls.Add($introLabel)
    [void]$form.Controls.Add($saveButton)
    [void]$form.Controls.Add($cancelButton)

    $saved = $false
    $saveButton.Add_Click({
        $categories = @()
        $seenKeys = @{}
        foreach ($row in $grid.Rows) {
            if ($row.IsNewRow) {
                continue
            }

            $rawKey = [string]$row.Cells[0].Value
            $label = [string]$row.Cells[1].Value
            $parent = Normalize-CategoryKey ([string]$row.Cells[2].Value)
            if ([string]::IsNullOrWhiteSpace($rawKey) -and [string]::IsNullOrWhiteSpace($label)) {
                continue
            }

            $key = Normalize-CategoryKey $rawKey
            if ([string]::IsNullOrWhiteSpace($key)) {
                [System.Windows.Forms.MessageBox]::Show("Each category needs a key. Example: coding, reading, messengers.", "Invalid category", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
                return
            }

            if ((Get-ParentCategoryKeys) -contains $key) {
                [System.Windows.Forms.MessageBox]::Show("Built-in parent keys cannot be reused as custom categories.", "Invalid category", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
                return
            }

            if ($seenKeys.ContainsKey($key)) {
                [System.Windows.Forms.MessageBox]::Show("Category keys must be unique.", "Invalid category", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
                return
            }

            if ([string]::IsNullOrWhiteSpace($label)) {
                $label = Convert-CategoryKeyToLabel -Key $key
            }

            if (-not ((Get-ParentCategoryKeys) -contains $parent)) {
                [System.Windows.Forms.MessageBox]::Show("Each custom category must have one of the 4 base rule parents.", "Invalid category", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
                return
            }

            $seenKeys[$key] = $true
            $categories += ,@{
                key = $key
                label = $label.Trim()
                parent = $parent
            }
        }

        $State.Settings["categories"] = @(Normalize-CustomCategories -Categories $categories)
        Save-Settings -Settings $State.Settings
        $State.Rules = @(Normalize-Rules -Rules $State.Rules)
        Save-Rules -Rules $State.Rules
        $saved = $true
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })

    $cancelButton.Add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.Close()
    })

    [void]$form.ShowDialog()
    return $saved
}

function Show-RulesEditor {
    param(
        [hashtable]$State
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Rule editor"
    $form.StartPosition = "CenterParent"
    $form.Size = New-Object System.Drawing.Size(1180, 560)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(16, 16)
    $grid.Size = New-Object System.Drawing.Size(1130, 440)
    $grid.AllowUserToAddRows = $true
    $grid.AllowUserToDeleteRows = $true
    $grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $grid.RowHeadersVisible = $false

    $enabledColumn = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $enabledColumn.HeaderText = "On"
    [void]$grid.Columns.Add($enabledColumn)

    $priorityColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $priorityColumn.HeaderText = "Priority"
    [void]$grid.Columns.Add($priorityColumn)

    $nameColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $nameColumn.HeaderText = "Name"
    [void]$grid.Columns.Add($nameColumn)

    $categoryColumn = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
    $categoryColumn.HeaderText = "Category"
    [void]$categoryColumn.Items.AddRange((Get-RuleCategoryChoices))
    [void]$grid.Columns.Add($categoryColumn)

    $targetColumn = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
    $targetColumn.HeaderText = "Target"
    [void]$targetColumn.Items.AddRange((Get-RuleTargetChoices))
    [void]$grid.Columns.Add($targetColumn)

    $matchModeColumn = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
    $matchModeColumn.HeaderText = "Match mode"
    [void]$matchModeColumn.Items.AddRange((Get-RuleMatchModeChoices))
    [void]$grid.Columns.Add($matchModeColumn)

    $patternsColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $patternsColumn.HeaderText = "Patterns (comma separated)"
    [void]$grid.Columns.Add($patternsColumn)

    foreach ($rule in $State.Rules) {
        [void]$grid.Rows.Add(@(
            [bool]$rule.Enabled,
            [int]$rule.Priority,
            [string]$rule.Name,
            [string]$rule.Category,
            [string]$rule.Target,
            [string]$rule.MatchMode,
            ([string[]]$rule.Patterns -join ", ")
        ))
    }

    $hint = New-SettingsValueLabel -Text "Lowest priority runs first. Use exact/regex only where you really need them." -Left 16 -Top 466 -Width 480
    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Text = "Save rules"
    $saveButton.Location = New-Object System.Drawing.Point(926, 470)
    $saveButton.Size = New-Object System.Drawing.Size(100, 34)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = New-Object System.Drawing.Point(1042, 470)
    $cancelButton.Size = New-Object System.Drawing.Size(100, 34)

    [void]$form.Controls.Add($grid)
    [void]$form.Controls.Add($hint)
    [void]$form.Controls.Add($saveButton)
    [void]$form.Controls.Add($cancelButton)

    $saved = $false
    $saveButton.Add_Click({
        $rules = @()
        foreach ($row in $grid.Rows) {
            if ($row.IsNewRow) {
                continue
            }

            $enabled = $true
            if ($null -ne $row.Cells[0].Value) {
                $enabled = [bool]$row.Cells[0].Value
            }
            $priorityText = [string]$row.Cells[1].Value
            $name = [string]$row.Cells[2].Value
            $category = [string]$row.Cells[3].Value
            $target = [string]$row.Cells[4].Value
            $matchMode = [string]$row.Cells[5].Value
            $patternText = [string]$row.Cells[6].Value

            if ([string]::IsNullOrWhiteSpace($name) -and [string]::IsNullOrWhiteSpace($patternText)) {
                continue
            }

            if ([string]::IsNullOrWhiteSpace($name) -or [string]::IsNullOrWhiteSpace($category) -or [string]::IsNullOrWhiteSpace($target) -or [string]::IsNullOrWhiteSpace($matchMode)) {
                [System.Windows.Forms.MessageBox]::Show("Each non-empty rule row needs Priority, Name, Category, Target, Match mode and Patterns.", "Invalid rule", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
                return
            }

            $priority = 100
            if (-not [int]::TryParse($priorityText, [ref]$priority)) {
                [System.Windows.Forms.MessageBox]::Show("Priority must be a whole number.", "Invalid rule", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
                return
            }

            $patterns = @()
            foreach ($pattern in ($patternText -split ",")) {
                $trimmed = $pattern.Trim()
                if (-not [string]::IsNullOrWhiteSpace($trimmed)) {
                    $patterns += ,$trimmed
                }
            }

            if ($patterns.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("Each rule needs at least one pattern.", "Invalid rule", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
                return
            }

            $rules += ,@{
                Enabled = $enabled
                Priority = $priority
                Name = $name
                Category = $category
                Target = $target
                MatchMode = $matchMode
                Patterns = $patterns
            }
        }

        Save-Rules -Rules $rules
        $State.Rules = @(Normalize-Rules -Rules $rules)
        $saved = $true
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })

    $cancelButton.Add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.Close()
    })

    [void]$form.ShowDialog()
    return $saved
}

function Show-ClassifyCurrentActivityDialog {
    param(
        [hashtable]$State,
        $ActivityOverride = $null,
        [string]$DefaultCategory = "",
        [string]$DialogTitle = "Classify current activity",
        [string]$SuggestedTarget = "",
        [string]$SuggestedPattern = "",
        [string]$SuggestedMatchMode = ""
    )

    $activity = $State.LastActivity
    if ($null -ne $ActivityOverride) {
        $activity = $ActivityOverride
    }

    if ($null -eq $activity) {
        [System.Windows.Forms.MessageBox]::Show("There is no current activity to classify yet.", "Classify current", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        return $false
    }

    $domain = Get-ActivityField -Activity $activity -Name "Domain"
    $url = Get-ActivityField -Activity $activity -Name "Url"

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $DialogTitle
    $form.StartPosition = "CenterParent"
    $form.Size = New-Object System.Drawing.Size(560, 408)
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false

    $processLabel = New-SettingsValueLabel -Text "Process: $([string]$activity.ProcessName)" -Left 20 -Top 20 -Width 500
    $titleLabel = New-SettingsValueLabel -Text "Title: $([string]$activity.WindowTitle)" -Left 20 -Top 52 -Width 500
    $siteText = $domain
    if ([string]::IsNullOrWhiteSpace($siteText)) {
        $siteText = $url
    }
    $siteLabel = New-SettingsValueLabel -Text ("Site: " + $siteText) -Left 20 -Top 84 -Width 500

    $categoryLabel = New-SettingsValueLabel -Text "Category" -Left 20 -Top 126
    $categoryBox = New-Object System.Windows.Forms.ComboBox
    $categoryBox.Location = New-Object System.Drawing.Point(200, 122)
    $categoryBox.Size = New-Object System.Drawing.Size(200, 26)
    $categoryBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$categoryBox.Items.AddRange((Get-RuleCategoryChoices))
    $selectedCategory = "other"
    if (-not [string]::IsNullOrWhiteSpace($DefaultCategory)) {
        $selectedCategory = $DefaultCategory
    }
    elseif ($null -ne $State.LastActivityCategory) {
        $selectedCategory = $State.LastActivityCategory
    }
    $categoryBox.SelectedItem = $selectedCategory

    $createRuleCheck = New-Object System.Windows.Forms.CheckBox
    $createRuleCheck.Text = "Create a permanent rule for this kind of activity"
    $createRuleCheck.Location = New-Object System.Drawing.Point(20, 164)
    $createRuleCheck.Size = New-Object System.Drawing.Size(360, 24)
    $createRuleCheck.Checked = $true

    $applyTodayCheck = New-Object System.Windows.Forms.CheckBox
    $applyTodayCheck.Text = "Also reclassify today's already tracked time for this item"
    $applyTodayCheck.Location = New-Object System.Drawing.Point(20, 194)
    $applyTodayCheck.Size = New-Object System.Drawing.Size(380, 24)
    $applyTodayCheck.Checked = $true

    $targetLabel = New-SettingsValueLabel -Text "Rule target" -Left 20 -Top 234
    $targetBox = New-Object System.Windows.Forms.ComboBox
    $targetBox.Location = New-Object System.Drawing.Point(200, 230)
    $targetBox.Size = New-Object System.Drawing.Size(120, 26)
    $targetBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$targetBox.Items.AddRange((Get-RuleTargetChoices))
    $targetSelection = "process"
    if (-not [string]::IsNullOrWhiteSpace($SuggestedTarget)) {
        $targetSelection = $SuggestedTarget
    }
    elseif (-not [string]::IsNullOrWhiteSpace($domain)) {
        $targetSelection = "domain"
    }
    $targetBox.SelectedItem = $targetSelection

    $matchModeLabel = New-SettingsValueLabel -Text "Match mode" -Left 20 -Top 268
    $matchModeBox = New-Object System.Windows.Forms.ComboBox
    $matchModeBox.Location = New-Object System.Drawing.Point(200, 264)
    $matchModeBox.Size = New-Object System.Drawing.Size(140, 26)
    $matchModeBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$matchModeBox.Items.AddRange((Get-RuleMatchModeChoices))
    $matchModeSelection = "contains"
    if (-not [string]::IsNullOrWhiteSpace($SuggestedMatchMode)) {
        $matchModeSelection = $SuggestedMatchMode
    }
    elseif ($targetSelection -eq "process") {
        $matchModeSelection = "exact"
    }
    $matchModeBox.SelectedItem = $matchModeSelection

    $patternLabel = New-SettingsValueLabel -Text "Pattern" -Left 20 -Top 302
    $patternBox = New-Object System.Windows.Forms.TextBox
    $patternBox.Location = New-Object System.Drawing.Point(200, 298)
    $patternBox.Size = New-Object System.Drawing.Size(320, 24)
    $patternBox.Text = $SuggestedPattern
    if ([string]::IsNullOrWhiteSpace($patternBox.Text)) {
        $patternBox.Text = $domain
    }
    if ([string]::IsNullOrWhiteSpace($patternBox.Text)) {
        $patternBox.Text = [string]$activity.ProcessName
    }

    $targetBox.Add_SelectedIndexChanged({
        switch ([string]$targetBox.SelectedItem) {
            "process" { $patternBox.Text = [string]$activity.ProcessName }
            "window" { $patternBox.Text = [string]$activity.WindowTitle }
            "title" { $patternBox.Text = [string]$activity.WindowTitle }
            "url" {
                $patternBox.Text = $url
                if ([string]::IsNullOrWhiteSpace($patternBox.Text)) {
                    $patternBox.Text = $domain
                }
            }
            "domain" {
                $patternBox.Text = $domain
                if ([string]::IsNullOrWhiteSpace($patternBox.Text)) {
                    $patternBox.Text = $url
                }
            }
            default {
                $patternBox.Text = $domain
                if ([string]::IsNullOrWhiteSpace($patternBox.Text)) {
                    $patternBox.Text = [string]$activity.ProcessName
                }
            }
        }

        switch ([string]$targetBox.SelectedItem) {
            "process" { $matchModeBox.SelectedItem = "exact" }
            "domain" { $matchModeBox.SelectedItem = "contains" }
            default { $matchModeBox.SelectedItem = "contains" }
        }
    })

    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Text = "Apply"
    $saveButton.Location = New-Object System.Drawing.Point(310, 338)
    $saveButton.Size = New-Object System.Drawing.Size(100, 34)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = New-Object System.Drawing.Point(420, 338)
    $cancelButton.Size = New-Object System.Drawing.Size(100, 34)

    foreach ($control in @($processLabel, $titleLabel, $siteLabel, $categoryLabel, $categoryBox, $createRuleCheck, $applyTodayCheck, $targetLabel, $targetBox, $matchModeLabel, $matchModeBox, $patternLabel, $patternBox, $saveButton, $cancelButton)) {
        [void]$form.Controls.Add($control)
    }

    $saved = $false
    $saveButton.Add_Click({
        $category = [string]$categoryBox.SelectedItem
        if ([string]::IsNullOrWhiteSpace($category)) {
            return
        }

        $target = [string]$targetBox.SelectedItem
        $pattern = $patternBox.Text
        $matchMode = [string]$matchModeBox.SelectedItem

        if ($createRuleCheck.Checked) {
            [void](Add-ManualRuleForActivity -State $State -Activity $activity -Category $category -Target $target -Pattern $pattern -MatchMode $matchMode)
        }

        if ($applyTodayCheck.Checked) {
            Move-ExistingActivitiesByRule -State $State -DateKey $State.CurrentDateKey -NewCategory $category -Target $target -Pattern $pattern -MatchMode $matchMode
        }

        $State.LastActivityCategory = $category
        $saved = $true
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })

    $cancelButton.Add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.Close()
    })

    [void]$form.ShowDialog()
    return $saved
}

function Invoke-SelfTestMode {
    Ensure-Storage
    $settings = Load-Settings
    $rules = Load-Rules

    if ($settings.limits.total -ne 10800) { throw "Expected total limit of 10800 seconds." }
    if (-not $settings.ContainsKey("categories")) { throw "Expected categories list in settings." }
    if ($settings.notifications.warningPercent -ne 80) { throw "Expected default warning threshold of 80%." }
    if (-not [bool]$settings.notifications.closeDistractingWindows) { throw "Expected distracting window blocking to default to enabled." }
    if ([int]$settings.notifications.blockCooldownSeconds -lt 1) { throw "Expected a positive block cooldown." }
    if (-not [bool]$settings.focusMode.closeDistractingWindows) { throw "Expected focus mode close behavior to default to enabled." }
    if (-not (Get-DesiredAutoStartEnabled -Settings $settings)) { throw "Expected autostart preference to default to enabled." }
    if ((Get-IdleSeconds) -lt 0) { throw "Idle detection returned a negative value." }
    if (@($rules).Count -eq 0) { throw "Rules failed to load." }
    $selfTestRules = @(Normalize-Rules -Rules @(
        @{ Name = "Study process"; Category = "study"; Target = "process"; MatchMode = "exact"; Priority = 5; Enabled = $true; Patterns = @("code") },
        @{ Name = "Social window"; Category = "socials"; Target = "window"; MatchMode = "contains"; Priority = 10; Enabled = $true; Patterns = @("discord") },
        @{ Name = "Fun window"; Category = "browser_fun"; Target = "window"; MatchMode = "contains"; Priority = 20; Enabled = $true; Patterns = @("mangadex") },
        @{ Name = "Social url"; Category = "socials"; Target = "domain"; MatchMode = "contains"; Priority = 1; Enabled = $true; Patterns = @("discord.com") }
    ))
    if ((Get-CategoryForActivity -Activity @{ ProcessName = "Code"; WindowTitle = "Algorithms notes" } -Rules $selfTestRules) -ne "study") { throw "Study process classification failed." }
    if ((Get-CategoryForActivity -Activity @{ ProcessName = "chrome"; WindowTitle = "Discord - chat" } -Rules $selfTestRules) -ne "socials") { throw "Social classification failed." }
    if ((Get-CategoryForActivity -Activity @{ ProcessName = "msedge"; WindowTitle = "MangaDex - chapter 12" } -Rules $selfTestRules) -ne "browser_fun") { throw "Browser fun classification failed." }
    if ((Get-CategoryForActivity -Activity @{ ProcessName = "chrome"; WindowTitle = "Tab"; Url = "https://discord.com/channels"; Domain = "discord.com" } -Rules $selfTestRules) -ne "socials") { throw "URL classification failed." }
    $discordResolution = Resolve-ActivityClassification -Activity @{ ProcessName = "chrome"; WindowTitle = "Tab"; Url = "https://discord.com/channels"; Domain = "discord.com" } -Rules $selfTestRules
    if (-not $discordResolution.Matched -or [string]::IsNullOrWhiteSpace([string]$discordResolution.RuleName) -or [string]::IsNullOrWhiteSpace([string]$discordResolution.Pattern)) { throw "Classification details failed." }
    $customSettings = ConvertTo-PlainData $settings
    $customSettings["categories"] = @(@{ key = "coding"; label = "Coding"; parent = "study" }, @{ key = "messengers"; label = "Messengers"; parent = "socials" })
    Save-Settings -Settings $customSettings
    $customReloadedSettings = Load-Settings
    if (@($customReloadedSettings.categories).Count -ne 2) { throw "Custom category save/load failed." }
    if ((Get-RuleCategoryChoices) -notcontains "coding") { throw "Custom category choices failed." }
    $customRules = @(@{ Name = "Coding rule"; Category = "coding"; Target = "process"; MatchMode = "exact"; Priority = 5; Enabled = $true; Patterns = @("code") })
    $normalizedCustomRules = @(Normalize-Rules -Rules $customRules)
    if ((Get-CategoryForActivity -Activity @{ ProcessName = "code"; WindowTitle = "Lesson" } -Rules $normalizedCustomRules) -ne "coding") { throw "Custom category classification failed." }
    $priorityRules = @(
        @{ Name = "Loose"; Category = "other"; Target = "process"; MatchMode = "contains"; Priority = 50; Enabled = $true; Patterns = @("code") },
        @{ Name = "Exact"; Category = "study"; Target = "process"; MatchMode = "exact"; Priority = 10; Enabled = $true; Patterns = @("code") },
        @{ Name = "Disabled"; Category = "socials"; Target = "process"; MatchMode = "exact"; Priority = 1; Enabled = $false; Patterns = @("code") }
    )
    $normalizedPriorityRules = @(Normalize-Rules -Rules $priorityRules)
    if ((Get-CategoryForActivity -Activity @{ ProcessName = "Code"; WindowTitle = "Lesson" } -Rules $normalizedPriorityRules) -ne "study") { throw "Rule priority or exact matching failed." }
    if (-not (Get-AutoStartCommand).Contains("-StartMinimized")) { throw "Autostart command is missing minimized mode." }
    if (-not (Get-AutoStartCommand).ToLowerInvariant().Contains("wscript.exe")) { throw "Autostart command should use wscript.exe." }
    Remove-Item -LiteralPath $script:ReopenSignalPath -ErrorAction SilentlyContinue
    if (-not (Request-ReopenRunningInstance)) { throw "Reopen request write failed." }
    if (-not (Consume-ReopenRunningInstanceRequest)) { throw "Reopen request consume failed." }
    if (Consume-ReopenRunningInstanceRequest) { throw "Reopen request should be one-shot." }
    $cooldownState = @{ ActionCooldowns = @{} }
    if (-not (Enter-ActionCooldown -State $cooldownState -Key "selftest" -CooldownSeconds 60)) { throw "Cooldown first entry failed." }
    if (Enter-ActionCooldown -State $cooldownState -Key "selftest" -CooldownSeconds 60) { throw "Cooldown repeat should have been blocked." }

    $selfTestDaysDirectory = Join-Path $script:ExportsDirectory "selftest-days"
    $selfTestSummaryPath = Join-Path $script:ExportsDirectory "selftest-usage-summary-cache.json"
    if (Test-Path -LiteralPath $selfTestDaysDirectory) {
        Remove-Item -LiteralPath $selfTestDaysDirectory -Recurse -Force -ErrorAction SilentlyContinue
    }
    New-Item -ItemType Directory -Path $selfTestDaysDirectory | Out-Null
    "{}" | Set-Content -LiteralPath $selfTestSummaryPath -Encoding UTF8
    $script:UsageDaysDirectory = $selfTestDaysDirectory
    $script:UsageSummaryPath = $selfTestSummaryPath
    $script:UsageSummaryCache = @{}
    $script:UsageDateIndex = @{}
    $usage = @{}
    $todayKey = Get-DateKey
    Add-UsageSample -UsageData $usage -DateKey $todayKey -Activity @{ ProcessName = "Code"; WindowTitle = "Lesson" } -Category "coding" -Seconds 60
    Add-UsageSample -UsageData $usage -DateKey $todayKey -Activity @{ ProcessName = "chrome"; WindowTitle = "Discord" } -Category "socials" -Seconds 30
    $day = Get-DayStats -UsageData $usage -DateKey $todayKey

    if ([math]::Round($day.totals.total) -ne 90) { throw "Total aggregation failed." }
    if ([math]::Round($day.totals.study) -ne 60) { throw "Study aggregation failed." }
    if ([math]::Round($day.totals.coding) -ne 60) { throw "Custom category totals failed." }
    if ([math]::Round($day.totals.socials) -ne 30) { throw "Social aggregation failed." }
    if (@(Get-RecentDaySummaries -UsageData $usage -Days 1).Count -ne 1) { throw "History summary failed." }
    if (-not ((Get-RemainingSummaryText -State @{ UsageData = $usage; CurrentDateKey = $todayKey; Settings = $settings }).Contains("Left today"))) { throw "Tray summary failed." }
    $exactToday = @(Get-TopExactCategories -UsageData $usage -DateKey $todayKey -Top 5)
    if ($exactToday.Count -eq 0 -or [string]$exactToday[0].category -ne "coding") { throw "Top exact categories for today failed." }
    $weekly = Get-WeeklyInsightSummary -UsageData $usage -Days 1
    if ([math]::Round($weekly.total) -ne 90) { throw "Weekly insights aggregation failed." }
    if (@($weekly.weeklyCategories).Count -eq 0 -or [string]$weekly.weeklyCategories[0].category -ne "coding") { throw "Weekly exact categories failed." }
    $reviewDayKey = (Get-Date).AddDays(-1).ToString("yyyy-MM-dd")
    Add-UsageSample -UsageData $usage -DateKey $reviewDayKey -Activity @{ ProcessName = "Telegram"; WindowTitle = "Chats" } -Category "socials" -Seconds 2000
    $weeklyReview = Get-WeeklyReviewSummary -UsageData $usage -Settings $settings -Days 2
    if (@($weeklyReview.slippedDays).Count -lt 1 -or @($weeklyReview.topDistractingApps).Count -lt 1) { throw "Weekly review summary failed." }
    $heatmapSummary = Get-CalendarHeatmapSummary -UsageData $usage -Settings $settings -Weeks 2
    if (@($heatmapSummary.cells).Count -ne 14 -or [string]::IsNullOrWhiteSpace([string]$heatmapSummary.legend)) { throw "Heatmap summary failed." }
    $analytics = Get-AnalyticsSummary -UsageData $usage -Days 1
    if ([math]::Round($analytics.totals.study) -ne 60) { throw "Analytics study aggregation failed." }
    if (@($analytics.topApps).Count -lt 2) { throw "Analytics top apps aggregation failed." }
    if (@($analytics.topExactCategories).Count -eq 0 -or [string]$analytics.topExactCategories[0].category -ne "coding") { throw "Analytics exact categories failed." }
    $goalSummary = Get-GoalsDashboardSummary -UsageData $usage -DateKey $todayKey -Settings $settings
    if ($goalSummary.goalsMet -lt 2) { throw "Goals dashboard summary failed." }
    $sessionDay = Get-DayStats -UsageData $usage -DateKey $todayKey
    $sessionDay.sessions += ,@{
        start = (Get-Date).AddMinutes(-20).ToString("o")
        end = (Get-Date).AddMinutes(-10).ToString("o")
        process = "Code"
        title = "Lesson"
        url = ""
        domain = ""
        category = "study"
        seconds = 600.0
    }
    if (@(Get-RecentSessions -UsageData $usage -Days 1 -Top 5).Count -lt 1) { throw "Recent session listing failed." }
    $timelineBuckets = @(Get-TodayTimelineBuckets -UsageData $usage -DateKey $todayKey)
    if ($timelineBuckets.Count -ne 24) { throw "Timeline buckets were not created." }
    if (($timelineBuckets | Measure-Object -Property total -Sum).Sum -le 0) { throw "Timeline bucket totals failed." }
    $usage[$todayKey]["sessions"] = @{
        start = (Get-Date).AddMinutes(-5).ToString("o")
        end = (Get-Date).ToString("o")
        process = "Explorer"
        title = "Window"
        url = ""
        domain = ""
        category = "other"
        seconds = 42.0
    }
    if (@((Get-DayStats -UsageData $usage -DateKey $todayKey).sessions).Count -ne 1) { throw "Session normalization failed." }
    $usage[$todayKey]["sessions"] = @(@{ start = ""; end = ""; process = ""; title = ""; url = ""; domain = ""; category = "other"; seconds = 0.0 })
    if (@((Get-DayStats -UsageData $usage -DateKey $todayKey).sessions).Count -ne 0) { throw "Invalid session cleanup failed." }
    $streaks = Get-SessionStreakSummary -UsageData $usage -Settings $settings -Days 1
    if ($streaks.bestUnderLimit -lt 1) { throw "Streak summary failed." }
    Add-UsageSample -UsageData $usage -DateKey $todayKey -Activity @{ ProcessName = "Telegram"; WindowTitle = "Friends chat" } -Category "other" -Seconds 45
    if (@(Get-UncategorizedActivities -UsageData $usage -DateKey $todayKey -Rules $normalizedPriorityRules -Top 10).Count -lt 1) { throw "Classification review list failed." }
    $uncategorizedSummary = Get-UncategorizedActivitySummary -UsageData $usage -DateKey $todayKey -Rules $normalizedPriorityRules
    if ($uncategorizedSummary.count -lt 1 -or [math]::Round($uncategorizedSummary.seconds) -lt 30) { throw "Uncategorized summary failed." }
    $suggestions = Get-RuleSuggestions -UsageData $usage -DateKey $todayKey -Rules $normalizedPriorityRules -Top 5
    if ($suggestions.count -lt 1 -or @($suggestions.items).Count -lt 1 -or [string]$suggestions.items[0].target -ne "process") { throw "Rule suggestion generation failed." }
    $healthSnapshot = Get-ClassificationHealthSnapshot -UsageData $usage -DateKey $todayKey -Rules $normalizedPriorityRules -Days 1
    if ($healthSnapshot.todayUnmatchedCount -lt 1 -or @($healthSnapshot.topRules).Count -lt 1 -or @($healthSnapshot.topSuggestions).Count -lt 1) { throw "Classification health snapshot failed." }
    if (-not (Format-ClassificationSummaryText -Resolution $discordResolution).Contains([string]$discordResolution.RuleName)) { throw "Classification summary text failed." }

    $csvPath = Join-Path $script:ExportsDirectory "selftest-report.csv"
    $jsonPath = Join-Path $script:ExportsDirectory "selftest-raw.json"
    $usagePath = Join-Path $script:ExportsDirectory "selftest-usage.json"
    $backupA = Join-Path $script:BackupsDirectory "selftest-backup-a.json"
    $backupB = Join-Path $script:BackupsDirectory "selftest-backup-b.json"
    $backupC = Join-Path $script:BackupsDirectory "selftest-backup-c.json"
    Export-AnalyticsCsv -UsageData $usage -Days 7 -Path $csvPath
    Export-UsageDataJson -UsageData $usage -Path $jsonPath
    Save-JsonFile -Path $usagePath -Data (New-UsageDataEnvelope -UsageData $usage)
    if (-not (Test-Path -LiteralPath $csvPath)) { throw "CSV export failed." }
    if (-not (Test-Path -LiteralPath $jsonPath)) { throw "JSON export failed." }
    $savedUsage = Read-JsonFile -Path $usagePath -Fallback @{}
    if (-not ($savedUsage -is [System.Collections.IDictionary]) -or [int]$savedUsage.schemaVersion -ne $script:UsageSchemaVersion) { throw "Usage schema envelope failed." }
    if (@(Get-Content -LiteralPath $usagePath).Count -gt 2) { throw "Compressed JSON save failed." }
    "{}" | Set-Content -LiteralPath $backupA -Encoding UTF8
    "{}" | Set-Content -LiteralPath $backupB -Encoding UTF8
    "{}" | Set-Content -LiteralPath $backupC -Encoding UTF8
    (Get-Item -LiteralPath $backupA).LastWriteTime = (Get-Date).AddDays(-3)
    (Get-Item -LiteralPath $backupB).LastWriteTime = (Get-Date).AddDays(-2)
    (Get-Item -LiteralPath $backupC).LastWriteTime = (Get-Date).AddDays(-1)
    Trim-BackupFiles -Directory $script:BackupsDirectory -Prefix "selftest-backup" -Keep 2
    if (@(Get-ChildItem -LiteralPath $script:BackupsDirectory -File -Filter "selftest-backup-*").Count -ne 2) { throw "Backup rotation failed." }
    Remove-Item -LiteralPath $csvPath -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $jsonPath -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $usagePath -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $backupA -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $backupB -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $backupC -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $selfTestSummaryPath -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $selfTestDaysDirectory -Recurse -Force -ErrorAction SilentlyContinue
    Save-Settings -Settings $settings

    Write-Output "Self-test passed."
}

function New-MetricRow {
    param(
        [string]$Title,
        [int]$Top
    )

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = $Title
    $titleLabel.Location = New-Object System.Drawing.Point(18, $Top)
    $titleLabel.Size = New-Object System.Drawing.Size(160, 20)
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)

    $valueLabel = New-Object System.Windows.Forms.Label
    $valueLabel.Text = "00:00:00 / 00:00:00"
    $valueLabel.Location = New-Object System.Drawing.Point(190, $Top)
    $valueLabel.Size = New-Object System.Drawing.Size(170, 20)

    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(18, ($Top + 24))
    $progressBar.Size = New-Object System.Drawing.Size(342, 18)
    $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous

    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Text = "Within limit"
    $statusLabel.Location = New-Object System.Drawing.Point(18, ($Top + 48))
    $statusLabel.Size = New-Object System.Drawing.Size(342, 18)

    return @{
        Controls = @($titleLabel, $valueLabel, $progressBar, $statusLabel)
        ValueLabel = $valueLabel
        ProgressBar = $progressBar
        StatusLabel = $statusLabel
    }
}

function New-TrackerState {
    Ensure-Storage

    $currentDateKey = Get-DateKey
    $state = @{
        Settings = Load-Settings
        Rules = Load-Rules
        UsageData = Load-UsageData
        CurrentDateKey = $currentDateKey
        TrackingEnabled = $true
        LastTickAt = Get-Date
        LastSavedAt = Get-Date
        LastActivityRefresh = Get-Date "2000-01-01"
        LastSidebarRefresh = Get-Date "2000-01-01"
        LastTrayTooltipText = ""
        LastActivity = $null
        LastActivityCategory = $null
        CurrentSession = $null
        CurrentIdleSeconds = 0.0
        IsIdle = $false
        Notifications = @{}
        IsDirty = $false
        DirtyDateKeys = @{}
        DerivedCaches = @{}
        Controls = @{}
        Form = $null
        QuickGlanceForm = $null
        BrowserBridgeJob = $null
        AllowExit = $false
        OverrideUntil = $null
        HardLimitHandledFor = @{}
        ActionCooldowns = @{}
        FocusMode = @{
            Enabled = $false
            Until = $null
            LastPromptAt = Get-Date "2000-01-01"
        }
        ListSorts = @{
            Activity = @{ Column = 3; Order = [System.Windows.Forms.SortOrder]::Descending }
            Apps = @{ Column = 2; Order = [System.Windows.Forms.SortOrder]::Descending }
            ExactCategories = @{ Column = 2; Order = [System.Windows.Forms.SortOrder]::Descending }
            History = @{ Column = 0; Order = [System.Windows.Forms.SortOrder]::Ascending }
            WeeklyReviewDays = @{ Column = 0; Order = [System.Windows.Forms.SortOrder]::Ascending }
            WeeklyReviewApps = @{ Column = 2; Order = [System.Windows.Forms.SortOrder]::Descending }
            AnalyticsApps = @{ Column = 2; Order = [System.Windows.Forms.SortOrder]::Descending }
            InsightCategories = @{ Column = 2; Order = [System.Windows.Forms.SortOrder]::Descending }
            TodaySessions = @{ Column = 0; Order = [System.Windows.Forms.SortOrder]::Descending }
            WeeklySessions = @{ Column = 1; Order = [System.Windows.Forms.SortOrder]::Descending }
            Review = @{ Column = 2; Order = [System.Windows.Forms.SortOrder]::Descending }
            HealthUncategorized = @{ Column = 1; Order = [System.Windows.Forms.SortOrder]::Descending }
            HealthRules = @{ Column = 2; Order = [System.Windows.Forms.SortOrder]::Descending }
            HealthSuggestions = @{ Column = 2; Order = [System.Windows.Forms.SortOrder]::Descending }
        }
    }

    Sync-AutoStartWithSettings -Settings $state.Settings
    [void](Get-DayStats -UsageData $state.UsageData -DateKey $currentDateKey)
    [void](Ensure-NotificationState -State $state -DateKey $currentDateKey)
    return $state
}

function Show-MainWindow {
    $state = New-TrackerState

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $script:WindowTitle
    $form.StartPosition = "CenterScreen"
    $form.Size = New-Object System.Drawing.Size(1300, 900)
    $form.MinimumSize = New-Object System.Drawing.Size(1220, 860)
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $form.BackColor = [System.Drawing.Color]::WhiteSmoke
    $form.Icon = Get-AppIcon
    $state.Form = $form

    $quickGlanceForm = New-Object System.Windows.Forms.Form
    $quickGlanceForm.Text = "Quick glance"
    $quickGlanceForm.StartPosition = "Manual"
    $quickGlanceForm.Size = New-Object System.Drawing.Size(340, 266)
    $quickGlanceForm.MinimumSize = New-Object System.Drawing.Size(340, 266)
    $quickGlanceForm.MaximumSize = New-Object System.Drawing.Size(340, 266)
    $quickGlanceForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedToolWindow
    $quickGlanceForm.TopMost = $true
    $quickGlanceForm.ShowInTaskbar = $false
    $quickGlanceForm.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $quickGlanceForm.BackColor = [System.Drawing.Color]::WhiteSmoke
    $quickGlanceForm.Icon = Get-AppIcon
    $quickGlanceForm.Visible = $false

    $quickTitleLabel = New-Object System.Windows.Forms.Label
    $quickTitleLabel.Location = New-Object System.Drawing.Point(16, 14)
    $quickTitleLabel.Size = New-Object System.Drawing.Size(292, 24)
    $quickTitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $quickTitleLabel.Text = "Quick glance"

    $quickDateLabel = New-Object System.Windows.Forms.Label
    $quickDateLabel.Location = New-Object System.Drawing.Point(16, 42)
    $quickDateLabel.Size = New-Object System.Drawing.Size(292, 20)

    $quickTotalLabel = New-Object System.Windows.Forms.Label
    $quickTotalLabel.Location = New-Object System.Drawing.Point(16, 72)
    $quickTotalLabel.Size = New-Object System.Drawing.Size(292, 20)
    $quickTotalLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)

    $quickStudyLabel = New-Object System.Windows.Forms.Label
    $quickStudyLabel.Location = New-Object System.Drawing.Point(16, 100)
    $quickStudyLabel.Size = New-Object System.Drawing.Size(292, 20)

    $quickFunLabel = New-Object System.Windows.Forms.Label
    $quickFunLabel.Location = New-Object System.Drawing.Point(16, 124)
    $quickFunLabel.Size = New-Object System.Drawing.Size(292, 20)

    $quickSocialsLabel = New-Object System.Windows.Forms.Label
    $quickSocialsLabel.Location = New-Object System.Drawing.Point(16, 148)
    $quickSocialsLabel.Size = New-Object System.Drawing.Size(292, 20)

    $quickCurrentLabel = New-Object System.Windows.Forms.Label
    $quickCurrentLabel.Location = New-Object System.Drawing.Point(16, 176)
    $quickCurrentLabel.Size = New-Object System.Drawing.Size(292, 34)

    $quickFocusLabel = New-Object System.Windows.Forms.Label
    $quickFocusLabel.Location = New-Object System.Drawing.Point(16, 208)
    $quickFocusLabel.Size = New-Object System.Drawing.Size(182, 20)

    $quickStatusLabel = New-Object System.Windows.Forms.Label
    $quickStatusLabel.Location = New-Object System.Drawing.Point(16, 228)
    $quickStatusLabel.Size = New-Object System.Drawing.Size(182, 20)

    $quickOpenButton = New-Object System.Windows.Forms.Button
    $quickOpenButton.Text = "Open"
    $quickOpenButton.Location = New-Object System.Drawing.Point(214, 204)
    $quickOpenButton.Size = New-Object System.Drawing.Size(94, 26)

    $quickPauseButton = New-Object System.Windows.Forms.Button
    $quickPauseButton.Text = "Pause"
    $quickPauseButton.Location = New-Object System.Drawing.Point(214, 232)
    $quickPauseButton.Size = New-Object System.Drawing.Size(94, 26)

    foreach ($control in @($quickTitleLabel, $quickDateLabel, $quickTotalLabel, $quickStudyLabel, $quickFunLabel, $quickSocialsLabel, $quickCurrentLabel, $quickFocusLabel, $quickStatusLabel, $quickOpenButton, $quickPauseButton)) {
        [void]$quickGlanceForm.Controls.Add($control)
    }
    $state.QuickGlanceForm = $quickGlanceForm

    $headerLabel = New-Object System.Windows.Forms.Label
    $headerLabel.Location = New-Object System.Drawing.Point(20, 16)
    $headerLabel.Size = New-Object System.Drawing.Size(220, 26)
    $headerLabel.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)

    $modeLabel = New-Object System.Windows.Forms.Label
    $modeLabel.Location = New-Object System.Drawing.Point(250, 20)
    $modeLabel.Size = New-Object System.Drawing.Size(240, 24)
    $modeLabel.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)

    $browserBridgeLabel = New-Object System.Windows.Forms.Label
    $browserBridgeLabel.Location = New-Object System.Drawing.Point(520, 20)
    $browserBridgeLabel.Size = New-Object System.Drawing.Size(720, 24)
    $browserBridgeLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Italic)

    $focusModeLabel = New-Object System.Windows.Forms.Label
    $focusModeLabel.Location = New-Object System.Drawing.Point(20, 744)
    $focusModeLabel.Size = New-Object System.Drawing.Size(220, 24)
    $focusModeLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)

    $summaryGroup = New-Object System.Windows.Forms.GroupBox
    $summaryGroup.Text = "Today's limits"
    $summaryGroup.Location = New-Object System.Drawing.Point(20, 60)
    $summaryGroup.Size = New-Object System.Drawing.Size(390, 350)

    $totalRow = New-MetricRow -Title "Total computer time" -Top 28
    $studyRow = New-MetricRow -Title "Study target (2:00-2:30)" -Top 108
    $browserFunRow = New-MetricRow -Title "Browser fun / manga" -Top 188
    $socialsRow = New-MetricRow -Title "Social media" -Top 268
    foreach ($control in ($totalRow.Controls + $studyRow.Controls + $browserFunRow.Controls + $socialsRow.Controls)) {
        [void]$summaryGroup.Controls.Add($control)
    }

    $otherLabel = New-Object System.Windows.Forms.Label
    $otherLabel.Location = New-Object System.Drawing.Point(20, 420)
    $otherLabel.Size = New-Object System.Drawing.Size(390, 34)
    $otherLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Italic)

    $topAppLabel = New-Object System.Windows.Forms.Label
    $topAppLabel.Location = New-Object System.Drawing.Point(20, 454)
    $topAppLabel.Size = New-Object System.Drawing.Size(390, 28)
    $topAppLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)

    $activeGroup = New-Object System.Windows.Forms.GroupBox
    $activeGroup.Text = "Current activity"
    $activeGroup.Location = New-Object System.Drawing.Point(20, 492)
    $activeGroup.Size = New-Object System.Drawing.Size(390, 246)

    $activeProcessLabel = New-Object System.Windows.Forms.Label
    $activeProcessLabel.Text = "Process:"
    $activeProcessLabel.Location = New-Object System.Drawing.Point(16, 32)
    $activeProcessLabel.Size = New-Object System.Drawing.Size(110, 20)

    $activeProcessValue = New-Object System.Windows.Forms.Label
    $activeProcessValue.Location = New-Object System.Drawing.Point(130, 32)
    $activeProcessValue.Size = New-Object System.Drawing.Size(240, 20)

    $activeTitleLabel = New-Object System.Windows.Forms.Label
    $activeTitleLabel.Text = "Window:"
    $activeTitleLabel.Location = New-Object System.Drawing.Point(16, 70)
    $activeTitleLabel.Size = New-Object System.Drawing.Size(110, 20)

    $activeTitleValue = New-Object System.Windows.Forms.Label
    $activeTitleValue.Location = New-Object System.Drawing.Point(130, 70)
    $activeTitleValue.Size = New-Object System.Drawing.Size(240, 44)
    $activeTitleValue.AutoEllipsis = $true

    $activeUrlLabel = New-Object System.Windows.Forms.Label
    $activeUrlLabel.Text = "Site / URL:"
    $activeUrlLabel.Location = New-Object System.Drawing.Point(16, 118)
    $activeUrlLabel.Size = New-Object System.Drawing.Size(110, 20)

    $activeUrlValue = New-Object System.Windows.Forms.Label
    $activeUrlValue.Location = New-Object System.Drawing.Point(130, 118)
    $activeUrlValue.Size = New-Object System.Drawing.Size(240, 28)
    $activeUrlValue.AutoEllipsis = $true

    $activeCategoryLabel = New-Object System.Windows.Forms.Label
    $activeCategoryLabel.Text = "Category / rule:"
    $activeCategoryLabel.Location = New-Object System.Drawing.Point(16, 154)
    $activeCategoryLabel.Size = New-Object System.Drawing.Size(110, 20)

    $activeCategoryValue = New-Object System.Windows.Forms.Label
    $activeCategoryValue.Location = New-Object System.Drawing.Point(130, 154)
    $activeCategoryValue.Size = New-Object System.Drawing.Size(240, 40)

    $idleLabel = New-Object System.Windows.Forms.Label
    $idleLabel.Text = "Status:"
    $idleLabel.Location = New-Object System.Drawing.Point(16, 202)
    $idleLabel.Size = New-Object System.Drawing.Size(110, 20)

    $idleValue = New-Object System.Windows.Forms.Label
    $idleValue.Location = New-Object System.Drawing.Point(130, 202)
    $idleValue.Size = New-Object System.Drawing.Size(240, 20)

    foreach ($control in @($activeProcessLabel, $activeProcessValue, $activeTitleLabel, $activeTitleValue, $activeUrlLabel, $activeUrlValue, $activeCategoryLabel, $activeCategoryValue, $idleLabel, $idleValue)) {
        [void]$activeGroup.Controls.Add($control)
    }

    $mainTabs = New-Object System.Windows.Forms.TabControl
    $mainTabs.Location = New-Object System.Drawing.Point(430, 60)
    $mainTabs.Size = New-Object System.Drawing.Size(840, 688)
    $mainTabs.Font = New-Object System.Drawing.Font("Segoe UI", 10)

    $todayTab = New-Object System.Windows.Forms.TabPage
    $todayTab.Text = "Today"
    $todayTab.BackColor = [System.Drawing.Color]::WhiteSmoke

    $historyTab = New-Object System.Windows.Forms.TabPage
    $historyTab.Text = "History"
    $historyTab.BackColor = [System.Drawing.Color]::WhiteSmoke

    $insightsTab = New-Object System.Windows.Forms.TabPage
    $insightsTab.Text = "Insights"
    $insightsTab.BackColor = [System.Drawing.Color]::WhiteSmoke

    $weekReviewTab = New-Object System.Windows.Forms.TabPage
    $weekReviewTab.Text = "Week review"
    $weekReviewTab.BackColor = [System.Drawing.Color]::WhiteSmoke

    $goalsTab = New-Object System.Windows.Forms.TabPage
    $goalsTab.Text = "Goals"
    $goalsTab.BackColor = [System.Drawing.Color]::WhiteSmoke

    $analyticsTab = New-Object System.Windows.Forms.TabPage
    $analyticsTab.Text = "Analytics"
    $analyticsTab.BackColor = [System.Drawing.Color]::WhiteSmoke

    $sessionsTab = New-Object System.Windows.Forms.TabPage
    $sessionsTab.Text = "Timeline"
    $sessionsTab.BackColor = [System.Drawing.Color]::WhiteSmoke

    $reviewTab = New-Object System.Windows.Forms.TabPage
    $reviewTab.Text = "Review"
    $reviewTab.BackColor = [System.Drawing.Color]::WhiteSmoke

    $healthTab = New-Object System.Windows.Forms.TabPage
    $healthTab.Text = "Health"
    $healthTab.BackColor = [System.Drawing.Color]::WhiteSmoke

    $rulesTab = New-Object System.Windows.Forms.TabPage
    $rulesTab.Text = "Rules"
    $rulesTab.BackColor = [System.Drawing.Color]::WhiteSmoke

    [void]$mainTabs.TabPages.Add($todayTab)
    [void]$mainTabs.TabPages.Add($historyTab)
    [void]$mainTabs.TabPages.Add($insightsTab)
    [void]$mainTabs.TabPages.Add($weekReviewTab)
    [void]$mainTabs.TabPages.Add($goalsTab)
    [void]$mainTabs.TabPages.Add($analyticsTab)
    [void]$mainTabs.TabPages.Add($sessionsTab)
    [void]$mainTabs.TabPages.Add($reviewTab)
    [void]$mainTabs.TabPages.Add($healthTab)
    [void]$mainTabs.TabPages.Add($rulesTab)

    $activityGroup = New-Object System.Windows.Forms.GroupBox
    $activityGroup.Text = "Top activities today"
    $activityGroup.Location = New-Object System.Drawing.Point(12, 12)
    $activityGroup.Size = New-Object System.Drawing.Size(796, 278)

    $activityList = New-Object System.Windows.Forms.ListView
    $activityList.Location = New-Object System.Drawing.Point(14, 28)
    $activityList.Size = New-Object System.Drawing.Size(766, 234)
    $activityList.View = [System.Windows.Forms.View]::Details
    $activityList.FullRowSelect = $true
    $activityList.GridLines = $true
    $activityList.HideSelection = $false
    $activityList.Tag = "string|string|string|time"
    [void]$activityList.Columns.Add("Category", 120)
    [void]$activityList.Columns.Add("Process", 120)
    [void]$activityList.Columns.Add("Window title / domain", 440)
    [void]$activityList.Columns.Add("Time", 90)
    [void]$activityGroup.Controls.Add($activityList)

    $appsGroup = New-Object System.Windows.Forms.GroupBox
    $appsGroup.Text = "Top apps today"
    $appsGroup.Location = New-Object System.Drawing.Point(12, 304)
    $appsGroup.Size = New-Object System.Drawing.Size(392, 312)

    $appList = New-Object System.Windows.Forms.ListView
    $appList.Location = New-Object System.Drawing.Point(14, 28)
    $appList.Size = New-Object System.Drawing.Size(362, 268)
    $appList.View = [System.Windows.Forms.View]::Details
    $appList.FullRowSelect = $true
    $appList.GridLines = $true
    $appList.HideSelection = $false
    $appList.Tag = "string|string|time|percent"
    [void]$appList.Columns.Add("Process", 136)
    [void]$appList.Columns.Add("Category", 92)
    [void]$appList.Columns.Add("Time", 80)
    [void]$appList.Columns.Add("Share", 54)
    [void]$appsGroup.Controls.Add($appList)

    $exactCategoriesGroup = New-Object System.Windows.Forms.GroupBox
    $exactCategoriesGroup.Text = "Top categories today"
    $exactCategoriesGroup.Location = New-Object System.Drawing.Point(416, 304)
    $exactCategoriesGroup.Size = New-Object System.Drawing.Size(392, 312)

    $exactCategoriesList = New-Object System.Windows.Forms.ListView
    $exactCategoriesList.Location = New-Object System.Drawing.Point(14, 28)
    $exactCategoriesList.Size = New-Object System.Drawing.Size(362, 268)
    $exactCategoriesList.View = [System.Windows.Forms.View]::Details
    $exactCategoriesList.FullRowSelect = $true
    $exactCategoriesList.GridLines = $true
    $exactCategoriesList.HideSelection = $false
    $exactCategoriesList.Tag = "string|string|time"
    [void]$exactCategoriesList.Columns.Add("Category", 166)
    [void]$exactCategoriesList.Columns.Add("Parent", 104)
    [void]$exactCategoriesList.Columns.Add("Time", 88)
    [void]$exactCategoriesGroup.Controls.Add($exactCategoriesList)

    [void]$todayTab.Controls.Add($activityGroup)
    [void]$todayTab.Controls.Add($appsGroup)
    [void]$todayTab.Controls.Add($exactCategoriesGroup)

    $historyGroup = New-Object System.Windows.Forms.GroupBox
    $historyGroup.Text = "Last 7 days"
    $historyGroup.Location = New-Object System.Drawing.Point(12, 12)
    $historyGroup.Size = New-Object System.Drawing.Size(796, 604)

    $historyList = New-Object System.Windows.Forms.ListView
    $historyList.Location = New-Object System.Drawing.Point(14, 28)
    $historyList.Size = New-Object System.Drawing.Size(766, 560)
    $historyList.View = [System.Windows.Forms.View]::Details
    $historyList.FullRowSelect = $true
    $historyList.GridLines = $true
    $historyList.HideSelection = $false
    $historyList.Tag = "date|time|time|time|time"
    [void]$historyList.Columns.Add("Day", 140)
    [void]$historyList.Columns.Add("Total", 120)
    [void]$historyList.Columns.Add("Study", 120)
    [void]$historyList.Columns.Add("Browser fun", 150)
    [void]$historyList.Columns.Add("Socials", 120)
    [void]$historyGroup.Controls.Add($historyList)
    [void]$historyTab.Controls.Add($historyGroup)

    $insightTotalLabel = New-SettingsValueLabel -Text "7-day total:" -Left 18 -Top 24 -Width 320
    $insightTotalLabel.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $insightStudyLabel = New-SettingsValueLabel -Text "Study:" -Left 18 -Top 58 -Width 320
    $insightAverageLabel = New-SettingsValueLabel -Text "Daily average:" -Left 18 -Top 92 -Width 320
    $insightStudyDaysLabel = New-SettingsValueLabel -Text "Days with 2h+ study:" -Left 18 -Top 126 -Width 320
    $insightBestDayLabel = New-SettingsValueLabel -Text "Best study day:" -Left 18 -Top 160 -Width 760
    $insightTopAppLabel = New-SettingsValueLabel -Text "Top app this week:" -Left 18 -Top 194 -Width 760

    $insightAppsGroup = New-Object System.Windows.Forms.GroupBox
    $insightAppsGroup.Text = "Top apps this week"
    $insightAppsGroup.Location = New-Object System.Drawing.Point(12, 236)
    $insightAppsGroup.Size = New-Object System.Drawing.Size(392, 380)

    $insightAppsList = New-Object System.Windows.Forms.ListView
    $insightAppsList.Location = New-Object System.Drawing.Point(14, 28)
    $insightAppsList.Size = New-Object System.Drawing.Size(362, 336)
    $insightAppsList.View = [System.Windows.Forms.View]::Details
    $insightAppsList.FullRowSelect = $true
    $insightAppsList.GridLines = $true
    $insightAppsList.HideSelection = $false
    [void]$insightAppsList.Columns.Add("Process", 150)
    [void]$insightAppsList.Columns.Add("Category", 120)
    [void]$insightAppsList.Columns.Add("Time", 88)
    [void]$insightAppsGroup.Controls.Add($insightAppsList)

    $insightCategoriesGroup = New-Object System.Windows.Forms.GroupBox
    $insightCategoriesGroup.Text = "Top categories this week"
    $insightCategoriesGroup.Location = New-Object System.Drawing.Point(416, 236)
    $insightCategoriesGroup.Size = New-Object System.Drawing.Size(392, 380)

    $insightCategoriesList = New-Object System.Windows.Forms.ListView
    $insightCategoriesList.Location = New-Object System.Drawing.Point(14, 28)
    $insightCategoriesList.Size = New-Object System.Drawing.Size(362, 336)
    $insightCategoriesList.View = [System.Windows.Forms.View]::Details
    $insightCategoriesList.FullRowSelect = $true
    $insightCategoriesList.GridLines = $true
    $insightCategoriesList.HideSelection = $false
    $insightCategoriesList.Tag = "string|string|time"
    [void]$insightCategoriesList.Columns.Add("Category", 166)
    [void]$insightCategoriesList.Columns.Add("Parent", 104)
    [void]$insightCategoriesList.Columns.Add("Time", 88)
    [void]$insightCategoriesGroup.Controls.Add($insightCategoriesList)

    foreach ($control in @($insightTotalLabel, $insightStudyLabel, $insightAverageLabel, $insightStudyDaysLabel, $insightBestDayLabel, $insightTopAppLabel, $insightAppsGroup, $insightCategoriesGroup)) {
        [void]$insightsTab.Controls.Add($control)
    }

    $weeklyReviewSummaryLabel = New-Object System.Windows.Forms.Label
    $weeklyReviewSummaryLabel.Location = New-Object System.Drawing.Point(18, 18)
    $weeklyReviewSummaryLabel.Size = New-Object System.Drawing.Size(760, 24)
    $weeklyReviewSummaryLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $weeklyReviewSummaryLabel.Text = "Week review:"

    $weeklyReviewWinLabel = New-Object System.Windows.Forms.Label
    $weeklyReviewWinLabel.Location = New-Object System.Drawing.Point(18, 50)
    $weeklyReviewWinLabel.Size = New-Object System.Drawing.Size(760, 22)
    $weeklyReviewWinLabel.Text = "Best win:"

    $weeklyReviewSlipLabel = New-Object System.Windows.Forms.Label
    $weeklyReviewSlipLabel.Location = New-Object System.Drawing.Point(18, 80)
    $weeklyReviewSlipLabel.Size = New-Object System.Drawing.Size(760, 22)
    $weeklyReviewSlipLabel.Text = "Biggest slip:"

    $weeklyReviewDistractionLabel = New-Object System.Windows.Forms.Label
    $weeklyReviewDistractionLabel.Location = New-Object System.Drawing.Point(18, 110)
    $weeklyReviewDistractionLabel.Size = New-Object System.Drawing.Size(760, 22)
    $weeklyReviewDistractionLabel.Text = "Main distraction:"

    $weeklyReviewCoachLabel = New-Object System.Windows.Forms.Label
    $weeklyReviewCoachLabel.Location = New-Object System.Drawing.Point(18, 140)
    $weeklyReviewCoachLabel.Size = New-Object System.Drawing.Size(760, 36)
    $weeklyReviewCoachLabel.Text = "Coach note:"

    $weeklyReviewDaysGroup = New-Object System.Windows.Forms.GroupBox
    $weeklyReviewDaysGroup.Text = "Days to revisit"
    $weeklyReviewDaysGroup.Location = New-Object System.Drawing.Point(12, 190)
    $weeklyReviewDaysGroup.Size = New-Object System.Drawing.Size(392, 208)

    $weeklyReviewDaysList = New-Object System.Windows.Forms.ListView
    $weeklyReviewDaysList.Location = New-Object System.Drawing.Point(14, 28)
    $weeklyReviewDaysList.Size = New-Object System.Drawing.Size(362, 164)
    $weeklyReviewDaysList.View = [System.Windows.Forms.View]::Details
    $weeklyReviewDaysList.FullRowSelect = $true
    $weeklyReviewDaysList.GridLines = $true
    $weeklyReviewDaysList.HideSelection = $false
    $weeklyReviewDaysList.Tag = "date|string|time|time"
    [void]$weeklyReviewDaysList.Columns.Add("Day", 84)
    [void]$weeklyReviewDaysList.Columns.Add("Issue", 136)
    [void]$weeklyReviewDaysList.Columns.Add("Total", 70)
    [void]$weeklyReviewDaysList.Columns.Add("Study", 70)
    [void]$weeklyReviewDaysGroup.Controls.Add($weeklyReviewDaysList)

    $weeklyReviewAppsGroup = New-Object System.Windows.Forms.GroupBox
    $weeklyReviewAppsGroup.Text = "Top distractions this week"
    $weeklyReviewAppsGroup.Location = New-Object System.Drawing.Point(416, 190)
    $weeklyReviewAppsGroup.Size = New-Object System.Drawing.Size(392, 208)

    $weeklyReviewAppsList = New-Object System.Windows.Forms.ListView
    $weeklyReviewAppsList.Location = New-Object System.Drawing.Point(14, 28)
    $weeklyReviewAppsList.Size = New-Object System.Drawing.Size(362, 164)
    $weeklyReviewAppsList.View = [System.Windows.Forms.View]::Details
    $weeklyReviewAppsList.FullRowSelect = $true
    $weeklyReviewAppsList.GridLines = $true
    $weeklyReviewAppsList.HideSelection = $false
    $weeklyReviewAppsList.Tag = "string|string|time"
    [void]$weeklyReviewAppsList.Columns.Add("Process", 144)
    [void]$weeklyReviewAppsList.Columns.Add("Category", 118)
    [void]$weeklyReviewAppsList.Columns.Add("Time", 86)
    [void]$weeklyReviewAppsGroup.Controls.Add($weeklyReviewAppsList)

    $weeklyReviewHeatmapGroup = New-Object System.Windows.Forms.GroupBox
    $weeklyReviewHeatmapGroup.Text = "Calendar heatmap"
    $weeklyReviewHeatmapGroup.Location = New-Object System.Drawing.Point(12, 412)
    $weeklyReviewHeatmapGroup.Size = New-Object System.Drawing.Size(796, 204)

    $weeklyReviewHeatmapLegendLabel = New-Object System.Windows.Forms.Label
    $weeklyReviewHeatmapLegendLabel.Location = New-Object System.Drawing.Point(14, 26)
    $weeklyReviewHeatmapLegendLabel.Size = New-Object System.Drawing.Size(766, 18)
    $weeklyReviewHeatmapLegendLabel.Text = "Legend:"

    $weeklyReviewHeatmapSummaryLabel = New-Object System.Windows.Forms.Label
    $weeklyReviewHeatmapSummaryLabel.Location = New-Object System.Drawing.Point(14, 46)
    $weeklyReviewHeatmapSummaryLabel.Size = New-Object System.Drawing.Size(766, 18)
    $weeklyReviewHeatmapSummaryLabel.Text = "Heatmap:"

    $weekdayNames = @("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")
    for ($weekdayIndex = 0; $weekdayIndex -lt $weekdayNames.Count; $weekdayIndex += 1) {
        $weekdayLabel = New-Object System.Windows.Forms.Label
        $weekdayLabel.Location = New-Object System.Drawing.Point((14 + ($weekdayIndex * 108)), 72)
        $weekdayLabel.Size = New-Object System.Drawing.Size(100, 18)
        $weekdayLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $weekdayLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $weekdayLabel.Text = [string]$weekdayNames[$weekdayIndex]
        [void]$weeklyReviewHeatmapGroup.Controls.Add($weekdayLabel)
    }

    $weeklyReviewHeatmapLabels = @()
    $heatmapToolTip = New-Object System.Windows.Forms.ToolTip
    for ($row = 0; $row -lt 6; $row += 1) {
        for ($col = 0; $col -lt 7; $col += 1) {
            $dayLabel = New-Object System.Windows.Forms.Label
            $dayLabel.Location = New-Object System.Drawing.Point((14 + ($col * 108)), (96 + ($row * 16)))
            $dayLabel.Size = New-Object System.Drawing.Size(100, 14)
            $dayLabel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
            $dayLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
            $dayLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
            $dayLabel.BackColor = [System.Drawing.Color]::WhiteSmoke
            $dayLabel.ForeColor = [System.Drawing.Color]::Black
            $dayLabel.Text = ""
            $weeklyReviewHeatmapLabels += ,$dayLabel
            [void]$weeklyReviewHeatmapGroup.Controls.Add($dayLabel)
        }
    }

    foreach ($control in @($weeklyReviewHeatmapLegendLabel, $weeklyReviewHeatmapSummaryLabel)) {
        [void]$weeklyReviewHeatmapGroup.Controls.Add($control)
    }

    foreach ($control in @($weeklyReviewSummaryLabel, $weeklyReviewWinLabel, $weeklyReviewSlipLabel, $weeklyReviewDistractionLabel, $weeklyReviewCoachLabel, $weeklyReviewDaysGroup, $weeklyReviewAppsGroup, $weeklyReviewHeatmapGroup)) {
        [void]$weekReviewTab.Controls.Add($control)
    }

    $goalsIntroLabel = New-Object System.Windows.Forms.Label
    $goalsIntroLabel.Location = New-Object System.Drawing.Point(18, 18)
    $goalsIntroLabel.Size = New-Object System.Drawing.Size(700, 22)
    $goalsIntroLabel.Text = "See whether today is still on plan: total budget, study target, fun budget and socials."

    $goalsOverviewGroup = New-Object System.Windows.Forms.GroupBox
    $goalsOverviewGroup.Text = "Today on plan?"
    $goalsOverviewGroup.Location = New-Object System.Drawing.Point(12, 52)
    $goalsOverviewGroup.Size = New-Object System.Drawing.Size(796, 128)

    $goalsSummaryLabel = New-Object System.Windows.Forms.Label
    $goalsSummaryLabel.Location = New-Object System.Drawing.Point(16, 28)
    $goalsSummaryLabel.Size = New-Object System.Drawing.Size(760, 24)
    $goalsSummaryLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)

    $goalsPrimaryLabel = New-Object System.Windows.Forms.Label
    $goalsPrimaryLabel.Location = New-Object System.Drawing.Point(16, 58)
    $goalsPrimaryLabel.Size = New-Object System.Drawing.Size(760, 24)

    $goalsUnlockLabel = New-Object System.Windows.Forms.Label
    $goalsUnlockLabel.Location = New-Object System.Drawing.Point(16, 86)
    $goalsUnlockLabel.Size = New-Object System.Drawing.Size(760, 22)
    $goalsUnlockLabel.ForeColor = [System.Drawing.Color]::SaddleBrown

    foreach ($control in @($goalsSummaryLabel, $goalsPrimaryLabel, $goalsUnlockLabel)) {
        [void]$goalsOverviewGroup.Controls.Add($control)
    }

    $goalsRowsGroup = New-Object System.Windows.Forms.GroupBox
    $goalsRowsGroup.Text = "Goal status today"
    $goalsRowsGroup.Location = New-Object System.Drawing.Point(12, 194)
    $goalsRowsGroup.Size = New-Object System.Drawing.Size(796, 422)

    $goalsTotalRow = New-MetricRow -Title "Stay under 3 hours total" -Top 24
    $goalsStudyRow = New-MetricRow -Title "Study between 2:00 and 2:30" -Top 104
    $goalsBrowserFunRow = New-MetricRow -Title "Fun / manga up to 0:30" -Top 184
    $goalsSocialsRow = New-MetricRow -Title "Socials up to 0:18" -Top 264
    foreach ($control in ($goalsTotalRow.Controls + $goalsStudyRow.Controls + $goalsBrowserFunRow.Controls + $goalsSocialsRow.Controls)) {
        [void]$goalsRowsGroup.Controls.Add($control)
    }

    foreach ($control in @($goalsIntroLabel, $goalsOverviewGroup, $goalsRowsGroup)) {
        [void]$goalsTab.Controls.Add($control)
    }

    $analyticsIntroLabel = New-Object System.Windows.Forms.Label
    $analyticsIntroLabel.Location = New-Object System.Drawing.Point(18, 18)
    $analyticsIntroLabel.Size = New-Object System.Drawing.Size(470, 22)
    $analyticsIntroLabel.Text = "See the last 7 days at a glance, 30-day trends and export your report."

    $exportCsvButton = New-Object System.Windows.Forms.Button
    $exportCsvButton.Text = "Export CSV"
    $exportCsvButton.Location = New-Object System.Drawing.Point(580, 14)
    $exportCsvButton.Size = New-Object System.Drawing.Size(108, 32)

    $exportJsonButton = New-Object System.Windows.Forms.Button
    $exportJsonButton.Text = "Export JSON"
    $exportJsonButton.Location = New-Object System.Drawing.Point(700, 14)
    $exportJsonButton.Size = New-Object System.Drawing.Size(108, 32)

    $analyticsTotalLabel = New-SettingsValueLabel -Text "30-day total:" -Left 18 -Top 58 -Width 340
    $analyticsTotalLabel.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $analyticsAverageLabel = New-SettingsValueLabel -Text "Daily average:" -Left 18 -Top 92 -Width 340
    $analyticsBestDayLabel = New-SettingsValueLabel -Text "Best study day:" -Left 18 -Top 126 -Width 760
    $analyticsDistractionLabel = New-SettingsValueLabel -Text "Top distraction:" -Left 18 -Top 160 -Width 760
    $analyticsLimitStreakLabel = New-SettingsValueLabel -Text "Under-limit streak:" -Left 360 -Top 58 -Width 430
    $analyticsStudyStreakLabel = New-SettingsValueLabel -Text "Study-goal streak:" -Left 360 -Top 92 -Width 430

    $analyticsTrendGroup = New-Object System.Windows.Forms.GroupBox
    $analyticsTrendGroup.Text = "Last 7 days by category"
    $analyticsTrendGroup.Location = New-Object System.Drawing.Point(12, 194)
    $analyticsTrendGroup.Size = New-Object System.Drawing.Size(796, 250)

    $analyticsTrendChart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
    $analyticsTrendChart.Location = New-Object System.Drawing.Point(14, 26)
    $analyticsTrendChart.Size = New-Object System.Drawing.Size(766, 208)
    $analyticsTrendChart.BackColor = [System.Drawing.Color]::White
    $analyticsTrendArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea("Trend")
    $analyticsTrendArea.BackColor = [System.Drawing.Color]::White
    $analyticsTrendArea.AxisX.Interval = 1
    $analyticsTrendArea.AxisX.MajorGrid.Enabled = $false
    $analyticsTrendArea.AxisY.MajorGrid.LineColor = [System.Drawing.Color]::Gainsboro
    $analyticsTrendArea.AxisY.LabelStyle.Format = "0.0h"
    $analyticsTrendArea.AxisY.Title = "Hours"
    [void]$analyticsTrendChart.ChartAreas.Add($analyticsTrendArea)
    [void]$analyticsTrendGroup.Controls.Add($analyticsTrendChart)

    $analyticsBreakdownGroup = New-Object System.Windows.Forms.GroupBox
    $analyticsBreakdownGroup.Text = "30-day category breakdown"
    $analyticsBreakdownGroup.Location = New-Object System.Drawing.Point(12, 458)
    $analyticsBreakdownGroup.Size = New-Object System.Drawing.Size(320, 158)

    $analyticsBreakdownChart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
    $analyticsBreakdownChart.Location = New-Object System.Drawing.Point(12, 24)
    $analyticsBreakdownChart.Size = New-Object System.Drawing.Size(294, 120)
    $analyticsBreakdownChart.BackColor = [System.Drawing.Color]::White
    $analyticsBreakdownArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea("Breakdown")
    $analyticsBreakdownArea.BackColor = [System.Drawing.Color]::White
    $analyticsBreakdownArea.Area3DStyle.Enable3D = $false
    [void]$analyticsBreakdownChart.ChartAreas.Add($analyticsBreakdownArea)
    [void]$analyticsBreakdownGroup.Controls.Add($analyticsBreakdownChart)

    $analyticsAppsGroup = New-Object System.Windows.Forms.GroupBox
    $analyticsAppsGroup.Text = "Top apps in the last 30 days"
    $analyticsAppsGroup.Location = New-Object System.Drawing.Point(346, 458)
    $analyticsAppsGroup.Size = New-Object System.Drawing.Size(462, 158)

    $analyticsAppsList = New-Object System.Windows.Forms.ListView
    $analyticsAppsList.Location = New-Object System.Drawing.Point(14, 26)
    $analyticsAppsList.Size = New-Object System.Drawing.Size(432, 116)
    $analyticsAppsList.View = [System.Windows.Forms.View]::Details
    $analyticsAppsList.FullRowSelect = $true
    $analyticsAppsList.GridLines = $true
    $analyticsAppsList.HideSelection = $false
    $analyticsAppsList.Tag = "string|string|time|percent"
    [void]$analyticsAppsList.Columns.Add("Process", 148)
    [void]$analyticsAppsList.Columns.Add("Category", 108)
    [void]$analyticsAppsList.Columns.Add("Time", 92)
    [void]$analyticsAppsList.Columns.Add("Share", 70)
    [void]$analyticsAppsGroup.Controls.Add($analyticsAppsList)

    foreach ($control in @($analyticsIntroLabel, $exportCsvButton, $exportJsonButton, $analyticsTotalLabel, $analyticsAverageLabel, $analyticsBestDayLabel, $analyticsDistractionLabel, $analyticsLimitStreakLabel, $analyticsStudyStreakLabel, $analyticsTrendGroup, $analyticsBreakdownGroup, $analyticsAppsGroup)) {
        [void]$analyticsTab.Controls.Add($control)
    }

    $sessionsIntroLabel = New-Object System.Windows.Forms.Label
    $sessionsIntroLabel.Location = New-Object System.Drawing.Point(18, 16)
    $sessionsIntroLabel.Size = New-Object System.Drawing.Size(700, 22)
    $sessionsIntroLabel.Text = "See how the day unfolded: hour-by-hour timeline, recent sessions and longest blocks this week."

    $sessionCountLabel = New-SettingsValueLabel -Text "Sessions today:" -Left 18 -Top 50 -Width 240
    $sessionLongestLabel = New-SettingsValueLabel -Text "Longest today:" -Left 280 -Top 50 -Width 300
    $timelinePeakLabel = New-SettingsValueLabel -Text "Peak hour:" -Left 540 -Top 50 -Width 260

    $timelineChartGroup = New-Object System.Windows.Forms.GroupBox
    $timelineChartGroup.Text = "Today by hour"
    $timelineChartGroup.Location = New-Object System.Drawing.Point(12, 84)
    $timelineChartGroup.Size = New-Object System.Drawing.Size(796, 258)

    $timelineChart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
    $timelineChart.Location = New-Object System.Drawing.Point(14, 26)
    $timelineChart.Size = New-Object System.Drawing.Size(766, 216)
    $timelineChart.BackColor = [System.Drawing.Color]::White
    $timelineArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea("Timeline")
    $timelineArea.BackColor = [System.Drawing.Color]::White
    $timelineArea.AxisX.Interval = 2
    $timelineArea.AxisX.MajorGrid.Enabled = $false
    $timelineArea.AxisY.MajorGrid.LineColor = [System.Drawing.Color]::Gainsboro
    $timelineArea.AxisY.LabelStyle.Format = "0.0h"
    $timelineArea.AxisY.Title = "Hours"
    [void]$timelineChart.ChartAreas.Add($timelineArea)
    [void]$timelineChartGroup.Controls.Add($timelineChart)

    $todaySessionsGroup = New-Object System.Windows.Forms.GroupBox
    $todaySessionsGroup.Text = "Recent sessions today"
    $todaySessionsGroup.Location = New-Object System.Drawing.Point(12, 356)
    $todaySessionsGroup.Size = New-Object System.Drawing.Size(392, 260)

    $todaySessionsList = New-Object System.Windows.Forms.ListView
    $todaySessionsList.Location = New-Object System.Drawing.Point(14, 28)
    $todaySessionsList.Size = New-Object System.Drawing.Size(362, 216)
    $todaySessionsList.View = [System.Windows.Forms.View]::Details
    $todaySessionsList.FullRowSelect = $true
    $todaySessionsList.GridLines = $true
    $todaySessionsList.HideSelection = $false
    $todaySessionsList.Tag = "string|time|string|string|string"
    [void]$todaySessionsList.Columns.Add("End", 80)
    [void]$todaySessionsList.Columns.Add("Duration", 78)
    [void]$todaySessionsList.Columns.Add("Category", 94)
    [void]$todaySessionsList.Columns.Add("Process", 96)
    [void]$todaySessionsList.Columns.Add("Window title / domain", 214)
    [void]$todaySessionsGroup.Controls.Add($todaySessionsList)

    $weeklySessionsGroup = New-Object System.Windows.Forms.GroupBox
    $weeklySessionsGroup.Text = "Longest sessions in the last 7 days"
    $weeklySessionsGroup.Location = New-Object System.Drawing.Point(416, 356)
    $weeklySessionsGroup.Size = New-Object System.Drawing.Size(392, 260)

    $weeklySessionsList = New-Object System.Windows.Forms.ListView
    $weeklySessionsList.Location = New-Object System.Drawing.Point(14, 28)
    $weeklySessionsList.Size = New-Object System.Drawing.Size(362, 216)
    $weeklySessionsList.View = [System.Windows.Forms.View]::Details
    $weeklySessionsList.FullRowSelect = $true
    $weeklySessionsList.GridLines = $true
    $weeklySessionsList.HideSelection = $false
    $weeklySessionsList.Tag = "date|time|string|string|string"
    [void]$weeklySessionsList.Columns.Add("Date", 84)
    [void]$weeklySessionsList.Columns.Add("Duration", 78)
    [void]$weeklySessionsList.Columns.Add("Category", 94)
    [void]$weeklySessionsList.Columns.Add("Process", 96)
    [void]$weeklySessionsList.Columns.Add("Window title / domain", 206)
    [void]$weeklySessionsGroup.Controls.Add($weeklySessionsList)

    foreach ($control in @($sessionsIntroLabel, $sessionCountLabel, $sessionLongestLabel, $timelinePeakLabel, $timelineChartGroup, $todaySessionsGroup, $weeklySessionsGroup)) {
        [void]$sessionsTab.Controls.Add($control)
    }

    $reviewIntroLabel = New-Object System.Windows.Forms.Label
    $reviewIntroLabel.Location = New-Object System.Drawing.Point(18, 16)
    $reviewIntroLabel.Size = New-Object System.Drawing.Size(700, 22)
    $reviewIntroLabel.Text = "These are activities from today that did not match any rule yet. Classify them once and the tracker will learn."

    $reviewSummaryLabel = New-Object System.Windows.Forms.Label
    $reviewSummaryLabel.Location = New-Object System.Drawing.Point(18, 46)
    $reviewSummaryLabel.Size = New-Object System.Drawing.Size(520, 24)
    $reviewSummaryLabel.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)

    $reviewClassifyButton = New-Object System.Windows.Forms.Button
    $reviewClassifyButton.Text = "Classify selected"
    $reviewClassifyButton.Location = New-Object System.Drawing.Point(656, 42)
    $reviewClassifyButton.Size = New-Object System.Drawing.Size(152, 32)
    $reviewClassifyButton.Enabled = $false

    $reviewGroup = New-Object System.Windows.Forms.GroupBox
    $reviewGroup.Text = "Uncategorized activities today"
    $reviewGroup.Location = New-Object System.Drawing.Point(12, 84)
    $reviewGroup.Size = New-Object System.Drawing.Size(796, 532)

    $reviewList = New-Object System.Windows.Forms.ListView
    $reviewList.Location = New-Object System.Drawing.Point(14, 28)
    $reviewList.Size = New-Object System.Drawing.Size(766, 488)
    $reviewList.View = [System.Windows.Forms.View]::Details
    $reviewList.FullRowSelect = $true
    $reviewList.GridLines = $true
    $reviewList.HideSelection = $false
    $reviewList.Tag = "string|string|time"
    [void]$reviewList.Columns.Add("Process", 160)
    [void]$reviewList.Columns.Add("Window title / domain", 486)
    [void]$reviewList.Columns.Add("Time", 100)
    [void]$reviewGroup.Controls.Add($reviewList)

    foreach ($control in @($reviewIntroLabel, $reviewSummaryLabel, $reviewClassifyButton, $reviewGroup)) {
        [void]$reviewTab.Controls.Add($control)
    }

    $healthIntroLabel = New-Object System.Windows.Forms.Label
    $healthIntroLabel.Location = New-Object System.Drawing.Point(18, 16)
    $healthIntroLabel.Size = New-Object System.Drawing.Size(760, 22)
    $healthIntroLabel.Text = "See how well the tracker understands your activity: coverage, missing rules and the rules doing most of the work."

    $healthTodayCoverageLabel = New-SettingsValueLabel -Text "Today classified:" -Left 18 -Top 48 -Width 760
    $healthTodayCoverageLabel.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $healthWeekCoverageLabel = New-SettingsValueLabel -Text "7-day classified:" -Left 18 -Top 80 -Width 760
    $healthReviewSummaryLabel = New-SettingsValueLabel -Text "Uncategorized today:" -Left 18 -Top 112 -Width 760
    $healthMainGapLabel = New-SettingsValueLabel -Text "Main gap today:" -Left 18 -Top 144 -Width 760

    $healthReviewButton = New-Object System.Windows.Forms.Button
    $healthReviewButton.Text = "Open review"
    $healthReviewButton.Location = New-Object System.Drawing.Point(594, 40)
    $healthReviewButton.Size = New-Object System.Drawing.Size(102, 32)

    $healthClassifyButton = New-Object System.Windows.Forms.Button
    $healthClassifyButton.Text = "Classify current"
    $healthClassifyButton.Location = New-Object System.Drawing.Point(706, 40)
    $healthClassifyButton.Size = New-Object System.Drawing.Size(102, 32)

    $healthUncategorizedGroup = New-Object System.Windows.Forms.GroupBox
    $healthUncategorizedGroup.Text = "Top uncategorized today"
    $healthUncategorizedGroup.Location = New-Object System.Drawing.Point(12, 186)
    $healthUncategorizedGroup.Size = New-Object System.Drawing.Size(392, 206)

    $healthUncategorizedList = New-Object System.Windows.Forms.ListView
    $healthUncategorizedList.Location = New-Object System.Drawing.Point(14, 28)
    $healthUncategorizedList.Size = New-Object System.Drawing.Size(362, 162)
    $healthUncategorizedList.View = [System.Windows.Forms.View]::Details
    $healthUncategorizedList.FullRowSelect = $true
    $healthUncategorizedList.GridLines = $true
    $healthUncategorizedList.HideSelection = $false
    $healthUncategorizedList.Tag = "string|time|number|string"
    [void]$healthUncategorizedList.Columns.Add("Process", 102)
    [void]$healthUncategorizedList.Columns.Add("Time", 78)
    [void]$healthUncategorizedList.Columns.Add("Hits", 54)
    [void]$healthUncategorizedList.Columns.Add("Example", 120)
    [void]$healthUncategorizedGroup.Controls.Add($healthUncategorizedList)

    $healthRulesGroup = New-Object System.Windows.Forms.GroupBox
    $healthRulesGroup.Text = "Rules doing the most work this week"
    $healthRulesGroup.Location = New-Object System.Drawing.Point(416, 186)
    $healthRulesGroup.Size = New-Object System.Drawing.Size(392, 206)

    $healthRulesList = New-Object System.Windows.Forms.ListView
    $healthRulesList.Location = New-Object System.Drawing.Point(14, 28)
    $healthRulesList.Size = New-Object System.Drawing.Size(362, 162)
    $healthRulesList.View = [System.Windows.Forms.View]::Details
    $healthRulesList.FullRowSelect = $true
    $healthRulesList.GridLines = $true
    $healthRulesList.HideSelection = $false
    $healthRulesList.Tag = "string|string|time|number"
    [void]$healthRulesList.Columns.Add("Rule", 126)
    [void]$healthRulesList.Columns.Add("Category", 82)
    [void]$healthRulesList.Columns.Add("Time", 78)
    [void]$healthRulesList.Columns.Add("Hits", 54)
    [void]$healthRulesGroup.Controls.Add($healthRulesList)

    $healthSuggestionsGroup = New-Object System.Windows.Forms.GroupBox
    $healthSuggestionsGroup.Text = "Suggested rules today"
    $healthSuggestionsGroup.Location = New-Object System.Drawing.Point(12, 406)
    $healthSuggestionsGroup.Size = New-Object System.Drawing.Size(796, 210)

    $healthSuggestionsSummaryLabel = New-Object System.Windows.Forms.Label
    $healthSuggestionsSummaryLabel.Location = New-Object System.Drawing.Point(14, 28)
    $healthSuggestionsSummaryLabel.Size = New-Object System.Drawing.Size(620, 22)
    $healthSuggestionsSummaryLabel.Text = "Suggested rules: no strong automatic suggestion yet."

    $healthUseSuggestionButton = New-Object System.Windows.Forms.Button
    $healthUseSuggestionButton.Text = "Use suggestion"
    $healthUseSuggestionButton.Location = New-Object System.Drawing.Point(666, 22)
    $healthUseSuggestionButton.Size = New-Object System.Drawing.Size(114, 30)
    $healthUseSuggestionButton.Enabled = $false

    $healthSuggestionsList = New-Object System.Windows.Forms.ListView
    $healthSuggestionsList.Location = New-Object System.Drawing.Point(14, 60)
    $healthSuggestionsList.Size = New-Object System.Drawing.Size(766, 134)
    $healthSuggestionsList.View = [System.Windows.Forms.View]::Details
    $healthSuggestionsList.FullRowSelect = $true
    $healthSuggestionsList.GridLines = $true
    $healthSuggestionsList.HideSelection = $false
    $healthSuggestionsList.Tag = "string|string|time|number|string|string"
    [void]$healthSuggestionsList.Columns.Add("Target", 78)
    [void]$healthSuggestionsList.Columns.Add("Pattern", 150)
    [void]$healthSuggestionsList.Columns.Add("Time", 82)
    [void]$healthSuggestionsList.Columns.Add("Hits", 54)
    [void]$healthSuggestionsList.Columns.Add("Why", 110)
    [void]$healthSuggestionsList.Columns.Add("Example", 270)
    foreach ($control in @($healthSuggestionsSummaryLabel, $healthUseSuggestionButton, $healthSuggestionsList)) {
        [void]$healthSuggestionsGroup.Controls.Add($control)
    }

    foreach ($control in @($healthIntroLabel, $healthTodayCoverageLabel, $healthWeekCoverageLabel, $healthReviewSummaryLabel, $healthMainGapLabel, $healthReviewButton, $healthClassifyButton, $healthUncategorizedGroup, $healthRulesGroup, $healthSuggestionsGroup)) {
        [void]$healthTab.Controls.Add($control)
    }

    $rulesIntroLabel = New-Object System.Windows.Forms.Label
    $rulesIntroLabel.Location = New-Object System.Drawing.Point(18, 16)
    $rulesIntroLabel.Size = New-Object System.Drawing.Size(760, 22)
    $rulesIntroLabel.Text = "The lowest priority rule wins first. Use match mode and enable/disable rules instead of deleting them."

    $rulesGroup = New-Object System.Windows.Forms.GroupBox
    $rulesGroup.Text = "Loaded rules"
    $rulesGroup.Location = New-Object System.Drawing.Point(12, 44)
    $rulesGroup.Size = New-Object System.Drawing.Size(796, 572)

    $ruleList = New-Object System.Windows.Forms.ListView
    $ruleList.Location = New-Object System.Drawing.Point(14, 28)
    $ruleList.Size = New-Object System.Drawing.Size(766, 528)
    $ruleList.View = [System.Windows.Forms.View]::Details
    $ruleList.FullRowSelect = $true
    $ruleList.GridLines = $true
    $ruleList.HideSelection = $false
    [void]$ruleList.Columns.Add("Rule", 130)
    [void]$ruleList.Columns.Add("Category", 110)
    [void]$ruleList.Columns.Add("Target", 70)
    [void]$ruleList.Columns.Add("Match", 70)
    [void]$ruleList.Columns.Add("Priority", 70)
    [void]$ruleList.Columns.Add("Enabled", 70)
    [void]$ruleList.Columns.Add("Patterns", 220)
    [void]$rulesGroup.Controls.Add($ruleList)

    [void]$rulesTab.Controls.Add($rulesIntroLabel)
    [void]$rulesTab.Controls.Add($rulesGroup)

    $autoStartLabel = New-Object System.Windows.Forms.Label
    $autoStartLabel.Location = New-Object System.Drawing.Point(20, 768)
    $autoStartLabel.Size = New-Object System.Drawing.Size(220, 24)
    $autoStartLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)

    $autoStartButton = New-Object System.Windows.Forms.Button
    $autoStartButton.Location = New-Object System.Drawing.Point(250, 764)
    $autoStartButton.Size = New-Object System.Drawing.Size(160, 34)

    $focusButton = New-Object System.Windows.Forms.Button
    $focusButton.Text = "Start focus"
    $focusButton.Location = New-Object System.Drawing.Point(250, 740)
    $focusButton.Size = New-Object System.Drawing.Size(160, 24)

    $settingsButton = New-Object System.Windows.Forms.Button
    $settingsButton.Text = "Settings"
    $settingsButton.Location = New-Object System.Drawing.Point(20, 798)
    $settingsButton.Size = New-Object System.Drawing.Size(110, 34)

    $pauseButton = New-Object System.Windows.Forms.Button
    $pauseButton.Text = "Pause tracking"
    $pauseButton.Location = New-Object System.Drawing.Point(144, 798)
    $pauseButton.Size = New-Object System.Drawing.Size(130, 34)

    $resetButton = New-Object System.Windows.Forms.Button
    $resetButton.Text = "Reset today"
    $resetButton.Location = New-Object System.Drawing.Point(288, 798)
    $resetButton.Size = New-Object System.Drawing.Size(110, 34)

    $editRulesButton = New-Object System.Windows.Forms.Button
    $editRulesButton.Text = "Rules"
    $editRulesButton.Location = New-Object System.Drawing.Point(412, 798)
    $editRulesButton.Size = New-Object System.Drawing.Size(80, 34)

    $categoriesButton = New-Object System.Windows.Forms.Button
    $categoriesButton.Text = "Categories"
    $categoriesButton.Location = New-Object System.Drawing.Point(506, 798)
    $categoriesButton.Size = New-Object System.Drawing.Size(100, 34)

    $classifyButton = New-Object System.Windows.Forms.Button
    $classifyButton.Text = "Classify current"
    $classifyButton.Location = New-Object System.Drawing.Point(620, 798)
    $classifyButton.Size = New-Object System.Drawing.Size(118, 34)

    $reloadRulesButton = New-Object System.Windows.Forms.Button
    $reloadRulesButton.Text = "Reload rules"
    $reloadRulesButton.Location = New-Object System.Drawing.Point(752, 798)
    $reloadRulesButton.Size = New-Object System.Drawing.Size(104, 34)

    $openDataButton = New-Object System.Windows.Forms.Button
    $openDataButton.Text = "Open data folder"
    $openDataButton.Location = New-Object System.Drawing.Point(870, 798)
    $openDataButton.Size = New-Object System.Drawing.Size(120, 34)

    $openExtensionButton = New-Object System.Windows.Forms.Button
    $openExtensionButton.Text = "Browser extension"
    $openExtensionButton.Location = New-Object System.Drawing.Point(1004, 798)
    $openExtensionButton.Size = New-Object System.Drawing.Size(126, 34)

    $hintLabel = New-Object System.Windows.Forms.Label
    $hintLabel.Location = New-Object System.Drawing.Point(1140, 804)
    $hintLabel.Size = New-Object System.Drawing.Size(130, 24)
    $hintLabel.Text = "Tip: app stays in tray."

    foreach ($control in @($headerLabel, $modeLabel, $browserBridgeLabel, $summaryGroup, $otherLabel, $topAppLabel, $activeGroup, $focusModeLabel, $focusButton, $mainTabs, $autoStartLabel, $autoStartButton, $settingsButton, $pauseButton, $resetButton, $editRulesButton, $categoriesButton, $classifyButton, $reloadRulesButton, $openDataButton, $openExtensionButton, $hintLabel)) {
        [void]$form.Controls.Add($control)
    }

    $script:NotifyIcon = New-Object System.Windows.Forms.NotifyIcon
    $script:NotifyIcon.Icon = Get-AppIcon
    $script:NotifyIcon.Text = $script:WindowTitle
    $script:NotifyIcon.Visible = $true

    $trayMenu = New-Object System.Windows.Forms.ContextMenuStrip
    $statusItem = $trayMenu.Items.Add("Left today")
    $statusItem.Enabled = $false
    $currentTrayItem = $trayMenu.Items.Add("Current: idle / no app")
    $currentTrayItem.Enabled = $false
    $topAppTrayItem = $trayMenu.Items.Add("Most used app today")
    $topAppTrayItem.Enabled = $false
    [void]$trayMenu.Items.Add("-")
    $reviewTrayItem = $trayMenu.Items.Add("Review uncategorized")
    $pauseTrayItem = $trayMenu.Items.Add("Pause tracking")
    $focusTrayItem = $trayMenu.Items.Add("Start focus mode")
    $classifyTrayItem = $trayMenu.Items.Add("Classify current")
    $quickGlanceTrayItem = $trayMenu.Items.Add("Show quick glance")
    $openTrayItem = $trayMenu.Items.Add("Open dashboard")
    $analyticsTrayItem = $trayMenu.Items.Add("Open analytics")
    $settingsTrayItem = $trayMenu.Items.Add("Settings")
    [void]$trayMenu.Items.Add("-")
    $exitTrayItem = $trayMenu.Items.Add("Exit")
    $script:NotifyIcon.ContextMenuStrip = $trayMenu

    $state.Controls = @{
        HeaderLabel = $headerLabel
        ModeLabel = $modeLabel
        TotalRow = $totalRow
        StudyRow = $studyRow
        BrowserFunRow = $browserFunRow
        SocialsRow = $socialsRow
        OtherLabel = $otherLabel
        TopAppLabel = $topAppLabel
        ActiveProcessValue = $activeProcessValue
        ActiveTitleValue = $activeTitleValue
        ActiveUrlValue = $activeUrlValue
        ActiveCategoryValue = $activeCategoryValue
        IdleValue = $idleValue
        FocusModeLabel = $focusModeLabel
        FocusButton = $focusButton
        MainTabs = $mainTabs
        ActivityList = $activityList
        AppList = $appList
        ExactCategoryList = $exactCategoriesList
        HistoryList = $historyList
        InsightTotalLabel = $insightTotalLabel
        InsightStudyLabel = $insightStudyLabel
        InsightAverageLabel = $insightAverageLabel
        InsightStudyDaysLabel = $insightStudyDaysLabel
        InsightBestDayLabel = $insightBestDayLabel
        InsightTopAppLabel = $insightTopAppLabel
        InsightAppsList = $insightAppsList
        InsightCategoriesList = $insightCategoriesList
        WeeklyReviewSummaryLabel = $weeklyReviewSummaryLabel
        WeeklyReviewWinLabel = $weeklyReviewWinLabel
        WeeklyReviewSlipLabel = $weeklyReviewSlipLabel
        WeeklyReviewDistractionLabel = $weeklyReviewDistractionLabel
        WeeklyReviewCoachLabel = $weeklyReviewCoachLabel
        WeeklyReviewDaysList = $weeklyReviewDaysList
        WeeklyReviewAppsList = $weeklyReviewAppsList
        WeeklyReviewHeatmapLegendLabel = $weeklyReviewHeatmapLegendLabel
        WeeklyReviewHeatmapSummaryLabel = $weeklyReviewHeatmapSummaryLabel
        WeeklyReviewHeatmapLabels = $weeklyReviewHeatmapLabels
        WeeklyReviewHeatmapToolTip = $heatmapToolTip
        GoalsSummaryLabel = $goalsSummaryLabel
        GoalsPrimaryLabel = $goalsPrimaryLabel
        GoalsUnlockLabel = $goalsUnlockLabel
        GoalsTotalRow = $goalsTotalRow
        GoalsStudyRow = $goalsStudyRow
        GoalsBrowserFunRow = $goalsBrowserFunRow
        GoalsSocialsRow = $goalsSocialsRow
        AnalyticsTotalLabel = $analyticsTotalLabel
        AnalyticsAverageLabel = $analyticsAverageLabel
        AnalyticsBestDayLabel = $analyticsBestDayLabel
        AnalyticsDistractionLabel = $analyticsDistractionLabel
        AnalyticsLimitStreakLabel = $analyticsLimitStreakLabel
        AnalyticsStudyStreakLabel = $analyticsStudyStreakLabel
        AnalyticsTrendChart = $analyticsTrendChart
        AnalyticsBreakdownChart = $analyticsBreakdownChart
        AnalyticsAppsList = $analyticsAppsList
        SessionCountLabel = $sessionCountLabel
        SessionLongestLabel = $sessionLongestLabel
        TimelinePeakLabel = $timelinePeakLabel
        TimelineChart = $timelineChart
        TodaySessionsList = $todaySessionsList
        WeeklySessionsList = $weeklySessionsList
        ReviewSummaryLabel = $reviewSummaryLabel
        ReviewList = $reviewList
        ReviewClassifyButton = $reviewClassifyButton
        HealthTodayCoverageLabel = $healthTodayCoverageLabel
        HealthWeekCoverageLabel = $healthWeekCoverageLabel
        HealthReviewSummaryLabel = $healthReviewSummaryLabel
        HealthMainGapLabel = $healthMainGapLabel
        HealthReviewButton = $healthReviewButton
        HealthClassifyButton = $healthClassifyButton
        HealthUncategorizedList = $healthUncategorizedList
        HealthRulesList = $healthRulesList
        HealthSuggestionsSummaryLabel = $healthSuggestionsSummaryLabel
        HealthSuggestionsList = $healthSuggestionsList
        HealthUseSuggestionButton = $healthUseSuggestionButton
        RuleList = $ruleList
        AutoStartLabel = $autoStartLabel
        AutoStartButton = $autoStartButton
        SettingsButton = $settingsButton
        CategoriesButton = $categoriesButton
        BrowserBridgeLabel = $browserBridgeLabel
        PauseButton = $pauseButton
        ExportCsvButton = $exportCsvButton
        ExportJsonButton = $exportJsonButton
        QuickGlanceForm = $quickGlanceForm
        QuickGlanceDateLabel = $quickDateLabel
        QuickGlanceTotalLabel = $quickTotalLabel
        QuickGlanceStudyLabel = $quickStudyLabel
        QuickGlanceFunLabel = $quickFunLabel
        QuickGlanceSocialsLabel = $quickSocialsLabel
        QuickGlanceCurrentLabel = $quickCurrentLabel
        QuickGlanceFocusLabel = $quickFocusLabel
        QuickGlanceStatusLabel = $quickStatusLabel
        QuickGlanceOpenButton = $quickOpenButton
        QuickGlancePauseButton = $quickPauseButton
    }
    $state.TrayItems = @{
        StatusItem = $statusItem
        CurrentItem = $currentTrayItem
        TopAppItem = $topAppTrayItem
        ReviewItem = $reviewTrayItem
        PauseItem = $pauseTrayItem
        FocusItem = $focusTrayItem
        ClassifyItem = $classifyTrayItem
        QuickGlanceItem = $quickGlanceTrayItem
        OpenItem = $openTrayItem
        AnalyticsItem = $analyticsTrayItem
        SettingsItem = $settingsTrayItem
        ExitItem = $exitTrayItem
    }

    Update-RuleList -RuleList $ruleList -Rules $state.Rules
    Update-Ui -State $state
    [void](Consume-ReopenRunningInstanceRequest)

    $startupTimer = New-Object System.Windows.Forms.Timer
    $startupTimer.Interval = 250
    $startupTimer.Add_Tick({
        $startupTimer.Stop()
        $startupTimer.Dispose()
        try {
            Ensure-BackgroundServicesStarted -State $state
            Update-BrowserBridgeUi -State $state
        }
        catch {
        }
    })
    $startupTimer.Start()

    $sampleTimer = New-Object System.Windows.Forms.Timer
    $sampleTimer.Interval = [int]([math]::Max(1, [double]$state.Settings.sampleIntervalSeconds) * 1000)
    $sampleTimer.Add_Tick({
        try {
            if (Consume-ReopenRunningInstanceRequest) {
                Show-TrackerWindow -Form $form
                Request-UiRefresh -State $state
            }

            $now = Get-Date
            $elapsed = ($now - $state.LastTickAt).TotalSeconds
            $state.LastTickAt = $now

            $newDateKey = Get-DateKey
            if ($newDateKey -ne $state.CurrentDateKey) {
                Stop-TrackedSession -State $state
                $state.CurrentDateKey = $newDateKey
                [void](Get-DayStats -UsageData $state.UsageData -DateKey $newDateKey)
                [void](Ensure-NotificationState -State $state -DateKey $newDateKey)
                $state.OverrideUntil = $null
                $state.HardLimitHandledFor = @{}
                Clear-DerivedCaches -State $state
                $state.IsDirty = $true
            }

            $activity = Get-EffectiveActivity -Activity (Get-ActiveWindowInfo) -State $state
            $idleSeconds = Get-IdleSeconds
            $state.CurrentIdleSeconds = $idleSeconds
            $state.IsIdle = $idleSeconds -ge [double]$state.Settings.idleThresholdSeconds

            if ($state.TrackingEnabled -and -not $state.IsIdle -and -not (Should-IgnoreActivity $activity)) {
                $category = Get-CategoryForActivity -Activity $activity -Rules $state.Rules
                $state.LastActivity = $activity
                $state.LastActivityCategory = $category

                if (Handle-FocusMode -State $state -Activity $activity -Category $category) {
                    Add-UsageSample -UsageData $state.UsageData -DateKey $state.CurrentDateKey -Activity $activity -Category $category -Seconds $elapsed -State $state
                    Update-TrackedSession -State $state -DateKey $state.CurrentDateKey -Activity $activity -Category $category -Seconds $elapsed
                    $state.IsDirty = $true
                }
                else {
                    Stop-TrackedSession -State $state
                }
            }
            elseif ($null -ne $activity) {
                Stop-TrackedSession -State $state
                $state.LastActivity = $activity
                $state.LastActivityCategory = Get-CategoryForActivity -Activity $activity -Rules $state.Rules
            }
            else {
                Stop-TrackedSession -State $state
            }

            Check-Limits -State $state
            Save-IfDirty -State $state
            Update-Ui -State $state
        }
        catch {
            $state.LastTickAt = Get-Date
            $state.CurrentIdleSeconds = 0
            $state.IsIdle = $false
            if ($null -ne $script:NotifyIcon) {
                $script:NotifyIcon.BalloonTipTitle = "Tracking error"
                $script:NotifyIcon.BalloonTipText = $_.Exception.Message
                $script:NotifyIcon.ShowBalloonTip(4000)
            }
        }
    })
    $sampleTimer.Start()

    $pauseButton.Add_Click({
        $state.TrackingEnabled = -not $state.TrackingEnabled
        if ($state.TrackingEnabled) {
            $state.HardLimitHandledFor = @{}
        }
        Update-Ui -State $state
    })

    $resetButton.Add_Click({
        $answer = [System.Windows.Forms.MessageBox]::Show("Reset all tracked time for today?", "Reset today", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) {
            Reset-Today -State $state
        }
    })

    $settingsButton.Add_Click({
        if (Show-SettingsEditor -State $state) {
            $sampleTimer.Interval = [int]([math]::Max(1, [double]$state.Settings.sampleIntervalSeconds) * 1000)
            Request-UiRefresh -State $state
            Update-Ui -State $state
        }
    })
    $focusButton.Add_Click({
        if ($state.FocusMode.Enabled) {
            Stop-FocusMode -State $state
        }
        else {
            [void](Show-FocusModeDialog -State $state)
        }
        Request-UiRefresh -State $state
        Update-Ui -State $state
    })
    $editRulesButton.Add_Click({
        if (Show-RulesEditor -State $state) {
            Update-RuleList -RuleList $ruleList -Rules $state.Rules
            Clear-DerivedCaches -State $state
            Request-UiRefresh -State $state
            Update-Ui -State $state
        }
    })
    $categoriesButton.Add_Click({
        if (Show-CategoriesEditor -State $state) {
            Update-RuleList -RuleList $ruleList -Rules $state.Rules
            Clear-DerivedCaches -State $state
            Request-UiRefresh -State $state
            Update-Ui -State $state
        }
    })
    $classifyButton.Add_Click({
        if (Show-ClassifyCurrentActivityDialog -State $state) {
            Update-RuleList -RuleList $ruleList -Rules $state.Rules
            Clear-DerivedCaches -State $state
            Request-UiRefresh -State $state
            Update-Ui -State $state
        }
    })
    $reloadRulesButton.Add_Click({ $state.Rules = Load-Rules; Update-RuleList -RuleList $ruleList -Rules $state.Rules; Clear-DerivedCaches -State $state; Request-UiRefresh -State $state; Update-Ui -State $state })
    $openDataButton.Add_Click({ Start-Process -FilePath "explorer.exe" -ArgumentList @($script:DataDirectory) })
    $openExtensionButton.Add_Click({ Start-Process -FilePath "explorer.exe" -ArgumentList @($script:BrowserExtensionDirectory) })
    $quickOpenButton.Add_Click({
        Show-TrackerWindow -Form $form
        Request-UiRefresh -State $state
        Update-Ui -State $state
    })
    $quickPauseButton.Add_Click({
        $state.TrackingEnabled = -not $state.TrackingEnabled
        if ($state.TrackingEnabled) {
            $state.HardLimitHandledFor = @{}
        }
        Request-UiRefresh -State $state
        Update-Ui -State $state
    })
    $exportCsvButton.Add_Click({
        $path = Show-SaveFileDialog -Title "Export analytics report" -Filter "CSV files (*.csv)|*.csv" -DefaultFileName ("screen-time-report-{0}.csv" -f (Get-Date -Format "yyyy-MM-dd"))
        if (-not [string]::IsNullOrWhiteSpace($path)) {
            Export-AnalyticsCsv -UsageData $state.UsageData -Days 30 -Path $path
            Show-TrackerNotification -Title "CSV export ready" -Message "Analytics report saved."
        }
    })
    $exportJsonButton.Add_Click({
        $path = Show-SaveFileDialog -Title "Export raw usage data" -Filter "JSON files (*.json)|*.json" -DefaultFileName ("screen-time-raw-{0}.json" -f (Get-Date -Format "yyyy-MM-dd"))
        if (-not [string]::IsNullOrWhiteSpace($path)) {
            Export-UsageDataJson -UsageData $state.UsageData -Path $path
            Show-TrackerNotification -Title "JSON export ready" -Message "Raw usage data saved."
        }
    })
    $autoStartButton.Add_Click({
        $desiredEnabled = -not (Get-DesiredAutoStartEnabled -Settings $state.Settings)
        if (-not $state.Settings.ContainsKey("autostart")) {
            $state.Settings["autostart"] = @{ enabled = $desiredEnabled }
        }
        else {
            $state.Settings.autostart["enabled"] = $desiredEnabled
        }

        Save-Settings -Settings $state.Settings
        Sync-AutoStartWithSettings -Settings $state.Settings
        Update-AutoStartUi -State $state
    })
    $reviewClassifyButton.Add_Click({
        if ($reviewList.SelectedItems.Count -eq 0) {
            return
        }

        $selectedActivity = $reviewList.SelectedItems[0].Tag
        if ($null -eq $selectedActivity) {
            return
        }

        if (Show-ClassifyCurrentActivityDialog -State $state -ActivityOverride $selectedActivity -DefaultCategory "other" -DialogTitle "Classify reviewed activity") {
            Update-RuleList -RuleList $ruleList -Rules $state.Rules
            Clear-DerivedCaches -State $state
            Request-UiRefresh -State $state
            Update-Ui -State $state
        }
    })
    $healthReviewButton.Add_Click({
        Open-TrackerSection -State $state -Form $form -TabKey "review"
    })
    $healthClassifyButton.Add_Click({
        if (Show-ClassifyCurrentActivityDialog -State $state) {
            Update-RuleList -RuleList $ruleList -Rules $state.Rules
            Clear-DerivedCaches -State $state
            Request-UiRefresh -State $state
            Update-Ui -State $state
        }
    })
    $healthUseSuggestionButton.Add_Click({
        if ($healthSuggestionsList.SelectedItems.Count -eq 0) {
            return
        }

        $suggestion = $healthSuggestionsList.SelectedItems[0].Tag
        if ($null -eq $suggestion -or $null -eq $suggestion.activity) {
            return
        }

        if (Show-ClassifyCurrentActivityDialog -State $state -ActivityOverride $suggestion.activity -DefaultCategory "other" -DialogTitle "Use suggested rule" -SuggestedTarget ([string]$suggestion.target) -SuggestedPattern ([string]$suggestion.pattern) -SuggestedMatchMode ([string]$suggestion.matchMode)) {
            Update-RuleList -RuleList $ruleList -Rules $state.Rules
            Clear-DerivedCaches -State $state
            Request-UiRefresh -State $state
            Update-Ui -State $state
        }
    })

    $activityList.Add_ColumnClick({
        param($sender, $eventArgs)
        $sort = $state.ListSorts.Activity
        if ($sort.Column -eq $eventArgs.Column) {
            $sort.Order = Get-ToggledSortOrder -CurrentOrder $sort.Order
        }
        else {
            $sort.Column = $eventArgs.Column
            $sort.Order = [System.Windows.Forms.SortOrder]::Ascending
        }
        Update-ActivityList -ActivityList $activityList -State $state
    })

    $appList.Add_ColumnClick({
        param($sender, $eventArgs)
        $sort = $state.ListSorts.Apps
        if ($sort.Column -eq $eventArgs.Column) {
            $sort.Order = Get-ToggledSortOrder -CurrentOrder $sort.Order
        }
        else {
            $sort.Column = $eventArgs.Column
            $sort.Order = [System.Windows.Forms.SortOrder]::Ascending
        }
        Update-AppUsageList -AppList $appList -State $state
    })

    $exactCategoriesList.Add_ColumnClick({
        param($sender, $eventArgs)
        $sort = $state.ListSorts.ExactCategories
        if ($sort.Column -eq $eventArgs.Column) {
            $sort.Order = Get-ToggledSortOrder -CurrentOrder $sort.Order
        }
        else {
            $sort.Column = $eventArgs.Column
            $sort.Order = [System.Windows.Forms.SortOrder]::Ascending
        }
        Update-ExactCategoryList -CategoryList $exactCategoriesList -Items (Get-TopExactCategories -UsageData $state.UsageData -DateKey $state.CurrentDateKey -Top 12) -SortState $sort
    })

    $historyList.Add_ColumnClick({
        param($sender, $eventArgs)
        $sort = $state.ListSorts.History
        if ($sort.Column -eq $eventArgs.Column) {
            $sort.Order = Get-ToggledSortOrder -CurrentOrder $sort.Order
        }
        else {
            $sort.Column = $eventArgs.Column
            $sort.Order = [System.Windows.Forms.SortOrder]::Ascending
        }
        Update-HistoryList -HistoryList $historyList -State $state
    })

    $analyticsAppsList.Add_ColumnClick({
        param($sender, $eventArgs)
        $sort = $state.ListSorts.AnalyticsApps
        if ($sort.Column -eq $eventArgs.Column) {
            $sort.Order = Get-ToggledSortOrder -CurrentOrder $sort.Order
        }
        else {
            $sort.Column = $eventArgs.Column
            $sort.Order = [System.Windows.Forms.SortOrder]::Ascending
        }
        Update-AnalyticsUi -State $state
    })

    $insightCategoriesList.Add_ColumnClick({
        param($sender, $eventArgs)
        $sort = $state.ListSorts.InsightCategories
        if ($sort.Column -eq $eventArgs.Column) {
            $sort.Order = Get-ToggledSortOrder -CurrentOrder $sort.Order
        }
        else {
            $sort.Column = $eventArgs.Column
            $sort.Order = [System.Windows.Forms.SortOrder]::Ascending
        }
        Update-InsightsUi -State $state
    })

    $weeklyReviewDaysList.Add_ColumnClick({
        param($sender, $eventArgs)
        $sort = $state.ListSorts.WeeklyReviewDays
        if ($sort.Column -eq $eventArgs.Column) {
            $sort.Order = Get-ToggledSortOrder -CurrentOrder $sort.Order
        }
        else {
            $sort.Column = $eventArgs.Column
            $sort.Order = [System.Windows.Forms.SortOrder]::Ascending
        }
        Update-WeeklyReviewUi -State $state
    })

    $weeklyReviewAppsList.Add_ColumnClick({
        param($sender, $eventArgs)
        $sort = $state.ListSorts.WeeklyReviewApps
        if ($sort.Column -eq $eventArgs.Column) {
            $sort.Order = Get-ToggledSortOrder -CurrentOrder $sort.Order
        }
        else {
            $sort.Column = $eventArgs.Column
            $sort.Order = [System.Windows.Forms.SortOrder]::Ascending
        }
        Update-WeeklyReviewUi -State $state
    })

    $todaySessionsList.Add_ColumnClick({
        param($sender, $eventArgs)
        $sort = $state.ListSorts.TodaySessions
        if ($sort.Column -eq $eventArgs.Column) {
            $sort.Order = Get-ToggledSortOrder -CurrentOrder $sort.Order
        }
        else {
            $sort.Column = $eventArgs.Column
            $sort.Order = [System.Windows.Forms.SortOrder]::Descending
        }
        Update-SessionsUi -State $state
    })

    $weeklySessionsList.Add_ColumnClick({
        param($sender, $eventArgs)
        $sort = $state.ListSorts.WeeklySessions
        if ($sort.Column -eq $eventArgs.Column) {
            $sort.Order = Get-ToggledSortOrder -CurrentOrder $sort.Order
        }
        else {
            $sort.Column = $eventArgs.Column
            $sort.Order = [System.Windows.Forms.SortOrder]::Descending
        }
        Update-SessionsUi -State $state
    })

    $reviewList.Add_ColumnClick({
        param($sender, $eventArgs)
        $sort = $state.ListSorts.Review
        if ($sort.Column -eq $eventArgs.Column) {
            $sort.Order = Get-ToggledSortOrder -CurrentOrder $sort.Order
        }
        else {
            $sort.Column = $eventArgs.Column
            $sort.Order = [System.Windows.Forms.SortOrder]::Descending
        }
        Update-ClassificationReviewUi -State $state
    })
    $reviewList.Add_SelectedIndexChanged({
        $reviewClassifyButton.Enabled = $reviewList.SelectedItems.Count -gt 0
    })
    $healthUncategorizedList.Add_ColumnClick({
        param($sender, $eventArgs)
        $sort = $state.ListSorts.HealthUncategorized
        if ($sort.Column -eq $eventArgs.Column) {
            $sort.Order = Get-ToggledSortOrder -CurrentOrder $sort.Order
        }
        else {
            $sort.Column = $eventArgs.Column
            $sort.Order = [System.Windows.Forms.SortOrder]::Ascending
        }
        Update-ClassificationHealthUi -State $state
    })
    $healthRulesList.Add_ColumnClick({
        param($sender, $eventArgs)
        $sort = $state.ListSorts.HealthRules
        if ($sort.Column -eq $eventArgs.Column) {
            $sort.Order = Get-ToggledSortOrder -CurrentOrder $sort.Order
        }
        else {
            $sort.Column = $eventArgs.Column
            $sort.Order = [System.Windows.Forms.SortOrder]::Ascending
        }
        Update-ClassificationHealthUi -State $state
    })
    $healthSuggestionsList.Add_ColumnClick({
        param($sender, $eventArgs)
        $sort = $state.ListSorts.HealthSuggestions
        if ($sort.Column -eq $eventArgs.Column) {
            $sort.Order = Get-ToggledSortOrder -CurrentOrder $sort.Order
        }
        else {
            $sort.Column = $eventArgs.Column
            $sort.Order = [System.Windows.Forms.SortOrder]::Ascending
        }
        Update-ClassificationHealthUi -State $state
    })
    $healthSuggestionsList.Add_SelectedIndexChanged({
        $healthUseSuggestionButton.Enabled = $healthSuggestionsList.SelectedItems.Count -gt 0
    })
    $healthSuggestionsList.Add_DoubleClick({
        if ($healthUseSuggestionButton.Enabled) {
            $healthUseSuggestionButton.PerformClick()
        }
    })
    $mainTabs.Add_SelectedIndexChanged({
        Request-UiRefresh -State $state
        Update-Ui -State $state
    })

    $script:NotifyIcon.Add_DoubleClick({
        Show-TrackerWindow -Form $form
    })
    $pauseTrayItem.Add_Click({
        $state.TrackingEnabled = -not $state.TrackingEnabled
        if ($state.TrackingEnabled) {
            $state.HardLimitHandledFor = @{}
        }
        Request-UiRefresh -State $state
        Update-Ui -State $state
    })
    $focusTrayItem.Add_Click({
        if ($state.FocusMode.Enabled) {
            Stop-FocusMode -State $state
        }
        else {
            [void](Show-FocusModeDialog -State $state)
        }
        Request-UiRefresh -State $state
        Update-Ui -State $state
    })
    $reviewTrayItem.Add_Click({
        Open-TrackerSection -State $state -Form $form -TabKey "review"
    })
    $quickGlanceTrayItem.Add_Click({
        Toggle-QuickGlanceWindow -State $state
        Request-UiRefresh -State $state
        Update-Ui -State $state
    })
    $classifyTrayItem.Add_Click({
        if (Show-ClassifyCurrentActivityDialog -State $state) {
            Update-RuleList -RuleList $ruleList -Rules $state.Rules
            Clear-DerivedCaches -State $state
            Request-UiRefresh -State $state
            Update-Ui -State $state
        }
    })
    $openTrayItem.Add_Click({
        Open-TrackerSection -State $state -Form $form -TabKey "today"
    })
    $analyticsTrayItem.Add_Click({
        Open-TrackerSection -State $state -Form $form -TabKey "analytics"
    })
    $settingsTrayItem.Add_Click({
        Show-TrackerWindow -Form $form
        if (Show-SettingsEditor -State $state) {
            $sampleTimer.Interval = [int]([math]::Max(1, [double]$state.Settings.sampleIntervalSeconds) * 1000)
            Request-UiRefresh -State $state
            Update-Ui -State $state
        }
    })
    $exitTrayItem.Add_Click({
        $state.AllowExit = $true
        $form.Close()
    })

    $form.Add_Resize({
        if ($form.WindowState -eq [System.Windows.Forms.FormWindowState]::Minimized) {
            Hide-TrackerWindowToTray -Form $form -State $state -Message "The tracker is minimized to tray and keeps recording your active screen time."
        }
    })

    $quickGlanceForm.Add_FormClosing({
        param($sender, $eventArgs)

        if (-not $state.AllowExit -and $eventArgs.CloseReason -ne [System.Windows.Forms.CloseReason]::WindowsShutDown -and $eventArgs.CloseReason -ne [System.Windows.Forms.CloseReason]::TaskManagerClosing) {
            $eventArgs.Cancel = $true
            Hide-QuickGlanceWindow -State $state
        }
    })

    $form.Add_FormClosing({
        param($sender, $eventArgs)

        if (-not $state.AllowExit -and $eventArgs.CloseReason -ne [System.Windows.Forms.CloseReason]::WindowsShutDown -and $eventArgs.CloseReason -ne [System.Windows.Forms.CloseReason]::TaskManagerClosing) {
            $eventArgs.Cancel = $true
            Hide-TrackerWindowToTray -Form $form -State $state -Message "The window is closed to tray. Use the tray icon to reopen or exit."
            return
        }

        $sampleTimer.Stop()
        Stop-TrackedSession -State $state
        Save-IfDirty -State $state -Force
        Stop-BrowserBridgeJob -Job $state.BrowserBridgeJob
        if ($null -ne $quickGlanceForm) {
            $quickGlanceForm.Hide()
            $quickGlanceForm.Dispose()
        }
        $script:NotifyIcon.Visible = $false
        $script:NotifyIcon.Dispose()
    })

    if ($StartMinimized) {
        $form.WindowState = [System.Windows.Forms.FormWindowState]::Minimized
        $form.Add_Shown({
            Hide-TrackerWindowToTray -Form $form -State $state -Message "The tracker started in tray and is already counting your active screen time."
        })
    }

    [System.Windows.Forms.Application]::Run($form)
}

if ($SelfTest) {
    Invoke-SelfTestMode
    return
}

if ($MigrateStorage) {
    Ensure-Storage
    if (Migrate-LegacyUsageDataToDayFiles) {
        Write-Output "Storage migration completed."
    }
    else {
        Write-Output "Storage migration skipped."
    }
    return
}

if ($RebuildSummaryCache) {
    Ensure-Storage
    $count = Rebuild-UsageSummaryCache
    Write-Output ("Summary cache rebuilt for {0} days." -f $count)
    return
}

if (-not (Acquire-AppMutex)) {
    [void](Request-ReopenRunningInstance)
    return
}

try {
    Show-MainWindow
}
catch {
    Show-StartupFailure -Exception $_.Exception
    throw
}
finally {
    Release-AppMutex
}
