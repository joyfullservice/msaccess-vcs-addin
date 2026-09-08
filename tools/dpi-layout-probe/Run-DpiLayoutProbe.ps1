[CmdletBinding()]
param(
    [string]$Label,
    [int]$Monitor = -1,
    [string]$FixtureDirectory,
    [string]$ResultsDirectory,
    [switch]$ListMonitors
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$acForm = 2
$acDesign = 1
$acSaveYes = 1
$acQuitSaveNone = 2
$probeVersion = 1

if (-not $FixtureDirectory) {
    $FixtureDirectory = Join-Path $PSScriptRoot "fixtures"
}
if (-not $ResultsDirectory) {
    $ResultsDirectory = Join-Path $PSScriptRoot "results"
}

Add-Type -AssemblyName System.Windows.Forms

if (-not ("DpiLayoutProbe.NativeMethods" -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

namespace DpiLayoutProbe
{
    public static class NativeMethods
    {
        [DllImport("user32.dll")]
        public static extern uint GetDpiForWindow(IntPtr hwnd);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool SetWindowPos(
            IntPtr hwnd,
            IntPtr hwndInsertAfter,
            int x,
            int y,
            int width,
            int height,
            uint flags);
    }
}
"@
}

function Get-Monitors {
    $index = 0
    foreach ($screen in [System.Windows.Forms.Screen]::AllScreens) {
        [pscustomobject]@{
            Index       = $index
            DeviceName  = $screen.DeviceName
            Primary     = $screen.Primary
            X           = $screen.Bounds.X
            Y           = $screen.Bounds.Y
            Width       = $screen.Bounds.Width
            Height      = $screen.Bounds.Height
            WorkingArea = "$($screen.WorkingArea.X),$($screen.WorkingArea.Y) " +
                "$($screen.WorkingArea.Width)x$($screen.WorkingArea.Height)"
        }
        $index++
    }
}

$monitorInfo = @(Get-Monitors)
if ($ListMonitors) {
    $monitorInfo | Format-Table -AutoSize
    return
}

if (-not $Label) {
    Write-Host ""
    $monitorInfo | Format-Table -AutoSize
    Write-Host ""
    $Label = Read-Host "Run label (for example 100pct or 4k-150pct)"
    $monitorAnswer = Read-Host "Monitor index (press Enter for the primary display)"
    if ($monitorAnswer) {
        $parsedMonitor = 0
        if (-not [int]::TryParse($monitorAnswer, [ref]$parsedMonitor)) {
            throw "Monitor index must be an integer."
        }
        $Monitor = $parsedMonitor
    }
}

if ($Label -notmatch "^[A-Za-z0-9][A-Za-z0-9._-]*$") {
    throw "Label must contain only letters, digits, period, underscore, or hyphen."
}

if ($Monitor -lt 0) {
    $primary = $monitorInfo | Where-Object Primary | Select-Object -First 1
    $Monitor = $primary.Index
}
if ($Monitor -ge $monitorInfo.Count) {
    throw "Monitor index $Monitor does not exist. Use -ListMonitors to list displays."
}

$existingAccess = @(Get-Process MSACCESS -ErrorAction SilentlyContinue)
if ($existingAccess.Count -gt 0) {
    throw "Close every Microsoft Access window before running the probe. " +
        "This prevents COM from attaching to an existing DPI context."
}

$fixtures = @(Get-ChildItem -LiteralPath $FixtureDirectory -Filter "*.form" -File |
    Sort-Object Name)
if ($fixtures.Count -eq 0) {
    throw "No .form fixtures found in $FixtureDirectory"
}

$runDirectory = Join-Path $ResultsDirectory $Label
if (Test-Path -LiteralPath $runDirectory) {
    throw "Result directory already exists: $runDirectory. Use a new label."
}

$inputDirectory = Join-Path $runDirectory "input"
$plainDirectory = Join-Path $runDirectory "plain"
$designDirectory = Join-Path $runDirectory "design-save"
[System.IO.Directory]::CreateDirectory($inputDirectory) | Out-Null
[System.IO.Directory]::CreateDirectory($plainDirectory) | Out-Null
[System.IO.Directory]::CreateDirectory($designDirectory) | Out-Null

foreach ($fixture in $fixtures) {
    [System.IO.File]::Copy(
        $fixture.FullName,
        (Join-Path $inputDirectory $fixture.Name),
        $false)
}
$fixtures = @(Get-ChildItem -LiteralPath $inputDirectory -Filter "*.form" -File |
    Sort-Object Name)

$databasePath = Join-Path $runDirectory "probe.accdb"
$metadataPath = Join-Path $runDirectory "metadata.json"
$failurePath = Join-Path $runDirectory "failure.json"
$selectedScreen = [System.Windows.Forms.Screen]::AllScreens[$Monitor]
$access = $null
$accessProcess = $null
$startedAt = Get-Date

try {
    Write-Host "Starting a fresh Access instance..."
    $access = New-Object -ComObject Access.Application
    $access.Visible = $true
    Start-Sleep -Milliseconds 750

    $hwnd = [IntPtr][int64]$access.hWndAccessApp()
    $area = $selectedScreen.WorkingArea
    $positioned = [DpiLayoutProbe.NativeMethods]::SetWindowPos(
        $hwnd,
        [IntPtr]::Zero,
        $area.X + 20,
        $area.Y + 20,
        [Math]::Max(800, $area.Width - 40),
        [Math]::Max(600, $area.Height - 40),
        0x0040)
    if (-not $positioned) {
        throw "Could not position Access on monitor $Monitor."
    }

    Start-Sleep -Milliseconds 750
    $access.NewCurrentDatabase($databasePath)
    Start-Sleep -Milliseconds 750

    # NewCurrentDatabase can recreate the top-level window; position it again.
    $hwnd = [IntPtr][int64]$access.hWndAccessApp()
    [void][DpiLayoutProbe.NativeMethods]::SetWindowPos(
        $hwnd,
        [IntPtr]::Zero,
        $area.X + 20,
        $area.Y + 20,
        [Math]::Max(800, $area.Width - 40),
        [Math]::Max(600, $area.Height - 40),
        0x0040)
    Start-Sleep -Milliseconds 750

    $effectiveDpi = [int][DpiLayoutProbe.NativeMethods]::GetDpiForWindow($hwnd)
    $effectiveScale = [Math]::Round(100.0 * $effectiveDpi / 96.0, 2)
    $accessProcess = Get-Process MSACCESS | Sort-Object StartTime -Descending |
        Select-Object -First 1

    Write-Host ("Access effective DPI: {0} ({1}% scale)" -f
        $effectiveDpi, $effectiveScale)
    Write-Host ("Selected display: {0}, {1}x{2}" -f
        $selectedScreen.DeviceName,
        $selectedScreen.Bounds.Width,
        $selectedScreen.Bounds.Height)
    Write-Host ""

    $fixtureResults = @()
    $fixtureNumber = 0
    foreach ($fixture in $fixtures) {
        $fixtureNumber++
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($fixture.Name)
        $plainName = "dpi_plain_{0:D2}" -f $fixtureNumber
        $designName = "dpi_design_{0:D2}" -f $fixtureNumber
        $plainPath = Join-Path $plainDirectory $fixture.Name
        $designPath = Join-Path $designDirectory $fixture.Name

        Write-Host ("[{0}/{1}] {2}" -f
            $fixtureNumber, $fixtures.Count, $fixture.Name)

        # Path 1: LoadFromText followed immediately by SaveAsText.
        $access.LoadFromText($acForm, $plainName, $fixture.FullName)
        $access.SaveAsText($acForm, $plainName, $plainPath)
        $access.DoCmd.DeleteObject($acForm, $plainName)

        # Path 2: mirror InitializeForms after import.
        $access.LoadFromText($acForm, $designName, $fixture.FullName)
        $access.DoCmd.OpenForm($designName, $acDesign)
        Start-Sleep -Milliseconds 150
        $form = $access.Forms.Item($designName)
        $form.Tag = $form.Tag
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($form)
        $form = $null
        $access.DoCmd.Close($acForm, $designName, $acSaveYes)
        $access.SaveAsText($acForm, $designName, $designPath)
        $access.DoCmd.DeleteObject($acForm, $designName)

        $fixtureResults += [pscustomobject]@{
            name       = $baseName
            input      = [pscustomobject]@{
                path   = $fixture.FullName
                sha256 = (Get-FileHash -LiteralPath $fixture.FullName -Algorithm SHA256).Hash
            }
            plain      = [pscustomobject]@{
                path   = $plainPath
                sha256 = (Get-FileHash -LiteralPath $plainPath -Algorithm SHA256).Hash
            }
            designSave = [pscustomobject]@{
                path   = $designPath
                sha256 = (Get-FileHash -LiteralPath $designPath -Algorithm SHA256).Hash
            }
        }
    }

    $fileVersion = $accessProcess.MainModule.FileVersionInfo
    $metadata = [ordered]@{
        probeVersion          = $probeVersion
        label                 = $Label
        startedAtLocal        = $startedAt.ToString("o")
        completedAtLocal      = (Get-Date).ToString("o")
        fixtureDirectory      = (Resolve-Path -LiteralPath $FixtureDirectory).Path
        runDirectory          = $runDirectory
        databasePath          = $databasePath
        accessVersion         = [string]$access.Version
        accessProductVersion  = $fileVersion.ProductVersion
        accessFileVersion     = $fileVersion.FileVersion
        accessExecutable      = $accessProcess.MainModule.FileName
        accessEffectiveDpi    = $effectiveDpi
        accessEffectiveScale  = $effectiveScale
        selectedMonitor       = $Monitor
        selectedDisplay       = [ordered]@{
            deviceName = $selectedScreen.DeviceName
            primary    = $selectedScreen.Primary
            x          = $selectedScreen.Bounds.X
            y          = $selectedScreen.Bounds.Y
            width      = $selectedScreen.Bounds.Width
            height     = $selectedScreen.Bounds.Height
        }
        allDisplays           = $monitorInfo
        windowsVersion        = [System.Environment]::OSVersion.VersionString
        powershellVersion     = $PSVersionTable.PSVersion.ToString()
        fixtures              = $fixtureResults
    }

    $metadataJson = $metadata | ConvertTo-Json -Depth 8
    [System.IO.File]::WriteAllText(
        $metadataPath,
        $metadataJson,
        [System.Text.UTF8Encoding]::new($false))

    Write-Host ""
    Write-Host "Probe complete: $runDirectory" -ForegroundColor Green
    Write-Host "Close Access before changing scaling. If Access does not see the intended DPI, sign out/in before rerunning with a new label."
}
catch {
    $failure = [ordered]@{
        probeVersion = $probeVersion
        label        = $Label
        failedAt     = (Get-Date).ToString("o")
        error        = $_.Exception.Message
        scriptStack  = $_.ScriptStackTrace
    }
    $failureJson = $failure | ConvertTo-Json -Depth 5
    [System.IO.File]::WriteAllText(
        $failurePath,
        $failureJson,
        [System.Text.UTF8Encoding]::new($false))
    throw
}
finally {
    if ($null -ne $access) {
        try {
            $access.CloseCurrentDatabase()
        }
        catch {
            Write-Warning "Could not close the scratch database: $($_.Exception.Message)"
        }
        try {
            $access.Quit($acQuitSaveNone)
        }
        catch {
            Write-Warning "Could not close Access: $($_.Exception.Message)"
        }
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($access)
        $access = $null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}
