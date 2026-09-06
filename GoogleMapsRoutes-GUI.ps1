#Requires -Version 5.1
<#
.SYNOPSIS
    Google Maps Route & Map Generator — Advanced WPF Application.
    Supports manual route calculation (Fastest, Shortest, Eco-friendly with Avoid options),
    interactive vector Leaflet map (WebView2) with static fallback, publication-quality PDF reports,
    GPX / KML GPS exports, address geocode validation preview, and batch processing.

.DESCRIPTION
    Architecture:
      Modularized orchestrator loading decoupled subsystem components:
        - RouteMapFunctions.ps1  : Core backend APIs (Routes API v2, Geocoding, Static Maps, GPX/KML, Cost calculation)
        - Modules\AppConfig.ps1      : DPAPI security, JSON config persistence, localization catalog, logging
        - Modules\InteractiveMap.ps1 : WebView2 Leaflet vector map generator & WPF host integration
        - Modules\ReportPdf.ps1      : Headless Microsoft Edge vector PDF report engine
        - Modules\AsyncWorkers.ps1   : Isolated MTA runspaces for background geocoding, route calculation, & validation
        - Modules\AppXaml.ps1        : Complete WPF UI layout, responsive themes, styles, & templates
        - Modules\UiSettingsTab.ps1  : Settings UI, API key verification, Dark/Light theme, dynamic localization, cost tracking
        - Modules\UiManualTab.ps1    : Single route interactive calculator, avoid controls, waypoint manager, PDF/GPX/KML export
        - Modules\UiBatchTab.ps1     : Batch data import (Excel/CSV/JSON), pre-batch address validation, batch export

.NOTES
    Encoding: UTF-8 with BOM
    Compatibility: Windows PowerShell 5.1 & PowerShell 7+
    PS2EXE: 100% self-sufficient compilation via Build-Exe.ps1
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [switch]$NoShowDialog
)

Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force -ErrorAction SilentlyContinue

# ── 1. Enforce TLS 1.2 / TLS 1.1 protocols for HTTPS requests ────────────────
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls

# ── 2. Enforce STA mode for WPF ───────────────────────────────────────────
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) {
    $currentProcess = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    if ($currentProcess -match 'powershell\.exe|pwsh\.exe') {
        Start-Process -FilePath $currentProcess -ArgumentList "-NoProfile -STA -ExecutionPolicy Bypass -File `"$PSCommandPath`""
        exit
    }
}

# Establish application directory for both .ps1 and PS2EXE compiled execution
$script:AppDir = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
    $PSScriptRoot
} else {
    [System.IO.Path]::GetDirectoryName([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
}
if ([string]::IsNullOrWhiteSpace($PSScriptRoot)) {
    $PSScriptRoot = $script:AppDir
}

# ── 3. Load GUI, Drawing, and Security assemblies ────────────────────────────
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms, System.Drawing, System.Security

# DWM Dark Mode for Windows 10/11 window title bar
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class DwmDarkWindow {
    [DllImport("dwmapi.dll", PreserveSig = true)]
    public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
}
"@ -ErrorAction SilentlyContinue

# ── 4. Dot-Source Modular Components ─────────────────────────────────────────
# Uses $PSScriptRoot pattern matching Build-Exe.ps1 bundler regex for PS2EXE compatibility
. "$PSScriptRoot\RouteMapFunctions.ps1"
. "$PSScriptRoot\Modules\AppConfig.ps1"
. "$PSScriptRoot\Modules\InteractiveMap.ps1"
. "$PSScriptRoot\Modules\ReportPdf.ps1"
. "$PSScriptRoot\Modules\AsyncWorkers.ps1"
. "$PSScriptRoot\Modules\AppXaml.ps1"
. "$PSScriptRoot\Modules\UiSettingsTab.ps1"
. "$PSScriptRoot\Modules\UiManualTab.ps1"
. "$PSScriptRoot\Modules\UiBatchTab.ps1"

# ── 5. Initialize Application State & Configuration ──────────────────────────
$script:AppConfig = Load-AppConfig
$script:LocCatalog = Load-LocalizationConfig

Write-AppLog "Google Maps Routes GUI initialized. Current language: $($script:AppConfig.Language), Theme: $($script:AppConfig.Theme)" "INFO"

# ── 6. Load XAML Interface ───────────────────────────────────────────────────
$xamlString = Get-AppXaml
$reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xamlString))
$window = [System.Windows.Markup.XamlReader]::Load($reader)
$script:MainWindow = $window

# ── 7. Map Controls into Hashtable ───────────────────────────────────────────
$Controls = [ordered]@{}

$ctrlNames = @(
    'btnAddWaypoint', 'btnBatchExportGpx', 'btnBatchExportKml', 'btnBatchExportPdf',
    'btnBrowseBatchFile', 'btnBrowseSettingsOutputDir', 'btnCalculateManual', 'btnClearManualEnd', 'btnClearManualStart',
    'btnCopyInvalidAddresses', 'btnCopyUrl', 'btnExportCsv', 'btnExportExcel', 'btnExportJson',
    'btnManualExportGpx', 'btnManualExportKml', 'btnManualExportPdf', 'btnOpenGoogleMaps', 'btnOpenLangFile',
    'btnOpenLogFile', 'btnOpenOutputDir', 'btnQuickSettings', 'btnReloadBatchFile', 'btnReloadLang',
    'btnResetApiCounters', 'btnResetOverlayConfig', 'btnSaveMapAs', 'btnSaveSettings', 'btnStartBatch',
    'btnStopBatch', 'btnTestApiKey', 'btnThemeToggle', 'btnToggleKeyVisibility', 'btnValidateBatchGeocoding',
    'btnWpClear', 'btnWpDown', 'btnWpRemove', 'btnWpUp',
    'chkEnableBottomOverlay', 'chkEnableTopOverlay', 'chkManualAvoidFerries', 'chkManualAvoidHighways', 'chkManualAvoidTolls',
    'chkRememberKey', 'chkSettingsAvoidFerries', 'chkSettingsAvoidHighways', 'chkSettingsAvoidTolls', 'chkTrafficAware',
    'cmbApiCurrency', 'cmbAppLanguage', 'cmbBatchRouteType', 'cmbDefaultEmission', 'cmbDefaultMapSize',
    'cmbDefaultRouteType', 'cmbEmission', 'cmbSettingsLanguage', 'cmbSettingsTheme',
    'dgBatchInput', 'dgBatchPoints', 'dgBatchResults', 'dgGeocodeValidation', 'gridOverlayConfig', 'imgMapPreview',
    'lblApiBadge', 'lblApiUsageMonthlyCalls', 'lblApiUsageSessionCalls', 'lblBatchDefaultRouteType', 'lblBatchFileInfo',
    'lblBatchInputFile', 'lblBatchProgressText', 'lblBatchStats', 'lblColPropAlign', 'lblColPropName',
    'lblColPropOrder', 'lblColPropPanel', 'lblColPropShow', 'lblEstimatedCostMonthly', 'lblFooterStatus',
    'lblFooterVersion', 'lblFreeTierInfo', 'lblGeocodeValidationSummary', 'lblGoogleUrlDisplay', 'lblHeaderDist',
    'lblHeaderDur', 'lblHeaderType', 'lblKeyTestResult', 'lblManualAvoidHeader', 'lblManualDestination',
    'lblManualDist', 'lblManualEmission', 'lblManualOptHeader', 'lblManualOrigin', 'lblManualRouteName',
    'lblManualRoutePointsHeader', 'lblManualStatus', 'lblManualTime', 'lblManualType', 'lblManualWaypoints',
    'lblMapPlaceholder', 'lblSettingsApiDesc', 'lblSettingsApiHeader', 'lblSettingsApiLabel', 'lblSettingsApiUsageDesc',
    'lblSettingsApiUsageHeader', 'lblSettingsAvoidHeader', 'lblSettingsDefaultEmission', 'lblSettingsDefaultMapSize',
    'lblSettingsDefaultRouteType', 'lblSettingsLangHeader', 'lblSettingsLangLabel', 'lblSettingsOutputDir',
    'lblSettingsOverlayDesc', 'lblSettingsOverlayHeader', 'lblSettingsPrefHeader', 'lblSettingsThemeHeader', 'lblSettingsThemeLabel',
    'lstWaypoints', 'pbBatchProgress', 'pnlEmission', 'pnlInteractiveMapHost', 'rbTypeEco', 'rbTypeFastest',
    'rbTypeShortest', 'rbViewInteractive', 'rbViewStatic', 'tabBatchSub', 'tabItemBatch', 'tabItemManual',
    'tabItemSettings', 'tabMain', 'tabSubInput', 'tabSubLog', 'tabSubPoints', 'tabSubResults', 'tabSubValidation',
    'txtBatchFilePath', 'txtBatchLog', 'txtHeaderSubtitle', 'txtHeaderTitle', 'txtManualEnd',
    'txtManualName', 'txtManualStart', 'txtNewWaypoint', 'txtSettingsApiKey', 'txtSettingsApiKeyVisible', 'txtSettingsOutputDir',
    'txtSettingsCartoApiKey', 'txtSettingsCartoApiKeyVisible', 'btnToggleCartoKeyVisibility',
    'lblSettingsCartoApiHeader', 'lblSettingsCartoApiDesc', 'lblSettingsCartoApiLabel'
)

foreach ($n in $ctrlNames) {
    $c = $window.FindName($n)
    if ($c) { $Controls[$n] = $c }
}

$script:OverlayPropKeys = @(
    'StartGeocoded', 'EndGeocoded', 'Distance', 'Duration', 'Timestamp',
    'RouteName', 'RouteType', 'Waypoints', 'StartRaw', 'EndRaw'
)
foreach ($k in $script:OverlayPropKeys) {
    foreach ($pfx in @('chkProp_', 'cmbPanel_', 'cmbAlign_', 'cmbOrder_', 'lblProp_')) {
        $n = "$pfx$k"
        $c = $window.FindName($n)
        if ($c) { $Controls[$n] = $c }
    }
}

$script:Controls = $Controls
foreach ($k in $Controls.Keys) {
    Set-Variable -Name $k -Value $Controls[$k] -Scope Script
}

# ── 8. Register Subsystem UI Events ──────────────────────────────────────────
Register-UiSettingsTabEvents -Controls $Controls -Window $window
Register-UiManualTabEvents   -Controls $Controls -Window $window
Register-UiBatchTabEvents    -Controls $Controls -Window $window

# ── 9. Populate Saved Settings into Controls ─────────────────────────────────
if ($script:AppConfig.ApiKey) {
    $Controls.txtSettingsApiKey.Password = $script:AppConfig.ApiKey
    $Controls.txtSettingsApiKeyVisible.Text = $script:AppConfig.ApiKey
    $Controls.lblApiBadge.Text = 'API: Configured'
    $Controls.lblApiBadge.Foreground = [System.Windows.Media.Brushes]::LightGreen
} else {
    $Controls.lblApiBadge.Text = 'API: Not Configured'
    $Controls.lblApiBadge.Foreground = [System.Windows.Media.Brushes]::Orange
}
if ($script:AppConfig.CartoApiKey) {
    if ($Controls.txtSettingsCartoApiKey) { $Controls.txtSettingsCartoApiKey.Password = $script:AppConfig.CartoApiKey }
    if ($Controls.txtSettingsCartoApiKeyVisible) { $Controls.txtSettingsCartoApiKeyVisible.Text = $script:AppConfig.CartoApiKey }
}
$Controls.chkRememberKey.IsChecked = if ($null -ne $script:AppConfig.RememberKey) { [bool]$script:AppConfig.RememberKey } else { $true }

# Preferences
if ($Controls.cmbDefaultRouteType -and $script:AppConfig.DefaultRouteType) {
    foreach ($it in $Controls.cmbDefaultRouteType.Items) {
        if ($it.Tag -eq $script:AppConfig.DefaultRouteType) {
            $Controls.cmbDefaultRouteType.SelectedItem = $it
            break
        }
    }
}

if ($Controls.cmbDefaultEmission -and $script:AppConfig.DefaultEmissionType) {
    foreach ($it in $Controls.cmbDefaultEmission.Items) {
        if ($it.Tag -eq $script:AppConfig.DefaultEmissionType) {
            $Controls.cmbDefaultEmission.SelectedItem = $it
            break
        }
    }
}

# Map size
if ($Controls.cmbDefaultMapSize -and $script:AppConfig.MapWidth -and $script:AppConfig.MapHeight) {
    $targetSize = "$($script:AppConfig.MapWidth)x$($script:AppConfig.MapHeight)"
    foreach ($it in $Controls.cmbDefaultMapSize.Items) {
        if ($it.Tag -eq $targetSize) {
            $Controls.cmbDefaultMapSize.SelectedItem = $it
            break
        }
    }
}

# Sync Default Route Type to Manual & Batch tabs
if ($script:AppConfig.DefaultRouteType) {
    if ($script:AppConfig.DefaultRouteType -eq 'Shortest' -and $Controls.rbTypeShortest) {
        $Controls.rbTypeShortest.IsChecked = $true
    } elseif ($script:AppConfig.DefaultRouteType -eq 'Eco' -and $Controls.rbTypeEco) {
        $Controls.rbTypeEco.IsChecked = $true
    } elseif ($Controls.rbTypeFastest) {
        $Controls.rbTypeFastest.IsChecked = $true
    }

    if ($Controls.cmbBatchRouteType) {
        foreach ($it in $Controls.cmbBatchRouteType.Items) {
            if ($it.Tag -eq $script:AppConfig.DefaultRouteType) {
                $Controls.cmbBatchRouteType.SelectedItem = $it
                break
            }
        }
    }
}

# Sync Default Emission to Manual tab
if ($Controls.cmbEmission -and $script:AppConfig.DefaultEmissionType) {
    foreach ($it in $Controls.cmbEmission.Items) {
        if ($it.Tag -eq $script:AppConfig.DefaultEmissionType) {
            $Controls.cmbEmission.SelectedItem = $it
            break
        }
    }
}

# Avoid options
$Controls.chkSettingsAvoidTolls.IsChecked = [bool]$script:AppConfig.AvoidTolls
$Controls.chkSettingsAvoidHighways.IsChecked = [bool]$script:AppConfig.AvoidHighways
$Controls.chkSettingsAvoidFerries.IsChecked = [bool]$script:AppConfig.AvoidFerries

$Controls.chkManualAvoidTolls.IsChecked = [bool]$script:AppConfig.AvoidTolls
$Controls.chkManualAvoidHighways.IsChecked = [bool]$script:AppConfig.AvoidHighways
$Controls.chkManualAvoidFerries.IsChecked = [bool]$script:AppConfig.AvoidFerries

# Output Directory
if ($script:AppConfig.OutputDirectory) {
    $Controls.txtSettingsOutputDir.Text = $script:AppConfig.OutputDirectory
}

# Overlay Config
if ($script:AppConfig.OverlayConfig) {
    Set-OverlayConfigUi -cfg $script:AppConfig.OverlayConfig -Controls $Controls
}

# Currency
if ($script:AppConfig.ApiUsage -and $script:AppConfig.ApiUsage.PreferredCurrency -and $Controls.cmbApiCurrency) {
    foreach ($it in $Controls.cmbApiCurrency.Items) {
        if ($it.Tag -eq $script:AppConfig.ApiUsage.PreferredCurrency) {
            $Controls.cmbApiCurrency.SelectedItem = $it
            break
        }
    }
}

# ── 10. Apply Active Theme & Language ────────────────────────────────────────
$activeTheme = if ($script:AppConfig.Theme) { $script:AppConfig.Theme } else { 'Dark' }
Set-AppTheme -Theme $activeTheme -Window $window -Controls $Controls

$activeLang = if ($script:AppConfig.Language) { $script:AppConfig.Language } else { 'en' }
Apply-AppLanguage -LanguageCode $activeLang -Window $window -Controls $Controls

# Update API usage counter displays
Update-ApiUsageBadgeText

# ── 11. Window Closing Lifecycle ─────────────────────────────────────────────
$window.Add_Closing({
    if ($script:ActiveBatchTimer) { try { $script:ActiveBatchTimer.Stop() } catch { } }
    if ($script:ActiveManualTimer) { try { $script:ActiveManualTimer.Stop() } catch { } }
    if ($script:ActiveTestTimer) { try { $script:ActiveTestTimer.Stop() } catch { } }
    if ($script:SyncState) { $script:SyncState.CancelRequested = $true }
})

# ── 12. Show Application Window ──────────────────────────────────────────────
if (-not $NoShowDialog) {
    $window.ShowDialog() | Out-Null
}
