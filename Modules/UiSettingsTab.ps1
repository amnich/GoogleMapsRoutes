#Requires -Version 5.1
<#
.SYNOPSIS
    Google Maps Routes & Map Generator — Settings Tab & System UI Handlers.
.DESCRIPTION
    Manages API key security (DPAPI), async API key tester, theme switching (Dark/Light),
    multi-language localization (EN/DE/PL), map overlay banner designer, and
    API usage tracking & cost estimation (Feature 6.S).
.NOTES
    Encoding: UTF-8 with BOM
#>

function Update-ApiUsageBadgeText {
    [CmdletBinding()]
    param()

    if (-not $script:AppConfig -or -not $script:AppConfig.ApiUsage) { return }
    $u = $script:AppConfig.ApiUsage
    $curr = if ($u.PreferredCurrency) { $u.PreferredCurrency } else { 'USD' }

    $costObj = Get-EstimatedApiCost -GeocodingCalls $u.MonthlyCallsGeocoding `
        -RoutesBasicCalls $u.MonthlyCallsRoutes `
        -RoutesAdvancedCalls 0 `
        -StaticMapsCalls $u.MonthlyCallsStatic `
        -Currency $curr

    if ($script:Controls -and $script:Controls.txtApiUsageBadge) {
        $script:Controls.txtApiUsageBadge.Text = "API: $($costObj.FormattedCost)"
    }
    if ($script:Controls -and $script:Controls.lblApiUsageSessionCalls) {
        $totSession = $u.SessionCallsGeocoding + $u.SessionCallsRoutes + $u.SessionCallsStatic
        $script:Controls.lblApiUsageSessionCalls.Text = "$totSession calls"
    }
    if ($script:Controls -and $script:Controls.lblApiUsageMonthlyCalls) {
        $script:Controls.lblApiUsageMonthlyCalls.Text = "$($costObj.TotalCalls) calls"
    }
    if ($script:Controls -and $script:Controls.lblEstimatedCostMonthly) {
        $script:Controls.lblEstimatedCostMonthly.Text = $costObj.FormattedCost
    }
    if ($script:Controls -and $script:Controls.lblFreeTierInfo) {
        $script:Controls.lblFreeTierInfo.Text = "Free Tier Balance: $($costObj.FreeTierRemaining)"
    }
}

function Set-OverlayConfigUi {
    param($cfg, [object]$Controls)
    if (-not $cfg) { return }

    if ($Controls.chkEnableTopOverlay)    { $Controls.chkEnableTopOverlay.IsChecked = [bool]$cfg.EnableTopOverlay }
    if ($Controls.chkEnableBottomOverlay) { $Controls.chkEnableBottomOverlay.IsChecked = [bool]$cfg.EnableBottomOverlay }

    if ($cfg.Properties) {
        foreach ($propName in $script:OverlayPropKeys) {
            $p = $cfg.Properties.$propName
            if (-not $p) { continue }

            $chkCtrl = $Controls["chkProp_$propName"]
            $cmbPnl  = $Controls["cmbPanel_$propName"]
            $cmbAln  = $Controls["cmbAlign_$propName"]
            $cmbOrd  = $Controls["cmbOrder_$propName"]

            if ($chkCtrl) { $chkCtrl.IsChecked = [bool]$p.Enabled }
            if ($cmbPnl) {
                foreach ($it in $cmbPnl.Items) {
                    if ($it.Tag -eq [string]$p.Panel) { $cmbPnl.SelectedItem = $it; break }
                }
            }
            if ($cmbAln) {
                foreach ($it in $cmbAln.Items) {
                    if ($it.Tag -eq [string]$p.Alignment) { $cmbAln.SelectedItem = $it; break }
                }
            }
            if ($cmbOrd) {
                foreach ($it in $cmbOrd.Items) {
                    if ($it.Tag -eq [string]$p.Order) { $cmbOrd.SelectedItem = $it; break }
                }
            }
        }
    }
}

function Get-CurrentOverlayConfigFromUi {
    param([object]$Controls)

    $props = [ordered]@{}
    foreach ($propName in $script:OverlayPropKeys) {
        $chkCtrl = $Controls["chkProp_$propName"]
        $cmbPnl  = $Controls["cmbPanel_$propName"]
        $cmbAln  = $Controls["cmbAlign_$propName"]
        $cmbOrd  = $Controls["cmbOrder_$propName"]

        $pnlVal = if ($cmbPnl -and $cmbPnl.SelectedItem) { [string]$cmbPnl.SelectedItem.Tag } else { 'Bottom' }
        $alnVal = if ($cmbAln -and $cmbAln.SelectedItem) { [string]$cmbAln.SelectedItem.Tag } else { 'Left' }
        $ordVal = if ($cmbOrd -and $cmbOrd.SelectedItem) { [int]$cmbOrd.SelectedItem.Tag } else { 1 }
        $enVal  = if ($chkCtrl) { [bool]$chkCtrl.IsChecked } else { $true }

        $props[$propName] = [ordered]@{
            Enabled   = $enVal
            Panel     = $pnlVal
            Alignment = $alnVal
            Order     = $ordVal
        }
    }

    return [ordered]@{
        EnableTopOverlay    = if ($Controls.chkEnableTopOverlay) { [bool]$Controls.chkEnableTopOverlay.IsChecked } else { $true }
        EnableBottomOverlay = if ($Controls.chkEnableBottomOverlay) { [bool]$Controls.chkEnableBottomOverlay.IsChecked } else { $true }
        Properties          = $props
    }
}

function Get-CurrentOverlayConfig {
    [CmdletBinding()]
    param([object]$Controls = $null)

    if (-not $Controls -and $script:Controls) {
        $Controls = $script:Controls
    }

    if ($Controls -and $Controls.chkEnableTopOverlay) {
        return (Get-CurrentOverlayConfigFromUi -Controls $Controls)
    }

    if ($script:AppConfig -and $script:AppConfig.OverlayConfig) {
        return $script:AppConfig.OverlayConfig
    }

    return (Get-DefaultOverlayConfig)
}

Set-Item -Path "function:global:Get-CurrentOverlayConfig" -Value (Get-Item "function:Get-CurrentOverlayConfig").ScriptBlock -ErrorAction SilentlyContinue
Set-Item -Path "function:global:Get-CurrentOverlayConfigFromUi" -Value (Get-Item "function:Get-CurrentOverlayConfigFromUi").ScriptBlock -ErrorAction SilentlyContinue
Set-Item -Path "function:global:Set-OverlayConfigUi" -Value (Get-Item "function:Set-OverlayConfigUi").ScriptBlock -ErrorAction SilentlyContinue

function Set-AppTheme {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][string]$Theme = 'Dark',
        [Parameter(Mandatory = $false)][System.Windows.Window]$Window = $null,
        [Parameter(Mandatory = $false)][object]$Controls = $null
    )

    if ($Theme -notmatch '(?i)light|dark') { $Theme = 'Dark' }
    $isLight = ($Theme -match '(?i)light')
    $script:CurrentTheme = if ($isLight) { 'Light' } else { 'Dark' }

    if (-not $Window -and $script:MainWindow) { $Window = $script:MainWindow }
    if (-not $Controls -and $script:Controls) { $Controls = $script:Controls }

    $palette = if ($isLight) {
        [ordered]@{
            'BgDark'           = '#F1F5F9'
            'BgCard'           = '#FFFFFF'
            'BgCardHover'      = '#F8FAFC'
            'BgCardAlt'        = '#F8FAFC'
            'BorderCard'       = '#CBD5E1'
            'TextPrimary'      = '#0F172A'
            'TextSecondary'    = '#475569'
            'AccentBlue'       = '#2563EB'
            'AccentGreen'      = '#059669'
            'AccentAmber'      = '#D97706'
            'AccentRed'        = '#DC2626'
            'BgInput'          = '#FFFFFF'
            'BorderInput'      = '#CBD5E1'
            'BtnSecondaryBg'   = '#E2E8F0'
            'BtnSecondaryFg'   = '#0F172A'
            'GridLines'        = '#E2E8F0'
            'LogBg'            = '#F8FAFC'
            'LogFg'            = '#0369A1'
            'DataGridHeaderBg' = '#E2E8F0'
            'DataGridHeaderFg' = '#334155'
            'DataGridRowBg'    = '#FFFFFF'
            'DataGridAltRowBg' = '#F8FAFC'
        }
    }
    else {
        [ordered]@{
            'BgDark'           = '#0F172A'
            'BgCard'           = '#1E293B'
            'BgCardHover'      = '#293548'
            'BgCardAlt'        = '#162032'
            'BorderCard'       = '#334155'
            'TextPrimary'      = '#F8FAFC'
            'TextSecondary'    = '#94A3B8'
            'AccentBlue'       = '#2563EB'
            'AccentGreen'      = '#10B981'
            'AccentAmber'      = '#F59E0B'
            'AccentRed'        = '#EF4444'
            'BgInput'          = '#1E293B'
            'BorderInput'      = '#334155'
            'BtnSecondaryBg'   = '#334155'
            'BtnSecondaryFg'   = '#F8FAFC'
            'GridLines'        = '#2D3748'
            'LogBg'            = '#0A0F1D'
            'LogFg'            = '#38BDF8'
            'DataGridHeaderBg' = '#0F172A'
            'DataGridHeaderFg' = '#94A3B8'
            'DataGridRowBg'    = '#1E293B'
            'DataGridAltRowBg' = '#162032'
        }
    }

    if ($Window) {
        foreach ($k in $palette.Keys) {
            $c = [System.Windows.Media.ColorConverter]::ConvertFromString($palette[$k])
            $brush = [System.Windows.Media.SolidColorBrush]::new($c)
            $brush.Freeze()
            $Window.Resources[$k] = $brush
            $Window.Resources["Theme_$k"] = $brush
        }
        $Window.Resources['Theme_BgApp'] = $Window.Resources['BgDark']
        $Window.Resources['Theme_Border'] = $Window.Resources['BorderCard']

        $Window.Background = $Window.Resources['BgDark']
        $Window.Foreground = $Window.Resources['TextPrimary']

        # Update DWM title bar chrome
        try {
            $helper = [System.Windows.Interop.WindowInteropHelper]::new($Window)
            if ($helper.Handle -ne [IntPtr]::Zero) {
                $val = if ($isLight) { 0 } else { 1 }
                [DwmDarkWindow]::DwmSetWindowAttribute($helper.Handle, 20, [ref]$val, 4)
            }
        }
        catch {}
    }

    if ($Controls) {
        # Update overlay property labels foreground
        if ($script:OverlayPropKeys) {
            foreach ($k in $script:OverlayPropKeys) {
                $lblCtrl = $Controls["lblProp_$k"]
                if ($lblCtrl -and $Window) {
                    $lblCtrl.Foreground = $Window.Resources['TextPrimary']
                }
            }
        }

        # Update Toggle Button text/icon
        if ($Controls.btnThemeToggle) {
            $Controls.btnThemeToggle.Content = if ($isLight) { (Get-LocText 'ThemeLight' '🌙 Dark') } else { (Get-LocText 'ThemeDark' '☀️ Light') }
        }

        # Update Settings ComboBox
        if ($Controls.cmbSettingsTheme) {
            $script:SuppressThemeEvents = $true
            try {
                foreach ($it in $Controls.cmbSettingsTheme.Items) {
                    if ($it.Tag -eq $script:CurrentTheme) {
                        $Controls.cmbSettingsTheme.SelectedItem = $it
                        break
                    }
                }
            }
            finally {
                $script:SuppressThemeEvents = $false
            }
        }

        # Dynamically refresh interactive map tile style if route is active
        if (Get-Command Update-InteractiveRouteMap -ErrorAction SilentlyContinue) {
            Update-InteractiveRouteMap
        }
    }
}

function Apply-AppLanguage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][string]$LanguageCode = 'en',
        [Parameter(Mandatory = $false)][System.Windows.Window]$Window = $null,
        [Parameter(Mandatory = $false)][object]$Controls = $null
    )

    if (-not $Window -and $script:MainWindow) { $Window = $script:MainWindow }
    if (-not $Controls -and $script:Controls) { $Controls = $script:Controls }

    if (-not [string]::IsNullOrWhiteSpace($LanguageCode) -and $script:LocCatalog -and $script:LocCatalog.Languages -and $script:LocCatalog.Languages.PSObject.Properties[$LanguageCode.ToLower()]) {
        $script:CurrentLanguage = $LanguageCode.ToLower()
    }
    else {
        $script:CurrentLanguage = 'en'
    }

    if (-not $Controls) { return }

    # Synchronize ComboBoxes without triggering duplicate events
    $script:SuppressLangEvents = $true
    try {
        if ($Controls.cmbAppLanguage) {
            foreach ($item in $Controls.cmbAppLanguage.Items) {
                if ($item.Tag -eq $script:CurrentLanguage) {
                    $Controls.cmbAppLanguage.SelectedItem = $item
                    break
                }
            }
        }
        if ($Controls.cmbSettingsLanguage) {
            foreach ($item in $Controls.cmbSettingsLanguage.Items) {
                if ($item.Tag -eq $script:CurrentLanguage) {
                    $Controls.cmbSettingsLanguage.SelectedItem = $item
                    break
                }
            }
        }
    }
    finally {
        $script:SuppressLangEvents = $false
    }

    # Window & Header
    if ($Window) { $Window.Title = (Get-LocText 'AppTitle') }
    if ($Controls.txtHeaderTitle)    { $Controls.txtHeaderTitle.Text = (Get-LocText 'AppTitle') }
    if ($Controls.txtHeaderSubtitle) { $Controls.txtHeaderSubtitle.Text = (Get-LocText 'AppSubtitle') }
    if ($Controls.btnQuickSettings)  { $Controls.btnQuickSettings.Content = (Get-LocText 'BtnQuickSettings') }
    if ($Controls.btnThemeToggle) {
        $Controls.btnThemeToggle.ToolTip = (Get-LocText 'ThemeToggleTip')
        $Controls.btnThemeToggle.Content = if ($script:CurrentTheme -eq 'Light') { (Get-LocText 'ThemeLight' '🌙 Dark') } else { (Get-LocText 'ThemeDark' '☀️ Light') }
    }

    # Main Tabs
    if ($Controls.tabItemManual)   { $Controls.tabItemManual.Header = (Get-LocText 'TabManual') }
    if ($Controls.tabItemBatch)    { $Controls.tabItemBatch.Header = (Get-LocText 'TabBatch') }
    if ($Controls.tabItemSettings) { $Controls.tabItemSettings.Header = (Get-LocText 'TabSettings') }

    # Tab 1: Manual Route
    if ($Controls.lblManualRoutePointsHeader) { $Controls.lblManualRoutePointsHeader.Text = (Get-LocText 'ManualHeaderRoutePoints') }
    if ($Controls.lblManualOrigin)            { $Controls.lblManualOrigin.Text = (Get-LocText 'ManualOrigin') }
    if ($Controls.lblManualWaypoints)         { $Controls.lblManualWaypoints.Text = (Get-LocText 'ManualWaypoints') }
    if ($Controls.txtNewWaypoint)             { $Controls.txtNewWaypoint.ToolTip = (Get-LocText 'ManualWaypointsTooltip') }
    if ($Controls.btnAddWaypoint)             { $Controls.btnAddWaypoint.Content = (Get-LocText 'ManualBtnAdd') }
    if ($Controls.btnWpUp)                    { $Controls.btnWpUp.Content = (Get-LocText 'ManualBtnUp') }
    if ($Controls.btnWpDown)                  { $Controls.btnWpDown.Content = (Get-LocText 'ManualBtnDown') }
    if ($Controls.btnWpRemove)                { $Controls.btnWpRemove.Content = (Get-LocText 'ManualBtnRemove') }
    if ($Controls.btnWpClear)                 { $Controls.btnWpClear.Content = (Get-LocText 'ManualBtnClear') }
    if ($Controls.lblManualDestination)       { $Controls.lblManualDestination.Text = (Get-LocText 'ManualDestination') }
    if ($Controls.lblManualRouteName)         { $Controls.lblManualRouteName.Text = (Get-LocText 'ManualRouteName') }
    if ($Controls.lblManualOptHeader)         { $Controls.lblManualOptHeader.Text = (Get-LocText 'ManualHeaderOptimization') }
    if ($Controls.rbTypeFastest)              { $Controls.rbTypeFastest.Content = (Get-LocText 'ManualOptFastest') }
    if ($Controls.rbTypeShortest)             { $Controls.rbTypeShortest.Content = (Get-LocText 'ManualOptShortest') }
    if ($Controls.rbTypeEco)                  { $Controls.rbTypeEco.Content = (Get-LocText 'ManualOptEco') }
    if ($Controls.lblManualEmission)          { $Controls.lblManualEmission.Text = (Get-LocText 'ManualEmission') }
    if ($Controls.chkTrafficAware)            { $Controls.chkTrafficAware.Content = (Get-LocText 'ManualTrafficAware') }

    # Avoid Options
    if ($Controls.lblManualAvoidHeader)   { $Controls.lblManualAvoidHeader.Text = (Get-LocText 'ManualAvoidHeader' 'Route Avoidance:') }
    if ($Controls.chkManualAvoidTolls)    { $Controls.chkManualAvoidTolls.Content = (Get-LocText 'AvoidTolls' '🚫 Avoid Tolls') }
    if ($Controls.chkManualAvoidHighways) { $Controls.chkManualAvoidHighways.Content = (Get-LocText 'AvoidHighways' '🚫 Avoid Highways') }
    if ($Controls.chkManualAvoidFerries)  { $Controls.chkManualAvoidFerries.Content = (Get-LocText 'AvoidFerries' '🚫 Avoid Ferries') }

    if ($Controls.btnCalculateManual -and $Controls.btnCalculateManual.IsEnabled) {
        $Controls.btnCalculateManual.Content = (Get-LocText 'ManualBtnCalculate')
    }
    if ($Controls.lblHeaderDist) { $Controls.lblHeaderDist.Text = (Get-LocText 'ManualStatDistance') }
    if ($Controls.lblHeaderDur)  { $Controls.lblHeaderDur.Text = (Get-LocText 'ManualStatDuration') }
    if ($Controls.lblHeaderType) { $Controls.lblHeaderType.Text = (Get-LocText 'ManualStatType') }
    if ($Controls.lblManualStatus -and $Controls.lblManualStatus.Text -match '(?i)idle|bereit|gotow') {
        $Controls.lblManualStatus.Text = (Get-LocText 'ManualStatusIdle')
    }
    if ($Controls.lblMapPlaceholder -and $Controls.lblMapPlaceholder.Visibility -eq [System.Windows.Visibility]::Visible) {
        $Controls.lblMapPlaceholder.Text = (Get-LocText 'ManualMapPlaceholder')
    }
    if ($Controls.lblGoogleUrlDisplay -and $Controls.lblGoogleUrlDisplay.Text -match '(?i)no generated|kein link|brak') {
        $Controls.lblGoogleUrlDisplay.Text = (Get-LocText 'ManualNoUrl')
    }

    # Manual Map Controls & Exports
    if ($Controls.rbViewStatic)         { $Controls.rbViewStatic.Content = (Get-LocText 'MapStatic' '🖼️ Static Map') }
    if ($Controls.rbViewInteractive)    { $Controls.rbViewInteractive.Content = (Get-LocText 'MapInteractive' '🗺️ Interactive Map') }
    if ($Controls.btnManualExportPdf)   { $Controls.btnManualExportPdf.Content = (Get-LocText 'BtnExportPdf' '📄 PDF Report') }
    if ($Controls.btnManualExportGpx)   { $Controls.btnManualExportGpx.Content = (Get-LocText 'BtnExportGpx' '💾 GPX') }
    if ($Controls.btnManualExportKml)   { $Controls.btnManualExportKml.Content = (Get-LocText 'BtnExportKml' '💾 KML') }
    if ($Controls.btnOpenGoogleMaps)    { $Controls.btnOpenGoogleMaps.Content = (Get-LocText 'ManualBtnGoogleMaps') }
    if ($Controls.btnCopyUrl)           { $Controls.btnCopyUrl.Content = (Get-LocText 'ManualBtnCopyUrl') }
    if ($Controls.btnSaveMapAs)         { $Controls.btnSaveMapAs.Content = (Get-LocText 'ManualBtnSaveMapAs') }

    # Tab 2: Batch Processing
    if ($Controls.lblBatchInputFile)        { $Controls.lblBatchInputFile.Text = (Get-LocText 'BatchInputFile') }
    if ($Controls.btnBrowseBatchFile)       { $Controls.btnBrowseBatchFile.Content = (Get-LocText 'BatchBtnBrowse') }
    if ($Controls.btnReloadBatchFile)       { $Controls.btnReloadBatchFile.Content = (Get-LocText 'BatchBtnReload') }
    if ($Controls.lblBatchFileInfo -and $Controls.lblBatchFileInfo.Text -match '(?i)no file|keine datei|brak') {
        $Controls.lblBatchFileInfo.Text = (Get-LocText 'BatchNoFileLoaded')
    }
    if ($Controls.lblBatchDefaultRouteType) { $Controls.lblBatchDefaultRouteType.Text = (Get-LocText 'BatchDefaultRouteType') }
    if ($Controls.btnStartBatch)            { $Controls.btnStartBatch.Content = (Get-LocText 'BatchBtnStart') }
    if ($Controls.btnStopBatch)             { $Controls.btnStopBatch.Content = (Get-LocText 'BatchBtnStop') }

    # Batch Subtabs
    if ($Controls.tabSubInput)      { $Controls.tabSubInput.Header = (Get-LocText 'BatchTabInputPreview') }
    if ($Controls.tabSubResults)    { $Controls.tabSubResults.Header = (Get-LocText 'BatchTabResults') }
    if ($Controls.tabSubPoints)     { $Controls.tabSubPoints.Header = (Get-LocText 'BatchTabPoints') }
    if ($Controls.tabSubValidation) { $Controls.tabSubValidation.Header = (Get-LocText 'BatchTabGeocodeValidation' '🔍 Geocode Validation') }
    if ($Controls.tabSubLog)        { $Controls.tabSubLog.Header = (Get-LocText 'BatchTabLog') }

    # Geocode Validation Subtab
    if ($Controls.btnValidateBatchGeocoding) { $Controls.btnValidateBatchGeocoding.Content = (Get-LocText 'BatchBtnValidate' '🔍 Validate Addresses') }
    if ($Controls.btnCopyInvalidAddresses)   { $Controls.btnCopyInvalidAddresses.Content = (Get-LocText 'BatchBtnCopyInvalid' '📋 Copy Invalid Addresses') }

    # Batch DataGrid Columns
    if ($Controls.dgBatchResults -and $Controls.dgBatchResults.Columns.Count -ge 10) {
        $Controls.dgBatchResults.Columns[0].Header = (Get-LocText 'BatchColId')
        $Controls.dgBatchResults.Columns[1].Header = (Get-LocText 'BatchColName')
        $Controls.dgBatchResults.Columns[2].Header = (Get-LocText 'BatchColOrigin')
        $Controls.dgBatchResults.Columns[3].Header = (Get-LocText 'BatchColDestination')
        $Controls.dgBatchResults.Columns[4].Header = (Get-LocText 'BatchColWaypoints')
        $Controls.dgBatchResults.Columns[5].Header = (Get-LocText 'BatchColType')
        $Controls.dgBatchResults.Columns[6].Header = (Get-LocText 'BatchColDistance')
        $Controls.dgBatchResults.Columns[7].Header = (Get-LocText 'BatchColDuration')
        $Controls.dgBatchResults.Columns[8].Header = (Get-LocText 'BatchColStatus')
        $Controls.dgBatchResults.Columns[9].Header = (Get-LocText 'BatchColMap')
    }

    # Points DataGrid Columns
    if ($Controls.dgBatchPoints -and $Controls.dgBatchPoints.Columns.Count -ge 9) {
        $Controls.dgBatchPoints.Columns[0].Header = (Get-LocText 'PointsColRouteId')
        $Controls.dgBatchPoints.Columns[1].Header = (Get-LocText 'PointsColRouteName')
        $Controls.dgBatchPoints.Columns[2].Header = (Get-LocText 'PointsColOrder')
        $Controls.dgBatchPoints.Columns[3].Header = (Get-LocText 'PointsColType')
        $Controls.dgBatchPoints.Columns[4].Header = (Get-LocText 'PointsColOriginalAddress')
        $Controls.dgBatchPoints.Columns[5].Header = (Get-LocText 'PointsColGeocodedAddress')
        $Controls.dgBatchPoints.Columns[6].Header = (Get-LocText 'PointsColGeocodeStatus')
        $Controls.dgBatchPoints.Columns[7].Header = (Get-LocText 'PointsColMatchType')
        $Controls.dgBatchPoints.Columns[8].Header = (Get-LocText 'PointsColIsFallback')
        if ($Controls.dgBatchPoints.Columns.Count -ge 11) {
            $Controls.dgBatchPoints.Columns[9].Header = (Get-LocText 'PointsColLatitude')
            $Controls.dgBatchPoints.Columns[10].Header = (Get-LocText 'PointsColLongitude')
        }
    }

    # Validation DataGrid Columns (8 columns: Index, Address, Role, Status, Precision, FormattedAddress, Latitude, Longitude)
    if ($Controls.dgGeocodeValidation -and $Controls.dgGeocodeValidation.Columns.Count -ge 8) {
        $Controls.dgGeocodeValidation.Columns[0].Header = '#'
        $Controls.dgGeocodeValidation.Columns[1].Header = (Get-LocText 'ValidationColAddress' 'Input Address')
        $Controls.dgGeocodeValidation.Columns[2].Header = (Get-LocText 'ValidationColRole' 'Role')
        $Controls.dgGeocodeValidation.Columns[3].Header = (Get-LocText 'ValidationColStatus' 'Status')
        $Controls.dgGeocodeValidation.Columns[4].Header = (Get-LocText 'ValidationColType' 'Precision')
        $Controls.dgGeocodeValidation.Columns[5].Header = (Get-LocText 'ValidationColFormatted' 'Google Normalized Address')
        $Controls.dgGeocodeValidation.Columns[6].Header = (Get-LocText 'ValidationColLat' 'Latitude')
        $Controls.dgGeocodeValidation.Columns[7].Header = (Get-LocText 'ValidationColLng' 'Longitude')
    }

    if ($Controls.lblBatchProgressText -and $Controls.lblBatchProgressText.Text -match '(?i)ready|bereit|gotow') {
        $Controls.lblBatchProgressText.Text = (Get-LocText 'BatchProgressReady')
    }
    if ($Controls.btnOpenOutputDir)   { $Controls.btnOpenOutputDir.Content = (Get-LocText 'BatchBtnOpenOutputDir') }
    if ($Controls.btnExportExcel)     { $Controls.btnExportExcel.Content = (Get-LocText 'BatchBtnExportExcel') }
    if ($Controls.btnExportCsv)       { $Controls.btnExportCsv.Content = (Get-LocText 'BatchBtnExportCsv') }
    if ($Controls.btnExportJson)      { $Controls.btnExportJson.Content = (Get-LocText 'BatchBtnExportJson') }
    if ($Controls.btnBatchExportPdf)  { $Controls.btnBatchExportPdf.Content = (Get-LocText 'BtnExportPdf' '📄 PDF Report') }
    if ($Controls.btnBatchExportGpx)  { $Controls.btnBatchExportGpx.Content = (Get-LocText 'BtnExportGpx' '💾 GPX') }
    if ($Controls.btnBatchExportKml)  { $Controls.btnBatchExportKml.Content = (Get-LocText 'BtnExportKml' '💾 KML') }

    # Tab 3: Settings
    if ($Controls.lblSettingsApiHeader)        { $Controls.lblSettingsApiHeader.Text = (Get-LocText 'SettingsHeaderApi') }
    if ($Controls.lblSettingsApiDesc)          { $Controls.lblSettingsApiDesc.Text = (Get-LocText 'SettingsApiDesc') }
    if ($Controls.lblSettingsApiLabel)         { $Controls.lblSettingsApiLabel.Text = (Get-LocText 'SettingsApiLabel') }
    if ($Controls.lblSettingsCartoApiHeader)   { $Controls.lblSettingsCartoApiHeader.Text = (Get-LocText 'SettingsHeaderCartoApi' 'CARTO API Key (Interactive Map)') }
    if ($Controls.lblSettingsCartoApiDesc)     { $Controls.lblSettingsCartoApiDesc.Text = (Get-LocText 'SettingsCartoApiDesc' 'Optional. Required only if using authenticated or commercial CARTO private basemap services. Default public raster tiles work without a key.') }
    if ($Controls.lblSettingsCartoApiLabel)    { $Controls.lblSettingsCartoApiLabel.Text = (Get-LocText 'SettingsCartoApiLabel' 'CARTO API Key / Access Token:') }
    if ($Controls.btnTestApiKey -and $Controls.btnTestApiKey.IsEnabled) {
        $Controls.btnTestApiKey.Content = (Get-LocText 'SettingsBtnTestKey')
    }
    if ($Controls.chkRememberKey)              { $Controls.chkRememberKey.Content = (Get-LocText 'SettingsChkRemember') }

    # API Usage Section
    if ($Controls.lblSettingsApiUsageHeader)   { $Controls.lblSettingsApiUsageHeader.Text = (Get-LocText 'ApiUsageHeader' 'API Usage & Cost Estimation') }
    if ($Controls.btnResetApiCounters)         { $Controls.btnResetApiCounters.Content = (Get-LocText 'ApiUsageBtnReset' '🔄 Reset Month') }

    # Preferences Section
    if ($Controls.lblSettingsPrefHeader)       { $Controls.lblSettingsPrefHeader.Text = (Get-LocText 'SettingsHeaderPreferences') }
    if ($Controls.lblSettingsDefaultRouteType) { $Controls.lblSettingsDefaultRouteType.Text = (Get-LocText 'SettingsDefaultRouteType') }
    if ($Controls.lblSettingsDefaultEmission)  { $Controls.lblSettingsDefaultEmission.Text = (Get-LocText 'SettingsDefaultEmission') }
    if ($Controls.lblSettingsDefaultMapSize)   { $Controls.lblSettingsDefaultMapSize.Text = (Get-LocText 'SettingsDefaultMapSize') }
    if ($Controls.lblSettingsOutputDir)        { $Controls.lblSettingsOutputDir.Text = (Get-LocText 'SettingsOutputDir') }
    if ($Controls.btnBrowseSettingsOutputDir)  { $Controls.btnBrowseSettingsOutputDir.Content = (Get-LocText 'SettingsBtnBrowseOutputDir') }

    # Avoid Defaults in Settings
    if ($Controls.lblSettingsAvoidHeader)      { $Controls.lblSettingsAvoidHeader.Text = (Get-LocText 'SettingsAvoidTitle' 'Default Route Avoidance Options:') }
    if ($Controls.chkSettingsAvoidTolls)       { $Controls.chkSettingsAvoidTolls.Content = (Get-LocText 'AvoidTolls' '🚫 Avoid Tolls') }
    if ($Controls.chkSettingsAvoidHighways)    { $Controls.chkSettingsAvoidHighways.Content = (Get-LocText 'AvoidHighways' '🚫 Avoid Highways') }
    if ($Controls.chkSettingsAvoidFerries)     { $Controls.chkSettingsAvoidFerries.Content = (Get-LocText 'AvoidFerries' '🚫 Avoid Ferries') }

    # Language & Theme Section
    if ($Controls.lblSettingsLangHeader)       { $Controls.lblSettingsLangHeader.Text = (Get-LocText 'SettingsHeaderLanguage') }
    if ($Controls.lblSettingsLangLabel)        { $Controls.lblSettingsLangLabel.Text = (Get-LocText 'SettingsLanguageLabel') }
    if ($Controls.btnOpenLangFile)             { $Controls.btnOpenLangFile.Content = (Get-LocText 'SettingsBtnOpenLangFile') }
    if ($Controls.btnReloadLang)               { $Controls.btnReloadLang.Content = (Get-LocText 'SettingsBtnReloadLang') }
    if ($Controls.lblSettingsThemeHeader)      { $Controls.lblSettingsThemeHeader.Text = (Get-LocText 'SettingsThemeLabel') }
    if ($Controls.lblSettingsThemeLabel)       { $Controls.lblSettingsThemeLabel.Text = (Get-LocText 'SettingsThemeLabel') }

    # Overlay Section
    if ($Controls.lblSettingsOverlayHeader)    { $Controls.lblSettingsOverlayHeader.Text = (Get-LocText 'SettingsHeaderOverlay') }
    if ($Controls.lblSettingsOverlayDesc)      { $Controls.lblSettingsOverlayDesc.Text = (Get-LocText 'SettingsOverlayDesc') }
    if ($Controls.chkEnableTopOverlay)         { $Controls.chkEnableTopOverlay.Content = (Get-LocText 'SettingsOverlayTopEnable') }
    if ($Controls.chkEnableBottomOverlay)      { $Controls.chkEnableBottomOverlay.Content = (Get-LocText 'SettingsOverlayBottomEnable') }
    if ($Controls.btnResetOverlayConfig)       { $Controls.btnResetOverlayConfig.Content = (Get-LocText 'SettingsOverlayBtnReset') }
    if ($Controls.lblColPropName)              { $Controls.lblColPropName.Text = (Get-LocText 'OverlayColProperty') }
    if ($Controls.lblColPropShow)              { $Controls.lblColPropShow.Text = (Get-LocText 'OverlayColShow') }
    if ($Controls.lblColPropPanel)             { $Controls.lblColPropPanel.Text = (Get-LocText 'OverlayColPanel') }
    if ($Controls.lblColPropAlign)             { $Controls.lblColPropAlign.Text = (Get-LocText 'OverlayColAlign') }
    if ($Controls.lblColPropOrder)             { $Controls.lblColPropOrder.Text = (Get-LocText 'OverlayColOrder') }

    # Dynamic Overlay Property Labels
    if ($script:OverlayPropKeys) {
        foreach ($propName in $script:OverlayPropKeys) {
            $lblCtrl = $Controls["lblProp_$propName"]
            if ($lblCtrl) {
                $locKey = "OverlayProp$propName"
                $lblCtrl.Text = (Get-LocText $locKey $propName)
            }
        }
    }

    # Save & Open Log
    if ($Controls.btnSaveSettings) { $Controls.btnSaveSettings.Content = (Get-LocText 'SettingsBtnSave') }
    if ($Controls.btnOpenLogFile)  { $Controls.btnOpenLogFile.Content = (Get-LocText 'SettingsBtnOpenLog') }

    # Footer
    if ($Controls.lblFooterStatus)  { $Controls.lblFooterStatus.Text = (Get-LocText 'FooterReady') }
    if ($Controls.lblFooterVersion) { $Controls.lblFooterVersion.Text = (Get-LocText 'FooterVersion') }

    # Cost Badge
    Update-ApiUsageBadgeText
}

Set-Item -Path "function:global:Set-AppTheme" -Value (Get-Item "function:Set-AppTheme").ScriptBlock -ErrorAction SilentlyContinue
Set-Item -Path "function:global:Apply-AppLanguage" -Value (Get-Item "function:Apply-AppLanguage").ScriptBlock -ErrorAction SilentlyContinue
Set-Item -Path "function:global:Update-ApiUsageBadgeText" -Value (Get-Item "function:Update-ApiUsageBadgeText").ScriptBlock -ErrorAction SilentlyContinue

function Register-UiSettingsTabEvents {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][hashtable]$Controls,
        [Parameter(Mandatory = $true)][System.Windows.Window]$Window
    )

    $script:Controls = $Controls
    $script:MainWindow = $Window
    foreach ($k in $Controls.Keys) {
        Set-Variable -Name $k -Value $Controls[$k] -Scope Script -Force
        Set-Variable -Name "script:$k" -Value $Controls[$k] -Scope Script -Force
    }
    $script:OverlayPropKeys = @(
        'StartGeocoded', 'EndGeocoded', 'Distance', 'Duration', 'Timestamp',
        'RouteName', 'RouteType', 'Waypoints', 'StartRaw', 'EndRaw'
    )

    $aliases = [ordered]@{
        'txtSettingsApiKey'             = $Controls.txtSettingsApiKey
        'txtSettingsApiKeyVisible'      = $Controls.txtSettingsApiKeyVisible
        'btnToggleKeyVisibility'        = $Controls.btnToggleKeyVisibility
        'txtSettingsCartoApiKey'        = $Controls.txtSettingsCartoApiKey
        'txtSettingsCartoApiKeyVisible' = $Controls.txtSettingsCartoApiKeyVisible
        'btnToggleCartoKeyVisibility'   = $Controls.btnToggleCartoKeyVisibility
        'btnTestApiKey'                 = $Controls.btnTestApiKey
        'chkRememberKey'           = $Controls.chkRememberKey
        'lblKeyTestResult'         = $Controls.lblKeyTestResult
        'lblApiBadge'              = $Controls.lblApiBadge
        'btnApiUsageBadge'         = $Controls.btnApiUsageBadge
        'cmbApiCurrency'           = $Controls.cmbApiCurrency
        'btnResetApiCounters'      = $Controls.btnResetApiCounters
        'cmbDefaultRouteType'      = $Controls.cmbDefaultRouteType
        'cmbDefaultEmission'       = $Controls.cmbDefaultEmission
        'chkSettingsAvoidTolls'    = $Controls.chkSettingsAvoidTolls
        'chkSettingsAvoidHighways' = $Controls.chkSettingsAvoidHighways
        'chkSettingsAvoidFerries'  = $Controls.chkSettingsAvoidFerries
        'cmbDefaultMapSize'        = $Controls.cmbDefaultMapSize
        'txtSettingsOutputDir'     = $Controls.txtSettingsOutputDir
        'btnBrowseOutputDir'       = $Controls.btnBrowseSettingsOutputDir
        'btnResetOverlay'          = $Controls.btnResetOverlayConfig
        'cmbAppLanguage'           = $Controls.cmbAppLanguage
        'cmbSettingsLanguage'      = $Controls.cmbSettingsLanguage
        'btnOpenLangFile'          = $Controls.btnOpenLangFile
        'btnReloadLang'            = $Controls.btnReloadLang
        'cmbSettingsTheme'         = $Controls.cmbSettingsTheme
        'btnThemeToggle'           = $Controls.btnThemeToggle
        'btnSaveSettings'          = $Controls.btnSaveSettings
        'btnOpenLogFile'           = $Controls.btnOpenLogFile
        'btnQuickSettings'         = $Controls.btnQuickSettings
        'tabMain'                  = $Controls.tabMain
    }
    foreach ($entry in $aliases.GetEnumerator()) {
        Set-Variable -Name $entry.Key -Value $entry.Value -Scope Script -Force
        Set-Variable -Name "script:$($entry.Key)" -Value $entry.Value -Scope Script -Force
    }

    $txtSettingsApiKey             = $aliases['txtSettingsApiKey']
    $txtSettingsApiKeyVisible      = $aliases['txtSettingsApiKeyVisible']
    $btnToggleKeyVisibility        = $aliases['btnToggleKeyVisibility']
    $txtSettingsCartoApiKey        = $aliases['txtSettingsCartoApiKey']
    $txtSettingsCartoApiKeyVisible = $aliases['txtSettingsCartoApiKeyVisible']
    $btnToggleCartoKeyVisibility   = $aliases['btnToggleCartoKeyVisibility']
    $btnTestApiKey                 = $aliases['btnTestApiKey']
    $chkRememberKey           = $aliases['chkRememberKey']
    $lblKeyTestResult         = $aliases['lblKeyTestResult']
    $lblApiBadge              = $aliases['lblApiBadge']
    $btnApiUsageBadge         = $aliases['btnApiUsageBadge']
    $cmbApiCurrency           = $aliases['cmbApiCurrency']
    $btnResetApiCounters      = $aliases['btnResetApiCounters']
    $cmbDefaultRouteType      = $aliases['cmbDefaultRouteType']
    $cmbDefaultEmission       = $aliases['cmbDefaultEmission']
    $chkSettingsAvoidTolls    = $aliases['chkSettingsAvoidTolls']
    $chkSettingsAvoidHighways = $aliases['chkSettingsAvoidHighways']
    $chkSettingsAvoidFerries  = $aliases['chkSettingsAvoidFerries']
    $cmbDefaultMapSize        = $aliases['cmbDefaultMapSize']
    $txtSettingsOutputDir     = $aliases['txtSettingsOutputDir']
    $btnBrowseOutputDir       = $aliases['btnBrowseOutputDir']
    $btnResetOverlay          = $aliases['btnResetOverlay']
    $cmbAppLanguage           = $aliases['cmbAppLanguage']
    $cmbSettingsLanguage      = $aliases['cmbSettingsLanguage']
    $btnOpenLangFile          = $aliases['btnOpenLangFile']
    $btnReloadLang            = $aliases['btnReloadLang']
    $cmbSettingsTheme         = $aliases['cmbSettingsTheme']
    $btnThemeToggle           = $aliases['btnThemeToggle']
    $btnSaveSettings          = $aliases['btnSaveSettings']
    $btnOpenLogFile           = $aliases['btnOpenLogFile']
    $btnQuickSettings         = $aliases['btnQuickSettings']
    $tabMain                  = $aliases['tabMain']

    # Populate Languages Dropdowns
    if ($cmbAppLanguage -and $cmbSettingsLanguage) {
        $cmbAppLanguage.Items.Clear()
        $cmbSettingsLanguage.Items.Clear()

        $langs = if ($script:LocCatalog -and $script:LocCatalog.Languages) {
            @($script:LocCatalog.Languages.PSObject.Properties.Name)
        } else { @('en', 'de', 'pl') }

        foreach ($l in $langs) {
            $dispName = switch ($l) {
                'de' { 'Deutsch' }
                'pl' { 'Polski' }
                default { 'English' }
            }
            $cbi1 = [System.Windows.Controls.ComboBoxItem]::new()
            $cbi1.Content = "[$($l.ToUpper())] $dispName"
            $cbi1.Tag = $l
            if ($l -eq $script:CurrentLanguage) { $cbi1.IsSelected = $true }
            $cmbAppLanguage.Items.Add($cbi1) | Out-Null

            $cbi2 = [System.Windows.Controls.ComboBoxItem]::new()
            $cbi2.Content = "[$($l.ToUpper())] $dispName"
            $cbi2.Tag = $l
            if ($l -eq $script:CurrentLanguage) { $cbi2.IsSelected = $true }
            $cmbSettingsLanguage.Items.Add($cbi2) | Out-Null
        }
    }

    # Theme toggle button
    $btnThemeToggle.Add_Click({
        $newTheme = if ($script:CurrentTheme -eq 'Dark') { 'Light' } else { 'Dark' }
        Set-AppTheme -Theme $newTheme -Window $Window -Controls $Controls
    })

    # Theme dropdown selection changed
    $cmbSettingsTheme.Add_SelectionChanged({
        if ($script:SuppressThemeEvents) { return }
        if ($cmbSettingsTheme.SelectedItem) {
            $th = [string]$cmbSettingsTheme.SelectedItem.Tag
            Set-AppTheme -Theme $th -Window $Window -Controls $Controls
        }
    })

    # Language dropdowns selection changed
    $cmbAppLanguage.Add_SelectionChanged({
        if ($script:SuppressLangEvents) { return }
        if ($cmbAppLanguage.SelectedItem) {
            $l = [string]$cmbAppLanguage.SelectedItem.Tag
            Apply-AppLanguage -LanguageCode $l -Window $Window -Controls $Controls
        }
    })

    $cmbSettingsLanguage.Add_SelectionChanged({
        if ($script:SuppressLangEvents) { return }
        if ($cmbSettingsLanguage.SelectedItem) {
            $l = [string]$cmbSettingsLanguage.SelectedItem.Tag
            Apply-AppLanguage -LanguageCode $l -Window $Window -Controls $Controls
        }
    })

    # API Key visibility toggle
    $btnToggleKeyVisibility.Add_Click({
        if ($txtSettingsApiKey.Visibility -eq [System.Windows.Visibility]::Visible) {
            $txtSettingsApiKeyVisible.Text = $txtSettingsApiKey.Password
            $txtSettingsApiKey.Visibility = [System.Windows.Visibility]::Collapsed
            $txtSettingsApiKeyVisible.Visibility = [System.Windows.Visibility]::Visible
            $btnToggleKeyVisibility.Content = (Get-LocText 'SettingsBtnHide' '🔒 Hide')
        }
        else {
            $txtSettingsApiKey.Password = $txtSettingsApiKeyVisible.Text
            $txtSettingsApiKeyVisible.Visibility = [System.Windows.Visibility]::Collapsed
            $txtSettingsApiKey.Visibility = [System.Windows.Visibility]::Visible
            $btnToggleKeyVisibility.Content = (Get-LocText 'SettingsBtnShow' '👁 Show')
        }
    })

    # CARTO API Key visibility toggle
    if ($btnToggleCartoKeyVisibility) {
        $btnToggleCartoKeyVisibility.Add_Click({
            if ($txtSettingsCartoApiKey.Visibility -eq [System.Windows.Visibility]::Visible) {
                $txtSettingsCartoApiKeyVisible.Text = $txtSettingsCartoApiKey.Password
                $txtSettingsCartoApiKey.Visibility = [System.Windows.Visibility]::Collapsed
                $txtSettingsCartoApiKeyVisible.Visibility = [System.Windows.Visibility]::Visible
                $btnToggleCartoKeyVisibility.Content = (Get-LocText 'SettingsBtnHide' '🔒 Hide')
            }
            else {
                $txtSettingsCartoApiKey.Password = $txtSettingsCartoApiKeyVisible.Text
                $txtSettingsCartoApiKeyVisible.Visibility = [System.Windows.Visibility]::Collapsed
                $txtSettingsCartoApiKey.Visibility = [System.Windows.Visibility]::Visible
                $btnToggleCartoKeyVisibility.Content = (Get-LocText 'SettingsBtnShow' '👁 Show')
            }
        })
    }

    # Test API Key Async
    $btnTestApiKey.Add_Click({
        $key = if ($txtSettingsApiKey.Visibility -eq [System.Windows.Visibility]::Visible) {
            $txtSettingsApiKey.Password.Trim()
        } else {
            $txtSettingsApiKeyVisible.Text.Trim()
        }

        if ([string]::IsNullOrWhiteSpace($key)) {
            $lblKeyTestResult.Text = '✕ Enter an API key first'
            $lblKeyTestResult.Foreground = [System.Windows.Media.Brushes]::Salmon
            return
        }

        $btnTestApiKey.IsEnabled = $false
        $btnTestApiKey.Content = '⏳ Testing...'
        $lblKeyTestResult.Text = 'Testing API key with Google Geocoding...'
        $lblKeyTestResult.Foreground = [System.Windows.Media.Brushes]::SkyBlue

        $testScript = {
            param($k, $lang)
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls
            if (Get-Command Test-GoogleApiKey -ErrorAction SilentlyContinue) {
                return (Test-GoogleApiKey -ApiKey $k -LanguageCode $lang)
            }
            try {
                $cleanLang = if ($lang) { ($lang -split '[-_]')[0].ToLower() } else { 'en' }
                $url = "https://maps.googleapis.com/maps/api/geocode/json?address=Warszawa&language=$cleanLang&key=$k"
                $resp = Invoke-RestMethod -Uri $url -Method GET -TimeoutSec 15
                if ($resp.status -eq 'OK' -or $resp.status -eq 'ZERO_RESULTS') {
                    return [PSCustomObject]@{ Valid = $true; Message = 'Klucz Google Maps API jest poprawny i aktywny.' }
                } elseif ($resp.status -eq 'REQUEST_DENIED') {
                    $m = if ($resp.error_message) { $resp.error_message } else { 'Żądanie odrzucone przez Google API.' }
                    return [PSCustomObject]@{ Valid = $false; Message = "Brak autoryzacji: $m" }
                } else {
                    return [PSCustomObject]@{ Valid = $false; Message = "Status API: $($resp.status)" }
                }
            } catch {
                return [PSCustomObject]@{ Valid = $false; Message = "Błąd połączenia: $($_.Exception.Message)" }
            }
        }

        $psTest = New-WorkerPowerShell -ScriptBlock $testScript
        $langCode = if ($script:CurrentLanguage) { $script:CurrentLanguage } else { 'en' }
        $psTest.AddArgument($key).AddArgument($langCode) | Out-Null
        $testHandle = $psTest.BeginInvoke()

        $testTimer = [System.Windows.Threading.DispatcherTimer]::new()
        $testTimer.Interval = [TimeSpan]::FromMilliseconds(150)
        $script:ActiveTestTimer = $testTimer
        $script:ActiveTestPs = $psTest
        $script:ActiveTestHandle = $testHandle
        $script:TestTimerTicks = 0

        $testTimer.Add_Tick({
            $localHandle = $script:ActiveTestHandle
            $localPs = $script:ActiveTestPs
            $script:TestTimerTicks++

            if ($localHandle -and $localHandle.IsCompleted) {
                if ($script:ActiveTestTimer) { try { $script:ActiveTestTimer.Stop() } catch { } }
                $btnTestApiKey.IsEnabled = $true
                $btnTestApiKey.Content = (Get-LocText 'SettingsBtnTestKey' '🔍 Test Key')

                try {
                    $res = $localPs.EndInvoke($localHandle)
                    $testResult = if ($res -and $res.Count -gt 0) { $res[0] } else { $null }
                    if ($testResult -and $testResult.Valid) {
                        $lblKeyTestResult.Text = "✓ Key Valid: $($testResult.Message)"
                        $lblKeyTestResult.Foreground = [System.Windows.Media.Brushes]::LightGreen
                        $lblApiBadge.Text = 'API: Active'
                        $lblApiBadge.Foreground = [System.Windows.Media.Brushes]::LightGreen
                        if ($script:AppConfig) {
                            $script:AppConfig.ApiKey = $key
                        }
                        Update-ApiUsageRecord -GeocodingInc 1
                        Update-ApiUsageBadgeText
                    }
                    elseif ($testResult) {
                        $lblKeyTestResult.Text = "✕ Test failed: $($testResult.Message)"
                        $lblKeyTestResult.Foreground = [System.Windows.Media.Brushes]::Salmon
                        $lblApiBadge.Text = 'API: Error'
                        $lblApiBadge.Foreground = [System.Windows.Media.Brushes]::Salmon
                    }
                    else {
                        $errMsg = 'No result returned from API verification.'
                        if ($localPs.Streams.Error.Count -gt 0) {
                            $errMsg = ($localPs.Streams.Error | ForEach-Object { $_.Exception.Message }) -join '; '
                        }
                        $lblKeyTestResult.Text = "✕ Test error: $errMsg"
                        $lblKeyTestResult.Foreground = [System.Windows.Media.Brushes]::Salmon
                        $lblApiBadge.Text = 'API: Error'
                        $lblApiBadge.Foreground = [System.Windows.Media.Brushes]::Salmon
                    }
                }
                catch {
                    $lblKeyTestResult.Text = "✕ Test exception: $($_.Exception.Message)"
                    $lblKeyTestResult.Foreground = [System.Windows.Media.Brushes]::Salmon
                }
                finally {
                    if ($localPs) { try { $localPs.Dispose() } catch { } }
                }
            }
            elseif ($script:TestTimerTicks -ge 140) {
                # 20s watchdog
                if ($script:ActiveTestTimer) { try { $script:ActiveTestTimer.Stop() } catch { } }
                $btnTestApiKey.IsEnabled = $true
                $btnTestApiKey.Content = (Get-LocText 'SettingsBtnTestKey' '🔍 Test Key')
                $lblKeyTestResult.Text = '✕ Test timed out (20s)'
                $lblKeyTestResult.Foreground = [System.Windows.Media.Brushes]::Salmon
                if ($localPs) { try { $localPs.Stop(); $localPs.Dispose() } catch { } }
            }
        })
        $testTimer.Start()
    })

    # Browse Output Dir
    $btnBrowseOutputDir.Add_Click({
        $fbd = [System.Windows.Forms.FolderBrowserDialog]::new()
        $fbd.Description = 'Select Results Output Folder'
        if (-not [string]::IsNullOrWhiteSpace($txtSettingsOutputDir.Text) -and (Test-Path $txtSettingsOutputDir.Text.Trim())) {
            $fbd.SelectedPath = $txtSettingsOutputDir.Text.Trim()
        }
        if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtSettingsOutputDir.Text = $fbd.SelectedPath
        }
    })

    # Reset Overlay Config to Defaults
    $btnResetOverlay.Add_Click({
        $def = Get-DefaultOverlayConfig
        Set-OverlayConfigUi -cfg $def -Controls $Controls
        [System.Windows.MessageBox]::Show('Map overlay banners reset to recommended default layout.', 'Overlay Reset', 'OK', 'Information')
    })

    # Header API Usage Badge Click -> Navigate to Settings (if present)
    if ($btnApiUsageBadge) {
        $btnApiUsageBadge.Add_Click({
            if ($tabMain) { $tabMain.SelectedIndex = 2 }
            elseif ($script:Controls -and $script:Controls.tabMain) { $script:Controls.tabMain.SelectedIndex = 2 }
        })
    }

    # Currency selection changed
    $cmbApiCurrency.Add_SelectionChanged({
        if ($cmbApiCurrency.SelectedItem -and $script:AppConfig -and $script:AppConfig.ApiUsage) {
            $curr = [string]$cmbApiCurrency.SelectedItem.Tag
            $script:AppConfig.ApiUsage.PreferredCurrency = $curr
            Update-ApiUsageBadgeText
        }
    })

    # Reset Monthly API Counters
    $btnResetApiCounters.Add_Click({
        $ask = [System.Windows.MessageBox]::Show("Are you sure you want to reset monthly API counters to 0?", "Reset Counters", "YesNo", "Question")
        if ($ask -eq [System.Windows.MessageBoxResult]::Yes -and $script:AppConfig -and $script:AppConfig.ApiUsage) {
            $script:AppConfig.ApiUsage.MonthlyCallsGeocoding = 0
            $script:AppConfig.ApiUsage.MonthlyCallsRoutes = 0
            $script:AppConfig.ApiUsage.MonthlyCallsStatic = 0
            Save-AppConfig -Config $script:AppConfig | Out-Null
            Update-ApiUsageBadgeText
            [System.Windows.MessageBox]::Show("Monthly counters reset to 0.", "Counters Reset", "OK", "Information")
        }
    })

    # Quick Settings button -> switch to Settings Tab
    $btnQuickSettings.Add_Click({
        if ($tabMain) { $tabMain.SelectedIndex = 2 }
        elseif ($script:Controls -and $script:Controls.tabMain) { $script:Controls.tabMain.SelectedIndex = 2 }
    })

    # Open Log File
    $btnOpenLogFile.Add_Click({
        if (Test-Path $script:LogFile) {
            Start-Process notepad.exe $script:LogFile
        }
    })

    # Open Localization JSON
    $btnOpenLangFile.Add_Click({
        $baseDir = if (-not [string]::IsNullOrWhiteSpace($script:AppDir)) {
            $script:AppDir
        } elseif (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
            $PSScriptRoot
        } else {
            [System.IO.Path]::GetDirectoryName([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
        }
        $locFile = if (-not [string]::IsNullOrWhiteSpace($script:AppDataDir)) { Join-Path $script:AppDataDir 'localization.json' } else { $null }
        if (-not $locFile -or -not (Test-Path $locFile)) {
            if (-not [string]::IsNullOrWhiteSpace($baseDir)) {
                $pLoc = Join-Path $baseDir 'localization.json'
                if (Test-Path $pLoc) {
                    $locFile = $pLoc
                } else {
                    $parentDir = Split-Path -Parent $baseDir
                    if (-not [string]::IsNullOrWhiteSpace($parentDir)) {
                        $pLoc = Join-Path $parentDir 'localization.json'
                        if (Test-Path $pLoc) { $locFile = $pLoc }
                    }
                }
            }
        }
        if ($locFile -and (Test-Path $locFile)) {
            Start-Process notepad.exe $locFile
        }
    })

    # Reload Localization JSON
    $btnReloadLang.Add_Click({
        $baseDir = if (-not [string]::IsNullOrWhiteSpace($script:AppDir)) {
            $script:AppDir
        } elseif (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
            $PSScriptRoot
        } else {
            [System.IO.Path]::GetDirectoryName([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
        }
        $locPath = if (-not [string]::IsNullOrWhiteSpace($script:AppDataDir)) { Join-Path $script:AppDataDir 'localization.json' } else { $null }
        if (-not $locPath -or -not (Test-Path $locPath)) {
            if (-not [string]::IsNullOrWhiteSpace($baseDir)) {
                $pLoc = Join-Path $baseDir 'localization.json'
                if (Test-Path $pLoc) {
                    $locPath = $pLoc
                } else {
                    $parentDir = Split-Path -Parent $baseDir
                    if (-not [string]::IsNullOrWhiteSpace($parentDir)) {
                        $pLoc = Join-Path $parentDir 'localization.json'
                        if (Test-Path $pLoc) { $locPath = $pLoc }
                    }
                }
            }
        }
        if ($locPath -and (Test-Path $locPath)) {
            $cat = Load-LocalizationConfig -FilePath $locPath
            if ($cat) {
                $script:LocCatalog = $cat
                Apply-AppLanguage -LanguageCode $script:CurrentLanguage -Window $Window -Controls $Controls
                [System.Windows.MessageBox]::Show((Get-LocText 'MsgLangReloaded' 'Language definitions reloaded successfully.'), (Get-LocText 'MsgExportTitle' 'Reload Complete'), 'OK', 'Information')
            }
        }
    })

    # Save Settings
    $btnSaveSettings.Add_Click({
        $key = if ($txtSettingsApiKey.Visibility -eq [System.Windows.Visibility]::Visible -and -not [string]::IsNullOrWhiteSpace($txtSettingsApiKey.Password)) {
            $txtSettingsApiKey.Password.Trim()
        } elseif (-not [string]::IsNullOrWhiteSpace($txtSettingsApiKeyVisible.Text)) {
            $txtSettingsApiKeyVisible.Text.Trim()
        } elseif (-not [string]::IsNullOrWhiteSpace($txtSettingsApiKey.Password)) {
            $txtSettingsApiKey.Password.Trim()
        } else {
            ''
        }

        # Keep both inputs in sync
        $txtSettingsApiKey.Password = $key
        $txtSettingsApiKeyVisible.Text = $key

        $cartoKey = if ($txtSettingsCartoApiKey -and $txtSettingsCartoApiKey.Visibility -eq [System.Windows.Visibility]::Visible -and -not [string]::IsNullOrWhiteSpace($txtSettingsCartoApiKey.Password)) {
            $txtSettingsCartoApiKey.Password.Trim()
        } elseif ($txtSettingsCartoApiKeyVisible -and -not [string]::IsNullOrWhiteSpace($txtSettingsCartoApiKeyVisible.Text)) {
            $txtSettingsCartoApiKeyVisible.Text.Trim()
        } elseif ($txtSettingsCartoApiKey -and -not [string]::IsNullOrWhiteSpace($txtSettingsCartoApiKey.Password)) {
            $txtSettingsCartoApiKey.Password.Trim()
        } else {
            ''
        }

        if ($txtSettingsCartoApiKey) { $txtSettingsCartoApiKey.Password = $cartoKey }
        if ($txtSettingsCartoApiKeyVisible) { $txtSettingsCartoApiKeyVisible.Text = $cartoKey }

        # Update API Badge
        if (-not [string]::IsNullOrWhiteSpace($key)) {
            $lblApiBadge.Text = 'API: Configured'
            $lblApiBadge.Foreground = [System.Windows.Media.Brushes]::LightGreen
        } else {
            $lblApiBadge.Text = 'API: Not Configured'
            $lblApiBadge.Foreground = [System.Windows.Media.Brushes]::Orange
        }

        $mapSize = if ($cmbDefaultMapSize.SelectedItem) { [string]$cmbDefaultMapSize.SelectedItem.Tag } else { '900x600' }
        $parts = $mapSize -split 'x'
        $w = if ($parts.Length -eq 2) { [int]$parts[0] } else { 900 }
        $h = if ($parts.Length -eq 2) { [int]$parts[1] } else { 600 }

        $overlayCfg = Get-CurrentOverlayConfigFromUi -Controls $Controls

        $selectedRouteType = if ($cmbDefaultRouteType.SelectedItem) { [string]$cmbDefaultRouteType.SelectedItem.Tag } else { 'Fastest' }
        $selectedEmission = if ($cmbDefaultEmission.SelectedItem) { [string]$cmbDefaultEmission.SelectedItem.Tag } else { 'GASOLINE' }
        $selectedLang = if ($cmbSettingsLanguage.SelectedItem) { [string]$cmbSettingsLanguage.SelectedItem.Tag } else { 'en' }
        $selectedTheme = if ($cmbSettingsTheme.SelectedItem) { [string]$cmbSettingsTheme.SelectedItem.Tag } else { 'Dark' }
        $selectedCurrency = if ($cmbApiCurrency -and $cmbApiCurrency.SelectedItem) { [string]$cmbApiCurrency.SelectedItem.Tag } else { 'USD' }

        if ($script:AppConfig -and $script:AppConfig.ApiUsage) {
            $script:AppConfig.ApiUsage.PreferredCurrency = $selectedCurrency
        }

        $newConfig = [ordered]@{
            ApiKey              = $key
            CartoApiKey         = $cartoKey
            RememberKey         = [bool]$chkRememberKey.IsChecked
            DefaultRouteType    = $selectedRouteType
            DefaultEmissionType = $selectedEmission
            MapWidth            = $w
            MapHeight           = $h
            OutputDirectory     = $txtSettingsOutputDir.Text.Trim()
            Language            = $selectedLang
            Theme               = $selectedTheme
            UseInteractiveMap   = $true
            AvoidTolls          = [bool]$chkSettingsAvoidTolls.IsChecked
            AvoidHighways       = [bool]$chkSettingsAvoidHighways.IsChecked
            AvoidFerries        = [bool]$chkSettingsAvoidFerries.IsChecked
            OverlayConfig       = $overlayCfg
            ApiUsage            = $script:AppConfig.ApiUsage
        }

        $ok = Save-AppConfig -Config $newConfig
        if ($ok) {
            # Synchronize active runtime settings across other tabs immediately:
            Set-AppTheme -Theme $selectedTheme -Window $Window -Controls $Controls
            Apply-AppLanguage -LanguageCode $selectedLang -Window $Window -Controls $Controls

            if ($Controls.chkManualAvoidTolls) { $Controls.chkManualAvoidTolls.IsChecked = [bool]$chkSettingsAvoidTolls.IsChecked }
            if ($Controls.chkManualAvoidHighways) { $Controls.chkManualAvoidHighways.IsChecked = [bool]$chkSettingsAvoidHighways.IsChecked }
            if ($Controls.chkManualAvoidFerries) { $Controls.chkManualAvoidFerries.IsChecked = [bool]$chkSettingsAvoidFerries.IsChecked }

            if ($Controls.cmbBatchRouteType) {
                foreach ($it in $Controls.cmbBatchRouteType.Items) {
                    if ($it.Tag -eq $selectedRouteType) {
                        $Controls.cmbBatchRouteType.SelectedItem = $it
                        break
                    }
                }
            }

            if ($selectedRouteType -eq 'Shortest' -and $Controls.rbTypeShortest) {
                $Controls.rbTypeShortest.IsChecked = $true
            } elseif ($selectedRouteType -eq 'Eco' -and $Controls.rbTypeEco) {
                $Controls.rbTypeEco.IsChecked = $true
            } elseif ($Controls.rbTypeFastest) {
                $Controls.rbTypeFastest.IsChecked = $true
            }

            if ($Controls.cmbEmission) {
                foreach ($it in $Controls.cmbEmission.Items) {
                    if ($it.Tag -eq $selectedEmission) {
                        $Controls.cmbEmission.SelectedItem = $it
                        break
                    }
                }
            }

            [System.Windows.MessageBox]::Show((Get-LocText 'MsgSettingsSaved' 'Settings saved successfully!'), (Get-LocText 'MsgSettingsSavedTitle' 'Saved'), 'OK', 'Information')
        }
    })
}
