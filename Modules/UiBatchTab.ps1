#Requires -Version 5.1
<#
.SYNOPSIS
    Google Maps Routes & Map Generator — Batch Processing Tab UI Handlers.
.DESCRIPTION
    Manages multi-format dataset import and preview (JSON, CSV, Excel),
    pre-batch geocode validation pass (Feature 3.J), asynchronous batch route calculation,
    real-time log streaming, results DataGrid binding, and batch reports export (PDF, GPX, KML, Excel, CSV, JSON).
.NOTES
    Encoding: UTF-8 with BOM
#>

$script:LoadedBatchData = $null
$script:BatchResultsList = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:BatchWorkerRunning = $false

function Write-BatchLog([string]$Message, [string]$Level = 'INFO') {
    $ts = (Get-Date).ToString('HH:mm:ss')
    $tag = switch ($Level) { 'OK' { '[OK]   ' } 'WARN' { '[WARN] ' } 'ERROR' { '[ERROR]' } default { '[INFO] ' } }
    $line = "$ts $tag $Message"
    $tb = if ($script:Controls -and $script:Controls.txtBatchLog) { $script:Controls.txtBatchLog } else { $null }
    if ($tb) {
        $tb.AppendText("$line`r`n")
        $tb.ScrollToEnd()
    }
    Write-AppLog $Message $Level
}
Set-Item -Path "function:global:Write-BatchLog" -Value (Get-Item "function:Write-BatchLog").ScriptBlock -ErrorAction SilentlyContinue

function Load-BatchFilePreviewInternal([string]$Path) {
    if (-not (Test-Path $Path)) { return }
    $dgBatchInput = if ($script:Controls) { $script:Controls.dgBatchInput } else { $null }
    $lblBatchFileInfo = if ($script:Controls) { $script:Controls.lblBatchFileInfo } else { $null }

    try {
        $data = Import-RouteDataFile -Path $Path
        $script:LoadedBatchData = $data

        if ($dgBatchInput) {
            $dgBatchInput.ItemsSource = $null
            $dgBatchInput.Columns.Clear()
        }

        if ($data.Mode -eq 'SequentialStops') {
            if ($lblBatchFileInfo) {
                $lblBatchFileInfo.Text = "Format: $($data.Format) | Mode: Sequential Stops ($($data.Routes.Count) Routes, $($data.TotalCount) Stops) | Total: $($data.TotalCount) stops"
                $lblBatchFileInfo.Foreground = [System.Windows.Media.Brushes]::LightGreen
            }

            if ($dgBatchInput) {
                $colStep = [System.Windows.Controls.DataGridTextColumn]::new()
                $colStep.Header = "#"
                $colStep.Binding = [System.Windows.Data.Binding]::new("Step")
                $colStep.Width = [System.Windows.Controls.DataGridLength]::new(50)
                $colStep.IsReadOnly = $true
                $dgBatchInput.Columns.Add($colStep)

                $colRouteName = [System.Windows.Controls.DataGridTextColumn]::new()
                $colRouteName.Header = (Get-LocText 'BatchColName' 'Route Name')
                $bindRouteName = [System.Windows.Data.Binding]::new("RouteName")
                $bindRouteName.Mode = [System.Windows.Data.BindingMode]::TwoWay
                $bindRouteName.UpdateSourceTrigger = [System.Windows.Data.UpdateSourceTrigger]::PropertyChanged
                $colRouteName.Binding = $bindRouteName
                $colRouteName.Width = [System.Windows.Controls.DataGridLength]::new(200)
                $colRouteName.IsReadOnly = $false
                $dgBatchInput.Columns.Add($colRouteName)

                $colRole = [System.Windows.Controls.DataGridTextColumn]::new()
                $colRole.Header = "Role / Point Type"
                $colRole.Binding = [System.Windows.Data.Binding]::new("Role")
                $colRole.Width = [System.Windows.Controls.DataGridLength]::new(160)
                $colRole.IsReadOnly = $true
                $dgBatchInput.Columns.Add($colRole)

                $colAddr = [System.Windows.Controls.DataGridTextColumn]::new()
                $colAddr.Header = "Address / Location"
                $colAddr.Binding = [System.Windows.Data.Binding]::new("Address")
                $colAddr.Width = [System.Windows.Controls.DataGridLength]::new(320)
                $colAddr.IsReadOnly = $true
                $dgBatchInput.Columns.Add($colAddr)

                $colRaw = [System.Windows.Controls.DataGridTextColumn]::new()
                $colRaw.Header = "Source Record Data"
                $colRaw.Binding = [System.Windows.Data.Binding]::new("RawSummary")
                $colRaw.Width = [System.Windows.Controls.DataGridLength]::new(1, [System.Windows.Controls.DataGridLengthUnitType]::Star)
                $colRaw.IsReadOnly = $true
                $dgBatchInput.Columns.Add($colRaw)

                $previewItems = [System.Collections.Generic.List[PSCustomObject]]::new()
                $stopsCount = $data.Stops.Count
                for ($i = 0; $i -lt $stopsCount; $i++) {
                    $st = $data.Stops[$i]
                    $role = if ($st.Sequence -eq '1' -or $i -eq 0) { "🟢 Origin (Start)" }
                    elseif ($i -eq ($stopsCount - 1) -or ($i + 1 -lt $stopsCount -and $data.Stops[$i + 1].Sequence -eq '1')) { "🔴 Destination (End)" }
                    else { "🟡 Waypoint $i" }

                    $rawProps = @()
                    if ($st.Raw) {
                        foreach ($p in $st.Raw.PSObject.Properties) {
                            $rawProps += "$($p.Name)=$($p.Value)"
                        }
                    }
                    $previewItems.Add([PSCustomObject]@{
                        Step       = ($i + 1)
                        RouteId    = [string]$st.RouteId
                        RouteName  = [string]$st.RouteName
                        Role       = $role
                        Address    = [string]$st.Address
                        RawSummary = ($rawProps -join '; ')
                    })
                }
                $dgBatchInput.ItemsSource = $previewItems
            }
            Write-BatchLog "Loaded file: $Path ($($data.TotalCount) sequential stops, format: $($data.Format), mode: $($data.Mode))" "OK"
        }
        else {
            if ($lblBatchFileInfo) {
                $lblBatchFileInfo.Text = "Format: $($data.Format) | Mode: Route List | Total Routes: $($data.TotalCount)"
                $lblBatchFileInfo.Foreground = [System.Windows.Media.Brushes]::LightGreen
            }

            if ($dgBatchInput) {
                $colId = [System.Windows.Controls.DataGridTextColumn]::new()
                $colId.Header = "ID"
                $colId.Binding = [System.Windows.Data.Binding]::new("Id")
                $colId.Width = [System.Windows.Controls.DataGridLength]::new(50)
                $colId.IsReadOnly = $true
                $dgBatchInput.Columns.Add($colId)

                $colName = [System.Windows.Controls.DataGridTextColumn]::new()
                $colName.Header = (Get-LocText 'BatchColName' 'Route Name')
                $bindName = [System.Windows.Data.Binding]::new("Name")
                $bindName.Mode = [System.Windows.Data.BindingMode]::TwoWay
                $bindName.UpdateSourceTrigger = [System.Windows.Data.UpdateSourceTrigger]::PropertyChanged
                $colName.Binding = $bindName
                $colName.Width = [System.Windows.Controls.DataGridLength]::new(180)
                $colName.IsReadOnly = $false
                $dgBatchInput.Columns.Add($colName)

                $colStart = [System.Windows.Controls.DataGridTextColumn]::new()
                $colStart.Header = "Origin (Start)"
                $colStart.Binding = [System.Windows.Data.Binding]::new("Start")
                $colStart.Width = [System.Windows.Controls.DataGridLength]::new(200)
                $colStart.IsReadOnly = $true
                $dgBatchInput.Columns.Add($colStart)

                $colEnd = [System.Windows.Controls.DataGridTextColumn]::new()
                $colEnd.Header = "Destination (End)"
                $colEnd.Binding = [System.Windows.Data.Binding]::new("End")
                $colEnd.Width = [System.Windows.Controls.DataGridLength]::new(200)
                $colEnd.IsReadOnly = $true
                $dgBatchInput.Columns.Add($colEnd)

                $colWpCount = [System.Windows.Controls.DataGridTextColumn]::new()
                $colWpCount.Header = "Waypoints"
                $colWpCount.Binding = [System.Windows.Data.Binding]::new("WaypointCount")
                $colWpCount.Width = [System.Windows.Controls.DataGridLength]::new(80)
                $colWpCount.IsReadOnly = $true
                $dgBatchInput.Columns.Add($colWpCount)

                $colWpText = [System.Windows.Controls.DataGridTextColumn]::new()
                $colWpText.Header = "Intermediate Stops"
                $colWpText.Binding = [System.Windows.Data.Binding]::new("WaypointsText")
                $colWpText.Width = [System.Windows.Controls.DataGridLength]::new(250)
                $colWpText.IsReadOnly = $true
                $dgBatchInput.Columns.Add($colWpText)

                $colType = [System.Windows.Controls.DataGridTextColumn]::new()
                $colType.Header = "Route Type"
                $colType.Binding = [System.Windows.Data.Binding]::new("RouteType")
                $colType.Width = [System.Windows.Controls.DataGridLength]::new(100)
                $colType.IsReadOnly = $true
                $dgBatchInput.Columns.Add($colType)

                $previewItems = [System.Collections.Generic.List[PSCustomObject]]::new()
                foreach ($r in @($data.Routes)) {
                    $wpText = if ($r.Waypoints -and @($r.Waypoints).Count -gt 0) { (@($r.Waypoints) -join ' | ') } else { '(none)' }
                    $wpCount = if ($r.Waypoints) { @($r.Waypoints).Count } else { 0 }
                    $rType = if ($r.RouteType) { $r.RouteType } else { 'Default' }

                    $previewItems.Add([PSCustomObject]@{
                        Id            = [string]$r.Id
                        Name          = [string]$r.Name
                        Start         = [string]$r.Start
                        End           = [string]$r.End
                        WaypointCount = $wpCount
                        WaypointsText = $wpText
                        RouteType     = $rType
                    })
                }
                $dgBatchInput.ItemsSource = $previewItems
            }
            Write-BatchLog "Loaded file: $Path ($($data.TotalCount) routes, format: $($data.Format), mode: $($data.Mode))" "OK"
        }
    }
    catch {
        if ($lblBatchFileInfo) {
            $lblBatchFileInfo.Text = "Load error: $($_.Exception.Message)"
            $lblBatchFileInfo.Foreground = [System.Windows.Media.Brushes]::Salmon
        }
        Write-BatchLog "Load error: $($_.Exception.Message)" "ERROR"
    }
}
Set-Item -Path "function:global:Load-BatchFilePreviewInternal" -Value (Get-Item "function:Load-BatchFilePreviewInternal").ScriptBlock -ErrorAction SilentlyContinue

function Sync-BatchPreviewToRoutes {
    if (-not $script:LoadedBatchData -or -not $script:Controls -or -not $script:Controls.dgBatchInput) { return }
    $dg = $script:Controls.dgBatchInput

    try {
        $dg.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row, $true)
    }
    catch { }

    $items = @($dg.ItemsSource)
    if ($items.Count -eq 0) { return }

    if ($script:LoadedBatchData.Mode -eq 'SequentialStops') {
        # "for multipoints in many rows its enough that the name is in first row"
        foreach ($route in @($script:LoadedBatchData.Routes)) {
            $matchingStops = @($items | Where-Object { $_.RouteId -eq $route.Id })
            if ($matchingStops.Count -eq 0 -and $script:LoadedBatchData.Routes.Count -eq 1) {
                $matchingStops = $items
            }

            # Check first row first
            $resolvedName = ''
            if ($matchingStops.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($matchingStops[0].RouteName)) {
                $resolvedName = $matchingStops[0].RouteName.Trim()
            }
            # Fallback to any other row in the group if first row was blank
            if ([string]::IsNullOrWhiteSpace($resolvedName)) {
                foreach ($st in $matchingStops) {
                    if (-not [string]::IsNullOrWhiteSpace($st.RouteName)) {
                        $resolvedName = $st.RouteName.Trim()
                        break
                    }
                }
            }

            if (-not [string]::IsNullOrWhiteSpace($resolvedName)) {
                $route.Name = $resolvedName
            }
        }
    }
    else {
        # RouteList mode
        $routes = @($script:LoadedBatchData.Routes)
        for ($i = 0; $i -lt [math]::Min($routes.Count, $items.Count); $i++) {
            $gridItem = $items[$i]
            if ($gridItem -and -not [string]::IsNullOrWhiteSpace($gridItem.Name)) {
                $routes[$i].Name = $gridItem.Name.Trim()
            }
        }
    }
}
Set-Item -Path "function:global:Sync-BatchPreviewToRoutes" -Value (Get-Item "function:Sync-BatchPreviewToRoutes").ScriptBlock -ErrorAction SilentlyContinue

function Register-UiBatchTabEvents {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][hashtable]$Controls,
        [Parameter(Mandatory = $false)][System.Windows.Window]$Window = $null
    )

    $script:Controls = $Controls
    foreach ($k in $Controls.Keys) {
        Set-Variable -Name $k -Value $Controls[$k] -Scope Script -Force
        Set-Variable -Name "script:$k" -Value $Controls[$k] -Scope Script -Force
    }

    $aliases = [ordered]@{
        'txtBatchFilePath'            = $Controls.txtBatchFilePath
        'btnBrowseBatchFile'          = $Controls.btnBrowseBatchFile
        'btnReloadBatchFile'          = $Controls.btnReloadBatchFile
        'lblBatchFileInfo'            = $Controls.lblBatchFileInfo
        'cmbBatchRouteType'           = $Controls.cmbBatchRouteType
        'btnValidateBatchGeocoding'   = $Controls.btnValidateBatchGeocoding
        'btnStartBatch'               = $Controls.btnStartBatch
        'btnStopBatch'                = $Controls.btnStopBatch
        'tabBatchSub'                 = $Controls.tabBatchSub
        'tabSubValidation'            = $Controls.tabSubValidation
        'dgBatchInput'                = $Controls.dgBatchInput
        'dgBatchResults'              = $Controls.dgBatchResults
        'dgBatchPoints'               = $Controls.dgBatchPoints
        'dgGeocodeValidation'         = $Controls.dgGeocodeValidation
        'lblGeocodeValidationSummary' = $Controls.lblGeocodeValidationSummary
        'btnCopyInvalidAddresses'     = $Controls.btnCopyInvalidAddresses
        'txtBatchLog'                 = $Controls.txtBatchLog
        'pbBatchProgress'             = $Controls.pbBatchProgress
        'lblBatchProgressText'        = $Controls.lblBatchProgressText
        'lblBatchStats'               = $Controls.lblBatchStats
        'btnOpenOutputDir'            = $Controls.btnOpenOutputDir
        'btnBatchExportPdf'           = $Controls.btnBatchExportPdf
        'btnBatchExportGpx'           = $Controls.btnBatchExportGpx
        'btnBatchExportKml'           = $Controls.btnBatchExportKml
        'btnExportExcel'              = $Controls.btnExportExcel
        'btnExportCsv'                = $Controls.btnExportCsv
        'btnExportJson'               = $Controls.btnExportJson
        'lblFooter'                   = $Controls.lblFooterStatus
        'txtOutputDir'                = $Controls.txtSettingsOutputDir
        'chkRememberKey'              = $Controls.chkRememberKey
        'cmbDefaultRouteType'         = $Controls.cmbDefaultRouteType
        'cmbDefaultEmission'          = $Controls.cmbDefaultEmission
        'tabMain'                     = $Controls.tabMain
    }

    foreach ($entry in $aliases.GetEnumerator()) {
        Set-Variable -Name $entry.Key -Value $entry.Value -Scope Script -Force
        Set-Variable -Name "script:$($entry.Key)" -Value $entry.Value -Scope Script -Force
    }

    $txtBatchFilePath            = $aliases['txtBatchFilePath']
    $btnBrowseBatchFile          = $aliases['btnBrowseBatchFile']
    $btnReloadBatchFile          = $aliases['btnReloadBatchFile']
    $lblBatchFileInfo            = $aliases['lblBatchFileInfo']
    $cmbBatchRouteType           = $aliases['cmbBatchRouteType']
    $btnValidateBatchGeocoding   = $aliases['btnValidateBatchGeocoding']
    $btnStartBatch               = $aliases['btnStartBatch']
    $btnStopBatch                = $aliases['btnStopBatch']
    $tabBatchSub                 = $aliases['tabBatchSub']
    $tabSubValidation            = $aliases['tabSubValidation']
    $dgBatchInput                = $aliases['dgBatchInput']
    $dgBatchResults              = $aliases['dgBatchResults']
    $dgBatchPoints               = $aliases['dgBatchPoints']
    $dgGeocodeValidation         = $aliases['dgGeocodeValidation']
    $lblGeocodeValidationSummary = $aliases['lblGeocodeValidationSummary']
    $btnCopyInvalidAddresses     = $aliases['btnCopyInvalidAddresses']
    $txtBatchLog                 = $aliases['txtBatchLog']
    $pbBatchProgress             = $aliases['pbBatchProgress']
    $lblBatchProgressText        = $aliases['lblBatchProgressText']
    $lblBatchStats               = $aliases['lblBatchStats']
    $btnOpenOutputDir            = $aliases['btnOpenOutputDir']
    $btnBatchExportPdf           = $aliases['btnBatchExportPdf']
    $btnBatchExportGpx           = $aliases['btnBatchExportGpx']
    $btnBatchExportKml           = $aliases['btnBatchExportKml']
    $btnExportExcel              = $aliases['btnExportExcel']
    $btnExportCsv                = $aliases['btnExportCsv']
    $btnExportJson               = $aliases['btnExportJson']
    $lblFooter                   = $aliases['lblFooter']
    $txtOutputDir                = $aliases['txtOutputDir']
    $chkRememberKey              = $aliases['chkRememberKey']
    $cmbDefaultRouteType         = $aliases['cmbDefaultRouteType']
    $cmbDefaultEmission          = $aliases['cmbDefaultEmission']
    $tabMain                     = $aliases['tabMain']

    $btnBrowseBatchFile.Add_Click({
        $initDir = $null
        if (-not [string]::IsNullOrWhiteSpace($txtBatchFilePath.Text) -and (Test-Path $txtBatchFilePath.Text.Trim())) {
            $initDir = Split-Path $txtBatchFilePath.Text.Trim() -Parent
        }
        $file = Select-InputDataFile -InitialDirectory $initDir
        if ($file) {
            $txtBatchFilePath.Text = $file
            Load-BatchFilePreviewInternal -Path $file
        }
    })

    $btnReloadBatchFile.Add_Click({
        if ($txtBatchFilePath.Text) {
            Load-BatchFilePreviewInternal -Path $txtBatchFilePath.Text.Trim()
        }
    })

    if ($dgBatchInput) {
        $dgBatchInput.Add_CellEditEnding({
            Sync-BatchPreviewToRoutes
        })
        $dgBatchInput.Add_RowEditEnding({
            Sync-BatchPreviewToRoutes
        })
    }

    # Feature 3.J: Pre-Batch Geocode Validation
    $btnValidateBatchGeocoding.Add_Click({
        Sync-BatchPreviewToRoutes
        $apiKey = if (Get-Command Get-CurrentApiKey -ErrorAction SilentlyContinue) { Get-CurrentApiKey } else { $script:AppConfig.ApiKey }
        if ([string]::IsNullOrWhiteSpace($apiKey)) {
            [System.Windows.MessageBox]::Show((Get-LocText 'MsgMissingApiKeyPrompt' 'Google Maps API Key is required.'), (Get-LocText 'MsgMissingApiKeyTitle' 'Missing API Key'), 'OK', 'Warning')
            if ($tabMain) { $tabMain.SelectedIndex = 2 }
            elseif ($script:Controls -and $script:Controls.tabMain) { $script:Controls.tabMain.SelectedIndex = 2 }
            return
        }

        if ($null -eq $script:LoadedBatchData) {
            [System.Windows.MessageBox]::Show((Get-LocText 'MsgNoDataFile' 'Please load an input data file first.'), (Get-LocText 'MsgNoDataFileTitle' 'No File'), 'OK', 'Warning')
            return
        }

        # Collect distinct addresses with roles
        $addressDict = [ordered]@{}
        if ($script:LoadedBatchData.Mode -eq 'SequentialStops') {
            $stops = $script:LoadedBatchData.Stops
            for ($i = 0; $i -lt $stops.Count; $i++) {
                $addr = [string]$stops[$i].Address
                if (-not [string]::IsNullOrWhiteSpace($addr)) {
                    $role = if ($i -eq 0) { "Start (Origin)" } elseif ($i -eq ($stops.Count - 1)) { "End (Destination)" } else { "Waypoint $i" }
                    if (-not $addressDict.Contains($addr)) {
                        $addressDict[$addr] = [PSCustomObject]@{ Address = $addr; Role = $role }
                    }
                }
            }
        }
        else {
            foreach ($r in @($script:LoadedBatchData.Routes)) {
                if ($r.Start -and -not $addressDict.Contains($r.Start)) {
                    $addressDict[$r.Start] = [PSCustomObject]@{ Address = [string]$r.Start; Role = "Origin" }
                }
                if ($r.End -and -not $addressDict.Contains($r.End)) {
                    $addressDict[$r.End] = [PSCustomObject]@{ Address = [string]$r.End; Role = "Destination" }
                }
                if ($r.Waypoints) {
                    foreach ($w in @($r.Waypoints)) {
                        if ($w -and -not $addressDict.Contains($w)) {
                            $addressDict[$w] = [PSCustomObject]@{ Address = [string]$w; Role = "Waypoint" }
                        }
                    }
                }
            }
        }

        $itemsToValidate = [System.Collections.Generic.List[PSCustomObject]]::new()
        foreach ($k in $addressDict.Keys) { $itemsToValidate.Add($addressDict[$k]) }

        if ($itemsToValidate.Count -eq 0) {
            [System.Windows.MessageBox]::Show("No valid addresses found in the loaded file.", "Validation", "OK", "Information")
            return
        }

        # Switch to Validation Tab (Index 1)
        if ($script:Controls -and $script:Controls.tabBatchSub) { $script:Controls.tabBatchSub.SelectedIndex = 1 }
        elseif ($tabBatchSub) { $tabBatchSub.SelectedIndex = 1 }
        if ($script:Controls -and $script:Controls.dgGeocodeValidation) { $script:Controls.dgGeocodeValidation.ItemsSource = $null }
        if ($script:Controls -and $script:Controls.lblGeocodeValidationSummary) { $script:Controls.lblGeocodeValidationSummary.Text = "Validating $($itemsToValidate.Count) unique addresses..." }
        if ($script:Controls -and $script:Controls.btnValidateBatchGeocoding) { $script:Controls.btnValidateBatchGeocoding.IsEnabled = $false }
        if ($script:Controls -and $script:Controls.btnStartBatch) { $script:Controls.btnStartBatch.IsEnabled = $false }
        if ($script:Controls -and $script:Controls.btnStopBatch) { $script:Controls.btnStopBatch.IsEnabled = $true }

        $logQueue = [System.Collections.Concurrent.ConcurrentQueue[PSCustomObject]]::new()
        $syncState = [hashtable]::Synchronized(@{
            CancelRequested = $false
            CurrentIndex    = 0
            TotalCount      = $itemsToValidate.Count
            LogQueue        = $logQueue
        })
        $script:SyncState = $syncState

        $psCmd = New-WorkerPowerShell -ScriptBlock $script:GeocodeValidationAsync
        $langCode = if ($script:CurrentLanguage) { $script:CurrentLanguage } else { 'en' }
        $psCmd.AddArgument($itemsToValidate).AddArgument($apiKey).AddArgument($langCode).AddArgument($syncState).AddArgument($script:LogFile) | Out-Null

        try {
            $asyncHandle = $psCmd.BeginInvoke()
        }
        catch {
            if ($script:Controls -and $script:Controls.btnValidateBatchGeocoding) { $script:Controls.btnValidateBatchGeocoding.IsEnabled = $true }
            if ($script:Controls -and $script:Controls.btnStartBatch) { $script:Controls.btnStartBatch.IsEnabled = $true }
            if ($script:Controls -and $script:Controls.btnStopBatch) { $script:Controls.btnStopBatch.IsEnabled = $false }
            [System.Windows.MessageBox]::Show("Failed to launch validation: $($_.Exception.Message)", "Error", "OK", "Error")
            return
        }

        $timerVal = [System.Windows.Threading.DispatcherTimer]::new()
        $timerVal.Interval = [TimeSpan]::FromMilliseconds(150)
        $script:ActiveValidationTimer = $timerVal
        $script:ActiveValidationPs = $psCmd
        $script:ActiveValidationAsyncHandle = $asyncHandle

        $timerVal.Add_Tick({
            $localHandle = $script:ActiveValidationAsyncHandle
            $localPs = $script:ActiveValidationPs
            $localSync = $script:SyncState

            if ($localSync -and $localSync.LogQueue) {
                $logItem = $null
                while ($localSync.LogQueue.TryDequeue([ref]$logItem)) {
                    if ($logItem) { Write-BatchLog $logItem.Message $logItem.Level }
                }
            }

            if ($localSync) {
                $curr = $localSync.CurrentIndex
                $tot  = $localSync.TotalCount
                $pct  = if ($tot -gt 0) { [math]::Min(100, [math]::Round(($curr / $tot) * 100, 0)) } else { 0 }
                if ($script:Controls -and $script:Controls.pbBatchProgress) { $script:Controls.pbBatchProgress.Value = $pct }
                if ($script:Controls -and $script:Controls.lblBatchProgressText) { $script:Controls.lblBatchProgressText.Text = "Validating addresses: $curr / $tot ($pct%)" }
            }

            if ($localHandle -and $localHandle.IsCompleted) {
                if ($script:ActiveValidationTimer) { try { $script:ActiveValidationTimer.Stop() } catch { } }
                if ($script:Controls -and $script:Controls.btnValidateBatchGeocoding) { $script:Controls.btnValidateBatchGeocoding.IsEnabled = $true }
                if ($script:Controls -and $script:Controls.btnStartBatch) { $script:Controls.btnStartBatch.IsEnabled = $true }
                if ($script:Controls -and $script:Controls.btnStopBatch) { $script:Controls.btnStopBatch.IsEnabled = $false }

                try {
                    $res = $localPs.EndInvoke($localHandle)
                    $valObj = if ($res -and $res.Count -gt 0) { $res[0] } else { $null }
                    if ($valObj -and $valObj.Results) {
                        $script:LastGeocodeValidationResults = $valObj.Results
                        if ($script:Controls -and $script:Controls.dgGeocodeValidation) {
                            $script:Controls.dgGeocodeValidation.ItemsSource = $valObj.Results
                        }

                        if ($valObj.ApiUsage) {
                            Update-ApiUsageRecord -GeocodingInc $valObj.ApiUsage.Geocoding
                            Update-ApiUsageBadgeText
                        }

                        $okCount = (@($valObj.Results) | Where-Object { $_.IsOk }).Count
                        $rooftopCount = (@($valObj.Results) | Where-Object { $_.IsRooftop }).Count
                        $approxCount = (@($valObj.Results) | Where-Object { $_.Precision -like '*APPROXIMATE*' -or $_.Precision -like '*CENTER*' }).Count
                        $failCount = (@($valObj.Results) | Where-Object { -not $_.IsOk }).Count

                        if ($script:Controls -and $script:Controls.lblGeocodeValidationSummary) {
                            $script:Controls.lblGeocodeValidationSummary.Text = "Validation complete: $okCount / $($valObj.Results.Count) verified ($rooftopCount Exact/Rooftop, $approxCount Approximate, $failCount Failed)"
                        }
                        if ($failCount -gt 0 -or $approxCount -gt 0) {
                            if ($script:Controls -and $script:Controls.btnCopyInvalidAddresses) { $script:Controls.btnCopyInvalidAddresses.IsEnabled = $true }
                        }
                        if ($script:Controls -and $script:Controls.lblBatchProgressText) {
                            $script:Controls.lblBatchProgressText.Text = "Validation complete ($failCount failed, $approxCount approximate)"
                        }
                        if ($script:Controls -and $script:Controls.tabBatchSub) {
                            $script:Controls.tabBatchSub.SelectedIndex = 1
                        }
                    }
                }
                catch {
                    Write-BatchLog "Validation read error: $($_.Exception.Message)" "ERROR"
                    [System.Windows.MessageBox]::Show("Validation read error: $($_.Exception.Message)", "Error", "OK", "Error")
                }
                finally {
                    if ($localPs) { try { $localPs.Dispose() } catch { } }
                }
            }
        })
        $timerVal.Start()
    })

    $btnCopyInvalidAddresses.Add_Click({
        if ($script:LastGeocodeValidationResults) {
            $invalids = @($script:LastGeocodeValidationResults) | Where-Object { -not $_.IsOk -or $_.Precision -like '*APPROXIMATE*' }
            if ($invalids.Count -gt 0) {
                $lines = $invalids | ForEach-Object { "$($_.Role): $($_.Address) -> [$($_.Status)] $($_.Precision)" }
                [System.Windows.Clipboard]::SetText(($lines -join "`r`n"))
                [System.Windows.MessageBox]::Show("Copied $($invalids.Count) invalid/approximate addresses to clipboard.", "Copied", "OK", "Information")
            }
        }
    })

    # Start Batch Processing
    $btnStartBatch.Add_Click({
        Sync-BatchPreviewToRoutes
        $apiKey = if (Get-Command Get-CurrentApiKey -ErrorAction SilentlyContinue) { Get-CurrentApiKey } else { $script:AppConfig.ApiKey }
        if ([string]::IsNullOrWhiteSpace($apiKey)) {
            [System.Windows.MessageBox]::Show((Get-LocText 'MsgMissingApiKeyPrompt'), (Get-LocText 'MsgMissingApiKeyTitle'), 'OK', 'Warning')
            if ($tabMain) { $tabMain.SelectedIndex = 2 }
            elseif ($script:Controls -and $script:Controls.tabMain) { $script:Controls.tabMain.SelectedIndex = 2 }
            return
        }

        if ($null -eq $script:LoadedBatchData -or $script:LoadedBatchData.Routes.Count -eq 0) {
            [System.Windows.MessageBox]::Show((Get-LocText 'MsgNoDataFile'), (Get-LocText 'MsgNoDataFileTitle'), 'OK', 'Warning')
            return
        }

        $outDir = $txtOutputDir.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($outDir)) { $outDir = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'TrasyGoogleMaps' }
        if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

        $defaultRouteType = if ($cmbBatchRouteType -and $cmbBatchRouteType.SelectedItem) { ($cmbBatchRouteType.SelectedItem.Tag -as [string]) } else { 'Fastest' }
        if ([string]::IsNullOrWhiteSpace($defaultRouteType)) { $defaultRouteType = 'Fastest' }

        $script:BatchWorkerRunning = $true
        $btnStartBatch.IsEnabled = $false
        $btnStopBatch.IsEnabled = $true
        $btnBrowseBatchFile.IsEnabled = $false
        $btnReloadBatchFile.IsEnabled = $false
        $btnValidateBatchGeocoding.IsEnabled = $false

        $script:BatchResultsList.Clear()
        $dgBatchResults.ItemsSource = $null
        if ($dgBatchPoints) { $dgBatchPoints.ItemsSource = $null }
        $pbBatchProgress.Value = 0
        $lblBatchProgressText.Text = "Starting batch processing (0 / $($script:LoadedBatchData.Routes.Count))..."
        $lblBatchStats.Text = "Success: 0 | Errors: 0"

        Write-BatchLog "=== Starting batch processing ($($script:LoadedBatchData.Routes.Count) routes) ===" "INFO"
        if ($tabBatchSub) { $tabBatchSub.SelectedIndex = 4 } # Switch to Activity Log

        $logQueue = [System.Collections.Concurrent.ConcurrentQueue[PSCustomObject]]::new()
        $syncState = [hashtable]::Synchronized(@{
            CancelRequested = $false
            CurrentIndex    = 0
            TotalCount      = $script:LoadedBatchData.Routes.Count
            SuccessCount    = 0
            FailCount       = 0
            LogQueue        = $logQueue
        })
        $script:SyncState = $syncState
        $routesToProcess = @($script:LoadedBatchData.Routes)

        try {
            $psCmdBatch = New-WorkerPowerShell -ScriptBlock $script:BatchCalcAsync
            $overlayCfgJson = ((Get-CurrentOverlayConfig -Controls $Controls) | ConvertTo-Json -Depth 6 -Compress)
            $langCode = if ($script:CurrentLanguage) { $script:CurrentLanguage } else { 'en' }

            $defAvoidTolls = if ($script:AppConfig) { [bool]$script:AppConfig.AvoidTolls } else { $false }
            $defAvoidHighways = if ($script:AppConfig) { [bool]$script:AppConfig.AvoidHighways } else { $false }
            $defAvoidFerries = if ($script:AppConfig) { [bool]$script:AppConfig.AvoidFerries } else { $false }

            $psCmdBatch.AddArgument($routesToProcess).AddArgument($apiKey).AddArgument($outDir).AddArgument($defaultRouteType).AddArgument($syncState).AddArgument($script:LogFile).AddArgument($langCode).AddArgument($overlayCfgJson).AddArgument($defAvoidTolls).AddArgument($defAvoidHighways).AddArgument($defAvoidFerries) | Out-Null
            $asyncBatchHandle = $psCmdBatch.BeginInvoke()
        }
        catch {
            Write-BatchLog "CRITICAL: Could not start batch worker: $($_.Exception.Message)" "ERROR"
            $btnStartBatch.IsEnabled = $true
            $btnStopBatch.IsEnabled = $false
            $btnBrowseBatchFile.IsEnabled = $true
            $btnReloadBatchFile.IsEnabled = $true
            $btnValidateBatchGeocoding.IsEnabled = $true
            $script:BatchWorkerRunning = $false
            return
        }

        $timerBatch = [System.Windows.Threading.DispatcherTimer]::new()
        $timerBatch.Interval = [TimeSpan]::FromMilliseconds(150)
        $script:ActiveBatchTimer = $timerBatch
        $script:ActiveBatchPs = $psCmdBatch
        $script:ActiveBatchAsyncHandle = $asyncBatchHandle

        $timerBatch.Add_Tick({
            $localBatchHandle = $script:ActiveBatchAsyncHandle
            $localBatchPs = $script:ActiveBatchPs
            $localSyncState = $script:SyncState

            if (-not $localSyncState) { return }

            if ($localSyncState.LogQueue) {
                $logItem = $null
                while ($localSyncState.LogQueue.TryDequeue([ref]$logItem)) {
                    if ($logItem) { Write-BatchLog $logItem.Message $logItem.Level }
                }
            }

            $curr = $localSyncState.CurrentIndex
            $tot  = $localSyncState.TotalCount
            $pct  = if ($tot -gt 0) { [math]::Min(100, [math]::Round(($curr / $tot) * 100, 0)) } else { 0 }
            $pbBatchProgress.Value = $pct
            $lblBatchProgressText.Text = "Processing: $curr / $tot ($pct%)"

            if ($localBatchHandle -and $localBatchHandle.IsCompleted) {
                $script:ActiveBatchTimer.Stop()
                $btnStartBatch.IsEnabled = $true
                $btnStopBatch.IsEnabled = $false
                $btnBrowseBatchFile.IsEnabled = $true
                $btnReloadBatchFile.IsEnabled = $true
                $btnValidateBatchGeocoding.IsEnabled = $true
                $script:BatchWorkerRunning = $false

                if ($localSyncState.LogQueue) {
                    $logItem = $null
                    while ($localSyncState.LogQueue.TryDequeue([ref]$logItem)) {
                        if ($logItem) { Write-BatchLog $logItem.Message $logItem.Level }
                    }
                }

                foreach ($streamErr in $localBatchPs.Streams.Error) {
                    Write-BatchLog "[Worker Stream Error] $($streamErr.Exception.Message)" "ERROR"
                }

                try {
                    $res = $localBatchPs.EndInvoke($localBatchHandle)
                    $batchObj = $res[0]
                    $script:BatchResultsList = [System.Collections.Generic.List[PSCustomObject]]::new()
                    foreach ($item in @($batchObj.Results)) { $script:BatchResultsList.Add($item) }
                    $dgBatchResults.ItemsSource = $script:BatchResultsList

                    if ($batchObj.ApiUsage) {
                        Update-ApiUsageRecord -GeocodingInc $batchObj.ApiUsage.Geocoding -RoutesInc $batchObj.ApiUsage.Routes -StaticMapsInc $batchObj.ApiUsage.StaticMaps
                        Update-ApiUsageBadgeText
                    }

                    # Populate Points Detail table
                    $allPointsList = [System.Collections.Generic.List[PSCustomObject]]::new()
                    foreach ($item in @($script:BatchResultsList)) {
                        if ($item.RoutePoints -and @($item.RoutePoints).Count -gt 0) {
                            foreach ($pt in @($item.RoutePoints)) {
                                $allPointsList.Add([PSCustomObject]@{
                                    RouteId         = $item.Id
                                    RouteName       = $item.Name
                                    PointOrder      = $pt.Order
                                    PointType       = $pt.PointType
                                    OriginalAddress = $pt.OriginalAddress
                                    GeocodedAddress = $pt.GeocodedAddress
                                    GeocodeStatus   = $pt.GeocodeStatus
                                    MatchType       = $pt.MatchType
                                    IsFallback      = if ($pt.IsFallback) { 'YES' } else { 'No' }
                                    Latitude        = $pt.Latitude
                                    Longitude       = $pt.Longitude
                                })
                            }
                        }
                    }
                    if ($dgBatchPoints) { $dgBatchPoints.ItemsSource = $allPointsList }

                    $statusMsg = if ($localSyncState.CancelRequested) { 'Stopped by user.' } else { 'Completed successfully.' }
                    $lblBatchProgressText.Text = $statusMsg
                    $lblFooter.Text = "Processing complete: $($script:BatchResultsList.Count) routes processed."

                    if ($script:BatchResultsList.Count -gt 0 -and $tabBatchSub) {
                        $tabBatchSub.SelectedIndex = 2 # Switch to Calculation Results
                    }
                }
                catch {
                    Write-BatchLog "Error finalizing batch results: $($_.Exception.Message)" "ERROR"
                }
                finally {
                    $localBatchPs.Dispose()
                }
            }
        })
        $timerBatch.Start()
    })

    $btnStopBatch.Add_Click({
        if ($script:SyncState) {
            $script:SyncState.CancelRequested = $true
            Write-BatchLog "Cancellation requested by user..." "WARN"
        }
    })

    # Double click on results row to open map PNG
    $dgBatchResults.Add_MouseDoubleClick({
        if ($dgBatchResults.SelectedItem) {
            $selected = $dgBatchResults.SelectedItem
            if ($selected.MapPath -and (Test-Path $selected.MapPath)) {
                Start-Process $selected.MapPath
            }
        }
    })

    $btnOpenOutputDir.Add_Click({
        $outDir = $txtOutputDir.Text.Trim()
        if (-not [string]::IsNullOrWhiteSpace($outDir) -and (Test-Path $outDir)) {
            Start-Process explorer.exe $outDir
        }
    })

    # Feature 5.O: Export Batch PDF Report
    $btnBatchExportPdf.Add_Click({
        if ($script:BatchResultsList.Count -eq 0) {
            [System.Windows.MessageBox]::Show("No calculation results available to export.", "Export PDF", "OK", "Information")
            return
        }
        $dlg = [System.Windows.Forms.SaveFileDialog]::new()
        $dlg.Title = 'Export Batch PDF Report'
        $dlg.Filter = 'PDF Document (*.pdf)|*.pdf'
        $dlg.FileName = "Batch_Routes_Report_$((Get-Date).ToString('yyyyMMdd_HHmm')).pdf"
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $srcLeaf = if (-not [string]::IsNullOrWhiteSpace($txtBatchFilePath.Text)) { Split-Path $txtBatchFilePath.Text.Trim() -Leaf } else { 'Batch_Input' }
                Export-BatchPdfReport -OutputPath $dlg.FileName -Routes $script:BatchResultsList -SourceFileName $srcLeaf
                $ask = [System.Windows.MessageBox]::Show("Batch PDF report created successfully!`n$($dlg.FileName)`n`nDo you want to open it now?", "PDF Export", "YesNo", "Information")
                if ($ask -eq [System.Windows.MessageBoxResult]::Yes) { Start-Process $dlg.FileName }
            }
            catch {
                [System.Windows.MessageBox]::Show("Failed to export batch PDF report:`n$($_.Exception.Message)", "Error", "OK", "Error")
            }
        }
    })

    # Feature 5.P: Export Batch GPX
    $btnBatchExportGpx.Add_Click({
        $validRoutes = @($script:BatchResultsList) | Where-Object { $_.EncodedPolyline }
        if ($validRoutes.Count -eq 0) {
            [System.Windows.MessageBox]::Show("No routes with valid polyline data found.", "Export GPX", "OK", "Information")
            return
        }
        $fbd = [System.Windows.Forms.FolderBrowserDialog]::new()
        $fbd.Description = 'Select Destination Folder for GPX Tracks'
        if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $exportDir = $fbd.SelectedPath
            $count = 0
            foreach ($r in $validRoutes) {
                $cleanName = ($r.Name -replace '[\\/:*?"<>|]', '_').Trim()
                $gpxFile = Join-Path $exportDir "Route_$($r.Id)_${cleanName}.gpx"
                Export-RouteGpx -OutputPath $gpxFile -RouteName $r.Name -EncodedPolyline $r.EncodedPolyline -DistanceKm $r.DistanceKm -DurationMin $r.DurationMin | Out-Null
                $count++
            }
            [System.Windows.MessageBox]::Show("Exported $count GPX track files to:`n$exportDir", "GPX Export", "OK", "Information")
        }
    })

    # Feature 5.P: Export Batch KML
    $btnBatchExportKml.Add_Click({
        $validRoutes = @($script:BatchResultsList) | Where-Object { $_.EncodedPolyline }
        if ($validRoutes.Count -eq 0) {
            [System.Windows.MessageBox]::Show("No routes with valid polyline data found.", "Export KML", "OK", "Information")
            return
        }
        $fbd = [System.Windows.Forms.FolderBrowserDialog]::new()
        $fbd.Description = 'Select Destination Folder for KML Tracks'
        if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $exportDir = $fbd.SelectedPath
            $count = 0
            foreach ($r in $validRoutes) {
                $cleanName = ($r.Name -replace '[\\/:*?"<>|]', '_').Trim()
                $kmlFile = Join-Path $exportDir "Route_$($r.Id)_${cleanName}.kml"
                Export-RouteKml -OutputPath $kmlFile -RouteName $r.Name -EncodedPolyline $r.EncodedPolyline -DistanceKm $r.DistanceKm -DurationMin $r.DurationMin | Out-Null
                $count++
            }
            [System.Windows.MessageBox]::Show("Exported $count KML track files to:`n$exportDir", "KML Export", "OK", "Information")
        }
    })

    # Existing Standard Exports (Excel, CSV, JSON)
    $btnExportExcel.Add_Click({
        if ($script:BatchResultsList.Count -eq 0) { return }
        $dlg = [System.Windows.Forms.SaveFileDialog]::new()
        $dlg.Filter = 'Excel Workbook (*.xlsx)|*.xlsx'
        $dlg.FileName = "Batch_Results_$((Get-Date).ToString('yyyyMMdd_HHmm')).xlsx"
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Export-RouteResults -Results $script:BatchResultsList -OutputPath $dlg.FileName -Format 'Excel'
            [System.Windows.MessageBox]::Show("Exported to $($dlg.FileName)", "Excel Export", "OK", "Information")
        }
    })

    $btnExportCsv.Add_Click({
        if ($script:BatchResultsList.Count -eq 0) { return }
        $dlg = [System.Windows.Forms.SaveFileDialog]::new()
        $dlg.Filter = 'CSV File (*.csv)|*.csv'
        $dlg.FileName = "Batch_Results_$((Get-Date).ToString('yyyyMMdd_HHmm')).csv"
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Export-RouteResults -Results $script:BatchResultsList -OutputPath $dlg.FileName -Format 'CSV'
            [System.Windows.MessageBox]::Show("Exported to $($dlg.FileName)", "CSV Export", "OK", "Information")
        }
    })

    $btnExportJson.Add_Click({
        if ($script:BatchResultsList.Count -eq 0) { return }
        $dlg = [System.Windows.Forms.SaveFileDialog]::new()
        $dlg.Filter = 'JSON File (*.json)|*.json'
        $dlg.FileName = "Batch_Results_$((Get-Date).ToString('yyyyMMdd_HHmm')).json"
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Export-RouteResults -Results $script:BatchResultsList -OutputPath $dlg.FileName -Format 'JSON'
            [System.Windows.MessageBox]::Show("Exported to $($dlg.FileName)", "JSON Export", "OK", "Information")
        }
    })
}
