#Requires -Version 5.1
<#
.SYNOPSIS
    Google Maps Routes & Map Generator — Manual Route Tab UI Handlers.
.DESCRIPTION
    Manages waypoint list operations, route optimization settings, avoid options,
    manual route calculation orchestration, WebView2/PNG map preview switching,
    and single-route exports (PDF, GPX, KML).
.NOTES
    Encoding: UTF-8 with BOM
#>

$script:LastManualResult = $null
$script:ActiveManualRouteName = ''

function Update-MapViewMode {
    [CmdletBinding()]
    param()

    $rbInteractive  = if ($script:Controls) { $script:Controls.rbViewInteractive } else { $null }
    $imgPreview     = if ($script:Controls) { $script:Controls.imgMapPreview } else { $null }
    $pnlHost        = if ($script:Controls) { $script:Controls.pnlInteractiveMapHost } else { $null }
    $lblPlaceholder = if ($script:Controls) { $script:Controls.lblMapPlaceholder } else { $null }

    if (-not $script:LastManualResult) {
        if ($lblPlaceholder) { $lblPlaceholder.Visibility = [System.Windows.Visibility]::Visible }
        if ($imgPreview)     { $imgPreview.Visibility     = [System.Windows.Visibility]::Collapsed }
        if ($pnlHost)        { $pnlHost.Visibility        = [System.Windows.Visibility]::Collapsed }
        return
    }

    if ($lblPlaceholder) { $lblPlaceholder.Visibility = [System.Windows.Visibility]::Collapsed }

    if ($rbInteractive -and $rbInteractive.IsChecked) {
        if ($imgPreview) { $imgPreview.Visibility = [System.Windows.Visibility]::Collapsed }
        if ($pnlHost)    { $pnlHost.Visibility    = [System.Windows.Visibility]::Visible }
    } else {
        if ($pnlHost)    { $pnlHost.Visibility    = [System.Windows.Visibility]::Collapsed }
        if ($imgPreview) { $imgPreview.Visibility = [System.Windows.Visibility]::Visible }
    }
}
Set-Item -Path "function:global:Update-MapViewMode" -Value (Get-Item "function:Update-MapViewMode").ScriptBlock -ErrorAction SilentlyContinue

function Register-UiManualTabEvents {
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
        'btnClearStart'      = $Controls.btnClearManualStart
        'btnClearEnd'        = $Controls.btnClearManualEnd
        'txtStart'           = $Controls.txtManualStart
        'txtEnd'             = $Controls.txtManualEnd
        'txtNewWp'           = $Controls.txtNewWaypoint
        'btnAddWp'           = $Controls.btnAddWaypoint
        'lstWp'              = $Controls.lstWaypoints
        'btnWpUp'            = $Controls.btnWpUp
        'btnWpDown'          = $Controls.btnWpDown
        'btnWpRemove'        = $Controls.btnWpRemove
        'btnWpClear'         = $Controls.btnWpClear
        'txtName'            = $Controls.txtManualName
        'rbFastest'          = $Controls.rbTypeFastest
        'rbShortest'         = $Controls.rbTypeShortest
        'rbEco'              = $Controls.rbTypeEco
        'pnlEmission'        = $Controls.pnlEmission
        'cmbEmission'        = $Controls.cmbEmission
        'chkTrafficAware'    = $Controls.chkTrafficAware
        'chkAvoidTolls'      = $Controls.chkManualAvoidTolls
        'chkAvoidHighways'   = $Controls.chkManualAvoidHighways
        'chkAvoidFerries'    = $Controls.chkManualAvoidFerries
        'btnCalculate'       = $Controls.btnCalculateManual
        'lblDist'            = $Controls.lblManualDist
        'lblTime'            = $Controls.lblManualTime
        'lblType'            = $Controls.lblManualType
        'lblStatus'          = $Controls.lblManualStatus
        'lblPlaceholder'     = $Controls.lblMapPlaceholder
        'imgPreview'         = $Controls.imgMapPreview
        'pnlInteractiveHost' = $Controls.pnlInteractiveMapHost
        'rbViewInteractive'  = $Controls.rbViewInteractive
        'rbViewStatic'       = $Controls.rbViewStatic
        'lblUrlDisplay'      = $Controls.lblGoogleUrlDisplay
        'btnOpenMaps'        = $Controls.btnOpenGoogleMaps
        'btnCopyUrl'         = $Controls.btnCopyUrl
        'btnSaveMapAs'       = $Controls.btnSaveMapAs
        'btnExportPdf'       = $Controls.btnManualExportPdf
        'btnExportGpx'       = $Controls.btnManualExportGpx
        'btnExportKml'       = $Controls.btnManualExportKml
        'lblFooter'          = $Controls.lblFooterStatus
        'txtOutputDir'       = $Controls.txtSettingsOutputDir
        'tabMain'            = $Controls.tabMain
    }

    foreach ($entry in $aliases.GetEnumerator()) {
        Set-Variable -Name $entry.Key -Value $entry.Value -Scope Script -Force
        Set-Variable -Name "script:$($entry.Key)" -Value $entry.Value -Scope Script -Force
    }

    $btnClearStart      = $aliases['btnClearStart']
    $btnClearEnd        = $aliases['btnClearEnd']
    $txtStart           = $aliases['txtStart']
    $txtEnd             = $aliases['txtEnd']
    $txtNewWp           = $aliases['txtNewWp']
    $btnAddWp           = $aliases['btnAddWp']
    $lstWp              = $aliases['lstWp']
    $btnWpUp            = $aliases['btnWpUp']
    $btnWpDown          = $aliases['btnWpDown']
    $btnWpRemove        = $aliases['btnWpRemove']
    $btnWpClear         = $aliases['btnWpClear']
    $txtName            = $aliases['txtName']
    $rbFastest          = $aliases['rbFastest']
    $rbShortest         = $aliases['rbShortest']
    $rbEco              = $aliases['rbEco']
    $pnlEmission        = $aliases['pnlEmission']
    $cmbEmission        = $aliases['cmbEmission']
    $chkTrafficAware    = $aliases['chkTrafficAware']
    $chkAvoidTolls      = $aliases['chkAvoidTolls']
    $chkAvoidHighways   = $aliases['chkAvoidHighways']
    $chkAvoidFerries    = $aliases['chkAvoidFerries']
    $btnCalculate       = $aliases['btnCalculate']
    $lblDist            = $aliases['lblDist']
    $lblTime            = $aliases['lblTime']
    $lblType            = $aliases['lblType']
    $lblStatus          = $aliases['lblStatus']
    $lblPlaceholder     = $aliases['lblPlaceholder']
    $imgPreview         = $aliases['imgPreview']
    $pnlInteractiveHost = $aliases['pnlInteractiveHost']
    $rbViewInteractive  = $aliases['rbViewInteractive']
    $rbViewStatic       = $aliases['rbViewStatic']
    $lblUrlDisplay      = $aliases['lblUrlDisplay']
    $btnOpenMaps        = $aliases['btnOpenMaps']
    $btnCopyUrl         = $aliases['btnCopyUrl']
    $btnSaveMapAs       = $aliases['btnSaveMapAs']
    $btnExportPdf       = $aliases['btnExportPdf']
    $btnExportGpx       = $aliases['btnExportGpx']
    $btnExportKml       = $aliases['btnExportKml']
    $lblFooter          = $aliases['lblFooter']
    $txtOutputDir       = $aliases['txtOutputDir']
    $tabMain            = $aliases['tabMain']

    # Clear address inputs
    $btnClearStart.Add_Click({ $txtStart.Clear() })
    $btnClearEnd.Add_Click({ $txtEnd.Clear() })

    # Waypoints list management
    $btnAddWp.Add_Click({
        $wp = $txtNewWp.Text.Trim()
        if (-not [string]::IsNullOrWhiteSpace($wp)) {
            if ($lstWp.Items.Count -ge 25) {
                [System.Windows.MessageBox]::Show((Get-LocText 'MsgMaxWaypoints' 'Maximum 25 waypoints allowed.'), (Get-LocText 'MsgMaxWaypointsTitle' 'Waypoints Limit'), 'OK', 'Warning')
                return
            }
            $null = $lstWp.Items.Add($wp)
            $txtNewWp.Clear()
        }
    })

    $txtNewWp.Add_KeyDown({
        if ($_.Key -eq [System.Windows.Input.Key]::Enter) {
            $btnAddWp.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        }
    })

    $btnWpRemove.Add_Click({
        if ($lstWp.SelectedIndex -ge 0) { $lstWp.Items.RemoveAt($lstWp.SelectedIndex) }
    })

    $btnWpClear.Add_Click({ $lstWp.Items.Clear() })

    $btnWpUp.Add_Click({
        $idx = $lstWp.SelectedIndex
        if ($idx -gt 0) {
            $item = $lstWp.Items[$idx]
            $lstWp.Items.RemoveAt($idx)
            $lstWp.Items.Insert($idx - 1, $item)
            $lstWp.SelectedIndex = $idx - 1
        }
    })

    $btnWpDown.Add_Click({
        $idx = $lstWp.SelectedIndex
        if ($idx -ge 0 -and $idx -lt ($lstWp.Items.Count - 1)) {
            $item = $lstWp.Items[$idx]
            $lstWp.Items.RemoveAt($idx)
            $lstWp.Items.Insert($idx + 1, $item)
            $lstWp.SelectedIndex = $idx + 1
        }
    })

    # Optimization mode toggles
    $rbEco.Add_Checked({ $pnlEmission.Visibility = [System.Windows.Visibility]::Visible })
    $rbFastest.Add_Checked({ $pnlEmission.Visibility = [System.Windows.Visibility]::Collapsed })
    $rbShortest.Add_Checked({ $pnlEmission.Visibility = [System.Windows.Visibility]::Collapsed })

    # Interactive vs Static Map Mode Toggle (Feature 4.K)
    $rbViewInteractive.Add_Checked({ Update-MapViewMode })
    $rbViewStatic.Add_Checked({ Update-MapViewMode })

    # Calculate Route Handler
    $btnCalculate.Add_Click({
        $apiKey = if (Get-Command Get-CurrentApiKey -ErrorAction SilentlyContinue) { Get-CurrentApiKey } else { $script:AppConfig.ApiKey }
        if ([string]::IsNullOrWhiteSpace($apiKey)) {
            [System.Windows.MessageBox]::Show((Get-LocText 'MsgMissingApiKeyPrompt' 'Google Maps API Key is required.'), (Get-LocText 'MsgMissingApiKeyTitle' 'Missing API Key'), 'OK', 'Warning')
            if ($tabMain) { $tabMain.SelectedIndex = 2 }
            elseif ($script:Controls -and $script:Controls.tabMain) { $script:Controls.tabMain.SelectedIndex = 2 }
            return
        }

        $start = $txtStart.Text.Trim()
        $end = $txtEnd.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($start) -or [string]::IsNullOrWhiteSpace($end)) {
            [System.Windows.MessageBox]::Show((Get-LocText 'MsgMissingData' 'Please provide both Origin and Destination.'), (Get-LocText 'MsgMissingDataTitle' 'Missing Data'), 'OK', 'Warning')
            return
        }

        $waypoints = [System.Collections.Generic.List[string]]::new()
        foreach ($item in $lstWp.Items) { $waypoints.Add([string]$item) }

        $routeType = if ($rbShortest.IsChecked) { 'Shortest' } elseif ($rbEco.IsChecked) { 'Eco' } else { 'Fastest' }
        $emission = if ($cmbEmission -and $cmbEmission.SelectedItem) { ($cmbEmission.SelectedItem.Tag -as [string]) } else { 'GASOLINE' }
        if ([string]::IsNullOrWhiteSpace($emission)) { $emission = 'GASOLINE' }
        $trafficAware = [bool]$chkTrafficAware.IsChecked
        $avoidTolls = [bool]$chkAvoidTolls.IsChecked
        $avoidHighways = [bool]$chkAvoidHighways.IsChecked
        $avoidFerries = [bool]$chkAvoidFerries.IsChecked

        $name = $txtName.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($name)) { $name = "Route $start -> $end" }
        $script:ActiveManualRouteName = $name

        $outDir = $txtOutputDir.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($outDir)) { $outDir = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'TrasyGoogleMaps' }
        if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

        $btnCalculate.IsEnabled = $false
        $btnCalculate.Content = '⏳ CALCULATING ROUTE...'
        $lblStatus.Text = 'Geocoding and calculating...'
        $lblStatus.Foreground = [System.Windows.Media.Brushes]::SkyBlue
        $lblFooter.Text = 'Calculating manual route...'

        Write-AppLog "Started manual route calculation: '$start' -> '$end' (Waypoints: $($waypoints.Count), Type: $routeType, Engine: $emission, AvoidTolls: $avoidTolls, AvoidHighways: $avoidHighways, AvoidFerries: $avoidFerries)..." "INFO"

        $psCmd = New-WorkerPowerShell -ScriptBlock $script:ManualCalcAsync
        $overlayCfgJson = ((Get-CurrentOverlayConfig -Controls $Controls) | ConvertTo-Json -Depth 6 -Compress)
        $langCode = if ($script:CurrentLanguage) { $script:CurrentLanguage } else { 'en' }

        $psCmd.AddArgument($start).AddArgument($end).AddArgument($waypoints).AddArgument($routeType).AddArgument($emission).AddArgument($trafficAware).AddArgument($name).AddArgument($apiKey).AddArgument($outDir).AddArgument($script:LogFile).AddArgument($langCode).AddArgument($overlayCfgJson).AddArgument($avoidTolls).AddArgument($avoidHighways).AddArgument($avoidFerries) | Out-Null

        try {
            $asyncHandle = $psCmd.BeginInvoke()
        }
        catch {
            Write-AppLog "CRITICAL: BeginInvoke() failed: $($_.Exception.Message)" "ERROR"
            $btnCalculate.IsEnabled = $true
            $btnCalculate.Content = '🚀 CALCULATE ROUTE & DOWNLOAD MAP'
            $lblStatus.Text = '✕ Launch error'
            $lblStatus.Foreground = [System.Windows.Media.Brushes]::Salmon
            $lblFooter.Text = "Error: $($_.Exception.Message)"
            return
        }

        $timer = [System.Windows.Threading.DispatcherTimer]::new()
        $timer.Interval = [TimeSpan]::FromMilliseconds(150)
        $script:ActiveManualTimer = $timer
        $script:ActiveManualPs = $psCmd
        $script:ActiveManualAsyncHandle = $asyncHandle
        $script:ManualTimerTicks = 0

        $timer.Add_Tick({
            $localHandle = $script:ActiveManualAsyncHandle
            $localPs = $script:ActiveManualPs
            $script:ManualTimerTicks++

            if ($localHandle -and $localHandle.IsCompleted) {
                if ($script:ActiveManualTimer) { try { $script:ActiveManualTimer.Stop() } catch { } }
                $btnCalculate.IsEnabled = $true
                $btnCalculate.Content = '🚀 CALCULATE ROUTE & DOWNLOAD MAP'

                foreach ($streamErr in $localPs.Streams.Error) {
                    Write-AppLog "[Stream.Error] $($streamErr.Exception.Message) @ $($streamErr.InvocationInfo.PositionMessage)" "ERROR"
                }

                try {
                    $res = $localPs.EndInvoke($localHandle)
                    $calc = $res[0]
                    if ($calc.Success) {
                        $script:LastManualResult = $calc
                        $lblDist.Text = "$($calc.DistanceKm) km"
                        $lblTime.Text = "$($calc.DurationMin) min"
                        $lblType.Text = switch ($script:CurrentLanguage) {
                            'de' { if ($calc.RouteType -eq 'Fastest') { 'Schnellste' } elseif ($calc.RouteType -eq 'Shortest') { 'Kürzeste' } else { 'Eco' } }
                            'pl' { if ($calc.RouteType -eq 'Fastest') { 'Najszybsza' } elseif ($calc.RouteType -eq 'Shortest') { 'Najkrótsza' } else { 'Eko' } }
                            default { [string]$calc.RouteType }
                        }
                        $lblStatus.Text = '✓ Success'
                        $lblStatus.Foreground = [System.Windows.Media.Brushes]::LightGreen
                        $lblFooter.Text = "Route ready: $($calc.DistanceKm) km, $($calc.DurationMin) min"

                        if ($calc.ApiUsage) {
                            Update-ApiUsageRecord -GeocodingInc $calc.ApiUsage.Geocoding -RoutesInc $calc.ApiUsage.Routes -StaticMapsInc $calc.ApiUsage.StaticMaps
                            Update-ApiUsageBadgeText
                        }

                        $script:LastGoogleMapsUrl = $calc.GoogleMapsUrl
                        $lblUrlDisplay.Text = $calc.GoogleMapsUrl
                        $btnOpenMaps.IsEnabled = $true
                        $btnCopyUrl.IsEnabled = $true
                        $btnExportPdf.IsEnabled = $true
                        $btnExportGpx.IsEnabled = $true
                        $btnExportKml.IsEnabled = $true

                        $lblPlaceholder.Visibility = [System.Windows.Visibility]::Collapsed

                        # Load Static Map PNG
                        if ($calc.MapPath -and (Test-Path $calc.MapPath)) {
                            $script:LastGeneratedMapPath = $calc.MapPath
                            $btnSaveMapAs.IsEnabled = $true

                            $imgBytes = [System.IO.File]::ReadAllBytes($calc.MapPath)
                            $ms = [System.IO.MemoryStream]::new($imgBytes)
                            $bi = [System.Windows.Media.Imaging.BitmapImage]::new()
                            $bi.BeginInit()
                            $bi.StreamSource = $ms
                            $bi.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
                            $bi.EndInit()
                            $bi.Freeze()
                            $imgPreview.Source = $bi
                        }

                        # Load Interactive Map (Feature 4.K)
                        if ($calc.EncodedPolyline) {
                            Update-InteractiveRouteMap -CalcResult $calc
                        }

                        Update-MapViewMode
                    }
                    else {
                        $lblStatus.Text = '✕ Error'
                        $lblStatus.Foreground = [System.Windows.Media.Brushes]::Salmon
                        $lblFooter.Text = "Error: $($calc.Error)"
                        [System.Windows.MessageBox]::Show($calc.Error, 'Route Error', 'OK', 'Error')
                    }
                }
                catch {
                    $lblStatus.Text = '✕ Error'
                    $lblStatus.Foreground = [System.Windows.Media.Brushes]::Salmon
                    $lblFooter.Text = "Exception: $($_.Exception.Message)"
                    [System.Windows.MessageBox]::Show($_.Exception.Message, 'Error', 'OK', 'Error')
                }
                finally {
                    $localPs.Dispose()
                }
            }
            elseif ($script:ManualTimerTicks -ge 400) {
                if ($script:ActiveManualTimer) { try { $script:ActiveManualTimer.Stop() } catch { } }
                $btnCalculate.IsEnabled = $true
                $btnCalculate.Content = '🚀 CALCULATE ROUTE & DOWNLOAD MAP'
                $lblStatus.Text = '✕ Timeout (60s)'
                $lblStatus.Foreground = [System.Windows.Media.Brushes]::Salmon
                $lblFooter.Text = 'Route calculation timed out (60s).'
                try { $localPs.Stop(); $localPs.Dispose() } catch { }
            }
        })
        $timer.Start()
    })

    # External navigation links
    $btnOpenMaps.Add_Click({
        if ($script:LastGoogleMapsUrl) { Start-Process $script:LastGoogleMapsUrl }
    })

    $btnCopyUrl.Add_Click({
        if ($script:LastGoogleMapsUrl) {
            [System.Windows.Clipboard]::SetText($script:LastGoogleMapsUrl)
            [System.Windows.MessageBox]::Show((Get-LocText 'MsgUrlCopied' 'Link copied to clipboard!'), (Get-LocText 'MsgUrlCopiedTitle' 'Copied'), 'OK', 'Information')
        }
    })

    $btnSaveMapAs.Add_Click({
        if ($script:LastGeneratedMapPath -and (Test-Path $script:LastGeneratedMapPath)) {
            $dlg = [System.Windows.Forms.SaveFileDialog]::new()
            $dlg.Title = 'Save PNG Map'
            $dlg.Filter = 'PNG Image (*.png)|*.png'
            $dlg.FileName = [System.IO.Path]::GetFileName($script:LastGeneratedMapPath)
            if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                Copy-Item -LiteralPath $script:LastGeneratedMapPath -Destination $dlg.FileName -Force
                [System.Windows.MessageBox]::Show(((Get-LocText 'MsgMapSaved' 'Map saved to: {0}') -f $dlg.FileName), (Get-LocText 'MsgMapSavedTitle' 'Saved'), 'OK', 'Information')
            }
        }
    })

    # Feature 5.O: Export PDF Report
    $btnExportPdf.Add_Click({
        if (-not $script:LastManualResult) { return }
        $r = $script:LastManualResult
        $dlg = [System.Windows.Forms.SaveFileDialog]::new()
        $dlg.Title = 'Export Route PDF Report'
        $dlg.Filter = 'PDF Document (*.pdf)|*.pdf'
        $dlg.FileName = "Route_$($r.DistanceKm)km.pdf"
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $rName = if ($txtName -and -not [string]::IsNullOrWhiteSpace($txtName.Text)) { $txtName.Text.Trim() } elseif ($script:ActiveManualRouteName) { $script:ActiveManualRouteName } else { 'Route' }
                Export-RoutePdfReport -OutputPath $dlg.FileName `
                    -RouteName $rName `
                    -DistanceKm $r.DistanceKm -DurationMin $r.DurationMin `
                    -RouteType $r.RouteType `
                    -AvoidTolls $r.AvoidTolls -AvoidHighways $r.AvoidHighways -AvoidFerries $r.AvoidFerries `
                    -OriginAddress $r.OriginAddress -OriginLat $r.OriginLat -OriginLng $r.OriginLng `
                    -DestAddress $r.DestAddress -DestLat $r.DestLat -DestLng $r.DestLng `
                    -Waypoints $r.Waypoints -MapImagePath $r.MapPath -GoogleMapsUrl $r.GoogleMapsUrl

                $ask = [System.Windows.MessageBox]::Show("PDF report generated successfully:`n$($dlg.FileName)`n`nDo you want to open it now?", "PDF Export", "YesNo", "Information")
                if ($ask -eq [System.Windows.MessageBoxResult]::Yes) {
                    Start-Process $dlg.FileName
                }
            }
            catch {
                [System.Windows.MessageBox]::Show("Failed to export PDF report:`n$($_.Exception.Message)", "Export Error", "OK", "Error")
            }
        }
    })

    # Feature 5.P: Export GPX
    $btnExportGpx.Add_Click({
        if (-not $script:LastManualResult -or -not $script:LastManualResult.EncodedPolyline) { return }
        $r = $script:LastManualResult
        $dlg = [System.Windows.Forms.SaveFileDialog]::new()
        $dlg.Title = 'Export Route GPX'
        $dlg.Filter = 'GPS Exchange Format (*.gpx)|*.gpx'
        $dlg.FileName = "Route_$($r.DistanceKm)km.gpx"
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $rName = if ($txtName -and -not [string]::IsNullOrWhiteSpace($txtName.Text)) { $txtName.Text.Trim() } elseif ($script:ActiveManualRouteName) { $script:ActiveManualRouteName } else { 'Route' }
                Export-RouteGpx -OutputPath $dlg.FileName `
                    -RouteName $rName `
                    -EncodedPolyline $r.EncodedPolyline `
                    -Waypoints $r.Waypoints `
                    -DistanceKm $r.DistanceKm -DurationMin $r.DurationMin
                [System.Windows.MessageBox]::Show("GPX track exported successfully:`n$($dlg.FileName)", "GPX Export", "OK", "Information")
            }
            catch {
                [System.Windows.MessageBox]::Show("Failed to export GPX:`n$($_.Exception.Message)", "Export Error", "OK", "Error")
            }
        }
    })

    # Feature 5.P: Export KML
    $btnExportKml.Add_Click({
        if (-not $script:LastManualResult -or -not $script:LastManualResult.EncodedPolyline) { return }
        $r = $script:LastManualResult
        $dlg = [System.Windows.Forms.SaveFileDialog]::new()
        $dlg.Title = 'Export Route KML'
        $dlg.Filter = 'Keyhole Markup Language (*.kml)|*.kml'
        $dlg.FileName = "Route_$($r.DistanceKm)km.kml"
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $rName = if ($txtName -and -not [string]::IsNullOrWhiteSpace($txtName.Text)) { $txtName.Text.Trim() } elseif ($script:ActiveManualRouteName) { $script:ActiveManualRouteName } else { 'Route' }
                Export-RouteKml -OutputPath $dlg.FileName `
                    -RouteName $rName `
                    -EncodedPolyline $r.EncodedPolyline `
                    -Waypoints $r.Waypoints `
                    -DistanceKm $r.DistanceKm -DurationMin $r.DurationMin
                [System.Windows.MessageBox]::Show("KML track exported successfully:`n$($dlg.FileName)", "KML Export", "OK", "Information")
            }
            catch {
                [System.Windows.MessageBox]::Show("Failed to export KML:`n$($_.Exception.Message)", "Export Error", "OK", "Error")
            }
        }
    })

    # Initialize map view mode based on configuration
    if ($script:AppConfig -and ($false -eq $script:AppConfig.UseInteractiveMap)) {
        if ($rbViewStatic) { $rbViewStatic.IsChecked = $true }
    } else {
        if ($rbViewInteractive) { $rbViewInteractive.IsChecked = $true }
    }
    Update-MapViewMode
}
