#Requires -Version 5.1
<#
.SYNOPSIS
    Google Maps Routes & Map Generator — Background Runspace Workers Subsystem.
.DESCRIPTION
    Provides thread-safe background execution for manual route calculations,
    batch dataset processing, and pre-batch geocode validation using isolated MTA runspaces.
.NOTES
    Encoding: UTF-8 with BOM
#>

# ══════════════════════════════════════════════════════════════════════════════
# 1. THREAD-SAFE BACKGROUND RUNSPACE WORKER FACTORY (INITIALSESSIONSTATE)
# ══════════════════════════════════════════════════════════════════════════════

function New-WorkerPowerShell {
    param([scriptblock]$ScriptBlock)
    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    Get-ChildItem function: | Where-Object {
        $_.Name -in @(
            'Protect-SecretString', 'Unprotect-SecretString', 'Test-GoogleApiKey',
            'Get-AddressComponentValue', 'Get-AddressCoordinates', 'Get-GeocodeStatusDescription',
            'Get-CarRouteData', 'Get-GoogleMapsUrl', 'Get-WrappedLines', 'Save-RouteMapPng',
            'Find-MatchingPropertyName', 'Import-RouteDataFile', 'Export-RouteResults',
            'ConvertFrom-GoogleEncodedPolyline', 'Export-RouteGpx', 'Export-RouteKml',
            'Get-EstimatedApiCost'
        )
    } | ForEach-Object {
        try {
            $iss.Commands.Add([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new($_.Name, $_.Definition))
        }
        catch { }
    }

    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($iss)
    $rs.ApartmentState = [System.Threading.ApartmentState]::MTA
    $rs.ThreadOptions  = [System.Management.Automation.Runspaces.PSThreadOptions]::UseNewThread
    $rs.Open()

    $ps = [PowerShell]::Create()
    $ps.Runspace = $rs
    # IMPORTANT: $null = prevents PowerShell object leakage into pipeline
    $null = $ps.AddScript($ScriptBlock.ToString())
    return $ps
}

# ══════════════════════════════════════════════════════════════════════════════
# 2. MANUAL CALCULATION WORKER SCRIPTBLOCK (Isolated Runspace, Top-level)
# ══════════════════════════════════════════════════════════════════════════════

$script:ManualCalcAsync = {
    param(
        $start, $end, $waypoints, $routeType, $emission, $trafficAware,
        $name, $apiKey, $outDir, $logFile, $languageCode = 'en',
        $overlayConfigJson = '',
        $avoidTolls = $false, $avoidHighways = $false, $avoidFerries = $false
    )

    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls
    Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

    $geoCount = 0
    $routesCount = 0
    $staticCount = 0

    $wlog = {
        param($msg, $lvl = 'INFO')
        if ($logFile) {
            $t = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff')
            try { [System.IO.File]::AppendAllText($logFile, "[$t] [$lvl] [ManualWorker] $msg`r`n", [System.Text.UTF8Encoding]::new($true)) } catch { }
        }
    }

    try {
        & $wlog "Geocoding origin: '$start'..." "INFO"
        $geoStart = Get-AddressCoordinates -Address $start -ApiKey $apiKey -LanguageCode $languageCode
        $geoCount++
        if ($geoStart.Status -ne 'OK') {
            & $wlog "Origin geocoding error: $($geoStart.Status)" "WARN"
            return [PSCustomObject]@{ Success = $false; Error = "Origin geocoding error: $($geoStart.Status)" }
        }
        & $wlog "Origin OK: $($geoStart.FormattedAddress) ($($geoStart.Latitude), $($geoStart.Longitude))" "INFO"

        & $wlog "Geocoding destination: '$end'..." "INFO"
        $geoEnd = Get-AddressCoordinates -Address $end -ApiKey $apiKey -LanguageCode $languageCode
        $geoCount++
        if ($geoEnd.Status -ne 'OK') {
            & $wlog "Destination geocoding error: $($geoEnd.Status)" "WARN"
            return [PSCustomObject]@{ Success = $false; Error = "Destination geocoding error: $($geoEnd.Status)" }
        }
        & $wlog "Destination OK: $($geoEnd.FormattedAddress) ($($geoEnd.Latitude), $($geoEnd.Longitude))" "INFO"

        $geoWp = [System.Collections.Generic.List[PSCustomObject]]::new()
        if ($waypoints) {
            foreach ($w in $waypoints) {
                if (-not [string]::IsNullOrWhiteSpace($w)) {
                    & $wlog "Geocoding waypoint: '$w'..." "INFO"
                    $g = Get-AddressCoordinates -Address $w -ApiKey $apiKey -LanguageCode $languageCode
                    $geoCount++
                    if ($g.Status -eq 'OK') {
                        $geoWp.Add($g)
                        & $wlog "Waypoint OK: $($g.FormattedAddress)" "INFO"
                    }
                    else {
                        & $wlog "Waypoint geocoding error '$w': $($g.Status)" "WARN"
                    }
                }
            }
        }

        & $wlog "Querying Google Routes API v2 (Type: $routeType, Engine: $emission, AvoidTolls: $avoidTolls, AvoidHighways: $avoidHighways, AvoidFerries: $avoidFerries)..." "INFO"
        $trasa = Get-CarRouteData -OriginLat $geoStart.Latitude -OriginLng $geoStart.Longitude `
            -DestLat $geoEnd.Latitude -DestLng $geoEnd.Longitude `
            -IntermediatePoints $geoWp -RouteType $routeType -EmissionType $emission `
            -ApiKey $apiKey -LanguageCode $languageCode -TrafficAware:$trafficAware `
            -AvoidTolls:$avoidTolls -AvoidHighways:$avoidHighways -AvoidFerries:$avoidFerries
        $routesCount++

        if ($trasa.Status -ne 'OK') {
            & $wlog "Routes API error: $($trasa.Status). $($trasa.ErrorMessage)" "WARN"
            return [PSCustomObject]@{ Success = $false; Error = "Routes API error: $($trasa.Status). $($trasa.ErrorMessage)" }
        }
        & $wlog "Routes API route found: $($trasa.OdlegloscKm) km, $($trasa.CzasMin) min" "INFO"

        $gUrl = Get-GoogleMapsUrl -Origin "$($geoStart.Latitude),$($geoStart.Longitude)" `
            -Destination "$($geoEnd.Latitude),$($geoEnd.Longitude)" `
            -Waypoints $geoWp

        $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
        $safeName = ($name -replace '[\\/:*?"<>|]', '_').Trim()
        $mapFileName = "${ts}_manual_route_${safeName}.png"
        $mapPath = Join-Path $outDir $mapFileName

        $allPts = [System.Collections.Generic.List[PSCustomObject]]::new()
        $allPts.Add($geoStart)
        foreach ($pt in $geoWp) { $allPts.Add($pt) }
        $allPts.Add($geoEnd)

        & $wlog "Rendering static map image: $mapPath..." "INFO"
        $hdrTypePrefix = switch ($languageCode) { 'de' { 'Typ: ' } 'pl' { 'Typ: ' } default { 'Type: ' } }
        $hdrTypeName = switch ($languageCode) {
            'de' { if ($routeType -eq 'Fastest') { 'Schnellste' } elseif ($routeType -eq 'Shortest') { 'Kürzeste' } else { 'Eco' } }
            'pl' { if ($routeType -eq 'Fastest') { 'Najszybsza' } elseif ($routeType -eq 'Shortest') { 'Najkrótsza' } else { 'Eko' } }
            default { $routeType }
        }
        $headerRightText = "$hdrTypePrefix$hdrTypeName"

        $saved = Save-RouteMapPng -EncodedPolyline $trasa.EncodedPolyline `
            -OriginLat $geoStart.Latitude -OriginLng $geoStart.Longitude `
            -DestLat $geoEnd.Latitude -DestLng $geoEnd.Longitude `
            -RoutePoints $allPts -OutputPath $mapPath -ApiKey $apiKey `
            -Width 900 -Height 600 `
            -AddressTextA $geoStart.FormattedAddress -AddressTextB $geoEnd.FormattedAddress `
            -DistanceText "$($trasa.OdlegloscKm) km" -DurationText "$($trasa.CzasMin) min" `
            -HeaderLeftText $name -HeaderRightText $headerRightText `
            -LanguageCode $languageCode `
            -StartRaw $start -StartGeocoded $geoStart.FormattedAddress `
            -EndRaw $end -EndGeocoded $geoEnd.FormattedAddress `
            -WaypointsList $geoWp -RouteName $name -RouteType $headerRightText `
            -OverlayConfig $overlayConfigJson
        $staticCount++

        & $wlog "Map rendering complete. Saved: $saved" "INFO"
        $resolvedMapPath = $(if ($saved) { $mapPath } else { $null })

        return [PSCustomObject]@{
            Success         = $true
            DistanceKm      = $trasa.OdlegloscKm
            DurationMin     = $trasa.CzasMin
            RouteType       = $routeType
            EncodedPolyline = $trasa.EncodedPolyline
            GoogleMapsUrl   = $gUrl
            MapPath         = $resolvedMapPath
            OriginLat       = $geoStart.Latitude
            OriginLng       = $geoStart.Longitude
            OriginAddress   = $geoStart.FormattedAddress
            DestLat         = $geoEnd.Latitude
            DestLng         = $geoEnd.Longitude
            DestAddress     = $geoEnd.FormattedAddress
            Waypoints       = $geoWp
            AvoidTolls      = $avoidTolls
            AvoidHighways   = $avoidHighways
            AvoidFerries    = $avoidFerries
            ApiUsage        = [PSCustomObject]@{
                Geocoding  = $geoCount
                Routes     = $routesCount
                StaticMaps = $staticCount
            }
            Error           = $null
        }
    }
    catch {
        $errFull = $_.Exception.ToString()
        & $wlog "Worker thread exception: $errFull" "ERROR"
        return [PSCustomObject]@{ Success = $false; Error = $_.Exception.Message }
    }
}

# ══════════════════════════════════════════════════════════════════════════════
# 3. BATCH PROCESSING WORKER SCRIPTBLOCK (Isolated Runspace, Top-level)
# ══════════════════════════════════════════════════════════════════════════════

$script:BatchCalcAsync = {
    param(
        $routes, $apiKey, $outDir, $defaultRouteType, $syncState, $logFile,
        $languageCode = 'en', $overlayConfigJson = '',
        $defaultAvoidTolls = $false, $defaultAvoidHighways = $false, $defaultAvoidFerries = $false
    )

    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls
    Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

    $geoCount = 0
    $routesCount = 0
    $staticCount = 0

    $wlog = {
        param($msg, $lvl = 'INFO')
        if ($syncState.LogQueue) {
            $syncState.LogQueue.Enqueue([PSCustomObject]@{ Level = $lvl; Message = $msg })
        }
        if ($logFile) {
            $t = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff')
            try { [System.IO.File]::AppendAllText($logFile, "[$t] [$lvl] [BatchWorker] $msg`r`n", [System.Text.UTF8Encoding]::new($true)) } catch { }
        }
    }

    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $results = [System.Collections.Generic.List[PSCustomObject]]::new()
    $total = $routes.Count

    for ($i = 0; $i -lt $total; $i++) {
        if ($syncState.CancelRequested) {
            & $wlog "Batch processing stopped by user at route $($i + 1)/$total." "WARN"
            break
        }
        $r = $routes[$i]
        $syncState.CurrentIndex = ($i + 1)

        $rType = if ($defaultRouteType -and $defaultRouteType -ne 'FromSource') { $defaultRouteType }
        elseif ($r.RouteType) { $r.RouteType }
        else { 'Fastest' }

        $avoidT = if ($null -ne $r.AvoidTolls) { [bool]$r.AvoidTolls } else { $defaultAvoidTolls }
        $avoidH = if ($null -ne $r.AvoidHighways) { [bool]$r.AvoidHighways } else { $defaultAvoidHighways }
        $avoidF = if ($null -ne $r.AvoidFerries) { [bool]$r.AvoidFerries } else { $defaultAvoidFerries }

        $routeName = if ($r.Name) { $r.Name } else { "Route $($i + 1)" }

        & $wlog "Route $($i + 1)/$($total): Processing '$($r.Start)' -> '$($r.End)' (Type: $rType)..." "INFO"

        try {
            $geoStart = Get-AddressCoordinates -Address $r.Start -ApiKey $apiKey -LanguageCode $languageCode
            $geoCount++
            $startStatus = Get-GeocodeStatusDescription -Geo $geoStart
            $isStartFallback = if ($geoStart -and ($geoStart.PartialMatch -or $geoStart.MatchType -in 'APPROXIMATE', 'GEOMETRIC_CENTER')) { $true } else { $false }

            $routePointsList = [System.Collections.Generic.List[PSCustomObject]]::new()
            $routePointsList.Add([PSCustomObject]@{
                Order           = 1
                PointType       = 'Start'
                OriginalAddress = $r.Start
                GeocodedAddress = if ($geoStart) { $geoStart.FormattedAddress } else { $null }
                GeocodeStatus   = $startStatus
                MatchType       = if ($geoStart) { $geoStart.MatchType } else { 'NOT_FOUND' }
                PartialMatch    = if ($geoStart) { [bool]$geoStart.PartialMatch } else { $false }
                IsFallback      = $isStartFallback
                Latitude        = if ($geoStart) { $geoStart.Latitude } else { $null }
                Longitude       = if ($geoStart) { $geoStart.Longitude } else { $null }
            })

            $geoEnd = Get-AddressCoordinates -Address $r.End -ApiKey $apiKey -LanguageCode $languageCode
            $geoCount++
            $endStatus = Get-GeocodeStatusDescription -Geo $geoEnd
            $isEndFallback = if ($geoEnd -and ($geoEnd.PartialMatch -or $geoEnd.MatchType -in 'APPROXIMATE', 'GEOMETRIC_CENTER')) { $true } else { $false }

            # Waypoints
            $geoWp = [System.Collections.Generic.List[PSCustomObject]]::new()
            $wpOrder = 2
            if ($r.Waypoints -and @($r.Waypoints).Count -gt 0) {
                foreach ($wpAddr in @($r.Waypoints)) {
                    if (-not [string]::IsNullOrWhiteSpace($wpAddr)) {
                        $gw = Get-AddressCoordinates -Address $wpAddr -ApiKey $apiKey -LanguageCode $languageCode
                        $geoCount++
                        $gwStatus = Get-GeocodeStatusDescription -Geo $gw
                        $isGwFallback = if ($gw -and ($gw.PartialMatch -or $gw.MatchType -in 'APPROXIMATE', 'GEOMETRIC_CENTER')) { $true } else { $false }
                        $routePointsList.Add([PSCustomObject]@{
                            Order           = $wpOrder
                            PointType       = "Waypoint $($wpOrder - 1)"
                            OriginalAddress = $wpAddr
                            GeocodedAddress = if ($gw) { $gw.FormattedAddress } else { $null }
                            GeocodeStatus   = $gwStatus
                            MatchType       = if ($gw) { $gw.MatchType } else { 'NOT_FOUND' }
                            PartialMatch    = if ($gw) { [bool]$gw.PartialMatch } else { $false }
                            IsFallback      = $isGwFallback
                            Latitude        = if ($gw) { $gw.Latitude } else { $null }
                            Longitude       = if ($gw) { $gw.Longitude } else { $null }
                        })
                        if ($gw -and $gw.Status -eq 'OK') { $geoWp.Add($gw) }
                        $wpOrder++
                    }
                }
            }

            $routePointsList.Add([PSCustomObject]@{
                Order           = $wpOrder
                PointType       = 'End'
                OriginalAddress = $r.End
                GeocodedAddress = if ($geoEnd) { $geoEnd.FormattedAddress } else { $null }
                GeocodeStatus   = $endStatus
                MatchType       = if ($geoEnd) { $geoEnd.MatchType } else { 'NOT_FOUND' }
                PartialMatch    = if ($geoEnd) { [bool]$geoEnd.PartialMatch } else { $false }
                IsFallback      = $isEndFallback
                Latitude        = if ($geoEnd) { $geoEnd.Latitude } else { $null }
                Longitude       = if ($geoEnd) { $geoEnd.Longitude } else { $null }
            })

            if (-not $geoStart -or $geoStart.Status -ne 'OK' -or -not $geoEnd -or $geoEnd.Status -ne 'OK') {
                & $wlog "Route $($i + 1): Geocoding failed (Start: $($geoStart.Status), End: $($geoEnd.Status)). Skipping route." "WARN"
                $results.Add([PSCustomObject]@{
                    Id               = if ($r.Id) { $r.Id } else { ($i + 1) }
                    Name             = $routeName
                    Start_Original   = $r.Start
                    Start_Geocoded   = if ($geoStart) { $geoStart.FormattedAddress } else { $null }
                    Start_Status     = $startStatus
                    End_Original     = $r.End
                    End_Geocoded     = if ($geoEnd) { $geoEnd.FormattedAddress } else { $null }
                    End_Status       = $endStatus
                    WaypointsCount   = $geoWp.Count
                    RouteType        = $rType
                    DistanceKm       = $null
                    DurationMin      = $null
                    Status           = "Geocode Error"
                    MapPath          = $null
                    EncodedPolyline  = $null
                    RoutePoints      = $routePointsList
                })
                continue
            }

            # Compute route
            $routeData = Get-CarRouteData -OriginLat $geoStart.Latitude -OriginLng $geoStart.Longitude `
                -DestLat $geoEnd.Latitude -DestLng $geoEnd.Longitude `
                -IntermediatePoints $geoWp -RouteType $rType `
                -ApiKey $apiKey -LanguageCode $languageCode `
                -AvoidTolls:$avoidT -AvoidHighways:$avoidH -AvoidFerries:$avoidF
            $routesCount++

            if ($routeData.Status -ne 'OK') {
                & $wlog "Route $($i + 1): Route calculation failed ($($routeData.Status))." "WARN"
                $results.Add([PSCustomObject]@{
                    Id               = if ($r.Id) { $r.Id } else { ($i + 1) }
                    Name             = $routeName
                    Start_Original   = $r.Start
                    Start_Geocoded   = $geoStart.FormattedAddress
                    Start_Status     = $startStatus
                    End_Original     = $r.End
                    End_Geocoded     = $geoEnd.FormattedAddress
                    End_Status       = $endStatus
                    WaypointsCount   = $geoWp.Count
                    RouteType        = $rType
                    DistanceKm       = $null
                    DurationMin      = $null
                    Status           = "Route Error ($($routeData.Status))"
                    MapPath          = $null
                    EncodedPolyline  = $null
                    RoutePoints      = $routePointsList
                })
                continue
            }

            # Render map
            $cleanName = ($routeName -replace '[\\/:*?"<>|]', '_').Trim()
            $mapFileName = "${ts}_route_$($i + 1)_${cleanName}.png"
            $mapPath = Join-Path $outDir $mapFileName

            $allPts = [System.Collections.Generic.List[PSCustomObject]]::new()
            $allPts.Add($geoStart)
            foreach ($pt in $geoWp) { $allPts.Add($pt) }
            $allPts.Add($geoEnd)

            $hdrTypePrefix = switch ($languageCode) { 'de' { 'Typ: ' } 'pl' { 'Typ: ' } default { 'Type: ' } }
            $hdrTypeName = switch ($languageCode) {
                'de' { if ($rType -eq 'Fastest') { 'Schnellste' } elseif ($rType -eq 'Shortest') { 'Kürzeste' } else { 'Eco' } }
                'pl' { if ($rType -eq 'Fastest') { 'Najszybsza' } elseif ($rType -eq 'Shortest') { 'Najkrótsza' } else { 'Eko' } }
                default { $rType }
            }
            $headerRightText = "$hdrTypePrefix$hdrTypeName"

            $saved = Save-RouteMapPng -EncodedPolyline $routeData.EncodedPolyline `
                -OriginLat $geoStart.Latitude -OriginLng $geoStart.Longitude `
                -DestLat $geoEnd.Latitude -DestLng $geoEnd.Longitude `
                -RoutePoints $allPts -OutputPath $mapPath -ApiKey $apiKey `
                -Width 900 -Height 600 `
                -AddressTextA $geoStart.FormattedAddress -AddressTextB $geoEnd.FormattedAddress `
                -DistanceText "$($routeData.OdlegloscKm) km" -DurationText "$($routeData.CzasMin) min" `
                -HeaderLeftText $routeName -HeaderRightText $headerRightText `
                -LanguageCode $languageCode `
                -StartRaw $r.Start -StartGeocoded $geoStart.FormattedAddress `
                -EndRaw $r.End -EndGeocoded $geoEnd.FormattedAddress `
                -WaypointsList $geoWp -RouteName $routeName -RouteType $headerRightText `
                -OverlayConfig $overlayConfigJson
            $staticCount++

            & $wlog "Route $($i + 1)/$total OK: $($routeData.OdlegloscKm) km, $($routeData.CzasMin) min" "OK"

            $results.Add([PSCustomObject]@{
                Id               = if ($r.Id) { $r.Id } else { ($i + 1) }
                Name             = $routeName
                Start_Original   = $r.Start
                Start_Geocoded   = $geoStart.FormattedAddress
                Start_Status     = $startStatus
                End_Original     = $r.End
                End_Geocoded     = $geoEnd.FormattedAddress
                End_Status       = $endStatus
                WaypointsCount   = $geoWp.Count
                RouteType        = $rType
                DistanceKm       = $routeData.OdlegloscKm
                DurationMin      = $routeData.CzasMin
                Status           = 'OK'
                MapPath          = $(if ($saved) { $mapPath } else { $null })
                EncodedPolyline  = $routeData.EncodedPolyline
                RoutePoints      = $routePointsList
                AvoidTolls       = $avoidT
                AvoidHighways    = $avoidH
                AvoidFerries     = $avoidF
            })
        }
        catch {
            & $wlog "Route $($i + 1) exception: $($_.Exception.Message)" "ERROR"
            $results.Add([PSCustomObject]@{
                Id               = if ($r.Id) { $r.Id } else { ($i + 1) }
                Name             = $routeName
                Start_Original   = $r.Start
                Start_Geocoded   = $null
                Start_Status     = "Exception"
                End_Original     = $r.End
                End_Geocoded     = $null
                End_Status       = "Exception"
                WaypointsCount   = 0
                RouteType        = $rType
                DistanceKm       = $null
                DurationMin      = $null
                Status           = "Exception: $($_.Exception.Message)"
                MapPath          = $null
                EncodedPolyline  = $null
                RoutePoints      = @()
            })
        }
    }

    return [PSCustomObject]@{
        Results  = $results
        ApiUsage = [PSCustomObject]@{
            Geocoding  = $geoCount
            Routes     = $routesCount
            StaticMaps = $staticCount
        }
    }
}

# ══════════════════════════════════════════════════════════════════════════════
# 4. GEOCODE VALIDATION PREVIEW WORKER SCRIPTBLOCK (Feature 3.J)
# ══════════════════════════════════════════════════════════════════════════════

$script:GeocodeValidationAsync = {
    param($addressItems, $apiKey, $languageCode = 'en', $syncState, $logFile)

    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls

    $geoCount = 0
    $wlog = {
        param($msg, $lvl = 'INFO')
        if ($syncState.LogQueue) {
            $syncState.LogQueue.Enqueue([PSCustomObject]@{ Level = $lvl; Message = $msg })
        }
        if ($logFile) {
            $t = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff')
            try { [System.IO.File]::AppendAllText($logFile, "[$t] [$lvl] [ValidationWorker] $msg`r`n", [System.Text.UTF8Encoding]::new($true)) } catch { }
        }
    }

    $results = [System.Collections.Generic.List[PSCustomObject]]::new()
    $total = $addressItems.Count
    & $wlog "Starting pre-batch geocode validation ($total unique addresses)..." "INFO"

    for ($i = 0; $i -lt $total; $i++) {
        if ($syncState.CancelRequested) {
            & $wlog "Geocode validation cancelled by user at address $($i + 1)/$total." "WARN"
            break
        }
        $item = $addressItems[$i]
        $syncState.CurrentIndex = ($i + 1)
        $rawAddr = [string]$item.Address

        & $wlog "Validating [$($i + 1)/$total]: '$rawAddr'..." "INFO"

        $geo = Get-AddressCoordinates -Address $rawAddr -ApiKey $apiKey -LanguageCode $languageCode
        $geoCount++

        $statusStr = if ($geo.Status -eq 'OK') { 'OK' } else { $geo.Status }
        $matchType = if ($geo.MatchType) { $geo.MatchType } else { 'NONE' }
        $lat = if ($geo.Latitude) { [double]$geo.Latitude } else { $null }
        $lng = if ($geo.Longitude) { [double]$geo.Longitude } else { $null }
        $formatted = if ($geo.FormattedAddress) { [string]$geo.FormattedAddress } else { '' }

        $precision = switch ($matchType) {
            'ROOFTOP'            { 'ROOFTOP (Exact)' }
            'RANGE_INTERPOLATED' { 'RANGE (Interpolated)' }
            'GEOMETRIC_CENTER'   { 'CENTER (Geometric)' }
            'APPROXIMATE'        { 'APPROXIMATE (Low Precision)' }
            default              { $matchType }
        }

        $results.Add([PSCustomObject]@{
            Index            = ($i + 1)
            Address          = $rawAddr
            Role             = [string]$item.Role
            Status           = $statusStr
            Precision        = $precision
            FormattedAddress = $formatted
            Latitude         = $lat
            Longitude        = $lng
            IsOk             = ($geo.Status -eq 'OK')
            IsRooftop        = ($matchType -eq 'ROOFTOP')
        })
    }

    & $wlog "Geocode validation completed. Processed $geoCount addresses." "OK"
    return [PSCustomObject]@{
        Results  = $results
        ApiUsage = [PSCustomObject]@{
            Geocoding  = $geoCount
            Routes     = 0
            StaticMaps = 0
        }
    }
}
