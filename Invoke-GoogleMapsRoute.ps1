#Requires -Version 5.1
<#
.SYNOPSIS
    Calculates Google Maps driving routes from manual parameters or data files (JSON, CSV, Excel).

.DESCRIPTION
    The script provides:
      1. Manual route calculation between origin, destination, and optional waypoints
      2. Batch loading of data files (JSON, CSV, Excel) with route lists or stop sequences
      3. Route optimization modes: Fastest (travel time), Shortest (distance), Eco-friendly (fuel/energy efficiency)
      4. Address geocoding using Google Geocoding API
      5. Route computation using Google Routes API v2
      6. High-resolution PNG map rendering with highlighted polyline and numbered markers
      7. Interactive Google Maps navigation URL generation
      8. Comprehensive batch summary export (Excel, CSV, JSON)

.PARAMETER StartPoint
    Address or latitude/longitude coordinates of origin (e.g., "Warszawa, Marszałkowska 1" or "52.2297, 21.0122").

.PARAMETER EndPoint
    Address or latitude/longitude coordinates of destination.

.PARAMETER Waypoints
    Array of addresses or coordinates for intermediate stops (up to 25 waypoints).

.PARAMETER Name
    Descriptive name or label for manual route.

.PARAMETER RouteType
    Route optimization mode:
      - 'Fastest': Optimized for shortest travel duration
      - 'Shortest': Optimized for shortest physical travel distance (km)
      - 'Eco': Optimized for lowest fuel or energy consumption
      - 'FromSource': Inherited from input data file or default

.PARAMETER EmissionType
    Vehicle powertrain type for Eco routes: 'GASOLINE', 'DIESEL', 'HYBRID', 'ELECTRIC'. Default: 'GASOLINE'.

.PARAMETER InputFile
    Path to input route data file (.xlsx, .xls, .csv, .tsv, .json).

.PARAMETER ExportFormat
    Export report format for batch processing: 'Excel', 'CSV', 'JSON', 'All', 'None'. Default: 'Excel'.

.PARAMETER ApiKey
    Google Maps API key. If omitted, resolved from GOOGLE_MAPS_API_KEY environment variable or DPAPI storage.

.PARAMETER OutputFolder
    Output directory for generated PNG maps and export files. Default: .\Results

.PARAMETER GenerateMap
    Switch indicating whether to download and render static PNG maps. Default: $true.

.PARAMETER MapWidth
    Width of the rendered PNG map in pixels (100-2048). Default: 900.

.PARAMETER MapHeight
    Height of the rendered PNG map in pixels (100-2048). Default: 600.

.PARAMETER TrafficAware
    Switch indicating whether to incorporate live real-time traffic conditions into route calculation.

.PARAMETER OpenBrowser
    Switch indicating whether to open the calculated route in the default web browser.

.EXAMPLE
    # Fastest route with intermediate waypoints
    .\Invoke-GoogleMapsRoute.ps1 -StartPoint "Warszawa, Marszałkowska 1" -EndPoint "Kraków, Rynek Główny 1" `
        -Waypoints "Radom, Żeromskiego 5", "Kielce, Sienkiewicza 10" -RouteType Fastest -GenerateMap

.EXAMPLE
    # Shortest distance route
    .\Invoke-GoogleMapsRoute.ps1 -StartPoint "Gdańsk, Długa 1" -EndPoint "Toruń, Szeroka 1" -RouteType Shortest

.EXAMPLE
    # Batch processing of Excel or JSON data file
    .\Invoke-GoogleMapsRoute.ps1 -InputFile ".\Samplesoutes_sample.xlsx" -RouteType Fastest -ExportFormat Excel

.NOTES
    Encoding: UTF-8 with BOM
#>

[CmdletBinding(DefaultParameterSetName = 'Manual')]
param(
    # --- Parameter Set: Manual ---
    [Parameter(Mandatory = $true, ParameterSetName = 'Manual', Position = 0)]
    [string]$StartPoint,

    [Parameter(Mandatory = $true, ParameterSetName = 'Manual', Position = 1)]
    [string]$EndPoint,

    [Parameter(Mandatory = $false, ParameterSetName = 'Manual')]
    [string[]]$Waypoints = @(),

    [Parameter(Mandatory = $false, ParameterSetName = 'Manual')]
    [string]$Name = 'Manual Route',

    # --- Parameter Set: File ---
    [Parameter(Mandatory = $true, ParameterSetName = 'File', Position = 0)]
    [string]$InputFile,

    [Parameter(Mandatory = $false, ParameterSetName = 'File')]
    [ValidateSet('Excel', 'CSV', 'JSON', 'All', 'None')]
    [string]$ExportFormat = 'Excel',

    # --- Common Parameters ---
    [Parameter(Mandatory = $false)]
    [ValidateSet('Fastest', 'Shortest', 'Eco', 'FromSource')]
    [string]$RouteType = 'Fastest',

    [Parameter(Mandatory = $false)]
    [ValidateSet('GASOLINE', 'DIESEL', 'HYBRID', 'ELECTRIC')]
    [string]$EmissionType = 'GASOLINE',

    [Parameter(Mandatory = $false)]
    [string]$ApiKey = $env:GOOGLE_MAPS_API_KEY,

    [Parameter(Mandatory = $false)]
    [string]$OutputFolder,

    [Parameter(Mandatory = $false)]
    [bool]$GenerateMap = $true,

    [Parameter(Mandatory = $false)]
    [ValidateRange(100, 2048)]
    [int]$MapWidth = 900,

    [Parameter(Mandatory = $false)]
    [ValidateRange(100, 2048)]
    [int]$MapHeight = 600,

    [Parameter(Mandatory = $false)]
    [switch]$TrafficAware,

    [Parameter(Mandatory = $false)]
    [switch]$OpenBrowser
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Load shared functions library
$FunctionsPath = Join-Path $PSScriptRoot 'RouteMapFunctions.ps1'
if (-not (Test-Path $FunctionsPath)) {
    throw "Shared functions library file not found: $FunctionsPath"
}
. $FunctionsPath

# Resolve API key
if ([string]::IsNullOrWhiteSpace($ApiKey)) {
    # Attempt resolution from local DPAPI configuration
    $AppDir = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'GoogleMapsRoutes'
    $CfgFile = Join-Path $AppDir 'config.json'
    if (Test-Path $CfgFile) {
        try {
            $cfg = Get-Content -LiteralPath $CfgFile -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($cfg.ApiKey) {
                $ApiKey = Unprotect-SecretString -EncryptedText $cfg.ApiKey
            }
        }
        catch { }
    }
}

if ([string]::IsNullOrWhiteSpace($ApiKey)) {
    throw "Brak klucza Google Maps API. Ustaw zmienną środowiskową GOOGLE_MAPS_API_KEY lub przekaż parametr -ApiKey."
}

# Resolve output folder
if ([string]::IsNullOrWhiteSpace($OutputFolder)) {
    $OutputFolder = Join-Path $PSScriptRoot 'Results'
}
if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}

$Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'

# ══════════════════════════════════════════════════════════════════════════════
# MODE 1: MANUAL PARAMETERS (MANUAL)
# ══════════════════════════════════════════════════════════════════════════════
if ($PSCmdlet.ParameterSetName -eq 'Manual') {
    Write-Host "`n[Google Maps Routes] Rozpoczynanie obliczania trasy..." -ForegroundColor Cyan
    Write-Host "  Start     : $StartPoint" -ForegroundColor White
    Write-Host "  Cel       : $EndPoint" -ForegroundColor White
    if ($Waypoints -and $Waypoints.Count -gt 0) {
        Write-Host "  Punkty pośrednie ($($Waypoints.Count)):" -ForegroundColor White
        for ($i = 0; $i -lt $Waypoints.Count; $i++) {
            Write-Host "    $($i + 1). $($Waypoints[$i])" -ForegroundColor Gray
        }
    }
    Write-Host "  Typ trasy : $RouteType" -ForegroundColor White
    if ($RouteType -eq 'Eco') {
        Write-Host "  Napęd     : $EmissionType" -ForegroundColor Gray
    }

    # Geokodowanie Startu
    Write-Host "`nGeokodowanie punktu początkowego: '$StartPoint'..." -ForegroundColor Yellow
    $GeoStart = Get-AddressCoordinates -Address $StartPoint -ApiKey $ApiKey
    if ($GeoStart.Status -ne 'OK') {
        throw "Nie udało się geokodować punktu startowego: '$StartPoint'. Status: $($GeoStart.Status)"
    }
    Write-Host "  OK: $($GeoStart.FormattedAddress) ($($GeoStart.Latitude), $($GeoStart.Longitude))" -ForegroundColor Green

    # Geokodowanie Celu
    Write-Host "Geokodowanie punktu docelowego: '$EndPoint'..." -ForegroundColor Yellow
    $GeoEnd = Get-AddressCoordinates -Address $EndPoint -ApiKey $ApiKey
    if ($GeoEnd.Status -ne 'OK') {
        throw "Nie udało się geokodować punktu docelowego: '$EndPoint'. Status: $($GeoEnd.Status)"
    }
    Write-Host "  OK: $($GeoEnd.FormattedAddress) ($($GeoEnd.Latitude), $($GeoEnd.Longitude))" -ForegroundColor Green

    # Geocoding Waypoints
    $GeocodedWaypoints = [System.Collections.Generic.List[PSCustomObject]]::new()
    if ($Waypoints -and $Waypoints.Count -gt 0) {
        $idx = 1
        foreach ($wp in $Waypoints) {
            if ([string]::IsNullOrWhiteSpace($wp)) { continue }
            Write-Host "Geokodowanie punktu pośredniego [$idx/$($Waypoints.Count)]: '$wp'..." -ForegroundColor Yellow
            $g = Get-AddressCoordinates -Address $wp -ApiKey $ApiKey
            if ($g.Status -ne 'OK') {
                throw "Nie udało się geokodować punktu pośredniego: '$wp'. Status: $($g.Status)"
            }
            Write-Host "  OK: $($g.FormattedAddress) ($($g.Latitude), $($g.Longitude))" -ForegroundColor Green
            $GeocodedWaypoints.Add($g)
            $idx++
            Start-Sleep -Milliseconds 150
        }
    }

    # Obliczenie trasy
    Write-Host "`nObliczanie trasy ($RouteType)..." -ForegroundColor Cyan
    $ActualRouteType = if ($RouteType -eq 'FromSource') { 'Fastest' } else { $RouteType }
    $Trasa = Get-CarRouteData -OriginLat $GeoStart.Latitude -OriginLng $GeoStart.Longitude `
        -DestLat $GeoEnd.Latitude -DestLng $GeoEnd.Longitude `
        -IntermediatePoints $GeocodedWaypoints -RouteType $ActualRouteType `
        -EmissionType $EmissionType -ApiKey $ApiKey -TrafficAware:$TrafficAware

    if ($Trasa.Status -ne 'OK') {
        throw "Błąd obliczania trasy: $($Trasa.Status). $($Trasa.ErrorMessage)"
    }

    Write-Host "  Odległość: $($Trasa.OdlegloscKm) km" -ForegroundColor Green
    Write-Host "  Czas     : $($Trasa.CzasMin) min" -ForegroundColor Green

    # Generowanie linku Google Maps
    $GoogleMapsUrl = Get-GoogleMapsUrl -Origin "$($GeoStart.Latitude),$($GeoStart.Longitude)" `
        -Destination "$($GeoEnd.Latitude),$($GeoEnd.Longitude)" `
        -Waypoints $GeocodedWaypoints

    # Generowanie mapy PNG
    $MapPath = $null
    if ($GenerateMap -and $Trasa.EncodedPolyline) {
        $SafeName = ($Name -replace '[\\/:*?"<>|]', '_').Trim()
        $MapFileName = "${Timestamp}_${SafeName}.png"
        $MapPath = Join-Path $OutputFolder $MapFileName

        $AllRoutePoints = [System.Collections.Generic.List[PSCustomObject]]::new()
        $AllRoutePoints.Add($GeoStart)
        foreach ($wp in $GeocodedWaypoints) { $AllRoutePoints.Add($wp) }
        $AllRoutePoints.Add($GeoEnd)

        $MapSaved = Save-RouteMapPng -EncodedPolyline $Trasa.EncodedPolyline `
            -OriginLat $GeoStart.Latitude -OriginLng $GeoStart.Longitude `
            -DestLat $GeoEnd.Latitude -DestLng $GeoEnd.Longitude `
            -RoutePoints $AllRoutePoints -OutputPath $MapPath -ApiKey $ApiKey `
            -Width $MapWidth -Height $MapHeight `
            -AddressTextA $GeoStart.FormattedAddress -AddressTextB $GeoEnd.FormattedAddress `
            -DistanceText "$($Trasa.OdlegloscKm) km" -DurationText "$($Trasa.CzasMin) min" `
            -HeaderLeftText $Name -HeaderRightText "Typ: $ActualRouteType" `
            -Legs $Trasa.Legs

        if ($MapSaved) {
            Write-Host "  Mapa PNG : $MapPath" -ForegroundColor Cyan
        }
    }

    Write-Host "  Link     : $GoogleMapsUrl" -ForegroundColor Cyan

    if ($OpenBrowser) {
        Start-Process $GoogleMapsUrl
    }

    $ResultObj = [PSCustomObject]@{
        Id               = '1'
        Nazwa            = $Name
        Start            = $StartPoint
        StartGeokodowany = $GeoStart.FormattedAddress
        StartKoordynaty  = "$($GeoStart.Latitude), $($GeoStart.Longitude)"
        Koniec           = $EndPoint
        KoniecGeokodowany= $GeoEnd.FormattedAddress
        KoniecKoordynaty = "$($GeoEnd.Latitude), $($GeoEnd.Longitude)"
        LiczbaPrzystankow= $GeocodedWaypoints.Count
        TypTrasy         = $ActualRouteType
        OdlegloscKm      = $Trasa.OdlegloscKm
        CzasMin          = $Trasa.CzasMin
        MapaPath         = $MapPath
        GoogleMapsUrl    = $GoogleMapsUrl
        Status           = 'OK'
    }

    return $ResultObj
}

# ══════════════════════════════════════════════════════════════════════════════
# TRYB 2: PLIK DANYCH (FILE)
# ══════════════════════════════════════════════════════════════════════════════
if ($PSCmdlet.ParameterSetName -eq 'File') {
    Write-Host "`n[Google Maps Routes] Wczytywanie pliku danych: $InputFile" -ForegroundColor Cyan
    $Imported = Import-RouteDataFile -Path $InputFile
    Write-Host "  Format        : $($Imported.Format)" -ForegroundColor White
    Write-Host "  Tryb danych   : $($Imported.Mode)" -ForegroundColor White
    Write-Host "  Liczba pozycji: $($Imported.TotalCount)" -ForegroundColor White

    if ($Imported.Routes.Count -eq 0) {
        Write-Warning "Plik nie zawiera żadnych tras do przetworzenia."
        return @()
    }

    $ResultsList = [System.Collections.Generic.List[PSCustomObject]]::new()
    $CurrentIdx = 0
    $TotalCount = $Imported.Routes.Count

    foreach ($r in $Imported.Routes) {
        $CurrentIdx++
        $RouteName = if ($r.Name) { $r.Name } else { "Trasa $CurrentIdx" }
        Write-Host "`n[$CurrentIdx/$TotalCount] Przetwarzanie: $RouteName" -ForegroundColor Cyan
        Write-Host "  Start: $($r.Start)" -ForegroundColor White
        Write-Host "  Cel  : $($r.End)" -ForegroundColor White

        # Resolve RouteType for row
        $RowRouteType = if ($RouteType -ne 'FromSource') {
            $RouteType
        }
        elseif ($r.RouteType) {
            $r.RouteType
        }
        else {
            'Fastest'
        }

        try {
            # Geokodowanie Startu
            $GeoStart = Get-AddressCoordinates -Address $r.Start -ApiKey $ApiKey
            $StartStatus = Get-GeocodeStatusDescription -Geo $GeoStart
            $IsStartFallback = if ($GeoStart -and ($GeoStart.PartialMatch -or $GeoStart.MatchType -in 'APPROXIMATE', 'GEOMETRIC_CENTER')) { $true } else { $false }
            $RoutePoints = [System.Collections.Generic.List[PSCustomObject]]::new()
            $RoutePoints.Add([PSCustomObject]@{
                Order           = 1
                PointType       = 'Start'
                LegDistanceKm   = 0
                LegDurationMin  = 0
                OriginalAddress = $r.Start
                GeocodedAddress = if ($GeoStart) { $GeoStart.FormattedAddress } else { $null }
                GeocodeStatus   = $StartStatus
                MatchType       = if ($GeoStart) { $GeoStart.MatchType } else { 'NOT_FOUND' }
                PartialMatch    = if ($GeoStart) { [bool]$GeoStart.PartialMatch } else { $false }
                IsFallback      = $IsStartFallback
                Latitude        = if ($GeoStart) { $GeoStart.Latitude } else { $null }
                Longitude       = if ($GeoStart) { $GeoStart.Longitude } else { $null }
            })

            if ($GeoStart.Status -ne 'OK') {
                Write-Warning "  Błąd geokodowania startu: $($r.Start)"
                $ResultsList.Add([PSCustomObject]@{
                    Id               = [string]$CurrentIdx
                    Nazwa            = $RouteName
                    Start            = $r.Start
                    StartGeokodowany = $null
                    StartStatus      = $StartStatus
                    Koniec           = $r.End
                    KoniecGeokodowany= $null
                    EndStatus        = 'NOT_PROCESSED'
                    LiczbaPrzystankow= 0
                    TypTrasy         = $RowRouteType
                    OdlegloscKm      = $null
                    CzasMin          = $null
                    MapaPath         = $null
                    GoogleMapsUrl    = $null
                    Status           = "Błąd geokodowania startu: $($GeoStart.Status)"
                    Points           = @($RoutePoints)
                })
                continue
            }

            # Geokodowanie Celu
            $GeoEnd = Get-AddressCoordinates -Address $r.End -ApiKey $ApiKey
            $EndStatus = Get-GeocodeStatusDescription -Geo $GeoEnd
            $IsEndFallback = if ($GeoEnd -and ($GeoEnd.PartialMatch -or $GeoEnd.MatchType -in 'APPROXIMATE', 'GEOMETRIC_CENTER')) { $true } else { $false }

            if ($GeoEnd.Status -ne 'OK') {
                Write-Warning "  Błąd geokodowania celu: $($r.End)"
                $RoutePoints.Add([PSCustomObject]@{
                    Order           = 2
                    PointType       = 'End'
                    LegDistanceKm   = $null
                    LegDurationMin  = $null
                    OriginalAddress = $r.End
                    GeocodedAddress = if ($GeoEnd) { $GeoEnd.FormattedAddress } else { $null }
                    GeocodeStatus   = $EndStatus
                    MatchType       = if ($GeoEnd) { $GeoEnd.MatchType } else { 'NOT_FOUND' }
                    PartialMatch    = if ($GeoEnd) { [bool]$GeoEnd.PartialMatch } else { $false }
                    IsFallback      = $IsEndFallback
                    Latitude        = if ($GeoEnd) { $GeoEnd.Latitude } else { $null }
                    Longitude       = if ($GeoEnd) { $GeoEnd.Longitude } else { $null }
                })
                $ResultsList.Add([PSCustomObject]@{
                    Id               = [string]$CurrentIdx
                    Nazwa            = $RouteName
                    Start            = $r.Start
                    StartGeokodowany = $GeoStart.FormattedAddress
                    StartStatus      = $StartStatus
                    Koniec           = $r.End
                    KoniecGeokodowany= $null
                    EndStatus        = $EndStatus
                    LiczbaPrzystankow= 0
                    TypTrasy         = $RowRouteType
                    OdlegloscKm      = $null
                    CzasMin          = $null
                    MapaPath         = $null
                    GoogleMapsUrl    = $null
                    Status           = "Błąd geokodowania celu: $($GeoEnd.Status)"
                    Points           = @($RoutePoints)
                })
                continue
            }

            # Geocoding Waypoints
            $GeocodedWaypoints = [System.Collections.Generic.List[PSCustomObject]]::new()
            $WpIdx = 1
            if ($r.Waypoints -and $r.Waypoints.Count -gt 0) {
                foreach ($wp in $r.Waypoints) {
                    if ([string]::IsNullOrWhiteSpace($wp)) { continue }
                    $g = Get-AddressCoordinates -Address $wp -ApiKey $ApiKey
                    $wpStatus = Get-GeocodeStatusDescription -Geo $g
                    $isWpFallback = if ($g -and ($g.PartialMatch -or $g.MatchType -in 'APPROXIMATE', 'GEOMETRIC_CENTER')) { $true } else { $false }

                    $RoutePoints.Add([PSCustomObject]@{
                        Order           = ($WpIdx + 1)
                        PointType       = "Waypoint $WpIdx"
                        LegDistanceKm   = $null
                        LegDurationMin  = $null
                        OriginalAddress = $wp
                        GeocodedAddress = if ($g) { $g.FormattedAddress } else { $null }
                        GeocodeStatus   = $wpStatus
                        MatchType       = if ($g) { $g.MatchType } else { 'NOT_FOUND' }
                        PartialMatch    = if ($g) { [bool]$g.PartialMatch } else { $false }
                        IsFallback      = $isWpFallback
                        Latitude        = if ($g) { $g.Latitude } else { $null }
                        Longitude       = if ($g) { $g.Longitude } else { $null }
                    })

                    if ($g.Status -eq 'OK' -and $null -ne $g.Latitude -and $null -ne $g.Longitude) {
                        $GeocodedWaypoints.Add($g)
                    }
                    else {
                        Write-Warning "  Ostrzeżenie: pomijanie nieznanego punktu pośredniego '$wp' ($wpStatus)"
                    }
                    $WpIdx++
                    Start-Sleep -Milliseconds 100
                }
            }

            # Add End point to structured points
            $RoutePoints.Add([PSCustomObject]@{
                Order           = ($RoutePoints.Count + 1)
                PointType       = 'End'
                LegDistanceKm   = $null
                LegDurationMin  = $null
                OriginalAddress = $r.End
                GeocodedAddress = if ($GeoEnd) { $GeoEnd.FormattedAddress } else { $null }
                GeocodeStatus   = $EndStatus
                MatchType       = if ($GeoEnd) { $GeoEnd.MatchType } else { 'NOT_FOUND' }
                PartialMatch    = if ($GeoEnd) { [bool]$GeoEnd.PartialMatch } else { $false }
                IsFallback      = $IsEndFallback
                Latitude        = if ($GeoEnd) { $GeoEnd.Latitude } else { $null }
                Longitude       = if ($GeoEnd) { $GeoEnd.Longitude } else { $null }
            })

            # Obliczenie trasy
            $Trasa = Get-CarRouteData -OriginLat $GeoStart.Latitude -OriginLng $GeoStart.Longitude `
                -DestLat $GeoEnd.Latitude -DestLng $GeoEnd.Longitude `
                -IntermediatePoints $GeocodedWaypoints -RouteType $RowRouteType `
                -EmissionType $EmissionType -ApiKey $ApiKey -TrafficAware:$TrafficAware

            if ($Trasa.Status -ne 'OK') {
                Write-Warning "  Błąd Routes API: $($Trasa.Status)"
                $ResultsList.Add([PSCustomObject]@{
                    Id               = [string]$CurrentIdx
                    Nazwa            = $RouteName
                    Start            = $r.Start
                    StartGeokodowany = $GeoStart.FormattedAddress
                    StartStatus      = $StartStatus
                    Koniec           = $r.End
                    KoniecGeokodowany= $GeoEnd.FormattedAddress
                    EndStatus        = $EndStatus
                    LiczbaPrzystankow= $GeocodedWaypoints.Count
                    TypTrasy         = $RowRouteType
                    OdlegloscKm      = $null
                    CzasMin          = $null
                    MapaPath         = $null
                    GoogleMapsUrl    = $null
                    Status           = "Błąd trasy: $($Trasa.Status)"
                    Points           = @($RoutePoints)
                })
                continue
            }

            Write-Host "  OK: $($Trasa.OdlegloscKm) km, $($Trasa.CzasMin) min" -ForegroundColor Green

            if ($Trasa.Legs -and $Trasa.Legs.Count -gt 0) {
                for ($p = 0; $p -lt $RoutePoints.Count; $p++) {
                    if ($p -eq 0) {
                        $RoutePoints[$p].LegDistanceKm = 0
                        $RoutePoints[$p].LegDurationMin = 0
                    }
                    elseif (($p - 1) -lt $Trasa.Legs.Count) {
                        $leg = $Trasa.Legs[$p - 1]
                        $RoutePoints[$p].LegDistanceKm = $leg.DistanceKm
                        $RoutePoints[$p].LegDurationMin = $leg.DurationMin
                    }
                }
                for ($w = 0; $w -lt $GeocodedWaypoints.Count; $w++) {
                    if ($w -lt $Trasa.Legs.Count) {
                        $GeocodedWaypoints[$w] | Add-Member -NotePropertyName 'LegDistanceKm' -NotePropertyValue $Trasa.Legs[$w].DistanceKm -Force
                        $GeocodedWaypoints[$w] | Add-Member -NotePropertyName 'LegDurationMin' -NotePropertyValue $Trasa.Legs[$w].DurationMin -Force
                    }
                }
            }

            $GoogleMapsUrl = Get-GoogleMapsUrl -Origin "$($GeoStart.Latitude),$($GeoStart.Longitude)" `
                -Destination "$($GeoEnd.Latitude),$($GeoEnd.Longitude)" `
                -Waypoints $GeocodedWaypoints

            $MapPath = $null
            if ($GenerateMap -and $Trasa.EncodedPolyline) {
                $SafeName = ($RouteName -replace '[\\/:*?"<>|]', '_').Trim().Trim('.') -replace '\s+', ' '
                if ([string]::IsNullOrWhiteSpace($SafeName)) { $SafeName = "trasa_${CurrentIdx}" }
                $MapFileName = "${Timestamp}_trasa_${CurrentIdx}_${SafeName}.png"
                $MapPath = Join-Path $OutputFolder $MapFileName

                $AllRoutePoints = [System.Collections.Generic.List[PSCustomObject]]::new()
                $AllRoutePoints.Add($GeoStart)
                foreach ($wp in $GeocodedWaypoints) { $AllRoutePoints.Add($wp) }
                $AllRoutePoints.Add($GeoEnd)

                $null = Save-RouteMapPng -EncodedPolyline $Trasa.EncodedPolyline `
                    -OriginLat $GeoStart.Latitude -OriginLng $GeoStart.Longitude `
                    -DestLat $GeoEnd.Latitude -DestLng $GeoEnd.Longitude `
                    -RoutePoints $AllRoutePoints -OutputPath $MapPath -ApiKey $ApiKey `
                    -Width $MapWidth -Height $MapHeight `
                    -AddressTextA $GeoStart.FormattedAddress -AddressTextB $GeoEnd.FormattedAddress `
                    -DistanceText "$($Trasa.OdlegloscKm) km" -DurationText "$($Trasa.CzasMin) min" `
                    -HeaderLeftText $RouteName -RouteName $RouteName -HeaderRightText "Typ: $RowRouteType" `
                    -Legs $Trasa.Legs
            }

            $ResultsList.Add([PSCustomObject]@{
                Id               = [string]$CurrentIdx
                Nazwa            = $RouteName
                Start            = $r.Start
                StartGeokodowany = $GeoStart.FormattedAddress
                StartStatus      = $StartStatus
                Koniec           = $r.End
                KoniecGeokodowany= $GeoEnd.FormattedAddress
                EndStatus        = $EndStatus
                LiczbaPrzystankow= $GeocodedWaypoints.Count
                TypTrasy         = $RowRouteType
                OdlegloscKm      = $Trasa.OdlegloscKm
                CzasMin          = $Trasa.CzasMin
                MapaPath         = $MapPath
                GoogleMapsUrl    = $GoogleMapsUrl
                Status           = 'OK'
                Points           = @($RoutePoints)
            })
        }
        catch {
            Write-Warning "  Wyjątek dla trasy $($RouteName): $($_.Exception.Message)"
            $ResultsList.Add([PSCustomObject]@{
                Id               = [string]$CurrentIdx
                Nazwa            = $RouteName
                Start            = $r.Start
                StartGeokodowany = $null
                StartStatus      = 'EXCEPTION'
                Koniec           = $r.End
                KoniecGeokodowany= $null
                EndStatus        = 'EXCEPTION'
                LiczbaPrzystankow= 0
                TypTrasy         = $RowRouteType
                OdlegloscKm      = $null
                CzasMin          = $null
                MapaPath         = $null
                GoogleMapsUrl    = $null
                Status           = "Błąd: $($_.Exception.Message)"
                Points           = @($RoutePoints)
            })
        }

        Start-Sleep -Milliseconds 200
    }

    # Eksport raportu zbiorczego
    if ($ExportFormat -ne 'None' -and $ResultsList.Count -gt 0) {
        $BaseReportName = "${Timestamp}_podsumowanie_tras"
        if ($ExportFormat -in 'Excel', 'All') {
            $XlsxPath = Join-Path $OutputFolder "${BaseReportName}.xlsx"
            $saved = Export-RouteResults -Results $ResultsList -OutputPath $XlsxPath -Format Excel
            Write-Host "`nZapisano raport Excel: $saved" -ForegroundColor Cyan
        }
        if ($ExportFormat -in 'CSV', 'All') {
            $CsvPath = Join-Path $OutputFolder "${BaseReportName}.csv"
            $saved = Export-RouteResults -Results $ResultsList -OutputPath $CsvPath -Format CSV
            Write-Host "Zapisano raport CSV  : $saved" -ForegroundColor Cyan
        }
        if ($ExportFormat -in 'JSON', 'All') {
            $JsonPath = Join-Path $OutputFolder "${BaseReportName}.json"
            $saved = Export-RouteResults -Results $ResultsList -OutputPath $JsonPath -Format JSON
            Write-Host "Zapisano raport JSON : $saved" -ForegroundColor Cyan
        }
    }

    Write-Host "`nZakończono przetwarzanie $($ResultsList.Count) tras." -ForegroundColor Green
    return @($ResultsList)
}
