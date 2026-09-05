#Requires -Version 5.1
<#
.SYNOPSIS
    Generates a single multi-stop driving route through multiple waypoints from an Excel file.

.DESCRIPTION
    Reads points from columns LP, Miejscowosc, and pickup/dropoff location.
    Stops are visited in order of LP. The script geocodes addresses, computes
    total distance and duration, generates a PNG route map, an Excel summary,
    and an interactive Google Maps URL.

.PARAMETER ApiKey
    Google Maps API key. Default: resolved from GOOGLE_MAPS_API_KEY environment variable.

.PARAMETER InputExcel
    Path to input Excel file. If omitted, an interactive file picker dialog opens.

.PARAMETER OutputFolder
    Output directory for generated route maps and summaries. Default: C:\Temp\MultiPointRoute

.PARAMETER KolumnaLP
    Column header name for stop sequence number in Excel. Default: "LP".

.PARAMETER KolumnaMiejscowosc
    Column header name for locality/city in Excel. Default: "Miejscowosc".

.PARAMETER KolumnaLokalizacja
    Column header name for stop location description in Excel. Default: "Lokalizacja miejsca odbioru lub dowozu".

.PARAMETER Kraj
    Country name appended to geocoding queries. Default: "Polska".

.PARAMETER MapWidth
    Rendered PNG map width in pixels (100-2048). Default: 900.

.PARAMETER MapHeight
    Rendered PNG map height in pixels (100-2048). Default: 600.

.EXAMPLE
    .\Get-MultiPointCarRoute_WithMap.ps1 -InputExcel "C:\Temp\points.xlsx"

.NOTES
    Required module: ImportExcel
    Required environment variable: GOOGLE_MAPS_API_KEY or -ApiKey parameter.
    Google Routes API supports up to 25 intermediate waypoints (27 points total).
    Encoding: UTF-8 with BOM
#>
[CmdletBinding()]
param(
    [Parameter()][string]$ApiKey = $env:GOOGLE_MAPS_API_KEY,
    [Parameter()][string]$InputExcel,
    [Parameter()][string]$OutputFolder = 'C:\Temp\MultiPointRoute',
    [Parameter()][string]$KolumnaLP = 'LP',
    [Parameter()][string]$KolumnaMiejscowosc = 'Miejscowosc',
    [Parameter()][string]$KolumnaLokalizacja = 'Lokalizacja miejsca odbioru lub dowozu',
    [Parameter()][string]$Kraj = 'Polska',
    [Parameter()][ValidateRange(100, 2048)][int]$MapWidth = 900,
    [Parameter()][ValidateRange(100, 2048)][int]$MapHeight = 600
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($ApiKey)) {
    throw 'Missing Google Maps API key. Set GOOGLE_MAPS_API_KEY environment variable or pass -ApiKey parameter.'
}

. "$PSScriptRoot\RouteMapFunctions.ps1"

if ([string]::IsNullOrWhiteSpace($InputExcel)) {
    Write-Host 'Opening Excel file selection dialog...' -ForegroundColor Cyan
    $InputExcel = Select-InputExcel
    if ([string]::IsNullOrWhiteSpace($InputExcel)) {
        Write-Warning 'No file selected. Script terminated.'
        exit 0
    }
}

if (-not (Test-Path -LiteralPath $InputExcel -PathType Leaf)) {
    throw "Input file does not exist: $InputExcel"
}
if (-not (Test-Path -LiteralPath $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}

try {
    $Dane = @(Import-Excel -Path $InputExcel)
}
catch {
    throw "Nie mozna wczytac pliku Excel. Sprawdz modul ImportExcel. Blad: $($_.Exception.Message)"
}
if ($Dane.Count -eq 0) {
    throw 'Plik Excel jest pusty.'
}

function Find-RouteColumn {
    param(
        [Parameter(Mandatory)][string[]]$Headers,
        [Parameter(Mandatory)][string]$PreferredName,
        [Parameter(Mandatory)][string]$FallbackPattern
    )

    $Column = $Headers | Where-Object { $_ -eq $PreferredName } | Select-Object -First 1
    if (-not $Column) {
        $Column = $Headers | Where-Object { $_ -match $FallbackPattern } | Select-Object -First 1
    }
    return $Column
}

$Headers = @($Dane[0].PSObject.Properties.Name)
$ColLP = Find-RouteColumn -Headers $Headers -PreferredName $KolumnaLP -FallbackPattern '(?i)^\s*(lp|l\.p\.|kolejnosc|numer)\s*$'
$ColMiejscowosc = Find-RouteColumn -Headers $Headers -PreferredName $KolumnaMiejscowosc -FallbackPattern '(?i)^\s*(miejscowo[sś][cć]|miasto)\s*$'
$ColLokalizacja = Find-RouteColumn -Headers $Headers -PreferredName $KolumnaLokalizacja -FallbackPattern '(?i)(lokalizacja|miejsce).*(odbioru|dowozu)|^\s*(adres|punkt)\s*$'

if (-not $ColLP -or -not $ColMiejscowosc -or -not $ColLokalizacja) {
    throw "Nie znaleziono wymaganych kolumn. Dostepne kolumny: $($Headers -join ', ')"
}

$RowsOrdered = @($Dane | Sort-Object { [int]($_.$ColLP) })
if ($RowsOrdered.Count -lt 2) {
    throw 'Trasa musi zawierac co najmniej 2 punkty.'
}
if ($RowsOrdered.Count -gt 27) {
    throw "Trasa zawiera $($RowsOrdered.Count) punktow. Maksymalna liczba to 27."
}

$GeocodedPoints = [System.Collections.Generic.List[PSCustomObject]]::new()
foreach ($Row in $RowsOrdered) {
    $LP = [string]$Row.$ColLP
    $Miejscowosc = ([string]$Row.$ColMiejscowosc).Trim()
    $Lokalizacja = ([string]$Row.$ColLokalizacja).Trim()
    if ([string]::IsNullOrWhiteSpace($Miejscowosc) -or [string]::IsNullOrWhiteSpace($Lokalizacja)) {
        throw "Punkt LP $LP ma pusta miejscowosc lub lokalizacje."
    }

    $Address = "$Lokalizacja, $Miejscowosc"
    if (-not [string]::IsNullOrWhiteSpace($Kraj)) {
        $Address += ", $Kraj"
    }

    Write-Host "[$LP/$($RowsOrdered.Count)] Geokodowanie: $Address" -ForegroundColor Yellow
    $Geo = Get-AddressCoordinates -Address $Address -ApiKey $ApiKey -RequireStreetNumber -Verbose:$false
    $GeoColor = if ($Geo.Status -eq 'OK') { 'Green' } else { 'Red' }
    Write-Host "  Geokodowanie status: $($Geo.Status), Typ: $($Geo.MatchType), Adres: $($Geo.FormattedAddress), Wspolrzedne: $($Geo.Latitude),$($Geo.Longitude)" -ForegroundColor $GeoColor
    $GeocodedPoints.Add([PSCustomObject]@{
            LP                 = $LP
            Miejscowosc        = $Miejscowosc
            Lokalizacja        = $Lokalizacja
            ZapytanieAdresowe  = $Address
            AdresGeokodowany   = $Geo.FormattedAddress
            Latitude           = $Geo.Latitude
            Longitude          = $Geo.Longitude
            TypDopasowania     = $Geo.MatchType
            CzescioweDopasowanie = $Geo.PartialMatch
            StatusGeokodowania = $Geo.Status
        })
    Start-Sleep -Milliseconds 200
}

$FailedPoints = @($GeocodedPoints | Where-Object { $_.StatusGeokodowania -ne 'OK' })
$DateStamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$OutputExcel = Join-Path $OutputFolder "${DateStamp}_trasa_wielopunktowa.xlsx"

if ($FailedPoints.Count -gt 0) {
    $GeocodedPoints | Export-Excel -Path $OutputExcel -WorksheetName 'Punkty' -TableName 'PunktyTrasy' -AutoSize -AutoFilter -FreezeTopRow
    throw "Nie udalo sie geokodowanie $($FailedPoints.Count) punktow. Szczegoly zapisano w: $OutputExcel"
}

$Origin = $GeocodedPoints[0]
$Destination = $GeocodedPoints[$GeocodedPoints.Count - 1]
$IntermediatePoints = if ($GeocodedPoints.Count -gt 2) {
    @($GeocodedPoints[1..($GeocodedPoints.Count - 2)])
}
else {
    @()
}

Write-Host "Obliczanie trasy przez $($GeocodedPoints.Count) punktow..." -ForegroundColor Cyan
$Route = Get-CarRouteData -OriginLat $Origin.Latitude -OriginLng $Origin.Longitude `
    -DestLat $Destination.Latitude -DestLng $Destination.Longitude `
    -IntermediatePoints $IntermediatePoints -ApiKey $ApiKey
if ($Route.Status -ne 'OK') {
    throw "Nie udalo sie obliczyc trasy: $($Route.Status)"
}

$MapPath = Join-Path $OutputFolder "${DateStamp}_trasa_wielopunktowa.png"
$MapSaved = Save-RouteMapPng -EncodedPolyline $Route.EncodedPolyline `
    -OriginLat $Origin.Latitude -OriginLng $Origin.Longitude `
    -DestLat $Destination.Latitude -DestLng $Destination.Longitude `
    -RoutePoints $GeocodedPoints -OutputPath $MapPath -ApiKey $ApiKey `
    -Width $MapWidth -Height $MapHeight `
    -AddressTextA $Origin.AdresGeokodowany -AddressTextB $Destination.AdresGeokodowany `
    -DistanceText "$($Route.OdlegloscKm) km"

$OriginValue = [System.Uri]::EscapeDataString("$($Origin.Latitude),$($Origin.Longitude)")
$DestinationValue = [System.Uri]::EscapeDataString("$($Destination.Latitude),$($Destination.Longitude)")
$GoogleMapsUrl = "https://www.google.com/maps/dir/?api=1&origin=$OriginValue&destination=$DestinationValue&travelmode=driving"
if ($IntermediatePoints.Count -gt 0) {
    $WaypointValues = @($IntermediatePoints | ForEach-Object { "$($_.Latitude),$($_.Longitude)" }) -join '|'
    $GoogleMapsUrl += '&waypoints=' + [System.Uri]::EscapeDataString($WaypointValues)
}

$Summary = [PSCustomObject]@{
    LiczbaPunktow = $GeocodedPoints.Count
    OdlegloscKm   = $Route.OdlegloscKm
    CzasMin       = $Route.CzasMin
    PlikMapy      = if ($MapSaved) { $MapPath } else { $null }
    GoogleMapsUrl = $GoogleMapsUrl
    Status        = 'OK'
}

$Summary | Export-Excel -Path $OutputExcel -WorksheetName 'Podsumowanie' -TableName 'PodsumowanieTrasy' -AutoSize -FreezeTopRow
$GeocodedPoints | Export-Excel -Path $OutputExcel -WorksheetName 'Punkty' -TableName 'PunktyTrasy' -AutoSize -AutoFilter -FreezeTopRow
$Dane | Export-Excel -Path $OutputExcel -WorksheetName 'DaneWejsciowe' -TableName 'DaneWejsciowe' -AutoSize -AutoFilter
$GoogleMapsUrl | Set-Content -LiteralPath (Join-Path $OutputFolder "${DateStamp}_GoogleMaps.url.txt") -Encoding UTF8

Write-Host "`nTrasa gotowa: $($Route.OdlegloscKm) km, $($Route.CzasMin) min" -ForegroundColor Green
Write-Host "Mapa:  $MapPath" -ForegroundColor Cyan
Write-Host "Excel: $OutputExcel" -ForegroundColor Cyan
Write-Host "Google Maps: $GoogleMapsUrl" -ForegroundColor Cyan

$Summary