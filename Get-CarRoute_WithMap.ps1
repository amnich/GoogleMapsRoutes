#Requires -Version 5.1
<#
.SYNOPSIS
    Calculates shortest passenger car routes between addresses A and B from an Excel file.

.DESCRIPTION
    Loads an Excel file containing origin ("Adres A") and destination ("Adres B") columns,
    geocodes addresses via Google Geocoding API, computes the shortest driving route using
    Google Routes API v2, exports results (distance km, duration) to Excel, and downloads
    static PNG route maps via Google Static Maps API.

.PARAMETER ApiKey
    Google Maps API key. Default: resolved from GOOGLE_MAPS_API_KEY environment variable.

.PARAMETER InputExcel
    Path to input Excel file. If omitted, an interactive file picker dialog opens.

.PARAMETER OutputFolder
    Output directory for Excel results and PNG maps. Default: D:\!zrobic\

.PARAMETER KolumnaAdresA
    Column header name for origin address in Excel. Default: "Adres A".

.PARAMETER KolumnaAdresB
    Column header name for destination address in Excel. Default: "Adres B".

.PARAMETER MapWidth
    Rendered PNG map width in pixels. Default: 900.

.PARAMETER MapHeight
    Rendered PNG map height in pixels. Default: 600.

.EXAMPLE
    .\Get-CarRoute_WithMap.ps1
    Launches script with interactive file picker dialog.

.EXAMPLE
    .\Get-CarRoute_WithMap.ps1 -InputExcel "C:\addresses.xlsx" -OutputFolder "C:\results"
    Processes specified Excel file and saves results to target folder.

.NOTES
    Required modules: ImportExcel
    Required environment variable: GOOGLE_MAPS_API_KEY or -ApiKey parameter
    Encoding: UTF-8 with BOM
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ApiKey = $env:GOOGLE_MAPS_API_KEY,

    [Parameter(Mandatory = $false)]
    [string]$InputExcel,

    [Parameter(Mandatory = $false)]
    [string]$OutputFolder = 'C:\Temp',

    [Parameter(Mandatory = $false)]
    [string]$KolumnaAdresA = 'Adres A',

    [Parameter(Mandatory = $false)]
    [string]$KolumnaAdresB = 'Adres B',

    [Parameter(Mandatory = $false)]
    [ValidateRange(100, 2048)]
    [int]$MapWidth = 900,

    [Parameter(Mandatory = $false)]
    [ValidateRange(100, 2048)]
    [int]$MapHeight = 600
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Initialization
if ([string]::IsNullOrWhiteSpace($ApiKey)) {
    throw "Missing Google Maps API key. Set GOOGLE_MAPS_API_KEY environment variable or pass -ApiKey."
}

if (-not (Test-Path -Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
    Write-Verbose "Created output folder: $OutputFolder"
}

$DateStamp = Get-Date -Format 'yyyyMMdd_HHmm'

# ── Import shared functions (Get-AddressCoordinates, Get-CarRouteData, Save-RouteMapPng, Select-InputExcel) ──
. "$PSScriptRoot\RouteMapFunctions.ps1"

# Select input file
if ([string]::IsNullOrWhiteSpace($InputExcel)) {
    Write-Host "Opening Excel file selection dialog..." -ForegroundColor Cyan
    $InputExcel = Select-InputExcel
    if ([string]::IsNullOrWhiteSpace($InputExcel)) {
        Write-Warning "No file selected. Script terminated."
        exit 0
    }
}

if (-not (Test-Path -Path $InputExcel)) { throw "Input file does not exist: $InputExcel" }

Write-Host "Loading file: $InputExcel" -ForegroundColor Cyan

try {
    $Dane = Import-Excel -Path $InputExcel
}
catch {
    throw "Cannot load Excel file. Ensure ImportExcel module is installed. Error: $($_.Exception.Message)"
}

if ($null -eq $Dane -or @($Dane).Count -eq 0) {
    Write-Warning "Excel file is empty or contains no rows."
    exit 0
}

$Headers = $Dane[0].PSObject.Properties.Name
Write-Verbose "Columns in file: $($Headers -join ', ')"

$ColA = $Headers | Where-Object { $_ -like $KolumnaAdresA } | Select-Object -First 1
$ColB = $Headers | Where-Object { $_ -like $KolumnaAdresB } | Select-Object -First 1
if (-not $ColA) { $ColA = $Headers | Where-Object { $_ -match 'adres.*a$|^a$|adres_a|adresa' } | Select-Object -First 1 }
if (-not $ColB) { $ColB = $Headers | Where-Object { $_ -match 'adres.*b$|^b$|adres_b|adresb' } | Select-Object -First 1 }

if (-not $ColA -or -not $ColB) {
    Write-Warning "Available columns: $($Headers -join ', ')"
    throw "Cannot find address columns. Use -KolumnaAdresA and -KolumnaAdresB parameters."
}

Write-Host "Found columns: '$ColA' and '$ColB'" -ForegroundColor Green

$Wyniki = [System.Collections.Generic.List[PSCustomObject]]::new()
$RowIndex = 0
$TotalRows = @($Dane).Count

foreach ($Row in $Dane) {
    $RowIndex++
    $AdresA = ($Row.$ColA) -as [string]
    $AdresB = ($Row.$ColB) -as [string]

    Write-Host "[$RowIndex/$TotalRows] $AdresA  ->  $AdresB" -ForegroundColor Yellow

    $Wynik = [PSCustomObject]@{
        LP                  = $RowIndex
        AdresA              = $AdresA
        AdresA_Geokodowany  = $null
        AdresA_UlicaINumer  = $null
        AdresA_KodPocztowy  = $null
        AdresA_Miasto       = $null
        AdresA_TypDopasowania = $null
        AdresA_CzescioweDopasowanie = $null
        AdresB              = $AdresB
        AdresB_Geokodowany  = $null
        AdresB_UlicaINumer  = $null
        AdresB_KodPocztowy  = $null
        AdresB_Miasto       = $null
        AdresB_TypDopasowania = $null
        AdresB_CzescioweDopasowanie = $null
        OdlegloscKm         = $null
        CzasPodrozyMin      = $null
        PlikMapy            = $null
        StatusGeokodowaniaA = $null
        StatusGeokodowaniaB = $null
        StatusTrasy         = $null
        BladOpis            = $null
    }

    if ([string]::IsNullOrWhiteSpace($AdresA) -or [string]::IsNullOrWhiteSpace($AdresB)) {
        $Wynik.StatusTrasy = 'PUSTE_DANE'
        $Wynik.BladOpis = 'Adres A lub B jest pusty.'
        Write-Warning "Wiersz $RowIndex`: Pusty adres - pomijam."
        $Wyniki.Add($Wynik)
        continue
    }

    # Geocode Address A
    $GeoA = Get-AddressCoordinates -Address $AdresA -ApiKey $ApiKey -RequireStreetNumber
    $Wynik.StatusGeokodowaniaA = $GeoA.Status
    $Wynik.AdresA_Geokodowany  = $GeoA.FormattedAddress
    $Wynik.AdresA_UlicaINumer  = $GeoA.UlicaINumer
    $Wynik.AdresA_KodPocztowy  = $GeoA.KodPocztowy
    $Wynik.AdresA_Miasto       = $GeoA.Miasto
    $Wynik.AdresA_TypDopasowania = $GeoA.MatchType
    $Wynik.AdresA_CzescioweDopasowanie = $GeoA.PartialMatch
    if ($GeoA.Status -ne 'OK') {
        $Wynik.StatusTrasy = 'BLAD_GEOCODING_A'
        $Wynik.BladOpis = "Nieprecyzyjny lub nieznany adres A: $($GeoA.Status); wynik: $($GeoA.FormattedAddress); typ: $($GeoA.MatchType)"
        Write-Warning "  Blad geokodowania A: $($Wynik.BladOpis)"
        $Wyniki.Add($Wynik); Start-Sleep -Milliseconds 200; continue
    }

    # Geocode Address B
    $GeoB = Get-AddressCoordinates -Address $AdresB -ApiKey $ApiKey -RequireStreetNumber
    $Wynik.StatusGeokodowaniaB = $GeoB.Status
    $Wynik.AdresB_Geokodowany  = $GeoB.FormattedAddress
    $Wynik.AdresB_UlicaINumer  = $GeoB.UlicaINumer
    $Wynik.AdresB_KodPocztowy  = $GeoB.KodPocztowy
    $Wynik.AdresB_Miasto       = $GeoB.Miasto
    $Wynik.AdresB_TypDopasowania = $GeoB.MatchType
    $Wynik.AdresB_CzescioweDopasowanie = $GeoB.PartialMatch
    if ($GeoB.Status -ne 'OK') {
        $Wynik.StatusTrasy = 'BLAD_GEOCODING_B'
        $Wynik.BladOpis = "Nieprecyzyjny lub nieznany adres B: $($GeoB.Status); wynik: $($GeoB.FormattedAddress); typ: $($GeoB.MatchType)"
        Write-Warning "  Blad geokodowania B: $($Wynik.BladOpis)"
        $Wyniki.Add($Wynik); Start-Sleep -Milliseconds 200; continue
    }

    Write-Verbose "  GeoA: $($GeoA.FormattedAddress) [$($GeoA.Latitude), $($GeoA.Longitude)]"
    Write-Verbose "  GeoB: $($GeoB.FormattedAddress) [$($GeoB.Latitude), $($GeoB.Longitude)]"

    # Trasa
    $Trasa = Get-CarRouteData -OriginLat $GeoA.Latitude -OriginLng $GeoA.Longitude `
        -DestLat   $GeoB.Latitude -DestLng   $GeoB.Longitude `
        -ApiKey    $ApiKey

    $Wynik.StatusTrasy = $Trasa.Status
    if ($Trasa.Status -ne 'OK') {
        $Wynik.BladOpis = "Brak trasy: $($Trasa.Status)"
        Write-Warning "  Brak trasy: $($Trasa.Status)"
        $Wyniki.Add($Wynik); Start-Sleep -Milliseconds 250; continue
    }

    $Wynik.OdlegloscKm = $Trasa.OdlegloscKm
    $Wynik.CzasPodrozyMin = $Trasa.CzasMin
    Write-Host "  OK  Odleglosc: $($Trasa.OdlegloscKm) km | Czas: $($Trasa.CzasMin) min" -ForegroundColor Green

    # Mapa PNG
    if ($Trasa.EncodedPolyline) {
        $PngFileName = "${DateStamp}_trasa_${RowIndex}.png"
        $PngPath = Join-Path -Path $OutputFolder -ChildPath $PngFileName
        $OdlTekst = if ($Trasa.OdlegloscKm) { "$($Trasa.OdlegloscKm) km" } else { '' }
        $TekstA_Mapa = if ($Wynik.AdresA_Geokodowany) { $Wynik.AdresA_Geokodowany } else { $AdresA }
        $TekstB_Mapa = if ($Wynik.AdresB_Geokodowany) { $Wynik.AdresB_Geokodowany } else { $AdresB }
        $MapSaved = Save-RouteMapPng -EncodedPolyline $Trasa.EncodedPolyline `
            -OriginLat $GeoA.Latitude -OriginLng $GeoA.Longitude `
            -DestLat   $GeoB.Latitude -DestLng   $GeoB.Longitude `
            -OutputPath $PngPath -ApiKey $ApiKey -Width $MapWidth -Height $MapHeight `
            -AddressTextA   $TekstA_Mapa `
            -AddressTextB   $TekstB_Mapa `
            -DistanceText   $OdlTekst
        if ($MapSaved) {
            $Wynik.PlikMapy = $PngPath
            Write-Host "  Mapa: $PngFileName" -ForegroundColor Cyan
        }
    }

    $Wyniki.Add($Wynik)
    Start-Sleep -Milliseconds 250
}

# Eksport do Excel
$OutputExcelPath = Join-Path -Path $OutputFolder -ChildPath "${DateStamp}_trasy_samochodowe.xlsx"

if ($Wyniki.Count -gt 0) {
    Write-Host "`nEksportowanie wynikow do: $OutputExcelPath" -ForegroundColor Cyan
    $Wyniki | Export-Excel -Path $OutputExcelPath -WorksheetName 'Trasy' -TableName 'TrasySamochodowe' `
        -AutoSize -AutoFilter -ClearSheet -NoNumberConversion * -FreezeTopRow
    $Dane | Export-Excel -Path $OutputExcelPath -WorksheetName 'DaneWejsciowe' -TableName 'DaneWejsciowe' `
        -AutoSize -AutoFilter -ClearSheet -NoNumberConversion *
    Write-Host "Wyniki zapisano: $OutputExcelPath" -ForegroundColor Green
}
else {
    Write-Warning "Brak wynikow do zapisania."
}

# Podsumowanie
$OK = @($Wyniki | Where-Object { $_.StatusTrasy -eq 'OK' }).Count
$Bledy = @($Wyniki | Where-Object { $_.StatusTrasy -ne 'OK' }).Count
$Mapy = @($Wyniki | Where-Object { $null -ne $_.PlikMapy }).Count

Write-Host "`n================================================" -ForegroundColor Magenta
Write-Host " PODSUMOWANIE" -ForegroundColor Magenta
Write-Host "================================================" -ForegroundColor Magenta
Write-Host " Przetworzone wiersze : $($Wyniki.Count)" -ForegroundColor White
Write-Host " Trasy OK             : $OK" -ForegroundColor Green
Write-Host " Bledy                : $Bledy" -ForegroundColor Red
Write-Host " Wygenerowane mapy    : $Mapy" -ForegroundColor Cyan
Write-Host " Plik Excel           : $OutputExcelPath" -ForegroundColor White
Write-Host " Folder wynikowy      : $OutputFolder" -ForegroundColor White
Write-Host "================================================`n" -ForegroundColor Magenta

$Wyniki