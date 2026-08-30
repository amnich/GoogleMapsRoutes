#Requires -Version 5.1
<#
.SYNOPSIS
    Oblicza najkrótszą trasę samochodem osobowym pomiędzy adresami A i B z pliku Excel.

.DESCRIPTION
    Skrypt wczytuje plik Excel z kolumnami "Adres A" i "Adres B", geokoduje adresy
    przez Google Geocoding API, oblicza najkrótszą trasę (samochód osobowy) przez
    Google Routes API v2, zapisuje wyniki (odległość km, czas) do pliku Excel oraz
    generuje mapę PNG dla każdej trasy przez Google Static Maps API.

.PARAMETER ApiKey
    Klucz Google Maps API. Domyślnie pobierany ze zmiennej środowiskowej GOOGLE_MAPS_API_KEY.

.PARAMETER InputExcel
    Ścieżka do pliku Excel wejściowego. Jeśli nie podano, otwiera się dialog wyboru pliku.

.PARAMETER OutputFolder
    Folder wynikowy dla Excel i PNG. Domyślnie D:\!zrobic\

.PARAMETER KolumnaAdresA
    Nazwa kolumny z adresem A w pliku Excel. Domyślnie "Adres A".

.PARAMETER KolumnaAdresB
    Nazwa kolumny z adresem B w pliku Excel. Domyślnie "Adres B".

.PARAMETER MapWidth
    Szerokość mapy PNG w pikselach. Domyślnie 900.

.PARAMETER MapHeight
    Wysokość mapy PNG w pikselach. Domyślnie 600.

.EXAMPLE
    .\Get-CarRoute_WithMap.ps1
    Uruchamia skrypt z dialogiem wyboru pliku.

.EXAMPLE
    .\Get-CarRoute_WithMap.ps1 -InputExcel "C:\adresy.xlsx" -OutputFolder "C:\wyniki"
    Przetwarza podany plik Excel i zapisuje wyniki w podanym folderze.

.NOTES
    Wymagane moduły: ImportExcel
    Wymagana zmienna środowiskowa: GOOGLE_MAPS_API_KEY lub parametr -ApiKey
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

# Inicjalizacja
if ([string]::IsNullOrWhiteSpace($ApiKey)) {
    throw "Brak klucza Google Maps API. Ustaw zmienną środowiskową GOOGLE_MAPS_API_KEY lub podaj parametr -ApiKey."
}

if (-not (Test-Path -Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
    Write-Verbose "Utworzono folder wynikowy: $OutputFolder"
}

$DateStamp = Get-Date -Format 'yyyyMMdd_HHmm'

# ── Import wspólnych funkcji (Get-AddressCoordinates, Get-CarRouteData, Save-RouteMapPng, Select-InputExcel) ──
. "$PSScriptRoot\RouteMapFunctions.ps1"

# Wybor pliku wejsciowego
if ([string]::IsNullOrWhiteSpace($InputExcel)) {
    Write-Host "Otwieranie dialogu wyboru pliku Excel..." -ForegroundColor Cyan
    $InputExcel = Select-InputExcel
    if ([string]::IsNullOrWhiteSpace($InputExcel)) {
        Write-Warning "Nie wybrano pliku. Skrypt zakonczony."
        exit 0
    }
}

if (-not (Test-Path -Path $InputExcel)) { throw "Plik wejsciowy nie istnieje: $InputExcel" }

Write-Host "Wczytywanie pliku: $InputExcel" -ForegroundColor Cyan

try {
    $Dane = Import-Excel -Path $InputExcel
}
catch {
    throw "Nie mozna wczytac pliku Excel. Upewnij sie ze modul ImportExcel jest zainstalowany. Blad: $($_.Exception.Message)"
}

if ($null -eq $Dane -or @($Dane).Count -eq 0) {
    Write-Warning "Plik Excel jest pusty lub nie zawiera danych."
    exit 0
}

$Headers = $Dane[0].PSObject.Properties.Name
Write-Verbose "Kolumny w pliku: $($Headers -join ', ')"

$ColA = $Headers | Where-Object { $_ -like $KolumnaAdresA } | Select-Object -First 1
$ColB = $Headers | Where-Object { $_ -like $KolumnaAdresB } | Select-Object -First 1
if (-not $ColA) { $ColA = $Headers | Where-Object { $_ -match 'adres.*a$|^a$|adres_a|adresa' } | Select-Object -First 1 }
if (-not $ColB) { $ColB = $Headers | Where-Object { $_ -match 'adres.*b$|^b$|adres_b|adresb' } | Select-Object -First 1 }

if (-not $ColA -or -not $ColB) {
    Write-Warning "Dostepne kolumny: $($Headers -join ', ')"
    throw "Nie mozna odnalezc kolumn adresowych. Uzyj parametrow -KolumnaAdresA i -KolumnaAdresB."
}

Write-Host "Znalezione kolumny: '$ColA' i '$ColB'" -ForegroundColor Green

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
        AdresB              = $AdresB
        AdresB_Geokodowany  = $null
        AdresB_UlicaINumer  = $null
        AdresB_KodPocztowy  = $null
        AdresB_Miasto       = $null
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

    # Geokodowanie A
    $GeoA = Get-AddressCoordinates -Address $AdresA -ApiKey $ApiKey
    $Wynik.StatusGeokodowaniaA = $GeoA.Status
    $Wynik.AdresA_Geokodowany  = $GeoA.FormattedAddress
    $Wynik.AdresA_UlicaINumer  = $GeoA.UlicaINumer
    $Wynik.AdresA_KodPocztowy  = $GeoA.KodPocztowy
    $Wynik.AdresA_Miasto       = $GeoA.Miasto
    if ($GeoA.Status -ne 'OK') {
        $Wynik.StatusTrasy = 'BLAD_GEOCODING_A'
        $Wynik.BladOpis = "Nie mozna znalezc adresu A: $($GeoA.Status)"
        Write-Warning "  Blad geokodowania A: $($GeoA.Status)"
        $Wyniki.Add($Wynik); Start-Sleep -Milliseconds 200; continue
    }

    # Geokodowanie B
    $GeoB = Get-AddressCoordinates -Address $AdresB -ApiKey $ApiKey
    $Wynik.StatusGeokodowaniaB = $GeoB.Status
    $Wynik.AdresB_Geokodowany  = $GeoB.FormattedAddress
    $Wynik.AdresB_UlicaINumer  = $GeoB.UlicaINumer
    $Wynik.AdresB_KodPocztowy  = $GeoB.KodPocztowy
    $Wynik.AdresB_Miasto       = $GeoB.Miasto
    if ($GeoB.Status -ne 'OK') {
        $Wynik.StatusTrasy = 'BLAD_GEOCODING_B'
        $Wynik.BladOpis = "Nie mozna znalezc adresu B: $($GeoB.Status)"
        Write-Warning "  Blad geokodowania B: $($GeoB.Status)"
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
            -TekstAdresA    $TekstA_Mapa `
            -TekstAdresB    $TekstB_Mapa `
            -TekstOdleglosc $OdlTekst
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