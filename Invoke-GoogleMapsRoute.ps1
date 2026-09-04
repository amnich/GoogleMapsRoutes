#Requires -Version 5.1
<#
.SYNOPSIS
    Oblicza trasę Google Maps dla parametrów ręcznych lub pliku danych (JSON, CSV, Excel).

.DESCRIPTION
    Skrypt umożliwia:
      1. Ręczne podanie punktu początkowego, końcowego oraz opcjonalnych punktów pośrednich
      2. Wczytanie pliku danych (JSON, CSV, Excel) z listą tras lub sekwencją przystanków
      3. Wybór optymalizacji: Najszybsza (Fastest), Najkrótsza (Shortest), Ekologiczna (Eco / Fuel Efficient)
      4. Geokodowanie adresów przez Google Geocoding API
      5. Obliczanie trasy przez Google Routes API v2
      6. Generowanie mapy PNG z naniesioną trasą i znacznikami
      7. Generowanie interaktywnego linku Google Maps
      8. Zapisywanie raportów zbiorczych (Excel, CSV, JSON)

.PARAMETER StartPoint
    Adres lub współrzędne punktu startowego (np. "Warszawa, Marszałkowska 1" lub "52.2297, 21.0122").

.PARAMETER EndPoint
    Adres lub współrzędne punktu docelowego.

.PARAMETER Waypoints
    Lista adresów lub współrzędnych punktów pośrednich (do 25 punktów).

.PARAMETER RouteType
    Typ trasy:
      - 'Fastest': Najszybsza czasowo
      - 'Shortest': Najkrótsza pod kątem odległości (km)
      - 'Eco': Ekologiczna / najniższe zużycie paliwa/energii

.PARAMETER EmissionType
    Typ napędu dla trasy Eco: 'GASOLINE', 'DIESEL', 'HYBRID', 'ELECTRIC'. Domyślnie 'GASOLINE'.

.PARAMETER InputFile
    Ścieżka do pliku wejściowego (.xlsx, .xls, .csv, .tsv, .json).

.PARAMETER ApiKey
    Klucz Google Maps API. Jeśli nie podano, pobierany ze zmiennej GOOGLE_MAPS_API_KEY lub DPAPI.

.PARAMETER OutputFolder
    Folder zapisu wygenerowanych map PNG i raportów. Domyślnie .\Results

.PARAMETER GenerateMap
    Przełącznik określający czy generować mapę statyczną PNG. Domyślnie włączony ($true).

.PARAMETER OpenBrowser
    Otwiera wygenerowaną trasę w przeglądarce internetowej.

.EXAMPLE
    # Trasa z punktami pośrednimi (Najszybsza)
    .\Invoke-GoogleMapsRoute.ps1 -StartPoint "Warszawa, Marszałkowska 1" -EndPoint "Kraków, Rynek Główny 1" `
        -Waypoints "Radom, Żeromskiego 5", "Kielce, Sienkiewicza 10" -RouteType Fastest -GenerateMap

.EXAMPLE
    # Trasa najkrótsza
    .\Invoke-GoogleMapsRoute.ps1 -StartPoint "Gdańsk, Długa 1" -EndPoint "Toruń, Szeroka 1" -RouteType Shortest

.EXAMPLE
    # Przetwarzanie wsadowe pliku Excel lub JSON
    .\Invoke-GoogleMapsRoute.ps1 -InputFile ".\Samples\routes_sample.xlsx" -RouteType Fastest -ExportFormat Excel

.NOTES
    Encoding: UTF-8 with BOM
#>

[CmdletBinding(DefaultParameterSetName = 'Manual')]
param(
    # --- Zestaw parametrów: Manual ---
    [Parameter(Mandatory = $true, ParameterSetName = 'Manual', Position = 0)]
    [string]$StartPoint,

    [Parameter(Mandatory = $true, ParameterSetName = 'Manual', Position = 1)]
    [string]$EndPoint,

    [Parameter(Mandatory = $false, ParameterSetName = 'Manual')]
    [string[]]$Waypoints = @(),

    [Parameter(Mandatory = $false, ParameterSetName = 'Manual')]
    [string]$Name = 'Trasa manualna',

    # --- Zestaw parametrów: File ---
    [Parameter(Mandatory = $true, ParameterSetName = 'File', Position = 0)]
    [string]$InputFile,

    [Parameter(Mandatory = $false, ParameterSetName = 'File')]
    [ValidateSet('Excel', 'CSV', 'JSON', 'All', 'None')]
    [string]$ExportFormat = 'Excel',

    # --- Parametry wspólne ---
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

# Załadowanie wspólnych funkcji
$FunctionsPath = Join-Path $PSScriptRoot 'RouteMapFunctions.ps1'
if (-not (Test-Path $FunctionsPath)) {
    throw "Nie odnaleziono pliku modułu funkcji: $FunctionsPath"
}
. $FunctionsPath

# Ustalenie klucza API
if ([string]::IsNullOrWhiteSpace($ApiKey)) {
    # Próba odczytania z konfiguracji lokalnej DPAPI
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

# Ustalenie folderu wyjściowego
if ([string]::IsNullOrWhiteSpace($OutputFolder)) {
    $OutputFolder = Join-Path $PSScriptRoot 'Results'
}
if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}

$Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'

# ══════════════════════════════════════════════════════════════════════════════
# TRYB 1: RĘCZNE PARAMETRY (MANUAL)
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

    # Geokodowanie Punktów Pośrednich
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
            -TekstAdresA $GeoStart.FormattedAddress -TekstAdresB $GeoEnd.FormattedAddress `
            -TekstOdleglosc "$($Trasa.OdlegloscKm) km" -TekstCzas "$($Trasa.CzasMin) min" `
            -TekstNaglowekLewy $Name -TekstNaglowekPrawy "Typ: $ActualRouteType"

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

        # Wybór RouteType dla wiersza
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
            if ($GeoStart.Status -ne 'OK') {
                Write-Warning "  Błąd geokodowania startu: $($r.Start)"
                $ResultsList.Add([PSCustomObject]@{
                    Id               = [string]$CurrentIdx
                    Nazwa            = $RouteName
                    Start            = $r.Start
                    StartGeokodowany = $null
                    Koniec           = $r.End
                    KoniecGeokodowany= $null
                    LiczbaPrzystankow= 0
                    TypTrasy         = $RowRouteType
                    OdlegloscKm      = $null
                    CzasMin          = $null
                    MapaPath         = $null
                    GoogleMapsUrl    = $null
                    Status           = "Błąd geokodowania startu: $($GeoStart.Status)"
                })
                continue
            }

            # Geokodowanie Celu
            $GeoEnd = Get-AddressCoordinates -Address $r.End -ApiKey $ApiKey
            if ($GeoEnd.Status -ne 'OK') {
                Write-Warning "  Błąd geokodowania celu: $($r.End)"
                $ResultsList.Add([PSCustomObject]@{
                    Id               = [string]$CurrentIdx
                    Nazwa            = $RouteName
                    Start            = $r.Start
                    StartGeokodowany = $GeoStart.FormattedAddress
                    Koniec           = $r.End
                    KoniecGeokodowany= $null
                    LiczbaPrzystankow= 0
                    TypTrasy         = $RowRouteType
                    OdlegloscKm      = $null
                    CzasMin          = $null
                    MapaPath         = $null
                    GoogleMapsUrl    = $null
                    Status           = "Błąd geokodowania celu: $($GeoEnd.Status)"
                })
                continue
            }

            # Geokodowanie Punktów Pośrednich
            $GeocodedWaypoints = [System.Collections.Generic.List[PSCustomObject]]::new()
            if ($r.Waypoints -and $r.Waypoints.Count -gt 0) {
                foreach ($wp in $r.Waypoints) {
                    if ([string]::IsNullOrWhiteSpace($wp)) { continue }
                    $g = Get-AddressCoordinates -Address $wp -ApiKey $ApiKey
                    if ($g.Status -eq 'OK') {
                        $GeocodedWaypoints.Add($g)
                    }
                    else {
                        Write-Warning "  Ostrzeżenie: pomijanie nieznanego punktu pośredniego '$wp'"
                    }
                    Start-Sleep -Milliseconds 100
                }
            }

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
                    Koniec           = $r.End
                    KoniecGeokodowany= $GeoEnd.FormattedAddress
                    LiczbaPrzystankow= $GeocodedWaypoints.Count
                    TypTrasy         = $RowRouteType
                    OdlegloscKm      = $null
                    CzasMin          = $null
                    MapaPath         = $null
                    GoogleMapsUrl    = $null
                    Status           = "Błąd trasy: $($Trasa.Status)"
                })
                continue
            }

            Write-Host "  OK: $($Trasa.OdlegloscKm) km, $($Trasa.CzasMin) min" -ForegroundColor Green

            $GoogleMapsUrl = Get-GoogleMapsUrl -Origin "$($GeoStart.Latitude),$($GeoStart.Longitude)" `
                -Destination "$($GeoEnd.Latitude),$($GeoEnd.Longitude)" `
                -Waypoints $GeocodedWaypoints

            $MapPath = $null
            if ($GenerateMap -and $Trasa.EncodedPolyline) {
                $SafeName = ($RouteName -replace '[\\/:*?"<>|]', '_').Trim()
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
                    -TekstAdresA $GeoStart.FormattedAddress -TekstAdresB $GeoEnd.FormattedAddress `
                    -TekstOdleglosc "$($Trasa.OdlegloscKm) km" -TekstCzas "$($Trasa.CzasMin) min" `
                    -TekstNaglowekLewy $RouteName -TekstNaglowekPrawy "Typ: $RowRouteType"
            }

            $ResultsList.Add([PSCustomObject]@{
                Id               = [string]$CurrentIdx
                Nazwa            = $RouteName
                Start            = $r.Start
                StartGeokodowany = $GeoStart.FormattedAddress
                Koniec           = $r.End
                KoniecGeokodowany= $GeoEnd.FormattedAddress
                LiczbaPrzystankow= $GeocodedWaypoints.Count
                TypTrasy         = $RowRouteType
                OdlegloscKm      = $Trasa.OdlegloscKm
                CzasMin          = $Trasa.CzasMin
                MapaPath         = $MapPath
                GoogleMapsUrl    = $GoogleMapsUrl
                Status           = 'OK'
            })
        }
        catch {
            Write-Warning "  Wyjątek dla trasy $($RouteName): $($_.Exception.Message)"
            $ResultsList.Add([PSCustomObject]@{
                Id               = [string]$CurrentIdx
                Nazwa            = $RouteName
                Start            = $r.Start
                StartGeokodowany = $null
                Koniec           = $r.End
                KoniecGeokodowany= $null
                LiczbaPrzystankow= 0
                TypTrasy         = $RowRouteType
                OdlegloscKm      = $null
                CzasMin          = $null
                MapaPath         = $null
                GoogleMapsUrl    = $null
                Status           = "Błąd: $($_.Exception.Message)"
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
