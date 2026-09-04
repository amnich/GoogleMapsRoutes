#Requires -Version 5.1
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force -ErrorAction SilentlyContinue
<#
.SYNOPSIS
    Google Maps Route & Map Generator — Zaawansowana aplikacja WPF Dark Mode.
    Obsługuje ręczne wprowadzanie tras (Start, Cel, Punkty pośrednie, Fastest/Shortest/Eco)
    oraz wsadowe przetwarzanie plików danych (JSON, CSV, Excel).

.DESCRIPTION
    Funkcje:
      - Tryb ręczny (Manual Input):
          * Punkt startowy i punkt końcowy (z walidacją i geokodowaniem)
          * Dynamiczna lista punktów pośrednich (dodawanie, usuwanie, zmiana kolejności)
          * Wybór optymalizacji: Najszybsza (Fastest), Najkrótsza (Shortest), Ekologiczna (Eco)
          * Wybór typu napędu dla trasy Eco (Benzyna, Diesel, Hybryda, Elektryczny)
          * Natychmiastowe obliczanie trasy, odległości (km) i czasu (min/godz)
          * Interaktywny podgląd mapy statycznej PNG z trasą i ponumerowanymi znacznikami
          * Kopiowanie i bezpośrednie otwieranie linku do nawigacji Google Maps w przeglądarce
      - Tryb wsadowy (Data Source / Batch):
          * Obsługa formatów Excel (.xlsx, .xls), CSV (.csv, .tsv), JSON (.json)
          * Automatyczne wykrywanie schematu pliku i mapowanie kolumn z możliwością korekty
          * Tabela podglądu danych wejściowych (DataGrid)
          * Pasek postępu, procenty, czas, asynchroniczny log zdarzeń w czasie rzeczywistym
          * Tabela wyników ze statusem i bezpośrednim dostępem do map
          * Eksport raportów zbiorczych do Excel, CSV i JSON
      - Bezpieczeństwo i ustawienia:
          * Szyfrowane przechowywanie klucza Google Maps API (Windows DPAPI per-user)
          * Asynchroniczny tester poprawności klucza API (nie zawiesza interfejsu)
          * Konfiguracja domyślnych wymiarów mapy, katalogów i typu trasy
      - Zgodność ze standardem PS2EXE (samodzielny plik .EXE bez zewnętrznych zależności).

.NOTES
    Encoding: UTF-8 with BOM
#>

# ── 1. Wymuszenie protokołów TLS 1.2 / TLS 1.1 dla zapytań HTTPS ───────────────
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls

# ── 2. Wymuszenie trybu STA dla WPF ──────────────────────────────────────────
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) {
    $currentProcess = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    if ($currentProcess -match 'powershell\.exe|pwsh\.exe') {
        Start-Process -FilePath $currentProcess -ArgumentList "-NoProfile -STA -ExecutionPolicy Bypass -File `"$PSCommandPath`""
        exit
    }
}

# ── 3. Ładowanie bibliotek GUI, Drawing i Security ───────────────────────────
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms, System.Drawing, System.Security

# DWM Dark Mode dla paska tytułu okna Windows 10/11
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class DwmDarkWindow {
    [DllImport("dwmapi.dll", PreserveSig = true)]
    public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
}
"@ -ErrorAction SilentlyContinue

# ══════════════════════════════════════════════════════════════════════════════
# 4. SAMODZIELNE FUNKCJE BAZOWE (EMBEDDED DLA ZGODNOŚCI Z PS2EXE)
# ══════════════════════════════════════════════════════════════════════════════

function Protect-SecretString {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$PlainText)
    if ([string]::IsNullOrEmpty($PlainText)) { return $null }
    try {
        Add-Type -AssemblyName System.Security
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($PlainText)
        $protected = [System.Security.Cryptography.ProtectedData]::Protect(
            $bytes, $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
        return [Convert]::ToBase64String($protected)
    }
    catch {
        try {
            $sec = ConvertTo-SecureString -String $PlainText -AsPlainText -Force
            return (ConvertFrom-SecureString -SecureString $sec)
        } catch {
            return $null
        }
    }
}

function Unprotect-SecretString {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$EncryptedText)
    if ([string]::IsNullOrWhiteSpace($EncryptedText)) { return $null }
    try {
        Add-Type -AssemblyName System.Security
        $bytes = [Convert]::FromBase64String($EncryptedText)
        $unprotected = [System.Security.Cryptography.ProtectedData]::Unprotect(
            $bytes, $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
        return [System.Text.Encoding]::UTF8.GetString($unprotected)
    }
    catch {
        try {
            $sec = ConvertTo-SecureString -String $EncryptedText
            $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)
            $str = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
            return $str
        }
        catch {
            return $null
        }
    }
}

function Test-GoogleApiKey {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ApiKey,
        [Parameter()][string]$LanguageCode = 'en'
    )
    if ([string]::IsNullOrWhiteSpace($ApiKey)) {
        return [PSCustomObject]@{ Valid = $false; Message = 'API key is empty.' }
    }
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls
    try {
        $lang = if ($LanguageCode) { ($LanguageCode -split '[-_]')[0].ToLower() } else { 'en' }
        $Url = "https://maps.googleapis.com/maps/api/geocode/json?address=Warszawa&language=$lang&key=$ApiKey"
        $Resp = Invoke-RestMethod -Uri $Url -Method GET -TimeoutSec 15
        if ($Resp.status -eq 'OK' -or $Resp.status -eq 'ZERO_RESULTS') {
            return [PSCustomObject]@{ Valid = $true; Message = 'Google Maps API key is valid and active.' }
        }
        elseif ($Resp.status -eq 'REQUEST_DENIED') {
            $msg = if ($Resp.error_message) { $Resp.error_message } else { 'Request denied by Google API.' }
            return [PSCustomObject]@{ Valid = $false; Message = "Unauthorized: $msg" }
        }
        else {
            return [PSCustomObject]@{ Valid = $false; Message = "Status API: $($Resp.status)" }
        }
    }
    catch {
        return [PSCustomObject]@{ Valid = $false; Message = "Connection error: $($_.Exception.Message)" }
    }
}

function Select-InputDataFile {
    param([string]$InitialDirectory)
    Add-Type -AssemblyName System.Windows.Forms
    $Dialog = [System.Windows.Forms.OpenFileDialog]::new()
    $Dialog.Title = 'Select route data file (JSON, CSV, Excel)'
    $Dialog.Filter = 'All Supported Files (*.xlsx;*.xls;*.csv;*.tsv;*.json)|*.xlsx;*.xls;*.csv;*.tsv;*.json|Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|CSV/TSV Files (*.csv;*.tsv)|*.csv;*.tsv|JSON Files (*.json)|*.json|All Files (*.*)|*.*'

    $chosenDir = $null
    if ($InitialDirectory -and (Test-Path $InitialDirectory)) {
        $chosenDir = $InitialDirectory
    }
    elseif ($script:LastDataDirectory -and (Test-Path $script:LastDataDirectory)) {
        $chosenDir = $script:LastDataDirectory
    }
    elseif ($script:Config -and $script:Config.LastInputFolder -and (Test-Path $script:Config.LastInputFolder)) {
        $chosenDir = $script:Config.LastInputFolder
    }
    elseif ($script:Config -and $script:Config.LastInputPath -and (Test-Path (Split-Path $script:Config.LastInputPath -Parent))) {
        $chosenDir = Split-Path $script:Config.LastInputPath -Parent
    }
    else {
        $samplesDir = if ($PSScriptRoot) { Join-Path $PSScriptRoot 'Samples' } else { Join-Path (Get-Location) 'Samples' }
        if (Test-Path $samplesDir) {
            $chosenDir = $samplesDir
        } else {
            $chosenDir = [Environment]::GetFolderPath('MyDocuments')
        }
    }

    $Dialog.InitialDirectory = $chosenDir
    $Dialog.RestoreDirectory = $true
    $Result = $Dialog.ShowDialog()
    if ($Result -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:LastDataDirectory = Split-Path $Dialog.FileName -Parent
        if ($script:Config) {
            $script:Config.LastInputFolder = $script:LastDataDirectory
            $script:Config.LastInputPath = $Dialog.FileName
        }
        return $Dialog.FileName
    }
    return $null
}

function Get-AddressComponentValue {
    param([object[]]$Components, [string[]]$Types)
    $Matches = @($Components) | Where-Object {
        $_ -and $_.PSObject.Properties.Name -contains 'types' -and
        (@($_.types) | Where-Object { $_ -in $Types } | Select-Object -First 1)
    } | Select-Object -First 1
    if (-not $Matches) { return $null }
    foreach ($FieldName in @('long_name', 'short_name', 'name')) {
        if ($Matches.PSObject.Properties.Name -contains $FieldName) {
            $Value = [string]$Matches.$FieldName
            if (-not [string]::IsNullOrWhiteSpace($Value)) { return $Value }
        }
    }
    return $null
}

function Get-AddressCoordinates {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Address,
        [Parameter(Mandatory)][string]$ApiKey,
        [Parameter()][string]$LanguageCode = 'en',
        [Parameter()][switch]$RequireStreetNumber
    )
    if ([string]::IsNullOrWhiteSpace($Address)) { return $null }

    if ($Address.Trim() -match '^\s*([+-]?\d+(?:\.\d+)?)\s*[,;\s]\s*([+-]?\d+(?:\.\d+)?)\s*$') {
        $lat = [double]$Matches[1]
        $lng = [double]$Matches[2]
        return [PSCustomObject]@{
            Latitude             = $lat
            Longitude            = $lng
            FormattedAddress     = "$lat, $lng"
            UlicaINumer          = $null
            KodPocztowy          = $null
            Miasto               = $null
            MatchType            = 'COORDINATES'
            PartialMatch         = $false
            Status               = 'OK'
            ErrorMessage         = $null
        }
    }

    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls
    $EncodedAddress = [System.Uri]::EscapeDataString($Address.Trim())
    $lang = if ($LanguageCode) { ($LanguageCode -split '[-_]')[0].ToLower() } else { 'en' }
    $Url = "https://maps.googleapis.com/maps/api/geocode/json?address=$EncodedAddress&language=$lang&key=$ApiKey"
    try {
        $Response = Invoke-RestMethod -Uri $Url -Method GET -TimeoutSec 30
        $Results = @($Response.results)
        if ($Response.status -eq 'OK' -and $Results.Count -gt 0) {
            $ResultItem = $Results[0]
            $Location   = $ResultItem.geometry.location

            $Components   = @($ResultItem.address_components)
            $StreetNumber = Get-AddressComponentValue -Components $Components -Types @('street_number')
            $Route        = Get-AddressComponentValue -Components $Components -Types @('route')
            $PostalCode   = Get-AddressComponentValue -Components $Components -Types @('postal_code')
            $City         = Get-AddressComponentValue -Components $Components -Types @('locality', 'postal_town')
            if ([string]::IsNullOrWhiteSpace($City)) {
                $City = Get-AddressComponentValue -Components $Components -Types @('administrative_area_level_3', 'administrative_area_level_2')
            }

            $StreetWithNumber = if ($Route -and $StreetNumber) { "$Route $StreetNumber" }
                                elseif ($Route) { $Route }
                                elseif ($StreetNumber) { $StreetNumber }
                                else { $null }

            $FormattedAddress = if ($lang -eq 'pl') {
                $ResultItem.formatted_address -replace ',\s*Poland$', ', Polska' -replace '\bPoland\b', 'Polska'
            } else {
                $ResultItem.formatted_address
            }
            $LocationType     = if ($ResultItem.geometry -and $ResultItem.geometry.location_type) { [string]$ResultItem.geometry.location_type } else { 'APPROXIMATE' }
            $PartialMatch     = if ($ResultItem.PSObject.Properties.Name -contains 'partial_match') { [bool]$ResultItem.partial_match } else { $false }

            return [PSCustomObject]@{
                Latitude             = [double]$Location.lat
                Longitude            = [double]$Location.lng
                FormattedAddress     = $FormattedAddress
                UlicaINumer          = $StreetWithNumber
                KodPocztowy          = $PostalCode
                Miasto               = $City
                MatchType            = $LocationType
                PartialMatch         = $PartialMatch
                Status               = 'OK'
                ErrorMessage         = $null
            }
        }
        else {
            return [PSCustomObject]@{
                Latitude             = $null; Longitude = $null; FormattedAddress = $null
                UlicaINumer          = $null; KodPocztowy = $null; Miasto = $null
                MatchType            = $null; PartialMatch = $null
                Status               = $Response.status
                ErrorMessage         = $Response.error_message
            }
        }
    }
    catch {
        $Message = $_.Exception.Message
        return [PSCustomObject]@{
            Latitude             = $null; Longitude = $null; FormattedAddress = $null
            UlicaINumer          = $null; KodPocztowy = $null; Miasto = $null
            MatchType            = $null; PartialMatch = $null
            Status               = "EXCEPTION: $Message"
            ErrorMessage         = $Message
        }
    }
}

function Get-GeocodeStatusDescription {
    [CmdletBinding()]
    param(
        [Parameter()][object]$Geo
    )
    if (-not $Geo) { return 'NOT_PROCESSED' }
    if ($Geo.Status -eq 'OK') {
        if ($Geo.PartialMatch -and $Geo.MatchType -in 'APPROXIMATE', 'GEOMETRIC_CENTER') {
            return "OK (Fallback: Approximate / Partial Match - $($Geo.MatchType))"
        }
        elseif ($Geo.PartialMatch) {
            return "OK (Fallback: Partial Match - $($Geo.MatchType))"
        }
        elseif ($Geo.MatchType -eq 'APPROXIMATE') {
            return 'OK (Fallback: Approximate)'
        }
        elseif ($Geo.MatchType -eq 'GEOMETRIC_CENTER') {
            return 'OK (Fallback: Geometric Center)'
        }
        elseif ($Geo.MatchType -eq 'RANGE_INTERPOLATED') {
            return 'OK (Interpolated)'
        }
        elseif ($Geo.MatchType -eq 'ROOFTOP') {
            return 'OK (Exact - ROOFTOP)'
        }
        elseif ($Geo.MatchType -eq 'COORDINATES') {
            return 'OK (Coordinates)'
        }
        else {
            return "OK ($($Geo.MatchType))"
        }
    }
    elseif ($Geo.Status -eq 'ZERO_RESULTS') {
        return 'ZERO_RESULTS (Address Not Found)'
    }
    else {
        return [string]$Geo.Status
    }
}

function Get-CarRouteData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][double]$OriginLat,
        [Parameter(Mandatory)][double]$OriginLng,
        [Parameter(Mandatory)][double]$DestLat,
        [Parameter(Mandatory)][double]$DestLng,
        [Parameter(Mandatory)][string]$ApiKey,
        [Parameter()][object[]]$IntermediatePoints = @(),
        [Parameter()][ValidateSet('Fastest', 'Shortest', 'Eco')][string]$RouteType = 'Fastest',
        [Parameter()][ValidateSet('GASOLINE', 'DIESEL', 'HYBRID', 'ELECTRIC')][string]$EmissionType = 'GASOLINE',
        [Parameter()][string]$LanguageCode = 'en',
        [Parameter()][string]$Units = 'METRIC',
        [Parameter()][switch]$TrafficAware
    )

    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls
    $RoutesUrl = 'https://routes.googleapis.com/directions/v2:computeRoutes'

    $RequestBody = [ordered]@{
        origin       = @{ location = @{ latLng = @{ latitude = $OriginLat; longitude = $OriginLng } } }
        destination  = @{ location = @{ latLng = @{ latitude = $DestLat; longitude = $DestLng } } }
        travelMode   = 'DRIVE'
        languageCode = if ($LanguageCode) { $LanguageCode } else { 'en' }
        units        = $Units
    }

    $HasIntermediates = $false
    if ($null -ne $IntermediatePoints -and @($IntermediatePoints).Count -gt 0) {
        $IntermediatesList = [System.Collections.Generic.List[hashtable]]::new()
        foreach ($pt in $IntermediatePoints) {
            if ($null -ne $pt -and $pt.Latitude -and $pt.Longitude) {
                $IntermediatesList.Add(@{
                    location = @{
                        latLng = @{
                            latitude  = [double]$pt.Latitude
                            longitude = [double]$pt.Longitude
                        }
                    }
                })
            }
        }
        if ($IntermediatesList.Count -gt 0) {
            $RequestBody['intermediates'] = $IntermediatesList
            $HasIntermediates = $true
        }
    }

    switch ($RouteType) {
        'Fastest' {
            $RequestBody['routingPreference'] = if ($TrafficAware) { 'TRAFFIC_AWARE' } else { 'TRAFFIC_UNAWARE' }
            if (-not $HasIntermediates) { $RequestBody['computeAlternativeRoutes'] = $true }
        }
        'Shortest' {
            $RequestBody['routingPreference'] = 'TRAFFIC_UNAWARE'
            if (-not $HasIntermediates) { $RequestBody['computeAlternativeRoutes'] = $true }
        }
        'Eco' {
            $RequestBody['routingPreference'] = 'TRAFFIC_AWARE_OPTIMAL'
            $RequestBody['requestedReferenceRoutes'] = @('FUEL_EFFICIENT')
            $RequestBody['routeModifiers'] = @{
                vehicleInfo = @{ emissionType = $EmissionType }
            }
        }
    }

    $Headers = @{
        'X-Goog-Api-Key'   = $ApiKey
        'Content-Type'     = 'application/json'
        'X-Goog-FieldMask' = 'routes.duration,routes.distanceMeters,routes.polyline.encodedPolyline,routes.description,routes.routeLabels'
    }

    try {
        $JsonBody = $RequestBody | ConvertTo-Json -Depth 10
        $Response = Invoke-RestMethod -Uri $RoutesUrl -Method POST -Headers $Headers -Body $JsonBody -TimeoutSec 60

        $Routes = @($Response.routes)
        if ($Routes.Count -eq 0) {
            return [PSCustomObject]@{
                OdlegloscKm     = $null
                CzasMin         = $null
                DurationSeconds = $null
                EncodedPolyline = $null
                RouteType       = $RouteType
                RouteLabels     = @()
                Status          = 'NO_ROUTES'
                ErrorMessage    = 'Google Routes API did not return any routes.'
            }
        }

        $SelectedRoute = $null
        if ($RouteType -eq 'Shortest') {
            $SelectedRoute = $Routes | Sort-Object -Property { [int64]($_.distanceMeters) } | Select-Object -First 1
        }
        elseif ($RouteType -eq 'Eco') {
            $EcoRoute = $Routes | Where-Object {
                $_.routeLabels -and (@($_.routeLabels) -contains 'FUEL_EFFICIENT')
            } | Select-Object -First 1

            $SelectedRoute = if ($EcoRoute) { $EcoRoute } else { $Routes[0] }
        }
        else {
            $SelectedRoute = $Routes | Sort-Object -Property {
                if ($_.duration) { [double]($_.duration.TrimEnd('s')) } else { [double]::MaxValue }
            } | Select-Object -First 1
        }

        $DistanceKm = if ($SelectedRoute.distanceMeters) { [math]::Round([double]$SelectedRoute.distanceMeters / 1000.0, 2) } else { $null }
        $DurationSec = if ($SelectedRoute.duration) { [double]($SelectedRoute.duration.TrimEnd('s')) } else { $null }
        $DurationMinutes = if ($null -ne $DurationSec) { [math]::Round($DurationSec / 60.0, 0) } else { $null }
        $Polyline = if ($SelectedRoute.polyline) { $SelectedRoute.polyline.encodedPolyline } else { $null }
        $Labels = if ($SelectedRoute.routeLabels) { @($SelectedRoute.routeLabels) } else { @() }

        return [PSCustomObject]@{
            OdlegloscKm     = $DistanceKm
            CzasMin         = $DurationMinutes
            DurationSeconds = $DurationSec
            EncodedPolyline = $Polyline
            RouteType       = $RouteType
            RouteLabels     = $Labels
            Status          = 'OK'
            ErrorMessage    = $null
        }
    }
    catch {
        $ErrorMsg = $_.Exception.Message
        return [PSCustomObject]@{
            OdlegloscKm     = $null
            CzasMin         = $null
            DurationSeconds = $null
            EncodedPolyline = $null
            RouteType       = $RouteType
            RouteLabels     = @()
            Status          = "EXCEPTION: $ErrorMsg"
            ErrorMessage    = $ErrorMsg
        }
    }
}

function Get-GoogleMapsUrl {
    param(
        [string]$Origin,
        [string]$Destination,
        [object[]]$Waypoints = @(),
        [string]$TravelMode = 'driving'
    )
    $OriginEnc = [System.Uri]::EscapeDataString($Origin.Trim())
    $DestEnc = [System.Uri]::EscapeDataString($Destination.Trim())
    $Url = "https://www.google.com/maps/dir/?api=1&origin=$OriginEnc&destination=$DestEnc&travelmode=$TravelMode"

    if ($null -ne $Waypoints -and @($Waypoints).Count -gt 0) {
        $WpStrings = [System.Collections.Generic.List[string]]::new()
        foreach ($wp in $Waypoints) {
            if ($wp -is [string] -and -not [string]::IsNullOrWhiteSpace($wp)) {
                $WpStrings.Add($wp.Trim())
            }
            elseif ($wp.Latitude -and $wp.Longitude) {
                $WpStrings.Add("$($wp.Latitude),$($wp.Longitude)")
            }
            elseif ($wp.ZapytanieAdresowe) {
                $WpStrings.Add([string]$wp.ZapytanieAdresowe)
            }
            elseif ($wp.AdresGeokodowany) {
                $WpStrings.Add([string]$wp.AdresGeokodowany)
            }
        }
        if ($WpStrings.Count -gt 0) {
            $Url += '&waypoints=' + [System.Uri]::EscapeDataString(($WpStrings -join '|'))
        }
    }
    return $Url
}

function Get-WrappedLines {
    param([System.Drawing.Graphics]$G, [string]$Text, [System.Drawing.Font]$F, [float]$MaxW)
    if ([string]::IsNullOrWhiteSpace($Text)) { return [string[]]@('') }
    if ($G.MeasureString($Text, $F).Width -le $MaxW) { return [string[]]@($Text) }
    $Words = $Text -split '\s+'
    $L1 = ''; $L2 = ''; $On2 = $false
    foreach ($W in $Words) {
        if (-not $On2) {
            $T = if ($L1) { "$L1 $W" } else { $W }
            if ($G.MeasureString($T, $F).Width -le $MaxW) { $L1 = $T }
            else { $On2 = $true; $L2 = $W }
        }
        else {
            $T2 = if ($L2) { "$L2 $W" } else { $W }
            if ($G.MeasureString($T2, $F).Width -le $MaxW) { $L2 = $T2 }
            else {
                if ($L2.Length -gt 3) { $L2 = $L2.Substring(0, $L2.Length - 3) + '...' }
                break
            }
        }
    }
    if ($L2) { return [string[]]@($L1, $L2) } else { return [string[]]@($L1) }
}

function Save-RouteMapPng {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$EncodedPolyline,
        [Parameter(Mandatory)][double]$OriginLat,
        [Parameter(Mandatory)][double]$OriginLng,
        [Parameter(Mandatory)][double]$DestLat,
        [Parameter(Mandatory)][double]$DestLng,
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)][string]$ApiKey,
        [Parameter()][int]$Width = 900,
        [Parameter()][int]$Height = 600,
        [Parameter()][object[]]$RoutePoints = @(),
        [Parameter()][Alias('TekstAdresA')][string]$AddressTextA = '',
        [Parameter()][Alias('TekstAdresB')][string]$AddressTextB = '',
        [Parameter()][Alias('TekstOdleglosc')][string]$DistanceText = '',
        [Parameter()][Alias('TekstCzas')][string]$DurationText = '',
        [Parameter()][Alias('TekstNaglowekLewy')][string]$HeaderLeftText = '',
        [Parameter()][Alias('TekstNaglowekPrawy')][string]$HeaderRightText = '',
        [Parameter()][Alias('TekstUmowa')][string]$ContractText = '',
        [Parameter()][Alias('TekstKierunek')][string]$DirectionText = '',
        [Parameter()][Alias('Opis')][string]$Description = '',
        [Parameter()][Alias('DataWygenerowania')][string]$GeneratedDate = '',
        [Parameter()][string]$LanguageCode = 'en',
        [Parameter()][string]$StartRaw = '',
        [Parameter()][string]$StartGeocoded = '',
        [Parameter()][string]$EndRaw = '',
        [Parameter()][string]$EndGeocoded = '',
        [Parameter()][object[]]$WaypointsList = @(),
        [Parameter()][string]$RouteName = '',
        [Parameter()][string]$RouteType = '',
        [Parameter()][object]$OverlayConfig = $null
    )

    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls
    $EncodedForUrl = [System.Uri]::EscapeDataString($EncodedPolyline)
    $MarkerParams = [System.Collections.Generic.List[string]]::new()

    $MarkerStart = [System.Uri]::EscapeDataString("color:green|label:A|$OriginLat,$OriginLng")
    $MarkerParams.Add("&markers=$MarkerStart")

    if ($null -ne $RoutePoints -and @($RoutePoints).Count -gt 0) {
        $IntermediatesOnly = @()
        if ($RoutePoints.Count -gt 2 -and
            [math]::Abs($RoutePoints[0].Latitude - $OriginLat) -lt 0.0001 -and
            [math]::Abs($RoutePoints[-1].Latitude - $DestLat) -lt 0.0001) {
            $IntermediatesOnly = @($RoutePoints[1..($RoutePoints.Count - 2)])
        }
        else {
            $IntermediatesOnly = @($RoutePoints | Where-Object {
                $null -ne $_.Latitude -and $null -ne $_.Longitude -and
                (-not ([math]::Abs($_.Latitude - $OriginLat) -lt 0.0001 -and [math]::Abs($_.Longitude - $OriginLng) -lt 0.0001)) -and
                (-not ([math]::Abs($_.Latitude - $DestLat) -lt 0.0001 -and [math]::Abs($_.Longitude - $DestLng) -lt 0.0001))
            })
        }

        $idx = 1
        foreach ($pt in $IntermediatesOnly) {
            if ($pt.Latitude -and $pt.Longitude) {
                $lbl = if ($idx -le 9) { [string]$idx }
                       elseif ($idx -le 35) { [string][char](55 + $idx) }
                       else { '' }
                $spec = if ($lbl) { "color:blue|label:$lbl|$($pt.Latitude),$($pt.Longitude)" }
                        else { "size:mid|color:blue|$($pt.Latitude),$($pt.Longitude)" }
                $MarkerParams.Add("&markers=" + [System.Uri]::EscapeDataString($spec))
                $idx++
            }
        }
    }

    $MarkerEnd = [System.Uri]::EscapeDataString("color:red|label:B|$DestLat,$DestLng")
    $MarkerParams.Add("&markers=$MarkerEnd")

    $lang = if ($LanguageCode) { ($LanguageCode -split '[-_]')[0].ToLower() } else { 'en' }
    $StaticMapUrl = ("https://maps.googleapis.com/maps/api/staticmap" +
        "?size=${Width}x${Height}" +
        "&language=$lang" +
        "&path=weight:4|color:0x0066FFff|enc:$EncodedForUrl" +
        ($MarkerParams -join '') +
        "&key=$ApiKey")

    try {
        $TargetDir = Split-Path -Parent $OutputPath
        if (-not [string]::IsNullOrWhiteSpace($TargetDir) -and -not (Test-Path $TargetDir)) {
            New-Item -ItemType Directory -Path $TargetDir -Force | Out-Null
        }

        $wc = [System.Net.WebClient]::new()
        try {
            $wc.DownloadFile($StaticMapUrl, $OutputPath)
        }
        finally {
            $wc.Dispose()
        }

        # Resolve overlay configuration
        if ($OverlayConfig -is [string] -and -not [string]::IsNullOrWhiteSpace($OverlayConfig)) {
            try { $OverlayConfig = $OverlayConfig | ConvertFrom-Json } catch { }
        }
        if (-not $OverlayConfig) {
            $OverlayConfig = [PSCustomObject]@{
                EnableTopOverlay    = $true
                EnableBottomOverlay = $true
                Items               = [PSCustomObject]@{
                    StartGeocoded = [PSCustomObject]@{ Enabled = $true;  Panel = 'Bottom'; Align = 'Left';   Order = 1 }
                    EndGeocoded   = [PSCustomObject]@{ Enabled = $true;  Panel = 'Bottom'; Align = 'Left';   Order = 2 }
                    Distance      = [PSCustomObject]@{ Enabled = $true;  Panel = 'Bottom'; Align = 'Left';   Order = 3 }
                    Duration      = [PSCustomObject]@{ Enabled = $true;  Panel = 'Bottom'; Align = 'Center'; Order = 3 }
                    Timestamp     = [PSCustomObject]@{ Enabled = $true;  Panel = 'Bottom'; Align = 'Right';  Order = 3 }
                    RouteName     = [PSCustomObject]@{ Enabled = $true;  Panel = 'Top';    Align = 'Left';   Order = 1 }
                    RouteType     = [PSCustomObject]@{ Enabled = $true;  Panel = 'Top';    Align = 'Right';  Order = 1 }
                    Waypoints     = [PSCustomObject]@{ Enabled = $false; Panel = 'Bottom'; Align = 'Left';   Order = 2 }
                    StartRaw      = [PSCustomObject]@{ Enabled = $false; Panel = 'None';   Align = 'Left';   Order = 1 }
                    EndRaw        = [PSCustomObject]@{ Enabled = $false; Panel = 'None';   Align = 'Left';   Order = 2 }
                }
            }
        }

        $enableTop = if ($null -ne $OverlayConfig.EnableTopOverlay) { [bool]$OverlayConfig.EnableTopOverlay } else { $true }
        $enableBtm = if ($null -ne $OverlayConfig.EnableBottomOverlay) { [bool]$OverlayConfig.EnableBottomOverlay } else { $true }

        # Resolve data values
        $addrStartGeo = if ($StartGeocoded) { $StartGeocoded } elseif ($AddressTextA) { $AddressTextA } else { '' }
        $addrStartRaw = if ($StartRaw) { $StartRaw } else { '' }
        $addrEndGeo   = if ($EndGeocoded) { $EndGeocoded } elseif ($AddressTextB) { $AddressTextB } else { '' }
        $addrEndRaw   = if ($EndRaw) { $EndRaw } else { '' }

        $nameVal = if ($RouteName) { $RouteName } elseif ($HeaderLeftText) { $HeaderLeftText } elseif ($Description) { $Description.Trim() } elseif ($ContractText) { $ContractText } else { '' }

        $typeVal = if ($RouteType) { $RouteType } elseif ($HeaderRightText) { $HeaderRightText } elseif ($DirectionText) { $DirectionText } else { '' }
        if ($typeVal -match '^(?:Type|Typ|Art):\s*(.+)$' -or $typeVal -match '^(Shortest|Fastest|Eco|Najkr[oó]tsza|Najszybsza|Eko|K[uü]rzeste|Schnellste)$') {
            $rawVal = if ($Matches[1]) { $Matches[1].Trim() } else { $Matches[0].Trim() }
            $normVal = if ($rawVal -match '(?i)short|kr[oó]t|k[uü]rz') { 'Shortest' }
                       elseif ($rawVal -match '(?i)eco|eko|fuel') { 'Eco' }
                       elseif ($rawVal -match '(?i)fast|szyb|schnell') { 'Fastest' }
                       else { $rawVal }
            $tPrefix = switch ($lang) { 'de' { 'Typ: ' } 'pl' { 'Typ: ' } default { 'Type: ' } }
            $tName = switch ($lang) {
                'de' { if ($normVal -eq 'Fastest') { 'Schnellste' } elseif ($normVal -eq 'Shortest') { 'Kürzeste' } elseif ($normVal -eq 'Eco') { 'Eco' } else { $normVal } }
                'pl' { if ($normVal -eq 'Fastest') { 'Najszybsza' } elseif ($normVal -eq 'Shortest') { 'Najkrótsza' } elseif ($normVal -eq 'Eco') { 'Eko' } else { $normVal } }
                default { $normVal }
            }
            $typeVal = "$tPrefix$tName"
        }

        $distPrefix = switch ($lang) { 'de' { 'Gesamt: ' } 'pl' { 'Razem: ' } default { 'Total: ' } }
        $distVal = if ($DistanceText) { $DistanceText } else { '' }

        $durVal = if ($DurationText) {
            if ($DurationText -match '^\(.*\)$') { $DurationText } else { "($DurationText)" }
        } else { '' }

        $dateVal = if ($GeneratedDate) { $GeneratedDate } else { (Get-Date -Format 'yyyy-MM-dd  HH:mm') }

        $wpItems = [System.Collections.Generic.List[PSCustomObject]]::new()
        $rawWpList = if ($WaypointsList -and @($WaypointsList).Count -gt 0) {
            $WaypointsList
        } elseif ($RoutePoints -and @($RoutePoints).Count -gt 2) {
            @($RoutePoints[1..($RoutePoints.Count - 2)])
        } else { @() }

        $wIdx = 1
        foreach ($w in $rawWpList) {
            $wText = if ($w -is [string]) { $w }
                     elseif ($w.FormattedAddress) { $w.FormattedAddress }
                     elseif ($w.Address) { $w.Address }
                     else { '' }
            if (-not [string]::IsNullOrWhiteSpace($wText)) {
                $wpItems.Add([PSCustomObject]@{
                    Index = $wIdx
                    Badge = "${wIdx}: "
                    Text  = $wText
                })
                $wIdx++
            }
        }

        # Build active property items map
        $propDataMap = @{
            'StartGeocoded' = @{ Id='StartGeocoded'; Kind='address'; Badge='A: '; BadgeColor='Green'; Text=$addrStartGeo }
            'StartRaw'      = @{ Id='StartRaw';      Kind='address'; Badge='A: '; BadgeColor='Green'; Text=$addrStartRaw }
            'EndGeocoded'   = @{ Id='EndGeocoded';   Kind='address'; Badge='B: '; BadgeColor='Red';   Text=$addrEndGeo }
            'EndRaw'        = @{ Id='EndRaw';        Kind='address'; Badge='B: '; BadgeColor='Red';   Text=$addrEndRaw }
            'Distance'      = @{ Id='Distance';      Kind='stat';    Prefix=$distPrefix; Value=$distVal }
            'Duration'      = @{ Id='Duration';      Kind='stat';    Value=$durVal }
            'Timestamp'     = @{ Id='Timestamp';     Kind='date';    Text=$dateVal }
            'RouteName'     = @{ Id='RouteName';     Kind='title';   Text=$nameVal }
            'RouteType'     = @{ Id='RouteType';     Kind='type';    Text=$typeVal }
            'Waypoints'     = @{ Id='Waypoints';     Kind='waypoints'; Items=$wpItems }
        }

        $topItems = [System.Collections.Generic.List[PSCustomObject]]::new()
        $btmItems = [System.Collections.Generic.List[PSCustomObject]]::new()

        if ($OverlayConfig.Items) {
            $propNames = if ($OverlayConfig.Items -is [System.Collections.IDictionary]) {
                $OverlayConfig.Items.Keys
            } else {
                $OverlayConfig.Items.PSObject.Properties.Name
            }
            foreach ($pName in $propNames) {
                $iCfg = if ($OverlayConfig.Items -is [System.Collections.IDictionary]) {
                    $OverlayConfig.Items[$pName]
                } else {
                    $OverlayConfig.Items.$pName
                }
                if (-not $iCfg) { continue }
                $pEnabled = if ($null -ne $iCfg.Enabled) { [bool]$iCfg.Enabled } else { $true }
                $pPanel   = if ($iCfg.Panel) { [string]$iCfg.Panel } else { 'None' }
                $pAlign   = if ($iCfg.Align) { [string]$iCfg.Align } else { 'Left' }
                $pOrder   = if ($iCfg.Order) { [int]$iCfg.Order } else { 1 }

                if (-not $pEnabled -or $pPanel -eq 'None') { continue }
                if (-not $propDataMap.ContainsKey($pName)) { continue }

                $pData = $propDataMap[$pName]
                $hasContent = $false
                if ($pData.Kind -eq 'waypoints') {
                    $hasContent = ($pData.Items -and $pData.Items.Count -gt 0)
                } elseif ($pData.Kind -eq 'stat') {
                    $hasContent = (-not [string]::IsNullOrWhiteSpace($pData.Value))
                } else {
                    $hasContent = (-not [string]::IsNullOrWhiteSpace($pData.Text))
                }
                if (-not $hasContent) { continue }

                $itemObj = [PSCustomObject]@{
                    Id         = $pName
                    Kind       = $pData.Kind
                    Badge      = $pData.Badge
                    BadgeColor = $pData.BadgeColor
                    Text       = $pData.Text
                    Prefix     = $pData.Prefix
                    Value      = $pData.Value
                    Items      = $pData.Items
                    Panel      = $pPanel
                    Align      = $pAlign
                    Order      = $pOrder
                }

                if ($pPanel -eq 'Top' -and $enableTop) {
                    $topItems.Add($itemObj)
                } elseif ($pPanel -eq 'Bottom' -and $enableBtm) {
                    $btmItems.Add($itemObj)
                }
            }
        }

        $MaTopOverlay = ($enableTop -and $topItems.Count -gt 0)
        $MaBottomOverlay = ($enableBtm -and $btmItems.Count -gt 0)

        if ($MaTopOverlay -or $MaBottomOverlay) {
            try {
                Add-Type -AssemblyName System.Drawing

                $FileBytes = [System.IO.File]::ReadAllBytes($OutputPath)
                $MemStream = [System.IO.MemoryStream]::new($FileBytes)
                $BitmapSrc = [System.Drawing.Bitmap]::new($MemStream)

                $ActualW = $BitmapSrc.Width
                $ActualH = $BitmapSrc.Height

                # Fonts definition
                $FontTopTitle = [System.Drawing.Font]::new('Segoe UI', 10.0, [System.Drawing.FontStyle]::Bold)
                $FontTopType  = [System.Drawing.Font]::new('Segoe UI', 10.0, [System.Drawing.FontStyle]::Bold)
                $FontBadge    = [System.Drawing.Font]::new('Segoe UI', 9.0,  [System.Drawing.FontStyle]::Bold)
                $FontAddr     = [System.Drawing.Font]::new('Segoe UI', 9.5,  [System.Drawing.FontStyle]::Regular)
                $FontDistLbl  = [System.Drawing.Font]::new('Segoe UI', 9.0,  [System.Drawing.FontStyle]::Bold)
                $FontDist     = [System.Drawing.Font]::new('Segoe UI', 12.0, [System.Drawing.FontStyle]::Bold)
                $FontDate     = [System.Drawing.Font]::new('Segoe UI', 8.5,  [System.Drawing.FontStyle]::Regular)

                $PadX  = 14
                $LineH = 20

                # Pre-measurement Graphics
                $dummyBmp = [System.Drawing.Bitmap]::new(1, 1)
                $measGfx  = [System.Drawing.Graphics]::FromImage($dummyBmp)

                # Helper scriptblock to group items by Order
                $BuildRows = {
                    param($items)
                    $orders = @($items | Select-Object -ExpandProperty Order -Unique | Sort-Object)
                    $rows = [System.Collections.Generic.List[PSCustomObject]]::new()
                    foreach ($ord in $orders) {
                        $rowItems = @($items | Where-Object { $_.Order -eq $ord })
                        $left   = [System.Collections.Generic.List[PSCustomObject]]::new()
                        $center = [System.Collections.Generic.List[PSCustomObject]]::new()
                        $right  = [System.Collections.Generic.List[PSCustomObject]]::new()
                        foreach ($it in $rowItems) {
                            if ($it.Align -eq 'Right') { $right.Add($it) }
                            elseif ($it.Align -eq 'Center') { $center.Add($it) }
                            else { $left.Add($it) }
                        }
                        $rows.Add([PSCustomObject]@{
                            Order  = $ord
                            Left   = $left
                            Center = $center
                            Right  = $right
                            Height = 20
                        })
                    }
                    return $rows.ToArray()
                }

                $topRows = @(if ($MaTopOverlay) { & $BuildRows $topItems } else { @() })
                $btmRows = @(if ($MaBottomOverlay) { & $BuildRows $btmItems } else { @() })

                # Measure row heights
                $MeasureRows = {
                    param($rows, $availWidth)
                    foreach ($row in @($rows)) {
                        $maxH = 20
                        $allItems = @($row.Left) + @($row.Center) + @($row.Right)
                        foreach ($it in $allItems) {
                            if ($it.Kind -eq 'address') {
                                $badgeSz = $measGfx.MeasureString($it.Badge, $FontBadge)
                                $addrW = [float]($availWidth - $badgeSz.Width)
                                $lines = @(Get-WrappedLines -G $measGfx -Text $it.Text -F $FontAddr -MaxW $addrW)
                                $it | Add-Member -NotePropertyName 'WrappedLines' -NotePropertyValue $lines -Force
                                $h = [math]::Max(1, $lines.Count) * $LineH
                                if ($h -gt $maxH) { $maxH = $h }
                            }
                            elseif ($it.Kind -eq 'waypoints') {
                                $totalWpH = 0
                                foreach ($wp in $it.Items) {
                                    $bSz = $measGfx.MeasureString($wp.Badge, $FontBadge)
                                    $wpMaxW = [float]($availWidth - $bSz.Width)
                                    $wpLines = @(Get-WrappedLines -G $measGfx -Text $wp.Text -F $FontAddr -MaxW $wpMaxW)
                                    $wp | Add-Member -NotePropertyName 'WrappedLines' -NotePropertyValue $wpLines -Force
                                    $totalWpH += [math]::Max(1, $wpLines.Count) * $LineH
                                }
                                if ($totalWpH -gt $maxH) { $maxH = $totalWpH }
                            }
                            elseif ($it.Kind -eq 'stat') {
                                if ($maxH -lt 24) { $maxH = 24 }
                            }
                            elseif ($it.Kind -in @('title', 'type')) {
                                if ($maxH -lt 22) { $maxH = 22 }
                            }
                        }
                        $row.Height = $maxH
                    }
                }

                $availContentW = [float]($ActualW - ($PadX * 2))
                & $MeasureRows $topRows $availContentW
                & $MeasureRows $btmRows $availContentW

                $measGfx.Dispose()
                $dummyBmp.Dispose()

                # Calculate banner heights
                $TopPad = 8; $TopBotPad = 8; $TopRowSpacing = 4
                $TopBarH = 0
                if ($MaTopOverlay -and @($topRows).Count -gt 0) {
                    $sumTopH = (@($topRows) | Measure-Object -Property Height -Sum).Sum
                    if (-not $sumTopH) { $sumTopH = 20 }
                    $TopBarH = [int]($TopPad + $sumTopH + ((@($topRows).Count - 1) * $TopRowSpacing) + $TopBotPad)
                    if ($TopBarH -lt 38) { $TopBarH = 38 }
                }

                $BtmPadTop = 10; $BtmPadBot = 10; $BtmRowSpacing = 6
                $BtmBarH = 0
                if ($MaBottomOverlay -and @($btmRows).Count -gt 0) {
                    $sumBtmH = (@($btmRows) | Measure-Object -Property Height -Sum).Sum
                    if (-not $sumBtmH) { $sumBtmH = 20 }
                    $BtmBarH = [int]($BtmPadTop + $sumBtmH + ((@($btmRows).Count - 1) * $BtmRowSpacing) + $BtmPadBot)
                }

                $FinalW = $ActualW
                $FinalH = $ActualH + $TopBarH + $BtmBarH

                $Bitmap = [System.Drawing.Bitmap]::new($FinalW, $FinalH, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
                $Graphics = [System.Drawing.Graphics]::FromImage($Bitmap)
                $Graphics.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
                $Graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit

                # 1. Background fill
                $BrushBg = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 15, 23, 42))
                $Graphics.FillRectangle($BrushBg, 0, 0, $FinalW, $FinalH)

                # 2. Draw map image in the middle
                $Graphics.DrawImage($BitmapSrc, 0, $TopBarH, $ActualW, $ActualH)

                # 3. Brushes & Pens
                $PenSep      = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(255, 51, 65, 85), 1.5)
                $BrushWhite  = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 248, 250, 252))
                $BrushYellow = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 250, 204, 21))
                $BrushCyan   = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 56, 189, 248))
                $BrushGreen  = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 16, 185, 129))
                $BrushRed    = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 239, 68, 68))
                $BrushMuted  = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 148, 163, 184))

                # Helper scriptblock to measure width of an item
                $MeasureItemWidth = {
                    param($it)
                    if ($it.Kind -eq 'address') {
                        $bSz = $Graphics.MeasureString($it.Badge, $FontBadge)
                        $tSz = $Graphics.MeasureString($it.Text, $FontAddr)
                        return ($bSz.Width + $tSz.Width)
                    }
                    elseif ($it.Kind -eq 'stat') {
                        $w = 0
                        if ($it.Prefix) { $w += $Graphics.MeasureString($it.Prefix, $FontDistLbl).Width }
                        if ($it.Value)  { $w += $Graphics.MeasureString($it.Value, $FontDist).Width }
                        return $w
                    }
                    elseif ($it.Kind -eq 'title') {
                        return $Graphics.MeasureString($it.Text, $FontTopTitle).Width
                    }
                    elseif ($it.Kind -eq 'type') {
                        return $Graphics.MeasureString($it.Text, $FontTopType).Width
                    }
                    elseif ($it.Kind -eq 'date') {
                        return $Graphics.MeasureString($it.Text, $FontDate).Width
                    }
                    elseif ($it.Kind -eq 'waypoints') {
                        return 200
                    }
                    return 0
                }

                # Helper scriptblock to draw an item at specified coordinates
                $DrawItem = {
                    param($it, [float]$x, [float]$y)
                    if ($it.Kind -eq 'address') {
                        $badgeBrush = if ($it.BadgeColor -eq 'Red') { $BrushRed } else { $BrushGreen }
                        $Graphics.DrawString($it.Badge, $FontBadge, $badgeBrush, $x, $y)
                        $bSz = $Graphics.MeasureString($it.Badge, $FontBadge)
                        $curLineY = $y
                        $lines = if ($it.WrappedLines) { $it.WrappedLines } else { @($it.Text) }
                        foreach ($line in $lines) {
                            $Graphics.DrawString($line, $FontAddr, $BrushWhite, ($x + $bSz.Width), $curLineY)
                            $curLineY += [float]$LineH
                        }
                    }
                    elseif ($it.Kind -eq 'waypoints') {
                        $wpY = $y
                        foreach ($wp in $it.Items) {
                            $Graphics.DrawString($wp.Badge, $FontBadge, $BrushCyan, $x, $wpY)
                            $bSz = $Graphics.MeasureString($wp.Badge, $FontBadge)
                            $lines = if ($wp.WrappedLines) { $wp.WrappedLines } else { @($wp.Text) }
                            foreach ($line in $lines) {
                                $Graphics.DrawString($line, $FontAddr, $BrushWhite, ($x + $bSz.Width), $wpY)
                                $wpY += [float]$LineH
                            }
                        }
                    }
                    elseif ($it.Kind -eq 'stat') {
                        $statX = $x
                        if ($it.Prefix) {
                            $pSz = $Graphics.MeasureString($it.Prefix, $FontDistLbl)
                            $Graphics.DrawString($it.Prefix, $FontDistLbl, $BrushCyan, $statX, ($y + 2))
                            $statX += $pSz.Width
                        }
                        if ($it.Value) {
                            $Graphics.DrawString($it.Value, $FontDist, $BrushYellow, $statX, $y)
                        }
                    }
                    elseif ($it.Kind -eq 'title') {
                        $Graphics.DrawString($it.Text, $FontTopTitle, $BrushWhite, $x, $y)
                    }
                    elseif ($it.Kind -eq 'type') {
                        $Graphics.DrawString($it.Text, $FontTopType, $BrushYellow, $x, $y)
                    }
                    elseif ($it.Kind -eq 'date') {
                        $Graphics.DrawString($it.Text, $FontDate, $BrushMuted, $x, ($y + 3))
                    }
                }

                # Helper scriptblock to render a banner's rows
                $RenderBannerRows = {
                    param($rows, [float]$startY, [float]$spacing)
                    $curY = $startY
                    foreach ($row in $rows) {
                        $leftX = [float]$PadX

                        # 1. Left items
                        foreach ($it in $row.Left) {
                            & $DrawItem $it $leftX $curY
                            $w = & $MeasureItemWidth $it
                            $leftX += [float]($w + 14)
                        }

                        # 2. Right items
                        $totalRightW = 0
                        foreach ($it in $row.Right) {
                            $totalRightW += [float]((& $MeasureItemWidth $it) + 12)
                        }
                        $rightX = [float]($FinalW - $PadX - $totalRightW + 12)
                        foreach ($it in $row.Right) {
                            & $DrawItem $it $rightX $curY
                            $w = & $MeasureItemWidth $it
                            $rightX += [float]($w + 12)
                        }

                        # 3. Center items
                        $totalCenterW = 0
                        foreach ($it in $row.Center) {
                            $totalCenterW += [float]((& $MeasureItemWidth $it) + 12)
                        }
                        $centerX = [float][math]::Max($leftX + 10, ($FinalW - $totalCenterW + 12) / 2)
                        foreach ($it in $row.Center) {
                            & $DrawItem $it $centerX $curY
                            $w = & $MeasureItemWidth $it
                            $centerX += [float]($w + 12)
                        }

                        $curY += [float]($row.Height + $spacing)
                    }
                }

                # 4. Draw Top Header Banner
                if ($MaTopOverlay -and $TopBarH -gt 0 -and @($topRows).Count -gt 0) {
                    $Graphics.DrawLine($PenSep, 0, $TopBarH, $FinalW, $TopBarH)
                    $topStartY = [float]$TopPad
                    if (@($topRows).Count -eq 1) {
                        $topStartY = [float][math]::Max(6, ($TopBarH - $topRows[0].Height) / 2)
                    }
                    & $RenderBannerRows $topRows $topStartY $TopRowSpacing
                }

                # 5. Draw Bottom Footer Banner
                if ($MaBottomOverlay -and $BtmBarH -gt 0 -and @($btmRows).Count -gt 0) {
                    $BtmBarY = $TopBarH + $ActualH
                    $Graphics.DrawLine($PenSep, 0, $BtmBarY, $FinalW, $BtmBarY)
                    $btmStartY = [float]($BtmBarY + $BtmPadTop)
                    & $RenderBannerRows $btmRows $btmStartY $BtmRowSpacing
                }

                # Dispose GDI+ objects
                $PenSep.Dispose()
                $BrushBg.Dispose(); $BrushWhite.Dispose(); $BrushYellow.Dispose()
                $BrushCyan.Dispose(); $BrushGreen.Dispose(); $BrushRed.Dispose(); $BrushMuted.Dispose()
                $FontTopTitle.Dispose(); $FontTopType.Dispose()
                $FontBadge.Dispose(); $FontAddr.Dispose(); $FontDistLbl.Dispose(); $FontDist.Dispose(); $FontDate.Dispose()
                $Graphics.Dispose()

                $Bitmap.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
                $Bitmap.Dispose()
                $BitmapSrc.Dispose()
                $MemStream.Dispose()
            }
            catch { }
        }
        return $true
    }
    catch {
        return $false
    }
}

function Find-MatchingPropertyName {
    param([string[]]$AvailableProperties, [string[]]$Patterns)
    foreach ($pattern in $Patterns) {
        $found = $AvailableProperties | Where-Object { $null -ne $_ -and $_.Trim() -match $pattern } | Select-Object -First 1
        if ($found) { return $found }
    }
    return $null
}

function Import-RouteDataFile {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path, [Parameter()][string]$Delimiter = '')

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Plik wejściowy nie istnieje: $Path"
    }

    $Extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    $RawRows = $null
    $Format = $null

    switch ($Extension) {
        { $_ -in '.xlsx', '.xls' } {
            $Format = 'Excel'
            if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
                throw "Wymagany moduł ImportExcel nie jest zainstalowany. Zainstaluj go poleceniem: Install-Module -Name ImportExcel -Scope CurrentUser"
            }
            Import-Module -Name ImportExcel -ErrorAction Stop
            $RawRows = @(Import-Excel -Path $Path)
        }
        { $_ -in '.csv', '.tsv', '.txt' } {
            $Format = 'CSV'
            $FirstLine = Get-Content -LiteralPath $Path -TotalCount 1
            $UsedDelimiter = if (-not [string]::IsNullOrWhiteSpace($Delimiter)) { $Delimiter }
            elseif ($Extension -eq '.tsv' -or $FirstLine -match "`t") { "`t" }
            elseif ($FirstLine -match ';') { ';' }
            else { ',' }
            $RawRows = @(Import-Csv -LiteralPath $Path -Delimiter $UsedDelimiter)
        }
        '.json' {
            $Format = 'JSON'
            $Content = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
            $ParsedJson = $Content | ConvertFrom-Json
            if ($ParsedJson -is [System.Collections.IEnumerable] -and -not ($ParsedJson -is [string])) {
                $RawRows = @($ParsedJson)
            }
            elseif ($ParsedJson.PSObject.Properties.Name -contains 'Routes') {
                $RawRows = @($ParsedJson.Routes)
            }
            elseif ($ParsedJson.PSObject.Properties.Name -contains 'Stops') {
                $RawRows = @($ParsedJson.Stops)
            }
            else {
                $RawRows = @($ParsedJson)
            }
        }
        default {
            throw "Nieobsługiwany format pliku: $Extension. Obsługiwane rozszerzenia: .xlsx, .xls, .csv, .tsv, .json"
        }
    }

    if ($null -eq $RawRows -or $RawRows.Count -eq 0) {
        return [PSCustomObject]@{
            Mode       = 'Empty'
            Routes     = @()
            RawData    = @()
            FilePath   = $Path
            Format     = $Format
            TotalCount = 0
        }
    }

    $PropNames = @($RawRows[0].PSObject.Properties.Name)

    # Sprawdzenie czy to sekwencja przystanków (SequentialStops)
    $ColSeq = Find-MatchingPropertyName -AvailableProperties $PropNames -Patterns @('^(lp|l\.p\.|kolejnosc|stop|sequence|order|nr)$')
    $ColAddrSeq = Find-MatchingPropertyName -AvailableProperties $PropNames -Patterns @('^(adres|address|lokalizacja|punkt|miejsce)$', 'lokalizacja.*(odbioru|dowozu)', 'adres.*(odbioru|dowozu)')
    $ColCitySeq = Find-MatchingPropertyName -AvailableProperties $PropNames -Patterns @('^(miejscowosc|miasto|city|town)$')

    $IsSequentialStops = ($ColSeq -and ($ColAddrSeq -or $ColCitySeq) -and -not (Find-MatchingPropertyName -AvailableProperties $PropNames -Patterns @('^(start|origin|adres.*a)$')))

    if ($IsSequentialStops) {
        $OrderedStops = @($RawRows | Sort-Object { [int]($_.$ColSeq) })
        $StopList = [System.Collections.Generic.List[PSCustomObject]]::new()
        foreach ($st in $OrderedStops) {
            $addr = if ($ColAddrSeq) { [string]$st.$ColAddrSeq } else { '' }
            $city = if ($ColCitySeq) { [string]$st.$ColCitySeq } else { '' }
            $fullAddr = if ($addr -and $city) { "$addr, $city" } elseif ($addr) { $addr } else { $city }
            $StopList.Add([PSCustomObject]@{
                Sequence = [string]$st.$ColSeq
                Address  = $fullAddr.Trim()
                Raw      = $st
            })
        }

        $RouteObj = $null
        if ($StopList.Count -ge 2) {
            $StartPoint = $StopList[0].Address
            $EndPoint = $StopList[$StopList.Count - 1].Address
            $Waypoints = if ($StopList.Count -gt 2) { @($StopList[1..($StopList.Count - 2)] | ForEach-Object { $_.Address }) } else { @() }
            $RouteObj = [PSCustomObject]@{
                Id          = '1'
                Name        = "Multi-point Route ($($StopList.Count) stops)"
                Start       = $StartPoint
                End         = $EndPoint
                Waypoints   = $Waypoints
                RouteType   = 'Fastest'
                OriginalRow = $OrderedStops
            }
        }

        $RoutesList = [System.Collections.Generic.List[PSCustomObject]]::new()
        if ($RouteObj) { $RoutesList.Add($RouteObj) }

        return [PSCustomObject]@{
            Mode       = 'SequentialStops'
            Stops      = $StopList
            Routes     = $RoutesList
            RawData    = $RawRows
            FilePath   = $Path
            Format     = $Format
            TotalCount = $StopList.Count
        }
    }

    # Tryb RouteList (wiersz = trasa)
    $ColStart = Find-MatchingPropertyName -AvailableProperties $PropNames -Patterns @(
        '(?i)^(start|origin|startpoint|poczat.*|poczatek|od|from|dom)$',
        '(?i)adres.*a|^a$',
        '(?i)punkt.*(poczat|start)'
    )
    $ColEnd = Find-MatchingPropertyName -AvailableProperties $PropNames -Patterns @(
        '(?i)^(end|dest|destination|endpoint|koniec.*|konic.*|cel|meta|do|to|szkola)$',
        '(?i)adres.*b|^b$',
        '(?i)punkt.*(konic|koniec|docel|cel)'
    )
    $ColWaypoints = Find-MatchingPropertyName -AvailableProperties $PropNames -Patterns @(
        '(?i)^(waypoints|waypoint|posredn.*|punkty.*posredn.*|przystank.*|via|stops|praca)$',
        '(?i)posrednie'
    )
    $ColName = Find-MatchingPropertyName -AvailableProperties $PropNames -Patterns @(
        '(?i)^(name|nazwa|umowa|contract|id|nr|opis|description|tytul)$',
        '(?i)numer.*umowy'
    )
    $ColRouteType = Find-MatchingPropertyName -AvailableProperties $PropNames -Patterns @(
        '(?i)^(routetype|typ|typtrasy|tryb|mode|optimization)$'
    )

    $NormalizedRoutes = [System.Collections.Generic.List[PSCustomObject]]::new()
    $idx = 1

    foreach ($row in $RawRows) {
        $startVal = if ($ColStart) { [string]$row.$ColStart } else { '' }
        $endVal   = if ($ColEnd) { [string]$row.$ColEnd } else { '' }
        if ([string]::IsNullOrWhiteSpace($startVal) -or [string]::IsNullOrWhiteSpace($endVal)) { continue }

        $nameVal = if ($ColName) { [string]$row.$ColName } else { "Route $idx" }
        $typeVal = if ($ColRouteType) { [string]$row.$ColRouteType } else { $null }

        if ($typeVal -match '(?i)eco|fuel|paliw|eko') { $typeVal = 'Eco' }
        elseif ($typeVal -match '(?i)short|krot|krót') { $typeVal = 'Shortest' }
        elseif ($typeVal -match '(?i)fast|szyb') { $typeVal = 'Fastest' }
        else { $typeVal = $null }

        $waypointsList = [System.Collections.Generic.List[string]]::new()
        if ($ColWaypoints -and -not [string]::IsNullOrWhiteSpace($row.$ColWaypoints)) {
            $rawWp = $row.$ColWaypoints
            if ($rawWp -is [System.Collections.IEnumerable] -and -not ($rawWp -is [string])) {
                foreach ($item in $rawWp) {
                    if (-not [string]::IsNullOrWhiteSpace([string]$item)) { $waypointsList.Add(([string]$item).Trim()) }
                }
            }
            else {
                $splits = ([string]$rawWp) -split '(?<!\\)[|;]'
                foreach ($s in $splits) {
                    $cleaned = $s.Trim()
                    if (-not [string]::IsNullOrWhiteSpace($cleaned)) { $waypointsList.Add($cleaned) }
                }
            }
        }

        $NormalizedRoutes.Add([PSCustomObject]@{
            Id          = [string]$idx
            Name        = $nameVal
            Start       = $startVal.Trim()
            End         = $endVal.Trim()
            Waypoints   = @($waypointsList)
            RouteType   = $typeVal
            OriginalRow = $row
        })
        $idx++
    }

    return [PSCustomObject]@{
        Mode       = 'RouteList'
        Routes     = $NormalizedRoutes
        RawData    = $RawRows
        FilePath   = $Path
        Format     = $Format
        TotalCount = $NormalizedRoutes.Count
        Columns    = [PSCustomObject]@{
            Start     = $ColStart
            End       = $ColEnd
            Waypoints = $ColWaypoints
            Name      = $ColName
            RouteType = $ColRouteType
        }
    }
}

function Export-RouteResults {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object[]]$Results,
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter()][ValidateSet('Excel', 'CSV', 'JSON')][string]$Format = 'Excel'
    )

    $TargetDir = Split-Path -Parent $OutputPath
    if (-not [string]::IsNullOrWhiteSpace($TargetDir) -and -not (Test-Path $TargetDir)) {
        New-Item -ItemType Directory -Path $TargetDir -Force | Out-Null
    }

    # Extract flat summary rows (excluding nested Points array from main sheet/file)
    $RoutesFlat = [System.Collections.Generic.List[PSCustomObject]]::new()
    $PointsFlat = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($r in $Results) {
        $routeId   = if ($null -ne $r.Id) { [string]$r.Id } else { '' }
        $routeName = if ($r.Name) { [string]$r.Name } elseif ($r.Nazwa) { [string]$r.Nazwa } else { "Route $routeId" }
        $startOrig = if ($r.Start) { [string]$r.Start } else { '' }
        $startGeo  = if ($r.StartGeocoded) { [string]$r.StartGeocoded } elseif ($r.StartGeokodowany) { [string]$r.StartGeokodowany } else { '' }
        $startStat = if ($r.StartStatus) { [string]$r.StartStatus } else { '' }
        $endOrig   = if ($r.End) { [string]$r.End } elseif ($r.Koniec) { [string]$r.Koniec } else { '' }
        $endGeo    = if ($r.EndGeocoded) { [string]$r.EndGeocoded } elseif ($r.KoniecGeokodowany) { [string]$r.KoniecGeokodowany } else { '' }
        $endStat   = if ($r.EndStatus) { [string]$r.EndStatus } else { '' }
        $wpCount   = if ($null -ne $r.WaypointsCount) { [int]$r.WaypointsCount } elseif ($null -ne $r.LiczbaPrzystankow) { [int]$r.LiczbaPrzystankow } else { 0 }
        $rType     = if ($r.RouteType) { [string]$r.RouteType } elseif ($r.TypTrasy) { [string]$r.TypTrasy } else { '' }
        $dist      = if ($null -ne $r.DistanceKm) { $r.DistanceKm } elseif ($null -ne $r.OdlegloscKm) { $r.OdlegloscKm } else { $null }
        $dur       = if ($null -ne $r.DurationMin) { $r.DurationMin } elseif ($null -ne $r.CzasMin) { $r.CzasMin } else { $null }
        $status    = if ($r.Status) { [string]$r.Status } else { '' }
        $map       = if ($r.MapPath) { [string]$r.MapPath } elseif ($r.MapaPath) { [string]$r.MapaPath } else { '' }
        $url       = if ($r.GoogleMapsUrl) { [string]$r.GoogleMapsUrl } else { '' }

        # Build waypoints summary text
        $wpSummaryList = [System.Collections.Generic.List[string]]::new()
        if ($r.Points -and ($r.Points -is [System.Collections.IEnumerable])) {
            foreach ($pt in $r.Points) {
                if ($pt.PointType -like 'Waypoint*') {
                    $ptSummary = "$($pt.PointType): '$($pt.OriginalAddress)'"
                    if ($pt.GeocodedAddress) { $ptSummary += " -> '$($pt.GeocodedAddress)'" }
                    if ($pt.GeocodeStatus) { $ptSummary += " [$($pt.GeocodeStatus)]" }
                    $wpSummaryList.Add($ptSummary)
                }

                $PointsFlat.Add([PSCustomObject]@{
                    RouteId         = $routeId
                    RouteName       = $routeName
                    PointOrder      = $pt.Order
                    PointType       = $pt.PointType
                    OriginalAddress = $pt.OriginalAddress
                    GeocodedAddress = $pt.GeocodedAddress
                    GeocodeStatus   = $pt.GeocodeStatus
                    MatchType       = $pt.MatchType
                    IsFallback      = if ($null -ne $pt.IsFallback) { [bool]$pt.IsFallback } else { $false }
                    Latitude        = $pt.Latitude
                    Longitude       = $pt.Longitude
                })
            }
        }

        $wpSummaryText = $wpSummaryList -join ' | '

        $RoutesFlat.Add([PSCustomObject]@{
            Id               = $routeId
            Name             = $routeName
            Start_Original   = $startOrig
            Start_Geocoded   = $startGeo
            Start_Status     = $startStat
            End_Original     = $endOrig
            End_Geocoded     = $endGeo
            End_Status       = $endStat
            WaypointsCount   = $wpCount
            RouteType        = $rType
            DistanceKm       = $dist
            DurationMin      = $dur
            Status           = $status
            WaypointsSummary = $wpSummaryText
            MapPath          = $map
            GoogleMapsUrl    = $url
        })
    }

    $csvEncoding = if ($PSVersionTable.PSVersion.Major -ge 7) { 'utf8BOM' } else { 'UTF8' }

    switch ($Format) {
        'Excel' {
            if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
                Write-Warning "Moduł ImportExcel nie jest zainstalowany. Eksportowanie do CSV zamiast Excel."
                $CsvPath = [System.IO.Path]::ChangeExtension($OutputPath, '.csv')
                $RoutesFlat | Export-Csv -LiteralPath $CsvPath -NoTypeInformation -Encoding $csvEncoding -Delimiter ';'
                if ($PointsFlat.Count -gt 0) {
                    $PtsCsv = [System.IO.Path]::Combine($TargetDir, "$([System.IO.Path]::GetFileNameWithoutExtension($CsvPath))_punkty.csv")
                    $PointsFlat | Export-Csv -LiteralPath $PtsCsv -NoTypeInformation -Encoding $csvEncoding -Delimiter ';'
                }
                return $CsvPath
            }
            Import-Module -Name ImportExcel -ErrorAction Stop
            if (Test-Path -LiteralPath $OutputPath) {
                Remove-Item -LiteralPath $OutputPath -Force -ErrorAction SilentlyContinue
            }
            $RoutesFlat | Export-Excel -Path $OutputPath -WorksheetName 'Trasy' -TableName 'WynikiTras' -AutoSize -AutoFilter -FreezeTopRow
            if ($PointsFlat.Count -gt 0) {
                $PointsFlat | Export-Excel -Path $OutputPath -WorksheetName 'PunktyTrasy' -TableName 'PunktyTrasy' -AutoSize -AutoFilter -FreezeTopRow
            }
            return $OutputPath
        }
        'CSV' {
            $RoutesFlat | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding $csvEncoding -Delimiter ';'
            if ($PointsFlat.Count -gt 0) {
                $PtsCsv = [System.IO.Path]::Combine($TargetDir, "$([System.IO.Path]::GetFileNameWithoutExtension($OutputPath))_punkty.csv")
                $PointsFlat | Export-Csv -LiteralPath $PtsCsv -NoTypeInformation -Encoding $csvEncoding -Delimiter ';'
            }
            return $OutputPath
        }
        'JSON' {
            $JsonContent = $Results | ConvertTo-Json -Depth 5
            [System.IO.File]::WriteAllText($OutputPath, $JsonContent, [System.Text.Encoding]::UTF8)
            return $OutputPath
        }
    }
}

# ══════════════════════════════════════════════════════════════════════════════
# 5. KONFIGURACJA I DPAPI SECURITY
# ══════════════════════════════════════════════════════════════════════════════

$script:AppDirName = 'GoogleMapsRoutes'
$script:LocalConfigFolder = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) $script:AppDirName
if (-not (Test-Path $script:LocalConfigFolder)) {
    New-Item -ItemType Directory -Path $script:LocalConfigFolder -Force | Out-Null
}
$script:ConfigFile = Join-Path $script:LocalConfigFolder 'config.json'
$script:LogFile    = Join-Path $script:LocalConfigFolder 'GoogleMapsRoutes.log'

# External Localization File resolution:
# Priority 1: $PSScriptRoot\localization.json
# Priority 2: EXE directory\localization.json (for compiled PS2EXE binaries)
# Priority 3: %LOCALAPPDATA%\GoogleMapsRoutes\localization.json
$script:ExeDir = try {
    Split-Path ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) -Parent
} catch { $null }

$script:LocalizationFile = if ($PSScriptRoot -and (Test-Path (Join-Path $PSScriptRoot 'localization.json'))) {
    Join-Path $PSScriptRoot 'localization.json'
} elseif ($script:ExeDir -and (Test-Path (Join-Path $script:ExeDir 'localization.json'))) {
    Join-Path $script:ExeDir 'localization.json'
} elseif (Test-Path (Join-Path $script:LocalConfigFolder 'localization.json')) {
    Join-Path $script:LocalConfigFolder 'localization.json'
} elseif ($PSScriptRoot) {
    Join-Path $PSScriptRoot 'localization.json'
} elseif ($script:ExeDir) {
    Join-Path $script:ExeDir 'localization.json'
} else {
    Join-Path $script:LocalConfigFolder 'localization.json'
}

# ==============================================================================
# BUNDLER: EMBEDDED LOCALIZATION CATALOG (SELF-SUFFICIENT PS2EXE EXECUTION)
# ==============================================================================
$script:EmbeddedLocalizationJson = @'
{
  "DefaultLanguage": "en",
  "Languages": {
    "en": {
      "DisplayName": "English",
      "GoogleCode": "en",
      "Strings": {
        "AppTitle": "Google Maps Route & Map Generator",
        "AppSubtitle": "Multi-point driving routes: Fastest, Shortest, Eco-friendly | Import JSON, CSV, Excel",
        "ApiBadgeChecking": "API: Checking...",
        "ApiBadgeActive": "API: Active",
        "ApiBadgeMissing": "API: Missing Key",
        "ApiBadgeError": "API: Error",
        "BtnQuickSettings": "⚙ API Settings",
        "FooterReady": "Ready.",
        "FooterVersion": "Google Maps Routes v2.0",
        "TabManual": "📍 Manual Route",
        "TabBatch": "📁 Batch File Processing",
        "TabSettings": "⚙ Settings & API Key",
        "ManualHeaderRoutePoints": "Route Points",
        "ManualOrigin": "Origin (Start / A):",
        "ManualWaypoints": "Intermediate Stops (optional up to 25):",
        "ManualWaypointsTooltip": "Enter waypoint address and click Add",
        "ManualBtnAdd": "➕ Add",
        "ManualBtnUp": "▲ Up",
        "ManualBtnDown": "▼ Down",
        "ManualBtnRemove": "✕ Remove",
        "ManualBtnClear": "🗑 Clear",
        "ManualDestination": "Destination (End / B):",
        "ManualRouteName": "Route Name / Description:",
        "ManualHeaderOptimization": "Route Optimization",
        "ManualOptFastest": "⚡ Fastest",
        "ManualOptShortest": "📏 Shortest",
        "ManualOptEco": "🌿 Eco",
        "ManualEmission": "Vehicle Engine Type (for Eco route):",
        "ManualFuelGasoline": "Gasoline",
        "ManualFuelDiesel": "Diesel",
        "ManualFuelHybrid": "Hybrid",
        "ManualFuelElectric": "Electric",
        "ManualTrafficAware": "Real-time traffic awareness (Live Traffic)",
        "ManualBtnCalculate": "🚀 CALCULATE ROUTE & DOWNLOAD MAP",
        "ManualBtnCalculating": "⏳ CALCULATING ROUTE...",
        "ManualStatDistance": "DISTANCE",
        "ManualStatDuration": "DURATION",
        "ManualStatType": "ROUTE TYPE",
        "ManualStatusIdle": "Idle",
        "ManualStatusCalculating": "Calculating...",
        "ManualStatusSuccess": "Route calculated",
        "ManualStatusError": "Calculation error",
        "ManualMapPlaceholder": "Map preview will appear here after route calculation...",
        "ManualNoUrl": "No generated link",
        "ManualBtnGoogleMaps": "🌐 Google Maps",
        "ManualBtnCopyUrl": "📋 Copy Link",
        "ManualBtnSaveMapAs": "💾 Save Map As...",
        "BatchInputFile": "Input File (JSON/CSV/XLSX):",
        "BatchBtnBrowse": "📂 Browse File...",
        "BatchBtnReload": "🔄 Reload",
        "BatchNoFileLoaded": "No file loaded.",
        "BatchDefaultRouteType": "Default route type:",
        "BatchOptFromSource": "From Source / Default",
        "BatchBtnStart": "▶ Start Processing",
        "BatchBtnStop": "⏹ Stop",
        "BatchTabInputPreview": "📋 Input Data Preview",
        "BatchTabResults": "📊 Calculation Results",
        "BatchTabLog": "📝 Activity Log",
        "BatchColId": "ID",
        "BatchColName": "Route Name",
        "BatchColOrigin": "Origin (Start)",
        "BatchColDestination": "Destination (End)",
        "BatchColWaypoints": "Waypoints",
        "BatchColType": "Type",
        "BatchColDistance": "Distance (km)",
        "BatchColDuration": "Duration (min)",
        "BatchColStatus": "Status",
        "BatchColMap": "PNG Map",
        "BatchProgressReady": "Ready",
        "BatchBtnOpenOutputDir": "📂 Open Output Folder",
        "BatchBtnExportExcel": "📊 Export Excel",
        "BatchBtnExportCsv": "📄 CSV",
        "BatchBtnExportJson": "📋 JSON",
        "SettingsHeaderApi": "Google Maps API Key",
        "SettingsApiDesc": "Required for Geocoding API, Routes API v2, and Static Maps API.",
        "SettingsApiLabel": "API Key:",
        "SettingsBtnShow": "👁 Show",
        "SettingsBtnHide": "🔒 Hide",
        "SettingsBtnTestKey": "🔍 Test Key",
        "SettingsBtnTestingKey": "⏳ Testing...",
        "SettingsChkRemember": "Remember securely on this computer (DPAPI CurrentUser encryption)",
        "SettingsHeaderPreferences": "Default Generation Preferences",
        "SettingsDefaultRouteType": "Default route type:",
        "SettingsDefaultEmission": "Default engine type for Eco routes:",
        "SettingsDefaultMapSize": "Default dimensions for generated PNG map:",
        "SettingsOutputDir": "Results Output Folder:",
        "SettingsBtnBrowseOutputDir": "📂 Browse...",
        "SettingsHeaderLanguage": "Language & Localization",
        "SettingsLanguageLabel": "Application & Google Maps API Language:",
        "SettingsBtnOpenLangFile": "📂 Open Localization File (localization.json)",
        "SettingsBtnReloadLang": "🔄 Reload Languages",
        "SettingsBtnSave": "💾 SAVE SETTINGS",
        "SettingsBtnOpenLog": "📋 OPEN LOG FILE",
        "MapLabelTotal": "Total: ",
        "MapLabelType": "Type: ",
        "MapLabelContract": "Contract: ",
        "MapLabelDirection": "Direction: ",
        "MsgMissingApiKey": "Please enter an API key before testing.",
        "MsgMissingApiKeyTitle": "Missing API Key",
        "MsgMissingApiKeyPrompt": "Please enter and save a Google Maps API key in Settings.",
        "MsgMissingData": "Please enter both an origin (start) and a destination (end).",
        "MsgMissingDataTitle": "Missing Data",
        "MsgNoDataFile": "Please load a valid data file first (JSON, CSV, or Excel).",
        "MsgNoDataFileTitle": "No Data",
        "MsgMaxWaypoints": "Maximum number of waypoints is 25.",
        "MsgMaxWaypointsTitle": "Waypoint Limit",
        "MsgSettingsSaved": "Settings have been saved successfully.",
        "MsgSettingsSavedTitle": "Saved",
        "MsgUrlCopied": "Google Maps navigation link copied to clipboard.",
        "MsgUrlCopiedTitle": "Copied",
        "MsgMapSaved": "Map saved: {0}",
        "MsgMapSavedTitle": "Saved",
        "MsgNoExportResults": "No results to export.",
        "MsgNoExportResultsTitle": "Empty Results",
        "MsgExportExcelComplete": "Exported to Excel:\n{0}",
        "MsgExportCsvComplete": "Exported to CSV:\n{0}",
        "MsgExportJsonComplete": "Exported to JSON:\n{0}",
        "MsgExportTitle": "Export Complete",
        "MsgLangReloaded": "Language definitions reloaded successfully ({0} languages found).",
        "SettingsHeaderOverlay": "Map Overlay & Banners (Top / Bottom)",
        "SettingsOverlayDesc": "Configure whether to display top and bottom banner panels, and choose which properties appear on each panel, line order, and alignment.",
        "SettingsOverlayTopEnable": "Enable Top Banner",
        "SettingsOverlayBottomEnable": "Enable Bottom Banner",
        "SettingsOverlayBtnReset": "🔄 Reset to Default Layout",
        "OverlayColProperty": "Property",
        "OverlayColShow": "Show",
        "OverlayColPanel": "Panel",
        "OverlayColAlign": "Alignment",
        "OverlayColOrder": "Line / Order",
        "OverlayPropStartGeocoded": "Start Address (Geocoded)",
        "OverlayPropEndGeocoded": "End Address (Geocoded)",
        "OverlayPropDistance": "Total Distance",
        "OverlayPropDuration": "Total Time",
        "OverlayPropTimestamp": "Generation Timestamp",
        "OverlayPropRouteName": "Route Name",
        "OverlayPropRouteType": "Route Type",
        "OverlayPropWaypoints": "Intermediate Stops (Waypoints)",
        "OverlayPropStartRaw": "Start Address (Raw Input)",
        "OverlayPropEndRaw": "End Address (Raw Input)",
        "RouteTypeFastest": "Fastest",
        "RouteTypeShortest": "Shortest",
        "RouteTypeEco": "Eco",
        "ThemeToggle": "Theme:",
        "ThemeDark": "🌙 Dark",
        "ThemeLight": "☀️ Light",
        "ThemeToggleTip": "Toggle Light / Dark theme",
        "SettingsThemeLabel": "Application Theme (Color Scheme):",
        "BatchTabPoints": "📍 Points Detail",
        "PointsColRouteId": "Route ID",
        "PointsColRouteName": "Route Name",
        "PointsColOrder": "No.",
        "PointsColType": "Point Type",
        "PointsColOriginalAddress": "Original Address",
        "PointsColGeocodedAddress": "Geocoded Address",
        "PointsColGeocodeStatus": "Geocode Status",
        "PointsColMatchType": "Match Type",
        "PointsColIsFallback": "Fallback?",
        "PointsColLatitude": "Latitude",
        "PointsColLongitude": "Longitude"
      }
    },
    "de": {
      "DisplayName": "Deutsch",
      "GoogleCode": "de",
      "Strings": {
        "AppTitle": "Google Maps Routen- & Kartengenerator",
        "AppSubtitle": "Mehrpunkt-Fahrrouten: Schnellste, Kürzeste, Sparsamste | Import JSON, CSV, Excel",
        "ApiBadgeChecking": "API: Prüfe...",
        "ApiBadgeActive": "API: Aktiv",
        "ApiBadgeMissing": "API: Fehlender Schlüssel",
        "ApiBadgeError": "API: Fehler",
        "BtnQuickSettings": "⚙ API-Einstellungen",
        "FooterReady": "Bereit.",
        "FooterVersion": "Google Maps Routes v2.0",
        "TabManual": "📍 Manuelle Route",
        "TabBatch": "📁 Stapelverarbeitung",
        "TabSettings": "⚙ Einstellungen & API-Schlüssel",
        "ManualHeaderRoutePoints": "Routenpunkte",
        "ManualOrigin": "Startpunkt (Start / A):",
        "ManualWaypoints": "Zwischenstopps (optional bis zu 25):",
        "ManualWaypointsTooltip": "Adresse für Zwischenstopp eingeben und Hinzufügen klicken",
        "ManualBtnAdd": "➕ Hinzufügen",
        "ManualBtnUp": "▲ Hoch",
        "ManualBtnDown": "▼ Runter",
        "ManualBtnRemove": "✕ Entfernen",
        "ManualBtnClear": "🗑 Leeren",
        "ManualDestination": "Zielort (Ziel / B):",
        "ManualRouteName": "Routenname / Beschreibung:",
        "ManualHeaderOptimization": "Routenoptimierung",
        "ManualOptFastest": "⚡ Schnellste",
        "ManualOptShortest": "📏 Kürzeste",
        "ManualOptEco": "🌿 Sparsam (Eco)",
        "ManualEmission": "Fahrzeugantrieb (für Eco-Route):",
        "ManualFuelGasoline": "Benzin",
        "ManualFuelDiesel": "Diesel",
        "ManualFuelHybrid": "Hybrid",
        "ManualFuelElectric": "Elektrisch",
        "ManualTrafficAware": "Echtzeit-Verkehrsberücksichtigung (Live Traffic)",
        "ManualBtnCalculate": "🚀 ROUTE BERECHNEN & KARTE HERUNTERLADEN",
        "ManualBtnCalculating": "⏳ BERECHNE ROUTE...",
        "ManualStatDistance": "DISTANZ",
        "ManualStatDuration": "DAUER",
        "ManualStatType": "ROUTENTYP",
        "ManualStatusIdle": "Bereit",
        "ManualStatusCalculating": "Berechnung läuft...",
        "ManualStatusSuccess": "Route berechnet",
        "ManualStatusError": "Berechnungsfehler",
        "ManualMapPlaceholder": "Die Kartenvorschau wird hier nach der Routenberechnung angezeigt...",
        "ManualNoUrl": "Kein Link vorhanden",
        "ManualBtnGoogleMaps": "🌐 Google Maps",
        "ManualBtnCopyUrl": "📋 Link kopieren",
        "ManualBtnSaveMapAs": "💾 Karte speichern unter...",
        "BatchInputFile": "Eingabedatei (JSON/CSV/XLSX):",
        "BatchBtnBrowse": "📂 Datei durchsuchen...",
        "BatchBtnReload": "🔄 Neu laden",
        "BatchNoFileLoaded": "Keine Datei geladen.",
        "BatchDefaultRouteType": "Standard-Routentyp:",
        "BatchOptFromSource": "Aus Quelle / Standard",
        "BatchBtnStart": "▶ Verarbeitung starten",
        "BatchBtnStop": "⏹ Anhalten",
        "BatchTabInputPreview": "📋 Datenvorschau",
        "BatchTabResults": "📊 Berechnungsergebnisse",
        "BatchTabLog": "📝 Aktivitätsprotokoll",
        "BatchColId": "ID",
        "BatchColName": "Routenname",
        "BatchColOrigin": "Startpunkt (Start)",
        "BatchColDestination": "Zielort (Ende)",
        "BatchColWaypoints": "Stopps",
        "BatchColType": "Typ",
        "BatchColDistance": "Distanz (km)",
        "BatchColDuration": "Dauer (min)",
        "BatchColStatus": "Status",
        "BatchColMap": "PNG-Karte",
        "BatchProgressReady": "Bereit",
        "BatchBtnOpenOutputDir": "📂 Ausgabeordner öffnen",
        "BatchBtnExportExcel": "📊 Excel exportieren",
        "BatchBtnExportCsv": "📄 CSV",
        "BatchBtnExportJson": "📋 JSON",
        "SettingsHeaderApi": "Google Maps API-Schlüssel",
        "SettingsApiDesc": "Erforderlich für Geocoding API, Routes API v2 und Static Maps API.",
        "SettingsApiLabel": "API-Schlüssel:",
        "SettingsBtnShow": "👁 Anzeigen",
        "SettingsBtnHide": "🔒 Verbergen",
        "SettingsBtnTestKey": "🔍 Schlüssel testen",
        "SettingsBtnTestingKey": "⏳ Prüfe...",
        "SettingsChkRemember": "Auf diesem Computer sicher speichern (DPAPI CurrentUser-Verschlüsselung)",
        "SettingsHeaderPreferences": "Standardeinstellungen für Generierung",
        "SettingsDefaultRouteType": "Standard-Routentyp:",
        "SettingsDefaultEmission": "Standard-Antriebsart für Eco-Routen:",
        "SettingsDefaultMapSize": "Standardabmessungen für generierte PNG-Karte:",
        "SettingsOutputDir": "Ausgabeordner für Ergebnisse:",
        "SettingsBtnBrowseOutputDir": "📂 Durchsuchen...",
        "SettingsHeaderLanguage": "Sprache & Lokalisierung",
        "SettingsLanguageLabel": "Sprache für Anwendung & Google Maps API:",
        "SettingsBtnOpenLangFile": "📂 Lokalisierungsdatei öffnen (localization.json)",
        "SettingsBtnReloadLang": "🔄 Sprachen neu laden",
        "SettingsBtnSave": "💾 EINSTELLUNGEN SPEICHERN",
        "SettingsBtnOpenLog": "📋 PROTOKOLLDATEI ÖFFNEN",
        "MapLabelTotal": "Gesamt: ",
        "MapLabelType": "Typ: ",
        "MapLabelContract": "Vertrag: ",
        "MapLabelDirection": "Richtung: ",
        "MsgMissingApiKey": "Bitte vor dem Testen einen API-Schlüssel eingeben.",
        "MsgMissingApiKeyTitle": "Fehlender API-Schlüssel",
        "MsgMissingApiKeyPrompt": "Bitte geben Sie einen Google Maps API-Schlüssel in den Einstellungen ein.",
        "MsgMissingData": "Bitte sowohl einen Startpunkt als auch einen Zielort eingeben.",
        "MsgMissingDataTitle": "Fehlende Angaben",
        "MsgNoDataFile": "Bitte laden Sie zuerst eine gültige Datendatei (JSON, CSV oder Excel).",
        "MsgNoDataFileTitle": "Keine Daten",
        "MsgMaxWaypoints": "Die maximale Anzahl an Zwischenstopps beträgt 25.",
        "MsgMaxWaypointsTitle": "Stopp-Limit",
        "MsgSettingsSaved": "Einstellungen wurden erfolgreich gespeichert.",
        "MsgSettingsSavedTitle": "Gespeichert",
        "MsgUrlCopied": "Google Maps Navigationslink in die Zwischenablage kopiert.",
        "MsgUrlCopiedTitle": "Kopiert",
        "MsgMapSaved": "Karte gespeichert: {0}",
        "MsgMapSavedTitle": "Gespeichert",
        "MsgNoExportResults": "Keine Ergebnisse zum Exportieren vorhanden.",
        "MsgNoExportResultsTitle": "Leere Ergebnisse",
        "MsgExportExcelComplete": "Nach Excel exportiert:\n{0}",
        "MsgExportCsvComplete": "Nach CSV exportiert:\n{0}",
        "MsgExportJsonComplete": "Nach JSON exportiert:\n{0}",
        "MsgExportTitle": "Export abgeschlossen",
        "MsgLangReloaded": "Sprachdefinitionen erfolgreich neu geladen ({0} Sprachen gefunden).",
        "MsgLangReloadedTitle": "Sprachen neu geladen",
        "SettingsHeaderOverlay": "Karten-Overlay & Banner (Oben / Unten)",
        "SettingsOverlayDesc": "Legen Sie fest, ob obere und untere Banner angezeigt werden und welche Eigenschaften in welchem Bereich, in welcher Zeile und Ausrichtung erscheinen.",
        "SettingsOverlayTopEnable": "Oberes Banner aktivieren",
        "SettingsOverlayBottomEnable": "Unteres Banner aktivieren",
        "SettingsOverlayBtnReset": "🔄 Standardlayout wiederherstellen",
        "OverlayColProperty": "Eigenschaft",
        "OverlayColShow": "Aktiv",
        "OverlayColPanel": "Bereich",
        "OverlayColAlign": "Ausrichtung",
        "OverlayColOrder": "Zeile / Reihenfolge",
        "OverlayPropStartGeocoded": "Startadresse (Geokodiert)",
        "OverlayPropEndGeocoded": "Zieladresse (Geokodiert)",
        "OverlayPropDistance": "Gesamtdistanz",
        "OverlayPropDuration": "Gesamtzeit",
        "OverlayPropTimestamp": "Erstellungszeitstempel",
        "OverlayPropRouteName": "Routenname",
        "OverlayPropRouteType": "Routentyp",
        "OverlayPropWaypoints": "Zwischenstopps (Waypoints)",
        "OverlayPropStartRaw": "Startadresse (Rohdaten)",
        "OverlayPropEndRaw": "Zieladresse (Rohdaten)",
        "RouteTypeFastest": "Schnellste",
        "RouteTypeShortest": "Kürzeste",
        "RouteTypeEco": "Eco",
        "ThemeToggle": "Design:",
        "ThemeDark": "🌙 Dunkel",
        "ThemeLight": "☀️ Hell",
        "ThemeToggleTip": "Hell- / Dunkelmodus umschalten",
        "SettingsThemeLabel": "Anwendungs-Design (Farbschema):",
        "BatchTabPoints": "📍 Routenpunkte",
        "PointsColRouteId": "Routen-ID",
        "PointsColRouteName": "Routenname",
        "PointsColOrder": "Nr.",
        "PointsColType": "Punkttyp",
        "PointsColOriginalAddress": "Originaladresse",
        "PointsColGeocodedAddress": "Geokodierte Adresse",
        "PointsColGeocodeStatus": "Geokodierungsstatus",
        "PointsColMatchType": "Übereinstimmungstyp",
        "PointsColIsFallback": "Fallback?",
        "PointsColLatitude": "Breitengrad",
        "PointsColLongitude": "Längengrad"
      }
    },
    "pl": {
      "DisplayName": "Polski",
      "GoogleCode": "pl",
      "Strings": {
        "AppTitle": "Generator Tras i Map Google Maps",
        "AppSubtitle": "Wielopunktowe trasy samochodowe: Najszybsza, Najkrótsza, Eko | Import JSON, CSV, Excel",
        "ApiBadgeChecking": "API: Sprawdzanie...",
        "ApiBadgeActive": "API: Aktywny",
        "ApiBadgeMissing": "API: Brak klucza",
        "ApiBadgeError": "API: Błąd",
        "BtnQuickSettings": "⚙ Ustawienia API",
        "FooterReady": "Gotowy.",
        "FooterVersion": "Google Maps Routes v2.0",
        "TabManual": "📍 Trasa ręczna",
        "TabBatch": "📁 Przetwarzanie wsadowe",
        "TabSettings": "⚙ Ustawienia i klucz API",
        "ManualHeaderRoutePoints": "Punkty trasy",
        "ManualOrigin": "Punkt początkowy (Start / A):",
        "ManualWaypoints": "Punkty pośrednie (opcjonalnie do 25):",
        "ManualWaypointsTooltip": "Wpisz adres punktu pośredniego i kliknij Dodaj",
        "ManualBtnAdd": "➕ Dodaj",
        "ManualBtnUp": "▲ W górę",
        "ManualBtnDown": "▼ W dół",
        "ManualBtnRemove": "✕ Usuń",
        "ManualBtnClear": "🗑 Wyczyść",
        "ManualDestination": "Punkt docelowy (Cel / B):",
        "ManualRouteName": "Nazwa trasy / Opis:",
        "ManualHeaderOptimization": "Optymalizacja trasy",
        "ManualOptFastest": "⚡ Najszybsza",
        "ManualOptShortest": "📏 Najkrótsza",
        "ManualOptEco": "🌿 Ekologiczna (Eko)",
        "ManualEmission": "Typ napędu pojazdu (dla trasy Eko):",
        "ManualFuelGasoline": "Benzyna",
        "ManualFuelDiesel": "Diesel",
        "ManualFuelHybrid": "Hybryda",
        "ManualFuelElectric": "Elektryczny",
        "ManualTrafficAware": "Uwzględniaj ruch drogowy w czasie rzeczywistym (Live Traffic)",
        "ManualBtnCalculate": "🚀 OBLICZ TRASĘ I POBIERZ MAPĘ",
        "ManualBtnCalculating": "⏳ OBLICZANIE TRASY...",
        "ManualStatDistance": "DYSTANS",
        "ManualStatDuration": "CZAS TRWANIA",
        "ManualStatType": "TYP TRASY",
        "ManualStatusIdle": "Gotowy",
        "ManualStatusCalculating": "Obliczanie trasy...",
        "ManualStatusSuccess": "Trasa obliczona",
        "ManualStatusError": "Błąd obliczania",
        "ManualMapPlaceholder": "Podgląd mapy pojawi się tutaj po obliczeniu trasy...",
        "ManualNoUrl": "Brak wygenerowanego linku",
        "ManualBtnGoogleMaps": "🌐 Google Maps",
        "ManualBtnCopyUrl": "📋 Kopiuj link",
        "ManualBtnSaveMapAs": "💾 Zapisz mapę jako...",
        "BatchInputFile": "Plik wejściowy (JSON/CSV/XLSX):",
        "BatchBtnBrowse": "📂 Przeglądaj plik...",
        "BatchBtnReload": "🔄 Przeładuj",
        "BatchNoFileLoaded": "Brak wczytanego pliku.",
        "BatchDefaultRouteType": "Domyślny typ trasy:",
        "BatchOptFromSource": "Ze źródła / Domyślny",
        "BatchBtnStart": "▶ Rozpocznij przetwarzanie",
        "BatchBtnStop": "⏹ Zatrzymaj",
        "BatchTabInputPreview": "📋 Podgląd danych wejściowych",
        "BatchTabResults": "📊 Wyniki obliczeń",
        "BatchTabLog": "📝 Dziennik zdarzeń",
        "BatchColId": "ID",
        "BatchColName": "Nazwa trasy",
        "BatchColOrigin": "Start (Początek)",
        "BatchColDestination": "Cel (Koniec)",
        "BatchColWaypoints": "Punkty",
        "BatchColType": "Typ",
        "BatchColDistance": "Dystans (km)",
        "BatchColDuration": "Czas (min)",
        "BatchColStatus": "Status",
        "BatchColMap": "Mapa PNG",
        "BatchProgressReady": "Gotowy",
        "BatchBtnOpenOutputDir": "📂 Otwórz folder wyników",
        "BatchBtnExportExcel": "📊 Eksportuj do Excel",
        "BatchBtnExportCsv": "📄 CSV",
        "BatchBtnExportJson": "📋 JSON",
        "SettingsHeaderApi": "Klucz Google Maps API",
        "SettingsApiDesc": "Wymagany do Geocoding API, Routes API v2 oraz Static Maps API.",
        "SettingsApiLabel": "Klucz API:",
        "SettingsBtnShow": "👁 Pokaż",
        "SettingsBtnHide": "🔒 Ukryj",
        "SettingsBtnTestKey": "🔍 Testuj klucz",
        "SettingsBtnTestingKey": "⏳ Sprawdzanie...",
        "SettingsChkRemember": "Zapamiętaj bezpiecznie na tym komputerze (szyfrowanie DPAPI CurrentUser)",
        "SettingsHeaderPreferences": "Domyślne preferencje generowania",
        "SettingsDefaultRouteType": "Domyślny typ trasy:",
        "SettingsDefaultEmission": "Domyślny typ napędu dla tras Eko:",
        "SettingsDefaultMapSize": "Domyślne wymiary generowanej mapy PNG:",
        "SettingsOutputDir": "Folder zapisu wyników:",
        "SettingsBtnBrowseOutputDir": "📂 Przeglądaj...",
        "SettingsHeaderLanguage": "Język i lokalizacja",
        "SettingsLanguageLabel": "Język aplikacji oraz zapytań Google Maps API:",
        "SettingsBtnOpenLangFile": "📂 Otwórz plik lokalizacji (localization.json)",
        "SettingsBtnReloadLang": "🔄 Przeładuj języki",
        "SettingsBtnSave": "💾 ZAPISZ USTAWIENIA",
        "SettingsBtnOpenLog": "📋 OTWÓRZ PLIK LOGU",
        "MapLabelTotal": "Razem: ",
        "MapLabelType": "Typ: ",
        "MapLabelContract": "Umowa: ",
        "MapLabelDirection": "Kierunek: ",
        "MsgMissingApiKey": "Wprowadź klucz API przed rozpoczęciem testu.",
        "MsgMissingApiKeyTitle": "Brak klucza API",
        "MsgMissingApiKeyPrompt": "Wprowadź i zapisz klucz Google Maps API w Ustawieniach.",
        "MsgMissingData": "Wprowadź zarówno punkt początkowy (start), jak i docelowy (cel).",
        "MsgMissingDataTitle": "Brakujące dane",
        "MsgNoDataFile": "Najpierw wczytaj poprawny plik z danymi (JSON, CSV lub Excel).",
        "MsgNoDataFileTitle": "Brak danych",
        "MsgMaxWaypoints": "Maksymalna dopuszczalna liczba punktów pośrednich wynosi 25.",
        "MsgMaxWaypointsTitle": "Limit punktów pośrednich",
        "MsgSettingsSaved": "Ustawienia zostały pomyślnie zapisane.",
        "MsgSettingsSavedTitle": "Zapisano",
        "MsgUrlCopied": "Link do nawigacji Google Maps został skopiowany do schowka.",
        "MsgUrlCopiedTitle": "Skopiowano",
        "MsgMapSaved": "Zapisano mapę: {0}",
        "MsgMapSavedTitle": "Zapisano",
        "MsgNoExportResults": "Brak wyników do wyeksportowania.",
        "MsgNoExportResultsTitle": "Puste wyniki",
        "MsgExportExcelComplete": "Wyeksportowano do pliku Excel:\n{0}",
        "MsgExportCsvComplete": "Wyeksportowano do pliku CSV:\n{0}",
        "MsgExportJsonComplete": "Wyeksportowano do pliku JSON:\n{0}",
        "MsgExportTitle": "Eksport zakończony",
        "MsgLangReloaded": "Pomyślnie przeładowano definicje językowe (znaleziono {0} języków).",
        "MsgLangReloadedTitle": "Przeładowano języki",
        "SettingsHeaderOverlay": "Nakładka na mapie i nagłówki (Góra / Dół)",
        "SettingsOverlayDesc": "Skonfiguruj wyświetlanie górnego i dolnego paska informacyjnego oraz wybierz, które dane mają się pojawić w którym panelu, wierszu i wyrównaniu.",
        "SettingsOverlayTopEnable": "Włącz górny nagłówek na mapie",
        "SettingsOverlayBottomEnable": "Włącz dolny nagłówek na mapie",
        "SettingsOverlayBtnReset": "🔄 Przywróć domyślny układ",
        "OverlayColProperty": "Właściwość",
        "OverlayColShow": "Pokaż",
        "OverlayColPanel": "Panel",
        "OverlayColAlign": "Wyrównanie",
        "OverlayColOrder": "Wiersz / Kolejność",
        "OverlayPropStartGeocoded": "Adres początkowy (Zgeokodowany)",
        "OverlayPropEndGeocoded": "Adres docelowy (Zgeokodowany)",
        "OverlayPropDistance": "Całkowity dystans",
        "OverlayPropDuration": "Całkowity czas",
        "OverlayPropTimestamp": "Data i czas wygenerowania",
        "OverlayPropRouteName": "Nazwa trasy",
        "OverlayPropRouteType": "Typ trasy",
        "OverlayPropWaypoints": "Punkty pośrednie (przystanki)",
        "OverlayPropStartRaw": "Adres początkowy (Wprowadzony)",
        "OverlayPropEndRaw": "Adres docelowy (Wprowadzony)",
        "RouteTypeFastest": "Najszybsza",
        "RouteTypeShortest": "Najkrótsza",
        "RouteTypeEco": "Eko",
        "ThemeToggle": "Motyw:",
        "ThemeDark": "🌙 Ciemny",
        "ThemeLight": "☀️ Jasny",
        "ThemeToggleTip": "Przełącz tryb jasny / ciemny",
        "SettingsThemeLabel": "Motyw aplikacji (schemat kolorów):",
        "BatchTabPoints": "📍 Punkty Tras",
        "PointsColRouteId": "ID Trasy",
        "PointsColRouteName": "Nazwa Trasy",
        "PointsColOrder": "Lp.",
        "PointsColType": "Typ Punktu",
        "PointsColOriginalAddress": "Adres Oryginalny",
        "PointsColGeocodedAddress": "Adres Zgeokodowany",
        "PointsColGeocodeStatus": "Status Geokodowania",
        "PointsColMatchType": "Typ Dopasowania",
        "PointsColIsFallback": "Fallback?",
        "PointsColLatitude": "Szerokość",
        "PointsColLongitude": "Długość"
      }
    }
  }
}
'@

function Load-LocalizationConfig {
    [CmdletBinding()]
    param()

    $script:LanguagesCatalog = [ordered]@{}
    $script:DefaultStrings = @{}

    # 1. Parse embedded localization catalog (guarantees all languages & keys in-memory)
    if (-not [string]::IsNullOrWhiteSpace($script:EmbeddedLocalizationJson)) {
        try {
            $embParsed = $script:EmbeddedLocalizationJson | ConvertFrom-Json
            if ($embParsed.Languages) {
                foreach ($prop in $embParsed.Languages.PSObject.Properties) {
                    $code = $prop.Name.ToLower()
                    $langData = $prop.Value
                    $disp = if ($langData.DisplayName) { [string]$langData.DisplayName } else { $code.ToUpper() }
                    $gCode = if ($langData.GoogleCode) { [string]$langData.GoogleCode } else { $code }

                    $strMap = @{}
                    if ($langData.Strings) {
                        foreach ($sProp in $langData.Strings.PSObject.Properties) {
                            $strMap[$sProp.Name] = [string]$sProp.Value
                        }
                    }

                    $script:LanguagesCatalog[$code] = [PSCustomObject]@{
                        Code        = $code
                        DisplayName = $disp
                        GoogleCode  = $gCode
                        Strings     = $strMap
                    }
                }
            }
        } catch {
            Write-AppLog "Error parsing embedded localization catalog: $(#Requires -Version 5.1
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force -ErrorAction SilentlyContinue
<#
.SYNOPSIS
    Google Maps Route & Map Generator — Zaawansowana aplikacja WPF Dark Mode.
    Obsługuje ręczne wprowadzanie tras (Start, Cel, Punkty pośrednie, Fastest/Shortest/Eco)
    oraz wsadowe przetwarzanie plików danych (JSON, CSV, Excel).

.DESCRIPTION
    Funkcje:
      - Tryb ręczny (Manual Input):
          * Punkt startowy i punkt końcowy (z walidacją i geokodowaniem)
          * Dynamiczna lista punktów pośrednich (dodawanie, usuwanie, zmiana kolejności)
          * Wybór optymalizacji: Najszybsza (Fastest), Najkrótsza (Shortest), Ekologiczna (Eco)
          * Wybór typu napędu dla trasy Eco (Benzyna, Diesel, Hybryda, Elektryczny)
          * Natychmiastowe obliczanie trasy, odległości (km) i czasu (min/godz)
          * Interaktywny podgląd mapy statycznej PNG z trasą i ponumerowanymi znacznikami
          * Kopiowanie i bezpośrednie otwieranie linku do nawigacji Google Maps w przeglądarce
      - Tryb wsadowy (Data Source / Batch):
          * Obsługa formatów Excel (.xlsx, .xls), CSV (.csv, .tsv), JSON (.json)
          * Automatyczne wykrywanie schematu pliku i mapowanie kolumn z możliwością korekty
          * Tabela podglądu danych wejściowych (DataGrid)
          * Pasek postępu, procenty, czas, asynchroniczny log zdarzeń w czasie rzeczywistym
          * Tabela wyników ze statusem i bezpośrednim dostępem do map
          * Eksport raportów zbiorczych do Excel, CSV i JSON
      - Bezpieczeństwo i ustawienia:
          * Szyfrowane przechowywanie klucza Google Maps API (Windows DPAPI per-user)
          * Asynchroniczny tester poprawności klucza API (nie zawiesza interfejsu)
          * Konfiguracja domyślnych wymiarów mapy, katalogów i typu trasy
      - Zgodność ze standardem PS2EXE (samodzielny plik .EXE bez zewnętrznych zależności).

.NOTES
    Encoding: UTF-8 with BOM
#>

# ── 1. Wymuszenie protokołów TLS 1.2 / TLS 1.1 dla zapytań HTTPS ───────────────
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls

# ── 2. Wymuszenie trybu STA dla WPF ──────────────────────────────────────────
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) {
    $currentProcess = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    if ($currentProcess -match 'powershell\.exe|pwsh\.exe') {
        Start-Process -FilePath $currentProcess -ArgumentList "-NoProfile -STA -ExecutionPolicy Bypass -File `"$PSCommandPath`""
        exit
    }
}

# ── 3. Ładowanie bibliotek GUI, Drawing i Security ───────────────────────────
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms, System.Drawing, System.Security

# DWM Dark Mode dla paska tytułu okna Windows 10/11
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class DwmDarkWindow {
    [DllImport("dwmapi.dll", PreserveSig = true)]
    public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
}
"@ -ErrorAction SilentlyContinue

# ══════════════════════════════════════════════════════════════════════════════
# 4. SAMODZIELNE FUNKCJE BAZOWE (EMBEDDED DLA ZGODNOŚCI Z PS2EXE)
# ══════════════════════════════════════════════════════════════════════════════

function Protect-SecretString {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$PlainText)
    if ([string]::IsNullOrEmpty($PlainText)) { return $null }
    try {
        Add-Type -AssemblyName System.Security
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($PlainText)
        $protected = [System.Security.Cryptography.ProtectedData]::Protect(
            $bytes, $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
        return [Convert]::ToBase64String($protected)
    }
    catch {
        try {
            $sec = ConvertTo-SecureString -String $PlainText -AsPlainText -Force
            return (ConvertFrom-SecureString -SecureString $sec)
        } catch {
            return $null
        }
    }
}

function Unprotect-SecretString {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$EncryptedText)
    if ([string]::IsNullOrWhiteSpace($EncryptedText)) { return $null }
    try {
        Add-Type -AssemblyName System.Security
        $bytes = [Convert]::FromBase64String($EncryptedText)
        $unprotected = [System.Security.Cryptography.ProtectedData]::Unprotect(
            $bytes, $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
        return [System.Text.Encoding]::UTF8.GetString($unprotected)
    }
    catch {
        try {
            $sec = ConvertTo-SecureString -String $EncryptedText
            $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)
            $str = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
            return $str
        }
        catch {
            return $null
        }
    }
}

function Test-GoogleApiKey {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ApiKey,
        [Parameter()][string]$LanguageCode = 'en'
    )
    if ([string]::IsNullOrWhiteSpace($ApiKey)) {
        return [PSCustomObject]@{ Valid = $false; Message = 'API key is empty.' }
    }
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls
    try {
        $lang = if ($LanguageCode) { ($LanguageCode -split '[-_]')[0].ToLower() } else { 'en' }
        $Url = "https://maps.googleapis.com/maps/api/geocode/json?address=Warszawa&language=$lang&key=$ApiKey"
        $Resp = Invoke-RestMethod -Uri $Url -Method GET -TimeoutSec 15
        if ($Resp.status -eq 'OK' -or $Resp.status -eq 'ZERO_RESULTS') {
            return [PSCustomObject]@{ Valid = $true; Message = 'Google Maps API key is valid and active.' }
        }
        elseif ($Resp.status -eq 'REQUEST_DENIED') {
            $msg = if ($Resp.error_message) { $Resp.error_message } else { 'Request denied by Google API.' }
            return [PSCustomObject]@{ Valid = $false; Message = "Unauthorized: $msg" }
        }
        else {
            return [PSCustomObject]@{ Valid = $false; Message = "Status API: $($Resp.status)" }
        }
    }
    catch {
        return [PSCustomObject]@{ Valid = $false; Message = "Connection error: $($_.Exception.Message)" }
    }
}

function Select-InputDataFile {
    param([string]$InitialDirectory)
    Add-Type -AssemblyName System.Windows.Forms
    $Dialog = [System.Windows.Forms.OpenFileDialog]::new()
    $Dialog.Title = 'Select route data file (JSON, CSV, Excel)'
    $Dialog.Filter = 'All Supported Files (*.xlsx;*.xls;*.csv;*.tsv;*.json)|*.xlsx;*.xls;*.csv;*.tsv;*.json|Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|CSV/TSV Files (*.csv;*.tsv)|*.csv;*.tsv|JSON Files (*.json)|*.json|All Files (*.*)|*.*'

    $chosenDir = $null
    if ($InitialDirectory -and (Test-Path $InitialDirectory)) {
        $chosenDir = $InitialDirectory
    }
    elseif ($script:LastDataDirectory -and (Test-Path $script:LastDataDirectory)) {
        $chosenDir = $script:LastDataDirectory
    }
    elseif ($script:Config -and $script:Config.LastInputFolder -and (Test-Path $script:Config.LastInputFolder)) {
        $chosenDir = $script:Config.LastInputFolder
    }
    elseif ($script:Config -and $script:Config.LastInputPath -and (Test-Path (Split-Path $script:Config.LastInputPath -Parent))) {
        $chosenDir = Split-Path $script:Config.LastInputPath -Parent
    }
    else {
        $samplesDir = if ($PSScriptRoot) { Join-Path $PSScriptRoot 'Samples' } else { Join-Path (Get-Location) 'Samples' }
        if (Test-Path $samplesDir) {
            $chosenDir = $samplesDir
        } else {
            $chosenDir = [Environment]::GetFolderPath('MyDocuments')
        }
    }

    $Dialog.InitialDirectory = $chosenDir
    $Dialog.RestoreDirectory = $true
    $Result = $Dialog.ShowDialog()
    if ($Result -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:LastDataDirectory = Split-Path $Dialog.FileName -Parent
        if ($script:Config) {
            $script:Config.LastInputFolder = $script:LastDataDirectory
            $script:Config.LastInputPath = $Dialog.FileName
        }
        return $Dialog.FileName
    }
    return $null
}

function Get-AddressComponentValue {
    param([object[]]$Components, [string[]]$Types)
    $Matches = @($Components) | Where-Object {
        $_ -and $_.PSObject.Properties.Name -contains 'types' -and
        (@($_.types) | Where-Object { $_ -in $Types } | Select-Object -First 1)
    } | Select-Object -First 1
    if (-not $Matches) { return $null }
    foreach ($FieldName in @('long_name', 'short_name', 'name')) {
        if ($Matches.PSObject.Properties.Name -contains $FieldName) {
            $Value = [string]$Matches.$FieldName
            if (-not [string]::IsNullOrWhiteSpace($Value)) { return $Value }
        }
    }
    return $null
}

function Get-AddressCoordinates {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Address,
        [Parameter(Mandatory)][string]$ApiKey,
        [Parameter()][string]$LanguageCode = 'en',
        [Parameter()][switch]$RequireStreetNumber
    )
    if ([string]::IsNullOrWhiteSpace($Address)) { return $null }

    if ($Address.Trim() -match '^\s*([+-]?\d+(?:\.\d+)?)\s*[,;\s]\s*([+-]?\d+(?:\.\d+)?)\s*$') {
        $lat = [double]$Matches[1]
        $lng = [double]$Matches[2]
        return [PSCustomObject]@{
            Latitude             = $lat
            Longitude            = $lng
            FormattedAddress     = "$lat, $lng"
            UlicaINumer          = $null
            KodPocztowy          = $null
            Miasto               = $null
            MatchType            = 'COORDINATES'
            PartialMatch         = $false
            Status               = 'OK'
            ErrorMessage         = $null
        }
    }

    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls
    $EncodedAddress = [System.Uri]::EscapeDataString($Address.Trim())
    $lang = if ($LanguageCode) { ($LanguageCode -split '[-_]')[0].ToLower() } else { 'en' }
    $Url = "https://maps.googleapis.com/maps/api/geocode/json?address=$EncodedAddress&language=$lang&key=$ApiKey"
    try {
        $Response = Invoke-RestMethod -Uri $Url -Method GET -TimeoutSec 30
        $Results = @($Response.results)
        if ($Response.status -eq 'OK' -and $Results.Count -gt 0) {
            $ResultItem = $Results[0]
            $Location   = $ResultItem.geometry.location

            $Components   = @($ResultItem.address_components)
            $StreetNumber = Get-AddressComponentValue -Components $Components -Types @('street_number')
            $Route        = Get-AddressComponentValue -Components $Components -Types @('route')
            $PostalCode   = Get-AddressComponentValue -Components $Components -Types @('postal_code')
            $City         = Get-AddressComponentValue -Components $Components -Types @('locality', 'postal_town')
            if ([string]::IsNullOrWhiteSpace($City)) {
                $City = Get-AddressComponentValue -Components $Components -Types @('administrative_area_level_3', 'administrative_area_level_2')
            }

            $StreetWithNumber = if ($Route -and $StreetNumber) { "$Route $StreetNumber" }
                                elseif ($Route) { $Route }
                                elseif ($StreetNumber) { $StreetNumber }
                                else { $null }

            $FormattedAddress = if ($lang -eq 'pl') {
                $ResultItem.formatted_address -replace ',\s*Poland$', ', Polska' -replace '\bPoland\b', 'Polska'
            } else {
                $ResultItem.formatted_address
            }
            $LocationType     = if ($ResultItem.geometry -and $ResultItem.geometry.location_type) { [string]$ResultItem.geometry.location_type } else { 'APPROXIMATE' }
            $PartialMatch     = if ($ResultItem.PSObject.Properties.Name -contains 'partial_match') { [bool]$ResultItem.partial_match } else { $false }

            return [PSCustomObject]@{
                Latitude             = [double]$Location.lat
                Longitude            = [double]$Location.lng
                FormattedAddress     = $FormattedAddress
                UlicaINumer          = $StreetWithNumber
                KodPocztowy          = $PostalCode
                Miasto               = $City
                MatchType            = $LocationType
                PartialMatch         = $PartialMatch
                Status               = 'OK'
                ErrorMessage         = $null
            }
        }
        else {
            return [PSCustomObject]@{
                Latitude             = $null; Longitude = $null; FormattedAddress = $null
                UlicaINumer          = $null; KodPocztowy = $null; Miasto = $null
                MatchType            = $null; PartialMatch = $null
                Status               = $Response.status
                ErrorMessage         = $Response.error_message
            }
        }
    }
    catch {
        $Message = $_.Exception.Message
        return [PSCustomObject]@{
            Latitude             = $null; Longitude = $null; FormattedAddress = $null
            UlicaINumer          = $null; KodPocztowy = $null; Miasto = $null
            MatchType            = $null; PartialMatch = $null
            Status               = "EXCEPTION: $Message"
            ErrorMessage         = $Message
        }
    }
}

function Get-GeocodeStatusDescription {
    [CmdletBinding()]
    param(
        [Parameter()][object]$Geo
    )
    if (-not $Geo) { return 'NOT_PROCESSED' }
    if ($Geo.Status -eq 'OK') {
        if ($Geo.PartialMatch -and $Geo.MatchType -in 'APPROXIMATE', 'GEOMETRIC_CENTER') {
            return "OK (Fallback: Approximate / Partial Match - $($Geo.MatchType))"
        }
        elseif ($Geo.PartialMatch) {
            return "OK (Fallback: Partial Match - $($Geo.MatchType))"
        }
        elseif ($Geo.MatchType -eq 'APPROXIMATE') {
            return 'OK (Fallback: Approximate)'
        }
        elseif ($Geo.MatchType -eq 'GEOMETRIC_CENTER') {
            return 'OK (Fallback: Geometric Center)'
        }
        elseif ($Geo.MatchType -eq 'RANGE_INTERPOLATED') {
            return 'OK (Interpolated)'
        }
        elseif ($Geo.MatchType -eq 'ROOFTOP') {
            return 'OK (Exact - ROOFTOP)'
        }
        elseif ($Geo.MatchType -eq 'COORDINATES') {
            return 'OK (Coordinates)'
        }
        else {
            return "OK ($($Geo.MatchType))"
        }
    }
    elseif ($Geo.Status -eq 'ZERO_RESULTS') {
        return 'ZERO_RESULTS (Address Not Found)'
    }
    else {
        return [string]$Geo.Status
    }
}

function Get-CarRouteData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][double]$OriginLat,
        [Parameter(Mandatory)][double]$OriginLng,
        [Parameter(Mandatory)][double]$DestLat,
        [Parameter(Mandatory)][double]$DestLng,
        [Parameter(Mandatory)][string]$ApiKey,
        [Parameter()][object[]]$IntermediatePoints = @(),
        [Parameter()][ValidateSet('Fastest', 'Shortest', 'Eco')][string]$RouteType = 'Fastest',
        [Parameter()][ValidateSet('GASOLINE', 'DIESEL', 'HYBRID', 'ELECTRIC')][string]$EmissionType = 'GASOLINE',
        [Parameter()][string]$LanguageCode = 'en',
        [Parameter()][string]$Units = 'METRIC',
        [Parameter()][switch]$TrafficAware
    )

    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls
    $RoutesUrl = 'https://routes.googleapis.com/directions/v2:computeRoutes'

    $RequestBody = [ordered]@{
        origin       = @{ location = @{ latLng = @{ latitude = $OriginLat; longitude = $OriginLng } } }
        destination  = @{ location = @{ latLng = @{ latitude = $DestLat; longitude = $DestLng } } }
        travelMode   = 'DRIVE'
        languageCode = if ($LanguageCode) { $LanguageCode } else { 'en' }
        units        = $Units
    }

    $HasIntermediates = $false
    if ($null -ne $IntermediatePoints -and @($IntermediatePoints).Count -gt 0) {
        $IntermediatesList = [System.Collections.Generic.List[hashtable]]::new()
        foreach ($pt in $IntermediatePoints) {
            if ($null -ne $pt -and $pt.Latitude -and $pt.Longitude) {
                $IntermediatesList.Add(@{
                    location = @{
                        latLng = @{
                            latitude  = [double]$pt.Latitude
                            longitude = [double]$pt.Longitude
                        }
                    }
                })
            }
        }
        if ($IntermediatesList.Count -gt 0) {
            $RequestBody['intermediates'] = $IntermediatesList
            $HasIntermediates = $true
        }
    }

    switch ($RouteType) {
        'Fastest' {
            $RequestBody['routingPreference'] = if ($TrafficAware) { 'TRAFFIC_AWARE' } else { 'TRAFFIC_UNAWARE' }
            if (-not $HasIntermediates) { $RequestBody['computeAlternativeRoutes'] = $true }
        }
        'Shortest' {
            $RequestBody['routingPreference'] = 'TRAFFIC_UNAWARE'
            if (-not $HasIntermediates) { $RequestBody['computeAlternativeRoutes'] = $true }
        }
        'Eco' {
            $RequestBody['routingPreference'] = 'TRAFFIC_AWARE_OPTIMAL'
            $RequestBody['requestedReferenceRoutes'] = @('FUEL_EFFICIENT')
            $RequestBody['routeModifiers'] = @{
                vehicleInfo = @{ emissionType = $EmissionType }
            }
        }
    }

    $Headers = @{
        'X-Goog-Api-Key'   = $ApiKey
        'Content-Type'     = 'application/json'
        'X-Goog-FieldMask' = 'routes.duration,routes.distanceMeters,routes.polyline.encodedPolyline,routes.description,routes.routeLabels'
    }

    try {
        $JsonBody = $RequestBody | ConvertTo-Json -Depth 10
        $Response = Invoke-RestMethod -Uri $RoutesUrl -Method POST -Headers $Headers -Body $JsonBody -TimeoutSec 60

        $Routes = @($Response.routes)
        if ($Routes.Count -eq 0) {
            return [PSCustomObject]@{
                OdlegloscKm     = $null
                CzasMin         = $null
                DurationSeconds = $null
                EncodedPolyline = $null
                RouteType       = $RouteType
                RouteLabels     = @()
                Status          = 'NO_ROUTES'
                ErrorMessage    = 'Google Routes API did not return any routes.'
            }
        }

        $SelectedRoute = $null
        if ($RouteType -eq 'Shortest') {
            $SelectedRoute = $Routes | Sort-Object -Property { [int64]($_.distanceMeters) } | Select-Object -First 1
        }
        elseif ($RouteType -eq 'Eco') {
            $EcoRoute = $Routes | Where-Object {
                $_.routeLabels -and (@($_.routeLabels) -contains 'FUEL_EFFICIENT')
            } | Select-Object -First 1

            $SelectedRoute = if ($EcoRoute) { $EcoRoute } else { $Routes[0] }
        }
        else {
            $SelectedRoute = $Routes | Sort-Object -Property {
                if ($_.duration) { [double]($_.duration.TrimEnd('s')) } else { [double]::MaxValue }
            } | Select-Object -First 1
        }

        $DistanceKm = if ($SelectedRoute.distanceMeters) { [math]::Round([double]$SelectedRoute.distanceMeters / 1000.0, 2) } else { $null }
        $DurationSec = if ($SelectedRoute.duration) { [double]($SelectedRoute.duration.TrimEnd('s')) } else { $null }
        $DurationMinutes = if ($null -ne $DurationSec) { [math]::Round($DurationSec / 60.0, 0) } else { $null }
        $Polyline = if ($SelectedRoute.polyline) { $SelectedRoute.polyline.encodedPolyline } else { $null }
        $Labels = if ($SelectedRoute.routeLabels) { @($SelectedRoute.routeLabels) } else { @() }

        return [PSCustomObject]@{
            OdlegloscKm     = $DistanceKm
            CzasMin         = $DurationMinutes
            DurationSeconds = $DurationSec
            EncodedPolyline = $Polyline
            RouteType       = $RouteType
            RouteLabels     = $Labels
            Status          = 'OK'
            ErrorMessage    = $null
        }
    }
    catch {
        $ErrorMsg = $_.Exception.Message
        return [PSCustomObject]@{
            OdlegloscKm     = $null
            CzasMin         = $null
            DurationSeconds = $null
            EncodedPolyline = $null
            RouteType       = $RouteType
            RouteLabels     = @()
            Status          = "EXCEPTION: $ErrorMsg"
            ErrorMessage    = $ErrorMsg
        }
    }
}

function Get-GoogleMapsUrl {
    param(
        [string]$Origin,
        [string]$Destination,
        [object[]]$Waypoints = @(),
        [string]$TravelMode = 'driving'
    )
    $OriginEnc = [System.Uri]::EscapeDataString($Origin.Trim())
    $DestEnc = [System.Uri]::EscapeDataString($Destination.Trim())
    $Url = "https://www.google.com/maps/dir/?api=1&origin=$OriginEnc&destination=$DestEnc&travelmode=$TravelMode"

    if ($null -ne $Waypoints -and @($Waypoints).Count -gt 0) {
        $WpStrings = [System.Collections.Generic.List[string]]::new()
        foreach ($wp in $Waypoints) {
            if ($wp -is [string] -and -not [string]::IsNullOrWhiteSpace($wp)) {
                $WpStrings.Add($wp.Trim())
            }
            elseif ($wp.Latitude -and $wp.Longitude) {
                $WpStrings.Add("$($wp.Latitude),$($wp.Longitude)")
            }
            elseif ($wp.ZapytanieAdresowe) {
                $WpStrings.Add([string]$wp.ZapytanieAdresowe)
            }
            elseif ($wp.AdresGeokodowany) {
                $WpStrings.Add([string]$wp.AdresGeokodowany)
            }
        }
        if ($WpStrings.Count -gt 0) {
            $Url += '&waypoints=' + [System.Uri]::EscapeDataString(($WpStrings -join '|'))
        }
    }
    return $Url
}

function Get-WrappedLines {
    param([System.Drawing.Graphics]$G, [string]$Text, [System.Drawing.Font]$F, [float]$MaxW)
    if ([string]::IsNullOrWhiteSpace($Text)) { return [string[]]@('') }
    if ($G.MeasureString($Text, $F).Width -le $MaxW) { return [string[]]@($Text) }
    $Words = $Text -split '\s+'
    $L1 = ''; $L2 = ''; $On2 = $false
    foreach ($W in $Words) {
        if (-not $On2) {
            $T = if ($L1) { "$L1 $W" } else { $W }
            if ($G.MeasureString($T, $F).Width -le $MaxW) { $L1 = $T }
            else { $On2 = $true; $L2 = $W }
        }
        else {
            $T2 = if ($L2) { "$L2 $W" } else { $W }
            if ($G.MeasureString($T2, $F).Width -le $MaxW) { $L2 = $T2 }
            else {
                if ($L2.Length -gt 3) { $L2 = $L2.Substring(0, $L2.Length - 3) + '...' }
                break
            }
        }
    }
    if ($L2) { return [string[]]@($L1, $L2) } else { return [string[]]@($L1) }
}

function Save-RouteMapPng {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$EncodedPolyline,
        [Parameter(Mandatory)][double]$OriginLat,
        [Parameter(Mandatory)][double]$OriginLng,
        [Parameter(Mandatory)][double]$DestLat,
        [Parameter(Mandatory)][double]$DestLng,
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)][string]$ApiKey,
        [Parameter()][int]$Width = 900,
        [Parameter()][int]$Height = 600,
        [Parameter()][object[]]$RoutePoints = @(),
        [Parameter()][Alias('TekstAdresA')][string]$AddressTextA = '',
        [Parameter()][Alias('TekstAdresB')][string]$AddressTextB = '',
        [Parameter()][Alias('TekstOdleglosc')][string]$DistanceText = '',
        [Parameter()][Alias('TekstCzas')][string]$DurationText = '',
        [Parameter()][Alias('TekstNaglowekLewy')][string]$HeaderLeftText = '',
        [Parameter()][Alias('TekstNaglowekPrawy')][string]$HeaderRightText = '',
        [Parameter()][Alias('TekstUmowa')][string]$ContractText = '',
        [Parameter()][Alias('TekstKierunek')][string]$DirectionText = '',
        [Parameter()][Alias('Opis')][string]$Description = '',
        [Parameter()][Alias('DataWygenerowania')][string]$GeneratedDate = '',
        [Parameter()][string]$LanguageCode = 'en',
        [Parameter()][string]$StartRaw = '',
        [Parameter()][string]$StartGeocoded = '',
        [Parameter()][string]$EndRaw = '',
        [Parameter()][string]$EndGeocoded = '',
        [Parameter()][object[]]$WaypointsList = @(),
        [Parameter()][string]$RouteName = '',
        [Parameter()][string]$RouteType = '',
        [Parameter()][object]$OverlayConfig = $null
    )

    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls
    $EncodedForUrl = [System.Uri]::EscapeDataString($EncodedPolyline)
    $MarkerParams = [System.Collections.Generic.List[string]]::new()

    $MarkerStart = [System.Uri]::EscapeDataString("color:green|label:A|$OriginLat,$OriginLng")
    $MarkerParams.Add("&markers=$MarkerStart")

    if ($null -ne $RoutePoints -and @($RoutePoints).Count -gt 0) {
        $IntermediatesOnly = @()
        if ($RoutePoints.Count -gt 2 -and
            [math]::Abs($RoutePoints[0].Latitude - $OriginLat) -lt 0.0001 -and
            [math]::Abs($RoutePoints[-1].Latitude - $DestLat) -lt 0.0001) {
            $IntermediatesOnly = @($RoutePoints[1..($RoutePoints.Count - 2)])
        }
        else {
            $IntermediatesOnly = @($RoutePoints | Where-Object {
                $null -ne $_.Latitude -and $null -ne $_.Longitude -and
                (-not ([math]::Abs($_.Latitude - $OriginLat) -lt 0.0001 -and [math]::Abs($_.Longitude - $OriginLng) -lt 0.0001)) -and
                (-not ([math]::Abs($_.Latitude - $DestLat) -lt 0.0001 -and [math]::Abs($_.Longitude - $DestLng) -lt 0.0001))
            })
        }

        $idx = 1
        foreach ($pt in $IntermediatesOnly) {
            if ($pt.Latitude -and $pt.Longitude) {
                $lbl = if ($idx -le 9) { [string]$idx }
                       elseif ($idx -le 35) { [string][char](55 + $idx) }
                       else { '' }
                $spec = if ($lbl) { "color:blue|label:$lbl|$($pt.Latitude),$($pt.Longitude)" }
                        else { "size:mid|color:blue|$($pt.Latitude),$($pt.Longitude)" }
                $MarkerParams.Add("&markers=" + [System.Uri]::EscapeDataString($spec))
                $idx++
            }
        }
    }

    $MarkerEnd = [System.Uri]::EscapeDataString("color:red|label:B|$DestLat,$DestLng")
    $MarkerParams.Add("&markers=$MarkerEnd")

    $lang = if ($LanguageCode) { ($LanguageCode -split '[-_]')[0].ToLower() } else { 'en' }
    $StaticMapUrl = ("https://maps.googleapis.com/maps/api/staticmap" +
        "?size=${Width}x${Height}" +
        "&language=$lang" +
        "&path=weight:4|color:0x0066FFff|enc:$EncodedForUrl" +
        ($MarkerParams -join '') +
        "&key=$ApiKey")

    try {
        $TargetDir = Split-Path -Parent $OutputPath
        if (-not [string]::IsNullOrWhiteSpace($TargetDir) -and -not (Test-Path $TargetDir)) {
            New-Item -ItemType Directory -Path $TargetDir -Force | Out-Null
        }

        $wc = [System.Net.WebClient]::new()
        try {
            $wc.DownloadFile($StaticMapUrl, $OutputPath)
        }
        finally {
            $wc.Dispose()
        }

        # Resolve overlay configuration
        if ($OverlayConfig -is [string] -and -not [string]::IsNullOrWhiteSpace($OverlayConfig)) {
            try { $OverlayConfig = $OverlayConfig | ConvertFrom-Json } catch { }
        }
        if (-not $OverlayConfig) {
            $OverlayConfig = [PSCustomObject]@{
                EnableTopOverlay    = $true
                EnableBottomOverlay = $true
                Items               = [PSCustomObject]@{
                    StartGeocoded = [PSCustomObject]@{ Enabled = $true;  Panel = 'Bottom'; Align = 'Left';   Order = 1 }
                    EndGeocoded   = [PSCustomObject]@{ Enabled = $true;  Panel = 'Bottom'; Align = 'Left';   Order = 2 }
                    Distance      = [PSCustomObject]@{ Enabled = $true;  Panel = 'Bottom'; Align = 'Left';   Order = 3 }
                    Duration      = [PSCustomObject]@{ Enabled = $true;  Panel = 'Bottom'; Align = 'Center'; Order = 3 }
                    Timestamp     = [PSCustomObject]@{ Enabled = $true;  Panel = 'Bottom'; Align = 'Right';  Order = 3 }
                    RouteName     = [PSCustomObject]@{ Enabled = $true;  Panel = 'Top';    Align = 'Left';   Order = 1 }
                    RouteType     = [PSCustomObject]@{ Enabled = $true;  Panel = 'Top';    Align = 'Right';  Order = 1 }
                    Waypoints     = [PSCustomObject]@{ Enabled = $false; Panel = 'Bottom'; Align = 'Left';   Order = 2 }
                    StartRaw      = [PSCustomObject]@{ Enabled = $false; Panel = 'None';   Align = 'Left';   Order = 1 }
                    EndRaw        = [PSCustomObject]@{ Enabled = $false; Panel = 'None';   Align = 'Left';   Order = 2 }
                }
            }
        }

        $enableTop = if ($null -ne $OverlayConfig.EnableTopOverlay) { [bool]$OverlayConfig.EnableTopOverlay } else { $true }
        $enableBtm = if ($null -ne $OverlayConfig.EnableBottomOverlay) { [bool]$OverlayConfig.EnableBottomOverlay } else { $true }

        # Resolve data values
        $addrStartGeo = if ($StartGeocoded) { $StartGeocoded } elseif ($AddressTextA) { $AddressTextA } else { '' }
        $addrStartRaw = if ($StartRaw) { $StartRaw } else { '' }
        $addrEndGeo   = if ($EndGeocoded) { $EndGeocoded } elseif ($AddressTextB) { $AddressTextB } else { '' }
        $addrEndRaw   = if ($EndRaw) { $EndRaw } else { '' }

        $nameVal = if ($RouteName) { $RouteName } elseif ($HeaderLeftText) { $HeaderLeftText } elseif ($Description) { $Description.Trim() } elseif ($ContractText) { $ContractText } else { '' }

        $typeVal = if ($RouteType) { $RouteType } elseif ($HeaderRightText) { $HeaderRightText } elseif ($DirectionText) { $DirectionText } else { '' }
        if ($typeVal -match '^(?:Type|Typ|Art):\s*(.+)$' -or $typeVal -match '^(Shortest|Fastest|Eco|Najkr[oó]tsza|Najszybsza|Eko|K[uü]rzeste|Schnellste)$') {
            $rawVal = if ($Matches[1]) { $Matches[1].Trim() } else { $Matches[0].Trim() }
            $normVal = if ($rawVal -match '(?i)short|kr[oó]t|k[uü]rz') { 'Shortest' }
                       elseif ($rawVal -match '(?i)eco|eko|fuel') { 'Eco' }
                       elseif ($rawVal -match '(?i)fast|szyb|schnell') { 'Fastest' }
                       else { $rawVal }
            $tPrefix = switch ($lang) { 'de' { 'Typ: ' } 'pl' { 'Typ: ' } default { 'Type: ' } }
            $tName = switch ($lang) {
                'de' { if ($normVal -eq 'Fastest') { 'Schnellste' } elseif ($normVal -eq 'Shortest') { 'Kürzeste' } elseif ($normVal -eq 'Eco') { 'Eco' } else { $normVal } }
                'pl' { if ($normVal -eq 'Fastest') { 'Najszybsza' } elseif ($normVal -eq 'Shortest') { 'Najkrótsza' } elseif ($normVal -eq 'Eco') { 'Eko' } else { $normVal } }
                default { $normVal }
            }
            $typeVal = "$tPrefix$tName"
        }

        $distPrefix = switch ($lang) { 'de' { 'Gesamt: ' } 'pl' { 'Razem: ' } default { 'Total: ' } }
        $distVal = if ($DistanceText) { $DistanceText } else { '' }

        $durVal = if ($DurationText) {
            if ($DurationText -match '^\(.*\)$') { $DurationText } else { "($DurationText)" }
        } else { '' }

        $dateVal = if ($GeneratedDate) { $GeneratedDate } else { (Get-Date -Format 'yyyy-MM-dd  HH:mm') }

        $wpItems = [System.Collections.Generic.List[PSCustomObject]]::new()
        $rawWpList = if ($WaypointsList -and @($WaypointsList).Count -gt 0) {
            $WaypointsList
        } elseif ($RoutePoints -and @($RoutePoints).Count -gt 2) {
            @($RoutePoints[1..($RoutePoints.Count - 2)])
        } else { @() }

        $wIdx = 1
        foreach ($w in $rawWpList) {
            $wText = if ($w -is [string]) { $w }
                     elseif ($w.FormattedAddress) { $w.FormattedAddress }
                     elseif ($w.Address) { $w.Address }
                     else { '' }
            if (-not [string]::IsNullOrWhiteSpace($wText)) {
                $wpItems.Add([PSCustomObject]@{
                    Index = $wIdx
                    Badge = "${wIdx}: "
                    Text  = $wText
                })
                $wIdx++
            }
        }

        # Build active property items map
        $propDataMap = @{
            'StartGeocoded' = @{ Id='StartGeocoded'; Kind='address'; Badge='A: '; BadgeColor='Green'; Text=$addrStartGeo }
            'StartRaw'      = @{ Id='StartRaw';      Kind='address'; Badge='A: '; BadgeColor='Green'; Text=$addrStartRaw }
            'EndGeocoded'   = @{ Id='EndGeocoded';   Kind='address'; Badge='B: '; BadgeColor='Red';   Text=$addrEndGeo }
            'EndRaw'        = @{ Id='EndRaw';        Kind='address'; Badge='B: '; BadgeColor='Red';   Text=$addrEndRaw }
            'Distance'      = @{ Id='Distance';      Kind='stat';    Prefix=$distPrefix; Value=$distVal }
            'Duration'      = @{ Id='Duration';      Kind='stat';    Value=$durVal }
            'Timestamp'     = @{ Id='Timestamp';     Kind='date';    Text=$dateVal }
            'RouteName'     = @{ Id='RouteName';     Kind='title';   Text=$nameVal }
            'RouteType'     = @{ Id='RouteType';     Kind='type';    Text=$typeVal }
            'Waypoints'     = @{ Id='Waypoints';     Kind='waypoints'; Items=$wpItems }
        }

        $topItems = [System.Collections.Generic.List[PSCustomObject]]::new()
        $btmItems = [System.Collections.Generic.List[PSCustomObject]]::new()

        if ($OverlayConfig.Items) {
            $propNames = if ($OverlayConfig.Items -is [System.Collections.IDictionary]) {
                $OverlayConfig.Items.Keys
            } else {
                $OverlayConfig.Items.PSObject.Properties.Name
            }
            foreach ($pName in $propNames) {
                $iCfg = if ($OverlayConfig.Items -is [System.Collections.IDictionary]) {
                    $OverlayConfig.Items[$pName]
                } else {
                    $OverlayConfig.Items.$pName
                }
                if (-not $iCfg) { continue }
                $pEnabled = if ($null -ne $iCfg.Enabled) { [bool]$iCfg.Enabled } else { $true }
                $pPanel   = if ($iCfg.Panel) { [string]$iCfg.Panel } else { 'None' }
                $pAlign   = if ($iCfg.Align) { [string]$iCfg.Align } else { 'Left' }
                $pOrder   = if ($iCfg.Order) { [int]$iCfg.Order } else { 1 }

                if (-not $pEnabled -or $pPanel -eq 'None') { continue }
                if (-not $propDataMap.ContainsKey($pName)) { continue }

                $pData = $propDataMap[$pName]
                $hasContent = $false
                if ($pData.Kind -eq 'waypoints') {
                    $hasContent = ($pData.Items -and $pData.Items.Count -gt 0)
                } elseif ($pData.Kind -eq 'stat') {
                    $hasContent = (-not [string]::IsNullOrWhiteSpace($pData.Value))
                } else {
                    $hasContent = (-not [string]::IsNullOrWhiteSpace($pData.Text))
                }
                if (-not $hasContent) { continue }

                $itemObj = [PSCustomObject]@{
                    Id         = $pName
                    Kind       = $pData.Kind
                    Badge      = $pData.Badge
                    BadgeColor = $pData.BadgeColor
                    Text       = $pData.Text
                    Prefix     = $pData.Prefix
                    Value      = $pData.Value
                    Items      = $pData.Items
                    Panel      = $pPanel
                    Align      = $pAlign
                    Order      = $pOrder
                }

                if ($pPanel -eq 'Top' -and $enableTop) {
                    $topItems.Add($itemObj)
                } elseif ($pPanel -eq 'Bottom' -and $enableBtm) {
                    $btmItems.Add($itemObj)
                }
            }
        }

        $MaTopOverlay = ($enableTop -and $topItems.Count -gt 0)
        $MaBottomOverlay = ($enableBtm -and $btmItems.Count -gt 0)

        if ($MaTopOverlay -or $MaBottomOverlay) {
            try {
                Add-Type -AssemblyName System.Drawing

                $FileBytes = [System.IO.File]::ReadAllBytes($OutputPath)
                $MemStream = [System.IO.MemoryStream]::new($FileBytes)
                $BitmapSrc = [System.Drawing.Bitmap]::new($MemStream)

                $ActualW = $BitmapSrc.Width
                $ActualH = $BitmapSrc.Height

                # Fonts definition
                $FontTopTitle = [System.Drawing.Font]::new('Segoe UI', 10.0, [System.Drawing.FontStyle]::Bold)
                $FontTopType  = [System.Drawing.Font]::new('Segoe UI', 10.0, [System.Drawing.FontStyle]::Bold)
                $FontBadge    = [System.Drawing.Font]::new('Segoe UI', 9.0,  [System.Drawing.FontStyle]::Bold)
                $FontAddr     = [System.Drawing.Font]::new('Segoe UI', 9.5,  [System.Drawing.FontStyle]::Regular)
                $FontDistLbl  = [System.Drawing.Font]::new('Segoe UI', 9.0,  [System.Drawing.FontStyle]::Bold)
                $FontDist     = [System.Drawing.Font]::new('Segoe UI', 12.0, [System.Drawing.FontStyle]::Bold)
                $FontDate     = [System.Drawing.Font]::new('Segoe UI', 8.5,  [System.Drawing.FontStyle]::Regular)

                $PadX  = 14
                $LineH = 20

                # Pre-measurement Graphics
                $dummyBmp = [System.Drawing.Bitmap]::new(1, 1)
                $measGfx  = [System.Drawing.Graphics]::FromImage($dummyBmp)

                # Helper scriptblock to group items by Order
                $BuildRows = {
                    param($items)
                    $orders = @($items | Select-Object -ExpandProperty Order -Unique | Sort-Object)
                    $rows = [System.Collections.Generic.List[PSCustomObject]]::new()
                    foreach ($ord in $orders) {
                        $rowItems = @($items | Where-Object { $_.Order -eq $ord })
                        $left   = [System.Collections.Generic.List[PSCustomObject]]::new()
                        $center = [System.Collections.Generic.List[PSCustomObject]]::new()
                        $right  = [System.Collections.Generic.List[PSCustomObject]]::new()
                        foreach ($it in $rowItems) {
                            if ($it.Align -eq 'Right') { $right.Add($it) }
                            elseif ($it.Align -eq 'Center') { $center.Add($it) }
                            else { $left.Add($it) }
                        }
                        $rows.Add([PSCustomObject]@{
                            Order  = $ord
                            Left   = $left
                            Center = $center
                            Right  = $right
                            Height = 20
                        })
                    }
                    return $rows.ToArray()
                }

                $topRows = @(if ($MaTopOverlay) { & $BuildRows $topItems } else { @() })
                $btmRows = @(if ($MaBottomOverlay) { & $BuildRows $btmItems } else { @() })

                # Measure row heights
                $MeasureRows = {
                    param($rows, $availWidth)
                    foreach ($row in @($rows)) {
                        $maxH = 20
                        $allItems = @($row.Left) + @($row.Center) + @($row.Right)
                        foreach ($it in $allItems) {
                            if ($it.Kind -eq 'address') {
                                $badgeSz = $measGfx.MeasureString($it.Badge, $FontBadge)
                                $addrW = [float]($availWidth - $badgeSz.Width)
                                $lines = @(Get-WrappedLines -G $measGfx -Text $it.Text -F $FontAddr -MaxW $addrW)
                                $it | Add-Member -NotePropertyName 'WrappedLines' -NotePropertyValue $lines -Force
                                $h = [math]::Max(1, $lines.Count) * $LineH
                                if ($h -gt $maxH) { $maxH = $h }
                            }
                            elseif ($it.Kind -eq 'waypoints') {
                                $totalWpH = 0
                                foreach ($wp in $it.Items) {
                                    $bSz = $measGfx.MeasureString($wp.Badge, $FontBadge)
                                    $wpMaxW = [float]($availWidth - $bSz.Width)
                                    $wpLines = @(Get-WrappedLines -G $measGfx -Text $wp.Text -F $FontAddr -MaxW $wpMaxW)
                                    $wp | Add-Member -NotePropertyName 'WrappedLines' -NotePropertyValue $wpLines -Force
                                    $totalWpH += [math]::Max(1, $wpLines.Count) * $LineH
                                }
                                if ($totalWpH -gt $maxH) { $maxH = $totalWpH }
                            }
                            elseif ($it.Kind -eq 'stat') {
                                if ($maxH -lt 24) { $maxH = 24 }
                            }
                            elseif ($it.Kind -in @('title', 'type')) {
                                if ($maxH -lt 22) { $maxH = 22 }
                            }
                        }
                        $row.Height = $maxH
                    }
                }

                $availContentW = [float]($ActualW - ($PadX * 2))
                & $MeasureRows $topRows $availContentW
                & $MeasureRows $btmRows $availContentW

                $measGfx.Dispose()
                $dummyBmp.Dispose()

                # Calculate banner heights
                $TopPad = 8; $TopBotPad = 8; $TopRowSpacing = 4
                $TopBarH = 0
                if ($MaTopOverlay -and @($topRows).Count -gt 0) {
                    $sumTopH = (@($topRows) | Measure-Object -Property Height -Sum).Sum
                    if (-not $sumTopH) { $sumTopH = 20 }
                    $TopBarH = [int]($TopPad + $sumTopH + ((@($topRows).Count - 1) * $TopRowSpacing) + $TopBotPad)
                    if ($TopBarH -lt 38) { $TopBarH = 38 }
                }

                $BtmPadTop = 10; $BtmPadBot = 10; $BtmRowSpacing = 6
                $BtmBarH = 0
                if ($MaBottomOverlay -and @($btmRows).Count -gt 0) {
                    $sumBtmH = (@($btmRows) | Measure-Object -Property Height -Sum).Sum
                    if (-not $sumBtmH) { $sumBtmH = 20 }
                    $BtmBarH = [int]($BtmPadTop + $sumBtmH + ((@($btmRows).Count - 1) * $BtmRowSpacing) + $BtmPadBot)
                }

                $FinalW = $ActualW
                $FinalH = $ActualH + $TopBarH + $BtmBarH

                $Bitmap = [System.Drawing.Bitmap]::new($FinalW, $FinalH, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
                $Graphics = [System.Drawing.Graphics]::FromImage($Bitmap)
                $Graphics.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
                $Graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit

                # 1. Background fill
                $BrushBg = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 15, 23, 42))
                $Graphics.FillRectangle($BrushBg, 0, 0, $FinalW, $FinalH)

                # 2. Draw map image in the middle
                $Graphics.DrawImage($BitmapSrc, 0, $TopBarH, $ActualW, $ActualH)

                # 3. Brushes & Pens
                $PenSep      = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(255, 51, 65, 85), 1.5)
                $BrushWhite  = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 248, 250, 252))
                $BrushYellow = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 250, 204, 21))
                $BrushCyan   = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 56, 189, 248))
                $BrushGreen  = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 16, 185, 129))
                $BrushRed    = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 239, 68, 68))
                $BrushMuted  = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 148, 163, 184))

                # Helper scriptblock to measure width of an item
                $MeasureItemWidth = {
                    param($it)
                    if ($it.Kind -eq 'address') {
                        $bSz = $Graphics.MeasureString($it.Badge, $FontBadge)
                        $tSz = $Graphics.MeasureString($it.Text, $FontAddr)
                        return ($bSz.Width + $tSz.Width)
                    }
                    elseif ($it.Kind -eq 'stat') {
                        $w = 0
                        if ($it.Prefix) { $w += $Graphics.MeasureString($it.Prefix, $FontDistLbl).Width }
                        if ($it.Value)  { $w += $Graphics.MeasureString($it.Value, $FontDist).Width }
                        return $w
                    }
                    elseif ($it.Kind -eq 'title') {
                        return $Graphics.MeasureString($it.Text, $FontTopTitle).Width
                    }
                    elseif ($it.Kind -eq 'type') {
                        return $Graphics.MeasureString($it.Text, $FontTopType).Width
                    }
                    elseif ($it.Kind -eq 'date') {
                        return $Graphics.MeasureString($it.Text, $FontDate).Width
                    }
                    elseif ($it.Kind -eq 'waypoints') {
                        return 200
                    }
                    return 0
                }

                # Helper scriptblock to draw an item at specified coordinates
                $DrawItem = {
                    param($it, [float]$x, [float]$y)
                    if ($it.Kind -eq 'address') {
                        $badgeBrush = if ($it.BadgeColor -eq 'Red') { $BrushRed } else { $BrushGreen }
                        $Graphics.DrawString($it.Badge, $FontBadge, $badgeBrush, $x, $y)
                        $bSz = $Graphics.MeasureString($it.Badge, $FontBadge)
                        $curLineY = $y
                        $lines = if ($it.WrappedLines) { $it.WrappedLines } else { @($it.Text) }
                        foreach ($line in $lines) {
                            $Graphics.DrawString($line, $FontAddr, $BrushWhite, ($x + $bSz.Width), $curLineY)
                            $curLineY += [float]$LineH
                        }
                    }
                    elseif ($it.Kind -eq 'waypoints') {
                        $wpY = $y
                        foreach ($wp in $it.Items) {
                            $Graphics.DrawString($wp.Badge, $FontBadge, $BrushCyan, $x, $wpY)
                            $bSz = $Graphics.MeasureString($wp.Badge, $FontBadge)
                            $lines = if ($wp.WrappedLines) { $wp.WrappedLines } else { @($wp.Text) }
                            foreach ($line in $lines) {
                                $Graphics.DrawString($line, $FontAddr, $BrushWhite, ($x + $bSz.Width), $wpY)
                                $wpY += [float]$LineH
                            }
                        }
                    }
                    elseif ($it.Kind -eq 'stat') {
                        $statX = $x
                        if ($it.Prefix) {
                            $pSz = $Graphics.MeasureString($it.Prefix, $FontDistLbl)
                            $Graphics.DrawString($it.Prefix, $FontDistLbl, $BrushCyan, $statX, ($y + 2))
                            $statX += $pSz.Width
                        }
                        if ($it.Value) {
                            $Graphics.DrawString($it.Value, $FontDist, $BrushYellow, $statX, $y)
                        }
                    }
                    elseif ($it.Kind -eq 'title') {
                        $Graphics.DrawString($it.Text, $FontTopTitle, $BrushWhite, $x, $y)
                    }
                    elseif ($it.Kind -eq 'type') {
                        $Graphics.DrawString($it.Text, $FontTopType, $BrushYellow, $x, $y)
                    }
                    elseif ($it.Kind -eq 'date') {
                        $Graphics.DrawString($it.Text, $FontDate, $BrushMuted, $x, ($y + 3))
                    }
                }

                # Helper scriptblock to render a banner's rows
                $RenderBannerRows = {
                    param($rows, [float]$startY, [float]$spacing)
                    $curY = $startY
                    foreach ($row in $rows) {
                        $leftX = [float]$PadX

                        # 1. Left items
                        foreach ($it in $row.Left) {
                            & $DrawItem $it $leftX $curY
                            $w = & $MeasureItemWidth $it
                            $leftX += [float]($w + 14)
                        }

                        # 2. Right items
                        $totalRightW = 0
                        foreach ($it in $row.Right) {
                            $totalRightW += [float]((& $MeasureItemWidth $it) + 12)
                        }
                        $rightX = [float]($FinalW - $PadX - $totalRightW + 12)
                        foreach ($it in $row.Right) {
                            & $DrawItem $it $rightX $curY
                            $w = & $MeasureItemWidth $it
                            $rightX += [float]($w + 12)
                        }

                        # 3. Center items
                        $totalCenterW = 0
                        foreach ($it in $row.Center) {
                            $totalCenterW += [float]((& $MeasureItemWidth $it) + 12)
                        }
                        $centerX = [float][math]::Max($leftX + 10, ($FinalW - $totalCenterW + 12) / 2)
                        foreach ($it in $row.Center) {
                            & $DrawItem $it $centerX $curY
                            $w = & $MeasureItemWidth $it
                            $centerX += [float]($w + 12)
                        }

                        $curY += [float]($row.Height + $spacing)
                    }
                }

                # 4. Draw Top Header Banner
                if ($MaTopOverlay -and $TopBarH -gt 0 -and @($topRows).Count -gt 0) {
                    $Graphics.DrawLine($PenSep, 0, $TopBarH, $FinalW, $TopBarH)
                    $topStartY = [float]$TopPad
                    if (@($topRows).Count -eq 1) {
                        $topStartY = [float][math]::Max(6, ($TopBarH - $topRows[0].Height) / 2)
                    }
                    & $RenderBannerRows $topRows $topStartY $TopRowSpacing
                }

                # 5. Draw Bottom Footer Banner
                if ($MaBottomOverlay -and $BtmBarH -gt 0 -and @($btmRows).Count -gt 0) {
                    $BtmBarY = $TopBarH + $ActualH
                    $Graphics.DrawLine($PenSep, 0, $BtmBarY, $FinalW, $BtmBarY)
                    $btmStartY = [float]($BtmBarY + $BtmPadTop)
                    & $RenderBannerRows $btmRows $btmStartY $BtmRowSpacing
                }

                # Dispose GDI+ objects
                $PenSep.Dispose()
                $BrushBg.Dispose(); $BrushWhite.Dispose(); $BrushYellow.Dispose()
                $BrushCyan.Dispose(); $BrushGreen.Dispose(); $BrushRed.Dispose(); $BrushMuted.Dispose()
                $FontTopTitle.Dispose(); $FontTopType.Dispose()
                $FontBadge.Dispose(); $FontAddr.Dispose(); $FontDistLbl.Dispose(); $FontDist.Dispose(); $FontDate.Dispose()
                $Graphics.Dispose()

                $Bitmap.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
                $Bitmap.Dispose()
                $BitmapSrc.Dispose()
                $MemStream.Dispose()
            }
            catch { }
        }
        return $true
    }
    catch {
        return $false
    }
}

function Find-MatchingPropertyName {
    param([string[]]$AvailableProperties, [string[]]$Patterns)
    foreach ($pattern in $Patterns) {
        $found = $AvailableProperties | Where-Object { $null -ne $_ -and $_.Trim() -match $pattern } | Select-Object -First 1
        if ($found) { return $found }
    }
    return $null
}

function Import-RouteDataFile {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path, [Parameter()][string]$Delimiter = '')

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Plik wejściowy nie istnieje: $Path"
    }

    $Extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    $RawRows = $null
    $Format = $null

    switch ($Extension) {
        { $_ -in '.xlsx', '.xls' } {
            $Format = 'Excel'
            if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
                throw "Wymagany moduł ImportExcel nie jest zainstalowany. Zainstaluj go poleceniem: Install-Module -Name ImportExcel -Scope CurrentUser"
            }
            Import-Module -Name ImportExcel -ErrorAction Stop
            $RawRows = @(Import-Excel -Path $Path)
        }
        { $_ -in '.csv', '.tsv', '.txt' } {
            $Format = 'CSV'
            $FirstLine = Get-Content -LiteralPath $Path -TotalCount 1
            $UsedDelimiter = if (-not [string]::IsNullOrWhiteSpace($Delimiter)) { $Delimiter }
            elseif ($Extension -eq '.tsv' -or $FirstLine -match "`t") { "`t" }
            elseif ($FirstLine -match ';') { ';' }
            else { ',' }
            $RawRows = @(Import-Csv -LiteralPath $Path -Delimiter $UsedDelimiter)
        }
        '.json' {
            $Format = 'JSON'
            $Content = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
            $ParsedJson = $Content | ConvertFrom-Json
            if ($ParsedJson -is [System.Collections.IEnumerable] -and -not ($ParsedJson -is [string])) {
                $RawRows = @($ParsedJson)
            }
            elseif ($ParsedJson.PSObject.Properties.Name -contains 'Routes') {
                $RawRows = @($ParsedJson.Routes)
            }
            elseif ($ParsedJson.PSObject.Properties.Name -contains 'Stops') {
                $RawRows = @($ParsedJson.Stops)
            }
            else {
                $RawRows = @($ParsedJson)
            }
        }
        default {
            throw "Nieobsługiwany format pliku: $Extension. Obsługiwane rozszerzenia: .xlsx, .xls, .csv, .tsv, .json"
        }
    }

    if ($null -eq $RawRows -or $RawRows.Count -eq 0) {
        return [PSCustomObject]@{
            Mode       = 'Empty'
            Routes     = @()
            RawData    = @()
            FilePath   = $Path
            Format     = $Format
            TotalCount = 0
        }
    }

    $PropNames = @($RawRows[0].PSObject.Properties.Name)

    # Sprawdzenie czy to sekwencja przystanków (SequentialStops)
    $ColSeq = Find-MatchingPropertyName -AvailableProperties $PropNames -Patterns @('^(lp|l\.p\.|kolejnosc|stop|sequence|order|nr)$')
    $ColAddrSeq = Find-MatchingPropertyName -AvailableProperties $PropNames -Patterns @('^(adres|address|lokalizacja|punkt|miejsce)$', 'lokalizacja.*(odbioru|dowozu)', 'adres.*(odbioru|dowozu)')
    $ColCitySeq = Find-MatchingPropertyName -AvailableProperties $PropNames -Patterns @('^(miejscowosc|miasto|city|town)$')

    $IsSequentialStops = ($ColSeq -and ($ColAddrSeq -or $ColCitySeq) -and -not (Find-MatchingPropertyName -AvailableProperties $PropNames -Patterns @('^(start|origin|adres.*a)$')))

    if ($IsSequentialStops) {
        $OrderedStops = @($RawRows | Sort-Object { [int]($_.$ColSeq) })
        $StopList = [System.Collections.Generic.List[PSCustomObject]]::new()
        foreach ($st in $OrderedStops) {
            $addr = if ($ColAddrSeq) { [string]$st.$ColAddrSeq } else { '' }
            $city = if ($ColCitySeq) { [string]$st.$ColCitySeq } else { '' }
            $fullAddr = if ($addr -and $city) { "$addr, $city" } elseif ($addr) { $addr } else { $city }
            $StopList.Add([PSCustomObject]@{
                Sequence = [string]$st.$ColSeq
                Address  = $fullAddr.Trim()
                Raw      = $st
            })
        }

        $RouteObj = $null
        if ($StopList.Count -ge 2) {
            $StartPoint = $StopList[0].Address
            $EndPoint = $StopList[$StopList.Count - 1].Address
            $Waypoints = if ($StopList.Count -gt 2) { @($StopList[1..($StopList.Count - 2)] | ForEach-Object { $_.Address }) } else { @() }
            $RouteObj = [PSCustomObject]@{
                Id          = '1'
                Name        = "Multi-point Route ($($StopList.Count) stops)"
                Start       = $StartPoint
                End         = $EndPoint
                Waypoints   = $Waypoints
                RouteType   = 'Fastest'
                OriginalRow = $OrderedStops
            }
        }

        $RoutesList = [System.Collections.Generic.List[PSCustomObject]]::new()
        if ($RouteObj) { $RoutesList.Add($RouteObj) }

        return [PSCustomObject]@{
            Mode       = 'SequentialStops'
            Stops      = $StopList
            Routes     = $RoutesList
            RawData    = $RawRows
            FilePath   = $Path
            Format     = $Format
            TotalCount = $StopList.Count
        }
    }

    # Tryb RouteList (wiersz = trasa)
    $ColStart = Find-MatchingPropertyName -AvailableProperties $PropNames -Patterns @(
        '(?i)^(start|origin|startpoint|poczat.*|poczatek|od|from|dom)$',
        '(?i)adres.*a|^a$',
        '(?i)punkt.*(poczat|start)'
    )
    $ColEnd = Find-MatchingPropertyName -AvailableProperties $PropNames -Patterns @(
        '(?i)^(end|dest|destination|endpoint|koniec.*|konic.*|cel|meta|do|to|szkola)$',
        '(?i)adres.*b|^b$',
        '(?i)punkt.*(konic|koniec|docel|cel)'
    )
    $ColWaypoints = Find-MatchingPropertyName -AvailableProperties $PropNames -Patterns @(
        '(?i)^(waypoints|waypoint|posredn.*|punkty.*posredn.*|przystank.*|via|stops|praca)$',
        '(?i)posrednie'
    )
    $ColName = Find-MatchingPropertyName -AvailableProperties $PropNames -Patterns @(
        '(?i)^(name|nazwa|umowa|contract|id|nr|opis|description|tytul)$',
        '(?i)numer.*umowy'
    )
    $ColRouteType = Find-MatchingPropertyName -AvailableProperties $PropNames -Patterns @(
        '(?i)^(routetype|typ|typtrasy|tryb|mode|optimization)$'
    )

    $NormalizedRoutes = [System.Collections.Generic.List[PSCustomObject]]::new()
    $idx = 1

    foreach ($row in $RawRows) {
        $startVal = if ($ColStart) { [string]$row.$ColStart } else { '' }
        $endVal   = if ($ColEnd) { [string]$row.$ColEnd } else { '' }
        if ([string]::IsNullOrWhiteSpace($startVal) -or [string]::IsNullOrWhiteSpace($endVal)) { continue }

        $nameVal = if ($ColName) { [string]$row.$ColName } else { "Route $idx" }
        $typeVal = if ($ColRouteType) { [string]$row.$ColRouteType } else { $null }

        if ($typeVal -match '(?i)eco|fuel|paliw|eko') { $typeVal = 'Eco' }
        elseif ($typeVal -match '(?i)short|krot|krót') { $typeVal = 'Shortest' }
        elseif ($typeVal -match '(?i)fast|szyb') { $typeVal = 'Fastest' }
        else { $typeVal = $null }

        $waypointsList = [System.Collections.Generic.List[string]]::new()
        if ($ColWaypoints -and -not [string]::IsNullOrWhiteSpace($row.$ColWaypoints)) {
            $rawWp = $row.$ColWaypoints
            if ($rawWp -is [System.Collections.IEnumerable] -and -not ($rawWp -is [string])) {
                foreach ($item in $rawWp) {
                    if (-not [string]::IsNullOrWhiteSpace([string]$item)) { $waypointsList.Add(([string]$item).Trim()) }
                }
            }
            else {
                $splits = ([string]$rawWp) -split '(?<!\\)[|;]'
                foreach ($s in $splits) {
                    $cleaned = $s.Trim()
                    if (-not [string]::IsNullOrWhiteSpace($cleaned)) { $waypointsList.Add($cleaned) }
                }
            }
        }

        $NormalizedRoutes.Add([PSCustomObject]@{
            Id          = [string]$idx
            Name        = $nameVal
            Start       = $startVal.Trim()
            End         = $endVal.Trim()
            Waypoints   = @($waypointsList)
            RouteType   = $typeVal
            OriginalRow = $row
        })
        $idx++
    }

    return [PSCustomObject]@{
        Mode       = 'RouteList'
        Routes     = $NormalizedRoutes
        RawData    = $RawRows
        FilePath   = $Path
        Format     = $Format
        TotalCount = $NormalizedRoutes.Count
        Columns    = [PSCustomObject]@{
            Start     = $ColStart
            End       = $ColEnd
            Waypoints = $ColWaypoints
            Name      = $ColName
            RouteType = $ColRouteType
        }
    }
}

function Export-RouteResults {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object[]]$Results,
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter()][ValidateSet('Excel', 'CSV', 'JSON')][string]$Format = 'Excel'
    )

    $TargetDir = Split-Path -Parent $OutputPath
    if (-not [string]::IsNullOrWhiteSpace($TargetDir) -and -not (Test-Path $TargetDir)) {
        New-Item -ItemType Directory -Path $TargetDir -Force | Out-Null
    }

    # Extract flat summary rows (excluding nested Points array from main sheet/file)
    $RoutesFlat = [System.Collections.Generic.List[PSCustomObject]]::new()
    $PointsFlat = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($r in $Results) {
        $routeId   = if ($null -ne $r.Id) { [string]$r.Id } else { '' }
        $routeName = if ($r.Name) { [string]$r.Name } elseif ($r.Nazwa) { [string]$r.Nazwa } else { "Route $routeId" }
        $startOrig = if ($r.Start) { [string]$r.Start } else { '' }
        $startGeo  = if ($r.StartGeocoded) { [string]$r.StartGeocoded } elseif ($r.StartGeokodowany) { [string]$r.StartGeokodowany } else { '' }
        $startStat = if ($r.StartStatus) { [string]$r.StartStatus } else { '' }
        $endOrig   = if ($r.End) { [string]$r.End } elseif ($r.Koniec) { [string]$r.Koniec } else { '' }
        $endGeo    = if ($r.EndGeocoded) { [string]$r.EndGeocoded } elseif ($r.KoniecGeokodowany) { [string]$r.KoniecGeokodowany } else { '' }
        $endStat   = if ($r.EndStatus) { [string]$r.EndStatus } else { '' }
        $wpCount   = if ($null -ne $r.WaypointsCount) { [int]$r.WaypointsCount } elseif ($null -ne $r.LiczbaPrzystankow) { [int]$r.LiczbaPrzystankow } else { 0 }
        $rType     = if ($r.RouteType) { [string]$r.RouteType } elseif ($r.TypTrasy) { [string]$r.TypTrasy } else { '' }
        $dist      = if ($null -ne $r.DistanceKm) { $r.DistanceKm } elseif ($null -ne $r.OdlegloscKm) { $r.OdlegloscKm } else { $null }
        $dur       = if ($null -ne $r.DurationMin) { $r.DurationMin } elseif ($null -ne $r.CzasMin) { $r.CzasMin } else { $null }
        $status    = if ($r.Status) { [string]$r.Status } else { '' }
        $map       = if ($r.MapPath) { [string]$r.MapPath } elseif ($r.MapaPath) { [string]$r.MapaPath } else { '' }
        $url       = if ($r.GoogleMapsUrl) { [string]$r.GoogleMapsUrl } else { '' }

        # Build waypoints summary text
        $wpSummaryList = [System.Collections.Generic.List[string]]::new()
        if ($r.Points -and ($r.Points -is [System.Collections.IEnumerable])) {
            foreach ($pt in $r.Points) {
                if ($pt.PointType -like 'Waypoint*') {
                    $ptSummary = "$($pt.PointType): '$($pt.OriginalAddress)'"
                    if ($pt.GeocodedAddress) { $ptSummary += " -> '$($pt.GeocodedAddress)'" }
                    if ($pt.GeocodeStatus) { $ptSummary += " [$($pt.GeocodeStatus)]" }
                    $wpSummaryList.Add($ptSummary)
                }

                $PointsFlat.Add([PSCustomObject]@{
                    RouteId         = $routeId
                    RouteName       = $routeName
                    PointOrder      = $pt.Order
                    PointType       = $pt.PointType
                    OriginalAddress = $pt.OriginalAddress
                    GeocodedAddress = $pt.GeocodedAddress
                    GeocodeStatus   = $pt.GeocodeStatus
                    MatchType       = $pt.MatchType
                    IsFallback      = if ($null -ne $pt.IsFallback) { [bool]$pt.IsFallback } else { $false }
                    Latitude        = $pt.Latitude
                    Longitude       = $pt.Longitude
                })
            }
        }

        $wpSummaryText = $wpSummaryList -join ' | '

        $RoutesFlat.Add([PSCustomObject]@{
            Id               = $routeId
            Name             = $routeName
            Start_Original   = $startOrig
            Start_Geocoded   = $startGeo
            Start_Status     = $startStat
            End_Original     = $endOrig
            End_Geocoded     = $endGeo
            End_Status       = $endStat
            WaypointsCount   = $wpCount
            RouteType        = $rType
            DistanceKm       = $dist
            DurationMin      = $dur
            Status           = $status
            WaypointsSummary = $wpSummaryText
            MapPath          = $map
            GoogleMapsUrl    = $url
        })
    }

    $csvEncoding = if ($PSVersionTable.PSVersion.Major -ge 7) { 'utf8BOM' } else { 'UTF8' }

    switch ($Format) {
        'Excel' {
            if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
                Write-Warning "Moduł ImportExcel nie jest zainstalowany. Eksportowanie do CSV zamiast Excel."
                $CsvPath = [System.IO.Path]::ChangeExtension($OutputPath, '.csv')
                $RoutesFlat | Export-Csv -LiteralPath $CsvPath -NoTypeInformation -Encoding $csvEncoding -Delimiter ';'
                if ($PointsFlat.Count -gt 0) {
                    $PtsCsv = [System.IO.Path]::Combine($TargetDir, "$([System.IO.Path]::GetFileNameWithoutExtension($CsvPath))_punkty.csv")
                    $PointsFlat | Export-Csv -LiteralPath $PtsCsv -NoTypeInformation -Encoding $csvEncoding -Delimiter ';'
                }
                return $CsvPath
            }
            Import-Module -Name ImportExcel -ErrorAction Stop
            if (Test-Path -LiteralPath $OutputPath) {
                Remove-Item -LiteralPath $OutputPath -Force -ErrorAction SilentlyContinue
            }
            $RoutesFlat | Export-Excel -Path $OutputPath -WorksheetName 'Trasy' -TableName 'WynikiTras' -AutoSize -AutoFilter -FreezeTopRow
            if ($PointsFlat.Count -gt 0) {
                $PointsFlat | Export-Excel -Path $OutputPath -WorksheetName 'PunktyTrasy' -TableName 'PunktyTrasy' -AutoSize -AutoFilter -FreezeTopRow
            }
            return $OutputPath
        }
        'CSV' {
            $RoutesFlat | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding $csvEncoding -Delimiter ';'
            if ($PointsFlat.Count -gt 0) {
                $PtsCsv = [System.IO.Path]::Combine($TargetDir, "$([System.IO.Path]::GetFileNameWithoutExtension($OutputPath))_punkty.csv")
                $PointsFlat | Export-Csv -LiteralPath $PtsCsv -NoTypeInformation -Encoding $csvEncoding -Delimiter ';'
            }
            return $OutputPath
        }
        'JSON' {
            $JsonContent = $Results | ConvertTo-Json -Depth 5
            [System.IO.File]::WriteAllText($OutputPath, $JsonContent, [System.Text.Encoding]::UTF8)
            return $OutputPath
        }
    }
}

# ══════════════════════════════════════════════════════════════════════════════
# 5. KONFIGURACJA I DPAPI SECURITY
# ══════════════════════════════════════════════════════════════════════════════

$script:AppDirName = 'GoogleMapsRoutes'
$script:LocalConfigFolder = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) $script:AppDirName
if (-not (Test-Path $script:LocalConfigFolder)) {
    New-Item -ItemType Directory -Path $script:LocalConfigFolder -Force | Out-Null
}
$script:ConfigFile = Join-Path $script:LocalConfigFolder 'config.json'
$script:LogFile    = Join-Path $script:LocalConfigFolder 'GoogleMapsRoutes.log'

# External Localization File resolution:
# Priority 1: $PSScriptRoot\localization.json
# Priority 2: EXE directory\localization.json (for compiled PS2EXE binaries)
# Priority 3: %LOCALAPPDATA%\GoogleMapsRoutes\localization.json
$script:ExeDir = try {
    Split-Path ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) -Parent
} catch { $null }

$script:LocalizationFile = if ($PSScriptRoot -and (Test-Path (Join-Path $PSScriptRoot 'localization.json'))) {
    Join-Path $PSScriptRoot 'localization.json'
} elseif ($script:ExeDir -and (Test-Path (Join-Path $script:ExeDir 'localization.json'))) {
    Join-Path $script:ExeDir 'localization.json'
} elseif (Test-Path (Join-Path $script:LocalConfigFolder 'localization.json')) {
    Join-Path $script:LocalConfigFolder 'localization.json'
} elseif ($PSScriptRoot) {
    Join-Path $PSScriptRoot 'localization.json'
} elseif ($script:ExeDir) {
    Join-Path $script:ExeDir 'localization.json'
} else {
    Join-Path $script:LocalConfigFolder 'localization.json'
}

function Load-LocalizationConfig {
    [CmdletBinding()]
    param()

    # If the file does not exist on disk, create default template
    if (-not (Test-Path $script:LocalizationFile)) {
        try {
            $defaultJson = @'
{
  "DefaultLanguage": "en",
  "Languages": {
    "en": {
      "DisplayName": "English",
      "GoogleCode": "en",
      "Strings": {
        "AppTitle": "Google Maps Route & Map Generator",
        "AppSubtitle": "Multi-point driving routes: Fastest, Shortest, Eco-friendly | Import JSON, CSV, Excel",
        "ApiBadgeChecking": "API: Checking...",
        "ApiBadgeActive": "API: Active",
        "ApiBadgeMissing": "API: Missing Key",
        "ApiBadgeError": "API: Error",
        "BtnQuickSettings": "⚙ API Settings",
        "FooterReady": "Ready.",
        "FooterVersion": "Google Maps Routes v2.0",
        "TabManual": "📍 Manual Route",
        "TabBatch": "📁 Batch File Processing",
        "TabSettings": "⚙ Settings & API Key",
        "ManualHeaderRoutePoints": "Route Points",
        "ManualOrigin": "Origin (Start / A):",
        "ManualWaypoints": "Intermediate Stops (optional up to 25):",
        "ManualWaypointsTooltip": "Enter waypoint address and click Add",
        "ManualBtnAdd": "➕ Add",
        "ManualBtnUp": "▲ Up",
        "ManualBtnDown": "▼ Down",
        "ManualBtnRemove": "✕ Remove",
        "ManualBtnClear": "🗑 Clear",
        "ManualDestination": "Destination (End / B):",
        "ManualRouteName": "Route Name / Description:",
        "ManualHeaderOptimization": "Route Optimization",
        "ManualOptFastest": "⚡ Fastest",
        "ManualOptShortest": "📏 Shortest",
        "ManualOptEco": "🌿 Eco",
        "ManualEmission": "Vehicle Engine Type (for Eco route):",
        "ManualFuelGasoline": "Gasoline",
        "ManualFuelDiesel": "Diesel",
        "ManualFuelHybrid": "Hybrid",
        "ManualFuelElectric": "Electric",
        "ManualTrafficAware": "Real-time traffic awareness (Live Traffic)",
        "ManualBtnCalculate": "🚀 CALCULATE ROUTE & DOWNLOAD MAP",
        "ManualBtnCalculating": "⏳ CALCULATING ROUTE...",
        "ManualStatDistance": "DISTANCE",
        "ManualStatDuration": "DURATION",
        "ManualStatType": "ROUTE TYPE",
        "ManualStatusIdle": "Idle",
        "ManualStatusCalculating": "Calculating...",
        "ManualStatusSuccess": "Route calculated",
        "ManualStatusError": "Calculation error",
        "ManualMapPlaceholder": "Map preview will appear here after route calculation...",
        "ManualNoUrl": "No generated link",
        "ManualBtnGoogleMaps": "🌐 Google Maps",
        "ManualBtnCopyUrl": "📋 Copy Link",
        "ManualBtnSaveMapAs": "💾 Save Map As...",
        "BatchInputFile": "Input File (JSON/CSV/XLSX):",
        "BatchBtnBrowse": "📂 Browse File...",
        "BatchBtnReload": "🔄 Reload",
        "BatchNoFileLoaded": "No file loaded.",
        "BatchDefaultRouteType": "Default route type:",
        "BatchOptFromSource": "From Source / Default",
        "BatchBtnStart": "▶ Start Processing",
        "BatchBtnStop": "⏹ Stop",
        "BatchTabInputPreview": "📋 Input Data Preview",
        "BatchTabResults": "📊 Calculation Results",
        "BatchTabLog": "📝 Activity Log",
        "BatchColId": "ID",
        "BatchColName": "Route Name",
        "BatchColOrigin": "Origin (Start)",
        "BatchColDestination": "Destination (End)",
        "BatchColWaypoints": "Waypoints",
        "BatchColType": "Type",
        "BatchColDistance": "Distance (km)",
        "BatchColDuration": "Duration (min)",
        "BatchColStatus": "Status",
        "BatchColMap": "PNG Map",
        "BatchProgressReady": "Ready",
        "BatchBtnOpenOutputDir": "📂 Open Output Folder",
        "BatchBtnExportExcel": "📊 Export Excel",
        "BatchBtnExportCsv": "📄 CSV",
        "BatchBtnExportJson": "📋 JSON",
        "SettingsHeaderApi": "Google Maps API Key",
        "SettingsApiDesc": "Required for Geocoding API, Routes API v2, and Static Maps API.",
        "SettingsApiLabel": "API Key:",
        "SettingsBtnShow": "👁 Show",
        "SettingsBtnHide": "🔒 Hide",
        "SettingsBtnTestKey": "🔍 Test Key",
        "SettingsChkRemember": "Remember securely on this computer (DPAPI CurrentUser encryption)",
        "SettingsHeaderPreferences": "Default Generation Preferences",
        "SettingsDefaultRouteType": "Default route type:",
        "SettingsDefaultEmission": "Default engine type for Eco routes:",
        "SettingsDefaultMapSize": "Default dimensions for generated PNG map:",
        "SettingsOutputDir": "Results Output Folder:",
        "SettingsBtnBrowseOutputDir": "📂 Browse...",
        "SettingsHeaderLanguage": "Language & Localization",
        "SettingsLanguageLabel": "Application & Google Maps API Language:",
        "SettingsBtnOpenLangFile": "📂 Open Localization File (localization.json)",
        "SettingsBtnReloadLang": "🔄 Reload Languages",
        "SettingsBtnSave": "💾 SAVE SETTINGS",
        "SettingsBtnOpenLog": "📋 OPEN LOG FILE",
        "MapLabelTotal": "Total: ",
        "MapLabelType": "Type: ",
        "MapLabelContract": "Contract: ",
        "MapLabelDirection": "Direction: ",
        "MsgMissingApiKey": "Please enter an API key before testing.",
        "MsgMissingApiKeyTitle": "Missing API Key",
        "MsgMissingApiKeyPrompt": "Please enter and save a Google Maps API key in Settings.",
        "MsgMissingData": "Please enter both an origin (start) and a destination (end).",
        "MsgMissingDataTitle": "Missing Data",
        "MsgNoDataFile": "Please load a valid data file first (JSON, CSV, or Excel).",
        "MsgNoDataFileTitle": "No Data",
        "MsgMaxWaypoints": "Maximum number of waypoints is 25.",
        "MsgMaxWaypointsTitle": "Waypoint Limit",
        "MsgSettingsSaved": "Settings have been saved successfully.",
        "MsgSettingsSavedTitle": "Saved",
        "MsgUrlCopied": "Google Maps navigation link copied to clipboard.",
        "MsgUrlCopiedTitle": "Copied",
        "MsgMapSaved": "Map saved: {0}",
        "MsgMapSavedTitle": "Saved",
        "MsgNoExportResults": "No results to export.",
        "MsgNoExportResultsTitle": "Empty Results",
        "MsgExportExcelComplete": "Exported to Excel:
{0}",
        "MsgExportCsvComplete": "Exported to CSV:
{0}",
        "MsgExportJsonComplete": "Exported to JSON:
{0}",
        "MsgExportTitle": "Export Complete",
        "MsgLangReloaded": "Language definitions reloaded successfully ({0} languages found).",
        "MsgLangReloadedTitle": "Languages Reloaded",
        "ThemeToggle": "Theme:",
        "ThemeDark": "🌙 Dark",
        "ThemeLight": "☀️ Light",
        "ThemeToggleTip": "Toggle Light / Dark theme",
        "SettingsThemeLabel": "Application Theme (Color Scheme):",
        "BatchTabPoints": "📍 Points Detail",
        "PointsColRouteId": "Route ID",
        "PointsColRouteName": "Route Name",
        "PointsColOrder": "No.",
        "PointsColType": "Point Type",
        "PointsColOriginalAddress": "Original Address",
        "PointsColGeocodedAddress": "Geocoded Address",
        "PointsColGeocodeStatus": "Geocode Status",
        "PointsColMatchType": "Match Type",
        "PointsColIsFallback": "Fallback?",
        "PointsColLatitude": "Latitude",
        "PointsColLongitude": "Longitude"
      }
    }
  }
}
'@
            $dir = Split-Path -Parent $script:LocalizationFile
            if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
            [System.IO.File]::WriteAllText($script:LocalizationFile, $defaultJson, [System.Text.UTF8Encoding]::new($true))
        } catch { }
    }

    # Auto-sync: If master localization.json exists in ExeDir or PSScriptRoot and is newer than active file, copy it
    $masterLoc = if ($script:ExeDir -and (Test-Path (Join-Path $script:ExeDir 'localization.json'))) {
        Join-Path $script:ExeDir 'localization.json'
    } elseif ($PSScriptRoot -and (Test-Path (Join-Path $PSScriptRoot 'localization.json'))) {
        Join-Path $PSScriptRoot 'localization.json'
    } else { $null }

    if ($masterLoc -and $script:LocalizationFile -ne $masterLoc -and (Test-Path $masterLoc)) {
        try {
            $masterTime = (Get-Item $masterLoc).LastWriteTimeUtc
            $activeTime = if (Test-Path $script:LocalizationFile) { (Get-Item $script:LocalizationFile).LastWriteTimeUtc } else { [DateTime]::MinValue }
            if ($masterTime -gt $activeTime) {
                Copy-Item -Path $masterLoc -Destination $script:LocalizationFile -Force
            }
        } catch { }
    }

    $script:LanguagesCatalog = [ordered]@{}
    $script:DefaultStrings = @{}

    if (Test-Path $script:LocalizationFile) {
        try {
            $raw = [System.IO.File]::ReadAllText($script:LocalizationFile, [System.Text.Encoding]::UTF8)
            $parsed = $raw | ConvertFrom-Json
            if ($parsed.Languages) {
                foreach ($prop in $parsed.Languages.PSObject.Properties) {
                    $code = $prop.Name.ToLower()
                    $langData = $prop.Value
                    $disp = if ($langData.DisplayName) { [string]$langData.DisplayName } else { $code.ToUpper() }
                    $gCode = if ($langData.GoogleCode) { [string]$langData.GoogleCode } else { $code }

                    $strMap = @{}
                    if ($langData.Strings) {
                        foreach ($sProp in $langData.Strings.PSObject.Properties) {
                            $strMap[$sProp.Name] = [string]$sProp.Value
                        }
                    }

                    $script:LanguagesCatalog[$code] = [PSCustomObject]@{
                        Code        = $code
                        DisplayName = $disp
                        GoogleCode  = $gCode
                        Strings     = $strMap
                    }
                }
            }
        } catch {
            Write-AppLog "Error parsing localization file $script:LocalizationFile : $($_.Exception.Message)" "WARN"
        }
    }

    # Ensure fallback EN exists
    if (-not $script:LanguagesCatalog.Contains('en')) {
        $script:LanguagesCatalog['en'] = [PSCustomObject]@{
            Code        = 'en'
            DisplayName = 'English'
            GoogleCode  = 'en'
            Strings     = @{ 'AppTitle' = 'Google Maps Route & Map Generator' }
        }
    }
    $script:DefaultStrings = $script:LanguagesCatalog['en'].Strings
}

function Get-LocText {
    param(
        [Parameter(Mandatory)][string]$Key,
        [Parameter()][string]$Default = $null
    )
    if ($script:CurrentStrings -and $script:CurrentStrings.ContainsKey($Key) -and -not [string]::IsNullOrWhiteSpace($script:CurrentStrings[$Key])) {
        return $script:CurrentStrings[$Key]
    }
    if ($script:DefaultStrings -and $script:DefaultStrings.ContainsKey($Key) -and -not [string]::IsNullOrWhiteSpace($script:DefaultStrings[$Key])) {
        return $script:DefaultStrings[$Key]
    }
    if ($Default) { return $Default }
    return $Key
}

function Get-MaskedKey([string]$Key) {
    if ([string]::IsNullOrWhiteSpace($Key)) { return '(brak)' }
    if ($Key.Length -le 8) { return '***' }
    return "$($Key.Substring(0, 4))...$($Key.Substring($Key.Length - 4, 4))"
}

function Write-AppLog {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'OK', 'WARN', 'ERROR', 'DEBUG')]
        [string]$Level = 'INFO',
        [switch]$ToBatchWindow
    )
    $now = Get-Date
    $timeStr = $now.ToString('yyyy-MM-dd HH:mm:ss.fff')
    $prefix = switch ($Level) {
        'OK'    { '[OK]   ' }
        'WARN'  { '[WARN] ' }
        'ERROR' { '[ERROR]' }
        'DEBUG' { '[DEBUG]' }
        default { '[INFO] ' }
    }
    $entry = "[$timeStr] $prefix $Message"

    try {
        [System.IO.File]::AppendAllText($script:LogFile, "$entry`r`n", [System.Text.UTF8Encoding]::new($true))
    } catch { }

    if ($ToBatchWindow -and $txtBatchLog) {
        try {
            $batchTime = $now.ToString('HH:mm:ss')
            $line = "$batchTime $prefix $Message`r`n"
            $txtBatchLog.Dispatcher.Invoke([Action]{
                $txtBatchLog.AppendText($line)
                $txtBatchLog.ScrollToEnd()
            })
        } catch { }
    }
}

Write-AppLog "================================================================================" "INFO"
Write-AppLog "Uruchomienie Google Maps Route & Map Generator v2.0" "INFO"
Write-AppLog "Środowisko: PowerShell $($PSVersionTable.PSVersion), OS: $([System.Environment]::OSVersion.VersionString)" "INFO"
Write-AppLog "Plik konfiguracji: $script:ConfigFile" "INFO"
Write-AppLog "Plik dziennika zdarzeń (log): $script:LogFile" "INFO"

$script:OverlayPropKeys = @('StartGeocoded', 'EndGeocoded', 'Distance', 'Duration', 'Timestamp', 'RouteName', 'RouteType', 'Waypoints', 'StartRaw', 'EndRaw')

function Get-DefaultOverlayConfig {
    return [ordered]@{
        EnableTopOverlay    = $true
        EnableBottomOverlay = $true
        Items               = [ordered]@{
            StartGeocoded = [ordered]@{ Enabled = $true;  Panel = 'Bottom'; Align = 'Left';   Order = 1 }
            EndGeocoded   = [ordered]@{ Enabled = $true;  Panel = 'Bottom'; Align = 'Left';   Order = 2 }
            Distance      = [ordered]@{ Enabled = $true;  Panel = 'Bottom'; Align = 'Left';   Order = 3 }
            Duration      = [ordered]@{ Enabled = $true;  Panel = 'Bottom'; Align = 'Center'; Order = 3 }
            Timestamp     = [ordered]@{ Enabled = $true;  Panel = 'Bottom'; Align = 'Right';  Order = 3 }
            RouteName     = [ordered]@{ Enabled = $true;  Panel = 'Top';    Align = 'Left';   Order = 1 }
            RouteType     = [ordered]@{ Enabled = $true;  Panel = 'Top';    Align = 'Right';  Order = 1 }
            Waypoints     = [ordered]@{ Enabled = $false; Panel = 'Bottom'; Align = 'Left';   Order = 2 }
            StartRaw      = [ordered]@{ Enabled = $false; Panel = 'None';   Align = 'Left';   Order = 1 }
            EndRaw        = [ordered]@{ Enabled = $false; Panel = 'None';   Align = 'Left';   Order = 2 }
        }
    }
}

function Load-AppConfig {
    $defaultResults = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'GoogleMapsRoutes\Results'
    $defaultOverlay = Get-DefaultOverlayConfig

    $cfg = [PSCustomObject]@{
        ApiKey           = ''
        RememberApiKey   = $true
        LastOutputFolder = $defaultResults
        LastInputFolder  = ''
        LastInputPath    = ''
        DefaultRouteType = 'Fastest'
        DefaultEmission  = 'GASOLINE'
        MapWidth         = 900
        MapHeight        = 600
        Language         = 'en'
        OverlayConfig    = $defaultOverlay
        Theme            = 'Dark'
    }

    if (Test-Path $script:ConfigFile) {
        try {
            $jsonText = [System.IO.File]::ReadAllText($script:ConfigFile, [System.Text.Encoding]::UTF8)
            $raw = $jsonText | ConvertFrom-Json
            if ($raw.ApiKeyEncrypted -is [string] -and -not [string]::IsNullOrWhiteSpace($raw.ApiKeyEncrypted)) {
                $dec = Unprotect-SecretString -EncryptedText $raw.ApiKeyEncrypted
                if (-not [string]::IsNullOrWhiteSpace($dec)) { $cfg.ApiKey = $dec }
            }
            elseif ($raw.ApiKey -is [string] -and -not [string]::IsNullOrWhiteSpace($raw.ApiKey)) {
                $dec = Unprotect-SecretString -EncryptedText $raw.ApiKey
                if (-not [string]::IsNullOrWhiteSpace($dec)) { $cfg.ApiKey = $dec }
            }

            if ($null -ne $raw.RememberApiKey) { $cfg.RememberApiKey = [bool]$raw.RememberApiKey }
            if ($raw.LastOutputFolder -is [string] -and -not [string]::IsNullOrWhiteSpace($raw.LastOutputFolder)) {
                $cfg.LastOutputFolder = $raw.LastOutputFolder
            }
            if ($raw.LastInputFolder -is [string] -and -not [string]::IsNullOrWhiteSpace($raw.LastInputFolder)) {
                $cfg.LastInputFolder = $raw.LastInputFolder
            }
            if ($raw.LastInputPath -is [string] -and -not [string]::IsNullOrWhiteSpace($raw.LastInputPath)) {
                $cfg.LastInputPath = $raw.LastInputPath
            }
            if ($raw.DefaultRouteType -is [string]) { $cfg.DefaultRouteType = $raw.DefaultRouteType }
            if ($raw.DefaultEmission -is [string]) { $cfg.DefaultEmission = $raw.DefaultEmission }
            if ($raw.MapWidth) { $cfg.MapWidth = [int]$raw.MapWidth }
            if ($raw.MapHeight) { $cfg.MapHeight = [int]$raw.MapHeight }
            if ($raw.Language -is [string] -and -not [string]::IsNullOrWhiteSpace($raw.Language)) {
                $cfg.Language = $raw.Language.Trim().ToLower()
            }
            if ($raw.Theme -is [string] -and $raw.Theme -in @('Dark', 'Light')) {
                $cfg.Theme = $raw.Theme
            }

            if ($raw.OverlayConfig) {
                if ($null -ne $raw.OverlayConfig.EnableTopOverlay) {
                    $defaultOverlay.EnableTopOverlay = [bool]$raw.OverlayConfig.EnableTopOverlay
                }
                if ($null -ne $raw.OverlayConfig.EnableBottomOverlay) {
                    $defaultOverlay.EnableBottomOverlay = [bool]$raw.OverlayConfig.EnableBottomOverlay
                }
                if ($raw.OverlayConfig.Items) {
                    foreach ($k in $script:OverlayPropKeys) {
                        $rawItem = if ($raw.OverlayConfig.Items.PSObject.Properties[$k]) {
                            $raw.OverlayConfig.Items.$k
                        } elseif ($raw.OverlayConfig.Items[$k]) {
                            $raw.OverlayConfig.Items[$k]
                        } else { $null }

                        if ($rawItem) {
                            $en = if ($null -ne $rawItem.Enabled) { [bool]$rawItem.Enabled } else { $defaultOverlay.Items[$k].Enabled }
                            $pn = if ($rawItem.Panel) { [string]$rawItem.Panel } else { $defaultOverlay.Items[$k].Panel }
                            $al = if ($rawItem.Align) { [string]$rawItem.Align } else { $defaultOverlay.Items[$k].Align }
                            $od = if ($rawItem.Order) { [int]$rawItem.Order } else { $defaultOverlay.Items[$k].Order }
                            $defaultOverlay.Items[$k] = [ordered]@{ Enabled = $en; Panel = $pn; Align = $al; Order = $od }
                        }
                    }
                }
            }
            $cfg.OverlayConfig = $defaultOverlay
        } catch { }
    }

    if ([string]::IsNullOrWhiteSpace($cfg.ApiKey) -and -not [string]::IsNullOrWhiteSpace($env:GOOGLE_MAPS_API_KEY)) {
        $cfg.ApiKey = $env:GOOGLE_MAPS_API_KEY
    }

    return $cfg
}

function Save-AppConfig {
    param(
        [string]$ApiKey,
        [bool]$RememberApiKey,
        [string]$OutputFolder,
        [string]$LastInputFolder = '',
        [string]$LastInputPath = '',
        [string]$DefaultRouteType = 'Fastest',
        [string]$DefaultEmission = 'GASOLINE',
        [int]$MapWidth = 900,
        [int]$MapHeight = 600,
        [string]$Language = '',
        [object]$OverlayConfig = $null,
        [string]$Theme = ''
    )
    $encKey = ''
    if ($RememberApiKey -and -not [string]::IsNullOrWhiteSpace($ApiKey)) {
        try {
            $protected = Protect-SecretString -PlainText $ApiKey
            if ($protected -is [string] -and -not [string]::IsNullOrWhiteSpace($protected)) {
                $encKey = $protected
            }
        } catch {
            $encKey = ''
        }
    }

    $finalInputFolder = if ($LastInputFolder) { $LastInputFolder } elseif ($script:Config -and $script:Config.LastInputFolder) { $script:Config.LastInputFolder } else { '' }
    $finalInputPath   = if ($LastInputPath) { $LastInputPath } elseif ($script:Config -and $script:Config.LastInputPath) { $script:Config.LastInputPath } else { '' }
    $finalLang        = if ($Language) { $Language } elseif ($script:Config -and $script:Config.Language) { $script:Config.Language } else { 'en' }
    $finalOverlay     = if ($OverlayConfig) { $OverlayConfig } elseif ($script:Config -and $script:Config.OverlayConfig) { $script:Config.OverlayConfig } else { Get-DefaultOverlayConfig }
    $finalTheme       = if ($Theme -in @('Dark', 'Light')) { $Theme } elseif ($script:Config -and $script:Config.Theme) { $script:Config.Theme } else { 'Dark' }

    $cfg = [ordered]@{
        ApiKeyEncrypted  = $encKey
        RememberApiKey   = $RememberApiKey
        LastOutputFolder = $OutputFolder
        LastInputFolder  = $finalInputFolder
        LastInputPath    = $finalInputPath
        DefaultRouteType = $DefaultRouteType
        DefaultEmission  = $DefaultEmission
        MapWidth         = $MapWidth
        MapHeight        = $MapHeight
        Language         = $finalLang
        OverlayConfig    = $finalOverlay
        Theme            = $finalTheme
    }
    $json = $cfg | ConvertTo-Json -Depth 6
    [System.IO.File]::WriteAllText($script:ConfigFile, $json, [System.Text.UTF8Encoding]::new($true))
}

function Get-CurrentOverlayConfig {
    $topEn = if ($chkEnableTopOverlay) { [bool]$chkEnableTopOverlay.IsChecked } else { $true }
    $btmEn = if ($chkEnableBottomOverlay) { [bool]$chkEnableBottomOverlay.IsChecked } else { $true }
    $cfg = [ordered]@{
        EnableTopOverlay    = $topEn
        EnableBottomOverlay = $btmEn
        Items               = [ordered]@{}
    }
    foreach ($key in $script:OverlayPropKeys) {
        $chk  = Get-Variable -Name "chkProp_$key"  -ValueOnly -ErrorAction SilentlyContinue
        $cmbP = Get-Variable -Name "cmbPanel_$key" -ValueOnly -ErrorAction SilentlyContinue
        $cmbA = Get-Variable -Name "cmbAlign_$key" -ValueOnly -ErrorAction SilentlyContinue
        $cmbO = Get-Variable -Name "cmbOrder_$key" -ValueOnly -ErrorAction SilentlyContinue

        $enabled = if ($chk) { [bool]$chk.IsChecked } else { $true }
        $panel   = if ($cmbP -and $cmbP.SelectedItem) { [string]$cmbP.SelectedItem.Tag } else { 'Bottom' }
        $align   = if ($cmbA -and $cmbA.SelectedItem) { [string]$cmbA.SelectedItem.Tag } else { 'Left' }
        $order   = if ($cmbO -and $cmbO.SelectedItem) { [int]$cmbO.SelectedItem.Tag } else { 1 }

        $cfg.Items[$key] = [ordered]@{
            Enabled = $enabled
            Panel   = $panel
            Align   = $align
            Order   = $order
        }
    }
    return $cfg
}

function Set-OverlayConfigUi($cfg) {
    if (-not $cfg) { return }
    if ($null -ne $cfg.EnableTopOverlay -and $chkEnableTopOverlay) {
        $chkEnableTopOverlay.IsChecked = [bool]$cfg.EnableTopOverlay
    }
    if ($null -ne $cfg.EnableBottomOverlay -and $chkEnableBottomOverlay) {
        $chkEnableBottomOverlay.IsChecked = [bool]$cfg.EnableBottomOverlay
    }
    if ($cfg.Items) {
        foreach ($key in $script:OverlayPropKeys) {
            $itemCfg = if ($cfg.Items.PSObject.Properties[$key]) {
                $cfg.Items.$key
            } elseif ($cfg.Items[$key]) {
                $cfg.Items[$key]
            } else { $null }

            if (-not $itemCfg) { continue }

            $chk  = Get-Variable -Name "chkProp_$key"  -ValueOnly -ErrorAction SilentlyContinue
            $cmbP = Get-Variable -Name "cmbPanel_$key" -ValueOnly -ErrorAction SilentlyContinue
            $cmbA = Get-Variable -Name "cmbAlign_$key" -ValueOnly -ErrorAction SilentlyContinue
            $cmbO = Get-Variable -Name "cmbOrder_$key" -ValueOnly -ErrorAction SilentlyContinue

            if ($chk -and $null -ne $itemCfg.Enabled) {
                $chk.IsChecked = [bool]$itemCfg.Enabled
            }
            if ($cmbP -and $itemCfg.Panel) {
                foreach ($opt in $cmbP.Items) {
                    if ($opt.Tag -eq $itemCfg.Panel) { $cmbP.SelectedItem = $opt; break }
                }
            }
            if ($cmbA -and $itemCfg.Align) {
                foreach ($opt in $cmbA.Items) {
                    if ($opt.Tag -eq $itemCfg.Align) { $cmbA.SelectedItem = $opt; break }
                }
            }
            if ($cmbO -and $itemCfg.Order) {
                foreach ($opt in $cmbO.Items) {
                    if ([int]$opt.Tag -eq [int]$itemCfg.Order) { $cmbO.SelectedItem = $opt; break }
                }
            }
        }
    }
}

function Reset-OverlayConfigUi {
    $defaultCfg = Get-DefaultOverlayConfig
    Set-OverlayConfigUi $defaultCfg
}

$script:Config = Load-AppConfig
Load-LocalizationConfig
$script:CurrentLanguage = if ($script:Config -and $script:Config.Language -and $script:LanguagesCatalog.Contains($script:Config.Language.ToLower())) {
    $script:Config.Language.ToLower()
} else {
    'en'
}
$script:CurrentGoogleLang = if ($script:LanguagesCatalog.Contains($script:CurrentLanguage)) {
    $script:LanguagesCatalog[$script:CurrentLanguage].GoogleCode
} else {
    'en'
}
$script:CurrentStrings = if ($script:LanguagesCatalog.Contains($script:CurrentLanguage)) {
    $script:LanguagesCatalog[$script:CurrentLanguage].Strings
} else {
    @{}
}
Write-AppLog "Active language: $script:CurrentLanguage (Google API code: $script:CurrentGoogleLang)" "INFO" 

# ══════════════════════════════════════════════════════════════════════════════
# 6. FABRYKA BEZPIECZNYCH WĄTKÓW TŁA (INITIALSESSIONSTATE RUNSPACE)
# ══════════════════════════════════════════════════════════════════════════════

function New-WorkerPowerShell {
    param([scriptblock]$ScriptBlock)
    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    Get-ChildItem function: | Where-Object {
        $_.Name -in @('Protect-SecretString', 'Unprotect-SecretString', 'Test-GoogleApiKey',
                      'Get-AddressComponentValue', 'Get-AddressCoordinates', 'Get-GeocodeStatusDescription', 'Get-CarRouteData',
                      'Get-GoogleMapsUrl', 'Get-WrappedLines', 'Save-RouteMapPng',
                      'Find-MatchingPropertyName', 'Import-RouteDataFile', 'Export-RouteResults')
    } | ForEach-Object {
        try {
            $iss.Commands.Add([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new($_.Name, $_.Definition))
        } catch { }
    }
    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($iss)
    $rs.ApartmentState = [System.Threading.ApartmentState]::MTA
    $rs.ThreadOptions  = [System.Management.Automation.Runspaces.PSThreadOptions]::UseNewThread
    $rs.Open()
    $ps = [PowerShell]::Create()
    $ps.Runspace = $rs
    # WAŻNE: [void] lub $null = zapobiega wyciekowi obiektu PowerShell do pipeline funkcji.
    # Bez tego funkcja zwraca tablicę @($ps, $ps), co przy wywołaniu .BeginInvoke()
    # powoduje próbę ponownego uruchomienia tej samej instancji i błąd:
    # "The operation cannot be performed because a command has already been started."
    $null = $ps.AddScript($ScriptBlock.ToString())
    return $ps
}

# ══════════════════════════════════════════════════════════════════════════════
# ══════════════════════════════════════════════════════════════════════════════
# 6b. MANUAL CALC WORKER SCRIPTBLOCK (Isolated Runspace, Top-level)
# ══════════════════════════════════════════════════════════════════════════════
$script:ManualCalcAsync = {
    param($start, $end, $waypoints, $routeType, $emission, $trafficAware, $name, $apiKey, $outDir, $logFile, $languageCode = 'en', $overlayConfigJson = '')

    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls
    Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

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
        if ($geoStart.Status -ne 'OK') {
            & $wlog "Origin geocoding error: $($geoStart.Status)" "WARN"
            return [PSCustomObject]@{ Success = $false; Error = "Origin geocoding error: $($geoStart.Status)" }
        }
        & $wlog "Origin OK: $($geoStart.FormattedAddress) ($($geoStart.Latitude), $($geoStart.Longitude))" "INFO"

        & $wlog "Geocoding destination: '$end'..." "INFO"
        $geoEnd = Get-AddressCoordinates -Address $end -ApiKey $apiKey -LanguageCode $languageCode
        if ($geoEnd.Status -ne 'OK') {
            & $wlog "Destination geocoding error: $($geoEnd.Status)" "WARN"
            return [PSCustomObject]@{ Success = $false; Error = "Destination geocoding error: $($geoEnd.Status)" }
        }
        & $wlog "Destination OK: $($geoEnd.FormattedAddress) ($($geoEnd.Latitude), $($geoEnd.Longitude))" "INFO"

        $geoWp = [System.Collections.Generic.List[PSCustomObject]]::new()
        foreach ($w in $waypoints) {
            & $wlog "Geocoding waypoint: '$w'..." "INFO"
            $g = Get-AddressCoordinates -Address $w -ApiKey $apiKey -LanguageCode $languageCode
            if ($g.Status -eq 'OK') {
                $geoWp.Add($g)
                & $wlog "Waypoint OK: $($g.FormattedAddress)" "INFO"
            } else {
                & $wlog "Waypoint geocoding error '$w': $($g.Status)" "WARN"
            }
        }

        & $wlog "Querying Google Routes API v2 (Type: $routeType, Engine: $emission)..." "INFO"
        $trasa = Get-CarRouteData -OriginLat $geoStart.Latitude -OriginLng $geoStart.Longitude `
            -DestLat $geoEnd.Latitude -DestLng $geoEnd.Longitude `
            -IntermediatePoints $geoWp -RouteType $routeType -EmissionType $emission `
            -ApiKey $apiKey -LanguageCode $languageCode -TrafficAware:$trafficAware

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

        & $wlog "Map rendering complete. Saved: $saved" "INFO"

        $resolvedMapPath = $(if ($saved) { $mapPath } else { $null })
        return [PSCustomObject]@{
            Success       = $true
            DistanceKm    = $trasa.OdlegloscKm
            DurationMin   = $trasa.CzasMin
            RouteType     = $routeType
            GoogleMapsUrl = $gUrl
            MapPath       = $resolvedMapPath
            Error         = $null
        }
    }
    catch {
        $errFull = $_.Exception.ToString()
        & $wlog "Worker thread exception: $errFull" "ERROR"
        return [PSCustomObject]@{ Success = $false; Error = $_.Exception.Message }
    }
}

# ══════════════════════════════════════════════════════════════════════════════
# 6c. BATCH CALC WORKER SCRIPTBLOCK (Isolated Runspace, Top-level)
# ══════════════════════════════════════════════════════════════════════════════
$script:BatchCalcAsync = {
    param($routes, $apiKey, $outDir, $defaultRouteType, $syncState, $logFile, $languageCode = 'en', $overlayConfigJson = '')

    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls
    Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

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

        $routeName = if ($r.Name) { $r.Name } else { "Route $($i + 1)" }

        & $wlog "Route $($i + 1)/$($total): Processing '$($r.Start)' -> '$($r.End)' (Type: $rType)..." "INFO"

        try {
            $geoStart = Get-AddressCoordinates -Address $r.Start -ApiKey $apiKey -LanguageCode $languageCode
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

            $geoEnd   = Get-AddressCoordinates -Address $r.End -ApiKey $apiKey -LanguageCode $languageCode
            $endStatus = Get-GeocodeStatusDescription -Geo $geoEnd
            $isEndFallback = if ($geoEnd -and ($geoEnd.PartialMatch -or $geoEnd.MatchType -in 'APPROXIMATE', 'GEOMETRIC_CENTER')) { $true } else { $false }

            if ($geoStart.Status -ne 'OK' -or $geoEnd.Status -ne 'OK') {
                $errReason = "Geocoding failed (Start=$($geoStart.Status), End=$($geoEnd.Status))"
                & $wlog "Route $($i + 1)/$($total): $errReason" "WARN"

                $routePointsList.Add([PSCustomObject]@{
                    Order           = 2
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

                $results.Add([PSCustomObject]@{
                    Id                = [string]($i + 1)
                    Name              = $routeName
                    Nazwa             = $routeName
                    Start             = $r.Start
                    StartGeocoded     = if ($geoStart) { $geoStart.FormattedAddress } else { $null }
                    StartStatus       = $startStatus
                    End               = $r.End
                    EndGeocoded       = if ($geoEnd) { $geoEnd.FormattedAddress } else { $null }
                    EndStatus         = $endStatus
                    Koniec            = $r.End
                    WaypointsCount    = 0
                    LiczbaPrzystankow = 0
                    RouteType         = $rType
                    TypTrasy          = $rType
                    DistanceKm        = $null
                    OdlegloscKm       = $null
                    DurationMin       = $null
                    CzasMin           = $null
                    Status            = $errReason
                    MapPath           = $null
                    MapaPath          = $null
                    Points            = @($routePointsList)
                })
                $syncState.FailCount++
                continue
            }

            $geoWp = [System.Collections.Generic.List[PSCustomObject]]::new()
            $wpIdx = 1
            if ($r.Waypoints) {
                foreach ($w in $r.Waypoints) {
                    if ([string]::IsNullOrWhiteSpace($w)) { continue }
                    $g = Get-AddressCoordinates -Address $w -ApiKey $apiKey -LanguageCode $languageCode
                    $wpStatus = Get-GeocodeStatusDescription -Geo $g
                    $isWpFallback = if ($g -and ($g.PartialMatch -or $g.MatchType -in 'APPROXIMATE', 'GEOMETRIC_CENTER')) { $true } else { $false }

                    $routePointsList.Add([PSCustomObject]@{
                        Order           = ($wpIdx + 1)
                        PointType       = "Waypoint $wpIdx"
                        OriginalAddress = $w
                        GeocodedAddress = if ($g) { $g.FormattedAddress } else { $null }
                        GeocodeStatus   = $wpStatus
                        MatchType       = if ($g) { $g.MatchType } else { 'NOT_FOUND' }
                        PartialMatch    = if ($g) { [bool]$g.PartialMatch } else { $false }
                        IsFallback      = $isWpFallback
                        Latitude        = if ($g) { $g.Latitude } else { $null }
                        Longitude       = if ($g) { $g.Longitude } else { $null }
                    })

                    if ($g -and $g.Status -eq 'OK' -and $null -ne $g.Latitude -and $null -ne $g.Longitude) {
                        $geoWp.Add($g)
                    } else {
                        & $wlog "Route $($i + 1)/$($total): Waypoint '$w' cannot be located ($wpStatus). Proceeding without it in driving directions." "WARN"
                    }
                    $wpIdx++
                    Start-Sleep -Milliseconds 60
                }
            }

            # Add End point to structured points
            $routePointsList.Add([PSCustomObject]@{
                Order           = ($routePointsList.Count + 1)
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

            & $wlog "Route $($i + 1)/$($total): Querying Google Routes API..." "INFO"
            $trasa = Get-CarRouteData -OriginLat $geoStart.Latitude -OriginLng $geoStart.Longitude `
                -DestLat $geoEnd.Latitude -DestLng $geoEnd.Longitude `
                -IntermediatePoints $geoWp -RouteType $rType -ApiKey $apiKey `
                -LanguageCode $languageCode

            if ($trasa.Status -ne 'OK') {
                $errReason = "Routes API: $($trasa.Status). $($trasa.ErrorMessage)"
                & $wlog "Route $($i + 1)/$($total): $errReason" "WARN"
                $results.Add([PSCustomObject]@{
                    Id                = [string]($i + 1)
                    Name              = $routeName
                    Nazwa             = $routeName
                    Start             = $r.Start
                    StartGeocoded     = if ($geoStart) { $geoStart.FormattedAddress } else { $null }
                    StartStatus       = $startStatus
                    End               = $r.End
                    EndGeocoded       = if ($geoEnd) { $geoEnd.FormattedAddress } else { $null }
                    EndStatus         = $endStatus
                    Koniec            = $r.End
                    WaypointsCount    = $geoWp.Count
                    LiczbaPrzystankow = $geoWp.Count
                    RouteType         = $rType
                    TypTrasy          = $rType
                    DistanceKm        = $null
                    OdlegloscKm       = $null
                    DurationMin       = $null
                    CzasMin           = $null
                    Status            = $errReason
                    MapPath           = $null
                    MapaPath          = $null
                    Points            = @($routePointsList)
                })
                $syncState.FailCount++
                continue
            }

            $safeName = ($routeName -replace '[\\/:*?"<>|]', '_').Trim()
            $mapPath = Join-Path $outDir "${ts}_route_$($i + 1)_${safeName}.png"

            $allPts = [System.Collections.Generic.List[PSCustomObject]]::new()
            $allPts.Add($geoStart)
            foreach ($wp in $geoWp) { $allPts.Add($wp) }
            $allPts.Add($geoEnd)

            & $wlog "Route $($i + 1)/$($total): Route OK ($($trasa.OdlegloscKm) km, $($trasa.CzasMin) min). Rendering static map..." "INFO"

            $hdrBatchPrefix = switch ($languageCode) { 'de' { 'Typ: ' } 'pl' { 'Typ: ' } default { 'Type: ' } }
            $hdrBatchName = switch ($languageCode) {
                'de' { if ($rType -eq 'Fastest') { 'Schnellste' } elseif ($rType -eq 'Shortest') { 'Kürzeste' } else { 'Eco' } }
                'pl' { if ($rType -eq 'Fastest') { 'Najszybsza' } elseif ($rType -eq 'Shortest') { 'Najkrótsza' } else { 'Eko' } }
                default { $rType }
            }
            $hdrBatchRightText = "$hdrBatchPrefix$hdrBatchName"

            $saved = Save-RouteMapPng -EncodedPolyline $trasa.EncodedPolyline `
                -OriginLat $geoStart.Latitude -OriginLng $geoStart.Longitude `
                -DestLat $geoEnd.Latitude -DestLng $geoEnd.Longitude `
                -RoutePoints $allPts -OutputPath $mapPath -ApiKey $apiKey `
                -AddressTextA $geoStart.FormattedAddress -AddressTextB $geoEnd.FormattedAddress `
                -DistanceText "$($trasa.OdlegloscKm) km" -DurationText "$($trasa.CzasMin) min" `
                -HeaderLeftText $routeName -HeaderRightText $hdrBatchRightText `
                -LanguageCode $languageCode `
                -StartRaw $r.Start -StartGeocoded $geoStart.FormattedAddress `
                -EndRaw $r.End -EndGeocoded $geoEnd.FormattedAddress `
                -WaypointsList $geoWp -RouteName $routeName -RouteType $hdrBatchRightText `
                -OverlayConfig $overlayConfigJson

            $resolvedMapPath = if ($saved -and (Test-Path $mapPath)) { $mapPath } else { $null }

            $results.Add([PSCustomObject]@{
                Id                = [string]($i + 1)
                Name              = $routeName
                Nazwa             = $routeName
                Start             = $r.Start
                StartGeocoded     = if ($geoStart) { $geoStart.FormattedAddress } else { $null }
                StartStatus       = $startStatus
                End               = $r.End
                EndGeocoded       = if ($geoEnd) { $geoEnd.FormattedAddress } else { $null }
                EndStatus         = $endStatus
                Koniec            = $r.End
                WaypointsCount    = $geoWp.Count
                LiczbaPrzystankow = $geoWp.Count
                RouteType         = $rType
                TypTrasy          = $rType
                DistanceKm        = $trasa.OdlegloscKm
                OdlegloscKm       = $trasa.OdlegloscKm
                DurationMin       = $trasa.CzasMin
                CzasMin           = $trasa.CzasMin
                Status            = 'OK'
                MapPath           = $resolvedMapPath
                MapaPath          = $resolvedMapPath
                Points            = @($routePointsList)
            })
            $syncState.SuccessCount++
            & $wlog "Route $($i + 1)/$($total): Complete! Map saved: $(Split-Path $mapPath -Leaf)" "OK"
        }
        catch {
            $errDetail = $_.Exception.Message
            & $wlog "Route $($i + 1)/$($total): Exception: $errDetail" "ERROR"
            $results.Add([PSCustomObject]@{
                Id                = [string]($i + 1)
                Name              = $routeName
                Nazwa             = $routeName
                Start             = $r.Start
                StartGeocoded     = if ($geoStart) { $geoStart.FormattedAddress } else { $null }
                StartStatus       = if ($startStatus) { $startStatus } else { 'EXCEPTION' }
                End               = $r.End
                EndGeocoded       = if ($geoEnd) { $geoEnd.FormattedAddress } else { $null }
                EndStatus         = if ($endStatus) { $endStatus } else { 'EXCEPTION' }
                Koniec            = $r.End
                WaypointsCount    = 0
                LiczbaPrzystankow = 0
                RouteType         = $rType
                TypTrasy          = $rType
                DistanceKm        = $null
                OdlegloscKm       = $null
                DurationMin       = $null
                CzasMin           = $null
                Status            = "Exception: $errDetail"
                MapPath           = $null
                MapaPath          = $null
                Points            = if ($routePointsList) { @($routePointsList) } else { @() }
            })
            $syncState.FailCount++
        }

        Start-Sleep -Milliseconds 100
    }

    return $results
}

# ══════════════════════════════════════════════════════════════════════════════
# 7. DEFINICJA INTERFEJSU WPF XAML (MODERN DARK THEME)
# ══════════════════════════════════════════════════════════════════════════════

[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Google Maps Route &amp; Map Generator"
    Height="880" Width="1100"
    MinHeight="700" MinWidth="900"
    WindowStartupLocation="CenterScreen"
    Background="{DynamicResource BgDark}" Foreground="{DynamicResource TextPrimary}"
    FontFamily="Segoe UI Variable, Segoe UI, sans-serif">

    <Window.Resources>
        <SolidColorBrush x:Key="BgDark" Color="#0F172A"/>
        <SolidColorBrush x:Key="BgCard" Color="#1E293B"/>
        <SolidColorBrush x:Key="BgCardHover" Color="#293548"/>
        <SolidColorBrush x:Key="BgCardAlt" Color="#162032"/>
        <SolidColorBrush x:Key="BorderCard" Color="#334155"/>
        <SolidColorBrush x:Key="TextPrimary" Color="#F8FAFC"/>
        <SolidColorBrush x:Key="TextSecondary" Color="#94A3B8"/>
        <SolidColorBrush x:Key="AccentBlue" Color="#2563EB"/>
        <SolidColorBrush x:Key="AccentGreen" Color="#10B981"/>
        <SolidColorBrush x:Key="AccentAmber" Color="#F59E0B"/>
        <SolidColorBrush x:Key="AccentRed" Color="#EF4444"/>
        <SolidColorBrush x:Key="BgInput" Color="#1E293B"/>
        <SolidColorBrush x:Key="BorderInput" Color="#334155"/>
        <SolidColorBrush x:Key="BtnSecondaryBg" Color="#334155"/>
        <SolidColorBrush x:Key="BtnSecondaryFg" Color="#F8FAFC"/>
        <SolidColorBrush x:Key="GridLines" Color="#2D3748"/>
        <SolidColorBrush x:Key="LogBg" Color="#0A0F1D"/>
        <SolidColorBrush x:Key="LogFg" Color="#38BDF8"/>
        <SolidColorBrush x:Key="DataGridHeaderBg" Color="#0F172A"/>
        <SolidColorBrush x:Key="DataGridHeaderFg" Color="#94A3B8"/>
        <SolidColorBrush x:Key="DataGridRowBg" Color="#1E293B"/>
        <SolidColorBrush x:Key="DataGridAltRowBg" Color="#162032"/>

        <Style TargetType="TextBox">
            <Setter Property="Background" Value="{DynamicResource BgInput}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderInput}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="9,6"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>

        <Style TargetType="PasswordBox">
            <Setter Property="Background" Value="{DynamicResource BgInput}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderInput}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="9,6"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>

        <Style TargetType="Button">
            <Setter Property="Background" Value="#2563EB"/>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="14,8"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="5"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Modern Dynamic ComboBox with Dropdown Popup Template -->
        <ControlTemplate x:Key="ComboBoxToggleButtonTemplate" TargetType="ToggleButton">
            <Border x:Name="TemplateRoot" Background="{TemplateBinding Background}" BorderBrush="{DynamicResource BorderInput}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="5">
                <Border x:Name="SplitBorder" Width="26" HorizontalAlignment="Right" Background="Transparent">
                    <Path x:Name="Arrow" HorizontalAlignment="Center" VerticalAlignment="Center" Fill="{DynamicResource TextSecondary}" Data="M 0 0 L 4 4 L 8 0 Z"/>
                </Border>
            </Border>
            <ControlTemplate.Triggers>
                <Trigger Property="IsMouseOver" Value="true">
                    <Setter TargetName="TemplateRoot" Property="BorderBrush" Value="{DynamicResource AccentBlue}"/>
                    <Setter TargetName="Arrow" Property="Fill" Value="{DynamicResource TextPrimary}"/>
                </Trigger>
                <Trigger Property="IsChecked" Value="true">
                    <Setter TargetName="TemplateRoot" Property="BorderBrush" Value="{DynamicResource AccentBlue}"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="false">
                    <Setter TargetName="TemplateRoot" Property="Opacity" Value="0.5"/>
                </Trigger>
            </ControlTemplate.Triggers>
        </ControlTemplate>

        <Style TargetType="ComboBox">
            <Setter Property="Background" Value="{DynamicResource BgInput}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderInput}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="ScrollViewer.HorizontalScrollBarVisibility" Value="Auto"/>
            <Setter Property="ScrollViewer.VerticalScrollBarVisibility" Value="Auto"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBox">
                        <Grid x:Name="MainGrid" SnapsToDevicePixels="true">
                            <ToggleButton x:Name="ToggleButton"
                                          Template="{StaticResource ComboBoxToggleButtonTemplate}"
                                          Background="{TemplateBinding Background}"
                                          BorderBrush="{TemplateBinding BorderBrush}"
                                          BorderThickness="{TemplateBinding BorderThickness}"
                                          Focusable="false"
                                          IsChecked="{Binding Path=IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"
                                          ClickMode="Press"/>
                            <ContentPresenter x:Name="ContentSite"
                                              IsHitTestVisible="false"
                                              Content="{TemplateBinding SelectionBoxItem}"
                                              ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                                              ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}"
                                              Margin="{TemplateBinding Padding}"
                                              VerticalAlignment="Center"
                                              HorizontalAlignment="Left"/>
                            <Popup x:Name="Popup"
                                   Placement="Bottom"
                                   IsOpen="{TemplateBinding IsDropDownOpen}"
                                   AllowsTransparency="true"
                                   Focusable="false"
                                   PopupAnimation="Slide">
                                <Grid x:Name="DropDown" SnapsToDevicePixels="true" MinWidth="{TemplateBinding ActualWidth}" MaxHeight="{TemplateBinding MaxDropDownHeight}">
                                    <Border x:Name="DropDownBorder" Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="5" Margin="0,2,0,0">
                                        <ScrollViewer Margin="2" SnapsToDevicePixels="true">
                                            <StackPanel IsItemsHost="true" KeyboardNavigation.DirectionalNavigation="Contained"/>
                                        </ScrollViewer>
                                    </Border>
                                </Grid>
                            </Popup>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="HasItems" Value="false">
                                <Setter TargetName="DropDownBorder" Property="MinHeight" Value="40"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="false">
                                <Setter Property="Opacity" Value="0.6"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="ComboBoxItem">
            <Setter Property="Background" Value="{DynamicResource BgCard}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="Padding" Value="10,7"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBoxItem">
                        <Border x:Name="ItemBorder" Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}" SnapsToDevicePixels="true">
                            <ContentPresenter Content="{TemplateBinding Content}" ContentTemplate="{TemplateBinding ContentTemplate}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsHighlighted" Value="true">
                                <Setter TargetName="ItemBorder" Property="Background" Value="{DynamicResource AccentBlue}"/>
                                <Setter Property="Foreground" Value="#FFFFFF"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="true">
                                <Setter TargetName="ItemBorder" Property="Background" Value="{DynamicResource AccentBlue}"/>
                                <Setter Property="Foreground" Value="#FFFFFF"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="false">
                                <Setter Property="Foreground" Value="{DynamicResource TextSecondary}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Modern Dynamic ListBox & Items -->
        <Style TargetType="ListBox">
            <Setter Property="Background" Value="{DynamicResource BgDark}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderCard}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="ScrollViewer.HorizontalScrollBarVisibility" Value="Disabled"/>
        </Style>

        <Style TargetType="ListBoxItem">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="Padding" Value="8,5"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ListBoxItem">
                        <Border x:Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" Padding="{TemplateBinding Padding}" CornerRadius="3" SnapsToDevicePixels="true">
                            <ContentPresenter Content="{TemplateBinding Content}" ContentTemplate="{TemplateBinding ContentTemplate}" HorizontalAlignment="Left" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="true">
                                <Setter TargetName="Bd" Property="Background" Value="{DynamicResource AccentBlue}"/>
                                <Setter Property="Foreground" Value="#FFFFFF"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="true">
                                <Setter TargetName="Bd" Property="Background" Value="{DynamicResource BgCardHover}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Modern Dynamic DataGrid & Elements -->
        <Style TargetType="DataGrid">
            <Setter Property="Background" Value="{DynamicResource BgDark}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderCard}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="RowBackground" Value="{DynamicResource DataGridRowBg}"/>
            <Setter Property="AlternatingRowBackground" Value="{DynamicResource DataGridAltRowBg}"/>
            <Setter Property="GridLinesVisibility" Value="Horizontal"/>
            <Setter Property="HorizontalGridLinesBrush" Value="{DynamicResource GridLines}"/>
            <Setter Property="HeadersVisibility" Value="Column"/>
            <Setter Property="AutoGenerateColumns" Value="False"/>
            <Setter Property="IsReadOnly" Value="True"/>
            <Setter Property="CanUserAddRows" Value="False"/>
            <Setter Property="CanUserDeleteRows" Value="False"/>
            <Setter Property="SelectionMode" Value="Single"/>
            <Setter Property="SelectionUnit" Value="FullRow"/>
        </Style>

        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="{DynamicResource DataGridHeaderBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource DataGridHeaderFg}"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="10,7"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderCard}"/>
            <Setter Property="BorderThickness" Value="0,0,1,1"/>
        </Style>

        <Style TargetType="DataGridRow">
            <Setter Property="Background" Value="{DynamicResource DataGridRowBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="SnapsToDevicePixels" Value="true"/>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="true">
                    <Setter Property="Background" Value="{DynamicResource AccentBlue}"/>
                    <Setter Property="Foreground" Value="#FFFFFF"/>
                </Trigger>
                <Trigger Property="IsMouseOver" Value="true">
                    <Setter Property="Background" Value="{DynamicResource BgCardHover}"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style TargetType="DataGridCell">
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="DataGridCell">
                        <Border Background="{TemplateBinding Background}" BorderThickness="0" Padding="{TemplateBinding Padding}" SnapsToDevicePixels="true">
                            <ContentPresenter SnapsToDevicePixels="{TemplateBinding SnapsToDevicePixels}" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="true">
                    <Setter Property="Background" Value="{DynamicResource AccentBlue}"/>
                    <Setter Property="Foreground" Value="#FFFFFF"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style TargetType="TabItem">
            <Setter Property="Background" Value="{DynamicResource BgCard}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextSecondary}"/>
            <Setter Property="Padding" Value="18,10"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Border Name="TabBorder"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{DynamicResource BorderCard}"
                                BorderThickness="1,1,1,0"
                                CornerRadius="6,6,0,0"
                                Margin="0,0,4,0"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter ContentSource="Header" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="TabBorder" Property="Background" Value="{DynamicResource AccentBlue}"/>
                                <Setter Property="Foreground" Value="#FFFFFF"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <Border Grid.Row="0" Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="14,10" Margin="0,0,0,12">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Orientation="Vertical">
                    <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                        <TextBlock Text="🗺️" FontSize="20" Margin="0,0,8,0" VerticalAlignment="Center"/>
                        <TextBlock Name="txtHeaderTitle" Text="Google Maps Route &amp; Map Generator" FontSize="18" FontWeight="Bold" Foreground="{DynamicResource TextPrimary}"/>
                    </StackPanel>
                    <TextBlock Name="txtHeaderSubtitle" Text="Multi-point driving routes: Fastest, Shortest, Eco-friendly | Import JSON, CSV, Excel" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="28,2,0,0"/>
                </StackPanel>
                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Name="lblApiBadge" Text="API: Checking..." Foreground="#EF4444" FontSize="12" FontWeight="SemiBold" VerticalAlignment="Center" Margin="0,0,10,0"/>
                    <ComboBox Name="cmbAppLanguage" Width="135" Height="30" Margin="0,0,10,0" VerticalAlignment="Center" ToolTip="Select Language / Sprache wählen / Wybierz język"/>
                    <Button Name="btnQuickSettings" Content="⚙ API Settings" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Padding="10,5" FontSize="12" Margin="0,0,10,0"/>
                    <Button Name="btnThemeToggle" Content="🌙 Dark" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" Padding="10,5" FontSize="12" ToolTip="Toggle Light / Dark theme"/>
                </StackPanel>
            </Grid>
        </Border>


        <!-- Main TabControl -->
        <TabControl Name="tabMain" Grid.Row="1" Background="Transparent" BorderThickness="0">

            <!-- TAB 1: MANUAL ROUTE -->
            <TabItem Name="tabItemManual" Header="📍 Manual Route">
                <Grid Margin="0,10,0,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="420" MinWidth="360"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <ScrollViewer VerticalScrollBarVisibility="Auto" Grid.Column="0" Margin="0,0,10,0">
                        <StackPanel>
                            <Border Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="14" Margin="0,0,0,12">
                                <StackPanel>
                                    <TextBlock Name="lblManualRoutePointsHeader" Text="Route Points" FontSize="15" FontWeight="Bold" Foreground="{DynamicResource TextPrimary}" Margin="0,0,0,10"/>

                                    <TextBlock Name="lblManualOrigin" Text="Origin (Start / A):" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,4"/>
                                    <Grid Margin="0,0,0,10">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="Auto"/>
                                        </Grid.ColumnDefinitions>
                                        <TextBox Name="txtManualStart" Text="Warszawa, Plac Defilad 1"/>
                                        <Button Name="btnClearManualStart" Grid.Column="1" Content="✕" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Padding="8,6" Margin="4,0,0,0" ToolTip="Clear"/>
                                    </Grid>

                                    <TextBlock Name="lblManualWaypoints" Text="Intermediate Stops (optional up to 25):" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,4"/>
                                    <Grid Margin="0,0,0,6">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="Auto"/>
                                        </Grid.ColumnDefinitions>
                                        <TextBox Name="txtNewWaypoint" ToolTip="Enter waypoint address and click Add"/>
                                        <Button Name="btnAddWaypoint" Grid.Column="1" Content="➕ Add" Background="#10B981" Margin="4,0,0,0"/>
                                    </Grid>

                                    <ListBox Name="lstWaypoints" Height="110" Margin="0,0,0,6"/>
                                    <Grid Margin="0,0,0,10">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="*"/>
                                        </Grid.ColumnDefinitions>
                                        <Button Name="btnWpUp" Content="▲ Up" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="0,0,2,0" Padding="4,4" FontSize="11"/>
                                        <Button Name="btnWpDown" Grid.Column="1" Content="▼ Down" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="2,0,2,0" Padding="4,4" FontSize="11"/>
                                        <Button Name="btnWpRemove" Grid.Column="2" Content="✕ Remove" Background="#EF4444" Margin="2,0,2,0" Padding="4,4" FontSize="11"/>
                                        <Button Name="btnWpClear" Grid.Column="3" Content="🗑 Clear" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="2,0,0,0" Padding="4,4" FontSize="11"/>
                                    </Grid>

                                    <TextBlock Name="lblManualDestination" Text="Destination (End / B):" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,4"/>
                                    <Grid Margin="0,0,0,6">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="Auto"/>
                                        </Grid.ColumnDefinitions>
                                        <TextBox Name="txtManualEnd" Text="Kraków, Rynek Główny 1"/>
                                        <Button Name="btnClearManualEnd" Grid.Column="1" Content="✕" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Padding="8,6" Margin="4,0,0,0" ToolTip="Clear"/>
                                    </Grid>

                                    <TextBlock Name="lblManualRouteName" Text="Route Name / Description:" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,4,0,4"/>
                                    <TextBox Name="txtManualName" Text="Route Warsaw - Krakow" Margin="0,0,0,6"/>
                                </StackPanel>
                            </Border>

                            <Border Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="14" Margin="0,0,0,12">
                                <StackPanel>
                                    <TextBlock Name="lblManualOptHeader" Text="Route Optimization" FontSize="15" FontWeight="Bold" Foreground="{DynamicResource TextPrimary}" Margin="0,0,0,10"/>

                                    <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                                        <RadioButton Name="rbTypeFastest" Content="⚡ Fastest" IsChecked="True" Foreground="{DynamicResource TextPrimary}" FontSize="13" Margin="0,0,16,0"/>
                                        <RadioButton Name="rbTypeShortest" Content="📏 Shortest" Foreground="{DynamicResource TextPrimary}" FontSize="13" Margin="0,0,16,0"/>
                                        <RadioButton Name="rbTypeEco" Content="🌿 Eco" Foreground="{DynamicResource TextPrimary}" FontSize="13"/>
                                    </StackPanel>

                                    <StackPanel Name="pnlEmission" Orientation="Vertical" Visibility="Collapsed" Margin="0,0,0,8">
                                        <TextBlock Name="lblManualEmission" Text="Vehicle Engine Type (for Eco route):" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,4"/>
                                        <ComboBox Name="cmbEmission">
                                            <ComboBoxItem Content="Gasoline (Benzyna)" Tag="GASOLINE" IsSelected="True"/>
                                            <ComboBoxItem Content="Diesel" Tag="DIESEL"/>
                                            <ComboBoxItem Content="Hybrid" Tag="HYBRID"/>
                                            <ComboBoxItem Content="Electric" Tag="ELECTRIC"/>
                                        </ComboBox>
                                    </StackPanel>

                                    <CheckBox Name="chkTrafficAware" Content="Real-time traffic awareness (Live Traffic)" Foreground="{DynamicResource TextSecondary}" FontSize="12" Margin="0,4,0,4"/>
                                </StackPanel>
                            </Border>

                            <Button Name="btnCalculateManual" Content="🚀 CALCULATE ROUTE &amp; DOWNLOAD MAP" Background="#2563EB" Foreground="#FFFFFF" Padding="16,12" FontSize="14" FontWeight="Bold"/>
                        </StackPanel>
                    </ScrollViewer>

                    <Grid Grid.Column="1" Margin="10,0,0,0">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>

                        <Border Grid.Row="0" Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="14" Margin="0,0,0,10">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>

                                <StackPanel Grid.Column="0">
                                    <TextBlock Name="lblHeaderDist" Text="DISTANCE" FontSize="11" FontWeight="Bold" Foreground="{DynamicResource TextSecondary}"/>
                                    <TextBlock Name="lblManualDist" Text="— km" FontSize="20" FontWeight="Bold" Foreground="#10B981"/>
                                </StackPanel>

                                <StackPanel Grid.Column="1">
                                    <TextBlock Name="lblHeaderDur" Text="DURATION" FontSize="11" FontWeight="Bold" Foreground="{DynamicResource TextSecondary}"/>
                                    <TextBlock Name="lblManualTime" Text="— min" FontSize="20" FontWeight="Bold" Foreground="#F59E0B"/>
                                </StackPanel>

                                <StackPanel Grid.Column="2">
                                    <TextBlock Name="lblHeaderType" Text="ROUTE TYPE" FontSize="11" FontWeight="Bold" Foreground="{DynamicResource TextSecondary}"/>
                                    <TextBlock Name="lblManualType" Text="Fastest" FontSize="16" FontWeight="SemiBold" Foreground="#38BDF8"/>
                                </StackPanel>

                                <StackPanel Grid.Column="3" VerticalAlignment="Center">
                                    <TextBlock Name="lblManualStatus" Text="Idle" FontSize="12" Foreground="{DynamicResource TextSecondary}" HorizontalAlignment="Right"/>
                                </StackPanel>
                            </Grid>
                        </Border>

                        <Border Grid.Row="1" Background="{DynamicResource BgDark}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="6" Margin="0,0,0,10">
                            <Grid>
                                <TextBlock Name="lblMapPlaceholder" Text="Map preview will appear here after route calculation..."
                                           Foreground="{DynamicResource TextSecondary}" FontSize="14" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                <Image Name="imgMapPreview" Stretch="Uniform" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                            </Grid>
                        </Border>

                        <Border Grid.Row="2" Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="10">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Name="lblGoogleUrlDisplay" Text="No generated link" Foreground="{DynamicResource TextSecondary}" FontSize="12" VerticalAlignment="Center" TextTrimming="CharacterEllipsis" Margin="0,0,10,0"/>
                                <Button Name="btnOpenGoogleMaps" Grid.Column="1" Content="🌐 Google Maps" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="0,0,6,0" Padding="10,6" IsEnabled="False"/>
                                <Button Name="btnCopyUrl" Grid.Column="2" Content="📋 Copy Link" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="0,0,6,0" Padding="10,6" IsEnabled="False"/>
                                <Button Name="btnSaveMapAs" Grid.Column="3" Content="💾 Save Map As..." Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Padding="10,6" IsEnabled="False"/>
                            </Grid>
                        </Border>
                    </Grid>
                </Grid>
            </TabItem>

            <!-- TAB 2: BATCH DATA PROCESSING -->
            <TabItem Name="tabItemBatch" Header="📁 Batch File Processing">
                <Grid Margin="0,10,0,0">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <Border Grid.Row="0" Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="14" Margin="0,0,0,10">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>

                            <Grid Grid.Row="0" Margin="0,0,0,10">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Name="lblBatchInputFile" Text="Input File (JSON/CSV/XLSX):" VerticalAlignment="Center" Foreground="{DynamicResource TextSecondary}" Margin="0,0,10,0"/>
                                <TextBox Name="txtBatchFilePath" Grid.Column="1" VerticalAlignment="Center"/>
                                <Button Name="btnBrowseBatchFile" Grid.Column="2" Content="📂 Browse File..." Background="#2563EB" Margin="6,0,0,0"/>
                                <Button Name="btnReloadBatchFile" Grid.Column="3" Content="🔄 Reload" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="6,0,0,0"/>
                            </Grid>

                            <Grid Grid.Row="1">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                                    <TextBlock Name="lblBatchFileInfo" Text="No file loaded." Foreground="{DynamicResource TextSecondary}" FontSize="12"/>
                                </StackPanel>
                                <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center" Margin="10,0">
                                    <TextBlock Name="lblBatchDefaultRouteType" Text="Default route type:" Foreground="{DynamicResource TextSecondary}" FontSize="12" VerticalAlignment="Center" Margin="0,0,6,0"/>
                                    <ComboBox Name="cmbBatchRouteType" Width="170">
                                        <ComboBoxItem Content="From Source / Default" Tag="FromSource" IsSelected="True"/>
                                        <ComboBoxItem Content="Fastest (Najszybsza)" Tag="Fastest"/>
                                        <ComboBoxItem Content="Shortest (Najkrótsza)" Tag="Shortest"/>
                                        <ComboBoxItem Content="Eco (Fuel Efficient)" Tag="Eco"/>
                                    </ComboBox>
                                </StackPanel>
                                <StackPanel Grid.Column="3" Orientation="Horizontal">
                                    <Button Name="btnStartBatch" Content="▶ Start Processing" Background="#10B981" Foreground="#FFFFFF" Padding="14,7" FontWeight="Bold" Margin="0,0,6,0"/>
                                    <Button Name="btnStopBatch" Content="⏹ Stop" Background="#EF4444" Foreground="#FFFFFF" Padding="12,7" IsEnabled="False"/>
                                </StackPanel>
                            </Grid>
                        </Grid>
                    </Border>

                    <TabControl Name="tabBatchSub" Grid.Row="1" Background="Transparent" BorderThickness="0">
                        <TabItem Name="tabSubInput" Header="📋 Input Data Preview">
                            <Border Background="{DynamicResource BgDark}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="6" Margin="0,6,0,0">
                                <DataGrid Name="dgBatchInput"/>
                            </Border>
                        </TabItem>

                        <TabItem Name="tabSubResults" Header="📊 Calculation Results">
                            <Border Background="{DynamicResource BgDark}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="6" Margin="0,6,0,0">
                                <DataGrid Name="dgBatchResults">
                                    <DataGrid.Columns>
                                        <DataGridTextColumn Header="ID" Binding="{Binding Id}" Width="45"/>
                                        <DataGridTextColumn Header="Route Name" Binding="{Binding Name}" Width="170"/>
                                        <DataGridTextColumn Header="Origin (Start)" Binding="{Binding Start}" Width="190"/>
                                        <DataGridTextColumn Header="Destination (End)" Binding="{Binding End}" Width="190"/>
                                        <DataGridTextColumn Header="Waypoints" Binding="{Binding WaypointsCount}" Width="75"/>
                                        <DataGridTextColumn Header="Type" Binding="{Binding RouteType}" Width="75"/>
                                        <DataGridTextColumn Header="Distance (km)" Binding="{Binding DistanceKm}" Width="95"/>
                                        <DataGridTextColumn Header="Duration (min)" Binding="{Binding DurationMin}" Width="85"/>
                                        <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="110"/>
                                        <DataGridTextColumn Header="PNG Map" Binding="{Binding MapPath}" Width="*"/>
                                    </DataGrid.Columns>
                                </DataGrid>
                            </Border>
                        </TabItem>

                        <TabItem Name="tabSubPoints" Header="📍 Points Detail">
                            <Border Background="{DynamicResource BgDark}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="6" Margin="0,6,0,0">
                                <DataGrid Name="dgBatchPoints">
                                    <DataGrid.Columns>
                                        <DataGridTextColumn Header="Route ID" Binding="{Binding RouteId}" Width="65"/>
                                        <DataGridTextColumn Header="Route Name" Binding="{Binding RouteName}" Width="150"/>
                                        <DataGridTextColumn Header="No." Binding="{Binding PointOrder}" Width="45"/>
                                        <DataGridTextColumn Header="Point Type" Binding="{Binding PointType}" Width="90"/>
                                        <DataGridTextColumn Header="Original Address" Binding="{Binding OriginalAddress}" Width="220"/>
                                        <DataGridTextColumn Header="Geocoded Address" Binding="{Binding GeocodedAddress}" Width="240"/>
                                        <DataGridTextColumn Header="Geocode Status" Binding="{Binding GeocodeStatus}" Width="170"/>
                                        <DataGridTextColumn Header="Match Type" Binding="{Binding MatchType}" Width="110"/>
                                        <DataGridTextColumn Header="Fallback?" Binding="{Binding IsFallback}" Width="75"/>
                                        <DataGridTextColumn Header="Latitude" Binding="{Binding Latitude}" Width="85"/>
                                        <DataGridTextColumn Header="Longitude" Binding="{Binding Longitude}" Width="85"/>
                                    </DataGrid.Columns>
                                </DataGrid>
                            </Border>
                        </TabItem>

                        <TabItem Name="tabSubLog" Header="📝 Activity Log">
                            <Border Background="{DynamicResource BgDark}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="6" Margin="0,6,0,0">
                                <TextBox Name="txtBatchLog" IsReadOnly="True" TextWrapping="Wrap"
                                         VerticalScrollBarVisibility="Auto" FontFamily="Consolas, monospace"
                                         FontSize="12" Background="{DynamicResource LogBg}" Foreground="{DynamicResource LogFg}"/>
                            </Border>
                        </TabItem>
                    </TabControl>

                    <Border Grid.Row="2" Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="12" Margin="0,10,0,0">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="230"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>

                            <StackPanel Grid.Column="0" VerticalAlignment="Center">
                                <TextBlock Name="lblBatchProgressText" Text="Ready" FontSize="12" Foreground="{DynamicResource TextSecondary}"/>
                                <ProgressBar Name="pbBatchProgress" Height="14" Minimum="0" Maximum="100" Value="0" Margin="0,4,0,0" Foreground="#10B981" Background="{DynamicResource BgDark}"/>
                            </StackPanel>

                            <TextBlock Name="lblBatchStats" Grid.Column="1" Text="" Foreground="#10B981" FontSize="13" FontWeight="SemiBold" VerticalAlignment="Center" Margin="20,0"/>

                            <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
                                <Button Name="btnOpenOutputDir" Content="📂 Open Output Folder" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="0,0,6,0" Padding="10,6"/>
                                <Button Name="btnExportExcel" Content="📊 Export Excel" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="0,0,6,0" Padding="10,6"/>
                                <Button Name="btnExportCsv" Content="📄 CSV" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="0,0,6,0" Padding="10,6"/>
                                <Button Name="btnExportJson" Content="📋 JSON" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Padding="10,6"/>
                            </StackPanel>
                        </Grid>
                    </Border>
                </Grid>
            </TabItem>

            <!-- TAB 3: SETTINGS & API KEY -->
            <TabItem Name="tabItemSettings" Header="⚙ Settings &amp; API Key">
                <ScrollViewer VerticalScrollBarVisibility="Auto" Margin="0,10,0,0">
                    <StackPanel MaxWidth="780" HorizontalAlignment="Left">
                        <Border Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="16" Margin="0,0,0,14">
                            <StackPanel>
                                <TextBlock Name="lblSettingsApiHeader" Text="Google Maps API Key" FontSize="16" FontWeight="Bold" Foreground="{DynamicResource TextPrimary}" Margin="0,0,0,8"/>
                                <TextBlock Name="lblSettingsApiDesc" Text="Required for Geocoding API, Routes API v2, and Static Maps API." FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,10"/>

                                <TextBlock Name="lblSettingsApiLabel" Text="API Key:" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,4"/>
                                <Grid Margin="0,0,0,8">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <PasswordBox Name="txtSettingsApiKey"/>
                                    <TextBox Name="txtSettingsApiKeyVisible" Visibility="Collapsed"/>
                                    <Button Name="btnToggleKeyVisibility" Grid.Column="1" Content="👁 Show" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="6,0,0,0" Padding="10,6"/>
                                    <Button Name="btnTestApiKey" Grid.Column="2" Content="🔍 Test Key" Background="#2563EB" Margin="6,0,0,0" Padding="12,6"/>
                                </Grid>

                                <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                                    <CheckBox Name="chkRememberKey" Content="Remember securely on this computer (DPAPI CurrentUser encryption)" IsChecked="True" Foreground="{DynamicResource TextSecondary}" FontSize="12"/>
                                </StackPanel>

                                <TextBlock Name="lblKeyTestResult" Text="" FontSize="12" FontWeight="SemiBold"/>
                            </StackPanel>
                        </Border>

                        <Border Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="16" Margin="0,0,0,14">
                            <StackPanel>
                                <TextBlock Name="lblSettingsPrefHeader" Text="Default Generation Preferences" FontSize="16" FontWeight="Bold" Foreground="{DynamicResource TextPrimary}" Margin="0,0,0,12"/>

                                <TextBlock Name="lblSettingsDefaultRouteType" Text="Default route type:" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,4"/>
                                <ComboBox Name="cmbDefaultRouteType" Margin="0,0,0,12">
                                    <ComboBoxItem Content="Fastest (Najszybsza)" Tag="Fastest" IsSelected="True"/>
                                    <ComboBoxItem Content="Shortest (Najkrótsza)" Tag="Shortest"/>
                                    <ComboBoxItem Content="Eco (Fuel Efficient)" Tag="Eco"/>
                                </ComboBox>

                                <TextBlock Name="lblSettingsDefaultEmission" Text="Default engine type for Eco routes:" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,4"/>
                                <ComboBox Name="cmbDefaultEmission" Margin="0,0,0,12">
                                    <ComboBoxItem Content="Gasoline (Benzyna)" Tag="GASOLINE" IsSelected="True"/>
                                    <ComboBoxItem Content="Diesel" Tag="DIESEL"/>
                                    <ComboBoxItem Content="Hybrid" Tag="HYBRID"/>
                                    <ComboBoxItem Content="Electric" Tag="ELECTRIC"/>
                                </ComboBox>

                                <TextBlock Name="lblSettingsDefaultMapSize" Text="Default dimensions for generated PNG map:" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,4"/>
                                <ComboBox Name="cmbDefaultMapSize" Margin="0,0,0,12">
                                    <ComboBoxItem Content="900 x 600 px (Recommended Standard)" Tag="900x600" IsSelected="True"/>
                                    <ComboBoxItem Content="1024 x 768 px (High Res)" Tag="1024x768"/>
                                    <ComboBoxItem Content="1280 x 720 px (HD 16:9)" Tag="1280x720"/>
                                    <ComboBoxItem Content="640 x 640 px (Square)" Tag="640x640"/>
                                    <ComboBoxItem Content="1600 x 900 px (Full HD 16:9)" Tag="1600x900"/>
                                </ComboBox>

                                <TextBlock Name="lblSettingsOutputDir" Text="Results Output Folder:" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,4"/>
                                <Grid Margin="0,0,0,12">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBox Name="txtSettingsOutputDir"/>
                                    <Button Name="btnBrowseSettingsOutputDir" Grid.Column="1" Content="📂 Browse..." Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="6,0,0,0"/>
                                </Grid>
                            </StackPanel>
                        </Border>

                        <!-- MAP OVERLAY CARD -->
                        <Border Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="16" Margin="0,0,0,14">
                            <StackPanel>
                                <TextBlock Name="lblSettingsOverlayHeader" Text="Map Overlay &amp; Banners (Top / Bottom)" FontSize="16" FontWeight="Bold" Foreground="{DynamicResource TextPrimary}" Margin="0,0,0,6"/>
                                <TextBlock Name="lblSettingsOverlayDesc" Text="Configure whether to display top and bottom banner panels, and choose which properties appear on each panel, line order, and alignment." FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,12" TextWrapping="Wrap"/>

                                <StackPanel Orientation="Horizontal" Margin="0,0,0,12">
                                    <CheckBox Name="chkEnableTopOverlay" Content="Enable Top Banner" IsChecked="True" Foreground="{DynamicResource TextPrimary}" FontSize="12" FontWeight="SemiBold" Margin="0,0,24,0"/>
                                    <CheckBox Name="chkEnableBottomOverlay" Content="Enable Bottom Banner" IsChecked="True" Foreground="{DynamicResource TextPrimary}" FontSize="12" FontWeight="SemiBold"/>
                                </StackPanel>

                                <Border Background="{DynamicResource BgDark}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="6" Padding="10" Margin="0,0,0,10">
                                    <Grid Name="gridOverlayConfig">
                                        <Grid.RowDefinitions>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                        </Grid.RowDefinitions>
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="220"/>
                                            <ColumnDefinition Width="65"/>
                                            <ColumnDefinition Width="135"/>
                                            <ColumnDefinition Width="135"/>
                                            <ColumnDefinition Width="110"/>
                                        </Grid.ColumnDefinitions>

                                        <!-- Header Row -->
                                        <TextBlock Name="lblColPropName" Grid.Row="0" Grid.Column="0" Text="Property" FontWeight="Bold" Foreground="{DynamicResource TextSecondary}" FontSize="12" Margin="4,2,4,8"/>
                                        <TextBlock Name="lblColPropShow" Grid.Row="0" Grid.Column="1" Text="Show" FontWeight="Bold" Foreground="{DynamicResource TextSecondary}" FontSize="12" Margin="4,2,4,8" HorizontalAlignment="Center"/>
                                        <TextBlock Name="lblColPropPanel" Grid.Row="0" Grid.Column="2" Text="Panel" FontWeight="Bold" Foreground="{DynamicResource TextSecondary}" FontSize="12" Margin="4,2,4,8"/>
                                        <TextBlock Name="lblColPropAlign" Grid.Row="0" Grid.Column="3" Text="Alignment" FontWeight="Bold" Foreground="{DynamicResource TextSecondary}" FontSize="12" Margin="4,2,4,8"/>
                                        <TextBlock Name="lblColPropOrder" Grid.Row="0" Grid.Column="4" Text="Line / Order" FontWeight="Bold" Foreground="{DynamicResource TextSecondary}" FontSize="12" Margin="4,2,4,8" HorizontalAlignment="Center"/>

                                        <!-- Row 1: StartGeocoded -->
                                        <TextBlock Name="lblProp_StartGeocoded" Grid.Row="1" Grid.Column="0" Text="Start Address (Geocoded)" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_StartGeocoded" Grid.Row="1" Grid.Column="1" IsChecked="True" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_StartGeocoded" Grid.Row="1" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom" IsSelected="True"/>
                                            <ComboBoxItem Content="Top" Tag="Top"/>
                                            <ComboBoxItem Content="None" Tag="None"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_StartGeocoded" Grid.Row="1" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left" IsSelected="True"/>
                                            <ComboBoxItem Content="Center" Tag="Center"/>
                                            <ComboBoxItem Content="Right" Tag="Right"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_StartGeocoded" Grid.Row="1" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1" IsSelected="True"/>
                                            <ComboBoxItem Content="2" Tag="2"/>
                                            <ComboBoxItem Content="3" Tag="3"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>

                                        <!-- Row 2: EndGeocoded -->
                                        <TextBlock Name="lblProp_EndGeocoded" Grid.Row="2" Grid.Column="0" Text="End Address (Geocoded)" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_EndGeocoded" Grid.Row="2" Grid.Column="1" IsChecked="True" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_EndGeocoded" Grid.Row="2" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom" IsSelected="True"/>
                                            <ComboBoxItem Content="Top" Tag="Top"/>
                                            <ComboBoxItem Content="None" Tag="None"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_EndGeocoded" Grid.Row="2" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left" IsSelected="True"/>
                                            <ComboBoxItem Content="Center" Tag="Center"/>
                                            <ComboBoxItem Content="Right" Tag="Right"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_EndGeocoded" Grid.Row="2" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1"/>
                                            <ComboBoxItem Content="2" Tag="2" IsSelected="True"/>
                                            <ComboBoxItem Content="3" Tag="3"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>

                                        <!-- Row 3: Distance -->
                                        <TextBlock Name="lblProp_Distance" Grid.Row="3" Grid.Column="0" Text="Total Distance" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_Distance" Grid.Row="3" Grid.Column="1" IsChecked="True" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_Distance" Grid.Row="3" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom" IsSelected="True"/>
                                            <ComboBoxItem Content="Top" Tag="Top"/>
                                            <ComboBoxItem Content="None" Tag="None"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_Distance" Grid.Row="3" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left" IsSelected="True"/>
                                            <ComboBoxItem Content="Center" Tag="Center"/>
                                            <ComboBoxItem Content="Right" Tag="Right"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_Distance" Grid.Row="3" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1"/>
                                            <ComboBoxItem Content="2" Tag="2"/>
                                            <ComboBoxItem Content="3" Tag="3" IsSelected="True"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>

                                        <!-- Row 4: Duration -->
                                        <TextBlock Name="lblProp_Duration" Grid.Row="4" Grid.Column="0" Text="Total Time" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_Duration" Grid.Row="4" Grid.Column="1" IsChecked="True" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_Duration" Grid.Row="4" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom" IsSelected="True"/>
                                            <ComboBoxItem Content="Top" Tag="Top"/>
                                            <ComboBoxItem Content="None" Tag="None"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_Duration" Grid.Row="4" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left"/>
                                            <ComboBoxItem Content="Center" Tag="Center" IsSelected="True"/>
                                            <ComboBoxItem Content="Right" Tag="Right"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_Duration" Grid.Row="4" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1"/>
                                            <ComboBoxItem Content="2" Tag="2"/>
                                            <ComboBoxItem Content="3" Tag="3" IsSelected="True"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>

                                        <!-- Row 5: Timestamp -->
                                        <TextBlock Name="lblProp_Timestamp" Grid.Row="5" Grid.Column="0" Text="Generation Timestamp" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_Timestamp" Grid.Row="5" Grid.Column="1" IsChecked="True" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_Timestamp" Grid.Row="5" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom" IsSelected="True"/>
                                            <ComboBoxItem Content="Top" Tag="Top"/>
                                            <ComboBoxItem Content="None" Tag="None"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_Timestamp" Grid.Row="5" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left"/>
                                            <ComboBoxItem Content="Center" Tag="Center"/>
                                            <ComboBoxItem Content="Right" Tag="Right" IsSelected="True"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_Timestamp" Grid.Row="5" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1"/>
                                            <ComboBoxItem Content="2" Tag="2"/>
                                            <ComboBoxItem Content="3" Tag="3" IsSelected="True"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>

                                        <!-- Row 6: RouteName -->
                                        <TextBlock Name="lblProp_RouteName" Grid.Row="6" Grid.Column="0" Text="Route Name" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_RouteName" Grid.Row="6" Grid.Column="1" IsChecked="True" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_RouteName" Grid.Row="6" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom"/>
                                            <ComboBoxItem Content="Top" Tag="Top" IsSelected="True"/>
                                            <ComboBoxItem Content="None" Tag="None"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_RouteName" Grid.Row="6" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left" IsSelected="True"/>
                                            <ComboBoxItem Content="Center" Tag="Center"/>
                                            <ComboBoxItem Content="Right" Tag="Right"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_RouteName" Grid.Row="6" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1" IsSelected="True"/>
                                            <ComboBoxItem Content="2" Tag="2"/>
                                            <ComboBoxItem Content="3" Tag="3"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>

                                        <!-- Row 7: RouteType -->
                                        <TextBlock Name="lblProp_RouteType" Grid.Row="7" Grid.Column="0" Text="Route Type" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_RouteType" Grid.Row="7" Grid.Column="1" IsChecked="True" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_RouteType" Grid.Row="7" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom"/>
                                            <ComboBoxItem Content="Top" Tag="Top" IsSelected="True"/>
                                            <ComboBoxItem Content="None" Tag="None"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_RouteType" Grid.Row="7" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left"/>
                                            <ComboBoxItem Content="Center" Tag="Center"/>
                                            <ComboBoxItem Content="Right" Tag="Right" IsSelected="True"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_RouteType" Grid.Row="7" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1" IsSelected="True"/>
                                            <ComboBoxItem Content="2" Tag="2"/>
                                            <ComboBoxItem Content="3" Tag="3"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>

                                        <!-- Row 8: Waypoints -->
                                        <TextBlock Name="lblProp_Waypoints" Grid.Row="8" Grid.Column="0" Text="Intermediate Stops (Waypoints)" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_Waypoints" Grid.Row="8" Grid.Column="1" IsChecked="False" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_Waypoints" Grid.Row="8" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom" IsSelected="True"/>
                                            <ComboBoxItem Content="Top" Tag="Top"/>
                                            <ComboBoxItem Content="None" Tag="None"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_Waypoints" Grid.Row="8" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left" IsSelected="True"/>
                                            <ComboBoxItem Content="Center" Tag="Center"/>
                                            <ComboBoxItem Content="Right" Tag="Right"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_Waypoints" Grid.Row="8" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1"/>
                                            <ComboBoxItem Content="2" Tag="2" IsSelected="True"/>
                                            <ComboBoxItem Content="3" Tag="3"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>

                                        <!-- Row 9: StartRaw -->
                                        <TextBlock Name="lblProp_StartRaw" Grid.Row="9" Grid.Column="0" Text="Start Address (Raw Input)" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_StartRaw" Grid.Row="9" Grid.Column="1" IsChecked="False" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_StartRaw" Grid.Row="9" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom"/>
                                            <ComboBoxItem Content="Top" Tag="Top"/>
                                            <ComboBoxItem Content="None" Tag="None" IsSelected="True"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_StartRaw" Grid.Row="9" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left" IsSelected="True"/>
                                            <ComboBoxItem Content="Center" Tag="Center"/>
                                            <ComboBoxItem Content="Right" Tag="Right"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_StartRaw" Grid.Row="9" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1" IsSelected="True"/>
                                            <ComboBoxItem Content="2" Tag="2"/>
                                            <ComboBoxItem Content="3" Tag="3"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>

                                        <!-- Row 10: EndRaw -->
                                        <TextBlock Name="lblProp_EndRaw" Grid.Row="10" Grid.Column="0" Text="End Address (Raw Input)" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_EndRaw" Grid.Row="10" Grid.Column="1" IsChecked="False" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_EndRaw" Grid.Row="10" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom"/>
                                            <ComboBoxItem Content="Top" Tag="Top"/>
                                            <ComboBoxItem Content="None" Tag="None" IsSelected="True"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_EndRaw" Grid.Row="10" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left" IsSelected="True"/>
                                            <ComboBoxItem Content="Center" Tag="Center"/>
                                            <ComboBoxItem Content="Right" Tag="Right"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_EndRaw" Grid.Row="10" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1"/>
                                            <ComboBoxItem Content="2" Tag="2" IsSelected="True"/>
                                            <ComboBoxItem Content="3" Tag="3"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>
                                    </Grid>
                                </Border>

                                <StackPanel Orientation="Horizontal" Margin="0,2,0,0">
                                    <Button Name="btnResetOverlayConfig" Content="🔄 Reset to Default Layout" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Padding="12,6"/>
                                </StackPanel>
                            </StackPanel>
                        </Border>

                        <Border Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="16" Margin="0,0,0,14">
                            <StackPanel>
                                <TextBlock Name="lblSettingsLangHeader" Text="Language &amp; Localization" FontSize="16" FontWeight="Bold" Foreground="{DynamicResource TextPrimary}" Margin="0,0,0,8"/>
                                <TextBlock Name="lblSettingsLangLabel" Text="Application and Google Maps API Language:" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,6"/>
                                <ComboBox Name="cmbSettingsLanguage" Margin="0,0,0,10"/>
                                <StackPanel Orientation="Horizontal">
                                    <Button Name="btnOpenLangFile" Content="📂 Open Localization File (localization.json)" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Padding="10,6" Margin="0,0,8,0" ToolTip="Open the external localization file to edit or add new languages"/>
                                    <Button Name="btnReloadLang" Content="🔄 Reload Languages" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Padding="10,6" ToolTip="Reload language definitions from disk"/>
                                </StackPanel>
                            </StackPanel>
                        </Border>

                        <Border Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="16" Margin="0,0,0,14">
                            <StackPanel>
                                <TextBlock Name="lblSettingsThemeHeader" Text="Appearance &amp; Theme" FontSize="16" FontWeight="Bold" Foreground="{DynamicResource TextPrimary}" Margin="0,0,0,8"/>
                                <TextBlock Name="lblSettingsThemeLabel" Text="Application Theme (Color Scheme):" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,6"/>
                                <ComboBox Name="cmbSettingsTheme" Margin="0,0,0,0">
                                    <ComboBoxItem Content="🌙 Dark" Tag="Dark" IsSelected="True"/>
                                    <ComboBoxItem Content="☀️ Light" Tag="Light"/>
                                </ComboBox>
                            </StackPanel>
                        </Border>

                        <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                            <Button Name="btnSaveSettings" Content="💾 SAVE SETTINGS" Background="#10B981" Foreground="#FFFFFF" Padding="14,10" FontWeight="Bold" Width="200"/>
                            <Button Name="btnOpenLogFile" Content="📋 OPEN LOG FILE" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Padding="14,10" FontWeight="SemiBold" Margin="10,0,0,0"/>
                        </StackPanel>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
        </TabControl>

        <!-- Footer -->
        <Border Grid.Row="2" Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="6" Padding="10,6" Margin="0,10,0,0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Name="lblFooterStatus" Text="Ready." Foreground="{DynamicResource TextSecondary}" FontSize="12" VerticalAlignment="Center"/>
                <TextBlock Name="lblFooterVersion" Grid.Column="1" Text="Google Maps Routes v2.0" Foreground="{DynamicResource TextSecondary}" FontSize="12" VerticalAlignment="Center"/>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# ── 8. Tworzenie okna WPF z XAML ─────────────────────────────────────────────
$reader = [System.Xml.XmlNodeReader]::new($xaml)
$window = [System.Windows.Markup.XamlReader]::Load($reader)

# Zastosowanie DWM Dark Mode dla okna (zgodnie z motywem)
$window.Add_SourceInitialized({
    try {
        $helper = [System.Windows.Interop.WindowInteropHelper]::new($window)
        $val = if ($script:CurrentTheme -eq 'Light') { 0 } else { 1 }
        [DwmDarkWindow]::DwmSetWindowAttribute($helper.Handle, 20, [ref]$val, 4)
    } catch {}
})

# ── 9. Pobranie referencji do elementów UI ───────────────────────────────────
$txtHeaderTitle      = $window.FindName('txtHeaderTitle')
$txtHeaderSubtitle   = $window.FindName('txtHeaderSubtitle')
$cmbAppLanguage      = $window.FindName('cmbAppLanguage')
$btnThemeToggle      = $window.FindName('btnThemeToggle')
$lblApiBadge         = $window.FindName('lblApiBadge')
$btnQuickSettings    = $window.FindName('btnQuickSettings')
$tabMain             = $window.FindName('tabMain')
$tabItemManual       = $window.FindName('tabItemManual')
$tabItemBatch        = $window.FindName('tabItemBatch')
$tabItemSettings     = $window.FindName('tabItemSettings')
$lblFooterStatus     = $window.FindName('lblFooterStatus')
$lblFooterVersion    = $window.FindName('lblFooterVersion')

# Tab 1: Manual
$lblManualRoutePointsHeader = $window.FindName('lblManualRoutePointsHeader')
$lblManualOrigin            = $window.FindName('lblManualOrigin')
$lblManualWaypoints         = $window.FindName('lblManualWaypoints')
$lblManualDestination       = $window.FindName('lblManualDestination')
$lblManualRouteName         = $window.FindName('lblManualRouteName')
$lblManualOptHeader         = $window.FindName('lblManualOptHeader')
$lblManualEmission          = $window.FindName('lblManualEmission')
$lblHeaderDist              = $window.FindName('lblHeaderDist')
$lblHeaderDur               = $window.FindName('lblHeaderDur')
$lblHeaderType              = $window.FindName('lblHeaderType')

# Tab 1: Manual
$txtManualStart      = $window.FindName('txtManualStart')
$btnClearManualStart = $window.FindName('btnClearManualStart')
$txtNewWaypoint      = $window.FindName('txtNewWaypoint')
$btnAddWaypoint      = $window.FindName('btnAddWaypoint')
$lstWaypoints        = $window.FindName('lstWaypoints')
$btnWpUp             = $window.FindName('btnWpUp')
$btnWpDown           = $window.FindName('btnWpDown')
$btnWpRemove         = $window.FindName('btnWpRemove')
$btnWpClear          = $window.FindName('btnWpClear')
$txtManualEnd        = $window.FindName('txtManualEnd')
$btnClearManualEnd   = $window.FindName('btnClearManualEnd')
$txtManualName       = $window.FindName('txtManualName')
$rbTypeFastest       = $window.FindName('rbTypeFastest')
$rbTypeShortest      = $window.FindName('rbTypeShortest')
$rbTypeEco           = $window.FindName('rbTypeEco')
$pnlEmission         = $window.FindName('pnlEmission')
$cmbEmission         = $window.FindName('cmbEmission')
$chkTrafficAware     = $window.FindName('chkTrafficAware')
$btnCalculateManual  = $window.FindName('btnCalculateManual')
$lblManualDist       = $window.FindName('lblManualDist')
$lblManualTime       = $window.FindName('lblManualTime')
$lblManualType       = $window.FindName('lblManualType')
$lblManualStatus     = $window.FindName('lblManualStatus')
$lblMapPlaceholder   = $window.FindName('lblMapPlaceholder')
$imgMapPreview       = $window.FindName('imgMapPreview')
$lblGoogleUrlDisplay = $window.FindName('lblGoogleUrlDisplay')
$btnOpenGoogleMaps   = $window.FindName('btnOpenGoogleMaps')
$btnCopyUrl          = $window.FindName('btnCopyUrl')
$btnSaveMapAs        = $window.FindName('btnSaveMapAs')

# Tab 2: Batch
$lblBatchInputFile       = $window.FindName('lblBatchInputFile')
$lblBatchDefaultRouteType= $window.FindName('lblBatchDefaultRouteType')
$tabSubInput             = $window.FindName('tabSubInput')
$tabSubResults           = $window.FindName('tabSubResults')
$tabSubPoints            = $window.FindName('tabSubPoints')
$tabSubLog               = $window.FindName('tabSubLog')
$txtBatchFilePath    = $window.FindName('txtBatchFilePath')
$btnBrowseBatchFile  = $window.FindName('btnBrowseBatchFile')
$btnReloadBatchFile  = $window.FindName('btnReloadBatchFile')
$lblBatchFileInfo    = $window.FindName('lblBatchFileInfo')
$cmbBatchRouteType   = $window.FindName('cmbBatchRouteType')
$btnStartBatch       = $window.FindName('btnStartBatch')
$btnStopBatch        = $window.FindName('btnStopBatch')
$tabBatchSub        = $window.FindName('tabBatchSub')
$dgBatchInput        = $window.FindName('dgBatchInput')
$dgBatchResults      = $window.FindName('dgBatchResults')
$dgBatchPoints       = $window.FindName('dgBatchPoints')
$txtBatchLog         = $window.FindName('txtBatchLog')
$lblBatchProgressText= $window.FindName('lblBatchProgressText')
$pbBatchProgress     = $window.FindName('pbBatchProgress')
$lblBatchStats       = $window.FindName('lblBatchStats')
$btnOpenOutputDir    = $window.FindName('btnOpenOutputDir')
$btnExportExcel      = $window.FindName('btnExportExcel')
$btnExportCsv        = $window.FindName('btnExportCsv')
$btnExportJson       = $window.FindName('btnExportJson')

# Tab 3: Settings
$txtSettingsApiKey          = $window.FindName('txtSettingsApiKey')
$txtSettingsApiKeyVisible   = $window.FindName('txtSettingsApiKeyVisible')
$btnToggleKeyVisibility     = $window.FindName('btnToggleKeyVisibility')
$btnTestApiKey              = $window.FindName('btnTestApiKey')
$chkRememberKey             = $window.FindName('chkRememberKey')
$lblKeyTestResult           = $window.FindName('lblKeyTestResult')
$cmbDefaultRouteType        = $window.FindName('cmbDefaultRouteType')
$cmbDefaultEmission         = $window.FindName('cmbDefaultEmission')
$cmbDefaultMapSize          = $window.FindName('cmbDefaultMapSize')
$txtSettingsOutputDir       = $window.FindName('txtSettingsOutputDir')
$btnBrowseSettingsOutputDir = $window.FindName('btnBrowseSettingsOutputDir')
$btnSaveSettings            = $window.FindName('btnSaveSettings')
$btnOpenLogFile             = $window.FindName('btnOpenLogFile')
$lblSettingsApiHeader       = $window.FindName('lblSettingsApiHeader')
$lblSettingsApiDesc         = $window.FindName('lblSettingsApiDesc')
$lblSettingsApiLabel        = $window.FindName('lblSettingsApiLabel')
$lblSettingsPrefHeader      = $window.FindName('lblSettingsPrefHeader')
$lblSettingsDefaultRouteType= $window.FindName('lblSettingsDefaultRouteType')
$lblSettingsDefaultEmission = $window.FindName('lblSettingsDefaultEmission')
$lblSettingsDefaultMapSize  = $window.FindName('lblSettingsDefaultMapSize')
$lblSettingsOutputDir       = $window.FindName('lblSettingsOutputDir')
$lblSettingsLangHeader      = $window.FindName('lblSettingsLangHeader')
$lblSettingsLangLabel       = $window.FindName('lblSettingsLangLabel')
$cmbSettingsLanguage        = $window.FindName('cmbSettingsLanguage')
$btnOpenLangFile            = $window.FindName('btnOpenLangFile')
$btnReloadLang              = $window.FindName('btnReloadLang')
$lblSettingsThemeHeader     = $window.FindName('lblSettingsThemeHeader')
$lblSettingsThemeLabel      = $window.FindName('lblSettingsThemeLabel')
$cmbSettingsTheme           = $window.FindName('cmbSettingsTheme')

# Tab 3: Overlay Settings
$lblSettingsOverlayHeader    = $window.FindName('lblSettingsOverlayHeader')
$lblSettingsOverlayDesc      = $window.FindName('lblSettingsOverlayDesc')
$chkEnableTopOverlay         = $window.FindName('chkEnableTopOverlay')
$chkEnableBottomOverlay      = $window.FindName('chkEnableBottomOverlay')
$lblColPropName              = $window.FindName('lblColPropName')
$lblColPropShow              = $window.FindName('lblColPropShow')
$lblColPropPanel             = $window.FindName('lblColPropPanel')
$lblColPropAlign             = $window.FindName('lblColPropAlign')
$lblColPropOrder             = $window.FindName('lblColPropOrder')
$btnResetOverlayConfig       = $window.FindName('btnResetOverlayConfig')

foreach ($key in $script:OverlayPropKeys) {
    Set-Variable -Name "lblProp_$key"  -Value ($window.FindName("lblProp_$key"))  -Scope Script
    Set-Variable -Name "chkProp_$key"  -Value ($window.FindName("chkProp_$key"))  -Scope Script
    Set-Variable -Name "cmbPanel_$key" -Value ($window.FindName("cmbPanel_$key")) -Scope Script
    Set-Variable -Name "cmbAlign_$key" -Value ($window.FindName("cmbAlign_$key")) -Scope Script
    Set-Variable -Name "cmbOrder_$key" -Value ($window.FindName("cmbOrder_$key")) -Scope Script
}

# ── 10. System wielojęzyczności i funkcje pomocnicze stanu UI ─────────────────

function Populate-LanguageDropdowns {
    $script:SuppressLangEvents = $true
    try {
        if ($cmbAppLanguage) {
            $cmbAppLanguage.Items.Clear()
            foreach ($langCode in $script:LanguagesCatalog.Keys) {
                $langObj = $script:LanguagesCatalog[$langCode]
                $cbi = [System.Windows.Controls.ComboBoxItem]::new()
                $cbi.Content = "[$($langObj.Code.ToUpper())] $($langObj.DisplayName)"
                $cbi.Tag = $langObj.Code
                if ($langObj.Code -eq $script:CurrentLanguage) { $cbi.IsSelected = $true }
                $null = $cmbAppLanguage.Items.Add($cbi)
            }
        }
        if ($cmbSettingsLanguage) {
            $cmbSettingsLanguage.Items.Clear()
            foreach ($langCode in $script:LanguagesCatalog.Keys) {
                $langObj = $script:LanguagesCatalog[$langCode]
                $cbi = [System.Windows.Controls.ComboBoxItem]::new()
                $cbi.Content = "[$($langObj.Code.ToUpper())] $($langObj.DisplayName)"
                $cbi.Tag = $langObj.Code
                if ($langObj.Code -eq $script:CurrentLanguage) { $cbi.IsSelected = $true }
                $null = $cmbSettingsLanguage.Items.Add($cbi)
            }
        }
    } finally {
        $script:SuppressLangEvents = $false
    }
}

function Set-AppTheme {
    param(
        [Parameter(Mandatory = $false)]
        [string]$Theme = 'Dark'
    )

    if ($Theme -notmatch '(?i)light|dark') { $Theme = 'Dark' }
    $isLight = ($Theme -match '(?i)light')
    $script:CurrentTheme = if ($isLight) { 'Light' } else { 'Dark' }

    $palette = if ($isLight) {
        [ordered]@{
            'BgDark'                 = '#F1F5F9'
            'BgCard'                 = '#FFFFFF'
            'BgCardHover'            = '#F8FAFC'
            'BgCardAlt'              = '#F8FAFC'
            'BorderCard'             = '#CBD5E1'
            'TextPrimary'            = '#0F172A'
            'TextSecondary'          = '#475569'
            'AccentBlue'             = '#2563EB'
            'AccentGreen'            = '#059669'
            'AccentAmber'            = '#D97706'
            'AccentRed'              = '#DC2626'
            'BgInput'                = '#FFFFFF'
            'BorderInput'            = '#CBD5E1'
            'BtnSecondaryBg'         = '#E2E8F0'
            'BtnSecondaryFg'         = '#0F172A'
            'GridLines'              = '#E2E8F0'
            'LogBg'                  = '#F8FAFC'
            'LogFg'                  = '#0369A1'
            'DataGridHeaderBg'       = '#E2E8F0'
            'DataGridHeaderFg'       = '#334155'
            'DataGridRowBg'          = '#FFFFFF'
            'DataGridAltRowBg'       = '#F8FAFC'
        }
    } else {
        [ordered]@{
            'BgDark'                 = '#0F172A'
            'BgCard'                 = '#1E293B'
            'BgCardHover'            = '#293548'
            'BgCardAlt'              = '#162032'
            'BorderCard'             = '#334155'
            'TextPrimary'            = '#F8FAFC'
            'TextSecondary'          = '#94A3B8'
            'AccentBlue'             = '#2563EB'
            'AccentGreen'            = '#10B981'
            'AccentAmber'            = '#F59E0B'
            'AccentRed'              = '#EF4444'
            'BgInput'                = '#1E293B'
            'BorderInput'            = '#334155'
            'BtnSecondaryBg'         = '#334155'
            'BtnSecondaryFg'         = '#F8FAFC'
            'GridLines'              = '#2D3748'
            'LogBg'                  = '#0A0F1D'
            'LogFg'                  = '#38BDF8'
            'DataGridHeaderBg'       = '#0F172A'
            'DataGridHeaderFg'       = '#94A3B8'
            'DataGridRowBg'          = '#1E293B'
            'DataGridAltRowBg'       = '#162032'
        }
    }

    foreach ($k in $palette.Keys) {
        $c = [System.Windows.Media.ColorConverter]::ConvertFromString($palette[$k])
        $brush = [System.Windows.Media.SolidColorBrush]::new($c)
        $brush.Freeze()
        $window.Resources[$k] = $brush
        $window.Resources["Theme_$k"] = $brush
    }
    $window.Resources['Theme_BgApp'] = $window.Resources['BgDark']
    $window.Resources['Theme_Border'] = $window.Resources['BorderCard']

    if ($window) {
        $window.Background = $window.Resources['BgDark']
        $window.Foreground = $window.Resources['TextPrimary']
    }

    # Update overlay property labels foreground
    if ($script:OverlayPropKeys) {
        foreach ($k in $script:OverlayPropKeys) {
            $lblCtrl = Get-Variable -Name "lblProp_$k" -ValueOnly -ErrorAction SilentlyContinue
            if ($lblCtrl) {
                $lblCtrl.Foreground = $window.Resources['TextPrimary']
            }
        }
    }

    # Update DWM title bar chrome
    try {
        $helper = [System.Windows.Interop.WindowInteropHelper]::new($window)
        if ($helper.Handle -ne [IntPtr]::Zero) {
            $val = if ($isLight) { 0 } else { 1 }
            [DwmDarkWindow]::DwmSetWindowAttribute($helper.Handle, 20, [ref]$val, 4)
        }
    } catch {}

    # Update Toggle Button text/icon
    if ($btnThemeToggle) {
        $btnThemeToggle.Content = if ($isLight) { (Get-LocText 'ThemeLight') } else { (Get-LocText 'ThemeDark') }
    }

    # Update Settings ComboBox
    if ($cmbSettingsTheme) {
        $script:SuppressThemeEvents = $true
        try {
            foreach ($it in $cmbSettingsTheme.Items) {
                if ($it.Tag -eq $script:CurrentTheme) {
                    $cmbSettingsTheme.SelectedItem = $it
                    break
                }
            }
        } finally {
            $script:SuppressThemeEvents = $false
        }
    }
}

function Apply-AppLanguage {
    param([string]$LanguageCode)

    if (-not [string]::IsNullOrWhiteSpace($LanguageCode) -and $script:LanguagesCatalog.Contains($LanguageCode.ToLower())) {
        $script:CurrentLanguage = $LanguageCode.ToLower()
    } else {
        $script:CurrentLanguage = 'en'
    }

    $langObj = $script:LanguagesCatalog[$script:CurrentLanguage]
    $script:CurrentGoogleLang = if ($langObj -and $langObj.GoogleCode) { $langObj.GoogleCode } else { $script:CurrentLanguage }
    $script:CurrentStrings = if ($langObj -and $langObj.Strings) { $langObj.Strings } else { @{} }

    # Synchronize ComboBoxes without triggering duplicate events
    $script:SuppressLangEvents = $true
    try {
        if ($cmbAppLanguage) {
            foreach ($item in $cmbAppLanguage.Items) {
                if ($item.Tag -eq $script:CurrentLanguage) {
                    $cmbAppLanguage.SelectedItem = $item
                    break
                }
            }
        }
        if ($cmbSettingsLanguage) {
            foreach ($item in $cmbSettingsLanguage.Items) {
                if ($item.Tag -eq $script:CurrentLanguage) {
                    $cmbSettingsLanguage.SelectedItem = $item
                    break
                }
            }
        }
    } finally {
        $script:SuppressLangEvents = $false
    }

    # Window & Header
    if ($window) { $window.Title = (Get-LocText 'AppTitle') }
    if ($txtHeaderTitle) { $txtHeaderTitle.Text = (Get-LocText 'AppTitle') }
    if ($txtHeaderSubtitle) { $txtHeaderSubtitle.Text = (Get-LocText 'AppSubtitle') }
    if ($btnQuickSettings) { $btnQuickSettings.Content = (Get-LocText 'BtnQuickSettings') }
    if ($btnThemeToggle) {
        $btnThemeToggle.ToolTip = (Get-LocText 'ThemeToggleTip')
        $btnThemeToggle.Content = if ($script:CurrentTheme -eq 'Light') { (Get-LocText 'ThemeLight') } else { (Get-LocText 'ThemeDark') }
    }

    # Tab Headers
    if ($tabItemManual) { $tabItemManual.Header = (Get-LocText 'TabManual') }
    if ($tabItemBatch) { $tabItemBatch.Header = (Get-LocText 'TabBatch') }
    if ($tabItemSettings) { $tabItemSettings.Header = (Get-LocText 'TabSettings') }

    # Tab 1: Manual Route
    if ($lblManualRoutePointsHeader) { $lblManualRoutePointsHeader.Text = (Get-LocText 'ManualHeaderRoutePoints') }
    if ($lblManualOrigin) { $lblManualOrigin.Text = (Get-LocText 'ManualOrigin') }
    if ($lblManualWaypoints) { $lblManualWaypoints.Text = (Get-LocText 'ManualWaypoints') }
    if ($txtNewWaypoint) { $txtNewWaypoint.ToolTip = (Get-LocText 'ManualWaypointsTooltip') }
    if ($btnAddWaypoint) { $btnAddWaypoint.Content = (Get-LocText 'ManualBtnAdd') }
    if ($btnWpUp) { $btnWpUp.Content = (Get-LocText 'ManualBtnUp') }
    if ($btnWpDown) { $btnWpDown.Content = (Get-LocText 'ManualBtnDown') }
    if ($btnWpRemove) { $btnWpRemove.Content = (Get-LocText 'ManualBtnRemove') }
    if ($btnWpClear) { $btnWpClear.Content = (Get-LocText 'ManualBtnClear') }
    if ($lblManualDestination) { $lblManualDestination.Text = (Get-LocText 'ManualDestination') }
    if ($lblManualRouteName) { $lblManualRouteName.Text = (Get-LocText 'ManualRouteName') }
    if ($lblManualOptHeader) { $lblManualOptHeader.Text = (Get-LocText 'ManualHeaderOptimization') }
    if ($rbTypeFastest) { $rbTypeFastest.Content = (Get-LocText 'ManualOptFastest') }
    if ($rbTypeShortest) { $rbTypeShortest.Content = (Get-LocText 'ManualOptShortest') }
    if ($rbTypeEco) { $rbTypeEco.Content = (Get-LocText 'ManualOptEco') }
    if ($lblManualEmission) { $lblManualEmission.Text = (Get-LocText 'ManualEmission') }
    if ($chkTrafficAware) { $chkTrafficAware.Content = (Get-LocText 'ManualTrafficAware') }
    if ($btnCalculateManual -and $btnCalculateManual.IsEnabled) { $btnCalculateManual.Content = (Get-LocText 'ManualBtnCalculate') }
    if ($lblHeaderDist) { $lblHeaderDist.Text = (Get-LocText 'ManualStatDistance') }
    if ($lblHeaderDur) { $lblHeaderDur.Text = (Get-LocText 'ManualStatDuration') }
    if ($lblHeaderType) { $lblHeaderType.Text = (Get-LocText 'ManualStatType') }
    if ($lblManualType) {
        $curT = $lblManualType.Text
        if ($curT -match '(?i)fast|szyb|schnell') {
            $lblManualType.Text = switch ($script:CurrentLanguage) { 'de' { 'Schnellste' } 'pl' { 'Najszybsza' } default { 'Fastest' } }
        } elseif ($curT -match '(?i)short|kr[oó]t|k[uü]rz') {
            $lblManualType.Text = switch ($script:CurrentLanguage) { 'de' { 'Kürzeste' } 'pl' { 'Najkrótsza' } default { 'Shortest' } }
        } elseif ($curT -match '(?i)eco|eko') {
            $lblManualType.Text = switch ($script:CurrentLanguage) { 'pl' { 'Eko' } default { 'Eco' } }
        }
    }
    if ($lblManualStatus -and $lblManualStatus.Text -match '(?i)idle|bereit|gotow') { $lblManualStatus.Text = (Get-LocText 'ManualStatusIdle') }
    if ($lblMapPlaceholder -and $lblMapPlaceholder.Visibility -eq [System.Windows.Visibility]::Visible) { $lblMapPlaceholder.Text = (Get-LocText 'ManualMapPlaceholder') }
    if ($lblGoogleUrlDisplay -and $lblGoogleUrlDisplay.Text -match '(?i)no generated|kein link|brak') { $lblGoogleUrlDisplay.Text = (Get-LocText 'ManualNoUrl') }
    if ($btnOpenGoogleMaps) { $btnOpenGoogleMaps.Content = (Get-LocText 'ManualBtnGoogleMaps') }
    if ($btnCopyUrl) { $btnCopyUrl.Content = (Get-LocText 'ManualBtnCopyUrl') }
    if ($btnSaveMapAs) { $btnSaveMapAs.Content = (Get-LocText 'ManualBtnSaveMapAs') }

    # Tab 2: Batch Processing
    if ($lblBatchInputFile) { $lblBatchInputFile.Text = (Get-LocText 'BatchInputFile') }
    if ($btnBrowseBatchFile) { $btnBrowseBatchFile.Content = (Get-LocText 'BatchBtnBrowse') }
    if ($btnReloadBatchFile) { $btnReloadBatchFile.Content = (Get-LocText 'BatchBtnReload') }
    if ($lblBatchFileInfo -and $lblBatchFileInfo.Text -match '(?i)no file|keine datei|brak') { $lblBatchFileInfo.Text = (Get-LocText 'BatchNoFileLoaded') }
    if ($lblBatchDefaultRouteType) { $lblBatchDefaultRouteType.Text = (Get-LocText 'BatchDefaultRouteType') }
    if ($btnStartBatch) { $btnStartBatch.Content = (Get-LocText 'BatchBtnStart') }
    if ($btnStopBatch) { $btnStopBatch.Content = (Get-LocText 'BatchBtnStop') }
    if ($tabSubInput) { $tabSubInput.Header = (Get-LocText 'BatchTabInputPreview') }
    if ($tabSubResults) { $tabSubResults.Header = (Get-LocText 'BatchTabResults') }
    if ($tabSubPoints) { $tabSubPoints.Header = (Get-LocText 'BatchTabPoints') }
    if ($tabSubLog) { $tabSubLog.Header = (Get-LocText 'BatchTabLog') }

    # Batch DataGrid Columns
    if ($dgBatchResults -and $dgBatchResults.Columns.Count -ge 10) {
        $dgBatchResults.Columns[0].Header = (Get-LocText 'BatchColId')
        $dgBatchResults.Columns[1].Header = (Get-LocText 'BatchColName')
        $dgBatchResults.Columns[2].Header = (Get-LocText 'BatchColOrigin')
        $dgBatchResults.Columns[3].Header = (Get-LocText 'BatchColDestination')
        $dgBatchResults.Columns[4].Header = (Get-LocText 'BatchColWaypoints')
        $dgBatchResults.Columns[5].Header = (Get-LocText 'BatchColType')
        $dgBatchResults.Columns[6].Header = (Get-LocText 'BatchColDistance')
        $dgBatchResults.Columns[7].Header = (Get-LocText 'BatchColDuration')
        $dgBatchResults.Columns[8].Header = (Get-LocText 'BatchColStatus')
        $dgBatchResults.Columns[9].Header = (Get-LocText 'BatchColMap')
    }

    # Points DataGrid Columns
    if ($dgBatchPoints -and $dgBatchPoints.Columns.Count -ge 9) {
        $dgBatchPoints.Columns[0].Header = (Get-LocText 'PointsColRouteId')
        $dgBatchPoints.Columns[1].Header = (Get-LocText 'PointsColRouteName')
        $dgBatchPoints.Columns[2].Header = (Get-LocText 'PointsColOrder')
        $dgBatchPoints.Columns[3].Header = (Get-LocText 'PointsColType')
        $dgBatchPoints.Columns[4].Header = (Get-LocText 'PointsColOriginalAddress')
        $dgBatchPoints.Columns[5].Header = (Get-LocText 'PointsColGeocodedAddress')
        $dgBatchPoints.Columns[6].Header = (Get-LocText 'PointsColGeocodeStatus')
        $dgBatchPoints.Columns[7].Header = (Get-LocText 'PointsColMatchType')
        $dgBatchPoints.Columns[8].Header = (Get-LocText 'PointsColIsFallback')
        if ($dgBatchPoints.Columns.Count -ge 11) {
            $dgBatchPoints.Columns[9].Header = (Get-LocText 'PointsColLatitude')
            $dgBatchPoints.Columns[10].Header = (Get-LocText 'PointsColLongitude')
        }
    }

    if ($lblBatchProgressText -and $lblBatchProgressText.Text -match '(?i)ready|bereit|gotow') { $lblBatchProgressText.Text = (Get-LocText 'BatchProgressReady') }
    if ($btnOpenOutputDir) { $btnOpenOutputDir.Content = (Get-LocText 'BatchBtnOpenOutputDir') }
    if ($btnExportExcel) { $btnExportExcel.Content = (Get-LocText 'BatchBtnExportExcel') }
    if ($btnExportCsv) { $btnExportCsv.Content = (Get-LocText 'BatchBtnExportCsv') }
    if ($btnExportJson) { $btnExportJson.Content = (Get-LocText 'BatchBtnExportJson') }

    # Tab 3: Settings
    if ($lblSettingsApiHeader) { $lblSettingsApiHeader.Text = (Get-LocText 'SettingsHeaderApi') }
    if ($lblSettingsApiDesc) { $lblSettingsApiDesc.Text = (Get-LocText 'SettingsApiDesc') }
    if ($lblSettingsApiLabel) { $lblSettingsApiLabel.Text = (Get-LocText 'SettingsApiLabel') }
    if ($btnTestApiKey -and $btnTestApiKey.IsEnabled) { $btnTestApiKey.Content = (Get-LocText 'SettingsBtnTestKey') }
    if ($chkRememberKey) { $chkRememberKey.Content = (Get-LocText 'SettingsChkRemember') }
    if ($lblSettingsPrefHeader) { $lblSettingsPrefHeader.Text = (Get-LocText 'SettingsHeaderPreferences') }
    if ($lblSettingsDefaultRouteType) { $lblSettingsDefaultRouteType.Text = (Get-LocText 'SettingsDefaultRouteType') }
    if ($lblSettingsDefaultEmission) { $lblSettingsDefaultEmission.Text = (Get-LocText 'SettingsDefaultEmission') }
    if ($lblSettingsDefaultMapSize) { $lblSettingsDefaultMapSize.Text = (Get-LocText 'SettingsDefaultMapSize') }
    if ($lblSettingsOutputDir) { $lblSettingsOutputDir.Text = (Get-LocText 'SettingsOutputDir') }
    if ($btnBrowseSettingsOutputDir) { $btnBrowseSettingsOutputDir.Content = (Get-LocText 'SettingsBtnBrowseOutputDir') }
    if ($lblSettingsOverlayHeader) { $lblSettingsOverlayHeader.Text = (Get-LocText 'SettingsHeaderOverlay') }
    if ($lblSettingsOverlayDesc) { $lblSettingsOverlayDesc.Text = (Get-LocText 'SettingsOverlayDesc') }
    if ($chkEnableTopOverlay) { $chkEnableTopOverlay.Content = (Get-LocText 'SettingsOverlayTopEnable') }
    if ($chkEnableBottomOverlay) { $chkEnableBottomOverlay.Content = (Get-LocText 'SettingsOverlayBottomEnable') }
    if ($lblColPropName) { $lblColPropName.Text = (Get-LocText 'OverlayColProperty') }
    if ($lblColPropShow) { $lblColPropShow.Text = (Get-LocText 'OverlayColShow') }
    if ($lblColPropPanel) { $lblColPropPanel.Text = (Get-LocText 'OverlayColPanel') }
    if ($lblColPropAlign) { $lblColPropAlign.Text = (Get-LocText 'OverlayColAlign') }
    if ($lblColPropOrder) { $lblColPropOrder.Text = (Get-LocText 'OverlayColOrder') }
    if ($btnResetOverlayConfig) { $btnResetOverlayConfig.Content = (Get-LocText 'SettingsOverlayBtnReset') }

    foreach ($k in $script:OverlayPropKeys) {
        $lblCtrl = Get-Variable -Name "lblProp_$k" -ValueOnly -ErrorAction SilentlyContinue
        if ($lblCtrl) {
            $lblCtrl.Text = (Get-LocText "OverlayProp$k")
        }
    }

    if ($lblSettingsLangHeader) { $lblSettingsLangHeader.Text = (Get-LocText 'SettingsHeaderLanguage') }
    if ($lblSettingsLangLabel) { $lblSettingsLangLabel.Text = (Get-LocText 'SettingsLanguageLabel') }
    if ($btnOpenLangFile) { $btnOpenLangFile.Content = (Get-LocText 'SettingsBtnOpenLangFile') }
    if ($btnReloadLang) { $btnReloadLang.Content = (Get-LocText 'SettingsBtnReloadLang') }
    if ($lblSettingsThemeHeader) { $lblSettingsThemeHeader.Text = switch ($script:CurrentLanguage) { 'de' { 'Erscheinungsbild & Design' } 'pl' { 'Wygląd i motyw' } default { 'Appearance & Theme' } } }
    if ($lblSettingsThemeLabel) { $lblSettingsThemeLabel.Text = (Get-LocText 'SettingsThemeLabel') }
    if ($cmbSettingsTheme -and $cmbSettingsTheme.Items.Count -ge 2) {
        $cmbSettingsTheme.Items[0].Content = (Get-LocText 'ThemeDark')
        $cmbSettingsTheme.Items[1].Content = (Get-LocText 'ThemeLight')
    }
    if ($btnSaveSettings) { $btnSaveSettings.Content = (Get-LocText 'SettingsBtnSave') }
    if ($btnOpenLogFile) { $btnOpenLogFile.Content = (Get-LocText 'SettingsBtnOpenLog') }

    # Footer
    if ($lblFooterStatus -and $lblFooterStatus.Text -match '(?i)ready|bereit|gotow') { $lblFooterStatus.Text = (Get-LocText 'FooterReady') }
    if ($lblFooterVersion) { $lblFooterVersion.Text = (Get-LocText 'FooterVersion') }

    # Update API badge text if checking
    if ($lblApiBadge -and $lblApiBadge.Text -match '(?i)checking|prüfe|sprawdz') {
        $lblApiBadge.Text = (Get-LocText 'ApiBadgeChecking')
    }
}

function Set-CurrentApiKey([string]$Key) {
    $txtSettingsApiKey.Password = $Key
    $txtSettingsApiKeyVisible.Text = $Key
    $script:CurrentApiKey = $Key
}

function Get-CurrentApiKey {
    if ($txtSettingsApiKeyVisible.Visibility -eq [System.Windows.Visibility]::Visible) {
        return $txtSettingsApiKeyVisible.Text.Trim()
    }
    return $txtSettingsApiKey.Password.Trim()
}

function Write-BatchLog([string]$Message, [string]$Level = 'INFO') {
    Write-AppLog -Message $Message -Level $Level -ToBatchWindow
}

function Update-ApiStatusBadge($IsValid, [string]$Message) {
    $validBool = [bool]$IsValid
    if ($validBool) {
        $lblApiBadge.Text = 'API: Active'
        $lblApiBadge.Foreground = [System.Windows.Media.Brushes]::LightGreen
        $lblKeyTestResult.Text = "✓ $Message"
        $lblKeyTestResult.Foreground = [System.Windows.Media.Brushes]::LightGreen
    } else {
        $lblApiBadge.Text = 'API: Invalid'
        $lblApiBadge.Foreground = [System.Windows.Media.Brushes]::Salmon
        $lblKeyTestResult.Text = "✕ $Message"
        $lblKeyTestResult.Foreground = [System.Windows.Media.Brushes]::Salmon
    }
}

# Inicjalizacja wartości kontrolek z zapisanego configu
$script:CurrentApiKey = ''
if (-not [string]::IsNullOrWhiteSpace($script:Config.ApiKey)) {
    Set-CurrentApiKey -Key $script:Config.ApiKey
    $lblApiBadge.Text = 'API: Configured'
    $lblApiBadge.Foreground = [System.Windows.Media.Brushes]::LightGreen
}

$chkRememberKey.IsChecked = [bool]$script:Config.RememberApiKey
$txtSettingsOutputDir.Text = if ($script:Config.LastOutputFolder) { $script:Config.LastOutputFolder } else { Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'GoogleMapsRoutes\Results' }

Populate-LanguageDropdowns
Apply-AppLanguage -LanguageCode $script:CurrentLanguage
Set-AppTheme -Theme $script:Config.Theme

# Ustawienie domyślnych ComboBoxów
foreach ($item in $cmbDefaultRouteType.Items) {
    if ($item.Tag -eq $script:Config.DefaultRouteType) { $item.IsSelected = $true; break }
}
foreach ($item in $cmbDefaultEmission.Items) {
    if ($item.Tag -eq $script:Config.DefaultEmission) { $item.IsSelected = $true; break }
}
$targetMapTag = "$($script:Config.MapWidth)x$($script:Config.MapHeight)"
foreach ($item in $cmbDefaultMapSize.Items) {
    if ($item.Tag -eq $targetMapTag) { $item.IsSelected = $true; break }
}

# Ustawienie kontrolek nakładki mapy (Overlay)
if ($script:Config.OverlayConfig) {
    Set-OverlayConfigUi $script:Config.OverlayConfig
} else {
    Reset-OverlayConfigUi
}

$script:LastGeneratedMapPath = $null
$script:LastGoogleMapsUrl    = $null
$script:LoadedBatchData      = $null
$script:BatchResultsList     = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:BatchWorkerRunning   = $false
$script:BatchCancelRequested = $false
$script:LastDataDirectory    = if ($script:Config.LastInputFolder) { $script:Config.LastInputFolder } else { $null }

# Restore last used batch input file if available
if ($script:Config.LastInputPath -and (Test-Path $script:Config.LastInputPath)) {
    $txtBatchFilePath.Text = $script:Config.LastInputPath
    try {
        Load-BatchFilePreview -Path $script:Config.LastInputPath
    } catch { }
}

# ── 11. Zdarzenia: Ustawienia i Klucz API ────────────────────────────────────
$btnQuickSettings.Add_Click({
    $tabMain.SelectedItem = $tabItemSettings
})

if ($btnThemeToggle) {
    $btnThemeToggle.Add_Click({
        $newTheme = if ($script:CurrentTheme -eq 'Light') { 'Dark' } else { 'Light' }
        Set-AppTheme -Theme $newTheme
        $script:Config.Theme = $newTheme
        $routeType = if ($cmbDefaultRouteType.SelectedItem) { $cmbDefaultRouteType.SelectedItem.Tag -as [string] } else { $script:Config.DefaultRouteType }
        $emission = if ($cmbDefaultEmission.SelectedItem) { $cmbDefaultEmission.SelectedItem.Tag -as [string] } else { $script:Config.DefaultEmission }
        Save-AppConfig -ApiKey (Get-CurrentApiKey) -RememberApiKey $chkRememberKey.IsChecked `
            -OutputFolder $txtSettingsOutputDir.Text.Trim() `
            -LastInputFolder $script:Config.LastInputFolder -LastInputPath $script:Config.LastInputPath `
            -DefaultRouteType $routeType `
            -DefaultEmission $emission `
            -Language $script:CurrentLanguage `
            -Theme $newTheme
        Write-AppLog "Theme toggled to: $newTheme" "INFO"
    })
}

if ($cmbSettingsTheme) {
    $cmbSettingsTheme.Add_SelectionChanged({
        if ($script:SuppressThemeEvents) { return }
        $sel = $cmbSettingsTheme.SelectedItem
        if ($sel -and $sel.Tag) {
            $newTheme = [string]$sel.Tag
            if ($newTheme -ne $script:CurrentTheme) {
                Set-AppTheme -Theme $newTheme
                $script:Config.Theme = $newTheme
                $routeType = if ($cmbDefaultRouteType.SelectedItem) { $cmbDefaultRouteType.SelectedItem.Tag -as [string] } else { $script:Config.DefaultRouteType }
                $emission = if ($cmbDefaultEmission.SelectedItem) { $cmbDefaultEmission.SelectedItem.Tag -as [string] } else { $script:Config.DefaultEmission }
                Save-AppConfig -ApiKey (Get-CurrentApiKey) -RememberApiKey $chkRememberKey.IsChecked `
                    -OutputFolder $txtSettingsOutputDir.Text.Trim() `
                    -LastInputFolder $script:Config.LastInputFolder -LastInputPath $script:Config.LastInputPath `
                    -DefaultRouteType $routeType `
                    -DefaultEmission $emission `
                    -Language $script:CurrentLanguage `
                    -Theme $newTheme
                Write-AppLog "Theme changed to: $newTheme" "INFO"
            }
        }
    })
}

$cmbAppLanguage.Add_SelectionChanged({
    if ($script:SuppressLangEvents) { return }
    $sel = $cmbAppLanguage.SelectedItem
    if ($sel -and $sel.Tag) {
        $newLang = [string]$sel.Tag
        Apply-AppLanguage -LanguageCode $newLang
        Save-AppConfig -ApiKey (Get-CurrentApiKey) -RememberApiKey $chkRememberKey.IsChecked `
            -OutputFolder $txtSettingsOutputDir.Text.Trim() `
            -LastInputFolder $script:Config.LastInputFolder -LastInputPath $script:Config.LastInputPath `
            -DefaultRouteType ($cmbDefaultRouteType.SelectedItem.Tag -as [string]) `
            -DefaultEmission ($cmbDefaultEmission.SelectedItem.Tag -as [string]) `
            -Language $newLang
        Write-AppLog "Language changed to: $newLang (Google API: $script:CurrentGoogleLang)" "INFO"
    }
})

$cmbSettingsLanguage.Add_SelectionChanged({
    if ($script:SuppressLangEvents) { return }
    $sel = $cmbSettingsLanguage.SelectedItem
    if ($sel -and $sel.Tag) {
        $newLang = [string]$sel.Tag
        Apply-AppLanguage -LanguageCode $newLang
        Save-AppConfig -ApiKey (Get-CurrentApiKey) -RememberApiKey $chkRememberKey.IsChecked `
            -OutputFolder $txtSettingsOutputDir.Text.Trim() `
            -LastInputFolder $script:Config.LastInputFolder -LastInputPath $script:Config.LastInputPath `
            -DefaultRouteType ($cmbDefaultRouteType.SelectedItem.Tag -as [string]) `
            -DefaultEmission ($cmbDefaultEmission.SelectedItem.Tag -as [string]) `
            -Language $newLang
        Write-AppLog "Language changed to: $newLang (Google API: $script:CurrentGoogleLang)" "INFO"
    }
})

$btnOpenLangFile.Add_Click({
    if (-not (Test-Path $script:LocalizationFile)) {
        Load-LocalizationConfig
    }
    try {
        Start-Process -FilePath "notepad.exe" -ArgumentList "`"$script:LocalizationFile`""
    } catch {
        try { Start-Process -FilePath $script:LocalizationFile } catch {
            [System.Windows.MessageBox]::Show("Cannot open localization file:`r`n$($script:LocalizationFile)`r`n$($_.Exception.Message)", 'Error', 'OK', 'Error')
        }
    }
})

$btnReloadLang.Add_Click({
    Load-LocalizationConfig
    Populate-LanguageDropdowns
    Apply-AppLanguage -LanguageCode $script:CurrentLanguage
    $count = $script:LanguagesCatalog.Count
    $msg = (Get-LocText 'MsgLangReloaded') -f $count
    $title = (Get-LocText 'MsgLangReloadedTitle')
    [System.Windows.MessageBox]::Show($msg, $title, 'OK', 'Information')
})

$btnToggleKeyVisibility.Add_Click({
    if ($txtSettingsApiKeyVisible.Visibility -eq [System.Windows.Visibility]::Visible) {
        $txtSettingsApiKey.Password = $txtSettingsApiKeyVisible.Text
        $txtSettingsApiKeyVisible.Visibility = [System.Windows.Visibility]::Collapsed
        $txtSettingsApiKey.Visibility = [System.Windows.Visibility]::Visible
        $btnToggleKeyVisibility.Content = '👁 Show'
    } else {
        $txtSettingsApiKeyVisible.Text = $txtSettingsApiKey.Password
        $txtSettingsApiKey.Visibility = [System.Windows.Visibility]::Collapsed
        $txtSettingsApiKeyVisible.Visibility = [System.Windows.Visibility]::Visible
        $btnToggleKeyVisibility.Content = '🔒 Hide'
    }
})

$btnOpenLogFile.Add_Click({
    if (-not (Test-Path $script:LogFile)) {
        Write-AppLog "Creating new application log file." "INFO"
    }
    try {
        Start-Process -FilePath "notepad.exe" -ArgumentList "`"$script:LogFile`""
    } catch {
        try { Start-Process -FilePath $script:LogFile } catch {
            [System.Windows.MessageBox]::Show("Cannot open log file:`r`n$($script:LogFile)`r`n$($_.Exception.Message)", 'Log Open Error', 'OK', 'Error')
        }
    }
})

$btnTestApiKey.Add_Click({
    $key = Get-CurrentApiKey
    if ([string]::IsNullOrWhiteSpace($key)) {
        [System.Windows.MessageBox]::Show((Get-LocText 'MsgMissingApiKey'), (Get-LocText 'MsgMissingApiKeyTitle'), 'OK', 'Warning')
        return
    }

    $maskedKey = Get-MaskedKey $key
    Write-AppLog "Testing API key (Key: $maskedKey)..." "INFO"

    $lblKeyTestResult.Text = 'Testing API key...'
    $lblKeyTestResult.Foreground = [System.Windows.Media.Brushes]::SkyBlue
    $btnTestApiKey.IsEnabled = $false

    # Asynchroniczny test połączenia z Google API w osobnym runspace
    $testScript = {
        param($apiKeyToTest, $apiLanguage = 'en')
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls
        if ([string]::IsNullOrWhiteSpace($apiKeyToTest)) {
            return [PSCustomObject]@{ Valid = $false; Message = 'API key is empty.' }
        }
        try {
            $lang = if ($apiLanguage) { $apiLanguage } else { 'en' }
            $Url = "https://maps.googleapis.com/maps/api/geocode/json?address=Warszawa&language=$lang&key=$apiKeyToTest"
            $Resp = Invoke-RestMethod -Uri $Url -Method GET -TimeoutSec 15
            if ($Resp.status -eq 'OK' -or $Resp.status -eq 'ZERO_RESULTS') {
                return [PSCustomObject]@{ Valid = $true; Message = 'Google Maps API key is valid and active.' }
            }
            elseif ($Resp.status -eq 'REQUEST_DENIED') {
                $msg = if ($Resp.error_message) { $Resp.error_message } else { 'Request denied by Google API.' }
                return [PSCustomObject]@{ Valid = $false; Message = "Unauthorized: $msg" }
            }
            else {
                return [PSCustomObject]@{ Valid = $false; Message = "API Status: $($Resp.status)" }
            }
        }
        catch {
            return [PSCustomObject]@{ Valid = $false; Message = "Connection error: $($_.Exception.Message)" }
        }
    }

    $psTest = [PowerShell]::Create().AddScript($testScript).AddArgument($key).AddArgument($script:CurrentGoogleLang)
    $testHandle = $psTest.BeginInvoke()
    $testTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $testTimer.Interval = [TimeSpan]::FromMilliseconds(100)
    $script:ActiveTestTimer = $testTimer
    $script:ActiveTestPs = $psTest
    $script:ActiveTestHandle = $testHandle
    $script:TestTimerTicks = 0

    $testTimer.Add_Tick({
        $localTestHandle = $script:ActiveTestHandle
        $localTestPs     = $script:ActiveTestPs
        $script:TestTimerTicks++
        if ($localTestHandle -and $localTestHandle.IsCompleted) {
            if ($script:ActiveTestTimer) { try { $script:ActiveTestTimer.Stop() } catch { } }
            $btnTestApiKey.IsEnabled = $true
            try {
                $res = $localTestPs.EndInvoke($localTestHandle)
                $testResult = $res[0]
                $isValid = [bool]$testResult.Valid
                $msg = [string]$testResult.Message
                Update-ApiStatusBadge -IsValid $isValid -Message $msg
                $logLevel = if ($isValid) { 'OK' } else { 'WARN' }
                Write-AppLog "API key test completed: Valid=$isValid, Message='$msg'" $logLevel
            }
            catch {
                $errDetail = $_.Exception.Message
                Update-ApiStatusBadge -IsValid $false -Message "Test error: $errDetail"
                Write-AppLog "Exception while retrieving API test result: $($_.Exception.ToString())" "ERROR"
            }
            finally {
                $localTestPs.Dispose()
            }
        }
        elseif ($script:TestTimerTicks -ge 200) { # 20s timeout limit
            if ($script:ActiveTestTimer) { try { $script:ActiveTestTimer.Stop() } catch { } }
            $btnTestApiKey.IsEnabled = $true
            Update-ApiStatusBadge -IsValid $false -Message "Google API response timeout (20s)."
            Write-AppLog "API key test timeout (20s watchdog timeout)." "WARN"
            try { $localTestPs.Stop(); $localTestPs.Dispose() } catch { }
        }
    })
    $testTimer.Start()
})

$btnBrowseSettingsOutputDir.Add_Click({
    $dlg = [System.Windows.Forms.FolderBrowserDialog]::new()
    $dlg.Description = 'Select default folder for calculation results'
    $dlg.SelectedPath = $txtSettingsOutputDir.Text
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtSettingsOutputDir.Text = $dlg.SelectedPath
    }
})

$btnSaveSettings.Add_Click({
    $key = Get-CurrentApiKey
    $remember = [bool]$chkRememberKey.IsChecked
    $outDir = $txtSettingsOutputDir.Text.Trim()
    $routeType = ($cmbDefaultRouteType.SelectedItem.Tag -as [string])
    $emission = ($cmbDefaultEmission.SelectedItem.Tag -as [string])
    $dims = ($cmbDefaultMapSize.SelectedItem.Tag -as [string]) -split 'x'
    $mapW = [int]$dims[0]
    $mapH = [int]$dims[1]

    $langToSave = if ($cmbSettingsLanguage.SelectedItem) { [string]$cmbSettingsLanguage.SelectedItem.Tag } else { $script:CurrentLanguage }
    $themeToSave = if ($cmbSettingsTheme.SelectedItem) { [string]$cmbSettingsTheme.SelectedItem.Tag } else { $script:CurrentTheme }
    $overlayCfg = Get-CurrentOverlayConfig
    Save-AppConfig -ApiKey $key -RememberApiKey $remember -OutputFolder $outDir `
        -LastInputFolder $script:Config.LastInputFolder -LastInputPath $script:Config.LastInputPath `
        -DefaultRouteType $routeType -DefaultEmission $emission -MapWidth $mapW -MapHeight $mapH -Language $langToSave `
        -OverlayConfig $overlayCfg -Theme $themeToSave
    $script:Config.OverlayConfig = $overlayCfg
    $script:Config.Theme = $themeToSave

    Set-CurrentApiKey -Key $key
    if ($remember -and -not [string]::IsNullOrWhiteSpace($key)) {
        $lblApiBadge.Text = 'API: Configured'
        $lblApiBadge.Foreground = [System.Windows.Media.Brushes]::LightGreen
    }
    elseif (-not $remember -and -not [string]::IsNullOrWhiteSpace($key)) {
        $lblApiBadge.Text = 'API: Session key (unsaved)'
        $lblApiBadge.Foreground = [System.Windows.Media.Brushes]::SkyBlue
    }
    else {
        $lblApiBadge.Text = 'API: Unverified'
        $lblApiBadge.Foreground = [System.Windows.Media.Brushes]::Orange
    }

    [System.Windows.MessageBox]::Show((Get-LocText 'MsgSettingsSaved'), (Get-LocText 'MsgSettingsSavedTitle'), 'OK', 'Information')
})

if ($btnResetOverlayConfig) {
    $btnResetOverlayConfig.Add_Click({
        Reset-OverlayConfigUi
    })
}

# ── 12. Zdarzenia: Tab 1 (Manual Input) ───────────────────────────────────────
$btnClearManualStart.Add_Click({ $txtManualStart.Clear() })
$btnClearManualEnd.Add_Click({ $txtManualEnd.Clear() })

$btnAddWaypoint.Add_Click({
    $wp = $txtNewWaypoint.Text.Trim()
    if (-not [string]::IsNullOrWhiteSpace($wp)) {
        if ($lstWaypoints.Items.Count -ge 25) {
            [System.Windows.MessageBox]::Show((Get-LocText 'MsgMaxWaypoints'), (Get-LocText 'MsgMaxWaypointsTitle'), 'OK', 'Warning')
            return
        }
        $null = $lstWaypoints.Items.Add($wp)
        $txtNewWaypoint.Clear()
    }
})

$txtNewWaypoint.Add_KeyDown({
    if ($_.Key -eq [System.Windows.Input.Key]::Enter) {
        $btnAddWaypoint.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
    }
})

$btnWpRemove.Add_Click({
    if ($lstWaypoints.SelectedIndex -ge 0) {
        $lstWaypoints.Items.RemoveAt($lstWaypoints.SelectedIndex)
    }
})

$btnWpClear.Add_Click({
    $lstWaypoints.Items.Clear()
})

$btnWpUp.Add_Click({
    $idx = $lstWaypoints.SelectedIndex
    if ($idx -gt 0) {
        $item = $lstWaypoints.Items[$idx]
        $lstWaypoints.Items.RemoveAt($idx)
        $lstWaypoints.Items.Insert($idx - 1, $item)
        $lstWaypoints.SelectedIndex = $idx - 1
    }
})

$btnWpDown.Add_Click({
    $idx = $lstWaypoints.SelectedIndex
    if ($idx -ge 0 -and $idx -lt ($lstWaypoints.Items.Count - 1)) {
        $item = $lstWaypoints.Items[$idx]
        $lstWaypoints.Items.RemoveAt($idx)
        $lstWaypoints.Items.Insert($idx + 1, $item)
        $lstWaypoints.SelectedIndex = $idx + 1
    }
})

$rbTypeEco.Add_Checked({ $pnlEmission.Visibility = [System.Windows.Visibility]::Visible })
$rbTypeFastest.Add_Checked({ $pnlEmission.Visibility = [System.Windows.Visibility]::Collapsed })
$rbTypeShortest.Add_Checked({ $pnlEmission.Visibility = [System.Windows.Visibility]::Collapsed })

$btnCalculateManual.Add_Click({
    $apiKey = Get-CurrentApiKey
    if ([string]::IsNullOrWhiteSpace($apiKey)) {
        [System.Windows.MessageBox]::Show((Get-LocText 'MsgMissingApiKeyPrompt'), (Get-LocText 'MsgMissingApiKeyTitle'), 'OK', 'Warning')
        $tabMain.SelectedIndex = 2
        return
    }

    $start = $txtManualStart.Text.Trim()
    $end = $txtManualEnd.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($start) -or [string]::IsNullOrWhiteSpace($end)) {
        [System.Windows.MessageBox]::Show((Get-LocText 'MsgMissingData'), (Get-LocText 'MsgMissingDataTitle'), 'OK', 'Warning')
        return
    }

    $waypoints = [System.Collections.Generic.List[string]]::new()
    foreach ($item in $lstWaypoints.Items) {
        $waypoints.Add([string]$item)
    }

    $routeType = if ($rbTypeShortest.IsChecked) { 'Shortest' } elseif ($rbTypeEco.IsChecked) { 'Eco' } else { 'Fastest' }
    $emission = ($cmbEmission.SelectedItem.Tag -as [string])
    if ([string]::IsNullOrWhiteSpace($emission)) { $emission = 'GASOLINE' }
    $trafficAware = [bool]$chkTrafficAware.IsChecked
    $name = $txtManualName.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($name)) { $name = "Route $start -> $end" }

    $outDir = $txtSettingsOutputDir.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($outDir)) { $outDir = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'GoogleMapsRoutes\Results' }
    if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

    if ($txtBatchFilePath.Text -and (Test-Path $txtBatchFilePath.Text.Trim())) {
        $script:Config.LastInputPath = $txtBatchFilePath.Text.Trim()
        $script:Config.LastInputFolder = Split-Path $script:Config.LastInputPath -Parent
        $script:LastDataDirectory = $script:Config.LastInputFolder
        Save-AppConfig -ApiKey $apiKey -RememberApiKey $chkRememberKey.IsChecked -OutputFolder $outDir `
            -LastInputFolder $script:Config.LastInputFolder -LastInputPath $script:Config.LastInputPath `
            -DefaultRouteType ($cmbDefaultRouteType.SelectedItem.Tag -as [string]) `
            -DefaultEmission ($cmbDefaultEmission.SelectedItem.Tag -as [string])
    }

    $btnCalculateManual.IsEnabled = $false
    $btnCalculateManual.Content = '⏳ CALCULATING ROUTE...'
    $lblManualStatus.Text = 'Geocoding and calculating...'
    $lblManualStatus.Foreground = [System.Windows.Media.Brushes]::SkyBlue
    $lblFooterStatus.Text = 'Calculating manual route...'

    Write-AppLog "Started manual route calculation: '$start' -> '$end' (Waypoints: $($waypoints.Count), Type: $routeType, Engine: $emission, LiveTraffic: $trafficAware)..." "INFO"

    # $script:ManualCalcAsync is defined at top-level (section 6b) — used directly below
    $psCmd = New-WorkerPowerShell -ScriptBlock $script:ManualCalcAsync
    $overlayCfgJson = (Get-CurrentOverlayConfig | ConvertTo-Json -Depth 6 -Compress)
    $psCmd.AddArgument($start).AddArgument($end).AddArgument($waypoints).AddArgument($routeType).AddArgument($emission).AddArgument($trafficAware).AddArgument($name).AddArgument($apiKey).AddArgument($outDir).AddArgument($script:LogFile).AddArgument($script:CurrentGoogleLang).AddArgument($overlayCfgJson) | Out-Null

    try {
        $asyncHandle = $psCmd.BeginInvoke()
    } catch {
        Write-AppLog "CRITICAL: BeginInvoke() threw exception: $($_.Exception.Message)" "ERROR"
        $btnCalculateManual.IsEnabled = $true
        $btnCalculateManual.Content = '🚀 CALCULATE ROUTE & DOWNLOAD MAP'
        $lblManualStatus.Text = '✕ Launch error'
        $lblManualStatus.Foreground = [System.Windows.Media.Brushes]::Salmon
        $lblFooterStatus.Text = "Error: $($_.Exception.Message)"
        return
    }
    if (-not $asyncHandle) {
        Write-AppLog "CRITICAL: BeginInvoke() returned null — runspace may be invalid." "ERROR"
        $btnCalculateManual.IsEnabled = $true
        $btnCalculateManual.Content = '🚀 CALCULATE ROUTE & DOWNLOAD MAP'
        $lblManualStatus.Text = '✕ Runspace error'
        $lblManualStatus.Foreground = [System.Windows.Media.Brushes]::Salmon
        return
    }
    $timer = [System.Windows.Threading.DispatcherTimer]::new()
    $timer.Interval = [TimeSpan]::FromMilliseconds(150)
    $script:ActiveManualTimer = $timer
    $script:ActiveManualPs = $psCmd
    $script:ActiveManualAsyncHandle = $asyncHandle
    $script:ManualTimerTicks = 0

    Write-AppLog "Worker started (BeginInvoke). IsCompleted=$($asyncHandle.IsCompleted)" "INFO"

    $timer.Add_Tick({
        $localHandle = $script:ActiveManualAsyncHandle
        $localPs     = $script:ActiveManualPs
        $script:ManualTimerTicks++
        if ($localHandle -and $localHandle.IsCompleted) {
            if ($script:ActiveManualTimer) { try { $script:ActiveManualTimer.Stop() } catch { } }
            $btnCalculateManual.IsEnabled = $true
            $btnCalculateManual.Content = '🚀 CALCULATE ROUTE & DOWNLOAD MAP'

            # Log any stream errors from the worker runspace before inspecting result
            foreach ($streamErr in $localPs.Streams.Error) {
                Write-AppLog "[Stream.Error] $($streamErr.Exception.Message) @ $($streamErr.InvocationInfo.PositionMessage)" "ERROR"
            }

            try {
                $res = $localPs.EndInvoke($localHandle)
                $calc = $res[0]
                if ($calc.Success) {
                    $lblManualDist.Text = "$($calc.DistanceKm) km"
                    $lblManualTime.Text = "$($calc.DurationMin) min"
                    $lblManualType.Text = switch ($script:CurrentLanguage) {
                        'de' { if ($calc.RouteType -eq 'Fastest') { 'Schnellste' } elseif ($calc.RouteType -eq 'Shortest') { 'Kürzeste' } else { 'Eco' } }
                        'pl' { if ($calc.RouteType -eq 'Fastest') { 'Najszybsza' } elseif ($calc.RouteType -eq 'Shortest') { 'Najkrótsza' } else { 'Eko' } }
                        default { [string]$calc.RouteType }
                    }
                    $lblManualStatus.Text = '✓ Success'
                    $lblManualStatus.Foreground = [System.Windows.Media.Brushes]::LightGreen
                    $lblFooterStatus.Text = "Route ready: $($calc.DistanceKm) km, $($calc.DurationMin) min"
                    Write-AppLog "Manual route calculation completed successfully: $($calc.DistanceKm) km, $($calc.DurationMin) min (Map file: $($calc.MapPath))" "OK"

                    $script:LastGoogleMapsUrl = $calc.GoogleMapsUrl
                    $lblGoogleUrlDisplay.Text = $calc.GoogleMapsUrl
                    $btnOpenGoogleMaps.IsEnabled = $true
                    $btnCopyUrl.IsEnabled = $true

                    if ($calc.MapPath -and (Test-Path $calc.MapPath)) {
                        $script:LastGeneratedMapPath = $calc.MapPath
                        $btnSaveMapAs.IsEnabled = $true
                        $lblMapPlaceholder.Visibility = [System.Windows.Visibility]::Collapsed

                        $imgBytes = [System.IO.File]::ReadAllBytes($calc.MapPath)
                        $ms = [System.IO.MemoryStream]::new($imgBytes)
                        $bi = [System.Windows.Media.Imaging.BitmapImage]::new()
                        $bi.BeginInit()
                        $bi.StreamSource = $ms
                        $bi.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
                        $bi.EndInit()
                        $bi.Freeze()
                        $imgMapPreview.Source = $bi
                    }
                } else {
                    $lblManualStatus.Text = '✕ Error'
                    $lblManualStatus.Foreground = [System.Windows.Media.Brushes]::Salmon
                    $lblFooterStatus.Text = "Error: $($calc.Error)"
                    Write-AppLog "Manual route calculation failed: $($calc.Error)" "ERROR"
                    [System.Windows.MessageBox]::Show($calc.Error, 'Route Error', 'OK', 'Error')
                }
            }
            catch {
                $errDetail = $_.Exception.ToString()
                $lblManualStatus.Text = '✕ Error'
                $lblManualStatus.Foreground = [System.Windows.Media.Brushes]::Salmon
                $lblFooterStatus.Text = "Exception: $($_.Exception.Message)"
                Write-AppLog "UI exception while reading route result: $errDetail" "ERROR"
                [System.Windows.MessageBox]::Show($_.Exception.Message, 'Error', 'OK', 'Error')
            }
            finally {
                $localPs.Dispose()
            }
        }
        elseif ($script:ManualTimerTicks -ge 400) { # 60 seconds watchdog timeout
            if ($script:ActiveManualTimer) { try { $script:ActiveManualTimer.Stop() } catch { } }
            $btnCalculateManual.IsEnabled = $true
            $btnCalculateManual.Content = '🚀 CALCULATE ROUTE & DOWNLOAD MAP'
            $lblManualStatus.Text = '✕ Timeout (60s)'
            $lblManualStatus.Foreground = [System.Windows.Media.Brushes]::Salmon
            $lblFooterStatus.Text = 'Route calculation timed out (60s).'
            Write-AppLog "Manual route calculation timed out (60s watchdog timeout). Ticks=$($script:ManualTimerTicks)" "WARN"
            try { $localPs.Stop(); $localPs.Dispose() } catch { }
        }
    })
    $timer.Start()
})

$btnOpenGoogleMaps.Add_Click({
    if ($script:LastGoogleMapsUrl) {
        Start-Process $script:LastGoogleMapsUrl
    }
})

$btnCopyUrl.Add_Click({
    if ($script:LastGoogleMapsUrl) {
        [System.Windows.Clipboard]::SetText($script:LastGoogleMapsUrl)
        [System.Windows.MessageBox]::Show((Get-LocText 'MsgUrlCopied'), (Get-LocText 'MsgUrlCopiedTitle'), 'OK', 'Information')
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
            [System.Windows.MessageBox]::Show(((Get-LocText 'MsgMapSaved') -f $dlg.FileName), (Get-LocText 'MsgMapSavedTitle'), 'OK', 'Information')
        }
    }
})

# ── 13. Zdarzenia: Tab 2 (Batch Processing) ──────────────────────────────────
function Load-BatchFilePreview([string]$Path) {
    if (-not (Test-Path $Path)) { return }
    try {
        $data = Import-RouteDataFile -Path $Path
        $script:LoadedBatchData = $data

        # Configure dynamic columns and clear previous items
        $dgBatchInput.ItemsSource = $null
        $dgBatchInput.Columns.Clear()

        if ($data.Mode -eq 'SequentialStops') {
            $lblBatchFileInfo.Text = "Format: $($data.Format) | Mode: Sequential Stops (1 Multi-point Route, $($data.TotalCount) Stops) | Total: $($data.TotalCount) stops"
            $lblBatchFileInfo.Foreground = [System.Windows.Media.Brushes]::LightGreen

            $colStep = [System.Windows.Controls.DataGridTextColumn]::new()
            $colStep.Header = "#"
            $colStep.Binding = [System.Windows.Data.Binding]::new("Step")
            $colStep.Width = [System.Windows.Controls.DataGridLength]::new(55)
            $dgBatchInput.Columns.Add($colStep)

            $colRole = [System.Windows.Controls.DataGridTextColumn]::new()
            $colRole.Header = "Role / Point Type"
            $colRole.Binding = [System.Windows.Data.Binding]::new("Role")
            $colRole.Width = [System.Windows.Controls.DataGridLength]::new(180)
            $dgBatchInput.Columns.Add($colRole)

            $colAddr = [System.Windows.Controls.DataGridTextColumn]::new()
            $colAddr.Header = "Address / Location"
            $colAddr.Binding = [System.Windows.Data.Binding]::new("Address")
            $colAddr.Width = [System.Windows.Controls.DataGridLength]::new(340)
            $dgBatchInput.Columns.Add($colAddr)

            $colRaw = [System.Windows.Controls.DataGridTextColumn]::new()
            $colRaw.Header = "Source Record Data"
            $colRaw.Binding = [System.Windows.Data.Binding]::new("RawSummary")
            $colRaw.Width = [System.Windows.Controls.DataGridLength]::new(1, [System.Windows.Controls.DataGridLengthUnitType]::Star)
            $dgBatchInput.Columns.Add($colRaw)

            $previewItems = [System.Collections.Generic.List[PSCustomObject]]::new()
            $stopsCount = $data.Stops.Count
            for ($i = 0; $i -lt $stopsCount; $i++) {
                $st = $data.Stops[$i]
                $role = if ($i -eq 0) { "🟢 Origin (Start)" }
                        elseif ($i -eq ($stopsCount - 1)) { "🔴 Destination (End)" }
                        else { "🟡 Waypoint $i" }

                $rawProps = @()
                if ($st.Raw) {
                    foreach ($p in $st.Raw.PSObject.Properties) {
                        $rawProps += "$($p.Name)=$($p.Value)"
                    }
                }
                $rawSummaryText = $rawProps -join '; '

                $previewItems.Add([PSCustomObject]@{
                    Step       = ($i + 1)
                    Role       = $role
                    Address    = [string]$st.Address
                    RawSummary = $rawSummaryText
                })
            }
            $dgBatchInput.ItemsSource = $previewItems
            Write-BatchLog "Loaded file: $Path ($($data.TotalCount) sequential stops, format: $($data.Format), mode: $($data.Mode))" "OK"
        }
        else {
            $lblBatchFileInfo.Text = "Format: $($data.Format) | Mode: Route List | Total Routes: $($data.TotalCount)"
            $lblBatchFileInfo.Foreground = [System.Windows.Media.Brushes]::LightGreen

            $colId = [System.Windows.Controls.DataGridTextColumn]::new()
            $colId.Header = "ID"
            $colId.Binding = [System.Windows.Data.Binding]::new("Id")
            $colId.Width = [System.Windows.Controls.DataGridLength]::new(50)
            $dgBatchInput.Columns.Add($colId)

            $colName = [System.Windows.Controls.DataGridTextColumn]::new()
            $colName.Header = "Route Name"
            $colName.Binding = [System.Windows.Data.Binding]::new("Name")
            $colName.Width = [System.Windows.Controls.DataGridLength]::new(180)
            $dgBatchInput.Columns.Add($colName)

            $colStart = [System.Windows.Controls.DataGridTextColumn]::new()
            $colStart.Header = "Origin (Start)"
            $colStart.Binding = [System.Windows.Data.Binding]::new("Start")
            $colStart.Width = [System.Windows.Controls.DataGridLength]::new(200)
            $dgBatchInput.Columns.Add($colStart)

            $colEnd = [System.Windows.Controls.DataGridTextColumn]::new()
            $colEnd.Header = "Destination (End)"
            $colEnd.Binding = [System.Windows.Data.Binding]::new("End")
            $colEnd.Width = [System.Windows.Controls.DataGridLength]::new(200)
            $dgBatchInput.Columns.Add($colEnd)

            $colWpCount = [System.Windows.Controls.DataGridTextColumn]::new()
            $colWpCount.Header = "Waypoints"
            $colWpCount.Binding = [System.Windows.Data.Binding]::new("WaypointCount")
            $colWpCount.Width = [System.Windows.Controls.DataGridLength]::new(80)
            $dgBatchInput.Columns.Add($colWpCount)

            $colWpText = [System.Windows.Controls.DataGridTextColumn]::new()
            $colWpText.Header = "Intermediate Stops"
            $colWpText.Binding = [System.Windows.Data.Binding]::new("WaypointsText")
            $colWpText.Width = [System.Windows.Controls.DataGridLength]::new(250)
            $dgBatchInput.Columns.Add($colWpText)

            $colType = [System.Windows.Controls.DataGridTextColumn]::new()
            $colType.Header = "Route Type"
            $colType.Binding = [System.Windows.Data.Binding]::new("RouteType")
            $colType.Width = [System.Windows.Controls.DataGridLength]::new(100)
            $dgBatchInput.Columns.Add($colType)

            $previewItems = [System.Collections.Generic.List[PSCustomObject]]::new()
            foreach ($r in @($data.Routes)) {
                $wpText = if ($r.Waypoints -and @($r.Waypoints).Count -gt 0) {
                    (@($r.Waypoints) -join ' | ')
                } else {
                    '(none)'
                }
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
            Write-BatchLog "Loaded file: $Path ($($data.TotalCount) routes, format: $($data.Format), mode: $($data.Mode))" "OK"
        }
    }
    catch {
        $lblBatchFileInfo.Text = "Load error: $($_.Exception.Message)"
        $lblBatchFileInfo.Foreground = [System.Windows.Media.Brushes]::Salmon
        Write-BatchLog "Load error: $($_.Exception.Message)" "ERROR"
    }
}

$btnBrowseBatchFile.Add_Click({
    $initDir = $null
    if (-not [string]::IsNullOrWhiteSpace($txtBatchFilePath.Text) -and (Test-Path $txtBatchFilePath.Text.Trim())) {
        $initDir = Split-Path $txtBatchFilePath.Text.Trim() -Parent
    }
    elseif (-not [string]::IsNullOrWhiteSpace($txtBatchFilePath.Text) -and (Test-Path (Split-Path $txtBatchFilePath.Text.Trim() -Parent))) {
        $initDir = Split-Path $txtBatchFilePath.Text.Trim() -Parent
    }
    elseif ($script:Config.LastInputFolder -and (Test-Path $script:Config.LastInputFolder)) {
        $initDir = $script:Config.LastInputFolder
    }
    elseif ($script:Config.LastInputPath -and (Test-Path (Split-Path $script:Config.LastInputPath -Parent))) {
        $initDir = Split-Path $script:Config.LastInputPath -Parent
    }

    $file = Select-InputDataFile -InitialDirectory $initDir
    if ($file) {
        $txtBatchFilePath.Text = $file
        $script:Config.LastInputPath = $file
        $script:Config.LastInputFolder = Split-Path $file -Parent
        $script:LastDataDirectory = $script:Config.LastInputFolder

        Save-AppConfig -ApiKey (Get-CurrentApiKey) `
            -RememberApiKey $chkRememberKey.IsChecked `
            -OutputFolder $txtSettingsOutputDir.Text.Trim() `
            -LastInputFolder $script:Config.LastInputFolder `
            -LastInputPath $script:Config.LastInputPath `
            -DefaultRouteType ($cmbDefaultRouteType.SelectedItem.Tag -as [string]) `
            -DefaultEmission ($cmbDefaultEmission.SelectedItem.Tag -as [string])

        Load-BatchFilePreview -Path $file
    }
})

$btnReloadBatchFile.Add_Click({
    if ($txtBatchFilePath.Text) {
        $path = $txtBatchFilePath.Text.Trim()
        if (Test-Path $path) {
            $script:Config.LastInputPath = $path
            $script:Config.LastInputFolder = Split-Path $path -Parent
            $script:LastDataDirectory = $script:Config.LastInputFolder

            Save-AppConfig -ApiKey (Get-CurrentApiKey) `
                -RememberApiKey $chkRememberKey.IsChecked `
                -OutputFolder $txtSettingsOutputDir.Text.Trim() `
                -LastInputFolder $script:Config.LastInputFolder `
                -LastInputPath $script:Config.LastInputPath `
                -DefaultRouteType ($cmbDefaultRouteType.SelectedItem.Tag -as [string]) `
                -DefaultEmission ($cmbDefaultEmission.SelectedItem.Tag -as [string])
        }
        Load-BatchFilePreview -Path $path
    }
})

$btnStartBatch.Add_Click({
    $apiKey = Get-CurrentApiKey
    if ([string]::IsNullOrWhiteSpace($apiKey)) {
        [System.Windows.MessageBox]::Show((Get-LocText 'MsgMissingApiKeyPrompt'), (Get-LocText 'MsgMissingApiKeyTitle'), 'OK', 'Warning')
        $tabMain.SelectedIndex = 2
        return
    }

    if ($null -eq $script:LoadedBatchData -or $script:LoadedBatchData.Routes.Count -eq 0) {
        [System.Windows.MessageBox]::Show((Get-LocText 'MsgNoDataFile'), (Get-LocText 'MsgNoDataFileTitle'), 'OK', 'Warning')
        return
    }

    $outDir = $txtSettingsOutputDir.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($outDir)) { $outDir = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'GoogleMapsRoutes\Results' }
    if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

    if ($txtBatchFilePath.Text -and (Test-Path $txtBatchFilePath.Text.Trim())) {
        $script:Config.LastInputPath = $txtBatchFilePath.Text.Trim()
        $script:Config.LastInputFolder = Split-Path $script:Config.LastInputPath -Parent
        $script:LastDataDirectory = $script:Config.LastInputFolder
        Save-AppConfig -ApiKey $apiKey -RememberApiKey $chkRememberKey.IsChecked -OutputFolder $outDir `
            -LastInputFolder $script:Config.LastInputFolder -LastInputPath $script:Config.LastInputPath `
            -DefaultRouteType ($cmbDefaultRouteType.SelectedItem.Tag -as [string]) `
            -DefaultEmission ($cmbDefaultEmission.SelectedItem.Tag -as [string])
    }

    $defaultRouteType = ($cmbBatchRouteType.SelectedItem.Tag -as [string])
    if ([string]::IsNullOrWhiteSpace($defaultRouteType)) { $defaultRouteType = 'Fastest' }

    $script:BatchWorkerRunning = $true
    $btnStartBatch.IsEnabled = $false
    $btnStopBatch.IsEnabled = $true
    $btnBrowseBatchFile.IsEnabled = $false
    $btnReloadBatchFile.IsEnabled = $false

    $script:BatchResultsList.Clear()
    $dgBatchResults.ItemsSource = $null
    if ($dgBatchPoints) { $dgBatchPoints.ItemsSource = $null }
    $pbBatchProgress.Value = 0
    $lblBatchProgressText.Text = "Starting batch processing (0 / $($script:LoadedBatchData.Routes.Count))..."
    $lblBatchStats.Text = "Success: 0 | Errors: 0"

    Write-BatchLog "=== Starting batch processing ($($script:LoadedBatchData.Routes.Count) routes) ===" "INFO"

    # Switch to Activity Log tab so user sees live execution
    if ($tabBatchSub) {
        $tabBatchSub.SelectedIndex = 2
    }

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
        $overlayCfgJson = (Get-CurrentOverlayConfig | ConvertTo-Json -Depth 6 -Compress)
        $psCmdBatch.AddArgument($routesToProcess).AddArgument($apiKey).AddArgument($outDir).AddArgument($defaultRouteType).AddArgument($syncState).AddArgument($script:LogFile).AddArgument($script:CurrentGoogleLang).AddArgument($overlayCfgJson) | Out-Null
        $asyncBatchHandle = $psCmdBatch.BeginInvoke()
    }
    catch {
        Write-BatchLog "CRITICAL: Could not start batch worker: $($_.Exception.Message)" "ERROR"
        $btnStartBatch.IsEnabled = $true
        $btnStopBatch.IsEnabled = $false
        $btnBrowseBatchFile.IsEnabled = $true
        $btnReloadBatchFile.IsEnabled = $true
        $script:BatchWorkerRunning = $false
        $lblBatchProgressText.Text = "Launch error"
        return
    }

    if (-not $asyncBatchHandle) {
        Write-BatchLog "CRITICAL: BeginInvoke returned null handle." "ERROR"
        $btnStartBatch.IsEnabled = $true
        $btnStopBatch.IsEnabled = $false
        $btnBrowseBatchFile.IsEnabled = $true
        $btnReloadBatchFile.IsEnabled = $true
        $script:BatchWorkerRunning = $false
        return
    }

    $timerBatch = [System.Windows.Threading.DispatcherTimer]::new()
    $timerBatch.Interval = [TimeSpan]::FromMilliseconds(150)
    $script:ActiveBatchTimer = $timerBatch
    $script:ActiveBatchPs = $psCmdBatch
    $script:ActiveBatchAsyncHandle = $asyncBatchHandle

    # Direct script-scope reference (NO .GetNewClosure()!)
    $timerBatch.Add_Tick({
        $localBatchHandle = $script:ActiveBatchAsyncHandle
        $localBatchPs     = $script:ActiveBatchPs
        $localSyncState   = $script:SyncState

        if (-not $localSyncState) { return }

        # Flush real-time worker logs to UI Activity Log
        if ($localSyncState.LogQueue) {
            $logItem = $null
            while ($localSyncState.LogQueue.TryDequeue([ref]$logItem)) {
                if ($logItem) {
                    Write-BatchLog $logItem.Message $logItem.Level
                }
            }
        }

        $curr = $localSyncState.CurrentIndex
        $tot  = $localSyncState.TotalCount
        $pct  = if ($tot -gt 0) { [math]::Min(100, [math]::Round(($curr / $tot) * 100, 0)) } else { 0 }
        $pbBatchProgress.Value = $pct
        $lblBatchProgressText.Text = "Processing: $curr / $tot ($pct%)"
        $lblBatchStats.Text = "Success: $($localSyncState.SuccessCount) | Errors: $($localSyncState.FailCount)"

        if ($localBatchHandle -and $localBatchHandle.IsCompleted) {
            $script:ActiveBatchTimer.Stop()
            $btnStartBatch.IsEnabled = $true
            $btnStopBatch.IsEnabled = $false
            $btnBrowseBatchFile.IsEnabled = $true
            $btnReloadBatchFile.IsEnabled = $true
            $script:BatchWorkerRunning = $false

            # Flush any remaining logs
            if ($localSyncState.LogQueue) {
                $logItem = $null
                while ($localSyncState.LogQueue.TryDequeue([ref]$logItem)) {
                    if ($logItem) {
                        Write-BatchLog $logItem.Message $logItem.Level
                    }
                }
            }

            # Check stream errors
            foreach ($streamErr in $localBatchPs.Streams.Error) {
                Write-BatchLog "[Worker Stream Error] $($streamErr.Exception.Message)" "ERROR"
            }

            try {
                $res = $localBatchPs.EndInvoke($localBatchHandle)
                $script:BatchResultsList = [System.Collections.Generic.List[PSCustomObject]]::new()
                foreach ($item in @($res)) { $script:BatchResultsList.Add($item) }
                $dgBatchResults.ItemsSource = $script:BatchResultsList

                # Populate Points Detail table
                $allPointsList = [System.Collections.Generic.List[PSCustomObject]]::new()
                foreach ($item in @($script:BatchResultsList)) {
                    if ($item.Points -and @($item.Points).Count -gt 0) {
                        foreach ($pt in @($item.Points)) {
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
                if ($dgBatchPoints) {
                    $dgBatchPoints.ItemsSource = $allPointsList
                }

                $statusMsg = if ($localSyncState.CancelRequested) { 'Stopped by user.' } else { 'Completed successfully.' }
                $lblBatchProgressText.Text = $statusMsg
                Write-BatchLog "=== Batch processing completed. $statusMsg Success: $($localSyncState.SuccessCount), Errors: $($localSyncState.FailCount) ===" "OK"
                $lblFooterStatus.Text = "Processing complete: $($localSyncState.SuccessCount) routes generated."

                # Automatically switch to Calculation Results tab
                if ($script:BatchResultsList.Count -gt 0 -and $tabBatchSub) {
                    $tabBatchSub.SelectedIndex = 1
                }
            }
            catch {
                Write-BatchLog "Error reading batch results: $($_.Exception.Message)" "ERROR"
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
        $lblBatchProgressText.Text = 'Stopping...'
        Write-BatchLog "Stop requested by user..." "WARN"
    }
})

$dgBatchResults.Add_MouseDoubleClick({
    $sel = $dgBatchResults.SelectedItem
    if ($sel -and $sel.MapPath -and (Test-Path $sel.MapPath)) {
        Start-Process $sel.MapPath
    }
})

$btnOpenOutputDir.Add_Click({
    $outDir = $txtSettingsOutputDir.Text.Trim()
    if (Test-Path $outDir) {
        Start-Process explorer.exe -ArgumentList "`"$outDir`""
    }
})

$btnExportExcel.Add_Click({
    if ($script:BatchResultsList.Count -eq 0) {
        [System.Windows.MessageBox]::Show((Get-LocText 'MsgNoExportResults'), (Get-LocText 'MsgNoExportResultsTitle'), 'OK', 'Information')
        return
    }
    $outDir = $txtSettingsOutputDir.Text.Trim()
    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $path = Join-Path $outDir "${ts}_route_results.xlsx"
    $saved = Export-RouteResults -Results $script:BatchResultsList -OutputPath $path -Format Excel
    [System.Windows.MessageBox]::Show(((Get-LocText 'MsgExportExcelComplete') -f $saved), (Get-LocText 'MsgExportTitle'), 'OK', 'Information')
})

$btnExportCsv.Add_Click({
    if ($script:BatchResultsList.Count -eq 0) {
        [System.Windows.MessageBox]::Show((Get-LocText 'MsgNoExportResults'), (Get-LocText 'MsgNoExportResultsTitle'), 'OK', 'Information')
        return
    }
    $outDir = $txtSettingsOutputDir.Text.Trim()
    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $path = Join-Path $outDir "${ts}_route_results.csv"
    $saved = Export-RouteResults -Results $script:BatchResultsList -OutputPath $path -Format CSV
    [System.Windows.MessageBox]::Show(((Get-LocText 'MsgExportCsvComplete') -f $saved), (Get-LocText 'MsgExportTitle'), 'OK', 'Information')
})

$btnExportJson.Add_Click({
    if ($script:BatchResultsList.Count -eq 0) {
        [System.Windows.MessageBox]::Show((Get-LocText 'MsgNoExportResults'), (Get-LocText 'MsgNoExportResultsTitle'), 'OK', 'Information')
        return
    }
    $outDir = $txtSettingsOutputDir.Text.Trim()
    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $path = Join-Path $outDir "${ts}_route_results.json"
    $saved = Export-RouteResults -Results $script:BatchResultsList -OutputPath $path -Format JSON
    [System.Windows.MessageBox]::Show(((Get-LocText 'MsgExportJsonComplete') -f $saved), (Get-LocText 'MsgExportTitle'), 'OK', 'Information')
})

# ── 13. Obsługa zamykania okna ───────────────────────────────────────────────
$window.Add_Closing({
    if ($script:ActiveBatchTimer) { try { $script:ActiveBatchTimer.Stop() } catch { } }
    if ($script:ActiveManualTimer) { try { $script:ActiveManualTimer.Stop() } catch { } }
    if ($script:SyncState) { $script:SyncState.CancelRequested = $true }
})

# ── 14. Uruchomienie okna ────────────────────────────────────────────────────
$window.ShowDialog() | Out-Null
.Exception.Message)" "WARN"
        }
    }

    # 2. Auto-sync: If master localization.json exists in ExeDir or PSScriptRoot and is newer than active file, copy it
    $masterLoc = if ($script:ExeDir -and (Test-Path (Join-Path $script:ExeDir 'localization.json'))) {
        Join-Path $script:ExeDir 'localization.json'
    } elseif ($PSScriptRoot -and (Test-Path (Join-Path $PSScriptRoot 'localization.json'))) {
        Join-Path $PSScriptRoot 'localization.json'
    } else { $null }

    if ($masterLoc -and $script:LocalizationFile -ne $masterLoc -and (Test-Path $masterLoc)) {
        try {
            $masterTime = (Get-Item $masterLoc).LastWriteTimeUtc
            $activeTime = if (Test-Path $script:LocalizationFile) { (Get-Item $script:LocalizationFile).LastWriteTimeUtc } else { [DateTime]::MinValue }
            if ($masterTime -gt $activeTime) {
                Copy-Item -Path $masterLoc -Destination $script:LocalizationFile -Force
            }
        } catch { }
    }

    # 3. If file does not exist on disk, save embedded template so user can customize it locally
    if (-not (Test-Path $script:LocalizationFile)) {
        try {
            $dir = Split-Path -Parent $script:LocalizationFile
            if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
            if (-not [string]::IsNullOrWhiteSpace($script:EmbeddedLocalizationJson)) {
                [System.IO.File]::WriteAllText($script:LocalizationFile, $script:EmbeddedLocalizationJson, [System.Text.UTF8Encoding]::new($true))
            }
        } catch { }
    } else {
        # 4. Read and overlay external localization file (allows user overrides / custom keys)
        try {
            $raw = [System.IO.File]::ReadAllText($script:LocalizationFile, [System.Text.Encoding]::UTF8)
            $parsed = $raw | ConvertFrom-Json
            if ($parsed.Languages) {
                foreach ($prop in $parsed.Languages.PSObject.Properties) {
                    $code = $prop.Name.ToLower()
                    $langData = $prop.Value
                    $disp = if ($langData.DisplayName) { [string]$langData.DisplayName } else { $code.ToUpper() }
                    $gCode = if ($langData.GoogleCode) { [string]$langData.GoogleCode } else { $code }

                    if (-not $script:LanguagesCatalog.Contains($code)) {
                        $script:LanguagesCatalog[$code] = [PSCustomObject]@{
                            Code        = $code
                            DisplayName = $disp
                            GoogleCode  = $gCode
                            Strings     = @{}
                        }
                    } else {
                        $script:LanguagesCatalog[$code].DisplayName = $disp
                        $script:LanguagesCatalog[$code].GoogleCode  = $gCode
                    }

                    if ($langData.Strings) {
                        foreach ($sProp in $langData.Strings.PSObject.Properties) {
                            $script:LanguagesCatalog[$code].Strings[$sProp.Name] = [string]$sProp.Value
                        }
                    }
                }
            }
        } catch {
            Write-AppLog "Error parsing localization file $script:LocalizationFile : $(#Requires -Version 5.1
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force -ErrorAction SilentlyContinue
<#
.SYNOPSIS
    Google Maps Route & Map Generator — Zaawansowana aplikacja WPF Dark Mode.
    Obsługuje ręczne wprowadzanie tras (Start, Cel, Punkty pośrednie, Fastest/Shortest/Eco)
    oraz wsadowe przetwarzanie plików danych (JSON, CSV, Excel).

.DESCRIPTION
    Funkcje:
      - Tryb ręczny (Manual Input):
          * Punkt startowy i punkt końcowy (z walidacją i geokodowaniem)
          * Dynamiczna lista punktów pośrednich (dodawanie, usuwanie, zmiana kolejności)
          * Wybór optymalizacji: Najszybsza (Fastest), Najkrótsza (Shortest), Ekologiczna (Eco)
          * Wybór typu napędu dla trasy Eco (Benzyna, Diesel, Hybryda, Elektryczny)
          * Natychmiastowe obliczanie trasy, odległości (km) i czasu (min/godz)
          * Interaktywny podgląd mapy statycznej PNG z trasą i ponumerowanymi znacznikami
          * Kopiowanie i bezpośrednie otwieranie linku do nawigacji Google Maps w przeglądarce
      - Tryb wsadowy (Data Source / Batch):
          * Obsługa formatów Excel (.xlsx, .xls), CSV (.csv, .tsv), JSON (.json)
          * Automatyczne wykrywanie schematu pliku i mapowanie kolumn z możliwością korekty
          * Tabela podglądu danych wejściowych (DataGrid)
          * Pasek postępu, procenty, czas, asynchroniczny log zdarzeń w czasie rzeczywistym
          * Tabela wyników ze statusem i bezpośrednim dostępem do map
          * Eksport raportów zbiorczych do Excel, CSV i JSON
      - Bezpieczeństwo i ustawienia:
          * Szyfrowane przechowywanie klucza Google Maps API (Windows DPAPI per-user)
          * Asynchroniczny tester poprawności klucza API (nie zawiesza interfejsu)
          * Konfiguracja domyślnych wymiarów mapy, katalogów i typu trasy
      - Zgodność ze standardem PS2EXE (samodzielny plik .EXE bez zewnętrznych zależności).

.NOTES
    Encoding: UTF-8 with BOM
#>

# ── 1. Wymuszenie protokołów TLS 1.2 / TLS 1.1 dla zapytań HTTPS ───────────────
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls

# ── 2. Wymuszenie trybu STA dla WPF ──────────────────────────────────────────
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) {
    $currentProcess = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    if ($currentProcess -match 'powershell\.exe|pwsh\.exe') {
        Start-Process -FilePath $currentProcess -ArgumentList "-NoProfile -STA -ExecutionPolicy Bypass -File `"$PSCommandPath`""
        exit
    }
}

# ── 3. Ładowanie bibliotek GUI, Drawing i Security ───────────────────────────
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms, System.Drawing, System.Security

# DWM Dark Mode dla paska tytułu okna Windows 10/11
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class DwmDarkWindow {
    [DllImport("dwmapi.dll", PreserveSig = true)]
    public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
}
"@ -ErrorAction SilentlyContinue

# ══════════════════════════════════════════════════════════════════════════════
# 4. SAMODZIELNE FUNKCJE BAZOWE (EMBEDDED DLA ZGODNOŚCI Z PS2EXE)
# ══════════════════════════════════════════════════════════════════════════════

function Protect-SecretString {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$PlainText)
    if ([string]::IsNullOrEmpty($PlainText)) { return $null }
    try {
        Add-Type -AssemblyName System.Security
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($PlainText)
        $protected = [System.Security.Cryptography.ProtectedData]::Protect(
            $bytes, $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
        return [Convert]::ToBase64String($protected)
    }
    catch {
        try {
            $sec = ConvertTo-SecureString -String $PlainText -AsPlainText -Force
            return (ConvertFrom-SecureString -SecureString $sec)
        } catch {
            return $null
        }
    }
}

function Unprotect-SecretString {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$EncryptedText)
    if ([string]::IsNullOrWhiteSpace($EncryptedText)) { return $null }
    try {
        Add-Type -AssemblyName System.Security
        $bytes = [Convert]::FromBase64String($EncryptedText)
        $unprotected = [System.Security.Cryptography.ProtectedData]::Unprotect(
            $bytes, $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
        return [System.Text.Encoding]::UTF8.GetString($unprotected)
    }
    catch {
        try {
            $sec = ConvertTo-SecureString -String $EncryptedText
            $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)
            $str = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
            return $str
        }
        catch {
            return $null
        }
    }
}

function Test-GoogleApiKey {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ApiKey,
        [Parameter()][string]$LanguageCode = 'en'
    )
    if ([string]::IsNullOrWhiteSpace($ApiKey)) {
        return [PSCustomObject]@{ Valid = $false; Message = 'API key is empty.' }
    }
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls
    try {
        $lang = if ($LanguageCode) { ($LanguageCode -split '[-_]')[0].ToLower() } else { 'en' }
        $Url = "https://maps.googleapis.com/maps/api/geocode/json?address=Warszawa&language=$lang&key=$ApiKey"
        $Resp = Invoke-RestMethod -Uri $Url -Method GET -TimeoutSec 15
        if ($Resp.status -eq 'OK' -or $Resp.status -eq 'ZERO_RESULTS') {
            return [PSCustomObject]@{ Valid = $true; Message = 'Google Maps API key is valid and active.' }
        }
        elseif ($Resp.status -eq 'REQUEST_DENIED') {
            $msg = if ($Resp.error_message) { $Resp.error_message } else { 'Request denied by Google API.' }
            return [PSCustomObject]@{ Valid = $false; Message = "Unauthorized: $msg" }
        }
        else {
            return [PSCustomObject]@{ Valid = $false; Message = "Status API: $($Resp.status)" }
        }
    }
    catch {
        return [PSCustomObject]@{ Valid = $false; Message = "Connection error: $($_.Exception.Message)" }
    }
}

function Select-InputDataFile {
    param([string]$InitialDirectory)
    Add-Type -AssemblyName System.Windows.Forms
    $Dialog = [System.Windows.Forms.OpenFileDialog]::new()
    $Dialog.Title = 'Select route data file (JSON, CSV, Excel)'
    $Dialog.Filter = 'All Supported Files (*.xlsx;*.xls;*.csv;*.tsv;*.json)|*.xlsx;*.xls;*.csv;*.tsv;*.json|Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|CSV/TSV Files (*.csv;*.tsv)|*.csv;*.tsv|JSON Files (*.json)|*.json|All Files (*.*)|*.*'

    $chosenDir = $null
    if ($InitialDirectory -and (Test-Path $InitialDirectory)) {
        $chosenDir = $InitialDirectory
    }
    elseif ($script:LastDataDirectory -and (Test-Path $script:LastDataDirectory)) {
        $chosenDir = $script:LastDataDirectory
    }
    elseif ($script:Config -and $script:Config.LastInputFolder -and (Test-Path $script:Config.LastInputFolder)) {
        $chosenDir = $script:Config.LastInputFolder
    }
    elseif ($script:Config -and $script:Config.LastInputPath -and (Test-Path (Split-Path $script:Config.LastInputPath -Parent))) {
        $chosenDir = Split-Path $script:Config.LastInputPath -Parent
    }
    else {
        $samplesDir = if ($PSScriptRoot) { Join-Path $PSScriptRoot 'Samples' } else { Join-Path (Get-Location) 'Samples' }
        if (Test-Path $samplesDir) {
            $chosenDir = $samplesDir
        } else {
            $chosenDir = [Environment]::GetFolderPath('MyDocuments')
        }
    }

    $Dialog.InitialDirectory = $chosenDir
    $Dialog.RestoreDirectory = $true
    $Result = $Dialog.ShowDialog()
    if ($Result -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:LastDataDirectory = Split-Path $Dialog.FileName -Parent
        if ($script:Config) {
            $script:Config.LastInputFolder = $script:LastDataDirectory
            $script:Config.LastInputPath = $Dialog.FileName
        }
        return $Dialog.FileName
    }
    return $null
}

function Get-AddressComponentValue {
    param([object[]]$Components, [string[]]$Types)
    $Matches = @($Components) | Where-Object {
        $_ -and $_.PSObject.Properties.Name -contains 'types' -and
        (@($_.types) | Where-Object { $_ -in $Types } | Select-Object -First 1)
    } | Select-Object -First 1
    if (-not $Matches) { return $null }
    foreach ($FieldName in @('long_name', 'short_name', 'name')) {
        if ($Matches.PSObject.Properties.Name -contains $FieldName) {
            $Value = [string]$Matches.$FieldName
            if (-not [string]::IsNullOrWhiteSpace($Value)) { return $Value }
        }
    }
    return $null
}

function Get-AddressCoordinates {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Address,
        [Parameter(Mandatory)][string]$ApiKey,
        [Parameter()][string]$LanguageCode = 'en',
        [Parameter()][switch]$RequireStreetNumber
    )
    if ([string]::IsNullOrWhiteSpace($Address)) { return $null }

    if ($Address.Trim() -match '^\s*([+-]?\d+(?:\.\d+)?)\s*[,;\s]\s*([+-]?\d+(?:\.\d+)?)\s*$') {
        $lat = [double]$Matches[1]
        $lng = [double]$Matches[2]
        return [PSCustomObject]@{
            Latitude             = $lat
            Longitude            = $lng
            FormattedAddress     = "$lat, $lng"
            UlicaINumer          = $null
            KodPocztowy          = $null
            Miasto               = $null
            MatchType            = 'COORDINATES'
            PartialMatch         = $false
            Status               = 'OK'
            ErrorMessage         = $null
        }
    }

    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls
    $EncodedAddress = [System.Uri]::EscapeDataString($Address.Trim())
    $lang = if ($LanguageCode) { ($LanguageCode -split '[-_]')[0].ToLower() } else { 'en' }
    $Url = "https://maps.googleapis.com/maps/api/geocode/json?address=$EncodedAddress&language=$lang&key=$ApiKey"
    try {
        $Response = Invoke-RestMethod -Uri $Url -Method GET -TimeoutSec 30
        $Results = @($Response.results)
        if ($Response.status -eq 'OK' -and $Results.Count -gt 0) {
            $ResultItem = $Results[0]
            $Location   = $ResultItem.geometry.location

            $Components   = @($ResultItem.address_components)
            $StreetNumber = Get-AddressComponentValue -Components $Components -Types @('street_number')
            $Route        = Get-AddressComponentValue -Components $Components -Types @('route')
            $PostalCode   = Get-AddressComponentValue -Components $Components -Types @('postal_code')
            $City         = Get-AddressComponentValue -Components $Components -Types @('locality', 'postal_town')
            if ([string]::IsNullOrWhiteSpace($City)) {
                $City = Get-AddressComponentValue -Components $Components -Types @('administrative_area_level_3', 'administrative_area_level_2')
            }

            $StreetWithNumber = if ($Route -and $StreetNumber) { "$Route $StreetNumber" }
                                elseif ($Route) { $Route }
                                elseif ($StreetNumber) { $StreetNumber }
                                else { $null }

            $FormattedAddress = if ($lang -eq 'pl') {
                $ResultItem.formatted_address -replace ',\s*Poland$', ', Polska' -replace '\bPoland\b', 'Polska'
            } else {
                $ResultItem.formatted_address
            }
            $LocationType     = if ($ResultItem.geometry -and $ResultItem.geometry.location_type) { [string]$ResultItem.geometry.location_type } else { 'APPROXIMATE' }
            $PartialMatch     = if ($ResultItem.PSObject.Properties.Name -contains 'partial_match') { [bool]$ResultItem.partial_match } else { $false }

            return [PSCustomObject]@{
                Latitude             = [double]$Location.lat
                Longitude            = [double]$Location.lng
                FormattedAddress     = $FormattedAddress
                UlicaINumer          = $StreetWithNumber
                KodPocztowy          = $PostalCode
                Miasto               = $City
                MatchType            = $LocationType
                PartialMatch         = $PartialMatch
                Status               = 'OK'
                ErrorMessage         = $null
            }
        }
        else {
            return [PSCustomObject]@{
                Latitude             = $null; Longitude = $null; FormattedAddress = $null
                UlicaINumer          = $null; KodPocztowy = $null; Miasto = $null
                MatchType            = $null; PartialMatch = $null
                Status               = $Response.status
                ErrorMessage         = $Response.error_message
            }
        }
    }
    catch {
        $Message = $_.Exception.Message
        return [PSCustomObject]@{
            Latitude             = $null; Longitude = $null; FormattedAddress = $null
            UlicaINumer          = $null; KodPocztowy = $null; Miasto = $null
            MatchType            = $null; PartialMatch = $null
            Status               = "EXCEPTION: $Message"
            ErrorMessage         = $Message
        }
    }
}

function Get-GeocodeStatusDescription {
    [CmdletBinding()]
    param(
        [Parameter()][object]$Geo
    )
    if (-not $Geo) { return 'NOT_PROCESSED' }
    if ($Geo.Status -eq 'OK') {
        if ($Geo.PartialMatch -and $Geo.MatchType -in 'APPROXIMATE', 'GEOMETRIC_CENTER') {
            return "OK (Fallback: Approximate / Partial Match - $($Geo.MatchType))"
        }
        elseif ($Geo.PartialMatch) {
            return "OK (Fallback: Partial Match - $($Geo.MatchType))"
        }
        elseif ($Geo.MatchType -eq 'APPROXIMATE') {
            return 'OK (Fallback: Approximate)'
        }
        elseif ($Geo.MatchType -eq 'GEOMETRIC_CENTER') {
            return 'OK (Fallback: Geometric Center)'
        }
        elseif ($Geo.MatchType -eq 'RANGE_INTERPOLATED') {
            return 'OK (Interpolated)'
        }
        elseif ($Geo.MatchType -eq 'ROOFTOP') {
            return 'OK (Exact - ROOFTOP)'
        }
        elseif ($Geo.MatchType -eq 'COORDINATES') {
            return 'OK (Coordinates)'
        }
        else {
            return "OK ($($Geo.MatchType))"
        }
    }
    elseif ($Geo.Status -eq 'ZERO_RESULTS') {
        return 'ZERO_RESULTS (Address Not Found)'
    }
    else {
        return [string]$Geo.Status
    }
}

function Get-CarRouteData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][double]$OriginLat,
        [Parameter(Mandatory)][double]$OriginLng,
        [Parameter(Mandatory)][double]$DestLat,
        [Parameter(Mandatory)][double]$DestLng,
        [Parameter(Mandatory)][string]$ApiKey,
        [Parameter()][object[]]$IntermediatePoints = @(),
        [Parameter()][ValidateSet('Fastest', 'Shortest', 'Eco')][string]$RouteType = 'Fastest',
        [Parameter()][ValidateSet('GASOLINE', 'DIESEL', 'HYBRID', 'ELECTRIC')][string]$EmissionType = 'GASOLINE',
        [Parameter()][string]$LanguageCode = 'en',
        [Parameter()][string]$Units = 'METRIC',
        [Parameter()][switch]$TrafficAware
    )

    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls
    $RoutesUrl = 'https://routes.googleapis.com/directions/v2:computeRoutes'

    $RequestBody = [ordered]@{
        origin       = @{ location = @{ latLng = @{ latitude = $OriginLat; longitude = $OriginLng } } }
        destination  = @{ location = @{ latLng = @{ latitude = $DestLat; longitude = $DestLng } } }
        travelMode   = 'DRIVE'
        languageCode = if ($LanguageCode) { $LanguageCode } else { 'en' }
        units        = $Units
    }

    $HasIntermediates = $false
    if ($null -ne $IntermediatePoints -and @($IntermediatePoints).Count -gt 0) {
        $IntermediatesList = [System.Collections.Generic.List[hashtable]]::new()
        foreach ($pt in $IntermediatePoints) {
            if ($null -ne $pt -and $pt.Latitude -and $pt.Longitude) {
                $IntermediatesList.Add(@{
                    location = @{
                        latLng = @{
                            latitude  = [double]$pt.Latitude
                            longitude = [double]$pt.Longitude
                        }
                    }
                })
            }
        }
        if ($IntermediatesList.Count -gt 0) {
            $RequestBody['intermediates'] = $IntermediatesList
            $HasIntermediates = $true
        }
    }

    switch ($RouteType) {
        'Fastest' {
            $RequestBody['routingPreference'] = if ($TrafficAware) { 'TRAFFIC_AWARE' } else { 'TRAFFIC_UNAWARE' }
            if (-not $HasIntermediates) { $RequestBody['computeAlternativeRoutes'] = $true }
        }
        'Shortest' {
            $RequestBody['routingPreference'] = 'TRAFFIC_UNAWARE'
            if (-not $HasIntermediates) { $RequestBody['computeAlternativeRoutes'] = $true }
        }
        'Eco' {
            $RequestBody['routingPreference'] = 'TRAFFIC_AWARE_OPTIMAL'
            $RequestBody['requestedReferenceRoutes'] = @('FUEL_EFFICIENT')
            $RequestBody['routeModifiers'] = @{
                vehicleInfo = @{ emissionType = $EmissionType }
            }
        }
    }

    $Headers = @{
        'X-Goog-Api-Key'   = $ApiKey
        'Content-Type'     = 'application/json'
        'X-Goog-FieldMask' = 'routes.duration,routes.distanceMeters,routes.polyline.encodedPolyline,routes.description,routes.routeLabels'
    }

    try {
        $JsonBody = $RequestBody | ConvertTo-Json -Depth 10
        $Response = Invoke-RestMethod -Uri $RoutesUrl -Method POST -Headers $Headers -Body $JsonBody -TimeoutSec 60

        $Routes = @($Response.routes)
        if ($Routes.Count -eq 0) {
            return [PSCustomObject]@{
                OdlegloscKm     = $null
                CzasMin         = $null
                DurationSeconds = $null
                EncodedPolyline = $null
                RouteType       = $RouteType
                RouteLabels     = @()
                Status          = 'NO_ROUTES'
                ErrorMessage    = 'Google Routes API did not return any routes.'
            }
        }

        $SelectedRoute = $null
        if ($RouteType -eq 'Shortest') {
            $SelectedRoute = $Routes | Sort-Object -Property { [int64]($_.distanceMeters) } | Select-Object -First 1
        }
        elseif ($RouteType -eq 'Eco') {
            $EcoRoute = $Routes | Where-Object {
                $_.routeLabels -and (@($_.routeLabels) -contains 'FUEL_EFFICIENT')
            } | Select-Object -First 1

            $SelectedRoute = if ($EcoRoute) { $EcoRoute } else { $Routes[0] }
        }
        else {
            $SelectedRoute = $Routes | Sort-Object -Property {
                if ($_.duration) { [double]($_.duration.TrimEnd('s')) } else { [double]::MaxValue }
            } | Select-Object -First 1
        }

        $DistanceKm = if ($SelectedRoute.distanceMeters) { [math]::Round([double]$SelectedRoute.distanceMeters / 1000.0, 2) } else { $null }
        $DurationSec = if ($SelectedRoute.duration) { [double]($SelectedRoute.duration.TrimEnd('s')) } else { $null }
        $DurationMinutes = if ($null -ne $DurationSec) { [math]::Round($DurationSec / 60.0, 0) } else { $null }
        $Polyline = if ($SelectedRoute.polyline) { $SelectedRoute.polyline.encodedPolyline } else { $null }
        $Labels = if ($SelectedRoute.routeLabels) { @($SelectedRoute.routeLabels) } else { @() }

        return [PSCustomObject]@{
            OdlegloscKm     = $DistanceKm
            CzasMin         = $DurationMinutes
            DurationSeconds = $DurationSec
            EncodedPolyline = $Polyline
            RouteType       = $RouteType
            RouteLabels     = $Labels
            Status          = 'OK'
            ErrorMessage    = $null
        }
    }
    catch {
        $ErrorMsg = $_.Exception.Message
        return [PSCustomObject]@{
            OdlegloscKm     = $null
            CzasMin         = $null
            DurationSeconds = $null
            EncodedPolyline = $null
            RouteType       = $RouteType
            RouteLabels     = @()
            Status          = "EXCEPTION: $ErrorMsg"
            ErrorMessage    = $ErrorMsg
        }
    }
}

function Get-GoogleMapsUrl {
    param(
        [string]$Origin,
        [string]$Destination,
        [object[]]$Waypoints = @(),
        [string]$TravelMode = 'driving'
    )
    $OriginEnc = [System.Uri]::EscapeDataString($Origin.Trim())
    $DestEnc = [System.Uri]::EscapeDataString($Destination.Trim())
    $Url = "https://www.google.com/maps/dir/?api=1&origin=$OriginEnc&destination=$DestEnc&travelmode=$TravelMode"

    if ($null -ne $Waypoints -and @($Waypoints).Count -gt 0) {
        $WpStrings = [System.Collections.Generic.List[string]]::new()
        foreach ($wp in $Waypoints) {
            if ($wp -is [string] -and -not [string]::IsNullOrWhiteSpace($wp)) {
                $WpStrings.Add($wp.Trim())
            }
            elseif ($wp.Latitude -and $wp.Longitude) {
                $WpStrings.Add("$($wp.Latitude),$($wp.Longitude)")
            }
            elseif ($wp.ZapytanieAdresowe) {
                $WpStrings.Add([string]$wp.ZapytanieAdresowe)
            }
            elseif ($wp.AdresGeokodowany) {
                $WpStrings.Add([string]$wp.AdresGeokodowany)
            }
        }
        if ($WpStrings.Count -gt 0) {
            $Url += '&waypoints=' + [System.Uri]::EscapeDataString(($WpStrings -join '|'))
        }
    }
    return $Url
}

function Get-WrappedLines {
    param([System.Drawing.Graphics]$G, [string]$Text, [System.Drawing.Font]$F, [float]$MaxW)
    if ([string]::IsNullOrWhiteSpace($Text)) { return [string[]]@('') }
    if ($G.MeasureString($Text, $F).Width -le $MaxW) { return [string[]]@($Text) }
    $Words = $Text -split '\s+'
    $L1 = ''; $L2 = ''; $On2 = $false
    foreach ($W in $Words) {
        if (-not $On2) {
            $T = if ($L1) { "$L1 $W" } else { $W }
            if ($G.MeasureString($T, $F).Width -le $MaxW) { $L1 = $T }
            else { $On2 = $true; $L2 = $W }
        }
        else {
            $T2 = if ($L2) { "$L2 $W" } else { $W }
            if ($G.MeasureString($T2, $F).Width -le $MaxW) { $L2 = $T2 }
            else {
                if ($L2.Length -gt 3) { $L2 = $L2.Substring(0, $L2.Length - 3) + '...' }
                break
            }
        }
    }
    if ($L2) { return [string[]]@($L1, $L2) } else { return [string[]]@($L1) }
}

function Save-RouteMapPng {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$EncodedPolyline,
        [Parameter(Mandatory)][double]$OriginLat,
        [Parameter(Mandatory)][double]$OriginLng,
        [Parameter(Mandatory)][double]$DestLat,
        [Parameter(Mandatory)][double]$DestLng,
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)][string]$ApiKey,
        [Parameter()][int]$Width = 900,
        [Parameter()][int]$Height = 600,
        [Parameter()][object[]]$RoutePoints = @(),
        [Parameter()][Alias('TekstAdresA')][string]$AddressTextA = '',
        [Parameter()][Alias('TekstAdresB')][string]$AddressTextB = '',
        [Parameter()][Alias('TekstOdleglosc')][string]$DistanceText = '',
        [Parameter()][Alias('TekstCzas')][string]$DurationText = '',
        [Parameter()][Alias('TekstNaglowekLewy')][string]$HeaderLeftText = '',
        [Parameter()][Alias('TekstNaglowekPrawy')][string]$HeaderRightText = '',
        [Parameter()][Alias('TekstUmowa')][string]$ContractText = '',
        [Parameter()][Alias('TekstKierunek')][string]$DirectionText = '',
        [Parameter()][Alias('Opis')][string]$Description = '',
        [Parameter()][Alias('DataWygenerowania')][string]$GeneratedDate = '',
        [Parameter()][string]$LanguageCode = 'en',
        [Parameter()][string]$StartRaw = '',
        [Parameter()][string]$StartGeocoded = '',
        [Parameter()][string]$EndRaw = '',
        [Parameter()][string]$EndGeocoded = '',
        [Parameter()][object[]]$WaypointsList = @(),
        [Parameter()][string]$RouteName = '',
        [Parameter()][string]$RouteType = '',
        [Parameter()][object]$OverlayConfig = $null
    )

    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls
    $EncodedForUrl = [System.Uri]::EscapeDataString($EncodedPolyline)
    $MarkerParams = [System.Collections.Generic.List[string]]::new()

    $MarkerStart = [System.Uri]::EscapeDataString("color:green|label:A|$OriginLat,$OriginLng")
    $MarkerParams.Add("&markers=$MarkerStart")

    if ($null -ne $RoutePoints -and @($RoutePoints).Count -gt 0) {
        $IntermediatesOnly = @()
        if ($RoutePoints.Count -gt 2 -and
            [math]::Abs($RoutePoints[0].Latitude - $OriginLat) -lt 0.0001 -and
            [math]::Abs($RoutePoints[-1].Latitude - $DestLat) -lt 0.0001) {
            $IntermediatesOnly = @($RoutePoints[1..($RoutePoints.Count - 2)])
        }
        else {
            $IntermediatesOnly = @($RoutePoints | Where-Object {
                $null -ne $_.Latitude -and $null -ne $_.Longitude -and
                (-not ([math]::Abs($_.Latitude - $OriginLat) -lt 0.0001 -and [math]::Abs($_.Longitude - $OriginLng) -lt 0.0001)) -and
                (-not ([math]::Abs($_.Latitude - $DestLat) -lt 0.0001 -and [math]::Abs($_.Longitude - $DestLng) -lt 0.0001))
            })
        }

        $idx = 1
        foreach ($pt in $IntermediatesOnly) {
            if ($pt.Latitude -and $pt.Longitude) {
                $lbl = if ($idx -le 9) { [string]$idx }
                       elseif ($idx -le 35) { [string][char](55 + $idx) }
                       else { '' }
                $spec = if ($lbl) { "color:blue|label:$lbl|$($pt.Latitude),$($pt.Longitude)" }
                        else { "size:mid|color:blue|$($pt.Latitude),$($pt.Longitude)" }
                $MarkerParams.Add("&markers=" + [System.Uri]::EscapeDataString($spec))
                $idx++
            }
        }
    }

    $MarkerEnd = [System.Uri]::EscapeDataString("color:red|label:B|$DestLat,$DestLng")
    $MarkerParams.Add("&markers=$MarkerEnd")

    $lang = if ($LanguageCode) { ($LanguageCode -split '[-_]')[0].ToLower() } else { 'en' }
    $StaticMapUrl = ("https://maps.googleapis.com/maps/api/staticmap" +
        "?size=${Width}x${Height}" +
        "&language=$lang" +
        "&path=weight:4|color:0x0066FFff|enc:$EncodedForUrl" +
        ($MarkerParams -join '') +
        "&key=$ApiKey")

    try {
        $TargetDir = Split-Path -Parent $OutputPath
        if (-not [string]::IsNullOrWhiteSpace($TargetDir) -and -not (Test-Path $TargetDir)) {
            New-Item -ItemType Directory -Path $TargetDir -Force | Out-Null
        }

        $wc = [System.Net.WebClient]::new()
        try {
            $wc.DownloadFile($StaticMapUrl, $OutputPath)
        }
        finally {
            $wc.Dispose()
        }

        # Resolve overlay configuration
        if ($OverlayConfig -is [string] -and -not [string]::IsNullOrWhiteSpace($OverlayConfig)) {
            try { $OverlayConfig = $OverlayConfig | ConvertFrom-Json } catch { }
        }
        if (-not $OverlayConfig) {
            $OverlayConfig = [PSCustomObject]@{
                EnableTopOverlay    = $true
                EnableBottomOverlay = $true
                Items               = [PSCustomObject]@{
                    StartGeocoded = [PSCustomObject]@{ Enabled = $true;  Panel = 'Bottom'; Align = 'Left';   Order = 1 }
                    EndGeocoded   = [PSCustomObject]@{ Enabled = $true;  Panel = 'Bottom'; Align = 'Left';   Order = 2 }
                    Distance      = [PSCustomObject]@{ Enabled = $true;  Panel = 'Bottom'; Align = 'Left';   Order = 3 }
                    Duration      = [PSCustomObject]@{ Enabled = $true;  Panel = 'Bottom'; Align = 'Center'; Order = 3 }
                    Timestamp     = [PSCustomObject]@{ Enabled = $true;  Panel = 'Bottom'; Align = 'Right';  Order = 3 }
                    RouteName     = [PSCustomObject]@{ Enabled = $true;  Panel = 'Top';    Align = 'Left';   Order = 1 }
                    RouteType     = [PSCustomObject]@{ Enabled = $true;  Panel = 'Top';    Align = 'Right';  Order = 1 }
                    Waypoints     = [PSCustomObject]@{ Enabled = $false; Panel = 'Bottom'; Align = 'Left';   Order = 2 }
                    StartRaw      = [PSCustomObject]@{ Enabled = $false; Panel = 'None';   Align = 'Left';   Order = 1 }
                    EndRaw        = [PSCustomObject]@{ Enabled = $false; Panel = 'None';   Align = 'Left';   Order = 2 }
                }
            }
        }

        $enableTop = if ($null -ne $OverlayConfig.EnableTopOverlay) { [bool]$OverlayConfig.EnableTopOverlay } else { $true }
        $enableBtm = if ($null -ne $OverlayConfig.EnableBottomOverlay) { [bool]$OverlayConfig.EnableBottomOverlay } else { $true }

        # Resolve data values
        $addrStartGeo = if ($StartGeocoded) { $StartGeocoded } elseif ($AddressTextA) { $AddressTextA } else { '' }
        $addrStartRaw = if ($StartRaw) { $StartRaw } else { '' }
        $addrEndGeo   = if ($EndGeocoded) { $EndGeocoded } elseif ($AddressTextB) { $AddressTextB } else { '' }
        $addrEndRaw   = if ($EndRaw) { $EndRaw } else { '' }

        $nameVal = if ($RouteName) { $RouteName } elseif ($HeaderLeftText) { $HeaderLeftText } elseif ($Description) { $Description.Trim() } elseif ($ContractText) { $ContractText } else { '' }

        $typeVal = if ($RouteType) { $RouteType } elseif ($HeaderRightText) { $HeaderRightText } elseif ($DirectionText) { $DirectionText } else { '' }
        if ($typeVal -match '^(?:Type|Typ|Art):\s*(.+)$' -or $typeVal -match '^(Shortest|Fastest|Eco|Najkr[oó]tsza|Najszybsza|Eko|K[uü]rzeste|Schnellste)$') {
            $rawVal = if ($Matches[1]) { $Matches[1].Trim() } else { $Matches[0].Trim() }
            $normVal = if ($rawVal -match '(?i)short|kr[oó]t|k[uü]rz') { 'Shortest' }
                       elseif ($rawVal -match '(?i)eco|eko|fuel') { 'Eco' }
                       elseif ($rawVal -match '(?i)fast|szyb|schnell') { 'Fastest' }
                       else { $rawVal }
            $tPrefix = switch ($lang) { 'de' { 'Typ: ' } 'pl' { 'Typ: ' } default { 'Type: ' } }
            $tName = switch ($lang) {
                'de' { if ($normVal -eq 'Fastest') { 'Schnellste' } elseif ($normVal -eq 'Shortest') { 'Kürzeste' } elseif ($normVal -eq 'Eco') { 'Eco' } else { $normVal } }
                'pl' { if ($normVal -eq 'Fastest') { 'Najszybsza' } elseif ($normVal -eq 'Shortest') { 'Najkrótsza' } elseif ($normVal -eq 'Eco') { 'Eko' } else { $normVal } }
                default { $normVal }
            }
            $typeVal = "$tPrefix$tName"
        }

        $distPrefix = switch ($lang) { 'de' { 'Gesamt: ' } 'pl' { 'Razem: ' } default { 'Total: ' } }
        $distVal = if ($DistanceText) { $DistanceText } else { '' }

        $durVal = if ($DurationText) {
            if ($DurationText -match '^\(.*\)$') { $DurationText } else { "($DurationText)" }
        } else { '' }

        $dateVal = if ($GeneratedDate) { $GeneratedDate } else { (Get-Date -Format 'yyyy-MM-dd  HH:mm') }

        $wpItems = [System.Collections.Generic.List[PSCustomObject]]::new()
        $rawWpList = if ($WaypointsList -and @($WaypointsList).Count -gt 0) {
            $WaypointsList
        } elseif ($RoutePoints -and @($RoutePoints).Count -gt 2) {
            @($RoutePoints[1..($RoutePoints.Count - 2)])
        } else { @() }

        $wIdx = 1
        foreach ($w in $rawWpList) {
            $wText = if ($w -is [string]) { $w }
                     elseif ($w.FormattedAddress) { $w.FormattedAddress }
                     elseif ($w.Address) { $w.Address }
                     else { '' }
            if (-not [string]::IsNullOrWhiteSpace($wText)) {
                $wpItems.Add([PSCustomObject]@{
                    Index = $wIdx
                    Badge = "${wIdx}: "
                    Text  = $wText
                })
                $wIdx++
            }
        }

        # Build active property items map
        $propDataMap = @{
            'StartGeocoded' = @{ Id='StartGeocoded'; Kind='address'; Badge='A: '; BadgeColor='Green'; Text=$addrStartGeo }
            'StartRaw'      = @{ Id='StartRaw';      Kind='address'; Badge='A: '; BadgeColor='Green'; Text=$addrStartRaw }
            'EndGeocoded'   = @{ Id='EndGeocoded';   Kind='address'; Badge='B: '; BadgeColor='Red';   Text=$addrEndGeo }
            'EndRaw'        = @{ Id='EndRaw';        Kind='address'; Badge='B: '; BadgeColor='Red';   Text=$addrEndRaw }
            'Distance'      = @{ Id='Distance';      Kind='stat';    Prefix=$distPrefix; Value=$distVal }
            'Duration'      = @{ Id='Duration';      Kind='stat';    Value=$durVal }
            'Timestamp'     = @{ Id='Timestamp';     Kind='date';    Text=$dateVal }
            'RouteName'     = @{ Id='RouteName';     Kind='title';   Text=$nameVal }
            'RouteType'     = @{ Id='RouteType';     Kind='type';    Text=$typeVal }
            'Waypoints'     = @{ Id='Waypoints';     Kind='waypoints'; Items=$wpItems }
        }

        $topItems = [System.Collections.Generic.List[PSCustomObject]]::new()
        $btmItems = [System.Collections.Generic.List[PSCustomObject]]::new()

        if ($OverlayConfig.Items) {
            $propNames = if ($OverlayConfig.Items -is [System.Collections.IDictionary]) {
                $OverlayConfig.Items.Keys
            } else {
                $OverlayConfig.Items.PSObject.Properties.Name
            }
            foreach ($pName in $propNames) {
                $iCfg = if ($OverlayConfig.Items -is [System.Collections.IDictionary]) {
                    $OverlayConfig.Items[$pName]
                } else {
                    $OverlayConfig.Items.$pName
                }
                if (-not $iCfg) { continue }
                $pEnabled = if ($null -ne $iCfg.Enabled) { [bool]$iCfg.Enabled } else { $true }
                $pPanel   = if ($iCfg.Panel) { [string]$iCfg.Panel } else { 'None' }
                $pAlign   = if ($iCfg.Align) { [string]$iCfg.Align } else { 'Left' }
                $pOrder   = if ($iCfg.Order) { [int]$iCfg.Order } else { 1 }

                if (-not $pEnabled -or $pPanel -eq 'None') { continue }
                if (-not $propDataMap.ContainsKey($pName)) { continue }

                $pData = $propDataMap[$pName]
                $hasContent = $false
                if ($pData.Kind -eq 'waypoints') {
                    $hasContent = ($pData.Items -and $pData.Items.Count -gt 0)
                } elseif ($pData.Kind -eq 'stat') {
                    $hasContent = (-not [string]::IsNullOrWhiteSpace($pData.Value))
                } else {
                    $hasContent = (-not [string]::IsNullOrWhiteSpace($pData.Text))
                }
                if (-not $hasContent) { continue }

                $itemObj = [PSCustomObject]@{
                    Id         = $pName
                    Kind       = $pData.Kind
                    Badge      = $pData.Badge
                    BadgeColor = $pData.BadgeColor
                    Text       = $pData.Text
                    Prefix     = $pData.Prefix
                    Value      = $pData.Value
                    Items      = $pData.Items
                    Panel      = $pPanel
                    Align      = $pAlign
                    Order      = $pOrder
                }

                if ($pPanel -eq 'Top' -and $enableTop) {
                    $topItems.Add($itemObj)
                } elseif ($pPanel -eq 'Bottom' -and $enableBtm) {
                    $btmItems.Add($itemObj)
                }
            }
        }

        $MaTopOverlay = ($enableTop -and $topItems.Count -gt 0)
        $MaBottomOverlay = ($enableBtm -and $btmItems.Count -gt 0)

        if ($MaTopOverlay -or $MaBottomOverlay) {
            try {
                Add-Type -AssemblyName System.Drawing

                $FileBytes = [System.IO.File]::ReadAllBytes($OutputPath)
                $MemStream = [System.IO.MemoryStream]::new($FileBytes)
                $BitmapSrc = [System.Drawing.Bitmap]::new($MemStream)

                $ActualW = $BitmapSrc.Width
                $ActualH = $BitmapSrc.Height

                # Fonts definition
                $FontTopTitle = [System.Drawing.Font]::new('Segoe UI', 10.0, [System.Drawing.FontStyle]::Bold)
                $FontTopType  = [System.Drawing.Font]::new('Segoe UI', 10.0, [System.Drawing.FontStyle]::Bold)
                $FontBadge    = [System.Drawing.Font]::new('Segoe UI', 9.0,  [System.Drawing.FontStyle]::Bold)
                $FontAddr     = [System.Drawing.Font]::new('Segoe UI', 9.5,  [System.Drawing.FontStyle]::Regular)
                $FontDistLbl  = [System.Drawing.Font]::new('Segoe UI', 9.0,  [System.Drawing.FontStyle]::Bold)
                $FontDist     = [System.Drawing.Font]::new('Segoe UI', 12.0, [System.Drawing.FontStyle]::Bold)
                $FontDate     = [System.Drawing.Font]::new('Segoe UI', 8.5,  [System.Drawing.FontStyle]::Regular)

                $PadX  = 14
                $LineH = 20

                # Pre-measurement Graphics
                $dummyBmp = [System.Drawing.Bitmap]::new(1, 1)
                $measGfx  = [System.Drawing.Graphics]::FromImage($dummyBmp)

                # Helper scriptblock to group items by Order
                $BuildRows = {
                    param($items)
                    $orders = @($items | Select-Object -ExpandProperty Order -Unique | Sort-Object)
                    $rows = [System.Collections.Generic.List[PSCustomObject]]::new()
                    foreach ($ord in $orders) {
                        $rowItems = @($items | Where-Object { $_.Order -eq $ord })
                        $left   = [System.Collections.Generic.List[PSCustomObject]]::new()
                        $center = [System.Collections.Generic.List[PSCustomObject]]::new()
                        $right  = [System.Collections.Generic.List[PSCustomObject]]::new()
                        foreach ($it in $rowItems) {
                            if ($it.Align -eq 'Right') { $right.Add($it) }
                            elseif ($it.Align -eq 'Center') { $center.Add($it) }
                            else { $left.Add($it) }
                        }
                        $rows.Add([PSCustomObject]@{
                            Order  = $ord
                            Left   = $left
                            Center = $center
                            Right  = $right
                            Height = 20
                        })
                    }
                    return $rows.ToArray()
                }

                $topRows = @(if ($MaTopOverlay) { & $BuildRows $topItems } else { @() })
                $btmRows = @(if ($MaBottomOverlay) { & $BuildRows $btmItems } else { @() })

                # Measure row heights
                $MeasureRows = {
                    param($rows, $availWidth)
                    foreach ($row in @($rows)) {
                        $maxH = 20
                        $allItems = @($row.Left) + @($row.Center) + @($row.Right)
                        foreach ($it in $allItems) {
                            if ($it.Kind -eq 'address') {
                                $badgeSz = $measGfx.MeasureString($it.Badge, $FontBadge)
                                $addrW = [float]($availWidth - $badgeSz.Width)
                                $lines = @(Get-WrappedLines -G $measGfx -Text $it.Text -F $FontAddr -MaxW $addrW)
                                $it | Add-Member -NotePropertyName 'WrappedLines' -NotePropertyValue $lines -Force
                                $h = [math]::Max(1, $lines.Count) * $LineH
                                if ($h -gt $maxH) { $maxH = $h }
                            }
                            elseif ($it.Kind -eq 'waypoints') {
                                $totalWpH = 0
                                foreach ($wp in $it.Items) {
                                    $bSz = $measGfx.MeasureString($wp.Badge, $FontBadge)
                                    $wpMaxW = [float]($availWidth - $bSz.Width)
                                    $wpLines = @(Get-WrappedLines -G $measGfx -Text $wp.Text -F $FontAddr -MaxW $wpMaxW)
                                    $wp | Add-Member -NotePropertyName 'WrappedLines' -NotePropertyValue $wpLines -Force
                                    $totalWpH += [math]::Max(1, $wpLines.Count) * $LineH
                                }
                                if ($totalWpH -gt $maxH) { $maxH = $totalWpH }
                            }
                            elseif ($it.Kind -eq 'stat') {
                                if ($maxH -lt 24) { $maxH = 24 }
                            }
                            elseif ($it.Kind -in @('title', 'type')) {
                                if ($maxH -lt 22) { $maxH = 22 }
                            }
                        }
                        $row.Height = $maxH
                    }
                }

                $availContentW = [float]($ActualW - ($PadX * 2))
                & $MeasureRows $topRows $availContentW
                & $MeasureRows $btmRows $availContentW

                $measGfx.Dispose()
                $dummyBmp.Dispose()

                # Calculate banner heights
                $TopPad = 8; $TopBotPad = 8; $TopRowSpacing = 4
                $TopBarH = 0
                if ($MaTopOverlay -and @($topRows).Count -gt 0) {
                    $sumTopH = (@($topRows) | Measure-Object -Property Height -Sum).Sum
                    if (-not $sumTopH) { $sumTopH = 20 }
                    $TopBarH = [int]($TopPad + $sumTopH + ((@($topRows).Count - 1) * $TopRowSpacing) + $TopBotPad)
                    if ($TopBarH -lt 38) { $TopBarH = 38 }
                }

                $BtmPadTop = 10; $BtmPadBot = 10; $BtmRowSpacing = 6
                $BtmBarH = 0
                if ($MaBottomOverlay -and @($btmRows).Count -gt 0) {
                    $sumBtmH = (@($btmRows) | Measure-Object -Property Height -Sum).Sum
                    if (-not $sumBtmH) { $sumBtmH = 20 }
                    $BtmBarH = [int]($BtmPadTop + $sumBtmH + ((@($btmRows).Count - 1) * $BtmRowSpacing) + $BtmPadBot)
                }

                $FinalW = $ActualW
                $FinalH = $ActualH + $TopBarH + $BtmBarH

                $Bitmap = [System.Drawing.Bitmap]::new($FinalW, $FinalH, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
                $Graphics = [System.Drawing.Graphics]::FromImage($Bitmap)
                $Graphics.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
                $Graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit

                # 1. Background fill
                $BrushBg = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 15, 23, 42))
                $Graphics.FillRectangle($BrushBg, 0, 0, $FinalW, $FinalH)

                # 2. Draw map image in the middle
                $Graphics.DrawImage($BitmapSrc, 0, $TopBarH, $ActualW, $ActualH)

                # 3. Brushes & Pens
                $PenSep      = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(255, 51, 65, 85), 1.5)
                $BrushWhite  = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 248, 250, 252))
                $BrushYellow = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 250, 204, 21))
                $BrushCyan   = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 56, 189, 248))
                $BrushGreen  = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 16, 185, 129))
                $BrushRed    = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 239, 68, 68))
                $BrushMuted  = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 148, 163, 184))

                # Helper scriptblock to measure width of an item
                $MeasureItemWidth = {
                    param($it)
                    if ($it.Kind -eq 'address') {
                        $bSz = $Graphics.MeasureString($it.Badge, $FontBadge)
                        $tSz = $Graphics.MeasureString($it.Text, $FontAddr)
                        return ($bSz.Width + $tSz.Width)
                    }
                    elseif ($it.Kind -eq 'stat') {
                        $w = 0
                        if ($it.Prefix) { $w += $Graphics.MeasureString($it.Prefix, $FontDistLbl).Width }
                        if ($it.Value)  { $w += $Graphics.MeasureString($it.Value, $FontDist).Width }
                        return $w
                    }
                    elseif ($it.Kind -eq 'title') {
                        return $Graphics.MeasureString($it.Text, $FontTopTitle).Width
                    }
                    elseif ($it.Kind -eq 'type') {
                        return $Graphics.MeasureString($it.Text, $FontTopType).Width
                    }
                    elseif ($it.Kind -eq 'date') {
                        return $Graphics.MeasureString($it.Text, $FontDate).Width
                    }
                    elseif ($it.Kind -eq 'waypoints') {
                        return 200
                    }
                    return 0
                }

                # Helper scriptblock to draw an item at specified coordinates
                $DrawItem = {
                    param($it, [float]$x, [float]$y)
                    if ($it.Kind -eq 'address') {
                        $badgeBrush = if ($it.BadgeColor -eq 'Red') { $BrushRed } else { $BrushGreen }
                        $Graphics.DrawString($it.Badge, $FontBadge, $badgeBrush, $x, $y)
                        $bSz = $Graphics.MeasureString($it.Badge, $FontBadge)
                        $curLineY = $y
                        $lines = if ($it.WrappedLines) { $it.WrappedLines } else { @($it.Text) }
                        foreach ($line in $lines) {
                            $Graphics.DrawString($line, $FontAddr, $BrushWhite, ($x + $bSz.Width), $curLineY)
                            $curLineY += [float]$LineH
                        }
                    }
                    elseif ($it.Kind -eq 'waypoints') {
                        $wpY = $y
                        foreach ($wp in $it.Items) {
                            $Graphics.DrawString($wp.Badge, $FontBadge, $BrushCyan, $x, $wpY)
                            $bSz = $Graphics.MeasureString($wp.Badge, $FontBadge)
                            $lines = if ($wp.WrappedLines) { $wp.WrappedLines } else { @($wp.Text) }
                            foreach ($line in $lines) {
                                $Graphics.DrawString($line, $FontAddr, $BrushWhite, ($x + $bSz.Width), $wpY)
                                $wpY += [float]$LineH
                            }
                        }
                    }
                    elseif ($it.Kind -eq 'stat') {
                        $statX = $x
                        if ($it.Prefix) {
                            $pSz = $Graphics.MeasureString($it.Prefix, $FontDistLbl)
                            $Graphics.DrawString($it.Prefix, $FontDistLbl, $BrushCyan, $statX, ($y + 2))
                            $statX += $pSz.Width
                        }
                        if ($it.Value) {
                            $Graphics.DrawString($it.Value, $FontDist, $BrushYellow, $statX, $y)
                        }
                    }
                    elseif ($it.Kind -eq 'title') {
                        $Graphics.DrawString($it.Text, $FontTopTitle, $BrushWhite, $x, $y)
                    }
                    elseif ($it.Kind -eq 'type') {
                        $Graphics.DrawString($it.Text, $FontTopType, $BrushYellow, $x, $y)
                    }
                    elseif ($it.Kind -eq 'date') {
                        $Graphics.DrawString($it.Text, $FontDate, $BrushMuted, $x, ($y + 3))
                    }
                }

                # Helper scriptblock to render a banner's rows
                $RenderBannerRows = {
                    param($rows, [float]$startY, [float]$spacing)
                    $curY = $startY
                    foreach ($row in $rows) {
                        $leftX = [float]$PadX

                        # 1. Left items
                        foreach ($it in $row.Left) {
                            & $DrawItem $it $leftX $curY
                            $w = & $MeasureItemWidth $it
                            $leftX += [float]($w + 14)
                        }

                        # 2. Right items
                        $totalRightW = 0
                        foreach ($it in $row.Right) {
                            $totalRightW += [float]((& $MeasureItemWidth $it) + 12)
                        }
                        $rightX = [float]($FinalW - $PadX - $totalRightW + 12)
                        foreach ($it in $row.Right) {
                            & $DrawItem $it $rightX $curY
                            $w = & $MeasureItemWidth $it
                            $rightX += [float]($w + 12)
                        }

                        # 3. Center items
                        $totalCenterW = 0
                        foreach ($it in $row.Center) {
                            $totalCenterW += [float]((& $MeasureItemWidth $it) + 12)
                        }
                        $centerX = [float][math]::Max($leftX + 10, ($FinalW - $totalCenterW + 12) / 2)
                        foreach ($it in $row.Center) {
                            & $DrawItem $it $centerX $curY
                            $w = & $MeasureItemWidth $it
                            $centerX += [float]($w + 12)
                        }

                        $curY += [float]($row.Height + $spacing)
                    }
                }

                # 4. Draw Top Header Banner
                if ($MaTopOverlay -and $TopBarH -gt 0 -and @($topRows).Count -gt 0) {
                    $Graphics.DrawLine($PenSep, 0, $TopBarH, $FinalW, $TopBarH)
                    $topStartY = [float]$TopPad
                    if (@($topRows).Count -eq 1) {
                        $topStartY = [float][math]::Max(6, ($TopBarH - $topRows[0].Height) / 2)
                    }
                    & $RenderBannerRows $topRows $topStartY $TopRowSpacing
                }

                # 5. Draw Bottom Footer Banner
                if ($MaBottomOverlay -and $BtmBarH -gt 0 -and @($btmRows).Count -gt 0) {
                    $BtmBarY = $TopBarH + $ActualH
                    $Graphics.DrawLine($PenSep, 0, $BtmBarY, $FinalW, $BtmBarY)
                    $btmStartY = [float]($BtmBarY + $BtmPadTop)
                    & $RenderBannerRows $btmRows $btmStartY $BtmRowSpacing
                }

                # Dispose GDI+ objects
                $PenSep.Dispose()
                $BrushBg.Dispose(); $BrushWhite.Dispose(); $BrushYellow.Dispose()
                $BrushCyan.Dispose(); $BrushGreen.Dispose(); $BrushRed.Dispose(); $BrushMuted.Dispose()
                $FontTopTitle.Dispose(); $FontTopType.Dispose()
                $FontBadge.Dispose(); $FontAddr.Dispose(); $FontDistLbl.Dispose(); $FontDist.Dispose(); $FontDate.Dispose()
                $Graphics.Dispose()

                $Bitmap.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
                $Bitmap.Dispose()
                $BitmapSrc.Dispose()
                $MemStream.Dispose()
            }
            catch { }
        }
        return $true
    }
    catch {
        return $false
    }
}

function Find-MatchingPropertyName {
    param([string[]]$AvailableProperties, [string[]]$Patterns)
    foreach ($pattern in $Patterns) {
        $found = $AvailableProperties | Where-Object { $null -ne $_ -and $_.Trim() -match $pattern } | Select-Object -First 1
        if ($found) { return $found }
    }
    return $null
}

function Import-RouteDataFile {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path, [Parameter()][string]$Delimiter = '')

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Plik wejściowy nie istnieje: $Path"
    }

    $Extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    $RawRows = $null
    $Format = $null

    switch ($Extension) {
        { $_ -in '.xlsx', '.xls' } {
            $Format = 'Excel'
            if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
                throw "Wymagany moduł ImportExcel nie jest zainstalowany. Zainstaluj go poleceniem: Install-Module -Name ImportExcel -Scope CurrentUser"
            }
            Import-Module -Name ImportExcel -ErrorAction Stop
            $RawRows = @(Import-Excel -Path $Path)
        }
        { $_ -in '.csv', '.tsv', '.txt' } {
            $Format = 'CSV'
            $FirstLine = Get-Content -LiteralPath $Path -TotalCount 1
            $UsedDelimiter = if (-not [string]::IsNullOrWhiteSpace($Delimiter)) { $Delimiter }
            elseif ($Extension -eq '.tsv' -or $FirstLine -match "`t") { "`t" }
            elseif ($FirstLine -match ';') { ';' }
            else { ',' }
            $RawRows = @(Import-Csv -LiteralPath $Path -Delimiter $UsedDelimiter)
        }
        '.json' {
            $Format = 'JSON'
            $Content = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
            $ParsedJson = $Content | ConvertFrom-Json
            if ($ParsedJson -is [System.Collections.IEnumerable] -and -not ($ParsedJson -is [string])) {
                $RawRows = @($ParsedJson)
            }
            elseif ($ParsedJson.PSObject.Properties.Name -contains 'Routes') {
                $RawRows = @($ParsedJson.Routes)
            }
            elseif ($ParsedJson.PSObject.Properties.Name -contains 'Stops') {
                $RawRows = @($ParsedJson.Stops)
            }
            else {
                $RawRows = @($ParsedJson)
            }
        }
        default {
            throw "Nieobsługiwany format pliku: $Extension. Obsługiwane rozszerzenia: .xlsx, .xls, .csv, .tsv, .json"
        }
    }

    if ($null -eq $RawRows -or $RawRows.Count -eq 0) {
        return [PSCustomObject]@{
            Mode       = 'Empty'
            Routes     = @()
            RawData    = @()
            FilePath   = $Path
            Format     = $Format
            TotalCount = 0
        }
    }

    $PropNames = @($RawRows[0].PSObject.Properties.Name)

    # Sprawdzenie czy to sekwencja przystanków (SequentialStops)
    $ColSeq = Find-MatchingPropertyName -AvailableProperties $PropNames -Patterns @('^(lp|l\.p\.|kolejnosc|stop|sequence|order|nr)$')
    $ColAddrSeq = Find-MatchingPropertyName -AvailableProperties $PropNames -Patterns @('^(adres|address|lokalizacja|punkt|miejsce)$', 'lokalizacja.*(odbioru|dowozu)', 'adres.*(odbioru|dowozu)')
    $ColCitySeq = Find-MatchingPropertyName -AvailableProperties $PropNames -Patterns @('^(miejscowosc|miasto|city|town)$')

    $IsSequentialStops = ($ColSeq -and ($ColAddrSeq -or $ColCitySeq) -and -not (Find-MatchingPropertyName -AvailableProperties $PropNames -Patterns @('^(start|origin|adres.*a)$')))

    if ($IsSequentialStops) {
        $OrderedStops = @($RawRows | Sort-Object { [int]($_.$ColSeq) })
        $StopList = [System.Collections.Generic.List[PSCustomObject]]::new()
        foreach ($st in $OrderedStops) {
            $addr = if ($ColAddrSeq) { [string]$st.$ColAddrSeq } else { '' }
            $city = if ($ColCitySeq) { [string]$st.$ColCitySeq } else { '' }
            $fullAddr = if ($addr -and $city) { "$addr, $city" } elseif ($addr) { $addr } else { $city }
            $StopList.Add([PSCustomObject]@{
                Sequence = [string]$st.$ColSeq
                Address  = $fullAddr.Trim()
                Raw      = $st
            })
        }

        $RouteObj = $null
        if ($StopList.Count -ge 2) {
            $StartPoint = $StopList[0].Address
            $EndPoint = $StopList[$StopList.Count - 1].Address
            $Waypoints = if ($StopList.Count -gt 2) { @($StopList[1..($StopList.Count - 2)] | ForEach-Object { $_.Address }) } else { @() }
            $RouteObj = [PSCustomObject]@{
                Id          = '1'
                Name        = "Multi-point Route ($($StopList.Count) stops)"
                Start       = $StartPoint
                End         = $EndPoint
                Waypoints   = $Waypoints
                RouteType   = 'Fastest'
                OriginalRow = $OrderedStops
            }
        }

        $RoutesList = [System.Collections.Generic.List[PSCustomObject]]::new()
        if ($RouteObj) { $RoutesList.Add($RouteObj) }

        return [PSCustomObject]@{
            Mode       = 'SequentialStops'
            Stops      = $StopList
            Routes     = $RoutesList
            RawData    = $RawRows
            FilePath   = $Path
            Format     = $Format
            TotalCount = $StopList.Count
        }
    }

    # Tryb RouteList (wiersz = trasa)
    $ColStart = Find-MatchingPropertyName -AvailableProperties $PropNames -Patterns @(
        '(?i)^(start|origin|startpoint|poczat.*|poczatek|od|from|dom)$',
        '(?i)adres.*a|^a$',
        '(?i)punkt.*(poczat|start)'
    )
    $ColEnd = Find-MatchingPropertyName -AvailableProperties $PropNames -Patterns @(
        '(?i)^(end|dest|destination|endpoint|koniec.*|konic.*|cel|meta|do|to|szkola)$',
        '(?i)adres.*b|^b$',
        '(?i)punkt.*(konic|koniec|docel|cel)'
    )
    $ColWaypoints = Find-MatchingPropertyName -AvailableProperties $PropNames -Patterns @(
        '(?i)^(waypoints|waypoint|posredn.*|punkty.*posredn.*|przystank.*|via|stops|praca)$',
        '(?i)posrednie'
    )
    $ColName = Find-MatchingPropertyName -AvailableProperties $PropNames -Patterns @(
        '(?i)^(name|nazwa|umowa|contract|id|nr|opis|description|tytul)$',
        '(?i)numer.*umowy'
    )
    $ColRouteType = Find-MatchingPropertyName -AvailableProperties $PropNames -Patterns @(
        '(?i)^(routetype|typ|typtrasy|tryb|mode|optimization)$'
    )

    $NormalizedRoutes = [System.Collections.Generic.List[PSCustomObject]]::new()
    $idx = 1

    foreach ($row in $RawRows) {
        $startVal = if ($ColStart) { [string]$row.$ColStart } else { '' }
        $endVal   = if ($ColEnd) { [string]$row.$ColEnd } else { '' }
        if ([string]::IsNullOrWhiteSpace($startVal) -or [string]::IsNullOrWhiteSpace($endVal)) { continue }

        $nameVal = if ($ColName) { [string]$row.$ColName } else { "Route $idx" }
        $typeVal = if ($ColRouteType) { [string]$row.$ColRouteType } else { $null }

        if ($typeVal -match '(?i)eco|fuel|paliw|eko') { $typeVal = 'Eco' }
        elseif ($typeVal -match '(?i)short|krot|krót') { $typeVal = 'Shortest' }
        elseif ($typeVal -match '(?i)fast|szyb') { $typeVal = 'Fastest' }
        else { $typeVal = $null }

        $waypointsList = [System.Collections.Generic.List[string]]::new()
        if ($ColWaypoints -and -not [string]::IsNullOrWhiteSpace($row.$ColWaypoints)) {
            $rawWp = $row.$ColWaypoints
            if ($rawWp -is [System.Collections.IEnumerable] -and -not ($rawWp -is [string])) {
                foreach ($item in $rawWp) {
                    if (-not [string]::IsNullOrWhiteSpace([string]$item)) { $waypointsList.Add(([string]$item).Trim()) }
                }
            }
            else {
                $splits = ([string]$rawWp) -split '(?<!\\)[|;]'
                foreach ($s in $splits) {
                    $cleaned = $s.Trim()
                    if (-not [string]::IsNullOrWhiteSpace($cleaned)) { $waypointsList.Add($cleaned) }
                }
            }
        }

        $NormalizedRoutes.Add([PSCustomObject]@{
            Id          = [string]$idx
            Name        = $nameVal
            Start       = $startVal.Trim()
            End         = $endVal.Trim()
            Waypoints   = @($waypointsList)
            RouteType   = $typeVal
            OriginalRow = $row
        })
        $idx++
    }

    return [PSCustomObject]@{
        Mode       = 'RouteList'
        Routes     = $NormalizedRoutes
        RawData    = $RawRows
        FilePath   = $Path
        Format     = $Format
        TotalCount = $NormalizedRoutes.Count
        Columns    = [PSCustomObject]@{
            Start     = $ColStart
            End       = $ColEnd
            Waypoints = $ColWaypoints
            Name      = $ColName
            RouteType = $ColRouteType
        }
    }
}

function Export-RouteResults {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object[]]$Results,
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter()][ValidateSet('Excel', 'CSV', 'JSON')][string]$Format = 'Excel'
    )

    $TargetDir = Split-Path -Parent $OutputPath
    if (-not [string]::IsNullOrWhiteSpace($TargetDir) -and -not (Test-Path $TargetDir)) {
        New-Item -ItemType Directory -Path $TargetDir -Force | Out-Null
    }

    # Extract flat summary rows (excluding nested Points array from main sheet/file)
    $RoutesFlat = [System.Collections.Generic.List[PSCustomObject]]::new()
    $PointsFlat = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($r in $Results) {
        $routeId   = if ($null -ne $r.Id) { [string]$r.Id } else { '' }
        $routeName = if ($r.Name) { [string]$r.Name } elseif ($r.Nazwa) { [string]$r.Nazwa } else { "Route $routeId" }
        $startOrig = if ($r.Start) { [string]$r.Start } else { '' }
        $startGeo  = if ($r.StartGeocoded) { [string]$r.StartGeocoded } elseif ($r.StartGeokodowany) { [string]$r.StartGeokodowany } else { '' }
        $startStat = if ($r.StartStatus) { [string]$r.StartStatus } else { '' }
        $endOrig   = if ($r.End) { [string]$r.End } elseif ($r.Koniec) { [string]$r.Koniec } else { '' }
        $endGeo    = if ($r.EndGeocoded) { [string]$r.EndGeocoded } elseif ($r.KoniecGeokodowany) { [string]$r.KoniecGeokodowany } else { '' }
        $endStat   = if ($r.EndStatus) { [string]$r.EndStatus } else { '' }
        $wpCount   = if ($null -ne $r.WaypointsCount) { [int]$r.WaypointsCount } elseif ($null -ne $r.LiczbaPrzystankow) { [int]$r.LiczbaPrzystankow } else { 0 }
        $rType     = if ($r.RouteType) { [string]$r.RouteType } elseif ($r.TypTrasy) { [string]$r.TypTrasy } else { '' }
        $dist      = if ($null -ne $r.DistanceKm) { $r.DistanceKm } elseif ($null -ne $r.OdlegloscKm) { $r.OdlegloscKm } else { $null }
        $dur       = if ($null -ne $r.DurationMin) { $r.DurationMin } elseif ($null -ne $r.CzasMin) { $r.CzasMin } else { $null }
        $status    = if ($r.Status) { [string]$r.Status } else { '' }
        $map       = if ($r.MapPath) { [string]$r.MapPath } elseif ($r.MapaPath) { [string]$r.MapaPath } else { '' }
        $url       = if ($r.GoogleMapsUrl) { [string]$r.GoogleMapsUrl } else { '' }

        # Build waypoints summary text
        $wpSummaryList = [System.Collections.Generic.List[string]]::new()
        if ($r.Points -and ($r.Points -is [System.Collections.IEnumerable])) {
            foreach ($pt in $r.Points) {
                if ($pt.PointType -like 'Waypoint*') {
                    $ptSummary = "$($pt.PointType): '$($pt.OriginalAddress)'"
                    if ($pt.GeocodedAddress) { $ptSummary += " -> '$($pt.GeocodedAddress)'" }
                    if ($pt.GeocodeStatus) { $ptSummary += " [$($pt.GeocodeStatus)]" }
                    $wpSummaryList.Add($ptSummary)
                }

                $PointsFlat.Add([PSCustomObject]@{
                    RouteId         = $routeId
                    RouteName       = $routeName
                    PointOrder      = $pt.Order
                    PointType       = $pt.PointType
                    OriginalAddress = $pt.OriginalAddress
                    GeocodedAddress = $pt.GeocodedAddress
                    GeocodeStatus   = $pt.GeocodeStatus
                    MatchType       = $pt.MatchType
                    IsFallback      = if ($null -ne $pt.IsFallback) { [bool]$pt.IsFallback } else { $false }
                    Latitude        = $pt.Latitude
                    Longitude       = $pt.Longitude
                })
            }
        }

        $wpSummaryText = $wpSummaryList -join ' | '

        $RoutesFlat.Add([PSCustomObject]@{
            Id               = $routeId
            Name             = $routeName
            Start_Original   = $startOrig
            Start_Geocoded   = $startGeo
            Start_Status     = $startStat
            End_Original     = $endOrig
            End_Geocoded     = $endGeo
            End_Status       = $endStat
            WaypointsCount   = $wpCount
            RouteType        = $rType
            DistanceKm       = $dist
            DurationMin      = $dur
            Status           = $status
            WaypointsSummary = $wpSummaryText
            MapPath          = $map
            GoogleMapsUrl    = $url
        })
    }

    $csvEncoding = if ($PSVersionTable.PSVersion.Major -ge 7) { 'utf8BOM' } else { 'UTF8' }

    switch ($Format) {
        'Excel' {
            if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
                Write-Warning "Moduł ImportExcel nie jest zainstalowany. Eksportowanie do CSV zamiast Excel."
                $CsvPath = [System.IO.Path]::ChangeExtension($OutputPath, '.csv')
                $RoutesFlat | Export-Csv -LiteralPath $CsvPath -NoTypeInformation -Encoding $csvEncoding -Delimiter ';'
                if ($PointsFlat.Count -gt 0) {
                    $PtsCsv = [System.IO.Path]::Combine($TargetDir, "$([System.IO.Path]::GetFileNameWithoutExtension($CsvPath))_punkty.csv")
                    $PointsFlat | Export-Csv -LiteralPath $PtsCsv -NoTypeInformation -Encoding $csvEncoding -Delimiter ';'
                }
                return $CsvPath
            }
            Import-Module -Name ImportExcel -ErrorAction Stop
            if (Test-Path -LiteralPath $OutputPath) {
                Remove-Item -LiteralPath $OutputPath -Force -ErrorAction SilentlyContinue
            }
            $RoutesFlat | Export-Excel -Path $OutputPath -WorksheetName 'Trasy' -TableName 'WynikiTras' -AutoSize -AutoFilter -FreezeTopRow
            if ($PointsFlat.Count -gt 0) {
                $PointsFlat | Export-Excel -Path $OutputPath -WorksheetName 'PunktyTrasy' -TableName 'PunktyTrasy' -AutoSize -AutoFilter -FreezeTopRow
            }
            return $OutputPath
        }
        'CSV' {
            $RoutesFlat | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding $csvEncoding -Delimiter ';'
            if ($PointsFlat.Count -gt 0) {
                $PtsCsv = [System.IO.Path]::Combine($TargetDir, "$([System.IO.Path]::GetFileNameWithoutExtension($OutputPath))_punkty.csv")
                $PointsFlat | Export-Csv -LiteralPath $PtsCsv -NoTypeInformation -Encoding $csvEncoding -Delimiter ';'
            }
            return $OutputPath
        }
        'JSON' {
            $JsonContent = $Results | ConvertTo-Json -Depth 5
            [System.IO.File]::WriteAllText($OutputPath, $JsonContent, [System.Text.Encoding]::UTF8)
            return $OutputPath
        }
    }
}

# ══════════════════════════════════════════════════════════════════════════════
# 5. KONFIGURACJA I DPAPI SECURITY
# ══════════════════════════════════════════════════════════════════════════════

$script:AppDirName = 'GoogleMapsRoutes'
$script:LocalConfigFolder = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) $script:AppDirName
if (-not (Test-Path $script:LocalConfigFolder)) {
    New-Item -ItemType Directory -Path $script:LocalConfigFolder -Force | Out-Null
}
$script:ConfigFile = Join-Path $script:LocalConfigFolder 'config.json'
$script:LogFile    = Join-Path $script:LocalConfigFolder 'GoogleMapsRoutes.log'

# External Localization File resolution:
# Priority 1: $PSScriptRoot\localization.json
# Priority 2: EXE directory\localization.json (for compiled PS2EXE binaries)
# Priority 3: %LOCALAPPDATA%\GoogleMapsRoutes\localization.json
$script:ExeDir = try {
    Split-Path ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) -Parent
} catch { $null }

$script:LocalizationFile = if ($PSScriptRoot -and (Test-Path (Join-Path $PSScriptRoot 'localization.json'))) {
    Join-Path $PSScriptRoot 'localization.json'
} elseif ($script:ExeDir -and (Test-Path (Join-Path $script:ExeDir 'localization.json'))) {
    Join-Path $script:ExeDir 'localization.json'
} elseif (Test-Path (Join-Path $script:LocalConfigFolder 'localization.json')) {
    Join-Path $script:LocalConfigFolder 'localization.json'
} elseif ($PSScriptRoot) {
    Join-Path $PSScriptRoot 'localization.json'
} elseif ($script:ExeDir) {
    Join-Path $script:ExeDir 'localization.json'
} else {
    Join-Path $script:LocalConfigFolder 'localization.json'
}

function Load-LocalizationConfig {
    [CmdletBinding()]
    param()

    # If the file does not exist on disk, create default template
    if (-not (Test-Path $script:LocalizationFile)) {
        try {
            $defaultJson = @'
{
  "DefaultLanguage": "en",
  "Languages": {
    "en": {
      "DisplayName": "English",
      "GoogleCode": "en",
      "Strings": {
        "AppTitle": "Google Maps Route & Map Generator",
        "AppSubtitle": "Multi-point driving routes: Fastest, Shortest, Eco-friendly | Import JSON, CSV, Excel",
        "ApiBadgeChecking": "API: Checking...",
        "ApiBadgeActive": "API: Active",
        "ApiBadgeMissing": "API: Missing Key",
        "ApiBadgeError": "API: Error",
        "BtnQuickSettings": "⚙ API Settings",
        "FooterReady": "Ready.",
        "FooterVersion": "Google Maps Routes v2.0",
        "TabManual": "📍 Manual Route",
        "TabBatch": "📁 Batch File Processing",
        "TabSettings": "⚙ Settings & API Key",
        "ManualHeaderRoutePoints": "Route Points",
        "ManualOrigin": "Origin (Start / A):",
        "ManualWaypoints": "Intermediate Stops (optional up to 25):",
        "ManualWaypointsTooltip": "Enter waypoint address and click Add",
        "ManualBtnAdd": "➕ Add",
        "ManualBtnUp": "▲ Up",
        "ManualBtnDown": "▼ Down",
        "ManualBtnRemove": "✕ Remove",
        "ManualBtnClear": "🗑 Clear",
        "ManualDestination": "Destination (End / B):",
        "ManualRouteName": "Route Name / Description:",
        "ManualHeaderOptimization": "Route Optimization",
        "ManualOptFastest": "⚡ Fastest",
        "ManualOptShortest": "📏 Shortest",
        "ManualOptEco": "🌿 Eco",
        "ManualEmission": "Vehicle Engine Type (for Eco route):",
        "ManualFuelGasoline": "Gasoline",
        "ManualFuelDiesel": "Diesel",
        "ManualFuelHybrid": "Hybrid",
        "ManualFuelElectric": "Electric",
        "ManualTrafficAware": "Real-time traffic awareness (Live Traffic)",
        "ManualBtnCalculate": "🚀 CALCULATE ROUTE & DOWNLOAD MAP",
        "ManualBtnCalculating": "⏳ CALCULATING ROUTE...",
        "ManualStatDistance": "DISTANCE",
        "ManualStatDuration": "DURATION",
        "ManualStatType": "ROUTE TYPE",
        "ManualStatusIdle": "Idle",
        "ManualStatusCalculating": "Calculating...",
        "ManualStatusSuccess": "Route calculated",
        "ManualStatusError": "Calculation error",
        "ManualMapPlaceholder": "Map preview will appear here after route calculation...",
        "ManualNoUrl": "No generated link",
        "ManualBtnGoogleMaps": "🌐 Google Maps",
        "ManualBtnCopyUrl": "📋 Copy Link",
        "ManualBtnSaveMapAs": "💾 Save Map As...",
        "BatchInputFile": "Input File (JSON/CSV/XLSX):",
        "BatchBtnBrowse": "📂 Browse File...",
        "BatchBtnReload": "🔄 Reload",
        "BatchNoFileLoaded": "No file loaded.",
        "BatchDefaultRouteType": "Default route type:",
        "BatchOptFromSource": "From Source / Default",
        "BatchBtnStart": "▶ Start Processing",
        "BatchBtnStop": "⏹ Stop",
        "BatchTabInputPreview": "📋 Input Data Preview",
        "BatchTabResults": "📊 Calculation Results",
        "BatchTabLog": "📝 Activity Log",
        "BatchColId": "ID",
        "BatchColName": "Route Name",
        "BatchColOrigin": "Origin (Start)",
        "BatchColDestination": "Destination (End)",
        "BatchColWaypoints": "Waypoints",
        "BatchColType": "Type",
        "BatchColDistance": "Distance (km)",
        "BatchColDuration": "Duration (min)",
        "BatchColStatus": "Status",
        "BatchColMap": "PNG Map",
        "BatchProgressReady": "Ready",
        "BatchBtnOpenOutputDir": "📂 Open Output Folder",
        "BatchBtnExportExcel": "📊 Export Excel",
        "BatchBtnExportCsv": "📄 CSV",
        "BatchBtnExportJson": "📋 JSON",
        "SettingsHeaderApi": "Google Maps API Key",
        "SettingsApiDesc": "Required for Geocoding API, Routes API v2, and Static Maps API.",
        "SettingsApiLabel": "API Key:",
        "SettingsBtnShow": "👁 Show",
        "SettingsBtnHide": "🔒 Hide",
        "SettingsBtnTestKey": "🔍 Test Key",
        "SettingsChkRemember": "Remember securely on this computer (DPAPI CurrentUser encryption)",
        "SettingsHeaderPreferences": "Default Generation Preferences",
        "SettingsDefaultRouteType": "Default route type:",
        "SettingsDefaultEmission": "Default engine type for Eco routes:",
        "SettingsDefaultMapSize": "Default dimensions for generated PNG map:",
        "SettingsOutputDir": "Results Output Folder:",
        "SettingsBtnBrowseOutputDir": "📂 Browse...",
        "SettingsHeaderLanguage": "Language & Localization",
        "SettingsLanguageLabel": "Application & Google Maps API Language:",
        "SettingsBtnOpenLangFile": "📂 Open Localization File (localization.json)",
        "SettingsBtnReloadLang": "🔄 Reload Languages",
        "SettingsBtnSave": "💾 SAVE SETTINGS",
        "SettingsBtnOpenLog": "📋 OPEN LOG FILE",
        "MapLabelTotal": "Total: ",
        "MapLabelType": "Type: ",
        "MapLabelContract": "Contract: ",
        "MapLabelDirection": "Direction: ",
        "MsgMissingApiKey": "Please enter an API key before testing.",
        "MsgMissingApiKeyTitle": "Missing API Key",
        "MsgMissingApiKeyPrompt": "Please enter and save a Google Maps API key in Settings.",
        "MsgMissingData": "Please enter both an origin (start) and a destination (end).",
        "MsgMissingDataTitle": "Missing Data",
        "MsgNoDataFile": "Please load a valid data file first (JSON, CSV, or Excel).",
        "MsgNoDataFileTitle": "No Data",
        "MsgMaxWaypoints": "Maximum number of waypoints is 25.",
        "MsgMaxWaypointsTitle": "Waypoint Limit",
        "MsgSettingsSaved": "Settings have been saved successfully.",
        "MsgSettingsSavedTitle": "Saved",
        "MsgUrlCopied": "Google Maps navigation link copied to clipboard.",
        "MsgUrlCopiedTitle": "Copied",
        "MsgMapSaved": "Map saved: {0}",
        "MsgMapSavedTitle": "Saved",
        "MsgNoExportResults": "No results to export.",
        "MsgNoExportResultsTitle": "Empty Results",
        "MsgExportExcelComplete": "Exported to Excel:
{0}",
        "MsgExportCsvComplete": "Exported to CSV:
{0}",
        "MsgExportJsonComplete": "Exported to JSON:
{0}",
        "MsgExportTitle": "Export Complete",
        "MsgLangReloaded": "Language definitions reloaded successfully ({0} languages found).",
        "MsgLangReloadedTitle": "Languages Reloaded",
        "ThemeToggle": "Theme:",
        "ThemeDark": "🌙 Dark",
        "ThemeLight": "☀️ Light",
        "ThemeToggleTip": "Toggle Light / Dark theme",
        "SettingsThemeLabel": "Application Theme (Color Scheme):",
        "BatchTabPoints": "📍 Points Detail",
        "PointsColRouteId": "Route ID",
        "PointsColRouteName": "Route Name",
        "PointsColOrder": "No.",
        "PointsColType": "Point Type",
        "PointsColOriginalAddress": "Original Address",
        "PointsColGeocodedAddress": "Geocoded Address",
        "PointsColGeocodeStatus": "Geocode Status",
        "PointsColMatchType": "Match Type",
        "PointsColIsFallback": "Fallback?",
        "PointsColLatitude": "Latitude",
        "PointsColLongitude": "Longitude"
      }
    }
  }
}
'@
            $dir = Split-Path -Parent $script:LocalizationFile
            if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
            [System.IO.File]::WriteAllText($script:LocalizationFile, $defaultJson, [System.Text.UTF8Encoding]::new($true))
        } catch { }
    }

    # Auto-sync: If master localization.json exists in ExeDir or PSScriptRoot and is newer than active file, copy it
    $masterLoc = if ($script:ExeDir -and (Test-Path (Join-Path $script:ExeDir 'localization.json'))) {
        Join-Path $script:ExeDir 'localization.json'
    } elseif ($PSScriptRoot -and (Test-Path (Join-Path $PSScriptRoot 'localization.json'))) {
        Join-Path $PSScriptRoot 'localization.json'
    } else { $null }

    if ($masterLoc -and $script:LocalizationFile -ne $masterLoc -and (Test-Path $masterLoc)) {
        try {
            $masterTime = (Get-Item $masterLoc).LastWriteTimeUtc
            $activeTime = if (Test-Path $script:LocalizationFile) { (Get-Item $script:LocalizationFile).LastWriteTimeUtc } else { [DateTime]::MinValue }
            if ($masterTime -gt $activeTime) {
                Copy-Item -Path $masterLoc -Destination $script:LocalizationFile -Force
            }
        } catch { }
    }

    $script:LanguagesCatalog = [ordered]@{}
    $script:DefaultStrings = @{}

    if (Test-Path $script:LocalizationFile) {
        try {
            $raw = [System.IO.File]::ReadAllText($script:LocalizationFile, [System.Text.Encoding]::UTF8)
            $parsed = $raw | ConvertFrom-Json
            if ($parsed.Languages) {
                foreach ($prop in $parsed.Languages.PSObject.Properties) {
                    $code = $prop.Name.ToLower()
                    $langData = $prop.Value
                    $disp = if ($langData.DisplayName) { [string]$langData.DisplayName } else { $code.ToUpper() }
                    $gCode = if ($langData.GoogleCode) { [string]$langData.GoogleCode } else { $code }

                    $strMap = @{}
                    if ($langData.Strings) {
                        foreach ($sProp in $langData.Strings.PSObject.Properties) {
                            $strMap[$sProp.Name] = [string]$sProp.Value
                        }
                    }

                    $script:LanguagesCatalog[$code] = [PSCustomObject]@{
                        Code        = $code
                        DisplayName = $disp
                        GoogleCode  = $gCode
                        Strings     = $strMap
                    }
                }
            }
        } catch {
            Write-AppLog "Error parsing localization file $script:LocalizationFile : $($_.Exception.Message)" "WARN"
        }
    }

    # Ensure fallback EN exists
    if (-not $script:LanguagesCatalog.Contains('en')) {
        $script:LanguagesCatalog['en'] = [PSCustomObject]@{
            Code        = 'en'
            DisplayName = 'English'
            GoogleCode  = 'en'
            Strings     = @{ 'AppTitle' = 'Google Maps Route & Map Generator' }
        }
    }
    $script:DefaultStrings = $script:LanguagesCatalog['en'].Strings
}

function Get-LocText {
    param(
        [Parameter(Mandatory)][string]$Key,
        [Parameter()][string]$Default = $null
    )
    if ($script:CurrentStrings -and $script:CurrentStrings.ContainsKey($Key) -and -not [string]::IsNullOrWhiteSpace($script:CurrentStrings[$Key])) {
        return $script:CurrentStrings[$Key]
    }
    if ($script:DefaultStrings -and $script:DefaultStrings.ContainsKey($Key) -and -not [string]::IsNullOrWhiteSpace($script:DefaultStrings[$Key])) {
        return $script:DefaultStrings[$Key]
    }
    if ($Default) { return $Default }
    return $Key
}

function Get-MaskedKey([string]$Key) {
    if ([string]::IsNullOrWhiteSpace($Key)) { return '(brak)' }
    if ($Key.Length -le 8) { return '***' }
    return "$($Key.Substring(0, 4))...$($Key.Substring($Key.Length - 4, 4))"
}

function Write-AppLog {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'OK', 'WARN', 'ERROR', 'DEBUG')]
        [string]$Level = 'INFO',
        [switch]$ToBatchWindow
    )
    $now = Get-Date
    $timeStr = $now.ToString('yyyy-MM-dd HH:mm:ss.fff')
    $prefix = switch ($Level) {
        'OK'    { '[OK]   ' }
        'WARN'  { '[WARN] ' }
        'ERROR' { '[ERROR]' }
        'DEBUG' { '[DEBUG]' }
        default { '[INFO] ' }
    }
    $entry = "[$timeStr] $prefix $Message"

    try {
        [System.IO.File]::AppendAllText($script:LogFile, "$entry`r`n", [System.Text.UTF8Encoding]::new($true))
    } catch { }

    if ($ToBatchWindow -and $txtBatchLog) {
        try {
            $batchTime = $now.ToString('HH:mm:ss')
            $line = "$batchTime $prefix $Message`r`n"
            $txtBatchLog.Dispatcher.Invoke([Action]{
                $txtBatchLog.AppendText($line)
                $txtBatchLog.ScrollToEnd()
            })
        } catch { }
    }
}

Write-AppLog "================================================================================" "INFO"
Write-AppLog "Uruchomienie Google Maps Route & Map Generator v2.0" "INFO"
Write-AppLog "Środowisko: PowerShell $($PSVersionTable.PSVersion), OS: $([System.Environment]::OSVersion.VersionString)" "INFO"
Write-AppLog "Plik konfiguracji: $script:ConfigFile" "INFO"
Write-AppLog "Plik dziennika zdarzeń (log): $script:LogFile" "INFO"

$script:OverlayPropKeys = @('StartGeocoded', 'EndGeocoded', 'Distance', 'Duration', 'Timestamp', 'RouteName', 'RouteType', 'Waypoints', 'StartRaw', 'EndRaw')

function Get-DefaultOverlayConfig {
    return [ordered]@{
        EnableTopOverlay    = $true
        EnableBottomOverlay = $true
        Items               = [ordered]@{
            StartGeocoded = [ordered]@{ Enabled = $true;  Panel = 'Bottom'; Align = 'Left';   Order = 1 }
            EndGeocoded   = [ordered]@{ Enabled = $true;  Panel = 'Bottom'; Align = 'Left';   Order = 2 }
            Distance      = [ordered]@{ Enabled = $true;  Panel = 'Bottom'; Align = 'Left';   Order = 3 }
            Duration      = [ordered]@{ Enabled = $true;  Panel = 'Bottom'; Align = 'Center'; Order = 3 }
            Timestamp     = [ordered]@{ Enabled = $true;  Panel = 'Bottom'; Align = 'Right';  Order = 3 }
            RouteName     = [ordered]@{ Enabled = $true;  Panel = 'Top';    Align = 'Left';   Order = 1 }
            RouteType     = [ordered]@{ Enabled = $true;  Panel = 'Top';    Align = 'Right';  Order = 1 }
            Waypoints     = [ordered]@{ Enabled = $false; Panel = 'Bottom'; Align = 'Left';   Order = 2 }
            StartRaw      = [ordered]@{ Enabled = $false; Panel = 'None';   Align = 'Left';   Order = 1 }
            EndRaw        = [ordered]@{ Enabled = $false; Panel = 'None';   Align = 'Left';   Order = 2 }
        }
    }
}

function Load-AppConfig {
    $defaultResults = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'GoogleMapsRoutes\Results'
    $defaultOverlay = Get-DefaultOverlayConfig

    $cfg = [PSCustomObject]@{
        ApiKey           = ''
        RememberApiKey   = $true
        LastOutputFolder = $defaultResults
        LastInputFolder  = ''
        LastInputPath    = ''
        DefaultRouteType = 'Fastest'
        DefaultEmission  = 'GASOLINE'
        MapWidth         = 900
        MapHeight        = 600
        Language         = 'en'
        OverlayConfig    = $defaultOverlay
        Theme            = 'Dark'
    }

    if (Test-Path $script:ConfigFile) {
        try {
            $jsonText = [System.IO.File]::ReadAllText($script:ConfigFile, [System.Text.Encoding]::UTF8)
            $raw = $jsonText | ConvertFrom-Json
            if ($raw.ApiKeyEncrypted -is [string] -and -not [string]::IsNullOrWhiteSpace($raw.ApiKeyEncrypted)) {
                $dec = Unprotect-SecretString -EncryptedText $raw.ApiKeyEncrypted
                if (-not [string]::IsNullOrWhiteSpace($dec)) { $cfg.ApiKey = $dec }
            }
            elseif ($raw.ApiKey -is [string] -and -not [string]::IsNullOrWhiteSpace($raw.ApiKey)) {
                $dec = Unprotect-SecretString -EncryptedText $raw.ApiKey
                if (-not [string]::IsNullOrWhiteSpace($dec)) { $cfg.ApiKey = $dec }
            }

            if ($null -ne $raw.RememberApiKey) { $cfg.RememberApiKey = [bool]$raw.RememberApiKey }
            if ($raw.LastOutputFolder -is [string] -and -not [string]::IsNullOrWhiteSpace($raw.LastOutputFolder)) {
                $cfg.LastOutputFolder = $raw.LastOutputFolder
            }
            if ($raw.LastInputFolder -is [string] -and -not [string]::IsNullOrWhiteSpace($raw.LastInputFolder)) {
                $cfg.LastInputFolder = $raw.LastInputFolder
            }
            if ($raw.LastInputPath -is [string] -and -not [string]::IsNullOrWhiteSpace($raw.LastInputPath)) {
                $cfg.LastInputPath = $raw.LastInputPath
            }
            if ($raw.DefaultRouteType -is [string]) { $cfg.DefaultRouteType = $raw.DefaultRouteType }
            if ($raw.DefaultEmission -is [string]) { $cfg.DefaultEmission = $raw.DefaultEmission }
            if ($raw.MapWidth) { $cfg.MapWidth = [int]$raw.MapWidth }
            if ($raw.MapHeight) { $cfg.MapHeight = [int]$raw.MapHeight }
            if ($raw.Language -is [string] -and -not [string]::IsNullOrWhiteSpace($raw.Language)) {
                $cfg.Language = $raw.Language.Trim().ToLower()
            }
            if ($raw.Theme -is [string] -and $raw.Theme -in @('Dark', 'Light')) {
                $cfg.Theme = $raw.Theme
            }

            if ($raw.OverlayConfig) {
                if ($null -ne $raw.OverlayConfig.EnableTopOverlay) {
                    $defaultOverlay.EnableTopOverlay = [bool]$raw.OverlayConfig.EnableTopOverlay
                }
                if ($null -ne $raw.OverlayConfig.EnableBottomOverlay) {
                    $defaultOverlay.EnableBottomOverlay = [bool]$raw.OverlayConfig.EnableBottomOverlay
                }
                if ($raw.OverlayConfig.Items) {
                    foreach ($k in $script:OverlayPropKeys) {
                        $rawItem = if ($raw.OverlayConfig.Items.PSObject.Properties[$k]) {
                            $raw.OverlayConfig.Items.$k
                        } elseif ($raw.OverlayConfig.Items[$k]) {
                            $raw.OverlayConfig.Items[$k]
                        } else { $null }

                        if ($rawItem) {
                            $en = if ($null -ne $rawItem.Enabled) { [bool]$rawItem.Enabled } else { $defaultOverlay.Items[$k].Enabled }
                            $pn = if ($rawItem.Panel) { [string]$rawItem.Panel } else { $defaultOverlay.Items[$k].Panel }
                            $al = if ($rawItem.Align) { [string]$rawItem.Align } else { $defaultOverlay.Items[$k].Align }
                            $od = if ($rawItem.Order) { [int]$rawItem.Order } else { $defaultOverlay.Items[$k].Order }
                            $defaultOverlay.Items[$k] = [ordered]@{ Enabled = $en; Panel = $pn; Align = $al; Order = $od }
                        }
                    }
                }
            }
            $cfg.OverlayConfig = $defaultOverlay
        } catch { }
    }

    if ([string]::IsNullOrWhiteSpace($cfg.ApiKey) -and -not [string]::IsNullOrWhiteSpace($env:GOOGLE_MAPS_API_KEY)) {
        $cfg.ApiKey = $env:GOOGLE_MAPS_API_KEY
    }

    return $cfg
}

function Save-AppConfig {
    param(
        [string]$ApiKey,
        [bool]$RememberApiKey,
        [string]$OutputFolder,
        [string]$LastInputFolder = '',
        [string]$LastInputPath = '',
        [string]$DefaultRouteType = 'Fastest',
        [string]$DefaultEmission = 'GASOLINE',
        [int]$MapWidth = 900,
        [int]$MapHeight = 600,
        [string]$Language = '',
        [object]$OverlayConfig = $null,
        [string]$Theme = ''
    )
    $encKey = ''
    if ($RememberApiKey -and -not [string]::IsNullOrWhiteSpace($ApiKey)) {
        try {
            $protected = Protect-SecretString -PlainText $ApiKey
            if ($protected -is [string] -and -not [string]::IsNullOrWhiteSpace($protected)) {
                $encKey = $protected
            }
        } catch {
            $encKey = ''
        }
    }

    $finalInputFolder = if ($LastInputFolder) { $LastInputFolder } elseif ($script:Config -and $script:Config.LastInputFolder) { $script:Config.LastInputFolder } else { '' }
    $finalInputPath   = if ($LastInputPath) { $LastInputPath } elseif ($script:Config -and $script:Config.LastInputPath) { $script:Config.LastInputPath } else { '' }
    $finalLang        = if ($Language) { $Language } elseif ($script:Config -and $script:Config.Language) { $script:Config.Language } else { 'en' }
    $finalOverlay     = if ($OverlayConfig) { $OverlayConfig } elseif ($script:Config -and $script:Config.OverlayConfig) { $script:Config.OverlayConfig } else { Get-DefaultOverlayConfig }
    $finalTheme       = if ($Theme -in @('Dark', 'Light')) { $Theme } elseif ($script:Config -and $script:Config.Theme) { $script:Config.Theme } else { 'Dark' }

    $cfg = [ordered]@{
        ApiKeyEncrypted  = $encKey
        RememberApiKey   = $RememberApiKey
        LastOutputFolder = $OutputFolder
        LastInputFolder  = $finalInputFolder
        LastInputPath    = $finalInputPath
        DefaultRouteType = $DefaultRouteType
        DefaultEmission  = $DefaultEmission
        MapWidth         = $MapWidth
        MapHeight        = $MapHeight
        Language         = $finalLang
        OverlayConfig    = $finalOverlay
        Theme            = $finalTheme
    }
    $json = $cfg | ConvertTo-Json -Depth 6
    [System.IO.File]::WriteAllText($script:ConfigFile, $json, [System.Text.UTF8Encoding]::new($true))
}

function Get-CurrentOverlayConfig {
    $topEn = if ($chkEnableTopOverlay) { [bool]$chkEnableTopOverlay.IsChecked } else { $true }
    $btmEn = if ($chkEnableBottomOverlay) { [bool]$chkEnableBottomOverlay.IsChecked } else { $true }
    $cfg = [ordered]@{
        EnableTopOverlay    = $topEn
        EnableBottomOverlay = $btmEn
        Items               = [ordered]@{}
    }
    foreach ($key in $script:OverlayPropKeys) {
        $chk  = Get-Variable -Name "chkProp_$key"  -ValueOnly -ErrorAction SilentlyContinue
        $cmbP = Get-Variable -Name "cmbPanel_$key" -ValueOnly -ErrorAction SilentlyContinue
        $cmbA = Get-Variable -Name "cmbAlign_$key" -ValueOnly -ErrorAction SilentlyContinue
        $cmbO = Get-Variable -Name "cmbOrder_$key" -ValueOnly -ErrorAction SilentlyContinue

        $enabled = if ($chk) { [bool]$chk.IsChecked } else { $true }
        $panel   = if ($cmbP -and $cmbP.SelectedItem) { [string]$cmbP.SelectedItem.Tag } else { 'Bottom' }
        $align   = if ($cmbA -and $cmbA.SelectedItem) { [string]$cmbA.SelectedItem.Tag } else { 'Left' }
        $order   = if ($cmbO -and $cmbO.SelectedItem) { [int]$cmbO.SelectedItem.Tag } else { 1 }

        $cfg.Items[$key] = [ordered]@{
            Enabled = $enabled
            Panel   = $panel
            Align   = $align
            Order   = $order
        }
    }
    return $cfg
}

function Set-OverlayConfigUi($cfg) {
    if (-not $cfg) { return }
    if ($null -ne $cfg.EnableTopOverlay -and $chkEnableTopOverlay) {
        $chkEnableTopOverlay.IsChecked = [bool]$cfg.EnableTopOverlay
    }
    if ($null -ne $cfg.EnableBottomOverlay -and $chkEnableBottomOverlay) {
        $chkEnableBottomOverlay.IsChecked = [bool]$cfg.EnableBottomOverlay
    }
    if ($cfg.Items) {
        foreach ($key in $script:OverlayPropKeys) {
            $itemCfg = if ($cfg.Items.PSObject.Properties[$key]) {
                $cfg.Items.$key
            } elseif ($cfg.Items[$key]) {
                $cfg.Items[$key]
            } else { $null }

            if (-not $itemCfg) { continue }

            $chk  = Get-Variable -Name "chkProp_$key"  -ValueOnly -ErrorAction SilentlyContinue
            $cmbP = Get-Variable -Name "cmbPanel_$key" -ValueOnly -ErrorAction SilentlyContinue
            $cmbA = Get-Variable -Name "cmbAlign_$key" -ValueOnly -ErrorAction SilentlyContinue
            $cmbO = Get-Variable -Name "cmbOrder_$key" -ValueOnly -ErrorAction SilentlyContinue

            if ($chk -and $null -ne $itemCfg.Enabled) {
                $chk.IsChecked = [bool]$itemCfg.Enabled
            }
            if ($cmbP -and $itemCfg.Panel) {
                foreach ($opt in $cmbP.Items) {
                    if ($opt.Tag -eq $itemCfg.Panel) { $cmbP.SelectedItem = $opt; break }
                }
            }
            if ($cmbA -and $itemCfg.Align) {
                foreach ($opt in $cmbA.Items) {
                    if ($opt.Tag -eq $itemCfg.Align) { $cmbA.SelectedItem = $opt; break }
                }
            }
            if ($cmbO -and $itemCfg.Order) {
                foreach ($opt in $cmbO.Items) {
                    if ([int]$opt.Tag -eq [int]$itemCfg.Order) { $cmbO.SelectedItem = $opt; break }
                }
            }
        }
    }
}

function Reset-OverlayConfigUi {
    $defaultCfg = Get-DefaultOverlayConfig
    Set-OverlayConfigUi $defaultCfg
}

$script:Config = Load-AppConfig
Load-LocalizationConfig
$script:CurrentLanguage = if ($script:Config -and $script:Config.Language -and $script:LanguagesCatalog.Contains($script:Config.Language.ToLower())) {
    $script:Config.Language.ToLower()
} else {
    'en'
}
$script:CurrentGoogleLang = if ($script:LanguagesCatalog.Contains($script:CurrentLanguage)) {
    $script:LanguagesCatalog[$script:CurrentLanguage].GoogleCode
} else {
    'en'
}
$script:CurrentStrings = if ($script:LanguagesCatalog.Contains($script:CurrentLanguage)) {
    $script:LanguagesCatalog[$script:CurrentLanguage].Strings
} else {
    @{}
}
Write-AppLog "Active language: $script:CurrentLanguage (Google API code: $script:CurrentGoogleLang)" "INFO" 

# ══════════════════════════════════════════════════════════════════════════════
# 6. FABRYKA BEZPIECZNYCH WĄTKÓW TŁA (INITIALSESSIONSTATE RUNSPACE)
# ══════════════════════════════════════════════════════════════════════════════

function New-WorkerPowerShell {
    param([scriptblock]$ScriptBlock)
    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    Get-ChildItem function: | Where-Object {
        $_.Name -in @('Protect-SecretString', 'Unprotect-SecretString', 'Test-GoogleApiKey',
                      'Get-AddressComponentValue', 'Get-AddressCoordinates', 'Get-GeocodeStatusDescription', 'Get-CarRouteData',
                      'Get-GoogleMapsUrl', 'Get-WrappedLines', 'Save-RouteMapPng',
                      'Find-MatchingPropertyName', 'Import-RouteDataFile', 'Export-RouteResults')
    } | ForEach-Object {
        try {
            $iss.Commands.Add([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new($_.Name, $_.Definition))
        } catch { }
    }
    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($iss)
    $rs.ApartmentState = [System.Threading.ApartmentState]::MTA
    $rs.ThreadOptions  = [System.Management.Automation.Runspaces.PSThreadOptions]::UseNewThread
    $rs.Open()
    $ps = [PowerShell]::Create()
    $ps.Runspace = $rs
    # WAŻNE: [void] lub $null = zapobiega wyciekowi obiektu PowerShell do pipeline funkcji.
    # Bez tego funkcja zwraca tablicę @($ps, $ps), co przy wywołaniu .BeginInvoke()
    # powoduje próbę ponownego uruchomienia tej samej instancji i błąd:
    # "The operation cannot be performed because a command has already been started."
    $null = $ps.AddScript($ScriptBlock.ToString())
    return $ps
}

# ══════════════════════════════════════════════════════════════════════════════
# ══════════════════════════════════════════════════════════════════════════════
# 6b. MANUAL CALC WORKER SCRIPTBLOCK (Isolated Runspace, Top-level)
# ══════════════════════════════════════════════════════════════════════════════
$script:ManualCalcAsync = {
    param($start, $end, $waypoints, $routeType, $emission, $trafficAware, $name, $apiKey, $outDir, $logFile, $languageCode = 'en', $overlayConfigJson = '')

    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls
    Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

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
        if ($geoStart.Status -ne 'OK') {
            & $wlog "Origin geocoding error: $($geoStart.Status)" "WARN"
            return [PSCustomObject]@{ Success = $false; Error = "Origin geocoding error: $($geoStart.Status)" }
        }
        & $wlog "Origin OK: $($geoStart.FormattedAddress) ($($geoStart.Latitude), $($geoStart.Longitude))" "INFO"

        & $wlog "Geocoding destination: '$end'..." "INFO"
        $geoEnd = Get-AddressCoordinates -Address $end -ApiKey $apiKey -LanguageCode $languageCode
        if ($geoEnd.Status -ne 'OK') {
            & $wlog "Destination geocoding error: $($geoEnd.Status)" "WARN"
            return [PSCustomObject]@{ Success = $false; Error = "Destination geocoding error: $($geoEnd.Status)" }
        }
        & $wlog "Destination OK: $($geoEnd.FormattedAddress) ($($geoEnd.Latitude), $($geoEnd.Longitude))" "INFO"

        $geoWp = [System.Collections.Generic.List[PSCustomObject]]::new()
        foreach ($w in $waypoints) {
            & $wlog "Geocoding waypoint: '$w'..." "INFO"
            $g = Get-AddressCoordinates -Address $w -ApiKey $apiKey -LanguageCode $languageCode
            if ($g.Status -eq 'OK') {
                $geoWp.Add($g)
                & $wlog "Waypoint OK: $($g.FormattedAddress)" "INFO"
            } else {
                & $wlog "Waypoint geocoding error '$w': $($g.Status)" "WARN"
            }
        }

        & $wlog "Querying Google Routes API v2 (Type: $routeType, Engine: $emission)..." "INFO"
        $trasa = Get-CarRouteData -OriginLat $geoStart.Latitude -OriginLng $geoStart.Longitude `
            -DestLat $geoEnd.Latitude -DestLng $geoEnd.Longitude `
            -IntermediatePoints $geoWp -RouteType $routeType -EmissionType $emission `
            -ApiKey $apiKey -LanguageCode $languageCode -TrafficAware:$trafficAware

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

        & $wlog "Map rendering complete. Saved: $saved" "INFO"

        $resolvedMapPath = $(if ($saved) { $mapPath } else { $null })
        return [PSCustomObject]@{
            Success       = $true
            DistanceKm    = $trasa.OdlegloscKm
            DurationMin   = $trasa.CzasMin
            RouteType     = $routeType
            GoogleMapsUrl = $gUrl
            MapPath       = $resolvedMapPath
            Error         = $null
        }
    }
    catch {
        $errFull = $_.Exception.ToString()
        & $wlog "Worker thread exception: $errFull" "ERROR"
        return [PSCustomObject]@{ Success = $false; Error = $_.Exception.Message }
    }
}

# ══════════════════════════════════════════════════════════════════════════════
# 6c. BATCH CALC WORKER SCRIPTBLOCK (Isolated Runspace, Top-level)
# ══════════════════════════════════════════════════════════════════════════════
$script:BatchCalcAsync = {
    param($routes, $apiKey, $outDir, $defaultRouteType, $syncState, $logFile, $languageCode = 'en', $overlayConfigJson = '')

    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls
    Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

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

        $routeName = if ($r.Name) { $r.Name } else { "Route $($i + 1)" }

        & $wlog "Route $($i + 1)/$($total): Processing '$($r.Start)' -> '$($r.End)' (Type: $rType)..." "INFO"

        try {
            $geoStart = Get-AddressCoordinates -Address $r.Start -ApiKey $apiKey -LanguageCode $languageCode
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

            $geoEnd   = Get-AddressCoordinates -Address $r.End -ApiKey $apiKey -LanguageCode $languageCode
            $endStatus = Get-GeocodeStatusDescription -Geo $geoEnd
            $isEndFallback = if ($geoEnd -and ($geoEnd.PartialMatch -or $geoEnd.MatchType -in 'APPROXIMATE', 'GEOMETRIC_CENTER')) { $true } else { $false }

            if ($geoStart.Status -ne 'OK' -or $geoEnd.Status -ne 'OK') {
                $errReason = "Geocoding failed (Start=$($geoStart.Status), End=$($geoEnd.Status))"
                & $wlog "Route $($i + 1)/$($total): $errReason" "WARN"

                $routePointsList.Add([PSCustomObject]@{
                    Order           = 2
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

                $results.Add([PSCustomObject]@{
                    Id                = [string]($i + 1)
                    Name              = $routeName
                    Nazwa             = $routeName
                    Start             = $r.Start
                    StartGeocoded     = if ($geoStart) { $geoStart.FormattedAddress } else { $null }
                    StartStatus       = $startStatus
                    End               = $r.End
                    EndGeocoded       = if ($geoEnd) { $geoEnd.FormattedAddress } else { $null }
                    EndStatus         = $endStatus
                    Koniec            = $r.End
                    WaypointsCount    = 0
                    LiczbaPrzystankow = 0
                    RouteType         = $rType
                    TypTrasy          = $rType
                    DistanceKm        = $null
                    OdlegloscKm       = $null
                    DurationMin       = $null
                    CzasMin           = $null
                    Status            = $errReason
                    MapPath           = $null
                    MapaPath          = $null
                    Points            = @($routePointsList)
                })
                $syncState.FailCount++
                continue
            }

            $geoWp = [System.Collections.Generic.List[PSCustomObject]]::new()
            $wpIdx = 1
            if ($r.Waypoints) {
                foreach ($w in $r.Waypoints) {
                    if ([string]::IsNullOrWhiteSpace($w)) { continue }
                    $g = Get-AddressCoordinates -Address $w -ApiKey $apiKey -LanguageCode $languageCode
                    $wpStatus = Get-GeocodeStatusDescription -Geo $g
                    $isWpFallback = if ($g -and ($g.PartialMatch -or $g.MatchType -in 'APPROXIMATE', 'GEOMETRIC_CENTER')) { $true } else { $false }

                    $routePointsList.Add([PSCustomObject]@{
                        Order           = ($wpIdx + 1)
                        PointType       = "Waypoint $wpIdx"
                        OriginalAddress = $w
                        GeocodedAddress = if ($g) { $g.FormattedAddress } else { $null }
                        GeocodeStatus   = $wpStatus
                        MatchType       = if ($g) { $g.MatchType } else { 'NOT_FOUND' }
                        PartialMatch    = if ($g) { [bool]$g.PartialMatch } else { $false }
                        IsFallback      = $isWpFallback
                        Latitude        = if ($g) { $g.Latitude } else { $null }
                        Longitude       = if ($g) { $g.Longitude } else { $null }
                    })

                    if ($g -and $g.Status -eq 'OK' -and $null -ne $g.Latitude -and $null -ne $g.Longitude) {
                        $geoWp.Add($g)
                    } else {
                        & $wlog "Route $($i + 1)/$($total): Waypoint '$w' cannot be located ($wpStatus). Proceeding without it in driving directions." "WARN"
                    }
                    $wpIdx++
                    Start-Sleep -Milliseconds 60
                }
            }

            # Add End point to structured points
            $routePointsList.Add([PSCustomObject]@{
                Order           = ($routePointsList.Count + 1)
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

            & $wlog "Route $($i + 1)/$($total): Querying Google Routes API..." "INFO"
            $trasa = Get-CarRouteData -OriginLat $geoStart.Latitude -OriginLng $geoStart.Longitude `
                -DestLat $geoEnd.Latitude -DestLng $geoEnd.Longitude `
                -IntermediatePoints $geoWp -RouteType $rType -ApiKey $apiKey `
                -LanguageCode $languageCode

            if ($trasa.Status -ne 'OK') {
                $errReason = "Routes API: $($trasa.Status). $($trasa.ErrorMessage)"
                & $wlog "Route $($i + 1)/$($total): $errReason" "WARN"
                $results.Add([PSCustomObject]@{
                    Id                = [string]($i + 1)
                    Name              = $routeName
                    Nazwa             = $routeName
                    Start             = $r.Start
                    StartGeocoded     = if ($geoStart) { $geoStart.FormattedAddress } else { $null }
                    StartStatus       = $startStatus
                    End               = $r.End
                    EndGeocoded       = if ($geoEnd) { $geoEnd.FormattedAddress } else { $null }
                    EndStatus         = $endStatus
                    Koniec            = $r.End
                    WaypointsCount    = $geoWp.Count
                    LiczbaPrzystankow = $geoWp.Count
                    RouteType         = $rType
                    TypTrasy          = $rType
                    DistanceKm        = $null
                    OdlegloscKm       = $null
                    DurationMin       = $null
                    CzasMin           = $null
                    Status            = $errReason
                    MapPath           = $null
                    MapaPath          = $null
                    Points            = @($routePointsList)
                })
                $syncState.FailCount++
                continue
            }

            $safeName = ($routeName -replace '[\\/:*?"<>|]', '_').Trim()
            $mapPath = Join-Path $outDir "${ts}_route_$($i + 1)_${safeName}.png"

            $allPts = [System.Collections.Generic.List[PSCustomObject]]::new()
            $allPts.Add($geoStart)
            foreach ($wp in $geoWp) { $allPts.Add($wp) }
            $allPts.Add($geoEnd)

            & $wlog "Route $($i + 1)/$($total): Route OK ($($trasa.OdlegloscKm) km, $($trasa.CzasMin) min). Rendering static map..." "INFO"

            $hdrBatchPrefix = switch ($languageCode) { 'de' { 'Typ: ' } 'pl' { 'Typ: ' } default { 'Type: ' } }
            $hdrBatchName = switch ($languageCode) {
                'de' { if ($rType -eq 'Fastest') { 'Schnellste' } elseif ($rType -eq 'Shortest') { 'Kürzeste' } else { 'Eco' } }
                'pl' { if ($rType -eq 'Fastest') { 'Najszybsza' } elseif ($rType -eq 'Shortest') { 'Najkrótsza' } else { 'Eko' } }
                default { $rType }
            }
            $hdrBatchRightText = "$hdrBatchPrefix$hdrBatchName"

            $saved = Save-RouteMapPng -EncodedPolyline $trasa.EncodedPolyline `
                -OriginLat $geoStart.Latitude -OriginLng $geoStart.Longitude `
                -DestLat $geoEnd.Latitude -DestLng $geoEnd.Longitude `
                -RoutePoints $allPts -OutputPath $mapPath -ApiKey $apiKey `
                -AddressTextA $geoStart.FormattedAddress -AddressTextB $geoEnd.FormattedAddress `
                -DistanceText "$($trasa.OdlegloscKm) km" -DurationText "$($trasa.CzasMin) min" `
                -HeaderLeftText $routeName -HeaderRightText $hdrBatchRightText `
                -LanguageCode $languageCode `
                -StartRaw $r.Start -StartGeocoded $geoStart.FormattedAddress `
                -EndRaw $r.End -EndGeocoded $geoEnd.FormattedAddress `
                -WaypointsList $geoWp -RouteName $routeName -RouteType $hdrBatchRightText `
                -OverlayConfig $overlayConfigJson

            $resolvedMapPath = if ($saved -and (Test-Path $mapPath)) { $mapPath } else { $null }

            $results.Add([PSCustomObject]@{
                Id                = [string]($i + 1)
                Name              = $routeName
                Nazwa             = $routeName
                Start             = $r.Start
                StartGeocoded     = if ($geoStart) { $geoStart.FormattedAddress } else { $null }
                StartStatus       = $startStatus
                End               = $r.End
                EndGeocoded       = if ($geoEnd) { $geoEnd.FormattedAddress } else { $null }
                EndStatus         = $endStatus
                Koniec            = $r.End
                WaypointsCount    = $geoWp.Count
                LiczbaPrzystankow = $geoWp.Count
                RouteType         = $rType
                TypTrasy          = $rType
                DistanceKm        = $trasa.OdlegloscKm
                OdlegloscKm       = $trasa.OdlegloscKm
                DurationMin       = $trasa.CzasMin
                CzasMin           = $trasa.CzasMin
                Status            = 'OK'
                MapPath           = $resolvedMapPath
                MapaPath          = $resolvedMapPath
                Points            = @($routePointsList)
            })
            $syncState.SuccessCount++
            & $wlog "Route $($i + 1)/$($total): Complete! Map saved: $(Split-Path $mapPath -Leaf)" "OK"
        }
        catch {
            $errDetail = $_.Exception.Message
            & $wlog "Route $($i + 1)/$($total): Exception: $errDetail" "ERROR"
            $results.Add([PSCustomObject]@{
                Id                = [string]($i + 1)
                Name              = $routeName
                Nazwa             = $routeName
                Start             = $r.Start
                StartGeocoded     = if ($geoStart) { $geoStart.FormattedAddress } else { $null }
                StartStatus       = if ($startStatus) { $startStatus } else { 'EXCEPTION' }
                End               = $r.End
                EndGeocoded       = if ($geoEnd) { $geoEnd.FormattedAddress } else { $null }
                EndStatus         = if ($endStatus) { $endStatus } else { 'EXCEPTION' }
                Koniec            = $r.End
                WaypointsCount    = 0
                LiczbaPrzystankow = 0
                RouteType         = $rType
                TypTrasy          = $rType
                DistanceKm        = $null
                OdlegloscKm       = $null
                DurationMin       = $null
                CzasMin           = $null
                Status            = "Exception: $errDetail"
                MapPath           = $null
                MapaPath          = $null
                Points            = if ($routePointsList) { @($routePointsList) } else { @() }
            })
            $syncState.FailCount++
        }

        Start-Sleep -Milliseconds 100
    }

    return $results
}

# ══════════════════════════════════════════════════════════════════════════════
# 7. DEFINICJA INTERFEJSU WPF XAML (MODERN DARK THEME)
# ══════════════════════════════════════════════════════════════════════════════

[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Google Maps Route &amp; Map Generator"
    Height="880" Width="1100"
    MinHeight="700" MinWidth="900"
    WindowStartupLocation="CenterScreen"
    Background="{DynamicResource BgDark}" Foreground="{DynamicResource TextPrimary}"
    FontFamily="Segoe UI Variable, Segoe UI, sans-serif">

    <Window.Resources>
        <SolidColorBrush x:Key="BgDark" Color="#0F172A"/>
        <SolidColorBrush x:Key="BgCard" Color="#1E293B"/>
        <SolidColorBrush x:Key="BgCardHover" Color="#293548"/>
        <SolidColorBrush x:Key="BgCardAlt" Color="#162032"/>
        <SolidColorBrush x:Key="BorderCard" Color="#334155"/>
        <SolidColorBrush x:Key="TextPrimary" Color="#F8FAFC"/>
        <SolidColorBrush x:Key="TextSecondary" Color="#94A3B8"/>
        <SolidColorBrush x:Key="AccentBlue" Color="#2563EB"/>
        <SolidColorBrush x:Key="AccentGreen" Color="#10B981"/>
        <SolidColorBrush x:Key="AccentAmber" Color="#F59E0B"/>
        <SolidColorBrush x:Key="AccentRed" Color="#EF4444"/>
        <SolidColorBrush x:Key="BgInput" Color="#1E293B"/>
        <SolidColorBrush x:Key="BorderInput" Color="#334155"/>
        <SolidColorBrush x:Key="BtnSecondaryBg" Color="#334155"/>
        <SolidColorBrush x:Key="BtnSecondaryFg" Color="#F8FAFC"/>
        <SolidColorBrush x:Key="GridLines" Color="#2D3748"/>
        <SolidColorBrush x:Key="LogBg" Color="#0A0F1D"/>
        <SolidColorBrush x:Key="LogFg" Color="#38BDF8"/>
        <SolidColorBrush x:Key="DataGridHeaderBg" Color="#0F172A"/>
        <SolidColorBrush x:Key="DataGridHeaderFg" Color="#94A3B8"/>
        <SolidColorBrush x:Key="DataGridRowBg" Color="#1E293B"/>
        <SolidColorBrush x:Key="DataGridAltRowBg" Color="#162032"/>

        <Style TargetType="TextBox">
            <Setter Property="Background" Value="{DynamicResource BgInput}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderInput}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="9,6"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>

        <Style TargetType="PasswordBox">
            <Setter Property="Background" Value="{DynamicResource BgInput}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderInput}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="9,6"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>

        <Style TargetType="Button">
            <Setter Property="Background" Value="#2563EB"/>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="14,8"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="5"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Modern Dynamic ComboBox with Dropdown Popup Template -->
        <ControlTemplate x:Key="ComboBoxToggleButtonTemplate" TargetType="ToggleButton">
            <Border x:Name="TemplateRoot" Background="{TemplateBinding Background}" BorderBrush="{DynamicResource BorderInput}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="5">
                <Border x:Name="SplitBorder" Width="26" HorizontalAlignment="Right" Background="Transparent">
                    <Path x:Name="Arrow" HorizontalAlignment="Center" VerticalAlignment="Center" Fill="{DynamicResource TextSecondary}" Data="M 0 0 L 4 4 L 8 0 Z"/>
                </Border>
            </Border>
            <ControlTemplate.Triggers>
                <Trigger Property="IsMouseOver" Value="true">
                    <Setter TargetName="TemplateRoot" Property="BorderBrush" Value="{DynamicResource AccentBlue}"/>
                    <Setter TargetName="Arrow" Property="Fill" Value="{DynamicResource TextPrimary}"/>
                </Trigger>
                <Trigger Property="IsChecked" Value="true">
                    <Setter TargetName="TemplateRoot" Property="BorderBrush" Value="{DynamicResource AccentBlue}"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="false">
                    <Setter TargetName="TemplateRoot" Property="Opacity" Value="0.5"/>
                </Trigger>
            </ControlTemplate.Triggers>
        </ControlTemplate>

        <Style TargetType="ComboBox">
            <Setter Property="Background" Value="{DynamicResource BgInput}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderInput}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="ScrollViewer.HorizontalScrollBarVisibility" Value="Auto"/>
            <Setter Property="ScrollViewer.VerticalScrollBarVisibility" Value="Auto"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBox">
                        <Grid x:Name="MainGrid" SnapsToDevicePixels="true">
                            <ToggleButton x:Name="ToggleButton"
                                          Template="{StaticResource ComboBoxToggleButtonTemplate}"
                                          Background="{TemplateBinding Background}"
                                          BorderBrush="{TemplateBinding BorderBrush}"
                                          BorderThickness="{TemplateBinding BorderThickness}"
                                          Focusable="false"
                                          IsChecked="{Binding Path=IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"
                                          ClickMode="Press"/>
                            <ContentPresenter x:Name="ContentSite"
                                              IsHitTestVisible="false"
                                              Content="{TemplateBinding SelectionBoxItem}"
                                              ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                                              ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}"
                                              Margin="{TemplateBinding Padding}"
                                              VerticalAlignment="Center"
                                              HorizontalAlignment="Left"/>
                            <Popup x:Name="Popup"
                                   Placement="Bottom"
                                   IsOpen="{TemplateBinding IsDropDownOpen}"
                                   AllowsTransparency="true"
                                   Focusable="false"
                                   PopupAnimation="Slide">
                                <Grid x:Name="DropDown" SnapsToDevicePixels="true" MinWidth="{TemplateBinding ActualWidth}" MaxHeight="{TemplateBinding MaxDropDownHeight}">
                                    <Border x:Name="DropDownBorder" Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="5" Margin="0,2,0,0">
                                        <ScrollViewer Margin="2" SnapsToDevicePixels="true">
                                            <StackPanel IsItemsHost="true" KeyboardNavigation.DirectionalNavigation="Contained"/>
                                        </ScrollViewer>
                                    </Border>
                                </Grid>
                            </Popup>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="HasItems" Value="false">
                                <Setter TargetName="DropDownBorder" Property="MinHeight" Value="40"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="false">
                                <Setter Property="Opacity" Value="0.6"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="ComboBoxItem">
            <Setter Property="Background" Value="{DynamicResource BgCard}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="Padding" Value="10,7"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBoxItem">
                        <Border x:Name="ItemBorder" Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}" SnapsToDevicePixels="true">
                            <ContentPresenter Content="{TemplateBinding Content}" ContentTemplate="{TemplateBinding ContentTemplate}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsHighlighted" Value="true">
                                <Setter TargetName="ItemBorder" Property="Background" Value="{DynamicResource AccentBlue}"/>
                                <Setter Property="Foreground" Value="#FFFFFF"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="true">
                                <Setter TargetName="ItemBorder" Property="Background" Value="{DynamicResource AccentBlue}"/>
                                <Setter Property="Foreground" Value="#FFFFFF"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="false">
                                <Setter Property="Foreground" Value="{DynamicResource TextSecondary}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Modern Dynamic ListBox & Items -->
        <Style TargetType="ListBox">
            <Setter Property="Background" Value="{DynamicResource BgDark}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderCard}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="ScrollViewer.HorizontalScrollBarVisibility" Value="Disabled"/>
        </Style>

        <Style TargetType="ListBoxItem">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="Padding" Value="8,5"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ListBoxItem">
                        <Border x:Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" Padding="{TemplateBinding Padding}" CornerRadius="3" SnapsToDevicePixels="true">
                            <ContentPresenter Content="{TemplateBinding Content}" ContentTemplate="{TemplateBinding ContentTemplate}" HorizontalAlignment="Left" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="true">
                                <Setter TargetName="Bd" Property="Background" Value="{DynamicResource AccentBlue}"/>
                                <Setter Property="Foreground" Value="#FFFFFF"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="true">
                                <Setter TargetName="Bd" Property="Background" Value="{DynamicResource BgCardHover}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Modern Dynamic DataGrid & Elements -->
        <Style TargetType="DataGrid">
            <Setter Property="Background" Value="{DynamicResource BgDark}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderCard}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="RowBackground" Value="{DynamicResource DataGridRowBg}"/>
            <Setter Property="AlternatingRowBackground" Value="{DynamicResource DataGridAltRowBg}"/>
            <Setter Property="GridLinesVisibility" Value="Horizontal"/>
            <Setter Property="HorizontalGridLinesBrush" Value="{DynamicResource GridLines}"/>
            <Setter Property="HeadersVisibility" Value="Column"/>
            <Setter Property="AutoGenerateColumns" Value="False"/>
            <Setter Property="IsReadOnly" Value="True"/>
            <Setter Property="CanUserAddRows" Value="False"/>
            <Setter Property="CanUserDeleteRows" Value="False"/>
            <Setter Property="SelectionMode" Value="Single"/>
            <Setter Property="SelectionUnit" Value="FullRow"/>
        </Style>

        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="{DynamicResource DataGridHeaderBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource DataGridHeaderFg}"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="10,7"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderCard}"/>
            <Setter Property="BorderThickness" Value="0,0,1,1"/>
        </Style>

        <Style TargetType="DataGridRow">
            <Setter Property="Background" Value="{DynamicResource DataGridRowBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="SnapsToDevicePixels" Value="true"/>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="true">
                    <Setter Property="Background" Value="{DynamicResource AccentBlue}"/>
                    <Setter Property="Foreground" Value="#FFFFFF"/>
                </Trigger>
                <Trigger Property="IsMouseOver" Value="true">
                    <Setter Property="Background" Value="{DynamicResource BgCardHover}"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style TargetType="DataGridCell">
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="DataGridCell">
                        <Border Background="{TemplateBinding Background}" BorderThickness="0" Padding="{TemplateBinding Padding}" SnapsToDevicePixels="true">
                            <ContentPresenter SnapsToDevicePixels="{TemplateBinding SnapsToDevicePixels}" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="true">
                    <Setter Property="Background" Value="{DynamicResource AccentBlue}"/>
                    <Setter Property="Foreground" Value="#FFFFFF"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style TargetType="TabItem">
            <Setter Property="Background" Value="{DynamicResource BgCard}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextSecondary}"/>
            <Setter Property="Padding" Value="18,10"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Border Name="TabBorder"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{DynamicResource BorderCard}"
                                BorderThickness="1,1,1,0"
                                CornerRadius="6,6,0,0"
                                Margin="0,0,4,0"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter ContentSource="Header" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="TabBorder" Property="Background" Value="{DynamicResource AccentBlue}"/>
                                <Setter Property="Foreground" Value="#FFFFFF"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <Border Grid.Row="0" Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="14,10" Margin="0,0,0,12">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Orientation="Vertical">
                    <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                        <TextBlock Text="🗺️" FontSize="20" Margin="0,0,8,0" VerticalAlignment="Center"/>
                        <TextBlock Name="txtHeaderTitle" Text="Google Maps Route &amp; Map Generator" FontSize="18" FontWeight="Bold" Foreground="{DynamicResource TextPrimary}"/>
                    </StackPanel>
                    <TextBlock Name="txtHeaderSubtitle" Text="Multi-point driving routes: Fastest, Shortest, Eco-friendly | Import JSON, CSV, Excel" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="28,2,0,0"/>
                </StackPanel>
                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Name="lblApiBadge" Text="API: Checking..." Foreground="#EF4444" FontSize="12" FontWeight="SemiBold" VerticalAlignment="Center" Margin="0,0,10,0"/>
                    <ComboBox Name="cmbAppLanguage" Width="135" Height="30" Margin="0,0,10,0" VerticalAlignment="Center" ToolTip="Select Language / Sprache wählen / Wybierz język"/>
                    <Button Name="btnQuickSettings" Content="⚙ API Settings" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Padding="10,5" FontSize="12" Margin="0,0,10,0"/>
                    <Button Name="btnThemeToggle" Content="🌙 Dark" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" Padding="10,5" FontSize="12" ToolTip="Toggle Light / Dark theme"/>
                </StackPanel>
            </Grid>
        </Border>


        <!-- Main TabControl -->
        <TabControl Name="tabMain" Grid.Row="1" Background="Transparent" BorderThickness="0">

            <!-- TAB 1: MANUAL ROUTE -->
            <TabItem Name="tabItemManual" Header="📍 Manual Route">
                <Grid Margin="0,10,0,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="420" MinWidth="360"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <ScrollViewer VerticalScrollBarVisibility="Auto" Grid.Column="0" Margin="0,0,10,0">
                        <StackPanel>
                            <Border Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="14" Margin="0,0,0,12">
                                <StackPanel>
                                    <TextBlock Name="lblManualRoutePointsHeader" Text="Route Points" FontSize="15" FontWeight="Bold" Foreground="{DynamicResource TextPrimary}" Margin="0,0,0,10"/>

                                    <TextBlock Name="lblManualOrigin" Text="Origin (Start / A):" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,4"/>
                                    <Grid Margin="0,0,0,10">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="Auto"/>
                                        </Grid.ColumnDefinitions>
                                        <TextBox Name="txtManualStart" Text="Warszawa, Plac Defilad 1"/>
                                        <Button Name="btnClearManualStart" Grid.Column="1" Content="✕" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Padding="8,6" Margin="4,0,0,0" ToolTip="Clear"/>
                                    </Grid>

                                    <TextBlock Name="lblManualWaypoints" Text="Intermediate Stops (optional up to 25):" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,4"/>
                                    <Grid Margin="0,0,0,6">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="Auto"/>
                                        </Grid.ColumnDefinitions>
                                        <TextBox Name="txtNewWaypoint" ToolTip="Enter waypoint address and click Add"/>
                                        <Button Name="btnAddWaypoint" Grid.Column="1" Content="➕ Add" Background="#10B981" Margin="4,0,0,0"/>
                                    </Grid>

                                    <ListBox Name="lstWaypoints" Height="110" Margin="0,0,0,6"/>
                                    <Grid Margin="0,0,0,10">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="*"/>
                                        </Grid.ColumnDefinitions>
                                        <Button Name="btnWpUp" Content="▲ Up" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="0,0,2,0" Padding="4,4" FontSize="11"/>
                                        <Button Name="btnWpDown" Grid.Column="1" Content="▼ Down" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="2,0,2,0" Padding="4,4" FontSize="11"/>
                                        <Button Name="btnWpRemove" Grid.Column="2" Content="✕ Remove" Background="#EF4444" Margin="2,0,2,0" Padding="4,4" FontSize="11"/>
                                        <Button Name="btnWpClear" Grid.Column="3" Content="🗑 Clear" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="2,0,0,0" Padding="4,4" FontSize="11"/>
                                    </Grid>

                                    <TextBlock Name="lblManualDestination" Text="Destination (End / B):" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,4"/>
                                    <Grid Margin="0,0,0,6">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="Auto"/>
                                        </Grid.ColumnDefinitions>
                                        <TextBox Name="txtManualEnd" Text="Kraków, Rynek Główny 1"/>
                                        <Button Name="btnClearManualEnd" Grid.Column="1" Content="✕" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Padding="8,6" Margin="4,0,0,0" ToolTip="Clear"/>
                                    </Grid>

                                    <TextBlock Name="lblManualRouteName" Text="Route Name / Description:" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,4,0,4"/>
                                    <TextBox Name="txtManualName" Text="Route Warsaw - Krakow" Margin="0,0,0,6"/>
                                </StackPanel>
                            </Border>

                            <Border Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="14" Margin="0,0,0,12">
                                <StackPanel>
                                    <TextBlock Name="lblManualOptHeader" Text="Route Optimization" FontSize="15" FontWeight="Bold" Foreground="{DynamicResource TextPrimary}" Margin="0,0,0,10"/>

                                    <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                                        <RadioButton Name="rbTypeFastest" Content="⚡ Fastest" IsChecked="True" Foreground="{DynamicResource TextPrimary}" FontSize="13" Margin="0,0,16,0"/>
                                        <RadioButton Name="rbTypeShortest" Content="📏 Shortest" Foreground="{DynamicResource TextPrimary}" FontSize="13" Margin="0,0,16,0"/>
                                        <RadioButton Name="rbTypeEco" Content="🌿 Eco" Foreground="{DynamicResource TextPrimary}" FontSize="13"/>
                                    </StackPanel>

                                    <StackPanel Name="pnlEmission" Orientation="Vertical" Visibility="Collapsed" Margin="0,0,0,8">
                                        <TextBlock Name="lblManualEmission" Text="Vehicle Engine Type (for Eco route):" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,4"/>
                                        <ComboBox Name="cmbEmission">
                                            <ComboBoxItem Content="Gasoline (Benzyna)" Tag="GASOLINE" IsSelected="True"/>
                                            <ComboBoxItem Content="Diesel" Tag="DIESEL"/>
                                            <ComboBoxItem Content="Hybrid" Tag="HYBRID"/>
                                            <ComboBoxItem Content="Electric" Tag="ELECTRIC"/>
                                        </ComboBox>
                                    </StackPanel>

                                    <CheckBox Name="chkTrafficAware" Content="Real-time traffic awareness (Live Traffic)" Foreground="{DynamicResource TextSecondary}" FontSize="12" Margin="0,4,0,4"/>
                                </StackPanel>
                            </Border>

                            <Button Name="btnCalculateManual" Content="🚀 CALCULATE ROUTE &amp; DOWNLOAD MAP" Background="#2563EB" Foreground="#FFFFFF" Padding="16,12" FontSize="14" FontWeight="Bold"/>
                        </StackPanel>
                    </ScrollViewer>

                    <Grid Grid.Column="1" Margin="10,0,0,0">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>

                        <Border Grid.Row="0" Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="14" Margin="0,0,0,10">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>

                                <StackPanel Grid.Column="0">
                                    <TextBlock Name="lblHeaderDist" Text="DISTANCE" FontSize="11" FontWeight="Bold" Foreground="{DynamicResource TextSecondary}"/>
                                    <TextBlock Name="lblManualDist" Text="— km" FontSize="20" FontWeight="Bold" Foreground="#10B981"/>
                                </StackPanel>

                                <StackPanel Grid.Column="1">
                                    <TextBlock Name="lblHeaderDur" Text="DURATION" FontSize="11" FontWeight="Bold" Foreground="{DynamicResource TextSecondary}"/>
                                    <TextBlock Name="lblManualTime" Text="— min" FontSize="20" FontWeight="Bold" Foreground="#F59E0B"/>
                                </StackPanel>

                                <StackPanel Grid.Column="2">
                                    <TextBlock Name="lblHeaderType" Text="ROUTE TYPE" FontSize="11" FontWeight="Bold" Foreground="{DynamicResource TextSecondary}"/>
                                    <TextBlock Name="lblManualType" Text="Fastest" FontSize="16" FontWeight="SemiBold" Foreground="#38BDF8"/>
                                </StackPanel>

                                <StackPanel Grid.Column="3" VerticalAlignment="Center">
                                    <TextBlock Name="lblManualStatus" Text="Idle" FontSize="12" Foreground="{DynamicResource TextSecondary}" HorizontalAlignment="Right"/>
                                </StackPanel>
                            </Grid>
                        </Border>

                        <Border Grid.Row="1" Background="{DynamicResource BgDark}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="6" Margin="0,0,0,10">
                            <Grid>
                                <TextBlock Name="lblMapPlaceholder" Text="Map preview will appear here after route calculation..."
                                           Foreground="{DynamicResource TextSecondary}" FontSize="14" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                <Image Name="imgMapPreview" Stretch="Uniform" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                            </Grid>
                        </Border>

                        <Border Grid.Row="2" Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="10">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Name="lblGoogleUrlDisplay" Text="No generated link" Foreground="{DynamicResource TextSecondary}" FontSize="12" VerticalAlignment="Center" TextTrimming="CharacterEllipsis" Margin="0,0,10,0"/>
                                <Button Name="btnOpenGoogleMaps" Grid.Column="1" Content="🌐 Google Maps" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="0,0,6,0" Padding="10,6" IsEnabled="False"/>
                                <Button Name="btnCopyUrl" Grid.Column="2" Content="📋 Copy Link" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="0,0,6,0" Padding="10,6" IsEnabled="False"/>
                                <Button Name="btnSaveMapAs" Grid.Column="3" Content="💾 Save Map As..." Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Padding="10,6" IsEnabled="False"/>
                            </Grid>
                        </Border>
                    </Grid>
                </Grid>
            </TabItem>

            <!-- TAB 2: BATCH DATA PROCESSING -->
            <TabItem Name="tabItemBatch" Header="📁 Batch File Processing">
                <Grid Margin="0,10,0,0">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <Border Grid.Row="0" Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="14" Margin="0,0,0,10">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>

                            <Grid Grid.Row="0" Margin="0,0,0,10">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Name="lblBatchInputFile" Text="Input File (JSON/CSV/XLSX):" VerticalAlignment="Center" Foreground="{DynamicResource TextSecondary}" Margin="0,0,10,0"/>
                                <TextBox Name="txtBatchFilePath" Grid.Column="1" VerticalAlignment="Center"/>
                                <Button Name="btnBrowseBatchFile" Grid.Column="2" Content="📂 Browse File..." Background="#2563EB" Margin="6,0,0,0"/>
                                <Button Name="btnReloadBatchFile" Grid.Column="3" Content="🔄 Reload" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="6,0,0,0"/>
                            </Grid>

                            <Grid Grid.Row="1">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                                    <TextBlock Name="lblBatchFileInfo" Text="No file loaded." Foreground="{DynamicResource TextSecondary}" FontSize="12"/>
                                </StackPanel>
                                <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center" Margin="10,0">
                                    <TextBlock Name="lblBatchDefaultRouteType" Text="Default route type:" Foreground="{DynamicResource TextSecondary}" FontSize="12" VerticalAlignment="Center" Margin="0,0,6,0"/>
                                    <ComboBox Name="cmbBatchRouteType" Width="170">
                                        <ComboBoxItem Content="From Source / Default" Tag="FromSource" IsSelected="True"/>
                                        <ComboBoxItem Content="Fastest (Najszybsza)" Tag="Fastest"/>
                                        <ComboBoxItem Content="Shortest (Najkrótsza)" Tag="Shortest"/>
                                        <ComboBoxItem Content="Eco (Fuel Efficient)" Tag="Eco"/>
                                    </ComboBox>
                                </StackPanel>
                                <StackPanel Grid.Column="3" Orientation="Horizontal">
                                    <Button Name="btnStartBatch" Content="▶ Start Processing" Background="#10B981" Foreground="#FFFFFF" Padding="14,7" FontWeight="Bold" Margin="0,0,6,0"/>
                                    <Button Name="btnStopBatch" Content="⏹ Stop" Background="#EF4444" Foreground="#FFFFFF" Padding="12,7" IsEnabled="False"/>
                                </StackPanel>
                            </Grid>
                        </Grid>
                    </Border>

                    <TabControl Name="tabBatchSub" Grid.Row="1" Background="Transparent" BorderThickness="0">
                        <TabItem Name="tabSubInput" Header="📋 Input Data Preview">
                            <Border Background="{DynamicResource BgDark}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="6" Margin="0,6,0,0">
                                <DataGrid Name="dgBatchInput"/>
                            </Border>
                        </TabItem>

                        <TabItem Name="tabSubResults" Header="📊 Calculation Results">
                            <Border Background="{DynamicResource BgDark}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="6" Margin="0,6,0,0">
                                <DataGrid Name="dgBatchResults">
                                    <DataGrid.Columns>
                                        <DataGridTextColumn Header="ID" Binding="{Binding Id}" Width="45"/>
                                        <DataGridTextColumn Header="Route Name" Binding="{Binding Name}" Width="170"/>
                                        <DataGridTextColumn Header="Origin (Start)" Binding="{Binding Start}" Width="190"/>
                                        <DataGridTextColumn Header="Destination (End)" Binding="{Binding End}" Width="190"/>
                                        <DataGridTextColumn Header="Waypoints" Binding="{Binding WaypointsCount}" Width="75"/>
                                        <DataGridTextColumn Header="Type" Binding="{Binding RouteType}" Width="75"/>
                                        <DataGridTextColumn Header="Distance (km)" Binding="{Binding DistanceKm}" Width="95"/>
                                        <DataGridTextColumn Header="Duration (min)" Binding="{Binding DurationMin}" Width="85"/>
                                        <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="110"/>
                                        <DataGridTextColumn Header="PNG Map" Binding="{Binding MapPath}" Width="*"/>
                                    </DataGrid.Columns>
                                </DataGrid>
                            </Border>
                        </TabItem>

                        <TabItem Name="tabSubPoints" Header="📍 Points Detail">
                            <Border Background="{DynamicResource BgDark}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="6" Margin="0,6,0,0">
                                <DataGrid Name="dgBatchPoints">
                                    <DataGrid.Columns>
                                        <DataGridTextColumn Header="Route ID" Binding="{Binding RouteId}" Width="65"/>
                                        <DataGridTextColumn Header="Route Name" Binding="{Binding RouteName}" Width="150"/>
                                        <DataGridTextColumn Header="No." Binding="{Binding PointOrder}" Width="45"/>
                                        <DataGridTextColumn Header="Point Type" Binding="{Binding PointType}" Width="90"/>
                                        <DataGridTextColumn Header="Original Address" Binding="{Binding OriginalAddress}" Width="220"/>
                                        <DataGridTextColumn Header="Geocoded Address" Binding="{Binding GeocodedAddress}" Width="240"/>
                                        <DataGridTextColumn Header="Geocode Status" Binding="{Binding GeocodeStatus}" Width="170"/>
                                        <DataGridTextColumn Header="Match Type" Binding="{Binding MatchType}" Width="110"/>
                                        <DataGridTextColumn Header="Fallback?" Binding="{Binding IsFallback}" Width="75"/>
                                        <DataGridTextColumn Header="Latitude" Binding="{Binding Latitude}" Width="85"/>
                                        <DataGridTextColumn Header="Longitude" Binding="{Binding Longitude}" Width="85"/>
                                    </DataGrid.Columns>
                                </DataGrid>
                            </Border>
                        </TabItem>

                        <TabItem Name="tabSubLog" Header="📝 Activity Log">
                            <Border Background="{DynamicResource BgDark}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="6" Margin="0,6,0,0">
                                <TextBox Name="txtBatchLog" IsReadOnly="True" TextWrapping="Wrap"
                                         VerticalScrollBarVisibility="Auto" FontFamily="Consolas, monospace"
                                         FontSize="12" Background="{DynamicResource LogBg}" Foreground="{DynamicResource LogFg}"/>
                            </Border>
                        </TabItem>
                    </TabControl>

                    <Border Grid.Row="2" Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="12" Margin="0,10,0,0">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="230"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>

                            <StackPanel Grid.Column="0" VerticalAlignment="Center">
                                <TextBlock Name="lblBatchProgressText" Text="Ready" FontSize="12" Foreground="{DynamicResource TextSecondary}"/>
                                <ProgressBar Name="pbBatchProgress" Height="14" Minimum="0" Maximum="100" Value="0" Margin="0,4,0,0" Foreground="#10B981" Background="{DynamicResource BgDark}"/>
                            </StackPanel>

                            <TextBlock Name="lblBatchStats" Grid.Column="1" Text="" Foreground="#10B981" FontSize="13" FontWeight="SemiBold" VerticalAlignment="Center" Margin="20,0"/>

                            <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
                                <Button Name="btnOpenOutputDir" Content="📂 Open Output Folder" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="0,0,6,0" Padding="10,6"/>
                                <Button Name="btnExportExcel" Content="📊 Export Excel" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="0,0,6,0" Padding="10,6"/>
                                <Button Name="btnExportCsv" Content="📄 CSV" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="0,0,6,0" Padding="10,6"/>
                                <Button Name="btnExportJson" Content="📋 JSON" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Padding="10,6"/>
                            </StackPanel>
                        </Grid>
                    </Border>
                </Grid>
            </TabItem>

            <!-- TAB 3: SETTINGS & API KEY -->
            <TabItem Name="tabItemSettings" Header="⚙ Settings &amp; API Key">
                <ScrollViewer VerticalScrollBarVisibility="Auto" Margin="0,10,0,0">
                    <StackPanel MaxWidth="780" HorizontalAlignment="Left">
                        <Border Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="16" Margin="0,0,0,14">
                            <StackPanel>
                                <TextBlock Name="lblSettingsApiHeader" Text="Google Maps API Key" FontSize="16" FontWeight="Bold" Foreground="{DynamicResource TextPrimary}" Margin="0,0,0,8"/>
                                <TextBlock Name="lblSettingsApiDesc" Text="Required for Geocoding API, Routes API v2, and Static Maps API." FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,10"/>

                                <TextBlock Name="lblSettingsApiLabel" Text="API Key:" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,4"/>
                                <Grid Margin="0,0,0,8">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <PasswordBox Name="txtSettingsApiKey"/>
                                    <TextBox Name="txtSettingsApiKeyVisible" Visibility="Collapsed"/>
                                    <Button Name="btnToggleKeyVisibility" Grid.Column="1" Content="👁 Show" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="6,0,0,0" Padding="10,6"/>
                                    <Button Name="btnTestApiKey" Grid.Column="2" Content="🔍 Test Key" Background="#2563EB" Margin="6,0,0,0" Padding="12,6"/>
                                </Grid>

                                <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                                    <CheckBox Name="chkRememberKey" Content="Remember securely on this computer (DPAPI CurrentUser encryption)" IsChecked="True" Foreground="{DynamicResource TextSecondary}" FontSize="12"/>
                                </StackPanel>

                                <TextBlock Name="lblKeyTestResult" Text="" FontSize="12" FontWeight="SemiBold"/>
                            </StackPanel>
                        </Border>

                        <Border Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="16" Margin="0,0,0,14">
                            <StackPanel>
                                <TextBlock Name="lblSettingsPrefHeader" Text="Default Generation Preferences" FontSize="16" FontWeight="Bold" Foreground="{DynamicResource TextPrimary}" Margin="0,0,0,12"/>

                                <TextBlock Name="lblSettingsDefaultRouteType" Text="Default route type:" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,4"/>
                                <ComboBox Name="cmbDefaultRouteType" Margin="0,0,0,12">
                                    <ComboBoxItem Content="Fastest (Najszybsza)" Tag="Fastest" IsSelected="True"/>
                                    <ComboBoxItem Content="Shortest (Najkrótsza)" Tag="Shortest"/>
                                    <ComboBoxItem Content="Eco (Fuel Efficient)" Tag="Eco"/>
                                </ComboBox>

                                <TextBlock Name="lblSettingsDefaultEmission" Text="Default engine type for Eco routes:" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,4"/>
                                <ComboBox Name="cmbDefaultEmission" Margin="0,0,0,12">
                                    <ComboBoxItem Content="Gasoline (Benzyna)" Tag="GASOLINE" IsSelected="True"/>
                                    <ComboBoxItem Content="Diesel" Tag="DIESEL"/>
                                    <ComboBoxItem Content="Hybrid" Tag="HYBRID"/>
                                    <ComboBoxItem Content="Electric" Tag="ELECTRIC"/>
                                </ComboBox>

                                <TextBlock Name="lblSettingsDefaultMapSize" Text="Default dimensions for generated PNG map:" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,4"/>
                                <ComboBox Name="cmbDefaultMapSize" Margin="0,0,0,12">
                                    <ComboBoxItem Content="900 x 600 px (Recommended Standard)" Tag="900x600" IsSelected="True"/>
                                    <ComboBoxItem Content="1024 x 768 px (High Res)" Tag="1024x768"/>
                                    <ComboBoxItem Content="1280 x 720 px (HD 16:9)" Tag="1280x720"/>
                                    <ComboBoxItem Content="640 x 640 px (Square)" Tag="640x640"/>
                                    <ComboBoxItem Content="1600 x 900 px (Full HD 16:9)" Tag="1600x900"/>
                                </ComboBox>

                                <TextBlock Name="lblSettingsOutputDir" Text="Results Output Folder:" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,4"/>
                                <Grid Margin="0,0,0,12">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBox Name="txtSettingsOutputDir"/>
                                    <Button Name="btnBrowseSettingsOutputDir" Grid.Column="1" Content="📂 Browse..." Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="6,0,0,0"/>
                                </Grid>
                            </StackPanel>
                        </Border>

                        <!-- MAP OVERLAY CARD -->
                        <Border Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="16" Margin="0,0,0,14">
                            <StackPanel>
                                <TextBlock Name="lblSettingsOverlayHeader" Text="Map Overlay &amp; Banners (Top / Bottom)" FontSize="16" FontWeight="Bold" Foreground="{DynamicResource TextPrimary}" Margin="0,0,0,6"/>
                                <TextBlock Name="lblSettingsOverlayDesc" Text="Configure whether to display top and bottom banner panels, and choose which properties appear on each panel, line order, and alignment." FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,12" TextWrapping="Wrap"/>

                                <StackPanel Orientation="Horizontal" Margin="0,0,0,12">
                                    <CheckBox Name="chkEnableTopOverlay" Content="Enable Top Banner" IsChecked="True" Foreground="{DynamicResource TextPrimary}" FontSize="12" FontWeight="SemiBold" Margin="0,0,24,0"/>
                                    <CheckBox Name="chkEnableBottomOverlay" Content="Enable Bottom Banner" IsChecked="True" Foreground="{DynamicResource TextPrimary}" FontSize="12" FontWeight="SemiBold"/>
                                </StackPanel>

                                <Border Background="{DynamicResource BgDark}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="6" Padding="10" Margin="0,0,0,10">
                                    <Grid Name="gridOverlayConfig">
                                        <Grid.RowDefinitions>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                        </Grid.RowDefinitions>
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="220"/>
                                            <ColumnDefinition Width="65"/>
                                            <ColumnDefinition Width="135"/>
                                            <ColumnDefinition Width="135"/>
                                            <ColumnDefinition Width="110"/>
                                        </Grid.ColumnDefinitions>

                                        <!-- Header Row -->
                                        <TextBlock Name="lblColPropName" Grid.Row="0" Grid.Column="0" Text="Property" FontWeight="Bold" Foreground="{DynamicResource TextSecondary}" FontSize="12" Margin="4,2,4,8"/>
                                        <TextBlock Name="lblColPropShow" Grid.Row="0" Grid.Column="1" Text="Show" FontWeight="Bold" Foreground="{DynamicResource TextSecondary}" FontSize="12" Margin="4,2,4,8" HorizontalAlignment="Center"/>
                                        <TextBlock Name="lblColPropPanel" Grid.Row="0" Grid.Column="2" Text="Panel" FontWeight="Bold" Foreground="{DynamicResource TextSecondary}" FontSize="12" Margin="4,2,4,8"/>
                                        <TextBlock Name="lblColPropAlign" Grid.Row="0" Grid.Column="3" Text="Alignment" FontWeight="Bold" Foreground="{DynamicResource TextSecondary}" FontSize="12" Margin="4,2,4,8"/>
                                        <TextBlock Name="lblColPropOrder" Grid.Row="0" Grid.Column="4" Text="Line / Order" FontWeight="Bold" Foreground="{DynamicResource TextSecondary}" FontSize="12" Margin="4,2,4,8" HorizontalAlignment="Center"/>

                                        <!-- Row 1: StartGeocoded -->
                                        <TextBlock Name="lblProp_StartGeocoded" Grid.Row="1" Grid.Column="0" Text="Start Address (Geocoded)" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_StartGeocoded" Grid.Row="1" Grid.Column="1" IsChecked="True" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_StartGeocoded" Grid.Row="1" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom" IsSelected="True"/>
                                            <ComboBoxItem Content="Top" Tag="Top"/>
                                            <ComboBoxItem Content="None" Tag="None"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_StartGeocoded" Grid.Row="1" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left" IsSelected="True"/>
                                            <ComboBoxItem Content="Center" Tag="Center"/>
                                            <ComboBoxItem Content="Right" Tag="Right"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_StartGeocoded" Grid.Row="1" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1" IsSelected="True"/>
                                            <ComboBoxItem Content="2" Tag="2"/>
                                            <ComboBoxItem Content="3" Tag="3"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>

                                        <!-- Row 2: EndGeocoded -->
                                        <TextBlock Name="lblProp_EndGeocoded" Grid.Row="2" Grid.Column="0" Text="End Address (Geocoded)" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_EndGeocoded" Grid.Row="2" Grid.Column="1" IsChecked="True" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_EndGeocoded" Grid.Row="2" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom" IsSelected="True"/>
                                            <ComboBoxItem Content="Top" Tag="Top"/>
                                            <ComboBoxItem Content="None" Tag="None"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_EndGeocoded" Grid.Row="2" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left" IsSelected="True"/>
                                            <ComboBoxItem Content="Center" Tag="Center"/>
                                            <ComboBoxItem Content="Right" Tag="Right"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_EndGeocoded" Grid.Row="2" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1"/>
                                            <ComboBoxItem Content="2" Tag="2" IsSelected="True"/>
                                            <ComboBoxItem Content="3" Tag="3"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>

                                        <!-- Row 3: Distance -->
                                        <TextBlock Name="lblProp_Distance" Grid.Row="3" Grid.Column="0" Text="Total Distance" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_Distance" Grid.Row="3" Grid.Column="1" IsChecked="True" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_Distance" Grid.Row="3" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom" IsSelected="True"/>
                                            <ComboBoxItem Content="Top" Tag="Top"/>
                                            <ComboBoxItem Content="None" Tag="None"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_Distance" Grid.Row="3" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left" IsSelected="True"/>
                                            <ComboBoxItem Content="Center" Tag="Center"/>
                                            <ComboBoxItem Content="Right" Tag="Right"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_Distance" Grid.Row="3" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1"/>
                                            <ComboBoxItem Content="2" Tag="2"/>
                                            <ComboBoxItem Content="3" Tag="3" IsSelected="True"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>

                                        <!-- Row 4: Duration -->
                                        <TextBlock Name="lblProp_Duration" Grid.Row="4" Grid.Column="0" Text="Total Time" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_Duration" Grid.Row="4" Grid.Column="1" IsChecked="True" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_Duration" Grid.Row="4" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom" IsSelected="True"/>
                                            <ComboBoxItem Content="Top" Tag="Top"/>
                                            <ComboBoxItem Content="None" Tag="None"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_Duration" Grid.Row="4" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left"/>
                                            <ComboBoxItem Content="Center" Tag="Center" IsSelected="True"/>
                                            <ComboBoxItem Content="Right" Tag="Right"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_Duration" Grid.Row="4" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1"/>
                                            <ComboBoxItem Content="2" Tag="2"/>
                                            <ComboBoxItem Content="3" Tag="3" IsSelected="True"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>

                                        <!-- Row 5: Timestamp -->
                                        <TextBlock Name="lblProp_Timestamp" Grid.Row="5" Grid.Column="0" Text="Generation Timestamp" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_Timestamp" Grid.Row="5" Grid.Column="1" IsChecked="True" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_Timestamp" Grid.Row="5" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom" IsSelected="True"/>
                                            <ComboBoxItem Content="Top" Tag="Top"/>
                                            <ComboBoxItem Content="None" Tag="None"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_Timestamp" Grid.Row="5" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left"/>
                                            <ComboBoxItem Content="Center" Tag="Center"/>
                                            <ComboBoxItem Content="Right" Tag="Right" IsSelected="True"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_Timestamp" Grid.Row="5" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1"/>
                                            <ComboBoxItem Content="2" Tag="2"/>
                                            <ComboBoxItem Content="3" Tag="3" IsSelected="True"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>

                                        <!-- Row 6: RouteName -->
                                        <TextBlock Name="lblProp_RouteName" Grid.Row="6" Grid.Column="0" Text="Route Name" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_RouteName" Grid.Row="6" Grid.Column="1" IsChecked="True" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_RouteName" Grid.Row="6" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom"/>
                                            <ComboBoxItem Content="Top" Tag="Top" IsSelected="True"/>
                                            <ComboBoxItem Content="None" Tag="None"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_RouteName" Grid.Row="6" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left" IsSelected="True"/>
                                            <ComboBoxItem Content="Center" Tag="Center"/>
                                            <ComboBoxItem Content="Right" Tag="Right"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_RouteName" Grid.Row="6" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1" IsSelected="True"/>
                                            <ComboBoxItem Content="2" Tag="2"/>
                                            <ComboBoxItem Content="3" Tag="3"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>

                                        <!-- Row 7: RouteType -->
                                        <TextBlock Name="lblProp_RouteType" Grid.Row="7" Grid.Column="0" Text="Route Type" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_RouteType" Grid.Row="7" Grid.Column="1" IsChecked="True" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_RouteType" Grid.Row="7" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom"/>
                                            <ComboBoxItem Content="Top" Tag="Top" IsSelected="True"/>
                                            <ComboBoxItem Content="None" Tag="None"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_RouteType" Grid.Row="7" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left"/>
                                            <ComboBoxItem Content="Center" Tag="Center"/>
                                            <ComboBoxItem Content="Right" Tag="Right" IsSelected="True"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_RouteType" Grid.Row="7" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1" IsSelected="True"/>
                                            <ComboBoxItem Content="2" Tag="2"/>
                                            <ComboBoxItem Content="3" Tag="3"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>

                                        <!-- Row 8: Waypoints -->
                                        <TextBlock Name="lblProp_Waypoints" Grid.Row="8" Grid.Column="0" Text="Intermediate Stops (Waypoints)" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_Waypoints" Grid.Row="8" Grid.Column="1" IsChecked="False" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_Waypoints" Grid.Row="8" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom" IsSelected="True"/>
                                            <ComboBoxItem Content="Top" Tag="Top"/>
                                            <ComboBoxItem Content="None" Tag="None"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_Waypoints" Grid.Row="8" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left" IsSelected="True"/>
                                            <ComboBoxItem Content="Center" Tag="Center"/>
                                            <ComboBoxItem Content="Right" Tag="Right"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_Waypoints" Grid.Row="8" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1"/>
                                            <ComboBoxItem Content="2" Tag="2" IsSelected="True"/>
                                            <ComboBoxItem Content="3" Tag="3"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>

                                        <!-- Row 9: StartRaw -->
                                        <TextBlock Name="lblProp_StartRaw" Grid.Row="9" Grid.Column="0" Text="Start Address (Raw Input)" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_StartRaw" Grid.Row="9" Grid.Column="1" IsChecked="False" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_StartRaw" Grid.Row="9" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom"/>
                                            <ComboBoxItem Content="Top" Tag="Top"/>
                                            <ComboBoxItem Content="None" Tag="None" IsSelected="True"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_StartRaw" Grid.Row="9" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left" IsSelected="True"/>
                                            <ComboBoxItem Content="Center" Tag="Center"/>
                                            <ComboBoxItem Content="Right" Tag="Right"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_StartRaw" Grid.Row="9" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1" IsSelected="True"/>
                                            <ComboBoxItem Content="2" Tag="2"/>
                                            <ComboBoxItem Content="3" Tag="3"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>

                                        <!-- Row 10: EndRaw -->
                                        <TextBlock Name="lblProp_EndRaw" Grid.Row="10" Grid.Column="0" Text="End Address (Raw Input)" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_EndRaw" Grid.Row="10" Grid.Column="1" IsChecked="False" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_EndRaw" Grid.Row="10" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom"/>
                                            <ComboBoxItem Content="Top" Tag="Top"/>
                                            <ComboBoxItem Content="None" Tag="None" IsSelected="True"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_EndRaw" Grid.Row="10" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left" IsSelected="True"/>
                                            <ComboBoxItem Content="Center" Tag="Center"/>
                                            <ComboBoxItem Content="Right" Tag="Right"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_EndRaw" Grid.Row="10" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1"/>
                                            <ComboBoxItem Content="2" Tag="2" IsSelected="True"/>
                                            <ComboBoxItem Content="3" Tag="3"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>
                                    </Grid>
                                </Border>

                                <StackPanel Orientation="Horizontal" Margin="0,2,0,0">
                                    <Button Name="btnResetOverlayConfig" Content="🔄 Reset to Default Layout" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Padding="12,6"/>
                                </StackPanel>
                            </StackPanel>
                        </Border>

                        <Border Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="16" Margin="0,0,0,14">
                            <StackPanel>
                                <TextBlock Name="lblSettingsLangHeader" Text="Language &amp; Localization" FontSize="16" FontWeight="Bold" Foreground="{DynamicResource TextPrimary}" Margin="0,0,0,8"/>
                                <TextBlock Name="lblSettingsLangLabel" Text="Application and Google Maps API Language:" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,6"/>
                                <ComboBox Name="cmbSettingsLanguage" Margin="0,0,0,10"/>
                                <StackPanel Orientation="Horizontal">
                                    <Button Name="btnOpenLangFile" Content="📂 Open Localization File (localization.json)" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Padding="10,6" Margin="0,0,8,0" ToolTip="Open the external localization file to edit or add new languages"/>
                                    <Button Name="btnReloadLang" Content="🔄 Reload Languages" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Padding="10,6" ToolTip="Reload language definitions from disk"/>
                                </StackPanel>
                            </StackPanel>
                        </Border>

                        <Border Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="16" Margin="0,0,0,14">
                            <StackPanel>
                                <TextBlock Name="lblSettingsThemeHeader" Text="Appearance &amp; Theme" FontSize="16" FontWeight="Bold" Foreground="{DynamicResource TextPrimary}" Margin="0,0,0,8"/>
                                <TextBlock Name="lblSettingsThemeLabel" Text="Application Theme (Color Scheme):" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,6"/>
                                <ComboBox Name="cmbSettingsTheme" Margin="0,0,0,0">
                                    <ComboBoxItem Content="🌙 Dark" Tag="Dark" IsSelected="True"/>
                                    <ComboBoxItem Content="☀️ Light" Tag="Light"/>
                                </ComboBox>
                            </StackPanel>
                        </Border>

                        <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                            <Button Name="btnSaveSettings" Content="💾 SAVE SETTINGS" Background="#10B981" Foreground="#FFFFFF" Padding="14,10" FontWeight="Bold" Width="200"/>
                            <Button Name="btnOpenLogFile" Content="📋 OPEN LOG FILE" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Padding="14,10" FontWeight="SemiBold" Margin="10,0,0,0"/>
                        </StackPanel>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
        </TabControl>

        <!-- Footer -->
        <Border Grid.Row="2" Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="6" Padding="10,6" Margin="0,10,0,0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Name="lblFooterStatus" Text="Ready." Foreground="{DynamicResource TextSecondary}" FontSize="12" VerticalAlignment="Center"/>
                <TextBlock Name="lblFooterVersion" Grid.Column="1" Text="Google Maps Routes v2.0" Foreground="{DynamicResource TextSecondary}" FontSize="12" VerticalAlignment="Center"/>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# ── 8. Tworzenie okna WPF z XAML ─────────────────────────────────────────────
$reader = [System.Xml.XmlNodeReader]::new($xaml)
$window = [System.Windows.Markup.XamlReader]::Load($reader)

# Zastosowanie DWM Dark Mode dla okna (zgodnie z motywem)
$window.Add_SourceInitialized({
    try {
        $helper = [System.Windows.Interop.WindowInteropHelper]::new($window)
        $val = if ($script:CurrentTheme -eq 'Light') { 0 } else { 1 }
        [DwmDarkWindow]::DwmSetWindowAttribute($helper.Handle, 20, [ref]$val, 4)
    } catch {}
})

# ── 9. Pobranie referencji do elementów UI ───────────────────────────────────
$txtHeaderTitle      = $window.FindName('txtHeaderTitle')
$txtHeaderSubtitle   = $window.FindName('txtHeaderSubtitle')
$cmbAppLanguage      = $window.FindName('cmbAppLanguage')
$btnThemeToggle      = $window.FindName('btnThemeToggle')
$lblApiBadge         = $window.FindName('lblApiBadge')
$btnQuickSettings    = $window.FindName('btnQuickSettings')
$tabMain             = $window.FindName('tabMain')
$tabItemManual       = $window.FindName('tabItemManual')
$tabItemBatch        = $window.FindName('tabItemBatch')
$tabItemSettings     = $window.FindName('tabItemSettings')
$lblFooterStatus     = $window.FindName('lblFooterStatus')
$lblFooterVersion    = $window.FindName('lblFooterVersion')

# Tab 1: Manual
$lblManualRoutePointsHeader = $window.FindName('lblManualRoutePointsHeader')
$lblManualOrigin            = $window.FindName('lblManualOrigin')
$lblManualWaypoints         = $window.FindName('lblManualWaypoints')
$lblManualDestination       = $window.FindName('lblManualDestination')
$lblManualRouteName         = $window.FindName('lblManualRouteName')
$lblManualOptHeader         = $window.FindName('lblManualOptHeader')
$lblManualEmission          = $window.FindName('lblManualEmission')
$lblHeaderDist              = $window.FindName('lblHeaderDist')
$lblHeaderDur               = $window.FindName('lblHeaderDur')
$lblHeaderType              = $window.FindName('lblHeaderType')

# Tab 1: Manual
$txtManualStart      = $window.FindName('txtManualStart')
$btnClearManualStart = $window.FindName('btnClearManualStart')
$txtNewWaypoint      = $window.FindName('txtNewWaypoint')
$btnAddWaypoint      = $window.FindName('btnAddWaypoint')
$lstWaypoints        = $window.FindName('lstWaypoints')
$btnWpUp             = $window.FindName('btnWpUp')
$btnWpDown           = $window.FindName('btnWpDown')
$btnWpRemove         = $window.FindName('btnWpRemove')
$btnWpClear          = $window.FindName('btnWpClear')
$txtManualEnd        = $window.FindName('txtManualEnd')
$btnClearManualEnd   = $window.FindName('btnClearManualEnd')
$txtManualName       = $window.FindName('txtManualName')
$rbTypeFastest       = $window.FindName('rbTypeFastest')
$rbTypeShortest      = $window.FindName('rbTypeShortest')
$rbTypeEco           = $window.FindName('rbTypeEco')
$pnlEmission         = $window.FindName('pnlEmission')
$cmbEmission         = $window.FindName('cmbEmission')
$chkTrafficAware     = $window.FindName('chkTrafficAware')
$btnCalculateManual  = $window.FindName('btnCalculateManual')
$lblManualDist       = $window.FindName('lblManualDist')
$lblManualTime       = $window.FindName('lblManualTime')
$lblManualType       = $window.FindName('lblManualType')
$lblManualStatus     = $window.FindName('lblManualStatus')
$lblMapPlaceholder   = $window.FindName('lblMapPlaceholder')
$imgMapPreview       = $window.FindName('imgMapPreview')
$lblGoogleUrlDisplay = $window.FindName('lblGoogleUrlDisplay')
$btnOpenGoogleMaps   = $window.FindName('btnOpenGoogleMaps')
$btnCopyUrl          = $window.FindName('btnCopyUrl')
$btnSaveMapAs        = $window.FindName('btnSaveMapAs')

# Tab 2: Batch
$lblBatchInputFile       = $window.FindName('lblBatchInputFile')
$lblBatchDefaultRouteType= $window.FindName('lblBatchDefaultRouteType')
$tabSubInput             = $window.FindName('tabSubInput')
$tabSubResults           = $window.FindName('tabSubResults')
$tabSubPoints            = $window.FindName('tabSubPoints')
$tabSubLog               = $window.FindName('tabSubLog')
$txtBatchFilePath    = $window.FindName('txtBatchFilePath')
$btnBrowseBatchFile  = $window.FindName('btnBrowseBatchFile')
$btnReloadBatchFile  = $window.FindName('btnReloadBatchFile')
$lblBatchFileInfo    = $window.FindName('lblBatchFileInfo')
$cmbBatchRouteType   = $window.FindName('cmbBatchRouteType')
$btnStartBatch       = $window.FindName('btnStartBatch')
$btnStopBatch        = $window.FindName('btnStopBatch')
$tabBatchSub        = $window.FindName('tabBatchSub')
$dgBatchInput        = $window.FindName('dgBatchInput')
$dgBatchResults      = $window.FindName('dgBatchResults')
$dgBatchPoints       = $window.FindName('dgBatchPoints')
$txtBatchLog         = $window.FindName('txtBatchLog')
$lblBatchProgressText= $window.FindName('lblBatchProgressText')
$pbBatchProgress     = $window.FindName('pbBatchProgress')
$lblBatchStats       = $window.FindName('lblBatchStats')
$btnOpenOutputDir    = $window.FindName('btnOpenOutputDir')
$btnExportExcel      = $window.FindName('btnExportExcel')
$btnExportCsv        = $window.FindName('btnExportCsv')
$btnExportJson       = $window.FindName('btnExportJson')

# Tab 3: Settings
$txtSettingsApiKey          = $window.FindName('txtSettingsApiKey')
$txtSettingsApiKeyVisible   = $window.FindName('txtSettingsApiKeyVisible')
$btnToggleKeyVisibility     = $window.FindName('btnToggleKeyVisibility')
$btnTestApiKey              = $window.FindName('btnTestApiKey')
$chkRememberKey             = $window.FindName('chkRememberKey')
$lblKeyTestResult           = $window.FindName('lblKeyTestResult')
$cmbDefaultRouteType        = $window.FindName('cmbDefaultRouteType')
$cmbDefaultEmission         = $window.FindName('cmbDefaultEmission')
$cmbDefaultMapSize          = $window.FindName('cmbDefaultMapSize')
$txtSettingsOutputDir       = $window.FindName('txtSettingsOutputDir')
$btnBrowseSettingsOutputDir = $window.FindName('btnBrowseSettingsOutputDir')
$btnSaveSettings            = $window.FindName('btnSaveSettings')
$btnOpenLogFile             = $window.FindName('btnOpenLogFile')
$lblSettingsApiHeader       = $window.FindName('lblSettingsApiHeader')
$lblSettingsApiDesc         = $window.FindName('lblSettingsApiDesc')
$lblSettingsApiLabel        = $window.FindName('lblSettingsApiLabel')
$lblSettingsPrefHeader      = $window.FindName('lblSettingsPrefHeader')
$lblSettingsDefaultRouteType= $window.FindName('lblSettingsDefaultRouteType')
$lblSettingsDefaultEmission = $window.FindName('lblSettingsDefaultEmission')
$lblSettingsDefaultMapSize  = $window.FindName('lblSettingsDefaultMapSize')
$lblSettingsOutputDir       = $window.FindName('lblSettingsOutputDir')
$lblSettingsLangHeader      = $window.FindName('lblSettingsLangHeader')
$lblSettingsLangLabel       = $window.FindName('lblSettingsLangLabel')
$cmbSettingsLanguage        = $window.FindName('cmbSettingsLanguage')
$btnOpenLangFile            = $window.FindName('btnOpenLangFile')
$btnReloadLang              = $window.FindName('btnReloadLang')
$lblSettingsThemeHeader     = $window.FindName('lblSettingsThemeHeader')
$lblSettingsThemeLabel      = $window.FindName('lblSettingsThemeLabel')
$cmbSettingsTheme           = $window.FindName('cmbSettingsTheme')

# Tab 3: Overlay Settings
$lblSettingsOverlayHeader    = $window.FindName('lblSettingsOverlayHeader')
$lblSettingsOverlayDesc      = $window.FindName('lblSettingsOverlayDesc')
$chkEnableTopOverlay         = $window.FindName('chkEnableTopOverlay')
$chkEnableBottomOverlay      = $window.FindName('chkEnableBottomOverlay')
$lblColPropName              = $window.FindName('lblColPropName')
$lblColPropShow              = $window.FindName('lblColPropShow')
$lblColPropPanel             = $window.FindName('lblColPropPanel')
$lblColPropAlign             = $window.FindName('lblColPropAlign')
$lblColPropOrder             = $window.FindName('lblColPropOrder')
$btnResetOverlayConfig       = $window.FindName('btnResetOverlayConfig')

foreach ($key in $script:OverlayPropKeys) {
    Set-Variable -Name "lblProp_$key"  -Value ($window.FindName("lblProp_$key"))  -Scope Script
    Set-Variable -Name "chkProp_$key"  -Value ($window.FindName("chkProp_$key"))  -Scope Script
    Set-Variable -Name "cmbPanel_$key" -Value ($window.FindName("cmbPanel_$key")) -Scope Script
    Set-Variable -Name "cmbAlign_$key" -Value ($window.FindName("cmbAlign_$key")) -Scope Script
    Set-Variable -Name "cmbOrder_$key" -Value ($window.FindName("cmbOrder_$key")) -Scope Script
}

# ── 10. System wielojęzyczności i funkcje pomocnicze stanu UI ─────────────────

function Populate-LanguageDropdowns {
    $script:SuppressLangEvents = $true
    try {
        if ($cmbAppLanguage) {
            $cmbAppLanguage.Items.Clear()
            foreach ($langCode in $script:LanguagesCatalog.Keys) {
                $langObj = $script:LanguagesCatalog[$langCode]
                $cbi = [System.Windows.Controls.ComboBoxItem]::new()
                $cbi.Content = "[$($langObj.Code.ToUpper())] $($langObj.DisplayName)"
                $cbi.Tag = $langObj.Code
                if ($langObj.Code -eq $script:CurrentLanguage) { $cbi.IsSelected = $true }
                $null = $cmbAppLanguage.Items.Add($cbi)
            }
        }
        if ($cmbSettingsLanguage) {
            $cmbSettingsLanguage.Items.Clear()
            foreach ($langCode in $script:LanguagesCatalog.Keys) {
                $langObj = $script:LanguagesCatalog[$langCode]
                $cbi = [System.Windows.Controls.ComboBoxItem]::new()
                $cbi.Content = "[$($langObj.Code.ToUpper())] $($langObj.DisplayName)"
                $cbi.Tag = $langObj.Code
                if ($langObj.Code -eq $script:CurrentLanguage) { $cbi.IsSelected = $true }
                $null = $cmbSettingsLanguage.Items.Add($cbi)
            }
        }
    } finally {
        $script:SuppressLangEvents = $false
    }
}

function Set-AppTheme {
    param(
        [Parameter(Mandatory = $false)]
        [string]$Theme = 'Dark'
    )

    if ($Theme -notmatch '(?i)light|dark') { $Theme = 'Dark' }
    $isLight = ($Theme -match '(?i)light')
    $script:CurrentTheme = if ($isLight) { 'Light' } else { 'Dark' }

    $palette = if ($isLight) {
        [ordered]@{
            'BgDark'                 = '#F1F5F9'
            'BgCard'                 = '#FFFFFF'
            'BgCardHover'            = '#F8FAFC'
            'BgCardAlt'              = '#F8FAFC'
            'BorderCard'             = '#CBD5E1'
            'TextPrimary'            = '#0F172A'
            'TextSecondary'          = '#475569'
            'AccentBlue'             = '#2563EB'
            'AccentGreen'            = '#059669'
            'AccentAmber'            = '#D97706'
            'AccentRed'              = '#DC2626'
            'BgInput'                = '#FFFFFF'
            'BorderInput'            = '#CBD5E1'
            'BtnSecondaryBg'         = '#E2E8F0'
            'BtnSecondaryFg'         = '#0F172A'
            'GridLines'              = '#E2E8F0'
            'LogBg'                  = '#F8FAFC'
            'LogFg'                  = '#0369A1'
            'DataGridHeaderBg'       = '#E2E8F0'
            'DataGridHeaderFg'       = '#334155'
            'DataGridRowBg'          = '#FFFFFF'
            'DataGridAltRowBg'       = '#F8FAFC'
        }
    } else {
        [ordered]@{
            'BgDark'                 = '#0F172A'
            'BgCard'                 = '#1E293B'
            'BgCardHover'            = '#293548'
            'BgCardAlt'              = '#162032'
            'BorderCard'             = '#334155'
            'TextPrimary'            = '#F8FAFC'
            'TextSecondary'          = '#94A3B8'
            'AccentBlue'             = '#2563EB'
            'AccentGreen'            = '#10B981'
            'AccentAmber'            = '#F59E0B'
            'AccentRed'              = '#EF4444'
            'BgInput'                = '#1E293B'
            'BorderInput'            = '#334155'
            'BtnSecondaryBg'         = '#334155'
            'BtnSecondaryFg'         = '#F8FAFC'
            'GridLines'              = '#2D3748'
            'LogBg'                  = '#0A0F1D'
            'LogFg'                  = '#38BDF8'
            'DataGridHeaderBg'       = '#0F172A'
            'DataGridHeaderFg'       = '#94A3B8'
            'DataGridRowBg'          = '#1E293B'
            'DataGridAltRowBg'       = '#162032'
        }
    }

    foreach ($k in $palette.Keys) {
        $c = [System.Windows.Media.ColorConverter]::ConvertFromString($palette[$k])
        $brush = [System.Windows.Media.SolidColorBrush]::new($c)
        $brush.Freeze()
        $window.Resources[$k] = $brush
        $window.Resources["Theme_$k"] = $brush
    }
    $window.Resources['Theme_BgApp'] = $window.Resources['BgDark']
    $window.Resources['Theme_Border'] = $window.Resources['BorderCard']

    if ($window) {
        $window.Background = $window.Resources['BgDark']
        $window.Foreground = $window.Resources['TextPrimary']
    }

    # Update overlay property labels foreground
    if ($script:OverlayPropKeys) {
        foreach ($k in $script:OverlayPropKeys) {
            $lblCtrl = Get-Variable -Name "lblProp_$k" -ValueOnly -ErrorAction SilentlyContinue
            if ($lblCtrl) {
                $lblCtrl.Foreground = $window.Resources['TextPrimary']
            }
        }
    }

    # Update DWM title bar chrome
    try {
        $helper = [System.Windows.Interop.WindowInteropHelper]::new($window)
        if ($helper.Handle -ne [IntPtr]::Zero) {
            $val = if ($isLight) { 0 } else { 1 }
            [DwmDarkWindow]::DwmSetWindowAttribute($helper.Handle, 20, [ref]$val, 4)
        }
    } catch {}

    # Update Toggle Button text/icon
    if ($btnThemeToggle) {
        $btnThemeToggle.Content = if ($isLight) { (Get-LocText 'ThemeLight') } else { (Get-LocText 'ThemeDark') }
    }

    # Update Settings ComboBox
    if ($cmbSettingsTheme) {
        $script:SuppressThemeEvents = $true
        try {
            foreach ($it in $cmbSettingsTheme.Items) {
                if ($it.Tag -eq $script:CurrentTheme) {
                    $cmbSettingsTheme.SelectedItem = $it
                    break
                }
            }
        } finally {
            $script:SuppressThemeEvents = $false
        }
    }
}

function Apply-AppLanguage {
    param([string]$LanguageCode)

    if (-not [string]::IsNullOrWhiteSpace($LanguageCode) -and $script:LanguagesCatalog.Contains($LanguageCode.ToLower())) {
        $script:CurrentLanguage = $LanguageCode.ToLower()
    } else {
        $script:CurrentLanguage = 'en'
    }

    $langObj = $script:LanguagesCatalog[$script:CurrentLanguage]
    $script:CurrentGoogleLang = if ($langObj -and $langObj.GoogleCode) { $langObj.GoogleCode } else { $script:CurrentLanguage }
    $script:CurrentStrings = if ($langObj -and $langObj.Strings) { $langObj.Strings } else { @{} }

    # Synchronize ComboBoxes without triggering duplicate events
    $script:SuppressLangEvents = $true
    try {
        if ($cmbAppLanguage) {
            foreach ($item in $cmbAppLanguage.Items) {
                if ($item.Tag -eq $script:CurrentLanguage) {
                    $cmbAppLanguage.SelectedItem = $item
                    break
                }
            }
        }
        if ($cmbSettingsLanguage) {
            foreach ($item in $cmbSettingsLanguage.Items) {
                if ($item.Tag -eq $script:CurrentLanguage) {
                    $cmbSettingsLanguage.SelectedItem = $item
                    break
                }
            }
        }
    } finally {
        $script:SuppressLangEvents = $false
    }

    # Window & Header
    if ($window) { $window.Title = (Get-LocText 'AppTitle') }
    if ($txtHeaderTitle) { $txtHeaderTitle.Text = (Get-LocText 'AppTitle') }
    if ($txtHeaderSubtitle) { $txtHeaderSubtitle.Text = (Get-LocText 'AppSubtitle') }
    if ($btnQuickSettings) { $btnQuickSettings.Content = (Get-LocText 'BtnQuickSettings') }
    if ($btnThemeToggle) {
        $btnThemeToggle.ToolTip = (Get-LocText 'ThemeToggleTip')
        $btnThemeToggle.Content = if ($script:CurrentTheme -eq 'Light') { (Get-LocText 'ThemeLight') } else { (Get-LocText 'ThemeDark') }
    }

    # Tab Headers
    if ($tabItemManual) { $tabItemManual.Header = (Get-LocText 'TabManual') }
    if ($tabItemBatch) { $tabItemBatch.Header = (Get-LocText 'TabBatch') }
    if ($tabItemSettings) { $tabItemSettings.Header = (Get-LocText 'TabSettings') }

    # Tab 1: Manual Route
    if ($lblManualRoutePointsHeader) { $lblManualRoutePointsHeader.Text = (Get-LocText 'ManualHeaderRoutePoints') }
    if ($lblManualOrigin) { $lblManualOrigin.Text = (Get-LocText 'ManualOrigin') }
    if ($lblManualWaypoints) { $lblManualWaypoints.Text = (Get-LocText 'ManualWaypoints') }
    if ($txtNewWaypoint) { $txtNewWaypoint.ToolTip = (Get-LocText 'ManualWaypointsTooltip') }
    if ($btnAddWaypoint) { $btnAddWaypoint.Content = (Get-LocText 'ManualBtnAdd') }
    if ($btnWpUp) { $btnWpUp.Content = (Get-LocText 'ManualBtnUp') }
    if ($btnWpDown) { $btnWpDown.Content = (Get-LocText 'ManualBtnDown') }
    if ($btnWpRemove) { $btnWpRemove.Content = (Get-LocText 'ManualBtnRemove') }
    if ($btnWpClear) { $btnWpClear.Content = (Get-LocText 'ManualBtnClear') }
    if ($lblManualDestination) { $lblManualDestination.Text = (Get-LocText 'ManualDestination') }
    if ($lblManualRouteName) { $lblManualRouteName.Text = (Get-LocText 'ManualRouteName') }
    if ($lblManualOptHeader) { $lblManualOptHeader.Text = (Get-LocText 'ManualHeaderOptimization') }
    if ($rbTypeFastest) { $rbTypeFastest.Content = (Get-LocText 'ManualOptFastest') }
    if ($rbTypeShortest) { $rbTypeShortest.Content = (Get-LocText 'ManualOptShortest') }
    if ($rbTypeEco) { $rbTypeEco.Content = (Get-LocText 'ManualOptEco') }
    if ($lblManualEmission) { $lblManualEmission.Text = (Get-LocText 'ManualEmission') }
    if ($chkTrafficAware) { $chkTrafficAware.Content = (Get-LocText 'ManualTrafficAware') }
    if ($btnCalculateManual -and $btnCalculateManual.IsEnabled) { $btnCalculateManual.Content = (Get-LocText 'ManualBtnCalculate') }
    if ($lblHeaderDist) { $lblHeaderDist.Text = (Get-LocText 'ManualStatDistance') }
    if ($lblHeaderDur) { $lblHeaderDur.Text = (Get-LocText 'ManualStatDuration') }
    if ($lblHeaderType) { $lblHeaderType.Text = (Get-LocText 'ManualStatType') }
    if ($lblManualType) {
        $curT = $lblManualType.Text
        if ($curT -match '(?i)fast|szyb|schnell') {
            $lblManualType.Text = switch ($script:CurrentLanguage) { 'de' { 'Schnellste' } 'pl' { 'Najszybsza' } default { 'Fastest' } }
        } elseif ($curT -match '(?i)short|kr[oó]t|k[uü]rz') {
            $lblManualType.Text = switch ($script:CurrentLanguage) { 'de' { 'Kürzeste' } 'pl' { 'Najkrótsza' } default { 'Shortest' } }
        } elseif ($curT -match '(?i)eco|eko') {
            $lblManualType.Text = switch ($script:CurrentLanguage) { 'pl' { 'Eko' } default { 'Eco' } }
        }
    }
    if ($lblManualStatus -and $lblManualStatus.Text -match '(?i)idle|bereit|gotow') { $lblManualStatus.Text = (Get-LocText 'ManualStatusIdle') }
    if ($lblMapPlaceholder -and $lblMapPlaceholder.Visibility -eq [System.Windows.Visibility]::Visible) { $lblMapPlaceholder.Text = (Get-LocText 'ManualMapPlaceholder') }
    if ($lblGoogleUrlDisplay -and $lblGoogleUrlDisplay.Text -match '(?i)no generated|kein link|brak') { $lblGoogleUrlDisplay.Text = (Get-LocText 'ManualNoUrl') }
    if ($btnOpenGoogleMaps) { $btnOpenGoogleMaps.Content = (Get-LocText 'ManualBtnGoogleMaps') }
    if ($btnCopyUrl) { $btnCopyUrl.Content = (Get-LocText 'ManualBtnCopyUrl') }
    if ($btnSaveMapAs) { $btnSaveMapAs.Content = (Get-LocText 'ManualBtnSaveMapAs') }

    # Tab 2: Batch Processing
    if ($lblBatchInputFile) { $lblBatchInputFile.Text = (Get-LocText 'BatchInputFile') }
    if ($btnBrowseBatchFile) { $btnBrowseBatchFile.Content = (Get-LocText 'BatchBtnBrowse') }
    if ($btnReloadBatchFile) { $btnReloadBatchFile.Content = (Get-LocText 'BatchBtnReload') }
    if ($lblBatchFileInfo -and $lblBatchFileInfo.Text -match '(?i)no file|keine datei|brak') { $lblBatchFileInfo.Text = (Get-LocText 'BatchNoFileLoaded') }
    if ($lblBatchDefaultRouteType) { $lblBatchDefaultRouteType.Text = (Get-LocText 'BatchDefaultRouteType') }
    if ($btnStartBatch) { $btnStartBatch.Content = (Get-LocText 'BatchBtnStart') }
    if ($btnStopBatch) { $btnStopBatch.Content = (Get-LocText 'BatchBtnStop') }
    if ($tabSubInput) { $tabSubInput.Header = (Get-LocText 'BatchTabInputPreview') }
    if ($tabSubResults) { $tabSubResults.Header = (Get-LocText 'BatchTabResults') }
    if ($tabSubPoints) { $tabSubPoints.Header = (Get-LocText 'BatchTabPoints') }
    if ($tabSubLog) { $tabSubLog.Header = (Get-LocText 'BatchTabLog') }

    # Batch DataGrid Columns
    if ($dgBatchResults -and $dgBatchResults.Columns.Count -ge 10) {
        $dgBatchResults.Columns[0].Header = (Get-LocText 'BatchColId')
        $dgBatchResults.Columns[1].Header = (Get-LocText 'BatchColName')
        $dgBatchResults.Columns[2].Header = (Get-LocText 'BatchColOrigin')
        $dgBatchResults.Columns[3].Header = (Get-LocText 'BatchColDestination')
        $dgBatchResults.Columns[4].Header = (Get-LocText 'BatchColWaypoints')
        $dgBatchResults.Columns[5].Header = (Get-LocText 'BatchColType')
        $dgBatchResults.Columns[6].Header = (Get-LocText 'BatchColDistance')
        $dgBatchResults.Columns[7].Header = (Get-LocText 'BatchColDuration')
        $dgBatchResults.Columns[8].Header = (Get-LocText 'BatchColStatus')
        $dgBatchResults.Columns[9].Header = (Get-LocText 'BatchColMap')
    }

    # Points DataGrid Columns
    if ($dgBatchPoints -and $dgBatchPoints.Columns.Count -ge 9) {
        $dgBatchPoints.Columns[0].Header = (Get-LocText 'PointsColRouteId')
        $dgBatchPoints.Columns[1].Header = (Get-LocText 'PointsColRouteName')
        $dgBatchPoints.Columns[2].Header = (Get-LocText 'PointsColOrder')
        $dgBatchPoints.Columns[3].Header = (Get-LocText 'PointsColType')
        $dgBatchPoints.Columns[4].Header = (Get-LocText 'PointsColOriginalAddress')
        $dgBatchPoints.Columns[5].Header = (Get-LocText 'PointsColGeocodedAddress')
        $dgBatchPoints.Columns[6].Header = (Get-LocText 'PointsColGeocodeStatus')
        $dgBatchPoints.Columns[7].Header = (Get-LocText 'PointsColMatchType')
        $dgBatchPoints.Columns[8].Header = (Get-LocText 'PointsColIsFallback')
        if ($dgBatchPoints.Columns.Count -ge 11) {
            $dgBatchPoints.Columns[9].Header = (Get-LocText 'PointsColLatitude')
            $dgBatchPoints.Columns[10].Header = (Get-LocText 'PointsColLongitude')
        }
    }

    if ($lblBatchProgressText -and $lblBatchProgressText.Text -match '(?i)ready|bereit|gotow') { $lblBatchProgressText.Text = (Get-LocText 'BatchProgressReady') }
    if ($btnOpenOutputDir) { $btnOpenOutputDir.Content = (Get-LocText 'BatchBtnOpenOutputDir') }
    if ($btnExportExcel) { $btnExportExcel.Content = (Get-LocText 'BatchBtnExportExcel') }
    if ($btnExportCsv) { $btnExportCsv.Content = (Get-LocText 'BatchBtnExportCsv') }
    if ($btnExportJson) { $btnExportJson.Content = (Get-LocText 'BatchBtnExportJson') }

    # Tab 3: Settings
    if ($lblSettingsApiHeader) { $lblSettingsApiHeader.Text = (Get-LocText 'SettingsHeaderApi') }
    if ($lblSettingsApiDesc) { $lblSettingsApiDesc.Text = (Get-LocText 'SettingsApiDesc') }
    if ($lblSettingsApiLabel) { $lblSettingsApiLabel.Text = (Get-LocText 'SettingsApiLabel') }
    if ($btnTestApiKey -and $btnTestApiKey.IsEnabled) { $btnTestApiKey.Content = (Get-LocText 'SettingsBtnTestKey') }
    if ($chkRememberKey) { $chkRememberKey.Content = (Get-LocText 'SettingsChkRemember') }
    if ($lblSettingsPrefHeader) { $lblSettingsPrefHeader.Text = (Get-LocText 'SettingsHeaderPreferences') }
    if ($lblSettingsDefaultRouteType) { $lblSettingsDefaultRouteType.Text = (Get-LocText 'SettingsDefaultRouteType') }
    if ($lblSettingsDefaultEmission) { $lblSettingsDefaultEmission.Text = (Get-LocText 'SettingsDefaultEmission') }
    if ($lblSettingsDefaultMapSize) { $lblSettingsDefaultMapSize.Text = (Get-LocText 'SettingsDefaultMapSize') }
    if ($lblSettingsOutputDir) { $lblSettingsOutputDir.Text = (Get-LocText 'SettingsOutputDir') }
    if ($btnBrowseSettingsOutputDir) { $btnBrowseSettingsOutputDir.Content = (Get-LocText 'SettingsBtnBrowseOutputDir') }
    if ($lblSettingsOverlayHeader) { $lblSettingsOverlayHeader.Text = (Get-LocText 'SettingsHeaderOverlay') }
    if ($lblSettingsOverlayDesc) { $lblSettingsOverlayDesc.Text = (Get-LocText 'SettingsOverlayDesc') }
    if ($chkEnableTopOverlay) { $chkEnableTopOverlay.Content = (Get-LocText 'SettingsOverlayTopEnable') }
    if ($chkEnableBottomOverlay) { $chkEnableBottomOverlay.Content = (Get-LocText 'SettingsOverlayBottomEnable') }
    if ($lblColPropName) { $lblColPropName.Text = (Get-LocText 'OverlayColProperty') }
    if ($lblColPropShow) { $lblColPropShow.Text = (Get-LocText 'OverlayColShow') }
    if ($lblColPropPanel) { $lblColPropPanel.Text = (Get-LocText 'OverlayColPanel') }
    if ($lblColPropAlign) { $lblColPropAlign.Text = (Get-LocText 'OverlayColAlign') }
    if ($lblColPropOrder) { $lblColPropOrder.Text = (Get-LocText 'OverlayColOrder') }
    if ($btnResetOverlayConfig) { $btnResetOverlayConfig.Content = (Get-LocText 'SettingsOverlayBtnReset') }

    foreach ($k in $script:OverlayPropKeys) {
        $lblCtrl = Get-Variable -Name "lblProp_$k" -ValueOnly -ErrorAction SilentlyContinue
        if ($lblCtrl) {
            $lblCtrl.Text = (Get-LocText "OverlayProp$k")
        }
    }

    if ($lblSettingsLangHeader) { $lblSettingsLangHeader.Text = (Get-LocText 'SettingsHeaderLanguage') }
    if ($lblSettingsLangLabel) { $lblSettingsLangLabel.Text = (Get-LocText 'SettingsLanguageLabel') }
    if ($btnOpenLangFile) { $btnOpenLangFile.Content = (Get-LocText 'SettingsBtnOpenLangFile') }
    if ($btnReloadLang) { $btnReloadLang.Content = (Get-LocText 'SettingsBtnReloadLang') }
    if ($lblSettingsThemeHeader) { $lblSettingsThemeHeader.Text = switch ($script:CurrentLanguage) { 'de' { 'Erscheinungsbild & Design' } 'pl' { 'Wygląd i motyw' } default { 'Appearance & Theme' } } }
    if ($lblSettingsThemeLabel) { $lblSettingsThemeLabel.Text = (Get-LocText 'SettingsThemeLabel') }
    if ($cmbSettingsTheme -and $cmbSettingsTheme.Items.Count -ge 2) {
        $cmbSettingsTheme.Items[0].Content = (Get-LocText 'ThemeDark')
        $cmbSettingsTheme.Items[1].Content = (Get-LocText 'ThemeLight')
    }
    if ($btnSaveSettings) { $btnSaveSettings.Content = (Get-LocText 'SettingsBtnSave') }
    if ($btnOpenLogFile) { $btnOpenLogFile.Content = (Get-LocText 'SettingsBtnOpenLog') }

    # Footer
    if ($lblFooterStatus -and $lblFooterStatus.Text -match '(?i)ready|bereit|gotow') { $lblFooterStatus.Text = (Get-LocText 'FooterReady') }
    if ($lblFooterVersion) { $lblFooterVersion.Text = (Get-LocText 'FooterVersion') }

    # Update API badge text if checking
    if ($lblApiBadge -and $lblApiBadge.Text -match '(?i)checking|prüfe|sprawdz') {
        $lblApiBadge.Text = (Get-LocText 'ApiBadgeChecking')
    }
}

function Set-CurrentApiKey([string]$Key) {
    $txtSettingsApiKey.Password = $Key
    $txtSettingsApiKeyVisible.Text = $Key
    $script:CurrentApiKey = $Key
}

function Get-CurrentApiKey {
    if ($txtSettingsApiKeyVisible.Visibility -eq [System.Windows.Visibility]::Visible) {
        return $txtSettingsApiKeyVisible.Text.Trim()
    }
    return $txtSettingsApiKey.Password.Trim()
}

function Write-BatchLog([string]$Message, [string]$Level = 'INFO') {
    Write-AppLog -Message $Message -Level $Level -ToBatchWindow
}

function Update-ApiStatusBadge($IsValid, [string]$Message) {
    $validBool = [bool]$IsValid
    if ($validBool) {
        $lblApiBadge.Text = 'API: Active'
        $lblApiBadge.Foreground = [System.Windows.Media.Brushes]::LightGreen
        $lblKeyTestResult.Text = "✓ $Message"
        $lblKeyTestResult.Foreground = [System.Windows.Media.Brushes]::LightGreen
    } else {
        $lblApiBadge.Text = 'API: Invalid'
        $lblApiBadge.Foreground = [System.Windows.Media.Brushes]::Salmon
        $lblKeyTestResult.Text = "✕ $Message"
        $lblKeyTestResult.Foreground = [System.Windows.Media.Brushes]::Salmon
    }
}

# Inicjalizacja wartości kontrolek z zapisanego configu
$script:CurrentApiKey = ''
if (-not [string]::IsNullOrWhiteSpace($script:Config.ApiKey)) {
    Set-CurrentApiKey -Key $script:Config.ApiKey
    $lblApiBadge.Text = 'API: Configured'
    $lblApiBadge.Foreground = [System.Windows.Media.Brushes]::LightGreen
}

$chkRememberKey.IsChecked = [bool]$script:Config.RememberApiKey
$txtSettingsOutputDir.Text = if ($script:Config.LastOutputFolder) { $script:Config.LastOutputFolder } else { Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'GoogleMapsRoutes\Results' }

Populate-LanguageDropdowns
Apply-AppLanguage -LanguageCode $script:CurrentLanguage
Set-AppTheme -Theme $script:Config.Theme

# Ustawienie domyślnych ComboBoxów
foreach ($item in $cmbDefaultRouteType.Items) {
    if ($item.Tag -eq $script:Config.DefaultRouteType) { $item.IsSelected = $true; break }
}
foreach ($item in $cmbDefaultEmission.Items) {
    if ($item.Tag -eq $script:Config.DefaultEmission) { $item.IsSelected = $true; break }
}
$targetMapTag = "$($script:Config.MapWidth)x$($script:Config.MapHeight)"
foreach ($item in $cmbDefaultMapSize.Items) {
    if ($item.Tag -eq $targetMapTag) { $item.IsSelected = $true; break }
}

# Ustawienie kontrolek nakładki mapy (Overlay)
if ($script:Config.OverlayConfig) {
    Set-OverlayConfigUi $script:Config.OverlayConfig
} else {
    Reset-OverlayConfigUi
}

$script:LastGeneratedMapPath = $null
$script:LastGoogleMapsUrl    = $null
$script:LoadedBatchData      = $null
$script:BatchResultsList     = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:BatchWorkerRunning   = $false
$script:BatchCancelRequested = $false
$script:LastDataDirectory    = if ($script:Config.LastInputFolder) { $script:Config.LastInputFolder } else { $null }

# Restore last used batch input file if available
if ($script:Config.LastInputPath -and (Test-Path $script:Config.LastInputPath)) {
    $txtBatchFilePath.Text = $script:Config.LastInputPath
    try {
        Load-BatchFilePreview -Path $script:Config.LastInputPath
    } catch { }
}

# ── 11. Zdarzenia: Ustawienia i Klucz API ────────────────────────────────────
$btnQuickSettings.Add_Click({
    $tabMain.SelectedItem = $tabItemSettings
})

if ($btnThemeToggle) {
    $btnThemeToggle.Add_Click({
        $newTheme = if ($script:CurrentTheme -eq 'Light') { 'Dark' } else { 'Light' }
        Set-AppTheme -Theme $newTheme
        $script:Config.Theme = $newTheme
        $routeType = if ($cmbDefaultRouteType.SelectedItem) { $cmbDefaultRouteType.SelectedItem.Tag -as [string] } else { $script:Config.DefaultRouteType }
        $emission = if ($cmbDefaultEmission.SelectedItem) { $cmbDefaultEmission.SelectedItem.Tag -as [string] } else { $script:Config.DefaultEmission }
        Save-AppConfig -ApiKey (Get-CurrentApiKey) -RememberApiKey $chkRememberKey.IsChecked `
            -OutputFolder $txtSettingsOutputDir.Text.Trim() `
            -LastInputFolder $script:Config.LastInputFolder -LastInputPath $script:Config.LastInputPath `
            -DefaultRouteType $routeType `
            -DefaultEmission $emission `
            -Language $script:CurrentLanguage `
            -Theme $newTheme
        Write-AppLog "Theme toggled to: $newTheme" "INFO"
    })
}

if ($cmbSettingsTheme) {
    $cmbSettingsTheme.Add_SelectionChanged({
        if ($script:SuppressThemeEvents) { return }
        $sel = $cmbSettingsTheme.SelectedItem
        if ($sel -and $sel.Tag) {
            $newTheme = [string]$sel.Tag
            if ($newTheme -ne $script:CurrentTheme) {
                Set-AppTheme -Theme $newTheme
                $script:Config.Theme = $newTheme
                $routeType = if ($cmbDefaultRouteType.SelectedItem) { $cmbDefaultRouteType.SelectedItem.Tag -as [string] } else { $script:Config.DefaultRouteType }
                $emission = if ($cmbDefaultEmission.SelectedItem) { $cmbDefaultEmission.SelectedItem.Tag -as [string] } else { $script:Config.DefaultEmission }
                Save-AppConfig -ApiKey (Get-CurrentApiKey) -RememberApiKey $chkRememberKey.IsChecked `
                    -OutputFolder $txtSettingsOutputDir.Text.Trim() `
                    -LastInputFolder $script:Config.LastInputFolder -LastInputPath $script:Config.LastInputPath `
                    -DefaultRouteType $routeType `
                    -DefaultEmission $emission `
                    -Language $script:CurrentLanguage `
                    -Theme $newTheme
                Write-AppLog "Theme changed to: $newTheme" "INFO"
            }
        }
    })
}

$cmbAppLanguage.Add_SelectionChanged({
    if ($script:SuppressLangEvents) { return }
    $sel = $cmbAppLanguage.SelectedItem
    if ($sel -and $sel.Tag) {
        $newLang = [string]$sel.Tag
        Apply-AppLanguage -LanguageCode $newLang
        Save-AppConfig -ApiKey (Get-CurrentApiKey) -RememberApiKey $chkRememberKey.IsChecked `
            -OutputFolder $txtSettingsOutputDir.Text.Trim() `
            -LastInputFolder $script:Config.LastInputFolder -LastInputPath $script:Config.LastInputPath `
            -DefaultRouteType ($cmbDefaultRouteType.SelectedItem.Tag -as [string]) `
            -DefaultEmission ($cmbDefaultEmission.SelectedItem.Tag -as [string]) `
            -Language $newLang
        Write-AppLog "Language changed to: $newLang (Google API: $script:CurrentGoogleLang)" "INFO"
    }
})

$cmbSettingsLanguage.Add_SelectionChanged({
    if ($script:SuppressLangEvents) { return }
    $sel = $cmbSettingsLanguage.SelectedItem
    if ($sel -and $sel.Tag) {
        $newLang = [string]$sel.Tag
        Apply-AppLanguage -LanguageCode $newLang
        Save-AppConfig -ApiKey (Get-CurrentApiKey) -RememberApiKey $chkRememberKey.IsChecked `
            -OutputFolder $txtSettingsOutputDir.Text.Trim() `
            -LastInputFolder $script:Config.LastInputFolder -LastInputPath $script:Config.LastInputPath `
            -DefaultRouteType ($cmbDefaultRouteType.SelectedItem.Tag -as [string]) `
            -DefaultEmission ($cmbDefaultEmission.SelectedItem.Tag -as [string]) `
            -Language $newLang
        Write-AppLog "Language changed to: $newLang (Google API: $script:CurrentGoogleLang)" "INFO"
    }
})

$btnOpenLangFile.Add_Click({
    if (-not (Test-Path $script:LocalizationFile)) {
        Load-LocalizationConfig
    }
    try {
        Start-Process -FilePath "notepad.exe" -ArgumentList "`"$script:LocalizationFile`""
    } catch {
        try { Start-Process -FilePath $script:LocalizationFile } catch {
            [System.Windows.MessageBox]::Show("Cannot open localization file:`r`n$($script:LocalizationFile)`r`n$($_.Exception.Message)", 'Error', 'OK', 'Error')
        }
    }
})

$btnReloadLang.Add_Click({
    Load-LocalizationConfig
    Populate-LanguageDropdowns
    Apply-AppLanguage -LanguageCode $script:CurrentLanguage
    $count = $script:LanguagesCatalog.Count
    $msg = (Get-LocText 'MsgLangReloaded') -f $count
    $title = (Get-LocText 'MsgLangReloadedTitle')
    [System.Windows.MessageBox]::Show($msg, $title, 'OK', 'Information')
})

$btnToggleKeyVisibility.Add_Click({
    if ($txtSettingsApiKeyVisible.Visibility -eq [System.Windows.Visibility]::Visible) {
        $txtSettingsApiKey.Password = $txtSettingsApiKeyVisible.Text
        $txtSettingsApiKeyVisible.Visibility = [System.Windows.Visibility]::Collapsed
        $txtSettingsApiKey.Visibility = [System.Windows.Visibility]::Visible
        $btnToggleKeyVisibility.Content = '👁 Show'
    } else {
        $txtSettingsApiKeyVisible.Text = $txtSettingsApiKey.Password
        $txtSettingsApiKey.Visibility = [System.Windows.Visibility]::Collapsed
        $txtSettingsApiKeyVisible.Visibility = [System.Windows.Visibility]::Visible
        $btnToggleKeyVisibility.Content = '🔒 Hide'
    }
})

$btnOpenLogFile.Add_Click({
    if (-not (Test-Path $script:LogFile)) {
        Write-AppLog "Creating new application log file." "INFO"
    }
    try {
        Start-Process -FilePath "notepad.exe" -ArgumentList "`"$script:LogFile`""
    } catch {
        try { Start-Process -FilePath $script:LogFile } catch {
            [System.Windows.MessageBox]::Show("Cannot open log file:`r`n$($script:LogFile)`r`n$($_.Exception.Message)", 'Log Open Error', 'OK', 'Error')
        }
    }
})

$btnTestApiKey.Add_Click({
    $key = Get-CurrentApiKey
    if ([string]::IsNullOrWhiteSpace($key)) {
        [System.Windows.MessageBox]::Show((Get-LocText 'MsgMissingApiKey'), (Get-LocText 'MsgMissingApiKeyTitle'), 'OK', 'Warning')
        return
    }

    $maskedKey = Get-MaskedKey $key
    Write-AppLog "Testing API key (Key: $maskedKey)..." "INFO"

    $lblKeyTestResult.Text = 'Testing API key...'
    $lblKeyTestResult.Foreground = [System.Windows.Media.Brushes]::SkyBlue
    $btnTestApiKey.IsEnabled = $false

    # Asynchroniczny test połączenia z Google API w osobnym runspace
    $testScript = {
        param($apiKeyToTest, $apiLanguage = 'en')
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls
        if ([string]::IsNullOrWhiteSpace($apiKeyToTest)) {
            return [PSCustomObject]@{ Valid = $false; Message = 'API key is empty.' }
        }
        try {
            $lang = if ($apiLanguage) { $apiLanguage } else { 'en' }
            $Url = "https://maps.googleapis.com/maps/api/geocode/json?address=Warszawa&language=$lang&key=$apiKeyToTest"
            $Resp = Invoke-RestMethod -Uri $Url -Method GET -TimeoutSec 15
            if ($Resp.status -eq 'OK' -or $Resp.status -eq 'ZERO_RESULTS') {
                return [PSCustomObject]@{ Valid = $true; Message = 'Google Maps API key is valid and active.' }
            }
            elseif ($Resp.status -eq 'REQUEST_DENIED') {
                $msg = if ($Resp.error_message) { $Resp.error_message } else { 'Request denied by Google API.' }
                return [PSCustomObject]@{ Valid = $false; Message = "Unauthorized: $msg" }
            }
            else {
                return [PSCustomObject]@{ Valid = $false; Message = "API Status: $($Resp.status)" }
            }
        }
        catch {
            return [PSCustomObject]@{ Valid = $false; Message = "Connection error: $($_.Exception.Message)" }
        }
    }

    $psTest = [PowerShell]::Create().AddScript($testScript).AddArgument($key).AddArgument($script:CurrentGoogleLang)
    $testHandle = $psTest.BeginInvoke()
    $testTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $testTimer.Interval = [TimeSpan]::FromMilliseconds(100)
    $script:ActiveTestTimer = $testTimer
    $script:ActiveTestPs = $psTest
    $script:ActiveTestHandle = $testHandle
    $script:TestTimerTicks = 0

    $testTimer.Add_Tick({
        $localTestHandle = $script:ActiveTestHandle
        $localTestPs     = $script:ActiveTestPs
        $script:TestTimerTicks++
        if ($localTestHandle -and $localTestHandle.IsCompleted) {
            if ($script:ActiveTestTimer) { try { $script:ActiveTestTimer.Stop() } catch { } }
            $btnTestApiKey.IsEnabled = $true
            try {
                $res = $localTestPs.EndInvoke($localTestHandle)
                $testResult = $res[0]
                $isValid = [bool]$testResult.Valid
                $msg = [string]$testResult.Message
                Update-ApiStatusBadge -IsValid $isValid -Message $msg
                $logLevel = if ($isValid) { 'OK' } else { 'WARN' }
                Write-AppLog "API key test completed: Valid=$isValid, Message='$msg'" $logLevel
            }
            catch {
                $errDetail = $_.Exception.Message
                Update-ApiStatusBadge -IsValid $false -Message "Test error: $errDetail"
                Write-AppLog "Exception while retrieving API test result: $($_.Exception.ToString())" "ERROR"
            }
            finally {
                $localTestPs.Dispose()
            }
        }
        elseif ($script:TestTimerTicks -ge 200) { # 20s timeout limit
            if ($script:ActiveTestTimer) { try { $script:ActiveTestTimer.Stop() } catch { } }
            $btnTestApiKey.IsEnabled = $true
            Update-ApiStatusBadge -IsValid $false -Message "Google API response timeout (20s)."
            Write-AppLog "API key test timeout (20s watchdog timeout)." "WARN"
            try { $localTestPs.Stop(); $localTestPs.Dispose() } catch { }
        }
    })
    $testTimer.Start()
})

$btnBrowseSettingsOutputDir.Add_Click({
    $dlg = [System.Windows.Forms.FolderBrowserDialog]::new()
    $dlg.Description = 'Select default folder for calculation results'
    $dlg.SelectedPath = $txtSettingsOutputDir.Text
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtSettingsOutputDir.Text = $dlg.SelectedPath
    }
})

$btnSaveSettings.Add_Click({
    $key = Get-CurrentApiKey
    $remember = [bool]$chkRememberKey.IsChecked
    $outDir = $txtSettingsOutputDir.Text.Trim()
    $routeType = ($cmbDefaultRouteType.SelectedItem.Tag -as [string])
    $emission = ($cmbDefaultEmission.SelectedItem.Tag -as [string])
    $dims = ($cmbDefaultMapSize.SelectedItem.Tag -as [string]) -split 'x'
    $mapW = [int]$dims[0]
    $mapH = [int]$dims[1]

    $langToSave = if ($cmbSettingsLanguage.SelectedItem) { [string]$cmbSettingsLanguage.SelectedItem.Tag } else { $script:CurrentLanguage }
    $themeToSave = if ($cmbSettingsTheme.SelectedItem) { [string]$cmbSettingsTheme.SelectedItem.Tag } else { $script:CurrentTheme }
    $overlayCfg = Get-CurrentOverlayConfig
    Save-AppConfig -ApiKey $key -RememberApiKey $remember -OutputFolder $outDir `
        -LastInputFolder $script:Config.LastInputFolder -LastInputPath $script:Config.LastInputPath `
        -DefaultRouteType $routeType -DefaultEmission $emission -MapWidth $mapW -MapHeight $mapH -Language $langToSave `
        -OverlayConfig $overlayCfg -Theme $themeToSave
    $script:Config.OverlayConfig = $overlayCfg
    $script:Config.Theme = $themeToSave

    Set-CurrentApiKey -Key $key
    if ($remember -and -not [string]::IsNullOrWhiteSpace($key)) {
        $lblApiBadge.Text = 'API: Configured'
        $lblApiBadge.Foreground = [System.Windows.Media.Brushes]::LightGreen
    }
    elseif (-not $remember -and -not [string]::IsNullOrWhiteSpace($key)) {
        $lblApiBadge.Text = 'API: Session key (unsaved)'
        $lblApiBadge.Foreground = [System.Windows.Media.Brushes]::SkyBlue
    }
    else {
        $lblApiBadge.Text = 'API: Unverified'
        $lblApiBadge.Foreground = [System.Windows.Media.Brushes]::Orange
    }

    [System.Windows.MessageBox]::Show((Get-LocText 'MsgSettingsSaved'), (Get-LocText 'MsgSettingsSavedTitle'), 'OK', 'Information')
})

if ($btnResetOverlayConfig) {
    $btnResetOverlayConfig.Add_Click({
        Reset-OverlayConfigUi
    })
}

# ── 12. Zdarzenia: Tab 1 (Manual Input) ───────────────────────────────────────
$btnClearManualStart.Add_Click({ $txtManualStart.Clear() })
$btnClearManualEnd.Add_Click({ $txtManualEnd.Clear() })

$btnAddWaypoint.Add_Click({
    $wp = $txtNewWaypoint.Text.Trim()
    if (-not [string]::IsNullOrWhiteSpace($wp)) {
        if ($lstWaypoints.Items.Count -ge 25) {
            [System.Windows.MessageBox]::Show((Get-LocText 'MsgMaxWaypoints'), (Get-LocText 'MsgMaxWaypointsTitle'), 'OK', 'Warning')
            return
        }
        $null = $lstWaypoints.Items.Add($wp)
        $txtNewWaypoint.Clear()
    }
})

$txtNewWaypoint.Add_KeyDown({
    if ($_.Key -eq [System.Windows.Input.Key]::Enter) {
        $btnAddWaypoint.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
    }
})

$btnWpRemove.Add_Click({
    if ($lstWaypoints.SelectedIndex -ge 0) {
        $lstWaypoints.Items.RemoveAt($lstWaypoints.SelectedIndex)
    }
})

$btnWpClear.Add_Click({
    $lstWaypoints.Items.Clear()
})

$btnWpUp.Add_Click({
    $idx = $lstWaypoints.SelectedIndex
    if ($idx -gt 0) {
        $item = $lstWaypoints.Items[$idx]
        $lstWaypoints.Items.RemoveAt($idx)
        $lstWaypoints.Items.Insert($idx - 1, $item)
        $lstWaypoints.SelectedIndex = $idx - 1
    }
})

$btnWpDown.Add_Click({
    $idx = $lstWaypoints.SelectedIndex
    if ($idx -ge 0 -and $idx -lt ($lstWaypoints.Items.Count - 1)) {
        $item = $lstWaypoints.Items[$idx]
        $lstWaypoints.Items.RemoveAt($idx)
        $lstWaypoints.Items.Insert($idx + 1, $item)
        $lstWaypoints.SelectedIndex = $idx + 1
    }
})

$rbTypeEco.Add_Checked({ $pnlEmission.Visibility = [System.Windows.Visibility]::Visible })
$rbTypeFastest.Add_Checked({ $pnlEmission.Visibility = [System.Windows.Visibility]::Collapsed })
$rbTypeShortest.Add_Checked({ $pnlEmission.Visibility = [System.Windows.Visibility]::Collapsed })

$btnCalculateManual.Add_Click({
    $apiKey = Get-CurrentApiKey
    if ([string]::IsNullOrWhiteSpace($apiKey)) {
        [System.Windows.MessageBox]::Show((Get-LocText 'MsgMissingApiKeyPrompt'), (Get-LocText 'MsgMissingApiKeyTitle'), 'OK', 'Warning')
        $tabMain.SelectedIndex = 2
        return
    }

    $start = $txtManualStart.Text.Trim()
    $end = $txtManualEnd.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($start) -or [string]::IsNullOrWhiteSpace($end)) {
        [System.Windows.MessageBox]::Show((Get-LocText 'MsgMissingData'), (Get-LocText 'MsgMissingDataTitle'), 'OK', 'Warning')
        return
    }

    $waypoints = [System.Collections.Generic.List[string]]::new()
    foreach ($item in $lstWaypoints.Items) {
        $waypoints.Add([string]$item)
    }

    $routeType = if ($rbTypeShortest.IsChecked) { 'Shortest' } elseif ($rbTypeEco.IsChecked) { 'Eco' } else { 'Fastest' }
    $emission = ($cmbEmission.SelectedItem.Tag -as [string])
    if ([string]::IsNullOrWhiteSpace($emission)) { $emission = 'GASOLINE' }
    $trafficAware = [bool]$chkTrafficAware.IsChecked
    $name = $txtManualName.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($name)) { $name = "Route $start -> $end" }

    $outDir = $txtSettingsOutputDir.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($outDir)) { $outDir = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'GoogleMapsRoutes\Results' }
    if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

    if ($txtBatchFilePath.Text -and (Test-Path $txtBatchFilePath.Text.Trim())) {
        $script:Config.LastInputPath = $txtBatchFilePath.Text.Trim()
        $script:Config.LastInputFolder = Split-Path $script:Config.LastInputPath -Parent
        $script:LastDataDirectory = $script:Config.LastInputFolder
        Save-AppConfig -ApiKey $apiKey -RememberApiKey $chkRememberKey.IsChecked -OutputFolder $outDir `
            -LastInputFolder $script:Config.LastInputFolder -LastInputPath $script:Config.LastInputPath `
            -DefaultRouteType ($cmbDefaultRouteType.SelectedItem.Tag -as [string]) `
            -DefaultEmission ($cmbDefaultEmission.SelectedItem.Tag -as [string])
    }

    $btnCalculateManual.IsEnabled = $false
    $btnCalculateManual.Content = '⏳ CALCULATING ROUTE...'
    $lblManualStatus.Text = 'Geocoding and calculating...'
    $lblManualStatus.Foreground = [System.Windows.Media.Brushes]::SkyBlue
    $lblFooterStatus.Text = 'Calculating manual route...'

    Write-AppLog "Started manual route calculation: '$start' -> '$end' (Waypoints: $($waypoints.Count), Type: $routeType, Engine: $emission, LiveTraffic: $trafficAware)..." "INFO"

    # $script:ManualCalcAsync is defined at top-level (section 6b) — used directly below
    $psCmd = New-WorkerPowerShell -ScriptBlock $script:ManualCalcAsync
    $overlayCfgJson = (Get-CurrentOverlayConfig | ConvertTo-Json -Depth 6 -Compress)
    $psCmd.AddArgument($start).AddArgument($end).AddArgument($waypoints).AddArgument($routeType).AddArgument($emission).AddArgument($trafficAware).AddArgument($name).AddArgument($apiKey).AddArgument($outDir).AddArgument($script:LogFile).AddArgument($script:CurrentGoogleLang).AddArgument($overlayCfgJson) | Out-Null

    try {
        $asyncHandle = $psCmd.BeginInvoke()
    } catch {
        Write-AppLog "CRITICAL: BeginInvoke() threw exception: $($_.Exception.Message)" "ERROR"
        $btnCalculateManual.IsEnabled = $true
        $btnCalculateManual.Content = '🚀 CALCULATE ROUTE & DOWNLOAD MAP'
        $lblManualStatus.Text = '✕ Launch error'
        $lblManualStatus.Foreground = [System.Windows.Media.Brushes]::Salmon
        $lblFooterStatus.Text = "Error: $($_.Exception.Message)"
        return
    }
    if (-not $asyncHandle) {
        Write-AppLog "CRITICAL: BeginInvoke() returned null — runspace may be invalid." "ERROR"
        $btnCalculateManual.IsEnabled = $true
        $btnCalculateManual.Content = '🚀 CALCULATE ROUTE & DOWNLOAD MAP'
        $lblManualStatus.Text = '✕ Runspace error'
        $lblManualStatus.Foreground = [System.Windows.Media.Brushes]::Salmon
        return
    }
    $timer = [System.Windows.Threading.DispatcherTimer]::new()
    $timer.Interval = [TimeSpan]::FromMilliseconds(150)
    $script:ActiveManualTimer = $timer
    $script:ActiveManualPs = $psCmd
    $script:ActiveManualAsyncHandle = $asyncHandle
    $script:ManualTimerTicks = 0

    Write-AppLog "Worker started (BeginInvoke). IsCompleted=$($asyncHandle.IsCompleted)" "INFO"

    $timer.Add_Tick({
        $localHandle = $script:ActiveManualAsyncHandle
        $localPs     = $script:ActiveManualPs
        $script:ManualTimerTicks++
        if ($localHandle -and $localHandle.IsCompleted) {
            if ($script:ActiveManualTimer) { try { $script:ActiveManualTimer.Stop() } catch { } }
            $btnCalculateManual.IsEnabled = $true
            $btnCalculateManual.Content = '🚀 CALCULATE ROUTE & DOWNLOAD MAP'

            # Log any stream errors from the worker runspace before inspecting result
            foreach ($streamErr in $localPs.Streams.Error) {
                Write-AppLog "[Stream.Error] $($streamErr.Exception.Message) @ $($streamErr.InvocationInfo.PositionMessage)" "ERROR"
            }

            try {
                $res = $localPs.EndInvoke($localHandle)
                $calc = $res[0]
                if ($calc.Success) {
                    $lblManualDist.Text = "$($calc.DistanceKm) km"
                    $lblManualTime.Text = "$($calc.DurationMin) min"
                    $lblManualType.Text = switch ($script:CurrentLanguage) {
                        'de' { if ($calc.RouteType -eq 'Fastest') { 'Schnellste' } elseif ($calc.RouteType -eq 'Shortest') { 'Kürzeste' } else { 'Eco' } }
                        'pl' { if ($calc.RouteType -eq 'Fastest') { 'Najszybsza' } elseif ($calc.RouteType -eq 'Shortest') { 'Najkrótsza' } else { 'Eko' } }
                        default { [string]$calc.RouteType }
                    }
                    $lblManualStatus.Text = '✓ Success'
                    $lblManualStatus.Foreground = [System.Windows.Media.Brushes]::LightGreen
                    $lblFooterStatus.Text = "Route ready: $($calc.DistanceKm) km, $($calc.DurationMin) min"
                    Write-AppLog "Manual route calculation completed successfully: $($calc.DistanceKm) km, $($calc.DurationMin) min (Map file: $($calc.MapPath))" "OK"

                    $script:LastGoogleMapsUrl = $calc.GoogleMapsUrl
                    $lblGoogleUrlDisplay.Text = $calc.GoogleMapsUrl
                    $btnOpenGoogleMaps.IsEnabled = $true
                    $btnCopyUrl.IsEnabled = $true

                    if ($calc.MapPath -and (Test-Path $calc.MapPath)) {
                        $script:LastGeneratedMapPath = $calc.MapPath
                        $btnSaveMapAs.IsEnabled = $true
                        $lblMapPlaceholder.Visibility = [System.Windows.Visibility]::Collapsed

                        $imgBytes = [System.IO.File]::ReadAllBytes($calc.MapPath)
                        $ms = [System.IO.MemoryStream]::new($imgBytes)
                        $bi = [System.Windows.Media.Imaging.BitmapImage]::new()
                        $bi.BeginInit()
                        $bi.StreamSource = $ms
                        $bi.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
                        $bi.EndInit()
                        $bi.Freeze()
                        $imgMapPreview.Source = $bi
                    }
                } else {
                    $lblManualStatus.Text = '✕ Error'
                    $lblManualStatus.Foreground = [System.Windows.Media.Brushes]::Salmon
                    $lblFooterStatus.Text = "Error: $($calc.Error)"
                    Write-AppLog "Manual route calculation failed: $($calc.Error)" "ERROR"
                    [System.Windows.MessageBox]::Show($calc.Error, 'Route Error', 'OK', 'Error')
                }
            }
            catch {
                $errDetail = $_.Exception.ToString()
                $lblManualStatus.Text = '✕ Error'
                $lblManualStatus.Foreground = [System.Windows.Media.Brushes]::Salmon
                $lblFooterStatus.Text = "Exception: $($_.Exception.Message)"
                Write-AppLog "UI exception while reading route result: $errDetail" "ERROR"
                [System.Windows.MessageBox]::Show($_.Exception.Message, 'Error', 'OK', 'Error')
            }
            finally {
                $localPs.Dispose()
            }
        }
        elseif ($script:ManualTimerTicks -ge 400) { # 60 seconds watchdog timeout
            if ($script:ActiveManualTimer) { try { $script:ActiveManualTimer.Stop() } catch { } }
            $btnCalculateManual.IsEnabled = $true
            $btnCalculateManual.Content = '🚀 CALCULATE ROUTE & DOWNLOAD MAP'
            $lblManualStatus.Text = '✕ Timeout (60s)'
            $lblManualStatus.Foreground = [System.Windows.Media.Brushes]::Salmon
            $lblFooterStatus.Text = 'Route calculation timed out (60s).'
            Write-AppLog "Manual route calculation timed out (60s watchdog timeout). Ticks=$($script:ManualTimerTicks)" "WARN"
            try { $localPs.Stop(); $localPs.Dispose() } catch { }
        }
    })
    $timer.Start()
})

$btnOpenGoogleMaps.Add_Click({
    if ($script:LastGoogleMapsUrl) {
        Start-Process $script:LastGoogleMapsUrl
    }
})

$btnCopyUrl.Add_Click({
    if ($script:LastGoogleMapsUrl) {
        [System.Windows.Clipboard]::SetText($script:LastGoogleMapsUrl)
        [System.Windows.MessageBox]::Show((Get-LocText 'MsgUrlCopied'), (Get-LocText 'MsgUrlCopiedTitle'), 'OK', 'Information')
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
            [System.Windows.MessageBox]::Show(((Get-LocText 'MsgMapSaved') -f $dlg.FileName), (Get-LocText 'MsgMapSavedTitle'), 'OK', 'Information')
        }
    }
})

# ── 13. Zdarzenia: Tab 2 (Batch Processing) ──────────────────────────────────
function Load-BatchFilePreview([string]$Path) {
    if (-not (Test-Path $Path)) { return }
    try {
        $data = Import-RouteDataFile -Path $Path
        $script:LoadedBatchData = $data

        # Configure dynamic columns and clear previous items
        $dgBatchInput.ItemsSource = $null
        $dgBatchInput.Columns.Clear()

        if ($data.Mode -eq 'SequentialStops') {
            $lblBatchFileInfo.Text = "Format: $($data.Format) | Mode: Sequential Stops (1 Multi-point Route, $($data.TotalCount) Stops) | Total: $($data.TotalCount) stops"
            $lblBatchFileInfo.Foreground = [System.Windows.Media.Brushes]::LightGreen

            $colStep = [System.Windows.Controls.DataGridTextColumn]::new()
            $colStep.Header = "#"
            $colStep.Binding = [System.Windows.Data.Binding]::new("Step")
            $colStep.Width = [System.Windows.Controls.DataGridLength]::new(55)
            $dgBatchInput.Columns.Add($colStep)

            $colRole = [System.Windows.Controls.DataGridTextColumn]::new()
            $colRole.Header = "Role / Point Type"
            $colRole.Binding = [System.Windows.Data.Binding]::new("Role")
            $colRole.Width = [System.Windows.Controls.DataGridLength]::new(180)
            $dgBatchInput.Columns.Add($colRole)

            $colAddr = [System.Windows.Controls.DataGridTextColumn]::new()
            $colAddr.Header = "Address / Location"
            $colAddr.Binding = [System.Windows.Data.Binding]::new("Address")
            $colAddr.Width = [System.Windows.Controls.DataGridLength]::new(340)
            $dgBatchInput.Columns.Add($colAddr)

            $colRaw = [System.Windows.Controls.DataGridTextColumn]::new()
            $colRaw.Header = "Source Record Data"
            $colRaw.Binding = [System.Windows.Data.Binding]::new("RawSummary")
            $colRaw.Width = [System.Windows.Controls.DataGridLength]::new(1, [System.Windows.Controls.DataGridLengthUnitType]::Star)
            $dgBatchInput.Columns.Add($colRaw)

            $previewItems = [System.Collections.Generic.List[PSCustomObject]]::new()
            $stopsCount = $data.Stops.Count
            for ($i = 0; $i -lt $stopsCount; $i++) {
                $st = $data.Stops[$i]
                $role = if ($i -eq 0) { "🟢 Origin (Start)" }
                        elseif ($i -eq ($stopsCount - 1)) { "🔴 Destination (End)" }
                        else { "🟡 Waypoint $i" }

                $rawProps = @()
                if ($st.Raw) {
                    foreach ($p in $st.Raw.PSObject.Properties) {
                        $rawProps += "$($p.Name)=$($p.Value)"
                    }
                }
                $rawSummaryText = $rawProps -join '; '

                $previewItems.Add([PSCustomObject]@{
                    Step       = ($i + 1)
                    Role       = $role
                    Address    = [string]$st.Address
                    RawSummary = $rawSummaryText
                })
            }
            $dgBatchInput.ItemsSource = $previewItems
            Write-BatchLog "Loaded file: $Path ($($data.TotalCount) sequential stops, format: $($data.Format), mode: $($data.Mode))" "OK"
        }
        else {
            $lblBatchFileInfo.Text = "Format: $($data.Format) | Mode: Route List | Total Routes: $($data.TotalCount)"
            $lblBatchFileInfo.Foreground = [System.Windows.Media.Brushes]::LightGreen

            $colId = [System.Windows.Controls.DataGridTextColumn]::new()
            $colId.Header = "ID"
            $colId.Binding = [System.Windows.Data.Binding]::new("Id")
            $colId.Width = [System.Windows.Controls.DataGridLength]::new(50)
            $dgBatchInput.Columns.Add($colId)

            $colName = [System.Windows.Controls.DataGridTextColumn]::new()
            $colName.Header = "Route Name"
            $colName.Binding = [System.Windows.Data.Binding]::new("Name")
            $colName.Width = [System.Windows.Controls.DataGridLength]::new(180)
            $dgBatchInput.Columns.Add($colName)

            $colStart = [System.Windows.Controls.DataGridTextColumn]::new()
            $colStart.Header = "Origin (Start)"
            $colStart.Binding = [System.Windows.Data.Binding]::new("Start")
            $colStart.Width = [System.Windows.Controls.DataGridLength]::new(200)
            $dgBatchInput.Columns.Add($colStart)

            $colEnd = [System.Windows.Controls.DataGridTextColumn]::new()
            $colEnd.Header = "Destination (End)"
            $colEnd.Binding = [System.Windows.Data.Binding]::new("End")
            $colEnd.Width = [System.Windows.Controls.DataGridLength]::new(200)
            $dgBatchInput.Columns.Add($colEnd)

            $colWpCount = [System.Windows.Controls.DataGridTextColumn]::new()
            $colWpCount.Header = "Waypoints"
            $colWpCount.Binding = [System.Windows.Data.Binding]::new("WaypointCount")
            $colWpCount.Width = [System.Windows.Controls.DataGridLength]::new(80)
            $dgBatchInput.Columns.Add($colWpCount)

            $colWpText = [System.Windows.Controls.DataGridTextColumn]::new()
            $colWpText.Header = "Intermediate Stops"
            $colWpText.Binding = [System.Windows.Data.Binding]::new("WaypointsText")
            $colWpText.Width = [System.Windows.Controls.DataGridLength]::new(250)
            $dgBatchInput.Columns.Add($colWpText)

            $colType = [System.Windows.Controls.DataGridTextColumn]::new()
            $colType.Header = "Route Type"
            $colType.Binding = [System.Windows.Data.Binding]::new("RouteType")
            $colType.Width = [System.Windows.Controls.DataGridLength]::new(100)
            $dgBatchInput.Columns.Add($colType)

            $previewItems = [System.Collections.Generic.List[PSCustomObject]]::new()
            foreach ($r in @($data.Routes)) {
                $wpText = if ($r.Waypoints -and @($r.Waypoints).Count -gt 0) {
                    (@($r.Waypoints) -join ' | ')
                } else {
                    '(none)'
                }
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
            Write-BatchLog "Loaded file: $Path ($($data.TotalCount) routes, format: $($data.Format), mode: $($data.Mode))" "OK"
        }
    }
    catch {
        $lblBatchFileInfo.Text = "Load error: $($_.Exception.Message)"
        $lblBatchFileInfo.Foreground = [System.Windows.Media.Brushes]::Salmon
        Write-BatchLog "Load error: $($_.Exception.Message)" "ERROR"
    }
}

$btnBrowseBatchFile.Add_Click({
    $initDir = $null
    if (-not [string]::IsNullOrWhiteSpace($txtBatchFilePath.Text) -and (Test-Path $txtBatchFilePath.Text.Trim())) {
        $initDir = Split-Path $txtBatchFilePath.Text.Trim() -Parent
    }
    elseif (-not [string]::IsNullOrWhiteSpace($txtBatchFilePath.Text) -and (Test-Path (Split-Path $txtBatchFilePath.Text.Trim() -Parent))) {
        $initDir = Split-Path $txtBatchFilePath.Text.Trim() -Parent
    }
    elseif ($script:Config.LastInputFolder -and (Test-Path $script:Config.LastInputFolder)) {
        $initDir = $script:Config.LastInputFolder
    }
    elseif ($script:Config.LastInputPath -and (Test-Path (Split-Path $script:Config.LastInputPath -Parent))) {
        $initDir = Split-Path $script:Config.LastInputPath -Parent
    }

    $file = Select-InputDataFile -InitialDirectory $initDir
    if ($file) {
        $txtBatchFilePath.Text = $file
        $script:Config.LastInputPath = $file
        $script:Config.LastInputFolder = Split-Path $file -Parent
        $script:LastDataDirectory = $script:Config.LastInputFolder

        Save-AppConfig -ApiKey (Get-CurrentApiKey) `
            -RememberApiKey $chkRememberKey.IsChecked `
            -OutputFolder $txtSettingsOutputDir.Text.Trim() `
            -LastInputFolder $script:Config.LastInputFolder `
            -LastInputPath $script:Config.LastInputPath `
            -DefaultRouteType ($cmbDefaultRouteType.SelectedItem.Tag -as [string]) `
            -DefaultEmission ($cmbDefaultEmission.SelectedItem.Tag -as [string])

        Load-BatchFilePreview -Path $file
    }
})

$btnReloadBatchFile.Add_Click({
    if ($txtBatchFilePath.Text) {
        $path = $txtBatchFilePath.Text.Trim()
        if (Test-Path $path) {
            $script:Config.LastInputPath = $path
            $script:Config.LastInputFolder = Split-Path $path -Parent
            $script:LastDataDirectory = $script:Config.LastInputFolder

            Save-AppConfig -ApiKey (Get-CurrentApiKey) `
                -RememberApiKey $chkRememberKey.IsChecked `
                -OutputFolder $txtSettingsOutputDir.Text.Trim() `
                -LastInputFolder $script:Config.LastInputFolder `
                -LastInputPath $script:Config.LastInputPath `
                -DefaultRouteType ($cmbDefaultRouteType.SelectedItem.Tag -as [string]) `
                -DefaultEmission ($cmbDefaultEmission.SelectedItem.Tag -as [string])
        }
        Load-BatchFilePreview -Path $path
    }
})

$btnStartBatch.Add_Click({
    $apiKey = Get-CurrentApiKey
    if ([string]::IsNullOrWhiteSpace($apiKey)) {
        [System.Windows.MessageBox]::Show((Get-LocText 'MsgMissingApiKeyPrompt'), (Get-LocText 'MsgMissingApiKeyTitle'), 'OK', 'Warning')
        $tabMain.SelectedIndex = 2
        return
    }

    if ($null -eq $script:LoadedBatchData -or $script:LoadedBatchData.Routes.Count -eq 0) {
        [System.Windows.MessageBox]::Show((Get-LocText 'MsgNoDataFile'), (Get-LocText 'MsgNoDataFileTitle'), 'OK', 'Warning')
        return
    }

    $outDir = $txtSettingsOutputDir.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($outDir)) { $outDir = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'GoogleMapsRoutes\Results' }
    if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

    if ($txtBatchFilePath.Text -and (Test-Path $txtBatchFilePath.Text.Trim())) {
        $script:Config.LastInputPath = $txtBatchFilePath.Text.Trim()
        $script:Config.LastInputFolder = Split-Path $script:Config.LastInputPath -Parent
        $script:LastDataDirectory = $script:Config.LastInputFolder
        Save-AppConfig -ApiKey $apiKey -RememberApiKey $chkRememberKey.IsChecked -OutputFolder $outDir `
            -LastInputFolder $script:Config.LastInputFolder -LastInputPath $script:Config.LastInputPath `
            -DefaultRouteType ($cmbDefaultRouteType.SelectedItem.Tag -as [string]) `
            -DefaultEmission ($cmbDefaultEmission.SelectedItem.Tag -as [string])
    }

    $defaultRouteType = ($cmbBatchRouteType.SelectedItem.Tag -as [string])
    if ([string]::IsNullOrWhiteSpace($defaultRouteType)) { $defaultRouteType = 'Fastest' }

    $script:BatchWorkerRunning = $true
    $btnStartBatch.IsEnabled = $false
    $btnStopBatch.IsEnabled = $true
    $btnBrowseBatchFile.IsEnabled = $false
    $btnReloadBatchFile.IsEnabled = $false

    $script:BatchResultsList.Clear()
    $dgBatchResults.ItemsSource = $null
    if ($dgBatchPoints) { $dgBatchPoints.ItemsSource = $null }
    $pbBatchProgress.Value = 0
    $lblBatchProgressText.Text = "Starting batch processing (0 / $($script:LoadedBatchData.Routes.Count))..."
    $lblBatchStats.Text = "Success: 0 | Errors: 0"

    Write-BatchLog "=== Starting batch processing ($($script:LoadedBatchData.Routes.Count) routes) ===" "INFO"

    # Switch to Activity Log tab so user sees live execution
    if ($tabBatchSub) {
        $tabBatchSub.SelectedIndex = 2
    }

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
        $overlayCfgJson = (Get-CurrentOverlayConfig | ConvertTo-Json -Depth 6 -Compress)
        $psCmdBatch.AddArgument($routesToProcess).AddArgument($apiKey).AddArgument($outDir).AddArgument($defaultRouteType).AddArgument($syncState).AddArgument($script:LogFile).AddArgument($script:CurrentGoogleLang).AddArgument($overlayCfgJson) | Out-Null
        $asyncBatchHandle = $psCmdBatch.BeginInvoke()
    }
    catch {
        Write-BatchLog "CRITICAL: Could not start batch worker: $($_.Exception.Message)" "ERROR"
        $btnStartBatch.IsEnabled = $true
        $btnStopBatch.IsEnabled = $false
        $btnBrowseBatchFile.IsEnabled = $true
        $btnReloadBatchFile.IsEnabled = $true
        $script:BatchWorkerRunning = $false
        $lblBatchProgressText.Text = "Launch error"
        return
    }

    if (-not $asyncBatchHandle) {
        Write-BatchLog "CRITICAL: BeginInvoke returned null handle." "ERROR"
        $btnStartBatch.IsEnabled = $true
        $btnStopBatch.IsEnabled = $false
        $btnBrowseBatchFile.IsEnabled = $true
        $btnReloadBatchFile.IsEnabled = $true
        $script:BatchWorkerRunning = $false
        return
    }

    $timerBatch = [System.Windows.Threading.DispatcherTimer]::new()
    $timerBatch.Interval = [TimeSpan]::FromMilliseconds(150)
    $script:ActiveBatchTimer = $timerBatch
    $script:ActiveBatchPs = $psCmdBatch
    $script:ActiveBatchAsyncHandle = $asyncBatchHandle

    # Direct script-scope reference (NO .GetNewClosure()!)
    $timerBatch.Add_Tick({
        $localBatchHandle = $script:ActiveBatchAsyncHandle
        $localBatchPs     = $script:ActiveBatchPs
        $localSyncState   = $script:SyncState

        if (-not $localSyncState) { return }

        # Flush real-time worker logs to UI Activity Log
        if ($localSyncState.LogQueue) {
            $logItem = $null
            while ($localSyncState.LogQueue.TryDequeue([ref]$logItem)) {
                if ($logItem) {
                    Write-BatchLog $logItem.Message $logItem.Level
                }
            }
        }

        $curr = $localSyncState.CurrentIndex
        $tot  = $localSyncState.TotalCount
        $pct  = if ($tot -gt 0) { [math]::Min(100, [math]::Round(($curr / $tot) * 100, 0)) } else { 0 }
        $pbBatchProgress.Value = $pct
        $lblBatchProgressText.Text = "Processing: $curr / $tot ($pct%)"
        $lblBatchStats.Text = "Success: $($localSyncState.SuccessCount) | Errors: $($localSyncState.FailCount)"

        if ($localBatchHandle -and $localBatchHandle.IsCompleted) {
            $script:ActiveBatchTimer.Stop()
            $btnStartBatch.IsEnabled = $true
            $btnStopBatch.IsEnabled = $false
            $btnBrowseBatchFile.IsEnabled = $true
            $btnReloadBatchFile.IsEnabled = $true
            $script:BatchWorkerRunning = $false

            # Flush any remaining logs
            if ($localSyncState.LogQueue) {
                $logItem = $null
                while ($localSyncState.LogQueue.TryDequeue([ref]$logItem)) {
                    if ($logItem) {
                        Write-BatchLog $logItem.Message $logItem.Level
                    }
                }
            }

            # Check stream errors
            foreach ($streamErr in $localBatchPs.Streams.Error) {
                Write-BatchLog "[Worker Stream Error] $($streamErr.Exception.Message)" "ERROR"
            }

            try {
                $res = $localBatchPs.EndInvoke($localBatchHandle)
                $script:BatchResultsList = [System.Collections.Generic.List[PSCustomObject]]::new()
                foreach ($item in @($res)) { $script:BatchResultsList.Add($item) }
                $dgBatchResults.ItemsSource = $script:BatchResultsList

                # Populate Points Detail table
                $allPointsList = [System.Collections.Generic.List[PSCustomObject]]::new()
                foreach ($item in @($script:BatchResultsList)) {
                    if ($item.Points -and @($item.Points).Count -gt 0) {
                        foreach ($pt in @($item.Points)) {
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
                if ($dgBatchPoints) {
                    $dgBatchPoints.ItemsSource = $allPointsList
                }

                $statusMsg = if ($localSyncState.CancelRequested) { 'Stopped by user.' } else { 'Completed successfully.' }
                $lblBatchProgressText.Text = $statusMsg
                Write-BatchLog "=== Batch processing completed. $statusMsg Success: $($localSyncState.SuccessCount), Errors: $($localSyncState.FailCount) ===" "OK"
                $lblFooterStatus.Text = "Processing complete: $($localSyncState.SuccessCount) routes generated."

                # Automatically switch to Calculation Results tab
                if ($script:BatchResultsList.Count -gt 0 -and $tabBatchSub) {
                    $tabBatchSub.SelectedIndex = 1
                }
            }
            catch {
                Write-BatchLog "Error reading batch results: $($_.Exception.Message)" "ERROR"
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
        $lblBatchProgressText.Text = 'Stopping...'
        Write-BatchLog "Stop requested by user..." "WARN"
    }
})

$dgBatchResults.Add_MouseDoubleClick({
    $sel = $dgBatchResults.SelectedItem
    if ($sel -and $sel.MapPath -and (Test-Path $sel.MapPath)) {
        Start-Process $sel.MapPath
    }
})

$btnOpenOutputDir.Add_Click({
    $outDir = $txtSettingsOutputDir.Text.Trim()
    if (Test-Path $outDir) {
        Start-Process explorer.exe -ArgumentList "`"$outDir`""
    }
})

$btnExportExcel.Add_Click({
    if ($script:BatchResultsList.Count -eq 0) {
        [System.Windows.MessageBox]::Show((Get-LocText 'MsgNoExportResults'), (Get-LocText 'MsgNoExportResultsTitle'), 'OK', 'Information')
        return
    }
    $outDir = $txtSettingsOutputDir.Text.Trim()
    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $path = Join-Path $outDir "${ts}_route_results.xlsx"
    $saved = Export-RouteResults -Results $script:BatchResultsList -OutputPath $path -Format Excel
    [System.Windows.MessageBox]::Show(((Get-LocText 'MsgExportExcelComplete') -f $saved), (Get-LocText 'MsgExportTitle'), 'OK', 'Information')
})

$btnExportCsv.Add_Click({
    if ($script:BatchResultsList.Count -eq 0) {
        [System.Windows.MessageBox]::Show((Get-LocText 'MsgNoExportResults'), (Get-LocText 'MsgNoExportResultsTitle'), 'OK', 'Information')
        return
    }
    $outDir = $txtSettingsOutputDir.Text.Trim()
    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $path = Join-Path $outDir "${ts}_route_results.csv"
    $saved = Export-RouteResults -Results $script:BatchResultsList -OutputPath $path -Format CSV
    [System.Windows.MessageBox]::Show(((Get-LocText 'MsgExportCsvComplete') -f $saved), (Get-LocText 'MsgExportTitle'), 'OK', 'Information')
})

$btnExportJson.Add_Click({
    if ($script:BatchResultsList.Count -eq 0) {
        [System.Windows.MessageBox]::Show((Get-LocText 'MsgNoExportResults'), (Get-LocText 'MsgNoExportResultsTitle'), 'OK', 'Information')
        return
    }
    $outDir = $txtSettingsOutputDir.Text.Trim()
    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $path = Join-Path $outDir "${ts}_route_results.json"
    $saved = Export-RouteResults -Results $script:BatchResultsList -OutputPath $path -Format JSON
    [System.Windows.MessageBox]::Show(((Get-LocText 'MsgExportJsonComplete') -f $saved), (Get-LocText 'MsgExportTitle'), 'OK', 'Information')
})

# ── 13. Obsługa zamykania okna ───────────────────────────────────────────────
$window.Add_Closing({
    if ($script:ActiveBatchTimer) { try { $script:ActiveBatchTimer.Stop() } catch { } }
    if ($script:ActiveManualTimer) { try { $script:ActiveManualTimer.Stop() } catch { } }
    if ($script:SyncState) { $script:SyncState.CancelRequested = $true }
})

# ── 14. Uruchomienie okna ────────────────────────────────────────────────────
$window.ShowDialog() | Out-Null
.Exception.Message)" "WARN"
        }
    }

    # 5. Fallback safeguard: Ensure fallback EN exists
    if (-not $script:LanguagesCatalog.Contains('en')) {
        $script:LanguagesCatalog['en'] = [PSCustomObject]@{
            Code        = 'en'
            DisplayName = 'English'
            GoogleCode  = 'en'
            Strings     = @{ 'AppTitle' = 'Google Maps Route & Map Generator' }
        }
    }
    $script:DefaultStrings = $script:LanguagesCatalog['en'].Strings
}

function Get-LocText {
    param(
        [Parameter(Mandatory)][string]$Key,
        [Parameter()][string]$Default = $null
    )
    if ($script:CurrentStrings -and $script:CurrentStrings.ContainsKey($Key) -and -not [string]::IsNullOrWhiteSpace($script:CurrentStrings[$Key])) {
        return $script:CurrentStrings[$Key]
    }
    if ($script:DefaultStrings -and $script:DefaultStrings.ContainsKey($Key) -and -not [string]::IsNullOrWhiteSpace($script:DefaultStrings[$Key])) {
        return $script:DefaultStrings[$Key]
    }
    if ($Default) { return $Default }
    return $Key
}

function Get-MaskedKey([string]$Key) {
    if ([string]::IsNullOrWhiteSpace($Key)) { return '(brak)' }
    if ($Key.Length -le 8) { return '***' }
    return "$($Key.Substring(0, 4))...$($Key.Substring($Key.Length - 4, 4))"
}

function Write-AppLog {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'OK', 'WARN', 'ERROR', 'DEBUG')]
        [string]$Level = 'INFO',
        [switch]$ToBatchWindow
    )
    $now = Get-Date
    $timeStr = $now.ToString('yyyy-MM-dd HH:mm:ss.fff')
    $prefix = switch ($Level) {
        'OK'    { '[OK]   ' }
        'WARN'  { '[WARN] ' }
        'ERROR' { '[ERROR]' }
        'DEBUG' { '[DEBUG]' }
        default { '[INFO] ' }
    }
    $entry = "[$timeStr] $prefix $Message"

    try {
        [System.IO.File]::AppendAllText($script:LogFile, "$entry`r`n", [System.Text.UTF8Encoding]::new($true))
    } catch { }

    if ($ToBatchWindow -and $txtBatchLog) {
        try {
            $batchTime = $now.ToString('HH:mm:ss')
            $line = "$batchTime $prefix $Message`r`n"
            $txtBatchLog.Dispatcher.Invoke([Action]{
                $txtBatchLog.AppendText($line)
                $txtBatchLog.ScrollToEnd()
            })
        } catch { }
    }
}

Write-AppLog "================================================================================" "INFO"
Write-AppLog "Uruchomienie Google Maps Route & Map Generator v2.0" "INFO"
Write-AppLog "Środowisko: PowerShell $($PSVersionTable.PSVersion), OS: $([System.Environment]::OSVersion.VersionString)" "INFO"
Write-AppLog "Plik konfiguracji: $script:ConfigFile" "INFO"
Write-AppLog "Plik dziennika zdarzeń (log): $script:LogFile" "INFO"

$script:OverlayPropKeys = @('StartGeocoded', 'EndGeocoded', 'Distance', 'Duration', 'Timestamp', 'RouteName', 'RouteType', 'Waypoints', 'StartRaw', 'EndRaw')

function Get-DefaultOverlayConfig {
    return [ordered]@{
        EnableTopOverlay    = $true
        EnableBottomOverlay = $true
        Items               = [ordered]@{
            StartGeocoded = [ordered]@{ Enabled = $true;  Panel = 'Bottom'; Align = 'Left';   Order = 1 }
            EndGeocoded   = [ordered]@{ Enabled = $true;  Panel = 'Bottom'; Align = 'Left';   Order = 2 }
            Distance      = [ordered]@{ Enabled = $true;  Panel = 'Bottom'; Align = 'Left';   Order = 3 }
            Duration      = [ordered]@{ Enabled = $true;  Panel = 'Bottom'; Align = 'Center'; Order = 3 }
            Timestamp     = [ordered]@{ Enabled = $true;  Panel = 'Bottom'; Align = 'Right';  Order = 3 }
            RouteName     = [ordered]@{ Enabled = $true;  Panel = 'Top';    Align = 'Left';   Order = 1 }
            RouteType     = [ordered]@{ Enabled = $true;  Panel = 'Top';    Align = 'Right';  Order = 1 }
            Waypoints     = [ordered]@{ Enabled = $false; Panel = 'Bottom'; Align = 'Left';   Order = 2 }
            StartRaw      = [ordered]@{ Enabled = $false; Panel = 'None';   Align = 'Left';   Order = 1 }
            EndRaw        = [ordered]@{ Enabled = $false; Panel = 'None';   Align = 'Left';   Order = 2 }
        }
    }
}

function Load-AppConfig {
    $defaultResults = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'GoogleMapsRoutes\Results'
    $defaultOverlay = Get-DefaultOverlayConfig

    $cfg = [PSCustomObject]@{
        ApiKey           = ''
        RememberApiKey   = $true
        LastOutputFolder = $defaultResults
        LastInputFolder  = ''
        LastInputPath    = ''
        DefaultRouteType = 'Fastest'
        DefaultEmission  = 'GASOLINE'
        MapWidth         = 900
        MapHeight        = 600
        Language         = 'en'
        OverlayConfig    = $defaultOverlay
        Theme            = 'Dark'
    }

    if (Test-Path $script:ConfigFile) {
        try {
            $jsonText = [System.IO.File]::ReadAllText($script:ConfigFile, [System.Text.Encoding]::UTF8)
            $raw = $jsonText | ConvertFrom-Json
            if ($raw.ApiKeyEncrypted -is [string] -and -not [string]::IsNullOrWhiteSpace($raw.ApiKeyEncrypted)) {
                $dec = Unprotect-SecretString -EncryptedText $raw.ApiKeyEncrypted
                if (-not [string]::IsNullOrWhiteSpace($dec)) { $cfg.ApiKey = $dec }
            }
            elseif ($raw.ApiKey -is [string] -and -not [string]::IsNullOrWhiteSpace($raw.ApiKey)) {
                $dec = Unprotect-SecretString -EncryptedText $raw.ApiKey
                if (-not [string]::IsNullOrWhiteSpace($dec)) { $cfg.ApiKey = $dec }
            }

            if ($null -ne $raw.RememberApiKey) { $cfg.RememberApiKey = [bool]$raw.RememberApiKey }
            if ($raw.LastOutputFolder -is [string] -and -not [string]::IsNullOrWhiteSpace($raw.LastOutputFolder)) {
                $cfg.LastOutputFolder = $raw.LastOutputFolder
            }
            if ($raw.LastInputFolder -is [string] -and -not [string]::IsNullOrWhiteSpace($raw.LastInputFolder)) {
                $cfg.LastInputFolder = $raw.LastInputFolder
            }
            if ($raw.LastInputPath -is [string] -and -not [string]::IsNullOrWhiteSpace($raw.LastInputPath)) {
                $cfg.LastInputPath = $raw.LastInputPath
            }
            if ($raw.DefaultRouteType -is [string]) { $cfg.DefaultRouteType = $raw.DefaultRouteType }
            if ($raw.DefaultEmission -is [string]) { $cfg.DefaultEmission = $raw.DefaultEmission }
            if ($raw.MapWidth) { $cfg.MapWidth = [int]$raw.MapWidth }
            if ($raw.MapHeight) { $cfg.MapHeight = [int]$raw.MapHeight }
            if ($raw.Language -is [string] -and -not [string]::IsNullOrWhiteSpace($raw.Language)) {
                $cfg.Language = $raw.Language.Trim().ToLower()
            }
            if ($raw.Theme -is [string] -and $raw.Theme -in @('Dark', 'Light')) {
                $cfg.Theme = $raw.Theme
            }

            if ($raw.OverlayConfig) {
                if ($null -ne $raw.OverlayConfig.EnableTopOverlay) {
                    $defaultOverlay.EnableTopOverlay = [bool]$raw.OverlayConfig.EnableTopOverlay
                }
                if ($null -ne $raw.OverlayConfig.EnableBottomOverlay) {
                    $defaultOverlay.EnableBottomOverlay = [bool]$raw.OverlayConfig.EnableBottomOverlay
                }
                if ($raw.OverlayConfig.Items) {
                    foreach ($k in $script:OverlayPropKeys) {
                        $rawItem = if ($raw.OverlayConfig.Items.PSObject.Properties[$k]) {
                            $raw.OverlayConfig.Items.$k
                        } elseif ($raw.OverlayConfig.Items[$k]) {
                            $raw.OverlayConfig.Items[$k]
                        } else { $null }

                        if ($rawItem) {
                            $en = if ($null -ne $rawItem.Enabled) { [bool]$rawItem.Enabled } else { $defaultOverlay.Items[$k].Enabled }
                            $pn = if ($rawItem.Panel) { [string]$rawItem.Panel } else { $defaultOverlay.Items[$k].Panel }
                            $al = if ($rawItem.Align) { [string]$rawItem.Align } else { $defaultOverlay.Items[$k].Align }
                            $od = if ($rawItem.Order) { [int]$rawItem.Order } else { $defaultOverlay.Items[$k].Order }
                            $defaultOverlay.Items[$k] = [ordered]@{ Enabled = $en; Panel = $pn; Align = $al; Order = $od }
                        }
                    }
                }
            }
            $cfg.OverlayConfig = $defaultOverlay
        } catch { }
    }

    if ([string]::IsNullOrWhiteSpace($cfg.ApiKey) -and -not [string]::IsNullOrWhiteSpace($env:GOOGLE_MAPS_API_KEY)) {
        $cfg.ApiKey = $env:GOOGLE_MAPS_API_KEY
    }

    return $cfg
}

function Save-AppConfig {
    param(
        [string]$ApiKey,
        [bool]$RememberApiKey,
        [string]$OutputFolder,
        [string]$LastInputFolder = '',
        [string]$LastInputPath = '',
        [string]$DefaultRouteType = 'Fastest',
        [string]$DefaultEmission = 'GASOLINE',
        [int]$MapWidth = 900,
        [int]$MapHeight = 600,
        [string]$Language = '',
        [object]$OverlayConfig = $null,
        [string]$Theme = ''
    )
    $encKey = ''
    if ($RememberApiKey -and -not [string]::IsNullOrWhiteSpace($ApiKey)) {
        try {
            $protected = Protect-SecretString -PlainText $ApiKey
            if ($protected -is [string] -and -not [string]::IsNullOrWhiteSpace($protected)) {
                $encKey = $protected
            }
        } catch {
            $encKey = ''
        }
    }

    $finalInputFolder = if ($LastInputFolder) { $LastInputFolder } elseif ($script:Config -and $script:Config.LastInputFolder) { $script:Config.LastInputFolder } else { '' }
    $finalInputPath   = if ($LastInputPath) { $LastInputPath } elseif ($script:Config -and $script:Config.LastInputPath) { $script:Config.LastInputPath } else { '' }
    $finalLang        = if ($Language) { $Language } elseif ($script:Config -and $script:Config.Language) { $script:Config.Language } else { 'en' }
    $finalOverlay     = if ($OverlayConfig) { $OverlayConfig } elseif ($script:Config -and $script:Config.OverlayConfig) { $script:Config.OverlayConfig } else { Get-DefaultOverlayConfig }
    $finalTheme       = if ($Theme -in @('Dark', 'Light')) { $Theme } elseif ($script:Config -and $script:Config.Theme) { $script:Config.Theme } else { 'Dark' }

    $cfg = [ordered]@{
        ApiKeyEncrypted  = $encKey
        RememberApiKey   = $RememberApiKey
        LastOutputFolder = $OutputFolder
        LastInputFolder  = $finalInputFolder
        LastInputPath    = $finalInputPath
        DefaultRouteType = $DefaultRouteType
        DefaultEmission  = $DefaultEmission
        MapWidth         = $MapWidth
        MapHeight        = $MapHeight
        Language         = $finalLang
        OverlayConfig    = $finalOverlay
        Theme            = $finalTheme
    }
    $json = $cfg | ConvertTo-Json -Depth 6
    [System.IO.File]::WriteAllText($script:ConfigFile, $json, [System.Text.UTF8Encoding]::new($true))
}

function Get-CurrentOverlayConfig {
    $topEn = if ($chkEnableTopOverlay) { [bool]$chkEnableTopOverlay.IsChecked } else { $true }
    $btmEn = if ($chkEnableBottomOverlay) { [bool]$chkEnableBottomOverlay.IsChecked } else { $true }
    $cfg = [ordered]@{
        EnableTopOverlay    = $topEn
        EnableBottomOverlay = $btmEn
        Items               = [ordered]@{}
    }
    foreach ($key in $script:OverlayPropKeys) {
        $chk  = Get-Variable -Name "chkProp_$key"  -ValueOnly -ErrorAction SilentlyContinue
        $cmbP = Get-Variable -Name "cmbPanel_$key" -ValueOnly -ErrorAction SilentlyContinue
        $cmbA = Get-Variable -Name "cmbAlign_$key" -ValueOnly -ErrorAction SilentlyContinue
        $cmbO = Get-Variable -Name "cmbOrder_$key" -ValueOnly -ErrorAction SilentlyContinue

        $enabled = if ($chk) { [bool]$chk.IsChecked } else { $true }
        $panel   = if ($cmbP -and $cmbP.SelectedItem) { [string]$cmbP.SelectedItem.Tag } else { 'Bottom' }
        $align   = if ($cmbA -and $cmbA.SelectedItem) { [string]$cmbA.SelectedItem.Tag } else { 'Left' }
        $order   = if ($cmbO -and $cmbO.SelectedItem) { [int]$cmbO.SelectedItem.Tag } else { 1 }

        $cfg.Items[$key] = [ordered]@{
            Enabled = $enabled
            Panel   = $panel
            Align   = $align
            Order   = $order
        }
    }
    return $cfg
}

function Set-OverlayConfigUi($cfg) {
    if (-not $cfg) { return }
    if ($null -ne $cfg.EnableTopOverlay -and $chkEnableTopOverlay) {
        $chkEnableTopOverlay.IsChecked = [bool]$cfg.EnableTopOverlay
    }
    if ($null -ne $cfg.EnableBottomOverlay -and $chkEnableBottomOverlay) {
        $chkEnableBottomOverlay.IsChecked = [bool]$cfg.EnableBottomOverlay
    }
    if ($cfg.Items) {
        foreach ($key in $script:OverlayPropKeys) {
            $itemCfg = if ($cfg.Items.PSObject.Properties[$key]) {
                $cfg.Items.$key
            } elseif ($cfg.Items[$key]) {
                $cfg.Items[$key]
            } else { $null }

            if (-not $itemCfg) { continue }

            $chk  = Get-Variable -Name "chkProp_$key"  -ValueOnly -ErrorAction SilentlyContinue
            $cmbP = Get-Variable -Name "cmbPanel_$key" -ValueOnly -ErrorAction SilentlyContinue
            $cmbA = Get-Variable -Name "cmbAlign_$key" -ValueOnly -ErrorAction SilentlyContinue
            $cmbO = Get-Variable -Name "cmbOrder_$key" -ValueOnly -ErrorAction SilentlyContinue

            if ($chk -and $null -ne $itemCfg.Enabled) {
                $chk.IsChecked = [bool]$itemCfg.Enabled
            }
            if ($cmbP -and $itemCfg.Panel) {
                foreach ($opt in $cmbP.Items) {
                    if ($opt.Tag -eq $itemCfg.Panel) { $cmbP.SelectedItem = $opt; break }
                }
            }
            if ($cmbA -and $itemCfg.Align) {
                foreach ($opt in $cmbA.Items) {
                    if ($opt.Tag -eq $itemCfg.Align) { $cmbA.SelectedItem = $opt; break }
                }
            }
            if ($cmbO -and $itemCfg.Order) {
                foreach ($opt in $cmbO.Items) {
                    if ([int]$opt.Tag -eq [int]$itemCfg.Order) { $cmbO.SelectedItem = $opt; break }
                }
            }
        }
    }
}

function Reset-OverlayConfigUi {
    $defaultCfg = Get-DefaultOverlayConfig
    Set-OverlayConfigUi $defaultCfg
}

$script:Config = Load-AppConfig
Load-LocalizationConfig
$script:CurrentLanguage = if ($script:Config -and $script:Config.Language -and $script:LanguagesCatalog.Contains($script:Config.Language.ToLower())) {
    $script:Config.Language.ToLower()
} else {
    'en'
}
$script:CurrentGoogleLang = if ($script:LanguagesCatalog.Contains($script:CurrentLanguage)) {
    $script:LanguagesCatalog[$script:CurrentLanguage].GoogleCode
} else {
    'en'
}
$script:CurrentStrings = if ($script:LanguagesCatalog.Contains($script:CurrentLanguage)) {
    $script:LanguagesCatalog[$script:CurrentLanguage].Strings
} else {
    @{}
}
Write-AppLog "Active language: $script:CurrentLanguage (Google API code: $script:CurrentGoogleLang)" "INFO" 

# ══════════════════════════════════════════════════════════════════════════════
# 6. FABRYKA BEZPIECZNYCH WĄTKÓW TŁA (INITIALSESSIONSTATE RUNSPACE)
# ══════════════════════════════════════════════════════════════════════════════

function New-WorkerPowerShell {
    param([scriptblock]$ScriptBlock)
    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    Get-ChildItem function: | Where-Object {
        $_.Name -in @('Protect-SecretString', 'Unprotect-SecretString', 'Test-GoogleApiKey',
                      'Get-AddressComponentValue', 'Get-AddressCoordinates', 'Get-GeocodeStatusDescription', 'Get-CarRouteData',
                      'Get-GoogleMapsUrl', 'Get-WrappedLines', 'Save-RouteMapPng',
                      'Find-MatchingPropertyName', 'Import-RouteDataFile', 'Export-RouteResults')
    } | ForEach-Object {
        try {
            $iss.Commands.Add([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new($_.Name, $_.Definition))
        } catch { }
    }
    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($iss)
    $rs.ApartmentState = [System.Threading.ApartmentState]::MTA
    $rs.ThreadOptions  = [System.Management.Automation.Runspaces.PSThreadOptions]::UseNewThread
    $rs.Open()
    $ps = [PowerShell]::Create()
    $ps.Runspace = $rs
    # WAŻNE: [void] lub $null = zapobiega wyciekowi obiektu PowerShell do pipeline funkcji.
    # Bez tego funkcja zwraca tablicę @($ps, $ps), co przy wywołaniu .BeginInvoke()
    # powoduje próbę ponownego uruchomienia tej samej instancji i błąd:
    # "The operation cannot be performed because a command has already been started."
    $null = $ps.AddScript($ScriptBlock.ToString())
    return $ps
}

# ══════════════════════════════════════════════════════════════════════════════
# ══════════════════════════════════════════════════════════════════════════════
# 6b. MANUAL CALC WORKER SCRIPTBLOCK (Isolated Runspace, Top-level)
# ══════════════════════════════════════════════════════════════════════════════
$script:ManualCalcAsync = {
    param($start, $end, $waypoints, $routeType, $emission, $trafficAware, $name, $apiKey, $outDir, $logFile, $languageCode = 'en', $overlayConfigJson = '')

    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls
    Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

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
        if ($geoStart.Status -ne 'OK') {
            & $wlog "Origin geocoding error: $($geoStart.Status)" "WARN"
            return [PSCustomObject]@{ Success = $false; Error = "Origin geocoding error: $($geoStart.Status)" }
        }
        & $wlog "Origin OK: $($geoStart.FormattedAddress) ($($geoStart.Latitude), $($geoStart.Longitude))" "INFO"

        & $wlog "Geocoding destination: '$end'..." "INFO"
        $geoEnd = Get-AddressCoordinates -Address $end -ApiKey $apiKey -LanguageCode $languageCode
        if ($geoEnd.Status -ne 'OK') {
            & $wlog "Destination geocoding error: $($geoEnd.Status)" "WARN"
            return [PSCustomObject]@{ Success = $false; Error = "Destination geocoding error: $($geoEnd.Status)" }
        }
        & $wlog "Destination OK: $($geoEnd.FormattedAddress) ($($geoEnd.Latitude), $($geoEnd.Longitude))" "INFO"

        $geoWp = [System.Collections.Generic.List[PSCustomObject]]::new()
        foreach ($w in $waypoints) {
            & $wlog "Geocoding waypoint: '$w'..." "INFO"
            $g = Get-AddressCoordinates -Address $w -ApiKey $apiKey -LanguageCode $languageCode
            if ($g.Status -eq 'OK') {
                $geoWp.Add($g)
                & $wlog "Waypoint OK: $($g.FormattedAddress)" "INFO"
            } else {
                & $wlog "Waypoint geocoding error '$w': $($g.Status)" "WARN"
            }
        }

        & $wlog "Querying Google Routes API v2 (Type: $routeType, Engine: $emission)..." "INFO"
        $trasa = Get-CarRouteData -OriginLat $geoStart.Latitude -OriginLng $geoStart.Longitude `
            -DestLat $geoEnd.Latitude -DestLng $geoEnd.Longitude `
            -IntermediatePoints $geoWp -RouteType $routeType -EmissionType $emission `
            -ApiKey $apiKey -LanguageCode $languageCode -TrafficAware:$trafficAware

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

        & $wlog "Map rendering complete. Saved: $saved" "INFO"

        $resolvedMapPath = $(if ($saved) { $mapPath } else { $null })
        return [PSCustomObject]@{
            Success       = $true
            DistanceKm    = $trasa.OdlegloscKm
            DurationMin   = $trasa.CzasMin
            RouteType     = $routeType
            GoogleMapsUrl = $gUrl
            MapPath       = $resolvedMapPath
            Error         = $null
        }
    }
    catch {
        $errFull = $_.Exception.ToString()
        & $wlog "Worker thread exception: $errFull" "ERROR"
        return [PSCustomObject]@{ Success = $false; Error = $_.Exception.Message }
    }
}

# ══════════════════════════════════════════════════════════════════════════════
# 6c. BATCH CALC WORKER SCRIPTBLOCK (Isolated Runspace, Top-level)
# ══════════════════════════════════════════════════════════════════════════════
$script:BatchCalcAsync = {
    param($routes, $apiKey, $outDir, $defaultRouteType, $syncState, $logFile, $languageCode = 'en', $overlayConfigJson = '')

    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls
    Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

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

        $routeName = if ($r.Name) { $r.Name } else { "Route $($i + 1)" }

        & $wlog "Route $($i + 1)/$($total): Processing '$($r.Start)' -> '$($r.End)' (Type: $rType)..." "INFO"

        try {
            $geoStart = Get-AddressCoordinates -Address $r.Start -ApiKey $apiKey -LanguageCode $languageCode
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

            $geoEnd   = Get-AddressCoordinates -Address $r.End -ApiKey $apiKey -LanguageCode $languageCode
            $endStatus = Get-GeocodeStatusDescription -Geo $geoEnd
            $isEndFallback = if ($geoEnd -and ($geoEnd.PartialMatch -or $geoEnd.MatchType -in 'APPROXIMATE', 'GEOMETRIC_CENTER')) { $true } else { $false }

            if ($geoStart.Status -ne 'OK' -or $geoEnd.Status -ne 'OK') {
                $errReason = "Geocoding failed (Start=$($geoStart.Status), End=$($geoEnd.Status))"
                & $wlog "Route $($i + 1)/$($total): $errReason" "WARN"

                $routePointsList.Add([PSCustomObject]@{
                    Order           = 2
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

                $results.Add([PSCustomObject]@{
                    Id                = [string]($i + 1)
                    Name              = $routeName
                    Nazwa             = $routeName
                    Start             = $r.Start
                    StartGeocoded     = if ($geoStart) { $geoStart.FormattedAddress } else { $null }
                    StartStatus       = $startStatus
                    End               = $r.End
                    EndGeocoded       = if ($geoEnd) { $geoEnd.FormattedAddress } else { $null }
                    EndStatus         = $endStatus
                    Koniec            = $r.End
                    WaypointsCount    = 0
                    LiczbaPrzystankow = 0
                    RouteType         = $rType
                    TypTrasy          = $rType
                    DistanceKm        = $null
                    OdlegloscKm       = $null
                    DurationMin       = $null
                    CzasMin           = $null
                    Status            = $errReason
                    MapPath           = $null
                    MapaPath          = $null
                    Points            = @($routePointsList)
                })
                $syncState.FailCount++
                continue
            }

            $geoWp = [System.Collections.Generic.List[PSCustomObject]]::new()
            $wpIdx = 1
            if ($r.Waypoints) {
                foreach ($w in $r.Waypoints) {
                    if ([string]::IsNullOrWhiteSpace($w)) { continue }
                    $g = Get-AddressCoordinates -Address $w -ApiKey $apiKey -LanguageCode $languageCode
                    $wpStatus = Get-GeocodeStatusDescription -Geo $g
                    $isWpFallback = if ($g -and ($g.PartialMatch -or $g.MatchType -in 'APPROXIMATE', 'GEOMETRIC_CENTER')) { $true } else { $false }

                    $routePointsList.Add([PSCustomObject]@{
                        Order           = ($wpIdx + 1)
                        PointType       = "Waypoint $wpIdx"
                        OriginalAddress = $w
                        GeocodedAddress = if ($g) { $g.FormattedAddress } else { $null }
                        GeocodeStatus   = $wpStatus
                        MatchType       = if ($g) { $g.MatchType } else { 'NOT_FOUND' }
                        PartialMatch    = if ($g) { [bool]$g.PartialMatch } else { $false }
                        IsFallback      = $isWpFallback
                        Latitude        = if ($g) { $g.Latitude } else { $null }
                        Longitude       = if ($g) { $g.Longitude } else { $null }
                    })

                    if ($g -and $g.Status -eq 'OK' -and $null -ne $g.Latitude -and $null -ne $g.Longitude) {
                        $geoWp.Add($g)
                    } else {
                        & $wlog "Route $($i + 1)/$($total): Waypoint '$w' cannot be located ($wpStatus). Proceeding without it in driving directions." "WARN"
                    }
                    $wpIdx++
                    Start-Sleep -Milliseconds 60
                }
            }

            # Add End point to structured points
            $routePointsList.Add([PSCustomObject]@{
                Order           = ($routePointsList.Count + 1)
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

            & $wlog "Route $($i + 1)/$($total): Querying Google Routes API..." "INFO"
            $trasa = Get-CarRouteData -OriginLat $geoStart.Latitude -OriginLng $geoStart.Longitude `
                -DestLat $geoEnd.Latitude -DestLng $geoEnd.Longitude `
                -IntermediatePoints $geoWp -RouteType $rType -ApiKey $apiKey `
                -LanguageCode $languageCode

            if ($trasa.Status -ne 'OK') {
                $errReason = "Routes API: $($trasa.Status). $($trasa.ErrorMessage)"
                & $wlog "Route $($i + 1)/$($total): $errReason" "WARN"
                $results.Add([PSCustomObject]@{
                    Id                = [string]($i + 1)
                    Name              = $routeName
                    Nazwa             = $routeName
                    Start             = $r.Start
                    StartGeocoded     = if ($geoStart) { $geoStart.FormattedAddress } else { $null }
                    StartStatus       = $startStatus
                    End               = $r.End
                    EndGeocoded       = if ($geoEnd) { $geoEnd.FormattedAddress } else { $null }
                    EndStatus         = $endStatus
                    Koniec            = $r.End
                    WaypointsCount    = $geoWp.Count
                    LiczbaPrzystankow = $geoWp.Count
                    RouteType         = $rType
                    TypTrasy          = $rType
                    DistanceKm        = $null
                    OdlegloscKm       = $null
                    DurationMin       = $null
                    CzasMin           = $null
                    Status            = $errReason
                    MapPath           = $null
                    MapaPath          = $null
                    Points            = @($routePointsList)
                })
                $syncState.FailCount++
                continue
            }

            $safeName = ($routeName -replace '[\\/:*?"<>|]', '_').Trim()
            $mapPath = Join-Path $outDir "${ts}_route_$($i + 1)_${safeName}.png"

            $allPts = [System.Collections.Generic.List[PSCustomObject]]::new()
            $allPts.Add($geoStart)
            foreach ($wp in $geoWp) { $allPts.Add($wp) }
            $allPts.Add($geoEnd)

            & $wlog "Route $($i + 1)/$($total): Route OK ($($trasa.OdlegloscKm) km, $($trasa.CzasMin) min). Rendering static map..." "INFO"

            $hdrBatchPrefix = switch ($languageCode) { 'de' { 'Typ: ' } 'pl' { 'Typ: ' } default { 'Type: ' } }
            $hdrBatchName = switch ($languageCode) {
                'de' { if ($rType -eq 'Fastest') { 'Schnellste' } elseif ($rType -eq 'Shortest') { 'Kürzeste' } else { 'Eco' } }
                'pl' { if ($rType -eq 'Fastest') { 'Najszybsza' } elseif ($rType -eq 'Shortest') { 'Najkrótsza' } else { 'Eko' } }
                default { $rType }
            }
            $hdrBatchRightText = "$hdrBatchPrefix$hdrBatchName"

            $saved = Save-RouteMapPng -EncodedPolyline $trasa.EncodedPolyline `
                -OriginLat $geoStart.Latitude -OriginLng $geoStart.Longitude `
                -DestLat $geoEnd.Latitude -DestLng $geoEnd.Longitude `
                -RoutePoints $allPts -OutputPath $mapPath -ApiKey $apiKey `
                -AddressTextA $geoStart.FormattedAddress -AddressTextB $geoEnd.FormattedAddress `
                -DistanceText "$($trasa.OdlegloscKm) km" -DurationText "$($trasa.CzasMin) min" `
                -HeaderLeftText $routeName -HeaderRightText $hdrBatchRightText `
                -LanguageCode $languageCode `
                -StartRaw $r.Start -StartGeocoded $geoStart.FormattedAddress `
                -EndRaw $r.End -EndGeocoded $geoEnd.FormattedAddress `
                -WaypointsList $geoWp -RouteName $routeName -RouteType $hdrBatchRightText `
                -OverlayConfig $overlayConfigJson

            $resolvedMapPath = if ($saved -and (Test-Path $mapPath)) { $mapPath } else { $null }

            $results.Add([PSCustomObject]@{
                Id                = [string]($i + 1)
                Name              = $routeName
                Nazwa             = $routeName
                Start             = $r.Start
                StartGeocoded     = if ($geoStart) { $geoStart.FormattedAddress } else { $null }
                StartStatus       = $startStatus
                End               = $r.End
                EndGeocoded       = if ($geoEnd) { $geoEnd.FormattedAddress } else { $null }
                EndStatus         = $endStatus
                Koniec            = $r.End
                WaypointsCount    = $geoWp.Count
                LiczbaPrzystankow = $geoWp.Count
                RouteType         = $rType
                TypTrasy          = $rType
                DistanceKm        = $trasa.OdlegloscKm
                OdlegloscKm       = $trasa.OdlegloscKm
                DurationMin       = $trasa.CzasMin
                CzasMin           = $trasa.CzasMin
                Status            = 'OK'
                MapPath           = $resolvedMapPath
                MapaPath          = $resolvedMapPath
                Points            = @($routePointsList)
            })
            $syncState.SuccessCount++
            & $wlog "Route $($i + 1)/$($total): Complete! Map saved: $(Split-Path $mapPath -Leaf)" "OK"
        }
        catch {
            $errDetail = $_.Exception.Message
            & $wlog "Route $($i + 1)/$($total): Exception: $errDetail" "ERROR"
            $results.Add([PSCustomObject]@{
                Id                = [string]($i + 1)
                Name              = $routeName
                Nazwa             = $routeName
                Start             = $r.Start
                StartGeocoded     = if ($geoStart) { $geoStart.FormattedAddress } else { $null }
                StartStatus       = if ($startStatus) { $startStatus } else { 'EXCEPTION' }
                End               = $r.End
                EndGeocoded       = if ($geoEnd) { $geoEnd.FormattedAddress } else { $null }
                EndStatus         = if ($endStatus) { $endStatus } else { 'EXCEPTION' }
                Koniec            = $r.End
                WaypointsCount    = 0
                LiczbaPrzystankow = 0
                RouteType         = $rType
                TypTrasy          = $rType
                DistanceKm        = $null
                OdlegloscKm       = $null
                DurationMin       = $null
                CzasMin           = $null
                Status            = "Exception: $errDetail"
                MapPath           = $null
                MapaPath          = $null
                Points            = if ($routePointsList) { @($routePointsList) } else { @() }
            })
            $syncState.FailCount++
        }

        Start-Sleep -Milliseconds 100
    }

    return $results
}

# ══════════════════════════════════════════════════════════════════════════════
# 7. DEFINICJA INTERFEJSU WPF XAML (MODERN DARK THEME)
# ══════════════════════════════════════════════════════════════════════════════

[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Google Maps Route &amp; Map Generator"
    Height="880" Width="1100"
    MinHeight="700" MinWidth="900"
    WindowStartupLocation="CenterScreen"
    Background="{DynamicResource BgDark}" Foreground="{DynamicResource TextPrimary}"
    FontFamily="Segoe UI Variable, Segoe UI, sans-serif">

    <Window.Resources>
        <SolidColorBrush x:Key="BgDark" Color="#0F172A"/>
        <SolidColorBrush x:Key="BgCard" Color="#1E293B"/>
        <SolidColorBrush x:Key="BgCardHover" Color="#293548"/>
        <SolidColorBrush x:Key="BgCardAlt" Color="#162032"/>
        <SolidColorBrush x:Key="BorderCard" Color="#334155"/>
        <SolidColorBrush x:Key="TextPrimary" Color="#F8FAFC"/>
        <SolidColorBrush x:Key="TextSecondary" Color="#94A3B8"/>
        <SolidColorBrush x:Key="AccentBlue" Color="#2563EB"/>
        <SolidColorBrush x:Key="AccentGreen" Color="#10B981"/>
        <SolidColorBrush x:Key="AccentAmber" Color="#F59E0B"/>
        <SolidColorBrush x:Key="AccentRed" Color="#EF4444"/>
        <SolidColorBrush x:Key="BgInput" Color="#1E293B"/>
        <SolidColorBrush x:Key="BorderInput" Color="#334155"/>
        <SolidColorBrush x:Key="BtnSecondaryBg" Color="#334155"/>
        <SolidColorBrush x:Key="BtnSecondaryFg" Color="#F8FAFC"/>
        <SolidColorBrush x:Key="GridLines" Color="#2D3748"/>
        <SolidColorBrush x:Key="LogBg" Color="#0A0F1D"/>
        <SolidColorBrush x:Key="LogFg" Color="#38BDF8"/>
        <SolidColorBrush x:Key="DataGridHeaderBg" Color="#0F172A"/>
        <SolidColorBrush x:Key="DataGridHeaderFg" Color="#94A3B8"/>
        <SolidColorBrush x:Key="DataGridRowBg" Color="#1E293B"/>
        <SolidColorBrush x:Key="DataGridAltRowBg" Color="#162032"/>

        <Style TargetType="TextBox">
            <Setter Property="Background" Value="{DynamicResource BgInput}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderInput}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="9,6"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>

        <Style TargetType="PasswordBox">
            <Setter Property="Background" Value="{DynamicResource BgInput}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderInput}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="9,6"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>

        <Style TargetType="Button">
            <Setter Property="Background" Value="#2563EB"/>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="14,8"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="5"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Modern Dynamic ComboBox with Dropdown Popup Template -->
        <ControlTemplate x:Key="ComboBoxToggleButtonTemplate" TargetType="ToggleButton">
            <Border x:Name="TemplateRoot" Background="{TemplateBinding Background}" BorderBrush="{DynamicResource BorderInput}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="5">
                <Border x:Name="SplitBorder" Width="26" HorizontalAlignment="Right" Background="Transparent">
                    <Path x:Name="Arrow" HorizontalAlignment="Center" VerticalAlignment="Center" Fill="{DynamicResource TextSecondary}" Data="M 0 0 L 4 4 L 8 0 Z"/>
                </Border>
            </Border>
            <ControlTemplate.Triggers>
                <Trigger Property="IsMouseOver" Value="true">
                    <Setter TargetName="TemplateRoot" Property="BorderBrush" Value="{DynamicResource AccentBlue}"/>
                    <Setter TargetName="Arrow" Property="Fill" Value="{DynamicResource TextPrimary}"/>
                </Trigger>
                <Trigger Property="IsChecked" Value="true">
                    <Setter TargetName="TemplateRoot" Property="BorderBrush" Value="{DynamicResource AccentBlue}"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="false">
                    <Setter TargetName="TemplateRoot" Property="Opacity" Value="0.5"/>
                </Trigger>
            </ControlTemplate.Triggers>
        </ControlTemplate>

        <Style TargetType="ComboBox">
            <Setter Property="Background" Value="{DynamicResource BgInput}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderInput}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="ScrollViewer.HorizontalScrollBarVisibility" Value="Auto"/>
            <Setter Property="ScrollViewer.VerticalScrollBarVisibility" Value="Auto"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBox">
                        <Grid x:Name="MainGrid" SnapsToDevicePixels="true">
                            <ToggleButton x:Name="ToggleButton"
                                          Template="{StaticResource ComboBoxToggleButtonTemplate}"
                                          Background="{TemplateBinding Background}"
                                          BorderBrush="{TemplateBinding BorderBrush}"
                                          BorderThickness="{TemplateBinding BorderThickness}"
                                          Focusable="false"
                                          IsChecked="{Binding Path=IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"
                                          ClickMode="Press"/>
                            <ContentPresenter x:Name="ContentSite"
                                              IsHitTestVisible="false"
                                              Content="{TemplateBinding SelectionBoxItem}"
                                              ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                                              ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}"
                                              Margin="{TemplateBinding Padding}"
                                              VerticalAlignment="Center"
                                              HorizontalAlignment="Left"/>
                            <Popup x:Name="Popup"
                                   Placement="Bottom"
                                   IsOpen="{TemplateBinding IsDropDownOpen}"
                                   AllowsTransparency="true"
                                   Focusable="false"
                                   PopupAnimation="Slide">
                                <Grid x:Name="DropDown" SnapsToDevicePixels="true" MinWidth="{TemplateBinding ActualWidth}" MaxHeight="{TemplateBinding MaxDropDownHeight}">
                                    <Border x:Name="DropDownBorder" Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="5" Margin="0,2,0,0">
                                        <ScrollViewer Margin="2" SnapsToDevicePixels="true">
                                            <StackPanel IsItemsHost="true" KeyboardNavigation.DirectionalNavigation="Contained"/>
                                        </ScrollViewer>
                                    </Border>
                                </Grid>
                            </Popup>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="HasItems" Value="false">
                                <Setter TargetName="DropDownBorder" Property="MinHeight" Value="40"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="false">
                                <Setter Property="Opacity" Value="0.6"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="ComboBoxItem">
            <Setter Property="Background" Value="{DynamicResource BgCard}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="Padding" Value="10,7"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBoxItem">
                        <Border x:Name="ItemBorder" Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}" SnapsToDevicePixels="true">
                            <ContentPresenter Content="{TemplateBinding Content}" ContentTemplate="{TemplateBinding ContentTemplate}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsHighlighted" Value="true">
                                <Setter TargetName="ItemBorder" Property="Background" Value="{DynamicResource AccentBlue}"/>
                                <Setter Property="Foreground" Value="#FFFFFF"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="true">
                                <Setter TargetName="ItemBorder" Property="Background" Value="{DynamicResource AccentBlue}"/>
                                <Setter Property="Foreground" Value="#FFFFFF"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="false">
                                <Setter Property="Foreground" Value="{DynamicResource TextSecondary}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Modern Dynamic ListBox & Items -->
        <Style TargetType="ListBox">
            <Setter Property="Background" Value="{DynamicResource BgDark}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderCard}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="ScrollViewer.HorizontalScrollBarVisibility" Value="Disabled"/>
        </Style>

        <Style TargetType="ListBoxItem">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="Padding" Value="8,5"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ListBoxItem">
                        <Border x:Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" Padding="{TemplateBinding Padding}" CornerRadius="3" SnapsToDevicePixels="true">
                            <ContentPresenter Content="{TemplateBinding Content}" ContentTemplate="{TemplateBinding ContentTemplate}" HorizontalAlignment="Left" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="true">
                                <Setter TargetName="Bd" Property="Background" Value="{DynamicResource AccentBlue}"/>
                                <Setter Property="Foreground" Value="#FFFFFF"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="true">
                                <Setter TargetName="Bd" Property="Background" Value="{DynamicResource BgCardHover}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Modern Dynamic DataGrid & Elements -->
        <Style TargetType="DataGrid">
            <Setter Property="Background" Value="{DynamicResource BgDark}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderCard}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="RowBackground" Value="{DynamicResource DataGridRowBg}"/>
            <Setter Property="AlternatingRowBackground" Value="{DynamicResource DataGridAltRowBg}"/>
            <Setter Property="GridLinesVisibility" Value="Horizontal"/>
            <Setter Property="HorizontalGridLinesBrush" Value="{DynamicResource GridLines}"/>
            <Setter Property="HeadersVisibility" Value="Column"/>
            <Setter Property="AutoGenerateColumns" Value="False"/>
            <Setter Property="IsReadOnly" Value="True"/>
            <Setter Property="CanUserAddRows" Value="False"/>
            <Setter Property="CanUserDeleteRows" Value="False"/>
            <Setter Property="SelectionMode" Value="Single"/>
            <Setter Property="SelectionUnit" Value="FullRow"/>
        </Style>

        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="{DynamicResource DataGridHeaderBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource DataGridHeaderFg}"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="10,7"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderCard}"/>
            <Setter Property="BorderThickness" Value="0,0,1,1"/>
        </Style>

        <Style TargetType="DataGridRow">
            <Setter Property="Background" Value="{DynamicResource DataGridRowBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="SnapsToDevicePixels" Value="true"/>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="true">
                    <Setter Property="Background" Value="{DynamicResource AccentBlue}"/>
                    <Setter Property="Foreground" Value="#FFFFFF"/>
                </Trigger>
                <Trigger Property="IsMouseOver" Value="true">
                    <Setter Property="Background" Value="{DynamicResource BgCardHover}"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style TargetType="DataGridCell">
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="DataGridCell">
                        <Border Background="{TemplateBinding Background}" BorderThickness="0" Padding="{TemplateBinding Padding}" SnapsToDevicePixels="true">
                            <ContentPresenter SnapsToDevicePixels="{TemplateBinding SnapsToDevicePixels}" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="true">
                    <Setter Property="Background" Value="{DynamicResource AccentBlue}"/>
                    <Setter Property="Foreground" Value="#FFFFFF"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style TargetType="TabItem">
            <Setter Property="Background" Value="{DynamicResource BgCard}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextSecondary}"/>
            <Setter Property="Padding" Value="18,10"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Border Name="TabBorder"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{DynamicResource BorderCard}"
                                BorderThickness="1,1,1,0"
                                CornerRadius="6,6,0,0"
                                Margin="0,0,4,0"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter ContentSource="Header" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="TabBorder" Property="Background" Value="{DynamicResource AccentBlue}"/>
                                <Setter Property="Foreground" Value="#FFFFFF"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <Border Grid.Row="0" Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="14,10" Margin="0,0,0,12">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Orientation="Vertical">
                    <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                        <TextBlock Text="🗺️" FontSize="20" Margin="0,0,8,0" VerticalAlignment="Center"/>
                        <TextBlock Name="txtHeaderTitle" Text="Google Maps Route &amp; Map Generator" FontSize="18" FontWeight="Bold" Foreground="{DynamicResource TextPrimary}"/>
                    </StackPanel>
                    <TextBlock Name="txtHeaderSubtitle" Text="Multi-point driving routes: Fastest, Shortest, Eco-friendly | Import JSON, CSV, Excel" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="28,2,0,0"/>
                </StackPanel>
                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Name="lblApiBadge" Text="API: Checking..." Foreground="#EF4444" FontSize="12" FontWeight="SemiBold" VerticalAlignment="Center" Margin="0,0,10,0"/>
                    <ComboBox Name="cmbAppLanguage" Width="135" Height="30" Margin="0,0,10,0" VerticalAlignment="Center" ToolTip="Select Language / Sprache wählen / Wybierz język"/>
                    <Button Name="btnQuickSettings" Content="⚙ API Settings" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Padding="10,5" FontSize="12" Margin="0,0,10,0"/>
                    <Button Name="btnThemeToggle" Content="🌙 Dark" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" Padding="10,5" FontSize="12" ToolTip="Toggle Light / Dark theme"/>
                </StackPanel>
            </Grid>
        </Border>


        <!-- Main TabControl -->
        <TabControl Name="tabMain" Grid.Row="1" Background="Transparent" BorderThickness="0">

            <!-- TAB 1: MANUAL ROUTE -->
            <TabItem Name="tabItemManual" Header="📍 Manual Route">
                <Grid Margin="0,10,0,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="420" MinWidth="360"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <ScrollViewer VerticalScrollBarVisibility="Auto" Grid.Column="0" Margin="0,0,10,0">
                        <StackPanel>
                            <Border Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="14" Margin="0,0,0,12">
                                <StackPanel>
                                    <TextBlock Name="lblManualRoutePointsHeader" Text="Route Points" FontSize="15" FontWeight="Bold" Foreground="{DynamicResource TextPrimary}" Margin="0,0,0,10"/>

                                    <TextBlock Name="lblManualOrigin" Text="Origin (Start / A):" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,4"/>
                                    <Grid Margin="0,0,0,10">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="Auto"/>
                                        </Grid.ColumnDefinitions>
                                        <TextBox Name="txtManualStart" Text="Warszawa, Plac Defilad 1"/>
                                        <Button Name="btnClearManualStart" Grid.Column="1" Content="✕" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Padding="8,6" Margin="4,0,0,0" ToolTip="Clear"/>
                                    </Grid>

                                    <TextBlock Name="lblManualWaypoints" Text="Intermediate Stops (optional up to 25):" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,4"/>
                                    <Grid Margin="0,0,0,6">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="Auto"/>
                                        </Grid.ColumnDefinitions>
                                        <TextBox Name="txtNewWaypoint" ToolTip="Enter waypoint address and click Add"/>
                                        <Button Name="btnAddWaypoint" Grid.Column="1" Content="➕ Add" Background="#10B981" Margin="4,0,0,0"/>
                                    </Grid>

                                    <ListBox Name="lstWaypoints" Height="110" Margin="0,0,0,6"/>
                                    <Grid Margin="0,0,0,10">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="*"/>
                                        </Grid.ColumnDefinitions>
                                        <Button Name="btnWpUp" Content="▲ Up" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="0,0,2,0" Padding="4,4" FontSize="11"/>
                                        <Button Name="btnWpDown" Grid.Column="1" Content="▼ Down" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="2,0,2,0" Padding="4,4" FontSize="11"/>
                                        <Button Name="btnWpRemove" Grid.Column="2" Content="✕ Remove" Background="#EF4444" Margin="2,0,2,0" Padding="4,4" FontSize="11"/>
                                        <Button Name="btnWpClear" Grid.Column="3" Content="🗑 Clear" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="2,0,0,0" Padding="4,4" FontSize="11"/>
                                    </Grid>

                                    <TextBlock Name="lblManualDestination" Text="Destination (End / B):" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,4"/>
                                    <Grid Margin="0,0,0,6">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="Auto"/>
                                        </Grid.ColumnDefinitions>
                                        <TextBox Name="txtManualEnd" Text="Kraków, Rynek Główny 1"/>
                                        <Button Name="btnClearManualEnd" Grid.Column="1" Content="✕" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Padding="8,6" Margin="4,0,0,0" ToolTip="Clear"/>
                                    </Grid>

                                    <TextBlock Name="lblManualRouteName" Text="Route Name / Description:" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,4,0,4"/>
                                    <TextBox Name="txtManualName" Text="Route Warsaw - Krakow" Margin="0,0,0,6"/>
                                </StackPanel>
                            </Border>

                            <Border Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="14" Margin="0,0,0,12">
                                <StackPanel>
                                    <TextBlock Name="lblManualOptHeader" Text="Route Optimization" FontSize="15" FontWeight="Bold" Foreground="{DynamicResource TextPrimary}" Margin="0,0,0,10"/>

                                    <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                                        <RadioButton Name="rbTypeFastest" Content="⚡ Fastest" IsChecked="True" Foreground="{DynamicResource TextPrimary}" FontSize="13" Margin="0,0,16,0"/>
                                        <RadioButton Name="rbTypeShortest" Content="📏 Shortest" Foreground="{DynamicResource TextPrimary}" FontSize="13" Margin="0,0,16,0"/>
                                        <RadioButton Name="rbTypeEco" Content="🌿 Eco" Foreground="{DynamicResource TextPrimary}" FontSize="13"/>
                                    </StackPanel>

                                    <StackPanel Name="pnlEmission" Orientation="Vertical" Visibility="Collapsed" Margin="0,0,0,8">
                                        <TextBlock Name="lblManualEmission" Text="Vehicle Engine Type (for Eco route):" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,4"/>
                                        <ComboBox Name="cmbEmission">
                                            <ComboBoxItem Content="Gasoline (Benzyna)" Tag="GASOLINE" IsSelected="True"/>
                                            <ComboBoxItem Content="Diesel" Tag="DIESEL"/>
                                            <ComboBoxItem Content="Hybrid" Tag="HYBRID"/>
                                            <ComboBoxItem Content="Electric" Tag="ELECTRIC"/>
                                        </ComboBox>
                                    </StackPanel>

                                    <CheckBox Name="chkTrafficAware" Content="Real-time traffic awareness (Live Traffic)" Foreground="{DynamicResource TextSecondary}" FontSize="12" Margin="0,4,0,4"/>
                                </StackPanel>
                            </Border>

                            <Button Name="btnCalculateManual" Content="🚀 CALCULATE ROUTE &amp; DOWNLOAD MAP" Background="#2563EB" Foreground="#FFFFFF" Padding="16,12" FontSize="14" FontWeight="Bold"/>
                        </StackPanel>
                    </ScrollViewer>

                    <Grid Grid.Column="1" Margin="10,0,0,0">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>

                        <Border Grid.Row="0" Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="14" Margin="0,0,0,10">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>

                                <StackPanel Grid.Column="0">
                                    <TextBlock Name="lblHeaderDist" Text="DISTANCE" FontSize="11" FontWeight="Bold" Foreground="{DynamicResource TextSecondary}"/>
                                    <TextBlock Name="lblManualDist" Text="— km" FontSize="20" FontWeight="Bold" Foreground="#10B981"/>
                                </StackPanel>

                                <StackPanel Grid.Column="1">
                                    <TextBlock Name="lblHeaderDur" Text="DURATION" FontSize="11" FontWeight="Bold" Foreground="{DynamicResource TextSecondary}"/>
                                    <TextBlock Name="lblManualTime" Text="— min" FontSize="20" FontWeight="Bold" Foreground="#F59E0B"/>
                                </StackPanel>

                                <StackPanel Grid.Column="2">
                                    <TextBlock Name="lblHeaderType" Text="ROUTE TYPE" FontSize="11" FontWeight="Bold" Foreground="{DynamicResource TextSecondary}"/>
                                    <TextBlock Name="lblManualType" Text="Fastest" FontSize="16" FontWeight="SemiBold" Foreground="#38BDF8"/>
                                </StackPanel>

                                <StackPanel Grid.Column="3" VerticalAlignment="Center">
                                    <TextBlock Name="lblManualStatus" Text="Idle" FontSize="12" Foreground="{DynamicResource TextSecondary}" HorizontalAlignment="Right"/>
                                </StackPanel>
                            </Grid>
                        </Border>

                        <Border Grid.Row="1" Background="{DynamicResource BgDark}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="6" Margin="0,0,0,10">
                            <Grid>
                                <TextBlock Name="lblMapPlaceholder" Text="Map preview will appear here after route calculation..."
                                           Foreground="{DynamicResource TextSecondary}" FontSize="14" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                <Image Name="imgMapPreview" Stretch="Uniform" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                            </Grid>
                        </Border>

                        <Border Grid.Row="2" Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="10">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Name="lblGoogleUrlDisplay" Text="No generated link" Foreground="{DynamicResource TextSecondary}" FontSize="12" VerticalAlignment="Center" TextTrimming="CharacterEllipsis" Margin="0,0,10,0"/>
                                <Button Name="btnOpenGoogleMaps" Grid.Column="1" Content="🌐 Google Maps" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="0,0,6,0" Padding="10,6" IsEnabled="False"/>
                                <Button Name="btnCopyUrl" Grid.Column="2" Content="📋 Copy Link" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="0,0,6,0" Padding="10,6" IsEnabled="False"/>
                                <Button Name="btnSaveMapAs" Grid.Column="3" Content="💾 Save Map As..." Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Padding="10,6" IsEnabled="False"/>
                            </Grid>
                        </Border>
                    </Grid>
                </Grid>
            </TabItem>

            <!-- TAB 2: BATCH DATA PROCESSING -->
            <TabItem Name="tabItemBatch" Header="📁 Batch File Processing">
                <Grid Margin="0,10,0,0">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <Border Grid.Row="0" Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="14" Margin="0,0,0,10">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>

                            <Grid Grid.Row="0" Margin="0,0,0,10">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Name="lblBatchInputFile" Text="Input File (JSON/CSV/XLSX):" VerticalAlignment="Center" Foreground="{DynamicResource TextSecondary}" Margin="0,0,10,0"/>
                                <TextBox Name="txtBatchFilePath" Grid.Column="1" VerticalAlignment="Center"/>
                                <Button Name="btnBrowseBatchFile" Grid.Column="2" Content="📂 Browse File..." Background="#2563EB" Margin="6,0,0,0"/>
                                <Button Name="btnReloadBatchFile" Grid.Column="3" Content="🔄 Reload" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="6,0,0,0"/>
                            </Grid>

                            <Grid Grid.Row="1">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                                    <TextBlock Name="lblBatchFileInfo" Text="No file loaded." Foreground="{DynamicResource TextSecondary}" FontSize="12"/>
                                </StackPanel>
                                <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center" Margin="10,0">
                                    <TextBlock Name="lblBatchDefaultRouteType" Text="Default route type:" Foreground="{DynamicResource TextSecondary}" FontSize="12" VerticalAlignment="Center" Margin="0,0,6,0"/>
                                    <ComboBox Name="cmbBatchRouteType" Width="170">
                                        <ComboBoxItem Content="From Source / Default" Tag="FromSource" IsSelected="True"/>
                                        <ComboBoxItem Content="Fastest (Najszybsza)" Tag="Fastest"/>
                                        <ComboBoxItem Content="Shortest (Najkrótsza)" Tag="Shortest"/>
                                        <ComboBoxItem Content="Eco (Fuel Efficient)" Tag="Eco"/>
                                    </ComboBox>
                                </StackPanel>
                                <StackPanel Grid.Column="3" Orientation="Horizontal">
                                    <Button Name="btnStartBatch" Content="▶ Start Processing" Background="#10B981" Foreground="#FFFFFF" Padding="14,7" FontWeight="Bold" Margin="0,0,6,0"/>
                                    <Button Name="btnStopBatch" Content="⏹ Stop" Background="#EF4444" Foreground="#FFFFFF" Padding="12,7" IsEnabled="False"/>
                                </StackPanel>
                            </Grid>
                        </Grid>
                    </Border>

                    <TabControl Name="tabBatchSub" Grid.Row="1" Background="Transparent" BorderThickness="0">
                        <TabItem Name="tabSubInput" Header="📋 Input Data Preview">
                            <Border Background="{DynamicResource BgDark}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="6" Margin="0,6,0,0">
                                <DataGrid Name="dgBatchInput"/>
                            </Border>
                        </TabItem>

                        <TabItem Name="tabSubResults" Header="📊 Calculation Results">
                            <Border Background="{DynamicResource BgDark}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="6" Margin="0,6,0,0">
                                <DataGrid Name="dgBatchResults">
                                    <DataGrid.Columns>
                                        <DataGridTextColumn Header="ID" Binding="{Binding Id}" Width="45"/>
                                        <DataGridTextColumn Header="Route Name" Binding="{Binding Name}" Width="170"/>
                                        <DataGridTextColumn Header="Origin (Start)" Binding="{Binding Start}" Width="190"/>
                                        <DataGridTextColumn Header="Destination (End)" Binding="{Binding End}" Width="190"/>
                                        <DataGridTextColumn Header="Waypoints" Binding="{Binding WaypointsCount}" Width="75"/>
                                        <DataGridTextColumn Header="Type" Binding="{Binding RouteType}" Width="75"/>
                                        <DataGridTextColumn Header="Distance (km)" Binding="{Binding DistanceKm}" Width="95"/>
                                        <DataGridTextColumn Header="Duration (min)" Binding="{Binding DurationMin}" Width="85"/>
                                        <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="110"/>
                                        <DataGridTextColumn Header="PNG Map" Binding="{Binding MapPath}" Width="*"/>
                                    </DataGrid.Columns>
                                </DataGrid>
                            </Border>
                        </TabItem>

                        <TabItem Name="tabSubPoints" Header="📍 Points Detail">
                            <Border Background="{DynamicResource BgDark}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="6" Margin="0,6,0,0">
                                <DataGrid Name="dgBatchPoints">
                                    <DataGrid.Columns>
                                        <DataGridTextColumn Header="Route ID" Binding="{Binding RouteId}" Width="65"/>
                                        <DataGridTextColumn Header="Route Name" Binding="{Binding RouteName}" Width="150"/>
                                        <DataGridTextColumn Header="No." Binding="{Binding PointOrder}" Width="45"/>
                                        <DataGridTextColumn Header="Point Type" Binding="{Binding PointType}" Width="90"/>
                                        <DataGridTextColumn Header="Original Address" Binding="{Binding OriginalAddress}" Width="220"/>
                                        <DataGridTextColumn Header="Geocoded Address" Binding="{Binding GeocodedAddress}" Width="240"/>
                                        <DataGridTextColumn Header="Geocode Status" Binding="{Binding GeocodeStatus}" Width="170"/>
                                        <DataGridTextColumn Header="Match Type" Binding="{Binding MatchType}" Width="110"/>
                                        <DataGridTextColumn Header="Fallback?" Binding="{Binding IsFallback}" Width="75"/>
                                        <DataGridTextColumn Header="Latitude" Binding="{Binding Latitude}" Width="85"/>
                                        <DataGridTextColumn Header="Longitude" Binding="{Binding Longitude}" Width="85"/>
                                    </DataGrid.Columns>
                                </DataGrid>
                            </Border>
                        </TabItem>

                        <TabItem Name="tabSubLog" Header="📝 Activity Log">
                            <Border Background="{DynamicResource BgDark}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="6" Margin="0,6,0,0">
                                <TextBox Name="txtBatchLog" IsReadOnly="True" TextWrapping="Wrap"
                                         VerticalScrollBarVisibility="Auto" FontFamily="Consolas, monospace"
                                         FontSize="12" Background="{DynamicResource LogBg}" Foreground="{DynamicResource LogFg}"/>
                            </Border>
                        </TabItem>
                    </TabControl>

                    <Border Grid.Row="2" Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="12" Margin="0,10,0,0">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="230"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>

                            <StackPanel Grid.Column="0" VerticalAlignment="Center">
                                <TextBlock Name="lblBatchProgressText" Text="Ready" FontSize="12" Foreground="{DynamicResource TextSecondary}"/>
                                <ProgressBar Name="pbBatchProgress" Height="14" Minimum="0" Maximum="100" Value="0" Margin="0,4,0,0" Foreground="#10B981" Background="{DynamicResource BgDark}"/>
                            </StackPanel>

                            <TextBlock Name="lblBatchStats" Grid.Column="1" Text="" Foreground="#10B981" FontSize="13" FontWeight="SemiBold" VerticalAlignment="Center" Margin="20,0"/>

                            <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
                                <Button Name="btnOpenOutputDir" Content="📂 Open Output Folder" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="0,0,6,0" Padding="10,6"/>
                                <Button Name="btnExportExcel" Content="📊 Export Excel" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="0,0,6,0" Padding="10,6"/>
                                <Button Name="btnExportCsv" Content="📄 CSV" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="0,0,6,0" Padding="10,6"/>
                                <Button Name="btnExportJson" Content="📋 JSON" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Padding="10,6"/>
                            </StackPanel>
                        </Grid>
                    </Border>
                </Grid>
            </TabItem>

            <!-- TAB 3: SETTINGS & API KEY -->
            <TabItem Name="tabItemSettings" Header="⚙ Settings &amp; API Key">
                <ScrollViewer VerticalScrollBarVisibility="Auto" Margin="0,10,0,0">
                    <StackPanel MaxWidth="780" HorizontalAlignment="Left">
                        <Border Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="16" Margin="0,0,0,14">
                            <StackPanel>
                                <TextBlock Name="lblSettingsApiHeader" Text="Google Maps API Key" FontSize="16" FontWeight="Bold" Foreground="{DynamicResource TextPrimary}" Margin="0,0,0,8"/>
                                <TextBlock Name="lblSettingsApiDesc" Text="Required for Geocoding API, Routes API v2, and Static Maps API." FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,10"/>

                                <TextBlock Name="lblSettingsApiLabel" Text="API Key:" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,4"/>
                                <Grid Margin="0,0,0,8">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <PasswordBox Name="txtSettingsApiKey"/>
                                    <TextBox Name="txtSettingsApiKeyVisible" Visibility="Collapsed"/>
                                    <Button Name="btnToggleKeyVisibility" Grid.Column="1" Content="👁 Show" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="6,0,0,0" Padding="10,6"/>
                                    <Button Name="btnTestApiKey" Grid.Column="2" Content="🔍 Test Key" Background="#2563EB" Margin="6,0,0,0" Padding="12,6"/>
                                </Grid>

                                <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                                    <CheckBox Name="chkRememberKey" Content="Remember securely on this computer (DPAPI CurrentUser encryption)" IsChecked="True" Foreground="{DynamicResource TextSecondary}" FontSize="12"/>
                                </StackPanel>

                                <TextBlock Name="lblKeyTestResult" Text="" FontSize="12" FontWeight="SemiBold"/>
                            </StackPanel>
                        </Border>

                        <Border Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="16" Margin="0,0,0,14">
                            <StackPanel>
                                <TextBlock Name="lblSettingsPrefHeader" Text="Default Generation Preferences" FontSize="16" FontWeight="Bold" Foreground="{DynamicResource TextPrimary}" Margin="0,0,0,12"/>

                                <TextBlock Name="lblSettingsDefaultRouteType" Text="Default route type:" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,4"/>
                                <ComboBox Name="cmbDefaultRouteType" Margin="0,0,0,12">
                                    <ComboBoxItem Content="Fastest (Najszybsza)" Tag="Fastest" IsSelected="True"/>
                                    <ComboBoxItem Content="Shortest (Najkrótsza)" Tag="Shortest"/>
                                    <ComboBoxItem Content="Eco (Fuel Efficient)" Tag="Eco"/>
                                </ComboBox>

                                <TextBlock Name="lblSettingsDefaultEmission" Text="Default engine type for Eco routes:" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,4"/>
                                <ComboBox Name="cmbDefaultEmission" Margin="0,0,0,12">
                                    <ComboBoxItem Content="Gasoline (Benzyna)" Tag="GASOLINE" IsSelected="True"/>
                                    <ComboBoxItem Content="Diesel" Tag="DIESEL"/>
                                    <ComboBoxItem Content="Hybrid" Tag="HYBRID"/>
                                    <ComboBoxItem Content="Electric" Tag="ELECTRIC"/>
                                </ComboBox>

                                <TextBlock Name="lblSettingsDefaultMapSize" Text="Default dimensions for generated PNG map:" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,4"/>
                                <ComboBox Name="cmbDefaultMapSize" Margin="0,0,0,12">
                                    <ComboBoxItem Content="900 x 600 px (Recommended Standard)" Tag="900x600" IsSelected="True"/>
                                    <ComboBoxItem Content="1024 x 768 px (High Res)" Tag="1024x768"/>
                                    <ComboBoxItem Content="1280 x 720 px (HD 16:9)" Tag="1280x720"/>
                                    <ComboBoxItem Content="640 x 640 px (Square)" Tag="640x640"/>
                                    <ComboBoxItem Content="1600 x 900 px (Full HD 16:9)" Tag="1600x900"/>
                                </ComboBox>

                                <TextBlock Name="lblSettingsOutputDir" Text="Results Output Folder:" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,4"/>
                                <Grid Margin="0,0,0,12">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBox Name="txtSettingsOutputDir"/>
                                    <Button Name="btnBrowseSettingsOutputDir" Grid.Column="1" Content="📂 Browse..." Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="6,0,0,0"/>
                                </Grid>
                            </StackPanel>
                        </Border>

                        <!-- MAP OVERLAY CARD -->
                        <Border Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="16" Margin="0,0,0,14">
                            <StackPanel>
                                <TextBlock Name="lblSettingsOverlayHeader" Text="Map Overlay &amp; Banners (Top / Bottom)" FontSize="16" FontWeight="Bold" Foreground="{DynamicResource TextPrimary}" Margin="0,0,0,6"/>
                                <TextBlock Name="lblSettingsOverlayDesc" Text="Configure whether to display top and bottom banner panels, and choose which properties appear on each panel, line order, and alignment." FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,12" TextWrapping="Wrap"/>

                                <StackPanel Orientation="Horizontal" Margin="0,0,0,12">
                                    <CheckBox Name="chkEnableTopOverlay" Content="Enable Top Banner" IsChecked="True" Foreground="{DynamicResource TextPrimary}" FontSize="12" FontWeight="SemiBold" Margin="0,0,24,0"/>
                                    <CheckBox Name="chkEnableBottomOverlay" Content="Enable Bottom Banner" IsChecked="True" Foreground="{DynamicResource TextPrimary}" FontSize="12" FontWeight="SemiBold"/>
                                </StackPanel>

                                <Border Background="{DynamicResource BgDark}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="6" Padding="10" Margin="0,0,0,10">
                                    <Grid Name="gridOverlayConfig">
                                        <Grid.RowDefinitions>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                        </Grid.RowDefinitions>
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="220"/>
                                            <ColumnDefinition Width="65"/>
                                            <ColumnDefinition Width="135"/>
                                            <ColumnDefinition Width="135"/>
                                            <ColumnDefinition Width="110"/>
                                        </Grid.ColumnDefinitions>

                                        <!-- Header Row -->
                                        <TextBlock Name="lblColPropName" Grid.Row="0" Grid.Column="0" Text="Property" FontWeight="Bold" Foreground="{DynamicResource TextSecondary}" FontSize="12" Margin="4,2,4,8"/>
                                        <TextBlock Name="lblColPropShow" Grid.Row="0" Grid.Column="1" Text="Show" FontWeight="Bold" Foreground="{DynamicResource TextSecondary}" FontSize="12" Margin="4,2,4,8" HorizontalAlignment="Center"/>
                                        <TextBlock Name="lblColPropPanel" Grid.Row="0" Grid.Column="2" Text="Panel" FontWeight="Bold" Foreground="{DynamicResource TextSecondary}" FontSize="12" Margin="4,2,4,8"/>
                                        <TextBlock Name="lblColPropAlign" Grid.Row="0" Grid.Column="3" Text="Alignment" FontWeight="Bold" Foreground="{DynamicResource TextSecondary}" FontSize="12" Margin="4,2,4,8"/>
                                        <TextBlock Name="lblColPropOrder" Grid.Row="0" Grid.Column="4" Text="Line / Order" FontWeight="Bold" Foreground="{DynamicResource TextSecondary}" FontSize="12" Margin="4,2,4,8" HorizontalAlignment="Center"/>

                                        <!-- Row 1: StartGeocoded -->
                                        <TextBlock Name="lblProp_StartGeocoded" Grid.Row="1" Grid.Column="0" Text="Start Address (Geocoded)" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_StartGeocoded" Grid.Row="1" Grid.Column="1" IsChecked="True" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_StartGeocoded" Grid.Row="1" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom" IsSelected="True"/>
                                            <ComboBoxItem Content="Top" Tag="Top"/>
                                            <ComboBoxItem Content="None" Tag="None"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_StartGeocoded" Grid.Row="1" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left" IsSelected="True"/>
                                            <ComboBoxItem Content="Center" Tag="Center"/>
                                            <ComboBoxItem Content="Right" Tag="Right"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_StartGeocoded" Grid.Row="1" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1" IsSelected="True"/>
                                            <ComboBoxItem Content="2" Tag="2"/>
                                            <ComboBoxItem Content="3" Tag="3"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>

                                        <!-- Row 2: EndGeocoded -->
                                        <TextBlock Name="lblProp_EndGeocoded" Grid.Row="2" Grid.Column="0" Text="End Address (Geocoded)" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_EndGeocoded" Grid.Row="2" Grid.Column="1" IsChecked="True" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_EndGeocoded" Grid.Row="2" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom" IsSelected="True"/>
                                            <ComboBoxItem Content="Top" Tag="Top"/>
                                            <ComboBoxItem Content="None" Tag="None"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_EndGeocoded" Grid.Row="2" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left" IsSelected="True"/>
                                            <ComboBoxItem Content="Center" Tag="Center"/>
                                            <ComboBoxItem Content="Right" Tag="Right"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_EndGeocoded" Grid.Row="2" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1"/>
                                            <ComboBoxItem Content="2" Tag="2" IsSelected="True"/>
                                            <ComboBoxItem Content="3" Tag="3"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>

                                        <!-- Row 3: Distance -->
                                        <TextBlock Name="lblProp_Distance" Grid.Row="3" Grid.Column="0" Text="Total Distance" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_Distance" Grid.Row="3" Grid.Column="1" IsChecked="True" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_Distance" Grid.Row="3" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom" IsSelected="True"/>
                                            <ComboBoxItem Content="Top" Tag="Top"/>
                                            <ComboBoxItem Content="None" Tag="None"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_Distance" Grid.Row="3" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left" IsSelected="True"/>
                                            <ComboBoxItem Content="Center" Tag="Center"/>
                                            <ComboBoxItem Content="Right" Tag="Right"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_Distance" Grid.Row="3" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1"/>
                                            <ComboBoxItem Content="2" Tag="2"/>
                                            <ComboBoxItem Content="3" Tag="3" IsSelected="True"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>

                                        <!-- Row 4: Duration -->
                                        <TextBlock Name="lblProp_Duration" Grid.Row="4" Grid.Column="0" Text="Total Time" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_Duration" Grid.Row="4" Grid.Column="1" IsChecked="True" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_Duration" Grid.Row="4" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom" IsSelected="True"/>
                                            <ComboBoxItem Content="Top" Tag="Top"/>
                                            <ComboBoxItem Content="None" Tag="None"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_Duration" Grid.Row="4" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left"/>
                                            <ComboBoxItem Content="Center" Tag="Center" IsSelected="True"/>
                                            <ComboBoxItem Content="Right" Tag="Right"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_Duration" Grid.Row="4" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1"/>
                                            <ComboBoxItem Content="2" Tag="2"/>
                                            <ComboBoxItem Content="3" Tag="3" IsSelected="True"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>

                                        <!-- Row 5: Timestamp -->
                                        <TextBlock Name="lblProp_Timestamp" Grid.Row="5" Grid.Column="0" Text="Generation Timestamp" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_Timestamp" Grid.Row="5" Grid.Column="1" IsChecked="True" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_Timestamp" Grid.Row="5" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom" IsSelected="True"/>
                                            <ComboBoxItem Content="Top" Tag="Top"/>
                                            <ComboBoxItem Content="None" Tag="None"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_Timestamp" Grid.Row="5" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left"/>
                                            <ComboBoxItem Content="Center" Tag="Center"/>
                                            <ComboBoxItem Content="Right" Tag="Right" IsSelected="True"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_Timestamp" Grid.Row="5" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1"/>
                                            <ComboBoxItem Content="2" Tag="2"/>
                                            <ComboBoxItem Content="3" Tag="3" IsSelected="True"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>

                                        <!-- Row 6: RouteName -->
                                        <TextBlock Name="lblProp_RouteName" Grid.Row="6" Grid.Column="0" Text="Route Name" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_RouteName" Grid.Row="6" Grid.Column="1" IsChecked="True" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_RouteName" Grid.Row="6" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom"/>
                                            <ComboBoxItem Content="Top" Tag="Top" IsSelected="True"/>
                                            <ComboBoxItem Content="None" Tag="None"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_RouteName" Grid.Row="6" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left" IsSelected="True"/>
                                            <ComboBoxItem Content="Center" Tag="Center"/>
                                            <ComboBoxItem Content="Right" Tag="Right"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_RouteName" Grid.Row="6" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1" IsSelected="True"/>
                                            <ComboBoxItem Content="2" Tag="2"/>
                                            <ComboBoxItem Content="3" Tag="3"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>

                                        <!-- Row 7: RouteType -->
                                        <TextBlock Name="lblProp_RouteType" Grid.Row="7" Grid.Column="0" Text="Route Type" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_RouteType" Grid.Row="7" Grid.Column="1" IsChecked="True" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_RouteType" Grid.Row="7" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom"/>
                                            <ComboBoxItem Content="Top" Tag="Top" IsSelected="True"/>
                                            <ComboBoxItem Content="None" Tag="None"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_RouteType" Grid.Row="7" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left"/>
                                            <ComboBoxItem Content="Center" Tag="Center"/>
                                            <ComboBoxItem Content="Right" Tag="Right" IsSelected="True"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_RouteType" Grid.Row="7" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1" IsSelected="True"/>
                                            <ComboBoxItem Content="2" Tag="2"/>
                                            <ComboBoxItem Content="3" Tag="3"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>

                                        <!-- Row 8: Waypoints -->
                                        <TextBlock Name="lblProp_Waypoints" Grid.Row="8" Grid.Column="0" Text="Intermediate Stops (Waypoints)" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_Waypoints" Grid.Row="8" Grid.Column="1" IsChecked="False" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_Waypoints" Grid.Row="8" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom" IsSelected="True"/>
                                            <ComboBoxItem Content="Top" Tag="Top"/>
                                            <ComboBoxItem Content="None" Tag="None"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_Waypoints" Grid.Row="8" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left" IsSelected="True"/>
                                            <ComboBoxItem Content="Center" Tag="Center"/>
                                            <ComboBoxItem Content="Right" Tag="Right"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_Waypoints" Grid.Row="8" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1"/>
                                            <ComboBoxItem Content="2" Tag="2" IsSelected="True"/>
                                            <ComboBoxItem Content="3" Tag="3"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>

                                        <!-- Row 9: StartRaw -->
                                        <TextBlock Name="lblProp_StartRaw" Grid.Row="9" Grid.Column="0" Text="Start Address (Raw Input)" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_StartRaw" Grid.Row="9" Grid.Column="1" IsChecked="False" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_StartRaw" Grid.Row="9" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom"/>
                                            <ComboBoxItem Content="Top" Tag="Top"/>
                                            <ComboBoxItem Content="None" Tag="None" IsSelected="True"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_StartRaw" Grid.Row="9" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left" IsSelected="True"/>
                                            <ComboBoxItem Content="Center" Tag="Center"/>
                                            <ComboBoxItem Content="Right" Tag="Right"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_StartRaw" Grid.Row="9" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1" IsSelected="True"/>
                                            <ComboBoxItem Content="2" Tag="2"/>
                                            <ComboBoxItem Content="3" Tag="3"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>

                                        <!-- Row 10: EndRaw -->
                                        <TextBlock Name="lblProp_EndRaw" Grid.Row="10" Grid.Column="0" Text="End Address (Raw Input)" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_EndRaw" Grid.Row="10" Grid.Column="1" IsChecked="False" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_EndRaw" Grid.Row="10" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom"/>
                                            <ComboBoxItem Content="Top" Tag="Top"/>
                                            <ComboBoxItem Content="None" Tag="None" IsSelected="True"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_EndRaw" Grid.Row="10" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left" IsSelected="True"/>
                                            <ComboBoxItem Content="Center" Tag="Center"/>
                                            <ComboBoxItem Content="Right" Tag="Right"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_EndRaw" Grid.Row="10" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1"/>
                                            <ComboBoxItem Content="2" Tag="2" IsSelected="True"/>
                                            <ComboBoxItem Content="3" Tag="3"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>
                                    </Grid>
                                </Border>

                                <StackPanel Orientation="Horizontal" Margin="0,2,0,0">
                                    <Button Name="btnResetOverlayConfig" Content="🔄 Reset to Default Layout" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Padding="12,6"/>
                                </StackPanel>
                            </StackPanel>
                        </Border>

                        <Border Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="16" Margin="0,0,0,14">
                            <StackPanel>
                                <TextBlock Name="lblSettingsLangHeader" Text="Language &amp; Localization" FontSize="16" FontWeight="Bold" Foreground="{DynamicResource TextPrimary}" Margin="0,0,0,8"/>
                                <TextBlock Name="lblSettingsLangLabel" Text="Application and Google Maps API Language:" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,6"/>
                                <ComboBox Name="cmbSettingsLanguage" Margin="0,0,0,10"/>
                                <StackPanel Orientation="Horizontal">
                                    <Button Name="btnOpenLangFile" Content="📂 Open Localization File (localization.json)" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Padding="10,6" Margin="0,0,8,0" ToolTip="Open the external localization file to edit or add new languages"/>
                                    <Button Name="btnReloadLang" Content="🔄 Reload Languages" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Padding="10,6" ToolTip="Reload language definitions from disk"/>
                                </StackPanel>
                            </StackPanel>
                        </Border>

                        <Border Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="16" Margin="0,0,0,14">
                            <StackPanel>
                                <TextBlock Name="lblSettingsThemeHeader" Text="Appearance &amp; Theme" FontSize="16" FontWeight="Bold" Foreground="{DynamicResource TextPrimary}" Margin="0,0,0,8"/>
                                <TextBlock Name="lblSettingsThemeLabel" Text="Application Theme (Color Scheme):" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,6"/>
                                <ComboBox Name="cmbSettingsTheme" Margin="0,0,0,0">
                                    <ComboBoxItem Content="🌙 Dark" Tag="Dark" IsSelected="True"/>
                                    <ComboBoxItem Content="☀️ Light" Tag="Light"/>
                                </ComboBox>
                            </StackPanel>
                        </Border>

                        <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                            <Button Name="btnSaveSettings" Content="💾 SAVE SETTINGS" Background="#10B981" Foreground="#FFFFFF" Padding="14,10" FontWeight="Bold" Width="200"/>
                            <Button Name="btnOpenLogFile" Content="📋 OPEN LOG FILE" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Padding="14,10" FontWeight="SemiBold" Margin="10,0,0,0"/>
                        </StackPanel>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
        </TabControl>

        <!-- Footer -->
        <Border Grid.Row="2" Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="6" Padding="10,6" Margin="0,10,0,0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Name="lblFooterStatus" Text="Ready." Foreground="{DynamicResource TextSecondary}" FontSize="12" VerticalAlignment="Center"/>
                <TextBlock Name="lblFooterVersion" Grid.Column="1" Text="Google Maps Routes v2.0" Foreground="{DynamicResource TextSecondary}" FontSize="12" VerticalAlignment="Center"/>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# ── 8. Tworzenie okna WPF z XAML ─────────────────────────────────────────────
$reader = [System.Xml.XmlNodeReader]::new($xaml)
$window = [System.Windows.Markup.XamlReader]::Load($reader)

# Zastosowanie DWM Dark Mode dla okna (zgodnie z motywem)
$window.Add_SourceInitialized({
    try {
        $helper = [System.Windows.Interop.WindowInteropHelper]::new($window)
        $val = if ($script:CurrentTheme -eq 'Light') { 0 } else { 1 }
        [DwmDarkWindow]::DwmSetWindowAttribute($helper.Handle, 20, [ref]$val, 4)
    } catch {}
})

# ── 9. Pobranie referencji do elementów UI ───────────────────────────────────
$txtHeaderTitle      = $window.FindName('txtHeaderTitle')
$txtHeaderSubtitle   = $window.FindName('txtHeaderSubtitle')
$cmbAppLanguage      = $window.FindName('cmbAppLanguage')
$btnThemeToggle      = $window.FindName('btnThemeToggle')
$lblApiBadge         = $window.FindName('lblApiBadge')
$btnQuickSettings    = $window.FindName('btnQuickSettings')
$tabMain             = $window.FindName('tabMain')
$tabItemManual       = $window.FindName('tabItemManual')
$tabItemBatch        = $window.FindName('tabItemBatch')
$tabItemSettings     = $window.FindName('tabItemSettings')
$lblFooterStatus     = $window.FindName('lblFooterStatus')
$lblFooterVersion    = $window.FindName('lblFooterVersion')

# Tab 1: Manual
$lblManualRoutePointsHeader = $window.FindName('lblManualRoutePointsHeader')
$lblManualOrigin            = $window.FindName('lblManualOrigin')
$lblManualWaypoints         = $window.FindName('lblManualWaypoints')
$lblManualDestination       = $window.FindName('lblManualDestination')
$lblManualRouteName         = $window.FindName('lblManualRouteName')
$lblManualOptHeader         = $window.FindName('lblManualOptHeader')
$lblManualEmission          = $window.FindName('lblManualEmission')
$lblHeaderDist              = $window.FindName('lblHeaderDist')
$lblHeaderDur               = $window.FindName('lblHeaderDur')
$lblHeaderType              = $window.FindName('lblHeaderType')

# Tab 1: Manual
$txtManualStart      = $window.FindName('txtManualStart')
$btnClearManualStart = $window.FindName('btnClearManualStart')
$txtNewWaypoint      = $window.FindName('txtNewWaypoint')
$btnAddWaypoint      = $window.FindName('btnAddWaypoint')
$lstWaypoints        = $window.FindName('lstWaypoints')
$btnWpUp             = $window.FindName('btnWpUp')
$btnWpDown           = $window.FindName('btnWpDown')
$btnWpRemove         = $window.FindName('btnWpRemove')
$btnWpClear          = $window.FindName('btnWpClear')
$txtManualEnd        = $window.FindName('txtManualEnd')
$btnClearManualEnd   = $window.FindName('btnClearManualEnd')
$txtManualName       = $window.FindName('txtManualName')
$rbTypeFastest       = $window.FindName('rbTypeFastest')
$rbTypeShortest      = $window.FindName('rbTypeShortest')
$rbTypeEco           = $window.FindName('rbTypeEco')
$pnlEmission         = $window.FindName('pnlEmission')
$cmbEmission         = $window.FindName('cmbEmission')
$chkTrafficAware     = $window.FindName('chkTrafficAware')
$btnCalculateManual  = $window.FindName('btnCalculateManual')
$lblManualDist       = $window.FindName('lblManualDist')
$lblManualTime       = $window.FindName('lblManualTime')
$lblManualType       = $window.FindName('lblManualType')
$lblManualStatus     = $window.FindName('lblManualStatus')
$lblMapPlaceholder   = $window.FindName('lblMapPlaceholder')
$imgMapPreview       = $window.FindName('imgMapPreview')
$lblGoogleUrlDisplay = $window.FindName('lblGoogleUrlDisplay')
$btnOpenGoogleMaps   = $window.FindName('btnOpenGoogleMaps')
$btnCopyUrl          = $window.FindName('btnCopyUrl')
$btnSaveMapAs        = $window.FindName('btnSaveMapAs')

# Tab 2: Batch
$lblBatchInputFile       = $window.FindName('lblBatchInputFile')
$lblBatchDefaultRouteType= $window.FindName('lblBatchDefaultRouteType')
$tabSubInput             = $window.FindName('tabSubInput')
$tabSubResults           = $window.FindName('tabSubResults')
$tabSubPoints            = $window.FindName('tabSubPoints')
$tabSubLog               = $window.FindName('tabSubLog')
$txtBatchFilePath    = $window.FindName('txtBatchFilePath')
$btnBrowseBatchFile  = $window.FindName('btnBrowseBatchFile')
$btnReloadBatchFile  = $window.FindName('btnReloadBatchFile')
$lblBatchFileInfo    = $window.FindName('lblBatchFileInfo')
$cmbBatchRouteType   = $window.FindName('cmbBatchRouteType')
$btnStartBatch       = $window.FindName('btnStartBatch')
$btnStopBatch        = $window.FindName('btnStopBatch')
$tabBatchSub        = $window.FindName('tabBatchSub')
$dgBatchInput        = $window.FindName('dgBatchInput')
$dgBatchResults      = $window.FindName('dgBatchResults')
$dgBatchPoints       = $window.FindName('dgBatchPoints')
$txtBatchLog         = $window.FindName('txtBatchLog')
$lblBatchProgressText= $window.FindName('lblBatchProgressText')
$pbBatchProgress     = $window.FindName('pbBatchProgress')
$lblBatchStats       = $window.FindName('lblBatchStats')
$btnOpenOutputDir    = $window.FindName('btnOpenOutputDir')
$btnExportExcel      = $window.FindName('btnExportExcel')
$btnExportCsv        = $window.FindName('btnExportCsv')
$btnExportJson       = $window.FindName('btnExportJson')

# Tab 3: Settings
$txtSettingsApiKey          = $window.FindName('txtSettingsApiKey')
$txtSettingsApiKeyVisible   = $window.FindName('txtSettingsApiKeyVisible')
$btnToggleKeyVisibility     = $window.FindName('btnToggleKeyVisibility')
$btnTestApiKey              = $window.FindName('btnTestApiKey')
$chkRememberKey             = $window.FindName('chkRememberKey')
$lblKeyTestResult           = $window.FindName('lblKeyTestResult')
$cmbDefaultRouteType        = $window.FindName('cmbDefaultRouteType')
$cmbDefaultEmission         = $window.FindName('cmbDefaultEmission')
$cmbDefaultMapSize          = $window.FindName('cmbDefaultMapSize')
$txtSettingsOutputDir       = $window.FindName('txtSettingsOutputDir')
$btnBrowseSettingsOutputDir = $window.FindName('btnBrowseSettingsOutputDir')
$btnSaveSettings            = $window.FindName('btnSaveSettings')
$btnOpenLogFile             = $window.FindName('btnOpenLogFile')
$lblSettingsApiHeader       = $window.FindName('lblSettingsApiHeader')
$lblSettingsApiDesc         = $window.FindName('lblSettingsApiDesc')
$lblSettingsApiLabel        = $window.FindName('lblSettingsApiLabel')
$lblSettingsPrefHeader      = $window.FindName('lblSettingsPrefHeader')
$lblSettingsDefaultRouteType= $window.FindName('lblSettingsDefaultRouteType')
$lblSettingsDefaultEmission = $window.FindName('lblSettingsDefaultEmission')
$lblSettingsDefaultMapSize  = $window.FindName('lblSettingsDefaultMapSize')
$lblSettingsOutputDir       = $window.FindName('lblSettingsOutputDir')
$lblSettingsLangHeader      = $window.FindName('lblSettingsLangHeader')
$lblSettingsLangLabel       = $window.FindName('lblSettingsLangLabel')
$cmbSettingsLanguage        = $window.FindName('cmbSettingsLanguage')
$btnOpenLangFile            = $window.FindName('btnOpenLangFile')
$btnReloadLang              = $window.FindName('btnReloadLang')
$lblSettingsThemeHeader     = $window.FindName('lblSettingsThemeHeader')
$lblSettingsThemeLabel      = $window.FindName('lblSettingsThemeLabel')
$cmbSettingsTheme           = $window.FindName('cmbSettingsTheme')

# Tab 3: Overlay Settings
$lblSettingsOverlayHeader    = $window.FindName('lblSettingsOverlayHeader')
$lblSettingsOverlayDesc      = $window.FindName('lblSettingsOverlayDesc')
$chkEnableTopOverlay         = $window.FindName('chkEnableTopOverlay')
$chkEnableBottomOverlay      = $window.FindName('chkEnableBottomOverlay')
$lblColPropName              = $window.FindName('lblColPropName')
$lblColPropShow              = $window.FindName('lblColPropShow')
$lblColPropPanel             = $window.FindName('lblColPropPanel')
$lblColPropAlign             = $window.FindName('lblColPropAlign')
$lblColPropOrder             = $window.FindName('lblColPropOrder')
$btnResetOverlayConfig       = $window.FindName('btnResetOverlayConfig')

foreach ($key in $script:OverlayPropKeys) {
    Set-Variable -Name "lblProp_$key"  -Value ($window.FindName("lblProp_$key"))  -Scope Script
    Set-Variable -Name "chkProp_$key"  -Value ($window.FindName("chkProp_$key"))  -Scope Script
    Set-Variable -Name "cmbPanel_$key" -Value ($window.FindName("cmbPanel_$key")) -Scope Script
    Set-Variable -Name "cmbAlign_$key" -Value ($window.FindName("cmbAlign_$key")) -Scope Script
    Set-Variable -Name "cmbOrder_$key" -Value ($window.FindName("cmbOrder_$key")) -Scope Script
}

# ── 10. System wielojęzyczności i funkcje pomocnicze stanu UI ─────────────────

function Populate-LanguageDropdowns {
    $script:SuppressLangEvents = $true
    try {
        if ($cmbAppLanguage) {
            $cmbAppLanguage.Items.Clear()
            foreach ($langCode in $script:LanguagesCatalog.Keys) {
                $langObj = $script:LanguagesCatalog[$langCode]
                $cbi = [System.Windows.Controls.ComboBoxItem]::new()
                $cbi.Content = "[$($langObj.Code.ToUpper())] $($langObj.DisplayName)"
                $cbi.Tag = $langObj.Code
                if ($langObj.Code -eq $script:CurrentLanguage) { $cbi.IsSelected = $true }
                $null = $cmbAppLanguage.Items.Add($cbi)
            }
        }
        if ($cmbSettingsLanguage) {
            $cmbSettingsLanguage.Items.Clear()
            foreach ($langCode in $script:LanguagesCatalog.Keys) {
                $langObj = $script:LanguagesCatalog[$langCode]
                $cbi = [System.Windows.Controls.ComboBoxItem]::new()
                $cbi.Content = "[$($langObj.Code.ToUpper())] $($langObj.DisplayName)"
                $cbi.Tag = $langObj.Code
                if ($langObj.Code -eq $script:CurrentLanguage) { $cbi.IsSelected = $true }
                $null = $cmbSettingsLanguage.Items.Add($cbi)
            }
        }
    } finally {
        $script:SuppressLangEvents = $false
    }
}

function Set-AppTheme {
    param(
        [Parameter(Mandatory = $false)]
        [string]$Theme = 'Dark'
    )

    if ($Theme -notmatch '(?i)light|dark') { $Theme = 'Dark' }
    $isLight = ($Theme -match '(?i)light')
    $script:CurrentTheme = if ($isLight) { 'Light' } else { 'Dark' }

    $palette = if ($isLight) {
        [ordered]@{
            'BgDark'                 = '#F1F5F9'
            'BgCard'                 = '#FFFFFF'
            'BgCardHover'            = '#F8FAFC'
            'BgCardAlt'              = '#F8FAFC'
            'BorderCard'             = '#CBD5E1'
            'TextPrimary'            = '#0F172A'
            'TextSecondary'          = '#475569'
            'AccentBlue'             = '#2563EB'
            'AccentGreen'            = '#059669'
            'AccentAmber'            = '#D97706'
            'AccentRed'              = '#DC2626'
            'BgInput'                = '#FFFFFF'
            'BorderInput'            = '#CBD5E1'
            'BtnSecondaryBg'         = '#E2E8F0'
            'BtnSecondaryFg'         = '#0F172A'
            'GridLines'              = '#E2E8F0'
            'LogBg'                  = '#F8FAFC'
            'LogFg'                  = '#0369A1'
            'DataGridHeaderBg'       = '#E2E8F0'
            'DataGridHeaderFg'       = '#334155'
            'DataGridRowBg'          = '#FFFFFF'
            'DataGridAltRowBg'       = '#F8FAFC'
        }
    } else {
        [ordered]@{
            'BgDark'                 = '#0F172A'
            'BgCard'                 = '#1E293B'
            'BgCardHover'            = '#293548'
            'BgCardAlt'              = '#162032'
            'BorderCard'             = '#334155'
            'TextPrimary'            = '#F8FAFC'
            'TextSecondary'          = '#94A3B8'
            'AccentBlue'             = '#2563EB'
            'AccentGreen'            = '#10B981'
            'AccentAmber'            = '#F59E0B'
            'AccentRed'              = '#EF4444'
            'BgInput'                = '#1E293B'
            'BorderInput'            = '#334155'
            'BtnSecondaryBg'         = '#334155'
            'BtnSecondaryFg'         = '#F8FAFC'
            'GridLines'              = '#2D3748'
            'LogBg'                  = '#0A0F1D'
            'LogFg'                  = '#38BDF8'
            'DataGridHeaderBg'       = '#0F172A'
            'DataGridHeaderFg'       = '#94A3B8'
            'DataGridRowBg'          = '#1E293B'
            'DataGridAltRowBg'       = '#162032'
        }
    }

    foreach ($k in $palette.Keys) {
        $c = [System.Windows.Media.ColorConverter]::ConvertFromString($palette[$k])
        $brush = [System.Windows.Media.SolidColorBrush]::new($c)
        $brush.Freeze()
        $window.Resources[$k] = $brush
        $window.Resources["Theme_$k"] = $brush
    }
    $window.Resources['Theme_BgApp'] = $window.Resources['BgDark']
    $window.Resources['Theme_Border'] = $window.Resources['BorderCard']

    if ($window) {
        $window.Background = $window.Resources['BgDark']
        $window.Foreground = $window.Resources['TextPrimary']
    }

    # Update overlay property labels foreground
    if ($script:OverlayPropKeys) {
        foreach ($k in $script:OverlayPropKeys) {
            $lblCtrl = Get-Variable -Name "lblProp_$k" -ValueOnly -ErrorAction SilentlyContinue
            if ($lblCtrl) {
                $lblCtrl.Foreground = $window.Resources['TextPrimary']
            }
        }
    }

    # Update DWM title bar chrome
    try {
        $helper = [System.Windows.Interop.WindowInteropHelper]::new($window)
        if ($helper.Handle -ne [IntPtr]::Zero) {
            $val = if ($isLight) { 0 } else { 1 }
            [DwmDarkWindow]::DwmSetWindowAttribute($helper.Handle, 20, [ref]$val, 4)
        }
    } catch {}

    # Update Toggle Button text/icon
    if ($btnThemeToggle) {
        $btnThemeToggle.Content = if ($isLight) { (Get-LocText 'ThemeLight') } else { (Get-LocText 'ThemeDark') }
    }

    # Update Settings ComboBox
    if ($cmbSettingsTheme) {
        $script:SuppressThemeEvents = $true
        try {
            foreach ($it in $cmbSettingsTheme.Items) {
                if ($it.Tag -eq $script:CurrentTheme) {
                    $cmbSettingsTheme.SelectedItem = $it
                    break
                }
            }
        } finally {
            $script:SuppressThemeEvents = $false
        }
    }
}

function Apply-AppLanguage {
    param([string]$LanguageCode)

    if (-not [string]::IsNullOrWhiteSpace($LanguageCode) -and $script:LanguagesCatalog.Contains($LanguageCode.ToLower())) {
        $script:CurrentLanguage = $LanguageCode.ToLower()
    } else {
        $script:CurrentLanguage = 'en'
    }

    $langObj = $script:LanguagesCatalog[$script:CurrentLanguage]
    $script:CurrentGoogleLang = if ($langObj -and $langObj.GoogleCode) { $langObj.GoogleCode } else { $script:CurrentLanguage }
    $script:CurrentStrings = if ($langObj -and $langObj.Strings) { $langObj.Strings } else { @{} }

    # Synchronize ComboBoxes without triggering duplicate events
    $script:SuppressLangEvents = $true
    try {
        if ($cmbAppLanguage) {
            foreach ($item in $cmbAppLanguage.Items) {
                if ($item.Tag -eq $script:CurrentLanguage) {
                    $cmbAppLanguage.SelectedItem = $item
                    break
                }
            }
        }
        if ($cmbSettingsLanguage) {
            foreach ($item in $cmbSettingsLanguage.Items) {
                if ($item.Tag -eq $script:CurrentLanguage) {
                    $cmbSettingsLanguage.SelectedItem = $item
                    break
                }
            }
        }
    } finally {
        $script:SuppressLangEvents = $false
    }

    # Window & Header
    if ($window) { $window.Title = (Get-LocText 'AppTitle') }
    if ($txtHeaderTitle) { $txtHeaderTitle.Text = (Get-LocText 'AppTitle') }
    if ($txtHeaderSubtitle) { $txtHeaderSubtitle.Text = (Get-LocText 'AppSubtitle') }
    if ($btnQuickSettings) { $btnQuickSettings.Content = (Get-LocText 'BtnQuickSettings') }
    if ($btnThemeToggle) {
        $btnThemeToggle.ToolTip = (Get-LocText 'ThemeToggleTip')
        $btnThemeToggle.Content = if ($script:CurrentTheme -eq 'Light') { (Get-LocText 'ThemeLight') } else { (Get-LocText 'ThemeDark') }
    }

    # Tab Headers
    if ($tabItemManual) { $tabItemManual.Header = (Get-LocText 'TabManual') }
    if ($tabItemBatch) { $tabItemBatch.Header = (Get-LocText 'TabBatch') }
    if ($tabItemSettings) { $tabItemSettings.Header = (Get-LocText 'TabSettings') }

    # Tab 1: Manual Route
    if ($lblManualRoutePointsHeader) { $lblManualRoutePointsHeader.Text = (Get-LocText 'ManualHeaderRoutePoints') }
    if ($lblManualOrigin) { $lblManualOrigin.Text = (Get-LocText 'ManualOrigin') }
    if ($lblManualWaypoints) { $lblManualWaypoints.Text = (Get-LocText 'ManualWaypoints') }
    if ($txtNewWaypoint) { $txtNewWaypoint.ToolTip = (Get-LocText 'ManualWaypointsTooltip') }
    if ($btnAddWaypoint) { $btnAddWaypoint.Content = (Get-LocText 'ManualBtnAdd') }
    if ($btnWpUp) { $btnWpUp.Content = (Get-LocText 'ManualBtnUp') }
    if ($btnWpDown) { $btnWpDown.Content = (Get-LocText 'ManualBtnDown') }
    if ($btnWpRemove) { $btnWpRemove.Content = (Get-LocText 'ManualBtnRemove') }
    if ($btnWpClear) { $btnWpClear.Content = (Get-LocText 'ManualBtnClear') }
    if ($lblManualDestination) { $lblManualDestination.Text = (Get-LocText 'ManualDestination') }
    if ($lblManualRouteName) { $lblManualRouteName.Text = (Get-LocText 'ManualRouteName') }
    if ($lblManualOptHeader) { $lblManualOptHeader.Text = (Get-LocText 'ManualHeaderOptimization') }
    if ($rbTypeFastest) { $rbTypeFastest.Content = (Get-LocText 'ManualOptFastest') }
    if ($rbTypeShortest) { $rbTypeShortest.Content = (Get-LocText 'ManualOptShortest') }
    if ($rbTypeEco) { $rbTypeEco.Content = (Get-LocText 'ManualOptEco') }
    if ($lblManualEmission) { $lblManualEmission.Text = (Get-LocText 'ManualEmission') }
    if ($chkTrafficAware) { $chkTrafficAware.Content = (Get-LocText 'ManualTrafficAware') }
    if ($btnCalculateManual -and $btnCalculateManual.IsEnabled) { $btnCalculateManual.Content = (Get-LocText 'ManualBtnCalculate') }
    if ($lblHeaderDist) { $lblHeaderDist.Text = (Get-LocText 'ManualStatDistance') }
    if ($lblHeaderDur) { $lblHeaderDur.Text = (Get-LocText 'ManualStatDuration') }
    if ($lblHeaderType) { $lblHeaderType.Text = (Get-LocText 'ManualStatType') }
    if ($lblManualType) {
        $curT = $lblManualType.Text
        if ($curT -match '(?i)fast|szyb|schnell') {
            $lblManualType.Text = switch ($script:CurrentLanguage) { 'de' { 'Schnellste' } 'pl' { 'Najszybsza' } default { 'Fastest' } }
        } elseif ($curT -match '(?i)short|kr[oó]t|k[uü]rz') {
            $lblManualType.Text = switch ($script:CurrentLanguage) { 'de' { 'Kürzeste' } 'pl' { 'Najkrótsza' } default { 'Shortest' } }
        } elseif ($curT -match '(?i)eco|eko') {
            $lblManualType.Text = switch ($script:CurrentLanguage) { 'pl' { 'Eko' } default { 'Eco' } }
        }
    }
    if ($lblManualStatus -and $lblManualStatus.Text -match '(?i)idle|bereit|gotow') { $lblManualStatus.Text = (Get-LocText 'ManualStatusIdle') }
    if ($lblMapPlaceholder -and $lblMapPlaceholder.Visibility -eq [System.Windows.Visibility]::Visible) { $lblMapPlaceholder.Text = (Get-LocText 'ManualMapPlaceholder') }
    if ($lblGoogleUrlDisplay -and $lblGoogleUrlDisplay.Text -match '(?i)no generated|kein link|brak') { $lblGoogleUrlDisplay.Text = (Get-LocText 'ManualNoUrl') }
    if ($btnOpenGoogleMaps) { $btnOpenGoogleMaps.Content = (Get-LocText 'ManualBtnGoogleMaps') }
    if ($btnCopyUrl) { $btnCopyUrl.Content = (Get-LocText 'ManualBtnCopyUrl') }
    if ($btnSaveMapAs) { $btnSaveMapAs.Content = (Get-LocText 'ManualBtnSaveMapAs') }

    # Tab 2: Batch Processing
    if ($lblBatchInputFile) { $lblBatchInputFile.Text = (Get-LocText 'BatchInputFile') }
    if ($btnBrowseBatchFile) { $btnBrowseBatchFile.Content = (Get-LocText 'BatchBtnBrowse') }
    if ($btnReloadBatchFile) { $btnReloadBatchFile.Content = (Get-LocText 'BatchBtnReload') }
    if ($lblBatchFileInfo -and $lblBatchFileInfo.Text -match '(?i)no file|keine datei|brak') { $lblBatchFileInfo.Text = (Get-LocText 'BatchNoFileLoaded') }
    if ($lblBatchDefaultRouteType) { $lblBatchDefaultRouteType.Text = (Get-LocText 'BatchDefaultRouteType') }
    if ($btnStartBatch) { $btnStartBatch.Content = (Get-LocText 'BatchBtnStart') }
    if ($btnStopBatch) { $btnStopBatch.Content = (Get-LocText 'BatchBtnStop') }
    if ($tabSubInput) { $tabSubInput.Header = (Get-LocText 'BatchTabInputPreview') }
    if ($tabSubResults) { $tabSubResults.Header = (Get-LocText 'BatchTabResults') }
    if ($tabSubPoints) { $tabSubPoints.Header = (Get-LocText 'BatchTabPoints') }
    if ($tabSubLog) { $tabSubLog.Header = (Get-LocText 'BatchTabLog') }

    # Batch DataGrid Columns
    if ($dgBatchResults -and $dgBatchResults.Columns.Count -ge 10) {
        $dgBatchResults.Columns[0].Header = (Get-LocText 'BatchColId')
        $dgBatchResults.Columns[1].Header = (Get-LocText 'BatchColName')
        $dgBatchResults.Columns[2].Header = (Get-LocText 'BatchColOrigin')
        $dgBatchResults.Columns[3].Header = (Get-LocText 'BatchColDestination')
        $dgBatchResults.Columns[4].Header = (Get-LocText 'BatchColWaypoints')
        $dgBatchResults.Columns[5].Header = (Get-LocText 'BatchColType')
        $dgBatchResults.Columns[6].Header = (Get-LocText 'BatchColDistance')
        $dgBatchResults.Columns[7].Header = (Get-LocText 'BatchColDuration')
        $dgBatchResults.Columns[8].Header = (Get-LocText 'BatchColStatus')
        $dgBatchResults.Columns[9].Header = (Get-LocText 'BatchColMap')
    }

    # Points DataGrid Columns
    if ($dgBatchPoints -and $dgBatchPoints.Columns.Count -ge 9) {
        $dgBatchPoints.Columns[0].Header = (Get-LocText 'PointsColRouteId')
        $dgBatchPoints.Columns[1].Header = (Get-LocText 'PointsColRouteName')
        $dgBatchPoints.Columns[2].Header = (Get-LocText 'PointsColOrder')
        $dgBatchPoints.Columns[3].Header = (Get-LocText 'PointsColType')
        $dgBatchPoints.Columns[4].Header = (Get-LocText 'PointsColOriginalAddress')
        $dgBatchPoints.Columns[5].Header = (Get-LocText 'PointsColGeocodedAddress')
        $dgBatchPoints.Columns[6].Header = (Get-LocText 'PointsColGeocodeStatus')
        $dgBatchPoints.Columns[7].Header = (Get-LocText 'PointsColMatchType')
        $dgBatchPoints.Columns[8].Header = (Get-LocText 'PointsColIsFallback')
        if ($dgBatchPoints.Columns.Count -ge 11) {
            $dgBatchPoints.Columns[9].Header = (Get-LocText 'PointsColLatitude')
            $dgBatchPoints.Columns[10].Header = (Get-LocText 'PointsColLongitude')
        }
    }

    if ($lblBatchProgressText -and $lblBatchProgressText.Text -match '(?i)ready|bereit|gotow') { $lblBatchProgressText.Text = (Get-LocText 'BatchProgressReady') }
    if ($btnOpenOutputDir) { $btnOpenOutputDir.Content = (Get-LocText 'BatchBtnOpenOutputDir') }
    if ($btnExportExcel) { $btnExportExcel.Content = (Get-LocText 'BatchBtnExportExcel') }
    if ($btnExportCsv) { $btnExportCsv.Content = (Get-LocText 'BatchBtnExportCsv') }
    if ($btnExportJson) { $btnExportJson.Content = (Get-LocText 'BatchBtnExportJson') }

    # Tab 3: Settings
    if ($lblSettingsApiHeader) { $lblSettingsApiHeader.Text = (Get-LocText 'SettingsHeaderApi') }
    if ($lblSettingsApiDesc) { $lblSettingsApiDesc.Text = (Get-LocText 'SettingsApiDesc') }
    if ($lblSettingsApiLabel) { $lblSettingsApiLabel.Text = (Get-LocText 'SettingsApiLabel') }
    if ($btnTestApiKey -and $btnTestApiKey.IsEnabled) { $btnTestApiKey.Content = (Get-LocText 'SettingsBtnTestKey') }
    if ($chkRememberKey) { $chkRememberKey.Content = (Get-LocText 'SettingsChkRemember') }
    if ($lblSettingsPrefHeader) { $lblSettingsPrefHeader.Text = (Get-LocText 'SettingsHeaderPreferences') }
    if ($lblSettingsDefaultRouteType) { $lblSettingsDefaultRouteType.Text = (Get-LocText 'SettingsDefaultRouteType') }
    if ($lblSettingsDefaultEmission) { $lblSettingsDefaultEmission.Text = (Get-LocText 'SettingsDefaultEmission') }
    if ($lblSettingsDefaultMapSize) { $lblSettingsDefaultMapSize.Text = (Get-LocText 'SettingsDefaultMapSize') }
    if ($lblSettingsOutputDir) { $lblSettingsOutputDir.Text = (Get-LocText 'SettingsOutputDir') }
    if ($btnBrowseSettingsOutputDir) { $btnBrowseSettingsOutputDir.Content = (Get-LocText 'SettingsBtnBrowseOutputDir') }
    if ($lblSettingsOverlayHeader) { $lblSettingsOverlayHeader.Text = (Get-LocText 'SettingsHeaderOverlay') }
    if ($lblSettingsOverlayDesc) { $lblSettingsOverlayDesc.Text = (Get-LocText 'SettingsOverlayDesc') }
    if ($chkEnableTopOverlay) { $chkEnableTopOverlay.Content = (Get-LocText 'SettingsOverlayTopEnable') }
    if ($chkEnableBottomOverlay) { $chkEnableBottomOverlay.Content = (Get-LocText 'SettingsOverlayBottomEnable') }
    if ($lblColPropName) { $lblColPropName.Text = (Get-LocText 'OverlayColProperty') }
    if ($lblColPropShow) { $lblColPropShow.Text = (Get-LocText 'OverlayColShow') }
    if ($lblColPropPanel) { $lblColPropPanel.Text = (Get-LocText 'OverlayColPanel') }
    if ($lblColPropAlign) { $lblColPropAlign.Text = (Get-LocText 'OverlayColAlign') }
    if ($lblColPropOrder) { $lblColPropOrder.Text = (Get-LocText 'OverlayColOrder') }
    if ($btnResetOverlayConfig) { $btnResetOverlayConfig.Content = (Get-LocText 'SettingsOverlayBtnReset') }

    foreach ($k in $script:OverlayPropKeys) {
        $lblCtrl = Get-Variable -Name "lblProp_$k" -ValueOnly -ErrorAction SilentlyContinue
        if ($lblCtrl) {
            $lblCtrl.Text = (Get-LocText "OverlayProp$k")
        }
    }

    if ($lblSettingsLangHeader) { $lblSettingsLangHeader.Text = (Get-LocText 'SettingsHeaderLanguage') }
    if ($lblSettingsLangLabel) { $lblSettingsLangLabel.Text = (Get-LocText 'SettingsLanguageLabel') }
    if ($btnOpenLangFile) { $btnOpenLangFile.Content = (Get-LocText 'SettingsBtnOpenLangFile') }
    if ($btnReloadLang) { $btnReloadLang.Content = (Get-LocText 'SettingsBtnReloadLang') }
    if ($lblSettingsThemeHeader) { $lblSettingsThemeHeader.Text = switch ($script:CurrentLanguage) { 'de' { 'Erscheinungsbild & Design' } 'pl' { 'Wygląd i motyw' } default { 'Appearance & Theme' } } }
    if ($lblSettingsThemeLabel) { $lblSettingsThemeLabel.Text = (Get-LocText 'SettingsThemeLabel') }
    if ($cmbSettingsTheme -and $cmbSettingsTheme.Items.Count -ge 2) {
        $cmbSettingsTheme.Items[0].Content = (Get-LocText 'ThemeDark')
        $cmbSettingsTheme.Items[1].Content = (Get-LocText 'ThemeLight')
    }
    if ($btnSaveSettings) { $btnSaveSettings.Content = (Get-LocText 'SettingsBtnSave') }
    if ($btnOpenLogFile) { $btnOpenLogFile.Content = (Get-LocText 'SettingsBtnOpenLog') }

    # Footer
    if ($lblFooterStatus -and $lblFooterStatus.Text -match '(?i)ready|bereit|gotow') { $lblFooterStatus.Text = (Get-LocText 'FooterReady') }
    if ($lblFooterVersion) { $lblFooterVersion.Text = (Get-LocText 'FooterVersion') }

    # Update API badge text if checking
    if ($lblApiBadge -and $lblApiBadge.Text -match '(?i)checking|prüfe|sprawdz') {
        $lblApiBadge.Text = (Get-LocText 'ApiBadgeChecking')
    }
}

function Set-CurrentApiKey([string]$Key) {
    $txtSettingsApiKey.Password = $Key
    $txtSettingsApiKeyVisible.Text = $Key
    $script:CurrentApiKey = $Key
}

function Get-CurrentApiKey {
    if ($txtSettingsApiKeyVisible.Visibility -eq [System.Windows.Visibility]::Visible) {
        return $txtSettingsApiKeyVisible.Text.Trim()
    }
    return $txtSettingsApiKey.Password.Trim()
}

function Write-BatchLog([string]$Message, [string]$Level = 'INFO') {
    Write-AppLog -Message $Message -Level $Level -ToBatchWindow
}

function Update-ApiStatusBadge($IsValid, [string]$Message) {
    $validBool = [bool]$IsValid
    if ($validBool) {
        $lblApiBadge.Text = 'API: Active'
        $lblApiBadge.Foreground = [System.Windows.Media.Brushes]::LightGreen
        $lblKeyTestResult.Text = "✓ $Message"
        $lblKeyTestResult.Foreground = [System.Windows.Media.Brushes]::LightGreen
    } else {
        $lblApiBadge.Text = 'API: Invalid'
        $lblApiBadge.Foreground = [System.Windows.Media.Brushes]::Salmon
        $lblKeyTestResult.Text = "✕ $Message"
        $lblKeyTestResult.Foreground = [System.Windows.Media.Brushes]::Salmon
    }
}

# Inicjalizacja wartości kontrolek z zapisanego configu
$script:CurrentApiKey = ''
if (-not [string]::IsNullOrWhiteSpace($script:Config.ApiKey)) {
    Set-CurrentApiKey -Key $script:Config.ApiKey
    $lblApiBadge.Text = 'API: Configured'
    $lblApiBadge.Foreground = [System.Windows.Media.Brushes]::LightGreen
}

$chkRememberKey.IsChecked = [bool]$script:Config.RememberApiKey
$txtSettingsOutputDir.Text = if ($script:Config.LastOutputFolder) { $script:Config.LastOutputFolder } else { Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'GoogleMapsRoutes\Results' }

Populate-LanguageDropdowns
Apply-AppLanguage -LanguageCode $script:CurrentLanguage
Set-AppTheme -Theme $script:Config.Theme

# Ustawienie domyślnych ComboBoxów
foreach ($item in $cmbDefaultRouteType.Items) {
    if ($item.Tag -eq $script:Config.DefaultRouteType) { $item.IsSelected = $true; break }
}
foreach ($item in $cmbDefaultEmission.Items) {
    if ($item.Tag -eq $script:Config.DefaultEmission) { $item.IsSelected = $true; break }
}
$targetMapTag = "$($script:Config.MapWidth)x$($script:Config.MapHeight)"
foreach ($item in $cmbDefaultMapSize.Items) {
    if ($item.Tag -eq $targetMapTag) { $item.IsSelected = $true; break }
}

# Ustawienie kontrolek nakładki mapy (Overlay)
if ($script:Config.OverlayConfig) {
    Set-OverlayConfigUi $script:Config.OverlayConfig
} else {
    Reset-OverlayConfigUi
}

$script:LastGeneratedMapPath = $null
$script:LastGoogleMapsUrl    = $null
$script:LoadedBatchData      = $null
$script:BatchResultsList     = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:BatchWorkerRunning   = $false
$script:BatchCancelRequested = $false
$script:LastDataDirectory    = if ($script:Config.LastInputFolder) { $script:Config.LastInputFolder } else { $null }

# Restore last used batch input file if available
if ($script:Config.LastInputPath -and (Test-Path $script:Config.LastInputPath)) {
    $txtBatchFilePath.Text = $script:Config.LastInputPath
    try {
        Load-BatchFilePreview -Path $script:Config.LastInputPath
    } catch { }
}

# ── 11. Zdarzenia: Ustawienia i Klucz API ────────────────────────────────────
$btnQuickSettings.Add_Click({
    $tabMain.SelectedItem = $tabItemSettings
})

if ($btnThemeToggle) {
    $btnThemeToggle.Add_Click({
        $newTheme = if ($script:CurrentTheme -eq 'Light') { 'Dark' } else { 'Light' }
        Set-AppTheme -Theme $newTheme
        $script:Config.Theme = $newTheme
        $routeType = if ($cmbDefaultRouteType.SelectedItem) { $cmbDefaultRouteType.SelectedItem.Tag -as [string] } else { $script:Config.DefaultRouteType }
        $emission = if ($cmbDefaultEmission.SelectedItem) { $cmbDefaultEmission.SelectedItem.Tag -as [string] } else { $script:Config.DefaultEmission }
        Save-AppConfig -ApiKey (Get-CurrentApiKey) -RememberApiKey $chkRememberKey.IsChecked `
            -OutputFolder $txtSettingsOutputDir.Text.Trim() `
            -LastInputFolder $script:Config.LastInputFolder -LastInputPath $script:Config.LastInputPath `
            -DefaultRouteType $routeType `
            -DefaultEmission $emission `
            -Language $script:CurrentLanguage `
            -Theme $newTheme
        Write-AppLog "Theme toggled to: $newTheme" "INFO"
    })
}

if ($cmbSettingsTheme) {
    $cmbSettingsTheme.Add_SelectionChanged({
        if ($script:SuppressThemeEvents) { return }
        $sel = $cmbSettingsTheme.SelectedItem
        if ($sel -and $sel.Tag) {
            $newTheme = [string]$sel.Tag
            if ($newTheme -ne $script:CurrentTheme) {
                Set-AppTheme -Theme $newTheme
                $script:Config.Theme = $newTheme
                $routeType = if ($cmbDefaultRouteType.SelectedItem) { $cmbDefaultRouteType.SelectedItem.Tag -as [string] } else { $script:Config.DefaultRouteType }
                $emission = if ($cmbDefaultEmission.SelectedItem) { $cmbDefaultEmission.SelectedItem.Tag -as [string] } else { $script:Config.DefaultEmission }
                Save-AppConfig -ApiKey (Get-CurrentApiKey) -RememberApiKey $chkRememberKey.IsChecked `
                    -OutputFolder $txtSettingsOutputDir.Text.Trim() `
                    -LastInputFolder $script:Config.LastInputFolder -LastInputPath $script:Config.LastInputPath `
                    -DefaultRouteType $routeType `
                    -DefaultEmission $emission `
                    -Language $script:CurrentLanguage `
                    -Theme $newTheme
                Write-AppLog "Theme changed to: $newTheme" "INFO"
            }
        }
    })
}

$cmbAppLanguage.Add_SelectionChanged({
    if ($script:SuppressLangEvents) { return }
    $sel = $cmbAppLanguage.SelectedItem
    if ($sel -and $sel.Tag) {
        $newLang = [string]$sel.Tag
        Apply-AppLanguage -LanguageCode $newLang
        Save-AppConfig -ApiKey (Get-CurrentApiKey) -RememberApiKey $chkRememberKey.IsChecked `
            -OutputFolder $txtSettingsOutputDir.Text.Trim() `
            -LastInputFolder $script:Config.LastInputFolder -LastInputPath $script:Config.LastInputPath `
            -DefaultRouteType ($cmbDefaultRouteType.SelectedItem.Tag -as [string]) `
            -DefaultEmission ($cmbDefaultEmission.SelectedItem.Tag -as [string]) `
            -Language $newLang
        Write-AppLog "Language changed to: $newLang (Google API: $script:CurrentGoogleLang)" "INFO"
    }
})

$cmbSettingsLanguage.Add_SelectionChanged({
    if ($script:SuppressLangEvents) { return }
    $sel = $cmbSettingsLanguage.SelectedItem
    if ($sel -and $sel.Tag) {
        $newLang = [string]$sel.Tag
        Apply-AppLanguage -LanguageCode $newLang
        Save-AppConfig -ApiKey (Get-CurrentApiKey) -RememberApiKey $chkRememberKey.IsChecked `
            -OutputFolder $txtSettingsOutputDir.Text.Trim() `
            -LastInputFolder $script:Config.LastInputFolder -LastInputPath $script:Config.LastInputPath `
            -DefaultRouteType ($cmbDefaultRouteType.SelectedItem.Tag -as [string]) `
            -DefaultEmission ($cmbDefaultEmission.SelectedItem.Tag -as [string]) `
            -Language $newLang
        Write-AppLog "Language changed to: $newLang (Google API: $script:CurrentGoogleLang)" "INFO"
    }
})

$btnOpenLangFile.Add_Click({
    if (-not (Test-Path $script:LocalizationFile)) {
        Load-LocalizationConfig
    }
    try {
        Start-Process -FilePath "notepad.exe" -ArgumentList "`"$script:LocalizationFile`""
    } catch {
        try { Start-Process -FilePath $script:LocalizationFile } catch {
            [System.Windows.MessageBox]::Show("Cannot open localization file:`r`n$($script:LocalizationFile)`r`n$($_.Exception.Message)", 'Error', 'OK', 'Error')
        }
    }
})

$btnReloadLang.Add_Click({
    Load-LocalizationConfig
    Populate-LanguageDropdowns
    Apply-AppLanguage -LanguageCode $script:CurrentLanguage
    $count = $script:LanguagesCatalog.Count
    $msg = (Get-LocText 'MsgLangReloaded') -f $count
    $title = (Get-LocText 'MsgLangReloadedTitle')
    [System.Windows.MessageBox]::Show($msg, $title, 'OK', 'Information')
})

$btnToggleKeyVisibility.Add_Click({
    if ($txtSettingsApiKeyVisible.Visibility -eq [System.Windows.Visibility]::Visible) {
        $txtSettingsApiKey.Password = $txtSettingsApiKeyVisible.Text
        $txtSettingsApiKeyVisible.Visibility = [System.Windows.Visibility]::Collapsed
        $txtSettingsApiKey.Visibility = [System.Windows.Visibility]::Visible
        $btnToggleKeyVisibility.Content = '👁 Show'
    } else {
        $txtSettingsApiKeyVisible.Text = $txtSettingsApiKey.Password
        $txtSettingsApiKey.Visibility = [System.Windows.Visibility]::Collapsed
        $txtSettingsApiKeyVisible.Visibility = [System.Windows.Visibility]::Visible
        $btnToggleKeyVisibility.Content = '🔒 Hide'
    }
})

$btnOpenLogFile.Add_Click({
    if (-not (Test-Path $script:LogFile)) {
        Write-AppLog "Creating new application log file." "INFO"
    }
    try {
        Start-Process -FilePath "notepad.exe" -ArgumentList "`"$script:LogFile`""
    } catch {
        try { Start-Process -FilePath $script:LogFile } catch {
            [System.Windows.MessageBox]::Show("Cannot open log file:`r`n$($script:LogFile)`r`n$($_.Exception.Message)", 'Log Open Error', 'OK', 'Error')
        }
    }
})

$btnTestApiKey.Add_Click({
    $key = Get-CurrentApiKey
    if ([string]::IsNullOrWhiteSpace($key)) {
        [System.Windows.MessageBox]::Show((Get-LocText 'MsgMissingApiKey'), (Get-LocText 'MsgMissingApiKeyTitle'), 'OK', 'Warning')
        return
    }

    $maskedKey = Get-MaskedKey $key
    Write-AppLog "Testing API key (Key: $maskedKey)..." "INFO"

    $lblKeyTestResult.Text = 'Testing API key...'
    $lblKeyTestResult.Foreground = [System.Windows.Media.Brushes]::SkyBlue
    $btnTestApiKey.IsEnabled = $false

    # Asynchroniczny test połączenia z Google API w osobnym runspace
    $testScript = {
        param($apiKeyToTest, $apiLanguage = 'en')
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls
        if ([string]::IsNullOrWhiteSpace($apiKeyToTest)) {
            return [PSCustomObject]@{ Valid = $false; Message = 'API key is empty.' }
        }
        try {
            $lang = if ($apiLanguage) { $apiLanguage } else { 'en' }
            $Url = "https://maps.googleapis.com/maps/api/geocode/json?address=Warszawa&language=$lang&key=$apiKeyToTest"
            $Resp = Invoke-RestMethod -Uri $Url -Method GET -TimeoutSec 15
            if ($Resp.status -eq 'OK' -or $Resp.status -eq 'ZERO_RESULTS') {
                return [PSCustomObject]@{ Valid = $true; Message = 'Google Maps API key is valid and active.' }
            }
            elseif ($Resp.status -eq 'REQUEST_DENIED') {
                $msg = if ($Resp.error_message) { $Resp.error_message } else { 'Request denied by Google API.' }
                return [PSCustomObject]@{ Valid = $false; Message = "Unauthorized: $msg" }
            }
            else {
                return [PSCustomObject]@{ Valid = $false; Message = "API Status: $($Resp.status)" }
            }
        }
        catch {
            return [PSCustomObject]@{ Valid = $false; Message = "Connection error: $($_.Exception.Message)" }
        }
    }

    $psTest = [PowerShell]::Create().AddScript($testScript).AddArgument($key).AddArgument($script:CurrentGoogleLang)
    $testHandle = $psTest.BeginInvoke()
    $testTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $testTimer.Interval = [TimeSpan]::FromMilliseconds(100)
    $script:ActiveTestTimer = $testTimer
    $script:ActiveTestPs = $psTest
    $script:ActiveTestHandle = $testHandle
    $script:TestTimerTicks = 0

    $testTimer.Add_Tick({
        $localTestHandle = $script:ActiveTestHandle
        $localTestPs     = $script:ActiveTestPs
        $script:TestTimerTicks++
        if ($localTestHandle -and $localTestHandle.IsCompleted) {
            if ($script:ActiveTestTimer) { try { $script:ActiveTestTimer.Stop() } catch { } }
            $btnTestApiKey.IsEnabled = $true
            try {
                $res = $localTestPs.EndInvoke($localTestHandle)
                $testResult = $res[0]
                $isValid = [bool]$testResult.Valid
                $msg = [string]$testResult.Message
                Update-ApiStatusBadge -IsValid $isValid -Message $msg
                $logLevel = if ($isValid) { 'OK' } else { 'WARN' }
                Write-AppLog "API key test completed: Valid=$isValid, Message='$msg'" $logLevel
            }
            catch {
                $errDetail = $_.Exception.Message
                Update-ApiStatusBadge -IsValid $false -Message "Test error: $errDetail"
                Write-AppLog "Exception while retrieving API test result: $($_.Exception.ToString())" "ERROR"
            }
            finally {
                $localTestPs.Dispose()
            }
        }
        elseif ($script:TestTimerTicks -ge 200) { # 20s timeout limit
            if ($script:ActiveTestTimer) { try { $script:ActiveTestTimer.Stop() } catch { } }
            $btnTestApiKey.IsEnabled = $true
            Update-ApiStatusBadge -IsValid $false -Message "Google API response timeout (20s)."
            Write-AppLog "API key test timeout (20s watchdog timeout)." "WARN"
            try { $localTestPs.Stop(); $localTestPs.Dispose() } catch { }
        }
    })
    $testTimer.Start()
})

$btnBrowseSettingsOutputDir.Add_Click({
    $dlg = [System.Windows.Forms.FolderBrowserDialog]::new()
    $dlg.Description = 'Select default folder for calculation results'
    $dlg.SelectedPath = $txtSettingsOutputDir.Text
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtSettingsOutputDir.Text = $dlg.SelectedPath
    }
})

$btnSaveSettings.Add_Click({
    $key = Get-CurrentApiKey
    $remember = [bool]$chkRememberKey.IsChecked
    $outDir = $txtSettingsOutputDir.Text.Trim()
    $routeType = ($cmbDefaultRouteType.SelectedItem.Tag -as [string])
    $emission = ($cmbDefaultEmission.SelectedItem.Tag -as [string])
    $dims = ($cmbDefaultMapSize.SelectedItem.Tag -as [string]) -split 'x'
    $mapW = [int]$dims[0]
    $mapH = [int]$dims[1]

    $langToSave = if ($cmbSettingsLanguage.SelectedItem) { [string]$cmbSettingsLanguage.SelectedItem.Tag } else { $script:CurrentLanguage }
    $themeToSave = if ($cmbSettingsTheme.SelectedItem) { [string]$cmbSettingsTheme.SelectedItem.Tag } else { $script:CurrentTheme }
    $overlayCfg = Get-CurrentOverlayConfig
    Save-AppConfig -ApiKey $key -RememberApiKey $remember -OutputFolder $outDir `
        -LastInputFolder $script:Config.LastInputFolder -LastInputPath $script:Config.LastInputPath `
        -DefaultRouteType $routeType -DefaultEmission $emission -MapWidth $mapW -MapHeight $mapH -Language $langToSave `
        -OverlayConfig $overlayCfg -Theme $themeToSave
    $script:Config.OverlayConfig = $overlayCfg
    $script:Config.Theme = $themeToSave

    Set-CurrentApiKey -Key $key
    if ($remember -and -not [string]::IsNullOrWhiteSpace($key)) {
        $lblApiBadge.Text = 'API: Configured'
        $lblApiBadge.Foreground = [System.Windows.Media.Brushes]::LightGreen
    }
    elseif (-not $remember -and -not [string]::IsNullOrWhiteSpace($key)) {
        $lblApiBadge.Text = 'API: Session key (unsaved)'
        $lblApiBadge.Foreground = [System.Windows.Media.Brushes]::SkyBlue
    }
    else {
        $lblApiBadge.Text = 'API: Unverified'
        $lblApiBadge.Foreground = [System.Windows.Media.Brushes]::Orange
    }

    [System.Windows.MessageBox]::Show((Get-LocText 'MsgSettingsSaved'), (Get-LocText 'MsgSettingsSavedTitle'), 'OK', 'Information')
})

if ($btnResetOverlayConfig) {
    $btnResetOverlayConfig.Add_Click({
        Reset-OverlayConfigUi
    })
}

# ── 12. Zdarzenia: Tab 1 (Manual Input) ───────────────────────────────────────
$btnClearManualStart.Add_Click({ $txtManualStart.Clear() })
$btnClearManualEnd.Add_Click({ $txtManualEnd.Clear() })

$btnAddWaypoint.Add_Click({
    $wp = $txtNewWaypoint.Text.Trim()
    if (-not [string]::IsNullOrWhiteSpace($wp)) {
        if ($lstWaypoints.Items.Count -ge 25) {
            [System.Windows.MessageBox]::Show((Get-LocText 'MsgMaxWaypoints'), (Get-LocText 'MsgMaxWaypointsTitle'), 'OK', 'Warning')
            return
        }
        $null = $lstWaypoints.Items.Add($wp)
        $txtNewWaypoint.Clear()
    }
})

$txtNewWaypoint.Add_KeyDown({
    if ($_.Key -eq [System.Windows.Input.Key]::Enter) {
        $btnAddWaypoint.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
    }
})

$btnWpRemove.Add_Click({
    if ($lstWaypoints.SelectedIndex -ge 0) {
        $lstWaypoints.Items.RemoveAt($lstWaypoints.SelectedIndex)
    }
})

$btnWpClear.Add_Click({
    $lstWaypoints.Items.Clear()
})

$btnWpUp.Add_Click({
    $idx = $lstWaypoints.SelectedIndex
    if ($idx -gt 0) {
        $item = $lstWaypoints.Items[$idx]
        $lstWaypoints.Items.RemoveAt($idx)
        $lstWaypoints.Items.Insert($idx - 1, $item)
        $lstWaypoints.SelectedIndex = $idx - 1
    }
})

$btnWpDown.Add_Click({
    $idx = $lstWaypoints.SelectedIndex
    if ($idx -ge 0 -and $idx -lt ($lstWaypoints.Items.Count - 1)) {
        $item = $lstWaypoints.Items[$idx]
        $lstWaypoints.Items.RemoveAt($idx)
        $lstWaypoints.Items.Insert($idx + 1, $item)
        $lstWaypoints.SelectedIndex = $idx + 1
    }
})

$rbTypeEco.Add_Checked({ $pnlEmission.Visibility = [System.Windows.Visibility]::Visible })
$rbTypeFastest.Add_Checked({ $pnlEmission.Visibility = [System.Windows.Visibility]::Collapsed })
$rbTypeShortest.Add_Checked({ $pnlEmission.Visibility = [System.Windows.Visibility]::Collapsed })

$btnCalculateManual.Add_Click({
    $apiKey = Get-CurrentApiKey
    if ([string]::IsNullOrWhiteSpace($apiKey)) {
        [System.Windows.MessageBox]::Show((Get-LocText 'MsgMissingApiKeyPrompt'), (Get-LocText 'MsgMissingApiKeyTitle'), 'OK', 'Warning')
        $tabMain.SelectedIndex = 2
        return
    }

    $start = $txtManualStart.Text.Trim()
    $end = $txtManualEnd.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($start) -or [string]::IsNullOrWhiteSpace($end)) {
        [System.Windows.MessageBox]::Show((Get-LocText 'MsgMissingData'), (Get-LocText 'MsgMissingDataTitle'), 'OK', 'Warning')
        return
    }

    $waypoints = [System.Collections.Generic.List[string]]::new()
    foreach ($item in $lstWaypoints.Items) {
        $waypoints.Add([string]$item)
    }

    $routeType = if ($rbTypeShortest.IsChecked) { 'Shortest' } elseif ($rbTypeEco.IsChecked) { 'Eco' } else { 'Fastest' }
    $emission = ($cmbEmission.SelectedItem.Tag -as [string])
    if ([string]::IsNullOrWhiteSpace($emission)) { $emission = 'GASOLINE' }
    $trafficAware = [bool]$chkTrafficAware.IsChecked
    $name = $txtManualName.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($name)) { $name = "Route $start -> $end" }

    $outDir = $txtSettingsOutputDir.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($outDir)) { $outDir = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'GoogleMapsRoutes\Results' }
    if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

    if ($txtBatchFilePath.Text -and (Test-Path $txtBatchFilePath.Text.Trim())) {
        $script:Config.LastInputPath = $txtBatchFilePath.Text.Trim()
        $script:Config.LastInputFolder = Split-Path $script:Config.LastInputPath -Parent
        $script:LastDataDirectory = $script:Config.LastInputFolder
        Save-AppConfig -ApiKey $apiKey -RememberApiKey $chkRememberKey.IsChecked -OutputFolder $outDir `
            -LastInputFolder $script:Config.LastInputFolder -LastInputPath $script:Config.LastInputPath `
            -DefaultRouteType ($cmbDefaultRouteType.SelectedItem.Tag -as [string]) `
            -DefaultEmission ($cmbDefaultEmission.SelectedItem.Tag -as [string])
    }

    $btnCalculateManual.IsEnabled = $false
    $btnCalculateManual.Content = '⏳ CALCULATING ROUTE...'
    $lblManualStatus.Text = 'Geocoding and calculating...'
    $lblManualStatus.Foreground = [System.Windows.Media.Brushes]::SkyBlue
    $lblFooterStatus.Text = 'Calculating manual route...'

    Write-AppLog "Started manual route calculation: '$start' -> '$end' (Waypoints: $($waypoints.Count), Type: $routeType, Engine: $emission, LiveTraffic: $trafficAware)..." "INFO"

    # $script:ManualCalcAsync is defined at top-level (section 6b) — used directly below
    $psCmd = New-WorkerPowerShell -ScriptBlock $script:ManualCalcAsync
    $overlayCfgJson = (Get-CurrentOverlayConfig | ConvertTo-Json -Depth 6 -Compress)
    $psCmd.AddArgument($start).AddArgument($end).AddArgument($waypoints).AddArgument($routeType).AddArgument($emission).AddArgument($trafficAware).AddArgument($name).AddArgument($apiKey).AddArgument($outDir).AddArgument($script:LogFile).AddArgument($script:CurrentGoogleLang).AddArgument($overlayCfgJson) | Out-Null

    try {
        $asyncHandle = $psCmd.BeginInvoke()
    } catch {
        Write-AppLog "CRITICAL: BeginInvoke() threw exception: $($_.Exception.Message)" "ERROR"
        $btnCalculateManual.IsEnabled = $true
        $btnCalculateManual.Content = '🚀 CALCULATE ROUTE & DOWNLOAD MAP'
        $lblManualStatus.Text = '✕ Launch error'
        $lblManualStatus.Foreground = [System.Windows.Media.Brushes]::Salmon
        $lblFooterStatus.Text = "Error: $($_.Exception.Message)"
        return
    }
    if (-not $asyncHandle) {
        Write-AppLog "CRITICAL: BeginInvoke() returned null — runspace may be invalid." "ERROR"
        $btnCalculateManual.IsEnabled = $true
        $btnCalculateManual.Content = '🚀 CALCULATE ROUTE & DOWNLOAD MAP'
        $lblManualStatus.Text = '✕ Runspace error'
        $lblManualStatus.Foreground = [System.Windows.Media.Brushes]::Salmon
        return
    }
    $timer = [System.Windows.Threading.DispatcherTimer]::new()
    $timer.Interval = [TimeSpan]::FromMilliseconds(150)
    $script:ActiveManualTimer = $timer
    $script:ActiveManualPs = $psCmd
    $script:ActiveManualAsyncHandle = $asyncHandle
    $script:ManualTimerTicks = 0

    Write-AppLog "Worker started (BeginInvoke). IsCompleted=$($asyncHandle.IsCompleted)" "INFO"

    $timer.Add_Tick({
        $localHandle = $script:ActiveManualAsyncHandle
        $localPs     = $script:ActiveManualPs
        $script:ManualTimerTicks++
        if ($localHandle -and $localHandle.IsCompleted) {
            if ($script:ActiveManualTimer) { try { $script:ActiveManualTimer.Stop() } catch { } }
            $btnCalculateManual.IsEnabled = $true
            $btnCalculateManual.Content = '🚀 CALCULATE ROUTE & DOWNLOAD MAP'

            # Log any stream errors from the worker runspace before inspecting result
            foreach ($streamErr in $localPs.Streams.Error) {
                Write-AppLog "[Stream.Error] $($streamErr.Exception.Message) @ $($streamErr.InvocationInfo.PositionMessage)" "ERROR"
            }

            try {
                $res = $localPs.EndInvoke($localHandle)
                $calc = $res[0]
                if ($calc.Success) {
                    $lblManualDist.Text = "$($calc.DistanceKm) km"
                    $lblManualTime.Text = "$($calc.DurationMin) min"
                    $lblManualType.Text = switch ($script:CurrentLanguage) {
                        'de' { if ($calc.RouteType -eq 'Fastest') { 'Schnellste' } elseif ($calc.RouteType -eq 'Shortest') { 'Kürzeste' } else { 'Eco' } }
                        'pl' { if ($calc.RouteType -eq 'Fastest') { 'Najszybsza' } elseif ($calc.RouteType -eq 'Shortest') { 'Najkrótsza' } else { 'Eko' } }
                        default { [string]$calc.RouteType }
                    }
                    $lblManualStatus.Text = '✓ Success'
                    $lblManualStatus.Foreground = [System.Windows.Media.Brushes]::LightGreen
                    $lblFooterStatus.Text = "Route ready: $($calc.DistanceKm) km, $($calc.DurationMin) min"
                    Write-AppLog "Manual route calculation completed successfully: $($calc.DistanceKm) km, $($calc.DurationMin) min (Map file: $($calc.MapPath))" "OK"

                    $script:LastGoogleMapsUrl = $calc.GoogleMapsUrl
                    $lblGoogleUrlDisplay.Text = $calc.GoogleMapsUrl
                    $btnOpenGoogleMaps.IsEnabled = $true
                    $btnCopyUrl.IsEnabled = $true

                    if ($calc.MapPath -and (Test-Path $calc.MapPath)) {
                        $script:LastGeneratedMapPath = $calc.MapPath
                        $btnSaveMapAs.IsEnabled = $true
                        $lblMapPlaceholder.Visibility = [System.Windows.Visibility]::Collapsed

                        $imgBytes = [System.IO.File]::ReadAllBytes($calc.MapPath)
                        $ms = [System.IO.MemoryStream]::new($imgBytes)
                        $bi = [System.Windows.Media.Imaging.BitmapImage]::new()
                        $bi.BeginInit()
                        $bi.StreamSource = $ms
                        $bi.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
                        $bi.EndInit()
                        $bi.Freeze()
                        $imgMapPreview.Source = $bi
                    }
                } else {
                    $lblManualStatus.Text = '✕ Error'
                    $lblManualStatus.Foreground = [System.Windows.Media.Brushes]::Salmon
                    $lblFooterStatus.Text = "Error: $($calc.Error)"
                    Write-AppLog "Manual route calculation failed: $($calc.Error)" "ERROR"
                    [System.Windows.MessageBox]::Show($calc.Error, 'Route Error', 'OK', 'Error')
                }
            }
            catch {
                $errDetail = $_.Exception.ToString()
                $lblManualStatus.Text = '✕ Error'
                $lblManualStatus.Foreground = [System.Windows.Media.Brushes]::Salmon
                $lblFooterStatus.Text = "Exception: $($_.Exception.Message)"
                Write-AppLog "UI exception while reading route result: $errDetail" "ERROR"
                [System.Windows.MessageBox]::Show($_.Exception.Message, 'Error', 'OK', 'Error')
            }
            finally {
                $localPs.Dispose()
            }
        }
        elseif ($script:ManualTimerTicks -ge 400) { # 60 seconds watchdog timeout
            if ($script:ActiveManualTimer) { try { $script:ActiveManualTimer.Stop() } catch { } }
            $btnCalculateManual.IsEnabled = $true
            $btnCalculateManual.Content = '🚀 CALCULATE ROUTE & DOWNLOAD MAP'
            $lblManualStatus.Text = '✕ Timeout (60s)'
            $lblManualStatus.Foreground = [System.Windows.Media.Brushes]::Salmon
            $lblFooterStatus.Text = 'Route calculation timed out (60s).'
            Write-AppLog "Manual route calculation timed out (60s watchdog timeout). Ticks=$($script:ManualTimerTicks)" "WARN"
            try { $localPs.Stop(); $localPs.Dispose() } catch { }
        }
    })
    $timer.Start()
})

$btnOpenGoogleMaps.Add_Click({
    if ($script:LastGoogleMapsUrl) {
        Start-Process $script:LastGoogleMapsUrl
    }
})

$btnCopyUrl.Add_Click({
    if ($script:LastGoogleMapsUrl) {
        [System.Windows.Clipboard]::SetText($script:LastGoogleMapsUrl)
        [System.Windows.MessageBox]::Show((Get-LocText 'MsgUrlCopied'), (Get-LocText 'MsgUrlCopiedTitle'), 'OK', 'Information')
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
            [System.Windows.MessageBox]::Show(((Get-LocText 'MsgMapSaved') -f $dlg.FileName), (Get-LocText 'MsgMapSavedTitle'), 'OK', 'Information')
        }
    }
})

# ── 13. Zdarzenia: Tab 2 (Batch Processing) ──────────────────────────────────
function Load-BatchFilePreview([string]$Path) {
    if (-not (Test-Path $Path)) { return }
    try {
        $data = Import-RouteDataFile -Path $Path
        $script:LoadedBatchData = $data

        # Configure dynamic columns and clear previous items
        $dgBatchInput.ItemsSource = $null
        $dgBatchInput.Columns.Clear()

        if ($data.Mode -eq 'SequentialStops') {
            $lblBatchFileInfo.Text = "Format: $($data.Format) | Mode: Sequential Stops (1 Multi-point Route, $($data.TotalCount) Stops) | Total: $($data.TotalCount) stops"
            $lblBatchFileInfo.Foreground = [System.Windows.Media.Brushes]::LightGreen

            $colStep = [System.Windows.Controls.DataGridTextColumn]::new()
            $colStep.Header = "#"
            $colStep.Binding = [System.Windows.Data.Binding]::new("Step")
            $colStep.Width = [System.Windows.Controls.DataGridLength]::new(55)
            $dgBatchInput.Columns.Add($colStep)

            $colRole = [System.Windows.Controls.DataGridTextColumn]::new()
            $colRole.Header = "Role / Point Type"
            $colRole.Binding = [System.Windows.Data.Binding]::new("Role")
            $colRole.Width = [System.Windows.Controls.DataGridLength]::new(180)
            $dgBatchInput.Columns.Add($colRole)

            $colAddr = [System.Windows.Controls.DataGridTextColumn]::new()
            $colAddr.Header = "Address / Location"
            $colAddr.Binding = [System.Windows.Data.Binding]::new("Address")
            $colAddr.Width = [System.Windows.Controls.DataGridLength]::new(340)
            $dgBatchInput.Columns.Add($colAddr)

            $colRaw = [System.Windows.Controls.DataGridTextColumn]::new()
            $colRaw.Header = "Source Record Data"
            $colRaw.Binding = [System.Windows.Data.Binding]::new("RawSummary")
            $colRaw.Width = [System.Windows.Controls.DataGridLength]::new(1, [System.Windows.Controls.DataGridLengthUnitType]::Star)
            $dgBatchInput.Columns.Add($colRaw)

            $previewItems = [System.Collections.Generic.List[PSCustomObject]]::new()
            $stopsCount = $data.Stops.Count
            for ($i = 0; $i -lt $stopsCount; $i++) {
                $st = $data.Stops[$i]
                $role = if ($i -eq 0) { "🟢 Origin (Start)" }
                        elseif ($i -eq ($stopsCount - 1)) { "🔴 Destination (End)" }
                        else { "🟡 Waypoint $i" }

                $rawProps = @()
                if ($st.Raw) {
                    foreach ($p in $st.Raw.PSObject.Properties) {
                        $rawProps += "$($p.Name)=$($p.Value)"
                    }
                }
                $rawSummaryText = $rawProps -join '; '

                $previewItems.Add([PSCustomObject]@{
                    Step       = ($i + 1)
                    Role       = $role
                    Address    = [string]$st.Address
                    RawSummary = $rawSummaryText
                })
            }
            $dgBatchInput.ItemsSource = $previewItems
            Write-BatchLog "Loaded file: $Path ($($data.TotalCount) sequential stops, format: $($data.Format), mode: $($data.Mode))" "OK"
        }
        else {
            $lblBatchFileInfo.Text = "Format: $($data.Format) | Mode: Route List | Total Routes: $($data.TotalCount)"
            $lblBatchFileInfo.Foreground = [System.Windows.Media.Brushes]::LightGreen

            $colId = [System.Windows.Controls.DataGridTextColumn]::new()
            $colId.Header = "ID"
            $colId.Binding = [System.Windows.Data.Binding]::new("Id")
            $colId.Width = [System.Windows.Controls.DataGridLength]::new(50)
            $dgBatchInput.Columns.Add($colId)

            $colName = [System.Windows.Controls.DataGridTextColumn]::new()
            $colName.Header = "Route Name"
            $colName.Binding = [System.Windows.Data.Binding]::new("Name")
            $colName.Width = [System.Windows.Controls.DataGridLength]::new(180)
            $dgBatchInput.Columns.Add($colName)

            $colStart = [System.Windows.Controls.DataGridTextColumn]::new()
            $colStart.Header = "Origin (Start)"
            $colStart.Binding = [System.Windows.Data.Binding]::new("Start")
            $colStart.Width = [System.Windows.Controls.DataGridLength]::new(200)
            $dgBatchInput.Columns.Add($colStart)

            $colEnd = [System.Windows.Controls.DataGridTextColumn]::new()
            $colEnd.Header = "Destination (End)"
            $colEnd.Binding = [System.Windows.Data.Binding]::new("End")
            $colEnd.Width = [System.Windows.Controls.DataGridLength]::new(200)
            $dgBatchInput.Columns.Add($colEnd)

            $colWpCount = [System.Windows.Controls.DataGridTextColumn]::new()
            $colWpCount.Header = "Waypoints"
            $colWpCount.Binding = [System.Windows.Data.Binding]::new("WaypointCount")
            $colWpCount.Width = [System.Windows.Controls.DataGridLength]::new(80)
            $dgBatchInput.Columns.Add($colWpCount)

            $colWpText = [System.Windows.Controls.DataGridTextColumn]::new()
            $colWpText.Header = "Intermediate Stops"
            $colWpText.Binding = [System.Windows.Data.Binding]::new("WaypointsText")
            $colWpText.Width = [System.Windows.Controls.DataGridLength]::new(250)
            $dgBatchInput.Columns.Add($colWpText)

            $colType = [System.Windows.Controls.DataGridTextColumn]::new()
            $colType.Header = "Route Type"
            $colType.Binding = [System.Windows.Data.Binding]::new("RouteType")
            $colType.Width = [System.Windows.Controls.DataGridLength]::new(100)
            $dgBatchInput.Columns.Add($colType)

            $previewItems = [System.Collections.Generic.List[PSCustomObject]]::new()
            foreach ($r in @($data.Routes)) {
                $wpText = if ($r.Waypoints -and @($r.Waypoints).Count -gt 0) {
                    (@($r.Waypoints) -join ' | ')
                } else {
                    '(none)'
                }
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
            Write-BatchLog "Loaded file: $Path ($($data.TotalCount) routes, format: $($data.Format), mode: $($data.Mode))" "OK"
        }
    }
    catch {
        $lblBatchFileInfo.Text = "Load error: $($_.Exception.Message)"
        $lblBatchFileInfo.Foreground = [System.Windows.Media.Brushes]::Salmon
        Write-BatchLog "Load error: $($_.Exception.Message)" "ERROR"
    }
}

$btnBrowseBatchFile.Add_Click({
    $initDir = $null
    if (-not [string]::IsNullOrWhiteSpace($txtBatchFilePath.Text) -and (Test-Path $txtBatchFilePath.Text.Trim())) {
        $initDir = Split-Path $txtBatchFilePath.Text.Trim() -Parent
    }
    elseif (-not [string]::IsNullOrWhiteSpace($txtBatchFilePath.Text) -and (Test-Path (Split-Path $txtBatchFilePath.Text.Trim() -Parent))) {
        $initDir = Split-Path $txtBatchFilePath.Text.Trim() -Parent
    }
    elseif ($script:Config.LastInputFolder -and (Test-Path $script:Config.LastInputFolder)) {
        $initDir = $script:Config.LastInputFolder
    }
    elseif ($script:Config.LastInputPath -and (Test-Path (Split-Path $script:Config.LastInputPath -Parent))) {
        $initDir = Split-Path $script:Config.LastInputPath -Parent
    }

    $file = Select-InputDataFile -InitialDirectory $initDir
    if ($file) {
        $txtBatchFilePath.Text = $file
        $script:Config.LastInputPath = $file
        $script:Config.LastInputFolder = Split-Path $file -Parent
        $script:LastDataDirectory = $script:Config.LastInputFolder

        Save-AppConfig -ApiKey (Get-CurrentApiKey) `
            -RememberApiKey $chkRememberKey.IsChecked `
            -OutputFolder $txtSettingsOutputDir.Text.Trim() `
            -LastInputFolder $script:Config.LastInputFolder `
            -LastInputPath $script:Config.LastInputPath `
            -DefaultRouteType ($cmbDefaultRouteType.SelectedItem.Tag -as [string]) `
            -DefaultEmission ($cmbDefaultEmission.SelectedItem.Tag -as [string])

        Load-BatchFilePreview -Path $file
    }
})

$btnReloadBatchFile.Add_Click({
    if ($txtBatchFilePath.Text) {
        $path = $txtBatchFilePath.Text.Trim()
        if (Test-Path $path) {
            $script:Config.LastInputPath = $path
            $script:Config.LastInputFolder = Split-Path $path -Parent
            $script:LastDataDirectory = $script:Config.LastInputFolder

            Save-AppConfig -ApiKey (Get-CurrentApiKey) `
                -RememberApiKey $chkRememberKey.IsChecked `
                -OutputFolder $txtSettingsOutputDir.Text.Trim() `
                -LastInputFolder $script:Config.LastInputFolder `
                -LastInputPath $script:Config.LastInputPath `
                -DefaultRouteType ($cmbDefaultRouteType.SelectedItem.Tag -as [string]) `
                -DefaultEmission ($cmbDefaultEmission.SelectedItem.Tag -as [string])
        }
        Load-BatchFilePreview -Path $path
    }
})

$btnStartBatch.Add_Click({
    $apiKey = Get-CurrentApiKey
    if ([string]::IsNullOrWhiteSpace($apiKey)) {
        [System.Windows.MessageBox]::Show((Get-LocText 'MsgMissingApiKeyPrompt'), (Get-LocText 'MsgMissingApiKeyTitle'), 'OK', 'Warning')
        $tabMain.SelectedIndex = 2
        return
    }

    if ($null -eq $script:LoadedBatchData -or $script:LoadedBatchData.Routes.Count -eq 0) {
        [System.Windows.MessageBox]::Show((Get-LocText 'MsgNoDataFile'), (Get-LocText 'MsgNoDataFileTitle'), 'OK', 'Warning')
        return
    }

    $outDir = $txtSettingsOutputDir.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($outDir)) { $outDir = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'GoogleMapsRoutes\Results' }
    if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

    if ($txtBatchFilePath.Text -and (Test-Path $txtBatchFilePath.Text.Trim())) {
        $script:Config.LastInputPath = $txtBatchFilePath.Text.Trim()
        $script:Config.LastInputFolder = Split-Path $script:Config.LastInputPath -Parent
        $script:LastDataDirectory = $script:Config.LastInputFolder
        Save-AppConfig -ApiKey $apiKey -RememberApiKey $chkRememberKey.IsChecked -OutputFolder $outDir `
            -LastInputFolder $script:Config.LastInputFolder -LastInputPath $script:Config.LastInputPath `
            -DefaultRouteType ($cmbDefaultRouteType.SelectedItem.Tag -as [string]) `
            -DefaultEmission ($cmbDefaultEmission.SelectedItem.Tag -as [string])
    }

    $defaultRouteType = ($cmbBatchRouteType.SelectedItem.Tag -as [string])
    if ([string]::IsNullOrWhiteSpace($defaultRouteType)) { $defaultRouteType = 'Fastest' }

    $script:BatchWorkerRunning = $true
    $btnStartBatch.IsEnabled = $false
    $btnStopBatch.IsEnabled = $true
    $btnBrowseBatchFile.IsEnabled = $false
    $btnReloadBatchFile.IsEnabled = $false

    $script:BatchResultsList.Clear()
    $dgBatchResults.ItemsSource = $null
    if ($dgBatchPoints) { $dgBatchPoints.ItemsSource = $null }
    $pbBatchProgress.Value = 0
    $lblBatchProgressText.Text = "Starting batch processing (0 / $($script:LoadedBatchData.Routes.Count))..."
    $lblBatchStats.Text = "Success: 0 | Errors: 0"

    Write-BatchLog "=== Starting batch processing ($($script:LoadedBatchData.Routes.Count) routes) ===" "INFO"

    # Switch to Activity Log tab so user sees live execution
    if ($tabBatchSub) {
        $tabBatchSub.SelectedIndex = 2
    }

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
        $overlayCfgJson = (Get-CurrentOverlayConfig | ConvertTo-Json -Depth 6 -Compress)
        $psCmdBatch.AddArgument($routesToProcess).AddArgument($apiKey).AddArgument($outDir).AddArgument($defaultRouteType).AddArgument($syncState).AddArgument($script:LogFile).AddArgument($script:CurrentGoogleLang).AddArgument($overlayCfgJson) | Out-Null
        $asyncBatchHandle = $psCmdBatch.BeginInvoke()
    }
    catch {
        Write-BatchLog "CRITICAL: Could not start batch worker: $($_.Exception.Message)" "ERROR"
        $btnStartBatch.IsEnabled = $true
        $btnStopBatch.IsEnabled = $false
        $btnBrowseBatchFile.IsEnabled = $true
        $btnReloadBatchFile.IsEnabled = $true
        $script:BatchWorkerRunning = $false
        $lblBatchProgressText.Text = "Launch error"
        return
    }

    if (-not $asyncBatchHandle) {
        Write-BatchLog "CRITICAL: BeginInvoke returned null handle." "ERROR"
        $btnStartBatch.IsEnabled = $true
        $btnStopBatch.IsEnabled = $false
        $btnBrowseBatchFile.IsEnabled = $true
        $btnReloadBatchFile.IsEnabled = $true
        $script:BatchWorkerRunning = $false
        return
    }

    $timerBatch = [System.Windows.Threading.DispatcherTimer]::new()
    $timerBatch.Interval = [TimeSpan]::FromMilliseconds(150)
    $script:ActiveBatchTimer = $timerBatch
    $script:ActiveBatchPs = $psCmdBatch
    $script:ActiveBatchAsyncHandle = $asyncBatchHandle

    # Direct script-scope reference (NO .GetNewClosure()!)
    $timerBatch.Add_Tick({
        $localBatchHandle = $script:ActiveBatchAsyncHandle
        $localBatchPs     = $script:ActiveBatchPs
        $localSyncState   = $script:SyncState

        if (-not $localSyncState) { return }

        # Flush real-time worker logs to UI Activity Log
        if ($localSyncState.LogQueue) {
            $logItem = $null
            while ($localSyncState.LogQueue.TryDequeue([ref]$logItem)) {
                if ($logItem) {
                    Write-BatchLog $logItem.Message $logItem.Level
                }
            }
        }

        $curr = $localSyncState.CurrentIndex
        $tot  = $localSyncState.TotalCount
        $pct  = if ($tot -gt 0) { [math]::Min(100, [math]::Round(($curr / $tot) * 100, 0)) } else { 0 }
        $pbBatchProgress.Value = $pct
        $lblBatchProgressText.Text = "Processing: $curr / $tot ($pct%)"
        $lblBatchStats.Text = "Success: $($localSyncState.SuccessCount) | Errors: $($localSyncState.FailCount)"

        if ($localBatchHandle -and $localBatchHandle.IsCompleted) {
            $script:ActiveBatchTimer.Stop()
            $btnStartBatch.IsEnabled = $true
            $btnStopBatch.IsEnabled = $false
            $btnBrowseBatchFile.IsEnabled = $true
            $btnReloadBatchFile.IsEnabled = $true
            $script:BatchWorkerRunning = $false

            # Flush any remaining logs
            if ($localSyncState.LogQueue) {
                $logItem = $null
                while ($localSyncState.LogQueue.TryDequeue([ref]$logItem)) {
                    if ($logItem) {
                        Write-BatchLog $logItem.Message $logItem.Level
                    }
                }
            }

            # Check stream errors
            foreach ($streamErr in $localBatchPs.Streams.Error) {
                Write-BatchLog "[Worker Stream Error] $($streamErr.Exception.Message)" "ERROR"
            }

            try {
                $res = $localBatchPs.EndInvoke($localBatchHandle)
                $script:BatchResultsList = [System.Collections.Generic.List[PSCustomObject]]::new()
                foreach ($item in @($res)) { $script:BatchResultsList.Add($item) }
                $dgBatchResults.ItemsSource = $script:BatchResultsList

                # Populate Points Detail table
                $allPointsList = [System.Collections.Generic.List[PSCustomObject]]::new()
                foreach ($item in @($script:BatchResultsList)) {
                    if ($item.Points -and @($item.Points).Count -gt 0) {
                        foreach ($pt in @($item.Points)) {
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
                if ($dgBatchPoints) {
                    $dgBatchPoints.ItemsSource = $allPointsList
                }

                $statusMsg = if ($localSyncState.CancelRequested) { 'Stopped by user.' } else { 'Completed successfully.' }
                $lblBatchProgressText.Text = $statusMsg
                Write-BatchLog "=== Batch processing completed. $statusMsg Success: $($localSyncState.SuccessCount), Errors: $($localSyncState.FailCount) ===" "OK"
                $lblFooterStatus.Text = "Processing complete: $($localSyncState.SuccessCount) routes generated."

                # Automatically switch to Calculation Results tab
                if ($script:BatchResultsList.Count -gt 0 -and $tabBatchSub) {
                    $tabBatchSub.SelectedIndex = 1
                }
            }
            catch {
                Write-BatchLog "Error reading batch results: $($_.Exception.Message)" "ERROR"
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
        $lblBatchProgressText.Text = 'Stopping...'
        Write-BatchLog "Stop requested by user..." "WARN"
    }
})

$dgBatchResults.Add_MouseDoubleClick({
    $sel = $dgBatchResults.SelectedItem
    if ($sel -and $sel.MapPath -and (Test-Path $sel.MapPath)) {
        Start-Process $sel.MapPath
    }
})

$btnOpenOutputDir.Add_Click({
    $outDir = $txtSettingsOutputDir.Text.Trim()
    if (Test-Path $outDir) {
        Start-Process explorer.exe -ArgumentList "`"$outDir`""
    }
})

$btnExportExcel.Add_Click({
    if ($script:BatchResultsList.Count -eq 0) {
        [System.Windows.MessageBox]::Show((Get-LocText 'MsgNoExportResults'), (Get-LocText 'MsgNoExportResultsTitle'), 'OK', 'Information')
        return
    }
    $outDir = $txtSettingsOutputDir.Text.Trim()
    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $path = Join-Path $outDir "${ts}_route_results.xlsx"
    $saved = Export-RouteResults -Results $script:BatchResultsList -OutputPath $path -Format Excel
    [System.Windows.MessageBox]::Show(((Get-LocText 'MsgExportExcelComplete') -f $saved), (Get-LocText 'MsgExportTitle'), 'OK', 'Information')
})

$btnExportCsv.Add_Click({
    if ($script:BatchResultsList.Count -eq 0) {
        [System.Windows.MessageBox]::Show((Get-LocText 'MsgNoExportResults'), (Get-LocText 'MsgNoExportResultsTitle'), 'OK', 'Information')
        return
    }
    $outDir = $txtSettingsOutputDir.Text.Trim()
    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $path = Join-Path $outDir "${ts}_route_results.csv"
    $saved = Export-RouteResults -Results $script:BatchResultsList -OutputPath $path -Format CSV
    [System.Windows.MessageBox]::Show(((Get-LocText 'MsgExportCsvComplete') -f $saved), (Get-LocText 'MsgExportTitle'), 'OK', 'Information')
})

$btnExportJson.Add_Click({
    if ($script:BatchResultsList.Count -eq 0) {
        [System.Windows.MessageBox]::Show((Get-LocText 'MsgNoExportResults'), (Get-LocText 'MsgNoExportResultsTitle'), 'OK', 'Information')
        return
    }
    $outDir = $txtSettingsOutputDir.Text.Trim()
    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $path = Join-Path $outDir "${ts}_route_results.json"
    $saved = Export-RouteResults -Results $script:BatchResultsList -OutputPath $path -Format JSON
    [System.Windows.MessageBox]::Show(((Get-LocText 'MsgExportJsonComplete') -f $saved), (Get-LocText 'MsgExportTitle'), 'OK', 'Information')
})

# ── 13. Obsługa zamykania okna ───────────────────────────────────────────────
$window.Add_Closing({
    if ($script:ActiveBatchTimer) { try { $script:ActiveBatchTimer.Stop() } catch { } }
    if ($script:ActiveManualTimer) { try { $script:ActiveManualTimer.Stop() } catch { } }
    if ($script:SyncState) { $script:SyncState.CancelRequested = $true }
})

# ── 14. Uruchomienie okna ────────────────────────────────────────────────────
$window.ShowDialog() | Out-Null
