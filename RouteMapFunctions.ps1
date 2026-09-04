#Requires -Version 5.1
<#
.SYNOPSIS
    Wspólne funkcje do geokodowania adresów, obliczania tras (Fastest, Shortest, Eco)
    oraz generowania map PNG i obsługi plików danych (JSON, CSV, Excel).

.DESCRIPTION
    Moduł funkcji wykorzystywany przez:
      - GoogleMapsRoutes-GUI.ps1
      - Invoke-GoogleMapsRoute.ps1
      - Get-CarRoute_WithMap.ps1
      - Get-MultiPointCarRoute_WithMap.ps1
      - Process-SchoolTransportRoutes.ps1

    Główne funkcje:
      - Select-InputDataFile / Select-InputExcel : Dialogi wyboru plików
      - Protect-SecretString / Unprotect-SecretString : Bezpieczne przechowywanie klucza API (DPAPI)
      - Test-GoogleApiKey : Weryfikacja poprawności klucza API
      - Get-AddressCoordinates : Geokodowanie adresu przez Google Geocoding API
      - Get-CarRouteData : Obliczanie trasy przez Google Routes API v2 (Fastest, Shortest, Eco, multipoint)
      - Get-GoogleMapsUrl : Generowanie linku do Google Maps
      - Save-RouteMapPng : Pobieranie mapy z Google Static Maps API z wieloma znacznikami i nakładką
      - Import-RouteDataFile : Uniwersalny import tras z JSON, CSV, Excel (XLSX, XLS)
      - Export-RouteResults : Uniwersalny eksport wyników do Excel, CSV, JSON

.NOTES
    Wymagana zmienna środowiskowa GOOGLE_MAPS_API_KEY lub parametr -ApiKey.
    Encoding: UTF-8 with BOM
#>

# ══════════════════════════════════════════════════════════════════════════════
# 1. DIALOGI I BEZPIECZEŃSTWO (DPAPI)
# ══════════════════════════════════════════════════════════════════════════════

function Select-InputExcel {
    param([string]$InitialDirectory)
    Add-Type -AssemblyName System.Windows.Forms
    $Dialog = [System.Windows.Forms.OpenFileDialog]::new()
    $Dialog.Title = 'Select Excel file with addresses'
    $Dialog.Filter = 'Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*'
    if ($InitialDirectory -and (Test-Path $InitialDirectory)) {
        $Dialog.InitialDirectory = $InitialDirectory
    } else {
        $Dialog.InitialDirectory = [Environment]::GetFolderPath('MyDocuments')
    }
    $Dialog.RestoreDirectory = $true
    $Result = $Dialog.ShowDialog()
    if ($Result -eq [System.Windows.Forms.DialogResult]::OK) { return $Dialog.FileName }
    return $null
}

function Select-InputDataFile {
    param([string]$InitialDirectory)
    Add-Type -AssemblyName System.Windows.Forms
    $Dialog = [System.Windows.Forms.OpenFileDialog]::new()
    $Dialog.Title = 'Select route data file (JSON, CSV, Excel)'
    $Dialog.Filter = 'All Supported Files (*.xlsx;*.xls;*.csv;*.tsv;*.json)|*.xlsx;*.xls;*.csv;*.tsv;*.json|Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|CSV/TSV Files (*.csv;*.tsv)|*.csv;*.tsv|JSON Files (*.json)|*.json|All Files (*.*)|*.*'
    if ($InitialDirectory -and (Test-Path $InitialDirectory)) {
        $Dialog.InitialDirectory = $InitialDirectory
    } else {
        $Dialog.InitialDirectory = [Environment]::GetFolderPath('MyDocuments')
    }
    $Dialog.RestoreDirectory = $true
    $Result = $Dialog.ShowDialog()
    if ($Result -eq [System.Windows.Forms.DialogResult]::OK) { return $Dialog.FileName }
    return $null
}

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
        $sec = ConvertTo-SecureString -String $PlainText -AsPlainText -Force
        return (ConvertFrom-SecureString -SecureString $sec)
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
        return [PSCustomObject]@{ Valid = $false; Message = 'Klucz API jest pusty.' }
    }
    try {
        $lang = if ($LanguageCode) { ($LanguageCode -split '[-_]')[0].ToLower() } else { 'en' }
        $Url = "https://maps.googleapis.com/maps/api/geocode/json?address=Warszawa&language=$lang&key=$ApiKey"
        $Resp = Invoke-RestMethod -Uri $Url -Method GET -TimeoutSec 15
        if ($Resp.status -eq 'OK' -or $Resp.status -eq 'ZERO_RESULTS') {
            return [PSCustomObject]@{ Valid = $true; Message = 'Klucz Google Maps API jest poprawny i aktywny.' }
        }
        elseif ($Resp.status -eq 'REQUEST_DENIED') {
            $msg = if ($Resp.error_message) { $Resp.error_message } else { 'Żądanie odrzucone przez Google API.' }
            return [PSCustomObject]@{ Valid = $false; Message = "Brak autoryzacji: $msg" }
        }
        else {
            return [PSCustomObject]@{ Valid = $false; Message = "Status API: $($Resp.status)" }
        }
    }
    catch {
        return [PSCustomObject]@{ Valid = $false; Message = "Błąd połączenia: $($_.Exception.Message)" }
    }
}

# ══════════════════════════════════════════════════════════════════════════════
# 2. GEOKODOWANIE ADRESÓW (GOOGLE GEOCODING API)
# ══════════════════════════════════════════════════════════════════════════════

function Get-AddressComponentValue {
    [CmdletBinding()]
    param(
        [Parameter()][object[]]$Components,
        [Parameter(Mandatory)][string[]]$Types
    )

    $Matches = @($Components) | Where-Object {
        $_ -and $_.PSObject.Properties.Name -contains 'types' -and
        (@($_.types) | Where-Object { $_ -in $Types } | Select-Object -First 1)
    } | Select-Object -First 1

    if (-not $Matches) { return $null }

    foreach ($FieldName in @('long_name', 'short_name', 'name')) {
        if ($Matches.PSObject.Properties.Name -contains $FieldName) {
            $Value = [string]$Matches.$FieldName
            if (-not [string]::IsNullOrWhiteSpace($Value)) {
                return $Value
            }
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

    # Jeśli podano bezpośrednio współrzędne w formacie "52.2297, 21.0122"
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

    $EncodedAddress = [System.Uri]::EscapeDataString($Address.Trim())
    $lang = if ($LanguageCode) { ($LanguageCode -split '[-_]')[0].ToLower() } else { 'en' }
    $Url = "https://maps.googleapis.com/maps/api/geocode/json?address=$EncodedAddress&language=$lang&key=$ApiKey"
    try {
        Write-Verbose "Geokodowanie: '$Address'"
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
            Write-Warning "Geokodowanie nieudane dla '$Address'. Status API: $($Response.status)"
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
        Write-Warning "Błąd geokodowania '$Address': $Message"
        return [PSCustomObject]@{
            Latitude             = $null; Longitude = $null; FormattedAddress = $null
            UlicaINumer          = $null; KodPocztowy = $null; Miasto = $null
            MatchType            = $null; PartialMatch = $null
            Status               = "EXCEPTION: $Message"
            ErrorMessage         = $Message
        }
    }
}

# ══════════════════════════════════════════════════════════════════════════════
# 3. OBLICZANIE TRAS (GOOGLE ROUTES API v2)
# ══════════════════════════════════════════════════════════════════════════════

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

    $RoutesUrl = 'https://routes.googleapis.com/directions/v2:computeRoutes'

    $RequestBody = [ordered]@{
        origin       = @{ location = @{ latLng = @{ latitude = $OriginLat; longitude = $OriginLng } } }
        destination  = @{ location = @{ latLng = @{ latitude = $DestLat; longitude = $DestLng } } }
        travelMode   = 'DRIVE'
        languageCode = $LanguageCode
        units        = $Units
    }

    # Obsługa punktów pośrednich (Intermediates)
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

    # Konfiguracja parametrów w zależności od typu trasy ($RouteType)
    switch ($RouteType) {
        'Fastest' {
            # Google Routes API domyślnie optymalizuje czas trasy
            $RequestBody['routingPreference'] = if ($TrafficAware) { 'TRAFFIC_AWARE' } else { 'TRAFFIC_UNAWARE' }
            # Alternatywy są dozwolone tylko gdy nie ma punktów pośrednich
            if (-not $HasIntermediates) {
                $RequestBody['computeAlternativeRoutes'] = $true
            }
        }
        'Shortest' {
            # Pobieramy alternatywne trasy i wybieramy tę z minimalną odległością w metrach
            $RequestBody['routingPreference'] = 'TRAFFIC_UNAWARE'
            if (-not $HasIntermediates) {
                $RequestBody['computeAlternativeRoutes'] = $true
            }
        }
        'Eco' {
            # Trasa ekologiczna (najmniejsze zużycie paliwa/energii)
            # Wymaga TRAFFIC_AWARE_OPTIMAL oraz routeModifiers.vehicleInfo
            $RequestBody['routingPreference'] = 'TRAFFIC_AWARE_OPTIMAL'
            $RequestBody['requestedReferenceRoutes'] = @('FUEL_EFFICIENT')
            $RequestBody['routeModifiers'] = @{
                vehicleInfo = @{
                    emissionType = $EmissionType
                }
            }
        }
    }

    $Headers = @{
        'X-Goog-Api-Key'   = $ApiKey
        'Content-Type'     = 'application/json'
        'X-Goog-FieldMask' = 'routes.duration,routes.distanceMeters,routes.polyline.encodedPolyline,routes.description,routes.routeLabels'
    }

    try {
        Write-Verbose "Routes API ($RouteType): ($OriginLat,$OriginLng) -> ($DestLat,$DestLng), Punkty pośrednie: $($IntermediatePoints.Count)"
        $JsonBody = $RequestBody | ConvertTo-Json -Depth 10
        $Response = Invoke-RestMethod -Uri $RoutesUrl -Method POST -Headers $Headers -Body $JsonBody -TimeoutSec 60

        $Routes = @($Response.routes)
        if ($Routes.Count -eq 0) {
            Write-Warning "Routes API nie zwróciło żadnej trasy."
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

        # Wybór odpowiedniej trasy
        $SelectedRoute = $null
        if ($RouteType -eq 'Shortest') {
            # Wybieramy najkrótszą pod kątem odległości
            $SelectedRoute = $Routes | Sort-Object -Property { [int64]($_.distanceMeters) } | Select-Object -First 1
        }
        elseif ($RouteType -eq 'Eco') {
            # Sprawdzamy czy API zwróciło trasę oznaczoną jako FUEL_EFFICIENT
            $EcoRoute = $Routes | Where-Object {
                $_.routeLabels -and (@($_.routeLabels) -contains 'FUEL_EFFICIENT')
            } | Select-Object -First 1

            if ($EcoRoute) {
                $SelectedRoute = $EcoRoute
            }
            else {
                # Jeśli nie ma specyficznej etykiety, bierzemy pierwszą (optymalną)
                $SelectedRoute = $Routes[0]
            }
        }
        else {
            # Fastest: sortujemy po minimalnym czasie
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
        Write-Error "Błąd Routes API ($RouteType): $ErrorMsg"
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
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Origin,
        [Parameter(Mandatory)][string]$Destination,
        [Parameter()][object[]]$Waypoints = @(),
        [Parameter()][string]$TravelMode = 'driving'
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

# ══════════════════════════════════════════════════════════════════════════════
# 4. GENEROWANIE I ZAPIS MAPY PNG (GOOGLE STATIC MAPS API + GDI+ NAKŁADKA)
# ══════════════════════════════════════════════════════════════════════════════

function Get-WrappedLines {
    param(
        [System.Drawing.Graphics]$G,
        [string]$Text,
        [System.Drawing.Font]$F,
        [float]$MaxW
    )
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
        [Parameter()][string]$TekstAdresA = '',
        [Parameter()][string]$TekstAdresB = '',
        [Parameter()][string]$TekstOdleglosc = '',
        [Parameter()][string]$TekstCzas = '',
        [Parameter()][string]$TekstNaglowekLewy = '',
        [Parameter()][string]$TekstNaglowekPrawy = '',
        [Parameter()][string]$TekstUmowa = '',
        [Parameter()][string]$TekstKierunek = '',
        [Parameter()][string]$Opis = '',
        [Parameter()][string]$DataWygenerowania = '',
        [Parameter()][string]$LanguageCode = 'en'
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

        # Header Left: Description / Title / Contract
        if ([string]::IsNullOrWhiteSpace($TekstNaglowekLewy)) {
            if (-not [string]::IsNullOrWhiteSpace($Opis)) {
                $TekstNaglowekLewy = $Opis.Trim()
            } elseif (-not [string]::IsNullOrWhiteSpace($TekstUmowa)) {
                $TekstNaglowekLewy = if ($TekstUmowa -match '^(numer\s*umowy|contract|umowa|nr\s*umowy):') { $TekstUmowa } else { "Contract: $TekstUmowa" }
            }
        }

        # Header Right: Direction / Route Type / Date
        if ([string]::IsNullOrWhiteSpace($TekstNaglowekPrawy)) {
            if (-not [string]::IsNullOrWhiteSpace($TekstKierunek)) {
                $prefixDir = switch ($lang) { 'de' { 'Richtung: ' } 'pl' { 'Kierunek: ' } default { 'Direction: ' } }
                $TekstNaglowekPrawy = if ($TekstKierunek -match '^(kierunek|direction|route|trasa|richtung):') { $TekstKierunek } else { "$prefixDir$TekstKierunek" }
            }
        }
        else {
            # Localize passed Type / Typ string according to $LanguageCode
            if ($TekstNaglowekPrawy -match '^(?:Type|Typ|Art):\s*(.+)$' -or $TekstNaglowekPrawy -match '^(Shortest|Fastest|Eco|Najkr[oó]tsza|Najszybsza|Eko|K[uü]rzeste|Schnellste)$') {
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
                $TekstNaglowekPrawy = "$tPrefix$tName"
            }
        }

        # Format distance and duration string
        $TekstOdlegloscWyswietlana = $TekstOdleglosc
        if (-not [string]::IsNullOrWhiteSpace($TekstCzas)) {
            $TekstOdlegloscWyswietlana = if ($TekstOdlegloscWyswietlana) { "$TekstOdlegloscWyswietlana  ($TekstCzas)" } else { $TekstCzas }
        }

        $MaTopOverlay = (-not [string]::IsNullOrWhiteSpace($TekstNaglowekLewy)) -or (-not [string]::IsNullOrWhiteSpace($TekstNaglowekPrawy))
        $MaBottomOverlay = (-not [string]::IsNullOrWhiteSpace($TekstAdresA)) -or (-not [string]::IsNullOrWhiteSpace($TekstAdresB)) -or (-not [string]::IsNullOrWhiteSpace($TekstOdlegloscWyswietlana))

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

                $TopBarH = if ($MaTopOverlay) { 38 } else { 0 }

                # Pre-measure bottom text lines using temporary Graphics
                $dummyBmp = [System.Drawing.Bitmap]::new(1, 1)
                $measGfx  = [System.Drawing.Graphics]::FromImage($dummyBmp)

                $PadX       = 14
                $LineH      = 20
                $LabelASize = $measGfx.MeasureString('A: ', $FontBadge)
                $LabelBSize = $measGfx.MeasureString('B: ', $FontBadge)
                $LabelMaxW  = [float][math]::Max($LabelASize.Width, $LabelBSize.Width)
                $AvailW     = [float]($ActualW - ($PadX * 2))
                $AddrMaxW   = [float]($AvailW - $LabelMaxW)

                $LinesA = @(Get-WrappedLines -G $measGfx -Text $TekstAdresA -F $FontAddr -MaxW $AddrMaxW)
                $LinesB = @(Get-WrappedLines -G $measGfx -Text $TekstAdresB -F $FontAddr -MaxW $AddrMaxW)
                $measGfx.Dispose()
                $dummyBmp.Dispose()

                $LinesACount = [math]::Max(1, $LinesA.Count)
                $LinesBCount = [math]::Max(1, $LinesB.Count)

                $BtmPadTop   = 10
                $BtmPadBot   = 10
                $SpacingAB   = 4
                $SpacingDist = 8
                $DistLineH   = 24

                $BtmBarH = if ($MaBottomOverlay) {
                    $BtmPadTop + ($LinesACount * $LineH) + $SpacingAB + ($LinesBCount * $LineH) + $SpacingDist + $DistLineH + $BtmPadBot
                } else { 0 }

                # Canvas extension: Map is preserved 100% in the middle with extra height on top and bottom
                $FinalW = $ActualW
                $FinalH = $ActualH + $TopBarH + $BtmBarH

                $Bitmap = [System.Drawing.Bitmap]::new($FinalW, $FinalH, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
                $Graphics = [System.Drawing.Graphics]::FromImage($Bitmap)
                $Graphics.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
                $Graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit

                # 1. Fill solid background with dark slate #0F172A
                $BrushBg = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 15, 23, 42))
                $Graphics.FillRectangle($BrushBg, 0, 0, $FinalW, $FinalH)

                # 2. Draw map in the middle — untouched and 100% visible (no overlays covering map content)
                $Graphics.DrawImage($BitmapSrc, 0, $TopBarH, $ActualW, $ActualH)

                # 3. Separator lines and color brushes
                $PenSep      = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(255, 51, 65, 85), 1.5)
                $BrushWhite  = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 248, 250, 252))
                $BrushYellow = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 250, 204, 21))
                $BrushCyan   = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 56, 189, 248))
                $BrushGreen  = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 16, 185, 129))
                $BrushRed    = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 239, 68, 68))
                $BrushMuted  = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 148, 163, 184))

                # 4. Top Header Banner (Extended on top)
                if ($MaTopOverlay) {
                    $Graphics.DrawLine($PenSep, 0, $TopBarH, $FinalW, $TopBarH)
                    $TopTextY = [float](($TopBarH - $FontTopTitle.Height) / 2)
                    $CurLeftX = [float]$PadX

                    if (-not [string]::IsNullOrWhiteSpace($TekstNaglowekLewy)) {
                        $Graphics.DrawString($TekstNaglowekLewy, $FontTopTitle, $BrushWhite, $CurLeftX, $TopTextY)
                        $CurLeftX += $Graphics.MeasureString($TekstNaglowekLewy, $FontTopTitle).Width
                    }

                    if (-not [string]::IsNullOrWhiteSpace($TekstNaglowekPrawy)) {
                        $SizeR = $Graphics.MeasureString($TekstNaglowekPrawy, $FontTopType)
                        $RightStartX = [float][math]::Max($CurLeftX + 15, $FinalW - $PadX - $SizeR.Width)
                        $Graphics.DrawString($TekstNaglowekPrawy, $FontTopType, $BrushYellow, $RightStartX, $TopTextY)
                    }
                }

                # 5. Bottom Footer Banner (Extended on bottom, distance & time a line lower)
                if ($MaBottomOverlay) {
                    $BtmBarY = $TopBarH + $ActualH
                    $Graphics.DrawLine($PenSep, 0, $BtmBarY, $FinalW, $BtmBarY)

                    $CurY = [float]($BtmBarY + $BtmPadTop)
                    $LabelXOffset = [float]($PadX + $LabelMaxW)

                    # Line 1: Origin [A]
                    $Graphics.DrawString('A: ', $FontBadge, $BrushGreen, [float]$PadX, $CurY)
                    foreach ($Line in $LinesA) {
                        $Graphics.DrawString($Line, $FontAddr, $BrushWhite, $LabelXOffset, $CurY)
                        $CurY += [float]$LineH
                    }

                    # Line 2: Destination [B]
                    $CurY += [float]$SpacingAB
                    $Graphics.DrawString('B: ', $FontBadge, $BrushRed, [float]$PadX, $CurY)
                    foreach ($Line in $LinesB) {
                        $Graphics.DrawString($Line, $FontAddr, $BrushWhite, $LabelXOffset, $CurY)
                        $CurY += [float]$LineH
                    }

                    # Line 3: Distance and Duration A LINE LOWER (never overlaps with addresses)
                    $CurY += [float]$SpacingDist
                    if (-not [string]::IsNullOrWhiteSpace($TekstOdlegloscWyswietlana)) {
                        $distLblText = switch ($LanguageCode) {
                            'de'    { "Gesamt: " }
                            'pl'    { "Razem: " }
                            default { "Total: " }
                        }
                        $distLblSize = $Graphics.MeasureString($distLblText, $FontDistLbl)
                        $Graphics.DrawString($distLblText, $FontDistLbl, $BrushCyan, [float]$PadX, [float]($CurY + 2))
                        $Graphics.DrawString($TekstOdlegloscWyswietlana, $FontDist, $BrushYellow, [float]($PadX + $distLblSize.Width), $CurY)
                    }

                    # Timestamp on Line 3 (Right aligned)
                    $DateStr   = if ($DataWygenerowania) { $DataWygenerowania } else { (Get-Date -Format 'yyyy-MM-dd  HH:mm') }
                    $DateSizeF = $Graphics.MeasureString($DateStr, $FontDate)
                    $DateX     = [float]($FinalW - $DateSizeF.Width - $PadX)
                    $DateY     = [float]($CurY + 3)
                    $Graphics.DrawString($DateStr, $FontDate, $BrushMuted, $DateX, $DateY)
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

# ══════════════════════════════════════════════════════════════════════════════
# 5. UNIWERSALNY IMPORT DANYCH (JSON, CSV, EXCEL)
# ══════════════════════════════════════════════════════════════════════════════

function Find-MatchingPropertyName {
    param(
        [Parameter(Mandatory)][string[]]$AvailableProperties,
        [Parameter(Mandatory)][string[]]$Patterns
    )
    foreach ($pattern in $Patterns) {
        $found = $AvailableProperties | Where-Object {
            $null -ne $_ -and $_.Trim() -match $pattern
        } | Select-Object -First 1
        if ($found) { return $found }
    }
    return $null
}

function Import-RouteDataFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter()][string]$Delimiter = ''
    )

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
            $UsedDelimiter = if (-not [string]::IsNullOrWhiteSpace($Delimiter)) {
                $Delimiter
            }
            elseif ($Extension -eq '.tsv' -or $FirstLine -match "`t") {
                "`t"
            }
            elseif ($FirstLine -match ';') {
                ';'
            }
            else {
                ','
            }
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

    # Analiza kolumn pierwszego wiersza
    $PropNames = @($RawRows[0].PSObject.Properties.Name)

    # 1. Sprawdzamy czy plik to sekwencja punktów jednej trasy wielopunktowej (SequentialStops)
    # np. kolumny: LP / Kolejnosc + Adres / Lokalizacja
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

        # Generujemy z sekwencji pojedynczą trasę z punktami pośrednimi
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

    # 2. Tryb RouteList (każdy wiersz to osobna trasa)
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
        if ([string]::IsNullOrWhiteSpace($startVal) -or [string]::IsNullOrWhiteSpace($endVal)) {
            continue
        }

        $nameVal = if ($ColName) { [string]$row.$ColName } else { "Trasa $idx" }
        $typeVal = if ($ColRouteType) { [string]$row.$ColRouteType } else { $null }

        # Normalizacja RouteType
        if ($typeVal -match '(?i)eco|fuel|paliw|eko') { $typeVal = 'Eco' }
        elseif ($typeVal -match '(?i)short|krot|krót') { $typeVal = 'Shortest' }
        elseif ($typeVal -match '(?i)fast|szyb') { $typeVal = 'Fastest' }
        else { $typeVal = $null }

        # Obsługa punktów pośrednich
        $waypointsList = [System.Collections.Generic.List[string]]::new()
        if ($ColWaypoints -and -not [string]::IsNullOrWhiteSpace($row.$ColWaypoints)) {
            $rawWp = $row.$ColWaypoints
            if ($rawWp -is [System.Collections.IEnumerable] -and -not ($rawWp -is [string])) {
                foreach ($item in $rawWp) {
                    if (-not [string]::IsNullOrWhiteSpace([string]$item)) {
                        $waypointsList.Add(([string]$item).Trim())
                    }
                }
            }
            else {
                $splits = ([string]$rawWp) -split '(?<!\\)[|;]'
                foreach ($s in $splits) {
                    $cleaned = $s.Trim()
                    if (-not [string]::IsNullOrWhiteSpace($cleaned)) {
                        $waypointsList.Add($cleaned)
                    }
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

# ══════════════════════════════════════════════════════════════════════════════
# 6. UNIWERSALNY EKSPORT WYNIKÓW (EXCEL, CSV, JSON)
# ══════════════════════════════════════════════════════════════════════════════

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

    switch ($Format) {
        'Excel' {
            if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
                Write-Warning "Moduł ImportExcel nie jest zainstalowany. Eksportowanie do CSV zamiast Excel."
                $CsvPath = [System.IO.Path]::ChangeExtension($OutputPath, '.csv')
                $Results | Export-Csv -LiteralPath $CsvPath -NoTypeInformation -Encoding UTF8 -Delimiter ';'
                return $CsvPath
            }
            Import-Module -Name ImportExcel -ErrorAction Stop
            $Results | Export-Excel -Path $OutputPath -WorksheetName 'Trasy' -TableName 'WynikiTras' -AutoSize -AutoFilter -FreezeTopRow
            return $OutputPath
        }
        'CSV' {
            $Results | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8 -Delimiter ';'
            return $OutputPath
        }
        'JSON' {
            $JsonContent = $Results | ConvertTo-Json -Depth 5
            [System.IO.File]::WriteAllText($OutputPath, $JsonContent, [System.Text.Encoding]::UTF8)
            return $OutputPath
        }
    }
}
