#Requires -Version 5.1
<#
.SYNOPSIS
    Processes school transport contracts — generates route maps and Excel distance summaries.

.DESCRIPTION
    Loads an Excel file with contract columns: Umowa (Contract), Szkoła (School), Dom (Home),
    Praca (Work), Tryb (Mode), and Wariant (Variant).
    Creates a dedicated subfolder per contract, renders PNG route maps (Home<->School, optional Work stop),
    and generates an individual Excel summary. At completion, creates a consolidated Excel workbook for all contracts.

.PARAMETER ApiKey
    Google Maps API key. Default: resolved from GOOGLE_MAPS_API_KEY environment variable.

.PARAMETER InputExcel
    Path to input Excel file. If omitted, an interactive file picker dialog opens.

.PARAMETER OutputFolder
    Output directory for contract subfolders. Default: C:\Temp\SchoolRoutes

.PARAMETER MapWidth
    Rendered PNG map width in pixels (100-2048). Default: 900.

.PARAMETER MapHeight
    Rendered PNG map height in pixels (100-2048). Default: 600.

.EXAMPLE
    .\Process-SchoolTransportRoutes.ps1
    Launches script with interactive file picker dialog.

.EXAMPLE
    .\Process-SchoolTransportRoutes.ps1 -InputExcel "C:\contract_data.xlsx" -OutputFolder "C:\results"
    Processes specified contract Excel file and outputs results to target directory.

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
    [string]$OutputFolder = 'C:\Temp\SchoolRoutes',

    [Parameter(Mandatory = $false)]
    [ValidateRange(100, 2048)]
    [int]$MapWidth = 900,

    [Parameter(Mandatory = $false)]
    [ValidateRange(100, 2048)]
    [int]$MapHeight = 600
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Validate API Key ──────────────────────────────────────────────────────────
if ([string]::IsNullOrWhiteSpace($ApiKey)) {
    throw "Missing Google Maps API key. Set GOOGLE_MAPS_API_KEY environment variable or pass -ApiKey parameter."
}

# ── Import shared functions ────────────────────────────────────────────────────
. "$PSScriptRoot\RouteMapFunctions.ps1"

# ── Tworzenie folderu wynikowego ──────────────────────────────────────────────
if (-not (Test-Path -Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
    Write-Verbose "Utworzono folder wynikowy: $OutputFolder"
}

# ── Funkcja pomocnicza: sanityzacja nazwy pliku/folderu ───────────────────────
function ConvertTo-SafeFileName {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Name)
    # Replace forbidden Windows path characters with '-'
    $InvalidChars = [System.IO.Path]::GetInvalidFileNameChars()
    $Safe = $Name
    foreach ($Char in $InvalidChars) {
        $Safe = $Safe.Replace([string]$Char, '-')
    }
    # Replace spaces with '-' for readability
    $Safe = $Safe.Trim()
    return $Safe
}

# ── Helper: verify if work address is valid (not empty, not "Not applicable") ──
function Test-PracaAddress {
    [CmdletBinding()]
    param([Parameter()][string]$Praca)
    if ([string]::IsNullOrWhiteSpace($Praca)) { return $false }
    $Normalized = $Praca.Trim().ToLower()
    if ($Normalized -in @('nie dotyczy', 'nd', 'n/d', 'nd.', '-', 'brak')) { return $false }
    return $true
}

# ── Helper: calculate route and render PNG map ─────────────────────────────────
function Invoke-RouteAndMap {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][PSCustomObject]$GeoStart,
        [Parameter(Mandatory)][PSCustomObject]$GeoEnd,
        [Parameter(Mandatory)][string]$PngPath,
        [Parameter(Mandatory)][string]$LabelStart,
        [Parameter(Mandatory)][string]$LabelEnd,
        [Parameter(Mandatory)][string]$ApiKey,
        [Parameter()][int]$Width = 900,
        [Parameter()][int]$Height = 600,
        [Parameter()][string]$NumerUmowy = '',
        [Parameter()][string]$Opis = '',
        [Parameter()][string]$DataWygenerowania = ''
    )

    $Trasa = Get-CarRouteData -OriginLat $GeoStart.Latitude -OriginLng $GeoStart.Longitude `
        -DestLat $GeoEnd.Latitude -DestLng $GeoEnd.Longitude -ApiKey $ApiKey

    if ($Trasa.Status -ne 'OK') {
        Write-Warning "  Brak trasy ${LabelStart} -> ${LabelEnd}: $($Trasa.Status)"
        return [PSCustomObject]@{ OdlegloscKm = $null; Status = $Trasa.Status; MapSaved = $false }
    }

    Write-Host "    $LabelStart -> $LabelEnd : $($Trasa.OdlegloscKm) km" -ForegroundColor Green

    $MapSaved = $false
    if ($Trasa.EncodedPolyline) {
        $OdlTekst = if ($Trasa.OdlegloscKm) { "$($Trasa.OdlegloscKm) km" } else { '' }
        $TekstA = if ($GeoStart.FormattedAddress) { $GeoStart.FormattedAddress } else { $LabelStart }
        $TekstB = if ($GeoEnd.FormattedAddress) { $GeoEnd.FormattedAddress } else { $LabelEnd }
        $KierunekTekst = "$LabelStart -> $LabelEnd"

        $MapSaved = Save-RouteMapPng -EncodedPolyline $Trasa.EncodedPolyline `
            -OriginLat $GeoStart.Latitude -OriginLng $GeoStart.Longitude `
            -DestLat $GeoEnd.Latitude -DestLng $GeoEnd.Longitude `
            -OutputPath $PngPath -ApiKey $ApiKey -Width $Width -Height $Height `
            -AddressTextA $TekstA -AddressTextB $TekstB -DistanceText $OdlTekst `
            -ContractText $NumerUmowy -DirectionText $KierunekTekst `
            -Description $Opis -GeneratedDate $DataWygenerowania

        if ($MapSaved) {
            Write-Host "    Mapa: $(Split-Path -Leaf $PngPath)" -ForegroundColor Cyan
        }
    }

    return [PSCustomObject]@{ OdlegloscKm = $Trasa.OdlegloscKm; Status = $Trasa.Status; MapSaved = $MapSaved }
}

# ══════════════════════════════════════════════════════════════════════════════
# MAIN EXECUTION LOGIC
# ══════════════════════════════════════════════════════════════════════════════

# ── Select Input File ──────────────────────────────────────────────────────────
if ([string]::IsNullOrWhiteSpace($InputExcel)) {
    Write-Host "Otwieranie dialogu wyboru pliku Excel..." -ForegroundColor Cyan
    $InputExcel = Select-InputExcel
    if ([string]::IsNullOrWhiteSpace($InputExcel)) {
        Write-Warning "Nie wybrano pliku. Skrypt zakończony."
        exit 0
    }
}

if (-not (Test-Path -Path $InputExcel)) { throw "Plik wejściowy nie istnieje: $InputExcel" }

Write-Host "Wczytywanie pliku: $InputExcel" -ForegroundColor Cyan

try {
    $Dane = Import-Excel -Path $InputExcel
}
catch {
    throw "Nie można wczytać pliku Excel. Upewnij się że moduł ImportExcel jest zainstalowany. Błąd: $($_.Exception.Message)"
}

if ($null -eq $Dane -or @($Dane).Count -eq 0) {
    Write-Warning "Plik Excel jest pusty lub nie zawiera danych."
    exit 0
}

$Headers = $Dane[0].PSObject.Properties.Name
Write-Verbose "Kolumny w pliku: $($Headers -join ', ')"

# ── Mapowanie kolumn ──────────────────────────────────────────────────────────
function Find-ColumnHeader {
    param(
        [Parameter(Mandatory)][string[]]$Headers,
        [Parameter(Mandatory)][string[]]$Patterns
    )
    foreach ($pattern in $Patterns) {
        $found = $Headers | Where-Object {
            if ($null -ne $_) {
                $cleaned = $_.Trim()
                $cleaned -match $pattern
            }
        } | Select-Object -First 1
        if ($found) { return $found }
    }
    return $null
}

$ColMap = @{
    Umowa   = Find-ColumnHeader -Headers $Headers -Patterns @('(?i)^\s*(umow[ay]|nr\s*umow[ay]|numer\s*umow[ay]|id|numer|nr)\s*$', '(?i)umow')
    Opis    = Find-ColumnHeader -Headers $Headers -Patterns @('(?i)^\s*(opis|nazwa|nazwa\s*folderu|folder|identyfikator)\s*$', '(?i)opis')
    Szkola  = Find-ColumnHeader -Headers $Headers -Patterns @('(?i)^\s*(szko[łl][ay]|adres\s*szko[łl][yi]|plac[oó]wk[ay]|przedszkol[ea]|o[sś]rodek|szko[łl]a\s*adres)\s*$', '(?i)szko[łl]|plac[oó]wk|przedszkol')
    Dom     = Find-ColumnHeader -Headers $Headers -Patterns @('(?i)^\s*(dom|domu|adres\s*domu|adres\s*zamieszkania|zamieszkani[ea]|miejsce\s*zamieszkania|dom\s*adres)\s*$', '(?i)dom|zamieszkan')
    Praca   = Find-ColumnHeader -Headers $Headers -Patterns @('(?i)^\s*(prac[ay]|adres\s*prac[ay]|miejsce\s*prac[ay]|zak[łl]ad\s*prac[ay]|firma|praca\s*adres)\s*$', '(?i)prac|firm')
    Tryb    = Find-ColumnHeader -Headers $Headers -Patterns @('(?i)^\s*(tryb|tryb\s*przejazdu|spos[oó]b)\s*$', '(?i)tryb')
    Wariant = Find-ColumnHeader -Headers $Headers -Patterns @('(?i)^\s*(wariant|wariant\s*trasy|opcja)\s*$', '(?i)wariant|opcj')
}

# Walidacja wymaganych kolumn
$RequiredCols = @('Umowa', 'Szkola', 'Dom')
foreach ($ColName in $RequiredCols) {
    if (-not $ColMap[$ColName]) {
        Write-Warning "Dostępne kolumny: $($Headers -join ', ')"
        throw "Nie można odnaleźć wymaganej kolumny '$ColName' w pliku Excel."
    }
}

Write-Host "Znalezione kolumny:" -ForegroundColor Green
$ColMap.GetEnumerator() | Sort-Object Name | ForEach-Object {
    $Status = if ($_.Value) { "'$($_.Value)'" } else { '(nie znaleziono)' }
    Write-Host "  $($_.Name) -> $Status" -ForegroundColor $(if ($_.Value) { 'Green' } else { 'Yellow' })
}

# ── Main Processing Loop ─────────────────────────────────────────────────────
$WszystkieWyniki = [System.Collections.Generic.List[PSCustomObject]]::new()
$RowIndex = 0
$TotalRows = @($Dane).Count

# Geocoding cache — addresses frequently repeat across contracts
$GeoCache = @{}

foreach ($Row in $Dane) {
    $RowIndex++

    $NumerUmowy = if ($ColMap.Umowa -and $null -ne $Row.($ColMap.Umowa)) { ($Row.($ColMap.Umowa) -as [string]).Trim() } else { $null }
    $Opis = if ($ColMap.Opis -and $null -ne $Row.($ColMap.Opis)) { ($Row.($ColMap.Opis) -as [string]).Trim() } else { $null }
    $AdresSzkoly = if ($ColMap.Szkola -and $null -ne $Row.($ColMap.Szkola)) { ($Row.($ColMap.Szkola) -as [string]).Trim() } else { $null }
    $AdresDomu = if ($ColMap.Dom -and $null -ne $Row.($ColMap.Dom)) { ($Row.($ColMap.Dom) -as [string]).Trim() } else { $null }
    $AdresPracy = if ($ColMap.Praca -and $null -ne $Row.($ColMap.Praca)) { ($Row.($ColMap.Praca) -as [string]).Trim() } else { $null }
    $Tryb = if ($ColMap.Tryb -and $null -ne $Row.($ColMap.Tryb)) { ($Row.($ColMap.Tryb) -as [string]).Trim() } else { $null }
    $Wariant = if ($ColMap.Wariant -and $null -ne $Row.($ColMap.Wariant)) { ($Row.($ColMap.Wariant) -as [string]).Trim() } else { $null }

    if ([string]::IsNullOrWhiteSpace($NumerUmowy)) {
        Write-Warning "Wiersz ${RowIndex}: Pusty numer umowy — pomijam."
        continue
    }

    $SafeIdentifier = if (-not [string]::IsNullOrWhiteSpace($Opis)) { $Opis } else { $NumerUmowy }
    $SafeName = ConvertTo-SafeFileName -Name $SafeIdentifier
    $DisplayLabel = if ($Opis) { "$NumerUmowy ($Opis)" } else { $NumerUmowy }
    Write-Host "`n[$RowIndex/$TotalRows] Umowa: $DisplayLabel (Folder: $SafeName)" -ForegroundColor Yellow
    Write-Host "  Dom:    $AdresDomu" -ForegroundColor White
    Write-Host "  Szkoła: $AdresSzkoly" -ForegroundColor White
    if (Test-PracaAddress -Praca $AdresPracy) {
        Write-Host "  Praca:  $AdresPracy" -ForegroundColor White
    }
    if ($Tryb) { Write-Host "  Tryb:   $Tryb" -ForegroundColor DarkGray }
    if ($Wariant) { Write-Host "  Wariant: $Wariant" -ForegroundColor DarkGray }

    # ── Tworzenie folderu umowy ───────────────────────────────────────────────
    $UmowaFolder = Join-Path -Path $OutputFolder -ChildPath $SafeName
    if (-not (Test-Path -Path $UmowaFolder)) {
        New-Item -ItemType Directory -Path $UmowaFolder -Force | Out-Null
    }

    # ── Walidacja adresów ─────────────────────────────────────────────────────
    if ([string]::IsNullOrWhiteSpace($AdresDomu) -or [string]::IsNullOrWhiteSpace($AdresSzkoly)) {
        Write-Warning "  Wiersz ${RowIndex}: Pusty adres Domu lub Szkoły — pomijam."
        $WszystkieWyniki.Add([PSCustomObject]@{
                'Numer umowy'        = $NumerUmowy
                'Opis'               = $Opis
                'Tryb'               = $Tryb
                'Wariant'            = $Wariant
                'AdresDomu'          = $AdresDomu
                'Dom_Geokodowany'    = $null
                'Dom_UlicaINumer'    = $null
                'Dom_KodPocztowy'    = $null
                'Dom_Miasto'         = $null
                'AdresSzkoly'        = $AdresSzkoly
                'Szkoła_Geokodowany'  = $null
                'Szkoła_UlicaINumer'  = $null
                'Szkoła_KodPocztowy'  = $null
                'Szkoła_Miasto'      = $null
                'AdresPracy'         = $AdresPracy
                'Praca_Geokodowany'  = $null
                'Praca_UlicaINumer'  = $null
                'Praca_KodPocztowy'  = $null
                'Praca_Miasto'       = $null
                'Dom→Szkoła [km]'    = $null
                'Szkoła→Dom [km]'    = $null
                'Dom→Praca [km]'     = $null
                'Praca→Dom [km]'     = $null
                'Szkoła→Praca [km]'  = $null
                'Praca→Szkoła [km]'  = $null
                'Status'             = 'PUSTE_ADRESY'
            })
        continue
    }

    # ── Geokodowanie (z cache) ────────────────────────────────────────────────
    $GeoDom = $null; $GeoSzkola = $null; $GeoPraca = $null
    $HasError = $false

    # Dom
    if ($GeoCache.ContainsKey($AdresDomu)) {
        $GeoDom = $GeoCache[$AdresDomu]
    }
    else {
        $GeoDom = Get-AddressCoordinates -Address $AdresDomu -ApiKey $ApiKey
        $GeoCache[$AdresDomu] = $GeoDom
        Start-Sleep -Milliseconds 200
    }
    if (-not $GeoDom -or $GeoDom.Status -ne 'OK') {
        Write-Warning "  Błąd geokodowania Domu: $($GeoDom.Status)"
        $HasError = $true
    }

    # School
    if ($GeoCache.ContainsKey($AdresSzkoly)) {
        $GeoSzkola = $GeoCache[$AdresSzkoly]
    }
    else {
        $GeoSzkola = Get-AddressCoordinates -Address $AdresSzkoly -ApiKey $ApiKey
        $GeoCache[$AdresSzkoly] = $GeoSzkola
        Start-Sleep -Milliseconds 200
    }
    if (-not $GeoSzkola -or $GeoSzkola.Status -ne 'OK') {
        Write-Warning "  Błąd geokodowania Szkoły: $($GeoSzkola.Status)"
        $HasError = $true
    }

    # Praca (opcjonalnie)
    $MaPrace = Test-PracaAddress -Praca $AdresPracy
    if ($MaPrace) {
        if ($GeoCache.ContainsKey($AdresPracy)) {
            $GeoPraca = $GeoCache[$AdresPracy]
        }
        else {
            $GeoPraca = Get-AddressCoordinates -Address $AdresPracy -ApiKey $ApiKey
            $GeoCache[$AdresPracy] = $GeoPraca
            Start-Sleep -Milliseconds 200
        }
        if (-not $GeoPraca -or $GeoPraca.Status -ne 'OK') {
            Write-Warning "  Błąd geokodowania Pracy: $($GeoPraca.Status)"
            # Work is optional — continue without work routes if geocoding fails
            $MaPrace = $false
        }
    }

    if ($HasError) {
        $WszystkieWyniki.Add([PSCustomObject]@{
                'Numer umowy'        = $NumerUmowy
                'Opis'               = $Opis
                'Tryb'               = $Tryb
                'Wariant'            = $Wariant
                'AdresDomu'          = $AdresDomu
                'Dom_Geokodowany'    = if ($GeoDom) { $GeoDom.FormattedAddress } else { $null }
                'Dom_UlicaINumer'    = if ($GeoDom) { $GeoDom.UlicaINumer } else { $null }
                'Dom_KodPocztowy'    = if ($GeoDom) { $GeoDom.KodPocztowy } else { $null }
                'Dom_Miasto'         = if ($GeoDom) { $GeoDom.Miasto } else { $null }
                'AdresSzkoly'        = $AdresSzkoly
                'Szkoła_Geokodowany'  = if ($GeoSzkola) { $GeoSzkola.FormattedAddress } else { $null }
                'Szkoła_UlicaINumer'  = if ($GeoSzkola) { $GeoSzkola.UlicaINumer } else { $null }
                'Szkoła_KodPocztowy'  = if ($GeoSzkola) { $GeoSzkola.KodPocztowy } else { $null }
                'Szkoła_Miasto'      = if ($GeoSzkola) { $GeoSzkola.Miasto } else { $null }
                'AdresPracy'         = $AdresPracy
                'Praca_Geokodowany'  = if ($GeoPraca) { $GeoPraca.FormattedAddress } else { $null }
                'Praca_UlicaINumer'  = if ($GeoPraca) { $GeoPraca.UlicaINumer } else { $null }
                'Praca_KodPocztowy'  = if ($GeoPraca) { $GeoPraca.KodPocztowy } else { $null }
                'Praca_Miasto'       = if ($GeoPraca) { $GeoPraca.Miasto } else { $null }
                'Dom→Szkoła [km]'    = $null
                'Szkoła→Dom [km]'    = $null
                'Dom→Praca [km]'     = $null
                'Praca→Dom [km]'     = $null
                'Szkoła→Praca [km]'  = $null
                'Praca→Szkoła [km]'  = $null
                'Status'             = 'BLAD_GEOKODOWANIA'
            })
        continue
    }

    # ── Inicjalizacja wyników odległości ──────────────────────────────────────
    $KmDomSzkola = $null
    $KmSzkolaDom = $null
    $KmDomPraca = $null
    $KmPracaDom = $null
    $KmSzkolaPraca = $null
    $KmPracaSzkola = $null

    $curDate = (Get-Date).ToString('yyyy-MM-dd')

    # ── Trasa 1: Dom → Szkoła ─────────────────────────────────────────────────
    $PngName = "${SafeName}_Dom_Szkoła.png"
    $PngPath = Join-Path -Path $UmowaFolder -ChildPath $PngName
    $Result = Invoke-RouteAndMap -GeoStart $GeoDom -GeoEnd $GeoSzkola `
        -PngPath $PngPath -LabelStart 'Dom' -LabelEnd 'Szkoła' `
        -ApiKey $ApiKey -Width $MapWidth -Height $MapHeight `
        -NumerUmowy $NumerUmowy -Opis $Opis -DataWygenerowania $curDate
    $KmDomSzkola = $Result.OdlegloscKm
    Start-Sleep -Milliseconds 250

    # ── Trasa 2: Szkoła → Dom ─────────────────────────────────────────────────
    $PngName = "${SafeName}_Szkoła_Dom.png"
    $PngPath = Join-Path -Path $UmowaFolder -ChildPath $PngName
    $Result = Invoke-RouteAndMap -GeoStart $GeoSzkola -GeoEnd $GeoDom `
        -PngPath $PngPath -LabelStart 'Szkoła' -LabelEnd 'Dom' `
        -ApiKey $ApiKey -Width $MapWidth -Height $MapHeight `
        -NumerUmowy $NumerUmowy -Opis $Opis -DataWygenerowania $curDate
    $KmSzkolaDom = $Result.OdlegloscKm
    Start-Sleep -Milliseconds 250

    # ── Trasy z Pracą (jeśli dotyczy) ─────────────────────────────────────────
    if ($MaPrace) {
        # Trasa 3: Dom → Praca
        $PngName = "${SafeName}_Dom_Praca.png"
        $PngPath = Join-Path -Path $UmowaFolder -ChildPath $PngName
        $Result = Invoke-RouteAndMap -GeoStart $GeoDom -GeoEnd $GeoPraca `
            -PngPath $PngPath -LabelStart 'Dom' -LabelEnd 'Praca' `
            -ApiKey $ApiKey -Width $MapWidth -Height $MapHeight `
            -NumerUmowy $NumerUmowy -Opis $Opis -DataWygenerowania $curDate
        $KmDomPraca = $Result.OdlegloscKm
        Start-Sleep -Milliseconds 250

        # Trasa 4: Praca → Dom
        $PngName = "${SafeName}_Praca_Dom.png"
        $PngPath = Join-Path -Path $UmowaFolder -ChildPath $PngName
        $Result = Invoke-RouteAndMap -GeoStart $GeoPraca -GeoEnd $GeoDom `
            -PngPath $PngPath -LabelStart 'Praca' -LabelEnd 'Dom' `
            -ApiKey $ApiKey -Width $MapWidth -Height $MapHeight `
            -NumerUmowy $NumerUmowy -Opis $Opis -DataWygenerowania $curDate
        $KmPracaDom = $Result.OdlegloscKm
        Start-Sleep -Milliseconds 250

        # Route 5: School -> Work
        $PngName = "${SafeName}_Szkoła_Praca.png"
        $PngPath = Join-Path -Path $UmowaFolder -ChildPath $PngName
        $Result = Invoke-RouteAndMap -GeoStart $GeoSzkola -GeoEnd $GeoPraca `
            -PngPath $PngPath -LabelStart 'Szkoła' -LabelEnd 'Praca' `
            -ApiKey $ApiKey -Width $MapWidth -Height $MapHeight `
            -NumerUmowy $NumerUmowy -Opis $Opis -DataWygenerowania $curDate
        $KmSzkolaPraca = $Result.OdlegloscKm
        Start-Sleep -Milliseconds 250

        # Route 6: Work -> School
        $PngName = "${SafeName}_Praca_Szkoła.png"
        $PngPath = Join-Path -Path $UmowaFolder -ChildPath $PngName
        $Result = Invoke-RouteAndMap -GeoStart $GeoPraca -GeoEnd $GeoSzkola `
            -PngPath $PngPath -LabelStart 'Praca' -LabelEnd 'Szkoła' `
            -ApiKey $ApiKey -Width $MapWidth -Height $MapHeight `
            -NumerUmowy $NumerUmowy -Opis $Opis -DataWygenerowania $curDate
        $KmPracaSzkola = $Result.OdlegloscKm
        Start-Sleep -Milliseconds 250
    }

    # ── Wynik dla tego wiersza ────────────────────────────────────────────────
    $WierszWynik = [PSCustomObject]@{
        'Numer umowy'        = $NumerUmowy
        'Opis'               = $Opis
        'Tryb'               = $Tryb
        'Wariant'            = $Wariant
        'AdresDomu'          = $AdresDomu
        'Dom_Geokodowany'    = if ($GeoDom) { $GeoDom.FormattedAddress } else { $null }
        'Dom_UlicaINumer'    = if ($GeoDom) { $GeoDom.UlicaINumer } else { $null }
        'Dom_KodPocztowy'    = if ($GeoDom) { $GeoDom.KodPocztowy } else { $null }
        'Dom_Miasto'         = if ($GeoDom) { $GeoDom.Miasto } else { $null }
        'AdresSzkoly'        = $AdresSzkoly
        'Szkoła_Geokodowany'  = if ($GeoSzkola) { $GeoSzkola.FormattedAddress } else { $null }
        'Szkoła_UlicaINumer'  = if ($GeoSzkola) { $GeoSzkola.UlicaINumer } else { $null }
        'Szkoła_KodPocztowy'  = if ($GeoSzkola) { $GeoSzkola.KodPocztowy } else { $null }
        'Szkoła_Miasto'      = if ($GeoSzkola) { $GeoSzkola.Miasto } else { $null }
        'AdresPracy'         = $AdresPracy
        'Praca_Geokodowany'  = if ($GeoPraca) { $GeoPraca.FormattedAddress } else { $null }
        'Praca_UlicaINumer'  = if ($GeoPraca) { $GeoPraca.UlicaINumer } else { $null }
        'Praca_KodPocztowy'  = if ($GeoPraca) { $GeoPraca.KodPocztowy } else { $null }
        'Praca_Miasto'       = if ($GeoPraca) { $GeoPraca.Miasto } else { $null }
        'Dom→Szkoła [km]'    = $KmDomSzkola
        'Szkoła→Dom [km]'    = $KmSzkolaDom
        'Dom→Praca [km]'     = $KmDomPraca
        'Praca→Dom [km]'     = $KmPracaDom
        'Szkoła→Praca [km]'  = $KmSzkolaPraca
        'Praca→Szkoła [km]'  = $KmPracaSzkola
        'Status'             = 'OK'
    }

    $WszystkieWyniki.Add($WierszWynik)

    # ── Excel podsumowania w folderze umowy ───────────────────────────────────
    $UmowaExcelPath = Join-Path -Path $UmowaFolder -ChildPath "${SafeName}.xlsx"
    $WierszWynik | Export-Excel -Path $UmowaExcelPath -WorksheetName 'Podsumowanie' `
        -TableName 'Podsumowanie' -AutoSize -AutoFilter -ClearSheet `
        -NoNumberConversion 'Numer umowy' -FreezeTopRow
    Write-Host "  Excel: ${SafeName}.xlsx" -ForegroundColor Cyan
}

# ══════════════════════════════════════════════════════════════════════════════
# ZBIORCZY EXCEL
# ══════════════════════════════════════════════════════════════════════════════

$DateStamp = Get-Date -Format 'yyyyMMdd_HHmm'
$ZbiorczyExcelPath = Join-Path -Path $OutputFolder -ChildPath "${DateStamp}_Podsumowanie_wszystkie_umowy.xlsx"

if ($WszystkieWyniki.Count -gt 0) {
    Write-Host "`nEksportowanie zbiorczego podsumowania: $ZbiorczyExcelPath" -ForegroundColor Cyan

    $WszystkieWyniki | Export-Excel -Path $ZbiorczyExcelPath -WorksheetName 'Wszystkie umowy' `
        -TableName 'WszystkieUmowy' -AutoSize -AutoFilter -ClearSheet `
        -NoNumberConversion 'Numer umowy' -FreezeTopRow

    # Add input data as second worksheet
    $Dane | Export-Excel -Path $ZbiorczyExcelPath -WorksheetName 'Dane wejściowe' `
        -TableName 'DaneWejsciowe' -AutoSize -AutoFilter -ClearSheet `
        -NoNumberConversion *

    Write-Host "Zbiorczy Excel zapisano: $ZbiorczyExcelPath" -ForegroundColor Green
}
else {
    Write-Warning "Brak wyników do zapisania."
}

# ══════════════════════════════════════════════════════════════════════════════
# PODSUMOWANIE
# ══════════════════════════════════════════════════════════════════════════════

$OK = @($WszystkieWyniki | Where-Object { $_.Status -eq 'OK' }).Count
$Bledy = @($WszystkieWyniki | Where-Object { $_.Status -ne 'OK' }).Count
$Umowy = $WszystkieWyniki.Count

Write-Host "`n════════════════════════════════════════════════" -ForegroundColor Magenta
Write-Host " PODSUMOWANIE PRZETWARZANIA UMÓW" -ForegroundColor Magenta
Write-Host "════════════════════════════════════════════════" -ForegroundColor Magenta
Write-Host " Przetworzone umowy    : $Umowy" -ForegroundColor White
Write-Host " OK                    : $OK" -ForegroundColor Green
Write-Host " Błędy                 : $Bledy" -ForegroundColor $(if ($Bledy -gt 0) { 'Red' } else { 'Green' })
Write-Host " Zbiorczy Excel        : $ZbiorczyExcelPath" -ForegroundColor White
Write-Host " Folder wynikowy       : $OutputFolder" -ForegroundColor White
Write-Host "════════════════════════════════════════════════`n" -ForegroundColor Magenta

$WszystkieWyniki

