#Requires -Version 5.1
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force -ErrorAction SilentlyContinue
<#
.SYNOPSIS
    WPF GUI application for processing school transport contracts using Google Maps.
    Designed for direct execution in PowerShell and standalone EXE compilation via PS2EXE.

.DESCRIPTION
    - Secure Google Maps API key storage (Windows DPAPI per-user encryption)
    - Automatic detection and installation of the ImportExcel module (Scope: CurrentUser)
    - Required column reference (Dom, Szkola, Praca, Umowa, Opis) and live header validation
    - Remembers last used output folder and user settings in local config file
    - Responsive WPF Dark Mode interface with real-time progress bar and event log (background worker)
    - 100% PS2EXE compatible (produces standalone single-file EXE)

.NOTES
    Encoding: UTF-8 with BOM
#>

# ── Force STA mode for WPF ───────────────────────────────────────────────────
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) {
    $currentProcess = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    if ($currentProcess -match 'powershell\.exe|pwsh\.exe') {
        Start-Process -FilePath $currentProcess -ArgumentList "-NoProfile -STA -ExecutionPolicy Bypass -File `"$PSCommandPath`""
        exit
    }
}

# ── Import GUI and Drawing assemblies ─────────────────────────────────────────
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms, System.Drawing, System.Security

# ── P/Invoke for Windows DWM Dark Mode window title bar ───────────────────────
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class DwmDarkWindow {
    [DllImport("dwmapi.dll", PreserveSig = true)]
    public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
}
"@ -ErrorAction SilentlyContinue

# ══════════════════════════════════════════════════════════════════════════════
# CONFIGURATION AND DPAPI SECURITY
# ══════════════════════════════════════════════════════════════════════════════

$script:AppDirName = 'SchoolTransportRoutes'
$script:LocalConfigFolder = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) $script:AppDirName
$script:ConfigFilePath = if (Test-Path "$PSScriptRoot\config.json") {
    "$PSScriptRoot\config.json"
} else {
    Join-Path $script:LocalConfigFolder 'config.json'
}

function Protect-SecretString {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$PlainText)
    if ([string]::IsNullOrEmpty($PlainText)) { return $null }
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($PlainText)
        $protected = [System.Security.Cryptography.ProtectedData]::Protect(
            $bytes, $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
        return [Convert]::ToBase64String($protected)
    } catch {
        # Fallback for restricted environments
        $sec = ConvertTo-SecureString -String $PlainText -AsPlainText -Force
        return (ConvertFrom-SecureString -SecureString $sec)
    }
}

function Unprotect-SecretString {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$EncryptedText)
    if ([string]::IsNullOrWhiteSpace($EncryptedText)) { return $null }
    try {
        $bytes = [Convert]::FromBase64String($EncryptedText)
        $unprotected = [System.Security.Cryptography.ProtectedData]::Unprotect(
            $bytes, $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
        return [System.Text.Encoding]::UTF8.GetString($unprotected)
    } catch {
        try {
            $sec = ConvertTo-SecureString -String $EncryptedText
            $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)
            $str = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
            return $str
        } catch {
            return $null
        }
    }
}

function Load-AppConfig {
    $defaultFolder = Join-Path $PSScriptRoot 'Results'
    $config = [PSCustomObject]@{
        ApiKey           = ''
        RememberApiKey   = $true
        LastOutputFolder = $defaultFolder
        LastExcelPath    = ''
        MapWidth         = 900
        MapHeight        = 600
    }

    if (Test-Path $script:ConfigFilePath) {
        try {
            $jsonContent = [System.IO.File]::ReadAllText($script:ConfigFilePath, [System.Text.Encoding]::UTF8)
            $raw = $jsonContent | ConvertFrom-Json
            if ($raw.ApiKeyEncrypted) {
                $decrypted = Unprotect-SecretString -EncryptedText $raw.ApiKeyEncrypted
                if ($decrypted) { $config.ApiKey = $decrypted }
            }
            if ($null -ne $raw.RememberApiKey) { $config.RememberApiKey = [bool]$raw.RememberApiKey }
            if ($raw.LastOutputFolder) { $config.LastOutputFolder = [string]$raw.LastOutputFolder }
            if ($raw.LastExcelPath) { $config.LastExcelPath = [string]$raw.LastExcelPath }
            if ($raw.MapWidth) { $config.MapWidth = [int]$raw.MapWidth }
            if ($raw.MapHeight) { $config.MapHeight = [int]$raw.MapHeight }
        } catch {}
    }

    # If key is not in config, check environment variable
    if ([string]::IsNullOrWhiteSpace($config.ApiKey) -and -not [string]::IsNullOrWhiteSpace($env:GOOGLE_MAPS_API_KEY)) {
        $config.ApiKey = $env:GOOGLE_MAPS_API_KEY
    }

    return $config
}

function Save-AppConfig {
    [CmdletBinding()]
    param(
        [string]$ApiKey,
        [bool]$RememberApiKey,
        [string]$OutputFolder,
        [string]$ExcelPath = '',
        [int]$MapWidth = 900,
        [int]$MapHeight = 600
    )
    try {
        $dir = Split-Path -Path $script:ConfigFilePath -Parent
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

        $encryptedKey = if ($RememberApiKey -and -not [string]::IsNullOrWhiteSpace($ApiKey)) {
            Protect-SecretString -PlainText $ApiKey
        } else {
            $null
        }

        $obj = [ordered]@{
            ApiKeyEncrypted  = $encryptedKey
            RememberApiKey   = $RememberApiKey
            LastOutputFolder = $OutputFolder
            LastExcelPath    = $ExcelPath
            MapWidth         = $MapWidth
            MapHeight        = $MapHeight
            LastUpdated      = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        }
        $json = $obj | ConvertTo-Json -Depth 5
        [System.IO.File]::WriteAllText($script:ConfigFilePath, $json, [System.Text.Encoding]::UTF8)
    } catch {}
}

# ══════════════════════════════════════════════════════════════════════════════
# IMPORTEXCEL MODULE MANAGEMENT
# ══════════════════════════════════════════════════════════════════════════════

function Test-ImportExcelAvailable {
    return [bool](Get-Module -ListAvailable -Name ImportExcel)
}

function Install-ImportExcelModule {
    param([System.Action[string]]$OnLog)
    
    if ($OnLog) { $OnLog.Invoke("Configuring TLS 1.2 protocols for PSGallery...") }
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls

    if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
        if ($OnLog) { $OnLog.Invoke("Installing NuGet package provider for CurrentUser...") }
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -ErrorAction Stop | Out-Null
    }

    if ($OnLog) { $OnLog.Invoke("Downloading and installing ImportExcel from PSGallery (Scope: CurrentUser)...") }
    Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber -Repository PSGallery -ErrorAction Stop
    Import-Module -Name ImportExcel -Force -ErrorAction Stop | Out-Null
    if ($OnLog) { $OnLog.Invoke("ImportExcel module successfully installed!") }
}

# ══════════════════════════════════════════════════════════════════════════════
# ROUTE AND GEOCODING CORE FUNCTIONS (PS2EXE COMPLIANT)
# ══════════════════════════════════════════════════════════════════════════════

function ConvertTo-SafeFileName {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Name)
    $InvalidChars = [System.IO.Path]::GetInvalidFileNameChars()
    $Safe = $Name
    foreach ($Char in $InvalidChars) {
        $Safe = $Safe.Replace([string]$Char, '-')
    }
    return $Safe.Trim()
}

function Test-PracaAddress {
    [CmdletBinding()]
    param([Parameter()][string]$Praca)
    if ([string]::IsNullOrWhiteSpace($Praca)) { return $false }
    $Normalized = $Praca.Trim().ToLower()
    if ($Normalized -in @('nie dotyczy', 'nd', 'n/d', 'nd.', '-', 'brak', 'none', 'n/a', 'na')) { return $false }
    return $true
}

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
            if (-not [string]::IsNullOrWhiteSpace($Value)) { return $Value }
        }
    }
    return $null
}

function Get-AddressCoordinates {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Address,
        [Parameter(Mandatory)][string]$ApiKey
    )
    if ([string]::IsNullOrWhiteSpace($Address)) { return $null }
    $EncodedAddress = [System.Uri]::EscapeDataString($Address.Trim())
    $Url = "https://maps.googleapis.com/maps/api/geocode/json?address=$EncodedAddress&language=pl&key=$ApiKey"
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

            $FormattedAddress = $ResultItem.formatted_address -replace ',\s*Poland$', ', Polska' -replace '\bPoland\b', 'Polska'

            return [PSCustomObject]@{
                Latitude         = [double]$Location.lat
                Longitude        = [double]$Location.lng
                FormattedAddress = $FormattedAddress
                UlicaINumer      = $StreetWithNumber
                KodPocztowy      = $PostalCode
                Miasto           = $City
                Status           = 'OK'
            }
        }
        else {
            return [PSCustomObject]@{
                Latitude     = $null; Longitude = $null; FormattedAddress = $null
                UlicaINumer  = $null; KodPocztowy = $null; Miasto = $null
                Status       = $Response.status
                ErrorMessage = $Response.error_message
            }
        }
    }
    catch {
        return [PSCustomObject]@{
            Latitude     = $null; Longitude = $null; FormattedAddress = $null
            UlicaINumer  = $null; KodPocztowy = $null; Miasto = $null
            Status       = "EXCEPTION: $($_.Exception.Message)"
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Get-CarRouteData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][double]$OriginLat,
        [Parameter(Mandatory)][double]$OriginLng,
        [Parameter(Mandatory)][double]$DestLat,
        [Parameter(Mandatory)][double]$DestLng,
        [Parameter(Mandatory)][string]$ApiKey
    )
    $RoutesUrl = 'https://routes.googleapis.com/directions/v2:computeRoutes'
    $RequestBody = @{
        origin                   = @{ location = @{ latLng = @{ latitude = $OriginLat; longitude = $OriginLng } } }
        destination              = @{ location = @{ latLng = @{ latitude = $DestLat; longitude = $DestLng } } }
        travelMode               = 'DRIVE'
        routingPreference        = 'TRAFFIC_UNAWARE'
        computeAlternativeRoutes = $true
        languageCode             = 'pl'
        units                    = 'METRIC'
    }
    $Headers = @{
        'X-Goog-Api-Key'   = $ApiKey
        'Content-Type'     = 'application/json'
        'X-Goog-FieldMask' = 'routes.duration,routes.distanceMeters,routes.polyline.encodedPolyline'
    }
    try {
        $Response = Invoke-RestMethod -Uri $RoutesUrl -Method POST -Headers $Headers `
            -Body ($RequestBody | ConvertTo-Json -Depth 10) -TimeoutSec 60
        $Routes = @($Response.routes)
        if ($Routes.Count -eq 0) {
            return [PSCustomObject]@{ OdlegloscKm = $null; CzasMin = $null; EncodedPolyline = $null; Status = 'NO_ROUTES' }
        }
        $Route = $Routes | Sort-Object -Property distanceMeters | Select-Object -First 1
        $DistanceKm = if ($Route.distanceMeters) { [math]::Round($Route.distanceMeters / 1000.0, 2) } else { $null }
        $DurationMinutes = $null
        if ($Route.duration) {
            $Seconds = [double]($Route.duration.TrimEnd('s'))
            $DurationMinutes = [math]::Round($Seconds / 60.0, 0)
        }
        $Polyline = if ($Route.polyline) { $Route.polyline.encodedPolyline } else { $null }
        return [PSCustomObject]@{ OdlegloscKm = $DistanceKm; CzasMin = $DurationMinutes; EncodedPolyline = $Polyline; Status = 'OK' }
    }
    catch {
        return [PSCustomObject]@{ OdlegloscKm = $null; CzasMin = $null; EncodedPolyline = $null; Status = "EXCEPTION: $($_.Exception.Message)" }
    }
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
        [Parameter()][string]$TekstAdresA = '',
        [Parameter()][string]$TekstAdresB = '',
        [Parameter()][string]$TekstOdleglosc = '',
        [Parameter()][string]$TekstUmowa = '',
        [Parameter()][string]$TekstKierunek = '',
        [Parameter()][string]$TekstNaglowekLewy = '',
        [Parameter()][string]$TekstNaglowekPrawy = '',
        [Parameter()][string]$Opis = '',
        [Parameter()][string]$DataWygenerowania = ''
    )
    $EncodedForUrl = [System.Uri]::EscapeDataString($EncodedPolyline)
    $MarkerStart = [System.Uri]::EscapeDataString("color:green|label:A|$OriginLat,$OriginLng")
    $MarkerEnd = [System.Uri]::EscapeDataString("color:red|label:B|$DestLat,$DestLng")
    $StaticMapUrl = ("https://maps.googleapis.com/maps/api/staticmap" +
        "?size=${Width}x${Height}" +
        "&path=weight:4|color:0x0066FFff|enc:$EncodedForUrl" +
        "&markers=$MarkerStart" +
        "&markers=$MarkerEnd" +
        "&key=$ApiKey")
    try {
        Invoke-WebRequest -Uri $StaticMapUrl -OutFile $OutputPath -TimeoutSec 30 | Out-Null

        # Header Left: Opis (preferred), fallback to NumerUmowy (no 'Contract: ' prefix)
        if ([string]::IsNullOrWhiteSpace($TekstNaglowekLewy)) {
            if (-not [string]::IsNullOrWhiteSpace($Opis)) {
                $TekstNaglowekLewy = $Opis.Trim()
            } elseif (-not [string]::IsNullOrWhiteSpace($TekstUmowa)) {
                $TekstNaglowekLewy = $TekstUmowa.Trim()
            }
        }

        # Header Right: Date of generation (upper right corner)
        if ([string]::IsNullOrWhiteSpace($TekstNaglowekPrawy)) {
            if ([string]::IsNullOrWhiteSpace($DataWygenerowania)) {
                $DataWygenerowania = (Get-Date).ToString('yyyy-MM-dd')
            }
            $TekstNaglowekPrawy = "Data: $DataWygenerowania"
        }

        $MaTopOverlay = (-not [string]::IsNullOrWhiteSpace($TekstNaglowekLewy)) -or
                        (-not [string]::IsNullOrWhiteSpace($TekstNaglowekPrawy))

        $MaBottomOverlay = (-not [string]::IsNullOrWhiteSpace($TekstAdresA)) -or
                           (-not [string]::IsNullOrWhiteSpace($TekstAdresB)) -or
                           (-not [string]::IsNullOrWhiteSpace($TekstOdleglosc))

        if ($MaBottomOverlay -or $MaTopOverlay) {
            try {
                $FileBytes = [System.IO.File]::ReadAllBytes($OutputPath)
                $MemStream = [System.IO.MemoryStream]::new($FileBytes)
                $BitmapSrc = [System.Drawing.Bitmap]::new($MemStream)

                $ActualW = $BitmapSrc.Width
                $ActualH = $BitmapSrc.Height

                # Calculate extra canvas space for top and bottom extensions
                $TopBarH = if ($MaTopOverlay) { 36 } else { 0 }
                $BtmBarH = if ($MaBottomOverlay) { 60 } else { 0 }
                $FinalW  = $ActualW
                $FinalH  = $ActualH + $TopBarH + $BtmBarH

                # Create enlarged canvas
                $Bitmap = [System.Drawing.Bitmap]::new($FinalW, $FinalH, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
                $Graphics = [System.Drawing.Graphics]::FromImage($Bitmap)
                $Graphics.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
                $Graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit

                # 1. Fill canvas with solid clean white
                $BrushWhite = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::White)
                $Graphics.FillRectangle($BrushWhite, 0, 0, $FinalW, $FinalH)

                # 2. Draw map in the middle — untouched and 100% visible
                $Graphics.DrawImage($BitmapSrc, 0, $TopBarH, $ActualW, $ActualH)

                # 3. Separator lines
                $PenBorder = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(210, 220, 230), 1)
                if ($TopBarH -gt 0) {
                    $Graphics.DrawLine($PenBorder, 0, $TopBarH, $FinalW, $TopBarH)
                }
                if ($BtmBarH -gt 0) {
                    $BtmBarY = $TopBarH + $ActualH
                    $Graphics.DrawLine($PenBorder, 0, $BtmBarY, $FinalW, $BtmBarY)
                }

                # 4. Top banner (Opis on left, Date of generation on upper right)
                if ($MaTopOverlay) {
                    $FontTopTitle = [System.Drawing.Font]::new('Segoe UI', 11.0, [System.Drawing.FontStyle]::Bold)
                    $FontTopDate  = [System.Drawing.Font]::new('Segoe UI', 9.5,  [System.Drawing.FontStyle]::Regular)
                    $BrushDark    = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 15, 23, 42))
                    $BrushGray    = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 100, 116, 139))

                    if (-not [string]::IsNullOrWhiteSpace($TekstNaglowekLewy)) {
                        $Graphics.DrawString($TekstNaglowekLewy, $FontTopTitle, $BrushDark, 14.0, 7.0)
                    }
                    if (-not [string]::IsNullOrWhiteSpace($TekstNaglowekPrawy)) {
                        $SizePrawego = $Graphics.MeasureString($TekstNaglowekPrawy, $FontTopDate)
                        $XPrawy = [math]::Max(0.0, [double]($FinalW - $SizePrawego.Width - 14.0))
                        $Graphics.DrawString($TekstNaglowekPrawy, $FontTopDate, $BrushGray, [single]$XPrawy, 9.0)
                    }

                    $FontTopTitle.Dispose()
                    $FontTopDate.Dispose()
                    $BrushDark.Dispose()
                    $BrushGray.Dispose()
                }

                # 5. Bottom banner (Address [A] on line 1, Address [B] on line 2, Total km in right corner)
                if ($MaBottomOverlay) {
                    $BtmBarY = $TopBarH + $ActualH
                    $FontBadge      = [System.Drawing.Font]::new('Segoe UI', 8.5, [System.Drawing.FontStyle]::Bold)
                    $FontAddr       = [System.Drawing.Font]::new('Segoe UI', 9.0, [System.Drawing.FontStyle]::Regular)
                    $FontDistLabel  = [System.Drawing.Font]::new('Segoe UI', 7.5, [System.Drawing.FontStyle]::Regular)
                    $FontDistValue  = [System.Drawing.Font]::new('Segoe UI', 13.5, [System.Drawing.FontStyle]::Bold)

                    $BrushDark      = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 15, 23, 42))
                    $BrushGray      = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 100, 116, 139))
                    $BrushBadgeA    = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 22, 163, 74))
                    $BrushBadgeB    = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 220, 38, 38))
                    $BrushDist      = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 2, 132, 199))

                    # Distance display in bottom right corner
                    $distDisplay = if ($TekstOdleglosc -match 'km\s*$') {
                        $TekstOdleglosc.Trim()
                    } elseif (-not [string]::IsNullOrWhiteSpace($TekstOdleglosc)) {
                        "$($TekstOdleglosc.Trim()) km"
                    } else {
                        ''
                    }

                    $distValSize = if ($distDisplay) { $Graphics.MeasureString($distDisplay, $FontDistValue) } else { [System.Drawing.SizeF]::Empty }
                    $distLblSize = if ($distDisplay) { $Graphics.MeasureString('Dystans:', $FontDistLabel) } else { [System.Drawing.SizeF]::Empty }
                    $distBoxW    = [math]::Max($distValSize.Width, $distLblSize.Width)
                    $distRightX  = [single]($FinalW - 14.0)

                    if (-not [string]::IsNullOrWhiteSpace($distDisplay)) {
                        $Graphics.DrawString('Dystans:', $FontDistLabel, $BrushGray, [single]($distRightX - $distLblSize.Width), [single]($BtmBarY + 8.0))
                        $Graphics.DrawString($distDisplay, $FontDistValue, $BrushDist, [single]($distRightX - $distValSize.Width), [single]($BtmBarY + 23.0))
                    }

                    # Address [A] and Address [B] calculation (stacked on 2 lines, cannot trim or exceed frame)
                    $reservedRight = if ($distBoxW -gt 0) { $distBoxW + 28.0 } else { 14.0 }
                    $maxAddrW = [single]([math]::Max(160.0, ($FinalW - 40.0 - $reservedRight)))

                    $fmt = [System.Drawing.StringFormat]::new()
                    $fmt.Trimming = [System.Drawing.StringTrimming]::EllipsisCharacter
                    $fmt.FormatFlags = [System.Drawing.StringFormatFlags]::NoWrap

                    # Line 1: [A] + Address A
                    if (-not [string]::IsNullOrWhiteSpace($TekstAdresA)) {
                        $Graphics.DrawString('[A]', $FontBadge, $BrushBadgeA, 14.0, [single]($BtmBarY + 8.0))
                        $fontA = $FontAddr
                        if ($Graphics.MeasureString($TekstAdresA, $FontAddr).Width -gt $maxAddrW) {
                            $fontA = [System.Drawing.Font]::new('Segoe UI', 8.0, [System.Drawing.FontStyle]::Regular)
                        }
                        $rectA = [System.Drawing.RectangleF]::new(40.0, [single]($BtmBarY + 8.0), $maxAddrW, 20.0)
                        $Graphics.DrawString($TekstAdresA, $fontA, $BrushDark, $rectA, $fmt)
                        if ($fontA -ne $FontAddr) { $fontA.Dispose() }
                    }

                    # Line 2: [B] + Address B
                    if (-not [string]::IsNullOrWhiteSpace($TekstAdresB)) {
                        $Graphics.DrawString('[B]', $FontBadge, $BrushBadgeB, 14.0, [single]($BtmBarY + 32.0))
                        $fontB = $FontAddr
                        if ($Graphics.MeasureString($TekstAdresB, $FontAddr).Width -gt $maxAddrW) {
                            $fontB = [System.Drawing.Font]::new('Segoe UI', 8.0, [System.Drawing.FontStyle]::Regular)
                        }
                        $rectB = [System.Drawing.RectangleF]::new(40.0, [single]($BtmBarY + 32.0), $maxAddrW, 20.0)
                        $Graphics.DrawString($TekstAdresB, $fontB, $BrushDark, $rectB, $fmt)
                        if ($fontB -ne $FontAddr) { $fontB.Dispose() }
                    }

                    $fmt.Dispose()
                    $FontBadge.Dispose()
                    $FontAddr.Dispose()
                    $FontDistLabel.Dispose()
                    $FontDistValue.Dispose()
                    $BrushDark.Dispose()
                    $BrushGray.Dispose()
                    $BrushBadgeA.Dispose()
                    $BrushBadgeB.Dispose()
                    $BrushDist.Dispose()
                }

                $BrushWhite.Dispose()
                $PenBorder.Dispose()
                $Graphics.Dispose()

                $Bitmap.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
                $Bitmap.Dispose()
                $BitmapSrc.Dispose()
                $MemStream.Dispose()
            } catch {}
        }
        return $true
    }
    catch {
        return $false
    }
}

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
        return [PSCustomObject]@{ OdlegloscKm = $null; Status = $Trasa.Status; MapSaved = $false }
    }

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
            -TekstAdresA $TekstA -TekstAdresB $TekstB -TekstOdleglosc $OdlTekst `
            -TekstUmowa $NumerUmowy -TekstKierunek $KierunekTekst `
            -Opis $Opis -DataWygenerowania $DataWygenerowania
    }

    return [PSCustomObject]@{ OdlegloscKm = $Trasa.OdlegloscKm; Status = $Trasa.Status; MapSaved = $MapSaved }
}

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

# ══════════════════════════════════════════════════════════════════════════════
# WPF XAML USER INTERFACE DEFINITION (MODERN DARK THEME)
# ══════════════════════════════════════════════════════════════════════════════

[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="School Transport Route &amp; Map Generator — Google Maps"
    Height="880" Width="960"
    MinHeight="720" MinWidth="820"
    WindowStartupLocation="CenterScreen"
    Background="#0F172A" Foreground="#F8FAFC"
    FontFamily="Segoe UI Variable, Segoe UI, sans-serif">

    <Window.Resources>
        <!-- Styles for controls -->
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="#1E293B"/>
            <Setter Property="Foreground" Value="#F8FAFC"/>
            <Setter Property="BorderBrush" Value="#334155"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="10,7"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>

        <Style TargetType="PasswordBox">
            <Setter Property="Background" Value="#1E293B"/>
            <Setter Property="Foreground" Value="#F8FAFC"/>
            <Setter Property="BorderBrush" Value="#334155"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="10,7"/>
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
        </Style>

        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="#CBD5E1"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>
    </Window.Resources>

    <Grid Margin="20">
        <Grid.RowDefinitions>
            <!-- 0: Header -->
            <RowDefinition Height="Auto"/>
            <!-- 1: Card 1 - Google Maps API Key -->
            <RowDefinition Height="Auto"/>
            <!-- 2: Card 2 - Excel Input File and Columns Guide -->
            <RowDefinition Height="Auto"/>
            <!-- 3: Card 3 - Output Folder and Map Dimensions -->
            <RowDefinition Height="Auto"/>
            <!-- 4: Card 4 - Action Buttons and Progress Bar -->
            <RowDefinition Height="Auto"/>
            <!-- 5: Card 5 - Log Console -->
            <RowDefinition Height="*"/>
            <!-- 6: Bottom Status Bar -->
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- HEADER -->
        <Border Grid.Row="0" Margin="0,0,0,15" Padding="0,0,0,10" BorderBrush="#1E293B" BorderThickness="0,0,0,1">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0">
                    <StackPanel Orientation="Horizontal">
                        <TextBlock Text="🗺️" FontSize="24" Margin="0,0,10,0" VerticalAlignment="Center"/>
                        <TextBlock Text="School Transport Route &amp; Map Generator" FontSize="20" FontWeight="Bold" Foreground="#38BDF8" VerticalAlignment="Center"/>
                    </StackPanel>
                    <TextBlock Text="Automatic Google Maps route calculation, PNG map generation, and Excel summary reports" 
                               FontSize="12" Foreground="#94A3B8" Margin="34,2,0,0"/>
                </StackPanel>
                <Border Grid.Column="1" Background="#1E293B" CornerRadius="6" Padding="10,4" VerticalAlignment="Center">
                    <TextBlock Text="Version 1.0 (PS2EXE Ready)" FontSize="11" Foreground="#38BDF8" FontWeight="SemiBold"/>
                </Border>
            </Grid>
        </Border>

        <!-- CARD 1: GOOGLE MAPS API KEY -->
        <Border Grid.Row="1" Background="#1E293B" CornerRadius="8" Padding="15" Margin="0,0,0,12" BorderBrush="#334155" BorderThickness="1">
            <StackPanel>
                <Grid Margin="0,0,0,8">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <StackPanel Orientation="Horizontal" Grid.Column="0">
                        <TextBlock Text="🔑" FontSize="14" Margin="0,0,6,0" VerticalAlignment="Center"/>
                        <TextBlock Text="Google Maps API Key" FontWeight="Bold" FontSize="13" Foreground="#F8FAFC" VerticalAlignment="Center"/>
                        <TextBlock Text=" (Required for Geocoding, Routes v2, and Static Maps)" FontSize="11" Foreground="#94A3B8" VerticalAlignment="Center" Margin="5,0,0,0"/>
                    </StackPanel>
                    <Button Name="btnToggleMask" Grid.Column="1" Content="👁 Show Key" Background="#334155" Foreground="#E2E8F0" Padding="8,4" FontSize="11"/>
                </Grid>

                <Grid Margin="0,0,0,8">
                    <PasswordBox Name="txtApiKeyPass"/>
                    <TextBox Name="txtApiKeyText" Visibility="Collapsed"/>
                </Grid>

                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <CheckBox Name="chkRememberApiKey" Grid.Column="0" IsChecked="True" Content="Remember API key securely in local profile (Windows DPAPI encryption)"/>
                    <TextBlock Name="lblKeyStatus" Grid.Column="1" Text="" FontSize="11" Foreground="#10B981" FontWeight="SemiBold"/>
                </Grid>
            </StackPanel>
        </Border>

        <!-- CARD 2: EXCEL INPUT FILE AND COLUMN GUIDE -->
        <Border Grid.Row="2" Background="#1E293B" CornerRadius="8" Padding="15" Margin="0,0,0,12" BorderBrush="#334155" BorderThickness="1">
            <StackPanel>
                <Grid Margin="0,0,0,8">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <StackPanel Orientation="Horizontal" Grid.Column="0">
                        <TextBlock Text="📊" FontSize="14" Margin="0,0,6,0" VerticalAlignment="Center"/>
                        <TextBlock Text="Input Excel File with Contracts" FontWeight="Bold" FontSize="13" Foreground="#F8FAFC" VerticalAlignment="Center"/>
                    </StackPanel>
                    <Button Name="btnBrowseExcel" Grid.Column="1" Content="📁 Select Excel File..." Background="#2563EB" Foreground="#FFFFFF" Padding="12,5" FontSize="12"/>
                </Grid>

                <TextBox Name="txtExcelPath" Margin="0,0,0,10" ToolTip="Path to the Excel file. You can also drag and drop a file directly onto this window."/>

                <!-- REQUIRED COLUMNS GUIDE BANNER -->
                <Border Background="#0F172A" CornerRadius="6" Padding="12,10" BorderBrush="#334155" BorderThickness="1">
                    <StackPanel>
                        <TextBlock Text="ℹ️ Required column structure in the Excel sheet:" FontWeight="Bold" FontSize="12" Foreground="#38BDF8" Margin="0,0,0,6"/>
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <StackPanel Grid.Column="0" Margin="0,0,10,0">
                                <TextBlock Text="• Dom — Home address (required for route)" FontSize="11" Foreground="#E2E8F0" Margin="0,1"/>
                                <TextBlock Text="• Szkola — School address (required for route)" FontSize="11" Foreground="#E2E8F0" Margin="0,1"/>
                                <TextBlock Text="• Praca — Work address (optional, combined routes)" FontSize="11" Foreground="#94A3B8" Margin="0,1"/>
                            </StackPanel>
                            <StackPanel Grid.Column="1">
                                <TextBlock Text="• Umowa — Contract number (written on PNG map header)" FontSize="11" Foreground="#E2E8F0" Margin="0,1"/>
                                <TextBlock Text="• Opis — Used as subfolder name and PNG/Excel file names" FontSize="11" Foreground="#FDE047" Margin="0,1"/>
                            </StackPanel>
                        </Grid>

                        <!-- Live Column Header Status Bar -->
                        <Border Name="borderColStatus" Background="#1E293B" CornerRadius="4" Padding="8,5" Margin="0,8,0,0" Visibility="Collapsed">
                            <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                                <TextBlock Text="Detected columns:" FontSize="11" FontWeight="Bold" Foreground="#94A3B8" Margin="0,0,8,0"/>
                                <TextBlock Name="badgeUmowa" Text="Umowa: ?" FontSize="11" Margin="0,0,10,0"/>
                                <TextBlock Name="badgeOpis" Text="Opis: ?" FontSize="11" Margin="0,0,10,0"/>
                                <TextBlock Name="badgeDom" Text="Dom: ?" FontSize="11" Margin="0,0,10,0"/>
                                <TextBlock Name="badgeSzkola" Text="Szkola: ?" FontSize="11" Margin="0,0,10,0"/>
                                <TextBlock Name="badgePraca" Text="Praca: ?" FontSize="11"/>
                            </StackPanel>
                        </Border>
                    </StackPanel>
                </Border>
            </StackPanel>
        </Border>

        <!-- CARD 3: OUTPUT FOLDER AND MAP DIMENSIONS -->
        <Border Grid.Row="3" Background="#1E293B" CornerRadius="8" Padding="15" Margin="0,0,0,12" BorderBrush="#334155" BorderThickness="1">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="220"/>
                </Grid.ColumnDefinitions>

                <StackPanel Grid.Column="0" Margin="0,0,10,0">
                    <TextBlock Text="📂 Output Folder (saved in configuration)" FontWeight="Bold" FontSize="13" Foreground="#F8FAFC" Margin="0,0,0,6"/>
                    <TextBox Name="txtOutputFolder"/>
                </StackPanel>

                <Button Name="btnBrowseOutput" Grid.Column="1" Content="Browse..." Background="#334155" Foreground="#E2E8F0" VerticalAlignment="Bottom" Height="36" Margin="0,0,15,0" Padding="12,0"/>

                <StackPanel Grid.Column="2">
                    <TextBlock Text="📐 PNG Map Dimensions" FontWeight="Bold" FontSize="13" Foreground="#F8FAFC" Margin="0,0,0,6"/>
                    <ComboBox Name="cmbMapDimensions" Height="36" Background="#1E293B" Foreground="#0F172A" FontSize="12">
                        <ComboBoxItem Content="900 x 600 px (Recommended)" IsSelected="True" Tag="900x600"/>
                        <ComboBoxItem Content="1024 x 768 px" Tag="1024x768"/>
                        <ComboBoxItem Content="1200 x 800 px" Tag="1200x800"/>
                        <ComboBoxItem Content="800 x 500 px" Tag="800x500"/>
                    </ComboBox>
                </StackPanel>
            </Grid>
        </Border>

        <!-- CARD 4: ACTION BUTTONS AND PROGRESS -->
        <Border Grid.Row="4" Background="#1E293B" CornerRadius="8" Padding="15" Margin="0,0,0,12" BorderBrush="#334155" BorderThickness="1">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <!-- Action Buttons -->
                <Grid Grid.Row="0" Margin="0,0,0,12">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>

                    <Button Name="btnStart" Grid.Column="0" Content="▶ Start Processing Routes &amp; Maps" Background="#059669" Foreground="#FFFFFF" Height="42" FontSize="14" FontWeight="Bold" Margin="0,0,10,0"/>
                    <Button Name="btnStop" Grid.Column="1" Content="⏹ Stop" Background="#DC2626" Foreground="#FFFFFF" Height="42" FontSize="13" Width="110" Margin="0,0,10,0" IsEnabled="False"/>
                    <Button Name="btnOpenOutput" Grid.Column="2" Content="📁 Open Output Folder" Background="#334155" Foreground="#E2E8F0" Height="42" FontSize="12"/>
                </Grid>

                <!-- Progress Bar and Status -->
                <StackPanel Grid.Row="1">
                    <Grid Margin="0,0,0,4">
                        <TextBlock Name="lblProgressText" Text="Ready to start" FontSize="12" Foreground="#CBD5E1"/>
                        <TextBlock Name="lblProgressPercent" Text="0%" HorizontalAlignment="Right" FontSize="12" FontWeight="Bold" Foreground="#38BDF8"/>
                    </Grid>
                    <ProgressBar Name="progBar" Height="14" Minimum="0" Maximum="100" Value="0" Background="#0F172A" Foreground="#38BDF8"/>
                </StackPanel>
            </Grid>
        </Border>

        <!-- CARD 5: EVENT LOG (CONSOLE) -->
        <Border Grid.Row="5" Background="#0B0F17" CornerRadius="8" Padding="10" BorderBrush="#1E293B" BorderThickness="1" MinHeight="140">
            <TextBox Name="txtLog" Background="Transparent" Foreground="#E2E8F0" BorderThickness="0" 
                     FontFamily="Cascadia Mono, Consolas, Courier New" FontSize="11" 
                     IsReadOnly="True" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"
                     TextWrapping="NoWrap" AcceptsReturn="True" MinHeight="120"/>
        </Border>

        <!-- BOTTOM STATUS BAR -->
        <Grid Grid.Row="6" Margin="0,8,0,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBlock Name="lblStatusModule" Text="ImportExcel Module: Checking..." FontSize="11" Foreground="#94A3B8"/>
            <StackPanel Grid.Column="1" Orientation="Horizontal">
                <TextBlock Name="lblStatsProcessed" Text="Processed: 0" FontSize="11" Foreground="#94A3B8" Margin="0,0,15,0"/>
                <TextBlock Name="lblStatsSuccess" Text="OK: 0" FontSize="11" Foreground="#10B981" Margin="0,0,15,0"/>
                <TextBlock Name="lblStatsError" Text="Errors: 0" FontSize="11" Foreground="#EF4444"/>
            </StackPanel>
        </Grid>
    </Grid>
</Window>
"@

# ══════════════════════════════════════════════════════════════════════════════
# WINDOW AND CONTROLS INITIALIZATION
# ══════════════════════════════════════════════════════════════════════════════

$reader = [System.Xml.XmlNodeReader]::new($xaml)
$window = [System.Windows.Markup.XamlReader]::Load($reader)

# Retrieve controls by name
$txtApiKeyPass       = $window.FindName('txtApiKeyPass')
$txtApiKeyText       = $window.FindName('txtApiKeyText')
$btnToggleMask       = $window.FindName('btnToggleMask')
$chkRememberApiKey   = $window.FindName('chkRememberApiKey')
$lblKeyStatus        = $window.FindName('lblKeyStatus')

$btnBrowseExcel      = $window.FindName('btnBrowseExcel')
$txtExcelPath        = $window.FindName('txtExcelPath')
$borderColStatus     = $window.FindName('borderColStatus')
$badgeUmowa          = $window.FindName('badgeUmowa')
$badgeOpis           = $window.FindName('badgeOpis')
$badgeDom            = $window.FindName('badgeDom')
$badgeSzkola         = $window.FindName('badgeSzkola')
$badgePraca          = $window.FindName('badgePraca')

$txtOutputFolder     = $window.FindName('txtOutputFolder')
$btnBrowseOutput     = $window.FindName('btnBrowseOutput')
$cmbMapDimensions    = $window.FindName('cmbMapDimensions')

$btnStart            = $window.FindName('btnStart')
$btnStop             = $window.FindName('btnStop')
$btnOpenOutput       = $window.FindName('btnOpenOutput')

$lblProgressText     = $window.FindName('lblProgressText')
$lblProgressPercent  = $window.FindName('lblProgressPercent')
$progBar             = $window.FindName('progBar')

$txtLog              = $window.FindName('txtLog')
$lblStatusModule     = $window.FindName('lblStatusModule')
$lblStatsProcessed   = $window.FindName('lblStatsProcessed')
$lblStatsSuccess     = $window.FindName('lblStatsSuccess')
$lblStatsError       = $window.FindName('lblStatsError')

# Global state variables
$script:IsKeyMasked = $true
$script:CancellationRequested = $false
$script:IsWorking = $false
$script:CurrentExcelInspection = $null

# ── GUI helper functions ──────────────────────────────────────────────────────
function Write-GuiLog {
    param([string]$Message, [string]$Level = 'INFO')
    try {
        $time = (Get-Date).ToString('HH:mm:ss')
        $prefix = switch ($Level) {
            'OK'    { '[OK]    ' }
            'WARN'  { '[WARN]  ' }
            'ERROR' { '[ERROR] ' }
            default { '[INFO]  ' }
        }
        $line = "$time $prefix $Message`r`n"
        if ($window -and $window.Dispatcher -and $txtLog) {
            if ($window.Dispatcher.CheckAccess()) {
                $txtLog.AppendText($line)
                $txtLog.ScrollToEnd()
            } else {
                $window.Dispatcher.Invoke([Action]{
                    if ($txtLog) {
                        $txtLog.AppendText($line)
                        $txtLog.ScrollToEnd()
                    }
                })
            }
        }
    } catch { }
}

function Get-CurrentApiKey {
    if ($script:IsKeyMasked) {
        return $txtApiKeyPass.Password.Trim()
    } else {
        return $txtApiKeyText.Text.Trim()
    }
}

function Set-CurrentApiKey([string]$Key) {
    $txtApiKeyPass.Password = $Key
    $txtApiKeyText.Text = $Key
}

function Save-CurrentGuiConfig {
    try {
        $key = Get-CurrentApiKey
        $remember = [bool]$chkRememberApiKey.IsChecked
        $outFolder = if (-not [string]::IsNullOrWhiteSpace($txtOutputFolder.Text)) { $txtOutputFolder.Text.Trim() } else { '' }
        $excelPath = if (-not [string]::IsNullOrWhiteSpace($txtExcelPath.Text)) { $txtExcelPath.Text.Trim() } else { '' }
        $mW = 900; $mH = 600
        if ($cmbMapDimensions.SelectedItem) {
            $parts = ($cmbMapDimensions.SelectedItem.Tag -as [string]) -split 'x'
            if ($parts.Count -eq 2) { $mW = [int]$parts[0]; $mH = [int]$parts[1] }
        }
        Save-AppConfig -ApiKey $key -RememberApiKey $remember -OutputFolder $outFolder -ExcelPath $excelPath -MapWidth $mW -MapHeight $mH
    } catch { }
}

# ── Handle DWM Dark Mode for window title bar ─────────────────────────────────
$window.Add_SourceInitialized({
    try {
        $helper = New-Object System.Windows.Interop.WindowInteropHelper($window)
        $val = 1
        [DwmDarkWindow]::DwmSetWindowAttribute($helper.Handle, 20, [ref]$val, 4)
        [DwmDarkWindow]::DwmSetWindowAttribute($helper.Handle, 19, [ref]$val, 4)
    } catch {}
})

# ── Load saved configuration ──────────────────────────────────────────────────
$loadedConfig = Load-AppConfig
if (-not [string]::IsNullOrWhiteSpace($loadedConfig.ApiKey)) {
    Set-CurrentApiKey -Key $loadedConfig.ApiKey
    $lblKeyStatus.Text = '🔒 Securely loaded from DPAPI'
}
$chkRememberApiKey.IsChecked = $loadedConfig.RememberApiKey
$txtOutputFolder.Text = $loadedConfig.LastOutputFolder

if (-not [string]::IsNullOrWhiteSpace($loadedConfig.LastExcelPath)) {
    $txtExcelPath.Text = $loadedConfig.LastExcelPath
    if (Test-Path $loadedConfig.LastExcelPath) {
        Update-ExcelColumnStatus -FilePath $loadedConfig.LastExcelPath
    }
}

# Set map dimensions preset
$dimKey = "$($loadedConfig.MapWidth)x$($loadedConfig.MapHeight)"
foreach ($item in $cmbMapDimensions.Items) {
    if ($item.Tag -eq $dimKey) {
        $cmbMapDimensions.SelectedItem = $item
        break
    }
}

# ── Auto-save configuration on changes and window close ──────────────────────
$window.Add_Closing({
    Save-CurrentGuiConfig
})

$txtApiKeyPass.Add_PasswordChanged({
    if ($script:IsKeyMasked) {
        $txtApiKeyText.Text = $txtApiKeyPass.Password
    }
    if ($chkRememberApiKey.IsChecked) { Save-CurrentGuiConfig }
})

$txtApiKeyText.Add_TextChanged({
    if (-not $script:IsKeyMasked) {
        $txtApiKeyPass.Password = $txtApiKeyText.Text
    }
    if ($chkRememberApiKey.IsChecked) { Save-CurrentGuiConfig }
})

$chkRememberApiKey.Add_Click({
    Save-CurrentGuiConfig
    if ($chkRememberApiKey.IsChecked) {
        $lblKeyStatus.Text = '🔒 Securely saved to DPAPI'
    } else {
        $lblKeyStatus.Text = '⚠️ Key will not be remembered'
    }
})

$txtOutputFolder.Add_LostFocus({
    Save-CurrentGuiConfig
})

$cmbMapDimensions.Add_SelectionChanged({
    Save-CurrentGuiConfig
})

# ── Check ImportExcel module at startup ───────────────────────────────────────
if (Test-ImportExcelAvailable) {
    $lblStatusModule.Text = "ImportExcel Module: Ready"
    $lblStatusModule.Foreground = [System.Windows.Media.Brushes]::LightGreen
} else {
    $lblStatusModule.Text = "ImportExcel Module: Missing! (Will install automatically)"
    $lblStatusModule.Foreground = [System.Windows.Media.Brushes]::Yellow
}

# ── Toggle API key visibility ─────────────────────────────────────────────────
$btnToggleMask.Add_Click({
    if ($script:IsKeyMasked) {
        $txtApiKeyText.Text = $txtApiKeyPass.Password
        $txtApiKeyPass.Visibility = [System.Windows.Visibility]::Collapsed
        $txtApiKeyText.Visibility = [System.Windows.Visibility]::Visible
        $btnToggleMask.Content = '🔒 Hide Key'
        $script:IsKeyMasked = $false
    } else {
        $txtApiKeyPass.Password = $txtApiKeyText.Text
        $txtApiKeyText.Visibility = [System.Windows.Visibility]::Collapsed
        $txtApiKeyPass.Visibility = [System.Windows.Visibility]::Visible
        $btnToggleMask.Content = '👁 Show Key'
        $script:IsKeyMasked = $true
    }
})

# ── Validate Excel columns ────────────────────────────────────────────────────
function Update-ExcelColumnStatus([string]$FilePath) {
    if (-not (Test-Path $FilePath)) {
        $borderColStatus.Visibility = [System.Windows.Visibility]::Collapsed
        return
    }

    if (-not (Test-ImportExcelAvailable)) {
        $borderColStatus.Visibility = [System.Windows.Visibility]::Visible
        $badgeUmowa.Text = "ImportExcel not loaded"
        $badgeUmowa.Foreground = [System.Windows.Media.Brushes]::Yellow
        return
    }

    try {
        $firstRows = Import-Excel -Path $FilePath -EndRow 2 -ErrorAction Stop
        if (-not $firstRows) {
            $firstRows = Import-Excel -Path $FilePath -ErrorAction Stop
        }
        if (-not $firstRows -or @($firstRows).Count -eq 0) {
            Write-GuiLog "Excel file appears to be empty." "WARN"
            return
        }

        $headers = @($firstRows[0].PSObject.Properties.Name)

        $patterns = @{
            Umowa   = @('(?i)^\s*(umow[ay]|nr\s*umow[ay]|numer\s*umow[ay]|id|numer|nr|contract)\s*$', '(?i)umow|contract')
            Opis    = @('(?i)^\s*(opis|nazwa|nazwa\s*folderu|folder|identyfikator|description)\s*$', '(?i)opis|desc')
            Szkola  = @('(?i)^\s*(szko[łl][ay]|adres\s*szko[łl][yi]|plac[oó]wk[ay]|przedszkol[ea]|o[sś]rodek|szko[łl]a\s*adres|school)\s*$', '(?i)szko[łl]|plac[oó]wk|przedszkol|school')
            Dom     = @('(?i)^\s*(dom|domu|adres\s*domu|adres\s*zamieszkania|zamieszkani[ea]|miejsce\s*zamieszkania|dom\s*adres|home)\s*$', '(?i)dom|zamieszkan|home')
            Praca   = @('(?i)^\s*(prac[ay]|adres\s*prac[ay]|miejsce\s*prac[ay]|zak[łl]ad\s*prac[ay]|firma|praca\s*adres|work)\s*$', '(?i)prac|firm|work')
        }

        $colUmowa  = Find-ColumnHeader -Headers $headers -Patterns $patterns.Umowa
        $colOpis   = Find-ColumnHeader -Headers $headers -Patterns $patterns.Opis
        $colDom    = Find-ColumnHeader -Headers $headers -Patterns $patterns.Dom
        $colSzkola = Find-ColumnHeader -Headers $headers -Patterns $patterns.Szkola
        $colPraca  = Find-ColumnHeader -Headers $headers -Patterns $patterns.Praca

        $borderColStatus.Visibility = [System.Windows.Visibility]::Visible

        # Umowa (required)
        if ($colUmowa) {
            $badgeUmowa.Text = "Umowa: ✓ ($colUmowa)"
            $badgeUmowa.Foreground = [System.Windows.Media.Brushes]::LightGreen
        } else {
            $badgeUmowa.Text = "Umowa: ✗ Missing!"
            $badgeUmowa.Foreground = [System.Windows.Media.Brushes]::Tomato
        }

        # Opis (optional - falls back to Umowa)
        if ($colOpis) {
            $badgeOpis.Text = "Opis: ✓ ($colOpis)"
            $badgeOpis.Foreground = [System.Windows.Media.Brushes]::LightGreen
        } else {
            $badgeOpis.Text = "Opis: ~ Falls back to Umowa"
            $badgeOpis.Foreground = [System.Windows.Media.Brushes]::Khaki
        }

        # Dom (required)
        if ($colDom) {
            $badgeDom.Text = "Dom: ✓ ($colDom)"
            $badgeDom.Foreground = [System.Windows.Media.Brushes]::LightGreen
        } else {
            $badgeDom.Text = "Dom: ✗ Missing!"
            $badgeDom.Foreground = [System.Windows.Media.Brushes]::Tomato
        }

        # Szkola (required)
        if ($colSzkola) {
            $badgeSzkola.Text = "Szkola: ✓ ($colSzkola)"
            $badgeSzkola.Foreground = [System.Windows.Media.Brushes]::LightGreen
        } else {
            $badgeSzkola.Text = "Szkola: ✗ Missing!"
            $badgeSzkola.Foreground = [System.Windows.Media.Brushes]::Tomato
        }

        # Praca (optional)
        if ($colPraca) {
            $badgePraca.Text = "Praca: ✓ ($colPraca)"
            $badgePraca.Foreground = [System.Windows.Media.Brushes]::LightGreen
        } else {
            $badgePraca.Text = "Praca: (none)"
            $badgePraca.Foreground = [System.Windows.Media.Brushes]::Gray
        }

        Write-GuiLog "Loaded Excel headers: $($headers -join ', ')" "INFO"
    } catch {
        Write-GuiLog "Error reading Excel columns: $($_.Exception.Message)" "WARN"
    }
}

# ── Select Excel file ─────────────────────────────────────────────────────────
$btnBrowseExcel.Add_Click({
    $dlg = [System.Windows.Forms.OpenFileDialog]::new()
    $dlg.Title = 'Select input Excel file with contracts'
    $dlg.Filter = 'Excel Workbooks (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*'
    if (-not [string]::IsNullOrWhiteSpace($txtExcelPath.Text) -and (Test-Path $txtExcelPath.Text)) {
        $dlg.InitialDirectory = Split-Path -Path $txtExcelPath.Text -Parent
    } elseif (-not [string]::IsNullOrWhiteSpace($loadedConfig.LastExcelPath) -and (Test-Path (Split-Path -Path $loadedConfig.LastExcelPath -Parent))) {
        $dlg.InitialDirectory = Split-Path -Path $loadedConfig.LastExcelPath -Parent
    } else {
        $dlg.InitialDirectory = [Environment]::GetFolderPath('MyDocuments')
    }

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtExcelPath.Text = $dlg.FileName
        Update-ExcelColumnStatus -FilePath $dlg.FileName
        Save-CurrentGuiConfig
    }
})

$txtExcelPath.Add_LostFocus({
    if (-not [string]::IsNullOrWhiteSpace($txtExcelPath.Text)) {
        Update-ExcelColumnStatus -FilePath $txtExcelPath.Text.Trim()
    }
    Save-CurrentGuiConfig
})

$txtExcelPath.Add_TextChanged({
    Save-CurrentGuiConfig
})

# Drag & Drop Excel file onto the application window
$window.AllowDrop = $true
$window.Add_Drop({
    param($s, $e)
    if ($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
        $files = $e.Data.GetData([System.Windows.DataFormats]::FileDrop)
        if ($files -and $files.Count -gt 0) {
            $file = $files[0]
            if ($file -match '\.(xlsx|xls)$') {
                $txtExcelPath.Text = $file
                Update-ExcelColumnStatus -FilePath $file
                Save-CurrentGuiConfig
            } else {
                [System.Windows.MessageBox]::Show('The dropped file is not an Excel workbook (.xlsx or .xls).', 'Invalid File', 'OK', 'Warning')
            }
        }
    }
})

# ── Select output folder ──────────────────────────────────────────────────────
$btnBrowseOutput.Add_Click({
    $fbd = [System.Windows.Forms.FolderBrowserDialog]::new()
    $fbd.Description = 'Select destination folder for generated routes and maps'
    if (-not [string]::IsNullOrWhiteSpace($txtOutputFolder.Text) -and (Test-Path $txtOutputFolder.Text)) {
        $fbd.SelectedPath = $txtOutputFolder.Text
    }
    if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtOutputFolder.Text = $fbd.SelectedPath
        Save-CurrentGuiConfig
    }
})

$btnOpenOutput.Add_Click({
    $target = if (-not [string]::IsNullOrWhiteSpace($txtOutputFolder.Text)) {
        $txtOutputFolder.Text.Trim()
    } elseif ($script:CurrentOutputFolder) {
        $script:CurrentOutputFolder
    } else {
        $PSScriptRoot
    }
    if (-not (Test-Path $target)) {
        New-Item -ItemType Directory -Path $target -Force | Out-Null
    }
    Start-Process -FilePath 'explorer.exe' -ArgumentList "`"$target`""
})

# ── Stop processing ───────────────────────────────────────────────────────────
$btnStop.Add_Click({
    if ($script:IsWorking) {
        $btnStop.IsEnabled = $false
        $script:CancellationRequested = $true
        if ($script:SyncState) {
            $script:SyncState.CancellationRequested = $true
        }
        Write-GuiLog "Cancellation requested. Waiting for current operation to complete..." "WARN"
        $lblProgressText.Text = "Stopping..."
    }
})

# ══════════════════════════════════════════════════════════════════════════════
# MAIN PROCESSING ACTION (BACKGROUND THREAD - NON-BLOCKING WPF)
# ══════════════════════════════════════════════════════════════════════════════

$btnStart.Add_Click({
    $apiKey       = (Get-CurrentApiKey)
    $excelPath    = $txtExcelPath.Text.Trim()
    $outputFolder = $txtOutputFolder.Text.Trim()
    $rememberKey  = [bool]$chkRememberApiKey.IsChecked
    $script:CurrentOutputFolder = $outputFolder

    # Pre-flight validations
    if ([string]::IsNullOrWhiteSpace($apiKey)) {
        [System.Windows.MessageBox]::Show("Please provide a Google Maps API key before starting.", "Missing API Key", "OK", "Warning")
        return
    }

    if ([string]::IsNullOrWhiteSpace($excelPath) -or -not (Test-Path $excelPath)) {
        [System.Windows.MessageBox]::Show("Please select an existing Excel file with contracts.", "Missing Input File", "OK", "Warning")
        return
    }

    if ([string]::IsNullOrWhiteSpace($outputFolder)) {
        $outputFolder = Join-Path $PSScriptRoot 'Results'
        $txtOutputFolder.Text = $outputFolder
    }

    # Map dimensions
    $mapW = 900
    $mapH = 600
    if ($cmbMapDimensions.SelectedItem) {
        $tagParts = ($cmbMapDimensions.SelectedItem.Tag -as [string]) -split 'x'
        if ($tagParts.Count -eq 2) {
            $mapW = [int]$tagParts[0]
            $mapH = [int]$tagParts[1]
        }
    }

    # Save configuration
    Save-AppConfig -ApiKey $apiKey -RememberApiKey $rememberKey -OutputFolder $outputFolder -MapWidth $mapW -MapHeight $mapH
    if ($rememberKey) {
        $lblKeyStatus.Text = '🔒 Securely saved to DPAPI'
    } else {
        $lblKeyStatus.Text = 'API key will not be remembered'
    }

    # Lock UI during execution
    $script:IsWorking = $true
    $script:CancellationRequested = $false
    $btnStart.IsEnabled = $false
    $btnStop.IsEnabled = $true
    $btnBrowseExcel.IsEnabled = $false
    $btnBrowseOutput.IsEnabled = $false
    $txtLog.Clear()

    Write-GuiLog "Starting contract processing task..." "INFO"
    Write-GuiLog "Input file    : $excelPath" "INFO"
    Write-GuiLog "Output folder : $outputFolder" "INFO"
    Write-GuiLog "Map dimensions: ${mapW}x${mapH} px" "INFO"

    # Create output directory
    if (-not (Test-Path $outputFolder)) {
        New-Item -ItemType Directory -Path $outputFolder -Force | Out-Null
    }

    # Shared synchronized state for progress and cancellation
    $syncState = [hashtable]::Synchronized(@{
        CancellationRequested = $false
        Result                = $null
        Error                 = $null
    })
    $script:SyncState = $syncState

    # Set up InitialSessionState with helper functions from this script
    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    Get-ChildItem function: | ForEach-Object {
        if ($_.Name -notin @('prompt', 'TabExpansion2', 'Clear-Host', 'help', 'mkdir', 'more')) {
            try {
                $iss.Commands.Add([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new($_.Name, $_.Definition))
            } catch { }
        }
    }

    $workerRS = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($iss)
    $workerRS.ApartmentState = [System.Threading.ApartmentState]::STA
    $workerRS.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
    $workerRS.Open()

    # Pass UI references and parameters into the Runspace's SessionState
    $workerRS.SessionStateProxy.SetVariable('window', $window)
    $workerRS.SessionStateProxy.SetVariable('txtLog', $txtLog)
    $workerRS.SessionStateProxy.SetVariable('progBar', $progBar)
    $workerRS.SessionStateProxy.SetVariable('lblProgressPercent', $lblProgressPercent)
    $workerRS.SessionStateProxy.SetVariable('lblProgressText', $lblProgressText)
    $workerRS.SessionStateProxy.SetVariable('lblStatsProcessed', $lblStatsProcessed)
    $workerRS.SessionStateProxy.SetVariable('lblStatsSuccess', $lblStatsSuccess)
    $workerRS.SessionStateProxy.SetVariable('lblStatsError', $lblStatsError)
        $workerRS.SessionStateProxy.SetVariable('lblStatusModule', $lblStatusModule)
    $workerRS.SessionStateProxy.SetVariable('lblKeyStatus', $lblKeyStatus)
    $workerRS.SessionStateProxy.SetVariable('excelPath', $excelPath)
    $workerRS.SessionStateProxy.SetVariable('outputFolder', $outputFolder)
    $workerRS.SessionStateProxy.SetVariable('mapW', $mapW)
    $workerRS.SessionStateProxy.SetVariable('mapH', $mapH)
    $workerRS.SessionStateProxy.SetVariable('apiKey', $apiKey)
    $workerRS.SessionStateProxy.SetVariable('syncState', $syncState)

    $workerPS = [System.Management.Automation.PowerShell]::Create()
    $workerPS.Runspace = $workerRS

    $workerScript = {
        try {

        # 1. Check ImportExcel module
        if (-not (Test-ImportExcelAvailable)) {
            Write-GuiLog "ImportExcel module not found. Starting installation..." "WARN"
            $window.Dispatcher.Invoke([Action]{
                $lblStatusModule.Text = "Installing ImportExcel module from PSGallery..."
            })
            Install-ImportExcelModule -OnLog { param($msg) Write-GuiLog $msg "INFO" }
            $window.Dispatcher.Invoke([Action]{
                $lblStatusModule.Text = "ImportExcel Module: Ready"
                $lblStatusModule.Foreground = [System.Windows.Media.Brushes]::LightGreen
            })
        }

        # 1b. Pre-flight verification of Google Maps API Key
        Write-GuiLog "Validating Google Maps API key..." "INFO"
        $testKeyResult = Get-AddressCoordinates -Address "Polska" -ApiKey $apiKey
        if (-not $testKeyResult -or $testKeyResult.Status -eq 'REQUEST_DENIED' -or $testKeyResult.Status -match 'INVALID') {
            $apiErrMsg = if ($testKeyResult -and $testKeyResult.ErrorMessage) { $testKeyResult.ErrorMessage.Trim() } else { "Access denied (Status: $($testKeyResult.Status))" }
            $window.Dispatcher.Invoke([Action]{
                $lblKeyStatus.Text = '❌ Invalid API Key'
                $lblKeyStatus.Foreground = [System.Windows.Media.Brushes]::Salmon
                $lblStatsError.Text = "Errors: 1"
            })
            Write-GuiLog "Google Maps API validation failed: $apiErrMsg" "ERROR"
            throw "Google Maps API Key Error: $apiErrMsg`nPlease check your API key in the top box."
        }
        Write-GuiLog "Google Maps API key validated successfully." "OK"

        # 2. Read Excel data
        Write-GuiLog "Reading rows from Excel file..." "INFO"
        $Dane = Import-Excel -Path $excelPath
        if ($null -eq $Dane -or @($Dane).Count -eq 0) {
            throw "The Excel file is empty or contains no data rows."
        }

        $Headers = $Dane[0].PSObject.Properties.Name

        # 3. Map column headers
        $ColMap = @{
            Umowa   = Find-ColumnHeader -Headers $Headers -Patterns @('(?i)^\s*(umow[ay]|nr\s*umow[ay]|numer\s*umow[ay]|id|numer|nr|contract)\s*$', '(?i)umow|contract')
            Opis    = Find-ColumnHeader -Headers $Headers -Patterns @('(?i)^\s*(opis|nazwa|nazwa\s*folderu|folder|identyfikator|description)\s*$', '(?i)opis|desc')
            Szkola  = Find-ColumnHeader -Headers $Headers -Patterns @('(?i)^\s*(szko[łl][ay]|adres\s*szko[łl][yi]|plac[oó]wk[ay]|przedszkol[ea]|o[sś]rodek|szko[łl]a\s*adres|school)\s*$', '(?i)szko[łl]|plac[oó]wk|przedszkol|school')
            Dom     = Find-ColumnHeader -Headers $Headers -Patterns @('(?i)^\s*(dom|domu|adres\s*domu|adres\s*zamieszkania|zamieszkani[ea]|miejsce\s*zamieszkania|dom\s*adres|home)\s*$', '(?i)dom|zamieszkan|home')
            Praca   = Find-ColumnHeader -Headers $Headers -Patterns @('(?i)^\s*(prac[ay]|adres\s*prac[ay]|miejsce\s*prac[ay]|zak[łl]ad\s*prac[ay]|firma|praca\s*adres|work)\s*$', '(?i)prac|firm|work')
            Tryb    = Find-ColumnHeader -Headers $Headers -Patterns @('(?i)^\s*(tryb|tryb\s*przejazdu|spos[oó]b|mode)\s*$', '(?i)tryb|mode')
            Wariant = Find-ColumnHeader -Headers $Headers -Patterns @('(?i)^\s*(wariant|wariant\s*trasy|opcja|variant)\s*$', '(?i)wariant|opcj|variant')
        }

        # Validate required columns
        $Required = @('Umowa', 'Szkola', 'Dom')
        foreach ($c in $Required) {
            if (-not $ColMap[$c]) {
                throw "Could not find required column '$c' in Excel. Available columns: $($Headers -join ', ')"
            }
        }

        Write-GuiLog "Mapped columns: Umowa='$($ColMap.Umowa)', Opis='$($ColMap.Opis)', Dom='$($ColMap.Dom)', Szkola='$($ColMap.Szkola)', Praca='$($ColMap.Praca)'" "OK"

        $WszystkieWyniki = [System.Collections.Generic.List[PSCustomObject]]::new()
        $RowIndex = 0
        $TotalRows = @($Dane).Count
        $GeoCache = @{}
        $SuccessCount = 0
        $ErrorCount = 0

        foreach ($Row in $Dane) {
            if ($syncState.CancellationRequested) {
                Write-GuiLog "Processing cancelled by user after row $RowIndex." "WARN"
                break
            }

            $RowIndex++

            $NumerUmowy  = if ($ColMap.Umowa -and $null -ne $Row.($ColMap.Umowa)) { ($Row.($ColMap.Umowa) -as [string]).Trim() } else { $null }
            $Opis        = if ($ColMap.Opis -and $null -ne $Row.($ColMap.Opis)) { ($Row.($ColMap.Opis) -as [string]).Trim() } else { $null }
            $AdresSzkoly = if ($ColMap.Szkola -and $null -ne $Row.($ColMap.Szkola)) { ($Row.($ColMap.Szkola) -as [string]).Trim() } else { $null }
            $AdresDomu   = if ($ColMap.Dom -and $null -ne $Row.($ColMap.Dom)) { ($Row.($ColMap.Dom) -as [string]).Trim() } else { $null }
            $AdresPracy  = if ($ColMap.Praca -and $null -ne $Row.($ColMap.Praca)) { ($Row.($ColMap.Praca) -as [string]).Trim() } else { $null }
            $Tryb        = if ($ColMap.Tryb -and $null -ne $Row.($ColMap.Tryb)) { ($Row.($ColMap.Tryb) -as [string]).Trim() } else { $null }
            $Wariant     = if ($ColMap.Wariant -and $null -ne $Row.($ColMap.Wariant)) { ($Row.($ColMap.Wariant) -as [string]).Trim() } else { $null }

            # Update progress in UI
            $pct = [math]::Round(($RowIndex / $TotalRows) * 100)
            $window.Dispatcher.Invoke([Action]{
                $progBar.Value = $pct
                $lblProgressPercent.Text = "$pct%"
                $lblProgressText.Text = "Processing [$RowIndex/$TotalRows]: Contract $NumerUmowy"
                $lblStatsProcessed.Text = "Processed: $RowIndex / $TotalRows"
            })

            if ([string]::IsNullOrWhiteSpace($NumerUmowy)) {
                Write-GuiLog "Row ${RowIndex}: Empty contract number — skipping." "WARN"
                $ErrorCount++
                continue
            }

            $SafeIdentifier = if (-not [string]::IsNullOrWhiteSpace($Opis)) { $Opis } else { $NumerUmowy }
            $SafeName = ConvertTo-SafeFileName -Name $SafeIdentifier
            $DisplayLabel = if ($Opis) { "$NumerUmowy ($Opis)" } else { $NumerUmowy }

            Write-GuiLog "[$RowIndex/$TotalRows] Contract: $DisplayLabel (Folder: $SafeName)" "INFO"

            # Create contract subfolder
            $UmowaFolder = Join-Path -Path $outputFolder -ChildPath $SafeName
            if (-not (Test-Path -Path $UmowaFolder)) {
                New-Item -ItemType Directory -Path $UmowaFolder -Force | Out-Null
            }

            # Address validation
            if ([string]::IsNullOrWhiteSpace($AdresDomu) -or [string]::IsNullOrWhiteSpace($AdresSzkoly)) {
                Write-GuiLog "  Row ${RowIndex}: Empty Home or School address — error." "WARN"
                $ErrorCount++
                $WszystkieWyniki.Add([PSCustomObject]@{
                    'Numer umowy' = $NumerUmowy; 'Opis' = $Opis; 'Tryb' = $Tryb; 'Wariant' = $Wariant
                    'AdresDomu' = $AdresDomu; 'AdresSzkoly' = $AdresSzkoly; 'AdresPracy' = $AdresPracy
                    'Dom→Szkoła [km]' = $null; 'Szkoła→Dom [km]' = $null; 'Dom→Praca [km]' = $null; 'Praca→Dom [km]' = $null
                    'Szkoła→Praca [km]' = $null; 'Praca→Szkoła [km]' = $null; 'Status' = 'EMPTY_ADDRESSES'
                })
                continue
            }

            # Geocode Home
            $GeoDom = if ($GeoCache.ContainsKey($AdresDomu)) { $GeoCache[$AdresDomu] } else {
                $g = Get-AddressCoordinates -Address $AdresDomu -ApiKey $apiKey
                $GeoCache[$AdresDomu] = $g
                Start-Sleep -Milliseconds 150
                $g
            }

            # Geocode School
            $GeoSzkola = if ($GeoCache.ContainsKey($AdresSzkoly)) { $GeoCache[$AdresSzkoly] } else {
                $g = Get-AddressCoordinates -Address $AdresSzkoly -ApiKey $apiKey
                $GeoCache[$AdresSzkoly] = $g
                Start-Sleep -Milliseconds 150
                $g
            }

            $HasError = $false
            if (-not $GeoDom -or $GeoDom.Status -ne 'OK') {
                $detail = if ($GeoDom -and $GeoDom.ErrorMessage) { " - $($GeoDom.ErrorMessage)" } else { "" }
                Write-GuiLog "  Home geocoding error ($AdresDomu): $($GeoDom.Status)$detail" "ERROR"
                $HasError = $true
            }
            if (-not $GeoSzkola -or $GeoSzkola.Status -ne 'OK') {
                $detail = if ($GeoSzkola -and $GeoSzkola.ErrorMessage) { " - $($GeoSzkola.ErrorMessage)" } else { "" }
                Write-GuiLog "  School geocoding error ($AdresSzkoly): $($GeoSzkola.Status)$detail" "ERROR"
                $HasError = $true
            }

            # Work (optional)
            $MaPrace = Test-PracaAddress -Praca $AdresPracy
            $GeoPraca = $null
            if ($MaPrace) {
                $GeoPraca = if ($GeoCache.ContainsKey($AdresPracy)) { $GeoCache[$AdresPracy] } else {
                    $g = Get-AddressCoordinates -Address $AdresPracy -ApiKey $apiKey
                    $GeoCache[$AdresPracy] = $g
                    Start-Sleep -Milliseconds 150
                    $g
                }
                if (-not $GeoPraca -or $GeoPraca.Status -ne 'OK') {
                    $wDetail = if ($GeoPraca -and $GeoPraca.ErrorMessage) { " - $($GeoPraca.ErrorMessage)" } else { "" }
                    Write-GuiLog "  Work geocoding error ($AdresPracy): $($GeoPraca.Status)$wDetail (work routes skipped)" "WARN"
                    $MaPrace = $false
                }
            }

            if ($HasError) {
                $ErrorCount++
                $window.Dispatcher.Invoke([Action]{
                    $lblStatsSuccess.Text = "OK: $SuccessCount"
                    $lblStatsError.Text   = "Errors: $ErrorCount"
                })
                $WszystkieWyniki.Add([PSCustomObject]@{
                    'Numer umowy' = $NumerUmowy; 'Opis' = $Opis; 'Tryb' = $Tryb; 'Wariant' = $Wariant
                    'AdresDomu' = $AdresDomu; 'AdresSzkoly' = $AdresSzkoly; 'AdresPracy' = $AdresPracy
                    'Status' = 'GEOCODING_ERROR'
                })
                continue
            }

            $curDate = (Get-Date).ToString('yyyy-MM-dd')

            # ── Route 1: Dom -> Szkola ──
            $PngDomSzkola = Join-Path $UmowaFolder "${SafeName}_Dom_Szkoła.png"
            $Res1 = Invoke-RouteAndMap -GeoStart $GeoDom -GeoEnd $GeoSzkola -PngPath $PngDomSzkola `
                -LabelStart 'Dom' -LabelEnd 'Szkoła' -ApiKey $apiKey -Width $mapW -Height $mapH `
                -NumerUmowy $NumerUmowy -Opis $Opis -DataWygenerowania $curDate
            $KmDomSzkola = $Res1.OdlegloscKm
            Write-GuiLog "  Dom -> Szkoła: $KmDomSzkola km" "OK"
            Start-Sleep -Milliseconds 200

            # ── Route 2: Szkola -> Dom ──
            $PngSzkolaDom = Join-Path $UmowaFolder "${SafeName}_Szkoła_Dom.png"
            $Res2 = Invoke-RouteAndMap -GeoStart $GeoSzkola -GeoEnd $GeoDom -PngPath $PngSzkolaDom `
                -LabelStart 'Szkoła' -LabelEnd 'Dom' -ApiKey $apiKey -Width $mapW -Height $mapH `
                -NumerUmowy $NumerUmowy -Opis $Opis -DataWygenerowania $curDate
            $KmSzkolaDom = $Res2.OdlegloscKm
            Write-GuiLog "  Szkoła -> Dom: $KmSzkolaDom km" "OK"
            Start-Sleep -Milliseconds 200

            # Combined routes with Work
            $KmDomPraca = $null; $KmPracaDom = $null; $KmSzkolaPraca = $null; $KmPracaSzkola = $null
            if ($MaPrace) {
                # Dom -> Praca
                $Png3 = Join-Path $UmowaFolder "${SafeName}_Dom_Praca.png"
                $Res3 = Invoke-RouteAndMap -GeoStart $GeoDom -GeoEnd $GeoPraca -PngPath $Png3 `
                    -LabelStart 'Dom' -LabelEnd 'Praca' -ApiKey $apiKey -Width $mapW -Height $mapH `
                    -NumerUmowy $NumerUmowy -Opis $Opis -DataWygenerowania $curDate
                $KmDomPraca = $Res3.OdlegloscKm

                # Praca -> Dom
                $Png4 = Join-Path $UmowaFolder "${SafeName}_Praca_Dom.png"
                $Res4 = Invoke-RouteAndMap -GeoStart $GeoPraca -GeoEnd $GeoDom -PngPath $Png4 `
                    -LabelStart 'Praca' -LabelEnd 'Dom' -ApiKey $apiKey -Width $mapW -Height $mapH `
                    -NumerUmowy $NumerUmowy -Opis $Opis -DataWygenerowania $curDate
                $KmPracaDom = $Res4.OdlegloscKm

                # Szkola -> Praca
                $Png5 = Join-Path $UmowaFolder "${SafeName}_Szkoła_Praca.png"
                $Res5 = Invoke-RouteAndMap -GeoStart $GeoSzkola -GeoEnd $GeoPraca -PngPath $Png5 `
                    -LabelStart 'Szkoła' -LabelEnd 'Praca' -ApiKey $apiKey -Width $mapW -Height $mapH `
                    -NumerUmowy $NumerUmowy -Opis $Opis -DataWygenerowania $curDate
                $KmSzkolaPraca = $Res5.OdlegloscKm

                # Praca -> Szkola
                $Png6 = Join-Path $UmowaFolder "${SafeName}_Praca_Szkoła.png"
                $Res6 = Invoke-RouteAndMap -GeoStart $GeoPraca -GeoEnd $GeoSzkola -PngPath $Png6 `
                    -LabelStart 'Praca' -LabelEnd 'Szkoła' -ApiKey $apiKey -Width $mapW -Height $mapH `
                    -NumerUmowy $NumerUmowy -Opis $Opis -DataWygenerowania $curDate
                $KmPracaSzkola = $Res6.OdlegloscKm

                Write-GuiLog "  Work routes: Dom-Praca ($KmDomPraca km), Szk-Praca ($KmSzkolaPraca km)" "OK"
            }

            # Row summary
            $WierszWynik = [PSCustomObject]@{
                'Numer umowy'       = $NumerUmowy
                'Opis'              = $Opis
                'Tryb'              = $Tryb
                'Wariant'           = $Wariant
                'AdresDomu'         = $AdresDomu
                'Dom_Geokodowany'   = if ($GeoDom) { $GeoDom.FormattedAddress } else { $null }
                'Dom_UlicaINumer'   = if ($GeoDom) { $GeoDom.UlicaINumer } else { $null }
                'Dom_KodPocztowy'   = if ($GeoDom) { $GeoDom.KodPocztowy } else { $null }
                'Dom_Miasto'        = if ($GeoDom) { $GeoDom.Miasto } else { $null }
                'AdresSzkoly'       = $AdresSzkoly
                'Szkoła_Geokodowany' = if ($GeoSzkola) { $GeoSzkola.FormattedAddress } else { $null }
                'Szkoła_UlicaINumer' = if ($GeoSzkola) { $GeoSzkola.UlicaINumer } else { $null }
                'Szkoła_KodPocztowy' = if ($GeoSzkola) { $GeoSzkola.KodPocztowy } else { $null }
                'Szkoła_Miasto'     = if ($GeoSzkola) { $GeoSzkola.Miasto } else { $null }
                'AdresPracy'        = $AdresPracy
                'Praca_Geokodowany' = if ($GeoPraca) { $GeoPraca.FormattedAddress } else { $null }
                'Praca_UlicaINumer' = if ($GeoPraca) { $GeoPraca.UlicaINumer } else { $null }
                'Praca_KodPocztowy' = if ($GeoPraca) { $GeoPraca.KodPocztowy } else { $null }
                'Praca_Miasto'      = if ($GeoPraca) { $GeoPraca.Miasto } else { $null }
                'Dom→Szkoła [km]'   = $KmDomSzkola
                'Szkoła→Dom [km]'   = $KmSzkolaDom
                'Dom→Praca [km]'    = $KmDomPraca
                'Praca→Dom [km]'    = $KmPracaDom
                'Szkoła→Praca [km]' = $KmSzkolaPraca
                'Praca→Szkoła [km]' = $KmPracaSzkola
                'Status'            = 'OK'
            }

            $WszystkieWyniki.Add($WierszWynik)
            $SuccessCount++

            # Individual contract Excel summary
            $UmowaExcelPath = Join-Path -Path $UmowaFolder -ChildPath "${SafeName}.xlsx"
            $WierszWynik | Export-Excel -Path $UmowaExcelPath -WorksheetName 'Podsumowanie' `
                -TableName 'Podsumowanie' -AutoSize -AutoFilter -ClearSheet `
                -NoNumberConversion 'Numer umowy' -FreezeTopRow

            $window.Dispatcher.Invoke([Action]{
                $lblStatsSuccess.Text = "OK: $SuccessCount"
                $lblStatsError.Text   = "Errors: $ErrorCount"
            })
        }

        # 4. Master Excel summary
        if ($WszystkieWyniki.Count -gt 0) {
            $DateStamp = Get-Date -Format 'yyyyMMdd_HHmm'
            $ZbiorczyExcelPath = Join-Path -Path $outputFolder -ChildPath "${DateStamp}_Podsumowanie_wszystkie_umowy.xlsx"
            Write-GuiLog "Generating combined Excel report: $ZbiorczyExcelPath" "INFO"

            $WszystkieWyniki | Export-Excel -Path $ZbiorczyExcelPath -WorksheetName 'Wszystkie umowy' `
                -TableName 'WszystkieUmowy' -AutoSize -AutoFilter -ClearSheet `
                -NoNumberConversion 'Numer umowy' -FreezeTopRow

            $Dane | Export-Excel -Path $ZbiorczyExcelPath -WorksheetName 'Dane wejściowe' `
                -TableName 'DaneWejsciowe' -AutoSize -AutoFilter -ClearSheet `
                -NoNumberConversion *

            Write-GuiLog "Saved combined Excel report: $ZbiorczyExcelPath" "OK"
        }

            $syncState.Result = @{
                Total   = $TotalRows
                Success = $SuccessCount
                Errors  = $ErrorCount
            }
        }
        catch {
            $syncState.Error = $_
            Write-GuiLog "Execution halted due to error: $($_.Exception.Message)" "ERROR"
        }
    }

    $workerPS.AddScript($workerScript) | Out-Null
    $script:WorkerPS = $workerPS
    $script:WorkerRS = $workerRS
    $script:AsyncResult = $workerPS.BeginInvoke()

    # Poll for completion on UI thread using DispatcherTimer
    $script:PollTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $script:PollTimer.Interval = [TimeSpan]::FromMilliseconds(150)
    $script:PollTimer.Add_Tick({
        param($sender, $e)
        try {
            if ($script:AsyncResult -and $script:AsyncResult.IsCompleted) {
                if ($sender) { $sender.Stop() }
                if ($script:PollTimer) { $script:PollTimer.Stop() }

                $pipelineError = $null
                if ($script:WorkerPS) {
                    try {
                        $script:WorkerPS.EndInvoke($script:AsyncResult) | Out-Null
                    } catch {
                        $pipelineError = $_
                    }

                    if ($script:WorkerPS.Streams.Error.Count -gt 0 -and -not $pipelineError) {
                        $pipelineError = $script:WorkerPS.Streams.Error[0]
                    }
                }

                $execError = if ($script:SyncState -and $script:SyncState.Error) { $script:SyncState.Error } else { $pipelineError }

                # Cleanup Runspace and PowerShell instance
                try { if ($script:WorkerPS) { $script:WorkerPS.Dispose() } } catch { }
                try { if ($script:WorkerRS) { $script:WorkerRS.Close(); $script:WorkerRS.Dispose() } } catch { }
                $script:WorkerPS = $null
                $script:WorkerRS = $null
                $script:AsyncResult = $null

                $script:IsWorking = $false
                $btnStart.IsEnabled = $true
                $btnStop.IsEnabled = $false
                $btnBrowseExcel.IsEnabled = $true
                $btnBrowseOutput.IsEnabled = $true

                if ($execError) {
                    $errMsg = if ($execError.Exception) { $execError.Exception.Message } else { "$execError" }
                    Write-GuiLog "Critical error: $errMsg" "ERROR"
                    $lblProgressText.Text = "Error: $errMsg"
                    $lblStatsError.Text = "Errors: 1"
                    [System.Windows.MessageBox]::Show("Processing Error:`n`n$errMsg", "Google Maps Error", "OK", "Error")
                } elseif ($script:SyncState -and $script:SyncState.CancellationRequested) {
                    Write-GuiLog "Task was cancelled." "WARN"
                    $lblProgressText.Text = "Cancelled by user."
                } else {
                    $res = if ($script:SyncState) { $script:SyncState.Result } else { $null }
                    $tot = if ($res -and $null -ne $res.Total) { $res.Total } else { 0 }
                    $suc = if ($res -and $null -ne $res.Success) { $res.Success } else { 0 }
                    $errCount = if ($res -and $null -ne $res.Errors) { $res.Errors } else { 0 }

                    $lblStatsProcessed.Text = "Processed: $tot / $tot"
                    $lblStatsSuccess.Text   = "OK: $suc"
                    $lblStatsError.Text     = "Errors: $errCount"

                    if ($suc -gt 0 -and $errCount -eq 0) {
                        Write-GuiLog "Processing finished successfully! Processed: $tot, OK: $suc, Errors: 0" "OK"
                        $lblProgressText.Text = "Finished successfully! (OK: $suc)"
                        $progBar.Value = 100
                        $lblProgressPercent.Text = "100%"

                        $dialogResult = [System.Windows.MessageBox]::Show(
                            "Processing finished successfully!`n`nProcessed contracts: $tot`nSuccess (OK): $suc`nErrors: 0`n`nWould you like to open the output folder?",
                            "Success", "YesNo", "Information")
                        if ($dialogResult -eq [System.Windows.MessageBoxResult]::Yes -or $dialogResult -eq 'Yes' -or "$dialogResult" -eq 'Yes') {
                            $targetFolder = if (-not [string]::IsNullOrWhiteSpace($txtOutputFolder.Text)) {
                                $txtOutputFolder.Text.Trim()
                            } elseif ($script:CurrentOutputFolder) {
                                $script:CurrentOutputFolder
                            } else {
                                'C:\Temp\SchoolRoutes'
                            }
                            if (-not (Test-Path $targetFolder)) {
                                New-Item -ItemType Directory -Path $targetFolder -Force | Out-Null
                            }
                            Start-Process -FilePath 'explorer.exe' -ArgumentList "`"$targetFolder`""
                        }
                    } elseif ($suc -gt 0 -and $errCount -gt 0) {
                        Write-GuiLog "Processing completed with some errors. Processed: $tot, OK: $suc, Errors: $errCount" "WARN"
                        $lblProgressText.Text = "Completed with warnings! (OK: $suc, Errors: $errCount)"
                        $progBar.Value = 100
                        $lblProgressPercent.Text = "100%"

                        $dialogResult = [System.Windows.MessageBox]::Show(
                            "Processing completed with warnings!`n`nProcessed contracts: $tot`nSuccess (OK): $suc`nErrors: $errCount`n`nSome contracts encountered errors. Would you like to open the output folder?",
                            "Completed with Warnings", "YesNo", "Warning")
                        if ($dialogResult -eq [System.Windows.MessageBoxResult]::Yes -or $dialogResult -eq 'Yes' -or "$dialogResult" -eq 'Yes') {
                            $targetFolder = if (-not [string]::IsNullOrWhiteSpace($txtOutputFolder.Text)) {
                                $txtOutputFolder.Text.Trim()
                            } elseif ($script:CurrentOutputFolder) {
                                $script:CurrentOutputFolder
                            } else {
                                'C:\Temp\SchoolRoutes'
                            }
                            if (-not (Test-Path $targetFolder)) {
                                New-Item -ItemType Directory -Path $targetFolder -Force | Out-Null
                            }
                            Start-Process -FilePath 'explorer.exe' -ArgumentList "`"$targetFolder`""
                        }
                    } else {
                        Write-GuiLog "Processing failed! Processed: $tot, OK: 0, Errors: $errCount" "ERROR"
                        $lblProgressText.Text = "Failed: 0 contracts processed ($errCount errors)."
                        $progBar.Value = 100
                        $lblProgressPercent.Text = "100%"

                        [System.Windows.MessageBox]::Show(
                            "Processing FAILED!`n`nProcessed contracts: $tot`nSuccess (OK): 0`nErrors: $errCount`n`nAll contracts failed to process. Please check the event log and verify your Google Maps API key and contract addresses.",
                            "Processing Failed", "OK", "Error")
                    }
                }
            }
        }
        catch {
            try { if ($sender) { $sender.Stop() } } catch {}
            try { if ($script:PollTimer) { $script:PollTimer.Stop() } } catch {}
            $script:AsyncResult = $null
            $script:WorkerPS = $null
            $script:WorkerRS = $null
            $script:IsWorking = $false
            Write-GuiLog "Error finalizing task: $($_.Exception.Message)" "ERROR"
            $btnStart.IsEnabled = $true
            $btnStop.IsEnabled = $false
            $btnBrowseExcel.IsEnabled = $true
            $btnBrowseOutput.IsEnabled = $true
        }
    })
    $script:PollTimer.Start()
})

# ── Display application window ────────────────────────────────────────────────
Write-GuiLog "Application started successfully. Select an Excel file and start route calculation." "INFO"
$window.ShowDialog() | Out-Null
