#Requires -Version 5.1
<#
.SYNOPSIS
    Google Maps Routes & Map Generator — Configuration, DPAPI Security & Logging Subsystem.
.DESCRIPTION
    Manages DPAPI-encrypted API keys, application configuration persistence,
    multi-language localization catalogs, overlay banner templates, and Google Maps API usage tracking.
.NOTES
    Encoding: UTF-8 with BOM
#>

# ══════════════════════════════════════════════════════════════════════════════
# 1. DPAPI SECURITY & LOGGING
# ══════════════════════════════════════════════════════════════════════════════

$script:AppDataDir      = Join-Path $env:LOCALAPPDATA 'GoogleMapsRoutes'
$script:ConfigFile      = Join-Path $script:AppDataDir 'config.json'
$script:LogFile         = Join-Path $script:AppDataDir 'GoogleMapsRoutes.log'
$script:AppConfig       = $null
$script:LocCatalog      = $null
$script:CurrentLanguage = 'en'
$script:CurrentTheme    = 'Dark'

if (-not (Test-Path $script:AppDataDir)) {
    New-Item -ItemType Directory -Path $script:AppDataDir -Force | Out-Null
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
        try {
            $sec = ConvertTo-SecureString -String $PlainText -AsPlainText -Force
            return (ConvertFrom-SecureString -SecureString $sec)
        }
        catch {
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
            try {
                return [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
            }
            finally {
                [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
            }
        }
        catch {
            return $null
        }
    }
}

function Get-MaskedKey([string]$Key) {
    if ([string]::IsNullOrWhiteSpace($Key)) { return '(none)' }
    if ($Key.Length -le 8) { return '***' }
    return "$($Key.Substring(0, 4))...$($Key.Substring($Key.Length - 4))"
}

function Get-CurrentApiKey {
    [CmdletBinding()]
    param()

    # 1. Check UI inputs if window/controls are active
    if ($script:Controls) {
        if ($script:Controls.txtSettingsApiKey -and $script:Controls.txtSettingsApiKey.Visibility -eq [System.Windows.Visibility]::Visible) {
            $k = [string]$script:Controls.txtSettingsApiKey.Password
            if (-not [string]::IsNullOrWhiteSpace($k)) { return $k.Trim() }
        }
        if ($script:Controls.txtSettingsApiKeyVisible -and $script:Controls.txtSettingsApiKeyVisible.Visibility -eq [System.Windows.Visibility]::Visible) {
            $k = [string]$script:Controls.txtSettingsApiKeyVisible.Text
            if (-not [string]::IsNullOrWhiteSpace($k)) { return $k.Trim() }
        }
        if ($script:Controls.txtSettingsApiKey -and -not [string]::IsNullOrWhiteSpace($script:Controls.txtSettingsApiKey.Password)) {
            return $script:Controls.txtSettingsApiKey.Password.Trim()
        }
        if ($script:Controls.txtSettingsApiKeyVisible -and -not [string]::IsNullOrWhiteSpace($script:Controls.txtSettingsApiKeyVisible.Text)) {
            return $script:Controls.txtSettingsApiKeyVisible.Text.Trim()
        }
    }

    # 2. Check loaded AppConfig
    if ($script:AppConfig -and -not [string]::IsNullOrWhiteSpace($script:AppConfig.ApiKey)) {
        return $script:AppConfig.ApiKey.Trim()
    }

    # 3. Check environment variable
    if (-not [string]::IsNullOrWhiteSpace($env:GOOGLE_MAPS_API_KEY)) {
        return $env:GOOGLE_MAPS_API_KEY.Trim()
    }

    return ''
}
Set-Item -Path "function:global:Get-CurrentApiKey" -Value (Get-Item "function:Get-CurrentApiKey").ScriptBlock -ErrorAction SilentlyContinue

function Set-CurrentApiKey([string]$Key) {
    if ($script:Controls) {
        if ($script:Controls.txtSettingsApiKey) { $script:Controls.txtSettingsApiKey.Password = $Key }
        if ($script:Controls.txtSettingsApiKeyVisible) { $script:Controls.txtSettingsApiKeyVisible.Text = $Key }
    }
}
Set-Item -Path "function:global:Set-CurrentApiKey" -Value (Get-Item "function:Set-CurrentApiKey").ScriptBlock -ErrorAction SilentlyContinue

function Get-CurrentCartoApiKey {
    [CmdletBinding()]
    param()

    if ($script:Controls) {
        if ($script:Controls.txtSettingsCartoApiKey -and $script:Controls.txtSettingsCartoApiKey.Visibility -eq [System.Windows.Visibility]::Visible) {
            $k = [string]$script:Controls.txtSettingsCartoApiKey.Password
            if (-not [string]::IsNullOrWhiteSpace($k)) { return $k.Trim() }
        }
        if ($script:Controls.txtSettingsCartoApiKeyVisible -and $script:Controls.txtSettingsCartoApiKeyVisible.Visibility -eq [System.Windows.Visibility]::Visible) {
            $k = [string]$script:Controls.txtSettingsCartoApiKeyVisible.Text
            if (-not [string]::IsNullOrWhiteSpace($k)) { return $k.Trim() }
        }
        if ($script:Controls.txtSettingsCartoApiKey -and -not [string]::IsNullOrWhiteSpace($script:Controls.txtSettingsCartoApiKey.Password)) {
            return $script:Controls.txtSettingsCartoApiKey.Password.Trim()
        }
        if ($script:Controls.txtSettingsCartoApiKeyVisible -and -not [string]::IsNullOrWhiteSpace($script:Controls.txtSettingsCartoApiKeyVisible.Text)) {
            return $script:Controls.txtSettingsCartoApiKeyVisible.Text.Trim()
        }
    }

    if ($script:AppConfig -and -not [string]::IsNullOrWhiteSpace($script:AppConfig.CartoApiKey)) {
        return $script:AppConfig.CartoApiKey.Trim()
    }

    if (-not [string]::IsNullOrWhiteSpace($env:CARTO_API_KEY)) {
        return $env:CARTO_API_KEY.Trim()
    }

    return ''
}
Set-Item -Path "function:global:Get-CurrentCartoApiKey" -Value (Get-Item "function:Get-CurrentCartoApiKey").ScriptBlock -ErrorAction SilentlyContinue

function Set-CurrentCartoApiKey([string]$Key) {
    if ($script:Controls) {
        if ($script:Controls.txtSettingsCartoApiKey) { $script:Controls.txtSettingsCartoApiKey.Password = $Key }
        if ($script:Controls.txtSettingsCartoApiKeyVisible) { $script:Controls.txtSettingsCartoApiKeyVisible.Text = $Key }
    }
}
Set-Item -Path "function:global:Set-CurrentCartoApiKey" -Value (Get-Item "function:Set-CurrentCartoApiKey").ScriptBlock -ErrorAction SilentlyContinue

function Write-AppLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [Parameter()][ValidateSet('INFO', 'OK', 'WARN', 'ERROR', 'DEBUG')][string]$Level = 'INFO'
    )
    $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff')
    $safeMsg = $Message
    $cfg = Get-Variable -Scope Script -Name 'AppConfig' -ValueOnly -ErrorAction SilentlyContinue
    if ($cfg -and $cfg.ApiKey -and $cfg.ApiKey.Length -gt 8) {
        $masked = Get-MaskedKey $cfg.ApiKey
        $safeMsg = $safeMsg.Replace($cfg.ApiKey, $masked)
    }
    $line = "[$ts] [$Level] $safeMsg"
    try {
        if ($script:LogFile) {
            [System.IO.File]::AppendAllText($script:LogFile, "$line`r`n", [System.Text.Encoding]::UTF8)
        }
    }
    catch { }
}
Set-Item -Path "function:global:Write-AppLog" -Value (Get-Item "function:Write-AppLog").ScriptBlock -ErrorAction SilentlyContinue

# ══════════════════════════════════════════════════════════════════════════════
# 2. LOCALIZATION CATALOG SUBSYSTEM
# ══════════════════════════════════════════════════════════════════════════════

$script:LocCatalog = $null
$script:CurrentLanguage = 'en'

function Load-LocalizationConfig {
    [CmdletBinding()]
    param()

    $locPath = $null
    $baseDir = if (-not [string]::IsNullOrWhiteSpace($script:AppDir)) {
        $script:AppDir
    } elseif (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        $PSScriptRoot
    } else {
        [System.IO.Path]::GetDirectoryName([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
    }

    $candidates = [System.Collections.Generic.List[string]]::new()
    if (-not [string]::IsNullOrWhiteSpace($baseDir)) {
        $candidates.Add((Join-Path $baseDir 'localization.json'))
        $parentDir = Split-Path -Parent $baseDir
        if (-not [string]::IsNullOrWhiteSpace($parentDir)) {
            $candidates.Add((Join-Path $parentDir 'localization.json'))
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($script:AppDataDir)) {
        $candidates.Add((Join-Path $script:AppDataDir 'localization.json'))
    }

    foreach ($cand in $candidates) {
        if (-not [string]::IsNullOrWhiteSpace($cand) -and (Test-Path $cand)) {
            $locPath = (Resolve-Path $cand).Path
            break
        }
    }

    if ($locPath -and (Test-Path $locPath)) {
        try {
            $jsonRaw = [System.IO.File]::ReadAllText($locPath, [System.Text.Encoding]::UTF8)
            $parsed = $jsonRaw | ConvertFrom-Json
            if ($parsed.Languages) {
                $script:LocCatalog = $parsed
                Write-AppLog "Loaded localization catalog from: $locPath" "OK"
                return $parsed
            }
        }
        catch {
            Write-AppLog "Error parsing $($locPath): $($_.Exception.Message)" "WARN"
        }
    }

    # In-memory fallback
    if ($script:EmbeddedLocalizationJson) {
        try {
            $script:LocCatalog = $script:EmbeddedLocalizationJson | ConvertFrom-Json
            return $script:LocCatalog
        }
        catch { }
    }
    return $null
}

function Get-LocText([string]$Key, [string]$DefaultText = '') {
    if (-not $script:LocCatalog -or -not $script:LocCatalog.Languages) {
        return $(if ($DefaultText) { $DefaultText } else { $Key })
    }
    $lang = if ($script:CurrentLanguage) { $script:CurrentLanguage } else { 'en' }
    $langObj = $script:LocCatalog.Languages.$lang
    if ($langObj -and $langObj.Strings -and ($langObj.Strings.PSObject.Properties.Name -contains $Key)) {
        $val = $langObj.Strings.$Key
        if (-not [string]::IsNullOrWhiteSpace($val)) { return [string]$val }
    }
    # Fallback to EN
    $enObj = $script:LocCatalog.Languages.en
    if ($enObj -and $enObj.Strings -and ($enObj.Strings.PSObject.Properties.Name -contains $Key)) {
        $val = $enObj.Strings.$Key
        if (-not [string]::IsNullOrWhiteSpace($val)) { return [string]$val }
    }
    return $(if ($DefaultText) { $DefaultText } else { $Key })
}

# ══════════════════════════════════════════════════════════════════════════════
# 3. OVERLAY CONFIGURATION DEFAULTS
# ══════════════════════════════════════════════════════════════════════════════

function Get-DefaultOverlayConfig {
    return [ordered]@{
        EnableTopOverlay    = $true
        EnableBottomOverlay = $true
        Properties          = [ordered]@{
            RouteName     = [ordered]@{ Enabled = $true;  Panel = 'Top';    Alignment = 'Left';   Order = 1 }
            RouteType     = [ordered]@{ Enabled = $true;  Panel = 'Top';    Alignment = 'Right';  Order = 1 }
            Timestamp     = [ordered]@{ Enabled = $false; Panel = 'Top';    Alignment = 'Right';  Order = 2 }
            StartGeocoded = [ordered]@{ Enabled = $true;  Panel = 'Bottom'; Alignment = 'Left';   Order = 1 }
            EndGeocoded   = [ordered]@{ Enabled = $true;  Panel = 'Bottom'; Alignment = 'Left';   Order = 2 }
            Distance      = [ordered]@{ Enabled = $true;  Panel = 'Bottom'; Alignment = 'Left';   Order = 3 }
            Duration      = [ordered]@{ Enabled = $true;  Panel = 'Bottom'; Alignment = 'Left';   Order = 3 }
            Waypoints      = [ordered]@{ Enabled = $false; Panel = 'Bottom'; Alignment = 'Left';   Order = 4 }
            PointDistances = [ordered]@{ Enabled = $false; Panel = 'None';   Alignment = 'Left';   Order = 5 }
            StartRaw       = [ordered]@{ Enabled = $false; Panel = 'Bottom'; Alignment = 'Left';   Order = 6 }
            EndRaw         = [ordered]@{ Enabled = $false; Panel = 'Bottom'; Alignment = 'Left';   Order = 7 }
        }
    }
}

# ══════════════════════════════════════════════════════════════════════════════
# 4. APPLICATION CONFIGURATION & API USAGE TRACKING
# ══════════════════════════════════════════════════════════════════════════════

function Get-DefaultAppConfig {
    [CmdletBinding()]
    param()

    $defaultOutput = [System.IO.Path]::Combine([System.Environment]::GetFolderPath('MyDocuments'), 'TrasyGoogleMaps')
    $currentMonth = (Get-Date).ToString('yyyy-MM')

    return [ordered]@{
        ApiKey               = ''
        CartoApiKey          = ''
        RememberKey          = $true
        DefaultRouteType     = 'Fastest'
        DefaultEmissionType  = 'GASOLINE'
        MapWidth             = 900
        MapHeight            = 600
        OutputDirectory      = $defaultOutput
        Language             = 'en'
        Theme                = 'Dark'
        UseInteractiveMap    = $true
        AvoidTolls           = $false
        AvoidHighways        = $false
        AvoidFerries         = $false
        OverlayConfig        = (Get-DefaultOverlayConfig)
        ApiUsage             = [ordered]@{
            CurrentMonth          = $currentMonth
            MonthlyCallsGeocoding = 0
            MonthlyCallsRoutes    = 0
            MonthlyCallsStatic    = 0
            SessionCallsGeocoding = 0
            SessionCallsRoutes    = 0
            SessionCallsStatic    = 0
            BudgetLimitUsd        = 50.0
            PreferredCurrency     = 'USD'
        }
    }
}
Set-Item -Path "function:global:Get-DefaultAppConfig" -Value (Get-Item "function:Get-DefaultAppConfig").ScriptBlock -ErrorAction SilentlyContinue

function Load-AppConfig {
    [CmdletBinding()]
    param()

    $result = Get-DefaultAppConfig

    if (-not (Test-Path $script:ConfigFile)) {
        $script:AppConfig = $result
        $script:CurrentLanguage = $result.Language
        return $result
    }

    try {
        $raw = [System.IO.File]::ReadAllText($script:ConfigFile, [System.Text.Encoding]::UTF8)
        $cfg = $raw | ConvertFrom-Json

        if ($cfg.ApiKey) {
            $decrypted = Unprotect-SecretString -EncryptedText $cfg.ApiKey
            $result.ApiKey = if ($decrypted) { $decrypted } else { '' }
        }
        if ($cfg.CartoApiKey) {
            $decryptedCarto = Unprotect-SecretString -EncryptedText $cfg.CartoApiKey
            $result.CartoApiKey = if ($decryptedCarto) { $decryptedCarto } else { [string]$cfg.CartoApiKey }
        }
        if ($null -ne $cfg.RememberKey)         { $result.RememberKey = [bool]$cfg.RememberKey }
        if ($cfg.DefaultRouteType)             { $result.DefaultRouteType = [string]$cfg.DefaultRouteType }
        if ($cfg.DefaultEmissionType)          { $result.DefaultEmissionType = [string]$cfg.DefaultEmissionType }
        if ($cfg.MapWidth -gt 0)               { $result.MapWidth = [int]$cfg.MapWidth }
        if ($cfg.MapHeight -gt 0)              { $result.MapHeight = [int]$cfg.MapHeight }
        if ($cfg.OutputDirectory)              { $result.OutputDirectory = [string]$cfg.OutputDirectory }
        if ($cfg.Language)                     { $result.Language = [string]$cfg.Language }
        if ($cfg.Theme)                        { $result.Theme = [string]$cfg.Theme }
        if ($null -ne $cfg.UseInteractiveMap)  { $result.UseInteractiveMap = [bool]$cfg.UseInteractiveMap }
        if ($null -ne $cfg.AvoidTolls)         { $result.AvoidTolls = [bool]$cfg.AvoidTolls }
        if ($null -ne $cfg.AvoidHighways)      { $result.AvoidHighways = [bool]$cfg.AvoidHighways }
        if ($null -ne $cfg.AvoidFerries)       { $result.AvoidFerries = [bool]$cfg.AvoidFerries }

        if ($cfg.OverlayConfig) {
            $result.OverlayConfig = $cfg.OverlayConfig
        }

        # API Usage loading
        if ($cfg.ApiUsage) {
            $u = $cfg.ApiUsage
            $currentMonth = (Get-Date).ToString('yyyy-MM')
            if ($u.CurrentMonth -eq $currentMonth) {
                $result.ApiUsage.CurrentMonth          = $currentMonth
                $result.ApiUsage.MonthlyCallsGeocoding = [int]$u.MonthlyCallsGeocoding
                $result.ApiUsage.MonthlyCallsRoutes    = [int]$u.MonthlyCallsRoutes
                $result.ApiUsage.MonthlyCallsStatic    = [int]$u.MonthlyCallsStatic
            } else {
                # New month reset
                $result.ApiUsage.CurrentMonth          = $currentMonth
                $result.ApiUsage.MonthlyCallsGeocoding = 0
                $result.ApiUsage.MonthlyCallsRoutes    = 0
                $result.ApiUsage.MonthlyCallsStatic    = 0
            }
            if ($u.BudgetLimitUsd)      { $result.ApiUsage.BudgetLimitUsd = [double]$u.BudgetLimitUsd }
            if ($u.PreferredCurrency)  { $result.ApiUsage.PreferredCurrency = [string]$u.PreferredCurrency }
        }

        $script:AppConfig = $result
        $script:CurrentLanguage = $result.Language
        return $result
    }
    catch {
        Write-AppLog "Error loading config: $($_.Exception.Message)" "ERROR"
        $script:AppConfig = Get-DefaultAppConfig
        $script:CurrentLanguage = $script:AppConfig.Language
        return $script:AppConfig
    }
}
Set-Item -Path "function:global:Load-AppConfig" -Value (Get-Item "function:Load-AppConfig").ScriptBlock -ErrorAction SilentlyContinue

function Save-AppConfig {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][object]$Config)

    try {
        $plainKey = if ($Config.ApiKey) { [string]$Config.ApiKey } else { '' }
        $plainCartoKey = if ($Config.CartoApiKey) { [string]$Config.CartoApiKey } else { '' }
        $remember = if ($null -ne $Config.RememberKey) { [bool]$Config.RememberKey } else { $true }

        $encryptedKey = if ($remember -and -not [string]::IsNullOrWhiteSpace($plainKey)) {
            Protect-SecretString -PlainText $plainKey
        } else {
            ''
        }

        $encryptedCartoKey = if ($remember -and -not [string]::IsNullOrWhiteSpace($plainCartoKey)) {
            Protect-SecretString -PlainText $plainCartoKey
        } else {
            ''
        }

        $toSave = [ordered]@{
            ApiKey              = $encryptedKey
            CartoApiKey         = $encryptedCartoKey
            RememberKey         = $remember
            DefaultRouteType    = [string]$Config.DefaultRouteType
            DefaultEmissionType = [string]$Config.DefaultEmissionType
            MapWidth            = [int]$Config.MapWidth
            MapHeight           = [int]$Config.MapHeight
            OutputDirectory     = [string]$Config.OutputDirectory
            Language            = [string]$Config.Language
            Theme               = [string]$Config.Theme
            UseInteractiveMap   = [bool]$Config.UseInteractiveMap
            AvoidTolls          = [bool]$Config.AvoidTolls
            AvoidHighways       = [bool]$Config.AvoidHighways
            AvoidFerries        = [bool]$Config.AvoidFerries
            OverlayConfig       = $Config.OverlayConfig
            ApiUsage            = $Config.ApiUsage
        }

        $json = $toSave | ConvertTo-Json -Depth 10
        [System.IO.File]::WriteAllText($script:ConfigFile, $json, [System.Text.UTF8Encoding]::new($true))
        $script:AppConfig = $Config
        if ($Config.Language) { $script:CurrentLanguage = [string]$Config.Language }
        if ($Config.Theme)    { $script:CurrentTheme = [string]$Config.Theme }
        Write-AppLog "Configuration saved to: $script:ConfigFile" "OK"
        return $true
    }
    catch {
        Write-AppLog "Error saving configuration: $($_.Exception.Message)" "ERROR"
        return $false
    }
}
Set-Item -Path "function:global:Save-AppConfig" -Value (Get-Item "function:Save-AppConfig").ScriptBlock -ErrorAction SilentlyContinue

function Update-ApiUsageRecord {
    [CmdletBinding()]
    param(
        [Parameter()][int]$GeocodingInc = 0,
        [Parameter()][int]$RoutesInc = 0,
        [Parameter()][int]$StaticMapsInc = 0
    )

    if (-not $script:AppConfig -or -not $script:AppConfig.ApiUsage) { return }

    # If in-memory ApiKey is empty, attempt to preserve key from UI or environment
    if ([string]::IsNullOrWhiteSpace($script:AppConfig.ApiKey)) {
        $activeKey = Get-CurrentApiKey
        if (-not [string]::IsNullOrWhiteSpace($activeKey)) {
            $script:AppConfig.ApiKey = $activeKey
        }
    }

    $u = $script:AppConfig.ApiUsage
    $currentMonth = (Get-Date).ToString('yyyy-MM')
    if ($u.CurrentMonth -ne $currentMonth) {
        $u.CurrentMonth = $currentMonth
        $u.MonthlyCallsGeocoding = 0
        $u.MonthlyCallsRoutes = 0
        $u.MonthlyCallsStatic = 0
    }

    $u.SessionCallsGeocoding += $GeocodingInc
    $u.SessionCallsRoutes    += $RoutesInc
    $u.SessionCallsStatic    += $StaticMapsInc

    $u.MonthlyCallsGeocoding += $GeocodingInc
    $u.MonthlyCallsRoutes    += $RoutesInc
    $u.MonthlyCallsStatic    += $StaticMapsInc

    Save-AppConfig -Config $script:AppConfig | Out-Null
}
Set-Item -Path "function:global:Update-ApiUsageRecord" -Value (Get-Item "function:Update-ApiUsageRecord").ScriptBlock -ErrorAction SilentlyContinue
