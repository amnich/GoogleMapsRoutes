#Requires -Version 5.1
<#
.SYNOPSIS
    Compiles GoogleMapsRoutes-GUI.ps1 or Process-SchoolTransportRoutes-GUI.ps1
    into a self-sufficient standalone executable (.EXE) using the PS2EXE module.

.DESCRIPTION
    Automatically integrates all helper scripts (.ps1) and translation files (localization.json)
    into a single unified, self-contained PowerShell script prior to compilation.
    The resulting executable operates anywhere with zero external file dependencies.
    Automatically verifies the presence of the ps2exe module (installs from PSGallery if needed)
    and produces a windowed application without console (-noConsole) in Single-Threaded Apartment (-sta) mode.

.PARAMETER Target
    Target application to compile:
      - 'GoogleMapsRoutes'       : Main universal routing application (GoogleMapsRoutes-GUI.ps1) -> GoogleMapsRoutes.exe
      - 'SchoolTransportRoutes'  : School transport contract processor (Process-SchoolTransportRoutes-GUI.ps1) -> SchoolTransportRoutes.exe

.PARAMETER InputScript
    Path to input script (optional, overrides -Target).

.PARAMETER OutputFile
    Path to output executable (.EXE) (optional, overrides -Target).

.PARAMETER IconFile
    Path to .ico icon file (optional).

.PARAMETER NoBundle
    Skip the automated script bundling step and compile the input script directly.

.PARAMETER BundledScriptPath
    Custom path for the generated intermediate bundled script (optional, defaults to Build\<Target>.bundled.ps1).

.PARAMETER KeepBundledScript
    Preserve the intermediate bundled .ps1 script in the Build directory for audit/inspection (default: $true).

.EXAMPLE
    .\Build-Exe.ps1
    Bundles helper scripts and localization.json, then compiles GoogleMapsRoutes.exe

.EXAMPLE
    .\Build-Exe.ps1 -Target SchoolTransportRoutes
    Bundles dependencies and compiles SchoolTransportRoutes.exe

.EXAMPLE
    .\Build-Exe.ps1 -InputScript .\Get-CarRoute_WithMap.ps1 -OutputFile .\Get-CarRoute.exe
    Inlines RouteMapFunctions.ps1 and compiles a standalone CLI executable

.NOTES
    Encoding: UTF-8 with BOM
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet('GoogleMapsRoutes', 'SchoolTransportRoutes')]
    [string]$Target = 'GoogleMapsRoutes',

    [Parameter(Mandatory = $false)]
    [string]$InputScript,

    [Parameter(Mandatory = $false)]
    [string]$OutputFile,

    [Parameter(Mandatory = $false)]
    [string]$IconFile = "D:\Skrypty\GoogleMapsRoutes\Res\Logo_AM6.ico",

    [Parameter(Mandatory = $false)]
    [switch]$NoBundle,

    [Parameter(Mandatory = $false)]
    [string]$BundledScriptPath,

    [Parameter(Mandatory = $false)]
    [bool]$KeepBundledScript = $true
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }

# ══════════════════════════════════════════════════════════════════════════════
# 1. SCRIPT BUNDLER PIPELINE (SELF-SUFFICIENT PS2EXE INTEGRATION)
# ══════════════════════════════════════════════════════════════════════════════

function New-BundledScript {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceScriptPath,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath,

        [Parameter(Mandatory = $false)]
        [string]$LocalizationJsonPath
    )

    $resolvedSource = (Resolve-Path $SourceScriptPath).Path
    $sourceDir = Split-Path -Parent $resolvedSource
    $outDir = Split-Path -Parent $OutputPath
    if (-not (Test-Path $outDir)) {
        New-Item -ItemType Directory -Path $outDir -Force | Out-Null
    }

    $inlinedFiles = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $inlinedList = [System.Collections.Generic.List[PSCustomObject]]::new()

    function Resolve-AndInlineContent {
        param(
            [string]$Content,
            [string]$CurrentDir
        )

        $patterns = @(
            # Pattern 1: . "$PSScriptRoot\Helper.ps1" or . '.\Helper.ps1'
            '(?m)^\s*\.\s+["'']?((?:\$PSScriptRoot|\$ScriptDir|\.)[\\/]([^\r\n"'']+?\.ps1))["'']?\s*$',
            # Pattern 2: . (Join-Path $PSScriptRoot 'Helper.ps1')
            '(?m)^\s*\.\s*\(\s*Join-Path\s+(?:\$PSScriptRoot|\$ScriptDir)\s+[''"]([^''"]+\.ps1)[''"]\s*\)\s*$',
            # Pattern 3: $FunctionsPath = Join-Path $PSScriptRoot 'Helper.ps1' ... . $FunctionsPath
            '(?s)\$([a-zA-Z0-9_]+)\s*=\s*Join-Path\s+(?:\$PSScriptRoot|\$ScriptDir)\s+[''"]([^''"]+\.ps1)[''"].*?\.\s+\$\1\b'
        )

        $updatedContent = $Content

        # Process Pattern 3 first (multiline variable assignment + dot-source)
        $m3 = [regex]::Matches($updatedContent, $patterns[2])
        foreach ($match in $m3) {
            $helperLeaf = $match.Groups[2].Value
            $helperFullPath = [System.IO.Path]::GetFullPath((Join-Path $CurrentDir $helperLeaf))
            if (Test-Path $helperFullPath) {
                $helperCanonical = (Resolve-Path $helperFullPath).Path
                if (-not $inlinedFiles.Contains($helperCanonical)) {
                    $inlinedFiles.Add($helperCanonical) | Out-Null
                    $helperRaw = [System.IO.File]::ReadAllText($helperCanonical, [System.Text.Encoding]::UTF8)
                    $lineCount = ($helperRaw -split '\r?\n').Count
                    $inlinedList.Add([PSCustomObject]@{ Name = $helperLeaf; Lines = $lineCount; Path = $helperCanonical }) | Out-Null

                    # Recursively inline nested dependencies
                    $helperInlined = Resolve-AndInlineContent -Content $helperRaw -CurrentDir (Split-Path -Parent $helperCanonical)
                    # Strip script-level #Requires from helper
                    $helperInlined = $helperInlined -replace '(?m)^\s*#Requires\s+-[^\r\n]+', '# [Bundler stripped nested #Requires]'

                    $replacement = @"
# ==============================================================================
# >>> BUNDLER: BEGIN INLINED HELPER [$helperLeaf]
# >>> Source: $helperCanonical ($lineCount lines)
# ==============================================================================
$helperInlined
# ==============================================================================
# <<< BUNDLER: END INLINED HELPER [$helperLeaf]
# ==============================================================================
"@
                    $updatedContent = $updatedContent.Replace($match.Value, $replacement)
                } else {
                    $updatedContent = $updatedContent.Replace($match.Value, "# [Bundler: Already inlined $helperLeaf]")
                }
            }
        }

        # Process Patterns 1 & 2
        for ($pIdx = 0; $pIdx -le 1; $pIdx++) {
            $p = $patterns[$pIdx]
            $matches = [regex]::Matches($updatedContent, $p)
            foreach ($match in $matches) {
                $helperLeaf = if ($pIdx -eq 0) { $match.Groups[2].Value } else { $match.Groups[1].Value }
                $helperFullPath = [System.IO.Path]::GetFullPath((Join-Path $CurrentDir $helperLeaf))
                if (Test-Path $helperFullPath) {
                    $helperCanonical = (Resolve-Path $helperFullPath).Path
                    if (-not $inlinedFiles.Contains($helperCanonical)) {
                        $inlinedFiles.Add($helperCanonical) | Out-Null
                        $helperRaw = [System.IO.File]::ReadAllText($helperCanonical, [System.Text.Encoding]::UTF8)
                        $lineCount = ($helperRaw -split '\r?\n').Count
                        $inlinedList.Add([PSCustomObject]@{ Name = $helperLeaf; Lines = $lineCount; Path = $helperCanonical }) | Out-Null

                        # Recursively inline nested dependencies
                        $helperInlined = Resolve-AndInlineContent -Content $helperRaw -CurrentDir (Split-Path -Parent $helperCanonical)
                        # Strip script-level #Requires from helper
                        $helperInlined = $helperInlined -replace '(?m)^\s*#Requires\s+-[^\r\n]+', '# [Bundler stripped nested #Requires]'

                        $replacement = @"
# ==============================================================================
# >>> BUNDLER: BEGIN INLINED HELPER [$helperLeaf]
# >>> Source: $helperCanonical ($lineCount lines)
# ==============================================================================
$helperInlined
# ==============================================================================
# <<< BUNDLER: END INLINED HELPER [$helperLeaf]
# ==============================================================================
"@
                        $updatedContent = $updatedContent.Replace($match.Value, $replacement)
                    } else {
                        $updatedContent = $updatedContent.Replace($match.Value, "# [Bundler: Already inlined $helperLeaf]")
                    }
                }
            }
        }

        return $updatedContent
    }

    Write-Host "`nBundling dependencies for self-sufficient execution..." -ForegroundColor Yellow
    Write-Host "  Source file : $resolvedSource" -ForegroundColor Gray

    $rawContent = [System.IO.File]::ReadAllText($resolvedSource, [System.Text.Encoding]::UTF8)
    $bundled = Resolve-AndInlineContent -Content $rawContent -CurrentDir $sourceDir

    if ($inlinedList.Count -gt 0) {
        foreach ($item in $inlinedList) {
            Write-Host "  + Inlined helper script: $($item.Name) ($($item.Lines) lines)" -ForegroundColor Green
        }
    } else {
        Write-Host "  * No external dot-sourced helper scripts found." -ForegroundColor DarkGray
    }

    # Embed Localization JSON if applicable
    $locPath = if ($LocalizationJsonPath -and (Test-Path $LocalizationJsonPath)) {
        (Resolve-Path $LocalizationJsonPath).Path
    } elseif (Test-Path (Join-Path $sourceDir 'localization.json')) {
        (Resolve-Path (Join-Path $sourceDir 'localization.json')).Path
    } elseif (Test-Path (Join-Path $ScriptDir 'localization.json')) {
        (Resolve-Path (Join-Path $ScriptDir 'localization.json')).Path
    } else { $null }

    $embeddedLocInfo = $null
    if ($locPath -and ($bundled -match 'localization\.json' -or $bundled -match 'EmbeddedLocalizationJson' -or $bundled -match 'Load-LocalizationConfig')) {
        $locJsonRaw = [System.IO.File]::ReadAllText($locPath, [System.Text.Encoding]::UTF8).Trim()
        $locParsed = $locJsonRaw | ConvertFrom-Json
        $langNames = @($locParsed.Languages.PSObject.Properties.Name)
        $embeddedLocInfo = [PSCustomObject]@{
            Path      = $locPath
            Languages = $langNames
            SizeKb    = [math]::Round($locJsonRaw.Length / 1KB, 1)
        }

        $embeddedBlock = @"
# ==============================================================================
# BUNDLER: EMBEDDED LOCALIZATION CATALOG (SELF-SUFFICIENT PS2EXE EXECUTION)
# ==============================================================================
`$script:EmbeddedLocalizationJson = @'
$locJsonRaw
'@
"@

        $existingPattern = '(?s)# ==============================================================================\s*# BUNDLER: EMBEDDED LOCALIZATION CATALOG.*?\r?\n''@'
        if ($bundled -match $existingPattern) {
            $bundled = [regex]::Replace($bundled, $existingPattern, $embeddedBlock)
        } elseif ($bundled -match '(?m)^\s*function\s+Load-LocalizationConfig\b') {
            $bundled = [regex]::Replace($bundled, '(?m)^\s*function\s+Load-LocalizationConfig\b', "$embeddedBlock`r`n`r`nfunction Load-LocalizationConfig")
        } else {
            $bundled = "$embeddedBlock`r`n`r`n$bundled"
        }

        Write-Host "  + Embedded localization catalog: $(Split-Path $locPath -Leaf) ($($langNames -join ', '), $($embeddedLocInfo.SizeKb) KB)" -ForegroundColor Green
    }

    # AST Syntax Validation
    $tokens = $null
    $errors = $null
    $null = [System.Management.Automation.Language.Parser]::ParseInput($bundled, [ref]$tokens, [ref]$errors)
    if ($errors -and $errors.Count -gt 0) {
        $firstErr = $errors[0]
        throw "Bundled script syntax error at line $($firstErr.Extent.StartLineNumber), col $($firstErr.Extent.StartColumnNumber): $($firstErr.Message)"
    }
    Write-Host "  + Syntax validation: PASSED (0 errors, $($tokens.Count) tokens)" -ForegroundColor Green

    # Write output file as UTF-8 with BOM (with retry loop for transient locks)
    $written = $false
    for ($retry = 0; $retry -lt 5; $retry++) {
        try {
            [System.IO.File]::WriteAllText($OutputPath, $bundled, [System.Text.UTF8Encoding]::new($true))
            $written = $true
            break
        } catch [System.IO.IOException] {
            Start-Sleep -Milliseconds 400
        }
    }
    if (-not $written) {
        [System.IO.File]::WriteAllText($OutputPath, $bundled, [System.Text.UTF8Encoding]::new($true))
    }

    $outItem = Get-Item $OutputPath
    $sizeKb = [math]::Round($outItem.Length / 1KB, 1)
    Write-Host "  + Bundled script written: $OutputPath ($sizeKb KB)`n" -ForegroundColor Green

    return [PSCustomObject]@{
        OutputPath           = $OutputPath
        SourcePath           = $resolvedSource
        InlinedHelpers       = $inlinedList.ToArray()
        EmbeddedLocalization = $embeddedLocInfo
        SizeKb               = $sizeKb
    }
}

# ══════════════════════════════════════════════════════════════════════════════
# 2. TARGET CONFIGURATION & PATH RESOLUTION
# ══════════════════════════════════════════════════════════════════════════════

if (-not $InputScript) {
    $InputScript = if ($Target -eq 'SchoolTransportRoutes') {
        Join-Path $ScriptDir 'Process-SchoolTransportRoutes-GUI.ps1'
    } else {
        Join-Path $ScriptDir 'GoogleMapsRoutes-GUI.ps1'
    }
}

if (-not $OutputFile) {
    $OutputFile = if ($Target -eq 'SchoolTransportRoutes') {
        Join-Path $ScriptDir 'SchoolTransportRoutes.exe'
    } else {
        Join-Path $ScriptDir 'GoogleMapsRoutes.exe'
    }
}

if (-not $BundledScriptPath) {
    $buildFolder = Join-Path $ScriptDir 'Build'
    $scriptBase = [System.IO.Path]::GetFileNameWithoutExtension($InputScript)
    $BundledScriptPath = Join-Path $buildFolder "$scriptBase.bundled.ps1"
}

$AppTitle = if ($Target -eq 'SchoolTransportRoutes') {
    'School Transport Route & Map Generator'
} else {
    'Google Maps Route & Map Generator'
}

$AppDesc = if ($Target -eq 'SchoolTransportRoutes') {
    'Tool for calculating Google Maps routes and generating maps for school transport contracts'
} else {
    'Generic Google Maps multi-point route calculator and map generator supporting Fastest, Shortest, and Eco routes with JSON, CSV, and Excel input'
}

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " PS2EXE SELF-SUFFICIENT COMPILATION: $Target" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

# Close any running instance of the output executable to avoid file lock
$targetProcName = [System.IO.Path]::GetFileNameWithoutExtension($OutputFile)
$runningProc = Get-Process -Name $targetProcName -ErrorAction SilentlyContinue
if ($runningProc) {
    Write-Host "Stopping running instance of $targetProcName (PID: $($runningProc.Id))..." -ForegroundColor Yellow
    $runningProc | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 600
}

# 3. Verify PS2EXE module
if (-not (Get-Module -ListAvailable -Name ps2exe)) {
    Write-Host "ps2exe module not found. Installing from PSGallery (CurrentUser)..." -ForegroundColor Yellow
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls
    if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -ErrorAction Stop | Out-Null
    }
    Install-Module -Name ps2exe -Scope CurrentUser -Force -AllowClobber -Repository PSGallery -ErrorAction Stop
    Write-Host "ps2exe module installed successfully." -ForegroundColor Green
}

Import-Module -Name ps2exe -Force

# 4. Validate source script
if (-not (Test-Path $InputScript)) {
    throw "Source file does not exist: $InputScript"
}

# 5. Execute Bundling Pipeline (integrate helper .ps1 and localization.json)
$CompileInput = $InputScript
if (-not $NoBundle) {
    $bundleResult = New-BundledScript -SourceScriptPath $InputScript -OutputPath $BundledScriptPath
    $CompileInput = $bundleResult.OutputPath
} else {
    Write-Host "`n[Notice] Script bundling bypassed (-NoBundle). Using raw input script." -ForegroundColor Yellow
}

Write-Host "Compilation Source : $CompileInput" -ForegroundColor White
Write-Host "Compilation Output : $OutputFile" -ForegroundColor White

# 6. PS2EXE compilation parameters
$CompileParams = @{
    InputFile   = $CompileInput
    OutputFile  = $OutputFile
    NoConsole   = $true
    STA         = $true
    Title       = $AppTitle
    Description = $AppDesc
    Company     = 'Adam Mnich'
    Product     = $Target
    Copyright   = 'Copyright (c) 2026'
    Version     = '2.0.0.0'
    NoError     = $true
    NoOutput    = $true
}

if (-not [string]::IsNullOrWhiteSpace($IconFile) -and (Test-Path $IconFile)) {
    $CompileParams['IconFile'] = $IconFile
    Write-Host "Icon file          : $IconFile" -ForegroundColor White
}

# 7. Invoke compiler
Write-Host "`nStarting PS2EXE compilation..." -ForegroundColor Yellow
Invoke-PS2EXE @CompileParams

# Clean up bundled script if requested
if (-not $KeepBundledScript -and -not $NoBundle -and (Test-Path $BundledScriptPath)) {
    Remove-Item -Path $BundledScriptPath -Force -ErrorAction SilentlyContinue
}

# 8. Report results
if (Test-Path $OutputFile) {
    $item = Get-Item $OutputFile
    $sizeKb = [math]::Round($item.Length / 1KB, 1)
    Write-Host "`n==========================================================" -ForegroundColor Green
    Write-Host " [SUCCESS] Executable created successfully!" -ForegroundColor Green
    Write-Host "==========================================================" -ForegroundColor Green
    Write-Host "  Executable : $($item.FullName)" -ForegroundColor Green
    Write-Host "  File Size  : $sizeKb KB" -ForegroundColor Green
    if (-not $NoBundle -and $KeepBundledScript) {
        Write-Host "  Bundled PS1: $BundledScriptPath" -ForegroundColor DarkGray
    }
    Write-Host "  Standalone : 100% self-sufficient (no external .ps1 or .json needed)`n" -ForegroundColor Cyan
} else {
    throw "Compilation completed, but output file $OutputFile was not found."
}
