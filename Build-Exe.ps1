#Requires -Version 5.1
<#
.SYNOPSIS
    Compiles the Process-SchoolTransportRoutes-GUI.ps1 script into a standalone EXE using PS2EXE.

.DESCRIPTION
    The script automatically verifies the presence of the ps2exe module (installs from PSGallery if needed)
    and produces a windowed application without a console (-noConsole) running in single-threaded apartment mode (-sta).

.NOTES
    Encoding: UTF-8 with BOM
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$InputScript,

    [Parameter(Mandatory = $false)]
    [string]$OutputFile,

    [Parameter(Mandatory = $false)]
    [string]$IconFile
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
if (-not $InputScript) { $InputScript = Join-Path $ScriptDir 'Process-SchoolTransportRoutes-GUI.ps1' }
if (-not $OutputFile)  { $OutputFile  = Join-Path $ScriptDir 'SchoolTransportRoutes.exe' }

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " PS2EXE COMPILATION: SchoolTransportRoutes" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

# 1. Check for PS2EXE module
if (-not (Get-Module -ListAvailable -Name ps2exe)) {
    Write-Host "ps2exe module not found. Installing from PSGallery for CurrentUser..." -ForegroundColor Yellow
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls
    if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -ErrorAction Stop | Out-Null
    }
    Install-Module -Name ps2exe -Scope CurrentUser -Force -AllowClobber -Repository PSGallery -ErrorAction Stop
    Write-Host "ps2exe module installed successfully." -ForegroundColor Green
}

Import-Module -Name ps2exe -Force

# 2. Validate input script
if (-not (Test-Path $InputScript)) {
    throw "Input file does not exist: $InputScript"
}

Write-Host "Source file : $InputScript" -ForegroundColor White
Write-Host "Output file : $OutputFile" -ForegroundColor White

# 3. PS2EXE compilation parameters
$CompileParams = @{
    InputFile       = $InputScript
    OutputFile      = $OutputFile
    NoConsole       = $true
    STA             = $true
    Title           = 'School Transport Route & Map Generator'
    Description     = 'Tool for calculating Google Maps routes and generating maps for school transport contracts'
    Company         = 'Admin'
    Product         = 'School Transport Routes'
    Copyright       = 'Copyright (c) 2026'
    Version         = '1.0.0.0'
    NoError         = $true
    NoOutput        = $true
}

if (-not [string]::IsNullOrWhiteSpace($IconFile) -and (Test-Path $IconFile)) {
    $CompileParams['IconFile'] = $IconFile
    Write-Host "Icon file   : $IconFile" -ForegroundColor White
}

# 4. Invoke PS2EXE compiler
Write-Host "`nStarting compilation..." -ForegroundColor Yellow
Invoke-PS2EXE @CompileParams

if (Test-Path $OutputFile) {
    $item = Get-Item $OutputFile
    $sizeKb = [math]::Round($item.Length / 1KB, 1)
    Write-Host "`n[SUCCESS] Executable created successfully!" -ForegroundColor Green
    Write-Host "Path: $($item.FullName)" -ForegroundColor Green
    Write-Host "Size: $sizeKb KB" -ForegroundColor Green
} else {
    throw "Compilation completed, but output file $OutputFile was not found."
}
