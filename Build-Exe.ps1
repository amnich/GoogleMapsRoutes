#Requires -Version 5.1
<#
.SYNOPSIS
    Compiles GoogleMapsRoutes-GUI.ps1 or Process-SchoolTransportRoutes-GUI.ps1
    into a standalone executable (.EXE) using the PS2EXE module.

.DESCRIPTION
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

.EXAMPLE
    .\Build-Exe.ps1
    Compiles the main application GoogleMapsRoutes.exe

.EXAMPLE
    .\Build-Exe.ps1 -Target SchoolTransportRoutes
    Compiles the dedicated SchoolTransportRoutes.exe

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
    [string]$IconFile
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }

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
Write-Host " PS2EXE COMPILATION: $Target" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

# Close any running instance of the output executable to avoid file lock
$targetProcName = [System.IO.Path]::GetFileNameWithoutExtension($OutputFile)
$runningProc = Get-Process -Name $targetProcName -ErrorAction SilentlyContinue
if ($runningProc) {
    Write-Host "Stopping running instance of $targetProcName (PID: $($runningProc.Id))..." -ForegroundColor Yellow
    $runningProc | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 600
}

# 1. Verify PS2EXE module
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

# 2. Validate source script
if (-not (Test-Path $InputScript)) {
    throw "Source file does not exist: $InputScript"
}

Write-Host "Source file : $InputScript" -ForegroundColor White
Write-Host "Output file : $OutputFile" -ForegroundColor White

# 3. PS2EXE compilation parameters
$CompileParams = @{
    InputFile       = $InputScript
    OutputFile      = $OutputFile
    NoConsole       = $true
    STA             = $true
    Title           = $AppTitle
    Description     = $AppDesc
    Company         = 'Admin'
    Product         = $Target
    Copyright       = 'Copyright (c) 2026'
    Version         = '2.0.0.0'
    NoError         = $true
    NoOutput        = $true
}

if (-not [string]::IsNullOrWhiteSpace($IconFile) -and (Test-Path $IconFile)) {
    $CompileParams['IconFile'] = $IconFile
    Write-Host "Icon file   : $IconFile" -ForegroundColor White
}

# 4. Invoke compiler
Write-Host "`nStarting compilation..." -ForegroundColor Yellow
Invoke-PS2EXE @CompileParams

if (Test-Path $OutputFile) {
    $item = Get-Item $OutputFile
    $sizeKb = [math]::Round($item.Length / 1KB, 1)
    Write-Host "`n[SUCCESS] Executable created successfully!" -ForegroundColor Green
    Write-Host "Path : $($item.FullName)" -ForegroundColor Green
    Write-Host "Size : $sizeKb KB" -ForegroundColor Green
} else {
    throw "Compilation completed, but output file $OutputFile was not found."
}
