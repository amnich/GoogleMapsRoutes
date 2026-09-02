# RouteMapFunctions — Architecture

<!-- AUTO:metadata -->
- **Script Path:** `D:\Skrypty\Mnich_Adam_Skrypty\Mapy Google\RouteMapFunctions.ps1`
- **Last Synced:** `2026-09-02` (`Uncommitted`)
- **Type:** PowerShell Script (.ps1)
<!-- /AUTO -->

## Overview
<!-- AUTO:overview -->
Wspólne funkcje do geokodowania adresów, obliczania tras i generowania map PNG. Moduł funkcji wykorzystywany przez Get-CarRoute_WithMap.ps1 oraz Process-SchoolTransportRoutes.ps1. Zawiera: - Select-InputExcel — dialog wyboru pliku Excel - Get-AddressCoordinates — geokodowanie adresu przez Google Geocoding API - Get-CarRouteData — obliczanie trasy przez Google Routes API v2 - Save-RouteMapPng — generowanie mapy PNG przez Google Static Maps API
<!-- /AUTO -->

## Parameters
<!-- AUTO:parameters -->
| Name | Type | Mandatory | Default | Description |
| :--- | :--- | :--- | :--- | :--- |
| *None* | N/A | N/A | N/A | Script accepts no CLI parameters; configuration is loaded internally or uses defaults |
<!-- /AUTO -->

## Functions
<!-- AUTO:functions -->
| Name | Purpose | Calls |
| :--- | :--- | :--- |
| `Select-InputExcel` | Internal routine supporting script workflow | `Add-Type` |
| `Get-AddressComponentValue` | Internal routine supporting script workflow | `Where-Object`, `Select-Object` |
| `Get-AddressCoordinates` | Internal routine supporting script workflow | `Write-Verbose`, `Invoke-RestMethod`, `Get-AddressComponentValue` |
| `Get-CarRouteData` | Internal routine supporting script workflow | `Write-Verbose`, `Invoke-RestMethod`, `ConvertTo-Json` |
| `Save-RouteMapPng` | Internal routine supporting script workflow | `Write-Verbose`, `Invoke-WebRequest`, `Out-Null` |
| `Get-WrappedLines` | Internal routine supporting script workflow | `Internal Logic` |
<!-- /AUTO -->

## Dependencies
<!-- AUTO:dependencies -->
- **Modules:** `ActiveDirectory`
- **External Tools:** REST/Web API Client (`System.Net`)
- **Config & Environment:** *None*
- **Infrastructure:** *Local System Execution*
<!-- /AUTO -->

## Data Flow
<!-- AUTO:dataflow -->
1. **Input / Init:** Initializes runtime environment, loads configuration or input parameters, and verifies required modules.
2. **Processing:** Executes core queries, Business Central / SQL Server operations, and data transformations.
3. **Output:** Outputs formatted results, generates file artifacts (CSV/JSON/XLSX), or applies configuration changes.
<!-- /AUTO -->

## Error Handling & Logging
<!-- AUTO:errors-and-logging -->
- **Errors:** Structured `try/catch` error trapping wrapping critical operations
- **Logging:** Standard console stream output (`Write-Host`, `Write-Verbose`, `Write-Warning`)
<!-- /AUTO -->

## Security Notes
<!-- AUTO:security -->
- **Privileges:** Standard user permissions; requires appropriate domain read/write access to target resources
- **Credentials:** Operates under current caller Windows Integrated Security context
- **Sensitive Ops:** Read-only inspection and reporting operations
<!-- /AUTO -->

## Usage Example
<!-- AUTO:usage -->
```powershell
.\RouteMapFunctions.ps1
```
<!-- /AUTO -->

## Notes
<!-- MANUAL -->
<!-- Agents: preserve this section verbatim during updates. -->
- Funkcja `Save-RouteMapPng` obsługuje opcjonalny górny pasek nagłówka informacyjnego (parametry `-TekstUmowa` / `-TekstKierunek` lub `-TekstNaglowekLewy` / `-TekstNaglowekPrawy`).
<!-- /MANUAL -->

