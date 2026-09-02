# Process-SchoolTransportRoutes — Architecture

<!-- AUTO:metadata -->
- **Script Path:** `D:\Skrypty\Mnich_Adam_Skrypty\Mapy Google\Process-SchoolTransportRoutes.ps1`
- **Last Synced:** `2026-09-02` (`Uncommitted`)
- **Type:** PowerShell Script (.ps1, Advanced Function)
<!-- /AUTO -->

## Overview
<!-- AUTO:overview -->
Przetwarza umowy na dowozy szkolne — generuje mapy tras i podsumowania Excel. Skrypt wczytuje plik Excel z kolumnami: Umowa, Szkoła, Dom, Praca, Tryb, Wariant. Dla każdej umowy tworzy osobny folder w OutputFolder, generuje mapy PNG tras (Dom↔Szkoła, opcjonalnie z Pracą) oraz plik Excel z podsumowaniem odległości. Na końcu tworzy zbiorczy Excel ze wszystkimi umowami.
<!-- /AUTO -->

## Parameters
<!-- AUTO:parameters -->
| Name | Type | Mandatory | Default | Description |
| :--- | :--- | :--- | :--- | :--- |
| `-ApiKey` | `[String]` | Yes | `$env:GOOGLE_MAPS_API_KEY` | Klucz Google Maps API. Domyślnie pobierany ze zmiennej środowiskowej GOOGLE_MAPS_API_KEY. |
| `-InputExcel` | `[String]` | Yes | `$null` | Ścieżka do pliku Excel wejściowego. Jeśli nie podano, otwiera się dialog wyboru pliku. |
| `-OutputFolder` | `[String]` | Yes | `'C:\Temp\SchoolRoutes'` | Folder wynikowy dla podfolderów umów. Domyślnie C:\Temp\SchoolRoutes |
| `-MapWidth` | `[Int32]` | Yes | `900` | Szerokość mapy PNG w pikselach. Domyślnie 900. |
| `-MapHeight` | `[Int32]` | Yes | `600` | Wysokość mapy PNG w pikselach. Domyślnie 600. |
<!-- /AUTO -->

## Functions
<!-- AUTO:functions -->
| Name | Purpose | Calls |
| :--- | :--- | :--- |
| `ConvertTo-SafeFileName` | Internal routine supporting script workflow | `Internal Logic` |
| `Test-PracaAddress` | Internal routine supporting script workflow | `Internal Logic` |
| `Invoke-RouteAndMap` | Internal routine supporting script workflow | `Get-CarRouteData`, `Write-Warning`, `Write-Host` |
| `Find-ColumnHeader` | Internal routine supporting script workflow | `Where-Object`, `Select-Object` |
<!-- /AUTO -->

## Dependencies
<!-- AUTO:dependencies -->
- **Modules:** `ActiveDirectory`, `ImportExcel`
- **External Tools:** *None*
- **Config & Environment:** Environment variables (`$env:COMPUTERNAME`, `$env:USERDOMAIN`)
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
.\Process-SchoolTransportRoutes.ps1
    Uruchamia skrypt z dialogiem wyboru pliku.
```
<!-- /AUTO -->

## Notes
<!-- MANUAL -->
<!-- Agents: preserve this section verbatim during updates. -->
- Generowane mapy tras PNG zawierają górną belkę informacyjną z numerem umowy oraz kierunkiem (np. Praca -> Dom, Dom -> Szkoła).
<!-- /MANUAL -->

