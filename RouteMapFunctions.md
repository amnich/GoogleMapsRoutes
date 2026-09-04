# RouteMapFunctions — Architecture

<!-- AUTO:metadata -->
- **Script Path:** `D:\Skrypty\GoogleMapsRoutes\RouteMapFunctions.ps1`
- **Last Synced:** `2026-09-04`
- **Type:** PowerShell Script (.ps1)
<!-- /AUTO -->

## Overview
<!-- AUTO:overview -->
Wspólne funkcje do geokodowania adresów, obliczania tras (Fastest, Shortest, Eco), generowania map poglądowych PNG oraz uniwersalnego importu i eksportu danych tras (JSON, CSV, Excel). Moduł wykorzystywany przez `GoogleMapsRoutes-GUI.ps1`, `Invoke-GoogleMapsRoute.ps1`, `Get-CarRoute_WithMap.ps1`, `Get-MultiPointCarRoute_WithMap.ps1` oraz `Process-SchoolTransportRoutes.ps1`.
<!-- /AUTO -->

## Functions
<!-- AUTO:functions -->
| Name | Purpose | Parameters | Returns |
| :--- | :--- | :--- | :--- |
| `Select-InputDataFile` | Okno dialogowe wyboru pliku z danymi (JSON, CSV, TSV, XLSX, XLS) | *None* | `[string]` Ścieżka pliku |
| `Select-InputExcel` | Okno dialogowe wyboru pliku Excel (.xlsx, .xls) | *None* | `[string]` Ścieżka pliku |
| `Protect-SecretString` | Szyfruje tekst (klucz API) za pomocą Windows DPAPI per-user | `-PlainText` | `[string]` Base64 szyfrogramu |
| `Unprotect-SecretString` | Odszyfrowuje tekst z Windows DPAPI per-user | `-EncryptedText` | `[string]` Tekst jawny |
| `Test-GoogleApiKey` | Sprawdza poprawność klucza API wykonując zapytanie testowe | `-ApiKey` | `[PSCustomObject]` (Valid, Message) |
| `Get-AddressCoordinates` | Geokoduje adres na współrzędne geograficzne (lat/lng) lub weryfikuje podane współrzędne | `-Address`, `-ApiKey`, `-RequireStreetNumber` | `[PSCustomObject]` (Latitude, Longitude, FormattedAddress, Status) |
| `Get-CarRouteData` | Oblicza trasę samochodową (Fastest, Shortest, Eco) przez Google Routes API v2 | `-OriginLat`, `-OriginLng`, `-DestLat`, `-DestLng`, `-ApiKey`, `-IntermediatePoints`, `-RouteType`, `-EmissionType`, `-TrafficAware` | `[PSCustomObject]` (OdlegloscKm, CzasMin, EncodedPolyline, RouteType, Status) |
| `Get-GoogleMapsUrl` | Generuje bezpośredni link do trasy w Google Maps z uwzględnieniem punktów pośrednich | `-Origin`, `-Destination`, `-Waypoints`, `-TravelMode` | `[string]` URL |
| `Save-RouteMapPng` | Pobiera mapę PNG przez Static Maps API z wieloma znacznikami (A, 1..N, B) i konfigurowalną nakładką | `-EncodedPolyline`, `-OriginLat`, `-OriginLng`, `-DestLat`, `-DestLng`, `-OutputPath`, `-ApiKey`, `-RoutePoints`, `-Width`, `-Height`, `-AddressTextA`, `-AddressTextB`, `-DistanceText`, `-DurationText`, `-HeaderLeftText`, `-HeaderRightText`, `-OverlayConfig` (aliasy: `-TekstAdresA`, `-TekstAdresB`, `-TekstOdleglosc`, `-TekstCzas`, `-TekstNaglowekLewy`, `-TekstNaglowekPrawy`, `-TekstUmowa`, `-TekstKierunek`, `-Opis`, `-DataWygenerowania`) | `[bool]` Sukces zapisu |
| `Import-RouteDataFile` | Wczytuje i automatycznie mapuje trasy z plików JSON, CSV, TSV oraz Excel | `-Path`, `-Delimiter` | `[PSCustomObject]` (Mode, Routes, Format, TotalCount) |
| `Export-RouteResults` | Zapisuje wyniki obliczeń do formatu Excel (.xlsx), CSV (.csv) lub JSON (.json) | `-Results`, `-OutputPath`, `-Format` | `[string]` Ścieżka pliku wyjściowego |
<!-- /AUTO -->

## Dependencies
<!-- AUTO:dependencies -->
- **Modules:** `ImportExcel` (opcjonalny, do obsługi plików .xlsx)
- **External Tools:** Google Maps API (Routes API v2, Geocoding API, Static Maps API)
- **Infrastructure:** Windows Presentation Framework (WPF), System.Drawing (GDI+), System.Security (DPAPI)
<!-- /AUTO -->

## Notes
<!-- MANUAL -->
<!-- Agents: preserve this section verbatim during updates. -->
- Funkcja `Save-RouteMapPng` obsługuje zarówno trasy proste (A -> B), jak i trasy wielopunktowe (A -> 1..N -> B).
- Formatowanie tekstu nakładki graficznej GDI+ jest realizowane dynamicznie na podstawie obliczonych linii adresu i wielkości czcionki Segoe UI.
<!-- /MANUAL -->
