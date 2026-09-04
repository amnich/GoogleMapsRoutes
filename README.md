# Google Maps Routes & Map Generator v2.0

Zaawansowane, uniwersalne narzędzie PowerShell do obliczania tras samochodowych (jedno- i wielopunktowych), optymalizacji pod kątem czasu (**Fastest**), odległości (**Shortest**) i zużycia energii (**Eco / Fuel Efficient**) oraz generowania estetycznych map poglądowych PNG za pomocą **Google Maps Platform** (Routes API v2, Geocoding API, Static Maps API).

Aplikacja oferuje dwa niezależne interfejsy:
1. **Nowoczesną aplikację okienkową WPF Dark Mode** (`GoogleMapsRoutes-GUI.ps1`) obsługującą wprowadzanie ręczne oraz przetwarzanie wsadowe.
2. **Zaawansowany skrypt wiersza poleceń (CLI)** (`Invoke-GoogleMapsRoute.ps1`) do skryptowania, automatyzacji i integracji potokowej PowerShell.

---

## Główne Możliwości

- **Wielopunktowe trasy (Multipoint Routing)**:
  - Punkt początkowy (Start / A)
  - Do 25 uporządkowanych punktów pośrednich (Waypoints)
  - Punkt końcowy (Cel / B)
- **Typy optymalizacji trasy**:
  - `Fastest` (**Najszybsza**): Minimalizuje łączny czas przejazdu (z opcją Live Traffic).
  - `Shortest` (**Najkrótsza**): Wybiera najkrótszy fizyczny dystans (km) spośród alternatywnych tras Google.
  - `Eco` (**Ekologiczna / Fuel Efficient**): Żąda od Google Routes API trasy o najniższym zużyciu paliwa i emisji CO2 z uwzględnieniem typu napędu (`GASOLINE`, `DIESEL`, `HYBRID`, `ELECTRIC`).
- **Uniwersalne źródła danych**:
  - **JSON** (`.json`): Obsługuje listę tras oraz sekwencję przystanków pojedynczej podróży.
  - **CSV / TSV** (`.csv`, `.tsv`): Automatyczne wykrywanie separatorów (średnik, przecinek, tabulator) i polskich/angielskich nagłówków.
  - **Excel** (`.xlsx`, `.xls`): Pełne wsparcie dla arkuszy kalkulacyjnych z automatycznym filtrowaniem i stylizacją.
- **Ręczne wprowadzanie parametrów (Manual Input)**:
  - W wierszu poleceń (parametry `-StartPoint`, `-EndPoint`, `-Waypoints`).
  - W interfejsie graficznym (wyszukiwanie adresów, dynamiczna lista przystanków, zmiana kolejności góra/dół).
- **Wizualizacja i podgląd**:
  - Generowanie mapy PNG (Google Static Maps) z zielonym punktem `A`, niebieskimi numerowanymi przystankami `1..N`, czerwonym punktem `B` i górną/dolną belką informacyjną.
  - Generowanie klikalnego linku do oficjalnej nawigacji Google Maps.
- **Bezpieczeństwo**:
  - Bezpieczne przechowywanie klucza Google Maps API przy użyciu szyfrowania Windows DPAPI per-user.
- **Kompilacja do EXE**:
  - Kompilacja do samodzielnego pliku wykonywalnego `GoogleMapsRoutes.exe` za pomocą modułu PS2EXE.

---

## Wymagania

- **System operacyjny**: Windows 10 / Windows 11 / Windows Server
- **Środowisko**: PowerShell 5.1 (Desktop) lub PowerShell 7+ (Core)
- **Moduły**: `ImportExcel` (do obsługi plików .xlsx), opcjonalnie `ps2exe` (do kompilacji EXE)
- **Klucz API**: Google Maps API z włączonymi usługami:
  - Routes API (v2)
  - Geocoding API
  - Maps Static API

---

## Szybki Start (GUI)

Uruchomienie aplikacji okienkowej:

```powershell
.\GoogleMapsRoutes-GUI.ps1
```

### Tryb Ręczny (Manual)
1. Wprowadź punkt początkowy (np. `Warszawa, Plac Defilad 1`).
2. Wprowadź punkt docelowy (np. `Kraków, Rynek Główny 1`).
3. Opcjonalnie dodaj punkty pośrednie (np. `Radom`, `Kielce`) i ustal ich kolejność przyciskami ▲/▼.
4. Wybierz typ trasy: **Najszybsza**, **Najkrótsza** lub **Eco**.
5. Kliknij **🚀 Oblicz trasę i pobierz mapę**.
6. Wyświetli się łączny dystans, czas, podgląd mapy oraz przyciski do otwarcia trasy w Google Maps.

### Tryb Wsadowy (Batch)
1. Przejdź do zakładki **📁 Z Pliku Danych (Batch)**.
2. Kliknij **Wybierz plik...** i wskaż plik `.xlsx`, `.csv` lub `.json` (przykłady znajdują się w katalogu `.\Samples`).
3. Sprawdź podgląd wczytanych rekordów w tabeli.
4. Kliknij **▶ Rozpocznij przetwarzanie**.
5. Po zakończeniu kliknij **Otwórz folder wyników** lub wyeksportuj raport do Excel / CSV / JSON.

---

## Użycie z Wiersza Poleceń (CLI)

Głównym skryptem CLI jest `Invoke-GoogleMapsRoute.ps1`.

### 1. Trasa ręczna z punktami pośrednimi

```powershell
.\Invoke-GoogleMapsRoute.ps1 `
    -StartPoint "Warszawa, Plac Defilad 1" `
    -EndPoint "Kraków, Rynek Główny 1" `
    -Waypoints "Radom, Plac Konstytucji 1", "Kielce, Sienkiewicza 1" `
    -RouteType Fastest `
    -GenerateMap `
    -OpenBrowser
```

### 2. Trasa najkrótsza (Shortest)

```powershell
.\Invoke-GoogleMapsRoute.ps1 `
    -StartPoint "Gdańsk, Długa 1" `
    -EndPoint "Toruń, Rynek Staromiejski 1" `
    -RouteType Shortest `
    -GenerateMap
```

### 3. Trasa ekologiczna (Eco) z wybranym silnikiem

```powershell
.\Invoke-GoogleMapsRoute.ps1 `
    -StartPoint "Poznań, Stary Rynek 1" `
    -EndPoint "Wrocław, Rynek 1" `
    -RouteType Eco `
    -EmissionType HYBRID `
    -GenerateMap
```

### 4. Przetwarzanie wsadowe pliku (JSON, CSV, Excel)

```powershell
# Przetworzenie pliku Excel i wygenerowanie raportu zbiorczego:
.\Invoke-GoogleMapsRoute.ps1 `
    -InputFile ".\Samples\routes_sample.xlsx" `
    -ExportFormat Excel `
    -OutputFolder ".\Results"

# Przetworzenie pliku CSV:
.\Invoke-GoogleMapsRoute.ps1 `
    -InputFile ".\Samples\routes_sample.csv" `
    -RouteType Fastest `
    -ExportFormat All
```

---

## Schematy Plików Wejściowych

### 1. Format JSON

#### Tryb A: Lista tras
```json
[
  {
    "Nazwa": "Trasa Warszawa - Kraków",
    "Start": "Warszawa, Plac Defilad 1",
    "Koniec": "Kraków, Rynek Główny 1",
    "PunktyPosrednie": [
      "Radom, Plac Konstytucji 3 Maja 1",
      "Kielce, Sienkiewicza 1"
    ],
    "TypTrasy": "Fastest"
  }
]
```

#### Tryb B: Sekwencja przystanków jednej podróży
```json
[
  { "LP": 1, "Lokalizacja": "Warszawa, Marszałkowska 1" },
  { "LP": 2, "Lokalizacja": "Radom, Żeromskiego 1" },
  { "LP": 3, "Lokalizacja": "Kraków, Rynek Główny 1" }
]
```

### 2. Format CSV
Nagłówki są rozpoznawane automatycznie w języku polskim i angielskim:
```csv
Nazwa;Start;Koniec;PunktyPosrednie;TypTrasy
Trasa Katowice - Zakopane;Katowice, Rynek 1;Zakopane, Krupówki 1;Kraków, Rynek Główny 1;Fastest
Trasa Wrocław - Opole;Wrocław, Rynek 1;Opole, Rynek 1;;Shortest
```

### 3. Format Excel (.xlsx)
Kolumny:
- `Nazwa` (lub `Umowa`, `ID`, `Opis`)
- `Start` (lub `Adres A`, `Początek`, `Od`)
- `Koniec` (lub `Adres B`, `Cel`, `Do`)
- `PunktyPosrednie` (opcjonalnie, oddzielone znakiem `|` lub `;`)
- `TypTrasy` (opcjonalnie: `Fastest`, `Shortest`, `Eco`)

---

## Kompilacja do Pliku EXE

Aby skompilować aplikację do samodzielnego pliku wykonywalnego `GoogleMapsRoutes.exe`:

```powershell
.\Build-Exe.ps1 -Target GoogleMapsRoutes
```

Skompilowany program nie wymaga uruchamiania konsoli PowerShell i może być dystrybuowany jako samodzielne narzędzie.

---

## Struktura Projektu

- `GoogleMapsRoutes-GUI.ps1` — Główna aplikacja okienkowa WPF Dark Mode.
- `Invoke-GoogleMapsRoute.ps1` — Skrypt wiersza poleceń (CLI).
- `RouteMapFunctions.ps1` — Moduł silnika (geokodowanie, Routes API, Static Maps, import/eksport).
- `Build-Exe.ps1` — Skrypt kompilatora PS2EXE.
- `Samples/` — Przykładowe pliki testowe (`routes_sample.json`, `routes_sample.csv`, `routes_sample.xlsx`, `multipoint_trip_sample.json`).
- `Process-SchoolTransportRoutes*` — Dedykowane skrypty dla umów na dowozy szkolne (zachowane dla kompatybilności wstecznej).
