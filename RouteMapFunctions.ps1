#Requires -Version 5.1
<#
.SYNOPSIS
    Wspólne funkcje do geokodowania adresów, obliczania tras i generowania map PNG.

.DESCRIPTION
    Moduł funkcji wykorzystywany przez Get-CarRoute_WithMap.ps1
    oraz Process-SchoolTransportRoutes.ps1.
    Zawiera:
      - Select-InputExcel       — dialog wyboru pliku Excel
      - Get-AddressCoordinates  — geokodowanie adresu przez Google Geocoding API
      - Get-CarRouteData        — obliczanie trasy przez Google Routes API v2
      - Save-RouteMapPng        — generowanie mapy PNG przez Google Static Maps API

.NOTES
    Wymagana zmienna środowiskowa: GOOGLE_MAPS_API_KEY lub przekazanie klucza API jako parametru.
    Encoding: UTF-8 with BOM
#>

function Select-InputExcel {
    Add-Type -AssemblyName System.Windows.Forms
    $Dialog = [System.Windows.Forms.OpenFileDialog]::new()
    $Dialog.Title = 'Wybierz plik Excel z adresami'
    $Dialog.Filter = 'Pliki Excel (*.xlsx;*.xls)|*.xlsx;*.xls|Wszystkie pliki (*.*)|*.*'
    $Dialog.InitialDirectory = [Environment]::GetFolderPath('MyDocuments')
    $Result = $Dialog.ShowDialog()
    if ($Result -eq [System.Windows.Forms.DialogResult]::OK) { return $Dialog.FileName }
    return $null
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
            if (-not [string]::IsNullOrWhiteSpace($Value)) {
                return $Value
            }
        }
    }

    return $null
}

function Get-AddressCoordinates {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Address,
        [Parameter(Mandatory)][string]$ApiKey,
        [Parameter()][switch]$RequireStreetNumber
    )
    if ([string]::IsNullOrWhiteSpace($Address)) { return $null }
    $EncodedAddress = [System.Uri]::EscapeDataString($Address.Trim())
    $Url = "https://maps.googleapis.com/maps/api/geocode/json?address=$EncodedAddress&language=pl&key=$ApiKey"
    try {
        Write-Verbose "Geokodowanie: '$Address'"
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
            $ResultTypes = @($ResultItem.types)
            $IsPartialMatch = $ResultItem.PSObject.Properties.Name -contains 'partial_match' -and $ResultItem.partial_match -eq $true
            $MatchStatus = if ($RequireStreetNumber -and [string]::IsNullOrWhiteSpace($StreetNumber)) {
                'IMPRECISE_MATCH'
            }
            else {
                'OK'
            }

            return [PSCustomObject]@{
                Latitude         = [double]$Location.lat
                Longitude        = [double]$Location.lng
                FormattedAddress = $FormattedAddress
                UlicaINumer      = $StreetWithNumber
                KodPocztowy      = $PostalCode
                Miasto           = $City
                MatchType        = $ResultTypes -join ','
                PartialMatch     = $IsPartialMatch
                Status           = $MatchStatus
            }
        }
        else {
            Write-Warning "Geokodowanie nieudane dla '$Address'. Status API: $($Response.status)"
            return [PSCustomObject]@{
                Latitude = $null; Longitude = $null; FormattedAddress = $null
                UlicaINumer = $null; KodPocztowy = $null; Miasto = $null
                MatchType = $null; PartialMatch = $null; Status = $Response.status
            }
        }
    }
    catch {
        $Message = $_.Exception.Message
        Write-Warning "Błąd geokodowania '$Address': $Message"
        return [PSCustomObject]@{
            Latitude = $null; Longitude = $null; FormattedAddress = $null
            UlicaINumer = $null; KodPocztowy = $null; Miasto = $null
            MatchType = $null; PartialMatch = $null; Status = "EXCEPTION: $Message"
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
        [Parameter()][object[]]$IntermediatePoints = @(),
        [Parameter(Mandatory)][string]$ApiKey
    )
    $RoutesUrl = 'https://routes.googleapis.com/directions/v2:computeRoutes'
    $RequestBody = @{
        origin                   = @{ location = @{ latLng = @{ latitude = $OriginLat; longitude = $OriginLng } } }
        destination              = @{ location = @{ latLng = @{ latitude = $DestLat; longitude = $DestLng } } }
        travelMode               = 'DRIVE'
        routingPreference        = 'TRAFFIC_UNAWARE'
        computeAlternativeRoutes = @($IntermediatePoints).Count -eq 0
        languageCode             = 'pl'
        units                    = 'METRIC'
    }
    if (@($IntermediatePoints).Count -gt 0) {
        $RequestBody.intermediates = @($IntermediatePoints | ForEach-Object {
                @{
                    location = @{
                        latLng = @{
                            latitude  = [double]$_.Latitude
                            longitude = [double]$_.Longitude
                        }
                    }
                }
            })
    }
    $Headers = @{
        'X-Goog-Api-Key'   = $ApiKey
        'Content-Type'     = 'application/json'
        'X-Goog-FieldMask' = 'routes.duration,routes.distanceMeters,routes.polyline.encodedPolyline'
    }
    try {
        Write-Verbose "Routes API: ($OriginLat,$OriginLng) -> ($DestLat,$DestLng)"
        $Response = Invoke-RestMethod -Uri $RoutesUrl -Method POST -Headers $Headers `
            -Body ($RequestBody | ConvertTo-Json -Depth 10) -TimeoutSec 60
        $Routes = @($Response.routes)
        if ($Routes.Count -eq 0) {
            Write-Warning "Routes API nie zwrocilo tras."
            return [PSCustomObject]@{ OdlegloscKm = $null; CzasMin = $null; EncodedPolyline = $null; Status = 'NO_ROUTES' }
        }
        # Google optymalizuje trasy pod czas, nie dystans - wybieramy najkrotsza w km sposrod zwroconych alternatyw
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
        Write-Error "Blad Routes API: $($_.Exception.Message)"
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
        [Parameter()][object[]]$RoutePoints = @(),
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)][string]$ApiKey,
        [Parameter()][int]$Width = 900,
        [Parameter()][int]$Height = 600,
        # Parametry tekstu nakładki (opcjonalne)
        [Parameter()][string]$TekstAdresA = '',
        [Parameter()][string]$TekstAdresB = '',
        [Parameter()][string]$TekstOdleglosc = ''
    )
    $EncodedForUrl = [System.Uri]::EscapeDataString($EncodedPolyline)
    $MarkerParameters = if (@($RoutePoints).Count -gt 0) {
        $PointCount = @($RoutePoints).Count
        for ($PointIndex = 0; $PointIndex -lt $PointCount; $PointIndex++) {
            $Point = $RoutePoints[$PointIndex]
            $MarkerColor = if ($PointIndex -eq 0) { 'green' } elseif ($PointIndex -eq ($PointCount - 1)) { 'red' } else { 'blue' }
            $MarkerNumber = $PointIndex + 1
            $MarkerLabel = if ($MarkerNumber -le 9) { [string]$MarkerNumber } else { [char](55 + $MarkerNumber) }
            $MarkerValue = [System.Uri]::EscapeDataString("color:$MarkerColor|label:$MarkerLabel|$($Point.Latitude),$($Point.Longitude)")
            "&markers=$MarkerValue"
        }
    }
    else {
        $MarkerStart = [System.Uri]::EscapeDataString("color:green|label:A|$OriginLat,$OriginLng")
        $MarkerEnd = [System.Uri]::EscapeDataString("color:red|label:B|$DestLat,$DestLng")
        @("&markers=$MarkerStart", "&markers=$MarkerEnd")
    }
    $StaticMapUrl = ("https://maps.googleapis.com/maps/api/staticmap" +
        "?size=${Width}x${Height}" +
        "&path=weight:4|color:0x0066FFff|enc:$EncodedForUrl" +
        ($MarkerParameters -join '') +
        "&key=$ApiKey")
    try {
        Write-Verbose "Pobieranie mapy PNG: $OutputPath"
        Invoke-WebRequest -Uri $StaticMapUrl -OutFile $OutputPath -TimeoutSec 30 | Out-Null
        Write-Verbose "Mapa pobrana: $OutputPath"

        # Nakladka tekstowa przez System.Drawing
        $MaTextOverlay = (-not [string]::IsNullOrWhiteSpace($TekstAdresA)) -or
        (-not [string]::IsNullOrWhiteSpace($TekstAdresB)) -or
        (-not [string]::IsNullOrWhiteSpace($TekstOdleglosc))

        if ($MaTextOverlay) {
            try {
                Add-Type -AssemblyName System.Drawing

                # WAŻNE: ładujemy przez MemoryStream — Bitmap otwarty ze ścieżki
                # blokuje plik (GDI+ file lock), uniemożliwiając zapis pod tę samą nazwę.
                $FileBytes = [System.IO.File]::ReadAllBytes($OutputPath)
                $MemStream = [System.IO.MemoryStream]::new($FileBytes)
                $BitmapSrc = [System.Drawing.Bitmap]::new($MemStream)

                # Jeśli format jest indeksowany (np. 8-bit palette), konwertuj do 32bpp
                # bo Graphics.FromImage nie obsługuje formatów indeksowanych
                if ($BitmapSrc.PixelFormat -band [System.Drawing.Imaging.PixelFormat]::Indexed) {
                    $Bitmap = [System.Drawing.Bitmap]::new($BitmapSrc.Width, $BitmapSrc.Height, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
                    $GfxConv = [System.Drawing.Graphics]::FromImage($Bitmap)
                    $GfxConv.DrawImage($BitmapSrc, 0, 0, $BitmapSrc.Width, $BitmapSrc.Height)
                    $GfxConv.Dispose()
                    $BitmapSrc.Dispose()
                }
                else {
                    $Bitmap = $BitmapSrc
                }

                $Graphics = [System.Drawing.Graphics]::FromImage($Bitmap)
                $Graphics.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
                $Graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit

                # ── Czcionki ──────────────────────────────────────────────────
                $FontLabel = [System.Drawing.Font]::new('Segoe UI', 8,  [System.Drawing.FontStyle]::Bold)
                $FontValue = [System.Drawing.Font]::new('Segoe UI', 9,  [System.Drawing.FontStyle]::Regular)
                $FontDist  = [System.Drawing.Font]::new('Segoe UI', 13, [System.Drawing.FontStyle]::Bold)
                $FontDate  = [System.Drawing.Font]::new('Segoe UI', 7,  [System.Drawing.FontStyle]::Italic)

                # ── Kolory ────────────────────────────────────────────────────
                $BrushWhite   = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::White)
                $BrushYellow  = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 255, 220, 0))
                $BrushGray    = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 180, 220, 255))
                $BrushDateClr = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(200, 180, 190, 210))

                # ── Layout: strefa prawa zarezerwowana dla odległości ─────────
                $Pad         = 10
                $LineH       = 18    # wysokość wiersza tekstu adresu
                $LabelH      = 18    # wysokość etykiety (A/B)
                $DateH       = 14    # wysokość wiersza daty
                $PadTop      = 6
                $PadBot      = 5
                $RightW      = 110   # szerokość kolumny prawej (odległość + data)

                # Oblicz dostępną szerokość dla tekstu adresu
                $LabelASize = $Graphics.MeasureString('A: ', $FontLabel)
                $LabelBSize = $Graphics.MeasureString('B: ', $FontLabel)
                $LabelMaxW  = [float][math]::Max($LabelASize.Width, $LabelBSize.Width)
                $TextMaxW   = [float]($Bitmap.Width - $RightW - $LabelMaxW - $Pad * 2)

                # ── Word-wrap inline (bez scriptblock — PS5 safe) ────────────
                # Zamienia dlugi adres na tablice 1 lub 2 linii
                function Get-WrappedLines {
                    param(
                        [System.Drawing.Graphics]$G,
                        [string]$Text,
                        [System.Drawing.Font]$F,
                        [float]$MaxW
                    )
                    if ([string]::IsNullOrWhiteSpace($Text)) { return [string[]]@('') }
                    if ($G.MeasureString($Text, $F).Width -le $MaxW) { return [string[]]@($Text) }
                    $Words = $Text -split ' '
                    $L1 = ''; $L2 = ''; $On2 = $false
                    foreach ($W in $Words) {
                        if (-not $On2) {
                            $T = if ($L1) { "$L1 $W" } else { $W }
                            if ($G.MeasureString($T, $F).Width -le $MaxW) { $L1 = $T }
                            else { $On2 = $true; $L2 = $W }
                        } else {
                            $T2 = if ($L2) { "$L2 $W" } else { $W }
                            if ($G.MeasureString($T2, $F).Width -le $MaxW) { $L2 = $T2 }
                            else {
                                if ($L2.Length -gt 3) { $L2 = $L2.Substring(0, $L2.Length - 3) + '...' }
                                break
                            }
                        }
                    }
                    if ($L2) { return [string[]]@($L1, $L2) } else { return [string[]]@($L1) }
                }

                $LinesA = @(Get-WrappedLines -G $Graphics -Text $TekstAdresA -F $FontValue -MaxW $TextMaxW)
                $LinesB = @(Get-WrappedLines -G $Graphics -Text $TekstAdresB -F $FontValue -MaxW $TextMaxW)

                # ── Oblicz dynamiczną wysokość paska ──────────────────────────
                $RowsA  = $LinesA.Count
                $RowsB  = $LinesB.Count
                $ContentH   = $PadTop + $LabelH + (($RowsA - 1) * $LineH) + 4 +
                              $LabelH + (($RowsB - 1) * $LineH) + $DateH + $PadBot
                $BannerHeight = [int][math]::Max(88, $ContentH)
                $BannerY      = $Bitmap.Height - $BannerHeight

                # ── Tło paska ─────────────────────────────────────────────────
                $BrushBg = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(210, 12, 17, 32))
                $Graphics.FillRectangle($BrushBg, 0, $BannerY, $Bitmap.Width, $BannerHeight)

                $PenLine = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(200, 0, 100, 255), 2)
                $Graphics.DrawLine($PenLine, 0, $BannerY, $Bitmap.Width, $BannerY)

                # ── Rysuj A ───────────────────────────────────────────────────
                $CurY = [float]($BannerY + $PadTop)
                $Graphics.DrawString('A: ', $FontLabel, $BrushGray, [float]$Pad, $CurY)
                $LabelXOffset = [float]($Pad + $LabelMaxW)
                foreach ($Line in $LinesA) {
                    $Graphics.DrawString($Line, $FontValue, $BrushWhite, $LabelXOffset, $CurY)
                    $CurY += [float]$LineH
                }

                # ── Rysuj B ───────────────────────────────────────────────────
                $CurY += 4.0
                $Graphics.DrawString('B: ', $FontLabel, $BrushGray, [float]$Pad, $CurY)
                foreach ($Line in $LinesB) {
                    $Graphics.DrawString($Line, $FontValue, $BrushWhite, $LabelXOffset, $CurY)
                    $CurY += [float]$LineH
                }

                # ── Odległość (prawa strona, wyśrodkowana pionowo) ────────────
                if (-not [string]::IsNullOrWhiteSpace($TekstOdleglosc)) {
                    $DistSizeF  = $Graphics.MeasureString($TekstOdleglosc, $FontDist)
                    $DistX      = [float]($Bitmap.Width - $DistSizeF.Width - $Pad)
                    $DistY      = [float]($BannerY + ($BannerHeight - $DateH - $PadBot - $DistSizeF.Height) / 2)
                    $Graphics.DrawString($TekstOdleglosc, $FontDist, $BrushYellow, $DistX, $DistY)
                }

                # ── Data generacji (prawa strona, sam dół paska) ──────────────
                $DateStr     = Get-Date -Format 'yyyy-MM-dd  HH:mm'
                $DateSizeF   = $Graphics.MeasureString($DateStr, $FontDate)
                $DateX       = [float]($Bitmap.Width - $DateSizeF.Width - $Pad)
                $DateY       = [float]($Bitmap.Height - $PadBot - $DateSizeF.Height)
                $Graphics.DrawString($DateStr, $FontDate, $BrushDateClr, $DateX, $DateY)

                # ── Zwolnij zasoby GDI+ ───────────────────────────────────────
                $PenLine.Dispose()
                $BrushBg.Dispose(); $BrushWhite.Dispose(); $BrushYellow.Dispose()
                $BrushGray.Dispose(); $BrushDateClr.Dispose()
                $FontLabel.Dispose(); $FontValue.Dispose(); $FontDist.Dispose(); $FontDate.Dispose()
                $Graphics.Dispose()

                # Zapisz z powrotem — plik NIE jest zablokowany (ładowany przez MemoryStream)
                $Bitmap.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
                $Bitmap.Dispose()
                $MemStream.Dispose()

                Write-Verbose "Nakladka tekstowa dodana: $OutputPath"
            }
            catch {
                Write-Host "  [BLAD nakladki PNG] $($_.Exception.Message)" -ForegroundColor Red
                Write-Warning "Nie udalo sie dodac nakladki tekstowej do PNG: $($_.Exception.Message)"
                # PNG bez nakladki zostaje — nie przerywamy
            }
        }

        return $true
    }
    catch {
        Write-Warning "Nie udalo sie pobrac/przetworzyc mapy PNG: $($_.Exception.Message)"
        return $false
    }
}
