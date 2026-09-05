#Requires -Version 5.1
<#
.SYNOPSIS
    Google Maps Routes & Map Generator — Interactive WebView2 Map Subsystem.
.DESCRIPTION
    Initializes and manages the Microsoft Edge WebView2 Chromium control for WPF,
    generating dynamic vector route visualizations with Leaflet.js, markers, popups,
    and adaptive Dark/Light tile themes with seamless fallback to static PNG images.
.NOTES
    Encoding: UTF-8 with BOM
#>

$script:HasWebView2 = $false
$script:WebView2Control = $null
$script:CoreWebView2Env = $null
$script:FallbackWebBrowser = $null

function Initialize-WebView2Environment {
    [CmdletBinding()]
    param()

    if ($script:HasWebView2 -and $script:WebView2Control -and $script:CoreWebView2Env) { return $true }

    $baseDir = if (-not [string]::IsNullOrWhiteSpace($script:AppDir)) {
        $script:AppDir
    } elseif (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        $PSScriptRoot
    } else {
        [System.IO.Path]::GetDirectoryName([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
    }

    $parentDir = if (-not [string]::IsNullOrWhiteSpace($baseDir)) { Split-Path -Parent $baseDir } else { $null }

    # Candidates for Microsoft.Web.WebView2 assemblies
    $candidates = [System.Collections.Generic.List[string]]::new()
    if (-not [string]::IsNullOrWhiteSpace($baseDir)) {
        $candidates.Add((Join-Path $baseDir 'lib\Microsoft.Web.WebView2.Wpf.dll'))
    }
    if (-not [string]::IsNullOrWhiteSpace($parentDir)) {
        $candidates.Add((Join-Path $parentDir 'lib\Microsoft.Web.WebView2.Wpf.dll'))
    }
    $systemPaths = @(
        # Standard system-installed products embedding WebView2
        "C:\Program Files\Fortinet\FortiClient\Microsoft.Web.WebView2.Wpf.dll",
        "C:\Program Files\PowerToys\Microsoft.Web.WebView2.Wpf.dll",
        "C:\Program Files\Surfshark\Microsoft.Web.WebView2.Wpf.dll",
        "C:\Program Files\Microsoft SQL Server Management Studio 21\Release\Common7\IDE\PrivateAssemblies\Microsoft.Web.WebView2.Wpf.dll",
        "C:\Program Files\Microsoft Office\root\Office16\WritingAssistant\Microsoft.Web.WebView2.Wpf.dll"
    )
    foreach ($sp in $systemPaths) {
        $candidates.Add([string]$sp)
    }

    $loaded = $false
    foreach ($cand in $candidates) {
        if ($cand -and (Test-Path $cand)) {
            $dir = Split-Path -Parent $cand
            $core = Join-Path $dir 'Microsoft.Web.WebView2.Core.dll'
            if (Test-Path $core) {
                try {
                    Add-Type -Path $core -ErrorAction Stop
                    Add-Type -Path $cand -ErrorAction Stop
                    # Verify instantiability with current CLR / runtime
                    $testControl = [Microsoft.Web.WebView2.Wpf.WebView2]::new()
                    $loaded = $true
                    Write-AppLog "Loaded WebView2 WPF assemblies from: $dir" "OK"
                    break
                }
                catch {
                    # Continue checking other candidates if incompatible with current runtime
                }
            }
        }
    }

    if (-not $loaded) {
        Write-AppLog "WebView2 WPF assembly not found or incompatible. Static PNG mode / fallback will be active." "INFO"
        $script:HasWebView2 = $false
        return $false
    }

    # Crucial: Configure writable UserDataFolder in LocalAppData to avoid E_ACCESSDENIED (0x80070005)
    # when running from powershell.exe in C:\Windows\System32
    $userDataFolder = Join-Path $env:LOCALAPPDATA "GoogleMapsRoutes\WebView2Data"
    if (-not (Test-Path $userDataFolder)) {
        try { [System.IO.Directory]::CreateDirectory($userDataFolder) | Out-Null } catch { }
    }

    try {
        $envTask = [Microsoft.Web.WebView2.Core.CoreWebView2Environment]::CreateAsync($null, $userDataFolder, $null)
        $script:CoreWebView2Env = $envTask.GetAwaiter().GetResult()
    }
    catch {
        Write-AppLog "CoreWebView2Environment creation failed: $($_.Exception.Message)" "WARN"
        $script:CoreWebView2Env = $null
    }

    try {
        $wv = [Microsoft.Web.WebView2.Wpf.WebView2]::new()
        $wv.Add_Loaded({
            if ($script:CoreWebView2Env -and (-not $script:WebView2Initialized)) {
                try {
                    $script:WebView2Control.EnsureCoreWebView2Async($script:CoreWebView2Env) | Out-Null
                    $script:WebView2Initialized = $true
                } catch { }
            }
        })
        $wv.Add_NavigationCompleted({
            param($s, $e)
            if ($e.IsSuccess) {
                Write-AppLog "Interactive Leaflet map rendered successfully." "INFO"
            } else {
                Write-AppLog "Interactive map navigation status: $($e.WebErrorStatus)" "WARN"
            }
        })
        $script:WebView2Control = $wv
        $script:HasWebView2 = $true
        return $true
    }
    catch {
        Write-AppLog "WebView2 instantiation failed: $($_.Exception.Message)" "WARN"
        $script:HasWebView2 = $false
        return $false
    }
}

function ConvertFrom-GoogleEncodedPolyline {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$EncodedPolyline)

    $points = [System.Collections.Generic.List[PSCustomObject]]::new()
    if ([string]::IsNullOrWhiteSpace($EncodedPolyline)) { return $points }

    $len = $EncodedPolyline.Length
    $index = 0
    $lat = 0
    $lng = 0

    while ($index -lt $len) {
        $b = 0
        $shift = 0
        $result = 0
        do {
            $b = [int][char]$EncodedPolyline[$index++] - 63
            $result = $result -bor (($b -band 0x1f) -shl $shift)
            $shift += 5
        } while ($b -ge 0x20 -and $index -lt $len)
        $dlat = if (($result -band 1) -ne 0) { -bnot ($result -shr 1) } else { ($result -shr 1) }
        $lat += $dlat

        $shift = 0
        $result = 0
        do {
            $b = [int][char]$EncodedPolyline[$index++] - 63
            $result = $result -bor (($b -band 0x1f) -shl $shift)
            $shift += 5
        } while ($b -ge 0x20 -and $index -lt $len)
        $dlng = if (($result -band 1) -ne 0) { -bnot ($result -shr 1) } else { ($result -shr 1) }
        $lng += $dlng

        $points.Add([PSCustomObject]@{
            Latitude  = [math]::Round($lat * 1e-5, 6)
            Longitude = [math]::Round($lng * 1e-5, 6)
        })
    }

    return $points
}

function New-RouteHtmlMap {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][string]$RouteName = 'Route',
        [Parameter(Mandatory = $true)][string]$EncodedPolyline,
        [Parameter()][double]$OriginLat = 0,
        [Parameter()][double]$OriginLng = 0,
        [Parameter()][string]$OriginAddress = '',
        [Parameter()][double]$DestLat = 0,
        [Parameter()][double]$DestLng = 0,
        [Parameter()][string]$DestAddress = '',
        [Parameter()][object[]]$Waypoints = @(),
        [Parameter()][double]$DistanceKm = 0,
        [Parameter()][int]$DurationMin = 0,
        [Parameter()][string]$RouteType = 'Fastest',
        [Parameter()][bool]$IsDarkMode = $true,
        [Parameter()][string]$OutputPath = '',
        [Parameter()][string]$CartoApiKey = ''
    )

    if ([string]::IsNullOrWhiteSpace($RouteName)) {
        $RouteName = 'Route'
    }
    if (-not $OutputPath) {
        $ts = (Get-Date).ToString('yyyyMMdd_HHmmssfff')
        $OutputPath = Join-Path $env:TEMP "gmaps_interactive_route_${ts}.html"
    }

    # Clean up older temporary route map HTML files (> 30 mins)
    try {
        Get-ChildItem -Path $env:TEMP -Filter 'gmaps_interactive_route_*.html' -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -lt (Get-Date).AddMinutes(-30) } |
            Remove-Item -Force -ErrorAction SilentlyContinue
    } catch { }

    $resolvedCartoKey = if (-not [string]::IsNullOrWhiteSpace($CartoApiKey)) {
        $CartoApiKey.Trim()
    } elseif ($script:AppConfig -and -not [string]::IsNullOrWhiteSpace($script:AppConfig.CartoApiKey)) {
        $script:AppConfig.CartoApiKey.Trim()
    } elseif (-not [string]::IsNullOrWhiteSpace($env:CARTO_API_KEY)) {
        $env:CARTO_API_KEY.Trim()
    } else {
        ''
    }

    $points = if ($EncodedPolyline) { ConvertFrom-GoogleEncodedPolyline -EncodedPolyline $EncodedPolyline } else { @() }
    
    # Coordinates array for Leaflet: [[lat, lng], [lat, lng], ...]
    $coordJsonList = [System.Collections.Generic.List[string]]::new()
    foreach ($pt in $points) {
        $lat = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:F6}", [double]$pt.Latitude)
        $lng = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:F6}", [double]$pt.Longitude)
        $coordJsonList.Add("[$lat, $lng]")
    }
    $coordJson = "[$($coordJsonList -join ',')]"

    # Tile layer URL and attribution based on theme
    $tileUrl = if ($IsDarkMode) {
        'https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png'
    } else {
        'https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png'
    }

    if (-not [string]::IsNullOrWhiteSpace($resolvedCartoKey)) {
        $escapedCarto = [System.Uri]::EscapeDataString($resolvedCartoKey)
        $tileUrl += "?key=$escapedCarto"
    }
    $bgStyle = if ($IsDarkMode) { 'background:#0f172a;color:#f8fafc;' } else { 'background:#f8fafc;color:#0f172a;' }
    $polyColor = if ($IsDarkMode) { '#38bdf8' } else { '#2563eb' }

    $safeRouteName = [System.Net.WebUtility]::HtmlEncode($RouteName)
    $safeOrigin = [System.Net.WebUtility]::HtmlEncode($OriginAddress)
    $safeDest = [System.Net.WebUtility]::HtmlEncode($DestAddress)

    # Waypoints JSON
    $wpJsonList = [System.Collections.Generic.List[string]]::new()
    if ($Waypoints -and @($Waypoints).Count -gt 0) {
        $idx = 1
        foreach ($wp in $Waypoints) {
            if ($wp.Latitude -and $wp.Longitude) {
                $lat = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:F6}", [double]$wp.Latitude)
                $lng = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:F6}", [double]$wp.Longitude)
                $addr = [System.Net.WebUtility]::HtmlEncode($(if ($wp.Address) { [string]$wp.Address } else { "Stop $idx" }))
                $wpJsonList.Add("{ lat: $lat, lng: $lng, label: '$idx', address: '$addr' }")
                $idx++
            }
        }
    }
    $wpJson = "[$($wpJsonList -join ',')]"

    $origLatStr = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:F6}", $OriginLat)
    $origLngStr = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:F6}", $OriginLng)
    $destLatStr = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:F6}", $DestLat)
    $destLngStr = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:F6}", $DestLng)

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>$safeRouteName</title>
  <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"/>
  <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
  <style>
    html, body, #map { width: 100%; height: 100%; margin: 0; padding: 0; overflow: hidden; $bgStyle font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; }
    .route-info-card {
      position: absolute;
      top: 12px;
      right: 12px;
      z-index: 1000;
      background: rgba(15, 23, 42, 0.88);
      color: #f8fafc;
      padding: 10px 14px;
      border-radius: 8px;
      backdrop-filter: blur(8px);
      box-shadow: 0 4px 16px rgba(0,0,0,0.3);
      font-size: 13px;
      line-height: 1.4;
      border: 1px solid rgba(255,255,255,0.1);
      max-width: 300px;
      pointer-events: none;
    }
    .route-info-card strong { color: #38bdf8; }
    .marker-pin {
      width: 30px;
      height: 30px;
      border-radius: 50% 50% 50% 0;
      background: #0284c7;
      position: absolute;
      transform: rotate(-45deg);
      left: 50%;
      top: 50%;
      margin: -15px 0 0 -15px;
      box-shadow: 0 2px 5px rgba(0,0,0,0.4);
    }
    .marker-pin::after {
      content: '';
      width: 14px;
      height: 14px;
      margin: 8px 0 0 8px;
      background: #fff;
      position: absolute;
      border-radius: 50%;
    }
    .marker-pin span {
      position: absolute;
      transform: rotate(45deg);
      top: 5px;
      left: 7px;
      font-size: 11px;
      font-weight: bold;
      color: #0f172a;
      z-index: 10;
    }
    .pin-start { background: #10b981; }
    .pin-dest { background: #ef4444; }
    .pin-wp { background: #3b82f6; }
    .leaflet-popup-content-wrapper { background: #1e293b; color: #f8fafc; border: 1px solid #334155; border-radius: 8px; }
    .leaflet-popup-tip { background: #1e293b; }
  </style>
</head>
<body>
  <div id="map"></div>
  <div class="route-info-card">
    <div><strong>$safeRouteName</strong></div>
    <div>Distance: $DistanceKm km | Time: $DurationMin min</div>
    <div style="font-size:11px;opacity:0.8;">Type: $RouteType</div>
  </div>

  <script>
    const coords = $coordJson;
    const waypoints = $wpJson;

    const map = L.map('map', { zoomControl: true });

    L.tileLayer('$tileUrl', {
      maxZoom: 19,
      attribution: '&copy; <a href="https://carto.com/">CARTO</a> &copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a>'
    }).addTo(map);

    function createIcon(label, cls) {
      return L.divIcon({
        className: 'custom-pin',
        html: '<div class="marker-pin ' + cls + '"><span>' + label + '</span></div>',
        iconSize: [30, 42],
        iconAnchor: [15, 38],
        popupAnchor: [0, -34]
      });
    }

    window.addEventListener('resize', () => {
      map.invalidateSize();
    });

    setTimeout(() => {
      map.invalidateSize();
      if (coords && coords.length > 0 && typeof poly !== 'undefined') {
        map.fitBounds(poly.getBounds(), { padding: [40, 40] });
      }
    }, 200);

    let poly = null;
    if (coords && coords.length > 0) {
      poly = L.polyline(coords, {
        color: '$polyColor',
        weight: 5,
        opacity: 0.9,
        lineJoin: 'round'
      }).addTo(map);

      map.fitBounds(poly.getBounds(), { padding: [40, 40] });
    } else {
      map.setView([$origLatStr || 52.0, $origLngStr || 19.5], 10);
    }

    // Origin Marker
    if ($origLatStr && $origLngStr) {
      L.marker([$origLatStr, $origLngStr], { icon: createIcon('A', 'pin-start') })
        .addTo(map)
        .bindPopup('<strong>Origin (Start)</strong><br>$safeOrigin');
    }

    // Waypoints Markers
    waypoints.forEach(wp => {
      L.marker([wp.lat, wp.lng], { icon: createIcon(wp.label, 'pin-wp') })
        .addTo(map)
        .bindPopup('<strong>Waypoint ' + wp.label + '</strong><br>' + wp.address);
    });

    // Destination Marker
    if ($destLatStr && $destLngStr) {
      L.marker([$destLatStr, $destLngStr], { icon: createIcon('B', 'pin-dest') })
        .addTo(map)
        .bindPopup('<strong>Destination (End)</strong><br>$safeDest');
    }
  </script>
</body>
</html>
"@

    [System.IO.File]::WriteAllText($OutputPath, $html, [System.Text.UTF8Encoding]::new($true))
    return $OutputPath
}

function Navigate-RouteInteractiveMap {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$HtmlPath,
        [Parameter()][object]$HostPanel = $null
    )

    if (Initialize-WebView2Environment) {
        if ($HostPanel -and $script:WebView2Control) {
            if (-not $HostPanel.Child -or $HostPanel.Child -ne $script:WebView2Control) {
                $HostPanel.Child = $script:WebView2Control
            }
        }
        if ($script:WebView2Control) {
            try {
                if ($script:CoreWebView2Env -and (-not $script:WebView2Initialized)) {
                    try {
                        $script:WebView2Control.EnsureCoreWebView2Async($script:CoreWebView2Env) | Out-Null
                        $script:WebView2Initialized = $true
                    } catch {
                        # Outside active event loop; Loaded event handler will initialize CoreWebView2
                    }
                }

                $targetUri = [System.Uri]::new($HtmlPath)
                if ($script:WebView2Control.CoreWebView2) {
                    $script:WebView2Control.CoreWebView2.Navigate($targetUri.AbsoluteUri)
                }
                $script:WebView2Control.Source = $targetUri
                return $true
            }
            catch {
                Write-AppLog "WebView2 navigation error: $($_.Exception.Message)" "WARN"
            }
        }
    }

    # Fallback to WPF WebBrowser if WebView2 is unavailable or failed
    if ($HostPanel) {
        if (-not $script:FallbackWebBrowser) {
            $script:FallbackWebBrowser = [System.Windows.Controls.WebBrowser]::new()
        }
        if (-not $HostPanel.Child -or $HostPanel.Child -ne $script:FallbackWebBrowser) {
            $HostPanel.Child = $script:FallbackWebBrowser
        }
        try {
            $script:FallbackWebBrowser.Navigate([System.Uri]::new($HtmlPath))
            return $true
        }
        catch {
            Write-AppLog "WebBrowser fallback navigation failed: $($_.Exception.Message)" "WARN"
        }
    }

    return $false
}

function Update-InteractiveRouteMap {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object]$CalcResult = $null
    )

    $calc = if ($CalcResult) { $CalcResult } else { $script:LastManualResult }
    if (-not $calc -or -not $calc.EncodedPolyline) { return }

    $isDark = ($script:CurrentTheme -ne 'Light')
    $resolvedRouteName = if (-not [string]::IsNullOrWhiteSpace($script:ActiveManualRouteName)) {
        $script:ActiveManualRouteName
    } elseif ($script:Controls -and $script:Controls.txtManualName -and -not [string]::IsNullOrWhiteSpace($script:Controls.txtManualName.Text)) {
        $script:Controls.txtManualName.Text.Trim()
    } else {
        'Route'
    }
    if ([string]::IsNullOrWhiteSpace($resolvedRouteName)) { $resolvedRouteName = 'Route' }

    $cartoKey = if (Get-Command Get-CurrentCartoApiKey -ErrorAction SilentlyContinue) {
        Get-CurrentCartoApiKey
    } elseif ($script:AppConfig -and -not [string]::IsNullOrWhiteSpace($script:AppConfig.CartoApiKey)) {
        $script:AppConfig.CartoApiKey.Trim()
    } else {
        ''
    }

    $htmlMap = New-RouteHtmlMap -RouteName $resolvedRouteName `
        -EncodedPolyline $calc.EncodedPolyline `
        -OriginLat $calc.OriginLat -OriginLng $calc.OriginLng -OriginAddress $calc.OriginAddress `
        -DestLat $calc.DestLat -DestLng $calc.DestLng -DestAddress $calc.DestAddress `
        -Waypoints $calc.Waypoints -DistanceKm $calc.DistanceKm -DurationMin $calc.DurationMin `
        -RouteType $calc.RouteType -IsDarkMode $isDark -CartoApiKey $cartoKey

    $script:LastInteractiveMapPath = $htmlMap
    $mapHostPanel = if ($script:Controls) { $script:Controls.pnlInteractiveMapHost } else { $null }

    Navigate-RouteInteractiveMap -HtmlPath $htmlMap -HostPanel $mapHostPanel | Out-Null
}

Set-Item -Path "function:global:Initialize-WebView2Environment" -Value (Get-Item "function:Initialize-WebView2Environment").ScriptBlock -ErrorAction SilentlyContinue
Set-Item -Path "function:global:New-RouteHtmlMap" -Value (Get-Item "function:New-RouteHtmlMap").ScriptBlock -ErrorAction SilentlyContinue
Set-Item -Path "function:global:Navigate-RouteInteractiveMap" -Value (Get-Item "function:Navigate-RouteInteractiveMap").ScriptBlock -ErrorAction SilentlyContinue
Set-Item -Path "function:global:Update-InteractiveRouteMap" -Value (Get-Item "function:Update-InteractiveRouteMap").ScriptBlock -ErrorAction SilentlyContinue
