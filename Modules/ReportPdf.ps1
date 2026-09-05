#Requires -Version 5.1
<#
.SYNOPSIS
    Google Maps Routes & Map Generator — PDF Report Generation Subsystem.
.DESCRIPTION
    Generates publication-quality PDF reports for single routes and batch summaries
    using native headless Microsoft Edge, embedding base64 route maps and stop itineraries.
.NOTES
    Encoding: UTF-8 with BOM
#>

function Find-EdgeExecutable {
    [CmdletBinding()]
    param()

    $candidates = @(
        "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        "C:\Program Files\Microsoft\Edge\Application\msedge.exe"
    )
    foreach ($cand in $candidates) {
        if (Test-Path $cand) { return $cand }
    }
    $cmd = Get-Command msedge.exe -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    return $null
}

function Export-RoutePdfReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $false)][string]$RouteName = 'Route',
        [Parameter()][double]$DistanceKm = 0,
        [Parameter()][int]$DurationMin = 0,
        [Parameter()][string]$RouteType = 'Fastest',
        [Parameter()][string]$EmissionType = '',
        [Parameter()][bool]$AvoidTolls = $false,
        [Parameter()][bool]$AvoidHighways = $false,
        [Parameter()][bool]$AvoidFerries = $false,
        [Parameter()][string]$OriginAddress = '',
        [Parameter()][double]$OriginLat = 0,
        [Parameter()][double]$OriginLng = 0,
        [Parameter()][string]$DestAddress = '',
        [Parameter()][double]$DestLat = 0,
        [Parameter()][double]$DestLng = 0,
        [Parameter()][object[]]$Waypoints = @(),
        [Parameter()][string]$MapImagePath = '',
        [Parameter()][string]$GoogleMapsUrl = '',
        [Parameter()][string]$Language = 'en',
        [Parameter()][string]$ReportTitle = 'Route Report'
    )

    if ([string]::IsNullOrWhiteSpace($RouteName)) {
        $RouteName = 'Route'
    }

    $edgePath = Find-EdgeExecutable
    if (-not $edgePath) {
        throw "Microsoft Edge executable (msedge.exe) not found. PDF export requires Edge."
    }

    $outDir = Split-Path -Parent $OutputPath
    if ($outDir -and -not (Test-Path $outDir)) {
        New-Item -ItemType Directory -Path $outDir -Force | Out-Null
    }

    # Map image base64 data URI
    $mapImgTag = ''
    if ($MapImagePath -and (Test-Path $MapImagePath)) {
        try {
            $bytes = [System.IO.File]::ReadAllBytes($MapImagePath)
            $b64 = [Convert]::ToBase64String($bytes)
            $mapImgTag = "<div class='map-container'><img src='data:image/png;base64,$b64' alt='Route Map' /></div>"
        }
        catch { }
    }

    $genDate = (Get-Date).ToString("yyyy-MM-dd HH:mm")
    $safeName = [System.Net.WebUtility]::HtmlEncode($RouteName)
    $safeOrigin = [System.Net.WebUtility]::HtmlEncode($OriginAddress)
    $safeDest = [System.Net.WebUtility]::HtmlEncode($DestAddress)

    # Avoid options badges
    $avoidBadges = [System.Collections.Generic.List[string]]::new()
    if ($AvoidTolls)    { $avoidBadges.Add("<span class='badge badge-warn'>🚫 Avoid Tolls</span>") }
    if ($AvoidHighways) { $avoidBadges.Add("<span class='badge badge-warn'>🚫 Avoid Highways</span>") }
    if ($AvoidFerries)  { $avoidBadges.Add("<span class='badge badge-warn'>🚫 Avoid Ferries</span>") }
    $avoidHtml = if ($avoidBadges.Count -gt 0) { $avoidBadges -join ' ' } else { "<span class='badge badge-info'>None</span>" }

    # Stop itinerary rows
    $rowsHtml = [System.Text.StringBuilder]::new()
    # 1. Origin
    $oLatStr = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:F6}", $OriginLat)
    $oLngStr = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:F6}", $OriginLng)
    $null = $rowsHtml.AppendLine("<tr><td><span class='pin-badge pin-start'>A</span></td><td><strong>Origin (Start)</strong></td><td>$safeOrigin</td><td>$oLatStr, $oLngStr</td></tr>")

    # 2. Waypoints
    if ($Waypoints -and @($Waypoints).Count -gt 0) {
        $wIdx = 1
        foreach ($wp in $Waypoints) {
            $wLatStr = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:F6}", [double]$wp.Latitude)
            $wLngStr = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:F6}", [double]$wp.Longitude)
            $wAddr = [System.Net.WebUtility]::HtmlEncode($(if ($wp.Address) { [string]$wp.Address } else { "Waypoint $wIdx" }))
            $null = $rowsHtml.AppendLine("<tr><td><span class='pin-badge pin-wp'>$wIdx</span></td><td>Waypoint $wIdx</td><td>$wAddr</td><td>$wLatStr, $wLngStr</td></tr>")
            $wIdx++
        }
    }

    # 3. Destination
    $dLatStr = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:F6}", $DestLat)
    $dLngStr = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:F6}", $DestLng)
    $null = $rowsHtml.AppendLine("<tr><td><span class='pin-badge pin-dest'>B</span></td><td><strong>Destination (End)</strong></td><td>$safeDest</td><td>$dLatStr, $dLngStr</td></tr>")

    $tempHtml = Join-Path $env:TEMP "gmaps_report_$([System.Guid]::NewGuid().ToString('N')).html"

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>$safeName - Route Report</title>
  <style>
    @page { size: A4 portrait; margin: 12mm 14mm; }
    * { margin: 0; padding: 0; box-sizing: border-box; font-family: 'Segoe UI', Arial, sans-serif; }
    body { color: #1e293b; background: #ffffff; font-size: 12px; line-height: 1.4; }
    .header { border-bottom: 2px solid #2563eb; padding-bottom: 10px; margin-bottom: 14px; display: flex; justify-content: space-between; align-items: flex-end; }
    .header h1 { font-size: 20px; color: #0f172a; margin-bottom: 2px; }
    .header .subtitle { font-size: 11px; color: #64748b; }
    .header .meta { text-align: right; font-size: 11px; color: #64748b; }
    .kpi-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px; margin-bottom: 14px; }
    .kpi-card { background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px; padding: 8px 12px; }
    .kpi-label { font-size: 10px; text-transform: uppercase; color: #64748b; font-weight: 600; margin-bottom: 2px; }
    .kpi-value { font-size: 16px; font-weight: bold; color: #0f172a; }
    .kpi-value.blue { color: #2563eb; }
    .kpi-value.green { color: #059669; }
    .map-container { width: 100%; border: 1px solid #cbd5e1; border-radius: 6px; overflow: hidden; margin-bottom: 14px; page-break-inside: avoid; text-align: center; background: #0f172a; }
    .map-container img { width: 100%; height: auto; max-height: 480px; object-fit: contain; display: block; margin: 0 auto; }
    .section-title { font-size: 13px; font-weight: bold; color: #0f172a; border-bottom: 1px solid #e2e8f0; padding-bottom: 4px; margin-bottom: 8px; margin-top: 6px; }
    table { width: 100%; border-collapse: collapse; margin-bottom: 14px; font-size: 11px; }
    th { background: #f1f5f9; text-align: left; padding: 6px 8px; font-weight: 600; color: #475569; border-bottom: 1px solid #cbd5e1; }
    td { padding: 6px 8px; border-bottom: 1px solid #f1f5f9; vertical-align: middle; }
    tr:nth-child(even) td { background: #fafafa; }
    .pin-badge { display: inline-block; width: 20px; height: 20px; border-radius: 50%; color: white; text-align: center; line-height: 20px; font-weight: bold; font-size: 10px; }
    .pin-start { background: #10b981; }
    .pin-dest { background: #ef4444; }
    .pin-wp { background: #3b82f6; }
    .badge { display: inline-block; padding: 2px 6px; border-radius: 4px; font-size: 10px; font-weight: 600; }
    .badge-warn { background: #fef3c7; color: #92400e; border: 1px solid #fde68a; }
    .badge-info { background: #e0f2fe; color: #0369a1; border: 1px solid #bae6fd; }
    .footer { margin-top: 15px; border-top: 1px solid #e2e8f0; padding-top: 6px; font-size: 10px; color: #94a3b8; display: flex; justify-content: space-between; }
  </style>
</head>
<body>
  <div class="header">
    <div>
      <h1>$safeName</h1>
      <div class="subtitle">Google Maps Route & Navigation Report</div>
    </div>
    <div class="meta">
      Generated: <strong>$genDate</strong><br>
      Engine: <strong>Google Routes API v2</strong>
    </div>
  </div>

  <div class="kpi-grid">
    <div class="kpi-card">
      <div class="kpi-label">Total Distance</div>
      <div class="kpi-value blue">$DistanceKm km</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">Estimated Duration</div>
      <div class="kpi-value green">$DurationMin min</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">Optimization</div>
      <div class="kpi-value">$RouteType</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">Avoid Options</div>
      <div style="margin-top:2px;">$avoidHtml</div>
    </div>
  </div>

  $mapImgTag

  <div class="section-title">Route Itinerary & Stops</div>
  <table>
    <thead>
      <tr>
        <th style="width: 35px;">#</th>
        <th style="width: 140px;">Role / Point</th>
        <th>Address / Location</th>
        <th style="width: 160px;">Coordinates (Lat, Lng)</th>
      </tr>
    </thead>
    <tbody>
      $($rowsHtml.ToString())
    </tbody>
  </table>

  <div class="footer">
    <span>Google Maps Route & Map Generator v2.1</span>
    <span>$(if ($GoogleMapsUrl) { "<a href='$GoogleMapsUrl' style='color:#2563eb;text-decoration:none;'>Open in Google Maps &rarr;</a>" } else { '' })</span>
  </div>
</body>
</html>
"@

    [System.IO.File]::WriteAllText($tempHtml, $html, [System.Text.UTF8Encoding]::new($true))

    try {
        $p = Start-Process -FilePath $edgePath -ArgumentList "--headless", "--disable-gpu", "--no-pdf-header-footer", "--print-to-pdf=`"$OutputPath`"", "`"$tempHtml`"" -Wait -PassThru
        if ($p.ExitCode -eq 0 -and (Test-Path $OutputPath)) {
            if (Get-Command Write-AppLog -ErrorAction SilentlyContinue) { Write-AppLog "Generated PDF report successfully: $OutputPath" "OK" }
            return $OutputPath
        } else {
            throw "Edge process exited with code $($p.ExitCode) and PDF was not created."
        }
    }
    finally {
        if (Test-Path $tempHtml) {
            Remove-Item -Path $tempHtml -Force -ErrorAction SilentlyContinue
        }
    }
}

function Export-BatchPdfReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $true)][object[]]$Routes,
        [Parameter()][string]$BatchTitle = 'Batch Processing Summary Report',
        [Parameter()][string]$SourceFileName = ''
    )

    $edgePath = Find-EdgeExecutable
    if (-not $edgePath) {
        throw "Microsoft Edge executable (msedge.exe) not found. PDF export requires Edge."
    }

    $outDir = Split-Path -Parent $OutputPath
    if ($outDir -and -not (Test-Path $outDir)) {
        New-Item -ItemType Directory -Path $outDir -Force | Out-Null
    }

    $genDate = (Get-Date).ToString("yyyy-MM-dd HH:mm")
    $totalCount = @($Routes).Count
    $totalDist = 0.0
    $totalDur = 0
    $successCount = 0

    $rowsSb = [System.Text.StringBuilder]::new()
    foreach ($r in @($Routes)) {
        $dist = if ($r.DistanceKm) { [double]$r.DistanceKm } else { 0.0 }
        $dur  = if ($r.DurationMin) { [int]$r.DurationMin } else { 0 }
        $totalDist += $dist
        $totalDur  += $dur
        if ($r.Status -eq 'OK' -or $r.Status -eq 'Success') { $successCount++ }

        $statBadge = if ($r.Status -eq 'OK' -or $r.Status -eq 'Success') {
            "<span class='badge' style='background:#dcfce7;color:#166534;'>OK</span>"
        } else {
            "<span class='badge' style='background:#fee2e2;color:#991b1b;'>$([System.Net.WebUtility]::HtmlEncode([string]$r.Status))</span>"
        }

        $id = [System.Net.WebUtility]::HtmlEncode([string]$r.Id)
        $name = [System.Net.WebUtility]::HtmlEncode([string]$r.Name)
        $start = [System.Net.WebUtility]::HtmlEncode([string]$r.Start_Original)
        $end = [System.Net.WebUtility]::HtmlEncode([string]$r.End_Original)
        $type = [System.Net.WebUtility]::HtmlEncode([string]$r.RouteType)

        $null = $rowsSb.AppendLine("<tr><td>$id</td><td><strong>$name</strong></td><td>$start</td><td>$end</td><td>$type</td><td style='text-align:right;'><strong>$dist km</strong></td><td style='text-align:right;'>$dur min</td><td style='text-align:center;'>$statBadge</td></tr>")
    }
    $totalDistRound = [math]::Round($totalDist, 2)

    $tempHtml = Join-Path $env:TEMP "gmaps_batch_report_$([System.Guid]::NewGuid().ToString('N')).html"

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>$BatchTitle</title>
  <style>
    @page { size: A4 landscape; margin: 12mm 12mm; }
    * { margin: 0; padding: 0; box-sizing: border-box; font-family: 'Segoe UI', Arial, sans-serif; }
    body { color: #1e293b; background: #ffffff; font-size: 11px; line-height: 1.4; }
    .header { border-bottom: 2px solid #2563eb; padding-bottom: 8px; margin-bottom: 12px; display: flex; justify-content: space-between; align-items: flex-end; }
    .header h1 { font-size: 18px; color: #0f172a; }
    .header .meta { text-align: right; font-size: 10px; color: #64748b; }
    .kpi-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px; margin-bottom: 12px; }
    .kpi-card { background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px; padding: 6px 10px; }
    .kpi-label { font-size: 9px; text-transform: uppercase; color: #64748b; font-weight: 600; }
    .kpi-value { font-size: 15px; font-weight: bold; color: #0f172a; }
    table { width: 100%; border-collapse: collapse; font-size: 10px; }
    th { background: #f1f5f9; text-align: left; padding: 5px 6px; font-weight: 600; color: #475569; border-bottom: 1px solid #cbd5e1; }
    td { padding: 5px 6px; border-bottom: 1px solid #f1f5f9; vertical-align: middle; }
    tr:nth-child(even) td { background: #fafafa; }
    .badge { display: inline-block; padding: 2px 5px; border-radius: 4px; font-size: 9px; font-weight: bold; }
  </style>
</head>
<body>
  <div class="header">
    <div>
      <h1>$BatchTitle</h1>
      <div style="font-size: 10px; color: #64748b;">Source: $([System.Net.WebUtility]::HtmlEncode($SourceFileName))</div>
    </div>
    <div class="meta">Generated: <strong>$genDate</strong></div>
  </div>

  <div class="kpi-grid">
    <div class="kpi-card"><div class="kpi-label">Total Routes</div><div class="kpi-value">$totalCount</div></div>
    <div class="kpi-card"><div class="kpi-label">Successful Routes</div><div class="kpi-value" style="color:#059669;">$successCount / $totalCount</div></div>
    <div class="kpi-card"><div class="kpi-label">Total Distance</div><div class="kpi-value" style="color:#2563eb;">$totalDistRound km</div></div>
    <div class="kpi-card"><div class="kpi-label">Total Travel Time</div><div class="kpi-value">$totalDur min ($([math]::Round($totalDur / 60.0, 1)) h)</div></div>
  </div>

  <table>
    <thead>
      <tr>
        <th style="width: 30px;">#</th>
        <th>Route Name</th>
        <th>Origin</th>
        <th>Destination</th>
        <th style="width: 70px;">Type</th>
        <th style="width: 75px; text-align:right;">Distance</th>
        <th style="width: 65px; text-align:right;">Duration</th>
        <th style="width: 60px; text-align:center;">Status</th>
      </tr>
    </thead>
    <tbody>
      $($rowsSb.ToString())
    </tbody>
  </table>
</body>
</html>
"@

    [System.IO.File]::WriteAllText($tempHtml, $html, [System.Text.UTF8Encoding]::new($true))

    try {
        $p = Start-Process -FilePath $edgePath -ArgumentList "--headless", "--disable-gpu", "--no-pdf-header-footer", "--print-to-pdf=`"$OutputPath`"", "`"$tempHtml`"" -Wait -PassThru
        if ($p.ExitCode -eq 0 -and (Test-Path $OutputPath)) {
            if (Get-Command Write-AppLog -ErrorAction SilentlyContinue) { Write-AppLog "Generated batch PDF report successfully: $OutputPath" "OK" }
            return $OutputPath
        } else {
            throw "Edge process exited with code $($p.ExitCode) and PDF was not created."
        }
    }
    finally {
        if (Test-Path $tempHtml) {
            Remove-Item -Path $tempHtml -Force -ErrorAction SilentlyContinue
        }
    }
}

Set-Item -Path "function:global:Find-EdgeExecutable" -Value (Get-Item "function:Find-EdgeExecutable").ScriptBlock -ErrorAction SilentlyContinue
Set-Item -Path "function:global:Export-RoutePdfReport" -Value (Get-Item "function:Export-RoutePdfReport").ScriptBlock -ErrorAction SilentlyContinue
Set-Item -Path "function:global:Export-BatchPdfReport" -Value (Get-Item "function:Export-BatchPdfReport").ScriptBlock -ErrorAction SilentlyContinue
