#Requires -Version 5.1
<#
.SYNOPSIS
    Generates an interactive HTML uptime report for servers and network devices,
    with 30/60/90-day window switching, and posts it to a NinjaOne KB folder.

.DESCRIPTION
    Uses the NinjaOne Public API v2 with Client Credentials (silent, no browser).

    Device scope:
      Servers      : WINDOWS_SERVER, LINUX_SERVER, MAC_SERVER
      Network      : NMS_ROUTER, NMS_SWITCH, NMS_FIREWALL, NMS_OTHER,
                     NMS_UNKNOWN, NMS_PRINTER, NMS_STORAGE

    Uptime calculation:
      Pulls device activity logs for 90 days looking for offline/online events.
      Pairs OFFLINE events with the next ONLINE event to calculate downtime.
      If an OFFLINE event has no matching ONLINE event, the device is assumed to
      have recovered at its last contact time or end of window.
      Uptime % = (window seconds - downtime seconds) / window seconds * 100

    NOTE: Uptime is calculated from NinjaOne agent/NMS check-in events, not
    true packet-level availability. Accuracy depends on your polling interval.

    Three datasets (30/60/90 days) are pre-calculated and embedded in one HTML
    file. The dropdown in the report switches instantly with no re-fetch.

    Report is saved locally as a timestamped HTML file and opened automatically.
    KB posting will be added in a future revision.

    API APP SETUP (one-time):
    Administration > Apps > API > Client App IDs > Add
      Platform      : API Services (Machine-to-Machine)
      Allowed Scopes: monitoring
      Redirect URI  : Leave blank

    REGIONAL URLS:
    US: https://app.ninjarmm.com  |  EU: https://eu.ninjarmm.com
    OC: https://oc.ninjarmm.com   |  CA: https://ca.ninjarmm.com
#>

# ==============================================================================
#  CONFIGURATION -- Fill in ALL values before running
# ==============================================================================

$BaseUrl       = 'https://<your Login URL>'       # No trailing slash
$TokenEndpoint = 'https://<your Login URL>/ws/oauth/token'
$ClientId      = '<Your Client ID>'
$ClientSecret  = '<Your Client Secret>'

$UptimeWarnThreshold     = 99.0   # Below this % shows yellow
$UptimeCriticalThreshold = 95.0   # Below this % shows red

# ==============================================================================
#  DO NOT EDIT BELOW THIS LINE
# ==============================================================================

try { [Net.ServicePointManager]::SecurityProtocol =
    [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch {}
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName System.Web            -ErrorAction SilentlyContinue

$ServerClasses    = @('WINDOWS_SERVER','LINUX_SERVER','MAC_SERVER')
$NetworkClasses   = @('NMS_ROUTER','NMS_SWITCH','NMS_FIREWALL','NMS_OTHER',
                      'NMS_UNKNOWN','NMS_PRINTER','NMS_STORAGE')
$AllTargetClasses = $ServerClasses + $NetworkClasses

$OfflineTypes = @('OFFLINE','AGENT_OFFLINE','DEVICE_OFFLINE','NMS_OFFLINE',
                  'CONDITION_NMS_NETWORK_STATUS','CONDITION_NMS_NETWORK_STATUS_CHANGE',
                  'CONDITION_SYSTEM_UPTIME','CONDITION_NMS_SYSTEM_UPTIME')
$OnlineTypes  = @('ONLINE','AGENT_ONLINE','DEVICE_ONLINE','NMS_ONLINE',
                  'ONLINE_RESTORED','CONNECTION_RESTORED')

# -- Config validation ---------------------------------------------------------
$ConfigErrors = @()
if ($BaseUrl      -like '*<*') { $ConfigErrors += 'Fill in $BaseUrl' }
if ($ClientId     -like '*<*') { $ConfigErrors += 'Fill in $ClientId' }
if ($ClientSecret -like '*<*') { $ConfigErrors += 'Fill in $ClientSecret' }
if ($ConfigErrors.Count -gt 0) {
    Write-Host ''; Write-Host '  [!] Configuration errors:' -ForegroundColor Red
    $ConfigErrors | ForEach-Object { Write-Host "      - $_" -ForegroundColor Red }
    exit 1
}

# -- Helpers -------------------------------------------------------------------
function Get-Prop {
    param([object]$Obj, [string]$Name, [object]$Default = $null)
    if ($null -eq $Obj) { return $Default }
    $p = $Obj.PSObject.Properties[$Name]
    if ($null -eq $p -or $null -eq $p.Value) { return $Default }
    return $p.Value
}

$script:AuthHeaders = $null

function Invoke-NinjaApi {
    param([string]$Path, [string]$Method = 'GET', [string]$Body = $null, [int]$Retries = 3)
    $Attempt = 0
    while ($true) {
        $Attempt++
        $Params = @{ Uri = "$BaseUrl/v2/$Path"; Method = $Method; Headers = $script:AuthHeaders }
        if ($Body) { $Params.Body = $Body; $Params.ContentType = 'application/json' }
        try { return Invoke-RestMethod @Params }
        catch {
            $sc = $null; try { $sc = [int]$_.Exception.Response.StatusCode } catch {}
            if ($sc -ge 400 -and $sc -lt 500 -and $sc -ne 429) {
                throw "HTTP $sc on $Method /v2/$Path -- $_" }
            if ($Attempt -gt $Retries) { throw "Failed after $Retries retries -- $_" }
            $delay = if ($sc -eq 429) { 10 } else { [int][Math]::Pow(2, $Attempt) }
            Write-Host "    [~] Retry $Attempt in ${delay}s..." -ForegroundColor Yellow
            Start-Sleep -Seconds $delay
        }
    }
}

Write-Host ''
Write-Host '  ================================================================' -ForegroundColor Cyan
Write-Host '  NinjaOne Uptime Report -- Servers and Network Devices'           -ForegroundColor Cyan
Write-Host "  Warn: <${UptimeWarnThreshold}%  Critical: <${UptimeCriticalThreshold}%"  -ForegroundColor Cyan
Write-Host '  ================================================================' -ForegroundColor Cyan
Write-Host ''

# =============================================================================
#  STEP 1: Authenticate
# =============================================================================
Write-Host '  [1/5] Authenticating...' -ForegroundColor Cyan
try {
    $Token = Invoke-RestMethod -Uri $TokenEndpoint -Method POST `
        -ContentType 'application/x-www-form-urlencoded' `
        -Body "grant_type=client_credentials&client_id=$ClientId&client_secret=$ClientSecret&scope=monitoring"
    $script:AuthHeaders = @{ Authorization = "Bearer $($Token.access_token)"; Accept = 'application/json' }
} catch {
    Write-Host "  [!] Authentication failed: $_" -ForegroundColor Red; exit 1
}
Write-Host '  [OK] Authenticated.' -ForegroundColor Green

# =============================================================================
#  STEP 2: Load organizations and devices
# =============================================================================
Write-Host ''; Write-Host '  [2/5] Loading organizations and devices...' -ForegroundColor Cyan

$OrgMap = @{}
try {
    $After = $null
    do {
        $QS   = "organizations?pageSize=200$(if ($After) { "&after=$After" })"
        $Page = Invoke-NinjaApi -Path $QS
        $Items = if ($Page -is [array]) { $Page } else { @($Page) }
        foreach ($o in $Items) {
            $id = Get-Prop $o 'id'; $n = Get-Prop $o 'name'
            if ($id -and $n) { $OrgMap[[string]$id] = $n }
        }
        $After = if ($Items.Count -eq 200) { Get-Prop $Items[-1] 'id' } else { $null }
    } while ($After)
} catch { Write-Host "  [i] Could not load orgs: $_" -ForegroundColor Gray }

$AllDevices = New-Object System.Collections.ArrayList
$After = $null
do {
    $QS = "devices?pageSize=200$(if ($After) { "&after=$After" })"
    try {
        $Page  = Invoke-NinjaApi -Path $QS
        $Items = if ($Page -is [array]) { $Page } else { @($Page) }
        foreach ($d in $Items) { [void]$AllDevices.Add($d) }
        $After = if ($Items.Count -eq 200) { Get-Prop $Items[-1] 'id' } else { $null }
    } catch { Write-Host "  [!] Device page failed: $_" -ForegroundColor Red; break }
} while ($After)

$TargetDevices = $AllDevices | Where-Object {
    $nc = (Get-Prop $_ 'nodeClass' -Default '').ToUpperInvariant()
    $AllTargetClasses -contains $nc
}

$AllNodeClasses = $AllDevices | ForEach-Object { Get-Prop $_ 'nodeClass' -Default 'NULL' } | Sort-Object -Unique
Write-Host "  [OK] $($OrgMap.Count) org(s)  |  $($AllDevices.Count) total devices  |  $($TargetDevices.Count) matched" -ForegroundColor Green
Write-Host "       Node classes in your instance : $($AllNodeClasses -join ', ')" -ForegroundColor Gray
Write-Host "       Node classes this script uses : $($AllTargetClasses -join ', ')" -ForegroundColor Gray
if ($TargetDevices.Count -eq 0) {
    Write-Host "  [!] No devices matched. Check node classes above and add any missing ones to AllTargetClasses." -ForegroundColor Yellow
}

# =============================================================================
#  STEP 3: Pull activity logs and calculate uptime per device
# =============================================================================
Write-Host ''; Write-Host '  [3/5] Pulling activity logs and calculating uptime...' -ForegroundColor Cyan
Write-Host "       ($($TargetDevices.Count) devices -- one API call each)" -ForegroundColor Gray

$NowSec    = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
$From90Sec = [DateTimeOffset]::UtcNow.AddDays(-90).ToUnixTimeSeconds()
$From60Sec = [DateTimeOffset]::UtcNow.AddDays(-60).ToUnixTimeSeconds()
$From30Sec = [DateTimeOffset]::UtcNow.AddDays(-30).ToUnixTimeSeconds()
$Win90Sec  = $NowSec - $From90Sec
$Win60Sec  = $NowSec - $From60Sec
$Win30Sec  = $NowSec - $From30Sec

function Get-DowntimeSecs {
    param($Periods, [long]$WinStart, [long]$WinEnd)
    $total = 0
    foreach ($p in $Periods) {
        $s = [Math]::Max([long]$p.Start, $WinStart)
        $e = [Math]::Min([long]$p.End,   $WinEnd)
        if ($e -gt $s) { $total += ($e - $s) }
    }
    return $total
}

$DeviceResults = New-Object System.Collections.ArrayList
$TotalDev = @($TargetDevices).Count; $i = 0

foreach ($Device in $TargetDevices) {
    $i++
    $DevId      = Get-Prop $Device 'id'
    $DevName    = Get-Prop $Device 'systemName' -Default (Get-Prop $Device 'dnsName' -Default "Device $DevId")
    $OrgId      = Get-Prop $Device 'organizationId'
    $OrgName    = if ($OrgId -and $OrgMap.ContainsKey([string]$OrgId)) { $OrgMap[[string]$OrgId] } else { 'Unassigned' }
    $NodeClass  = (Get-Prop $Device 'nodeClass' -Default 'UNKNOWN').ToUpperInvariant()
    $DeviceType = if ($ServerClasses -contains $NodeClass) { 'Server' } else { 'Network' }
    $LastContact = Get-Prop $Device 'lastContact'
    $IsOffline   = [bool](Get-Prop $Device 'offline' -Default $false)

    $pct = if ($TotalDev -gt 0) { [int](($i / $TotalDev) * 100) } else { 100 }
    Write-Progress -Activity 'Processing devices' -Status "[$i/$TotalDev] $OrgName :: $DevName" -PercentComplete $pct

    $Events = @()
    try {
        $ActResp = Invoke-NinjaApi -Path "device/$DevId/activities?pageSize=1000&after=$From90Sec"
        $ActArr  = if ($ActResp.PSObject.Properties['activities']) { $ActResp.activities }
                   elseif ($ActResp -is [array]) { $ActResp }
                   else { @($ActResp) }

        $Events = @($ActArr | Where-Object {
            $t  = (Get-Prop $_ 'type'           -Default '').ToUpperInvariant()
            $st = (Get-Prop $_ 'statusCode'     -Default '').ToUpperInvariant()
            $sr = (Get-Prop $_ 'activityResult' -Default '').ToUpperInvariant()
            $c  = "$t $st $sr"
            ($OfflineTypes | Where-Object { $c -like "*$_*" }) -or
            ($OnlineTypes  | Where-Object { $c -like "*$_*" })
        } | Sort-Object { Get-Prop $_ 'activityTime' })
    } catch {
        Write-Host "    [i] No activities for $DevName" -ForegroundColor Gray
    }

    # If no activity events were logged but the device is currently offline,
    # synthesize an outage from lastContact to now so the device doesn't
    # incorrectly show 100% uptime just because nothing was logged.
    if ($Events.Count -eq 0 -and $IsOffline -and $LastContact) {
        $OutagePeriods = New-Object System.Collections.ArrayList
        [void]$OutagePeriods.Add(@{ Start = [long]$LastContact; End = $NowSec })
        Write-Host "    [i] $DevName -- offline with no events, outage assumed from lastContact ($LastSeen)" -ForegroundColor Yellow
    } else {

    $OutagePeriods = New-Object System.Collections.ArrayList
    $OutageStart   = $null

    foreach ($Ev in $Events) {
        $EvTime = [long](Get-Prop $Ev 'activityTime' -Default 0)
        if ($EvTime -eq 0) { continue }
        $c      = "$(Get-Prop $Ev 'type' -Default '') $(Get-Prop $Ev 'statusCode' -Default '') $(Get-Prop $Ev 'activityResult' -Default '')".ToUpperInvariant()
        $IsOff  = ($OfflineTypes | Where-Object { $c -like "*$_*" }) -as [bool]
        $IsOn   = ($OnlineTypes  | Where-Object { $c -like "*$_*" }) -as [bool]

        if ($IsOff -and -not $OutageStart) {
            $OutageStart = $EvTime
        } elseif ($IsOn -and $OutageStart) {
            [void]$OutagePeriods.Add(@{ Start = $OutageStart; End = $EvTime })
            $OutageStart = $null
        }
    }

    if ($OutageStart) {
        $CloseAt = if ($IsOffline) { $NowSec }
                   elseif ($LastContact -and [long]$LastContact -gt [long]$OutageStart) { [long]$LastContact }
                   else { $NowSec }
        [void]$OutagePeriods.Add(@{ Start = $OutageStart; End = $CloseAt })
    }

    } # end else (has events or is online)

    $Down90 = Get-DowntimeSecs -Periods $OutagePeriods -WinStart $From90Sec -WinEnd $NowSec
    $Down60 = Get-DowntimeSecs -Periods $OutagePeriods -WinStart $From60Sec -WinEnd $NowSec
    $Down30 = Get-DowntimeSecs -Periods $OutagePeriods -WinStart $From30Sec -WinEnd $NowSec

    $Up90 = [math]::Min([math]::Round(($Win90Sec - $Down90) / $Win90Sec * 100, 3), 100)
    $Up60 = [math]::Min([math]::Round(($Win60Sec - $Down60) / $Win60Sec * 100, 3), 100)
    $Up30 = [math]::Min([math]::Round(($Win30Sec - $Down30) / $Win30Sec * 100, 3), 100)

    $LastSeen = if ($LastContact) {
        [DateTimeOffset]::FromUnixTimeSeconds([long]$LastContact).ToLocalTime().ToString('yyyy-MM-dd HH:mm')
    } else { 'Unknown' }

    [void]$DeviceResults.Add([PSCustomObject]@{
        DeviceName      = $DevName
        OrgName         = $OrgName
        NodeClass       = $NodeClass
        DeviceType      = $DeviceType
        IsOffline       = $IsOffline
        LastSeen        = $LastSeen
        OutageCount     = $OutagePeriods.Count
        UptimePct90     = $Up90
        UptimePct60     = $Up60
        UptimePct30     = $Up30
        DowntimeMin90   = [math]::Round($Down90 / 60, 1)
        DowntimeMin60   = [math]::Round($Down60 / 60, 1)
        DowntimeMin30   = [math]::Round($Down30 / 60, 1)
        HasActivityData = ($Events.Count -gt 0)
        OutageSynthesized = ($Events.Count -eq 0 -and $IsOffline -and $LastContact)
    })
}
Write-Progress -Activity 'Processing devices' -Completed
Write-Host "  [OK] Uptime calculated for $($DeviceResults.Count) device(s)." -ForegroundColor Green

# =============================================================================
#  STEP 4: Build datasets and HTML
# =============================================================================
Write-Host ''; Write-Host '  [4/5] Building HTML report...' -ForegroundColor Cyan

function Build-Dataset {
    param($Rows, [string]$Win)
    $UF = "UptimePct$Win"; $DF = "DowntimeMin$Win"
    $Out = New-Object System.Collections.ArrayList
    foreach ($R in ($Rows | Sort-Object OrgName, DeviceName)) {
        $UV = $R.PSObject.Properties[$UF]; $DV = $R.PSObject.Properties[$DF]
        [void]$Out.Add([PSCustomObject]@{
            DeviceName  = $R.DeviceName
            OrgName     = $R.OrgName
            NodeClass   = $R.NodeClass
            DeviceType  = $R.DeviceType
            IsOffline   = $R.IsOffline
            LastSeen    = $R.LastSeen
            OutageCount = $R.OutageCount
            UptimePct   = if ($UV -and $null -ne $UV.Value) { [math]::Max([double]$UV.Value, 0.0) } else { 100.0 }
            DowntimeMin = if ($DV -and $null -ne $DV.Value) { [math]::Max([double]$DV.Value, 0.0) } else { 0.0 }
            HasData     = $R.HasActivityData
            Synthesized = if ($R.PSObject.Properties['OutageSynthesized']) { [bool]$R.OutageSynthesized } else { $false }
        })
    }
    return $Out.ToArray()
}

$Data90 = @(Build-Dataset -Rows $DeviceResults -Win '90')
$Data60 = @(Build-Dataset -Rows $DeviceResults -Win '60')
$Data30 = @(Build-Dataset -Rows $DeviceResults -Win '30')

$Json90 = $Data90 | ConvertTo-Json -Depth 3 -Compress
$Json60 = $Data60 | ConvertTo-Json -Depth 3 -Compress
$Json30 = $Data30 | ConvertTo-Json -Depth 3 -Compress

$GeneratedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
$UniqueOrgs  = ($Data90 | Select-Object -ExpandProperty OrgName -Unique | Sort-Object |
    ForEach-Object { "<option>$([System.Web.HttpUtility]::HtmlEncode($_))</option>" }) -join ''
$UniqueTypes = ($Data90 | Select-Object -ExpandProperty DeviceType -Unique | Sort-Object |
    ForEach-Object { "<option>$_</option>" }) -join ''

# NOTE: JS template literals use ${...} which PowerShell would expand inside a
# double-quoted here-string. All JS ${} are backtick-escaped as `${} so
# PowerShell leaves them alone. The PS variables ($Json90 etc.) are embedded
# directly without ${} wrapping so they expand correctly.
$Html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Server and Network Device Uptime Report</title>
<style>
:root{
  --blue:#2E75B6;--dark:#1F4E79;--border:#c8d8ee;--card:#f0f4fa;
  --text:#1a1a1a;--muted:#6b7a99;
  --green:#1a7a3f;--green-bg:#e8f5ee;--green-bar:#27ae60;
  --yellow:#b45309;--yellow-bg:#fffbeb;--yellow-bar:#f39c12;
  --red:#b91c1c;--red-bg:#fef2f2;--red-bar:#e74c3c;
  --gray:#4b5563;--gray-bg:#f3f4f6;
}
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#f8faff;color:var(--text);min-height:100vh;}
.header{background:linear-gradient(135deg,var(--dark),var(--blue));padding:22px 32px;}
.header h1{color:#fff;font-size:20px;font-weight:700;}
.header p{color:#bdd7ee;font-size:12px;margin-top:4px;}
.controls{background:#fff;border-bottom:1px solid var(--border);padding:12px 32px;display:flex;flex-wrap:wrap;gap:10px;align-items:center;}
select,input[type=text]{background:#f8faff;border:1px solid var(--border);border-radius:6px;padding:6px 10px;font-size:13px;color:var(--text);outline:none;}
select:focus,input:focus{border-color:var(--blue);}
input[type=text]{width:200px;}
.window-select{font-weight:700;color:var(--dark);background:#EEF4FB;border-color:var(--blue);}
.btn-ghost{background:transparent;color:var(--muted);border:1px solid var(--border);border-radius:6px;padding:6px 14px;font-size:13px;cursor:pointer;}
.btn-ghost:hover{background:var(--card);}
.count-label{font-size:12px;color:var(--muted);margin-left:auto;}
.stats{display:flex;gap:14px;padding:16px 32px;flex-wrap:wrap;}
.stat{background:#fff;border:1px solid var(--border);border-radius:8px;padding:14px 20px;min-width:140px;}
.stat-label{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;font-weight:600;}
.stat-value{font-size:26px;font-weight:700;color:var(--blue);line-height:1.2;margin-top:3px;}
.stat-value.green{color:var(--green);}.stat-value.yellow{color:var(--yellow);}.stat-value.red{color:var(--red);}
.stat-sub{font-size:11px;color:var(--muted);margin-top:2px;}
.main{padding:0 32px 40px;}
.org-section{margin-bottom:24px;}
.org-header{font-size:14px;font-weight:700;color:var(--dark);padding:10px 16px;background:#EEF4FB;border:1px solid var(--border);border-radius:6px 6px 0 0;border-bottom:none;display:flex;justify-content:space-between;align-items:center;}
.org-avg{font-size:12px;font-weight:600;color:var(--muted);}
.device-list{background:#fff;border:1px solid var(--border);border-radius:0 0 6px 6px;overflow:hidden;}
.device-row{display:grid;grid-template-columns:220px 1fr 90px 90px 100px 110px;align-items:center;padding:10px 16px;border-top:1px solid var(--border);gap:12px;font-size:13px;}
.device-row:hover{background:var(--card);}
.device-row:first-child{border-top:none;}
.device-name{font-weight:600;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.device-class{font-size:11px;color:var(--muted);margin-top:2px;}
.bar-wrap{background:#e8edf2;border-radius:4px;height:8px;overflow:hidden;}
.bar-fill{height:100%;border-radius:4px;transition:width 0.4s ease;}
.bar-green{background:var(--green-bar);}.bar-yellow{background:var(--yellow-bar);}.bar-red{background:var(--red-bar);}
.uptime-pct{font-weight:700;font-size:14px;text-align:right;}
.pct-green{color:var(--green);}.pct-yellow{color:var(--yellow);}.pct-red{color:var(--red);}
.downtime-cell{font-size:12px;color:var(--muted);text-align:right;}
.status-cell{text-align:center;}
.badge{display:inline-block;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:700;}
.badge-online{background:var(--green-bg);color:var(--green);}
.badge-offline{background:var(--red-bg);color:var(--red);}
.badge-nodata{background:var(--gray-bg);color:var(--gray);}
.last-seen{font-size:11px;color:var(--muted);}
.nodata-note{font-size:11px;color:var(--yellow);margin-top:2px;}
.col-header{display:grid;grid-template-columns:220px 1fr 90px 90px 100px 110px;padding:8px 16px;gap:12px;background:#f8faff;border-bottom:2px solid var(--border);}
.col-header span{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--muted);}
.col-header span:nth-child(3),.col-header span:nth-child(4){text-align:right;}
.col-header span:nth-child(5){text-align:center;}
.empty{text-align:center;padding:40px;color:var(--muted);font-style:italic;}
.footer{text-align:center;padding:16px;font-size:11px;color:var(--muted);border-top:1px solid var(--border);}
.window-badge{display:inline-block;background:var(--blue);color:#fff;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:700;margin-left:8px;vertical-align:middle;}
</style>
</head>
<body>
<div class="header">
  <h1>&#128200; Server and Network Uptime Report
    <span class="window-badge" id="windowBadge">90 Days</span>
  </h1>
  <p>Generated: $GeneratedAt &nbsp;&bull;&nbsp; Warn: &lt;${UptimeWarnThreshold}% &nbsp;&bull;&nbsp; Critical: &lt;${UptimeCriticalThreshold}% &nbsp;&bull;&nbsp; $BaseUrl</p>
</div>
<div class="controls">
  <select class="window-select" id="fWindow" onchange="switchWindow()">
    <option value="30">Last 30 Days</option>
    <option value="60">Last 60 Days</option>
    <option value="90" selected>Last 90 Days</option>
  </select>
  <input type="text" id="fSearch" placeholder="&#128269; Search device or org..." oninput="render()">
  <select id="fOrg" onchange="render()"><option value="">All Organizations</option>$UniqueOrgs</select>
  <select id="fType" onchange="render()"><option value="">All Device Types</option>$UniqueTypes</select>
  <select id="fStatus" onchange="render()">
    <option value="">All Status</option>
    <option value="online">Online Only</option>
    <option value="offline">Offline Only</option>
  </select>
  <button class="btn-ghost" onclick="resetFilters()">Reset</button>
  <span class="count-label" id="countLabel"></span>
</div>
<div class="stats" id="statBar"></div>
<div class="main" id="mainContent"></div>
<div class="footer">
  NinjaOne Uptime Report &mdash; Servers and Network Devices &mdash; Get-NinjaUptimeReport.ps1
  &mdash; Uptime is calculated from NinjaOne agent/NMS check-in events, not true packet-level availability.
</div>
<script>
const DATASETS = { '30':$Json30, '60':$Json60, '90':$Json90 };
const WARN = $UptimeWarnThreshold;
const CRIT = $UptimeCriticalThreshold;
let currentData = DATASETS['90'];

function switchWindow() {
  const w = document.getElementById('fWindow').value;
  currentData = DATASETS[w];
  document.getElementById('windowBadge').textContent = 'Last ' + w + ' Days';
  render();
}
function resetFilters() {
  document.getElementById('fSearch').value  = '';
  document.getElementById('fOrg').value     = '';
  document.getElementById('fType').value    = '';
  document.getElementById('fStatus').value  = '';
  render();
}
function pctClass(p) { return p < CRIT ? 'pct-red' : p < WARN ? 'pct-yellow' : 'pct-green'; }
function barClass(p) { return p < CRIT ? 'bar-red' : p < WARN ? 'bar-yellow' : 'bar-green'; }

function render() {
  const q    = document.getElementById('fSearch').value.trim().toLowerCase();
  const org  = document.getElementById('fOrg').value;
  const type = document.getElementById('fType').value;
  const stat = document.getElementById('fStatus').value;

  let rows = currentData.filter(r => {
    if (q    && !r.DeviceName.toLowerCase().includes(q) && !r.OrgName.toLowerCase().includes(q)) return false;
    if (org  && r.OrgName    !== org)  return false;
    if (type && r.DeviceType !== type) return false;
    if (stat === 'online'  &&  r.IsOffline) return false;
    if (stat === 'offline' && !r.IsOffline) return false;
    return true;
  });

  const total   = rows.length;
  const online  = rows.filter(r => !r.IsOffline).length;
  const offline = rows.filter(r =>  r.IsOffline).length;
  const noData  = rows.filter(r => !r.HasData).length;
  const avgUp   = total > 0 ? (rows.reduce((s,r) => s + r.UptimePct, 0) / total).toFixed(2) : '100.00';
  const worst   = total > 0 ? rows.slice().sort((a,b) => a.UptimePct - b.UptimePct)[0] : null;
  const avgClass = avgUp < CRIT ? 'red' : avgUp < WARN ? 'yellow' : 'green';

  document.getElementById('statBar').innerHTML =
    '<div class="stat"><div class="stat-label">Devices</div><div class="stat-value">' + total + '</div><div class="stat-sub">in view</div></div>' +
    '<div class="stat"><div class="stat-label">Fleet Avg Uptime</div><div class="stat-value ' + avgClass + '">' + avgUp + '%</div><div class="stat-sub">across all devices</div></div>' +
    '<div class="stat"><div class="stat-label">Currently Online</div><div class="stat-value green">' + online + '</div></div>' +
    '<div class="stat"><div class="stat-label">Currently Offline</div><div class="stat-value ' + (offline > 0 ? 'red' : 'green') + '">' + offline + '</div></div>' +
    (noData > 0 ? '<div class="stat"><div class="stat-label">No Activity Data</div><div class="stat-value yellow">' + noData + '</div><div class="stat-sub">100% assumed</div></div>' : '') +
    (worst ? '<div class="stat"><div class="stat-label">Lowest Uptime</div><div class="stat-value red">' + worst.UptimePct + '%</div><div class="stat-sub">' + worst.DeviceName + '</div></div>' : '');

  document.getElementById('countLabel').textContent = total + ' devices';

  const main = document.getElementById('mainContent');
  if (total === 0) {
    main.innerHTML = '<div class="empty">No devices match your filters.' +
      (currentData.length === 0 ? ' No server or network devices were found -- check node class output in the PowerShell console.' : '') +
      '</div>';
    return;
  }

  const orgs = {};
  rows.forEach(r => { if (!orgs[r.OrgName]) orgs[r.OrgName] = []; orgs[r.OrgName].push(r); });

  main.innerHTML = Object.keys(orgs).sort().map(orgName => {
    const devices  = orgs[orgName];
    const orgAvg   = (devices.reduce((s,d) => s + d.UptimePct, 0) / devices.length).toFixed(2);
    const orgAC    = orgAvg < CRIT ? 'pct-red' : orgAvg < WARN ? 'pct-yellow' : 'pct-green';
    const devRows  = devices.map(d => {
      const bc = barClass(d.UptimePct);
      const pc = pctClass(d.UptimePct);
      const badge = d.Synthesized
        ? '<span class="badge badge-offline">Offline*</span>'
        : !d.HasData
          ? '<span class="badge badge-nodata">No Data</span>'
          : d.IsOffline
            ? '<span class="badge badge-offline">Offline</span>'
            : '<span class="badge badge-online">Online</span>';
      const ndNote = d.Synthesized
        ? '<div class="nodata-note">&#9888; No events logged; downtime estimated from last contact</div>'
        : !d.HasData
          ? '<div class="nodata-note">&#9888; No events logged; 100% assumed (device online)</div>'
          : '';
      return '<div class="device-row">' +
        '<div><div class="device-name" title="' + d.DeviceName + '">' + d.DeviceName + '</div>' +
        '<div class="device-class">' + d.NodeClass + ' &bull; ' + d.OutageCount + ' outage(s)</div>' + ndNote + '</div>' +
        '<div><div class="bar-wrap"><div class="bar-fill ' + bc + '" style="width:' + Math.min(d.UptimePct,100) + '%"></div></div></div>' +
        '<div class="uptime-pct ' + pc + '">' + d.UptimePct + '%</div>' +
        '<div class="downtime-cell">' + d.DowntimeMin + ' min<br>downtime</div>' +
        '<div class="status-cell">' + badge + '<div class="last-seen">' + d.LastSeen + '</div></div>' +
        '</div>';
    }).join('');
    return '<div class="org-section">' +
      '<div class="org-header"><span>' + orgName + ' (' + devices.length + ' device' + (devices.length !== 1 ? 's' : '') + ')</span>' +
      '<span class="org-avg ' + orgAC + '">Org avg: ' + orgAvg + '%</span></div>' +
      '<div class="device-list"><div class="col-header"><span>Device</span><span>Uptime Bar</span><span>Uptime %</span><span>Downtime</span><span>Status</span></div>' +
      devRows + '</div></div>';
  }).join('');
}
render();
</script>
</body>
</html>
"@

# =============================================================================
#  STEP 5: Save report locally
# =============================================================================
Write-Host ''; Write-Host '  [5/5] Saving report...' -ForegroundColor Cyan

# -- Save local copy -----------------------------------------------------------
$Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.Path }
$LocalPath = Join-Path $ScriptDir "NinjaUptimeReport_$Timestamp.html"
try {
    [System.IO.File]::WriteAllText($LocalPath, $Html, [System.Text.Encoding]::UTF8)
    Write-Host "  [OK] Local copy saved: $LocalPath" -ForegroundColor Gray
} catch {
    $LocalPath = Join-Path $env:TEMP "NinjaUptimeReport_$Timestamp.html"
    [System.IO.File]::WriteAllText($LocalPath, $Html, [System.Text.Encoding]::UTF8)
    Write-Host "  [OK] Local copy saved to TEMP: $LocalPath" -ForegroundColor Gray
}

$FleetAvg90 = if ($Data90.Count -gt 0) {
    [math]::Round(($Data90 | Measure-Object UptimePct -Average).Average, 2)
} else { 100 }

Write-Host ''
Write-Host '  ================================================================' -ForegroundColor Green
Write-Host '  [OK] COMPLETE'                                                    -ForegroundColor Green
Write-Host "       Devices processed : $($DeviceResults.Count)"                -ForegroundColor Green
Write-Host "       Fleet avg uptime  : $FleetAvg90% (90-day)"                  -ForegroundColor Green
Write-Host "       Local file        : $LocalPath"                             -ForegroundColor Green
Write-Host '  ================================================================' -ForegroundColor Green
Write-Host ''
try { Start-Process $LocalPath } catch {}
