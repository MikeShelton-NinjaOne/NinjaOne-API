#Requires -Version 5.1
<#
.SYNOPSIS
    Generates an HTML report showing tickets that sat unassigned for longer
    than a configurable threshold (default: 15 minutes) before being assigned,
    plus any tickets that are still unassigned today.

.DESCRIPTION
    Uses the NinjaOne Public API v2 with Client Credentials (silent, no browser).

    For each ticket in the lookback window the script:
      1. Records the ticket creation time
      2. Pulls the ticket log entries
      3. Finds the FIRST assignment log entry
      4. Calculates the gap between ticket creation and first assignment
      5. Flags the ticket if that gap exceeds $ThresholdMinutes
      6. Also flags tickets that have NO assignment log entry at all
         (still unassigned as of right now)

    The HTML report includes:
      - Summary stat cards (total breached, still unassigned, avg wait time,
        worst offender)
      - Filterable / sortable table (by org, technician, wait time, priority)
      - Colour-coded wait-time badges (yellow > threshold, red > 2x threshold)
      - Sortable columns
      - Date range filter controls

    Output: self-contained HTML file saved to the same folder as this script.

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

# Your NinjaOne login URL (no trailing slash)
$BaseUrl       = 'https://<your Login URL>'

# Same URL with /ws/oauth/token appended
$TokenEndpoint = 'https://<your Login URL>/ws/oauth/token'

# From Administration > Apps > API > Client App IDs
$ClientId      = '<Your Client ID>'

# From Administration > Apps > API > Client App IDs (shown once at creation)
$ClientSecret  = '<Your Client Secret>'

# How many minutes before a ticket is considered "slow to assign"
# Change this to suit your SLA -- e.g. 30 for half an hour, 60 for one hour
$ThresholdMinutes = 15

# How many days back to look for tickets (default: last 30 days)
$LookbackDays = 30

# ==============================================================================
#  DO NOT EDIT BELOW THIS LINE
# ==============================================================================

# Force TLS 1.2
try {
    [Net.ServicePointManager]::SecurityProtocol = `
        [Net.ServicePointManager]::SecurityProtocol -bor `
        [Net.SecurityProtocolType]::Tls12
} catch {}

try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$ErrorActionPreference = 'Continue'

Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue

# -- Validate config ----------------------------------------------------------
$ConfigErrors = @()
if ($BaseUrl      -like '*<*') { $ConfigErrors += 'Fill in $BaseUrl' }
if ($ClientId     -like '*<*') { $ConfigErrors += 'Fill in $ClientId' }
if ($ClientSecret -like '*<*') { $ConfigErrors += 'Fill in $ClientSecret' }
if ($ConfigErrors.Count -gt 0) {
    Write-Host ''
    Write-Host '  [!] Configuration errors:' -ForegroundColor Red
    $ConfigErrors | ForEach-Object { Write-Host "      - $_" -ForegroundColor Red }
    exit 1
}

# -- Safe property helper -----------------------------------------------------
function Get-Prop {
    param([object]$Obj, [string]$Name, [object]$Default = $null)
    if ($null -eq $Obj) { return $Default }
    $p = $Obj.PSObject.Properties[$Name]
    if ($null -eq $p -or $null -eq $p.Value) { return $Default }
    return $p.Value
}

# -- Epoch helpers ------------------------------------------------------------
function ConvertFrom-EpochMs {
    param([long]$Ms)
    [DateTimeOffset]::FromUnixTimeMilliseconds($Ms).ToLocalTime()
}

function Format-LocalTime {
    param([DateTimeOffset]$Dt)
    $Dt.ToString('yyyy-MM-dd HH:mm')
}

# -- API helper with retry ----------------------------------------------------
$script:Headers = $null

function Invoke-NinjaApi {
    param([string]$Path, [string]$Method = 'GET', [string]$Body = $null, [int]$Retries = 3)
    $Attempt = 0
    while ($true) {
        $Attempt++
        $Params = @{ Uri = "$BaseUrl/v2/$Path"; Method = $Method; Headers = $script:Headers }
        if ($Body) { $Params.Body = $Body; $Params.ContentType = 'application/json' }
        try {
            return Invoke-RestMethod @Params
        } catch {
            $sc = $null; try { $sc = [int]$_.Exception.Response.StatusCode } catch {}
            if ($sc -ge 400 -and $sc -lt 500 -and $sc -ne 429) { throw "HTTP $sc on $Method /v2/$Path -- $_" }
            if ($Attempt -gt $Retries) { throw "Failed after $Retries retries on $Method /v2/$Path -- $_" }
            $delay = if ($sc -eq 429) { 10 } else { [int][Math]::Pow(2, $Attempt) }
            Write-Host "    [~] Retry $Attempt/$Retries in ${delay}s (HTTP $sc)..." -ForegroundColor Yellow
            Start-Sleep -Seconds $delay
        }
    }
}

Write-Host ''
Write-Host '  ================================================================' -ForegroundColor Cyan
Write-Host "  NinjaOne Unassigned Ticket Report  [Threshold: ${ThresholdMinutes}min]" -ForegroundColor Cyan
Write-Host "  Lookback: $LookbackDays days" -ForegroundColor Cyan
Write-Host '  ================================================================' -ForegroundColor Cyan
Write-Host ''

# =============================================================================
#  STEP 1: Authenticate
# =============================================================================
Write-Host '  [1/5] Authenticating...' -ForegroundColor Cyan
try {
    $Token = Invoke-RestMethod -Uri $TokenEndpoint -Method POST `
        -ContentType 'application/x-www-form-urlencoded' `
        -Body @{
            grant_type    = 'client_credentials'
            client_id     = $ClientId
            client_secret = $ClientSecret
            scope         = 'monitoring'
        }
    $script:Headers = @{
        Authorization = "Bearer $($Token.access_token)"
        Accept        = 'application/json'
    }
} catch {
    Write-Host "  [!] Authentication failed: $_" -ForegroundColor Red
    exit 1
}
Write-Host '  [OK] Authenticated.' -ForegroundColor Green

# =============================================================================
#  STEP 2: Load orgs and technicians for display names
# =============================================================================
Write-Host ''
Write-Host '  [2/5] Loading organizations and technicians...' -ForegroundColor Cyan

$OrgMap  = @{}  # orgId -> name
$TechMap = @{}  # techId -> name

try {
    $AllOrgs = New-Object System.Collections.ArrayList
    $After = $null
    do {
        $QS   = "organizations?pageSize=200$(if ($After) { "&after=$After" })"
        $Page = Invoke-NinjaApi -Path $QS
        $Items = if ($Page -is [array]) { $Page } else { @($Page) }
        foreach ($o in $Items) { [void]$AllOrgs.Add($o) }
        $After = if ($Items.Count -eq 200) { Get-Prop $Items[-1] 'id' } else { $null }
    } while ($After)
    foreach ($o in $AllOrgs) {
        $id = Get-Prop $o 'id'; $n = Get-Prop $o 'name'
        if ($id -and $n) { $OrgMap[[string]$id] = $n }
    }
} catch { Write-Host "  [i] Could not load orgs: $_" -ForegroundColor Gray }

try {
    $Techs = Invoke-NinjaApi -Path 'technicians'
    $TechArr = if ($Techs -is [array]) { $Techs } else { @($Techs) }
    foreach ($t in $TechArr) {
        $id = Get-Prop $t 'id'
        $fn = Get-Prop $t 'firstName' -Default ''
        $ln = Get-Prop $t 'lastName'  -Default ''
        $nm = Get-Prop $t 'name'
        if ($id) {
            $TechMap[[string]$id] = if ($nm) { $nm } else { "$fn $ln".Trim() }
        }
    }
} catch { Write-Host "  [i] Could not load technicians: $_" -ForegroundColor Gray }

Write-Host "  [OK] $($OrgMap.Count) org(s), $($TechMap.Count) technician(s)." -ForegroundColor Green

# =============================================================================
#  STEP 3: Pull tickets in the lookback window
# =============================================================================
Write-Host ''
Write-Host "  [3/5] Fetching tickets from the last $LookbackDays days..." -ForegroundColor Cyan

$FromMs   = [DateTimeOffset]::UtcNow.AddDays(-$LookbackDays).ToUnixTimeMilliseconds()
$AllTickets = New-Object System.Collections.ArrayList
$LastCursor = $null
$StopPaging = $false

do {
    $Path = 'ticketing/ticket?pageSize=200'
    if ($LastCursor) { $Path += "&after=$LastCursor" }
    try {
        $Resp = Invoke-NinjaApi -Path $Path
    } catch {
        Write-Host "  [!] Failed to fetch tickets: $_" -ForegroundColor Red
        break
    }

    $Data = Get-Prop $Resp 'data'
    if (-not $Data -or $Data.Count -eq 0) { $StopPaging = $true; break }

    $DataArr = if ($Data -is [array]) { $Data } else { @($Data) }
    foreach ($t in $DataArr) { [void]$AllTickets.Add($t) }

    # Stop paging when tickets go older than our window
    $OldestOnPage = ($DataArr | Sort-Object { Get-Prop $_ 'createTime' } | Select-Object -First 1)
    $OldestCreate = Get-Prop $OldestOnPage 'createTime'
    if ($OldestCreate -and $OldestCreate -lt $FromMs) { $StopPaging = $true }

    $Meta = Get-Prop $Resp 'metadata'
    $LastCursor = if ($Meta) { Get-Prop $Meta 'lastCursorId' } else { $null }
    if (-not $LastCursor) { $StopPaging = $true }

} while (-not $StopPaging)

# Filter to the lookback window
$WindowTickets = $AllTickets | Where-Object { (Get-Prop $_ 'createTime') -ge $FromMs }
Write-Host "  [OK] $($WindowTickets.Count) ticket(s) in the last $LookbackDays days." -ForegroundColor Green

# =============================================================================
#  STEP 4: Analyse each ticket -- check assignment timing via log entries
# =============================================================================
Write-Host ''
Write-Host '  [4/5] Analysing assignment timing per ticket...' -ForegroundColor Cyan

$ThresholdMs   = $ThresholdMinutes * 60 * 1000
$ReportRows    = New-Object System.Collections.ArrayList
$Total         = @($WindowTickets).Count
$i             = 0

foreach ($Ticket in $WindowTickets) {
    $i++
    $TicketId  = Get-Prop $Ticket 'id'
    $Subject   = Get-Prop $Ticket 'subject'    -Default "(no subject)"
    $Priority  = Get-Prop $Ticket 'priority'   -Default 'NONE'
    $Status    = Get-Prop $Ticket 'status'     -Default 'UNKNOWN'
    $OrgId     = Get-Prop $Ticket 'clientId'
    $CreateMs  = Get-Prop $Ticket 'createTime'
    $OrgName   = if ($OrgId -and $OrgMap.ContainsKey([string]$OrgId)) { $OrgMap[[string]$OrgId] } else { "Org $OrgId" }

    $pct = if ($Total -gt 0) { [int](($i / $Total) * 100) } else { 100 }
    Write-Progress -Activity 'Analysing tickets' `
                   -Status "[$i/$Total] $Subject" `
                   -PercentComplete $pct

    if (-not $CreateMs) { continue }

    # Pull log entries to find the first assignment event
    $FirstAssignMs    = $null
    $AssignedTechName = $null
    $AssignedByName   = $null   # the person who PERFORMED the assignment

    try {
        $Logs = Invoke-NinjaApi -Path "ticketing/ticket/$TicketId/log-entry"
        $LogArr = if ($Logs -is [array]) { $Logs } else { @($Logs) }

        # Look for log entries that indicate an assignment
        # Type can be ASSIGNMENT, or STATUS_CHANGE entries that include assignee data
        $AssignLogs = $LogArr | Where-Object {
            $t = (Get-Prop $_ 'type' -Default '').ToUpperInvariant()
            $t -eq 'ASSIGNMENT' -or $t -eq 'TECHNICIAN_CHANGED' -or $t -eq 'ASSIGNED'
        } | Sort-Object { Get-Prop $_ 'createTime' }

        # Fallback: if no explicit ASSIGNMENT type, look for any log entry
        # that has appUserContactType = TECHNICIAN and references a user change
        if (-not $AssignLogs -or $AssignLogs.Count -eq 0) {
            $AssignLogs = $LogArr | Where-Object {
                $t    = (Get-Prop $_ 'type' -Default '').ToUpperInvariant()
                $ct   = (Get-Prop $_ 'appUserContactType' -Default '').ToUpperInvariant()
                $body = Get-Prop $_ 'body' -Default ''
                # Some NinjaOne instances log assignments as status changes with body text
                ($t -eq 'ACTIVITY' -or $t -eq 'STATUS') -and (
                    $body -match 'assign' -or $ct -eq 'TECHNICIAN'
                )
            } | Sort-Object { Get-Prop $_ 'createTime' }
        }

        if ($AssignLogs -and $AssignLogs.Count -gt 0) {
            $FirstAssign   = $AssignLogs | Select-Object -First 1
            $FirstAssignMs = Get-Prop $FirstAssign 'createTime'

            # Who the ticket was assigned TO -- appUserContactId
            $AssignedById  = Get-Prop $FirstAssign 'appUserContactId'
            $AssignedTechName = if ($AssignedById -and $TechMap.ContainsKey([string]$AssignedById)) {
                $TechMap[[string]$AssignedById]
            } elseif ($AssignedById) {
                "Tech $AssignedById"
            } else {
                'Unknown'
            }

            # Who PERFORMED the assignment -- createdBy object on the log entry
            # NinjaOne returns this as { id, name } on the log entry actor
            $CreatedByObj = Get-Prop $FirstAssign 'createdBy'
            if ($CreatedByObj) {
                # createdBy is a nested object with id and name
                $ActorName = Get-Prop $CreatedByObj 'name'
                $ActorId   = Get-Prop $CreatedByObj 'id'
                if ($ActorName) {
                    $AssignedByName = $ActorName
                } elseif ($ActorId -and $TechMap.ContainsKey([string]$ActorId)) {
                    $AssignedByName = $TechMap[[string]$ActorId]
                } elseif ($ActorId) {
                    $AssignedByName = "User $ActorId"
                }
            }
            # Fallback: some versions use actorUserId as a flat field
            if (-not $AssignedByName) {
                $ActorId = Get-Prop $FirstAssign 'actorUserId'
                if (-not $ActorId) { $ActorId = Get-Prop $FirstAssign 'userId' }
                if ($ActorId -and $TechMap.ContainsKey([string]$ActorId)) {
                    $AssignedByName = $TechMap[[string]$ActorId]
                } elseif ($ActorId) {
                    $AssignedByName = "User $ActorId"
                }
            }
            if (-not $AssignedByName) { $AssignedByName = 'Unknown' }
        }
    } catch {
        # Non-fatal -- log entry fetch failed for this ticket
    }

    # Calculate the gap
    $GapMs         = $null
    $GapMinutes    = $null
    $StillUnassigned = $false

    if ($FirstAssignMs -and $CreateMs) {
        $GapMs      = [long]$FirstAssignMs - [long]$CreateMs
        $GapMinutes = [math]::Round($GapMs / 1000 / 60, 1)
    } elseif (-not $FirstAssignMs) {
        # Check current assignee on ticket object
        $CurrentAssignee = Get-Prop $Ticket 'assignedAppUserId'
        if (-not $CurrentAssignee) {
            $StillUnassigned = $true
            # Gap = now - createTime
            $NowMs      = [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()
            $GapMs      = $NowMs - [long]$CreateMs
            $GapMinutes = [math]::Round($GapMs / 1000 / 60, 1)
            $AssignedTechName = '— Still Unassigned —'
            $AssignedByName   = '—'
        }
    }

    # Include only tickets that breached the threshold
    if ($null -eq $GapMinutes) { continue }
    if ($GapMinutes -le $ThresholdMinutes -and -not $StillUnassigned) { continue }

    $CreateDto = ConvertFrom-EpochMs -Ms ([long]$CreateMs)

    [void]$ReportRows.Add([PSCustomObject]@{
        TicketId         = $TicketId
        Subject          = $Subject
        Priority         = $Priority
        Status           = $Status
        OrgName          = $OrgName
        CreatedAt        = Format-LocalTime $CreateDto
        CreatedAtSort    = $CreateMs
        GapMinutes       = $GapMinutes
        AssignedTo       = $AssignedTechName
        AssignedBy       = $AssignedByName
        StillUnassigned  = $StillUnassigned
    })
}

Write-Progress -Activity 'Analysing tickets' -Completed

$BreachedCount     = ($ReportRows | Where-Object { -not $_.StillUnassigned }).Count
$StillUnassigned   = ($ReportRows | Where-Object { $_.StillUnassigned }).Count
$AvgWait           = if ($ReportRows.Count -gt 0) { [math]::Round(($ReportRows | Measure-Object GapMinutes -Average).Average, 1) } else { 0 }
$WorstGap          = if ($ReportRows.Count -gt 0) { ($ReportRows | Sort-Object GapMinutes -Descending | Select-Object -First 1) } else { $null }

Write-Host "  [OK] $($ReportRows.Count) ticket(s) breached the ${ThresholdMinutes}-minute threshold." -ForegroundColor Green
Write-Host "       $BreachedCount eventually assigned, $StillUnassigned still unassigned." -ForegroundColor Green

# =============================================================================
#  STEP 5: Build HTML report
# =============================================================================
Write-Host ''
Write-Host '  [5/5] Building HTML report...' -ForegroundColor Cyan

$GeneratedAt   = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
$DoubleThresh  = $ThresholdMinutes * 2
$JsonData      = $ReportRows | Sort-Object GapMinutes -Descending | ConvertTo-Json -Depth 3 -Compress
$WorstStr      = if ($WorstGap) { "$($WorstGap.GapMinutes) min ($($WorstGap.Subject))" } else { 'N/A' }

# Collect unique orgs and priorities for filter dropdowns
$UniqueOrgs   = ($ReportRows | Select-Object -ExpandProperty OrgName -Unique | Sort-Object |
    ForEach-Object { "<option>$([System.Web.HttpUtility]::HtmlEncode($_))</option>" }) -join ''
$UniquePris   = ($ReportRows | Select-Object -ExpandProperty Priority -Unique | Sort-Object |
    ForEach-Object { "<option>$([System.Web.HttpUtility]::HtmlEncode($_))</option>" }) -join ''

$Html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Unassigned Ticket Report</title>
<style>
:root {
  --blue:#2E75B6; --dark:#1F4E79; --card:#f0f4fa; --border:#c8d8ee;
  --text:#1a1a1a; --muted:#6b7a99;
  --green:#1a7a3f; --green-bg:#e8f5ee;
  --red:#b91c1c; --red-bg:#fef2f2;
  --yellow:#b45309; --yellow-bg:#fffbeb;
  --orange:#c2410c; --orange-bg:#fff7ed;
  --gray:#4b5563; --gray-bg:#f3f4f6;
}
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
       background: #f8faff; color: var(--text); min-height: 100vh; }
.header { background: linear-gradient(135deg, var(--dark), var(--blue));
          padding: 22px 32px; }
.header h1 { color: #fff; font-size: 20px; font-weight: 700; }
.header p  { color: #bdd7ee; font-size: 12px; margin-top: 4px; }
.stats { display: flex; gap: 14px; padding: 16px 32px; flex-wrap: wrap; }
.stat { background: #fff; border: 1px solid var(--border); border-radius: 8px;
        padding: 14px 20px; min-width: 140px; }
.stat-label { font-size: 11px; color: var(--muted); text-transform: uppercase;
              letter-spacing: .5px; font-weight: 600; }
.stat-value { font-size: 26px; font-weight: 700; color: var(--blue);
              line-height: 1.2; margin-top: 3px; }
.stat-value.red    { color: var(--red); }
.stat-value.yellow { color: var(--yellow); }
.stat-value.green  { color: var(--green); }
.stat-sub { font-size: 11px; color: var(--muted); margin-top: 2px; }
.controls { background: #fff; border-bottom: 1px solid var(--border);
            padding: 12px 32px; display: flex; flex-wrap: wrap; gap: 10px;
            align-items: center; }
select, input[type=text] { background: #f8faff; border: 1px solid var(--border);
  border-radius: 6px; padding: 6px 10px; font-size: 13px; color: var(--text);
  outline: none; }
select:focus, input:focus { border-color: var(--blue); }
input[type=text] { width: 240px; }
.btn-ghost { background: transparent; color: var(--muted); border: 1px solid var(--border);
             border-radius: 6px; padding: 6px 14px; font-size: 13px; cursor: pointer; }
.btn-ghost:hover { background: var(--card); }
.count-label { font-size: 12px; color: var(--muted); margin-left: auto; }
.main { padding: 16px 32px 40px; }
.table-wrap { background: #fff; border: 1px solid var(--border);
              border-radius: 8px; overflow: hidden; }
table { width: 100%; border-collapse: collapse; font-size: 13px; }
th { background: var(--dark); color: #fff; padding: 10px 14px; text-align: left;
     font-size: 11px; font-weight: 700; text-transform: uppercase;
     letter-spacing: .5px; cursor: pointer; user-select: none; white-space: nowrap; }
th:hover { background: var(--blue); }
td { padding: 9px 14px; border-top: 1px solid var(--border); vertical-align: middle; }
tr:hover td { background: var(--card); }
.badge { display: inline-block; padding: 2px 9px; border-radius: 20px;
         font-size: 11px; font-weight: 700; text-transform: uppercase; }
.badge-ok       { background: var(--green-bg); color: var(--green); }
.badge-warn     { background: var(--yellow-bg); color: var(--yellow); }
.badge-bad      { background: var(--orange-bg); color: var(--orange); }
.badge-critical { background: var(--red-bg); color: var(--red); }
.badge-unassigned { background: var(--red-bg); color: var(--red); font-style: italic; }
.pri-CRITICAL { color: var(--red); font-weight: 700; }
.pri-HIGH     { color: var(--orange); font-weight: 600; }
.pri-MEDIUM   { color: var(--yellow); }
.pri-LOW      { color: var(--green); }
.pri-NONE     { color: var(--muted); }
.empty { text-align: center; padding: 60px 20px; color: var(--muted);
         font-style: italic; }
.hidden { display: none !important; }
.footer { text-align: center; padding: 16px; font-size: 11px; color: var(--muted);
          border-top: 1px solid var(--border); }
</style>
</head>
<body>

<div class="header">
  <h1>&#9203; Unassigned Ticket Report</h1>
  <p>Generated: $GeneratedAt &nbsp;&bull;&nbsp; Threshold: ${ThresholdMinutes} minutes &nbsp;&bull;&nbsp; Lookback: $LookbackDays days &nbsp;&bull;&nbsp; Instance: $BaseUrl</p>
</div>

<div class="stats" id="statBar"></div>

<div class="controls">
  <input type="text" id="fSearch" placeholder="&#128269; Search subject or org..." oninput="render()">
  <select id="fOrg"      onchange="render()"><option value="">All Organizations</option>$UniqueOrgs</select>
  <select id="fPriority" onchange="render()"><option value="">All Priorities</option>$UniquePris</select>
  <select id="fStatus"   onchange="render()">
    <option value="">All</option>
    <option value="unassigned">Still Unassigned Only</option>
    <option value="assigned">Eventually Assigned Only</option>
  </select>
  <button class="btn-ghost" onclick="resetFilters()">Reset</button>
  <span class="count-label" id="countLabel"></span>
</div>

<div class="main">
  <div class="table-wrap">
    <table>
      <thead><tr>
        <th onclick="sortBy('TicketId')">Ticket # &#8597;</th>
        <th onclick="sortBy('Subject')">Subject &#8597;</th>
        <th onclick="sortBy('OrgName')">Organization &#8597;</th>
        <th onclick="sortBy('Priority')">Priority &#8597;</th>
        <th onclick="sortBy('Status')">Status &#8597;</th>
        <th onclick="sortBy('CreatedAt')">Created &#8597;</th>
        <th onclick="sortBy('GapMinutes')">Wait Time &#8597;</th>
        <th onclick="sortBy('AssignedTo')">Assigned To &#8597;</th>
        <th onclick="sortBy('AssignedBy')">Assigned By &#8597;</th>
      </tr></thead>
      <tbody id="tbody"></tbody>
    </table>
    <div class="empty" id="emptyMsg" style="display:none;">No tickets match your filters.</div>
  </div>
</div>

<div class="footer">
  NinjaOne Unassigned Ticket Report &mdash; Threshold: ${ThresholdMinutes} min &mdash; Generated by Get-NinjaUnassignedTicketReport.ps1
</div>

<script>
const DATA      = $JsonData;
const THRESHOLD = $ThresholdMinutes;
const DOUBLE    = $DoubleThresh;

let sortCol = 'GapMinutes', sortDir = -1;  // default: worst first

function sortBy(col) {
  if (sortCol === col) { sortDir *= -1; } else { sortCol = col; sortDir = 1; }
  render();
}

function resetFilters() {
  document.getElementById('fSearch').value   = '';
  document.getElementById('fOrg').value      = '';
  document.getElementById('fPriority').value = '';
  document.getElementById('fStatus').value   = '';
  render();
}

function waitBadge(mins, unassigned) {
  if (unassigned) return '<span class="badge badge-unassigned">&#128308; Still Unassigned</span>';
  if (mins > DOUBLE * 2) return '<span class="badge badge-critical">' + mins + ' min</span>';
  if (mins > DOUBLE)     return '<span class="badge badge-bad">'      + mins + ' min</span>';
  return                        '<span class="badge badge-warn">'     + mins + ' min</span>';
}

function priClass(p) {
  switch ((p || '').toUpperCase()) {
    case 'CRITICAL': return 'pri-CRITICAL';
    case 'HIGH':     return 'pri-HIGH';
    case 'MEDIUM':   return 'pri-MEDIUM';
    case 'LOW':      return 'pri-LOW';
    default:         return 'pri-NONE';
  }
}

function render() {
  const q    = document.getElementById('fSearch').value.trim().toLowerCase();
  const org  = document.getElementById('fOrg').value;
  const pri  = document.getElementById('fPriority').value;
  const stat = document.getElementById('fStatus').value;

  let rows = DATA.filter(r => {
    if (q   && !r.Subject.toLowerCase().includes(q) && !r.OrgName.toLowerCase().includes(q)) return false;
    if (org  && r.OrgName  !== org)  return false;
    if (pri  && r.Priority !== pri)  return false;
    if (stat === 'unassigned' && !r.StillUnassigned)  return false;
    if (stat === 'assigned'   &&  r.StillUnassigned)  return false;
    return true;
  });

  rows.sort((a, b) => {
    const av = a[sortCol]; const bv = b[sortCol];
    if (av == null && bv == null) return 0;
    if (av == null) return sortDir;
    if (bv == null) return -sortDir;
    return av < bv ? -sortDir : av > bv ? sortDir : 0;
  });

  // Update stats
  const total       = rows.length;
  const unassigned  = rows.filter(r => r.StillUnassigned).length;
  const assigned    = total - unassigned;
  const avgWait     = total > 0
    ? (rows.reduce((s, r) => s + r.GapMinutes, 0) / total).toFixed(1)
    : '0';
  const worst = rows.length > 0
    ? rows.slice().sort((a,b) => b.GapMinutes - a.GapMinutes)[0]
    : null;

  document.getElementById('statBar').innerHTML = `
    <div class="stat"><div class="stat-label">Showing</div><div class="stat-value">${total}</div><div class="stat-sub">tickets</div></div>
    <div class="stat"><div class="stat-label">Still Unassigned</div><div class="stat-value red">${unassigned}</div><div class="stat-sub">right now</div></div>
    <div class="stat"><div class="stat-label">Eventually Assigned</div><div class="stat-value yellow">${assigned}</div><div class="stat-sub">after threshold</div></div>
    <div class="stat"><div class="stat-label">Avg Wait</div><div class="stat-value">${avgWait}</div><div class="stat-sub">minutes</div></div>
    ${worst ? '<div class="stat"><div class="stat-label">Longest Wait</div><div class="stat-value red">' + worst.GapMinutes + '</div><div class="stat-sub">minutes</div></div>' : ''}
  `;

  document.getElementById('countLabel').textContent = total + ' records';

  const tbody = document.getElementById('tbody');
  if (total === 0) {
    tbody.innerHTML = '';
    document.getElementById('emptyMsg').style.display = 'block';
    return;
  }
  document.getElementById('emptyMsg').style.display = 'none';

  tbody.innerHTML = rows.map(r => `<tr>
    <td><strong>#${r.TicketId}</strong></td>
    <td>${r.Subject}</td>
    <td>${r.OrgName}</td>
    <td><span class="${priClass(r.Priority)}">${r.Priority}</span></td>
    <td>${r.Status}</td>
    <td>${r.CreatedAt}</td>
    <td>${waitBadge(r.GapMinutes, r.StillUnassigned)}</td>
    <td>${r.AssignedTo || '&mdash;'}</td>
    <td>${r.AssignedBy || '&mdash;'}</td>
  </tr>`).join('');
}

render();
</script>
</body>
</html>
"@

# Save the report
$Timestamp  = Get-Date -Format 'yyyyMMdd_HHmmss'
$ScriptDir  = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.Path }
$OutputPath = Join-Path $ScriptDir "NinjaUnassignedTicketReport_$Timestamp.html"

try {
    [System.IO.File]::WriteAllText($OutputPath, $Html, [System.Text.Encoding]::UTF8)
} catch {
    $OutputPath = Join-Path $env:TEMP "NinjaUnassignedTicketReport_$Timestamp.html"
    [System.IO.File]::WriteAllText($OutputPath, $Html, [System.Text.Encoding]::UTF8)
}

Write-Host ''
Write-Host '  ================================================================' -ForegroundColor Green
Write-Host '  [OK] REPORT COMPLETE' -ForegroundColor Green
Write-Host "       Tickets analysed  : $Total" -ForegroundColor Green
Write-Host "       Breached threshold: $($ReportRows.Count)" -ForegroundColor Green
Write-Host "       Still unassigned  : $StillUnassigned" -ForegroundColor Green
Write-Host "       Avg wait time     : $AvgWait min" -ForegroundColor Green
if ($WorstGap) {
Write-Host "       Longest wait      : $($WorstGap.GapMinutes) min -- $($WorstGap.Subject)" -ForegroundColor Green
}
Write-Host "       Report saved to   : $OutputPath" -ForegroundColor Green
Write-Host '  ================================================================' -ForegroundColor Green
Write-Host ''

try { Start-Process $OutputPath } catch {}
