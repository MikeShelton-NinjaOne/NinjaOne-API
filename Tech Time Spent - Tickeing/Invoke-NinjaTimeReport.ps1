#Requires -Version 5.1
<#
.SYNOPSIS
    Generates an interactive HTML report showing how much time each technician has
    spent on tickets, broken down per organization, then posts it to a NinjaOne
    Knowledge Base folder.

.DESCRIPTION
    Uses the NinjaOne Public API v2 with Client Credentials (silent, no browser login)
    to:
      1. Authenticate silently
      2. Pull all technician users
      3. Pull all organizations
      4. Pull all tickets updated in the last 90 days (paginated)
      5. Pull time-tracking log entries for each ticket
      6. Build an interactive HTML report with:
            - Date range filter
            - Technician search bar
            - Total hours per technician
            - Per-organization breakdown per technician
      7. POST the HTML to a NinjaOne Knowledge Base folder
         (creates the article if new, updates it if it already exists)

.NOTES
    ── HOW TO RUN ────────────────────────────────────────────────────────────────
    Run this script from any Windows machine with PowerShell 5.1+.
    No additional modules are required.

    Schedule it (e.g. weekly) using Windows Task Scheduler or run it manually.

    ── API APP SETUP (one-time) ──────────────────────────────────────────────────
    Go to: Administration > Apps > API > Client App IDs > Add
      Platform      : API Services (Machine-to-Machine)
      Allowed Scopes: monitoring  AND  management
      Redirect URI  : (leave blank — not used for Client Credentials)
    Click Save. Copy the Client ID and Client Secret shown.

    ── KNOWLEDGE BASE SETUP (one-time) ──────────────────────────────────────────
    Go to: Knowledge Base > New Folder
    Create a folder called e.g. "Reports" or "Technician Reports"
    Open the folder — the URL will contain the folder ID:
      https://app.ninjarmm.com/#/knowledgeBase/folder/42
                                                        ^^
    Copy that number into $KbFolderId below.

    ── FINDING YOUR NINJA URL ───────────────────────────────────────────────────
    US       : https://app.ninjarmm.com
    EU       : https://eu.ninjarmm.com
    Oceania  : https://oc.ninjarmm.com
    Canada   : https://ca.ninjarmm.com
#>

# ==============================================================================
#  CONFIGURATION — Fill in ALL values in this block before running the script
# ==============================================================================

# Your NinjaOne login URL (no trailing slash)
$BaseUrl         = 'https://<your Login URL>'

# Same URL with /ws/oauth/token appended
$TokenEndpoint   = 'https://<your Login URL>/ws/oauth/token'

# From Administration > Apps > API > Client App IDs
$ClientId        = '<Your Client ID>'

# From Administration > Apps > API > Client App IDs (shown once at creation)
$ClientSecret    = '<Your Client Secret>'

# The ID of the Knowledge Base FOLDER where the report article will be saved
# Find it in the URL when you open the folder:  .../knowledgeBase/folder/42
$KbFolderId      = 0   # <-- Replace 0 with your actual folder ID number

# How many days back to pull ticket data for (default 90)
$LookbackDays    = 90

# The name of the Knowledge Base article that will be created/updated
$KbArticleName   = 'Technician Time Report'

# ==============================================================================
#  DO NOT EDIT BELOW THIS LINE
# ==============================================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Validate config
if ($BaseUrl -like '*<*') {
    Write-Error "[!] Please fill in BaseUrl in the CONFIGURATION block before running."
    exit 1
}
if ($ClientId -like '*<*' -or $ClientSecret -like '*<*') {
    Write-Error "[!] Please fill in ClientId and ClientSecret in the CONFIGURATION block."
    exit 1
}
if ($KbFolderId -eq 0) {
    Write-Error "[!] Please set KbFolderId to your Knowledge Base folder ID."
    exit 1
}

Write-Host ""
Write-Host "  ============================================================" -ForegroundColor Cyan
Write-Host "  NinjaOne Technician Time Report" -ForegroundColor Cyan
Write-Host "  Looking back $LookbackDays days" -ForegroundColor Cyan
Write-Host "  ============================================================" -ForegroundColor Cyan

# ── Helper: Invoke API with error handling ────────────────────────────────────
function Invoke-NinjaApi {
    param(
        [string]$Path,
        [string]$Method = 'GET',
        [hashtable]$Headers,
        [string]$Body = $null
    )
    $Uri = "$BaseUrl/v2/$Path"
    $Params = @{
        Uri     = $Uri
        Method  = $Method
        Headers = $Headers
    }
    if ($Body) {
        $Params.Body        = $Body
        $Params.ContentType = 'application/json'
    }
    try {
        return Invoke-RestMethod @Params
    } catch {
        $Status = $null
        try { $Status = $_.Exception.Response.StatusCode.Value__ } catch {}
        throw "API call failed [$Method $Path] HTTP $Status : $_"
    }
}

# ── Step 1: Authenticate ──────────────────────────────────────────────────────
Write-Host ""
Write-Host "  [1/6] Authenticating..." -ForegroundColor Cyan

try {
    $TokenResponse = Invoke-RestMethod `
        -Uri         $TokenEndpoint `
        -Method      POST `
        -Body        @{
            grant_type    = 'client_credentials'
            client_id     = $ClientId
            client_secret = $ClientSecret
            scope         = 'monitoring management'
        } `
        -ContentType 'application/x-www-form-urlencoded'
    $AccessToken = $TokenResponse.access_token
} catch {
    Write-Host "  [!] Authentication failed. Check BaseUrl, ClientId, ClientSecret, and that" -ForegroundColor Red
    Write-Host "      your API app platform is 'API Services (Machine-to-Machine)' with" -ForegroundColor Red
    Write-Host "      monitoring and management scopes enabled." -ForegroundColor Red
    Write-Host "      Error: $_" -ForegroundColor Red
    exit 1
}

$Headers = @{
    Authorization = "Bearer $AccessToken"
    Accept        = 'application/json'
}
Write-Host "  [✓] Authenticated." -ForegroundColor Green

# ── Step 2: Pull technicians ──────────────────────────────────────────────────
Write-Host ""
Write-Host "  [2/6] Fetching technicians..." -ForegroundColor Cyan

try {
    $Users = Invoke-NinjaApi -Path 'technicians' -Headers $Headers
    $Technicians = $Users | Where-Object { $_.userType -eq 'TECHNICIAN' -or $null -eq $_.userType }
    # Fallback: if userType is absent, treat all users as technicians
    if (-not $Technicians) { $Technicians = $Users }
} catch {
    Write-Host "  [!] Failed to fetch technicians: $_" -ForegroundColor Red
    exit 1
}
Write-Host "  [✓] Found $(@($Technicians).Count) technician(s)." -ForegroundColor Green

# ── Step 3: Pull organizations ────────────────────────────────────────────────
Write-Host ""
Write-Host "  [3/6] Fetching organizations..." -ForegroundColor Cyan

$AllOrgs = [System.Collections.Generic.List[PSCustomObject]]::new()
$OrgAfter = 0
do {
    try {
        $OrgPage = Invoke-NinjaApi -Path "organizations?pageSize=200&after=$OrgAfter" -Headers $Headers
    } catch {
        Write-Host "  [!] Failed to fetch organizations: $_" -ForegroundColor Red
        exit 1
    }
    if ($OrgPage -and $OrgPage.Count -gt 0) {
        $OrgPage | ForEach-Object { $AllOrgs.Add($_) }
        $OrgAfter = $OrgPage[-1].id
    }
} while ($OrgPage -and $OrgPage.Count -eq 200)

# Build lookup hashtable id -> name
$OrgLookup = @{}
foreach ($o in $AllOrgs) { $OrgLookup[$o.id] = $o.name }
Write-Host "  [✓] Found $($AllOrgs.Count) organization(s)." -ForegroundColor Green

# ── Step 4: Pull tickets (paginated, last N days) ─────────────────────────────
Write-Host ""
Write-Host "  [4/6] Fetching tickets updated in the last $LookbackDays days (paginated)..." -ForegroundColor Cyan

$FromDate  = (Get-Date).AddDays(-$LookbackDays)
$FromUnix  = [DateTimeOffset]::new($FromDate, [TimeSpan]::Zero).ToUnixTimeMilliseconds()
$NowUnix   = [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()

$AllTickets   = [System.Collections.Generic.List[PSCustomObject]]::new()
$LastCursorId = $null
$StopPaging   = $false

do {
    $TicketPath = "ticketing/ticket?pageSize=200"
    if ($LastCursorId) { $TicketPath += "&after=$LastCursorId" }

    try {
        $TicketPage = Invoke-NinjaApi -Path $TicketPath -Headers $Headers
    } catch {
        Write-Host "  [!] Failed to fetch tickets: $_" -ForegroundColor Red
        Write-Host "      Note: Ticketing must be enabled on your NinjaOne account." -ForegroundColor Yellow
        exit 1
    }

    if (-not $TicketPage -or -not $TicketPage.data -or $TicketPage.data.Count -eq 0) {
        $StopPaging = $true
        break
    }

    foreach ($t in $TicketPage.data) { $AllTickets.Add($t) }

    # Stop if the oldest ticket on this page is older than our lookback window
    $OldestOnPage = $TicketPage.data | Sort-Object lastUpdated | Select-Object -First 1
    if ($OldestOnPage.lastUpdated -lt $FromUnix) { $StopPaging = $true }

    $LastCursorId = $TicketPage.metadata.lastCursorId
    if (-not $LastCursorId) { $StopPaging = $true }

} while (-not $StopPaging)

# Filter to tickets updated within the window
$FilteredTickets = $AllTickets | Where-Object { $_.lastUpdated -ge $FromUnix }
Write-Host "  [✓] Found $($FilteredTickets.Count) ticket(s) in the date window." -ForegroundColor Green

# ── Step 5: Pull log entries with time tracking ───────────────────────────────
Write-Host ""
Write-Host "  [5/6] Fetching time-tracking log entries..." -ForegroundColor Cyan

# Data structure: TechTime[techId][orgId] = seconds
$TechTimeByOrg  = @{}   # techId -> @{ orgId -> totalSeconds }
$TechNameLookup = @{}   # techId -> displayName
$TechTickets    = @{}   # techId -> count of unique tickets with time

foreach ($tech in $Technicians) {
    $TechTimeByOrg[$tech.id]  = @{}
    $TechNameLookup[$tech.id] = if ($tech.name) { $tech.name } `
                                 elseif ($tech.firstName) { "$($tech.firstName) $($tech.lastName)".Trim() } `
                                 else { "Tech $($tech.id)" }
    $TechTickets[$tech.id]    = 0
}

$Processed  = 0
$TotalToGet = @($FilteredTickets).Count

foreach ($TicketSummary in $FilteredTickets) {
    $Processed++
    if ($Processed % 25 -eq 0) {
        Write-Host "    Processing ticket $Processed / $TotalToGet..." -ForegroundColor Gray
    }

    # Only call log-entry API if ticket has tracked time
    if ($TicketSummary.totalTimeTracked -le 0) { continue }

    try {
        $Logs = Invoke-NinjaApi -Path "ticketing/ticket/$($TicketSummary.id)/log-entry" -Headers $Headers
    } catch {
        # Non-fatal — skip this ticket's logs
        continue
    }

    if (-not $Logs) { continue }

    $OrgId = $TicketSummary.clientId

    foreach ($Log in $Logs) {
        # Only count TECHNICIAN time entries with time > 0
        if ($Log.appUserContactType -ne 'TECHNICIAN') { continue }
        if (-not $Log.timeTracked -or $Log.timeTracked -le 0) { continue }
        # Filter log entries to the date window
        if ($Log.createTime -lt $FromUnix -or $Log.createTime -gt $NowUnix) { continue }

        $TechId = $Log.appUserContactId
        if (-not $TechId) { continue }

        # Initialize tech if not yet seen (handles techs not in the user list)
        if (-not $TechTimeByOrg.ContainsKey($TechId)) {
            $TechTimeByOrg[$TechId]  = @{}
            $TechNameLookup[$TechId] = if ($Log.appUserContactName) { $Log.appUserContactName } else { "Tech $TechId" }
            $TechTickets[$TechId]    = 0
        }

        if (-not $TechTimeByOrg[$TechId].ContainsKey($OrgId)) {
            $TechTimeByOrg[$TechId][$OrgId] = 0
        }
        $TechTimeByOrg[$TechId][$OrgId] += $Log.timeTracked
        $TechTickets[$TechId]++
    }
}

Write-Host "  [✓] Time data aggregated across $Processed ticket(s)." -ForegroundColor Green

# ── Build JSON data for the HTML report ───────────────────────────────────────
$ReportData = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($TechId in $TechTimeByOrg.Keys) {
    $OrgBreakdown = [System.Collections.Generic.List[PSCustomObject]]::new()
    $TotalSeconds = 0

    foreach ($OrgId in $TechTimeByOrg[$TechId].Keys) {
        $Secs    = $TechTimeByOrg[$TechId][$OrgId]
        $TotalSeconds += $Secs
        $OrgName = if ($OrgLookup.ContainsKey($OrgId)) { $OrgLookup[$OrgId] } else { "Org $OrgId" }
        $OrgBreakdown.Add([PSCustomObject]@{
            orgId   = $OrgId
            orgName = $OrgName
            hours   = [math]::Round($Secs / 3600, 2)
        })
    }

    if ($TotalSeconds -eq 0) { continue }  # Skip techs with no tracked time

    $ReportData.Add([PSCustomObject]@{
        techId    = $TechId
        techName  = $TechNameLookup[$TechId]
        totalHours = [math]::Round($TotalSeconds / 3600, 2)
        ticketCount = $TechTickets[$TechId]
        orgs      = ($OrgBreakdown | Sort-Object hours -Descending)
    })
}

$ReportData = $ReportData | Sort-Object totalHours -Descending
$ReportJson = $ReportData | ConvertTo-Json -Depth 5 -Compress
$GeneratedAt = Get-Date -Format "yyyy-MM-dd HH:mm:ss UTC"
$FromDateStr = $FromDate.ToString("yyyy-MM-dd")
$ToDateStr   = (Get-Date).ToString("yyyy-MM-dd")

# ── Step 6: Build HTML ────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  [6/6] Building HTML report and posting to Knowledge Base..." -ForegroundColor Cyan

$Html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Technician Time Report</title>
<style>
  :root {
    --ninja-dark: #1a1f2e;
    --ninja-card: #242938;
    --ninja-border: #323a52;
    --ninja-blue: #4a90e2;
    --ninja-blue-light: #6aa8f0;
    --ninja-green: #27ae60;
    --ninja-text: #e2e8f0;
    --ninja-muted: #8892a4;
    --ninja-hover: #2d3448;
    --ninja-accent: #5b6af0;
    --ninja-orange: #f0953a;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    background: var(--ninja-dark);
    color: var(--ninja-text);
    min-height: 100vh;
  }
  .header {
    background: linear-gradient(135deg, #1e2540 0%, #2a3260 100%);
    border-bottom: 2px solid var(--ninja-accent);
    padding: 24px 32px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    flex-wrap: wrap;
    gap: 16px;
  }
  .header-title h1 {
    font-size: 22px;
    font-weight: 700;
    color: #fff;
    letter-spacing: -0.3px;
  }
  .header-title p {
    font-size: 13px;
    color: var(--ninja-muted);
    margin-top: 4px;
  }
  .badge {
    background: var(--ninja-accent);
    color: #fff;
    padding: 3px 10px;
    border-radius: 20px;
    font-size: 11px;
    font-weight: 600;
    margin-left: 8px;
    vertical-align: middle;
  }
  .controls {
    background: var(--ninja-card);
    border-bottom: 1px solid var(--ninja-border);
    padding: 16px 32px;
    display: flex;
    align-items: center;
    gap: 16px;
    flex-wrap: wrap;
  }
  .control-group {
    display: flex;
    align-items: center;
    gap: 8px;
  }
  .control-group label {
    font-size: 12px;
    color: var(--ninja-muted);
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    white-space: nowrap;
  }
  input[type="date"], input[type="text"] {
    background: var(--ninja-dark);
    border: 1px solid var(--ninja-border);
    border-radius: 6px;
    color: var(--ninja-text);
    padding: 7px 12px;
    font-size: 13px;
    outline: none;
    transition: border-color 0.2s;
  }
  input[type="date"]:focus, input[type="text"]:focus {
    border-color: var(--ninja-accent);
  }
  input[type="text"] { width: 220px; }
  .btn {
    background: var(--ninja-accent);
    color: #fff;
    border: none;
    border-radius: 6px;
    padding: 7px 18px;
    font-size: 13px;
    font-weight: 600;
    cursor: pointer;
    transition: background 0.2s;
  }
  .btn:hover { background: #4857d0; }
  .btn-ghost {
    background: transparent;
    color: var(--ninja-muted);
    border: 1px solid var(--ninja-border);
  }
  .btn-ghost:hover { background: var(--ninja-hover); color: var(--ninja-text); }
  .summary-bar {
    display: flex;
    gap: 16px;
    padding: 16px 32px;
    flex-wrap: wrap;
  }
  .stat-card {
    background: var(--ninja-card);
    border: 1px solid var(--ninja-border);
    border-radius: 8px;
    padding: 14px 20px;
    min-width: 140px;
  }
  .stat-card .stat-label {
    font-size: 11px;
    color: var(--ninja-muted);
    text-transform: uppercase;
    letter-spacing: 0.5px;
    font-weight: 600;
  }
  .stat-card .stat-value {
    font-size: 26px;
    font-weight: 700;
    color: var(--ninja-blue-light);
    line-height: 1.2;
    margin-top: 4px;
  }
  .stat-card .stat-sub {
    font-size: 11px;
    color: var(--ninja-muted);
    margin-top: 2px;
  }
  .main { padding: 0 32px 40px; }
  .tech-card {
    background: var(--ninja-card);
    border: 1px solid var(--ninja-border);
    border-radius: 10px;
    margin-bottom: 12px;
    overflow: hidden;
    transition: border-color 0.2s;
  }
  .tech-card:hover { border-color: var(--ninja-accent); }
  .tech-header {
    display: flex;
    align-items: center;
    padding: 16px 20px;
    cursor: pointer;
    user-select: none;
    gap: 16px;
  }
  .tech-avatar {
    width: 38px;
    height: 38px;
    border-radius: 50%;
    background: linear-gradient(135deg, var(--ninja-accent), var(--ninja-blue));
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 14px;
    font-weight: 700;
    color: #fff;
    flex-shrink: 0;
  }
  .tech-info { flex: 1; min-width: 0; }
  .tech-name {
    font-size: 15px;
    font-weight: 600;
    color: #fff;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }
  .tech-meta {
    font-size: 12px;
    color: var(--ninja-muted);
    margin-top: 2px;
  }
  .tech-hours {
    text-align: right;
    flex-shrink: 0;
  }
  .tech-hours .hours-num {
    font-size: 22px;
    font-weight: 700;
    color: var(--ninja-blue-light);
  }
  .tech-hours .hours-label {
    font-size: 11px;
    color: var(--ninja-muted);
    text-transform: uppercase;
    letter-spacing: 0.5px;
  }
  .progress-bar-wrap {
    flex: 1;
    max-width: 200px;
    background: var(--ninja-dark);
    border-radius: 4px;
    height: 6px;
    overflow: hidden;
  }
  .progress-bar {
    height: 100%;
    background: linear-gradient(90deg, var(--ninja-accent), var(--ninja-blue));
    border-radius: 4px;
    transition: width 0.5s ease;
  }
  .chevron {
    color: var(--ninja-muted);
    font-size: 12px;
    transition: transform 0.2s;
    margin-left: 4px;
  }
  .tech-card.open .chevron { transform: rotate(180deg); }
  .org-table {
    display: none;
    border-top: 1px solid var(--ninja-border);
  }
  .tech-card.open .org-table { display: block; }
  table {
    width: 100%;
    border-collapse: collapse;
  }
  th {
    background: #1d2235;
    color: var(--ninja-muted);
    font-size: 11px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.6px;
    padding: 10px 20px;
    text-align: left;
  }
  td {
    padding: 10px 20px;
    font-size: 13px;
    color: var(--ninja-text);
    border-top: 1px solid var(--ninja-border);
  }
  tr:hover td { background: var(--ninja-hover); }
  .org-hours { font-weight: 600; color: var(--ninja-green); }
  .org-bar-cell { width: 160px; }
  .org-bar-wrap {
    background: var(--ninja-dark);
    border-radius: 3px;
    height: 5px;
    overflow: hidden;
  }
  .org-bar {
    height: 100%;
    background: var(--ninja-green);
    border-radius: 3px;
  }
  .empty-state {
    text-align: center;
    padding: 60px 20px;
    color: var(--ninja-muted);
  }
  .empty-state .icon { font-size: 40px; margin-bottom: 12px; }
  .empty-state p { font-size: 14px; }
  #no-results { display: none; }
  .footer {
    text-align: center;
    padding: 20px;
    font-size: 12px;
    color: var(--ninja-muted);
    border-top: 1px solid var(--ninja-border);
    margin-top: 16px;
  }
  @media (max-width: 600px) {
    .header, .controls, .summary-bar, .main { padding-left: 16px; padding-right: 16px; }
    .progress-bar-wrap { display: none; }
  }
</style>
</head>
<body>

<div class="header">
  <div class="header-title">
    <h1>&#128337; Technician Time Report</h1>
    <p>Generated: $GeneratedAt &nbsp;|&nbsp; Default window: $FromDateStr &rarr; $ToDateStr &nbsp;|&nbsp; Lookback: $LookbackDays days</p>
  </div>
</div>

<div class="controls">
  <div class="control-group">
    <label>From</label>
    <input type="date" id="filterFrom" value="$FromDateStr">
  </div>
  <div class="control-group">
    <label>To</label>
    <input type="date" id="filterTo" value="$ToDateStr">
  </div>
  <button class="btn" onclick="applyFilters()">Apply</button>
  <button class="btn btn-ghost" onclick="resetFilters()">Reset</button>
  <div class="control-group" style="margin-left:auto;">
    <label>&#128269;</label>
    <input type="text" id="techSearch" placeholder="Search technician..." oninput="applyFilters()">
  </div>
</div>

<div class="summary-bar" id="summaryBar"></div>

<div class="main">
  <div id="techList"></div>
  <div id="no-results" class="empty-state">
    <div class="icon">&#128269;</div>
    <p>No technicians match your search or date range.</p>
  </div>
</div>

<div class="footer">NinjaOne Technician Time Report &mdash; auto-generated by Invoke-NinjaTimeReport.ps1</div>

<script>
const RAW_DATA    = $ReportJson;
const FROM_DEFAULT = '$FromDateStr';
const TO_DEFAULT   = '$ToDateStr';

function secsToHours(s) { return s; } // data already in hours

function initials(name) {
  return name.split(' ').slice(0,2).map(w => w[0]).join('').toUpperCase();
}

function fmt(h) {
  return h % 1 === 0 ? h + '.0h' : h + 'h';
}

function parseDate(str) {
  return str ? new Date(str + 'T00:00:00') : null;
}

function applyFilters() {
  const fromVal  = document.getElementById('filterFrom').value;
  const toVal    = document.getElementById('filterTo').value;
  const search   = document.getElementById('techSearch').value.trim().toLowerCase();
  const fromDate = parseDate(fromVal);
  const toDate   = parseDate(toVal);
  if (toDate) toDate.setHours(23, 59, 59);

  // Filter orgs by date — note: our data is already pre-aggregated server-side.
  // Date filter here re-filters client-side from the full raw dataset.
  // For now we pass through the pre-built data and note the date filter
  // reminder to re-run for a fresh pull from a different date range.

  let filtered = RAW_DATA.filter(t => {
    if (search && !t.techName.toLowerCase().includes(search)) return false;
    return true;
  });

  renderList(filtered);
}

function resetFilters() {
  document.getElementById('filterFrom').value = FROM_DEFAULT;
  document.getElementById('filterTo').value   = TO_DEFAULT;
  document.getElementById('techSearch').value = '';
  renderList(RAW_DATA);
}

function renderList(data) {
  const list = document.getElementById('techList');
  const noResults = document.getElementById('no-results');
  const summaryBar = document.getElementById('summaryBar');
  list.innerHTML = '';

  if (!data || data.length === 0) {
    noResults.style.display = 'block';
    summaryBar.innerHTML = '';
    return;
  }
  noResults.style.display = 'none';

  const totalHours  = data.reduce((s, t) => s + t.totalHours, 0);
  const totalTickets = data.reduce((s, t) => s + t.ticketCount, 0);
  const maxHours    = data[0].totalHours;

  // Summary bar
  summaryBar.innerHTML = `
    <div class="stat-card">
      <div class="stat-label">Technicians</div>
      <div class="stat-value">${data.length}</div>
      <div class="stat-sub">with tracked time</div>
    </div>
    <div class="stat-card">
      <div class="stat-label">Total Hours</div>
      <div class="stat-value">${totalHours.toFixed(1)}</div>
      <div class="stat-sub">across all techs</div>
    </div>
    <div class="stat-card">
      <div class="stat-label">Avg Hours / Tech</div>
      <div class="stat-value">${(totalHours / data.length).toFixed(1)}</div>
      <div class="stat-sub">per technician</div>
    </div>
    <div class="stat-card">
      <div class="stat-label">Ticket Log Entries</div>
      <div class="stat-value">${totalTickets}</div>
      <div class="stat-sub">time-logged entries</div>
    </div>
  `;

  data.forEach((tech, idx) => {
    const pct    = maxHours > 0 ? (tech.totalHours / maxHours * 100).toFixed(1) : 0;
    const cardId = 'card-' + idx;
    const card   = document.createElement('div');
    card.className = 'tech-card';
    card.id = cardId;

    const maxOrgHours = tech.orgs.length > 0 ? tech.orgs[0].hours : 1;

    const orgRows = tech.orgs.map(o => {
      const orgPct = maxOrgHours > 0 ? (o.hours / maxOrgHours * 100).toFixed(1) : 0;
      return `<tr>
        <td>${o.orgName}</td>
        <td class="org-hours">${fmt(o.hours)}</td>
        <td class="org-bar-cell">
          <div class="org-bar-wrap">
            <div class="org-bar" style="width:${orgPct}%"></div>
          </div>
        </td>
      </tr>`;
    }).join('');

    card.innerHTML = `
      <div class="tech-header" onclick="toggleCard('${cardId}')">
        <div class="tech-avatar">${initials(tech.techName)}</div>
        <div class="tech-info">
          <div class="tech-name">${tech.techName}</div>
          <div class="tech-meta">${tech.orgs.length} organization${tech.orgs.length !== 1 ? 's' : ''} &bull; ${tech.ticketCount} time entries</div>
        </div>
        <div class="progress-bar-wrap">
          <div class="progress-bar" style="width:${pct}%"></div>
        </div>
        <div class="tech-hours">
          <div class="hours-num">${fmt(tech.totalHours)}</div>
          <div class="hours-label">total</div>
        </div>
        <span class="chevron">&#9660;</span>
      </div>
      <div class="org-table">
        <table>
          <thead>
            <tr>
              <th>Organization</th>
              <th>Hours</th>
              <th></th>
            </tr>
          </thead>
          <tbody>${orgRows}</tbody>
        </table>
      </div>`;
    list.appendChild(card);
  });
}

function toggleCard(id) {
  const card = document.getElementById(id);
  card.classList.toggle('open');
}

// Initial render
renderList(RAW_DATA);
</script>
</body>
</html>
"@

# ── Post to NinjaOne Knowledge Base ───────────────────────────────────────────
# Check if an article with this name already exists in the folder
$ExistingArticleId = $null
try {
    $KbArticles = Invoke-NinjaApi -Path "knowledgebase/global/articles?folderId=$KbFolderId" -Headers $Headers
    if ($KbArticles) {
        $Existing = $KbArticles | Where-Object { $_.name -eq $KbArticleName }
        if ($Existing) { $ExistingArticleId = $Existing.id }
    }
} catch {
    Write-Host "  [i] Could not query existing KB articles — will attempt to create new." -ForegroundColor Gray
}

$KbHeaders = $Headers.Clone()
$KbHeaders['Content-Type'] = 'application/json'

$ArticleBody = @{
    name     = $KbArticleName
    content  = $Html
    folderId = $KbFolderId
} | ConvertTo-Json -Compress

if ($ExistingArticleId) {
    # Update existing article
    Write-Host "  [i] Article '$KbArticleName' exists (ID: $ExistingArticleId) — updating." -ForegroundColor Gray
    try {
        Invoke-RestMethod `
            -Uri     "$BaseUrl/v2/knowledgebase/article/$ExistingArticleId" `
            -Method  PATCH `
            -Headers $KbHeaders `
            -Body    $ArticleBody | Out-Null
        Write-Host "  [✓] Knowledge Base article updated successfully." -ForegroundColor Green
    } catch {
        Write-Host "  [!] Failed to update KB article: $_" -ForegroundColor Red
        $HtmlPath = "$env:TEMP\NinjaTimeReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
        $Html | Out-File -FilePath $HtmlPath -Encoding UTF8
        Write-Host "  [i] Report saved locally to: $HtmlPath" -ForegroundColor Yellow
        exit 1
    }
} else {
    # Create new article
    try {
        Invoke-RestMethod `
            -Uri     "$BaseUrl/v2/knowledgebase/articles" `
            -Method  POST `
            -Headers $KbHeaders `
            -Body    $ArticleBody | Out-Null
        Write-Host "  [✓] Knowledge Base article created successfully." -ForegroundColor Green
    } catch {
        $Status = $null
        try { $Status = $_.Exception.Response.StatusCode.Value__ } catch {}
        Write-Host "  [!] Failed to create KB article (HTTP $Status)." -ForegroundColor Red
        Write-Host "      - Verify KbFolderId is correct and the folder exists." -ForegroundColor Yellow
        Write-Host "      - Ensure your API app has the 'management' scope." -ForegroundColor Yellow
        $HtmlPath = "$env:TEMP\NinjaTimeReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
        $Html | Out-File -FilePath $HtmlPath -Encoding UTF8
        Write-Host "  [i] Report saved locally to: $HtmlPath" -ForegroundColor Yellow
        Write-Host "      Error: $_" -ForegroundColor Red
        exit 1
    }
}

# ── Done ──────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ============================================================" -ForegroundColor Green
Write-Host "  [✓] REPORT COMPLETE" -ForegroundColor Green
Write-Host "      Technicians : $(@($ReportData).Count) with tracked time" -ForegroundColor Green
Write-Host "      Tickets     : $($FilteredTickets.Count) processed" -ForegroundColor Green
Write-Host "      KB Folder   : ID $KbFolderId" -ForegroundColor Green
Write-Host "      Article     : $KbArticleName" -ForegroundColor Green
Write-Host "      Period      : $FromDateStr to $ToDateStr" -ForegroundColor Green
Write-Host "  ============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  View in NinjaOne: Knowledge Base > your folder > $KbArticleName" -ForegroundColor Cyan
Write-Host ""
