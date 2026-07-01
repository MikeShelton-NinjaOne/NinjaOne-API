#Requires -Version 5.1
<#
.SYNOPSIS
    Pulls software patch data across ALL organizations and ALL devices, cross-references
    installed software versions, and posts an interactive filterable HTML report to a
    NinjaOne Knowledge Base folder.

.DESCRIPTION
    Uses the NinjaOne Public API v2 with Client Credentials (silent, no browser login).

    Data sources:
      GET /v2/queries/software-patches   — patch records (status, impact, installedAt, productIdentifier)
      GET /v2/queries/software           — installed software per device (name, version, publisher)
      GET /v2/organizations              — org names for display
      GET /v2/devices                    — device names and org mapping

    Version enrichment: software-patches does not return a version field.
    This script cross-references the installed software inventory by matching
    productIdentifier / name to attach the currently installed version to each patch record.

    Report features (all client-side, no server round-trips):
      - Filter by status (Installed / Pending / Failed / Rejected)
      - Filter by severity / impact (Critical / Important / Moderate / Low)
      - Filter by organization
      - Search bar (patch name or device name)
      - Summary stat cards at the top
      - Sortable colour-coded table
      - Posts to NinjaOne KB — creates article on first run, updates on subsequent runs

.NOTES
    ── API APP SETUP (one-time) ──────────────────────────────────────────────────
    Go to: Administration > Apps > API > Client App IDs > Add
      Platform      : API Services (Machine-to-Machine)
      Allowed Scopes: monitoring  AND  management
      Redirect URI  : Leave blank
    Click Save. Copy the Client ID and Client Secret.

    ── KB FOLDER ID ─────────────────────────────────────────────────────────────
    Open the target KB folder in NinjaOne. Folder ID is in the URL:
      https://app.ninjarmm.com/#/knowledgeBase/folder/42
                                                       ^^
    ── REGIONAL URLS ────────────────────────────────────────────────────────────
    US: https://app.ninjarmm.com  |  EU: https://eu.ninjarmm.com
    OC: https://oc.ninjarmm.com   |  CA: https://ca.ninjarmm.com
#>

# ==============================================================================
#  CONFIGURATION — Fill in ALL five values before running
# ==============================================================================

$BaseUrl       = 'https://<your Login URL>'
$TokenEndpoint = 'https://<your Login URL>/ws/oauth/token'
$ClientId      = '<Your Client ID>'
$ClientSecret  = '<Your Client Secret>'
$KbFolderId    = 0    # <-- Replace with your KB folder ID number

# ==============================================================================
#  DO NOT EDIT BELOW THIS LINE
# ==============================================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ($BaseUrl -like '*<*')                            { Write-Error 'Fill in BaseUrl';       exit 1 }
if ($ClientId -like '*<*' -or $ClientSecret -like '*<*') { Write-Error 'Fill in credentials'; exit 1 }
if ($KbFolderId -eq 0)                               { Write-Error 'Fill in KbFolderId';   exit 1 }

$ArticleName = 'Software Patch Report — All Organizations'

Write-Host ""
Write-Host "  ============================================================" -ForegroundColor Cyan
Write-Host "  NinjaOne Software Patch Report  [All Organizations]" -ForegroundColor Cyan
Write-Host "  ============================================================" -ForegroundColor Cyan

# ── Auth ──────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  [1/6] Authenticating..." -ForegroundColor Cyan
try {
    $Token = Invoke-RestMethod -Uri $TokenEndpoint -Method POST -ContentType 'application/x-www-form-urlencoded' -Body @{
        grant_type = 'client_credentials'; client_id = $ClientId
        client_secret = $ClientSecret; scope = 'monitoring management'
    }
    $H = @{ Authorization = "Bearer $($Token.access_token)"; Accept = 'application/json'; 'Content-Type' = 'application/json' }
} catch { Write-Host "  [!] Auth failed: $_" -ForegroundColor Red; exit 1 }
Write-Host "  [✓] Authenticated." -ForegroundColor Green

function Get-NinjaPagedQuery {
    param([string]$Path, [hashtable]$Headers)
    $all = [System.Collections.Generic.List[PSCustomObject]]::new()
    $cursor = $null
    do {
        $url = "$BaseUrl/v2/$Path"
        if ($cursor) { $url += if ($url -match '\?') { "&after=$cursor" } else { "?after=$cursor" } }
        $page = Invoke-RestMethod -Uri $url -Method GET -Headers $Headers
        $items = if ($page.PSObject.Properties['results']) { $page.results } elseif ($page -is [array]) { $page } else { @($page) }
        foreach ($i in $items) { $all.Add($i) }
        $cursor = if ($page.PSObject.Properties['lastCursor']) { $page.lastCursor } else { $null }
    } while ($cursor -and $items.Count -gt 0)
    return $all
}

# ── Organizations ─────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  [2/6] Fetching organizations..." -ForegroundColor Cyan
$Orgs = Get-NinjaPagedQuery -Path 'organizations?pageSize=200' -Headers $H
$OrgMap = @{}; foreach ($o in $Orgs) { $OrgMap[$o.id] = $o.name }
Write-Host "  [✓] $($Orgs.Count) organizations." -ForegroundColor Green

# ── Devices ───────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  [3/6] Fetching devices..." -ForegroundColor Cyan
$Devices = Get-NinjaPagedQuery -Path 'devices?pageSize=200' -Headers $H
$DevMap = @{}; foreach ($d in $Devices) { $DevMap[$d.id] = $d }
Write-Host "  [✓] $($Devices.Count) devices." -ForegroundColor Green

# ── Software inventory (for version cross-reference) ──────────────────────────
Write-Host ""
Write-Host "  [4/6] Fetching software inventory (version data)..." -ForegroundColor Cyan
$SoftwareRaw = Get-NinjaPagedQuery -Path 'queries/software?pageSize=1000' -Headers $H
# Build lookup: deviceId -> productIdentifier/name -> version
$VersionMap = @{}
foreach ($sw in $SoftwareRaw) {
    $devId = $sw.deviceId
    if (-not $VersionMap.ContainsKey($devId)) { $VersionMap[$devId] = @{} }
    $key = if ($sw.productCode) { $sw.productCode } else { $sw.name }
    if ($key -and $sw.version) { $VersionMap[$devId][$key] = $sw.version }
}
Write-Host "  [✓] $($SoftwareRaw.Count) software records." -ForegroundColor Green

# ── Software patches ──────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  [5/6] Fetching software patch records..." -ForegroundColor Cyan
$Patches = Get-NinjaPagedQuery -Path 'queries/software-patches?pageSize=1000' -Headers $H
Write-Host "  [✓] $($Patches.Count) patch records." -ForegroundColor Green

# ── Build report dataset ──────────────────────────────────────────────────────
$Rows = [System.Collections.Generic.List[PSCustomObject]]::new()
foreach ($p in $Patches) {
    $dev     = if ($DevMap.ContainsKey($p.deviceId)) { $DevMap[$p.deviceId] } else { $null }
    $devName = if ($dev) { $dev.systemName } else { "Device $($p.deviceId)" }
    $orgId   = if ($dev) { $dev.organizationId } else { 0 }
    $orgName = if ($OrgMap.ContainsKey($orgId)) { $OrgMap[$orgId] } else { 'Unknown Org' }

    # Version cross-reference
    $ver = 'N/A'
    if ($VersionMap.ContainsKey($p.deviceId)) {
        $devSw = $VersionMap[$p.deviceId]
        if ($p.identifier -and $devSw.ContainsKey($p.identifier))   { $ver = $devSw[$p.identifier] }
        elseif ($p.name   -and $devSw.ContainsKey($p.name))         { $ver = $devSw[$p.name] }
    }

    $installedAt = if ($p.installedAt) {
        [DateTimeOffset]::FromUnixTimeMilliseconds($p.installedAt).ToLocalTime().ToString('yyyy-MM-dd HH:mm')
    } else { '' }

    $Rows.Add([PSCustomObject]@{
        patchName   = $p.name
        version     = $ver
        status      = $p.status
        impact      = $p.impact
        type        = $p.type
        deviceName  = $devName
        deviceId    = $p.deviceId
        orgName     = $orgName
        identifier  = $p.identifier
        installedAt = $installedAt
    })
}

$JsonData    = $Rows | ConvertTo-Json -Depth 3 -Compress
$GeneratedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
$OrgOptions  = ($OrgMap.Values | Sort-Object | ForEach-Object { "<option value='$_'>$_</option>" }) -join ''

# ── Build HTML ────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  [6/6] Building HTML and posting to KB..." -ForegroundColor Cyan

. "$PSScriptRoot\shared_html_builder.ps1" -ErrorAction SilentlyContinue

$Html = @"
<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<title>Software Patch Report — All Organizations</title>
<style>
:root{--blue:#2E75B6;--dark:#1F4E79;--card:#f0f4fa;--border:#c8d8ee;--text:#1a1a1a;--muted:#6b7a99;
--green:#1a7a3f;--green-bg:#e8f5ee;--red:#b91c1c;--red-bg:#fef2f2;--yellow:#b45309;--yellow-bg:#fffbeb;
--gray:#4b5563;--gray-bg:#f3f4f6;--purple:#7c3aed;--purple-bg:#f5f3ff;}
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#f8faff;color:var(--text);min-height:100vh;}
.header{background:linear-gradient(135deg,var(--dark),var(--blue));padding:22px 32px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px;}
.header h1{color:#fff;font-size:20px;font-weight:700;}
.header p{color:#bdd7ee;font-size:12px;margin-top:4px;}
.controls{background:#fff;border-bottom:1px solid var(--border);padding:14px 32px;display:flex;flex-wrap:wrap;gap:12px;align-items:center;}
select,input[type=text]{background:#f8faff;border:1px solid var(--border);border-radius:6px;padding:7px 11px;font-size:13px;color:var(--text);outline:none;}
select:focus,input:focus{border-color:var(--blue);}
input[type=text]{width:220px;}
.btn{background:var(--blue);color:#fff;border:none;border-radius:6px;padding:7px 16px;font-size:13px;font-weight:600;cursor:pointer;}
.btn:hover{background:var(--dark);}
.btn-ghost{background:transparent;color:var(--muted);border:1px solid var(--border);}
.btn-ghost:hover{background:var(--card);}
.stats{display:flex;gap:14px;padding:16px 32px;flex-wrap:wrap;}
.stat{background:#fff;border:1px solid var(--border);border-radius:8px;padding:14px 20px;min-width:130px;}
.stat-label{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;font-weight:600;}
.stat-value{font-size:26px;font-weight:700;color:var(--blue);line-height:1.2;margin-top:3px;}
.stat-sub{font-size:11px;color:var(--muted);margin-top:2px;}
.main{padding:0 32px 40px;}
.table-wrap{background:#fff;border:1px solid var(--border);border-radius:8px;overflow:hidden;}
table{width:100%;border-collapse:collapse;font-size:13px;}
th{background:var(--dark);color:#fff;padding:10px 14px;text-align:left;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;cursor:pointer;user-select:none;white-space:nowrap;}
th:hover{background:var(--blue);}
td{padding:9px 14px;border-top:1px solid var(--border);vertical-align:middle;}
tr:hover td{background:var(--card);}
.badge{display:inline-block;padding:2px 9px;border-radius:20px;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.4px;}
.badge-INSTALLED{background:var(--green-bg);color:var(--green);}
.badge-PENDING,.badge-APPROVED{background:var(--yellow-bg);color:var(--yellow);}
.badge-FAILED{background:var(--red-bg);color:var(--red);}
.badge-REJECTED{background:var(--gray-bg);color:var(--gray);}
.sev-CRITICAL{background:#fef2f2;color:#b91c1c;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:700;}
.sev-IMPORTANT,.sev-HIGH{background:#fff7ed;color:#b45309;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:700;}
.sev-MODERATE,.sev-MEDIUM{background:#fffbeb;color:#92400e;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:700;}
.sev-LOW{background:#f0fdf4;color:#166534;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:700;}
.sev-OTHER{background:var(--gray-bg);color:var(--gray);padding:2px 8px;border-radius:4px;font-size:11px;}
.empty{text-align:center;padding:60px 20px;color:var(--muted);}
.ver{font-family:'Courier New',monospace;font-size:12px;color:var(--purple);background:var(--purple-bg);padding:1px 6px;border-radius:3px;}
</style></head><body>
<div class="header">
  <div><h1>&#129520; Software Patch Report — All Organizations</h1><p>Generated: $GeneratedAt</p></div>
</div>
<div class="controls">
  <select id="fStatus" onchange="render()"><option value="">All Statuses</option><option>INSTALLED</option><option>PENDING</option><option>FAILED</option><option>REJECTED</option><option>APPROVED</option></select>
  <select id="fImpact" onchange="render()"><option value="">All Severities</option><option>CRITICAL</option><option>IMPORTANT</option><option>MODERATE</option><option>LOW</option></select>
  <select id="fOrg" onchange="render()"><option value="">All Organizations</option>$OrgOptions</select>
  <input type="text" id="fSearch" placeholder="&#128269; Search patch or device..." oninput="render()">
  <button class="btn btn-ghost" onclick="resetFilters()">Reset</button>
  <span id="countLabel" style="font-size:12px;color:var(--muted);margin-left:auto;"></span>
</div>
<div class="stats" id="stats"></div>
<div class="main">
  <div class="table-wrap">
    <table id="patchTable">
      <thead><tr>
        <th onclick="sortBy('patchName')">Patch Name &#8597;</th>
        <th onclick="sortBy('version')">Version &#8597;</th>
        <th onclick="sortBy('status')">Status &#8597;</th>
        <th onclick="sortBy('impact')">Severity &#8597;</th>
        <th onclick="sortBy('type')">Type &#8597;</th>
        <th onclick="sortBy('deviceName')">Device &#8597;</th>
        <th onclick="sortBy('orgName')">Organization &#8597;</th>
        <th onclick="sortBy('installedAt')">Installed At &#8597;</th>
      </tr></thead>
      <tbody id="tbody"></tbody>
    </table>
    <div class="empty" id="emptyMsg" style="display:none;">No records match your filters.</div>
  </div>
</div>
<script>
const DATA = $JsonData;
let sortCol = 'patchName', sortDir = 1;
function sortBy(col) { if(sortCol===col){sortDir*=-1;}else{sortCol=col;sortDir=1;} render(); }
function resetFilters(){
  document.getElementById('fStatus').value='';
  document.getElementById('fImpact').value='';
  document.getElementById('fOrg').value='';
  document.getElementById('fSearch').value='';
  render();
}
function render() {
  const st=document.getElementById('fStatus').value.toUpperCase();
  const im=document.getElementById('fImpact').value.toUpperCase();
  const org=document.getElementById('fOrg').value;
  const q=document.getElementById('fSearch').value.toLowerCase();
  let rows=DATA.filter(r=>{
    if(st && r.status.toUpperCase()!==st) return false;
    if(im && (r.impact||'').toUpperCase()!==im) return false;
    if(org && r.orgName!==org) return false;
    if(q && !r.patchName.toLowerCase().includes(q) && !r.deviceName.toLowerCase().includes(q)) return false;
    return true;
  });
  rows.sort((a,b)=>{
    const av=a[sortCol]||''; const bv=b[sortCol]||'';
    return av<bv?-sortDir:av>bv?sortDir:0;
  });
  const total=rows.length;
  const installed=rows.filter(r=>r.status==='INSTALLED').length;
  const failed=rows.filter(r=>r.status==='FAILED').length;
  const pending=rows.filter(r=>r.status==='PENDING'||r.status==='APPROVED').length;
  document.getElementById('stats').innerHTML=`
    <div class="stat"><div class="stat-label">Showing</div><div class="stat-value">${total}</div><div class="stat-sub">records</div></div>
    <div class="stat"><div class="stat-label">Installed</div><div class="stat-value" style="color:var(--green)">${installed}</div><div class="stat-sub">patches</div></div>
    <div class="stat"><div class="stat-label">Failed</div><div class="stat-value" style="color:var(--red)">${failed}</div><div class="stat-sub">patches</div></div>
    <div class="stat"><div class="stat-label">Pending</div><div class="stat-value" style="color:var(--yellow)">${pending}</div><div class="stat-sub">patches</div></div>`;
  document.getElementById('countLabel').textContent=total+' records';
  const tbody=document.getElementById('tbody');
  if(total===0){tbody.innerHTML='';document.getElementById('emptyMsg').style.display='block';return;}
  document.getElementById('emptyMsg').style.display='none';
  tbody.innerHTML=rows.map(r=>`<tr>
    <td>${r.patchName}</td>
    <td><span class="ver">${r.version}</span></td>
    <td><span class="badge badge-${r.status}">${r.status}</span></td>
    <td><span class="sev-${(r.impact||'OTHER').toUpperCase()}">${r.impact||'—'}</span></td>
    <td>${r.type||'—'}</td>
    <td>${r.deviceName}</td>
    <td>${r.orgName}</td>
    <td>${r.installedAt||'—'}</td>
  </tr>`).join('');
}
render();
</script>
</body></html>
"@

# ── Post to KB ────────────────────────────────────────────────────────────────
$ExistingId = $null
try {
    $Articles = Invoke-RestMethod -Uri "$BaseUrl/v2/knowledgebase/global/articles?folderId=$KbFolderId" -Headers $H
    $Ex = $Articles | Where-Object { $_.name -eq $ArticleName } | Select-Object -First 1
    if ($Ex) { $ExistingId = $Ex.id }
} catch {}

$Body = @{ name = $ArticleName; content = $Html; folderId = $KbFolderId } | ConvertTo-Json -Compress

try {
    if ($ExistingId) {
        Invoke-RestMethod -Uri "$BaseUrl/v2/knowledgebase/article/$ExistingId" -Method PATCH -Headers $H -Body $Body | Out-Null
        Write-Host "  [✓] KB article updated (ID: $ExistingId)." -ForegroundColor Green
    } else {
        Invoke-RestMethod -Uri "$BaseUrl/v2/knowledgebase/articles" -Method POST -Headers $H -Body $Body | Out-Null
        Write-Host "  [✓] KB article created." -ForegroundColor Green
    }
} catch {
    $sc = $null; try { $sc = $_.Exception.Response.StatusCode.Value__ } catch {}
    Write-Host "  [!] Failed to post to KB (HTTP $sc): $_" -ForegroundColor Red
    $f = "$env:TEMP\NinjaPatchReport_AllOrgs_$(Get-Date -Format yyyyMMdd_HHmmss).html"
    $Html | Out-File -FilePath $f -Encoding UTF8
    Write-Host "  [i] Saved locally: $f" -ForegroundColor Yellow
    exit 1
}

Write-Host ""
Write-Host "  ============================================================" -ForegroundColor Green
Write-Host "  [✓] COMPLETE" -ForegroundColor Green
Write-Host "      Records  : $($Rows.Count) patch records" -ForegroundColor Green
Write-Host "      Devices  : $($Devices.Count)" -ForegroundColor Green
Write-Host "      Orgs     : $($Orgs.Count)" -ForegroundColor Green
Write-Host "      Article  : $ArticleName" -ForegroundColor Green
Write-Host "  ============================================================" -ForegroundColor Green
Write-Host "  View: NinjaOne > Knowledge Base > folder $KbFolderId > $ArticleName" -ForegroundColor Cyan
Write-Host ""
