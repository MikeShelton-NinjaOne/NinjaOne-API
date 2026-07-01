#Requires -Version 5.1
<#
.SYNOPSIS
    Pulls software patch data for a SINGLE device, cross-references installed software
    versions, and posts an interactive filterable HTML report to a NinjaOne KB folder.

.DESCRIPTION
    Data sources:
      GET /v2/device/{id}/software-patches  — patch records for this device
      GET /v2/device/{id}/software          — installed software (for version data)
      GET /v2/device/{id}                   — device name and org ID
      GET /v2/organization/{id}             — org name

    Run this standalone OR trigger it from NinjaOne using a Script Variable for Device ID.

.NOTES
    ── API APP SETUP (one-time) ──────────────────────────────────────────────────
    Administration > Apps > API > Client App IDs > Add
      Platform: API Services (Machine-to-Machine)
      Scopes  : monitoring  AND  management

    ── HOW TO RUN FROM NINJAONE ─────────────────────────────────────────────────
    Create one Script Variable:
      Name: targetDeviceId  |  Type: Integer  |  Label: Target Device ID

    ── FINDING YOUR DEVICE ID ───────────────────────────────────────────────────
    Open the device in NinjaOne. Look at the URL:
      https://app.ninjarmm.com/#/deviceDashboard/12345/overview
                                                   ^^^^^

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

# When running OUTSIDE NinjaOne, set this directly. When running IN NinjaOne,
# leave as $null — it will be read from the Script Variable automatically.
$ManualDeviceId = $null   # e.g. 12345

# ==============================================================================
#  DO NOT EDIT BELOW THIS LINE
# ==============================================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ($BaseUrl -like '*<*')                            { Write-Error 'Fill in BaseUrl';       exit 1 }
if ($ClientId -like '*<*' -or $ClientSecret -like '*<*') { Write-Error 'Fill in credentials'; exit 1 }
if ($KbFolderId -eq 0)                               { Write-Error 'Fill in KbFolderId';   exit 1 }

# Resolve device ID — NinjaOne Script Variable takes priority
$TargetDeviceId = if ($env:targetDeviceId) { $env:targetDeviceId } else { $ManualDeviceId }
if (-not $TargetDeviceId) {
    Write-Error "[!] No device ID provided. Set ManualDeviceId in the config or use the 'targetDeviceId' Script Variable."
    exit 1
}

Write-Host ""
Write-Host "  ============================================================" -ForegroundColor Cyan
Write-Host "  NinjaOne Software Patch Report  [Single Device: $TargetDeviceId]" -ForegroundColor Cyan
Write-Host "  ============================================================" -ForegroundColor Cyan

# ── Auth ──────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  [1/5] Authenticating..." -ForegroundColor Cyan
try {
    $Token = Invoke-RestMethod -Uri $TokenEndpoint -Method POST -ContentType 'application/x-www-form-urlencoded' -Body @{
        grant_type = 'client_credentials'; client_id = $ClientId
        client_secret = $ClientSecret; scope = 'monitoring management'
    }
    $H = @{ Authorization = "Bearer $($Token.access_token)"; Accept = 'application/json'; 'Content-Type' = 'application/json' }
} catch { Write-Host "  [!] Auth failed: $_" -ForegroundColor Red; exit 1 }
Write-Host "  [✓] Authenticated." -ForegroundColor Green

# ── Device info ───────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  [2/5] Fetching device and org info..." -ForegroundColor Cyan
try {
    $Device = Invoke-RestMethod -Uri "$BaseUrl/v2/device/$TargetDeviceId" -Headers $H
} catch {
    $sc = $null; try { $sc = $_.Exception.Response.StatusCode.Value__ } catch {}
    if ($sc -eq 404) { Write-Host "  [!] Device $TargetDeviceId not found. Check the ID from the device URL." -ForegroundColor Red }
    else { Write-Host "  [!] Failed: $_" -ForegroundColor Red }
    exit 1
}
$Hostname = if ($Device.systemName) { $Device.systemName } else { "Device $TargetDeviceId" }
$OrgId    = $Device.organizationId
$OrgName  = "Organization $OrgId"
try { $OrgName = (Invoke-RestMethod -Uri "$BaseUrl/v2/organization/$OrgId" -Headers $H).name } catch {}
Write-Host "  [✓] Device: $Hostname  |  Org: $OrgName" -ForegroundColor Green

# ── Software inventory (version data) ─────────────────────────────────────────
Write-Host ""
Write-Host "  [3/5] Fetching installed software..." -ForegroundColor Cyan
$Software = @()
try { $Software = Invoke-RestMethod -Uri "$BaseUrl/v2/device/$TargetDeviceId/software" -Headers $H } catch {
    Write-Host "  [i] Software list not available for this device type." -ForegroundColor Gray
}
$VersionMap = @{}
foreach ($sw in $Software) {
    $key = if ($sw.productCode) { $sw.productCode } else { $sw.name }
    if ($key -and $sw.version) { $VersionMap[$key] = $sw.version }
}
Write-Host "  [✓] $($Software.Count) software records." -ForegroundColor Green

# ── Patch data ────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  [4/5] Fetching software patches..." -ForegroundColor Cyan
$Patches = @()
try { $Patches = Invoke-RestMethod -Uri "$BaseUrl/v2/device/$TargetDeviceId/software-patches" -Headers $H } catch {
    Write-Host "  [!] Failed to fetch patches: $_" -ForegroundColor Red; exit 1
}
Write-Host "  [✓] $($Patches.Count) patch records." -ForegroundColor Green

# ── Build dataset ─────────────────────────────────────────────────────────────
$Rows = $Patches | ForEach-Object {
    $ver = 'N/A'
    if ($_.identifier -and $VersionMap.ContainsKey($_.identifier)) { $ver = $VersionMap[$_.identifier] }
    elseif ($_.name   -and $VersionMap.ContainsKey($_.name))       { $ver = $VersionMap[$_.name] }
    $ia = if ($_.installedAt) {
        [DateTimeOffset]::FromUnixTimeMilliseconds($_.installedAt).ToLocalTime().ToString('yyyy-MM-dd HH:mm')
    } else { '' }
    [PSCustomObject]@{
        patchName  = $_.name; version = $ver; status = $_.status
        impact = $_.impact; type = $_.type; identifier = $_.identifier; installedAt = $ia
    }
}
$JsonData    = $Rows | ConvertTo-Json -Depth 3 -Compress
$GeneratedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
$ArticleName = "Software Patch Report — $Hostname"

# ── Build HTML ────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  [5/5] Building HTML and posting to KB..." -ForegroundColor Cyan

$Html = @"
<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<title>Software Patch Report — $Hostname</title>
<style>
:root{--blue:#2E75B6;--dark:#1F4E79;--card:#f0f4fa;--border:#c8d8ee;--text:#1a1a1a;--muted:#6b7a99;
--green:#1a7a3f;--green-bg:#e8f5ee;--red:#b91c1c;--red-bg:#fef2f2;--yellow:#b45309;--yellow-bg:#fffbeb;
--gray:#4b5563;--gray-bg:#f3f4f6;--purple:#7c3aed;--purple-bg:#f5f3ff;}
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#f8faff;color:var(--text);}
.header{background:linear-gradient(135deg,var(--dark),var(--blue));padding:22px 32px;}
.header h1{color:#fff;font-size:20px;font-weight:700;}
.header p{color:#bdd7ee;font-size:12px;margin-top:4px;}
.controls{background:#fff;border-bottom:1px solid var(--border);padding:14px 32px;display:flex;flex-wrap:wrap;gap:12px;align-items:center;}
select,input[type=text]{background:#f8faff;border:1px solid var(--border);border-radius:6px;padding:7px 11px;font-size:13px;outline:none;}
select:focus,input:focus{border-color:var(--blue);}
input[type=text]{width:220px;}
.btn-ghost{background:transparent;color:var(--muted);border:1px solid var(--border);border-radius:6px;padding:7px 14px;font-size:13px;cursor:pointer;}
.btn-ghost:hover{background:var(--card);}
.stats{display:flex;gap:14px;padding:16px 32px;flex-wrap:wrap;}
.stat{background:#fff;border:1px solid var(--border);border-radius:8px;padding:14px 20px;min-width:120px;}
.stat-label{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;font-weight:600;}
.stat-value{font-size:26px;font-weight:700;color:var(--blue);line-height:1.2;margin-top:3px;}
.main{padding:0 32px 40px;}
.table-wrap{background:#fff;border:1px solid var(--border);border-radius:8px;overflow:hidden;}
table{width:100%;border-collapse:collapse;font-size:13px;}
th{background:var(--dark);color:#fff;padding:10px 14px;text-align:left;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;cursor:pointer;user-select:none;white-space:nowrap;}
th:hover{background:var(--blue);}
td{padding:9px 14px;border-top:1px solid var(--border);}
tr:hover td{background:var(--card);}
.badge{display:inline-block;padding:2px 9px;border-radius:20px;font-size:11px;font-weight:700;text-transform:uppercase;}
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
  <h1>&#129520; Software Patch Report</h1>
  <p>Device: <strong style="color:#fff">$Hostname</strong> &nbsp;&bull;&nbsp; Org: $OrgName &nbsp;&bull;&nbsp; Generated: $GeneratedAt</p>
</div>
<div class="controls">
  <select id="fStatus" onchange="render()"><option value="">All Statuses</option><option>INSTALLED</option><option>PENDING</option><option>FAILED</option><option>REJECTED</option><option>APPROVED</option></select>
  <select id="fImpact" onchange="render()"><option value="">All Severities</option><option>CRITICAL</option><option>IMPORTANT</option><option>MODERATE</option><option>LOW</option></select>
  <input type="text" id="fSearch" placeholder="&#128269; Search patch name..." oninput="render()">
  <button class="btn-ghost" onclick="resetFilters()">Reset</button>
  <span id="countLabel" style="font-size:12px;color:var(--muted);margin-left:auto;"></span>
</div>
<div class="stats" id="stats"></div>
<div class="main">
  <div class="table-wrap">
    <table><thead><tr>
      <th onclick="sortBy('patchName')">Patch Name &#8597;</th>
      <th onclick="sortBy('version')">Version &#8597;</th>
      <th onclick="sortBy('status')">Status &#8597;</th>
      <th onclick="sortBy('impact')">Severity &#8597;</th>
      <th onclick="sortBy('type')">Type &#8597;</th>
      <th onclick="sortBy('installedAt')">Installed At &#8597;</th>
    </tr></thead><tbody id="tbody"></tbody></table>
    <div class="empty" id="emptyMsg" style="display:none;">No records match your filters.</div>
  </div>
</div>
<script>
const DATA=$JsonData;
let sortCol='patchName',sortDir=1;
function sortBy(c){if(sortCol===c){sortDir*=-1;}else{sortCol=c;sortDir=1;}render();}
function resetFilters(){['fStatus','fImpact'].forEach(id=>document.getElementById(id).value='');document.getElementById('fSearch').value='';render();}
function render(){
  const st=document.getElementById('fStatus').value.toUpperCase();
  const im=document.getElementById('fImpact').value.toUpperCase();
  const q=document.getElementById('fSearch').value.toLowerCase();
  let rows=DATA.filter(r=>{
    if(st&&r.status.toUpperCase()!==st)return false;
    if(im&&(r.impact||'').toUpperCase()!==im)return false;
    if(q&&!r.patchName.toLowerCase().includes(q))return false;
    return true;
  });
  rows.sort((a,b)=>{const av=a[sortCol]||'';const bv=b[sortCol]||'';return av<bv?-sortDir:av>bv?sortDir:0;});
  const tot=rows.length;
  document.getElementById('stats').innerHTML=`
    <div class="stat"><div class="stat-label">Showing</div><div class="stat-value">${tot}</div><div class="stat-sub">records</div></div>
    <div class="stat"><div class="stat-label">Installed</div><div class="stat-value" style="color:var(--green)">${rows.filter(r=>r.status==='INSTALLED').length}</div></div>
    <div class="stat"><div class="stat-label">Failed</div><div class="stat-value" style="color:var(--red)">${rows.filter(r=>r.status==='FAILED').length}</div></div>
    <div class="stat"><div class="stat-label">Pending</div><div class="stat-value" style="color:var(--yellow)">${rows.filter(r=>r.status==='PENDING'||r.status==='APPROVED').length}</div></div>`;
  document.getElementById('countLabel').textContent=tot+' records';
  if(tot===0){document.getElementById('tbody').innerHTML='';document.getElementById('emptyMsg').style.display='block';return;}
  document.getElementById('emptyMsg').style.display='none';
  document.getElementById('tbody').innerHTML=rows.map(r=>`<tr>
    <td>${r.patchName}</td>
    <td><span class="ver">${r.version}</span></td>
    <td><span class="badge badge-${r.status}">${r.status}</span></td>
    <td><span class="sev-${(r.impact||'OTHER').toUpperCase()}">${r.impact||'—'}</span></td>
    <td>${r.type||'—'}</td>
    <td>${r.installedAt||'—'}</td>
  </tr>`).join('');
}
render();
</script></body></html>
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
        Write-Host "  [✓] KB article updated." -ForegroundColor Green
    } else {
        Invoke-RestMethod -Uri "$BaseUrl/v2/knowledgebase/articles" -Method POST -Headers $H -Body $Body | Out-Null
        Write-Host "  [✓] KB article created." -ForegroundColor Green
    }
} catch {
    $sc=$null;try{$sc=$_.Exception.Response.StatusCode.Value__}catch{}
    Write-Host "  [!] Failed to post to KB (HTTP $sc): $_" -ForegroundColor Red
    $f="$env:TEMP\NinjaPatchReport_${Hostname}_$(Get-Date -Format yyyyMMdd_HHmmss).html"
    $Html | Out-File -FilePath $f -Encoding UTF8
    Write-Host "  [i] Saved locally: $f" -ForegroundColor Yellow; exit 1
}

Write-Host ""
Write-Host "  ============================================================" -ForegroundColor Green
Write-Host "  [✓] COMPLETE" -ForegroundColor Green
Write-Host "      Device   : $Hostname (ID: $TargetDeviceId)" -ForegroundColor Green
Write-Host "      Org      : $OrgName" -ForegroundColor Green
Write-Host "      Patches  : $($Patches.Count) records" -ForegroundColor Green
Write-Host "      Article  : $ArticleName" -ForegroundColor Green
Write-Host "  ============================================================" -ForegroundColor Green
Write-Host "  View: NinjaOne > Knowledge Base > folder $KbFolderId > $ArticleName" -ForegroundColor Cyan
Write-Host ""
