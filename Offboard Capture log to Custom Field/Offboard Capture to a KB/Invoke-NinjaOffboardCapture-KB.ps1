#Requires -Version 5.1
<#
.SYNOPSIS
    Captures device information and an offboard reason, then creates or updates
    a Knowledge Base article in NinjaOne with a fully formatted HTML report.

.DESCRIPTION
    Uses the NinjaOne Public API v2 with Client Credentials (silent, no browser
    login required) to:

      1. Authenticate silently using your API Client ID and Secret
      2. Pull detailed device information for the target device ID (read from
         NinjaOne Script Variables set when triggering the script)
      3. Collect the offboard reason entered as a NinjaOne Script Variable
      4. Build a clean, formatted HTML report of all captured data
      5. Check whether a Knowledge Base article already exists for this device
         in the configured KB folder — create it if new, update it if it exists

    Each offboarded device gets its own KB article named after the hostname.
    Articles live permanently in the KB folder you specify and can be viewed
    by any technician with Knowledge Base access.

.NOTES
    ── HOW TO RUN THIS IN NINJAONE ──────────────────────────────────────────────
    1. Fill in the CONFIGURATION block below — ALL six values
    2. Go to Administration > Scripting > Scripts > Add Script
    3. Paste this entire script, set language to PowerShell, Run As: System
    4. Create two Script Variables (Administration > Scripting > Script Variables):
         Name: targetDeviceId  | Type: Integer | Label: Target Device ID
         Name: offboardReason  | Type: Text    | Label: Offboard Reason
    5. Run the script against any managed device — it targets the device ID you
       enter in the Script Variable, not necessarily the device it physically runs on

    ── FINDING YOUR DEVICE ID ────────────────────────────────────────────────────
    Open the device in NinjaOne. The device ID is in the browser URL:
      https://app.ninjarmm.com/#/deviceDashboard/12345/overview
                                                   ^^^^^
    That number is your device ID. Enter it in the Script Variable when running.

    ── FINDING YOUR KB FOLDER ID ─────────────────────────────────────────────────
    In NinjaOne, go to Knowledge Base and open or create a folder for offboard
    records. The folder ID is in the URL:
      https://app.ninjarmm.com/#/knowledgeBase/folder/42
                                                       ^^
    Copy that number into $KbFolderId below.

    ── API APP SETUP (one-time) ──────────────────────────────────────────────────
    Go to: Administration > Apps > API > Client App IDs > Add
      Platform      : API Services (Machine-to-Machine)
      Allowed Scopes: monitoring  AND  management
      Redirect URI  : Leave blank — not needed for Client Credentials
    Click Save. Copy the Client ID and Client Secret shown.

    ── REGIONAL BASE URLS ────────────────────────────────────────────────────────
    US       : https://app.ninjarmm.com
    EU       : https://eu.ninjarmm.com
    Oceania  : https://oc.ninjarmm.com
    Canada   : https://ca.ninjarmm.com
#>

# ==============================================================================
#  CONFIGURATION — Fill in ALL six values below before saving/running the script
# ==============================================================================

# Your NinjaOne login URL (no trailing slash)
# Example: https://app.ninjarmm.com
$BaseUrl       = 'https://<your Login URL>'

# Same URL with /ws/oauth/token appended
# Example: https://app.ninjarmm.com/ws/oauth/token
$TokenEndpoint = 'https://<your Login URL>/ws/oauth/token'

# From Administration > Apps > API > Client App IDs
$ClientId      = '<Your Client ID>'

# From Administration > Apps > API > Client App IDs (shown once at creation)
$ClientSecret  = '<Your Client Secret>'

# The Knowledge Base folder ID where offboard articles will be saved.
# Find it in the URL when you open the folder: .../knowledgeBase/folder/42
# Create the folder first in NinjaOne if it does not exist yet.
$KbFolderId    = 0   # <-- Replace 0 with your actual folder ID number

# The name prefix for KB articles — device hostname is appended automatically.
# Example result: "Offboard Report — DESKTOP-ABC123"
$ArticlePrefix = 'Offboard Report'

# ==============================================================================
#  DO NOT EDIT BELOW THIS LINE
# ==============================================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Validate config ────────────────────────────────────────────────────────────
if ($BaseUrl -like '*<*') {
    Write-Error "[!] Please fill in BaseUrl in the CONFIGURATION block before running."
    exit 1
}
if ($ClientId -like '*<*' -or $ClientSecret -like '*<*') {
    Write-Error "[!] Please fill in ClientId and ClientSecret in the CONFIGURATION block."
    exit 1
}
if ($KbFolderId -eq 0) {
    Write-Error "[!] Please set KbFolderId to your Knowledge Base folder ID (not 0)."
    exit 1
}

# ── Read NinjaOne Script Variables ────────────────────────────────────────────
# NinjaOne injects these as environment variables when the script runs.
# They map to the Script Variables created in Administration > Scripting.
$TargetDeviceId = $env:targetDeviceId
$OffboardReason = $env:offboardReason

if (-not $TargetDeviceId) {
    Write-Error "[!] Script Variable 'targetDeviceId' is empty. Set it before running."
    exit 1
}
if (-not $OffboardReason) {
    Write-Error "[!] Script Variable 'offboardReason' is empty. Enter a reason before running."
    exit 1
}

Write-Host ""
Write-Host "  ============================================================" -ForegroundColor Cyan
Write-Host "  NinjaOne Offboard Capture  →  Knowledge Base" -ForegroundColor Cyan
Write-Host "  Device ID  : $TargetDeviceId" -ForegroundColor Cyan
Write-Host "  KB Folder  : $KbFolderId" -ForegroundColor Cyan
Write-Host "  ============================================================" -ForegroundColor Cyan
Write-Host ""

# ── Helper: call the API with error handling ───────────────────────────────────
function Invoke-NinjaApi {
    param(
        [string]$Path,
        [string]$Method = 'GET',
        [hashtable]$Headers,
        [string]$Body   = $null
    )
    $Params = @{
        Uri     = "$BaseUrl/v2/$Path"
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
        throw "API [$Method /v2/$Path] HTTP $Status — $_"
    }
}

# ── Step 1: Authenticate via Client Credentials ───────────────────────────────
Write-Host "  [1/5] Authenticating (Client Credentials)..." -ForegroundColor Cyan

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
    Write-Host "  [!] Authentication failed." -ForegroundColor Red
    Write-Host "      Check BaseUrl, ClientId, and ClientSecret." -ForegroundColor Yellow
    Write-Host "      API app must be 'API Services (Machine-to-Machine)' with" -ForegroundColor Yellow
    Write-Host "      monitoring AND management scopes enabled." -ForegroundColor Yellow
    Write-Host "      Error: $_" -ForegroundColor Red
    exit 1
}

$Headers = @{
    Authorization  = "Bearer $AccessToken"
    Accept         = 'application/json'
    'Content-Type' = 'application/json'
}

Write-Host "  [✓] Authenticated." -ForegroundColor Green

# ── Step 2: Pull device information ───────────────────────────────────────────
Write-Host ""
Write-Host "  [2/5] Fetching device information (ID: $TargetDeviceId)..." -ForegroundColor Cyan

try {
    $Device = Invoke-NinjaApi -Path "device/$TargetDeviceId" -Headers $Headers
} catch {
    if ($_ -like '*404*') {
        Write-Host "  [!] Device ID $TargetDeviceId was not found in NinjaOne." -ForegroundColor Red
        Write-Host "      Verify the ID from the device URL in your browser." -ForegroundColor Yellow
    } else {
        Write-Host "  [!] Failed to fetch device: $_" -ForegroundColor Red
    }
    exit 1
}

# Extended hardware — non-fatal
$DeviceDetails = $null
try {
    $DeviceDetails = Invoke-NinjaApi -Path "device/$TargetDeviceId/system-info" -Headers $Headers
} catch {
    Write-Host "  [i] Extended system info not available for this device type." -ForegroundColor Gray
}

# Installed software — non-fatal
$SoftwareList = $null
try {
    $SoftwareList = Invoke-NinjaApi -Path "device/$TargetDeviceId/software" -Headers $Headers
} catch {
    Write-Host "  [i] Software list not available for this device type." -ForegroundColor Gray
}

# Last activity — non-fatal
$LastActivity = $null
try {
    $Activities = Invoke-NinjaApi -Path "device/$TargetDeviceId/activities?pageSize=1" -Headers $Headers
    if ($Activities.activities) { $LastActivity = $Activities.activities[0] }
} catch {}

$Hostname = if ($Device.systemName)  { $Device.systemName } else { "Device-$TargetDeviceId" }
$OrgId    = $Device.organizationId
$LastSeen = if ($Device.lastContact) {
    [DateTimeOffset]::FromUnixTimeMilliseconds($Device.lastContact).ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss')
} else { 'Unknown' }

Write-Host "  [✓] Device: $Hostname  (Org ID: $OrgId)" -ForegroundColor Green

# ── Step 3: Get organization name ─────────────────────────────────────────────
Write-Host ""
Write-Host "  [3/5] Fetching organization name (ID: $OrgId)..." -ForegroundColor Cyan

$OrgName = "Organization $OrgId"
try {
    $Org     = Invoke-NinjaApi -Path "organization/$OrgId" -Headers $Headers
    $OrgName = $Org.name
} catch {
    Write-Host "  [i] Could not retrieve org name — continuing." -ForegroundColor Gray
}

Write-Host "  [✓] Organization: $OrgName" -ForegroundColor Green

# ── Step 4: Build the HTML article content ────────────────────────────────────
Write-Host ""
Write-Host "  [4/5] Building HTML report..." -ForegroundColor Cyan

$Timestamp   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
$ArticleName = "$ArticlePrefix — $Hostname"

# Hardware section
$HardwareRows = '<tr><td colspan="2"><em>Not available for this device type.</em></td></tr>'
if ($DeviceDetails) {
    $RamGb   = if ($DeviceDetails.memory.memoryMb) {
        [math]::Round($DeviceDetails.memory.memoryMb / 1024, 1)
    } else { 'N/A' }
    $CpuName = if ($DeviceDetails.processors) {
        [System.Web.HttpUtility]::HtmlEncode(($DeviceDetails.processors | Select-Object -First 1).name)
    } else { 'N/A' }
    $HardwareRows = @"
      <tr><td>Manufacturer</td><td>$([System.Web.HttpUtility]::HtmlEncode($DeviceDetails.system.manufacturer))</td></tr>
      <tr><td>Model</td><td>$([System.Web.HttpUtility]::HtmlEncode($DeviceDetails.system.model))</td></tr>
      <tr><td>Serial Number</td><td>$([System.Web.HttpUtility]::HtmlEncode($DeviceDetails.bios.serialNumber))</td></tr>
      <tr><td>CPU</td><td>$CpuName</td></tr>
      <tr><td>RAM</td><td>$RamGb GB</td></tr>
"@
}

# Last activity section
$ActivityRows = '<tr><td colspan="2"><em>No recent activity recorded.</em></td></tr>'
if ($LastActivity) {
    $ActivityTime = [DateTimeOffset]::FromUnixTimeMilliseconds(
        $LastActivity.activityTime).ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss')
    $ActivityRows = @"
      <tr><td>Time</td><td>$ActivityTime</td></tr>
      <tr><td>Type</td><td>$([System.Web.HttpUtility]::HtmlEncode($LastActivity.type))</td></tr>
      <tr><td>Message</td><td>$([System.Web.HttpUtility]::HtmlEncode($LastActivity.message))</td></tr>
"@
}

# Software section
$SoftwareRows = '<tr><td colspan="2"><em>Software list not available.</em></td></tr>'
if ($SoftwareList -and $SoftwareList.Count -gt 0) {
    $SoftwareRows = ($SoftwareList | Select-Object -First 20 | ForEach-Object {
        "<tr><td>$([System.Web.HttpUtility]::HtmlEncode($_.name))</td>" +
        "<td>$([System.Web.HttpUtility]::HtmlEncode($_.version))</td></tr>"
    }) -join "`n      "
    if ($SoftwareList.Count -gt 20) {
        $SoftwareRows += "`n      <tr><td colspan='2'><em>Showing first 20 of $($SoftwareList.Count) applications.</em></td></tr>"
    }
}

# Build full HTML — NinjaOne KB renders HTML content directly in the article viewer
$ArticleHtml = @"
<div style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;max-width:900px;color:#1a1a1a;">

  <!-- Header banner -->
  <div style="background:linear-gradient(135deg,#1F4E79,#2E75B6);border-radius:8px;padding:24px 28px;margin-bottom:24px;">
    <h1 style="margin:0;color:#ffffff;font-size:22px;font-weight:700;">&#128683; Device Offboard Report</h1>
    <p style="margin:6px 0 0;color:#BDD7EE;font-size:13px;">
      $Hostname &nbsp;&bull;&nbsp; $OrgName &nbsp;&bull;&nbsp; Captured: $Timestamp
    </p>
  </div>

  <!-- Offboard Reason -->
  <div style="background:#FFF8E1;border-left:4px solid #F0953A;border-radius:4px;padding:16px 20px;margin-bottom:20px;">
    <p style="margin:0 0 6px;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:0.6px;color:#C55A11;">
      Offboard Reason
    </p>
    <p style="margin:0;font-size:14px;color:#3D2800;line-height:1.6;">
      $([System.Web.HttpUtility]::HtmlEncode($OffboardReason))
    </p>
  </div>

  <!-- Two-column: Device + Hardware -->
  <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:20px;">

    <!-- Device Details -->
    <div style="background:#F8FAFF;border:1px solid #D0DCF0;border-radius:6px;overflow:hidden;">
      <div style="background:#2E75B6;padding:10px 16px;">
        <span style="color:#fff;font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:0.6px;">
          &#128187; Device Details
        </span>
      </div>
      <table style="width:100%;border-collapse:collapse;font-size:13px;">
        <tr style="background:#EEF4FB;">
          <td style="padding:8px 14px;font-weight:600;color:#1F4E79;width:38%;border-bottom:1px solid #D0DCF0;">System Name</td>
          <td style="padding:8px 14px;border-bottom:1px solid #D0DCF0;">$([System.Web.HttpUtility]::HtmlEncode($Device.systemName))</td>
        </tr>
        <tr>
          <td style="padding:8px 14px;font-weight:600;color:#1F4E79;border-bottom:1px solid #D0DCF0;">DNS Name</td>
          <td style="padding:8px 14px;border-bottom:1px solid #D0DCF0;">$([System.Web.HttpUtility]::HtmlEncode($Device.dnsName))</td>
        </tr>
        <tr style="background:#EEF4FB;">
          <td style="padding:8px 14px;font-weight:600;color:#1F4E79;border-bottom:1px solid #D0DCF0;">IP Address(es)</td>
          <td style="padding:8px 14px;border-bottom:1px solid #D0DCF0;">$([System.Web.HttpUtility]::HtmlEncode(($Device.ipAddresses -join ', ')))</td>
        </tr>
        <tr>
          <td style="padding:8px 14px;font-weight:600;color:#1F4E79;border-bottom:1px solid #D0DCF0;">OS</td>
          <td style="padding:8px 14px;border-bottom:1px solid #D0DCF0;">$([System.Web.HttpUtility]::HtmlEncode("$($Device.os.name) $($Device.os.servicePack)".Trim()))</td>
        </tr>
        <tr style="background:#EEF4FB;">
          <td style="padding:8px 14px;font-weight:600;color:#1F4E79;border-bottom:1px solid #D0DCF0;">Device Class</td>
          <td style="padding:8px 14px;border-bottom:1px solid #D0DCF0;">$([System.Web.HttpUtility]::HtmlEncode($Device.nodeClass))</td>
        </tr>
        <tr>
          <td style="padding:8px 14px;font-weight:600;color:#1F4E79;border-bottom:1px solid #D0DCF0;">Last Contact</td>
          <td style="padding:8px 14px;border-bottom:1px solid #D0DCF0;">$LastSeen</td>
        </tr>
        <tr style="background:#EEF4FB;">
          <td style="padding:8px 14px;font-weight:600;color:#1F4E79;">Agent Version</td>
          <td style="padding:8px 14px;">$([System.Web.HttpUtility]::HtmlEncode($Device.agentVersion))</td>
        </tr>
      </table>
    </div>

    <!-- Hardware Info -->
    <div style="background:#F8FAFF;border:1px solid #D0DCF0;border-radius:6px;overflow:hidden;">
      <div style="background:#2E75B6;padding:10px 16px;">
        <span style="color:#fff;font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:0.6px;">
          &#9881; Hardware Info
        </span>
      </div>
      <table style="width:100%;border-collapse:collapse;font-size:13px;">
        $HardwareRows
      </table>
    </div>
  </div>

  <!-- Last Activity -->
  <div style="background:#F8FAFF;border:1px solid #D0DCF0;border-radius:6px;overflow:hidden;margin-bottom:20px;">
    <div style="background:#2E75B6;padding:10px 16px;">
      <span style="color:#fff;font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:0.6px;">
        &#128337; Last Recorded Activity
      </span>
    </div>
    <table style="width:100%;border-collapse:collapse;font-size:13px;">
      $ActivityRows
    </table>
  </div>

  <!-- Installed Software -->
  <div style="background:#F8FAFF;border:1px solid #D0DCF0;border-radius:6px;overflow:hidden;margin-bottom:20px;">
    <div style="background:#2E75B6;padding:10px 16px;">
      <span style="color:#fff;font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:0.6px;">
        &#128230; Installed Software (first 20)
      </span>
    </div>
    <table style="width:100%;border-collapse:collapse;font-size:13px;">
      <tr style="background:#1F4E79;">
        <th style="padding:8px 14px;color:#fff;text-align:left;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:0.5px;">Application</th>
        <th style="padding:8px 14px;color:#fff;text-align:left;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:0.5px;">Version</th>
      </tr>
      $SoftwareRows
    </table>
  </div>

  <!-- Footer -->
  <div style="text-align:center;padding:14px;border-top:1px solid #D0DCF0;color:#8892A4;font-size:11px;">
    Generated by Invoke-NinjaOffboardCapture-KB.ps1 &nbsp;&bull;&nbsp; Device ID: $TargetDeviceId &nbsp;&bull;&nbsp; Org ID: $OrgId
  </div>

</div>
"@

Write-Host "  [✓] HTML report built." -ForegroundColor Green

# ── Step 5: Create or update the KB article ───────────────────────────────────
Write-Host ""
Write-Host "  [5/5] Posting to Knowledge Base folder (ID: $KbFolderId)..." -ForegroundColor Cyan

# Check if an article with this name already exists in the folder
$ExistingArticleId = $null
try {
    $KbArticles = Invoke-NinjaApi -Path "knowledgebase/global/articles?folderId=$KbFolderId" -Headers $Headers
    if ($KbArticles) {
        $Existing = $KbArticles | Where-Object { $_.name -eq $ArticleName } | Select-Object -First 1
        if ($Existing) { $ExistingArticleId = $Existing.id }
    }
} catch {
    Write-Host "  [i] Could not check for existing articles — will attempt to create." -ForegroundColor Gray
}

$ArticleBody = @{
    name     = $ArticleName
    content  = $ArticleHtml
    folderId = $KbFolderId
} | ConvertTo-Json -Compress

if ($ExistingArticleId) {
    Write-Host "  [i] Article '$ArticleName' exists (ID: $ExistingArticleId) — updating." -ForegroundColor Gray
    try {
        Invoke-RestMethod `
            -Uri     "$BaseUrl/v2/knowledgebase/article/$ExistingArticleId" `
            -Method  PATCH `
            -Headers $Headers `
            -Body    $ArticleBody `
            -ContentType 'application/json' | Out-Null
        Write-Host "  [✓] KB article updated." -ForegroundColor Green
    } catch {
        Write-Host "  [!] Failed to update KB article: $_" -ForegroundColor Red
        exit 1
    }
} else {
    try {
        Invoke-RestMethod `
            -Uri     "$BaseUrl/v2/knowledgebase/articles" `
            -Method  POST `
            -Headers $Headers `
            -Body    $ArticleBody `
            -ContentType 'application/json' | Out-Null
        Write-Host "  [✓] KB article created." -ForegroundColor Green
    } catch {
        $Status = $null
        try { $Status = $_.Exception.Response.StatusCode.Value__ } catch {}
        Write-Host "  [!] Failed to create KB article (HTTP $Status)." -ForegroundColor Red
        Write-Host "      Check that KbFolderId $KbFolderId exists in your Knowledge Base." -ForegroundColor Yellow
        Write-Host "      Ensure the API app has the 'management' scope." -ForegroundColor Yellow
        Write-Host "      Error: $_" -ForegroundColor Red
        exit 1
    }
}

# ── Summary ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ============================================================" -ForegroundColor Green
Write-Host "  [✓] OFFBOARD CAPTURE COMPLETE" -ForegroundColor Green
Write-Host "      Device      : $Hostname (ID: $TargetDeviceId)" -ForegroundColor Green
Write-Host "      Org         : $OrgName (ID: $OrgId)" -ForegroundColor Green
Write-Host "      KB Folder   : $KbFolderId" -ForegroundColor Green
Write-Host "      Article     : $ArticleName" -ForegroundColor Green
Write-Host "      Timestamp   : $Timestamp" -ForegroundColor Green
Write-Host "  ============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  To view: NinjaOne > Knowledge Base > your folder > $ArticleName" -ForegroundColor Cyan
Write-Host ""
