#Requires -Version 5.1
<#
.SYNOPSIS
    Captures device information and an offboard reason, then writes a structured
    offboard report to an Apps & Services document on the device's organization
    in NinjaOne.

.DESCRIPTION
    Uses the NinjaOne Public API v2 with Client Credentials (silent, no browser
    login required) to:

      1. Authenticate silently using your API Client ID and Secret
      2. Pull detailed device information for the target device ID (via NinjaOne
         Script Variables set when triggering the script)
      3. Collect the offboard reason entered as a NinjaOne Script Variable
      4. Check whether the "Device Offboard Report" document template exists
         on the organization — create it automatically via API if it does not
      5. Create a new Apps & Services document on the organization named after
         the device hostname, populating all template fields with captured data

    Each offboarded device gets its own named document under the organization's
    Apps & Services tab. Documents stack up over time as a permanent offboard log.

.NOTES
    ── HOW TO RUN THIS IN NINJAONE ──────────────────────────────────────────────
    1. Fill in the CONFIGURATION block below (the four credential lines)
    2. Go to Administration > Scripting > Scripts > Add Script
    3. Paste this entire script, set language to PowerShell, Run As: System
    4. Create two Script Variables (Administration > Scripting > Script Variables):
         Name: targetDeviceId  | Type: Integer | Label: Target Device ID
         Name: offboardReason  | Type: Text    | Label: Offboard Reason
    5. Run the script against any managed device — it targets the device ID you
       enter in the Script Variable, not necessarily the device it runs on

    ── FINDING YOUR DEVICE ID ────────────────────────────────────────────────────
    Open the device in NinjaOne. The device ID is in the URL:
      https://app.ninjarmm.com/#/deviceDashboard/12345/overview
                                                   ^^^^^
    That number is your device ID.

    ── API APP SETUP (one-time) ──────────────────────────────────────────────────
    Go to: Administration > Apps > API > Client App IDs > Add
      Platform      : API Services (Machine-to-Machine)
      Allowed Scopes: monitoring  AND  management
      Redirect URI  : Leave blank — not needed for Client Credentials
    Click Save. Copy the Client ID and Client Secret.

    ── REGIONAL BASE URLS ────────────────────────────────────────────────────────
    US       : https://app.ninjarmm.com
    EU       : https://eu.ninjarmm.com
    Oceania  : https://oc.ninjarmm.com
    Canada   : https://ca.ninjarmm.com
#>

# ==============================================================================
#  CONFIGURATION — Fill in ALL four values below before saving/running the script
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

# ==============================================================================
#  DO NOT EDIT BELOW THIS LINE
# ==============================================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# The exact name of the document template this script will create/use.
# If you change this, update it in the README too.
$TemplateName = 'Device Offboard Report'

# ── Validate config ────────────────────────────────────────────────────────────
if ($BaseUrl -like '*<*') {
    Write-Error "[!] Please fill in BaseUrl in the CONFIGURATION block before running."
    exit 1
}
if ($ClientId -like '*<*' -or $ClientSecret -like '*<*') {
    Write-Error "[!] Please fill in ClientId and ClientSecret in the CONFIGURATION block."
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
Write-Host "  NinjaOne Offboard Capture" -ForegroundColor Cyan
Write-Host "  Device ID   : $TargetDeviceId" -ForegroundColor Cyan
Write-Host "  ============================================================" -ForegroundColor Cyan
Write-Host ""

# ── Helper: call the API with error handling ───────────────────────────────────
function Invoke-NinjaApi {
    param(
        [string]$Path,
        [string]$Method = 'GET',
        [hashtable]$Headers,
        [string]$Body = $null
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
Write-Host "  [1/6] Authenticating (Client Credentials)..." -ForegroundColor Cyan

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
    Write-Host "      Check BaseUrl, ClientId, ClientSecret." -ForegroundColor Yellow
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
Write-Host "  [2/6] Fetching device information (ID: $TargetDeviceId)..." -ForegroundColor Cyan

try {
    $Device = Invoke-NinjaApi -Path "device/$TargetDeviceId" -Headers $Headers
} catch {
    if ($_ -like '*404*') {
        Write-Host "  [!] Device ID $TargetDeviceId not found." -ForegroundColor Red
        Write-Host "      Verify the ID from the device URL in NinjaOne." -ForegroundColor Yellow
    } else {
        Write-Host "  [!] Failed to fetch device: $_" -ForegroundColor Red
    }
    exit 1
}

# Pull extended hardware/system info — non-fatal if unavailable
$DeviceDetails = $null
try {
    $DeviceDetails = Invoke-NinjaApi -Path "device/$TargetDeviceId/system-info" -Headers $Headers
} catch {
    Write-Host "  [i] Extended system info not available for this device type." -ForegroundColor Gray
}

# Pull installed software — non-fatal
$SoftwareList = $null
try {
    $SoftwareList = Invoke-NinjaApi -Path "device/$TargetDeviceId/software" -Headers $Headers
} catch {
    Write-Host "  [i] Software list not available for this device type." -ForegroundColor Gray
}

# Pull last activity — non-fatal
$LastActivity = $null
try {
    $Activities = Invoke-NinjaApi -Path "device/$TargetDeviceId/activities?pageSize=1" -Headers $Headers
    if ($Activities.activities) { $LastActivity = $Activities.activities[0] }
} catch {}

$Hostname  = if ($Device.systemName)  { $Device.systemName } else { "Device-$TargetDeviceId" }
$OrgId     = $Device.organizationId
$LastSeen  = if ($Device.lastContact) {
    [DateTimeOffset]::FromUnixTimeMilliseconds($Device.lastContact).ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss')
} else { 'Unknown' }

Write-Host "  [✓] Device: $Hostname  (Org ID: $OrgId)" -ForegroundColor Green

# ── Step 3: Get organization name ─────────────────────────────────────────────
Write-Host ""
Write-Host "  [3/6] Fetching organization info (ID: $OrgId)..." -ForegroundColor Cyan

$OrgName = "Organization $OrgId"
try {
    $Org     = Invoke-NinjaApi -Path "organization/$OrgId" -Headers $Headers
    $OrgName = $Org.name
} catch {
    Write-Host "  [i] Could not retrieve org name — continuing." -ForegroundColor Gray
}

Write-Host "  [✓] Organization: $OrgName" -ForegroundColor Green

# ── Step 4: Find or create the document template ───────────────────────────────
Write-Host ""
Write-Host "  [4/6] Checking for '$TemplateName' document template..." -ForegroundColor Cyan

$Template     = $null
$TemplateId   = $null
$AttributeMap = @{}   # fieldName -> attributeId

try {
    $Templates = Invoke-NinjaApi -Path 'document-templates' -Headers $Headers
    $Template  = $Templates | Where-Object { $_.name -eq $TemplateName } | Select-Object -First 1
} catch {
    Write-Host "  [i] Could not query templates — will attempt creation." -ForegroundColor Gray
}

if ($Template) {
    $TemplateId = $Template.id
    Write-Host "  [i] Template '$TemplateName' already exists (ID: $TemplateId)." -ForegroundColor Gray

    # Fetch the template with its attribute definitions so we can map field names to IDs
    try {
        $TemplateDetail = Invoke-NinjaApi -Path "document-templates/$TemplateId" -Headers $Headers
        foreach ($attr in $TemplateDetail.fields) {
            $AttributeMap[$attr.name] = $attr.attributeId
        }
    } catch {
        Write-Host "  [i] Could not fetch template attribute IDs — will use field names directly." -ForegroundColor Gray
    }

} else {
    Write-Host "  [i] Template not found — creating it now..." -ForegroundColor Gray

    # Build the template definition
    # Field types: TEXT, WYSIWYG, DATETIME, CHECKBOX, DROPDOWN, NUMERIC, ATTACHMENT
    $TemplateBody = @{
        name        = $TemplateName
        description = 'Automatically generated offboard capture report for a device.'
        fields      = @(
            @{ name = 'offboardReason';   label = 'Offboard Reason';    fieldType = 'WYSIWYG';  required = $false }
            @{ name = 'captureTimestamp'; label = 'Capture Date/Time';  fieldType = 'TEXT';     required = $false }
            @{ name = 'deviceDetails';    label = 'Device Details';     fieldType = 'WYSIWYG';  required = $false }
            @{ name = 'hardwareInfo';     label = 'Hardware Info';      fieldType = 'WYSIWYG';  required = $false }
            @{ name = 'lastActivity';     label = 'Last Activity';      fieldType = 'WYSIWYG';  required = $false }
            @{ name = 'softwareSummary';  label = 'Installed Software'; fieldType = 'WYSIWYG';  required = $false }
        )
    } | ConvertTo-Json -Depth 5

    try {
        $CreatedTemplate = Invoke-NinjaApi -Path 'document-templates' -Method POST -Headers $Headers -Body $TemplateBody
        $TemplateId      = $CreatedTemplate.id
        Write-Host "  [✓] Template created (ID: $TemplateId)." -ForegroundColor Green

        # Re-fetch to get attribute IDs
        Start-Sleep -Seconds 2
        $TemplateDetail = Invoke-NinjaApi -Path "document-templates/$TemplateId" -Headers $Headers
        foreach ($attr in $TemplateDetail.fields) {
            $AttributeMap[$attr.name] = $attr.attributeId
        }
    } catch {
        Write-Host "  [!] Failed to create document template: $_" -ForegroundColor Red
        Write-Host "      Ensure your API app has the 'management' scope." -ForegroundColor Yellow
        exit 1
    }
}

# ── Step 5: Build the field content ───────────────────────────────────────────
Write-Host ""
Write-Host "  [5/6] Building offboard report content..." -ForegroundColor Cyan

$Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

# Offboard Reason — plain text wrapped in HTML for WYSIWYG field
$ReasonHtml = "<p><strong>Reason:</strong> $([System.Web.HttpUtility]::HtmlEncode($OffboardReason))</p>"

# Device Details
$DeviceHtml = @"
<table>
  <tr><td><strong>System Name</strong></td><td>$($Device.systemName)</td></tr>
  <tr><td><strong>DNS Name</strong></td><td>$($Device.dnsName)</td></tr>
  <tr><td><strong>IP Address(es)</strong></td><td>$($Device.ipAddresses -join ', ')</td></tr>
  <tr><td><strong>OS</strong></td><td>$($Device.os.name) $($Device.os.servicePack)</td></tr>
  <tr><td><strong>Device Class</strong></td><td>$($Device.nodeClass)</td></tr>
  <tr><td><strong>Last Contact</strong></td><td>$LastSeen</td></tr>
  <tr><td><strong>Agent Version</strong></td><td>$($Device.agentVersion)</td></tr>
  <tr><td><strong>Organization</strong></td><td>$OrgName (ID: $OrgId)</td></tr>
</table>
"@

# Hardware info
$HardwareHtml = '<p><em>Not available for this device type.</em></p>'
if ($DeviceDetails) {
    $RamGb        = if ($DeviceDetails.memory.memoryMb) { [math]::Round($DeviceDetails.memory.memoryMb / 1024, 1) } else { 'N/A' }
    $CpuName      = if ($DeviceDetails.processors)      { ($DeviceDetails.processors | Select-Object -First 1).name } else { 'N/A' }
    $HardwareHtml = @"
<table>
  <tr><td><strong>Manufacturer</strong></td><td>$($DeviceDetails.system.manufacturer)</td></tr>
  <tr><td><strong>Model</strong></td><td>$($DeviceDetails.system.model)</td></tr>
  <tr><td><strong>Serial Number</strong></td><td>$($DeviceDetails.bios.serialNumber)</td></tr>
  <tr><td><strong>CPU</strong></td><td>$CpuName</td></tr>
  <tr><td><strong>RAM (GB)</strong></td><td>$RamGb</td></tr>
</table>
"@
}

# Last activity
$ActivityHtml = '<p><em>No recent activity recorded.</em></p>'
if ($LastActivity) {
    $ActivityTime = [DateTimeOffset]::FromUnixTimeMilliseconds($LastActivity.activityTime).ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss')
    $ActivityHtml = @"
<table>
  <tr><td><strong>Time</strong></td><td>$ActivityTime</td></tr>
  <tr><td><strong>Type</strong></td><td>$($LastActivity.type)</td></tr>
  <tr><td><strong>Message</strong></td><td>$($LastActivity.message)</td></tr>
</table>
"@
}

# Software list
$SoftwareHtml = '<p><em>Software list not available.</em></p>'
if ($SoftwareList -and $SoftwareList.Count -gt 0) {
    $Rows = ($SoftwareList | Select-Object -First 20 | ForEach-Object {
        "<tr><td>$([System.Web.HttpUtility]::HtmlEncode($_.name))</td><td>$($_.version)</td></tr>"
    }) -join ''
    $SoftwareHtml = "<table><tr><th>Application</th><th>Version</th></tr>$Rows</table>"
    if ($SoftwareList.Count -gt 20) {
        $SoftwareHtml += "<p><em>Showing first 20 of $($SoftwareList.Count) installed applications.</em></p>"
    }
}

Write-Host "  [✓] Content built." -ForegroundColor Green

# ── Step 6: Create the Apps & Services document on the organization ────────────
Write-Host ""
Write-Host "  [6/6] Creating Apps & Services document on organization '$OrgName'..." -ForegroundColor Cyan

# Document name = hostname so each device gets its own named document
$DocumentName = "Offboard — $Hostname"

# Helper: build a field entry — tries attributeId lookup first, falls back to name
function Build-Field {
    param([string]$FieldName, [string]$Value)
    $Entry = @{ value = $Value }
    if ($AttributeMap.ContainsKey($FieldName) -and $AttributeMap[$FieldName]) {
        $Entry['attributeId'] = $AttributeMap[$FieldName]
    } else {
        $Entry['name'] = $FieldName
    }
    return $Entry
}

$DocBody = @{
    name               = $DocumentName
    documentTemplateId = $TemplateId
    organizationId     = $OrgId
    fields             = @(
        (Build-Field 'offboardReason'   $ReasonHtml)
        (Build-Field 'captureTimestamp' $Timestamp)
        (Build-Field 'deviceDetails'    $DeviceHtml)
        (Build-Field 'hardwareInfo'     $HardwareHtml)
        (Build-Field 'lastActivity'     $ActivityHtml)
        (Build-Field 'softwareSummary'  $SoftwareHtml)
    )
} | ConvertTo-Json -Depth 6

# Check if a document for this device already exists on this org
$ExistingDocId = $null
try {
    $ExistingDocs  = Invoke-NinjaApi -Path "organization/$OrgId/documents" -Headers $Headers
    $ExistingDoc   = $ExistingDocs | Where-Object {
        $_.name -eq $DocumentName -and $_.documentTemplateId -eq $TemplateId
    } | Select-Object -First 1
    if ($ExistingDoc) { $ExistingDocId = $ExistingDoc.id }
} catch {
    Write-Host "  [i] Could not check for existing documents — will create new." -ForegroundColor Gray
}

if ($ExistingDocId) {
    Write-Host "  [i] Document '$DocumentName' already exists (ID: $ExistingDocId) — updating." -ForegroundColor Gray
    try {
        Invoke-NinjaApi -Path "organization/documents" -Method PATCH -Headers $Headers -Body $DocBody | Out-Null
        Write-Host "  [✓] Document updated." -ForegroundColor Green
    } catch {
        Write-Host "  [!] Failed to update document: $_" -ForegroundColor Red
        exit 1
    }
} else {
    try {
        Invoke-NinjaApi -Path "organization/documents" -Method POST -Headers $Headers -Body $DocBody | Out-Null
        Write-Host "  [✓] Document created." -ForegroundColor Green
    } catch {
        Write-Host "  [!] Failed to create Apps & Services document: $_" -ForegroundColor Red
        Write-Host "      Ensure management scope is enabled on the API app." -ForegroundColor Yellow
        Write-Host "      Error detail: $_" -ForegroundColor Red
        exit 1
    }
}

# ── Summary ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ============================================================" -ForegroundColor Green
Write-Host "  [✓] OFFBOARD CAPTURE COMPLETE" -ForegroundColor Green
Write-Host "      Device       : $Hostname (ID: $TargetDeviceId)" -ForegroundColor Green
Write-Host "      Organization : $OrgName (ID: $OrgId)" -ForegroundColor Green
Write-Host "      Document     : $DocumentName" -ForegroundColor Green
Write-Host "      Template     : $TemplateName (ID: $TemplateId)" -ForegroundColor Green
Write-Host "      Timestamp    : $Timestamp" -ForegroundColor Green
Write-Host "  ============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  To view: NinjaOne > Organizations > $OrgName > Documentation > Apps & Services" -ForegroundColor Cyan
Write-Host ""
