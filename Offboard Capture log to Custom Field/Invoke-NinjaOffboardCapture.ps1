#Requires -Version 5.1
<#
.SYNOPSIS
    Captures device information and an offboard reason, creates an org-level custom field
    named after the device hostname, and writes all collected data to that field.

.DESCRIPTION
    This script is designed to be triggered manually by a technician from within NinjaOne
    on a specific device. It uses the NinjaOne Public API v2 with Client Credentials
    (silent, no browser login required) to:

      1. Authenticate silently using your API Client ID and Secret
      2. Pull detailed device information for the target device ID
      3. Collect the offboard reason entered as a NinjaOne Script Variable
      4. Create a new Organization-scoped custom field named after the device hostname
         (if it does not already exist)
      5. Write all collected data + offboard reason to that custom field on the org record

.NOTES
    ── HOW TO RUN THIS IN NINJAONE ──────────────────────────────────────────────
    1. Go to Administration > Scripting > Scripts > Add Script
    2. Paste this script in, set language to PowerShell
    3. Add two Script Variables (Administration > Scripting > Script Variables):
         - Name: targetDeviceId   | Type: Integer  | Label: "Target Device ID"
         - Name: offboardReason   | Type: Text     | Label: "Offboard Reason"
    4. Fill in the CONFIGURATION block below before saving
    5. Run the script against any managed device (it uses targetDeviceId, not the
       device it runs on — though it can run on the same device)

    ── FINDING YOUR DEVICE ID ────────────────────────────────────────────────────
    In NinjaOne, open the device. The URL will contain the device ID:
      https://app.ninjarmm.com/#/deviceDashboard/12345/overview
                                                   ^^^^^
    Or use the NinjaOne API:  GET /v2/devices  and look for the "id" field.

    ── API APP SETUP (one-time) ──────────────────────────────────────────────────
    Go to: Administration > Apps > API > Client App IDs > Add
      - Platform:      API Services (Machine-to-Machine)
      - Allowed Scopes: monitoring, management
      - Click Save, copy the Client ID and Client Secret
    No Redirect URI is needed for Client Credentials.
#>

# ==============================================================================
#  CONFIGURATION — Fill in ALL four values below before saving/running this script
# ==============================================================================

$BaseUrl      = 'https://<your Login URL>'          # e.g. https://app.ninjarmm.com
$TokenEndpoint = 'https://<your Login URL>/ws/oauth/token'  # same URL + /ws/oauth/token
$ClientId     = '<Your Client ID>'                  # From Administration > Apps > API
$ClientSecret = '<Your Client Secret>'              # From Administration > Apps > API

# ==============================================================================
#  DO NOT EDIT BELOW THIS LINE
# ==============================================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Read NinjaOne Script Variables ────────────────────────────────────────────
# These are populated automatically by NinjaOne when the script runs.
# They map to the Script Variables you create in Administration > Scripting.
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
Write-Host "  ============================================================" -ForegroundColor Cyan
Write-Host ""

# ── Step 1: Authenticate via Client Credentials ───────────────────────────────
Write-Host "  [1/5] Authenticating (Client Credentials)..." -ForegroundColor Cyan

$TokenBody = @{
    grant_type    = 'client_credentials'
    client_id     = $ClientId
    client_secret = $ClientSecret
    scope         = 'monitoring management'
}

try {
    $TokenResponse = Invoke-RestMethod `
        -Uri         $TokenEndpoint `
        -Method      POST `
        -Body        $TokenBody `
        -ContentType 'application/x-www-form-urlencoded'

    $AccessToken = $TokenResponse.access_token
} catch {
    Write-Host "  [!] Authentication failed." -ForegroundColor Red
    Write-Host "      Check BaseUrl, ClientId, ClientSecret, and that your API app" -ForegroundColor Yellow
    Write-Host "      has 'monitoring' and 'management' scopes enabled." -ForegroundColor Yellow
    Write-Host "      Error: $_" -ForegroundColor Red
    exit 1
}

$Headers = @{
    Authorization  = "Bearer $AccessToken"
    'Content-Type' = 'application/json'
    Accept         = 'application/json'
}

Write-Host "  [✓] Authenticated successfully." -ForegroundColor Green

# ── Step 2: Pull device information ───────────────────────────────────────────
Write-Host ""
Write-Host "  [2/5] Fetching device information for ID: $TargetDeviceId..." -ForegroundColor Cyan

try {
    $Device = Invoke-RestMethod `
        -Uri     "$BaseUrl/v2/device/$TargetDeviceId" `
        -Method  GET `
        -Headers $Headers
} catch {
    $Status = $_.Exception.Response.StatusCode.Value__
    if ($Status -eq 404) {
        Write-Host "  [!] Device ID $TargetDeviceId was not found in NinjaOne." -ForegroundColor Red
        Write-Host "      Verify the ID in the NinjaOne portal (check the device URL)." -ForegroundColor Yellow
    } else {
        Write-Host "  [!] Failed to retrieve device (HTTP $Status)." -ForegroundColor Red
        Write-Host "      Error: $_" -ForegroundColor Red
    }
    exit 1
}

# Pull extended device details (OS, hardware, last seen, etc.)
try {
    $DeviceDetails = Invoke-RestMethod `
        -Uri     "$BaseUrl/v2/device/$TargetDeviceId/system-info" `
        -Method  GET `
        -Headers $Headers
} catch {
    # system-info may not be available for all device types — non-fatal
    $DeviceDetails = $null
    Write-Host "  [i] Note: Extended system info not available for this device type." -ForegroundColor Gray
}

# Pull installed software
try {
    $SoftwareList = Invoke-RestMethod `
        -Uri     "$BaseUrl/v2/device/$TargetDeviceId/software" `
        -Method  GET `
        -Headers $Headers
} catch {
    $SoftwareList = $null
}

# Pull last logged-on user from activities
try {
    $Activities = Invoke-RestMethod `
        -Uri     "$BaseUrl/v2/device/$TargetDeviceId/activities?pageSize=10" `
        -Method  GET `
        -Headers $Headers
    $LastActivity = if ($Activities.activities) { $Activities.activities[0] } else { $null }
} catch {
    $LastActivity = $null
}

$Hostname   = if ($Device.systemName)  { $Device.systemName }  else { "Device-$TargetDeviceId" }
$OrgId      = $Device.organizationId
$LastSeen   = if ($Device.lastContact) {
    [DateTimeOffset]::FromUnixTimeMilliseconds($Device.lastContact).ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss")
} else { "Unknown" }

Write-Host "  [✓] Device found: $Hostname  (Org ID: $OrgId)" -ForegroundColor Green

# ── Step 3: Get organization information ──────────────────────────────────────
Write-Host ""
Write-Host "  [3/5] Fetching organization information (Org ID: $OrgId)..." -ForegroundColor Cyan

try {
    $Org = Invoke-RestMethod `
        -Uri     "$BaseUrl/v2/organization/$OrgId" `
        -Method  GET `
        -Headers $Headers
    $OrgName = $Org.name
} catch {
    $OrgName = "Unknown Organization"
    Write-Host "  [i] Could not retrieve org name — continuing." -ForegroundColor Gray
}

Write-Host "  [✓] Organization: $OrgName" -ForegroundColor Green

# ── Step 4: Create the custom field if it doesn't exist ───────────────────────
Write-Host ""
Write-Host "  [4/5] Checking/creating custom field for: $Hostname..." -ForegroundColor Cyan

# NinjaOne field names must be camelCase, no spaces or special chars
# Build a safe field name: strip non-alphanumeric, camelCase from hostname
$RawName   = ($Hostname -replace '[^a-zA-Z0-9]', ' ').Trim()
$Words     = $RawName -split '\s+'
$FieldName = ($Words[0].ToLower()) + (($Words[1..($Words.Count - 1)] | ForEach-Object {
    $_.Substring(0,1).ToUpper() + $_.Substring(1).ToLower()
}) -join '')

# Truncate to 50 chars (NinjaOne limit)
if ($FieldName.Length -gt 50) { $FieldName = $FieldName.Substring(0, 50) }
$FieldLabel = "Offboard - $Hostname"

Write-Host "  [i] Field name will be: $FieldName" -ForegroundColor Gray
Write-Host "  [i] Field label will be: $FieldLabel" -ForegroundColor Gray

# Check if field already exists
$FieldExists = $false
try {
    $ExistingFields = Invoke-RestMethod `
        -Uri     "$BaseUrl/v2/custom-fields?scope=organization" `
        -Method  GET `
        -Headers $Headers

    $ExistingField = $ExistingFields | Where-Object { $_.name -eq $FieldName }
    if ($ExistingField) {
        $FieldExists = $true
        Write-Host "  [i] Custom field '$FieldName' already exists — skipping creation." -ForegroundColor Gray
    }
} catch {
    Write-Host "  [i] Could not query existing fields — will attempt creation." -ForegroundColor Gray
}

if (-not $FieldExists) {
    # Create the organization-scoped WYSIWYG/multiline text custom field
    $NewFieldBody = @{
        name                 = $FieldName
        label                = $FieldLabel
        description          = "Offboard data captured for device: $Hostname (ID: $TargetDeviceId)"
        fieldType            = 'TEXT_MULTILINE'
        definitionScopes     = @('ORGANIZATION')
        technicianPermission = 'READ_ONLY'
        scriptPermission     = 'READ_WRITE'
        apiPermission        = 'READ_WRITE'
    } | ConvertTo-Json

    try {
        Invoke-RestMethod `
            -Uri     "$BaseUrl/v2/custom-fields" `
            -Method  POST `
            -Headers $Headers `
            -Body    $NewFieldBody | Out-Null

        Write-Host "  [✓] Custom field '$FieldName' created successfully." -ForegroundColor Green
        # Brief pause to allow NinjaOne to propagate the new field definition
        Start-Sleep -Seconds 3

    } catch {
        $Status = $_.Exception.Response.StatusCode.Value__
        if ($Status -eq 409) {
            # Race condition — field was created between our check and POST
            Write-Host "  [i] Field already existed (409 conflict) — continuing." -ForegroundColor Gray
        } else {
            Write-Host "  [!] Failed to create custom field (HTTP $Status)." -ForegroundColor Red
            Write-Host "      Ensure your API app has the 'management' scope." -ForegroundColor Yellow
            Write-Host "      Error: $_" -ForegroundColor Red
            exit 1
        }
    }
}

# ── Step 5: Build output and write to custom field ────────────────────────────
Write-Host ""
Write-Host "  [5/5] Writing offboard data to organization custom field..." -ForegroundColor Cyan

$Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# Gather software summary (top 20 installed apps)
$SoftwareSummary = "Not available"
if ($SoftwareList -and $SoftwareList.Count -gt 0) {
    $SoftwareSummary = ($SoftwareList | Select-Object -First 20 | ForEach-Object {
        "$($_.name) $($_.version)"
    }) -join "`n"
}

# Build the field value — plain text, structured for readability
$FieldValue = @"
========================================
 OFFBOARD CAPTURE REPORT
========================================
Captured      : $Timestamp
Device Name   : $Hostname
Device ID     : $TargetDeviceId
Organization  : $OrgName (ID: $OrgId)

----------------------------------------
 OFFBOARD REASON
----------------------------------------
$OffboardReason

----------------------------------------
 DEVICE DETAILS
----------------------------------------
System Name   : $($Device.systemName)
DNS Name      : $($Device.dnsName)
IP Address    : $($Device.ipAddresses -join ', ')
OS            : $($Device.os.name) $($Device.os.servicePack)
Device Class  : $($Device.nodeClass)
Last Contact  : $LastSeen
Agent Version : $($Device.agentVersion)
"@

# Append system info if available
if ($DeviceDetails) {
    $FieldValue += @"

----------------------------------------
 HARDWARE
----------------------------------------
Manufacturer  : $($DeviceDetails.system.manufacturer)
Model         : $($DeviceDetails.system.model)
Serial Number : $($DeviceDetails.bios.serialNumber)
CPU           : $($DeviceDetails.processors | Select-Object -First 1 -ExpandProperty name)
RAM (GB)      : $([math]::Round($DeviceDetails.memory.memoryMb / 1024, 1))
"@
}

# Append last activity if available
if ($LastActivity) {
    $ActivityTime = [DateTimeOffset]::FromUnixTimeMilliseconds($LastActivity.activityTime).ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss")
    $FieldValue += @"

----------------------------------------
 LAST RECORDED ACTIVITY
----------------------------------------
Time          : $ActivityTime
Type          : $($LastActivity.type)
Message       : $($LastActivity.message)
"@
}

$FieldValue += @"

----------------------------------------
 INSTALLED SOFTWARE (first 20)
----------------------------------------
$SoftwareSummary

========================================
 END OF REPORT
========================================
"@

# Write the value to the organization's custom field
$PatchBody = @{ $FieldName = $FieldValue } | ConvertTo-Json

try {
    Invoke-RestMethod `
        -Uri     "$BaseUrl/v2/organization/$OrgId/custom-fields" `
        -Method  PATCH `
        -Headers $Headers `
        -Body    $PatchBody | Out-Null
} catch {
    $Status = $_.Exception.Response.StatusCode.Value__
    Write-Host "  [!] Failed to write to custom field (HTTP $Status)." -ForegroundColor Red
    Write-Host "      Error: $_" -ForegroundColor Red
    Write-Host ""
    Write-Host "  Captured data (for manual copy):" -ForegroundColor Yellow
    Write-Host $FieldValue
    exit 1
}

Write-Host "  [✓] Data written to organization custom field: $FieldName" -ForegroundColor Green

# ── Summary ───────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ============================================================" -ForegroundColor Green
Write-Host "  [✓] OFFBOARD CAPTURE COMPLETE" -ForegroundColor Green
Write-Host "      Device     : $Hostname (ID: $TargetDeviceId)" -ForegroundColor Green
Write-Host "      Org        : $OrgName (ID: $OrgId)" -ForegroundColor Green
Write-Host "      Field      : $FieldName" -ForegroundColor Green
Write-Host "      Timestamp  : $Timestamp" -ForegroundColor Green
Write-Host "  ============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  To view the data:" -ForegroundColor Cyan
Write-Host "  NinjaOne > Organizations > $OrgName > Custom Fields > $FieldLabel" -ForegroundColor Cyan
Write-Host ""
