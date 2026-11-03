#Requires -Version 5.1
<#
.SYNOPSIS
    Reads the warranty start date from a NinjaOne device and writes it to the
    ITAM Purchase Date field on the same device.

.DESCRIPTION
    Uses the NinjaOne Public API v2 with Client Credentials (silent, no browser).

    For a given device ID the script:
      1. Fetches the device warranty using GET /v2/device/{id}?expand=warranty
      2. Reads warranty.startDate (Unix epoch seconds)
      3. Converts to milliseconds (required by the ITAM field)
      4. Writes to the itamAssetPurchaseDate field via
         PATCH /v2/device/{id}/custom-fields

    Confirmed endpoints and field names from NinjaOne official documentation:
      GET  /v2/device/{id}?expand=warranty
           Returns: warranty.startDate (epoch seconds)
      PATCH /v2/device/{id}/custom-fields
           Field:   itamAssetPurchaseDate (epoch milliseconds)

    API APP SETUP (one-time):
    Administration > Apps > API > Client App IDs > Add
      Platform      : API Services (Machine-to-Machine)
      Allowed Scopes: monitoring  AND  management
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

# The device ID to update
# Find this in NinjaOne by opening the device and reading the number from the URL:
#   https://app.ninjarmm.com/#/deviceDashboard/3/overview
#                                              ^ this number
$DeviceId      = 0   # <-- Replace with your device ID

# ==============================================================================
#  DO NOT EDIT BELOW THIS LINE
# ==============================================================================

try { [Net.ServicePointManager]::SecurityProtocol =
    [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch {}
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
$ErrorActionPreference = 'Stop'

# -- Config validation ---------------------------------------------------------
$ConfigErrors = @()
if ($BaseUrl      -like '*<*') { $ConfigErrors += 'Fill in $BaseUrl' }
if ($TokenEndpoint -like '*<*') { $ConfigErrors += 'Fill in $TokenEndpoint' }
if ($ClientId     -like '*<*') { $ConfigErrors += 'Fill in $ClientId' }
if ($ClientSecret -like '*<*') { $ConfigErrors += 'Fill in $ClientSecret' }
if ($DeviceId     -eq 0)       { $ConfigErrors += 'Fill in $DeviceId' }
if ($ConfigErrors.Count -gt 0) {
    Write-Host ''
    Write-Host '  [!] Configuration errors:' -ForegroundColor Red
    $ConfigErrors | ForEach-Object { Write-Host "      - $_" -ForegroundColor Red }
    exit 1
}

# -- Helper -------------------------------------------------------------------
function Get-Prop {
    param([object]$Obj, [string]$Name, [object]$Default = $null)
    if ($null -eq $Obj) { return $Default }
    $p = $Obj.PSObject.Properties[$Name]
    if ($null -eq $p -or $null -eq $p.Value) { return $Default }
    return $p.Value
}

Write-Host ''
Write-Host '  ================================================================' -ForegroundColor Cyan
Write-Host '  NinjaOne -- Write Warranty Start Date to ITAM Purchase Date'     -ForegroundColor Cyan
Write-Host "  Device ID: $DeviceId"                                            -ForegroundColor Cyan
Write-Host '  ================================================================' -ForegroundColor Cyan
Write-Host ''

# =============================================================================
#  STEP 1: Authenticate
# =============================================================================
Write-Host '  [1/4] Authenticating...' -ForegroundColor Cyan
$Headers = $null
try {
    $Token = Invoke-RestMethod -Uri $TokenEndpoint -Method POST `
        -ContentType 'application/x-www-form-urlencoded' `
        -Body "grant_type=client_credentials&client_id=$ClientId&client_secret=$ClientSecret&scope=monitoring management"
    $Headers = @{ Authorization = "Bearer $($Token.access_token)"; Accept = 'application/json' }
} catch {
    Write-Host "  [!] Authentication failed: $_" -ForegroundColor Red
    exit 1
}
Write-Host '  [OK] Authenticated.' -ForegroundColor Green

# =============================================================================
#  STEP 2: Fetch device and warranty
# =============================================================================
Write-Host ''; Write-Host "  [2/4] Fetching device $DeviceId with warranty data..." -ForegroundColor Cyan

$Device = $null
try {
    $Device = Invoke-RestMethod `
        -Uri "$BaseUrl/v2/device/$DeviceId`?expand=warranty" `
        -Method GET -Headers $Headers -ErrorAction Stop
} catch {
    $sc = $null; try { $sc = [int]$_.Exception.Response.StatusCode } catch {}
    if ($sc -eq 404) {
        Write-Host "  [!] Device $DeviceId not found. Check the ID and try again." -ForegroundColor Red
    } else {
        Write-Host "  [!] Failed to fetch device (HTTP $sc): $_" -ForegroundColor Red
    }
    exit 1
}

$DeviceName    = Get-Prop $Device 'systemName' -Default (Get-Prop $Device 'dnsName' -Default "Device $DeviceId")
$DeviceOrg     = Get-Prop $Device 'organizationId'
$WarrantyObj   = Get-Prop $Device 'warranty'
$WarrantyStart = if ($WarrantyObj) { Get-Prop $WarrantyObj 'startDate' } else { $null }

Write-Host "  [OK] Device found:" -ForegroundColor Green
Write-Host "       Name      : $DeviceName"
Write-Host "       Org ID    : $DeviceOrg"

if (-not $WarrantyObj -or $null -eq $WarrantyStart -or $WarrantyStart -eq 0) {
    Write-Host "  [!] No warranty start date found on device $DeviceId." -ForegroundColor Red
    Write-Host "      Ensure Warranty Tracking is enabled and the device has a warranty start date set." -ForegroundColor Yellow
    Write-Host "      In NinjaOne: device Overview tab > Warranty date > set start date or Sync with manufacturer." -ForegroundColor Yellow
    exit 1
}

# Convert epoch seconds to a readable date for display
$WarrantyStartDate = [DateTimeOffset]::FromUnixTimeSeconds([long]$WarrantyStart).ToLocalTime()
Write-Host "       Warranty start : $($WarrantyStartDate.ToString('yyyy-MM-dd')) (epoch: $WarrantyStart seconds)"

# =============================================================================
#  STEP 3: Convert and confirm
# =============================================================================
Write-Host ''; Write-Host '  [3/4] Converting warranty start date for ITAM field...' -ForegroundColor Cyan

# itamAssetPurchaseDate requires epoch milliseconds -- warranty.startDate is seconds
$PurchaseDateMs = [long]$WarrantyStart * 1000

Write-Host "  [OK] Warranty start  : $($WarrantyStartDate.ToString('yyyy-MM-dd'))"
Write-Host "       ITAM value (ms) : $PurchaseDateMs"

# =============================================================================
#  STEP 4: Write to ITAM Purchase Date field
# =============================================================================
Write-Host ''; Write-Host "  [4/4] Writing Purchase Date to device $DeviceId..." -ForegroundColor Cyan

# Body is a flat object keyed by the ITAM internal field name
# Confirmed from NinjaOne ITAM API docs:
#   itamAssetPurchaseDate = Purchase Date field (epoch milliseconds)
$Body    = "{`"itamAssetPurchaseDate`":$PurchaseDateMs}"
$PatchH  = $Headers.Clone()
$PatchH['Content-Type'] = 'application/json'

try {
    Invoke-RestMethod `
        -Uri "$BaseUrl/v2/device/$DeviceId/custom-fields" `
        -Method PATCH -Headers $PatchH -Body $Body -ErrorAction Stop | Out-Null

    Write-Host "  [OK] Purchase Date written successfully." -ForegroundColor Green
    Write-Host "       Device    : $DeviceName (ID: $DeviceId)"
    Write-Host "       Field     : itamAssetPurchaseDate"
    Write-Host "       Value set : $($WarrantyStartDate.ToString('yyyy-MM-dd'))"
} catch {
    $sc     = $null; try { $sc = [int]$_.Exception.Response.StatusCode }     catch {}
    $detail = '';    try { $detail = $_.ErrorDetails.Message }                catch {}
    Write-Host "  [!] Failed to write Purchase Date (HTTP $sc): $_" -ForegroundColor Red
    if ($detail) { Write-Host "      API detail: $detail" -ForegroundColor Red }
    Write-Host "      Check: management scope on API app, ITAM is enabled, device has API permission on itamAssetPurchaseDate field." -ForegroundColor Yellow
    exit 1
}

Write-Host ''
Write-Host '  ================================================================' -ForegroundColor Green
Write-Host '  [OK] COMPLETE'                                                    -ForegroundColor Green
Write-Host "       Device    : $DeviceName (ID: $DeviceId)"                    -ForegroundColor Green
Write-Host "       Purchase Date set to : $($WarrantyStartDate.ToString('yyyy-MM-dd'))" -ForegroundColor Green
Write-Host '  ================================================================' -ForegroundColor Green
Write-Host ''
