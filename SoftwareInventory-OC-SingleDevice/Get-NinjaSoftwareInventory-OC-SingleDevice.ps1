#Requires -Version 5.1
<#
.SYNOPSIS
    Pulls the software inventory for a single device from the NinjaOne Oceania
    (OC) instance and exports the results to a CSV file.

.DESCRIPTION
    Uses the NinjaOne Public API v2 with Client Credentials (silent, no browser
    login required). Hard-coded to the Oceania instance: https://oc.ninjarmm.com

    API calls made:
      GET /v2/device/{id}                    — hostname, OS type, locationId, orgId
      GET /v2/device/{id}/software           — full software list for this device
      GET /v2/organization/{id}/locations    — resolve locationId to a location name

    Output columns (CSV):
      Publisher | SoftwareName | Version | OSType | InstallDate | Hostname | LocationName

    ── HOW TO PROVIDE THE DEVICE ID ─────────────────────────────────────────────
    Three ways — the script checks them in this order:

      1. NinjaOne Script Variable  : Create a Script Variable named 'targetDeviceId'
                                     (Type: Integer). NinjaOne injects it as an
                                     environment variable automatically when run
                                     as an automation.

      2. Config block below        : Set $ManualDeviceId to the device ID number.
                                     Use this when running the script manually from
                                     PowerShell outside of NinjaOne.

      3. Command-line parameter    : Pass -DeviceId 12345 when calling the script
                                     directly from a PowerShell prompt.

    ── FINDING THE DEVICE ID ────────────────────────────────────────────────────
    Open the device in NinjaOne and look at the browser URL:
      https://oc.ninjarmm.com/#/deviceDashboard/12345/overview
                                                 ^^^^^
    That number is the Device ID.

    ── A NOTE ON FIRST INSTALL vs LAST INSTALL ───────────────────────────────────
    The /v2/device/{id}/software endpoint returns one record per installed
    application with a single installDate field. The NinjaOne API does not store
    separate first-install and last-install timestamps. The InstallDate column
    reflects what the OS reported at time of collection.

    ── API APP SETUP (one-time) ──────────────────────────────────────────────────
    Go to: Administration > Apps > API > Client App IDs > Add
      Platform      : API Services (Machine-to-Machine)
      Allowed Scopes: monitoring
      Redirect URI  : Leave blank
    Click Save. Copy the Client ID and Client Secret.

    ── REGIONAL NOTE ────────────────────────────────────────────────────────────
    This script is hard-coded to the Oceania (OC) instance: https://oc.ninjarmm.com
    Do not change the BaseUrl or TokenEndpoint unless moving to a different region.
#>

[CmdletBinding()]
param(
    # Optional: pass the Device ID directly as a command-line parameter
    # e.g.  .\Get-NinjaSoftwareInventory-OC-SingleDevice.ps1 -DeviceId 12345
    [Parameter(Mandatory = $false)]
    [int]$DeviceId = 0
)

# ==============================================================================
#  CONFIGURATION — Fill in your Client ID and Secret below
# ==============================================================================

# Hard-coded to Oceania instance — do not change
$BaseUrl       = 'https://oc.ninjarmm.com'
$TokenEndpoint = 'https://oc.ninjarmm.com/ws/oauth/token'

# From Administration > Apps > API > Client App IDs
$ClientId      = '<Your Client ID>'

# From Administration > Apps > API > Client App IDs (shown once at creation)
$ClientSecret  = '<Your Client Secret>'

# Set this when running manually outside of NinjaOne.
# Leave as 0 if using a NinjaOne Script Variable or the -DeviceId parameter.
$ManualDeviceId = 0   # e.g. 12345

# ==============================================================================
#  DO NOT EDIT BELOW THIS LINE
# ==============================================================================

# Note: Set-StrictMode is intentionally NOT used.
# NinjaOne API responses are dynamic PSObjects whose available properties vary
# by device type. StrictMode throws errors when a property is absent rather
# than returning $null, which breaks on devices that don't report every field.
$ErrorActionPreference = 'Stop'

# ── Safe property helper ──────────────────────────────────────────────────────
# Safely reads a property from an API response PSObject.
# Returns $Default if the property is missing or null — never throws.
function Get-Prop {
    param(
        [object]$Obj,
        [string]$Name,
        [object]$Default = $null
    )
    if ($null -eq $Obj) { return $Default }
    $prop = $Obj.PSObject.Properties[$Name]
    if ($null -eq $prop -or $null -eq $prop.Value -or $prop.Value -eq '') {
        return $Default
    }
    return $prop.Value
}

# ── Validate credentials ──────────────────────────────────────────────────────
if ($ClientId -like '*<*' -or $ClientSecret -like '*<*') {
    Write-Error "[!] Please fill in ClientId and ClientSecret in the CONFIGURATION block."
    exit 1
}

# ── Resolve device ID — checks three sources in priority order ────────────────
#   1. NinjaOne Script Variable ($env:targetDeviceId)
#   2. -DeviceId command-line parameter
#   3. $ManualDeviceId in the config block
$ResolvedDeviceId = 0

if ($env:targetDeviceId -and [int]::TryParse($env:targetDeviceId, [ref]$null)) {
    $ResolvedDeviceId = [int]$env:targetDeviceId
    Write-Host "  [i] Using device ID from NinjaOne Script Variable: $ResolvedDeviceId" -ForegroundColor Gray
} elseif ($DeviceId -gt 0) {
    $ResolvedDeviceId = $DeviceId
    Write-Host "  [i] Using device ID from -DeviceId parameter: $ResolvedDeviceId" -ForegroundColor Gray
} elseif ($ManualDeviceId -gt 0) {
    $ResolvedDeviceId = $ManualDeviceId
    Write-Host "  [i] Using device ID from config block: $ResolvedDeviceId" -ForegroundColor Gray
} else {
    Write-Error @"
[!] No device ID provided. Supply one of the following:
    1. NinjaOne Script Variable named 'targetDeviceId' (when running as automation)
    2. -DeviceId parameter: .\Get-NinjaSoftwareInventory-OC-SingleDevice.ps1 -DeviceId 12345
    3. Set `$ManualDeviceId = 12345 in the CONFIGURATION block

To find a Device ID: open the device in NinjaOne and check the URL:
  https://oc.ninjarmm.com/#/deviceDashboard/12345/overview  <- that number
"@
    exit 1
}

Write-Host ""
Write-Host "  ============================================================" -ForegroundColor Cyan
Write-Host "  NinjaOne Software Inventory — Single Device" -ForegroundColor Cyan
Write-Host "  Instance  : oc.ninjarmm.com (Oceania)" -ForegroundColor Cyan
Write-Host "  Device ID : $ResolvedDeviceId" -ForegroundColor Cyan
Write-Host "  ============================================================" -ForegroundColor Cyan
Write-Host ""

# ── Step 1: Authenticate ──────────────────────────────────────────────────────
Write-Host "  [1/4] Authenticating (Client Credentials)..." -ForegroundColor Cyan
try {
    $TokenResponse = Invoke-RestMethod `
        -Uri         $TokenEndpoint `
        -Method      POST `
        -ContentType 'application/x-www-form-urlencoded' `
        -Body        @{
            grant_type    = 'client_credentials'
            client_id     = $ClientId
            client_secret = $ClientSecret
            scope         = 'monitoring'
        }
    $AccessToken = $TokenResponse.access_token
} catch {
    Write-Host "  [!] Authentication failed." -ForegroundColor Red
    Write-Host "      Verify ClientId, ClientSecret, and that the API app" -ForegroundColor Yellow
    Write-Host "      platform is 'API Services (Machine-to-Machine)' with" -ForegroundColor Yellow
    Write-Host "      the monitoring scope enabled." -ForegroundColor Yellow
    Write-Host "      Error: $_" -ForegroundColor Red
    exit 1
}

$Headers = @{
    Authorization = "Bearer $AccessToken"
    Accept        = 'application/json'
}
Write-Host "  [✓] Authenticated." -ForegroundColor Green

# ── Step 2: Fetch device info (hostname, OS type, locationId, orgId) ──────────
Write-Host ""
Write-Host "  [2/4] Fetching device info (ID: $ResolvedDeviceId)..." -ForegroundColor Cyan
try {
    $Device = Invoke-RestMethod `
        -Uri     "$BaseUrl/v2/device/$ResolvedDeviceId" `
        -Method  GET `
        -Headers $Headers
} catch {
    $Status = $null
    try { $Status = $_.Exception.Response.StatusCode.Value__ } catch {}
    if ($Status -eq 404) {
        Write-Host "  [!] Device ID $ResolvedDeviceId was not found in NinjaOne." -ForegroundColor Red
        Write-Host "      Verify the ID from the device URL in the OC portal." -ForegroundColor Yellow
    } else {
        Write-Host "  [!] Failed to fetch device (HTTP $Status): $_" -ForegroundColor Red
    }
    exit 1
}

$Hostname  = Get-Prop -Obj $Device -Name 'systemName'     -Default "Device $ResolvedDeviceId"
$OSType    = Get-Prop -Obj $Device -Name 'nodeClass'      -Default 'Unknown'
$LocationId = Get-Prop -Obj $Device -Name 'locationId'    -Default $null
$OrgId     = Get-Prop -Obj $Device -Name 'organizationId' -Default $null

Write-Host "  [✓] Device : $Hostname" -ForegroundColor Green
Write-Host "      OS Type: $OSType" -ForegroundColor Green
Write-Host "      Org ID : $OrgId" -ForegroundColor Green

# ── Resolve location name ─────────────────────────────────────────────────────
$LocationName = 'Unknown'
if ($OrgId -and $LocationId) {
    try {
        $Locations = Invoke-RestMethod `
            -Uri     "$BaseUrl/v2/organization/$OrgId/locations" `
            -Method  GET `
            -Headers $Headers
        foreach ($Loc in $Locations) {
            $locId = Get-Prop -Obj $Loc -Name 'id'
            if ($locId -eq $LocationId) {
                $LocationName = Get-Prop -Obj $Loc -Name 'name' -Default "Location $LocationId"
                break
            }
        }
    } catch {
        Write-Host "  [i] Could not resolve location name — continuing." -ForegroundColor Gray
    }
}
Write-Host "      Location: $LocationName" -ForegroundColor Green

# ── Step 3: Fetch software for this device ────────────────────────────────────
Write-Host ""
Write-Host "  [3/4] Fetching software inventory for $Hostname..." -ForegroundColor Cyan
try {
    $SoftwareRaw = Invoke-RestMethod `
        -Uri     "$BaseUrl/v2/device/$ResolvedDeviceId/software" `
        -Method  GET `
        -Headers $Headers
} catch {
    $Status = $null
    try { $Status = $_.Exception.Response.StatusCode.Value__ } catch {}
    Write-Host "  [!] Failed to fetch software list (HTTP $Status): $_" -ForegroundColor Red
    Write-Host "      Note: Some device types (NMS, cloud monitors) do not" -ForegroundColor Yellow
    Write-Host "      report a software inventory to NinjaOne." -ForegroundColor Yellow
    exit 1
}

# Normalise — the endpoint returns an array directly
$SoftwareList = if ($SoftwareRaw -is [array]) { $SoftwareRaw } else { @($SoftwareRaw) }
Write-Host "  [✓] $($SoftwareList.Count) software record(s) found." -ForegroundColor Green

# ── Step 4: Build output and export ───────────────────────────────────────────
Write-Host ""
Write-Host "  [4/4] Building output and exporting to CSV..." -ForegroundColor Cyan

$Output = $SoftwareList | ForEach-Object {
    $sw = $_

    $RawDate     = Get-Prop -Obj $sw -Name 'installDate'
    $InstallDate = if ($RawDate) {
        try { ([datetime]$RawDate).ToString('yyyy-MM-dd') } catch { $RawDate }
    } else { '' }

    [PSCustomObject]@{
        Publisher    = Get-Prop -Obj $sw -Name 'publisher' -Default ''
        SoftwareName = Get-Prop -Obj $sw -Name 'name'      -Default ''
        Version      = Get-Prop -Obj $sw -Name 'version'   -Default ''
        OSType       = $OSType
        InstallDate  = $InstallDate
        Hostname     = $Hostname
        LocationName = $LocationName
    }
} | Sort-Object SoftwareName

# ── Console preview ───────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  Preview (first 25 records):" -ForegroundColor Cyan
$Output | Select-Object -First 25 |
    Format-Table Publisher, SoftwareName, Version, OSType, InstallDate, Hostname, LocationName -AutoSize

# ── CSV export ────────────────────────────────────────────────────────────────
$Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$SafeName  = $Hostname -replace '[^a-zA-Z0-9_-]', '_'
$CsvPath   = "C:\Windows\Temp\NinjaSoftware_${SafeName}_$Timestamp.csv"

$Output | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "  ============================================================" -ForegroundColor Green
Write-Host "  [✓] COMPLETE" -ForegroundColor Green
Write-Host "      Device    : $Hostname (ID: $ResolvedDeviceId)" -ForegroundColor Green
Write-Host "      OS Type   : $OSType" -ForegroundColor Green
Write-Host "      Location  : $LocationName" -ForegroundColor Green
Write-Host "      Records   : $($Output.Count) software records" -ForegroundColor Green
Write-Host "      CSV saved : $CsvPath" -ForegroundColor Green
Write-Host "  ============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  NOTE: InstallDate is the OS-reported install date per record." -ForegroundColor Yellow
Write-Host "  The NinjaOne API does not provide separate first/last install" -ForegroundColor Yellow
Write-Host "  timestamps — only a single installDate per application." -ForegroundColor Yellow
Write-Host ""
