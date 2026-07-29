#Requires -Version 5.1
<#
.SYNOPSIS
    Pulls the software inventory for EVERY device across the entire NinjaOne
    Oceania (OC) tenant and exports the results to a single CSV file.

.DESCRIPTION
    Uses the NinjaOne Public API v2 with Client Credentials (silent, no browser
    login required). Hard-coded to the Oceania instance: https://oc.ninjarmm.com

    This is the tenant-wide version of the single-device script. Instead of
    looping per-device (slow, one API call per machine), it uses NinjaOne's
    bulk "queries" endpoint, which returns software records for the whole
    tenant in a handful of paginated calls.

    API calls made:
      GET /v2/devices           — paginated list of every device (with
                                   organization + location expanded inline)
      GET /v2/queries/software  — paginated, tenant-wide software inventory
                                   (one record per app per device)

    Output columns (CSV):
      Publisher | SoftwareName | Version | OSType | InstallDate | Hostname |
      OrganizationName | LocationName | DeviceId

    ── HOW TO SCOPE THE EXPORT (OPTIONAL) ───────────────────────────────────────
    By default this pulls software for ALL devices in the tenant. To narrow it
    to one organization, set $OrganizationFilterId in the CONFIGURATION block
    to that org's ID (find it by opening the organization in NinjaOne and
    checking the URL, e.g. .../organization/7/... -> 7). Leave it at 0 to
    include every organization.

    ── A NOTE ON FIRST INSTALL vs LAST INSTALL ───────────────────────────────────
    The queries/software endpoint returns one record per installed application
    with a single installDate field. The NinjaOne API does not store separate
    first-install and last-install timestamps. InstallDate reflects what the
    OS reported at time of collection.

    ── A NOTE ON MISSING DATA ────────────────────────────────────────────────────
    Not every device reports every field (this varies by OS and device type).
    Any field NinjaOne didn't report is left blank in the CSV rather than
    guessed or filled in — do not assume a blank means "none installed" or
    "not applicable", it means the API didn't return a value for that field.

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

# Optional: restrict the export to a single organization by ID.
# Leave at 0 to export software for every organization in the tenant.
$OrganizationFilterId = 0   # e.g. 7

# Page size used for both device and software list calls. 1000 is a safe
# default well within NinjaOne's allowed range.
$PageSize = 1000

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
# Returns $Default (default: $null) if the property is missing, null, or blank
# — never throws, and never invents a value that wasn't actually returned.
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

Write-Host ""
Write-Host "  ============================================================" -ForegroundColor Cyan
Write-Host "  NinjaOne Software Inventory — Tenant-Wide" -ForegroundColor Cyan
Write-Host "  Instance : oc.ninjarmm.com (Oceania)" -ForegroundColor Cyan
if ($OrganizationFilterId -gt 0) {
    Write-Host "  Scope    : Organization ID $OrganizationFilterId only" -ForegroundColor Cyan
} else {
    Write-Host "  Scope    : Entire tenant (all organizations)" -ForegroundColor Cyan
}
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

# ── Step 2: Fetch every device in the tenant (paginated) ──────────────────────
# We ask NinjaOne to expand organization + location inline so we don't need a
# separate lookup call per org — one pass over /v2/devices gets us hostname,
# OS type, org name, and location name for every machine.
Write-Host ""
Write-Host "  [2/4] Fetching device list (this may take a few pages on large tenants)..." -ForegroundColor Cyan

$DeviceLookup = @{}   # deviceId -> pscustomobject { Hostname, OSType, OrgName, LocationName }
$AfterId      = 0
$TotalDevices = 0

do {
    $uri = "$BaseUrl/v2/devices?pageSize=$PageSize&after=$AfterId&expand=organization,location"
    if ($OrganizationFilterId -gt 0) {
        $filter = [uri]::EscapeDataString("org = $OrganizationFilterId")
        $uri += "&df=$filter"
    }

    try {
        $Page = Invoke-RestMethod -Uri $uri -Method GET -Headers $Headers
    } catch {
        Write-Host "  [!] Failed to fetch device list: $_" -ForegroundColor Red
        exit 1
    }

    $PageDevices = if ($Page -is [array]) { $Page } else { @($Page) }
    if ($PageDevices.Count -eq 0) { break }

    foreach ($d in $PageDevices) {
        $devId = Get-Prop -Obj $d -Name 'id'
        if ($null -eq $devId) { continue }

        $orgObj = Get-Prop -Obj $d -Name 'organization'
        $locObj = Get-Prop -Obj $d -Name 'location'

        $DeviceLookup[[string]$devId] = [pscustomobject]@{
            Hostname     = Get-Prop -Obj $d -Name 'systemName' -Default "Device $devId"
            OSType       = Get-Prop -Obj $d -Name 'nodeClass'
            OrgName      = if ($orgObj) { Get-Prop -Obj $orgObj -Name 'name' } else { $null }
            LocationName = if ($locObj) { Get-Prop -Obj $locObj -Name 'name' } else { $null }
        }
        $TotalDevices++
        $AfterId = $devId
    }

    Write-Host "      ...$TotalDevices device(s) so far" -ForegroundColor Gray

} while ($PageDevices.Count -ge $PageSize)

Write-Host "  [✓] $TotalDevices device(s) indexed." -ForegroundColor Green

if ($TotalDevices -eq 0) {
    Write-Host "  [!] No devices found — nothing to export." -ForegroundColor Red
    exit 1
}

# ── Step 3: Fetch the full tenant-wide software inventory (cursor-paginated) ──
Write-Host ""
Write-Host "  [3/4] Fetching software inventory for the whole tenant..." -ForegroundColor Cyan

$SoftwareList = [System.Collections.Generic.List[object]]::new()
$Cursor       = $null

do {
    $uri = "$BaseUrl/v2/queries/software?pageSize=$PageSize"
    if ($OrganizationFilterId -gt 0) {
        $filter = [uri]::EscapeDataString("org = $OrganizationFilterId")
        $uri += "&df=$filter"
    }
    if ($Cursor) {
        $uri += "&cursor=$([uri]::EscapeDataString($Cursor))"
    }

    try {
        $Response = Invoke-RestMethod -Uri $uri -Method GET -Headers $Headers
    } catch {
        Write-Host "  [!] Failed to fetch software inventory: $_" -ForegroundColor Red
        exit 1
    }

    $Results = Get-Prop -Obj $Response -Name 'results' -Default @()
    foreach ($r in $Results) { $SoftwareList.Add($r) }

    Write-Host "      ...$($SoftwareList.Count) software record(s) so far" -ForegroundColor Gray

    $CursorObj = Get-Prop -Obj $Response -Name 'cursor'
    $Cursor    = if ($CursorObj) { Get-Prop -Obj $CursorObj -Name 'name' } else { $null }

} while ($Cursor -and $Results.Count -gt 0)

Write-Host "  [✓] $($SoftwareList.Count) total software record(s) retrieved." -ForegroundColor Green

# ── Step 4: Join software records to device info and export ───────────────────
Write-Host ""
Write-Host "  [4/4] Building output and exporting to CSV..." -ForegroundColor Cyan

$UnmatchedDeviceIds = [System.Collections.Generic.HashSet[string]]::new()

$Output = $SoftwareList | ForEach-Object {
    $sw = $_

    $devId = Get-Prop -Obj $sw -Name 'deviceId'
    $devInfo = if ($null -ne $devId -and $DeviceLookup.ContainsKey([string]$devId)) {
        $DeviceLookup[[string]$devId]
    } else {
        if ($null -ne $devId) { [void]$UnmatchedDeviceIds.Add([string]$devId) }
        $null
    }

    $RawDate     = Get-Prop -Obj $sw -Name 'installDate'
    $InstallDate = if ($RawDate) {
        try { ([datetime]$RawDate).ToString('yyyy-MM-dd') } catch { $RawDate }
    } else { $null }

    [PSCustomObject]@{
        Publisher        = Get-Prop -Obj $sw -Name 'publisher'
        SoftwareName     = Get-Prop -Obj $sw -Name 'name'
        Version          = Get-Prop -Obj $sw -Name 'version'
        OSType           = if ($devInfo) { $devInfo.OSType } else { $null }
        InstallDate      = $InstallDate
        Hostname         = if ($devInfo) { $devInfo.Hostname } else { $null }
        OrganizationName = if ($devInfo) { $devInfo.OrgName } else { $null }
        LocationName     = if ($devInfo) { $devInfo.LocationName } else { $null }
        DeviceId         = $devId
    }
} | Sort-Object OrganizationName, Hostname, SoftwareName

if ($UnmatchedDeviceIds.Count -gt 0) {
    Write-Host "  [i] $($UnmatchedDeviceIds.Count) software record(s) referenced a device ID not found in the device list (left blank)." -ForegroundColor Gray
}

# ── Console preview ───────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  Preview (first 25 records):" -ForegroundColor Cyan
$Output | Select-Object -First 25 |
    Format-Table Publisher, SoftwareName, Version, OSType, InstallDate, Hostname, OrganizationName, LocationName -AutoSize

# ── CSV export ────────────────────────────────────────────────────────────────
$Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$ScopeTag  = if ($OrganizationFilterId -gt 0) { "Org${OrganizationFilterId}" } else { "AllOrgs" }
$CsvPath   = "C:\Windows\Temp\NinjaSoftware_${ScopeTag}_$Timestamp.csv"

$Output | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "  ============================================================" -ForegroundColor Green
Write-Host "  [✓] COMPLETE" -ForegroundColor Green
Write-Host "      Devices covered : $TotalDevices" -ForegroundColor Green
Write-Host "      Records         : $($Output.Count) software records" -ForegroundColor Green
Write-Host "      CSV saved       : $CsvPath" -ForegroundColor Green
Write-Host "  ============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  NOTE: InstallDate is the OS-reported install date per record." -ForegroundColor Yellow
Write-Host "  The NinjaOne API does not provide separate first/last install" -ForegroundColor Yellow
Write-Host "  timestamps — only a single installDate per application." -ForegroundColor Yellow
Write-Host "  Any blank field means NinjaOne did not report a value for it" -ForegroundColor Yellow
Write-Host "  — it is not filled in or guessed." -ForegroundColor Yellow
Write-Host ""
