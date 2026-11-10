# ==============================================================================
#  Tenable.sc -> NinjaOne — Vulnerability Scan Group Upload
# ==============================================================================
#
#  WHAT THIS SCRIPT DOES:
#
#  1. Logs in to NinjaOne and finds your Scan Group automatically (by name).
#  2. Connects to your Tenable.sc instance and pulls vulnerability scan data.
#  3. Builds a CSV file using the exact column headings your Scan Group expects.
#  4. Uploads the CSV to the Scan Group and waits for NinjaOne to finish importing.
#
#  The first time you run this script it will open your browser so you can log
#  in to NinjaOne. After that it saves a refresh token so you never have to
#  log in again — it renews its own access automatically on every run.
#
#  REQUIREMENTS:
#
#  PowerShell 7 or newer. (Check with:  $PSVersionTable.PSVersion )
#
#  BEFORE YOU RUN THIS SCRIPT THE FIRST TIME:
#
#  1. Log in to NinjaOne and go to:
#        Administration > Apps > API > Client App
#     Create or open an app and make sure the Redirect URI is set to exactly:
#        https://localhost
#     Also make sure Refresh Token (offline_access) is enabled on the app.
#
#  2. In Tenable.sc, generate an API key for your account:
#        Username menu (top right) > Profile > API Keys > Generate
#     Copy both the Access Key and Secret Key.
#
#  3. In NinjaOne, create the Scan Group ONE TIME by hand. NinjaOne does not
#     allow a Scan Group to be created through the API.
#        - Enable the Tenable (or other vendor) Vulnerability Importer app.
#        - Open its "Scan groups" tab and click "Create scan group".
#        - Give it a name and choose the scanning vendor.
#        - Upload a CSV file when asked, and map the columns.
#     TIP: If you don't have a CSV yet, run this script once. If it cannot
#     find your Scan Group, it still builds a starter CSV file you can use
#     for this one-time setup, and tells you exactly what to do.
#
#  4. Fill in all values in the CONFIGURATION section below.
#
# ==============================================================================
#  CONFIGURATION — only edit values in this section
# ==============================================================================

# --- NinjaOne credentials -----------------------------------------------------

$NinjaBaseUrl       = 'https://<your NinjaOne login URL>'        # e.g. https://app.ninjarmm.com
$NinjaTokenEndpoint = 'https://<your NinjaOne login URL>/ws/oauth/token'
$NinjaClientId      = '<Your NinjaOne Client ID>'
$NinjaClientSecret  = '<Your NinjaOne Client Secret>'

# The NAME of the Scan Group in NinjaOne (exactly as shown on the Scan groups tab).
# The script looks up the Scan Group ID for you.
# Leave this blank ('') only if you have exactly ONE Scan Group in NinjaOne.
$NinjaScanGroupName = '<Your Scan Group Name>'                   # e.g. Tenable - All Servers

# --- Tenable.sc credentials ---------------------------------------------------

$TenableBaseUrl     = 'https://<your Tenable.sc hostname or IP>' # e.g. https://tenable.company.com
$TenableAccessKey   = '<Your Tenable.sc Access Key>'
$TenableSecretKey   = '<Your Tenable.sc Secret Key>'

# --- Scan settings ------------------------------------------------------------

# The Tenable.sc Repository ID to pull vulnerabilities from.
# In Tenable.sc go to: Repositories — the ID is shown in the repository list.
$TenableRepositoryId = '<Your Repository ID>'                    # e.g. 1

# How devices are identified in the CSV.
#   'Auto'       - (recommended) use whatever your NinjaOne Scan Group is set up to use.
#   'Hostname'   - force the Tenable.sc DNS name
#   'IPAddress'  - force the Tenable.sc IP address
#   'MACAddress' - force the Tenable.sc MAC address
$DeviceIdentifier   = 'Auto'

# Name for the CSV file that will be created and uploaded to NinjaOne.
$CsvFileName        = "Tenable-Vulnerabilities-$(Get-Date -Format 'yyyy-MM-dd').csv"

# Where to save the CSV on this computer before uploading.
# Defaults to the same folder as this script.
$CsvOutputPath      = Join-Path $PSScriptRoot $CsvFileName

# Where to save the NinjaOne refresh token on this computer.
$TokenFile          = Join-Path $PSScriptRoot 'ninja_refresh_token.txt'

# How many minutes to wait for NinjaOne to finish processing the upload.
$ImportTimeoutMinutes = 10

# ==============================================================================
#  DO NOT EDIT BELOW THIS LINE
# ==============================================================================

$NinjaRedirectUri   = 'https://localhost'
$NinjaAuthEndpoint  = "$NinjaBaseUrl/ws/oauth/authorize"
$NinjaApiBase       = "$NinjaBaseUrl/api/v2"
$NinjaScope         = 'monitoring management control offline_access'

# ------------------------------------------------------------------------------
#  FUNCTION: Get a NinjaOne access token using a saved refresh token
# ------------------------------------------------------------------------------

function Get-NinjaAccessTokenFromRefresh {
    param ([string]$RefreshToken)

    $Body = @{
        grant_type    = 'refresh_token'
        refresh_token = $RefreshToken
        client_id     = $NinjaClientId
        client_secret = $NinjaClientSecret
    }

    try {
        return Invoke-RestMethod -Method Post -Uri $NinjaTokenEndpoint `
            -ContentType 'application/x-www-form-urlencoded' -Body $Body
    } catch {
        return $null
    }
}

# ------------------------------------------------------------------------------
#  FUNCTION: Run the full NinjaOne browser-based authorization code flow
# ------------------------------------------------------------------------------

function Invoke-NinjaAuthCodeFlow {
    Write-Host ''
    Write-Host '---------------------------------------------------------------'
    Write-Host ' NinjaOne Login — Step 1 of 3'
    Write-Host ' Opening your browser to log in to NinjaOne...'
    Write-Host '---------------------------------------------------------------'

    $State     = [System.Guid]::NewGuid().ToString('N')
    $AuthQuery = "response_type=code" +
                 "&client_id=$([Uri]::EscapeDataString($NinjaClientId))" +
                 "&redirect_uri=$([Uri]::EscapeDataString($NinjaRedirectUri))" +
                 "&scope=$([Uri]::EscapeDataString($NinjaScope))" +
                 "&state=$State"

    Start-Process "$NinjaAuthEndpoint`?$AuthQuery"

    Write-Host ''
    Write-Host ' Your browser should open a NinjaOne login page.'
    Write-Host ' Log in and click Authorize. The browser will then show'
    Write-Host ' a blank or "connection refused" page — that is normal.'
    Write-Host ''
    Write-Host '---------------------------------------------------------------'
    Write-Host ' NinjaOne Login — Step 2 of 3'
    Write-Host ' Paste the URL from your browser address bar'
    Write-Host '---------------------------------------------------------------'
    Write-Host ''
    Write-Host ' After authorizing, your browser will show a blank page.'
    Write-Host ' Copy the entire URL from the address bar and paste it here:'
    Write-Host ''

    $PastedUrl = Read-Host ' Paste URL here'
    $Uri       = [System.Uri]$PastedUrl
    $Params    = [System.Web.HttpUtility]::ParseQueryString($Uri.Query)

    if ($Params['state'] -ne $State) {
        Write-Host ''
        Write-Host ' ERROR: The URL does not look right (state mismatch).' -ForegroundColor Red
        Write-Host '        Make sure you copied the full URL from the address bar.' -ForegroundColor Red
        exit 1
    }

    $AuthCode = $Params['code']
    if (-not $AuthCode) {
        Write-Host ''
        Write-Host ' ERROR: No authorization code found in the URL.' -ForegroundColor Red
        Write-Host '        Make sure you copied the full URL from the address bar.' -ForegroundColor Red
        exit 1
    }

    Write-Host ''
    Write-Host ' Authorization code received.'
    Write-Host ''
    Write-Host '---------------------------------------------------------------'
    Write-Host ' NinjaOne Login — Step 3 of 3'
    Write-Host ' Exchanging code for access token...'
    Write-Host '---------------------------------------------------------------'

    $TokenBody = @{
        grant_type    = 'authorization_code'
        code          = $AuthCode
        redirect_uri  = $NinjaRedirectUri
        client_id     = $NinjaClientId
        client_secret = $NinjaClientSecret
    }

    try {
        return Invoke-RestMethod -Method Post -Uri $NinjaTokenEndpoint `
            -ContentType 'application/x-www-form-urlencoded' -Body $TokenBody
    } catch {
        Write-Host ''
        Write-Host " ERROR: Could not get an access token. $_" -ForegroundColor Red
        Write-Host '        Double-check your NinjaClientId, NinjaClientSecret, and NinjaTokenEndpoint.' -ForegroundColor Red
        exit 1
    }
}

# ------------------------------------------------------------------------------
#  FUNCTION: Call the NinjaOne API (GET requests)
# ------------------------------------------------------------------------------

function Invoke-NinjaApiGet {
    param ([string]$Path)

    return Invoke-RestMethod -Method Get -Uri "$NinjaApiBase/$Path" -Headers @{
        Authorization = "Bearer $NinjaAccessToken"
        Accept        = 'application/json'
    }
}

# ------------------------------------------------------------------------------
#  FUNCTION: Work out which Tenable.sc field matches the Scan Group's setting
#  Returns 'Hostname', 'IPAddress', 'MACAddress', or $null if not recognized.
# ------------------------------------------------------------------------------

function Resolve-DeviceIdentifier {
    param ([string]$Mapping)

    $m = ($Mapping -replace '[^a-zA-Z]', '').ToLower()

    if ($m -like '*mac*')                                       { return 'MACAddress' }
    if ($m -like 'ip*')                                         { return 'IPAddress'  }
    if ($m -like '*host*' -or $m -like '*dns*' -or $m -like '*name*') { return 'Hostname' }
    return $null
}

# ------------------------------------------------------------------------------
#  STEP 1 — Authenticate to NinjaOne (refresh token or full login)
# ------------------------------------------------------------------------------

Write-Host ''
Write-Host '============================================================='
Write-Host ' Tenable.sc -> NinjaOne Scan Group Uploader'
Write-Host '============================================================='

$NinjaAccessToken  = $null
$NinjaRefreshToken = $null

if (Test-Path $TokenFile) {
    Write-Host ''
    Write-Host ' [NinjaOne] Found saved refresh token. Renewing access...'
    $Saved    = (Get-Content $TokenFile -Raw).Trim()
    $Response = Get-NinjaAccessTokenFromRefresh -RefreshToken $Saved

    if ($Response -and $Response.access_token) {
        $NinjaAccessToken  = $Response.access_token
        $NinjaRefreshToken = $Response.refresh_token
        Write-Host ' [NinjaOne] Access token renewed successfully.'
    } else {
        Write-Host ' [NinjaOne] Saved token expired. A browser login is required.' -ForegroundColor Yellow
        Remove-Item $TokenFile -Force
    }
}

if (-not $NinjaAccessToken) {
    $Response = Invoke-NinjaAuthCodeFlow
    if (-not $Response.access_token) {
        Write-Host ' ERROR: No access token received after login.' -ForegroundColor Red
        exit 1
    }
    $NinjaAccessToken  = $Response.access_token
    $NinjaRefreshToken = $Response.refresh_token
    Write-Host ' [NinjaOne] Access token obtained successfully.'
}

if ($NinjaRefreshToken) {
    $NinjaRefreshToken | Set-Content $TokenFile -Force
    Write-Host " [NinjaOne] Refresh token saved to: $TokenFile"
} else {
    Write-Host ' [NinjaOne] WARNING: No refresh token returned. Check offline_access is enabled on your API app.' -ForegroundColor Yellow
}

# ------------------------------------------------------------------------------
#  STEP 2 — Find the Scan Group in NinjaOne automatically
# ------------------------------------------------------------------------------

Write-Host ''
Write-Host ' [NinjaOne] Looking up Scan Groups...'

try {
    $ScanGroups = @(Invoke-NinjaApiGet -Path 'vulnerability/scan-groups')
} catch {
    Write-Host ''
    Write-Host " ERROR: Could not retrieve Scan Groups from NinjaOne. $_" -ForegroundColor Red
    Write-Host '        Check your NinjaBaseUrl and that the API app has the monitoring and management scopes.' -ForegroundColor Red
    exit 1
}

$ScanGroup = $null

if ($NinjaScanGroupName) {
    $Matches = @($ScanGroups | Where-Object { $_.groupName -eq $NinjaScanGroupName })
    if ($Matches.Count -eq 1) {
        $ScanGroup = $Matches[0]
    } elseif ($Matches.Count -gt 1) {
        Write-Host ''
        Write-Host " ERROR: More than one Scan Group is named '$NinjaScanGroupName'. Rename one so the name is unique:" -ForegroundColor Red
        $Matches | ForEach-Object { Write-Host "        ID $($_.id) : $($_.groupName)" -ForegroundColor Red }
        exit 1
    }
} elseif ($ScanGroups.Count -eq 1) {
    $ScanGroup = $ScanGroups[0]
} elseif ($ScanGroups.Count -gt 1) {
    Write-Host ''
    Write-Host ' ERROR: NinjaScanGroupName is blank and you have more than one Scan Group. Set the name to one of:' -ForegroundColor Red
    $ScanGroups | ForEach-Object { Write-Host "        ID $($_.id) : $($_.groupName)" -ForegroundColor Red }
    exit 1
}

$DeviceHeader = $null
$CveHeader    = $null

if ($ScanGroup) {
    $ScanGroupId  = $ScanGroup.id
    $DeviceHeader = $ScanGroup.deviceIdHeader
    $CveHeader    = $ScanGroup.cveIdHeader
    Write-Host " [NinjaOne] Found Scan Group '$($ScanGroup.groupName)' (ID $ScanGroupId)."
    Write-Host " [NinjaOne] It expects the columns: Device = '$DeviceHeader', CVE = '$CveHeader'."
    Write-Host " [NinjaOne] Device matching setting in NinjaOne: '$($ScanGroup.deviceInternalMapping)'."
} else {
    Write-Host ''
    Write-Host " [NinjaOne] No Scan Group found$(if ($NinjaScanGroupName) { " named '$NinjaScanGroupName'" })." -ForegroundColor Yellow
    Write-Host ' [NinjaOne] The script will still build a starter CSV for you to create the Scan Group.' -ForegroundColor Yellow
}

# Decide which Tenable.sc field identifies each device
$Identifier = $DeviceIdentifier
if ($Identifier -eq 'Auto') {
    if ($ScanGroup) {
        $Identifier = Resolve-DeviceIdentifier -Mapping $ScanGroup.deviceInternalMapping
        if (-not $Identifier) {
            Write-Host ''
            Write-Host " ERROR: Could not tell which device identifier the Scan Group uses (its setting reads '$($ScanGroup.deviceInternalMapping)')." -ForegroundColor Red
            Write-Host "        Set DeviceIdentifier at the top of this script to 'Hostname', 'IPAddress' or 'MACAddress'." -ForegroundColor Red
            exit 1
        }
    } else {
        $Identifier = 'Hostname'
    }
}
if ($Identifier -notin @('Hostname', 'IPAddress', 'MACAddress')) {
    Write-Host " ERROR: DeviceIdentifier '$DeviceIdentifier' is not valid. Use Auto, Hostname, IPAddress or MACAddress." -ForegroundColor Red
    exit 1
}
Write-Host " [Script] Devices will be identified by: $Identifier"

# Column headings: use the Scan Group's own, or sensible defaults for a starter CSV
if (-not $DeviceHeader) { $DeviceHeader = $Identifier }
if (-not $CveHeader)    { $CveHeader    = 'CVE ID' }

# ------------------------------------------------------------------------------
#  STEP 3 — Pull vulnerability data from Tenable.sc
# ------------------------------------------------------------------------------

Write-Host ''
Write-Host ' [Tenable.sc] Querying vulnerabilities...'

$TenableHeaders = @{
    'x-apikey' = "accesskey=$TenableAccessKey; secretkey=$TenableSecretKey"
    Accept     = 'application/json'
}

$TenableFields = @('pluginID', 'ip', 'dnsName', 'cve')
if ($Identifier -eq 'MACAddress') { $TenableFields += 'macAddress' }

$Vulnerabilities = @()
$PageSize        = 5000
$Offset          = 0

do {
    # Build the analysis query — all severities, from the specified repository
    $AnalysisBody = @{
        type       = 'vuln'
        sourceType = 'cumulative'
        query      = @{
            type        = 'vuln'
            tool        = 'vulndetails'
            startOffset = $Offset
            endOffset   = $Offset + $PageSize
            filters     = @(
                @{
                    id           = 'repositoryIDs'
                    filterName   = 'repositoryIDs'
                    operator     = '='
                    value        = $TenableRepositoryId
                }
            )
            fields = $TenableFields
        }
    } | ConvertTo-Json -Depth 10

    try {
        $TenableResponse = Invoke-RestMethod `
            -Method Post `
            -Uri "$TenableBaseUrl/rest/analysis" `
            -Headers $TenableHeaders `
            -Body $AnalysisBody `
            -ContentType 'application/json' `
            -SkipCertificateCheck   # Remove this line if Tenable.sc has a valid trusted certificate
    } catch {
        Write-Host ''
        Write-Host " ERROR: Could not connect to Tenable.sc. $_" -ForegroundColor Red
        Write-Host '        Check your TenableBaseUrl, TenableAccessKey, and TenableSecretKey.' -ForegroundColor Red
        exit 1
    }

    $Page  = @($TenableResponse.response.results) | Where-Object { $null -ne $_ }
    $Total = 0
    [void][int]::TryParse([string]$TenableResponse.response.totalRecords, [ref]$Total)

    $Vulnerabilities += $Page
    $Offset          += $PageSize

    # Keep going while Tenable.sc says there are more records. If it does not
    # report a total, keep going while pages come back full.
    $MoreRecords = if ($Total -gt 0) { $Vulnerabilities.Count -lt $Total } else { $Page.Count -ge $PageSize }

} while ($Page.Count -gt 0 -and $MoreRecords)

if ($Vulnerabilities.Count -eq 0) {
    Write-Host ' [Tenable.sc] No vulnerabilities returned. Check your Repository ID and that scans have been run.' -ForegroundColor Yellow
    exit 0
}

Write-Host " [Tenable.sc] Retrieved $($Vulnerabilities.Count) vulnerability records."

# ------------------------------------------------------------------------------
#  STEP 4 — Build and save the CSV (one row per device + CVE)
# ------------------------------------------------------------------------------

Write-Host ''
Write-Host ' [CSV] Building CSV file...'

$CsvRows        = @()
$SkippedNoCve   = 0
$SkippedNoDevice = 0

foreach ($Vuln in $Vulnerabilities) {

    $DeviceValue = switch ($Identifier) {
        'Hostname'   { $Vuln.dnsName }
        'IPAddress'  { $Vuln.ip }
        'MACAddress' { $Vuln.macAddress }
    }

    if ([string]::IsNullOrWhiteSpace($DeviceValue)) { $SkippedNoDevice++; continue }

    # Tenable.sc may return CVEs as a list or as one comma-separated string
    $CveIds = @()
    if ($Vuln.cve -is [string]) {
        $CveIds = $Vuln.cve -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    } elseif ($Vuln.cve) {
        $CveIds = @($Vuln.cve) | ForEach-Object { "$_".Trim() } | Where-Object { $_ }
    }

    if ($CveIds.Count -eq 0) { $SkippedNoCve++; continue }

    foreach ($Cve in $CveIds) {
        $Row = [ordered]@{}
        $Row[$DeviceHeader] = $DeviceValue.Trim()
        $Row[$CveHeader]    = $Cve
        $CsvRows += [PSCustomObject]$Row
    }
}

# Remove duplicate device + CVE pairs
$CsvRows = @($CsvRows | Sort-Object -Property $DeviceHeader, $CveHeader -Unique)

Write-Host " [CSV] Skipped $SkippedNoDevice records with no $Identifier and $SkippedNoCve records with no CVE ID."

if ($CsvRows.Count -eq 0) {
    Write-Host ' [CSV] No rows to upload after filtering. Nothing was uploaded.' -ForegroundColor Yellow
    exit 0
}

$CsvRows | Export-Csv -Path $CsvOutputPath -NoTypeInformation -Encoding UTF8

Write-Host " [CSV] Saved to: $CsvOutputPath"
Write-Host " [CSV] Total rows: $($CsvRows.Count)"

# ------------------------------------------------------------------------------
#  If there is no Scan Group yet, stop here and explain how to create one
# ------------------------------------------------------------------------------

if (-not $ScanGroup) {
    Write-Host ''
    Write-Host '=============================================================' -ForegroundColor Yellow
    Write-Host ' ONE-TIME SETUP NEEDED — create the Scan Group in NinjaOne' -ForegroundColor Yellow
    Write-Host '=============================================================' -ForegroundColor Yellow
    Write-Host " 1. In NinjaOne open the vulnerability importer app > Scan groups > Create scan group."
    Write-Host ' 2. Name it, choose the scanning vendor, and click Upload CSV file.'
    Write-Host " 3. Select this file: $CsvOutputPath"
    Write-Host " 4. Map '$DeviceHeader' to the device identifier and '$CveHeader' to the CVE ID, then finish."
    Write-Host " 5. Put the Scan Group's name in NinjaScanGroupName at the top of this script and run it again."
    exit 0
}

# ------------------------------------------------------------------------------
#  STEP 5 — Upload the CSV to the Scan Group
# ------------------------------------------------------------------------------

Write-Host ''
Write-Host " [NinjaOne] Uploading CSV to Scan Group '$($ScanGroup.groupName)' (ID $ScanGroupId)..."

$UploadedBefore = $ScanGroup.lastFileUploadDate

try {
    $UploadResponse = Invoke-RestMethod `
        -Method Post `
        -Uri "$NinjaApiBase/vulnerability/scan-groups/$ScanGroupId/upload" `
        -Headers @{ Authorization = "Bearer $NinjaAccessToken" } `
        -Form @{ csv = Get-Item -Path $CsvOutputPath }
} catch {
    Write-Host ''
    Write-Host " ERROR: Failed to upload CSV to NinjaOne. $_" -ForegroundColor Red
    Write-Host "        Check that Scan Group ID $ScanGroupId exists and that the CSV columns '$DeviceHeader' and '$CveHeader' match the Scan Group." -ForegroundColor Red
    exit 1
}

Write-Host " [NinjaOne] Upload accepted (response: $UploadResponse). Waiting for NinjaOne to process it..."

# ------------------------------------------------------------------------------
#  STEP 6 — Wait for NinjaOne to finish importing and report the result
# ------------------------------------------------------------------------------

$Deadline      = (Get-Date).AddMinutes($ImportTimeoutMinutes)
$SawInProgress = $false
$FinalState    = $null

while ((Get-Date) -lt $Deadline) {
    Start-Sleep -Seconds 10

    try {
        $Current = Invoke-NinjaApiGet -Path "vulnerability/scan-groups/$ScanGroupId"
    } catch {
        continue
    }

    if ($Current.status -eq 'IN_PROGRESS') { $SawInProgress = $true; continue }

    $IsNewImport = $SawInProgress -or ($Current.lastFileUploadDate -ne $UploadedBefore)
    if ($IsNewImport -and $Current.status -in @('COMPLETE', 'FAILED')) {
        $FinalState = $Current
        break
    }
}

Write-Host ''
Write-Host '============================================================='

if (-not $FinalState) {
    Write-Host ' Upload sent, but NinjaOne did not report a result in time.' -ForegroundColor Yellow
    Write-Host "  Check the Scan Group '$($ScanGroup.groupName)' in NinjaOne to confirm the import finished."
} elseif ($FinalState.status -eq 'FAILED') {
    Write-Host ' NinjaOne reported the import FAILED.' -ForegroundColor Red
    Write-Host "  Check that the CSV columns '$DeviceHeader' and '$CveHeader' match the Scan Group's settings."
    Write-Host '============================================================='
    exit 1
} else {
    Write-Host ' Done!'
    Write-Host "  Scan Group             : $($ScanGroup.groupName) (ID $ScanGroupId)"
    Write-Host "  Records processed      : $($FinalState.recordsProcessed)"
}

Write-Host "  Vulnerability rows sent: $($CsvRows.Count)"
Write-Host "  CSV saved to           : $CsvOutputPath"
Write-Host '============================================================='
