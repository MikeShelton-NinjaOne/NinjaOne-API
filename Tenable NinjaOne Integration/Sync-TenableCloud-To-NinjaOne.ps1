# ==============================================================================
#  Tenable Vulnerability Management (Cloud) -> NinjaOne — Scan Group Upload
# ==============================================================================
#
#  FOR: Tenable's CLOUD product (Tenable Vulnerability Management, formerly
#       Tenable.io). If you run Tenable.sc on your own server, use the
#       Tenable.sc version of this script instead.
#
#  WHAT THIS SCRIPT DOES:
#
#  1. Logs in to NinjaOne and looks up your Scan Group by name to get its ID.
#  2. Asks Tenable to export your open vulnerabilities, waits for the export
#     to finish, and downloads the results.
#  3. Builds a CSV file using the exact column headings your Scan Group expects.
#  4. Uploads the CSV to the Scan Group and reports what NinjaOne says back.
#
#  The first time you run this script it will open your browser so you can log
#  in to NinjaOne. After that it saves a refresh token so you do not have to
#  log in again. It renews its own access on every run.
#
#  REQUIREMENTS:
#
#  PowerShell 7 or newer. (Check with:  $PSVersionTable.PSVersion )
#
#  BEFORE YOU RUN THIS SCRIPT THE FIRST TIME:
#
#  1. Log in to NinjaOne and go to:
#        Administration > Apps > API > Client App
#     Create or open an app. The app must be a type that allows the
#     Authorization Code and Refresh Token grant types (a Web or Native app).
#     Set the Redirect URI to exactly:
#        https://localhost
#     Make sure the offline_access scope is allowed on the app.
#
#  2. In Tenable Vulnerability Management, generate an API key (per Tenable's docs):
#        - Log in at cloud.tenable.com.
#        - Click the user profile icon in the upper-right corner (My Account opens).
#        - In the left navigation, click "API Keys".
#        - In the bottom-left corner, click "Generate", then "Continue".
#        - Copy the Access Key and Secret Key. You cannot view them again after
#          you close the My Account page.
#     WARNING: Each Tenable account can only have ONE API key at a time. Generating
#     a new key makes the existing key stop working, which will break anything
#     else that uses it. Consider using a dedicated Tenable user for this script.
#     The user needs at least the Basic [16] role plus the "Can View" permission
#     on the assets to export (Administrator [64] users do not need that extra
#     permission).
#
#  3. In NinjaOne, create the Scan Group ONE TIME by hand. NinjaOne requires the
#     first creation and import of a Scan Group to be done in the console.
#        - Go to Administration > Apps and open the Vulnerability Importer app
#          (or the Tenable app) and enable it.
#        - Open the "Scan groups" tab and click "Create scan group".
#        - Enter a name, choose the scanning vendor, and upload a CSV file.
#        - Map the device column, choose the "Device ID type", and map the CVE column.
#     If you do not have a CSV yet, run this script once. If it cannot find your
#     Scan Group, it still builds a starter CSV and prints the exact steps.
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

# The NAME of the Scan Group in NinjaOne, exactly as shown on the Scan groups tab.
# The script looks up the Scan Group ID for you from this name.
$NinjaScanGroupName = '<Your Scan Group Name>'                   # e.g. Tenable - All Servers

# --- Tenable Vulnerability Management (cloud) credentials ---------------------

# The address of Tenable's cloud API. Use the address of the Tenable cloud you log in to.
$TenableBaseUrl     = 'https://cloud.tenable.com'
$TenableAccessKey   = '<Your Tenable Access Key>'
$TenableSecretKey   = '<Your Tenable Secret Key>'

# --- Scan settings ------------------------------------------------------------

# How devices are identified. This MUST match the "Device ID type" you chose when
# you created the Scan Group in NinjaOne. Type exactly one of these four words:
#   Hostname     - the Tenable hostname
#   FQDN         - the Tenable fully qualified domain name
#   IPAddress    - the Tenable IPv4 address
#   MACAddress   - the Tenable MAC address
$DeviceIdentifier   = '<Hostname, FQDN, IPAddress or MACAddress>'

# How many days back to look for open vulnerabilities.
# Tenable only exports findings it last saw within this window. A device that
# has not been scanned for longer than this many days will NOT be included.
# Tenable's own default window is 30 days.
$LookbackDays       = 30

# How many minutes to wait for Tenable to finish preparing the export.
$TenableExportTimeoutMinutes = 30

# Name for the CSV file that will be created and uploaded to NinjaOne.
$CsvFileName        = "Tenable-Vulnerabilities-$(Get-Date -Format 'yyyy-MM-dd').csv"

# Where to save the CSV on this computer before uploading.
# Defaults to the same folder as this script.
$CsvOutputPath      = Join-Path $PSScriptRoot $CsvFileName

# Where to save the NinjaOne refresh token on this computer.
$TokenFile          = Join-Path $PSScriptRoot 'ninja_refresh_token.txt'

# How many minutes to wait for NinjaOne to report the result of the upload.
$ImportTimeoutMinutes = 10

# ==============================================================================
#  DO NOT EDIT BELOW THIS LINE
# ==============================================================================

# ------------------------------------------------------------------------------
#  CHECKS — make sure the script can run and the settings are filled in
# ------------------------------------------------------------------------------

if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host ''
    Write-Host " ERROR: This script needs PowerShell 7 or newer. You are running $($PSVersionTable.PSVersion)." -ForegroundColor Red
    exit 1
}

$RequiredSettings = [ordered]@{
    NinjaBaseUrl        = $NinjaBaseUrl
    NinjaTokenEndpoint  = $NinjaTokenEndpoint
    NinjaClientId       = $NinjaClientId
    NinjaClientSecret   = $NinjaClientSecret
    NinjaScanGroupName  = $NinjaScanGroupName
    TenableBaseUrl      = $TenableBaseUrl
    TenableAccessKey    = $TenableAccessKey
    TenableSecretKey    = $TenableSecretKey
    DeviceIdentifier    = $DeviceIdentifier
}
$NotFilledIn = @($RequiredSettings.Keys | Where-Object { [string]::IsNullOrWhiteSpace($RequiredSettings[$_]) -or $RequiredSettings[$_] -match '<.+>' })
if ($NotFilledIn.Count -gt 0) {
    Write-Host ''
    Write-Host ' ERROR: These settings in the CONFIGURATION section still need a value:' -ForegroundColor Red
    $NotFilledIn | ForEach-Object { Write-Host "        $_" -ForegroundColor Red }
    exit 1
}

if ($DeviceIdentifier -notin @('Hostname', 'FQDN', 'IPAddress', 'MACAddress')) {
    Write-Host ''
    Write-Host " ERROR: DeviceIdentifier '$DeviceIdentifier' is not valid. Use exactly: Hostname, FQDN, IPAddress or MACAddress." -ForegroundColor Red
    exit 1
}

if (-not ($LookbackDays -is [int]) -or $LookbackDays -lt 1) {
    Write-Host ''
    Write-Host ' ERROR: LookbackDays must be a whole number of 1 or more.' -ForegroundColor Red
    exit 1
}

$TenableBaseUrl     = $TenableBaseUrl.TrimEnd('/')
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
        Write-Host " [NinjaOne] Token renewal failed: $($_.Exception.Message)" -ForegroundColor Yellow
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
#  FUNCTION: Call the Tenable API. Waits and retries when Tenable says
#  "too many requests" (429) or reports an internal error (500 and above).
# ------------------------------------------------------------------------------

$TenableHeaders = @{
    'X-ApiKeys' = "accessKey=$TenableAccessKey; secretKey=$TenableSecretKey;"
    Accept      = 'application/json'
}

function Invoke-TenableApi {
    param (
        [string]$Method,
        [string]$Path,
        [string]$Body = $null,
        [int]$MaxAttempts = 5
    )

    for ($Attempt = 1; $Attempt -le $MaxAttempts; $Attempt++) {
        try {
            $Request = @{
                Method      = $Method
                Uri         = "$TenableBaseUrl$Path"
                Headers     = $TenableHeaders
                ContentType = 'application/json'
            }
            if ($Body) { $Request['Body'] = $Body }
            return Invoke-RestMethod @Request
        } catch {
            $Code = $null
            if ($_.Exception.Response) { $Code = [int]$_.Exception.Response.StatusCode }

            $Retryable = ($Code -eq 429) -or ($Code -ge 500)
            if ($Retryable -and $Attempt -lt $MaxAttempts) {
                $Wait = [math]::Pow(2, $Attempt) * 5
                try {
                    $RetryAfter = $_.Exception.Response.Headers.RetryAfter.Delta.TotalSeconds
                    if ($RetryAfter -gt 0) { $Wait = [math]::Ceiling($RetryAfter) }
                } catch { }
                Write-Host " [Tenable] HTTP $Code. Waiting $Wait seconds and trying again (attempt $Attempt of $MaxAttempts)..." -ForegroundColor Yellow
                Start-Sleep -Seconds $Wait
                continue
            }
            throw "HTTP $Code : $($_.Exception.Message)"
        }
    }
}

# ------------------------------------------------------------------------------
#  FUNCTIONS: Read values out of Tenable records (which are nested, e.g. asset.fqdn)
# ------------------------------------------------------------------------------

function Get-NestedValue {
    param ($Object, [string]$Path)

    $Current = $Object
    foreach ($Part in $Path.Split('.')) {
        if ($null -eq $Current) { return $null }
        $Prop = $Current.PSObject.Properties[$Part]
        if (-not $Prop) { return $null }
        $Current = $Prop.Value
    }
    return $Current
}

# Returns the values from the first path that has any, as a list of trimmed text.
function Get-FirstValues {
    param ($Object, [string[]]$Paths)

    foreach ($Path in $Paths) {
        $Value = Get-NestedValue -Object $Object -Path $Path
        $List  = @(@($Value) | ForEach-Object { "$_".Trim() } | Where-Object { $_ })
        if ($List.Count -gt 0) { return $List }
    }
    return @()
}

$IdentifierPaths = @{
    Hostname   = @('asset.hostname')
    FQDN       = @('asset.fqdn')
    IPAddress  = @('asset.ipv4', 'asset.ipv4_addresses')
    MACAddress = @('asset.mac_address', 'asset.mac_addresses')
}

# ------------------------------------------------------------------------------
#  STEP 1 — Authenticate to NinjaOne (refresh token or full login)
# ------------------------------------------------------------------------------

Write-Host ''
Write-Host '============================================================='
Write-Host ' Tenable Cloud -> NinjaOne Scan Group Uploader'
Write-Host '============================================================='

$NinjaAccessToken  = $null
$NinjaRefreshToken = $null
$Renewed           = $false

if (Test-Path $TokenFile) {
    Write-Host ''
    Write-Host ' [NinjaOne] Found saved refresh token. Renewing access...'
    $Saved    = (Get-Content $TokenFile -Raw).Trim()
    $Response = Get-NinjaAccessTokenFromRefresh -RefreshToken $Saved

    if ($Response -and $Response.access_token) {
        $NinjaAccessToken  = $Response.access_token
        $NinjaRefreshToken = $Response.refresh_token
        $Renewed           = $true
        Write-Host ' [NinjaOne] Access token renewed successfully.'
    } else {
        Write-Host ' [NinjaOne] Could not renew with the saved token. A browser login is required.' -ForegroundColor Yellow
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
} elseif ($Renewed) {
    Write-Host ' [NinjaOne] NinjaOne did not return a new refresh token. Keeping the saved one.'
} else {
    Write-Host ' [NinjaOne] WARNING: No refresh token returned. Check that offline_access is allowed and Refresh Token is enabled on your API app. You will need to log in through the browser on every run.' -ForegroundColor Yellow
}

# ------------------------------------------------------------------------------
#  STEP 2 — Look up the Scan Group in NinjaOne by name
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

$NameMatches = @($ScanGroups | Where-Object { $_.groupName -eq $NinjaScanGroupName })

if ($NameMatches.Count -gt 1) {
    Write-Host ''
    Write-Host " ERROR: More than one Scan Group is named '$NinjaScanGroupName'. Rename one so the name is unique:" -ForegroundColor Red
    $NameMatches | ForEach-Object { Write-Host "        ID $($_.id) : $($_.groupName)" -ForegroundColor Red }
    exit 1
}

$ScanGroup    = $null
$DeviceHeader = $null
$CveHeader    = $null

if ($NameMatches.Count -eq 1) {
    $ScanGroup    = $NameMatches[0]
    $ScanGroupId  = $ScanGroup.id
    $DeviceHeader = $ScanGroup.deviceIdHeader
    $CveHeader    = $ScanGroup.cveIdHeader

    if ([string]::IsNullOrWhiteSpace($DeviceHeader) -or [string]::IsNullOrWhiteSpace($CveHeader)) {
        Write-Host ''
        Write-Host " ERROR: Scan Group '$($ScanGroup.groupName)' (ID $ScanGroupId) did not report its device and CVE column names, so the CSV cannot be matched to it." -ForegroundColor Red
        Write-Host "        NinjaOne returned: deviceIdHeader='$DeviceHeader', cveIdHeader='$CveHeader'." -ForegroundColor Red
        exit 1
    }

    Write-Host " [NinjaOne] Found Scan Group '$($ScanGroup.groupName)' (ID $ScanGroupId)."
    Write-Host " [NinjaOne] It expects the columns: device = '$DeviceHeader', CVE = '$CveHeader'."
    Write-Host " [NinjaOne] Its device matching setting reads: '$($ScanGroup.deviceInternalMapping)'."
    Write-Host " [Script]   Your DeviceIdentifier setting is: $DeviceIdentifier. These two should refer to the same thing."
} else {
    Write-Host ''
    Write-Host " [NinjaOne] No Scan Group named '$NinjaScanGroupName' was found." -ForegroundColor Yellow
    if ($ScanGroups.Count -gt 0) {
        Write-Host ' [NinjaOne] Scan Groups that do exist:' -ForegroundColor Yellow
        $ScanGroups | ForEach-Object { Write-Host "            ID $($_.id) : $($_.groupName)" -ForegroundColor Yellow }
    }
    Write-Host ' [NinjaOne] The script will still build a starter CSV so you can create the Scan Group.' -ForegroundColor Yellow

    $DeviceHeader = 'Device ID'
    $CveHeader    = 'CVE ID'
}

# ------------------------------------------------------------------------------
#  STEP 3 — Ask Tenable to export vulnerabilities
# ------------------------------------------------------------------------------

Write-Host ''
Write-Host " [Tenable] Requesting an export of open vulnerabilities seen in the last $LookbackDays days..."

$SinceEpoch = [DateTimeOffset]::UtcNow.AddDays(-$LookbackDays).ToUnixTimeSeconds()

# state: only vulnerabilities that are currently open (OPEN) or came back (REOPENED).
# since: only findings last found on or after this date (Unix time).
# include_plugin_output = false: leaves out bulky scan output we do not need.
$ExportRequest = @{
    num_assets            = 500
    include_plugin_output = $false
    filters               = @{
        state = @('OPEN', 'REOPENED')
        since = $SinceEpoch
    }
} | ConvertTo-Json -Depth 5

try {
    $ExportJob = Invoke-TenableApi -Method Post -Path '/vulns/export' -Body $ExportRequest
} catch {
    Write-Host ''
    if ("$_" -like 'HTTP 409*') {
        Write-Host ' ERROR: Tenable says an identical export is already running. Wait for it to finish and run this script again.' -ForegroundColor Red
    } else {
        Write-Host " ERROR: Could not start the Tenable export. $_" -ForegroundColor Red
        Write-Host '        Check your TenableBaseUrl, TenableAccessKey and TenableSecretKey, and that the key''s user is allowed to export vulnerabilities.' -ForegroundColor Red
    }
    exit 1
}

$ExportUuid = $ExportJob.export_uuid
if (-not $ExportUuid) {
    Write-Host ' ERROR: Tenable did not return an export ID.' -ForegroundColor Red
    exit 1
}
Write-Host " [Tenable] Export started (ID $ExportUuid). Waiting for it to finish..."

# ------------------------------------------------------------------------------
#  STEP 4 — Wait for the export to finish
# ------------------------------------------------------------------------------

$ExportDeadline = (Get-Date).AddMinutes($TenableExportTimeoutMinutes)
$ExportStatus   = $null

while ($true) {
    if ((Get-Date) -gt $ExportDeadline) {
        Write-Host ''
        Write-Host " ERROR: Tenable did not finish the export within $TenableExportTimeoutMinutes minutes. Last status: $($ExportStatus.status)." -ForegroundColor Red
        Write-Host '        Increase TenableExportTimeoutMinutes or try again later.' -ForegroundColor Red
        exit 1
    }

    try {
        $ExportStatus = Invoke-TenableApi -Method Get -Path "/vulns/export/$ExportUuid/status"
    } catch {
        Write-Host ''
        Write-Host " ERROR: Could not check the Tenable export status. $_" -ForegroundColor Red
        exit 1
    }

    if ($ExportStatus.status -eq 'FINISHED') { break }

    if ($ExportStatus.status -in @('ERROR', 'CANCELLED')) {
        Write-Host ''
        Write-Host " ERROR: The Tenable export ended with status $($ExportStatus.status)." -ForegroundColor Red
        exit 1
    }

    Write-Host " [Tenable] Status: $($ExportStatus.status). Chunks ready: $(@($ExportStatus.chunks_available).Count)..."
    Start-Sleep -Seconds 10
}

$ChunkIds = @($ExportStatus.chunks_available)

if (@($ExportStatus.chunks_failed).Count -gt 0 -or @($ExportStatus.chunks_cancelled).Count -gt 0) {
    Write-Host ''
    Write-Host " ERROR: Tenable reported failed or cancelled chunks (failed: $(@($ExportStatus.chunks_failed) -join ', '); cancelled: $(@($ExportStatus.chunks_cancelled) -join ', ')). The export is incomplete, so nothing was uploaded." -ForegroundColor Red
    exit 1
}

if ($ChunkIds.Count -eq 0) {
    Write-Host ' [Tenable] The export finished but contains no vulnerabilities. Check LookbackDays and that scans have run.' -ForegroundColor Yellow
    exit 0
}

Write-Host " [Tenable] Export finished. Downloading $($ChunkIds.Count) chunk(s)..."

# ------------------------------------------------------------------------------
#  STEP 5 — Download each chunk and build the rows (one per device + CVE)
# ------------------------------------------------------------------------------

$CsvRows         = [System.Collections.Generic.List[object]]::new()
$TotalFindings   = 0
$SkippedNoDevice = 0
$SkippedNoCve    = 0
$FirstRecord     = $null

foreach ($ChunkId in $ChunkIds) {
    try {
        $Records = @(Invoke-TenableApi -Method Get -Path "/vulns/export/$ExportUuid/chunks/$ChunkId")
    } catch {
        Write-Host ''
        Write-Host " ERROR: Could not download chunk $ChunkId from Tenable. $_" -ForegroundColor Red
        exit 1
    }

    foreach ($Record in $Records) {
        if ($null -eq $Record) { continue }
        $TotalFindings++
        if ($null -eq $FirstRecord) { $FirstRecord = $Record }

        $DeviceValues = @(Get-FirstValues -Object $Record -Paths $IdentifierPaths[$DeviceIdentifier])
        if ($DeviceValues.Count -eq 0) { $SkippedNoDevice++; continue }

        # The CVE field is a list; if it ever arrives as text, split it on commas/semicolons.
        $CveIds = @(Get-FirstValues -Object $Record -Paths @('plugin.cve') | ForEach-Object { $_ -split '[,;]' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        if ($CveIds.Count -eq 0) { $SkippedNoCve++; continue }

        foreach ($Device in $DeviceValues) {
            foreach ($Cve in $CveIds) {
                $Row = [ordered]@{}
                $Row[$DeviceHeader] = $Device
                $Row[$CveHeader]    = $Cve
                $CsvRows.Add([PSCustomObject]$Row)
            }
        }
    }

    Write-Host " [Tenable] Downloaded chunk $ChunkId ($TotalFindings findings so far)."
}

Write-Host " [CSV] Skipped $SkippedNoDevice findings with no $DeviceIdentifier value and $SkippedNoCve findings with no CVE ID."

$UniqueRows = @($CsvRows | Sort-Object -Property $DeviceHeader, $CveHeader -Unique)

if ($UniqueRows.Count -eq 0) {
    Write-Host ''
    Write-Host ' [CSV] No rows to upload. Nothing was uploaded.' -ForegroundColor Yellow
    if ($FirstRecord) {
        $TopLevel = ($FirstRecord.PSObject.Properties.Name) -join ', '
        $AssetProps = if ($FirstRecord.asset) { ($FirstRecord.asset.PSObject.Properties.Name) -join ', ' } else { '(none)' }
        $PluginProps = if ($FirstRecord.plugin) { ($FirstRecord.plugin.PSObject.Properties.Name) -join ', ' } else { '(none)' }
        Write-Host ' [CSV] To help find out why, here is what Tenable returned for the first finding:' -ForegroundColor Yellow
        Write-Host "       Top-level fields : $TopLevel"
        Write-Host "       asset fields     : $AssetProps"
        Write-Host "       plugin fields    : $PluginProps"
    }
    exit 0
}

# ------------------------------------------------------------------------------
#  STEP 6 — Save the CSV
# ------------------------------------------------------------------------------

Write-Host ''
Write-Host ' [CSV] Saving CSV file...'
$UniqueRows | Export-Csv -Path $CsvOutputPath -NoTypeInformation -Encoding UTF8

Write-Host " [CSV] Saved to: $CsvOutputPath"
Write-Host " [CSV] Total rows (duplicates removed): $($UniqueRows.Count)"

# ------------------------------------------------------------------------------
#  If there is no Scan Group yet, stop here and explain how to create one
# ------------------------------------------------------------------------------

if (-not $ScanGroup) {
    Write-Host ''
    Write-Host '=============================================================' -ForegroundColor Yellow
    Write-Host ' ONE-TIME SETUP NEEDED — create the Scan Group in NinjaOne' -ForegroundColor Yellow
    Write-Host '=============================================================' -ForegroundColor Yellow
    Write-Host ' 1. In NinjaOne go to Administration > Apps, open the Vulnerability Importer app and enable it.'
    Write-Host ' 2. Open the Scan groups tab and click Create scan group.'
    Write-Host " 3. Enter the name '$NinjaScanGroupName' and choose the scanning vendor."
    Write-Host " 4. Click Upload CSV file and select: $CsvOutputPath"
    Write-Host " 5. Map the device column ('$DeviceHeader') and set 'Device ID type' to match '$DeviceIdentifier'."
    Write-Host "    Map the CVE column ('$CveHeader') to the CVE ID, then click Finish."
    Write-Host ' 6. Run this script again.'
    exit 0
}

# ------------------------------------------------------------------------------
#  STEP 7 — Upload the CSV to the Scan Group
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

Write-Host " [NinjaOne] Upload accepted. NinjaOne returned: $UploadResponse"
Write-Host ' [NinjaOne] Checking the Scan Group status...'

# ------------------------------------------------------------------------------
#  STEP 8 — Report what NinjaOne says about the import
# ------------------------------------------------------------------------------

$Deadline = (Get-Date).AddMinutes($ImportTimeoutMinutes)
$Latest   = $null
$Finished = $false

while ((Get-Date) -lt $Deadline) {
    Start-Sleep -Seconds 10

    try {
        $Latest = Invoke-NinjaApiGet -Path "vulnerability/scan-groups/$ScanGroupId"
    } catch {
        continue
    }

    # Only trust a finished status once NinjaOne shows a new upload date.
    $NewUploadDate = $Latest.lastFileUploadDate -ne $UploadedBefore
    if ($NewUploadDate -and $Latest.status -in @('COMPLETE', 'FAILED')) {
        $Finished = $true
        break
    }
}

Write-Host ''
Write-Host '============================================================='

if (-not $Finished) {
    Write-Host " NinjaOne did not report a finished import within $ImportTimeoutMinutes minutes." -ForegroundColor Yellow
    if ($Latest) {
        Write-Host "  Last reported status   : $($Latest.status)"
        Write-Host "  Records processed      : $($Latest.recordsProcessed)"
        Write-Host "  Last file upload date  : $($Latest.lastFileUploadDate)"
    }
    Write-Host "  Check Scan Group '$($ScanGroup.groupName)' in NinjaOne to confirm the result."
} elseif ($Latest.status -eq 'FAILED') {
    Write-Host ' NinjaOne reported the import FAILED.' -ForegroundColor Red
    Write-Host "  Scan Group             : $($ScanGroup.groupName) (ID $ScanGroupId)"
    Write-Host "  Check that the CSV columns '$DeviceHeader' and '$CveHeader' match the Scan Group's settings."
    Write-Host '============================================================='
    exit 1
} else {
    Write-Host ' NinjaOne reported the import COMPLETE.'
    Write-Host "  Scan Group             : $($ScanGroup.groupName) (ID $ScanGroupId)"
    Write-Host "  Records processed      : $($Latest.recordsProcessed)"
    Write-Host "  Last file upload date  : $($Latest.lastFileUploadDate)"
}

Write-Host "  Rows sent              : $($UniqueRows.Count)"
Write-Host "  CSV saved to           : $CsvOutputPath"
Write-Host '============================================================='
