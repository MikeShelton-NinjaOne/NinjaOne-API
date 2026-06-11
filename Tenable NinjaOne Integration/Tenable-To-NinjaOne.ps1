# ==============================================================================
#  Tenable.sc -> NinjaOne — Vulnerability CSV Upload
# ==============================================================================
#
#  WHAT THIS SCRIPT DOES:
#
#  1. Connects to your Tenable.sc instance and pulls vulnerability scan data.
#  2. Formats the results as a CSV file.
#  3. Uploads the CSV to a device group in NinjaOne as a document.
#
#  The first time you run this script it will open your browser so you can log
#  in to NinjaOne. After that it saves a refresh token so you never have to
#  log in again — it renews its own access automatically on every run.
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
#  3. Fill in all values in the CONFIGURATION section below.
#
# ==============================================================================
#  CONFIGURATION — only edit values in this section
# ==============================================================================

# --- NinjaOne credentials -----------------------------------------------------

$NinjaBaseUrl       = 'https://<your NinjaOne login URL>'        # e.g. https://app.ninjarmm.com
$NinjaTokenEndpoint = 'https://<your NinjaOne login URL>/ws/oauth/token'
$NinjaClientId      = '<Your NinjaOne Client ID>'
$NinjaClientSecret  = '<Your NinjaOne Client Secret>'

# The NinjaOne Organization ID that contains the device group to upload to.
# Find it in NinjaOne under Organizations — the number in the URL.
$NinjaOrganizationId = '<Your Organization ID>'                  # e.g. 42

# --- Tenable.sc credentials ---------------------------------------------------

$TenableBaseUrl     = 'https://<your Tenable.sc hostname or IP>' # e.g. https://tenable.company.com
$TenableAccessKey   = '<Your Tenable.sc Access Key>'
$TenableSecretKey   = '<Your Tenable.sc Secret Key>'

# --- Scan settings ------------------------------------------------------------

# The Tenable.sc Repository ID to pull vulnerabilities from.
# In Tenable.sc go to: Repositories — the ID is shown in the repository list.
$TenableRepositoryId = '<Your Repository ID>'                    # e.g. 1

# Name for the CSV file that will be created and uploaded to NinjaOne.
$CsvFileName        = "Tenable-Vulnerabilities-$(Get-Date -Format 'yyyy-MM-dd').csv"

# Where to save the CSV on this computer before uploading.
# Defaults to the same folder as this script.
$CsvOutputPath      = Join-Path $PSScriptRoot $CsvFileName

# Where to save the NinjaOne refresh token on this computer.
$TokenFile          = Join-Path $PSScriptRoot 'ninja_refresh_token.txt'

# ==============================================================================
#  DO NOT EDIT BELOW THIS LINE
# ==============================================================================

$NinjaRedirectUri   = 'https://localhost'
$NinjaAuthEndpoint  = "$NinjaBaseUrl/ws/oauth/authorize"
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
#  STEP 1 — Authenticate to NinjaOne (refresh token or full login)
# ------------------------------------------------------------------------------

Write-Host ''
Write-Host '============================================================='
Write-Host ' Tenable.sc -> NinjaOne Vulnerability CSV Uploader'
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

$NinjaHeaders = @{
    Authorization  = "Bearer $NinjaAccessToken"
    'Content-Type' = 'application/json'
}

# ------------------------------------------------------------------------------
#  STEP 2 — Pull vulnerability data from Tenable.sc
# ------------------------------------------------------------------------------

Write-Host ''
Write-Host ' [Tenable.sc] Querying vulnerabilities...'

$TenableHeaders = @{
    'x-apikey' = "accesskey=$TenableAccessKey; secretkey=$TenableSecretKey"
    Accept      = 'application/json'
}

# Build the analysis query — all severities, from the specified repository
$AnalysisBody = @{
    type       = 'vuln'
    sourceType = 'cumulative'
    query      = @{
        type        = 'vuln'
        tool        = 'vulndetails'
        startOffset = 0
        endOffset   = 5000
        filters     = @(
            @{
                id           = 'repositoryIDs'
                filterName   = 'repositoryIDs'
                operator     = '='
                value        = $TenableRepositoryId
            }
        )
        fields = @(
            'pluginID',
            'pluginName',
            'severity',
            'ip',
            'dnsName',
            'cve',
            'firstSeen',
            'lastSeen'
        )
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

$Vulnerabilities = $TenableResponse.response.results

if (-not $Vulnerabilities -or $Vulnerabilities.Count -eq 0) {
    Write-Host ' [Tenable.sc] No vulnerabilities returned. Check your Repository ID and that scans have been run.' -ForegroundColor Yellow
    exit 0
}

Write-Host " [Tenable.sc] Retrieved $($Vulnerabilities.Count) vulnerability records."

# ------------------------------------------------------------------------------
#  STEP 3 — Build and save the CSV
# ------------------------------------------------------------------------------

Write-Host ''
Write-Host " [CSV] Building CSV file..."

$SeverityMap = @{
    '0' = 'Informational'
    '1' = 'Low'
    '2' = 'Medium'
    '3' = 'High'
    '4' = 'Critical'
}

$CsvRows = $Vulnerabilities | ForEach-Object {
    $SeverityLabel = if ($SeverityMap.ContainsKey($_.severity.id)) {
        $SeverityMap[$_.severity.id]
    } else {
        $_.severity.name
    }

    $CveList = if ($_.cve) { ($_.cve -join '; ') } else { '' }

    [PSCustomObject]@{
        'Plugin ID'   = $_.pluginID
        'Name'        = $_.pluginName
        'Severity'    = $SeverityLabel
        'IP Address'  = $_.ip
        'Hostname'    = $_.dnsName
        'CVE'         = $CveList
        'First Seen'  = if ($_.firstSeen) { (Get-Date '1970-01-01').AddSeconds($_.firstSeen).ToString('yyyy-MM-dd') } else { '' }
        'Last Seen'   = if ($_.lastSeen)  { (Get-Date '1970-01-01').AddSeconds($_.lastSeen).ToString('yyyy-MM-dd')  } else { '' }
    }
}

$CsvRows | Export-Csv -Path $CsvOutputPath -NoTypeInformation -Encoding UTF8

Write-Host " [CSV] Saved to: $CsvOutputPath"
Write-Host " [CSV] Total rows: $($CsvRows.Count)"

# ------------------------------------------------------------------------------
#  STEP 4 — Upload the CSV to NinjaOne as a document on the Organization
# ------------------------------------------------------------------------------

Write-Host ''
Write-Host " [NinjaOne] Uploading CSV to Organization ID: $NinjaOrganizationId..."

# Upload the file and get back a file token
$FileBytes    = [System.IO.File]::ReadAllBytes($CsvOutputPath)
$Boundary     = [System.Guid]::NewGuid().ToString('N')
$FileName     = [System.IO.Path]::GetFileName($CsvOutputPath)

$BodyLines = (
    "--$Boundary",
    "Content-Disposition: form-data; name=`"file`"; filename=`"$FileName`"",
    "Content-Type: text/csv",
    "",
    [System.Text.Encoding]::UTF8.GetString($FileBytes),
    "--$Boundary--"
) -join "`r`n"

$UploadHeaders = @{
    Authorization  = "Bearer $NinjaAccessToken"
    'Content-Type' = "multipart/form-data; boundary=$Boundary"
}

try {
    $UploadResponse = Invoke-RestMethod `
        -Method Post `
        -Uri "$NinjaBaseUrl/v2/organization/$NinjaOrganizationId/document" `
        -Headers $UploadHeaders `
        -Body ([System.Text.Encoding]::UTF8.GetBytes($BodyLines))
} catch {
    Write-Host ''
    Write-Host " ERROR: Failed to upload CSV to NinjaOne. $_" -ForegroundColor Red
    Write-Host '        Check your NinjaOrganizationId and that your account has permission to upload documents.' -ForegroundColor Red
    exit 1
}

Write-Host ''
Write-Host '============================================================='
Write-Host ' Done!'
Write-Host "  Vulnerabilities pulled : $($CsvRows.Count)"
Write-Host "  CSV saved to           : $CsvOutputPath"
Write-Host "  Uploaded to NinjaOne   : Organization $NinjaOrganizationId"
Write-Host '============================================================='
