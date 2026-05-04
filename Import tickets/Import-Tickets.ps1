<#
.SYNOPSIS
    Import-Tickets.ps1 — Bulk-import tickets from an Excel spreadsheet into NinjaOne via REST API.

.DESCRIPTION
    Designed for new NinjaOne customers migrating ticket data from another ticketing system.
    Reads tickets from an .xlsx file and creates them via the NinjaOne API one at a time, with
    validation, retry logic, throttling, and a full run log.

    AUTHENTICATION:
    NinjaOne requires the Authorization Code OAuth flow for write operations like creating
    tickets. On your FIRST run, this script will open a browser window so you can log in to
    NinjaOne and approve the application. After that, your login is saved (encrypted) so you
    won't be prompted again until it expires.

.HOW TO USE THIS SCRIPT (for first-time PowerShell users)
    Step 1.  Install the two prerequisite modules ONE TIME by opening PowerShell and running:
                Install-Module ImportExcel    -Scope CurrentUser
                Install-Module PSAuthClient   -Scope CurrentUser

    Step 2.  Create an API Client App in NinjaOne under Administration > Apps > API >
             Client App IDs. Use these settings:
                - Application Platform : API Services (machine-to-machine)
                - Redirect URI         : https://localhost   (no port, no slash, no path)
                - Scopes               : Monitoring + Management
                - Allowed Grant Types  : Authorization Code + Refresh Token
             Save the Client ID and Client Secret somewhere safe — the secret is only
             shown once.

    Step 3.  Open this .ps1 file in Notepad or VS Code, scroll to the
             "# === CONFIGURATION ===" block, and replace every value wrapped in angle
             brackets (for example <Your Client ID>) with your real information.
             Do NOT remove the single quotes around the values.

    Step 4.  Save the file. Make sure DryRun is set to $true for your first run — this
             validates your spreadsheet WITHOUT sending any data to NinjaOne.

    Step 5.  Run the script from a PowerShell window:
                .\Import-Tickets.ps1
             The first time you run it, your browser will open. Log in to NinjaOne and
             approve the application. The browser will redirect to a page that probably
             shows "site can't be reached" — that is expected. The script captures what
             it needs and continues automatically.

    After a successful dry run, change DryRun to $false and run the script again to
    perform the real import.

    YOU SHOULD ONLY EVER NEED TO EDIT THE # === CONFIGURATION === BLOCK BELOW.
#>

# === CONFIGURATION ===
# Edit the values below before running the script.
# Replace everything inside the angle brackets < > with your own information.
# Do NOT remove the quotation marks.

$Config = @{
    # Your NinjaOne portal URL (the address you use to log in).
    # Examples: https://app.ninjarmm.com  (US primary)
    #           https://us2.ninjarmm.com  (US secondary)
    #           https://ca.ninjarmm.com   (Canada)
    #           https://eu.ninjarmm.com   (Europe)
    #           https://oc.ninjarmm.com   (Oceania / Australia)
    BaseUrl       = 'https://<your NinjaOne URL>'

    # Your API Client ID — from NinjaOne under Administration > Apps > API > Client App IDs
    ClientId      = '<Your Client ID>'

    # Your API Client Secret — shown ONCE when you saved the API client app. Keep it private.
    ClientSecret  = '<Your Client Secret>'

    # Redirect URI — leave this exactly as shown.
    # IMPORTANT: this MUST match what you set in NinjaOne EXACTLY.
    # No port number, no path, no trailing slash. Just: https://localhost
    RedirectUri   = 'https://localhost'

    # Full path to the spreadsheet you want to import (example: C:\Migration\Tickets.xlsx)
    SpreadsheetPath = '<Full path to your .xlsx file>'

    # The name of the worksheet/tab inside the spreadsheet that contains the tickets
    WorksheetName = 'Tickets'

    # Where to save the run log (the script will create this file)
    LogFilePath   = '.\ticket-import-log.txt'

    # Where to save the encrypted refresh token cache.
    # This file is encrypted with Windows DPAPI — only THIS Windows user account on THIS
    # computer can decrypt it. Safe to leave at the default. Delete this file to force a
    # fresh browser login on the next run.
    TokenCachePath = '.\ninja-token-cache.xml'

    # Set to $true for a test run that validates the file but does NOT send any data to the API.
    # Always run with $true the FIRST time to make sure your data looks right.
    DryRun        = $true
}

# ============================================================================
# DO NOT EDIT BELOW THIS LINE
# ============================================================================

# ---- Script-wide state ----
$Script:AccessToken    = $null     # In-memory cached access token for this run
$Script:Stats = @{
    TotalRows  = 0
    Imported   = 0
    Skipped    = [System.Collections.Generic.List[object]]::new()
    Failed     = [System.Collections.Generic.List[object]]::new()
}

# OAuth scope: 'monitoring' to read, 'management' to write tickets, 'offline_access' to get a refresh token
$Script:Scope = 'monitoring management offline_access'

# Throttle: max 5 requests per second  =>  minimum 200 ms between requests
$Script:MinIntervalMs  = 200
$Script:LastRequestAt  = [DateTime]::MinValue


# ----------------------------------------------------------------------------
# Custom exception used to signal "refresh token didn't work, fall back to
# interactive browser login". This lets Get-AuthToken cleanly distinguish a
# recoverable refresh failure from a hard error.
# ----------------------------------------------------------------------------
class RefreshTokenInvalidException : System.Exception {
    RefreshTokenInvalidException([string]$msg) : base($msg) {}
}


# ----------------------------------------------------------------------------
# Write-Log
#   Logs a message to both the console (with color) and the log file (plain text).
#   Levels: INFO (cyan), SUCCESS (green), WARN (yellow), ERROR (red)
# ----------------------------------------------------------------------------
function Write-Log {
    param(
        [Parameter(Mandatory)] [string] $Message,
        [ValidateSet('INFO','SUCCESS','WARN','ERROR')] [string] $Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line      = "[$timestamp] [$Level] $Message"

    # Map level to console color
    switch ($Level) {
        'SUCCESS' { Write-Host $line -ForegroundColor Green }
        'WARN'    { Write-Host $line -ForegroundColor Yellow }
        'ERROR'   { Write-Host $line -ForegroundColor Red }
        default   { Write-Host $line -ForegroundColor Cyan }
    }

    # Always append to log file (plain text, no color codes)
    try {
        Add-Content -Path $Config.LogFilePath -Value $line -ErrorAction Stop
    } catch {
        # Don't crash the whole script if logging fails — just warn once on the console
        Write-Host "[$timestamp] [WARN] Could not write to log file '$($Config.LogFilePath)': $($_.Exception.Message)" -ForegroundColor Yellow
    }
}


# ----------------------------------------------------------------------------
# Test-Configuration
#   Pre-flight checks. Verifies that the user has filled in all placeholders,
#   the spreadsheet exists, the required modules are installed, and the
#   worksheet name is correct. Exits the script if anything is wrong.
# ----------------------------------------------------------------------------
function Test-Configuration {
    Write-Log 'Running pre-flight configuration checks...' 'INFO'

    # 1. Check for unreplaced <placeholder> values in any string config field
    $placeholdersFound = @()
    foreach ($key in $Config.Keys) {
        $value = $Config[$key]
        if ($value -is [string] -and $value -match '[<>]') {
            $placeholdersFound += $key
        }
    }
    if ($placeholdersFound.Count -gt 0) {
        Write-Log 'The following configuration field(s) still contain placeholder values that you need to replace:' 'ERROR'
        foreach ($f in $placeholdersFound) {
            Write-Log "    - $f  (current value: $($Config[$f]))" 'ERROR'
        }
        Write-Log 'Open this script in Notepad or VS Code and edit the # === CONFIGURATION === block at the top.' 'ERROR'
        exit 1
    }

    # 2. Confirm the spreadsheet file exists
    if (-not (Test-Path -LiteralPath $Config.SpreadsheetPath)) {
        Write-Log "Cannot find the spreadsheet file at: $($Config.SpreadsheetPath)" 'ERROR'
        Write-Log 'Double-check the SpreadsheetPath value in the configuration block.' 'ERROR'
        exit 1
    }

    # 3. Confirm the ImportExcel module is installed
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Log "The required PowerShell module 'ImportExcel' is not installed." 'ERROR'
        Write-Log 'To install it, open PowerShell and run this command:' 'ERROR'
        Write-Log '    Install-Module ImportExcel -Scope CurrentUser' 'ERROR'
        Write-Log 'Then re-run this script.' 'ERROR'
        exit 1
    }
    Import-Module ImportExcel -ErrorAction Stop

    # 4. Confirm the PSAuthClient module is installed (for OAuth browser flow)
    if (-not (Get-Module -ListAvailable -Name PSAuthClient)) {
        Write-Log "The required PowerShell module 'PSAuthClient' is not installed." 'ERROR'
        Write-Log 'To install it, open PowerShell and run this command:' 'ERROR'
        Write-Log '    Install-Module PSAuthClient -Scope CurrentUser' 'ERROR'
        Write-Log 'Then re-run this script.' 'ERROR'
        exit 1
    }
    Import-Module PSAuthClient -ErrorAction Stop

    # 5. Confirm the worksheet exists inside the workbook
    try {
        $sheetNames = Get-ExcelSheetInfo -Path $Config.SpreadsheetPath -ErrorAction Stop |
                      Select-Object -ExpandProperty Name
    } catch {
        Write-Log "Could not read the spreadsheet at $($Config.SpreadsheetPath): $($_.Exception.Message)" 'ERROR'
        exit 1
    }
    if ($sheetNames -notcontains $Config.WorksheetName) {
        Write-Log "The workbook does not contain a worksheet named '$($Config.WorksheetName)'." 'ERROR'
        Write-Log "Sheets found in the workbook: $($sheetNames -join ', ')" 'ERROR'
        Write-Log 'Update the WorksheetName value in the configuration block.' 'ERROR'
        exit 1
    }

    Write-Log 'Pre-flight checks passed.' 'SUCCESS'
}


# ----------------------------------------------------------------------------
# Save-TokenCache
#   Encrypts a refresh token using Windows DPAPI (current-user scope) and
#   writes it to disk. Only the same Windows user on the same machine can
#   decrypt the file.
# ----------------------------------------------------------------------------
function Save-TokenCache {
    param([Parameter(Mandatory)] [string] $RefreshToken)

    try {
        # Convert plain string -> SecureString -> encrypted standard string (DPAPI).
        # Without -Key, ConvertFrom-SecureString uses the current user's DPAPI key.
        $secure    = ConvertTo-SecureString -String $RefreshToken -AsPlainText -Force
        $encrypted = ConvertFrom-SecureString -SecureString $secure
        Set-Content -Path $Config.TokenCachePath -Value $encrypted -Encoding UTF8 -ErrorAction Stop
        Write-Log "Saved encrypted refresh token to $($Config.TokenCachePath)." 'INFO'
    } catch {
        Write-Log "Could not save the token cache: $($_.Exception.Message)" 'WARN'
    }
}


# ----------------------------------------------------------------------------
# Get-CachedRefreshToken
#   Reads and decrypts the cached refresh token, or returns $null if the
#   cache doesn't exist or can't be decrypted.
# ----------------------------------------------------------------------------
function Get-CachedRefreshToken {
    if (-not (Test-Path -LiteralPath $Config.TokenCachePath)) {
        return $null
    }
    try {
        $encrypted = Get-Content -Path $Config.TokenCachePath -Raw -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($encrypted)) { return $null }

        $secure = ConvertTo-SecureString -String $encrypted.Trim() -ErrorAction Stop

        # Pull the plain string back out of the SecureString
        $bstr  = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
        try {
            return [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
        } finally {
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        }
    } catch {
        Write-Log "Cached token at $($Config.TokenCachePath) could not be decrypted (this is normal if you switched user accounts or machines). A fresh browser login will be required." 'WARN'
        return $null
    }
}


# ----------------------------------------------------------------------------
# Invoke-InteractiveLogin
#   Launches a browser via PSAuthClient so the user can log in to NinjaOne
#   and approve the application. Exchanges the resulting auth code for an
#   access token + refresh token, and persists the refresh token to disk.
#   Returns the access token.
# ----------------------------------------------------------------------------
function Invoke-InteractiveLogin {
    Write-Log 'Opening your browser so you can log in to NinjaOne. Please log in and approve the application — this only needs to happen the first time.' 'INFO'
    Write-Log "(After approving, your browser may show a 'site can't be reached' page at https://localhost. That's expected — just leave it and come back here.)" 'INFO'

    $authParams = @{
        Uri              = "$($Config.BaseUrl)/ws/oauth/authorize"
        Redirect_uri     = $Config.RedirectUri
        Client_id        = $Config.ClientId
        Scope            = $Script:Scope
        UsePkce          = $false
        CustomParameters = @{ client_secret = $Config.ClientSecret }
    }

    try {
        $auth = Invoke-OAuth2AuthorizationEndpoint @authParams
    } catch {
        Write-Log "Browser login failed: $($_.Exception.Message)" 'ERROR'
        Write-Log "Most likely causes:" 'ERROR'
        Write-Log "    - Redirect URI mismatch — must be EXACTLY 'https://localhost' in both NinjaOne and this script (no port, no slash, no path)." 'ERROR'
        Write-Log "    - Client ID is wrong, or the API client app does not have the Authorization Code grant type enabled." 'ERROR'
        exit 1
    }

    if (-not $auth.code) {
        Write-Log 'Browser login did not return an authorization code. Aborting.' 'ERROR'
        exit 1
    }

    Write-Log 'Authorization code received. Exchanging it for an access token...' 'INFO'

    $tokenParams = @{
        uri           = "$($Config.BaseUrl)/ws/oauth/token"
        redirect_uri  = $Config.RedirectUri
        client_id     = $Config.ClientId
        client_secret = $Config.ClientSecret
        code          = $auth.code
    }

    try {
        $token = Invoke-OAuth2TokenEndpoint @tokenParams
    } catch {
        Write-Log "Token exchange failed: $($_.Exception.Message)" 'ERROR'
        exit 1
    }

    if (-not $token.access_token) {
        Write-Log 'Token endpoint did not return an access_token.' 'ERROR'
        exit 1
    }

    if ($token.refresh_token) {
        Save-TokenCache -RefreshToken $token.refresh_token
    } else {
        Write-Log 'No refresh token was returned. You will be prompted to log in again next run.' 'WARN'
        Write-Log "Make sure 'offline_access' is included in scope and 'Refresh Token' is checked in NinjaOne." 'WARN'
    }

    Write-Log 'Login successful.' 'SUCCESS'
    return [string]$token.access_token
}


# ----------------------------------------------------------------------------
# Invoke-RefreshTokenLogin
#   Uses the cached refresh token to silently obtain a new access token.
#   Throws RefreshTokenInvalidException if the saved token is missing or
#   has been rejected by NinjaOne (so the caller can fall back to the
#   browser login).
# ----------------------------------------------------------------------------
function Invoke-RefreshTokenLogin {
    $refreshToken = Get-CachedRefreshToken
    if (-not $refreshToken) {
        throw [RefreshTokenInvalidException]::new('No cached refresh token found.')
    }

    Write-Log 'Using saved login to get a new access token...' 'INFO'

    $body = @{
        grant_type    = 'refresh_token'
        refresh_token = $refreshToken
        client_id     = $Config.ClientId
        client_secret = $Config.ClientSecret
    }

    try {
        $token = Invoke-RestMethod `
            -Method      Post `
            -Uri         "$($Config.BaseUrl)/ws/oauth/token" `
            -Body        $body `
            -ContentType 'application/x-www-form-urlencoded' `
            -ErrorAction Stop
    } catch {
        # Try to detect "invalid_grant" / 400 which means the refresh token is dead
        $statusCode = $null
        if ($_.Exception.Response) {
            try { $statusCode = [int]$_.Exception.Response.StatusCode } catch {}
        }
        $msg = $_.Exception.Message
        if ($statusCode -eq 400 -or $msg -match 'invalid_grant') {
            throw [RefreshTokenInvalidException]::new('Saved login has expired or been revoked.')
        }
        # Anything else is a real error — let it bubble up
        throw
    }

    if (-not $token.access_token) {
        throw [RefreshTokenInvalidException]::new('Refresh response did not include an access_token.')
    }

    # Rotate the cached refresh token if NinjaOne issued a new one
    if ($token.refresh_token -and $token.refresh_token -ne $refreshToken) {
        Save-TokenCache -RefreshToken $token.refresh_token
    }

    Write-Log 'Got a new access token from saved login.' 'SUCCESS'
    return [string]$token.access_token
}


# ----------------------------------------------------------------------------
# Get-AuthToken
#   Main entry point for getting a usable access token. Tries (in order):
#     1. In-memory cached access token from earlier in this run
#     2. Refresh token flow (silent)
#     3. Interactive browser login
#   Pass -Force to skip the in-memory cache (used when a 401 comes back
#   mid-import and we need a fresh token).
# ----------------------------------------------------------------------------
function Get-AuthToken {
    param([switch] $Force)

    if (-not $Force -and $Script:AccessToken) {
        return $Script:AccessToken
    }

    # Try refresh first (silent path)
    try {
        $Script:AccessToken = Invoke-RefreshTokenLogin
        return $Script:AccessToken
    } catch [RefreshTokenInvalidException] {
        Write-Log "Saved login is not usable: $($_.Exception.Message)" 'INFO'
        Write-Log 'Falling back to browser login.' 'INFO'
    }

    # Fall back to interactive browser login
    $Script:AccessToken = Invoke-InteractiveLogin
    return $Script:AccessToken
}


# ----------------------------------------------------------------------------
# Read-TicketSpreadsheet
#   Reads tickets from the configured worksheet and returns them as objects.
#   Header matching is case-insensitive (handled by ImportExcel).
# ----------------------------------------------------------------------------
function Read-TicketSpreadsheet {
    Write-Log "Reading tickets from '$($Config.SpreadsheetPath)' [worksheet: $($Config.WorksheetName)]..." 'INFO'

    try {
        $rows = Import-Excel -Path $Config.SpreadsheetPath -WorksheetName $Config.WorksheetName -ErrorAction Stop
    } catch {
        Write-Log "Failed to read the spreadsheet: $($_.Exception.Message)" 'ERROR'
        exit 1
    }

    if (-not $rows) {
        Write-Log 'The worksheet is empty — no rows to import.' 'WARN'
        return @()
    }

    Write-Log "Found $($rows.Count) row(s) in the worksheet." 'INFO'
    return $rows
}


# ----------------------------------------------------------------------------
# Test-TicketRow
#   Validates a single row. Returns $null if valid, otherwise a string
#   describing what is missing/wrong.
# ----------------------------------------------------------------------------
function Test-TicketRow {
    param([Parameter(Mandatory)] $Row)

    $missing = @()
    foreach ($field in @('Subject','Description','RequesterEmail')) {
        $value = $Row.$field
        if ([string]::IsNullOrWhiteSpace([string]$value)) {
            $missing += $field
        }
    }

    if ($missing.Count -gt 0) {
        return "Missing required field(s): $($missing -join ', ')"
    }
    return $null
}


# ----------------------------------------------------------------------------
# ConvertTo-TicketPayload
#   Converts a spreadsheet row into the JSON payload expected by the API.
#   Empty/blank fields are sent as $null so the API doesn't store empty strings.
# ----------------------------------------------------------------------------
function ConvertTo-TicketPayload {
    param([Parameter(Mandatory)] $Row)

    # Helper: return $null for blank values, trimmed string otherwise
    $clean = {
        param($v)
        if ($null -eq $v) { return $null }
        $s = [string]$v
        if ([string]::IsNullOrWhiteSpace($s)) { return $null }
        return $s.Trim()
    }

    # Tags column is comma-separated; turn it into an array
    $tags = $null
    $rawTags = & $clean $Row.Tags
    if ($rawTags) {
        $tags = $rawTags -split ',' |
                ForEach-Object { $_.Trim() } |
                Where-Object   { $_ -ne '' }
    }

    [pscustomobject]@{
        external_id      = & $clean $Row.TicketID
        subject          = & $clean $Row.Subject
        description      = & $clean $Row.Description
        status           = & $clean $Row.Status
        priority         = & $clean $Row.Priority
        requester_email  = & $clean $Row.RequesterEmail
        assignee_email   = & $clean $Row.AssigneeEmail
        created_at       = & $clean $Row.CreatedDate
        category         = & $clean $Row.Category
        tags             = $tags
    }
}


# ----------------------------------------------------------------------------
# Wait-Throttle
#   Sleeps if necessary to stay under 5 requests per second.
# ----------------------------------------------------------------------------
function Wait-Throttle {
    $elapsedMs = (New-TimeSpan -Start $Script:LastRequestAt -End (Get-Date)).TotalMilliseconds
    if ($elapsedMs -lt $Script:MinIntervalMs) {
        Start-Sleep -Milliseconds ([int]($Script:MinIntervalMs - $elapsedMs))
    }
    $Script:LastRequestAt = Get-Date
}


# ----------------------------------------------------------------------------
# Import-Ticket
#   Sends a single ticket to the NinjaOne API with retry/backoff.
#   Returns $true on success, $false on permanent failure (after retries).
#   Stores any error message in [ref]$ErrorMessage for the caller to log.
# ----------------------------------------------------------------------------
function Import-Ticket {
    param(
        [Parameter(Mandatory)] $Payload,
        [Parameter(Mandatory)] [int] $RowNumber,
        [Parameter(Mandatory)] [ref] $ErrorMessage
    )

    $url    = "$($Config.BaseUrl.TrimEnd('/'))/api/v2/ticketing/ticket"
    $json   = $Payload | ConvertTo-Json -Depth 6
    $delays = @(2, 4, 8)   # Backoff seconds between attempts 1->2, 2->3, 3->fail
    $tokenAlreadyRefreshed = $false

    for ($attempt = 1; $attempt -le 3; $attempt++) {

        Wait-Throttle
        $headers = @{
            'Authorization' = "Bearer $(Get-AuthToken)"
            'Accept'        = 'application/json'
        }

        try {
            $null = Invoke-RestMethod `
                -Method      Post `
                -Uri         $url `
                -Headers     $headers `
                -Body        $json `
                -ContentType 'application/json' `
                -ErrorAction Stop
            return $true

        } catch {
            # Try to extract the HTTP status code from the response
            $statusCode = $null
            if ($_.Exception.Response) {
                try { $statusCode = [int]$_.Exception.Response.StatusCode } catch {}
            }

            # 401 -> token may have expired mid-run. Refresh once and retry the SAME row.
            if ($statusCode -eq 401 -and -not $tokenAlreadyRefreshed) {
                Write-Log "Row $RowNumber attempt $attempt — 401 Unauthorized. Refreshing token and retrying..." 'WARN'
                try {
                    Get-AuthToken -Force | Out-Null
                    $tokenAlreadyRefreshed = $true
                    continue
                } catch {
                    $ErrorMessage.Value = "Could not refresh token: $($_.Exception.Message)"
                    return $false
                }
            }

            # Retryable transient errors
            $transient = @(429, 500, 502, 503, 504)
            if ($transient -contains $statusCode -and $attempt -lt 3) {
                $wait = $delays[$attempt - 1]
                Write-Log "Row $RowNumber attempt $attempt — HTTP $statusCode. Waiting ${wait}s and retrying..." 'WARN'
                Start-Sleep -Seconds $wait
                continue
            }

            # Permanent failure — build a friendly error message
            $friendly = switch ($statusCode) {
                400     { '400 Bad Request — the ticket data was rejected. Check field values (Status, Priority, dates, etc.).' }
                401     { '401 Unauthorized — check your Client ID, Client Secret, and that the API client app has Authorization Code + Refresh Token grant types enabled.' }
                403     { '403 Forbidden — your API client app is missing the Management scope. Edit the app in NinjaOne to enable it.' }
                404     { '404 Not Found — check the BaseUrl in the configuration block.' }
                409     { '409 Conflict — a ticket with this ID may already exist.' }
                422     { '422 Unprocessable Entity — one or more fields failed validation on the server.' }
                default { if ($statusCode) { "HTTP $statusCode — $($_.Exception.Message)" } else { $_.Exception.Message } }
            }
            $ErrorMessage.Value = $friendly
            return $false
        }
    }

    $ErrorMessage.Value = 'Failed after 3 attempts'
    return $false
}


# ----------------------------------------------------------------------------
# Write-Summary
#   Prints a clean end-of-run summary to console and log file.
# ----------------------------------------------------------------------------
function Write-Summary {
    Write-Host ''
    Write-Host '============================================================' -ForegroundColor Cyan
    Write-Host '                    IMPORT SUMMARY' -ForegroundColor Cyan
    Write-Host '============================================================' -ForegroundColor Cyan

    if ($Config.DryRun) {
        Write-Host 'DRY RUN — no data was actually sent to the API.' -ForegroundColor Yellow
    }

    Write-Host ("Total rows read:        {0}" -f $Script:Stats.TotalRows)
    Write-Host ("Imported successfully:  {0}" -f $Script:Stats.Imported) -ForegroundColor Green
    Write-Host ("Skipped (validation):   {0}" -f $Script:Stats.Skipped.Count) -ForegroundColor Yellow
    Write-Host ("Failed (API errors):    {0}" -f $Script:Stats.Failed.Count)  -ForegroundColor Red

    if ($Script:Stats.Skipped.Count -gt 0) {
        Write-Host ''
        Write-Host 'Skipped rows:' -ForegroundColor Yellow
        foreach ($s in $Script:Stats.Skipped) {
            Write-Host ("  Row {0}: {1}" -f $s.Row, $s.Reason) -ForegroundColor Yellow
        }
    }

    if ($Script:Stats.Failed.Count -gt 0) {
        Write-Host ''
        Write-Host 'Failed rows:' -ForegroundColor Red
        foreach ($f in $Script:Stats.Failed) {
            Write-Host ("  Row {0}: {1}" -f $f.Row, $f.Reason) -ForegroundColor Red
        }
    }

    Write-Host '============================================================' -ForegroundColor Cyan
    Write-Host ("Full log saved to: {0}" -f $Config.LogFilePath) -ForegroundColor Cyan
    Write-Host ''

    # Mirror the summary into the log file
    Write-Log ("SUMMARY — Total: $($Script:Stats.TotalRows), Imported: $($Script:Stats.Imported), Skipped: $($Script:Stats.Skipped.Count), Failed: $($Script:Stats.Failed.Count)") 'INFO'
}


# ============================================================================
# MAIN
# ============================================================================

# Banner
Write-Host ''
Write-Host '============================================================' -ForegroundColor Cyan
Write-Host '       NinjaOne Ticket Import — Spreadsheet to API' -ForegroundColor Cyan
Write-Host '============================================================' -ForegroundColor Cyan
if ($Config.DryRun) {
    Write-Host 'DRY RUN MODE — no data will be sent to NinjaOne.' -ForegroundColor Yellow
    Write-Host 'Set $Config.DryRun = $false in the configuration block to run for real.' -ForegroundColor Yellow
}
Write-Host ''

# Start log file fresh for this run
"=== Ticket Import run started $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ===" |
    Out-File -FilePath $Config.LogFilePath -Encoding utf8 -Force

# 1. Validate configuration
Test-Configuration

# 2. Authenticate (skipped on dry runs so customers can validate without logging in yet)
if (-not $Config.DryRun) {
    Get-AuthToken | Out-Null
} else {
    Write-Log 'Skipping authentication because DryRun is enabled.' 'INFO'
    Write-Log '(On the real run, your browser will open the first time so you can log in to NinjaOne.)' 'INFO'
}

# 3. Read the spreadsheet
$rows = Read-TicketSpreadsheet
$Script:Stats.TotalRows = @($rows).Count

# 4. Process each row
$rowNumber = 1   # Header is row 1 in Excel; data rows start at 2
foreach ($row in $rows) {
    $rowNumber++

    # Validate
    $validationError = Test-TicketRow -Row $row
    if ($validationError) {
        Write-Log "Row $rowNumber — SKIPPED. $validationError" 'WARN'
        $Script:Stats.Skipped.Add([pscustomobject]@{ Row = $rowNumber; Reason = $validationError })
        continue
    }

    # Build payload
    $payload = ConvertTo-TicketPayload -Row $row
    $label   = if ($payload.external_id) { $payload.external_id } else { $payload.subject }

    if ($Config.DryRun) {
        Write-Log "Row $rowNumber — [DRY RUN] would import: $label" 'INFO'
        $Script:Stats.Imported++
        continue
    }

    # Send to API
    $errMsg = ''
    $success = Import-Ticket -Payload $payload -RowNumber $rowNumber -ErrorMessage ([ref]$errMsg)
    if ($success) {
        Write-Log "Row $rowNumber — imported: $label" 'SUCCESS'
        $Script:Stats.Imported++
    } else {
        Write-Log "Row $rowNumber — FAILED: $errMsg" 'ERROR'
        $Script:Stats.Failed.Add([pscustomobject]@{ Row = $rowNumber; Reason = $errMsg })
    }
}

# 5. Final summary
Write-Summary
