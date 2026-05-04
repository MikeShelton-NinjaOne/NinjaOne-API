<#
.SYNOPSIS
    Import-Tickets.ps1 — Bulk-import tickets from an Excel spreadsheet into our ticketing system via REST API.

.DESCRIPTION
    Designed for new customers migrating ticket data from another ticketing system.
    Reads tickets from an .xlsx file and creates them via the API one at a time, with
    validation, retry logic, throttling, and a full run log.

.HOW TO USE THIS SCRIPT (for first-time PowerShell users)
    Step 1.  Install the prerequisite module ONE TIME by opening PowerShell and running:
                Install-Module ImportExcel -Scope CurrentUser

    Step 2.  Open this .ps1 file in Notepad or VS Code.

    Step 3.  Scroll down to the "# === CONFIGURATION ===" block and replace every value
             that is wrapped in angle brackets (for example <Your Client ID>) with your
             real information. Do NOT remove the single quotes around the values.

    Step 4.  Save the file. Make sure DryRun is set to $true for your first run — this
             validates your spreadsheet WITHOUT sending any data to the API.

    Step 5.  Run the script. Either right-click the file and choose "Run with PowerShell",
             or open a PowerShell window in this folder and run:
                .\Import-Tickets.ps1

    After a successful dry run, change DryRun to $false and run the script again to
    perform the real import.

    YOU SHOULD ONLY EVER NEED TO EDIT THE # === CONFIGURATION === BLOCK BELOW.
#>

# === CONFIGURATION ===
# Edit the values below before running the script.
# Replace everything inside the angle brackets < > with your own information.
# Do NOT remove the quotation marks.

$Config = @{
    # The base URL of your ticketing system login page
    BaseUrl       = 'https://<your Login URL>'

    # The OAuth token endpoint (usually your login URL with /ws/oauth/token at the end)
    TokenEndpoint = 'https://<your Login URL>/ws/oauth/token'

    # Your API Client ID (provided by your administrator or found in the API/Integrations section of the portal)
    ClientId      = '<Your Client ID>'

    # Your API Client Secret (provided alongside the Client ID — keep this private)
    ClientSecret  = '<Your Client Secret>'

    # Redirect URI used during OAuth — leave this exactly as shown
    RedirectUri   = 'https://localhost'

    # Full path to the spreadsheet you want to import (example: C:\Migration\Tickets.xlsx)
    SpreadsheetPath = '<Full path to your .xlsx file>'

    # The name of the worksheet/tab inside the spreadsheet that contains the tickets
    WorksheetName = 'Tickets'

    # Where to save the run log (the script will create this file)
    LogFilePath   = '.\ticket-import-log.txt'

    # Set to $true for a test run that validates the file but does NOT send any data to the API.
    # Always run with $true the FIRST time to make sure your data looks right.
    DryRun        = $true
}

# ============================================================================
# DO NOT EDIT BELOW THIS LINE
# ============================================================================

# ---- Script-wide state ----
$Script:BearerToken    = $null     # Cached OAuth token for the duration of the run
$Script:Stats = @{
    TotalRows  = 0
    Imported   = 0
    Skipped    = [System.Collections.Generic.List[object]]::new()
    Failed     = [System.Collections.Generic.List[object]]::new()
}

# Throttle: max 5 requests per second  =>  minimum 200 ms between requests
$Script:MinIntervalMs  = 200
$Script:LastRequestAt  = [DateTime]::MinValue


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
#   the spreadsheet exists, the ImportExcel module is installed, and the
#   worksheet name is correct. Exits the script if anything is wrong.
# ----------------------------------------------------------------------------
function Test-Configuration {
    Write-Log "Running pre-flight configuration checks..." 'INFO'

    # 1. Check for unreplaced <placeholder> values in any string config field
    $placeholdersFound = @()
    foreach ($key in $Config.Keys) {
        $value = $Config[$key]
        if ($value -is [string] -and $value -match '[<>]') {
            $placeholdersFound += $key
        }
    }
    if ($placeholdersFound.Count -gt 0) {
        Write-Log "The following configuration field(s) still contain placeholder values that you need to replace:" 'ERROR'
        foreach ($f in $placeholdersFound) {
            Write-Log "    - $f  (current value: $($Config[$f]))" 'ERROR'
        }
        Write-Log "Open this script in Notepad or VS Code and edit the # === CONFIGURATION === block at the top." 'ERROR'
        exit 1
    }

    # 2. Confirm the spreadsheet file exists
    if (-not (Test-Path -LiteralPath $Config.SpreadsheetPath)) {
        Write-Log "Cannot find the spreadsheet file at: $($Config.SpreadsheetPath)" 'ERROR'
        Write-Log "Double-check the SpreadsheetPath value in the configuration block." 'ERROR'
        exit 1
    }

    # 3. Confirm the ImportExcel module is installed
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Log "The required PowerShell module 'ImportExcel' is not installed." 'ERROR'
        Write-Log "To install it, open PowerShell and run this command:" 'ERROR'
        Write-Log "    Install-Module ImportExcel -Scope CurrentUser" 'ERROR'
        Write-Log "Then re-run this script." 'ERROR'
        exit 1
    }
    Import-Module ImportExcel -ErrorAction Stop

    # 4. Confirm the worksheet exists inside the workbook
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
        Write-Log "Update the WorksheetName value in the configuration block." 'ERROR'
        exit 1
    }

    Write-Log "Pre-flight checks passed." 'SUCCESS'
}


# ----------------------------------------------------------------------------
# Get-AuthToken
#   Requests an OAuth 2.0 bearer token using the client_credentials grant type.
#   Caches the token in $Script:BearerToken so it can be reused across rows.
#   Pass -Force to ignore the cache and request a brand new token.
# ----------------------------------------------------------------------------
function Get-AuthToken {
    param([switch] $Force)

    if (-not $Force -and $Script:BearerToken) {
        return $Script:BearerToken
    }

    Write-Log "Requesting OAuth access token from $($Config.TokenEndpoint)..." 'INFO'

    # Build the form-encoded body for the token request
    $body = @{
        grant_type    = 'client_credentials'
        client_id     = $Config.ClientId
        client_secret = $Config.ClientSecret
        redirect_uri  = $Config.RedirectUri
    }

    try {
        $response = Invoke-RestMethod `
            -Method  Post `
            -Uri     $Config.TokenEndpoint `
            -Body    $body `
            -ContentType 'application/x-www-form-urlencoded' `
            -ErrorAction Stop

        if (-not $response.access_token) {
            Write-Log "Authentication response did not contain an access_token." 'ERROR'
            exit 1
        }

        $Script:BearerToken = $response.access_token
        Write-Log "Authentication successful." 'SUCCESS'
        return $Script:BearerToken

    } catch {
        Write-Log "Authentication failed: $($_.Exception.Message)" 'ERROR'
        Write-Log "Common causes:" 'ERROR'
        Write-Log "    - Wrong Client ID or Client Secret" 'ERROR'
        Write-Log "    - Wrong BaseUrl or TokenEndpoint" 'ERROR'
        Write-Log "    - Network or firewall blocking the request" 'ERROR'
        exit 1
    }
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
        Write-Log "The worksheet is empty — no rows to import." 'WARN'
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
#   Sends a single ticket to the API with retry/backoff.
#   Returns $true on success, $false on permanent failure (after retries).
#   Stores any error message in [ref]$ErrorMessage for the caller to log.
# ----------------------------------------------------------------------------
function Import-Ticket {
    param(
        [Parameter(Mandatory)] $Payload,
        [Parameter(Mandatory)] [int] $RowNumber,
        [Parameter(Mandatory)] [ref] $ErrorMessage
    )

    $url    = "$($Config.BaseUrl.TrimEnd('/'))/tickets"
    $json   = $Payload | ConvertTo-Json -Depth 6
    $delays = @(2, 4, 8)   # Backoff seconds between attempts 1->2, 2->3, 3->fail

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

            # 401 -> token may have expired mid-run. Force-refresh and retry once.
            if ($statusCode -eq 401) {
                Write-Log "Row $RowNumber attempt $attempt — 401 Unauthorized. Refreshing token and retrying..." 'WARN'
                Get-AuthToken -Force | Out-Null
                continue
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
                401     { '401 Unauthorized — check your Client ID and Client Secret.' }
                403     { '403 Forbidden — your API client does not have permission to create tickets.' }
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
Write-Host '          Ticket Import — Spreadsheet to API' -ForegroundColor Cyan
Write-Host '============================================================' -ForegroundColor Cyan
if ($Config.DryRun) {
    Write-Host 'DRY RUN MODE — no data will be sent to the API.' -ForegroundColor Yellow
    Write-Host 'Set $Config.DryRun = $false in the configuration block to run for real.' -ForegroundColor Yellow
}
Write-Host ''

# Start log file fresh for this run
"=== Ticket Import run started $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ===" |
    Out-File -FilePath $Config.LogFilePath -Encoding utf8 -Force

# 1. Validate configuration
Test-Configuration

# 2. Authenticate (skipped on dry runs so customers can validate without real credentials)
if (-not $Config.DryRun) {
    Get-AuthToken | Out-Null
} else {
    Write-Log "Skipping authentication because DryRun is enabled." 'INFO'
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
