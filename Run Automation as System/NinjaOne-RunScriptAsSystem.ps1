# ==============================================================================
#  NinjaOne — Run Script as SYSTEM via Authorization Code Flow
# ==============================================================================
#
#  BEFORE YOU RUN THIS SCRIPT:
#
#  1. Log in to NinjaOne and go to:
#        Administration > Apps > API > Client App
#     Create a new app (or open an existing one) to find your Client ID and Secret.
#
#  2. Set the Redirect URI in that same app to exactly:
#        https://localhost
#
#  3. Fill in the five fields below — that is the only section you need to touch.
#
#  4. Set $DeviceId to the numeric ID of the device you want to run the script on.
#     (Find it in NinjaOne under the device details URL — the number at the end.)
#
#  5. Put the commands you want to run inside the $ScriptBody block below.
#
# ==============================================================================
#  CONFIGURATION — only edit values in this section
# ==============================================================================

$BaseUrl         = 'https://<your login URL>'              # e.g. https://app.ninjarmm.com
$TokenEndpoint   = 'https://<your login URL>/ws/oauth/token'
$ClientId        = '<Your Client ID>'
$ClientSecret    = '<Your Client Secret>'

# The device ID number from NinjaOne (found in the device's URL)
$DeviceId        = '<Your Device ID>'

# The PowerShell commands you want to run on the device as SYSTEM
$ScriptBody      = @'
Write-Output "Hello from SYSTEM context"
whoami
'@

# ==============================================================================
#  DO NOT EDIT BELOW THIS LINE
# ==============================================================================

$RedirectUri   = 'https://localhost'
$AuthEndpoint  = "$BaseUrl/ws/oauth/authorize"
$Scope         = 'monitoring management control'
$ScriptType    = 'POWERSHELL'
$RunAs         = 'SYSTEM'

# --- Step 1: Open the NinjaOne login/authorization page in the browser --------

Write-Host ''
Write-Host '---------------------------------------------------------------'
Write-Host ' Step 1 of 3 — Opening your browser to log in to NinjaOne...'
Write-Host '---------------------------------------------------------------'

$State     = [System.Guid]::NewGuid().ToString('N')
$AuthQuery = "response_type=code" +
             "&client_id=$([Uri]::EscapeDataString($ClientId))" +
             "&redirect_uri=$([Uri]::EscapeDataString($RedirectUri))" +
             "&scope=$([Uri]::EscapeDataString($Scope))" +
             "&state=$State"

Start-Process "$AuthEndpoint`?$AuthQuery"

Write-Host ''
Write-Host ' Your browser should open a NinjaOne login page.'
Write-Host ' Log in and click Authorize. The browser will then show'
Write-Host ' a blank or "connection refused" page — that is normal.'
Write-Host ''

# --- Step 2: Read the authorization code from the URL the user pastes back ---
#
#  NOTE: Because https://localhost (no port) cannot be intercepted by a local
#  HTTP listener, the browser will land on a blank page after you authorize.
#  Copy the full URL from your browser address bar and paste it below.

Write-Host '---------------------------------------------------------------'
Write-Host ' Step 2 of 3 — Paste the URL from your browser address bar'
Write-Host '---------------------------------------------------------------'
Write-Host ''
Write-Host ' After authorizing, your browser will show a blank page.'
Write-Host ' Copy the entire URL from the address bar and paste it here:'
Write-Host ''

$PastedUrl = Read-Host ' Paste URL here'

$Uri    = [System.Uri]$PastedUrl
$Params = [System.Web.HttpUtility]::ParseQueryString($Uri.Query)

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
Write-Host ' Authorization code received. Exchanging for access token...'

# --- Step 3: Exchange the code for an access token ---------------------------

Write-Host ''
Write-Host '---------------------------------------------------------------'
Write-Host ' Step 3 of 3 — Connecting to NinjaOne and running your script'
Write-Host '---------------------------------------------------------------'

$TokenBody = @{
    grant_type    = 'authorization_code'
    code          = $AuthCode
    redirect_uri  = $RedirectUri
    client_id     = $ClientId
    client_secret = $ClientSecret
}

try {
    $TokenResponse = Invoke-RestMethod -Method Post -Uri $TokenEndpoint `
        -ContentType 'application/x-www-form-urlencoded' `
        -Body $TokenBody
} catch {
    Write-Host ''
    Write-Host " ERROR: Could not get an access token. $_" -ForegroundColor Red
    Write-Host '        Double-check your ClientId, ClientSecret, and TokenEndpoint.' -ForegroundColor Red
    exit 1
}

$AccessToken = $TokenResponse.access_token
Write-Host ' Access token obtained successfully.'

# --- Run the script on the target device as SYSTEM ---------------------------

$Headers = @{
    Authorization  = "Bearer $AccessToken"
    'Content-Type' = 'application/json'
}

$Payload = @{
    type  = $ScriptType
    runAs = $RunAs
    body  = $ScriptBody
} | ConvertTo-Json -Depth 5

Write-Host ''
Write-Host " Sending script to device $DeviceId as SYSTEM..."

try {
    $RunResponse = Invoke-RestMethod -Method Post `
        -Uri "$BaseUrl/v2/device/$DeviceId/script/run" `
        -Headers $Headers `
        -Body $Payload
} catch {
    Write-Host ''
    Write-Host " ERROR: Failed to send the script. $_" -ForegroundColor Red
    Write-Host '        Make sure the Device ID is correct and the device is online.' -ForegroundColor Red
    exit 1
}

$ActivityId = $RunResponse.activityId
Write-Host " Script queued! Activity ID: $ActivityId"
Write-Host ''

# --- Poll for the result ------------------------------------------------------

$ResultUrl    = "$BaseUrl/v2/device/$DeviceId/scripting/activity/$ActivityId/result"
$MaxWaitSec   = 120
$PollInterval = 5
$Elapsed      = 0

Write-Host ' Waiting for the script to finish (this may take up to 2 minutes)...'

do {
    Start-Sleep -Seconds $PollInterval
    $Elapsed += $PollInterval

    try {
        $Result = Invoke-RestMethod -Method Get -Uri $ResultUrl -Headers $Headers
    } catch {
        Write-Host "  Still running... ($Elapsed seconds elapsed)"
        continue
    }

    if ($Result.status -in @('SUCCESS', 'FAILED', 'TIMED_OUT')) {
        Write-Host ''
        Write-Host "==============================================================="
        Write-Host " Script completed with status: $($Result.status)"
        Write-Host "==============================================================="
        Write-Host ''
        Write-Host '--- Output ---'
        Write-Host $Result.output
        if ($Result.errorOutput) {
            Write-Host ''
            Write-Host '--- Errors ---'
            Write-Host $Result.errorOutput
        }
        break
    }

    Write-Host "  Still running... ($Elapsed seconds elapsed)"

} while ($Elapsed -lt $MaxWaitSec)

if ($Elapsed -ge $MaxWaitSec) {
    Write-Host ''
    Write-Host " The script is still running after $MaxWaitSec seconds." -ForegroundColor Yellow
    Write-Host " You can check the result in NinjaOne under the device's activity log." -ForegroundColor Yellow
    Write-Host " Activity ID: $ActivityId" -ForegroundColor Yellow
}
