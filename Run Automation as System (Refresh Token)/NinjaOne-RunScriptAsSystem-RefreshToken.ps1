# ==============================================================================
#  NinjaOne — Run Script as SYSTEM (Refresh Token Flow)
# ==============================================================================
#
#  HOW THIS SCRIPT WORKS:
#
#  The first time you run this script, it will open your browser so you can
#  log in to NinjaOne and authorize it. After that, it saves a refresh token
#  to a file on your computer so you never have to log in again — it will
#  renew its own access automatically every time you run it.
#
#  BEFORE YOU RUN THIS SCRIPT THE FIRST TIME:
#
#  1. Log in to NinjaOne and go to:
#        Administration > Apps > API > Client App
#     Create a new app (or open an existing one) to find your Client ID and Secret.
#
#  2. In that same app, make sure the Redirect URI is set to exactly:
#        https://localhost
#
#  3. Fill in the values below — that is the only section you need to touch.
#
#  4. Set $DeviceId to the numeric ID of the device you want to run the script on.
#     (Find it in NinjaOne by opening the device — the number at the end of the URL.)
#
#  5. Put the commands you want to run inside the $ScriptBody block below.
#
# ==============================================================================
#  CONFIGURATION — only edit values in this section
# ==============================================================================

$BaseUrl         = 'https://<your login URL>'               # e.g. https://app.ninjarmm.com
$TokenEndpoint   = 'https://<your login URL>/ws/oauth/token'
$ClientId        = '<Your Client ID>'
$ClientSecret    = '<Your Client Secret>'

# The device ID number from NinjaOne (found in the device's URL)
$DeviceId        = '<Your Device ID>'

# Where to save the refresh token on this computer.
# The default puts it in the same folder as this script.
# You can change this to any path you prefer, e.g. 'C:\Scripts\ninja_token.txt'
$TokenFile       = Join-Path $PSScriptRoot 'ninja_refresh_token.txt'

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
$Scope         = 'monitoring management control offline_access'
$ScriptType    = 'POWERSHELL'
$RunAs         = 'SYSTEM'

# ------------------------------------------------------------------------------
#  FUNCTION: Exchange a refresh token for a new access token
# ------------------------------------------------------------------------------

function Get-AccessTokenFromRefreshToken {
    param (
        [string]$RefreshToken
    )

    $Body = @{
        grant_type    = 'refresh_token'
        refresh_token = $RefreshToken
        client_id     = $ClientId
        client_secret = $ClientSecret
    }

    try {
        $Response = Invoke-RestMethod -Method Post -Uri $TokenEndpoint `
            -ContentType 'application/x-www-form-urlencoded' `
            -Body $Body
        return $Response
    } catch {
        return $null
    }
}

# ------------------------------------------------------------------------------
#  FUNCTION: Run the full browser-based authorization code flow
# ------------------------------------------------------------------------------

function Invoke-AuthCodeFlow {
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
    Write-Host ' Authorization code received. Exchanging for tokens...'

    Write-Host ''
    Write-Host '---------------------------------------------------------------'
    Write-Host ' Step 3 of 3 — Connecting to NinjaOne...'
    Write-Host '---------------------------------------------------------------'

    $TokenBody = @{
        grant_type    = 'authorization_code'
        code          = $AuthCode
        redirect_uri  = $RedirectUri
        client_id     = $ClientId
        client_secret = $ClientSecret
    }

    try {
        $Response = Invoke-RestMethod -Method Post -Uri $TokenEndpoint `
            -ContentType 'application/x-www-form-urlencoded' `
            -Body $TokenBody
    } catch {
        Write-Host ''
        Write-Host " ERROR: Could not get an access token. $_" -ForegroundColor Red
        Write-Host '        Double-check your ClientId, ClientSecret, and TokenEndpoint.' -ForegroundColor Red
        exit 1
    }

    return $Response
}

# ------------------------------------------------------------------------------
#  MAIN — Get a valid access token (refresh token if available, else full login)
# ------------------------------------------------------------------------------

$AccessToken   = $null
$RefreshToken  = $null

# --- Try to use an existing saved refresh token first -------------------------

if (Test-Path $TokenFile) {
    Write-Host ''
    Write-Host ' Found saved refresh token. Attempting to renew access...'

    $SavedRefreshToken = (Get-Content $TokenFile -Raw).Trim()
    $TokenResponse     = Get-AccessTokenFromRefreshToken -RefreshToken $SavedRefreshToken

    if ($TokenResponse -and $TokenResponse.access_token) {
        $AccessToken  = $TokenResponse.access_token
        $RefreshToken = $TokenResponse.refresh_token

        Write-Host ' Access token renewed successfully using saved refresh token.'
    } else {
        Write-Host ' Saved refresh token has expired or is invalid.' -ForegroundColor Yellow
        Write-Host ' A new browser login is required.'
        Remove-Item $TokenFile -Force
    }
}

# --- If no valid token yet, do the full browser login -------------------------

if (-not $AccessToken) {
    $TokenResponse = Invoke-AuthCodeFlow

    if (-not $TokenResponse.access_token) {
        Write-Host ''
        Write-Host ' ERROR: Did not receive an access token after login.' -ForegroundColor Red
        exit 1
    }

    $AccessToken  = $TokenResponse.access_token
    $RefreshToken = $TokenResponse.refresh_token

    Write-Host ' Access token obtained successfully.'
}

# --- Save the latest refresh token for next time ------------------------------

if ($RefreshToken) {
    $RefreshToken | Set-Content $TokenFile -Force
    Write-Host " Refresh token saved to: $TokenFile"
} else {
    Write-Host ''
    Write-Host ' WARNING: NinjaOne did not return a refresh token.' -ForegroundColor Yellow
    Write-Host '          You may need to log in again next time.' -ForegroundColor Yellow
    Write-Host '          Make sure "offline_access" is enabled on your API app.' -ForegroundColor Yellow
}

# ------------------------------------------------------------------------------
#  Run the script on the target device as SYSTEM
# ------------------------------------------------------------------------------

Write-Host ''
Write-Host '---------------------------------------------------------------'
Write-Host " Sending script to device $DeviceId as SYSTEM..."
Write-Host '---------------------------------------------------------------'

$Headers = @{
    Authorization  = "Bearer $AccessToken"
    'Content-Type' = 'application/json'
}

$Payload = @{
    type  = $ScriptType
    runAs = $RunAs
    body  = $ScriptBody
} | ConvertTo-Json -Depth 5

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

# ------------------------------------------------------------------------------
#  Poll for the result
# ------------------------------------------------------------------------------

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
        Write-Host '==============================================================='
        Write-Host " Script completed with status: $($Result.status)"
        Write-Host '==============================================================='
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
