<#
.SYNOPSIS
    Configures (or replaces) a NinjaOne live-feed webhook so Ninja activity
    events stream into a SIEM such as Splunk HEC. Pure PowerShell, no modules.

.DESCRIPTION
    Walks an admin through the OAuth 2.0 Authorization Code flow:
      1) Opens NinjaOne's /ws/oauth/authorize URL in the default browser.
      2) Admin logs in and consents. Ninja redirects to https://localhost/?code=...
         The browser will show a "site can't be reached" / cert warning page.
         That is EXPECTED. The admin pastes the full URL back into PowerShell.
      3) Exchanges the auth code for an access token at /ws/oauth/token.
      4) Calls PUT {BaseUrl}/v2/webhook with the configured payload.
      5) On 204 No Content, prints the verification hint.

    No external modules. Works on Windows PowerShell 5.1 and PowerShell 7+.

.NOTES
    Requirements on the NinjaOne side (one-time setup):
      - You are a NinjaOne System Administrator.
      - An OAuth Client App exists at: Administration -> Apps -> API ->
        Client App IDs, with:
            Application Platform : Web
            Redirect URI         : https://localhost      (no port, no slash)
            Allowed Grant Types  : Authorization Code, Refresh Token
            Scopes               : Monitoring  (and offline_access, optional)

    See README.md in this folder for a step-by-step walkthrough.
#>

[CmdletBinding()]
param()

# Stop on errors so failures don't get swallowed
$ErrorActionPreference = 'Stop'

# ============================================================================
# BLOCK 1 - CREDENTIALS
# ============================================================================
# Replace each <placeholder> below.
#
# BaseUrl is your NinjaOne login URL. Pick the one matching your region:
#     US        https://app.ninjarmm.com
#     US (alt)  https://us2.ninjarmm.com
#     Canada    https://ca.ninjarmm.com
#     EU        https://eu.ninjarmm.com
#     Oceania   https://oc.ninjarmm.com
#
# TokenEndpoint and AuthEndpoint MUST use the same host as BaseUrl.
#
# ClientId and ClientSecret come from the OAuth Client App you created in
# NinjaOne (Administration -> Apps -> API -> Client App IDs).
#
# RedirectUri MUST exactly match what's saved on the OAuth app in Ninja.
# Use https://localhost  (no port, no trailing slash).

$Config = [ordered]@{
    BaseUrl       = 'https://<your Login URL>'
    TokenEndpoint = 'https://<your login URL>/ws/oauth/token'
    AuthEndpoint  = 'https://<your login URL>/ws/oauth/authorize'
    ClientId      = '<Your Client ID>'
    ClientSecret  = '<Your Client Secert>'
    RedirectUri   = 'https://localhost'
    Scope         = 'monitoring offline_access'
}

# ============================================================================
# BLOCK 2 - WEBHOOK PAYLOAD
# ============================================================================
# This is the body of the PUT /v2/webhook request.
#
#   url        : where Ninja POSTs each activity event. For Splunk HEC this
#                MUST end in /services/collector/raw .
#
#   expand     : extra objects to inline in each event payload. Default is
#                deviceId only; this script asks for device, organization,
#                and location so SIEM events arrive enriched.
#
#   headers    : auth/custom headers Ninja attaches to outbound POSTs.
#                Splunk HEC expects:   Authorization: Splunk <hec-token>
#                Most other SIEMs use: Authorization: Bearer <token>
#
#   activities : which Ninja activity types to stream. Each value is ["*"]
#                meaning "all sub-types of this category" (per v2 webhook
#                schema). Common categories:
#                  CONDITION  - alerts and resets
#                  ACTIONSET  - automations / technician actions
#                  SYSTEM     - system events
#                  ANTIVIRUS, MDM, PATCH_MANAGEMENT,
#                  SOFTWARE_PATCH_MANAGEMENT, TICKETING, SECURITY,
#                  MONITOR, SCHEDULED_TASK, REMOTE_TOOLS, SCRIPTING
#                Only include what your SIEM actually needs.

$WebhookPayload = [ordered]@{
    url     = 'https://<your-siem-host>:8088/services/collector/raw'
    expand  = @('device', 'organization', 'location')
    headers = @(
        [ordered]@{
            name  = 'Authorization'
            value = 'Splunk <your-HEC-token-here>'
        }
    )
    activities = [ordered]@{
        CONDITION = @('*')
        SYSTEM    = @('*')
        ACTIONSET = @('*')
    }
}

# ============================================================================
# Helper: write a colored status line
# ============================================================================
function Write-Status {
    param(
        [Parameter(Mandatory)] [string] $Message,
        [ValidateSet('Info','Ok','Warn','Err','Step')] [string] $Level = 'Info'
    )
    $color = switch ($Level) {
        'Ok'   { 'Green' }
        'Warn' { 'Yellow' }
        'Err'  { 'Red' }
        'Step' { 'Cyan' }
        default { 'Gray' }
    }
    $prefix = switch ($Level) {
        'Ok'   { '[ OK ]  ' }
        'Warn' { '[WARN]  ' }
        'Err'  { '[FAIL]  ' }
        'Step' { '[STEP]  ' }
        default { '        ' }
    }
    Write-Host ($prefix + $Message) -ForegroundColor $color
}

# ============================================================================
# Pre-flight: make sure the user filled in the placeholders
# ============================================================================
function Assert-Config {
    $bad = @()

    foreach ($k in 'BaseUrl','TokenEndpoint','AuthEndpoint','ClientId','ClientSecret') {
        if ([string]::IsNullOrWhiteSpace($Config[$k]) -or $Config[$k] -match '<.*>') {
            $bad += "  - `$Config.$k still contains a placeholder."
        }
    }

    if ($Config.RedirectUri -ne 'https://localhost') {
        $bad += "  - `$Config.RedirectUri must be exactly 'https://localhost' (no port, no trailing slash)."
    }

    if ($WebhookPayload.url -match '<.*>' -or [string]::IsNullOrWhiteSpace($WebhookPayload.url)) {
        $bad += "  - `$WebhookPayload.url still contains a placeholder."
    }

    foreach ($h in $WebhookPayload.headers) {
        if ($h.value -match '<.*>') {
            $bad += "  - A `$WebhookPayload.headers entry still contains a placeholder ('$($h.value)')."
        }
    }

    # Cross-check: same host across the three Ninja URLs
    try {
        $h1 = ([uri]$Config.BaseUrl).Host
        $h2 = ([uri]$Config.TokenEndpoint).Host
        $h3 = ([uri]$Config.AuthEndpoint).Host
        if ($h1 -and $h2 -and $h3 -and -not ($h1 -eq $h2 -and $h2 -eq $h3)) {
            $bad += "  - BaseUrl, TokenEndpoint and AuthEndpoint must all use the same host (got '$h1', '$h2', '$h3')."
        }
    } catch {
        # URL parse failure will be caught by the placeholder check above
    }

    if ($bad.Count -gt 0) {
        Write-Host ""
        Write-Status "Configuration is incomplete:" -Level Err
        $bad | ForEach-Object { Write-Host $_ -ForegroundColor Red }
        Write-Host ""
        Write-Host "Open Set-NinjaWebhook.ps1, fix the items above, save, and re-run." -ForegroundColor Yellow
        Write-Host ""
        exit 1
    }
}

# ============================================================================
# STEP A - Get an authorization code from Ninja (browser round-trip)
# ============================================================================
function Get-NinjaAuthCode {
    # Random state so we can verify the redirect we get back is ours
    $state = [guid]::NewGuid().ToString('N')

    # Build the consent URL. All values URL-encoded.
    $authUrl = '{0}?response_type=code&client_id={1}&redirect_uri={2}&scope={3}&state={4}' -f `
        $Config.AuthEndpoint, `
        [uri]::EscapeDataString($Config.ClientId), `
        [uri]::EscapeDataString($Config.RedirectUri), `
        [uri]::EscapeDataString($Config.Scope), `
        $state

    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host " NinjaOne Webhook Setup - Authorization Code Flow" -ForegroundColor Cyan
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Status "Opening NinjaOne in your default browser..." -Level Step
    Write-Host ""
    Write-Host "What's about to happen:" -ForegroundColor Yellow
    Write-Host "  1) Log in to NinjaOne (if you aren't already)."
    Write-Host "  2) Click 'Allow' on the consent screen."
    Write-Host "  3) Your browser will try to load https://localhost/?code=..."
    Write-Host "     and show a 'site can't be reached' or certificate warning."
    Write-Host "     ** That is EXPECTED. Nothing is broken. **"
    Write-Host "  4) Copy the FULL URL from your browser's address bar"
    Write-Host "     (it will start with https://localhost/?code=) and paste"
    Write-Host "     it back here when prompted."
    Write-Host ""

    Start-Process $authUrl

    Write-Host ""
    $returnedUrl = Read-Host "Paste the full https://localhost/?code=... URL here"

    if ([string]::IsNullOrWhiteSpace($returnedUrl)) {
        throw "No URL was pasted. Aborting."
    }

    # Make sure System.Web is available for query parsing on PS 5.1 and 7
    try { Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue } catch { }

    try {
        $uri   = [uri]$returnedUrl
        $query = [System.Web.HttpUtility]::ParseQueryString($uri.Query)
    } catch {
        throw "Couldn't parse the URL you pasted. Make sure you copied the entire address from the browser. ($($_.Exception.Message))"
    }

    $code          = $query['code']
    $returnedState = $query['state']
    $errParam      = $query['error']

    if ($errParam) {
        $desc = $query['error_description']
        throw "Ninja returned an OAuth error '$errParam'$(if ($desc) { ": $desc" })."
    }

    if (-not $code) {
        throw "No 'code' parameter found in the URL. Did you paste the right one? It should look like https://localhost/?code=...&state=..."
    }

    if ($returnedState -ne $state) {
        Write-Status "State value didn't match what we sent ('$returnedState' vs '$state'). Continuing, but be aware this could indicate a stale or hijacked redirect." -Level Warn
    }

    Write-Status "Got authorization code." -Level Ok
    return $code
}

# ============================================================================
# STEP B - Exchange the code for an access token
# ============================================================================
function Get-NinjaAccessToken {
    param([Parameter(Mandatory)] [string] $AuthCode)

    Write-Status "Exchanging authorization code for an access token..." -Level Step

    $body = @{
        grant_type    = 'authorization_code'
        code          = $AuthCode
        redirect_uri  = $Config.RedirectUri
        client_id     = $Config.ClientId
        client_secret = $Config.ClientSecret
    }

    try {
        $resp = Invoke-RestMethod -Method Post `
                                  -Uri $Config.TokenEndpoint `
                                  -Body $body `
                                  -ContentType 'application/x-www-form-urlencoded'
    } catch {
        Write-Host ""
        Write-Status "Token exchange failed." -Level Err
        $detail = $_.ErrorDetails.Message
        if ($detail) {
            Write-Host "Response from Ninja:" -ForegroundColor Red
            Write-Host $detail -ForegroundColor Red
        } else {
            Write-Host $_.Exception.Message -ForegroundColor Red
        }
        Write-Host ""
        Write-Host "Common causes:" -ForegroundColor Yellow
        Write-Host "  invalid_grant      - auth code already used or expired (codes live ~60s). Re-run the script."
        Write-Host "  invalid_client     - wrong Client ID/Secret, or redirect URI on the Ninja app doesn't match exactly 'https://localhost'."
        Write-Host "  unauthorized_client - the OAuth app doesn't have 'Authorization Code' enabled under Allowed Grant Types."
        throw
    }

    if (-not $resp.access_token) {
        throw "Token endpoint responded but no access_token was in the response."
    }

    Write-Status ("Got access token (expires in {0}s)." -f $resp.expires_in) -Level Ok
    return $resp.access_token
}

# ============================================================================
# STEP C - Configure the webhook
# ============================================================================
function Set-NinjaWebhookConfig {
    param(
        [Parameter(Mandatory)] [string] $AccessToken,
        [Parameter(Mandatory)]          $Payload
    )

    $endpoint = '{0}/v2/webhook' -f $Config.BaseUrl.TrimEnd('/')
    $jsonBody = $Payload | ConvertTo-Json -Depth 10

    Write-Host ""
    Write-Status "Submitting webhook configuration:" -Level Step
    Write-Host "  PUT $endpoint"
    Write-Host ""
    Write-Host "Payload:" -ForegroundColor Yellow
    Write-Host $jsonBody
    Write-Host ""

    $headers = @{
        Authorization  = "Bearer $AccessToken"
        'Content-Type' = 'application/json'
    }

    try {
        # Invoke-WebRequest (not Invoke-RestMethod) so we can read StatusCode.
        # 204 No Content = success.
        $resp = Invoke-WebRequest -Method Put `
                                  -Uri $endpoint `
                                  -Headers $headers `
                                  -Body $jsonBody `
                                  -UseBasicParsing
    } catch {
        Write-Host ""
        Write-Status "Webhook configuration failed." -Level Err
        $detail = $_.ErrorDetails.Message
        if ($detail) {
            Write-Host "Response from Ninja:" -ForegroundColor Red
            Write-Host $detail -ForegroundColor Red
        } else {
            Write-Host $_.Exception.Message -ForegroundColor Red
        }
        Write-Host ""
        Write-Host "Common causes:" -ForegroundColor Yellow
        Write-Host "  401 Unauthorized - token invalid/expired. Re-run the script."
        Write-Host "  403 Forbidden    - your Ninja user is not a System Administrator, or the OAuth app's scope doesn't include 'monitoring'."
        Write-Host "  400 Bad Request  - check the payload (URL format, activity type names, header structure)."
        throw
    }

    if ($resp.StatusCode -eq 204) {
        Write-Host ""
        Write-Status "204 No Content - webhook configured successfully." -Level Ok
    } else {
        Write-Status ("Got HTTP {0}, expected 204. Body: {1}" -f $resp.StatusCode, $resp.Content) -Level Warn
    }
}

# ============================================================================
# MAIN
# ============================================================================
Assert-Config

$code  = Get-NinjaAuthCode
$token = Get-NinjaAccessToken -AuthCode $code
Set-NinjaWebhookConfig -AccessToken $token -Payload $WebhookPayload

Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host " Done. Verify in NinjaOne:"                                    -ForegroundColor Green
Write-Host "   Administration -> Notification Channels"                    -ForegroundColor Green
Write-Host ""
Write-Host " You'll see a webhook entry with NO NAME. That's the one."     -ForegroundColor Green
Write-Host " DO NOT rename, edit, or modify it - touching it will break"   -ForegroundColor Green
Write-Host " the live stream to your SIEM."                                -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""
