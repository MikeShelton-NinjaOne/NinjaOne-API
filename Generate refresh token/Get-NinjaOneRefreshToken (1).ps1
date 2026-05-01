<#
.SYNOPSIS
    Generates a NinjaOne refresh token using the OAuth 2.0 Authorization Code flow.

.DESCRIPTION
    Fill in the $Config block below with your NinjaOne details, then run the script.
    It opens your browser to the NinjaOne authorization page, captures the redirect
    on http://localhost/, and exchanges the code for an access token + refresh token.

    Requires PowerShell 5.1+ on Windows. Because the redirect URI has no explicit
    port, the local listener binds to port 80, which means this script MUST be run
    from an elevated (Administrator) PowerShell prompt.

.NOTES
    Run as Administrator. The redirect URI registered in NinjaOne must exactly
    match the one in $Config (default: http://localhost/).
#>

# =============================================================================
# CONFIGURATION — fill in these values before running
# =============================================================================
$Config = @{
    # Your NinjaOne login URL (no trailing slash). Examples:
    #   https://app.ninjarmm.com   (North America)
    #   https://eu.ninjarmm.com    (Europe)
    #   https://ca.ninjarmm.com    (Canada)
    #   https://oc.ninjarmm.com    (Australia / Oceania)
    BaseUrl       = 'https://<your Login URL>'

    # OAuth 2.0 token endpoint. Should be your BaseUrl + '/ws/oauth/token'.
    TokenEndpoint = 'https://<your Login URL>/ws/oauth/token'

    # Client ID and Secret from Administration > Apps > API > Client App IDs.
    ClientId      = '<Your Client ID>'
    ClientSecret  = '<Your Client Secret>'

    # Redirect URI. Must EXACTLY match what is registered on the NinjaOne client app.
    RedirectUri   = 'http://localhost/'

    # Space-delimited scopes. 'offline_access' is required to get a refresh token.
    Scope         = 'monitoring management control offline_access'

    # Optional: if set, the token bundle is saved DPAPI-encrypted to this path.
    # Leave $null to skip saving.
    OutputFile    = $null
}
# =============================================================================

# --- Sanity checks -----------------------------------------------------------

if ($Config.BaseUrl -like '*<your*' -or
    $Config.ClientId -like '<Your*' -or
    $Config.ClientSecret -like '<Your*') {
    Write-Error "Please fill in the `$Config block at the top of the script before running."
    return
}

if ($Config.Scope -notmatch '\boffline_access\b') {
    Write-Warning "Scope does not include 'offline_access' — NinjaOne will not return a refresh token."
}

# Helper: generate a CSRF-protection state value
function New-RandomState {
    $bytes = New-Object byte[] 24
    [System.Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($bytes)
    [Convert]::ToBase64String($bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
}

$state = New-RandomState

# --- Build the authorization URL --------------------------------------------

$authParams = @{
    response_type = 'code'
    client_id     = $Config.ClientId
    redirect_uri  = $Config.RedirectUri
    scope         = $Config.Scope
    state         = $state
}

$queryString = ($authParams.GetEnumerator() | ForEach-Object {
    "{0}={1}" -f [uri]::EscapeDataString($_.Key), [uri]::EscapeDataString($_.Value)
}) -join '&'

$authUrl = "$($Config.BaseUrl)/ws/oauth/authorize?$queryString"

# --- Start a local HTTP listener BEFORE opening the browser -----------------

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add($Config.RedirectUri)

try {
    $listener.Start()
}
catch {
    Write-Error @"
Could not start the local listener on $($Config.RedirectUri).
Most common cause: this script must be run as Administrator because port 80
is a privileged port on Windows. Right-click PowerShell and choose
'Run as Administrator', then re-run.

Other causes: another process (IIS, Skype, a dev server) is already bound
to port 80. Stop it or change the RedirectUri to include a free port
(e.g. 'http://localhost:8765/') AND update NinjaOne to match.

Underlying error: $($_.Exception.Message)
"@
    return
}

Write-Host ""
Write-Host "Opening browser for NinjaOne login..." -ForegroundColor Green
Start-Process $authUrl
Write-Host "Waiting for authorization redirect on $($Config.RedirectUri) ..." -ForegroundColor Green

# --- Wait for the redirect ---------------------------------------------------

try {
    $context  = $listener.GetContext()
    $request  = $context.Request
    $response = $context.Response

    $returnedCode  = $request.QueryString['code']
    $returnedState = $request.QueryString['state']
    $errorParam    = $request.QueryString['error']
    $errorDesc     = $request.QueryString['error_description']

    if ($errorParam) {
        $html = "<html><body style='font-family:sans-serif'><h2>Authorization failed</h2><p><b>$errorParam</b>: $errorDesc</p><p>You can close this window.</p></body></html>"
    }
    elseif (-not $returnedCode) {
        $html = "<html><body style='font-family:sans-serif'><h2>No authorization code returned</h2><p>You can close this window.</p></body></html>"
    }
    else {
        $html = "<html><body style='font-family:sans-serif'><h2>Authorization successful</h2><p>You can close this window and return to PowerShell.</p></body></html>"
    }

    $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
    $response.ContentType     = 'text/html; charset=utf-8'
    $response.ContentLength64 = $buffer.Length
    $response.OutputStream.Write($buffer, 0, $buffer.Length)
    $response.OutputStream.Close()
}
finally {
    $listener.Stop()
    $listener.Close()
}

if ($errorParam) {
    Write-Error "NinjaOne returned an OAuth error: $errorParam - $errorDesc"
    return
}
if (-not $returnedCode) {
    Write-Error "No authorization code was returned. Aborting."
    return
}
if ($returnedState -ne $state) {
    Write-Error "State parameter mismatch. Possible CSRF attempt. Aborting."
    return
}

Write-Host "Authorization code received. Exchanging for tokens..." -ForegroundColor Green

# --- Exchange the code for tokens -------------------------------------------

$body = @{
    grant_type    = 'authorization_code'
    code          = $returnedCode
    redirect_uri  = $Config.RedirectUri
    client_id     = $Config.ClientId
    client_secret = $Config.ClientSecret
}

try {
    $tokenResponse = Invoke-RestMethod -Method Post `
        -Uri $Config.TokenEndpoint `
        -Body $body `
        -ContentType 'application/x-www-form-urlencoded' `
        -ErrorAction Stop
}
catch {
    $errMsg = $_.Exception.Message
    if ($_.Exception.Response) {
        try {
            $stream = $_.Exception.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($stream)
            $errMsg = "$errMsg`nResponse: $($reader.ReadToEnd())"
        } catch { }
    }
    Write-Error "Token exchange failed: $errMsg"
    return
}

if (-not $tokenResponse.refresh_token) {
    Write-Warning "Token exchange succeeded but no refresh_token was returned."
    Write-Warning "Make sure 'offline_access' is in your scope AND that 'Refresh Token' is enabled as an allowed grant type on the NinjaOne client app."
}

Write-Host ""
Write-Host "=== Tokens ===" -ForegroundColor Cyan
Write-Host "Access Token  : $($tokenResponse.access_token)"
Write-Host "Refresh Token : $($tokenResponse.refresh_token)"
Write-Host "Expires In    : $($tokenResponse.expires_in) seconds"
Write-Host "Token Type    : $($tokenResponse.token_type)"
Write-Host "Scope         : $($tokenResponse.scope)"
Write-Host ""

# --- Optional: save encrypted bundle ----------------------------------------

if ($Config.OutputFile) {
    $payload = [PSCustomObject]@{
        base_url      = $Config.BaseUrl
        client_id     = $Config.ClientId
        scope         = $tokenResponse.scope
        access_token  = $tokenResponse.access_token
        refresh_token = $tokenResponse.refresh_token
        expires_in    = $tokenResponse.expires_in
        retrieved_at  = (Get-Date).ToString('o')
    }

    $json   = $payload | ConvertTo-Json
    $secure = ConvertTo-SecureString -String $json -AsPlainText -Force
    $secure | ConvertFrom-SecureString | Set-Content -Path $Config.OutputFile -Encoding ASCII

    Write-Host "Encrypted token bundle written to: $($Config.OutputFile)" -ForegroundColor Green
    Write-Host "(Decryptable only by the same Windows user account on the same machine.)" -ForegroundColor DarkGray
}

# Return the token object for pipeline use
$tokenResponse
