#Requires -Version 5.1
<#
.SYNOPSIS
    Assigns a device to a NinjaOne ticket using Authorization Code OAuth flow.

.DESCRIPTION
    Uses the NinjaOne Public API v2 with Authorization Code flow. This is
    required for ticket updates -- NinjaOne enforces user context on this
    endpoint and rejects Client Credentials (machine-to-machine) tokens.

    On first run:
      - Opens a browser window to the NinjaOne login page
      - You log in as a technician and approve the request
      - NinjaOne redirects to localhost with an authorization code
      - Script exchanges the code for an access + refresh token
      - Tokens are cached to a local file for future runs

    On subsequent runs:
      - Uses the cached refresh token silently -- no browser needed
      - If the refresh token has expired, re-opens the browser to log in again

    Token cache file is saved alongside the script as NinjaTokenCache.json.
    Delete that file to force a fresh login.

    REQUIRES PSAuthClient module (install once):
      Install-Module PSAuthClient -Scope CurrentUser -Confirm:$false

    API APP SETUP (one-time -- separate app from your machine-to-machine app):
    Administration > Apps > API > Client App IDs > Add
      Name          : Any name, e.g. TicketDeviceAssign
      Platform      : Web             <-- issues a client secret
      Allowed Scopes: monitoring AND management
      Grant Types   : Authorization Code AND Refresh Token
      Redirect URI  : https://localhost/   <-- exactly this, with trailing slash

    REGIONAL URLS:
    US: https://app.ninjarmm.com  |  EU: https://eu.ninjarmm.com
    OC: https://oc.ninjarmm.com   |  CA: https://ca.ninjarmm.com
#>

# ==============================================================================
#  CONFIGURATION -- Fill in ALL values before running
# ==============================================================================

$BaseUrl       = 'https://<your Login URL>'       # No trailing slash
$ClientId      = '<Your Web App Client ID>'       # From the Web platform app
$ClientSecret  = '<Your Web App Client Secret>'   # From the Web platform app

# The ticket you want to assign a device to
$TicketId      = 0    # <-- Replace with your ticket ID

# The device you want to assign to the ticket
$DeviceId      = 0    # <-- Replace with your device ID

# Set to $true to print the raw ticket JSON and exit without making changes
# Use this to verify field names before committing to an update
$DiagnosticMode = $false

# ==============================================================================
#  DO NOT EDIT BELOW THIS LINE
# ==============================================================================

try { [Net.ServicePointManager]::SecurityProtocol =
    [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch {}
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
$ErrorActionPreference = 'Stop'

$TokenEndpoint  = "$BaseUrl/ws/oauth/token"
$AuthEndpoint   = "$BaseUrl/ws/oauth/authorize"
$Scopes         = 'monitoring management offline_access'
$ScriptDir      = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.Path }
$CacheFile      = Join-Path $ScriptDir 'NinjaTokenCache.json'

# -- Config validation ---------------------------------------------------------
$ConfigErrors = @()
if ($BaseUrl      -like '*<*') { $ConfigErrors += 'Fill in $BaseUrl' }
if ($ClientId     -like '*<*') { $ConfigErrors += 'Fill in $ClientId' }
if ($ClientSecret -like '*<*') { $ConfigErrors += 'Fill in $ClientSecret' }
if ($TicketId  -eq 0)       { $ConfigErrors += 'Fill in $TicketId' }
if ($DeviceId  -eq 0)       { $ConfigErrors += 'Fill in $DeviceId' }
if ($ConfigErrors.Count -gt 0) {
    Write-Host ''
    Write-Host '  [!] Configuration errors:' -ForegroundColor Red
    $ConfigErrors | ForEach-Object { Write-Host "      - $_" -ForegroundColor Red }
    exit 1
}

# -- Helper -------------------------------------------------------------------
function Get-Prop {
    param([object]$Obj, [string]$Name, [object]$Default = $null)
    if ($null -eq $Obj) { return $Default }
    $p = $Obj.PSObject.Properties[$Name]
    if ($null -eq $p -or $null -eq $p.Value) { return $Default }
    return $p.Value
}

Write-Host ''
Write-Host '  ================================================================' -ForegroundColor Cyan
Write-Host '  NinjaOne -- Assign Device to Ticket'                             -ForegroundColor Cyan
Write-Host "  Ticket ID: $TicketId  |  Device ID: $DeviceId"                  -ForegroundColor Cyan
Write-Host '  ================================================================' -ForegroundColor Cyan
Write-Host ''

# =============================================================================
#  STEP 1: Get access token (from cache or via browser login)
# =============================================================================
Write-Host '  [1/4] Authenticating...' -ForegroundColor Cyan

$AccessToken = $null

# -- Try cached refresh token first -------------------------------------------
if (Test-Path $CacheFile) {
    try {
        $Cache = Get-Content $CacheFile -Raw | ConvertFrom-Json
        $RefreshToken = Get-Prop $Cache 'refresh_token'
        if ($RefreshToken) {
            Write-Host '       Found cached token -- attempting silent refresh...' -ForegroundColor Gray
            $RefreshResp = Invoke-RestMethod -Uri $TokenEndpoint -Method POST `
                -ContentType 'application/x-www-form-urlencoded' `
                -Body ("grant_type=refresh_token" +
                       "&client_id=$([Uri]::EscapeDataString($ClientId))" +
                       "&client_secret=$([Uri]::EscapeDataString($ClientSecret))" +
                       "&refresh_token=$([Uri]::EscapeDataString($RefreshToken))") `
                -ErrorAction Stop
            $AccessToken = $RefreshResp.access_token
            # Save updated tokens
            $RefreshResp | ConvertTo-Json | Set-Content $CacheFile -Encoding UTF8
            Write-Host '  [OK] Authenticated via cached refresh token.' -ForegroundColor Green
        }
    } catch {
        Write-Host '       Cached token expired or invalid -- will open browser.' -ForegroundColor Yellow
        $AccessToken = $null
    }
}

# -- Fall back to browser Authorization Code flow via PSAuthClient -----------
# PSAuthClient is NinjaOne's own recommended module for this flow.
# It handles the https://localhost/ callback correctly without admin rights.
if (-not $AccessToken) {
    if (-not (Get-Module -ListAvailable -Name PSAuthClient)) {
        Write-Host '       PSAuthClient module not found -- installing automatically...' -ForegroundColor Yellow
        try {
            Install-Module PSAuthClient -Scope CurrentUser -Force -AllowClobber `
                -Repository PSGallery -ErrorAction Stop
            Write-Host '       PSAuthClient installed successfully.' -ForegroundColor Green
        } catch {
            Write-Host "  [!] Failed to install PSAuthClient: $_" -ForegroundColor Red
            Write-Host "      Try manually: Install-Module PSAuthClient -Scope CurrentUser -Force" -ForegroundColor Yellow
            exit 1
        }
    }
    Import-Module PSAuthClient -ErrorAction Stop

    $CallbackUri = 'https://localhost/'

    Write-Host ''
    Write-Host '  [..] Opening browser for NinjaOne login...' -ForegroundColor Yellow
    Write-Host '       Log in and approve the request in your browser.' -ForegroundColor Yellow
    Write-Host '       This window will continue automatically after login.' -ForegroundColor Yellow
    Write-Host ''

    try {
        # Step 1: Authorization endpoint -- returns object with code, redirect_uri,
        # client_id etc. which is then splatted directly into the token endpoint.
        # This is the confirmed pattern from PSAuthClient docs and NinjaOne's own example.
        $Auth = Invoke-OAuth2AuthorizationEndpoint `
            -uri        $AuthEndpoint `
            -client_id  $ClientId `
            -redirect_uri $CallbackUri `
            -scope      $Scopes `
            -usePkce:$false

        if (-not $Auth.code) {
            Write-Host "  [!] No authorization code received." -ForegroundColor Red
            exit 1
        }

        # Step 2: Add client_secret to the auth object, then splat into token endpoint
        $Auth.Add('client_secret', $ClientSecret)
        $TokenResp   = Invoke-OAuth2TokenEndpoint -uri $TokenEndpoint @Auth
        $AccessToken = $TokenResp.access_token
        $TokenResp | ConvertTo-Json | Set-Content $CacheFile -Encoding UTF8
        Write-Host '  [OK] Login successful. Tokens cached for future runs.' -ForegroundColor Green
    } catch {
        Write-Host "  [!] Authorization flow failed: $_" -ForegroundColor Red
        exit 1
    }
}

$Headers = @{ Authorization = "Bearer $AccessToken"; Accept = 'application/json' }

# =============================================================================
#  STEP 2: Look up ticket
# =============================================================================
Write-Host ''; Write-Host "  [2/4] Looking up ticket $TicketId..." -ForegroundColor Cyan
$Ticket = $null
try {
    $Ticket = Invoke-RestMethod -Uri "$BaseUrl/v2/ticketing/ticket/$TicketId" `
        -Method GET -Headers $Headers -ErrorAction Stop
} catch {
    $sc = $null; try { $sc = [int]$_.Exception.Response.StatusCode } catch {}
    if ($sc -eq 404) {
        Write-Host "  [!] Ticket $TicketId not found." -ForegroundColor Red
    } else {
        Write-Host "  [!] Failed to retrieve ticket (HTTP $sc): $_" -ForegroundColor Red
    }
    exit 1
}

$TicketSubject      = [string](Get-Prop $Ticket 'subject'     -Default '(no subject)')
# status is an object { name, displayName, statusId } -- extract the name string
$TicketStatusObj    = Get-Prop $Ticket 'status'
$TicketStatus       = if ($TicketStatusObj) { [string](Get-Prop $TicketStatusObj 'name' -Default 'NEW') } else { 'NEW' }
$TicketOrg          = Get-Prop $Ticket 'clientId'     -Default 'Unknown'
$TicketRequesterUid = [string](Get-Prop $Ticket 'requesterUid')
$TicketFormId       = Get-Prop $Ticket 'ticketFormId'
$TicketVersion      = Get-Prop $Ticket 'version'
$CurrentDevice      = Get-Prop $Ticket 'nodeId'

Write-Host "  [OK] Found ticket:" -ForegroundColor Green
Write-Host "       Subject : $TicketSubject"
Write-Host "       Status  : $TicketStatus"
Write-Host "       Org ID  : $TicketOrg"
if ($CurrentDevice) {
    Write-Host "       Current node   : $CurrentDevice (will be replaced)" -ForegroundColor Yellow
} else {
    Write-Host "       Current node   : (none)"
}

if ($DiagnosticMode) {
    Write-Host ''
    Write-Host '  [DIAGNOSTIC] Raw ticket JSON:' -ForegroundColor Magenta
    $Ticket | ConvertTo-Json -Depth 3 | Write-Host
    Write-Host ''
    Write-Host '  [DIAGNOSTIC] Exiting without making changes.' -ForegroundColor Magenta
    Write-Host '               Set $DiagnosticMode = $false to run normally.' -ForegroundColor Magenta
    exit 0
}

# =============================================================================
#  STEP 3: Look up device
# =============================================================================
Write-Host ''; Write-Host "  [3/4] Looking up device $DeviceId..." -ForegroundColor Cyan
$Device = $null
try {
    $Device = Invoke-RestMethod -Uri "$BaseUrl/v2/device/$DeviceId" `
        -Method GET -Headers $Headers -ErrorAction Stop
} catch {
    $sc = $null; try { $sc = [int]$_.Exception.Response.StatusCode } catch {}
    if ($sc -eq 404) {
        Write-Host "  [!] Device $DeviceId not found." -ForegroundColor Red
    } else {
        Write-Host "  [!] Failed to retrieve device (HTTP $sc): $_" -ForegroundColor Red
    }
    exit 1
}

$DeviceName  = Get-Prop $Device 'systemName' -Default (Get-Prop $Device 'dnsName' -Default "Device $DeviceId")
$DeviceClass = Get-Prop $Device 'nodeClass'  -Default 'UNKNOWN'
$DeviceOrg   = Get-Prop $Device 'organizationId'

Write-Host "  [OK] Found device:" -ForegroundColor Green
Write-Host "       Name       : $DeviceName"
Write-Host "       Node class : $DeviceClass"
Write-Host "       Org ID     : $DeviceOrg"

# =============================================================================
#  STEP 4: Assign device to ticket
# =============================================================================
Write-Host ''; Write-Host "  [4/4] Assigning device $DeviceId ($DeviceName) to ticket $TicketId..." -ForegroundColor Cyan

$PutHeaders = $Headers.Clone()
$PutHeaders['Content-Type'] = 'application/json'
# Build PUT body using JavaScriptSerializer to safely handle special chars
# Include all fields the API marks as non-nullable so we don't get 400 errors
Add-Type -AssemblyName System.Web.Extensions -ErrorAction SilentlyContinue
$Jss = New-Object System.Web.Script.Serialization.JavaScriptSerializer

# Cast all values to plain strings/ints to avoid PSMethod circular reference
# in JavaScriptSerializer
# Confirmed from NinjaOne OpenAPI spec (homotechsual/NinjaOne.oa3.json):
# - Body must have a 'ticket' wrapper object
# - Device field is 'nodeId' (int) not 'deviceId'
# - Subject field is 'summary' not 'subject'
# - Required non-nullable fields: summary, status, requesterUid
# Body is a flat object -- no wrapper. Confirmed from debug output that
# wrapping in 'ticket' causes the API to fail to find 'subject'.
$PutDict = New-Object 'System.Collections.Generic.Dictionary[string,object]'
$PutDict.Add('subject',      $TicketSubject)
$PutDict.Add('status',       $TicketStatus)
$PutDict.Add('requesterUid', $TicketRequesterUid)
$PutDict.Add('nodeId',       [int]$DeviceId)
if ($TicketFormId) { $PutDict.Add('ticketFormId', [int]$TicketFormId) }
$TicketClientId = Get-Prop $Ticket 'clientId'
if ($TicketClientId -and [int]$TicketClientId -ne 0) {
    $PutDict.Add('clientId', [int]$TicketClientId)
} else {
    Write-Host "  [!] Ticket $TicketId has no organization assigned." -ForegroundColor Red
    Write-Host "      A device cannot be assigned to a ticket that is not linked to an organization." -ForegroundColor Yellow
    Write-Host "      Open the ticket in NinjaOne, assign it to an organization, then run this script again." -ForegroundColor Yellow
    exit 1
}
if ($TicketVersion) { $PutDict.Add('version', [int]$TicketVersion) }

$Body = $Jss.Serialize($PutDict)

try {
    $Result        = Invoke-RestMethod -Uri "$BaseUrl/v2/ticketing/ticket/$TicketId" `
        -Method PUT -Headers $PutHeaders -Body $Body -ErrorAction Stop
    # nodeId confirmed in response

    Write-Host "  [OK] Device assigned successfully." -ForegroundColor Green
    Write-Host "       Ticket ID      : $TicketId"
    Write-Host "       Ticket subject : $TicketSubject"
    Write-Host "       Node ID        : $DeviceId"
    Write-Host "       Device name    : $DeviceName"
} catch {
    $sc     = $null; try { $sc = [int]$_.Exception.Response.StatusCode }     catch {}
    $detail = '';    try { $detail = $_.ErrorDetails.Message }                catch {}
    Write-Host "  [!] Failed to assign device (HTTP $sc): $_" -ForegroundColor Red
    if ($detail) { Write-Host "      API detail: $detail" -ForegroundColor Red }
    exit 1
}

Write-Host ''
Write-Host '  ================================================================' -ForegroundColor Green
Write-Host '  [OK] COMPLETE'                                                    -ForegroundColor Green
Write-Host '  ================================================================' -ForegroundColor Green
Write-Host ''
