#Requires -Version 5.1
<#
.SYNOPSIS
    Assigns a device owner in NinjaOne by hostname, first name, and last name.

.DESCRIPTION
    This script authenticates to the NinjaOne API using OAuth2 (Authorization Code flow),
    finds a device by its hostname, finds a contact/end-user by first and last name,
    then assigns that contact as the device owner.

.NOTES
    Prerequisites:
      1. A NinjaOne API application must be created in your portal.
         Go to: Administration > Apps > API > Client App IDs > Add
         Set the redirect URI to: https://localhost
         Set the platform to: "Web" or "Application"

      2. The PSAuthClient PowerShell module is required for the OAuth2 flow.
         Install it by running (as Administrator, one time only):
           Install-Module PSAuthClient -Confirm:$false

      3. Run this script as Administrator (required for the local browser redirect).

.EXAMPLE
    .\Set-NinjaDeviceOwner.ps1 `
        -Hostname "DESKTOP-ABC123" `
        -FirstName "Jane" `
        -LastName "Smith"
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, HelpMessage = "The exact hostname of the device to update.")]
    [string]$Hostname,

    [Parameter(Mandatory = $true, HelpMessage = "First name of the user to assign as owner.")]
    [string]$FirstName,

    [Parameter(Mandatory = $true, HelpMessage = "Last name of the user to assign as owner.")]
    [string]$LastName
)

# ==============================================================================
#  CONFIGURATION — Fill in your NinjaOne details here before running the script
# ==============================================================================

# Your NinjaOne login URL  (e.g. https://app.ninjarmm.com  or  https://eu.ninjarmm.com)
$BaseUrl        = 'https://<your Login URL>'

# OAuth2 token endpoint — same base URL, just add /ws/oauth/token
$TokenEndpoint  = 'https://<your Login URL>/ws/oauth/token'

# Client ID from Administration > Apps > API > Client App IDs
$ClientId       = '<Your Client ID>'

# Client Secret from the same page
$ClientSecret   = '<Your Client Secret>'

# Redirect URL — must match EXACTLY what you entered in the NinjaOne portal
$RedirectUri    = 'https://localhost'

# ==============================================================================
#  DO NOT EDIT BELOW THIS LINE
# ==============================================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Helper: Check for PSAuthClient module ────────────────────────────────────
if (-not (Get-Module -ListAvailable -Name PSAuthClient)) {
    Write-Host ""
    Write-Host "  [!] The PSAuthClient module is not installed." -ForegroundColor Yellow
    Write-Host "      Run the following command as Administrator, then re-run this script:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "      Install-Module PSAuthClient -Confirm:`$false" -ForegroundColor Cyan
    Write-Host ""
    exit 1
}

Import-Module PSAuthClient -ErrorAction Stop

# ── Helper: Build API headers from an access token ───────────────────────────
function Get-AuthHeaders {
    param ([string]$AccessToken)
    return @{
        Authorization  = "Bearer $AccessToken"
        'Content-Type' = 'application/json'
        Accept         = 'application/json'
    }
}

# ── Step 1: Authenticate via OAuth2 Authorization Code flow ──────────────────
Write-Host ""
Write-Host "  [1/4] Authenticating with NinjaOne..." -ForegroundColor Cyan
Write-Host "        A browser window will open for you to log in." -ForegroundColor Gray
Write-Host ""

$AuthorizeEndpoint = "$BaseUrl/ws/oauth/authorize"

$AuthParams = @{
    Uri              = $AuthorizeEndpoint
    Redirect_uri     = $RedirectUri
    Client_id        = $ClientId
    Scope            = 'monitoring management offline_access'
    UsePkce          = $false
    CustomParameters = @{ client_secret = $ClientSecret }
}

try {
    $Auth     = Invoke-OAuth2AuthorizationEndpoint @AuthParams
    $AuthCode = $Auth.code
} catch {
    Write-Host "  [!] Authentication failed. Check your BaseUrl, ClientId, ClientSecret," -ForegroundColor Red
    Write-Host "      and that your redirect URI in NinjaOne is set to: $RedirectUri" -ForegroundColor Red
    Write-Host "      Error: $_" -ForegroundColor Red
    exit 1
}

$TokenParams = @{
    Uri              = $TokenEndpoint
    Redirect_uri     = $RedirectUri
    Client_id        = $ClientId
    Code             = $AuthCode
    CustomParameters = @{ client_secret = $ClientSecret }
}

try {
    $TokenResponse = Invoke-OAuth2TokenEndpoint @TokenParams
    $AccessToken   = $TokenResponse.access_token
} catch {
    Write-Host "  [!] Failed to exchange authorization code for access token." -ForegroundColor Red
    Write-Host "      Error: $_" -ForegroundColor Red
    exit 1
}

Write-Host "  [✓] Successfully authenticated." -ForegroundColor Green

$Headers = Get-AuthHeaders -AccessToken $AccessToken

# ── Step 2: Find the device by hostname ──────────────────────────────────────
Write-Host ""
Write-Host "  [2/4] Looking up device with hostname: '$Hostname'..." -ForegroundColor Cyan

try {
    $DevicesResponse = Invoke-RestMethod `
        -Uri     "$BaseUrl/v2/devices?pageSize=1000" `
        -Method  GET `
        -Headers $Headers
} catch {
    Write-Host "  [!] Failed to retrieve devices from NinjaOne." -ForegroundColor Red
    Write-Host "      Error: $_" -ForegroundColor Red
    exit 1
}

# The API returns an array directly
$MatchingDevices = $DevicesResponse | Where-Object {
    $_.systemName -like $Hostname -or $_.dnsName -like $Hostname
}

if (-not $MatchingDevices) {
    Write-Host "  [!] No device found with hostname '$Hostname'." -ForegroundColor Red
    Write-Host "      Tip: Hostname matching is case-insensitive. Check spelling and try again." -ForegroundColor Yellow
    exit 1
}

if (@($MatchingDevices).Count -gt 1) {
    Write-Host "  [!] Multiple devices matched hostname '$Hostname':" -ForegroundColor Yellow
    $MatchingDevices | ForEach-Object {
        Write-Host "       ID: $($_.id)  |  Name: $($_.systemName)  |  DNS: $($_.dnsName)" -ForegroundColor Gray
    }
    Write-Host "      Please use a more specific hostname and try again." -ForegroundColor Yellow
    exit 1
}

$Device   = $MatchingDevices
$DeviceId = $Device.id

Write-Host "  [✓] Device found — ID: $DeviceId  |  Name: $($Device.systemName)" -ForegroundColor Green

# ── Step 3: Find the contact/end-user by first and last name ─────────────────
Write-Host ""
Write-Host "  [3/4] Looking up contact: '$FirstName $LastName'..." -ForegroundColor Cyan

# Search end-users (contacts) — try the contacts endpoint first
try {
    $ContactsResponse = Invoke-RestMethod `
        -Uri     "$BaseUrl/v2/contacts?pageSize=1000" `
        -Method  GET `
        -Headers $Headers
} catch {
    # Fall back to the end-user endpoint used in older API versions
    try {
        $ContactsResponse = Invoke-RestMethod `
            -Uri     "$BaseUrl/v2/organization/end-users?pageSize=1000" `
            -Method  GET `
            -Headers $Headers
    } catch {
        Write-Host "  [!] Failed to retrieve contacts/end-users from NinjaOne." -ForegroundColor Red
        Write-Host "      Error: $_" -ForegroundColor Red
        exit 1
    }
}

$MatchingContacts = $ContactsResponse | Where-Object {
    $_.firstName -like $FirstName -and $_.lastName -like $LastName
}

if (-not $MatchingContacts) {
    Write-Host "  [!] No contact found matching '$FirstName $LastName'." -ForegroundColor Red
    Write-Host "      Tips:" -ForegroundColor Yellow
    Write-Host "        • Check spelling of the first and last name." -ForegroundColor Yellow
    Write-Host "        • The user must exist as a Contact or End User in NinjaOne." -ForegroundColor Yellow
    Write-Host "        • Add them in NinjaOne under the relevant Organization > End Users." -ForegroundColor Yellow
    exit 1
}

if (@($MatchingContacts).Count -gt 1) {
    Write-Host "  [!] Multiple contacts matched '$FirstName $LastName':" -ForegroundColor Yellow
    $MatchingContacts | ForEach-Object {
        Write-Host "       ID: $($_.id)  |  $($_.firstName) $($_.lastName)  |  Email: $($_.email)" -ForegroundColor Gray
    }
    Write-Host "      Please ensure first and last name uniquely identify the user." -ForegroundColor Yellow
    exit 1
}

$Contact   = $MatchingContacts
$ContactId = $Contact.id

Write-Host "  [✓] Contact found — ID: $ContactId  |  Name: $($Contact.firstName) $($Contact.lastName)" -ForegroundColor Green

# ── Step 4: Assign the contact as the device owner ───────────────────────────
Write-Host ""
Write-Host "  [4/4] Assigning '$($Contact.firstName) $($Contact.lastName)' as owner of '$($Device.systemName)'..." -ForegroundColor Cyan

$Body = @{ userId = $ContactId } | ConvertTo-Json

try {
    Invoke-RestMethod `
        -Uri     "$BaseUrl/v2/device/$DeviceId/owner" `
        -Method  POST `
        -Headers $Headers `
        -Body    $Body | Out-Null
} catch {
    $StatusCode = $_.Exception.Response.StatusCode.Value__
    Write-Host "  [!] Failed to assign device owner (HTTP $StatusCode)." -ForegroundColor Red
    Write-Host "      • Make sure your API application has the 'management' scope." -ForegroundColor Yellow
    Write-Host "      • Confirm the contact belongs to the same organization as the device." -ForegroundColor Yellow
    Write-Host "      Error: $_" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "  ================================================" -ForegroundColor Green
Write-Host "  [✓] SUCCESS" -ForegroundColor Green
Write-Host "      Device : $($Device.systemName)  (ID: $DeviceId)" -ForegroundColor Green
Write-Host "      Owner  : $($Contact.firstName) $($Contact.lastName)  (ID: $ContactId)" -ForegroundColor Green
Write-Host "  ================================================" -ForegroundColor Green
Write-Host ""
