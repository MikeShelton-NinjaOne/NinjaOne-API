# ==============================================================================
#  NinjaOne — Device Detail Export with Expanded Fields
#  (Organization, Location, Warranty)
# ==============================================================================
#
#  WHAT THIS SCRIPT DOES:
#
#  Pulls detailed information for every device in NinjaOne, including the
#  expanded Organization, Location, and Warranty data that is not returned
#  by default, and exports everything to a CSV file.
#
#  The first time you run this script it will open your browser so you can
#  log in to NinjaOne. After that it saves a refresh token so you never
#  have to log in again — it renews its own access automatically on every run.
#
#  BEFORE YOU RUN THIS SCRIPT THE FIRST TIME:
#
#  1. Log in to NinjaOne and go to:
#        Administration > Apps > API > Client App
#     Create or open an app and make sure the Redirect URI is set to exactly:
#        https://localhost
#     Also make sure Refresh Token (offline_access) is enabled on the app.
#
#  2. Fill in all values in the CONFIGURATION section below.
#
# ==============================================================================
#  CONFIGURATION — only edit values in this section
# ==============================================================================

$BaseUrl         = 'https://<your login URL>'               # e.g. https://app.ninjarmm.com
$TokenEndpoint   = 'https://<your login URL>/ws/oauth/token'
$ClientId        = '<Your Client ID>'
$ClientSecret    = '<Your Client Secret>'

# Where to save the exported CSV file.
# Defaults to the same folder as this script, with today's date in the name.
$CsvOutputPath   = Join-Path $PSScriptRoot "NinjaOne-Devices-$(Get-Date -Format 'yyyy-MM-dd').csv"

# Where to save the NinjaOne refresh token on this computer.
$TokenFile       = Join-Path $PSScriptRoot 'ninja_refresh_token.txt'

# ==============================================================================
#  DO NOT EDIT BELOW THIS LINE
# ==============================================================================

$RedirectUri     = 'https://localhost'
$AuthEndpoint    = "$BaseUrl/ws/oauth/authorize"
$Scope           = 'monitoring management control offline_access'

# ------------------------------------------------------------------------------
#  FUNCTION: Refresh an existing access token
# ------------------------------------------------------------------------------

function Get-AccessTokenFromRefresh {
    param ([string]$RefreshToken)

    $Body = @{
        grant_type    = 'refresh_token'
        refresh_token = $RefreshToken
        client_id     = $ClientId
        client_secret = $ClientSecret
    }

    try {
        return Invoke-RestMethod -Method Post -Uri $TokenEndpoint `
            -ContentType 'application/x-www-form-urlencoded' -Body $Body
    } catch {
        return $null
    }
}

# ------------------------------------------------------------------------------
#  FUNCTION: Full browser-based authorization code flow
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
    Write-Host ' Step 3 of 3 — Exchanging code for access token...'
    Write-Host '---------------------------------------------------------------'

    $TokenBody = @{
        grant_type    = 'authorization_code'
        code          = $AuthCode
        redirect_uri  = $RedirectUri
        client_id     = $ClientId
        client_secret = $ClientSecret
    }

    try {
        return Invoke-RestMethod -Method Post -Uri $TokenEndpoint `
            -ContentType 'application/x-www-form-urlencoded' -Body $TokenBody
    } catch {
        Write-Host ''
        Write-Host " ERROR: Could not get an access token. $_" -ForegroundColor Red
        Write-Host '        Double-check your ClientId, ClientSecret, and TokenEndpoint.' -ForegroundColor Red
        exit 1
    }
}

# ------------------------------------------------------------------------------
#  STEP 1 — Authenticate (use saved refresh token, or fall back to browser login)
# ------------------------------------------------------------------------------

Write-Host ''
Write-Host '============================================================='
Write-Host ' NinjaOne Device Detail Exporter'
Write-Host '============================================================='

$AccessToken  = $null
$RefreshToken = $null

if (Test-Path $TokenFile) {
    Write-Host ''
    Write-Host ' Found saved refresh token. Renewing access...'
    $Saved    = (Get-Content $TokenFile -Raw).Trim()
    $Response = Get-AccessTokenFromRefresh -RefreshToken $Saved

    if ($Response -and $Response.access_token) {
        $AccessToken  = $Response.access_token
        $RefreshToken = $Response.refresh_token
        Write-Host ' Access token renewed successfully.'
    } else {
        Write-Host ' Saved token has expired. A browser login is required.' -ForegroundColor Yellow
        Remove-Item $TokenFile -Force
    }
}

if (-not $AccessToken) {
    $Response = Invoke-AuthCodeFlow
    if (-not $Response.access_token) {
        Write-Host ' ERROR: No access token received.' -ForegroundColor Red
        exit 1
    }
    $AccessToken  = $Response.access_token
    $RefreshToken = $Response.refresh_token
    Write-Host ' Access token obtained successfully.'
}

if ($RefreshToken) {
    $RefreshToken | Set-Content $TokenFile -Force
    Write-Host " Refresh token saved to: $TokenFile"
} else {
    Write-Host ' WARNING: No refresh token returned. Check offline_access is enabled on your API app.' -ForegroundColor Yellow
}

$Headers = @{
    Authorization  = "Bearer $AccessToken"
    'Content-Type' = 'application/json'
}

# ------------------------------------------------------------------------------
#  STEP 2 — Pull all devices with expanded fields
# ------------------------------------------------------------------------------

Write-Host ''
Write-Host ' Fetching devices from NinjaOne...'

# Fetch all devices in one call — NinjaOne paginates at 1000 by default.
# The loop below handles environments with more than 1000 devices automatically.

$AllDevices  = [System.Collections.Generic.List[object]]::new()
$PageSize    = 1000
$After       = $null

do {
    $Url = "$BaseUrl/v2/devices?expand=organization,location,warranty&pageSize=$PageSize"
    if ($After) { $Url += "&after=$After" }

    try {
        $Page = Invoke-RestMethod -Method Get -Uri $Url -Headers $Headers
    } catch {
        Write-Host ''
        Write-Host " ERROR: Failed to retrieve devices. $_" -ForegroundColor Red
        Write-Host '        Check your BaseUrl and that your account has permission to view devices.' -ForegroundColor Red
        exit 1
    }

    if ($Page -and $Page.Count -gt 0) {
        $AllDevices.AddRange([object[]]$Page)
        $After = $Page[-1].id
        Write-Host "  Retrieved $($AllDevices.Count) devices so far..."
    } else {
        break
    }

} while ($Page.Count -eq $PageSize)

Write-Host " Total devices retrieved: $($AllDevices.Count)"

# ------------------------------------------------------------------------------
#  STEP 3 — Flatten and export to CSV
# ------------------------------------------------------------------------------

Write-Host ''
Write-Host ' Building CSV...'

function ConvertFrom-UnixSeconds {
    param ($Seconds)
    if ($null -eq $Seconds -or $Seconds -eq 0) { return '' }
    return (Get-Date '1970-01-01 00:00:00Z').AddSeconds($Seconds).ToLocalTime().ToString('yyyy-MM-dd')
}

$CsvRows = $AllDevices | ForEach-Object {
    $Device = $_

    # --- Core device fields ---------------------------------------------------
    $DeviceId       = $Device.id
    $Uid            = $Device.uid
    $DisplayName    = $Device.displayName
    $SystemName     = $Device.systemName
    $DnsName        = $Device.dnsName
    $NodeClass      = $Device.nodeClass
    $NodeRoleId     = $Device.nodeRoleId
    $Online         = $Device.online
    $LastContact    = if ($Device.lastContact) {
                          (Get-Date '1970-01-01 00:00:00Z').AddSeconds($Device.lastContact).ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss')
                      } else { '' }
    $Created        = if ($Device.created) {
                          (Get-Date '1970-01-01 00:00:00Z').AddSeconds($Device.created).ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss')
                      } else { '' }
    $IpAddresses    = if ($Device.ipAddresses)  { $Device.ipAddresses -join ', ' }  else { '' }
    $MacAddresses   = if ($Device.macAddresses) { $Device.macAddresses -join ', ' } else { '' }
    $PublicIp       = $Device.publicIP

    # --- System / OS fields ---------------------------------------------------
    $Manufacturer   = $Device.system.manufacturer
    $Model          = $Device.system.model
    $BiosSerial     = $Device.system.biosSerialNumber
    $OsName         = $Device.os.name
    $OsBuild        = $Device.os.buildNumber
    $OsVersion      = $Device.os.version
    $AgentVersion   = $Device.agentVersion

    # --- Expanded: Organization -----------------------------------------------
    $OrgId          = $Device.organization.id
    $OrgName        = $Device.organization.name
    $OrgDescription = $Device.organization.description
    $OrgWebsite     = $Device.organization.website

    # --- Expanded: Location ---------------------------------------------------
    $LocationId     = $Device.location.id
    $LocationName   = $Device.location.name
    $LocationAddr   = $Device.location.address
    $LocationCity   = $Device.location.city
    $LocationState  = $Device.location.state
    $LocationZip    = $Device.location.zipCode
    $LocationCountry = $Device.location.country

    # --- Expanded: Warranty ---------------------------------------------------
    $WarrantyStart  = ConvertFrom-UnixSeconds $Device.warranty.startDate
    $WarrantyEnd    = ConvertFrom-UnixSeconds $Device.warranty.endDate
    $WarrantyMfr    = ConvertFrom-UnixSeconds $Device.warranty.manufacturerFulfillmentDate

    [PSCustomObject]@{
        # Core
        'Device ID'                      = $DeviceId
        'UID'                            = $Uid
        'Display Name'                   = $DisplayName
        'System Name'                    = $SystemName
        'DNS Name'                       = $DnsName
        'Node Class'                     = $NodeClass
        'Node Role ID'                   = $NodeRoleId
        'Online'                         = $Online
        'Last Contact'                   = $LastContact
        'Created'                        = $Created
        'IP Addresses'                   = $IpAddresses
        'MAC Addresses'                  = $MacAddresses
        'Public IP'                      = $PublicIp
        'Agent Version'                  = $AgentVersion
        # System / OS
        'Manufacturer'                   = $Manufacturer
        'Model'                          = $Model
        'BIOS Serial Number'             = $BiosSerial
        'OS Name'                        = $OsName
        'OS Build Number'                = $OsBuild
        'OS Version'                     = $OsVersion
        # Organization (expanded)
        'Organization ID'                = $OrgId
        'Organization Name'              = $OrgName
        'Organization Description'       = $OrgDescription
        'Organization Website'           = $OrgWebsite
        # Location (expanded)
        'Location ID'                    = $LocationId
        'Location Name'                  = $LocationName
        'Location Address'               = $LocationAddr
        'Location City'                  = $LocationCity
        'Location State'                 = $LocationState
        'Location Zip'                   = $LocationZip
        'Location Country'               = $LocationCountry
        # Warranty (expanded)
        'Warranty Start Date'            = $WarrantyStart
        'Warranty End Date'              = $WarrantyEnd
        'Warranty Mfr Fulfillment Date'  = $WarrantyMfr
    }
}

$CsvRows | Export-Csv -Path $CsvOutputPath -NoTypeInformation -Encoding UTF8

Write-Host ''
Write-Host '============================================================='
Write-Host ' Done!'
Write-Host "  Devices exported : $($CsvRows.Count)"
Write-Host "  CSV saved to     : $CsvOutputPath"
Write-Host '============================================================='
