<#
.SYNOPSIS
    Moves NinjaRMM devices to the correct Location based on their Public IP address.

.DESCRIPTION
    Reads a CSV that maps Public IP addresses to Locations, compares it against every
    device's current Public IP in NinjaRMM, and moves any device that is in the wrong
    Location.

    HOW TO USE THIS SCRIPT:
    1. Fill in the "USER CONFIGURATION" section below (API credentials, region URL,
       folder path, and optionally an Organization ID).
    2. Make sure your input CSV file exists in that folder and has the required columns
       (see the CONFIGURATION section for the exact filename and column names).
    3. Run the script. It will print progress and save 3 result CSV files in the same
       folder.

.NOTES
    Author : TawTek
    Date   : 2024-11-04
    Version: 2.0 (restructured for clarity)
#>

<#==========================================================================================
  USER CONFIGURATION — EDIT THE VALUES IN THIS SECTION BEFORE RUNNING THE SCRIPT
==========================================================================================#>

# --- NinjaRMM API credentials ---------------------------------------------------------
# Create these in NinjaRMM under: Administration > Apps > API Client.
# The client must have "monitoring" and "management" scopes enabled.
$NinjaClientId     = ''   # <-- REQUIRED: paste your API Client ID here
$NinjaClientSecret = ''   # <-- REQUIRED: paste your API Client Secret here

# --- NinjaRMM region URL ---------------------------------------------------------------
# Must match the region your NinjaRMM instance actually uses.
# Common examples: https://app.ninjarmm.com (US) | https://eu.ninjarmm.com (EU)
# | https://oc.ninjarmm.com (Oceania) | https://ca.ninjarmm.com (Canada)
$NinjaRegionUrl = 'https://app.ninjarmm.com'   # <-- REQUIRED: confirm this is correct

# --- Folder containing the input CSV, and where the output CSVs will be saved ---------
$WorkingFolder = 'C:\Temp'   # <-- REQUIRED: folder path, no trailing slash

# --- Name of the input CSV file (must live inside $WorkingFolder) ---------------------
$InputCsvFileName = 'ninjapublicips.csv'   # only change this if your file is named differently

# --- Optional: limit the script to a single Organization ------------------------------
# Leave as $null to process devices across ALL organizations.
# Set to a number (e.g. 12) to only process devices belonging to that Organization ID.
$OrganizationIdFilter = $null   # <-- OPTIONAL: set to an Org ID, or leave as $null for all orgs

<#==========================================================================================
  INPUT CSV REQUIREMENTS (for reference — the file itself is not edited by the script)
  --------------------------------------------------------------------------------------
  The CSV named above must contain these columns, spelled exactly:
      PublicIP    - the public IP address tied to a location
      OrgID       - the NinjaRMM Organization ID that location belongs to
      Location    - a friendly name for the location (used only for display/reporting)
      LocationID  - the NinjaRMM Location ID that matching devices should be moved into

  Example:
      PublicIP,OrgID,Location,LocationID
      203.0.113.10,12,Main Office,45
      203.0.113.55,12,Branch Office,46
      198.51.100.2,18,Warehouse,52

  Rules:
      - No two rows may share the same PublicIP. The script will stop and tell you
        which IPs are duplicated if this happens.
      - Blank PublicIP rows are ignored automatically.
==========================================================================================#>

<#--------------------------------------------------------------------------------------
  SCRIPT FUNCTIONS — you shouldn't need to edit anything below this line
--------------------------------------------------------------------------------------#>

function Get-NinjaAuthHeader {
    <#
        Requests a fresh OAuth access token from NinjaRMM and returns it as a header
        object ready to use with Invoke-RestMethod / Invoke-WebRequest.
    #>
    param(
        [Parameter(Mandatory)][string]$RegionUrl,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$ClientSecret
    )

    $TokenRequestBody = @{
        'grant_type'    = 'client_credentials'
        'client_id'     = $ClientId
        'client_secret' = $ClientSecret
        'scope'         = 'monitoring management'
    }

    $TokenResponse = Invoke-WebRequest -Uri "$RegionUrl/ws/oauth/token" `
                                        -Method POST `
                                        -Body $TokenRequestBody `
                                        -ContentType 'application/x-www-form-urlencoded' |
                      Select-Object -ExpandProperty Content | ConvertFrom-Json

    return @{
        'Authorization' = "Bearer $($TokenResponse.access_token)"
        'Content-Type'  = 'application/json'
    }
}

function Import-LocationMap {
    <#
        Imports and validates the Public IP -> Location CSV.
        Exits the script if the file is missing or contains duplicate Public IPs.
    #>
    param(
        [Parameter(Mandatory)][string]$CsvPath
    )

    if (-not (Test-Path $CsvPath)) {
        Write-Host "FAIL: Input CSV not found at $CsvPath" -ForegroundColor DarkRed
        Write-Host "      Check `$WorkingFolder and `$InputCsvFileName in the configuration section." -ForegroundColor DarkRed
        exit
    }

    $LocationMap = Import-Csv -Path $CsvPath | Where-Object { -not [string]::IsNullOrWhiteSpace($_.PublicIP) }

    $DuplicateIPs = $LocationMap | Group-Object -Property PublicIP | Where-Object { $_.Count -gt 1 }
    if ($DuplicateIPs) {
        Write-Host "FAIL: Duplicate Public IPs found in the CSV. Fix these rows and re-run:" -ForegroundColor DarkRed
        $DuplicateIPs | ForEach-Object { Write-Host "      $($_.Name)  (appears $($_.Count) times)" -ForegroundColor DarkRed }
        exit
    }

    return $LocationMap
}

function Move-NinjaDevicesByPublicIP {
    <#
        Main entry point. Fetches devices from NinjaRMM, compares each device's Public IP
        against the imported Location map, and moves any device that's in the wrong Location.
    #>
    param(
        [Parameter(Mandatory)][string]$RegionUrl,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$ClientSecret,
        [Parameter(Mandatory)][string]$WorkingFolder,
        [Parameter(Mandatory)][string]$InputCsvFileName,
        [Nullable[int]]$OrganizationIdFilter = $null
    )

    $InputCsvPath = Join-Path $WorkingFolder $InputCsvFileName

    $MovedDevices      = @()
    $FailedDevices     = @()
    $UnmatchedDevices  = @()
    $MatchedCount      = 0
    $UnmatchedCount    = 0

    Write-Host "INFO: Importing Location map from $InputCsvPath"
    $LocationMap = Import-LocationMap -CsvPath $InputCsvPath

    Write-Host "INFO: Requesting NinjaRMM API access token."
    $AuthHeader = Get-NinjaAuthHeader -RegionUrl $RegionUrl -ClientId $ClientId -ClientSecret $ClientSecret

    Write-Host "INFO: Fetching device list from NinjaRMM."
    $AllDevices = Invoke-RestMethod -Uri "$RegionUrl/v2/devices-detailed" -Method GET -Headers $AuthHeader

    $DevicesToCheck = if ($null -ne $OrganizationIdFilter) {
        $AllDevices | Where-Object { $_.organizationId -eq $OrganizationIdFilter }
    } else {
        $AllDevices
    }

    Write-Host "INFO: Comparing $($DevicesToCheck.Count) device(s) against the Location map."

    foreach ($Device in $DevicesToCheck) {

        $MatchingRows = $LocationMap | Where-Object { $_.PublicIP -eq $Device.publicIP }

        # No Public IP match at all -> log as unmatched and move to the next device
        if (-not $MatchingRows) {
            $UnmatchedDevices += [PSCustomObject]@{
                Device       = $Device.systemName
                Org          = ($LocationMap | Where-Object { $_.OrgID -eq $Device.organizationId } | Select-Object -First 1).Location
                Old_Location = ($LocationMap | Where-Object { $_.LocationID -eq $Device.locationId }).Location
                PublicIP     = $Device.publicIP
            }
            $UnmatchedCount++
            continue
        }

        foreach ($MatchedRow in $MatchingRows) {

            $MatchedCount++
            $PercentComplete = [math]::Round(($MatchedCount / ($DevicesToCheck.Count - $UnmatchedCount)) * 100)
            Write-Progress -Activity "Comparing Public IP to Location..." `
                            -Status "$MatchedCount/$($DevicesToCheck.Count - $UnmatchedCount) | $PercentComplete% Complete | $($Device.systemName)" `
                            -PercentComplete $PercentComplete

            # Already in the correct Location -> nothing to do
            if ($Device.locationId -eq $MatchedRow.LocationID) {
                continue
            }

            # Wrong Location -> attempt to move the device
            $PatchBody = @{ 'locationId' = $MatchedRow.LocationID } | ConvertTo-Json

            try {
                # Request a fresh token here in case the original has expired during a long run
                $AuthHeader = Get-NinjaAuthHeader -RegionUrl $RegionUrl -ClientId $ClientId -ClientSecret $ClientSecret
                Invoke-RestMethod -Uri "$RegionUrl/v2/device/$($Device.id)" -Method PATCH -Headers $AuthHeader -Body $PatchBody

                $MovedDevices += [PSCustomObject]@{
                    Device       = $Device.systemName
                    Org          = ($LocationMap | Where-Object { $_.OrgID -eq $Device.organizationId } | Select-Object -First 1).Location
                    Old_Location = ($LocationMap | Where-Object { $_.LocationID -eq $Device.locationId }).Location
                    New_Location = $MatchedRow.Location
                    PublicIP     = $Device.publicIP
                }
            } catch {
                $FailedDevices += [PSCustomObject]@{
                    Device          = $Device.systemName
                    Org             = ($LocationMap | Where-Object { $_.OrgID -eq $Device.organizationId } | Select-Object -First 1).Location
                    Old_Location    = ($LocationMap | Where-Object { $_.LocationID -eq $Device.locationId }).Location
                    Failed_Location = $MatchedRow.Location
                    PublicIP        = $Device.publicIP
                    Error           = $_.Exception.Message
                }
            }
        }
    }

    Write-Progress -Activity "Comparing Public IP to Location..." -Completed

    if ($MovedDevices.Count -gt 0) {
        $MovedDevices | Sort-Object Org, New_Location | Format-Table -AutoSize
        $MovedDevices | Sort-Object Org, New_Location | Export-Csv -Path (Join-Path $WorkingFolder 'movedendpoints.csv') -NoTypeInformation
        Write-Host "PASS: Moved devices exported to $(Join-Path $WorkingFolder 'movedendpoints.csv')" -ForegroundColor DarkGreen
    }

    if ($UnmatchedDevices.Count -gt 0) {
        $UnmatchedDevices | Sort-Object Org, Device | Format-Table -AutoSize
        $UnmatchedDevices | Sort-Object Org, Device | Export-Csv -Path (Join-Path $WorkingFolder 'unmatchedendpoints.csv') -NoTypeInformation
        Write-Host "WARN: Unmatched devices exported to $(Join-Path $WorkingFolder 'unmatchedendpoints.csv')" -ForegroundColor DarkYellow
    }

    if ($FailedDevices.Count -gt 0) {
        $FailedDevices | Sort-Object Org, Device | Format-Table -AutoSize
        $FailedDevices | Sort-Object Org, Device | Export-Csv -Path (Join-Path $WorkingFolder 'failedendpoints.csv') -NoTypeInformation
        Write-Host "FAIL: Failed moves exported to $(Join-Path $WorkingFolder 'failedendpoints.csv')" -ForegroundColor DarkRed
    }

    if ($MovedDevices.Count -eq 0 -and $UnmatchedDevices.Count -eq 0 -and $FailedDevices.Count -eq 0) {
        Write-Host "INFO: No devices needed to move — all Public IPs matched their correct Location." -ForegroundColor DarkCyan
    }
}

<#--------------------------------------------------------------------------------------
  RUN
  --------------------------------------------------------------------------------------
  This calls the script using the values you filled in under USER CONFIGURATION above.
  You do not need to edit this part.
--------------------------------------------------------------------------------------#>

if ([string]::IsNullOrWhiteSpace($NinjaClientId) -or [string]::IsNullOrWhiteSpace($NinjaClientSecret)) {
    Write-Host "FAIL: `$NinjaClientId and `$NinjaClientSecret must be filled in under USER CONFIGURATION before running this script." -ForegroundColor DarkRed
    exit
}

Move-NinjaDevicesByPublicIP -RegionUrl $NinjaRegionUrl `
                             -ClientId $NinjaClientId `
                             -ClientSecret $NinjaClientSecret `
                             -WorkingFolder $WorkingFolder `
                             -InputCsvFileName $InputCsvFileName `
                             -OrganizationIdFilter $OrganizationIdFilter
