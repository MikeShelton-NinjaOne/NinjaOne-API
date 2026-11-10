<#
    Sync-InsightVM-To-NinjaOne.ps1

    WHAT THIS SCRIPT DOES
    ----------------------
    1. Connects to Rapid7 InsightVM (Insight Platform) using an API key and pulls the
       current list of vulnerabilities (CVEs) found on your assets.
    2. Throws out anything below the severity level you set below.
    3. Builds a CSV file in the format NinjaOne's Rapid7 Vulnerability Importer expects.
    4. Logs into NinjaOne and uploads that CSV into the Scan Group you specify.
    5. Writes a plain-English log entry so you can see what happened on each run.

    Every time this runs, it uploads the FULL current list of qualifying CVEs.
    NinjaOne always shows the data from the most recent upload, so you don't need
    to worry about old CVEs piling up or duplicating - each run replaces the last.

    REQUIREMENTS (one-time, before this script will work)
    --------------------------------------------------------
    This script needs PowerShell 7 or newer (not the older "PowerShell 5.1" that comes
    built into Windows). If you're not sure which you have, open a PowerShell window
    and type: $PSVersionTable.PSVersion
    If the first number is less than 7, download PowerShell 7 from Microsoft
    (search "install PowerShell 7 Windows") and run this script with that instead.

    ONE-TIME SETUP CHECKLIST (do all of this before running the script for the first time)
    -------------------------------------------------------------------------------------
    [ ] In Rapid7: create an API key.
        Log in at https://insight.rapid7.com, click the gear icon, choose "API Keys",
        then "Generate New User Key" (or create an organization key). Select your
        organization, give the key a name, and submit. Copy the key right away -
        Rapid7 only shows it once.
        The user the key belongs to must have a role that can read asset data
        (for example Asset Owner, Security Manager or Site Owner).
        Note which region your Rapid7 account is in (us, us2, us3, eu, ca, au, ap,
        aps2 or me1) - you'll need it below.

    [ ] In NinjaOne: create an OAuth API app.
        Go to Administration > Apps > API > Client App IDs > Add.
        Choose "API Services (machine-to-machine)" as the platform.
        Under Scopes, make sure "Monitoring" and "Management" are both checked.
        Save it, then copy the Client ID and Client Secret it gives you - you'll
        only be able to see the Secret once, so copy it somewhere safe right away.

    [ ] In NinjaOne: enable the Rapid7 Vulnerability Importer app.
        Go to Administration > Apps > Installed (or "Add Apps" if you don't see it),
        find Rapid7, and click Enable.

    [ ] In NinjaOne: create the Scan Group you want this script to update.
        Inside the Rapid7 app, go to the Scan Groups tab and click Create scan group.
        Give it a name (e.g. "Rapid7 - All Servers") and finish the setup wizard.
        Note the Scan Group ID shown for it - you'll need that number below.
        (This is a one-time manual step. After that, the script can update it via API.)

    Once all of that is done, fill in the settings below and you're ready to run it.
#>


# ======================================================================================
#  YOUR SETTINGS  -  This is the ONLY part of the file you should need to change.
#  Type your values between the quotation marks, then save the file.
#  WARNING: this file will contain your secrets. Do NOT upload it to GitHub (or share
#  it) with your real API key / Client Secret typed in.
# ======================================================================================

# --- Rapid7 connection ---

# Which Rapid7 region your account is in. Pick ONE of the lines below by removing the
# "#" in front of it, and put a "#" in front of the others. This must match your
# Rapid7 account, or the key will be rejected.
$Rapid7Region = "us"      # US
# $Rapid7Region = "us2"   # US2
# $Rapid7Region = "us3"   # US3
# $Rapid7Region = "eu"    # Europe
# $Rapid7Region = "ca"    # Canada
# $Rapid7Region = "au"    # Australia
# $Rapid7Region = "ap"    # Japan / Asia-Pacific
# $Rapid7Region = "aps2"  # India
# $Rapid7Region = "me1"   # Middle East

# The Rapid7 API key you created in the one-time setup checklist above.
# There is no username or password - the API key is all Rapid7 needs.
$Rapid7ApiKey = "your-rapid7-api-key"


# --- NinjaOne connection ---

# Which NinjaOne cloud your account lives in. This must match where you log in to
# NinjaOne. Pick ONE of the lines below by removing the "#" in front of it, and put
# a "#" in front of the others.
$NinjaOneBaseURL = "app.ninjarmm.com"      # US (default)
# $NinjaOneBaseURL = "us2.ninjarmm.com"    # US2
# $NinjaOneBaseURL = "eu.ninjarmm.com"     # Europe / Middle East
# $NinjaOneBaseURL = "ca.ninjarmm.com"     # Canada
# $NinjaOneBaseURL = "oc.ninjarmm.com"     # Australia / Oceania

# The Client ID and Client Secret from the NinjaOne API app you created in the
# one-time setup checklist above.
$NinjaOneClientID     = "your-ninjaone-client-id"
$NinjaOneClientSecret = "your-ninjaone-client-secret"

# The ID number of the Scan Group you created in NinjaOne (see setup checklist above).
# This must already exist - the script updates it, it does not create it.
# Example: 1234
$NinjaOneScanGroupID = 1234


# --- What to sync ---

# Only CVEs with a CVSS severity score at or above this number will be sent to NinjaOne.
# A higher number means fewer, more serious CVEs get synced.
#   10.0 = only the most critical CVEs
#    7.0 = "high" severity and above (a common starting point)
#    0.0 = sync everything, no filtering
$MinimumCVSSSeverity = 7.0

# How NinjaOne should identify which device each CVE belongs to.
# Most people should leave this as "Hostname". If your NinjaOne Scan Group was set up
# to match devices by IP address instead, change this to "IPAddress".
#   "Hostname"   - matches devices by computer name
#   "IPAddress"  - matches devices by IP address
$DeviceIdField = "Hostname"

# Where this script should write its log file. The default writes a file called
# sync-log.txt in the same folder as this script.
$LogFilePath = ".\sync-log.txt"


# ======================================================================================
#  DO NOT EDIT BELOW THIS LINE
#  Everything past this point is the part that does the actual work. You shouldn't
#  need to change anything below here unless you're comfortable with PowerShell.
# ======================================================================================

$ErrorActionPreference = "Stop"

# Rapid7's cloud API address for your region.
$InsightVMBaseURL = "https://$Rapid7Region.api.insight.rapid7.com/vm/v4/integration"

# ---------------------------------------------------------------------------------
# Small helper: writes a timestamped line to both the screen and the log file.
# ---------------------------------------------------------------------------------
function Write-Log {
    param([string]$Message)
    $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
    Write-Host $line
    Add-Content -Path $LogFilePath -Value $line
}

# ---------------------------------------------------------------------------------
# Builds the API key header Rapid7's API expects.
# ---------------------------------------------------------------------------------
function Get-InsightVMHeaders {
    return @{
        "X-Api-Key"    = $Rapid7ApiKey
        "Accept"       = "application/json"
        "Content-Type" = "application/json"
    }
}

# ---------------------------------------------------------------------------------
# Runs a Rapid7 "search" request and keeps asking for the next page (using the
# cursor Rapid7 hands back) until there isn't one. Returns every record found.
# ---------------------------------------------------------------------------------
function Invoke-InsightVMSearch {
    param(
        [string]$Path,
        [string]$ExtraQuery = ""
    )

    $headers = Get-InsightVMHeaders
    $allRecords = @()
    $cursor = $null
    $pageSize = 500

    do {
        $url = "$InsightVMBaseURL/$Path`?size=$pageSize"
        if ($ExtraQuery) { $url += "&$ExtraQuery" }
        if ($cursor)     { $url += "&cursor=" + [uri]::EscapeDataString($cursor) }

        $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Post -Body "{}"

        $records = @($response.data) | Where-Object { $null -ne $_ }
        $allRecords += $records
        $cursor = $response.metadata.cursor
    } while ($records.Count -gt 0 -and $cursor)

    return $allRecords
}

# ---------------------------------------------------------------------------------
# Pulls every asset from InsightVM. In Rapid7's cloud API, each asset already comes
# back with its vulnerability findings attached, so there is no separate
# per-asset request.
# ---------------------------------------------------------------------------------
function Get-InsightVMAssets {
    # includeSame=true asks Rapid7 to include findings that haven't changed,
    # not only new ones.
    return Invoke-InsightVMSearch -Path "assets" -ExtraQuery "includeSame=true"
}

# ---------------------------------------------------------------------------------
# Pulls the full list of vulnerability definitions (CVE number(s) and CVSS severity)
# once, and stores them in memory by vulnerability ID so each one can be looked up
# instantly for every device it appears on.
# ---------------------------------------------------------------------------------
$script:VulnDetailCache = @{}

function Initialize-InsightVMVulnerabilityDetails {
    $vulns = Invoke-InsightVMSearch -Path "vulnerabilities"
    foreach ($v in $vulns) {
        if ($v.id) { $script:VulnDetailCache[[string]$v.id] = $v }
    }
}

# ---------------------------------------------------------------------------------
# Logs into NinjaOne and returns a bearer token to use on the upload request.
# ---------------------------------------------------------------------------------
function Get-NinjaOneAccessToken {
    $tokenUrl = "https://$NinjaOneBaseURL/ws/oauth/token"
    $body = @{
        grant_type    = "client_credentials"
        client_id     = $NinjaOneClientID
        client_secret = $NinjaOneClientSecret
        scope         = "monitoring management"
    }

    $response = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $body -ContentType "application/x-www-form-urlencoded"
    return $response.access_token
}

# ---------------------------------------------------------------------------------
# Uploads the finished CSV file to the specified NinjaOne Scan Group.
# ---------------------------------------------------------------------------------
function Send-CsvToNinjaOne {
    param(
        [string]$AccessToken,
        [string]$CsvPath
    )

    $uploadUrl = "https://$NinjaOneBaseURL/api/v2/vulnerability/scan-groups/$NinjaOneScanGroupID/upload"
    $headers = @{ Authorization = "Bearer $AccessToken" }

    Invoke-RestMethod -Uri $uploadUrl -Headers $headers -Method Post -Form @{ csv = Get-Item -Path $CsvPath }
}

# ======================================================================================
#  MAIN - this is what actually runs when you double-click or execute the script.
# ======================================================================================

Write-Log "===== Starting InsightVM -> NinjaOne sync ====="

try {

    # --- Step 1: Pull assets and their vulnerability findings from InsightVM ---
    Write-Log "Connecting to Rapid7 (region: $Rapid7Region) and retrieving asset list..."
    try {
        $assets = Get-InsightVMAssets
        Initialize-InsightVMVulnerabilityDetails
    }
    catch {
        Write-Log "COULD NOT CONNECT TO RAPID7. Things to check: is the Rapid7 region correct for your account, is the API key correct and still valid, and does the key's user have permission to read asset data? Details: $($_.Exception.Message)"
        throw
    }
    Write-Log "Found $($assets.Count) assets in InsightVM."

    $rows = @()
    $totalRawFindings = 0

    foreach ($asset in $assets) {

        # Figure out which identifier to use for this device, based on your setting above.
        if ($DeviceIdField -eq "IPAddress") {
            $deviceId = $asset.ip
        }
        else {
            $deviceId = $asset.host_name
        }

        if ([string]::IsNullOrWhiteSpace($deviceId)) {
            # Skip assets that don't have the identifier you chose - there's nothing
            # useful to send NinjaOne for a device we can't identify.
            continue
        }

        # The asset already carries its findings. "new" and "same" are the findings
        # that are currently present; "remediated" (already fixed) is left out.
        $findings = @($asset.new) + @($asset.same) | Where-Object { $null -ne $_ }
        $totalRawFindings += $findings.Count

        foreach ($finding in $findings) {
            $detail = $script:VulnDetailCache[[string]$finding.vulnerability_id]
            if (-not $detail) { continue }

            $severity = $detail.cvss_v3_score
            if (-not $severity) { $severity = $detail.cvss_v2_score }
            if (-not $severity) { $severity = 0 }

            if ([double]$severity -lt $MinimumCVSSSeverity) {
                continue
            }

            # A single vulnerability entry can list more than one CVE ID. Rapid7 may
            # return them as a list or as one comma-separated piece of text, so
            # handle both.
            $cveIds = @()
            if ($detail.cves -is [string]) {
                $cveIds = $detail.cves -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ }
            }
            elseif ($detail.cves) {
                $cveIds = @($detail.cves)
            }

            if ($cveIds.Count -eq 0) {
                # Some vulnerabilities in InsightVM don't have a CVE assigned at all
                # (e.g. vendor-specific advisories) - skip those since NinjaOne's
                # importer is expecting a CVE ID.
                continue
            }

            foreach ($cveId in $cveIds) {
                $rows += [PSCustomObject]@{
                    Vendor  = "Rapid7"
                    "Device ID" = $deviceId
                    "CVE ID"    = $cveId
                }
            }
        }
    }

    Write-Log "Pulled $totalRawFindings raw findings from InsightVM. $($rows.Count) rows remain after applying your severity filter ($MinimumCVSSSeverity+) and removing entries without a CVE ID or device identifier."

    if ($rows.Count -eq 0) {
        Write-Log "No qualifying CVEs found - nothing to upload this run. Ending."
        Write-Log "===== Sync finished ====="
        return
    }

    # De-duplicate in case the same CVE shows up more than once for the same device
    # (this can happen if InsightVM reports the same finding under slightly different
    # scan contexts).
    $rows = $rows | Sort-Object "Device ID", "CVE ID" -Unique
    Write-Log "$($rows.Count) unique device/CVE rows after removing duplicates."

    # --- Step 2: Write the CSV file NinjaOne expects ---
    $csvPath = Join-Path -Path (Split-Path -Path $LogFilePath -Parent -ErrorAction SilentlyContinue) -ChildPath "insightvm-export.csv"
    if ([string]::IsNullOrWhiteSpace($csvPath) -or $csvPath -eq "\insightvm-export.csv") {
        $csvPath = ".\insightvm-export.csv"
    }
    $rows | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Log "Wrote CSV file to $csvPath."

    # --- Step 3: Log into NinjaOne ---
    Write-Log "Connecting to NinjaOne..."
    try {
        $token = Get-NinjaOneAccessToken
    }
    catch {
        Write-Log "COULD NOT LOG IN TO NINJAONE. Things to check: is the NinjaOne region/base URL correct, and are the Client ID/Client Secret correct and still valid? Details: $($_.Exception.Message)"
        throw
    }
    Write-Log "Connected to NinjaOne successfully."

    # --- Step 4: Upload the CSV to the Scan Group ---
    Write-Log "Uploading CSV to Scan Group ID $NinjaOneScanGroupID..."
    try {
        Send-CsvToNinjaOne -AccessToken $token -CsvPath $csvPath
    }
    catch {
        Write-Log "THE UPLOAD TO NINJAONE FAILED. Things to check: does Scan Group ID $NinjaOneScanGroupID exist in NinjaOne, and does the API app have both the Monitoring and Management scopes enabled? Details: $($_.Exception.Message)"
        throw
    }
    Write-Log "Upload succeeded. $($rows.Count) rows sent to NinjaOne Scan Group $NinjaOneScanGroupID."

}
catch {
    Write-Log "Sync stopped early because of an error. See the message above for what to check."
}

Write-Log "===== Sync finished ====="
