#################################################################################
#  SIMPLE NINJAONE CUSTOM FIELD IMPORT
#
#  HOW TO USE:
#    1. Fill in the SETTINGS section below (only the 5 lines marked  <<< EDIT)
#    2. Save the file
#    3. Right-click it > "Run with PowerShell"  (or run it from a PowerShell window)
#    4. Enter your API Client Secret when asked
#    5. Review the results shown on screen and in the results CSV saved next to your import file.
#################################################################################


#=============================== SETTINGS ======================================

# 1. Path to your CSV file
$CsvPath      = "C:\Temp\import.csv"                       # <<< EDIT

# 2. Name of the CSV column that contains the NinjaOne Device ID (a number, e.g. 1234)
$DeviceIdColumn = "DeviceID"                               # <<< EDIT

# 3. Which CSV column goes into which Ninja custom field.
#    Format:  "CSV Column Name" = "ninjaFieldMachineName"
#    (add more lines if you have more than one field)
$FieldMap = @{
    "Notes" = "deviceNotes"                                # <<< EDIT
}

# 4. NinjaOne API Client ID
$ClientId     = "PASTE-CLIENT-ID-HERE"                     # <<< EDIT

# 5. Your NinjaOne region: app, eu, oc, ca, us2, or fedgov
$Region       = "app"                                      # <<< EDIT

#============================ END OF SETTINGS ===================================
# Nothing below this line needs to be changed.


$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# ---- Basic checks -------------------------------------------------------------
if (-not (Test-Path $CsvPath)) { Write-Host "ERROR: CSV file not found: $CsvPath" -ForegroundColor Red; Read-Host "Press Enter to exit"; exit }
if ($ClientId -like "PASTE*")  { Write-Host "ERROR: Please enter your Client ID in the settings." -ForegroundColor Red; Read-Host "Press Enter to exit"; exit }

$rows = Import-Csv -Path $CsvPath
if (-not $rows) { Write-Host "ERROR: The CSV is empty." -ForegroundColor Red; Read-Host "Press Enter to exit"; exit }

$columns = $rows[0].PSObject.Properties.Name
$needed  = @($DeviceIdColumn) + @($FieldMap.Keys)
foreach ($c in $needed) {
    if ($c -notin $columns) {
        Write-Host "ERROR: Column '$c' was not found in the CSV." -ForegroundColor Red
        Write-Host "Columns in your CSV are: $($columns -join ', ')" -ForegroundColor Yellow
        Read-Host "Press Enter to exit"; exit
    }
}

# ---- Ask for the secret -------------------------------------------------------
$secure = Read-Host "Enter your NinjaOne API Client Secret" -AsSecureString
$secret = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure))

$BaseUrl = switch ($Region) {
    'app'    { 'https://app.ninjarmm.com' }
    'eu'     { 'https://eu.ninjarmm.com' }
    'oc'     { 'https://oc.ninjarmm.com' }
    'ca'     { 'https://ca.ninjarmm.com' }
    'us2'    { 'https://us2.ninjarmm.com' }
    'fedgov' { 'https://fedgov.rmmservice.com' }
    default  { Write-Host "ERROR: Unknown region '$Region'" -ForegroundColor Red; Read-Host "Press Enter to exit"; exit }
}

# ---- Log in -------------------------------------------------------------------
try {
    # Build the login body by hand so special characters in the secret are encoded correctly
    $loginBody = "grant_type=client_credentials" +
        "&client_id="     + [uri]::EscapeDataString($ClientId.Trim()) +
        "&client_secret=" + [uri]::EscapeDataString($secret.Trim()) +
        "&scope="         + [uri]::EscapeDataString("monitoring management")
    $tokenResp = Invoke-RestMethod -Method Post -Uri "$BaseUrl/ws/oauth/token" `
        -ContentType 'application/x-www-form-urlencoded' -Body $loginBody
    $headers = @{ Authorization = "Bearer $($tokenResp.access_token)"; Accept = 'application/json' }
    Write-Host "Logged in to NinjaOne." -ForegroundColor Green
}
catch {
    Write-Host "ERROR: Could not log in. Check your Client ID, Secret and Region." -ForegroundColor Red
    Write-Host $_.Exception.Message
    if ($_.ErrorDetails.Message) { Write-Host "Details from NinjaOne: $($_.ErrorDetails.Message)" -ForegroundColor Yellow }
    Write-Host "Trying to log in at: $BaseUrl/ws/oauth/token" -ForegroundColor Gray
    Read-Host "Press Enter to exit"; exit
}

# ---- Get all devices ----------------------------------------------------------
$devices = Invoke-RestMethod -Method Get -Uri "$BaseUrl/v2/devices" -Headers $headers
Write-Host "Found $($devices.Count) devices in Ninja."

$lookup = @{}
foreach ($d in $devices) { $lookup[[string]$d.id] = $d }

# ---- Process each CSV row -----------------------------------------------------
$results = @()
foreach ($row in $rows) {
    $name = ([string]$row.$DeviceIdColumn).Trim()
    $status = ""; $note = ""

    $device = if ($name) { $lookup[$name] } else { $null }

    if (-not $name)        { $status = "SKIPPED"; $note = "No device ID in this row" }
    elseif (-not $device)  { $status = "NOT FOUND"; $note = "No device with this ID in Ninja" }
    else {
        $payload = @{}
        foreach ($csvCol in $FieldMap.Keys) {
            $val = $row.$csvCol
            if ([string]::IsNullOrEmpty($val)) { continue }                          # skip empty cells
            $val = ($val -replace "`r`n", "`n") -replace "`r", "`n"                  # clean line breaks
            $payload[$FieldMap[$csvCol]] = $val
        }

        if ($payload.Count -eq 0) { $status = "SKIPPED"; $note = "All cells empty" }
        else {
            try {
                Invoke-RestMethod -Method Patch -Uri "$BaseUrl/v2/device/$($device.id)/custom-fields" `
                    -Headers $headers -ContentType 'application/json' `
                    -Body ($payload | ConvertTo-Json -Compress) | Out-Null
                $status = "UPDATED"; $note = ($payload.Keys -join ", ")
            }
            catch {
                $status = "ERROR"; $note = $_.Exception.Message
            }
            Start-Sleep -Milliseconds 150
        }
    }

    $color = switch ($status) { "UPDATED" {"Green"} "ERROR" {"Red"} "NOT FOUND" {"Yellow"} default {"Gray"} }
    $deviceName = if ($device) { $device.systemName } else { "" }
    Write-Host ("{0,-13} {1}  {2}  {3}" -f $status, $name, $deviceName, $note) -ForegroundColor $color

    $results += [pscustomobject]@{ DeviceId = $name; DeviceName = $deviceName; Status = $status; Note = $note }
}

# ---- Summary ------------------------------------------------------------------
$logFile = Join-Path (Split-Path -Parent $CsvPath) ("NinjaImportResults_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
$results | Export-Csv -Path $logFile -NoTypeInformation

Write-Host "`n------ SUMMARY ------" -ForegroundColor Cyan
$results | Group-Object Status | ForEach-Object { Write-Host ("{0,-13} {1}" -f $_.Name, $_.Count) }
Write-Host "Results saved to: $logFile"

Read-Host "`nPress Enter to close"
