
# Get_EntraDevices.ps1
#
# Output:
#   All Microsoft Entra ID devices
#   Device name, status, OS, version, join type, owner,
#   MDM, compliance, registration date, last activity,
#   days inactive, and device ID
#
# Export:
#   CSV: Same output
#
# Read-only: Does not modify Entra ID or devices.


####################################################################################
#################### MICROSOFT GRAPH CONNECTION ####################################
####################################################################################
$modules = @(
    "Microsoft.Graph.Identity.DirectoryManagement"
)

foreach ($module in $modules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Install-Module $module -Scope CurrentUser -ErrorAction Stop
    }

    Import-Module $module -ErrorAction Stop
}

$requiredScopes = @(
    "Device.Read.All",
    "User.Read.All"
)

$connection = Get-MgContext

if (-not $connection -or
    @($requiredScopes | Where-Object { $_ -notin $connection.Scopes }).Count -gt 0) {

    Connect-MgGraph -Scopes $requiredScopes -NoWelcome
}


####################################################################################
#################### RETRIEVE DEVICES ##############################################
####################################################################################
Write-Host "`nRetrieving Entra ID devices..." -ForegroundColor Cyan

$devices = @(
    Get-MgDevice -All -ExpandProperty "registeredOwners" -Property @(
        "id",
        "deviceId",
        "displayName",
        "accountEnabled",
        "operatingSystem",
        "operatingSystemVersion",
        "trustType",
        "managementType",
        "isManaged",
        "isCompliant",
        "registrationDateTime",
        "approximateLastSignInDateTime"
    ) -ErrorAction Stop
)


####################################################################################
#################### PROCESS RESULTS ###############################################
####################################################################################
$today = Get-Date

$results = foreach ($device in $devices) {

    $joinType = switch ($device.TrustType) {
        "AzureAd" { "Microsoft Entra joined" }
        "Workplace" { "Microsoft Entra registered" }
        "ServerAd" { "Microsoft Entra hybrid joined" }
        default { $device.TrustType }
    }

    $owners = @(
        foreach ($owner in $device.RegisteredOwners) {
            $name = $owner.AdditionalProperties['displayName']

            if (-not $name) {
                $name = $owner.Id
            }

            $name
        }
    )

    $ownerName = if ($owners.Count -gt 0) {
        $owners -join "; "
    }
    else {
        "None"
    }

    $lastActivity = $device.ApproximateLastSignInDateTime


    $daysInactive = if ($lastActivity) {
        $activityDate = [datetime]$lastActivity
        [math]::Max(0, [math]::Floor(($today - $activityDate).TotalDays))
    }
    else {
        $null
    }


    $mdm = if (-not $device.IsManaged) {
        "None"
    }
    elseif ($device.ManagementType) {
        $device.ManagementType
    }
    else {
        "Managed"
    }

    $compliant = if ($null -eq $device.IsCompliant) {
        "N/A"
    }
    elseif ($device.IsCompliant) {
        "Yes"
    }
    else {
        "No"
    }

    [PSCustomObject]@{
        Name         = $device.DisplayName
        Enabled      = if ($device.AccountEnabled) { "Yes" } else { "No" }
        OS           = $device.OperatingSystem
        Version      = $device.OperatingSystemVersion
        JoinType     = $joinType
        Owner        = $ownerName
        MDM          = $mdm
        Compliant    = $compliant
        Registered   = if ($device.RegistrationDateTime) {
            $device.RegistrationDateTime.LocalDateTime
        }
        else { $null }
        Activity     = if ($lastActivity) {
            $lastActivity.LocalDateTime
        }
        else { $null }
        DaysInactive = $daysInactive
        DeviceId     = $device.DeviceId
    }
}

$results = @($results | Sort-Object Activity -Descending)



####################################################################################
#################### DISPLAY AND EXPORT ############################################
####################################################################################

Write-Host "`nSelect console sorting order:" -ForegroundColor Cyan
Write-Host "1. Owner"
Write-Host "2. Days Inactive"
Write-Host "3. Name"

$choice = Read-Host "`nEnter choice (1-3)"

switch ($choice) {
    "1" {
        $sortedResults = $results | Sort-Object Owner, Name
    }
    "2" {
        $sortedResults = $results | Sort-Object `
        @{Expression = { $null -eq $_.DaysInactive } },
        @{Expression = { $_.DaysInactive }; Descending = $true }
    }
    "3" {
        $sortedResults = $results | Sort-Object Name
    }
    default {
        Write-Host "Invalid choice. Sorting by Name." -ForegroundColor Yellow
        $sortedResults = $results | Sort-Object Name
    }
}

Write-Host "`nTotal devices: $($results.Count)" -ForegroundColor Green

$sortedResults |
Format-Table Name, Owner, Enabled, Compliant, DaysInactive -AutoSize |
Out-Host

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$csvPath = Join-Path $PSScriptRoot "EntraDevices_$timestamp.csv"

$results |
    Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

Write-Host "`nCSV exported to: $csvPath" -ForegroundColor Green


