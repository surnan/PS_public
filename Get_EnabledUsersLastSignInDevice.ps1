# Get_EnabledUsersLastSignInDevice.ps1
#
# Reports the latest successful interactive Entra sign-in
# for every enabled account, including enabled guests.
#
# Output:
#   Console summary
#   Timestamped CSV beside this script
#
# Read-only: Does not modify users or devices.
# Requires PowerShell 7.

####################################################################################
#################### REQUIREMENTS ###################################################
####################################################################################

if ($PSVersionTable.PSVersion.Major -lt 7) {
    throw "Run this script in PowerShell 7."
}

$modules = @(
    "Microsoft.Graph.Users",
    "Microsoft.Graph.Reports"
)

foreach ($module in $modules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Install-Module $module -Scope CurrentUser -ErrorAction Stop
    }

    Import-Module $module -ErrorAction Stop
}

####################################################################################
#################### MICROSOFT GRAPH CONNECTION ####################################
####################################################################################

Connect-MgGraph -Scopes "User.Read.All", "AuditLog.Read.All" `
    -ContextScope Process -NoWelcome -ErrorAction Stop

####################################################################################
#################### INPUT ##########################################################
####################################################################################

$inputDays = Read-Host "Number of days to search [Enter = 30]"

if ([string]::IsNullOrWhiteSpace($inputDays)) {
    $days = 30
}
else {
    $days = 0

    if (-not [int]::TryParse($inputDays, [ref]$days) -or $days -lt 1) {
        throw "Enter a positive whole number of days."
    }
}

$startDateFilter = (Get-Date).ToUniversalTime().AddDays(-$days).
    ToString("yyyy-MM-ddTHH:mm:ssZ")

# Number of users checked simultaneously.
$throttleLimit = 4

####################################################################################
#################### GET ENABLED USERS ##############################################
####################################################################################

Write-Host "`nGetting enabled accounts..." -ForegroundColor Cyan

$users = @(
    Get-MgUser -All `
        -Filter "accountEnabled eq true" `
        -Property Id, DisplayName, UserPrincipalName, AccountEnabled, UserType `
        -ErrorAction Stop |
        Sort-Object DisplayName
)

if ($users.Count -eq 0) {
    Write-Host "No enabled accounts found."
    return
}

####################################################################################
#################### GET LAST SIGN-IN DEVICE IN PARALLEL ############################
####################################################################################

Write-Host "Checking $($users.Count) enabled accounts..."
Write-Host "Concurrent requests: $throttleLimit"
Write-Host "Search period: Last $days days`n"

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

$results = @(
    $users | ForEach-Object -Parallel {
        $user = $_
        $searchStart = $using:startDateFilter

        # Build the same output structure for successful and failed queries.
        $row = [PSCustomObject]@{
            DisplayName       = $user.DisplayName
            UserPrincipalName = $user.UserPrincipalName
            AccountEnabled    = $user.AccountEnabled
            UserType          = $user.UserType
            LastSignIn        = $null
            DeviceName        = $null
            OperatingSystem   = $null
            Browser           = $null
            DeviceId          = $null
            Managed           = $null
            Compliant         = $null
            IPAddress         = $null
            Application       = $null
            QueryStatus       = $null
            ErrorMessage      = $null
        }

        try {
            Import-Module Microsoft.Graph.Reports -ErrorAction Stop

            $filter = "userId eq '$($user.Id)' " +
                      "and createdDateTime ge $searchStart " +
                      "and status/errorCode eq 0 " +
                      "and isInteractive eq true"

            # Graph returns the newest sign-in first by default.
            $signIn = Get-MgAuditLogSignIn `
                -Filter $filter `
                -Top 1 `
                -ErrorAction Stop |
                Select-Object -First 1

            if ($null -ne $signIn) {
                $row.LastSignIn = ([datetimeoffset]$signIn.CreatedDateTime).
                    ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss zzz")

                $row.DeviceName = $signIn.DeviceDetail.DisplayName

                if ([string]::IsNullOrWhiteSpace($row.DeviceName)) {
                    $row.DeviceName = "Device name not provided"
                }

                $row.OperatingSystem = $signIn.DeviceDetail.OperatingSystem
                $row.Browser         = $signIn.DeviceDetail.Browser
                $row.DeviceId        = $signIn.DeviceDetail.DeviceId
                $row.Managed         = $signIn.DeviceDetail.IsManaged
                $row.Compliant       = $signIn.DeviceDetail.IsCompliant
                $row.IPAddress       = $signIn.IpAddress
                $row.Application     = $signIn.AppDisplayName
                $row.QueryStatus     = "Success"
            }
            else {
                $row.DeviceName  = "No successful sign-in in available logs"
                $row.QueryStatus = "No matching sign-in"
            }
        }
        catch {
            $row.DeviceName   = "Query failed - see ErrorMessage"
            $row.QueryStatus  = "Failed"
            $row.ErrorMessage = $_.Exception.Message
        }

        $row
    } -ThrottleLimit $throttleLimit
)

$stopwatch.Stop()

# Parallel requests finish in varying order.
$results = @($results | Sort-Object DisplayName)

####################################################################################
#################### CSV EXPORT #####################################################
####################################################################################

$outputFolder = if ($PSScriptRoot) {
    $PSScriptRoot
}
else {
    (Get-Location).Path
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$csvPath = Join-Path $outputFolder "EnabledUsers_LastSignInDevice_$timestamp.csv"

$results |
    Export-Csv -Path $csvPath -NoTypeInformation -Encoding utf8 -ErrorAction Stop

####################################################################################
#################### CONSOLE OUTPUT #################################################
####################################################################################

$results |
    Format-Table DisplayName, LastSignIn, DeviceName, OperatingSystem -AutoSize |
    Out-Host

$failedCount = @(
    $results | Where-Object QueryStatus -eq "Failed"
).Count

$noSignInCount = @(
    $results | Where-Object QueryStatus -eq "No matching sign-in"
).Count

Write-Host "`nEnabled accounts reported: $($results.Count)" -ForegroundColor Green
Write-Host "No matching sign-in: $noSignInCount"
Write-Host "Failed queries: $failedCount"
Write-Host "Elapsed time: $([math]::Round($stopwatch.Elapsed.TotalSeconds, 1)) seconds"
Write-Host "CSV saved to: $csvPath" -ForegroundColor Green

if ($failedCount -gt 0) {
    Write-Warning "Some queries failed. Review the ErrorMessage column in the CSV."
}