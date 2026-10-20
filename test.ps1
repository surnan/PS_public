# Get_EnabledUsersLastSignInDevice.ps1
#
# Reports the latest successful interactive sign-in for every enabled account.
# Includes enabled guest accounts.
# Read-only: Does not modify users or devices.

####################################################################################
#################### MICROSOFT GRAPH CONNECTION ####################################
####################################################################################

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

Connect-MgGraph -Scopes "User.Read.All", "AuditLog.Read.All" -NoWelcome `
    -ErrorAction Stop

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

$startDate = (Get-Date).ToUniversalTime().AddDays(-$days)
$startDateFilter = $startDate.ToString("yyyy-MM-ddTHH:mm:ssZ")

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
#################### GET LAST SIGN-IN DEVICE ########################################
####################################################################################

$counter = 0

$results = @(
    foreach ($user in $users) {
        $counter++

        Write-Progress -Activity "Checking latest successful sign-ins" `
            -Status "$counter of $($users.Count): $($user.DisplayName)" `
            -PercentComplete (($counter / $users.Count) * 100)

        $filter = "userId eq '$($user.Id)' " +
                  "and createdDateTime ge $startDateFilter " +
                  "and status/errorCode eq 0 " +
                  "and isInteractive eq true"

        # Graph returns the newest sign-in first by default.
        # Stop on query errors instead of reporting them as missing activity.
        $signIn = Get-MgAuditLogSignIn `
            -Filter $filter `
            -Top 1 `
            -ErrorAction Stop |
            Select-Object -First 1

        $deviceName = "No successful sign-in in available logs"
        $lastSignIn = $null

        if ($null -ne $signIn) {
            $lastSignIn = ([datetimeoffset]$signIn.CreatedDateTime).
                ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss zzz")

            $deviceName = $signIn.DeviceDetail.DisplayName

            if ([string]::IsNullOrWhiteSpace($deviceName)) {
                $deviceName = "Device name not provided"
            }
        }

        [PSCustomObject]@{
            DisplayName       = $user.DisplayName
            UserPrincipalName = $user.UserPrincipalName
            AccountEnabled    = $user.AccountEnabled
            UserType          = $user.UserType
            LastSignIn        = $lastSignIn
            DeviceName        = $deviceName
            OperatingSystem   = $signIn.DeviceDetail.OperatingSystem
            Browser           = $signIn.DeviceDetail.Browser
            DeviceId          = $signIn.DeviceDetail.DeviceId
            Managed           = $signIn.DeviceDetail.IsManaged
            Compliant         = $signIn.DeviceDetail.IsCompliant
            IPAddress         = $signIn.IpAddress
            Application       = $signIn.AppDisplayName
        }
    }
)

Write-Progress -Activity "Checking latest successful sign-ins" -Completed

####################################################################################
#################### CONSOLE OUTPUT AND CSV EXPORT ##################################
####################################################################################

$results |
    Format-Table DisplayName, LastSignIn, DeviceName, OperatingSystem -AutoSize |
    Out-Host

$outputFolder = if ($PSScriptRoot) {
    $PSScriptRoot
}
else {
    (Get-Location).Path
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$csvPath = Join-Path $outputFolder "EnabledUsers_LastSignInDevice_$timestamp.csv"

$results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding utf8

Write-Host "`nEnabled accounts reported: $($results.Count)" -ForegroundColor Green
Write-Host "CSV saved to: $csvPath" -ForegroundColor Green