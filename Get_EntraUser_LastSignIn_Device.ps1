
# EntraUserLastSignInDevice.ps1
# Input: Number of days to search (default: 30)
#
# Output:
#   Display Name
#   User Principal Name
#   Last Sign-in
#   Computer Name
#   Operating System
#   Browser
#   Device ID
#   Managed
#   Compliant
#   IP Address
#   Application
#
# Export:
#   CSV: Same output
#
# Read-only: Does not modify Entra ID or devices.


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

$requiredScopes = @(
    "User.Read.All",
    "AuditLog.Read.All"
)

$context = Get-MgContext

if (-not $context -or
    @($requiredScopes | Where-Object { $_ -notin $context.Scopes }).Count -gt 0) {

    Connect-MgGraph -Scopes $requiredScopes -NoWelcome -ErrorAction Stop
}


####################################################################################
#################### USER INPUT ####################################################
####################################################################################
$inputDays = Read-Host "Enter number of days to search (default: 30)"

if ([string]::IsNullOrWhiteSpace($inputDays)) {
    $days = 30
}
else {
    $days = 0

    if (-not [int]::TryParse($inputDays, [ref]$days) -or $days -lt 1) {
        Write-Host "Invalid number of days." -ForegroundColor Red
        return
    }
}

$startDate = (Get-Date).ToUniversalTime().AddDays(-$days)
$startDateString = $startDate.ToString("yyyy-MM-ddTHH:mm:ssZ")

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()


####################################################################################
#################### RETRIEVE ENABLED USERS ########################################
####################################################################################
Write-Host "`nRetrieving enabled users..." -ForegroundColor Cyan

try {
    $users = @(
        Get-MgUser `
            -All `
            -Filter "accountEnabled eq true and userType eq 'Member'" `
            -Property "Id,DisplayName,UserPrincipalName" `
            -ErrorAction Stop
    )
}
catch {
    Write-Host "Unable to retrieve users: $_" -ForegroundColor Red
    return
}

Write-Host "Enabled users found: $($users.Count)"


####################################################################################
#################### RETRIEVE SIGN-IN LOGS #########################################
####################################################################################
Write-Host "`nRetrieving sign-in logs for the last $days days..." -ForegroundColor Cyan

$filter = "createdDateTime ge $startDateString and isInteractive eq true"

try {
    $signIns = @(
        Get-MgAuditLogSignIn -Filter $filter -All -PageSize 999 -ErrorAction Stop
    )
}
catch {
    Write-Host "Unable to retrieve sign-in logs: $_" -ForegroundColor Red
    return
}

# Successful sign-ins only.
$successfulSignIns = @(
    $signIns | Where-Object {
        $_.Status.ErrorCode -eq 0 -and
        $null -ne $_.UserId
    }
)

Write-Host "Successful interactive sign-ins: $($successfulSignIns.Count)"


####################################################################################
#################### FIND LATEST SIGN-IN PER USER ##################################
####################################################################################
$latestByUser = @{}

foreach ($signIn in $successfulSignIns) {

    $userId = [string]$signIn.UserId

    if (-not $latestByUser.ContainsKey($userId) -or
        $signIn.CreatedDateTime -gt $latestByUser[$userId].CreatedDateTime) {

        $latestByUser[$userId] = $signIn
    }
}


####################################################################################
#################### BUILD REPORT ##################################################
####################################################################################
$results = [System.Collections.Generic.List[object]]::new()

$total = $users.Count
$current = 0

foreach ($user in $users) {

    $current++

    Write-Progress `
        -Activity "Building Entra sign-in device report" `
        -Status "$current of $total - $($user.DisplayName)" `
        -PercentComplete (($current / [math]::Max($total, 1)) * 100)

    $signIn = $latestByUser[[string]$user.Id]

    $lastSignIn = $null
    $computerName = "Not available"
    $operatingSystem = ""
    $browser = ""
    $deviceId = ""
    $managed = ""
    $compliant = ""
    $ipAddress = ""
    $application = ""

    if ($signIn) {

        $lastSignIn = $signIn.CreatedDateTime.ToLocalTime().ToString(
            "yyyy-MM-dd HH:mm:ss"
        )

        if ($signIn.DeviceDetail) {

            $device = $signIn.DeviceDetail

            if (-not [string]::IsNullOrWhiteSpace($device.DisplayName)) {
                $computerName = $device.DisplayName
            }

            $operatingSystem = $device.OperatingSystem
            $browser = $device.Browser
            $deviceId = $device.DeviceId

            if ($null -ne $device.IsManaged) {
                $managed = $device.IsManaged
            }

            if ($null -ne $device.IsCompliant) {
                $compliant = $device.IsCompliant
            }
        }

        $ipAddress = $signIn.IpAddress
        $application = $signIn.AppDisplayName
    }

    $results.Add([PSCustomObject]@{
        DisplayName       = $user.DisplayName
        UserPrincipalName = $user.UserPrincipalName
        LastSignIn        = $lastSignIn
        ComputerName      = $computerName
        OperatingSystem   = $operatingSystem
        Browser           = $browser
        DeviceId          = $deviceId
        Managed           = $managed
        Compliant         = $compliant
        IpAddress         = $ipAddress
        Application       = $application
    })
}

Write-Progress -Activity "Building Entra sign-in device report" -Completed


####################################################################################
#################### CONSOLE OUTPUT ################################################
####################################################################################
# $results = @($results | Sort-Object DisplayName)
$results = @($results | Sort-Object LastSignIn -Descending)


Write-Host "`n================ REPORT SUMMARY ================" -ForegroundColor Cyan

Write-Host "Enabled users: $($users.Count)"
Write-Host "Users with recent sign-ins: $($latestByUser.Count)"
Write-Host "Users with computer names: $( @($results | Where-Object {
    $_.ComputerName -ne 'Not available'
}).Count )"

Write-Host "`nLAST SIGN-IN DEVICES" -ForegroundColor Cyan

# $results |
#     Select-Object DisplayName, LastSignIn, ComputerName, OperatingSystem |
#     Format-Table -AutoSize
$results |
    Sort-Object LastSignIn -Descending |
    Select-Object `
        @{Name = 'DisplayName'; Expression = {
            if ($_.DisplayName.Length -gt 20) {
                $_.DisplayName.Substring(0, 17) + '...'
            }
            else {
                $_.DisplayName
            }
        }},
        LastSignIn,
        ComputerName,
        OperatingSystem |
    Format-Table -AutoSize

####################################################################################
#################### EXPORT CSV ####################################################
####################################################################################
$timestamp = Get-Date -Format "yyyy-MM-dd_HHmmss"

$outputPath = Join-Path (Get-Location).Path (
    "EntraUserLastSignInDevice_$timestamp.csv"
)

try {
    $results |
        Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop

    Write-Host "`nCSV exported: $outputPath" -ForegroundColor Green
}
catch {
    Write-Host "CSV export failed: $_" -ForegroundColor Red
}

$stopwatch.Stop()

Write-Host "Execution time: $($stopwatch.Elapsed.ToString('hh\:mm\:ss'))"
