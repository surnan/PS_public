# Get_EntraUser_LastSignIn_Device.ps1
# Reports the latest successful interactive Entra sign-in for every enabled account,
# including enabled guest accounts, within the available sign-in log retention window.
# Read-only; PowerShell 7 required.
# CSV: Timestamped and stored beside the script.

####################################################################################
#################### REQUIREMENTS / GRAPH CONNECTION ###############################
####################################################################################
if ($PSVersionTable.PSVersion.Major -lt 7) {
    throw 'Run this script in PowerShell 7.'
}

$modules = @('Microsoft.Graph.Users', 'Microsoft.Graph.Reports')
try {
    foreach ($module in $modules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            Write-Host "Installing missing module: $module" -ForegroundColor Yellow
            Install-Module -Name $module -Scope CurrentUser -ErrorAction Stop
        }
        Import-Module -Name $module -ErrorAction Stop
    }

    $requiredScopes = @('User.Read.All', 'AuditLog.Read.All')
    $context = Get-MgContext
    $missingScopes = @($requiredScopes | Where-Object { $_ -notin @($context.Scopes) })
    if (-not $context -or $context.AuthType -ne 'Delegated' -or $missingScopes.Count -gt 0) {
        Connect-MgGraph -Scopes $requiredScopes -ContextScope Process -NoWelcome -ErrorAction Stop
    }
    $context = Get-MgContext
    $missingScopes = @($requiredScopes | Where-Object { $_ -notin @($context.Scopes) })
    if (-not $context -or $context.AuthType -ne 'Delegated' -or $missingScopes.Count -gt 0) {
        throw "Missing delegated Graph permissions: $($missingScopes -join ', ')"
    }
}
catch {
    Write-Host "Microsoft Graph initialization failed: $($_.Exception.Message)" -ForegroundColor Red
    return
}

####################################################################################
#################### INPUT ##########################################################
####################################################################################
$inputDays = Read-Host 'Number of days to search [Enter = 30]'
if ([string]::IsNullOrWhiteSpace($inputDays)) {
    $days = 30
}
else {
    $days = 0
    if (-not [int]::TryParse($inputDays, [ref]$days) -or $days -lt 1) {
        Write-Host 'Enter a positive whole number of days.' -ForegroundColor Red
        return
    }
}
$startDateFilter = (Get-Date).ToUniversalTime().AddDays(-$days).ToString('yyyy-MM-ddTHH:mm:ssZ')
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

####################################################################################
#################### GET ENABLED USERS ##############################################
####################################################################################
Write-Host "`nGetting enabled accounts..." -ForegroundColor Cyan
try {
    $users = @(
        Get-MgUser -All -Filter 'accountEnabled eq true' `
            -Property Id, DisplayName, UserPrincipalName, AccountEnabled, UserType `
            -ErrorAction Stop | Sort-Object DisplayName
    )
}
catch {
    Write-Host "Unable to retrieve enabled accounts: $($_.Exception.Message)" -ForegroundColor Red
    return
}
if ($users.Count -eq 0) {
    Write-Host 'No enabled accounts found.' -ForegroundColor Yellow
    return
}
Write-Host "Enabled accounts: $($users.Count)" -ForegroundColor Cyan

####################################################################################
#################### OLD: PER-USER PARALLEL QUERIES (DISABLED) #######################
####################################################################################
# Previously queried Get-MgAuditLogSignIn once per user inside ForEach-Object -Parallel.
# This could generate hundreds of requests and trigger Graph HTTP 429 throttling.
# $throttleLimit = 4
# $results = @(
#     $users | ForEach-Object -Parallel {
#         $user = $_
#         Import-Module Microsoft.Graph.Reports -ErrorAction Stop
#         $filter = "userId eq '$($user.Id)' and createdDateTime ge $using:startDateFilter " +
#                   'and status/errorCode eq 0 and isInteractive eq true'
#         $signIn = Get-MgAuditLogSignIn -Filter $filter -Top 1 -ErrorAction Stop |
#             Select-Object -First 1
#         # Former per-user row processing occurred here.
#     } -ThrottleLimit $throttleLimit
# )

####################################################################################
#################### NEW: RETRIEVE SIGN-IN LOGS IN BULK #############################
####################################################################################
Write-Host "`nRetrieving interactive sign-ins for the last $days days..." -ForegroundColor Cyan
Write-Host 'Microsoft Graph will page through matching events; large tenants may take time.'
$filter = "createdDateTime ge $startDateFilter and isInteractive eq true"
try {
    $signIns = @(
        Get-MgAuditLogSignIn -Filter $filter -All -PageSize 999 -ErrorAction Stop
    )
}
catch {
    Write-Host "Unable to retrieve sign-ins: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host 'No partial report will be exported.' -ForegroundColor Yellow
    return
}
Write-Host "Interactive sign-in events downloaded: $($signIns.Count)"

####################################################################################
#################### KEEP LATEST SUCCESSFUL SIGN-IN PER USER ########################
####################################################################################
$latestByUser = @{}
foreach ($signIn in $signIns) {
    if ($null -eq $signIn.Status -or $signIn.Status.ErrorCode -ne 0 -or
        [string]::IsNullOrWhiteSpace([string]$signIn.UserId)) {
        continue
    }
    $id = [string]$signIn.UserId
    if (-not $latestByUser.ContainsKey($id) -or
        $signIn.CreatedDateTime -gt $latestByUser[$id].CreatedDateTime) {
        $latestByUser[$id] = $signIn
    }
}
Write-Host "Unique accounts with successful interactive sign-ins: $($latestByUser.Count)"

####################################################################################
#################### BUILD REPORT (LOCAL MATCHING, NO PER-USER API CALLS) ###########
####################################################################################
$results = [System.Collections.Generic.List[object]]::new()
foreach ($user in $users) {
    $signIn = $latestByUser[[string]$user.Id]
    $row = [PSCustomObject]@{
        DisplayName       = $user.DisplayName
        UserPrincipalName = $user.UserPrincipalName
        AccountEnabled    = $user.AccountEnabled
        UserType          = $user.UserType
        LastSignIn        = $null
        DeviceName        = 'No successful sign-in in available logs'
        OperatingSystem   = $null
        Browser           = $null
        DeviceId          = $null
        Managed           = $null
        Compliant         = $null
        IPAddress         = $null
        Application       = $null
        QueryStatus       = 'No matching sign-in'
        ErrorMessage      = $null
    }
    if ($null -ne $signIn) {
        $row.LastSignIn = ([datetimeoffset]$signIn.CreatedDateTime).ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss zzz')
        $row.DeviceName = $signIn.DeviceDetail.DisplayName
        if ([string]::IsNullOrWhiteSpace($row.DeviceName)) {
            $row.DeviceName = 'Device name not provided'
        }
        $row.OperatingSystem = $signIn.DeviceDetail.OperatingSystem
        $row.Browser         = $signIn.DeviceDetail.Browser
        $row.DeviceId        = $signIn.DeviceDetail.DeviceId
        $row.Managed         = $signIn.DeviceDetail.IsManaged
        $row.Compliant       = $signIn.DeviceDetail.IsCompliant
        $row.IPAddress       = $signIn.IpAddress
        $row.Application     = $signIn.AppDisplayName
        $row.QueryStatus     = 'Success'
    }
    $results.Add($row)
}

# Keep CSV alphabetical (unchanged); console can use another sort.
$results = @($results | Sort-Object DisplayName)

####################################################################################
#################### CSV EXPORT #####################################################
####################################################################################
$outputFolder = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$csvPath = Join-Path $outputFolder "EnabledUsers_LastSignInDevice_$timestamp.csv"
try {
    $results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding utf8 -ErrorAction Stop
}
catch {
    Write-Host "CSV export failed: $($_.Exception.Message)" -ForegroundColor Red
    return
}

####################################################################################
#################### CONSOLE OUTPUT #################################################
####################################################################################
# OLD: alphabetical, full-width names (DISABLED)
# $results | Format-Table DisplayName, LastSignIn, DeviceName, OperatingSystem -AutoSize | Out-Host

# NEW: newest sign-in first; truncate display name to 20 chars only in console.
$results |
    Sort-Object @{ Expression = { [string]::IsNullOrWhiteSpace($_.LastSignIn) } }, `
                @{ Expression = { $_.LastSignIn }; Descending = $true } |
    Select-Object `
        @{Name = 'DisplayName'; Expression = {
            $name = [string]$_.DisplayName
            if ($name.Length -gt 20) { $name.Substring(0, 17) + '...' } else { $name }
        }},
        LastSignIn, DeviceName, OperatingSystem |
    Format-Table -AutoSize | Out-Host

$noSignInCount = @($results | Where-Object QueryStatus -eq 'No matching sign-in').Count
$successCount = @($results | Where-Object QueryStatus -eq 'Success').Count
$stopwatch.Stop()
Write-Host "`nEnabled accounts reported: $($results.Count)" -ForegroundColor Green
Write-Host "Accounts with matching sign-in: $successCount"
Write-Host "No matching sign-in: $noSignInCount"
Write-Host 'Per-user Graph queries: 0 (bulk retrieval used)'
Write-Host "Elapsed time: $([math]::Round($stopwatch.Elapsed.TotalSeconds, 1)) seconds"
Write-Host "CSV saved to: $csvPath" -ForegroundColor Green
Write-Host 'Note: Results are limited to sign-in events retained and accessible in Microsoft Entra.' -ForegroundColor DarkYellow
