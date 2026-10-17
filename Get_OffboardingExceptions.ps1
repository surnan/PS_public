# Get_OffboardingExceptions.ps1
# Input: None
#
# Output: All Disabled Accounts
#   DisplayName
#   User Principal Name
#   Email Address Count
#   Group Membership Count
#   License Count
#   Account Enabled Status
#
# Export: CSV with detailed information


####################################################################################
#################### CONFIGURATION #################################################
####################################################################################
$throttleLimit = 5
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# ForEach-Object -Parallel requires PowerShell 7
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "PowerShell 7 or later is required (1)." -ForegroundColor Red
    return
}



####################################################################################
#################### MICROSOFT GRAPH CONNECTION ####################################
####################################################################################
$modules = @(
    "Microsoft.Graph.Users"
)

$requiredScopes = @("User.Read.All", "Group.Read.All")

try {
    # Install and import required modules
    foreach ($module in $modules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            Install-Module $module -Scope CurrentUser -ErrorAction Stop
        }

        Import-Module $module -ErrorAction Stop
    }

    # Check existing connection and permissions
    $connection = Get-MgContext

    $missingScopes = @(
        $requiredScopes | Where-Object {
            $_ -notin @($connection.Scopes)
        }
    )

    # Connect if disconnected or missing required permissions
    if (-not $connection -or
        $connection.AuthType -ne "Delegated" -or
        $missingScopes.Count -gt 0) {

        Connect-MgGraph `
            -Scopes $requiredScopes `
            -ContextScope CurrentUser `
            -NoWelcome `
            -ErrorAction Stop

        $connection = Get-MgContext
    }

    # Verify authentication type
    if (-not $connection -or $connection.AuthType -ne "Delegated") {
        throw "A delegated Microsoft Graph connection is required."
    }

    # Verify required permissions
    $missingScopes = @(
        $requiredScopes | Where-Object {
            $_ -notin @($connection.Scopes)
        }
    )

    if ($missingScopes.Count -gt 0) {
        throw "Missing Graph scopes: $($missingScopes -join ', ')"
    }
}
catch {
    Write-Host "Microsoft Graph connection failed:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    return
}



####################################################################################
#################### FIND ALL DISABLED ACCOUNTS ####################################
####################################################################################
Write-Host "`nLoading disabled accounts..." -ForegroundColor Yellow

try {
    $disabledUsers = @(
        Get-MgUser `
            -Filter "accountEnabled eq false" `
            -Property Id, DisplayName, UserPrincipalName, AccountEnabled, ProxyAddresses, AssignedLicenses `
            -All `
            -ErrorAction Stop
    )
}
catch {
    Write-Host "Unable to retrieve disabled accounts (3):" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    return
}

if ($disabledUsers.Count -eq 0) {
    Write-Host "No disabled accounts were found (4)." -ForegroundColor Yellow
    $stopwatch.Stop()
    return
}

Write-Host "Disabled accounts found: $($disabledUsers.Count)" -ForegroundColor Cyan
Write-Host "Parallel workers: $throttleLimit" -ForegroundColor Cyan


####################################################################################
#################### PARALLEL PROCESSING ###########################################
####################################################################################
$processed = 0

$parallelResults = @(
    $disabledUsers |
    ForEach-Object -Parallel {

        $user = $_

        try {
            # Each parallel runspace needs its own module import
            Import-Module Microsoft.Graph.Users -ErrorAction Stop

            # Graph connections are not guaranteed to carry into parallel runspaces
            if (-not (Get-MgContext)) {
                Connect-MgGraph `
                    -Scopes "User.Read.All", "Group.Read.All" `
                    -ContextScope CurrentUser `
                    -NoWelcome `
                    -ErrorAction Stop | Out-Null
            }

            # Get all SMTP addresses
            $smtpAddresses = @(
                $user.ProxyAddresses |
                Where-Object { $_ -like "smtp:*" } |
                ForEach-Object { $_ -replace "^smtp:", "" }
            )

            # Get direct group memberships
            $groups = @(
                Get-MgUserMemberOfAsGroup `
                    -UserId $user.Id `
                    -All `
                    -ErrorAction Stop
            )

            # CSV display values
            $aliasDisplay = if ($smtpAddresses.Count -eq 1) {
                $smtpAddresses[0]
            }
            else {
                $smtpAddresses.Count
            }

            $groupDisplay = if ($groups.Count -eq 1) {
                $groups[0].DisplayName
            }
            else {
                $groups.Count
            }

            $licenseCount = @($user.AssignedLicenses).Count

            # Return one complete result per user
            [PSCustomObject]@{
                DisplayName        = $user.DisplayName
                UserPrincipalName  = $user.UserPrincipalName
                AccountEnabled     = $user.AccountEnabled
                EmailAddressCount  = $smtpAddresses.Count
                GroupCount         = $groups.Count
                LicenseCount       = $licenseCount
                SMTPAddresses      = $aliasDisplay
                Groups             = $groupDisplay
                Error              = $null
            }
        }
        catch {
            [PSCustomObject]@{
                DisplayName        = $user.DisplayName
                UserPrincipalName  = $user.UserPrincipalName
                Error              = $_.Exception.Message
            }
        }

    } -ThrottleLimit $throttleLimit |
    ForEach-Object {
        $processed++

        Write-Progress `
            -Activity "Checking disabled accounts" `
            -Status "$processed of $($disabledUsers.Count) completed" `
            -PercentComplete (($processed / $disabledUsers.Count) * 100)

        $_
    }
)

Write-Progress -Activity "Checking disabled accounts" -Completed


####################################################################################
#################### CHECK FOR ERRORS ##############################################
####################################################################################
$scanErrors = @($parallelResults | Where-Object { $_.Error })

if ($scanErrors.Count -gt 0) {
    Write-Host "`nErrors encountered: $($scanErrors.Count)" -ForegroundColor Red

    $scanErrors |
    Format-Table DisplayName, UserPrincipalName, Error -Wrap

    Write-Host "CSV export skipped to avoid an incomplete report." -ForegroundColor Yellow

    $stopwatch.Stop()
    Write-Host "Execution time (5): $($stopwatch.Elapsed)" -ForegroundColor Cyan
    return
}


####################################################################################
#################### PREPARE REPORTS ###############################################
####################################################################################
$results = @($parallelResults | Sort-Object DisplayName)

$consoleResults = @(
    $results |
    Select-Object `
        DisplayName,
        UserPrincipalName,
        EmailAddressCount,
        GroupCount,
        LicenseCount,
        AccountEnabled
)


####################################################################################
#################### CONSOLE OUTPUT ################################################
####################################################################################
Write-Host "`nTotal disabled accounts: $($results.Count)" -ForegroundColor Yellow
Write-Host ""

$consoleResults |
Format-Table `
    DisplayName, `
    UserPrincipalName, `
    @{Name = "[SMTP Count]"; Expression = { $_.EmailAddressCount }; Alignment = "Center"}, `
    @{Name = "[Groups]"; Expression = { $_.GroupCount }; Alignment = "Center"}, `
    @{Name = "[Licenses]"; Expression = { $_.LicenseCount }; Alignment = "Center"}, `
    AccountEnabled `
    -AutoSize


####################################################################################
#################### CREATE CSV ####################################################
####################################################################################
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$csvPath = ".\OffboardingExceptions_$timestamp.csv"

try {
    $results |
    Select-Object `
        DisplayName,
        UserPrincipalName,
        AccountEnabled,
        SMTPAddresses,
        Groups,
        LicenseCount |
    Export-Csv `
        -Path $csvPath `
        -NoTypeInformation `
        -Encoding utf8 `
        -ErrorAction Stop
}
catch {
    Write-Host "CSV export failed (6):" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    $stopwatch.Stop()
    return
}


####################################################################################
#################### EXECUTION TIME ################################################
####################################################################################
$stopwatch.Stop()

Write-Host ("`n{0,-25}" -f "File Name:") -ForegroundColor DarkCyan -NoNewline
Write-Host "Get_OffboardingExceptions.ps1" -ForegroundColor Cyan

Write-Host ("{0,-25}" -f "CSV file generated:") -ForegroundColor DarkYellow -NoNewline
Write-Host (Resolve-Path $csvPath) -ForegroundColor DarkYellow

Write-Host ("{0,-25}" -f "Total Execution Time:") -ForegroundColor DarkGreen -NoNewline
Write-Host "$($stopwatch.Elapsed)" -ForegroundColor DarkGreen
