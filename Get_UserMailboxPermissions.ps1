
# Get_UserMailboxPermissions.ps1
# Reports direct and security-group-based Full Access permissions.
# Uses 5 parallel sessions
# Mailbox retry max = 2 to help prevent access errors caused by
# multiple processes querying the same mailbox



####################################################################################
#################### EXCHANGE ONLINE CONNECTION ####################################
####################################################################################
$module = "ExchangeOnlineManagement"

try {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Install-Module $module -Scope CurrentUser -ErrorAction Stop
    }

    if (-not (Get-Module -Name $module)) {
        Import-Module $module -ErrorAction Stop
    }

    $connection = Get-ConnectionInformation -ErrorAction Stop |
        Where-Object { $_.State -eq "Connected" }

    if (-not $connection) {
        Connect-ExchangeOnline `
            -ShowBanner:$false `
            -ErrorAction Stop
    }
}
catch {
    Write-Host "Exchange Online connection failed (1):" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    return
}



####################################################################################
#################### GET INPUT PARAMETERS ##########################################
####################################################################################
$user_email = (Read-Host "Enter the user's email address").Trim()

if ($user_email -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
    Write-Host "Invalid email format (2)" -ForegroundColor Red -BackgroundColor White
    return
}


####################################################################################
#################### FIND TARGET USER ##############################################
####################################################################################
try {
    $user = Get-EXORecipient `
        -Identity $user_email `
        -ErrorAction Stop
}
catch {
    Write-Warning "Could not find user (3): $user_email"
    return
}

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()


####################################################################################
#################### BUILD USER IDENTIFIERS ########################################
####################################################################################
# User properties pushed into each parallel session
# Unique prevents duplicate identifiers

$userIdentifiers = @(
    [string]$user.PrimarySmtpAddress
    [string]$user.Guid
    [string]$user.DistinguishedName
    [string]$user.Name
    [string]$user.DisplayName
    [string]$user_email
) | Where-Object { $_ } | Select-Object -Unique


####################################################################################
#################### FIND SECURITY GROUP MEMBERSHIPS ###############################
####################################################################################
# Get-DistributionGroup returns distribution and mail-enabled security groups

Write-Host "`nChecking security group memberships..." -ForegroundColor Yellow

try {
    $allSecurityGroups = @(
        Get-DistributionGroup `
            -RecipientTypeDetails MailUniversalSecurityGroup `
            -ResultSize Unlimited `
            -ErrorAction Stop
    )
}
catch {
    Write-Warning "Could not retrieve security groups (4): $($_.Exception.Message)"
    return
}

$userGroups = @()
$groupErrors = @()

foreach ($currentGroup in $allSecurityGroups) {
    try {
        $members = @(
            Get-DistributionGroupMember `
                -Identity $currentGroup.Identity `
                -ResultSize Unlimited `
                -ErrorAction Stop
        )

        $isMember = $members | Where-Object {
            $_.Guid -eq $user.Guid -or
            $_.PrimarySmtpAddress -eq $user.PrimarySmtpAddress
        }

        if ($isMember) {
            $userGroups += $currentGroup

            Write-Host "Security group: $($currentGroup.DisplayName)" `
                -ForegroundColor Green
        }
    }
    catch {
        $groupErrors += [PSCustomObject]@{
            Group = $currentGroup.DisplayName
            Error = $_.Exception.Message
        }
    }
}

Write-Host "`nSecurity groups found: $($userGroups.Count)" `
    -ForegroundColor Magenta


####################################################################################
#################### BUILD SECURITY GROUP LOOKUP ###################################
####################################################################################
# Build lookup table for user's security groups
# Pushed into each parallel session

$groupLookup = @{}

foreach ($currentGroup in $userGroups) {
    $identifiers = @(
        [string]$currentGroup.PrimarySmtpAddress
        [string]$currentGroup.Guid
        [string]$currentGroup.DistinguishedName
        [string]$currentGroup.Name
        [string]$currentGroup.DisplayName
        [string]$currentGroup.Alias
    )

    foreach ($identifier in $identifiers) {
        if ($identifier) {
            $groupLookup[$identifier] = $currentGroup.DisplayName
        }
    }
}


####################################################################################
#################### RETRIEVE ALL MAILBOXES ########################################
####################################################################################
# Retrieve all mailboxes except user's own

$allMailboxes = @(
    Get-EXOMailbox -ResultSize Unlimited -ErrorAction Stop |
        Where-Object {
            $_.PrimarySmtpAddress -ne $user.PrimarySmtpAddress
        }
)

Write-Host "`nMailbox count = $($allMailboxes.Count)" -ForegroundColor Magenta

Write-Progress `
    -Activity "Checking mailbox permissions" `
    -Status "Scanning $($allMailboxes.Count) mailboxes using 5 parallel workers..."


####################################################################################
#################### SCAN MAILBOX PERMISSIONS ######################################
####################################################################################
try {
    $results = @(
        $allMailboxes | ForEach-Object -Parallel {
            $currentMailbox = $_

            # Access variables from main runspace
            $userIds = $using:userIdentifiers
            $groups = $using:groupLookup

            # Two attempts = initial query + one retry
            $maxAttempts = 2

            for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
                try {
                    # Retrieve mailbox permission entries
                    $permissions = @(
                        Get-EXOMailboxPermission `
                            -Identity $currentMailbox.PrimarySmtpAddress `
                            -ErrorAction Stop
                    )

                    foreach ($permission in $permissions) {
                        # Only check explicitly allowed Full Access
                        if ($permission.AccessRights -notcontains "FullAccess" -or
                            $permission.Deny -or
                            $permission.IsInherited) {
                            continue
                        }

                        $permissionUser = [string]$permission.User

                        # User has Direct Full Access
                        if ($userIds -contains $permissionUser) {
                            [PSCustomObject]@{
                                Mailbox   = [string]$currentMailbox.PrimarySmtpAddress
                                Access    = "Direct"
                                GrantedBy = "User account"
                                Status    = "FullAccess"
                                Error     = ""
                            }
                        }
                        elseif ($groups.ContainsKey($permissionUser)) {
                            # Full Access through security groups
                            [PSCustomObject]@{
                                Mailbox   = [string]$currentMailbox.PrimarySmtpAddress
                                Access    = "Security Group"
                                GrantedBy = $groups[$permissionUser]
                                Status    = "FullAccess"
                                Error     = ""
                            }
                        }
                    }

                    # Query succeeded; exit retry loop
                    break
                }
                catch {
                    if ($attempt -lt $maxAttempts) {
                        # Wait two seconds before retrying
                        Start-Sleep -Seconds 2
                    }
                    else {
                        # Both attempts failed
                        [PSCustomObject]@{
                            Mailbox   = [string]$currentMailbox.PrimarySmtpAddress
                            Access    = ""
                            GrantedBy = ""
                            Status    = "Error"
                            Error     = $_.Exception.Message
                        }
                    }
                }
            }
        } -ThrottleLimit 5
    )
}
finally {
    Write-Progress -Activity "Checking mailbox permissions" -Completed
}


####################################################################################
#################### FILTER FULL ACCESS RESULTS ####################################
####################################################################################
$fullAccess = @(
    $results |
        Where-Object { $_.Status -eq "FullAccess" } |
        Sort-Object Mailbox, Access
)


####################################################################################
#################### CONSOLE OUTPUT ################################################
####################################################################################
Write-Host "`nMailbox Full Access permissions:" -ForegroundColor Cyan

if ($fullAccess.Count -gt 0) {
    $fullAccess |
        Select-Object Mailbox, Access, GrantedBy |
        Format-Table -AutoSize
}
else {
    Write-Host "No Full Access permissions found." -ForegroundColor Yellow
}


####################################################################################
#################### DISPLAY ERRORS ################################################
####################################################################################
$errorsFound = @($results | Where-Object { $_.Status -eq "Error" })

foreach ($result in $errorsFound) {
    Write-Warning "Could not check $($result.Mailbox): $($result.Error)"
}

foreach ($groupError in $groupErrors) {
    Write-Warning "Could not check group $($groupError.Group): $($groupError.Error)"
}


####################################################################################
#################### EXECUTION TIME ################################################
####################################################################################
$stopwatch.Stop()

$directCount = @($fullAccess | Where-Object { $_.Access -eq "Direct" }).Count
$groupCount = @($fullAccess | Where-Object { $_.Access -eq "Security Group" }).Count

Write-Host "`nDirect Full Access grants: $directCount" -ForegroundColor Green
Write-Host "Security group Full Access grants: $groupCount" -ForegroundColor Green
Write-Host "Total Full Access grants: $($fullAccess.Count)" -ForegroundColor Magenta
Write-Host "Mailbox query errors: $($errorsFound.Count)" -ForegroundColor Yellow
Write-Host "Group lookup errors: $($groupErrors.Count)" -ForegroundColor Yellow
Write-Host "Total execution time: $($stopwatch.Elapsed)" -ForegroundColor Cyan
