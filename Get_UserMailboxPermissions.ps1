
# Get_UserMailboxPermissions.ps1
# Reports direct and security-group-based Full Access permissions.
# Uses 5 parallel session 
# Mailbox retry max = 2 to help prevent access errors caused by multiple processes on same mailbox

####################################################################################
#################### All Exchange Powershell Scripts START #########################
####################################################################################
# Check Exchange Online module
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Install-Module ExchangeOnlineManagement `
        -Scope CurrentUser `
        -ErrorAction Stop
}

# Module installed from if statement above
Import-Module ExchangeOnlineManagement -ErrorAction Stop

# Get-ConnectionInformation belongs to ExchangeOnlineManagement module
$connection = Get-ConnectionInformation -ErrorAction SilentlyContinue |
    Where-Object { $_.State -eq "Connected" }

if (-not $connection) {
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
}
####################################################################################
#################### All Exchange Powershell Scripts END ###########################
####################################################################################



## Convert input --> $user
$user_email = Read-Host "Enter the user's email address"
if ($user_email.Trim() -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
    Write-Host "Invalid email format." -ForegroundColor Red -BackgroundColor White
    exit
}

try {
    $user = Get-EXORecipient -Identity $user_email -ErrorAction Stop
}
catch {
    Write-Warning "Could not find user: $user_email"
    return
}

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# User properties push into each parallel session
# String Array of all user identifiers for future query inputs
# Unique prevents array from growing without new identifiers
$userIdentifiers = @(
    [string]$user.PrimarySmtpAddress    #otherwise, data-type: SmtpAddress
    [string]$user.Guid                  #otherwise, data-type: Guid
    [string]$user.DistinguishedName
    [string]$user.Name
    [string]$user.DisplayName
    [string]$user_email
) | Where-Object { $_ } | Select-Object -Unique

# Get-DistributionGroup = Only Distribution & mail-enabled security groups
# 'Get-MgGroup' = Entra ID + Microsoft 365 + mail-enabled groups
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
    Write-Warning "Could not retrieve security groups: $($_.Exception.Message)"
    return
}

$userGroups = @()   # Groups containing the user
$groupErrors = @()  # Array of Group iteration errors

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

# Retrieve all mailboxes except user's own
$allMailboxes = @(
    Get-EXOMailbox -ResultSize Unlimited -ErrorAction Stop |
        Where-Object {
            $_.PrimarySmtpAddress -ne $user.PrimarySmtpAddress
        }
)

Write-Host "`nMailbox count = $($allMailboxes.Count)" -ForegroundColor Magenta

# Display temporary progress indicator
Write-Progress `
    -Activity "Checking mailbox permissions" `
    -Status "Scanning $($allMailboxes.Count) mailboxes using 5 parallel workers..."

# Scan mailbox permissions in parallel
try {
    $results = @(
        $allMailboxes | ForEach-Object -Parallel {
            $currentMailbox = $_

            # Access variables from main runspace
            $userIds = $using:userIdentifiers
            $groups = $using:groupLookup

            # Two attempts = initial query + one retry
            # Retry mailbox queries from multi-thread conflict
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
                        # Hoping delay reduces chance this mailbox having another multi-thread conflict
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
    # Remove progress indicator
    Write-Progress -Activity "Checking mailbox permissions" -Completed
}


# Get successful Full Access matches
$fullAccess = @(
    $results |
        Where-Object { $_.Status -eq "FullAccess" } |
        Sort-Object Mailbox, Access
)

# Display Full Access results
Write-Host "`nMailbox Full Access permissions:" -ForegroundColor Cyan

if ($fullAccess.Count -gt 0) {
    $fullAccess |
        Select-Object Mailbox, Access, GrantedBy |
        Format-Table -AutoSize
}
else {
    Write-Host "No Full Access permissions found." -ForegroundColor Yellow
}

# Print out Errors
$errorsFound = @($results | Where-Object { $_.Status -eq "Error" })

foreach ($result in $errorsFound) {
    Write-Warning "Could not check $($result.Mailbox): $($result.Error)"
}

foreach ($groupError in $groupErrors) {
    Write-Warning "Could not check group $($groupError.Group): $($groupError.Error)"
}

$stopwatch.Stop()   # Stop timer

# Calculate summary totals
$directCount = @($fullAccess | Where-Object { $_.Access -eq "Direct" }).Count
$groupCount = @($fullAccess | Where-Object { $_.Access -eq "Security Group" }).Count

# Display summary
Write-Host "`nDirect Full Access grants: $directCount" -ForegroundColor Green
Write-Host "Security group Full Access grants: $groupCount" -ForegroundColor Green
Write-Host "Total Full Access grants: $($fullAccess.Count)" -ForegroundColor Magenta
Write-Host "Mailbox query errors: $($errorsFound.Count)" -ForegroundColor Yellow
Write-Host "Group lookup errors: $($groupErrors.Count)" -ForegroundColor Yellow
Write-Host "Total execution time: $($stopwatch.Elapsed)" -ForegroundColor Cyan