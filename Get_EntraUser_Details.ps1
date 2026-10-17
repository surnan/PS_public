# Get_EntraUser_Details.ps1
# Input: User email address (primary email, alias, or UPN)
#
# Output:
#   Display Name
#   Username
#   Primary Email Address
#   Job Title
#   Creation Date/Time
#   All Email Aliases
#   Group Memberships:
#       Group Name
#       Group Type
#       Group Email
#       Membership (Direct/Nested)
#       Member Through (nested group path)
#
# Export:
#   CSV: Same Output


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
#################### GET INPUT PARAMETERS ##########################################
####################################################################################
$inputEmail = (Read-Host "Enter user's email address").Trim()

if ($inputEmail -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
    Write-Host "Invalid email format (1A)" -ForegroundColor Red
    return
}


####################################################################################
#################### FIND TARGET USER ##############################################
####################################################################################
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

try {
    $targetUser = @(
        Get-MgUser `
            -Filter "userPrincipalName eq '$inputEmail' or mail eq '$inputEmail'" `
            -Property Id, DisplayName, UserPrincipalName, Mail, JobTitle, CreatedDateTime, ProxyAddresses `
            -ErrorAction Stop
    )
}
catch {
    Write-Warning "Microsoft Graph error (2): $($_.Exception.Message)"
    return
}

if ($targetUser.Count -eq 0) {
    try {
        $targetUser = @(
            Get-MgUser `
                -Filter "proxyAddresses/any(address:address eq 'smtp:$inputEmail')" `
                -Property Id, DisplayName, UserPrincipalName, Mail, JobTitle, CreatedDateTime, ProxyAddresses `
                -ErrorAction Stop
        )
    }
    catch {
        Write-Warning "Microsoft Graph error (3): $($_.Exception.Message)"
        return
    }
}

if ($targetUser.Count -eq 0) {
    Write-Host "No user found with email address (4): $inputEmail" -ForegroundColor Red
    return
}

if ($targetUser.Count -gt 1) {
    Write-Warning "Multiple users matched this email address (5)"
    $targetUser |
    Select-Object DisplayName, UserPrincipalName, Mail |
    Format-Table -AutoSize |
    Out-Host
    return
}
# convert array into regular variable
$targetUser = $targetUser[0]


####################################################################################
#################### FIND DIRECT AND NESTED GROUP MEMBERSHIPS ######################
####################################################################################
try {
    $directGroups = @(
        Get-MgUserMemberOfAsGroup `
            -UserId $targetUser.Id `
            -Property Id `
            -All `
            -ErrorAction Stop
    )
}
catch {
    Write-Warning "Unable to retrieve direct group memberships (6): $($_.Exception.Message)"
    return
}


# Retrieve direct and nested group memberships
try {
    $allGroups = @(
        Get-MgUserTransitiveMemberOfAsGroup `
            -UserId $targetUser.Id `
            -Property Id, DisplayName, Mail, GroupTypes, MailEnabled, SecurityEnabled `
            -All `
            -ErrorAction Stop
    )
}
catch {
    Write-Warning "Unable to retrieve transitive group memberships (6A): $($_.Exception.Message)"
    return
}


####################################################################################
#################### RESOLVE NESTED GROUP PATHS ####################################
####################################################################################
$directGroupIds = @($directGroups | ForEach-Object { $_.Id })

# Array of all the User's group's ID
$groupById = @{}
foreach ($group in $allGroups) {
    $groupById[$group.Id] = $group
}

$groupPaths = @{}
$parentCache = @{}
$queue = [System.Collections.Generic.Queue[object]]::new()

# create queue with with each group ID and grow with each connected group
foreach ($group in $directGroups) {
    $queue.Enqueue(@($group.Id))
}

try {
    while ($queue.Count -gt 0) {
        $path = @($queue.Dequeue())
        $currentId = $path[-1]
        if (-not $parentCache.ContainsKey($currentId)) {
            $parentCache[$currentId] = @(
                Get-MgGroupMemberOfAsGroup `
                    -GroupId $currentId `
                    -Property Id `
                    -All `
                    -ErrorAction Stop
            )
        }

        foreach ($parent in $parentCache[$currentId]) {
            $parentId = $parent.Id
            if ($parentId -in $path -or -not $groupById.ContainsKey($parentId)) {
                continue
            }

            $newPath = @($path) + @($parentId)
            $through = @(
                $path | ForEach-Object {
                    if ($groupById.ContainsKey($_)) {
                        $groupById[$_].DisplayName
                    }
                }
            ) -join " -> "

            if (-not $groupPaths.ContainsKey($parentId)) {
                $groupPaths[$parentId] = @()
            }

            if ($through -notin $groupPaths[$parentId]) {
                $groupPaths[$parentId] += $through
            }
            $queue.Enqueue($newPath)
        }
    }
}
catch {
    Write-Host "Unable to resolve nested group paths (7): $($_.Exception.Message)" -ForegroundColor Red
    return
}


####################################################################################
#################### BUILD USER DETAILS OUTPUT OBJECT ##############################
####################################################################################
$tuDetails = [PSCustomObject]@{
    Name            = $targetUser.DisplayName
    UPN             = $targetUser.UserPrincipalName
    Email           = $targetUser.Mail
    JobTitle        = $targetUser.JobTitle
    CreatedDateTime = $targetUser.CreatedDateTime

    Aliases         = @(
        $targetUser.ProxyAddresses |
        Where-Object { $_ -cmatch '^smtp:' } |
        ForEach-Object { $_ -creplace '^smtp:', '' } |
        Sort-Object
    )

    AllGroups       = @(
        $allGroups | ForEach-Object {
            $group = $_

            $groupType = if ($group.GroupTypes -contains "Unified") {
                "Microsoft 365"
            }
            elseif ($group.MailEnabled -and $group.SecurityEnabled) {
                "Mail-enabled Security"
            }
            elseif ($group.MailEnabled) {
                "Distribution"
            }
            elseif ($group.SecurityEnabled) {
                "Security"
            }
            else {
                "Unknown"
            }

            $membership = if ($group.Id -in $directGroupIds) {
                "Direct"
            }
            else {
                "Nested"
            }

            $memberThrough = if ($membership -eq "Direct") {
                ""
            }
            elseif ($groupPaths.ContainsKey($group.Id)) {
                $groupPaths[$group.Id] -join " | "
            }
            else {
                "Path not resolved"
            }

            [PSCustomObject]@{
                "Group Name"     = $group.DisplayName
                "Group Type"     = $groupType
                "Group Email"    = $group.Mail
                "Membership"     = $membership
                "Member Through" = $memberThrough
            }
        }
    )
}


####################################################################################
#################### CONSOLE OUTPUT ################################################
####################################################################################
Write-Host ("`n{0,-20}" -f "Name:") -ForegroundColor Green -NoNewline
Write-Host $tuDetails.Name

Write-Host ("{0,-20}" -f "UPN:") -ForegroundColor Green -NoNewline
Write-Host $tuDetails.UPN

Write-Host ("{0,-20}" -f "Primary Email:") -ForegroundColor Green -NoNewline
Write-Host $tuDetails.Email

Write-Host ("{0,-20}" -f "Job Title:") -ForegroundColor Green -NoNewline
Write-Host $tuDetails.JobTitle

Write-Host ("{0,-20}" -f "Creation Date/Time:") -ForegroundColor Green -NoNewline
Write-Host $tuDetails.CreatedDateTime

Write-Host "`nEmail Aliases:" -ForegroundColor Cyan

if ($tuDetails.Aliases.Count -gt 0) {
    $tuDetails.Aliases | Out-Host
}
else {
    Write-Host "No email aliases found." -ForegroundColor Yellow
}

Write-Host "`n{Group Memberships}" -ForegroundColor Cyan

if ($tuDetails.AllGroups.Count -gt 0) {
    $tuDetails.AllGroups |
    Sort-Object Membership, "Group Name" |
    Format-Table `
    @{Label = "Group Name"; Expression = { $_.'Group Name' }; Width = 32 },
    @{Label = "Group Type"; Expression = { $_.'Group Type' }; Width = 24 },
    @{Label = "Membership"; Expression = { $_.Membership }; Width = 10 },
    @{Label = "Member Through"; Expression = { $_.'Member Through' }; Width = 55 } `
        -Wrap:$false |
    Out-Host
}
else {
    Write-Host "No group memberships found." -ForegroundColor Yellow
}

$nestedCount = @(
    $tuDetails.AllGroups | Where-Object Membership -eq "Nested"
).Count

Write-Host "`nTotal groups: $($tuDetails.AllGroups.Count)" -ForegroundColor Yellow
Write-Host "Direct groups: $($directGroups.Count)" -ForegroundColor Yellow
Write-Host "Nested-only groups: $nestedCount" -ForegroundColor Yellow


####################################################################################
#################### CREATE CSV ####################################################
####################################################################################
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"

$filename = $targetUser.UserPrincipalName.Split("@")[0]
$filename = $filename -replace '[^\w\.-]', '_'

$csvPath = ".\UserDetails_${filename}_$timestamp.csv"

try {
    @(
        "Name,$($tuDetails.Name)"
        "UPN,$($tuDetails.UPN)"
        "Primary Email,$($tuDetails.Email)"
        "Job Title,$($tuDetails.JobTitle)"
        "Creation Date/Time,$($tuDetails.CreatedDateTime)"
        ""
        "Email Aliases"
    ) | Set-Content -Path $csvPath -Encoding utf8 -ErrorAction Stop

    if ($tuDetails.Aliases.Count -gt 0) {
        $tuDetails.Aliases |
        Add-Content -Path $csvPath -Encoding utf8 -ErrorAction Stop
    }
    else {
        "No email aliases found." |
        Add-Content -Path $csvPath -Encoding utf8 -ErrorAction Stop
    }

    "" | Add-Content -Path $csvPath -Encoding utf8 -ErrorAction Stop
    "Group Memberships" | Add-Content -Path $csvPath -Encoding utf8 -ErrorAction Stop

    if ($tuDetails.AllGroups.Count -gt 0) {
        $tuDetails.AllGroups |
        Sort-Object "Group Name" |
        ConvertTo-Csv -NoTypeInformation |
        Add-Content -Path $csvPath -Encoding utf8 -ErrorAction Stop
    }
    else {
        "Group Name,Group Type,Group Email,Membership,Member Through" |
        Add-Content -Path $csvPath -Encoding utf8 -ErrorAction Stop
    }
}
catch {
    Write-Host "Unable to create CSV file (8): $($_.Exception.Message)" -ForegroundColor Red
    return
}


####################################################################################
#################### EXECUTION TIME ################################################
####################################################################################
$stopwatch.Stop()

Write-Host ("`n{0,-25}" -f "File Name:") -ForegroundColor DarkCyan -NoNewline
Write-Host "Get_EntraUser_Details.ps1" -ForegroundColor Cyan

Write-Host ("{0,-25}" -f "CSV file generated:") -ForegroundColor DarkYellow -NoNewline
Write-Host (Resolve-Path $csvPath) -ForegroundColor DarkYellow

Write-Host ("{0,-25}" -f "Total Execution Time:") -ForegroundColor DarkGreen -NoNewline
Write-Host "($($stopwatch.Elapsed))" -ForegroundColor DarkGreen
