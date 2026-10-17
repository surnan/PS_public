
# Get_EntraGroup_Members.ps1
# Input: Group email address (primary email or alias)
#
# Output: $tgDetails
#   Group Name
#   Group Email
#   Group Type
#   All Email Aliases
#   Group Members (Name, UPN, Email)
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
$inputEmail = (Read-Host "Enter group's email address").Trim()
if ($inputEmail -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
    Write-Host "Invalid email format (1A)" -ForegroundColor Red
    return
}


####################################################################################
#################### FIND TARGET GROUP #############################################
####################################################################################
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

try {
    $targetGroup = Get-MgGroup `
        -Filter "mail eq '$inputEmail'" `
        -Property Id, DisplayName, Mail, ProxyAddresses, GroupTypes, MailEnabled, SecurityEnabled `
        -ErrorAction Stop
}
catch {
    Write-Warning "Microsoft Graph error (2): $($_.Exception.Message)"
    return
}

# Check alias emails if targetGroup not found
if (-not $targetGroup) {
    Write-Host "Not found as primary email. Checking aliases..." -ForegroundColor Yellow

    try {
        # SMTP = primary; smtp = alias
        $targetGroup = Get-MgGroup `
            -Filter "proxyAddresses/any(address:address eq 'smtp:$inputEmail')" `
            -Property Id, DisplayName, Mail, ProxyAddresses, GroupTypes, MailEnabled, SecurityEnabled `
            -ErrorAction Stop
    }
    catch {
        Write-Warning "Microsoft Graph error (3): $($_.Exception.Message)"
        return
    }
}

if (-not $targetGroup) {
    Write-Host "No group found with email address (4): $inputEmail" -ForegroundColor Red
    return
}


####################################################################################
#################### GET GROUP MEMBERS #############################################
####################################################################################
try {
    $groupMembers = @(
        Get-MgGroupMemberAsUser `
            -GroupId $targetGroup.Id `
            -Property DisplayName, UserPrincipalName, Mail `
            -All `
            -ErrorAction Stop
    )
}
catch {
    Write-Host "Unable to retrieve group members (5): $($_.Exception.Message)" -ForegroundColor Red
    return
}


####################################################################################
################# BUILD GROUP DETAILS OUTPUT OBJECT ################################
####################################################################################
$tgDetails = [PSCustomObject]@{
    Name      = $targetGroup.DisplayName
    Email     = $targetGroup.Mail

    GroupType = if ($targetGroup.GroupTypes -contains "Unified") {
        "Microsoft 365"
    }
    elseif ($targetGroup.MailEnabled -and $targetGroup.SecurityEnabled) {
        "Mail-enabled Security"
    }
    elseif ($targetGroup.MailEnabled) {
        "Distribution"
    }
    elseif ($targetGroup.SecurityEnabled) {
        "Security"
    }
    else {
        "Unknown"
    }

    Aliases   = @(
        $targetGroup.ProxyAddresses |
        Where-Object { $_ -cmatch '^smtp:' } |
        ForEach-Object { $_ -creplace '^smtp:', '' } |
        Sort-Object
    )

    Members   = @($groupMembers |
        ForEach-Object { [PSCustomObject]@{
                Name  = $_.DisplayName
                UPN   = $_.UserPrincipalName
                Email = $_.Mail
            }
        }
    )
}


####################################################################################
#################### CONSOLE OUTPUT ################################################
####################################################################################
Write-Host ("`n{0,-20}" -f "Group Name:") -ForegroundColor Green -NoNewline
Write-Host $tgDetails.Name
Write-Host ("{0,-20}" -f "Group Email:") -ForegroundColor Green -NoNewline
Write-Host $tgDetails.Email
Write-Host ("{0,-20}" -f "Group Type:") -ForegroundColor Green -NoNewline
Write-Host $tgDetails.GroupType

Write-Host "`nEmail Aliases:" -ForegroundColor Green

if ($tgDetails.Aliases.Count -gt 0) {
    $tgDetails.Aliases | Out-Host
}
else {
    Write-Host "No email aliases found." -ForegroundColor Yellow
}

Write-Host "`n{Group Members Table}" -ForegroundColor Green

if ($tgDetails.Members.Count -gt 0) {
    $tgDetails.Members |
    Sort-Object Name |
    Format-Table -AutoSize |
    Out-Host
}
else {
    Write-Host "No user members found in this group." -ForegroundColor Yellow
}


####################################################################################
#################### CREATE CSV ####################################################
####################################################################################
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"

# Replace invalid filename characters
$groupName_2_fileName = $tgDetails.Name -replace '[^\w\.-]', '_'
$csvPath = ".\GroupDetails_${groupName_2_fileName}_$timestamp.csv"
$csvCreated = $false

try {
    @(
        "Group Name,$($tgDetails.Name)"
        "Group Email,$($tgDetails.Email)"
        "Group Type,$($tgDetails.GroupType)"
    ) | Set-Content -Path $csvPath -Encoding utf8

    # Email aliases
    "" | Add-Content -Path $csvPath -Encoding utf8
    "Email Aliases" | Add-Content -Path $csvPath -Encoding utf8

    if ($tgDetails.Aliases.Count -gt 0) {
        $tgDetails.Aliases |
        Add-Content -Path $csvPath -Encoding utf8
    }
    else {
        "No email aliases found." |
        Add-Content -Path $csvPath -Encoding utf8
    }

    # Group members
    "" | Add-Content -Path $csvPath -Encoding utf8
    "Group Members" | Add-Content -Path $csvPath -Encoding utf8

    if ($tgDetails.Members.Count -gt 0) {
        $tgDetails.Members |
        Sort-Object Name |
        ConvertTo-Csv -NoTypeInformation |
        Add-Content -Path $csvPath -Encoding utf8
    }
    else {
        "Name,UPN,Email" |
        Add-Content -Path $csvPath -Encoding utf8
    }
    $csvCreated = $true
}
catch {
    Write-Host "Unable to create CSV file: $($_.Exception.Message)" -ForegroundColor Red
}


####################################################################################
#################### EXECUTION TIME ################################################
####################################################################################
$stopwatch.Stop()

Write-Host ("`n{0,-25}" -f "File Name:") -ForegroundColor DarkCyan -NoNewline
Write-Host "Get_EntraGroup_Members.ps1" -ForegroundColor Cyan

if ($csvCreated) {
    Write-Host ("{0,-25}" -f "CSV file generated:") -ForegroundColor DarkYellow -NoNewline
    Write-Host (Resolve-Path $csvPath) -ForegroundColor DarkYellow
}

Write-Host ("{0,-25}" -f "Total Execution Time:") -ForegroundColor DarkGreen -NoNewline
Write-Host "($($stopwatch.Elapsed))" -ForegroundColor DarkGreen
