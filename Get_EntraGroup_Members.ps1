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
#   CSV: group details, aliases, and members


####################################################################################
#################### All Microsoft Graph Powershell Scripts START ##################
####################################################################################
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -ErrorAction Stop
}
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

$connection = Get-MgContext
if (-not $connection) {
    Connect-MgGraph `
        -Scopes "User.Read.All", "Group.Read.All" `
        -NoWelcome `
        -ErrorAction Stop
}


####################################################################################
#################### GET INPUT PARAMETERS ##########################################
####################################################################################
$inputEmail = (Read-Host "Enter group's email address").Trim()
if ($inputEmail -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
    Write-Host "Invalid email format (1)" -ForegroundColor Red
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
    # Throws on MicrosoftGraph error.  NOT triggered if no match found
    Write-Warning "Microsoft Graph error (2): $($_.Exception.Message)"
    return
}

# Check Alias emails if targetGroup not found
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
        # Throws on MicrosoftGraph error.  NOT triggered if no match found
        Write-Warning "Microsoft Graph error (3): $($_.Exception.Message)"
        return
    }
}

# targetGroup not found, invalid input
if (-not $targetGroup) {
    Write-Host "No group found with email address (4): $inputEmail" -ForegroundColor Red
    return
}


####################################################################################
#################### DETERMINE GROUP TYPE ##########################################
####################################################################################
$tgType = if ($targetGroup.GroupTypes -contains "Unified") {
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



####################################################################################
#################### RETRIEVE GROUP MEMBERS ########################################
####################################################################################
try {
    $tgMembers = @(
        Get-MgGroupMemberAsUser `
            -GroupId $targetGroup.Id `
            -Property Id, DisplayName, UserPrincipalName, Mail `
            -All `
            -ErrorAction Stop
    )
}
catch {
    Write-Host "Unable to retrieve group members (5)." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    return
}


####################################################################################
#################### BUILD GROUP DETAILS OBJECT ####################################
####################################################################################
# Building object for Output
$tgDetails = [PSCustomObject]@{
    Name      = $targetGroup.DisplayName
    Email     = $targetGroup.Mail
    GroupType = $tgType

    Aliases   = @(
        $targetGroup.ProxyAddresses |
        Where-Object { $_ -cmatch '^smtp:' } |
        ForEach-Object { $_ -creplace '^smtp:', '' } |
        Sort-Object
    )

    Members   = @(
        foreach ($member in $tgMembers) {
            [PSCustomObject]@{
                Name  = $member.DisplayName
                UPN   = $member.UserPrincipalName
                Email = $member.Mail
            }
        }
    )
}


####################################################################################
#################### CONSOLE OUTPUT ################################################
####################################################################################

Write-Host ("`n{0,-18}" -f "Group Name:") -ForegroundColor Green -NoNewline
Write-Host $tgDetails.Name
Write-Host ("{0,-18}" -f "Group Email:") -ForegroundColor Green -NoNewline
Write-Host $tgDetails.Email
Write-Host ("{0,-18}" -f "Group Type:") -ForegroundColor Green -NoNewline
Write-Host $tgDetails.GroupType


# No table, so each array element also newLine
Write-Host "`nEmail Aliases:" -ForegroundColor Green
if ($tgDetails.Aliases.Count -gt 0) {
    $tgDetails.Aliases | 
    Out-Host
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
# [^\w\.-] --> any character not {letter/number/underscore/period/hyphen}
$groupName_2_fileName = $tgDetails.Name -replace '[^\w\.-]', '_'
$csvPath = ".\GroupDetails_${groupName_2_fileName}_$timestamp.csv"

try {
    @(
        "Group Name,$($tgDetails.Name)"
        "Group Email,$($tgDetails.Email)"
        "Group Type,$($tgDetails.GroupType)"
        ""
        "Email Aliases"
    ) | Set-Content -Path $csvPath -Encoding utf8

    if ($tgDetails.Aliases.Count -gt 0) {
        $tgDetails.Aliases |
        Add-Content -Path $csvPath -Encoding utf8
    }
    else {
        "No email aliases found." |
        Add-Content -Path $csvPath -Encoding utf8
    }

    "" | Add-Content -Path $csvPath -Encoding utf8
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
    
    Write-Host "`nCSV file created:" -ForegroundColor Green
    Write-Host (Resolve-Path $csvPath)
}
catch {
    Write-Host "Unable to create CSV file: $($_.Exception.Message)" -ForegroundColor Red
}


####################################################################################
#################### EXECUTION TIME ################################################
####################################################################################
$stopwatch.Stop()
Write-Host "`nTotal execution time: $($stopwatch.Elapsed)" -ForegroundColor Cyan