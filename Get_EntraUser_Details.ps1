# Get_EntraUser_Details.ps1
# Input = User email address (primary email or alias or username)
#
# Output: 
#   Display Name
#   Username
#   Primary Email Address
#   Job Title
#   Creation Date/Time
#   All Email Aliases
#   Group Memberships (Name, Type, Primary Email)
#
# Export:
#   CSV: Same Output


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
        -Scopes "User.Read.All", "Group.Read.All"`
        -NoWelcome `
        -ErrorAction Stop
}


####################################################################################
#################### Get Input Parameters ##########################################
####################################################################################
$inputEmail = Read-Host "Enter user's email address"
if ($inputEmail -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
    Write-Host "Invalid email format (1)" -ForegroundColor Red
    return
}


####################################################################################
#################### FIND TARGET USER ##############################################
####################################################################################
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

try {
    # Search both UPN & Mail properties
    # ConvergeDirect vs ConvergeMarketing
    $targetUser = Get-MgUser `
        -Filter "userPrincipalName eq '$inputEmail' or mail eq '$inputEmail'" `
        -Property Id, DisplayName, UserPrincipalName, Mail, JobTitle, CreatedDateTime, ProxyAddresses `
        -ErrorAction Stop
}
catch {
    # Throws on MicrosoftGraph error.  NOT triggered if no match found
    Write-Warning "Microsoft Graph error (2): $($_.Exception.Message)"
    return
}

# Check Alias emails if targetUser not found
if (-not $targetUser) {
    Write-Host "Not found as primary email or UPN. Checking aliases..." -ForegroundColor Yellow
    try {
        # SMTP = primary; smtp = alias
        $targetUser = Get-MgUser `
            -Filter "proxyAddresses/any(address:address eq 'smtp:$inputEmail')" `
            -Property Id, DisplayName, UserPrincipalName, Mail, JobTitle, CreatedDateTime, ProxyAddresses `
            -ErrorAction Stop
    }
    catch {
        # Throws on MicrosoftGraph error.  NOT triggered if no match found
        Write-Warning "Microsoft Graph error (3): $($_.Exception.Message)"
        return
    }
}

if (-not $targetUser) {
    Write-Host "No user found with email address (4): $inputEmail" -ForegroundColor Red
    return
}

# Only users have UPN/username & email properties
if (@($targetUser).Count -gt 1) {
    Write-Warning "Multiple users matched this email address (5)"
    $targetUser | Select-Object DisplayName, UserPrincipalName, Mail |
    Format-Table -AutoSize
    return
}


####################################################################################
#################### FIND GROUPS CONTAINING TARGET USER ############################
####################################################################################
try {
    # "-All" avoids pagination when query return multiple pages
    $tuAllGroups = Get-MgUserMemberOf `
        -UserId $targetUser.Id `
        -All `
        -ErrorAction Stop
}
catch {
    # Throws on MicrosoftGraph error.  NOT triggered if no match found
    Write-Warning "Microsoft Graph error: $($_.Exception.Message)"
    return
}

####################################################################################
################ FIND DETAILS FOR GROUPS CONTAINING TARGET USER ####################
####################################################################################
$tuagDetails = $tuAllGroups | ForEach-Object { 
    $groupDetails = Get-MgGroup `
        -GroupId $_.Id `
        -Property DisplayName, GroupTypes, MailEnabled, SecurityEnabled, Mail `
        -ErrorAction SilentlyContinue

    if ($groupDetails) {
        $groupResult = [PSCustomObject]@{
            "Group Name"  = $groupDetails.DisplayName
            "Group Email" = $groupDetails.Mail
        }

        $groupType = if ($targetGroup.GroupTypes -contains "Unified") {
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

        $groupResult | Add-Member -NotePropertyName "Group Type" -NotePropertyValue $groupType
        $groupResult
    }
}

####################################################################################
#################### BUILD USER DETAILS OBJECT ####################################
####################################################################################
$tuDetails = [PSCustomObject]@{
    Name            = $targetUser.DisplayName
    UPN             = $targetUser.UserPrincipalName
    JobTitle        = $targetUser.JobTitle
    CreatedDateTime = $targetUser.CreatedDateTime

    Aliases         = @(
        $targetUser.ProxyAddresses |
        Where-Object { $_ -cmatch "^smtp:" } |
        ForEach-Object { $_ -replace "^smtp:", "" } |
        Sort-Object
    )
    
    AllGroups       = @(
        foreach ($member in $tuagDetails) {
            [PSCustomObject]@{
                Name  = $member.DisplayName
                Email = $member.Mail
                GroupType = $member."Group Type"
            }
        }
    )
}




####################################################################################
#################### CONSOLE OUTPUT ################################################
####################################################################################
Write-Host ("{0,-20}" -f "Name:") -ForegroundColor Green -NoNewline
Write-Host $targetUser.DisplayName
Write-Host ("{0,-20}" -f "UPN:") -ForegroundColor Green -NoNewline
Write-Host $targetUser.UserPrincipalName
Write-Host ("{0,-20}" -f "Job Title:") -ForegroundColor Green -NoNewline
Write-Host $targetUser.JobTitle
Write-Host ("{0,-20}" -f "Creation Date/Time:") -ForegroundColor Green -NoNewline
Write-Host $targetUser.CreatedDateTime

Write-Host "`nEmail Aliases:" -ForegroundColor Cyan
if ($targetAliases) {
    $targetAliases |
    Sort-Object |
    ForEach-Object { Write-Host $_ }
}
else {
    Write-Host "No email aliases found." -ForegroundColor Yellow
}

Write-Host "`nGroup Memberships:" -ForegroundColor Cyan

if ($tuagDetails) {
    $tuagDetails |
    Sort-Object "Group Name" |
    Format-Table "Group Name", "Group Type", "Group Email" -AutoSize
}
else {
    Write-Host "No group memberships found." -ForegroundColor Yellow
}

Write-Host "`nTotal execution time: $($stopwatch.Elapsed)" -ForegroundColor Cyan
Write-Host "`nCSV file created:" -ForegroundColor Green
Write-Host $csvPath


# =========================================================
# CREATE CSV REPORT
# =========================================================
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$userName = $targetUser.UserPrincipalName.Split("@")[0]
$csvPath = ".\UserDetails_${userName}_$timestamp.csv"

# User information
"Name,$($targetUser.DisplayName)" | Out-File -FilePath $csvPath -Encoding utf8
"UPN,$($targetUser.UserPrincipalName)" | Out-File -FilePath $csvPath -Append -Encoding utf8
"Job Title,$($targetUser.JobTitle)" | Out-File -FilePath $csvPath -Append -Encoding utf8
"Creation Date/Time,$($targetUser.CreatedDateTime)" | Out-File -FilePath $csvPath -Append -Encoding utf8
"" | Out-File -FilePath $csvPath -Append -Encoding utf8
"Email Aliases" | Out-File -FilePath $csvPath -Append -Encoding utf8

if ($targetAliases) {
    foreach ($alias in ($targetAliases | Sort-Object)) {
        $alias | Out-File -FilePath $csvPath -Append -Encoding utf8
    }
}
else {
    "No email aliases found." |
    Out-File -FilePath $csvPath -Append -Encoding utf8
}


"" | Out-File -FilePath $csvPath -Append -Encoding utf8


# Group table header
"Group Name,Group Type,Group Email" | Out-File -FilePath $csvPath -Append -Encoding utf8


# Group data
if ($tuagDetails) {
    foreach ($groupResult in ($tuagDetails | Sort-Object "Group Name")) {
        "$($groupResult.'Group Name'),$($groupResult.'Group Type'),$($groupResult.'Group Email')" |
        Out-File -FilePath $csvPath -Append -Encoding utf8
    }
}
























# ####################################################################################
# #################### Find User's Email Aliases #####################################
# ####################################################################################
# # SMTP = primary; smtp = alias
# $targetAliases = $targetUser.ProxyAddresses |
# Where-Object { $_ -cmatch "^smtp:" } |
# ForEach-Object { $_ -replace "^smtp:", "" }