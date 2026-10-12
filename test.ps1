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
    Write-Warning "Microsoft Graph error (6): $($_.Exception.Message)"
    return
}

####################################################################################
################# BUILD USER DETAILS OUTPUT OBJECT #################################
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
        $tuAllGroups | ForEach-Object { 
            $groupDetails = Get-MgGroup `
                -GroupId $_.Id `
                -Property DisplayName, GroupTypes, MailEnabled, SecurityEnabled, Mail `
                -ErrorAction SilentlyContinue

            if ($groupDetails) {
                $groupResult = [PSCustomObject]@{
                    "Group Name"  = $groupDetails.DisplayName
                    "Group Email" = $groupDetails.Mail
                }

                $groupType = if ($groupDetails.GroupTypes -contains "Unified") {
                    "Microsoft 365"
                }
                elseif ($groupDetails.MailEnabled -and $groupDetails.SecurityEnabled) {
                    "Mail-enabled Security"
                }
                elseif ($groupDetails.MailEnabled) {
                    "Distribution"
                }
                elseif ($groupDetails.SecurityEnabled) {
                    "Security"
                }
                else {
                    "Unknown"
                }

                $groupResult | Add-Member -NotePropertyName "Group Type" -NotePropertyValue $groupType
                $groupResult
            }
        }
    )
}


####################################################################################
#################### CONSOLE OUTPUT ################################################
####################################################################################
Write-Host ("{0,-20}" -f "Name:") -ForegroundColor Green -NoNewline
Write-Host $tuDetails.Name
Write-Host ("{0,-20}" -f "UPN:") -ForegroundColor Green -NoNewline
Write-Host $tuDetails.UPN
Write-Host ("{0,-20}" -f "Job Title:") -ForegroundColor Green -NoNewline
Write-Host $tuDetails.JobTitle
Write-Host ("{0,-20}" -f "Creation Date/Time:") -ForegroundColor Green -NoNewline
Write-Host $tuDetails.CreatedDateTime

Write-Host "`nEmail Aliases:" -ForegroundColor Cyan
if ($tuDetails.Aliases.Count -gt 0) {
    $tuDetails.Aliases |
    Out-Host
}
else {
    Write-Host "No email aliases found." -ForegroundColor Yellow
}

Write-Host "`n{Group Memberships}" -ForegroundColor Cyan

if ($tuDetails.AllGroups.Count -gt 0) {
    $tuDetails.AllGroups |
    Sort-Object "Group Name" |
    Format-Table "Group Name", "Group Type", "Group Email" -AutoSize
}
else {
    Write-Host "No group memberships found." -ForegroundColor Yellow
}


####################################################################################
#################### CREATE CSV ####################################################
####################################################################################
$stopwatch.Stop()
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"

$filename = $targetUser.UserPrincipalName.Split("@")[0]
$csvPath = ".\UserDetails_${filename}_$timestamp.csv"

try {
    @(
        "Name,$($tuDetails.Name)"
        "UPN,$($tuDetails.UPN)"
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

    # Export group membership table
    if ($tuDetails.AllGroups.Count -gt 0) {
        $tuDetails.AllGroups |
        Sort-Object "Group Name" |
        Select-Object "Group Name", "Group Type", "Group Email" |
        ConvertTo-Csv -NoTypeInformation |
        Add-Content -Path $csvPath -Encoding utf8 -ErrorAction Stop
    }
    else {
        "Group Name,Group Type,Group Email" |
        Add-Content -Path $csvPath -Encoding utf8 -ErrorAction Stop
    }
}
catch {
    Write-Host "Unable to create CSV file (7): $($_.Exception.Message)" -ForegroundColor Red
    return
}





####################################################################################
#################### EXECUTION TIME ################################################
####################################################################################
Write-Host ("`n{0,-25}" -f "File Name:") -ForegroundColor DarkCyan -NoNewline
Write-Host "Get_EntraGroup_Members.ps1" -ForegroundColor Cyan

Write-Host ("{0,-25}" -f "CSV file generated:") -ForegroundColor DarkYellow -NoNewline
Write-Host (Resolve-Path $csvPath) -ForegroundColor DarkYellow

Write-Host ("{0,-25}" -f "Total Execution Time:") -ForegroundColor DarkGreen -NoNewline
Write-Host "($stopwatch.Elapsed)" -ForegroundColor DarkGreen