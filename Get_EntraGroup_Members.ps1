# Get_EntraGroup_Members.ps1
# Input = $inputEmail
# Output = Display Name, Primary Email, Group Type
# Output = All email aliases
# Output = Each member (play Name, UPN, Primary Email Address)


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
#################### Get Input Parameters ##########################################
####################################################################################

$inputEmail = Read-Host "Enter group's email address"
if ($inputEmail -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
    Write-Host "Invalid email format." -ForegroundColor Red
    exit
}


####################################################################################
#################### FIND targetGroup - Start ######################################
####################################################################################

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# No match, Get-MgGroup returns nothing but doesn't trigger catch
$targetGroup = $null

try {
    $targetGroup = Get-MgGroup `
        -Filter "mail eq '$inputEmail'" `
        -Property Id, DisplayName, Mail, ProxyAddresses, GroupTypes, MailEnabled, SecurityEnabled `
        -ErrorAction Stop
}
catch {
    # Throws on Graph error.  NOT when no match found
    Write-Warning "Microsoft Graph error: $($_.Exception.Message)"
}

# If not found by primary email, check aliases
if (-not $targetGroup) {
    Write-Host "Not found as primary email. Checking aliases..." -ForegroundColor Yellow
    try {
        $targetGroup = Get-MgGroup `
            -Filter "proxyAddresses/any(address:address eq 'smtp:$inputEmail')" `
            -Property Id, DisplayName, Mail, ProxyAddresses, GroupTypes, MailEnabled, SecurityEnabled `
            -ErrorAction Stop
    }
    catch {
        # Throws on Graph error.  NOT when no match found
        Write-Warning "Microsoft Graph error: $($_.Exception.Message)"
        exit
    }
}

if (-not $targetGroup) {
    Write-Host "Not found as alias email" -ForegroundColor Yellow
    Write-Host "Email address not found in Exchange Online System" -ForegroundColor Red
    exit
}


# =========================================================
# DETERMINE GROUP TYPE - Start
# =========================================================

if ($targetGroup.GroupTypes -contains "Unified") {
    $groupType = "Microsoft 365"
}
elseif ($targetGroup.MailEnabled -and $targetGroup.SecurityEnabled) {
    $groupType = "Mail-enabled Security"
}
elseif ($targetGroup.MailEnabled -and -not $targetGroup.SecurityEnabled) {
    $groupType = "Distribution"
}
elseif (-not $targetGroup.MailEnabled -and $targetGroup.SecurityEnabled) {
    $groupType = "Security"
}
else {
    $groupType = "Unknown"
}


# =========================================================
# Target Group found.  Load all aliases
# =========================================================

$aliases = $targetGroup.ProxyAddresses |
Where-Object { $_ -cmatch '^smtp:' } |
ForEach-Object { $_ -replace '^smtp:', '' }

# =========================================================
# OUTPUT 
# =========================================================

# Display group details
Write-Host ("`n{0,-18}" -f "Group Name:") -ForegroundColor Green -NoNewline
Write-Host $targetGroup.DisplayName

Write-Host ("{0,-18}" -f "Group Email:") -ForegroundColor Green -NoNewline
Write-Host $targetGroup.Mail

Write-Host ("{0,-18}" -f "Group Type:") -ForegroundColor Green -NoNewline
Write-Host $groupType

# Display group aliases
Write-Host "`nEmail Aliases:" -ForegroundColor Green
if ($aliases) {
    $aliases | Sort-Object
}
else {
    Write-Host "No email aliases found." -ForegroundColor Yellow
}

# Display group members
Write-Host "`n`nGroup Members:" -ForegroundColor Green

try {
    $members = Get-MgGroupMemberAsUser `
        -GroupId $targetGroup.Id `
        -Property Id, DisplayName, UserPrincipalName, Mail `
        -All `
        -ErrorAction Stop
}
catch {
    Write-Host "Unable to retrieve group members." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit
}

if ($members) {
    $memberResults = foreach ($member in $members) {
        [PSCustomObject]@{
            "Name"  = $member.DisplayName
            "UPN"   = $member.UserPrincipalName
            "Email" = $member.Mail
        }
    }
    $memberResults |
    Sort-Object Name |
    Format-Table -AutoSize
}
else {
    $memberResults = @()
    Write-Host "No user members found in this group." -ForegroundColor Yellow
}


# =========================================================
# CREATE CSV REPORT
# =========================================================

$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$groupNameSafe = $targetGroup.DisplayName -replace '[^\w\.-]', '_'
$csvPath = ".\GroupDetails_${groupNameSafe}_$timestamp.csv"

# Group information
"Group Name,$($targetGroup.DisplayName)" |
Out-File `
    -FilePath $csvPath `
    -Encoding utf8

"Group Email,$($targetGroup.Mail)" |
Out-File `
    -FilePath $csvPath `
    -Append `
    -Encoding utf8

"Group Type,$groupType" |
Out-File `
    -FilePath $csvPath `
    -Append `
    -Encoding utf8


# Blank line
"" |
Out-File `
    -FilePath $csvPath `
    -Append `
    -Encoding utf8

# Email aliases
"Email Aliases" |
Out-File `
    -FilePath $csvPath `
    -Append `
    -Encoding utf8

if ($aliases) {

    foreach ($alias in ($aliases | Sort-Object)) {

        $alias |
        Out-File `
            -FilePath $csvPath `
            -Append `
            -Encoding utf8
    }
}
else {
    "No email aliases found." |
    Out-File `
        -FilePath $csvPath `
        -Append `
        -Encoding utf8
}


# Blank line
"" |
Out-File `
    -FilePath $csvPath `
    -Append `
    -Encoding utf8

# Group members table header
"Name,UPN,Email" |
Out-File `
    -FilePath $csvPath `
    -Append `
    -Encoding utf8

# Group member data
if ($memberResults) {
    foreach ($memberResult in ($memberResults | Sort-Object Name)) {
        "$($memberResult.Name),$($memberResult.UPN),$($memberResult.Email)" |
        Out-File `
            -FilePath $csvPath `
            -Append `
            -Encoding utf8
    }
}


# =========================================================
# OUTPUT part 2
# =========================================================

$stopwatch.Stop()

Write-Host ""
Write-Host "CSV file created:" -ForegroundColor Green
Write-Host $csvPath

Write-Host "`nTotal execution time: $($stopwatch.Elapsed)" -ForegroundColor Cyan
