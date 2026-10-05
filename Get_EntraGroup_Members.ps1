$groupEmail = Read-Host "Enter group's email address"


# Verify email format
if ($groupEmail -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
    Write-Host "Invalid email format." -ForegroundColor Red
    exit
}


# Connect to Microsoft Graph if necessary
$context = Get-MgContext
if (-not $context) {
    Write-Host "Not connected to Microsoft Graph. Connecting..." -ForegroundColor Gray
    Connect-MgGraph -Scopes `
        "Group.Read.All", `
        "User.Read.All"
}


# =========================================================
# FIND GROUP
# =========================================================

$group = $null

try {
    # Checking primary email address
    $group = Get-MgGroup `
        -Filter "mail eq '$groupEmail'" `
        -Property Id, DisplayName, Mail, ProxyAddresses, GroupTypes, MailEnabled, SecurityEnabled `
        -ErrorAction Stop
}
catch {
    Write-Host "Primary email lookup failed." -ForegroundColor Yellow
}


# If not found by primary email, check aliases
if (-not $group) {

    Write-Host "Not found by primary email. Checking aliases..." -ForegroundColor Yellow

    try {

        $group = Get-MgGroup `
            -Filter "proxyAddresses/any(address:address eq 'smtp:$groupEmail')" `
            -Property Id, DisplayName, Mail, ProxyAddresses, GroupTypes, MailEnabled, SecurityEnabled `
            -ErrorAction Stop

    }
    catch {

        Write-Host "Alias lookup failed." -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        exit
    }
}


if (-not $group) {

    Write-Host "No group was found with email address $groupEmail" -ForegroundColor Red
    exit
}


# =========================================================
# DETERMINE GROUP TYPE
# =========================================================

if ($group.GroupTypes -contains "Unified") {
    $groupType = "Microsoft 365"
}
elseif ($group.MailEnabled -and $group.SecurityEnabled) {
    $groupType = "Mail-enabled Security"
}
elseif ($group.MailEnabled -and -not $group.SecurityEnabled) {
    $groupType = "Distribution"
}
elseif (-not $group.MailEnabled -and $group.SecurityEnabled) {
    $groupType = "Security"
}
else {
    $groupType = "Unknown"
}


# =========================================================
# GET GROUP ALIASES
# =========================================================

$aliases = $group.ProxyAddresses |
    Where-Object { $_ -cmatch '^smtp:' } |
    ForEach-Object { $_ -replace '^smtp:', '' }


# =========================================================
# DISPLAY GROUP INFORMATION
# =========================================================

Write-Host ""
Write-Host "Group found:" -ForegroundColor DarkBlue

Write-Host ("{0,-18}" -f "Group Name:") -ForegroundColor Green -NoNewline
Write-Host $group.DisplayName

Write-Host ("{0,-18}" -f "Group Email:") -ForegroundColor Green -NoNewline
Write-Host $group.Mail

Write-Host ("{0,-18}" -f "Group Type:") -ForegroundColor Green -NoNewline
Write-Host $groupType


Write-Host ""
Write-Host "Email Aliases:" -ForegroundColor Green

if ($aliases) {

    $aliases | Sort-Object

}
else {

    Write-Host "No email aliases found." -ForegroundColor Yellow

}


# =========================================================
# GET GROUP MEMBERS
# =========================================================

Write-Host ""
Write-Host "Group Members:" -ForegroundColor Green


try {
    $members = Get-MgGroupMemberAsUser `
        -GroupId $group.Id `
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
$groupNameSafe = $group.DisplayName -replace '[^\w\.-]', '_'
$csvPath = ".\GroupDetails_${groupNameSafe}_$timestamp.csv"

# Group information
"Group Name,$($group.DisplayName)" |
    Out-File `
        -FilePath $csvPath `
        -Encoding utf8

"Group Email,$($group.Mail)" |
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
# FINISHED
# =========================================================

Write-Host ""
Write-Host "CSV file created:" -ForegroundColor Green
Write-Host $csvPath