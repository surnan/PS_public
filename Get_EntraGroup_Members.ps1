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
    Write-Host "Invalid email format (1)" -ForegroundColor Red
    return
}


####################################################################################
#################### FIND targetGroup - Start ######################################
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

if (-not $targetGroup) {
    Write-Host "No group found with email address (4): $inputEmail" -ForegroundColor Red
    return
}


####################################################################################
#################### Target Group found.  Load all aliases #########################
####################################################################################
# SMTP = primary; smtp = alias
$tgAliases = $targetGroup.ProxyAddresses |
Where-Object { $_ -cmatch '^smtp:' } |
ForEach-Object { $_ -replace '^smtp:', '' }

####################################################################################
#################### DETERMINE GROUP TYPE - Start ##################################
####################################################################################
if ($targetGroup.GroupTypes -contains "Unified") {
    $tgType = "Microsoft 365"
}
elseif ($targetGroup.MailEnabled -and $targetGroup.SecurityEnabled) {
    $tgType = "Mail-enabled Security"
}
elseif ($targetGroup.MailEnabled -and -not $targetGroup.SecurityEnabled) {
    $tgType = "Distribution"
}
elseif (-not $targetGroup.MailEnabled -and $targetGroup.SecurityEnabled) {
    $tgType = "Security"
} 
else {
    $tgType = "Unknown"
}

####################################################################################
#################### OUTPUT - start ################################################
####################################################################################
$stopwatch.Stop()

Write-Host ("`n{0,-18}" -f "Group Name:") -ForegroundColor Green -NoNewline
Write-Host $targetGroup.DisplayName

Write-Host ("{0,-18}" -f "Group Email:") -ForegroundColor Green -NoNewline
Write-Host $targetGroup.Mail

Write-Host ("{0,-18}" -f "Group Type:") -ForegroundColor Green -NoNewline
Write-Host $tgType

#Only one propery, so Sort-Object doesn't need property
Write-Host "`nEmail Aliases:" -ForegroundColor Green
if ($tgAliases) {
    $tgAliases | Sort-Object
}
else {
    Write-Host "No email aliases found." -ForegroundColor Yellow
}

# Display group members
Write-Host "`n`nGroup Members:" -ForegroundColor Green
try {
    $tgMembers = Get-MgGroupMemberAsUser `
        -GroupId $targetGroup.Id `
        -Property Id, DisplayName, UserPrincipalName, Mail `
        -All `
        -ErrorAction Stop
}
catch {
    Write-Host "Unable to retrieve group members (5)" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    return
}

if ($tgMembers) {
    $memberResults = foreach ($member in $tgMembers) {
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

####################################################################################
#################### Create CSV ####################################################
####################################################################################
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$groupNameSafe = $targetGroup.DisplayName -replace '[^\w\.-]', '_'
$csvPath = ".\GroupDetails_${groupNameSafe}_$timestamp.csv"

"Group Name,$($targetGroup.DisplayName)" | Out-File -FilePath $csvPath -Encoding utf8
"Group Email,$($targetGroup.Mail)" | Out-File -FilePath $csvPath -Append -Encoding utf8
"Group Type,$tgType" | Out-File -FilePath $csvPath -Append -Encoding utf8
"" | Out-File -FilePath $csvPath -Append -Encoding utf8
"Email Aliases" | Out-File -FilePath $csvPath -Append -Encoding utf8

# Aliases
if ($tgAliases) {
    foreach ($alias in ($tgAliases | Sort-Object)) {
        $alias | Out-File -FilePath $csvPath -Append -Encoding utf8
    }
}
else {
    "No email aliases found." | Out-File -FilePath $csvPath -Append -Encoding utf8
}


# Groupmembers
"" | Out-File -FilePath $csvPath -Append -Encoding utf8
"Name,UPN,Email" | Out-File -FilePath $csvPath -Append -Encoding utf8
if ($memberResults) {
    foreach ($memberResult in ($memberResults | Sort-Object Name)) {
        "$($memberResult.Name),$($memberResult.UPN),$($memberResult.Email)" | Out-File -FilePath $csvPath -Append -Encoding utf8
    }
}




Write-Host "`nCSV file created:" -ForegroundColor Green
Write-Host $csvPath
Write-Host "`nTotal execution time: $($stopwatch.Elapsed)" -ForegroundColor Cyan