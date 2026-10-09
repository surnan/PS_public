# Get_EntraUser_Details.ps1
# Input = $inputEmail
# Output = Display Name, Primary Email, Job Title, Creation Date/Time
# Output = All email aliases
# Output = All Group memberships (Name, Group Type, Group Email)

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
    Write-Host "Invalid email format." -ForegroundColor Red
    exit
}

####################################################################################
#################### FIND targetUser - Start #######################################
####################################################################################

try {
    #Get-MgUser cmdlet; "UserID" accepts GUID & UserPrincipalName
    $user = Get-MgUser `
        -UserId $inputEmail `
        -Property Id, DisplayName, UserPrincipalName, JobTitle, CreatedDateTime, ProxyAddresses `
        -ErrorAction Stop
}
catch {
    # UPN lookup failed, now check aliases
    try {
        $user = Get-MgUser `
            -Filter "proxyAddresses/any(address:address eq 'smtp:$inputEmail')" `
            -Property Id, DisplayName, UserPrincipalName, JobTitle, CreatedDateTime, ProxyAddresses `
            -ErrorAction Stop
    }
    catch {
        $user = $null
        Write-Host "Alias lookup failed." -ForegroundColor Yellow
        Write-Host $_.Exception.Message -ForegroundColor Red
    }
}


if ($user) {
    # =========================================================
    # USER INFORMATION
    # =========================================================
    Write-Host ""
    Write-Host "User found in Entra:" -ForegroundColor DarkBlue

    Write-Host ("{0,-20}" -f "Name:") -ForegroundColor Green -NoNewline
    Write-Host $user.DisplayName

    Write-Host ("{0,-20}" -f "UPN:") -ForegroundColor Green -NoNewline
    Write-Host $user.UserPrincipalName

    Write-Host ("{0,-20}" -f "Job Title:") -ForegroundColor Green -NoNewline
    Write-Host $user.JobTitle

    Write-Host ("{0,-20}" -f "Creation Date/Time:") -ForegroundColor Green -NoNewline
    Write-Host $user.CreatedDateTime


    # =========================================================
    # EMAIL ALIASES
    # =========================================================
    Write-Host ""
    Write-Host "Email Aliases:" -ForegroundColor Green

    #SMTP = primary; smtp = alias
    $aliases = $user.ProxyAddresses |
        Where-Object { $_ -cmatch "^smtp:" } |
        ForEach-Object { $_ -replace "^smtp:", "" }

    #Only one propery, so Sort-Object doesn't need property
    if ($aliases) {
        $aliases | Sort-Object
    }
    else {
        Write-Host "No email aliases found." -ForegroundColor Yellow
    }


    # =========================================================
    # GROUP MEMBERSHIPS
    # =========================================================


    Write-Host ""
    Write-Host "Group Memberships Below:" -ForegroundColor Green

    try {
        # $groups = Groups w/ $user.Id as member
        # "-All" avoids pagination when query return multiple
        $groups = Get-MgUserMemberOf `
            -UserId $user.Id `
            -All `
            -ErrorAction Stop

        if ($groups) {
            # Determine group type
            $groupResults = foreach ($group in $groups) {

                $groupDetails = Get-MgGroup `
                    -GroupId $group.Id `
                    -Property DisplayName, GroupTypes, MailEnabled, SecurityEnabled, Mail `
                    -ErrorAction SilentlyContinue

                if ($groupDetails) {

                    # Determine group type
                    if ($groupDetails.GroupTypes -contains "Unified") {
                        $groupType = "Microsoft 365"
                    }
                    elseif ($groupDetails.MailEnabled -and $groupDetails.SecurityEnabled) {
                        $groupType = "Mail-enabled Security"
                    }
                    elseif ($groupDetails.MailEnabled -and -not $groupDetails.SecurityEnabled) {
                        $groupType = "Distribution"
                    }
                    elseif (-not $groupDetails.MailEnabled -and $groupDetails.SecurityEnabled) {
                        $groupType = "Security"
                    }
                    else {
                        $groupType = "Unknown"
                    }
                    
                    # Create object for $groupResults
                    $groupResult = [PSCustomObject]@{
                        "Group Name"  = $groupDetails.DisplayName
                        "Group Type"  = $groupType
                        "Group Email" = $groupDetails.Mail
                    }
                    # Output goes to console if it's not captured
                    # ForEach loop captures $groupResult for $groupResults 
                    $groupResult
                }
            }
            $groupResults |
                Sort-Object "Group Name" |
                Format-Table -AutoSize
        }
        else {
            $groupResults = @()
            Write-Host "User is not a member of any groups." -ForegroundColor Yellow
        }
    }
    catch {
        $groupResults = @()
        Write-Host "Unable to retrieve group memberships." -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
    }

    # =========================================================
    # CREATE CSV REPORT
    # =========================================================

    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"

    $userName = $user.UserPrincipalName.Split("@")[0]

    $csvPath = ".\UserDetails_${userName}_$timestamp.csv"


    # User information
    "Name,$($user.DisplayName)" |
        Out-File `
            -FilePath $csvPath `
            -Encoding utf8

    "UPN,$($user.UserPrincipalName)" |
        Out-File `
            -FilePath $csvPath `
            -Append `
            -Encoding utf8

    "Job Title,$($user.JobTitle)" |
        Out-File `
            -FilePath $csvPath `
            -Append `
            -Encoding utf8

    "Creation Date/Time,$($user.CreatedDateTime)" |
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


    # Group table header
    "Group Name,Group Type,Group Email" |
        Out-File `
            -FilePath $csvPath `
            -Append `
            -Encoding utf8


    # Group data
    if ($groupResults) {

        foreach ($groupResult in ($groupResults | Sort-Object "Group Name")) {

            "$($groupResult.'Group Name'),$($groupResult.'Group Type'),$($groupResult.'Group Email')" |
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
}
else {
    Write-Host "No Entra user was found with that address." -ForegroundColor Red
}