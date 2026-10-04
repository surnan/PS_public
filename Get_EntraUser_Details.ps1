
$userEmail = Read-Host "Enter user's email address"

# Verify string = email format
if ($userEmail -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
    Write-Host "Invalid email format." -ForegroundColor Red -BackgroundColor White
    exit
}


# Connect to Graph if necessary
$context = Get-MgContext
if (-not $context) {
    Write-Host "Not connected to Microsoft Graph. Connecting..." -ForegroundColor Gray
    Connect-MgGraph -Scopes "User.Read.All", "Group.Read.All"
}

# Check: Input = User Account UPN or User Alias
try {
    #Get-MgUser cmdlet; "UserID" accepts GUID & UserPrincipalName
    $user = Get-MgUser `
        -UserId $userEmail `
        -Property Id, DisplayName, UserPrincipalName, JobTitle, CreatedDateTime, ProxyAddresses `
        -ErrorAction Stop
}
catch {
    # $userEmail not found as UserID, now checking aliases.
    try {
        $user = Get-MgUser `
            -Filter "proxyAddresses/any(address:address eq 'smtp:$userEmail')" `
            -Property Id, DisplayName, UserPrincipalName, jobTitle, createdDateTime, ProxyAddresses `
            -ErrorAction Stop
    }
    catch {
        $user = $null  #just in-case user = blank from filter returning zero results
        Write-Host "Alias lookup failed" -ForegroundColor Yellow
        Write-Host $_.Exception.Message -ForegroundColor Red
    }
}

if ($user) {
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

    # Get all of user's alias
    Write-Host ""
    Write-Host ""
    Write-Host "Email Aliases:" -ForegroundColor Green
    
    #SMTP = primary; smtp = alias
    $aliases = $user.ProxyAddresses |
    Where-Object { $_ -cmatch '^smtp:' } |
    ForEach-Object { $_ -replace "^smtp:", "" }

    #Only one propery, so Sort-Object doesn't need property
    if ($aliases) {
        $aliases | Sort-Object
    }
    else {
        Write-Host "No email aliases found." -ForegroundColor Red
    }
    
    Write-Host ""
    Write-Host "Group Memberships Below:" -ForegroundColor DarkGreen
    
    # $groups = Groups w/ $user.Id as member
    # "-All" avoids pagination & necessary when pulling multiple from query
    try {
        $groups = Get-MgUserMemberOf `
            -UserId $user.Id `
            -All `
            -ErrorAction Stop

        if ($groups) {
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

                    [PSCustomObject]@{
                        "Group Name"  = $groupDetails.DisplayName
                        "Group Type"  = $groupType
                        "Group Email" = $groupDetails.Mail
                    }
                }
            }

            $groupResults |
            Sort-Object "Group Name" |
            Format-Table -AutoSize

        }
        else {
            Write-Host "User is not a member of any groups." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "Unable to retrieve group memberships." -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
    }
}
else {
    Write-Host "No Entra user was found with that address." -ForegroundColor Red
}




# Learn > Microsoft Graph > User Resource Type
# https://learn.microsoft.com/en-us/graph/api/resources/user?view=graph-rest-1.0


# -Filter "proxyAddresses/any(address:address eq 'smtp:$userEmail')" `
# -ConsistencyLevel eventual ` 
# -Property Id,DisplayName,UserPrincipalName,jobTitle, createdDateTime, ProxyAddresses `
# -ErrorAction Stop