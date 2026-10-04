$userEmail = Read-Host "Enter user's email address"


# Check: Input = Valid Email
if ($userEmail -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
    Write-Host "Invalid email format." -ForegroundColor Red -BackgroundColor White
    exit
}

Write-Host "Email format looks valid." -ForegroundColor Yellow


# Connect to Graph if necessary
$context = Get-MgContext

if (-not $context) {
    Write-Host "Not connected to Microsoft Graph. Connecting..." -ForegroundColor Gray

    Connect-MgGraph -Scopes `
        "User.Read.All", `
        "Group.Read.All"
}


# Check: Input = User Account UPN or User Alias
try {

    $user = Get-MgUser `
        -UserId $userEmail `
        -Property Id,DisplayName,UserPrincipalName,JobTitle,CreatedDateTime,ProxyAddresses `
        -ErrorAction Stop

}
catch {

    Write-Host "Not UPN. Checking email aliases..." -ForegroundColor Yellow

    try {

        $user = Get-MgUser `
            -Filter "proxyAddresses/any(address:address eq 'smtp:$userEmail')" `
            -Property Id,DisplayName,UserPrincipalName,JobTitle,CreatedDateTime,ProxyAddresses `
            -ErrorAction Stop

    }
    catch {

        $user = $null
        Write-Host "Alias lookup failed" -ForegroundColor Yellow
        Write-Host $_.Exception.Message -ForegroundColor Red

    }
}



if ($user) {

    Write-Host ""
    Write-Host "User found in Entra:" -ForegroundColor DarkBlue

    Write-Host "Name: " -ForegroundColor Green -BackgroundColor White -NoNewline
    Write-Host "   $($user.DisplayName)"

    Write-Host "UPN: " -ForegroundColor Green -BackgroundColor White -NoNewline
    Write-Host $user.UserPrincipalName

    Write-Host "Job Title: " -ForegroundColor Green -BackgroundColor White -NoNewline
    Write-Host $user.JobTitle

    Write-Host "Creation Date/Time: " -ForegroundColor Green -BackgroundColor White -NoNewline
    Write-Host $user.CreatedDateTime


    # Get user's group memberships
    Write-Host ""
    Write-Host "Group Memberships:" -ForegroundColor Green -BackgroundColor White

    try {

        $groups = Get-MgUserMemberOf `
            -UserId $user.Id `
            -All `
            -ErrorAction Stop

        if ($groups) {

            foreach ($group in $groups) {

                $groupDetails = Get-MgGroup `
                    -GroupId $group.Id `
                    -Property DisplayName,Id `
                    -ErrorAction SilentlyContinue

                if ($groupDetails) {
                    Write-Host "  $($groupDetails.DisplayName)"
                }
            }

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