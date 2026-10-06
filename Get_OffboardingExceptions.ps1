# Connect to Microsoft Graph if necessary
$context = Get-MgContext
if (-not $context) {
    Write-Host "Not connected to Microsoft Graph. Connecting..." -ForegroundColor Gray
    Connect-MgGraph -Scopes `
        "User.Read.All", `
        "Group.Read.All", `
        "LicenseAssignment.Read.All"
}


# -"accountEnabled eq false"  = users w/ sign-in disabled.
$disabledUsers = Get-MgUser `
    -Filter "accountEnabled eq false" `
    -Property Id, DisplayName, UserPrincipalName, AccountEnabled, ProxyAddresses `
    -All `
    -ErrorAction Stop


Write-Host "Loading account." -ForegroundColor Yellow -NoNewline
$results = @(
    foreach ($user in $disabledUsers) {
        # Write-Host "`nUser: $($user.DisplayName)" -ForegroundColor Cyan
        # Write-Host "Sign-in name: $($user.UserPrincipalName)"

        # Get all SMTP addresses, including the primary email address
        $smtpAddresses = @(
            $user.ProxyAddresses |
            Where-Object { $_ -like "smtp:*" } |
            ForEach-Object { $_ -replace "^smtp:", "" }
        )

        # Show one address, otherwise show the count
        if ($smtpAddresses.Count -eq 1) {
            $aliasDisplay = $smtpAddresses[0]
        }
        else {
            $aliasDisplay = $smtpAddresses.Count
        }        
        

        # Get the user's direct group memberships
        $groups = @(
            Get-MgUserMemberOfAsGroup `
                -UserId $user.Id `
                -All `
                -ErrorAction Stop
        )

        # Show one group name, otherwise show the count
        if ($groups.Count -eq 1) {
            $groupDisplay = $groups[0].DisplayName
        }
        else {
            $groupDisplay = $groups.Count
        }

        $result = [PSCustomObject]@{
            DisplayName       = $user.DisplayName
            UserPrincipalName = $user.UserPrincipalName
            AccountEnabled    = $user.AccountEnabled
            SMTPAddresses     = $aliasDisplay
            Groups            = $groupDisplay
        }
        $result
        Write-Host "." -ForegroundColor Yellow -NoNewline
    }
)

Write-Host ""

if ($results.Count -eq 0) {
    Write-Host "No users with sign-in disabled were found." -ForegroundColor Yellow
    return
}
    

# Display the report in the console
# $results |
# Format-Table -AutoSize -Wrap


# Display the report in the console
$results |
    Format-Table `
        DisplayName, `
        UserPrincipalName, `
        AccountEnabled, `
        @{Name = "Email Address Count"; Expression = { $_.SMTPAddresses }; Alignment = "Center"}, `
        @{Name = "Group Count"; Expression = { $_.SMTPAddresses }; Alignment = "Center"} `
        -AutoSize `
        -Wrap




# Export the same report to CSV
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$csvPath = ".\OffboardingExceptions_$timestamp.csv"

$results |
Export-Csv `
    -Path $csvPath `
    -NoTypeInformation `
    -Encoding utf8 `
    -ErrorAction Stop

Write-Host ""
Write-Host "CSV file created:" -ForegroundColor Green
Write-Host $csvPath