# Connect to Microsoft Graph if necessary
$context = Get-MgContext
if (-not $context) {
    Write-Host "Not connected to Microsoft Graph. Connecting..." -ForegroundColor Gray
    Connect-MgGraph -Scopes `
        "User.Read.All", `
        "Group.Read.All",`
        "LicenseAssignment.Read.All"
}


# -"accountEnabled eq false"  = users w/ sign-in disabled.
$disabledUsers = Get-MgUser `
    -Filter "accountEnabled eq false" `
    -Property DisplayName,UserPrincipalName,AccountEnabled, ProxyAddresses `
    -All

foreach ($user in $disabledUsers) {
    Write-Host "`nUser: $($user.DisplayName)" -ForegroundColor Cyan
    Write-Host "Sign-in name: $($user.UserPrincipalName)"

    $user.ProxyAddresses |
        Where-Object { $_ -like "smtp:*" } |
        ForEach-Object { $_ -replace "^smtp:", "" }
}

    

# Display the results
$disabledUsers |
    Format-Table DisplayName,UserPrincipalName,AccountEnabled -AutoSize


$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$csvPath = ".\OffboardingExceptions_$timestamp.csv"

foreach ($user in $disabledUsers) {
    "$($user.UserPrincipalName)" | Out-File `
                -FilePath $csvPath `
                -Append `
                -Encoding utf8
}




Write-Host ""
Write-Host "CSV file created:" -ForegroundColor Green
Write-Host $csvPath