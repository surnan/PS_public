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
    -Property DisplayName,UserPrincipalName,AccountEnabled `
    -All

# Display the results
$disabledUsers |
    Format-Table DisplayName,UserPrincipalName,AccountEnabled -AutoSize
