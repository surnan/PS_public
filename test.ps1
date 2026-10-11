# Connect to Microsoft Graph if necessary
$context = Get-MgContext

if (-not $context) {
    Write-Host "Not connected to Microsoft Graph. Connecting..." -ForegroundColor Gray

    Connect-MgGraph -Scopes `
        "User.Read.All", `
        "Group.Read.All", `
        "LicenseAssignment.Read.All"
}

# Get users with sign-in disabled
$disabledUsers = Get-MgUser `
    -Filter "accountEnabled eq false" `
    -Property Id,DisplayName,UserPrincipalName,AccountEnabled,ProxyAddresses,AssignedLicenses `
    -All `
    -ErrorAction Stop

# Store separate rows for the console
$consoleResults = @()

Write-Host "Loading accounts" -ForegroundColor Yellow -NoNewline

# Build the CSV report and console report
$results = @(
    foreach ($user in $disabledUsers) {

        # Get all SMTP addresses, including the primary email address
        $smtpAddresses = @(
            $user.ProxyAddresses |
                Where-Object { $_ -like "smtp:*" } |
                ForEach-Object { $_ -replace "^smtp:", "" }
        )

        # CSV: show one address, otherwise show the count
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

        # CSV: show one group name, otherwise show the count
        if ($groups.Count -eq 1) {
            $groupDisplay = $groups[0].DisplayName
        }
        else {
            $groupDisplay = $groups.Count
        }

        # Count assigned product licenses
        $licenseCount = @($user.AssignedLicenses).Count

        # Console: always show numeric counts
        $consoleResults += [PSCustomObject]@{
            DisplayName       = $user.DisplayName
            UserPrincipalName = $user.UserPrincipalName
            AccountEnabled    = $user.AccountEnabled
            EmailAddressCount = $smtpAddresses.Count
            GroupCount        = $groups.Count
            LicenseCount      = $licenseCount
        }

        # CSV: preserve the existing report values
        $result = [PSCustomObject]@{
            DisplayName       = $user.DisplayName
            UserPrincipalName = $user.UserPrincipalName
            AccountEnabled    = $user.AccountEnabled
            SMTPAddresses     = $aliasDisplay
            Groups            = $groupDisplay
            LicenseCount      = $licenseCount
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

# Sort both reports alphabetically by display name
$consoleResults = @($consoleResults | Sort-Object DisplayName)
$results = @($results | Sort-Object DisplayName)

# Display numeric counts in the console
$consoleResults |
    Format-Table `
        DisplayName, `
        UserPrincipalName, `
        @{Name = "[Email Address #]"; Expression = { $_.EmailAddressCount }; Alignment = "Center"}, `
        @{Name = "[Group #]"; Expression = { $_.GroupCount }; Alignment = "Center"}, `
        @{Name = "[License #]"; Expression = { $_.LicenseCount }; Alignment = "Center"}, `
        AccountEnabled `
        -AutoSize `
        -Wrap

# Export the existing detailed report to CSV
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