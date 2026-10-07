# Get-StaleUser.ps1
# Prompts for a user's email address and number of days.
# Checks whether the user's account is stale.


# Connect to Microsoft Graph if necessary
$context = Get-MgContext

if (-not $context) {

    Write-Host "Not connected to Microsoft Graph. Connecting..." `
        -ForegroundColor Gray

    Connect-MgGraph -Scopes `
        "User.Read.All", `
        "AuditLog.Read.All"
}


# Prompt for user's email address
$userEmail = Read-Host "Enter the user's email address"

$userEmail = $userEmail.Trim()


# Prompt for number of days
do {

    $daysInput = Read-Host `
        "Enter the number of days to search back (example: 60 or 90)"

    $days = 0

    $isValid = [int]::TryParse(
        $daysInput,
        [ref]$days
    )

    if (-not $isValid -or $days -le 0) {

        Write-Host `
            "Please enter a positive whole number." `
            -ForegroundColor Red
    }

}
until ($isValid -and $days -gt 0)


# Calculate cutoff date
$cutoffDate = (Get-Date).AddDays(-$days)


Write-Host `
    "`nSearching for $userEmail..." `
    -ForegroundColor Yellow


# Find user
try {

    $user = Get-MgUser `
        -Filter "userPrincipalName eq '$userEmail'" `
        -Property `
            Id,
            DisplayName,
            UserPrincipalName,
            Mail,
            AccountEnabled,
            CreatedDateTime,
            AssignedLicenses,
            SignInActivity `
        -ErrorAction Stop

}
catch {

    Write-Host "`nMicrosoft Graph returned an error:" `
        -ForegroundColor Red

    Write-Host $_.Exception.Message `
        -ForegroundColor Red

    return
}


# Check whether a user was returned
if (-not $user) {

    Write-Host `
        "`nNo user was found with UPN: $userEmail" `
        -ForegroundColor Red

    return
}


# Get last sign-in
$lastSignIn = $user.SignInActivity.LastSignInDateTime


# Determine stale status
if ($lastSignIn) {

    $isStale = $lastSignIn -lt $cutoffDate

}
else {

    # Never signed in.
    # Only stale if account itself is older than cutoff.
    $isStale = $user.CreatedDateTime -lt $cutoffDate
}


# Create result
$result = [PSCustomObject]@{

    DisplayName = $user.DisplayName

    UserPrincipalName = $user.UserPrincipalName

    LastSignIn = if ($lastSignIn) {

        $lastSignIn.ToLocalTime()

    }
    else {

        "Never"

    }

    CreatedDate = $user.CreatedDateTime.ToLocalTime()

    LicenseCount = $user.AssignedLicenses.Count

    Stale = if ($isStale) {

        "Yes"

    }
    else {

        "No"

    }
}


# Display result
Write-Host "`nAccount Details" `
    -ForegroundColor Cyan

Write-Host "---------------" `
    -ForegroundColor Cyan


$result |
    Format-Table `
        DisplayName,
        UserPrincipalName,
        LastSignIn,
        CreatedDate,
        LicenseCount,
        Stale `
        -AutoSize


# Display final stale status
if ($isStale) {

    Write-Host `
        "`nSTALE: This account has not signed in within the last $days days." `
        -ForegroundColor Red

}
else {

    Write-Host `
        "`nACTIVE: This account does not meet the $days-day stale threshold." `
        -ForegroundColor Green
}