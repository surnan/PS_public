# Get_UserSignIns.ps1
# Prompts for a user's email address and number of days.
# Returns the user's Entra ID sign-in activity during that period.
# Exports full results to a CSV file.


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
$userEmail = Read-Host "Enter user's email address"

$userEmail = $userEmail.Trim()


# Verify email format
if ($userEmail -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {

    Write-Host `
        "Invalid email format." `
        -ForegroundColor Red `
        -BackgroundColor White

    exit
}


# Prompt for number of days
$daysInput = Read-Host "Enter the number of days to search through in logs"

$days = 0

$isValid = [int]::TryParse(
    $daysInput,
    [ref]$days
)


# Verify that input is a positive integer
if (-not $isValid -or $days -le 0) {

    Write-Host `
        "Invalid input. Please enter a positive whole number." `
        -ForegroundColor Red `
        -BackgroundColor White

    exit
}


# Verify that the user exists
Write-Host `
    "`nSearching for $userEmail..." `
    -ForegroundColor Yellow


try {

    $user = Get-MgUser `
        -Filter "userPrincipalName eq '$userEmail'" `
        -Property `
            Id,
            DisplayName,
            UserPrincipalName `
        -ErrorAction Stop

}
catch {

    Write-Host `
        "`nMicrosoft Graph returned an error:" `
        -ForegroundColor Red

    Write-Host `
        $_.Exception.Message `
        -ForegroundColor Red

    exit
}


# Make sure a user was returned
if (-not $user) {

    Write-Host `
        "`nNo user was found with UPN: $userEmail" `
        -ForegroundColor Red

    exit
}


Write-Host `
    "`nUser found: $($user.DisplayName)" `
    -ForegroundColor Green


# Calculate cutoff date in UTC
$cutoffDate = (Get-Date).AddDays(-$days).ToUniversalTime()

$cutoffDateString = $cutoffDate.ToString(
    "yyyy-MM-ddTHH:mm:ssZ"
)


Write-Host `
    "`nRetrieving sign-ins from the last $days days..." `
    -ForegroundColor Yellow


# Get sign-in logs
try {

    Write-Progress `
        -Activity "Searching Entra sign-in logs" `
        -Status "Retrieving sign-in events for $userEmail..."

    $signIns = Get-MgAuditLogSignIn `
        -Filter "userPrincipalName eq '$userEmail' and createdDateTime ge $cutoffDateString" `
        -All `
        -ErrorAction Stop

    Write-Progress `
        -Activity "Searching Entra sign-in logs" `
        -Completed

}
catch {

    Write-Progress `
        -Activity "Searching Entra sign-in logs" `
        -Completed

    Write-Host `
        "`nUnable to retrieve sign-in logs." `
        -ForegroundColor Red

    Write-Host `
        $_.Exception.Message `
        -ForegroundColor Red

    exit
}


# Check for results
if (-not $signIns) {

    Write-Host `
        "`nNo sign-in activity found for $userEmail during the last $days days." `
        -ForegroundColor Yellow

    exit
}


# Build results
$results = foreach ($signIn in $signIns) {

    [PSCustomObject]@{

        SignInTime = $signIn.CreatedDateTime.ToLocalTime()

        User = $signIn.UserPrincipalName

        Application = $signIn.AppDisplayName

        IPAddress = $signIn.IpAddress

        ClientApp = $signIn.ClientAppUsed

        Device = $signIn.DeviceDetail.DisplayName

        OperatingSystem = $signIn.DeviceDetail.OperatingSystem

        Browser = $signIn.DeviceDetail.Browser

        City = $signIn.Location.City

        State = $signIn.Location.State

        Country = $signIn.Location.CountryOrRegion

        Status = if ($signIn.Status.ErrorCode -eq 0) {
            "Success"
        }
        else {
            "Failed"
        }

        FailureReason = if ($signIn.Status.ErrorCode -eq 0) {
            ""
        }
        else {
            $signIn.Status.FailureReason
        }
    }
}


# Sort results newest to oldest
$results = $results |
    Sort-Object SignInTime -Descending


# Display results on console
Write-Host `
    "`nSign-in activity for $($user.DisplayName)" `
    -ForegroundColor Cyan

Write-Host `
    "Last $days days" `
    -ForegroundColor Cyan


$results |
    Format-Table `
        SignInTime,
        Application,
        IPAddress,
        ClientApp,
        OperatingSystem,
        Status `
        -AutoSize


Write-Host `
    "`nTotal sign-in events: $($results.Count)" `
    -ForegroundColor Cyan


# Create safe version of username for CSV filename
$safeUserName = $user.UserPrincipalName `
    -replace '@', '_' `
    -replace '\.', '_'


# Create timestamp
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"


# Create CSV path
$csvPath = ".\UserSignIns_$($safeUserName)_$timestamp.csv"


# Export full PSCustomObject results to CSV
$results |
    Export-Csv `
        -Path $csvPath `
        -NoTypeInformation


Write-Host `
    "`nCSV file created: $csvPath" `
    -ForegroundColor Green