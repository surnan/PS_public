function Convert_To_Date {
    param (
        [string]$DateString
    )

    try {
        # [System.DateTime] from .NET
        return [System.DateTime]::ParseExact(
            $DateString,
            "MM-dd-yyyy",
            $null
        )
    }
    catch {
        return $null
    }
}


$today = (Get-Date).Date
$startDateInput = Read-Host "Enter start date (MM-DD-YYYY) or blank for today"
$endDateInput   = Read-Host "Enter end date (MM-DD-YYYY) or blank for today"

# String has value --> DateTime object.
# String is blank --> use current date.
$startDate = [System.String]::IsNullOrWhiteSpace($startDateInput) `
    ? $today `
    : (Convert_To_Date $startDateInput)

$endDate = [System.String]::IsNullOrWhiteSpace($endDateInput) `
    ? $today `
    : (Convert_To_Date $endDateInput)


# Verify valid values for Start & End Dates
if ($null -eq $startDate -or $null -eq $endDate) {
    Write-Host "Invalid date value entered." -ForegroundColor Red
    exit
}


# Don't allow future dates
if ($startDate -gt $today -or $endDate -gt $today) {
    Write-Host "Date cannot be after today." -ForegroundColor Red
    exit
}


# Verify Start Date not after End date
if ($endDate -lt $startDate) {
    Write-Host "End date cannot be before start date." -ForegroundColor Red
    exit
}


# Convert date into UTC ISO 8601 format for OData query
$startDateOData = $startDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")

# Forward to next day start and go backwards in smallest time unit
# Forced to use Less than or Equal because Less than isn't available
$endDateOData = $endDate.Date.AddDays(1).AddTicks(-1).ToUniversalTime().ToString(
    "yyyy-MM-ddTHH:mm:ss.fffffffZ"
)


# Building OData filter
$filter = "createdDateTime ge $startDateOData and createdDateTime le $endDateOData"


# Display selected dates
# Write-Host ""
# Write-Host "Start Date: " -ForegroundColor Blue -NoNewline
# Write-Host $startDate.ToString("MM-dd-yyyy") -ForegroundColor Green
# Write-Host "End Date: " -ForegroundColor Blue -NoNewline
# Write-Host $endDate.ToString("MM-dd-yyyy") -ForegroundColor Green


# Connect to Microsoft Graph if necessary
$context = Get-MgContext

if (-not $context) {
    Write-Host ""
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
    Connect-MgGraph -Scopes "User.Read.All"
}


# Query Entra users
try {
    $users = Get-MgUser `
        -Filter $filter `
        -Property DisplayName,UserPrincipalName,JobTitle,CreatedDateTime `
        -All `
        -ErrorAction Stop
}
catch {
    Write-Host ""
    Write-Host "Microsoft Graph query failed." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host $_.InvocationInfo.PositionMessage -ForegroundColor Yellow
    exit
}

Write-Host ""

if ($users) {
    Write-Host "User accounts created between " -ForegroundColor Blue -NoNewline
    Write-Host $startDate.ToString("MM-dd-yyyy") -ForegroundColor Green -NoNewline
    Write-Host " and " -ForegroundColor Blue -NoNewline
    Write-Host $endDate.ToString("MM-dd-yyyy") -ForegroundColor Green

    Write-Host ""
    Write-Host "Total users found: $($users.Count)" -ForegroundColor Yellow
    Write-Host ""


    # Create a reusable result set for both console output and CSV export
    $results = $users |
        Sort-Object CreatedDateTime |
        Select-Object `
            DisplayName,
            UserPrincipalName,
            JobTitle,
            CreatedDateTime


    # Print results to console
    $results |
        Format-Table `
            DisplayName,
            UserPrincipalName,
            CreatedDateTime `
            -AutoSize

    # Build CSV filename
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csvPath = ".\EntraUsersByCreatedDate_$($startDate.ToString('yyyy-MM-dd'))_to_" `
                + "$($endDate.ToString('yyyy-MM-dd'))_$timestamp.csv"


    # Export-Csv is PowerShell cmdlet
    $results |
        Export-Csv `
            -Path $csvPath `
            -NoTypeInformation


    Write-Host ""
    Write-Host "CSV created: " -ForegroundColor Blue -NoNewline
    Write-Host $csvPath -ForegroundColor Green
}
else {

    Write-Host "No user accounts were created during this date range." -ForegroundColor Yellow
}