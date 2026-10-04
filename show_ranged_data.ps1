function Convert-ToDate {
    param (
        [string]$DateString
    )

    try {
        # [System.DateTime] = .NET
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


# Today's date as a DateTime object
$today = (Get-Date).Date


# Get date range from user
$startDateInput = Read-Host "Enter start date (MM-DD-YYYY) or blank for today"
$endDateInput   = Read-Host "Enter end date (MM-DD-YYYY) or blank for today"


# If input is blank, use today's date.
# Otherwise, convert the string into a DateTime object.
$startDate = [string]::IsNullOrWhiteSpace($startDateInput) `
    ? $today `
    : (Convert-ToDate $startDateInput)

$endDate = [string]::IsNullOrWhiteSpace($endDateInput) `
    ? $today `
    : (Convert-ToDate $endDateInput)


# Check for invalid date input
if ($null -eq $startDate -or $null -eq $endDate) {
    Write-Host "Invalid date value entered." -ForegroundColor Red
    exit
}


# Dates cannot be in the future
if ($startDate -gt $today -or $endDate -gt $today) {
    Write-Host "Date cannot be after today." -ForegroundColor Red
    exit
}


# End date must be the same as or later than start date
if ($endDate -lt $startDate) {
    Write-Host "End date cannot be before start date." -ForegroundColor Red
    exit
}


# Convert local date boundaries to UTC ISO 8601 format
# for Microsoft Graph / OData
$startDateOData = $startDate.ToUniversalTime().ToString(
    "yyyy-MM-ddTHH:mm:ssZ"
)

# Use the very end of the selected end date because Graph
# supports "le" for createdDateTime
$endDateOData = $endDate.Date.AddDays(1).AddTicks(-1).ToUniversalTime().ToString(
    "yyyy-MM-ddTHH:mm:ss.fffffffZ"
)


# Build OData filter
$filter = "createdDateTime ge $startDateOData and createdDateTime le $endDateOData"


# Display selected dates
Write-Host ""
Write-Host "Start Date: " -ForegroundColor Blue -NoNewline
Write-Host $startDate.ToString("MM-dd-yyyy") -ForegroundColor Green

Write-Host "End Date: " -ForegroundColor Blue -NoNewline
Write-Host $endDate.ToString("MM-dd-yyyy") -ForegroundColor Green


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
        -Property Id,DisplayName,UserPrincipalName,JobTitle,CreatedDateTime `
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


# Display and export results
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
            CreatedDateTime,
            @{Name="ReportStartDate"; Expression={$startDate.ToString("MM-dd-yyyy")}},
            @{Name="ReportEndDate"; Expression={$endDate.ToString("MM-dd-yyyy")}}


    # Print results to console
    $results |
        Format-Table `
            DisplayName,
            UserPrincipalName,
            JobTitle,
            CreatedDateTime `
            -AutoSize

    # Build CSV filename
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$csvPath = ".\CreatedUsers_$($startDate.ToString('yyyy-MM-dd'))_to_$($endDate.ToString('yyyy-MM-dd'))_$timestamp.csv"


    # Export results to CSV
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