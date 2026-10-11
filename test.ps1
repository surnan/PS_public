
# Get_EntraUsersByCreatedDate.ps1
# Input: Start Date and End Date (MM-DD-YYYY), blank = Today
#
# Output:
#   Display Name
#   UPN
#   Job Title
#   Created Date/Time
#
# Export:
#   CSV: Same Output


####################################################################################
#################### All Microsoft Graph Powershell Scripts START ##################
####################################################################################
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -ErrorAction Stop
}
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

$connection = Get-MgContext
if (-not $connection) {
    Connect-MgGraph `
        -Scopes "User.Read.All"`
        -NoWelcome `
        -ErrorAction Stop
}


####################################################################################
#################### GET INPUT PARAMETERS ##########################################
####################################################################################
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$today = (Get-Date).Date

function Read_Date {
    param ([string]$Prompt)
    $inputDate = Read-Host $Prompt

    if ([string]::IsNullOrWhiteSpace($inputDate)) {
        return $today
    }

    try {
        [datetime]::ParseExact($inputDate, "MM-dd-yyyy", $null)
    }
    catch {
        Write-Host "Invalid date value entered. Use MM-DD-YYYY." -ForegroundColor Red
        return $null
    }
}

$startDate = Read_Date "Enter start date (MM-DD-YYYY) or blank for today"
if ($null -eq $startDate) { return }

$endDate = Read_Date "Enter end date (MM-DD-YYYY) or blank for today"
if ($null -eq $endDate) { return }

if ($startDate -gt $today -or $endDate -gt $today) {
    Write-Host "Date cannot be after today." -ForegroundColor Red
    return
}

if ($endDate -lt $startDate) {
    Write-Host "End date cannot be before start date." -ForegroundColor Red
    return
}


####################################################################################
#################### BUILD ODATA FILTER ############################################
####################################################################################
$startDateOData = $startDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")

# Convoluted because Microsoft Graph accepts "le" and not "lt" for date comparisons
$endDateOData = $endDate.Date.AddDays(1).AddTicks(-1).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ")

$filter = "createdDateTime ge $startDateOData and createdDateTime le $endDateOData"


####################################################################################
#################### FIND TARGET ACCOUNTS ##########################################
####################################################################################
try {
    $users = @(
        Get-MgUser `
            -Filter $filter `
            -Property DisplayName, UserPrincipalName, JobTitle, CreatedDateTime `
            -All `
            -ErrorAction Stop
    )
}
catch {
    Write-Host "Microsoft Graph error: $($_.Exception.Message)" -ForegroundColor Red
    return
}


####################################################################################
#################### BUILD RESULTS OBJECT ##########################################
####################################################################################
$results = @(
    $users |
    Sort-Object CreatedDateTime
)


####################################################################################
#################### CONSOLE OUTPUT ################################################
####################################################################################
Write-Host "`nUser accounts created between " -ForegroundColor Blue -NoNewline
Write-Host $startDate.ToString("MM-dd-yyyy") -ForegroundColor Green -NoNewline
Write-Host " and " -ForegroundColor Blue -NoNewline
Write-Host $endDate.ToString("MM-dd-yyyy") -ForegroundColor Green

Write-Host "`nTotal users found: $($results.Count)" -ForegroundColor Yellow

if ($results.Count -gt 0) {
    Write-Host ""

    # NOT three arrays
    # '@'{} === three hashtables that expose width column
    $results |
    Format-Table `
        @{Label="DisplayName"; Expression={$_.DisplayName}; Width=30},
        @{Label="UserPrincipalName"; Expression={$_.UserPrincipalName}; Width=60},
        @{Label="CreatedDateTime"; Expression={$_.CreatedDateTime}; Width=23} |
    Out-Host
}
else {
    Write-Host "`nNo user accounts were created within the specified date range." -ForegroundColor Yellow
}


####################################################################################
#################### CREATE CSV ####################################################
####################################################################################
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

$csvPath = ".\EntraUsersByCreatedDate_$($startDate.ToString('yyyy-MM-dd'))_to_$($endDate.ToString('yyyy-MM-dd'))_$timestamp.csv"

try {
    $results |
    Export-Csv `
        -Path $csvPath `
        -NoTypeInformation `
        -Encoding utf8 `
        -ErrorAction Stop
}
catch {
    Write-Host "Unable to create CSV file: $($_.Exception.Message)" -ForegroundColor Red
    return
}


####################################################################################
#################### EXECUTION TIME ################################################
####################################################################################
$stopwatch.Stop()

Write-Host ("`n{0,-25}" -f "File Name:") -ForegroundColor DarkCyan -NoNewline
Write-Host "Get_EntraUsersByCreatedDate.ps1" -ForegroundColor Cyan

Write-Host ("{0,-25}" -f "CSV file generated:") -ForegroundColor DarkYellow -NoNewline
Write-Host (Resolve-Path $csvPath) -ForegroundColor DarkYellow

Write-Host ("{0,-25}" -f "Total Execution Time:") -ForegroundColor DarkGreen -NoNewline
Write-Host $stopwatch.Elapsed -ForegroundColor DarkGreen
