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
#################### MICROSOFT GRAPH CONNECTION ####################################
####################################################################################
$modules = @(
    "Microsoft.Graph.Users"
)

$requiredScopes = @("User.Read.All")

try {
    foreach ($module in $modules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            Install-Module $module -Scope CurrentUser -ErrorAction Stop
        }

        Import-Module $module -ErrorAction Stop
    }

    $connection = Get-MgContext

    if (-not $connection) {
        Connect-MgGraph `
            -Scopes $requiredScopes `
            -ContextScope CurrentUser `
            -NoWelcome `
            -ErrorAction Stop

        $connection = Get-MgContext
    }

    if ($connection.AuthType -ne "Delegated") {
        throw "A delegated Microsoft Graph connection is required."
    }

    $missingScopes = @(
        $requiredScopes | Where-Object {
            $_ -notin @($connection.Scopes)
        }
    )

    if ($missingScopes.Count -gt 0) {
        throw "Missing Graph scopes: $($missingScopes -join ', ')"
    }
}
catch {
    Write-Host "Microsoft Graph connection failed:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    return
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

# Microsoft Graph accepts "le" rather than "lt" for this date filter
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
    Write-Host "Microsoft Graph error (1): $($_.Exception.Message)" -ForegroundColor Red
    return
}


####################################################################################
#################### BUILD RESULTS OBJECT ##########################################
####################################################################################
$results = [PSCustomObject]@{
    startDate = $startDate
    endDate   = $endDate
    allUsers  = @($users | Sort-Object CreatedDateTime)
}


####################################################################################
#################### CONSOLE OUTPUT ################################################
####################################################################################
Write-Host "`nUser accounts created between " -ForegroundColor Blue -NoNewline
Write-Host $results.startDate -ForegroundColor Green -NoNewline
Write-Host " and " -ForegroundColor Blue -NoNewline
Write-Host $results.endDate -ForegroundColor Green

Write-Host "`nTotal users found: $($results.allUsers.Count)" -ForegroundColor Yellow
Write-Host ""

if ($results.allUsers.Count -gt 0) {
    $results.allUsers |
    Format-Table DisplayName, UserPrincipalName, JobTitle, CreatedDateTime -AutoSize |
    Out-Host
}
else {
    Write-Host "`nNo user accounts were created within the specified date range." -ForegroundColor Yellow
}


####################################################################################
#################### CREATE CSV ####################################################
####################################################################################
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$startString = $results.startDate.ToString('yyyy-MM-dd')
$endString = $results.endDate.ToString('yyyy-MM-dd')
$csvPath = ".\EntraUsersByCreatedDate_${startString}_to_${endString}_$timestamp.csv"

try {
    $results.allUsers |
    Export-Csv `
        -Path $csvPath `
        -NoTypeInformation `
        -Encoding utf8 `
        -ErrorAction Stop
}
catch {
    Write-Host "Unable to create CSV file (2): $($_.Exception.Message)" -ForegroundColor Red
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