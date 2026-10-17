# Get_UserSignIns.ps1
# Input: User email address (UPN) and number of days
#
# Output:
#   Sign-in time
#   User
#   Application
#   IP address
#   Client application
#   Device information
#   Location
#   Sign-in status and failure reason
#
# Export:
#   CSV: Full sign-in results



####################################################################################
#################### MICROSOFT GRAPH CONNECTION ####################################
####################################################################################
$modules = @(
    "Microsoft.Graph.Users"
)

$requiredScopes = @("User.Read.All", "Group.Read.All")

try {
    # Install and import required modules
    foreach ($module in $modules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            Install-Module $module -Scope CurrentUser -ErrorAction Stop
        }

        Import-Module $module -ErrorAction Stop
    }

    # Check existing connection and permissions
    $connection = Get-MgContext

    $missingScopes = @(
        $requiredScopes | Where-Object {
            $_ -notin @($connection.Scopes)
        }
    )

    # Connect if disconnected or missing required permissions
    if (-not $connection -or
        $connection.AuthType -ne "Delegated" -or
        $missingScopes.Count -gt 0) {

        Connect-MgGraph `
            -Scopes $requiredScopes `
            -ContextScope CurrentUser `
            -NoWelcome `
            -ErrorAction Stop

        $connection = Get-MgContext
    }

    # Verify authentication type
    if (-not $connection -or $connection.AuthType -ne "Delegated") {
        throw "A delegated Microsoft Graph connection is required."
    }

    # Verify required permissions
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
$userEmail = (Read-Host "Enter user's email address").Trim()

if ($userEmail -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
    Write-Host "Invalid email format (2):" -ForegroundColor Red
    return
}

$daysInput = Read-Host "Enter the number of days to search through in logs"

$days = 0

$isValid = [int]::TryParse($daysInput, [ref]$days)

if (-not $isValid -or $days -le 0) {
    Write-Host "Invalid input. Please enter a positive integer (3):" -ForegroundColor Red
    return
}


####################################################################################
#################### FIND TARGET USER ##############################################
####################################################################################
Write-Host "`nSearching for $userEmail..." -ForegroundColor Yellow

try {
    $user = Get-MgUser `
        -Filter "userPrincipalName eq '$userEmail'" `
        -Property Id, DisplayName, UserPrincipalName `
        -ErrorAction Stop
}
catch {
    Write-Host "Microsoft Graph returned an error (4):" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    return
}

if (-not $user) {
    Write-Host "`nNo user was found with UPN (5): $userEmail" -ForegroundColor Red
    return
}

Write-Host "`nUser found: $($user.DisplayName)" -ForegroundColor Green


####################################################################################
#################### BUILD DATE FILTER #############################################
####################################################################################
$cutoffDate = (Get-Date).AddDays(-$days).ToUniversalTime()
$cutoffDateString = $cutoffDate.ToString("yyyy-MM-ddTHH:mm:ssZ")

Write-Host "`nRetrieving sign-ins from the last $days days..." -ForegroundColor Yellow


####################################################################################
#################### GET SIGN-IN LOGS ##############################################
####################################################################################
try {
    Write-Progress `
        -Activity "Searching Entra sign-in logs" `
        -Status "Retrieving sign-in events for $userEmail..."

    $signIns = @(
        Get-MgAuditLogSignIn `
            -Filter "userPrincipalName eq '$userEmail' and createdDateTime ge $cutoffDateString" `
            -All `
            -ErrorAction Stop
    )
}
catch {
    Write-Host "Unable to retrieve sign-in logs (6):" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    return
}
finally {
    Write-Progress `
        -Activity "Searching Entra sign-in logs" `
        -Completed
}


####################################################################################
#################### CHECK RESULTS #################################################
####################################################################################
if ($signIns.Count -eq 0) {
    Write-Host "`nNo sign-in activity found for $userEmail during the last $days days (7):" -ForegroundColor Yellow
    return
}


####################################################################################
#################### BUILD RESULTS OBJECT ##########################################
####################################################################################
$results = @(
    foreach ($signIn in $signIns) {
        [PSCustomObject]@{
            SignInTime      = $signIn.CreatedDateTime.ToLocalTime()
            User            = $signIn.UserPrincipalName
            Application     = $signIn.AppDisplayName
            IPAddress       = $signIn.IpAddress
            ClientApp       = $signIn.ClientAppUsed
            Device          = $signIn.DeviceDetail.DisplayName
            OperatingSystem = $signIn.DeviceDetail.OperatingSystem
            Browser         = $signIn.DeviceDetail.Browser
            City            = $signIn.Location.City
            State           = $signIn.Location.State
            Country         = $signIn.Location.CountryOrRegion

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
)

$results = @($results | Sort-Object SignInTime -Descending)


####################################################################################
#################### CONSOLE OUTPUT ################################################
####################################################################################
Write-Host "`nSign-in activity for $($user.DisplayName)" -ForegroundColor Cyan
Write-Host "Last $days days" -ForegroundColor Cyan

$results |
Format-Table `
    SignInTime,
    Application,
    IPAddress,
    ClientApp,
    OperatingSystem,
    Status `
    -AutoSize |
Out-Host

Write-Host "`nTotal sign-in events: $($results.Count)" -ForegroundColor Cyan


####################################################################################
#################### CREATE CSV ####################################################
####################################################################################
$safeUserName = $user.UserPrincipalName `
    -replace '@', '_' `
    -replace '\.', '_'

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

$csvPath = ".\UserSignIns_$($safeUserName)_$timestamp.csv"

try {
    $results |
    Export-Csv `
        -Path $csvPath `
        -NoTypeInformation `
        -Encoding utf8 `
        -ErrorAction Stop
}
catch {
    Write-Host "Unable to create CSV file: (8)" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    return
}

Write-Host "`nCSV file created: $csvPath" -ForegroundColor Green
