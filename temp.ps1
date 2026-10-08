# Get-TeamsUserMembership.ps1
# Lists direct Teams memberships and Owner/Member/Guest roles.
# Does not add users or change any memberships.
# Does not report individual channel memberships.

$userEmail = (Read-Host "Enter user's email address").Trim()

# Verify email format
if ($userEmail -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
    Write-Host "Invalid email format." -ForegroundColor Red
    return
}

# Install/load required modules, then connect if necessary
try {
    $modules = @(
        'Microsoft.Graph.Authentication'
        'Microsoft.Graph.Users'
        'Microsoft.Graph.Teams'
    )

    foreach ($module in $modules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            Write-Host "Installing $module..." -ForegroundColor Yellow

            Install-Module $module `
                -Scope CurrentUser `
                -ErrorAction Stop
        }

        # Load Authentication before the other Graph modules
        Import-Module $module -ErrorAction Stop
    }

    $requiredScopes = @(
        'User.Read.All'
        'Team.ReadBasic.All'
        'TeamMember.Read.All'
    )

    $context = Get-MgContext

    $missingScopes = @(
        $requiredScopes |
            Where-Object { $_ -notin $context.Scopes }
    )

    if (-not $context -or $missingScopes.Count -gt 0) {
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow

        $connectionParameters = @{
            Scopes       = $requiredScopes
            ContextScope = 'Process'
            NoWelcome    = $true
            ErrorAction  = 'Stop'
        }

        # Reuse the existing tenant when requesting additional scopes
        if ($context -and $context.TenantId) {
            $connectionParameters.TenantId = $context.TenantId
        }

        Connect-MgGraph @connectionParameters
    }
}
catch {
    Write-Host "Unable to load modules or connect: $($_.Exception.Message)" `
        -ForegroundColor Red
    return
}

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

try {
    # Resolve the user by sign-in address first
    try {
        $userAccount = Get-MgUser `
            -UserId $userEmail `
            -Property Id, DisplayName, UserPrincipalName `
            -ErrorAction Stop
    }
    catch {
        # Check primary email and aliases if the UPN lookup fails
        $escapedEmail = $userEmail.Replace("'", "''")

        $filter = (
            "mail eq '$escapedEmail' or " +
            "proxyAddresses/any(a:a eq 'smtp:$escapedEmail') or " +
            "proxyAddresses/any(a:a eq 'SMTP:$escapedEmail')"
        )

        $matches = @(
            Get-MgUser `
                -Filter $filter `
                -Property Id, DisplayName, UserPrincipalName `
                -ConsistencyLevel eventual `
                -CountVariable matchCount `
                -All `
                -ErrorAction Stop
        )

        if ($matches.Count -ne 1) {
            Write-Host "Expected one user; found $($matches.Count) for $userEmail." `
                -ForegroundColor Red
            return
        }

        $userAccount = $matches[0]
    }

    Write-Host "`nRetrieving teams for $($userAccount.DisplayName)..." `
        -ForegroundColor Yellow

    $teams = @(
        Get-MgUserJoinedTeam `
            -UserId $userAccount.Id `
            -All `
            -ErrorAction Stop
    )

    if ($teams.Count -eq 0) {
        Write-Host "No direct Teams memberships found." -ForegroundColor Yellow
        return
    }

    $index = 0

    $results = @(
        foreach ($team in ($teams | Sort-Object DisplayName)) {
            $index++

            Write-Progress `
                -Activity "Checking Teams membership roles" `
                -Status "$index of $($teams.Count): $($team.DisplayName)" `
                -PercentComplete (($index / $teams.Count) * 100)

            $role = 'Unknown'
            $status = 'OK'
            $errorMessage = ''

            try {
                # Retrieve only the entered user's membership record
                $membership = @(
                    Get-MgTeamMember `
                        -TeamId $team.Id `
                        -Filter "microsoft.graph.aadUserConversationMember/userId eq '$($userAccount.Id)'" `
                        -All `
                        -ErrorAction Stop
                )

                if ($membership.Count -ne 1) {
                    throw "Expected one membership record; found $($membership.Count)."
                }

                if ('owner' -in $membership[0].Roles) {
                    $role = 'Owner'
                }
                elseif ('guest' -in $membership[0].Roles) {
                    $role = 'Guest'
                }
                else {
                    $role = 'Member'
                }
            }
            catch {
                $status = 'Role lookup failed'
                $errorMessage = $_.Exception.Message

                Write-Warning "Could not check role for $($team.DisplayName): $errorMessage"
            }

            [PSCustomObject]@{
                DisplayName       = $userAccount.DisplayName
                UserPrincipalName = $userAccount.UserPrincipalName
                TeamName          = $team.DisplayName
                Role              = $role
                TeamId            = $team.Id
                Status            = $status
                Error             = $errorMessage
            }
        }
    )

    Write-Progress -Activity "Checking Teams membership roles" -Completed

    # Display team names and roles in the console
    Write-Host "`nTeams for $($userAccount.DisplayName):" -ForegroundColor Cyan

    $results |
        Format-Table TeamName, Role, Status -AutoSize -Wrap

    Write-Host "Total teams: $($results.Count)" -ForegroundColor Cyan

    # Export detailed results alongside the script.
    # If pasted into the terminal, use the current directory.
    $outputDirectory = if ($PSScriptRoot) {
        $PSScriptRoot
    }
    else {
        (Get-Location).Path
    }

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
    $safeUserName = $userAccount.UserPrincipalName -replace '[^a-zA-Z0-9._-]', '_'

    $csvPath = Join-Path $outputDirectory `
        "TeamsMembership_${safeUserName}_$timestamp.csv"

    try {
        $results |
            Export-Csv `
                -Path $csvPath `
                -NoTypeInformation `
                -Encoding UTF8 `
                -ErrorAction Stop

        Write-Host "`nCSV saved: $csvPath" -ForegroundColor Green
    }
    catch {
        Write-Warning "Unable to export CSV: $($_.Exception.Message)"
    }
}
catch {
    Write-Host "Unable to complete report: $($_.Exception.Message)" `
        -ForegroundColor Red
}
finally {
    Write-Progress -Activity "Checking Teams membership roles" -Completed

    $stopwatch.Stop()

    Write-Host "`nTotal execution time: $($stopwatch.Elapsed)" `
        -ForegroundColor Cyan
}