# Get-TeamsUserMembership.ps1
# Read-only: direct teams and channels in/shared with those teams.
# Shared-channel-only access to other teams is outside this report's scope.

$userEmail = (Read-Host "Enter user's email address").Trim()

if ($userEmail -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
    Write-Host 'Invalid email format.' -ForegroundColor Red
    return
}

# Load only Authentication to reduce Graph submodule version conflicts.
try {
    if (-not (Get-Module -ListAvailable Microsoft.Graph.Authentication)) {
        Write-Host 'Installing Microsoft.Graph.Authentication...' `
            -ForegroundColor Yellow

        Install-Module Microsoft.Graph.Authentication `
            -Scope CurrentUser `
            -ErrorAction Stop
    }

    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    $scopes = @(
        'User.Read.All'
        'Team.ReadBasic.All'
        'TeamMember.Read.All'
        'Channel.ReadBasic.All'
        'ChannelMember.Read.All'
    )

    $context = Get-MgContext

    $missingScopes = @(
        $scopes |
            Where-Object { $_ -notin $context.Scopes }
    )

    if (-not $context -or $missingScopes.Count -gt 0) {
        Write-Host 'Connecting to Microsoft Graph...' `
            -ForegroundColor Yellow

        $connectionParameters = @{
            Scopes       = $scopes
            ContextScope = 'Process'
            NoWelcome    = $true
            ErrorAction  = 'Stop'
        }

        if ($context -and $context.TenantId) {
            $connectionParameters.TenantId = $context.TenantId
        }

        Connect-MgGraph @connectionParameters
    }

    $tenantId = (Get-MgContext).TenantId
}
catch {
    Write-Host "Unable to load modules or connect: $($_.Exception.Message)" `
        -ForegroundColor Red
    return
}

# Read every page before returning the complete collection.
function Get-GraphCollection {
    param(
        [string]$Uri,
        [hashtable]$Headers = @{}
    )

    $items = [System.Collections.Generic.List[object]]::new()

    do {
        $page = Invoke-MgGraphRequest `
            -Method GET `
            -Uri $Uri `
            -Headers $Headers `
            -ErrorAction Stop

        foreach ($item in $page.value) {
            $items.Add($item)
        }

        $Uri = $page.'@odata.nextLink'
    } while ($Uri)

    $items.ToArray()
}

function Get-MembershipRole {
    param($Membership)

    if ('owner' -in $Membership.roles) {
        'Owner'
    }
    elseif ('guest' -in $Membership.roles) {
        'Guest'
    }
    else {
        'Member'
    }
}

$baseUri = 'https://graph.microsoft.com/v1.0'
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

try {
    # Resolve the sign-in address first.
    $encodedEmail = [Uri]::EscapeDataString($userEmail)

    try {
        $user = Invoke-MgGraphRequest `
            -Method GET `
            -Uri "$baseUri/users/${encodedEmail}?`$select=id,displayName,userPrincipalName" `
            -ErrorAction Stop
    }
    catch {
        # Fall back to primary email and aliases.
        $escapedEmail = $userEmail.Replace("'", "''")

        $filter = (
            "mail eq '$escapedEmail' or " +
            "proxyAddresses/any(a:a eq 'smtp:$escapedEmail') or " +
            "proxyAddresses/any(a:a eq 'SMTP:$escapedEmail')"
        )

        $encodedFilter = [Uri]::EscapeDataString($filter)

        $matches = @(
            Get-GraphCollection `
                -Headers @{ ConsistencyLevel = 'eventual' } `
                -Uri "$baseUri/users?`$filter=$encodedFilter&`$count=true&`$select=id,displayName,userPrincipalName"
        )

        if ($matches.Count -ne 1) {
            throw "Expected one user; found $($matches.Count) for $userEmail."
        }

        $user = $matches[0]
    }

    Write-Host "`nRetrieving teams for $($user.displayName)..." `
        -ForegroundColor Yellow

    $teams = @(
        Get-GraphCollection `
            -Uri "$baseUri/users/$($user.id)/joinedTeams"
    )

    if ($teams.Count -eq 0) {
        Write-Host 'No direct Teams memberships found.' `
            -ForegroundColor Yellow
        return
    }

    $results = [System.Collections.Generic.List[object]]::new()

    # Add identically structured rows to the report.
    function Add-ReportRow {
        param(
            $Team,
            $TeamRole,
            $RecordType,
            $ChannelName = '',
            $ChannelType = '',
            $ChannelRole = '',
            $Access = 'Yes',
            $ChannelId = '',
            $ErrorMessage = ''
        )

        $results.Add([PSCustomObject]@{
            UserPrincipalName = $user.userPrincipalName
            TeamName          = $Team.displayName
            TeamRole          = $TeamRole
            TeamId            = $Team.id
            RecordType        = $RecordType
            ChannelName       = $ChannelName
            ChannelType       = $ChannelType
            ChannelRole       = $ChannelRole
            Access            = $Access
            ChannelId         = $ChannelId
            Error             = $ErrorMessage
        })
    }

    $index = 0

    foreach ($team in ($teams | Sort-Object displayName)) {
        $index++

        Write-Progress `
            -Activity 'Checking teams and channels' `
            -Status "$index of $($teams.Count): $($team.displayName)" `
            -PercentComplete (($index / $teams.Count) * 100)

        $teamRole = 'Unknown'
        $teamError = ''

        # Check the user's role in this team.
        try {
            $membershipFilter = [Uri]::EscapeDataString(
                "microsoft.graph.aadUserConversationMember/userId eq '$($user.id)'"
            )

            $membership = @(
                Get-GraphCollection `
                    -Uri "$baseUri/teams/$($team.id)/members?`$filter=$membershipFilter"
            )

            if ($membership.Count -ne 1) {
                throw "Expected one team membership; found $($membership.Count)."
            }

            $teamRole = Get-MembershipRole $membership[0]
        }
        catch {
            $teamError = $_.Exception.Message

            Write-Warning "$($team.displayName): $teamError"
        }

        # Keep a team row even if channel retrieval fails.
        Add-ReportRow `
            -Team $team `
            -TeamRole $teamRole `
            -RecordType 'Team' `
            -ErrorMessage $teamError

        # Includes this team's channels and incoming shared channels.
        try {
            $channels = @(
                Get-GraphCollection `
                    -Uri "$baseUri/teams/$($team.id)/allChannels?`$select=id,displayName,membershipType"
            )
        }
        catch {
            $errorMessage = $_.Exception.Message

            Write-Warning "Could not list channels for $($team.displayName): $errorMessage"

            Add-ReportRow `
                -Team $team `
                -TeamRole $teamRole `
                -RecordType 'Channel' `
                -ChannelName '(list unavailable)' `
                -ChannelRole 'Unknown' `
                -Access 'Unknown' `
                -ErrorMessage $errorMessage

            continue
        }

        foreach ($channel in ($channels | Sort-Object displayName)) {
            # Standard channels inherit team access.
            $channelRole = 'Inherited from team'
            $access = 'Yes'
            $errorMessage = ''

            # Private/shared channels require a separate membership check.
            if ($channel.membershipType -ne 'standard') {
                try {
                    # Incoming shared channels belong to their host team.
                    $channelUri = $channel.'@odata.id'

                    if ($channelUri) {
                        $channelUri = $channelUri -replace '/tenants/[^/]+', ''
                    }
                    else {
                        if ($channel.membershipType -eq 'shared') {
                            throw 'Shared channel host URL was not returned; access could not be verified.'
                        }

                        $encodedChannelId = [Uri]::EscapeDataString(
                            $channel.id
                        )

                        $channelUri = "$baseUri/teams/$($team.id)/channels/$encodedChannelId"
                    }

                    # allMembers includes direct membership and shared
                    # channel access inherited through another team.
                    $channelMembers = @(
                        Get-GraphCollection `
                            -Uri "$channelUri/allMembers"
                    )

                    $matchingMembers = @(
                        $channelMembers |
                            Where-Object {
                                $_.userId -eq $user.id -and
                                (
                                    -not $_.tenantId -or
                                    $_.tenantId -eq $tenantId
                                )
                            }
                    )

                    # Membership check succeeded, but user is not a member.
                    # Do not include this channel.
                    if ($matchingMembers.Count -eq 0) {
                        continue
                    }

                    $channelRole = 'Member'

                    $ownerMemberships = @(
                        $matchingMembers |
                            Where-Object { 'owner' -in $_.roles }
                    )

                    $guestMemberships = @(
                        $matchingMembers |
                            Where-Object { 'guest' -in $_.roles }
                    )

                    if ($ownerMemberships.Count -gt 0) {
                        $channelRole = 'Owner'
                    }
                    elseif ($guestMemberships.Count -gt 0) {
                        $channelRole = 'Guest'
                    }
                }
                catch {
                    # A failed check does not prove the user lacks access.
                    $channelRole = 'Unknown'
                    $access = 'Unknown'
                    $errorMessage = $_.Exception.Message

                    Write-Warning "$($team.displayName) / $($channel.displayName): $errorMessage"
                }
            }

            Add-ReportRow `
                -Team $team `
                -TeamRole $teamRole `
                -RecordType 'Channel' `
                -ChannelName $channel.displayName `
                -ChannelType $channel.membershipType `
                -ChannelRole $channelRole `
                -Access $access `
                -ChannelId $channel.id `
                -ErrorMessage $errorMessage
        }
    }

    Write-Progress -Activity 'Checking teams and channels' -Completed

    # Console: team names and roles.
    Write-Host "`nTeams for $($user.displayName):" `
        -ForegroundColor Cyan

    $results |
        Where-Object RecordType -eq 'Team' |
        Format-Table TeamName, TeamRole -AutoSize -Wrap

    # Console: channel names, types, roles, and verified access.
    Write-Host 'Channels (Unknown means access could not be verified):' `
        -ForegroundColor Cyan

    $results |
        Where-Object RecordType -eq 'Channel' |
        Format-Table `
            TeamName, ChannelName, ChannelType, ChannelRole, Access `
            -AutoSize -Wrap

    # Export the complete report.
    $outputDirectory = if ($PSScriptRoot) {
        $PSScriptRoot
    }
    else {
        (Get-Location).Path
    }

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss_fff'

    $safeUserName = $user.userPrincipalName -replace '[^a-zA-Z0-9._-]', '_'

    $csvPath = Join-Path $outputDirectory `
        "TeamsChannels_${safeUserName}_$timestamp.csv"

    $results |
        Export-Csv `
            -Path $csvPath `
            -NoTypeInformation `
            -Encoding UTF8 `
            -ErrorAction Stop

    Write-Host "`nCSV saved: $csvPath" -ForegroundColor Green

    Write-Host "Direct teams found: $($teams.Count)" `
        -ForegroundColor Cyan
}
catch {
    Write-Host "Unable to complete report: $($_.Exception.Message)" `
        -ForegroundColor Red
}
finally {
    Write-Progress -Activity 'Checking teams and channels' -Completed

    $stopwatch.Stop()

    Write-Host "`nTotal execution time: $($stopwatch.Elapsed)" `
        -ForegroundColor Cyan
}