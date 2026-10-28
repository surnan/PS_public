# Get_EntraGroup_Members.ps1
# Read-only report for an Entra ID group and its direct members.
# Installs/imports required Graph modules, exports a timestamped CSV.
# Compatible with Windows PowerShell 5.1 and PowerShell 7.

$ErrorActionPreference = 'Stop'
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

####################################################################################
#################### MICROSOFT GRAPH CONNECTION ####################################
####################################################################################

$modules = @(
    'Microsoft.Graph.Groups',
    'Microsoft.Graph.Users'
)

try {
    foreach ($module in $modules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            Write-Host "Installing missing module: $module" -ForegroundColor Yellow
            if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
                Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser | Out-Null
            }
            Install-Module -Name $module -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop
        }
        Import-Module -Name $module -ErrorAction Stop
    }

    $requiredCommands = @('Connect-MgGraph', 'Get-MgContext', 'Get-MgGroup', 'Get-MgGroupMember', 'Get-MgUser')
    foreach ($command in $requiredCommands) {
        if (-not (Get-Command -Name $command -ErrorAction SilentlyContinue)) {
            throw "Required Microsoft Graph command '$command' is unavailable after importing modules."
        }
    }

    $requiredScopes = @('Group.Read.All', 'User.Read.All', 'Directory.Read.All')
    $context = Get-MgContext
    $missingScopes = @($requiredScopes | Where-Object { -not $context -or $_ -notin @($context.Scopes) })

    if (-not $context -or $missingScopes.Count -gt 0) {
        Write-Host 'Connecting to Microsoft Graph...' -ForegroundColor Cyan
        Connect-MgGraph -Scopes $requiredScopes -NoWelcome -ErrorAction Stop
    }

    $context = Get-MgContext
    if (-not $context) { throw 'No active Microsoft Graph connection.' }
    $missingScopes = @($requiredScopes | Where-Object { $_ -notin @($context.Scopes) })
    if ($missingScopes.Count -gt 0) {
        throw "Missing Microsoft Graph permissions: $($missingScopes -join ', ')"
    }
}
catch {
    Write-Host "Microsoft Graph setup/connection failed: $($_.Exception.Message)" -ForegroundColor Red
    return
}

####################################################################################
#################### INPUT / GROUP LOOKUP ##########################################
####################################################################################

$groupInput = (Read-Host "Enter group's primary email address").Trim()
if ([string]::IsNullOrWhiteSpace($groupInput)) {
    Write-Host 'No group address entered.' -ForegroundColor Red
    return
}

try {
    # Escape single quotes for Microsoft Graph OData filtering.
    $safeAddress = $groupInput.Replace("'", "''")
    $groups = @(Get-MgGroup -Filter "mail eq '$safeAddress'" -Property 'id,displayName,mail,mailEnabled,securityEnabled,groupTypes,proxyAddresses' -All -ErrorAction Stop)
    if ($groups.Count -eq 0) {
        Write-Host "No group found with primary email address: $groupInput" -ForegroundColor Yellow
        return
    }
    $group = $groups[0]
    $members = @(Get-MgGroupMember -GroupId $group.Id -All -ErrorAction Stop)
}
catch {
    Write-Host "Unable to retrieve group or members: $($_.Exception.Message)" -ForegroundColor Red
    return
}

####################################################################################
#################### BUILD REPORT ##################################################
####################################################################################

$groupType = if (@($group.GroupTypes) -contains 'Unified') { 'Microsoft 365' }
             elseif ($group.SecurityEnabled -and $group.MailEnabled) { 'Mail-enabled security' }
             elseif ($group.SecurityEnabled) { 'Security' }
             elseif ($group.MailEnabled) { 'Distribution' }
             else { 'Other' }

$aliases = @($group.ProxyAddresses | Where-Object { $_ -match '^(?i)smtp:' } | ForEach-Object { $_ -replace '^(?i)smtp:', '' }) -join '; '
$results = @(
    foreach ($member in $members) {
        $memberType = 'Other'
        if ($member.AdditionalProperties -and $member.AdditionalProperties.ContainsKey('@odata.type')) {
            $memberType = [string]$member.AdditionalProperties['@odata.type'] -replace '^#microsoft.graph\.', ''
        }

        $name = $null
        $upn = $null
        $email = $null
        $memberId = $member.Id

        if ($memberType -eq 'user') {
            try {
                $user = Get-MgUser -UserId $memberId -Property 'displayName,userPrincipalName,mail' -ErrorAction Stop
                $name = $user.DisplayName
                $upn = $user.UserPrincipalName
                $email = $user.Mail
            }
            catch {
                Write-Warning "Could not read user $memberId : $($_.Exception.Message)"
            }
        }
        else {
            if ($member.AdditionalProperties) {
                $name = [string]$member.AdditionalProperties['displayName']
                $upn = [string]$member.AdditionalProperties['userPrincipalName']
                $email = [string]$member.AdditionalProperties['mail']
            }
        }

        [PSCustomObject]@{
            GroupName       = $group.DisplayName
            GroupEmail      = $group.Mail
            GroupType       = $groupType
            MailEnabled     = $group.MailEnabled
            SecurityEnabled = $group.SecurityEnabled
            GroupAliases    = $aliases
            MemberName      = $name
            MemberUPN       = $upn
            MemberEmail     = $email
            MemberType      = $memberType
            MemberId        = $memberId
        }
    }
)

# Preserve group metadata in the CSV even when there are no direct members.
if ($results.Count -eq 0) {
    $results = @([PSCustomObject]@{
        GroupName = $group.DisplayName; GroupEmail = $group.Mail; GroupType = $groupType
        MailEnabled = $group.MailEnabled; SecurityEnabled = $group.SecurityEnabled
        GroupAliases = $aliases; MemberName = $null; MemberUPN = $null
        MemberEmail = $null; MemberType = $null; MemberId = $null
    })
}

####################################################################################
#################### CONSOLE AND CSV OUTPUT ########################################
####################################################################################

Write-Host "`nGroup: $($group.DisplayName)" -ForegroundColor Cyan
Write-Host "Email: $($group.Mail)"
Write-Host "Type: $groupType"
Write-Host "Mail enabled: $($group.MailEnabled)"
Write-Host "Security enabled: $($group.SecurityEnabled)"
Write-Host "Aliases: $aliases"
Write-Host "Direct members: $($members.Count)"

$results | Format-Table MemberName, MemberUPN, MemberEmail, MemberType -AutoSize

$timestamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$filename = "Get_EntraGroup_Members_$timestamp.csv"
$csvPath = Join-Path -Path $PSScriptRoot -ChildPath $filename

try {
    $results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
    Write-Host "`nCSV saved: $csvPath" -ForegroundColor Green
}
catch {
    Write-Host "CSV export failed: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    $stopwatch.Stop()
    Write-Host "Execution time: $($stopwatch.Elapsed.ToString('hh\:mm\:ss'))" -ForegroundColor Cyan
}
