# Get-UserMailboxRules.ps1
# Input: Mailbox email address
#
# Output:
#   All inbox rules
#   Rule name, status, priority, conditions, and actions
#
# Read-only: Does not modify mailbox rules.


####################################################################################
#################### EXCHANGE ONLINE CONNECTION ####################################
####################################################################################
$modules = @(
    "ExchangeOnlineManagement"
)

try {
    foreach ($module in $modules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            Install-Module $module -Scope CurrentUser -ErrorAction Stop
        }

        Import-Module $module -ErrorAction Stop
    }

    $connection = Get-ConnectionInformation -ErrorAction Stop |
        Where-Object { $_.State -eq "Connected" }

    if (-not $connection) {
        Connect-ExchangeOnline `
            -ShowBanner:$false `
            -ErrorAction Stop
    }
}
catch {
    Write-Host "Exchange Online connection failed (1):" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    return
}


####################################################################################
#################### GET INPUT PARAMETERS ##########################################
####################################################################################
$mailbox = (Read-Host "Enter user's email address").Trim()

if ($mailbox -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
    Write-Host "Invalid email format (2):" -ForegroundColor Red
    return
}


####################################################################################
#################### GET MAILBOX RULES #############################################
####################################################################################
Write-Host "`nSearching mailbox rules for: $mailbox" -ForegroundColor Cyan

try {
    $rules = @(
        Get-InboxRule `
            -Mailbox $mailbox `
            -ErrorAction Stop
    )
}
catch {
    Write-Host "Unable to retrieve Inbox Rules (3):" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    return
}


####################################################################################
#################### CONSOLE OUTPUT ################################################
####################################################################################
if ($rules.Count -eq 0) {
    Write-Host "`nNo Inbox Rules were found for $mailbox. (4)" -ForegroundColor Yellow
    return
}

Write-Host "`n$($rules.Count) Inbox Rule(s) found.`n" -ForegroundColor Green

$rules |
    Select-Object `
        Name,
        Enabled,
        Priority,
        Description,
        From,
        SentTo,
        SubjectContainsWords,
        BodyContainsWords,
        MoveToFolder,
        ForwardTo,
        RedirectTo,
        ForwardAsAttachmentTo,
        DeleteMessage,
        MarkAsRead,
        StopProcessingRules |
    Format-Table -AutoSize -Wrap |
    Out-Host
