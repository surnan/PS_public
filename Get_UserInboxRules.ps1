
# Get-UserMailboxRules.ps1
# Input: Mailbox email address
#
# Output:
#   Inbox rules with compact console formatting
#   Long values truncated for readability
#
# Export:
#   CSV with all original rule properties and full values
#
# Read-only: Does not modify mailbox rules.


####################################################################################
#################### EXCHANGE ONLINE CONNECTION ####################################
####################################################################################
$module = "ExchangeOnlineManagement"

try {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Install-Module $module -Scope CurrentUser -ErrorAction Stop
    }

    if (-not (Get-Module -Name $module)) {
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
    Write-Host "Invalid email format (2)." -ForegroundColor Red
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

if ($rules.Count -eq 0) {
    Write-Host "`nNo Inbox Rules were found for $mailbox." -ForegroundColor Yellow
    return
}


####################################################################################
#################### PREPARE RESULTS ################################################
####################################################################################

# Convert values and collections into readable strings
function Convert-RuleValue {
    param($Value)

    if ($null -eq $Value) {
        return ""
    }

    return (@(
        $Value | ForEach-Object { [string]$_ }
    ) -join "; ")
}

$results = @(
    foreach ($rule in $rules) {
        [PSCustomObject]@{
            Name                  = $rule.Name
            Enabled               = $rule.Enabled
            Priority              = $rule.Priority
            Description           = Convert-RuleValue $rule.Description
            From                  = Convert-RuleValue $rule.From
            SentTo                = Convert-RuleValue $rule.SentTo
            SubjectContainsWords  = Convert-RuleValue $rule.SubjectContainsWords
            BodyContainsWords     = Convert-RuleValue $rule.BodyContainsWords
            MoveToFolder          = Convert-RuleValue $rule.MoveToFolder
            ForwardTo             = Convert-RuleValue $rule.ForwardTo
            RedirectTo            = Convert-RuleValue $rule.RedirectTo
            ForwardAsAttachmentTo = Convert-RuleValue $rule.ForwardAsAttachmentTo
            DeleteMessage         = $rule.DeleteMessage
            MarkAsRead            = $rule.MarkAsRead
            StopProcessingRules   = $rule.StopProcessingRules
        }
    }
)



####################################################################################
#################### CONSOLE OUTPUT ################################################
####################################################################################
Write-Host "`n$($results.Count) Inbox Rule(s) found." -ForegroundColor Green
Write-Host "`nMailbox: $mailbox" -ForegroundColor Cyan
Write-Host ""

$results |
Sort-Object Priority |
Format-Table `
    @{Label = "Name"; Expression = { $_.Name }; Width = 30 },
    @{Label = "Enabled"; Expression = { $_.Enabled }; Width = 7 },
    @{Label = "Priority"; Expression = { $_.Priority }; Width = 8 },
    @{Label = "Forward To"; Expression = { $_.ForwardTo }; Width = 30 },
    @{Label = "Redirect To"; Expression = { $_.RedirectTo }; Width = 30 },
    @{Label = "Description"; Expression = { $_.Description }; Width = 50 } `
    -Wrap:$false |
Out-Host



####################################################################################
#################### CREATE CSV ####################################################
####################################################################################
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

$safeMailbox = $mailbox -replace '[^\w\.-]', '_'

$csvPath = ".\MailboxRules_${safeMailbox}_$timestamp.csv"

try {
    $results |
    Sort-Object Priority |
    Export-Csv `
        -Path $csvPath `
        -NoTypeInformation `
        -Encoding utf8 `
        -ErrorAction Stop
}
catch {
    Write-Host "Unable to create CSV file (4):" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    return
}


####################################################################################
#################### EXECUTION SUMMARY #############################################
####################################################################################
Write-Host ("`n{0,-25}" -f "File Name:") -ForegroundColor DarkCyan -NoNewline
Write-Host "Get-UserMailboxRules.ps1" -ForegroundColor Cyan

Write-Host ("{0,-25}" -f "CSV file generated:") -ForegroundColor DarkYellow -NoNewline
Write-Host (Resolve-Path $csvPath) -ForegroundColor DarkYellow

Write-Host ("{0,-25}" -f "Total Rules:") -ForegroundColor DarkGreen -NoNewline
Write-Host $results.Count -ForegroundColor Green
