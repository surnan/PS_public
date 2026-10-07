# Check whether Exchange Online is already connected
$connection = Get-ConnectionInformation -ErrorAction SilentlyContinue

if (-not $connection) {
    Write-Host "Not connected to Exchange Online. Connecting..." -ForegroundColor Yellow
    Connect-ExchangeOnline
}


# Prompt user for mailbox email address
$mailbox = Read-Host "Enter the user's email address"
Write-Host "`nSearching mailbox rules for: $mailbox" -ForegroundColor Cyan

try {
    # Get all Inbox Rules for the mailbox
    $rules = Get-InboxRule `
        -Mailbox $mailbox `
        -ErrorAction Stop
    if (-not $rules) {
        Write-Host "`nNo Inbox Rules were found for $mailbox." -ForegroundColor Yellow
        return
    }

    Write-Host "`n$($rules.Count) Inbox Rule(s) found.`n" -ForegroundColor Green
    # Display rule information
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
        Format-Table -AutoSize -Wrap
}
catch {
    Write-Host "`nUnable to retrieve Inbox Rules." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}