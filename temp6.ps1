# Check whether Exchange Online is already connected
$connection = Get-ConnectionInformation -ErrorAction SilentlyContinue

if (-not $connection) {
    Write-Host "Not connected to Exchange Online. Connecting..." -ForegroundColor Yellow
    Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -ErrorAction Stop
    Connect-ExchangeOnline
}


# Prompt user for mailbox email address
$userMailbox = Read-Host "Enter the user's email address"
Write-Host "`nSearching mailbox permissions for: $userMailbox" -ForegroundColor Cyan

$allMailboxes = @(Get-EXOMailbox -ResultSize Unlimited -ErrorAction Stop)

Write-Host("`n Mailbox count = $($allMailboxes.Count)") -ForegroundColor Magenta

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

foreach ($mailbox in $allMailboxes) {
    if ($mailbox.PrimarySmtpAddress -eq $userMailbox) { continue }


    try {
        $permissions = Get-EXOMailboxPermission `
            -Identity $mailbox.PrimarySmtpAddress `
            -User $userMailbox `
            -ErrorAction Stop

        $userPermissions = $permissions | Where-Object {
            $_.User -eq $userMailbox -and
            $_.AccessRights -contains "FullAccess" -and
            -not $_.Deny
        }

        if ($userPermissions.accessrights -contains "FullAccess" -and -not $userPermissions.deny) {
            Write-Host "Full Access: $($mailbox.PrimarySmtpAddress)" -ForegroundColor Green
        }
    }
    catch {

        if ($_.Exception.Message -match '"code":"NotFound"' -and
            $_.Exception.Message -match 'No permissions were found') {
            continue
        }

        Write-Warning "Could not check $($mailbox.PrimarySmtpAddress): $($_.Exception.Message)"
    }
}

$stopwatch.Stop()

Write-Host "Total execution time: $($stopwatch.Elapsed)" `
    -ForegroundColor Cyan