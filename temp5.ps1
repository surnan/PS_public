
# Check whether Exchange Online is already connected
$connection = Get-ConnectionInformation -ErrorAction SilentlyContinue |
    Where-Object { $_.State -eq "Connected" }

if (-not $connection) {
    Write-Host "Not connected to Exchange Online. Connecting..." -ForegroundColor Yellow

    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -ErrorAction Stop
    }

    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
}

# Prompt user for mailbox email address
$userMailbox = Read-Host "Enter the user's email address"

Write-Host "`nSearching mailbox permissions for: $userMailbox" -ForegroundColor Cyan

# Retrieve all mailboxes
$allMailboxes = @(
    Get-EXOMailbox -ResultSize Unlimited -ErrorAction Stop |
        Where-Object {
            $_.PrimarySmtpAddress -ne $userMailbox
        }
)

Write-Host "`nMailbox count = $($allMailboxes.Count)" -ForegroundColor Magenta

# Start timer
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# Check mailbox permissions using 5 parallel workers
$results = @(
    $allMailboxes | ForEach-Object -Parallel {

        $mailbox = $_
        $userEmail = $using:userMailbox

        try {
            $permissions = Get-EXOMailboxPermission `
                -Identity $mailbox.PrimarySmtpAddress `
                -User $userEmail `
                -ErrorAction Stop

            $userPermissions = $permissions | Where-Object {
                $_.User -eq $userEmail -and
                $_.AccessRights -contains "FullAccess" -and
                -not $_.Deny
            }

            if ($userPermissions) {
                [PSCustomObject]@{
                    Mailbox = [string]$mailbox.PrimarySmtpAddress
                    Status  = "FullAccess"
                    Error   = ""
                }
            }
        }
        catch {
            # Ignore expected errors when no permissions exist
            if ($_.Exception.Message -match '"code":"NotFound"' -and
                $_.Exception.Message -match 'No permissions were found') {
                return
            }

            # Return unexpected errors for review
            [PSCustomObject]@{
                Mailbox = [string]$mailbox.PrimarySmtpAddress
                Status  = "Error"
                Error   = $_.Exception.Message
            }
        }

    } -ThrottleLimit 5
)

# Display successful matches
$fullAccess = @(
    $results | Where-Object { $_.Status -eq "FullAccess" }
)

foreach ($result in $fullAccess) {
    Write-Host "Full Access: $($result.Mailbox)" -ForegroundColor Green
}

# Display unexpected errors
$errorsFound = @(
    $results | Where-Object { $_.Status -eq "Error" }
)

foreach ($result in $errorsFound) {
    Write-Warning "Could not check $($result.Mailbox): $($result.Error)"
}

# Stop timer
$stopwatch.Stop()

Write-Host "`nTotal Full Access matches: $($fullAccess.Count)" -ForegroundColor Magenta
Write-Host "Unexpected errors: $($errorsFound.Count)" -ForegroundColor Yellow
Write-Host "Total execution time: $($stopwatch.Elapsed)" -ForegroundColor Cyan
