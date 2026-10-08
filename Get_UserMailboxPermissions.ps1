# Don't create connection if it already exists
$connection = Get-ConnectionInformation -ErrorAction SilentlyContinue |
Where-Object { $_.State -eq "Connected" }

if (-not $connection) {
    #Don't install module if it already exists
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -ErrorAction Stop
    }
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
}


# Prompt user for mailbox email address
$userMailbox = Read-Host "Enter the user's email address"
Write-Host "`nSearching mailbox permissions for: $userMailbox" -ForegroundColor Cyan


# Retrieve all mailboxes except the user's own mailbox
# Don't need to check mailbox in check mailbox permissionloop
$allMailboxes = @(
    Get-EXOMailbox -ResultSize Unlimited -ErrorAction Stop |
    Where-Object {
        $_.PrimarySmtpAddress -ne $userMailbox
    }
)


Write-Host "`nMailbox count = $($allMailboxes.Count)" -ForegroundColor Magenta


# Start timer
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()


# Display temporary progress message
Write-Progress `
    -Activity "Checking mailbox permissions" `
    -Status "Scanning $($allMailboxes.Count) mailboxes using 5 parallel workers..."


# Check mailbox permissions using parallel processing
try {

    # ForEach-Object better for parallel processing
    $results = @(
        $allMailboxes | ForEach-Object -Parallel {
            $currentMailbox = $_
            # Parallel processing has different run spaces.  
            # Not guaranteed to have access to main session variables
            # $using: forces value into each runspace as $userEmail 
            $userEmail = $using:userMailbox

            try {
                $permissions = Get-EXOMailboxPermission `
                    -Identity $currentMailbox.PrimarySmtpAddress `
                    -User $userEmail `
                    -ErrorAction Stop

                if ($permissions | Where-Object {
                        $_.User -eq $userEmail -and
                        $_.AccessRights -contains "FullAccess" -and
                        -not $_.Deny
                    }) {
                    [PSCustomObject]@{
                        Mailbox = [string]$currentMailbox.PrimarySmtpAddress
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

                # Return unexpected errors
                [PSCustomObject]@{
                    Mailbox = [string]$currentMailbox.PrimarySmtpAddress
                    Status  = "Error"
                    Error   = $_.Exception.Message
                }
            }
        } -ThrottleLimit 5
    )

}
finally {
    Write-Progress -Activity "Checking mailbox permissions" -Completed
}


# Get successful matches
$fullAccess = @(
    $results | Where-Object {
        $_.Status -eq "FullAccess"
    }
)


# Display Full Access mailboxes
foreach ($result in $fullAccess) {
    Write-Host "Full Access: $($result.Mailbox)" `
        -ForegroundColor Green
}


################################################################
###############
# Get unexpected errors
$errorsFound = @(
    $results | Where-Object {
        $_.Status -eq "Error"
    }
)
# Display unexpected errors
foreach ($result in $errorsFound) {
    Write-Warning "Could not check $($result.Mailbox): $($result.Error)"
}
###############
################################################################


# Stop timer
$stopwatch.Stop()


# Display summary
Write-Host "`nTotal Full Access matches: $($fullAccess.Count)" -ForegroundColor Magenta
Write-Host "Unexpected errors: $($errorsFound.Count)" -ForegroundColor Yellow
Write-Host "Total execution time: $($stopwatch.Elapsed)" -ForegroundColor Cyan
