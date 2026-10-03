function Test-ValidDate {param ([string]$DateString)
    # Out-Null to discard return value of System.DateTime
    try {
        [System.DateTime]::ParseExact(
            $DateString,
            "MM-dd-yyyy",
            $null
        ) | Out-Null
        return $true
    }
    catch {
        return $false
    }
}


$today = (Get-Date).ToString("MM-dd-yyyy")

$startDate = Read-Host "Enter start date (MM-DD-YYYY) or blank for today"
if ([string]::IsNullOrWhiteSpace($startDate)){
    $startDate = $today
}

$endDate = Read-Host "Enter end date (MM-DD-YYYY) or blank for today"
if ([string]::IsNullOrWhiteSpace($endDate)){
    $endDate = $today
}



if (-not (Test-ValidDate $startDate) -or  -not (Test-ValidDate $endDate)){
    Write-Host "Invalid date entered" -ForegroundColor Red
    exit
}



Write-Host("Start Date: ") -ForegroundColor Blue -NoNewline
Write-Host("$($startDate)") -ForegroundColor Green

Write-Host("End Date: ") -ForegroundColor Blue -NoNewline
Write-Host("$($endDate)") -ForegroundColor Green