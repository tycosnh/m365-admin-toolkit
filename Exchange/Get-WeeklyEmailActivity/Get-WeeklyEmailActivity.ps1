<#
.SYNOPSIS
    Reconstructs a user's sent and received email counts, week by week, from message trace.

.DESCRIPTION
    Microsoft 365 usage reports only keep a short window of per-user detail. When someone
    asks "how much email did this person send each week last month," message trace is the
    reliable source. This script splits any date range into weekly buckets, queries
    Get-MessageTraceV2 for each, and returns a clean table plus a CSV.

    - Sent mail is deduplicated by MessageId, since trace returns one row per recipient.
    - Received mail counts messages with Status = Delivered to the user (Junk included).
    - Counts come from transport events, so they will be close to, but not identical to,
      the "Email activity" usage report.

.PARAMETER User
    The mailbox to report on, e.g. jane.doe@contoso.com

.PARAMETER StartDate
    First day of the range (local time). Must be within the last 90 days.

.PARAMETER EndDate
    Last day of the range, inclusive. Defaults to today.

.PARAMETER OutputPath
    Folder for the CSV. Defaults to the current directory.

.EXAMPLE
    Connect-ExchangeOnline
    .\Get-WeeklyEmailActivity.ps1 -User jane.doe@contoso.com -StartDate 2026-05-15 -EndDate 2026-06-15

.NOTES
    Requires ExchangeOnlineManagement 3.7.0+ and an active Connect-ExchangeOnline session.
    Get-MessageTraceV2 keeps 90 days of data and allows at most 10 days per query,
    so weekly buckets stay well inside both limits.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$User,

    [Parameter(Mandatory)]
    [datetime]$StartDate,

    [datetime]$EndDate = (Get-Date).Date,

    [string]$OutputPath = (Get-Location).Path
)

$ErrorActionPreference = 'Stop'

if (-not (Get-Command Get-MessageTraceV2 -ErrorAction SilentlyContinue)) {
    throw "Get-MessageTraceV2 not found. Install ExchangeOnlineManagement 3.7.0+ and run Connect-ExchangeOnline first."
}

$StartDate = $StartDate.Date
$EndDate   = $EndDate.Date

if ($EndDate -lt $StartDate) {
    throw "EndDate ($($EndDate.ToShortDateString())) is before StartDate ($($StartDate.ToShortDateString()))."
}

$oldestAllowed = (Get-Date).Date.AddDays(-90)
if ($StartDate -lt $oldestAllowed) {
    throw ("StartDate is older than 90 days. Message trace only retains 90 days. " +
           "Use Start-HistoricalSearch for older data.")
}

# Build weekly buckets. The last one may be a partial week.
$weeks = [System.Collections.Generic.List[object]]::new()
$cursor = $StartDate
while ($cursor -le $EndDate) {
    $bucketEnd = $cursor.AddDays(6)
    if ($bucketEnd -gt $EndDate) { $bucketEnd = $EndDate }
    $weeks.Add([pscustomobject]@{
        Label = ("{0:M/d} - {1:M/d}" -f $cursor, $bucketEnd)
        Start = $cursor
        End   = $bucketEnd.AddDays(1).AddSeconds(-1)   # inclusive through end of day
        Days  = ($bucketEnd - $cursor).Days + 1
    })
    $cursor = $bucketEnd.AddDays(1)
}

$results = foreach ($week in $weeks) {
    Write-Host "Querying $($week.Label) ..." -ForegroundColor Cyan

    $sentRaw = @(Get-MessageTraceV2 -SenderAddress $User `
        -StartDate $week.Start -EndDate $week.End -ResultSize 5000)
    $sentCount = @($sentRaw | Sort-Object MessageId -Unique).Count

    $receivedRaw = @(Get-MessageTraceV2 -RecipientAddress $User -Status Delivered `
        -StartDate $week.Start -EndDate $week.End -ResultSize 5000)

    # 5000 is the hard cap per query. Flag it so the week can be split smaller.
    if ($sentRaw.Count -ge 5000 -or $receivedRaw.Count -ge 5000) {
        Write-Warning "$($week.Label): hit the 5000-row cap. Counts are truncated. Split this range into smaller windows."
    }

    [pscustomobject]@{
        Week     = $week.Label
        Days     = $week.Days
        Sent     = $sentCount
        Received = $receivedRaw.Count
    }
}

$results | Format-Table -AutoSize

$safeUser = ($User -split '@')[0] -replace '[^\w\-]', ''
$fileName = "{0}_WeeklyEmailActivity_{1:yyyyMMdd}-{2:yyyyMMdd}.csv" -f $safeUser, $StartDate, $EndDate
$csvPath  = Join-Path $OutputPath $fileName

$results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
Write-Host "Exported to $csvPath" -ForegroundColor Green
