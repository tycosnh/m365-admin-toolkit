# Get-WeeklyEmailActivity

Rebuilds how much email a user sent and received, week by week, over any range in the last 90 days.

Managers and HR ask for this during reviews and investigations. The built-in usage reports often don't go back far enough or don't break down by week. Message trace does, so this script turns it into a clean table.

## What it does

- Splits your date range into weekly buckets, with a partial week at the end if needed.
- Counts **sent** mail, deduplicated by message. Trace returns one row per recipient, so raw counts would inflate.
- Counts **received** mail delivered to the user, Junk included.
- Warns if a week hits the 5,000-row query cap.
- Exports a CSV.

## Usage

```powershell
Connect-ExchangeOnline
.\Get-WeeklyEmailActivity.ps1 -User jane.doe@contoso.com -StartDate 2026-05-15 -EndDate 2026-06-15
```

Example output:

```
Week        Days Sent Received
----        ---- ---- --------
5/15 - 5/21    7   42      187
5/22 - 5/28    7   38      201
5/29 - 6/4     7   51      176
6/5 - 6/11     7   44      190
6/12 - 6/15    4   19       88
```

## Requirements

- `ExchangeOnlineManagement` 3.7.0 or later
- An account that can run message trace

## Notes

- Counts come from transport events, so they'll differ slightly from the usage report.
- For data older than 90 days, use `Start-HistoricalSearch` instead.
