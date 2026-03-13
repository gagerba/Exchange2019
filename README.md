# Exchange 2019 - Mailbox Activity Checker

A pair of PowerShell scripts for auditing mailbox activity in Exchange 2019 using message tracking logs. Useful for identifying inactive accounts, stale mailboxes, or shared mailboxes that are no longer in use.

---

## Scripts

| Script | Purpose |
|---|---|
| `Check-RecipientActivity.ps1` | Checks whether addresses have **received** mail in the last N days |
| `Check-SenderActivity.ps1` | Checks whether addresses have **sent** mail in the last N days |

---

## Requirements

- Exchange 2019 (on-premises)
- Exchange Management Shell, or a remote PowerShell session with Exchange snap-ins loaded
- Read access to message tracking logs on the target server
- PowerShell 5.1 or later

---

## Input File Format

Both scripts accept a plain text file with one email address per line. Blank lines are ignored.

```
john.doe@contoso.com
jane.smith@contoso.com
shared-mailbox@contoso.com
```

---

## Usage

### Check-RecipientActivity.ps1

```powershell
# Basic - check last 30 days
.\Check-RecipientActivity.ps1 -RecipientFile "C:\recipients.txt"

# Custom window + export results
.\Check-RecipientActivity.ps1 -RecipientFile "C:\recipients.txt" -DaysBack 60 -ExportCsv "C:\results.csv"
```

### Check-SenderActivity.ps1

```powershell
# Basic - check last 30 days
.\Check-SenderActivity.ps1 -SenderFile "C:\senders.txt"

# Custom window + export results
.\Check-SenderActivity.ps1 -SenderFile "C:\senders.txt" -DaysBack 60 -ExportCsv "C:\results.csv"
```

### Parameters

| Parameter | Required | Default | Description |
|---|---|---|---|
| `-RecipientFile` / `-SenderFile` | Yes | - | Path to the input text file |
| `-DaysBack` | No | `30` | Number of days to look back |
| `-ExportCsv` | No | - | Path to export results as a CSV file |

---

## Output

Each script prints a colour-coded result per address and a summary at the end.

```
Checking 4 sender(s) - activity since 2025-02-11

  Checking: john.doe@contoso.com          >>  ACTIVE
  Checking: jane.smith@contoso.com        >>  NO MAIL
  Checking: shared-mailbox@contoso.com    >>  ACTIVE
  Checking: old-account@contoso.com       >>  NO MAIL

--- Summary -------------------------------------------
  Total checked : 4
  Active        : 2
  No mail / err : 2
```

If `-ExportCsv` is specified, the results are written to a UTF-8 CSV with the following columns:

| Column | Description |
|---|---|
| `Sender` / `Recipient` | The email address checked |
| `Active` | `True` if activity was found, `False` otherwise |
| `LastSent` / `LastReceived` | Timestamp of the most recent matching message |
| `Notes` | Error message if the lookup failed |

---

## How It Works

Both scripts call `Get-MessageTrackingLog` with `-ResultSize 1` — retrieving only a single matching entry per address. This keeps execution fast even in large environments with extensive log histories.

| Script | Event type queried |
|---|---|
| `Check-RecipientActivity.ps1` | `DELIVER` — message delivered to a local mailbox |
| `Check-SenderActivity.ps1` | `SEND` — message handed off by the local transport service |

Errors for individual addresses are caught and recorded without interrupting the rest of the run.

---

## License

MIT
