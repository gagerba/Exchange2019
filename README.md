# Exchange 2019 - Mailbox Activity Checker

A set of PowerShell scripts for auditing mailbox activity in Exchange 2019 using message tracking logs. Useful for identifying inactive accounts, stale mailboxes, or shared mailboxes that are no longer in use.

---

## Scripts

| Script | Purpose |
|---|---|
| `Check-RecipientActivity.ps1` | Checks whether addresses have **received** mail in the last N days |
| `Check-SenderActivity.ps1` | Checks whether addresses have **sent** mail in the last N days |
| `Check-MailActivity.ps1` | Checks both **sent and received** activity in a single pass |

All three scripts query **all Hub Transport / Mailbox servers** in the organisation automatically, ensuring no DAG node is missed.

---

## Requirements

- Exchange 2019 (on-premises)
- Exchange Management Shell, or a remote PowerShell session with Exchange snap-ins loaded
- At least View-Only Organization Management rights (required to query remote servers)
- PowerShell 5.1 or later

---

## Input File Format

All scripts accept a plain text file with one email address per line. Blank lines are ignored.

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

### Check-MailActivity.ps1

```powershell
# Basic - check last 30 days
.\Check-MailActivity.ps1 -AddressFile "C:\addresses.txt"

# Custom window + export results
.\Check-MailActivity.ps1 -AddressFile "C:\addresses.txt" -DaysBack 60 -ExportCsv "C:\results.csv"
```

### Parameters

| Parameter | Script | Required | Default | Description |
|---|---|---|---|---|
| `-RecipientFile` | Check-RecipientActivity | Yes | - | Path to the input text file |
| `-SenderFile` | Check-SenderActivity | Yes | - | Path to the input text file |
| `-AddressFile` | Check-MailActivity | Yes | - | Path to the input text file |
| `-DaysBack` | All | No | `30` | Number of days to look back |
| `-ExportCsv` | All | No | - | Path to export results as a CSV file |

---

## Output

### Check-RecipientActivity.ps1 / Check-SenderActivity.ps1

Each script prints a colour-coded result per address and a summary at the end.

```
Checking 4 sender(s) - activity since 2025-02-11

  Checking: john.doe@contoso.com          >>  ACTIVE
  Checking: jane.smith@contoso.com        >>  NO MAIL
  Checking: shared-mailbox@contoso.com    >>  ACTIVE
  Checking: old-account@contoso.com       >>  NO MAIL

--- Summary -------------------------------------------
  Servers queried : 2 (EXCH01, EXCH02)
  Total checked   : 4
  Active          : 2
  No mail         : 2
```

### Check-MailActivity.ps1

Prints two lines per address (sent / received) and a more detailed summary.

```
Checking 3 address(es) - activity since 2025-02-11

  Checking: john.doe@contoso.com
    Sent     : ACTIVE
    Received : ACTIVE
  Checking: jane.smith@contoso.com
    Sent     : NO MAIL
    Received : ACTIVE
  Checking: old-account@contoso.com
    Sent     : NO MAIL
    Received : NO MAIL

--- Summary -------------------------------------------
  Servers queried   : 2 (EXCH01, EXCH02)
  Total checked     : 3
  Sent and received : 1
  Sent only         : 0
  Received only     : 1
  No activity       : 1
```

If a server is unreachable during a query, a `[!]` warning is shown inline and the affected server is listed in the `Notes` column of the CSV export.

### CSV Columns

| Column | Scripts | Description |
|---|---|---|
| `Recipient` / `Sender` / `Address` | Per script | The email address checked |
| `Active` | Check-RecipientActivity, Check-SenderActivity | `True` if activity was found |
| `Sent` | Check-MailActivity | `True` if the address has sent mail |
| `Received` | Check-MailActivity | `True` if the address has received mail |
| `EitherActive` | Check-MailActivity | `True` if sent or received activity was found |
| `LastSent` | Check-SenderActivity, Check-MailActivity | Timestamp of the most recent sent message |
| `LastReceived` | Check-RecipientActivity, Check-MailActivity | Timestamp of the most recent received message |
| `ServersQueried` | All | Number of servers included in the query |
| `ServersFailed` | All | Number of servers that could not be reached |
| `Notes` | All | Error or warning details if any server was unreachable |

---

## How It Works

All scripts discover transport servers dynamically at startup using `Get-ExchangeServer`, then query each server individually so that every DAG node is covered. Unreachable servers are caught per address and recorded as warnings rather than terminating the run.

| Script | Event type queried |
|---|---|
| `Check-RecipientActivity.ps1` | `DELIVER` -- message delivered to a local mailbox |
| `Check-SenderActivity.ps1` | `SEND` -- message handed off by the local transport service |
| `Check-MailActivity.ps1` | Both `DELIVER` and `SEND` in a single pass |

---

## License

MIT
