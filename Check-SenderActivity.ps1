#Requires -Version 5.1
<#
.SYNOPSIS
    Checks Exchange 2019 message tracking logs for sender activity in the last 30 days.

.DESCRIPTION
    Reads a list of email addresses from a text file (one per line) and queries
    Exchange message tracking logs to determine if each sender has sent
    any mail in the last 30 days.

.PARAMETER SenderFile
    Path to the text file containing one email address per line.

.PARAMETER DaysBack
    Number of days to look back in message tracking logs. Default is 30.

.PARAMETER ExportCsv
    Optional path to export results as a CSV file.

.EXAMPLE
    .\Check-SenderActivity.ps1 -SenderFile "C:\senders.txt"
    .\Check-SenderActivity.ps1 -SenderFile "C:\senders.txt" -DaysBack 60 -ExportCsv "C:\results.csv"
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$SenderFile,

    [int]$DaysBack = 30,

    [string]$ExportCsv
)

# ── Initialise ────────────────────────────────────────────────────────────────

$StartDate = (Get-Date).AddDays(-$DaysBack)
$Senders   = Get-Content $SenderFile | Where-Object { $_ -match '\S' } | ForEach-Object { $_.Trim() }
$Results   = [System.Collections.Generic.List[PSCustomObject]]::new()

Write-Host "`nChecking $($Senders.Count) sender(s) — activity since $($StartDate.ToString('yyyy-MM-dd'))`n" -ForegroundColor Cyan

# ── Main loop ─────────────────────────────────────────────────────────────────

foreach ($Sender in $Senders) {

    Write-Host "  Checking: $Sender" -NoNewline

    try {
        $Hits = Get-MessageTrackingLog `
            -EventId     SEND `
            -Sender      $Sender `
            -Start       $StartDate `
            -ResultSize  1 `
            -ErrorAction Stop

        $Active   = $Hits.Count -gt 0
        $LastSent = if ($Active) { ($Hits | Sort-Object Timestamp -Descending | Select-Object -First 1).Timestamp } else { $null }

        Write-Host ("  →  {0}" -f $(if ($Active) { "ACTIVE" } else { "NO MAIL" })) `
            -ForegroundColor $(if ($Active) { 'Green' } else { 'Yellow' })

        $Results.Add([PSCustomObject]@{
            Sender    = $Sender
            Active    = $Active
            LastSent  = $LastSent
            Notes     = $null
        })
    }
    catch {
        Write-Host "  →  ERROR" -ForegroundColor Red
        $Results.Add([PSCustomObject]@{
            Sender    = $Sender
            Active    = $false
            LastSent  = $null
            Notes     = $_.Exception.Message
        })
    }
}

# ── Summary ───────────────────────────────────────────────────────────────────

$ActiveCount   = ($Results | Where-Object Active).Count
$InactiveCount = $Results.Count - $ActiveCount

Write-Host "`n─── Summary ────────────────────────────────" -ForegroundColor Cyan
Write-Host "  Total checked : $($Results.Count)"
Write-Host "  Active        : $ActiveCount"  -ForegroundColor Green
Write-Host "  No mail / err : $InactiveCount`n" -ForegroundColor Yellow

$Results | Format-Table -AutoSize

# ── Optional CSV export ───────────────────────────────────────────────────────

if ($ExportCsv) {
    $Results | Export-Csv -Path $ExportCsv -NoTypeInformation -Encoding UTF8
    Write-Host "Results exported to: $ExportCsv`n" -ForegroundColor Cyan
}
