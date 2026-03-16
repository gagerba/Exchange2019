#Requires -Version 5.1
<#
.SYNOPSIS
    Checks Exchange 2019 message tracking logs for recipient activity in the last 30 days.

.DESCRIPTION
    Reads a list of email addresses from a text file (one per line) and queries
    Exchange message tracking logs across ALL transport/mailbox servers to determine
    if each recipient has received any mail in the specified period.

.PARAMETER RecipientFile
    Path to the text file containing one email address per line.

.PARAMETER DaysBack
    Number of days to look back in message tracking logs. Default is 30.

.PARAMETER ExportCsv
    Optional path to export results as a CSV file.

.EXAMPLE
    .\Check-RecipientActivity.ps1 -RecipientFile "C:\recipients.txt"
    .\Check-RecipientActivity.ps1 -RecipientFile "C:\recipients.txt" -DaysBack 60 -ExportCsv "C:\results.csv"
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$RecipientFile,

    [int]$DaysBack = 30,

    [string]$ExportCsv
)

# -- Initialise --

$StartDate  = (Get-Date).AddDays(-$DaysBack)
$Recipients = Get-Content $RecipientFile | Where-Object { $_ -match '\S' } | ForEach-Object { $_.Trim() }
$Results    = [System.Collections.Generic.List[PSCustomObject]]::new()

# -- Resolve transport servers once, up front --

Write-Host "`nResolving Exchange transport servers..." -ForegroundColor Cyan

$TransportServers = @(
    Get-ExchangeServer -ErrorAction Stop |
        Where-Object { $_.IsHubTransportServer -or $_.IsMailboxServer } |
        Select-Object -ExpandProperty Name |
        Sort-Object
)

if ($TransportServers.Count -eq 0) {
    Write-Error "No Hub Transport / Mailbox servers found. Ensure you are running this inside the Exchange Management Shell."
    exit 1
}

Write-Host "  Found $($TransportServers.Count) server(s): $($TransportServers -join ', ')`n" -ForegroundColor Cyan
Write-Host "Checking $($Recipients.Count) recipient(s) - activity since $($StartDate.ToString('yyyy-MM-dd'))`n" -ForegroundColor Cyan

# -- Main loop --

foreach ($Recipient in $Recipients) {

    Write-Host "  Checking: $Recipient" -NoNewline

    $AllHits        = [System.Collections.Generic.List[object]]::new()
    $FailedServers  = [System.Collections.Generic.List[string]]::new()

    foreach ($Server in $TransportServers) {
        try {
            $ServerHits = Get-MessageTrackingLog `
                -Server      $Server `
                -EventId     DELIVER `
                -Recipients  $Recipient `
                -Start       $StartDate `
                -ResultSize  Unlimited `
                -ErrorAction Stop

            if ($ServerHits) {
                foreach ($hit in $ServerHits) { $AllHits.Add($hit) }
            }
        }
        catch {
            $FailedServers.Add("$Server ($($_.Exception.Message))")
        }
    }

    $Active       = $AllHits.Count -gt 0
    $LastReceived = if ($Active) {
        ($AllHits | Sort-Object Timestamp -Descending | Select-Object -First 1).Timestamp
    } else { $null }

    $Notes = if ($FailedServers.Count -gt 0) {
        "Unreachable: " + ($FailedServers -join '; ')
    } else { $null }

    $StatusLabel = if ($Active) { "ACTIVE" } else { "NO MAIL" }
    $StatusColor = if ($Active) { 'Green' } elseif ($FailedServers.Count -gt 0) { 'Red' } else { 'Yellow' }

    Write-Host ("  >>  {0}{1}" -f $StatusLabel, $(if ($FailedServers.Count -gt 0) { " [!]" } else { "" })) `
        -ForegroundColor $StatusColor

    if ($FailedServers.Count -gt 0) {
        Write-Host "       Warn: could not query $($FailedServers.Count) server(s)" -ForegroundColor DarkYellow
    }

    $Results.Add([PSCustomObject]@{
        Recipient      = $Recipient
        Active         = $Active
        LastReceived   = $LastReceived
        ServersQueried = $TransportServers.Count
        ServersFailed  = $FailedServers.Count
        Notes          = $Notes
    })
}

# -- Summary --

$ActiveCount      = ($Results | Where-Object Active).Count
$InactiveCount    = ($Results | Where-Object { -not $_.Active -and $_.ServersFailed -eq 0 }).Count
$PartialCount     = ($Results | Where-Object { $_.ServersFailed -gt 0 }).Count

Write-Host "`n-- Summary --" -ForegroundColor Cyan
Write-Host "  Servers queried : $($TransportServers.Count) ($($TransportServers -join ', '))"
Write-Host "  Total checked   : $($Results.Count)"
Write-Host "  Active          : $ActiveCount"   -ForegroundColor Green
Write-Host "  No mail         : $InactiveCount" -ForegroundColor Yellow
if ($PartialCount -gt 0) {
    Write-Host "  Partial / error : $PartialCount (results may be incomplete)" -ForegroundColor Red
}
Write-Host ""

$Results | Format-Table -AutoSize

# -- Optional CSV export --

if ($ExportCsv) {
    $Results | Export-Csv -Path $ExportCsv -NoTypeInformation -Encoding UTF8
    Write-Host "Results exported to: $ExportCsv`n" -ForegroundColor Cyan
}
