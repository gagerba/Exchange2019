#Requires -Version 5.1
<#
.SYNOPSIS
    Checks Exchange 2019 message tracking logs for both sender and recipient activity.

.DESCRIPTION
    Reads a list of email addresses from a text file (one per line) and queries
    Exchange message tracking logs across ALL transport/mailbox servers to determine
    whether each address has sent or received mail in the specified period.
    Results include separate sent/received flags and timestamps for each address.

.PARAMETER AddressFile
    Path to the text file containing one email address per line.

.PARAMETER DaysBack
    Number of days to look back in message tracking logs. Default is 30.

.PARAMETER ExportCsv
    Optional path to export results as a CSV file.

.EXAMPLE
    .\Check-MailActivity.ps1 -AddressFile "C:\addresses.txt"
    .\Check-MailActivity.ps1 -AddressFile "C:\addresses.txt" -DaysBack 60 -ExportCsv "C:\results.csv"
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$AddressFile,

    [int]$DaysBack = 30,

    [string]$ExportCsv
)

# -- Initialise ----------------------------------------------------------------

$StartDate = (Get-Date).AddDays(-$DaysBack)
$Addresses = Get-Content $AddressFile | Where-Object { $_ -match '\S' } | ForEach-Object { $_.Trim() }
$Results   = [System.Collections.Generic.List[PSCustomObject]]::new()

# -- Resolve transport servers once, up front ----------------------------------

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
Write-Host "Checking $($Addresses.Count) address(es) - activity since $($StartDate.ToString('yyyy-MM-dd'))`n" -ForegroundColor Cyan

# -- Helper: query one event type across all servers ---------------------------

function Get-TrackingHits {
    param (
        [string]   $Address,
        [string]   $EventId,
        [string]   $Role,       # 'Sender' or 'Recipient'
        [string[]] $Servers,
        [datetime] $Since
    )

    $hits          = [System.Collections.Generic.List[object]]::new()
    $failedServers = [System.Collections.Generic.List[string]]::new()

    foreach ($Server in $Servers) {
        try {
            $params = @{
                Server      = $Server
                EventId     = $EventId
                Start       = $Since
                ResultSize  = 'Unlimited'
                ErrorAction = 'Stop'
            }

            if ($Role -eq 'Sender') {
                $params['Sender'] = $Address
            } else {
                $params['Recipients'] = $Address
            }

            $serverHits = Get-MessageTrackingLog @params
            if ($serverHits) {
                foreach ($h in $serverHits) { $hits.Add($h) }
            }
        }
        catch {
            $failedServers.Add("$Server ($($_.Exception.Message))")
        }
    }

    return [PSCustomObject]@{
        Hits          = $hits
        FailedServers = $failedServers
    }
}

# -- Main loop -----------------------------------------------------------------

foreach ($Address in $Addresses) {

    Write-Host "  Checking: $Address"

    # Sent (SEND event, address as sender)
    Write-Host "    Sent     : " -NoNewline
    $SentResult  = Get-TrackingHits -Address $Address -EventId 'SEND'    -Role 'Sender'    -Servers $TransportServers -Since $StartDate
    $HasSent     = $SentResult.Hits.Count -gt 0
    $LastSent    = if ($HasSent) { ($SentResult.Hits | Sort-Object Timestamp -Descending | Select-Object -First 1).Timestamp } else { $null }
    $SentColor   = if ($HasSent) { 'Green' } elseif ($SentResult.FailedServers.Count -gt 0) { 'Red' } else { 'Yellow' }
    Write-Host ("{0}{1}" -f $(if ($HasSent) { "ACTIVE" } else { "NO MAIL" }), $(if ($SentResult.FailedServers.Count -gt 0) { " [!]" } else { "" })) -ForegroundColor $SentColor

    # Received (DELIVER event, address as recipient)
    Write-Host "    Received : " -NoNewline
    $RecvResult  = Get-TrackingHits -Address $Address -EventId 'DELIVER' -Role 'Recipient' -Servers $TransportServers -Since $StartDate
    $HasReceived = $RecvResult.Hits.Count -gt 0
    $LastReceived= if ($HasReceived) { ($RecvResult.Hits | Sort-Object Timestamp -Descending | Select-Object -First 1).Timestamp } else { $null }
    $RecvColor   = if ($HasReceived) { 'Green' } elseif ($RecvResult.FailedServers.Count -gt 0) { 'Red' } else { 'Yellow' }
    Write-Host ("{0}{1}" -f $(if ($HasReceived) { "ACTIVE" } else { "NO MAIL" }), $(if ($RecvResult.FailedServers.Count -gt 0) { " [!]" } else { "" })) -ForegroundColor $RecvColor

    # Merge failed servers from both queries (deduplicated by server name)
    $AllFailed = @(
        $SentResult.FailedServers +
        $RecvResult.FailedServers |
            Sort-Object -Unique
    )

    $Notes = if ($AllFailed.Count -gt 0) {
        "Unreachable: " + ($AllFailed -join '; ')
    } else { $null }

    if ($AllFailed.Count -gt 0) {
        Write-Host "       Warn: could not query $($AllFailed.Count) server(s) for one or both checks" -ForegroundColor DarkYellow
    }

    $Results.Add([PSCustomObject]@{
        Address        = $Address
        Sent           = $HasSent
        LastSent       = $LastSent
        Received       = $HasReceived
        LastReceived   = $LastReceived
        EitherActive   = ($HasSent -or $HasReceived)
        ServersQueried = $TransportServers.Count
        ServersFailed  = $AllFailed.Count
        Notes          = $Notes
    })
}

# -- Summary -------------------------------------------------------------------

$BothActive    = ($Results | Where-Object { $_.Sent -and $_.Received }).Count
$SentOnly      = ($Results | Where-Object { $_.Sent -and -not $_.Received }).Count
$ReceivedOnly  = ($Results | Where-Object { -not $_.Sent -and $_.Received }).Count
$NeitherActive = ($Results | Where-Object { -not $_.Sent -and -not $_.Received -and $_.ServersFailed -eq 0 }).Count
$PartialCount  = ($Results | Where-Object { $_.ServersFailed -gt 0 }).Count

Write-Host "`n--- Summary -------------------------------------------" -ForegroundColor Cyan
Write-Host "  Servers queried   : $($TransportServers.Count) ($($TransportServers -join ', '))"
Write-Host "  Total checked     : $($Results.Count)"
Write-Host "  Sent and received : $BothActive"    -ForegroundColor Green
Write-Host "  Sent only         : $SentOnly"      -ForegroundColor Green
Write-Host "  Received only     : $ReceivedOnly"  -ForegroundColor Green
Write-Host "  No activity       : $NeitherActive" -ForegroundColor Yellow
if ($PartialCount -gt 0) {
    Write-Host "  Partial / error   : $PartialCount (results may be incomplete)" -ForegroundColor Red
}
Write-Host ""

$Results | Format-Table -AutoSize

# -- Optional CSV export -------------------------------------------------------

if ($ExportCsv) {
    $Results | Export-Csv -Path $ExportCsv -NoTypeInformation -Encoding UTF8
    Write-Host "Results exported to: $ExportCsv`n" -ForegroundColor Cyan
}
