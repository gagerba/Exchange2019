#Requires -Version 5.1
<#
.SYNOPSIS
    Checks Exchange 2019 message tracking logs for both sender and recipient activity.
    Parallel edition — uses a RunspacePool to process multiple addresses concurrently.

.DESCRIPTION
    Reads a list of email addresses from a text file (one per line) and queries
    Exchange message tracking logs across ALL transport/mailbox servers to determine
    whether each address has sent or received mail in the specified period.
    Results include separate sent/received flags and timestamps for each address.

.PARAMETER AddressFile
    Path to the text file containing one email address per line.

.PARAMETER DaysBack
    Number of days to look back in message tracking logs. Default is 30.

.PARAMETER ThrottleLimit
    Maximum number of addresses processed in parallel. Default is 10.
    Raise carefully — each runspace opens Exchange connections to every server.

.PARAMETER ExportCsv
    Optional path to export results as a CSV file.

.EXAMPLE
    .\Check-MailActivity.ps1 -AddressFile "C:\addresses.txt"
    .\Check-MailActivity.ps1 -AddressFile "C:\addresses.txt" -DaysBack 60 -ThrottleLimit 15 -ExportCsv "C:\results.csv"
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$AddressFile,

    [int]$DaysBack = 30,

    [ValidateRange(1, 50)]
    [int]$ThrottleLimit = 10,

    [string]$ExportCsv
)

# -- Initialise ----------------------------------------------------------------

$StartDate = (Get-Date).AddDays(-$DaysBack)
$Addresses = Get-Content $AddressFile |
                 Where-Object { $_ -match '\S' } |
                 ForEach-Object { $_.Trim() }

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

Write-Host "  Found $($TransportServers.Count) server(s): $($TransportServers -join ', ')" -ForegroundColor Cyan
Write-Host "  Checking $($Addresses.Count) address(es) - activity since $($StartDate.ToString('yyyy-MM-dd'))" -ForegroundColor Cyan
Write-Host "  Parallel throttle : $ThrottleLimit concurrent address(es)`n" -ForegroundColor Cyan

# -- Scriptblock executed inside each runspace ---------------------------------
#    Each runspace is responsible for one address (both SEND + DELIVER checks).
#    The Exchange snap-in is loaded once per runspace.

$RunspaceScript = {
    param (
        [string]   $Address,
        [string[]] $TransportServers,
        [datetime] $StartDate
    )

    # Load Exchange cmdlets in this runspace.
    # Works when the script runs on an Exchange server (or where the snap-in is registered).
    # Falls back silently if already loaded (e.g. inherited session).
    try {
        if (-not (Get-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue)) {
            Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction Stop
        }
    }
    catch {
        # Snap-in path varies; try the E15/E19 module path as a fallback
        $exSetup = [System.Environment]::GetEnvironmentVariable('ExchangeInstallPath')
        if ($exSetup) {
            $modulePath = Join-Path $exSetup 'bin\RemoteExchange.ps1'
            if (Test-Path $modulePath) {
                . $modulePath
                Connect-ExchangeServer -auto -ClientApplication:ManagementShell 2>$null
            }
        }
    }

    # ---- inner helper --------------------------------------------------------
    function Invoke-TrackingQuery {
        param (
            [string]   $Address,
            [string]   $EventId,
            [string]   $Role,
            [string[]] $Servers,
            [datetime] $Since
        )

        $hits   = [System.Collections.Generic.List[object]]::new()
        $failed = [System.Collections.Generic.List[string]]::new()

        foreach ($Server in $Servers) {
            try {
                $params = @{
                    Server      = $Server
                    EventId     = $EventId
                    Start       = $Since
                    ResultSize  = 'Unlimited'
                    ErrorAction = 'Stop'
                }
                if ($Role -eq 'Sender')    { $params['Sender']     = $Address }
                else                       { $params['Recipients'] = $Address }

                $found = Get-MessageTrackingLog @params
                if ($found) { foreach ($h in $found) { $hits.Add($h) } }
            }
            catch {
                $failed.Add("$Server ($($_.Exception.Message))")
            }
        }

        return [PSCustomObject]@{ Hits = $hits; Failed = $failed }
    }
    # --------------------------------------------------------------------------

    $sentResult = Invoke-TrackingQuery -Address $Address -EventId 'SEND'    -Role 'Sender'    -Servers $TransportServers -Since $StartDate
    $recvResult = Invoke-TrackingQuery -Address $Address -EventId 'DELIVER' -Role 'Recipient' -Servers $TransportServers -Since $StartDate

    $hasSent     = $sentResult.Hits.Count -gt 0
    $hasReceived = $recvResult.Hits.Count -gt 0

    $lastSent     = if ($hasSent)     { ($sentResult.Hits | Sort-Object Timestamp -Descending | Select-Object -First 1).Timestamp } else { $null }
    $lastReceived = if ($hasReceived) { ($recvResult.Hits | Sort-Object Timestamp -Descending | Select-Object -First 1).Timestamp } else { $null }

    $allFailed = @(
        $sentResult.Failed + $recvResult.Failed | Sort-Object -Unique
    )

    return [PSCustomObject]@{
        Address        = $Address
        Sent           = $hasSent
        LastSent       = $lastSent
        Received       = $hasReceived
        LastReceived   = $lastReceived
        EitherActive   = ($hasSent -or $hasReceived)
        ServersQueried = $TransportServers.Count
        ServersFailed  = $allFailed.Count
        Notes          = if ($allFailed.Count -gt 0) { "Unreachable: " + ($allFailed -join '; ') } else { $null }
    }
}

# -- Build the RunspacePool and dispatch one job per address -------------------

$Pool = [RunspaceFactory]::CreateRunspacePool(1, $ThrottleLimit)
$Pool.Open()

$Jobs = [System.Collections.Generic.List[hashtable]]::new()

foreach ($Address in $Addresses) {
    $ps = [PowerShell]::Create()
    $ps.RunspacePool = $Pool
    [void]$ps.AddScript($RunspaceScript)
    [void]$ps.AddParameters([ordered]@{
        Address          = $Address
        TransportServers = $TransportServers
        StartDate        = $StartDate
    })
    $Jobs.Add(@{
        PS      = $ps
        Handle  = $ps.BeginInvoke()
        Address = $Address
    })
}

$total     = $Jobs.Count
$completed = 0
$pending   = [System.Collections.Generic.List[hashtable]]::new($Jobs)
$Results   = [System.Collections.Generic.List[PSCustomObject]]::new()

Write-Host "  All $total job(s) dispatched.`n" -ForegroundColor Cyan

# Helper: render one address result row to the console
function Write-AddressResult {
    param ([PSCustomObject]$Row)
    $sentLabel = if ($Row.Sent)     { 'ACTIVE'  } else { 'NO MAIL' }
    $recvLabel = if ($Row.Received) { 'ACTIVE'  } else { 'NO MAIL' }
    $sentColor = if ($Row.Sent)     { 'Green'   } elseif ($Row.ServersFailed -gt 0) { 'Red' } else { 'Yellow' }
    $recvColor = if ($Row.Received) { 'Green'   } elseif ($Row.ServersFailed -gt 0) { 'Red' } else { 'Yellow' }
    $warn      = if ($Row.ServersFailed -gt 0)  { ' [!]' } else { '' }
    Write-Host "  $($Row.Address)"
    Write-Host ("    Sent     : {0}{1}" -f $sentLabel, $warn) -ForegroundColor $sentColor
    Write-Host ("    Received : {0}{1}" -f $recvLabel, $warn) -ForegroundColor $recvColor
    if ($Row.Notes) { Write-Host "       Warn  : $($Row.Notes)" -ForegroundColor DarkYellow }
}

# -- Collect results: poll all pending jobs, harvest whichever finishes first --
#    This means the progress bar advances the moment ANY job completes,
#    rather than stalling on a slow address while faster ones sit idle.

while ($pending.Count -gt 0) {

    # How many runspaces are currently active (started but not yet done)
    $running = ($pending | Where-Object { $_.Handle.IsCompleted -eq $false }).Count

    # Build a live status line that shows in-flight addresses
    $inFlight = ($pending |
        Where-Object { -not $_.Handle.IsCompleted } |
        Select-Object -ExpandProperty Address |
        Select-Object -First 5) -join ', '
    if ($pending.Count - ($pending | Where-Object { $_.Handle.IsCompleted }).Count -gt 5) {
        $inFlight += ', ...'
    }

    $pct = [int](($completed / $total) * 100)
    Write-Progress `
        -Activity         "Checking mail activity  ($completed / $total complete)" `
        -Status           "Running: $running job(s) - $inFlight" `
        -PercentComplete  $pct `
        -SecondsRemaining -1     # suppress unreliable ETA

    # Find jobs that have finished since last pass
    $done = $pending | Where-Object { $_.Handle.IsCompleted }

    if ($done) {
        foreach ($job in @($done)) {

            $raw = $job.PS.EndInvoke($job.Handle)

            if ($raw) {
                $row = $raw
                if ($row -is [System.Collections.IEnumerable] -and $row -isnot [string]) {
                    $row = @($row)[0]
                }
                Write-AddressResult -Row $row
                $Results.Add($row)
            }
            else {
                Write-Host "  $($job.Address) — no result returned (runspace error)" -ForegroundColor Red
                $Results.Add([PSCustomObject]@{
                    Address        = $job.Address
                    Sent           = $false
                    LastSent       = $null
                    Received       = $false
                    LastReceived   = $null
                    EitherActive   = $false
                    ServersQueried = $TransportServers.Count
                    ServersFailed  = $TransportServers.Count
                    Notes          = 'Runspace returned no data'
                })
            }

            if ($job.PS.HadErrors) {
                foreach ($err in $job.PS.Streams.Error) {
                    Write-Warning "  Runspace error for $($job.Address): $err"
                }
            }

            $job.PS.Dispose()
            [void]$pending.Remove($job)
            $completed++
        }
    }
    else {
        # Nothing finished yet — yield the CPU briefly before polling again
        Start-Sleep -Milliseconds 200
    }
}

Write-Progress -Activity "Checking mail activity  ($total / $total complete)" -Completed

$Pool.Close()
$Pool.Dispose()

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

$Results | Sort-Object Address | Format-Table -AutoSize

# -- Optional CSV export -------------------------------------------------------

if ($ExportCsv) {
    $Results | Sort-Object Address | Export-Csv -Path $ExportCsv -NoTypeInformation -Encoding UTF8
    Write-Host "Results exported to: $ExportCsv`n" -ForegroundColor Cyan
}
