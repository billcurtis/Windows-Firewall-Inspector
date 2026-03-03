<#
.SYNOPSIS
    Firewall Log Collector - Enables firewall logging, Security Event Log
    auditing, and writes firewall log entries to date-stamped CSV files.

.DESCRIPTION
    Standalone script designed to run locally OR via Invoke-AzVMRunCommand.
      * Enables Windows Firewall logging (allowed + dropped connections)
        on all profiles (Domain, Private, Public).
      * Enables Windows Filtering Platform audit policies so firewall
        allow/drop events appear in the Security Event Log.
      * Periodically reads the firewall log (pfirewall.log) and appends
        entries to a date-stamped CSV file compatible with
        FirewallLogViewer.ps1.
      * Can be installed as a Scheduled Task that runs as SYSTEM at
        startup so it persists across reboots and runs whether or not
        a user is logged in.

    Compatible with Invoke-AzVMRunCommand -Parameter @{Action='...'} syntax
    (all parameters are string-typed key-value pairs).

.PARAMETER Action
    The operation to perform. One of:
      Enable    - (default) Enables logging/auditing and starts the
                  continuous collection loop (for scheduled task use).
      Disable   - Disables Security Event Log auditing and restores
                  default firewall logging settings.
      Install   - Copies this script to a well-known path on the VM
                  and registers a Scheduled Task that runs it as SYSTEM
                  at startup.
      Uninstall - Removes the scheduled task and disables logging.
      Collect   - One-shot mode: enables logging, performs a single
                  collection pass, writes the CSV, then exits.
                  Ideal for Invoke-AzVMRunCommand.

.PARAMETER OutputPath
    Directory where CSV files are written.
    Default: C:\ProgramData\FirewallLogCollector\Logs

.PARAMETER IntervalSeconds
    Collection & upload interval in seconds (used by Enable mode).
    Default: 60

.EXAMPLE
    # Local: enable logging and start the collection loop
    .\FirewallLogCollector.ps1 -Action Enable

.EXAMPLE
    # Local: install the scheduled task
    .\FirewallLogCollector.ps1 -Action Install

.EXAMPLE
    # Via Invoke-AzVMRunCommand: one-shot collection
    Invoke-AzVMRunCommand -ResourceGroupName 'rg' -VMName 'vm' `
        -CommandId 'RunPowerShellScript' `
        -ScriptPath '.\FirewallLogCollector.ps1' `
        -Parameter @{ Action = 'Collect' }

.EXAMPLE
    # Via Invoke-AzVMRunCommand: install the scheduled task on a remote VM
    Invoke-AzVMRunCommand -ResourceGroupName 'rg' -VMName 'vm' `
        -CommandId 'RunPowerShellScript' `
        -ScriptPath '.\FirewallLogCollector.ps1' `
        -Parameter @{ Action = 'Install'; IntervalSeconds = '30' }

.EXAMPLE
    # Via Invoke-AzVMRunCommand: disable everything
    Invoke-AzVMRunCommand -ResourceGroupName 'rg' -VMName 'vm' `
        -CommandId 'RunPowerShellScript' `
        -ScriptPath '.\FirewallLogCollector.ps1' `
        -Parameter @{ Action = 'Uninstall' }
#>
#Requires -Version 5.1

[CmdletBinding()]
param(
    [ValidateSet('Enable','Disable','Install','Uninstall','Collect')]
    [string]$Action = 'Enable',

    [string]$OutputPath = 'C:\ProgramData\FirewallLogCollector\Logs',

    [int]$IntervalSeconds = 60
)

# ------------------------------------------------------------------
# Constants
# ------------------------------------------------------------------
$TaskName           = "FirewallLogCollector"
$DeployDir          = "C:\ProgramData\FirewallLogCollector"
$DeployedScriptPath = Join-Path $DeployDir "FirewallLogCollector.ps1"
$FirewallLogPath    = "$env:SystemRoot\System32\LogFiles\Firewall\pfirewall.log"
$script:LastFilePos = 0
$script:LastEventTime = $null                     # tracks Security Event Log read position
$script:RuleNameCache = @{}                        # FilterId -> display name
$script:EventRuleMap  = @{}                        # "proto-srcip-srcport-dstip-dstport" -> RuleName
$script:EventAppMap   = @{}                        # "proto-srcip-srcport-dstip-dstport" -> Application path

# ------------------------------------------------------------------
# Require elevation
# ------------------------------------------------------------------
function Assert-Administrator {
    $principal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Error "This script must be run as Administrator (elevated)."
        exit 1
    }
}

# ------------------------------------------------------------------
# Firewall logging helpers
# ------------------------------------------------------------------
function Enable-FirewallAndAuditLogging {
    <#
    .SYNOPSIS Enables Windows Firewall log file logging and Security Event Log auditing.
    #>
    [CmdletBinding()]
    param()

    Write-Host "[INIT] Enabling Windows Firewall logging on all profiles..." -ForegroundColor Cyan

    # Enable pfirewall.log for allowed + dropped on all profiles
    $cmds = @(
        "netsh advfirewall set allprofiles logging filename `"$FirewallLogPath`"",
        "netsh advfirewall set allprofiles logging maxfilesize 32767",
        "netsh advfirewall set allprofiles logging droppedconnections enable",
        "netsh advfirewall set allprofiles logging allowedconnections enable"
    )
    foreach ($cmd in $cmds) {
        $result = Invoke-Expression $cmd 2>&1
        if ($LASTEXITCODE -ne 0) { Write-Warning "  Command failed: $cmd  $result" }
    }
    Write-Host "[INIT] Firewall file logging enabled (pfirewall.log)." -ForegroundColor Green

    # Enable Security Event Log auditing for Windows Filtering Platform
    Write-Host "[INIT] Enabling Security Event Log auditing for firewall events..." -ForegroundColor Cyan

    # Filtering Platform Packet Drop - Success and Failure
    & auditpol /set /subcategory:"Filtering Platform Packet Drop" /success:enable /failure:enable 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) { Write-Warning "  Failed to set audit policy for Filtering Platform Packet Drop." }

    # Filtering Platform Connection - Success and Failure
    & auditpol /set /subcategory:"Filtering Platform Connection" /success:enable /failure:enable 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) { Write-Warning "  Failed to set audit policy for Filtering Platform Connection." }

    Write-Host "[INIT] Security Event Log auditing enabled (Success + Failure)." -ForegroundColor Green
}

function Disable-FirewallAndAuditLogging {
    <#
    .SYNOPSIS Disables Security Event Log auditing and restores default firewall logging.
    #>
    [CmdletBinding()]
    param()

    Write-Host "[STOP] Disabling Security Event Log auditing for firewall events..." -ForegroundColor Yellow

    & auditpol /set /subcategory:"Filtering Platform Packet Drop" /success:disable /failure:disable 2>&1 | Out-Null
    & auditpol /set /subcategory:"Filtering Platform Connection" /success:disable /failure:disable 2>&1 | Out-Null

    Write-Host "[STOP] Security Event Log auditing disabled." -ForegroundColor Green

    Write-Host "[STOP] Restoring default firewall logging settings..." -ForegroundColor Yellow
    & netsh advfirewall set allprofiles logging droppedconnections disable 2>&1 | Out-Null
    & netsh advfirewall set allprofiles logging allowedconnections disable 2>&1 | Out-Null
    Write-Host "[STOP] Firewall logging restored to defaults." -ForegroundColor Green
}

# ------------------------------------------------------------------
# Security Event Log → firewall rule name resolution
# ------------------------------------------------------------------

function Update-FirewallRuleCache {
    <#
    .SYNOPSIS Builds/refreshes a hashtable of WFP Filter Run-Time IDs to
             friendly firewall rule display names.
    #>
    [CmdletBinding()]
    param()

    try {
        $rules = Get-NetFirewallRule -ErrorAction SilentlyContinue
        foreach ($r in $rules) {
            $script:RuleNameCache[$r.Name] = $r.DisplayName
        }
    } catch {
        Write-Warning "Could not refresh firewall rule cache: $_"
    }
}

function Read-SecurityFirewallEvents {
    <#
    .SYNOPSIS Reads recent Security Event Log events (5156 = allow, 5157 = block)
             and populates a lookup table keyed on connection 5-tuple so that
             pfirewall.log entries can be enriched with the matching rule name.
    #>
    [CmdletBinding()]
    param()

    # Only read events since our last check (or last 2 minutes on first run)
    $since = if ($script:LastEventTime) { $script:LastEventTime } else { (Get-Date).AddSeconds(-($IntervalSeconds + 30)) }
    $script:LastEventTime = Get-Date

    try {
        # Event IDs: 5156 (allowed), 5157 (blocked)
        $xpathFilter = "*[System[(EventID=5156 or EventID=5157) and TimeCreated[@SystemTime>='$($since.ToUniversalTime().ToString('o'))' ]]]"
        $events = Get-WinEvent -LogName Security -FilterXPath $xpathFilter -ErrorAction SilentlyContinue

        if (-not $events) { return }

        foreach ($evt in $events) {
            try {
                $xmlData  = [xml]$evt.ToXml()
                $ns       = New-Object Xml.XmlNamespaceManager $xmlData.NameTable
                $ns.AddNamespace('e', 'http://schemas.microsoft.com/win/2004/08/events/event')

                $dataNodes = $xmlData.SelectNodes('//e:EventData/e:Data', $ns)
                $props = @{}
                foreach ($d in $dataNodes) {
                    $props[$d.GetAttribute('Name')] = $d.InnerText
                }

                # Build a 5-tuple key matching how pfirewall.log entries look
                $proto = switch ($props['Protocol']) {
                    '6'  { 'TCP' }
                    '17' { 'UDP' }
                    '1'  { 'ICMP' }
                    default { $props['Protocol'] }
                }
                $key = "$proto-$($props['SourceAddress'])-$($props['SourcePort'])-$($props['DestAddress'])-$($props['DestPort'])"

                # Resolve the filter/rule name
                $filterName = $props['FilterRTID']
                $layerName  = $props['LayerRTID']
                $ruleName   = ''

                # Try direct lookup by FilterRTID in our rule name cache
                if ($filterName -and $script:RuleNameCache.ContainsKey($filterName)) {
                    $ruleName = $script:RuleNameCache[$filterName]
                }

                # Fallback: use LayerName or FilterRTID as-is
                if (-not $ruleName -and $filterName) {
                    # Try Get-NetFirewallRule by name (FilterRTID is sometimes the rule Name)
                    try {
                        $r = Get-NetFirewallRule -Name $filterName -ErrorAction SilentlyContinue
                        if ($r) {
                            $ruleName = $r.DisplayName
                            $script:RuleNameCache[$filterName] = $ruleName
                        }
                    } catch {}
                }

                # Last resort: try to match via the port filter rules
                if (-not $ruleName) {
                    try {
                        $portFilters = Get-NetFirewallPortFilter -ErrorAction SilentlyContinue | Where-Object {
                            ($_.LocalPort -contains $props['DestPort']) -or ($_.LocalPort -contains $props['SourcePort'])
                        }
                        if ($portFilters) {
                            $associatedRule = $portFilters | ForEach-Object {
                                $_ | Get-NetFirewallRule -ErrorAction SilentlyContinue
                            } | Select-Object -First 1
                            if ($associatedRule) {
                                $ruleName = $associatedRule.DisplayName
                            }
                        }
                    } catch {}
                }

                if (-not $ruleName) { $ruleName = "FilterRTID:$filterName" }

                $script:EventRuleMap[$key] = $ruleName

                # Extract the Application Name from the event details
                $appName = $props['Application']
                if ($appName) {
                    # The Application field is a device path like \device\harddiskvolume3\windows\system32\svchost.exe
                    # Convert to a friendly name: just the filename, or keep full path
                    $appName = $appName -replace '^\\device\\harddiskvolume\d+', ''
                    $appName = $appName.TrimStart('\')
                }
                else {
                    $appName = ''
                }
                $script:EventAppMap[$key] = $appName

                # Also store a reverse-direction key for SEND vs RECEIVE matching
                $reverseKey = "$proto-$($props['DestAddress'])-$($props['DestPort'])-$($props['SourceAddress'])-$($props['SourcePort'])"
                if (-not $script:EventRuleMap.ContainsKey($reverseKey)) {
                    $script:EventRuleMap[$reverseKey] = $ruleName
                }
                if (-not $script:EventAppMap.ContainsKey($reverseKey)) {
                    $script:EventAppMap[$reverseKey] = $appName
                }
            } catch {
                # Skip unparseable events
                continue
            }
        }

        # Evict old entries (keep map manageable)
        if ($script:EventRuleMap.Count -gt 50000) {
            $keysToRemove = @($script:EventRuleMap.Keys | Select-Object -First 25000)
            foreach ($k in $keysToRemove) {
                $script:EventRuleMap.Remove($k)
                $script:EventAppMap.Remove($k)
            }
        }
    }
    catch {
        Write-Warning "Could not read Security Event Log firewall events: $_"
    }
}

function Resolve-RuleNameForEntry {
    <#
    .SYNOPSIS Looks up the firewall rule name and application name for a parsed
             log entry using the 5-tuple key built from Security Event Log events.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Entry)

    $proto = $Entry.Protocol.ToUpper()

    # Try source->dest key
    $key1 = "$proto-$($Entry.SrcIP)-$($Entry.SrcPort)-$($Entry.DstIP)-$($Entry.DstPort)"
    if ($script:EventRuleMap.ContainsKey($key1)) {
        $Entry.RuleName     = $script:EventRuleMap[$key1]
        $Entry.EventAppName = if ($script:EventAppMap.ContainsKey($key1)) { $script:EventAppMap[$key1] } else { '' }
        return $Entry
    }

    # Try dest->source (reverse)
    $key2 = "$proto-$($Entry.DstIP)-$($Entry.DstPort)-$($Entry.SrcIP)-$($Entry.SrcPort)"
    if ($script:EventRuleMap.ContainsKey($key2)) {
        $Entry.RuleName     = $script:EventRuleMap[$key2]
        $Entry.EventAppName = if ($script:EventAppMap.ContainsKey($key2)) { $script:EventAppMap[$key2] } else { '' }
        return $Entry
    }

    $Entry.RuleName     = ''
    $Entry.EventAppName = ''
    return $Entry
}

# ------------------------------------------------------------------
# Process resolution (port → PID → process name)
# ------------------------------------------------------------------
$script:ProcessCache = @{}

function Update-ConnectionProcessMap {
    [CmdletBinding()]
    param()

    $now = Get-Date
    $map = @{}

    try {
        $tcpConns = Get-NetTCPConnection -ErrorAction SilentlyContinue
        foreach ($c in $tcpConns) {
            $key = "TCP-$($c.LocalAddress)-$($c.LocalPort)"
            $map[$key] = @{ PID = $c.OwningProcess; Timestamp = $now }
            $key2 = "TCP-$($c.RemoteAddress)-$($c.RemotePort)-$($c.LocalPort)"
            $map[$key2] = @{ PID = $c.OwningProcess; Timestamp = $now }
        }
    } catch {}

    try {
        $udpEps = Get-NetUDPEndpoint -ErrorAction SilentlyContinue
        foreach ($u in $udpEps) {
            $key = "UDP-$($u.LocalAddress)-$($u.LocalPort)"
            $map[$key] = @{ PID = $u.OwningProcess; Timestamp = $now }
        }
    } catch {}

    $pidNames = @{}
    foreach ($entry in $map.Values) {
        $procId = $entry.PID
        if ($procId -and -not $pidNames.ContainsKey($procId)) {
            try {
                $proc = Get-Process -Id $procId -ErrorAction SilentlyContinue
                $pidNames[$procId] = if ($proc) { $proc.ProcessName } else { "Unknown" }
            } catch { $pidNames[$procId] = "Unknown" }
        }
    }

    foreach ($key in $map.Keys) {
        $entry = $map[$key]
        $entry.ProcessName = $pidNames[$entry.PID]
        $script:ProcessCache[$key] = $entry
    }

    # Evict stale entries (> 2 minutes old)
    $cutoff = $now.AddSeconds(-120)
    $staleKeys = @($script:ProcessCache.Keys | Where-Object { $script:ProcessCache[$_].Timestamp -lt $cutoff })
    foreach ($k in $staleKeys) { $script:ProcessCache.Remove($k) }
}

function Resolve-ProcessForEntry {
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Entry)

    $proto = $Entry.Protocol.ToUpper()
    $dir   = $Entry.Direction

    $keys = @()
    if ($dir -eq "SEND") {
        $keys += "$proto-$($Entry.SrcIP)-$($Entry.SrcPort)"
    }
    elseif ($dir -eq "RECEIVE") {
        $keys += "$proto-$($Entry.DstIP)-$($Entry.DstPort)"
    }
    $keys += "$proto-$($Entry.SrcIP)-$($Entry.SrcPort)"
    $keys += "$proto-$($Entry.DstIP)-$($Entry.DstPort)"

    foreach ($key in $keys) {
        if ($script:ProcessCache.ContainsKey($key)) {
            $match = $script:ProcessCache[$key]
            $Entry.ProcessId   = $match.PID
            $Entry.ProcessName = $match.ProcessName
            return $Entry
        }
    }

    $Entry.ProcessId   = 0
    $Entry.ProcessName = "Unknown"
    return $Entry
}

# ------------------------------------------------------------------
# Log file parsing
# ------------------------------------------------------------------
function Read-NewFirewallLogEntries {
    <#
    .SYNOPSIS Reads new lines from pfirewall.log since the last read position.
    .OUTPUTS Array of hashtables with parsed fields.
    #>
    [CmdletBinding()]
    param([switch]$ResetPosition)

    if ($ResetPosition) { $script:LastFilePos = 0 }

    if (-not (Test-Path $FirewallLogPath)) {
        Write-Warning "Firewall log not found at $FirewallLogPath"
        return @()
    }

    Update-ConnectionProcessMap

    $entries = [System.Collections.Generic.List[hashtable]]::new()
    $computerName = $env:COMPUTERNAME

    try {
        $fs = [System.IO.FileStream]::new(
            $FirewallLogPath,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::Read,
            [System.IO.FileShare]::ReadWrite
        )
        $reader = [System.IO.StreamReader]::new($fs)

        if ($script:LastFilePos -gt 0 -and $script:LastFilePos -le $fs.Length) {
            $fs.Seek($script:LastFilePos, [System.IO.SeekOrigin]::Begin) | Out-Null
        }
        elseif ($script:LastFilePos -gt $fs.Length) {
            $script:LastFilePos = 0
            $fs.Seek(0, [System.IO.SeekOrigin]::Begin) | Out-Null
        }

        while ($null -ne ($line = $reader.ReadLine())) {
            if ($line.StartsWith("#") -or [string]::IsNullOrWhiteSpace($line)) { continue }

            $parsed = ConvertFrom-FirewallLogLine -Line $line -ComputerName $computerName
            if ($parsed) {
                $parsed = Resolve-ProcessForEntry -Entry $parsed
                $entries.Add($parsed)
            }
        }

        $script:LastFilePos = $fs.Position
        $reader.Close()
        $fs.Close()
    }
    catch {
        Write-Warning "Error reading firewall log: $_"
    }

    return $entries.ToArray()
}

function ConvertFrom-FirewallLogLine {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Line,
        [string]$ComputerName = $env:COMPUTERNAME
    )

    # Standard fields: date time action protocol src-ip dst-ip src-port dst-port
    #                   size tcpflags tcpsyn tcpack tcpwin icmptype icmpcode info path
    $fields = $Line -split '\s+'
    if ($fields.Count -lt 17) {
        while ($fields.Count -lt 17) { $fields += "-" }
    }

    $srcPort = 0; [int]::TryParse($fields[6], [ref]$srcPort) | Out-Null
    $dstPort = 0; [int]::TryParse($fields[7], [ref]$dstPort) | Out-Null
    $size    = 0; [int]::TryParse($fields[8], [ref]$size)    | Out-Null

    return @{
        Date         = $fields[0]
        Time         = $fields[1]
        Action       = $fields[2]
        Protocol     = $fields[3]
        SrcIP        = $fields[4]
        DstIP        = $fields[5]
        SrcPort      = $srcPort
        DstPort      = $dstPort
        Size         = $size
        TcpFlags     = $fields[9]
        TcpSyn       = $fields[10]
        TcpAck       = $fields[11]
        TcpWin       = $fields[12]
        IcmpType     = $fields[13]
        IcmpCode     = $fields[14]
        Info         = $fields[15]
        Direction    = $fields[16]
        ProcessId    = 0
        ProcessName  = "Unknown"
        RuleName     = ""
        EventAppName = ""
        ComputerName = $ComputerName
    }
}

# ------------------------------------------------------------------
# CSV output
# ------------------------------------------------------------------
$script:CSVColumns = @(
    "Date","Time","Action","Protocol","SrcIP","DstIP","SrcPort","DstPort",
    "Direction","Size","ProcessName","ProcessId","RuleName","EventAppName","ComputerName"
)

function Write-EntriesToCSV {
    <#
    .SYNOPSIS Appends parsed firewall log entries to a date-stamped CSV file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable[]]$Entries,
        [Parameter(Mandatory)][string]$OutputDirectory
    )

    if ($Entries.Count -eq 0) { return }

    # Create output directory if it does not exist
    if (-not (Test-Path $OutputDirectory)) {
        New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null
    }

    # File name: FirewallLog_YYYY-MM-DD.csv  (one file per day)
    $dateStamp = Get-Date -Format "yyyy-MM-dd"
    $csvPath   = Join-Path $OutputDirectory "FirewallLog_$dateStamp.csv"

    $needsHeader = -not (Test-Path $csvPath)

    # Build rows as PSCustomObjects so Export-Csv works cleanly
    $rows = foreach ($e in $Entries) {
        [PSCustomObject]@{
            Date         = $e.Date
            Time         = $e.Time
            Action       = $e.Action
            Protocol     = $e.Protocol
            SrcIP        = $e.SrcIP
            DstIP        = $e.DstIP
            SrcPort      = $e.SrcPort
            DstPort      = $e.DstPort
            Direction    = $e.Direction
            Size         = $e.Size
            ProcessName  = $e.ProcessName
            ProcessId    = $e.ProcessId
            RuleName     = $e.RuleName
            EventAppName = $e.EventAppName
            ComputerName = $e.ComputerName
        }
    }

    if ($needsHeader) {
        $rows | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Force
    }
    else {
        $rows | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Append
    }

    return $csvPath
}

# ------------------------------------------------------------------
# Scheduled Task management
# ------------------------------------------------------------------
function Install-CollectorTask {
    <#
    .SYNOPSIS Copies this script to a well-known path and registers a
             scheduled task that runs it as SYSTEM at startup.
    #>
    [CmdletBinding()]
    param()

    # --- Deploy script to a persistent, well-known location ---
    if (-not (Test-Path $DeployDir)) {
        New-Item -Path $DeployDir -ItemType Directory -Force | Out-Null
    }

    # Determine the source of this script.  When invoked via
    # Invoke-AzVMRunCommand the engine writes it to a temp folder;
    # $PSCommandPath is still set to that temp path so we can copy it.
    $sourcePath = $PSCommandPath
    if (-not $sourcePath) { $sourcePath = $MyInvocation.ScriptName }
    if (-not $sourcePath) {
        # Last resort: try to locate ourselves via call stack
        $sourcePath = (Get-PSCallStack | Select-Object -Last 1).ScriptName
    }

    if ($sourcePath -and (Test-Path $sourcePath)) {
        Copy-Item -Path $sourcePath -Destination $DeployedScriptPath -Force
        Write-Host "[TASK] Script deployed to $DeployedScriptPath" -ForegroundColor Green
    }
    elseif (Test-Path $DeployedScriptPath) {
        Write-Host "[TASK] Using previously deployed script at $DeployedScriptPath" -ForegroundColor Yellow
    }
    else {
        Write-Error "Could not determine script source path for deployment. Ensure the script exists at $DeployedScriptPath or re-run with a valid -ScriptPath."
        return
    }

    # --- Register the scheduled task ---
    $argString = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$DeployedScriptPath`" -Action Enable -OutputPath `"$OutputPath`" -IntervalSeconds $IntervalSeconds"

    Write-Host "[TASK] Registering scheduled task '$TaskName'..." -ForegroundColor Cyan

    # Remove existing task if present
    $existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if ($existing) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        Write-Host "[TASK] Removed existing task." -ForegroundColor Yellow
    }

    $action  = New-ScheduledTaskAction -Execute "powershell.exe" -Argument $argString
    $trigger = New-ScheduledTaskTrigger -AtStartup
    $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest

    $settings = New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -StartWhenAvailable `
        -RestartCount 3 `
        -RestartInterval (New-TimeSpan -Minutes 1) `
        -ExecutionTimeLimit ([TimeSpan]::Zero)

    Register-ScheduledTask `
        -TaskName $TaskName `
        -Action $action `
        -Trigger $trigger `
        -Principal $principal `
        -Settings $settings `
        -Description "Collects Windows Firewall log entries and writes date-stamped CSV files. Enables firewall logging and Security Event Log auditing." |
        Out-Null

    Write-Host "[TASK] Scheduled task '$TaskName' registered successfully." -ForegroundColor Green
    Write-Host "       Runs as:    SYSTEM (at startup, no user login required)" -ForegroundColor Gray
    Write-Host "       Script:     $DeployedScriptPath" -ForegroundColor Gray
    Write-Host "       Output:     $OutputPath" -ForegroundColor Gray
    Write-Host "       Interval:   ${IntervalSeconds}s" -ForegroundColor Gray
}

function Uninstall-CollectorTask {
    <#
    .SYNOPSIS Removes the scheduled task and disables logging.
    #>
    [CmdletBinding()]
    param()

    Write-Host "[TASK] Removing scheduled task '$TaskName'..." -ForegroundColor Yellow

    $existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if ($existing) {
        # Stop if running
        if ($existing.State -eq 'Running') {
            Stop-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
        }
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        Write-Host "[TASK] Scheduled task removed." -ForegroundColor Green
    }
    else {
        Write-Host "[TASK] No scheduled task found with name '$TaskName'." -ForegroundColor Gray
    }

    Disable-FirewallAndAuditLogging
}

# ------------------------------------------------------------------
# Single-pass collection (for Invoke-AzVMRunCommand / one-shot use)
# ------------------------------------------------------------------
function Invoke-SingleCollection {
    <#
    .SYNOPSIS Enables logging, performs ONE collection pass, writes CSV, then exits.
             Safe for Invoke-AzVMRunCommand which expects the script to complete.
    #>
    [CmdletBinding()]
    param()

    if (-not (Test-Path $OutputPath)) {
        New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
    }

    Enable-FirewallAndAuditLogging

    Write-Host "[INIT] Building firewall rule name cache..." -ForegroundColor Gray
    Update-FirewallRuleCache
    Write-Host "[INIT] Cached $($script:RuleNameCache.Count) firewall rule(s)." -ForegroundColor Green

    # Read everything currently in the log
    $null = Read-NewFirewallLogEntries -ResetPosition

    # Brief pause so the log and Security events have fresh data
    Write-Host "[COLLECT] Waiting $IntervalSeconds seconds for new firewall activity..." -ForegroundColor Cyan
    Start-Sleep -Seconds $IntervalSeconds

    # Collect Security Event Log rule mappings
    Read-SecurityFirewallEvents

    # Read new entries
    $entries = Read-NewFirewallLogEntries
    for ($i = 0; $i -lt $entries.Count; $i++) {
        $entries[$i] = Resolve-RuleNameForEntry -Entry $entries[$i]
    }

    if ($entries.Count -gt 0) {
        $csvFile = Write-EntriesToCSV -Entries $entries -OutputDirectory $OutputPath
        Write-Host "[COLLECT] Wrote $($entries.Count) entries to $(Split-Path $csvFile -Leaf)" -ForegroundColor Green
    }
    else {
        Write-Host "[COLLECT] No new entries found." -ForegroundColor Gray
    }

    Write-Host "[COLLECT] Single collection pass complete." -ForegroundColor Green
}

# ------------------------------------------------------------------
# Main collection loop (continuous — for scheduled task)
# ------------------------------------------------------------------
function Start-CollectionLoop {
    [CmdletBinding()]
    param()

    # Ensure output directory exists
    if (-not (Test-Path $OutputPath)) {
        New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
    }

    Enable-FirewallAndAuditLogging

    # Pre-load firewall rule name cache
    Write-Host "[INIT] Building firewall rule name cache..." -ForegroundColor Gray
    Update-FirewallRuleCache
    Write-Host "[INIT] Cached $($script:RuleNameCache.Count) firewall rule(s)." -ForegroundColor Green

    # Initialize file position (skip existing content)
    $null = Read-NewFirewallLogEntries -ResetPosition
    # Seek to end so we only capture new entries from this point
    if (Test-Path $FirewallLogPath) {
        try {
            $fi = [System.IO.FileInfo]::new($FirewallLogPath)
            $script:LastFilePos = $fi.Length
        } catch {}
    }

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  Firewall Log Collector"                 -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  Output:   $OutputPath"                  -ForegroundColor Gray
    Write-Host "  Interval: ${IntervalSeconds}s"          -ForegroundColor Gray
    Write-Host "  Log:      $FirewallLogPath"             -ForegroundColor Gray
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Collecting firewall log entries. Press Ctrl+C to stop." -ForegroundColor Cyan
    Write-Host ""

    $totalEntries = 0

    # Graceful shutdown
    trap {
        Write-Host "`n[EXIT] Shutting down collector..." -ForegroundColor Yellow
        Write-Host "[EXIT] Total entries collected: $totalEntries" -ForegroundColor Green
        break
    }

    while ($true) {
        $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

        try {
            # Read Security Event Log to get rule names for recent connections
            Read-SecurityFirewallEvents

            $entries = Read-NewFirewallLogEntries

            # Enrich each entry with the firewall rule name
            for ($i = 0; $i -lt $entries.Count; $i++) {
                $entries[$i] = Resolve-RuleNameForEntry -Entry $entries[$i]
            }

            $count   = $entries.Count

            if ($count -gt 0) {
                $totalEntries += $count
                $csvFile = Write-EntriesToCSV -Entries $entries -OutputDirectory $OutputPath
                Write-Host "[$ts] Wrote $count entries to $(Split-Path $csvFile -Leaf). Total: $totalEntries" -ForegroundColor Green
            }
            else {
                Write-Host "[$ts] No new entries." -ForegroundColor Gray
            }
        }
        catch {
            Write-Host "[$ts] ERROR: $_" -ForegroundColor Red
        }

        Start-Sleep -Seconds $IntervalSeconds
    }
}

# ==================================================================
# Entry point
# ==================================================================
Assert-Administrator

switch ($Action) {
    'Disable' {
        Disable-FirewallAndAuditLogging
        Write-Host ""
        Write-Host "Firewall logging and Security Event Log auditing have been disabled." -ForegroundColor Green
    }
    'Install' {
        Install-CollectorTask
    }
    'Uninstall' {
        Uninstall-CollectorTask
    }
    'Collect' {
        Invoke-SingleCollection
    }
    default {
        # 'Enable' — continuous collection loop (for scheduled task / interactive use)
        Start-CollectionLoop
    }
}
