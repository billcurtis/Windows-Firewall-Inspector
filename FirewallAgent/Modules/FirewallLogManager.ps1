<#
.SYNOPSIS
    Windows Firewall Log Manager Module.
.DESCRIPTION
    Enables/disables Windows Firewall logging, parses the log file incrementally,
    and resolves active connections to process names (PID -> executable).
#>

# ---------- State ----------
$script:OriginalLoggingState = @{}
$script:LastFilePosition     = 0
$script:LogFilePath          = "$env:SystemRoot\System32\LogFiles\Firewall\pfirewall.log"
$script:ProcessCache         = @{}          # port-key -> @{PID; ProcessName; Timestamp}
$script:ProcessCacheMaxAge   = 30           # seconds
$script:LastEventTime        = $null        # tracks Security Event Log read position
$script:RuleNameCache        = @{}          # FilterId -> display name
$script:EventRuleMap         = @{}          # "proto-srcip-srcport-dstip-dstport" -> RuleName
$script:EventAppMap          = @{}          # "proto-srcip-srcport-dstip-dstport" -> Application path

# ---------- Firewall Logging Control ----------

function Save-FirewallLoggingState {
    <# Captures the current logging settings for all profiles so they can be restored later. #>
    [CmdletBinding()]
    param()

    $profiles = @("domainprofile", "privateprofile", "publicprofile")
    foreach ($p in $profiles) {
        $output = & netsh advfirewall show $p logging 2>&1 | Out-String
        $state  = @{
            AllowedConnections = ($output -match "LogAllowedConnections\s+Enable")
            DroppedConnections = ($output -match "LogDroppedConnections\s+Enable")
        }
        $script:OriginalLoggingState[$p] = $state
    }
    return $script:OriginalLoggingState
}

function Enable-FirewallLogging {
    <# Enables full firewall logging (allowed + dropped) on all profiles. Requires elevation. #>
    [CmdletBinding()]
    param([string]$LogPath = $script:LogFilePath)

    Save-FirewallLoggingState | Out-Null

    $cmds = @(
        "netsh advfirewall set allprofiles logging filename `"$LogPath`"",
        "netsh advfirewall set allprofiles logging maxfilesize 32767",
        "netsh advfirewall set allprofiles logging droppedconnections enable",
        "netsh advfirewall set allprofiles logging allowedconnections enable"
    )

    foreach ($cmd in $cmds) {
        $result = Invoke-Expression $cmd 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Warning "Command failed: $cmd`n$result"
        }
    }
    Write-Verbose "Firewall logging enabled on all profiles."

    # Enable Security Event Log auditing for Windows Filtering Platform
    Write-Verbose "Enabling Security Event Log auditing for firewall events..."
    & auditpol /set /subcategory:"Filtering Platform Packet Drop" /success:enable /failure:enable 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) { Write-Warning "Failed to set audit policy for Filtering Platform Packet Drop." }
    & auditpol /set /subcategory:"Filtering Platform Connection" /success:enable /failure:enable 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) { Write-Warning "Failed to set audit policy for Filtering Platform Connection." }
    Write-Verbose "Security Event Log auditing enabled (Success + Failure)."
}

function Restore-FirewallLogging {
    <# Restores the firewall logging settings that were saved at startup. #>
    [CmdletBinding()]
    param()

    foreach ($p in $script:OriginalLoggingState.Keys) {
        $s = $script:OriginalLoggingState[$p]
        $allowed = if ($s.AllowedConnections) { "enable" } else { "disable" }
        $dropped = if ($s.DroppedConnections) { "enable" } else { "disable" }

        & netsh advfirewall set $p logging allowedconnections $allowed 2>&1 | Out-Null
        & netsh advfirewall set $p logging droppedconnections $dropped 2>&1 | Out-Null
    }

    # Disable Security Event Log auditing
    & auditpol /set /subcategory:"Filtering Platform Packet Drop" /success:disable /failure:disable 2>&1 | Out-Null
    & auditpol /set /subcategory:"Filtering Platform Connection" /success:disable /failure:disable 2>&1 | Out-Null
    Write-Verbose "Security Event Log auditing disabled."
    Write-Verbose "Firewall logging settings restored."
}

# ---------- Process Resolution ----------

function Update-ConnectionProcessMap {
    <# Refreshes the in-memory map of local-port -> (PID, ProcessName). #>
    [CmdletBinding()]
    param()

    $now = Get-Date
    $map = @{}

    # TCP connections
    try {
        $tcpConns = Get-NetTCPConnection -ErrorAction SilentlyContinue
        foreach ($c in $tcpConns) {
            $key = "TCP-$($c.LocalAddress)-$($c.LocalPort)"
            $map[$key] = @{ PID = $c.OwningProcess; Timestamp = $now }
            $key2 = "TCP-$($c.RemoteAddress)-$($c.RemotePort)-$($c.LocalPort)"
            $map[$key2] = @{ PID = $c.OwningProcess; Timestamp = $now }
        }
    } catch {}

    # UDP endpoints
    try {
        $udpEps = Get-NetUDPEndpoint -ErrorAction SilentlyContinue
        foreach ($u in $udpEps) {
            $key = "UDP-$($u.LocalAddress)-$($u.LocalPort)"
            $map[$key] = @{ PID = $u.OwningProcess; Timestamp = $now }
        }
    } catch {}

    # Resolve process names
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

    # Merge into cache
    foreach ($key in $map.Keys) {
        $entry = $map[$key]
        $entry.ProcessName = $pidNames[$entry.PID]
        $script:ProcessCache[$key] = $entry
    }

    # Evict stale entries
    $cutoff = $now.AddSeconds(-120)
    $staleKeys = $script:ProcessCache.Keys | Where-Object { $script:ProcessCache[$_].Timestamp -lt $cutoff }
    foreach ($k in $staleKeys) { $script:ProcessCache.Remove($k) }
}

function Resolve-LogEntryProcess {
    <#
    .SYNOPSIS Attempts to resolve the PID and process name for a parsed firewall log entry.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Entry)

    $proto = $Entry.Protocol.ToUpper()
    $dir   = $Entry.Direction

    # Build lookup keys based on direction
    $keys = @()
    if ($dir -eq "SEND") {
        $keys += "$proto-$($Entry.SrcIP)-$($Entry.SrcPort)"
    }
    elseif ($dir -eq "RECEIVE") {
        $keys += "$proto-$($Entry.DstIP)-$($Entry.DstPort)"
    }
    # Fallback keys
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

# ---------- Security Event Log → Rule Name & App Name Resolution ----------

function Update-FirewallRuleCache {
    <# Builds/refreshes a hashtable of WFP Filter Run-Time IDs to friendly firewall rule display names. #>
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
             and populates lookup tables keyed on connection 5-tuple so that
             pfirewall.log entries can be enriched with rule name and app name.
    #>
    [CmdletBinding()]
    param()

    # Only read events since our last check (or last 90 seconds on first run)
    $since = if ($script:LastEventTime) { $script:LastEventTime } else { (Get-Date).AddSeconds(-90) }
    $script:LastEventTime = Get-Date

    try {
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
                $ruleName   = ''

                if ($filterName -and $script:RuleNameCache.ContainsKey($filterName)) {
                    $ruleName = $script:RuleNameCache[$filterName]
                }

                if (-not $ruleName -and $filterName) {
                    try {
                        $r = Get-NetFirewallRule -Name $filterName -ErrorAction SilentlyContinue
                        if ($r) {
                            $ruleName = $r.DisplayName
                            $script:RuleNameCache[$filterName] = $ruleName
                        }
                    } catch {}
                }

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
                    $appName = $appName -replace '^\\device\\harddiskvolume\d+', ''
                    $appName = $appName.TrimStart('\\')
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

# ---------- Log Parsing ----------

function Get-NewFirewallLogEntries {
    <#
    .SYNOPSIS Reads new lines from the firewall log file since the last read position.
    .OUTPUTS Array of hashtables, each representing a parsed log entry.
    #>
    [CmdletBinding()]
    param([switch]$ResetPosition)

    if ($ResetPosition) { $script:LastFilePosition = 0 }

    if (-not (Test-Path $script:LogFilePath)) {
        Write-Warning "Firewall log not found at $($script:LogFilePath)"
        return @()
    }

    # Refresh PID map and Security Event Log data before parsing
    Update-ConnectionProcessMap
    Read-SecurityFirewallEvents

    $entries = [System.Collections.Generic.List[hashtable]]::new()

    try {
        $fs = [System.IO.FileStream]::new(
            $script:LogFilePath,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::Read,
            [System.IO.FileShare]::ReadWrite   # critical - file is locked by Windows Firewall
        )
        $reader = [System.IO.StreamReader]::new($fs)

        # Seek to last known position
        if ($script:LastFilePosition -gt 0 -and $script:LastFilePosition -le $fs.Length) {
            $fs.Seek($script:LastFilePosition, [System.IO.SeekOrigin]::Begin) | Out-Null
        }
        elseif ($script:LastFilePosition -gt $fs.Length) {
            # File was rotated / truncated - start from beginning
            $script:LastFilePosition = 0
            $fs.Seek(0, [System.IO.SeekOrigin]::Begin) | Out-Null
        }

        $computerName = $env:COMPUTERNAME
        $lineCount    = 0

        while ($null -ne ($line = $reader.ReadLine())) {
            # Skip comments / headers
            if ($line.StartsWith("#") -or [string]::IsNullOrWhiteSpace($line)) { continue }

            $parsed = ConvertFrom-FirewallLogLine -Line $line -ComputerName $computerName
            if ($parsed) {
                $parsed = Resolve-LogEntryProcess -Entry $parsed
                $parsed = Resolve-RuleNameForEntry -Entry $parsed
                $entries.Add($parsed)
            }
            $lineCount++
        }

        $script:LastFilePosition = $fs.Position
        $reader.Close()
        $fs.Close()
    }
    catch {
        Write-Warning "Error reading firewall log: $_"
    }

    return $entries.ToArray()
}

function ConvertFrom-FirewallLogLine {
    <#
    .SYNOPSIS Parses a single line from the Windows Firewall log into a hashtable.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Line,
        [string]$ComputerName = $env:COMPUTERNAME
    )

    # Standard fields: date time action protocol src-ip dst-ip src-port dst-port size tcpflags tcpsyn tcpack tcpwin icmptype icmpcode info path
    $fields = $Line -split '\s+'
    if ($fields.Count -lt 17) {
        # Might be a short line; pad with dashes
        while ($fields.Count -lt 17) { $fields += "-" }
    }

    $date      = $fields[0]
    $time      = $fields[1]
    $action    = $fields[2]
    $protocol  = $fields[3]
    $srcIP     = $fields[4]
    $dstIP     = $fields[5]
    $srcPort   = 0; [int]::TryParse($fields[6], [ref]$srcPort) | Out-Null
    $dstPort   = 0; [int]::TryParse($fields[7], [ref]$dstPort) | Out-Null
    $size      = 0; [int]::TryParse($fields[8], [ref]$size)    | Out-Null
    $tcpFlags  = $fields[9]
    $tcpSyn    = $fields[10]
    $tcpAck    = $fields[11]
    $tcpWin    = $fields[12]
    $icmpType  = $fields[13]
    $icmpCode  = $fields[14]
    $info      = $fields[15]
    $direction = $fields[16]

    # Generate unique RowKey
    $ts     = "$date`T$time" -replace '[^0-9T]', ''
    $rowKey = "$ts-$([guid]::NewGuid().ToString('N').Substring(0,8))"

    return @{
        PartitionKey = "${ComputerName}_$date"
        RowKey       = $rowKey
        Date         = $date
        Time         = $time
        Action       = $action
        Protocol     = $protocol
        SrcIP        = $srcIP
        DstIP        = $dstIP
        SrcPort      = $srcPort
        DstPort      = $dstPort
        Size         = $size
        TcpFlags     = $tcpFlags
        TcpSyn       = $tcpSyn
        TcpAck       = $tcpAck
        TcpWin       = $tcpWin
        IcmpType     = $icmpType
        IcmpCode     = $icmpCode
        Info         = $info
        Direction    = $direction
        ProcessId    = 0
        ProcessName  = "Unknown"
        RuleName     = ""
        EventAppName = ""
        ComputerName = $ComputerName
    }
}
