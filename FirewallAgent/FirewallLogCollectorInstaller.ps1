<#
    Toggle this before running the script.
    $true  = install the scheduled collector
    $false = uninstall the scheduled collector
#>

#Requires -Version 5.1

$install = $true

$taskName = 'FirewallBlockedCollector'
$deployDir = 'C:\ProgramData\FirewallBlockedCollector'
$runtimePath = Join-Path $deployDir 'FirewallBlockedCollector.ps1'
$outputPath = Join-Path $deployDir 'Logs'
$intervalSeconds = 60
$firewallLogPath = "$env:SystemRoot\System32\LogFiles\Firewall\pfirewall.log"

function Assert-Administrator {
    $principal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw 'This script must be run as Administrator.'
    }
}

function Enable-BlockedOnlyLogging {
    & netsh advfirewall set allprofiles logging filename "$firewallLogPath" | Out-Null
    & netsh advfirewall set allprofiles logging maxfilesize 32767 | Out-Null
    & netsh advfirewall set allprofiles logging allowedconnections disable | Out-Null
    & netsh advfirewall set allprofiles logging droppedconnections enable | Out-Null
}

function Disable-BlockedOnlyLogging {
    & netsh advfirewall set allprofiles logging allowedconnections disable | Out-Null
    & netsh advfirewall set allprofiles logging droppedconnections disable | Out-Null
}

function Get-RuntimeScript {
    return @"
`$outputPath = '$outputPath'
`$firewallLogPath = '$firewallLogPath'
`$intervalSeconds = $intervalSeconds
`$lastFilePos = 0
`$processCache = @{}

if (-not (Test-Path `$outputPath)) {
    New-Item -Path `$outputPath -ItemType Directory -Force | Out-Null
}

function Update-ProcessCache {
    `$now = Get-Date
    `$connections = @{}

    try {
        foreach (`$conn in Get-NetTCPConnection -ErrorAction SilentlyContinue) {
            `$key = "TCP-`$(`$conn.LocalAddress)-`$(`$conn.LocalPort)"
            `$connections[`$key] = `$conn.OwningProcess
        }
    }
    catch {
    }

    try {
        foreach (`$endpoint in Get-NetUDPEndpoint -ErrorAction SilentlyContinue) {
            `$key = "UDP-`$(`$endpoint.LocalAddress)-`$(`$endpoint.LocalPort)"
            `$connections[`$key] = `$endpoint.OwningProcess
        }
    }
    catch {
    }

    foreach (`$key in `$connections.Keys) {
        `$pid = `$connections[`$key]
        `$processName = 'Unknown'
        `$executablePath = ''

        if (`$pid) {
            try {
                `$process = Get-Process -Id `$pid -ErrorAction SilentlyContinue
                if (`$process) {
                    `$processName = `$process.ProcessName
                    `$executablePath = `$process.Path
                    if (-not `$executablePath) {
                        `$executablePath = `$processName + '.exe'
                    }
                }
            }
            catch {
            }
        }

        `$processCache[`$key] = @{
            ProcessId = `$pid
            ProcessName = `$processName
            ExecutablePath = `$executablePath
            Timestamp = `$now
        }
    }

    `$cutoff = `$now.AddMinutes(-2)
    foreach (`$staleKey in @(`$processCache.Keys | Where-Object { `$processCache[`$_].Timestamp -lt `$cutoff })) {
        `$processCache.Remove(`$staleKey)
    }
}

function Resolve-ProcessInfo {
    param(
        [string]`$Protocol,
        [string]`$SrcIP,
        [int]`$SrcPort,
        [string]`$DstIP,
        [int]`$DstPort
    )

    `$keys = @(
        "`$Protocol-`$SrcIP-`$SrcPort",
        "`$Protocol-`$DstIP-`$DstPort",
        "`$Protocol-0.0.0.0-`$SrcPort",
        "`$Protocol-0.0.0.0-`$DstPort",
        "`$Protocol-::-`$SrcPort",
        "`$Protocol-::-`$DstPort"
    )

    foreach (`$key in `$keys) {
        if (`$processCache.ContainsKey(`$key)) {
            return `$processCache[`$key]
        }
    }

    return @{
        ProcessId = 0
        ProcessName = 'Unknown'
        ExecutablePath = ''
    }
}

function Read-NewDropLines {
    if (-not (Test-Path `$firewallLogPath)) {
        return @()
    }

    Update-ProcessCache
    `$lines = [System.Collections.Generic.List[object]]::new()
    `$stream = [System.IO.FileStream]::new(`$firewallLogPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    `$reader = [System.IO.StreamReader]::new(`$stream)

    if (`$lastFilePos -gt 0 -and `$lastFilePos -le `$stream.Length) {
        `$stream.Seek(`$lastFilePos, [System.IO.SeekOrigin]::Begin) | Out-Null
    }

    while (`$null -ne (`$line = `$reader.ReadLine())) {
        if (`$line.StartsWith('#') -or [string]::IsNullOrWhiteSpace(`$line)) {
            continue
        }

        `$fields = `$line -split '\s+'
        if (`$fields.Count -lt 8 -or `$fields[2] -ne 'DROP') {
            continue
        }

        `$srcPort = 0
        [int]::TryParse(`$fields[6], [ref]`$srcPort) | Out-Null
        `$dstPort = 0
        [int]::TryParse(`$fields[7], [ref]`$dstPort) | Out-Null
        `$processInfo = Resolve-ProcessInfo -Protocol `$fields[3] -SrcIP `$fields[4] -SrcPort `$srcPort -DstIP `$fields[5] -DstPort `$dstPort

        `$lines.Add(([pscustomobject]@{
            Date = `$fields[0]
            Time = `$fields[1]
            Action = `$fields[2]
            Protocol = `$fields[3]
            SrcIP = `$fields[4]
            DstIP = `$fields[5]
            SrcPort = `$srcPort
            DstPort = `$dstPort
            ProcessId = `$processInfo.ProcessId
            ProcessName = `$processInfo.ProcessName
            ExecutablePath = `$processInfo.ExecutablePath
            ComputerName = `$env:COMPUTERNAME
        }))
    }

    `$lastFilePos = `$stream.Position
    `$reader.Close()
    `$stream.Close()
    return `$lines.ToArray()
}

`$null = Read-NewDropLines
if (Test-Path `$firewallLogPath) {
    `$lastFilePos = ([System.IO.FileInfo]::new(`$firewallLogPath)).Length
}

while (`$true) {
    try {
        `$rows = Read-NewDropLines
        if (`$rows.Count -gt 0) {
            `$csvPath = Join-Path `$outputPath ('FirewallBlocked_' + (Get-Date -Format 'yyyy-MM-dd') + '.csv')
            if (Test-Path `$csvPath) {
                `$rows | Export-Csv -Path `$csvPath -NoTypeInformation -Encoding UTF8 -Append
            }
            else {
                `$rows | Export-Csv -Path `$csvPath -NoTypeInformation -Encoding UTF8
            }
        }
    }
    catch {
    }

    Start-Sleep -Seconds `$intervalSeconds
}
"@
}

function Install-Collector {
    if (-not (Test-Path $deployDir)) {
        New-Item -Path $deployDir -ItemType Directory -Force | Out-Null
    }

    Set-Content -Path $runtimePath -Value (Get-RuntimeScript) -Encoding UTF8
    Enable-BlockedOnlyLogging

    $existing = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    if ($existing) {
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    }

    $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$runtimePath`""
    $trigger = New-ScheduledTaskTrigger -AtStartup
    $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -ExecutionTimeLimit ([TimeSpan]::Zero)

    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description 'Collect blocked Windows Firewall traffic to CSV.' | Out-Null
    Write-Host "Installed $taskName" -ForegroundColor Green
}

function Uninstall-Collector {
    $existing = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    if ($existing) {
        if ($existing.State -eq 'Running') {
            Stop-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
        }
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    }

    Disable-BlockedOnlyLogging
    if (Test-Path $runtimePath) {
        Remove-Item -Path $runtimePath -Force
    }
    Write-Host "Uninstalled $taskName" -ForegroundColor Yellow
}

Assert-Administrator

if ($install) {
    Install-Collector
}
else {
    Uninstall-Collector
}