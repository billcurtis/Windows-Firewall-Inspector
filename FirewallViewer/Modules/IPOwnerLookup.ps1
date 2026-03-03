<#
.SYNOPSIS
    IP Owner Lookup Module using ip-api.com batch API.
.DESCRIPTION
    Resolves IP addresses to their owning organization using the free ip-api.com service.
    Batch endpoint processes up to 100 IPs per request (counts as 1 API call).
    Results are cached to avoid redundant lookups.
#>

$script:IPOwnerCache = @{}

function Test-PrivateIP {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$IPAddress)

    if ([string]::IsNullOrEmpty($IPAddress) -or $IPAddress -eq '-' -or $IPAddress -eq '*') { return $true }

    try {
        $ip = [System.Net.IPAddress]::Parse($IPAddress)
        $bytes = $ip.GetAddressBytes()

        if ($bytes.Count -gt 4) { return ($IPAddress -eq '::1') }

        $first  = [int]$bytes[0]
        $second = [int]$bytes[1]

        if ($first -eq 10)  { return $true }
        if ($first -eq 172 -and $second -ge 16 -and $second -le 31) { return $true }
        if ($first -eq 192 -and $second -eq 168) { return $true }
        if ($first -eq 127) { return $true }
        if ($first -eq 169 -and $second -eq 254) { return $true }
        if ($first -eq 0)   { return $true }
        if ($first -eq 100 -and $second -ge 64 -and $second -le 127) { return $true }
        if ($first -ge 224) { return $true }
        if ($first -eq 168 -and $second -eq 63 -and [int]$bytes[2] -eq 129 -and [int]$bytes[3] -eq 16) { return $true }

        return $false
    }
    catch { return $false }
}

function Resolve-IPOwners {
    <#
    .SYNOPSIS Batch-resolves IP owners. Returns hashtable of IP -> OwnerInfo.
    .DESCRIPTION
        Uses ip-api.com batch endpoint (up to 100 IPs per POST).
        Private/reserved IPs are resolved locally. Results are cached.
        Output objects have a .Name property for compatibility with the viewer.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$IPAddresses
    )

    $unique  = @($IPAddresses | Select-Object -Unique)
    $results = @{}
    $publicIPs = [System.Collections.ArrayList]::new()

    # Separate cached / private / public
    foreach ($ip in $unique) {
        if ($script:IPOwnerCache.ContainsKey($ip)) {
            $results[$ip] = $script:IPOwnerCache[$ip]
            continue
        }
        if (Test-PrivateIP -IPAddress $ip) {
            $entry = [PSCustomObject]@{
                IP      = $ip
                Name    = "Private/Reserved"
                ISP     = "N/A"
                Country = "N/A"
                City    = "N/A"
                Status  = "private"
            }
            $script:IPOwnerCache[$ip] = $entry
            $results[$ip] = $entry
            continue
        }
        $null = $publicIPs.Add($ip)
    }

    if ($publicIPs.Count -eq 0) { return $results }

    # Process in batches of 100
    $batchSize = 100
    for ($i = 0; $i -lt $publicIPs.Count; $i += $batchSize) {
        $end   = [Math]::Min($i + $batchSize - 1, $publicIPs.Count - 1)
        $batch = $publicIPs[$i..$end]

        try {
            $body = ($batch | ForEach-Object {
                @{ query = $_; fields = "status,message,country,city,isp,org,as,query" }
            })
            $json = $body | ConvertTo-Json -Depth 3 -Compress
            if ($batch.Count -eq 1) { $json = "[$json]" }

            $resp = Invoke-RestMethod -Uri "http://ip-api.com/batch?fields=status,message,country,city,isp,org,as,query" `
                        -Method Post -ContentType "application/json" -Body $json -TimeoutSec 30 -ErrorAction Stop

            foreach ($r in $resp) {
                $ipAddr = $r.query
                if ($r.status -eq 'success') {
                    $entry = [PSCustomObject]@{
                        IP      = $ipAddr
                        Name    = if ($r.org) { $r.org } else { $r.isp }
                        ISP     = $r.isp
                        Country = $r.country
                        City    = $r.city
                        Status  = "success"
                    }
                }
                else {
                    $entry = [PSCustomObject]@{
                        IP      = $ipAddr
                        Name    = "Lookup failed"
                        ISP     = "N/A"
                        Country = "N/A"
                        City    = "N/A"
                        Status  = "failed"
                    }
                }
                $script:IPOwnerCache[$ipAddr] = $entry
                $results[$ipAddr] = $entry
            }
        }
        catch {
            foreach ($ipAddr in $batch) {
                if (-not $results.ContainsKey($ipAddr)) {
                    $entry = [PSCustomObject]@{
                        IP      = $ipAddr
                        Name    = "Batch error"
                        ISP     = "N/A"
                        Country = "N/A"
                        City    = "N/A"
                        Status  = "error"
                    }
                    $script:IPOwnerCache[$ipAddr] = $entry
                    $results[$ipAddr] = $entry
                }
            }
        }

        # Rate-limit pause between batches (stay under 45 req/min)
        if ($i + $batchSize -lt $publicIPs.Count) {
            Start-Sleep -Milliseconds 1500
        }
    }

    return $results
}

function Clear-IPOwnerCache {
    [CmdletBinding()]
    param()
    $script:IPOwnerCache.Clear()
}
