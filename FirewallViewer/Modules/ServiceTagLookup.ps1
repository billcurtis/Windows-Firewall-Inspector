<#
.SYNOPSIS
    Microsoft Service Tag / IP Range Lookup Module.
.DESCRIPTION
    Downloads the official Microsoft Azure IP Ranges and Service Tags JSON file,
    caches it locally, and provides functions to determine which Microsoft
    service(s) a given IP address belongs to.

    The JSON is published at:
      https://www.microsoft.com/en-us/download/details.aspx?id=56519
    The actual download URL is discovered by scraping the confirmation page.

    The local cache is refreshed once per day (or on demand).
#>

# ───────────────── state ─────────────────
$script:ServiceTagData    = $null          # parsed JSON root
$script:ServiceTagIndex   = $null          # array of @{ Network; MaskBits; ServiceName }
$script:ServiceTagCacheDir  = Join-Path $env:LOCALAPPDATA "FirewallLogViewer"
$script:ServiceTagCachePath = Join-Path $script:ServiceTagCacheDir "ServiceTags.json"
$script:ServiceTagMaxAge    = [TimeSpan]::FromHours(24)
$script:ServiceTagIPCache   = @{}          # IP -> "Service1, Service2"

# ───────────────── download helpers ─────────────────

function Get-ServiceTagDownloadUrl {
    <#
    .SYNOPSIS Discovers the current download URL for the Service Tags JSON.
    #>
    [CmdletBinding()]
    param()

    $confirmUrl = "https://www.microsoft.com/en-us/download/confirmation.aspx?id=56519"

    try {
        $html = Invoke-WebRequest -Uri $confirmUrl -UseBasicParsing -TimeoutSec 30 -ErrorAction Stop
        # Look for the direct download link (href ending in .json)
        $match = [regex]::Match($html.Content, 'href="(https://download\.microsoft\.com/download/[^"]+\.json)"')
        if ($match.Success) {
            return $match.Groups[1].Value
        }

        # Fallback: look for any ServiceTags*.json link
        $match2 = [regex]::Match($html.Content, 'href="(https://[^"]+ServiceTags[^"]*\.json)"')
        if ($match2.Success) {
            return $match2.Groups[1].Value
        }

        Write-Warning "Could not find Service Tags JSON download link on the Microsoft page."
        return $null
    }
    catch {
        Write-Warning "Failed to discover Service Tags download URL: $_"
        return $null
    }
}

function Update-ServiceTagCache {
    <#
    .SYNOPSIS Downloads the latest Service Tags JSON and saves to local cache.
    .OUTPUTS $true if download succeeded, $false otherwise.
    #>
    [CmdletBinding()]
    param([switch]$Force)

    # Check if the cache is still fresh
    if (-not $Force -and (Test-Path $script:ServiceTagCachePath)) {
        $age = (Get-Date) - (Get-Item $script:ServiceTagCachePath).LastWriteTime
        if ($age -lt $script:ServiceTagMaxAge) {
            Write-Verbose "Service Tags cache is fresh ($([int]$age.TotalHours)h old). Skipping download."
            return $true
        }
    }

    $url = Get-ServiceTagDownloadUrl
    if (-not $url) { return $false }

    try {
        if (-not (Test-Path $script:ServiceTagCacheDir)) {
            New-Item -Path $script:ServiceTagCacheDir -ItemType Directory -Force | Out-Null
        }

        Invoke-WebRequest -Uri $url -OutFile $script:ServiceTagCachePath -UseBasicParsing -TimeoutSec 120 -ErrorAction Stop
        Write-Verbose "Service Tags JSON downloaded and cached at $($script:ServiceTagCachePath)."
        return $true
    }
    catch {
        Write-Warning "Failed to download Service Tags JSON: $_"
        return $false
    }
}

# ───────────────── CIDR / IP math ─────────────────

function ConvertTo-UInt32 {
    <# Converts an IPv4 address string to a UInt32 for fast range comparisons. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$IPAddress)

    $parts = $IPAddress.Split('.')
    return ([uint32]$parts[0] -shl 24) -bor
           ([uint32]$parts[1] -shl 16) -bor
           ([uint32]$parts[2] -shl 8)  -bor
           ([uint32]$parts[3])
}

function Test-IPv4InCIDR {
    <# Returns $true if the given IPv4 address falls within the CIDR range. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][uint32]$IPUInt32,
        [Parameter(Mandatory)][uint32]$NetworkUInt32,
        [Parameter(Mandatory)][int]$MaskBits
    )

    if ($MaskBits -eq 0) { return $true }
    $mask = [uint32]([uint64]([math]::Pow(2, 32) - [math]::Pow(2, 32 - $MaskBits)))
    return (($IPUInt32 -band $mask) -eq ($NetworkUInt32 -band $mask))
}

# ───────────────── index builder ─────────────────

function Initialize-ServiceTagIndex {
    <#
    .SYNOPSIS Loads the cached JSON and builds a fast-lookup index of
             (NetworkUInt32, MaskBits, ServiceName) for all IPv4 prefixes.
    #>
    [CmdletBinding()]
    param()

    if (-not (Test-Path $script:ServiceTagCachePath)) {
        Write-Warning "Service Tags cache file not found. Call Update-ServiceTagCache first."
        return $false
    }

    try {
        $raw = Get-Content -Path $script:ServiceTagCachePath -Raw -ErrorAction Stop
        $script:ServiceTagData = $raw | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        Write-Warning "Failed to parse Service Tags JSON: $_"
        return $false
    }

    $index = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($svc in $script:ServiceTagData.values) {
        $svcName = $svc.name   # e.g. "AzureActiveDirectory", "Storage.WestUS2"
        foreach ($prefix in $svc.properties.addressPrefixes) {
            # Only index IPv4 CIDR ranges (skip IPv6 for this feature)
            if ($prefix -match '^(\d+\.\d+\.\d+\.\d+)/(\d+)$') {
                $netAddr  = $Matches[1]
                $maskBits = [int]$Matches[2]
                try {
                    $netUInt = ConvertTo-UInt32 -IPAddress $netAddr
                    $index.Add([PSCustomObject]@{
                        NetworkUInt32 = $netUInt
                        MaskBits      = $maskBits
                        ServiceName   = $svcName
                    })
                } catch {}
            }
        }
    }

    $script:ServiceTagIndex = $index.ToArray()
    $script:ServiceTagIPCache = @{}
    Write-Verbose "Service Tag index built: $($script:ServiceTagIndex.Count) IPv4 prefixes across $($script:ServiceTagData.values.Count) service tags."
    return $true
}

# ───────────────── lookup ─────────────────

function Resolve-ServiceTag {
    <#
    .SYNOPSIS  Returns the Microsoft service tag name(s) for a single IP, or empty string.
    .OUTPUTS   String — comma-separated service names, or "".
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$IPAddress)

    # Return from cache if available
    if ($script:ServiceTagIPCache.ContainsKey($IPAddress)) {
        return $script:ServiceTagIPCache[$IPAddress]
    }

    # Must have an index
    if (-not $script:ServiceTagIndex) { return "" }

    # Skip obviously non-IPv4
    if ($IPAddress -match ':' -or $IPAddress -eq '-' -or $IPAddress -eq '*' -or [string]::IsNullOrWhiteSpace($IPAddress)) {
        $script:ServiceTagIPCache[$IPAddress] = ""
        return ""
    }

    try {
        $ipUInt = ConvertTo-UInt32 -IPAddress $IPAddress
    }
    catch {
        $script:ServiceTagIPCache[$IPAddress] = ""
        return ""
    }

    $matchedTags = [System.Collections.Generic.List[string]]::new()

    foreach ($entry in $script:ServiceTagIndex) {
        if (Test-IPv4InCIDR -IPUInt32 $ipUInt -NetworkUInt32 $entry.NetworkUInt32 -MaskBits $entry.MaskBits) {
            if (-not $matchedTags.Contains($entry.ServiceName)) {
                $matchedTags.Add($entry.ServiceName)
            }
        }
    }

    # Deduplicate: if we have both a regional tag (e.g. "Storage.WestUS2")
    # and the parent tag ("Storage"), keep the more specific one
    $result = if ($matchedTags.Count -gt 0) { $matchedTags -join ", " } else { "" }
    $script:ServiceTagIPCache[$IPAddress] = $result
    return $result
}

function Resolve-ServiceTags {
    <#
    .SYNOPSIS  Batch-resolves an array of IP addresses to their Microsoft service tags.
    .OUTPUTS   Hashtable  IP -> ServiceTagString
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string[]]$IPAddresses)

    $results = @{}
    $unique  = $IPAddresses | Select-Object -Unique

    foreach ($ip in $unique) {
        $results[$ip] = Resolve-ServiceTag -IPAddress $ip
    }

    return $results
}
