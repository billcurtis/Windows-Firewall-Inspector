<#
.SYNOPSIS
    Azure Authentication Module for Firewall Log Agent.
.DESCRIPTION
    Provides Device Code Flow authentication for both ARM (management) and
    Storage scopes, plus Access Key fallback. Pure REST — no Az modules / CLI.
#>

# Well-known Azure PowerShell public client ID (supports device code flow)
$script:AzureADClientId  = "1950a258-227b-4e31-a9cf-717495945fc2"
$script:AzureADAuthority = "https://login.microsoftonline.com"
$script:ArmScope         = "https://management.azure.com/.default offline_access"
$script:StorageScope     = "https://storage.azure.com/.default"

# ────────────────── Auth Context ──────────────────

function New-AzureAuthContext {
    <#
    .SYNOPSIS Creates a blank authentication context.
    .DESCRIPTION StorageAccountName is optional at creation; set it later after the user picks one.
    #>
    [CmdletBinding()]
    param([string]$StorageAccountName = "")

    return [PSCustomObject]@{
        StorageAccountName = $StorageAccountName
        AuthType           = $null          # 'AccessKey' or 'AzureAD'
        AccessKey          = $null
        # ARM (management-plane) token
        ArmToken           = $null
        ArmTokenExpiry     = $null
        # Storage (data-plane) token
        BearerToken        = $null
        TokenExpiry        = $null
        # Shared
        RefreshToken       = $null
        TenantId           = "common"
    }
}

# ────────────────── Access Key ──────────────────

function Set-AccessKeyAuth {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][PSCustomObject]$AuthContext,
        [Parameter(Mandatory)][string]$AccessKey
    )
    $AuthContext.AuthType  = 'AccessKey'
    $AuthContext.AccessKey = $AccessKey
    return $AuthContext
}

# ────────────────── Device Code Flow (ARM scope) ──────────────────

function Request-DeviceCode {
    <#
    .SYNOPSIS  Initiates the device code flow (fast, non-blocking).
    .DESCRIPTION Returns the raw device-code response from Azure AD.
    #>
    [CmdletBinding()]
    param(
        [string]$TenantId = "common"
    )

    $deviceCodeUrl = "$script:AzureADAuthority/$TenantId/oauth2/v2.0/devicecode"
    $body = "client_id=$script:AzureADClientId&scope=$([uri]::EscapeDataString($script:ArmScope))"

    try {
        return Invoke-RestMethod -Uri $deviceCodeUrl -Method Post -Body $body `
            -ContentType "application/x-www-form-urlencoded" -ErrorAction Stop
    }
    catch { throw "Failed to initiate device code flow: $($_.Exception.Message)" }
}

function Poll-DeviceCodeToken {
    <#
    .SYNOPSIS  Blocks until the user completes sign-in or the code expires.
    .DESCRIPTION Designed to run inside a background runspace so the UI stays responsive.
                 Returns a hashtable with token fields on success, or throws.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$DeviceCode,
        [Parameter(Mandatory)][int]$ExpiresIn,
        [int]$Interval = 5
    )

    $clientId  = "1950a258-227b-4e31-a9cf-717495945fc2"
    $authority = "https://login.microsoftonline.com"
    $tokenUrl  = "$authority/$TenantId/oauth2/v2.0/token"
    $pollBody  = "grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Adevice_code&client_id=$clientId&device_code=$DeviceCode"

    $expiresAt = (Get-Date).AddSeconds($ExpiresIn)
    $sleepSec  = [Math]::Max($Interval, 5)

    while ((Get-Date) -lt $expiresAt) {
        Start-Sleep -Seconds $sleepSec
        try {
            $r = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $pollBody `
                -ContentType "application/x-www-form-urlencoded" -ErrorAction Stop

            return @{
                access_token  = $r.access_token
                refresh_token = $r.refresh_token
                expires_in    = $r.expires_in
            }
        }
        catch {
            $errBody = $null
            try { $errBody = $_.ErrorDetails.Message | ConvertFrom-Json } catch {}
            if ($errBody.error -eq "authorization_pending") { continue }
            elseif ($errBody.error -eq "slow_down")         { $sleepSec += 5; continue }
            else { throw "Device code auth failed: $($errBody.error_description)" }
        }
    }
    throw "Device code authentication timed out."
}

function Connect-AzureAD {
    <#
    .SYNOPSIS  Blocking convenience wrapper (for console / headless mode).
    .DESCRIPTION Calls Request-DeviceCode then Poll-DeviceCodeToken synchronously.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][PSCustomObject]$AuthContext,
        [string]$TenantId = "common",
        [scriptblock]$OnUserCodeReceived
    )

    $AuthContext.TenantId = $TenantId
    $dc = Request-DeviceCode -TenantId $TenantId

    if ($OnUserCodeReceived) {
        & $OnUserCodeReceived $dc.user_code $dc.verification_uri $dc.message
    }
    else { Write-Host $dc.message -ForegroundColor Yellow }

    $tokenResult = Poll-DeviceCodeToken -TenantId $TenantId `
        -DeviceCode $dc.device_code -ExpiresIn $dc.expires_in -Interval $dc.interval

    $AuthContext.AuthType       = 'AzureAD'
    $AuthContext.ArmToken       = $tokenResult.access_token
    $AuthContext.RefreshToken   = $tokenResult.refresh_token
    $AuthContext.ArmTokenExpiry = (Get-Date).AddSeconds($tokenResult.expires_in - 120)
    return $AuthContext
}

# ────────────────── Acquire Storage Token (from refresh token) ──────────────────

function Get-StorageToken {
    <#
    .SYNOPSIS Exchanges the refresh token for a Storage data-plane bearer token.
    .DESCRIPTION Call this after Connect-AzureAD and after the user has selected a storage account.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][PSCustomObject]$AuthContext)

    if (-not $AuthContext.RefreshToken) { throw "No refresh token. Call Connect-AzureAD first." }

    $tokenUrl = "$script:AzureADAuthority/$($AuthContext.TenantId)/oauth2/v2.0/token"
    $body = "grant_type=refresh_token&client_id=$script:AzureADClientId&refresh_token=$($AuthContext.RefreshToken)&scope=$([uri]::EscapeDataString($script:StorageScope))"

    try {
        $r = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $body `
            -ContentType "application/x-www-form-urlencoded" -ErrorAction Stop

        $AuthContext.AuthType     = 'AzureAD'
        $AuthContext.BearerToken  = $r.access_token
        $AuthContext.RefreshToken = $r.refresh_token   # may be rotated
        $AuthContext.TokenExpiry  = (Get-Date).AddSeconds($r.expires_in - 120)
        return $AuthContext
    }
    catch { throw "Failed to acquire storage token: $($_.Exception.Message)" }
}

# ────────────────── Token Refresh Helpers ──────────────────

function Update-AzureADToken {
    <# Refreshes the storage data-plane token. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][PSCustomObject]$AuthContext)

    if ($AuthContext.AuthType -ne 'AzureAD' -or -not $AuthContext.RefreshToken) {
        throw "No refresh token available."
    }

    $tokenUrl = "$script:AzureADAuthority/$($AuthContext.TenantId)/oauth2/v2.0/token"
    $body = "grant_type=refresh_token&client_id=$script:AzureADClientId&refresh_token=$($AuthContext.RefreshToken)&scope=$([uri]::EscapeDataString($script:StorageScope))"

    try {
        $r = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $body `
            -ContentType "application/x-www-form-urlencoded" -ErrorAction Stop
        $AuthContext.BearerToken  = $r.access_token
        $AuthContext.RefreshToken = $r.refresh_token
        $AuthContext.TokenExpiry  = (Get-Date).AddSeconds($r.expires_in - 120)
        return $AuthContext
    }
    catch { throw "Storage token refresh failed: $($_.Exception.Message)" }
}

# ────────────────── Auth Headers for Table Storage ──────────────────

function Get-AzureAuthHeaders {
    <#
    .SYNOPSIS Returns HTTP headers for an Azure Table Storage REST call.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][PSCustomObject]$AuthContext,
        [Parameter(Mandatory)][string]$ResourcePath,
        [string]$Method      = "GET",
        [string]$ContentType = "application/json"
    )

    $dateString = [DateTime]::UtcNow.ToString("R")

    $headers = @{
        "x-ms-date"          = $dateString
        "x-ms-version"       = "2021-12-02"
        "Accept"             = "application/json;odata=nometadata"
        "DataServiceVersion" = "3.0;NetFx"
    }

    if ($AuthContext.AuthType -eq 'AccessKey') {
        $stringToSign = "$dateString`n/$($AuthContext.StorageAccountName)/$ResourcePath"
        $keyBytes     = [Convert]::FromBase64String($AuthContext.AccessKey)
        $hmac         = [System.Security.Cryptography.HMACSHA256]::new($keyBytes)
        $sigBytes     = $hmac.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($stringToSign))
        $signature    = [Convert]::ToBase64String($sigBytes)
        $hmac.Dispose()
        $headers["Authorization"] = "SharedKeyLite $($AuthContext.StorageAccountName):$signature"
    }
    elseif ($AuthContext.AuthType -eq 'AzureAD') {
        # Auto-refresh storage token
        if ($AuthContext.TokenExpiry -and (Get-Date) -ge $AuthContext.TokenExpiry) {
            try { $AuthContext = Update-AzureADToken -AuthContext $AuthContext } catch { Write-Warning $_ }
        }
        $headers["Authorization"] = "Bearer $($AuthContext.BearerToken)"
    }
    else {
        throw "Authentication not configured."
    }

    return $headers
}
