<#
.SYNOPSIS
    Azure Authentication Module for Windows Firewall Viewer (Az PowerShell Cmdlets).
.DESCRIPTION
    Uses the Az PowerShell module (Connect-AzAccount, Get-AzSubscription,
    Get-AzStorageAccount, etc.) for authentication and resource browsing.
    Provides SharedKeyLite / Bearer header generation for Azure Table Storage REST calls.
#>

# ─────────────────────────────────────────────────
# Connection helpers (Az module wrappers)
# ─────────────────────────────────────────────────

function Connect-AzureViewer {
    <#
    .SYNOPSIS Opens an interactive browser login via Connect-AzAccount.
    .OUTPUTS Boolean $true on success, throws on failure.
    #>
    [CmdletBinding()]
    param()

    try {
        $ctx = Connect-AzAccount -ErrorAction Stop
        if (-not $ctx) { throw "Connect-AzAccount returned no context." }
        return $true
    }
    catch {
        throw "Azure login failed: $($_.Exception.Message)"
    }
}

function Get-ViewerSubscriptions {
    <#
    .SYNOPSIS Returns all enabled Azure subscriptions for the logged-in user.
    #>
    [CmdletBinding()]
    param()

    try {
        $subs = Get-AzSubscription -ErrorAction Stop |
            Where-Object { $_.State -eq 'Enabled' } |
            Sort-Object Name
        return $subs
    }
    catch {
        throw "Failed to list subscriptions: $($_.Exception.Message)"
    }
}

function Get-ViewerStorageAccounts {
    <#
    .SYNOPSIS Lists table-capable storage accounts in a subscription.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SubscriptionId
    )

    try {
        $null = Set-AzContext -SubscriptionId $SubscriptionId -ErrorAction Stop

        $accounts = Get-AzStorageAccount -ErrorAction Stop |
            Where-Object {
                $_.Kind -notin @('BlobStorage','FileStorage','BlockBlobStorage') -and
                $_.PrimaryEndpoints.Table
            } |
            Sort-Object StorageAccountName

        return $accounts
    }
    catch {
        throw "Failed to list storage accounts: $($_.Exception.Message)"
    }
}

function Get-ViewerStorageAccountKeys {
    <#
    .SYNOPSIS Gets the access keys for a storage account.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ResourceGroupName,
        [Parameter(Mandatory)][string]$StorageAccountName
    )

    try {
        $keys = Get-AzStorageAccountKey -ResourceGroupName $ResourceGroupName `
                    -Name $StorageAccountName -ErrorAction Stop
        return $keys
    }
    catch {
        throw "Failed to get storage keys for '$StorageAccountName': $($_.Exception.Message)"
    }
}

function Get-ViewerStorageTables {
    <#
    .SYNOPSIS Lists table names in a storage account using REST API.
    .DESCRIPTION
        Tries the current auth method first. If SharedKey returns 403,
        falls back to Bearer token via Get-AzAccessToken.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][PSCustomObject]$AuthContext
    )

    $resource = "Tables"
    $uri      = "https://$($AuthContext.StorageAccountName).table.core.windows.net/$resource"
    $headers  = Get-AzureAuthHeaders -AuthContext $AuthContext -ResourcePath $resource -Method "GET"

    try {
        $resp = Invoke-RestMethod -Uri $uri -Method Get -Headers $headers -ErrorAction Stop
        return ($resp.value | ForEach-Object { $_.TableName })
    }
    catch {
        $code = 0
        try { $code = $_.Exception.Response.StatusCode.value__ } catch {}

        if ($code -eq 403 -and $AuthContext.AuthType -eq 'AccessKey') {
            Write-Warning "Shared key access denied. Falling back to Azure AD Bearer token."
            try {
                $AuthContext = Set-AzureADAuth -AuthContext $AuthContext
                $headers = Get-AzureAuthHeaders -AuthContext $AuthContext -ResourcePath $resource -Method "GET"
                $resp = Invoke-RestMethod -Uri $uri -Method Get -Headers $headers -ErrorAction Stop
                return ($resp.value | ForEach-Object { $_.TableName })
            }
            catch {
                throw "Failed to list tables (RBAC fallback): $($_.Exception.Message)"
            }
        }
        throw "Failed to list tables: $($_.Exception.Message)"
    }
}

# ─────────────────────────────────────────────────
# Auth context helpers (used by AzureTableStorage.ps1)
# ─────────────────────────────────────────────────

function New-AzureAuthContext {
    <#
    .SYNOPSIS Creates a new auth context object for Table Storage REST calls.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$StorageAccountName
    )
    return [PSCustomObject]@{
        StorageAccountName = $StorageAccountName
        AuthType           = $null        # 'AccessKey' or 'AzureAD'
        AccessKey          = $null
        BearerToken        = $null
        TokenExpiry        = $null
    }
}

function Set-AccessKeyAuth {
    <#
    .SYNOPSIS Configures the auth context with a shared access key.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][PSCustomObject]$AuthContext,
        [Parameter(Mandatory)][string]$AccessKey
    )
    $AuthContext.AuthType  = 'AccessKey'
    $AuthContext.AccessKey = $AccessKey
    return $AuthContext
}

function Get-PlainAccessToken {
    <#
    .SYNOPSIS Gets a plain-text storage token from the Az module.
    .DESCRIPTION Handles both Az.Accounts 2.x (plain string) and 3.x+ (SecureString).
    .OUTPUTS PSCustomObject with .Token (string) and .ExpiresOn (DateTime).
    #>
    [CmdletBinding()]
    param()

    $result = Get-AzAccessToken -ResourceUrl "https://storage.azure.com" -ErrorAction Stop

    # Az.Accounts 3.x returns Token as SecureString
    $tokenStr = if ($result.Token -is [System.Security.SecureString]) {
        [System.Net.NetworkCredential]::new('', $result.Token).Password
    } else {
        $result.Token
    }

    $expiry = if ($result.ExpiresOn -is [DateTimeOffset]) {
        $result.ExpiresOn.UtcDateTime
    } else {
        $result.ExpiresOn
    }

    return [PSCustomObject]@{
        Token     = $tokenStr
        ExpiresOn = $expiry
    }
}

function Set-AzureADAuth {
    <#
    .SYNOPSIS Configures the auth context with a Bearer token from the Az module.
    .DESCRIPTION Uses Get-AzAccessToken to obtain a storage-scoped token.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][PSCustomObject]$AuthContext
    )

    try {
        $token = Get-PlainAccessToken
        $AuthContext.AuthType    = 'AzureAD'
        $AuthContext.BearerToken = $token.Token
        $AuthContext.TokenExpiry = $token.ExpiresOn.AddMinutes(-2)
        return $AuthContext
    }
    catch {
        throw "Failed to get storage token: $($_.Exception.Message)"
    }
}

# ─────────────────────────────────────────────────
# REST header generation for Table Storage
# ─────────────────────────────────────────────────

function Get-AzureAuthHeaders {
    <#
    .SYNOPSIS Generates authentication headers for Azure Table Storage REST calls.
    .DESCRIPTION
        Returns SharedKeyLite headers when AuthType is AccessKey,
        or Bearer token headers when AuthType is AzureAD.
        Automatically refreshes expired Bearer tokens via Get-AzAccessToken.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][PSCustomObject]$AuthContext,
        [Parameter(Mandatory)][string]$ResourcePath,
        [string]$Method = "GET",
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
        # Auto-refresh expired tokens via Az module
        if ($AuthContext.TokenExpiry -and (Get-Date) -ge $AuthContext.TokenExpiry) {
            try {
                $token = Get-PlainAccessToken
                $AuthContext.BearerToken = $token.Token
                $AuthContext.TokenExpiry = $token.ExpiresOn.AddMinutes(-2)
            }
            catch {
                Write-Warning "Token refresh failed: $_"
            }
        }
        $headers["Authorization"] = "Bearer $($AuthContext.BearerToken)"
    }
    else {
        throw "Authentication not configured."
    }

    return $headers
}
