<#
.SYNOPSIS
    Azure Resource Manager REST API Module.
.DESCRIPTION
    Lists subscriptions, storage accounts, retrieves storage account keys,
    and lists tables — all via pure ARM REST calls (no Az modules / CLI).
#>

$script:ArmBaseUri = "https://management.azure.com"

function Get-ArmAuthHeaders {
    <# Returns Authorization header for ARM calls using the ARM bearer token. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][PSCustomObject]$AuthContext)

    if (-not $AuthContext.ArmToken) {
        throw "No ARM token. Call Connect-AzureAD first."
    }

    # Refresh if near expiry
    if ($AuthContext.ArmTokenExpiry -and (Get-Date) -ge $AuthContext.ArmTokenExpiry -and $AuthContext.RefreshToken) {
        $AuthContext = Update-ArmToken -AuthContext $AuthContext
    }

    return @{ "Authorization" = "Bearer $($AuthContext.ArmToken)" }
}

function Update-ArmToken {
    [CmdletBinding()]
    param([Parameter(Mandatory)][PSCustomObject]$AuthContext)

    $tokenUrl = "$($script:AzureADAuthority)/$($AuthContext.TenantId)/oauth2/v2.0/token"
    $armScope = "https://management.azure.com/.default"
    $body = "grant_type=refresh_token&client_id=$script:AzureADClientId&refresh_token=$($AuthContext.RefreshToken)&scope=$([uri]::EscapeDataString($armScope))"

    try {
        $r = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $body `
            -ContentType "application/x-www-form-urlencoded" -ErrorAction Stop
        $AuthContext.ArmToken       = $r.access_token
        $AuthContext.RefreshToken   = $r.refresh_token
        $AuthContext.ArmTokenExpiry = (Get-Date).AddSeconds($r.expires_in - 120)
        return $AuthContext
    }
    catch { throw "ARM token refresh failed: $($_.Exception.Message)" }
}

# ────────────────── Subscriptions ──────────────────

function Get-AzureSubscriptions {
    <#
    .SYNOPSIS Lists all Azure subscriptions accessible to the authenticated user.
    .OUTPUTS Array of objects with subscriptionId, displayName, state.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][PSCustomObject]$AuthContext)

    $headers = Get-ArmAuthHeaders -AuthContext $AuthContext
    $uri     = "$script:ArmBaseUri/subscriptions?api-version=2022-12-01"

    $all = [System.Collections.Generic.List[object]]::new()

    do {
        $resp = Invoke-RestMethod -Uri $uri -Method Get -Headers $headers -ErrorAction Stop
        if ($resp.value) { foreach ($s in $resp.value) { $all.Add($s) } }
        $uri = $resp.nextLink
    } while ($uri)

    return $all.ToArray()
}

# ────────────────── Storage Accounts ──────────────────

function Get-AzureStorageAccounts {
    <#
    .SYNOPSIS Lists all storage accounts in a subscription.
    .OUTPUTS Array of storage account objects (id, name, location, resourceGroup parsed from id).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][PSCustomObject]$AuthContext,
        [Parameter(Mandatory)][string]$SubscriptionId
    )

    $headers = Get-ArmAuthHeaders -AuthContext $AuthContext
    $uri     = "$script:ArmBaseUri/subscriptions/$SubscriptionId/providers/Microsoft.Storage/storageAccounts?api-version=2023-05-01"

    $all = [System.Collections.Generic.List[object]]::new()

    do {
        $resp = Invoke-RestMethod -Uri $uri -Method Get -Headers $headers -ErrorAction Stop
        if ($resp.value) { foreach ($sa in $resp.value) { $all.Add($sa) } }
        $uri = $resp.nextLink
    } while ($uri)

    # Attach resourceGroup parsed from id for convenience
    foreach ($sa in $all) {
        if ($sa.id -match '/resourceGroups/([^/]+)/') {
            $sa | Add-Member -NotePropertyName "resourceGroup" -NotePropertyValue $Matches[1] -Force
        }
    }

    # Filter to accounts that support Table Storage (have a table endpoint)
    $tableCapable = $all | Where-Object {
        $_.properties.primaryEndpoints.table -and
        $_.kind -notin @('BlobStorage','FileStorage','BlockBlobStorage')
    }

    return @($tableCapable)
}

# ────────────────── Storage Account Keys ──────────────────

function Get-AzureStorageAccountKeys {
    <#
    .SYNOPSIS Retrieves the access keys for a storage account via ARM.
    .OUTPUTS Array of key objects with keyName and value.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][PSCustomObject]$AuthContext,
        [Parameter(Mandatory)][string]$SubscriptionId,
        [Parameter(Mandatory)][string]$ResourceGroup,
        [Parameter(Mandatory)][string]$StorageAccountName
    )

    $headers = Get-ArmAuthHeaders -AuthContext $AuthContext
    $uri     = "$script:ArmBaseUri/subscriptions/$SubscriptionId/resourceGroups/$ResourceGroup/providers/Microsoft.Storage/storageAccounts/$StorageAccountName/listKeys?api-version=2023-05-01"

    $resp = Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body "" `
        -ContentType "application/json" -ErrorAction Stop

    return $resp.keys
}

# ────────────────── List Tables in a Storage Account ──────────────────

function Get-AzureStorageTables {
    <#
    .SYNOPSIS Lists all tables in a storage account using the Table Storage REST API.
    .DESCRIPTION Uses the storage access key (SharedKeyLite) for authentication.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][PSCustomObject]$AuthContext)

    $resource = "Tables"
    $uri      = "https://$($AuthContext.StorageAccountName).table.core.windows.net/$resource"
    $headers  = Get-AzureAuthHeaders -AuthContext $AuthContext -ResourcePath $resource -Method "GET"

    try {
        $resp = Invoke-RestMethod -Uri $uri -Method Get -Headers $headers -ErrorAction Stop
        if ($resp.value) {
            return ($resp.value | ForEach-Object { $_.TableName })
        }
        return @()
    }
    catch {
        # If no tables exist yet, some responses may be empty
        $code = 0
        try { $code = $_.Exception.Response.StatusCode.value__ } catch {}
        if ($code -eq 404) { return @() }
        throw "Failed to list tables: $($_.Exception.Message)"
    }
}
