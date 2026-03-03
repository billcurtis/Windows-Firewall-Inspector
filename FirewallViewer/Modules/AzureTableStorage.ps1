<#
.SYNOPSIS
    Azure Table Storage REST API Module for Windows Firewall Viewer.
.DESCRIPTION
    Provides query and CRUD operations for Azure Table Storage using pure REST API calls.
#>

function New-StorageTable {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][PSCustomObject]$AuthContext,
        [Parameter(Mandatory)][string]$TableName
    )
    $resource = "Tables"
    $uri      = "https://$($AuthContext.StorageAccountName).table.core.windows.net/$resource"
    $headers  = Get-AzureAuthHeaders -AuthContext $AuthContext -ResourcePath $resource -Method "POST"
    $body     = @{ TableName = $TableName } | ConvertTo-Json

    try {
        $null = Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $body `
            -ContentType "application/json" -ErrorAction Stop
        return $true
    }
    catch {
        $code = 0
        try { $code = $_.Exception.Response.StatusCode.value__ } catch {}
        if ($code -eq 409) { return $true }

        # 401/403 → fall back / refresh RBAC bearer token
        if ($code -in @(401,403)) {
            Write-Warning "Auth denied on table create ($code). Switching/refreshing Azure AD token."
            try {
                $null = Set-AzureADAuth -AuthContext $AuthContext
                $headers = Get-AzureAuthHeaders -AuthContext $AuthContext -ResourcePath $resource -Method "POST"
                $null = Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $body `
                    -ContentType "application/json" -ErrorAction Stop
                return $true
            }
            catch {
                $c2 = 0; try { $c2 = $_.Exception.Response.StatusCode.value__ } catch {}
                if ($c2 -eq 409) { return $true }
                throw "Failed to create table (RBAC fallback): $($_.Exception.Message)"
            }
        }
        throw "Failed to create table '$TableName': $($_.Exception.Message)"
    }
}

function Get-StorageTableEntities {
    <#
    .SYNOPSIS Queries entities from a table with optional OData filter. Handles continuation/pagination.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][PSCustomObject]$AuthContext,
        [Parameter(Mandatory)][string]$TableName,
        [string]$Filter  = "",
        [int]$Top        = 0,
        [string]$Select  = ""
    )

    $baseResource = "$TableName()"
    $baseUri      = "https://$($AuthContext.StorageAccountName).table.core.windows.net/$baseResource"

    $qp = @()
    if ($Filter) { $qp += "`$filter=$([uri]::EscapeDataString($Filter))" }
    if ($Top -gt 0) { $qp += "`$top=$Top" }
    if ($Select) { $qp += "`$select=$([uri]::EscapeDataString($Select))" }

    $uri = $baseUri
    if ($qp.Count -gt 0) { $uri += "?" + ($qp -join "&") }

    $allEntities = [System.Collections.Generic.List[object]]::new()

    do {
        $headers = Get-AzureAuthHeaders -AuthContext $AuthContext -ResourcePath $baseResource -Method "GET"

        try {
            $resp = Invoke-WebRequest -Uri $uri -Method Get -Headers $headers -UseBasicParsing -ErrorAction Stop
        }
        catch {
            $code = 0
            try { $code = $_.Exception.Response.StatusCode.value__ } catch {}

            if ($code -in @(401,403)) {
                if ($AuthContext.AuthType -eq 'AccessKey') {
                    # Shared key blocked → switch to RBAC bearer
                    Write-Warning "Shared key denied during query. Switching to Azure AD."
                    $null = Set-AzureADAuth -AuthContext $AuthContext
                }
                else {
                    # AzureAD token may be expired → get fresh token
                    Write-Warning "Bearer token rejected ($code). Refreshing token."
                    $null = Set-AzureADAuth -AuthContext $AuthContext
                }
                $headers = Get-AzureAuthHeaders -AuthContext $AuthContext -ResourcePath $baseResource -Method "GET"
                try {
                    $resp = Invoke-WebRequest -Uri $uri -Method Get -Headers $headers -UseBasicParsing -ErrorAction Stop
                }
                catch { throw "Table query failed (auth retry): $($_.Exception.Message)" }
            }
            else { throw "Table query failed: $($_.Exception.Message)" }
        }

        $content = $resp.Content | ConvertFrom-Json
        if ($content.value) {
            foreach ($e in $content.value) { $allEntities.Add($e) }
        }

        $npk = $null; $nrk = $null
        try { $npk = $resp.Headers["x-ms-continuation-NextPartitionKey"] } catch {}
        try { $nrk = $resp.Headers["x-ms-continuation-NextRowKey"] }       catch {}

        if ($npk) {
            $contParams = @("NextPartitionKey=$([uri]::EscapeDataString($npk))")
            if ($nrk) { $contParams += "NextRowKey=$([uri]::EscapeDataString($nrk))" }
            if ($Filter) { $contParams += "`$filter=$([uri]::EscapeDataString($Filter))" }
            if ($Top -gt 0) { $contParams += "`$top=$Top" }
            if ($Select) { $contParams += "`$select=$([uri]::EscapeDataString($Select))" }
            $uri = $baseUri + "?" + ($contParams -join "&")
        }
        else { $uri = $null }
    } while ($uri)

    return $allEntities.ToArray()
}

function Get-StorageTableNames {
    <#
    .SYNOPSIS Lists all table names in the storage account.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][PSCustomObject]$AuthContext)

    $resource = "Tables"
    $uri      = "https://$($AuthContext.StorageAccountName).table.core.windows.net/$resource"
    $headers  = Get-AzureAuthHeaders -AuthContext $AuthContext -ResourcePath $resource -Method "GET"

    try {
        $resp = Invoke-RestMethod -Uri $uri -Method Get -Headers $headers -ErrorAction Stop
        return ($resp.value | ForEach-Object { $_.TableName })
    }
    catch { throw "Failed to list tables: $($_.Exception.Message)" }
}
