<#
.SYNOPSIS
    Azure Table Storage REST API Module for Firewall Log Agent.
.DESCRIPTION
    Provides CRUD operations for Azure Table Storage using pure REST API calls.
    Supports SharedKeyLite and Azure AD Bearer token authentication.
#>

function New-StorageTable {
    <#
    .SYNOPSIS Creates a new table in the storage account. Returns $true if created or already exists.
    #>
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
        if ($code -eq 409) { return $true }   # TableAlreadyExists
        throw "Failed to create table '$TableName': $($_.Exception.Message)"
    }
}

function Add-StorageTableEntity {
    <#
    .SYNOPSIS Inserts a single entity into the specified table.
    .PARAMETER Entity
        Hashtable with PartitionKey, RowKey, and custom properties.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][PSCustomObject]$AuthContext,
        [Parameter(Mandatory)][string]$TableName,
        [Parameter(Mandatory)][hashtable]$Entity
    )

    $resource = $TableName
    $uri      = "https://$($AuthContext.StorageAccountName).table.core.windows.net/$resource"
    $headers  = Get-AzureAuthHeaders -AuthContext $AuthContext -ResourcePath $resource -Method "POST"
    $body     = $Entity | ConvertTo-Json -Depth 5

    try {
        $null = Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $body `
            -ContentType "application/json" -ErrorAction Stop
        return $true
    }
    catch {
        $code = 0
        try { $code = $_.Exception.Response.StatusCode.value__ } catch {}
        if ($code -eq 409) { return $true }   # EntityAlreadyExists
        throw "Failed to insert entity: $($_.Exception.Message)"
    }
}

function Add-StorageTableEntities {
    <#
    .SYNOPSIS Inserts multiple entities sequentially. Returns count of successful inserts.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][PSCustomObject]$AuthContext,
        [Parameter(Mandatory)][string]$TableName,
        [Parameter(Mandatory)][hashtable[]]$Entities
    )

    $successCount = 0
    $errorCount   = 0

    foreach ($entity in $Entities) {
        try {
            $null = Add-StorageTableEntity -AuthContext $AuthContext -TableName $TableName -Entity $entity
            $successCount++
        }
        catch {
            $errorCount++
            Write-Warning "Entity insert failed: $_"
        }
    }
    return @{ Success = $successCount; Errors = $errorCount }
}

function Get-StorageTableEntities {
    <#
    .SYNOPSIS Queries entities from a table with optional OData filter. Handles pagination.
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
            throw "Table query failed: $($_.Exception.Message)"
        }

        $content = $resp.Content | ConvertFrom-Json
        if ($content.value) {
            foreach ($e in $content.value) { $allEntities.Add($e) }
        }

        # Continuation tokens
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
        else {
            $uri = $null
        }
    } while ($uri)

    return $allEntities.ToArray()
}

function Remove-StorageTableEntity {
    <#
    .SYNOPSIS Deletes a single entity by PartitionKey and RowKey.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][PSCustomObject]$AuthContext,
        [Parameter(Mandatory)][string]$TableName,
        [Parameter(Mandatory)][string]$PartitionKey,
        [Parameter(Mandatory)][string]$RowKey
    )

    $resource = "$TableName(PartitionKey='$PartitionKey',RowKey='$RowKey')"
    $uri      = "https://$($AuthContext.StorageAccountName).table.core.windows.net/$resource"
    $headers  = Get-AzureAuthHeaders -AuthContext $AuthContext -ResourcePath $resource -Method "DELETE"
    $headers["If-Match"] = "*"

    try {
        $null = Invoke-RestMethod -Uri $uri -Method Delete -Headers $headers -ErrorAction Stop
        return $true
    }
    catch { throw "Failed to delete entity: $($_.Exception.Message)" }
}
