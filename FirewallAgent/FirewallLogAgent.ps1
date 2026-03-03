<#
.SYNOPSIS
    Firewall Log Agent - Main Entry Point.
.DESCRIPTION
    Collects Windows Firewall log entries, resolves processes, and uploads
    to Azure Table Storage via REST API. Supports GUI and headless modes.
.PARAMETER NoGUI
    Run in headless / console mode (no WPF window).
.PARAMETER StorageAccount
    Azure Storage Account name (required for NoGUI mode).
.PARAMETER AccessKey
    Azure Storage Account access key. If omitted in NoGUI mode, Azure AD
    device code flow is used.
.PARAMETER TableName
    Azure Table name to store logs in. Default: FirewallLogs.
.PARAMETER IntervalSeconds
    Collection & upload interval in seconds. Default: 60.
.EXAMPLE
    .\FirewallLogAgent.ps1
    # Launches the GUI
.EXAMPLE
    .\FirewallLogAgent.ps1 -NoGUI -StorageAccount "mystorageacct" -AccessKey "base64key=="
    # Headless mode with access key auth
.EXAMPLE
    .\FirewallLogAgent.ps1 -NoGUI -StorageAccount "mystorageacct"
    # Headless mode with Azure AD device code auth
#>
#Requires -Version 5.1

[CmdletBinding()]
param(
    [switch]$NoGUI,
    [string]$StorageAccount,
    [string]$AccessKey,
    [string]$TableName       = "FirewallLogs",
    [int]$IntervalSeconds    = 60,
    [string]$TenantId        = "common"
)

# ------------------------------------------------------------------
# Ensure running as Administrator (netsh + log file access require it)
# ------------------------------------------------------------------
$currentPrincipal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Re-launch elevated, preserving all original arguments
    $scriptPath = $MyInvocation.MyCommand.Definition
    $argList = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
    foreach ($key in $PSBoundParameters.Keys) {
        $val = $PSBoundParameters[$key]
        if ($val -is [switch]) { if ($val) { $argList += " -$key" } }
        else { $argList += " -$key `"$val`"" }
    }
    try {
        Start-Process -FilePath "powershell.exe" -ArgumentList $argList -Verb RunAs
    }
    catch {
        Write-Error "This script requires Administrator privileges. Right-click and Run as Administrator."
    }
    exit
}

# ------------------------------------------------------------------
# Load modules
# ------------------------------------------------------------------
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition
$modulePath = Join-Path $scriptRoot "Modules"

. (Join-Path $modulePath "AzureAuth.ps1")
. (Join-Path $modulePath "AzureTableStorage.ps1")
. (Join-Path $modulePath "AzureResourceManager.ps1")
. (Join-Path $modulePath "FirewallLogManager.ps1")

# ------------------------------------------------------------------
# GUI Mode
# ------------------------------------------------------------------
if (-not $NoGUI) {
    . (Join-Path $modulePath "AgentGUI.ps1")
    Show-AgentWindow
    return
}

# ------------------------------------------------------------------
# Headless / Console Mode
# ------------------------------------------------------------------
if (-not $StorageAccount) {
    Write-Error "StorageAccount is required in NoGUI mode. Use -StorageAccount 'name'."
    exit 1
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Firewall Log Agent (Console Mode)"     -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# --- Authenticate ---
$authCtx = New-AzureAuthContext -StorageAccountName $StorageAccount

if ($AccessKey) {
    $authCtx = Set-AccessKeyAuth -AuthContext $authCtx -AccessKey $AccessKey
    Write-Host "[AUTH] Using Access Key authentication." -ForegroundColor Green
}
else {
    Write-Host "[AUTH] Starting Azure AD Device Code Flow..." -ForegroundColor Yellow
    $authCtx = Start-DeviceCodeAuth -AuthContext $authCtx -TenantId $TenantId
    Write-Host "[AUTH] Azure AD authentication successful." -ForegroundColor Green
}

# --- Create table ---
Write-Host "[INIT] Ensuring table '$TableName' exists..." -ForegroundColor Gray
$null = New-StorageTable -AuthContext $authCtx -TableName $TableName
Write-Host "[INIT] Table ready." -ForegroundColor Green

# --- Enable firewall logging ---
Write-Host "[INIT] Enabling Windows Firewall logging..." -ForegroundColor Gray
try {
    Enable-FirewallLogging
    Write-Host "[INIT] Firewall logging enabled." -ForegroundColor Green
}
catch {
    Write-Warning "Could not enable firewall logging (run as Administrator): $_"
}

# --- Build firewall rule name cache ---
Write-Host "[INIT] Building firewall rule name cache..." -ForegroundColor Gray
try { Update-FirewallRuleCache } catch { Write-Warning "Could not build rule cache: $_" }

# --- Initial file position ---
$null = Get-NewFirewallLogEntries -ResetPosition
Write-Host "[INIT] Log file position initialized." -ForegroundColor Gray
Write-Host ""
Write-Host "Collecting every $IntervalSeconds seconds. Press Ctrl+C to stop." -ForegroundColor Cyan
Write-Host ""

# --- Graceful shutdown via Ctrl+C ---
$keepRunning = $true

$null = Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action {
    $keepRunning = $false
}

try {
    [Console]::TreatControlCAsInput = $false
}
catch {}

trap {
    Write-Host "`n[EXIT] Restoring firewall logging settings..." -ForegroundColor Yellow
    try { Restore-FirewallLogging } catch { Write-Warning $_ }
    Write-Host "[EXIT] Done." -ForegroundColor Green
    break
}

# --- Main collection loop ---
$totalEntries = 0
$totalUploads = 0

while ($keepRunning) {
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    try {
        $entries = Get-NewFirewallLogEntries
        $count   = $entries.Count

        if ($count -gt 0) {
            $totalEntries += $count
            Write-Host "[$ts] Found $count new entries. Uploading..." -ForegroundColor White

            $result = Add-StorageTableEntities -AuthContext $authCtx -TableName $TableName -Entities $entries
            $totalUploads += $result.Success

            if ($result.Errors -gt 0) {
                Write-Host "[$ts] Uploaded $($result.Success), failed $($result.Errors). Total: $totalUploads" -ForegroundColor Yellow
            }
            else {
                Write-Host "[$ts] Uploaded $($result.Success) entities. Total: $totalUploads" -ForegroundColor Green
            }
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

# Restore on normal exit
Write-Host "[EXIT] Restoring firewall logging settings..." -ForegroundColor Yellow
try { Restore-FirewallLogging } catch { Write-Warning $_ }
Write-Host "[EXIT] Agent stopped. Total entries: $totalEntries, Total uploaded: $totalUploads" -ForegroundColor Green
