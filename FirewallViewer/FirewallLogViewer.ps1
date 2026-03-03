<#
.SYNOPSIS
    Windows Firewall Viewer - Main Entry Point.
.DESCRIPTION
    Dark-mode WPF dashboard for viewing, filtering, and visualising
    Windows Firewall log data stored in Azure Table Storage by the
    Firewall Log Agent. Includes bar/pie charts, IP owner lookup, and CSV export.
.EXAMPLE
    .\FirewallLogViewer.ps1
    # Launches the viewer GUI.
#>
#Requires -Version 5.1

[CmdletBinding()]
param()

# ------------------------------------------------------------------
# Load modules
# ------------------------------------------------------------------
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition
$modulePath = Join-Path $scriptRoot "Modules"

. (Join-Path $modulePath "AzureAuth.ps1")
. (Join-Path $modulePath "AzureTableStorage.ps1")
. (Join-Path $modulePath "IPOwnerLookup.ps1")
. (Join-Path $modulePath "ServiceTagLookup.ps1")
. (Join-Path $modulePath "ViewerGUI.ps1")

# ------------------------------------------------------------------
# Launch viewer
# ------------------------------------------------------------------
Show-ViewerWindow
