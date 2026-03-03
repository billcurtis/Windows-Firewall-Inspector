<#
.SYNOPSIS
    WPF GUI for the Firewall Log Agent.
.DESCRIPTION
    Provides a graphical interface with "Connect to Azure" device-code login,
    then cascading dropdowns for Subscription, Storage Account, and Table.
#>

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

function Show-AgentWindow {
    [CmdletBinding()]
    param(
        [scriptblock]$OnStart,
        [scriptblock]$OnStop,
        [scriptblock]$OnExit
    )

    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Firewall Log Agent" Height="700" Width="780"
        WindowStartupLocation="CenterScreen" ResizeMode="CanResizeWithGrip"
        Background="#1E1E2E">
    <Window.Resources>
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="#E4E4E7"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="FontSize"   Value="13"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background"  Value="#2D2D44"/>
            <Setter Property="Foreground"  Value="#E4E4E7"/>
            <Setter Property="BorderBrush" Value="#3F3F5C"/>
            <Setter Property="Padding"     Value="6,4"/>
            <Setter Property="FontFamily"  Value="Segoe UI"/>
        </Style>
        <Style TargetType="PasswordBox">
            <Setter Property="Background"  Value="#2D2D44"/>
            <Setter Property="Foreground"  Value="#E4E4E7"/>
            <Setter Property="BorderBrush" Value="#3F3F5C"/>
            <Setter Property="Padding"     Value="6,4"/>
            <Setter Property="FontFamily"  Value="Segoe UI"/>
        </Style>
        <Style TargetType="ComboBox">
            <Setter Property="Background"  Value="#2D2D44"/>
            <Setter Property="Foreground"  Value="#000000"/>
            <Setter Property="BorderBrush" Value="#3F3F5C"/>
            <Setter Property="Padding"     Value="4,3"/>
            <Setter Property="FontFamily"  Value="Segoe UI"/>
        </Style>
        <Style TargetType="ComboBoxItem">
            <Setter Property="Background"  Value="#2D2D44"/>
            <Setter Property="Foreground"  Value="#E4E4E7"/>
            <Setter Property="FontFamily"  Value="Segoe UI"/>
            <Setter Property="Padding"     Value="4,3"/>
        </Style>
        <Style TargetType="Button">
            <Setter Property="Background"  Value="#7C3AED"/>
            <Setter Property="Foreground"  Value="White"/>
            <Setter Property="BorderBrush" Value="#6D28D9"/>
            <Setter Property="Padding"     Value="16,8"/>
            <Setter Property="FontFamily"  Value="Segoe UI"/>
            <Setter Property="FontWeight"  Value="SemiBold"/>
            <Setter Property="Cursor"      Value="Hand"/>
        </Style>
        <Style TargetType="GroupBox">
            <Setter Property="Foreground"  Value="#E4E4E7"/>
            <Setter Property="BorderBrush" Value="#3F3F5C"/>
            <Setter Property="FontFamily"  Value="Segoe UI"/>
        </Style>
    </Window.Resources>

    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Row 0: Header -->
        <StackPanel Grid.Row="0" Margin="0,0,0,12">
            <TextBlock Text="Firewall Log Agent" FontSize="22" FontWeight="Bold" Foreground="#7C3AED"/>
            <TextBlock Text="Collects Windows Firewall logs and uploads to Azure Table Storage" FontSize="12" Foreground="#A1A1AA" Margin="0,4,0,0"/>
        </StackPanel>

        <!-- Row 1: Azure Login -->
        <GroupBox Grid.Row="1" Header="  Step 1: Connect to Azure  " Margin="0,0,0,10" Padding="10">
            <StackPanel>
                <StackPanel Orientation="Horizontal">
                    <Button Name="btnConnectAzure" Content="Connect to Azure" Width="200"/>
                    <TextBlock Name="lblLoginStatus" Text="  Not signed in" VerticalAlignment="Center" Margin="12,0,0,0" Foreground="#F59E0B"/>
                </StackPanel>
                <TextBlock Name="lblDeviceCode" Text="" Margin="0,6,0,0" FontSize="12" Foreground="#A1A1AA" TextWrapping="Wrap"/>
            </StackPanel>
        </GroupBox>

        <!-- Row 2: Resource Selection -->
        <GroupBox Grid.Row="2" Header="  Step 2: Select Resources  " Margin="0,0,0,10" Padding="10" Name="grpResources" IsEnabled="False">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="140"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <TextBlock Grid.Row="0" Grid.Column="0" Text="Subscription:" VerticalAlignment="Center" Margin="0,4"/>
                <ComboBox  Grid.Row="0" Grid.Column="1" Name="cmbSubscription" Margin="0,4" DisplayMemberPath="displayName" Grid.ColumnSpan="2">
                    <ComboBox.ItemContainerStyle>
                        <Style TargetType="ComboBoxItem">
                            <Setter Property="Background" Value="#1E1E2E"/>
                            <Setter Property="Foreground" Value="#000000"/>
                        </Style>
                    </ComboBox.ItemContainerStyle>
                </ComboBox>

                <TextBlock Grid.Row="1" Grid.Column="0" Text="Storage Account:" VerticalAlignment="Center" Margin="0,4"/>
                <ComboBox  Grid.Row="1" Grid.Column="1" Name="cmbStorageAccount" Margin="0,4" DisplayMemberPath="name" Grid.ColumnSpan="2">
                    <ComboBox.ItemContainerStyle>
                        <Style TargetType="ComboBoxItem">
                            <Setter Property="Background" Value="#1E1E2E"/>
                            <Setter Property="Foreground" Value="#000000"/>
                        </Style>
                    </ComboBox.ItemContainerStyle>
                </ComboBox>

                <TextBlock Grid.Row="2" Grid.Column="0" Text="Table Name:" VerticalAlignment="Center" Margin="0,4"/>
                <ComboBox  Grid.Row="2" Grid.Column="1" Name="cmbTable" Margin="0,4" IsEditable="True" Text="FirewallLogs">
                    <ComboBox.ItemContainerStyle>
                        <Style TargetType="ComboBoxItem">
                            <Setter Property="Background" Value="#1E1E2E"/>
                            <Setter Property="Foreground" Value="#000000"/>
                        </Style>
                    </ComboBox.ItemContainerStyle>
                </ComboBox>
                <Button    Grid.Row="2" Grid.Column="2" Name="btnSetup" Content="Setup" Width="80" Margin="6,4,0,4"
                           Background="#10B981" BorderBrush="#059669"/>
            </Grid>
        </GroupBox>

        <!-- Row 3: Monitoring Controls -->
        <StackPanel Grid.Row="3" Orientation="Horizontal" Margin="0,0,0,10">
            <Button Name="btnStart" Content="Start Monitoring" Width="180" IsEnabled="False"/>
            <Button Name="btnStop"  Content="Stop Monitoring"  Width="180" Margin="10,0,0,0" IsEnabled="False"
                    Background="#EF4444" BorderBrush="#DC2626"/>
            <TextBlock Name="lblStatus" Text="" VerticalAlignment="Center" Margin="16,0,0,0" FontSize="13"/>
        </StackPanel>

        <!-- Row 4: Activity Log -->
        <GroupBox Grid.Row="4" Header="  Activity Log  " Padding="6">
            <TextBox Name="txtLog" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
                     Background="#0F0F1A" Foreground="#A1A1AA" BorderThickness="0" FontFamily="Cascadia Mono,Consolas" FontSize="11"/>
        </GroupBox>

        <!-- Row 5: Status Bar -->
        <Grid Grid.Row="5" Margin="0,8,0,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBlock Grid.Column="0" Name="lblEntries"     Text="Entries collected: 0" FontSize="11" Foreground="#A1A1AA"/>
            <TextBlock Grid.Column="1" Name="lblLastUpload"  Text="Last upload: --"      FontSize="11" Foreground="#A1A1AA" Margin="20,0,0,0"/>
            <TextBlock Grid.Column="2" Name="lblUploadCount" Text="Uploaded: 0"           FontSize="11" Foreground="#A1A1AA" Margin="20,0,0,0"/>
        </Grid>
    </Grid>
</Window>
"@

    $reader = [System.Xml.XmlNodeReader]::new($xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)

    # ── Locate named elements ──
    $ui = @{}
    $names = @(
        "btnConnectAzure","lblLoginStatus","lblDeviceCode",
        "grpResources","cmbSubscription","cmbStorageAccount","cmbTable","btnSetup",
        "btnStart","btnStop","lblStatus",
        "txtLog","lblEntries","lblLastUpload","lblUploadCount"
    )
    foreach ($n in $names) { $ui[$n] = $window.FindName($n) }

    # ── State ──
    $state = @{
        AuthContext      = $null
        Subscriptions    = @()
        StorageAccounts  = @()
        SelectedSA       = $null
        IsMonitoring     = $false
        Timer            = $null
        EntryCount       = 0
        UploadCount      = 0
    }

    # ── Helper: append log ──
    $appendLog = {
        param([string]$Message)
        $ts = Get-Date -Format "HH:mm:ss"
        $ui["txtLog"].AppendText("[$ts] $Message`r`n")
        $ui["txtLog"].ScrollToEnd()
    }

    # ══════════════════════════════════════════
    # CONNECT TO AZURE (non-blocking via background runspace)
    # ══════════════════════════════════════════
    $ui["btnConnectAzure"].Add_Click({
        $ui["btnConnectAzure"].IsEnabled = $false
        $ui["btnConnectAzure"].Content   = "Signing in..."
        $ui["lblLoginStatus"].Text       = "  Requesting device code..."
        $ui["lblLoginStatus"].Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#3B82F6")
        & $appendLog "Initiating Azure AD Device Code Flow..."

        try {
            # Step 1 (fast): request the device code on the UI thread
            $dc = Request-DeviceCode -TenantId "common"

            # Display the code and copy to clipboard
            $ui["lblDeviceCode"].Text = "Code: $($dc.user_code)   |   Open: $($dc.verification_uri)   (code copied to clipboard)"
            [System.Windows.Clipboard]::SetText($dc.user_code)
            & $appendLog "--------------------------------------------"
            & $appendLog "AZURE LOGIN: Enter code  $($dc.user_code)  at  $($dc.verification_uri)"
            & $appendLog "--------------------------------------------"

            $ui["lblLoginStatus"].Text = "  Waiting for sign-in..."

            # Step 2: poll for the token in a background runspace (returns JSON string to avoid serialization issues)
            $runspace = [runspacefactory]::CreateRunspace()
            $runspace.Open()
            $ps = [powershell]::Create().AddScript({
                param($TenantId, $DeviceCode, $ExpiresIn, $Interval)

                $clientId  = "1950a258-227b-4e31-a9cf-717495945fc2"
                $authority = "https://login.microsoftonline.com"
                $tokenUrl  = "$authority/$TenantId/oauth2/v2.0/token"
                $pollBody  = "grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Adevice_code&client_id=$clientId&device_code=$DeviceCode"

                $expiresAt = (Get-Date).AddSeconds($ExpiresIn)
                $sleepSec  = [Math]::Max($Interval, 5)

                while ((Get-Date) -lt $expiresAt) {
                    Start-Sleep -Seconds $sleepSec
                    try {
                        # Use Invoke-WebRequest to get raw JSON string (avoids deserialization issues)
                        $resp = Invoke-WebRequest -Uri $tokenUrl -Method Post -Body $pollBody `
                            -ContentType "application/x-www-form-urlencoded" -ErrorAction Stop -UseBasicParsing
                        return "OK|$($resp.Content)"
                    }
                    catch {
                        $errBody = $null
                        try { $errBody = $_.ErrorDetails.Message | ConvertFrom-Json } catch {}
                        if ($errBody.error -eq "authorization_pending") { continue }
                        elseif ($errBody.error -eq "slow_down")         { $sleepSec += 5; continue }
                        else {
                            return "ERR|Device code auth failed: $($errBody.error_description)"
                        }
                    }
                }
                return "ERR|Device code authentication timed out."
            }).AddArgument("common").AddArgument($dc.device_code).AddArgument($dc.expires_in).AddArgument([int]$dc.interval)

            $ps.Runspace = $runspace
            $asyncResult = $ps.BeginInvoke()

            # Store handles so the timer can access them
            $state.LoginRunspace    = $runspace
            $state.LoginPowerShell  = $ps
            $state.LoginAsyncResult = $asyncResult

            # Step 3: DispatcherTimer checks every 2s if the runspace is done
            $loginTimer = New-Object System.Windows.Threading.DispatcherTimer
            $loginTimer.Interval = [TimeSpan]::FromSeconds(2)
            $loginTimer.Add_Tick({
                if (-not $state.LoginAsyncResult.IsCompleted) { return }

                # Stop this timer immediately
                $this.Stop()

                try {
                    $rawResults = $state.LoginPowerShell.EndInvoke($state.LoginAsyncResult)

                    # Clean up runspace
                    $state.LoginPowerShell.Dispose()
                    $state.LoginRunspace.Dispose()

                    $rawString = [string]$rawResults[0]
                    if (-not $rawString -or -not $rawString.StartsWith("OK|")) {
                        $errMsg = if ($rawString -and $rawString.StartsWith("ERR|")) { $rawString.Substring(4) } else { "Unknown error" }
                        throw $errMsg
                    }

                    # Parse the JSON token response (came as raw string to avoid runspace serialization issues)
                    $jsonBody = $rawString.Substring(3)
                    $result   = $jsonBody | ConvertFrom-Json

                    $authCtx = New-AzureAuthContext
                    $authCtx.AuthType       = 'AzureAD'
                    $authCtx.TenantId       = "common"
                    $authCtx.ArmToken       = $result.access_token
                    $authCtx.RefreshToken   = $result.refresh_token
                    $authCtx.ArmTokenExpiry = (Get-Date).AddSeconds($result.expires_in - 120)

                    & $appendLog "DEBUG: refresh_token present = $([bool]$authCtx.RefreshToken)"

                    $state.AuthContext = $authCtx
                    $ui["lblLoginStatus"].Text       = "  Signed in"
                    $ui["lblLoginStatus"].Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#10B981")
                    $ui["lblDeviceCode"].Text        = ""
                    $ui["btnConnectAzure"].Content   = "Connected"
                    & $appendLog "Azure AD authentication successful."

                    # Load subscriptions
                    & $appendLog "Loading subscriptions..."
                    $subs = Get-AzureSubscriptions -AuthContext $authCtx
                    $state.Subscriptions = $subs

                    $ui["cmbSubscription"].ItemsSource = $subs
                    if ($subs.Count -gt 0) {
                        $ui["cmbSubscription"].SelectedIndex = 0
                    }

                    $ui["grpResources"].IsEnabled = $true
                    & $appendLog "Found $($subs.Count) subscription(s)."
                }
                catch {
                    $ui["lblLoginStatus"].Text       = "  Sign-in failed"
                    $ui["lblLoginStatus"].Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#EF4444")
                    $ui["lblDeviceCode"].Text        = ""
                    $ui["btnConnectAzure"].Content   = "Connect to Azure"
                    $ui["btnConnectAzure"].IsEnabled = $true
                    & $appendLog "ERROR: $_"
                }
            })
            $loginTimer.Start()
        }
        catch {
            $ui["lblLoginStatus"].Text       = "  Sign-in failed"
            $ui["lblLoginStatus"].Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#EF4444")
            $ui["btnConnectAzure"].Content   = "Connect to Azure"
            $ui["btnConnectAzure"].IsEnabled = $true
            & $appendLog "ERROR: $_"
        }
    })

    # ══════════════════════════════════════════
    # SUBSCRIPTION CHANGED → load storage accounts
    # ══════════════════════════════════════════
    $ui["cmbSubscription"].Add_SelectionChanged({
        $sel = $ui["cmbSubscription"].SelectedItem
        if (-not $sel -or -not $state.AuthContext) { return }

        $ui["cmbStorageAccount"].ItemsSource = $null
        $ui["cmbTable"].ItemsSource = $null
        & $appendLog "Loading storage accounts for '$($sel.displayName)'..."

        try {
            $accounts = Get-AzureStorageAccounts -AuthContext $state.AuthContext -SubscriptionId $sel.subscriptionId
            $state.StorageAccounts = $accounts

            $ui["cmbStorageAccount"].ItemsSource = $accounts
            if ($accounts.Count -gt 0) {
                $ui["cmbStorageAccount"].SelectedIndex = 0
            }
            & $appendLog "Found $($accounts.Count) storage account(s)."
        }
        catch {
            & $appendLog "ERROR loading storage accounts: $_"
        }
    })

    # ══════════════════════════════════════════
    # STORAGE ACCOUNT CHANGED → retrieve keys + list tables
    # ══════════════════════════════════════════
    # Helper: list tables and populate the Table combobox
    $loadTables = {
        & $appendLog "Listing tables..."
        $tables = Get-AzureStorageTables -AuthContext $state.AuthContext
        if ($tables -and $tables.Count -gt 0) {
            $ui["cmbTable"].ItemsSource = $tables
            $idx = [Array]::IndexOf($tables, "FirewallLogs")
            if ($idx -ge 0) { $ui["cmbTable"].SelectedIndex = $idx }
            else { $ui["cmbTable"].SelectedIndex = 0 }
            & $appendLog "Found $($tables.Count) table(s)."
        }
        else {
            $ui["cmbTable"].ItemsSource = $null
            $ui["cmbTable"].Text = "FirewallLogs"
            & $appendLog "No existing tables. A new one will be created."
        }
    }

    # Helper: switch auth context to RBAC bearer token
    $switchToRbac = {
        param([string]$AccountName)
        & $appendLog "Switching to RBAC bearer token for '$AccountName'..."
        $state.AuthContext.StorageAccountName = $AccountName
        $state.AuthContext = Get-StorageToken -AuthContext $state.AuthContext
        & $appendLog "Storage bearer token acquired (RBAC)."
    }

    $ui["cmbStorageAccount"].Add_SelectionChanged({
        $sa  = $ui["cmbStorageAccount"].SelectedItem
        $sub = $ui["cmbSubscription"].SelectedItem
        if (-not $sa -or -not $sub -or -not $state.AuthContext) { return }

        $state.SelectedSA = $sa
        $state.AuthContext.StorageAccountName = $sa.name

        # Check if shared-key access is disabled on this account
        $keyAuthAllowed = $true
        try {
            if ($sa.properties.allowSharedKeyAccess -eq $false) {
                $keyAuthAllowed = $false
            }
        } catch {}

        if (-not $keyAuthAllowed) {
            & $appendLog "Key-based auth is disabled on '$($sa.name)'. Using RBAC."
            try {
                & $switchToRbac $sa.name
                & $loadTables
            }
            catch {
                & $appendLog "ERROR setting up RBAC auth: $_"
            }
            return
        }

        # Key auth allowed — try getting keys
        & $appendLog "Retrieving keys for '$($sa.name)'..."
        try {
            $keys = Get-AzureStorageAccountKeys -AuthContext $state.AuthContext `
                -SubscriptionId $sub.subscriptionId `
                -ResourceGroup  $sa.resourceGroup `
                -StorageAccountName $sa.name

            $firstKey = ($keys | Select-Object -First 1).value
            if (-not $firstKey) { throw "No access keys returned." }

            $state.AuthContext = Set-AccessKeyAuth -AuthContext $state.AuthContext -AccessKey $firstKey
            & $appendLog "Access key retrieved for '$($sa.name)'."

            # Try listing tables with the key — if 403, fall back to RBAC
            try {
                & $loadTables
            }
            catch {
                & $appendLog "Table listing failed: $_"
                & $appendLog "Falling back to RBAC..."
                & $switchToRbac $sa.name
                try { & $loadTables } catch {
                    $ui["cmbTable"].ItemsSource = $null
                    $ui["cmbTable"].Text = "FirewallLogs"
                    & $appendLog "Could not list tables (will create): $_"
                }
            }
        }
        catch {
            & $appendLog "Could not use key auth: $_"
            & $appendLog "Attempting RBAC fallback..."
            try {
                & $switchToRbac $sa.name
                & $loadTables
            }
            catch {
                & $appendLog "RBAC fallback also failed: $_"
            }
        }
    })

    # ══════════════════════════════════════════
    # SETUP BUTTON → create/confirm table
    # ══════════════════════════════════════════
    $ui["btnSetup"].Add_Click({
        if (-not $state.AuthContext -or -not $state.AuthContext.StorageAccountName) {
            & $appendLog "ERROR: Select a storage account first."
            return
        }

        $tableName = $ui["cmbTable"].Text.Trim()
        if (-not $tableName) { $tableName = "FirewallLogs" }

        & $appendLog "Ensuring table '$tableName' exists..."
        try {
            $null = New-StorageTable -AuthContext $state.AuthContext -TableName $tableName
            & $appendLog "Table '$tableName' is ready."
            $ui["btnStart"].IsEnabled = $true
            $ui["lblStatus"].Text     = "Ready to monitor"
            $ui["lblStatus"].Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#10B981")
        }
        catch {
            # If 403 (key auth not permitted), retry with RBAC bearer
            $is403 = $false
            try { $is403 = $_.Exception.Response.StatusCode.value__ -eq 403 } catch {}
            if ($is403 -and $state.AuthContext.AuthType -eq 'AccessKey' -and $state.AuthContext.RefreshToken) {
                & $appendLog "Key auth rejected (403). Switching to RBAC bearer token..."
                try {
                    $state.AuthContext = Get-StorageToken -AuthContext $state.AuthContext
                    & $appendLog "RBAC token acquired. Retrying table creation..."
                    $null = New-StorageTable -AuthContext $state.AuthContext -TableName $tableName
                    & $appendLog "Table '$tableName' is ready."
                    $ui["btnStart"].IsEnabled = $true
                    $ui["lblStatus"].Text     = "Ready to monitor"
                    $ui["lblStatus"].Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#10B981")
                }
                catch {
                    & $appendLog "ERROR creating table with RBAC: $_"
                }
            }
            else {
                & $appendLog "ERROR creating table: $_"
            }
        }
    })

    # ══════════════════════════════════════════
    # START MONITORING
    # ══════════════════════════════════════════
    $ui["btnStart"].Add_Click({
        if ($state.IsMonitoring) { return }

        & $appendLog "Enabling Windows Firewall logging..."
        try {
            Enable-FirewallLogging
            & $appendLog "Firewall logging enabled."
        }
        catch {
            & $appendLog "WARNING: Could not enable firewall logging (need admin?): $_"
        }

        & $appendLog "Building firewall rule name cache..."
        try { Update-FirewallRuleCache } catch { & $appendLog "WARNING: Could not build rule cache: $_" }

        $null = Get-NewFirewallLogEntries -ResetPosition
        & $appendLog "File position initialized. Starting collection timer (60s interval)."

        $timer = New-Object System.Windows.Threading.DispatcherTimer
        $timer.Interval = [TimeSpan]::FromSeconds(60)
        $timer.Add_Tick({
            if (-not $state.IsMonitoring) { return }

            try {
                & $appendLog "Reading new log entries..."
                $entries = Get-NewFirewallLogEntries
                $count   = $entries.Count

                if ($count -gt 0) {
                    $state.EntryCount += $count
                    $ui["lblEntries"].Text = "Entries collected: $($state.EntryCount)"
                    & $appendLog "Found $count new entries. Uploading (AuthType=$($state.AuthContext.AuthType))..."

                    $tableName = $ui["cmbTable"].Text.Trim()

                    # Try inserting; on first 403, switch to RBAC and retry
                    $result = Add-StorageTableEntities -AuthContext $state.AuthContext -TableName $tableName -Entities $entries

                    if ($result.Errors -gt 0 -and $state.AuthContext.RefreshToken) {
                        & $appendLog "$($result.Errors) errors. Switching to RBAC and retrying..."
                        try {
                            $state.AuthContext = Get-StorageToken -AuthContext $state.AuthContext
                            $result = Add-StorageTableEntities -AuthContext $state.AuthContext -TableName $tableName -Entities $entries
                        }
                        catch {
                            & $appendLog "RBAC retry failed: $_"
                        }
                    }

                    $state.UploadCount += $result.Success
                    $ui["lblUploadCount"].Text = "Uploaded: $($state.UploadCount)"
                    $ui["lblLastUpload"].Text  = "Last upload: $(Get-Date -Format 'HH:mm:ss')"

                    if ($result.Errors -gt 0) {
                        & $appendLog "Upload: $($result.Success) OK, $($result.Errors) errors."
                    }
                    else {
                        & $appendLog "Upload complete: $($result.Success) entities."
                    }
                }
                else {
                    & $appendLog "No new entries."
                }
            }
            catch {
                & $appendLog "UPLOAD ERROR: $_"
            }
        })

        $timer.Start()
        $state.Timer        = $timer
        $state.IsMonitoring = $true
        $ui["btnStart"].IsEnabled   = $false
        $ui["btnStop"].IsEnabled    = $true
        $ui["grpResources"].IsEnabled = $false
        $ui["btnConnectAzure"].IsEnabled = $false
        $ui["lblStatus"].Text       = "Monitoring active"
        $ui["lblStatus"].Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#10B981")
        & $appendLog "Monitoring started."
    })

    # ══════════════════════════════════════════
    # STOP MONITORING
    # ══════════════════════════════════════════
    $ui["btnStop"].Add_Click({
        if (-not $state.IsMonitoring) { return }

        if ($state.Timer) { $state.Timer.Stop(); $state.Timer = $null }
        $state.IsMonitoring = $false

        & $appendLog "Restoring firewall logging settings..."
        try { Restore-FirewallLogging } catch { & $appendLog "WARNING: $_" }

        $ui["btnStart"].IsEnabled   = $true
        $ui["btnStop"].IsEnabled    = $false
        $ui["grpResources"].IsEnabled = $true
        $ui["btnConnectAzure"].IsEnabled = $true
        $ui["lblStatus"].Text       = "Monitoring stopped"
        $ui["lblStatus"].Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#F59E0B")
        & $appendLog "Monitoring stopped."
    })

    # ── Window Close ──
    $window.Add_Closing({
        if ($state.IsMonitoring) {
            if ($state.Timer) { $state.Timer.Stop() }
            try { Restore-FirewallLogging } catch {}
        }
        if ($OnExit) { & $OnExit }
    })

    $window.ShowDialog() | Out-Null
}
