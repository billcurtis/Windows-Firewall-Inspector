<#
.SYNOPSIS
    WPF Dark-Mode GUI for the Windows Firewall Viewer / Dashboard.
.DESCRIPTION
    Modern dark-themed WPF interface with data grid, filters, bar/pie charts,
    IP owner resolution, and CSV export. Self-contained drawing routines for charts.
#>

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms   # for SaveFileDialog

# ───────────────── colour palette ─────────────────
$script:Colors = @{
    BgDark      = "#1A1008"
    BgSurface   = "#231A0E"
    BgElevated  = "#352812"
    BgCard      = "#2C2010"
    Primary     = "#C2650A"
    PrimaryDim  = "#A85608"
    Accent      = "#E8944A"
    Text        = "#E8E0D6"
    SubText     = "#A89882"
    Border      = "#4A3520"
    Success     = "#10B981"
    Danger      = "#EF4444"
    Warning     = "#F59E0B"
    Info        = "#3B82F6"
}

$script:ModuleDir = $PSScriptRoot   # resolved at dot-source time

# ───────────────── main window ─────────────────

function Show-ViewerWindow {
    [CmdletBinding()]
    param()

    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Windows Firewall Viewer" Height="960" Width="1440"
        WindowStartupLocation="CenterScreen" ResizeMode="CanResizeWithGrip"
        Background="$($script:Colors.BgDark)">

    <Window.Resources>
        <!-- Global text style -->
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="$($script:Colors.Text)"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
        </Style>
        <Style TargetType="TextBox" x:Key="DarkTextBox">
            <Setter Property="Background"  Value="$($script:Colors.BgElevated)"/>
            <Setter Property="Foreground"  Value="$($script:Colors.Text)"/>
            <Setter Property="BorderBrush" Value="$($script:Colors.Border)"/>
            <Setter Property="Padding"     Value="8,5"/>
            <Setter Property="FontFamily"  Value="Segoe UI"/>
            <Setter Property="FontSize"    Value="13"/>
        </Style>
        <Style TargetType="PasswordBox" x:Key="DarkPassword">
            <Setter Property="Background"  Value="$($script:Colors.BgElevated)"/>
            <Setter Property="Foreground"  Value="$($script:Colors.Text)"/>
            <Setter Property="BorderBrush" Value="$($script:Colors.Border)"/>
            <Setter Property="Padding"     Value="8,5"/>
            <Setter Property="FontFamily"  Value="Segoe UI"/>
        </Style>
        <Style TargetType="Button" x:Key="PrimaryBtn">
            <Setter Property="Background"   Value="$($script:Colors.Primary)"/>
            <Setter Property="Foreground"   Value="White"/>
            <Setter Property="BorderBrush"  Value="$($script:Colors.PrimaryDim)"/>
            <Setter Property="Padding"      Value="18,8"/>
            <Setter Property="FontFamily"   Value="Segoe UI"/>
            <Setter Property="FontWeight"   Value="SemiBold"/>
            <Setter Property="FontSize"     Value="13"/>
            <Setter Property="Cursor"       Value="Hand"/>
            <Setter Property="BorderThickness" Value="0"/>
        </Style>
        <Style TargetType="Button" x:Key="SecondaryBtn" BasedOn="{StaticResource PrimaryBtn}">
            <Setter Property="Background"  Value="$($script:Colors.BgElevated)"/>
            <Setter Property="BorderBrush" Value="$($script:Colors.Border)"/>
            <Setter Property="BorderThickness" Value="1"/>
        </Style>
        <Style TargetType="ComboBox" x:Key="DarkCombo">
            <Setter Property="Background"  Value="$($script:Colors.BgElevated)"/>
            <Setter Property="Foreground"  Value="#000000"/>
            <Setter Property="BorderBrush" Value="$($script:Colors.Border)"/>
            <Setter Property="Padding"     Value="6,4"/>
            <Setter Property="FontFamily"  Value="Segoe UI"/>
            <Setter Property="ItemContainerStyle">
                <Setter.Value>
                    <Style TargetType="ComboBoxItem">
                        <Setter Property="Foreground" Value="#000000"/>
                    </Style>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="DatePicker" x:Key="DarkDate">
            <Setter Property="Background"  Value="$($script:Colors.BgElevated)"/>
            <Setter Property="Foreground"  Value="#000000"/>
            <Setter Property="BorderBrush" Value="$($script:Colors.Border)"/>
            <Setter Property="FontFamily"  Value="Segoe UI"/>
        </Style>
        <Style TargetType="DataGrid" x:Key="DarkGrid">
            <Setter Property="Background"            Value="$($script:Colors.BgSurface)"/>
            <Setter Property="Foreground"            Value="$($script:Colors.Text)"/>
            <Setter Property="BorderBrush"           Value="$($script:Colors.Border)"/>
            <Setter Property="RowBackground"         Value="$($script:Colors.BgSurface)"/>
            <Setter Property="AlternatingRowBackground" Value="$($script:Colors.BgCard)"/>
            <Setter Property="GridLinesVisibility"   Value="None"/>
            <Setter Property="HeadersVisibility"     Value="Column"/>
            <Setter Property="FontFamily"            Value="Segoe UI"/>
            <Setter Property="FontSize"              Value="12"/>
            <Setter Property="IsReadOnly"            Value="True"/>
            <Setter Property="AutoGenerateColumns"   Value="False"/>
            <Setter Property="SelectionMode"         Value="Extended"/>
            <Setter Property="CanUserSortColumns"    Value="True"/>
        </Style>
        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="$($script:Colors.BgElevated)"/>
            <Setter Property="Foreground" Value="$($script:Colors.Accent)"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding"    Value="8,6"/>
            <Setter Property="BorderBrush" Value="$($script:Colors.Border)"/>
            <Setter Property="BorderThickness" Value="0,0,1,1"/>
        </Style>
        <Style TargetType="MenuItem">
            <Setter Property="Background" Value="$($script:Colors.BgSurface)"/>
            <Setter Property="Foreground" Value="$($script:Colors.Text)"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="Padding"    Value="6,4"/>
            <Style.Triggers>
                <Trigger Property="IsHighlighted" Value="True">
                    <Setter Property="Background" Value="$($script:Colors.BgElevated)"/>
                    <Setter Property="Foreground" Value="$($script:Colors.Accent)"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="Separator" x:Key="{x:Static MenuItem.SeparatorStyleKey}">
            <Setter Property="Background" Value="$($script:Colors.Border)"/>
            <Setter Property="Margin" Value="4,2"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Separator">
                        <Border Height="1" Background="{TemplateBinding Background}" Margin="{TemplateBinding Margin}"/>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <DockPanel>
        <!-- ═══ MENU BAR ═══ -->
        <Menu DockPanel.Dock="Top" Background="$($script:Colors.BgSurface)" Foreground="$($script:Colors.Text)" FontFamily="Segoe UI" FontSize="13" Padding="4,2">
            <MenuItem Header="_File">
                <MenuItem Name="mnuLoadCsv"   Header="Load CSV File"      InputGestureText="Ctrl+O"/>
                <MenuItem Name="mnuExportCsv"  Header="Export to CSV"     InputGestureText="Ctrl+S"/>
                <Separator/>
                <MenuItem Name="mnuExit"       Header="Exit"             InputGestureText="Alt+F4"/>
            </MenuItem>
            <MenuItem Header="_Actions">
                <MenuItem Name="mnuQuery"       Header="Query / Apply Filters"  InputGestureText="F5"/>
                <MenuItem Name="mnuClear"       Header="Clear Filters"/>
                <Separator/>
                <MenuItem Name="mnuLookupIPs"   Header="Lookup IP Owners"/>
                <MenuItem Name="mnuServiceTags" Header="Lookup MS Service Tags"/>
                <Separator/>
                <MenuItem Name="mnuConnectAzure" Header="Connect to Azure"/>
            </MenuItem>
        </Menu>

        <!-- ═══ TOP BAR ═══ -->
        <Border DockPanel.Dock="Top" Background="$($script:Colors.BgSurface)" Padding="16,10" BorderBrush="$($script:Colors.Border)" BorderThickness="0,0,0,1">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Text="&#x1F525; Windows Firewall Viewer" FontSize="20" FontWeight="Bold" Foreground="$($script:Colors.Primary)" VerticalAlignment="Center"/>
                <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
                    <Ellipse Name="statusDot" Width="10" Height="10" Fill="$($script:Colors.Danger)" Margin="0,0,6,0"/>
                    <TextBlock Name="lblConnStatus" Text="Disconnected" FontSize="12" Foreground="$($script:Colors.SubText)"/>
                </StackPanel>
            </Grid>
        </Border>

        <!-- ═══ STATUS BAR ═══ -->
        <Border DockPanel.Dock="Bottom" Background="$($script:Colors.BgSurface)" Padding="16,6" BorderBrush="$($script:Colors.Border)" BorderThickness="0,1,0,0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Name="lblStatusMsg" Text="Ready" FontSize="11" Foreground="$($script:Colors.SubText)"/>
                <TextBlock Grid.Column="1" Name="lblRecordCount" Text="Records: 0" FontSize="11" Foreground="$($script:Colors.SubText)" Margin="20,0"/>
                <TextBlock Grid.Column="2" Name="lblFilterCount" Text="" FontSize="11" Foreground="$($script:Colors.Info)" Margin="20,0"/>
            </Grid>
        </Border>

        <!-- ═══ MAIN CONTENT ═══ -->
        <Grid>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="300"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <!-- ──── LEFT SIDEBAR ──── -->
            <Border Grid.Column="0" Background="$($script:Colors.BgSurface)" BorderBrush="$($script:Colors.Border)" BorderThickness="0,0,1,0">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <StackPanel Margin="14">

                        <!-- Connection -->
                        <TextBlock Text="CONNECTION" FontSize="11" FontWeight="Bold" Foreground="$($script:Colors.Accent)" Margin="0,0,0,8"/>

                        <Button Name="btnLogin" Content="&#x1F512;  Connect to Azure" Style="{StaticResource PrimaryBtn}" Margin="0,0,0,10" HorizontalAlignment="Stretch"/>

                        <TextBlock Text="Subscription" FontSize="11" Foreground="$($script:Colors.SubText)" Margin="0,0,0,4"/>
                        <ComboBox Name="cmbSubscription" Style="{StaticResource DarkCombo}" IsEnabled="False"/>

                        <TextBlock Text="Storage Account" FontSize="11" Foreground="$($script:Colors.SubText)" Margin="0,8,0,4"/>
                        <ComboBox Name="cmbStorageAccount" Style="{StaticResource DarkCombo}" IsEnabled="False"/>

                        <TextBlock Text="Table" FontSize="11" Foreground="$($script:Colors.SubText)" Margin="0,8,0,4"/>
                        <ComboBox Name="cmbTable" Style="{StaticResource DarkCombo}" IsEnabled="False"/>

                        <!-- Separator -->
                        <Rectangle Height="1" Fill="$($script:Colors.Border)" Margin="0,16"/>

                        <!-- Filters -->
                        <TextBlock Text="FILTERS" FontSize="11" FontWeight="Bold" Foreground="$($script:Colors.Accent)" Margin="0,0,0,8"/>

                        <TextBlock Text="Date From" FontSize="11" Foreground="$($script:Colors.SubText)" Margin="0,0,0,4"/>
                        <DatePicker Name="dpFrom" Style="{StaticResource DarkDate}"/>

                        <TextBlock Text="Date To" FontSize="11" Foreground="$($script:Colors.SubText)" Margin="0,8,0,4"/>
                        <DatePicker Name="dpTo" Style="{StaticResource DarkDate}"/>

                        <TextBlock Text="Action" FontSize="11" Foreground="$($script:Colors.SubText)" Margin="0,8,0,4"/>
                        <ComboBox Name="cmbAction" Style="{StaticResource DarkCombo}">
                            <ComboBoxItem Content="All" IsSelected="True"/>
                            <ComboBoxItem Content="ALLOW"/>
                            <ComboBoxItem Content="DROP"/>
                        </ComboBox>

                        <TextBlock Text="Protocol" FontSize="11" Foreground="$($script:Colors.SubText)" Margin="0,8,0,4"/>
                        <ComboBox Name="cmbProtocol" Style="{StaticResource DarkCombo}">
                            <ComboBoxItem Content="All" IsSelected="True"/>
                            <ComboBoxItem Content="TCP"/>
                            <ComboBoxItem Content="UDP"/>
                            <ComboBoxItem Content="ICMP"/>
                        </ComboBox>

                        <TextBlock Text="Direction" FontSize="11" Foreground="$($script:Colors.SubText)" Margin="0,8,0,4"/>
                        <ComboBox Name="cmbDirection" Style="{StaticResource DarkCombo}">
                            <ComboBoxItem Content="All" IsSelected="True"/>
                            <ComboBoxItem Content="SEND"/>
                            <ComboBoxItem Content="RECEIVE"/>
                        </ComboBox>

                        <TextBlock Text="IP Address (contains)" FontSize="11" Foreground="$($script:Colors.SubText)" Margin="0,8,0,4"/>
                        <TextBox Name="txtIPFilter" Style="{StaticResource DarkTextBox}"/>

                        <TextBlock Text="Port" FontSize="11" Foreground="$($script:Colors.SubText)" Margin="0,8,0,4"/>
                        <TextBox Name="txtPortFilter" Style="{StaticResource DarkTextBox}"/>

                        <TextBlock Text="Process Name" FontSize="11" Foreground="$($script:Colors.SubText)" Margin="0,8,0,4"/>
                        <TextBox Name="txtProcessFilter" Style="{StaticResource DarkTextBox}"/>

                        <Button Name="btnQuery" Content="&#x1F50D;  Query Data" Style="{StaticResource PrimaryBtn}" Margin="0,12,0,0" HorizontalAlignment="Stretch"/>
                        <Button Name="btnClear" Content="Clear Filters" Style="{StaticResource SecondaryBtn}" Margin="0,6,0,0" HorizontalAlignment="Stretch"/>

                        <!-- Separator -->
                        <Rectangle Height="1" Fill="$($script:Colors.Border)" Margin="0,16"/>

                        <!-- Actions -->
                        <TextBlock Text="ACTIONS" FontSize="11" FontWeight="Bold" Foreground="$($script:Colors.Accent)" Margin="0,0,0,8"/>
                        <Button Name="btnLoadCsv"    Content="&#x1F4C2;  Load CSV File"   Style="{StaticResource SecondaryBtn}" Margin="0,0,0,6" HorizontalAlignment="Stretch"/>
                        <Button Name="btnExportCsv"  Content="&#x1F4BE;  Export to CSV"  Style="{StaticResource SecondaryBtn}" Margin="0,0,0,6" HorizontalAlignment="Stretch"/>
                        <Button Name="btnLookupIPs"  Content="&#x1F310;  Lookup IP Owners" Style="{StaticResource SecondaryBtn}" Margin="0,0,0,6" HorizontalAlignment="Stretch"/>
                        <Button Name="btnServiceTags" Content="&#x2601;  Lookup MS Service Tags" Style="{StaticResource SecondaryBtn}" HorizontalAlignment="Stretch"/>
                    </StackPanel>
                </ScrollViewer>
            </Border>

            <!-- ──── RIGHT CONTENT: Data Grid ──── -->
            <DataGrid Name="dgData" Grid.Column="1" Style="{StaticResource DarkGrid}" Margin="8">
                <DataGrid.Columns>
                    <DataGridTextColumn Header="Date"        Binding="{Binding Date}"        Width="90"/>
                    <DataGridTextColumn Header="Time"        Binding="{Binding Time}"        Width="80"/>
                    <DataGridTextColumn Header="Action"      Binding="{Binding Action}"      Width="70"/>
                    <DataGridTextColumn Header="Protocol"    Binding="{Binding Protocol}"    Width="70"/>
                    <DataGridTextColumn Header="Source IP"   Binding="{Binding SrcIP}"       Width="130"/>
                    <DataGridTextColumn Header="Dest IP"     Binding="{Binding DstIP}"       Width="130"/>
                    <DataGridTextColumn Header="Src Port"    Binding="{Binding SrcPort}"     Width="70"/>
                    <DataGridTextColumn Header="Dst Port"    Binding="{Binding DstPort}"     Width="70"/>
                    <DataGridTextColumn Header="Direction"   Binding="{Binding Direction}"   Width="80"/>
                    <DataGridTextColumn Header="Size"        Binding="{Binding Size}"        Width="60"/>
                    <DataGridTextColumn Header="Process"     Binding="{Binding ProcessName}" Width="110"/>
                    <DataGridTextColumn Header="PID"         Binding="{Binding ProcessId}"   Width="55"/>
                    <DataGridTextColumn Header="Rule Name"   Binding="{Binding RuleName}"     Width="180"/>
                    <DataGridTextColumn Header="App Name"    Binding="{Binding EventAppName}"  Width="200"/>
                    <DataGridTextColumn Header="SrcIP Owner" Binding="{Binding SrcIPOwner}"    Width="160"/>
                    <DataGridTextColumn Header="DstIP Owner" Binding="{Binding DstIPOwner}"    Width="160"/>
                    <DataGridTextColumn Header="Dst Location" Binding="{Binding DstIPLocation}" Width="140"/>
                    <DataGridTextColumn Header="Dst MS Service" Binding="{Binding DstServiceTag}" Width="200"/>
                    <DataGridTextColumn Header="Computer"    Binding="{Binding ComputerName}"   Width="110"/>
                </DataGrid.Columns>
            </DataGrid>
        </Grid>
    </DockPanel>
</Window>
"@

    $reader = [System.Xml.XmlNodeReader]::new($xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)

    # ── locate named controls ──
    $ui = @{}
    $elementNames = @(
        "statusDot","lblConnStatus",
        "lblStatusMsg","lblRecordCount","lblFilterCount",
        "btnLogin","cmbSubscription","cmbStorageAccount","cmbTable",
        "dpFrom","dpTo","cmbAction","cmbProtocol","cmbDirection",
        "txtIPFilter","txtPortFilter","txtProcessFilter",
        "btnQuery","btnClear","btnLoadCsv","btnExportCsv","btnLookupIPs","btnServiceTags",
        "dgData",
        "mnuLoadCsv","mnuExportCsv","mnuExit","mnuQuery","mnuClear","mnuLookupIPs","mnuServiceTags","mnuConnectAzure"
    )
    foreach ($n in $elementNames) { $ui[$n] = $window.FindName($n) }

    # ── state ──
    $state = @{
        AuthContext    = $null
        RawData        = @()
        FilteredData   = @()
        IPOwners       = @{}
        ServiceTags    = @{}
        Subscriptions  = @()
        StorageAccounts = @()
    }

    # Helper: update status bar
    $setStatus = {
        param([string]$Msg, [string]$Color = $script:Colors.SubText)
        $ui["lblStatusMsg"].Text = $Msg
        $ui["lblStatusMsg"].Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($Color)
    }

    # ─── Reusable: apply all client-side filters to $state.RawData ───
    $applyFilters = {
        $filtered = $state.RawData

        # Date range
        $dateFrom = $ui["dpFrom"].SelectedDate
        $dateTo   = $ui["dpTo"].SelectedDate
        if ($dateFrom) {
            $fromStr = $dateFrom.ToString('yyyy-MM-dd')
            $filtered = $filtered | Where-Object { $_.Date -ge $fromStr }
        }
        if ($dateTo) {
            $toStr = $dateTo.ToString('yyyy-MM-dd')
            $filtered = $filtered | Where-Object { $_.Date -le $toStr }
        }

        # Dropdown filters
        $actionSel = $ui["cmbAction"].SelectedItem.Content
        if ($actionSel -and $actionSel -ne "All") {
            $filtered = $filtered | Where-Object { $_.Action -eq $actionSel }
        }
        $protoSel = $ui["cmbProtocol"].SelectedItem.Content
        if ($protoSel -and $protoSel -ne "All") {
            $filtered = $filtered | Where-Object { $_.Protocol -eq $protoSel }
        }
        $dirSel = $ui["cmbDirection"].SelectedItem.Content
        if ($dirSel -and $dirSel -ne "All") {
            $filtered = $filtered | Where-Object { $_.Direction -eq $dirSel }
        }

        # Text filters
        $ipFilter = $ui["txtIPFilter"].Text.Trim()
        if ($ipFilter) {
            $filtered = $filtered | Where-Object { $_.SrcIP -like "*$ipFilter*" -or $_.DstIP -like "*$ipFilter*" }
        }
        $portFilter = $ui["txtPortFilter"].Text.Trim()
        if ($portFilter) {
            $filtered = $filtered | Where-Object { "$($_.SrcPort)" -eq $portFilter -or "$($_.DstPort)" -eq $portFilter }
        }
        $procFilter = $ui["txtProcessFilter"].Text.Trim()
        if ($procFilter) {
            $filtered = $filtered | Where-Object { $_.ProcessName -like "*$procFilter*" }
        }

        # Attach enrichment data already in state
        foreach ($row in $filtered) {
            $srcInfo = if ($state.IPOwners.ContainsKey($row.SrcIP)) { $state.IPOwners[$row.SrcIP] } else { $null }
            $dstInfo = if ($state.IPOwners.ContainsKey($row.DstIP)) { $state.IPOwners[$row.DstIP] } else { $null }
            $row | Add-Member -NotePropertyName "SrcIPOwner"    -NotePropertyValue $(if ($srcInfo) { $srcInfo.Name } else { "" }) -Force
            $row | Add-Member -NotePropertyName "DstIPOwner"    -NotePropertyValue $(if ($dstInfo) { $dstInfo.Name } else { "" }) -Force
            $row | Add-Member -NotePropertyName "DstIPLocation" -NotePropertyValue $(if ($dstInfo -and $dstInfo.City -ne 'N/A') { "$($dstInfo.City), $($dstInfo.Country)" } else { "" }) -Force
            $row | Add-Member -NotePropertyName "DstServiceTag" -NotePropertyValue $(if ($state.ServiceTags.ContainsKey($row.DstIP)) { $state.ServiceTags[$row.DstIP] } else { "" }) -Force
        }

        $state.FilteredData = @($filtered)
        $ui["dgData"].ItemsSource = $state.FilteredData
        $totalCount = $state.RawData.Count
        $ui["lblRecordCount"].Text = "Records: $totalCount"
        $ui["lblFilterCount"].Text = if ($filtered.Count -ne $totalCount) { "Showing: $($filtered.Count)" } else { "" }
    }

    # ─── Login to Azure ───
    $ui["btnLogin"].Add_Click({
        & $setStatus "Logging in to Azure..." $script:Colors.Info

        try {
            $null = Connect-AzureViewer
            & $setStatus "Logged in. Loading subscriptions..." $script:Colors.Info

            $subs = @(Get-ViewerSubscriptions)
            $state.Subscriptions = $subs

            $ui["cmbSubscription"].Items.Clear()
            foreach ($s in $subs) {
                $item = [System.Windows.Controls.ComboBoxItem]::new()
                $item.Content    = "$($s.Name)  ($($s.Id))"
                $item.Tag        = $s.Id
                $ui["cmbSubscription"].Items.Add($item)
            }
            $ui["cmbSubscription"].IsEnabled = $true
            $ui["cmbStorageAccount"].Items.Clear()
            $ui["cmbStorageAccount"].IsEnabled = $false
            $ui["cmbTable"].Items.Clear()
            $ui["cmbTable"].IsEnabled = $false

            $ui["statusDot"].Fill = [System.Windows.Media.BrushConverter]::new().ConvertFromString($script:Colors.Warning)
            $ui["lblConnStatus"].Text = "Logged In"
            & $setStatus "Select a subscription." $script:Colors.Success
        }
        catch {
            $ui["statusDot"].Fill = [System.Windows.Media.BrushConverter]::new().ConvertFromString($script:Colors.Danger)
            $ui["lblConnStatus"].Text = "Login Failed"
            & $setStatus "Login failed: $_" $script:Colors.Danger
        }
    })

    # ─── Subscription changed → load storage accounts ───
    $ui["cmbSubscription"].Add_SelectionChanged({
        $selected = $ui["cmbSubscription"].SelectedItem
        if (-not $selected) { return }
        $subId = $selected.Tag

        $ui["cmbStorageAccount"].Items.Clear()
        $ui["cmbStorageAccount"].IsEnabled = $false
        $ui["cmbTable"].Items.Clear()
        $ui["cmbTable"].IsEnabled = $false

        & $setStatus "Loading storage accounts..." $script:Colors.Info

        try {
            $accounts = @(Get-ViewerStorageAccounts -SubscriptionId $subId)
            $state.StorageAccounts = $accounts

            foreach ($a in $accounts) {
                $item = [System.Windows.Controls.ComboBoxItem]::new()
                $item.Content    = $a.StorageAccountName
                $item.Tag        = "$($a.ResourceGroupName)|$($a.StorageAccountName)"
                $ui["cmbStorageAccount"].Items.Add($item)
            }
            $ui["cmbStorageAccount"].IsEnabled = $true
            & $setStatus "Found $($accounts.Count) table-capable storage account(s). Select one." $script:Colors.Success
        }
        catch {
            & $setStatus "Failed to load storage accounts: $_" $script:Colors.Danger
        }
    })

    # ─── Storage account changed → load tables ───
    $ui["cmbStorageAccount"].Add_SelectionChanged({
        $selected = $ui["cmbStorageAccount"].SelectedItem
        if (-not $selected) { return }

        $parts = $selected.Tag -split '\|'
        $rg    = $parts[0]
        $acct  = $parts[1]

        $ui["cmbTable"].Items.Clear()
        $ui["cmbTable"].IsEnabled = $false

        & $setStatus "Loading tables for '$acct'..." $script:Colors.Info

        try {
            # Build auth context — try access key first, fall back to bearer
            $ctx = New-AzureAuthContext -StorageAccountName $acct

            try {
                $keys = Get-ViewerStorageAccountKeys -ResourceGroupName $rg -StorageAccountName $acct
                $ctx  = Set-AccessKeyAuth -AuthContext $ctx -AccessKey $keys[0].Value
            }
            catch {
                Write-Warning "Could not retrieve access keys, using Azure AD bearer token."
                $ctx = Set-AzureADAuth -AuthContext $ctx
            }

            $tables = @(Get-ViewerStorageTables -AuthContext $ctx)

            foreach ($t in $tables) {
                $item = [System.Windows.Controls.ComboBoxItem]::new()
                $item.Content    = $t
                $ui["cmbTable"].Items.Add($item)
            }
            $ui["cmbTable"].IsEnabled = $true

            # Store the auth context so queries can use it
            $state.AuthContext = $ctx

            $ui["statusDot"].Fill = [System.Windows.Media.BrushConverter]::new().ConvertFromString($script:Colors.Success)
            $ui["lblConnStatus"].Text = "Connected"
            & $setStatus "Found $($tables.Count) table(s). Select one and query." $script:Colors.Success
        }
        catch {
            $ui["statusDot"].Fill = [System.Windows.Media.BrushConverter]::new().ConvertFromString($script:Colors.Danger)
            $ui["lblConnStatus"].Text = "Error"
            & $setStatus "Failed to load tables: $_" $script:Colors.Danger
        }
    })

    # ─── Query / Apply Filters ───
    $ui["btnQuery"].Add_Click({
        # Mode 1: Azure Table Storage query
        if ($state.AuthContext) {
            $tableItem = $ui["cmbTable"].SelectedItem
            if (-not $tableItem) { & $setStatus "Select a table first." $script:Colors.Warning; return }
            $tableName = "$($tableItem.Content)".Trim()

            & $setStatus "Querying table '$tableName'..." $script:Colors.Info

            # Build OData filter
            $filters = @()
            $dateFrom = $ui["dpFrom"].SelectedDate
            $dateTo   = $ui["dpTo"].SelectedDate
            if ($dateFrom) { $filters += "Date ge '$($dateFrom.ToString('yyyy-MM-dd'))'" }
            if ($dateTo)   { $filters += "Date le '$($dateTo.ToString('yyyy-MM-dd'))'"   }

            $actionSel = $ui["cmbAction"].SelectedItem.Content
            if ($actionSel -and $actionSel -ne "All") { $filters += "Action eq '$actionSel'" }

            $protoSel = $ui["cmbProtocol"].SelectedItem.Content
            if ($protoSel -and $protoSel -ne "All") { $filters += "Protocol eq '$protoSel'" }

            $dirSel = $ui["cmbDirection"].SelectedItem.Content
            if ($dirSel -and $dirSel -ne "All") { $filters += "Direction eq '$dirSel'" }

            $odata = $filters -join " and "

            try {
                $entities = Get-StorageTableEntities -AuthContext $state.AuthContext -TableName $tableName -Filter $odata
                $state.RawData = $entities

                & $applyFilters
                & $setStatus "Loaded $($state.RawData.Count) records." $script:Colors.Info

                # Auto-trigger IP owner lookup
                & $startIPLookup
            }
            catch {
                & $setStatus "Query failed: $_" $script:Colors.Danger
            }
            return
        }

        # Mode 2: Apply filters to already-loaded data (CSV or previous query)
        if ($state.RawData.Count -gt 0) {
            & $applyFilters
            & $setStatus "Filters applied. Showing $($state.FilteredData.Count) of $($state.RawData.Count) records." $script:Colors.Success
            return
        }

        & $setStatus "Load a CSV file or connect to Azure first." $script:Colors.Warning
    })

    # ─── Clear filters ───
    $ui["btnClear"].Add_Click({
        $ui["dpFrom"].SelectedDate = $null
        $ui["dpTo"].SelectedDate   = $null
        $ui["cmbAction"].SelectedIndex    = 0
        $ui["cmbProtocol"].SelectedIndex  = 0
        $ui["cmbDirection"].SelectedIndex = 0
        $ui["txtIPFilter"].Text      = ""
        $ui["txtPortFilter"].Text    = ""
        $ui["txtProcessFilter"].Text = ""
        & $setStatus "Filters cleared." $script:Colors.SubText
    })

    # ─── Load CSV (from FirewallLogCollector output) ───
    $ui["btnLoadCsv"].Add_Click({
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter   = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
        $dlg.Title    = "Open Firewall Log CSV"
        $dlg.Multiselect = $true
        # Default to the collector output directory if it exists
        $defaultCsvDir = "C:\ProgramData\FirewallLogCollector\Logs"
        if (Test-Path $defaultCsvDir) { $dlg.InitialDirectory = $defaultCsvDir }

        if ($dlg.ShowDialog() -eq "OK") {
            & $setStatus "Loading CSV file(s)..." $script:Colors.Info
            try {
                $allRows = @()
                foreach ($file in $dlg.FileNames) {
                    $csvData = Import-Csv -Path $file -Encoding UTF8
                    foreach ($row in $csvData) {
                        # Normalise property names and add IP owner placeholders
                        $row | Add-Member -NotePropertyName "SrcIPOwner"    -NotePropertyValue "" -Force -ErrorAction SilentlyContinue
                        $row | Add-Member -NotePropertyName "DstIPOwner"    -NotePropertyValue "" -Force -ErrorAction SilentlyContinue
                        $row | Add-Member -NotePropertyName "DstIPLocation" -NotePropertyValue "" -Force -ErrorAction SilentlyContinue
                        # Ensure ComputerName exists
                        if (-not $row.PSObject.Properties['ComputerName']) {
                            $row | Add-Member -NotePropertyName "ComputerName" -NotePropertyValue "" -Force
                        }
                        # Ensure ProcessName, ProcessId, RuleName, EventAppName exist
                        if (-not $row.PSObject.Properties['ProcessName']) {
                            $row | Add-Member -NotePropertyName "ProcessName" -NotePropertyValue "Unknown" -Force
                        }
                        if (-not $row.PSObject.Properties['ProcessId']) {
                            $row | Add-Member -NotePropertyName "ProcessId" -NotePropertyValue 0 -Force
                        }
                        if (-not $row.PSObject.Properties['RuleName']) {
                            $row | Add-Member -NotePropertyName "RuleName" -NotePropertyValue "" -Force
                        }
                        if (-not $row.PSObject.Properties['EventAppName']) {
                            $row | Add-Member -NotePropertyName "EventAppName" -NotePropertyValue "" -Force
                        }
                        if (-not $row.PSObject.Properties['DstServiceTag']) {
                            $row | Add-Member -NotePropertyName "DstServiceTag" -NotePropertyValue "" -Force
                        }
                    }
                    $allRows += $csvData
                }

                $state.RawData = $allRows

                & $applyFilters

                $fileCount = $dlg.FileNames.Count
                $ui["statusDot"].Fill = [System.Windows.Media.BrushConverter]::new().ConvertFromString($script:Colors.Success)
                $ui["lblConnStatus"].Text = "CSV Loaded"
                & $setStatus "Loaded $($allRows.Count) records from $fileCount CSV file(s). Running IP & Service Tag lookups..." $script:Colors.Success

                # Auto-trigger IP Owner and Service Tag lookups
                & $startIPLookup
                & $startServiceTagLookup
            }
            catch {
                & $setStatus "Failed to load CSV: $_" $script:Colors.Danger
            }
        }
    })

    # ─── Export CSV ───
    $ui["btnExportCsv"].Add_Click({
        if ($state.FilteredData.Count -eq 0) { & $setStatus "No data to export." $script:Colors.Warning; return }

        $dlg = New-Object System.Windows.Forms.SaveFileDialog
        $dlg.Filter   = "CSV Files (*.csv)|*.csv"
        $dlg.FileName = "FirewallLogs_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        if ($dlg.ShowDialog() -eq "OK") {
            try {
                $state.FilteredData | Select-Object Date,Time,Action,Protocol,SrcIP,DstIP,SrcPort,DstPort,Direction,Size,ProcessName,ProcessId,RuleName,EventAppName,SrcIPOwner,DstIPOwner,DstIPLocation,DstServiceTag,ComputerName |
                    Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8
                & $setStatus "Exported $($state.FilteredData.Count) records to $($dlg.FileName)" $script:Colors.Success
            }
            catch { & $setStatus "Export failed: $_" $script:Colors.Danger }
        }
    })

    # ─── IP Owner Lookup (reusable scriptblock) ───
    $startIPLookup = {
        if ($state.FilteredData.Count -eq 0) { return }

        # Collect unique IPs
        $allIPs = @()
        foreach ($row in $state.FilteredData) {
            if ($row.SrcIP -and $row.SrcIP -ne '-') { $allIPs += $row.SrcIP }
            if ($row.DstIP -and $row.DstIP -ne '-') { $allIPs += $row.DstIP }
        }
        $uniqueIPs = @($allIPs | Select-Object -Unique)
        if ($uniqueIPs.Count -eq 0) { return }

        $ui["btnLookupIPs"].IsEnabled = $false
        & $setStatus "Looking up $($uniqueIPs.Count) unique IP(s) in background..." $script:Colors.Info

        # Store references at script scope so the timer tick can reach them
        $script:_ipLookupUI    = $ui
        $script:_ipLookupState = $state
        $script:_ipSetStatus   = $setStatus

        # Run lookup in a background runspace so the UI stays responsive
        $script:_ipRunspace = [runspacefactory]::CreateRunspace()
        $script:_ipRunspace.Open()

        $script:_ipPS = [powershell]::Create().AddScript({
            param($IPs, $ModulePath)
            . $ModulePath   # load IPOwnerLookup.ps1 into this runspace
            $results = Resolve-IPOwners -IPAddresses $IPs
            return $results
        }).AddArgument($uniqueIPs).AddArgument((Join-Path $script:ModuleDir "IPOwnerLookup.ps1"))

        $script:_ipPS.Runspace = $script:_ipRunspace
        $script:_ipHandle = $script:_ipPS.BeginInvoke()

        # Poll for completion via DispatcherTimer (keeps UI alive)
        $script:_ipTimer = New-Object System.Windows.Threading.DispatcherTimer
        $script:_ipTimer.Interval = [TimeSpan]::FromMilliseconds(500)
        $script:_ipTimer.Add_Tick({
            if ($script:_ipHandle.IsCompleted) {
                $script:_ipTimer.Stop()

                try {
                    $rawResult = $script:_ipPS.EndInvoke($script:_ipHandle)

                    # EndInvoke returns PSDataCollection — unwrap to the actual hashtable
                    $owners = $null
                    foreach ($item in $rawResult) {
                        if ($item -is [hashtable]) { $owners = $item; break }
                    }
                    if (-not $owners) { $owners = @{} }

                    # Filter out private/reserved entries for the count
                    $publicCount = ($owners.Values | Where-Object { $_.Status -eq 'success' }).Count

                    # Merge results into state
                    foreach ($k in $owners.Keys) { $script:_ipLookupState.IPOwners[$k] = $owners[$k] }

                    # Re-attach to grid data
                    foreach ($row in $script:_ipLookupState.FilteredData) {
                        $srcI = if ($script:_ipLookupState.IPOwners.ContainsKey($row.SrcIP)) { $script:_ipLookupState.IPOwners[$row.SrcIP] } else { $null }
                        $dstI = if ($script:_ipLookupState.IPOwners.ContainsKey($row.DstIP)) { $script:_ipLookupState.IPOwners[$row.DstIP] } else { $null }
                        $row | Add-Member -NotePropertyName "SrcIPOwner"    -NotePropertyValue $(if ($srcI) { $srcI.Name } else { "" }) -Force
                        $row | Add-Member -NotePropertyName "DstIPOwner"    -NotePropertyValue $(if ($dstI) { $dstI.Name } else { "" }) -Force
                        $row | Add-Member -NotePropertyName "DstIPLocation" -NotePropertyValue $(if ($dstI -and $dstI.City -ne 'N/A') { "$($dstI.City), $($dstI.Country)" } else { "" }) -Force
                        $row | Add-Member -NotePropertyName "DstServiceTag" -NotePropertyValue $(if ($script:_ipLookupState.ServiceTags.ContainsKey($row.DstIP)) { $script:_ipLookupState.ServiceTags[$row.DstIP] } else { "" }) -Force
                    }

                    $script:_ipLookupUI["dgData"].Items.Refresh()
                    & $script:_ipSetStatus "IP lookup complete. Resolved $publicCount public address(es) ($($owners.Count) total unique)." $script:Colors.Success
                }
                catch {
                    & $script:_ipSetStatus "IP lookup failed: $_" $script:Colors.Danger
                }
                finally {
                    $script:_ipLookupUI["btnLookupIPs"].IsEnabled = $true
                    $script:_ipPS.Dispose()
                    $script:_ipRunspace.Dispose()
                }
            }
        })
        $script:_ipTimer.Start()
    }

    $ui["btnLookupIPs"].Add_Click({ & $startIPLookup })

    # ─── Service Tag Lookup ───
    $startServiceTagLookup = {
        if ($state.FilteredData.Count -eq 0) { & $setStatus "No data loaded." $script:Colors.Warning; return }

        # Collect unique dest IPs
        $dstIPs = @()
        foreach ($row in $state.FilteredData) {
            if ($row.DstIP -and $row.DstIP -ne '-') { $dstIPs += $row.DstIP }
        }
        $uniqueDstIPs = @($dstIPs | Select-Object -Unique)
        if ($uniqueDstIPs.Count -eq 0) { & $setStatus "No destination IPs to look up." $script:Colors.Warning; return }

        $ui["btnServiceTags"].IsEnabled = $false
        & $setStatus "Downloading Microsoft Service Tags and resolving $($uniqueDstIPs.Count) dest IP(s)..." $script:Colors.Info

        $script:_stLookupUI    = $ui
        $script:_stLookupState = $state
        $script:_stSetStatus   = $setStatus

        # Run in background runspace
        $script:_stRunspace = [runspacefactory]::CreateRunspace()
        $script:_stRunspace.Open()

        $script:_stPS = [powershell]::Create().AddScript({
            param($IPs, $ModulePath)
            . $ModulePath   # load ServiceTagLookup.ps1
            $downloaded = Update-ServiceTagCache
            if (-not $downloaded) { return @{ Error = "Failed to download Service Tags" } }
            $indexed = Initialize-ServiceTagIndex
            if (-not $indexed) { return @{ Error = "Failed to parse Service Tags" } }
            $results = Resolve-ServiceTags -IPAddresses $IPs
            return $results
        }).AddArgument($uniqueDstIPs).AddArgument((Join-Path $script:ModuleDir "ServiceTagLookup.ps1"))

        $script:_stPS.Runspace = $script:_stRunspace
        $script:_stHandle = $script:_stPS.BeginInvoke()

        $script:_stTimer = New-Object System.Windows.Threading.DispatcherTimer
        $script:_stTimer.Interval = [TimeSpan]::FromMilliseconds(500)
        $script:_stTimer.Add_Tick({
            if ($script:_stHandle.IsCompleted) {
                $script:_stTimer.Stop()

                try {
                    $rawResult = $script:_stPS.EndInvoke($script:_stHandle)

                    $tags = $null
                    foreach ($item in $rawResult) {
                        if ($item -is [hashtable]) { $tags = $item; break }
                    }
                    if (-not $tags) { $tags = @{} }

                    # Check for error
                    if ($tags.ContainsKey('Error')) {
                        & $script:_stSetStatus "Service Tag lookup failed: $($tags.Error)" $script:Colors.Danger
                        return
                    }

                    # Merge into state
                    foreach ($k in $tags.Keys) { $script:_stLookupState.ServiceTags[$k] = $tags[$k] }

                    # Count how many IPs matched a service
                    $matchCount = ($tags.Values | Where-Object { $_ -ne '' }).Count

                    # Enrich grid rows
                    foreach ($row in $script:_stLookupState.FilteredData) {
                        $svcTag = if ($script:_stLookupState.ServiceTags.ContainsKey($row.DstIP)) { $script:_stLookupState.ServiceTags[$row.DstIP] } else { "" }
                        $row | Add-Member -NotePropertyName "DstServiceTag" -NotePropertyValue $svcTag -Force
                    }

                    $script:_stLookupUI["dgData"].Items.Refresh()
                    & $script:_stSetStatus "Service Tag lookup complete. $matchCount of $($tags.Count) dest IP(s) matched Microsoft services." $script:Colors.Success
                }
                catch {
                    & $script:_stSetStatus "Service Tag lookup failed: $_" $script:Colors.Danger
                }
                finally {
                    $script:_stLookupUI["btnServiceTags"].IsEnabled = $true
                    $script:_stPS.Dispose()
                    $script:_stRunspace.Dispose()
                }
            }
        })
        $script:_stTimer.Start()
    }

    $ui["btnServiceTags"].Add_Click({ & $startServiceTagLookup })

    # ═══ Menu bar wiring ═══
    # File menu
    $ui["mnuLoadCsv"].Add_Click({   $ui["btnLoadCsv"].RaiseEvent(  [System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) })
    $ui["mnuExportCsv"].Add_Click({ $ui["btnExportCsv"].RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) })
    $ui["mnuExit"].Add_Click({ $window.Close() })

    # Actions menu
    $ui["mnuQuery"].Add_Click({       $ui["btnQuery"].RaiseEvent(      [System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) })
    $ui["mnuClear"].Add_Click({       $ui["btnClear"].RaiseEvent(      [System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) })
    $ui["mnuLookupIPs"].Add_Click({   $ui["btnLookupIPs"].RaiseEvent(  [System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) })
    $ui["mnuServiceTags"].Add_Click({ $ui["btnServiceTags"].RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) })
    $ui["mnuConnectAzure"].Add_Click({ $ui["btnLogin"].RaiseEvent(    [System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) })

    $window.ShowDialog() | Out-Null
}
