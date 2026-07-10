#Requires -Version 5.1
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore

# --- 1. AppData Configuration Management ---
$AppName = "Clientreport"
$ConfigDir = Join-Path -Path $env:APPDATA -ChildPath $AppName
$ConfigFile = Join-Path -Path $ConfigDir -ChildPath "config.json"

if (-not (Test-Path -Path $ConfigDir)) {
    New-Item -ItemType Directory -Path $ConfigDir | Out-Null
}

function Get-AppConfig {
    if (Test-Path -Path $ConfigFile) {
        try {
            return Get-Content -Path $ConfigFile -Raw | ConvertFrom-Json
        }
        catch { }
    }
    return [PSCustomObject]@{ Tenants = @(); ModuleImported = $false }
}

function Save-AppConfig ($Config) {
    $Config | ConvertTo-Json -Depth 2 | Set-Content -Path $ConfigFile -Encoding UTF8
}

function Get-CacheFilePath ($TenantId) {
    if ([string]::IsNullOrWhiteSpace($TenantId)) { return $null }
    return Join-Path -Path $ConfigDir -ChildPath "$($TenantId -replace '[^a-zA-Z0-9-]', '')_cache.json"
}

$AppConfig = Get-AppConfig

if ($null -eq $AppConfig.ModuleImported) {
    $AppConfig | Add-Member -NotePropertyName ModuleImported -NotePropertyValue $false -Force
}

if (-not $AppConfig.ModuleImported) {
    try {
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
        $AppConfig.ModuleImported = $true
        Save-AppConfig $AppConfig
    }
    catch {
        Write-Warning "Failed to import Microsoft.Graph.Authentication: $($_.Exception.Message)"
    }
}

# --- 2. XAML Definition ---
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Clientreport" Height="750" Width="1480" Background="#1E1E1E"
        WindowStartupLocation="CenterScreen" WindowStyle="None">

    <WindowChrome.WindowChrome>
        <WindowChrome CaptionHeight="32" ResizeBorderThickness="6" GlassFrameThickness="0" UseAeroCaptionButtons="False"/>
    </WindowChrome.WindowChrome>

    <Window.Resources>
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="#E0E0E0"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
        </Style>
        <Style TargetType="Label">
            <Setter Property="Foreground" Value="#E0E0E0"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
        </Style>
        
        <Style TargetType="Button">
            <Setter Property="Background" Value="#333333"/>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="BorderBrush" Value="#555555"/>
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="Margin" Value="5,0"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#505050"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Background" Value="#252525"/>
                    <Setter Property="Foreground" Value="#777777"/>
                    <Setter Property="BorderBrush" Value="#333333"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style x:Key="ChromeButtonStyle" TargetType="Button">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="Width" Value="45"/>
            <Setter Property="Height" Value="32"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#3E3E40"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style x:Key="CloseButtonStyle" TargetType="Button" BasedOn="{StaticResource ChromeButtonStyle}">
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#E81123"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <SolidColorBrush x:Key="StandardScrollBarTrack" Color="#1E1E1E"/>
        <SolidColorBrush x:Key="StandardScrollBarThumb" Color="#424242"/>
        <SolidColorBrush x:Key="StandardScrollBarThumbHover" Color="#5E5E5E"/>
        <Style TargetType="ScrollBar">
            <Setter Property="Background" Value="{StaticResource StandardScrollBarTrack}"/>
            <Setter Property="Foreground" Value="{StaticResource StandardScrollBarThumb}"/>
            <Setter Property="Width" Value="14"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ScrollBar">
                        <Grid Background="{TemplateBinding Background}">
                            <Track x:Name="PART_Track" IsDirectionReversed="true">
                                <Track.Thumb>
                                    <Thumb Background="{TemplateBinding Foreground}">
                                        <Thumb.Template>
                                            <ControlTemplate TargetType="Thumb">
                                                <Border Background="{TemplateBinding Background}" CornerRadius="6" Margin="2"/>
                                            </ControlTemplate>
                                        </Thumb.Template>
                                    </Thumb>
                                </Track.Thumb>
                            </Track>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <ControlTemplate x:Key="ComboBoxToggleButton" TargetType="ToggleButton">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition />
                    <ColumnDefinition Width="20" />
                </Grid.ColumnDefinitions>
                <Border x:Name="Border" Grid.ColumnSpan="2" CornerRadius="4" Background="#2D2D2D" BorderBrush="#555555" BorderThickness="1" />
                <Border Grid.Column="0" CornerRadius="4,0,0,4" Margin="1" Background="#2D2D2D" />
                <Path x:Name="Arrow" Grid.Column="1" Fill="#E0E0E0" HorizontalAlignment="Center" VerticalAlignment="Center" Data="M 0 0 L 4 4 L 8 0 Z"/>
            </Grid>
            <ControlTemplate.Triggers>
                <Trigger Property="IsMouseOver" Value="true">
                    <Setter TargetName="Border" Property="Background" Value="#3D3D3D" />
                </Trigger>
                <Trigger Property="IsChecked" Value="true">
                    <Setter TargetName="Border" Property="Background" Value="#3D3D3D" />
                </Trigger>
            </ControlTemplate.Triggers>
        </ControlTemplate>
        <Style TargetType="ComboBox">
            <Setter Property="Foreground" Value="#E0E0E0"/>
            <Setter Property="Background" Value="#2D2D2D"/>
            <Setter Property="BorderBrush" Value="#555555"/>
            <Setter Property="Height" Value="28"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBox">
                        <Grid>
                            <ToggleButton Name="ToggleButton" Template="{StaticResource ComboBoxToggleButton}" Grid.Column="2" Focusable="false" IsChecked="{Binding Path=IsDropDownOpen,Mode=TwoWay,RelativeSource={RelativeSource TemplatedParent}}" ClickMode="Press"/>
                            <ContentPresenter Name="ContentSite" IsHitTestVisible="False" Content="{TemplateBinding SelectionBoxItem}" ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}" ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}" Margin="8,3,23,3" VerticalAlignment="Center" HorizontalAlignment="Left" />
                            <TextBox x:Name="PART_EditableTextBox" 
                                     Style="{x:Null}" 
                                     BorderThickness="0" 
                                     Background="Transparent" 
                                     Foreground="#E0E0E0" 
                                     CaretBrush="#E0E0E0" 
                                     HorizontalAlignment="Stretch" 
                                     VerticalAlignment="Center" 
                                     Margin="8,0,23,0" 
                                     Focusable="True" 
                                     Visibility="Hidden" 
                                     IsReadOnly="{TemplateBinding IsReadOnly}"/>
                            <Popup Name="Popup" Placement="Bottom" IsOpen="{TemplateBinding IsDropDownOpen}" AllowsTransparency="True" Focusable="False" PopupAnimation="Slide">
                                <Grid Name="DropDown" SnapsToDevicePixels="True" MinWidth="{TemplateBinding ActualWidth}" MaxHeight="{TemplateBinding MaxDropDownHeight}">
                                    <Border x:Name="DropDownBorder" Background="#252525" BorderThickness="1" BorderBrush="#555555"/>
                                    <ScrollViewer Margin="1" SnapsToDevicePixels="True">
                                        <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Contained" />
                                    </ScrollViewer>
                                </Grid>
                            </Popup>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="HasItems" Value="false">
                                <Setter TargetName="DropDownBorder" Property="MinHeight" Value="95"/>
                            </Trigger>
                            <Trigger Property="IsGrouping" Value="true">
                                <Setter Property="ScrollViewer.CanContentScroll" Value="false"/>
                            </Trigger>
                            <Trigger Property="IsEditable" Value="true">
                                <Setter Property="IsTabStop" Value="false"/>
                                <Setter TargetName="PART_EditableTextBox" Property="Visibility" Value="Visible"/>
                                <Setter TargetName="ContentSite" Property="Visibility" Value="Hidden"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="ComboBoxItem">
            <Setter Property="Foreground" Value="#E0E0E0"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Padding" Value="5"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBoxItem">
                        <Border Name="Border" Padding="{TemplateBinding Padding}" Background="{TemplateBinding Background}">
                            <ContentPresenter />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsHighlighted" Value="true">
                                <Setter TargetName="Border" Property="Background" Value="#007ACC"/>
                                <Setter Property="Foreground" Value="#FFFFFF"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="DataGrid">
            <Setter Property="Background" Value="#1E1E1E"/>
            <Setter Property="Foreground" Value="#E0E0E0"/>
            <Setter Property="RowBackground" Value="#252525"/>
            <Setter Property="AlternatingRowBackground" Value="#2A2A2A"/>
            <Setter Property="GridLinesVisibility" Value="None"/>
            <Setter Property="BorderBrush" Value="#333333"/>
            <Setter Property="HeadersVisibility" Value="Column"/>
        </Style>
        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="#333333"/>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="BorderThickness" Value="0,0,1,1"/>
            <Setter Property="BorderBrush" Value="#252525"/>
        </Style>
        <Style TargetType="DataGridCell">
            <Setter Property="Padding" Value="5"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="DataGridCell">
                        <Border Padding="{TemplateBinding Padding}" Background="{TemplateBinding Background}">
                            <ContentPresenter VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#007ACC"/>
                    <Setter Property="Foreground" Value="White"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style x:Key="DarkTooltip" TargetType="ToolTip">
            <Setter Property="Background" Value="#2D2D2D"/>
            <Setter Property="Foreground" Value="#E0E0E0"/>
            <Setter Property="BorderBrush" Value="#555555"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="10,8"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="FontSize" Value="12"/>
        </Style>
    </Window.Resources>

    <Border BorderBrush="#333333" BorderThickness="1">
        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="32"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <Grid Grid.Row="0" Background="#2D2D2D">
                <StackPanel Orientation="Horizontal" VerticalAlignment="Center" Margin="10,0,0,0">
                    <TextBlock Text="Clientreport" FontWeight="SemiBold" FontSize="12"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" WindowChrome.IsHitTestVisibleInChrome="True">
                    <Button x:Name="btnMinimize" Content="&#151;" Style="{StaticResource ChromeButtonStyle}" FontSize="10"/>
                    <Button x:Name="btnMaximize" Content="&#9634;" Style="{StaticResource ChromeButtonStyle}" FontSize="14"/>
                    <Button x:Name="btnClose"    Content="&#10005;" Style="{StaticResource CloseButtonStyle}" FontSize="11"/>
                </StackPanel>
            </Grid>

            <Grid Grid.Row="1" Margin="15,15,15,10">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="220"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="220"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <Label    Grid.Column="0" Content="Tenant:" VerticalAlignment="Center" Margin="0,0,5,0"/>
                <ComboBox x:Name="cboTenant" Grid.Column="1" IsEditable="True" VerticalAlignment="Center"/>
                <Button   x:Name="btnConnect" Grid.Column="2" Content="Connect" Width="90"/>

                <Label   Grid.Column="4" Content="Search:" VerticalAlignment="Center" Margin="0,0,5,0"/>
                <TextBox x:Name="txtSearch" Grid.Column="5" Height="28" Background="#2D2D2D" Foreground="#FFFFFF" BorderBrush="#555555" Padding="5,3" VerticalAlignment="Center"/>

                <StackPanel Grid.Column="6" Orientation="Vertical" VerticalAlignment="Center" Margin="5,0">
                    <Button x:Name="btnFetch" Content="Generate Report" Width="130" IsEnabled="False" Background="#007ACC" Height="28"/>
                    <TextBlock x:Name="txtCacheStatus" Text="No cache found." Foreground="#888888" HorizontalAlignment="Center" Margin="0,3,0,0" FontSize="10"/>
                </StackPanel>
                <Button x:Name="btnExport" Grid.Column="7" Content="Export CSV"      Width="100" IsEnabled="False" VerticalAlignment="Center"/>
            </Grid>

            <DataGrid x:Name="dgResults" Grid.Row="2" Margin="15,5,15,5"
                      AutoGenerateColumns="False" CanUserSortColumns="True" IsReadOnly="True"
                      ScrollViewer.CanContentScroll="True" ScrollViewer.VerticalScrollBarVisibility="Auto">
                <DataGrid.Columns>
                    <DataGridTextColumn Header="Display Name"        Binding="{Binding DisplayName}"       Width="1.2*"/>
                    <DataGridTextColumn Header="Last Logged On User" Binding="{Binding LastLoggedOnUser}"  Width="1.6*"/>
                    <DataGridTextColumn Header="Last Logon Date"     Binding="{Binding LastLogOnDateTime}" Width="1.1*"/>

                    <DataGridTemplateColumn Header="Logon History" Width="1.8*" SortMemberPath="AllLoggedOnUsers">
                        <DataGridTemplateColumn.CellTemplate>
                            <DataTemplate>
                                <TextBlock Text="{Binding AllLoggedOnUsers}"
                                           TextTrimming="CharacterEllipsis"
                                           VerticalAlignment="Center"
                                           Margin="5,0">
                                    <TextBlock.ToolTip>
                                        <ToolTip Style="{StaticResource DarkTooltip}"
                                                 Content="{Binding LogonHistoryTooltip}"/>
                                    </TextBlock.ToolTip>
                                </TextBlock>
                            </DataTemplate>
                        </DataGridTemplateColumn.CellTemplate>
                    </DataGridTemplateColumn>

                    <DataGridTextColumn Header="Model"          Binding="{Binding ModelName}"          Width="1.2*"/>
                    <DataGridTextColumn Header="Serial Number"  Binding="{Binding SerialNumber}"       Width="1*"/>
                    <DataGridTextColumn Header="Wi-Fi MAC"      Binding="{Binding WiFiMacAddress}"     Width="1.1*"/>
                    <DataGridTextColumn Header="Ethernet MAC"   Binding="{Binding EthernetMacAddress}" Width="1.1*"/>
                    <DataGridTextColumn Header="Days Sync"      Binding="{Binding DaysSinceLastSync}"  Width="0.8*"/>
                    <DataGridTextColumn Header="OS Version"     Binding="{Binding OSVersion}"          Width="0.9*"/>
                </DataGrid.Columns>
            </DataGrid>

            <Grid Grid.Row="3" Margin="15,10,15,15">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="200"/>
                </Grid.ColumnDefinitions>
                <TextBlock    x:Name="txtStatus" Grid.Column="0" Text="Ready" VerticalAlignment="Center" FontSize="13"/>
                <ProgressBar  x:Name="pbStatus"  Grid.Column="1" Height="18" Minimum="0" Maximum="100" Value="0"
                              Background="#333333" Foreground="#007ACC" BorderThickness="0" Visibility="Collapsed"/>
            </Grid>
        </Grid>
    </Border>
</Window>
"@

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
try {
    $Window = [Windows.Markup.XamlReader]::Load($reader)
}
catch {
    Write-Warning "Failed to load XAML. $($_.Exception.Message)"
    exit
}

# --- 3. UI Element Mapping ---
$cboTenant = $Window.FindName("cboTenant")
$btnConnect = $Window.FindName("btnConnect")
$txtCacheStatus = $Window.FindName("txtCacheStatus")
$btnFetch = $Window.FindName("btnFetch")
$btnExport = $Window.FindName("btnExport")
$txtSearch = $Window.FindName("txtSearch")
$dgResults = $Window.FindName("dgResults")
$txtStatus = $Window.FindName("txtStatus")
$pbStatus = $Window.FindName("pbStatus")
$btnMinimize = $Window.FindName("btnMinimize")
$btnMaximize = $Window.FindName("btnMaximize")
$btnClose = $Window.FindName("btnClose")

# --- 4. State Management & Cache Checking ---
$Script:IsConnected = $false
$Script:ConnectedTenant = ""
$Script:ReportData = @()

function Update-CacheUI {
    $tenantId = if ($null -ne $cboTenant.SelectedItem) { $cboTenant.SelectedItem.ToString() } else { $cboTenant.Text }
    $tenantId = $tenantId.Trim()
    
    if ($Script:IsConnected -and -not [string]::IsNullOrEmpty($tenantId) -and $tenantId -ne $Script:ConnectedTenant) {
        try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch { }
        $Script:IsConnected = $false
        $Script:ConnectedTenant = ""
        $btnConnect.Content = "Connect"
        $btnFetch.IsEnabled = $false
        $txtStatus.Text = "Tenant changed. Connection closed."
        $pbStatus.Visibility = [System.Windows.Visibility]::Collapsed
        $pbStatus.Value = 0
    }

    $cacheFile = Get-CacheFilePath $tenantId

    if ($cacheFile -and (Test-Path -Path $cacheFile)) {
        $lastUpdate = (Get-Item $cacheFile).LastWriteTime.ToString("dd.MM.yyyy HH:mm")
        $txtCacheStatus.Text = "Cached data timestamp: $lastUpdate"
        $txtCacheStatus.Foreground = "#00A000"
        
        try {
            $txtStatus.Text = "Auto-loading cached data..."
            $Window.Dispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::Render)
            
            $cachedJson = Get-Content -Path $cacheFile -Raw | ConvertFrom-Json
            
            $Script:ReportData = @()
            foreach ($item in $cachedJson) {
                $Script:ReportData += [PSCustomObject]@{
                    DisplayName         = $item.DisplayName
                    LastLoggedOnUser    = $item.LastLoggedOnUser
                    LastLogOnDateTime   = $item.LastLogOnDateTime
                    AllLoggedOnUsers    = $item.AllLoggedOnUsers
                    LogonHistoryTooltip = $item.LogonHistoryTooltip
                    ModelName           = $item.ModelName
                    SerialNumber        = $item.SerialNumber
                    WiFiMacAddress      = $item.WiFiMacAddress
                    EthernetMacAddress  = $item.EthernetMacAddress
                    DaysSinceLastSync   = $item.DaysSinceLastSync
                    OSVersion           = $item.OSVersion
                }
            }
            
            $dgResults.ItemsSource = $Script:ReportData
            $btnExport.IsEnabled = $true
            
            # Unconditionally update the status message with exact cache timestamp
            $txtStatus.Text = "Loaded $($Script:ReportData.Count) records from offline cache ($lastUpdate)."
        }
        catch {
            $txtStatus.Text = "Failed to auto-load cache: $($_.Exception.Message)"
        }
    }
    else {
        $txtCacheStatus.Text = "No cache found."
        $txtCacheStatus.Foreground = "#888888"
        $dgResults.ItemsSource = $null
        $btnExport.IsEnabled = $false
        $Script:ReportData = @()
        if (-not $Script:IsConnected) {
            $txtStatus.Text = "Ready"
        }
    }
}

$cboTenant.Add_SelectionChanged({ 
        $Window.Dispatcher.InvokeAsync({ Update-CacheUI }, [System.Windows.Threading.DispatcherPriority]::Render) | Out-Null
    })
$cboTenant.Add_KeyUp({ Update-CacheUI })

$AppConfig.Tenants | ForEach-Object { $cboTenant.Items.Add($_) | Out-Null }
if ($AppConfig.Tenants.Count -gt 0) { 
    $cboTenant.SelectedIndex = 0 
    Update-CacheUI
}

$syncHash = [hashtable]::Synchronized(@{
        TenantId        = ""
        StatusText      = "Ready"
        PercentComplete = 0
        IsRunning       = $false
        IsFinished      = $false
        Error           = $null
        Results         = $null
    })

# --- 5. Background Runspace ScriptBlock ---
$backgroundJob = {
    param($syncHash)

    function Format-MacAddress ($raw) {
        if ([string]::IsNullOrEmpty($raw)) { return "N/A" }
        return $raw.ToUpper()
    }

    function Format-LogOnDate ($raw) {
        if ([string]::IsNullOrEmpty($raw)) { return "N/A" }
        try { 
            return ([datetime]$raw).ToLocalTime().ToString("dd.MM.yyyy HH:mm")
        }
        catch { return $raw }
    }

    try {
        $syncHash.StatusText = "Loading authentication parameters..."
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

        $syncHash.StatusText = "Establishing graph session..."
        Connect-MgGraph -TenantId $syncHash.TenantId -Scopes "DeviceManagementManagedDevices.Read.All", "User.Read.All" -NoWelcome -ClientTimeout 600

        $now = Get-Date
        
        $ManagedDevices = [System.Collections.Generic.List[object]]::new()
        $uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices"
        
        $syncHash.StatusText = "Fetching endpoints..."
        $syncHash.PercentComplete = 5

        do {
            $response = Invoke-MgGraphRequest -Method Get -Uri $uri -OutputType PSObject
            if ($response.value) {
                $ManagedDevices.AddRange($response.value)
            }
            $uri = $response.'@odata.nextLink'
            if ($uri) {
                $syncHash.StatusText = "Fetching endpoints... ($($ManagedDevices.Count) devices found so far)"
            }
        } while ($uri)

        if ($ManagedDevices.Count -eq 0) {
            $syncHash.StatusText = "No endpoints returned."
            $syncHash.PercentComplete = 100
            $syncHash.IsFinished = $true
            return
        }

        $ManagedDevices = $ManagedDevices | Sort-Object -Property deviceName

        $deviceHardwareMap = @{}
        $batchSize = 20
        $totalHwBatches = [math]::Ceiling($ManagedDevices.Count / $batchSize)
        
        for ($i = 0; $i -lt $ManagedDevices.Count; $i += $batchSize) {
            $batchNum = [math]::Floor($i / $batchSize) + 1
            $syncHash.StatusText = "Fetching hardware info - batch $batchNum of $totalHwBatches"
            $syncHash.PercentComplete = 10 + ((($batchNum - 1) / $totalHwBatches) * 35)

            $endIdx = [math]::Min($i + $batchSize - 1, $ManagedDevices.Count - 1)
            $chunk = $ManagedDevices[$i..$endIdx]

            $batchRequests = for ($j = 0; $j -lt $chunk.Count; $j++) {
                @{ 
                    id     = "$j"
                    method = "GET"
                    url    = "/deviceManagement/managedDevices/$($chunk[$j].id)?`$select=id,hardwareInformation,ethernetMacAddress" 
                }
            }

            $batchBody = @{ requests = @($batchRequests) } | ConvertTo-Json -Depth 10
            $batchResponse = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/`$batch" -Body $batchBody -OutputType PSObject -ErrorAction Stop

            foreach ($resp in $batchResponse.responses) {
                if ($resp.status -eq 200) {
                    $devId = $resp.body.id
                    $ethMac = $resp.body.ethernetMacAddress
                    
                    if (-not $ethMac -and $resp.body.hardwareInformation -and $resp.body.hardwareInformation.ethernetMacAddress) {
                        $ethMac = $resp.body.hardwareInformation.ethernetMacAddress
                    }
                    $deviceHardwareMap[$devId] = $ethMac
                }
            }
        }

        $uniqueUserIds = $ManagedDevices |
        ForEach-Object { $_.usersLoggedOn | ForEach-Object { $_.userId } } |
        Where-Object { -not [string]::IsNullOrEmpty($_) } |
        Select-Object  -Unique

        $userUpnMap = @{}

        if ($uniqueUserIds.Count -gt 0) {
            $totalUpnBatches = [math]::Ceiling($uniqueUserIds.Count / $batchSize)

            for ($i = 0; $i -lt $uniqueUserIds.Count; $i += $batchSize) {
                $batchNum = [math]::Floor($i / $batchSize) + 1
                $syncHash.StatusText = "Resolving directory principals - batch $batchNum of $totalUpnBatches"
                $syncHash.PercentComplete = 45 + ((($batchNum - 1) / $totalUpnBatches) * 35)

                $endIdx = [math]::Min($i + $batchSize - 1, $uniqueUserIds.Count - 1)
                $chunk = $uniqueUserIds[$i..$endIdx]

                $batchRequests = for ($j = 0; $j -lt $chunk.Count; $j++) {
                    @{ id = "$j"; method = "GET"; url = "/users/$($chunk[$j])?`$select=userPrincipalName" }
                }

                $batchBody = @{ requests = @($batchRequests) } | ConvertTo-Json -Depth 10
                $batchResponse = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/`$batch" -Body $batchBody -OutputType PSObject -ErrorAction Stop

                foreach ($resp in $batchResponse.responses) {
                    $uid = $chunk[[int]$resp.id]
                    $userUpnMap[$uid] = if ($resp.status -eq 200) { $resp.body.userPrincipalName } else { "unknown ($uid)" }
                }
            }
        }

        $results = [System.Collections.Generic.List[object]]::new()
        $deviceCount = $ManagedDevices.Count
        $devIndex = 0

        foreach ($ManagedDevice in $ManagedDevices) {
            $devIndex++
            $syncHash.StatusText = "Compiling records: $devIndex / $deviceCount"
            $syncHash.PercentComplete = 80 + (($devIndex / $deviceCount) * 20)

            $syncDate = try { [datetime]$ManagedDevice.lastSyncDateTime } catch { $now }
            $timediff = New-TimeSpan -Start $syncDate -End $now
            
            $lastLogOn = $ManagedDevice.usersLoggedOn | Select-Object -Last 1

            $lastUserId = $lastLogOn.userId
            $lastUserUpn = if ([string]::IsNullOrEmpty($lastUserId)) { "N/A" } else { $userUpnMap[$lastUserId] }
            $lastDateFmt = Format-LogOnDate $lastLogOn.lastLogOnDateTime

            # Sort descending by lastLogOnDateTime before formatting
            $allUsersFormatted = $ManagedDevice.usersLoggedOn | 
            Sort-Object -Property lastLogOnDateTime -Descending | 
            ForEach-Object {
                $uid = $_.userId
                $dateStr = Format-LogOnDate $_.lastLogOnDateTime
                if (-not [string]::IsNullOrEmpty($uid)) {
                    $upn = if ($userUpnMap.ContainsKey($uid)) { $userUpnMap[$uid] } else { "unknown ($uid)" }
                    "$upn ($dateStr)"
                }
            } | Select-Object -Unique

            $allUsersFlat = if ($allUsersFormatted) { $allUsersFormatted -join "; " }    else { "N/A" }
            $allUsersTooltip = if ($allUsersFormatted) { $allUsersFormatted -join "`n" }    else { "N/A" }

            $wifiMac = Format-MacAddress $ManagedDevice.wiFiMacAddress
            $ethMac = Format-MacAddress $deviceHardwareMap[$ManagedDevice.id]

            $results.Add([PSCustomObject]@{
                    DisplayName         = $ManagedDevice.deviceName
                    LastLoggedOnUser    = $lastUserUpn
                    LastLogOnDateTime   = $lastDateFmt
                    AllLoggedOnUsers    = $allUsersFlat
                    LogonHistoryTooltip = $allUsersTooltip
                    ModelName           = $ManagedDevice.model
                    SerialNumber        = $ManagedDevice.serialNumber
                    WiFiMacAddress      = $wifiMac
                    EthernetMacAddress  = $ethMac
                    DaysSinceLastSync   = $timediff.Days
                    OSVersion           = $ManagedDevice.osVersion
                })
        }

        $syncHash.Results = $results
        $syncHash.StatusText = "Data fetching completed."
        $syncHash.PercentComplete = 100

    }
    catch {
        $syncHash.Error = $_.Exception.Message
        $syncHash.StatusText = "Engine error detected."
    }
    finally {
        $syncHash.IsFinished = $true
        $syncHash.IsRunning = $false
    }
}

# --- 6. Execution ---
function Invoke-ReportFetch {
    $btnFetch.IsEnabled = $false
    $btnConnect.IsEnabled = $false
    $btnExport.IsEnabled = $false
    $dgResults.ItemsSource = $null
    $txtSearch.Text = ""

    $syncHash.TenantId = $cboTenant.Text.Trim()
    $syncHash.StatusText = "Spawning worker thread..."
    $syncHash.PercentComplete = 0
    $syncHash.IsRunning = $true
    $syncHash.IsFinished = $false
    $syncHash.Error = $null
    $syncHash.Results = $null

    $pbStatus.Value = 0
    $pbStatus.Visibility = [System.Windows.Visibility]::Visible

    $Script:runspace = [runspacefactory]::CreateRunspace()
    $Script:runspace.ApartmentState = "MTA"
    $Script:runspace.ThreadOptions = "ReuseThread"
    $Script:runspace.Open()

    $Script:powerShell = [powershell]::Create()
    $Script:powerShell.Runspace = $Script:runspace
    $Script:powerShell.AddScript($backgroundJob).AddArgument($syncHash) | Out-Null
    $Script:powerShell.BeginInvoke() | Out-Null
    $timer.Start()
}

# --- 7. Timer ---
$timer = New-Object System.Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromMilliseconds(200)
$timer.Add_Tick({
        $txtStatus.Text = $syncHash.StatusText
        $pbStatus.Value = $syncHash.PercentComplete

        if ($syncHash.IsFinished) {
            $timer.Stop()
            $syncHash.IsFinished = $false

            $pbStatus.Visibility = [System.Windows.Visibility]::Collapsed
            $pbStatus.Value = 0

            $btnFetch.IsEnabled = $true
            $btnConnect.IsEnabled = $true
            Update-CacheUI

            if ($syncHash.Error) {
                [System.Windows.MessageBox]::Show($syncHash.Error, "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            }
            elseif ($syncHash.Results) {
                $Script:ReportData = $syncHash.Results
                $dgResults.ItemsSource = $Script:ReportData
                $btnExport.IsEnabled = $true
            
                $cacheFile = Get-CacheFilePath $cboTenant.Text.Trim()
                if ($cacheFile) {
                    try {
                        $Script:ReportData | ConvertTo-Json -Depth 10 | Set-Content -Path $cacheFile -Encoding UTF8
                        Update-CacheUI
                    }
                    catch {
                        Write-Warning "Failed to save cache file: $($_.Exception.Message)"
                    }
                }
            }

            if ($Script:runspace) {
                $Script:runspace.Dispose()
                $Script:powerShell.Dispose()
            }
        }
    })

# --- 8. Event Handlers ---
$btnMinimize.Add_Click({ $Window.WindowState = [System.Windows.WindowState]::Minimized })
$btnMaximize.Add_Click({
        if ($Window.WindowState -eq [System.Windows.WindowState]::Maximized) {
            $Window.WindowState = [System.Windows.WindowState]::Normal
        }
        else {
            $Window.WindowState = [System.Windows.WindowState]::Maximized
        }
    })
$btnClose.Add_Click({ $Window.Close() })

$txtSearch.Add_TextChanged({
        if (-not $Script:ReportData -or $Script:ReportData.Count -eq 0) { return }
        $query = $txtSearch.Text.Trim()

        if ([string]::IsNullOrEmpty($query)) {
            $dgResults.ItemsSource = $Script:ReportData
        }
        else {
            $filtered = [System.Collections.Generic.List[object]]::new()
            foreach ($row in $Script:ReportData) {
                if (($row.DisplayName -like "*$query*") -or
                    ($row.LastLoggedOnUser -like "*$query*") -or
                    ($row.AllLoggedOnUsers -like "*$query*") -or
                    ($row.ModelName -like "*$query*") -or
                    ($row.SerialNumber -like "*$query*") -or
                    ($row.WiFiMacAddress -like "*$query*") -or
                    ($row.EthernetMacAddress -like "*$query*") -or
                    ($row.OSVersion -like "*$query*")) {
                    $filtered.Add($row) | Out-Null
                }
            }
            $dgResults.ItemsSource = $filtered
        }
    })

$btnConnect.Add_Click({
        if ($Script:IsConnected) {
            try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch { }
            $Script:IsConnected = $false
            $Script:ConnectedTenant = ""
            $Script:ReportData = @()
            $dgResults.ItemsSource = $null
            $btnConnect.Content = "Connect"
            $btnFetch.IsEnabled = $false
            $btnExport.IsEnabled = $false
            $txtStatus.Text = "Disconnected."
            $pbStatus.Visibility = [System.Windows.Visibility]::Collapsed
            $pbStatus.Value = 0
            Update-CacheUI
        }
        else {
            $tenantId = $cboTenant.Text.Trim()
            if (-not $tenantId) {
                [System.Windows.MessageBox]::Show("Please supply a valid Tenant string.", "Input Check Failed", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }

            $txtStatus.Text = "Connecting..."
            $Window.Dispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::Render)

            try {
                Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
                Connect-MgGraph -TenantId $tenantId -Scopes "DeviceManagementManagedDevices.Read.All", "User.Read.All" -NoWelcome -ClientTimeout 600

                $Script:IsConnected = $true
                $Script:ConnectedTenant = $tenantId
                $btnConnect.Content = "Disconnect"
                $btnFetch.IsEnabled = $true
                $txtStatus.Text = "Session established. Fetching data..."

                if ($AppConfig.Tenants -notcontains $tenantId) {
                    $AppConfig.Tenants += $tenantId
                    Save-AppConfig $AppConfig
                    $cboTenant.Items.Add($tenantId) | Out-Null
                }

                Invoke-ReportFetch

            }
            catch {
                [System.Windows.MessageBox]::Show("Connection error: $($_.Exception.Message)", "Identity Exception", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                $txtStatus.Text = "Authentication aborted."
            }
        }
    })

$btnFetch.Add_Click({ Invoke-ReportFetch })

$btnExport.Add_Click({
        if (-not $dgResults.ItemsSource) { return }

        $dialog = New-Object Microsoft.Win32.SaveFileDialog
        $dialog.Filter = "CSV Files (*.csv)|*.csv"
        $dialog.FileName = "Clientreport_$(Get-Date -F yyyyMMdd_HHmmss).csv"
        $dialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')

        if ($dialog.ShowDialog() -eq $true) {
            try {
                $dgResults.ItemsSource | Export-Csv -Path $dialog.FileName -NoTypeInformation -Encoding UTF8
                $txtStatus.Text = "Written: $($dialog.FileName)"
            }
            catch {
                [System.Windows.MessageBox]::Show("Export error: $($_.Exception.Message)", "IO Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            }
        }
    })

$Window.Add_Closed({
        try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch { }
        if ($Script:runspace) { $Script:runspace.Dispose() }
        if ($Script:powerShell) { $Script:powerShell.Dispose() }
    })

# --- 9. Launch ---
$Window.ShowDialog() | Out-Null