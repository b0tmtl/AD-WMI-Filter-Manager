#Requires -Modules ActiveDirectory
<#
.SYNOPSIS
    WMI Filter Management and Testing Tool for Active Directory - Modern GUI Version (Optimized)
.DESCRIPTION
    This tool retrieves all WMI filters from a specified domain, displays their details
    including linked GPOs, and allows testing filters against specific computers.
    Features a modern WPF-based graphical user interface with optimized performance.
.EXAMPLE
    .\Get-WMIFilterToolGUI.ps1
.NOTES
    Name: PowerShell WMI Filter Tool
    Author: Younes Cheriti
    Version: 3.1 (GUI Version - Performance Optimized)
    Requires: Active Directory module and appropriate permissions
#>

[CmdletBinding()]
param()

# Add required assemblies
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

# Check for required module
if (!(Get-Module -ListAvailable -Name ActiveDirectory)) {
    [System.Windows.MessageBox]::Show(
        "Active Directory PowerShell module is not installed.`nPlease install RSAT or run this on a domain controller.",
        "Missing Prerequisites",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Error
    )
    exit 1
}

Import-Module ActiveDirectory -ErrorAction Stop

# Define XAML for the modern GUI
$xaml = @'
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="WMI Filter Management Tool" 
    Height="800" Width="1200"
    WindowStartupLocation="CenterScreen"
    Background="#1e1e1e">
    
    <Window.Resources>
        <!-- Modern Color Scheme -->
        <SolidColorBrush x:Key="PrimaryColor" Color="#0078D4"/>
        <SolidColorBrush x:Key="SecondaryColor" Color="#106EBE"/>
        <SolidColorBrush x:Key="AccentColor" Color="#40E0D0"/>
        <SolidColorBrush x:Key="BackgroundDark" Color="#1e1e1e"/>
        <SolidColorBrush x:Key="BackgroundMedium" Color="#2d2d30"/>
        <SolidColorBrush x:Key="BackgroundLight" Color="#3e3e42"/>
        <SolidColorBrush x:Key="ForegroundLight" Color="#f1f1f1"/>
        <SolidColorBrush x:Key="ForegroundDim" Color="#cccccc"/>
        
        <!-- Modern Button Style -->
        <Style x:Key="ModernButton" TargetType="Button">
            <Setter Property="Background" Value="{StaticResource PrimaryColor}"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="15,8"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="FontWeight" Value="Medium"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" 
                                CornerRadius="4"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" 
                                            VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="{StaticResource SecondaryColor}"/>
                </Trigger>
                <Trigger Property="IsPressed" Value="True">
                    <Setter Property="Background" Value="#0E5A9E"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Background" Value="#666666"/>
                    <Setter Property="Foreground" Value="#999999"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        
        <!-- TextBox Style -->
        <Style x:Key="ModernTextBox" TargetType="TextBox">
            <Setter Property="Background" Value="{StaticResource BackgroundMedium}"/>
            <Setter Property="Foreground" Value="{StaticResource ForegroundLight}"/>
            <Setter Property="BorderBrush" Value="#555555"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="FontSize" Value="14"/>
            <Style.Triggers>
                <Trigger Property="IsFocused" Value="True">
                    <Setter Property="BorderBrush" Value="{StaticResource PrimaryColor}"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        
        <!-- ComboBox Style -->
        <Style x:Key="ModernComboBox" TargetType="ComboBox">
            <Setter Property="Background" Value="{StaticResource BackgroundMedium}"/>
            <Setter Property="Foreground" Value="{StaticResource ForegroundLight}"/>
            <Setter Property="BorderBrush" Value="#555555"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="FontSize" Value="14"/>
        </Style>
    </Window.Resources>
    
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <Border Grid.Row="0" Background="{StaticResource BackgroundMedium}" 
                BorderBrush="{StaticResource BackgroundLight}" BorderThickness="0,0,0,1">
            <Grid Margin="20,15">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                    <StackPanel Grid.Column="0" Orientation="Horizontal">
                        
                        <StackPanel Orientation="Vertical">
                            <TextBlock Text="WMI Filter Management Tool" 
                                     FontSize="24" FontWeight="Light" 
                                     Foreground="{StaticResource ForegroundLight}"/>
                            <TextBlock Text="Manage and test Active Directory WMI Filters" 
                                     FontSize="14" 
                                     Foreground="{StaticResource ForegroundDim}" 
                                     Margin="0,5,0,0"/>
                        </StackPanel>
                    </StackPanel>
                
                <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="Domain:" Foreground="{StaticResource ForegroundDim}" 
                             VerticalAlignment="Center" Margin="0,0,10,0"/>
                    <TextBox Name="txtDomain" Width="250" Style="{StaticResource ModernTextBox}"
                           VerticalAlignment="Center" Margin="0,0,10,0"/>
                    <Button Name="btnConnect" Content="Connect" Style="{StaticResource ModernButton}"
                          Width="100" Margin="0,0,10,0"/>
                    <Button Name="btnRefresh" Content="↻ Refresh" Style="{StaticResource ModernButton}"
                          Width="100" IsEnabled="False"/>
                </StackPanel>
            </Grid>
        </Border>
        
        <!-- Filter Bar -->
        <Border Grid.Row="1" Background="{StaticResource BackgroundLight}" Padding="20,10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <StackPanel Grid.Column="0" Orientation="Horizontal">
                    <TextBlock Text="🔍" FontSize="16" Foreground="{StaticResource ForegroundDim}" 
                             VerticalAlignment="Center" Margin="0,0,10,0"/>
                    <TextBox Name="txtSearch" Width="300" Style="{StaticResource ModernTextBox}"
                           Tag="Search WMI filters..." VerticalAlignment="Center"/>
                    <TextBlock Name="lblStats" Foreground="{StaticResource AccentColor}" 
                             VerticalAlignment="Center" Margin="20,0,0,0" FontWeight="Medium"/>
                </StackPanel>
                
                <CheckBox Grid.Column="1" Name="chkShowOnlyLinked" Content="Show only filters with linked GPOs"
                        Foreground="{StaticResource ForegroundLight}" VerticalAlignment="Center"
                        FontSize="14"/>
            </Grid>
        </Border>
        
        <!-- Main Content -->
        <Grid Grid.Row="2">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="450"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            
            <!-- WMI Filters List -->
            <Border Grid.Column="0" Background="{StaticResource BackgroundMedium}" 
                    BorderBrush="{StaticResource BackgroundLight}" BorderThickness="0,0,1,0">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    
                    <Border Grid.Row="0" Background="{StaticResource BackgroundLight}" Padding="15,10">
                        <TextBlock Text="WMI Filters" FontSize="16" FontWeight="Medium"
                                 Foreground="{StaticResource ForegroundLight}"/>
                    </Border>
                    
                    <ListBox Grid.Row="1" Name="lstFilters" Background="Transparent" 
                           BorderThickness="0" ScrollViewer.HorizontalScrollBarVisibility="Disabled">
                        <ListBox.ItemContainerStyle>
                            <Style TargetType="ListBoxItem">
                                <Setter Property="Padding" Value="0"/>
                                <Setter Property="Margin" Value="5,2"/>
                                <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
                                <Setter Property="Template">
                                    <Setter.Value>
                                        <ControlTemplate TargetType="ListBoxItem">
                                            <Border Name="Border" Background="Transparent" 
                                                  CornerRadius="4" Padding="15,12">
                                                <ContentPresenter/>
                                            </Border>
                                            <ControlTemplate.Triggers>
                                                <Trigger Property="IsMouseOver" Value="True">
                                                    <Setter TargetName="Border" Property="Background" 
                                                          Value="{StaticResource BackgroundLight}"/>
                                                </Trigger>
                                                <Trigger Property="IsSelected" Value="True">
                                                    <Setter TargetName="Border" Property="Background" 
                                                          Value="{StaticResource PrimaryColor}"/>
                                                </Trigger>
                                            </ControlTemplate.Triggers>
                                        </ControlTemplate>
                                    </Setter.Value>
                                </Setter>
                            </Style>
                        </ListBox.ItemContainerStyle>
                        <ListBox.ItemTemplate>
                            <DataTemplate>
                                <Grid>
                                    <Grid.RowDefinitions>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                    </Grid.RowDefinitions>
                                    
                                    <TextBlock Grid.Row="0" Text="{Binding Name}" FontSize="15" 
                                             FontWeight="Medium" Foreground="{StaticResource ForegroundLight}"/>
                                    <TextBlock Grid.Row="1" Text="{Binding Description}" FontSize="12" 
                                             Foreground="{StaticResource ForegroundDim}" Margin="0,2,0,0"
                                             TextTrimming="CharacterEllipsis"/>
                                    <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,4,0,0">
                                        <Border Background="{StaticResource AccentColor}" CornerRadius="3" 
                                              Padding="6,2" Margin="0,0,8,0">
                                            <TextBlock Text="{Binding LinkedGPOCount}" FontSize="11" 
                                                     Foreground="{StaticResource BackgroundDark}" FontWeight="Medium"/>
                                        </Border>
                                        <TextBlock Text="{Binding QueryPreview}" FontSize="11" 
                                                 Foreground="{StaticResource ForegroundDim}" 
                                                 TextTrimming="CharacterEllipsis"/>
                                    </StackPanel>
                                </Grid>
                            </DataTemplate>
                        </ListBox.ItemTemplate>
                    </ListBox>
                </Grid>
            </Border>
            
            <!-- Details Panel -->
            <Grid Grid.Column="1">
                <Grid.RowDefinitions>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                
                <!-- Filter Details -->
                <ScrollViewer Grid.Row="0" VerticalScrollBarVisibility="Auto">
                    <StackPanel Name="pnlDetails" Margin="20" Visibility="Collapsed">
                        <TextBlock Text="Filter Details" FontSize="20" FontWeight="Light"
                                 Foreground="{StaticResource ForegroundLight}" Margin="0,0,0,20"/>
                        
                        <!-- Basic Info -->
                        <Border Background="{StaticResource BackgroundMedium}" CornerRadius="4" Padding="15">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="120"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                
                                <TextBlock Grid.Row="0" Grid.Column="0" Text="Name:" 
                                         Foreground="{StaticResource ForegroundDim}" Margin="0,0,0,8"/>
                                <TextBlock Grid.Row="0" Grid.Column="1" Name="lblName" 
                                         Foreground="{StaticResource ForegroundLight}" FontWeight="Medium" 
                                         Margin="0,0,0,8" TextWrapping="Wrap"/>
                                
                                <TextBlock Grid.Row="1" Grid.Column="0" Text="Description:" 
                                         Foreground="{StaticResource ForegroundDim}" Margin="0,0,0,8"/>
                                <TextBlock Grid.Row="1" Grid.Column="1" Name="lblDescription" 
                                         Foreground="{StaticResource ForegroundLight}" Margin="0,0,0,8" 
                                         TextWrapping="Wrap"/>
                                
                                <TextBlock Grid.Row="2" Grid.Column="0" Text="Author:" 
                                         Foreground="{StaticResource ForegroundDim}" Margin="0,0,0,8"/>
                                <TextBlock Grid.Row="2" Grid.Column="1" Name="lblAuthor" 
                                         Foreground="{StaticResource ForegroundLight}" Margin="0,0,0,8"/>
                                
                                <TextBlock Grid.Row="3" Grid.Column="0" Text="Filter ID:" 
                                         Foreground="{StaticResource ForegroundDim}"/>
                                <TextBlock Grid.Row="3" Grid.Column="1" Name="lblFilterID" 
                                         Foreground="{StaticResource ForegroundLight}" FontFamily="Consolas"/>
                            </Grid>
                        </Border>
                        
                        <!-- WQL Query -->
                        <TextBlock Text="WQL Query" FontSize="16" FontWeight="Medium"
                                 Foreground="{StaticResource ForegroundLight}" Margin="0,20,0,10"/>
                        <Border Background="{StaticResource BackgroundMedium}" CornerRadius="4" Padding="15">
                            <TextBox Name="txtQuery" IsReadOnly="True" TextWrapping="Wrap"
                                   Background="Transparent" BorderThickness="0"
                                   Foreground="{StaticResource AccentColor}" FontFamily="Consolas"
                                   FontSize="13"/>
                        </Border>
                        
                        <!-- Linked GPOs -->
                        <StackPanel Name="pnlLinkedGPOs">
                            <TextBlock Text="Linked GPOs" FontSize="16" FontWeight="Medium"
                                     Foreground="{StaticResource ForegroundLight}" Margin="0,20,0,10"/>
                            <Border Background="{StaticResource BackgroundMedium}" CornerRadius="4" Padding="15">
                                <ItemsControl Name="lstLinkedGPOs">
                                    <ItemsControl.ItemTemplate>
                                        <DataTemplate>
                                            <Border Background="{StaticResource BackgroundLight}" CornerRadius="3"
                                                  Padding="10,8" Margin="0,0,0,5">
                                                <TextBlock Text="{Binding}" Foreground="{StaticResource ForegroundLight}"/>
                                            </Border>
                                        </DataTemplate>
                                    </ItemsControl.ItemTemplate>
                                </ItemsControl>
                            </Border>
                        </StackPanel>
                        
                        <!-- Test Section -->
                        <TextBlock Text="Test Filter" FontSize="16" FontWeight="Medium"
                                 Foreground="{StaticResource ForegroundLight}" Margin="0,20,0,10"/>
                        <Border Background="{StaticResource BackgroundMedium}" CornerRadius="4" Padding="15">
                            <StackPanel>
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
    
                                    <TextBox Grid.Column="0" Name="txtTestComputer" 
                                           Style="{StaticResource ModernTextBox}"
                                           Tag="Enter computer name..." Margin="0,0,10,0"/>
                                    <CheckBox Grid.Column="1" Name="chkUseCredentials" 
                                            Content="Use alternate credentials" 
                                            Foreground="{StaticResource ForegroundLight}" 
                                            VerticalAlignment="Center" Margin="0,0,10,0"/>
                                    <Button Grid.Column="2" Name="btnTest" Content="Test Filter" 
                                          Style="{StaticResource ModernButton}" Width="120"/>
                                </Grid>
                                
                                <!-- Test Results -->
                                <Border Name="pnlTestResults" Visibility="Collapsed" Margin="0,15,0,0"
                                      Background="{StaticResource BackgroundLight}" CornerRadius="4" Padding="15">
                                    <StackPanel>
                                        <TextBlock Name="lblTestStatus" FontSize="16" FontWeight="Medium"
                                                 Margin="0,0,0,10"/>
                                        <TextBox Name="txtTestResults" IsReadOnly="True" TextWrapping="Wrap"
                                               Background="Transparent" BorderThickness="0"
                                               Foreground="{StaticResource ForegroundLight}" FontFamily="Consolas"
                                               MaxHeight="200" VerticalScrollBarVisibility="Auto"/>
                                    </StackPanel>
                                </Border>
                            </StackPanel>
                        </Border>
                    </StackPanel>
                    
                   
                </ScrollViewer>
                 <!-- No Selection Message -->
                    <Grid Name="pnlNoSelection" HorizontalAlignment="Center" VerticalAlignment="Center">
                        <StackPanel>
                            <TextBlock Text="📋" FontSize="48" Foreground="{StaticResource ForegroundDim}" 
                                     HorizontalAlignment="Center" Margin="0,0,0,20"/>
                            <TextBlock Text="Select a WMI filter to view details" FontSize="18" 
                                     Foreground="{StaticResource ForegroundDim}" HorizontalAlignment="Center"/>
                        </StackPanel>
                    </Grid>
            </Grid>
        </Grid>
        
        <!-- Status Bar -->
        <Border Grid.Row="3" Background="{StaticResource BackgroundMedium}" 
                BorderBrush="{StaticResource BackgroundLight}" BorderThickness="0,1,0,0">
            <Grid Margin="20,10">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <TextBlock Grid.Column="0" Name="lblStatus" Foreground="{StaticResource ForegroundDim}"
                         VerticalAlignment="Center"/>
                <TextBlock Grid.Column="1" Name="lblDomainInfo" Foreground="{StaticResource ForegroundDim}"
                         VerticalAlignment="Center"/>
            </Grid>
        </Border>
        
        <!-- Loading Overlay -->
        <Grid Name="LoadingOverlay" Grid.RowSpan="4" Background="#AA000000" Visibility="Collapsed">
            <Border Background="{StaticResource BackgroundMedium}" CornerRadius="8" 
                    HorizontalAlignment="Center" VerticalAlignment="Center" Padding="40,30">
                <StackPanel>
                    <ProgressBar IsIndeterminate="True" Width="200" Height="4" 
                               Foreground="{StaticResource PrimaryColor}"/>
                    <TextBlock Name="lblLoadingText" Text="Loading..." FontSize="16" 
                             Foreground="{StaticResource ForegroundLight}" 
                             HorizontalAlignment="Center" Margin="0,15,0,0"/>
                </StackPanel>
            </Border>
        </Grid>
    </Grid>
</Window>
'@

# Create WPF Window
$reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Set window icon using a system icon (simpler and more reliable)
Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;
    public class Win32 {
        [DllImport("shell32.dll")]
        public static extern IntPtr ExtractIcon(IntPtr hInst, string lpszExeFileName, int nIconIndex);
    }
"@

# Use a built-in Windows icon (13 is the settings/gear icon)
$iconHandle = [Win32]::ExtractIcon([System.IntPtr]::Zero, "shell32.dll", 12)
if ($iconHandle -ne [System.IntPtr]::Zero) {
    $window.Icon = [System.Windows.Interop.Imaging]::CreateBitmapSourceFromHIcon(
        $iconHandle,
        [System.Windows.Int32Rect]::Empty,
        [System.Windows.Media.Imaging.BitmapSizeOptions]::FromEmptyOptions()
    )
}

# Get controls
$chkUseCredentials = $window.FindName("chkUseCredentials")
$pnlLinkedGPOs = $window.FindName("pnlLinkedGPOs")
$txtDomain = $window.FindName("txtDomain")
$btnConnect = $window.FindName("btnConnect")
$btnRefresh = $window.FindName("btnRefresh")
$txtSearch = $window.FindName("txtSearch")
$lblStats = $window.FindName("lblStats")
$chkShowOnlyLinked = $window.FindName("chkShowOnlyLinked")
$lstFilters = $window.FindName("lstFilters")
$pnlDetails = $window.FindName("pnlDetails")
$pnlNoSelection = $window.FindName("pnlNoSelection")
$lblName = $window.FindName("lblName")
$lblDescription = $window.FindName("lblDescription")
$lblAuthor = $window.FindName("lblAuthor")
$lblFilterID = $window.FindName("lblFilterID")
$txtQuery = $window.FindName("txtQuery")
$lstLinkedGPOs = $window.FindName("lstLinkedGPOs")
$txtTestComputer = $window.FindName("txtTestComputer")
$btnTest = $window.FindName("btnTest")
$pnlTestResults = $window.FindName("pnlTestResults")
$lblTestStatus = $window.FindName("lblTestStatus")
$txtTestResults = $window.FindName("txtTestResults")
$lblStatus = $window.FindName("lblStatus")
$lblDomainInfo = $window.FindName("lblDomainInfo")
$LoadingOverlay = $window.FindName("LoadingOverlay")
$lblLoadingText = $window.FindName("lblLoadingText")

# Global variables
$script:DomainData = $null
$script:AllFilters = @()
$script:FilteredFilters = @()

# Helper Functions
function Show-Loading {
    param([string]$Text = "Loading...")
    $LoadingOverlay.Visibility = "Visible"
    $lblLoadingText.Text = $Text
    $window.UpdateLayout()
    [System.Windows.Forms.Application]::DoEvents()
}

function Hide-Loading {
    $LoadingOverlay.Visibility = "Collapsed"
}

function Update-Status {
    param([string]$Text, [string]$Color = "ForegroundDim")
    $lblStatus.Text = $Text
    if ($Color -eq "Error") {
        $lblStatus.Foreground = "#FF6B6B"
    } elseif ($Color -eq "Success") {
        $lblStatus.Foreground = $window.Resources["AccentColor"]
    } else {
        $lblStatus.Foreground = $window.Resources[$Color]
    }
}

# Function to parse GUID from gPCWQLFilter format [domain;{GUID};0]
function Get-WMIFilterGUID {
    param($filterString)
    if ($filterString -match '\{([^}]+)\}') {
        return $matches[1]
    }
    return $null
}

function Get-DomainData {
    param([string]$DomainDN)
    
    try {
        # Convert DN to domain name
        $DomainName = $DomainDN.Replace('DC=','').Replace(',','.')
        
        # Get domain controller
        try {
            $DC = (Get-ADDomainController -DomainName $DomainName -Discover -NextClosestSite -ErrorAction Stop).HostName[0]
        }
        catch {
            $DC = (Get-ADDomain -Identity $DomainDN -ErrorAction Stop).PDCEmulator
        }
        
        # Get all WMI filters and cache them by GUID
        $wmiSearcher = New-Object DirectoryServices.DirectorySearcher
        $wmiSearcher.SearchRoot = [ADSI]"LDAP://$DC/CN=SOM,CN=WMIPolicy,CN=System,$DomainDN"
        $wmiSearcher.Filter = "(objectClass=msWMI-Som)"
        $wmiSearcher.PropertiesToLoad.AddRange(@(
            "mswmi-id", 
            "mswmi-name", 
            "mswmi-parm1", 
            "mswmi-parm2", 
            "mswmi-author",
            "mswmi-creationdate",
            "mswmi-changedate",
            "distinguishedname"
        ))
        
        $wmiFiltersHash = @{}
        $wmiFilters = @()
        
        $wmiSearcher.FindAll() | ForEach-Object {
            $props = $_.Properties
            $guid = if ($props["mswmi-id"] -and $props["mswmi-id"].Count -gt 0) { $props["mswmi-id"][0] } else { $null }
            $name = if ($props["mswmi-name"] -and $props["mswmi-name"].Count -gt 0) { $props["mswmi-name"][0] } else { "Unknown" }
            
            if ($guid -and $name) {
                # Remove braces from GUID if present for consistent matching
                $cleanGuid = $guid -replace '[{}]', ''
                $wmiFiltersHash[$cleanGuid] = $name
                
                $filter = [PSCustomObject]@{
                    Name = $name
                    Description = if ($props["mswmi-parm1"] -and $props["mswmi-parm1"].Count -gt 0) { $props["mswmi-parm1"][0] } else { $null }
                    Query = if ($props["mswmi-parm2"] -and $props["mswmi-parm2"].Count -gt 0) { $props["mswmi-parm2"][0] } else { $null }
                    Author = if ($props["mswmi-author"] -and $props["mswmi-author"].Count -gt 0) { $props["mswmi-author"][0] } else { $null }
                    ID = $guid
                    DistinguishedName = if ($props["distinguishedname"] -and $props["distinguishedname"].Count -gt 0) { $props["distinguishedname"][0] } else { $null }
                    CreatedTime = if ($props["mswmi-creationdate"] -and $props["mswmi-creationdate"].Count -gt 0) { $props["mswmi-creationdate"][0] } else { $null }
                    ModifiedTime = if ($props["mswmi-changedate"] -and $props["mswmi-changedate"].Count -gt 0) { $props["mswmi-changedate"][0] } else { $null }
                }
                $wmiFilters += $filter
            }
        }
        
        # Now get GPOs with WMI filters using LDAP search (much faster than Get-GPO -All)
        $gpoSearcher = New-Object DirectoryServices.DirectorySearcher
        $gpoSearcher.SearchRoot = [ADSI]"LDAP://$DC/CN=Policies,CN=System,$DomainDN"
        $gpoSearcher.Filter = "(&(objectClass=groupPolicyContainer)(gPCWQLFilter=*))"
        $gpoSearcher.PageSize = 1000
        $gpoSearcher.PropertiesToLoad.AddRange(@(
            "displayName",
            "gPCWQLFilter",
            "cn"
        ))
        
        $gpoWmiLinks = @{}
        $gpoSearcher.FindAll() | ForEach-Object {
            $props = $_.Properties
            $gpoName = if ($props["displayname"] -and $props["displayname"].Count -gt 0) { $props["displayname"][0] } else { "Unknown GPO" }
            $wmiFilterRef = if ($props["gpcwqlfilter"] -and $props["gpcwqlfilter"].Count -gt 0) { $props["gpcwqlfilter"][0] } else { $null }
            
            if ($wmiFilterRef) {
                $wmiFilterGUID = Get-WMIFilterGUID $wmiFilterRef
                if ($wmiFilterGUID) {
                    $cleanGuid = $wmiFilterGUID -replace '[{}]', ''
                    if (-not $gpoWmiLinks.ContainsKey($cleanGuid)) {
    $gpoWmiLinks[$cleanGuid] = @()
}
$gpoWmiLinks[$cleanGuid] += $gpoName
                }
            }
        }
        
        return @{
            DomainDN = $DomainDN
            DomainName = $DomainName
            DomainController = $DC
            WMIFilters = $wmiFilters
            WMIFiltersHash = $wmiFiltersHash
            GPOWmiLinks = $gpoWmiLinks
        }
    }
    catch {
        throw $_
    }
}

function Format-WMIFiltersForDisplay {
    param($Filters, $GPOWmiLinks, $WMIFiltersHash)
    
    $DisplayFilters = @()
    foreach ($Filter in $Filters) {
        $cleanGuid = $Filter.ID -replace '[{}]', ''
        $LinkedGPOs = if ($GPOWmiLinks.ContainsKey($cleanGuid)) { $GPOWmiLinks[$cleanGuid] } else { @() }
        $QueryPreview = if ($Filter.Query.Length -gt 50) { $Filter.Query.Substring(0, 47) + "..." } else { $Filter.Query }
        
        $DisplayFilters += [PSCustomObject]@{
            Name = $Filter.Name
            Description = if ($Filter.Description) { $Filter.Description } else { "No description" }
            Query = $Filter.Query
            Author = $Filter.Author
            ID = $Filter.ID
            LinkedGPOs = $LinkedGPOs
            LinkedGPOCount = "$($LinkedGPOs.Count) GPOs"
            QueryPreview = $QueryPreview
            OriginalFilter = $Filter
        }
    }
    return $DisplayFilters
}

function Update-FilterList {
    $searchText = if ($txtSearch.Text -eq $txtSearch.Tag) { "" } else { $txtSearch.Text.ToLower() }
    $showOnlyLinked = $chkShowOnlyLinked.IsChecked
    
    $filtered = $script:AllFilters | Where-Object {
        $matchesSearch = [string]::IsNullOrWhiteSpace($searchText) -or 
                        $_.Name.ToLower().Contains($searchText) -or 
                        $_.Description.ToLower().Contains($searchText) -or
                        $_.Query.ToLower().Contains($searchText)
        
        $matchesLinked = -not $showOnlyLinked -or $_.LinkedGPOs.Count -gt 0
        
        return $matchesSearch -and $matchesLinked
    }
    
    $script:FilteredFilters = @($filtered)
    $lstFilters.ItemsSource = @($filtered)
    
    # Update stats
    $totalFilters = $script:AllFilters.Count
    $displayedFilters = $filtered.Count
    $totalLinked = ($script:AllFilters | Where-Object { $_.LinkedGPOs.Count -gt 0 }).Count
    
    $lblStats.Text = "Showing $displayedFilters of $totalFilters filters | $totalLinked with linked GPOs"
}

function Test-WMIFilter {
    param($Query, $ComputerName, $Credential = $null)
    
    try {
        # Parse WMI filter format
        $QueryParts = $Query -split ';'
        $Queries = @()
        
        $i = 0
        while ($i -lt $QueryParts.Count) {
            if ($QueryParts[$i] -match '^\d+$' -and $i+6 -lt $QueryParts.Count -and $QueryParts[$i+4] -eq 'WQL') {
                $Length = [int]$QueryParts[$i+3]
                $Namespace = $QueryParts[$i+5]
                $WQLQuery = $QueryParts[$i+6]
                
                # Handle queries that might be split across multiple parts
                $CurrentLength = $WQLQuery.Length
                $j = $i + 7
                while ($CurrentLength -lt $Length -and $j -lt $QueryParts.Count) {
                    $WQLQuery += ";" + $QueryParts[$j]
                    $CurrentLength = $WQLQuery.Length
                    $j++
                }
                
                $Queries += @{
                    Namespace = $Namespace
                    Query = $WQLQuery
                }
                
                $i = $j
            }
            else {
                $i++
            }
        }
        
        # If no queries found, try simple parse
        if ($Queries.Count -eq 0) {
            $SimpleQueries = $Query -split ';' | Where-Object { $_ -match 'SELECT' }
            foreach ($SimpleQuery in $SimpleQueries) {
                $Queries += @{
                    Namespace = "root\cimv2"
                    Query = $SimpleQuery.Trim()
                }
            }
        }
        
        # Test each query
        $AllResults = $true
        $ResultDetails = @()
        
        foreach ($QueryInfo in $Queries) {
            $ResultDetails += "Testing query in namespace '$($QueryInfo.Namespace)':"
            $ResultDetails += $QueryInfo.Query
            
            try {
                if ($Credential) {
                        $Result = Get-WmiObject -Query $QueryInfo.Query -ComputerName $ComputerName -Namespace $QueryInfo.Namespace -Credential $Credential -ErrorAction Stop
                    } else {
                        $Result = Get-WmiObject -Query $QueryInfo.Query -ComputerName $ComputerName -Namespace $QueryInfo.Namespace -ErrorAction Stop
                    }
                
                if ($Result) {
                    $ResultDetails += "✓ Query returned $(@($Result).Count) result(s)"
                }
                else {
                    $ResultDetails += "✗ Query returned no results"
                    $AllResults = $false
                }
            }
            catch {
                $ResultDetails += "✗ Query failed: $_"
                $AllResults = $false
            }
            $ResultDetails += ""
        }
        
        return @{
            Result = $AllResults
            Details = $ResultDetails -join "`r`n"
        }
    }
    catch {
        return @{
            Result = $false
            Details = "Error testing WMI filter: $_"
        }
    }
}

# Event Handlers
$btnConnect.Add_Click({
    $domainDN = $txtDomain.Text.Trim()

    # Check if we're already connected (disconnect mode)
if ($btnConnect.Content -eq "Disconnect") {
    # Reset everything
    $script:DomainData = $null
    $script:AllFilters = @()
    $script:FilteredFilters = @()
    $lstFilters.ItemsSource = $null
    $pnlDetails.Visibility = "Collapsed"
    $pnlNoSelection.Visibility = "Visible"
    
    # Reset UI elements
    $txtDomain.IsEnabled = $true
    $btnConnect.Content = "Connect"
    $btnRefresh.IsEnabled = $false
    $lblDomainInfo.Text = ""
    $lblStats.Text = ""
    Update-Status "Disconnected from domain"
    
    return
}
    
    if ([string]::IsNullOrWhiteSpace($domainDN)) {
        [System.Windows.MessageBox]::Show(
            "Please enter a domain DN (e.g., DC=cofomo,DC=microsoft,DC=com)",
            "Domain Required",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        )
        return
    }
    
    if ($domainDN -notmatch '^DC=.+') {
        [System.Windows.MessageBox]::Show(
            "Invalid domain DN format. Expected format: DC=domain,DC=com",
            "Invalid Format",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        )
        return
    }
    
    Show-Loading "Connecting to domain..."
    
    try {
        $script:DomainData = Get-DomainData -DomainDN $domainDN
        
        if ($script:DomainData -and $script:DomainData.WMIFilters.Count -gt 0) {
            # Format filters for display
            $script:AllFilters = Format-WMIFiltersForDisplay -Filters $script:DomainData.WMIFilters -GPOWmiLinks $script:DomainData.GPOWmiLinks -WMIFiltersHash $script:DomainData.WMIFiltersHash
            
            # Update UI
            Update-FilterList
            $btnRefresh.IsEnabled = $true
            $txtDomain.IsEnabled = $false
            $btnConnect.Content = "Disconnect"
            
            $lblDomainInfo.Text = "Connected to: $($script:DomainData.DomainController)"
            Update-Status "Successfully connected to domain" "Success"
        }
        else {
            Update-Status "No WMI filters found in the specified domain" "Error"
        }
    }
    catch {
        Update-Status "Failed to connect: $_" "Error"
        [System.Windows.MessageBox]::Show(
            "Failed to connect to domain:`n`n$_",
            "Connection Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
    finally {
        Hide-Loading
    }
})

$btnRefresh.Add_Click({
    if ($script:DomainData -eq $null) { return }
    
    Show-Loading "Refreshing data..."
    
    try {
        $script:DomainData = Get-DomainData -DomainDN $script:DomainData.DomainDN
        $script:AllFilters = Format-WMIFiltersForDisplay -Filters $script:DomainData.WMIFilters -GPOWmiLinks $script:DomainData.GPOWmiLinks -WMIFiltersHash $script:DomainData.WMIFiltersHash
        Update-FilterList
        Update-Status "Data refreshed successfully" "Success"
    }
    catch {
        Update-Status "Failed to refresh: $_" "Error"
    }
    finally {
        Hide-Loading
    }
})

$txtSearch.Add_TextChanged({
    Update-FilterList
})

$chkShowOnlyLinked.Add_Click({
    Update-FilterList
})

$lstFilters.Add_SelectionChanged({
    $selected = $lstFilters.SelectedItem
    
    if ($selected -eq $null) {
        $pnlDetails.Visibility = "Collapsed"
        $pnlNoSelection.Visibility = "Visible"
        return
    }
    
    $pnlDetails.Visibility = "Visible"
    $pnlNoSelection.Visibility = "Collapsed"
    $pnlTestResults.Visibility = "Collapsed"
    
    # Update details
    $lblName.Text = $selected.Name
    $lblDescription.Text = if ($selected.Description) { $selected.Description } else { "No description" }
    $lblAuthor.Text = if ($selected.Author) { $selected.Author } else { "Unknown" }
    $lblFilterID.Text = $selected.ID
    $txtQuery.Text = $selected.Query
    $lstLinkedGPOs.ItemsSource = @($selected.LinkedGPOs)
    $pnlLinkedGPOs.Visibility = if ($selected.LinkedGPOs.Count -gt 0) { "Visible" } else { "Collapsed" }

    
    Update-Status "Selected filter: $($selected.Name)"
})

$btnTest.Add_Click({
    $computerName = $txtTestComputer.Text.Trim()
    $selected = $lstFilters.SelectedItem
    
    if ([string]::IsNullOrWhiteSpace($computerName) -or $computerName -eq 'Enter computer name...') {
        [System.Windows.MessageBox]::Show(
            "Please enter a computer name to test",
            "Computer Name Required",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        )
        return
    }
    
    if ($selected -eq $null) { return }
    
    $credential = $null
    if ($chkUseCredentials.IsChecked) {
        $credential = Get-Credential -Message "Enter credentials for $computerName"
        if ($null -eq $credential) {
            return  # User cancelled
        }
    }
    Show-Loading "Testing WMI filter..."
    Update-Status "Testing filter on $computerName..."
                $testResult = Test-WMIFilter -Query $selected.Query -ComputerName $computerName -Credential $credential
                Hide-Loading
                
                if ($testResult.Success -eq $false) {
                    $lblTestStatus.Text = "❌ Test Failed"
                    $lblTestStatus.Foreground = "#FF6B6B"
                    $txtTestResults.Text = $testResult.Message
                    Update-Status "Test failed: $($testResult.Message)" "Error"
                }
                else {
                    if ($testResult.Result) {
                        $lblTestStatus.Text = "✅ Filter Result: TRUE - Would Apply"
                        $lblTestStatus.Foreground = $window.Resources["AccentColor"]
                        Update-Status "Filter would apply to $computerName" "Success"
                    }
                    else {
                        $lblTestStatus.Text = "❌ Filter Result: FALSE - Would Not Apply"
                        $lblTestStatus.Foreground = "#FF6B6B"
                        Update-Status "Filter would not apply to $computerName"
                    }
                    $txtTestResults.Text = $result.Details
                }
                
                $pnlTestResults.Visibility = "Visible"
                $pnlTestResults.BringIntoView()

})

# Add placeholder text functionality
$txtSearch.Add_GotFocus({
    if ($txtSearch.Text -eq $txtSearch.Tag) {
        $txtSearch.Text = ""
    }
})

$txtSearch.Add_LostFocus({
    if ([string]::IsNullOrWhiteSpace($txtSearch.Text)) {
        $txtSearch.Text = $txtSearch.Tag
    }
})

$txtTestComputer.Add_GotFocus({
    if ($txtTestComputer.Text -eq $txtTestComputer.Tag) {
        $txtTestComputer.Text = ""
    }
})

$txtTestComputer.Add_LostFocus({
    if ([string]::IsNullOrWhiteSpace($txtTestComputer.Text)) {
        $txtTestComputer.Text = $txtTestComputer.Tag
    }
})

# Initialize placeholder text
$txtSearch.Text = $txtSearch.Tag
$txtTestComputer.Text = $txtTestComputer.Tag

# Set initial status
Update-Status "Ready - Enter domain DN to begin"

# Pre-populate domain if available
try {
    $currentDomain = (Get-ADDomain -Current LocalComputer -ErrorAction SilentlyContinue).DistinguishedName
    if ($currentDomain) {
        $txtDomain.Text = $currentDomain
    }
} catch {
    # Ignore if can't get current domain
}

# Show the window
$window.ShowDialog() | Out-Null
