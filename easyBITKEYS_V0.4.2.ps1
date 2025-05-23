# filepath: c:\Users\PhinITAndreasHepp\OneDrive - PhinIT\05-Administratives\10-Scripte\PowerShell\# EASYit\easyBITkeys\easyBITKEYS_V0.1.0.ps1
# ---------------------------------------------------------------------
# Author: ANDREAS HEPP
# 
# This script reads all BitLocker Keys (msFVE-RecoveryInformation) from Active Directory
# and displays them in a color-coded GUI.
# Status colors:
#   - No key available: LightGray
#   - Key current up to 6 months: LightGreen
#   - Key older than 7 months: Orange
#   - Key older than 12 months: LightCoral
# 
# The application provides buttons to export as CSV, delete the selected entry
# (after confirmation), and close the application.
#
# The DataGrid displays 4 columns:
#   "Nr" (only as wide as the content),
#   "ComputerName" (fills the remaining space),
#   "BitLockerKey" (fills the remaining space),
#   "Datum" (only as wide as the content).
#
# The footer (with URL) is displayed as a panel below the buttons and
# has the same background color as the header (as defined in the INI file).
#
# NOTE: The application can also run in client mode without the AD module,
# displaying sample data for demonstration purposes.
# ---------------------------------------------------------------------

#region Script Configuration
# Diese Konfigurationswerte sollten am Anfang des Skripts stehen.

# Allgemeine Skriptinformationen
$ScriptVersion   = "0.4.2"
$LastUpdate      = "2025-05-23" # Datum der letzten Aktualisierung
$Author          = "ANDREAS HEPP"
$APPName         = "easyBITKEYS"
$APPDescription  = "easyBITKEYS displays BitLocker keys from Active Directory in a color-coded manner."

# GUI-Konfiguration
$ThemeColor      = "#164360" # Hauptfarbe für Header/Footer (Hex-Code)
$FontFamily      = "Segoe UI" # Standard-Schriftart für die GUI
$FontSize        = [float]12   # Standard-Schriftgröße für die GUI
$ConfigHeaderLogoFileName = "APPICON.png"  # Dateiname für das Logo im Header. Muss sich im selben Verzeichnis wie das Skript befinden oder einen relativen/absoluten Pfad haben. Wenn leer, wird kein Logo angezeigt.
$HeaderLogoURL   = "https://www.phinit.de" # URL, die beim Klick auf das Logo geöffnet wird.

# Debug-Modus (0 = Aus, 1 = An) - Stellen Sie sicher, dass die Write-DebugMessage Funktion existiert
$global:DebugMode = 0 # Standardmäßig deaktiviert
$script:Debug = $global:DebugMode # Debug-Modus für das Skript setzen

# Log-Datei Konfiguration
$LogFile = Join-Path $PSScriptRoot "easyBITKEYS.log" # Pfad zur Log-Datei
$MaxLogSizeMB = 5 # Maximale Größe der Log-Datei in MB

#endregion Script Configuration

#region Logging Functions
# Module for logging
function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("INFO", "WARN", "ERROR", "DEBUG")]
        [string]$Level = "INFO"
    )
    
    # Zeitstempel außerhalb des try-Blocks definieren, damit er im catch-Block zuverlässig verfügbar ist.
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    # $logMessage wird später im try-Block initialisiert.
    # Für den Fall, dass ein Fehler vor der Initialisierung von $logMessage auftritt,
    # wird eine Basis-Nachricht für das Fallback-Logging vorbereitet.
    $baseFallbackMessage = "[$timestamp] [$Level] $Message" 

    try {
        # Debug-Meldungen überspringen, wenn der Debug-Modus deaktiviert ist
        if ($Level -eq "DEBUG" -and -not $script:Debug) {
            return
        }
        
        $logMessage = "[$timestamp] [$Level] $Message"
        
        # Sicherstellen, dass das Logs-Verzeichnis existiert
        $logDir = Join-Path $PSScriptRoot "Logs"
        if (-not (Test-Path $logDir)) {
            # -ErrorAction Stop verwenden, um sicherzustellen, dass Fehler beim Erstellen des Verzeichnisses
            # den catch-Block auslösen.
            New-Item -Path $logDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }
        
        # Log-Datei mit Datum
        $logFile = Join-Path $logDir "easyBITKEYS_$(Get-Date -Format 'yyyy-MM-dd').log"
        
        # In Datei schreiben, -ErrorAction Stop verwenden, um Fehler abzufangen
        Add-Content -Path $logFile -Value $logMessage -Encoding UTF8 -ErrorAction Stop
        
        # Auch auf der Konsole ausgeben für Debug-Level und Fehler
        if ($Level -eq "DEBUG" -or $Level -eq "ERROR") {
            Write-Host $logMessage -ForegroundColor $(
                switch ($Level) {
                    "DEBUG" { "Cyan" }
                    "INFO"  { "White" } # Standard-Farbe, kann je nach Host variieren
                    "WARN"  { "Yellow" }
                    "ERROR" { "Red" }
                }
            )
        }
    }
    catch {
        # Fallback-Logging bei Fehlern
        $fallbackLog = Join-Path $env:TEMP "easyBITKEYS_fallback.log"
        
        # Verwende $logMessage, falls es bereits konstruiert wurde, ansonsten die Basis-Nachricht.
        $messageForFallback = if ($logMessage) { $logMessage } else { $baseFallbackMessage }
        
        $errorDetails = "Error during logging operation: $($_.Exception.ToString()). Original log content attempt: $messageForFallback"
        $fallbackEntry = "[$timestamp] [FALLBACK_ERROR] $errorDetails"
        
        try {
            # Versuche, die detaillierte Fehlermeldung in die Fallback-Logdatei zu schreiben.
            # -ErrorAction SilentlyContinue, um zu verhindern, dass ein Fehler im Fallback-Logging das Skript abbricht.
            Add-Content -Path $fallbackLog -Value $fallbackEntry -Encoding UTF8 -ErrorAction SilentlyContinue
        }
        catch {
            # Wenn selbst das Schreiben in die Fallback-Logdatei fehlschlägt, eine Warnung ausgeben.
            Write-Warning "Primary logging failed, and fallback logging to '$fallbackLog' also failed. Error: $($_.Exception.Message)"
        }
    }
}

# Function for debug messages that respects debug setting
function Write-DebugMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message
    )
    
    try {
        if ($script:Debug) {
            Write-Log -Message $Message -Level DEBUG
        }
    }
    catch {
        # Silently fail if debug output fails
    }
}
#endregion

#region Initialization and Requirements Check
# Ensure script runs in PowerShell 5.1
if ($PSVersionTable.PSVersion.Major -ne 5) {
    Write-Host "Switching to PowerShell 5.1..." -ForegroundColor Yellow
    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# Self-Diagnostic at startup
function Test-Requirements {
    try {
        # Check if all icons in assets folder exist
        $assetsPath = Join-Path $PSScriptRoot "assets"
        $requiredIcons = @("info.png", "close.png")
        
        if (-not (Test-Path $assetsPath)) {
            Write-Log -Message "Assets folder not found: $assetsPath" -Level ERROR
            return $false
        }
        
        foreach ($icon in $requiredIcons) {
            $iconPath = Join-Path $assetsPath $icon
            if (-not (Test-Path $iconPath)) {
                Write-Log -Message "Icon not found: $iconPath" -Level ERROR
                return $false
            }
        }
        
        # Check if ActiveDirectory module is available
        if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
            Write-Log -Message "ActiveDirectory module not available" -Level ERROR
            $script:ClientMode = $true
            return $true
        }
        
        # Check permissions (simplified)
        try {
            # Try to read AD objects as permission test
            Get-ADObject -Filter {objectClass -eq "computer"} -ResultSetSize 1 | Out-Null
        }
        catch {
            Write-Log -Message "Insufficient AD permissions: $_" -Level ERROR
            return $false
        }
        
        Write-Log -Message "Self-Diagnostic completed successfully" -Level INFO
        return $true
    }
    catch {
        Write-Log -Message "Error in Self-Diagnostic: $_" -Level ERROR
        return $false
    }
}

# Load required assemblies
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
#endregion

#region Configuration Management
# Die INI-Datei-Verwaltung wurde entfernt. Die Konfiguration erfolgt nun direkt am Anfang des Skripts.
#endregion

#region XAML Definition
# Define XAML markup for the GUI as a plain string first
$xamlString = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="$APPName"
    Width="1200"
    Height="800"
    FontFamily="$FontFamily"
    FontSize="$FontSize"
    ResizeMode="CanMinimize"
    WindowStartupLocation="CenterScreen">

    <Window.Resources>
        <Style x:Key="NavigationButtonStyleUp" TargetType="{x:Type Button}">
            <Setter Property="Height" Value="35" />
            <Setter Property="Margin" Value="10,5" />
            <Setter Property="Background" Value="#3b85b6" />
            <Setter Property="Foreground" Value="White" />
            <Setter Property="BorderThickness" Value="0" />
            <Setter Property="Cursor" Value="Hand" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type Button}">
                        <Border
                            x:Name="border"
                            Background="{TemplateBinding Background}"
                            BorderBrush="{TemplateBinding BorderBrush}"
                            BorderThickness="{TemplateBinding BorderThickness}"
                            CornerRadius="8">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Opacity" Value="0.8" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="NavigationButtonStyleDown" TargetType="{x:Type Button}">
            <Setter Property="Height" Value="25" />
            <Setter Property="Margin" Value="10,5" />
            <Setter Property="Background" Value="#3b85b6" />
            <Setter Property="Foreground" Value="White" />
            <Setter Property="BorderThickness" Value="0" />
            <Setter Property="Cursor" Value="Hand" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type Button}">
                        <Border
                            x:Name="border"
                            Background="{TemplateBinding Background}"
                            BorderBrush="{TemplateBinding BorderBrush}"
                            BorderThickness="{TemplateBinding BorderThickness}"
                            CornerRadius="8">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Opacity" Value="0.8" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="80" />
            <RowDefinition Height="*" />
            <RowDefinition Height="60" />
        </Grid.RowDefinitions>

        <!--  Header area  -->
        <Border
            x:Name="headerBorder"
            Grid.Row="0"
            Background="$ThemeColor">
            <Grid>
                <StackPanel
                    Margin="20,0,0,0"
                    VerticalAlignment="Center"
                    Orientation="Horizontal">
                    <StackPanel>
                        <TextBlock
                            x:Name="headerTitle"
                            FontSize="24"
                            FontWeight="Bold"
                            Foreground="White"
                            Text="$APPName" />
                        <TextBlock
                            x:Name="headerSubtitle"
                            Margin="0,5,0,0"
                            FontSize="12"
                            Foreground="White"
                            Text="Version $ScriptVersion" />
                    </StackPanel>
                </StackPanel>

                <Border
                    Width="175"
                    Height="60"
                    Margin="0,0,20,0"
                    HorizontalAlignment="Right"
                    VerticalAlignment="Center">
                    <Image
                        x:Name="headerLogo"
                        Cursor="Hand"
                        Stretch="Uniform" />
                </Border>
            </Grid>
        </Border>

        <!--  Main area with navigation and content  -->
        <Grid Grid.Row="1">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="150" />
                <ColumnDefinition Width="*" />
            </Grid.ColumnDefinitions>

            <!--  Navigation area  -->
            <Border Grid.Column="0" Background="#F0F0F0">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="*" />
                        <RowDefinition Height="Auto" />
                    </Grid.RowDefinitions>

                    <StackPanel Grid.Row="0">
                        <!--  Obere Gruppe: Allgemeine Aktionen  -->
                        <Separator Margin="10,35,10,15" Visibility="Hidden" />
                        <Button
                            x:Name="btnRefresh"
                            Background="Green"
                            Content="Refresh List"
                            Style="{StaticResource NavigationButtonStyleUp}" />

                        <Button
                            x:Name="btnExport"
                            Background="DarkBlue"
                            Content="Export CSV"
                            Style="{StaticResource NavigationButtonStyleUp}" />

                        <!--  Überschrift und Trennung für untere Gruppe  -->
                        <TextBlock
                            Margin="0,40,0,10"
                            HorizontalAlignment="Center"
                            VerticalAlignment="Center"
                            FontSize="11"
                            FontWeight="SemiBold"
                            Foreground="#444444"
                            Text="FOR SELECTED KEY" />
                        <Separator Margin="10,0,10,15" />

                        <!--  Untere Gruppe: Aktionen für ausgewählten Schlüssel  -->
                        <Button
                            x:Name="btnCopyKey"
                            Background="#2196F3"
                            Content="Copy Key"
                            Style="{StaticResource NavigationButtonStyleDown}" />

                        <Button
                            x:Name="btnDelete"
                            Background="#FF7043"
                            Content="Remove Key"
                            Style="{StaticResource NavigationButtonStyleDown}" />

                        <!--  Überschrift und Trennung für Farblegende  -->
                        <TextBlock
                            Margin="0,80,0,10"
                            HorizontalAlignment="Center"
                            VerticalAlignment="Center"
                            FontSize="11"
                            FontWeight="SemiBold"
                            Foreground="#444444"
                            Text="COLOR LEGEND" />
                        <Separator Margin="10,0,10,5" />

                        <!--  Farblegende Einträge  -->
                        <StackPanel Margin="10,5,10,0" HorizontalAlignment="Left">
                            <StackPanel Margin="0,2" Orientation="Horizontal">
                                <Rectangle
                                    Width="15"
                                    Height="15"
                                    Margin="0,0,5,0"
                                    VerticalAlignment="Center"
                                    Fill="LightGreen" />
                                <TextBlock
                                    VerticalAlignment="Center"
                                    FontSize="11"
                                    Foreground="#555555"
                                    Text="Key &lt; 7 months" />
                            </StackPanel>
                            <StackPanel Margin="0,2" Orientation="Horizontal">
                                <Rectangle
                                    Width="15"
                                    Height="15"
                                    Margin="0,0,5,0"
                                    VerticalAlignment="Center"
                                    Fill="Orange" />
                                <TextBlock
                                    VerticalAlignment="Center"
                                    FontSize="11"
                                    Foreground="#555555"
                                    Text="Key &gt; 6 months" />
                            </StackPanel>
                            <StackPanel Margin="0,2" Orientation="Horizontal">
                                <Rectangle
                                    Width="15"
                                    Height="15"
                                    Margin="0,0,5,0"
                                    VerticalAlignment="Center"
                                    Fill="LightCoral" />
                                <TextBlock
                                    VerticalAlignment="Center"
                                    FontSize="11"
                                    Foreground="#555555"
                                    Text="Key &gt; 12 months" />
                            </StackPanel>
                        </StackPanel>
                    </StackPanel>

                    <!--  Navigation icons at the bottom  -->
                    <StackPanel
                        Grid.Row="1"
                        Margin="0,10,0,10"
                        HorizontalAlignment="Center"
                        Orientation="Horizontal">
                        <Button
                            x:Name="btnInfo"
                            Width="32"
                            Height="32"
                            Margin="5,0"
                            ToolTip="Information">
                            <Image
                                x:Name="infoIcon"
                                Width="24"
                                Height="24" />
                        </Button>
                        <Button
                            x:Name="btnExit"
                            Width="32"
                            Height="32"
                            Margin="5,0"
                            ToolTip="Close Application">
                            <Image
                                x:Name="closeIcon"
                                Width="24"
                                Height="24" />
                        </Button>
                    </StackPanel>
                </Grid>
            </Border>

            <!--  Main content  -->
            <Border
                Grid.Column="1"
                Margin="5,5,5,5"
                Background="White">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto" />
                        <RowDefinition Height="*" />
                    </Grid.RowDefinitions>

                    <!--  Search bar and filter  -->
                    <Border
                        Grid.Row="0"
                        Padding="10"
                        Background="#F5F5F5"
                        BorderBrush="#E0E0E0"
                        BorderThickness="0,0,0,1">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*" />
                                <ColumnDefinition Width="Auto" />
                                <ColumnDefinition Width="Auto" />
                            </Grid.ColumnDefinitions>

                            <!--  Search field  -->
                            <StackPanel Grid.Column="0" Orientation="Horizontal">
                                <TextBlock
                                    Margin="0,0,10,0"
                                    VerticalAlignment="Center"
                                    Text="Search:" />
                                <TextBox
                                    x:Name="txtSearch"
                                    Width="250"
                                    Margin="0,0,10,0"
                                    VerticalContentAlignment="Center" />
                                <ComboBox
                                    x:Name="cmbSearchField"
                                    Width="120"
                                    VerticalContentAlignment="Center"
                                    SelectedIndex="0">
                                    <!--  ComboBox items with proper Accessibility  -->
                                    <ComboBoxItem
                                        AutomationProperties.Name="Search by computer name"
                                        Content="ComputerName"
                                        TabIndex="2"
                                        ToolTip="Search by computer name" />
                                    <ComboBoxItem
                                        AutomationProperties.Name="Search by BitLocker key"
                                        Content="BitLockerKey"
                                        TabIndex="3"
                                        ToolTip="Search by BitLocker key" />
                                    <ComboBoxItem
                                        AutomationProperties.Name="Search all fields"
                                        Content="All Fields"
                                        TabIndex="4"
                                        ToolTip="Search all fields" />
                                </ComboBox>
                                <Button
                                    x:Name="btnSearch"
                                    Margin="10,0,0,0"
                                    Padding="10,0"
                                    Content="Search" />
                                <Button
                                    x:Name="btnClearSearch"
                                    Margin="5,0,0,0"
                                    Padding="5,0"
                                    Content="X"
                                    ToolTip="Reset search" />
                            </StackPanel>

                            <!--  Filter by age  -->
                            <StackPanel
                                Grid.Column="1"
                                Margin="10,0,0,0"
                                Orientation="Horizontal">
                                <TextBlock
                                    Margin="0,0,10,0"
                                    VerticalAlignment="Center"
                                    Text="Filter:" />
                                <ComboBox
                                    x:Name="cmbFilter"
                                    Width="150"
                                    VerticalContentAlignment="Center"
                                    SelectedIndex="0">
                                    <ComboBoxItem Content="Show all" />
                                    <ComboBoxItem Content="Current (&lt; 7 months)" />
                                    <ComboBoxItem Content="Older (7-12 months)" />
                                    <ComboBoxItem Content="Critical (&gt; 12 months)" />
                                </ComboBox>
                                <Button
                                    x:Name="btnApplyFilter"
                                    Margin="10,0,0,0"
                                    Padding="10,0"
                                    Content="Apply" />
                            </StackPanel>
                        </Grid>
                    </Border>

                    <!--  DataGrid for BitLocker Keys  -->
                    <DataGrid
                        x:Name="dataGrid"
                        Grid.Row="1"
                        Margin="5,10,5,5"
                        AlternatingRowBackground="#F5F5F5"
                        AutoGenerateColumns="False"
                        Background="White"
                        BorderBrush="#E0E0E0"
                        BorderThickness="1"
                        CanUserResizeRows="False"
                        GridLinesVisibility="Horizontal"
                        HeadersVisibility="Column"
                        HorizontalScrollBarVisibility="Auto"
                        IsReadOnly="True"
                        RowHeaderWidth="0"
                        SelectionMode="Single"
                        SelectionUnit="FullRow"
                        VerticalScrollBarVisibility="Auto">
                        <DataGrid.Resources>
                            <Style TargetType="{x:Type DataGridColumnHeader}">
                                <Setter Property="Background" Value="#F0F0F0" />
                                <Setter Property="Padding" Value="5,10,5,10" />
                                <Setter Property="FontWeight" Value="SemiBold" />
                            </Style>
                        </DataGrid.Resources>

                        <DataGrid.Columns>
                            <DataGridTextColumn
                                Width="50"
                                Binding="{Binding Nr}"
                                Header="Nr">
                                <DataGridTextColumn.HeaderStyle>
                                    <Style BasedOn="{StaticResource {x:Type DataGridColumnHeader}}" TargetType="DataGridColumnHeader">
                                        <Setter Property="Tag" Value="Nr" />
                                    </Style>
                                </DataGridTextColumn.HeaderStyle>
                            </DataGridTextColumn>
                            <DataGridTextColumn
                                Width="*"
                                Binding="{Binding ComputerName}"
                                Header="ComputerName">
                                <DataGridTextColumn.HeaderStyle>
                                    <Style BasedOn="{StaticResource {x:Type DataGridColumnHeader}}" TargetType="DataGridColumnHeader">
                                        <Setter Property="Tag" Value="ComputerName" />
                                    </Style>
                                </DataGridTextColumn.HeaderStyle>
                            </DataGridTextColumn>
                            <DataGridTextColumn
                                Width="*"
                                Binding="{Binding BitLockerKey}"
                                Header="BitLockerKey">
                                <DataGridTextColumn.HeaderStyle>
                                    <Style BasedOn="{StaticResource {x:Type DataGridColumnHeader}}" TargetType="DataGridColumnHeader">
                                        <Setter Property="Tag" Value="BitLockerKey" />
                                    </Style>
                                </DataGridTextColumn.HeaderStyle>
                            </DataGridTextColumn>
                            <DataGridTextColumn
                                Width="120"
                                Binding="{Binding Datum}"
                                Header="Date">
                                <DataGridTextColumn.HeaderStyle>
                                    <Style BasedOn="{StaticResource {x:Type DataGridColumnHeader}}" TargetType="DataGridColumnHeader">
                                        <Setter Property="Tag" Value="Date" />
                                    </Style>
                                </DataGridTextColumn.HeaderStyle>
                            </DataGridTextColumn>
                        </DataGrid.Columns>
                    </DataGrid>
                </Grid>
            </Border>
        </Grid>

        <!--  Footer area  -->
        <Border
            x:Name="footerBorder"
            Grid.Row="2"
            Background="$ThemeColor">
            <Grid>
                <TextBlock
                    x:Name="versionText"
                    HorizontalAlignment="Center"
                    VerticalAlignment="Center"
                    FontSize="12"
                    Foreground="White"
                    Text="$APPName - Version $ScriptVersion  |  Last Update: $LastUpdate  |  Author: $Author" />

            </Grid>
        </Border>
    </Grid>
</Window>
'@
#endregion

#region Initialize GUI Elements
# Expand PowerShell variables in the XAML string
$expandedXaml = $ExecutionContext.InvokeCommand.ExpandString($xamlString)

# Create XAML reader from the expanded string
try {
    [xml]$xmlDoc = $expandedXaml # Convert to XML object to ensure it's well-formed
} catch {
    Write-Log -Message "Error parsing expanded XAML string. It might not be well-formed XML. Error: $($_.Exception.Message)" -Level ERROR
    Write-Log -Message "Problematic XAML string (first 500 chars): $($expandedXaml.Substring(0, [System.Math]::Min($expandedXaml.Length, 500)))" -Level DEBUG
    # Further error handling or exit might be needed here
    exit 1
}

$reader = New-Object System.Xml.XmlNodeReader $xmlDoc
$window = [Windows.Markup.XamlReader]::Load($reader)

# Reference elements from XAML
$headerBorder = $window.FindName("headerBorder")
$headerTitle = $window.FindName("headerTitle")
$headerSubtitle = $window.FindName("headerSubtitle")
$headerLogo = $window.FindName("headerLogo")
$footerBorder = $window.FindName("footerBorder")
$versionText = $window.FindName("versionText")
$dataGrid = $window.FindName("dataGrid")
$btnRefresh = $window.FindName("btnRefresh")
$btnCopyKey = $window.FindName("btnCopyKey")
$btnExport = $window.FindName("btnExport")
$btnDelete = $window.FindName("btnDelete")
$infoIcon = $window.FindName("infoIcon")
$closeIcon = $window.FindName("closeIcon")
$btnInfo = $window.FindName("btnInfo")
$btnExit = $window.FindName("btnExit")


# Update GUI elements
try {
    Write-DebugMessage -Message "Updating UI elements with configuration values"

    # Set global font family and size for the window from configuration
    if ($null -ne $window) {
        if (-not ([string]::IsNullOrEmpty($FontFamily))) {
            $window.FontFamily = New-Object System.Windows.Media.FontFamily($FontFamily)
            Write-DebugMessage -Message "Set window FontFamily to '$FontFamily'"
        }
        # Ensure FontSize is a valid double for WPF
        if ($FontSize -is [double] -or $FontSize -is [int] -or $FontSize -is [float]) {
            $window.FontSize = [double]$FontSize
            Write-DebugMessage -Message "Set window FontSize to '$FontSize'"
        } else {
            Write-Log -Message "Invalid FontSize value: '$FontSize'. Must be a number." -Level WARN
        }
    }

    # Set theme color
    $themeBrushConverter = New-Object System.Windows.Media.BrushConverter
    $themeColorBrush = $themeBrushConverter.ConvertFromString($ThemeColor)
    
    if ($null -ne $headerBorder) { 
        $headerBorder.Background = $themeColorBrush
        Write-DebugMessage -Message "Set header background to '$ThemeColor'"
    }
    if ($null -ne $footerBorder) { 
        $footerBorder.Background = $themeColorBrush 
        Write-DebugMessage -Message "Set footer background to '$ThemeColor'"
    }
    
    # Set text content
    if ($null -ne $headerTitle) { 
        $headerTitle.Text = $APPName 
        Write-DebugMessage -Message "Set header title to '$APPName'"
    }
    if ($null -ne $headerSubtitle) { 
        $headerSubtitle.Text = $APPDescription 
        Write-DebugMessage -Message "Set header subtitle to '$APPDescription'"
    }
    # The following block for $footerText was removed as $footerText was not initialized
    # and $FooterText variable was not defined. The version information is handled by $versionText.
    
    # GUI footer text with replacements (version, update date, author)
    $versionString = "Version: $ScriptVersion  -  last updated: $LastUpdate  |  Author: $Author"
    if ($null -ne $versionText) { 
        $versionText.Text = $versionString 
        Write-DebugMessage -Message "Set version text to '$versionString'"
    }
    
    # Load logo
    if ($null -ne $headerLogo) {
        if (-not ([string]::IsNullOrEmpty($ConfigHeaderLogoFileName))) { 
            $logoFileOrPathFromConfig = $ConfigHeaderLogoFileName 
            
            $resolvedLogoPath = $logoFileOrPathFromConfig
            # If the configured path isn't absolute, assume it's relative to the script's root directory
            if (-not ([System.IO.Path]::IsPathRooted($logoFileOrPathFromConfig))) {
                $resolvedLogoPath = Join-Path $PSScriptRoot $logoFileOrPathFromConfig
            }

            if (Test-Path $resolvedLogoPath) {
                try {
                    $headerImage = New-Object System.Windows.Media.Imaging.BitmapImage
                    $headerImage.BeginInit()
                    # BitmapImage requires an absolute URI. $resolvedLogoPath should be absolute here.
                    $uri = New-Object System.Uri($resolvedLogoPath, [System.UriKind]::Absolute)
                    $headerImage.UriSource = $uri
                    $headerImage.EndInit()
                    $headerLogo.Source = $headerImage # Set the source of the XAML Image control
                    Write-DebugMessage -Message "Loaded header logo from '$resolvedLogoPath'"
                } catch {
                    Write-Log -Message "Error creating image object for logo. Configured: '$logoFileOrPathFromConfig', Resolved: '$resolvedLogoPath'. Error: $($_.Exception.Message)" -Level ERROR
                }
            }
            else {
                Write-Log -Message "Logo file (configured as '$logoFileOrPathFromConfig') not found. Attempted to load from '$resolvedLogoPath'." -Level WARN
            }
        }
        else {
            Write-DebugMessage -Message "Configuration for header logo file/path ('$ConfigHeaderLogoFileName') is empty. Skipping logo load."
        }
    }
    else {
        # This case means $window.FindName("headerLogo") returned null.
        Write-Log -Message "XAML Image element for the header logo (expected name 'headerLogo') was not found in the window." -Level WARN
    }
    # Load icons
    $iconsPath = Join-Path $PSScriptRoot "assets"
    
    # Load info icon
    if ($null -ne $infoIcon) {
        $infoIconPath = Join-Path $iconsPath "info.png"
        if (Test-Path $infoIconPath) {
            $infoImage = New-Object System.Windows.Media.Imaging.BitmapImage
            $infoImage.BeginInit()
            $infoImage.UriSource = New-Object System.Uri($infoIconPath)
            $infoImage.EndInit()
            $infoIcon.Source = $infoImage
            Write-DebugMessage -Message "Loaded info icon from $infoIconPath"
        }
        else {
            Write-Log -Message "Info icon not found: $infoIconPath" -Level WARN
        }
    }
    
    # Load close icon
    if ($null -ne $closeIcon) {
        $closeIconPath = Join-Path $iconsPath "close.png"
        if (Test-Path $closeIconPath) {
            $closeImage = New-Object System.Windows.Media.Imaging.BitmapImage
            $closeImage.BeginInit()
            $closeImage.UriSource = New-Object System.Uri($closeIconPath)
            $closeImage.EndInit()
            $closeIcon.Source = $closeImage
            Write-DebugMessage -Message "Loaded close icon from $closeIconPath"
        }
        else {
            Write-Log -Message "Close icon not found: $closeIconPath" -Level WARN
        }
    }
}
catch {
    Write-Log -Message "Error customizing the GUI: $_" -Level ERROR
    [System.Windows.MessageBox]::Show("An unexpected error occurred. Please contact support.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
}
#endregion

#region Event Handlers for Navigation
# Event handler for logo click
if ($null -ne $headerLogo) {
    $headerLogo.AddHandler([System.Windows.Controls.Image]::MouseLeftButtonDownEvent, [System.Windows.Input.MouseButtonEventHandler]{
        try {
            Write-DebugMessage -Message "Header logo clicked, opening URL: $HeaderLogoURL"
            Start-Process $HeaderLogoURL
        }
        catch {
            Write-Log -Message "Error opening URL $HeaderLogoURL : $_" -Level ERROR
        }
    })
}

# Event handler for footer click
# if ($null -ne $footerText) {
#    $footerText.Add_MouseLeftButtonDown({
#        try {
#            Write-DebugMessage -Message "Footer text clicked, opening URL: $HeaderLogoURL"
#            Start-Process $HeaderLogoURL
#        }
#        catch {
#            Write-Log -Message "Error opening URL $HeaderLogoURL - $_" -Level ERROR
#        }
#    })
#    # Set cursor style to indicate clickable text
#    $footerText.Cursor = [System.Windows.Input.Cursors]::Hand
# }
#endregion

#region ActiveDirectory Integration
# Import ActiveDirectory module
try {
    Import-Module ActiveDirectory -ErrorAction Stop
    Write-Log -Message "ActiveDirectory module loaded successfully" -Level INFO
}
catch {
    Write-Log -Message "Error loading ActiveDirectory module: $_" -Level ERROR
    [System.Windows.MessageBox]::Show("ActiveDirectory module is not available. Please install RSAT tools.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    $script:ClientMode = $true
}

# Function to load BitLocker keys
function Get-BitLockerKeys {
    try {
        Write-Log -Message "Starting AD query for BitLocker keys..." -Level INFO
        Write-DebugMessage -Message "Executing AD query for msFVE-RecoveryInformation objects"
        
        # AD query: All BitLocker keys (msFVE-RecoveryInformation)
        $bitLockerKeys = Get-ADObject -Filter 'objectClass -eq "msFVE-RecoveryInformation"' -Properties msFVE-RecoveryPassword, whenCreated, distinguishedName
        
        # Clear and refill DataGrid
        $tableData = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
        
        $counter = 1
        foreach ($key in $bitLockerKeys) {
            # Extract the actual computer name from the DistinguishedName
            $dnParts = $key.distinguishedName.Split(',')
            if ($dnParts.Length -ge 2) {
                $computerName = $dnParts[1] -replace "^CN=", ""
            }
            else {
                $computerName = "Unknown"
            }
            
            # Extract BitLocker key and date
            $bitLockerKey = $key.'msFVE-RecoveryPassword'
            $date = $key.whenCreated
            
            # Create object for table data
            $rowObj = New-Object PSObject -Property @{
                Nr = $counter
                ComputerName = $computerName
                BitLockerKey = $bitLockerKey
                Datum = $date.ToString("dd.MM.yyyy")
                DistinguishedName = $key.distinguishedName  # Store for later delete operations
                RowBackground = $null  # Will be set shortly
            }
            
            # Background color based on date
            if ([string]::IsNullOrEmpty($bitLockerKey)) {
                $rowObj.RowBackground = "#D3D3D3"  # LightGray
                Write-DebugMessage -Message "Empty key for $computerName, setting gray background"
            }
            else {
                $ageDays = (Get-Date) - $date
                $ageMonths = $ageDays.Days / 30
                
                if ($ageMonths -ge 12) {
                    $rowObj.RowBackground = "#F08080"  # LightCoral
                    Write-DebugMessage -Message "Key for $computerName is $ageMonths months old (critical)"
                }
                elseif ($ageMonths -ge 7) {
                    $rowObj.RowBackground = "#FFA500"  # Orange
                    Write-DebugMessage -Message "Key for $computerName is $ageMonths months old (warning)"
                }
                else {
                    $rowObj.RowBackground = "#90EE90"  # LightGreen
                    Write-DebugMessage -Message "Key for $computerName is $ageMonths months old (current)"
                }
            }
            
            $tableData.Add($rowObj)
            $counter++
        }
        
        Write-Log -Message "AD query completed: $($tableData.Count) entries found" -Level INFO
        return $tableData
    }
    catch {
        Write-Log -Message "Error in AD query: $_" -Level ERROR
        [System.Windows.MessageBox]::Show("Error in AD query: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return New-Object System.Collections.ObjectModel.ObservableCollection[Object]
    }
}

# Function to load sample BitLocker keys for client mode
function Get-SampleBitLockerKeys {
    Write-Log -Message "Loading sample BitLocker keys for client mode..." -Level INFO
    $tableData = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
    
    $sampleData = @(
        @{ Nr = 1; ComputerName = "ClientPC01"; BitLockerKey = "1234-5678-9012-3456"; Datum = (Get-Date).AddMonths(-2).ToString("dd.MM.yyyy"); RowBackground = "#90EE90" },
        @{ Nr = 2; ComputerName = "ClientPC02"; BitLockerKey = "2345-6789-0123-4567"; Datum = (Get-Date).AddMonths(-8).ToString("dd.MM.yyyy"); RowBackground = "#FFA500" },
        @{ Nr = 3; ComputerName = "ClientPC03"; BitLockerKey = "3456-7890-1234-5678"; Datum = (Get-Date).AddMonths(-13).ToString("dd.MM.yyyy"); RowBackground = "#F08080" },
        @{ Nr = 4; ComputerName = "ClientPC04"; BitLockerKey = ""; Datum = (Get-Date).AddMonths(-1).ToString("dd.MM.yyyy"); RowBackground = "#D3D3D3" }
    )
    
    foreach ($item in $sampleData) {
        $rowObj = New-Object PSObject -Property $item
        $tableData.Add($rowObj)
    }
    
    Write-Log -Message "Sample data loaded: $($tableData.Count) entries" -Level INFO
    return $tableData
}

# Fill DataGrid with BitLocker keys
function Update-BitLockerKeyTable {
    try {
        Write-DebugMessage -Message "Updating BitLocker key table"
        $tableData = if ($script:ClientMode) { Get-SampleBitLockerKeys } else { Get-BitLockerKeys }
        
        # Convert the collection to a proper ObservableCollection
        $observableCollection = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
        
        foreach ($item in $tableData) {
            $observableCollection.Add($item)
        }
        
        # Dispatcher for thread-safe UI updates
        $window.Dispatcher.Invoke([action]{
            $dataGrid.ItemsSource = $observableCollection
            
            # Style adjustment for rows based on RowBackground value
            Update-RowStyling -DataGrid $dataGrid
            
            Write-DebugMessage -Message "DataGrid updated with $($observableCollection.Count) rows"
        })
    }
    catch {
        Write-Log -Message "Error updating table: $_" -Level ERROR
    }
}

function Update-RowStyling {
    param (
        [Parameter(Mandatory=$true)]
        [System.Windows.Controls.DataGrid]$DataGrid
    )
    
    try {
        Write-DebugMessage -Message "Updating row styling for DataGrid"
        
        # Sicherstellen, dass der Handler nur einmal registriert wird
        if (-not $script:RowStylingHandlerRegistered) {
            $DataGrid.add_LoadingRow({ 
                param($sender, $e)
                
                $row = $e.Row
                $item = $row.DataContext
                
                # Überprüfe, ob das Item eine RowBackground-Eigenschaft hat
                if ($null -ne $item -and $null -ne $item.PSObject.Properties['RowBackground']) {
                    $colorCode = $item.RowBackground
                    if (-not [string]::IsNullOrEmpty($colorCode)) {
                        try {
                            $brushConverter = New-Object System.Windows.Media.BrushConverter
                            $brush = $brushConverter.ConvertFromString($colorCode)
                            $row.Background = $brush
                        }
                        catch {
                            Write-Log -Message "Fehler beim Konvertieren der Farbe: $_" -Level WARN
                        }
                    }
                }
            })
            $script:RowStylingHandlerRegistered = $true
            Write-DebugMessage -Message "Row styling event handler wurde registriert"
        }
    }
    catch {
        Write-Log -Message "Fehler beim Aktualisieren des Row-Stylings: $_" -Level ERROR
    }
}
#endregion

#region Button Event Handlers
# Button events
if ($null -ne $btnExport) {
    $btnExport.Add_Click({
        try {
            Write-DebugMessage -Message "Export button clicked"
            $dialog = New-Object Microsoft.Win32.SaveFileDialog
            $dialog.DefaultExt = ".csv"
            $dialog.Filter = "CSV files (*.csv)|*.csv"
            $dialog.Title = "Save BitLocker keys as CSV"
            
            $result = $dialog.ShowDialog()
            
            if ($result -eq $true) {
                $filePath = $dialog.FileName
                
                $csvData = @()
                foreach ($item in $dataGrid.ItemsSource) {
                    $csvData += [PSCustomObject]@{
                        Nr = $item.Nr
                        ComputerName = $item.ComputerName
                        BitLockerKey = $item.BitLockerKey
                        Datum = $item.Datum
                    }
                }
                
                $csvData | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8
                Write-Log -Message "CSV export to $filePath completed" -Level INFO
                Write-DebugMessage -Message "Exported $($csvData.Count) records to $filePath"
                [System.Windows.MessageBox]::Show("Export completed!", "Export", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            }
        }
        catch {
            Write-Log -Message "Error during CSV export: $_" -Level ERROR
            [System.Windows.MessageBox]::Show("Error during export: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    })
} else {
    Write-Log -Message "GUI Element 'btnExport' (Export Button) not found or initialized. Export functionality will be unavailable." -Level WARN
}

if ($null -ne $btnDelete) {
    $btnDelete.Add_Click({
        try {
            Write-DebugMessage -Message "Delete button clicked"
            $selectedItem = $dataGrid.SelectedItem
            
            if ($null -eq $selectedItem) {
                [System.Windows.MessageBox]::Show("Please select an entry from the list!", "Note", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                return
            }
            
            $confirm = [System.Windows.MessageBox]::Show("Really remove the selected entry?", "Remove", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
            
            if ($confirm -eq [System.Windows.MessageBoxResult]::Yes) {
                $dn = $selectedItem.DistinguishedName
                try {
                    # Delete AD object
                    Write-DebugMessage -Message "Removing AD object: $dn"
                    Remove-ADObject -Identity $dn -Confirm:$false
                    Write-Log -Message "AD object removed: $dn" -Level INFO
                    
                    # Update table
                    Update-BitLockerKeyTable
                    
                    [System.Windows.MessageBox]::Show("Entry has been removed.", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                }
                catch {
                    Write-Log -Message "Error removing entry $dn : $_" -Level ERROR
                    [System.Windows.MessageBox]::Show("Error removing entry: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                }
            }
        }
        catch {
            Write-Log -Message "Error in delete function: $_" -Level ERROR
        }
    })
} else {
    Write-Log -Message "GUI Element 'btnDelete' (Delete Button) not found or initialized. Delete functionality will be unavailable." -Level WARN
}

if ($null -ne $btnExit) {
    $btnExit.Add_Click({
        Write-DebugMessage -Message "Exit button clicked, closing application"
        $window.Close()
    })
} else {
    Write-Log -Message "GUI Element 'btnExit' (Exit Button) not found or initialized. Exit functionality will be unavailable." -Level WARN
}

if ($null -ne $btnExit) {
    $btnExit.Add_Click({
        Write-DebugMessage -Message "Close button clicked, closing application"
        $window.Close()
    })
} else {
    Write-Log -Message "GUI Element 'btnExit' (Close Button) not found or initialized. Close functionality will be unavailable." -Level WARN
}

if ($null -ne $btnInfo) {
    $btnInfo.Add_Click({
        try {
            Write-DebugMessage -Message "Info button clicked, showing application information"
            $infoText = @"
$APPName - $ScriptVersion
Last update: $LastUpdate
Author: $Author

This tool displays all BitLocker keys from Active Directory.
Color coding:
- Green: Current key (< 7 months)
- Orange: Key older than 7 months
- Red: Key older than 12 months
- Gray: No key available
"@
            [System.Windows.MessageBox]::Show($infoText, "Info", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        }
        catch {
            Write-Log -Message "Error displaying info: $_" -Level ERROR
        }
    })
} else {
    Write-Log -Message "GUI Element 'btnInfo' (Info Button) not found or initialized. Info functionality will be unavailable." -Level WARN
}

# Refresh function to update AD query
if ($null -ne $btnRefresh) {
    $btnRefresh.Add_Click({
        try {
            Write-DebugMessage -Message "Refresh button clicked"
            Write-Log -Message "Manual refresh of BitLocker keys requested" -Level INFO
            Update-BitLockerKeyTable
            [System.Windows.MessageBox]::Show("Data has been refreshed successfully.", "Refresh", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        }
        catch {
            Write-Log -Message "Error during manual refresh: $_" -Level ERROR
            [System.Windows.MessageBox]::Show("Error during refresh: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    })
} else {
    Write-Log -Message "GUI Element 'btnRefresh' (Refresh Button) not found or initialized. Refresh functionality will be unavailable." -Level WARN
}

# Function to copy selected BitLocker key to clipboard
if ($null -ne $btnCopyKey) {
    $btnCopyKey.Add_Click({
        try {
            Write-DebugMessage -Message "Copy key button clicked"
            $selectedItem = $dataGrid.SelectedItem
            
            if ($null -eq $selectedItem) {
                [System.Windows.MessageBox]::Show("Please select an entry from the list!", "Note", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                return
            }
            
            # Copy BitLocker key to clipboard
            $key = $selectedItem.BitLockerKey
            if ([string]::IsNullOrEmpty($key)) {
                [System.Windows.MessageBox]::Show("The selected entry does not contain a BitLocker key.", "Note", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }
            
            [System.Windows.Forms.Clipboard]::SetText($key)
            Write-Log -Message "BitLocker key for computer $($selectedItem.ComputerName) copied to clipboard" -Level INFO
            Write-DebugMessage -Message "Copied BitLocker key for $($selectedItem.ComputerName) to clipboard"
            [System.Windows.MessageBox]::Show("The BitLocker key has been copied to the clipboard.", "Copied", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        }
        catch {
            Write-Log -Message "Error copying BitLocker key: $_" -Level ERROR
            [System.Windows.MessageBox]::Show("Error copying key: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    })
} else {
    Write-Log -Message "GUI Element 'btnCopyKey' (Copy Key Button) not found or initialized. Copy key functionality will be unavailable." -Level WARN
}
#endregion

#region Search and Filter Functions
# References for search and filter components
$txtSearch = $window.FindName("txtSearch")
$cmbSearchField = $window.FindName("cmbSearchField")
$btnSearch = $window.FindName("btnSearch")
$btnClearSearch = $window.FindName("btnClearSearch")
$cmbFilter = $window.FindName("cmbFilter")
$btnApplyFilter = $window.FindName("btnApplyFilter")

# Store original data for search
$global:originalData = $null

# Event handler for search
if ($null -ne $btnSearch) {
    $btnSearch.Add_Click({
        try {
            Write-DebugMessage -Message "Search button clicked"
            # Ensure original data is saved
            if ($null -eq $global:originalData) {
                $global:originalData = $dataGrid.ItemsSource
            }
            
            $searchText = $txtSearch.Text.Trim()
            if ([string]::IsNullOrEmpty($searchText)) {
                [System.Windows.MessageBox]::Show("Please enter a search term.", "Search", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                return
            }
            
            $searchField = ($cmbSearchField.SelectedItem).Content.ToString()
            Write-DebugMessage -Message "Searching for '$searchText' in field '$searchField'"
            
            # Create new filtered collection
            $filteredData = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
            
            foreach ($item in $global:originalData) {
                $match = $false
                
                switch ($searchField) {
                    "ComputerName" {
                        if ($item.ComputerName -like "*$searchText*") { $match = $true }
                    }
                    "BitLockerKey" {
                        if ($item.BitLockerKey -like "*$searchText*") { $match = $true }
                    }
                    "All Fields" {
                        if (
                            ($item.ComputerName -like "*$searchText*") -or
                            ($item.BitLockerKey -like "*$searchText*") -or
                            ($item.Datum -like "*$searchText*")
                        ) { $match = $true }
                    }
                }
                
                if ($match) {
                    $filteredData.Add($item)
                }
            }
            
            $dataGrid.ItemsSource = $filteredData
            
            Write-Log -Message "Search performed: $($filteredData.Count) matches for '$searchText' in field '$searchField'" -Level INFO
            
            if ($filteredData.Count -eq 0) {
                [System.Windows.MessageBox]::Show("No entries found.", "Search", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            }
        }
        catch {
            Write-Log -Message "Error during search: $_" -Level ERROR
            [System.Windows.MessageBox]::Show("Error during search: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    })
} else {
    Write-Log -Message "GUI Element 'btnSearch' (Search Button) not found in XAML. Search functionality will be unavailable." -Level WARN
}

# Event handler for clearing search
if ($null -ne $btnClearSearch) {
    $btnClearSearch.Add_Click({
        try {
            Write-DebugMessage -Message "Clear search button clicked"
            $txtSearch.Text = ""
            
            if ($null -ne $global:originalData) {
                $dataGrid.ItemsSource = $global:originalData
                Write-Log -Message "Search reset" -Level INFO
            }
            else {
                # If no original data stored, reload
                Update-BitLockerKeyTable
            }
        }
        catch {
            Write-Log -Message "Error resetting search: $_" -Level ERROR
        }
    })
} else {
    Write-Log -Message "GUI Element 'btnClearSearch' (Clear Search Button) not found in XAML. Clear search functionality will be unavailable." -Level WARN
}

# Event handler for applying filter
if ($null -ne $btnApplyFilter) {
    $btnApplyFilter.Add_Click({
        try {
            Write-DebugMessage -Message "Apply filter button clicked"
            # Ensure original data is saved
            if ($null -eq $global:originalData) {
                $global:originalData = $dataGrid.ItemsSource
            }
            
            $filterOption = ($cmbFilter.SelectedItem).Content.ToString()
            Write-DebugMessage -Message "Applying filter: $filterOption"
            
            # If "Show all" is selected, show original data
            if ($filterOption -eq "Show all") {
                $dataGrid.ItemsSource = $global:originalData
                Write-Log -Message "Filter reset: Showing all entries" -Level INFO
                return
            }
            
            # Create filtered collection
            $filteredData = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
            
            foreach ($item in $global:originalData) {
                try {
                    # Safely parse date with various formats
                    $datum = $null
                    $dateFormats = @("dd.MM.yyyy", "yyyy-MM-dd", "MM/dd/yyyy")
                    
                    foreach ($format in $dateFormats) {
                        try {
                            $datum = [DateTime]::ParseExact($item.Datum, $format, [System.Globalization.CultureInfo]::InvariantCulture)
                            break # Stop once a format works
                        } catch {
                            # Continue to next format
                        }
                    }
                    
                    # If no format worked, try general parsing
                    if ($null -eq $datum) {
                        if ([DateTime]::TryParse($item.Datum, [ref]$datum)) {
                            # Successfully parsed
                        } else {
                            # Set to minimum date if unparseable
                            $datum = [DateTime]::MinValue
                            Write-DebugMessage -Message "Could not parse date: $($item.Datum) for computer $($item.ComputerName)"
                        }
                    }
                    
                    $ageInMonths = (New-TimeSpan -Start $datum -End (Get-Date)).Days / 30
                    
                    $match = $false
                    
                    switch ($filterOption) {
                        "Current (< 7 months)" {
                            if ($ageInMonths -lt 7) { $match = $true }
                        }
                        "Older (7-12 months)" {
                            if ($ageInMonths -ge 7 -and $ageInMonths -lt 12) { $match = $true }
                        }
                        "Critical (> 12 months)" {
                            if ($ageInMonths -ge 12) { $match = $true }
                        }
                    }
                    
                    if ($match) {
                        $filteredData.Add($item)
                    }
                }
                catch {
                    Write-Log -Message "Error filtering an entry: $_" -Level WARN
                }
            }
            
            $dataGrid.ItemsSource = $filteredData
            
            Write-Log -Message "Filter applied: $($filteredData.Count) entries for option '$filterOption'" -Level INFO
            
            if ($filteredData.Count -eq 0) {
                [System.Windows.MessageBox]::Show("No entries found for the selected filter.", "Filter", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            }
        }
        catch {
            Write-Log -Message "Error applying filter: $_" -Level ERROR
            [System.Windows.MessageBox]::Show("Error filtering: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    })
} else {
    Write-Log -Message "GUI Element 'btnApplyFilter' (Apply Filter Button) not found in XAML. Filter functionality will be unavailable." -Level WARN
}
#endregion

#region DataGrid Interactions
# Implement column sorting
function SortColumn_Click {
    param(
        [Parameter(Mandatory=$true)]
        [System.Object]$sender,
        [Parameter(Mandatory=$true)]
        [System.Windows.RoutedEventArgs]$e
    )
    
    try {
        if ($null -eq $sender -or $null -eq $sender.Column) {
            Write-Log -Message "Invalid column header clicked" -Level WARN
            return
        }
        
        $headerClicked = $sender
        $columnName = $headerClicked.Column.Header.ToString()
        
        Write-Log -Message "Sorting by column: $columnName" -Level INFO
        Write-DebugMessage -Message "Sorting table by column: $columnName"
        
        # Get current data
        $data = $dataGrid.ItemsSource
        
        # Sorting logic based on column
        $sortedData = $null
        
        switch ($columnName) {
            "Nr" {
                $sortedData = [System.Linq.Enumerable]::OrderBy($data, [Func[Object, int]] { param($x) $x.Nr })
            }
            "ComputerName" {
                $sortedData = [System.Linq.Enumerable]::OrderBy($data, [Func[Object, string]] { param($x) $x.ComputerName })
            }
            "BitLockerKey" {
                $sortedData = [System.Linq.Enumerable]::OrderBy($data, [Func[Object, string]] { param($x) $x.BitLockerKey })
            }
            "Date" {
                # Sort date as string in correct format (dd.MM.yyyy)
                $sortedData = [System.Linq.Enumerable]::OrderBy($data, [Func[Object, DateTime]] { 
                    param($x) 
                    try {
                        [DateTime]::ParseExact($x.Datum, "dd.MM.yyyy", [System.Globalization.CultureInfo]::InvariantCulture)
                    } 
                    catch {
                        [DateTime]::MinValue
                    }
                })
            }
        }
        
        if ($null -ne $sortedData) {
            # Convert to ObservableCollection
            $sortedCollection = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
            foreach ($item in $sortedData) {
                $sortedCollection.Add($item)
            }
            
            $dataGrid.ItemsSource = $sortedCollection
            Write-DebugMessage -Message "Table sorted by $columnName"
        }
    }
    catch {
        Write-Log -Message "Error sorting column $columnName : $_" -Level ERROR
    }
}

# Add column sorting event handlers
try {
    # Ensure Windows.Controls namespace is properly loaded
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase

    # Add handler using correct namespace
    $dataGrid.AddHandler([System.Windows.Controls.Primitives.DataGridColumnHeader]::MouseLeftButtonUpEvent, 
        [System.Windows.Input.MouseButtonEventHandler]{
        param($sender, $e)
        
        try {
            $header = [System.Windows.Media.VisualTreeHelper]::HitTest($dataGrid, $e.GetPosition($dataGrid)).VisualHit
            
            # Navigate up the tree to find the column header
            while ($null -ne $header -and -not ($header -is [System.Windows.Controls.Primitives.DataGridColumnHeader])) {
                $header = [System.Windows.Media.VisualTreeHelper]::GetParent($header)
            }
            
            if ($null -ne $header -and $header -is [System.Windows.Controls.Primitives.DataGridColumnHeader]) {
                $columnName = $header.Column.Header.ToString()
                Write-DebugMessage -Message "Column header clicked: $columnName"
                
                # Call the sort function
                $params = @{
                    sender = $header
                    e = $e
                }
                SortColumn_Click @params
            }
        }
        catch {
            Write-Log -Message "Error in column header click handler: $_" -Level ERROR
        }
    })
}
catch {
    Write-Log -Message "Error setting up column sorting: $_" -Level ERROR
}
#endregion

#region Application Initialization and Cleanup
# Run self-diagnostic
if (-not (Test-Requirements)) {
    [System.Windows.MessageBox]::Show("The requirements for execution are not met. See log file for details.", "Self-Diagnostic Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    exit
}

# Initialize
Update-BitLockerKeyTable

# Event for cleanup when closing window
$window.Add_Closed({
    try {
        # Stop INI monitoring
        # if ($null -ne $global:IniWatcherEvent) {
        #     Unregister-Event -SourceIdentifier $global:IniWatcherEvent.Name
        # }
        Write-Log -Message "Application is shutting down" -Level INFO
    }
    catch {
        Write-Log -Message "Error shutting down application: $_" -Level ERROR
    }
})

# Display GUI
$window.Title = $APPName
Write-DebugMessage -Message "Starting GUI display"
[void]$window.ShowDialog()
#endregion

# Search functionality enhancement
# Add key press event to search textbox
if ($null -ne $txtSearch) {
    $txtSearch.Add_KeyDown({
        param($sender, $e)
        
        try {
            # If Enter key is pressed, trigger search
            if ($e.Key -eq 'Enter') {
                $btnSearch.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
            }
        }
        catch {
            Write-Log -Message "Error in search textbox key handler: $_" -Level ERROR
        }
    })
}

# Ensure DataGrid row style is correctly applied on initial load
$dataGrid.Add_Loaded({
    try {
        Write-DebugMessage -Message "DataGrid loaded, applying row styling"
        if ($null -ne $dataGrid.ItemsSource) {
            Update-RowStyling -DataGrid $dataGrid
        }
    }
    catch {
        Write-Log -Message "Error applying row styling after DataGrid load: $_" -Level ERROR
    }
})
