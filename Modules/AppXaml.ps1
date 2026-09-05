#Requires -Version 5.1
<#
.SYNOPSIS
    Google Maps Routes & Map Generator — WPF XAML Layout & Theme Subsystem.
.DESCRIPTION
    Provides the complete XAML definition for the modern WPF Dark/Light mode UI,
    including interactive map container, geocode validation grid, avoid option controls,
    PDF/GPX/KML export buttons, and API cost tracking displays.
.NOTES
    Encoding: UTF-8 with BOM
#>

function Get-AppXaml {
    [CmdletBinding()]
    param()

    [string]$xaml = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Google Maps Route &amp; Map Generator"
    Height="880" Width="1120"
    MinHeight="700" MinWidth="920"
    WindowStartupLocation="CenterScreen"
    Background="{DynamicResource BgDark}" Foreground="{DynamicResource TextPrimary}"
    FontFamily="Segoe UI Variable, Segoe UI, sans-serif">

    <Window.Resources>
        <SolidColorBrush x:Key="BgDark" Color="#0F172A"/>
        <SolidColorBrush x:Key="BgCard" Color="#1E293B"/>
        <SolidColorBrush x:Key="BgCardHover" Color="#293548"/>
        <SolidColorBrush x:Key="BgCardAlt" Color="#162032"/>
        <SolidColorBrush x:Key="BorderCard" Color="#334155"/>
        <SolidColorBrush x:Key="TextPrimary" Color="#F8FAFC"/>
        <SolidColorBrush x:Key="TextSecondary" Color="#94A3B8"/>
        <SolidColorBrush x:Key="AccentBlue" Color="#2563EB"/>
        <SolidColorBrush x:Key="AccentGreen" Color="#10B981"/>
        <SolidColorBrush x:Key="AccentAmber" Color="#F59E0B"/>
        <SolidColorBrush x:Key="AccentRed" Color="#EF4444"/>
        <SolidColorBrush x:Key="BgInput" Color="#1E293B"/>
        <SolidColorBrush x:Key="BorderInput" Color="#334155"/>
        <SolidColorBrush x:Key="BtnSecondaryBg" Color="#334155"/>
        <SolidColorBrush x:Key="BtnSecondaryFg" Color="#F8FAFC"/>
        <SolidColorBrush x:Key="GridLines" Color="#2D3748"/>
        <SolidColorBrush x:Key="LogBg" Color="#0A0F1D"/>
        <SolidColorBrush x:Key="LogFg" Color="#38BDF8"/>
        <SolidColorBrush x:Key="DataGridHeaderBg" Color="#0F172A"/>
        <SolidColorBrush x:Key="DataGridHeaderFg" Color="#94A3B8"/>
        <SolidColorBrush x:Key="DataGridRowBg" Color="#1E293B"/>
        <SolidColorBrush x:Key="DataGridAltRowBg" Color="#162032"/>

        <Style TargetType="TextBox">
            <Setter Property="Background" Value="{DynamicResource BgInput}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderInput}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="9,6"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>

        <Style TargetType="PasswordBox">
            <Setter Property="Background" Value="{DynamicResource BgInput}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderInput}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="9,6"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>

        <Style TargetType="Button">
            <Setter Property="Background" Value="#2563EB"/>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="14,8"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="6"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Custom Dark ScrollBar -->
        <Style TargetType="ScrollBar">
            <Setter Property="Background" Value="#0F172A"/>
            <Setter Property="Foreground" Value="#334155"/>
            <Setter Property="Width" Value="8"/>
        </Style>

        <!-- DataGrid Styles -->
        <Style TargetType="DataGrid">
            <Setter Property="Background" Value="{DynamicResource BgCard}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderCard}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="RowBackground" Value="{DynamicResource DataGridRowBg}"/>
            <Setter Property="AlternatingRowBackground" Value="{DynamicResource DataGridAltRowBg}"/>
            <Setter Property="GridLinesVisibility" Value="Horizontal"/>
            <Setter Property="HorizontalGridLinesBrush" Value="{DynamicResource GridLines}"/>
            <Setter Property="VerticalGridLinesBrush" Value="{DynamicResource GridLines}"/>
            <Setter Property="HeadersVisibility" Value="Column"/>
            <Setter Property="CanUserAddRows" Value="False"/>
            <Setter Property="IsReadOnly" Value="True"/>
            <Setter Property="SelectionMode" Value="Single"/>
            <Setter Property="SelectionUnit" Value="FullRow"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="AutoGenerateColumns" Value="False"/>
        </Style>

        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="{DynamicResource DataGridHeaderBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource DataGridHeaderFg}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderCard}"/>
            <Setter Property="BorderThickness" Value="0,0,1,1"/>
            <Setter Property="Padding" Value="8,7"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="FontSize" Value="12"/>
        </Style>

        <Style TargetType="DataGridRow">
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#1D4ED8"/>
                    <Setter Property="Foreground" Value="#FFFFFF"/>
                </Trigger>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="{DynamicResource BgCardHover}"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style TargetType="DataGridCell">
            <Setter Property="Padding" Value="6,4"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="DataGridCell">
                        <Border Background="{TemplateBinding Background}" BorderThickness="0" Padding="{TemplateBinding Padding}">
                            <ContentPresenter VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#1D4ED8"/>
                    <Setter Property="Foreground" Value="#FFFFFF"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style TargetType="TabItem">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="{DynamicResource TextSecondary}"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="16,10"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Border Name="TabBorder"
                                Background="{TemplateBinding Background}"
                                BorderThickness="0,0,0,2"
                                BorderBrush="Transparent"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter ContentSource="Header" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="TabBorder" Property="BorderBrush" Value="#2563EB"/>
                                <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Dark Theme ComboBox -->
        <ControlTemplate x:Key="DarkComboBoxToggleButton" TargetType="ToggleButton">
            <Border Name="Border" Background="{DynamicResource BgInput}" BorderBrush="{DynamicResource BorderInput}" BorderThickness="1" CornerRadius="4">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="20"/>
                    </Grid.ColumnDefinitions>
                    <Path Name="Arrow" Grid.Column="1" HorizontalAlignment="Center" VerticalAlignment="Center"
                          Data="M 0 0 L 4 4 L 8 0 Z" Fill="{DynamicResource TextSecondary}"/>
                </Grid>
            </Border>
            <ControlTemplate.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter TargetName="Border" Property="BorderBrush" Value="{DynamicResource AccentBlue}"/>
                    <Setter TargetName="Arrow" Property="Fill" Value="{DynamicResource TextPrimary}"/>
                </Trigger>
            </ControlTemplate.Triggers>
        </ControlTemplate>

        <Style TargetType="ComboBox">
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="Background" Value="{DynamicResource BgInput}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderInput}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Padding" Value="8,5"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBox">
                        <Grid>
                            <ToggleButton Name="ToggleButton"
                                          Template="{StaticResource DarkComboBoxToggleButton}"
                                          Focusable="false"
                                          IsChecked="{Binding Path=IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"
                                          ClickMode="Press"/>
                            <ContentPresenter Name="ContentSite"
                                              IsHitTestVisible="False"
                                              Content="{TemplateBinding SelectionBoxItem}"
                                              ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                                              ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}"
                                              Margin="8,4,22,4"
                                              VerticalAlignment="Center"
                                              HorizontalAlignment="Left"/>
                            <Popup Name="Popup"
                                   Placement="Bottom"
                                   IsOpen="{TemplateBinding IsDropDownOpen}"
                                   AllowsTransparency="True"
                                   Focusable="False"
                                   PopupAnimation="Slide">
                                <Grid Name="DropDown"
                                      SnapsToDevicePixels="True"
                                      MinWidth="{TemplateBinding ActualWidth}"
                                      MaxHeight="{TemplateBinding MaxDropDownHeight}">
                                    <Border Name="DropDownBorder"
                                            Background="{DynamicResource BgCard}"
                                            BorderBrush="{DynamicResource BorderCard}"
                                            BorderThickness="1"
                                            CornerRadius="4"/>
                                    <ScrollViewer Margin="2,4" SnapsToDevicePixels="True">
                                        <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Cycle"/>
                                    </ScrollViewer>
                                </Grid>
                            </Popup>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="ComboBoxItem">
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="Background" Value="{DynamicResource BgCard}"/>
            <Setter Property="Padding" Value="8,5"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBoxItem">
                        <Border Name="Border" Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}">
                            <ContentPresenter/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsHighlighted" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="{DynamicResource AccentBlue}"/>
                                <Setter Property="Foreground" Value="#FFFFFF"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="#1D4ED8"/>
                                <Setter Property="Foreground" Value="#FFFFFF"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="ListBox">
            <Setter Property="Background" Value="{DynamicResource BgDark}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderCard}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="FontSize" Value="12"/>
        </Style>

        <Style TargetType="ListBoxItem">
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="Padding" Value="6,4"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ListBoxItem">
                        <Border Name="Border" Background="Transparent" Padding="{TemplateBinding Padding}">
                            <ContentPresenter/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="{DynamicResource BgCardHover}"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="{DynamicResource AccentBlue}"/>
                                <Setter Property="Foreground" Value="#FFFFFF"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <Border Grid.Row="0" Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="14,10" Margin="0,0,0,12">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Orientation="Vertical">
                    <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                        <TextBlock Text="🗺️" FontSize="20" Margin="0,0,8,0" VerticalAlignment="Center"/>
                        <TextBlock Name="txtHeaderTitle" Text="Google Maps Route &amp; Map Generator" FontSize="18" FontWeight="Bold" Foreground="{DynamicResource TextPrimary}"/>
                    </StackPanel>
                    <TextBlock Name="txtHeaderSubtitle" Text="Multi-point driving routes: Fastest, Shortest, Eco-friendly | Import JSON, CSV, Excel" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="28,2,0,0"/>
                </StackPanel>
                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                    <Button Name="btnApiUsageBadge" Background="#1E293B" BorderBrush="#334155" BorderThickness="1" Padding="8,4" Margin="0,0,10,0" Cursor="Hand" ToolTip="Click to view API Usage &amp; Cost Details">
                        <TextBlock Name="txtApiUsageBadge" Text="API: $0.00" Foreground="#38BDF8" FontSize="11" FontWeight="SemiBold"/>
                    </Button>
                    <TextBlock Name="lblApiBadge" Text="API: Checking..." Foreground="#EF4444" FontSize="12" FontWeight="SemiBold" VerticalAlignment="Center" Margin="0,0,10,0"/>
                    <ComboBox Name="cmbAppLanguage" Width="130" Height="30" Margin="0,0,10,0" VerticalAlignment="Center" ToolTip="Select Language"/>
                    <Button Name="btnQuickSettings" Content="⚙ API Settings" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Padding="10,5" FontSize="12" Margin="0,0,10,0"/>
                    <Button Name="btnThemeToggle" Content="🌙 Dark" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" Padding="10,5" FontSize="12" ToolTip="Toggle Light / Dark theme"/>
                </StackPanel>
            </Grid>
        </Border>

        <!-- Main TabControl -->
        <TabControl Name="tabMain" Grid.Row="1" Background="Transparent" BorderThickness="0">

            <!-- TAB 1: MANUAL ROUTE -->
            <TabItem Name="tabItemManual" Header="📍 Manual Route">
                <Grid Margin="0,10,0,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="430" MinWidth="370"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <ScrollViewer VerticalScrollBarVisibility="Auto" Grid.Column="0" Margin="0,0,10,0">
                        <StackPanel>
                            <!-- Route Points Card -->
                            <Border Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="14" Margin="0,0,0,12">
                                <StackPanel>
                                    <TextBlock Name="lblManualRoutePointsHeader" Text="Route Points" FontSize="15" FontWeight="Bold" Foreground="{DynamicResource TextPrimary}" Margin="0,0,0,10"/>

                                    <TextBlock Name="lblManualOrigin" Text="Origin (Start / A):" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,4"/>
                                    <Grid Margin="0,0,0,10">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="Auto"/>
                                        </Grid.ColumnDefinitions>
                                        <TextBox Name="txtManualStart" Text="Warszawa, Plac Defilad 1"/>
                                        <Button Name="btnClearManualStart" Grid.Column="1" Content="✕" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Padding="8,6" Margin="4,0,0,0" ToolTip="Clear"/>
                                    </Grid>

                                    <TextBlock Name="lblManualWaypoints" Text="Intermediate Stops (optional up to 25):" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,4"/>
                                    <Grid Margin="0,0,0,6">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="Auto"/>
                                        </Grid.ColumnDefinitions>
                                        <TextBox Name="txtNewWaypoint" ToolTip="Enter waypoint address and click Add"/>
                                        <Button Name="btnAddWaypoint" Grid.Column="1" Content="➕ Add" Background="#10B981" Margin="4,0,0,0"/>
                                    </Grid>

                                    <ListBox Name="lstWaypoints" Height="100" Margin="0,0,0,6"/>
                                    <Grid Margin="0,0,0,10">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="*"/>
                                        </Grid.ColumnDefinitions>
                                        <Button Name="btnWpUp" Content="▲ Up" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="0,0,2,0" Padding="4,4" FontSize="11"/>
                                        <Button Name="btnWpDown" Grid.Column="1" Content="▼ Down" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="2,0,2,0" Padding="4,4" FontSize="11"/>
                                        <Button Name="btnWpRemove" Grid.Column="2" Content="✕ Remove" Background="#EF4444" Margin="2,0,2,0" Padding="4,4" FontSize="11"/>
                                        <Button Name="btnWpClear" Grid.Column="3" Content="🗑 Clear" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="2,0,0,0" Padding="4,4" FontSize="11"/>
                                    </Grid>

                                    <TextBlock Name="lblManualDestination" Text="Destination (End / B):" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,4"/>
                                    <Grid Margin="0,0,0,6">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="Auto"/>
                                        </Grid.ColumnDefinitions>
                                        <TextBox Name="txtManualEnd" Text="Kraków, Rynek Główny 1"/>
                                        <Button Name="btnClearManualEnd" Grid.Column="1" Content="✕" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Padding="8,6" Margin="4,0,0,0" ToolTip="Clear"/>
                                    </Grid>

                                    <TextBlock Name="lblManualRouteName" Text="Route Name / Description:" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,4,0,4"/>
                                    <TextBox Name="txtManualName" Text="Route Warsaw - Krakow" Margin="0,0,0,4"/>
                                </StackPanel>
                            </Border>

                            <!-- Route Optimization & Avoid Modifiers Card -->
                            <Border Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="14" Margin="0,0,0,12">
                                <StackPanel>
                                    <TextBlock Name="lblManualOptHeader" Text="Route Optimization" FontSize="15" FontWeight="Bold" Foreground="{DynamicResource TextPrimary}" Margin="0,0,0,10"/>

                                    <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                                        <RadioButton Name="rbTypeFastest" Content="⚡ Fastest" IsChecked="True" Foreground="{DynamicResource TextPrimary}" FontSize="13" Margin="0,0,16,0"/>
                                        <RadioButton Name="rbTypeShortest" Content="📏 Shortest" Foreground="{DynamicResource TextPrimary}" FontSize="13" Margin="0,0,16,0"/>
                                        <RadioButton Name="rbTypeEco" Content="🌿 Eco" Foreground="{DynamicResource TextPrimary}" FontSize="13"/>
                                    </StackPanel>

                                    <StackPanel Name="pnlEmission" Orientation="Vertical" Visibility="Collapsed" Margin="0,0,0,8">
                                        <TextBlock Name="lblManualEmission" Text="Vehicle Engine Type (for Eco route):" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,4"/>
                                        <ComboBox Name="cmbEmission">
                                            <ComboBoxItem Content="Gasoline (Benzyna)" Tag="GASOLINE" IsSelected="True"/>
                                            <ComboBoxItem Content="Diesel" Tag="DIESEL"/>
                                            <ComboBoxItem Content="Hybrid" Tag="HYBRID"/>
                                            <ComboBoxItem Content="Electric" Tag="ELECTRIC"/>
                                        </ComboBox>
                                    </StackPanel>

                                    <CheckBox Name="chkTrafficAware" Content="Real-time traffic awareness (Live Traffic)" Foreground="{DynamicResource TextSecondary}" FontSize="12" Margin="0,2,0,8"/>

                                    <!-- Avoid Options (Feature 2.G) -->
                                    <TextBlock Name="lblManualAvoidHeader" Text="Route Avoidance:" FontSize="12" FontWeight="SemiBold" Foreground="{DynamicResource TextSecondary}" Margin="0,4,0,4"/>
                                    <WrapPanel Margin="0,0,0,4">
                                        <CheckBox Name="chkManualAvoidTolls" Content="🚫 Avoid Tolls" Foreground="{DynamicResource TextPrimary}" FontSize="12" Margin="0,0,14,4"/>
                                        <CheckBox Name="chkManualAvoidHighways" Content="🚫 Avoid Highways" Foreground="{DynamicResource TextPrimary}" FontSize="12" Margin="0,0,14,4"/>
                                        <CheckBox Name="chkManualAvoidFerries" Content="🚫 Avoid Ferries" Foreground="{DynamicResource TextPrimary}" FontSize="12" Margin="0,0,14,4"/>
                                    </WrapPanel>
                                </StackPanel>
                            </Border>

                            <Button Name="btnCalculateManual" Content="🚀 CALCULATE ROUTE &amp; DOWNLOAD MAP" Background="#2563EB" Foreground="#FFFFFF" Padding="16,12" FontSize="14" FontWeight="Bold"/>
                        </StackPanel>
                    </ScrollViewer>

                    <!-- Right Column: Results & Map Preview -->
                    <Grid Grid.Column="1" Margin="10,0,0,0">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>

                        <!-- Stat Boxes -->
                        <Border Grid.Row="0" Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="14" Margin="0,0,0,8">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>

                                <StackPanel Grid.Column="0">
                                    <TextBlock Name="lblHeaderDist" Text="DISTANCE" FontSize="11" FontWeight="Bold" Foreground="{DynamicResource TextSecondary}"/>
                                    <TextBlock Name="lblManualDist" Text="— km" FontSize="20" FontWeight="Bold" Foreground="#10B981"/>
                                </StackPanel>

                                <StackPanel Grid.Column="1">
                                    <TextBlock Name="lblHeaderDur" Text="DURATION" FontSize="11" FontWeight="Bold" Foreground="{DynamicResource TextSecondary}"/>
                                    <TextBlock Name="lblManualTime" Text="— min" FontSize="20" FontWeight="Bold" Foreground="#F59E0B"/>
                                </StackPanel>

                                <StackPanel Grid.Column="2">
                                    <TextBlock Name="lblHeaderType" Text="ROUTE TYPE" FontSize="11" FontWeight="Bold" Foreground="{DynamicResource TextSecondary}"/>
                                    <TextBlock Name="lblManualType" Text="Fastest" FontSize="16" FontWeight="SemiBold" Foreground="#38BDF8"/>
                                </StackPanel>

                                <StackPanel Grid.Column="3" VerticalAlignment="Center">
                                    <TextBlock Name="lblManualStatus" Text="Idle" FontSize="12" Foreground="{DynamicResource TextSecondary}" HorizontalAlignment="Right"/>
                                </StackPanel>
                            </Grid>
                        </Border>

                        <!-- Map View Mode Selector (Feature 4.K) -->
                        <Grid Grid.Row="1" Margin="0,0,0,6">
                            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                                <RadioButton Name="rbViewInteractive" Content="🗺️ Interactive Map" GroupName="ManualMapMode" IsChecked="True" Foreground="{DynamicResource TextPrimary}" FontSize="12" Margin="0,0,12,0"/>
                                <RadioButton Name="rbViewStatic" Content="🖼️ Static Map (PNG)" GroupName="ManualMapMode" Foreground="{DynamicResource TextPrimary}" FontSize="12"/>
                            </StackPanel>
                        </Grid>

                        <!-- Map Container -->
                        <Border Grid.Row="2" Background="{DynamicResource BgDark}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="2" Margin="0,0,0,8">
                            <Grid Name="grdMapHost">
                                <TextBlock Name="lblMapPlaceholder" Text="Map preview will appear here after route calculation..."
                                           Foreground="{DynamicResource TextSecondary}" FontSize="14" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                <Image Name="imgMapPreview" Stretch="Uniform" HorizontalAlignment="Center" VerticalAlignment="Center" Visibility="Collapsed"/>
                                <Border Name="pnlInteractiveMapHost" Background="#0A0F1D" CornerRadius="6" Visibility="Collapsed"/>
                            </Grid>
                        </Border>

                        <!-- Action Bar (PDF, GPX, KML, Google Maps) -->
                        <Border Grid.Row="3" Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="10">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Name="lblGoogleUrlDisplay" Text="No generated link" Foreground="{DynamicResource TextSecondary}" FontSize="12" VerticalAlignment="Center" TextTrimming="CharacterEllipsis" Margin="0,0,10,0"/>
                                <StackPanel Grid.Column="1" Orientation="Horizontal">
                                    <Button Name="btnManualExportPdf" Content="📄 PDF" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="0,0,6,0" Padding="10,6" IsEnabled="False" ToolTip="Export route report as PDF"/>
                                    <Button Name="btnManualExportGpx" Content="💾 GPX" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="0,0,6,0" Padding="10,6" IsEnabled="False" ToolTip="Export route track as GPX"/>
                                    <Button Name="btnManualExportKml" Content="💾 KML" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="0,0,6,0" Padding="10,6" IsEnabled="False" ToolTip="Export route track as KML"/>
                                    <Button Name="btnOpenGoogleMaps" Content="🌐 Google Maps" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="0,0,6,0" Padding="10,6" IsEnabled="False"/>
                                    <Button Name="btnCopyUrl" Content="📋 Copy Link" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="0,0,6,0" Padding="10,6" IsEnabled="False"/>
                                    <Button Name="btnSaveMapAs" Content="💾 Save Map As..." Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Padding="10,6" IsEnabled="False"/>
                                </StackPanel>
                            </Grid>
                        </Border>
                    </Grid>
                </Grid>
            </TabItem>

            <!-- TAB 2: BATCH DATA PROCESSING -->
            <TabItem Name="tabItemBatch" Header="📁 Batch File Processing">
                <Grid Margin="0,10,0,0">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <!-- Batch Controls Card -->
                    <Border Grid.Row="0" Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="14" Margin="0,0,0,10">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>

                            <Grid Grid.Row="0" Margin="0,0,0,10">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Name="lblBatchInputFile" Text="Input File (JSON/CSV/XLSX):" VerticalAlignment="Center" Foreground="{DynamicResource TextSecondary}" Margin="0,0,10,0"/>
                                <TextBox Name="txtBatchFilePath" Grid.Column="1" VerticalAlignment="Center"/>
                                <Button Name="btnBrowseBatchFile" Grid.Column="2" Content="📂 Browse File..." Background="#2563EB" Margin="6,0,0,0"/>
                                <Button Name="btnReloadBatchFile" Grid.Column="3" Content="🔄 Reload" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="6,0,0,0"/>
                            </Grid>

                            <Grid Grid.Row="1">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                                    <TextBlock Name="lblBatchFileInfo" Text="No file loaded." Foreground="{DynamicResource TextSecondary}" FontSize="12"/>
                                </StackPanel>
                                <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center" Margin="10,0">
                                    <TextBlock Name="lblBatchDefaultRouteType" Text="Default route type:" Foreground="{DynamicResource TextSecondary}" FontSize="12" VerticalAlignment="Center" Margin="0,0,6,0"/>
                                    <ComboBox Name="cmbBatchRouteType" Width="170">
                                        <ComboBoxItem Content="From Source / Default" Tag="FromSource" IsSelected="True"/>
                                        <ComboBoxItem Content="Fastest (Najszybsza)" Tag="Fastest"/>
                                        <ComboBoxItem Content="Shortest (Najkrótsza)" Tag="Shortest"/>
                                        <ComboBoxItem Content="Eco (Fuel Efficient)" Tag="Eco"/>
                                    </ComboBox>
                                </StackPanel>
                                <StackPanel Grid.Column="3" Orientation="Horizontal">
                                    <!-- Feature 3.J: Validate Addresses button -->
                                    <Button Name="btnValidateBatchGeocoding" Content="🔍 Validate Addresses" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Padding="12,7" FontWeight="SemiBold" Margin="0,0,8,0"/>
                                    <Button Name="btnStartBatch" Content="▶ Start Processing" Background="#10B981" Foreground="#FFFFFF" Padding="14,7" FontWeight="Bold" Margin="0,0,6,0"/>
                                    <Button Name="btnStopBatch" Content="⏹ Stop" Background="#EF4444" Foreground="#FFFFFF" Padding="12,7" IsEnabled="False"/>
                                </StackPanel>
                            </Grid>
                        </Grid>
                    </Border>

                    <!-- Batch Sub Tabs -->
                    <TabControl Name="tabBatchSub" Grid.Row="1" Background="Transparent" BorderThickness="0">
                        <TabItem Name="tabSubInput" Header="📋 Input Data Preview">
                            <Border Background="{DynamicResource BgDark}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="6" Margin="0,6,0,0">
                                <DataGrid Name="dgBatchInput"/>
                            </Border>
                        </TabItem>

                        <!-- Feature 3.J: Geocode Validation Tab -->
                        <TabItem Name="tabSubValidation" Header="🔍 Geocode Validation">
                            <Grid Margin="0,6,0,0">
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                </Grid.RowDefinitions>
                                <Border Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="6" Padding="10,6" Margin="0,0,0,6">
                                    <Grid>
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="Auto"/>
                                        </Grid.ColumnDefinitions>
                                        <TextBlock Name="lblGeocodeValidationSummary" Text="Address validation: Not run yet. Click 'Validate Addresses' to inspect accuracy." Foreground="{DynamicResource TextSecondary}" VerticalAlignment="Center" FontSize="12"/>
                                        <Button Name="btnCopyInvalidAddresses" Grid.Column="1" Content="📋 Copy Invalid Addresses" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Padding="8,4" FontSize="11" IsEnabled="False"/>
                                    </Grid>
                                </Border>
                                <Border Grid.Row="1" Background="{DynamicResource BgDark}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="6">
                                    <DataGrid Name="dgGeocodeValidation" AutoGenerateColumns="False" CanUserAddRows="False" CanUserDeleteRows="False" IsReadOnly="True" SelectionMode="Single" HeadersVisibility="Column">
                                        <DataGrid.Columns>
                                            <DataGridTextColumn Header="#" Binding="{Binding Index}" Width="45"/>
                                            <DataGridTextColumn Header="Address (Input)" Binding="{Binding Address}" Width="280"/>
                                            <DataGridTextColumn Header="Role" Binding="{Binding Role}" Width="120"/>
                                            <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="90"/>
                                            <DataGridTextColumn Header="Precision" Binding="{Binding Precision}" Width="160"/>
                                            <DataGridTextColumn Header="Google Normalized Address" Binding="{Binding FormattedAddress}" Width="*"/>
                                            <DataGridTextColumn Header="Latitude" Binding="{Binding Latitude}" Width="85"/>
                                            <DataGridTextColumn Header="Longitude" Binding="{Binding Longitude}" Width="85"/>
                                        </DataGrid.Columns>
                                    </DataGrid>
                                </Border>
                            </Grid>
                        </TabItem>

                        <TabItem Name="tabSubResults" Header="📊 Calculation Results">
                            <Border Background="{DynamicResource BgDark}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="6" Margin="0,6,0,0">
                                <DataGrid Name="dgBatchResults">
                                    <DataGrid.Columns>
                                        <DataGridTextColumn Header="ID" Binding="{Binding Id}" Width="45"/>
                                        <DataGridTextColumn Header="Route Name" Binding="{Binding Name}" Width="170"/>
                                        <DataGridTextColumn Header="Origin (Start)" Binding="{Binding Start_Original}" Width="180"/>
                                        <DataGridTextColumn Header="Destination (End)" Binding="{Binding End_Original}" Width="180"/>
                                        <DataGridTextColumn Header="Waypoints" Binding="{Binding WaypointsCount}" Width="75"/>
                                        <DataGridTextColumn Header="Type" Binding="{Binding RouteType}" Width="75"/>
                                        <DataGridTextColumn Header="Distance (km)" Binding="{Binding DistanceKm}" Width="95"/>
                                        <DataGridTextColumn Header="Duration (min)" Binding="{Binding DurationMin}" Width="85"/>
                                        <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="110"/>
                                        <DataGridTextColumn Header="PNG Map" Binding="{Binding MapPath}" Width="*"/>
                                    </DataGrid.Columns>
                                </DataGrid>
                            </Border>
                        </TabItem>

                        <TabItem Name="tabSubPoints" Header="📍 Points Detail">
                            <Border Background="{DynamicResource BgDark}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="6" Margin="0,6,0,0">
                                <DataGrid Name="dgBatchPoints">
                                    <DataGrid.Columns>
                                        <DataGridTextColumn Header="Route ID" Binding="{Binding RouteId}" Width="65"/>
                                        <DataGridTextColumn Header="Route Name" Binding="{Binding RouteName}" Width="150"/>
                                        <DataGridTextColumn Header="No." Binding="{Binding PointOrder}" Width="45"/>
                                        <DataGridTextColumn Header="Point Type" Binding="{Binding PointType}" Width="90"/>
                                        <DataGridTextColumn Header="Original Address" Binding="{Binding OriginalAddress}" Width="220"/>
                                        <DataGridTextColumn Header="Geocoded Address" Binding="{Binding GeocodedAddress}" Width="240"/>
                                        <DataGridTextColumn Header="Geocode Status" Binding="{Binding GeocodeStatus}" Width="170"/>
                                        <DataGridTextColumn Header="Match Type" Binding="{Binding MatchType}" Width="110"/>
                                        <DataGridTextColumn Header="Fallback?" Binding="{Binding IsFallback}" Width="75"/>
                                        <DataGridTextColumn Header="Latitude" Binding="{Binding Latitude}" Width="85"/>
                                        <DataGridTextColumn Header="Longitude" Binding="{Binding Longitude}" Width="85"/>
                                    </DataGrid.Columns>
                                </DataGrid>
                            </Border>
                        </TabItem>

                        <TabItem Name="tabSubLog" Header="📝 Activity Log">
                            <Border Background="{DynamicResource BgDark}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="6" Margin="0,6,0,0">
                                <TextBox Name="txtBatchLog" IsReadOnly="True" TextWrapping="Wrap"
                                         VerticalScrollBarVisibility="Auto" FontFamily="Consolas, monospace"
                                         FontSize="12" Background="{DynamicResource LogBg}" Foreground="{DynamicResource LogFg}"/>
                            </Border>
                        </TabItem>
                    </TabControl>

                    <!-- Batch Status & Export Bar -->
                    <Border Grid.Row="2" Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="12" Margin="0,10,0,0">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="230"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>

                            <StackPanel Grid.Column="0" VerticalAlignment="Center">
                                <TextBlock Name="lblBatchProgressText" Text="Ready" FontSize="12" Foreground="{DynamicResource TextSecondary}"/>
                                <ProgressBar Name="pbBatchProgress" Height="14" Minimum="0" Maximum="100" Value="0" Margin="0,4,0,0" Foreground="#10B981" Background="{DynamicResource BgDark}"/>
                            </StackPanel>

                            <TextBlock Name="lblBatchStats" Grid.Column="1" Text="" Foreground="#10B981" FontSize="13" FontWeight="SemiBold" VerticalAlignment="Center" Margin="20,0"/>

                            <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
                                <Button Name="btnOpenOutputDir" Content="📂 Open Output Folder" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="0,0,6,0" Padding="10,6"/>
                                <Button Name="btnBatchExportPdf" Content="📄 PDF" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="0,0,6,0" Padding="10,6"/>
                                <Button Name="btnBatchExportGpx" Content="💾 GPX" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="0,0,6,0" Padding="10,6"/>
                                <Button Name="btnBatchExportKml" Content="💾 KML" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="0,0,6,0" Padding="10,6"/>
                                <Button Name="btnExportExcel" Content="📊 Export Excel" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="0,0,6,0" Padding="10,6"/>
                                <Button Name="btnExportCsv" Content="📄 CSV" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="0,0,6,0" Padding="10,6"/>
                                <Button Name="btnExportJson" Content="📋 JSON" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Padding="10,6"/>
                            </StackPanel>
                        </Grid>
                    </Border>
                </Grid>
            </TabItem>

            <!-- TAB 3: SETTINGS & API KEY -->
            <TabItem Name="tabItemSettings" Header="⚙ Settings &amp; API Key">
                <ScrollViewer VerticalScrollBarVisibility="Auto" Margin="0,10,0,0">
                    <StackPanel MaxWidth="800" HorizontalAlignment="Left">
                        <!-- API Key Card -->
                        <Border Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="16" Margin="0,0,0,14">
                            <StackPanel>
                                <TextBlock Name="lblSettingsApiHeader" Text="Google Maps API Key" FontSize="16" FontWeight="Bold" Foreground="{DynamicResource TextPrimary}" Margin="0,0,0,8"/>
                                <TextBlock Name="lblSettingsApiDesc" Text="Required for Geocoding API, Routes API v2, and Static Maps API." FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,10"/>

                                <TextBlock Name="lblSettingsApiLabel" Text="API Key:" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,4"/>
                                <Grid Margin="0,0,0,8">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <PasswordBox Name="txtSettingsApiKey"/>
                                    <TextBox Name="txtSettingsApiKeyVisible" Visibility="Collapsed"/>
                                    <Button Name="btnToggleKeyVisibility" Grid.Column="1" Content="👁 Show" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="6,0,0,0" Padding="10,6"/>
                                    <Button Name="btnTestApiKey" Grid.Column="2" Content="🔍 Test Key" Background="#2563EB" Margin="6,0,0,0" Padding="12,6"/>
                                </Grid>

                                <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                                    <CheckBox Name="chkRememberKey" Content="Remember securely on this computer (DPAPI CurrentUser encryption)" IsChecked="True" Foreground="{DynamicResource TextSecondary}" FontSize="12"/>
                                </StackPanel>

                                <TextBlock Name="lblKeyTestResult" Text="" FontSize="12" FontWeight="SemiBold"/>
                            </StackPanel>
                        </Border>

                        <!-- CARTO API Key Card -->
                        <Border Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="16" Margin="0,0,0,14">
                            <StackPanel>
                                <TextBlock Name="lblSettingsCartoApiHeader" Text="CARTO API Key (Interactive Map)" FontSize="16" FontWeight="Bold" Foreground="{DynamicResource TextPrimary}" Margin="0,0,0,8"/>
                                <TextBlock Name="lblSettingsCartoApiDesc" Text="Optional. Required only if using authenticated or commercial CARTO private basemap services. Default public raster tiles work without a key." FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,10" TextWrapping="Wrap"/>

                                <TextBlock Name="lblSettingsCartoApiLabel" Text="CARTO API Key / Access Token:" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,4"/>
                                <Grid Margin="0,0,0,8">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <PasswordBox Name="txtSettingsCartoApiKey"/>
                                    <TextBox Name="txtSettingsCartoApiKeyVisible" Visibility="Collapsed"/>
                                    <Button Name="btnToggleCartoKeyVisibility" Grid.Column="1" Content="👁 Show" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="6,0,0,0" Padding="10,6"/>
                                </Grid>
                            </StackPanel>
                        </Border>

                        <!-- API Usage & Cost Estimation Card (Feature 6.S) -->
                        <Border Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="16" Margin="0,0,0,14">
                            <StackPanel>
                                <TextBlock Name="lblSettingsApiUsageHeader" Text="Google Maps Platform API Usage &amp; Cost Tracker" FontSize="16" FontWeight="Bold" Foreground="{DynamicResource TextPrimary}" Margin="0,0,0,6"/>
                                <TextBlock Name="lblSettingsApiUsageDesc" Text="Live tracking of API calls and estimated billable amounts against Google Maps Platform rates ($5/1k Geocoding, $5/1k Routes, $2/1k Static Maps)." FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,12" TextWrapping="Wrap"/>

                                <Grid Margin="0,0,0,10">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    <Border Background="{DynamicResource BgDark}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="6" Padding="10" Margin="0,0,6,0">
                                        <StackPanel>
                                            <TextBlock Text="SESSION CALLS" FontSize="10" FontWeight="Bold" Foreground="{DynamicResource TextSecondary}"/>
                                            <TextBlock Name="lblApiUsageSessionCalls" Text="0 calls" FontSize="16" FontWeight="Bold" Foreground="#38BDF8"/>
                                        </StackPanel>
                                    </Border>
                                    <Border Grid.Column="1" Background="{DynamicResource BgDark}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="6" Padding="10" Margin="3,0,3,0">
                                        <StackPanel>
                                            <TextBlock Text="MONTHLY CALLS" FontSize="10" FontWeight="Bold" Foreground="{DynamicResource TextSecondary}"/>
                                            <TextBlock Name="lblApiUsageMonthlyCalls" Text="0 calls" FontSize="16" FontWeight="Bold" Foreground="#10B981"/>
                                        </StackPanel>
                                    </Border>
                                    <Border Grid.Column="2" Background="{DynamicResource BgDark}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="6" Padding="10" Margin="6,0,0,0">
                                        <StackPanel>
                                            <TextBlock Text="ESTIMATED COST" FontSize="10" FontWeight="Bold" Foreground="{DynamicResource TextSecondary}"/>
                                            <TextBlock Name="lblEstimatedCostMonthly" Text="$0.00" FontSize="16" FontWeight="Bold" Foreground="#F59E0B"/>
                                        </StackPanel>
                                    </Border>
                                </Grid>

                                <Grid Margin="0,4,0,10">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="140"/>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBlock Text="Currency:" VerticalAlignment="Center" Foreground="{DynamicResource TextSecondary}" FontSize="12" Margin="0,0,10,0"/>
                                    <ComboBox Name="cmbApiCurrency" Grid.Column="1" VerticalAlignment="Center">
                                        <ComboBoxItem Content="USD ($)" Tag="USD" IsSelected="True"/>
                                        <ComboBoxItem Content="EUR (€)" Tag="EUR"/>
                                        <ComboBoxItem Content="PLN (zł)" Tag="PLN"/>
                                    </ComboBox>
                                    <TextBlock Name="lblFreeTierInfo" Grid.Column="2" Text="Free tier credit: $200.00 / month" Foreground="{DynamicResource TextSecondary}" FontSize="12" VerticalAlignment="Center" Margin="14,0,0,0"/>
                                    <Button Name="btnResetApiCounters" Grid.Column="3" Content="🔄 Reset Month" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Padding="10,5" FontSize="11"/>
                                </Grid>
                            </StackPanel>
                        </Border>

                        <!-- Preferences Card -->
                        <Border Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="16" Margin="0,0,0,14">
                            <StackPanel>
                                <TextBlock Name="lblSettingsPrefHeader" Text="Default Generation Preferences" FontSize="16" FontWeight="Bold" Foreground="{DynamicResource TextPrimary}" Margin="0,0,0,12"/>

                                <TextBlock Name="lblSettingsDefaultRouteType" Text="Default route type:" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,4"/>
                                <ComboBox Name="cmbDefaultRouteType" Margin="0,0,0,12">
                                    <ComboBoxItem Content="Fastest (Najszybsza)" Tag="Fastest" IsSelected="True"/>
                                    <ComboBoxItem Content="Shortest (Najkrótsza)" Tag="Shortest"/>
                                    <ComboBoxItem Content="Eco (Fuel Efficient)" Tag="Eco"/>
                                </ComboBox>

                                <TextBlock Name="lblSettingsDefaultEmission" Text="Default engine type for Eco routes:" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,4"/>
                                <ComboBox Name="cmbDefaultEmission" Margin="0,0,0,12">
                                    <ComboBoxItem Content="Gasoline (Benzyna)" Tag="GASOLINE" IsSelected="True"/>
                                    <ComboBoxItem Content="Diesel" Tag="DIESEL"/>
                                    <ComboBoxItem Content="Hybrid" Tag="HYBRID"/>
                                    <ComboBoxItem Content="Electric" Tag="ELECTRIC"/>
                                </ComboBox>

                                <!-- Default Avoid Options (Feature 2.G) -->
                                <TextBlock Name="lblSettingsAvoidHeader" Text="Default Route Avoidance Options:" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,6"/>
                                <WrapPanel Margin="0,0,0,12">
                                    <CheckBox Name="chkSettingsAvoidTolls" Content="🚫 Avoid Tolls by default" Foreground="{DynamicResource TextPrimary}" FontSize="12" Margin="0,0,14,4"/>
                                    <CheckBox Name="chkSettingsAvoidHighways" Content="🚫 Avoid Highways by default" Foreground="{DynamicResource TextPrimary}" FontSize="12" Margin="0,0,14,4"/>
                                    <CheckBox Name="chkSettingsAvoidFerries" Content="🚫 Avoid Ferries by default" Foreground="{DynamicResource TextPrimary}" FontSize="12" Margin="0,0,14,4"/>
                                </WrapPanel>

                                <TextBlock Name="lblSettingsDefaultMapSize" Text="Default dimensions for generated PNG map:" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,4"/>
                                <ComboBox Name="cmbDefaultMapSize" Margin="0,0,0,12">
                                    <ComboBoxItem Content="900 x 600 px (Recommended Standard)" Tag="900x600" IsSelected="True"/>
                                    <ComboBoxItem Content="1024 x 768 px (High Res)" Tag="1024x768"/>
                                    <ComboBoxItem Content="1280 x 720 px (HD 16:9)" Tag="1280x720"/>
                                    <ComboBoxItem Content="640 x 640 px (Square)" Tag="640x640"/>
                                    <ComboBoxItem Content="1600 x 900 px (Full HD 16:9)" Tag="1600x900"/>
                                </ComboBox>

                                <TextBlock Name="lblSettingsOutputDir" Text="Results Output Folder:" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,4"/>
                                <Grid Margin="0,0,0,12">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBox Name="txtSettingsOutputDir"/>
                                    <Button Name="btnBrowseSettingsOutputDir" Grid.Column="1" Content="📂 Browse..." Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Margin="6,0,0,0"/>
                                </Grid>
                            </StackPanel>
                        </Border>

                        <!-- MAP OVERLAY CARD -->
                        <Border Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="16" Margin="0,0,0,14">
                            <StackPanel>
                                <TextBlock Name="lblSettingsOverlayHeader" Text="Map Overlay &amp; Banners (Top / Bottom)" FontSize="16" FontWeight="Bold" Foreground="{DynamicResource TextPrimary}" Margin="0,0,0,6"/>
                                <TextBlock Name="lblSettingsOverlayDesc" Text="Configure whether to display top and bottom banner panels, and choose which properties appear on each panel, line order, and alignment." FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,12" TextWrapping="Wrap"/>

                                <StackPanel Orientation="Horizontal" Margin="0,0,0,12">
                                    <CheckBox Name="chkEnableTopOverlay" Content="Enable Top Banner" IsChecked="True" Foreground="{DynamicResource TextPrimary}" FontSize="12" FontWeight="SemiBold" Margin="0,0,24,0"/>
                                    <CheckBox Name="chkEnableBottomOverlay" Content="Enable Bottom Banner" IsChecked="True" Foreground="{DynamicResource TextPrimary}" FontSize="12" FontWeight="SemiBold"/>
                                </StackPanel>

                                <Border Background="{DynamicResource BgDark}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="6" Padding="10" Margin="0,0,0,10">
                                    <Grid Name="gridOverlayConfig">
                                        <Grid.RowDefinitions>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                        </Grid.RowDefinitions>
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="220"/>
                                            <ColumnDefinition Width="65"/>
                                            <ColumnDefinition Width="135"/>
                                            <ColumnDefinition Width="135"/>
                                            <ColumnDefinition Width="110"/>
                                        </Grid.ColumnDefinitions>

                                        <!-- Header Row -->
                                        <TextBlock Name="lblColPropName" Grid.Row="0" Grid.Column="0" Text="Property" FontWeight="Bold" Foreground="{DynamicResource TextSecondary}" FontSize="12" Margin="4,2,4,8"/>
                                        <TextBlock Name="lblColPropShow" Grid.Row="0" Grid.Column="1" Text="Show" FontWeight="Bold" Foreground="{DynamicResource TextSecondary}" FontSize="12" Margin="4,2,4,8" HorizontalAlignment="Center"/>
                                        <TextBlock Name="lblColPropPanel" Grid.Row="0" Grid.Column="2" Text="Panel" FontWeight="Bold" Foreground="{DynamicResource TextSecondary}" FontSize="12" Margin="4,2,4,8"/>
                                        <TextBlock Name="lblColPropAlign" Grid.Row="0" Grid.Column="3" Text="Alignment" FontWeight="Bold" Foreground="{DynamicResource TextSecondary}" FontSize="12" Margin="4,2,4,8"/>
                                        <TextBlock Name="lblColPropOrder" Grid.Row="0" Grid.Column="4" Text="Line / Order" FontWeight="Bold" Foreground="{DynamicResource TextSecondary}" FontSize="12" Margin="4,2,4,8" HorizontalAlignment="Center"/>

                                        <!-- Row 1: StartGeocoded -->
                                        <TextBlock Name="lblProp_StartGeocoded" Grid.Row="1" Grid.Column="0" Text="Start Address (Geocoded)" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_StartGeocoded" Grid.Row="1" Grid.Column="1" IsChecked="True" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_StartGeocoded" Grid.Row="1" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom" IsSelected="True"/>
                                            <ComboBoxItem Content="Top" Tag="Top"/>
                                            <ComboBoxItem Content="None" Tag="None"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_StartGeocoded" Grid.Row="1" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left" IsSelected="True"/>
                                            <ComboBoxItem Content="Center" Tag="Center"/>
                                            <ComboBoxItem Content="Right" Tag="Right"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_StartGeocoded" Grid.Row="1" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1" IsSelected="True"/>
                                            <ComboBoxItem Content="2" Tag="2"/>
                                            <ComboBoxItem Content="3" Tag="3"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>

                                        <!-- Row 2: EndGeocoded -->
                                        <TextBlock Name="lblProp_EndGeocoded" Grid.Row="2" Grid.Column="0" Text="End Address (Geocoded)" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_EndGeocoded" Grid.Row="2" Grid.Column="1" IsChecked="True" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_EndGeocoded" Grid.Row="2" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom" IsSelected="True"/>
                                            <ComboBoxItem Content="Top" Tag="Top"/>
                                            <ComboBoxItem Content="None" Tag="None"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_EndGeocoded" Grid.Row="2" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left" IsSelected="True"/>
                                            <ComboBoxItem Content="Center" Tag="Center"/>
                                            <ComboBoxItem Content="Right" Tag="Right"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_EndGeocoded" Grid.Row="2" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1"/>
                                            <ComboBoxItem Content="2" Tag="2" IsSelected="True"/>
                                            <ComboBoxItem Content="3" Tag="3"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>

                                        <!-- Row 3: Distance -->
                                        <TextBlock Name="lblProp_Distance" Grid.Row="3" Grid.Column="0" Text="Total Distance" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_Distance" Grid.Row="3" Grid.Column="1" IsChecked="True" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_Distance" Grid.Row="3" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom" IsSelected="True"/>
                                            <ComboBoxItem Content="Top" Tag="Top"/>
                                            <ComboBoxItem Content="None" Tag="None"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_Distance" Grid.Row="3" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left" IsSelected="True"/>
                                            <ComboBoxItem Content="Center" Tag="Center"/>
                                            <ComboBoxItem Content="Right" Tag="Right"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_Distance" Grid.Row="3" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1"/>
                                            <ComboBoxItem Content="2" Tag="2"/>
                                            <ComboBoxItem Content="3" Tag="3" IsSelected="True"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>

                                        <!-- Row 4: Duration -->
                                        <TextBlock Name="lblProp_Duration" Grid.Row="4" Grid.Column="0" Text="Total Time" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_Duration" Grid.Row="4" Grid.Column="1" IsChecked="True" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_Duration" Grid.Row="4" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom" IsSelected="True"/>
                                            <ComboBoxItem Content="Top" Tag="Top"/>
                                            <ComboBoxItem Content="None" Tag="None"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_Duration" Grid.Row="4" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left"/>
                                            <ComboBoxItem Content="Center" Tag="Center" IsSelected="True"/>
                                            <ComboBoxItem Content="Right" Tag="Right"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_Duration" Grid.Row="4" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1"/>
                                            <ComboBoxItem Content="2" Tag="2"/>
                                            <ComboBoxItem Content="3" Tag="3" IsSelected="True"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>

                                        <!-- Row 5: Timestamp -->
                                        <TextBlock Name="lblProp_Timestamp" Grid.Row="5" Grid.Column="0" Text="Generation Timestamp" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_Timestamp" Grid.Row="5" Grid.Column="1" IsChecked="True" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_Timestamp" Grid.Row="5" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom" IsSelected="True"/>
                                            <ComboBoxItem Content="Top" Tag="Top"/>
                                            <ComboBoxItem Content="None" Tag="None"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_Timestamp" Grid.Row="5" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left"/>
                                            <ComboBoxItem Content="Center" Tag="Center"/>
                                            <ComboBoxItem Content="Right" Tag="Right" IsSelected="True"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_Timestamp" Grid.Row="5" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1"/>
                                            <ComboBoxItem Content="2" Tag="2"/>
                                            <ComboBoxItem Content="3" Tag="3" IsSelected="True"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>

                                        <!-- Row 6: RouteName -->
                                        <TextBlock Name="lblProp_RouteName" Grid.Row="6" Grid.Column="0" Text="Route Name" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_RouteName" Grid.Row="6" Grid.Column="1" IsChecked="True" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_RouteName" Grid.Row="6" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom"/>
                                            <ComboBoxItem Content="Top" Tag="Top" IsSelected="True"/>
                                            <ComboBoxItem Content="None" Tag="None"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_RouteName" Grid.Row="6" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left" IsSelected="True"/>
                                            <ComboBoxItem Content="Center" Tag="Center"/>
                                            <ComboBoxItem Content="Right" Tag="Right"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_RouteName" Grid.Row="6" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1" IsSelected="True"/>
                                            <ComboBoxItem Content="2" Tag="2"/>
                                            <ComboBoxItem Content="3" Tag="3"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>

                                        <!-- Row 7: RouteType -->
                                        <TextBlock Name="lblProp_RouteType" Grid.Row="7" Grid.Column="0" Text="Route Type" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_RouteType" Grid.Row="7" Grid.Column="1" IsChecked="True" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_RouteType" Grid.Row="7" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom"/>
                                            <ComboBoxItem Content="Top" Tag="Top" IsSelected="True"/>
                                            <ComboBoxItem Content="None" Tag="None"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_RouteType" Grid.Row="7" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left"/>
                                            <ComboBoxItem Content="Center" Tag="Center"/>
                                            <ComboBoxItem Content="Right" Tag="Right" IsSelected="True"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_RouteType" Grid.Row="7" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1" IsSelected="True"/>
                                            <ComboBoxItem Content="2" Tag="2"/>
                                            <ComboBoxItem Content="3" Tag="3"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>

                                        <!-- Row 8: Waypoints -->
                                        <TextBlock Name="lblProp_Waypoints" Grid.Row="8" Grid.Column="0" Text="Intermediate Stops (Waypoints)" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_Waypoints" Grid.Row="8" Grid.Column="1" IsChecked="False" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_Waypoints" Grid.Row="8" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom" IsSelected="True"/>
                                            <ComboBoxItem Content="Top" Tag="Top"/>
                                            <ComboBoxItem Content="None" Tag="None"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_Waypoints" Grid.Row="8" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left" IsSelected="True"/>
                                            <ComboBoxItem Content="Center" Tag="Center"/>
                                            <ComboBoxItem Content="Right" Tag="Right"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_Waypoints" Grid.Row="8" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1"/>
                                            <ComboBoxItem Content="2" Tag="2" IsSelected="True"/>
                                            <ComboBoxItem Content="3" Tag="3"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>

                                        <!-- Row 9: StartRaw -->
                                        <TextBlock Name="lblProp_StartRaw" Grid.Row="9" Grid.Column="0" Text="Start Address (Raw Input)" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_StartRaw" Grid.Row="9" Grid.Column="1" IsChecked="False" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_StartRaw" Grid.Row="9" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom"/>
                                            <ComboBoxItem Content="Top" Tag="Top"/>
                                            <ComboBoxItem Content="None" Tag="None" IsSelected="True"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_StartRaw" Grid.Row="9" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left" IsSelected="True"/>
                                            <ComboBoxItem Content="Center" Tag="Center"/>
                                            <ComboBoxItem Content="Right" Tag="Right"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_StartRaw" Grid.Row="9" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1" IsSelected="True"/>
                                            <ComboBoxItem Content="2" Tag="2"/>
                                            <ComboBoxItem Content="3" Tag="3"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>

                                        <!-- Row 10: EndRaw -->
                                        <TextBlock Name="lblProp_EndRaw" Grid.Row="10" Grid.Column="0" Text="End Address (Raw Input)" Foreground="{DynamicResource TextPrimary}" FontSize="12" VerticalAlignment="Center" Margin="4,4"/>
                                        <CheckBox Name="chkProp_EndRaw" Grid.Row="10" Grid.Column="1" IsChecked="False" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        <ComboBox Name="cmbPanel_EndRaw" Grid.Row="10" Grid.Column="2" Margin="3,2">
                                            <ComboBoxItem Content="Bottom" Tag="Bottom"/>
                                            <ComboBoxItem Content="Top" Tag="Top"/>
                                            <ComboBoxItem Content="None" Tag="None" IsSelected="True"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbAlign_EndRaw" Grid.Row="10" Grid.Column="3" Margin="3,2">
                                            <ComboBoxItem Content="Left" Tag="Left" IsSelected="True"/>
                                            <ComboBoxItem Content="Center" Tag="Center"/>
                                            <ComboBoxItem Content="Right" Tag="Right"/>
                                        </ComboBox>
                                        <ComboBox Name="cmbOrder_EndRaw" Grid.Row="10" Grid.Column="4" Margin="3,2">
                                            <ComboBoxItem Content="1" Tag="1"/>
                                            <ComboBoxItem Content="2" Tag="2" IsSelected="True"/>
                                            <ComboBoxItem Content="3" Tag="3"/>
                                            <ComboBoxItem Content="4" Tag="4"/>
                                            <ComboBoxItem Content="5" Tag="5"/>
                                            <ComboBoxItem Content="6" Tag="6"/>
                                            <ComboBoxItem Content="7" Tag="7"/>
                                            <ComboBoxItem Content="8" Tag="8"/>
                                            <ComboBoxItem Content="9" Tag="9"/>
                                        </ComboBox>
                                    </Grid>
                                </Border>

                                <StackPanel Orientation="Horizontal" Margin="0,2,0,0">
                                    <Button Name="btnResetOverlayConfig" Content="🔄 Reset to Default Layout" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Padding="12,6"/>
                                </StackPanel>
                            </StackPanel>
                        </Border>

                        <!-- Language Card -->
                        <Border Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="16" Margin="0,0,0,14">
                            <StackPanel>
                                <TextBlock Name="lblSettingsLangHeader" Text="Language &amp; Localization" FontSize="16" FontWeight="Bold" Foreground="{DynamicResource TextPrimary}" Margin="0,0,0,8"/>
                                <TextBlock Name="lblSettingsLangLabel" Text="Application and Google Maps API Language:" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,6"/>
                                <ComboBox Name="cmbSettingsLanguage" Margin="0,0,0,10"/>
                                <StackPanel Orientation="Horizontal">
                                    <Button Name="btnOpenLangFile" Content="📂 Open Localization File (localization.json)" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Padding="10,6" Margin="0,0,8,0" ToolTip="Open the external localization file to edit or add new languages"/>
                                    <Button Name="btnReloadLang" Content="🔄 Reload Languages" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Padding="10,6" ToolTip="Reload language definitions from disk"/>
                                </StackPanel>
                            </StackPanel>
                        </Border>

                        <!-- Theme Card -->
                        <Border Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="16" Margin="0,0,0,14">
                            <StackPanel>
                                <TextBlock Name="lblSettingsThemeHeader" Text="Appearance &amp; Theme" FontSize="16" FontWeight="Bold" Foreground="{DynamicResource TextPrimary}" Margin="0,0,0,8"/>
                                <TextBlock Name="lblSettingsThemeLabel" Text="Application Theme (Color Scheme):" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,6"/>
                                <ComboBox Name="cmbSettingsTheme" Margin="0,0,0,0">
                                    <ComboBoxItem Content="🌙 Dark" Tag="Dark" IsSelected="True"/>
                                    <ComboBoxItem Content="☀️ Light" Tag="Light"/>
                                </ComboBox>
                            </StackPanel>
                        </Border>

                        <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                            <Button Name="btnSaveSettings" Content="💾 SAVE SETTINGS" Background="#10B981" Foreground="#FFFFFF" Padding="14,10" FontWeight="Bold" Width="200"/>
                            <Button Name="btnOpenLogFile" Content="📋 OPEN LOG FILE" Background="{DynamicResource BtnSecondaryBg}" Foreground="{DynamicResource BtnSecondaryFg}" Padding="14,10" FontWeight="SemiBold" Margin="10,0,0,0"/>
                        </StackPanel>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
        </TabControl>

        <!-- Footer -->
        <Border Grid.Row="2" Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="6" Padding="10,6" Margin="0,10,0,0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Name="lblFooterStatus" Text="Ready." Foreground="{DynamicResource TextSecondary}" FontSize="12" VerticalAlignment="Center"/>
                <TextBlock Name="lblFooterVersion" Grid.Column="1" Text="Google Maps Routes v2.1" Foreground="{DynamicResource TextSecondary}" FontSize="12" VerticalAlignment="Center"/>
            </Grid>
        </Border>
    </Grid>
</Window>
'@

    return $xaml
}
