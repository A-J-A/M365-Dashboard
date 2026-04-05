#Requires -Version 5.1
<#
.SYNOPSIS
    M365 Dashboard - Deployment Wizard (WPF GUI)
.DESCRIPTION
    Visual wizard for deploying the M365 Dashboard to Azure.
    Run via Start-Deployment.cmd or directly:
        powershell -ExecutionPolicy Bypass -File scripts\Deploy-Wizard.ps1
#>

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

# ---------------------------------------------------------------------------
# Colour tokens
# ---------------------------------------------------------------------------
$C = @{
    Bg       = "#0D1117"   # window background
    Panel    = "#161B22"   # card background
    Panel2   = "#1C2333"   # slightly lighter panel
    Border   = "#30363D"   # border
    Accent   = "#0078D4"   # Microsoft blue
    Accent2  = "#58A6FF"   # lighter blue
    Cyan     = "#39D353"   # success green
    Text     = "#E6EDF3"   # primary text
    Sub      = "#8B949E"   # secondary text
    Success  = "#3FB950"   # green
    Warn     = "#D29922"   # amber
    Error    = "#F85149"   # red
    Active   = "#1F6FEB"   # active step
    Done     = "#238636"   # done step
}

# ---------------------------------------------------------------------------
# XAML
# ---------------------------------------------------------------------------
[xml]$Xaml = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="M365 Dashboard — Deployment Wizard"
    Width="900" Height="660" MinWidth="800" MinHeight="580"
    WindowStartupLocation="CenterScreen"
    Background="#0D1117" Foreground="#E6EDF3"
    FontFamily="Segoe UI" FontSize="13">

  <Window.Resources>

    <!-- Button: Primary -->
    <Style x:Key="BtnPrimary" TargetType="Button">
      <Setter Property="Background"  Value="#0078D4"/>
      <Setter Property="Foreground"  Value="White"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding"     Value="22,9"/>
      <Setter Property="Cursor"      Value="Hand"/>
      <Setter Property="FontSize"    Value="13"/>
      <Setter Property="FontWeight"  Value="SemiBold"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bd" Background="{TemplateBinding Background}"
                    CornerRadius="5" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="Background" Value="#106EBE"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="bd" Property="Background" Value="#21262D"/>
                <Setter Property="Foreground" Value="#484F58"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- Button: Secondary -->
    <Style x:Key="BtnSecondary" TargetType="Button">
      <Setter Property="Background"  Value="Transparent"/>
      <Setter Property="Foreground"  Value="#58A6FF"/>
      <Setter Property="BorderBrush" Value="#30363D"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding"     Value="20,8"/>
      <Setter Property="Cursor"      Value="Hand"/>
      <Setter Property="FontSize"    Value="13"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bd" Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="5" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="Background" Value="#161B22"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Foreground" Value="#484F58"/>
                <Setter TargetName="bd" Property="BorderBrush" Value="#21262D"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- RadioButton: Mode card -->
    <Style x:Key="CardRadio" TargetType="RadioButton">
      <Setter Property="Background"      Value="#161B22"/>
      <Setter Property="BorderBrush"     Value="#30363D"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding"         Value="18,14"/>
      <Setter Property="Cursor"          Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="RadioButton">
            <Border x:Name="bd" Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="7" Padding="{TemplateBinding Padding}">
              <ContentPresenter/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="bd" Property="BorderBrush"  Value="#0078D4"/>
                <Setter TargetName="bd" Property="Background"   Value="#0D2137"/>
              </Trigger>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="Background" Value="#1C2333"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- TextBox -->
    <Style TargetType="TextBox">
      <Setter Property="Background"      Value="#0D1117"/>
      <Setter Property="Foreground"      Value="#E6EDF3"/>
      <Setter Property="BorderBrush"     Value="#30363D"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding"         Value="10,7"/>
      <Setter Property="CaretBrush"      Value="White"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TextBox">
            <Border Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="5">
              <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}"/>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- PasswordBox -->
    <Style TargetType="PasswordBox">
      <Setter Property="Background"      Value="#0D1117"/>
      <Setter Property="Foreground"      Value="#E6EDF3"/>
      <Setter Property="BorderBrush"     Value="#30363D"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding"         Value="10,7"/>
      <Setter Property="CaretBrush"      Value="White"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="PasswordBox">
            <Border Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="5">
              <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}"/>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- ComboBox -->
    <Style TargetType="ComboBox">
      <Setter Property="Background"      Value="#0D1117"/>
      <Setter Property="Foreground"      Value="#E6EDF3"/>
      <Setter Property="BorderBrush"     Value="#30363D"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding"         Value="10,7"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ComboBox">
            <Grid>
              <ToggleButton x:Name="ToggleButton" Focusable="false"
                            IsChecked="{Binding Path=IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"
                            ClickMode="Press">
                <ToggleButton.Template>
                  <ControlTemplate TargetType="ToggleButton">
                    <Border x:Name="border" Background="#0D1117" BorderBrush="#30363D" BorderThickness="1" CornerRadius="5">
                      <Grid>
                        <Grid.ColumnDefinitions>
                          <ColumnDefinition Width="*"/>
                          <ColumnDefinition Width="24"/>
                        </Grid.ColumnDefinitions>
                        <ContentPresenter Grid.Column="0" Margin="10,7" HorizontalAlignment="Left" VerticalAlignment="Center"/>
                        <Path Grid.Column="1" Data="M 0 0 L 4 4 L 8 0 Z" Fill="#8B949E"
                              HorizontalAlignment="Center" VerticalAlignment="Center"/>
                      </Grid>
                    </Border>
                    <ControlTemplate.Triggers>
                      <Trigger Property="IsMouseOver" Value="True">
                        <Setter TargetName="border" Property="BorderBrush" Value="#58A6FF"/>
                      </Trigger>
                    </ControlTemplate.Triggers>
                  </ControlTemplate>
                </ToggleButton.Template>
              </ToggleButton>
              <ContentPresenter x:Name="ContentSite" IsHitTestVisible="False"
                                Content="{TemplateBinding SelectionBoxItem}"
                                ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                                Margin="12,7,30,7" VerticalAlignment="Center"
                                HorizontalAlignment="Left"/>
              <Popup x:Name="Popup" Placement="Bottom" IsOpen="{TemplateBinding IsDropDownOpen}"
                     AllowsTransparency="True" Focusable="False" PopupAnimation="Slide">
                <Grid x:Name="DropDown" SnapsToDevicePixels="True"
                      MinWidth="{TemplateBinding ActualWidth}" MaxHeight="300">
                  <Border x:Name="DropDownBorder" Background="#161B22"
                          BorderBrush="#30363D" BorderThickness="1" CornerRadius="0,0,5,5"/>
                  <ScrollViewer Margin="1" SnapsToDevicePixels="True">
                    <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Contained"/>
                  </ScrollViewer>
                </Grid>
              </Popup>
            </Grid>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- ComboBoxItem -->
    <Style TargetType="ComboBoxItem">
      <Setter Property="Background"  Value="#161B22"/>
      <Setter Property="Foreground"  Value="#E6EDF3"/>
      <Setter Property="Padding"     Value="12,8"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ComboBoxItem">
            <Border x:Name="bd" Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}">
              <ContentPresenter/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsHighlighted" Value="True">
                <Setter TargetName="bd" Property="Background" Value="#0D2137"/>
                <Setter Property="Foreground" Value="#58A6FF"/>
              </Trigger>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="bd" Property="Background" Value="#0D2137"/>
                <Setter Property="Foreground" Value="#58A6FF"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- Separator -->
    <Style TargetType="Separator">
      <Setter Property="Background" Value="#30363D"/>
      <Setter Property="Height"     Value="1"/>
      <Setter Property="Margin"     Value="0,8"/>
    </Style>

  </Window.Resources>

  <Grid>
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="210"/>
      <ColumnDefinition Width="*"/>
    </Grid.ColumnDefinitions>

    <!-- ═══════════════════════════════════════════════
         LEFT SIDEBAR
    ═══════════════════════════════════════════════ -->
    <Border Grid.Column="0" Background="#0D1117"
            BorderBrush="#30363D" BorderThickness="0,0,1,0">
      <DockPanel>

        <!-- Logo block -->
        <StackPanel DockPanel.Dock="Top" Margin="22,28,22,20">
          <!-- M365 grid icon -->
          <Border Width="46" Height="46" CornerRadius="10"
                  Background="#0078D4" HorizontalAlignment="Left" Margin="0,0,0,14">
            <Grid>
              <Rectangle Fill="White" Width="9" Height="9" RadiusX="2" RadiusY="2"
                         HorizontalAlignment="Left" VerticalAlignment="Top" Margin="7,7,0,0"/>
              <Rectangle Fill="White" Width="9" Height="9" RadiusX="2" RadiusY="2"
                         HorizontalAlignment="Right" VerticalAlignment="Top" Margin="0,7,7,0"/>
              <Rectangle Fill="White" Width="9" Height="9" RadiusX="2" RadiusY="2"
                         HorizontalAlignment="Left" VerticalAlignment="Bottom" Margin="7,0,0,7"/>
              <Rectangle Fill="White" Width="9" Height="9" RadiusX="2" RadiusY="2"
                         HorizontalAlignment="Right" VerticalAlignment="Bottom" Margin="0,0,7,7"/>
            </Grid>
          </Border>
          <TextBlock Text="M365 Dashboard" FontSize="15" FontWeight="Bold" Foreground="White"/>
          <TextBlock Text="Deployment Wizard" FontSize="11" Foreground="#8B949E" Margin="0,3,0,0"/>
        </StackPanel>

        <Rectangle DockPanel.Dock="Top" Height="1" Fill="#30363D" Margin="18,0"/>

        <!-- Step list -->
        <StackPanel DockPanel.Dock="Top" Margin="0,18,0,0">
          <!-- Step 1 -->
          <Grid x:Name="SideStep1" Margin="14,4">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="30"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Border x:Name="SideDot1" Width="24" Height="24" CornerRadius="12"
                    Background="#0078D4" HorizontalAlignment="Center" VerticalAlignment="Center">
              <TextBlock x:Name="SideNum1" Text="1" Foreground="White"
                         FontWeight="Bold" FontSize="11"
                         HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <StackPanel Grid.Column="1" Margin="10,0,0,0" VerticalAlignment="Center">
              <TextBlock x:Name="SideLbl1" Text="Welcome" Foreground="White"
                         FontWeight="SemiBold" FontSize="12"/>
              <TextBlock Text="Prerequisites &amp; mode" Foreground="#8B949E" FontSize="10"/>
            </StackPanel>
          </Grid>
          <!-- Step 2 -->
          <Grid x:Name="SideStep2" Margin="14,4">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="30"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Border x:Name="SideDot2" Width="24" Height="24" CornerRadius="12"
                    Background="#30363D" HorizontalAlignment="Center" VerticalAlignment="Center">
              <TextBlock x:Name="SideNum2" Text="2" Foreground="#8B949E"
                         FontWeight="Bold" FontSize="11"
                         HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <StackPanel Grid.Column="1" Margin="10,0,0,0" VerticalAlignment="Center">
              <TextBlock x:Name="SideLbl2" Text="Configuration" Foreground="#8B949E"
                         FontWeight="Normal" FontSize="12"/>
              <TextBlock Text="Resources &amp; credentials" Foreground="#484F58" FontSize="10"/>
            </StackPanel>
          </Grid>
          <!-- Step 3 -->
          <Grid x:Name="SideStep3" Margin="14,4">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="30"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Border x:Name="SideDot3" Width="24" Height="24" CornerRadius="12"
                    Background="#30363D" HorizontalAlignment="Center" VerticalAlignment="Center">
              <TextBlock x:Name="SideNum3" Text="3" Foreground="#8B949E"
                         FontWeight="Bold" FontSize="11"
                         HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <StackPanel Grid.Column="1" Margin="10,0,0,0" VerticalAlignment="Center">
              <TextBlock x:Name="SideLbl3" Text="Review" Foreground="#8B949E"
                         FontWeight="Normal" FontSize="12"/>
              <TextBlock Text="Confirm settings" Foreground="#484F58" FontSize="10"/>
            </StackPanel>
          </Grid>
          <!-- Step 4 -->
          <Grid x:Name="SideStep4" Margin="14,4">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="30"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Border x:Name="SideDot4" Width="24" Height="24" CornerRadius="12"
                    Background="#30363D" HorizontalAlignment="Center" VerticalAlignment="Center">
              <TextBlock x:Name="SideNum4" Text="4" Foreground="#8B949E"
                         FontWeight="Bold" FontSize="11"
                         HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <StackPanel Grid.Column="1" Margin="10,0,0,0" VerticalAlignment="Center">
              <TextBlock x:Name="SideLbl4" Text="Deploying" Foreground="#8B949E"
                         FontWeight="Normal" FontSize="12"/>
              <TextBlock Text="Live progress" Foreground="#484F58" FontSize="10"/>
            </StackPanel>
          </Grid>
          <!-- Step 5 -->
          <Grid x:Name="SideStep5" Margin="14,4">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="30"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Border x:Name="SideDot5" Width="24" Height="24" CornerRadius="12"
                    Background="#30363D" HorizontalAlignment="Center" VerticalAlignment="Center">
              <TextBlock x:Name="SideNum5" Text="5" Foreground="#8B949E"
                         FontWeight="Bold" FontSize="11"
                         HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <StackPanel Grid.Column="1" Margin="10,0,0,0" VerticalAlignment="Center">
              <TextBlock x:Name="SideLbl5" Text="Complete" Foreground="#8B949E"
                         FontWeight="Normal" FontSize="12"/>
              <TextBlock Text="Next steps" Foreground="#484F58" FontSize="10"/>
            </StackPanel>
          </Grid>
        </StackPanel>

        <!-- Version at bottom -->
        <StackPanel DockPanel.Dock="Bottom" Margin="22,0,22,18">
          <Rectangle Height="1" Fill="#30363D" Margin="0,0,0,12"/>
          <TextBlock x:Name="TxtVersion" Text="Version checking..." Foreground="#484F58" FontSize="10"/>
          <TextBlock Text="github.com" Foreground="#484F58" FontSize="10" Margin="0,2,0,0"/>
        </StackPanel>

        <Grid/>
      </DockPanel>
    </Border>

    <!-- ═══════════════════════════════════════════════
         RIGHT CONTENT AREA
    ═══════════════════════════════════════════════ -->
    <Grid Grid.Column="1">
      <Grid.RowDefinitions>
        <RowDefinition Height="*"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>

      <!-- Page host (ScrollViewer wraps all pages) -->
      <ScrollViewer Grid.Row="0" VerticalScrollBarVisibility="Auto"
                    HorizontalScrollBarVisibility="Disabled" Padding="0,0,4,0">
        <Grid>

          <!-- ─────────────────────────────────────────
               PAGE 1 — WELCOME
          ───────────────────────────────────────── -->
          <StackPanel x:Name="PageWelcome" Margin="40,36,40,24" Visibility="Visible">
            <TextBlock Text="Welcome" FontSize="28" FontWeight="Bold" Foreground="White" Margin="0,0,0,6"/>
            <TextBlock TextWrapping="Wrap" Foreground="#8B949E" Margin="0,0,0,28">
              This wizard will deploy the M365 Dashboard to Azure Container Apps.
              Deployment takes approximately 10–15 minutes and requires an Azure subscription
              and Global Administrator access to your Microsoft 365 tenant.
            </TextBlock>

            <!-- Prerequisites card -->
            <TextBlock Text="P R E R E Q U I S I T E S" FontSize="10" FontWeight="Bold"
                       Foreground="#8B949E" Margin="0,0,0,10"/>
            <Border Background="#161B22" BorderBrush="#30363D" BorderThickness="1"
                    CornerRadius="8" Padding="20,16" Margin="0,0,0,28">
              <StackPanel>
                <!-- az -->
                <Grid Margin="0,5">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="26"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                  </Grid.ColumnDefinitions>
                  <TextBlock x:Name="PIconAz" Text="●" Foreground="#D29922" FontSize="10"
                             VerticalAlignment="Center"/>
                  <StackPanel Grid.Column="1" VerticalAlignment="Center">
                    <TextBlock Text="Azure CLI" Foreground="#E6EDF3" FontWeight="SemiBold"/>
                    <TextBlock x:Name="PTextAz" Text="Checking..." Foreground="#8B949E" FontSize="11"/>
                  </StackPanel>
                  <TextBlock x:Name="PBadgeAz" Grid.Column="2" Text="Checking"
                             Foreground="#D29922" FontSize="11" VerticalAlignment="Center"/>
                </Grid>
                <Rectangle Height="1" Fill="#21262D"/>
                <!-- git -->
                <Grid Margin="0,5">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="26"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                  </Grid.ColumnDefinitions>
                  <TextBlock x:Name="PIconGit" Text="●" Foreground="#D29922" FontSize="10"
                             VerticalAlignment="Center"/>
                  <StackPanel Grid.Column="1" VerticalAlignment="Center">
                    <TextBlock Text="Git" Foreground="#E6EDF3" FontWeight="SemiBold"/>
                    <TextBlock x:Name="PTextGit" Text="Checking..." Foreground="#8B949E" FontSize="11"/>
                  </StackPanel>
                  <TextBlock x:Name="PBadgeGit" Grid.Column="2" Text="Checking"
                             Foreground="#D29922" FontSize="11" VerticalAlignment="Center"/>
                </Grid>
                <Rectangle Height="1" Fill="#21262D"/>
                <!-- gh cli -->
                <Grid Margin="0,5">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="26"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                  </Grid.ColumnDefinitions>
                  <TextBlock x:Name="PIconGh" Text="●" Foreground="#D29922" FontSize="10"
                             VerticalAlignment="Center"/>
                  <StackPanel Grid.Column="1" VerticalAlignment="Center">
                    <TextBlock Text="GitHub CLI (gh)" Foreground="#E6EDF3" FontWeight="SemiBold"/>
                    <TextBlock x:Name="PTextGh" Text="Checking..." Foreground="#8B949E" FontSize="11"/>
                  </StackPanel>
                  <TextBlock x:Name="PBadgeGh" Grid.Column="2" Text="Checking"
                             Foreground="#D29922" FontSize="11" VerticalAlignment="Center"/>
                </Grid>
                <Rectangle Height="1" Fill="#21262D"/>
                <!-- repo -->
                <Grid Margin="0,5">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="26"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                  </Grid.ColumnDefinitions>
                  <TextBlock x:Name="PIconRepo" Text="●" Foreground="#D29922" FontSize="10"
                             VerticalAlignment="Center"/>
                  <StackPanel Grid.Column="1" VerticalAlignment="Center">
                    <TextBlock Text="GitHub Repository" Foreground="#E6EDF3" FontWeight="SemiBold"/>
                    <TextBlock x:Name="PTextRepo" Text="Checking..." Foreground="#8B949E" FontSize="11"/>
                  </StackPanel>
                  <TextBlock x:Name="PBadgeRepo" Grid.Column="2" Text="Checking"
                             Foreground="#D29922" FontSize="11" VerticalAlignment="Center"/>
                </Grid>
              </StackPanel>
            </Border>

            <!-- Deployment mode -->
            <TextBlock Text="D E P L O Y M E N T   M O D E" FontSize="10" FontWeight="Bold"
                       Foreground="#8B949E" Margin="0,0,0,10"/>
            <Grid Margin="0,0,0,8">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="14"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>

              <RadioButton x:Name="ModeStandard" Grid.Column="0"
                           Style="{StaticResource CardRadio}" IsChecked="True" GroupName="Mode">
                <StackPanel>
                  <TextBlock Text="🏢" FontSize="26" Margin="0,0,0,10"/>
                  <TextBlock Text="Standard" FontWeight="Bold" FontSize="14" Foreground="White" Margin="0,0,0,6"/>
                  <TextBlock TextWrapping="Wrap" Foreground="#8B949E" FontSize="12">
                    App registration and Azure resources in the same tenant.
                    Best for single-organisation deployments.
                  </TextBlock>
                </StackPanel>
              </RadioButton>

              <RadioButton x:Name="ModeMsp" Grid.Column="2"
                           Style="{StaticResource CardRadio}" GroupName="Mode">
                <StackPanel>
                  <TextBlock Text="🔗" FontSize="26" Margin="0,0,0,10"/>
                  <TextBlock Text="MSP / Multi-tenant" FontWeight="Bold" FontSize="14" Foreground="White" Margin="0,0,0,6"/>
                  <TextBlock TextWrapping="Wrap" Foreground="#8B949E" FontSize="12">
                    App registration in the client tenant, Azure resources
                    in your MSP subscription.
                  </TextBlock>
                </StackPanel>
              </RadioButton>
            </Grid>
          </StackPanel>

          <!-- ─────────────────────────────────────────
               PAGE 2 — CONFIGURATION
          ───────────────────────────────────────── -->
          <StackPanel x:Name="PageConfig" Margin="40,36,40,24" Visibility="Collapsed">
            <TextBlock Text="Configuration" FontSize="28" FontWeight="Bold" Foreground="White" Margin="0,0,0,6"/>
            <TextBlock TextWrapping="Wrap" Foreground="#8B949E" Margin="0,0,0,28">
              Provide details for your Azure deployment. These settings will be used to create
              and name all Azure resources.
            </TextBlock>

            <!-- Resource prefix -->
            <TextBlock Text="RESOURCE NAME PREFIX" FontSize="10" FontWeight="Bold"
                       Foreground="#8B949E" Margin="0,0,0,8"/>
            <TextBox x:Name="TxtPrefix" Margin="0,0,0,4"/>
            <TextBlock TextWrapping="Wrap" Foreground="#8B949E" FontSize="11" Margin="0,0,0,20">
              3–12 characters, letters and numbers only, must start with a letter.
              Used to name all Azure resources (e.g. myorg → myorg-prod-app).
            </TextBlock>

            <!-- Azure subscription -->
            <TextBlock Text="AZURE SUBSCRIPTION" FontSize="10" FontWeight="Bold"
                       Foreground="#8B949E" Margin="0,0,0,8"/>
            <ComboBox x:Name="CmbSubscription" Margin="0,0,0,4"/>
            <TextBlock x:Name="TxtSubNote" TextWrapping="Wrap" Foreground="#8B949E" FontSize="11" Margin="0,0,0,20">
              Loading subscriptions...
            </TextBlock>

            <!-- Azure region -->
            <TextBlock Text="AZURE REGION" FontSize="10" FontWeight="Bold"
                       Foreground="#8B949E" Margin="0,0,0,8"/>
            <ComboBox x:Name="CmbRegion" Margin="0,0,0,20">
              <ComboBoxItem Content="UK South"                     Tag="uksouth"       IsSelected="True"/>
              <ComboBoxItem Content="UK West"                      Tag="ukwest"/>
              <ComboBoxItem Content="North Europe (Ireland)"       Tag="northeurope"/>
              <ComboBoxItem Content="West Europe (Netherlands)"    Tag="westeurope"/>
              <ComboBoxItem Content="East US"                      Tag="eastus"/>
              <ComboBoxItem Content="East US 2"                    Tag="eastus2"/>
              <ComboBoxItem Content="West US 2"                    Tag="westus2"/>
              <ComboBoxItem Content="Australia East"               Tag="australiaeast"/>
              <ComboBoxItem Content="Southeast Asia (Singapore)"   Tag="southeastasia"/>
              <ComboBoxItem Content="Japan East"                   Tag="japaneast"/>
            </ComboBox>

            <!-- Credential type -->
            <TextBlock Text="APP CREDENTIAL TYPE" FontSize="10" FontWeight="Bold"
                       Foreground="#8B949E" Margin="0,0,0,8"/>
            <Grid Margin="0,0,0,20">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="14"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>
              <RadioButton x:Name="CredSecret" Grid.Column="0"
                           Style="{StaticResource CardRadio}" IsChecked="True" GroupName="Cred">
                <StackPanel>
                  <TextBlock Text="🔑" FontSize="20" Margin="0,0,0,8"/>
                  <TextBlock Text="Client Secret" FontWeight="SemiBold" Foreground="White" FontSize="13" Margin="0,0,0,6"/>
                  <TextBlock TextWrapping="Wrap" Foreground="#8B949E" FontSize="11">
                    Simpler setup. May be blocked if your tenant restricts app secrets.
                  </TextBlock>
                </StackPanel>
              </RadioButton>
              <RadioButton x:Name="CredCert" Grid.Column="2"
                           Style="{StaticResource CardRadio}" GroupName="Cred">
                <StackPanel>
                  <TextBlock Text="📜" FontSize="20" Margin="0,0,0,8"/>
                  <TextBlock Text="Certificate" FontWeight="SemiBold" Foreground="White" FontSize="13" Margin="0,0,0,6"/>
                  <TextBlock TextWrapping="Wrap" Foreground="#8B949E" FontSize="11">
                    More secure. Works even when client secrets are blocked by policy.
                  </TextBlock>
                </StackPanel>
              </RadioButton>
            </Grid>

            <!-- Login accounts (optional hints) -->
            <TextBlock Text="LOGIN ACCOUNTS  (optional)" FontSize="10" FontWeight="Bold"
                       Foreground="#8B949E" Margin="0,0,0,8"/>
            <Border Background="#161B22" BorderBrush="#30363D" BorderThickness="1"
                    CornerRadius="8" Padding="16,14" Margin="0,0,0,20">
              <StackPanel>
                <TextBlock TextWrapping="Wrap" Foreground="#8B949E" FontSize="11" Margin="0,0,0,12">
                  Enter the accounts you will use to log in. These are shown in the login prompts
                  so you pick the right account each time.
                </TextBlock>

                <!-- Client tenant — MSP only -->
                <StackPanel x:Name="PanelClientUser">
                  <TextBlock Text="Client tenant admin" Foreground="#58A6FF" FontSize="11"
                             FontWeight="SemiBold" Margin="0,0,0,4"/>
                  <TextBox x:Name="TxtClientUser" Margin="0,0,0,12"
                           Tag="e.g. admin@clientdomain.com"/>
                </StackPanel>

                <!-- Your Azure account -->
                <TextBlock Text="Your Azure account" Foreground="#3FB950" FontSize="11"
                           FontWeight="SemiBold" Margin="0,0,0,4"/>
                <TextBox x:Name="TxtAzureUser" Margin="0,0,0,12"
                         Tag="e.g. you@yourdomain.com"/>

                <!-- GitHub username -->
                <TextBlock Text="GitHub username" Foreground="#D29922" FontSize="11"
                           FontWeight="SemiBold" Margin="0,0,0,4"/>
                <TextBox x:Name="TxtGitHubUser" Margin="0,0,0,0"
                         Tag="e.g. your-github-username"/>
              </StackPanel>
            </Border>

            <!-- SQL Password -->
            <TextBlock Text="SQL ADMIN PASSWORD" FontSize="10" FontWeight="Bold"
                       Foreground="#8B949E" Margin="0,0,0,8"/>
            <TextBlock TextWrapping="Wrap" Foreground="#8B949E" FontSize="11" Margin="0,0,0,8">
              Used for the Azure SQL database. Min 12 chars with uppercase, lowercase,
              number and special character. Store this securely — it will not be shown again.
            </TextBlock>
            <TextBlock Text="Password" Foreground="#8B949E" FontSize="11" Margin="0,0,0,4"/>
            <PasswordBox x:Name="TxtPwd1" Margin="0,0,0,10"/>
            <TextBlock Text="Confirm Password" Foreground="#8B949E" FontSize="11" Margin="0,0,0,4"/>
            <PasswordBox x:Name="TxtPwd2" Margin="0,0,0,4"/>
            <TextBlock x:Name="TxtPwdErr" Foreground="#F85149" FontSize="11"
                       Margin="0,0,0,0" Visibility="Collapsed" TextWrapping="Wrap"/>
          </StackPanel>

          <!-- ─────────────────────────────────────────
               PAGE 3 — REVIEW
          ───────────────────────────────────────── -->
          <StackPanel x:Name="PageReview" Margin="40,36,40,24" Visibility="Collapsed">
            <TextBlock Text="Review &amp; Deploy" FontSize="28" FontWeight="Bold"
                       Foreground="White" Margin="0,0,0,6"/>
            <TextBlock TextWrapping="Wrap" Foreground="#8B949E" Margin="0,0,0,28">
              Review your settings before deployment begins. Click Deploy to start.
            </TextBlock>

            <!-- Summary card -->
            <Border Background="#161B22" BorderBrush="#30363D" BorderThickness="1"
                    CornerRadius="8" Padding="24,20" Margin="0,0,0,20">
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="160"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <!-- Row: Mode -->
                <TextBlock Grid.Row="0" Grid.Column="0" Text="Deployment Mode"
                           Foreground="#8B949E" Margin="0,0,0,14" VerticalAlignment="Center"/>
                <StackPanel Grid.Row="0" Grid.Column="1" Orientation="Horizontal" Margin="0,0,0,14">
                  <Border Background="#0D2137" BorderBrush="#0078D4" BorderThickness="1"
                          CornerRadius="4" Padding="8,3">
                    <TextBlock x:Name="RevMode" Text="" Foreground="#58A6FF" FontWeight="SemiBold" FontSize="12"/>
                  </Border>
                </StackPanel>

                <Rectangle Grid.Row="1" Grid.ColumnSpan="2" Height="1" Fill="#21262D" Margin="0,0,0,14"/>

                <!-- Row: Subscription -->
                <TextBlock Grid.Row="2" Grid.Column="0" Text="Subscription"
                           Foreground="#8B949E" Margin="0,0,0,14" VerticalAlignment="Center"/>
                <TextBlock x:Name="RevSub" Grid.Row="2" Grid.Column="1"
                           Foreground="White" FontWeight="SemiBold"
                           Margin="0,0,0,14" VerticalAlignment="Center" TextWrapping="Wrap"/>

                <!-- Row: Prefix -->
                <TextBlock Grid.Row="3" Grid.Column="0" Text="Resource Prefix"
                           Foreground="#8B949E" Margin="0,0,0,14" VerticalAlignment="Center"/>
                <TextBlock x:Name="RevPrefix" Grid.Row="3" Grid.Column="1"
                           Foreground="White" FontWeight="SemiBold" FontFamily="Consolas"
                           Margin="0,0,0,14" VerticalAlignment="Center"/>

                <!-- Row: Region -->
                <TextBlock Grid.Row="4" Grid.Column="0" Text="Azure Region"
                           Foreground="#8B949E" Margin="0,0,0,14" VerticalAlignment="Center"/>
                <TextBlock x:Name="RevRegion" Grid.Row="4" Grid.Column="1"
                           Foreground="White" FontWeight="SemiBold"
                           Margin="0,0,0,14" VerticalAlignment="Center"/>

                <!-- Row: Cred -->
                <TextBlock Grid.Row="5" Grid.Column="0" Text="Credential Type"
                           Foreground="#8B949E" Margin="0,0,0,14" VerticalAlignment="Center"/>
                <TextBlock x:Name="RevCred" Grid.Row="5" Grid.Column="1"
                           Foreground="White" FontWeight="SemiBold"
                           Margin="0,0,0,14" VerticalAlignment="Center"/>

                <!-- Row: Repo -->
                <TextBlock Grid.Row="6" Grid.Column="0" Text="GitHub Repo"
                           Foreground="#8B949E" VerticalAlignment="Center"/>
                <TextBlock x:Name="RevRepo" Grid.Row="6" Grid.Column="1"
                           Foreground="White" FontWeight="SemiBold"
                           FontFamily="Consolas" VerticalAlignment="Center"/>
              </Grid>
            </Border>

            <!-- Info banner -->
            <Border Background="#0D1F0D" BorderBrush="#238636" BorderThickness="1"
                    CornerRadius="7" Padding="16,12">
              <StackPanel Orientation="Horizontal">
                <TextBlock Text="ℹ" Foreground="#3FB950" FontSize="16" Margin="0,0,12,0" VerticalAlignment="Top"/>
                <TextBlock TextWrapping="Wrap" Foreground="#7EE787" FontSize="12">
                  Clicking Deploy will open Azure login prompts and begin the deployment.
                  Do not close this window during deployment. The process takes 10–15 minutes.
                </TextBlock>
              </StackPanel>
            </Border>
          </StackPanel>

          <!-- ─────────────────────────────────────────
               PAGE 4 — DEPLOYING
          ───────────────────────────────────────── -->
          <StackPanel x:Name="PageDeploy" Margin="40,36,40,24" Visibility="Collapsed">
            <TextBlock Text="Deploying..." FontSize="28" FontWeight="Bold"
                       Foreground="White" Margin="0,0,0,6"/>
            <TextBlock TextWrapping="Wrap" Foreground="#8B949E" Margin="0,0,0,20">
              Deployment is running. Do not close this window.
            </TextBlock>

            <!-- Animated progress bar -->
            <Grid Margin="0,0,0,20">
              <ProgressBar x:Name="PBar" Height="8" Minimum="0" Maximum="100" Value="0"
                           Background="#21262D" Foreground="#0078D4" BorderThickness="0"/>
              <Border Height="8" CornerRadius="4" Background="Transparent"
                      BorderBrush="#30363D" BorderThickness="1"/>
            </Grid>

            <!-- Step status checklist -->
            <Border Background="#161B22" BorderBrush="#30363D" BorderThickness="1"
                    CornerRadius="8" Padding="20,16" Margin="0,0,0,16">
              <StackPanel>
                <Grid x:Name="DS1" Margin="0,6">
                  <Grid.ColumnDefinitions><ColumnDefinition Width="28"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                  <TextBlock x:Name="DI1" Text="○" Foreground="#484F58" FontSize="16" VerticalAlignment="Center"/>
                  <TextBlock x:Name="DT1" Grid.Column="1" Text="Azure login" Foreground="#8B949E" VerticalAlignment="Center"/>
                </Grid>
                <Rectangle Height="1" Fill="#21262D"/>
                <Grid x:Name="DS2" Margin="0,6">
                  <Grid.ColumnDefinitions><ColumnDefinition Width="28"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                  <TextBlock x:Name="DI2" Text="○" Foreground="#484F58" FontSize="16" VerticalAlignment="Center"/>
                  <TextBlock x:Name="DT2" Grid.Column="1" Text="Create Entra ID app registration" Foreground="#8B949E" VerticalAlignment="Center"/>
                </Grid>
                <Rectangle Height="1" Fill="#21262D"/>
                <Grid x:Name="DS3" Margin="0,6">
                  <Grid.ColumnDefinitions><ColumnDefinition Width="28"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                  <TextBlock x:Name="DI3" Text="○" Foreground="#484F58" FontSize="16" VerticalAlignment="Center"/>
                  <TextBlock x:Name="DT3" Grid.Column="1" Text="Deploy Azure infrastructure (Container App, SQL, Key Vault, ACR)" Foreground="#8B949E" VerticalAlignment="Center"/>
                </Grid>
                <Rectangle Height="1" Fill="#21262D"/>
                <Grid x:Name="DS4" Margin="0,6">
                  <Grid.ColumnDefinitions><ColumnDefinition Width="28"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                  <TextBlock x:Name="DI4" Text="○" Foreground="#484F58" FontSize="16" VerticalAlignment="Center"/>
                  <TextBlock x:Name="DT4" Grid.Column="1" Text="Build and push Docker image to registry" Foreground="#8B949E" VerticalAlignment="Center"/>
                </Grid>
                <Rectangle Height="1" Fill="#21262D"/>
                <Grid x:Name="DS5" Margin="0,6">
                  <Grid.ColumnDefinitions><ColumnDefinition Width="28"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                  <TextBlock x:Name="DI5" Text="○" Foreground="#484F58" FontSize="16" VerticalAlignment="Center"/>
                  <TextBlock x:Name="DT5" Grid.Column="1" Text="Configure app registration (redirect URI, admin consent, logo)" Foreground="#8B949E" VerticalAlignment="Center"/>
                </Grid>
                <Rectangle Height="1" Fill="#21262D"/>
                <Grid x:Name="DS6" Margin="0,6">
                  <Grid.ColumnDefinitions><ColumnDefinition Width="28"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                  <TextBlock x:Name="DI6" Text="○" Foreground="#484F58" FontSize="16" VerticalAlignment="Center"/>
                  <TextBlock x:Name="DT6" Grid.Column="1" Text="Set GitHub Actions secrets for CI/CD" Foreground="#8B949E" VerticalAlignment="Center"/>
                </Grid>
              </StackPanel>
            </Border>

            <!-- Log output -->
            <TextBlock Text="DEPLOYMENT LOG" FontSize="10" FontWeight="Bold"
                       Foreground="#8B949E" Margin="0,0,0,8"/>
            <Border Background="#0D1117" BorderBrush="#30363D" BorderThickness="1" CornerRadius="7">
              <ScrollViewer x:Name="LogScroll" Height="200"
                            VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
                <TextBlock x:Name="LogBox" FontFamily="Consolas" FontSize="11.5"
                           Foreground="#58A6FF" Padding="12" TextWrapping="Wrap"/>
              </ScrollViewer>
            </Border>
          </StackPanel>

          <!-- ─────────────────────────────────────────
               PAGE 5 — COMPLETE
          ───────────────────────────────────────── -->
          <StackPanel x:Name="PageDone" Margin="40,36,40,24" Visibility="Collapsed">

            <!-- Success state -->
            <StackPanel x:Name="PanelSuccess" Visibility="Collapsed">
              <Border Background="#0D1F0D" BorderBrush="#238636" BorderThickness="1"
                      CornerRadius="10" Padding="28,24" Margin="0,0,0,24">
                <StackPanel HorizontalAlignment="Center">
                  <TextBlock Text="✓" FontSize="52" Foreground="#3FB950"
                             HorizontalAlignment="Center" Margin="0,0,0,10"/>
                  <TextBlock Text="Deployment Complete!" FontSize="22" FontWeight="Bold"
                             Foreground="White" HorizontalAlignment="Center" Margin="0,0,0,6"/>
                  <TextBlock Text="Your M365 Dashboard is live and ready to use."
                             Foreground="#7EE787" HorizontalAlignment="Center"/>
                </StackPanel>
              </Border>

              <!-- Dashboard URL -->
              <TextBlock Text="DASHBOARD URL" FontSize="10" FontWeight="Bold"
                         Foreground="#8B949E" Margin="0,0,0,8"/>
              <Border Background="#161B22" BorderBrush="#30363D" BorderThickness="1"
                      CornerRadius="7" Padding="16,12" Margin="0,0,0,8">
                <Grid>
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                  </Grid.ColumnDefinitions>
                  <TextBlock x:Name="TxtUrl" Foreground="#58A6FF" FontFamily="Consolas"
                             FontSize="12" VerticalAlignment="Center" TextWrapping="Wrap"/>
                  <Button x:Name="BtnCopyUrl" Grid.Column="1" Content="Copy"
                          Style="{StaticResource BtnSecondary}" Padding="12,6" Margin="10,0,0,0"/>
                </Grid>
              </Border>
              <Button x:Name="BtnOpenUrl" Content="🌐  Open Dashboard in Browser"
                      Style="{StaticResource BtnPrimary}" HorizontalAlignment="Left"
                      Margin="0,0,0,24"/>

              <!-- Next steps -->
              <TextBlock Text="MANUAL STEPS REQUIRED" FontSize="10" FontWeight="Bold"
                         Foreground="#8B949E" Margin="0,0,0,10"/>
              <Border Background="#161B22" BorderBrush="#30363D" BorderThickness="1"
                      CornerRadius="8" Padding="20,16">
                <StackPanel>
                  <StackPanel Orientation="Horizontal" Margin="0,0,0,12">
                    <Border Background="#0D2137" Width="26" Height="26" CornerRadius="13" Margin="0,0,12,0">
                      <TextBlock Text="1" Foreground="#58A6FF" FontWeight="Bold" FontSize="12"
                                 HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Border>
                    <StackPanel>
                      <TextBlock Text="Grant Admin Consent" Foreground="White" FontWeight="SemiBold"/>
                      <TextBlock x:Name="TxtConsentUrl" Foreground="#8B949E" FontSize="11"
                                 TextWrapping="Wrap"/>
                    </StackPanel>
                  </StackPanel>
                  <Rectangle Height="1" Fill="#21262D" Margin="0,0,0,12"/>
                  <StackPanel Orientation="Horizontal" Margin="0,0,0,12">
                    <Border Background="#0D2137" Width="26" Height="26" CornerRadius="13" Margin="0,0,12,0">
                      <TextBlock Text="2" Foreground="#58A6FF" FontWeight="Bold" FontSize="12"
                                 HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Border>
                    <StackPanel>
                      <TextBlock Text="Exchange Security Reader" Foreground="White" FontWeight="SemiBold"/>
                      <TextBlock Foreground="#8B949E" FontSize="11" TextWrapping="Wrap">
                        Exchange Admin Centre → Roles → View-Only Organization Management → Members → Add app registration
                      </TextBlock>
                    </StackPanel>
                  </StackPanel>
                  <Rectangle Height="1" Fill="#21262D" Margin="0,0,0,12"/>
                  <StackPanel Orientation="Horizontal">
                    <Border Background="#0D2137" Width="26" Height="26" CornerRadius="13" Margin="0,0,12,0">
                      <TextBlock Text="3" Foreground="#58A6FF" FontWeight="Bold" FontSize="12"
                                 HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Border>
                    <StackPanel>
                      <TextBlock Text="Assign Dashboard Admin Role" Foreground="White" FontWeight="SemiBold"/>
                      <TextBlock Foreground="#8B949E" FontSize="11" TextWrapping="Wrap">
                        Entra ID → Enterprise Applications → M365 Dashboard → Users and groups → Assign Dashboard Admin
                      </TextBlock>
                    </StackPanel>
                  </StackPanel>
                </StackPanel>
              </Border>
            </StackPanel>

            <!-- Error state -->
            <StackPanel x:Name="PanelError" Visibility="Collapsed">
              <Border Background="#1F0D0D" BorderBrush="#F85149" BorderThickness="1"
                      CornerRadius="10" Padding="28,24" Margin="0,0,0,24">
                <StackPanel HorizontalAlignment="Center">
                  <TextBlock Text="✗" FontSize="52" Foreground="#F85149"
                             HorizontalAlignment="Center" Margin="0,0,0,10"/>
                  <TextBlock Text="Deployment Failed" FontSize="22" FontWeight="Bold"
                             Foreground="White" HorizontalAlignment="Center" Margin="0,0,0,6"/>
                  <TextBlock Text="Check the deployment log for details. You can retry after fixing the issue."
                             Foreground="#FDA29B" HorizontalAlignment="Center" TextWrapping="Wrap"
                             TextAlignment="Center"/>
                </StackPanel>
              </Border>
              <Button x:Name="BtnRetry" Content="← Back to Review and Retry"
                      Style="{StaticResource BtnSecondary}" HorizontalAlignment="Left"/>
            </StackPanel>

          </StackPanel>

        </Grid>
      </ScrollViewer>

      <!-- ─────────────────────────────────────────
           FOOTER NAV BAR
      ───────────────────────────────────────── -->
      <Border Grid.Row="1" Background="#161B22" BorderBrush="#30363D"
              BorderThickness="0,1,0,0" Padding="28,14">
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="12"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <TextBlock x:Name="FooterMsg" Grid.Column="0" Foreground="#8B949E"
                     VerticalAlignment="Center" FontSize="12" TextWrapping="Wrap"/>
          <Button x:Name="BtnBack" Grid.Column="1" Content="← Back"
                  Style="{StaticResource BtnSecondary}" IsEnabled="False"/>
          <Button x:Name="BtnNext" Grid.Column="3" Content="Next →"
                  Style="{StaticResource BtnPrimary}"/>
        </Grid>
      </Border>

    </Grid>
  </Grid>
</Window>
'@

# ---------------------------------------------------------------------------
# Load window
# ---------------------------------------------------------------------------
$reader = New-Object System.Xml.XmlNodeReader $Xaml
$Win    = [Windows.Markup.XamlReader]::Load($reader)

function G($n) { $Win.FindName($n) }

# Pages
$Pages = @{
    Welcome = G "PageWelcome"
    Config  = G "PageConfig"
    Review  = G "PageReview"
    Deploy  = G "PageDeploy"
    Done    = G "PageDone"
}

# Prereq controls
$P = @{
    AzIcon  = G "PIconAz";   AzText  = G "PTextAz";   AzBadge  = G "PBadgeAz"
    GitIcon = G "PIconGit";  GitText = G "PTextGit";  GitBadge = G "PBadgeGit"
    GhIcon  = G "PIconGh";   GhText  = G "PTextGh";   GhBadge  = G "PBadgeGh"
    RepoIcon= G "PIconRepo"; RepoText= G "PTextRepo"; RepoBadge= G "PBadgeRepo"
}

# Mode / cred radios
$ModeStandard = G "ModeStandard";  $ModeMsp   = G "ModeMsp"
$CredSecret   = G "CredSecret";    $CredCert  = G "CredCert"

# Config inputs
$TxtPrefix       = G "TxtPrefix";  $CmbRegion = G "CmbRegion"
$CmbSubscription = G "CmbSubscription"; $TxtSubNote = G "TxtSubNote"
$TxtPwd1         = G "TxtPwd1";    $TxtPwd2   = G "TxtPwd2"
$TxtPwdErr       = G "TxtPwdErr"
$TxtClientUser   = G "TxtClientUser";  $TxtAzureUser = G "TxtAzureUser"
$TxtGitHubUser   = G "TxtGitHubUser";  $PanelClientUser = G "PanelClientUser"

# Subscription data store (populated when page 2 loads)
$script:Subscriptions = @()  # array of PSObjects with .id .name .isDefault

# Review labels
$RevMode = G "RevMode";  $RevPrefix = G "RevPrefix"
$RevRegion = G "RevRegion";  $RevCred = G "RevCred";  $RevRepo = G "RevRepo"
$RevSub    = G "RevSub"

# Deploy page
$PBar     = G "PBar"
$LogScroll= G "LogScroll"; $LogBox = G "LogBox"
$DI = 1..6 | ForEach-Object { G "DI$_" }
$DT = 1..6 | ForEach-Object { G "DT$_" }

# Done page
$PanelSuccess = G "PanelSuccess";  $PanelError = G "PanelError"
$TxtUrl       = G "TxtUrl";        $BtnCopyUrl = G "BtnCopyUrl"
$BtnOpenUrl   = G "BtnOpenUrl";    $BtnRetry   = G "BtnRetry"
$TxtConsentUrl= G "TxtConsentUrl"

# Sidebar
$SideDots = 1..5 | ForEach-Object { G "SideDot$_" }
$SideLabels = 1..5 | ForEach-Object { G "SideLbl$_" }

# Footer / nav
$BtnBack   = G "BtnBack";  $BtnNext = G "BtnNext";  $FooterMsg = G "FooterMsg"

# Version
$TxtVersion = G "TxtVersion"

# ---------------------------------------------------------------------------
# State
# ---------------------------------------------------------------------------
$script:Page       = 1
$script:RepoSlug   = ""
$script:ClientId   = ""
$script:DashUrl    = ""
$script:DeployJob  = $null

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function Set-Prereq($icon, $text, $badge, $state, $msg) {
    switch ($state) {
        "ok"   {
            $icon.Text = "●"; $icon.Foreground = "#3FB950"
            $text.Foreground = "#E6EDF3"; $text.Text = $msg
            $badge.Text = "OK"; $badge.Foreground = "#3FB950"
        }
        "warn" {
            $icon.Text = "●"; $icon.Foreground = "#D29922"
            $text.Foreground = "#8B949E"; $text.Text = $msg
            $badge.Text = "Optional"; $badge.Foreground = "#D29922"
        }
        "err"  {
            $icon.Text = "●"; $icon.Foreground = "#F85149"
            $text.Foreground = "#F85149"; $text.Text = $msg
            $badge.Text = "Missing"; $badge.Foreground = "#F85149"
        }
    }
}

function Update-Sidebar($page) {
    for ($i = 1; $i -le 5; $i++) {
        $dot = $SideDots[$i-1]; $lbl = $SideLabels[$i-1]
        if ($i -lt $page) {
            $dot.Background = "#238636"; $lbl.Foreground = "#3FB950"; $lbl.FontWeight = "Normal"
            $dot.Child.Text = "✓"; $dot.Child.Foreground = "White"
        } elseif ($i -eq $page) {
            $dot.Background = "#0078D4"; $lbl.Foreground = "White"; $lbl.FontWeight = "SemiBold"
            $dot.Child.Text = "$i"; $dot.Child.Foreground = "White"
        } else {
            $dot.Background = "#30363D"; $lbl.Foreground = "#8B949E"; $lbl.FontWeight = "Normal"
            $dot.Child.Text = "$i"; $dot.Child.Foreground = "#8B949E"
        }
    }
}

function Load-Subscriptions {
    $TxtSubNote.Text = "Loading subscriptions..."
    $CmbSubscription.Items.Clear()

    $ErrorActionPreference = "Continue"
    $rawSubs    = (cmd /c "az account list --query ""[?state=='Enabled']"" -o json 2>nul")
    $rawAccount = (cmd /c "az account show -o json 2>nul")
    $ErrorActionPreference = "Stop"

    # Extract currently logged-in user and pre-fill Azure account field if blank
    try {
        $accountJson = ($rawAccount | Where-Object { $_ -notmatch '^WARNING:' }) -join ""
        if ($accountJson -match '"user"') {
            $account = $accountJson | ConvertFrom-Json
            $loggedInUser = $account.user.name
            if ($loggedInUser -and [string]::IsNullOrWhiteSpace($TxtAzureUser.Text)) {
                $TxtAzureUser.Text = $loggedInUser
            }
        }
    } catch {}

    $script:Subscriptions = @()
    try {
        $subs = ($rawSubs -join "") | ConvertFrom-Json
        if ($subs -and $subs.Count -gt 0) {
            foreach ($sub in $subs) {
                $script:Subscriptions += $sub
                $label = if ($sub.isDefault) { "$($sub.name) (default)" } else { $sub.name }
                $item = New-Object System.Windows.Controls.ComboBoxItem
                $item.Content = $label
                $item.Tag     = $sub.id
                $item.Foreground = "#E6EDF3"
                $CmbSubscription.Items.Add($item) | Out-Null
                if ($sub.isDefault) { $CmbSubscription.SelectedItem = $item }
            }
            $userNote = if ($TxtAzureUser.Text) { " Logged in as: $($TxtAzureUser.Text)." } else { "" }
            if ($subs.Count -eq 1) {
                $TxtSubNote.Text = "One subscription found.$userNote"
            } else {
                $TxtSubNote.Text = "$($subs.Count) subscriptions found. Select the one to deploy into.$userNote"
            }
        } else {
            $TxtSubNote.Text = "No subscriptions found — run 'az login' first."
        }
    } catch {
        $TxtSubNote.Text = "Could not load subscriptions — run 'az login' first."
    }
}

function Show-Page($n) {
    foreach ($k in $Pages.Keys) { $Pages[$k].Visibility = "Collapsed" }
    switch ($n) {
        1 { $Pages.Welcome.Visibility = "Visible" }
        2 { $Pages.Config.Visibility  = "Visible"
            Load-Subscriptions
            # Show client tenant field only in MSP mode
            $PanelClientUser.Visibility = if ($ModeMsp.IsChecked) { "Visible" } else { "Collapsed" }
        }
        3 { $Pages.Review.Visibility  = "Visible" }
        4 { $Pages.Deploy.Visibility  = "Visible" }
        5 { $Pages.Done.Visibility    = "Visible" }
    }
    $script:Page = $n
    Update-Sidebar $n

    $BtnBack.IsEnabled = ($n -ge 2 -and $n -le 3)

    switch ($n) {
        1 { $BtnNext.Content = "Next →";      $BtnNext.IsEnabled = $true  }
        2 { $BtnNext.Content = "Next →";      $BtnNext.IsEnabled = $true  }
        3 { $BtnNext.Content = "🚀  Deploy";  $BtnNext.IsEnabled = $true  }
        4 { $BtnNext.Content = "Deploying…";  $BtnNext.IsEnabled = $false }
        5 { $BtnNext.Content = "Close";       $BtnNext.IsEnabled = $true  }
    }
    $FooterMsg.Text = ""
}

function Set-DeployStep($i, $state) {
    $icon = $DI[$i-1]; $txt = $DT[$i-1]
    switch ($state) {
        "pending" { $icon.Text = "○"; $icon.Foreground = "#484F58"; $txt.Foreground = "#8B949E" }
        "running" { $icon.Text = "◉"; $icon.Foreground = "#D29922"; $txt.Foreground = "White"   }
        "done"    { $icon.Text = "✓"; $icon.Foreground = "#3FB950"; $txt.Foreground = "#3FB950" }
        "error"   { $icon.Text = "✗"; $icon.Foreground = "#F85149"; $txt.Foreground = "#F85149" }
    }
}

function Add-Log($line) {
    if ([string]::IsNullOrWhiteSpace($line)) { return }
    $LogBox.Text = $LogBox.Text + [string]$line + "`n"
    $LogScroll.ScrollToEnd()
}

function Check-Prereqs {
    # Azure CLI
    $az = Get-Command az -ErrorAction SilentlyContinue
    if ($az) {
        try {
            $azVer = (& az --version 2>$null) | Select-Object -First 1
            $ver = if ($azVer -match 'azure-cli\s+([\d\.]+)') { $Matches[1] } else { $azVer.ToString().Trim() }
        } catch { $ver = "found" }
        Set-Prereq $P.AzIcon $P.AzText $P.AzBadge "ok" "Azure CLI $ver"
    } else {
        Set-Prereq $P.AzIcon $P.AzText $P.AzBadge "err" "Not found — install from aka.ms/installazurecliwindows"
        $BtnNext.IsEnabled = $false
        $FooterMsg.Text = "WARNING: Azure CLI is required. Install it and restart the wizard."
    }

    # Git
    $git = Get-Command git -ErrorAction SilentlyContinue
    if ($git) {
        try { $ver = (& git --version 2>$null) -replace "git version ","" } catch { $ver = "" }
        $ver = if ($ver) { $ver.ToString().Trim() } else { "(version unknown)" }
        Set-Prereq $P.GitIcon $P.GitText $P.GitBadge "ok" "Git $ver found"
    } else {
        Set-Prereq $P.GitIcon $P.GitText $P.GitBadge "err" "Not found — install from git-scm.com"
        $BtnNext.IsEnabled = $false
        $FooterMsg.Text = "WARNING: Git is required. Install it and restart the wizard."
    }

    # GitHub CLI
    $gh = Get-Command gh -ErrorAction SilentlyContinue
    if ($gh) {
        try { $ver = (& gh --version 2>$null | Select-Object -First 1) -replace "gh version ","" } catch { $ver = "" }
        $ver = if ($ver) { $ver.ToString().Trim() } else { "(version unknown)" }
        Set-Prereq $P.GhIcon $P.GhText $P.GhBadge "ok" "GitHub CLI $ver found"
    } else {
        Set-Prereq $P.GhIcon $P.GhText $P.GhBadge "warn" "Not found — will be installed automatically during deployment"
    }

    # GitHub repo slug
    try {
        $root   = Split-Path $PSScriptRoot -Parent
        $remote = (git -C $root remote get-url origin 2>$null).Trim()
        if ($remote -match "github\.com[:/](.+?)(\.git)?$") {
            $script:RepoSlug = $Matches[1].Trim()
            Set-Prereq $P.RepoIcon $P.RepoText $P.RepoBadge "ok" "github.com/$($script:RepoSlug)"
        } else {
            Set-Prereq $P.RepoIcon $P.RepoText $P.RepoBadge "warn" "No GitHub remote — CI/CD secrets will be printed manually"
        }
    } catch {
        Set-Prereq $P.RepoIcon $P.RepoText $P.RepoBadge "warn" "Could not detect git remote"
    }

    # Version
    try {
        $vFile = Join-Path (Split-Path $PSScriptRoot -Parent) "src\M365Dashboard.Api\Properties\AssemblyInfo.cs"
        if (Test-Path $vFile) {
            $match = (Get-Content $vFile | Select-String 'AssemblyVersion\("(.+?)"').Matches[0]
            if ($match) { $TxtVersion.Text = "v$($match.Groups[1].Value)" }
        } else {
            $TxtVersion.Text = "v1.x"
        }
    } catch { $TxtVersion.Text = "v1.x" }
}

function Validate-Page2 {
    $prefix = $TxtPrefix.Text.Trim()
    if ($prefix -notmatch "^[a-zA-Z][a-zA-Z0-9]{2,11}$") {
        $FooterMsg.Text = "Prefix must be 3-12 chars, start with a letter, letters/numbers only."
        return $false
    }
    $p1 = $TxtPwd1.Password; $p2 = $TxtPwd2.Password
    if ($p1.Length -lt 12) {
        $TxtPwdErr.Text = "Password must be at least 12 characters."; $TxtPwdErr.Visibility = "Visible"; return $false
    }
    if ($p1 -ne $p2) {
        $TxtPwdErr.Text = "Passwords do not match."; $TxtPwdErr.Visibility = "Visible"; return $false
    }
    if ($p1 -notmatch "[A-Z]" -or $p1 -notmatch "[a-z]" -or $p1 -notmatch "[0-9]" -or $p1 -notmatch "[^a-zA-Z0-9]") {
        $TxtPwdErr.Text = "Password must include uppercase, lowercase, number and special character."
        $TxtPwdErr.Visibility = "Visible"; return $false
    }
    $TxtPwdErr.Visibility = "Collapsed"
    $FooterMsg.Text = ""
    return $true
}

function Populate-Review {
    $RevMode.Text   = if ($ModeStandard.IsChecked) { "Standard" } else { "MSP / Multi-tenant" }
    $RevPrefix.Text = $TxtPrefix.Text.Trim().ToLower()
    $selItem = $CmbRegion.SelectedItem
    $RevRegion.Text = if ($selItem) { $selItem.Content } else { "UK South" }
    $RevCred.Text   = if ($CredSecret.IsChecked) { "Client Secret" } else { "Certificate" }
    $RevRepo.Text   = if ($script:RepoSlug) { "github.com/$($script:RepoSlug)" } else { "(not detected)" }
    $selSub = $CmbSubscription.SelectedItem
    if ($RevSub) { $RevSub.Text = if ($selSub) { $selSub.Content } else { "(default)" } }
}

function Show-MspLoginDialog($clientUser, $azureUser, $gitHubUser) {
    # Build account hint lines — only show if values were entered
    $hint1 = if ($clientUser) { "Account: $clientUser" } else { "" }
    $hint2 = if ($azureUser)  { "Account: $azureUser"  } else { "" }
    $hint3 = if ($gitHubUser) { "Username: $gitHubUser" } else { "" }
    # Show a clear modal dialog explaining the two MSP login steps before any popups appear
    $dlgXaml = [xml]@'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="MSP Deployment — Login Guide"
    Width="540" Height="600"
    WindowStartupLocation="CenterOwner"
    ResizeMode="NoResize"
    Background="#0D1117" Foreground="#E6EDF3"
    FontFamily="Segoe UI" FontSize="13">
  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <ScrollViewer Grid.Row="0" VerticalScrollBarVisibility="Auto" Padding="32,28,32,8">
      <StackPanel>
        <TextBlock Text="Logins required" FontSize="20" FontWeight="Bold" Foreground="White" Margin="0,0,0,6"/>
        <TextBlock TextWrapping="Wrap" Foreground="#8B949E" Margin="0,0,0,20" FontSize="12">
          This deployment requires three sign-ins. Browser windows will open automatically at each step.
        </TextBlock>

        <!-- Login 1 -->
        <Border Background="#0D2137" BorderBrush="#0078D4" BorderThickness="1" CornerRadius="7" Padding="16,14" Margin="0,0,0,10">
          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="36"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Border Width="28" Height="28" CornerRadius="14" Background="#0078D4" VerticalAlignment="Top" Margin="0,2,0,0">
              <TextBlock Text="1" Foreground="White" FontWeight="Bold" FontSize="13" HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <StackPanel Grid.Column="1">
              <TextBlock Text="CLIENT Tenant — Azure login" FontWeight="Bold" Foreground="#58A6FF" Margin="0,0,0,4"/>
              <TextBlock TextWrapping="Wrap" Foreground="#E6EDF3" FontSize="12">
                Sign in as a Global Admin of the CLIENT&apos;s Microsoft 365 tenant.
              </TextBlock>
              <TextBlock x:Name="Hint1" TextWrapping="Wrap" FontSize="12" FontWeight="SemiBold"
                         Foreground="#58A6FF" Margin="0,8,0,0"/>
              <TextBlock TextWrapping="Wrap" Foreground="#8B949E" FontSize="11" Margin="0,4,0,0">
                Used to create the app registration in their tenant. The client tenant may have no Azure subscription — that is normal.
              </TextBlock>
            </StackPanel>
          </Grid>
        </Border>

        <!-- Login 2 -->
        <Border Background="#0D1F0D" BorderBrush="#238636" BorderThickness="1" CornerRadius="7" Padding="16,14" Margin="0,0,0,10">
          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="36"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Border Width="28" Height="28" CornerRadius="14" Background="#238636" VerticalAlignment="Top" Margin="0,2,0,0">
              <TextBlock Text="2" Foreground="White" FontWeight="Bold" FontSize="13" HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <StackPanel Grid.Column="1">
              <TextBlock Text="YOUR Azure Subscription — Azure login" FontWeight="Bold" Foreground="#3FB950" Margin="0,0,0,4"/>
              <TextBlock TextWrapping="Wrap" Foreground="#E6EDF3" FontSize="12">
                Sign in to your own Azure subscription.
              </TextBlock>
              <TextBlock x:Name="Hint2" TextWrapping="Wrap" FontSize="12" FontWeight="SemiBold"
                         Foreground="#3FB950" Margin="0,8,0,0"/>
              <TextBlock TextWrapping="Wrap" Foreground="#8B949E" FontSize="11" Margin="0,4,0,0">
                This is where the Container App, SQL, Key Vault and ACR are created.
              </TextBlock>
            </StackPanel>
          </Grid>
        </Border>

        <!-- Login 3 -->
        <Border Background="#1A1207" BorderBrush="#7D4F12" BorderThickness="1" CornerRadius="7" Padding="16,14" Margin="0,0,0,20">
          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="36"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Border Width="28" Height="28" CornerRadius="14" Background="#7D4F12" VerticalAlignment="Top" Margin="0,2,0,0">
              <TextBlock Text="3" Foreground="White" FontWeight="Bold" FontSize="13" HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <StackPanel Grid.Column="1">
              <TextBlock Text="GitHub — gh auth login" FontWeight="Bold" Foreground="#D29922" Margin="0,0,0,4"/>
              <TextBlock TextWrapping="Wrap" Foreground="#E6EDF3" FontSize="12">
                Sign in to the GitHub account that owns the repository.
              </TextBlock>
              <TextBlock x:Name="Hint3" TextWrapping="Wrap" FontSize="12" FontWeight="SemiBold"
                         Foreground="#D29922" Margin="0,8,0,0"/>
              <TextBlock TextWrapping="Wrap" Foreground="#8B949E" FontSize="11" Margin="0,4,0,0">
                Used to automatically configure GitHub Actions CI/CD secrets so every push to main auto-deploys.
              </TextBlock>
            </StackPanel>
          </Grid>
        </Border>

        <Border Background="#161B22" BorderBrush="#30363D" BorderThickness="1" CornerRadius="6" Padding="12,10">
          <TextBlock TextWrapping="Wrap" Foreground="#8B949E" FontSize="11">
            Tip: have all three accounts ready before clicking Start. If you miss a login the wizard will still complete — GitHub secrets can be set manually afterwards.
          </TextBlock>
        </Border>
      </StackPanel>
    </ScrollViewer>

    <Border Grid.Row="1" Background="#161B22" BorderBrush="#30363D" BorderThickness="0,1,0,0" Padding="32,16">
      <Button x:Name="BtnOk" Content="Got it — start deployment"
              Height="40" FontSize="13" FontWeight="SemiBold"
              Foreground="White" Background="#0078D4" BorderThickness="0" Cursor="Hand"/>
    </Border>
  </Grid>
</Window>
'@
    $dlgReader = New-Object System.Xml.XmlNodeReader $dlgXaml
    $dlg = [Windows.Markup.XamlReader]::Load($dlgReader)
    $dlg.Owner = $Win
    # Populate hint labels with entered usernames
    if ($hint1) { $dlg.FindName("Hint1").Text = $hint1 }
    if ($hint2) { $dlg.FindName("Hint2").Text = $hint2 }
    if ($hint3) { $dlg.FindName("Hint3").Text = $hint3 }
    $dlg.FindName("BtnOk").Add_Click({ $dlg.Close() })
    [void]$dlg.ShowDialog()
}

function Start-Deploy {
    $prefix      = $TxtPrefix.Text.Trim().ToLower()
    $region      = ($CmbRegion.SelectedItem).Tag
    $useCert     = [bool]$CredCert.IsChecked
    $isMsp       = [bool]$ModeMsp.IsChecked
    $sqlPwd      = $TxtPwd1.Password
    $clientUser  = $TxtClientUser.Text.Trim()
    $azureUser   = $TxtAzureUser.Text.Trim()
    $gitHubUser  = $TxtGitHubUser.Text.Trim()
    $deployPs    = Join-Path $PSScriptRoot "Deploy-M365Dashboard.ps1"

    # Show MSP login briefing with entered account hints before starting
    if ($isMsp) { Show-MspLoginDialog $clientUser $azureUser $gitHubUser }

    Show-Page 4
    1..6 | ForEach-Object { Set-DeployStep $_ "pending" }
    Set-DeployStep 1 "running"
    $PBar.Value = 5

    # Pass SQL password via environment variable to avoid shell quoting issues
    $env:WIZARD_SQL_PASSWORD = $sqlPwd

    $selSub = $CmbSubscription.SelectedItem
    $subId  = if ($selSub -and $selSub.Tag) { $selSub.Tag } else { "" }

    $argList = @(
        "-NamePrefix",      $prefix,
        "-Location",        $region,
        "-DeployMode",      $(if ($isMsp) { "MSP" } else { "Standard" }),
        "-CredentialType",  $(if ($useCert) { "Certificate" } else { "Secret" }),
        "-SqlPassword",     $sqlPwd,
        "-NonInteractive"
    )
    if ($subId) { $argList += @("-SubscriptionId", $subId) }

    # Step patterns — matched against each output line
    $script:StepMap = @(
        @{ Re = "Checking Azure CLI|az login|Logged in as";                                Step=1; Pct=12 }
        @{ Re = "Creating app registration|App registration created|App created";          Step=2; Pct=26 }
        @{ Re = "Creating Azure resources|Deploying infrastructure|Container App created"; Step=3; Pct=55 }
        @{ Re = "Building Docker|az acr build|Docker image built|Successfully built";      Step=4; Pct=78 }
        @{ Re = "Configuring App Registration|Redirect URI|admin consent|logo upload";     Step=5; Pct=90 }
        @{ Re = "GitHub Actions|gh secret|All 8 GitHub|CI/CD";                            Step=6; Pct=97 }
        @{ Re = "Deployment Complete";                                                     Step=0; Pct=100 }
    )
    $script:CompletedSteps = @{}
    $script:RunningStep    = 1

    # Launch deploy script as background job
    # Note: avoid $args as param name — it's a reserved variable in PowerShell
    $script:DeployJob = Start-Job -ScriptBlock {
        param($scriptPath, $scriptArgs, $sqlPassword)
        $env:WIZARD_SQL_PASSWORD = $sqlPassword
        $ErrorActionPreference = 'Continue'
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $scriptPath @scriptArgs 2>&1
    } -ArgumentList $deployPs, $argList, $sqlPwd

    # Poll timer (every 600ms)
    $script:PollTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:PollTimer.Interval = [TimeSpan]::FromMilliseconds(600)
    $script:PollTimer.Add_Tick({
        if (-not $script:DeployJob) { $script:PollTimer.Stop(); return }

        $lines = Receive-Job $script:DeployJob -ErrorAction SilentlyContinue
        foreach ($line in $lines) {
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            Add-Log $line

            foreach ($sm in $script:StepMap) {
                if ($line -match $sm.Re) {
                    if ($sm.Step -eq 0) {
                        $PBar.Value = 100
                        if ($line -match "https://\S+azurecontainerapps") {
                            $script:DashUrl = ($line | Select-String "https://\S+").Matches[0].Value.TrimEnd('.')
                        }
                    } else {
                        if ($sm.Step -ne $script:RunningStep -and -not $script:CompletedSteps[$script:RunningStep]) {
                            Set-DeployStep $script:RunningStep "done"
                            $script:CompletedSteps[$script:RunningStep] = $true
                        }
                        $script:RunningStep = $sm.Step
                        Set-DeployStep $sm.Step "running"
                        if ($sm.Pct -gt $PBar.Value) { $PBar.Value = $sm.Pct }
                    }
                }
            }

            # Capture dashboard URL from "available at" line
            if ($line -match "available at:?\s+(https://\S+)") {
                $script:DashUrl = $Matches[1].Trim('.').Trim()
            }
            # Capture client ID for consent URL
            if ($line -match "Client ID[:\s]+([0-9a-f\-]{36})" -and -not $script:ClientId) {
                $script:ClientId = $Matches[1].Trim()
            }
        }

        if ($script:DeployJob.State -in @("Completed","Failed","Stopped")) {
            $script:PollTimer.Stop()
            $ok = ($script:DeployJob.State -eq "Completed")

            1..6 | ForEach-Object {
                if (-not $script:CompletedSteps[$_]) {
                    $stepState = if ($ok) { "done" } else { "error" }
                    Set-DeployStep $_ $stepState
                }
            }
            if ($ok) { $PBar.Value = 100 }

            Remove-Job $script:DeployJob -Force -ErrorAction SilentlyContinue
            $script:DeployJob = $null
            Finish-Deploy $ok
        }
    })
    $script:PollTimer.Start()
}

function Finish-Deploy($ok) {
    Show-Page 5
    if ($ok) {
        $PanelSuccess.Visibility = "Visible"
        $PanelError.Visibility   = "Collapsed"
        $TxtUrl.Text = if ($script:DashUrl) { $script:DashUrl } else { "(check log for URL)" }
        $consentUrl = if ($script:ClientId) {
            "entra.microsoft.com → App registrations → API permissions → Grant admin consent"
        } else { "Entra ID → App registrations → your app → API permissions → Grant admin consent" }
        $TxtConsentUrl.Text = $consentUrl
    } else {
        $PanelSuccess.Visibility = "Collapsed"
        $PanelError.Visibility   = "Visible"
    }
}

# ---------------------------------------------------------------------------
# Button wiring
# ---------------------------------------------------------------------------
$BtnNext.Add_Click({
    switch ($script:Page) {
        1 { Show-Page 2 }
        2 { if (Validate-Page2) { Populate-Review; Show-Page 3 } }
        3 { Start-Deploy }
        5 { $Win.Close() }
    }
})

$BtnBack.Add_Click({
    if ($script:Page -eq 2) { Show-Page 1 }
    if ($script:Page -eq 3) { Show-Page 2 }
})

$BtnOpenUrl.Add_Click({
    if ($script:DashUrl) { Start-Process $script:DashUrl }
})

$BtnCopyUrl.Add_Click({
    if ($script:DashUrl) {
        [System.Windows.Clipboard]::SetText($script:DashUrl)
        $BtnCopyUrl.Content = "Copied!"
        $timer2 = New-Object System.Windows.Threading.DispatcherTimer
        $timer2.Interval = [TimeSpan]::FromSeconds(2)
        $timer2.Add_Tick({ $BtnCopyUrl.Content = "Copy"; $timer2.Stop() })
        $timer2.Start()
    }
})

$BtnRetry.Add_Click({ Show-Page 3 })

# ---------------------------------------------------------------------------
# Window events
# ---------------------------------------------------------------------------
$Win.Add_Loaded({
    Show-Page 1
    Check-Prereqs
})

$Win.Add_Closing({
    if ($script:DeployJob) {
        $r = [System.Windows.MessageBox]::Show(
            "Deployment is still running. Are you sure you want to close?",
            "Deployment in Progress",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning)
        if ($r -eq [System.Windows.MessageBoxResult]::No) {
            $_.Cancel = $true; return
        }
        Stop-Job $script:DeployJob -ErrorAction SilentlyContinue
        Remove-Job $script:DeployJob -Force -ErrorAction SilentlyContinue
    }
    $env:WIZARD_SQL_PASSWORD = $null
})

# ---------------------------------------------------------------------------
# Launch
# ---------------------------------------------------------------------------
[void]$Win.ShowDialog()
