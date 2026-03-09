<#
    DISCLAIMER: This script is provided "AS IS" without warranty of any kind.
    Use it at your own risk. The author is not responsible for any damage or
    data loss caused by using this script. Always test in a non-production
    environment before deploying to production.
#>
#Requires -Modules PnP.PowerShell
<#
.SYNOPSIS
    GUI tool to register an Entra ID app with Sites.Selected permissions.

.DESCRIPTION
    Provides a WPF graphical interface to:
    - Register an Entra ID app with Sites.Selected permissions
    - Generate a self-signed certificate for certificate-based auth
    - Save registration details to a summary file

    After registration, a Global/Application Administrator must grant
    admin consent for the API permissions in the Azure portal.

.EXAMPLE
    .\Register-SitesSelectedAppGUI.ps1
#>

$ErrorActionPreference = 'Stop'

# ── Check / import PnP.PowerShell ───────────────────────────────────────────
if (-not (Get-Module -Name PnP.PowerShell)) {
    if (Get-Module -ListAvailable -Name PnP.PowerShell) {
        Import-Module PnP.PowerShell -ErrorAction Stop
    } else {
        throw "PnP.PowerShell module is not installed or loaded. Please install it (Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force) or load it manually before running this script."
    }
}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# ═════════════════════════════════════════════════════════════════════════════
#  XAML – Window Definition
# ═════════════════════════════════════════════════════════════════════════════
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Register Sites.Selected App"
        Width="700" Height="1020"
        WindowStartupLocation="CenterScreen"
        Background="#F5F5F5">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Padding" Value="12,6"/>
            <Setter Property="Margin" Value="4"/>
            <Setter Property="FontSize" Value="12"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Margin" Value="4"/>
            <Setter Property="Padding" Value="4,3"/>
            <Setter Property="FontSize" Value="12"/>
        </Style>
        <Style TargetType="PasswordBox">
            <Setter Property="Margin" Value="4"/>
            <Setter Property="Padding" Value="4,3"/>
            <Setter Property="FontSize" Value="12"/>
        </Style>
        <Style TargetType="TextBlock">
            <Setter Property="Margin" Value="4"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
            <Setter Property="FontSize" Value="12"/>
        </Style>
        <Style TargetType="ComboBox">
            <Setter Property="Margin" Value="4"/>
            <Setter Property="Padding" Value="4,3"/>
            <Setter Property="FontSize" Value="12"/>
        </Style>
    </Window.Resources>
    <DockPanel>
        <!-- ── Log Panel ────────────────────────────────────────────── -->
        <Border DockPanel.Dock="Bottom" Margin="10,4,10,10" Padding="4"
                BorderBrush="#CCCCCC" BorderThickness="1" CornerRadius="4"
                Background="White" Height="160">
            <DockPanel>
                <StackPanel DockPanel.Dock="Top" Orientation="Horizontal">
                    <TextBlock Text="Log" FontWeight="SemiBold" Margin="4,2"/>
                    <Button Name="btnClearLog" Content="Clear" FontSize="10"
                            Padding="8,2" Margin="8,0,0,0"
                            VerticalAlignment="Center"/>
                </StackPanel>
                <TextBox Name="txtLog" IsReadOnly="True"
                         VerticalScrollBarVisibility="Auto"
                         HorizontalScrollBarVisibility="Auto"
                         TextWrapping="Wrap" FontFamily="Consolas" FontSize="11"
                         Background="#FAFAFA" BorderThickness="0"/>
            </DockPanel>
        </Border>

        <!-- ── Registration Form ────────────────────────────────────── -->
        <Border Margin="10,10,10,4" Padding="16"
                BorderBrush="#CCCCCC" BorderThickness="1" CornerRadius="4"
                Background="White">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="16"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="160"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <!-- Title -->
                <TextBlock Grid.ColumnSpan="3" Text="Register Entra ID App with Sites.Selected"
                           FontSize="16" FontWeight="Bold" Margin="4,0,4,12"/>

                <!-- App Name -->
                <TextBlock Grid.Row="1" Text="Application Name:" FontWeight="SemiBold"/>
                <TextBox Name="txtAppName" Grid.Row="1" Grid.Column="1" Grid.ColumnSpan="2"
                         Text="SP-SitesSelected-App"/>

                <!-- Tenant -->
                <TextBlock Grid.Row="2" Text="Tenant:" FontWeight="SemiBold"/>
                <TextBox Name="txtTenant" Grid.Row="2" Grid.Column="1" Grid.ColumnSpan="2"
                         Text="contoso.onmicrosoft.com"/>

                <!-- Certificate Output Path -->
                <TextBlock Grid.Row="3" Text="Certificate Path:" FontWeight="SemiBold"/>
                <TextBox Name="txtOutPath" Grid.Row="3" Grid.Column="1" IsReadOnly="True"/>
                <Button Name="btnBrowse" Grid.Row="3" Grid.Column="2" Content="Browse..."
                        Padding="8,4"/>

                <!-- Certificate Password -->
                <TextBlock Grid.Row="4" Text="Certificate Password:" FontWeight="SemiBold"/>
                <PasswordBox Name="txtCertPassword" Grid.Row="4" Grid.Column="1"
                             Grid.ColumnSpan="2"/>

                <!-- Valid Years -->
                <TextBlock Grid.Row="5" Text="Certificate Valid (years):" FontWeight="SemiBold"/>
                <ComboBox Name="cmbValidYears" Grid.Row="5" Grid.Column="1"
                          Width="80" HorizontalAlignment="Left">
                    <ComboBoxItem Content="1"/>
                    <ComboBoxItem Content="2" IsSelected="True"/>
                    <ComboBoxItem Content="3"/>
                    <ComboBoxItem Content="5"/>
                    <ComboBoxItem Content="10"/>
                </ComboBox>

                <!-- Permissions -->
                <TextBlock Grid.Row="6" Text="Permissions:" FontWeight="SemiBold"/>
                <StackPanel Grid.Row="6" Grid.Column="1" Grid.ColumnSpan="2"
                            Orientation="Horizontal">
                    <CheckBox Name="chkGraph" Content="Microsoft Graph Sites.Selected"
                              IsChecked="True" IsEnabled="False" Margin="4"
                              VerticalAlignment="Center"/>
                    <CheckBox Name="chkSharePoint" Content="SharePoint Sites.Selected"
                              IsChecked="True" Margin="12,4"
                              VerticalAlignment="Center"/>
                </StackPanel>

                <!-- Separator row 7 is just spacing -->

                <!-- Register Button -->
                <Button Name="btnRegister" Grid.Row="8" Grid.ColumnSpan="3"
                        Content="Register App"
                        HorizontalAlignment="Right" Margin="4,4"
                        Background="#107C10" Foreground="White" FontWeight="Bold"
                        FontSize="14" Padding="20,8"/>

                <!-- Results -->
                <Border Name="pnlResults" Grid.Row="9" Grid.ColumnSpan="3"
                        Margin="0,8,0,0" Padding="10"
                        Background="#E6F4EA" CornerRadius="4"
                        BorderBrush="#A8D5A2" BorderThickness="1"
                        Visibility="Collapsed">
                    <StackPanel>
                        <TextBlock Text="Registration Successful" FontWeight="Bold"
                                   FontSize="13" Foreground="#2E7D32" Margin="0,0,0,6"/>
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="130"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            <TextBlock Text="Client ID:" FontWeight="SemiBold"/>
                            <TextBox Name="txtResultClientId" Grid.Column="1"
                                     IsReadOnly="True" BorderThickness="0"
                                     Background="Transparent" FontFamily="Consolas"/>
                            <TextBlock Grid.Row="1" Text="Thumbprint:" FontWeight="SemiBold"/>
                            <TextBox Name="txtResultThumbprint" Grid.Row="1" Grid.Column="1"
                                     IsReadOnly="True" BorderThickness="0"
                                     Background="Transparent" FontFamily="Consolas"/>
                            <TextBlock Grid.Row="2" Text="PFX File:" FontWeight="SemiBold"/>
                            <TextBox Name="txtResultPfx" Grid.Row="2" Grid.Column="1"
                                     IsReadOnly="True" BorderThickness="0"
                                     Background="Transparent" FontFamily="Consolas"/>
                        </Grid>
                    </StackPanel>
                </Border>
            </Grid>
        </Border>
    </DockPanel>
</Window>
"@

# ═════════════════════════════════════════════════════════════════════════════
#  Load XAML and get named controls
# ═════════════════════════════════════════════════════════════════════════════
$reader = [System.Xml.XmlNodeReader]::new($xaml)
$window = [System.Windows.Markup.XamlReader]::Load($reader)

$controls = @{}
$xaml.SelectNodes('//*[@Name]') | ForEach-Object {
    $controls[$_.Name] = $window.FindName($_.Name)
}

$txtLog              = $controls['txtLog']
$btnClearLog         = $controls['btnClearLog']
$txtAppName          = $controls['txtAppName']
$txtTenant           = $controls['txtTenant']
$txtOutPath          = $controls['txtOutPath']
$btnBrowse           = $controls['btnBrowse']
$txtCertPassword     = $controls['txtCertPassword']
$cmbValidYears       = $controls['cmbValidYears']
$chkSharePoint       = $controls['chkSharePoint']
$btnRegister         = $controls['btnRegister']
$pnlResults          = $controls['pnlResults']
$txtResultClientId   = $controls['txtResultClientId']
$txtResultThumbprint = $controls['txtResultThumbprint']
$txtResultPfx        = $controls['txtResultPfx']

# Default certificate path
$txtOutPath.Text = Join-Path $PSScriptRoot "Certificates"

# ═════════════════════════════════════════════════════════════════════════════
#  Helper Functions
# ═════════════════════════════════════════════════════════════════════════════

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "HH:mm:ss"
    $entry = "[$timestamp] [$Level] $Message`r`n"
    $txtLog.AppendText($entry)
    $txtLog.ScrollToEnd()
    $window.Dispatcher.Invoke(
        [Action]{},
        [System.Windows.Threading.DispatcherPriority]::Background
    )
}

# Track background process state
$script:bgProcess  = $null
$script:pollTimer  = $null
$script:regContext = $null   # stores UI values captured before async call
$script:resultFile = $null

# ═════════════════════════════════════════════════════════════════════════════
#  Event Handlers
# ═════════════════════════════════════════════════════════════════════════════

$btnClearLog.Add_Click({ $txtLog.Clear() })

# ── Browse for output folder ────────────────────────────────────────────────
$btnBrowse.Add_Click({
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = [System.Windows.Forms.FolderBrowserDialog]::new()
    $dialog.Description = "Select certificate output folder"
    $dialog.SelectedPath = $txtOutPath.Text
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtOutPath.Text = $dialog.SelectedPath
    }
})

# ── Register ────────────────────────────────────────────────────────────────
$btnRegister.Add_Click({
    $appName    = $txtAppName.Text.Trim()
    $tenant     = $txtTenant.Text.Trim()
    $outPath    = $txtOutPath.Text.Trim()
    $certPwd    = $txtCertPassword.Password
    $validYears = [int]$cmbValidYears.SelectedItem.Content
    $addSP      = $chkSharePoint.IsChecked

    # Validation
    if ([string]::IsNullOrWhiteSpace($appName)) {
        [System.Windows.MessageBox]::Show("Application Name is required.", "Validation")
        return
    }
    if ([string]::IsNullOrWhiteSpace($tenant)) {
        [System.Windows.MessageBox]::Show("Tenant is required.", "Validation")
        return
    }
    if ([string]::IsNullOrWhiteSpace($outPath)) {
        [System.Windows.MessageBox]::Show("Certificate output path is required.", "Validation")
        return
    }
    if ([string]::IsNullOrWhiteSpace($certPwd)) {
        [System.Windows.MessageBox]::Show("Certificate password is required.", "Validation")
        return
    }

    # Create output directory if needed
    if (-not (Test-Path $outPath)) {
        New-Item -ItemType Directory -Path $outPath -Force | Out-Null
        Write-Log "Created output directory: $outPath"
    }

    $permText = "Microsoft Graph Sites.Selected"
    if ($addSP) { $permText += " + SharePoint Sites.Selected" }

    Write-Log "Registering app '$appName' on tenant '$tenant'..."
    Write-Log "Permissions: $permText"
    Write-Log "Certificate valid for $validYears year(s), output: $outPath"
    Write-Log "A browser window will open for authentication..."

    $btnRegister.IsEnabled = $false
    $pnlResults.Visibility = [System.Windows.Visibility]::Collapsed

    # Save context for use by the poll-timer callback
    $script:regContext = @{
        AppName    = $appName
        Tenant     = $tenant
        OutPath    = $outPath
        ValidYears = $validYears
    }

    # Resolve the PnP module path from the already-loaded module
    $pnpModulePath = (Get-Module PnP.PowerShell).Path

    # ── Build a temp script and run in a separate pwsh process ───────
    # A separate process has a full interactive host so PnP's browser
    # auth and localhost redirect work correctly.
    $script:resultFile = Join-Path $env:TEMP "pnp-register-result-$([guid]::NewGuid().ToString('N')).json"

    $tempScript = Join-Path $env:TEMP "pnp-register-$([guid]::NewGuid().ToString('N')).ps1"
    $scriptBody = @"
`$ErrorActionPreference = 'Stop'
try {
    Import-Module '$($pnpModulePath -replace "'","''")' -ErrorAction Stop
    `$secPwd = ConvertTo-SecureString '$($certPwd -replace "'","''")' -AsPlainText -Force
    `$params = @{
        ApplicationName             = '$($appName -replace "'","''")'
        Tenant                      = '$($tenant -replace "'","''")'
        OutPath                     = '$($outPath -replace "'","''")'
        CertificatePassword         = `$secPwd
        ValidYears                  = $validYears
        GraphApplicationPermissions = 'Sites.Selected'
    }
    $(if ($addSP) { "`$params['SharePointApplicationPermissions'] = 'Sites.Selected'" })
    `$result = Register-PnPEntraIDApp @params
    @{
        Success    = `$true
        ClientId   = `$result.'AzureAppId/ClientId'
        Thumbprint = `$result.'Certificate Thumbprint'
        PfxFile    = `$result.'Pfx file'
    } | ConvertTo-Json | Set-Content -Path '$($script:resultFile -replace "'","''")' -Encoding UTF8
} catch {
    @{
        Success = `$false
        Error   = `$_.Exception.Message
    } | ConvertTo-Json | Set-Content -Path '$($script:resultFile -replace "'","''")' -Encoding UTF8
}
"@
    $scriptBody | Set-Content -Path $tempScript -Encoding UTF8

    $script:bgProcess = Start-Process pwsh -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $tempScript `
        -PassThru -WindowStyle Normal

    # ── Poll timer – checks every 500 ms if the process finished ────
    $script:pollTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $script:pollTimer.Interval = [TimeSpan]::FromMilliseconds(500)
    $script:pollTimer.Add_Tick({
        if (-not $script:bgProcess.HasExited) { return }

        $script:pollTimer.Stop()

        $ctx        = $script:regContext
        $appName    = $ctx.AppName
        $outPath    = $ctx.OutPath
        $tenant     = $ctx.Tenant
        $validYears = $ctx.ValidYears

        try {
            if (-not (Test-Path $script:resultFile)) {
                Write-Log "Registration failed: no result file produced. The process may have crashed." "ERROR"
                return
            }

            $result = Get-Content $script:resultFile -Raw | ConvertFrom-Json

            if (-not $result.Success) {
                Write-Log "Registration failed: $($result.Error)" "ERROR"
                [System.Windows.MessageBox]::Show(
                    "Failed to register the application.`n`n$($result.Error)`n`nEnsure you sign in with Global Admin or Application Admin credentials.",
                    "Registration Error",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Error
                )
                return
            }

            $clientId   = $result.ClientId
            $thumbprint = $result.Thumbprint
            $pfxFile    = $result.PfxFile

            Write-Log "Registration successful!"
            if ($clientId)   { Write-Log "  Client ID:   $clientId" }
            if ($thumbprint) { Write-Log "  Thumbprint:  $thumbprint" }
            if ($pfxFile)    { Write-Log "  PFX File:    $pfxFile" }

            # Show results panel
            $txtResultClientId.Text   = if ($clientId)   { $clientId }   else { "N/A" }
            $txtResultThumbprint.Text = if ($thumbprint) { $thumbprint } else { "N/A" }
            $txtResultPfx.Text        = if ($pfxFile)    { $pfxFile }    else { "N/A" }
            $pnlResults.Visibility = [System.Windows.Visibility]::Visible

            # Save summary file
            $timestamp   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $safeName    = $appName -replace '[\\/:*?"<>|]', '_'
            $summaryPath = Join-Path $outPath "$safeName-registration-details.txt"
            $summaryLines = @(
                "App Registration Details"
                "========================"
                "Application Name:       $appName"
                "Tenant:                 $tenant"
                "Date Created:           $timestamp"
                "Certificate Valid:      $validYears year(s)"
                "Certificate Path:       $outPath"
                ""
            )
            if ($clientId)   { $summaryLines += "Client ID (App ID):     $clientId" }
            if ($thumbprint) { $summaryLines += "Certificate Thumbprint: $thumbprint" }
            if ($pfxFile)    { $summaryLines += "PFX File:               $pfxFile" }
            $summaryLines += @(
                ""
                "NEXT STEPS:"
                "1. Grant admin consent in Azure Portal"
                "2. Use Manage-SitesSelectedGUI.ps1 to assign Sites.Selected permissions"
                "3. Store PFX file and password securely"
            )
            $summaryLines | Out-File -FilePath $summaryPath -Encoding UTF8
            Write-Log "Summary saved to: $summaryPath"

            Write-Log ""
            Write-Log "NEXT STEPS:"
            Write-Log "  1. Grant admin consent in Azure Portal > Entra ID > App registrations > '$appName' > API Permissions"
            Write-Log "  2. Use Manage-SitesSelectedGUI.ps1 to assign Sites.Selected permissions to specific sites"
            Write-Log "  3. Store the PFX file and password securely"
        }
        catch {
            Write-Log "Registration failed: $($_.Exception.Message)" "ERROR"
            [System.Windows.MessageBox]::Show(
                "Failed to register the application.`n`n$($_.Exception.Message)`n`nEnsure you sign in with Global Admin or Application Admin credentials.",
                "Registration Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
        }
        finally {
            $btnRegister.IsEnabled = $true
            # Cleanup temp files
            Remove-Item $script:resultFile -ErrorAction SilentlyContinue
            $script:bgProcess  = $null
        }
    })
    $script:pollTimer.Start()
})

# ═════════════════════════════════════════════════════════════════════════════
#  Show Window
# ═════════════════════════════════════════════════════════════════════════════
Write-Log "Register Sites.Selected App ready."
$window.ShowDialog() | Out-Null
