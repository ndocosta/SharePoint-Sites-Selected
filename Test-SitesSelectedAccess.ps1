<#
    DISCLAIMER: This script is provided "AS IS" without warranty of any kind.
    Use it at your own risk. The author is not responsible for any damage or
    data loss caused by using this script. Always test in a non-production
    environment before deploying to production.
#>
#Requires -Modules PnP.PowerShell
<#
.SYNOPSIS
    GUI tool to test Sites.Selected app permissions by performing read, create, upload,
    and delete operations on a SharePoint site using certificate-based authentication.

.DESCRIPTION
    Provides a WPF graphical interface to connect to a SharePoint site using an Entra ID
    app registration with Sites.Selected permissions and certificate authentication,
    then run selectable tests:

    - READ   : Get site properties, list all lists, read Documents library
    - CREATE : Create a test list with sample items, read them back
    - UPLOAD : Upload a test file to the Documents library, verify it
    - DELETE : Remove the test list and uploaded file

    Each test reports PASS/FAIL in a results grid so you can verify which permission
    levels (Read, Write, Manage, FullControl) are effective for the app on that site.

.EXAMPLE
    .\Test-SitesSelectedAccess.ps1
#>

$ErrorActionPreference = 'Stop'

# ── Check PnP.PowerShell ───────────────────────────────────────────────────
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
Add-Type -AssemblyName System.Windows.Forms

# ═════════════════════════════════════════════════════════════════════════════
#  XAML
# ═════════════════════════════════════════════════════════════════════════════
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Sites.Selected - Test App Permissions"
        Width="820" Height="720"
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
        <Style TargetType="TextBlock">
            <Setter Property="Margin" Value="4"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
            <Setter Property="FontSize" Value="12"/>
        </Style>
        <Style TargetType="CheckBox">
            <Setter Property="Margin" Value="4"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
            <Setter Property="FontSize" Value="12"/>
        </Style>
    </Window.Resources>
    <DockPanel>

        <!-- ── Connection Panel ─────────────────────────────────────── -->
        <Border DockPanel.Dock="Top" Margin="10,10,10,4" Padding="10,8"
                BorderBrush="#CCCCCC" BorderThickness="1" CornerRadius="4"
                Background="White">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="130"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <TextBlock Text="Site URL:" FontWeight="SemiBold"/>
                <TextBox Name="txtSiteUrl" Grid.Column="1" Grid.ColumnSpan="2"
                         Text="https://contoso.sharepoint.com/sites/YourSite"/>

                <TextBlock Text="Client ID:" Grid.Row="1" FontWeight="SemiBold"/>
                <TextBox Name="txtClientId" Grid.Row="1" Grid.Column="1" Grid.ColumnSpan="2"/>

                <TextBlock Text="Certificate:" Grid.Row="2" FontWeight="SemiBold"/>
                <Grid Grid.Row="2" Grid.Column="1" Grid.ColumnSpan="2">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <RadioButton Name="rbPfx" Content="PFX File:" IsChecked="True"
                                 VerticalAlignment="Center" Margin="4"/>
                    <TextBox Name="txtCertPath" Grid.Column="1"/>
                    <Button Name="btnBrowseCert" Grid.Column="2" Content="..."
                            Padding="8,4" FontWeight="Bold"/>
                    <RadioButton Name="rbThumbprint" Grid.Column="3" Content="Thumbprint:"
                                 VerticalAlignment="Center" Margin="12,4,4,4"/>
                    <TextBox Name="txtThumbprint" Grid.Column="4" Width="200"/>
                </Grid>
            </Grid>
        </Border>

        <!-- ── Action Buttons ───────────────────────────────────────── -->
        <Border DockPanel.Dock="Top" Margin="10,4,10,4" Padding="8,6"
                BorderBrush="#CCCCCC" BorderThickness="1" CornerRadius="4"
                Background="White">
            <DockPanel>
                <StackPanel Orientation="Horizontal" DockPanel.Dock="Left">
                    <Button Name="btnConnect" Content="Connect"
                            Background="#0078D4" Foreground="White" FontWeight="Bold"/>
                    <Button Name="btnDisconnect" Content="Disconnect" IsEnabled="False"/>
                    <TextBlock Name="txtStatus" FontWeight="Bold" Foreground="Red"
                               Text="  Not Connected" Margin="8,4"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal" DockPanel.Dock="Right"
                            HorizontalAlignment="Right">
                    <CheckBox Name="chkRead" Content="Read" IsChecked="True"/>
                    <CheckBox Name="chkCreate" Content="Create" IsChecked="True"/>
                    <CheckBox Name="chkUpload" Content="Upload" IsChecked="True"/>
                    <CheckBox Name="chkDelete" Content="Delete" IsChecked="True"/>
                    <CheckBox Name="chkSkipCleanup" Content="Skip Cleanup" Margin="12,4,4,4"/>
                    <Button Name="btnRunTests" Content="Run Selected" IsEnabled="False"
                            Background="#107C10" Foreground="White" FontWeight="Bold"/>
                    <Button Name="btnRunAll" Content="Run All" IsEnabled="False"
                            Background="#0078D4" Foreground="White" FontWeight="Bold"/>
                </StackPanel>
            </DockPanel>
        </Border>

        <!-- ── Results Grid ─────────────────────────────────────────── -->
        <Border DockPanel.Dock="Top" Margin="10,4,10,0" Padding="4"
                BorderBrush="#CCCCCC" BorderThickness="1" CornerRadius="4"
                Background="White">
            <DockPanel>
                <StackPanel DockPanel.Dock="Top" Orientation="Horizontal" Margin="4,2">
                    <TextBlock Text="Test Results" FontWeight="SemiBold"/>
                    <TextBlock Name="txtSummary" Foreground="Gray" Margin="16,4,4,4"/>
                    <Button Name="btnClearResults" Content="Clear" FontSize="10"
                            Padding="8,2" Margin="8,0,0,0" VerticalAlignment="Center"/>
                </StackPanel>
                <DataGrid Name="dgResults" AutoGenerateColumns="False"
                          IsReadOnly="True" SelectionMode="Single"
                          Background="White" GridLinesVisibility="Horizontal"
                          HeadersVisibility="Column" AlternatingRowBackground="#F9F9F9"
                          Height="200">
                    <DataGrid.RowStyle>
                        <Style TargetType="DataGridRow">
                            <Style.Triggers>
                                <DataTrigger Binding="{Binding Status}" Value="PASS">
                                    <Setter Property="Background" Value="#E6F4EA"/>
                                    <Setter Property="Foreground" Value="#1B7D2C"/>
                                </DataTrigger>
                                <DataTrigger Binding="{Binding Status}" Value="FAIL">
                                    <Setter Property="Background" Value="#FDEDED"/>
                                    <Setter Property="Foreground" Value="#C62828"/>
                                </DataTrigger>
                                <DataTrigger Binding="{Binding Status}" Value="SKIP">
                                    <Setter Property="Background" Value="#FFF8E1"/>
                                    <Setter Property="Foreground" Value="#F57F17"/>
                                </DataTrigger>
                            </Style.Triggers>
                        </Style>
                    </DataGrid.RowStyle>
                    <DataGrid.Columns>
                        <DataGridTextColumn Header="Result" Binding="{Binding Status}" Width="60"/>
                        <DataGridTextColumn Header="Category" Binding="{Binding Category}" Width="80"/>
                        <DataGridTextColumn Header="Test" Binding="{Binding TestName}" Width="200"/>
                        <DataGridTextColumn Header="Requires" Binding="{Binding RequiredPerm}" Width="80"/>
                        <DataGridTextColumn Header="Detail" Binding="{Binding Detail}" Width="*"/>
                    </DataGrid.Columns>
                </DataGrid>
            </DockPanel>
        </Border>

        <!-- ── Log Panel ────────────────────────────────────────────── -->
        <Border Margin="10,4,10,10" Padding="4"
                BorderBrush="#CCCCCC" BorderThickness="1" CornerRadius="4"
                Background="White">
            <DockPanel>
                <StackPanel DockPanel.Dock="Top" Orientation="Horizontal">
                    <TextBlock Text="Log" FontWeight="SemiBold" Margin="4,2"/>
                    <Button Name="btnClearLog" Content="Clear" FontSize="10"
                            Padding="8,2" Margin="8,0,0,0" VerticalAlignment="Center"/>
                </StackPanel>
                <TextBox Name="txtLog" IsReadOnly="True"
                         VerticalScrollBarVisibility="Auto"
                         HorizontalScrollBarVisibility="Auto"
                         TextWrapping="Wrap" FontFamily="Consolas" FontSize="11"
                         Background="#FAFAFA" BorderThickness="0"/>
            </DockPanel>
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

$txtSiteUrl      = $controls['txtSiteUrl']
$txtClientId     = $controls['txtClientId']
$rbPfx           = $controls['rbPfx']
$rbThumbprint    = $controls['rbThumbprint']
$txtCertPath     = $controls['txtCertPath']
$btnBrowseCert   = $controls['btnBrowseCert']
$txtThumbprint   = $controls['txtThumbprint']
$btnConnect      = $controls['btnConnect']
$btnDisconnect   = $controls['btnDisconnect']
$txtStatus       = $controls['txtStatus']
$chkRead         = $controls['chkRead']
$chkCreate       = $controls['chkCreate']
$chkUpload       = $controls['chkUpload']
$chkDelete       = $controls['chkDelete']
$chkSkipCleanup  = $controls['chkSkipCleanup']
$btnRunTests     = $controls['btnRunTests']
$btnRunAll       = $controls['btnRunAll']
$dgResults       = $controls['dgResults']
$txtSummary      = $controls['txtSummary']
$btnClearResults = $controls['btnClearResults']
$txtLog          = $controls['txtLog']
$btnClearLog     = $controls['btnClearLog']

# ═════════════════════════════════════════════════════════════════════════════
#  State
# ═════════════════════════════════════════════════════════════════════════════
$script:isConnected = $false
$script:testListName = $null
$script:testFileName = $null
$script:listCreated  = $false
$script:fileUploaded = $false
$script:defaultSiteUrl = $txtSiteUrl.Text

$resultItems = [System.Collections.ObjectModel.ObservableCollection[PSObject]]::new()
$dgResults.ItemsSource = $resultItems

# ═════════════════════════════════════════════════════════════════════════════
#  Helper functions
# ═════════════════════════════════════════════════════════════════════════════
function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "HH:mm:ss"
    $txtLog.AppendText("[$timestamp] $Message`r`n")
    $txtLog.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [Action]{ }
    )
}

function Add-TestResult {
    param(
        [string]$Category,
        [string]$TestName,
        [bool]$Success,
        [string]$Detail = "",
        [string]$RequiredPerm = "",
        [switch]$Skipped
    )
    $status = if ($Skipped) { "SKIP" } elseif ($Success) { "PASS" } else { "FAIL" }
    $resultItems.Add([PSCustomObject]@{
        Status       = $status
        Category     = $Category
        TestName     = $TestName
        RequiredPerm = $RequiredPerm
        Detail       = $Detail
    })
    Write-Log "$status - $TestName$(if ($Detail) { ": $Detail" })"
    Update-Summary
}

function Update-Summary {
    $total   = $resultItems.Count
    $passed  = ($resultItems | Where-Object { $_.Status -eq "PASS" }).Count
    $skipped = ($resultItems | Where-Object { $_.Status -eq "SKIP" }).Count
    $failed  = ($resultItems | Where-Object { $_.Status -eq "FAIL" }).Count
    $txtSummary.Text = "Total: $total  |  Passed: $passed  |  Skipped: $skipped  |  Failed: $failed"
}

function Set-Connected {
    param([bool]$Connected)
    $script:isConnected = $Connected
    $btnConnect.IsEnabled    = -not $Connected
    $btnDisconnect.IsEnabled = $Connected
    $btnRunTests.IsEnabled   = $Connected
    $btnRunAll.IsEnabled     = $Connected
    if ($Connected) {
        $txtStatus.Text       = "  Connected"
        $txtStatus.Foreground = [System.Windows.Media.Brushes]::Green
    } else {
        $txtStatus.Text       = "  Not Connected"
        $txtStatus.Foreground = [System.Windows.Media.Brushes]::Red
    }
}

# ═════════════════════════════════════════════════════════════════════════════
#  Browse certificate
# ═════════════════════════════════════════════════════════════════════════════
$btnBrowseCert.Add_Click({
    $rbPfx.IsChecked = $true
    $dialog = [System.Windows.Forms.OpenFileDialog]::new()
    $dialog.Title  = "Select PFX Certificate"
    $dialog.Filter = "PFX Files (*.pfx)|*.pfx|All Files (*.*)|*.*"
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtCertPath.Text = $dialog.FileName
    }
})

$txtThumbprint.Add_GotFocus({
    $rbThumbprint.IsChecked = $true
})

# ═════════════════════════════════════════════════════════════════════════════
#  Connect
# ═════════════════════════════════════════════════════════════════════════════
$btnConnect.Add_Click({
    $siteUrl  = $txtSiteUrl.Text.Trim()
    $clientId = $txtClientId.Text.Trim()

    if ($siteUrl -eq $script:defaultSiteUrl) {
        [System.Windows.MessageBox]::Show("Please change the Site URL to your actual SharePoint site before connecting.",
            "Default URL", "OK", "Warning")
        $txtSiteUrl.Focus()
        return
    }

    if (-not $siteUrl -or -not $clientId) {
        [System.Windows.MessageBox]::Show("Please fill in Site URL and Client ID.",
            "Missing Fields", "OK", "Warning")
        return
    }

    # Extract tenant from site URL (e.g. https://contoso.sharepoint.com/... → contoso.onmicrosoft.com)
    if ($siteUrl -match 'https://([^.]+)\.sharepoint\.com') {
        $tenant = "$($Matches[1]).onmicrosoft.com"
    } else {
        [System.Windows.MessageBox]::Show("Could not extract tenant from Site URL. Expected format: https://tenant.sharepoint.com/...",
            "Invalid URL", "OK", "Error")
        return
    }

    try {
        Write-Log "Connecting to $siteUrl ..."
        $connectParams = @{
            Url      = $siteUrl
            ClientId = $clientId
            Tenant   = $tenant
        }

        if ($rbPfx.IsChecked) {
            $certPath = $txtCertPath.Text.Trim()
            if (-not $certPath) {
                [System.Windows.MessageBox]::Show("Please select a PFX certificate file.",
                    "Missing Certificate", "OK", "Warning")
                return
            }
            if (-not (Test-Path $certPath)) {
                [System.Windows.MessageBox]::Show("Certificate file not found: $certPath",
                    "File Not Found", "OK", "Error")
                return
            }

            $credDialog = [System.Windows.Window]::new()
            $credDialog.Title = "Certificate Password"
            $credDialog.Width = 380
            $credDialog.Height = 150
            $credDialog.WindowStartupLocation = "CenterOwner"
            $credDialog.Owner = $window
            $credDialog.ResizeMode = "NoResize"
            $credPanel = [System.Windows.Controls.StackPanel]::new()
            $credPanel.Margin = [System.Windows.Thickness]::new(16)
            $credLabel = [System.Windows.Controls.TextBlock]::new()
            $credLabel.Text = "Enter PFX certificate password:"
            $credLabel.Margin = [System.Windows.Thickness]::new(0,0,0,8)
            $credBox = [System.Windows.Controls.PasswordBox]::new()
            $credBox.Margin = [System.Windows.Thickness]::new(0,0,0,12)
            $credBtn = [System.Windows.Controls.Button]::new()
            $credBtn.Content = "OK"
            $credBtn.Width = 80
            $credBtn.HorizontalAlignment = "Right"
            $credBtn.IsDefault = $true
            $credBtn.Add_Click({ $credDialog.DialogResult = $true; $credDialog.Close() })
            $credPanel.Children.Add($credLabel) | Out-Null
            $credPanel.Children.Add($credBox) | Out-Null
            $credPanel.Children.Add($credBtn) | Out-Null
            $credDialog.Content = $credPanel

            if ($credDialog.ShowDialog() -ne $true) {
                Write-Log "Connection cancelled by user."
                return
            }

            $secPass = $credBox.SecurePassword
            $connectParams["CertificatePath"]     = $certPath
            $connectParams["CertificatePassword"] = $secPass
        } else {
            $thumb = $txtThumbprint.Text.Trim()
            if (-not $thumb) {
                [System.Windows.MessageBox]::Show("Please enter a certificate thumbprint.",
                    "Missing Thumbprint", "OK", "Warning")
                return
            }
            $connectParams["Thumbprint"] = $thumb
        }

        Connect-PnPOnline @connectParams
        Set-Connected $true
        Write-Log "Connected successfully to $siteUrl"
    }
    catch {
        Write-Log "ERROR: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show(
            "Failed to connect:`n`n$($_.Exception.Message)`n`nPossible causes:`n- App not granted Sites.Selected on this site`n- Admin consent not granted`n- Wrong certificate or Client ID",
            "Connection Failed", "OK", "Error")
    }
})

# ═════════════════════════════════════════════════════════════════════════════
#  Disconnect
# ═════════════════════════════════════════════════════════════════════════════
$btnDisconnect.Add_Click({
    try { Disconnect-PnPOnline } catch { }
    Set-Connected $false
    Write-Log "Disconnected."
})

# ═════════════════════════════════════════════════════════════════════════════
#  Test runner functions
# ═════════════════════════════════════════════════════════════════════════════

function Invoke-ReadTests {
    Write-Log "── Running READ tests ──"

    # Read site properties
    try {
        $web = Get-PnPWeb -Includes Title, Url, Created
        Add-TestResult "READ" "Get site properties" $true "Title: $($web.Title)" -RequiredPerm "Read"
    }
    catch {
        Add-TestResult "READ" "Get site properties" $false $_.Exception.Message -RequiredPerm "Read"
    }

    # List existing lists
    try {
        $lists = Get-PnPList
        $listNames = ($lists | Select-Object -First 5 -ExpandProperty Title) -join ", "
        Add-TestResult "READ" "List all lists" $true "Found $($lists.Count) lists ($listNames...)" -RequiredPerm "Read"
    }
    catch {
        Add-TestResult "READ" "List all lists" $false $_.Exception.Message -RequiredPerm "Read"
    }

    # Read Documents library
    try {
        $items = Get-PnPListItem -List "Documents" -PageSize 5
        Add-TestResult "READ" "Read Documents library" $true "Found $($items.Count) item(s)" -RequiredPerm "Read"
    }
    catch {
        Add-TestResult "READ" "Read Documents library" $false $_.Exception.Message -RequiredPerm "Read"
    }
}

function Invoke-CreateTests {
    Write-Log "── Running CREATE tests ──"
    $script:testListName = "PnP-SitesSelected-Test-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
    $script:listCreated = $false

    # Create test list
    try {
        New-PnPList -Title $script:testListName -Template GenericList | Out-Null
        $script:listCreated = $true
        Add-TestResult "CREATE" "Create test list" $true "List: $($script:testListName)" -RequiredPerm "Manage"
    }
    catch {
        Add-TestResult "CREATE" "Create test list" $false $_.Exception.Message -RequiredPerm "Manage"
    }

    # Add items
    if ($script:listCreated) {
        try {
            Add-PnPListItem -List $script:testListName -Values @{ Title = "Test Item 1 - $(Get-Date -Format 'HH:mm:ss')" } | Out-Null
            Add-PnPListItem -List $script:testListName -Values @{ Title = "Test Item 2 - $(Get-Date -Format 'HH:mm:ss')" } | Out-Null
            Add-PnPListItem -List $script:testListName -Values @{ Title = "Test Item 3 - $(Get-Date -Format 'HH:mm:ss')" } | Out-Null
            Add-TestResult "CREATE" "Add list items" $true "Added 3 items to $($script:testListName)" -RequiredPerm "Write"
        }
        catch {
            Add-TestResult "CREATE" "Add list items" $false $_.Exception.Message -RequiredPerm "Write"
        }

        # Read back items
        try {
            $createdItems = Get-PnPListItem -List $script:testListName
            Add-TestResult "CREATE" "Read created items" $true "Retrieved $($createdItems.Count) item(s)" -RequiredPerm "Read"
        }
        catch {
            Add-TestResult "CREATE" "Read created items" $false $_.Exception.Message -RequiredPerm "Read"
        }
    } else {
        Add-TestResult "CREATE" "Add list items" $false "List creation failed" -RequiredPerm "Write" -Skipped
    }
}

function Invoke-UploadTests {
    Write-Log "── Running UPLOAD tests ──"
    $script:testFileName = "PnP-SitesSelected-TestFile-$(Get-Date -Format 'yyyyMMdd-HHmmss').txt"
    $script:fileUploaded = $false
    $tempFilePath = Join-Path $env:TEMP $script:testFileName

    # Upload file
    try {
        $fileContent = @"
Sites.Selected Permission Test File
====================================
Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Site:      $($txtSiteUrl.Text)
Client ID: $($txtClientId.Text)
This file was created by Test-SitesSelectedAccess.ps1 to verify upload permissions.
"@
        $fileContent | Out-File -FilePath $tempFilePath -Encoding UTF8
        Add-PnPFile -Path $tempFilePath -Folder "Shared Documents" | Out-Null
        $script:fileUploaded = $true
        Add-TestResult "UPLOAD" "Upload file to Documents" $true "File: $($script:testFileName)" -RequiredPerm "Write"
    }
    catch {
        Add-TestResult "UPLOAD" "Upload file to Documents" $false $_.Exception.Message -RequiredPerm "Write"
    }
    finally {
        if (Test-Path $tempFilePath) { Remove-Item $tempFilePath -Force }
    }

    # Verify
    if ($script:fileUploaded) {
        try {
            $uploadedFile = Get-PnPFile -Url "Shared Documents/$($script:testFileName)" -AsFileObject
            Add-TestResult "UPLOAD" "Verify uploaded file" $true "Size: $($uploadedFile.Length) bytes" -RequiredPerm "Read"
        }
        catch {
            Add-TestResult "UPLOAD" "Verify uploaded file" $false $_.Exception.Message -RequiredPerm "Read"
        }
    }
}

function Invoke-DeleteTests {
    Write-Log "── Running DELETE tests ──"

    if ($chkSkipCleanup.IsChecked) {
        Add-TestResult "DELETE" "Cleanup" $false "Skip Cleanup is checked" -RequiredPerm "Write" -Skipped
        if ($script:listCreated) { Write-Log "Test list '$($script:testListName)' left in place" }
        if ($script:fileUploaded) { Write-Log "Test file '$($script:testFileName)' left in Documents" }
        return
    }

    # Delete uploaded file
    if ($script:fileUploaded) {
        try {
            $webUrl = (Get-PnPWeb | Select-Object -ExpandProperty ServerRelativeUrl)
            Remove-PnPFile -ServerRelativeUrl "$webUrl/Shared Documents/$($script:testFileName)" -Force
            Add-TestResult "DELETE" "Delete uploaded file" $true "Removed $($script:testFileName)" -RequiredPerm "Write"
        }
        catch {
            Add-TestResult "DELETE" "Delete uploaded file" $false $_.Exception.Message -RequiredPerm "Write"
        }
    } else {
        Add-TestResult "DELETE" "Delete uploaded file" $false "No file was uploaded" -RequiredPerm "Write" -Skipped
    }

    # Delete test list
    if ($script:listCreated) {
        try {
            Remove-PnPList -Identity $script:testListName -Force
            Add-TestResult "DELETE" "Delete test list" $true "Removed $($script:testListName)" -RequiredPerm "Manage"
        }
        catch {
            Add-TestResult "DELETE" "Delete test list" $false $_.Exception.Message -RequiredPerm "Manage"
        }
    } else {
        Add-TestResult "DELETE" "Delete test list" $false "No list was created" -RequiredPerm "Manage" -Skipped
    }
}

# ═════════════════════════════════════════════════════════════════════════════
#  Run Tests (selected)
# ═════════════════════════════════════════════════════════════════════════════
$btnRunTests.Add_Click({
    if ($txtSiteUrl.Text.Trim() -eq $script:defaultSiteUrl) {
        [System.Windows.MessageBox]::Show("Please change the Site URL to your actual SharePoint site before running tests.",
            "Default URL", "OK", "Warning")
        $txtSiteUrl.Focus()
        return
    }

    if (-not $chkRead.IsChecked -and -not $chkCreate.IsChecked -and
        -not $chkUpload.IsChecked -and -not $chkDelete.IsChecked) {
        [System.Windows.MessageBox]::Show("Please select at least one test category.",
            "No Tests Selected", "OK", "Warning")
        return
    }

    $resultItems.Clear()
    $script:listCreated  = $false
    $script:fileUploaded = $false
    Write-Log "========== Starting selected tests =========="

    if ($chkRead.IsChecked)   { Invoke-ReadTests }
    if ($chkCreate.IsChecked) { Invoke-CreateTests }
    if ($chkUpload.IsChecked) { Invoke-UploadTests }
    if ($chkDelete.IsChecked) { Invoke-DeleteTests }

    Write-Log "========== Tests complete =========="
    Show-FinalSummary
})

# ═════════════════════════════════════════════════════════════════════════════
#  Run All
# ═════════════════════════════════════════════════════════════════════════════
$btnRunAll.Add_Click({
    if ($txtSiteUrl.Text.Trim() -eq $script:defaultSiteUrl) {
        [System.Windows.MessageBox]::Show("Please change the Site URL to your actual SharePoint site before running tests.",
            "Default URL", "OK", "Warning")
        $txtSiteUrl.Focus()
        return
    }

    $resultItems.Clear()
    $script:listCreated  = $false
    $script:fileUploaded = $false
    $chkRead.IsChecked    = $true
    $chkCreate.IsChecked  = $true
    $chkUpload.IsChecked  = $true
    $chkDelete.IsChecked  = $true
    Write-Log "========== Starting ALL tests =========="

    Invoke-ReadTests
    Invoke-CreateTests
    Invoke-UploadTests
    Invoke-DeleteTests

    Write-Log "========== All tests complete =========="
    Show-FinalSummary
})

function Show-FinalSummary {
    $passed  = ($resultItems | Where-Object { $_.Status -eq "PASS" }).Count
    $skipped = ($resultItems | Where-Object { $_.Status -eq "SKIP" }).Count
    $failed  = ($resultItems | Where-Object { $_.Status -eq "FAIL" }).Count
    $total   = $resultItems.Count

    if ($failed -eq 0 -and $skipped -eq 0) {
        Write-Log "All $total tests passed! The app has full access to this site."
    } elseif ($failed -eq 0) {
        Write-Log "Passed: $passed, Skipped: $skipped - All executed tests passed."
    } elseif ($passed -gt 0) {
        Write-Log "Passed: $passed, Skipped: $skipped, Failed: $failed - Check the app's permission level."
    } else {
        Write-Log "All $total tests failed. The app may not have any access to this site."
    }
}

# ═════════════════════════════════════════════════════════════════════════════
#  Clear buttons
# ═════════════════════════════════════════════════════════════════════════════
$btnClearResults.Add_Click({ $resultItems.Clear(); $txtSummary.Text = "" })
$btnClearLog.Add_Click({ $txtLog.Clear() })

# ═════════════════════════════════════════════════════════════════════════════
#  Cleanup on window close
# ═════════════════════════════════════════════════════════════════════════════
$window.Add_Closing({
    if ($script:isConnected) {
        try { Disconnect-PnPOnline } catch { }
    }
})

# ═════════════════════════════════════════════════════════════════════════════
#  Show window
# ═════════════════════════════════════════════════════════════════════════════
$window.ShowDialog() | Out-Null
