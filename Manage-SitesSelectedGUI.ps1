<#
    DISCLAIMER: This script is provided "AS IS" without warranty of any kind.
    Use it at your own risk. The author is not responsible for any damage or
    data loss caused by using this script. Always test in a non-production
    environment before deploying to production.
#>
#Requires -Modules PnP.PowerShell
<#
.SYNOPSIS
    GUI tool to manage Sites.Selected permissions for Entra ID app registrations.

.DESCRIPTION
    Provides a WPF graphical interface to:
    - Grant Sites.Selected permissions to an app on one or more site collections
    - View and edit app permissions on a specific site collection
    - Search all tenant sites for permissions granted to a specific app
    - Revoke app permissions from site collections

    Requires connection as a SharePoint Administrator with Sites.FullControl.All
    delegated permission (handled via interactive login with PnP Management Shell).

.EXAMPLE
    .\Manage-SitesSelectedGUI.ps1
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
        Title="Sites.Selected Permissions Manager"
        Width="960" Height="740"
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
        <Style TargetType="ComboBox">
            <Setter Property="Margin" Value="4"/>
            <Setter Property="Padding" Value="4,3"/>
            <Setter Property="FontSize" Value="12"/>
        </Style>
    </Window.Resources>
    <DockPanel>
        <!-- ── Connection Panel ─────────────────────────────────────── -->
        <Border DockPanel.Dock="Top" Margin="10,10,10,4" Padding="10,8"
                BorderBrush="#CCCCCC" BorderThickness="1" CornerRadius="4"
                Background="White">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Text="SharePoint Admin URL:" FontWeight="SemiBold"/>
                <TextBox Name="txtAdminUrl" Grid.Column="1"
                         Text="https://contoso-admin.sharepoint.com"/>
                <Button Name="btnConnect" Grid.Column="2" Content="&#x1F310; Connect"
                        Background="#0078D4" Foreground="White" FontWeight="Bold"/>
                <Button Name="btnDisconnect" Grid.Column="3" Content="Disconnect"
                        IsEnabled="False"/>
                <TextBlock Name="txtStatus" Grid.Column="4" FontWeight="Bold"
                           Foreground="Red" Text="  ● Not Connected"/>
            </Grid>
        </Border>

        <!-- ── Log Panel ────────────────────────────────────────────── -->
        <Border DockPanel.Dock="Bottom" Margin="10,4,10,10" Padding="4"
                BorderBrush="#CCCCCC" BorderThickness="1" CornerRadius="4"
                Background="White" Height="140">
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

        <!-- ── Tab Control ──────────────────────────────────────────── -->
        <TabControl Margin="10,4" FontSize="12">

            <!-- ============ TAB 1: Grant Permissions ============ -->
            <TabItem Header="  Grant Permissions  " FontWeight="SemiBold">
                <Grid Margin="12">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="130"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <TextBlock Text="App ID (GUID):" FontWeight="SemiBold"/>
                    <TextBox Name="txtGrantAppId" Grid.Column="1"/>

                    <TextBlock Text="Display Name:" Grid.Row="1" FontWeight="SemiBold"/>
                    <TextBox Name="txtGrantDisplayName" Grid.Row="1" Grid.Column="1"/>

                    <TextBlock Text="Permission:" Grid.Row="2" FontWeight="SemiBold"/>
                    <ComboBox Name="cmbGrantPermission" Grid.Row="2" Grid.Column="1"
                              Width="160" HorizontalAlignment="Left">
                        <ComboBoxItem Content="Read" IsSelected="True"/>
                        <ComboBoxItem Content="Write"/>
                        <ComboBoxItem Content="Manage"/>
                        <ComboBoxItem Content="FullControl"/>
                    </ComboBox>

                    <TextBlock Text="Site URL(s):" Grid.Row="3" FontWeight="SemiBold"
                               VerticalAlignment="Top" Margin="4,8,4,4"/>
                    <TextBlock Text="(one URL per line)" Grid.Row="3" Grid.Column="1"
                               Foreground="Gray" FontSize="10" HorizontalAlignment="Right"/>

                    <TextBox Name="txtGrantSiteUrls" Grid.Row="4" Grid.Column="0"
                             Grid.ColumnSpan="2" AcceptsReturn="True"
                             VerticalScrollBarVisibility="Auto"
                             TextWrapping="NoWrap" FontFamily="Consolas"/>

                    <Button Name="btnGrant" Grid.Row="5" Grid.ColumnSpan="2"
                            Content="Grant Permissions" IsEnabled="False"
                            HorizontalAlignment="Right" Margin="4,8"
                            Background="#107C10" Foreground="White" FontWeight="Bold"/>
                </Grid>
            </TabItem>

            <!-- ============ TAB 2: Manage by Site ============ -->
            <TabItem Header="  Manage by Site  " FontWeight="SemiBold">
                <DockPanel Margin="12">
                    <Grid DockPanel.Dock="Top">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Text="Site URL:" FontWeight="SemiBold"/>
                        <TextBox Name="txtSiteUrl" Grid.Column="1"/>
                        <Button Name="btnLoadSitePerms" Grid.Column="2"
                                Content="Load Permissions" IsEnabled="False"
                                Background="#0078D4" Foreground="White" FontWeight="Bold"/>
                    </Grid>

                    <Border DockPanel.Dock="Bottom" Margin="0,8,0,0" Padding="8"
                            Background="#F0F0F0" CornerRadius="4">
                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                            <TextBlock Text="Change to:" FontWeight="SemiBold"/>
                            <ComboBox Name="cmbEditPermission" Width="130">
                                <ComboBoxItem Content="Read"/>
                                <ComboBoxItem Content="Write"/>
                                <ComboBoxItem Content="Manage"/>
                                <ComboBoxItem Content="FullControl"/>
                            </ComboBox>
                            <Button Name="btnEditPerm" Content="Update Selected"
                                    IsEnabled="False"
                                    Background="#0078D4" Foreground="White"/>
                            <Button Name="btnRevokePerm" Content="Revoke Selected"
                                    IsEnabled="False"
                                    Background="#D13438" Foreground="White"/>
                        </StackPanel>
                    </Border>

                    <DataGrid Name="dgSitePerms" AutoGenerateColumns="False"
                              IsReadOnly="True" SelectionMode="Single"
                              Margin="0,8" Background="White"
                              GridLinesVisibility="Horizontal"
                              HeadersVisibility="Column"
                              AlternatingRowBackground="#F9F9F9">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="App ID"
                                Binding="{Binding AppId}" Width="270"/>
                            <DataGridTextColumn Header="Permission"
                                Binding="{Binding Roles}" Width="100"/>
                            <DataGridTextColumn Header="Permission ID"
                                Binding="{Binding Id}" Width="*"/>
                        </DataGrid.Columns>
                    </DataGrid>
                </DockPanel>
            </TabItem>

            <!-- ============ TAB 3: Search by App ============ -->
            <TabItem Header="  Search by App  " FontWeight="SemiBold">
                <DockPanel Margin="12">
                    <Grid DockPanel.Dock="Top">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Text="App ID or Name:" FontWeight="SemiBold"/>
                        <TextBox Name="txtSearchApp" Grid.Column="1"/>
                        <Button Name="btnSearchApp" Grid.Column="2"
                                Content="Search All Sites" IsEnabled="False"
                                Background="#0078D4" Foreground="White" FontWeight="Bold"/>
                        <TextBlock Name="txtSearchProgress" Grid.Column="3"
                                   Foreground="Gray" MinWidth="140"
                                   HorizontalAlignment="Right"/>
                    </Grid>

                    <DataGrid Name="dgAppSites" AutoGenerateColumns="False"
                              IsReadOnly="True" Margin="0,8"
                              Background="White"
                              GridLinesVisibility="Horizontal"
                              HeadersVisibility="Column"
                              AlternatingRowBackground="#F9F9F9">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Site URL"
                                Binding="{Binding SiteUrl}" Width="300"/>
                            <DataGridTextColumn Header="Site Title"
                                Binding="{Binding SiteTitle}" Width="160"/>
                            <DataGridTextColumn Header="App ID"
                                Binding="{Binding AppId}" Width="270"/>
                            <DataGridTextColumn Header="Permission"
                                Binding="{Binding Roles}" Width="100"/>
                            <DataGridTextColumn Header="Permission ID"
                                Binding="{Binding PermissionId}" Width="*"/>
                        </DataGrid.Columns>
                    </DataGrid>
                </DockPanel>
            </TabItem>

        </TabControl>
    </DockPanel>
</Window>
"@

# ═════════════════════════════════════════════════════════════════════════════
#  Load XAML and get named controls
# ═════════════════════════════════════════════════════════════════════════════
$reader = [System.Xml.XmlNodeReader]::new($xaml)
$window = [System.Windows.Markup.XamlReader]::Load($reader)

# Retrieve named elements
$controls = @{}
$xaml.SelectNodes('//*[@Name]') | ForEach-Object {
    $controls[$_.Name] = $window.FindName($_.Name)
}

$txtAdminUrl       = $controls['txtAdminUrl']
$btnConnect        = $controls['btnConnect']
$btnDisconnect     = $controls['btnDisconnect']
$txtStatus         = $controls['txtStatus']
$txtLog            = $controls['txtLog']
$btnClearLog       = $controls['btnClearLog']

$txtGrantAppId     = $controls['txtGrantAppId']
$txtGrantDisplayName = $controls['txtGrantDisplayName']
$cmbGrantPermission  = $controls['cmbGrantPermission']
$txtGrantSiteUrls  = $controls['txtGrantSiteUrls']
$btnGrant          = $controls['btnGrant']

$txtSiteUrl        = $controls['txtSiteUrl']
$btnLoadSitePerms  = $controls['btnLoadSitePerms']
$dgSitePerms       = $controls['dgSitePerms']
$cmbEditPermission = $controls['cmbEditPermission']
$btnEditPerm       = $controls['btnEditPerm']
$btnRevokePerm     = $controls['btnRevokePerm']

$txtSearchApp      = $controls['txtSearchApp']
$btnSearchApp      = $controls['btnSearchApp']
$txtSearchProgress = $controls['txtSearchProgress']
$dgAppSites        = $controls['dgAppSites']

# ═════════════════════════════════════════════════════════════════════════════
#  State
# ═════════════════════════════════════════════════════════════════════════════
$script:IsConnected = $false

# ═════════════════════════════════════════════════════════════════════════════
#  Helper Functions
# ═════════════════════════════════════════════════════════════════════════════

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "HH:mm:ss"
    $entry = "[$timestamp] [$Level] $Message`r`n"
    $txtLog.AppendText($entry)
    $txtLog.ScrollToEnd()
    # Allow UI to refresh
    $window.Dispatcher.Invoke(
        [Action]{},
        [System.Windows.Threading.DispatcherPriority]::Background
    )
}

function Test-IsConnected {
    if (-not $script:IsConnected) {
        [System.Windows.MessageBox]::Show(
            "Please connect to SharePoint first.",
            "Not Connected",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        )
        return $false
    }
    return $true
}

function Get-PermissionDetail {
    <#
    .SYNOPSIS
        Gets roles and app ID for a specific permission on a site using the
        Microsoft Graph API directly (via Invoke-PnPGraphMethod).
    #>
    param(
        [string]$PermissionId,
        [string]$SiteUrl
    )
    try {
        $uri = [System.Uri]$SiteUrl
        $hostname = $uri.Host
        $sitePath = $uri.AbsolutePath.TrimEnd('/')
        $graphUrl = "v1.0/sites/${hostname}:${sitePath}:/permissions/$PermissionId"

        $detail = Invoke-PnPGraphMethod -Url $graphUrl -Method Get

        $roles = "N/A"
        if ($detail.roles) {
            $roles = ($detail.roles -join ', ')
        }

        $appId = ""
        $identities = $detail.grantedToIdentitiesV2
        if (-not $identities) { $identities = $detail.grantedToIdentities }
        if ($identities) {
            $firstIdentity = $identities | Select-Object -First 1
            if ($firstIdentity.application) {
                $appId = $firstIdentity.application.id
            }
        }

        return @{ Roles = $roles; AppId = $appId }
    }
    catch {
        return @{ Roles = "Error"; AppId = "" }
    }
}

# ═════════════════════════════════════════════════════════════════════════════
#  Event Handlers
# ═════════════════════════════════════════════════════════════════════════════

# ── Connect ─────────────────────────────────────────────────────────────────
$btnConnect.Add_Click({
    $adminUrl = $txtAdminUrl.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($adminUrl)) {
        [System.Windows.MessageBox]::Show(
            "Please enter the SharePoint Admin URL.",
            "Missing URL",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        )
        return
    }

    Write-Log "Connecting to $adminUrl (interactive login)..."
    Write-Log "A browser window will open for authentication. The UI may briefly pause."
    $btnConnect.IsEnabled = $false

    # Flush pending UI updates before the blocking call
    $window.Dispatcher.Invoke(
        [Action]{},
        [System.Windows.Threading.DispatcherPriority]::Background
    )

    try {
        Connect-PnPOnline -Url $adminUrl -Interactive
        $script:IsConnected = $true
        $txtStatus.Text      = "  ● Connected"
        $txtStatus.Foreground = [System.Windows.Media.Brushes]::Green
        $btnDisconnect.IsEnabled   = $true
        $txtAdminUrl.IsEnabled     = $false
        $btnGrant.IsEnabled        = $true
        $btnLoadSitePerms.IsEnabled = $true
        $btnEditPerm.IsEnabled     = $true
        $btnRevokePerm.IsEnabled   = $true
        $btnSearchApp.IsEnabled    = $true
        Write-Log "Successfully connected to $adminUrl"
    }
    catch {
        $script:IsConnected = $false
        Write-Log "Connection failed: $($_.Exception.Message)" "ERROR"
        [System.Windows.MessageBox]::Show(
            "Failed to connect.`n`n$($_.Exception.Message)",
            "Connection Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
    finally {
        $btnConnect.IsEnabled = $true
    }
})

# ── Disconnect ──────────────────────────────────────────────────────────────
$btnDisconnect.Add_Click({
    try {
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
    } catch {}
    $script:IsConnected = $false
    $txtStatus.Text      = "  ● Not Connected"
    $txtStatus.Foreground = [System.Windows.Media.Brushes]::Red
    $btnDisconnect.IsEnabled   = $false
    $txtAdminUrl.IsEnabled     = $true
    $btnGrant.IsEnabled        = $false
    $btnLoadSitePerms.IsEnabled = $false
    $btnEditPerm.IsEnabled     = $false
    $btnRevokePerm.IsEnabled   = $false
    $btnSearchApp.IsEnabled    = $false
    Write-Log "Disconnected."
})

# ── Clear Log ───────────────────────────────────────────────────────────────
$btnClearLog.Add_Click({ $txtLog.Clear() })

# ═════════════════════════════════════════════════════════════════════════════
#  TAB 1: Grant Permissions
# ═════════════════════════════════════════════════════════════════════════════
$btnGrant.Add_Click({
    if (-not (Test-IsConnected)) { return }

    $appId       = $txtGrantAppId.Text.Trim()
    $displayName = $txtGrantDisplayName.Text.Trim()
    $permission  = $cmbGrantPermission.SelectedItem.Content
    $siteLines   = $txtGrantSiteUrls.Text.Trim() -split "`r?`n" |
                   Where-Object { $_.Trim() -ne '' } |
                   ForEach-Object { $_.Trim() }

    # Validate inputs
    if ([string]::IsNullOrWhiteSpace($appId)) {
        [System.Windows.MessageBox]::Show("App ID is required.", "Validation")
        return
    }
    $guidParsed = [guid]::Empty
    if (-not [guid]::TryParse($appId, [ref]$guidParsed)) {
        [System.Windows.MessageBox]::Show("App ID must be a valid GUID.", "Validation")
        return
    }
    if ([string]::IsNullOrWhiteSpace($displayName)) {
        [System.Windows.MessageBox]::Show("Display Name is required.", "Validation")
        return
    }
    if ($siteLines.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Enter at least one Site URL.", "Validation")
        return
    }

    $total   = $siteLines.Count
    $success = 0
    $failed  = 0

    Write-Log "Granting '$permission' permission for '$displayName' ($appId) on $total site(s)..."

    foreach ($siteUrl in $siteLines) {
        try {
            Grant-PnPEntraIDAppSitePermission `
                -AppId $appId `
                -DisplayName $displayName `
                -Permissions $permission `
                -Site $siteUrl

            $success++
            Write-Log "  ✓ Granted on $siteUrl"
        }
        catch {
            $failed++
            Write-Log "  ✗ Failed on $siteUrl : $($_.Exception.Message)" "ERROR"
        }
    }

    Write-Log "Grant complete. Success: $success, Failed: $failed."

    if ($failed -eq 0) {
        [System.Windows.MessageBox]::Show(
            "Successfully granted '$permission' permission on $success site(s).",
            "Grant Complete",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        )
    }
    else {
        [System.Windows.MessageBox]::Show(
            "Completed with errors.`nSuccess: $success, Failed: $failed`n`nCheck the log for details.",
            "Grant Complete",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        )
    }
})

# ═════════════════════════════════════════════════════════════════════════════
#  TAB 2: Manage by Site
# ═════════════════════════════════════════════════════════════════════════════

# ── Load Permissions ────────────────────────────────────────────────────────
$btnLoadSitePerms.Add_Click({
    if (-not (Test-IsConnected)) { return }

    $siteUrl = $txtSiteUrl.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($siteUrl)) {
        [System.Windows.MessageBox]::Show("Please enter a Site URL.", "Validation")
        return
    }

    Write-Log "Loading permissions for $siteUrl ..."
    $dgSitePerms.ItemsSource = $null

    try {
        $perms = @(Get-PnPEntraIDAppSitePermission -Site $siteUrl)
    }
    catch {
        Write-Log "Error loading permissions: $($_.Exception.Message)" "ERROR"
        [System.Windows.MessageBox]::Show(
            "Failed to load permissions.`n`n$($_.Exception.Message)",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        return
    }

    if ($perms.Count -eq 0) {
        Write-Log "No app permissions found on $siteUrl."
        [System.Windows.MessageBox]::Show(
            "No app permissions found on this site.",
            "No Results"
        )
        return
    }

    # Build display objects with roles and identity resolved from detail calls
    $items = [System.Collections.ObjectModel.ObservableCollection[PSObject]]::new()
    foreach ($p in $perms) {
        $detail = Get-PermissionDetail -PermissionId $p.Id -SiteUrl $siteUrl
        $items.Add([PSCustomObject]@{
            AppId       = $detail.AppId
            Roles       = $detail.Roles
            Id          = $p.Id
        })
    }
    $dgSitePerms.ItemsSource = $items
    Write-Log "Found $($items.Count) app permission(s) on $siteUrl."
})

# ── Edit Permission ─────────────────────────────────────────────────────────
$btnEditPerm.Add_Click({
    if (-not (Test-IsConnected)) { return }

    $selected = $dgSitePerms.SelectedItem
    if (-not $selected) {
        [System.Windows.MessageBox]::Show("Select a permission row first.", "No Selection")
        return
    }
    if (-not $cmbEditPermission.SelectedItem) {
        [System.Windows.MessageBox]::Show("Select the new permission level.", "Validation")
        return
    }

    $siteUrl       = $txtSiteUrl.Text.Trim()
    $permId        = $selected.Id
    $newPermission = $cmbEditPermission.SelectedItem.Content

    $confirm = [System.Windows.MessageBox]::Show(
        "Update permission for app '$($selected.AppId)' to '$newPermission' on:`n$siteUrl ?",
        "Confirm Update",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )
    if ($confirm -ne 'Yes') { return }

    Write-Log "Updating permission $permId to '$newPermission' on $siteUrl ..."
    try {
        $params = @{
            PermissionId = $permId
            Permissions  = $newPermission
        }
        if ($siteUrl) { $params['Site'] = $siteUrl }
        Set-PnPEntraIDAppSitePermission @params
        Write-Log "  ✓ Permission updated successfully."

        # Refresh the grid
        $btnLoadSitePerms.RaiseEvent(
            [System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent)
        )
    }
    catch {
        Write-Log "  ✗ Update failed: $($_.Exception.Message)" "ERROR"
        [System.Windows.MessageBox]::Show(
            "Failed to update permission.`n`n$($_.Exception.Message)",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
})

# ── Revoke Permission ──────────────────────────────────────────────────────
$btnRevokePerm.Add_Click({
    if (-not (Test-IsConnected)) { return }

    $selected = $dgSitePerms.SelectedItem
    if (-not $selected) {
        [System.Windows.MessageBox]::Show("Select a permission row first.", "No Selection")
        return
    }

    $siteUrl = $txtSiteUrl.Text.Trim()
    $permId  = $selected.Id

    $confirm = [System.Windows.MessageBox]::Show(
        "REVOKE permission for app '$($selected.AppId)' on:`n$siteUrl ?`n`nThis cannot be undone.",
        "Confirm Revoke",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Warning
    )
    if ($confirm -ne 'Yes') { return }

    Write-Log "Revoking permission $permId on $siteUrl ..."
    try {
        $params = @{ PermissionId = $permId; Force = $true }
        if ($siteUrl) { $params['Site'] = $siteUrl }
        Revoke-PnPEntraIDAppSitePermission @params
        Write-Log "  ✓ Permission revoked successfully."

        # Refresh the grid
        $btnLoadSitePerms.RaiseEvent(
            [System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent)
        )
    }
    catch {
        Write-Log "  ✗ Revoke failed: $($_.Exception.Message)" "ERROR"
        [System.Windows.MessageBox]::Show(
            "Failed to revoke permission.`n`n$($_.Exception.Message)",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
})

# ═════════════════════════════════════════════════════════════════════════════
#  TAB 3: Search by App
# ═════════════════════════════════════════════════════════════════════════════
$btnSearchApp.Add_Click({
    if (-not (Test-IsConnected)) { return }

    $appSearch = $txtSearchApp.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($appSearch)) {
        [System.Windows.MessageBox]::Show("Enter an App ID or Display Name.", "Validation")
        return
    }

    $dgAppSites.ItemsSource = $null
    $txtSearchProgress.Text = "Loading site list..."
    $btnSearchApp.IsEnabled = $false

    # Allow UI to update
    $window.Dispatcher.Invoke(
        [Action]{},
        [System.Windows.Threading.DispatcherPriority]::Background
    )

    Write-Log "Retrieving all tenant sites..."

    try {
        $allSites = @(Get-PnPTenantSite -ErrorAction Stop)
    }
    catch {
        Write-Log "Failed to get tenant sites: $($_.Exception.Message)" "ERROR"
        [System.Windows.MessageBox]::Show(
            "Failed to retrieve tenant sites.`nEnsure you are connected to the Admin site.`n`n$($_.Exception.Message)",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        $btnSearchApp.IsEnabled = $true
        $txtSearchProgress.Text = ""
        return
    }

    $total   = $allSites.Count
    $current = 0
    $results = [System.Collections.ObjectModel.ObservableCollection[PSObject]]::new()
    $dgAppSites.ItemsSource = $results

    Write-Log "Searching $total site(s) for app '$appSearch'..."

    foreach ($site in $allSites) {
        $current++
        $txtSearchProgress.Text = "Searching $current / $total ..."

        # Process UI events to keep responsive
        $window.Dispatcher.Invoke(
            [Action]{},
            [System.Windows.Threading.DispatcherPriority]::Background
        )

        try {
            $perms = @(Get-PnPEntraIDAppSitePermission `
                -AppIdentity $appSearch `
                -Site $site.Url `
                -ErrorAction Stop)

            foreach ($p in $perms) {
                $detail = Get-PermissionDetail -PermissionId $p.Id -SiteUrl $site.Url
                $results.Add([PSCustomObject]@{
                    SiteUrl      = $site.Url
                    SiteTitle    = $site.Title
                    AppId        = $detail.AppId
                    Roles        = $detail.Roles
                    PermissionId = $p.Id
                })
                Write-Log "  Found: $($site.Url) [$($detail.Roles)]"
            }
        }
        catch {
            # No permission found on this site or error — skip silently
        }
    }

    $txtSearchProgress.Text = "Done. Found $($results.Count) site(s)."
    $btnSearchApp.IsEnabled = $true
    Write-Log "Search complete. App '$appSearch' has permissions on $($results.Count) site(s)."

    if ($results.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            "No sites found with permissions for '$appSearch'.",
            "Search Complete"
        )
    }
})

# ═════════════════════════════════════════════════════════════════════════════
#  Show Window
# ═════════════════════════════════════════════════════════════════════════════
Write-Log "Sites.Selected Permissions Manager ready."
$window.ShowDialog() | Out-Null
