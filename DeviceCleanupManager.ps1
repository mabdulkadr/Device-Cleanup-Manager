<#!
.SYNOPSIS
    Device Cleanup Manager (WPF) - Active Directory Computer Cleanup and Control.

.DESCRIPTION
    A PowerShell 5.1 WPF tool to manage Active Directory computer objects:
      - Find inactive devices by LastLogonDate (lastLogonTimestamp-based) using a Days Inactive threshold
      - Find disabled devices (Enabled = False)
      - Simplified cleanup mode: choose either Inactive Devices or Disabled Computers
      - Load OU scope from LDAP using a searchable dropdown
      - Import device names from CSV (ComputerName/Name column) into Results for bulk actions
      - Safe actions with confirmation prompts + protected OU/DN rules
      - Export results to CSV and log actions to the Activity Log

    Notes about inactivity:
      - Get-ADComputer LastLogonDate is derived from lastLogonTimestamp replication and is suitable for cleanup reporting.
      - It is not the exact real-time "last logon" across all DCs.

.RUN AS
    Domain admin / delegated AD rights (recommended). A domain context is required for LDAP OU loading.
    RSAT ActiveDirectory module is required for AD actions (search/enable/disable/delete).

.EXAMPLE
    .\DeviceCleanupManager.ps1

.NOTES
    Author  : Mohammad Abdelkader
    Website : momar.tech
    Date    : 2026-02-25
    Version : 2.0
#>

#region ============================== CONFIG / CONSTANTS ===========================================
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$ToolName    = 'Device Cleanup Manager'
$AppVersion  = '2.0'
$AppDataRoot = Join-Path $env:ProgramData 'DeviceCleanupManager'
$LogsRoot    = Join-Path $AppDataRoot 'Logs'

if (-not (Test-Path $LogsRoot))     { New-Item -Path $LogsRoot -ItemType Directory -Force | Out-Null }

$script:ProtectedDNContains = @()

# Runtime state shared across UI handlers and background jobs
$script:ScanJob   = $null
$script:ScanTimer = $null
$script:OUAllItems = @()     # full OU list (objects)
$script:SelectedOUItem = $null
#endregion

#region ============================== WPF HELPERS / LOGGING =======================================
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase | Out-Null
Add-Type -AssemblyName System.DirectoryServices | Out-Null
Add-Type -AssemblyName Microsoft.VisualBasic | Out-Null

function New-Brush {
    param([string]$Hex)
    $bc = New-Object System.Windows.Media.BrushConverter
    return $bc.ConvertFromString($Hex)
}

function Add-LogLine {
    param(
        [Parameter(Mandatory=$true)][string]$Message,
        [ValidateSet('INFO','SUCCESS','WARNING','ERROR')][string]$Level = 'INFO'
    )

    if (-not $script:LogBox) { return }

    $ts = (Get-Date).ToString('HH:mm:ss')
    $prefix = "[{0}] [{1}] " -f $ts, $Level

    $color = '#E4E9F0'
    switch ($Level) {
        'INFO'    { $color = '#E4E9F0' }
        'SUCCESS' { $color = '#7CFFB2' }
        'WARNING' { $color = '#FFDF7C' }
        'ERROR'   { $color = '#FF8A8A' }
    }

    $para = New-Object System.Windows.Documents.Paragraph
    $para.Margin = '0,0,0,2'

    $run1 = New-Object System.Windows.Documents.Run($prefix)
    $run1.Foreground = New-Brush $color
    $run1.FontWeight = 'SemiBold'

    $run2 = New-Object System.Windows.Documents.Run($Message)
    $run2.Foreground = New-Brush '#E4E9F0'

    [void]$para.Inlines.Add($run1)
    [void]$para.Inlines.Add($run2)

    $doc = $script:LogBox.Document
    [void]$doc.Blocks.Add($para)
    $script:LogBox.ScrollToEnd()
}

function Show-WPFMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$Message,
        [string]$Title = 'Message',
        [ValidateSet('Green','Orange','Red','Blue')][string]$Color = 'Blue'
    )

    $header = '#0078D7'
    switch ($Color) {
        'Green'  { $header = '#28A745' }
        'Orange' { $header = '#FFA500' }
        'Red'    { $header = '#DC3545' }
        'Blue'   { $header = '#0078D7' }
    }

    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        WindowStartupLocation="CenterScreen"
        SizeToContent="Height"
        Width="560"
        Background="#FFFFFF"
        Title="$Title">
  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <Border Grid.Row="0" Background="$header" Padding="12">
      <TextBlock Text="$Title" Foreground="White" FontSize="16" FontWeight="Bold"/>
    </Border>

    <TextBlock Grid.Row="1" Margin="16" TextWrapping="Wrap" FontSize="13" Foreground="#111827" x:Name="Msg"/>

    <Button Grid.Row="2" Content="OK" Width="90" Height="30" Margin="0,8,0,14"
            HorizontalAlignment="Center" Background="$header" Foreground="White"
            BorderThickness="0" FontWeight="Bold" Cursor="Hand" ToolTip="Close this message" x:Name="OkBtn"/>
  </Grid>
</Window>
"@

    $xml = [xml]$xaml
    $reader = New-Object System.Xml.XmlNodeReader($xml)
    $w = [System.Windows.Markup.XamlReader]::Load($reader)

    $w.FindName('Msg').Text = $Message
    $w.FindName('OkBtn').Add_Click({ $w.Close() }) | Out-Null
    [void]$w.ShowDialog()
}

function Show-WPFConfirmation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$Message,
        [string]$Title = 'Confirmation',
        [ValidateSet('Green','Orange','Red','Blue')][string]$Color = 'Blue'
    )

    $header = '#0078D7'
    switch ($Color) {
        'Green'  { $header = '#28A745' }
        'Orange' { $header = '#FFA500' }
        'Red'    { $header = '#DC3545' }
        'Blue'   { $header = '#0078D7' }
    }

    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        WindowStartupLocation="CenterScreen"
        SizeToContent="WidthAndHeight"
        Background="#FFFFFF"
        Title="$Title">
  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <Border Grid.Row="0" Background="$header" Padding="12">
      <TextBlock Text="$Title" Foreground="White" FontSize="16" FontWeight="Bold" HorizontalAlignment="Center"/>
    </Border>

    <TextBlock Grid.Row="1" Margin="18" TextWrapping="Wrap" FontSize="13" Foreground="#111827" x:Name="Msg" Width="520"/>

    <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,6,0,14">
      <Button x:Name="YesBtn" Content="Yes" Width="92" Height="30" Margin="8,0"
              Background="$header" Foreground="White" BorderThickness="0" FontWeight="Bold" Cursor="Hand"
              ToolTip="Confirm and continue"/>
      <Button x:Name="NoBtn" Content="No" Width="92" Height="30" Margin="8,0"
              Background="#E5E7EB" Foreground="#111827" BorderThickness="0" FontWeight="Bold" Cursor="Hand"
              ToolTip="Cancel this action"/>
    </StackPanel>
  </Grid>
</Window>
"@

    $xml = [xml]$xaml
    $reader = New-Object System.Xml.XmlNodeReader($xml)
    $w = [System.Windows.Markup.XamlReader]::Load($reader)

    $w.FindName('Msg').Text = $Message

    $w.FindName('YesBtn').Add_Click({
        $w.DialogResult = $true
        $w.Close()
    }) | Out-Null

    $w.FindName('NoBtn').Add_Click({
        $w.DialogResult = $false
        $w.Close()
    }) | Out-Null

    [void]$w.ShowDialog()
    return $w.DialogResult
}

function Invoke-SafeUIAction {
    param(
        [Parameter(Mandatory=$true)][scriptblock]$Action,
        [Parameter(Mandatory=$true)][string]$ActionName
    )

    try {
        & $Action
    }
    catch {
        $err = $_.Exception.Message
        Add-LogLine -Message ("{0} failed: {1}" -f $ActionName, $err) -Level 'ERROR'
        Show-WPFMessage -Message ("{0} failed.`n{1}" -f $ActionName, $err) -Title 'Operation Error' -Color 'Red'
    }
}
#endregion

#region ============================== PROTECTION RULES ============================================
function Test-IsProtectedDN {
    param([string]$DistinguishedName)

    if ([string]::IsNullOrWhiteSpace($DistinguishedName)) { return $false }
    foreach ($p in $script:ProtectedDNContains) {
        if ($DistinguishedName -like "*$p*") { return $true }
    }
    return $false
}
#endregion

#region ============================== LDAP OU LOADER (FAST) =======================================
function Convert-DNToFriendlyPath {
    param([string]$DN)

    if ([string]::IsNullOrWhiteSpace($DN)) { return $DN }

    # Example:
    # OU=IT,OU=Faculty,DC=qassimu,DC=local  ->  Faculty / IT
    $parts = $DN -split ','
    $ous = @()
    foreach ($p in $parts) {
        if ($p.Trim().ToUpper().StartsWith('OU=')) {
            $ous += ($p.Trim().Substring(3))
        }
    }
    [array]::Reverse($ous)
    if ($ous.Count -eq 0) { return $DN }
    return ($ous -join ' / ')
}

function Convert-DNListToOUItems {
    param([string[]]$DNs)

    $results = @()
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($dn in $DNs) {
        if ([string]::IsNullOrWhiteSpace($dn)) { continue }
        if (-not $seen.Add($dn)) { continue }

        $friendly = Convert-DNToFriendlyPath -DN $dn
        $results += [PSCustomObject]@{
            Display = $friendly
            DN      = $dn
        }
    }

    $results = $results | Sort-Object -Property DN
    return @([PSCustomObject]@{ Display = 'Entire Domain'; DN = '' }) + $results
}

function Resolve-LDAPBaseContext {
    # Returns object: RootDN, Server, Source
    $candidates = @()

    # 1) Current domain context (works when token is domain-associated)
    try {
        $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
        if ($domain) {
            $dn = [string]$domain.GetDirectoryEntry().distinguishedName
            if (-not [string]::IsNullOrWhiteSpace($dn)) {
                $candidates += [PSCustomObject]@{
                    RootDN = $dn
                    Server = ''
                    Source = 'CurrentDomain'
                }
            }
        }
    } catch {}

    # 2) Local RootDSE
    try {
        $rootDse = [ADSI]'LDAP://RootDSE'
        $dn = [string]$rootDse.defaultNamingContext
        if (-not [string]::IsNullOrWhiteSpace($dn)) {
            $candidates += [PSCustomObject]@{
                RootDN = $dn
                Server = ''
                Source = 'RootDSE(Local)'
            }
        }
    } catch {}

    # 3) RootDSE via server hints from session environment
    $serverHints = @()
    if ($env:LOGONSERVER) { $serverHints += ($env:LOGONSERVER -replace '^\\\\', '') }
    if ($env:USERDNSDOMAIN) { $serverHints += $env:USERDNSDOMAIN }
    if ($env:USERDOMAIN -and ($env:USERDOMAIN -ne $env:COMPUTERNAME)) { $serverHints += $env:USERDOMAIN }
    $serverHints = $serverHints | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

    foreach ($srv in $serverHints) {
        try {
            $rootDse = [ADSI]("LDAP://{0}/RootDSE" -f $srv)
            $dn = [string]$rootDse.defaultNamingContext
            if (-not [string]::IsNullOrWhiteSpace($dn)) {
                $candidates += [PSCustomObject]@{
                    RootDN = $dn
                    Server = $srv
                    Source = ("RootDSE({0})" -f $srv)
                }
            }
        } catch {}
    }

    foreach ($c in $candidates) {
        if (-not [string]::IsNullOrWhiteSpace($c.RootDN)) { return $c }
    }
    return $null
}

function Get-OUsViaLDAP {
    # Returns OU items with Display text + DN value for ComboBox binding
    try {
        # Use the same LDAP approach that works interactively on this environment.
        $searchBaseDN = [string]([ADSI]'LDAP://RootDSE').defaultNamingContext
        if ([string]::IsNullOrWhiteSpace($searchBaseDN)) {
            throw 'Unable to read defaultNamingContext from LDAP://RootDSE.'
        }

        Add-LogLine -Message ("OU LDAP search base: {0}" -f $searchBaseDN) -Level 'INFO'

        $searchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$searchBaseDN")
        $ds = New-Object System.DirectoryServices.DirectorySearcher
        $ds.SearchRoot = $searchRoot
        $ds.Filter = '(objectClass=organizationalUnit)'
        $ds.PageSize = 1000
        [void]$ds.PropertiesToLoad.Add('distinguishedName')

        $dns = @()
        foreach ($r in $ds.FindAll()) {
            if ($r.Properties['distinguishedname'] -and $r.Properties['distinguishedname'].Count -gt 0) {
                $dn = [string]$r.Properties['distinguishedname'][0]
                if (-not [string]::IsNullOrWhiteSpace($dn)) { $dns += $dn }
            }
        }

        return Convert-DNListToOUItems -DNs $dns
    }
    catch {
        throw ("LDAP OU query failed: {0}" -f $_.Exception.Message)
    }
}

function Set-SelectedOUScope {
    param([object]$Item)

    if (-not $Item) { return }
    $script:SelectedOUItem = $Item

    if ($script:OUScopeText) {
        $script:OUScopeText.Text = ($Item.Display + '')
    }
}

function Show-OUSelectorDialog {
    param(
        [Parameter(Mandatory=$true)][object[]]$Items,
        [string]$CurrentDN = ''
    )

    if (-not $Items -or $Items.Count -eq 0) { return $null }

    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Select OU Scope"
        Width="820" Height="620"
        MinWidth="700" MinHeight="500"
        WindowStartupLocation="CenterOwner"
        ResizeMode="CanResizeWithGrip"
        Background="#F6F8FB">
  <Grid Margin="14">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <Border Grid.Row="0" Background="#1F2D3A" Padding="12" CornerRadius="6">
      <TextBlock Text="OU Scope Selector" Foreground="White" FontSize="16" FontWeight="SemiBold"/>
    </Border>

    <Grid Grid.Row="1" Margin="0,10,0,8">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="Auto"/>
      </Grid.ColumnDefinitions>
      <TextBlock Grid.Column="0" Text="Search OU:" VerticalAlignment="Center" Foreground="#1F2D3A" FontWeight="SemiBold"/>
      <TextBox Grid.Column="1" x:Name="FilterBox" Height="30" Margin="8,0,10,0"
               Background="White" BorderBrush="#D0D7E2" BorderThickness="1" Padding="10,4"
               ToolTip="Type OU name, path, or DN"/>
      <TextBlock Grid.Column="2" x:Name="CountTxt" VerticalAlignment="Center" Foreground="#5F6B7A"/>
    </Grid>

    <Border Grid.Row="2" Background="White" BorderBrush="#D0D7E2" BorderThickness="1" CornerRadius="6">
      <ListBox x:Name="OUList"
               DisplayMemberPath="Display"
               ScrollViewer.VerticalScrollBarVisibility="Auto"
               VirtualizingStackPanel.IsVirtualizing="True"
               VirtualizingStackPanel.VirtualizationMode="Recycling"
               BorderThickness="0"
               FontSize="13"
               Padding="6"/>
    </Border>

    <Grid Grid.Row="3" Margin="0,10,0,0">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="Auto"/>
      </Grid.ColumnDefinitions>
      <TextBlock Grid.Column="0" Text="Tip: Double-click an OU to select it quickly." Foreground="#5F6B7A" VerticalAlignment="Center"/>
      <StackPanel Grid.Column="1" Orientation="Horizontal">
        <Button x:Name="EntireBtn" Content="Entire Domain" Width="120" Height="30" Margin="0,0,8,0"
                Background="#E5E7EB" Foreground="#1F2D3A" BorderThickness="0" FontWeight="SemiBold"
                ToolTip="Set scope to entire domain"/>
        <Button x:Name="SelectBtn" Content="Select" Width="90" Height="30" Margin="0,0,8,0"
                Background="#9FAEF7" Foreground="#1F2D3A" BorderThickness="0" FontWeight="SemiBold"
                ToolTip="Use selected OU as search scope"/>
        <Button x:Name="CancelBtn" Content="Cancel" Width="90" Height="30"
                Background="#E5E7EB" Foreground="#1F2D3A" BorderThickness="0" FontWeight="SemiBold"
                ToolTip="Close without changing OU scope"/>
      </StackPanel>
    </Grid>
  </Grid>
</Window>
"@

    $xml = [xml]$xaml
    $reader = New-Object System.Xml.XmlNodeReader($xml)
    $w = [System.Windows.Markup.XamlReader]::Load($reader)
    if ($script:Window) { $w.Owner = $script:Window }

    $filterBox = $w.FindName('FilterBox')
    $listBox   = $w.FindName('OUList')
    $countTxt  = $w.FindName('CountTxt')
    $entireBtn = $w.FindName('EntireBtn')
    $selectBtn = $w.FindName('SelectBtn')
    $cancelBtn = $w.FindName('CancelBtn')

    $listBox.ItemsSource = $Items
    $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($listBox.ItemsSource)

    $view.Filter = {
        param($item)
        $q = ($filterBox.Text + '').Trim()
        if ([string]::IsNullOrWhiteSpace($q)) { return $true }
        $needle = $q.ToLowerInvariant()
        return (
            (($item.Display + '').ToLowerInvariant().Contains($needle)) -or
            (($item.DN + '').ToLowerInvariant().Contains($needle))
        )
    }

    $updateCount = {
        $count = 0
        foreach ($x in $view) { $count++ }
        $countTxt.Text = "Matches: $count / $($Items.Count)"
    }

    $filterBox.Add_TextChanged({
        $view.Refresh()
        & $updateCount
    }) | Out-Null

    $selected = $null

    if (-not [string]::IsNullOrWhiteSpace($CurrentDN)) {
        $currentItem = $Items | Where-Object { ($_.DN + '') -eq $CurrentDN } | Select-Object -First 1
        if ($currentItem) { $listBox.SelectedItem = $currentItem }
    }
    if (-not $listBox.SelectedItem -and $Items.Count -gt 0) { $listBox.SelectedIndex = 0 }
    if ($listBox.SelectedItem) { $listBox.ScrollIntoView($listBox.SelectedItem) }
    & $updateCount

    $doSelect = {
        if ($listBox.SelectedItem) {
            $script:__PickedOU = $listBox.SelectedItem
            $w.DialogResult = $true
            $w.Close()
        }
    }

    $selectBtn.Add_Click($doSelect) | Out-Null
    $listBox.Add_MouseDoubleClick($doSelect) | Out-Null

    $entireBtn.Add_Click({
        $entire = $Items | Where-Object { [string]::IsNullOrWhiteSpace($_.DN + '') } | Select-Object -First 1
        if ($entire) {
            $script:__PickedOU = $entire
            $w.DialogResult = $true
            $w.Close()
        }
    }) | Out-Null

    $cancelBtn.Add_Click({
        $w.DialogResult = $false
        $w.Close()
    }) | Out-Null

    $script:__PickedOU = $null
    [void]$w.ShowDialog()
    $selected = $script:__PickedOU
    $script:__PickedOU = $null
    return $selected
}

function Initialize-OUComboBox {
    try {
        Add-LogLine -Message 'Loading OUs via LDAP...' -Level 'INFO'
        $script:OUAllItems = @(Get-OUsViaLDAP)
        $ouCount = @($script:OUAllItems).Count

        $default = $script:OUAllItems | Select-Object -First 1
        Set-SelectedOUScope -Item $default
        Add-LogLine -Message ("OUs loaded: {0}" -f ([Math]::Max(0, ($ouCount - 1)))) -Level 'SUCCESS'
    }
    catch {
        Add-LogLine -Message ("OU load failed: {0}" -f $_.Exception.Message) -Level 'ERROR'
        $script:OUAllItems = @([PSCustomObject]@{ Display = 'Entire Domain'; DN = '' })
        Set-SelectedOUScope -Item $script:OUAllItems[0]
        Show-WPFMessage -Message ("Unable to load OU list from LDAP.`n{0}`n`nUse a domain account/session, or run from a domain-joined machine." -f $_.Exception.Message) -Title 'OU Load Error' -Color 'Red'
    }
}
#endregion

#region ============================== AD UTILITIES ================================================
function Ensure-ADModule {
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Show-WPFMessage -Message 'ActiveDirectory module (RSAT) is required for AD actions. Install RSAT or run on a management server/DC.' -Title 'Missing RSAT' -Color 'Red'
        return $false
    }
    Import-Module ActiveDirectory -ErrorAction Stop
    return $true
}

function Get-SearchBaseDNFromUI {
    $sel = $script:SelectedOUItem
    if (-not $sel) { return '' }
    return ($sel.DN + '')
}

function Parse-DaysInactive {
    $n = 0
    $ok = [int]::TryParse(($script:DaysInactiveBox.Text + ''), [ref]$n)
    if (-not $ok -or $n -le 0) { return $null }
    return $n
}

function Get-ModeFromUI {
    if ($script:ModeDisabledOnly.IsChecked) { return 'DisabledOnly' }
    return 'InactiveOnly'
}

function Update-ModeUIState {
    $mode = Get-ModeFromUI
    $isInactiveMode = ($mode -eq 'InactiveOnly')

    if ($script:DaysInactiveBox) {
        $script:DaysInactiveBox.IsEnabled = $isInactiveMode
    }

    if ($script:IncludeNeverLogon) {
        $script:IncludeNeverLogon.IsEnabled = $isInactiveMode
        if (-not $isInactiveMode) {
            $script:IncludeNeverLogon.IsChecked = $false
        }
    }
}
#endregion

#region ============================== SEARCH ENGINE (JOB + TIMER) =================================
function Start-Search {
    if (-not (Ensure-ADModule)) { return }

    $mode = Get-ModeFromUI
    $days = 0
    $includeNever = $false

    if ($mode -eq 'InactiveOnly') {
        $days = Parse-DaysInactive
        if ($days -eq $null) {
            $script:StatusLabel.Text = "Invalid Days Inactive value."
            Add-LogLine -Message 'Invalid Days Inactive input (must be greater than 0 for Inactive mode).' -Level 'WARNING'
            return
        }
        $includeNever = [bool]$script:IncludeNeverLogon.IsChecked
    }

    $searchBase = Get-SearchBaseDNFromUI
    $cutoff = if ($mode -eq 'InactiveOnly') { (Get-Date).AddDays(-$days) } else { Get-Date }

    # UI state
    $script:ProgressBar.Visibility = 'Visible'
    $script:ProgressBar.IsIndeterminate = $true
    $script:StatusLabel.Text = 'Searching...'
    $daysText = if ($mode -eq 'InactiveOnly') { $days } else { 'N/A' }
    $scopeText = if ([string]::IsNullOrWhiteSpace($searchBase)) { 'Entire Domain' } else { $searchBase }
    Add-LogLine -Message ("Search started. Mode={0}, Days={1}, IncludeNeverLogon={2}, Scope={3}" -f $mode, $daysText, $includeNever, $scopeText) -Level 'INFO'

    # Cleanup previous job/timer
    try {
        if ($script:ScanTimer) { $script:ScanTimer.Stop() }
        if ($script:ScanJob) {
            if ($script:ScanJob.State -eq 'Running') { Stop-Job -Job $script:ScanJob -Force | Out-Null }
            Remove-Job -Job $script:ScanJob -Force | Out-Null
        }
    } catch {}

    $script:ScanJob = Start-Job -ArgumentList $days, $cutoff, $searchBase, $mode, $includeNever -ScriptBlock {
        param($Days, $Cutoff, $SearchBaseDN, $Mode, $IncludeNeverLogon)

        Import-Module ActiveDirectory

        if ($Mode -eq 'DisabledOnly') {
            # Use server-side disabled filter for accuracy.
            $disabledSet = @()
            if ([string]::IsNullOrWhiteSpace($SearchBaseDN)) {
                $disabledSet = Get-ADComputer -Filter 'Enabled -eq $false' -Properties Name,Enabled,LastLogonDate,DistinguishedName
            } else {
                $disabledSet = Get-ADComputer -Filter 'Enabled -eq $false' -SearchBase $SearchBaseDN -Properties Name,Enabled,LastLogonDate,DistinguishedName
            }

            $finalDisabled = New-Object 'System.Collections.Generic.List[object]'
            foreach ($c in $disabledSet) {
                $ll = $c.LastLogonDate
                $hasLogon = [bool]$ll
                $inactiveDays = if ($hasLogon) { (New-TimeSpan -Start $ll -End (Get-Date)).Days } else { 'N/A' }

                $finalDisabled.Add([PSCustomObject]@{
                    ComputerName      = $c.Name
                    Enabled           = $false
                    LastLogonDate     = $(if ($hasLogon) { $ll } else { 'No Logon Information' })
                    InactiveDays      = $inactiveDays
                    DistinguishedName = $c.DistinguishedName
                    Source            = 'Search'
                })
            }
            return $finalDisabled
        }

        $all = @()
        if ([string]::IsNullOrWhiteSpace($SearchBaseDN)) {
            $all = Get-ADComputer -Filter * -Properties Name,Enabled,LastLogonDate,DistinguishedName
        } else {
            $all = Get-ADComputer -Filter * -SearchBase $SearchBaseDN -Properties Name,Enabled,LastLogonDate,DistinguishedName
        }

        $inactive = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($c in $all) {
            $ll = $c.LastLogonDate
            $hasLogon = [bool]$ll

            $isInactive = $false
            $inactiveDays = $null

            if ($hasLogon) {
                $inactiveDays = (New-TimeSpan -Start $ll -End (Get-Date)).Days
                if ([datetime]$ll -lt $Cutoff) { $isInactive = $true }
            } elseif ($IncludeNeverLogon) {
                $isInactive = $true
            }

            if ($isInactive) {
                $inactive[$c.DistinguishedName] = [PSCustomObject]@{
                    ComputerName      = $c.Name
                    Enabled           = [bool]$c.Enabled
                    LastLogonDate     = $(if ($hasLogon) { $ll } else { 'No Logon Information' })
                    InactiveDays      = $(if ($hasLogon) { $inactiveDays } else { 'N/A' })
                    DistinguishedName = $c.DistinguishedName
                    Source            = 'Search'
                }
            }
        }

        $finalInactive = New-Object 'System.Collections.Generic.List[object]'
        foreach ($k in $inactive.Keys) { $finalInactive.Add($inactive[$k]) }
        return $finalInactive
    }

    $script:ScanTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:ScanTimer.Interval = [TimeSpan]::FromMilliseconds(450)
    $script:ScanTimer.Add_Tick({
        if (-not $script:ScanJob) { return }

        if ($script:ScanJob.State -eq 'Completed') {
            $script:ScanTimer.Stop()
            $data = @()

            try { $data = @(Receive-Job -Job $script:ScanJob -Keep) } catch {}

            if (@($data).Count -eq 0) { $data = @() }

            # Add friendly OU Path column
            $out = @()
            foreach ($r in $data) {
                $dn = $r.DistinguishedName
                $ouPath = Convert-DNToFriendlyPath -DN $dn
                $out += [PSCustomObject]@{
                    IsSelected        = $false
                    ComputerName      = $r.ComputerName
                    Enabled           = $r.Enabled
                    LastLogonDate     = $r.LastLogonDate
                    InactiveDays      = $r.InactiveDays
                    OUPath            = $ouPath
                    DistinguishedName = $r.DistinguishedName
                    Source            = $r.Source
                }
            }

            $script:ComputerGrid.ItemsSource = @($out)
            if ($script:SelectAllGridCheck) { $script:SelectAllGridCheck.IsChecked = $false }
            $script:StatusLabel.Text = ("Done. Results: {0}" -f (@($out).Count))
            Add-LogLine -Message ("Search completed. Results: {0}" -f (@($out).Count)) -Level 'SUCCESS'

            $script:ProgressBar.IsIndeterminate = $false
            $script:ProgressBar.Visibility = 'Hidden'
        }

        if ($script:ScanJob.State -eq 'Failed') {
            $script:ScanTimer.Stop()
            $script:StatusLabel.Text = 'Search failed.'
            Add-LogLine -Message 'Search job failed.' -Level 'ERROR'
            $script:ProgressBar.IsIndeterminate = $false
            $script:ProgressBar.Visibility = 'Hidden'
        }
    }) | Out-Null

    $script:ScanTimer.Start()
}
#endregion

#region ============================== CSV IMPORT + DIRECT ACTIONS =================================
function Import-CsvToGrid {
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = 'CSV files (*.csv)|*.csv'
    $dlg.Multiselect = $false
    if ($dlg.ShowDialog() -ne $true) { return }

    $path = $dlg.FileName
    Add-LogLine -Message ("Loading CSV: {0}" -f $path) -Level 'INFO'

    $rows = @()
    try { $rows = @(Import-Csv -Path $path) } catch {}

    if (@($rows).Count -eq 0) {
        Show-WPFMessage -Message 'CSV is empty.' -Title 'CSV Import' -Color 'Orange'
        return
    }

    $list = @()
    foreach ($r in $rows) {
        $name = $null

        if ($r.PSObject.Properties.Name -contains 'ComputerName') { $name = $r.ComputerName }
        elseif ($r.PSObject.Properties.Name -contains 'Name')     { $name = $r.Name }

        $name = ($name + '').Trim()
        if ([string]::IsNullOrWhiteSpace($name)) { continue }

        $list += [PSCustomObject]@{
            IsSelected        = $false
            ComputerName      = $name
            Enabled           = ''
            LastLogonDate     = ''
            InactiveDays      = ''
            OUPath            = ''
            DistinguishedName = ''
            Source            = 'CSV'
        }
    }

    $script:ComputerGrid.ItemsSource = @($list)
    if ($script:SelectAllGridCheck) { $script:SelectAllGridCheck.IsChecked = $false }
    $script:StatusLabel.Text = ("CSV loaded. Items: {0}" -f (@($list).Count))
    Add-LogLine -Message ("CSV loaded. Items: {0}" -f (@($list).Count)) -Level 'SUCCESS'
}

function Resolve-ADComputerByName {
    param([string]$Name, [string]$SearchBaseDN)

    if ([string]::IsNullOrWhiteSpace($SearchBaseDN)) {
        return Get-ADComputer -LDAPFilter "(name=$Name)" -Properties DistinguishedName,Enabled,LastLogonDate
    } else {
        return Get-ADComputer -LDAPFilter "(name=$Name)" -SearchBase $SearchBaseDN -Properties DistinguishedName,Enabled,LastLogonDate
    }
}

function Refresh-GridAfterAction {
    param(
        [ValidateSet('Disable','Enable','Delete')][string]$Action,
        [string[]]$ComputerNames
    )

    $names = @($ComputerNames | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
    if (@($names).Count -eq 0) { return }

    $rows = @($script:ComputerGrid.ItemsSource)
    if (@($rows).Count -eq 0) { return }

    $nameSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($n in $names) { [void]$nameSet.Add(($n + '').Trim()) }

    if ($Action -eq 'Delete') {
        $newRows = @()
        foreach ($row in $rows) {
            $rowName = (($row.ComputerName + '').Trim())
            if ($nameSet.Contains($rowName)) { continue }
            $row.IsSelected = $false
            $newRows += $row
        }

        $script:ComputerGrid.ItemsSource = @($newRows)
        if ($script:SelectAllGridCheck) { $script:SelectAllGridCheck.IsChecked = $false }
        return
    }

    $searchBase = Get-SearchBaseDNFromUI
    foreach ($row in $rows) {
        $rowName = (($row.ComputerName + '').Trim())
        if (-not $nameSet.Contains($rowName)) { continue }

        try {
            $ad = Resolve-ADComputerByName -Name $rowName -SearchBaseDN $searchBase
            if (-not $ad) { continue }

            $ll = $ad.LastLogonDate
            $hasLogon = [bool]$ll

            $row.Enabled           = [bool]$ad.Enabled
            $row.LastLogonDate     = $(if ($hasLogon) { $ll } else { 'No Logon Information' })
            $row.InactiveDays      = $(if ($hasLogon) { (New-TimeSpan -Start $ll -End (Get-Date)).Days } else { 'N/A' })
            $row.DistinguishedName = $ad.DistinguishedName
            $row.OUPath            = Convert-DNToFriendlyPath -DN $ad.DistinguishedName
            $row.IsSelected        = $false
        }
        catch {
            Add-LogLine -Message ("Refresh skipped for {0}: {1}" -f $rowName, $_.Exception.Message) -Level 'WARNING'
        }
    }

    try { $script:ComputerGrid.Items.Refresh() } catch {}
    if ($script:SelectAllGridCheck) { $script:SelectAllGridCheck.IsChecked = $false }
}

function Invoke-DirectActionOnGrid {
    param(
        [ValidateSet('Disable','Enable','Delete')][string]$Action
    )

    if (-not (Ensure-ADModule)) { return }

    $data = @($script:ComputerGrid.ItemsSource)
    if (@($data).Count -eq 0) {
        $script:StatusLabel.Text = 'Grid is empty.'
        return
    }

    $color = 'Blue'
    if ($Action -eq 'Delete') { $color = 'Red' }
    $ok = Show-WPFConfirmation -Message ("Confirm {0} for ALL devices currently in the grid?" -f $Action) -Title ("Confirm {0}" -f $Action) -Color $color
    if ($ok -ne $true) { return }

    $searchBase = Get-SearchBaseDNFromUI
    $success = 0
    $successNames = @()

    foreach ($row in $data) {
        $name = ($row.ComputerName + '').Trim()
        if ([string]::IsNullOrWhiteSpace($name)) { continue }

        try {
            $ad = Resolve-ADComputerByName -Name $name -SearchBaseDN $searchBase
            if (-not $ad) {
                Add-LogLine -Message ("Not found in AD: {0}" -f $name) -Level 'WARNING'
                continue
            }

            if (Test-IsProtectedDN -DistinguishedName $ad.DistinguishedName) {
                Add-LogLine -Message ("Skipped (Protected DN): {0}" -f $name) -Level 'WARNING'
                continue
            }

            switch ($Action) {
                'Disable' {
                    Disable-ADAccount -Identity $ad.DistinguishedName -ErrorAction Stop
                    Add-LogLine -Message ("Disabled: {0}" -f $name) -Level 'SUCCESS'
                }
                'Enable' {
                    Enable-ADAccount -Identity $ad.DistinguishedName -ErrorAction Stop
                    Add-LogLine -Message ("Enabled: {0}" -f $name) -Level 'SUCCESS'
                }
                'Delete' {
                    Remove-ADComputer -Identity $ad.DistinguishedName -Confirm:$false -ErrorAction Stop
                    Add-LogLine -Message ("Deleted: {0}" -f $name) -Level 'SUCCESS'
                }
            }

            $success++
            $successNames += $name
        }
        catch {
            Add-LogLine -Message ("{0} failed for {1}: {2}" -f $Action, $name, $_.Exception.Message) -Level 'ERROR'
        }
    }

    if ($success -gt 0) {
        Refresh-GridAfterAction -Action $Action -ComputerNames $successNames
    }

    $script:StatusLabel.Text = ("{0} done. Success: {1}" -f $Action, $success)
}
#endregion

#region ============================== GRID ACTIONS (SELECTED) =====================================
function Set-GridCheckedState {
    param([bool]$Checked)

    $rows = @($script:ComputerGrid.ItemsSource)
    if (@($rows).Count -eq 0) { return }

    foreach ($row in $rows) {
        try {
            if ($row.PSObject.Properties.Name -contains 'IsSelected') {
                $row.IsSelected = $Checked
            }
        } catch {}
    }

    try { $script:ComputerGrid.Items.Refresh() } catch {}
}

function Get-SelectedGridItems {
    try {
        [void]$script:ComputerGrid.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Cell, $true)
        [void]$script:ComputerGrid.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row,  $true)
    } catch {}

    $checked = @()
    $rows = @($script:ComputerGrid.ItemsSource)
    foreach ($row in $rows) {
        try {
            if ($row.PSObject.Properties.Name -contains 'IsSelected' -and [bool]$row.IsSelected) {
                $checked += $row
            }
        } catch {}
    }
    return $checked
}

function Invoke-ActionOnSelected {
    param([ValidateSet('Disable','Enable','Delete')][string]$Action)

    if (-not (Ensure-ADModule)) { return }

    $selected = Get-SelectedGridItems
    if (@($selected).Count -eq 0) {
        $script:StatusLabel.Text = 'No checked devices.'
        return
    }

    $color = 'Blue'
    if ($Action -eq 'Delete') { $color = 'Red' }

    $ok = Show-WPFConfirmation -Message ("Confirm {0} for selected devices ({1})?" -f $Action, (@($selected).Count)) -Title ("Confirm {0}" -f $Action) -Color $color
    if ($ok -ne $true) { return }

    $success = 0
    $successNames = @()

    foreach ($item in $selected) {
        $name = ($item.ComputerName + '').Trim()
        $dn   = ($item.DistinguishedName + '').Trim()

        try {
            if (-not $dn) {
                # If DN missing (CSV), resolve
                $searchBase = Get-SearchBaseDNFromUI
                $ad = Resolve-ADComputerByName -Name $name -SearchBaseDN $searchBase
                if (-not $ad) {
                    Add-LogLine -Message ("Not found in AD: {0}" -f $name) -Level 'WARNING'
                    continue
                }
                $dn = $ad.DistinguishedName
            }

            if (Test-IsProtectedDN -DistinguishedName $dn) {
                Add-LogLine -Message ("Skipped (Protected DN): {0}" -f $name) -Level 'WARNING'
                continue
            }

            switch ($Action) {
                'Disable' { Disable-ADAccount -Identity $dn -ErrorAction Stop }
                'Enable'  { Enable-ADAccount  -Identity $dn -ErrorAction Stop }
                'Delete'  { Remove-ADComputer -Identity $dn -Confirm:$false -ErrorAction Stop }
            }

            Add-LogLine -Message ("{0}: {1}" -f $Action, $name) -Level 'SUCCESS'
            $success++
            $successNames += $name
        }
        catch {
            Add-LogLine -Message ("{0} failed for {1}: {2}" -f $Action, $name, $_.Exception.Message) -Level 'ERROR'
        }
    }

    if ($success -gt 0) {
        Refresh-GridAfterAction -Action $Action -ComputerNames $successNames
    }

    $script:StatusLabel.Text = ("{0} done. Success: {1}" -f $Action, $success)
}
#endregion

#region ============================== EXPORT / UTILS ==============================================
function Export-GridToCsv {
    $data = @($script:ComputerGrid.ItemsSource)
    if (@($data).Count -eq 0) {
        $script:StatusLabel.Text = 'No data to export.'
        return
    }

    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.FileName   = 'AD_Devices_Report'
    $dlg.DefaultExt = '.csv'
    $dlg.Filter     = 'CSV files (*.csv)|*.csv'
    if ($dlg.ShowDialog() -ne $true) { return }

    $path = $dlg.FileName
    $data | Select-Object ComputerName,Enabled,LastLogonDate,InactiveDays,OUPath,DistinguishedName,Source |
        Export-Csv -Path $path -NoTypeInformation -Encoding UTF8

    $script:StatusLabel.Text = ("Exported: {0}" -f $path)
    Add-LogLine -Message ("Exported CSV: {0}" -f $path) -Level 'SUCCESS'
}

function Clear-Results {
    $script:ComputerGrid.ItemsSource = @()
    if ($script:SelectAllGridCheck) { $script:SelectAllGridCheck.IsChecked = $false }
    $script:StatusLabel.Text = 'Cleared.'
    Add-LogLine -Message 'Results cleared.' -Level 'INFO'
}
#endregion

#region ============================== XAML UI ======================================================
$xamlText = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Device Cleanup Manager"
        Width="1200" Height="820"
        WindowStartupLocation="CenterScreen"
        Background="#F6F8FB"
        FontFamily="Segoe UI"
        FontSize="13"
        UseLayoutRounding="True"
        SnapsToDevicePixels="True">

  <Window.Resources>
    <DropShadowEffect x:Key="ShadowPrimary" BlurRadius="10" ShadowDepth="0" Opacity="0.55" Color="#9FAEF7"/>
    <DropShadowEffect x:Key="ShadowBlue"    BlurRadius="10" ShadowDepth="0" Opacity="0.55" Color="#8FB4FF"/>
    <DropShadowEffect x:Key="ShadowGreen"   BlurRadius="10" ShadowDepth="0" Opacity="0.55" Color="#9FD7B8"/>
    <DropShadowEffect x:Key="ShadowRed"     BlurRadius="10" ShadowDepth="0" Opacity="0.55" Color="#F7C4C4"/>

    <Style x:Key="BtnBase" TargetType="Button">
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Padding" Value="12,0"/>
      <Setter Property="Height" Value="32"/>
      <Style.Triggers>
        <Trigger Property="IsEnabled" Value="False">
          <Setter Property="Effect" Value="{x:Null}"/>
          <Setter Property="Background" Value="#ECEFF3"/>
          <Setter Property="Foreground" Value="#9CA3AF"/>
          <Setter Property="Opacity" Value="0.8"/>
        </Trigger>
      </Style.Triggers>
    </Style>

    <Style x:Key="BtnPrimary" TargetType="Button" BasedOn="{StaticResource BtnBase}">
      <Setter Property="Background" Value="#9FAEF7"/>
      <Setter Property="Foreground" Value="#1F2D3A"/>
      <Setter Property="Effect" Value="{StaticResource ShadowPrimary}"/>
    </Style>

    <Style x:Key="BtnBlue" TargetType="Button" BasedOn="{StaticResource BtnBase}">
      <Setter Property="Background" Value="#8FB4FF"/>
      <Setter Property="Foreground" Value="#1F2D3A"/>
      <Setter Property="Effect" Value="{StaticResource ShadowBlue}"/>
    </Style>

    <Style x:Key="BtnGreen" TargetType="Button" BasedOn="{StaticResource BtnBase}">
      <Setter Property="Background" Value="#9FD7B8"/>
      <Setter Property="Foreground" Value="#1F2D3A"/>
      <Setter Property="Effect" Value="{StaticResource ShadowGreen}"/>
    </Style>

    <Style x:Key="BtnRed" TargetType="Button" BasedOn="{StaticResource BtnBase}">
      <Setter Property="Background" Value="#F7C4C4"/>
      <Setter Property="Foreground" Value="#1F2D3A"/>
      <Setter Property="Effect" Value="{StaticResource ShadowRed}"/>
    </Style>

    <Style x:Key="Card" TargetType="Border">
      <Setter Property="Background" Value="#FFFFFF"/>
      <Setter Property="BorderBrush" Value="#E4E9F0"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="6"/>
      <Setter Property="Padding" Value="12"/>
      <Setter Property="Margin" Value="0,0,0,10"/>
    </Style>

    <Style x:Key="LabelSmall" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#475467"/>
      <Setter Property="FontSize" Value="12"/>
    </Style>

    <Style x:Key="TitleSmall" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#0F172A"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Margin" Value="0,0,0,8"/>
    </Style>

    <Style x:Key="GridHeaderStyle" TargetType="{x:Type DataGridColumnHeader}">
      <Setter Property="Background" Value="#F3F6FB"/>
      <Setter Property="Foreground" Value="#0F172A"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Padding" Value="8,6"/>
      <Setter Property="BorderBrush" Value="#DDE5F0"/>
      <Setter Property="BorderThickness" Value="0,0,0,1"/>
      <Setter Property="HorizontalContentAlignment" Value="Left"/>
    </Style>

    <Style x:Key="GridRowStyle" TargetType="{x:Type DataGridRow}">
      <Setter Property="MinHeight" Value="28"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Style.Triggers>
        <DataTrigger Binding="{Binding Enabled}" Value="False">
          <Setter Property="Background" Value="#FDECEC"/>
          <Setter Property="Foreground" Value="#7F1D1D"/>
        </DataTrigger>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Background" Value="#DBEAFE"/>
          <Setter Property="Foreground" Value="#0F172A"/>
        </Trigger>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Background" Value="#EFF6FF"/>
        </Trigger>
      </Style.Triggers>
    </Style>

    <Style x:Key="GridCellStyle" TargetType="{x:Type DataGridCell}">
      <Setter Property="Padding" Value="8,5"/>
      <Setter Property="BorderThickness" Value="0,0,0,1"/>
      <Setter Property="BorderBrush" Value="#EEF2F7"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Setter Property="Foreground" Value="#111827"/>
      <Setter Property="Background" Value="Transparent"/>
      <Style.Triggers>
        <DataTrigger Binding="{Binding Enabled}" Value="False">
          <Setter Property="Background" Value="#FDECEC"/>
        </DataTrigger>
      </Style.Triggers>
    </Style>

    <Style x:Key="GridTextTrim" TargetType="{x:Type TextBlock}">
      <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
    </Style>

    <Style x:Key="GridTextCenter" TargetType="{x:Type TextBlock}" BasedOn="{StaticResource GridTextTrim}">
      <Setter Property="TextAlignment" Value="Center"/>
    </Style>

    <Style x:Key="GridCheckBoxStyle" TargetType="{x:Type CheckBox}">
      <Setter Property="HorizontalAlignment" Value="Center"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="Focusable" Value="False"/>
    </Style>
  </Window.Resources>

  <Grid>
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="280"/>
      <ColumnDefinition Width="*"/>
    </Grid.ColumnDefinitions>

    <!-- ================= Sidebar ================= -->
    <Border Grid.Column="0" Background="#FFFFFF" BorderBrush="#E6EBF4" BorderThickness="0,0,1,0">
      <DockPanel LastChildFill="True">

        <!-- App Header -->
        <StackPanel DockPanel.Dock="Top" Margin="18,18,18,12">
          <StackPanel Orientation="Horizontal">
            <Border Width="36" Height="36" Background="#9AB8FF" CornerRadius="6">
              <TextBlock Text="D" Foreground="#1F2D3A" FontSize="18" FontWeight="Bold"
                         VerticalAlignment="Center" HorizontalAlignment="Center"/>
            </Border>
            <StackPanel Margin="10,0,0,0">
              <TextBlock Text="Device Cleanup Manager" FontSize="16" FontWeight="SemiBold" Foreground="#1F2D3A"/>
              <TextBlock Text="Inactive and Disabled AD Computer" FontSize="11" Foreground="#5F6B7A"/>
              <TextBlock Text="Management" FontSize="11" Foreground="#5F6B7A"/>
            </StackPanel>
          </StackPanel>
        </StackPanel>

        <!-- Footer -->
        <Border DockPanel.Dock="Bottom" BorderBrush="#E6EBF4" BorderThickness="0,1,0,0" Padding="14" Background="#FFFFFF">
          <StackPanel>
            <TextBlock Text="Device Cleanup Manager" FontSize="13" FontWeight="Bold" Foreground="#1F2D3A"/>
            <TextBlock x:Name="FooterVersion" Text="Version" FontSize="11" Foreground="#5F6B7A" Margin="0,4,0,0"/>
            <TextBlock FontSize="11" Foreground="#7C8BA1" Margin="0,8,0,0">
              <Run Text="© 2025 "/>
              <Hyperlink x:Name="FooterLink" NavigateUri="https://www.linkedin.com/in/mabdulkadr/">Mohammad Omar</Hyperlink>
            </TextBlock>
          </StackPanel>
        </Border>

        <!-- Nav -->
        <StackPanel DockPanel.Dock="Top" Margin="8,8">
          <TextBlock Text="OPERATIONS" Margin="14,10,0,6" FontSize="11" FontWeight="SemiBold" Foreground="#7C8BA1"/>
          <Button Content="Device Cleanup" FontWeight="SemiBold" Height="38" Margin="6" Padding="12,0"
                  HorizontalContentAlignment="Left"
                  Background="#D8E2F4" Foreground="#1F2D3A" BorderThickness="0"
                  ToolTip="Current operation view"/>
        </StackPanel>

        <!-- Session + About -->
        <Grid>
          <StackPanel VerticalAlignment="Bottom">

            <Border Style="{StaticResource Card}" Margin="12,0,12,8">
              <StackPanel>
                <TextBlock Text="Session" Style="{StaticResource TitleSmall}"/>
                <Grid Margin="0,4,0,0">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                  </Grid.ColumnDefinitions>
                  <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                  </Grid.RowDefinitions>

                  <TextBlock Grid.Row="0" Grid.Column="0" Text="Machine:" Foreground="#111827" FontWeight="SemiBold" Margin="0,0,8,6"/>
                  <Border Grid.Row="0" Grid.Column="1" Background="#EEF2FF" Padding="6,2" CornerRadius="4" Margin="0,0,0,6">
                    <TextBlock x:Name="SessionMachineTxt" Text="..." Foreground="#1D4ED8" FontSize="12"/>
                  </Border>

                  <TextBlock Grid.Row="1" Grid.Column="0" Text="User:" Foreground="#111827" FontWeight="SemiBold" Margin="0,0,8,6"/>
                  <Border Grid.Row="1" Grid.Column="1" Background="#ECFDF3" Padding="6,2" CornerRadius="4" Margin="0,0,0,6">
                    <TextBlock x:Name="SessionUserTxt" Text="..." Foreground="#166534" FontSize="12"/>
                  </Border>

                  <TextBlock Grid.Row="2" Grid.Column="0" Text="Elevation:" Foreground="#111827" FontWeight="SemiBold" Margin="0,0,8,0"/>
                  <Border Grid.Row="2" Grid.Column="1" x:Name="SessionElevationPill" Background="#ECFDF3" Padding="6,2" CornerRadius="4">
                    <TextBlock x:Name="SessionElevationTxt" Text="..." Foreground="#166534" FontSize="12"/>
                  </Border>
                </Grid>
              </StackPanel>
            </Border>

            <Border Style="{StaticResource Card}" Margin="12,0,12,14">
              <StackPanel>
                <TextBlock Text="About" Style="{StaticResource TitleSmall}"/>
                <TextBlock TextWrapping="Wrap" Foreground="#475467" FontSize="12"
                           Text="Centralized tool to locate stale AD computer objects, review status, and perform safe bulk disable or deletion with confirmation and protection rules."/>
              </StackPanel>
            </Border>

          </StackPanel>
        </Grid>
      </DockPanel>
    </Border>

    <!-- ================= Main Content ================= -->
    <Grid Grid.Column="1">
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
      </Grid.RowDefinitions>

      <!-- Header -->
      <Border Grid.Row="0" Padding="18,14,18,8" Background="#F6F8FB">
        <Grid>
          <StackPanel>
            <TextBlock Text="Device Cleanup Dashboard" FontSize="20" FontWeight="Bold" Foreground="#1F2D3A"/>
            <TextBlock Text="Identify inactive and disabled directory devices and execute controlled cleanup actions."
                       FontSize="14" Foreground="#5F6B7A" Margin="0,6,0,0"/>
          </StackPanel>
        </Grid>
      </Border>

      <Grid Grid.Row="1" Margin="16,0,16,12">
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="10"/>
          <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <!-- Search Criteria -->
        <Border Grid.Row="0" Style="{StaticResource Card}" Margin="0">
          <StackPanel>
            <TextBlock Text="Scope and Filters" Style="{StaticResource TitleSmall}"/>

            <Grid Margin="0,2,0,0">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="100"/>
                <ColumnDefinition Width="16"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
              </Grid.RowDefinitions>

              <TextBlock Grid.Row="0" Grid.Column="0" Text="Inactive Threshold (Days):" Foreground="#1F2D3A" VerticalAlignment="Center"/>
              <TextBox  Grid.Row="0" Grid.Column="1" x:Name="DaysInactiveBox" Height="28" Text="180" Margin="8,0,0,0"
                        Background="#FFFFFF" BorderBrush="#E4E9F0" BorderThickness="1" Padding="8,3"/>

              <TextBlock Grid.Row="0" Grid.Column="3" Text="OU Scope (LDAP):" Foreground="#1F2D3A" VerticalAlignment="Center"/>
              <Grid Grid.Row="0" Grid.Column="4" Margin="8,0,0,0">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBox x:Name="OUScopeText" Height="28"
                         IsReadOnly="True"
                         Background="#FFFFFF" BorderBrush="#E4E9F0" BorderThickness="1" Padding="8,3"
                         ToolTip="Selected OU scope for search"/>
                <Button Grid.Column="1" x:Name="PickOUBtn" Content="Browse..." MinWidth="90" Height="28" Margin="8,0,0,0"
                        Style="{StaticResource BtnBlue}" ToolTip="Open OU selector window with fast search"/>
              </Grid>
              <Button Grid.Row="0" Grid.Column="5" x:Name="SearchBtn" Content="Run Search"
                      MinWidth="124" Margin="10,0,0,0" Style="{StaticResource BtnPrimary}"
                      ToolTip="Run device search using current scope and filters"/>
            </Grid>

            <Grid Margin="0,8,0,0">
              <Border BorderBrush="#E4E9F0" BorderThickness="1" CornerRadius="6" Padding="12" Background="#FBFCFF">
                <StackPanel>
                  <TextBlock Text="Cleanup Target" Style="{StaticResource TitleSmall}" Margin="0,0,0,6"/>
                  <WrapPanel>
                    <RadioButton x:Name="ModeInactiveOnly" Content="Inactive Devices (by days)" IsChecked="True" Margin="0,0,14,0"/>
                    <RadioButton x:Name="ModeDisabledOnly" Content="Disabled Computers" Margin="0,0,14,0"/>
                    <CheckBox x:Name="IncludeNeverLogon" Content="Include devices with no logon timestamp" Margin="10,0,0,0" VerticalAlignment="Center"/>
                  </WrapPanel>
                </StackPanel>
              </Border>
            </Grid>
          </StackPanel>
        </Border>

        <!-- Main Panels -->
        <Grid Grid.Row="2">
          <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="130"/>
          </Grid.RowDefinitions>

          <!-- Right: Results -->
          <Border Grid.Row="0" Grid.ColumnSpan="2" Style="{StaticResource Card}" Margin="0,0,0,8">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
              </Grid.RowDefinitions>

              <Grid Grid.Row="0" Margin="0,0,0,8">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Text="Search Results" Style="{StaticResource TitleSmall}" Margin="0"/>
                <StackPanel Grid.Column="1" Orientation="Horizontal">
                  <CheckBox x:Name="SelectAllGridCheck" Content="Select All" VerticalAlignment="Center" Margin="0,0,10,0"
                            ToolTip="Select or clear all devices in the current grid"/>
                  <Button x:Name="ImportCsvBtn" Content="Import Names (CSV)" MinWidth="126" Height="28" FontSize="12" Style="{StaticResource BtnBlue}" Margin="0,0,8,0"
                          ToolTip="Import device names from CSV (ComputerName or Name column)"/>
                  <Button x:Name="ClearBtn" Content="Clear Results" MinWidth="98" Height="28" FontSize="12" Style="{StaticResource BtnBlue}" Margin="0,0,8,0"
                          ToolTip="Clear current results from the grid"/>
                  <Button x:Name="ExportBtn" Content="Export Results CSV" MinWidth="122" Height="28" FontSize="12" Style="{StaticResource BtnGreen}"
                          ToolTip="Export current results grid to CSV"/>
                </StackPanel>
              </Grid>

              <DataGrid Grid.Row="1" x:Name="ComputerGrid" AutoGenerateColumns="False" CanUserAddRows="False" IsReadOnly="False"
                        SelectionMode="Extended" SelectionUnit="FullRow"
                        CanUserResizeColumns="True" CanUserReorderColumns="True" CanUserSortColumns="True"
                        RowHeaderWidth="0" AlternationCount="2"
                        RowBackground="#FFFFFF" AlternatingRowBackground="#F8FAFD"
                        GridLinesVisibility="Horizontal" HorizontalGridLinesBrush="#EEF2F7"
                        ScrollViewer.HorizontalScrollBarVisibility="Auto"
                        ScrollViewer.VerticalScrollBarVisibility="Auto"
                        Background="#FFFFFF" Margin="0,0,0,8"
                        BorderBrush="#DDE5F0" BorderThickness="1"
                        ColumnHeaderStyle="{StaticResource GridHeaderStyle}"
                        RowStyle="{StaticResource GridRowStyle}"
                        CellStyle="{StaticResource GridCellStyle}">
                <DataGrid.Columns>
                  <DataGridCheckBoxColumn Header="Select"
                                          Binding="{Binding IsSelected, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"
                                          ElementStyle="{StaticResource GridCheckBoxStyle}"
                                          EditingElementStyle="{StaticResource GridCheckBoxStyle}"
                                          Width="62"/>
                  <DataGridTextColumn Header="Computer Name" Binding="{Binding ComputerName}" Width="155" IsReadOnly="True" ElementStyle="{StaticResource GridTextTrim}"/>
                  <DataGridTextColumn Header="Enabled" Binding="{Binding Enabled}" Width="80" IsReadOnly="True" ElementStyle="{StaticResource GridTextCenter}"/>
                  <DataGridTextColumn Header="Last Logon" Binding="{Binding LastLogonDate}" Width="175" IsReadOnly="True" ElementStyle="{StaticResource GridTextTrim}"/>
                  <DataGridTextColumn Header="Inactive Days" Binding="{Binding InactiveDays}" Width="95" IsReadOnly="True" ElementStyle="{StaticResource GridTextCenter}"/>
                  <DataGridTextColumn Header="OU Path" Binding="{Binding OUPath}" Width="230" IsReadOnly="True" ElementStyle="{StaticResource GridTextTrim}"/>
                  <DataGridTextColumn Header="Distinguished Name" Binding="{Binding DistinguishedName}" Width="620" IsReadOnly="True" ElementStyle="{StaticResource GridTextTrim}"/>
                  <DataGridTextColumn Header="Data Source" Binding="{Binding Source}" Width="100" IsReadOnly="True" ElementStyle="{StaticResource GridTextCenter}"/>
                </DataGrid.Columns>
              </DataGrid>

              <UniformGrid Grid.Row="2" Columns="3" Margin="0,2,0,0">
                <Button x:Name="DisableSelectedBtn" Content="Disable Selected" Style="{StaticResource BtnRed}" Margin="0,0,8,0"
                        ToolTip="Disable selected devices in Active Directory"/>
                <Button x:Name="EnableSelectedBtn" Content="Enable Selected" Style="{StaticResource BtnGreen}" Margin="0,0,8,0"
                        ToolTip="Enable selected devices in Active Directory"/>
                <Button x:Name="DeleteSelectedBtn" Content="Delete Selected" Style="{StaticResource BtnRed}"
                        ToolTip="Delete selected devices from Active Directory"/>
              </UniformGrid>

              <Grid Grid.Row="3" Margin="0,8,0,0">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" x:Name="StatusLabel" Text="Ready. Configure filters and click Run Search." Foreground="#111827" VerticalAlignment="Center"/>
                <ProgressBar Grid.Column="1" x:Name="ProgressBar" Width="170" Height="14" Visibility="Hidden" IsIndeterminate="True" Margin="8,0,0,0" VerticalAlignment="Center"/>
              </Grid>
            </Grid>
          </Border>

          <!-- Bottom Row: Message Center (full width) -->
          <Border Grid.Row="1" Grid.ColumnSpan="2" Style="{StaticResource Card}" Margin="0">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>
              <RichTextBox Grid.Row="1" x:Name="LogBox" IsReadOnly="True"
                           Background="#1F2D3A" Foreground="#E4E9F0"
                           BorderBrush="#1F2937" BorderThickness="1"
                           FontFamily="Consolas" FontSize="13"
                           VerticalScrollBarVisibility="Auto"
                           Padding="5"/>
            </Grid>
          </Border>
        </Grid>
      </Grid>
    </Grid>

  </Grid>
</Window>
'@
#endregion

#region ============================== LOAD UI / BIND CONTROLS =====================================
$xml = [xml]$xamlText
$reader = New-Object System.Xml.XmlNodeReader($xml)
$Window = [System.Windows.Markup.XamlReader]::Load($reader)
$script:Window = $Window

# Named controls
$script:DaysInactiveBox      = $Window.FindName('DaysInactiveBox')
$script:OUScopeText          = $Window.FindName('OUScopeText')
$script:PickOUBtn            = $Window.FindName('PickOUBtn')
$script:SearchBtn            = $Window.FindName('SearchBtn')
$script:ClearBtn             = $Window.FindName('ClearBtn')
$script:ExportBtn            = $Window.FindName('ExportBtn')

$script:ModeInactiveOnly     = $Window.FindName('ModeInactiveOnly')
$script:ModeDisabledOnly     = $Window.FindName('ModeDisabledOnly')
$script:IncludeNeverLogon    = $Window.FindName('IncludeNeverLogon')

$script:ComputerGrid         = $Window.FindName('ComputerGrid')
$script:DisableSelectedBtn   = $Window.FindName('DisableSelectedBtn')
$script:EnableSelectedBtn    = $Window.FindName('EnableSelectedBtn')
$script:DeleteSelectedBtn    = $Window.FindName('DeleteSelectedBtn')

$script:ProgressBar          = $Window.FindName('ProgressBar')
$script:StatusLabel          = $Window.FindName('StatusLabel')
$script:LogBox               = $Window.FindName('LogBox')

$script:ImportCsvBtn         = $Window.FindName('ImportCsvBtn')
$script:SelectAllGridCheck   = $Window.FindName('SelectAllGridCheck')

$FooterVersion               = $Window.FindName('FooterVersion')
$SessionMachineTxt           = $Window.FindName('SessionMachineTxt')
$SessionUserTxt              = $Window.FindName('SessionUserTxt')
$SessionElevationTxt         = $Window.FindName('SessionElevationTxt')
$SessionElevationPill        = $Window.FindName('SessionElevationPill')

# Session text
$FooterVersion.Text = "Version $AppVersion"
$SessionMachineTxt.Text = $env:COMPUTERNAME
$SessionUserTxt.Text    = $env:USERNAME

$IsAdmin = $false
try {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $p  = New-Object Security.Principal.WindowsPrincipal($id)
    $IsAdmin = $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
} catch {}

if ($IsAdmin) {
    $SessionElevationTxt.Text = 'Admin'
    $SessionElevationPill.Background = New-Brush '#ECFDF3'
} else {
    $SessionElevationTxt.Text = 'User'
    $SessionElevationPill.Background = New-Brush '#FEF3C7'
}
#endregion

#region ============================== EVENT WIRE-UP ===============================================
$script:SearchBtn.Add_Click({
    Invoke-SafeUIAction -Action { Start-Search } -ActionName 'Run Search'
}) | Out-Null
$script:ModeInactiveOnly.Add_Checked({ Update-ModeUIState }) | Out-Null
$script:ModeDisabledOnly.Add_Checked({ Update-ModeUIState }) | Out-Null
$script:PickOUBtn.Add_Click({
    Invoke-SafeUIAction -Action {
        if (@($script:OUAllItems).Count -eq 0) {
            Add-LogLine -Message 'OU list is empty. Load OUs first.' -Level 'WARNING'
            return
        }

        $currentDN = ''
        if ($script:SelectedOUItem) { $currentDN = ($script:SelectedOUItem.DN + '') }

        $picked = Show-OUSelectorDialog -Items $script:OUAllItems -CurrentDN $currentDN
        if ($picked) {
            Set-SelectedOUScope -Item $picked
            Add-LogLine -Message ("OU scope selected: {0}" -f ($picked.Display + '')) -Level 'INFO'
        }
    } -ActionName 'Browse OU'
}) | Out-Null

$script:ClearBtn.Add_Click({
    Invoke-SafeUIAction -Action { Clear-Results } -ActionName 'Clear Results'
}) | Out-Null
$script:ExportBtn.Add_Click({
    Invoke-SafeUIAction -Action { Export-GridToCsv } -ActionName 'Export Results CSV'
}) | Out-Null

$script:DisableSelectedBtn.Add_Click({
    Invoke-SafeUIAction -Action { Invoke-ActionOnSelected -Action 'Disable' } -ActionName 'Disable Selected'
}) | Out-Null
$script:EnableSelectedBtn.Add_Click({
    Invoke-SafeUIAction -Action { Invoke-ActionOnSelected -Action 'Enable' } -ActionName 'Enable Selected'
}) | Out-Null
$script:DeleteSelectedBtn.Add_Click({
    Invoke-SafeUIAction -Action { Invoke-ActionOnSelected -Action 'Delete' } -ActionName 'Delete Selected'
}) | Out-Null

$script:ImportCsvBtn.Add_Click({
    Invoke-SafeUIAction -Action { Import-CsvToGrid } -ActionName 'Import Names (CSV)'
}) | Out-Null

if ($script:SelectAllGridCheck) {
    $script:SelectAllGridCheck.Add_Checked({
        Invoke-SafeUIAction -Action { Set-GridCheckedState -Checked $true } -ActionName 'Select All'
    }) | Out-Null

    $script:SelectAllGridCheck.Add_Unchecked({
        Invoke-SafeUIAction -Action { Set-GridCheckedState -Checked $false } -ActionName 'Clear Selection'
    }) | Out-Null
}

# Ensure checkbox selection works on single click in the first grid column.
$script:ComputerGrid.Add_PreviewMouseLeftButtonDown({
    param($sender, $e)

    try {
        $dep = $e.OriginalSource
        while ($dep -and -not ($dep -is [System.Windows.Controls.DataGridCell])) {
            $dep = [System.Windows.Media.VisualTreeHelper]::GetParent($dep)
        }

        if (-not $dep) { return }
        $cell = [System.Windows.Controls.DataGridCell]$dep
        if (-not $cell.Column -or $cell.Column.DisplayIndex -ne 0) { return }

        $row = $cell.DataContext
        if (-not $row) { return }
        if ($row.PSObject.Properties.Name -notcontains 'IsSelected') { return }

        $row.IsSelected = -not [bool]$row.IsSelected
        $e.Handled = $true

        try {
            [void]$script:ComputerGrid.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Cell, $true)
            [void]$script:ComputerGrid.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row,  $true)
        } catch {}

        try { $script:ComputerGrid.Items.Refresh() } catch {}
    }
    catch {}
}) | Out-Null

# Cleanup on close
$Window.Add_Closing({
    try {
        if ($script:ScanTimer) { $script:ScanTimer.Stop() }
        if ($script:ScanJob) {
            if ($script:ScanJob.State -eq 'Running') { Stop-Job -Job $script:ScanJob -Force | Out-Null }
            Remove-Job -Job $script:ScanJob -Force | Out-Null
        }
    } catch {}
}) | Out-Null
#endregion

#region ============================== STARTUP ======================================================
Add-LogLine -Message ("{0} starting..." -f $ToolName) -Level 'INFO'

$script:StartupInitialized = $false
$script:StatusLabel.Text = 'Loading UI...'

# Show window first, then load OUs when the UI is fully rendered.
$Window.Add_ContentRendered({
    if ($script:StartupInitialized) { return }
    $script:StartupInitialized = $true

    try {
        Initialize-OUComboBox
        Update-ModeUIState
        $script:StatusLabel.Text = 'Ready. Configure filters and click Run Search.'
    }
    catch {
        $script:StatusLabel.Text = 'Ready (OU list load failed).'
        Add-LogLine -Message ("Startup OU load failed: {0}" -f $_.Exception.Message) -Level 'ERROR'
    }
}) | Out-Null

[void]$Window.ShowDialog()
#endregion
