param([switch]$SmokeTest)
#Requires -Version 5.1
# ==============================================================================
#  R.A - Inventory Search + Alter Provisioning  (versi PowerShell/WPF)
#  Cara pakai (gaya Chris Titus WinUtil, tanpa download manual):
#    irm https://raw.githubusercontent.com/USERNAME/REPO/main/RA.ps1 | iex
#  (jalankan di terminal yang Run as Administrator bila diwajibkan)
#  Versi EXE (CustomTkinter) tetap ada dan tidak berubah.
# ==============================================================================

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

$Script:SheetId   = '1EDQS4zmWKHM7MJf1BjZVB2BYgnwutnXZwB8e0SvnOYY'
$Script:SheetName = 'IP BOGOR'
$Script:Rows      = @()

# ------------------------------- fungsi inti --------------------------------

function Get-IsDarkMode {
    try {
        $v = Get-ItemPropertyValue -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -Name 'AppsUseLightTheme'
        return ($v -eq 0)
    } catch { return $false }
}

function Num-Eq {
    param([string]$A, [string]$B)
    $a = ($A | Out-String).Trim(); $b = ($B | Out-String).Trim()
    $ai = 0; $bi = 0
    if ([int]::TryParse($a, [ref]$ai) -and [int]::TryParse($b, [ref]$bi)) { return $ai -eq $bi }
    return ($a -ceq $b)
}

function Extract-IpSlotPort {
    param([string]$Text)
    $t = [string]$Text
    $out = @{ IP = ''; Slot = ''; Port = '' }
    $m = [regex]::Match($t, '\b(?:\d{1,3}\.){3}\d{1,3}\b')
    if ($m.Success) { $out.IP = $m.Value }
    $m = [regex]::Match($t, '(?i)slot\s*[:=\-]?\s*(\d{1,2})')
    if ($m.Success) { $out.Slot = $m.Groups[1].Value }
    $m = [regex]::Match($t, '(?i)port\s*[:=\-]?\s*(\d{1,3})')
    if ($m.Success) { $out.Port = $m.Groups[1].Value }
    if (-not $out.Slot -or -not $out.Port) {
        $m = [regex]::Match($t, '\b(\d{1,2})[\/\-:](\d{1,2})[\/\-:](\d{1,2})\b')
        if ($m.Success) {
            if (-not $out.Slot) { $out.Slot = $m.Groups[2].Value }
            if (-not $out.Port) { $out.Port = $m.Groups[3].Value }
        } else {
            $m = [regex]::Match($t, '\b(\d{1,2})/(\d{1,3})\b')
            if ($m.Success) {
                if (-not $out.Slot) { $out.Slot = $m.Groups[1].Value }
                if (-not $out.Port) { $out.Port = $m.Groups[2].Value }
            }
        }
    }
    return $out
}

function Extract-Services {
    param([string]$Text)
    $out = @{ INTERNET = ''; IPTV = ''; VOICE = '' }
    foreach ($m in [regex]::Matches([string]$Text, '([A-Za-z0-9][\w\-.]*_(?:INTERNET|IPTV|VOICE))\b', 'IgnoreCase')) {
        $token = $m.Groups[1].Value
        $tipe = $token.Substring($token.LastIndexOf('_') + 1).ToUpper()
        if ($out.ContainsKey($tipe) -and -not $out[$tipe]) { $out[$tipe] = $token }
    }
    return $out
}

function Write-ProvisionCsv($path, $rows) {
    $lines = @('"RESOURCE_ID","SERVICE_NAME","TARGET_ID","CONFIG_ITEM_NAME"')
    foreach ($row in $rows) {
        $f = @($row.RESOURCE_ID, $row.SERVICE_NAME, $row.TARGET_ID, $row.CONFIG) |
            ForEach-Object { ([string]$_).Replace('"', '""') }
        $lines += '"{0}","{1}","{2}","{3}"' -f $f[0], $f[1], $f[2], $f[3]
    }
    [IO.File]::WriteAllText($path, ($lines -join "`n"), (New-Object Text.UTF8Encoding $false))
}

function Get-DownloadFolder {
    try {
        $p = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path
        if ($p -and (Test-Path $p)) { return $p }
    } catch { }
    $p = Join-Path $home 'Downloads'
    if (Test-Path $p) { return $p }
    return $env:TEMP
}

function Get-SheetRows {
    $url = "https://docs.google.com/spreadsheets/d/$Script:SheetId/gviz/tq?tqx=out:csv&sheet=" + [uri]::EscapeDataString($Script:SheetName)
    $txt = Invoke-RestMethod -Uri $url -TimeoutSec 30
    if ($txt -match '^\s*<!DOCTYPE') { throw 'Sheet menolak akses (perlu Share publik).' }
    $rows = $txt | ConvertFrom-Csv
    if (-not $rows) { throw 'Sheet kosong.' }
    return @($rows)
}

function Get-FallbackRows {
    return @(
        [pscustomobject]@{ IP = '172.28.114.62'; SLOT = '1'; PORT = '1'; VLAN_NET = '3101'; VLAN_VOIP = '502'; ID_PORT = '12308999-75846529'; GPON = 'GPON04-D2-KHL-4'; VENDOR = 'HUAWEI'; WITEL = 'BOGOR' },
        [pscustomobject]@{ IP = '172.28.114.62'; SLOT = '1'; PORT = '2'; VLAN_NET = '3101'; VLAN_VOIP = '503'; ID_PORT = '12309000-75846530'; GPON = 'GPON04-D2-KHL-5'; VENDOR = 'HUAWEI'; WITEL = 'BOGOR' }
    )
}

# ------------------------------ smoke test ----------------------------------
if ($SmokeTest) {
    $gagal = 0
    $t1 = 'The TT is created for CRM order ID: S950VIS4XS41SIEWBT2MVWOEO-PDAk426092109423877932ca60_101072695-776390463~2026092109444743131946~43131946~144370441~3~WSA OSM ID: 124554512. Service ID is 3-4VDS67DL_122363300898_INTERNET,66250691_122363300898_IPTV'
    $t2 = 'The TT is created for CRM order ID: 2BYJ8AEQIBYWDRM4DATTDF28B-AOk4260505021154089b977c0_78541761-578023769~2026050514133036214749~36214749~120602104~3~WSA OSM ID: 111784540. Service ID is 1-QODQ3UG_122303221735_INTERNET,1-QODQ3UG_02512024599_VOICE'
    $s1 = Extract-Services $t1
    $s2 = Extract-Services $t2
    if ($s1.INTERNET -ne '3-4VDS67DL_122363300898_INTERNET') { 'FAIL T1.INTERNET'; $gagal++ }
    if ($s1.IPTV -ne '66250691_122363300898_IPTV') { 'FAIL T1.IPTV'; $gagal++ }
    if ($s1.VOICE -ne '') { 'FAIL T1.VOICE'; $gagal++ }
    if ($s2.INTERNET -ne '1-QODQ3UG_122303221735_INTERNET') { 'FAIL T2.INTERNET'; $gagal++ }
    if ($s2.VOICE -ne '1-QODQ3UG_02512024599_VOICE') { 'FAIL T2.VOICE'; $gagal++ }
    if ($s2.IPTV -ne '') { 'FAIL T2.IPTV'; $gagal++ }
    $v = Extract-IpSlotPort "IP OLT: 172.28.114.7`r`nSlot: 6`r`nPort: 8"
    if ($v.IP -ne '172.28.114.7' -or $v.Slot -ne '6' -or $v.Port -ne '8') { 'FAIL VALINS'; $gagal++ }
    $v = Extract-IpSlotPort 'ip 172.28.114.62 0/1/2'
    if ($v.Slot -ne '1' -or $v.Port -ne '2') { 'FAIL OLT'; $gagal++ }
    if (-not (Num-Eq '06' '6')) { 'FAIL NUMEQ'; $gagal++ }
    if (Num-Eq '16' '6') { 'FAIL NUMEQ2'; $gagal++ }
    $dlTest = Get-DownloadFolder
    if (-not (Test-Path $dlTest)) { 'FAIL DLPATH'; $gagal++ }
    if ($gagal -eq 0) { 'SMOKE OK' } else { "SMOKE GAGAL: $gagal" ; exit 1 }
    exit 0
}

# --------------------------------- XAML -------------------------------------
$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="R.A" Height="840" Width="1440" MinWidth="1200" MinHeight="700"
        WindowStartupLocation="CenterScreen">
  <Window.Resources>
    <Style x:Key="RndBtn" TargetType="Button">
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="FontWeight" Value="Bold"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="Padding" Value="8,6"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border Background="{TemplateBinding Background}" CornerRadius="10" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="RndBox" TargetType="TextBox">
      <Setter Property="Padding" Value="10,8"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="BorderBrush" Value="#E2E5F1"/>
      <Setter Property="BorderThickness" Value="1"/>
    </Style>
  </Window.Resources>
  <Grid x:Name="Root" Background="#EDEFF7">
    <Grid.RowDefinitions>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <Grid Grid.Row="0" Margin="14">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="480"/>
      </Grid.ColumnDefinitions>

      <!-- KARTU KIRI -->
      <Border Grid.Column="0" CornerRadius="16" Background="White" Margin="0,0,7,0" Padding="16,12,16,12" x:Name="CardLeft">
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>
          <DockPanel Grid.Row="0" Margin="0,0,0,6">
            <TextBlock Text="Inventory Search" FontSize="16" FontWeight="Bold" Foreground="#4F46E5" VerticalAlignment="Center"/>
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
              <Button x:Name="btnRefresh" Content="Refresh" Width="70" Height="26" Margin="0,0,6,0" Background="#E8E9FB" Foreground="#4F46E5" Style="{StaticResource RndBtn}" FontSize="11"/>
              <Button x:Name="btnReset" Content="Reset" Width="60" Height="26" Background="Transparent" Foreground="#4F46E5" Style="{StaticResource RndBtn}" FontSize="11"/>
            </StackPanel>
          </DockPanel>
          <Grid Grid.Row="1" Margin="0,0,0,8">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="90"/>
              <ColumnDefinition Width="90"/>
              <ColumnDefinition Width="120"/>
            </Grid.ColumnDefinitions>
            <TextBox x:Name="txtIP" Grid.Column="0" Margin="0,0,8,0" Height="40" Style="{StaticResource RndBox}"/>
            <TextBox x:Name="txtSlot" Grid.Column="1" Margin="0,0,8,0" Height="40" HorizontalContentAlignment="Center" Style="{StaticResource RndBox}"/>
            <TextBox x:Name="txtPort" Grid.Column="2" Margin="0,0,8,0" Height="40" HorizontalContentAlignment="Center" Style="{StaticResource RndBox}"/>
            <Button x:Name="btnCari" Grid.Column="3" Content="Cari" Height="40" Background="#4F46E5" Style="{StaticResource RndBtn}"/>
          </Grid>
          <Grid Grid.Row="2" Margin="0,0,0,8">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <TextBox x:Name="txtValins" Grid.Column="0" Margin="0,0,8,0" Height="86" AcceptsReturn="True" VerticalScrollBarVisibility="Auto" TextWrapping="Wrap" Style="{StaticResource RndBox}"/>
            <TextBox x:Name="txtBima" Grid.Column="1" Height="86" AcceptsReturn="True" VerticalScrollBarVisibility="Auto" TextWrapping="Wrap" Style="{StaticResource RndBox}"/>
          </Grid>
          <Grid Grid.Row="3" Margin="0,0,0,8">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Button x:Name="btnCariData" Grid.Column="0" Margin="0,0,8,0" Content="Cari Data" Height="40" Background="#E8E9FB" Foreground="#4F46E5" Style="{StaticResource RndBtn}"/>
            <Button x:Name="btnEkstrak" Grid.Column="1" Content="Ekstrak Service" Height="40" Background="#E8E9FB" Foreground="#4F46E5" Style="{StaticResource RndBtn}"/>
          </Grid>
          <TextBlock x:Name="lblCount" Grid.Row="4" Margin="2,0,0,4" FontSize="11" Foreground="Gray" Text="Memuat data..."/>
          <DataGrid x:Name="gridHasil" Grid.Row="5" AutoGenerateColumns="False" IsReadOnly="True"
                    SelectionMode="Single" HeadersVisibility="Column" GridLinesVisibility="Horizontal"
                    RowHeight="28" FontSize="12" Background="White" BorderBrush="#E2E5F1" BorderThickness="1">
            <DataGrid.ColumnHeaderStyle>
              <Style TargetType="DataGridColumnHeader">
                <Setter Property="Background" Value="#F1F2F9"/>
                <Setter Property="Foreground" Value="#33334D"/>
                <Setter Property="FontWeight" Value="Bold"/>
                <Setter Property="Padding" Value="8,6"/>
              </Style>
            </DataGrid.ColumnHeaderStyle>
            <DataGrid.Columns>
              <DataGridTextColumn Header="VLAN_NET" Binding="{Binding VLAN_NET}" Width="*"/>
              <DataGridTextColumn Header="VLAN_VOIP" Binding="{Binding VLAN_VOIP}" Width="*"/>
              <DataGridTextColumn Header="ID_PORT" Binding="{Binding ID_PORT}" Width="1.4*"/>
              <DataGridTextColumn Header="GPON" Binding="{Binding GPON}" Width="1.4*"/>
            </DataGrid.Columns>
          </DataGrid>
        </Grid>
      </Border>

      <!-- KARTU KANAN -->
      <Border Grid.Column="1" CornerRadius="16" Background="White" Margin="7,0,0,0" Padding="14,12,14,12" x:Name="CardRight">
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>
          <TextBlock Grid.Row="0" Text="Alter Provisioning" FontSize="16" FontWeight="Bold" Foreground="#4F46E5" Margin="0,0,0,8"/>
          <Grid Grid.Row="1" Margin="0,0,0,4">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="0.8*"/>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="34"/>
            </Grid.ColumnDefinitions>
            <TextBlock Grid.Column="0" Text="RESOURCE_ID" FontSize="9" FontWeight="Bold" Foreground="#6B7194"/>
            <TextBlock Grid.Column="1" Text="SERVICE_NAME" FontSize="9" FontWeight="Bold" Foreground="#6B7194"/>
            <TextBlock Grid.Column="2" Text="TARGET_ID" FontSize="9" FontWeight="Bold" Foreground="#6B7194"/>
            <TextBlock Grid.Column="3" Text="CONFIG_ITEM_NAME" FontSize="9" FontWeight="Bold" Foreground="#6B7194"/>
          </Grid>
          <DataGrid x:Name="gridAlter" Grid.Row="2" AutoGenerateColumns="False" CanUserAddRows="False"
                    HeadersVisibility="None" GridLinesVisibility="None" RowHeight="38" FontSize="11"
                    Background="White" BorderThickness="0" Margin="0,0,0,8">
            <DataGrid.Columns>
              <DataGridTextColumn Binding="{Binding RESOURCE_ID, UpdateSourceTrigger=PropertyChanged}" Width="*"/>
              <DataGridTextColumn Binding="{Binding SERVICE_NAME, UpdateSourceTrigger=PropertyChanged}" Width="*"/>
              <DataGridTextColumn Binding="{Binding TARGET_ID, UpdateSourceTrigger=PropertyChanged}" Width="0.8*"/>
              <DataGridComboBoxColumn x:Name="colCfg" SelectedItemBinding="{Binding CONFIG, UpdateSourceTrigger=PropertyChanged}" Width="*"/>
              <DataGridTemplateColumn Width="34">
                <DataGridTemplateColumn.CellTemplate>
                  <DataTemplate>
                    <Button Content="X" Width="28" Height="30" Background="#DC2626" Foreground="White" FontWeight="Bold" Style="{StaticResource RndBtn}" FontSize="11"/>
                  </DataTemplate>
                </DataGridTemplateColumn.CellTemplate>
              </DataGridTemplateColumn>
            </DataGrid.Columns>
          </DataGrid>
          <Grid Grid.Row="3">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Button x:Name="btnAdd" Grid.Column="0" Margin="0,0,6,0" Content="+ Add" Height="44" Background="#6D5CFF" Style="{StaticResource RndBtn}"/>
            <Button x:Name="btnDownload" Grid.Column="1" Margin="0,0,6,0" Content="Download" Height="44" Background="#15924E" Style="{StaticResource RndBtn}"/>
            <Button x:Name="btnClear" Grid.Column="2" Content="Clear" Height="44" Background="#DC2626" Style="{StaticResource RndBtn}"/>
          </Grid>
        </Grid>
      </Border>
    </Grid>
    <TextBlock Grid.Row="1" Text="(c) 2026, Yudha Prastyo. All rights reserved." HorizontalAlignment="Center" Margin="0,0,0,10" FontSize="10" Foreground="Gray"/>
  </Grid>
</Window>
'@

# ---------------------------------- UI --------------------------------------
$reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xaml))
$win = [System.Windows.Markup.XamlReader]::Load($reader)
$UI = @{}
$win.FindName('Root') | Out-Null
foreach ($n in @('txtIP','txtSlot','txtPort','btnCari','txtValins','txtBima','btnCariData','btnEkstrak',
                 'lblCount','gridHasil','btnRefresh','btnReset','gridAlter','colCfg',
                 'btnAdd','btnDownload','btnClear','CardLeft','CardRight','Root')) {
    $UI[$n] = $win.FindName($n)
}

$HintValins = 'Paste Valins...'
$HintBima   = 'Paste Detail BIMA...'
$UI.txtIP.Text = '172.28.114.62'
$UI.txtValins.Text = $HintValins; $UI.txtValins.Foreground = 'Gray'
$UI.txtBima.Text   = $HintBima;   $UI.txtBima.Foreground = 'Gray'
$UI.txtValins.Add_GotFocus({ if ($UI.txtValins.Text -eq $HintValins) { $UI.txtValins.Text = ''; $UI.txtValins.Foreground = 'Black' } })
$UI.txtValins.Add_LostFocus({ if (-not $UI.txtValins.Text.Trim()) { $UI.txtValins.Text = $HintValins; $UI.txtValins.Foreground = 'Gray' } })
$UI.txtBima.Add_GotFocus({ if ($UI.txtBima.Text -eq $HintBima) { $UI.txtBima.Text = ''; $UI.txtBima.Foreground = 'Black' } })
$UI.txtBima.Add_LostFocus({ if (-not $UI.txtBima.Text.Trim()) { $UI.txtBima.Text = $HintBima; $UI.txtBima.Foreground = 'Gray' } })
function Get-Paste($box, $hint) { $t = $box.Text.Trim(); if ($t -eq $hint) { '' } else { $t } }

# Tema ikut Windows
if (Get-IsDarkMode) {
    $UI.Root.Background = '#212121'
    $UI.CardLeft.Background = '#2B2B2B'; $UI.CardRight.Background = '#2B2B2B'
    foreach ($tb in @($UI.txtIP, $UI.txtSlot, $UI.txtPort, $UI.txtValins, $UI.txtBima)) {
        $tb.Background = '#343638'; $tb.Foreground = 'White'; $tb.BorderBrush = '#4A4A4A'
    }
    $UI.gridHasil.Background = '#2B2B2B'; $UI.gridHasil.Foreground = 'White'
    $UI.gridHasil.RowBackground = '#2B2B2B'; $UI.gridHasil.AlternatingRowBackground = '#323232'
    $UI.gridAlter.Background = '#2B2B2B'; $UI.gridAlter.Foreground = 'White'
    $UI.gridAlter.RowBackground = '#2B2B2B'; $UI.gridAlter.AlternatingRowBackground = '#323232'
    $UI.lblCount.Foreground = '#B0B0B0'
}

# Data Alter
$AlterRows = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
function Add-AlterRow {
    $AlterRows.Add([pscustomobject]@{ RESOURCE_ID = ''; SERVICE_NAME = ''; TARGET_ID = ''; CONFIG = 'Service_Port' }) | Out-Null
}
1..5 | ForEach-Object { Add-AlterRow }
$UI.gridAlter.ItemsSource = $AlterRows
$UI.colCfg.ItemsSource = @('Service_Port', 'S-Vlan')
$UI.gridAlter.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent,
    [System.Windows.RoutedEventHandler]{
        param($s, $e)
        $item = $e.OriginalSource.DataContext
        if ($item -and $AlterRows.Contains($item)) { $AlterRows.Remove($item) | Out-Null }
        if ($AlterRows.Count -eq 0) { Add-AlterRow }
    })

function Refresh-AlterGrid {
    $UI.gridAlter.ItemsSource = $null
    $UI.gridAlter.ItemsSource = $AlterRows
}

function Fill-AlterFromSearch($row) {
    while ($AlterRows.Count -lt 5) { Add-AlterRow }
    $map = @( @($row.ID_PORT, 'Service_Port'), @($row.ID_PORT, 'Service_Port'), @($row.ID_PORT, 'Service_Port'),
              @($row.VLAN_NET, 'S-Vlan'), @($row.VLAN_VOIP, 'S-Vlan') )
    for ($i = 0; $i -lt 5; $i++) { $AlterRows[$i].RESOURCE_ID = [string]$map[$i][0]; $AlterRows[$i].CONFIG = $map[$i][1] }
    Refresh-AlterGrid
}

function Fill-AlterFromServices($svc) {
    while ($AlterRows.Count -lt 5) { Add-AlterRow }
    $map = @( @($svc.INTERNET, 'Service_Port'), @($svc.VOICE, 'Service_Port'), @($svc.IPTV, 'Service_Port'),
              @($svc.INTERNET, 'S-Vlan'), @($svc.VOICE, 'S-Vlan') )
    for ($i = 0; $i -lt 5; $i++) { $AlterRows[$i].SERVICE_NAME = [string]$map[$i][0]; $AlterRows[$i].CONFIG = $map[$i][1] }
    Refresh-AlterGrid
}

function Do-Search {
    $ip = $UI.txtIP.Text.Trim(); $sl = $UI.txtSlot.Text.Trim(); $pt = $UI.txtPort.Text.Trim()
    if (-not $ip -and -not $sl -and -not $pt) {
        $UI.gridHasil.ItemsSource = $null
        $UI.lblCount.Text = 'Isi IP / Slot / Port dulu, lalu klik Cari'
        return
    }
    $res = @($Script:Rows | Where-Object {
        ($ip -eq '' -or ([string]$_.IP).Trim() -ceq $ip) -and
        ($sl -eq '' -or (Num-Eq $_.SLOT $sl)) -and
        ($pt -eq '' -or (Num-Eq $_.PORT $pt))
    })
    $UI.gridHasil.ItemsSource = $res
    $UI.lblCount.Text = "Menampilkan $($res.Count) dari $($Script:Rows.Count) data"
    if ($res.Count -gt 0) {
        Fill-AlterFromSearch $res[0]
        $UI.lblCount.Text += ' - Alter terisi otomatis'
    }
}

function Do-ValinsSearch {
    $t = Get-Paste $UI.txtValins $HintValins
    if (-not $t) { return }
    $p = Extract-IpSlotPort $t
    if (-not ($p.IP -or $p.Slot -or $p.Port)) { return }
    if ($p.IP) { $UI.txtIP.Text = $p.IP }
    if ($p.Slot) { $UI.txtSlot.Text = $p.Slot }
    if ($p.Port) { $UI.txtPort.Text = $p.Port }
    Do-Search
}

function Do-BimaExtract {
    $t = Get-Paste $UI.txtBima $HintBima
    if (-not $t) { return }
    $p = Extract-IpSlotPort $t
    if ($p.IP -or $p.Slot -or $p.Port) {
        if ($p.IP) { $UI.txtIP.Text = $p.IP }
        if ($p.Slot) { $UI.txtSlot.Text = $p.Slot }
        if ($p.Port) { $UI.txtPort.Text = $p.Port }
        Do-Search
    }
    $svc = Extract-Services $t
    if ($svc.INTERNET -or $svc.VOICE -or $svc.IPTV) {
        Fill-AlterFromServices $svc
        $info = "IP=$($p.IP) Slot=$($p.Slot) Port=$($p.Port)`r`nINTERNET=$($svc.INTERNET)`r`nVOICE=$($svc.VOICE)`r`nIPTV=$($svc.IPTV)"
        try { [System.Windows.Clipboard]::SetText($info) } catch { }
        $UI.lblCount.Text += ' - Service terisi'
    }
}

$UI.btnCari.Add_Click({ Do-Search })
$UI.btnCariData.Add_Click({ Do-ValinsSearch })
$UI.btnEkstrak.Add_Click({ Do-BimaExtract })
$UI.btnReset.Add_Click({
    $UI.txtIP.Text = ''; $UI.txtSlot.Text = ''; $UI.txtPort.Text = ''
    $UI.gridHasil.ItemsSource = $null
    $UI.lblCount.Text = "$($Script:Rows.Count) data siap - isi IP / Slot / Port lalu klik Cari"
})
$UI.btnRefresh.Add_Click({
    try {
        $Script:Rows = Get-SheetRows
        $UI.lblCount.Text = "$($Script:Rows.Count) data siap - isi IP / Slot / Port lalu klik Cari"
    } catch { $UI.lblCount.Text = 'Gagal refresh - periksa koneksi / Share publik' }
})
$UI.btnAdd.Add_Click({ Add-AlterRow })
$UI.btnClear.Add_Click({
    foreach ($brs in $AlterRows) { $brs.RESOURCE_ID = ''; $brs.SERVICE_NAME = ''; $brs.TARGET_ID = ''; $brs.CONFIG = 'Service_Port' }
    Refresh-AlterGrid
})
$UI.btnDownload.Add_Click({
    $data = @($AlterRows | Where-Object { $_.RESOURCE_ID -or $_.SERVICE_NAME -or $_.TARGET_ID })
    if ($data.Count -eq 0) { return }
    $dl = Join-Path (Get-DownloadFolder) 'Alter.csv'
    $i = 1
    $base = $dl
    while (Test-Path $dl) { $dl = $base -replace '\.csv$', "_$i.csv"; $i++ }
    Write-ProvisionCsv $dl $data
    $UI.lblCount.Text = "Tersimpan $($data.Count) baris -> $dl"
})
$UI.gridHasil.Add_MouseDoubleClick({
    $sel = $UI.gridHasil.SelectedItem
    if ($sel) { try { [System.Windows.Clipboard]::SetText("$($sel.VLAN_NET) | $($sel.VLAN_VOIP) | $($sel.ID_PORT) | $($sel.GPON)") } catch { } }
})
foreach ($tb in @($UI.txtIP, $UI.txtSlot, $UI.txtPort)) {
    $tb.Add_KeyDown({ param($s, $e) if ($e.Key -eq 'Enter') { Do-Search } })
}

# Muat data awal
$UI.lblCount.Text = 'Memuat data dari IP BOGOR...'
try { $Script:Rows = Get-SheetRows } catch { $Script:Rows = Get-FallbackRows }
$UI.lblCount.Text = "$($Script:Rows.Count) data siap - isi IP / Slot / Port lalu klik Cari"

# Pengaman: error di tombol jangan matikan aplikasi, tampilkan di info saja
[System.Windows.Threading.Dispatcher]::CurrentDispatcher.add_UnhandledException({
    param($s, $e)
    try { $UI.lblCount.Text = 'Error: ' + $e.Exception.Message } catch { }
    $e.Handled = $true
})

$win.ShowDialog() | Out-Null
