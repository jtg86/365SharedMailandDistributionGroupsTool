#requires -version 7.0
<#
M365 Mail & Groups Drift Toolkit (V2.6)

Nytt i V2.6:
- Mail-enabled security groups støttes som egen type: "MailSecurityGroup"
- Legg til/fjern medlemmer fungerer for både DistributionGroup og MailSecurityGroup
- Input godtar e-post, UPN og mail nickname/alias (uten @)

Prereqs:
  Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force

Run:
  pwsh -File .\M365-MailGroups-Tool.ps1
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ----------------- Config -----------------
$script:MinSearchLength   = 3
$script:MaxResultsPerType = 200
$script:Cache = @{
  Search  = @{}
  Details = @{}
  Cal     = @{}   # key: "Cal|smtp"
}

# ----------------- Helpers -----------------
function Ensure-Module {
  param([string]$Name)
  if (-not (Get-Module -ListAvailable -Name $Name)) {
    throw "Mangler modul: $Name. Kjør: Install-Module $Name -Scope CurrentUser"
  }
}

function Ensure-Directory {
  param([string]$Path)
  if (-not (Test-Path $Path)) {
    New-Item -ItemType Directory -Path $Path | Out-Null
  }
}

function Get-LogPath {
  $dir = Join-Path -Path $PSScriptRoot -ChildPath "logs"
  Ensure-Directory $dir
  Join-Path $dir ("run-{0}.log" -f (Get-Date -Format "yyyyMMdd-HHmmss"))
}

function Log-Line {
  param(
    [string]$LogFile,
    [string]$Message,
    [string]$Level = "INFO"
  )
  "[$(Get-Date -Format o)] [$Level] $Message" | Out-File -FilePath $LogFile -Append -Encoding utf8
}

function Write-UiStatus {
  param($TextBlock, [string]$Text)
  $TextBlock.Text = $Text
}

function Safe-Invoke {
  param(
    [scriptblock]$Action,
    [scriptblock]$OnError
  )
  try { & $Action } catch { & $OnError $_ }
}

function Escape-FilterValue {
  param([string]$Value)
  ($Value -replace "'", "''")
}

function Join-StringSafe {
  param($Array, [string]$Sep = ", ")
  if ($null -eq $Array) { return "" }
  $vals = @()
  foreach ($x in @($Array)) {
    if ($null -ne $x) { $vals += $x.ToString() }
  }
  ($vals -join $Sep)
}

function Is-MailboxPermissionType {
  param([string]$Type)
  $Type -in @("SharedMailbox","RoomMailbox","EquipmentMailbox")
}

function Is-ResourceMailboxType {
  param([string]$Type)
  $Type -in @("RoomMailbox","EquipmentMailbox")
}

function Is-GroupType {
  param([string]$Type)
  $Type -in @("DistributionGroup","MailSecurityGroup")
}

function Parse-Identities {
  param([string]$Text)

  if ([string]::IsNullOrWhiteSpace($Text)) { return @() }

  # Split on comma/semicolon/whitespace/newline
  $parts = $Text -split '[,\s;]+' |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
    ForEach-Object { $_.Trim() } |
    Select-Object -Unique

  # Accept:
  # - emails/UPN (contains @)
  # - mail nickname/alias: letters/digits + . _ - allowed, length >= 3
  $valid = foreach ($p in $parts) {
    if ($p -match '@') { $p; continue }
    if ($p.Length -ge 3 -and $p -match '^[a-zA-Z0-9][a-zA-Z0-9._-]{1,}$') { $p; continue }
  }

  @($valid | Select-Object -Unique)
}

# ----------------- EXO -----------------
Ensure-Module -Name "ExchangeOnlineManagement"
Import-Module ExchangeOnlineManagement -ErrorAction Stop

function Connect-IfNeeded {
  $ci = $null
  try { $ci = Get-ConnectionInformation -ErrorAction SilentlyContinue } catch {}
  if (-not $ci) {
    Connect-ExchangeOnline -ShowBanner:$false | Out-Null
  }
}

function Try-ResolveRecipient {
  param([string]$Identity)
  try {
    Get-Recipient -Identity $Identity -ErrorAction Stop -WarningAction SilentlyContinue
  } catch {
    $null
  }
}

function Resolve-RecipientInfo {
  param([string]$AnyIdentity)

  if ([string]::IsNullOrWhiteSpace($AnyIdentity)) {
    return [pscustomobject]@{ DisplayName=""; Email=""; TypeDetails="Unknown"; TargetId="" }
  }

  $raw = $AnyIdentity.Trim()

  if ($raw -in @("NT AUTHORITY\SELF","S-1-5-18")) {
    return [pscustomobject]@{ DisplayName=$raw; Email=""; TypeDetails="System"; TargetId=$raw }
  }

  if ($raw -in @("Default","Anonymous")) {
    return [pscustomobject]@{ DisplayName=$raw; Email=""; TypeDetails="Special"; TargetId=$raw }
  }

  $r = Try-ResolveRecipient -Identity $raw
  if ($null -ne $r) {
    $email = ""
    try { $email = $r.PrimarySmtpAddress.ToString() } catch {}
    return [pscustomobject]@{
      DisplayName = $r.DisplayName
      Email       = $email
      TypeDetails = $r.RecipientTypeDetails.ToString()
      TargetId    = $r.Identity.ToString()
    }
  }

  [pscustomobject]@{ DisplayName=$raw; Email=""; TypeDetails="Unresolved"; TargetId=$raw }
}

# ----------------- Search -----------------
function Search-DirectoryObjects {
  param([string]$SearchText)

  Connect-IfNeeded

  $s = $SearchText.Trim()
  if ($s.Length -lt $script:MinSearchLength) { return @() }
  if ($script:Cache.Search.ContainsKey($s)) { return $script:Cache.Search[$s] }

  $esc  = Escape-FilterValue $s
  $like = "*$esc*"

  $sharedFilter = "RecipientTypeDetails -eq 'SharedMailbox' -and (DisplayName -like '$like' -or PrimarySmtpAddress -like '$like' -or Alias -like '$like')"
  $roomFilter   = "RecipientTypeDetails -eq 'RoomMailbox' -and (DisplayName -like '$like' -or PrimarySmtpAddress -like '$like' -or Alias -like '$like')"
  $equipFilter  = "RecipientTypeDetails -eq 'EquipmentMailbox' -and (DisplayName -like '$like' -or PrimarySmtpAddress -like '$like' -or Alias -like '$like')"

  $dgFilter     = "DisplayName -like '$like' -or PrimarySmtpAddress -like '$like' -or Alias -like '$like'"
  $ddgFilter    = "DisplayName -like '$like' -or PrimarySmtpAddress -like '$like' -or Alias -like '$like'"

  $items = New-Object System.Collections.Generic.List[object]

  $shared = @(Get-Mailbox -Filter $sharedFilter -ResultSize $script:MaxResultsPerType -WarningAction SilentlyContinue |
    ForEach-Object { [pscustomobject]@{ Type="SharedMailbox"; Name=$_.DisplayName; Email=$_.PrimarySmtpAddress.ToString(); Identity=$_.Identity } })

  $rooms = @(Get-Mailbox -Filter $roomFilter -ResultSize $script:MaxResultsPerType -WarningAction SilentlyContinue |
    ForEach-Object { [pscustomobject]@{ Type="RoomMailbox"; Name=$_.DisplayName; Email=$_.PrimarySmtpAddress.ToString(); Identity=$_.Identity } })

  $equip = @(Get-Mailbox -Filter $equipFilter -ResultSize $script:MaxResultsPerType -WarningAction SilentlyContinue |
    ForEach-Object { [pscustomobject]@{ Type="EquipmentMailbox"; Name=$_.DisplayName; Email=$_.PrimarySmtpAddress.ToString(); Identity=$_.Identity } })

  # Viktig: Get-DistributionGroup returnerer både vanlige DG og mail-enabled security groups.
  $groups = @(Get-DistributionGroup -Filter $dgFilter -ResultSize $script:MaxResultsPerType -WarningAction SilentlyContinue)

  $dgs = @($groups | ForEach-Object {
      $rtype = $_.RecipientTypeDetails.ToString()
      $isSec = $false
      try {
        # Vanlig indikator: MailUniversalSecurityGroup
        if ($rtype -eq "MailUniversalSecurityGroup") { $isSec = $true }
        # Alternativ: GroupType kan inneholde SecurityEnabled
        if ($_.GroupType -and (@($_.GroupType) -contains "SecurityEnabled")) { $isSec = $true }
      } catch {}

      [pscustomobject]@{
        Type      = $(if ($isSec) { "MailSecurityGroup" } else { "DistributionGroup" })
        Name      = $_.DisplayName
        Email     = $_.PrimarySmtpAddress.ToString()
        Identity  = $_.Identity
        GroupKind = $rtype
      }
    })

  $ddgs = @(Get-DynamicDistributionGroup -Filter $ddgFilter -ResultSize $script:MaxResultsPerType -WarningAction SilentlyContinue |
    ForEach-Object { [pscustomobject]@{ Type="DynamicDistributionGroup"; Name=$_.DisplayName; Email=$_.PrimarySmtpAddress.ToString(); Identity=$_.Identity } })

  foreach ($x in $shared) { $items.Add($x) }
  foreach ($x in $rooms)  { $items.Add($x) }
  foreach ($x in $equip)  { $items.Add($x) }
  foreach ($x in $dgs)    { $items.Add($x) }
  foreach ($x in $ddgs)   { $items.Add($x) }

  $result = $items.ToArray()
  $script:Cache.Search[$s] = $result
  $result
}

# ----------------- Details: Mailbox permissions -----------------
function Get-MailboxAccessState {
  param([string]$MailboxIdentity)

  $faTrustees = @(Get-MailboxPermission -Identity $MailboxIdentity -WarningAction SilentlyContinue |
    Where-Object {
      $_.AccessRights -contains "FullAccess" -and -not $_.IsInherited -and $_.Deny -eq $false -and $_.User.ToString() -ne "NT AUTHORITY\SELF"
    } | ForEach-Object { $_.User.ToString() })

  $saTrustees = @(Get-RecipientPermission -Identity $MailboxIdentity -WarningAction SilentlyContinue |
    Where-Object {
      $_.AccessRights -contains "SendAs" -and $_.Trustee.ToString() -ne "NT AUTHORITY\SELF"
    } | ForEach-Object { $_.Trustee.ToString() })

  $all = @($faTrustees + $saTrustees | Select-Object -Unique)

  $faSet = @{}
  foreach ($t in $faTrustees) { $faSet[$t] = $true }
  $saSet = @{}
  foreach ($t in $saTrustees) { $saSet[$t] = $true }

  $rows = foreach ($t in $all) {
    $info = Resolve-RecipientInfo -AnyIdentity $t
    [pscustomobject]@{
      IsSelected  = $false
      DisplayName = $info.DisplayName
      Email       = $info.Email
      Type        = $info.TypeDetails
      TargetId    = $info.TargetId
      FullAccess  = [bool]$faSet[$t]
      SendAs      = [bool]$saSet[$t]
    }
  }

  $rows | Sort-Object Type, DisplayName
}

# ----------------- Details: Resource calendar permissions -----------------
function Get-CalendarFolderIdentity {
  param([string]$MailboxSmtp)

  $stats = @(Get-MailboxFolderStatistics -Identity $MailboxSmtp -FolderScope Calendar -WarningAction SilentlyContinue)
  if (@($stats).Count -eq 0) { return $null }

  $cal = $stats | Where-Object { $_.FolderType -eq "Calendar" } | Select-Object -First 1
  if (-not $cal) { $cal = $stats | Select-Object -First 1 }

  $fp = $cal.FolderPath
  if ([string]::IsNullOrWhiteSpace($fp)) { return $null }

  $fp2 = "\" + ($fp.TrimStart("/") -replace "/","\")
  "$MailboxSmtp`:$fp2"
}

function Get-ResourceCalendarPermissions {
  param([string]$MailboxSmtp)

  $cacheKey = "Cal|$MailboxSmtp"
  if ($script:Cache.Cal.ContainsKey($cacheKey)) { return $script:Cache.Cal[$cacheKey] }

  $folderId = Get-CalendarFolderIdentity -MailboxSmtp $MailboxSmtp
  if (-not $folderId) {
    $empty = @()
    $script:Cache.Cal[$cacheKey] = $empty
    return $empty
  }

  $perms = @(Get-MailboxFolderPermission -Identity $folderId -WarningAction SilentlyContinue)

  $rows = foreach ($p in $perms) {
    $userRaw = ""
    try { $userRaw = $p.User.ToString() } catch { $userRaw = "" }

    $info = Resolve-RecipientInfo -AnyIdentity $userRaw

    $rightsText = ""
    try { $rightsText = Join-StringSafe -Array $p.AccessRights -Sep ", " } catch { $rightsText = "" }

    $flagsText = ""
    try { $flagsText = $p.SharingPermissionFlags.ToString() } catch { $flagsText = "" }

    [pscustomobject]@{
      DisplayName  = $info.DisplayName
      Email        = $info.Email
      Type         = $info.TypeDetails
      Access       = $rightsText
      SharingFlags = $flagsText
    }
  }

  $result = @($rows | Sort-Object DisplayName)
  $script:Cache.Cal[$cacheKey] = $result
  $result
}

# ----------------- DG/DDG details -----------------
function Get-GroupMembers {
  param([string]$GroupIdentity)

  Get-DistributionGroupMember -Identity $GroupIdentity -ResultSize Unlimited -WarningAction SilentlyContinue |
    Select-Object @{n="IsSelected";e={$false}},
                  @{n="Name";e={$_.DisplayName}},
                  @{n="Email";e={$_.PrimarySmtpAddress.ToString()}},
                  @{n="Type";e={$_.RecipientType}}
}

function Get-DDGRule {
  param([string]$GroupIdentity)
  $g = Get-DynamicDistributionGroup -Identity $GroupIdentity -WarningAction SilentlyContinue
  [pscustomobject]@{
    Name               = $g.DisplayName
    Email              = $g.PrimarySmtpAddress.ToString()
    RecipientFilter    = $g.RecipientFilter
    RecipientContainer = $g.RecipientContainer
  }
}

function Get-DetailsCached {
  param([pscustomobject]$Obj)

  $key = "{0}|{1}" -f $Obj.Type, $Obj.Identity
  if ($script:Cache.Details.ContainsKey($key)) { return $script:Cache.Details[$key] }

  $details = switch ($Obj.Type) {
    "SharedMailbox" {
      [pscustomobject]@{ Type="SharedMailbox"; Header="Shared mailbox: $($Obj.Name)  <$($Obj.Email)>"; Perms=@(Get-MailboxAccessState -MailboxIdentity $Obj.Identity) }
    }
    "RoomMailbox" {
      [pscustomobject]@{ Type="RoomMailbox"; Header="Møterom (RoomMailbox): $($Obj.Name)  <$($Obj.Email)>"; Perms=@(Get-MailboxAccessState -MailboxIdentity $Obj.Identity) }
    }
    "EquipmentMailbox" {
      [pscustomobject]@{ Type="EquipmentMailbox"; Header="Ressurs (EquipmentMailbox): $($Obj.Name)  <$($Obj.Email)>"; Perms=@(Get-MailboxAccessState -MailboxIdentity $Obj.Identity) }
    }

    # Begge gruppetyper bruker samme medlems-API
    "DistributionGroup" {
      [pscustomobject]@{ Type="DistributionGroup"; Header="Distribusjonsgruppe: $($Obj.Name)  <$($Obj.Email)>"; Members=@(Get-GroupMembers -GroupIdentity $Obj.Identity) }
    }
    "MailSecurityGroup" {
      [pscustomobject]@{ Type="MailSecurityGroup"; Header="Mail-enabled sikkerhetsgruppe: $($Obj.Name)  <$($Obj.Email)>"; Members=@(Get-GroupMembers -GroupIdentity $Obj.Identity) }
    }

    "DynamicDistributionGroup" {
      $r = Get-DDGRule -GroupIdentity $Obj.Identity
      [pscustomobject]@{
        Type               = "DynamicDistributionGroup"
        Header             = "Dynamisk DG: $($r.Name)  <$($r.Email)>"
        RecipientFilter    = $r.RecipientFilter
        RecipientContainer = $r.RecipientContainer
      }
    }
  }

  $script:Cache.Details[$key] = $details
  $details
}

function Invalidate-DetailsCache {
  param([pscustomobject]$Obj)
  $key = "{0}|{1}" -f $Obj.Type, $Obj.Identity
  [void]$script:Cache.Details.Remove($key)
  if ($Obj -and (Is-ResourceMailboxType -Type $Obj.Type)) {
    [void]$script:Cache.Cal.Remove("Cal|$($Obj.Email)")
  }
}

# ----------------- Mutations -----------------
function Add-GroupMembers {
  param([string]$GroupIdentity, [string[]]$Users, [string]$LogFile)

  foreach ($u in @($Users)) {
    $rec = Try-ResolveRecipient -Identity $u
    if (-not $rec) { Log-Line $LogFile "Fant ikke mottaker: $u" "WARN"; continue }

    try {
      Add-DistributionGroupMember -Identity $GroupIdentity -Member $rec.Identity -Confirm:$false -WarningAction SilentlyContinue | Out-Null
      Log-Line $LogFile "Group: Add member -> $GroupIdentity : $($rec.Identity)"
    } catch {
      Log-Line $LogFile "Group: Add (mulig allerede medlem / blokkert) -> $GroupIdentity : $($rec.Identity) : $($_.Exception.Message)" "WARN"
    }
  }
}

function Remove-GroupMembers {
  param([string]$GroupIdentity, [string[]]$Users, [string]$LogFile)

  foreach ($u in @($Users)) {
    $rec = Try-ResolveRecipient -Identity $u
    if (-not $rec) { Log-Line $LogFile "Fant ikke mottaker: $u" "WARN"; continue }

    try {
      Remove-DistributionGroupMember -Identity $GroupIdentity -Member $rec.Identity -Confirm:$false -WarningAction SilentlyContinue | Out-Null
      Log-Line $LogFile "Group: Remove member -> $GroupIdentity : $($rec.Identity)"
    } catch {
      Log-Line $LogFile "Group: Remove (mulig ikke medlem / blokkert) -> $GroupIdentity : $($rec.Identity) : $($_.Exception.Message)" "WARN"
    }
  }
}

function Grant-MailboxRights {
  param(
    [string]$MailboxIdentity,
    [string[]]$Users,
    [bool]$DoFullAccess,
    [bool]$DoSendAs,
    [string]$LogFile
  )

  foreach ($u in @($Users)) {
    $rec = Try-ResolveRecipient -Identity $u
    if (-not $rec) { Log-Line $LogFile "Fant ikke mottaker: $u" "WARN"; continue }

    if ($DoFullAccess) {
      try {
        Add-MailboxPermission -Identity $MailboxIdentity -User $rec.Identity -AccessRights FullAccess -InheritanceType All -AutoMapping:$true -WarningAction SilentlyContinue | Out-Null
        Log-Line $LogFile "Mailbox: Grant FullAccess -> $MailboxIdentity : $($rec.Identity)"
      } catch {
        Log-Line $LogFile "Mailbox: FullAccess (mulig allerede satt) -> $MailboxIdentity : $($rec.Identity) : $($_.Exception.Message)" "WARN"
      }
    }

    if ($DoSendAs) {
      try {
        Add-RecipientPermission -Identity $MailboxIdentity -Trustee $rec.Identity -AccessRights SendAs -Confirm:$false -WarningAction SilentlyContinue | Out-Null
        Log-Line $LogFile "Mailbox: Grant SendAs -> $MailboxIdentity : $($rec.Identity)"
      } catch {
        Log-Line $LogFile "Mailbox: SendAs (mulig allerede satt) -> $MailboxIdentity : $($rec.Identity) : $($_.Exception.Message)" "WARN"
      }
    }
  }
}

function Revoke-MailboxRights {
  param(
    [string]$MailboxIdentity,
    [string[]]$UsersOrGroups,
    [bool]$DoFullAccess,
    [bool]$DoSendAs,
    [string]$LogFile
  )

  foreach ($u in @($UsersOrGroups)) {
    $rec = Try-ResolveRecipient -Identity $u
    $target = if ($rec) { $rec.Identity } else { $u }

    if ($DoFullAccess) {
      try {
        Remove-MailboxPermission -Identity $MailboxIdentity -User $target -AccessRights FullAccess -InheritanceType All -Confirm:$false -WarningAction SilentlyContinue | Out-Null
        Log-Line $LogFile "Mailbox: Revoke FullAccess -> $MailboxIdentity : $target"
      } catch {
        Log-Line $LogFile "Mailbox: FullAccess (mulig allerede fjernet) -> $MailboxIdentity : $target : $($_.Exception.Message)" "WARN"
      }
    }

    if ($DoSendAs) {
      try {
        Remove-RecipientPermission -Identity $MailboxIdentity -Trustee $target -AccessRights SendAs -Confirm:$false -WarningAction SilentlyContinue | Out-Null
        Log-Line $LogFile "Mailbox: Revoke SendAs -> $MailboxIdentity : $target"
      } catch {
        Log-Line $LogFile "Mailbox: SendAs (mulig allerede fjernet) -> $MailboxIdentity : $target : $($_.Exception.Message)" "WARN"
      }
    }
  }
}

# ----------------- WPF UI -----------------
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="M365 Mail &amp; Groups Drift Toolkit (V2.6)" Height="780" Width="1220" WindowStartupLocation="CenterScreen">
  <Grid Margin="10">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <DockPanel Grid.Row="0" Margin="0,0,0,10">
      <TextBox Name="TbSearch" Width="460" Height="28" Margin="0,0,10,0" VerticalContentAlignment="Center"
               ToolTip="Søk etter navn/alias/e-post. Minst 3 tegn." />
      <Button Name="BtnSearch" Content="Søk" Width="90" Height="28" Margin="0,0,10,0"/>
      <Button Name="BtnConnect" Content="Koble til EXO" Width="120" Height="28" Margin="0,0,10,0"/>
      <Button Name="BtnClearCache" Content="Tøm cache" Width="100" Height="28"/>
    </DockPanel>

    <Grid Grid.Row="1">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="470"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>

      <GroupBox Header="Treff (inkl. møterom/ressurser/grupper)" Grid.Column="0" Margin="0,0,10,0">
        <DataGrid Name="GridObjects" AutoGenerateColumns="False" IsReadOnly="True" SelectionMode="Single">
          <DataGrid.Columns>
            <DataGridTextColumn Header="Type" Binding="{Binding Type}" Width="190"/>
            <DataGridTextColumn Header="Navn" Binding="{Binding Name}" Width="*"/>
            <DataGridTextColumn Header="E-post" Binding="{Binding Email}" Width="*"/>
          </DataGrid.Columns>
        </DataGrid>
      </GroupBox>

      <GroupBox Header="Detaljer (load on demand)" Grid.Column="1">
        <TabControl Name="Tabs">

          <TabItem Name="TabMailboxPerms" Header="Mailbox permissions (FullAccess/SendAs)">
            <Grid Margin="10">
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
              </Grid.RowDefinitions>

              <TextBlock Name="TbMailboxHeader" FontSize="14" FontWeight="SemiBold" />

              <DataGrid Name="GridMailboxPerms" Grid.Row="1" AutoGenerateColumns="False" CanUserAddRows="False" SelectionMode="Extended">
                <DataGrid.Columns>
                  <DataGridCheckBoxColumn Header="Velg" Binding="{Binding IsSelected, Mode=TwoWay}" Width="60"/>
                  <DataGridTextColumn Header="Navn" Binding="{Binding DisplayName}" Width="*"/>
                  <DataGridTextColumn Header="E-post" Binding="{Binding Email}" Width="*"/>
                  <DataGridTextColumn Header="Type" Binding="{Binding Type}" Width="210"/>
                  <DataGridCheckBoxColumn Header="FullAccess" Binding="{Binding FullAccess}" Width="90" IsReadOnly="True"/>
                  <DataGridCheckBoxColumn Header="SendAs" Binding="{Binding SendAs}" Width="80" IsReadOnly="True"/>
                </DataGrid.Columns>
              </DataGrid>

              <StackPanel Grid.Row="2" Margin="0,10,0,0">
                <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                  <CheckBox Name="CbFullAccess" Content="FullAccess" Margin="0,0,15,0" IsChecked="True"/>
                  <CheckBox Name="CbSendAs" Content="SendAs" Margin="0,0,15,0" IsChecked="True"/>
                  <Button Name="BtnMailboxRemoveSelected" Content="Fjern valgte" Width="120" Margin="0,0,10,0"/>
                </StackPanel>

                <TextBox Name="TbMailboxAdd" Height="60" TextWrapping="Wrap" AcceptsReturn="True"
                         VerticalScrollBarVisibility="Auto"
                         ToolTip="Lim inn e-post/UPN/alias (komma/mellomrom/linjeskift). Grupper støttes også."/>
                <StackPanel Orientation="Horizontal" Margin="0,8,0,0">
                  <Button Name="BtnMailboxGrant" Content="Tildel valgt(e) rettighet(er)" Width="220" />
                </StackPanel>
              </StackPanel>
            </Grid>
          </TabItem>

          <TabItem Name="TabResourceCalendar" Header="Ressurskalender (Room/Equipment)">
            <Grid Margin="10">
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>

              <TextBlock Name="TbCalHeader" FontSize="14" FontWeight="SemiBold" />

              <DataGrid Name="GridCalPerms" Grid.Row="1" AutoGenerateColumns="False" CanUserAddRows="False" IsReadOnly="True">
                <DataGrid.Columns>
                  <DataGridTextColumn Header="Navn" Binding="{Binding DisplayName}" Width="*"/>
                  <DataGridTextColumn Header="E-post" Binding="{Binding Email}" Width="*"/>
                  <DataGridTextColumn Header="Type" Binding="{Binding Type}" Width="210"/>
                  <DataGridTextColumn Header="Access" Binding="{Binding Access}" Width="220"/>
                  <DataGridTextColumn Header="Flags" Binding="{Binding SharingFlags}" Width="160"/>
                </DataGrid.Columns>
              </DataGrid>
            </Grid>
          </TabItem>

          <TabItem Name="TabGroup" Header="Gruppe (DG / Mail-sikkerhet)">
            <Grid Margin="10">
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
              </Grid.RowDefinitions>

              <TextBlock Name="TbGroupHeader" FontSize="14" FontWeight="SemiBold" />

              <DataGrid Name="GridGroupMembers" Grid.Row="1" AutoGenerateColumns="False" CanUserAddRows="False">
                <DataGrid.Columns>
                  <DataGridCheckBoxColumn Header="Velg" Binding="{Binding IsSelected, Mode=TwoWay}" Width="60"/>
                  <DataGridTextColumn Header="Navn" Binding="{Binding Name}" Width="*"/>
                  <DataGridTextColumn Header="E-post" Binding="{Binding Email}" Width="*"/>
                  <DataGridTextColumn Header="Type" Binding="{Binding Type}" Width="170"/>
                </DataGrid.Columns>
              </DataGrid>

              <StackPanel Grid.Row="2" Margin="0,10,0,0">
                <Button Name="BtnGroupRemoveSelected" Content="Fjern valgte" Width="120" Margin="0,0,0,8"/>
                <TextBox Name="TbGroupAdd" Height="60" TextWrapping="Wrap" AcceptsReturn="True"
                         VerticalScrollBarVisibility="Auto"
                         ToolTip="Lim inn e-post/UPN/alias (komma/mellomrom/linjeskift)"/>
                <Button Name="BtnGroupAdd" Content="Legg til" Width="120" Margin="0,8,0,0"/>
              </StackPanel>
            </Grid>
          </TabItem>

          <TabItem Name="TabDDG" Header="Dynamisk distribusjonsgruppe">
            <Grid Margin="10">
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>

              <TextBlock Name="TbDdgHeader" FontSize="14" FontWeight="SemiBold" />
              <TextBlock Grid.Row="1" Text="Dette er en dynamisk gruppe. Medlemskap styres av regel og kan ikke endres manuelt."
                         Margin="0,8,0,8" FontWeight="SemiBold"/>

              <StackPanel Grid.Row="2">
                <TextBlock Text="RecipientFilter:" FontWeight="SemiBold"/>
                <TextBox Name="TbRecipientFilter" Height="150" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"/>
                <TextBlock Text="RecipientContainer:" FontWeight="SemiBold" Margin="0,10,0,0"/>
                <TextBox Name="TbRecipientContainer" Height="60" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"/>
              </StackPanel>
            </Grid>
          </TabItem>

        </TabControl>
      </GroupBox>
    </Grid>

    <TextBlock Name="TbStatus" Grid.Row="2" Margin="0,10,0,0" />
  </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
$win = [Windows.Markup.XamlReader]::Load($reader)

# Controls
$TbSearch      = $win.FindName("TbSearch")
$BtnSearch     = $win.FindName("BtnSearch")
$BtnConnect    = $win.FindName("BtnConnect")
$BtnClearCache = $win.FindName("BtnClearCache")
$GridObjects   = $win.FindName("GridObjects")
$Tabs          = $win.FindName("Tabs")

$TabMailboxPerms     = $win.FindName("TabMailboxPerms")
$TabResourceCalendar = $win.FindName("TabResourceCalendar")
$TabGroup            = $win.FindName("TabGroup")
$TabDDG              = $win.FindName("TabDDG")

$TbMailboxHeader          = $win.FindName("TbMailboxHeader")
$GridMailboxPerms         = $win.FindName("GridMailboxPerms")
$CbFullAccess             = $win.FindName("CbFullAccess")
$CbSendAs                 = $win.FindName("CbSendAs")
$BtnMailboxRemoveSelected = $win.FindName("BtnMailboxRemoveSelected")
$TbMailboxAdd             = $win.FindName("TbMailboxAdd")
$BtnMailboxGrant          = $win.FindName("BtnMailboxGrant")

$TbCalHeader  = $win.FindName("TbCalHeader")
$GridCalPerms = $win.FindName("GridCalPerms")

$TbGroupHeader         = $win.FindName("TbGroupHeader")
$GridGroupMembers      = $win.FindName("GridGroupMembers")
$BtnGroupRemoveSelected = $win.FindName("BtnGroupRemoveSelected")
$TbGroupAdd            = $win.FindName("TbGroupAdd")
$BtnGroupAdd           = $win.FindName("BtnGroupAdd")

$TbDdgHeader          = $win.FindName("TbDdgHeader")
$TbRecipientFilter    = $win.FindName("TbRecipientFilter")
$TbRecipientContainer = $win.FindName("TbRecipientContainer")

$TbStatus = $win.FindName("TbStatus")

# State
$script:SelectedObject = $null
$logFile = Get-LogPath
Log-Line $logFile "App started"
Write-UiStatus $TbStatus "Klar. Skriv minst $($script:MinSearchLength) tegn og trykk Søk. Logg: $logFile"

function Clear-Details {
  $TbMailboxHeader.Text = ""
  $GridMailboxPerms.ItemsSource = $null
  $TbCalHeader.Text = ""
  $GridCalPerms.ItemsSource = $null
  $TbGroupHeader.Text = ""
  $GridGroupMembers.ItemsSource = $null
  $TbDdgHeader.Text = ""
  $TbRecipientFilter.Text = ""
  $TbRecipientContainer.Text = ""
}

function Set-ButtonsEnabledForType {
  param([string]$Type)

  $BtnMailboxGrant.IsEnabled = $false
  $BtnMailboxRemoveSelected.IsEnabled = $false
  $BtnGroupAdd.IsEnabled = $false
  $BtnGroupRemoveSelected.IsEnabled = $false

  $TabResourceCalendar.IsEnabled = $false

  if (Is-MailboxPermissionType -Type $Type) {
    $BtnMailboxGrant.IsEnabled = $true
    $BtnMailboxRemoveSelected.IsEnabled = $true
  }

  if (Is-GroupType -Type $Type) {
    $BtnGroupAdd.IsEnabled = $true
    $BtnGroupRemoveSelected.IsEnabled = $true
  }

  if (Is-ResourceMailboxType -Type $Type) {
    $TabResourceCalendar.IsEnabled = $true
  }
}

function Load-DetailsForSelection {
  param($obj)

  Clear-Details
  $script:SelectedObject = $obj
  if (-not $obj) { return }

  Safe-Invoke -Action {
    Write-UiStatus $TbStatus "Laster detaljer for $($obj.Type): $($obj.Email)..."
    Set-ButtonsEnabledForType -Type $obj.Type

    $details = Get-DetailsCached -Obj $obj

    if (Is-MailboxPermissionType -Type $obj.Type) {
      if (Is-ResourceMailboxType -Type $obj.Type) { $Tabs.SelectedItem = $TabResourceCalendar }
      else { $Tabs.SelectedItem = $TabMailboxPerms }

      $TbMailboxHeader.Text = $details.Header
      $GridMailboxPerms.ItemsSource = $details.Perms

      if (Is-ResourceMailboxType -Type $obj.Type) {
        $TbCalHeader.Text = "Kalenderrettigheter: $($obj.Name)  <$($obj.Email)>"
        $GridCalPerms.ItemsSource = @(Get-ResourceCalendarPermissions -MailboxSmtp $obj.Email)
      }
    }
    elseif (Is-GroupType -Type $obj.Type) {
      $Tabs.SelectedItem = $TabGroup
      $TbGroupHeader.Text = $details.Header
      $GridGroupMembers.ItemsSource = $details.Members
    }
    elseif ($obj.Type -eq "DynamicDistributionGroup") {
      $Tabs.SelectedItem = $TabDDG
      $TbDdgHeader.Text = $details.Header
      $TbRecipientFilter.Text = $details.RecipientFilter
      $TbRecipientContainer.Text = $details.RecipientContainer
    }

    Write-UiStatus $TbStatus "Klar."
  } -OnError {
    param($err)
    Write-UiStatus $TbStatus "Feil ved lasting av detaljer: $($err.Exception.Message)"
    Log-Line $logFile "ERROR details: $($err.Exception.Message)" "ERROR"
  }
}

function Do-Search {
  $s = $TbSearch.Text.Trim()
  if ($s.Length -lt $script:MinSearchLength) {
    Write-UiStatus $TbStatus "Skriv minst $($script:MinSearchLength) tegn for å søke."
    return
  }

  Safe-Invoke -Action {
    Write-UiStatus $TbStatus "Søker etter '$s'..."
    Log-Line $logFile "Search: $s"
    $results = @(Search-DirectoryObjects -SearchText $s)
    $GridObjects.ItemsSource = $results
    if (@($results).Count -eq 0) { Write-UiStatus $TbStatus "Ingen treff." }
    else { Write-UiStatus $TbStatus "Fant $(@($results).Count) treff." }
  } -OnError {
    param($err)
    Write-UiStatus $TbStatus "Feil ved søk: $($err.Exception.Message)"
    Log-Line $logFile "ERROR search: $($err.Exception.Message)" "ERROR"
  }
}

# Events
$BtnConnect.Add_Click({
  Safe-Invoke -Action {
    Write-UiStatus $TbStatus "Kobler til Exchange Online..."
    Connect-ExchangeOnline -ShowBanner:$false | Out-Null
    Write-UiStatus $TbStatus "Tilkoblet."
    Log-Line $logFile "Connected EXO"
  } -OnError {
    param($err)
    Write-UiStatus $TbStatus "Kunne ikke koble til: $($err.Exception.Message)"
    Log-Line $logFile "ERROR connect: $($err.Exception.Message)" "ERROR"
  }
})

$BtnSearch.Add_Click({ Do-Search })
$TbSearch.Add_KeyDown({ if ($_.Key -eq "Enter") { Do-Search } })

$BtnClearCache.Add_Click({
  $script:Cache.Search.Clear()
  $script:Cache.Details.Clear()
  $script:Cache.Cal.Clear()
  Write-UiStatus $TbStatus "Cache tømt."
  Log-Line $logFile "Cache cleared"
})

$GridObjects.Add_SelectionChanged({
  $sel = $GridObjects.SelectedItem
  Load-DetailsForSelection -obj $sel
})

# Group actions (DG + MailSecurityGroup)
$BtnGroupAdd.Add_Click({
  if (-not $script:SelectedObject -or -not (Is-GroupType -Type $script:SelectedObject.Type)) { return }

  Safe-Invoke -Action {
    $users = @(Parse-Identities -Text $TbGroupAdd.Text)
    if (@($users).Count -eq 0) { Write-UiStatus $TbStatus "Ingen gyldige adresser/alias funnet i input."; return }

    Write-UiStatus $TbStatus "Legger til medlemmer..."
    Log-Line $logFile "Group Add: $($script:SelectedObject.Type) $($script:SelectedObject.Email) users=$($users -join ';')"

    Add-GroupMembers -GroupIdentity $script:SelectedObject.Identity -Users $users -LogFile $logFile

    $TbGroupAdd.Text = ""
    Invalidate-DetailsCache -Obj $script:SelectedObject
    Load-DetailsForSelection -obj $script:SelectedObject
    Write-UiStatus $TbStatus "Ferdig: lagt til."
  } -OnError {
    param($err)
    Write-UiStatus $TbStatus "Feil ved add: $($err.Exception.Message)"
    Log-Line $logFile "ERROR group add: $($err.Exception.Message)" "ERROR"
  }
})

$BtnGroupRemoveSelected.Add_Click({
  if (-not $script:SelectedObject -or -not (Is-GroupType -Type $script:SelectedObject.Type)) { return }

  Safe-Invoke -Action {
    $items = @($GridGroupMembers.ItemsSource) | Where-Object { $_.IsSelected -eq $true }
    if (@($items).Count -eq 0) { Write-UiStatus $TbStatus "Velg minst én medlem i lista."; return }

    $users = @($items.Email)

    Write-UiStatus $TbStatus "Fjerner medlemmer..."
    Log-Line $logFile "Group Remove: $($script:SelectedObject.Type) $($script:SelectedObject.Email) users=$($users -join ';')"

    Remove-GroupMembers -GroupIdentity $script:SelectedObject.Identity -Users $users -LogFile $logFile

    Invalidate-DetailsCache -Obj $script:SelectedObject
    Load-DetailsForSelection -obj $script:SelectedObject
    Write-UiStatus $TbStatus "Ferdig: fjernet."
  } -OnError {
    param($err)
    Write-UiStatus $TbStatus "Feil ved remove: $($err.Exception.Message)"
    Log-Line $logFile "ERROR group remove: $($err.Exception.Message)" "ERROR"
  }
})

# Mailbox actions
$BtnMailboxGrant.Add_Click({
  if (-not $script:SelectedObject -or -not (Is-MailboxPermissionType -Type $script:SelectedObject.Type)) { return }

  Safe-Invoke -Action {
    $users = @(Parse-Identities -Text $TbMailboxAdd.Text)
    if (@($users).Count -eq 0) { Write-UiStatus $TbStatus "Ingen gyldige adresser/alias funnet i input."; return }

    $doFA = [bool]$CbFullAccess.IsChecked
    $doSA = [bool]$CbSendAs.IsChecked
    if (-not $doFA -and -not $doSA) { Write-UiStatus $TbStatus "Huk av minst én rettighet (FullAccess/SendAs)."; return }

    Write-UiStatus $TbStatus "Tildeler rettigheter..."
    Log-Line $logFile "Mailbox Grant: $($script:SelectedObject.Type) $($script:SelectedObject.Email) users=$($users -join ';') FA=$doFA SA=$doSA"

    Grant-MailboxRights -MailboxIdentity $script:SelectedObject.Identity -Users $users -DoFullAccess:$doFA -DoSendAs:$doSA -LogFile $logFile

    $TbMailboxAdd.Text = ""
    Invalidate-DetailsCache -Obj $script:SelectedObject
    Load-DetailsForSelection -obj $script:SelectedObject
    Write-UiStatus $TbStatus "Ferdig: tildelt."
  } -OnError {
    param($err)
    Write-UiStatus $TbStatus "Feil ved tildeling: $($err.Exception.Message)"
    Log-Line $logFile "ERROR mailbox grant: $($err.Exception.Message)" "ERROR"
  }
})

$BtnMailboxRemoveSelected.Add_Click({
  if (-not $script:SelectedObject -or -not (Is-MailboxPermissionType -Type $script:SelectedObject.Type)) { return }

  Safe-Invoke -Action {
    $items = @($GridMailboxPerms.ItemsSource) | Where-Object { $_.IsSelected -eq $true }
    if (@($items).Count -eq 0) { Write-UiStatus $TbStatus "Velg minst én i lista."; return }

    $doFA = [bool]$CbFullAccess.IsChecked
    $doSA = [bool]$CbSendAs.IsChecked
    if (-not $doFA -and -not $doSA) { Write-UiStatus $TbStatus "Huk av minst én rettighet (FullAccess/SendAs)."; return }

    $targets = @($items.TargetId) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    Write-UiStatus $TbStatus "Fjerner rettigheter..."
    Log-Line $logFile "Mailbox Revoke: $($script:SelectedObject.Type) $($script:SelectedObject.Email) targets=$($targets -join ';') FA=$doFA SA=$doSA"

    Revoke-MailboxRights -MailboxIdentity $script:SelectedObject.Identity -UsersOrGroups $targets -DoFullAccess:$doFA -DoSendAs:$doSA -LogFile $logFile

    Invalidate-DetailsCache -Obj $script:SelectedObject
    Load-DetailsForSelection -obj $script:SelectedObject
    Write-UiStatus $TbStatus "Ferdig: fjernet."
  } -OnError {
    param($err)
    Write-UiStatus $TbStatus "Feil ved fjerning: $($err.Exception.Message)"
    Log-Line $logFile "ERROR mailbox revoke: $($err.Exception.Message)" "ERROR"
  }
})

# Show window
$win.ShowDialog() | Out-Null

# Cleanup
try { Disconnect-ExchangeOnline -Confirm:$false | Out-Null } catch {}
