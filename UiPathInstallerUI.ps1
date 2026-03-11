# =====================================================================
# UiPath Installer UI (PowerShell + WPF)
# Two screens: (1) Download & Offline validation, (2) Installation & Script
# - No install option checkboxes on Screen 1
# - Live popup log while installing, with Cancel + Close
# =====================================================================

# ---------- Ensure STA + elevate if needed ----------
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
  Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -STA -File `"$PSCommandPath`"" -Verb RunAs
  exit
}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# ---------- Load XAML ----------
$xamlPath = Join-Path $PSScriptRoot "UiPathInstallerUI.xaml"
if (-not (Test-Path $xamlPath)) {
  [System.Windows.MessageBox]::Show("Cannot find XAML: $xamlPath","Error") | Out-Null
  exit
}

[xml]$xaml = Get-Content $xamlPath -Raw
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

# ---------- Bind controls ----------
$Screen1Grid = $window.FindName("Screen1Grid")
$ProductBox = $window.FindName("ProductBox")
$VersionBox = $window.FindName("VersionBox")
$DownloadTargetBox = $window.FindName("DownloadTargetBox")
$BrowseDownloadTargetBtn = $window.FindName("BrowseDownloadTargetBtn")
$DownloadBtn = $window.FindName("DownloadBtn")
$DownloadProgress = $window.FindName("DownloadProgress")
$DownloadStatus = $window.FindName("DownloadStatus")

$InstallerFolderBox = $window.FindName("InstallerFolderBox")
$BrowseInstallerFolderBtn = $window.FindName("BrowseInstallerFolderBtn")
$ExpectedShaBox = $window.FindName("ExpectedShaBox")
$ValidateInstallerBtn = $window.FindName("ValidateInstallerBtn")
$InstallerValidationResult = $window.FindName("InstallerValidationResult")
$NextToInstallBtn = $window.FindName("NextToInstallBtn")

$Screen2Grid = $window.FindName("Screen2Grid")
$BackToScreen1Btn = $window.FindName("BackToScreen1Btn")
$ServiceModeCheck2 = $window.FindName("ServiceModeCheck2")
$AllUsersCheck2 = $window.FindName("AllUsersCheck2")
$RobotInstallCheck2 = $window.FindName("RobotInstallCheck2")
$StudioInstallCheck2 = $window.FindName("StudioInstallCheck2")

$ScriptOutputBox = $window.FindName("ScriptOutputBox")
$GenerateScriptBtn = $window.FindName("GenerateScriptBtn")
$CopyScriptBtn = $window.FindName("CopyScriptBtn")
$SaveScriptBtn = $window.FindName("SaveScriptBtn")
$RunInstallerBtn = $window.FindName("RunInstallerBtn")

# ---------- State ----------
$global:Validated = $false
$global:InstallerRoot = $null
$global:ProductMetadata = $null

# ---------- Helpers ----------
function Select-Folder($Initial="") {
  $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
  if ($Initial) { $dlg.SelectedPath = $Initial }
  if ($dlg.ShowDialog() -eq "OK") { return $dlg.SelectedPath }
  return ""
}

function Show-Error($m) { [System.Windows.MessageBox]::Show($m, "Error") | Out-Null }
function Show-Info ($m) { [System.Windows.MessageBox]::Show($m, "Information") | Out-Null }

function Test-Online {
  try {
    $req = [System.Net.WebRequest]::Create("https://www.microsoft.com")
    $req.Method = "HEAD"; $req.Timeout = 3000
    $resp = $req.GetResponse(); $resp.Close()
    return $true
  } catch { return $false }
}

function Validate-InstallerFiles($folder,$sha) {
  if (-not (Test-Path $folder)) { return "Folder does not exist." }
  $msi = Join-Path $folder "UiPathStudio.msi"
  if (-not (Test-Path $msi)) { return "UiPathStudio.msi NOT found." }
  if ($sha) {
    try {
      $hash=(Get-FileHash -Algorithm SHA256 $msi).Hash.ToLower()
      if ($hash -ne $sha.ToLower()) { return "SHA256 mismatch. Actual: $hash" }
    } catch { return "Could not compute SHA256: $($_.Exception.Message)" }
  }
  return "Installer OK"
}

function Test-StudioInstalled {
  if (Test-Path "C:\Program Files\UiPath\Studio") { return $true }
  $regPaths=@(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
  )
  foreach ($rp in $regPaths) {
    if (Test-Path $rp) {
      foreach ($it in Get-ChildItem $rp) {
        $dn=(Get-ItemProperty $it.PSPath -ErrorAction SilentlyContinue).DisplayName
        if ($dn -like "*UiPath Studio*") { return $true }
      }
    }
  }
  return $false
}

# ---------- Product metadata (optional) ----------
# Expecting a JSON structure like:
# {
#   "UiPath": {
#     "23.10.10": { "urls": ["https://.../UiPathStudio.msi"] },
#     "24.2.1":   { "urls": ["https://.../UiPathStudio.msi"] }
#   }
# }
$jsonPath = Join-Path $PSScriptRoot "json_files\product_versions.json"
if (Test-Path $jsonPath) {
  try {
    $global:ProductMetadata = Get-Content $jsonPath -Raw | ConvertFrom-Json
  } catch {
    $global:ProductMetadata = $null
  }
}

# ---------- Populate Product & Version comboboxes ----------
if ($global:ProductMetadata) {
  $ProductBox.Items.Clear()
  $global:ProductMetadata.PSObject.Properties.Name | ForEach-Object { [void]$ProductBox.Items.Add($_) }

  $ProductBox.Add_SelectionChanged({
    $VersionBox.Items.Clear()
    $sel = $ProductBox.SelectedItem
    if ($null -ne $sel) {
      $versions = $global:ProductMetadata.$sel.PSObject.Properties.Name
      foreach ($v in $versions) { [void]$VersionBox.Items.Add($v) }
    }
  })
}

# ---------- Build MSI command from Screen 2 ----------
function Get-SelectedFeatures {
  $features = [System.Collections.Generic.List[string]]::new()
  # Robot is typically required by UiPath robot/service — include if selected
  if ($RobotInstallCheck2.IsChecked) { $features.Add("Robot") }
  if ($StudioInstallCheck2.IsChecked) { $features.Add("Studio") }
  if ($ServiceModeCheck2.IsChecked) { $features.Add("RegisterService") }
  return $features
}

function Build-MSICommand {
  $msi = Join-Path $global:InstallerRoot "UiPathStudio.msi"
  $allUsers = if ($AllUsersCheck2.IsChecked) { "ALLUSERS=1" } else { "ALLUSERS=0" }

  $features = Get-SelectedFeatures
  $addlocal = ""
  if ($features.Count -gt 0) { $addlocal = " ADDLOCAL=" + ($features -join ",") }

  # If Studio unchecked but detected as installed => remove Studio
  $remove = ""
  if (-not $StudioInstallCheck2.IsChecked -and (Test-StudioInstalled)) {
    $remove = " REMOVE=Studio"
  }

  # Quiet + no restart; append MSI logging when running
  return "msiexec /i `"$msi`" /qn /norestart $allUsers$addlocal$remove"
}

function Build-FullScript {
  if (-not $global:InstallerRoot) { return "# ERROR: Missing installer folder" }
  $msiCmd = Build-MSICommand
  $msiLog = '$env:TEMP\UiPathStudio_{0}.log' -f ([Guid]::NewGuid())
  @"
# Install UiPath (silent)
$msi = Join-Path '$($global:InstallerRoot)' 'UiPathStudio.msi'
if (-not (Test-Path $msi)) { throw "UiPathStudio.msi not found in '$($global:InstallerRoot)'" }

$cmd = '$msiCmd /L*V "' + $msiLog + '"'
Write-Host "Running: $cmd"
cmd.exe /c $cmd
if (`$LASTEXITCODE -ne 0) { throw "Installer exited with code `$LASTEXITCODE. See: $msiLog" }
Write-Host "Install complete. Log: $msiLog"
"@
}

# ---------- Popup window for live log ----------
function New-InstallPopup {
  $x=@"
<Window xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
        xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'
        Title='Installing UiPath' Width='720' Height='440'
        WindowStartupLocation='CenterOwner'
        ResizeMode='CanMinimize'
        Topmost='True'>
  <Grid Margin='16'>
    <Grid.RowDefinitions>
      <RowDefinition Height='Auto'/>
      <RowDefinition Height='Auto'/>
      <RowDefinition Height='Auto'/>
      <RowDefinition Height='*'/>
      <RowDefinition Height='Auto'/>
    </Grid.RowDefinitions>
    <TextBlock x:Name='TitleBlock' Text='Installing UiPath...' FontSize='18' FontWeight='SemiBold'/>
    <TextBlock x:Name='StatusBlock' Grid.Row='1' Margin='0,10,0,4' Text='Starting...'/>
    <ProgressBar x:Name='ProgBar' Grid.Row='2' Height='16' IsIndeterminate='True'/>
    <TextBox x:Name='LogBox' Grid.Row='3' Margin='0,12,0,12'
             VerticalScrollBarVisibility='Auto'
             HorizontalScrollBarVisibility='Auto'
             IsReadOnly='True'
             TextWrapping='NoWrap'/>
    <StackPanel Grid.Row='4' Orientation='Horizontal' HorizontalAlignment='Right'>
      <Button x:Name='CancelBtn' Content='Cancel' Width='90' Margin='0,0,8,0'/>
      <Button x:Name='CloseBtn' Content='Close' Width='90' IsEnabled='False'/>
    </StackPanel>
  </Grid>
</Window>
"@
  [xml]$xml=$x
  $rd=New-Object System.Xml.XmlNodeReader $xml
  return [Windows.Markup.XamlReader]::Load($rd)
}
function Popup-Log($popup,$msg) {
  $popup.Dispatcher.Invoke([Action]{
    $lb=$popup.FindName("LogBox")
    $lb.AppendText($msg + [Environment]::NewLine)
    $lb.ScrollToEnd()
  })
}
function Popup-Status($popup,$msg) {
  $popup.Dispatcher.Invoke([Action]{ $popup.FindName("StatusBlock").Text = $msg })
}
function Popup-Finish($popup) {
  $popup.Dispatcher.Invoke([Action]{
    $popup.FindName("CancelBtn").IsEnabled=$false
    $popup.FindName("CloseBtn").IsEnabled=$true
    $p=$popup.FindName("ProgBar")
    $p.IsIndeterminate=$false; $p.Value=100
    $popup.FindName("StatusBlock").Text="Finished."
  })
}

# ---------- Events: Screen 1 ----------
$BrowseDownloadTargetBtn.Add_Click({
  $res=Select-Folder $DownloadTargetBox.Text
  if ($res) { $DownloadTargetBox.Text=$res }
})

$DownloadBtn.Add_Click({
  if (-not (Test-Online)) { Show-Error "No internet connectivity"; return }
  if (-not $DownloadTargetBox.Text) { Show-Error "Select a target folder"; return }
  if (-not (Test-Path $DownloadTargetBox.Text)) { Show-Error "Target folder does not exist"; return }

  # Resolve URLs from metadata
  $urls = @()
  $prod = $ProductBox.SelectedItem
  $ver  = $VersionBox.SelectedItem
  if ($global:ProductMetadata -and $prod -and $ver) {
    $node = $global:ProductMetadata.$prod.$ver
    if ($null -ne $node) {
      if ($node -is [string]) { $urls = @($node) }
      elseif ($node.urls) { $urls = @($node.urls) }
      elseif ($node -is [System.Array]) { $urls = $node }
    }
  }

  if (-not $urls -or $urls.Count -eq 0) {
    Show-Error "No download URLs found for the selected Product/Version. Add json_files\product_versions.json or use Offline Files."
    return
  }

  $DownloadProgress.Minimum=0
  $DownloadProgress.Maximum=$urls.Count
  $DownloadProgress.Value=0
  $DownloadStatus.Text="Starting..."

  foreach ($u in $urls) {
    try {
      $name = [System.IO.Path]::GetFileName((New-Object System.Uri($u)).AbsolutePath)
      if (-not $name) { $name = ("file_{0}" -f ([Guid]::NewGuid())) }
      $out = Join-Path $DownloadTargetBox.Text $name
      $DownloadStatus.Text="Downloading $name..."
      Invoke-WebRequest $u -OutFile $out -UseBasicParsing
      $DownloadProgress.Value+=1
    } catch {
      Show-Error "Failed to download $u`n$($_.Exception.Message)"
      break
    }
  }
  $DownloadStatus.Text="Download complete."
})

$BrowseInstallerFolderBtn.Add_Click({
  $res = Select-Folder $InstallerFolderBox.Text
  if ($res) { $InstallerFolderBox.Text=$res }
})

$ValidateInstallerBtn.Add_Click({
  $msg = Validate-InstallerFiles $InstallerFolderBox.Text $ExpectedShaBox.Text
  $InstallerValidationResult.Text = $msg
  if ($msg -like "*OK*") {
    $global:Validated=$true
    $global:InstallerRoot=$InstallerFolderBox.Text
    $NextToInstallBtn.IsEnabled=$true
  } else {
    $global:Validated=$false
    $NextToInstallBtn.IsEnabled=$false
  }
})

$NextToInstallBtn.Add_Click({
  if (-not $global:Validated) { Show-Error "Validate installer first"; return }
  $Screen1Grid.Visibility="Collapsed"
  $Screen2Grid.Visibility="Visible"
})

# ---------- Events: Screen 2 ----------
$BackToScreen1Btn.Add_Click({
  $Screen2Grid.Visibility="Collapsed"
  $Screen1Grid.Visibility="Visible"
})

$GenerateScriptBtn.Add_Click({
  if (-not $global:InstallerRoot) { Show-Error "Missing installer folder"; return }
  $ScriptOutputBox.Text = Build-FullScript
})

$CopyScriptBtn.Add_Click({
  [Windows.Clipboard]::SetText($ScriptOutputBox.Text)
})

$SaveScriptBtn.Add_Click({
  $dlg = New-Object Microsoft.Win32.SaveFileDialog
  $dlg.Filter = "PowerShell Script (*.ps1)|*.ps1"
  $dlg.FileName = "UiPathInstallScript.ps1"
  if ($dlg.ShowDialog() -eq $true) {
    Set-Content -Encoding UTF8 -Path $dlg.FileName -Value $ScriptOutputBox.Text
    Show-Info "Saved"
  }
})

# ---------- Run Installer with popup log ----------
$RunInstallerBtn.Add_Click({
  if (-not $global:InstallerRoot) { Show-Error "Missing installer folder"; return }

  $msiCmd = Build-MSICommand
  $msiLog = Join-Path $env:TEMP ("UiPathStudio_{0}.log" -f ([Guid]::NewGuid()))
  $msiCmd += " /L*V `"$msiLog`""

  $popup = New-InstallPopup
  $popup.Owner = $window
  $popup.Topmost = $true
  $popup.Show()
  $cancelBtn = $popup.FindName("CancelBtn")
  $closeBtn  = $popup.FindName("CloseBtn")
  Popup-Status $popup "Installing UiPath..."
  Popup-Log    $popup "Running: $msiCmd"

  # Live tail timer
  $lastLength=0
  $timer = New-Object System.Windows.Threading.DispatcherTimer
  $timer.Interval=[TimeSpan]::FromMilliseconds(250)
  $timer.Add_Tick({
    try {
      if (Test-Path $msiLog) {
        $len=(Get-Item $msiLog).Length
        if ($len -gt $lastLength) {
          $txt=Get-Content $msiLog -Raw
          $delta=$txt.Substring($lastLength)
          $lastLength=$len
          if ($delta) { Popup-Log $popup $delta }
        }
      }
    } catch {}
  })
  $timer.Start()

  # Run MSI in background job
  $job = Start-Job -ScriptBlock {
    param($cmd)
    $proc = Start-Process -FilePath "cmd.exe" -ArgumentList "/c $cmd" -PassThru -WindowStyle Hidden
    $procId=$proc.Id
    $proc.WaitForExit()
    [pscustomobject]@{ PID=$procId; Exit=$proc.ExitCode }
  } -ArgumentList $msiCmd

  # Cancel -> kill msiexec
  $cancelBtn.Add_Click({
    try {
      Get-Process msiexec -ErrorAction SilentlyContinue | Stop-Process -Force
      Popup-Log $popup "Cancel requested."
      Popup-Status $popup "Cancelling..."
    } catch {
      Popup-Log $popup "Cancel error: $($_.Exception.Message)"
    }
  })

  while ($job.State -eq "Running") { Start-Sleep -Milliseconds 200 }

  try { $timer.Stop() } catch {}
  $result = Receive-Job $job
  Remove-Job $job -Force -ErrorAction SilentlyContinue

  if ($result.Exit -ne 0) {
    Popup-Log $popup "Installer exited with code $($result.Exit)"
    Popup-Status $popup "Installation failed. See log."
    Popup-Finish $popup
    return
  }

  Popup-Log $popup "Installation completed successfully."
  Popup-Finish $popup

  $closeBtn.Add_Click({ try { $popup.Close() } catch {} })
})

# ---------- Show main window ----------
$window.ShowDialog() | Out-Null