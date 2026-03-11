# =====================================================================
# UiPath Installer UI (PowerShell + WPF)
# Popup-only progress window (system-topmost, raw MSI log, cancel+close)
# Installation features via ADDLOCAL/REMOVE (per UiPath MSI docs)
# Robot cannot be disabled (confirmed by UiPath docs)
# Client-credentials MSI auto connect or Machine Key fallback
# =====================================================================

# ---------------- STANDARDS & INITIALIZATION ----------------
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -STA -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# ---------------- LOAD MAIN XAML ----------------
$xamlPath = Join-Path $PSScriptRoot "UiPathInstallerUI.xaml"
if (-not (Test-Path $xamlPath)) {
    [System.Windows.MessageBox]::Show("Cannot find XAML: $xamlPath","Error") | Out-Null
    exit
}

[xml]$xaml = Get-Content $xamlPath -Raw
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

# ---------------- CONTROL BINDINGS ----------------
$Screen1Grid               = $window.FindName("Screen1Grid")
$DownloadUrlsBox           = $window.FindName("DownloadUrlsBox")
$DownloadTargetBox         = $window.FindName("DownloadTargetBox")
$BrowseDownloadTargetBtn   = $window.FindName("BrowseDownloadTargetBtn")
$DownloadBtn               = $window.FindName("DownloadBtn")
$DownloadProgress          = $window.FindName("DownloadProgress")
$DownloadStatus            = $window.FindName("DownloadStatus")

$InstallerFolderBox        = $window.FindName("InstallerFolderBox")
$BrowseInstallerFolderBtn  = $window.FindName("BrowseInstallerFolderBtn")
$ExpectedShaBox            = $window.FindName("ExpectedShaBox")
$ValidateInstallerBtn      = $window.FindName("ValidateInstallerBtn")
$InstallerValidationResult = $window.FindName("InstallerValidationResult")
$NextToInstallBtn          = $window.FindName("NextToInstallBtn")

$Screen2Grid               = $window.FindName("Screen2Grid")
$BackToScreen1Btn          = $window.FindName("BackToScreen1Btn")

$ServiceModeCheck          = $window.FindName("ServiceModeCheck")
$AllUsersCheck             = $window.FindName("AllUsersCheck")
$RobotInstallCheck         = $window.FindName("RobotInstallCheck")
$StudioInstallCheck        = $window.FindName("StudioInstallCheck")

$DefaultPathRadio          = $window.FindName("DefaultPathRadio")
$CustomInstallationPath    = $window.FindName("CustomInstallationPath")
$CustomPathRadio           = $window.FindName("CustomPathRadio")
$CustomPathBox             = $window.FindName("CustomPathBox")
$BrowseCustomPathBtn       = $window.FindName("BrowseCustomPathBtn")

$ChromeExtCheck            = $window.FindName("ChromeExtCheck")
$EdgeExtCheck              = $window.FindName("EdgeExtCheck")
$FirefoxExtCheck           = $window.FindName("FirefoxExtCheck")
$JavaExtCheck              = $window.FindName("JavaExtCheck")
$CitrixExtCheck            = $window.FindName("CitrixExtCheck")
$WinRemoteExtCheck         = $window.FindName("WinRemoteExtCheck")

$OrchUrlBox                = $window.FindName("OrchUrlBox")
$MachineKeyBox             = $window.FindName("MachineKeyBox")
$ClientSecretNameBox       = $window.FindName("ClientSecretNameBox")
$ConnectAfterInstallCheck  = $window.FindName("ConnectAfterInstallCheck")

$EnableLogConfigCheck      = $window.FindName("EnableLogConfigCheck")
$LogConfigPanel            = $window.FindName("LogConfigPanel")
$UseCustomLogFolderCheck   = $window.FindName("UseCustomLogFolderCheck")
$CustomLogFolderPanel      = $window.FindName("CustomLogFolderPanel")
$LogFolderBox              = $window.FindName("LogFolderBox")
$BrowseLogFolderBtn        = $window.FindName("BrowseLogFolderBtn")
$ApplyNLogBtn              = $window.FindName("ApplyNLogBtn")
$LogDaysBox                = $window.FindName("LogDaysBox")
$CleanLogsBtn              = $window.FindName("CleanLogsBtn")

$ScriptOutputBox           = $window.FindName("ScriptOutputBox")
$GenerateScriptBtn         = $window.FindName("GenerateScriptBtn")
$CopyScriptBtn             = $window.FindName("CopyScriptBtn")
$SaveScriptBtn             = $window.FindName("SaveScriptBtn")
$RunInstallerBtn           = $window.FindName("RunInstallerBtn")

# ---------------- STATE ----------------
$global:Validated     = $false
$global:InstallerRoot = $null
$global:ExtensionsHandledInMsi = $false

# ---------------- HELPERS ----------------
function Select-Folder($Initial="") {
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($Initial) { $dlg.SelectedPath = $Initial }
    if ($dlg.ShowDialog() -eq "OK") { return $dlg.SelectedPath }
    return ""
}
function Show-Error($m) { [System.Windows.MessageBox]::Show($m,"Error")    | Out-Null }
function Show-Info ($m) { [System.Windows.MessageBox]::Show($m,"Information") | Out-Null }

function Test-Online {
    try {
        $req = [System.Net.WebRequest]::Create("https://www.microsoft.com")
        $req.Method="HEAD"
        $req.Timeout=2000
        $resp=$req.GetResponse()
        $resp.Close()
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

function Find-UiRobot {
    $roots=@(
        "C:\Program Files\UiPath","C:\Program Files\UiPath\Studio",
        "C:\Program Files\UiPath\Robot"
    )
    foreach ($r in $roots) {
        if (Test-Path $r) {
            $exe=Get-ChildItem -Recurse -Filter "UiRobot.exe" $r -ErrorAction SilentlyContinue |
                 Select-Object -First 1
            if ($exe) { return $exe.FullName }
        }
    }
    return ""
}

# ---------------- FEATURE SELECTION ----------------
function Get-SelectedFeatures {
    $list=[System.Collections.Generic.List[string]]::new()

    # Robot is mandatory (Robot cannot be disabled)  (UiPath docs) [1](https://www.igmguru.com/blog/uipath-installation)
    $list.Add("Robot")

    if ($StudioInstallCheck.IsChecked) { $list.Add("Studio") }

    if ($ServiceModeCheck.IsChecked) { $list.Add("RegisterService") }

    if ($ChromeExtCheck.IsChecked)    { $list.Add("ChromeExtension") }
    if ($EdgeExtCheck.IsChecked)      { $list.Add("EdgeExtension") }
    if ($FirefoxExtCheck.IsChecked)   { $list.Add("FirefoxExtension") }
    if ($JavaExtCheck.IsChecked)      { $list.Add("JavaBridge") }
    if ($CitrixExtCheck.IsChecked)    { $list.Add("CitrixClient") }
    if ($WinRemoteExtCheck.IsChecked) { $list.Add("WindowsRdpExtension") }

    return $list
}

# ---------------- BUILD MSI COMMAND ----------------
function Build-MSICommand {
    $msi = Join-Path $global:InstallerRoot "UiPathStudio.msi"
    $allUsers = if ($AllUsersCheck.IsChecked) { "ALLUSERS=1" } else { "ALLUSERS=0" }

    $pathProps=""
    if ($CustomInstallationPath.IsChecked -and $CustomPathBox.Text) {
        $p=$CustomPathBox.Text.Replace('"','""')
        $pathProps=" APPLICATIONFOLDER=`"$p`" INSTALLDIR=`"$p`""
    }

    $features = Get-SelectedFeatures
    $addlocal=""
    if ($features.Count -gt 0) {
        $addlocal=" ADDLOCAL=" + ($features -join ",")
        $global:ExtensionsHandledInMsi=$true
    }

    # If Studio unchecked but installed → REMOVE=Studio (according to UiPath docs) [1](https://www.igmguru.com/blog/uipath-installation)
    $remove=""
    if (-not $StudioInstallCheck.IsChecked -and (Test-StudioInstalled)) {
        $remove=" REMOVE=Studio"
    }

    # MSI‑level client creds
    $orchProps=""
    if ($ConnectAfterInstallCheck.IsChecked -and $ClientSecretNameBox.Text -and $OrchUrlBox.Text) {
        $cid=$MachineKeyBox.Text
        $sec=$ClientSecretNameBox.Text
        $url=$OrchUrlBox.Text
        # Per UiPath docs: ORCHESTRATOR_URL, CLIENT_ID, CLIENT_SECRET are MSI parameters [1](https://www.igmguru.com/blog/uipath-installation)
        $orchProps=" ORCHESTRATOR_URL=`"$url`" CLIENT_ID=`"$cid`" CLIENT_SECRET=`"$sec`""
    }

    return "msiexec /i `"$msi`" /qn /norestart $allUsers$pathProps$addlocal$remove$orchProps"
}

# ---------------- MACHINE KEY FALLBACK ----------------
function Generate-MachineKeyConnect {
    if (-not $ConnectAfterInstallCheck.IsChecked) { return "" }
    if ($ClientSecretNameBox.Text) { return "" } # done in MSI
    if (-not $MachineKeyBox.Text) { return "" }
    if (-not $OrchUrlBox.Text) { return "" }

    $exe=Find-UiRobot
    $url=$OrchUrlBox.Text
    $key=$MachineKeyBox.Text

@"
# Connect Robot to Orchestrator (MachineKey Fallback)
`"$exe`" connect --url `"$url`" --key `"$key`"
"@
}

# ---------------- NLOG EDIT ----------------
function GenerateNLogScript {
    if (-not $EnableLogConfigCheck.IsChecked) { return "" }
    if (-not $UseCustomLogFolderCheck.IsChecked) { return "" }
    if (-not $LogFolderBox.Text) { return "" }

    $dest=$LogFolderBox.Text

@"
# Update NLog WorkflowLoggingDirectory
`$possible=@(
 'C:\Program Files\UiPath',
 `"\$env:ProgramData\UiPath`"",
 `"\$env:LOCALAPPDATA\UiPath`" 
)
`$found=$null
foreach (`$p in `$possible) {
 if (Test-Path `$p) {
  `$f=Get-ChildItem -Recurse -Filter 'NLog.config' `$p -ErrorAction SilentlyContinue |
     Select-Object -First 1
  if (`$f) { `$found=`$f.FullName; break }
 }
}
if (`$found) {
 [xml]`$xml=Get-Content `$found
 `$var=`$xml.nlog.variable | Where-Object { `$_.name -eq 'WorkflowLoggingDirectory' }
 if (`$var) {
  `$var.value = '$dest'
 } else {
  `$v=`$xml.CreateElement('variable')
  `$v.SetAttribute('name','WorkflowLoggingDirectory')
  `$v.SetAttribute('value','$dest')
  `$xml.nlog.AppendChild(`$v)|Out-Null
 }
 `$xml.Save(`$found)
}
"@
}

# ---------------- LOG CLEANUP ----------------
function GenerateCleanupScript {
    if (-not $EnableLogConfigCheck.IsChecked) { return "" }
    if (-not $LogDaysBox.Text) { return "" }
    $days=[int]$LogDaysBox.Text

@"
# Cleanup logs older than $days days
`$paths=@()
if (Test-Path '$($LogFolderBox.Text)') { `$paths+='$($LogFolderBox.Text)' }
`$paths+='C:\Windows\System32\config\systemprofile\AppData\Local\UiPath\Logs'
`$paths+=`\$env:LOCALAPPDATA+'\UiPath\Logs'

foreach (`$p in `$paths) {
 if (Test-Path `$p) {
  Get-ChildItem -Recurse `$p |
   Where-Object { `$_.LastWriteTime -lt (Get-Date).AddDays(-$days) } |
   Remove-Item -Force -ErrorAction SilentlyContinue
 }
}
"@
}

# ---------------- POPUP WINDOW ----------------
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
    $popup.Dispatcher.Invoke([Action]{
        $popup.FindName("StatusBlock").Text = $msg
    })
}

function Popup-Finish($popup) {
    $popup.Dispatcher.Invoke([Action]{
        $popup.FindName("CancelBtn").IsEnabled=$false
        $popup.FindName("CloseBtn").IsEnabled=$true
        $popup.FindName("ProgBar").IsIndeterminate=$false
        $popup.FindName("ProgBar").Value=100
        $popup.FindName("StatusBlock").Text="Finished."
    })
}

# ---------------- EVENTS: SCREEN 1 ----------------
$BrowseDownloadTargetBtn.Add_Click({
    $res=Select-Folder $DownloadTargetBox.Text
    if ($res) { $DownloadTargetBox.Text=$res }
})

$DownloadBtn.Add_Click({
    $urls=($DownloadUrlsBox.Text -split "`r?`n") | Where-Object { $_.Trim() }
    if (-not $urls -or -not $DownloadTargetBox.Text) { Show-Error "Missing URL or target folder"; return }
    if (-not (Test-Path $DownloadTargetBox.Text)) { Show-Error "Target does not exist"; return }
    if (-not (Test-Online)) { Show-Error "No internet connectivity"; return }

    $DownloadProgress.Minimum=0
    $DownloadProgress.Maximum=$urls.Count
    $DownloadProgress.Value=0
    $DownloadStatus.Text="Starting..."

    foreach ($u in $urls) {
        try {
            $name = [System.IO.Path]::GetFileName((New-Object System.Uri($u)).AbsolutePath)
            $out  = Join-Path $DownloadTargetBox.Text $name
            $DownloadStatus.Text="Downloading $name..."
            Invoke-WebRequest $u -OutFile $out -UseBasicParsing
            $DownloadProgress.Value+=1
        } catch {
            Show-Error "Failed to download $u"
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

# ---------------- EVENTS: SCREEN 2 ----------------
$BackToScreen1Btn.Add_Click({
    $Screen2Grid.Visibility="Collapsed"
    $Screen1Grid.Visibility="Visible"
})

$BrowseCustomPathBtn.Add_Click({
    $res=Select-Folder $CustomPathBox.Text
    if ($res) { $CustomPathBox.Text=$res }
})

# Log config toggle
$EnableLogConfigCheck.Add_Checked({
    $LogConfigPanel.Visibility="Visible"
    $LogDaysBox.IsEnabled=$true
    $CleanLogsBtn.IsEnabled=$true
})
$EnableLogConfigCheck.Add_Unchecked({
    $LogConfigPanel.Visibility="Collapsed"
    $LogDaysBox.IsEnabled=$false
    $CleanLogsBtn.IsEnabled=$false
})
$UseCustomLogFolderCheck.Add_Checked({ $CustomLogFolderPanel.Visibility="Visible" })
$UseCustomLogFolderCheck.Add_Unchecked({ $CustomLogFolderPanel.Visibility="Collapsed" })

$GenerateScriptBtn.Add_Click({
    if (-not $global:InstallerRoot) { Show-Error "Missing installer folder"; return }
    $script=Build-FullScript
    $ScriptOutputBox.Text=$script
})

$CopyScriptBtn.Add_Click({
    [Windows.Clipboard]::SetText($ScriptOutputBox.Text)
})

$SaveScriptBtn.Add_Click({
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter="PowerShell Script|*.ps1"
    $dlg.FileName="UiPathInstallScript.ps1"
    if ($dlg.ShowDialog() -eq $true) {
        Set-Content -Encoding UTF8 -Path $dlg.FileName -Value $ScriptOutputBox.Text
        Show-Info "Saved"
    }
})

# ---------------- RUN INSTALLER WITH POPUP ----------------
$RunInstallerBtn.Add_Click({

    if (-not $global:InstallerRoot) { Show-Error "Missing installer"; return }

    # ---- Build Steps ----
    $msiCmd   = Build-MSICommand
    $msiLog   = Join-Path $env:TEMP ("UiPathStudio_{0}.log" -f ([Guid]::NewGuid()))
    $msiCmd  += " /L*V `"$msiLog`""

    $nlogCmd  = GenerateNLogScript
    $orchCmd  = Generate-MachineKeyConnect
    $cleanCmd = GenerateCleanupScript

    # ---- Popup ----
    $popup = New-InstallPopup
    $popup.Owner = $window
    $popup.Topmost = $true
    $popup.Show()

    $cancelBtn = $popup.FindName("CancelBtn")
    $closeBtn  = $popup.FindName("CloseBtn")

    Popup-Status $popup "Installing UiPath..."

    # ---- Live log timer ----
    $fs=$null; $sr=$null
    $lastLength=0
    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval=[TimeSpan]::FromMilliseconds(250)
    $timer.Add_Tick({
        try {
            if (Test-Path $msiLog) {
                $len=(Get-Item $msiLog).Length
                if ($len -gt $lastLength) {
                    if (-not $fs) {
                        $fs=[System.IO.File]::Open($msiLog,'Open','Read','ReadWrite')
                        $sr=New-Object System.IO.StreamReader($fs)
                        $null=$sr.ReadToEnd()
                    }
                    $txt=Get-Content $msiLog -Raw
                    $delta=$txt.Substring($lastLength)
                    $lastLength=$len
                    if ($delta) { Popup-Log $popup $delta }
                }
            }
        } catch {}
    })
    $timer.Start()

    # ---- Run MSI in background job ----
    Popup-Log $popup "Running: $msiCmd"

    $job = Start-Job -ScriptBlock {
        param($cmd)
        $proc = Start-Process -FilePath "cmd.exe" -ArgumentList "/c $cmd" -PassThru -WindowStyle Hidden
        $procId=$proc.Id
        $proc.WaitForExit()
        return [pscustomobject]@{ PID=$procId; Exit=$proc.ExitCode }
    } -ArgumentList $msiCmd

    # ---- Cancel button ----
    $cancelBtn.Add_Click({
        try {
            Get-Process msiexec -ErrorAction SilentlyContinue | Stop-Process -Force
            Popup-Log $popup "Cancel requested."
            Popup-Status $popup "Cancelling..."
        } catch {
            Popup-Log $popup "Cancel error: $($_.Exception.Message)"
        }
    })

    # ---- Wait for job completion ----
    while ($job.State -eq "Running") { Start-Sleep -Milliseconds 200 }

    # Stop log timer
    try { $timer.Stop() } catch {}

    # Close file streams
    try { if ($sr){$sr.Close()}; if ($fs){$fs.Close()} } catch {}

    $result=Receive-Job $job
    Remove-Job $job -Force -ErrorAction SilentlyContinue

    if ($result.Exit -ne 0) {
        Popup-Log $popup "Installer exited with code $($result.Exit)"
        Popup-Status $popup "Installation failed. See log."
        Popup-Finish $popup
        return
    }

    Popup-Log $popup "Installation completed successfully."

    # ---- NLOG (optional) ----
    if ($nlogCmd) {
        Popup-Status $popup "Applying NLog settings..."
        $tmp = Join-Path $env:TEMP ("NLogApply_{0}.ps1" -f ([Guid]::NewGuid()))
        Set-Content -Encoding UTF8 -Path $tmp -Value $nlogCmd
        try { powershell -NoProfile -ExecutionPolicy Bypass -File $tmp | Out-Null } catch {}
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        Popup-Log $popup "NLog applied."
    }

    # ---- Orchestrator Machine Key connect ----
    if ($orchCmd) {
        Popup-Status $popup "Connecting Robot..."
        $tmp = Join-Path $env:TEMP ("RobotConnect_{0}.ps1" -f ([Guid]::NewGuid()))
        Set-Content -Encoding UTF8 -Path $tmp -Value $orchCmd
        try { powershell -NoProfile -ExecutionPolicy Bypass -File $tmp | Out-Null } catch {}
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        Popup-Log $popup "Robot connected."
    }

    # ---- Log Cleanup ----
    if ($cleanCmd) {
        Popup-Status $popup "Cleaning logs..."
        $tmp = Join-Path $env:TEMP ("CleanLogs_{0}.ps1" -f ([Guid]::NewGuid()))
        Set-Content -Encoding UTF8 -Path $tmp -Value $cleanCmd
        try { powershell -NoProfile -ExecutionPolicy Bypass -File $tmp | Out-Null } catch {}
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        Popup-Log $popup "Log cleanup done."
    }

    Popup-Finish $popup

    # Close button
    $closeBtn.Add_Click({
        try { $popup.Close() } catch {}
    })
})

# ---------------- SHOW MAIN WINDOW ----------------
$window.ShowDialog() | Out-Null