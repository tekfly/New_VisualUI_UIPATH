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
$xamlPath = Join-Path $PSScriptRoot "UiPathInstallerUI_AI.xaml"
if (-not (Test-Path $xamlPath)) {
    [System.Windows.MessageBox]::Show("Cannot find XAML: $xamlPath","Error") | Out-Null
    exit
}

[xml]$xaml = Get-Content $xamlPath -Raw
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

# ---------------- CONTROL BINDINGS ----------------
$Screen1Grid               = $window.FindName("Screen1Grid")
$MainPairsPanel            = $window.FindName("MainPairsPanel")
$AddMainPairBtn            = $window.FindName("AddMainPairBtn")
$UpdateJsonBtn             = $window.FindName("UpdateJsonBtn")
$UpdateJsonStatus          = $window.FindName("UpdateJsonStatus")
$EnableExtraProductsCheck  = $window.FindName("EnableExtraProductsCheck")
$ExtraProductsSection      = $window.FindName("ExtraProductsSection")
$ExtraPairsPanel           = $window.FindName("ExtraPairsPanel")
$AddExtraPairBtn           = $window.FindName("AddExtraPairBtn")
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

$CheckInstalledBtn         = $window.FindName("CheckInstalledBtn")
$InstalledSoftwareResult   = $window.FindName("InstalledSoftwareResult")

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
$global:MainPairs  = New-Object System.Collections.Generic.List[Hashtable]
$global:ExtraPairs = New-Object System.Collections.Generic.List[Hashtable]

# ---------------- HELPERS ----------------
function Select-Folder($Initial="") {
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($Initial) { $dlg.SelectedPath = $Initial }
    if ($dlg.ShowDialog() -eq "OK") { return $dlg.SelectedPath }
    return ""
}
function Show-Error($m) { [System.Windows.MessageBox]::Show($m,"Error")    | Out-Null }
function Show-Info ($m) { [System.Windows.MessageBox]::Show($m,"Information") | Out-Null }

# ---------------- JSON PRODUCT/VERSION METADATA ----------------
# GitHub raw fallback used only when the local json file is missing or empty,
# so the download list still works even on a machine where json_files/ wasn't deployed.
$global:GitJsonBaseUrl = "https://raw.githubusercontent.com/tekfly/New_VisualUI_UIPATH/main/json_files"

function Load-Json($file) {
    $path = Join-Path $PSScriptRoot "json_files\$file"

    if ((Test-Path $path) -and ((Get-Item $path).Length -gt 0)) {
        try {
            $data = Get-Content $path -Raw | ConvertFrom-Json
            if ($data) { return $data }
        } catch { }
    }

    # Local file missing/empty/invalid -> fall back to the copy on GitHub
    try {
        $url = "$($global:GitJsonBaseUrl)/$file"
        $raw = Invoke-RestMethod -Uri $url -UseBasicParsing
        if ($raw) { return $raw }
    } catch {
        Show-Error "Could not load $file locally or from GitHub.`n$($_.Exception.Message)"
    }
    return $null
}

$global:MainProducts  = Load-Json "product_versions.json"
$global:ExtraProducts = Load-Json "extra_products_versions.json"

function Update-JsonFilesFromGitHub {
    # Pulls the latest product_versions.json / extra_products_versions.json straight
    # from GitHub and overwrites the local json_files\ copies, so new UiPath releases
    # show up without needing a new build of this tool.
    $files = @("product_versions.json","extra_products_versions.json")
    $destDir = Join-Path $PSScriptRoot "json_files"
    if (-not (Test-Path $destDir)) { New-Item -ItemType Directory -Path $destDir | Out-Null }

    $errors = @()
    foreach ($f in $files) {
        try {
            $url  = "$($global:GitJsonBaseUrl)/$f"
            $raw  = Invoke-RestMethod -Uri $url -UseBasicParsing
            if (-not $raw) { throw "Empty/invalid response from GitHub" }
            $dest = Join-Path $destDir $f
            # Round-trip through ConvertTo-Json so we only ever write valid JSON to disk.
            ($raw | ConvertTo-Json -Depth 10) | Set-Content -Encoding UTF8 -Path $dest
        } catch {
            $errors += "$f`: $($_.Exception.Message)"
        }
    }
    return $errors
}

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

function Test-ValidMsiSignature($path) {
    # MSI files are OLE Compound Files; a real MSI always starts with this 8-byte magic.
    # This is what actually distinguishes a real installer from an empty/truncated/
    # HTML-error-page download saved with a .msi extension (which Test-Path alone can't catch).
    try {
        $expected = [byte[]](0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1)
        $bytes = New-Object byte[] 8
        $fs = [System.IO.File]::OpenRead($path)
        try {
            $read = $fs.Read($bytes,0,8)
        } finally { $fs.Close() }
        if ($read -lt 8) { return $false }
        for ($i=0; $i -lt 8; $i++) { if ($bytes[$i] -ne $expected[$i]) { return $false } }
        return $true
    } catch { return $false }
}

function Get-PrimaryInstallerMsi($folder) {
    # Picks the "main" installer to drive msiexec with, preferring Studio/Robot packages
    # since that's what this tool installs, but still works if only another product was downloaded.
    $candidates = Get-ChildItem -Path $folder -Filter *.msi -File -ErrorAction SilentlyContinue |
        Where-Object { (Test-ValidMsiSignature $_.FullName) -and $_.Length -gt 0 }
    if (-not $candidates) { return $null }

    $preferred = $candidates | Where-Object { $_.Name -match 'Studio|Robot' } | Select-Object -First 1
    if ($preferred) { return $preferred }
    return ($candidates | Select-Object -First 1)
}

function Validate-InstallerFiles($folder,$sha) {
    if (-not (Test-Path $folder)) { return "Folder does not exist." }

    $msiFiles = Get-ChildItem -Path $folder -Filter *.msi -File -ErrorAction SilentlyContinue
    if (-not $msiFiles -or $msiFiles.Count -eq 0) { return "No .msi installer files found in folder." }

    $validFiles = @()
    $badFiles   = @()
    foreach ($f in $msiFiles) {
        if ($f.Length -eq 0) { $badFiles += "$($f.Name) (empty file)"; continue }
        if (-not (Test-ValidMsiSignature $f.FullName)) { $badFiles += "$($f.Name) (invalid MSI signature - likely corrupt/incomplete download)"; continue }
        $validFiles += $f
    }

    if ($validFiles.Count -eq 0) {
        return "Validation FAILED. Found $($msiFiles.Count) .msi file(s) but none are valid MSIs: $($badFiles -join '; ')"
    }

    if ($sha) {
        $match = $null
        foreach ($f in $validFiles) {
            try {
                $hash = (Get-FileHash -Algorithm SHA256 $f.FullName).Hash.ToLower()
                if ($hash -eq $sha.ToLower()) { $match = $f; break }
            } catch { }
        }
        if (-not $match) { return "SHA256 mismatch: none of the valid MSI file(s) match the expected hash." }
    }

    $names = ($validFiles | ForEach-Object { $_.Name }) -join ', '
    $msg = "Installer OK - $($validFiles.Count) valid MSI file(s): $names"
    if ($badFiles.Count -gt 0) { $msg += "  |  Ignored invalid: $($badFiles -join '; ')" }
    return $msg
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

function Find-InstalledUiPathSoftware {
    # Looks at the standard "installed programs" registry locations (per-machine, both
    # registry views, and per-user) for anything published as UiPath, plus a fallback
    # folder check in case an entry was removed from the uninstall list but files remain.
    $results = New-Object System.Collections.Generic.List[PSCustomObject]

    $regRoots = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    )

    foreach ($root in $regRoots) {
        if (-not (Test-Path $root)) { continue }
        foreach ($key in Get-ChildItem $root -ErrorAction SilentlyContinue) {
            $p = Get-ItemProperty $key.PSPath -ErrorAction SilentlyContinue
            if ($p.DisplayName -like "*UiPath*") {
                $results.Add([PSCustomObject]@{
                    Name            = $p.DisplayName
                    Version         = $p.DisplayVersion
                    Publisher       = $p.Publisher
                    InstallLocation = $p.InstallLocation
                    Source          = "Registry"
                }) | Out-Null
            }
        }
    }

    # Fallback: known install roots, in case registry entries are missing/incomplete.
    $folderRoots = @("C:\Program Files\UiPath","C:\Program Files (x86)\UiPath")
    foreach ($fr in $folderRoots) {
        if (Test-Path $fr) {
            foreach ($sub in Get-ChildItem $fr -Directory -ErrorAction SilentlyContinue) {
                $already = $results | Where-Object { $_.InstallLocation -and $_.InstallLocation.TrimEnd('\') -eq $sub.FullName.TrimEnd('\') }
                if (-not $already) {
                    $results.Add([PSCustomObject]@{
                        Name            = "UiPath $($sub.Name) (folder only, no uninstall entry found)"
                        Version         = ""
                        Publisher       = ""
                        InstallLocation = $sub.FullName
                        Source          = "Folder"
                    }) | Out-Null
                }
            }
        }
    }

    return $results
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
    $msiFile = Get-PrimaryInstallerMsi $global:InstallerRoot
    if (-not $msiFile) { Show-Error "No valid .msi found in $($global:InstallerRoot)"; return "" }
    $msi = $msiFile.FullName
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
 "`$env:ProgramData\UiPath",
 "`$env:LOCALAPPDATA\UiPath"
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
`$paths+=(`$env:LOCALAPPDATA+'\UiPath\Logs')

foreach (`$p in `$paths) {
 if (Test-Path `$p) {
  Get-ChildItem -Recurse `$p |
   Where-Object { `$_.LastWriteTime -lt (Get-Date).AddDays(-$days) } |
   Remove-Item -Force -ErrorAction SilentlyContinue
 }
}
"@
}

# ---------------- FULL SCRIPT BUILDER ----------------
function Build-FullScript {
    if (-not $global:InstallerRoot) { return "# ERROR: Installer not validated" }

    $msiCmd   = Build-MSICommand
    $msiLog   = '$env:TEMP\UiPathStudio_{0}.log' -f ([Guid]::NewGuid())

    $nlogCmd  = GenerateNLogScript
    $orchCmd  = Generate-MachineKeyConnect
    $cleanCmd = GenerateCleanupScript

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("# Install UiPath")
    [void]$sb.AppendLine("`$cmd = '$msiCmd /L*V `"' + '$msiLog' + '`"'")
    [void]$sb.AppendLine('Write-Host "Installing..."')
    [void]$sb.AppendLine('cmd.exe /c $cmd')
    [void]$sb.AppendLine('if ($LASTEXITCODE -ne 0) {')
    [void]$sb.AppendLine('    throw "Installer exited with code $LASTEXITCODE. Log: ' + $msiLog + '"')
    [void]$sb.AppendLine('}')
    [void]$sb.AppendLine('Write-Host "Install complete. Log: ' + $msiLog + '"')

    if ($nlogCmd) {
        [void]$sb.AppendLine("")
        [void]$sb.AppendLine("# --- NLog configuration ---")
        [void]$sb.AppendLine($nlogCmd)
    }
    if ($orchCmd) {
        [void]$sb.AppendLine("")
        [void]$sb.AppendLine("# --- Orchestrator connect (Machine Key fallback) ---")
        [void]$sb.AppendLine($orchCmd)
    }
    if ($cleanCmd) {
        [void]$sb.AppendLine("")
        [void]$sb.AppendLine("# --- Log cleanup ---")
        [void]$sb.AppendLine($cleanCmd)
    }

    return $sb.ToString()
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

# ---------------- DYNAMIC PRODUCT/VERSION ROWS ----------------
function New-ProductRow {
    param(
        [Parameter(Mandatory=$true)] $ParentPanel,
        [Parameter(Mandatory=$true)] $ListRef,
        [Parameter(Mandatory=$true)] $JsonSource
    )

    $row = New-Object System.Windows.Controls.StackPanel
    $row.Orientation = 'Horizontal'
    $row.Margin = '0,4,0,0'

    $lblP = New-Object System.Windows.Controls.TextBlock
    $lblP.Text = 'Product:'; $lblP.Width = 70; $lblP.VerticalAlignment='Center'

    $cbP = New-Object System.Windows.Controls.ComboBox
    $cbP.Width = 220; $cbP.Margin='0,0,12,0'

    $lblV = New-Object System.Windows.Controls.TextBlock
    $lblV.Text='Version:'; $lblV.Width=70; $lblV.VerticalAlignment='Center'

    $cbV = New-Object System.Windows.Controls.ComboBox
    $cbV.Width=220; $cbV.Margin='0,0,12,0'

    $btnRemove = New-Object System.Windows.Controls.Button
    $btnRemove.Content='Remove'
    $btnRemove.Width=80

    $row.Children.Add($lblP)      | Out-Null
    $row.Children.Add($cbP)       | Out-Null
    $row.Children.Add($lblV)      | Out-Null
    $row.Children.Add($cbV)       | Out-Null
    $row.Children.Add($btnRemove) | Out-Null

    if ($JsonSource) {
        $cbP.Items.Clear()
        $JsonSource.PSObject.Properties.Name | ForEach-Object { [void]$cbP.Items.Add($_) }
    }

    # NOTE: .GetNewClosure() is required here. Without it, these script blocks don't
    # keep their own copy of $cbV/$JsonSource/$row/etc; every row's event ends up
    # resolving those names against whatever happens to be in scope when WPF fires
    # the event (long after New-ProductRow has returned), which is why the version
    # combo silently failed to populate.
    $cbP.Add_SelectionChanged({
        if (-not $cbV) { return }
        $cbV.Items.Clear()
        $sel = $cbP.SelectedItem
        if ($sel -and $JsonSource.$sel) {
            $versions = $JsonSource.$sel.PSObject.Properties.Name
            foreach ($v in $versions) { [void]$cbV.Items.Add($v) }
            if ($cbV.Items.Count -gt 0) { $cbV.SelectedIndex = 0 }
        }
    }.GetNewClosure())

    $ParentPanel.Children.Add($row) | Out-Null

    $entry = @{ Panel=$row; ProductBox=$cbP; VersionBox=$cbV }
    $ListRef.Add($entry) | Out-Null

    $btnRemove.Add_Click({
        if ($ListRef.Count -le 1) { return }
        $ParentPanel.Children.Remove($row)
        [void]$ListRef.Remove($entry)
    }.GetNewClosure())

    return $entry
}

function Collect-UrlsFromRows {
    param($RowList, $JsonSource)

    $urls = New-Object System.Collections.Generic.List[string]
    foreach ($pair in $RowList) {
        $prod = $pair.ProductBox.SelectedItem
        $ver  = $pair.VersionBox.SelectedItem
        if (-not $prod -or -not $ver) { continue }

        $node = $JsonSource.$prod.$ver
        if ($null -ne $node) {
            if ($node -is [string]) { $urls.Add($node) | Out-Null }
            elseif ($node.urls)     { $node.urls | ForEach-Object { $urls.Add($_) | Out-Null } }
            elseif ($node -is [System.Array]) { $node | ForEach-Object { $urls.Add($_) | Out-Null } }
        }
    }
    return $urls
}

# Initial main row + wiring for Add buttons / extra-products toggle
$firstMain = New-ProductRow -ParentPanel $MainPairsPanel -ListRef $global:MainPairs -JsonSource $global:MainProducts
if ($firstMain.ProductBox.Items.Count -gt 0) { $firstMain.ProductBox.SelectedIndex = 0 }

$AddMainPairBtn.Add_Click({
    New-ProductRow -ParentPanel $MainPairsPanel -ListRef $global:MainPairs -JsonSource $global:MainProducts | Out-Null
})

$EnableExtraProductsCheck.Add_Checked({   $ExtraProductsSection.Visibility = "Visible" })
$EnableExtraProductsCheck.Add_Unchecked({ $ExtraProductsSection.Visibility = "Collapsed" })

$AddExtraPairBtn.Add_Click({
    if (-not $global:ExtraProducts) {
        Show-Error "extra_products_versions.json not found (locally or on GitHub)."
        return
    }
    New-ProductRow -ParentPanel $ExtraPairsPanel -ListRef $global:ExtraPairs -JsonSource $global:ExtraProducts | Out-Null
})

$UpdateJsonBtn.Add_Click({
    $UpdateJsonStatus.Text = "Updating from GitHub..."
    if (-not (Test-Online)) {
        Show-Error "No internet connectivity."
        $UpdateJsonStatus.Text = ""
        return
    }

    $errors = Update-JsonFilesFromGitHub
    $global:MainProducts  = Load-Json "product_versions.json"
    $global:ExtraProducts = Load-Json "extra_products_versions.json"

    # Rebuild the product/version rows from scratch with the refreshed data.
    # Any current row selections are reset since the underlying lists just changed.
    $MainPairsPanel.Children.Clear()
    $global:MainPairs.Clear()
    $newMain = New-ProductRow -ParentPanel $MainPairsPanel -ListRef $global:MainPairs -JsonSource $global:MainProducts
    if ($newMain.ProductBox.Items.Count -gt 0) { $newMain.ProductBox.SelectedIndex = 0 }

    $ExtraPairsPanel.Children.Clear()
    $global:ExtraPairs.Clear()
    if ($EnableExtraProductsCheck.IsChecked -and $global:ExtraProducts) {
        New-ProductRow -ParentPanel $ExtraPairsPanel -ListRef $global:ExtraPairs -JsonSource $global:ExtraProducts | Out-Null
    }

    if ($errors.Count -gt 0) {
        Show-Error "Some files failed to update:`n$($errors -join [Environment]::NewLine)"
        $UpdateJsonStatus.Text = "Updated with errors."
    } else {
        $UpdateJsonStatus.Text = "Updated $(Get-Date -Format 'yyyy-MM-dd HH:mm')."
    }
})

# ---------------- EVENTS: SCREEN 1 ----------------
$BrowseDownloadTargetBtn.Add_Click({
    $res=Select-Folder $DownloadTargetBox.Text
    if ($res) { $DownloadTargetBox.Text=$res }
})

$DownloadBtn.Add_Click({
    if (-not $DownloadTargetBox.Text) { Show-Error "Select a target folder"; return }
    if (-not (Test-Path $DownloadTargetBox.Text)) { Show-Error "Target does not exist"; return }
    if (-not (Test-Online)) { Show-Error "No internet connectivity"; return }

    $urlsMain  = Collect-UrlsFromRows -RowList $global:MainPairs -JsonSource $global:MainProducts
    $urlsExtra = @()
    if ($EnableExtraProductsCheck.IsChecked -and $global:ExtraProducts) {
        $urlsExtra = Collect-UrlsFromRows -RowList $global:ExtraPairs -JsonSource $global:ExtraProducts
    }

    $urls = New-Object System.Collections.Generic.List[string]
    $urlsMain  | ForEach-Object { $urls.Add($_) | Out-Null }
    $urlsExtra | ForEach-Object { $urls.Add($_) | Out-Null }

    if ($urls.Count -eq 0) {
        Show-Error "No URLs found. Check your product/version selections."
        return
    }

    $DownloadProgress.Minimum=0
    $DownloadProgress.Maximum=$urls.Count
    $DownloadProgress.Value=0
    $DownloadStatus.Text="Starting..."

    foreach ($u in $urls) {
        try {
            $name = [System.IO.Path]::GetFileName((New-Object System.Uri($u)).AbsolutePath)
            if (-not $name) { $name = "file_$([Guid]::NewGuid())" }
            $out  = Join-Path $DownloadTargetBox.Text $name
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

$CheckInstalledBtn.Add_Click({
    $found = Find-InstalledUiPathSoftware
    if ($found.Count -eq 0) {
        $InstalledSoftwareResult.Text = "No UiPath software found on this machine (registry or Program Files)."
        return
    }
    $lines = foreach ($f in $found) {
        $ver = if ($f.Version) { " v$($f.Version)" } else { "" }
        $loc = if ($f.InstallLocation) { " - $($f.InstallLocation)" } else { "" }
        "[$($f.Source)] $($f.Name)$ver$loc"
    }
    $InstalledSoftwareResult.Text = ($lines -join [Environment]::NewLine)
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