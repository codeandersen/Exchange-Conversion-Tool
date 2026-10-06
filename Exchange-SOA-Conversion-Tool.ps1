#Requires -Version 5.1

<#
        .SYNOPSIS
        Exchange SOA Conversion Tool - Source of Authority management for Exchange mailboxes

        .DESCRIPTION
        GUI tool to manage Source of Authority (SOA) conversion for Exchange mailboxes
        (IsDirSynced users) via ExchangeOnlineManagement.

        Features:
        - Live search by display name or email address across all loaded data
        - Hide Converted filter
        - Pagination (100 per page)
        - Batch conversion with confirmation
        - Full logging to a single session log file

        .EXAMPLE
        .\Exchange-SOA-Conversion-Tool.ps1

        .NOTES
        Version: 1.10

        .LINK
        https://learn.microsoft.com/en-us/exchange/hybrid-deployment/enable-exchange-attributes-cloud-management

        .COPYRIGHT
        MIT License, feel free to distribute and use as you like, please leave author information.

        BLOG: http://www.hcandersen.net
        LinkedIn: https://www.linkedin.com/in/hanschrandersen/

        .DISCLAIMER
        This script is provided AS-IS, with no warranty - Use at own risk.
    #>

$script:Version       = "1.10"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:ScriptPath    = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:LogFile       = Join-Path $script:ScriptPath "SOAConversion_$(Get-Date -Format 'yyyyMMdd_HHmm').log"

# ---- State ----
$script:AllUsers             = @()
$script:AllUsersUnfiltered   = @()
$script:UsersCurrentPage     = 1
$script:UsersHideConverted   = $false
$script:UsersSortColumn      = ""
$script:UsersSortDirection   = "Ascending"
$script:ExoConnected         = $false
$script:ExoModuleLoaded      = $false

$script:PageSize = 100

# ==============================================================================
# SHARED HELPERS
# ==============================================================================

function Write-Log {
    param(
        [Parameter(Mandatory=$true)][string]$Message,
        [Parameter(Mandatory=$false)][ValidateSet('INFO','WARNING','ERROR')][string]$Level = 'INFO'
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry  = "[$timestamp] [$Level] $Message"
    $retries = 3
    for ($i = 0; $i -lt $retries; $i++) {
        try {
            [System.IO.File]::AppendAllText($script:LogFile, "$logEntry`r`n", [System.Text.Encoding]::UTF8)
            break
        } catch {
            if ($i -eq ($retries - 1)) { Write-Host "WARNING: Could not write to log: $($_.Exception.Message)" }
            Start-Sleep -Milliseconds 50
        }
    }
    Write-Host $logEntry
}

function Get-FilteredData {
    param(
        [array]$Source,
        [string]$SearchText,
        [bool]$HideConverted,
        [string]$ConvertedPropertyName,
        [string]$EmailPropertyName = "Mail"
    )
    $result = $Source

    if ($HideConverted) {
        $result = $result | Where-Object { $_.$ConvertedPropertyName -ne $true }
    }

    if (-not [string]::IsNullOrWhiteSpace($SearchText)) {
        $lower = $SearchText.ToLower()
        $result = $result | Where-Object {
            $dn   = ([string]$_.DisplayName).ToLower()
            $mail = ([string]$_.$EmailPropertyName).ToLower()
            ($dn.Contains($lower)) -or ($mail.Contains($lower))
        }
    }

    return @($result)
}

function New-StyledButton {
    param(
        [string]$Text,
        [int]$X, [int]$Y,
        [int]$Width = 160, [int]$Height = 38,
        [string]$BackHex = "#E1E1E1",
        [string]$ForeHex = "#1F1F1F",
        [bool]$Enabled = $true
    )
    $btn = New-Object System.Windows.Forms.Button
    $btn.Location  = New-Object System.Drawing.Point($X, $Y)
    $btn.Size      = New-Object System.Drawing.Size($Width, $Height)
    $btn.Text      = $Text
    $btn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btn.BackColor = [System.Drawing.ColorTranslator]::FromHtml($BackHex)
    $btn.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($ForeHex)
    $btn.Font      = New-Object System.Drawing.Font("Segoe UI", 10)
    $btn.FlatAppearance.BorderSize = 0
    $btn.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $btn.Enabled   = $Enabled
    return $btn
}

function New-DataGrid {
    $dgv = New-Object System.Windows.Forms.DataGridView
    $dgv.AllowUserToAddRows    = $false
    $dgv.AllowUserToDeleteRows = $false
    $dgv.ReadOnly              = $true
    $dgv.SelectionMode         = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $dgv.MultiSelect           = $true
    $dgv.AutoSizeColumnsMode   = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $dgv.BorderStyle           = [System.Windows.Forms.BorderStyle]::FixedSingle
    $dgv.BackgroundColor       = [System.Drawing.Color]::White
    $dgv.GridColor             = [System.Drawing.ColorTranslator]::FromHtml("#E1E1E1")
    $dgv.DefaultCellStyle.BackColor          = [System.Drawing.Color]::White
    $dgv.DefaultCellStyle.ForeColor          = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")
    $dgv.DefaultCellStyle.SelectionBackColor = [System.Drawing.ColorTranslator]::FromHtml("#0078D4")
    $dgv.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
    $dgv.DefaultCellStyle.Font               = New-Object System.Drawing.Font("Segoe UI", 9)
    $dgv.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#F5F5F5")
    $dgv.ColumnHeadersDefaultCellStyle.BackColor   = [System.Drawing.ColorTranslator]::FromHtml("#E1E1E1")
    $dgv.ColumnHeadersDefaultCellStyle.ForeColor   = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")
    $dgv.ColumnHeadersDefaultCellStyle.Font        = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $dgv.ColumnHeadersDefaultCellStyle.Alignment   = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleLeft
    $dgv.ColumnHeadersDefaultCellStyle.Padding     = New-Object System.Windows.Forms.Padding(5, 0, 0, 0)
    $dgv.ColumnHeadersHeight                = 40
    $dgv.ColumnHeadersHeightSizeMode        = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
    $dgv.EnableHeadersVisualStyles          = $false
    $dgv.RowHeadersVisible                  = $false
    $dgv.CellBorderStyle                    = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
    $dgv.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor
                  [System.Windows.Forms.AnchorStyles]::Bottom -bor
                  [System.Windows.Forms.AnchorStyles]::Left -bor
                  [System.Windows.Forms.AnchorStyles]::Right
    return $dgv
}

function Add-DgvTextColumn {
    param($Dgv, [string]$Name, [string]$Header, [int]$FillWeight, [bool]$Visible = $true)
    $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.Name       = $Name
    $col.HeaderText = $Header
    if ($FillWeight -gt 0) { $col.FillWeight = $FillWeight }
    $col.Visible    = $Visible
    [void]$Dgv.Columns.Add($col)
}

# ==============================================================================
# MODULE / CONNECTION CHECKS
# ==============================================================================

function Add-ModulePathIfMissing {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    $existing = $env:PSModulePath -split ';' | ForEach-Object { $_.TrimEnd('\') }
    if ($existing -notcontains $Path.TrimEnd('\')) {
        $env:PSModulePath = "$env:PSModulePath;$Path"
        Write-Log "Added '$Path' to PSModulePath for this session."
    }
}

function Import-ExchangeModule {
    Write-Log "PowerShell $($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition))"

    $psg = Get-Module -ListAvailable -Name PowerShellGet | Sort-Object Version -Descending | Select-Object -First 1
    if ($psg) { Write-Log "PowerShellGet v$($psg.Version)" } else { Write-Log "PowerShellGet module not found." -Level WARNING }

    if ($PSVersionTable.PSEdition -eq 'Core') {
        $userModuleDir     = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'PowerShell\Modules'
        $allUsersModuleDir = "$env:ProgramFiles\PowerShell\Modules"
    } else {
        $userModuleDir     = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'WindowsPowerShell\Modules'
        $allUsersModuleDir = "$env:ProgramFiles\WindowsPowerShell\Modules"
    }
    Add-ModulePathIfMissing $userModuleDir
    Add-ModulePathIfMissing $allUsersModuleDir
    Write-Log "PSModulePath: $env:PSModulePath"

    $dotNetTooOld = $false
    if ($PSVersionTable.PSEdition -ne 'Core') {
        $rel = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -ErrorAction SilentlyContinue).Release
        Write-Log ".NET Framework release: $rel"
        if ($null -eq $rel -or $rel -lt 461808) {
            $dotNetTooOld = $true
            Write-Log ".NET Framework 4.7.2 or later is required for the EXO v3 module." -Level WARNING
        }
    }

    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    $installScope = if ($isAdmin) { 'AllUsers' } else { 'CurrentUser' }
    Write-Log "Elevated: $isAdmin  -  install scope: $installScope"

    $script:ExoImportTried = @()

    $tryImport = {
        param([array]$Modules)
        foreach ($m in $Modules) {
            try {
                Import-Module -Name $m.Path -Force -DisableNameChecking -ErrorAction Stop
                Write-Log "Imported ExchangeOnlineManagement v$($m.Version) from '$($m.Path)'."
                return $true
            } catch {
                $script:ExoImportTried += "$($m.Path)  -  $($_.Exception.Message)"
                Write-Log "Failed to import ExchangeOnlineManagement from '$($m.Path)': $($_.Exception.Message)" -Level WARNING
            }
        }
        return $false
    }

    $getExoModules = {
        $found = @(Get-Module -ListAvailable -Name ExchangeOnlineManagement |
                   Sort-Object @{ Expression = 'Version'; Descending = $true },
                               @{ Expression = { $_.ModuleBase -like "$env:ProgramFiles*" }; Descending = $true })
        if ($found.Count -eq 0) {
            Get-InstalledModule -Name ExchangeOnlineManagement -AllVersions -ErrorAction SilentlyContinue | ForEach-Object {
                Write-Log "Get-InstalledModule found ExchangeOnlineManagement v$($_.Version) at '$($_.InstalledLocation)'."
                $psd1 = Join-Path $_.InstalledLocation 'ExchangeOnlineManagement.psd1'
                if (Test-Path $psd1) {
                    $found += [PSCustomObject]@{ Path = $psd1; Version = $_.Version; ModuleBase = $_.InstalledLocation }
                }
            }
        }
        return $found
    }

    $mods = @(& $getExoModules)
    if ($mods.Count -eq 0) {
        Write-Log "ExchangeOnlineManagement module not found." -Level WARNING
    } else {
        foreach ($m in $mods) { Write-Log "Found ExchangeOnlineManagement v$($m.Version) at '$($m.ModuleBase)'." }
        if (& $tryImport $mods) { return [bool]$true }
        Write-Log "No installed version of ExchangeOnlineManagement could be imported." -Level WARNING
    }

    $answer = [System.Windows.Forms.MessageBox]::Show(
        "The ExchangeOnlineManagement module is not available or could not be loaded.`n`nReinstall it now ($installScope scope)?",
        "Module Required", [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) {
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
            if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue | Where-Object { $_.Version -ge [version]'2.8.5.201' })) {
                Write-Log "Installing NuGet package provider ($installScope scope)..."
                Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope $installScope -Force -ErrorAction Stop | Out-Null
            }
            Write-Log "Installing ExchangeOnlineManagement ($installScope scope)..."
            Install-Module -Name ExchangeOnlineManagement -Scope $installScope -Force -AllowClobber -ErrorAction Stop | Out-Null
            Write-Log "ExchangeOnlineManagement installed successfully."
        } catch {
            Write-Log "Failed to install ExchangeOnlineManagement: $($_.Exception.Message)" -Level ERROR
        }

        $mods = @(& $getExoModules)
        foreach ($m in $mods) { Write-Log "Found ExchangeOnlineManagement v$($m.Version) at '$($m.ModuleBase)'." }
        if (& $tryImport $mods) { return [bool]$true }
    }

    Write-Log "ExchangeOnlineManagement could not be loaded. Paths tried: $($script:ExoImportTried -join ' | ')" -Level ERROR
    $triedText = if ($script:ExoImportTried.Count -gt 0) { "`n`nPaths tried:`n" + ($script:ExoImportTried -join "`n") } else { "" }
    $netNote   = if ($dotNetTooOld) { "`n`nNote: the EXO module requires .NET Framework 4.7.2 or later.`nInstall .NET Framework 4.7.2 or later (4.8 recommended)." } else { "" }
    [void][System.Windows.Forms.MessageBox]::Show(
        "ExchangeOnlineManagement could not be loaded.$triedText$netNote`n`nFix manually:`n  Uninstall-Module ExchangeOnlineManagement -AllVersions -Force`n  Install-Module ExchangeOnlineManagement -Scope $installScope -Force",
        "Module Load Failed", [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error)
    return [bool]$false
}

# ==============================================================================
# USERS -- BACKEND
# ==============================================================================

function Connect-EXOSession {
    Write-Log "Connecting to Exchange Online..."
    try {
        $connectParams = @{ ShowBanner = $false; ErrorAction = 'Stop' }
        if ((Get-Command Connect-ExchangeOnline).Parameters.ContainsKey('DisableWAM')) {
            $connectParams['DisableWAM'] = $true
            Write-Log "Using browser-based sign-in (WAM disabled)."
        }
        # Clear the WinForms SynchronizationContext so MSAL's async sign-in doesn't deadlock the UI thread.
        $prevSyncContext = [System.Threading.SynchronizationContext]::Current
        [System.Threading.SynchronizationContext]::SetSynchronizationContext($null)
        try { Connect-ExchangeOnline @connectParams }
        finally { [System.Threading.SynchronizationContext]::SetSynchronizationContext($prevSyncContext) }
        Write-Log "Connected to Exchange Online."
        return $true
    } catch {
        Write-Log "EXO connect failed: $($_.Exception.Message)" -Level ERROR
        [void][System.Windows.Forms.MessageBox]::Show(
            "Failed to connect to Exchange Online.`n`nError: $($_.Exception.Message)",
            "Connection Failed", [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
        return $false
    }
}

function Get-ExchangeUsers {
    Write-Log "Retrieving mailboxes from Exchange Online..."
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $list  = New-Object System.Collections.Generic.List[object]
        $total = 0
        Update-Busy "Requesting mailboxes from Exchange Online (first results can take a while)..."
        Get-Mailbox -ResultSize Unlimited -ErrorAction Stop | ForEach-Object {
            $total++
            if ($_.IsDirSynced -eq $true) {
                $list.Add([PSCustomObject]@{
                    DisplayName             = $_.DisplayName
                    PrimarySmtpAddress      = [string]$_.PrimarySmtpAddress
                    IsExchangeCloudManaged  = $_.IsExchangeCloudManaged
                    UserPrincipalName       = $_.UserPrincipalName
                    IsDirSynced             = $_.IsDirSynced
                })
            }
            if ($total % 25 -eq 0) {
                Update-Busy "Loading mailboxes... $total retrieved, $($list.Count) directory-synced"
            }
        }
        $watch.Stop()
        Write-Log "Retrieved $total total mailboxes in $([int]$watch.Elapsed.TotalSeconds)s."
        Write-Log "Filtered to $($list.Count) IsDirSynced mailboxes."
        return ,$list.ToArray()
    } catch {
        $watch.Stop()
        Write-Log "Failed to retrieve mailboxes: $($_.Exception.Message)" -Level ERROR
        [void][System.Windows.Forms.MessageBox]::Show(
            "Failed to retrieve mailboxes.`n`nError: $($_.Exception.Message)",
            "Retrieval Failed", [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
        return $null
    }
}

function Backup-MailboxAttributes {
    param($User)
    $mbx  = Get-Mailbox -Identity $User.UserPrincipalName -ErrorAction Stop
    $name = $User.DisplayName
    Write-Log "User '$name' - Alias (mailNickname): $($mbx.Alias)"
    Write-Log "User '$name' - Primary SMTP: $($mbx.PrimarySmtpAddress)"
    Write-Log "User '$name' - All Email Addresses: $(@($mbx.EmailAddresses) -join '; ')"
    Write-Log "User '$name' - HiddenFromAddressListsEnabled: $($mbx.HiddenFromAddressListsEnabled)"
    $custom = @()
    foreach ($n in 1..15) {
        $v = $mbx."CustomAttribute$n"
        if (-not [string]::IsNullOrWhiteSpace([string]$v)) { $custom += "CustomAttribute$n=$v" }
    }
    foreach ($n in 1..5) {
        $v = @($mbx."ExtensionCustomAttribute$n") -join ','
        if (-not [string]::IsNullOrWhiteSpace($v)) { $custom += "ExtensionCustomAttribute$n=$v" }
    }
    $customText = if ($custom.Count -gt 0) { $custom -join '; ' } else { 'None set' }
    Write-Log "User '$name' - Custom Attributes: $customText"
    return [string]$mbx.PrimarySmtpAddress
}

function Convert-UserToCloud {
    param($User)
    Write-Log "Converting user '$($User.DisplayName)' ($($User.UserPrincipalName)) to Cloud Managed..."
    try {
        $smtp = Backup-MailboxAttributes -User $User
        Set-Mailbox -Identity $User.UserPrincipalName -IsExchangeCloudManaged $true -ErrorAction Stop
        Write-Log "Successfully converted user '$($User.DisplayName)' ($smtp) to Cloud Managed." -Level INFO
        return $true
    } catch {
        Write-Log "Failed to convert '$($User.DisplayName)': $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

function Convert-UserToOnPrem {
    param($User)
    Write-Log "Converting user '$($User.DisplayName)' ($($User.UserPrincipalName)) to On-Prem Managed..."
    try {
        $smtp = Backup-MailboxAttributes -User $User
        Set-Mailbox -Identity $User.UserPrincipalName -IsExchangeCloudManaged $false -ErrorAction Stop
        Write-Log "Successfully converted user '$($User.DisplayName)' ($smtp) to On-Prem Managed." -Level INFO
        return $true
    } catch {
        Write-Log "Failed to roll back '$($User.DisplayName)': $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

Write-Log "========================================"
Write-Log "Exchange SOA Conversion Tool v$script:Version Started"
Write-Log "========================================"

$script:ExoModuleLoaded = Import-ExchangeModule
if (-not $script:ExoModuleLoaded) {
    Write-Log "ExchangeOnlineManagement module not loaded. Connect will be disabled." -Level WARNING
}

# ==============================================================================
# MAIN FORM
# ==============================================================================

$form = New-Object System.Windows.Forms.Form
$form.Text          = "Exchange SOA Conversion Tool"
# Build at the design size so all anchored child controls lay out correctly.
# The window is clamped to the screen working area at the end (before ShowDialog),
# which triggers the anchor re-layout so every column stays on-screen.
$form.MinimumSize   = New-Object System.Drawing.Size(820, 640)
$form.Size          = New-Object System.Drawing.Size(1150, 800)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
$form.BackColor     = [System.Drawing.ColorTranslator]::FromHtml("#F3F3F3")
$form.ShowIcon      = $false

# --- Header Panel ---
$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Height    = 85
$headerPanel.Dock      = [System.Windows.Forms.DockStyle]::Top
$headerPanel.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#E8E8E8")

$labelTitle = New-Object System.Windows.Forms.Label
$labelTitle.Location  = New-Object System.Drawing.Point(20, 12)
$labelTitle.Size      = New-Object System.Drawing.Size(900, 40)
$labelTitle.Text      = "Exchange SOA Conversion Tool"
$labelTitle.Font      = New-Object System.Drawing.Font("Segoe UI", 20, [System.Drawing.FontStyle]::Regular)
$labelTitle.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$labelTitle.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")
$labelTitle.BackColor = [System.Drawing.Color]::Transparent
$labelTitle.Anchor    = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$headerPanel.Controls.Add($labelTitle)

$labelDesc = New-Object System.Windows.Forms.Label
$labelDesc.Location  = New-Object System.Drawing.Point(20, 55)
$labelDesc.Size      = New-Object System.Drawing.Size(900, 22)
$labelDesc.Text      = "Manage Source of Authority for Exchange mailboxes"
$labelDesc.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
$labelDesc.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#605E5C")
$labelDesc.BackColor = [System.Drawing.Color]::Transparent
$labelDesc.Anchor    = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$headerPanel.Controls.Add($labelDesc)

$pictureBoxLogo = New-Object System.Windows.Forms.PictureBox
$pictureBoxLogo.Size     = New-Object System.Drawing.Size(80, 68)
$pictureBoxLogo.Location = New-Object System.Drawing.Point(1050, 8)
$pictureBoxLogo.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$pictureBoxLogo.BackColor = [System.Drawing.Color]::Transparent
$pictureBoxLogo.Anchor   = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$logoPath = Join-Path $script:ScriptPath "logo.png"
if (Test-Path $logoPath) {
    try {
        $fs  = [System.IO.File]::OpenRead($logoPath)
        $ms  = New-Object System.IO.MemoryStream
        $fs.CopyTo($ms); $fs.Close(); $fs.Dispose()
        $ms.Position = 0
        $pictureBoxLogo.Image = [System.Drawing.Image]::FromStream($ms)
    } catch { Write-Log "Failed to load logo: $($_.Exception.Message)" -Level WARNING }
}
$headerPanel.Controls.Add($pictureBoxLogo)

# ==============================================================================
# USERS -- GUI
# ==============================================================================

$uToolbar = New-Object System.Windows.Forms.Panel
$uToolbar.Height    = 52
$uToolbar.Dock      = [System.Windows.Forms.DockStyle]::Top
$uToolbar.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#F3F3F3")

$uBtnConnect    = New-StyledButton -Text "Connect to EXO"           -X 5   -Y 7 -Width 155 -Height 36 -BackHex "#0078D4" -ForeHex "#FFFFFF"
$uBtnRefresh    = New-StyledButton -Text "Refresh"                  -X 168 -Y 7 -Width 110 -Height 36 -Enabled $false
$uBtnDisconnect = New-StyledButton -Text "Disconnect"               -X 286 -Y 7 -Width 120 -Height 36 -Enabled $false

$uSearchLabel = New-Object System.Windows.Forms.Label
$uSearchLabel.Location  = New-Object System.Drawing.Point(420, 14)
$uSearchLabel.Size      = New-Object System.Drawing.Size(55, 22)
$uSearchLabel.Text      = "Search:"
$uSearchLabel.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
$uSearchLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")

$uSearchBox = New-Object System.Windows.Forms.TextBox
$uSearchBox.Location    = New-Object System.Drawing.Point(478, 11)
$uSearchBox.Size        = New-Object System.Drawing.Size(240, 26)
$uSearchBox.Font        = New-Object System.Drawing.Font("Segoe UI", 9)
$uSearchBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

$uChkHide = New-Object System.Windows.Forms.CheckBox
$uChkHide.Location  = New-Object System.Drawing.Point(730, 14)
$uChkHide.Size      = New-Object System.Drawing.Size(195, 24)
$uChkHide.Text      = "Hide Converted Users"
$uChkHide.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
$uChkHide.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")

$uToolbar.Controls.AddRange(@($uBtnConnect, $uBtnRefresh, $uBtnDisconnect, $uSearchLabel, $uSearchBox, $uChkHide))

$uGrid = New-DataGrid
$uGrid.Dock = [System.Windows.Forms.DockStyle]::Fill
Add-DgvTextColumn $uGrid "DisplayName"          "Display Name"   30
Add-DgvTextColumn $uGrid "Email"                "Email Address"  38
Add-DgvTextColumn $uGrid "IsExchangeCloudManaged" "Cloud Managed" 22
Add-DgvTextColumn $uGrid "UserPrincipalName"    "UPN"            0  $false

$uPaginPanel = New-Object System.Windows.Forms.Panel
$uPaginPanel.Height    = 36
$uPaginPanel.Dock      = [System.Windows.Forms.DockStyle]::Bottom
$uPaginPanel.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#F3F3F3")

$uBtnPrev = New-StyledButton -Text "< Previous" -X 0 -Y 1 -Width 100 -Height 28 -Enabled $false
$uBtnPrev.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom

$uPageInfo = New-Object System.Windows.Forms.Label
$uPageInfo.Location  = New-Object System.Drawing.Point(108, 6)
$uPageInfo.Size      = New-Object System.Drawing.Size(700, 22)
$uPageInfo.Text      = "No users loaded"
$uPageInfo.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
$uPageInfo.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#605E5C")
$uPageInfo.Anchor    = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom

$uBtnNext = New-StyledButton -Text "Next >" -X 1006 -Y 1 -Width 100 -Height 28 -Enabled $false
$uBtnNext.Anchor = [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom

$uPaginPanel.Controls.AddRange(@($uBtnPrev, $uPageInfo, $uBtnNext))
$uPaginPanel.Add_SizeChanged({
    $w = $uPaginPanel.ClientSize.Width
    $uBtnNext.Location  = New-Object System.Drawing.Point(($w - $uBtnNext.Width), 1)
    $uPageInfo.Location = New-Object System.Drawing.Point(108, 6)
    $uPageInfo.Size     = New-Object System.Drawing.Size([Math]::Max(50, $w - 108 - $uBtnNext.Width - 10), 22)
})

$uActionPanel = New-Object System.Windows.Forms.Panel
$uActionPanel.Height    = 52
$uActionPanel.Dock      = [System.Windows.Forms.DockStyle]::Bottom
$uActionPanel.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#F3F3F3")

$uBtnToCloud  = New-StyledButton -Text "Convert to Cloud Managed"   -X 0   -Y 4 -Width 220 -Height 40 -Enabled $false
$uBtnToOnPrem = New-StyledButton -Text "Convert to On-Prem Managed" -X 228 -Y 4 -Width 230 -Height 40 -Enabled $false
$uBtnOpenLog  = New-StyledButton -Text "Open Log"                   -X 466 -Y 4 -Width 120 -Height 40

$uStatusLabel = New-Object System.Windows.Forms.Label
$uStatusLabel.Location  = New-Object System.Drawing.Point(600, 12)
$uStatusLabel.Size      = New-Object System.Drawing.Size(500, 22)
$uStatusLabel.Text      = "Not connected"
$uStatusLabel.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
$uStatusLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#605E5C")
$uStatusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$uStatusLabel.Anchor    = [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom

$uVersionLabel = New-Object System.Windows.Forms.Label
$uVersionLabel.Location  = New-Object System.Drawing.Point(600, 34)
$uVersionLabel.Size      = New-Object System.Drawing.Size(500, 18)
$uVersionLabel.Text      = "Version $script:Version"
$uVersionLabel.Font      = New-Object System.Drawing.Font("Segoe UI", 8)
$uVersionLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#808080")
$uVersionLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$uVersionLabel.Anchor    = [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom

$uActionPanel.Controls.AddRange(@($uBtnToCloud, $uBtnToOnPrem, $uBtnOpenLog, $uStatusLabel, $uVersionLabel))
$uActionPanel.Add_SizeChanged({
    $w = $uActionPanel.ClientSize.Width
    $uStatusLabel.Location  = New-Object System.Drawing.Point([Math]::Max(0, $w - $uStatusLabel.Width), 6)
    $uVersionLabel.Location = New-Object System.Drawing.Point([Math]::Max(0, $w - $uVersionLabel.Width), 30)
})

$uProgress = New-Object System.Windows.Forms.ProgressBar
$uProgress.Height = 6
$uProgress.Dock   = [System.Windows.Forms.DockStyle]::Top
$uProgress.Visible = $false
$uProgress.MarqueeAnimationSpeed = 30

# Dock order: Fill grid first, then bottom panels (action outermost, pagination
# above), then top-docked progress bar + toolbar, then the header last so it
# claims the top edge.
$form.Controls.Add($uGrid)
$form.Controls.Add($uActionPanel)
$form.Controls.Add($uPaginPanel)
$form.Controls.Add($uProgress)
$form.Controls.Add($uToolbar)
$form.Controls.Add($headerPanel)

if (-not $script:ExoModuleLoaded) {
    $uBtnConnect.Enabled = $false
    $uStatusLabel.Text   = "ExchangeOnlineManagement not loaded - see log"
}

# ---- Busy-state helpers ----
function Start-Busy {
    param([string]$Text, [int]$Maximum = 0)
    $script:BusyWatch = [System.Diagnostics.Stopwatch]::StartNew()
    if ($Maximum -gt 0) {
        $uProgress.Style    = [System.Windows.Forms.ProgressBarStyle]::Continuous
        $uProgress.Minimum  = 0
        $uProgress.Maximum  = $Maximum
        $uProgress.Value    = 0
    } else {
        $uProgress.Style    = [System.Windows.Forms.ProgressBarStyle]::Marquee
    }
    $uProgress.Visible   = $true
    $form.UseWaitCursor  = $true
    $uSearchBox.Enabled  = $false
    $uChkHide.Enabled    = $false
    $uStatusLabel.Text   = $Text
    $uStatusLabel.Refresh()
    [System.Windows.Forms.Application]::DoEvents()
}

function Update-Busy {
    param([string]$Text, [int]$Value = -1)
    if ($Value -ge 0 -and $uProgress.Style -eq [System.Windows.Forms.ProgressBarStyle]::Continuous) {
        $uProgress.Value = [Math]::Min($Value, $uProgress.Maximum)
    }
    $uStatusLabel.Text = "$Text  ($([int]$script:BusyWatch.Elapsed.TotalSeconds)s)"
    $uStatusLabel.Refresh()
    [System.Windows.Forms.Application]::DoEvents()
}

function Stop-Busy {
    $uProgress.Visible  = $false
    $form.UseWaitCursor = $false
    $uSearchBox.Enabled = $true
    $uChkHide.Enabled   = $true
    if ($script:BusyWatch) { $script:BusyWatch.Stop() }
}

# ---- Users grid update helper ----
function Update-UserGrid {
    $display = Get-FilteredData `
        -Source              $script:AllUsersUnfiltered `
        -SearchText          $uSearchBox.Text `
        -HideConverted       $script:UsersHideConverted `
        -ConvertedPropertyName "IsExchangeCloudManaged" `
        -EmailPropertyName   "PrimarySmtpAddress"

    if ($script:UsersSortColumn -ne "") {
        $sortProp = if ($script:UsersSortColumn -eq "Email") { "PrimarySmtpAddress" } else { $script:UsersSortColumn }
        $display = if ($script:UsersSortDirection -eq "Ascending") {
            $display | Sort-Object -Property $sortProp
        } else {
            $display | Sort-Object -Property $sortProp -Descending
        }
    }

    $script:AllUsers = @($display)
    $total      = $script:AllUsers.Count
    $totalPages = [Math]::Ceiling($total / $script:PageSize)
    if ($totalPages -lt 1) { $totalPages = 1 }
    if ($script:UsersCurrentPage -gt $totalPages) { $script:UsersCurrentPage = $totalPages }
    if ($script:UsersCurrentPage -lt 1)           { $script:UsersCurrentPage = 1 }

    $startIdx = ($script:UsersCurrentPage - 1) * $script:PageSize
    $endIdx   = [Math]::Min($startIdx + $script:PageSize - 1, $total - 1)

    $uGrid.Rows.Clear()
    if ($total -gt 0) {
        for ($i = $startIdx; $i -le $endIdx; $i++) {
            $u   = $script:AllUsers[$i]
            $val = if ($u.IsExchangeCloudManaged) { "True" } else { "False" }
            $uGrid.Rows.Add($u.DisplayName, $u.PrimarySmtpAddress, $val, $u.UserPrincipalName)
        }
        $uPageInfo.Text = "Page $($script:UsersCurrentPage) of $totalPages  -  Showing $($endIdx - $startIdx + 1) of $total users"
    } else {
        $uPageInfo.Text = "No users to display"
    }
    $uBtnPrev.Enabled = ($script:UsersCurrentPage -gt 1)
    $uBtnNext.Enabled = ($script:UsersCurrentPage -lt $totalPages)
}

# ---- Users event wiring ----
$uBtnConnect.Add_Click({
    $uBtnConnect.Enabled    = $false
    $uBtnRefresh.Enabled    = $false
    $uBtnDisconnect.Enabled = $false
    $uBtnToCloud.Enabled    = $false
    $uBtnToOnPrem.Enabled   = $false
    try {
        Start-Busy "Waiting for sign-in..."
        if (Connect-EXOSession) {
            $script:ExoConnected   = $true
            $uBtnConnect.Text      = "Connected"
            $uBtnConnect.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#107C10")
            Update-Busy "Loading mailboxes..."
            $users = Get-ExchangeUsers
            if ($null -ne $users) {
                $script:AllUsersUnfiltered = $users
                $script:UsersCurrentPage   = 1
                Update-UserGrid
                $uStatusLabel.Text = "Connected  -  $($users.Count) mailboxes loaded"
            } else { $uStatusLabel.Text = "Connected but failed to load users" }
        }
    } finally {
        Stop-Busy
        $uBtnConnect.Enabled    = (-not $script:ExoConnected) -and $script:ExoModuleLoaded
        $uBtnRefresh.Enabled    = $script:ExoConnected
        $uBtnDisconnect.Enabled = $script:ExoConnected
        $uBtnToCloud.Enabled    = $script:ExoConnected
        $uBtnToOnPrem.Enabled   = $script:ExoConnected
    }
})

$uBtnRefresh.Add_Click({
    $uBtnConnect.Enabled    = $false
    $uBtnRefresh.Enabled    = $false
    $uBtnDisconnect.Enabled = $false
    $uBtnToCloud.Enabled    = $false
    $uBtnToOnPrem.Enabled   = $false
    try {
        Start-Busy "Refreshing mailboxes..."
        $users = Get-ExchangeUsers
        if ($null -ne $users) {
            $script:AllUsersUnfiltered = $users
            $script:UsersCurrentPage   = 1
            Update-UserGrid
            $uStatusLabel.Text = "Refreshed  -  $($users.Count) mailboxes loaded"
        } else { $uStatusLabel.Text = "Refresh failed - see log" }
    } finally {
        Stop-Busy
        $uBtnConnect.Enabled    = (-not $script:ExoConnected) -and $script:ExoModuleLoaded
        $uBtnRefresh.Enabled    = $script:ExoConnected
        $uBtnDisconnect.Enabled = $script:ExoConnected
        $uBtnToCloud.Enabled    = $script:ExoConnected
        $uBtnToOnPrem.Enabled   = $script:ExoConnected
    }
})

$uBtnDisconnect.Add_Click({
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
        Write-Log "Disconnected from Exchange Online."
    } catch {
        Write-Log "Disconnect warning: $($_.Exception.Message)" -Level WARNING
    }
    $script:ExoConnected    = $false
    $uBtnConnect.Text      = "Connect to EXO"
    $uBtnConnect.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#0078D4")
    $uBtnConnect.Enabled   = $script:ExoModuleLoaded
    $uBtnRefresh.Enabled   = $false
    $uBtnDisconnect.Enabled = $false
    $uBtnToCloud.Enabled    = $false
    $uBtnToOnPrem.Enabled   = $false
    $script:AllUsersUnfiltered = @()
    $script:AllUsers           = @()
    $uGrid.Rows.Clear()
    $uPageInfo.Text    = "No users loaded"
    $uStatusLabel.Text = "Disconnected"
    [void][System.Windows.Forms.MessageBox]::Show("Disconnected from Exchange Online.", "Disconnected",
        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

$uSearchBox.Add_TextChanged({
    $script:UsersCurrentPage = 1
    Update-UserGrid
})

$uChkHide.Add_CheckedChanged({
    $script:UsersHideConverted = $uChkHide.Checked
    $script:UsersCurrentPage   = 1
    Update-UserGrid
})

$uBtnPrev.Add_Click({
    if ($script:UsersCurrentPage -gt 1) { $script:UsersCurrentPage--; Update-UserGrid }
})
$uBtnNext.Add_Click({
    $tp = [Math]::Ceiling($script:AllUsers.Count / $script:PageSize)
    if ($script:UsersCurrentPage -lt $tp) { $script:UsersCurrentPage++; Update-UserGrid }
})

$uGrid.Add_ColumnHeaderMouseClick({
    param($dgv, $e)
    $col = $dgv.Columns[$e.ColumnIndex].Name
    if ($col -eq "UserPrincipalName") { return }
    if ($script:UsersSortColumn -eq $col) {
        $script:UsersSortDirection = if ($script:UsersSortDirection -eq "Ascending") { "Descending" } else { "Ascending" }
    } else {
        $script:UsersSortColumn    = $col
        $script:UsersSortDirection = "Ascending"
    }
    $script:UsersCurrentPage = 1
    Update-UserGrid
})

$uBtnToCloud.Add_Click({
    if ($uGrid.SelectedRows.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show("Please select at least one user.", "No Selection",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $cnt  = $uGrid.SelectedRows.Count
    $names = ($uGrid.SelectedRows | ForEach-Object { $_.Cells["DisplayName"].Value }) -join ", "
    $msg  = if ($cnt -eq 1) { "Convert '$names' to Cloud Managed?" } else { "Convert $cnt users to Cloud Managed?`n`n$names" }
    if ([System.Windows.Forms.MessageBox]::Show($msg, "Confirm", [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question) -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $uBtnConnect.Enabled    = $false
    $uBtnRefresh.Enabled    = $false
    $uBtnDisconnect.Enabled = $false
    $uBtnToCloud.Enabled    = $false
    $uBtnToOnPrem.Enabled   = $false
    Start-Busy "Converting..." -Maximum $cnt
    try {
        $ok = 0; $fail = 0; $i = 0
        foreach ($row in $uGrid.SelectedRows) {
            $i++
            $u = @{ DisplayName = $row.Cells["DisplayName"].Value; UserPrincipalName = $row.Cells["UserPrincipalName"].Value }
            Update-Busy "Converting $i of $cnt : $($u.DisplayName)" -Value $i
            if (Convert-UserToCloud -User $u) {
                $row.Cells["IsExchangeCloudManaged"].Value = "True"
                $match = $script:AllUsersUnfiltered | Where-Object { $_.UserPrincipalName -eq $u.UserPrincipalName }
                if ($match) { $match.IsExchangeCloudManaged = $true }
                $ok++
            } else { $fail++ }
        }
        Stop-Busy
        [void][System.Windows.Forms.MessageBox]::Show("Done.`n`nConverted: $ok`nFailed: $fail", "Batch Complete",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $uStatusLabel.Text = "Converted: $ok  Failed: $fail"
        Update-UserGrid
    } finally {
        Stop-Busy
        $uBtnConnect.Enabled    = (-not $script:ExoConnected) -and $script:ExoModuleLoaded
        $uBtnRefresh.Enabled    = $script:ExoConnected
        $uBtnDisconnect.Enabled = $script:ExoConnected
        $uBtnToCloud.Enabled    = $script:ExoConnected
        $uBtnToOnPrem.Enabled   = $script:ExoConnected
    }
})

$uBtnToOnPrem.Add_Click({
    if ($uGrid.SelectedRows.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show("Please select at least one user.", "No Selection",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $cnt  = $uGrid.SelectedRows.Count
    $names = ($uGrid.SelectedRows | ForEach-Object { $_.Cells["DisplayName"].Value }) -join ", "
    $msg  = if ($cnt -eq 1) { "Roll back '$names' to On-Prem Managed?" } else { "Roll back $cnt users to On-Prem Managed?`n`n$names" }
    if ([System.Windows.Forms.MessageBox]::Show($msg, "Confirm", [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question) -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $uBtnConnect.Enabled    = $false
    $uBtnRefresh.Enabled    = $false
    $uBtnDisconnect.Enabled = $false
    $uBtnToCloud.Enabled    = $false
    $uBtnToOnPrem.Enabled   = $false
    Start-Busy "Rolling back..." -Maximum $cnt
    try {
        $ok = 0; $fail = 0; $i = 0
        foreach ($row in $uGrid.SelectedRows) {
            $i++
            $u = @{ DisplayName = $row.Cells["DisplayName"].Value; UserPrincipalName = $row.Cells["UserPrincipalName"].Value }
            Update-Busy "Rolling back $i of $cnt : $($u.DisplayName)" -Value $i
            if (Convert-UserToOnPrem -User $u) {
                $row.Cells["IsExchangeCloudManaged"].Value = "False"
                $match = $script:AllUsersUnfiltered | Where-Object { $_.UserPrincipalName -eq $u.UserPrincipalName }
                if ($match) { $match.IsExchangeCloudManaged = $false }
                $ok++
            } else { $fail++ }
        }
        Stop-Busy
        [void][System.Windows.Forms.MessageBox]::Show("Done.`n`nRolled back: $ok`nFailed: $fail", "Batch Complete",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $uStatusLabel.Text = "Rolled back: $ok  Failed: $fail"
        Update-UserGrid
    } finally {
        Stop-Busy
        $uBtnConnect.Enabled    = (-not $script:ExoConnected) -and $script:ExoModuleLoaded
        $uBtnRefresh.Enabled    = $script:ExoConnected
        $uBtnDisconnect.Enabled = $script:ExoConnected
        $uBtnToCloud.Enabled    = $script:ExoConnected
        $uBtnToOnPrem.Enabled   = $script:ExoConnected
    }
})

$uBtnOpenLog.Add_Click({
    if (Test-Path $script:LogFile) { Start-Process notepad.exe -ArgumentList $script:LogFile }
    else { [void][System.Windows.Forms.MessageBox]::Show("Log file not found yet.", "Log", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) }
})

# ==============================================================================
# FORM CLOSING HANDLER
# ==============================================================================

$form.Add_FormClosing({
    Write-Log "Exchange SOA Conversion Tool closing..."
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}
    if ($pictureBoxLogo.Image) { $pictureBoxLogo.Image.Dispose() }
    Write-Log "========================================"
    Write-Log "Exchange SOA Conversion Tool Ended"
    Write-Log "========================================"
})

# ==============================================================================
# SHOW FORM
# ==============================================================================

# Open as a normal centered window at the design size (like the original tools).
# Docking guarantees the grid always fits its container, so no maximize is needed.
# Only clamp the size down if the screen is smaller than the design window.
$workingArea = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
if ($form.Width -gt $workingArea.Width -or $form.Height -gt $workingArea.Height) {
    $fitWidth  = [Math]::Min($form.Width,  $workingArea.Width  - 40)
    $fitHeight = [Math]::Min($form.Height, $workingArea.Height - 40)
    $form.Size = New-Object System.Drawing.Size($fitWidth, $fitHeight)
    Write-Log "Window clamped to fit screen working area ($($workingArea.Width)x$($workingArea.Height))."
}

[void]$form.ShowDialog()
