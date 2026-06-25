#Requires -Version 5.1

<#
        .SYNOPSIS
        Exchange SOA Conversion Tool - Combined Source of Authority management for Exchange users, Groups and Contacts

        .DESCRIPTION
        GUI tool to manage Source of Authority (SOA) conversion for:
          - Exchange mailboxes (IsDirSynced users) via ExchangeOnlineManagement
          - Mail-enabled Security Groups and Distribution Groups via Microsoft Graph
          - Org Contacts (mail contacts) via Microsoft Graph

        Features:
        - Three tabs: Users, Groups, Contacts
        - Live search by display name or email address across all loaded data
        - Per-tab Hide Converted filter
        - Pagination (100 per page)
        - Batch conversion with confirmation
        - Nested group detection and bottom-up conversion ordering (Groups tab)
        - Full logging to a single session log file

        .PARAMETER TenantId
        Optional. The Entra ID tenant ID (GUID). Recommended in multi-tenant / partner scenarios.

        .EXAMPLE
        .\Exchange-SOA-Conversion-Tool.ps1
        .\Exchange-SOA-Conversion-Tool.ps1 -TenantId "00000000-0000-0000-0000-000000000000"

        .NOTES
        Version: 1.00

        .LINK
        https://learn.microsoft.com/en-us/exchange/hybrid-deployment/enable-exchange-attributes-cloud-management
        https://learn.microsoft.com/en-us/entra/identity/hybrid/how-to-group-source-of-authority-configure
        https://learn.microsoft.com/en-us/entra/identity/hybrid/how-to-user-source-of-authority-configure#configure-contact-soa

        .COPYRIGHT
        MIT License, feel free to distribute and use as you like, please leave author information.

        BLOG: http://www.hcandersen.net
        LinkedIn: https://www.linkedin.com/in/hanschrandersen/

        .DISCLAIMER
        This script is provided AS-IS, with no warranty - Use at own risk.
    #>

param(
    [Parameter(Mandatory=$false, HelpMessage="Enter the Entra ID tenant ID (GUID) to connect to.")]
    [string]$TenantId
)

$script:Version       = "1.00"
$script:TenantId      = $TenantId

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:ScriptPath    = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:LogFile       = Join-Path $script:ScriptPath "SOAConversion_$(Get-Date -Format 'yyyyMMdd_HHmm').log"

# ---- Per-tab state ----
# Users
$script:AllUsers             = @()
$script:AllUsersUnfiltered   = @()
$script:UsersCurrentPage     = 1
$script:UsersHideConverted   = $false
$script:UsersSortColumn      = ""
$script:UsersSortDirection   = "Ascending"

# Groups
$script:AllGroups            = @()
$script:AllGroupsUnfiltered  = @()
$script:GroupsCurrentPage    = 1
$script:GroupsHideConverted  = $false
$script:GroupsSortColumn     = ""
$script:GroupsSortDirection  = "Ascending"
$script:PermissionOk         = $false
$script:NestingMap           = @{}
$script:NestingDepth         = @{}

# Contacts
$script:AllContacts           = @()
$script:AllContactsUnfiltered = @()
$script:ContactsCurrentPage   = 1
$script:ContactsHideConverted = $false
$script:ContactsSortColumn    = ""
$script:ContactsSortDirection = "Ascending"

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

function Test-ExchangeModule {
    $module = Get-Module -ListAvailable -Name ExchangeOnlineManagement
    if (-not $module) {
        Write-Log "ExchangeOnlineManagement module not found. Attempting install..." -Level WARNING
        try {
            [System.Windows.Forms.MessageBox]::Show(
                "ExchangeOnlineManagement module is not installed.`nThe tool will now attempt to install it.",
                "Module Required", [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information)
            Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            Write-Log "ExchangeOnlineManagement installed successfully."
            return $true
        } catch {
            Write-Log "Failed to install ExchangeOnlineManagement: $($_.Exception.Message)" -Level ERROR
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to install ExchangeOnlineManagement.`n`nError: $($_.Exception.Message)`n`nInstall manually:`nInstall-Module -Name ExchangeOnlineManagement -Scope CurrentUser",
                "Installation Failed", [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error)
            return $false
        }
    }
    Write-Log "ExchangeOnlineManagement found (v$($module.Version))."
    return $true
}

function Test-GraphModule {
    $module = Get-Module -ListAvailable -Name Microsoft.Graph.Groups
    if (-not $module) {
        Write-Log "Microsoft.Graph.Groups module not found. Attempting install..." -Level WARNING
        try {
            [System.Windows.Forms.MessageBox]::Show(
                "Microsoft.Graph.Groups module is not installed.`nThe tool will now attempt to install it.",
                "Module Required", [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information)
            Install-Module -Name Microsoft.Graph.Groups -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            Write-Log "Microsoft.Graph.Groups installed successfully."
            return $true
        } catch {
            Write-Log "Failed to install Microsoft.Graph.Groups: $($_.Exception.Message)" -Level ERROR
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to install Microsoft.Graph.Groups.`n`nError: $($_.Exception.Message)`n`nInstall manually:`nInstall-Module -Name Microsoft.Graph.Groups -Scope CurrentUser",
                "Installation Failed", [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error)
            return $false
        }
    }
    Write-Log "Microsoft.Graph.Groups found (v$($module.Version))."
    return $true
}

# ==============================================================================
# USERS TAB -- BACKEND
# ==============================================================================

function Connect-EXOSession {
    Write-Log "Connecting to Exchange Online..."
    try {
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        Connect-ExchangeOnline -ErrorAction Stop -ShowBanner:$false
        Write-Log "Connected to Exchange Online."
        return $true
    } catch {
        Write-Log "EXO connect failed: $($_.Exception.Message)" -Level ERROR
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to connect to Exchange Online.`n`nError: $($_.Exception.Message)",
            "Connection Failed", [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
        return $false
    }
}

function Get-ExchangeUsers {
    Write-Log "Retrieving mailboxes from Exchange Online..."
    try {
        $all = Get-Mailbox -ResultSize Unlimited -ErrorAction Stop |
               Select-Object DisplayName, PrimarySmtpAddress, IsExchangeCloudManaged, UserPrincipalName, IsDirSynced
        Write-Log "Retrieved $($all.Count) total mailboxes."
        $hybrid = $all | Where-Object { $_.IsDirSynced -eq $true }
        Write-Log "Filtered to $($hybrid.Count) IsDirSynced mailboxes."
        return @($hybrid)
    } catch {
        Write-Log "Failed to retrieve mailboxes: $($_.Exception.Message)" -Level ERROR
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to retrieve mailboxes.`n`nError: $($_.Exception.Message)",
            "Retrieval Failed", [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
        return $null
    }
}

function Convert-UserToCloud {
    param($User)
    Write-Log "Converting user '$($User.DisplayName)' to Cloud Managed..."
    try {
        Set-Mailbox -Identity $User.UserPrincipalName -IsExchangeCloudManaged $true -ErrorAction Stop
        Write-Log "Converted '$($User.DisplayName)' to Cloud Managed." -Level INFO
        return $true
    } catch {
        Write-Log "Failed to convert '$($User.DisplayName)': $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

function Convert-UserToOnPrem {
    param($User)
    Write-Log "Converting user '$($User.DisplayName)' to On-Prem Managed..."
    try {
        Set-Mailbox -Identity $User.UserPrincipalName -IsExchangeCloudManaged $false -ErrorAction Stop
        Write-Log "Converted '$($User.DisplayName)' to On-Prem Managed." -Level INFO
        return $true
    } catch {
        Write-Log "Failed to roll back '$($User.DisplayName)': $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

# ==============================================================================
# GROUPS TAB -- BACKEND
# ==============================================================================

function Test-GraphGroupPermission {
    try {
        $ctx = Get-MgContext
        if (-not $ctx) { $script:PermissionOk = $false; return $false }
        if ($ctx.Scopes -contains 'Group-OnPremisesSyncBehavior.ReadWrite.All') {
            $script:PermissionOk = $true; return $true
        }
        $script:PermissionOk = $false; return $false
    } catch {
        $script:PermissionOk = $false; return $false
    }
}

function Connect-GraphSession {
    $tenantMsg = if ($script:TenantId) { "TenantId: $($script:TenantId)" } else { "default tenant" }
    Write-Log "Connecting to Microsoft Graph ($tenantMsg)..."
    try {
        # Remove any previously loaded Microsoft.Graph modules to avoid assembly-version conflicts
        # (common when multiple versions are installed side-by-side).
        Get-Module 'Microsoft.Graph.*' | Remove-Module -Force -ErrorAction SilentlyContinue

        $graphModule = Get-Module -ListAvailable -Name Microsoft.Graph.Groups |
                       Sort-Object Version -Descending | Select-Object -First 1
        if (-not $graphModule) { throw "Microsoft.Graph.Groups module is not available." }
        $graphVersion = $graphModule.Version
        Write-Log "Importing Microsoft.Graph modules version $graphVersion..."
        Import-Module Microsoft.Graph.Authentication -RequiredVersion $graphVersion -Force -ErrorAction Stop
        Import-Module Microsoft.Graph.Groups      -RequiredVersion $graphVersion -Force -ErrorAction Stop
        Import-Module Microsoft.Graph.Identity.DirectoryManagement -RequiredVersion $graphVersion -Force -ErrorAction SilentlyContinue

        $scopes = @(
            'Group.ReadWrite.All',
            'Group-OnPremisesSyncBehavior.ReadWrite.All',
            'OrgContact.Read.All',
            'Contacts-OnPremisesSyncBehavior.ReadWrite.All'
        )
        $connectParams = @{ Scopes = $scopes; ErrorAction = 'Stop'; NoWelcome = $true }
        if ($script:TenantId) { $connectParams['TenantId'] = $script:TenantId }

        Connect-MgGraph @connectParams

        $ctx = Get-MgContext
        Write-Log "Connected to Graph. TenantId: $($ctx.TenantId)"

        if (-not (Test-GraphGroupPermission)) {
            Write-Log "Permission missing -- triggering consent re-flow..." -Level WARNING
            Disconnect-MgGraph -ErrorAction SilentlyContinue
            Connect-MgGraph @connectParams
            if (-not (Test-GraphGroupPermission)) {
                Write-Log "Required permission still not granted after consent." -Level WARNING
                [System.Windows.Forms.MessageBox]::Show(
                    "The permission 'Group-OnPremisesSyncBehavior.ReadWrite.All' was not granted.`n`nThis may require admin consent.`n`nTo grant manually:`n1. Entra admin center`n2. Enterprise Applications`n3. 'Microsoft Graph Command Line Tools'`n4. Permissions -> Grant admin consent",
                    "Permission Not Granted", [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning)
            }
        }
        return $true
    } catch {
        Write-Log "Graph connect failed: $($_.Exception.Message)" -Level ERROR
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to connect to Microsoft Graph.`n`nError: $($_.Exception.Message)",
            "Connection Failed", [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
        return $false
    }
}

function Get-ExchangeGroups {
    param($StatusLabelRef)
    Write-Log "Retrieving Exchange-relevant groups from Graph..."
    try {
        $allGroups = Get-MgGroup -All -Property Id,DisplayName,Mail,MailEnabled,SecurityEnabled,GroupTypes,OnPremisesSyncEnabled -ErrorAction Stop
        Write-Log "Retrieved $($allGroups.Count) total groups."

        $exchangeGroups = $allGroups | Where-Object {
            $_.MailEnabled -eq $true -and
            ($null -eq $_.GroupTypes -or $_.GroupTypes.Count -eq 0 -or $_.GroupTypes -notcontains "Unified")
        }
        Write-Log "Filtered to $($exchangeGroups.Count) Exchange-relevant groups."

        $results = @()
        $total   = $exchangeGroups.Count
        $idx     = 0

        foreach ($group in $exchangeGroups) {
            $idx++
            if ($idx % 10 -eq 0 -or $idx -eq $total) {
                if ($StatusLabelRef) { $StatusLabelRef.Text = "Loading group $idx of $total..." }
                [System.Windows.Forms.Application]::DoEvents()
            }

            $groupType    = if ($group.SecurityEnabled) { "Mail-Enabled Security Group" } else { "Distribution Group" }
            $isCloud      = $null
            for ($r = 0; $r -lt 3; $r++) {
                try {
                    $uri = "https://graph.microsoft.com/v1.0/groups/$($group.Id)/onPremisesSyncBehavior?`$select=isCloudManaged"
                    $resp = Invoke-MgGraphRequest -Uri $uri -Method GET -ErrorAction Stop
                    $isCloud = if ($null -ne $resp.isCloudManaged) { [bool]$resp.isCloudManaged } else { "Unknown" }
                    break
                } catch {
                    if ($r -lt 2) { Start-Sleep -Milliseconds 500 }
                    else          { $isCloud = "Unknown" }
                }
            }

            $results += [PSCustomObject]@{
                Id             = $group.Id
                DisplayName    = $group.DisplayName
                Mail           = $group.Mail
                GroupType      = $groupType
                IsCloudManaged = $isCloud
                SecurityEnabled = $group.SecurityEnabled
                MailEnabled    = $group.MailEnabled
            }
        }
        Write-Log "Finished loading SOA status for $($results.Count) groups."
        return $results
    } catch {
        Write-Log "Failed to retrieve groups: $($_.Exception.Message)" -Level ERROR
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to retrieve groups.`n`nError: $($_.Exception.Message)",
            "Retrieval Failed", [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
        return $null
    }
}

function Build-NestingMap {
    param([array]$Groups, $StatusLabelRef)
    Write-Log "Building nested group map..."
    $script:NestingMap   = @{}
    $script:NestingDepth = @{}
    $groupIds = @{}
    foreach ($g in $Groups) { $groupIds[$g.Id] = $g; $script:NestingMap[$g.Id] = @() }

    $total = $Groups.Count; $idx = 0
    foreach ($g in $Groups) {
        $idx++
        if ($idx % 10 -eq 0 -or $idx -eq $total) {
            if ($StatusLabelRef) { $StatusLabelRef.Text = "Analyzing nesting $idx of $total..." }
            [System.Windows.Forms.Application]::DoEvents()
        }
        try {
            $members = Get-MgGroupMember -GroupId $g.Id -All -Property Id -ErrorAction Stop
            foreach ($m in $members) {
                if ($groupIds.ContainsKey($m.Id)) { $script:NestingMap[$g.Id] += $m.Id }
            }
        } catch {
            Write-Log "Could not retrieve members for '$($g.DisplayName)': $($_.Exception.Message)" -Level WARNING
        }
    }

    foreach ($g in $Groups) { $script:NestingDepth[$g.Id] = 0 }
    $changed = $true; $iter = 0
    while ($changed -and $iter -lt 100) {
        $changed = $false; $iter++
        foreach ($parentId in $script:NestingMap.Keys) {
            foreach ($cid in $script:NestingMap[$parentId]) {
                $exp = $script:NestingDepth[$cid] + 1
                if ($exp -gt $script:NestingDepth[$parentId]) { $script:NestingDepth[$parentId] = $exp; $changed = $true }
            }
        }
    }
    $maxDepth = ($script:NestingDepth.Values | Measure-Object -Maximum).Maximum
    Write-Log "Nesting analysis complete. Max depth: $maxDepth."
}

function Get-UnconvertedChildren {
    param([string]$GroupId)
    $unconverted = @()
    if ($script:NestingMap.ContainsKey($GroupId)) {
        foreach ($cid in $script:NestingMap[$GroupId]) {
            $child = $script:AllGroups | Where-Object { $_.Id -eq $cid }
            if ($child -and $child.IsCloudManaged -ne $true) { $unconverted += $child }
        }
    }
    return $unconverted
}

function Convert-GroupToCloud {
    param($Group)
    Write-Log "Converting group '$($Group.DisplayName)' to Cloud Managed..."
    try {
        $body = @{ isCloudManaged = $true } | ConvertTo-Json
        Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/groups/$($Group.Id)/onPremisesSyncBehavior" -Method PATCH -Body $body -ContentType "application/json" -ErrorAction Stop
        Write-Log "Converted group '$($Group.DisplayName)'." -Level INFO
        return $true
    } catch {
        Write-Log "Failed to convert group '$($Group.DisplayName)': $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

function Convert-GroupToOnPrem {
    param($Group)
    Write-Log "Rolling back group '$($Group.DisplayName)' to On-Prem..."
    try {
        $body = @{ isCloudManaged = $false } | ConvertTo-Json
        Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/groups/$($Group.Id)/onPremisesSyncBehavior" -Method PATCH -Body $body -ContentType "application/json" -ErrorAction Stop
        Write-Log "Rolled back group '$($Group.DisplayName)'." -Level INFO
        return $true
    } catch {
        Write-Log "Failed to roll back group '$($Group.DisplayName)': $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

# ==============================================================================
# CONTACTS TAB -- BACKEND
# ==============================================================================

function Get-OrgContacts {
    param($StatusLabelRef)
    Write-Log "Retrieving org contacts from Graph..."
    try {
        $allContacts = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/contacts?`$select=id,displayName,mail,onPremisesSyncEnabled&`$top=999" -Method GET -ErrorAction Stop
        $contacts = [System.Collections.Generic.List[object]]::new()
        $page = $allContacts
        while ($page) {
            foreach ($c in $page.value) { $contacts.Add($c) }
            if ($page.'@odata.nextLink') {
                $page = Invoke-MgGraphRequest -Uri $page.'@odata.nextLink' -Method GET -ErrorAction Stop
            } else { break }
        }
        Write-Log "Retrieved $($contacts.Count) org contacts."

        $results = @()
        $total   = $contacts.Count
        $idx     = 0

        foreach ($contact in $contacts) {
            $idx++
            if ($idx % 20 -eq 0 -or $idx -eq $total) {
                if ($StatusLabelRef) { $StatusLabelRef.Text = "Loading contact $idx of $total..." }
                [System.Windows.Forms.Application]::DoEvents()
            }

            $isCloud = "N/A"
            if ($contact.onPremisesSyncEnabled -eq $true) {
                for ($r = 0; $r -lt 3; $r++) {
                    try {
                        $uri  = "https://graph.microsoft.com/v1.0/contacts/$($contact.id)/onPremisesSyncBehavior?`$select=isCloudManaged"
                        $resp = Invoke-MgGraphRequest -Uri $uri -Method GET -ErrorAction Stop
                        $isCloud = if ($null -ne $resp.isCloudManaged) { [bool]$resp.isCloudManaged } else { "Unknown" }
                        break
                    } catch {
                        if ($r -lt 2) { Start-Sleep -Milliseconds 300 }
                        else          { $isCloud = "Unknown" }
                    }
                }
            }

            $results += [PSCustomObject]@{
                Id                   = $contact.id
                DisplayName          = $contact.displayName
                Mail                 = $contact.mail
                OnPremisesSyncEnabled = $contact.onPremisesSyncEnabled
                IsCloudManaged       = $isCloud
            }
        }
        Write-Log "Finished loading SOA status for $($results.Count) contacts."
        return $results
    } catch {
        Write-Log "Failed to retrieve org contacts: $($_.Exception.Message)" -Level ERROR
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to retrieve org contacts.`n`nError: $($_.Exception.Message)`n`nNote: 'Contacts-OnPremisesSyncBehavior.ReadWrite.All' may require admin consent.",
            "Retrieval Failed", [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
        return $null
    }
}

function Convert-ContactToCloud {
    param($Contact)
    Write-Log "Converting contact '$($Contact.DisplayName)' to Cloud Managed..."
    try {
        $body = @{ isCloudManaged = $true } | ConvertTo-Json
        Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/contacts/$($Contact.Id)/onPremisesSyncBehavior" -Method PATCH -Body $body -ContentType "application/json" -ErrorAction Stop
        Write-Log "Converted contact '$($Contact.DisplayName)'." -Level INFO
        return $true
    } catch {
        Write-Log "Failed to convert contact '$($Contact.DisplayName)': $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

function Convert-ContactToOnPrem {
    param($Contact)
    Write-Log "Rolling back contact '$($Contact.DisplayName)' to On-Prem..."
    try {
        $body = @{ isCloudManaged = $false } | ConvertTo-Json
        Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/contacts/$($Contact.Id)/onPremisesSyncBehavior" -Method PATCH -Body $body -ContentType "application/json" -ErrorAction Stop
        Write-Log "Rolled back contact '$($Contact.DisplayName)'." -Level INFO
        return $true
    } catch {
        Write-Log "Failed to roll back contact '$($Contact.DisplayName)': $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

Write-Log "========================================"
Write-Log "Exchange SOA Conversion Tool v$script:Version Started"
Write-Log "========================================"

if (-not (Test-ExchangeModule)) {
    Write-Log "ExchangeOnlineManagement module missing. Users tab may not function." -Level WARNING
}
if (-not (Test-GraphModule)) {
    Write-Log "Microsoft.Graph.Groups module missing. Groups/Contacts tabs may not function." -Level WARNING
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
$labelDesc.Text      = "Manage Source of Authority for Exchange mailboxes, Groups, and Contacts"
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

# --- TabControl ---
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$tabControl.Padding = New-Object System.Drawing.Point(12, 4)
$tabControl.Dock = [System.Windows.Forms.DockStyle]::Fill

# Add the Fill control first, then the Top-docked header, so docking resolves
# correctly (the header claims the top edge, the tab control fills the rest).
$form.Controls.Add($tabControl)
$form.Controls.Add($headerPanel)

$tabUsers    = New-Object System.Windows.Forms.TabPage; $tabUsers.Text    = "  Users (Mailboxes)  ";    $tabUsers.BackColor    = [System.Drawing.ColorTranslator]::FromHtml("#F3F3F3")
$tabGroups   = New-Object System.Windows.Forms.TabPage; $tabGroups.Text   = "  Groups  ";               $tabGroups.BackColor   = [System.Drawing.ColorTranslator]::FromHtml("#F3F3F3")
$tabContacts = New-Object System.Windows.Forms.TabPage; $tabContacts.Text = "  Contacts  ";             $tabContacts.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#F3F3F3")
$tabControl.TabPages.Add($tabUsers)
$tabControl.TabPages.Add($tabGroups)
$tabControl.TabPages.Add($tabContacts)

# ==============================================================================
# USERS TAB -- GUI
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

$uBtnToCloud  = New-StyledButton -Text "Convert to Cloud Managed"   -X 0   -Y 4 -Width 220 -Height 40
$uBtnToOnPrem = New-StyledButton -Text "Convert to On-Prem Managed" -X 228 -Y 4 -Width 230 -Height 40
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

# Dock order: Fill grid first, then bottom panels (action outermost, pagination above), then top toolbar.
$tabUsers.Controls.Add($uGrid)
$tabUsers.Controls.Add($uActionPanel)
$tabUsers.Controls.Add($uPaginPanel)
$tabUsers.Controls.Add($uToolbar)

# ---- Users grid update helper ----
function Update-UserGrid {
    $display = Get-FilteredData `
        -Source              $script:AllUsersUnfiltered `
        -SearchText          $uSearchBox.Text `
        -HideConverted       $script:UsersHideConverted `
        -ConvertedPropertyName "IsExchangeCloudManaged" `
        -EmailPropertyName   "PrimarySmtpAddress"

    if ($script:UsersSortColumn -ne "") {
        $display = if ($script:UsersSortDirection -eq "Ascending") {
            $display | Sort-Object -Property $script:UsersSortColumn
        } else {
            $display | Sort-Object -Property $script:UsersSortColumn -Descending
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
    $uBtnConnect.Enabled = $false
    if (Connect-EXOSession) {
        $uBtnConnect.Text     = "Connected"
        $uBtnConnect.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#107C10")
        $uBtnRefresh.Enabled    = $true
        $uBtnDisconnect.Enabled = $true
        $uStatusLabel.Text = "Loading mailboxes..."
        $users = Get-ExchangeUsers
        if ($users) {
            $script:AllUsersUnfiltered = $users
            $script:UsersCurrentPage   = 1
            Update-UserGrid
            $uStatusLabel.Text = "Connected  -  $($users.Count) mailboxes loaded"
        } else { $uStatusLabel.Text = "Connected but failed to load users" }
    } else {
        $uBtnConnect.Enabled = $true
    }
})

$uBtnRefresh.Add_Click({
    $uBtnRefresh.Enabled = $false
    $uStatusLabel.Text = "Refreshing..."
    $users = Get-ExchangeUsers
    if ($users) {
        $script:AllUsersUnfiltered = $users
        $script:UsersCurrentPage   = 1
        Update-UserGrid
        $uStatusLabel.Text = "Refreshed  -  $($users.Count) mailboxes loaded"
    }
    $uBtnRefresh.Enabled = $true
})

$uBtnDisconnect.Add_Click({
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
        Write-Log "Disconnected from Exchange Online."
    } catch {
        Write-Log "Disconnect warning: $($_.Exception.Message)" -Level WARNING
    }
    $uBtnConnect.Text      = "Connect to EXO"
    $uBtnConnect.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#0078D4")
    $uBtnConnect.Enabled   = $true
    $uBtnRefresh.Enabled   = $false
    $uBtnDisconnect.Enabled = $false
    $script:AllUsersUnfiltered = @()
    $script:AllUsers           = @()
    $uGrid.Rows.Clear()
    $uPageInfo.Text    = "No users loaded"
    $uStatusLabel.Text = "Disconnected"
    [System.Windows.Forms.MessageBox]::Show("Disconnected from Exchange Online.", "Disconnected",
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
        [System.Windows.Forms.MessageBox]::Show("Please select at least one user.", "No Selection",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $cnt  = $uGrid.SelectedRows.Count
    $names = ($uGrid.SelectedRows | ForEach-Object { $_.Cells["DisplayName"].Value }) -join ", "
    $msg  = if ($cnt -eq 1) { "Convert '$names' to Cloud Managed?" } else { "Convert $cnt users to Cloud Managed?`n`n$names" }
    if ([System.Windows.Forms.MessageBox]::Show($msg, "Confirm", [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question) -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $ok = 0; $fail = 0
    foreach ($row in $uGrid.SelectedRows) {
        $u = @{ DisplayName = $row.Cells["DisplayName"].Value; UserPrincipalName = $row.Cells["UserPrincipalName"].Value }
        if (Convert-UserToCloud -User $u) {
            $row.Cells["IsExchangeCloudManaged"].Value = "True"
            $match = $script:AllUsersUnfiltered | Where-Object { $_.UserPrincipalName -eq $u.UserPrincipalName }
            if ($match) { $match.IsExchangeCloudManaged = $true }
            $ok++
        } else { $fail++ }
    }
    [System.Windows.Forms.MessageBox]::Show("Done.`n`nConverted: $ok`nFailed: $fail", "Batch Complete",
        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    $uStatusLabel.Text = "Converted: $ok  Failed: $fail"
    Update-UserGrid
})

$uBtnToOnPrem.Add_Click({
    if ($uGrid.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one user.", "No Selection",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $cnt  = $uGrid.SelectedRows.Count
    $names = ($uGrid.SelectedRows | ForEach-Object { $_.Cells["DisplayName"].Value }) -join ", "
    $msg  = if ($cnt -eq 1) { "Roll back '$names' to On-Prem Managed?" } else { "Roll back $cnt users to On-Prem Managed?`n`n$names" }
    if ([System.Windows.Forms.MessageBox]::Show($msg, "Confirm", [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question) -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $ok = 0; $fail = 0
    foreach ($row in $uGrid.SelectedRows) {
        $u = @{ DisplayName = $row.Cells["DisplayName"].Value; UserPrincipalName = $row.Cells["UserPrincipalName"].Value }
        if (Convert-UserToOnPrem -User $u) {
            $row.Cells["IsExchangeCloudManaged"].Value = "False"
            $match = $script:AllUsersUnfiltered | Where-Object { $_.UserPrincipalName -eq $u.UserPrincipalName }
            if ($match) { $match.IsExchangeCloudManaged = $false }
            $ok++
        } else { $fail++ }
    }
    [System.Windows.Forms.MessageBox]::Show("Done.`n`nRolled back: $ok`nFailed: $fail", "Batch Complete",
        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    $uStatusLabel.Text = "Rolled back: $ok  Failed: $fail"
    Update-UserGrid
})

$uBtnOpenLog.Add_Click({
    if (Test-Path $script:LogFile) { Start-Process notepad.exe -ArgumentList $script:LogFile }
    else { [System.Windows.Forms.MessageBox]::Show("Log file not found yet.", "Log", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) }
})

# ==============================================================================
# GROUPS TAB -- GUI
# ==============================================================================

$gToolbar = New-Object System.Windows.Forms.Panel
$gToolbar.Height    = 52
$gToolbar.Dock      = [System.Windows.Forms.DockStyle]::Top
$gToolbar.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#F3F3F3")

$gBtnConnect    = New-StyledButton -Text "Connect to Graph"  -X 5   -Y 7 -Width 155 -Height 36 -BackHex "#0078D4" -ForeHex "#FFFFFF"
$gBtnRefresh    = New-StyledButton -Text "Refresh"           -X 168 -Y 7 -Width 110 -Height 36 -Enabled $false
$gBtnDisconnect = New-StyledButton -Text "Disconnect"        -X 286 -Y 7 -Width 120 -Height 36 -Enabled $false

$gSearchLabel = New-Object System.Windows.Forms.Label
$gSearchLabel.Location  = New-Object System.Drawing.Point(420, 14)
$gSearchLabel.Size      = New-Object System.Drawing.Size(55, 22)
$gSearchLabel.Text      = "Search:"
$gSearchLabel.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
$gSearchLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")

$gSearchBox = New-Object System.Windows.Forms.TextBox
$gSearchBox.Location    = New-Object System.Drawing.Point(478, 11)
$gSearchBox.Size        = New-Object System.Drawing.Size(240, 26)
$gSearchBox.Font        = New-Object System.Drawing.Font("Segoe UI", 9)
$gSearchBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

$gChkHide = New-Object System.Windows.Forms.CheckBox
$gChkHide.Location  = New-Object System.Drawing.Point(730, 14)
$gChkHide.Size      = New-Object System.Drawing.Size(205, 24)
$gChkHide.Text      = "Hide Converted Groups"
$gChkHide.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
$gChkHide.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")

$gToolbar.Controls.AddRange(@($gBtnConnect, $gBtnRefresh, $gBtnDisconnect, $gSearchLabel, $gSearchBox, $gChkHide))

$gGrid = New-DataGrid
$gGrid.Dock = [System.Windows.Forms.DockStyle]::Fill
Add-DgvTextColumn $gGrid "DisplayName"    "Display Name"   26
Add-DgvTextColumn $gGrid "Email"          "Email Address"  26
Add-DgvTextColumn $gGrid "GroupType"      "Group Type"     20
Add-DgvTextColumn $gGrid "IsCloudManaged" "Cloud Managed"  12
Add-DgvTextColumn $gGrid "NestingDepth"   "Nesting Depth"  10
Add-DgvTextColumn $gGrid "ObjectId"       "Object ID"      0  $false

$gPaginPanel = New-Object System.Windows.Forms.Panel
$gPaginPanel.Height    = 36
$gPaginPanel.Dock      = [System.Windows.Forms.DockStyle]::Bottom
$gPaginPanel.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#F3F3F3")

$gBtnPrev = New-StyledButton -Text "< Previous" -X 0 -Y 1 -Width 100 -Height 28 -Enabled $false
$gBtnPrev.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom

$gPageInfo = New-Object System.Windows.Forms.Label
$gPageInfo.Location  = New-Object System.Drawing.Point(108, 6)
$gPageInfo.Size      = New-Object System.Drawing.Size(700, 22)
$gPageInfo.Text      = "No groups loaded"
$gPageInfo.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
$gPageInfo.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#605E5C")
$gPageInfo.Anchor    = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom

$gBtnNext = New-StyledButton -Text "Next >" -X 1006 -Y 1 -Width 100 -Height 28 -Enabled $false
$gBtnNext.Anchor = [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom

$gPaginPanel.Controls.AddRange(@($gBtnPrev, $gPageInfo, $gBtnNext))
$gPaginPanel.Add_SizeChanged({
    $w = $gPaginPanel.ClientSize.Width
    $gBtnNext.Location  = New-Object System.Drawing.Point(($w - $gBtnNext.Width), 1)
    $gPageInfo.Location = New-Object System.Drawing.Point(108, 6)
    $gPageInfo.Size     = New-Object System.Drawing.Size([Math]::Max(50, $w - 108 - $gBtnNext.Width - 10), 22)
})

$gActionPanel = New-Object System.Windows.Forms.Panel
$gActionPanel.Height    = 52
$gActionPanel.Dock      = [System.Windows.Forms.DockStyle]::Bottom
$gActionPanel.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#F3F3F3")

$gBtnToCloud  = New-StyledButton -Text "Convert to Cloud Managed"   -X 0   -Y 4 -Width 220 -Height 40
$gBtnToOnPrem = New-StyledButton -Text "Roll Back to On-Prem"       -X 228 -Y 4 -Width 200 -Height 40
$gBtnOpenLog2 = New-StyledButton -Text "Open Log"                   -X 436 -Y 4 -Width 120 -Height 40

$gStatusLabel = New-Object System.Windows.Forms.Label
$gStatusLabel.Location  = New-Object System.Drawing.Point(570, 12)
$gStatusLabel.Size      = New-Object System.Drawing.Size(530, 22)
$gStatusLabel.Text      = "Not connected"
$gStatusLabel.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
$gStatusLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#605E5C")
$gStatusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$gStatusLabel.Anchor    = [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom

$gVersionLabel = New-Object System.Windows.Forms.Label
$gVersionLabel.Location  = New-Object System.Drawing.Point(570, 34)
$gVersionLabel.Size      = New-Object System.Drawing.Size(530, 18)
$gVersionLabel.Text      = "Version $script:Version"
$gVersionLabel.Font      = New-Object System.Drawing.Font("Segoe UI", 8)
$gVersionLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#808080")
$gVersionLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$gVersionLabel.Anchor    = [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom

$gActionPanel.Controls.AddRange(@($gBtnToCloud, $gBtnToOnPrem, $gBtnOpenLog2, $gStatusLabel, $gVersionLabel))
$gActionPanel.Add_SizeChanged({
    $w = $gActionPanel.ClientSize.Width
    $gStatusLabel.Location  = New-Object System.Drawing.Point([Math]::Max(0, $w - $gStatusLabel.Width), 6)
    $gVersionLabel.Location = New-Object System.Drawing.Point([Math]::Max(0, $w - $gVersionLabel.Width), 30)
})

# Dock order: Fill grid first, then bottom panels (action outermost, pagination above), then top toolbar.
$tabGroups.Controls.Add($gGrid)
$tabGroups.Controls.Add($gActionPanel)
$tabGroups.Controls.Add($gPaginPanel)
$tabGroups.Controls.Add($gToolbar)

# ---- Groups grid update helper ----
function Update-GroupGrid {
    $display = Get-FilteredData `
        -Source                $script:AllGroupsUnfiltered `
        -SearchText            $gSearchBox.Text `
        -HideConverted         $script:GroupsHideConverted `
        -ConvertedPropertyName "IsCloudManaged" `
        -EmailPropertyName     "Mail"

    if ($script:GroupsSortColumn -ne "") {
        $sortProp = $script:GroupsSortColumn
        $display = if ($script:GroupsSortDirection -eq "Ascending") {
            if ($sortProp -eq "NestingDepth") {
                $display | Sort-Object { if ($script:NestingDepth.ContainsKey($_.Id)) { $script:NestingDepth[$_.Id] } else { 0 } }
            } else { $display | Sort-Object -Property $sortProp }
        } else {
            if ($sortProp -eq "NestingDepth") {
                $display | Sort-Object { if ($script:NestingDepth.ContainsKey($_.Id)) { $script:NestingDepth[$_.Id] } else { 0 } } -Descending
            } else { $display | Sort-Object -Property $sortProp -Descending }
        }
    }

    $script:AllGroups = @($display)
    $total      = $script:AllGroups.Count
    $totalPages = [Math]::Ceiling($total / $script:PageSize)
    if ($totalPages -lt 1) { $totalPages = 1 }
    if ($script:GroupsCurrentPage -gt $totalPages) { $script:GroupsCurrentPage = $totalPages }
    if ($script:GroupsCurrentPage -lt 1)           { $script:GroupsCurrentPage = 1 }

    $startIdx = ($script:GroupsCurrentPage - 1) * $script:PageSize
    $endIdx   = [Math]::Min($startIdx + $script:PageSize - 1, $total - 1)

    $gGrid.Rows.Clear()
    if ($total -gt 0) {
        for ($i = $startIdx; $i -le $endIdx; $i++) {
            $g     = $script:AllGroups[$i]
            $depth = if ($script:NestingDepth.ContainsKey($g.Id)) { $script:NestingDepth[$g.Id] } else { 0 }
            $val   = if ($g.IsCloudManaged -is [string]) { $g.IsCloudManaged } elseif ($g.IsCloudManaged) { "True" } else { "False" }
            $gGrid.Rows.Add($g.DisplayName, $g.Mail, $g.GroupType, $val, $depth, $g.Id)
        }
        $gPageInfo.Text = "Page $($script:GroupsCurrentPage) of $totalPages  -  Showing $($endIdx - $startIdx + 1) of $total groups"
    } else {
        $gPageInfo.Text = "No groups to display"
    }
    $gBtnPrev.Enabled = ($script:GroupsCurrentPage -gt 1)
    $gBtnNext.Enabled = ($script:GroupsCurrentPage -lt $totalPages)
}

# ---- Groups event wiring ----
$gBtnConnect.Add_Click({
    $gBtnConnect.Enabled = $false
    if (Connect-GraphSession) {
        $ctx = Get-MgContext
        $form.Text = "Exchange SOA Conversion Tool  -  Tenant: $($ctx.TenantId)"
        $gBtnConnect.Text      = "Connected"
        $gBtnConnect.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#107C10")
        $gBtnRefresh.Enabled    = $true
        $gBtnDisconnect.Enabled = $true
        $cBtnConnect.Text      = "Connected (Graph)"
        $cBtnConnect.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#107C10")
        $cBtnConnect.Enabled   = $false
        $cBtnRefresh.Enabled    = $true
        $cBtnDisconnect.Enabled = $true
        $gStatusLabel.Text = "Loading groups..."
        $groups = Get-ExchangeGroups -StatusLabelRef $gStatusLabel
        if ($groups) {
            $script:AllGroupsUnfiltered = $groups
            $gStatusLabel.Text = "Building nesting map..."
            Build-NestingMap -Groups $groups -StatusLabelRef $gStatusLabel
            $script:GroupsCurrentPage = 1
            Update-GroupGrid
            $gStatusLabel.Text = "Connected  -  $($groups.Count) groups loaded"
        } else { $gStatusLabel.Text = "Connected but failed to load groups" }
    } else {
        $gBtnConnect.Enabled = $true
    }
})

$gBtnRefresh.Add_Click({
    $gBtnRefresh.Enabled = $false
    $gStatusLabel.Text = "Refreshing..."
    $groups = Get-ExchangeGroups -StatusLabelRef $gStatusLabel
    if ($groups) {
        $script:AllGroupsUnfiltered = $groups
        $gStatusLabel.Text = "Building nesting map..."
        Build-NestingMap -Groups $groups -StatusLabelRef $gStatusLabel
        $script:GroupsCurrentPage = 1
        Update-GroupGrid
        $gStatusLabel.Text = "Refreshed  -  $($groups.Count) groups loaded"
    }
    $gBtnRefresh.Enabled = $true
})

$gBtnDisconnect.Add_Click({
    try { Disconnect-MgGraph -ErrorAction Stop; Write-Log "Disconnected from Graph." }
    catch { Write-Log "Graph disconnect warning: $($_.Exception.Message)" -Level WARNING }
    $form.Text = "Exchange SOA Conversion Tool"
    $gBtnConnect.Text      = "Connect to Graph"
    $gBtnConnect.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#0078D4")
    $gBtnConnect.Enabled   = $true
    $gBtnRefresh.Enabled   = $false
    $gBtnDisconnect.Enabled = $false
    $cBtnConnect.Text      = "Connect to Graph"
    $cBtnConnect.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#0078D4")
    $cBtnConnect.Enabled   = $true
    $cBtnRefresh.Enabled   = $false
    $cBtnDisconnect.Enabled = $false
    $script:AllGroupsUnfiltered   = @()
    $script:AllGroups             = @()
    $script:AllContactsUnfiltered = @()
    $script:AllContacts           = @()
    $script:PermissionOk          = $false
    $gGrid.Rows.Clear(); $cGrid.Rows.Clear()
    $gPageInfo.Text    = "No groups loaded"
    $cPageInfo.Text    = "No contacts loaded"
    $gStatusLabel.Text = "Disconnected"
    $cStatusLabel.Text = "Disconnected"
    [System.Windows.Forms.MessageBox]::Show("Disconnected from Microsoft Graph.", "Disconnected",
        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

$gSearchBox.Add_TextChanged({
    $script:GroupsCurrentPage = 1
    Update-GroupGrid
})

$gChkHide.Add_CheckedChanged({
    $script:GroupsHideConverted = $gChkHide.Checked
    $script:GroupsCurrentPage   = 1
    Update-GroupGrid
})

$gBtnPrev.Add_Click({
    if ($script:GroupsCurrentPage -gt 1) { $script:GroupsCurrentPage--; Update-GroupGrid }
})
$gBtnNext.Add_Click({
    $tp = [Math]::Ceiling($script:AllGroups.Count / $script:PageSize)
    if ($script:GroupsCurrentPage -lt $tp) { $script:GroupsCurrentPage++; Update-GroupGrid }
})

$gGrid.Add_ColumnHeaderMouseClick({
    param($dgv, $e)
    $col = $dgv.Columns[$e.ColumnIndex].Name
    if ($col -eq "ObjectId") { return }
    if ($script:GroupsSortColumn -eq $col) {
        $script:GroupsSortDirection = if ($script:GroupsSortDirection -eq "Ascending") { "Descending" } else { "Ascending" }
    } else {
        $script:GroupsSortColumn    = $col
        $script:GroupsSortDirection = "Ascending"
    }
    $script:GroupsCurrentPage = 1
    Update-GroupGrid
})

$gBtnToCloud.Add_Click({
    if (-not $script:PermissionOk) {
        [System.Windows.Forms.MessageBox]::Show(
            "Required permission 'Group-OnPremisesSyncBehavior.ReadWrite.All' is not consented.`nReconnect to Graph to grant it.",
            "Permission Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    if ($gGrid.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one group.", "No Selection",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $selectedGroups = @()
    foreach ($row in $gGrid.SelectedRows) {
        $gid = $row.Cells["ObjectId"].Value
        $g   = $script:AllGroups | Where-Object { $_.Id -eq $gid }
        if ($g) { $selectedGroups += $g }
    }

    $warnings = @()
    foreach ($g in $selectedGroups) {
        $uc = Get-UnconvertedChildren -GroupId $g.Id | Where-Object { $_.Id -notin ($selectedGroups | ForEach-Object { $_.Id }) }
        if ($uc.Count -gt 0) {
            $warnings += "Group '$($g.DisplayName)' has unconverted nested children not in selection: $(($uc | ForEach-Object { $_.DisplayName }) -join ', ')"
        }
    }
    if ($warnings.Count -gt 0) {
        $wText = "WARNING: Nested group ordering issue!`n`n$($warnings -join "`n`n")`n`nMicrosoft recommends converting children before parents.`n`nContinue anyway?"
        if ([System.Windows.Forms.MessageBox]::Show($wText, "Nested Group Warning",
            [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning) -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    }

    $sorted = $selectedGroups | Sort-Object { if ($script:NestingDepth.ContainsKey($_.Id)) { $script:NestingDepth[$_.Id] } else { 0 } }
    $cnt    = $sorted.Count
    $list   = ($sorted | ForEach-Object { "$($_.DisplayName) (Depth: $(if ($script:NestingDepth.ContainsKey($_.Id)) { $script:NestingDepth[$_.Id] } else { 0 }))" }) -join "`n"
    $msg    = if ($cnt -eq 1) { "Convert group '$($sorted[0].DisplayName)' to Cloud Managed?" } else { "Convert $cnt groups to Cloud Managed? Order (bottom-up):`n$list" }

    if ([System.Windows.Forms.MessageBox]::Show($msg, "Confirm", [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question) -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $ok = 0; $fail = 0
    foreach ($g in $sorted) {
        $gStatusLabel.Text = "Converting '$($g.DisplayName)'..."
        [System.Windows.Forms.Application]::DoEvents()
        if (Convert-GroupToCloud -Group $g) {
            $g.IsCloudManaged = $true
            foreach ($row in $gGrid.Rows) {
                if ($row.Cells["ObjectId"].Value -eq $g.Id) { $row.Cells["IsCloudManaged"].Value = "True"; break }
            }
            $ok++
        } else { $fail++ }
    }
    [System.Windows.Forms.MessageBox]::Show("Done.`n`nConverted: $ok`nFailed: $fail", "Batch Complete",
        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    $gStatusLabel.Text = "Converted: $ok  Failed: $fail"
    Update-GroupGrid
})

$gBtnToOnPrem.Add_Click({
    if (-not $script:PermissionOk) {
        [System.Windows.Forms.MessageBox]::Show(
            "Required permission 'Group-OnPremisesSyncBehavior.ReadWrite.All' is not consented.`nReconnect to Graph to grant it.",
            "Permission Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    if ($gGrid.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one group.", "No Selection",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $selectedGroups = @()
    foreach ($row in $gGrid.SelectedRows) {
        $gid = $row.Cells["ObjectId"].Value
        $g   = $script:AllGroups | Where-Object { $_.Id -eq $gid }
        if ($g) { $selectedGroups += $g }
    }

    $sorted = $selectedGroups | Sort-Object { if ($script:NestingDepth.ContainsKey($_.Id)) { $script:NestingDepth[$_.Id] } else { 0 } } -Descending
    $cnt    = $sorted.Count
    $list   = ($sorted | ForEach-Object { "$($_.DisplayName) (Depth: $(if ($script:NestingDepth.ContainsKey($_.Id)) { $script:NestingDepth[$_.Id] } else { 0 }))" }) -join "`n"
    $msg    = if ($cnt -eq 1) {
        "Roll back '$($sorted[0].DisplayName)' to On-Prem?`n`nIMPORTANT: Remove cloud users from the group and remove from access packages first."
    } else {
        "Roll back $cnt groups to On-Prem? Order (top-down):`n$list`n`nIMPORTANT: Remove cloud users and remove from access packages first."
    }

    if ([System.Windows.Forms.MessageBox]::Show($msg, "Confirm Rollback", [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question) -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $ok = 0; $fail = 0
    foreach ($g in $sorted) {
        $gStatusLabel.Text = "Rolling back '$($g.DisplayName)'..."
        [System.Windows.Forms.Application]::DoEvents()
        if (Convert-GroupToOnPrem -Group $g) {
            $g.IsCloudManaged = $false
            foreach ($row in $gGrid.Rows) {
                if ($row.Cells["ObjectId"].Value -eq $g.Id) { $row.Cells["IsCloudManaged"].Value = "False"; break }
            }
            $ok++
        } else { $fail++ }
    }
    [System.Windows.Forms.MessageBox]::Show("Done.`n`nRolled back: $ok`nFailed: $fail`n`nNote: Rollback is complete after the next Connect Sync run.", "Batch Complete",
        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    $gStatusLabel.Text = "Rolled back: $ok  Failed: $fail"
    Update-GroupGrid
})

$gBtnOpenLog2.Add_Click({
    if (Test-Path $script:LogFile) { Start-Process notepad.exe -ArgumentList $script:LogFile }
    else { [System.Windows.Forms.MessageBox]::Show("Log file not found yet.", "Log", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) }
})

# ==============================================================================
# CONTACTS TAB -- GUI
# ==============================================================================

$cToolbar = New-Object System.Windows.Forms.Panel
$cToolbar.Height    = 52
$cToolbar.Dock      = [System.Windows.Forms.DockStyle]::Top
$cToolbar.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#F3F3F3")

$cBtnConnect    = New-StyledButton -Text "Connect to Graph"  -X 5   -Y 7 -Width 155 -Height 36 -BackHex "#0078D4" -ForeHex "#FFFFFF"
$cBtnRefresh    = New-StyledButton -Text "Refresh"           -X 168 -Y 7 -Width 110 -Height 36 -Enabled $false
$cBtnDisconnect = New-StyledButton -Text "Disconnect"        -X 286 -Y 7 -Width 120 -Height 36 -Enabled $false

$cSearchLabel = New-Object System.Windows.Forms.Label
$cSearchLabel.Location  = New-Object System.Drawing.Point(420, 14)
$cSearchLabel.Size      = New-Object System.Drawing.Size(55, 22)
$cSearchLabel.Text      = "Search:"
$cSearchLabel.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
$cSearchLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")

$cSearchBox = New-Object System.Windows.Forms.TextBox
$cSearchBox.Location    = New-Object System.Drawing.Point(478, 11)
$cSearchBox.Size        = New-Object System.Drawing.Size(240, 26)
$cSearchBox.Font        = New-Object System.Drawing.Font("Segoe UI", 9)
$cSearchBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

$cChkHide = New-Object System.Windows.Forms.CheckBox
$cChkHide.Location  = New-Object System.Drawing.Point(730, 14)
$cChkHide.Size      = New-Object System.Drawing.Size(220, 24)
$cChkHide.Text      = "Hide Converted Contacts"
$cChkHide.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
$cChkHide.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")

$cToolbar.Controls.AddRange(@($cBtnConnect, $cBtnRefresh, $cBtnDisconnect, $cSearchLabel, $cSearchBox, $cChkHide))

$cGrid = New-DataGrid
$cGrid.Dock = [System.Windows.Forms.DockStyle]::Fill
Add-DgvTextColumn $cGrid "DisplayName"          "Display Name"    28
Add-DgvTextColumn $cGrid "Email"                "Email Address"   30
Add-DgvTextColumn $cGrid "SyncEnabled"          "On-Prem Synced"  18
Add-DgvTextColumn $cGrid "IsCloudManaged"       "Cloud Managed"   14
Add-DgvTextColumn $cGrid "ObjectId"             "Object ID"       0  $false

$cPaginPanel = New-Object System.Windows.Forms.Panel
$cPaginPanel.Height    = 36
$cPaginPanel.Dock      = [System.Windows.Forms.DockStyle]::Bottom
$cPaginPanel.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#F3F3F3")

$cBtnPrev = New-StyledButton -Text "< Previous" -X 0 -Y 1 -Width 100 -Height 28 -Enabled $false
$cBtnPrev.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom

$cPageInfo = New-Object System.Windows.Forms.Label
$cPageInfo.Location  = New-Object System.Drawing.Point(108, 6)
$cPageInfo.Size      = New-Object System.Drawing.Size(700, 22)
$cPageInfo.Text      = "No contacts loaded"
$cPageInfo.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
$cPageInfo.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#605E5C")
$cPageInfo.Anchor    = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom

$cBtnNext = New-StyledButton -Text "Next >" -X 1006 -Y 1 -Width 100 -Height 28 -Enabled $false
$cBtnNext.Anchor = [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom

$cPaginPanel.Controls.AddRange(@($cBtnPrev, $cPageInfo, $cBtnNext))
$cPaginPanel.Add_SizeChanged({
    $w = $cPaginPanel.ClientSize.Width
    $cBtnNext.Location  = New-Object System.Drawing.Point(($w - $cBtnNext.Width), 1)
    $cPageInfo.Location = New-Object System.Drawing.Point(108, 6)
    $cPageInfo.Size     = New-Object System.Drawing.Size([Math]::Max(50, $w - 108 - $cBtnNext.Width - 10), 22)
})

$cActionPanel = New-Object System.Windows.Forms.Panel
$cActionPanel.Height    = 52
$cActionPanel.Dock      = [System.Windows.Forms.DockStyle]::Bottom
$cActionPanel.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#F3F3F3")

$cBtnToCloud  = New-StyledButton -Text "Convert to Cloud Managed"   -X 0   -Y 4 -Width 220 -Height 40
$cBtnToOnPrem = New-StyledButton -Text "Roll Back to On-Prem"       -X 228 -Y 4 -Width 200 -Height 40
$cBtnOpenLog3 = New-StyledButton -Text "Open Log"                   -X 436 -Y 4 -Width 120 -Height 40

$cStatusLabel = New-Object System.Windows.Forms.Label
$cStatusLabel.Location  = New-Object System.Drawing.Point(570, 12)
$cStatusLabel.Size      = New-Object System.Drawing.Size(530, 22)
$cStatusLabel.Text      = "Not connected"
$cStatusLabel.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
$cStatusLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#605E5C")
$cStatusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$cStatusLabel.Anchor    = [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom

$cVersionLabel = New-Object System.Windows.Forms.Label
$cVersionLabel.Location  = New-Object System.Drawing.Point(570, 34)
$cVersionLabel.Size      = New-Object System.Drawing.Size(530, 18)
$cVersionLabel.Text      = "Version $script:Version"
$cVersionLabel.Font      = New-Object System.Drawing.Font("Segoe UI", 8)
$cVersionLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#808080")
$cVersionLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$cVersionLabel.Anchor    = [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom

$cActionPanel.Controls.AddRange(@($cBtnToCloud, $cBtnToOnPrem, $cBtnOpenLog3, $cStatusLabel, $cVersionLabel))
$cActionPanel.Add_SizeChanged({
    $w = $cActionPanel.ClientSize.Width
    $cStatusLabel.Location  = New-Object System.Drawing.Point([Math]::Max(0, $w - $cStatusLabel.Width), 6)
    $cVersionLabel.Location = New-Object System.Drawing.Point([Math]::Max(0, $w - $cVersionLabel.Width), 30)
})

# Dock order: Fill grid first, then bottom panels (action outermost, pagination above), then top toolbar.
$tabContacts.Controls.Add($cGrid)
$tabContacts.Controls.Add($cActionPanel)
$tabContacts.Controls.Add($cPaginPanel)
$tabContacts.Controls.Add($cToolbar)

# ---- Contacts grid update helper ----
function Update-ContactGrid {
    $display = Get-FilteredData `
        -Source                $script:AllContactsUnfiltered `
        -SearchText            $cSearchBox.Text `
        -HideConverted         $script:ContactsHideConverted `
        -ConvertedPropertyName "IsCloudManaged" `
        -EmailPropertyName     "Mail"

    if ($script:ContactsSortColumn -ne "") {
        $display = if ($script:ContactsSortDirection -eq "Ascending") {
            $display | Sort-Object -Property $script:ContactsSortColumn
        } else {
            $display | Sort-Object -Property $script:ContactsSortColumn -Descending
        }
    }

    $script:AllContacts = @($display)
    $total      = $script:AllContacts.Count
    $totalPages = [Math]::Ceiling($total / $script:PageSize)
    if ($totalPages -lt 1) { $totalPages = 1 }
    if ($script:ContactsCurrentPage -gt $totalPages) { $script:ContactsCurrentPage = $totalPages }
    if ($script:ContactsCurrentPage -lt 1)           { $script:ContactsCurrentPage = 1 }

    $startIdx = ($script:ContactsCurrentPage - 1) * $script:PageSize
    $endIdx   = [Math]::Min($startIdx + $script:PageSize - 1, $total - 1)

    $cGrid.Rows.Clear()
    if ($total -gt 0) {
        for ($i = $startIdx; $i -le $endIdx; $i++) {
            $c       = $script:AllContacts[$i]
            $synced  = if ($c.OnPremisesSyncEnabled) { "True" } else { "False" }
            $cmVal   = if ($c.IsCloudManaged -is [string]) { $c.IsCloudManaged } elseif ($c.IsCloudManaged) { "True" } else { "False" }
            $rowIdx  = $cGrid.Rows.Add($c.DisplayName, $c.Mail, $synced, $cmVal, $c.Id)

            if ($c.IsCloudManaged -eq "N/A") {
                $cGrid.Rows[$rowIdx].DefaultCellStyle.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#A0A0A0")
            }
        }
        $cPageInfo.Text = "Page $($script:ContactsCurrentPage) of $totalPages  -  Showing $($endIdx - $startIdx + 1) of $total contacts"
    } else {
        $cPageInfo.Text = "No contacts to display"
    }
    $cBtnPrev.Enabled = ($script:ContactsCurrentPage -gt 1)
    $cBtnNext.Enabled = ($script:ContactsCurrentPage -lt $totalPages)
}

# ---- Contacts event wiring ----
$cBtnConnect.Add_Click({
    $cBtnConnect.Enabled = $false
    if (Connect-GraphSession) {
        $ctx = Get-MgContext
        $form.Text = "Exchange SOA Conversion Tool  -  Tenant: $($ctx.TenantId)"
        $cBtnConnect.Text      = "Connected (Graph)"
        $cBtnConnect.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#107C10")
        $cBtnRefresh.Enabled    = $true
        $cBtnDisconnect.Enabled = $true
        $gBtnConnect.Text      = "Connected"
        $gBtnConnect.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#107C10")
        $gBtnConnect.Enabled   = $false
        $gBtnRefresh.Enabled    = $true
        $gBtnDisconnect.Enabled = $true
        $cStatusLabel.Text = "Loading contacts..."
        $contacts = Get-OrgContacts -StatusLabelRef $cStatusLabel
        if ($contacts) {
            $script:AllContactsUnfiltered = $contacts
            $script:ContactsCurrentPage   = 1
            Update-ContactGrid
            $cStatusLabel.Text = "Connected  -  $($contacts.Count) contacts loaded"
        } else { $cStatusLabel.Text = "Connected but failed to load contacts" }
    } else {
        $cBtnConnect.Enabled = $true
    }
})

$cBtnRefresh.Add_Click({
    $cBtnRefresh.Enabled = $false
    $cStatusLabel.Text = "Refreshing..."
    $contacts = Get-OrgContacts -StatusLabelRef $cStatusLabel
    if ($contacts) {
        $script:AllContactsUnfiltered = $contacts
        $script:ContactsCurrentPage   = 1
        Update-ContactGrid
        $cStatusLabel.Text = "Refreshed  -  $($contacts.Count) contacts loaded"
    }
    $cBtnRefresh.Enabled = $true
})

$cBtnDisconnect.Add_Click({
    try { Disconnect-MgGraph -ErrorAction Stop; Write-Log "Disconnected from Graph." }
    catch { Write-Log "Graph disconnect warning: $($_.Exception.Message)" -Level WARNING }
    $form.Text = "Exchange SOA Conversion Tool"
    $cBtnConnect.Text      = "Connect to Graph"
    $cBtnConnect.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#0078D4")
    $cBtnConnect.Enabled   = $true
    $cBtnRefresh.Enabled   = $false
    $cBtnDisconnect.Enabled = $false
    $gBtnConnect.Text      = "Connect to Graph"
    $gBtnConnect.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#0078D4")
    $gBtnConnect.Enabled   = $true
    $gBtnRefresh.Enabled   = $false
    $gBtnDisconnect.Enabled = $false
    $script:AllGroupsUnfiltered   = @()
    $script:AllGroups             = @()
    $script:AllContactsUnfiltered = @()
    $script:AllContacts           = @()
    $script:PermissionOk          = $false
    $gGrid.Rows.Clear(); $cGrid.Rows.Clear()
    $gPageInfo.Text    = "No groups loaded"
    $cPageInfo.Text    = "No contacts loaded"
    $gStatusLabel.Text = "Disconnected"
    $cStatusLabel.Text = "Disconnected"
    [System.Windows.Forms.MessageBox]::Show("Disconnected from Microsoft Graph.", "Disconnected",
        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

$cSearchBox.Add_TextChanged({
    $script:ContactsCurrentPage = 1
    Update-ContactGrid
})

$cChkHide.Add_CheckedChanged({
    $script:ContactsHideConverted = $cChkHide.Checked
    $script:ContactsCurrentPage   = 1
    Update-ContactGrid
})

$cBtnPrev.Add_Click({
    if ($script:ContactsCurrentPage -gt 1) { $script:ContactsCurrentPage--; Update-ContactGrid }
})
$cBtnNext.Add_Click({
    $tp = [Math]::Ceiling($script:AllContacts.Count / $script:PageSize)
    if ($script:ContactsCurrentPage -lt $tp) { $script:ContactsCurrentPage++; Update-ContactGrid }
})

$cGrid.Add_ColumnHeaderMouseClick({
    param($dgv, $e)
    $col = $dgv.Columns[$e.ColumnIndex].Name
    if ($col -eq "ObjectId") { return }
    if ($script:ContactsSortColumn -eq $col) {
        $script:ContactsSortDirection = if ($script:ContactsSortDirection -eq "Ascending") { "Descending" } else { "Ascending" }
    } else {
        $script:ContactsSortColumn    = $col
        $script:ContactsSortDirection = "Ascending"
    }
    $script:ContactsCurrentPage = 1
    Update-ContactGrid
})

$cBtnToCloud.Add_Click({
    if ($cGrid.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one contact.", "No Selection",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $selectedContacts = @()
    foreach ($row in $cGrid.SelectedRows) {
        $cid = $row.Cells["ObjectId"].Value
        $c   = $script:AllContacts | Where-Object { $_.Id -eq $cid }
        if ($c) {
            if ($c.OnPremisesSyncEnabled -ne $true) {
                [System.Windows.Forms.MessageBox]::Show(
                    "'$($c.DisplayName)' is not an on-premises synced contact (N/A).`nSOA conversion only applies to synced contacts. Skipping.",
                    "Not Applicable", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                continue
            }
            $selectedContacts += $c
        }
    }
    if ($selectedContacts.Count -eq 0) { return }

    $cnt   = $selectedContacts.Count
    $names = ($selectedContacts | ForEach-Object { $_.DisplayName }) -join ", "
    $msg   = if ($cnt -eq 1) { "Convert contact '$names' to Cloud Managed?" } else { "Convert $cnt contacts to Cloud Managed?`n`n$names" }
    if ([System.Windows.Forms.MessageBox]::Show($msg, "Confirm", [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question) -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $ok = 0; $fail = 0
    foreach ($c in $selectedContacts) {
        $cStatusLabel.Text = "Converting '$($c.DisplayName)'..."
        [System.Windows.Forms.Application]::DoEvents()
        if (Convert-ContactToCloud -Contact $c) {
            $c.IsCloudManaged = $true
            foreach ($row in $cGrid.Rows) {
                if ($row.Cells["ObjectId"].Value -eq $c.Id) { $row.Cells["IsCloudManaged"].Value = "True"; break }
            }
            $ok++
        } else { $fail++ }
    }
    [System.Windows.Forms.MessageBox]::Show("Done.`n`nConverted: $ok`nFailed: $fail", "Batch Complete",
        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    $cStatusLabel.Text = "Converted: $ok  Failed: $fail"
    Update-ContactGrid
})

$cBtnToOnPrem.Add_Click({
    if ($cGrid.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one contact.", "No Selection",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $selectedContacts = @()
    foreach ($row in $cGrid.SelectedRows) {
        $cid = $row.Cells["ObjectId"].Value
        $c   = $script:AllContacts | Where-Object { $_.Id -eq $cid }
        if ($c) {
            if ($c.OnPremisesSyncEnabled -ne $true) {
                [System.Windows.Forms.MessageBox]::Show(
                    "'$($c.DisplayName)' is not an on-premises synced contact (N/A).`nSOA conversion only applies to synced contacts. Skipping.",
                    "Not Applicable", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                continue
            }
            $selectedContacts += $c
        }
    }
    if ($selectedContacts.Count -eq 0) { return }

    $cnt   = $selectedContacts.Count
    $names = ($selectedContacts | ForEach-Object { $_.DisplayName }) -join ", "
    $msg   = if ($cnt -eq 1) { "Roll back contact '$names' to On-Prem Managed?" } else { "Roll back $cnt contacts to On-Prem Managed?`n`n$names" }
    if ([System.Windows.Forms.MessageBox]::Show($msg, "Confirm Rollback", [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question) -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $ok = 0; $fail = 0
    foreach ($c in $selectedContacts) {
        $cStatusLabel.Text = "Rolling back '$($c.DisplayName)'..."
        [System.Windows.Forms.Application]::DoEvents()
        if (Convert-ContactToOnPrem -Contact $c) {
            $c.IsCloudManaged = $false
            foreach ($row in $cGrid.Rows) {
                if ($row.Cells["ObjectId"].Value -eq $c.Id) { $row.Cells["IsCloudManaged"].Value = "False"; break }
            }
            $ok++
        } else { $fail++ }
    }
    [System.Windows.Forms.MessageBox]::Show("Done.`n`nRolled back: $ok`nFailed: $fail", "Batch Complete",
        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    $cStatusLabel.Text = "Rolled back: $ok  Failed: $fail"
    Update-ContactGrid
})

$cBtnOpenLog3.Add_Click({
    if (Test-Path $script:LogFile) { Start-Process notepad.exe -ArgumentList $script:LogFile }
    else { [System.Windows.Forms.MessageBox]::Show("Log file not found yet.", "Log", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) }
})

# ==============================================================================
# FORM CLOSING HANDLER
# ==============================================================================

$form.Add_FormClosing({
    Write-Log "Exchange SOA Conversion Tool closing..."
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
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
