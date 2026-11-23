<#
.SYNOPSIS
    Azure DevOps Server REST API Tool (Windows Forms)
    Compatible with Azure DevOps Server 2020 Update 2 through 25H2 (20260506.1) and later.

.DESCRIPTION
    Multi-environment support via JSON config files.
    Automatic or configurable API version selection.
    Automatic backup of current policy/repo state before any modifying operation.

.NOTES
    Place config files in a "configs" folder next to this script.
    Never commit live PATs. Sample configs are provided.
#>

#Requires -Version 5.1

###===================================================================================================
## Initialization
###===================================================================================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

# Cleanup previous form if re-run in same session
try {
    if (Get-Variable -Name Form -Scope Script -ErrorAction SilentlyContinue) {
        if ($null -ne $Form -and $Form -isnot [System.Array]) { $Form.Close(); $Form.Dispose() }
        Remove-Variable -Name Form -Scope Script -ErrorAction SilentlyContinue
    }
} catch {}

$script:ConfigsPath = Join-Path $PSScriptRoot "configs"
$script:BackupsPath = Join-Path $PSScriptRoot "backups"
$script:LogPath     = $PSScriptRoot

New-Item -ItemType Directory -Path $script:ConfigsPath -Force -ErrorAction SilentlyContinue | Out-Null
New-Item -ItemType Directory -Path $script:BackupsPath -Force -ErrorAction SilentlyContinue | Out-Null

$logDate = Get-Date -Format "yyyy-MM-dd"
$script:LogFile = Join-Path $script:LogPath "ADO_REST_Tools_Result_$logDate.log"
if (-not (Test-Path $script:LogFile)) { New-Item -Path $script:LogFile -ItemType File -Force | Out-Null }

###===================================================================================================
## Script-scoped state
###===================================================================================================
$script:CurrentConfig = $null
$script:ADOServerFQDN = ""
$script:Collection    = "DevOpsCollection"
$script:MyPat         = ""
$script:ApiVersion    = "6.0"
$script:UseHttps      = $true
$script:UrlPrefix     = ""
$script:AuthMode      = "PAT"   # PAT or Windows
$script:Headers       = @{}
$script:RepoListCache = @()
$script:SelectedRepoIds = @()
$script:SelectedRepoNames = @()

function Get-BaseUrl {
    $scheme = if ($script:UseHttps) { "https" } else { "http" }
    $prefix = if ($script:UrlPrefix) { "/" + $script:UrlPrefix } else { "" }
    return ($scheme + "://" + $script:ADOServerFQDN + $prefix)
}

function Update-Headers {
    if ([string]::IsNullOrWhiteSpace($script:MyPat)) {
        $script:Headers = @{}
        return
    }
    # Standard ADO PAT Basic auth: empty username + PAT as password
    $b64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes(":$($script:MyPat)"))
    $script:Headers = @{
        Authorization = "Basic $b64"
    }
}

function Write-Log {
    param([string]$Message)
    $ts = Get-Date -Format "yyyy/MM/dd HH:mm:ss"
    $line = "[$ts] $Message"
    Add-Content -Path $script:LogFile -Value $line -ErrorAction SilentlyContinue
    if ($null -ne $script:TextBoxResult) {
        $script:TextBoxResult.AppendText("$line`r`n")
    }
}

###===================================================================================================
## Configuration management
###===================================================================================================
function Get-ConfigFiles {
    if (-not (Test-Path $script:ConfigsPath)) { return @() }
    Get-ChildItem -Path $script:ConfigsPath -Filter "*.json" -ErrorAction SilentlyContinue |
        Sort-Object Name
}

function Read-Config {
    param([string]$Path)
    try {
        $raw = Get-Content -Path $Path -Raw -ErrorAction Stop
        $obj = $raw | ConvertFrom-Json
        $obj | Add-Member -NotePropertyName FilePath -NotePropertyValue $Path -Force
        return $obj
    } catch {
        Write-Log "Failed to read config $Path : $_"
        return $null
    }
}

function Get-ConfigEntries {
    param([object]$Obj)
    if ($null -eq $Obj) { return @() }
    # Support: single object, array of objects, or { Servers: [...] } / { Environments: [...] }
    if ($Obj -is [System.Array]) { return @($Obj) }
    if ($Obj.PSObject.Properties.Name -contains "Servers") { return @($Obj.Servers) }
    if ($Obj.PSObject.Properties.Name -contains "Environments") { return @($Obj.Environments) }
    if ($Obj.PSObject.Properties.Name -contains "ServerFQDN") { return @($Obj) }
    return @($Obj)
}

function Load-Config {
    param([object]$Config)
    if ($null -eq $Config) { return }
    $script:CurrentConfig = $Config
    $script:ADOServerFQDN = if ($Config.ServerFQDN) { $Config.ServerFQDN.Trim() } else { "" }
    $script:Collection    = if ($Config.Collection) { $Config.Collection.Trim() } else { "DevOpsCollection" }
    $rawPat = if ($Config.PAT) { $Config.PAT } elseif ($Config.Pat) { $Config.Pat } elseif ($Config.Token) { $Config.Token } else { "" }
    $script:MyPat = if ($rawPat) { $rawPat.ToString().Trim() } else { "" }
    $script:UseHttps      = if ($null -ne $Config.UseHttps) { [bool]$Config.UseHttps } else { $true }
    $script:UrlPrefix     = if ($Config.UrlPrefix) { $Config.UrlPrefix.Trim().Trim('/') } else { "" }
    $script:AuthMode      = if ($Config.AuthMode) { $Config.AuthMode } else { "PAT" }
    $script:ApiVersion    = if ($Config.ApiVersion -and $Config.ApiVersion -ne "auto") { $Config.ApiVersion } else { "6.0" }
    Update-Headers
    $patStatus = if ([string]::IsNullOrWhiteSpace($script:MyPat)) { "MISSING" } else { "set (len=$($script:MyPat.Length))" }
    $prefixLog = if ($script:UrlPrefix) { " | Prefix: $($script:UrlPrefix)" } else { "" }
    Write-Log "Configuration loaded successfully."
    Write-Log "  Name       : $($Config.Name)"
    Write-Log "  Server     : $($script:ADOServerFQDN)"
    Write-Log "  Collection : $($script:Collection)$prefixLog"
    Write-Log "  API Version: $($script:ApiVersion)"
    Write-Log "  Auth Mode  : $($script:AuthMode)"
    Write-Log "  PAT Status : $patStatus"
}

function Detect-ApiVersion {
    if ([string]::IsNullOrWhiteSpace($script:ADOServerFQDN) -or [string]::IsNullOrWhiteSpace($script:MyPat)) {
        return "6.0"
    }
    $base = Get-BaseUrl
    $testVersions = @("7.2", "7.1", "7.0", "6.0")
    $firstUrl = $null
    foreach ($ver in $testVersions) {
        $url = "$base/$($script:Collection)/_apis/projects?api-version=$ver&`$top=1"
        if (-not $firstUrl) { $firstUrl = $url }
        try {
            $detectParams = @{ Uri = $url; Method = "Get"; ErrorAction = "Stop" }
            if ($script:AuthMode -eq "Windows") { $detectParams.UseDefaultCredentials = $true } else { $detectParams.Headers = $script:Headers }
            $null = Invoke-RestMethod @detectParams
            Write-Log "API version $ver is supported by the server."
            return $ver
        } catch {
            continue
        }
    }
    Write-Log "Could not auto-detect API version. Falling back to 6.0 (compatible with Azure DevOps Server 2020)."
    if ($firstUrl) { Write-Log "Tried URL: $firstUrl" }
    return "6.0"
}

function Ensure-ApiVersion {
    if ($script:CurrentConfig -and $script:CurrentConfig.ApiVersion -eq "auto") {
        $detected = Detect-ApiVersion
        $script:ApiVersion = $detected
    }
}

###===================================================================================================
## Backup before any update
###===================================================================================================
function Backup-CurrentState {
    param(
        [string]$Operation,
        [string]$ProjectName
    )
    try {
        $ts = Get-Date -Format "yyyyMMdd_HHmmss"
        $safeProject = if ($ProjectName) { $ProjectName -replace '[^\w\-]', '_' } else { "unknown" }
        $backupName = "backup_${Operation}_${safeProject}_$ts.json"
        $backupFile = Join-Path $script:BackupsPath $backupName

        $backup = [ordered]@{
            Timestamp   = (Get-Date).ToString("o")
            Operation   = $Operation
            Server      = $script:ADOServerFQDN
            Collection  = $script:Collection
            Project     = $ProjectName
            ApiVersion  = $script:ApiVersion
            ConfigName  = if ($script:CurrentConfig) { $script:CurrentConfig.Name } else { "none" }
            Policies    = @()
            Repositories = @()
        }

        if (-not [string]::IsNullOrWhiteSpace($ProjectName)) {
            $projEnc = [uri]::EscapeDataString($ProjectName)
            $base = Get-BaseUrl
            $policiesUrl = "$base/$($script:Collection)/$projEnc/_apis/policy/configurations?api-version=$($script:ApiVersion)"
            try {
                $polResp = Invoke-RestMethod -Uri $policiesUrl -Headers $script:Headers -Method Get
                $backup.Policies = $polResp.value
            } catch {
                Write-Log "Backup: could not retrieve policies: $_"
            }

            $reposUrl = "$base/$($script:Collection)/$projEnc/_apis/git/repositories?api-version=$($script:ApiVersion)"
            try {
                $repoResp = Invoke-RestMethod -Uri $reposUrl -Headers $script:Headers -Method Get
                $backup.Repositories = $repoResp.value | Select-Object id, name, defaultBranch
            } catch {
                Write-Log "Backup: could not retrieve repositories: $_"
            }
        }

        $backup | ConvertTo-Json -Depth 10 | Set-Content -Path $backupFile -Encoding UTF8
        Write-Log "Backup of current policies and repository settings saved: $backupName"
        return $backupFile
    } catch {
        Write-Log "Backup failed: $_"
        return $null
    }
}

###===================================================================================================
## REST helpers
###===================================================================================================
function Invoke-AdoRest {
    param(
        [string]$Uri,
        [string]$Method = "Get",
        [object]$Body = $null
    )
    $params = @{
        Uri         = $Uri
        Method      = $Method
        ErrorAction = "Stop"
    }
    if ($script:AuthMode -eq "Windows") {
        $params.UseDefaultCredentials = $true
    } else {
        $params.Headers = $script:Headers
    }
    if ($Body) {
        $params.Body = ($Body | ConvertTo-Json -Depth 10)
        $params.ContentType = "application/json"
    }
    try {
        return Invoke-RestMethod @params
    } catch {
        $status = $null
        if ($_.Exception.Response) { $status = [int]$_.Exception.Response.StatusCode }
        Write-Log "REST $Method failed: $Uri"
        if ($status) { Write-Log "HTTP Status: $status" }
        if ($status -eq 401 -and $script:AuthMode -eq "PAT") {
            Write-Log "Hint: On-prem Azure DevOps Server often requires Basic Auth to be enabled in IIS, or use AuthMode: Windows with a domain account."
        }
        throw
    }
}

function Get-Projects {
    Ensure-ApiVersion
    $url = "$(Get-BaseUrl)/$($script:Collection)/_apis/projects?api-version=$($script:ApiVersion)"
    return (Invoke-AdoRest -Uri $url).value
}

function Get-Repositories {
    param([string]$ProjectName)
    Ensure-ApiVersion
    $enc = [uri]::EscapeDataString($ProjectName)
    $url = "$(Get-BaseUrl)/$($script:Collection)/$enc/_apis/git/repositories?api-version=$($script:ApiVersion)"
    return (Invoke-AdoRest -Uri $url).value
}

function Get-Policies {
    param([string]$ProjectName)
    Ensure-ApiVersion
    $enc = [uri]::EscapeDataString($ProjectName)
    $url = "$(Get-BaseUrl)/$($script:Collection)/$enc/_apis/policy/configurations?api-version=$($script:ApiVersion)"
    return (Invoke-AdoRest -Uri $url).value
}

###===================================================================================================
## UI
###===================================================================================================
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Azure DevOps Server REST API Tool (2020 Update 2 - 25H2+)"
$Form.StartPosition = "CenterScreen"
$Form.ClientSize = New-Object System.Drawing.Size(1320, 920)
$Form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$Form.MinimumSize = New-Object System.Drawing.Size(1100, 700)

# Consistent button size
$btnW = 180
$btnH = 32

# --- Group: Environment ---
$GroupEnv = New-Object System.Windows.Forms.GroupBox
$GroupEnv.Text = "Environment / Config"
$GroupEnv.Location = New-Object System.Drawing.Point(10, 8)
$GroupEnv.Size = New-Object System.Drawing.Size(1300, 70)
$GroupEnv.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

$LabelConfig = New-Object System.Windows.Forms.Label
$LabelConfig.Text = "Environment:"
$LabelConfig.Location = New-Object System.Drawing.Point(12, 28)
$LabelConfig.Size = New-Object System.Drawing.Size(80, 22)

$ComboBoxConfig = New-Object System.Windows.Forms.ComboBox
$ComboBoxConfig.Location = New-Object System.Drawing.Point(95, 26)
$ComboBoxConfig.Size = New-Object System.Drawing.Size(260, 24)
$ComboBoxConfig.DropDownStyle = "DropDownList"

$ButtonLoadConfig = New-Object System.Windows.Forms.Button
$ButtonLoadConfig.Text = "Load Config"
$ButtonLoadConfig.Location = New-Object System.Drawing.Point(365, 24)
$ButtonLoadConfig.Size = New-Object System.Drawing.Size($btnW, $btnH)

$ButtonEditConfig = New-Object System.Windows.Forms.Button
$ButtonEditConfig.Text = "Edit / New Config"
$ButtonEditConfig.Location = New-Object System.Drawing.Point(555, 24)
$ButtonEditConfig.Size = New-Object System.Drawing.Size($btnW, $btnH)

$LabelApiVersion = New-Object System.Windows.Forms.Label
$LabelApiVersion.Text = "API Version:"
$LabelApiVersion.Location = New-Object System.Drawing.Point(750, 28)
$LabelApiVersion.Size = New-Object System.Drawing.Size(80, 22)

$ComboBoxApiVersion = New-Object System.Windows.Forms.ComboBox
$ComboBoxApiVersion.Location = New-Object System.Drawing.Point(835, 26)
$ComboBoxApiVersion.Size = New-Object System.Drawing.Size(100, 24)
$ComboBoxApiVersion.DropDownStyle = "DropDownList"
@("auto","6.0","7.0","7.1","7.2") | ForEach-Object { [void]$ComboBoxApiVersion.Items.Add($_) }
$ComboBoxApiVersion.SelectedItem = "auto"

$ButtonExit = New-Object System.Windows.Forms.Button
$ButtonExit.Text = "Exit"
$ButtonExit.Location = New-Object System.Drawing.Point(1100, 24)
$ButtonExit.Size = New-Object System.Drawing.Size(120, $btnH)

$GroupEnv.Controls.AddRange(@($LabelConfig, $ComboBoxConfig, $ButtonLoadConfig, $ButtonEditConfig, $LabelApiVersion, $ComboBoxApiVersion, $ButtonExit))

# --- Group: Project ---
$GroupProject = New-Object System.Windows.Forms.GroupBox
$GroupProject.Text = "Project / Repositories"
$GroupProject.Location = New-Object System.Drawing.Point(10, 84)
$GroupProject.Size = New-Object System.Drawing.Size(1300, 70)
$GroupProject.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

$LabelProject = New-Object System.Windows.Forms.Label
$LabelProject.Text = "Project:"
$LabelProject.Location = New-Object System.Drawing.Point(12, 28)
$LabelProject.Size = New-Object System.Drawing.Size(55, 22)

$ComboBoxProject = New-Object System.Windows.Forms.ComboBox
$ComboBoxProject.Location = New-Object System.Drawing.Point(70, 26)
$ComboBoxProject.Size = New-Object System.Drawing.Size(280, 24)
$ComboBoxProject.DropDownStyle = "DropDownList"
$ComboBoxProject.Enabled = $false

$ButtonRefreshProjects = New-Object System.Windows.Forms.Button
$ButtonRefreshProjects.Text = "Refresh Projects"
$ButtonRefreshProjects.Location = New-Object System.Drawing.Point(360, 24)
$ButtonRefreshProjects.Size = New-Object System.Drawing.Size($btnW, $btnH)

$ButtonSelectRepos = New-Object System.Windows.Forms.Button
$ButtonSelectRepos.Text = "Select Repos"
$ButtonSelectRepos.Location = New-Object System.Drawing.Point(550, 24)
$ButtonSelectRepos.Size = New-Object System.Drawing.Size($btnW, $btnH)
$ButtonSelectRepos.Enabled = $false

$LabelSelectedRepos = New-Object System.Windows.Forms.Label
$LabelSelectedRepos.Text = "Repos: All"
$LabelSelectedRepos.Location = New-Object System.Drawing.Point(745, 28)
$LabelSelectedRepos.Size = New-Object System.Drawing.Size(500, 22)

$GroupProject.Controls.AddRange(@($LabelProject, $ComboBoxProject, $ButtonRefreshProjects, $ButtonSelectRepos, $LabelSelectedRepos))

# --- Group: Branch ---
$GroupBranch = New-Object System.Windows.Forms.GroupBox
$GroupBranch.Text = "Default Branch"
$GroupBranch.Location = New-Object System.Drawing.Point(10, 160)
$GroupBranch.Size = New-Object System.Drawing.Size(640, 80)
$GroupBranch.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left

$LabelBranch = New-Object System.Windows.Forms.Label
$LabelBranch.Text = "Branch:"
$LabelBranch.Location = New-Object System.Drawing.Point(12, 30)
$LabelBranch.Size = New-Object System.Drawing.Size(55, 22)

$ObjBoxBranch = New-Object System.Windows.Forms.ComboBox
$ObjBoxBranch.Location = New-Object System.Drawing.Point(70, 28)
$ObjBoxBranch.Size = New-Object System.Drawing.Size(140, 24)
$ObjBoxBranch.DropDownStyle = "DropDownList"
@("Develop","Main") | ForEach-Object { [void]$ObjBoxBranch.Items.Add($_) }
$ObjBoxBranch.SelectedIndex = 0
$ObjBoxBranch.Enabled = $false

$ButtonQueryAllRepo = New-Object System.Windows.Forms.Button
$ButtonQueryAllRepo.Text = "Query Default Branches"
$ButtonQueryAllRepo.Location = New-Object System.Drawing.Point(230, 26)
$ButtonQueryAllRepo.Size = New-Object System.Drawing.Size($btnW, $btnH)

$ButtonSetDefaultBranch = New-Object System.Windows.Forms.Button
$ButtonSetDefaultBranch.Text = "Set Default Branch"
$ButtonSetDefaultBranch.Location = New-Object System.Drawing.Point(420, 26)
$ButtonSetDefaultBranch.Size = New-Object System.Drawing.Size($btnW, $btnH)

$GroupBranch.Controls.AddRange(@($LabelBranch, $ObjBoxBranch, $ButtonQueryAllRepo, $ButtonSetDefaultBranch))

# --- Group: Policy Type filter ---
$GroupFilter = New-Object System.Windows.Forms.GroupBox
$GroupFilter.Text = "Policy Filter"
$GroupFilter.Location = New-Object System.Drawing.Point(660, 160)
$GroupFilter.Size = New-Object System.Drawing.Size(650, 80)
$GroupFilter.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

$LabelPolicyType = New-Object System.Windows.Forms.Label
$LabelPolicyType.Text = "Policy Type:"
$LabelPolicyType.Location = New-Object System.Drawing.Point(12, 30)
$LabelPolicyType.Size = New-Object System.Drawing.Size(80, 22)

$ObjBoxPolicyType = New-Object System.Windows.Forms.ComboBox
$ObjBoxPolicyType.Location = New-Object System.Drawing.Point(95, 28)
$ObjBoxPolicyType.Size = New-Object System.Drawing.Size(280, 24)
$ObjBoxPolicyType.DropDownStyle = "DropDownList"
@("All","Minimum number of reviewers","Required reviewers") | ForEach-Object { [void]$ObjBoxPolicyType.Items.Add($_) }
$ObjBoxPolicyType.SelectedIndex = 0

$GroupFilter.Controls.AddRange(@($LabelPolicyType, $ObjBoxPolicyType))

# --- Group: Project Policies ---
$GroupProjPolicy = New-Object System.Windows.Forms.GroupBox
$GroupProjPolicy.Text = "Project-Level Policies"
$GroupProjPolicy.Location = New-Object System.Drawing.Point(10, 248)
$GroupProjPolicy.Size = New-Object System.Drawing.Size(1300, 70)
$GroupProjPolicy.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

$ButtonQueryProjectPolicy = New-Object System.Windows.Forms.Button
$ButtonQueryProjectPolicy.Text = "Query Project Policies"
$ButtonQueryProjectPolicy.Location = New-Object System.Drawing.Point(12, 26)
$ButtonQueryProjectPolicy.Size = New-Object System.Drawing.Size($btnW, $btnH)

$ButtonEnableProjectPolicy = New-Object System.Windows.Forms.Button
$ButtonEnableProjectPolicy.Text = "Enable Project Policies"
$ButtonEnableProjectPolicy.Location = New-Object System.Drawing.Point(202, 26)
$ButtonEnableProjectPolicy.Size = New-Object System.Drawing.Size($btnW, $btnH)

$ButtonDisableProjectPolicy = New-Object System.Windows.Forms.Button
$ButtonDisableProjectPolicy.Text = "Disable Project Policies"
$ButtonDisableProjectPolicy.Location = New-Object System.Drawing.Point(392, 26)
$ButtonDisableProjectPolicy.Size = New-Object System.Drawing.Size($btnW, $btnH)

$ButtonDeleteProjectPolicy = New-Object System.Windows.Forms.Button
$ButtonDeleteProjectPolicy.Text = "Delete Project Policies"
$ButtonDeleteProjectPolicy.Location = New-Object System.Drawing.Point(582, 26)
$ButtonDeleteProjectPolicy.Size = New-Object System.Drawing.Size($btnW, $btnH)

$GroupProjPolicy.Controls.AddRange(@($ButtonQueryProjectPolicy, $ButtonEnableProjectPolicy, $ButtonDisableProjectPolicy, $ButtonDeleteProjectPolicy))

# --- Group: Repo Policies ---
$GroupRepoPolicy = New-Object System.Windows.Forms.GroupBox
$GroupRepoPolicy.Text = "Repository-Level Policies"
$GroupRepoPolicy.Location = New-Object System.Drawing.Point(10, 324)
$GroupRepoPolicy.Size = New-Object System.Drawing.Size(1300, 70)
$GroupRepoPolicy.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

$ButtonQueryReposPolicy = New-Object System.Windows.Forms.Button
$ButtonQueryReposPolicy.Text = "Query Repo Policies"
$ButtonQueryReposPolicy.Location = New-Object System.Drawing.Point(12, 26)
$ButtonQueryReposPolicy.Size = New-Object System.Drawing.Size($btnW, $btnH)

$ButtonEnableReposPolicy = New-Object System.Windows.Forms.Button
$ButtonEnableReposPolicy.Text = "Enable Repo Policies"
$ButtonEnableReposPolicy.Location = New-Object System.Drawing.Point(202, 26)
$ButtonEnableReposPolicy.Size = New-Object System.Drawing.Size($btnW, $btnH)

$ButtonDisableReposPolicy = New-Object System.Windows.Forms.Button
$ButtonDisableReposPolicy.Text = "Disable Repo Policies"
$ButtonDisableReposPolicy.Location = New-Object System.Drawing.Point(392, 26)
$ButtonDisableReposPolicy.Size = New-Object System.Drawing.Size($btnW, $btnH)

$ButtonDeleteReposPolicy = New-Object System.Windows.Forms.Button
$ButtonDeleteReposPolicy.Text = "Delete Repo Policies"
$ButtonDeleteReposPolicy.Location = New-Object System.Drawing.Point(582, 26)
$ButtonDeleteReposPolicy.Size = New-Object System.Drawing.Size($btnW, $btnH)

$GroupRepoPolicy.Controls.AddRange(@($ButtonQueryReposPolicy, $ButtonEnableReposPolicy, $ButtonDisableReposPolicy, $ButtonDeleteReposPolicy))

# --- Group: Self-Approval ---
$GroupSelf = New-Object System.Windows.Forms.GroupBox
$GroupSelf.Text = "Self-Approval"
$GroupSelf.Location = New-Object System.Drawing.Point(10, 400)
$GroupSelf.Size = New-Object System.Drawing.Size(1300, 70)
$GroupSelf.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

$ButtonQuerySelfApproval = New-Object System.Windows.Forms.Button
$ButtonQuerySelfApproval.Text = "Query Self-Approval"
$ButtonQuerySelfApproval.Location = New-Object System.Drawing.Point(12, 26)
$ButtonQuerySelfApproval.Size = New-Object System.Drawing.Size($btnW, $btnH)

$ButtonEnableSelfApproval = New-Object System.Windows.Forms.Button
$ButtonEnableSelfApproval.Text = "Enable Self-Approval"
$ButtonEnableSelfApproval.Location = New-Object System.Drawing.Point(202, 26)
$ButtonEnableSelfApproval.Size = New-Object System.Drawing.Size($btnW, $btnH)

$ButtonDisableSelfApproval = New-Object System.Windows.Forms.Button
$ButtonDisableSelfApproval.Text = "Disable Self-Approval"
$ButtonDisableSelfApproval.Location = New-Object System.Drawing.Point(392, 26)
$ButtonDisableSelfApproval.Size = New-Object System.Drawing.Size($btnW, $btnH)

$GroupSelf.Controls.AddRange(@($ButtonQuerySelfApproval, $ButtonEnableSelfApproval, $ButtonDisableSelfApproval))

# Result textbox
$TextBoxResult = New-Object System.Windows.Forms.TextBox
$TextBoxResult.Location = New-Object System.Drawing.Point(10, 480)
$TextBoxResult.Size = New-Object System.Drawing.Size(1300, 420)
$TextBoxResult.Multiline = $true
$TextBoxResult.ScrollBars = "Vertical"
$TextBoxResult.BackColor = "Black"
$TextBoxResult.ForeColor = "White"
$TextBoxResult.Font = New-Object System.Drawing.Font("Consolas", 9)
$TextBoxResult.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$script:TextBoxResult = $TextBoxResult

# Add groups and result box to form
$Form.Controls.AddRange(@(
    $GroupEnv, $GroupProject, $GroupBranch, $GroupFilter,
    $GroupProjPolicy, $GroupRepoPolicy, $GroupSelf, $TextBoxResult
))

###===================================================================================================
## UI Event Handlers
###===================================================================================================
function Refresh-ConfigList {
    $ComboBoxConfig.Items.Clear()
    $script:ConfigEntries = @()
    $files = Get-ConfigFiles
    foreach ($f in $files) {
        $cfg = Read-Config -Path $f.FullName
        if (-not $cfg) { continue }
        $entries = Get-ConfigEntries -Obj $cfg
        foreach ($e in $entries) {
            $display = if ($e.Name) { $e.Name } else { $f.BaseName }
            [void]$ComboBoxConfig.Items.Add($display)
            $script:ConfigEntries += $e
        }
    }
    if ($ComboBoxConfig.Items.Count -gt 0) {
        $ComboBoxConfig.SelectedIndex = 0
    } else {
        Write-Log "No config files found in $($script:ConfigsPath). Use 'Edit / New Config' to create one."
    }
}

$ButtonLoadConfig.Add_Click({
    if ($ComboBoxConfig.SelectedIndex -lt 0) { return }
    if (-not $script:ConfigEntries -or $script:ConfigEntries.Count -le $ComboBoxConfig.SelectedIndex) {
        Write-Log "No config entry available. Click Refresh or re-open the form."
        return
    }
    $cfg = $script:ConfigEntries[$ComboBoxConfig.SelectedIndex]
    if ($ComboBoxApiVersion.SelectedItem -ne "auto") {
        $cfg.ApiVersion = $ComboBoxApiVersion.SelectedItem
    }
    Load-Config -Config $cfg
    if ([string]::IsNullOrWhiteSpace($script:MyPat)) {
        Write-Log "WARNING: PAT is empty. Edit the config and set the PAT field before loading projects."
    }
    Ensure-ApiVersion
    $ComboBoxApiVersion.SelectedItem = $script:ApiVersion
    $ButtonRefreshProjects.PerformClick()
})

$ButtonEditConfig.Add_Click({
    $files = Get-ConfigFiles
    $targetPath = $null
    if ($ComboBoxConfig.SelectedIndex -ge 0 -and $files.Count -gt $ComboBoxConfig.SelectedIndex) {
        $targetPath = $files[$ComboBoxConfig.SelectedIndex].FullName
    } else {
        $name = [Microsoft.VisualBasic.Interaction]::InputBox("New config name (filename without .json):", "New Config", "new-environment")
        if ([string]::IsNullOrWhiteSpace($name)) { return }
        $targetPath = Join-Path $script:ConfigsPath "$name.json"
    }

    $existing = if (Test-Path $targetPath) { Get-Content $targetPath -Raw } else { @"
{
  "Name": "$([System.IO.Path]::GetFileNameWithoutExtension($targetPath))",
  "ServerFQDN": "",
  "Collection": "DevOpsCollection",
  "PAT": "",
  "ApiVersion": "auto",
  "UseHttps": true,
  "Notes": ""
}
"@
    }

    $editor = New-Object System.Windows.Forms.Form
    $editor.Text = "Edit Config: $(Split-Path $targetPath -Leaf)"
    $editor.Size = New-Object System.Drawing.Size(700, 500)
    $editor.StartPosition = "CenterParent"
    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Multiline = $true
    $tb.ScrollBars = "Both"
    $tb.Dock = "Fill"
    $tb.Font = New-Object System.Drawing.Font("Consolas", 10)
    $tb.Text = $existing
    $editor.Controls.Add($tb)
    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = "Save"
    $ok.Dock = "Bottom"
    $ok.Height = 40
    $ok.Add_Click({
        try {
            $null = $tb.Text | ConvertFrom-Json
            Set-Content -Path $targetPath -Value $tb.Text -Encoding UTF8
            $editor.DialogResult = "OK"
            $editor.Close()
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Invalid JSON: $_", "Error")
        }
    })
    $editor.Controls.Add($ok)
    if ($editor.ShowDialog() -eq "OK") {
        Refresh-ConfigList
        Write-Log "Config saved: $targetPath"
    }
})

$ButtonRefreshProjects.Add_Click({
    if ([string]::IsNullOrWhiteSpace($script:ADOServerFQDN)) {
        Write-Log "Load a valid config first (ServerFQDN required)."
        return
    }
    if ([string]::IsNullOrWhiteSpace($script:MyPat)) {
        Write-Log "PAT is missing or empty. Edit the config and set the PAT field, then Load Config again."
        return
    }
    try {
        $ComboBoxProject.Items.Clear()
        $projects = Get-Projects | Sort-Object name
        foreach ($p in $projects) { [void]$ComboBoxProject.Items.Add($p.name) }
        $ComboBoxProject.Enabled = ($ComboBoxProject.Items.Count -gt 0)
        $ButtonSelectRepos.Enabled = $ComboBoxProject.Enabled
        $ObjBoxBranch.Enabled = $ComboBoxProject.Enabled
        Write-Log "Successfully retrieved $($projects.Count) project(s) from the collection."
    } catch {
        $msg = $_.Exception.Message
        if ($msg -match "401") {
            Write-Log "Authentication failed (401 Unauthorized) while loading projects."
            Write-Log "Check: (1) PAT is correct and not expired, (2) PAT has Code (Read) or Project and Team (Read) scope, (3) server URL is reachable with HTTPS."
        } else {
            Write-Log "Failed to retrieve projects: $msg"
        }
    }
})

function Update-SelectedReposDisplay {
    if ($script:SelectedRepoNames.Count -eq 0) {
        $LabelSelectedRepos.Text = "Repos: All"
    } else {
        $LabelSelectedRepos.Text = "Repos: $($script:SelectedRepoNames.Count) selected"
    }
}

$ButtonSelectRepos.Add_Click({
    if ([string]::IsNullOrWhiteSpace($ComboBoxProject.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Select a project first.")
        return
    }
    try {
        $script:RepoListCache = Get-Repositories -ProjectName $ComboBoxProject.Text
    } catch {
        Write-Log "Failed to load repos: $_"
        return
    }
    if ($script:RepoListCache.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No repositories found.")
        return
    }

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Select Repositories"
    $dlg.Size = New-Object System.Drawing.Size(500, 450)
    $dlg.StartPosition = "CenterParent"
    $lb = New-Object System.Windows.Forms.ListBox
    $lb.SelectionMode = "MultiExtended"
    $lb.Dock = "Fill"
    foreach ($r in $script:RepoListCache) { [void]$lb.Items.Add($r.name) }
    $dlg.Controls.Add($lb)
    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = "OK"
    $ok.Dock = "Bottom"
    $ok.Add_Click({ $dlg.DialogResult = "OK"; $dlg.Close() })
    $dlg.Controls.Add($ok)
    if ($dlg.ShowDialog() -eq "OK") {
        $script:SelectedRepoNames = @($lb.SelectedItems)
        $script:SelectedRepoIds = @($script:RepoListCache | Where-Object { $script:SelectedRepoNames -contains $_.name } | Select-Object -ExpandProperty id)
        Update-SelectedReposDisplay
    }
})

function Get-TargetRepos {
    param($Repos)
    if ($script:SelectedRepoIds.Count -gt 0) {
        return @($Repos | Where-Object { $script:SelectedRepoIds -contains $_.id })
    }
    return @($Repos)
}

$ButtonExit.Add_Click({ $Form.Close() })

###===================================================================================================
## Operations (with automatic backup before updates)
###===================================================================================================
$ButtonQueryAllRepo.Add_Click({
    if ([string]::IsNullOrWhiteSpace($ComboBoxProject.Text)) { Write-Log "Select a project."; return }
    $proj = $ComboBoxProject.Text
    Write-Log "Querying default branch for each repository in project '$proj'..."
    try {
        $repos = Get-Repositories -ProjectName $proj
        $targets = Get-TargetRepos -Repos $repos
        foreach ($r in $targets) {
            Write-Log "  Repository: $($r.name) | Default Branch: $($r.defaultBranch)"
        }
        Write-Log "Default branch query completed."
    } catch { Write-Log "Query failed: $_" }
})

$ButtonSetDefaultBranch.Add_Click({
    if ([string]::IsNullOrWhiteSpace($ComboBoxProject.Text)) { Write-Log "Select a project."; return }
    $proj = $ComboBoxProject.Text
    $branch = if ($ObjBoxBranch.SelectedItem -eq "Main") { "refs/heads/main" } else { "refs/heads/develop" }
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Set default branch to '$branch' for selected repositories in '$proj'?`n`nA backup will be created first.",
        "Confirm", "YesNo")
    if ($confirm -ne "Yes") { return }

    Backup-CurrentState -Operation "SetDefaultBranch" -ProjectName $proj | Out-Null
    try {
        $repos = Get-Repositories -ProjectName $proj
        $targets = Get-TargetRepos -Repos $repos
        $enc = [uri]::EscapeDataString($proj)
        foreach ($r in $targets) {
            $url = "$(Get-BaseUrl)/$($script:Collection)/$enc/_apis/git/repositories/$($r.id)?api-version=$($script:ApiVersion)"
            $body = @{ defaultBranch = $branch }
            Invoke-AdoRest -Uri $url -Method Patch -Body $body | Out-Null
            Write-Log "Updated default branch for repository '$($r.name)' to '$branch'."
        }
        Write-Log "Default branch update completed for selected repositories."
    } catch { Write-Log "Set default branch failed: $_" }
})

function Invoke-PolicyOperation {
    param(
        [string]$Action,          # Query / Enable / Disable / Delete
        [string]$Scope,           # Project or Repo
        [bool]$IsModify = $false
    )
    if ([string]::IsNullOrWhiteSpace($ComboBoxProject.Text)) { Write-Log "Select a project."; return }
    $proj = $ComboBoxProject.Text
    if ($IsModify) {
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "$Action $Scope policies in '$proj'?`n`nA backup will be created first.",
            "Confirm", "YesNo")
        if ($confirm -ne "Yes") { return }
        Backup-CurrentState -Operation "${Action}${Scope}Policy" -ProjectName $proj | Out-Null
    }

    try {
        $policies = Get-Policies -ProjectName $proj
        $typeFilter = $ObjBoxPolicyType.SelectedItem
        $typeMap = @{
            "Minimum number of reviewers" = "fa4e907d-c16b-4a4c-9dfa-4906e5d171dd"
            "Required reviewers"          = "fd2167ab-b0be-447a-8ec8-39368250530e"
        }
        $filterId = if ($typeMap.ContainsKey($typeFilter)) { $typeMap[$typeFilter] } else { $null }

        $enc = [uri]::EscapeDataString($proj)
        foreach ($p in $policies) {
            $scopes = @()
            if ($p.settings -and $p.settings.scope) { $scopes = @($p.settings.scope) }
            $hasRepo = $false
            foreach ($s in $scopes) {
                if ($s.repositoryId -or $s.refName) { $hasRepo = $true; break }
            }
            $isProjectLevel = -not $hasRepo
            if (($Scope -eq "Project" -and -not $isProjectLevel) -or ($Scope -eq "Repo" -and $isProjectLevel)) { continue }
            if ($filterId -and $p.type.id -ne $filterId) { continue }

            $policyUrl = "$(Get-BaseUrl)/$($script:Collection)/$enc/_apis/policy/configurations/$($p.id)?api-version=$($script:ApiVersion)"

            switch ($Action) {
                "Query" {
                    Write-Log "  Policy: $($p.type.displayName) | ID: $($p.id) | Enabled: $($p.isEnabled) | Blocking: $($p.isBlocking)"
                }
                "Enable" {
                    $body = $p | ConvertTo-Json -Depth 10 | ConvertFrom-Json
                    $body.isEnabled = $true
                    $body.isBlocking = $true
                    # Clean read-only properties
                    foreach ($prop in @("createdBy","createdDate","_links","isDeleted","isEnterpriseManaged","revision","id","url")) {
                        $body.PSObject.Properties.Remove($prop)
                    }
                    if ($body.type) {
                        $body.type.PSObject.Properties.Remove("url")
                        $body.type.PSObject.Properties.Remove("displayName")
                    }
                    Invoke-AdoRest -Uri $policyUrl -Method Put -Body $body | Out-Null
                    Write-Log "Enabled policy '$($p.type.displayName)' (ID: $($p.id))."
                }
                "Disable" {
                    $body = $p | ConvertTo-Json -Depth 10 | ConvertFrom-Json
                    $body.isEnabled = $false
                    foreach ($prop in @("createdBy","createdDate","_links","isDeleted","isEnterpriseManaged","revision","id","url")) {
                        $body.PSObject.Properties.Remove($prop)
                    }
                    if ($body.type) {
                        $body.type.PSObject.Properties.Remove("url")
                        $body.type.PSObject.Properties.Remove("displayName")
                    }
                    Invoke-AdoRest -Uri $policyUrl -Method Put -Body $body | Out-Null
                    Write-Log "Disabled policy '$($p.type.displayName)' (ID: $($p.id))."
                }
                "Delete" {
                    Invoke-AdoRest -Uri $policyUrl -Method Delete | Out-Null
                    Write-Log "Deleted policy '$($p.type.displayName)' (ID: $($p.id))."
                }
            }
        }
        Write-Log "$Action operation on $Scope-level policies completed."
    } catch {
        Write-Log "$Action $Scope policies failed: $_"
    }
}

$ButtonQueryProjectPolicy.Add_Click({ Invoke-PolicyOperation -Action "Query" -Scope "Project" })
$ButtonEnableProjectPolicy.Add_Click({ Invoke-PolicyOperation -Action "Enable" -Scope "Project" -IsModify $true })
$ButtonDisableProjectPolicy.Add_Click({ Invoke-PolicyOperation -Action "Disable" -Scope "Project" -IsModify $true })
$ButtonDeleteProjectPolicy.Add_Click({ Invoke-PolicyOperation -Action "Delete" -Scope "Project" -IsModify $true })

$ButtonQueryReposPolicy.Add_Click({ Invoke-PolicyOperation -Action "Query" -Scope "Repo" })
$ButtonEnableReposPolicy.Add_Click({ Invoke-PolicyOperation -Action "Enable" -Scope "Repo" -IsModify $true })
$ButtonDisableReposPolicy.Add_Click({ Invoke-PolicyOperation -Action "Disable" -Scope "Repo" -IsModify $true })
$ButtonDeleteReposPolicy.Add_Click({ Invoke-PolicyOperation -Action "Delete" -Scope "Repo" -IsModify $true })

# Self-approval: controls creatorVoteCounts on "Minimum number of reviewers" policies
# Enable  = creatorVoteCounts true  (PR creator vote counts toward minimum)
# Disable = creatorVoteCounts false (PR creator vote does not count)
function Invoke-SelfApprovalOperation {
    param(
        [string]$Action,   # Query / Enable / Disable
        [bool]$IsModify = $false
    )
    if ([string]::IsNullOrWhiteSpace($ComboBoxProject.Text)) { Write-Log "Select a project."; return }
    $proj = $ComboBoxProject.Text
    $minReviewersTypeId = "fa4e907d-c16b-4a4c-9dfa-4906e5d171dd"

    if ($IsModify) {
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "$Action self-approval (creatorVoteCounts) for Minimum number of reviewers policies in '$proj'?`n`nA backup will be created first.",
            "Confirm", "YesNo")
        if ($confirm -ne "Yes") { return }
        Backup-CurrentState -Operation "${Action}SelfApproval" -ProjectName $proj | Out-Null
    }

    try {
        $policies = Get-Policies -ProjectName $proj
        $enc = [uri]::EscapeDataString($proj)
        $count = 0
        foreach ($p in $policies) {
            if ($p.type.id -ne $minReviewersTypeId) { continue }
            $count++
            $policyUrl = "$(Get-BaseUrl)/$($script:Collection)/$enc/_apis/policy/configurations/$($p.id)?api-version=$($script:ApiVersion)"
            $current = $p.settings.creatorVoteCounts

            switch ($Action) {
                "Query" {
                    $selfStatus = if ($current) { "Yes – the PR creator can approve and it counts" } else { "No – the PR creator's approval does not count" }
                    $blockStatus = if ($p.isBlocking) { "Yes – PR cannot complete until requirements are met" } else { "No – advisory only" }
                    Write-Log "  Policy #$($p.id): Self-approval allowed? $selfStatus"
                    Write-Log "              Blocks completion? $blockStatus"
                }
                "Enable" {
                    $body = $p | ConvertTo-Json -Depth 10 | ConvertFrom-Json
                    if (-not $body.settings) { $body.settings = [PSCustomObject]@{} }
                    $body.settings.creatorVoteCounts = $true
                    foreach ($prop in @("createdBy","createdDate","_links","isDeleted","isEnterpriseManaged","revision","id","url")) {
                        $body.PSObject.Properties.Remove($prop)
                    }
                    if ($body.type) {
                        $body.type.PSObject.Properties.Remove("url")
                        $body.type.PSObject.Properties.Remove("displayName")
                    }
                    Invoke-AdoRest -Uri $policyUrl -Method Put -Body $body | Out-Null
                    Write-Log "Policy #$($p.id): Self-approval is now allowed (PR creator vote counts toward the minimum)."
                }
                "Disable" {
                    $body = $p | ConvertTo-Json -Depth 10 | ConvertFrom-Json
                    if (-not $body.settings) { $body.settings = [PSCustomObject]@{} }
                    $body.settings.creatorVoteCounts = $false
                    foreach ($prop in @("createdBy","createdDate","_links","isDeleted","isEnterpriseManaged","revision","id","url")) {
                        $body.PSObject.Properties.Remove($prop)
                    }
                    if ($body.type) {
                        $body.type.PSObject.Properties.Remove("url")
                        $body.type.PSObject.Properties.Remove("displayName")
                    }
                    Invoke-AdoRest -Uri $policyUrl -Method Put -Body $body | Out-Null
                    Write-Log "Policy #$($p.id): Self-approval is now blocked (PR creator vote does not count)."
                }
            }
        }
        if ($count -eq 0) {
            Write-Log "No 'Minimum number of reviewers' policies found in project '$proj'."
        } else {
            Write-Log "Self-approval $Action finished. Processed $count Minimum-number-of-reviewers policy/policies."
        }
    } catch {
        Write-Log "$Action self-approval failed: $_"
    }
}

$ButtonQuerySelfApproval.Add_Click({ Invoke-SelfApprovalOperation -Action "Query" })
$ButtonEnableSelfApproval.Add_Click({ Invoke-SelfApprovalOperation -Action "Enable" -IsModify $true })
$ButtonDisableSelfApproval.Add_Click({ Invoke-SelfApprovalOperation -Action "Disable" -IsModify $true })

###===================================================================================================
## Startup
###===================================================================================================
Refresh-ConfigList
if ($ComboBoxConfig.Items.Count -eq 0) {
    Write-Log "No configuration files found."
    Write-Log "Use the 'Edit / New Config' button to create a config file for your Azure DevOps Server environment."
    Write-Log "Sample configs are in the 'configs' folder."
}

[void]$Form.ShowDialog()
