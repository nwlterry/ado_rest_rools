###===================================================================================================
##Load Windows Form Type
##---------------------------------------------------------------------------------------------------
#Add Windows Form Type
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#Add Visual Basic Form Type
[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

# If a previous run left a Form object in the session, close and remove it to avoid duplicate UI
try {
    if (Get-Variable -Name Form -Scope Script -ErrorAction SilentlyContinue) {
        if ($null -ne $Form) {
            try {
                if ($Form -is [System.Array]) { foreach ($f in $Form) { try { $f.Close(); $f.Dispose() } catch { } } }
                else { try { $Form.Close(); $Form.Dispose() } catch { } }
            } catch { }
        }
        Remove-Variable -Name Form -Scope Script -ErrorAction SilentlyContinue
    }
} catch { }

# Create the main Form if it doesn't exist
if (-not (Get-Variable -Name Form -Scope Script -ErrorAction SilentlyContinue)) {
    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = 'Azure DevOps Server REST API Tool'
    $Form.StartPosition = 'CenterScreen'
    $Form.ClientSize = New-Object System.Drawing.Size(1280,900)
    $Form.Font = New-Object System.Drawing.Font('Segoe UI',9)
}

###===================================================================================================
##Define Log Output
##---------------------------------------------------------------------------------------------------
#Define Log File Location
$logPath = $PSScriptRoot
$logDate = Get-date -Format yyyy-MM-dd
$logName = "ADO_REST_Tools_Result_$($logDate).log"
$logFile = "$logPath\$logName"

##---------------------------------------------------------------------------------------------------
#Check Log File Existence
if (-not (Test-Path -Path $logfile)) {
    New-Item -Path $logFile -ItemType File
}

###===================================================================================================
##---------------------------------------------------------------------------------------------------
#Policy Type
#"Minimum number of reviewers" id : fa4e907d-c16b-4a4c-9dfa-4906e5d171dd
#"Required reviewers" id : fd2167ab-b0be-447a-8ec8-39368250530e

###===================================================================================================
##Define Azure Devops Server
##---------------------------------------------------------------------------------------------------
#Define Azure Devops Server Information
$ADOServerFQDN = ""
$collection = "DevOpsCollection"
$projectName = ""
$repoName = ""
#$refNameMain = "refs/heads/main"
#$refNameDevelop = "refs/heads/develop"

##---------------------------------------------------------------------------------------------------
#Define Azure Devops Server Access Token
$MyPat = ''
$B64Pat = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes(":$MyPat"))
$headers = @{
    'Authorization' = 'Basic ' + $B64Pat
    'Content-Type' = 'application/json'
}
$ComboBoxProject = New-Object System.Windows.Forms.ComboBox
$ComboBoxProject.Location = New-Object System.Drawing.Point(130,18)
$ComboBoxProject.Size = New-Object System.Drawing.Size(300,24)
$ComboBoxProject.DropDownStyle = 'DropDownList'
$ComboBoxProject.Enabled = $false

$script:RepoListCache = @()
$script:SelectedRepoIds = @()
$script:SelectedRepoNames = @()

$ComboBoxProject.BringToFront()

###===================================================================================================
##Define Output Text Box
##---------------------------------------------------------------------------------------------------
#Define Result Output Text Box
$TextBoxResult = New-Object System.Windows.Forms.TextBox
$TextBoxResult.Location = New-Object System.Drawing.Point(20,200)
$TextBoxResult.Size = New-Object System.Drawing.Size(960,440)
$TextBoxResult.ForeColor = "White"
$TextBoxResult.BackColor = "Black"
$TextBoxResult.ScrollBars = "Vertical"
$TextBoxResult.MultiLine = $True

# Make result box resize with the form
$TextBoxResult.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

$null = $null # Defer adding the result TextBox until after layout so it's only added once

###===================================================================================================
##Define Form Label
##---------------------------------------------------------------------------------------------------
#Define Project Input Text Box Label
$LabelTextBoxProject = New-Object System.Windows.Forms.Label
$LabelTextBoxProject.Location = New-Object System.Drawing.Point(20,20)
$LabelTextBoxProject.Size = New-Object System.Drawing.Size(100,20)
$LabelTextBoxProject.Text = 'Project name:'

# (moved into `$TopTable`) $LabelTextBoxProject will be added to the top layout panel later
#Define Project Input Text Box Label
$LabelSelfApproval = New-Object System.Windows.Forms.Label
$LabelSelfApproval.Location = New-Object System.Drawing.Point(500,30)
$LabelSelfApproval.Size = New-Object System.Drawing.Size(100,40)
$LabelSelfApproval.Text = 'Enable/Disable Self Approval:'

# (moved into `$TopTable`) $LabelSelfApproval will be added to the top layout panel later
#Define Using Project Label
$LabelUsingProject = New-Object System.Windows.Forms.Label
$LabelUsingProject.Location = New-Object System.Drawing.Point(20,80)
$LabelUsingProject.Size = New-Object System.Drawing.Size(100,20)
$LabelUsingProject.Text = 'Using Project:'

# (moved into `$TopTable`) $LabelUsingProject will be added to the top layout panel later
#Define Display Project Label
$LabelDisplayProject = New-Object System.Windows.Forms.Label
$LabelDisplayProject.Location = New-Object System.Drawing.Point(20,100)
$LabelDisplayProject.Size = New-Object System.Drawing.Size(100,20)
$LabelDisplayProject.Text = ''

$LabelRepos = New-Object System.Windows.Forms.Label
$LabelRepos.Location = New-Object System.Drawing.Point(20,120)
$LabelRepos.Size = New-Object System.Drawing.Size(100,20)
$LabelRepos.Text = 'Selected repos:'

$LabelSelectedRepos = New-Object System.Windows.Forms.Label
$LabelSelectedRepos.Location = New-Object System.Drawing.Point(130,120)
$LabelSelectedRepos.Size = New-Object System.Drawing.Size(420,20)
$LabelSelectedRepos.Text = 'All'
$LabelSelectedRepos.AutoEllipsis = $true

$TextBoxSelectedReposSelf = New-Object System.Windows.Forms.TextBox
$TextBoxSelectedReposSelf.Location = New-Object System.Drawing.Point(620,120)
$TextBoxSelectedReposSelf.Size = New-Object System.Drawing.Size(320,22)
$TextBoxSelectedReposSelf.Text = 'All'
$TextBoxSelectedReposSelf.ReadOnly = $true
$TextBoxSelectedReposSelf.BorderStyle = 'FixedSingle'

$ButtonSelectRepos = New-Object System.Windows.Forms.Button
$ButtonSelectRepos.Location = New-Object System.Drawing.Point(560,116)
$ButtonSelectRepos.Size = New-Object System.Drawing.Size(120,28)
$ButtonSelectRepos.Text = 'Select Repos'
$ButtonSelectRepos.Enabled = $false
$ButtonSelectRepos.Add_Click({ showRepoSelector })

# (moved into `$TopTable`) $LabelDisplayProject will be added to the top layout panel later
###===================================================================================================
##Define Form Dropdown Menu
##---------------------------------------------------------------------------------------------------
#Define Branch Dropdown menu label
$LabelObjBoxBranch = New-Object System.Windows.Forms.Label
$LabelObjBoxBranch.Location = New-Object System.Drawing.Point(20,100)
$LabelObjBoxBranch.Size = New-Object System.Drawing.Size(100,20)
$LabelObjBoxBranch.Text = 'Set default branch:'

# (moved into `$TopTable`) $LabelObjBoxBranch will be added to the top layout panel later
#Define Branch Dropdown menu
$ObjBoxBranch = New-Object System.Windows.Forms.ComboBox
$ObjBoxBranch.Location = New-Object System.Drawing.Point(130,100)
$ObjBoxBranch.Size = New-Object System.Drawing.Size(120,24)
$ObjBoxBranch.Enabled = $false

#Define Branch Dropdoen menu value
$ObjBoxBranch.Items.Add('Develop') | Out-Null
$ObjBoxBranch.Items.Add('Main') | Out-Null

# (moved into `$TopTable`) $ObjBoxBranch will be added to the top layout panel later
##---------------------------------------------------------------------------------------------------
# Load projects from Azure DevOps and populate project dropdown
function loadProjects {
    try {
        $ComboBoxProject.Items.Clear()
        if ([string]::IsNullOrWhiteSpace($ADOServerFQDN) -or [string]::IsNullOrWhiteSpace($collection)) {
            $TextBoxResult.AppendText("Please set `ADOServerFQDN` and `collection` before loading projects.`r`n")
            return
        }
        $projectsUrl = "https://$ADOServerFQDN/$collection/_apis/projects?api-version=6.0"
        $projectsResponse = Invoke-RestMethod -Method Get -Uri $projectsUrl -Headers $headers
        $projects = $projectsResponse.value | Sort-Object -Property name
        foreach ($p in $projects) {
            $ComboBoxProject.Items.Add($p.name) | Out-Null
        }
        if ($ComboBoxProject.Items.Count -gt 0) { $ComboBoxProject.Enabled = $true }
    } catch {
        $TextBoxResult.AppendText("Failed to load projects: $_`r`n")
    }
}

function updateSelectedReposDisplay {
    if ($null -eq $LabelSelectedRepos) { return }
    if ($null -eq $script:SelectedRepoNames -or $script:SelectedRepoNames.Count -eq 0) {
        $LabelSelectedRepos.Text = 'All'
        if ($null -ne $TextBoxSelectedReposSelf) { $TextBoxSelectedReposSelf.Text = 'All' }
        return
    }
    if ($script:SelectedRepoNames.Count -le 3) {
        $LabelSelectedRepos.Text = ($script:SelectedRepoNames -join ', ')
        if ($null -ne $TextBoxSelectedReposSelf) { $TextBoxSelectedReposSelf.Text = ($script:SelectedRepoNames -join ', ') }
    } else {
        $LabelSelectedRepos.Text = "$($script:SelectedRepoNames.Count) selected"
        if ($null -ne $TextBoxSelectedReposSelf) { $TextBoxSelectedReposSelf.Text = "$($script:SelectedRepoNames.Count) selected" }
    }
}

function loadReposForProject {
    $script:RepoListCache = @()
    $script:SelectedRepoIds = @()
    $script:SelectedRepoNames = @()
    updateSelectedReposDisplay

    if ([string]::IsNullOrWhiteSpace($ComboBoxProject.Text)) { return }
    if ([string]::IsNullOrWhiteSpace($ADOServerFQDN) -or [string]::IsNullOrWhiteSpace($collection)) { return }

    try {
        $projectName = $ComboBoxProject.Text.ToString()
        $projectNameEnc = [uri]::EscapeDataString($projectName)
        $reposUrl = "https://$ADOServerFQDN/$collection/$projectNameEnc/_apis/git/repositories?api-version=6.0"
        $reposResponse = Invoke-RestMethod -Method Get -Uri $reposUrl -Headers $headers
        if ($null -ne $reposResponse -and $null -ne $reposResponse.value) {
            $script:RepoListCache = @($reposResponse.value | Sort-Object -Property name)
        }
    } catch {
        $TextBoxResult.AppendText("Failed to load repos: $_`r`n")
    }
}

function getTargetRepos {
    param(
        [Parameter(Mandatory=$true)]
        $Repos
    )
    if ($null -ne $script:SelectedRepoIds -and $script:SelectedRepoIds.Count -gt 0) {
        return @($Repos | Where-Object { $script:SelectedRepoIds -contains $_.id })
    }
    return @($Repos)
}

function showRepoSelector {
    if ($null -eq $ComboBoxProject -or [string]::IsNullOrWhiteSpace($ComboBoxProject.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Please select a project first.","Repo Selection")
        return
    }

    if ($null -eq $script:RepoListCache -or $script:RepoListCache.Count -eq 0) {
        loadReposForProject
    }

    if ($null -eq $script:RepoListCache -or $script:RepoListCache.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No repos found for the selected project.","Repo Selection")
        return
    }

    $repoForm = New-Object System.Windows.Forms.Form
    $repoForm.Text = "Select Repos"
    $repoForm.StartPosition = 'CenterParent'
    $repoForm.ClientSize = New-Object System.Drawing.Size(520,480)
    $repoForm.Font = New-Object System.Drawing.Font('Segoe UI',9)

    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Location = New-Object System.Drawing.Point(10,10)
    $listBox.Size = New-Object System.Drawing.Size(500,380)
    $listBox.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended
    $listBox.DisplayMember = 'name'

    foreach ($repo in $script:RepoListCache) {
        [void]$listBox.Items.Add($repo)
    }

    for ($i = 0; $i -lt $listBox.Items.Count; $i++) {
        $item = $listBox.Items[$i]
        if ($null -ne $item -and ($script:SelectedRepoIds -contains $item.id)) {
            $listBox.SetSelected($i, $true)
        }
    }

    $buttonSelectAll = New-Object System.Windows.Forms.Button
    $buttonSelectAll.Text = 'Select All'
    $buttonSelectAll.Location = New-Object System.Drawing.Point(10,400)
    $buttonSelectAll.Size = New-Object System.Drawing.Size(100,30)
    $buttonSelectAll.Add_Click({
        for ($i = 0; $i -lt $listBox.Items.Count; $i++) { $listBox.SetSelected($i, $true) }
    })

    $buttonClear = New-Object System.Windows.Forms.Button
    $buttonClear.Text = 'Clear'
    $buttonClear.Location = New-Object System.Drawing.Point(120,400)
    $buttonClear.Size = New-Object System.Drawing.Size(100,30)
    $buttonClear.Add_Click({ $listBox.ClearSelected() })

    $buttonOk = New-Object System.Windows.Forms.Button
    $buttonOk.Text = 'OK'
    $buttonOk.Location = New-Object System.Drawing.Point(320,400)
    $buttonOk.Size = New-Object System.Drawing.Size(90,30)
    $buttonOk.Add_Click({
        $script:SelectedRepoIds = @()
        $script:SelectedRepoNames = @()
        foreach ($item in $listBox.SelectedItems) {
            if ($null -ne $item) {
                $script:SelectedRepoIds += $item.id
                $script:SelectedRepoNames += $item.name
            }
        }
        updateSelectedReposDisplay
        $repoForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $repoForm.Close()
    })

    $buttonCancel = New-Object System.Windows.Forms.Button
    $buttonCancel.Text = 'Cancel'
    $buttonCancel.Location = New-Object System.Drawing.Point(420,400)
    $buttonCancel.Size = New-Object System.Drawing.Size(90,30)
    $buttonCancel.Add_Click({
        $repoForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $repoForm.Close()
    })

    $repoForm.Controls.Add($listBox)
    $repoForm.Controls.Add($buttonSelectAll)
    $repoForm.Controls.Add($buttonClear)
    $repoForm.Controls.Add($buttonOk)
    $repoForm.Controls.Add($buttonCancel)

    [void]$repoForm.ShowDialog($Form)
}

##---------------------------------------------------------------------------------------------------
#Define Policy Type Dropdown menu label
$LabelObjBoxPolicyType = New-Object System.Windows.Forms.Label
$LabelObjBoxPolicyType.Location = New-Object System.Drawing.Point(620,95)
$LabelObjBoxPolicyType.Size = New-Object System.Drawing.Size(70,20)
$LabelObjBoxPolicyType.Text = 'Policy Type:'
$LabelObjBoxPolicyType.Enabled = $false

# (moved into `$TopTable`) $LabelObjBoxPolicyType will be added to the top layout panel later
#Define Policy Type Dropdown menu
$ObjBoxPolicyType = New-Object System.Windows.Forms.ComboBox
$ObjBoxPolicyType.Location = New-Object System.Drawing.Point(690,95)
$ObjBoxPolicyType.Size = New-Object System.Drawing.Size(160,40)
$ObjBoxPolicyType.Enabled = $false

#Define Project Type Dropdown menu value
$ObjBoxPolicyType.Items.Add('Minimum number of reviewers') | Out-Null
$ObjBoxPolicyType.Items.Add('Main') | Out-Null

# (moved into `$TopTable`) $ObjBoxPolicyType will be added to the top layout panel later
###===================================================================================================
##Define Form Button
##---------------------------------------------------------------------------------------------------
#Define Button of query all repos
$ButtonQueryRepos = New-Object System.Windows.Forms.Button
$ButtonQueryRepos.Location = New-Object System.Drawing.Point(20,160)
$ButtonQueryRepos.Size = New-Object System.Drawing.Size(200,40)
$ButtonQueryRepos.Text = "Query Default Branch"
$ButtonQueryRepos.Enabled = $false
$ButtonQueryRepos.Add_Click( { queryAllRepo } )

# (moved into `$MiddlePanel`) $ButtonQueryRepos will be added to the middle layout panel later
##---------------------------------------------------------------------------------------------------
#Define Button to refresh projects list
$ButtonRefreshProjects = New-Object System.Windows.Forms.Button
$ButtonRefreshProjects.Location = New-Object System.Drawing.Point(340,20)
$ButtonRefreshProjects.Size = New-Object System.Drawing.Size(120,30)
$ButtonRefreshProjects.Text = "Refresh Projects"
$ButtonRefreshProjects.Enabled = $true
$ButtonRefreshProjects.Add_Click( {
    $ButtonRefreshProjects.Enabled = $false
    loadProjects
    $ButtonRefreshProjects.Enabled = $true
})

# (moved into `$TopTable`) $ButtonRefreshProjects will be added to the top layout panel later
$ButtonRefreshProjects.BringToFront()
##---------------------------------------------------------------------------------------------------
#Define Button of update repos default branch
$ButtonSetDefaultBranch = New-Object System.Windows.Forms.Button
$ButtonSetDefaultBranch.Location = New-Object System.Drawing.Point(240,160)
$ButtonSetDefaultBranch.Size = New-Object System.Drawing.Size(200,40)
$ButtonSetDefaultBranch.Text = "Set Default Branch"
$ButtonSetDefaultBranch.Enabled = $false
$ButtonSetDefaultBranch.Add_Click( { setDefaultBranch } )

# (moved into `$MiddlePanel`) $ButtonSetDefaultBranch will be added to the middle layout panel later
##---------------------------------------------------------------------------------------------------
#Define Button of query project policy
$ButtonQueryProjectPolicy = New-Object System.Windows.Forms.Button
$ButtonQueryProjectPolicy.Location = New-Object System.Drawing.Point(140,80)
$ButtonQueryProjectPolicy.Size = New-Object System.Drawing.Size(100,40)
$ButtonQueryProjectPolicy.Text = "Query Project Policy"
$ButtonQueryProjectPolicy.Enabled = $false
$ButtonQueryProjectPolicy.Add_Click( { queryProjectPolicy } )

# (moved into `$MiddlePanel`) $ButtonQueryProjectPolicy will be added to the middle layout panel later
##---------------------------------------------------------------------------------------------------
#Define Button of disable project policy
$ButtonDisableProjectPolicy = New-Object System.Windows.Forms.Button
$ButtonDisableProjectPolicy.Location = New-Object System.Drawing.Point(260,80)
$ButtonDisableProjectPolicy.Size = New-Object System.Drawing.Size(100,40)
$ButtonDisableProjectPolicy.Text = "Disable Project Policy"
$ButtonDisableProjectPolicy.Enabled = $false
$ButtonDisableProjectPolicy.Add_Click( { disableProjectPolicy } )

# (moved into `$MiddlePanel`) $ButtonDisableProjectPolicy will be added to the middle layout panel later
##---------------------------------------------------------------------------------------------------
#Define Button of enable project policy
$ButtonEnableProjectPolicy = New-Object System.Windows.Forms.Button
$ButtonEnableProjectPolicy.Location = New-Object System.Drawing.Point(380,80)
$ButtonEnableProjectPolicy.Size = New-Object System.Drawing.Size(100,40)
$ButtonEnableProjectPolicy.Text = "Enable Project Policy"
$ButtonEnableProjectPolicy.Enabled = $false
$ButtonEnableProjectPolicy.Add_Click( { enableProjectPolicy } )

# (moved into `$MiddlePanel`) $ButtonEnableProjectPolicy will be added to the middle layout panel later
##---------------------------------------------------------------------------------------------------
#Define Button of delete project policy
$ButtonDeleteProjectPolicy = New-Object System.Windows.Forms.Button
$ButtonDeleteProjectPolicy.Location = New-Object System.Drawing.Point(500,80)
$ButtonDeleteProjectPolicy.Size = New-Object System.Drawing.Size(100,40)
$ButtonDeleteProjectPolicy.Text = "Delete Project Policy"
$ButtonDeleteProjectPolicy.Enabled = $false
$ButtonDeleteProjectPolicy.Add_Click( { deleteProjectPolicy } )

# (moved into `$MiddlePanel`) $ButtonDeleteProjectPolicy will be added to the middle layout panel later
##---------------------------------------------------------------------------------------------------
#Define Button of query repositories policy
$ButtonQueryReposPolicy = New-Object System.Windows.Forms.Button
$ButtonQueryReposPolicy.Location = New-Object System.Drawing.Point(140,140)
$ButtonQueryReposPolicy.Size = New-Object System.Drawing.Size(100,40)
$ButtonQueryReposPolicy.Text = "Query Repos Policy"
$ButtonQueryReposPolicy.Enabled = $false
$ButtonQueryReposPolicy.Add_Click( { queryReposPolicy } )

# (moved into `$MiddlePanel`) $ButtonQueryReposPolicy will be added to the middle layout panel later
##---------------------------------------------------------------------------------------------------
#Define Button of disable repositories policy
$ButtonDisableReposPolicy = New-Object System.Windows.Forms.Button
$ButtonDisableReposPolicy.Location = New-Object System.Drawing.Point(260,140)
$ButtonDisableReposPolicy.Size = New-Object System.Drawing.Size(100,40)
$ButtonDisableReposPolicy.Text = "Disable Repos Policy"
$ButtonDisableReposPolicy.Enabled = $false
$ButtonDisableReposPolicy.Add_Click( { disableReposPolicy } )

# (moved into `$MiddlePanel`) $ButtonDisableReposPolicy will be added to the middle layout panel later
##---------------------------------------------------------------------------------------------------
#Define Button of enable repositories policy
$ButtonEnableReposPolicy = New-Object System.Windows.Forms.Button
$ButtonEnableReposPolicy.Location = New-Object System.Drawing.Point(380,140)
$ButtonEnableReposPolicy.Size = New-Object System.Drawing.Size(100,40)
$ButtonEnableReposPolicy.Text = "Enable Project Policy"
$ButtonEnableReposPolicy.Enabled = $false
$ButtonEnableReposPolicy.Add_Click( { enableReposPolicy } )

# (moved into `$MiddlePanel`) $ButtonEnableReposPolicy will be added to the middle layout panel later
##---------------------------------------------------------------------------------------------------
#Define Button of delete repositories policy
$ButtonDeleteReposPolicy = New-Object System.Windows.Forms.Button
$ButtonDeleteReposPolicy.Location = New-Object System.Drawing.Point(500,140)
$ButtonDeleteReposPolicy.Size = New-Object System.Drawing.Size(100,40)
$ButtonDeleteReposPolicy.Text = "Delete Repos Policy"
$ButtonDeleteReposPolicy.Enabled = $false
$ButtonDeleteReposPolicy.Add_Click( { deleteReposPolicy } )

# (moved into `$MiddlePanel`) $ButtonDeleteReposPolicy will be added to the middle layout panel later
##---------------------------------------------------------------------------------------------------
#Define Button of query repositories self approval
$ButtonQuerySelfApproval = New-Object System.Windows.Forms.Button
$ButtonQuerySelfApproval.Location = New-Object System.Drawing.Point(620,20)
$ButtonQuerySelfApproval.Size = New-Object System.Drawing.Size(100,40)
$ButtonQuerySelfApproval.Text = "Query Self Approval"
$ButtonQuerySelfApproval.Enabled = $false
$ButtonQuerySelfApproval.Add_Click( { querySelfApproval } )

# (moved into `$TopTable`) $ButtonQuerySelfApproval will be added to the top layout panel later
##---------------------------------------------------------------------------------------------------
#Define Button of enable repositories self approval
$ButtonEnableSelfApproval = New-Object System.Windows.Forms.Button
$ButtonEnableSelfApproval.Location = New-Object System.Drawing.Point(740,20)
$ButtonEnableSelfApproval.Size = New-Object System.Drawing.Size(100,40)
$ButtonEnableSelfApproval.Text = "Enable Self Approval"
$ButtonEnableSelfApproval.Enabled = $false
$ButtonEnableSelfApproval.Add_Click( { enableSelfApproval } )

# (moved into `$TopTable`) $ButtonEnableSelfApproval will be added to the top layout panel later
##---------------------------------------------------------------------------------------------------
#Define Button of disable repositories self approval
$ButtonDisableSelfApproval = New-Object System.Windows.Forms.Button
$ButtonDisableSelfApproval.Location = New-Object System.Drawing.Point(860,20)
$ButtonDisableSelfApproval.Size = New-Object System.Drawing.Size(100,40)
$ButtonDisableSelfApproval.Text = "Disable Self Approval"
$ButtonDisableSelfApproval.Enabled = $false
$ButtonDisableSelfApproval.Add_Click( { disableSelfApproval } )

# (moved into `$TopTable`) $ButtonDisableSelfApproval will be added to the top layout panel later
##---------------------------------------------------------------------------------------------------
#Define Button of exit form
$ButtonExit = New-Object System.Windows.Forms.Button
$ButtonExit.Location = New-Object System.Drawing.Point(860,140)
$ButtonExit.Size = New-Object System.Drawing.Size(100,40)
$ButtonExit.Text = "Exit"
$ButtonExit.Add_Click({$Form.Close()})

# (moved into `$MiddlePanel`) $ButtonExit will be added to the middle layout panel later
##Common Function
##---------------------------------------------------------------------------------------------------
#Define function to write log
function writeLog
{
    Add-content $logFile -value $logOutput
}

###===================================================================================================
##Define Default Branch Function
##---------------------------------------------------------------------------------------------------
#Define Function to query all repos
function queryAllRepo {
    $projectName = $ComboBoxProject.Text.ToString()
    $projectNameEnc = [uri]::EscapeDataString($projectName)
    $timeStamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")

    $TextBoxResult.Clear()

    $formTitle = "Query default branch of all repositories in project $($projectName) Result"
    $Form.Text = $formTitle

    $reposUrl = "https://$ADOServerFQDN/$collection/$projectNameEnc/_apis/git/repositories?api-version=6.0"
    $reposResponse = Invoke-RestMethod -Method Get -Uri $reposUrl -Headers $headers

    $targetRepos = getTargetRepos -Repos $reposResponse.value

    $TextBoxResult.Text = "$($timeStamp) The default branch of each repository:"

    foreach ($repo in $targetRepos) {
        $timeStamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
        $TextBoxResult.AppendText("`r`n")
        $TextBoxResult.Appendtext("$($timestamp) Project: $($projectName), Repository: $($repo.name), ID: $($repo.id), Default Branch: $($repo.defaultBranch)")
    }
    [System.Windows.Forms.MessageBox]::Show("Query of default branch in all repositories`n of project $projectName completed","Process Result")
    [string[]]$logOutput = $TextBoxResult.Text
    $logOutput = $logOutput.Replace("`r`n", "`r`n")
    writeLog
}

##---------------------------------------------------------------------------------------------------
#Define Set Default Branch Function
function setDefaultBranch {
    $projectName = $ComboBoxProject.Text.ToString()
    $projectNameEnc = [uri]::EscapeDataString($projectName)
    $timeStamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")

    $TextBoxResult.Clear()

    $formTitle = "Setting default branch on all repositories in Project $($projectName)"
    $Form.Text = $formTitle

    if ($ObjBoxBranch.SelectedItem -eq "Main"){
        $jsonData = @{"defaultBranch" = "refs/heads/main"}
    } else {
        $jsonData = @{"defaultBranch" = "refs/heads/develop"}
    }

    [string[]]$toBranch = $jsonData.Values
    if ($toBranch[0].Length -eq 15) {
        $toBranch = $toBranch.Substring(11,4)
    } else {
        $toBranch = $toBranch.Substring(11,7)
    }

    $reposUrl = "https://$ADOServerFQDN/$collection/$projectNameEnc/_apis/git/repositories?api-version=6.0"
    $reposResponse = Invoke-RestMethod -Method Get -Uri $reposUrl -Headers $headers

    $targetRepos = getTargetRepos -Repos $reposResponse.value

    $TextBoxResult.Text = "$($timeStamp) Setting default branch to $($toBranch.ToUpper()) of each repos:"

    foreach ($repo in $targetRepos) {
        $timeStamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
        $TextBoxResult.AppendText("`r`n")
        $TextBoxResult.AppendText("$($timeStamp) Setting default branch to $($toBranch.ToUpper()) of repository $($repo.name):")
        $TextBoxResult.AppendText("`r`n")
        $TextBoxResult.AppendText("$($timeStamp) Project: $($projectName), Repository: $($repo.name), ID: $($repo.id), Default Branch: $($repo.defaultBranch)")
        $branchUrl = "https://$ADOServerFQDN/$collection/$projectNameEnc/_apis/git/repositories/$($repo.id)?api-version=6.0"
        $TextBoxResult.AppendText("`r`n")
        $TextBoxResult.AppendText("$($timeStamp) Setting the default branch......")
        #$branchRESTPatch = Invoke-RestMethod -Uri $branchUrl -Method Patch -Headers $headers -ContentType "application/json" -Body ($jsonData | ConvertTo-Json)
        Invoke-RestMethod -Uri $branchUrl -Method Patch -Headers $headers -ContentType "application/json" -Body ($jsonData | ConvertTo-Json)
        $branchRESTGet = Invoke-RestMethod -Uri $branchUrl -Method Get -Headers $headers
        $TextBoxResult.AppendText("`r`n")
        $TextBoxResult.AppendText("$($timeStamp) Project: $($projectName), Repository: $($repo.name), ID: $($repo.id), Default Branch: $($branchRESTGet.defaultBranch)")
        [System.Windows.Forms.MessageBox]::Show("Setting default branch in repository`n $($repo.name) of project $projectName completed","Process Result")
    }

    [System.Windows.Forms.MessageBox]::Show("Setting default branch on all repositories of project $projectName completed","Process Result")
    [string[]]$logOutput = $TextBoxResult.Text
    $logOutput = $logOutput.Replace("`r`n", "`r`n")
    writeLog
}

###===================================================================================================
##Define Project Base Policy Function
##---------------------------------------------------------------------------------------------------
#Define Function to query project base policy
function queryProjectPolicy {
    $projectName = $ComboBoxProject.Text.ToString()
    $projectNameEnc = [uri]::EscapeDataString($projectName)
    $timeStamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")

    $TextBoxResult.Clear()

    $formTitle = "Query repository policy in project $($projectName) Result"
    $Form.Text = $formTitle

    $TextBoxResult.Text = "$($timeStamp) The repository policy on each repository:"
    $TextBoxResult.AppendText("`r`n")
    $TextBoxResult.Appendtext("$($timeStamp) Project: $($projectName)")
    $TextBoxResult.Appendtext("`n`n")

    $policiesObject = [System.Collections.Generic.List[object]]::new()

    #$reposUrl = "https://$ADOServerFQDN/$collection/$projectName/_apis/git/repositories?api-version=6.0"
    #$reposResponse = Invoke-RestMethod -Method Get -Uri $reposUrl -Headers $headers

    $policiesUrl = "https://$ADOServerFQDN/$collection/$projectNameEnc/_apis/policy/configurations?api-version=6.0"
    $policiesResponse = Invoke-RestMethod -Method Get -Uri $policiesUrl -Headers $headers

    foreach ($policy in $policiesResponse.value) {
        # settings.scope can be a single object or an array; project-level policies have no repo/ref scope
        $scopes = @()
        if ($null -ne $policy.settings -and $null -ne $policy.settings.scope) {
            $scopes = @($policy.settings.scope)
        }

        $hasRepoScope = $false
        foreach ($s in $scopes) {
            if ($null -ne $s.repositoryId -or $null -ne $s.refName) {
                $hasRepoScope = $true
                break
            }
        }

        if (-not $hasRepoScope) {
            $policyTemp = [pscustomobject]@{
                policyType = $($policy.Type.displayName)
                policyID = $($policy.id)
                policyEnabled = $($policy.isEnabled)
            }
            if ($null -ne $policyTemp.policyEnabled) {
                $policiesObject.add($policyTemp)
            }
        }
    }

    $policiesOutput = $policiesObject | Sort-Object #-Unique -Property policyRepo #policyProject,policyID,policyType,policyEabled

    foreach ($policyOutput in $policiesOutput) {
        $TextBoxResult.AppendText("`r`n")
        $TextBoxResult.Appendtext("$($timeStamp) Policy Type : $($policyOutput.policyType)")
        $TextBoxResult.AppendText("`r`n")
        $TextBoxResult.Appendtext("$($timeStamp) Policy ID : $($policyOutput.policyID)")
        $TextBoxResult.AppendText("`r`n")
        $TextBoxResult.Appendtext("$($timeStamp) Policy Enabled : $($policyOutput.policyEnabled)")
        $TextBoxResult.AppendText("`r`n")
    }

    $TextBoxResult.AppendText("`r`n")
    $TextBoxResult.AppendText("Query repositories policy of project $projectName completed")

    [System.Windows.Forms.MessageBox]::Show("Query reporistoty policies of project`n$projectName completed","Process Result")
    [string[]]$logOutput = $TextBoxResult.Text
    $logOutput = $logOutput.Replace("`r`n", "`r`n")
    writeLog
}

##---------------------------------------------------------------------------------------------------
#Define Function to disable project base policy
function disableProjectPolicy {
    $projectName = $ComboBoxProject.Text.ToString()
    $projectNameEnc = [uri]::EscapeDataString($projectName)
    $timeStamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")

    $TextBoxResult.Clear()

    $formTitle = "Disable project level policy in project $($projectName) Result"
    $Form.Text = $formTitle

    $TextBoxResult.Text = "$($timeStamp) Disable project level policy:"
    $TextBoxResult.AppendText("`r`n")
    $TextBoxResult.AppendText("$($timeStamp) Disabling project level policy......")
    $TextBoxResult.AppendText("`r`n")
    $TextBoxResult.Appendtext("$($timeStamp) Project: $($projectName)")
    $TextBoxResult.Appendtext("`n`n")

    #$reposUrl = "https://$ADOServerFQDN/$collection/$projectName/_apis/git/repositories?api-version=6.0"
    #$reposResponse = Invoke-RestMethod -Method Get -Uri $reposUrl -Headers $headers

    $policiesUrl = "https://$ADOServerFQDN/$collection/$projectNameEnc/_apis/policy/configurations?api-version=6.0"
    $policiesResponse = Invoke-RestMethod -Method Get -Uri $policiesUrl -Headers $headers

    foreach ($policy in $policiesResponse.value) {
        if ($null -eq $policy.settings.scope.repositoryId -And $null -eq $policy.settings.scope.refName) {
            $policiesPutUrl = "https://$ADOServerFQDN/DevOpsCollection/$projectNameEnc/_apis/policy/configurations/$($policy.id)?api-version=6.0"
            $policyJson = $policy | ConvertTo-Json -Depth 10
            $policyObject = ConvertFrom-Json -InputObject $policyJson
            $policyObject.PSObject.Methods.Remove("createdBy")
            $policyObject.PSObject.Methods.Remove("createdDate")
            $policyObject.PSObject.Methods.Remove("_links")
            $policyObject.PSObject.Methods.Remove("isDeleted")
            $policyObject.PSObject.Methods.Remove("isEnterpriseManaged")
            $policyObject.PSObject.Methods.Remove("revision")
            $policyObject.PSObject.Methods.Remove("id")
            $policyObject.PSObject.Methods.Remove("url")
            $policyObject.isEnabled = $false
            $policyObject.isBlocking = $false
            $policyObject.type.PSObject.Methods.Remove("url")
            $policyObject.type.PSObject.Methods.Remove("displayName")
            $jsonData = $policyObject | ConvertTo-Json -Depth 10
            #$policyPutResponse = Invoke-RestMethod $policiesPutUrl -Method Put -Headers $headers -ContentType "application/json" -Body $jsonData
            Invoke-RestMethod $policiesPutUrl -Method Put -Headers $headers -ContentType "application/json" -Body $jsonData
            $policyGetResponse = Invoke-RestMethod -Method Get -Uri $policiesUrl -Headers $headers
            foreach ($response in $policyGetResponse.value) {
                if ($response.id -eq $policy.id) {
                     $policyEnabled = $response.isEnabled
                }
            }
            $TextBoxResult.AppendText("`r`n")
            $TextBoxResult.Appendtext("$($timeStamp) Policy Type : $($policy.Type.displayName)")
            $TextBoxResult.AppendText("`r`n")
            $TextBoxResult.Appendtext("$($timeStamp) Policy ID : $($policy.id)")
            $TextBoxResult.AppendText("`r`n")
            $TextBoxResult.Appendtext("$($timeStamp) Policy Enabled : $($policyEnabled)")
            $TextBoxResult.AppendText("`r`n")
        }
    }

    $TextBoxResult.AppendText("`r`n")
    $TextBoxResult.AppendText("Disable project level policies of project $projectName completed")

    [System.Windows.Forms.MessageBox]::Show("Disable project level policies of project`n$projectName completed","Process Result")
    [string[]]$logOutput = $TextBoxResult.Text
    $logOutput = $logOutput.Replace("`r`n", "`r`n")
    writeLog
}

##---------------------------------------------------------------------------------------------------
#Define Function to enable project base policy
function enableProjectPolicy {
    $projectName = $ComboBoxProject.Text.ToString()
    $projectNameEnc = [uri]::EscapeDataString($projectName)
    $timeStamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")

    $TextBoxResult.Clear()

    $formTitle = "Disable project level policy in project $($projectName) Result"
    $Form.Text = $formTitle

    $TextBoxResult.Text = "$($timeStamp) Enable project level policy:"
    $TextBoxResult.AppendText("`r`n")
    $TextBoxResult.AppendText("$($timeStamp) Enabling project level policy......")
    $TextBoxResult.AppendText("`r`n")
    $TextBoxResult.Appendtext("$($timeStamp) Project: $($projectName)")
    $TextBoxResult.Appendtext("`n`n")

    #$reposUrl = "https://$ADOServerFQDN/$collection/$projectName/_apis/git/repositories?api-version=6.0"
    #$reposResponse = Invoke-RestMethod -Method Get -Uri $reposUrl -Headers $headers

    $policiesUrl = "https://$ADOServerFQDN/$collection/$projectNameEnc/_apis/policy/configurations?api-version=6.0"
    $policiesResponse = Invoke-RestMethod -Method Get -Uri $policiesUrl -Headers $headers

    foreach ($policy in $policiesResponse.value) {
        if ($null -eq $policy.settings.scope.repositoryId -And $null -eq $policy.settings.scope.refName) {
            $policiesPutUrl = "https://$ADOServerFQDN/DevOpsCollection/$projectNameEnc/_apis/policy/configurations/$($policy.id)?api-version=6.0"
            $policyJson = $policy | ConvertTo-Json -Depth 10
            $policyObject = ConvertFrom-Json -InputObject $policyJson
            $policyObject.PSObject.Methods.Remove("createdBy")
            $policyObject.PSObject.Methods.Remove("createdDate")
            $policyObject.PSObject.Methods.Remove("_links")
            $policyObject.PSObject.Methods.Remove("isDeleted")
            $policyObject.PSObject.Methods.Remove("isEnterpriseManaged")
            $policyObject.PSObject.Methods.Remove("revision")
            $policyObject.PSObject.Methods.Remove("id")
            $policyObject.PSObject.Methods.Remove("url")
            $policyObject.isEnabled = $true
            $policyObject.isBlocking = $true
            $policyObject.type.PSObject.Methods.Remove("url")
            $policyObject.type.PSObject.Methods.Remove("displayName")
            $jsonData = $policyObject | ConvertTo-Json -Depth 10
            #$policyPutResponse = Invoke-RestMethod $policiesPutUrl -Method Put -Headers $headers -ContentType "application/json" -Body $jsonData
            Invoke-RestMethod $policiesPutUrl -Method Put -Headers $headers -ContentType "application/json" -Body $jsonData
            $policyGetResponse = Invoke-RestMethod -Method Get -Uri $policiesUrl -Headers $headers
            foreach ($response in $policyGetResponse.value) {
                if ($response.id -eq $policy.id) {
                     $policyEnabled = $response.isEnabled
                }
            }
            $TextBoxResult.AppendText("`r`n")
            $TextBoxResult.Appendtext("$($timeStamp) Policy Type : $($policy.Type.displayName)")
            $TextBoxResult.AppendText("`r`n")
            $TextBoxResult.Appendtext("$($timeStamp) Policy ID : $($policy.id)")
            $TextBoxResult.AppendText("`r`n")
            $TextBoxResult.Appendtext("$($timeStamp) Policy Enabled : $($policyEnabled)")
            $TextBoxResult.AppendText("`r`n")
        }
    }

    $TextBoxResult.AppendText("`r`n")
    $TextBoxResult.AppendText("Enable project level policies of project $projectName completed")

    [System.Windows.Forms.MessageBox]::Show("Enable project level policies of project`n$projectName completed","Process Result")
    [string[]]$logOutput = $TextBoxResult.Text
    $logOutput = $logOutput.Replace("`r`n", "`r`n")
    writeLog
}

##---------------------------------------------------------------------------------------------------
#Define Function to delete project base policy
function deleteProjectPolicy {
    $projectName = $ComboBoxProject.Text.ToString()
    $projectNameEnc = [uri]::EscapeDataString($projectName)
    $timeStamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")

    $TextBoxResult.Text = "Please enter passcode to process"

    $passcodeInput = [Microsoft.VisualBasic.Interaction]::InputBox("Please enter passcode:", "User Input", "")
    if ($passcodeInput -eq "123456789") {

        $TextBoxResult.Clear()

        $formTitle = "Disable project level policy in project $($projectName) Result"
        $Form.Text = $formTitle

        $TextBoxResult.Text = "$($timeStamp) Delete project level policy:"
        $TextBoxResult.AppendText("`r`n")
        $TextBoxResult.AppendText("$($timeStamp) Delete project level policy......")

        #$reposUrl = "https://$ADOServerFQDN/$collection/$projectName/_apis/git/repositories?api-version=6.0"
        #$reposResponse = Invoke-RestMethod -Method Get -Uri $reposUrl -Headers $headers

        $policiesUrl = "https://$ADOServerFQDN/$collection/$projectNameEnc/_apis/policy/configurations?api-version=6.0"
        $policiesResponse = Invoke-RestMethod -Method Get -Uri $policiesUrl -Headers $headers

        foreach ($policy in $policiesResponse.value) {
            if ($null -eq $policy.settings.scope.repositoryId -And $null -eq $policy.settings.scope.refName) {
                $policyPutUrl = "https://$ADOServerFQDN/$collection/$projectNameEnc/_apis/policy/configurations/$($policy.id)?api-version=6.0"
                #$policyPutResponse = Invoke-RestMethod $policyPutUrl -Method Delete -Headers $headers -ContentType "application/json" -Body $jsonData
                Invoke-RestMethod $policyPutUrl -Method Delete -Headers $headers -ContentType "application/json" -Body $jsonData
            }
        }
        $TextBoxResult.AppendText("`r`n")
        $TextBoxResult.AppendText("Delete project level policies of project $projectName completed")
    } else {
        $TextBoxResult.AppendText("`r`n")
        $TextBoxResult.Text = "Passcode incorrect or user cancelled"
    }

    $TextBoxResult.AppendText("`r`n")
    $TextBoxResult.AppendText("Delete project level policies of project $projectName completed")

    [System.Windows.Forms.MessageBox]::Show("Delete project level policies of project`n$projectName completed","Process Result")
    [string[]]$logOutput = $TextBoxResult.Text
    $logOutput = $logOutput.Replace("`r`n", "`r`n")
    writeLog
}

###===================================================================================================
##Define Repository Base Policy Function
##---------------------------------------------------------------------------------------------------
#Define Function to query repository base policy
function queryReposPolicy {
    $projectName = $ComboBoxProject.Text.ToString()
    $projectNameEnc = [uri]::EscapeDataString($projectName)
    $timeStamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")

    $TextBoxResult.Clear()

    $formTitle = "Query repository policy in project $($projectName) Result"
    $Form.Text = $formTitle

    $TextBoxResult.Text = "$($timeStamp) The repository policy on each repository:"
    $TextBoxResult.AppendText("`r`n")
    $TextBoxResult.Appendtext("$($timeStamp) Project: $($projectName)")
    $TextBoxResult.Appendtext("`n`n")

    $policiesObject = [System.Collections.Generic.List[object]]::new()

    $reposUrl = "https://$ADOServerFQDN/$collection/$projectNameEnc/_apis/git/repositories?api-version=6.0"
    $reposResponse = Invoke-RestMethod -Method Get -Uri $reposUrl -Headers $headers
    $repoMap = @{}
    $targetRepos = getTargetRepos -Repos $reposResponse.value
    foreach ($repo in $targetRepos) {
        $repoMap[$repo.id] = $repo.name
    }

    $policiesUrl = "https://$ADOServerFQDN/$collection/$projectNameEnc/_apis/policy/configurations?api-version=6.0"
    $policiesResponse = Invoke-RestMethod -Method Get -Uri $policiesUrl -Headers $headers

    $hasSelection = ($null -ne $script:SelectedRepoIds -and $script:SelectedRepoIds.Count -gt 0)
    foreach ($policy in $policiesResponse.value) {
        # settings.scope can be a single object or an array; repo-level policies include repositoryId and refName
        $scopes = @()
        if ($null -ne $policy.settings -and $null -ne $policy.settings.scope) {
            $scopes = @($policy.settings.scope)
        }

        $matchedScope = $null
        foreach ($s in $scopes) {
            if ($null -ne $s.repositoryId -and $null -ne $s.refName) {
                $matchedScope = $s
                break
            }
        }

        if ($null -ne $matchedScope) {
            if ($hasSelection -and (-not ($script:SelectedRepoIds -contains $matchedScope.repositoryId))) {
                continue
            }
            $policyRepo = $repoMap[$matchedScope.repositoryId]
            if ([string]::IsNullOrWhiteSpace($policyRepo)) { $policyRepo = $matchedScope.repositoryId }

            if ($matchedScope.refName -eq "refs/heads/main") {
                $policyBranch = "Main"
            } elseif ($matchedScope.refName -eq "refs/heads/develop") {
                $policyBranch = "Develop"
            } else {
                $policyBranch = $matchedScope.refName
            }

            $policyTemp = [pscustomobject]@{
                policyRepo = $($policyRepo)
                policyBranch = $($policyBranch)
                policyType = $($policy.Type.displayName)
                policyID = $($policy.id)
                policyEnabled = $($policy.isEnabled)
            }
            if ($null -ne $policyTemp.policyEnabled) {
                $policiesObject.add($policyTemp)
            }
        }
    }

    $policiesOutput = $policiesObject | Sort-Object #-Unique -Property policyRepo #policyProject,policyID,policyType,policyEabled

    foreach ($policyOutput in $policiesOutput) {
        $TextBoxResult.AppendText("`r`n")
        $TextBoxResult.Appendtext("$($timeStamp) Reporistory: $($policyOutput.policyRepo)")
        $TextBoxResult.AppendText("`r`n")
        $TextBoxResult.Appendtext("$($timeStamp) Branch: $($policyOutput.policyBranch)")
        $TextBoxResult.AppendText("`r`n")
        $TextBoxResult.Appendtext("$($timeStamp) Policy Type : $($policyOutput.policyType)")
        $TextBoxResult.AppendText("`r`n")
        $TextBoxResult.Appendtext("$($timeStamp) Policy ID : $($policyOutput.policyID)")
        $TextBoxResult.AppendText("`r`n")
        $TextBoxResult.Appendtext("$($timeStamp) Policy Enabled : $($policyOutput.policyEnabled)")
        $TextBoxResult.AppendText("`r`n")
    }

    $TextBoxResult.AppendText("`r`n")
    $TextBoxResult.AppendText("Query repositories policy of project $projectName completed")

    [System.Windows.Forms.MessageBox]::Show("Query reporistoty policies of project`n$projectName completed","Process Result")
    [string[]]$logOutput = $TextBoxResult.Text
    $logOutput = $logOutput.Replace("`r`n", "`r`n")
    writeLog
}

##---------------------------------------------------------------------------------------------------
#Define Function to disable repository base policy
function disableReposPolicy {
    $projectName = $ComboBoxProject.Text.ToString()
    $projectNameEnc = [uri]::EscapeDataString($projectName)
    $timeStamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")

    $TextBoxResult.Clear()

    $formTitle = "Disable  repositories policy in project $($projectName) Result"
    $Form.Text = $formTitle

    $TextBoxResult.Text = "$($timeStamp) Disable repositories policy:"
    $TextBoxResult.AppendText("`r`n")
    $TextBoxResult.AppendText("$($timeStamp) Disabling repositories policy......")
    $TextBoxResult.AppendText("`r`n")
    $TextBoxResult.Appendtext("$($timeStamp) Project: $($projectName)")
    $TextBoxResult.Appendtext("`n`n")

    $reposUrl = "https://$ADOServerFQDN/$collection/$projectNameEnc/_apis/git/repositories?api-version=6.0"
    $reposResponse = Invoke-RestMethod -Method Get -Uri $reposUrl -Headers $headers

    $policiesUrl = "https://$ADOServerFQDN/$collection/$projectNameEnc/_apis/policy/configurations?api-version=6.0"
    $policiesResponse = Invoke-RestMethod -Method Get -Uri $policiesUrl -Headers $headers

    $hasSelection = ($null -ne $script:SelectedRepoIds -and $script:SelectedRepoIds.Count -gt 0)
    foreach ($policy in $policiesResponse.value) {
        if ($null -ne $policy.settings.scope.repositoryId -And $null -ne $policy.settings.scope.refName) {
            if ($hasSelection -and (-not ($script:SelectedRepoIds -contains $policy.settings.scope.repositoryId))) {
                continue
            }
            $policiesPutUrl = "https://$ADOServerFQDN/DevOpsCollection/$projectNameEnc/_apis/policy/configurations/$($policy.id)?api-version=6.0"
            $policyJson = $policy | ConvertTo-Json -Depth 10
            $policyObject = ConvertFrom-Json -InputObject $policyJson
            $policyObject.PSObject.Methods.Remove("createdBy")
            $policyObject.PSObject.Methods.Remove("createdDate")
            $policyObject.PSObject.Methods.Remove("_links")
            $policyObject.PSObject.Methods.Remove("isDeleted")
            $policyObject.PSObject.Methods.Remove("isEnterpriseManaged")
            $policyObject.PSObject.Methods.Remove("revision")
            $policyObject.PSObject.Methods.Remove("id")
            $policyObject.PSObject.Methods.Remove("url")
            $policyObject.isEnabled = $false
            $policyObject.isBlocking = $false
            $policyObject.type.PSObject.Methods.Remove("url")
            $policyObject.type.PSObject.Methods.Remove("displayName")
            $jsonData = $policyObject | ConvertTo-Json -Depth 10
            #$policyPutResponse = Invoke-RestMethod $policiesPutUrl -Method Put -Headers $headers -ContentType "application/json" -Body $jsonData
            Invoke-RestMethod $policiesPutUrl -Method Put -Headers $headers -ContentType "application/json" -Body $jsonData
            $policyGetResponse = Invoke-RestMethod -Method Get -Uri $policiesUrl -Headers $headers
            foreach ($response in $policyGetResponse.value) {
                if ($response.id -eq $policy.id) {
                     $policyEnabled = $response.isEnabled
                }
            }
            foreach ($repo in $reposResponse.value) {
                if ($repo.id -eq $policy.settings.scope.repositoryId) {
                     $policyRepo = $repo.name
                }
            }
            if ($policy.settings.scope.refName -eq "refs/heads/main") {
                $policyBranch = "Main"
            } elseif ($policy.settings.scope.refName -eq "refs/heads/develop") {
                $policyBranch = "Develop"
            }
            $TextBoxResult.AppendText("`r`n")
            $TextBoxResult.Appendtext("$($timeStamp) Reporistory: $($policyRepo)")
            $TextBoxResult.AppendText("`r`n")
            $TextBoxResult.Appendtext("$($timeStamp) Branch: $($policyBranch)")
            $TextBoxResult.AppendText("`r`n")
            $TextBoxResult.Appendtext("$($timeStamp) Policy Type : $($policy.Type.displayName)")
            $TextBoxResult.AppendText("`r`n")
            $TextBoxResult.Appendtext("$($timeStamp) Policy ID : $($policy.id)")
            $TextBoxResult.AppendText("`r`n")
            $TextBoxResult.Appendtext("$($timeStamp) Policy Enabled : $($policyEnabled)")
            $TextBoxResult.AppendText("`r`n")
        }
    }

    $TextBoxResult.AppendText("`r`n")
    $TextBoxResult.AppendText("Disable repositories policy of project $projectName completed")

    [System.Windows.Forms.MessageBox]::Show("Disable repositories policy of project`n$projectName completed","Process Result")
    [string[]]$logOutput = $TextBoxResult.Text
    $logOutput = $logOutput.Replace("`r`n", "`r`n")
    writeLog
}

##---------------------------------------------------------------------------------------------------
#Define Function to enable repository base policy
function enableReposPolicy {
    $projectName = $ComboBoxProject.Text.ToString()
    $projectNameEnc = [uri]::EscapeDataString($projectName)
    $timeStamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")

    $TextBoxResult.Clear()

    $formTitle = "Enable repositories policy in project $($projectName) Result"
    $Form.Text = $formTitle

    $TextBoxResult.Text = "$($timeStamp) Enable repositories policy:"
    $TextBoxResult.AppendText("`r`n")
    $TextBoxResult.AppendText("$($timeStamp) Enabling repositories policy......")
    $TextBoxResult.AppendText("`r`n")
    $TextBoxResult.Appendtext("$($timeStamp) Project: $($projectName)")
    $TextBoxResult.Appendtext("`n`n")

    $reposUrl = "https://$ADOServerFQDN/$collection/$projectNameEnc/_apis/git/repositories?api-version=6.0"
    $reposResponse = Invoke-RestMethod -Method Get -Uri $reposUrl -Headers $headers

    $policiesUrl = "https://$ADOServerFQDN/$collection/$projectNameEnc/_apis/policy/configurations?api-version=6.0"
    $policiesResponse = Invoke-RestMethod -Method Get -Uri $policiesUrl -Headers $headers

    $hasSelection = ($null -ne $script:SelectedRepoIds -and $script:SelectedRepoIds.Count -gt 0)
    foreach ($policy in $policiesResponse.value) {
        if ($null -ne $policy.settings.scope.repositoryId -And $null -ne $policy.settings.scope.refName) {
            if ($hasSelection -and (-not ($script:SelectedRepoIds -contains $policy.settings.scope.repositoryId))) {
                continue
            }
                    $policiesPutUrl = "https://$ADOServerFQDN/DevOpsCollection/$projectNameEnc/_apis/policy/configurations/$($policy.id)?api-version=6.0"
            $policyJson = $policy | ConvertTo-Json -Depth 10
            $policyObject = ConvertFrom-Json -InputObject $policyJson
            $policyObject.PSObject.Methods.Remove("createdBy")
            $policyObject.PSObject.Methods.Remove("createdDate")
            $policyObject.PSObject.Methods.Remove("_links")
            $policyObject.PSObject.Methods.Remove("isDeleted")
            $policyObject.PSObject.Methods.Remove("isEnterpriseManaged")
            $policyObject.PSObject.Methods.Remove("revision")
            $policyObject.PSObject.Methods.Remove("id")
            $policyObject.PSObject.Methods.Remove("url")
            $policyObject.isEnabled = $true
            $policyObject.isBlocking = $true
            $policyObject.type.PSObject.Methods.Remove("url")
            $policyObject.type.PSObject.Methods.Remove("displayName")
            $jsonData = $policyObject | ConvertTo-Json -Depth 10
            #$policyPutResponse = Invoke-RestMethod $policiesPutUrl -Method Put -Headers $headers -ContentType "application/json" -Body $jsonData
            Invoke-RestMethod $policiesPutUrl -Method Put -Headers $headers -ContentType "application/json" -Body $jsonData
            $policyGetResponse = Invoke-RestMethod -Method Get -Uri $policiesUrl -Headers $headers
            foreach ($response in $policyGetResponse.value) {
                if ($response.id -eq $policy.id) {
                     $policyEnabled = $response.isEnabled
                }
            }
            foreach ($repo in $reposResponse.value) {
                if ($repo.id -eq $policy.settings.scope.repositoryId) {
                     $policyRepo = $repo.name
                }
            }
            if ($policy.settings.scope.refName -eq "refs/heads/main") {
                $policyBranch = "Main"
            } elseif ($policy.settings.scope.refName -eq "refs/heads/develop") {
                $policyBranch = "Develop"
            }
            $TextBoxResult.AppendText("`r`n")
            $TextBoxResult.Appendtext("$($timeStamp) Reporistory: $($policyRepo)")
            $TextBoxResult.AppendText("`r`n")
            $TextBoxResult.Appendtext("$($timeStamp) Branch: $($policyBranch)")
            $TextBoxResult.AppendText("`r`n")
            $TextBoxResult.Appendtext("$($timeStamp) Policy Type : $($policy.Type.displayName)")
            $TextBoxResult.AppendText("`r`n")
            $TextBoxResult.Appendtext("$($timeStamp) Policy ID : $($policy.id)")
            $TextBoxResult.AppendText("`r`n")
            $TextBoxResult.Appendtext("$($timeStamp) Policy Enabled : $($policyEnabled)")
            $TextBoxResult.AppendText("`r`n")
        }
    }

    $TextBoxResult.AppendText("`r`n")
    $TextBoxResult.AppendText("Enable repositories policy of project $projectName completed")

    [System.Windows.Forms.MessageBox]::Show("Enable repositories policy of project`n$projectName completed","Process Result")
    [string[]]$logOutput = $TextBoxResult.Text
    $logOutput = $logOutput.Replace("`r`n", "`r`n")
    writeLog
}

##---------------------------------------------------------------------------------------------------
#Define Function to delete repository base policy
function deleteReposPolicy {
    $projectName = $ComboBoxProject.Text.ToString()
    $projectNameEnc = [uri]::EscapeDataString($projectName)
    $timeStamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")

    $TextBoxResult.Text = "Please enter passcode to process"

    $passcodeInput = [Microsoft.VisualBasic.Interaction]::InputBox("Please enter passcode:", "User Input", "")
    if ($passcodeInput -eq "123456789") {

        $TextBoxResult.Clear()

        $formTitle = "Disable project level policy in project $($projectName) Result"
        $Form.Text = $formTitle

        $TextBoxResult.Text = "$($timeStamp) Delete project level policy:"
        $TextBoxResult.AppendText("`r`n")
        $TextBoxResult.AppendText("$($timeStamp) Delete project level policy......")

        #$reposUrl = "https://$ADOServerFQDN/$collection/$projectName/_apis/git/repositories?api-version=6.0"
        #$reposResponse = Invoke-RestMethod -Method Get -Uri $reposUrl -Headers $headers

        $policiesUrl = "https://$ADOServerFQDN/$collection/$projectNameEnc/_apis/policy/configurations?api-version=6.0"
        $policiesResponse = Invoke-RestMethod -Method Get -Uri $policiesUrl -Headers $headers

        $hasSelection = ($null -ne $script:SelectedRepoIds -and $script:SelectedRepoIds.Count -gt 0)
        foreach ($policy in $policiesResponse.value) {
            if ($null -ne $policy.settings.scope.repositoryId -And $null -ne $policy.settings.scope.refName) {
                if ($hasSelection -and (-not ($script:SelectedRepoIds -contains $policy.settings.scope.repositoryId))) {
                    continue
                }
                $policyPutUrl = "https://$ADOServerFQDN/$collection/$projectNameEnc/_apis/policy/configurations/$($policy.id)?api-version=6.0"
                #$policyPutResponse = Invoke-RestMethod $policyPutUrl -Method Delete -Headers $headers -ContentType "application/json" -Body $jsonData
                Invoke-RestMethod $policyPutUrl -Method Delete -Headers $headers -ContentType "application/json" -Body $jsonData
            }
        }
        $TextBoxResult.AppendText("`r`n")
        $TextBoxResult.AppendText("Delete project level policies of project $projectName completed")
    } else {
        $TextBoxResult.AppendText("`r`n")
        $TextBoxResult.Text = "Passcode incorrect or user cancelled"
    }

    [System.Windows.Forms.MessageBox]::Show("Delete project level policies of project`n$projectName completed","Process Result")
    [string[]]$logOutput = $TextBoxResult.Text
    $logOutput = $logOutput.Replace("`r`n", "`r`n")
    writeLog
}

###===================================================================================================
##Define Function of Self Approval
##---------------------------------------------------------------------------------------------------
#Define Function to Query self approval
function querySelfApproval {
    $projectName = $ComboBoxProject.Text.ToString()
    $projectNameEnc = [uri]::EscapeDataString($projectName)
    $timeStamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")

    $TextBoxResult.Clear()

    $formTitle = "Query repositories self approval requirement in project $($projectName) Result"
    $Form.Text = $formTitle

    $TextBoxResult.Text = "$($timeStamp) Query repositories self approval requirement:"
    $TextBoxResult.AppendText("`r`n")
    $TextBoxResult.AppendText("$($timeStamp) Querying repositories self approval requirement......")
    $TextBoxResult.AppendText("`r`n")
    $TextBoxResult.Appendtext("$($timeStamp) Project: $($projectName)")
    $TextBoxResult.Appendtext("`n`n")

    $reposUrl = "https://$ADOServerFQDN/$collection/$projectNameEnc/_apis/git/repositories?api-version=6.0"
    $reposResponse = Invoke-RestMethod -Method Get -Uri $reposUrl -Headers $headers

    $policiesUrl = "https://$ADOServerFQDN/$collection/$projectNameEnc/_apis/policy/configurations?api-version=6.0"
    $policiesResponse = Invoke-RestMethod -Method Get -Uri $policiesUrl -Headers $headers

    $hasSelection = ($null -ne $script:SelectedRepoIds -and $script:SelectedRepoIds.Count -gt 0)
    foreach ($policy in $policiesResponse.value) {
        if ($policy.Type.id -eq "fa4e907d-c16b-4a4c-9dfa-4906e5d171dd") {
            if ($hasSelection -and (-not ($script:SelectedRepoIds -contains $policy.settings.scope.repositoryId))) {
                continue
            }
            $policyGetResponse = Invoke-RestMethod -Method Get -Uri $policiesUrl -Headers $headers
            foreach ($repo in $reposResponse.value) {
                if ($repo.id -eq $policy.settings.scope.repositoryId) {
                     $policyRepo = $repo.name
                }
            }
            foreach ($response in $policyGetResponse.value) {
                if ($response.id -eq $policy.id) {
                     $policyEnabled = $response.isEnabled
                     $policyApproval = $response.settings.creatorVoteCounts
                     if ($policyApproval -eq $true) {
                         $policyApprovalState = "Allow"
                     } else {
                         $policyApprovalState = "Disallow"
                     }
                }
            }
            if ($policy.settings.scope.refName -eq "refs/heads/main") {
                $policyBranch = "Main"
            } elseif ($policy.settings.scope.refName -eq "refs/heads/develop") {
                $policyBranch = "Develop"
            }
            $TextBoxResult.AppendText("`r`n")
            $TextBoxResult.Appendtext("$($timeStamp) Reporistory: $($policyRepo)")
            $TextBoxResult.AppendText("`r`n")
            $TextBoxResult.Appendtext("$($timeStamp) Branch: $($policyBranch)")
            $TextBoxResult.AppendText("`r`n")
            $TextBoxResult.Appendtext("$($timeStamp) Policy Type : $($policy.Type.displayName)")
            $TextBoxResult.AppendText("`r`n")
            $TextBoxResult.Appendtext("$($timeStamp) Policy ID : $($policy.id)")
            $TextBoxResult.AppendText("`r`n")
            $TextBoxResult.Appendtext("$($timeStamp) Policy Enabled : $($policyEnabled)")
            $TextBoxResult.AppendText("`r`n")
            $TextBoxResult.Appendtext("$($timeStamp) Policy Self Approval : $($policyApprovalState)")
            $TextBoxResult.AppendText("`r`n")
        }
    }

    $TextBoxResult.AppendText("`r`n")
    $TextBoxResult.AppendText("Query repositories self approval of project $projectName completed")

    [System.Windows.Forms.MessageBox]::Show("Query repositories self approval of project`n$projectName completed","Process Result")
    [string[]]$logOutput = $TextBoxResult.Text
    $logOutput = $logOutput.Replace("`r`n", "`r`n")
    writeLog
}

#$projectName = "data-platform"
#$repoName = "patron-profile-service-net-api", "patron-open-rating-net-consumer", "patron-summary-service-net-api", "patron-mpb-balance-net-consumer", "patron-personal-details-net-consumer", "patron-card-info-net-consumer", "patron-wallet-service-net-api", "patron-vvip-info-net-consumer", "patron-exclusion-net-consumer", "patron-ewallet-balance-net-consumer", "patron-coded-info-net-consumer", "patron-room-booking-net-consumer", "patron-preference-net-consumer", "patron-wechat-binding-net-consumer", "patron-srapp-status-net-consumer", "patron-uploaded-value-net-consumer", "patron-trip-service-net-api", "remove-open-ratings-job"

##---------------------------------------------------------------------------------------------------
#Define Function to Enable self approval
function enableSelfApproval {
    $projectName = $ComboBoxProject.Text.ToString()
    $projectNameEnc = [uri]::EscapeDataString($projectName)
    $timeStamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")

    $TextBoxResult.Clear()

    $formTitle = "Enable repositories self approval requirement in project $($projectName) Result"
    $Form.Text = $formTitle

    $TextBoxResult.Text = "$($timeStamp) Enable repositories self approval requirement:"
    $TextBoxResult.AppendText("`r`n")
    $TextBoxResult.AppendText("$($timeStamp) Enabling repositories self approval requirement......")
    $TextBoxResult.AppendText("`r`n")
    $TextBoxResult.Appendtext("$($timeStamp) Project: $($projectName)")
    $TextBoxResult.Appendtext("`n`n")

    $reposUrl = "https://$ADOServerFQDN/$collection/$projectNameEnc/_apis/git/repositories?api-version=6.0"
    $reposResponse = Invoke-RestMethod -Method Get -Uri $reposUrl -Headers $headers

    $policiesUrl = "https://$ADOServerFQDN/$collection/$projectNameEnc/_apis/policy/configurations?api-version=6.0"
    $policiesResponse = Invoke-RestMethod -Method Get -Uri $policiesUrl -Headers $headers

    $hasSelection = ($null -ne $script:SelectedRepoIds -and $script:SelectedRepoIds.Count -gt 0)
    foreach ($policy in $policiesResponse.value) {
        if ($policy.Type.id -eq "fa4e907d-c16b-4a4c-9dfa-4906e5d171dd") {
            if ($hasSelection -and (-not ($script:SelectedRepoIds -contains $policy.settings.scope.repositoryId))) {
                continue
            }
            $policiesPutUrl = "https://$ADOServerFQDN/DevOpsCollection/$projectNameEnc/_apis/policy/configurations/$($policy.id)?api-version=6.0"
            $policyJson = $policy | ConvertTo-Json -Depth 10
            $policyObject = ConvertFrom-Json -InputObject $policyJson
            $policyObject.PSObject.Methods.Remove("createdBy")
            $policyObject.PSObject.Methods.Remove("createdDate")
            $policyObject.PSObject.Methods.Remove("_links")
            $policyObject.PSObject.Methods.Remove("isDeleted")
            $policyObject.PSObject.Methods.Remove("isEnterpriseManaged")
            $policyObject.PSObject.Methods.Remove("revision")
            $policyObject.PSObject.Methods.Remove("id")
            $policyObject.PSObject.Methods.Remove("url")
            $policyObject.settings.creatorVoteCounts = $true
            $policyObject.type.PSObject.Methods.Remove("url")
            $policyObject.type.PSObject.Methods.Remove("displayName")
            $jsonData = $policyObject | ConvertTo-Json -Depth 10
            #$policyPutResponse = Invoke-RestMethod $policiesPutUrl -Method Put -Headers $headers -ContentType "application/json" -Body $jsonData
            Invoke-RestMethod $policiesPutUrl -Method Put -Headers $headers -ContentType "application/json" -Body $jsonData
            $policyGetResponse = Invoke-RestMethod -Method Get -Uri $policiesUrl -Headers $headers
            foreach ($repo in $reposResponse.value) {
                if ($repo.id -eq $policy.settings.scope.repositoryId) {
                     $policyRepo = $repo.name
                }
            }
            foreach ($response in $policyGetResponse.value) {
                if ($response.id -eq $policy.id) {
                     $policyEnabled = $response.isEnabled
                     $policyApproval = $response.settings.creatorVoteCounts
                     if ($policyApproval -eq $true) {
                         $policyApprovalState = "Allow"
                     } else {
                         $policyApprovalState = "Disallow"
                     }
                }
            }
            if ($policy.settings.scope.refName -eq "refs/heads/main") {
                $policyBranch = "Main"
            } elseif ($policy.settings.scope.refName -eq "refs/heads/develop") {
                $policyBranch = "Develop"
            }
            $TextBoxResult.AppendText("`r`n")
            $TextBoxResult.Appendtext("$($timeStamp) Reporistory: $($policyRepo)")
            $TextBoxResult.AppendText("`r`n")
            $TextBoxResult.Appendtext("$($timeStamp) Branch: $($policyBranch)")
            $TextBoxResult.AppendText("`r`n")
            $TextBoxResult.Appendtext("$($timeStamp) Policy Type : $($policy.Type.displayName)")
            $TextBoxResult.AppendText("`r`n")
            $TextBoxResult.Appendtext("$($timeStamp) Policy ID : $($policy.id)")
            $TextBoxResult.AppendText("`r`n")
            $TextBoxResult.Appendtext("$($timeStamp) Policy Enabled : $($policyEnabled)")
            $TextBoxResult.AppendText("`r`n")
            $TextBoxResult.Appendtext("$($timeStamp) Policy Self Approval : $($policyApprovalState)")
            $TextBoxResult.AppendText("`r`n")
        }
    }

    $TextBoxResult.AppendText("`r`n")
    $TextBoxResult.AppendText("Enable repositories self approval of project $projectName completed")

    [System.Windows.Forms.MessageBox]::Show("Enable repositories self approval of project`n$projectName completed","Process Result")
    [string[]]$logOutput = $TextBoxResult.Text
    $logOutput = $logOutput.Replace("`r`n", "`r`n")
    writeLog
}

##---------------------------------------------------------------------------------------------------
#Define Function to disable self approval
function disableSelfApproval {
    $projectName = $ComboBoxProject.Text.ToString()
    $projectNameEnc = [uri]::EscapeDataString($projectName)
    $timeStamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")

    $TextBoxResult.Clear()

    $formTitle = "Disable repositories self approval requirement in project $($projectName) Result"
    $Form.Text = $formTitle

    $TextBoxResult.Text = "$($timeStamp) Disable repositories self approval requirement:"
    $TextBoxResult.AppendText("`r`n")
    $TextBoxResult.AppendText("$($timeStamp) Disabling repositories self approval requirement......")
    $TextBoxResult.AppendText("`r`n")
    $TextBoxResult.Appendtext("$($timeStamp) Project: $($projectName)")
    $TextBoxResult.Appendtext("`n`n")

    $reposUrl = "https://$ADOServerFQDN/$collection/$projectNameEnc/_apis/git/repositories?api-version=6.0"
    $reposResponse = Invoke-RestMethod -Method Get -Uri $reposUrl -Headers $headers

    $policiesUrl = "https://$ADOServerFQDN/$collection/$projectNameEnc/_apis/policy/configurations?api-version=6.0"
    $policiesResponse = Invoke-RestMethod -Method Get -Uri $policiesUrl -Headers $headers

    $hasSelection = ($null -ne $script:SelectedRepoIds -and $script:SelectedRepoIds.Count -gt 0)
    foreach ($policy in $policiesResponse.value) {
       if ($policy.Type.id -eq "fa4e907d-c16b-4a4c-9dfa-4906e5d171dd") {
            if ($hasSelection -and (-not ($script:SelectedRepoIds -contains $policy.settings.scope.repositoryId))) {
                continue
            }
            $policiesPutUrl = "https://$ADOServerFQDN/DevOpsCollection/$projectNameEnc/_apis/policy/configurations/$($policy.id)?api-version=6.0"
            $policyJson = $policy | ConvertTo-Json -Depth 10
            $policyObject = ConvertFrom-Json -InputObject $policyJson
            $policyObject.PSObject.Methods.Remove("createdBy")
            $policyObject.PSObject.Methods.Remove("createdDate")
            $policyObject.PSObject.Methods.Remove("_links")
            $policyObject.PSObject.Methods.Remove("isDeleted")
            $policyObject.PSObject.Methods.Remove("isEnterpriseManaged")
            $policyObject.PSObject.Methods.Remove("revision")
            $policyObject.PSObject.Methods.Remove("id")
            $policyObject.PSObject.Methods.Remove("url")
            $policyObject.settings.creatorVoteCounts = $false
            $policyObject.type.PSObject.Methods.Remove("url")
            $policyObject.type.PSObject.Methods.Remove("displayName")
            $jsonData = $policyObject | ConvertTo-Json -Depth 10
            #$policyPutResponse = Invoke-RestMethod $policiesPutUrl -Method Put -Headers $headers -ContentType "application/json" -Body $jsonData
            Invoke-RestMethod $policiesPutUrl -Method Put -Headers $headers -ContentType "application/json" -Body $jsonData
            $policyGetResponse = Invoke-RestMethod -Method Get -Uri $policiesUrl -Headers $headers
            foreach ($repo in $reposResponse.value) {
                if ($repo.id -eq $policy.settings.scope.repositoryId) {
                     $policyRepo = $repo.name
                }
            }
            foreach ($response in $policyGetResponse.value) {
                if ($response.id -eq $policy.id) {
                     $policyEnabled = $response.isEnabled
                     $policyApproval = $response.settings.creatorVoteCounts
                     if ($policyApproval -eq $true) {
                         $policyApprovalState = "Allow"
                     } else {
                         $policyApprovalState = "Disallow"
                     }
                }
            }
            if ($policy.settings.scope.refName -eq "refs/heads/main") {
                $policyBranch = "Main"
            } elseif ($policy.settings.scope.refName -eq "refs/heads/develop") {
                $policyBranch = "Develop"
            }
            $TextBoxResult.AppendText("`r`n")
            $TextBoxResult.Appendtext("$($timeStamp) Reporistory: $($policyRepo)")
            $TextBoxResult.AppendText("`r`n")
            $TextBoxResult.Appendtext("$($timeStamp) Branch: $($policyBranch)")
            $TextBoxResult.AppendText("`r`n")
            $TextBoxResult.Appendtext("$($timeStamp) Policy Type : $($policy.Type.displayName)")
            $TextBoxResult.AppendText("`r`n")
            $TextBoxResult.Appendtext("$($timeStamp) Policy ID : $($policy.id)")
            $TextBoxResult.AppendText("`r`n")
            $TextBoxResult.Appendtext("$($timeStamp) Policy Enabled : $($policyEnabled)")
            $TextBoxResult.AppendText("`r`n")
            $TextBoxResult.Appendtext("$($timeStamp) Policy Self Approval : $($policyApprovalState)")
            $TextBoxResult.AppendText("`r`n")
        }
    }

    $TextBoxResult.AppendText("`r`n")
    $TextBoxResult.AppendText("Disable repositories self approval of project $projectName completed")

    [System.Windows.Forms.MessageBox]::Show("Disable repositories self approval of project`n$projectName completed","Process Result")
    [string[]]$logOutput = $TextBoxResult.Text
    $logOutput = $logOutput.Replace("`r`n", "`r`n")
    writeLog
}

###===================================================================================================
##Button Enable Disable Checking
##---------------------------------------------------------------------------------------------------
#Enable/Disbaled Button until project name is inputed
$ComboBoxProject.Add_SelectedIndexChanged({
    if ($ComboBoxProject.SelectedIndex -ne -1) {
        $ButtonSelectRepos.Enabled = $true
        $ButtonQueryRepos.Enabled = $true
        $ButtonQueryProjectPolicy.Enabled = $true
        $ButtonDisableProjectPolicy.Enabled = $true
        $ButtonEnableProjectPolicy.Enabled = $true
        $ButtonDeleteProjectPolicy.Enabled = $true
        $ButtonQueryReposPolicy.Enabled = $true
        $ButtonDisableReposPolicy.Enabled = $true
        $ButtonEnableReposPolicy.Enabled = $true
        $ButtonDeleteReposPolicy.Enabled = $true
        $ButtonQuerySelfApproval.Enabled = $true
        $ButtonEnableSelfApproval.Enabled = $true
        $ButtonDisableSelfApproval.Enabled = $true
        $ObjBoxBranch.Enabled = $true
        $LabelDisplayProject.Text = $ComboBoxProject.Text
        $LabelDisplayProject.BackColor = "Black"
        $LabelDisplayProject.ForeColor = "White"
        loadReposForProject
        if ($ObjBoxBranch.Selectedindex -ne -1) {
            $ButtonSetDefaultBranch.Enabled = $true
        } else {
            $ButtonSetDefaultBranch.Enabled = $false
        }
    } else {
        $ButtonSelectRepos.Enabled = $false
        $ButtonSetDefaultBranch.Enabled = $false
        $ButtonQueryRepos.Enabled = $false
        $ButtonQueryProjectPolicy.Enabled = $false
        $ButtonDisableProjectPolicy.Enabled = $false
        $ButtonEnableProjectPolicy.Enabled = $false
        $ButtonDeleteProjectPolicy.Enabled = $false
        $ButtonQueryReposPolicy.Enabled = $false
        $ButtonDisableReposPolicy.Enabled = $false
        $ButtonEnableReposPolicy.Enabled = $false
        $ButtonDeleteReposPolicy.Enabled = $false
        $ButtonQuerySelfApproval.Enabled = $true
        $ButtonEnableSelfApproval.Enabled = $false
        $ButtonDisableSelfApproval.Enabled = $false
        $ObjBoxBranch.Enabled = $false
        $LabelDisplayProject.Text = ""
        $LabelDisplayProject.BackColor = ""
        $LabelDisplayProject.ForeColor = ""
        $script:SelectedRepoIds = @()
        $script:SelectedRepoNames = @()
        $script:RepoListCache = @()
        updateSelectedReposDisplay
    }
})

##---------------------------------------------------------------------------------------------------
#Enable/Disbaled Button until default branch is selected
$ObjBoxBranch.Add_SelectedIndexChanged({
    if ($ObjBoxBranch.Selectedindex -ne -1) {
        $ButtonSetDefaultBranch.Enabled = $true
#        if ($ComboBoxProject.TextLength -gt 0) {
#            $ButtonQueryRepos.Enabled = $true
#            $ButtonQueryProjectPolicy.Enabled = $true
#            $ButtonDisableProjectPolicy.Enabled = $true
#            $ButtonEnableProjectPolicy.Enabled = $true
#            $ButtonDeleteProjectPolicy.Enabled = $true
#            $ButtonQueryReposPolicy.Enabled = $true
#            $ButtonDisableReposPolicy.Enabled = $true
#            $ButtonEnableReposPolicy.Enabled = $true
#            $ButtonDeleteReposPolicy.Enabled = $true
#            $ButtonEnableSelfApproval.Enabled = $true
#            $ButtonDisableSelfApproval.Enabled = $true
#        } else {
#            $ButtonSetDefaultBranch.Enabled = $false
#            $ButtonQueryRepos.Enabled = $false
#            $ButtonQueryProjectPolicy.Enabled = $false
#            $ButtonDisableProjectPolicy.Enabled = $false
#            $ButtonEnableProjectPolicy.Enabled = $false
#            $ButtonDeleteProjectPolicy.Enabled = $false
#            $ButtonQueryReposPolicy.Enabled = $false
#            $ButtonDisableReposPolicy.Enabled = $false
#            $ButtonEnableReposPolicy.Enabled = $false
#            $ButtonDeleteReposPolicy.Enabled = $false
#            $ButtonEnableSelfApproval.Enabled = $false
#            $ButtonDisableSelfApproval.Enabled = $false
        } else {
        $ButtonSetDefaultBranch.Enabled = $false
    }
})

###===================================================================================================
##Loading Windows Form Object
##---------------------------------------------------------------------------------------------------
# Loading Windows Form
# Apply explicit top-area layout overrides to prevent overlaps
# These overrides ensure the project controls and labels occupy the reserved
# area at the top of the form so the later reflow doesn't place controls
# on top of them.
$LabelTextBoxProject.Location = New-Object System.Drawing.Point(20,20)
$LabelTextBoxProject.Size = New-Object System.Drawing.Size(100,20)

if ($null -ne $ComboBoxProject) {
    $ComboBoxProject.Location = New-Object System.Drawing.Point(130,20)
    $ComboBoxProject.Size = New-Object System.Drawing.Size(300,24)
}

if ($null -ne $ButtonRefreshProjects) {
    $ButtonRefreshProjects.Location = New-Object System.Drawing.Point(340,20)
    $ButtonRefreshProjects.Size = New-Object System.Drawing.Size(120,30)
    $ButtonRefreshProjects.BringToFront()
}

$LabelUsingProject.Location = New-Object System.Drawing.Point(20,60)
$LabelUsingProject.Size = New-Object System.Drawing.Size(100,20)

$LabelDisplayProject.Location = New-Object System.Drawing.Point(130,60)
$LabelDisplayProject.Size = New-Object System.Drawing.Size(200,20)

$LabelObjBoxBranch.Location = New-Object System.Drawing.Point(20,100)
$LabelObjBoxBranch.Size = New-Object System.Drawing.Size(100,20)

if ($null -ne $ObjBoxBranch) {
    $ObjBoxBranch.Location = New-Object System.Drawing.Point(130,100)
    $ObjBoxBranch.Size = New-Object System.Drawing.Size(200,20)
}

if ($null -ne $ButtonQueryRepos) {
    $ButtonQueryRepos.Location = New-Object System.Drawing.Point(20,140)
    $ButtonQueryRepos.Size = New-Object System.Drawing.Size(200,40)
}

if ($null -ne $ButtonSetDefaultBranch) {
    $ButtonSetDefaultBranch.Location = New-Object System.Drawing.Point(240,140)
    $ButtonSetDefaultBranch.Size = New-Object System.Drawing.Size(200,40)
}

# Revert form size to a reasonable default so controls fit predictably
if ($null -ne $Form) {
    $Form.Font = New-Object System.Drawing.Font("Segoe UI",9)
    $Form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
    $Form.ClientSize = New-Object System.Drawing.Size(1280,900)

    # Create a TableLayoutPanel for the top controls for predictable column alignment
    $TopTable = New-Object System.Windows.Forms.TableLayoutPanel
    $TopTable.ColumnCount = 8
    $TopTable.RowCount = 3
    # Define column widths (percent)
    # Configure column styles: reserve an absolute wide column for the Project combo so its MinimumSize is respected
    $TopTable.ColumnStyles.Clear()
    $TopTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 5)))
    $TopTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 360)))
    $TopTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 130)))
    $TopTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
    $TopTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 220)))
    # Reserve three absolute columns for the self-approval buttons to avoid overlap
    $TopTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 180)))
    $TopTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 180)))
    $TopTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 180)))
    $TopTable.RowStyles.Clear()
    $TopTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 48)))
    $TopTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 48)))
    $TopTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 48)))
    $TopTable.Location = New-Object System.Drawing.Point(10,10)
    # Make table full width of the form (with small padding)
    try { $tableWidth = $Form.ClientSize.Width - 40 } catch { $tableWidth = 1200 }
    $TopTable.Size = New-Object System.Drawing.Size([Math]::Max(700,$tableWidth),168)
    $TopTable.Dock = 'Top'
    $TopTable.Padding = New-Object System.Windows.Forms.Padding(6)

    # Remove controls from form before reparenting
    $Form.Controls.Remove($LabelTextBoxProject) | Out-Null
    $Form.Controls.Remove($ComboBoxProject) | Out-Null
    $Form.Controls.Remove($ButtonRefreshProjects) | Out-Null
    $Form.Controls.Remove($LabelUsingProject) | Out-Null
    $Form.Controls.Remove($LabelDisplayProject) | Out-Null
    $Form.Controls.Remove($LabelObjBoxBranch) | Out-Null
    $Form.Controls.Remove($ObjBoxBranch) | Out-Null
    $Form.Controls.Remove($LabelObjBoxPolicyType) | Out-Null
    $Form.Controls.Remove($ObjBoxPolicyType) | Out-Null
    $Form.Controls.Remove($LabelSelfApproval) | Out-Null
    $Form.Controls.Remove($ButtonQuerySelfApproval) | Out-Null
    $Form.Controls.Remove($ButtonEnableSelfApproval) | Out-Null
    $Form.Controls.Remove($ButtonDisableSelfApproval) | Out-Null
    $Form.Controls.Remove($LabelRepos) | Out-Null
    $Form.Controls.Remove($LabelSelectedRepos) | Out-Null
    $Form.Controls.Remove($ButtonSelectRepos) | Out-Null
    $Form.Controls.Remove($TextBoxSelectedReposSelf) | Out-Null

    # Add controls into table cells (col,row) — ensure each control is removed from any previous parent first
    $tableAdds = @(
        @{ctrl=$LabelTextBoxProject; col=0; row=0},
        @{ctrl=$ComboBoxProject; col=1; row=0},
        @{ctrl=$ButtonRefreshProjects; col=2; row=0},
        @{ctrl=$LabelUsingProject; col=3; row=0},
        @{ctrl=$LabelDisplayProject; col=4; row=0},
        @{ctrl=$LabelObjBoxBranch; col=0; row=1},
        @{ctrl=$ObjBoxBranch; col=1; row=1},
        @{ctrl=$LabelObjBoxPolicyType; col=2; row=1},
        @{ctrl=$ObjBoxPolicyType; col=3; row=1},
        # Move LabelSelfApproval to the top row above the QuerySelfApproval button (col 5, row 0)
        @{ctrl=$LabelSelfApproval; col=5; row=0},
        @{ctrl=$ButtonQuerySelfApproval; col=5; row=1},
        @{ctrl=$ButtonEnableSelfApproval; col=6; row=1},
        @{ctrl=$ButtonDisableSelfApproval; col=7; row=1},
        @{ctrl=$LabelRepos; col=0; row=2},
        @{ctrl=$LabelSelectedRepos; col=1; row=2},
        @{ctrl=$ButtonSelectRepos; col=2; row=2},
        @{ctrl=$TextBoxSelectedReposSelf; col=5; row=2}
    )
    foreach ($entry in $tableAdds) {
        $c = $entry.ctrl
        if ($null -ne $c) {
            try { if ($null -ne $c.Parent) { $c.Parent.Controls.Remove($c) } } catch { }
            try { $TopTable.Controls.Add($c, $entry.col, $entry.row) | Out-Null } catch { }
        }
    }
    # Ensure the ComboBox span is set after adding
    try { $TopTable.SetColumnSpan($ComboBoxProject,1) } catch { }
    try { $TopTable.SetColumnSpan($TextBoxSelectedReposSelf,3) } catch { }
    # Adjust LabelSelfApproval to sit above the QuerySelfApproval button and center its text
    try {
        if ($null -ne $LabelSelfApproval) {
            $LabelSelfApproval.AutoSize = $false
            $LabelSelfApproval.Dock = 'Fill'
            $LabelSelfApproval.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
            $LabelSelfApproval.MinimumSize = New-Object System.Drawing.Size(160,24)
        }
    } catch { }

    # Set common margins and sizes for table children and ensure combos fill their cells
    foreach ($ctrl in $TopTable.Controls) {
        try { $ctrl.Margin = New-Object System.Windows.Forms.Padding(6) } catch { }
        if ($ctrl -is [System.Windows.Forms.Button]) {
            $ctrl.Size = New-Object System.Drawing.Size(160,36)
            try { $ctrl.MinimumSize = New-Object System.Drawing.Size(140,36) } catch { }
        }
        if ($ctrl -is [System.Windows.Forms.ComboBox]) {
            try { $ctrl.Dock = 'Fill' } catch { $ctrl.Size = New-Object System.Drawing.Size(300,24) }
            try { $ctrl.DropDownStyle = 'DropDownList' } catch { }
        }
        if ($ctrl -is [System.Windows.Forms.Label]) {
            $ctrl.AutoSize = $true
            try { $ctrl.Anchor = [System.Windows.Forms.AnchorStyles]::Left } catch { }
        }
    }

    # Ensure specific controls have adequate minimum widths for display
    try { $ComboBoxProject.MinimumSize = New-Object System.Drawing.Size(300,24) } catch { }
    try { $ObjBoxBranch.MinimumSize = New-Object System.Drawing.Size(220,24) } catch { }
    try { $ObjBoxPolicyType.MinimumSize = New-Object System.Drawing.Size(200,24) } catch { }
    try { $LabelDisplayProject.MinimumSize = New-Object System.Drawing.Size(200,20) } catch { }
    try {
        $TextBoxSelectedReposSelf.Dock = 'Fill'
        $TextBoxSelectedReposSelf.MinimumSize = New-Object System.Drawing.Size(240,22)
        $TextBoxSelectedReposSelf.BackColor = [System.Drawing.SystemColors]::Window
    } catch { }
    try {
        $LabelSelectedRepos.AutoSize = $false
        $LabelSelectedRepos.Dock = 'Fill'
        $LabelSelectedRepos.MinimumSize = New-Object System.Drawing.Size(320,20)
    } catch { }

    # Ensure the self-approval buttons occupy their table cells and are visible
    try { $ButtonQuerySelfApproval.Dock = 'Fill'; $ButtonQuerySelfApproval.MinimumSize = New-Object System.Drawing.Size(140,36); $ButtonQuerySelfApproval.Visible = $true } catch { }
    try { $ButtonEnableSelfApproval.Dock = 'Fill'; $ButtonEnableSelfApproval.MinimumSize = New-Object System.Drawing.Size(140,36); $ButtonEnableSelfApproval.Visible = $true } catch { }
    try { $ButtonDisableSelfApproval.Dock = 'Fill'; $ButtonDisableSelfApproval.MinimumSize = New-Object System.Drawing.Size(140,36); $ButtonDisableSelfApproval.Visible = $true } catch { }

    # Adjust Policy Type label and combo so the label aligns vertically with the dropdown
    try {
        $LabelObjBoxPolicyType.AutoSize = $false
        $LabelObjBoxPolicyType.Dock = 'Fill'
        $LabelObjBoxPolicyType.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
        $LabelObjBoxPolicyType.MinimumSize = New-Object System.Drawing.Size(100,24)
    } catch { }
    try {
        $ObjBoxPolicyType.Dock = 'Fill'
        $ObjBoxPolicyType.MinimumSize = New-Object System.Drawing.Size(200,24)
    } catch { }

    $Form.Controls.Add($TopTable)
    $TopTable.BringToFront()

    # Create a middle TableLayoutPanel with three rows (project policy, repo policy, default/exit)
    $MiddlePanel = New-Object System.Windows.Forms.TableLayoutPanel
    $middleY = $TopTable.Location.Y + $TopTable.Height + 6
    $MiddlePanel.Location = New-Object System.Drawing.Point -ArgumentList 10, $middleY
    try { $mpWidth = $Form.ClientSize.Width - 40 } catch { $mpWidth = 1200 }
    $MiddlePanel.Size = New-Object System.Drawing.Size([Math]::Max(600,$mpWidth),220)
    $MiddlePanel.ColumnCount = 1
    $MiddlePanel.RowCount = 3
    $MiddlePanel.ColumnStyles.Clear()
    $MiddlePanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $MiddlePanel.RowStyles.Clear()
    $MiddlePanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $MiddlePanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $MiddlePanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $MiddlePanel.Padding = New-Object System.Windows.Forms.Padding(6)
    $MiddlePanel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

    $projectPolicyPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $projectPolicyPanel.FlowDirection = 'LeftToRight'
    $projectPolicyPanel.WrapContents = $true
    $projectPolicyPanel.AutoSize = $true
    $projectPolicyPanel.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $projectPolicyPanel.Dock = 'Fill'
    $projectPolicyPanel.Padding = New-Object System.Windows.Forms.Padding(0)

    $repoPolicyPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $repoPolicyPanel.FlowDirection = 'LeftToRight'
    $repoPolicyPanel.WrapContents = $true
    $repoPolicyPanel.AutoSize = $true
    $repoPolicyPanel.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $repoPolicyPanel.Dock = 'Fill'
    $repoPolicyPanel.Padding = New-Object System.Windows.Forms.Padding(0)

    $defaultPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $defaultPanel.FlowDirection = 'LeftToRight'
    $defaultPanel.WrapContents = $true
    $defaultPanel.AutoSize = $true
    $defaultPanel.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $defaultPanel.Dock = 'Fill'
    $defaultPanel.Padding = New-Object System.Windows.Forms.Padding(0)

    $projectPolicyControls = @(
        $ButtonQueryProjectPolicy, $ButtonEnableProjectPolicy, $ButtonDisableProjectPolicy, $ButtonDeleteProjectPolicy
    )
    $repoPolicyControls = @(
        $ButtonQueryReposPolicy, $ButtonEnableReposPolicy, $ButtonDisableReposPolicy, $ButtonDeleteReposPolicy
    )
    $defaultControls = @(
        $ButtonQueryRepos, $ButtonSetDefaultBranch, $ButtonExit
    )

    $rowPanels = @(
        @{panel=$projectPolicyPanel; controls=$projectPolicyControls},
        @{panel=$repoPolicyPanel; controls=$repoPolicyControls},
        @{panel=$defaultPanel; controls=$defaultControls}
    )

    foreach ($row in $rowPanels) {
        $panel = $row.panel
        foreach ($lc in $row.controls) {
            if ($null -ne $lc) {
                try { $Form.Controls.Remove($lc) } catch { }
                try { if ($null -ne $lc.Parent) { $lc.Parent.Controls.Remove($lc) } } catch { }
                $lc.Size = New-Object System.Drawing.Size(160,40)
                try { $lc.MinimumSize = New-Object System.Drawing.Size(140,40) } catch { }
                $lc.Margin = New-Object System.Windows.Forms.Padding(8)
                try { $panel.Controls.Add($lc) | Out-Null } catch { }
            }
        }
    }

    $MiddlePanel.Controls.Add($projectPolicyPanel, 0, 0)
    $MiddlePanel.Controls.Add($repoPolicyPanel, 0, 1)
    $MiddlePanel.Controls.Add($defaultPanel, 0, 2)

    $Form.Controls.Add($MiddlePanel)
    $MiddlePanel.BringToFront()

    # Ensure Exit button is visible and has a minimum size inside the middle panel
    try { $ButtonExit.MinimumSize = New-Object System.Drawing.Size(140,40); $ButtonExit.Visible = $true } catch { }

    # Add the results TextBox after panels so it's parented only once and anchors correctly
    try { $Form.Controls.Remove($TextBoxResult) } catch { }
    $TextBoxResult.Location = New-Object System.Drawing.Point(20,($TopTable.Height + $MiddlePanel.Height + 24))

    # Safely determine the form client size (some hosts may leave $Form or ClientSize as an array)
    $formClientSize = $null
    if ($null -ne $Form) {
        try {
            $cs = $Form.ClientSize
            if ($cs -is [System.Array]) { $formClientSize = $cs[0] } else { $formClientSize = $cs }
        } catch {
            $formClientSize = New-Object System.Drawing.Size(1020,720)
        }
    } else {
        $formClientSize = New-Object System.Drawing.Size(1020,720)
    }

    $width = [int]$formClientSize.Width
    $height = [int]$formClientSize.Height
    $w = [Math]::Max(100, $width - 40)
    $h = [Math]::Max(100, $height - ($TopTable.Height + $MiddlePanel.Height + 48))

    $TextBoxResult.Size = New-Object System.Drawing.Size($w, $h)
    $TextBoxResult.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $Form.Controls.Add($TextBoxResult)
}

loadProjects
[void] $Form.ShowDialog()
