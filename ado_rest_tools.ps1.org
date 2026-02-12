###===================================================================================================
##Load Windows Form Type
##---------------------------------------------------------------------------------------------------
#Add Windows Form Type
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#Add Visual Basic Form Type
[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

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

###===================================================================================================
##Define Azure DevOps Server Response (Repos Base)
##---------------------------------------------------------------------------------------------------
#Define Repos URL and Response
#$reposUrl = "https://$ADOServerFQDN/$collection/$projectName/_apis/git/repositories?api-version=6.0"
#$reposResponse = Invoke-RestMethod -Method Get -Uri $reposUrl -Headers $headers

###===================================================================================================
##Define Windows Form
##---------------------------------------------------------------------------------------------------
#Define Main Windows Form
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Azure DevOps Server REST API Tool"
$Form.StartPosition = "CenterScreen"
$Form.BackColor = "DarkGray"
$Form.Size = New-Object System.Drawing.Size(1000,570)

###===================================================================================================
##Define Input Text Box
##---------------------------------------------------------------------------------------------------
#Define Project Input Text Box Label
$TextBoxProject = New-Object System.Windows.Forms.TextBox
$TextBoxProject.Location = New-Object System.Drawing.Point(20,40)
$TextBoxProject.Size = New-Object System.Drawing.Size(100,40)

$Form.Controls.Add($TextBoxProject)

###===================================================================================================
##Define Output Text Box
##---------------------------------------------------------------------------------------------------
#Define Result Output Text Box
$TextBoxResult = New-Object System.Windows.Forms.TextBox
$TextBoxResult.Location = New-Object System.Drawing.Size(20,200)
$TextBoxResult.Size = New-Object System.Drawing.Size(940,310)
$TextBoxResult.ForeColor = "White"
$TextBoxResult.BackColor = "Black"
$TextBoxResult.ScrollBars = "Vertical"
$TextBoxResult.MultiLine = $True

$Form.Controls.Add($TextBoxResult)

###===================================================================================================
##Define Form Label
##---------------------------------------------------------------------------------------------------
#Define Project Input Text Box Label
$LabelTextBoxProject = New-Object System.Windows.Forms.Label
$LabelTextBoxProject.Location = New-Object System.Drawing.Point(20,20)
$LabelTextBoxProject.Size = New-Object System.Drawing.Size(100,20)
$LabelTextBoxProject.Text = 'Project name:'

$Form.Controls.Add($LabelTextBoxProject)

#Define Project Input Text Box Label
$LabelSelfApproval = New-Object System.Windows.Forms.Label
$LabelSelfApproval.Location = New-Object System.Drawing.Point(500,30)
$LabelSelfApproval.Size = New-Object System.Drawing.Size(100,40)
$LabelSelfApproval.Text = 'Enable/Disable Self Approval:'

$Form.Controls.Add($LabelSelfApproval)

#Define Using Project Label
$LabelUsingProject = New-Object System.Windows.Forms.Label
$LabelUsingProject.Location = New-Object System.Drawing.Point(20,80)
$LabelUsingProject.Size = New-Object System.Drawing.Size(100,20)
$LabelUsingProject.Text = 'Using Project:'

$Form.Controls.Add($LabelUsingProject)

#Define Display Project Label
$LabelDisplayProject = New-Object System.Windows.Forms.Label
$LabelDisplayProject.Location = New-Object System.Drawing.Point(20,100)
$LabelDisplayProject.Size = New-Object System.Drawing.Size(100,20)
$LabelDisplayProject.Text = ''

$Form.Controls.Add($LabelDisplayProject)

###===================================================================================================
##Define Form Dropdown Menu
##---------------------------------------------------------------------------------------------------
#Define Branch Dropdown menu label
$LabelObjBoxBranch = New-Object System.Windows.Forms.Label
$LabelObjBoxBranch.Location = New-Object System.Drawing.Point(260,20)
$LabelObjBoxBranch.Size = New-Object System.Drawing.Size(100,20)
$LabelObjBoxBranch.Text = 'Set default branch:'

$Form.Controls.Add($LabelObjBoxBranch)

#Define Branch Dropdown menu
$ObjBoxBranch = New-Object System.Windows.Forms.ComboBox
$ObjBoxBranch.Location = New-Object System.Drawing.Size(260,40)
$ObjBoxBranch.Size = New-Object System.Drawing.Size(100,40)
$ObjBoxBranch.Enabled = $false

#Define Branch Dropdoen menu value
$ObjBoxBranch.Items.Add('Develop') | Out-Null
$ObjBoxBranch.Items.Add('Main') | Out-Null

$Form.Controls.Add($ObjBoxBranch)

##---------------------------------------------------------------------------------------------------
#Define Policy Type Dropdown menu label
$LabelObjBoxPolicyType = New-Object System.Windows.Forms.Label
$LabelObjBoxPolicyType.Location = New-Object System.Drawing.Point(620,95)
$LabelObjBoxPolicyType.Size = New-Object System.Drawing.Size(70,20)
$LabelObjBoxPolicyType.Text = 'Policy Type:'
$LabelObjBoxPolicyType.Enabled = $false

$Form.Controls.Add($LabelObjBoxPolicyType)

#Define Policy Type Dropdown menu
$ObjBoxPolicyType = New-Object System.Windows.Forms.ComboBox
$ObjBoxPolicyType.Location = New-Object System.Drawing.Size(690,95)
$ObjBoxPolicyType.Size = New-Object System.Drawing.Size(160,40)
$ObjBoxPolicyType.Enabled = $false

#Define Project Type Dropdown menu value
$ObjBoxPolicyType.Items.Add('Minimum number of reviewers') | Out-Null
$ObjBoxPolicyType.Items.Add('Main') | Out-Null

$Form.Controls.Add($ObjBoxPolicyType)

###===================================================================================================
##Define Form Button
##---------------------------------------------------------------------------------------------------
#Define Button of query all repos
$ButtonQueryRepos = New-Object System.Windows.Forms.Button
$ButtonQueryRepos.Location = New-Object System.Drawing.Size(140,20)
$ButtonQueryRepos.Size = New-Object System.Drawing.Size(100,40)
$ButtonQueryRepos.Text = "Query Default Branch"
$ButtonQueryRepos.Enabled = $false
$ButtonQueryRepos.Add_Click( { queryAllRepo } )

$Form.Controls.Add($ButtonQueryRepos)

##---------------------------------------------------------------------------------------------------
#Define Button of update repos default branch
$ButtonSetDefaultBranch = New-Object System.Windows.Forms.Button
$ButtonSetDefaultBranch.Location = New-Object System.Drawing.Size(380,20)
$ButtonSetDefaultBranch.Size = New-Object System.Drawing.Size(100,40)
$ButtonSetDefaultBranch.Text = "Set Default Branch"
$ButtonSetDefaultBranch.Enabled = $false
$ButtonSetDefaultBranch.Add_Click( { setDefaultBranch } )

$Form.Controls.Add($ButtonSetDefaultBranch)

##---------------------------------------------------------------------------------------------------
#Define Button of query project policy
$ButtonQueryProjectPolicy = New-Object System.Windows.Forms.Button
$ButtonQueryProjectPolicy.Location = New-Object System.Drawing.Size(140,80)
$ButtonQueryProjectPolicy.Size = New-Object System.Drawing.Size(100,40)
$ButtonQueryProjectPolicy.Text = "Query Project Policy"
$ButtonQueryProjectPolicy.Enabled = $false
$ButtonQueryProjectPolicy.Add_Click( { queryProjectPolicy } )

$Form.Controls.Add($ButtonQueryProjectPolicy)

##---------------------------------------------------------------------------------------------------
#Define Button of disable project policy
$ButtonDisableProjectPolicy = New-Object System.Windows.Forms.Button
$ButtonDisableProjectPolicy.Location = New-Object System.Drawing.Size(260,80)
$ButtonDisableProjectPolicy.Size = New-Object System.Drawing.Size(100,40)
$ButtonDisableProjectPolicy.Text = "Disable Project Policy"
$ButtonDisableProjectPolicy.Enabled = $false
$ButtonDisableProjectPolicy.Add_Click( { disableProjectPolicy } )

$Form.Controls.Add($ButtonDisableProjectPolicy)

##---------------------------------------------------------------------------------------------------
#Define Button of enable project policy
$ButtonEnableProjectPolicy = New-Object System.Windows.Forms.Button
$ButtonEnableProjectPolicy.Location = New-Object System.Drawing.Size(380,80)
$ButtonEnableProjectPolicy.Size = New-Object System.Drawing.Size(100,40)
$ButtonEnableProjectPolicy.Text = "Enable Project Policy"
$ButtonEnableProjectPolicy.Enabled = $false
$ButtonEnableProjectPolicy.Add_Click( { enableProjectPolicy } )

$Form.Controls.Add($ButtonEnableProjectPolicy)

##---------------------------------------------------------------------------------------------------
#Define Button of delete project policy
$ButtonDeleteProjectPolicy = New-Object System.Windows.Forms.Button
$ButtonDeleteProjectPolicy.Location = New-Object System.Drawing.Size(500,80)
$ButtonDeleteProjectPolicy.Size = New-Object System.Drawing.Size(100,40)
$ButtonDeleteProjectPolicy.Text = "Delete Project Policy"
$ButtonDeleteProjectPolicy.Enabled = $false
$ButtonDeleteProjectPolicy.Add_Click( { deleteProjectPolicy } )

$Form.Controls.Add($ButtonDeleteProjectPolicy)

##---------------------------------------------------------------------------------------------------
#Define Button of query repositories policy
$ButtonQueryReposPolicy = New-Object System.Windows.Forms.Button
$ButtonQueryReposPolicy.Location = New-Object System.Drawing.Size(140,140)
$ButtonQueryReposPolicy.Size = New-Object System.Drawing.Size(100,40)
$ButtonQueryReposPolicy.Text = "Query Repos Policy"
$ButtonQueryReposPolicy.Enabled = $false
$ButtonQueryReposPolicy.Add_Click( { queryReposPolicy } )

$Form.Controls.Add($ButtonQueryReposPolicy)

##---------------------------------------------------------------------------------------------------
#Define Button of disable repositories policy
$ButtonDisableReposPolicy = New-Object System.Windows.Forms.Button
$ButtonDisableReposPolicy.Location = New-Object System.Drawing.Size(260,140)
$ButtonDisableReposPolicy.Size = New-Object System.Drawing.Size(100,40)
$ButtonDisableReposPolicy.Text = "Disable Repos Policy"
$ButtonDisableReposPolicy.Enabled = $false
$ButtonDisableReposPolicy.Add_Click( { disableReposPolicy } )

$Form.Controls.Add($ButtonDisableReposPolicy)

##---------------------------------------------------------------------------------------------------
#Define Button of enable repositories policy
$ButtonEnableReposPolicy = New-Object System.Windows.Forms.Button
$ButtonEnableReposPolicy.Location = New-Object System.Drawing.Size(380,140)
$ButtonEnableReposPolicy.Size = New-Object System.Drawing.Size(100,40)
$ButtonEnableReposPolicy.Text = "Enable Project Policy"
$ButtonEnableReposPolicy.Enabled = $false
$ButtonEnableReposPolicy.Add_Click( { enableReposPolicy } )

$Form.Controls.Add($ButtonEnableReposPolicy)

##---------------------------------------------------------------------------------------------------
#Define Button of delete repositories policy
$ButtonDeleteReposPolicy = New-Object System.Windows.Forms.Button
$ButtonDeleteReposPolicy.Location = New-Object System.Drawing.Size(500,140)
$ButtonDeleteReposPolicy.Size = New-Object System.Drawing.Size(100,40)
$ButtonDeleteReposPolicy.Text = "Delete Repos Policy"
$ButtonDeleteReposPolicy.Enabled = $false
$ButtonDeleteReposPolicy.Add_Click( { deleteReposPolicy } )

$Form.Controls.Add($ButtonDeleteReposPolicy)

##---------------------------------------------------------------------------------------------------
#Define Button of query repositories self approval
$ButtonQuerySelfApproval = New-Object System.Windows.Forms.Button
$ButtonQuerySelfApproval.Location = New-Object System.Drawing.Size(620,20)
$ButtonQuerySelfApproval.Size = New-Object System.Drawing.Size(100,40)
$ButtonQuerySelfApproval.Text = "Query Self Approval"
$ButtonQuerySelfApproval.Enabled = $false
$ButtonQuerySelfApproval.Add_Click( { querySelfApproval } )

$Form.Controls.Add($ButtonQuerySelfApproval)

##---------------------------------------------------------------------------------------------------
#Define Button of enable repositories self approval
$ButtonEnableSelfApproval = New-Object System.Windows.Forms.Button
$ButtonEnableSelfApproval.Location = New-Object System.Drawing.Size(740,20)
$ButtonEnableSelfApproval.Size = New-Object System.Drawing.Size(100,40)
$ButtonEnableSelfApproval.Text = "Enable Self Approval"
$ButtonEnableSelfApproval.Enabled = $false
$ButtonEnableSelfApproval.Add_Click( { enableSelfApproval } )

$Form.Controls.Add($ButtonEnableSelfApproval)

##---------------------------------------------------------------------------------------------------
#Define Button of disable repositories self approval
$ButtonDisableSelfApproval = New-Object System.Windows.Forms.Button
$ButtonDisableSelfApproval.Location = New-Object System.Drawing.Size(860,20)
$ButtonDisableSelfApproval.Size = New-Object System.Drawing.Size(100,40)
$ButtonDisableSelfApproval.Text = "Disable Self Approval"
$ButtonDisableSelfApproval.Enabled = $false
$ButtonDisableSelfApproval.Add_Click( { disableSelfApproval } )

$Form.Controls.Add($ButtonDisableSelfApproval)

##---------------------------------------------------------------------------------------------------
#Define Button of exit form
$ButtonExit = New-Object System.Windows.Forms.Button
$ButtonExit.Location = New-Object System.Drawing.Size(860,140)
$ButtonExit.Size = New-Object System.Drawing.Size(100,40)
$ButtonExit.Text = "Exit"
$ButtonExit.Add_Click({$Form.Close()})

$Form.Controls.Add($ButtonExit)

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
    $projectName = $TextBoxProject.Text.ToString()
    $timeStamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")

    $TextBoxResult.Clear()

    $formTitle = "Query default branch of all repositories in project $($projectName) Result"
    $Form.Text = $formTitle

    $reposUrl = "https://$ADOServerFQDN/$collection/$projectName/_apis/git/repositories?api-version=6.0"
    $reposResponse = Invoke-RestMethod -Method Get -Uri $reposUrl -Headers $headers

    $TextBoxResult.Text = "$($timeStamp) The default branch of each repository:"

    foreach ($repo in $reposResponse.value) {
        $timeStamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
        $TextBoxResult.AppendText("`r`n")
        $TextBoxResult.Appendtext("$($timestamp) Project: $($projectName), Repository: $($repo.name), ID: $($repo.id), Default Branch: $($repo.defaultBranch)")
    }
    [System.Windows.MessageBox]::Show("Query of default branch in all repositories`n of project $projectName completed","Process Result")
    [string[]]$logOutput = $TextBoxResult.Text
    $logOutput = $logOutput.Replace("`r`n", "`r`n")
    writeLog
}

##---------------------------------------------------------------------------------------------------
#Define Set Default Branch Function
function setDefaultBranch {
    $projectName = $TextBoxProject.Text.ToString()
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

    $reposUrl = "https://$ADOServerFQDN/$collection/$projectName/_apis/git/repositories?api-version=6.0"
    $reposResponse = Invoke-RestMethod -Method Get -Uri $reposUrl -Headers $headers

    $TextBoxResult.Text = "$($timeStamp) Setting default branch to $($toBranch.ToUpper()) of each repos:"

    foreach ($repo in $reposResponse.value) {
        $timeStamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
        $TextBoxResult.AppendText("`r`n")
        $TextBoxResult.AppendText("$($timeStamp) Setting default branch to $($toBranch.ToUpper()) of repository $($repo.name):")
        $TextBoxResult.AppendText("`r`n")
        $TextBoxResult.AppendText("$($timeStamp) Project: $($projectName), Repository: $($repo.name), ID: $($repo.id), Default Branch: $($repo.defaultBranch)")
        $branchUrl = "https://$ADOServerFQDN/$collection/$projectName/_apis/git/repositories/$($repo.id)?api-version=6.0"
        $TextBoxResult.AppendText("`r`n")
        $TextBoxResult.AppendText("$($timeStamp) Setting the default branch......")
        #$branchRESTPatch = Invoke-RestMethod -Uri $branchUrl -Method Patch -Headers $headers -ContentType "application/json" -Body ($jsonData | ConvertTo-Json)
        Invoke-RestMethod -Uri $branchUrl -Method Patch -Headers $headers -ContentType "application/json" -Body ($jsonData | ConvertTo-Json)
        $branchRESTGet = Invoke-RestMethod -Uri $branchUrl -Method Get -Headers $headers
        $TextBoxResult.AppendText("`r`n")
        $TextBoxResult.AppendText("$($timeStamp) Project: $($projectName), Repository: $($repo.name), ID: $($repo.id), Default Branch: $($branchRESTGet.defaultBranch)")
        [System.Windows.MessageBox]::Show("Setting default branch in repository`n $($repo.name) of project $projectName completed","Process Result")
    }

    [System.Windows.MessageBox]::Show("Setting default branch on all repositories of project $projectName completed","Process Result")
    [string[]]$logOutput = $TextBoxResult.Text
    $logOutput = $logOutput.Replace("`r`n", "`r`n")
    writeLog
}

###===================================================================================================
##Define Project Base Policy Function
##---------------------------------------------------------------------------------------------------
#Define Function to query project base policy
function queryProjectPolicy {
    $projectName = $TextBoxProject.Text.ToString()
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

    $policiesUrl = "https://$ADOServerFQDN/$collection/$projectName/_apis/policy/configurations?api-version=6.0"
    $policiesResponse = Invoke-RestMethod -Method Get -Uri $policiesUrl -Headers $headers

    foreach ($policy in $policiesResponse.value) {
        if ($null -eq $policy.settings.scope.repositoryId -And $null -eq $policy.settings.scope.refName) {
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

    [System.Windows.MessageBox]::Show("Query reporistoty policies of project`n$projectName completed","Process Result")
    [string[]]$logOutput = $TextBoxResult.Text
    $logOutput = $logOutput.Replace("`r`n", "`r`n")
    writeLog
}

##---------------------------------------------------------------------------------------------------
#Define Function to disable project base policy
function disableProjectPolicy {
    $projectName = $TextBoxProject.Text.ToString()
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

    $policiesUrl = "https://$ADOServerFQDN/$collection/$projectName/_apis/policy/configurations?api-version=6.0"
    $policiesResponse = Invoke-RestMethod -Method Get -Uri $policiesUrl -Headers $headers

    foreach ($policy in $policiesResponse.value) {
        if ($null -eq $policy.settings.scope.repositoryId -And $null -eq $policy.settings.scope.refName) {
            $policiesPutUrl = "https://$ADOServerFQDN/DevOpsCollection/$projectName/_apis/policy/configurations/$($policy.id)?api-version=6.0"
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

    [System.Windows.MessageBox]::Show("Disable project level policies of project`n$projectName completed","Process Result")
    [string[]]$logOutput = $TextBoxResult.Text
    $logOutput = $logOutput.Replace("`r`n", "`r`n")
    writeLog
}

##---------------------------------------------------------------------------------------------------
#Define Function to enable project base policy
function enableProjectPolicy {
    $projectName = $TextBoxProject.Text.ToString()
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

    $policiesUrl = "https://$ADOServerFQDN/$collection/$projectName/_apis/policy/configurations?api-version=6.0"
    $policiesResponse = Invoke-RestMethod -Method Get -Uri $policiesUrl -Headers $headers

    foreach ($policy in $policiesResponse.value) {
        if ($null -eq $policy.settings.scope.repositoryId -And $null -eq $policy.settings.scope.refName) {
            $policiesPutUrl = "https://$ADOServerFQDN/DevOpsCollection/$projectName/_apis/policy/configurations/$($policy.id)?api-version=6.0"
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

    [System.Windows.MessageBox]::Show("Enable project level policies of project`n$projectName completed","Process Result")
    [string[]]$logOutput = $TextBoxResult.Text
    $logOutput = $logOutput.Replace("`r`n", "`r`n")
    writeLog
}

##---------------------------------------------------------------------------------------------------
#Define Function to delete project base policy
function deleteProjectPolicy {
    $projectName = $TextBoxProject.Text.ToString()
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

        $policiesUrl = "https://$ADOServerFQDN/$collection/$projectName/_apis/policy/configurations?api-version=6.0"
        $policiesResponse = Invoke-RestMethod -Method Get -Uri $policiesUrl -Headers $headers

        foreach ($policy in $policiesResponse.value) {
            if ($null -eq $policy.settings.scope.repositoryId -And $null -eq $policy.settings.scope.refName) {
                $policyPutUrl = "https://$ADOServerFQDN/$collection/$projectName/_apis/policy/configurations/$($policy.id)?api-version=6.0"
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

    [System.Windows.MessageBox]::Show("Delete project level policies of project`n$projectName completed","Process Result")
    [string[]]$logOutput = $TextBoxResult.Text
    $logOutput = $logOutput.Replace("`r`n", "`r`n")
    writeLog
}

###===================================================================================================
##Define Repository Base Policy Function
##---------------------------------------------------------------------------------------------------
#Define Function to query repository base policy
function queryReposPolicy {
    $projectName = $TextBoxProject.Text.ToString()
    $timeStamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")

    $TextBoxResult.Clear()

    $formTitle = "Query repository policy in project $($projectName) Result"
    $Form.Text = $formTitle

    $TextBoxResult.Text = "$($timeStamp) The repository policy on each repository:"
    $TextBoxResult.AppendText("`r`n")
    $TextBoxResult.Appendtext("$($timeStamp) Project: $($projectName)")
    $TextBoxResult.Appendtext("`n`n")

    $policiesObject = [System.Collections.Generic.List[object]]::new()

    $reposUrl = "https://$ADOServerFQDN/$collection/$projectName/_apis/git/repositories?api-version=6.0"
    $reposResponse = Invoke-RestMethod -Method Get -Uri $reposUrl -Headers $headers

    $policiesUrl = "https://$ADOServerFQDN/$collection/$projectName/_apis/policy/configurations?api-version=6.0"
    $policiesResponse = Invoke-RestMethod -Method Get -Uri $policiesUrl -Headers $headers

    foreach ($policy in $policiesResponse.value) {
        if ($null -eq $policy.settings.scope.repositoryId -And $null -eq $policy.settings.scope.refName) {
            foreach ($repo in $reposResponse.value) {
                if ($repo.id -eq $policy.settings.scope.repositoryId) {
                     $policyRepo = $repo.name
                }
                if ($policy.settings.scope.refName -eq "refs/heads/main") {
                    $policyBranch = "Main"
                } elseif ($policy.settings.scope.refName -eq "refs/heads/develop") {
                    $policyBranch = "Develop"
                }
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

    [System.Windows.MessageBox]::Show("Query reporistoty policies of project`n$projectName completed","Process Result")
    [string[]]$logOutput = $TextBoxResult.Text
    $logOutput = $logOutput.Replace("`r`n", "`r`n")
    writeLog
}

##---------------------------------------------------------------------------------------------------
#Define Function to disable repository base policy
function disableReposPolicy {
    $projectName = $TextBoxProject.Text.ToString()
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

    $reposUrl = "https://$ADOServerFQDN/$collection/$projectName/_apis/git/repositories?api-version=6.0"
    $reposResponse = Invoke-RestMethod -Method Get -Uri $reposUrl -Headers $headers

    $policiesUrl = "https://$ADOServerFQDN/$collection/$projectName/_apis/policy/configurations?api-version=6.0"
    $policiesResponse = Invoke-RestMethod -Method Get -Uri $policiesUrl -Headers $headers

    foreach ($policy in $policiesResponse.value) {
        if ($null -ne $policy.settings.scope.repositoryId -And $null -ne $policy.settings.scope.refName) {
            $policiesPutUrl = "https://$ADOServerFQDN/DevOpsCollection/$projectName/_apis/policy/configurations/$($policy.id)?api-version=6.0"
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

    [System.Windows.MessageBox]::Show("Disable repositories policy of project`n$projectName completed","Process Result")
    [string[]]$logOutput = $TextBoxResult.Text
    $logOutput = $logOutput.Replace("`r`n", "`r`n")
    writeLog
}

##---------------------------------------------------------------------------------------------------
#Define Function to enable repository base policy
function enableReposPolicy {
    $projectName = $TextBoxProject.Text.ToString()
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

    $reposUrl = "https://$ADOServerFQDN/$collection/$projectName/_apis/git/repositories?api-version=6.0"
    $reposResponse = Invoke-RestMethod -Method Get -Uri $reposUrl -Headers $headers

    $policiesUrl = "https://$ADOServerFQDN/$collection/$projectName/_apis/policy/configurations?api-version=6.0"
    $policiesResponse = Invoke-RestMethod -Method Get -Uri $policiesUrl -Headers $headers

    foreach ($policy in $policiesResponse.value) {
        if ($null -ne $policy.settings.scope.repositoryId -And $null -ne $policy.settings.scope.refName) {
            $policiesPutUrl = "https://$ADOServerFQDN/DevOpsCollection/$projectName/_apis/policy/configurations/$($policy.id)?api-version=6.0"
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

    [System.Windows.MessageBox]::Show("Enable repositories policy of project`n$projectName completed","Process Result")
    [string[]]$logOutput = $TextBoxResult.Text
    $logOutput = $logOutput.Replace("`r`n", "`r`n")
    writeLog
}

##---------------------------------------------------------------------------------------------------
#Define Function to delete repository base policy
function deleteReposPolicy {
    $projectName = $TextBoxProject.Text.ToString()
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

        $policiesUrl = "https://$ADOServerFQDN/$collection/$projectName/_apis/policy/configurations?api-version=6.0"
        $policiesResponse = Invoke-RestMethod -Method Get -Uri $policiesUrl -Headers $headers

        foreach ($policy in $policiesResponse.value) {
            if ($null -ne $policy.settings.scope.repositoryId -And $null -ne $policy.settings.scope.refName) {
                $policyPutUrl = "https://$ADOServerFQDN/$collection/$projectName/_apis/policy/configurations/$($policy.id)?api-version=6.0"
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

    [System.Windows.MessageBox]::Show("Delete project level policies of project`n$projectName completed","Process Result")
    [string[]]$logOutput = $TextBoxResult.Text
    $logOutput = $logOutput.Replace("`r`n", "`r`n")
    writeLog
}

###===================================================================================================
##Define Function of Self Approval
##---------------------------------------------------------------------------------------------------
#Define Function to Query self approval
function querySelfApproval {
    $projectName = $TextBoxProject.Text.ToString()
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

    $reposUrl = "https://$ADOServerFQDN/$collection/$projectName/_apis/git/repositories?api-version=6.0"
    $reposResponse = Invoke-RestMethod -Method Get -Uri $reposUrl -Headers $headers

    $policiesUrl = "https://$ADOServerFQDN/$collection/$projectName/_apis/policy/configurations?api-version=6.0"
    $policiesResponse = Invoke-RestMethod -Method Get -Uri $policiesUrl -Headers $headers

    foreach ($policy in $policiesResponse.value) {
        if ($policy.Type.id -eq "fa4e907d-c16b-4a4c-9dfa-4906e5d171dd") {
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

    [System.Windows.MessageBox]::Show("Query repositories self approval of project`n$projectName completed","Process Result")
    [string[]]$logOutput = $TextBoxResult.Text
    $logOutput = $logOutput.Replace("`r`n", "`r`n")
    writeLog
}

#$projectName = "data-platform"
#$repoName = "patron-profile-service-net-api", "patron-open-rating-net-consumer", "patron-summary-service-net-api", "patron-mpb-balance-net-consumer", "patron-personal-details-net-consumer", "patron-card-info-net-consumer", "patron-wallet-service-net-api", "patron-vvip-info-net-consumer", "patron-exclusion-net-consumer", "patron-ewallet-balance-net-consumer", "patron-coded-info-net-consumer", "patron-room-booking-net-consumer", "patron-preference-net-consumer", "patron-wechat-binding-net-consumer", "patron-srapp-status-net-consumer", "patron-uploaded-value-net-consumer", "patron-trip-service-net-api", "remove-open-ratings-job"

##---------------------------------------------------------------------------------------------------
#Define Function to Enable self approval
function enableSelfApproval {
    $projectName = $TextBoxProject.Text.ToString()
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

    $reposUrl = "https://$ADOServerFQDN/$collection/$projectName/_apis/git/repositories?api-version=6.0"
    $reposResponse = Invoke-RestMethod -Method Get -Uri $reposUrl -Headers $headers

    $policiesUrl = "https://$ADOServerFQDN/$collection/$projectName/_apis/policy/configurations?api-version=6.0"
    $policiesResponse = Invoke-RestMethod -Method Get -Uri $policiesUrl -Headers $headers

    foreach ($policy in $policiesResponse.value) {
        if ($policy.Type.id -eq "fa4e907d-c16b-4a4c-9dfa-4906e5d171dd") {
            $policiesPutUrl = "https://$ADOServerFQDN/DevOpsCollection/$projectName/_apis/policy/configurations/$($policy.id)?api-version=6.0"
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

    [System.Windows.MessageBox]::Show("Enable repositories self approval of project`n$projectName completed","Process Result")
    [string[]]$logOutput = $TextBoxResult.Text
    $logOutput = $logOutput.Replace("`r`n", "`r`n")
    writeLog
}

##---------------------------------------------------------------------------------------------------
#Define Function to disable self approval
function disableSelfApproval {
    $projectName = $TextBoxProject.Text.ToString()
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

    $reposUrl = "https://$ADOServerFQDN/$collection/$projectName/_apis/git/repositories?api-version=6.0"
    $reposResponse = Invoke-RestMethod -Method Get -Uri $reposUrl -Headers $headers

    $policiesUrl = "https://$ADOServerFQDN/$collection/$projectName/_apis/policy/configurations?api-version=6.0"
    $policiesResponse = Invoke-RestMethod -Method Get -Uri $policiesUrl -Headers $headers

    foreach ($policy in $policiesResponse.value) {
       if ($policy.Type.id -eq "fa4e907d-c16b-4a4c-9dfa-4906e5d171dd") {
            $policiesPutUrl = "https://$ADOServerFQDN/DevOpsCollection/$projectName/_apis/policy/configurations/$($policy.id)?api-version=6.0"
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

    [System.Windows.MessageBox]::Show("Disable repositories self approval of project`n$projectName completed","Process Result")
    [string[]]$logOutput = $TextBoxResult.Text
    $logOutput = $logOutput.Replace("`r`n", "`r`n")
    writeLog
}

###===================================================================================================
##Button Enable Disable Checking
##---------------------------------------------------------------------------------------------------
#Enable/Disbaled Button until project name is inputed
$TextBoxProject.Add_TextChanged({
    if ($TextBoxProject.TextLength -gt 0) {
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
        $LabelDisplayProject.Text = $TextBoxProject.Text
        $LabelDisplayProject.BackColor = "Black"
        $LabelDisplayProject.ForeColor = "White"
        if ($ObjBoxBranch.Selectedindex -ne -1) {
            $ButtonSetDefaultBranch.Enabled = $true
        } else {
            $ButtonSetDefaultBranch.Enabled = $false
        }
    } else {
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
    }
})

##---------------------------------------------------------------------------------------------------
#Enable/Disbaled Button until default branch is selected
$ObjBoxBranch.Add_SelectedIndexChanged({
    if ($ObjBoxBranch.Selectedindex -ne -1) {
        $ButtonSetDefaultBranch.Enabled = $true
#        if ($TextBoxProject.TextLength -gt 0) {
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
#Loading Windows Form
[void] $Form.ShowDialog()
