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
$script:AuthMode      = "PAT"   # PAT, Basic, or Windows
$script:SkipSslValidation = $false
$script:SslRetryUsed  = $false
$script:Headers       = @{}

# Azure DevOps Server requires TLS 1.2. Do not OR in TLS 1.3: on Windows PowerShell 5.1
# that flag often breaks the handshake ("unexpected error occurred on a send").
function Initialize-Tls {
    if (-not ("AdoServerTls" -as [type])) {
        Add-Type @"
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
public static class AdoServerTls {
    public static bool AcceptAny(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors) {
        return true;
    }
    public static void EnableTls12() {
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
        ServicePointManager.Expect100Continue = false;
    }
    public static void TrustAll() {
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
        ServicePointManager.Expect100Continue = false;
        ServicePointManager.ServerCertificateValidationCallback =
            new RemoteCertificateValidationCallback(AcceptAny);
    }
}
"@
    }
    [AdoServerTls]::EnableTls12()
}

function Enable-SkipSslValidation {
    Initialize-Tls
    # A PowerShell scriptblock callback is collected or throws, which closes the
    # connection with "An unexpected error occurred on a send". Use a static .NET method.
    [AdoServerTls]::TrustAll()
    $script:SkipSslValidation = $true
    $script:SslCallbackReady = $true
}

Initialize-Tls
$script:RepoListCache = @()
$script:SelectedRepoIds = @()
$script:SelectedRepoNames = @()
$script:ApiDetected = $false
$script:SslCallbackReady = $false

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

function Get-ErrorDetail {
    param($ErrorRecord)
    $parts = New-Object System.Collections.Generic.List[string]
    $ex = $ErrorRecord.Exception
    $guard = 0
    while ($null -ne $ex -and $guard -lt 6) {
        $line = $ex.GetType().FullName + ": " + $ex.Message
        if (-not $parts.Contains($line)) { [void]$parts.Add($line) }
        $ex = $ex.InnerException
        $guard++
    }
    return ($parts -join " | ")
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
    $mode = if ($Config.AuthMode) { $Config.AuthMode.Trim() } else { "PAT" }
    if ($mode -match '^(?i)(basic|pat)$') { $mode = "PAT" }
    elseif ($mode -match '^(?i)windows$') { $mode = "Windows" }
    $script:AuthMode = $mode
    $skipSsl = $false
    if ($null -ne $Config.SkipSslValidation) { $skipSsl = [bool]$Config.SkipSslValidation }
    elseif ($null -ne $Config.IgnoreSslErrors) { $skipSsl = [bool]$Config.IgnoreSslErrors }
    $script:SslRetryUsed = $false
    $script:ApiDetected = $false
    $script:PolicyTypesLoadedFromServer = $false
    if ($skipSsl) { Enable-SkipSslValidation }
    else {
        $script:SkipSslValidation = $false
        try { [Net.ServicePointManager]::ServerCertificateValidationCallback = $null } catch {}
    }
    $script:ApiVersion    = if ($Config.ApiVersion -and $Config.ApiVersion -ne "auto") { $Config.ApiVersion } else { "6.0" }
    Update-Headers
    $patStatus = if ([string]::IsNullOrWhiteSpace($script:MyPat)) { "MISSING" } else { "set (len=$($script:MyPat.Length))" }
    $prefixLog = if ($script:UrlPrefix) { " | Prefix: $($script:UrlPrefix)" } else { "" }
    $sslLog = if ($script:SkipSslValidation) { "skipped (lab or internal certificate)" } else { "enforced" }
    $authLog = if ($script:AuthMode -eq "PAT") { "PAT (Basic)" } else { $script:AuthMode }
    Write-Log "Configuration loaded successfully."
    Write-Log "  Name       : $($Config.Name)"
    Write-Log "  Server     : $($script:ADOServerFQDN)"
    Write-Log "  Collection : $($script:Collection)$prefixLog"
    Write-Log "  API Version: $($script:ApiVersion)"
    Write-Log "  Auth Mode  : $authLog"
    Write-Log "  SSL Check  : $sslLog"
    Write-Log "  PAT Status : $patStatus"
}

function Detect-ApiVersion {
    if ([string]::IsNullOrWhiteSpace($script:ADOServerFQDN)) { return "6.0" }
    if ($script:ApiDetected) { return $script:ApiVersion }
    $base = Get-BaseUrl
    $testVersions = @("7.2", "7.1", "7.0", "6.0")
    $firstUrl = $null
    $connectionError = $null
    foreach ($ver in $testVersions) {
        $url = "$base/$($script:Collection)/_apis/projects?api-version=$ver&`$top=1"
        if (-not $firstUrl) { $firstUrl = $url }
        try {
            $null = Invoke-AdoRest -Uri $url -Method Get -Quiet
            Write-Log "The server accepts API version $ver."
            $script:ApiVersion = $ver
            $script:ApiDetected = $true
            return $ver
        } catch {
            $detail = Get-ErrorDetail $_
            if ($detail -match "connection was closed|secure channel|trust relationship|unexpected error|timed out|Unable to connect|No such host|refused") {
                $connectionError = $detail
                break
            }
        }
    }
    $script:ApiDetected = $true
    if ($connectionError) {
        Write-Log "Cannot reach the server over HTTPS."
        Write-Log "  $connectionError"
        if ($firstUrl) { Write-Log "  URL tried: $firstUrl" }
        if (-not $script:UrlPrefix) {
            Write-Log "  If this server is installed under /tfs, set UrlPrefix to tfs and Collection to the real collection name."
        }
    } else {
        Write-Log "Could not auto-detect API version. Using 6.0 (Azure DevOps Server 2020 and later)."
        if ($firstUrl) { Write-Log "  URL tried: $firstUrl" }
    }
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
# PowerShell ConvertTo-Json turns a one-item array into an object. Policy "scope"
# must stay an array or the server rejects the update.
$script:JsonArrayProperties = @("scope","requiredReviewerIds","optionalReviewerIds","filenamePatterns")

function ConvertTo-PlainObject {
    param($Value, [string]$PropertyName)
    $forceArray = $PropertyName -and ($script:JsonArrayProperties -contains $PropertyName)
    if ($forceArray) {
        $arr = New-Object System.Collections.ArrayList
        foreach ($item in @($Value)) {
            if ($null -ne $item) { [void]$arr.Add((ConvertTo-PlainObject -Value $item)) }
        }
        return ,$arr.ToArray()
    }
    if ($null -eq $Value) { return $null }
    if ($Value -is [string]) { return $Value }
    if ($Value -is [datetime]) { return $Value.ToString("o") }
    if ($Value -is [ValueType]) { return $Value }
    if ($Value -is [System.Collections.IDictionary]) {
        $h = New-Object System.Collections.Hashtable
        foreach ($key in @($Value.Keys)) {
            $h[[string]$key] = ConvertTo-PlainObject -Value $Value[$key] -PropertyName ([string]$key)
        }
        return ,$h
    }
    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
        $arr = New-Object System.Collections.ArrayList
        foreach ($item in @($Value)) {
            if ($null -ne $item) { [void]$arr.Add((ConvertTo-PlainObject -Value $item)) }
        }
        return ,$arr.ToArray()
    }
    if ($Value -is [psobject]) {
        $h = New-Object System.Collections.Hashtable
        foreach ($prop in @($Value.PSObject.Properties)) {
            if ($prop.MemberType -ne "NoteProperty") { continue }
            $h[$prop.Name] = ConvertTo-PlainObject -Value $prop.Value -PropertyName $prop.Name
        }
        return ,$h
    }
    return $null
}

function ConvertTo-AdoJson {
    param($Value)
    # ConvertTo-Json understands PowerShell objects. JavaScriptSerializer must not
    # see them: passing a PSObject into Serialize walks PSParameterizedProperty
    # and throws a circular-reference error. Repair runs entirely in C#.
    if (-not ("AdoPolicyJson" -as [type])) {
        Add-Type -AssemblyName System.Web.Extensions
        $asm = [System.Web.Script.Serialization.JavaScriptSerializer].Assembly.Location
        Add-Type -ReferencedAssemblies $asm -TypeDefinition @"
using System;
using System.Collections;
using System.Collections.Generic;
using System.Web.Script.Serialization;
public static class AdoPolicyJson {
    static readonly string[] ForceArray = new string[] { "scope", "requiredReviewerIds", "optionalReviewerIds", "filenamePatterns" };
    public static string ToJson(string json) {
        var ser = new JavaScriptSerializer();
        ser.MaxJsonLength = 67108864;
        ser.RecursionLimit = 128;
        object graph = ser.DeserializeObject(json);
        graph = Repair(graph);
        return ser.Serialize(graph);
    }
    static object Repair(object node) {
        var dict = node as IDictionary<string, object>;
        if (dict != null) {
            var keys = new List<string>(dict.Keys);
            foreach (var key in keys) {
                object child = Repair(dict[key]);
                if (child != null && !(child is IList) && Array.IndexOf(ForceArray, key) >= 0) {
                    var list = new ArrayList();
                    list.Add(child);
                    child = list;
                }
                dict[key] = child;
            }
            return dict;
        }
        var listNode = node as IList;
        if (listNode != null && !(node is string)) {
            for (int i = 0; i < listNode.Count; i++) {
                listNode[i] = Repair(listNode[i]);
            }
            return listNode;
        }
        return node;
    }
}
"@
    }
    $json = $Value | ConvertTo-Json -Depth 100 -Compress
    return [AdoPolicyJson]::ToJson([string]$json)
}

function Get-RestErrorBody {
    param($ErrorRecord)
    if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
        return [string]$ErrorRecord.ErrorDetails.Message
    }
    try {
        $resp = $ErrorRecord.Exception.Response
        if (-not $resp) { return $null }
        $stream = $resp.GetResponseStream()
        if (-not $stream) { return $null }
        $reader = New-Object System.IO.StreamReader($stream)
        $text = $reader.ReadToEnd()
        $reader.Dispose()
        return $text
    } catch { return $null }
}

function Get-ScopeRepositoryId {
    param($Scope)
    if ($null -eq $Scope) { return "" }
    $raw = $Scope.repositoryId
    if ($null -eq $raw) { return "" }
    if ($raw -is [psobject] -and $raw.PSObject.Properties.Name -contains "id") { $raw = $raw.id }
    $text = ([string]$raw).Trim().Trim("{}")
    if ($text -match "^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$") {
        return $text.ToLower()
    }
    return ""
}

function Test-RealRepositoryId {
    param($Id)
    return -not [string]::IsNullOrWhiteSpace((Get-ScopeRepositoryId ([pscustomobject]@{ repositoryId = $Id })))
}

function New-PolicyUpdateBody {
    param(
        $Policy,
        [switch]$SetEnabled,
        [bool]$Enabled = $true
    )
    $body = $Policy | ConvertTo-Json -Depth 30 | ConvertFrom-Json
    if ($SetEnabled) { $body.isEnabled = [bool]$Enabled }
    foreach ($prop in @("createdBy","createdDate","_links","isDeleted","isEnterpriseManaged","revision","id","url")) {
        $body.PSObject.Properties.Remove($prop)
    }
    if ($body.type) {
        $typeId = [string]$body.type.id
        $body.PSObject.Properties.Remove("type")
        $body | Add-Member -NotePropertyName type -NotePropertyValue ([pscustomobject]@{ id = $typeId }) -Force
    }
    if ($body.settings -and ($body.settings.PSObject.Properties.Name -contains "scope") -and $null -ne $body.settings.scope) {
        $body.settings.scope = @($body.settings.scope)
    }
    return $body
}

function Invoke-AdoRest {
    param(
        [string]$Uri,
        [string]$Method = "Get",
        [object]$Body = $null,
        [switch]$Quiet
    )
    $params = @{
        Uri         = $Uri
        Method      = $Method
        TimeoutSec  = 20
        ErrorAction = "Stop"
    }
    if ($script:AuthMode -eq "Windows") {
        $params.UseDefaultCredentials = $true
    } else {
        $params.Headers = $script:Headers
    }
    if ($null -ne $Body) {
        $params.Body = ConvertTo-AdoJson -Value $Body
        $params.ContentType = "application/json"
    }
    try {
        return Invoke-RestMethod @params
    } catch {
        $detail = Get-ErrorDetail $_
        $isTls = $detail -match "trust relationship|SSL/TLS|secure channel|unexpected error occurred on a send"
        if ($isTls -and -not $script:SslCallbackReady) {
            $script:SslRetryUsed = $true
            Write-Log "HTTPS handshake failed. Retrying with certificate checks disabled."
            Enable-SkipSslValidation
            return Invoke-AdoRest -Uri $Uri -Method $Method -Body $Body -Quiet:$Quiet
        }
        if (-not $Quiet) {
            $status = $null
            if ($_.Exception.Response) {
                try { $status = [int]$_.Exception.Response.StatusCode } catch {}
            }
            Write-Log "REST $Method failed: $Uri"
            if ($status) { Write-Log "HTTP Status: $status" }
            Write-Log "  $detail"
            $errBody = Get-RestErrorBody $_
            if ($errBody) {
                $short = $errBody.Trim()
                if ($short.Length -gt 500) { $short = $short.Substring(0, 500) }
                Write-Log "  Server said: $short"
            }
            if ($status -eq 401 -and $script:AuthMode -eq "PAT") {
                Write-Log "Check the PAT (not expired, Code Read or Project and Team Read) and that Basic authentication is enabled in IIS."
            }
            if ($isTls) {
                Write-Log "Close this window and start the script again so the old certificate callback is cleared."
            }
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

# Policy types that Azure DevOps can apply to all repositories (project) or one repository.
# Anything else returned by /policy/types (TFVC check-in policies, and so on) is left out.
$script:ApplicablePolicyTypeIds = @{
    "fa4e907d-c16b-4a4c-9dfa-4906e5d171dd" = "Minimum number of reviewers"
    "fd2167ab-b0be-447a-8ec8-39368250530e" = "Required reviewers"
    "c6a1889d-b943-4856-b76f-9e46bb6b0df2" = "Comment requirements"
    "40e92b44-2fe1-4dd6-b3d8-74a9c21d0c6e" = "Work item linking"
    "0609b952-1397-4640-95ec-e00a01b2c241" = "Build validation"
    "fa4e907d-c16b-4a4c-9dfa-4916e5d171ab" = "Require a merge strategy"
    "cbdc66da-9728-4af8-aada-9a5a32e4a226" = "Status checks"
    "2e26e725-8201-4edd-8bf5-978563c34a80" = "Maximum file size"
    "7ed39669-655c-494e-b4a0-a08b4da0fcce" = "Case enforcement"
}

function Test-ApplicableGitPolicy {
    param($Policy)
    if ($null -eq $Policy -or $Policy.isDeleted) { return $false }
    if (-not $Policy.settings -or $null -eq $Policy.settings.scope) { return $false }
    $scopes = @($Policy.settings.scope)
    if ($scopes.Count -eq 0) { return $false }
    foreach ($s in $scopes) {
        if ($null -eq $s) { continue }
        $names = @($s.PSObject.Properties.Name)
        if ($names -contains "repositoryId" -or $names -contains "refName" -or $names -contains "matchKind") {
            return $true
        }
    }
    return $false
}

function Get-PolicyConfigurationPages {
    param([string]$FirstUrl)
    $collected = @()
    $url = $FirstUrl
    for ($page = 0; $page -lt 20; $page++) {
        $params = @{
            Uri             = $url
            Method          = "Get"
            TimeoutSec      = 30
            UseBasicParsing = $true
            ErrorAction     = "Stop"
        }
        if ($script:AuthMode -eq "Windows") { $params.UseDefaultCredentials = $true }
        else { $params.Headers = $script:Headers }
        $resp = Invoke-WebRequest @params
        $json = $resp.Content | ConvertFrom-Json
        if ($json.value) { $collected += @($json.value) }
        $token = $null
        foreach ($key in @($resp.Headers.Keys)) {
            if ($key -eq "x-ms-continuationtoken") { $token = [string]$resp.Headers[$key] }
        }
        if ([string]::IsNullOrWhiteSpace($token)) { break }
        $join = if ($FirstUrl -match "\?") { "&" } else { "?" }
        $url = "$FirstUrl${join}continuationToken=$([uri]::EscapeDataString($token))"
    }
    return @($collected)
}

function Get-Policies {
    param([string]$ProjectName)
    Ensure-ApiVersion
    $enc = [uri]::EscapeDataString($ProjectName)
    $root = "$(Get-BaseUrl)/$($script:Collection)/$enc"
    $byId = @{}
    $urls = @(
        "$root/_apis/policy/configurations?api-version=$($script:ApiVersion)&`$top=200",
        "$root/_apis/git/policy/configurations?api-version=$($script:ApiVersion)&`$top=200"
    )
    try { $script:RepoListCache = @(Get-Repositories -ProjectName $ProjectName) } catch { $script:RepoListCache = @() }
    foreach ($repo in @($script:RepoListCache)) {
        if (-not $repo.id) { continue }
        $urls += "$root/_apis/git/policy/configurations?repositoryId=$($repo.id)&api-version=$($script:ApiVersion)&`$top=200"
    }
    foreach ($url in $urls) {
        try {
            foreach ($item in @(Get-PolicyConfigurationPages -FirstUrl $url)) {
                if ($null -eq $item -or $item.isDeleted -or $null -eq $item.id) { continue }
                if (-not (Test-ApplicableGitPolicy $item)) { continue }
                $byId[[string]$item.id] = $item
            }
        } catch {
            if ($url -notmatch "/git/policy/") { throw }
        }
    }
    return @($byId.Values)
}

function Update-PolicyTypeCombo {
    param([hashtable]$Map)
    $current = [string]$ObjBoxPolicyType.SelectedItem
    $ObjBoxPolicyType.Items.Clear()
    $names = @("All") + @($Map.Keys | Where-Object { $_ -ne "All" } | Sort-Object)
    foreach ($name in $names) {
        if ($Map.ContainsKey($name)) { [void]$ObjBoxPolicyType.Items.Add($name) }
    }
    if ($ObjBoxPolicyType.Items.Contains($current)) {
        $ObjBoxPolicyType.SelectedItem = $current
    } elseif ($ObjBoxPolicyType.Items.Count -gt 0) {
        $ObjBoxPolicyType.SelectedIndex = 0
    }
}

function Refresh-PolicyTypesFromServer {
    param([string]$ProjectName)
    if ($script:PolicyTypesLoadedFromServer) { return }
    if ([string]::IsNullOrWhiteSpace($ProjectName)) { return }
    if ([string]::IsNullOrWhiteSpace($script:ADOServerFQDN)) { return }
    try {
        Ensure-ApiVersion
        $enc = [uri]::EscapeDataString($ProjectName)
        $url = "$(Get-BaseUrl)/$($script:Collection)/$enc/_apis/policy/types?api-version=$($script:ApiVersion)"
        $resp = Invoke-AdoRest -Uri $url -Method Get -Quiet
        $policies = @(Get-Policies -ProjectName $ProjectName)
        $previousName = [string]$ObjBoxPolicyType.SelectedItem
        $previousId = ""
        if ($previousName -and $script:PolicyTypeMap.ContainsKey($previousName)) {
            $previousId = [string]$script:PolicyTypeMap[$previousName]
        }
        $map = @{ "All" = "" }
        $seen = @{}
        foreach ($t in @($resp.value)) {
            if (-not $t.id) { continue }
            $id = ([string]$t.id).ToLower()
            if (-not $script:ApplicablePolicyTypeIds.ContainsKey($id)) { continue }
            $name = if ($t.displayName) { [string]$t.displayName } else { [string]$script:ApplicablePolicyTypeIds[$id] }
            if ($map.ContainsKey($name)) { $name = "$name ($id)" }
            $map[$name] = $id
            $seen[$id] = $true
        }
        foreach ($p in $policies) {
            if (-not $p.type -or -not $p.type.id) { continue }
            $id = ([string]$p.type.id).ToLower()
            if ($seen.ContainsKey($id)) { continue }
            $name = if ($p.type.displayName) { [string]$p.type.displayName } else { $id }
            if ($map.ContainsKey($name)) { $name = "$name ($id)" }
            $map[$name] = $id
            $seen[$id] = $true
        }
        if ($map.Count -le 1) {
            foreach ($id in $script:ApplicablePolicyTypeIds.Keys) {
                $map[$script:ApplicablePolicyTypeIds[$id]] = $id
            }
        }
        if ($map.Count -gt 1) {
            $script:PolicyTypeMap = $map
            $script:PolicyTypesLoadedFromServer = $true
            Update-PolicyTypeCombo -Map $map
            if ($previousId) {
                foreach ($key in @($map.Keys)) {
                    if ($map[$key] -and ([string]$map[$key]).ToLower() -eq $previousId.ToLower()) {
                        $ObjBoxPolicyType.SelectedItem = $key
                        break
                    }
                }
            }
            Write-Log "Policy types that can be applied in '$ProjectName': $($map.Count - 1)."
        }
    } catch {
        Write-Log "Could not load policy types from the server. Using the built-in list."
        Write-Log "  $(Get-ErrorDetail $_)"
    }
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
$ObjBoxPolicyType.Size = New-Object System.Drawing.Size(520, 24)
$ObjBoxPolicyType.DropDownWidth = 520
$ObjBoxPolicyType.DropDownStyle = "DropDownList"
$script:PolicyTypeMap = @{ "All" = "" }
foreach ($id in $script:ApplicablePolicyTypeIds.Keys) {
    $script:PolicyTypeMap[$script:ApplicablePolicyTypeIds[$id]] = $id
}
$script:PolicyTypesLoadedFromServer = $false
[void]$ObjBoxPolicyType.Items.Add("All")
@($script:ApplicablePolicyTypeIds.Values | Sort-Object) | ForEach-Object { [void]$ObjBoxPolicyType.Items.Add($_) }
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
  "AuthMode": "PAT",
  "SkipSslValidation": true,
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
        $projects = @(Get-Projects | Where-Object { $_ -and $_.name } | Sort-Object name)
        foreach ($p in $projects) { [void]$ComboBoxProject.Items.Add($p.name) }
        $ComboBoxProject.Enabled = ($projects.Count -gt 0)
        $ButtonSelectRepos.Enabled = $ComboBoxProject.Enabled
        $ObjBoxBranch.Enabled = $ComboBoxProject.Enabled
        Write-Log "Successfully retrieved $($projects.Count) project(s) from the collection."
        if ($projects.Count -gt 0) {
            $ComboBoxProject.SelectedIndex = 0
            Write-Log "Selected project: $($ComboBoxProject.Text)"
        }
    } catch {
        $msg = Get-ErrorDetail $_
        if ($msg -match "401") {
            Write-Log "Authentication failed (401 Unauthorized) while loading projects."
            Write-Log "Check the PAT (not expired, Code Read or Project and Team Read) and that Basic authentication is enabled in IIS."
        } else {
            Write-Log "Failed to retrieve projects."
            Write-Log "  $msg"
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

$ComboBoxProject.Add_SelectedIndexChanged({
    if ([string]::IsNullOrWhiteSpace($ComboBoxProject.Text)) { return }
    if ($script:PolicyTypeProject -ne $ComboBoxProject.Text) {
        $script:PolicyTypesLoadedFromServer = $false
        $script:PolicyTypeProject = $ComboBoxProject.Text
        Refresh-PolicyTypesFromServer -ProjectName $ComboBoxProject.Text
    }
})

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
    $typeFilter = [string]$ObjBoxPolicyType.SelectedItem
    if ([string]::IsNullOrWhiteSpace($typeFilter)) { $typeFilter = "All" }
    if ($IsModify) {
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "$Action $Scope-level '$typeFilter' policies in '$proj'?`n`nA backup is saved before the first change.",
            "Confirm", "YesNo")
        if ($confirm -ne "Yes") { return }
    }

    try {
        $policies = Get-Policies -ProjectName $proj
        Refresh-PolicyTypesFromServer -ProjectName $proj
        $typeFilter = [string]$ObjBoxPolicyType.SelectedItem
        if ([string]::IsNullOrWhiteSpace($typeFilter)) { $typeFilter = "All" }
        $filterId = ""
        if ($typeFilter -ne "All") {
            if ($script:PolicyTypeMap.ContainsKey($typeFilter)) {
                $filterId = [string]$script:PolicyTypeMap[$typeFilter]
                Write-Log "Policy type filter: $typeFilter ($filterId)"
            } else {
                Write-Log "Policy type '$typeFilter' is not a known type. Showing every configured policy."
                $typeFilter = "All"
                Write-Log "Policy type filter: All"
            }
        } else {
            Write-Log "Policy type filter: All"
        }
        if ($Scope -eq "Repo") {
            if ($script:SelectedRepoNames.Count -gt 0) {
                Write-Log "Repository filter: $($script:SelectedRepoNames -join ', ')"
            } else {
                Write-Log "Repository filter: all repositories"
            }
        }

        $enc = [uri]::EscapeDataString($proj)
        $policies = @($policies | Where-Object { $_ -ne $null })
        try { $script:RepoListCache = @(Get-Repositories -ProjectName $proj) } catch {}
        $repoNames = @{}
        foreach ($cached in @($script:RepoListCache)) {
            $rid = Get-ScopeRepositoryId ([pscustomobject]@{ repositoryId = $cached.id })
            if ($rid) { $repoNames[$rid] = [string]$cached.name }
        }
        $selectedIds = @()
        foreach ($id in @($script:SelectedRepoIds)) {
            $rid = Get-ScopeRepositoryId ([pscustomobject]@{ repositoryId = $id })
            if ($rid) { $selectedIds += $rid }
        }
        $matched = 0
        $changed = 0
        $failed = 0
        $projectCount = 0
        $repoCount = 0
        $elsewhere = New-Object System.Collections.Generic.List[string]
        $configured = New-Object System.Collections.Generic.List[object]
        $backedUp = $false
        foreach ($p in $policies) {
            $scopes = @()
            if ($p.settings -and $p.settings.scope) { $scopes = @($p.settings.scope) }
            $isRepoLevel = $false
            $scopeRepoIds = @()
            foreach ($s in $scopes) {
                $rid = Get-ScopeRepositoryId $s
                if ($rid) {
                    $isRepoLevel = $true
                    $scopeRepoIds += $rid
                }
            }
            $typeId = ""
            if ($p.type -and $p.type.id) { $typeId = ([string]$p.type.id).ToLower() }
            $scopeBits = @()
            foreach ($s in $scopes) {
                $ref = if ($s.refName) { $s.refName } else { "all branches" }
                if ($rid = Get-ScopeRepositoryId $s) {
                    if (-not $repoNames.ContainsKey($rid)) {
                        try {
                            $repoUrl = "$(Get-BaseUrl)/$($script:Collection)/$enc/_apis/git/repositories/${rid}?api-version=$($script:ApiVersion)"
                            $one = Invoke-AdoRest -Uri $repoUrl -Quiet
                            if ($one.name) { $repoNames[$rid] = [string]$one.name }
                        } catch {}
                    }
                    $repo = if ($repoNames.ContainsKey($rid)) { $repoNames[$rid] } else { "repository $rid" }
                    $scopeBits += "$ref, applies to $repo"
                } else {
                    $scopeBits += "$ref, applies to all repositories"
                }
            }
            $typeName = if ($p.type.displayName) { $p.type.displayName } else { $typeId }
            $onText = if ($p.isEnabled) { "On" } else { "Off" }
            $blockText = if ($p.isBlocking) { "blocks completion" } else { "advisory only" }
            $scopeText = if ($scopeBits.Count -gt 0) { $scopeBits -join "; " } else { "whole project" }
            $summary = "  $typeName | Policy #$($p.id) | $onText | $blockText | $scopeText"
            [void]$configured.Add([pscustomobject]@{
                Summary = $summary
                RepoIds = @($scopeRepoIds)
            })
            if ($filterId -and $typeId -ne $filterId.ToLower()) { continue }
            $wrongScope = ($Scope -eq "Project" -and $isRepoLevel) -or ($Scope -eq "Repo" -and -not $isRepoLevel)
            if ($wrongScope) {
                [void]$elsewhere.Add($summary)
                continue
            }
            if ($Scope -eq "Repo" -and $selectedIds.Count -gt 0) {
                $applies = $false
                foreach ($rid in $scopeRepoIds) {
                    if ($selectedIds -contains $rid) { $applies = $true; break }
                }
                if (-not $applies) { continue }
            }

            $matched++
            $policyUrl = "$(Get-BaseUrl)/$($script:Collection)/$enc/_apis/policy/configurations/$($p.id)?api-version=$($script:ApiVersion)"

            switch ($Action) {
                "Query" {
                    Write-Log $summary
                }
                "Enable" {
                    try {
                        if (-not $backedUp) {
                            Backup-CurrentState -Operation "${Action}${Scope}Policy" -ProjectName $proj | Out-Null
                            $backedUp = $true
                        }
                        $body = New-PolicyUpdateBody -Policy $p -SetEnabled -Enabled $true
                        $updated = Invoke-AdoRest -Uri $policyUrl -Method Put -Body $body
                        $now = "saved"
                        if ($null -ne $updated -and $null -ne $updated.isEnabled) {
                            $now = if ($updated.isEnabled) { "On" } else { "Off" }
                        }
                        Write-Log "Turned on '$typeName' (policy #$($p.id)). Server now reports $now."
                        $changed++
                    } catch {
                        $failed++
                        Write-Log "Could not turn on '$typeName' (policy #$($p.id))."
                        $why = $_.Exception.InnerException
                        if (-not $why) { $why = $_.Exception }
                        Write-Log "  $($why.Message)"
                    }
                }
                "Disable" {
                    try {
                        if (-not $backedUp) {
                            Backup-CurrentState -Operation "${Action}${Scope}Policy" -ProjectName $proj | Out-Null
                            $backedUp = $true
                        }
                        $body = New-PolicyUpdateBody -Policy $p -SetEnabled -Enabled $false
                        $updated = Invoke-AdoRest -Uri $policyUrl -Method Put -Body $body
                        $now = "saved"
                        if ($null -ne $updated -and $null -ne $updated.isEnabled) {
                            $now = if ($updated.isEnabled) { "On" } else { "Off" }
                        }
                        Write-Log "Turned off '$typeName' (policy #$($p.id)). Server now reports $now."
                        $changed++
                    } catch {
                        $failed++
                        Write-Log "Could not turn off '$typeName' (policy #$($p.id))."
                        $why = $_.Exception.InnerException
                        if (-not $why) { $why = $_.Exception }
                        Write-Log "  $($why.Message)"
                    }
                }
                "Delete" {
                    try {
                        if (-not $backedUp) {
                            Backup-CurrentState -Operation "${Action}${Scope}Policy" -ProjectName $proj | Out-Null
                            $backedUp = $true
                        }
                        Invoke-AdoRest -Uri $policyUrl -Method Delete | Out-Null
                        Write-Log "Deleted '$typeName' (policy #$($p.id))."
                        $changed++
                    } catch {
                        $failed++
                        Write-Log "Could not delete '$typeName' (policy #$($p.id))."
                        $why = $_.Exception.InnerException
                        if (-not $why) { $why = $_.Exception }
                        Write-Log "  $($why.Message)"
                    }
                }
            }
        }
        if ($matched -eq 0) {
            Write-Log "No $Scope-level policies matched '$typeFilter'."
            if ($elsewhere.Count -gt 0) {
                $other = if ($Scope -eq "Project") { "repository" } else { "project" }
                Write-Log "This type is set at $other level ($($elsewhere.Count)):"
                foreach ($line in $elsewhere) { Write-Log $line }
                if ($Scope -eq "Project") {
                    Write-Log "Use the Repository-Level buttons to change these."
                } else {
                    Write-Log "Use the Project-Level buttons to change these."
                }
            } else {
                if ($selectedIds.Count -gt 0) {
                    $wanted = @($script:SelectedRepoNames) -join ", "
                    Write-Log "No '$typeFilter' policy applies to: $wanted"
                } else {
                    Write-Log "'$typeFilter' is not configured, so there is nothing to turn on or off."
                }
                $onSelection = @()
                $others = @()
                foreach ($row in $configured) {
                    $hit = $false
                    if ($selectedIds.Count -eq 0) { $hit = $true }
                    else {
                        foreach ($rid in @($row.RepoIds)) {
                            if ($selectedIds -contains $rid) { $hit = $true; break }
                        }
                    }
                    if ($hit) { $onSelection += $row.Summary } else { $others += $row.Summary }
                }
                if ($selectedIds.Count -gt 0 -and $onSelection.Count -gt 0) {
                    Write-Log "Policies on the selected repository:"
                    foreach ($line in $onSelection) { Write-Log $line }
                } elseif ($onSelection.Count -eq 0 -and $others.Count -eq 0) {
                    Write-Log "This project has no Git branch or repository policies."
                } else {
                    Write-Log "Policies configured in this project:"
                    foreach ($line in ($onSelection + $others)) { Write-Log $line }
                }
                Write-Log "No '$typeFilter' policy exists yet. The list above is everything this project has."
            }
        } elseif ($Action -eq "Query") {
            Write-Log "Query finished. $matched $Scope-level policy/policies matched '$typeFilter'."
        } else {
            Write-Log "$Action finished. Updated $changed, failed $failed, matched $matched."
        }
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
                    $body = New-PolicyUpdateBody -Policy $p
                    if (-not $body.settings) { $body | Add-Member -NotePropertyName settings -NotePropertyValue ([pscustomobject]@{}) -Force }
                    $body.settings | Add-Member -NotePropertyName creatorVoteCounts -NotePropertyValue $true -Force
                    Invoke-AdoRest -Uri $policyUrl -Method Put -Body $body | Out-Null
                    Write-Log "Policy #$($p.id): Self-approval is now allowed (PR creator vote counts toward the minimum)."
                }
                "Disable" {
                    $body = New-PolicyUpdateBody -Policy $p
                    if (-not $body.settings) { $body | Add-Member -NotePropertyName settings -NotePropertyValue ([pscustomobject]@{}) -Force }
                    $body.settings | Add-Member -NotePropertyName creatorVoteCounts -NotePropertyValue $false -Force
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
