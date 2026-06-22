[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [ValidateSet('run', 'status', 'config', 'setcred')]
    [string]$Action = 'run',

    [string]$EmployeeNo,
    [string]$SpecRoot,
    [Nullable[double]]$MaxFunctionPoints,
    [Nullable[double]]$MaxCalculatedHours,
    [string]$AiOrderRoot,
    [string]$ProjectRoot,

    [string]$TargetWorkNo,
    [switch]$DryRun,
    [switch]$SkipSvnUpdate
)

$ErrorActionPreference = 'Stop'

$ConfigDir   = Join-Path $env:USERPROFILE '.winton-ai-order'
$ConfigFile  = Join-Path $ConfigDir 'config.json'
$CredFile    = Join-Path $ConfigDir 'cred.xml'
$LastRunFile = Join-Path $ConfigDir 'last-run.json'

function Ensure-ConfigDir {
    if (-not (Test-Path -LiteralPath $ConfigDir)) {
        New-Item -ItemType Directory -Path $ConfigDir -Force | Out-Null
    }
}

function Read-Config {
    $defaults = [ordered]@{
        employeeNo         = $null
        specRoot           = $null
        maxFunctionPoints  = 12
        maxCalculatedHours = 4
        aiOrderRoot        = $PSScriptRoot
        projectRoot        = 'C:\Project\agent1'
    }
    if (Test-Path -LiteralPath $ConfigFile) {
        $loaded = Get-Content -LiteralPath $ConfigFile -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($k in @($defaults.Keys)) {
            if (($loaded.PSObject.Properties.Name -contains $k) -and ($null -ne $loaded.$k)) {
                $defaults[$k] = $loaded.$k
            }
        }
    }
    return [pscustomobject]$defaults
}

function Write-Config($cfg) {
    Ensure-ConfigDir
    ($cfg | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath $ConfigFile -Encoding UTF8
}

switch ($Action) {

    'status' {
        $cfg = Read-Config
        $credValid  = Test-Path -LiteralPath $CredFile
        $configured = [bool]$cfg.employeeNo -and [bool]$cfg.specRoot
        [pscustomobject]@{
            action             = 'status'
            configured         = $configured
            credValid          = $credValid
            employeeNo         = $cfg.employeeNo
            specRoot           = $cfg.specRoot
            specRootExists     = [bool]($cfg.specRoot -and (Test-Path -LiteralPath $cfg.specRoot))
            maxFunctionPoints  = $cfg.maxFunctionPoints
            maxCalculatedHours = $cfg.maxCalculatedHours
            aiOrderRoot        = $cfg.aiOrderRoot
            projectRoot        = $cfg.projectRoot
        } | ConvertTo-Json -Compress
    }

    'config' {
        $cfg = Read-Config
        if ($PSBoundParameters.ContainsKey('EmployeeNo'))         { $cfg.employeeNo = $EmployeeNo }
        if ($PSBoundParameters.ContainsKey('SpecRoot'))           { $cfg.specRoot = $SpecRoot }
        if ($PSBoundParameters.ContainsKey('MaxFunctionPoints'))  { $cfg.maxFunctionPoints = [double]$MaxFunctionPoints }
        if ($PSBoundParameters.ContainsKey('MaxCalculatedHours')) { $cfg.maxCalculatedHours = [double]$MaxCalculatedHours }
        if ($PSBoundParameters.ContainsKey('AiOrderRoot'))        { $cfg.aiOrderRoot = $AiOrderRoot }
        if ($PSBoundParameters.ContainsKey('ProjectRoot'))        { $cfg.projectRoot = $ProjectRoot }
        Write-Config $cfg
        [pscustomobject]@{
            action             = 'config'
            employeeNo         = $cfg.employeeNo
            specRoot           = $cfg.specRoot
            maxFunctionPoints  = $cfg.maxFunctionPoints
            maxCalculatedHours = $cfg.maxCalculatedHours
            aiOrderRoot        = $cfg.aiOrderRoot
            projectRoot        = $cfg.projectRoot
        } | ConvertTo-Json -Compress
    }

    'setcred' {
        Ensure-ConfigDir
        Write-Host '=== 設定 RDLife / Tracker 帳號密碼（DPAPI 加密儲存）===' -ForegroundColor Cyan
        Write-Host '加密後只有「目前 Windows 使用者 + 本機」可解，存於：' -ForegroundColor DarkGray
        Write-Host "  $CredFile" -ForegroundColor DarkGray
        $cred = Get-Credential -Message 'RDLife / Tracker 登入帳密（帳號如 cq1pj）'
        if (-not $cred) {
            Write-Host '已取消，未儲存。' -ForegroundColor Yellow
            return
        }
        $cred | Export-Clixml -LiteralPath $CredFile
        Write-Host "已加密儲存：$CredFile" -ForegroundColor Green
        Write-Host '此視窗可關閉。' -ForegroundColor DarkGray
    }

    'run' {
        $cfg = Read-Config

        if (-not $cfg.employeeNo -or -not $cfg.specRoot) {
            Write-Output 'STATUS=not-configured'
            Write-Output 'HINT=請先設定工號與規格書路徑：ai-order.ps1 config -EmployeeNo <工號> -SpecRoot <路徑>'
            return
        }
        if (-not (Test-Path -LiteralPath $CredFile)) {
            Write-Output 'STATUS=no-credential'
            Write-Output 'HINT=請先在彈出視窗設定帳密：ai-order.ps1 setcred'
            return
        }
        if (-not (Test-Path -LiteralPath $cfg.specRoot)) {
            Write-Output 'STATUS=spec-root-missing'
            Write-Output ("HINT=規格書路徑不存在：" + $cfg.specRoot)
            return
        }

        $cred         = Import-Clixml -LiteralPath $CredFile
        $designerId   = ([string]$cfg.employeeNo).ToLower()
        $branchSuffix = '.' + $cfg.employeeNo
        $aiRoot       = if ($cfg.aiOrderRoot) { $cfg.aiOrderRoot } else { $PSScriptRoot }
        $importScript = Join-Path $aiRoot 'Invoke-TrackerRdLifeTaskImport.ps1'
        $exportScript = Join-Path $aiRoot 'Export-RDLifeSpecMht.ps1'

        $params = @{
            Account            = $cred.UserName
            Credential         = $cred
            ExportScript       = $exportScript
            SpecRoot           = $cfg.specRoot
            DesignerId         = $designerId
            CommitBranchSuffix = $branchSuffix
            ProjectRoot        = $cfg.projectRoot
            MaxFunctionPoints  = [double]$cfg.maxFunctionPoints
            MaxCalculatedHours = [double]$cfg.maxCalculatedHours
        }
        if ($TargetWorkNo)  { $params.TargetWorkNo = $TargetWorkNo }
        if ($DryRun)        { $params.DryRun = $true }
        if ($SkipSvnUpdate) { $params.SkipSvnUpdate = $true }

        try {
            $out = & $importScript @params
        } catch {
            Write-Output 'STATUS=error'
            Write-Output ('ERROR=' + $_.Exception.Message)
            return
        }

        $jsonText = ($out | Out-String).Trim()
        Ensure-ConfigDir
        $jsonText | Set-Content -LiteralPath $LastRunFile -Encoding UTF8

        try {
            $parsed = $jsonText | ConvertFrom-Json
        } catch {
            Write-Output 'STATUS=parse-error'
            Write-Output ('FULL_RESULT_FILE=' + $LastRunFile)
            return
        }

        $created = @($parsed.results | Where-Object { $_.action -eq 'created' })
        Write-Output 'STATUS=ok'
        Write-Output ('COUNTS=created={0};skipped={1};failed={2}' -f $parsed.createdCount, $parsed.skippedExistingCount, $parsed.failedCount)
        Write-Output ('CREATED_TASKS=' + (@($created | ForEach-Object { "$($_.createdTaskId):$($_.commitBranch)" }) -join ','))
        if ($parsed.failedCount -gt 0) {
            foreach ($f in @($parsed.results | Where-Object { $_.action -eq 'failed' })) {
                Write-Output ('FAILED_ITEM={0}|{1}' -f $f.itemNo, $f.error)
            }
        }
        Write-Output ('FULL_RESULT_FILE=' + $LastRunFile)
    }
}
