param(
    [pscredential]$Credential,

    [string]$TargetWorkNo = "",

    [switch]$DryRun,

    [switch]$StopAfterProcessSpec,

    [switch]$SkipSvnUpdate
)

$ErrorActionPreference = "Stop"

$scriptRoot = $PSScriptRoot
if (-not $scriptRoot) {
    $scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
}

$mainScript = Join-Path $scriptRoot "Invoke-TrackerRdLifeTaskImport.ps1"

$params = @{
    SpecRoot           = Join-Path ([Environment]::GetFolderPath("Desktop")) "規格文件"
    ExportScript       = Join-Path $scriptRoot "Export-RDLifeSpecMht.ps1"
    DesignerId         = "cq1"
    Statuses           = @("W04", "W06", "W14", "W03")
    MaxFunctionPoints  = 12
    MaxCalculatedHours = 4
}

if ($Credential) {
    $params.Credential = $Credential
}
if ($TargetWorkNo) {
    $params.TargetWorkNo = $TargetWorkNo
}
if ($DryRun) {
    $params.DryRun = $true
}
if ($StopAfterProcessSpec) {
    $params.StopAfterProcessSpec = $true
}
if ($SkipSvnUpdate) {
    $params.SkipSvnUpdate = $true
}

& $mainScript @params

