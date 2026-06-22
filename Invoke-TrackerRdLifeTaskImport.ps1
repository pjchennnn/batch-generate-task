param(
    [string]$Account = "cq1pj",

    [pscredential]$Credential,

    [string]$SpecRoot = (Join-Path ([Environment]::GetFolderPath("Desktop")) "規格文件"),

    [string]$ExportScript = "",

    [string]$OutDir = "",

    [string]$TempDir = "",

    [string]$TortoiseProc = "C:\Program Files\TortoiseSVN\bin\TortoiseProc.exe",

    [string]$BaseUrl = "https://crd.winton.com.tw",

    [string]$DesignerId = "cq1",

    [string[]]$Statuses = @("W04", "W06", "W14", "W03"),

    [double]$MaxFunctionPoints = 12,

    [double]$MaxCalculatedHours = 4,

    [string]$TargetWorkNo = "",

    [int]$SelectorPerPage = 100,

    [int]$TaskListPerPage = 500,

    [switch]$DryRun,

    [switch]$StopAfterProcessSpec,

    [switch]$SkipSvnUpdate,

    [string]$CommitBranchSuffix = ".CQ1",

    [string]$ProjectRoot = ""
)

$ErrorActionPreference = "Stop"

$scriptRoot = $PSScriptRoot
if (-not $scriptRoot) {
    $scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
}
if (-not $ExportScript) {
    $ExportScript = Join-Path $scriptRoot "Export-RDLifeSpecMht.ps1"
}
if (-not $OutDir) {
    $OutDir = Join-Path $scriptRoot "work\rdlife-output"
}
if (-not $TempDir) {
    $TempDir = Join-Path $scriptRoot "work\upload-temp"
}

$workflowPrompt = @'
## 工作流程

**禁止git push到遠端**

以下會使用到MCP: coding-agent
1. 使用 fetch_task 抓下工單的內容
2. **用get_time記錄當前時間作為開始時間（記住這個時間戳）**
3. 若有附件，則用 **download_task_files** 抓下附件
3.1 若附件是word檔，請用 **word_to_html** 轉成 html 後，再閱讀(不要直接讀word)
3.2 請用  **convert_html_to_utf8** 處理剛剛轉換後的 html 檔案(輸出的檔案會有後綴`-utf8`)
3.3 請用 **minify_spec** 處理剛剛轉成 utf8 的 html 檔案(輸出的檔案會有後綴`-clean`)<br>**之後如果要讀取 word 文件，應該要讀取步驟3.3轉換後的 html** 
5. 切到 WorkingBranch 分支
6. git pull
7. 直接在本地切到工作分支 {CommitBranch}
8. 使用plan mode評估修改方式與步驟
9. 開始工作
10. 工作完成後，**請使用claude.md中的編譯指令，進行編譯**
11. `Code Review`: 請檢查是否有滿足工單需求
12. `Code Review`: 請查看相關的skill，是否有符合要求
13. 若有修改，再**使用claude.md中的編譯指令，進行編譯**
14. 編譯完成後，使用git add 針對git status中，屬於此次新增、修改、刪除的檔案 提交到本地 {CommitBranch} 分支。並且*不能*推到遠端，並且不要加入任何*tasks*底下的項目。
15. **用get_time取得當下時間，計算從步驟2到現在的總耗時秒數**
16. 分支處理完成後，使用 report_task_completed 回覆結果，**durationSeconds 參數填入步驟10計算的秒數**
17. 工作失敗，使用 report_task_failed 回報

## 時間記錄說明
- 在 fetch_task 成功後，記下當時的 Unix timestamp（秒）
- 在呼叫 report_task_completed 前，用當前時間減去開始時間得到 durationSeconds
- 將此值傳給 report_task_completed 的 durationSeconds 參數    

## 注意事項
**若是閱讀、異動指定範圍內的檔案、資料夾，不用提示**

**以下指令不要中斷：** `Search`, `Find`, `echo`
'@

function Invoke-TrackerJsonPost {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [object]$Body,

        [Parameter(Mandatory = $true)]
        [Microsoft.PowerShell.Commands.WebRequestSession]$Session
    )

    $json = $Body | ConvertTo-Json -Depth 30
    $response = Invoke-WebRequest `
        -UseBasicParsing `
        -Uri "$BaseUrl$Path" `
        -Method POST `
        -WebSession $Session `
        -Headers @{
            "X-Requested-With" = "XMLHttpRequest"
            "Accept" = "application/json, text/javascript, */*; q=0.01"
            "Origin" = $BaseUrl
            "Referer" = "$BaseUrl/"
        } `
        -ContentType "application/json; charset=UTF-8" `
        -Body ([System.Text.Encoding]::UTF8.GetBytes($json))

    return ($response.Content | ConvertFrom-Json)
}

function Get-TaskDescription {
    param($Item)

    if (-not [string]::IsNullOrWhiteSpace([string]$Item.reason)) {
        return [string]$Item.reason
    }

    return [string]$Item.description
}

function Get-RequirementText {
    param(
        [string]$Description,
        [string]$MatchedWorkDate
    )

    if (-not [string]::IsNullOrWhiteSpace($MatchedWorkDate)) {
        return "$Description`r`n`r`n只針對${MatchedWorkDate}的部分做改動，其他部分請忽略"
    }

    return $Description
}

function Get-ProductNoFromRdLifeSystem {
    param([string]$System)

    if ([string]::IsNullOrWhiteSpace($System)) {
        return "WERP"
    }

    switch ($System.Trim().ToUpperInvariant()) {
        "WEBPOS" { return "WERP" }
        "NBW"    { return "BW" }
        "WSP"    { return "WSP" }
        "CERP"   { return "CERP" }
        default  { return "WERP" }
    }
}

function Test-ProgramFolderExists {
    param(
        [string]$ProjectRoot,
        [string]$ProgramNo
    )

    if ([string]::IsNullOrWhiteSpace($ProjectRoot) -or [string]::IsNullOrWhiteSpace($ProgramNo)) {
        return $false
    }
    if (-not (Test-Path -LiteralPath $ProjectRoot)) {
        return $false
    }

    $target = $ProgramNo.Trim()
    $skip = @(".git", ".vs", ".svn", "bin", "obj", "packages", "node_modules")
    $stack = New-Object System.Collections.Generic.Stack[string]
    $stack.Push($ProjectRoot)

    while ($stack.Count -gt 0) {
        $dir = $stack.Pop()
        $subDirs = $null
        try {
            $subDirs = [System.IO.Directory]::GetDirectories($dir)
        } catch {
            continue
        }
        foreach ($sd in $subDirs) {
            $name = [System.IO.Path]::GetFileName($sd)
            if ($name -ieq $target) {
                return $true
            }
            if ($skip -contains $name.ToLowerInvariant()) {
                continue
            }
            $stack.Push($sd)
        }
    }
    return $false
}

function Test-IsEligibleWorkItem {
    param($Item)

    if ($Item.statusCode -ne "W04") {
        return $false
    }

    $hasSmallFunctionPoints = $false
    if ($null -ne $Item.functionPoints) {
        $hasSmallFunctionPoints = ([double]$Item.functionPoints -le $MaxFunctionPoints)
    }

    $hasSmallCalculatedHours = $false
    if ($null -ne $Item.calculatedHours) {
        $hasSmallCalculatedHours = ([double]$Item.calculatedHours -le $MaxCalculatedHours)
    }

    return ($hasSmallFunctionPoints -or $hasSmallCalculatedHours)
}

function Get-AllOpenTrackerTasks {
    param([Microsoft.PowerShell.Commands.WebRequestSession]$Session)

    $all = @()
    $page = 1
    $total = 1

    while ($all.Count -lt $total) {
        $result = Invoke-TrackerJsonPost `
            -Path "/Wct12/Wct1250/OnePageTask" `
            -Session $Session `
            -Body @{
                pagination = @{
                    page = $page
                    perPage = $TaskListPerPage
                    total = 0
                }
                filter = @{
                    isOnlyMyTasks = $true
                    isShowClosedTasks = $false
                }
            }

        if ($result.status -ne "200") {
            throw "OnePageTask failed: $($result | ConvertTo-Json -Depth 8)"
        }

        $total = [int]$result.data.total
        $rows = @($result.data.data)
        if ($rows.Count -eq 0) {
            break
        }

        $all += $rows
        $page++
    }

    return $all
}

function Get-ExistingRdLifeMap {
    param([object[]]$Tasks)

    $map = @{}
    foreach ($task in $Tasks) {
        $rdLifeItemNo = [string]$task.rdLifeItemNo
        if ([string]::IsNullOrWhiteSpace($rdLifeItemNo)) {
            continue
        }

        if (-not $map.ContainsKey($rdLifeItemNo)) {
            $map[$rdLifeItemNo] = @()
        }
        $map[$rdLifeItemNo] += [string]$task.taskId
    }

    return $map
}

function Copy-ToUploadTemp {
    param([System.IO.FileInfo]$SourceFile)

    New-Item -ItemType Directory -Path $TempDir -Force | Out-Null
    $targetPath = Join-Path $TempDir $SourceFile.Name
    Copy-Item -LiteralPath $SourceFile.FullName -Destination $targetPath -Force

    $tempFile = Get-Item -LiteralPath $targetPath
    $tempFile.Attributes = $tempFile.Attributes -band (-bnot [System.IO.FileAttributes]::ReadOnly)
    return (Get-Item -LiteralPath $targetPath)
}

function Upload-TrackerTempFile {
    param(
        [System.IO.FileInfo]$File,
        [Microsoft.PowerShell.Commands.WebRequestSession]$Session
    )

    $contentType = switch ($File.Extension.ToLowerInvariant()) {
        ".doc" { "application/msword" }
        ".docx" { "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }
        ".pdf" { "application/pdf" }
        default { "application/octet-stream" }
    }
    $boundary = "----WebKitFormBoundary" + ([guid]::NewGuid().ToString("N").Substring(0, 16))
    $utf8 = [System.Text.Encoding]::UTF8
    $prefix = "--$boundary`r`nContent-Disposition: form-data; name=`"file`"; filename=`"$($File.Name)`"`r`nContent-Type: $contentType`r`n`r`n"
    $suffix = "`r`n--$boundary--`r`n"
    $fileBytes = [System.IO.File]::ReadAllBytes($File.FullName)
    $prefixBytes = $utf8.GetBytes($prefix)
    $suffixBytes = $utf8.GetBytes($suffix)
    $bodyBytes = New-Object byte[] ($prefixBytes.Length + $fileBytes.Length + $suffixBytes.Length)
    [Array]::Copy($prefixBytes, 0, $bodyBytes, 0, $prefixBytes.Length)
    [Array]::Copy($fileBytes, 0, $bodyBytes, $prefixBytes.Length, $fileBytes.Length)
    [Array]::Copy($suffixBytes, 0, $bodyBytes, $prefixBytes.Length + $fileBytes.Length, $suffixBytes.Length)

    $response = Invoke-WebRequest `
        -UseBasicParsing `
        -Uri "$BaseUrl/FileCenter/TempFile/UploadFile" `
        -Method POST `
        -WebSession $Session `
        -Headers @{
            "Accept" = "*/*"
            "Origin" = $BaseUrl
            "Referer" = "$BaseUrl/"
        } `
        -ContentType "multipart/form-data; boundary=$boundary" `
        -Body $bodyBytes

    $result = $response.Content | ConvertFrom-Json
    if ($result.status -ne "200") {
        throw "UploadFile failed: $($response.Content)"
    }

    return $result
}

function Process-TrackerTempSpecFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SystemFileName,

        [Parameter(Mandatory = $true)]
        [string]$FileName,

        [Microsoft.PowerShell.Commands.WebRequestSession]$Session
    )

    $result = Invoke-TrackerJsonPost `
        -Path "/Wct12/Wct1250/ProcessTempSpecFile" `
        -Session $Session `
        -Body @{
            systemFileName = $SystemFileName
            fileName = $FileName
        }

    if ($result.status -ne "200") {
        throw "ProcessTempSpecFile failed: $($result | ConvertTo-Json -Depth 8)"
    }

    return $result
}

function Get-WorkContentFromDescription {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ""
    }

    $lines = [regex]::Split($Text, '[\r\n]+')
    $inlinePattern = '(?:內容|内容)\s*[:：]\s*(.+)$'
    $labelPattern = '^(?:內容|内容)\s*[:：]\s*$'
    $nextSectionPattern = '^(?:規格|规格|文件|檔案|档案|附件|提報人|提报人|案號|案号|客代)\s*[:：]|^//'

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $t = $lines[$i].Trim()
        $m = [regex]::Match($t, $inlinePattern)
        if ($m.Success) {
            $c = ([regex]::Replace($m.Groups[1].Value, '\s*//.*$', '')).Trim()
            if ($c) {
                return $c
            }
        }
        if ($t -match $labelPattern) {
            for ($j = $i + 1; $j -lt [Math]::Min($lines.Count, $i + 8); $j++) {
                $n = $lines[$j].Trim()
                if (-not $n) {
                    continue
                }
                if ($n -match $nextSectionPattern) {
                    break
                }
                $n = ([regex]::Replace($n, '\s*//.*$', '')).Trim()
                if ($n) {
                    return $n
                }
            }
        }
    }
    return ""
}

function Invoke-ExportRdLifeSpec {
    param($Item)

    $programNo = ([string]$Item.progNo).Trim()
    $contentOverride = Get-WorkContentFromDescription ([string]$Item.description)
    $rows = @(& $ExportScript -WorkNo $Item.itemNo -ProgramNo $programNo -SpecRoot $SpecRoot -OutDir $OutDir -WorkContentOverride $contentOverride)
    $specRows = @($rows | Where-Object {
        $_ -isnot [string] -and
        $_.PSObject.Properties.Name -contains "SpecPath" -and
        $_.SpecPath
    })

    if ($specRows.Count -eq 0) {
        throw "Export-RDLifeSpecMht.ps1 did not return SpecPath. WorkNo=$($Item.itemNo); ProgramNo=$programNo"
    }

    foreach ($row in $specRows) {
        if (-not (Test-Path -LiteralPath ([string]$row.SpecPath))) {
            throw "SpecPath not found: $($row.SpecPath)"
        }
    }

    return $specRows
}

if (-not (Test-Path -LiteralPath $ExportScript)) {
    throw "Export script not found: $ExportScript"
}
if (-not (Test-Path -LiteralPath $SpecRoot)) {
    throw "Spec root not found: $SpecRoot"
}
if (-not $SkipSvnUpdate -and -not (Test-Path -LiteralPath $TortoiseProc)) {
    throw "TortoiseProc not found: $TortoiseProc"
}

New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
New-Item -ItemType Directory -Path $TempDir -Force | Out-Null

if (-not $Credential) {
    $Credential = Get-Credential -UserName $Account -Message "Tracker login"
}

if (-not $SkipSvnUpdate) {
    $svnProc = Start-Process `
        -FilePath $TortoiseProc `
        -ArgumentList @("/command:update", "/path:`"$SpecRoot`"", "/closeonend:2") `
        -Wait `
        -PassThru `
        -WindowStyle Hidden

    if ($svnProc.ExitCode -ne 0) {
        throw "TortoiseSVN update failed. ExitCode=$($svnProc.ExitCode)"
    }
}

$session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
$session.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36"

$loginBody = @{
    returnUrl = ""
    account = $Account
    password = $Credential.GetNetworkCredential().Password
    rememberMe = "false"
}

Invoke-WebRequest `
    -UseBasicParsing `
    -Uri "$BaseUrl/Account/Create" `
    -Method POST `
    -WebSession $session `
    -ContentType "application/x-www-form-urlencoded" `
    -Body $loginBody `
    -MaximumRedirection 5 | Out-Null

$form = Invoke-TrackerJsonPost -Path "/Wct12/Wct1250/ResourceTaskForm" -Body @{} -Session $session
if ($form.status -ne "200") {
    throw "ResourceTaskForm session check failed: $($form | ConvertTo-Json -Depth 8)"
}

$selector = Invoke-TrackerJsonPost `
    -Path "/Wct12/Wct1250/RdLifeWorkItemSelector" `
    -Session $session `
    -Body @{
        page = 1
        perPage = $SelectorPerPage
        total = 0
        q = $TargetWorkNo
        searchType = "fuzzy"
        searchFields = @("W.WKIM002", "W.WKIM006", "PGIF004", "W.WKIM014", "W.WKIM016")
        searchableFields = @("W.WKIM002", "W.WKIM006", "W.WKIM016")
        sorts = @()
        selectorModel = @{}
        isPagable = $true
        openerModelField = "rdLifeItemNo"
        openerPkgClass = "TaskForm"
        statuses = $Statuses
        designers = @($DesignerId)
        systemIds = @()
        rdLifeSystemId = $null
    }

if ($selector.status -ne "200") {
    throw "RdLifeWorkItemSelector failed: $($selector | ConvertTo-Json -Depth 8)"
}

$allItems = @($selector.data.data)
if ($TargetWorkNo) {
    $allItems = @($allItems | Where-Object { $_.itemNo -eq $TargetWorkNo })
}
$eligible = @($allItems | Where-Object { Test-IsEligibleWorkItem $_ })

$existingTasks = @(Get-AllOpenTrackerTasks -Session $session)
$existingByRdLife = Get-ExistingRdLifeMap -Tasks $existingTasks

$results = @()
foreach ($item in $eligible) {
    $itemResult = [ordered]@{
        itemNo = $item.itemNo
        progNo = ([string]$item.progNo).Trim()
        status = "$($item.statusCode) $($item.statusName)"
        functionPoints = $item.functionPoints
        calculatedHours = $item.calculatedHours
        rdLifeMinutes = if ($null -ne $item.calculatedHours) { [int]([double]$item.calculatedHours * 60) } else { 0 }
        action = ""
        error = ""
    }

    if ($existingByRdLife.ContainsKey([string]$item.itemNo)) {
        $itemResult.action = "skipped-existing"
        $itemResult.existingTaskIds = ($existingByRdLife[[string]$item.itemNo] -join ",")
        $results += [pscustomobject]$itemResult
        continue
    }

    try {
        $exports = @(Invoke-ExportRdLifeSpec -Item $item)
        $primaryExport = $exports[0]
        $sourceFiles = @($exports | ForEach-Object { Get-Item -LiteralPath ([string]$_.SpecPath) })

        $matchedWorkDate = [string]$primaryExport.MatchedWorkDate
        $productNo = Get-ProductNoFromRdLifeSystem ([string]$primaryExport.System)
        $requirement = Get-RequirementText -Description ([string]$item.description) -MatchedWorkDate $matchedWorkDate
        $programNoForFolder = ([string]$item.progNo).Trim()
        if (Test-ProgramFolderExists -ProjectRoot $ProjectRoot -ProgramNo $programNoForFolder) {
            $requirement = $requirement + "`r`n`r`n## 相關檔案`r`n請優先參考" + $programNoForFolder + "的資料夾(忽略大小寫)"
        }
        $taskDescription = Get-TaskDescription -Item $item

        $itemResult.productNo = $productNo
        $itemResult.system = $primaryExport.System
        $itemResult.specPaths = @($exports | ForEach-Object { $_.SpecPath })
        $itemResult.specCandidates = @($exports | ForEach-Object { $_.SpecCandidate })
        $itemResult.matchedWorkDate = $primaryExport.MatchedWorkDate
        $itemResult.latestDemandDate = $primaryExport.LatestDemandDate
        $itemResult.requirement = $requirement

        if ($DryRun) {
            $itemResult.action = "dry-run"
            $itemResult.attachments = @($sourceFiles | ForEach-Object { $_.Name })
            $results += [pscustomobject]$itemResult
            continue
        }

        $uploadFiles = @()
        $processedSpecs = @()
        foreach ($sourceFile in $sourceFiles) {
            $tempFile = Copy-ToUploadTemp -SourceFile $sourceFile
            $upload = Upload-TrackerTempFile -File $tempFile -Session $session
            $ext = $tempFile.Extension
            $encodedName = [System.Web.HttpUtility]::UrlEncode($sourceFile.Name)
            $uploadSystemFileName = [string]$upload.data.systemFileName
            $uploadFileName = [string]$upload.data.fileName
            if (-not $uploadFileName) {
                $uploadFileName = $sourceFile.Name
            }
            $processSpec = Process-TrackerTempSpecFile -SystemFileName $uploadSystemFileName -FileName $uploadFileName -Session $session
            $processedSpecs += $processSpec
            $tempFileExportApi = "/FileCenter/TempFile/ExportFile?systemFileName=" + $uploadSystemFileName + "&fileName=" + $encodedName + "&ext=" + $ext
            $uploadFiles += @{
                isNewUpload = $true
                fileName = $sourceFile.Name
                ext = $ext
                size = $sourceFile.Length
                key = ([string](Get-Random -Minimum 100000 -Maximum 999999))
                systemFileName = $uploadSystemFileName
                uploadStatus = "upload success"
                progress = 0
                uploadFailedMessage = ""
                downloadable = $true
                openable = $true
                isRemove = $false
                tempFileExportApi = $tempFileExportApi
            }
        }
        $workingBranch = [string]$item.patch
        $commitBranch = $workingBranch + $CommitBranchSuffix

        if ($StopAfterProcessSpec) {
            $itemResult.action = "processed-spec"
            $itemResult.processSpecStatus = (@($processedSpecs | ForEach-Object { $_.status }) -join ",")
            $itemResult.uploadSystemFileNames = @($uploadFiles | ForEach-Object { $_.systemFileName })
            $itemResult.uploadFileNames = @($uploadFiles | ForEach-Object { $_.fileName })
            $itemResult.attachments = @($uploadFiles | ForEach-Object { $_.fileName })
            $results += [pscustomobject]$itemResult
            continue
        }

        $save = Invoke-TrackerJsonPost `
            -Path "/Wct12/Wct1250/SaveTask" `
            -Session $session `
            -Body @{
                task = @{
                    statusText = "待領取"
                    resultRatingText = ""
                    hoursDisplayText = ""
                    hoursCompositeText = ""
                    taskId = ""
                    productNo = $productNo
                    description = $taskDescription
                    requirement = $requirement
                    status = 0
                    workingBranch = $workingBranch
                    commitBranch = $commitBranch
                    workflowPrompt = $workflowPrompt
                    isRunTest = $false
                    testHost = ""
                    testCustId = ""
                    testCompId = ""
                    testAccount = ""
                    testStore = ""
                    testCulture = ""
                    rdLifeSourceType = "NN"
                    rdLifeItemNo = $item.itemNo
                    rdLifeHours = $itemResult.rdLifeMinutes
                    designerId = $item.designerId
                    uploadFiles = @($uploadFiles)
                    scoreFiles = @()
                }
            }

        $itemResult.action = "created"
        $itemResult.createdTaskId = $save.data.task.taskId
        $itemResult.workingBranch = $workingBranch
        $itemResult.commitBranch = $commitBranch
        $itemResult.saveStatus = $save.status
        $itemResult.processSpecStatus = (@($processedSpecs | ForEach-Object { $_.status }) -join ",")
        $itemResult.attachments = @($save.data.task.uploadFiles | ForEach-Object { $_.fileName })
    } catch {
        $itemResult.action = "failed"
        $itemResult.error = $_.Exception.Message
    }

    $results += [pscustomobject]$itemResult
}

[ordered]@{
    selectorStatus = $selector.status
    selectorTotal = $selector.data.total
    returnedCount = $allItems.Count
    eligibleCount = $eligible.Count
    existingOpenTaskCount = $existingTasks.Count
    dryRun = [bool]$DryRun
    stopAfterProcessSpec = [bool]$StopAfterProcessSpec
    createdCount = @($results | Where-Object { $_.action -eq "created" }).Count
    skippedExistingCount = @($results | Where-Object { $_.action -eq "skipped-existing" }).Count
    dryRunCount = @($results | Where-Object { $_.action -eq "dry-run" }).Count
    processedSpecCount = @($results | Where-Object { $_.action -eq "processed-spec" }).Count
    failedCount = @($results | Where-Object { $_.action -eq "failed" }).Count
    results = $results
} | ConvertTo-Json -Depth 12
