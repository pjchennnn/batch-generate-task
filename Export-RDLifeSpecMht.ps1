param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$WorkNo,

    [string]$SpecRoot = "",

    [string]$Downloads = "",

    [string]$QueryScript = "",

    [string]$OutDir = "",

    [string]$ProgramNo = "",

    [switch]$SkipQuery,

    [switch]$ExportMht,

    [switch]$ShowQueryOutput,

    [string]$WorkContentOverride = ""
)

$ErrorActionPreference = "Stop"

function Join-Chars {
    param([int[]]$Codes)
    return -join ($Codes | ForEach-Object { [char]$_ })
}

function Get-AsciiStrings {
    param([byte[]]$Bytes, [int]$Start, [int]$End, [int]$MinLength = 2)

    $items = @()
    $i = $Start
    while ($i -lt $End) {
        if ($Bytes[$i] -ge 32 -and $Bytes[$i] -lt 127) {
            $s = New-Object System.Text.StringBuilder
            $offset = $i
            while ($i -lt $End -and $Bytes[$i] -ge 32 -and $Bytes[$i] -lt 127) {
                [void]$s.Append([char]$Bytes[$i])
                $i++
            }
            if ($s.Length -ge $MinLength) {
                $items += [pscustomobject]@{ Offset = $offset; Text = $s.ToString(); Encoding = "ascii" }
            }
        }
        $i++
    }
    return $items
}

function Get-Utf16Strings {
    param([byte[]]$Bytes, [int]$Start, [int]$End, [int]$MinLength = 2)

    $items = @()
    for ($alignment = 0; $alignment -lt 2; $alignment++) {
        $i = $Start
        if (($i % 2) -ne $alignment) { $i++ }
        while ($i + 1 -lt $End) {
            $code = [BitConverter]::ToUInt16($Bytes, $i)
            $ok = (($code -ge 0x20 -and $code -le 0x7e) -or ($code -ge 0x4e00 -and $code -le 0x9fff) -or ($code -ge 0xff00 -and $code -le 0xffef) -or ($code -ge 0x3000 -and $code -le 0x303f))
            if ($ok) {
                $s = New-Object System.Text.StringBuilder
                $offset = $i
                while ($i + 1 -lt $End) {
                    $code = [BitConverter]::ToUInt16($Bytes, $i)
                    $ok = (($code -ge 0x20 -and $code -le 0x7e) -or ($code -ge 0x4e00 -and $code -le 0x9fff) -or ($code -ge 0xff00 -and $code -le 0xffef) -or ($code -ge 0x3000 -and $code -le 0x303f))
                    if (-not $ok) { break }
                    [void]$s.Append([char]$code)
                    $i += 2
                }
                if ($s.Length -ge $MinLength) {
                    $items += [pscustomobject]@{ Offset = $offset; Text = $s.ToString(); Encoding = "utf16" }
                }
            }
            $i += 2
        }
    }
    return $items
}

function Find-Bytes {
    param([byte[]]$Bytes, [byte[]]$Needle)

    for ($i = 0; $i -le $Bytes.Length - $Needle.Length; $i++) {
        $found = $true
        for ($j = 0; $j -lt $Needle.Length; $j++) {
            if ($Bytes[$i + $j] -ne $Needle[$j]) {
                $found = $false
                break
            }
        }
        if ($found) { return $i }
    }
    return -1
}

function Normalize-Name {
    param([string]$Text)
    if (-not $Text) { return "" }
    $lower = $Text.ToLowerInvariant()
    return [regex]::Replace($lower, '[\s\u3000,，、:：;；"''“”\(\)\[\]{}._-]+', '')
}

function Get-FirstMatch {
    param($Items, [string]$Pattern)
    $hit = $Items | Where-Object { $_.Text -match $Pattern } | Select-Object -First 1
    if (-not $hit) { return $null }
    $m = [regex]::Match($hit.Text, $Pattern)
    if ($m.Success -and $m.Groups.Count -gt 1) { return $m.Groups[1].Value }
    return $hit.Text
}

function Clean-Text {
    param([string]$Text)
    if ($null -eq $Text) { return "" }
    return ([regex]::Replace($Text, '[\x00-\x08\x0B\x0C\x0E-\x1F]', '')).Trim()
}

function Normalize-DateSep {
    param([string]$Text)
    if ($null -eq $Text) { return "" }
    return ($Text -replace '[.\-]', '/')
}

function Get-WorkItemStrings {
    param([string]$BinPath, [string]$WorkNo)

    $bytes = [IO.File]::ReadAllBytes($BinPath)
    $needle = [Text.Encoding]::ASCII.GetBytes($WorkNo)
    $offset = Find-Bytes -Bytes $bytes -Needle $needle
    if ($offset -lt 0) {
        throw "Work item was not found in SysVerWorkItem dataset: $WorkNo"
    }

    $start = [Math]::Max(0, $offset - 64)
    $end = [Math]::Min($bytes.Length, $offset + 5000)
    $strings = @()
    $strings += Get-AsciiStrings -Bytes $bytes -Start $start -End $end
    $strings += Get-Utf16Strings -Bytes $bytes -Start $start -End $end
    return $strings | Sort-Object Offset, Encoding
}

function Get-SpecCandidatesFromStrings {
    param($Strings)

    $items = @()
    $specPattern = '(?:\u898f\u683c|\u89c4\u683c)\s*[:\uFF1A]\s*(.+?\.(?:doc|docx))'
    $docPattern = '([^\r\n\t]+?\.(?:doc|docx))'
    foreach ($item in $Strings) {
        $match = [regex]::Match($item.Text, $specPattern, [Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($match.Success) {
            $raw = $match.Groups[1].Value.Trim()
            $parts = [regex]::Split($raw, '[,\uFF0C\u3001;\uFF1B]')
            foreach ($part in $parts) {
                $name = $part.Trim()
                $docMatch = [regex]::Match($name, '(.+?\.(?:doc|docx))', [Text.RegularExpressions.RegexOptions]::IgnoreCase)
                if ($docMatch.Success) {
                    $items += [pscustomobject]@{
                        SourceLine = $item.Text
                        Candidate  = $docMatch.Groups[1].Value.Trim()
                    }
                }
            }

                $items += [pscustomobject]@{
                    SourceLine = $item.Text
                    Candidate  = $raw
                }
        }

        foreach ($docMatch in [regex]::Matches($item.Text, $docPattern, [Text.RegularExpressions.RegexOptions]::IgnoreCase)) {
            $candidate = $docMatch.Groups[1].Value.Trim()
            $candidate = [regex]::Replace($candidate, '^(?:\u898f\u683c|\u89c4\u683c)\s*[:\uFF1A]\s*', '')
            $items += [pscustomobject]@{
                SourceLine = $item.Text
                Candidate  = $candidate
            }
        }
    }

    $seen = @{}
    $unique = @()
    foreach ($item in $items) {
        $key = Normalize-Name $item.Candidate
        if ($key -and -not $seen.ContainsKey($key)) {
            $seen[$key] = $true
            $unique += $item
        }
    }
    return $unique
}

function Get-WorkContentFromStrings {
    param($Strings)

    $items = @($Strings | Sort-Object Offset, Encoding)
    $inlineContentPattern = '(?:\u5167\u5bb9|\u5185\u5bb9)\s*[:\uFF1A]\s*(.+)$'
    $contentLabelPattern = '^(?:\u5167\u5bb9|\u5185\u5bb9)\s*[:\uFF1A]\s*$'
    $nextSectionPattern = '^(?:\u898f\u683c|\u89c4\u683c|\u6587\u4ef6|\u6a94\u6848|\u6863\u6848|\u9644\u4ef6|\u63d0\u5831\u4eba|\u63d0\u62a5\u4eba)\s*[:\uFF1A]|^//'

    for ($i = 0; $i -lt $items.Count; $i++) {
        $text = Clean-Text $items[$i].Text
        $match = [regex]::Match($text, $inlineContentPattern)
        if ($match.Success) {
            $content = $match.Groups[1].Value.Trim()
            $content = [regex]::Replace($content, '\s*//.*$', '').Trim()
            if ($content) { return $content }
        }

        if ($text -match $contentLabelPattern) {
            for ($j = $i + 1; $j -lt [Math]::Min($items.Count, $i + 8); $j++) {
                $nextText = Clean-Text $items[$j].Text
                if (-not $nextText) { continue }
                if ($nextText -match $nextSectionPattern) { break }

                $nextText = [regex]::Replace($nextText, '\s*//.*$', '').Trim()
                if ($nextText) { return $nextText }
            }
        }
    }
    return ""
}

function Find-SpecDocuments {
    param(
        [string]$Root,
        $SpecCandidates,
        [string]$ProgramNo
    )

    if (-not (Test-Path -LiteralPath $Root)) {
        throw "Spec root was not found: $Root"
    }

    $docs = Get-ChildItem -LiteralPath $Root -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -match '^\.(doc|docx)$' -and $_.Name -notlike '~$*' }

    $found = @()
    $foundPaths = @{}

    foreach ($candidate in $SpecCandidates) {
        $leaf = [IO.Path]::GetFileName($candidate.Candidate)
        if (-not $leaf) { $leaf = $candidate.Candidate }
        $leafNorm = Normalize-Name $leaf
        $candidateNorm = Normalize-Name $candidate.Candidate

        $hits = @()
        $hits += $docs | Where-Object { $_.Name -eq $leaf } | ForEach-Object {
            [pscustomobject]@{ File = $_; Score = 100; Mode = "spec-exact"; Candidate = $candidate.Candidate; SourceLine = $candidate.SourceLine }
        }
        $hits += $docs | Where-Object { (Normalize-Name $_.Name) -eq $leafNorm -or (Normalize-Name $_.FullName) -like "*$candidateNorm*" } | ForEach-Object {
            [pscustomobject]@{ File = $_; Score = 90; Mode = "spec-normalized"; Candidate = $candidate.Candidate; SourceLine = $candidate.SourceLine }
        }

        $best = $hits | Sort-Object Score, @{ Expression = { $_.File.LastWriteTime }; Descending = $true } -Descending | Select-Object -First 1
        if ($best -and -not $foundPaths.ContainsKey($best.File.FullName)) {
            $foundPaths[$best.File.FullName] = $true
            $found += $best
        }
    }

    if ($found.Count -gt 0) { return $found }

    if ($ProgramNo) {
        $programHits = $docs | Where-Object { $_.Name -like "$ProgramNo*" -or $_.Name -like "*$ProgramNo*" } | ForEach-Object {
            $nameWithoutExt = [IO.Path]::GetFileNameWithoutExtension($_.Name)
            $escapedProgramNo = [regex]::Escape($ProgramNo)
            $score = if ($nameWithoutExt -eq $ProgramNo) {
                100
            } elseif ($_.Name -match "^$escapedProgramNo(\s|[\._-])") {
                95
            } elseif ($_.Name -like "$ProgramNo*") {
                70
            } else {
                60
            }
            [pscustomobject]@{ File = $_; Score = $score; Mode = "program-code"; Candidate = $ProgramNo; SourceLine = "" }
        }
        $bestProgram = $programHits | Sort-Object Score, @{ Expression = { $_.File.LastWriteTime }; Descending = $true } -Descending | Select-Object -First 1
        if ($bestProgram) { return @($bestProgram) }
    }

    return @()
}

function Convert-DocumentToMht {
    param(
        [string]$InputPath,
        [string]$OutputDir
    )

    $workDir = Join-Path $PSScriptRoot "rdlife-output\mht-work"
    New-Item -ItemType Directory -Path $workDir -Force | Out-Null

    $ext = [IO.Path]::GetExtension($InputPath)
    if (-not $ext) { $ext = ".doc" }
    $tempDoc = Join-Path $workDir ("input" + $ext)
    $tempMht = Join-Path $workDir "output.mht"
    $outputPath = Join-Path $OutputDir (([IO.Path]::GetFileNameWithoutExtension($InputPath)) + ".mht")

    Copy-Item -LiteralPath $InputPath -Destination $tempDoc -Force
    Remove-Item -LiteralPath $tempMht -Force -ErrorAction SilentlyContinue

    $worker = @'
$ErrorActionPreference = "Stop"
$docPath = "__DOC__"
$outPath = "__OUT__"
$word = $null
$doc = $null
try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $doc = $word.Documents.OpenNoRepairDialog($docPath, $false, $true, $false)
  $doc.SaveAs2($outPath, 9)
}
finally {
  if ($doc) { $doc.Close([ref]$false) | Out-Null }
  if ($word) { $word.Quit() | Out-Null }
}
'@

    $worker = $worker.Replace("__DOC__", $tempDoc.Replace("\", "\\")).Replace("__OUT__", $tempMht.Replace("\", "\\"))
    $workerPath = Join-Path $workDir "save-mht.ps1"
    Set-Content -LiteralPath $workerPath -Value $worker -Encoding ASCII

    powershell -NoProfile -Sta -ExecutionPolicy Bypass -File $workerPath
    if (-not (Test-Path -LiteralPath $tempMht)) {
        throw "MHT was not created."
    }

    Copy-Item -LiteralPath $tempMht -Destination $outputPath -Force
    return (Resolve-Path -LiteralPath $outputPath).Path
}

function Get-DocumentText {
    param([string]$InputPath)

    $workDir = Join-Path $PSScriptRoot "rdlife-output\mht-work"
    New-Item -ItemType Directory -Path $workDir -Force | Out-Null

    $ext = [IO.Path]::GetExtension($InputPath)
    if (-not $ext) { $ext = ".doc" }
    $tempDoc = Join-Path $workDir ("text-input" + $ext)
    $textPath = Join-Path $workDir "document-text.txt"

    Copy-Item -LiteralPath $InputPath -Destination $tempDoc -Force
    Remove-Item -LiteralPath $textPath -Force -ErrorAction SilentlyContinue

    $worker = @'
$ErrorActionPreference = "Stop"
$docPath = "__DOC__"
$textPath = "__TXT__"
$word = $null
$doc = $null
try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $doc = $word.Documents.OpenNoRepairDialog($docPath, $false, $true, $false)
  [IO.File]::WriteAllText($textPath, $doc.Content.Text, [Text.Encoding]::UTF8)
}
finally {
  if ($doc) { $doc.Close([ref]$false) | Out-Null }
  if ($word) { $word.Quit() | Out-Null }
}
'@

    $worker = $worker.Replace("__DOC__", $tempDoc.Replace("\", "\\")).Replace("__TXT__", $textPath.Replace("\", "\\"))
    $workerPath = Join-Path $workDir "extract-text.ps1"
    Set-Content -LiteralPath $workerPath -Value $worker -Encoding ASCII

    powershell -NoProfile -Sta -ExecutionPolicy Bypass -File $workerPath
    if (-not (Test-Path -LiteralPath $textPath)) {
        throw "Document text was not extracted."
    }

    return Get-Content -LiteralPath $textPath -Raw -Encoding UTF8
}

function Get-SpecDateInfo {
    param([string]$Text)

    if (-not $Text) { return @() }

    $datePattern = '\d{4}[/.\-]\d{1,2}[/.\-]\d{1,2}'
    $demandPattern = '\u9805\u76ee\u9700\u6c42|\u9879\u76ee\u9700\u6c42'
    $fallbackPattern = '\u8abf\u6574|\u8c03\u6574|\u9700\u6c42|\u4fee\u6539\u6a19\u8a18|\u4fee\u6539\u6807\u8bb0'
    $lines = [regex]::Split($Text, '[\r\n\v\f]+') | Where-Object { $_ -and $_.Trim() }

    $hits = @()
    foreach ($line in $lines) {
        if ($line -match $datePattern -and $line -match $demandPattern) {
            foreach ($m in [regex]::Matches($line, $datePattern)) {
                $hits += [pscustomobject]@{ Date = (Normalize-DateSep $m.Value); Strength = "demand"; Line = $line.Trim() }
            }
        }
    }

    if ($hits.Count -eq 0) {
        foreach ($line in $lines) {
            if ($line -match $datePattern -and $line -match $fallbackPattern) {
                foreach ($m in [regex]::Matches($line, $datePattern)) {
                    $hits += [pscustomobject]@{ Date = (Normalize-DateSep $m.Value); Strength = "fallback"; Line = $line.Trim() }
                }
            }
        }
    }

    $seen = @{}
    $unique = @()
    foreach ($hit in $hits) {
        $key = $hit.Date + "|" + $hit.Line
        if (-not $seen.ContainsKey($key)) {
            $seen[$key] = $true
            $unique += $hit
        }
    }
    return $unique
}

function Normalize-RequirementText {
    param([string]$Text)

    if (-not $Text) { return "" }
    $value = $Text.ToLowerInvariant()
    $value = $value -replace '\u5ba2\u88fd\u5316|\u5ba2\u5236\u5316', ([string][char]0x5ba2 + [string][char]0x88fd)
    $value = $value -replace '\u5f9e|\u4ece', ''
    $value = [regex]::Replace($value, '[\s\u3000,，、:：;；"''“”\(\)\[\]{}._\-/\\]+', '')
    return $value
}

function Get-MatchedWorkDateInfo {
    param(
        [string]$DocumentText,
        [string]$WorkContent
    )

    if (-not $DocumentText -or -not $WorkContent) { return $null }

    $target = Normalize-RequirementText $WorkContent
    if ($target.Length -lt 6) { return $null }

    $datePattern = '\d{4}[/.\-]\d{1,2}[/.\-]\d{1,2}'
    $lines = @([regex]::Split($DocumentText, '[\r\n\v\f]+') | ForEach-Object { Clean-Text $_ } | Where-Object { $_ })
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $windowLines = @()
        for ($j = $i; $j -lt [Math]::Min($lines.Count, $i + 4); $j++) {
            $windowLines += $lines[$j]
        }

        $window = ($windowLines -join " ")
        $windowNorm = Normalize-RequirementText $window
        if ($windowNorm -notlike "*$target*") { continue }

        $date = ""
        $dateLine = ""

        $windowDateHits = @()
        foreach ($line in $windowLines) {
            foreach ($match in [regex]::Matches($line, $datePattern)) {
                $dateNorm = Normalize-DateSep $match.Value
                $parsedDate = [datetime]::MinValue
                if ([datetime]::TryParseExact($dateNorm, "yyyy/M/d", [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$parsedDate) -or
                    [datetime]::TryParseExact($dateNorm, "yyyy/MM/dd", [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$parsedDate)) {
                    $windowDateHits += [pscustomobject]@{
                        DateValue = $parsedDate
                        DateText  = $dateNorm
                        Line      = $line
                    }
                }
            }
        }

        $bestWindowDate = $windowDateHits | Sort-Object DateValue -Descending | Select-Object -First 1
        if ($bestWindowDate) {
            $date = $bestWindowDate.DateText
            $dateLine = $bestWindowDate.Line
        } else {
            for ($k = $i; $k -ge [Math]::Max(0, $i - 5); $k--) {
                $match = [regex]::Match($lines[$k], $datePattern)
                if ($match.Success) {
                    $date = Normalize-DateSep $match.Value
                    $dateLine = $lines[$k]
                    break
                }
            }
        }

        if ($date) {
            return [pscustomobject]@{
                Date      = $date
                DateLine  = $dateLine
                MatchText = $window
            }
        }
    }

    return $null
}

function Get-LatestDateText {
    param([string[]]$DateTexts)

    $latest = $null
    $latestText = ""
    foreach ($text in $DateTexts) {
        $textNorm = Normalize-DateSep $text
        $date = [datetime]::MinValue
        if ([datetime]::TryParseExact($textNorm, "yyyy/M/d", [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$date) -or
            [datetime]::TryParseExact($textNorm, "yyyy/MM/dd", [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$date)) {
            if ($null -eq $latest -or $date -gt $latest) {
                $latest = $date
                $latestText = $date.ToString("yyyy/MM/dd")
            }
        }
    }
    return $latestText
}

if (-not $QueryScript) {
    $QueryScript = Join-Path $PSScriptRoot "query-rdlife.ps1"
}
if (-not $OutDir) {
    $OutDir = Join-Path $PSScriptRoot "rdlife-output"
}
if (-not $SpecRoot) {
    $SpecRoot = Join-Path ([Environment]::GetFolderPath("Desktop")) (Join-Chars @(0x898f, 0x683c, 0x6587, 0x4ef6))
}
if (-not $Downloads) {
    $Downloads = Join-Path $env:USERPROFILE "Downloads"
}

New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
if ($ExportMht) {
    New-Item -ItemType Directory -Path $Downloads -Force | Out-Null
}

if (-not $SkipQuery) {
    $queryOutput = & powershell -ExecutionPolicy Bypass -File $QueryScript $WorkNo -OutDir $OutDir
    if ($ShowQueryOutput) {
        $queryOutput | Out-Host
    }
}

$binPath = Join-Path $OutDir "$WorkNo.SysVerWorkItem.bin"
if (-not (Test-Path -LiteralPath $binPath)) {
    throw "SysVerWorkItem bin was not found: $binPath"
}

$strings = Get-WorkItemStrings -BinPath $binPath -WorkNo $WorkNo
if ($ProgramNo) {
    $programNo = $ProgramNo.Trim()
} else {
    $programNo = Get-FirstMatch $strings '^WBW\d+'
}
$workContent = Get-WorkContentFromStrings $strings
if ((Clean-Text $WorkContentOverride).Length -ge 6) {
    $workContent = (Clean-Text $WorkContentOverride)
}
$specCandidates = @(Get-SpecCandidatesFromStrings $strings)
$foundItems = @(Find-SpecDocuments -Root $SpecRoot -SpecCandidates $specCandidates -ProgramNo $programNo)

if ($foundItems.Count -eq 0) {
    throw "Spec document was not found. WorkNo=$WorkNo ProgramNo=$programNo"
}

foreach ($found in $foundItems) {
    $mhtPath = ""
    if ($ExportMht) {
        $mhtPath = Convert-DocumentToMht -InputPath $found.File.FullName -OutputDir $Downloads
    }
    $docText = Get-DocumentText -InputPath $found.File.FullName
    $dateInfo = @(Get-SpecDateInfo -Text $docText)
    $dateTexts = @($dateInfo | Select-Object -ExpandProperty Date -Unique)
    $latestDemandDate = Get-LatestDateText -DateTexts $dateTexts
    $matchedWorkDate = Get-MatchedWorkDateInfo -DocumentText $docText -WorkContent $workContent

    [pscustomobject]@{
        WorkNo        = $WorkNo
        ProgramNo     = $programNo
        WorkContent   = $workContent
        SearchMode    = $found.Mode
        SpecCandidate = $found.Candidate
        SpecLine      = $found.SourceLine
        SpecPath      = $found.File.FullName
        MhtPath       = $mhtPath
        MatchedWorkDate = if ($matchedWorkDate) { $matchedWorkDate.Date } else { "" }
        MatchedWorkDateLine = if ($matchedWorkDate) { $matchedWorkDate.DateLine } else { "" }
        MatchedWorkText = if ($matchedWorkDate) { $matchedWorkDate.MatchText } else { "" }
        DemandDates   = ($dateTexts -join ", ")
        LatestDemandDate = $latestDemandDate
        DemandDateLines = (($dateInfo | Select-Object -ExpandProperty Line -Unique) -join " | ")
    }
}
