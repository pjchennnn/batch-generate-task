param(
    [Parameter(Position = 0)]
    [string]$WorkNo,

    [switch]$Pending,

    [string]$User = "cq1",

    [string]$Server = "10.1.1.217:52800",

    [string]$AuthHeader = $env:RDLIFE_AUTH_HEADER,

    [string]$PcapPath = "C:\Project\123.pcapng",

    [string]$OutDir = ""
)

$ErrorActionPreference = "Stop"

if (-not $OutDir) {
    $scriptRoot = $PSScriptRoot
    if (-not $scriptRoot) { $scriptRoot = (Get-Location).Path }
    $OutDir = Join-Path $scriptRoot "rdlife-output"
}

function Get-TSharkPath {
    $cmd = Get-Command tshark -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }

    $default = "C:\Program Files\Wireshark\tshark.exe"
    if (Test-Path $default) { return $default }

    return $null
}

function Get-AuthHeaderFromPcap {
    param([string]$Path)

    $tshark = Get-TSharkPath
    if (-not $tshark) {
        throw "tshark.exe was not found. Install Wireshark, pass -AuthHeader, or set env:RDLIFE_AUTH_HEADER."
    }
    if (-not (Test-Path $Path)) {
        throw "Pcap file was not found: $Path. Pass -AuthHeader or set env:RDLIFE_AUTH_HEADER."
    }

    $auth = & $tshark -r $Path -Y "http.authorization && http.request.uri contains `"QueryCommon`"" -T fields -e http.authorization |
        Where-Object { $_ -like "Basic *" } |
        Select-Object -First 1

    if (-not $auth) {
        throw "Could not find a QueryCommon Basic Authorization header in the pcap."
    }

    return $auth.Trim()
}

function Get-AuthHeaderFromLocalFile {
    $scriptRoot = $PSScriptRoot
    if (-not $scriptRoot) { $scriptRoot = (Get-Location).Path }
    $path = Join-Path $scriptRoot ".rdlife-auth.local"
    if (-not (Test-Path -LiteralPath $path)) { return $null }

    $text = (Get-Content -LiteralPath $path -Raw).Trim()
    if (-not $text) { return $null }

    if ($text -match '(Basic\s+[A-Za-z0-9+/=]+)') {
        return $matches[1].Trim()
    }

    throw "Invalid local auth file format: $path"
}

function Expand-FDataValues {
    param($Node, [System.Collections.Generic.List[string]]$Values)

    if ($null -eq $Node) { return }

    if ($Node -is [System.Array]) {
        foreach ($item in $Node) { Expand-FDataValues $item $Values }
        return
    }

    if ($Node -is [pscustomobject]) {
        if ($Node.PSObject.Properties.Name -contains "fields") {
            $fields = $Node.fields
            if ($fields -and ($fields.PSObject.Properties.Name -contains "FData") -and ($fields.FData -is [string])) {
                [void]$Values.Add($fields.FData)
            }
        }

        foreach ($prop in $Node.PSObject.Properties) {
            Expand-FDataValues $prop.Value $Values
        }
    }
}

function Convert-Base64Deflate {
    param([string]$Text)

    $clean = $Text -replace "`r|`n", ""
    $compressed = [Convert]::FromBase64String($clean)

    # DataSnap returns zlib-wrapped deflate. Windows PowerShell's
    # DeflateStream expects the raw deflate payload, so skip the 2-byte
    # zlib header and 4-byte Adler32 trailer when present.
    if ($compressed.Length -gt 6 -and $compressed[0] -eq 0x78) {
        $raw = New-Object byte[] ($compressed.Length - 6)
        [Array]::Copy($compressed, 2, $raw, 0, $raw.Length)
    } else {
        $raw = $compressed
    }

    $inputStream = [IO.MemoryStream]::new($raw)
    $deflate = [IO.Compression.DeflateStream]::new($inputStream, [IO.Compression.CompressionMode]::Decompress)
    $output = [IO.MemoryStream]::new()
    $deflate.CopyTo($output)
    $deflate.Dispose()
    $inputStream.Dispose()
    return $output.ToArray()
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

function Get-WorkNoMatches {
    param([byte[]]$Bytes)

    $text = [Text.Encoding]::GetEncoding("ISO-8859-1").GetString($Bytes)
    $matches = [regex]::Matches($text, "W[a-zA-Z0-9]{1,3}\d{12,13}")
    $seen = @{}
    $items = @()
    foreach ($match in $matches) {
        if (-not $seen.ContainsKey($match.Value)) {
            $seen[$match.Value] = $true
            $items += [pscustomobject]@{ WorkNo = $match.Value; Offset = $match.Index }
        }
    }
    return $items
}

function Get-RecordStrings {
    param([byte[]]$Bytes, [int]$Start, [int]$End)

    $strings = @()
    $strings += Get-AsciiStrings -Bytes $Bytes -Start $Start -End $End
    $strings += Get-Utf16Strings -Bytes $Bytes -Start $Start -End $End
    return $strings | Sort-Object Offset, Encoding
}

function Get-FirstMatch {
    param($Items, [string]$Pattern)
    $hit = $Items | Where-Object { $_.Text -match $Pattern } | Select-Object -First 1
    if ($hit) {
        $m = [regex]::Match($hit.Text, $Pattern)
        if ($m.Success -and $m.Groups.Count -gt 1) { return $m.Groups[1].Value }
        return $hit.Text
    }
    return $null
}

function Get-LastMatch {
    param($Items, [string]$Pattern)
    $hit = $Items | Where-Object { $_.Text -match $Pattern } | Select-Object -Last 1
    if ($hit) {
        $m = [regex]::Match($hit.Text, $Pattern)
        if ($m.Success -and $m.Groups.Count -gt 1) { return $m.Groups[1].Value }
        return $hit.Text
    }
    return $null
}

function Get-NoteLines {
    param($Items)

    $notePatterns = @(
        '\u898f\u683c|\u89c4\u683c|\u5167\u5bb9|\u5185\u5bb9|\u6587\u4ef6|\u6a94\u6848|\u5099\u8a3b|\u5907\u6ce8|\u8bf4\u660e|\u8aaa\u660e',
        '\u6e2c\u8a66|\u6d4b\u8bd5|\u64b0\u5beb|\u64b0\u5199|\u65b0\u589e|\u4fee\u6539|\u522a\u9664|\u53d6\u6d88|\u7d50\u6848|\u7ed3\u6848',
        '\u63a5\u53e3|\u532f\u5165|\u5bfc\u5165|\u5c0e\u5165|\u50b3\u7968|\u51ed\u8bc1|\u767c\u7968|\u53eb\u8ca8|\u5ba2\u88fd|\u5ba2\u5236',
        '\u63d0\u5831\u4eba|\u63d0\u62a5\u4eba|\u8a2d\u8a08\u5e2b|\u8bbe\u8ba1\u5e08|\u54c1\u7ba1'
    )
    $knownNoise = '^(ADBS|CERP|webPOS|NBW|WSP|WBW\d+|s\d+|cq\d*|te|fk\d+|\d{2}\.\d{2}|W03|W04|W05|W06|W14)$'

    $seen = @{}
    $lines = @()
    foreach ($item in $Items) {
        $text = $item.Text.Trim()
        if (-not $text -or $text -match $knownNoise) { continue }

        $isNote = $false
        foreach ($pattern in $notePatterns) {
            if ($text -match $pattern) {
                $isNote = $true
                break
            }
        }

        if ($isNote -and -not $seen.ContainsKey($text)) {
            $seen[$text] = $true
            $lines += [pscustomobject]@{
                Offset   = $item.Offset
                Encoding = $item.Encoding
                Text     = $text
            }
        }
    }

    return $lines | Sort-Object Offset
}

if (-not $Pending -and -not $WorkNo) {
    throw "Pass a work item number, a search keyword, or use -Pending."
}

if (-not $AuthHeader) {
    $AuthHeader = Get-AuthHeaderFromLocalFile
}

if (-not $AuthHeader) {
    $AuthHeader = Get-AuthHeaderFromPcap -Path $PcapPath
}

if ($AuthHeader -notlike "Basic *") {
    throw "AuthHeader must use the 'Basic ...' format."
}

New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
$programNoLabel = -join ([char[]](0x7a0b, 0x5f0f, 0x4ee3, 0x865f))

if ($Pending) {
    $uri = "http://$Server/datasnap/rest/TServerMethods1/QuerySelfWorkItem/$User/'W03','W04','W06','W14','W05'/24,28,13,17,19//2/"
    $filePrefix = "pending"
} else {
    $encodedWorkNo = [Uri]::EscapeDataString($WorkNo)
    $uri = "http://$Server/datasnap/rest/TServerMethods1/QueryCommon///$encodedWorkNo/2/"
    $filePrefix = $WorkNo
}
$headers = @{ Authorization = $AuthHeader }

$response = Invoke-WebRequest -UseBasicParsing -Uri $uri -Headers $headers -TimeoutSec 20
$jsonPath = Join-Path $OutDir "$filePrefix.response.json"
$response.Content | Set-Content -Encoding UTF8 $jsonPath

$jsonText = (Get-Content $jsonPath -Raw) -replace "^\uFEFF", ""
$json = $jsonText | ConvertFrom-Json
$values = [System.Collections.Generic.List[string]]::new()
Expand-FDataValues $json $values

$datasets = @{}
for ($i = 0; $i -lt $values.Count - 1; $i += 2) {
    $name = $values[$i]
    $data = $values[$i + 1]
    if ($data -like "eJ*") {
        $bytes = Convert-Base64Deflate $data
        $binPath = Join-Path $OutDir "$filePrefix.$name.bin"
        [IO.File]::WriteAllBytes($binPath, $bytes)
        $datasets[$name] = [pscustomobject]@{ Path = $binPath; Bytes = $bytes }
    }
}

if ($Pending) {
    $pendingSet = $datasets["tqSelfWorkItem"]
    if (-not $pendingSet) {
        throw "The response does not contain a tqSelfWorkItem dataset."
    }

    $matches = Get-WorkNoMatches -Bytes $pendingSet.Bytes
    $rows = @()
    for ($i = 0; $i -lt $matches.Count; $i++) {
        $recordStart = [Math]::Max(0, $matches[$i].Offset - 80)
        $recordEnd = if ($i + 1 -lt $matches.Count) { [Math]::Min($pendingSet.Bytes.Length, $matches[$i + 1].Offset - 1) } else { [Math]::Min($pendingSet.Bytes.Length, $matches[$i].Offset + 1200) }
        $recordStrings = Get-RecordStrings -Bytes $pendingSet.Bytes -Start $recordStart -End $recordEnd
        $after = $recordStrings | Where-Object { $_.Offset -ge $matches[$i].Offset }
        $afterWorkNo = $recordStrings | Where-Object { $_.Offset -gt $matches[$i].Offset }

        $status = Get-FirstMatch $after '^(W03|W04|W05|W06|W14)$'
        if ($status -ne "W04") { continue }

        $program = Get-FirstMatch $afterWorkNo '^([A-Z][A-Z0-9_]{2,}\d{1,5})$'
        $analyze = Get-FirstMatch $afterWorkNo '^(s\d+)$'
        $system = Get-FirstMatch $afterWorkNo '^(webPOS|CERP|NBW|WSP)'
        $version = Get-LastMatch $afterWorkNo '^(\d{2}\.\d{2})'
        $listFinal = Get-FirstMatch $afterWorkNo '^(S\d+\.\d+\.[A-Z0-9]+)'

        $row = [pscustomobject]@{
            WorkNo      = $matches[$i].WorkNo
            ProgramNo   = $program
            AnalyzeCode = $analyze
            System      = $system
            Version     = $version
            ListFinal   = $listFinal
            Status      = ([char]0x5f85).ToString() + ([char]0x8655).ToString() + ([char]0x7406).ToString()
        }
        $row | Add-Member -NotePropertyName $programNoLabel -NotePropertyValue $program -Force
        $rows += $row
    }

    [pscustomobject]@{
        Mode       = "Pending"
        Api        = $uri
        HttpStatus = $response.StatusCode
        Count      = $rows.Count
        OutputDir  = (Resolve-Path $OutDir).Path
    }

    "`n--- pending work items ---"
    $rows | Format-Table -AutoSize
    return
}

$work = $datasets["SysVerWorkItem"]
if (-not $work) {
    throw "The response does not contain a SysVerWorkItem dataset."
}

$needle = [Text.Encoding]::ASCII.GetBytes($WorkNo)
$offset = Find-Bytes -Bytes $work.Bytes -Needle $needle
if ($offset -lt 0) {
    $matches = Get-WorkNoMatches -Bytes $work.Bytes
    if ($matches.Count -eq 0) {
        throw "No work items were found in SysVerWorkItem dataset for query: $WorkNo"
    }

    [pscustomobject]@{
        Query      = $WorkNo
        Api        = $uri
        HttpStatus = $response.StatusCode
        Count      = $matches.Count
        OutputDir  = (Resolve-Path $OutDir).Path
    }

    "`n--- matched work items ---"
    $matches | Select-Object WorkNo, Offset | Format-Table -AutoSize
    return
}

$start = [Math]::Max(0, $offset - 64)
$end = [Math]::Min($work.Bytes.Length, $offset + 5000)
$strings = @()
$strings += Get-AsciiStrings -Bytes $work.Bytes -Start $start -End $end
$strings += Get-Utf16Strings -Bytes $work.Bytes -Start $start -End $end
$strings = $strings | Sort-Object Offset, Encoding
$noteLines = Get-NoteLines $strings
$programNo = Get-FirstMatch $strings '^WBW\d+'

$summary = [pscustomobject]@{
    WorkNo       = $WorkNo
    Api          = $uri
    HttpStatus   = $response.StatusCode
    ProgramNo    = $programNo
    AnalyzeCode  = Get-FirstMatch $strings '^s\d+'
    System       = Get-FirstMatch $strings '^(webPOS|CERP|NBW|WSP)'
    Version      = Get-LastMatch $strings '^(\d{2}\.\d{2})'
    ListFinal    = Get-FirstMatch $strings '^(S\d+\.\d+\.[A-Z]\d+)'
    ProgramName  = Get-FirstMatch $strings '(API\u63a5\u53e3)'
    Status       = Get-FirstMatch $strings '(\u5f85\u8655\u7406|\u64b0\u5beb\u4e2d|\u64b0\u5beb\u5b8c\u6210|\u9000\u56de)'
    AnalystName  = Get-FirstMatch $strings '([\u5f90\u694a\u9673\u8449][\u4e00-\u9fff]{1,3})'
    OutputDir    = (Resolve-Path $OutDir).Path
}
$summary | Add-Member -NotePropertyName $programNoLabel -NotePropertyValue $programNo -Force

$summary

"`n--- description / notes ---"
if ($noteLines.Count -gt 0) {
    $noteLines | Select-Object Offset, Text | Format-Table -Wrap -AutoSize
} else {
    "(none found in SysVerWorkItem)"
}

"`n--- strings near record ---"
$strings |
    Where-Object {
        $_.Text -match $WorkNo -or
        $_.Text -match '^WBW|^s\d+|^cq|^te$|webPOS|^\d{2}\.\d{2}|^S\d+\.|API|\u5f85\u8655\u7406|\u64b0\u5beb|\u9000\u56de|\u5f90|\u694a|\u9673|\u8449'
    } |
    Select-Object Offset, Encoding, Text |
    Format-Table -AutoSize
