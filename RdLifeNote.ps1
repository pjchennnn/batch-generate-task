. (Join-Path $PSScriptRoot "RdLifeMcp.ps1")

$RdLifeNoteUri = "http://rdvm-srv2012cr.winton.com.tw:52800/datasnap/rest/TServerMethods1/%22UpdWorkItemNote%22/"
$RdLifeNoteTemplateDefault = Join-Path $env:USERPROFILE ".winton-ai-order\templates\updworkitemnote-frame101-body.json"
$RdLifeNoteTemplatePk = [guid]"11111111-1111-1111-1111-111111111111"
$RdLifeNoteTemplateValue = [string]::Concat([char]0xE000, [char]0xE001, [char]0xE002)

function Get-RdLifeDataSnapAuth {
    param([string]$AuthHeader)

    if ($AuthHeader) { return $AuthHeader.Trim() }
    if ($env:RDLIFE_AUTH_HEADER) { return $env:RDLIFE_AUTH_HEADER.Trim() }

    $local = Join-Path $PSScriptRoot ".rdlife-auth.local"
    if (Test-Path -LiteralPath $local) {
        $text = (Get-Content -LiteralPath $local -Raw).Trim()
        if ($text -match "(Basic\s+[A-Za-z0-9+/=]+)") { return $matches[1].Trim() }
        return $text
    }

    throw "RDLife DataSnap auth not found. Provide -AuthHeader, RDLIFE_AUTH_HEADER, or .rdlife-auth.local."
}

function Get-RdLifeNoteAdler32 {
    param([byte[]]$Data)

    $mod = 65521
    [uint32]$a = 1
    [uint32]$b = 0
    foreach ($value in $Data) {
        $a = ($a + [uint32]$value) % $mod
        $b = ($b + $a) % $mod
    }
    return (($b -shl 16) -bor $a)
}

function Expand-RdLifeNoteZlib {
    param([byte[]]$Compressed)

    $raw = New-Object byte[] ($Compressed.Length - 6)
    [Array]::Copy($Compressed, 2, $raw, 0, $raw.Length)
    $input = [IO.MemoryStream]::new($raw)
    $deflate = [IO.Compression.DeflateStream]::new($input, [IO.Compression.CompressionMode]::Decompress)
    $output = [IO.MemoryStream]::new()
    $deflate.CopyTo($output)
    $deflate.Dispose()
    $input.Dispose()
    return $output.ToArray()
}

function Compress-RdLifeNoteZlib {
    param([byte[]]$Data)

    $output = [IO.MemoryStream]::new()
    $deflate = [IO.Compression.DeflateStream]::new($output, [IO.Compression.CompressionMode]::Compress)
    $deflate.Write($Data, 0, $Data.Length)
    $deflate.Dispose()

    $raw = $output.ToArray()
    $adler = Get-RdLifeNoteAdler32 $Data
    $wrapped = New-Object byte[] (2 + $raw.Length + 4)
    $wrapped[0] = 0x78
    $wrapped[1] = 0x9C
    [Array]::Copy($raw, 0, $wrapped, 2, $raw.Length)
    $wrapped[$wrapped.Length - 4] = [byte](($adler -shr 24) -band 0xFF)
    $wrapped[$wrapped.Length - 3] = [byte](($adler -shr 16) -band 0xFF)
    $wrapped[$wrapped.Length - 2] = [byte](($adler -shr 8) -band 0xFF)
    $wrapped[$wrapped.Length - 1] = [byte]($adler -band 0xFF)
    return $wrapped
}

function Find-RdLifeNoteBytes {
    param([byte[]]$Haystack, [byte[]]$Needle)

    for ($i = 0; $i -le $Haystack.Length - $Needle.Length; $i++) {
        $ok = $true
        for ($j = 0; $j -lt $Needle.Length; $j++) {
            if ($Haystack[$i + $j] -ne $Needle[$j]) { $ok = $false; break }
        }
        if ($ok) { return $i }
    }
    return -1
}

function Set-RdLifeNoteBytesAll {
    param([byte[]]$Haystack, [byte[]]$Needle, [byte[]]$Replacement)

    $count = 0
    while ($true) {
        $idx = Find-RdLifeNoteBytes -Haystack $Haystack -Needle $Needle
        if ($idx -lt 0) { break }
        [Array]::Copy($Replacement, 0, $Haystack, $idx, $Replacement.Length)
        $count++
        if ($count -gt 16) { break }
    }
    return $count
}

function Get-RdLifeNoteFDataFields {
    param($Node)

    if ($null -eq $Node) { return $null }
    if ($Node -is [System.Array]) {
        foreach ($item in $Node) {
            $result = Get-RdLifeNoteFDataFields $item
            if ($result) { return $result }
        }
        return $null
    }
    if ($Node -is [pscustomobject]) {
        if ($Node.PSObject.Properties.Name -contains "fields") {
            $fields = $Node.fields
            if ($fields -and ($fields.PSObject.Properties.Name -contains "FData") -and ($fields.FData -is [string]) -and $fields.FData.Length -gt 0) {
                return $fields
            }
        }
        foreach ($prop in $Node.PSObject.Properties) {
            $result = Get-RdLifeNoteFDataFields $prop.Value
            if ($result) { return $result }
        }
    }
    return $null
}

function Set-RdLifeDesignerNote {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ItemNo,

        [string]$Note,

        [string]$ClaimTaskId,

        [string]$TemplatePath = $RdLifeNoteTemplateDefault,

        [string]$AuthHeader,

        [switch]$NoSend
    )

    $detail = Invoke-RdLifeMcpTool -Name "get_rdlife_workitem_detail" -Arguments @{ itemNo = $ItemNo }
    if (-not $detail.found) {
        throw "RDLife work item not found: $ItemNo"
    }
    $targetPk = [guid]([string]$detail.detail.Wkim001)

    $claimWord = [string]::Concat([char]0x5F85, [char]0x9818, [char]0x53D6)
    if ($ClaimTaskId) {
        $currentNote = [string]$detail.detail.Wkim017
        if ($currentNote.StartsWith($ClaimTaskId + $claimWord)) {
            return [pscustomobject]@{
                ItemNo     = $ItemNo
                TargetPk   = $targetPk.ToString()
                Note       = $currentNote
                DeltaBytes = 0
                PkReplaced = 0
                Sent       = $false
                Skipped    = $true
                Reason     = "already-claimed"
                StatusCode = $null
                Response   = $null
            }
        }
        $Note = $ClaimTaskId + $claimWord + "`r`n`r`n" + $currentNote
    }

    $templateText = Get-Content -LiteralPath $TemplatePath -Raw -Encoding UTF8
    $fields = Get-RdLifeNoteFDataFields ($templateText | ConvertFrom-Json)
    if (-not $fields) {
        throw "FData not found in template: $TemplatePath"
    }

    $binary = Expand-RdLifeNoteZlib ([Convert]::FromBase64String($fields.FData))

    $oldValue = [Text.Encoding]::Unicode.GetBytes($RdLifeNoteTemplateValue)
    $valueOffset = Find-RdLifeNoteBytes -Haystack $binary -Needle $oldValue
    if ($valueOffset -lt 0) {
        throw "Template note value marker was not found."
    }
    $fieldOffset = $valueOffset - 6
    if ($fieldOffset -lt 0 -or $binary[$fieldOffset] -ne 0x10 -or $binary[$fieldOffset + 1] -ne 0x00) {
        throw "Note field marker (10 00) was not found before the template note value."
    }

    $oldValueLength = [int][BitConverter]::ToUInt32($binary, $fieldOffset + 2)
    $newValue = [Text.Encoding]::Unicode.GetBytes($Note)
    $delta = $newValue.Length - $oldValueLength

    $prefix = $binary[0..($fieldOffset + 1)]
    $lengthBytes = [BitConverter]::GetBytes([uint32]$newValue.Length)
    $suffix = $binary[($valueOffset + $oldValueLength)..($binary.Length - 1)]
    $newBinary = [byte[]]($prefix + $lengthBytes + $newValue + $suffix)

    $blockLength = [BitConverter]::ToUInt32($newBinary, 8)
    [BitConverter]::GetBytes([uint32]($blockLength + $delta)).CopyTo($newBinary, 8)

    $pkHits = Set-RdLifeNoteBytesAll -Haystack $newBinary -Needle $RdLifeNoteTemplatePk.ToByteArray() -Replacement $targetPk.ToByteArray()
    if ($pkHits -lt 1) {
        throw "Template primary key bytes were not found; cannot retarget to $ItemNo."
    }

    $base64 = [Convert]::ToBase64String((Compress-RdLifeNoteZlib $newBinary))
    $parts = [System.Collections.Generic.List[string]]::new()
    for ($i = 0; $i -lt $base64.Length; $i += 76) {
        $parts.Add($base64.Substring($i, [Math]::Min(76, $base64.Length - $i)))
    }
    $base64Wrapped = $parts -join "`r`n"
    $base64Json = $base64Wrapped.Replace("`r", "\r").Replace("`n", "\n")

    $pattern = '"FData":"(?<data>.*?)","FLength":(?<len>\d+),"FMaxCapacity"'
    $replacement = '"FData":"' + $base64Json + '","FLength":' + $base64Wrapped.Length + ',"FMaxCapacity"'
    $body = [regex]::Replace($templateText, $pattern, $replacement, 1)

    $result = [ordered]@{
        ItemNo     = $ItemNo
        TargetPk   = $targetPk.ToString()
        Note       = $Note
        DeltaBytes = $delta
        PkReplaced = $pkHits
        Sent       = $false
        Skipped    = $false
        Reason     = ""
        StatusCode = $null
        Response   = $null
    }

    if (-not $NoSend) {
        $auth = Get-RdLifeDataSnapAuth -AuthHeader $AuthHeader
        $response = Invoke-WebRequest `
            -Uri $RdLifeNoteUri `
            -Method Post `
            -Headers @{ Authorization = $auth } `
            -ContentType "text/plain;charset=UTF-8" `
            -Body $body `
            -UseBasicParsing `
            -TimeoutSec 30
        $result.Sent = $true
        $result.StatusCode = [int]$response.StatusCode
        $result.Response = $response.Content
    }

    return [pscustomobject]$result
}
