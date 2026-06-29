$RdLifeMcpUrl = "https://crd.winton.com.tw/Mcp/Rdlife"

function Invoke-RdLifeMcpTool {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [hashtable]$Arguments = @{},

        [string]$Url = $RdLifeMcpUrl,

        [int]$TimeoutSec = 30
    )

    if (-not $env:WINTON_MCP_USER_ID -or -not $env:WINTON_MCP_TOKEN) {
        throw "RDLife MCP requires environment variables WINTON_MCP_USER_ID and WINTON_MCP_TOKEN."
    }

    $headers = @{
        "X-Mcp-UserId" = $env:WINTON_MCP_USER_ID
        "X-Mcp-Token"  = $env:WINTON_MCP_TOKEN
        "Accept"       = "application/json, text/event-stream"
    }

    $body = @{
        jsonrpc = "2.0"
        id      = 1
        method  = "tools/call"
        params  = @{
            name      = $Name
            arguments = $Arguments
        }
    } | ConvertTo-Json -Depth 20 -Compress

    $response = Invoke-WebRequest `
        -UseBasicParsing `
        -Uri $Url `
        -Method POST `
        -Headers $headers `
        -ContentType "application/json" `
        -Body ([System.Text.Encoding]::UTF8.GetBytes($body)) `
        -TimeoutSec $TimeoutSec

    $raw = $response.Content
    if ($raw -match "data:") {
        $raw = (($raw -split "`n") | Where-Object { $_ -like "data:*" } | ForEach-Object { $_.Substring(5).Trim() }) -join ""
    }

    $rpc = $raw | ConvertFrom-Json
    if ($rpc.error) {
        throw "RDLife MCP error for $($Name): $($rpc.error.message)"
    }

    $text = $rpc.result.content[0].text
    if ($rpc.result.isError) {
        throw "RDLife MCP tool $($Name) failed: $text"
    }

    return ($text | ConvertFrom-Json)
}

function ConvertTo-RdLifeItem {
    param(
        [Parameter(Mandatory = $true)]
        $Source,

        [bool]$IsBug = $false
    )

    return [pscustomobject]@{
        itemNo          = [string]$Source.ItemNo
        progNo          = ([string]$Source.ProgNo).Trim()
        statusCode      = [string]$Source.StatusCode
        statusName      = [string]$Source.StatusName
        functionPoints  = $Source.FunctionPoints
        calculatedHours = $Source.CalcHours
        description     = [string]$Source.Description
        reason          = [string]$Source.Reason
        patch           = [string]$Source.Patch
        designerId      = [string]$Source.PgId
        system          = [string]$Source.SystemCode
        version         = [string]$Source.VersionName
        isBug           = $IsBug
    }
}

function Get-RdLifeWorkItem {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ItemNo
    )

    $detail = Invoke-RdLifeMcpTool -Name "get_rdlife_workitem_detail" -Arguments @{ itemNo = $ItemNo }
    if (-not $detail.found) {
        return $null
    }

    return ConvertTo-RdLifeItem -Source $detail.summary -IsBug $false
}

function Get-RdLifeCandidateWorkItems {
    param(
        [string]$DesignerId,
        [string]$Status = "W04",
        [int]$Top = 200
    )

    $arguments = @{ status = $Status; top = $Top }
    if ($DesignerId) {
        $arguments.pgId = $DesignerId
    }

    $result = Invoke-RdLifeMcpTool -Name "search_rdlife_workitem" -Arguments $arguments

    return @($result.rows | ForEach-Object {
        ConvertTo-RdLifeItem -Source $_ -IsBug $false
    })
}
