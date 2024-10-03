
function Convert-Html2Table2 ($htmlFilePath) {
    # 將html檔案中的table轉換成hash table
    # 參數: $htmlFilePath: html檔案的路徑
    # 回傳: hash table

    # 讀取HTML檔案內容
    $html = Get-Content -Path $htmlFilePath -Raw -Encoding UTF8

    # 使用正則表達式匹配表格內容
    $tablePattern = "(?s)<table[^>]*>.*?</table>"
    $rowPattern = "(?s)<tr[^>]*>(.*?)</tr>"
    $cellPattern = "(?s)<t[hd][^>]*>(.*?)</t[hd]>"

    $result = @{}
    $tableMatches = [regex]::Matches($html, $tablePattern)

    for ($i = 0; $i -lt $tableMatches.Count; $i++) {
        $tableContent = $tableMatches[$i].Value
        $rows = [regex]::Matches($tableContent, $rowPattern)

        $headers = @()
        $tableData = @()

        for ($j = 0; $j -lt $rows.Count; $j++) {
            $rowContent = $rows[$j].Groups[1].Value
            $cells = [regex]::Matches($rowContent, $cellPattern)
            
            if ($j -eq 0) {
                # 假設第一行是表頭
                $headers = $cells | ForEach-Object { 
                    $_.Groups[1].Value -replace '<.*?>', '' -replace '&nbsp;', ' ' -replace '^\s+|\s+$', '' 
                }
            }
            else {
                $rowData = @{}
                for ($k = 0; $k -lt $headers.Count; $k++) {
                    if ($k -lt $cells.Count) {
                        $cellValue = $cells[$k].Groups[1].Value -replace '<.*?>', '' -replace '&nbsp;', ' ' -replace '^\s+|\s+$', ''
                        $rowData[$headers[$k]] = $cellValue
                    }
                    else {
                        $rowData[$headers[$k]] = $null
                    }
                }
                $tableData += $rowData
            }
        }

        $result["Table$($i+1)"] = $tableData
    }

    return $result
}
    

function convert-html2table ($htmlFilePath) {
    # 將html檔案中的table轉換成hash table
    # 參數: $htmlFilePath: html檔案的路徑
    # 回傳: hash table

    # 讀取HTML檔案內容
    $html = Get-Content -Path $htmlFilePath -Raw

    # 使用正則表達式匹配表格內容
    $tablePattern = "(?s)<table.*?>(.*?)</table>"
    $rowPattern = "(?s)<tr.*?>(.*?)</tr>"
    $cellPattern = "(?s)<t[hd].*?>(.*?)</t[hd]>"

    $result = @{}
    $tableMatches = [regex]::Matches($html, $tablePattern)

    for ($i = 0; $i -lt $tableMatches.Count; $i++) {
        $tableContent = $tableMatches[$i].Groups[1].Value
        $rows = [regex]::Matches($tableContent, $rowPattern)

        $headers = @()
        $tableData = @()

        for ($j = 0; $j -lt $rows.Count; $j++) {
            $rowContent = $rows[$j].Groups[1].Value
            $cells = [regex]::Matches($rowContent, $cellPattern)
            
            if ($j -eq 0) {
                # 假設第一行是表頭
                $headers = $cells | ForEach-Object { $_.Groups[1].Value.Trim() }
            }
            else {
                $rowData = @{}
                for ($k = 0; $k -lt $headers.Count; $k++) {
                    if ($k -lt $cells.Count) {
                        $rowData[$headers[$k]] = $cells[$k].Groups[1].Value.Trim()
                    }
                    else {
                        $rowData[$headers[$k]] = $null
                    }
                }
                $tableData += $rowData
            }
        }

        $result["Table$($i+1)"] = $tableData
    }

    return $result
}

function Convert-Html2Table3 ($htmlFilePath) {
    # 將html檔案中的table轉換成hash table
    # 參數: $htmlFilePath: html檔案的路徑
    # 回傳: hash table

    # 讀取HTML檔案內容
    $html = Get-Content -Path $htmlFilePath -Raw -Encoding UTF8

    # 使用正則表達式匹配表格內容
    $tablePattern = "(?s)<table[^>]*>.*?</table>"
    $rowPattern = "(?s)<tr[^>]*>(.*?)</tr>"
    $cellPattern = "(?s)<t[hd][^>]*>(.*?)</t[hd]>"

    # 函數：清理HTML內容
    function Clean-HtmlContent($content) {
        # 處理特殊情況，如 <a> 標籤
        $content = [regex]::Replace($content, '<a[^>]*>(.*?)</a>', '$1')
        
        # 移除其他HTML標籤
        $content = $content -replace '<[^>]+>', ''
        
        # 替換HTML實體和清理空白
        $content = $content -replace '&nbsp;', ' ' `
            -replace '&lt;', '<' `
            -replace '&gt;', '>' `
            -replace '&amp;', '&' `
            -replace '^\s+|\s+$', '' `
            -replace '\s+', ' '
        return $content
    }

    $result = @{}
    $tableMatches = [regex]::Matches($html, $tablePattern)

    for ($i = 0; $i -lt $tableMatches.Count; $i++) {
        $tableContent = $tableMatches[$i].Value
        $rows = [regex]::Matches($tableContent, $rowPattern)

        $headers = @()
        $tableData = @()

        for ($j = 0; $j -lt $rows.Count; $j++) {
            $rowContent = $rows[$j].Groups[1].Value
            $cells = [regex]::Matches($rowContent, $cellPattern)
            
            if ($j -eq 0) {
                # 假設第一行是表頭
                $headers = $cells | ForEach-Object { 
                    Clean-HtmlContent $_.Groups[1].Value
                }
            }
            else {
                $rowData = @{}
                for ($k = 0; $k -lt $headers.Count; $k++) {
                    if ($k -lt $cells.Count) {
                        $cellValue = Clean-HtmlContent $cells[$k].Groups[1].Value
                        $rowData[$headers[$k]] = $cellValue
                    }
                    else {
                        $rowData[$headers[$k]] = $null
                    }
                }
                $tableData += $rowData
            }
        }

        $result["Table$($i+1)"] = $tableData
    }

    return $result
}

$html = Convert-Html2Table2 -htmlFilePath .\html\showjob_vhwc_202410031237.html
$html.table1 | Format-List


# 從 HTML 檔案讀取表格資料
$html = Get-Content ".\html\showjob_vhwc_202410031237.html"

# 使用 ConvertFrom-StringData 將表格資料轉換成雜湊表
$hashtable = @{}
$html | Select-String -Pattern '<td(.*?)>(.*?)</td>' -AllMatches | ForEach-Object {
    $key = $_.Matches.Groups[2].Value.Trim()
    $value = $_.Matches.Groups[3].Value.Trim()
    $hashtable[$key] = $value
}

# 顯示雜湊表
$hashtable | format-list

function Convert-HtmlTableToHashtable {
    param(
        [string]$HtmlFilePath
    )

    # 從 HTML 檔案讀取表格資料
    $html = Get-Content $HtmlFilePath

    # 使用 ConvertFrom-StringData 將表格資料轉換成雜湊表
    $hashtable = @{}
    $header = $null
    $i = 0
    $tableFound = $false

    $html | Select-String -Pattern '<td>(.*?)</td>' -AllMatches | ForEach-Object {
        if ($_.Matches.Groups[1].Value.Trim() -eq 'gv1') {
            $tableFound = $true
            return
        }
        if ($tableFound) {
            if ($i -eq 0) {
                $header = $_.Matches.Groups[1].Value.Trim()
            }
            else {
                $value = $_.Matches.Groups[1].Value.Trim()
                $hashtable[$header] = $value
                $i = 0 
            }
            $i++
        }
    }

    # 傳回雜湊表
    return $hashtable
}
$aa = Convert-HtmlTableToHashtable -HtmlFilePath .\html\showjob_vhcy_202410031237.html

<#
.SYNOPSIS
    將包含批次作業執行資訊的 HTML 檔案轉換為 HashTable。

.DESCRIPTION
    此函數讀取指定的 HTML 檔案，解析其中的表格資料（假設為批次作業執行資訊），
    並將其轉換為 PowerShell HashTable 陣列。它使用表格中的 <th> 元素作為 HashTable 的鍵。

.PARAMETER HtmlFilePath
    要解析的 HTML 檔案的完整路徑。

.EXAMPLE
    $jobData = Convert-Html2HashTable_ShowJob -HtmlFilePath "C:\temp\batchjobs.html"
    $jobData | Format-Table -AutoSize

.NOTES
    作者: Assistant
    版本: 1.0
    最後更新: 2023-06-07
#>
function Convert-Html2HashTable_ShowJob {

    <#
    .SYNOPSIS
    將包含批次作業執行資訊的 HTML 檔案轉換為 HashTable。

    .DESCRIPTION
    此函數讀取指定的 HTML 檔案，解析其中的表格資料（假設為批次作業執行資訊），
    並將其轉換為 PowerShell HashTable 陣列。它使用表格中的 <th> 元素作為 HashTable 的鍵。

    .PARAMETER HtmlFilePath
    要解析的 HTML 檔案的完整路徑。

    .EXAMPLE
    $jobData = Convert-Html2HashTable_ShowJob -HtmlFilePath "C:\temp\batchjobs.html"
    $jobData | Format-Table -AutoSize

    #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$HtmlFilePath
    )

    # 檢查檔案是否存在
    if (-not (Test-Path $HtmlFilePath)) {
        throw "找不到檔案: $HtmlFilePath"
    }

    # 從檔案中讀取 HTML 內容
    $HtmlContent = Get-Content -Path $HtmlFilePath -Raw

    # 載入 HTML 內容
    $html = New-Object -ComObject "HTMLFile"
    $html.IHTMLDocument2_write($HtmlContent)

    # 找到第二個表格（索引為 1，因為索引從 0 開始）
    $table = $html.getElementsByTagName("table")[1]

    # 提取表頭
    $headers = @()
    foreach ($th in $table.getElementsByTagName("th")) {
        $headers += $th.innerText.Trim()
    }

    # 初始化一個空陣列來儲存結果
    $results = @()

    # 遍歷表格中的每一行（跳過表頭行）
    foreach ($row in $table.rows | Select-Object -Skip 1) {
        $rowData = @{}
        
        # 遍歷行中的每個單元格
        for ($i = 0; $i -lt $headers.Count; $i++) {
            $rowData[$headers[$i]] = $row.cells[$i].innerText.Trim()
        }

        # 將 HashTable 添加到結果陣列中
        $results += $rowData
    }

    return $results
}



Convert-Html2HashTable_ShowJob -HtmlFilePath .\html\showjob_vhcy_202410031220.html