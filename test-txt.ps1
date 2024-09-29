
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
            } else {
                $rowData = @{}
                for ($k = 0; $k -lt $headers.Count; $k++) {
                    if ($k -lt $cells.Count) {
                        $cellValue = $cells[$k].Groups[1].Value -replace '<.*?>', '' -replace '&nbsp;', ' ' -replace '^\s+|\s+$', ''
                        $rowData[$headers[$k]] = $cellValue
                    } else {
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
    

function convent-html2table ($htmlFilePath) {
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
            } else {
                $rowData = @{}
                for ($k = 0; $k -lt $headers.Count; $k++) {
                    if ($k -lt $cells.Count) {
                        $rowData[$headers[$k]] = $cells[$k].Groups[1].Value.Trim()
                    } else {
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
            } else {
                $rowData = @{}
                for ($k = 0; $k -lt $headers.Count; $k++) {
                    if ($k -lt $cells.Count) {
                        $cellValue = Clean-HtmlContent $cells[$k].Groups[1].Value
                        $rowData[$headers[$k]] = $cellValue
                    } else {
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

$html = Convert-Html2Table3 -htmlFilePath .\html\vhcy_showjob_vhcy_202409280837.html
$html.table1