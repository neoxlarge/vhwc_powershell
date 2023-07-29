function check-OPDList {

# 將變數 $json_file 設定為 JSON 檔案 "opd_list.json" 的路徑
$json_file = "opd_list.json"

# 讀取 JSON 檔案的內容並將其儲存在變數 $json_content 中
$json_content = Get-Content -Path $json_file -Raw

# 將 JSON 內容轉換為 PowerShell 物件並將其指派給變數 $opd_json
$opd_json = ConvertFrom-Json -InputObject $json_content

# 初始化變數 $opd，並將其設定為 null
$opd = $null

# 找出符合電腦名稱的資料.
foreach ($o in $opd_json.psobject.properties) {

    $result = $o.Value.name -eq "wnur-opd-pc02"
   
    if ($result) {
        $opd = $o.Value
        break
    }
 
}
}