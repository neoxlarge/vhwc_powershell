# 安裝office 2021
# office 2021 的安裝目前都改成需要用 office-deployment-tool 安裝.
# 簡單來說就是setup.exe 配合設定檔 .xml 來安裝.
# 細節請參考底下連結
# https://learn.microsoft.com/zh-tw/microsoft-365-apps/deploy/overview-office-deployment-tool
# xml裡的設定可以:
# 1. 安裝來源檔案可以指到綱路上的路徑, 不用下載到本機.
# 2. 自動輸入序號, 但為管控系統, 不建議使用.


# 路徑設定
$office_source_path = "\\172.20.5.187\mis\02-Office\Office_LTSC_2021_std_64bit"
$odt_exe = "setup.exe"
$odt_xml = "install_config.xml"

# 取得管理員權限
$credential = Get-Credential -UserName "vhcy\vhwcmis" -Message "請輸入管理員帳號密碼"

# 載入UI的組件
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# 檢查權限是否正確
try {
    $check_admin = new-PSDrive -Name check_admin -PSProvider FileSystem -Root "\\172.20.5.187\mis\02-office" -Credential $credential -ErrorAction stop
}
catch {
    # 彈出錯誤訊息
    
    [System.Windows.Forms.MessageBox]::Show(
        "權限不足, 請確認管理員權限正確", # 訊息內容
        "錯誤", # 視窗標題
        [System.Windows.Forms.MessageBoxButtons]::OK, # 按鈕
        [System.Windows.Forms.MessageBoxIcon]::Error # 圖示
    )

    write-output "權限不足, 請確認管理員權限正確"
    exit
}
finally {
    Remove-PSDrive -Name check_admin -ErrorAction SilentlyContinue
}



# 安裝office 2021
Start-Process -FilePath "$office_source_path\$odt_exe" -ArgumentList "/configure $office_source_path\$odt_xml" -Credential $credential