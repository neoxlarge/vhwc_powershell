# 安裝office 2021
# office 2021 的安裝目前都改成需要用 office-deployment-tool 安裝.
# 簡單來說就是setup.exe 配合設定檔 .xml 來安裝.
# 細節請參考底下連結
# https://learn.microsoft.com/zh-tw/microsoft-365-apps/deploy/overview-office-deployment-tool
# xml裡的設定可以:
# 1. 安裝來源檔案可以指到綱路上的路徑, 不用下載到本機.
# 2. 自動輸入序號, 但為管控系統, 不建議使用.

# 1. 會先檢查有無安裝office 2003, 有的話移掉.
# 2. 會先檢查有無安裝office 2021, 有的話移掉.

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


# 檢查其他MS office 是否安裝
function get-installedprogramlist {
    # 取得所有安裝的軟體,底下安裝軟體會用到.

    ### Win32_product的清單並不完整， Winnexus 並不在裡面.
    ### $all_installed_program = Get-WmiObject -Class Win32_Product

    ### 所有的軟體會在底下這三個登錄檔路徑中

    $software_reg_path = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    return (Get-ItemProperty -Path $software_reg_path -ErrorAction SilentlyContinue)
}

$office_installed = (get-installedprogramlist) | Where-Object { $_.DisplayName -like "*Microsoft office*" }

if ($office_installed -ne $null) {
    
    $office_installed.displayname
    #$install_confirm = Read-Host "以上Office已經安裝, 請移除後再安裝, 以免出錯(Y:裝/N:不裝)"
    # 創建確認視窗函數
    function Show-InstallConfirmDialog {
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "安裝確認"
        $form.Size = New-Object System.Drawing.Size(450, 500)
        $form.StartPosition = "CenterScreen"
        $form.FormBorderStyle = "FixedDialog"
        $form.MaximizeBox = $false
        $form.MinimizeBox = $false

        # 警告圖示
        $pictureBox = New-Object Windows.Forms.PictureBox
        $pictureBox.Location = New-Object Drawing.Point(20, 20)
        $pictureBox.Size = New-Object Drawing.Size(48, 48)
        $pictureBox.Image = [System.Drawing.SystemIcons]::Warning.ToBitmap()
        $form.Controls.Add($pictureBox)

        # 訊息標籤
        $label = New-Object System.Windows.Forms.Label
        $label.Location = New-Object System.Drawing.Point(80, 20)
        $label.Size = New-Object System.Drawing.Size(340, 100)
        $text = $office_installed.displayname + "`n`n 以上Office已經安裝, 請移除後再安裝, 以免出錯。"
        $label.Text = $text
        #"以上Office已經安裝,請移除後再安裝,以免出錯。`n`n是否要繼續安裝?"
        $label.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10)
        $form.Controls.Add($label)

        # 是按鈕
        $yesButton = New-Object System.Windows.Forms.Button
        $yesButton.Location = New-Object System.Drawing.Point(100, 420)
        $yesButton.Size = New-Object System.Drawing.Size(100, 30)
        $yesButton.Text = "是(Y)"
        $yesButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
        $form.Controls.Add($yesButton)

        # 否按鈕
        $noButton = New-Object System.Windows.Forms.Button
        $noButton.Location = New-Object System.Drawing.Point(250, 420)
        $noButton.Size = New-Object System.Drawing.Size(100, 30)
        $noButton.Text = "否(N)"
        $noButton.DialogResult = [System.Windows.Forms.DialogResult]::No
        $form.Controls.Add($noButton)

        # 設定預設按鈕
        $form.AcceptButton = $yesButton
        $form.CancelButton = $noButton

        # 加入按鍵事件
        $form.KeyPreview = $true
        $form.Add_KeyDown({
                if ($_.KeyCode -eq "Y") {
                    $form.DialogResult = [System.Windows.Forms.DialogResult]::Yes
                    $form.Close()
                }
                elseif ($_.KeyCode -eq "N") {
                    $form.DialogResult = [System.Windows.Forms.DialogResult]::No
                    $form.Close()
                }
            })

        # 顯示視窗並取得結果
        $result = $form.ShowDialog()

        # 回傳結果
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            return "Y"
        }
        else {
            return "N"
        }
    }

    # 使用這個函數替代原本的 Read-Host
    $install_confirm = Show-InstallConfirmDialog
}   
else {
    $install_confirm = "Y"
}

if ($install_confirm -eq "N") { 
    exit
}

# 安裝office 2021
Start-Process -FilePath "$office_source_path\$odt_exe" -ArgumentList "/configure $office_source_path\$odt_xml"