Write-Host "active the office 2020"
$Username = "vhcy\vhwcmis"
$Password = "Mis20190610"
$securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)

switch ($env:PROCESSOR_ARCHITECTURE) {
    "AMD64" { $officeospp_path = "C:\Program Files (x86)\Microsoft Office\Office14\OSPP.VBS"  }
    "x86" { $officeospp_path = "C:\Program Files\Microsoft Office\Office14\OSPP.VBS" }
    
}

$key_list = [ordered]@{
    1 = "BK63F-CDJGW-9MFY8-9J9D2-2VK3M"
    2 = "27P9Q-9CMHC-DPMPV-4XDQD-BFCYT"
    3 = "GR3XF-XM87M-QK2TX-6RR8Q-BPJWX"
    4 = "GVH89-PKWTH-MRCTP-CMYTJ-8D4WV"
    5 = "XHR87-BWT9M-W9HJG-D83V9-CPMHH"

    a = "27DWW-JQM62-4KBC2-HHFCG-VHRBB"
    b = "8W88H-RMRKC-XQ3DC-GWM6Q-H4YRX"
    e = "H6B73-3W3PD-XD3FD-V4BY7-7JF36"
    d = "HP72H-MJPQK-GWBPM-MDKTX-CBPC7"
    f = "J8JHM-KFX6C-G7RBX-47RM2-QC9C2"
    g = "J8JHM-KFX6C-G7RBX-47RM2-QC9C2"
    h = "THX9R-D6H2K-T9X66-B36FT-8YYFJ"
}


# 顯示金鑰清單
foreach ($key in $key_list.Keys) {
    Write-Host "$key. $($key_list[$key])"
}

# 提示使用者選擇
$userChoice = Read-Host "請選擇一個Key（輸入對應的數字或字母）"

# 獲取使用者選擇的金鑰
$office_Key = $key_list[$userChoice]

# 顯示使用者選擇的金鑰
Write-Host "使用者選擇的Key是: $office_key"


$result = Start-Process -FilePath cscript.exe -ArgumentList """$officeospp_path"" /inpkey:""$office_key""" -NoNewWindow -Credential $credential -PassThru -RedirectStandardOutput "$($ENV:TEMP)\officeospp.txt"
#$result = Start-Process -FilePath cscript.exe -ArgumentList """$officeospp_path"" /?" -NoNewWindow -Credential $credential -PassThru -RedirectStandardOutput c:\officeospp.txt
$result.WaitForExit()
Get-Content -Path "$($ENV:TEMP)\officeospp.txt"
Start-Sleep -s 3

$result_act = Start-Process -FilePath cscript.exe -ArgumentList """$officeospp_path"" /act" -NoNewWindow -Credential $credential -PassThru -RedirectStandardOutput "$($ENV:TEMP)\officeact.txt"
$result_act.WaitForExit()
Get-Content -Path "$($ENV:TEMP)\officeact.txt"
Start-Sleep -s 3

pause