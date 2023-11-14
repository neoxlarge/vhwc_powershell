
function Get-IPv4Address {
    <#
    只能在172.*才能用.
    #>

    $ip = Get-WmiObject -Class Win32_NetworkAdapterConfiguration |
    Where-Object { $_.IPAddress -ne $null -and $_.IPAddress[0] -like "172.*" } |
    Select-Object -ExpandProperty IPAddress |
    Where-Object { $_ -match "\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}" } |
    Select-Object -First 1
    return $ip
}



Write-Host "active the office 2020"

$Username = "vhcy\vhwcmis"
$Password = "Mis20190610"
$securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)

switch ($env:PROCESSOR_ARCHITECTURE) {
    "AMD64" { $officeospp_path = "C:\Program Files (x86)\Microsoft Office\Office14\OSPP.VBS" }
    "x86" { $officeospp_path = "C:\Program Files\Microsoft Office\Office14\OSPP.VBS" }
    
}

$key_list = [ordered]@{
    a = "BK63F-CDJGW-9MFY8-9J9D2-2VK3M"
    b = "27P9Q-9CMHC-DPMPV-4XDQD-BFCYT"
    c = "GR3XF-XM87M-QK2TX-6RR8Q-BPJWX"
    d = "GVH89-PKWTH-MRCTP-CMYTJ-8D4WV"
    e = "XHR87-BWT9M-W9HJG-D83V9-CPMHH"

    f = "27DWW-JQM62-4KBC2-HHFCG-VHRBB"
    g = "8W88H-RMRKC-XQ3DC-GWM6Q-H4YRX"
    h = "H6B73-3W3PD-XD3FD-V4BY7-7JF36"
    i = "HP72H-MJPQK-GWBPM-MDKTX-CBPC7"
    j = "J8JHM-KFX6C-G7RBX-47RM2-QC9C2"
    k = "J8JHM-KFX6C-G7RBX-47RM2-QC9C2"
    l = "THX9R-D6H2K-T9X66-B36FT-8YYFJ"
}


#檢查是否己經啟用
$run_ospp = Start-Process -FilePath cscript.exe -ArgumentList """$officeospp_path"" /dstatus" -NoNewWindow -Credential $credential -PassThru -RedirectStandardOutput "$($ENV:TEMP)\officehasactived.txt"
$run_ospp.waitforexit() 

$get_log = get-content -Path "$($ENV:TEMP)\officehasactived.txt"

$check_hasactived = $get_log | Select-String -Pattern "---LICENSED---"

if ($check_hasactived) {
    Write-Output $get_log
    Write-Output "己經啟用過office 2010了."
}
else {


    # 顯示金鑰清單
    foreach ($key in $key_list.Keys) {
        Write-Host "$key. $($key_list[$key])"
    }

    # 提示使用者選擇
    $userChoice = Read-Host "請選擇一個Key（輸入對應的字母）"

    # 獲取使用者選擇的金鑰
    $office_key = $key_list[$userChoice]
    # 顯示使用者選擇的金鑰
    Write-Host "使用者選擇的Key是: $office_key"

    #pause
    $result = Start-Process -FilePath cscript.exe -ArgumentList """$officeospp_path"" /inpkey:""$office_key""" -NoNewWindow -Credential $credential -PassThru -RedirectStandardOutput "$($ENV:TEMP)\officeospp.txt"

    $result.WaitForExit()
    Get-Content -Path "$($ENV:TEMP)\officeospp.txt"
    Start-Sleep -s 3

    for ($i = 0; $i -lt 2; $i++) {
        #先排除網路問題
        ipconfig /flushdns
        #ipconfig /renew

        $result_act = Start-Process -FilePath cscript.exe -ArgumentList """$officeospp_path"" /act" -NoNewWindow -Credential $credential -PassThru -RedirectStandardOutput "$($ENV:TEMP)\officeact.txt"
        $result_act.WaitForExit()
        $log = Get-Content -Path "$($ENV:TEMP)\officeact.txt"
        $check_active = $log | Select-String -Pattern "Product activation successful"

        Write-Host $log

        Start-Sleep -Seconds 1

        if ($check_active) {
            $save_tofile = "$env:COMPUTERNAME $(Get-IPv4Address) $office_key $(get-date)"
            Out-File -FilePath \\172.20.5.185\powershell\active_office.log -Append -InputObject $save_tofile
            break
        }

    }


}
Start-Sleep -s 30

