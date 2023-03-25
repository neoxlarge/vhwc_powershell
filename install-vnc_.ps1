
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

$all_installed_program = get-installedprogramlist



## 安裝VNC
$software_name = "UltraVnc"
$software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like ($software_name + "*") }

if ($null -eq $software_is_installed) {
    Write-Output ("Start to install" + $software_name)

    #來源路徑 ,要復制的路徑,and 安裝執行程式名稱
    $software_path = get-item -Path "\\172.20.5.187\mis\08-VNC\1_2_24"
    
    if (Test-Path -Path "d:\mis") {
        $software_copyto_path = "D:\mis"
    }
    else {
        $software_copyto_path = "C:\mis"
    }
    
    
    switch ($env:PROCESSOR_ARCHITECTURE) {
        "AMD64" { $software_exec = "UltraVNC_1_2_24_X64_Setup.exe" }
        "x86" { $software_exec = "UltraVNC_1_2_24_X86_Setup.exe" }
        default { Write-OutPut "Unsupport CPU or OS:"  $env:PROCESSOR_ARCHITECTURE; $software_exec = $null }
    }

    #復制檔案到D:\mis
    Copy-Item -Path $software_path -Destination $software_copyto_path -Recurse -Force 

    #installing...
    if ($software_exec) {
        Start-Process -FilePath ($software_copyto_path + "\" + $software_path.Name + "\" + $software_exec) -ArgumentList "/silent /loadinf=installvnc.inf /log=d:\mis\install_vnc.log" -Wait
        Start-Sleep -Seconds 5
    }
              
    #復制設定檔vltravnc.ini 到C:\Program Files\uvnc bvba\UltraVNC
    Copy-Item -Path ($software_copyto_path + "\" + $software_path.Name + "\ultravnc.ini") -Destination ($env:ProgramFiles + "\uvnc bvba\UltraVNC") -Force

    #安裝完, 刪除安裝檔案
    remove-item -Path ($software_copyto_path + "\" + $software_path.Name) -Recurse -Force

    #安裝完, 再重新取得安裝資訊
    $all_installed_program = get-installedprogramlist
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like ($software_name + "*") }

} 

Write-output ("software has installed:" + $software_is_installed.DisplayName )
Write-Output ("Version:" + $software_is_installed.DisplayVersion)

