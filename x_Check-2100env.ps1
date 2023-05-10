#check-2100
param($runadmin)


function check-2100env {
    #復制公文安裝檔到 d:\mis 備用.
    $software_path = get-item -Path "\\172.20.5.187\mis\08-2100公文系統\01.2100公文系統安裝包_Standard"
    if (Test-Path -Path "d:\mis") {
        $software_copyto_path = "D:\mis"
    }
    else {
        $software_copyto_path = "C:\mis"
    }

    #復制檔案到D:\mis

    if (!(Test-Path -Path ($software_copyto_path + "\" + $software_path.Name))) {
        Write-Output "復制公文系統到$software_copyto_path 備存"
        Copy-Item -Path $software_path -Destination $software_copyto_path -Recurse -Force 
    }

    <#
    hicos3.1版之後，會無法簽核公文的解決方法：
    請把下列這個檔案HiCOSCSPv32放在到６４位元電腦C:\Windows\SysWOW64，３２位元電腦C:\Windows\System32，另外３.1卡片管理工具裡的「設定」有４個都打勾，再執行一下環境檔，即可解決
    Hicoscspv32.dll 版本: 
    #>

    switch ($env:PROCESSOR_ARCHITECTURE) {
        "amd64" {$dll_path = "$env:windir\SysWoW64"}
        "x86" {$dll_path = "$env:windir\System32"}
        default {Write-Warning "Unknown processor architecture."}
    }

    $dll = Get-ItemPropertyValue -Path "$dll_path\HiCOSCSPv32.dll" -Name "VersionInfo"
    if ($dll.ProductVersion -ne "3.0.3.21207") {
        if (Test-Path -Path ($software_copyto_path+"\"+$software_path.Name + "\HiCOSCSPv32.dll")) {
            #覆蓋Hicoscspv32.cll到c:\windows\system32中.
            Write-Output "覆蓋Hicoscspv32.cll(3.0.3.21207)到$dll_path"
            copy-item -Path ($software_copyto_path+"\"+$software_path.Name + "\HiCOSCSPv32.dll") -Destination $dll_path -Force

        } else {write-warning "找不到正確的HiCOSCSPv32.dll檔案"}
    }

    Write-Output "執行 01公文環境檔.exe 及IE 設定"
    Start-Process -FilePath reg.exe -ArgumentList ("import " + $software_copyto_path + "\" + $software_path.Name + "\reg\IE9setting.reg") -Wait
    Start-Process -FilePath reg.exe -ArgumentList ("import " + $software_copyto_path + "\" + $software_path.Name + "\reg\IE9setting1.reg") -Wait
    Start-Process -FilePath ($software_copyto_path + "\" + $software_path.Name + "\01公文環境檔.exe") -Wait

    
}


 
#檔案獨立執行時會執行函式, 如果是被滙入時不會執行函式.
if ($run_main -eq $null) {

    #檢查是否管理員
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #如果非管理員, 就試著run as admin, 並傳入runadmin 參數1. 因為在網域一般使用者永遠拿不是管理員權限, 會造成無限重跑. 此參數用來輔助判斷只跑一次. 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    
    }

    check-2100env
    pause
}