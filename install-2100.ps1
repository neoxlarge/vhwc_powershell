# 安裝公文系統2100
# 安裝教學文件 http://172.22.250.179/ii/install/sysdeploy.htm
# 20240224 update insall 2100


param($runadmin)

$mymodule_path = "$(Split-Path $PSCommandPath)\"
Import-Module -name "$($mymodule_path)vhwcmis_module.psm1"


function install_msi {
    #$mode 是msiexec的參數, 預設i是安裝, fa是強制重新裝
    #msi是安裝的檔名
    param($mode = "i", $msi)
    Write-Output $msi
    if ($check_admin) {
        $proc = Start-Process -FilePath "msiexec.exe" -ArgumentList "/$mode $msi /passive /norestart" -PassThru
    } else {
        $credential = get-admin_cred
        $proc = Start-Process -FilePath "msiexec.exe" -ArgumentList "/$mode $msi /passive /norestart" -Credential $credential -PassThru
    }
    $proc.WaitForExit()
}


function install-2100 {

    $software_name = "電子公文系統"
    $software_path = "\\172.20.5.187\mis\08-2100公文系統\01.2100公文系統安裝包_Standard"
    
    ## 找出軟體是否己安裝

    $all_installed_program = get-installedprogramlist
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }

    #復制檔案到本機暫存"
    $software_path = get-item -Path $software_path
    #Copy-Item -Path $software_path -Destination $env:temp -Recurse -Force
    Start-Process -FilePath robocopy.exe -ArgumentList "$($software_path.FullName) $($env:temp + "\" +$software_path.Name) /e /R:3 /W:5" -Wait
    #要安裝的檔案
    $package_msi = @(
        "eDocSetup_Win7.msi",
        #"HC_Setup.msi", HCA醫事人員憑證驅動程式(提供埔里分院安裝), 這之後也會裝
        #"HiCOS.msi", #HiCOP之後會再裝, 這裡不裝
        "IPD21XSetup.msi",
        "soapsdk.msi",
        #"SetupXP.msi",  #XP用的不用裝
        "UniView.msi"
    )


    if ($null -eq $software_is_installed) {
    
        Write-Output "Start to insall: $software_name"
        #安裝MSI檔
        foreach ($p in $package_msi) {
            $p_path = $env:temp + "\" + $software_path.Name + "\package\" + $p 
            install_msi -mode "i" -msi $p_path
        }

        #安裝tablePC_SDK.exe
        if ($check_admin) {
            $proc = Start-Process -FilePath ($env:temp + "\" + $software_path.Name + "\package\tablePC_SDK.exe") -ArgumentList "/v/passive" -PassThru
        } else {
            $credential = get-admin_cred
            $proc =  Start-Process -FilePath ($env:temp + "\" + $software_path.Name + "\package\tablePC_SDK.exe") -ArgumentList "/v/passive" -Credential $credential -PassThru
        }
        $proc.WaitForExit()
    }
    #20230712
    #取消重裝
    <#
    else {
        Write-Output "Reinstall $software_name"

        foreach ($p in $package_msi) {
            $p_path = $env:temp + "\" + $software_path.Name + "\package\" + $p 
            install_msi -mode "fa" -msi $p_path
        }

        #安裝tablePC_SDK.exe
        Start-Process -FilePath ($env:temp + "\" + $software_path.Name + "\package\tablePC_SDK.exe") -ArgumentList "/v/passive" -wait

    }
    #>

    
    #安裝完, 再重新取得安裝資訊
    $all_installed_program = get-installedprogramlist
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }


    Write-Output ("Software has installed: " + $software_is_installed.DisplayName)
    Write-Output ("Version: " + $software_is_installed.DisplayVersion)

}

function set-2100_env {

    #先檢查檔案是否在,沒有要復制一下
    $software_path = "\\172.20.5.187\mis\08-2100公文系統\01.2100公文系統安裝包_Standard"
    $software_path = get-item -Path $software_path

    if (!(Test-Path -Path "$env:temp\$($software_path.name)")) {
        Start-Process -FilePath robocopy.exe -ArgumentList "$($software_path.FullName) $($env:temp + "\" +$software_path.Name) /e /R:3 /W:5" -Wait
    }

    #1.
    Write-Output "復制醫院設定檔ClientSetting.ini"    
    $path = $env:TEMP + "\01.2100公文系統安裝包_Standard\ClientSetting\ClientSetting_Chiayi.ini"
    if (Test-Path -Path $path) {
        copy-item -Path $path -Destination $env:SystemDrive\2100\SSO\ClientSetting.ini -Force
    }
    else {
        Write-Warning "找不到ClientSetting_Chiayi.ini檔案, 請檢查!!"
    }

    #2.
    <#
    hicos3.1版之後，會無法簽核公文的解決方法：
    請把下列這個檔案HiCOSCSPv32放在到６４位元電腦C:\Windows\SysWOW64，３２位元電腦C:\Windows\System32，另外３.1卡片管理工具裡的「設定」有４個都打勾，再執行一下環境檔，即可解決.
    Hicoscspv32.dll 版本: 
    #20240224, 理論上在安裝期間(管理員), 就會把hicosscpv32放好, 使用者權限不會執行到覆蓋動作. 但好像把管理者權限加入到覆蓋動作中比較好. 未完成.  
    #>

    switch ($env:PROCESSOR_ARCHITECTURE) {
        "amd64" { $dll_path = "$env:windir\SysWoW64" }
        "x86" { $dll_path = "$env:windir\System32" }
        default { Write-Warning "Unknown processor architecture." }
    }

    $dll = Get-ItemPropertyValue -Path "$dll_path\HiCOSCSPv32.dll" -Name "VersionInfo" -ErrorAction SilentlyContinue
    Write-Output "HiCOSCSPv32.dll version: $($dll.ProductVersion)"
    if ($dll.ProductVersion -ne "3.0.3.21207") {
        if (Test-Path -Path ($env:temp + "\01.2100公文系統安裝包_Standard\HiCOSCSPv32.dll")) {
            #覆蓋Hicoscspv32.cll到c:\windows\system32中.
            Write-Output "覆蓋Hicoscspv32.cll(3.0.3.21207)到$dll_path"
            copy-item -Path ($env:temp + "\01.2100公文系統安裝包_Standard\HiCOSCSPv32.dll") -Destination $dll_path -Force

        }
        else { write-warning "找不到正確的HiCOSCSPv32.dll檔案" }
    }

    #3.
    Write-Output "執行 01公文環境檔.exe "
    # Start-Process -FilePath reg.exe -ArgumentList ("import " + $env:temp + "\01.2100公文系統安裝包_Standard\reg\IE9setting.reg") -Wait
    # Start-Process -FilePath reg.exe -ArgumentList ("import " + $env:temp + "\01.2100公文系統安裝包_Standard\reg\IE9setting1.reg") -Wait
    Start-Process -FilePath $($env:temp + "\01.2100公文系統安裝包_Standard\01公文環境檔.exe") -Wait
  

}



#檔案獨立執行時會執行函式, 如果是被載入時不會執行函式.
if ($run_main -eq $null) {

    #檢查是否管理員
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #如果非管理員, 就試著run as admin, 並傳入runadmin 參數1. 因為在網域一般使用者永遠拿不是管理員權限, 會造成無限重跑. 此參數用來輔助判斷只跑一次. 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    
    }

    if ($check_admin) { 
        install-2100
        #Import-Module ((Split-Path $PSCommandPath) + "\Check-2100env.ps1")    
    }
    else {
        Write-Warning "無法取得管理員權限來安裝軟體, 請以管理員帳號重試."
    }

    set-2100_env
    pause
}