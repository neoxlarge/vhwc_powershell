## 安裝 HCAServiSign

param($runadmin)

Import-Module ((Split-Path $PSCommandPath) + "\get-installedprogramlist.psm1")

function install-HCA {

    ## 安裝 HCAServiSign
    ### 找出軟體是否己安裝

    $all_installed_program = get-installedprogramlist

    $software_name = "HCAServiSignAdapterSetup"
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like "$software_name*" }


    if ($null -eq $software_is_installed) {
        Write-Output "Start to install $software_name"

        #來源路徑 ,要復制的路徑,and 安裝執行程式名稱
        $software_path = get-item -Path "\\172.20.5.187\mis\05-HCAServiSign醫事卡解鎖"
        $software_exec = "HCAServiSignAdapterSetup.exe"
          

        #復制檔案到暫存資料夾
        Copy-Item -Path $software_path -Destination $env:TEMP -Recurse -Force 

        #installing...
        $process_id = Start-Process -FilePath ($env:TEMP + "\" + $software_path.Name + "\" + $software_exec) -PassThru
    
        #依安裝文件, HCAServiSignMonitor會最後被開啟, 所以檢查到該程序執行後, 表示安裝完成.
        $process_exist = $null
        while ($process_exist -eq $null) {
            $process_exist = Get-Process -Name HCAServiSignMonitor -ErrorAction SilentlyContinue
            if ($process_exist -ne $null) { Stop-Process -Name $process_id.Name }
            write-output ($process_id.Name + "is installing...wait 5 seconds.")
            Start-Sleep -Seconds 5
        }


        #安裝完, 再重新取得安裝資訊
        $all_installed_program = get-installedprogramlist
        $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like "$software_name*" }

    } 

    Write-output ("software has installed:" + $software_is_installed.DisplayName )
    Write-Output ("Version:" + $software_is_installed.DisplayVersion)

}


#檔案獨立執行時會執行函式, 如果是被?入時不會執行函式.
if ($run_main -eq $null) {

    #檢查是否管理員
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #如果非管理員, 就試著run as admin, 並傳入runadmin 參數1. 因為在網域一般使用者永遠拿不是管理員權限, 會造成無限重跑. 此參數用來輔助判斷只跑一次. 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    
    }

    if ($check_admin) { 
        install-HCA
    }
    else {
        Write-Warning "無法取得管理員權限來安裝軟體, 請以管理員帳號重試."
    }
    pause
}



