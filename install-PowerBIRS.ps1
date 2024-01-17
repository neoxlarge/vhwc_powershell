# 安裝PowerBI RS (Report Server)

param($runadmin)

Import-Module ((Split-Path $PSCommandPath) + "\get-installedprogramlist.psm1")


#取得OS的版本
function Get-OSVersion {
    $os = (Get-WmiObject -Class Win32_OperatingSystem).Caption

    if ($os -like "*Windows 7*") {
        return "Windows 7"
    }
    elseif ($os -like "*Windows 10*") {
        return "Windows 10"
    }
    elseif ($os -like "*Windows 11*") {
        return "Windows 11"
    }
    else {
        return "Unknown OS"
    }
}

function install-PowerBIRS {
    # 安裝Winnexus
    $software_name = "Microsoft PowerBI Desktop (x64) (September 2023)"
    $software_path = "\\172.20.5.187\mis\26-PowerBI"
    $software_exec = "PBIDesktopSetupRS_x64.exe"

    $Username = "vhcy\vhwcmis"
    $Password = "Mis20190610"
    $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)

    ## 找出軟體是否己安裝

    $all_installed_program = get-installedprogramlist

   
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like "$software_name*" }

    if ($software_is_installed -eq $null) {
        Write-OutPut "Start to install: $software_name"

        $software_path_name = $software_path.Split("\")[-1]
        
        #復制檔案到temp
        #copy-item 無法接認證, 須要從psdrive接, 所以要掛driver.
        $net_driver = "vhwcdrive" #只是給個driver名字而己.
        New-PSDrive -Name $net_driver -Root $software_path -PSProvider FileSystem -Credential $credential
        Copy-Item -Path "$($net_driver):\$($software_exec)" -Destination "$($env:TEMP)\" -Force -Verbose
        Remove-PSDrive -Name $net_driver

        #installing...
        if (!$check_admin) {
            $running = Start-Process -FilePath "$($env:TEMP)\$software_exec" -ArgumentList "-passive -norestart ACCEPT_EULA=1" -Credential $credential -PassThru
            $running.WaitForExit()
        }
        else {
            Start-Process -FilePath "$($env:TEMP)\$software_exec" -ArgumentList "-passive -norestart ACCEPT_EULA=1" -Wait
        }


        Start-Sleep -Seconds 5 
   
     
        #安裝完, 再重新取得安裝資訊
        $all_installed_program = get-installedprogramlist
        $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like "$software_name*" }
    } 

    Write-Output ("Software has installed: " + $software_is_installed.DisplayName)
    Write-Output ("Version: " + $software_is_installed.DisplayVersion)



}



#檔案獨立執行時會執行函式, 如果是被匯入時不會執行函式.
if ($run_main -eq $null) {

    #檢查是否管理員
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #如果非管理員, 就試著run as admin, 並傳入runadmin 參數1. 因為在網域一般使用者永遠拿不是管理員權限, 會造成無限重跑. 此參數用來輔助判斷只跑一次. 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    
    }

    if ($(Get-OSVersion) -in @("Windows 10", "Windows 11")) {
        if ($check_admin) { 
            install-PowerBIRS
        }
        
    } else {
        Write-Warning "非Windows 10 或 11 系統. 無法安裝POWER BI"
        
    }


    #pause
}


$(Get-OSVersion) -in @("Windows 10", "Windows 11")