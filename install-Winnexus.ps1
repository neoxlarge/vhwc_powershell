# 安裝Winnexus

param($runadmin)
Import-Module ((Split-Path $PSCommandPath) + "\get-installedprogramlist.psm1")

function install-WinNexus {
    # 安裝Winnexus
    $software_name = "WinNexus"
    # FIXME: $software_path = "\\172.20.5.187\mis\13-Winnexus\Winnexus_1.2.4.7\13-Winnexus"
    $software_path = "\\172.20.1.122\share\software\00newpc\13-Winnexus\Winnexus_1.2.4.7"
    $software_exec = "Install_Desktop.1.2.4.7.exe"

    $Username = "vhcy\vhwcmis"
    $Password = "Mis20190610"
    $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)

    ## 找出軟體是否己安裝

    $all_installed_program = get-installedprogramlist

   
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like "$software_name*" }

    if ($software_is_installed -eq $null) {
        Write-OutPut "Start to install: $software_name"

        #$software_path = get-item -Path $software_path
        
        #復制檔案到temp
        #copy-item 無法接認證, 須要從psdrive接, 所以要掛driver.
        $net_driver = "vhwcdrive" #只是給個driver名字而己.
        New-PSDrive -Name $net_driver -Root $software_path -PSProvider FileSystem -Credential $credential
        Copy-Item -Path "$($net_driver):\" -Destination "$($env:TEMP)\$software_name" -Recurse -Force -Verbose
        Remove-PSDrive -Name $net_driver

        #installing...

        Start-Process -FilePath "$($env:TEMP)\$software_Name\$software_exec" -ArgumentList ("/suppressmsgboxes /log:install_winnexus.log") -Wait
        Start-Sleep -Seconds 5 
   
     
        #安裝完, 再重新取得安裝資訊
        $all_installed_program = get-installedprogramlist
        $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like "$software_name *" }
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

    if ($check_admin) { 
        install-WinNexus
    }
    else {
        Write-Warning "無法取得管理員權限來安裝軟體, 請以管理員帳號重試."
    }
    pause
}