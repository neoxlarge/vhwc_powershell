<#  安裝31-庫賈氏病勾稽查詢系統
    * 此安裝程式不會在移除列表中留記錄. 在路徑中確認有無安裝.
    * 桌面捷徑更名.
#>


param($runadmin)

function install-cdcalert {


    $software_name = "庫賈氏病勾稽查詢系統"
    $software_path = "\\172.20.5.187\mis\31-庫賈氏病勾稽查詢系統\cdcClinic"
    $software_msi_x64 = "cdcalert.msi"  #64bit 32bit 都同一個
    $software_msi_x32 = "cdcalert.msi"

    $software_installed = "C:\Program Files (x86)\Changingtec\cdcClinic\cdcalert.exe"

    #用來連線172.20.1.112的認證
    $Username = "vhcy\vhwcmis"
    $Password = "Mis20190610"
    $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)

    $check_installedpath = Test-Path -Path $software_installed

    if ($check_installedpath -eq $false) {
        
        Write-Output "Start to install: $software_name"

        #復制檔案到temp
        $software_path = get-item -Path $software_path
                
        #copy-item 無法接認證, 須要從psdrive接, 所以要掛driver.
        $net_driver = "vhwcdrive" #只是給個driver名字而己.
        #掛載路徑
        New-PSDrive -Name $net_driver -Root $software_path -PSProvider FileSystem -Credential $credential
        #復制
        Copy-Item -Path "$($net_driver):\" -Destination "$($env:TEMP)\$($software_path.Name)" -Recurse -Force -Verbose 
        #unmount路徑
        Remove-PSDrive -Name $net_driver

        ## 判斷OS是32(x86)或是64(AMD64), 其他值(ARM64)不安裝  
        switch ($env:PROCESSOR_ARCHITECTURE) {
            "AMD64" { $software_exec = $software_msi_x64 }
            "x86" { $software_exec = $software_msi_x32 }
            default { Write-OutPut "Unsupport CPU or OS:"  $env:PROCESSOR_ARCHITECTURE; $software_exec = $null }
        }

        Start-Process -FilePath msiexec.exe Get-NetFirewallRule " /i $($env:TEMP)\$($software_path.Name)\$software_exec /passive" -Wait

        Start-Sleep -Seconds 2

        $software_property = Get-ItemProperty -Path $software_installed 

    }
    else {
        $software_property = Get-ItemProperty -Path $software_installed   
    }

    Write-Output ("Software has installed: " + $software_name)
    Write-Output ("Version: " + $software_property.versioninfo)

    #更名捷徑
    
    if (Test-Path "$($env:PUBLIC)\desktop\cdcalert.link") {
        Rename-Item -Path "$($env:PUBLIC)\desktop\cdcalert.link" -NewName "$software_name.lnk"
        }



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
        install-cdcalert
    }
    else {
        Write-Warning "無法取得管理員權限來安裝軟體, 請以管理員帳號重試."
    }
    pause
}