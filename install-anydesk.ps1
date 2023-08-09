<# 安裝anydesk
 1. anydesk exe package download link: https://anydesk.com/zht/downloads/windows
 2. anydesk msi package download link: https://download.anydesk.com/AnyDesk.msi
 3. anydesk command line https://support.anydesk.com/knowledge/command-line-interface-for-windows
    Parameter Description:
    --install <location>	
    Install AnyDesk to the specified <location>.
    e.g. C:\Program Files (x86)\AnyDesk

    --start-with-win	Automatically start AnyDesk with Windows. This is needed to be able to connect after restarting the system.
    --create-shortcuts	Create start menu entry.
    --create-desktop-icon	Create a link on the desktop for AnyDesk.
    --remove-first	Remove the current AnyDesk installation before installing the new one. e.g. when updating AnyDesk manually.
    --silent	Do not start AnyDesk after installation and do not display error message boxes during installation.
    --update-manually	Update AnyDesk manually
    (Default for custom clients).
    --update-disabled	Disable automatic update of AnyDesk.
    --update-auto	Update AnyDesk automatically
    (Default for standard clients, not available for custom clients).
#>


param($runadmin)

Import-Module ((Split-Path $PSCommandPath) + "\get-installedprogramlist.psm1")


function install-AnyDesk {


    $software_name = "AnyDesk"
    $software_path = "\\172.20.1.122\share\software\00newpc\35-anydesk"
    $software_msi_x64 = "AnyDesk.exe"  #64bit 32bit 都同一個
    $software_msi_x32 = "AnyDesk.exe"

    #用來連線172.20.1.112的認證
    $Username = "vhcy\vhwcmis"
    $Password = "Mis20190610"
    $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)

    $all_installed_program = get-installedprogramlist
      
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }

    if ($software_is_installed -eq $null) {
    
        Write-Output "Start to insall: $software_name"

        #復制檔案到temp
        $software_path = get-item -Path $software_path
                
        #copy-item 無法接認證, 須要從psdrive接, 所以要掛driver.
        $net_driver = "vhwcdrive" #只是給個driver名字而己.
        #掛載路徑
        New-PSDrive -Name $net_driver -Root $software_path -PSProvider FileSystem -Credential $credential
        #復制
        Copy-Item -Path "$($net_driver):\" -Destination $env:TEMP -Recurse -Force -Verbose 
        #unmount路徑
        Remove-PSDrive -Name $net_driver


        ## 判斷OS是32(x86)或是64(AMD64), 其他值(ARM64)不安裝  
        switch ($env:PROCESSOR_ARCHITECTURE) {
            "AMD64" { $software_exec = $software_msi_x64 }
            "x86" { $software_exec = $software_msi_x32 }
            default { Write-OutPut "Unsupport CPU or OS:"  $env:PROCESSOR_ARCHITECTURE; $software_exec = $null }
        }

        if ($software_exec -ne $null) {
            Start-Process -FilePath "$($env:temp)\$($software_path.Name)\$software_exec" -ArgumentList "--silent --create-desktop-icon " -Wait
            Start-Sleep -Seconds 5 
        }
        else {
            Write-Warning "$software_name 無法正常安裝."
        }
      
             
        #安裝完, 再重新取得安裝資訊
        $all_installed_program = get-installedprogramlist
        $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }
    
    }

    Write-Output ("Software has installed: " + $software_is_installed.DisplayName)
    Write-Output ("Version: " + $software_is_installed.DisplayVersion)

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
        install-SMAConnectAgent
        install-NX
    }
    else {
        Write-Warning "無法取得管理員權限來安裝軟體, 請以管理員帳號重試."
    }
    pause
}