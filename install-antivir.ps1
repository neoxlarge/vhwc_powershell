# 安裝防毒 Trend Micro Apex One Security Agent

param($runadmin)

$mymodule_path = "$(Split-Path $PSCommandPath)\"
Import-Module -name "$($mymodule_path)vhwcmis_module.psm1"


function install-AntiVir {
    # 安裝防毒 Trend Micro Apex One Security Agent
    ## 找出軟體是否己安裝

    $software_name = "Trend Micro Apex One Security Agent"
    $software_path = "\\172.20.1.122\share\software\00newpc\officescan_antivir"
    $software_msi_x64 = "agent_cloud_x64.msi"
    $software_msi_x32 = "agent_cloud_x86.msi"

    $Username = "vhcy\vhwcmis"
    $Password = "Mis20190610"
    $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)

    $all_installed_program = get-installedprogramlist
      
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }

    if ($software_is_installed -eq $null) {
    
        Write-Output "Start to insall: $software_name"

        #復制檔案到temp
        #$software_path = get-item -Path $software_path
                
        #copy-item 無法接認證, 須要從psdrive接, 所以要掛driver.
        $net_driver = "vhwcdrive" #只是給個driver名字而己.
        New-PSDrive -Name $net_driver -Root $software_path -PSProvider FileSystem -Credential $credential
        Copy-Item -Path "$($net_driver):\" -Destination "$($env:TEMP)\$software_name" -Recurse -Force -Verbose 
        Remove-PSDrive -Name $net_driver


        ## 判斷OS是32(x86)或是64(AMD64), 其他值(ARM64)不安裝  
        switch ($env:PROCESSOR_ARCHITECTURE) {
            "AMD64" { $software_exec = $software_msi_x64 }
            "x86" { $software_exec = $software_msi_x32 }
            default { Write-Warning "$software_name 無法正常安裝: 不支援的系統:  $($env:PROCESSOR_ARCHITECTURE)"; $software_exec = $null }
        }

        if ($software_exec -ne $null) {
            $argumentlist = "/i ""$($env:temp + "\" + $software_Name + "\" + $software_exec)"" /passive /log install_officescan.log"
            if ($check_admin) {
                $proc = Start-Process -FilePath "msiexec.exe" -ArgumentList $argumentlist -PassThru
            } else {
                $proc = Start-Process -FilePath "msiexec.exe" -ArgumentList $argumentlist -Credential $credential -PassThru
            }
            $proc.WaitForExit()
            Start-Sleep -Seconds 1 
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
    
    install-AntiVir
    pause
}