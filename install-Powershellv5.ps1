#install Powershell 5.1

param($runadmin)

function install-ps5 {
    if (($PSVersionTable.PSVersion.Major -lt 5) -and ($PSVersionTable.PSVersion.Minor -lt 1)) {
        Write-Output "Powershell 目前版本為 $($PSVersionTable.PSVersion), 升級版到到5.1"

        $software_name = "Powershell"
        $software_path = "\\172.20.5.185\powershell\powershell5.1forWin7"
        $software_msi_x64 = "Win7-KB3191566-x64.zip"
        $software_msi_x32 = "Win7-KB3191566-x86.zip"
    
        #檢查一下暫存目錄是否存在
        if (!(Test-Path -Path "$env:TEMP\$($software_path.Split("\")[-1])")) {
            switch ($env:PROCESSOR_ARCHITECTURE) {
                "AMD64" { $zip_path = "$software_path\$software_msi_x64"}
                "x86" { $zip_path = "$software_path\$software_msi_x32"}
            }

            New-Item -Path "$env:TEMP\$($software_path.Split("\")[-1])" -ItemType directory -Force
            
            Start-Process unzip.exe -ArgumentList "-o $zip_path -d $env:TEMP\$($software_path.Split("\")[-1])" -Wait -NoNewWindow

            Invoke-Expression "$env:TEMP\$($software_path.Split("\")[-1])\Install-WMF5.1.ps1 -AcceptEULA -AllowRestart" 
            
            write-output "Pause, please enter ..."
            
            $null = Read-Host
        }
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
        install-ps5    
    }
    else {
        Write-Warning "無法取得管理員權限來安裝軟體, 請以管理員帳號重試."
    }
    
}