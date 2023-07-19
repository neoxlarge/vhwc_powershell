# 安裝SmartIris

param($runadmin)

Import-Module ((Split-Path $PSCommandPath) + "\get-installedprogramlist.psm1")


function Parse-IniFile {
    <#
    code from chatgpt
    這個函數使用了 PowerShell 的進階參數處理機制來處理路徑參數，並在讀取 .ini 檔案的每一行時進行適當的解析。函數還使用了註解來說明每個部分的功能，讓程式碼更易於理解和維護。
    #>
    
        # 使用 [CmdletBinding()] 屬性來開啟進階的參數處理
        # 這允許在函數中使用進階參數，例如 Mandatory、ParameterSetName 等
        [CmdletBinding()]
        param (
            # 使用 [Parameter()] 屬性來指定必要的路徑參數
            [Parameter(Mandatory = $true)]
            [string]$Path
        )
     
        # 確保檔案存在
        if (-not (Test-Path $Path)) {
            throw "The file '$Path' does not exist."
        }
     
        # 初始化一個空的雜湊表，用於存儲解析後的 .ini 內容
        $ini = @{}
     
        # 初始化一個空的節點名稱變數，用於解析目前的節點
        $section = ""
     
        # 讀取 .ini 檔案中的每一行
        Get-Content $Path | ForEach-Object {
            # 去除每行前後的空格
            $line = $_.Trim()
     
            # 如果該行是節點名稱，則解析出節點名稱並初始化一個新的雜湊表
            if ($line -match "^\[.*\]$") {
                $section = $line.Substring(1, $line.Length - 2)
                $ini[$section] = @{}
            }
            # 如果該行是鍵值對，則解析出鍵和值，並將其存入目前節點的雜湊表中
            elseif ($line -match "^([^=]+)=(.*)$") {
                $key = $matches[1].Trim()
                $value = $matches[2].Trim()
                $ini[$section][$key] = $value
            }
        }
     
        # 返回解析後的 .ini 內容雜湊表
        return $ini
    
    }
    

    function  save-iniFile {

        param (
            [CmdletBinding()]
            $ini,
            [CmdletBinding()]
            $path
        )
    
        $ini_content = ""
    
        foreach ($i in $ini.keys) {
            $ini_content += "[$i] `n"
            
            foreach ($j in $ini.$i.keys) {
                $ini_content += "$j=$($ini.$i.$j.tostring()) `n"
            }
        }
    
        Write-Output $ini_content
        Out-File -InputObject $ini_content -FilePath $path
    }
    
    
    

function check-SmartIris {

    $ini_path = "C:\temp\TEDPC\SmartIris\UltraQuery\SysIni\LocalSetting.ini"

    if (Test-Path -Path $ini_path) {
        $ini = Parse-IniFile -Path $ini_path

        $ini.LocalSetting.AETitle = $env:COMPUTERNAME
        
        save-iniFile -ini $ini -path $ini_path
    }

}

function install-SmartIris {

    
    $software_name = "SmartIris"
    $software_path = "\\172.20.5.187\mis\02-SmartIris\SmartIris_V1.3.6.4_Beta7_UQ-1.1.0.19_R2_Install_20200701"
    $software_exe = "setup.exe"
    ## 找出軟體是否己安裝

    $all_installed_program = get-installedprogramlist
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }

    if ($null -eq $software_is_installed) {
    
        Write-Output "Start to insall: $software_name"

        #復制檔案到本機暫存"
        $software_path = get-item -Path $software_path
        Copy-Item -Path $software_path -Destination $env:temp -Recurse -Force -Verbose

        # 安裝  
        $runid = Start-Process -FilePath $($env:temp + "\" + $software_path.Name + "\" + $software_exe) -ArgumentList "/s /f1$($env:temp + "\" + $software_path.Name + "\vhwc.iss")" -PassThru
        
        # 安裝過程中, 這2支程會跳出來, 先砍了, 等會再設定即可.
        while (!($runid.HasExited)) {
            get-process -Name MonitorCfg -ErrorAction SilentlyContinue | Stop-Process
            get-process -Name UQ_Setting -ErrorAction SilentlyContinue | Stop-Process
            Start-Sleep -Seconds 1
        }

        #復制設定檔到本機.
        copy-item -Path "\\172.20.5.187\mis\02-SmartIris\vhwc_UltraQuery_SysIni\*" -Destination "C:\TEDPC\SmartIris\UltraQuery\SysIni" -Force
 

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
        install-SmartIris    
    }
    else {
        Write-Warning "無法取得管理員權限來安裝軟體, 請以管理員帳號重試."
    }
    pause
}