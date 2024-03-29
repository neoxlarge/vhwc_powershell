﻿#Check-SmartIris
#此檔執行以下動作. 
#1. 寫入AEtitle, 值是目前的電腦名稱. 
#
#底下設定可以不用修改.
#1. Localsetting.ini中的ip,程式執行時會自動抓取， 不用手動設定。 
#2. Serverlist.ini中的主機列表, 正常不會變動. 不用再設定. 
#
#20230809, HIS call PACS, 112年新電腦無法正確執行, 沂圻找到正確並修正ultraquery.ini後,正常, 但影响該功能的設定值不明. 暫無獨立檢查該設定的方法.


param($runadmin)

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
    #將ini寫回檔案
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
    #Write-Output $ini_content
    Out-File -InputObject $ini_content -FilePath $path
}
    
    
function check-SmartIris {

    $ini_path1 = "C:\TEDPC\SmartIris\UltraQuery\SysIni\LocalSetting.ini"
    $ini_path2 = "C:\TEDPC\SmartIris\UltraQuery\SysIni\UltraQuery.ini"

    
    Write-output "設定UltraQuery."

    #依照 \\172.19.1.14\Update\資訊室\共用程式\SmartIris\SmartIris更版SOP.doc
    #步驟三：
    #\\172.19.1.14\Update\資訊室\共用程式\SmartIris\UltraQuery_V1.1.1.0_Update_20200731
    # 全選複製到C:\TEDPC\SmartIris\UltraQuery並覆蓋

    copy-item -Path "\\172.20.5.187\mis\02-SmartIris\UltraQuery_V1.1.1.0_Update_20200731\*" -Destination "C:\TEDPC\SmartIris\UltraQuery" -Recurse -Force

    #復制設定檔到本機.
    copy-item -Path "\\172.20.5.187\mis\02-SmartIris\vhwc_UltraQuery_SysIni\*" -Destination "C:\TEDPC\SmartIris\UltraQuery\SysIni" -Force
 

    Write-Output "寫入電腦名稱$(($env:COMPUTERNAME).toLower())到AEtitle."

    $ini = Parse-IniFile -Path $ini_path1
        
    $ini.LocalSetting.AETitle = $($env:COMPUTERNAME).ToLower()
        
    save-iniFile -ini $ini -path $ini_path1

    #Write-Output "修改影像路徑."

    #$ini = Parse-IniFile -Path $ini_path2
        
    #$ini.PathSetting.ImageDir  = "C:\Image_Root\"

    #save-iniFile -ini $ini -path $ini_path2
    

}

#檔案獨立執行時會執行函式, 如果是被滙入時不會執行函式.
if ($run_main -eq $null) {

    #檢查是否管理員
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #如果非管理員, 就試著run as admin, 並傳入runadmin 參數1. 因為在網域一般使用者永遠拿不是管理員權限, 會造成無限重跑. 此參數用來輔助判斷只跑一次. 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    
    }

    Check-SmartIris
    
    pause
}