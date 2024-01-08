<# 20240108
剛問健保署窗口，因為健保署vpn查雲端藥歷的網頁，如果有問題，會去try COMX1-COMX10，
並寫入至C:\NHI\INI\csdll.ini，所以會有早上有的診間有問題，有的沒有，所以如果二院都是確認在COMX1，
健保署建議將C:\NHI\INI\csdll.ini設為唯讀，如果就不會發生早上的問題，所以 @沂圻 @隆昌 確認一下，
是否先將門診之電腦該檔案皆設為唯讀，避免被VPN改掉。另外如設唯讀後在控制程式的修訂如有異動檔案屬性，要注意記得改回來。
#>

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

function check_comx1 ($ini_path) {
    
    $log_file = "\\172.20.1.14\update\0001-中榮系統環境設定\set_csdllini.log"

    #電腦名稱限制條件
    $rule = "wmis-*"
    if ($env:COMPUTERNAME -like $rule) {
        $is_computername_rule = $true
    }
    else {
        $is_computername_rule = $false
    }

    #檢查電腦名稱是否符合rule 及檢查ini路徑, 如果ini不存在也不用跑
       
    if (Test-Path $ini_path -and $is_computername_rule) {

        #取得ini內容
        $ini_content = Parse-IniFile -Path $ini_path 
        #取得ini檔案屬性
        $ini_fileproperty_readonly = Get-ItemPropertyValue -Path $ini_path -Name "IsReadOnly"

        Write-Host "COM value is ""$($ini_content.CS.COM)"""
    
        #如果不是COMX1, 就改成COMX1
        if ($ini_content["CS"]["COM"] -ne "COMX1") {

            #如果有唯讀先解開
            if ($ini_fileproperty_readonly -eq $true) {
                Set-ItemProperty -Path $ini_path -Name IsReadOnly -Value $false -Force
                #改完再重取新得ini檔案屬性
                $ini_fileproperty_readonly = Get-ItemPropertyValue -Path $ini_path -Name "IsReadOnly"
            }

            Write-Host "Change COM value to ""COMX1""" -ForegroundColor Red
            $ini_content["CS"]["COM"] = "COMX1"

            #存檔
            save-iniFile -ini $ini_content -path $ini_path

            #寫一下log
            $log_string = "set csdll.ini COMX1: $env:COMPUTERNAME,$(Get-Date)" 
            $log_string | Add-Content -PassThru $log_file

        }

        #不是唯讀的都改唯讀
        if ($ini_fileproperty_readonly -eq $false) {
            Set-ItemProperty -Path $ini_path -Name IsReadOnly -Value $true -Force

            #寫一下log
            $log_string = "set csdll.ini readonly: $env:COMPUTERNAME,$(Get-Date)" 
            $log_string | Add-Content -PassThru $log_file

        }
    }
}

$ini_path = "C:\nhi\ini\csdll.ini"
check_comx1 -ini_path $ini_path