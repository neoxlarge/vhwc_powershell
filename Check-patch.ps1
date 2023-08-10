#修正

param($runadmin)

#管理者權限vhwcmis的證書.
$Username = "vhwcmis"
$Password = "Mis20190610"
$securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)

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


function check-patch {

    # 1.修正重新copy SAM檔.
    #copy灣橋SAM檔-至指定位置
    $sam_path1 = "C:\VGHTC\00_mis\CMS_CS5.1.5.5-讀卡機控制軟體\0640140012001000005984.SAM"
    $sam_path2 = "C:\NHI\SAM\COMX1\0640140012001000005984.SAM"
    Copy-Item -Path $sam_path1 -Destination $sam_path2 -Force

    # 2.移除虛擬健保卡連結, 只有診間需要預設執行.
    if ($check_admin) {
        Remove-Item -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\startup\虛擬健保卡控制軟體.lnk" -Force -ErrorAction SilentlyContinue
    }

    # 3.復制hisupdate 到預設執行
    if ($check_admin) {
        Copy-Item -Path C:\Users\Public\Desktop\HISUpdateLauncher.lnk 'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp' -Force
    }

}



#檔案獨立執行時會執行函式, 如果是被滙入時不會執行函式.
if ($run_main -eq $null) {

    #檢查是否管理員
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #如果非管理員, 就試著run as admin, 並傳入runadmin 參數1. 因為在網域一般使用者永遠拿不是管理員權限, 會造成無限重跑. 此參數用來輔助判斷只跑一次. 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    
    }

    Check-patch
    

    pause
}
