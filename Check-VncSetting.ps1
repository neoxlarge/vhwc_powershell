#check-VncSetting
#此檔會執行2個function.
#1.檢查vnc的設定檔ultravnc.ini內的值對不對. 
#2.檢查vnc的服務是否存在及正確. 不存在會建立一個vnc服務.

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



function check-vncSetting {
<#
此函式檢查vnc的設定檔ultravnc.ini內的值對不對. 
#>
    $local_vnc_ini = "$env:ProgramFiles\uvnc bvba\UltraVNC\ultravnc.ini"
    if (Test-Path -Path $local_vnc_ini) {

        $ini = Parse-IniFile -Path $local_vnc_ini

        if (
            #密碼設定 "581BE36E7F0108A3EF" 就是 sysc0012
            $ini["ultravnc"].passwd -eq "581BE36E7F0108A3EF" -and 
            $ini["ultravnc"].passwd2 -eq "581BE36E7F0108A3EF" -and 
            #port設定
            $ini["admin"].PortNumber -eq 5900 -and 
            $ini["admin"].HTTPPortNumber -eq 5800 -and
            #port自動
            $ini["admin"].AutoPortSelect -eq 1 -and
            #Display Query Windows 取消勾選
            $ini["admin"].QuerySetting -eq 2
        ) {
        
            #全對
            Write-Output "檢查VNC設定正確."

        } else {
            
            #有誤, 使用預設檔案來取代
            if ($check_admin) {
                $default_vnc_ini = "\\172.20.5.187\mis\08-VNC\1_2_24\ultravnc.ini"
                Write-Warning "VNC設定可能有誤,使用預設值取代. 預設檔來源: $default_vnc_ini "
        
                New-PSDrive -Name vnc -PSProvider FileSystem -Root ($env:ProgramFiles + "\uvnc bvba\UltraVNC\") -Credential $credential
                Copy-Item -Path $default_vnc_ini -Destination vnc: -Force 
                Write-Output ("取代檔案到:" + ($env:ProgramFiles + "\uvnc bvba\UltraVNC\"))
                Remove-PSDrive -Name vnc
            } else {
        
                Write-Warning "沒有系統管理員權限,檢查VNC設定可能有誤,請以系統管理員身分重新嘗試."
            }
        
        
        
        }


    } else {
        Write-Warning "VNC 設定檔不存在,請檢查 $env:ProgramFiles\uvnc bvba\UltraVNC\ultravnc.ini"
    } 


}


function Check-VncService {
<#
此函式檢查vnc的服務是否存在及正確. 不存在會建立一個vnc服務.
#>

    $serviceName = 'uvnc_service'
    $serviceDisplayName = 'uvnc_service'
    $serviceDescription = 'UltraVNC 遠端桌面伺服器'
    $exePath = '"C:\Program Files\uvnc bvba\UltraVNC\winvnc.exe" -service'

    # 檢查服務是否存在
    $service = Get-Service $serviceName -ErrorAction SilentlyContinue
    Write-Output "檢查VNC服務設定:"

    if ($service -ne $null) {
        #vnc服務存在, 檢查服務設定.
        $service | Format-Table -Property Name, Status, Starttype 
        
        #VNC服務啟動類型不正確
        if ($service.StartType -ne "Automatic") {
            Write-Warning "VNC服務啟動類型不正確:"
            if ($check_admin) {
                $service.StartType = "Automatic"
            } else {
                Write-Warning "沒有系統管理員權限,無法修改VNC服務啟動類型為Automatic"
            }
        }
        
        #VNC服務非執行中
        if ($service.status -ne "running") {
            Write-Warning "VNC服務非執行中:"
            if ($check_admin) {
                Write-Output "嘗試創建 $serviceName 服務, 請檢查服務是否正常."
                $service.Start()
            } else {
                Write-Warning "沒有系統管理員權限,無法啟動VNC服務為執行中."
            }
        }


    } else {

        if ($check_admin) {
            #vnc服務不存在, 建立一個
            Write-Warning "VNC服務不存在, 已嘗試創建 $serviceName 服務, 請檢查服務是否正常."
            $params = @{
                'Name' = $serviceName
                'BinaryPathName' = $exePath
                'DisplayName' = $serviceDisplayName
                'Description' = $serviceDescription
                'StartupType' = 'Automatic'
            }
            New-Service @params | Out-Null
            
        } else {
            Write-Warning "沒有系統管理員權限,無法建立VNC服務,請以系統管理員身分重新嘗試."
            
        }
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

    Check-VncSetting
    Check-VncService

    pause
}
