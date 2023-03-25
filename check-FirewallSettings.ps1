param($runadmin)

function import-module_func ($name) {
    #此function會檢查本機上是否有要載入的模組. 如果沒有, 就連線到wcdc2.vhcy.gov.tw上下載. 可能Win7沒有內建該模組. 
    $result = get-module -ListAvailable $name

    $Username = "vhwcmis"
    $Password = "Mis20190610"
    $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)

    if ($result -ne $null) {

        Import-Module -Name $name -ErrorAction Stop

    }
    else {

        $rsession = New-PSSession -ComputerName wcdc2.vhcy.gov.tw -Credential $credential
        Import-Module $name -PSSession $rsession -ErrorAction Stop
    }
    
}


function check-FirewallSettings {
    [CmdletBinding()]
    param()

    #載入模組 NetSecurity
    import-module_func NetSecurity
    
    Write-Output "檢查防火牆啟用狀態."    
    # 取得Windows防火牆物件
    $firewall = Get-NetFirewallProfile

    # 取得防火牆軟體名稱
    $fwName = $firewall.Name

    # 檢查不同網路類型的防火牆狀態
    $domainProfile = $firewall | Where-Object { $_.Name -eq 'Domain' }
    $privateProfile = $firewall | Where-Object { $_.Name -eq 'Private' }
    $publicProfile = $firewall | Where-Object { $_.Name -eq 'Public' }

    # 檢查每種網路類型的防火牆是否已啟用
    $domainEnabled = $domainProfile.Enabled
    $privateEnabled = $privateProfile.Enabled
    $publicEnabled = $publicProfile.Enabled

    # 輸出防火牆相關資訊
    #Write-Output "防火牆軟體名稱：$fwName"
    Write-Output "網域網路防火牆是否已啟用：$domainEnabled"
    Write-Output "私人網路防火牆是否已啟用：$privateEnabled"
    Write-Output "公用網路防火牆是否已啟用：$publicEnabled"

    # 檢查防火牆是否已在所有網路類型中啟用
    if ($domainEnabled -and $privateEnabled -and $publicEnabled) {
        Write-Output "防火牆已在所有網路類型中啟用。"
    }
    else {
        Write-Warning "防火牆未在所有網路類型中啟用。"
    }

    #檢查firewall中是否充許軟體通過.
    Write-Output "檢查firewall中是否充許軟體通過."

    #要檢查firewall中的軟體是否充許的關鍵字
    $Applications = @("vnc", "chrome", "edge")

    if ($check_admin) {
    
        foreach ($app in $Applications) {
            #.Enabled 1=True, 2=False, 比對true的話,在Win7中會失效. 比對1, 都可以成功.
            $appRules = Get-NetFirewallRule | Where-Object { $_.Enabled -eq 1 -and $_.DisplayName -like "*$App*" }
            Write-Output "檢查防火牆允許軟體: $app"
            if ($appRules) {
                $appRules | format-table -Property DisplayName, Enabled, Profile, Direction, Action
            }
            else {
                Write-Warning "未發現充許 $app 的設定."    
            }
        }
    }
    else {
        Write-Warning "沒有系統管理員權限,無法檢查允許軟體,請以系統管理員身分重新嘗試."
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

    check-FirewallSettings
    
    pause
}
