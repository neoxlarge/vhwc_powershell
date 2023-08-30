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
        Disconnect-PSSession -Session $rsession | Out-Null
    }
    
}


function check-FirewallSettings {
    [CmdletBinding()]
    param()

    if ($check_admin) {

        #載入模組 NetSecurity
        import-module_func NetSecurity
    
        Write-Output "檢查防火牆啟用狀態."    
        # 取得Windows防火牆物件
        $firewallProfiles = Get-NetFirewallProfile

        foreach ($profile in $firewallProfiles) {
            if ($($profile.Enabled)) {
                Write-output "配置檔案：$($profile.Name) : $($profile.Enabled)" 
            }
            else {
                write-warning "配置檔案：$($profile.Name) : $($profile.Enabled)" 
            }
        
        }


        #檢查firewall中是否充許軟體通過.
        Write-Output "檢查firewall中是否充許軟體通過."

        #要檢查firewall中的軟體是否充許的關鍵字
        $Applications = @("vnc", "chrome", "edge", "bbb")

    
    
        foreach ($app in $Applications) {
            $allowed = $false
        
            $firewallRules = Get-NetFirewallRule
        
            foreach ($rule in $firewallRules) {
                if ($rule.Enabled -and $rule.Action -eq 'Allow' -and $rule.DisplayName -like "*$app*") {
                    $allowed = $true
                    break
                }
            }
        
            if ($allowed) {
                Write-output "防火牆允許應用程式 '$app' 通過。"
            }
            else {
                Write-Warning "防火牆不允許應用程式 '$app' 通過。" 
            }
        }

        write-output "檢查是否回應Ping:"

        $icmpRule = Get-NetFirewallRule | Where-Object { $_.Name -like "*CoreNet-Diag-ICMP?-EchoRequest-In*" }
        #系統預設的應該有個, 全都打開.
    
        foreach ($i in $icmpRule) {
    
            if ($i.enabled -eq $false) {
                Write-Output "進行啟用firewall rule: $($i.Displayname)"
                $i | Set-NetFirewallRule -Enabled $true
            }
            else {
                Write-Output "已啟用: $($i.DisplayName)"
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
