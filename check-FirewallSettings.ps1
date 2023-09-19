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
    
        Write-Output "檢查防火牆啟用狀態:"    
        # 取得Windows防火牆物件
        $firewallProfiles = Get-NetFirewallProfile

        foreach ($profile in $firewallProfiles) {
            if ($($profile.Enabled)) {
                Write-output "Profile: $($profile.Name) : $($profile.Enabled)" 
            }
            else {
                write-warning "Profile:$($profile.Name) : $($profile.Enabled)" 
            }
        
        }


        #檢查firewall中是否充許軟體通過.
        Write-Output "檢查firewall中是否充許軟體通過:"

        #要檢查firewall中的軟體的rule是否正確, profile 有domain, action 是allow.
        
        $Applications = @(
            @{"DisplayName" = "winvnc.exe" ; "Path" = "C:\Program Files\uvnc bvba\UltraVNC\winvnc.exe" },
            @{"DisplayName" = "csFSIM 5.1版" ; "Path" = "C:\nhi\bin\csfsim.exe" } ,
            @{"DisplayName" = "Allow 虛擬健保卡控制軟體"; "Path" = "C:\users\%username%\appdata\local\temp\c\vnhi\虛擬健保卡控制軟體 .exe" },
            @{"DisplayName" = "Allow 虛擬健保卡控制軟體silent"; "path" = "C:\users\%username%\appdata\local\programs\virtual-nhicard\resources\app\vhcnhi_slient\vhcnhi_slient.exe" },
            @{"DisplayName" = "IccPrj"; "Path" = "C:\iccard_his\iccprj.exe" },
            @{"DisplayName" = "UltraQuery DICOM Query/Retrieve software"; "Path" = "C:\tedpc\smartiris\ultraquery\ultraquery.exe" },
            @{"DisplayName" = "vncviewer.exe"; "Path" = "C:\Program Files\uvnc bvba\UltraVNC\vncviewer.exe" }
            
        )

        # 取得所有firewall rule.
        $firewallRules = Get-NetFirewallRule

        foreach ($app in $Applications) {
            # 先以 displayname  和 方向 過濾
            $rule = $firewallRules | Where-Object -FilterScript { $_.DisplayName -like "*$($app.DisplayName)*" `
                    -and $_.Direction -eq "Inbound" `
                                                                    
            }
            if (!$rule) {
                
                Write-warning "Firewall rule 不存在, :$($app.DisplayName)"
                
                $info = $null
                New-NetFirewallRule -DisplayName $app.DisplayName -Direction "Inbound" -Action "Allow" -program $app.Path -OutVariable info | Out-Null
                Write-Warning "新增 Firewall rule: $($info.DisplayName) $($info.Direction) $($info.action)"
                
            }
            else {
                
                # firewall rule 存在.

                $in_profle = $rule | Where-Object { $_.profile.tostring() -like "*Domain*" }
                if (!$in_profle) {
                    Write-Warning "Firewall rule: $($rule[0].DisplayName), set profile Domain, Public, Private."
                    Set-NetFirewallRule -InputObject $rule[0] -Profile Domain, Public, Private
                }
            
                $is_block = $rule | Where-Object { $_.action -eq "Block" }
                foreach ($b in $is_block) {
                    Write-Warning "Firewall rule: $($b.DisplayName), set action allow."
                    Set-NetFirewallRule -InputObject $b -Action allow
                }

                if ($in_profle -and !$is_block) {
                    Write-Output "Firewall rule: $($rule.DisplayName) , OK"
                }

            }
        }



        write-output "檢查是否回應Ping:"

        $icmpRule = Get-NetFirewallRule | Where-Object { $_.Name -like "*CoreNet-Diag-ICMP?-EchoRequest-In*" }
        #系統預設的應該有這些imcp rule, 全都打開.
    
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

function Check-FirewallPortSetting {

    [CmdletBinding()]
    param()

    if ($check_admin) {

        $port_list = @(
            5900,
            5800,
            3033    #虛擬健保卡用
        )
        $firewallrules = Get-NetFirewallRule
        $firewallports = Get-NetFirewallPortFilter -Protocol TCP 

        foreach ($port in $port_list) {

            $result = $firewallports | Where-Object -FilterScript { $_.LocalPort -contains $port }

            if ($result) {
                #rule中有port, 檢查是否有啟用且allow
                Write-Output "檢查Port: $($result.LocalPort)"
                $result_rule = $firewallrules | Where-Object -FilterScript { $_.InstanceID -eq $result.InstanceID }
                Write-Output "firewall rule 名稱:$($result_rule.DisplayName), 啟用:$($result_rule.Enabled), 方向:$($result_rule.Direction), $($result_rule.Action) "
                
            }

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

    check-FirewallSettings
    
    pause
}
