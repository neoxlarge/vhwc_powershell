#check-firewallport
#Test-NetConnection 在模組 NetTCPIP, 所以須滙入讓模組.

param($runadmin)

function import-module_func ($name) {
#此function會檢查本機上是否有要載入的模組. 如果沒有, 就連線到wcdc2.vhcy.gov.tw上下載. 可能Win7沒有內建該模組. 
    $result = get-module -ListAvailable $name

    if ($result -ne $null) {

        Import-Module -Name $name -ErrorAction Stop

    } else {
        
        #管理者權限vhwcmis的證書.
        $Username = "vhwcmis"
        $Password = "Mis20190610"
        $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
        $credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)

        $rsession = New-PSSession -ComputerName wcdc2.vhcy.gov.tw -Credential $credential
        Import-Module $name -PSSession $rsession -ErrorAction Stop
        Disconnect-PSSession -Session $rsession | Out-Null
    }
    
}

function Get-IPv4Address {
    $ip = Get-WmiObject -Class Win32_NetworkAdapterConfiguration |
          Where-Object {$_.IPAddress -ne $null -and $_.IPAddress[0] -like "172.*" } |
          Select-Object -ExpandProperty IPAddress |
          Where-Object { $_ -match "\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}" } |
          Select-Object -First 1
    return $ip
}


function check-firewallport {
    param (
    $computerName = (Get-IPv4Address),
    $ports = @(5800,5900)
        )
    
    Write-Output "檢查 $computerName firewall port 是否有開啟:"

    #載入NetTCPIP模組, 
    import-module_func NetTCPIP

    #載入NetSecurity模組, 檢查firewall時會用到.
    import-module_func NetSecurity

    foreach ($port in $ports) {

        $result = Test-NetConnection -ComputerName $computerName -Port $port
        if ($result.TcpTestSucceeded) {
            Write-Output "Port $port is open."
        }
        else {
            Write-Output "Port $port is closed."
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

    check-firewallport
    pause
}