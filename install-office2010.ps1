param($runadmin)

Import-Module "\\172.20.5.185\powershell\vhwc_powershell\get-installedprogramlist.psm1"


function Get-IPv4Address {
    <#
    只能在172.*才能用.
    #>

    $ip = Get-WmiObject -Class Win32_NetworkAdapterConfiguration |
    Where-Object { $_.IPAddress -ne $null -and $_.IPAddress[0] -like "172.*" } |
    Select-Object -ExpandProperty IPAddress |
    Where-Object { $_ -match "\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}" } |
    Select-Object -First 1
    return $ip
}


function install-WinNexus {
    # 安裝Winnexus
    $software_name = "WinNexus"
    #$software_path = "\\172.20.5.187\mis\13-Winnexus\Winnexus_1.2.4.7\13-Winnexus"
    $software_path = "\\172.20.1.122\share\software\00newpc\13-Winnexus\Winnexus_1.2.4.7"
    $software_exec = "Install_Desktop.1.2.4.7.exe"

    $Username = "vhcy\vhwcmis"
    $Password = "Mis20190610"
    $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)

    ## 找出軟體是否己安裝

    $all_installed_program = get-installedprogramlist

   
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like "$software_name*" }

    if ($software_is_installed -eq $null) {
        Write-OutPut "Start to install: $software_name"

        #$software_path = get-item -Path $software_path
        
        #復制檔案到temp

        Copy-Item -Path $software_path -Destination "$($env:TEMP)\$software_name" -Recurse -Force -Verbose
       
        #installing...

        $run_app = Start-Process -FilePath "$($env:TEMP)\$software_Name\$software_exec" -ArgumentList ("/suppressmsgboxes /log:install_winnexus.log") -PassThru -NoNewWindow -Credential $credential
        $run_app.WaitForExit()
        Start-Sleep -Seconds 5 
   
     
        #安裝完, 再重新取得安裝資訊
        $all_installed_program = get-installedprogramlist
        $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like "$software_name *" }
    } 

    Write-Output ("Software has installed: " + $software_is_installed.DisplayName)
    Write-Output ("Version: " + $software_is_installed.DisplayVersion)



}



function install-office2010 {

    # install office 2010 

    $Username = "vhcy\vhwcmis"
    $Password = "Mis20190610"
    $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)

    #install office 2010
    $software_name = "Microsoft Office Standard 2010*"

    $all_installed_program = get-installedprogramlist
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }

    # not install office 2010 yet
    if ($software_is_installed -eq $null) {
        
        
        install-Winnexus    

        #check winnexus loader
        $app_path = 'C:\Program Files\WinNexus\Desktop\bin\WinNexusLoader.exe'
        $wdf_path = "\\172.20.5.185\powershell\vhwc_powershell\office2010_winnexus.wdf"

        Remove-Item -Path "C:\ProgramData\WinNexus\WPD\office2010_vhcy" -Recurse -Force -ErrorAction SilentlyContinue

        
        if ((Test-Path -Path $app_path) -and (Test-Path -Path $wdf_path)) {
            #use winnexus to install office 2010
            Start-Process -FilePath $app_path -ArgumentList  $wdf_path -NoNewWindow -Wait
        }
        else {
            Write-Error "Files not exist: $app_path , $wdf_path "
        }
        
        do {
            
            # 觀察winnexus的行為. 在winnexusloader啟動安裝後,有以下步?:
            # 1.下載office2010的壓縮檔 .wpd
            # 2.解壓到office2010_vhcy
            # 3.安裝中.
            # 4.裝完刪暫存, 底下office2010_vhcy\Script\install 會刪掉.
            # 結論, 資料夾存在且是空的, 即表安裝完成, 閒置中.

            Write-Host "Office 2010 in installing, please wait..."
            Start-Sleep -Seconds 10

            $path_property = Get-ItemProperty -Path "C:\ProgramData\WinNexus\WPD\office2010_vhcy\Script" -ErrorAction SilentlyContinue
           
        } until (
            $path_property.Exists -and $path_property.GetDirectories().count -eq $null

        )
        
    }

    Write-Host "Office 2010 has installed."

}

function choice-officekey {
    #預先選擇office key
    $pscompatiale = $PSVersionTable.PSCompatibleVersions | Where-Object { $_ -match "^5\.1" }
    if ($pscompatiale -ne $null) {
        $key_list = [ordered]@{}
    }
    else {
        $key_list = @{}
    }

    $key_list.Add("a", "BK63F-CDJGW-9MFY8-9J9D2-2VK3M")
    $key_list.Add("b", "27P9Q-9CMHC-DPMPV-4XDQD-BFCYT")
    $key_list.Add("c", "GR3XF-XM87M-QK2TX-6RR8Q-BPJWX")
    $key_list.Add("d", "GVH89-PKWTH-MRCTP-CMYTJ-8D4WV")
    $key_list.Add("e", "XHR87-BWT9M-W9HJG-D83V9-CPMHH")
    $key_list.Add("f", "27DWW-JQM62-4KBC2-HHFCG-VHRBB")
    $key_list.Add("g", "8W88H-RMRKC-XQ3DC-GWM6Q-H4YRX")
    $key_list.Add("h", "H6B73-3W3PD-XD3FD-V4BY7-7JF36")
    $key_list.Add("i", "HP72H-MJPQK-GWBPM-MDKTX-CBPC7")
    $key_list.Add("j", "J8JHM-KFX6C-G7RBX-47RM2-QC9C2")
    $key_list.Add("k", "J8JHM-KFX6C-G7RBX-47RM2-QC9C2")
    $key_list.Add("l", "THX9R-D6H2K-T9X66-B36FT-8YYFJ")

    # 顯示金鑰清單
    foreach ($key in $key_list.Keys) {
        Write-Host "$key. $($key_list[$key])"
    }
    
    # 提示使用者選擇
    if ($pscompatiale -eq $null) {
        Write-Host "Win7系統請小心字母順序" -ForegroundColor Red
    }
    $userChoice = Read-Host "請選擇一個Key（輸入對應的字母）"

    return $key_list["$userChoice"]

}

function active-office ($key) {
    Write-Host "Activing office 2010."

    $Username = "vhcy\vhwcmis"
    $Password = "Mis20190610"
    $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)

    switch ($env:PROCESSOR_ARCHITECTURE) {
        "AMD64" { $officeospp_path = "C:\Program Files (x86)\Microsoft Office\Office14\OSPP.VBS" }
        "x86" { $officeospp_path = "C:\Program Files\Microsoft Office\Office14\OSPP.VBS" }
    
    }

    $run_ospp = Start-Process -FilePath cscript.exe -ArgumentList """$officeospp_path"" /dstatus" -NoNewWindow -Credential $credential -PassThru -RedirectStandardOutput "$($ENV:TEMP)\officehasactived.txt"
    $run_ospp.waitforexit() 

    $get_log = get-content -Path "$($ENV:TEMP)\officehasactived.txt"

    $check_hasactived = $get_log | Select-String -Pattern "---LICENSED---"

    if ($check_hasactived) {
        Write-Output $get_log
        Write-Output "己經啟用過office 2010了."
    }
    else {

        $result = Start-Process -FilePath cscript.exe -ArgumentList """$officeospp_path"" /inpkey:""$key""" -NoNewWindow -Credential $credential -PassThru -RedirectStandardOutput "$($ENV:TEMP)\officeospp.txt"

        $result.WaitForExit()
        Get-Content -Path "$($ENV:TEMP)\officeospp.txt"
        Start-Sleep -s 3
    
        for ($i = 0; $i -lt 3; $i++) {
            #先排除網路問題
            ipconfig /flushdns
            #ipconfig /renew
    
            $result_act = Start-Process -FilePath cscript.exe -ArgumentList """$officeospp_path"" /act" -NoNewWindow -Credential $credential -PassThru -RedirectStandardOutput "$($ENV:TEMP)\officeact.txt"
            $result_act.WaitForExit()
            $log = Get-Content -Path "$($ENV:TEMP)\officeact.txt"
            $check_active = $log | Select-String -Pattern "Product activation successful"
    
            Write-Host $log
    
            Start-Sleep -Seconds 1
    
            if ($check_active) {
                $save_tofile = "$env:COMPUTERNAME $(Get-IPv4Address) $office_key $(get-date)"
                Out-File -FilePath \\172.20.5.185\powershell\active_office.log -Append -InputObject $save_tofile
                break
            }
    
        }

    }
}

#檔??????????????????????, 如????被??入??????執行函??
if ($run_main -eq $null) {

    #檢查??否管????
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #如????管??員, 就試??run as admin, 並傳??runadmin ??數1. ??為??網??????使??者永??拿不是管???????? ??造??????????. 此????用來????判??只跑???? 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    }

    $key = choice-officekey
    install-office2010
    active-office -key $key
    start-sleep -Seconds 10
    
}
