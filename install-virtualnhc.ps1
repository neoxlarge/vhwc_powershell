#�ɯŵ������O�dSDK 2.5.4

param($runadmin)

Import-Module ((Split-Path $PSCommandPath) + "\get-installedprogramlist.psm1")


#���oOS������
function Get-OSVersion {
    $os = (Get-WmiObject -Class Win32_OperatingSystem).Caption

    if ($os -like "*Windows 7*") {
        return "Windows 7"
    }
    elseif ($os -like "*Windows 10*") {
        return "Windows 10"
    }
    else {
        return "Unknown OS"
    }
}

function Get-IPv4Address {
    <#
    �^�ǧ�쪺IP,�u��b172.*�~���. 
    #>

    $ip = Get-WmiObject -Class Win32_NetworkAdapterConfiguration |
    Where-Object { $_.IPAddress -ne $null -and $_.IPAddress[0] -like "172.20.*" } |
    Select-Object -ExpandProperty IPAddress |
    Where-Object { $_ -match "\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}" } |
    Select-Object -First 1

    if ($ip -eq $null) {
        return $null
    }
    else {     
        return $ip
    }
}


function Create-Shortcut {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TargetPath,

        [Parameter(Mandatory = $true)]
        [string]$ShortcutPath
    )

    # �Ы� WScript.Shell ����
    $shell = New-Object -ComObject WScript.Shell

    # �ˬd�ֱ��覡�O�_�w�s�b�A�p�G�s�b��R��
    if ($shell.CreateShortcut($ShortcutPath).FullName) {
        Remove-Item $ShortcutPath -Force -ErrorAction SilentlyContinue
    }

    # �Ыاֱ�
    $shortcut = $shell.CreateShortcut($ShortcutPath)
    $shortcut.TargetPath = $TargetPath
    $shortcut.Save()
    
}


function check-OPDList {
    # �ˬd�ثe�q���W�٬O�_�b�ɯŦW��(opd_list.json)��. ������,�^�Ǹ��.
   
    # �N�ܼ� $json_file �]�w�� JSON �ɮ� "opd_list.json" �����|
    $path = get-item -Path $PSCommandPath
    $json_file = "$($Path.DirectoryName)\opd_list.json"
    #Write-Output $json_file
    
    # �q���W��
    $pc_name = $env:COMPUTERNAME

    # Ū�� JSON �ɮת����e�ñN���x�s�b�ܼ� $json_content ��
    $json_content = Get-Content -Path $json_file -Raw

    # �N JSON ���e�ഫ�� PowerShell ����ñN��������ܼ� $opd_json
    $opd_json = ConvertFrom-Json -InputObject $json_content

    # ��l���ܼ� $opd�A�ñN��]�w�� null
    $opd = $null

    # ��X�ŦX�q���W�٪����.
    foreach ($o in $opd_json.psobject.properties) {

        $result = $o.Value.name -eq $pc_name
   
        if ($result) {
            Write-Output $o.Value
            return $o.Value
            break
        }
 
    }
}

function install-virtualnhc {
    #�ɯŵ������O�dSDK 2.5.4

    #�ɯű���:
    # 1. �E����͹q���~�ݭn�w��.
    # 2. �E����͹q���ثe�Ҭ�win10.
    # �ҥH�n�ѦҦw�˦C��opd_list.json 
    
    # �s�����������O�dSDK 2.5.4���Φw��, �u�n�_���Ƨ��A����{���Y�i

    # �ˬd�ثe�q���W�٬O�_�b�ɯŦW��(opd_list.json)��.
    # ���Ucheck-OPDList��powershell v5 �y�k,�L�k�bwin7������. �N�����w�bwin10����.
    
    if ($(Get-OSVersion) -eq "Windows 10") {

        $ipv4 = Get-IPv4Address 

        $opd = check-OPDList

        #if ($opd.virturl_NHIcard) {
        if ($true) {    
            #�p�Gvirtual_NHIcar�ȬO$true ��ܥi�ɯ�.
            #�o��2.5.4, ���n�P�_�O�_�v�g�w��, �N�����A�ˤ@��.
            #�E����v�q�������̰��v��, �b���N�Ȥ��Ҽ{�v�����D. �@��ϥΪ̷|�L�k�ϥ�.

            #���o�ª��n��
            $installed_vhc = get-installedprogramlist | Where-Object -FilterScript { $_.Displayname -like "�������O�d*" }

            $log_file = "\\172.20.1.14\update\0001-���a�t�����ҳ]�w\VirtualHC_254.log"
            
            
            #�����ª��n��
            if ($installed_vhc -ne $null) {
                
                $unistalll_strign = $installed_vhc.QuietUninstallString.Split("""")
                Start-Process -FilePath $unistalll_strign[1] -ArgumentList $unistalll_strign[2] -Wait -ErrorAction SilentlyContinue -NoNewWindow
                #$installed_vhc.uninstall()
                Start-Sleep -Seconds 3
            }


            #�ˬd�s���O�_�v�g�_��2.5.4���쥻��
            $vhc_path = "\\172.20.5.187\mis\25-�������O�d\�E��\VHIC_virtual-nhicard+SDK+Setup-2.5.4"
            $vhc_path = Get-Item $vhc_path

            Copy-Item -Path $vhc_path -Destination "c:\NHI\$($vhc_path.name)" -Recurse -Force -Verbose
                            
            #�_��|��ୱ�αҰ�
            $target_path = "C:\users\public\desktop\�������O�d����n��.lnk"
            $result = !(Test-Path -Path $target_path) -and $check_admin
            if ($result) {
            Create-Shortcut -TargetPath "C:\NHI\VHIC_virtual-nhicard+SDK+Setup-2.5.4\�������O�d����n��-������.2.5.4.exe" -ShortcutPath $target_path
            #�Ȯɮ����Ұʪ�, ���O�Ҧ��q�����ݭn�]
            #Create-Shortcut -TargetPath "C:\NHI\VHIC_virtual-nhicard+SDK+Setup-2.5.4\�������O�d����n��-������.2.5.4.exe" -ShortcutPath "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\�������O�d����n��.lnk"
            }

            #open firewall
            #netsh advfirewall firewall add rule name='Allow �������O�d����n��' dir=in action=allow program='C:\NHI\VHIC_virtual-nhicard+SDK+Setup-2.5.4\�������O�d����n��-������.2.5.4.exe'
            $exe_path1 = "$($env:USERPROFILE)\AppData\Local\Temp\C\vNHI\�������O�d����n��.exe"
            $exe_path2 = "$($env:USERPROFILE)\AppData\Local\Temp\C\vNHI\resources\app\VHCNHI_Slient\VHCNHI_Slient.exe"

            $rule_name1 = "Allow �������O�d����n�� ($($env:USERNAME) main) "
            $rule_name2 = "Allow �������O�d����n�� ($($env:USERNAME) VHCNHI_Silent)"

            #�޲z���v��vhwcmis���Ү�.
            $Username = "vhwcmis"
            $Password = "Mis20190610"
            $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
            $credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)
            
            <#
            $cimsession = New-CimSession -ComputerName $env:COMPUTERNAME -Credential $credential

            if (Get-NetFirewallRule -DisplayName $rule_name1 -CimSession $cimsession -ErrorAction SilentlyContinue) {
                Remove-NetFirewallRule -DisplayName -CimSession $cimsession $rule_name1
            }
            if (Get-NetFirewallRule -DisplayName $rule_name2 -CimSession $cimsession -ErrorAction SilentlyContinue) {
                Remove-NetFirewallRule -DisplayName $rule_name2 -CimSession $cimsession
            }

            $arg1 = "advfirewall firewall add rule name=""$rule_name1 "" dir=in action=allow program=""$exe_path1"""
            Write-Output  "�إ�firewall rule: $arg1"
            #Start-Process netsh.exe -ArgumentList $arg1 -NoNewWindow -wait -Credential $credential
            Invoke-Command -ComputerName $env:COMPUTERNAME -Credential $credential -ScriptBlock {
                param ($argument1)
                netsh.exe $argument1
            } -ArgumentList $arg1



            $arg2 = "advfirewall firewall add rule name=""$rule_name2"" dir=in action=allow program=""$exe_path2"""
            write-output "�إ�firewall rule: $arg2"
            #Start-Process netsh.exe -ArgumentList $arg2 -NoNewWindow -wait -Credential $credential
            Invoke-Command -ComputerName $env:COMPUTERNAME -Credential $credential -ScriptBlock {
                param ($argument2)
                netsh.exe $argument2
            } -ArgumentList $arg2

#>

            $log_string = "update V-NHICard,$env:COMPUTERNAME,$ipv4,$(Get-OSVersion),$env:PROCESSOR_ARCHITECTURE,$(Get-Date)" 
            $log_string | Add-Content -PassThru $log_file

        }
    
            
        
    }
    

}


#�ɮ׿W�߰���ɷ|����禡, �p�G�O�Q�פJ�ɤ��|����禡.
if ($run_main -eq $null) {

    #�ˬd�O�_�޲z��
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #�p�G�D�޲z��, �N�յ�run as admin, �öǤJrunadmin �Ѽ�1. �]���b����@��ϥΪ̥û������O�޲z���v��, �|�y���L�����]. ���ѼƥΨӻ��U�P�_�u�]�@��. 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    
    }

    if ($check_admin) { 
        install-virtualnhc
    }
    else {
        Write-Warning "�L�k���o�޲z���v���Ӧw�˳n��, �ХH�޲z���b������."
    }
    pause
}