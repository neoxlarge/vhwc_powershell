<#  �w��31-�w���f�Ľ]�d�ߨt��
    * ���w�˵{�����|�b�����C���d�O��. �b���|���T�{���L�w��.
    * �ୱ���|��W.
#>


param($runadmin)

function install-cdcalert {


    $software_name = "�w���f�Ľ]�d�ߨt��"
    $software_path = "\\172.20.5.187\mis\31-�w���f�Ľ]�d�ߨt��\cdcClinic"
    $software_msi_x64 = "cdcalert.msi"  #64bit 32bit ���P�@��
    $software_msi_x32 = "cdcalert.msi"

    $software_installed = "C:\Program Files (x86)\Changingtec\cdcClinic\cdcalert.exe"

    #�Ψӳs�u172.20.1.112���{��
    $Username = "vhcy\vhwcmis"
    $Password = "Mis20190610"
    $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)

    $check_installedpath = Test-Path -Path $software_installed

    if ($check_installedpath -eq $false) {
        
        Write-Output "Start to install: $software_name"

        #�_���ɮר�temp
        $software_path = get-item -Path $software_path
                
        #copy-item �L�k���{��, ���n�qpsdrive��, �ҥH�n��driver.
        $net_driver = "vhwcdrive" #�u�O����driver�W�r�Ӥv.
        #�������|
        New-PSDrive -Name $net_driver -Root $software_path -PSProvider FileSystem -Credential $credential
        #�_��
        Copy-Item -Path "$($net_driver):\" -Destination "$($env:TEMP)\$($software_path.Name)" -Recurse -Force -Verbose 
        #unmount���|
        Remove-PSDrive -Name $net_driver

        ## �P�_OS�O32(x86)�άO64(AMD64), ��L��(ARM64)���w��  
        switch ($env:PROCESSOR_ARCHITECTURE) {
            "AMD64" { $software_exec = $software_msi_x64 }
            "x86" { $software_exec = $software_msi_x32 }
            default { Write-OutPut "Unsupport CPU or OS:"  $env:PROCESSOR_ARCHITECTURE; $software_exec = $null }
        }

        Start-Process -FilePath msiexec.exe Get-NetFirewallRule " /i $($env:TEMP)\$($software_path.Name)\$software_exec /passive" -Wait

        Start-Sleep -Seconds 2

        $software_property = Get-ItemProperty -Path $software_installed 

    }
    else {
        $software_property = Get-ItemProperty -Path $software_installed   
    }

    Write-Output ("Software has installed: " + $software_name)
    Write-Output ("Version: " + $software_property.versioninfo)

    #��W���|
    
    if (Test-Path "$($env:PUBLIC)\desktop\cdcalert.link") {
        Rename-Item -Path "$($env:PUBLIC)\desktop\cdcalert.link" -NewName "$software_name.lnk"
        }



}



#�ɮ׿W�߰���ɷ|����禡, �p�G�O�Q���J�ɤ��|����禡.
if ($run_main -eq $null) {

    #�ˬd�O�_�޲z��
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #�p�G�D�޲z��, �N�յ�run as admin, �öǤJrunadmin �Ѽ�1. �]���b����@��ϥΪ̥û������O�޲z���v��, �|�y���L�����]. ���ѼƥΨӻ��U�P�_�u�]�@��. 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    
    }

    if ($check_admin) { 
        install-cdcalert
    }
    else {
        Write-Warning "�L�k���o�޲z���v���Ӧw�˳n��, �ХH�޲z���b������."
    }
    pause
}