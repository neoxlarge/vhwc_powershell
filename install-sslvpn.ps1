# �w��SSLVPN�γn��
# SMA connect Ageng �M Net Extender

param($runadmin)

Import-Module ((Split-Path $PSCommandPath) + "\get-installedprogramlist.psm1")


function install-SMAConnectAgent {


    $software_name = "SMA Connect Agent"
    $software_path = "\\172.20.1.122\share\software\00newpc\34-SSLVPN"
    $software_msi_x64 = "SMAConnectAgent.msi"  #64bit 32bit ���P�@��
    $software_msi_x32 = "SMAConnectAgent.msi"

    $Username = "vhcy\vhwcmis"
    $Password = "Mis20190610"
    $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)

    $all_installed_program = get-installedprogramlist
      
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }

    if ($software_is_installed -eq $null) {
    
        Write-Output "Start to insall: $software_name"

        #�_���ɮר�temp
        $software_path = get-item -Path $software_path
                
        #copy-item �L�k���{��, ���n�qpsdrive��, �ҥH�n��driver.
        $net_driver = "vhwcdrive" #�u�O����driver�W�r�Ӥv.
        New-PSDrive -Name $net_driver -Root $software_path -PSProvider FileSystem -Credential $credential
        Copy-Item -Path "$($net_driver):\" -Destination $env:TEMP -Recurse -Force -Verbose 
        Remove-PSDrive -Name $net_driver


        ## �P�_OS�O32(x86)�άO64(AMD64), ��L��(ARM64)���w��  
        switch ($env:PROCESSOR_ARCHITECTURE) {
            "AMD64" { $software_exec = $software_msi_x64 }
            "x86" { $software_exec = $software_msi_x32 }
            default { Write-OutPut "Unsupport CPU or OS:"  $env:PROCESSOR_ARCHITECTURE; $software_exec = $null }
        }

        if ($software_exec -ne $null) {
            Start-Process -FilePath "msiexec.exe" -ArgumentList "/i $($env:temp + "\" + $software_path.Name + "\" + $software_exec) /passive /log install_officescan.log" -Wait
            Start-Sleep -Seconds 5 
        }
        else {
            Write-Warning "$software_name �L�k���`�w��."
        }
      
             
        #�w�˧�, �A���s���o�w�˸�T
        $all_installed_program = get-installedprogramlist
        $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }
    
    }

    Write-Output ("Software has installed: " + $software_is_installed.DisplayName)
    Write-Output ("Version: " + $software_is_installed.DisplayVersion)

}

function install-NX {
    #

    $software_name = "SonicWall NetExtender"
    $software_path = "\\172.20.1.122\share\software\00newpc\34-SSLVPN"
    $software_msi_x64 = "NXSetupU.exe"  #64bit 32bit ���P�@��
    $software_msi_x32 = "NXSetupU.exe"

    $Username = "vhcy\vhwcmis"
    $Password = "Mis20190610"
    $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)

    $all_installed_program = get-installedprogramlist
      
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }

    if ($software_is_installed -eq $null) {
    
        Write-Output "Start to insall: $software_name"

        #�_���ɮר�temp
        $software_path = get-item -Path $software_path
                
        #copy-item �L�k���{��, ���n�qpsdrive��, �ҥH�n��driver.
        $net_driver = "vhwcdrive" #�u�O����driver�W�r�Ӥv.
        New-PSDrive -Name $net_driver -Root $software_path -PSProvider FileSystem -Credential $credential
        Copy-Item -Path "$($net_driver):\" -Destination $env:TEMP -Recurse -Force -Verbose 
        Remove-PSDrive -Name $net_driver


        ## �P�_OS�O32(x86)�άO64(AMD64), ��L��(ARM64)���w��  
        switch ($env:PROCESSOR_ARCHITECTURE) {
            "AMD64" { $software_exec = $software_msi_x64 }
            "x86" { $software_exec = $software_msi_x32 }
            default { Write-OutPut "Unsupport CPU or OS:"  $env:PROCESSOR_ARCHITECTURE; $software_exec = $null }
        }

        if ($software_exec -ne $null) {
            Start-Process -FilePath "$($env:temp)\$($software_path.name)\$software_exec"    -ArgumentList "/S" -Wait
            Start-Sleep -Seconds 5 
        }
        else {
            Write-Warning "$software_name �L�k���`�w��."
        }
      
             
        #�w�˧�, �A���s���o�w�˸�T
        $all_installed_program = get-installedprogramlist
        $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }
    
    }

    Write-Output ("Software has installed: " + $software_is_installed.DisplayName)
    Write-Output ("Version: " + $software_is_installed.DisplayVersion)

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
        install-SMAConnectAgent
        install-NX
    }
    else {
        Write-Warning "�L�k���o�޲z���v���Ӧw�˳n��, �ХH�޲z���b������."
    }
    pause
}