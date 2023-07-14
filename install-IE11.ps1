#install IE11

param($runadmin)

Import-Module ((Split-Path $PSCommandPath) + "\get-installedprogramlist.psm1")

function install-IE11 {

    $software_name = "IE11"
    $software_path = "\\172.20.5.187\mis\20-IE"
    $software_msi_x64 = "x64\IE11-Windows6.1-x64-zh-tw.exe"
    $software_msi_x32 = "x86\IE11-Windows6.1-x86-zh-tw.exe"


    # ���o�ثe�t�Τ� IE ������
    #���R�O�N���ձq���U�� "HKLM:\Software\Microsoft\Internet Explorer" ��m���˯� IE ������T�A�ñN��s�x�b $ieVersion �ܼƤ��C
    #�Ъ`�N�A�ϥ� svcVersion �� Version �ݩʨ��M�� IE �������CsvcVersion �ݩʾA�Ω� IE 10 �Χ�s�����A�� Version �ݩʫh�A�Ω� IE 9 �θ��ª����C
    $ieVersion = $null
    $ieVersion = (Get-ItemProperty -Path "HKLM:\Software\Microsoft\Internet Explorer" | Select-Object -Property "svcVersion").svcVersion
    if ($ieVersion -eq $null) {
        $ieVersion = (Get-ItemPropertyValue -Path "HKLM:\Software\Microsoft\Internet Explorer" -Name "Version")
    }

    Write-Output "IE version: $ieVersion"
    
    # �ˬd�����æw�� IE 11�]�p�G�����p�� 11�^
    if ([int16]$ieVersion.Split(".")[0] -lt 11) {
        
        Write-Output "Start to insall: $software_name"

        #�_���ɮר쥻���Ȧs"
        $software_path = get-item -Path $software_path
        if (!(Test-Path "$env:temp\$($software_path.Name)")) {
            Copy-Item -Path $software_path -Destination $env:temp -Recurse -Force -Verbose
        }

        ## �P�_OS�O32(x86)�άO64(AMD64), ��L��(ARM64)���w��  
        switch ($env:PROCESSOR_ARCHITECTURE) {
            "AMD64" { $software_exec = $software_msi_x64 }
            "x86" { $software_exec = $software_msi_x32 }
            default { Write-Warning "Unsupport CPU or OS:"  $env:PROCESSOR_ARCHITECTURE; $software_exec = $null }
        }

        if ($null -ne $software_exec) {
            Start-Process -FilePath "$env:temp\$($software_path.Name)\$software_exec" -ArgumentList "/passive /norestart" -Wait
            Start-Sleep -Seconds 5 

            Write-Output "IE 11 �w�˵���."
            write-output "�w��IE 11��hotfix"

            # �w�� MSU�]Microsoft Update Standalone Package�^�w�˥]�A
            # �z�i�H�ϥ� Start-Process �R�O�Ӱ��� wusa.exe�]Windows Update Standalone Installer�^�u��A�ë��w�n�w�˪� MSU �ɮסC
            
            $hotfix = get-childitem -Path "$env:temp\$($software_path.Name)\$($software_exec.Split("\")[1])\*" -Include "*.msu"

            foreach ($h in $hotfix) {
                Write-Output "Installing hotfix: $($h.Name) "
                Start-Process -FilePath wusa.exe -ArgumentList "$($h.fullname) /quiet /norestart" -Wait
                Start-Sleep -Seconds 3 
            }

        }
        else {
            Write-Warning "$software_name �L�k���`�w��."
        }
        
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
        install-IE11
    }
    else {
        Write-Warning "�L�k���o�޲z���v���Ӧw�˳n��, �ХH�޲z���b������."
    }
    pause
}