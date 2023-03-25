# �w�˨��r Trend Micro Apex One Security Agent

param($runadmin)

Import-Module ((Split-Path $PSCommandPath) + "\get-installedprogramlist.psm1")


function install-OfficeScan {
    # �w�˨��r Trend Micro Apex One Security Agent
    ## ��X�n��O�_�v�w��

    $all_installed_program = get-installedprogramlist
  
    $software_name = "Trend Micro Apex One Security Agent"
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }

    if ($software_is_installed -eq $null) {
    
        Write-Output "Start to insall: $software_name"

        #�_���ɮר�"D:\mis"
        $software_path = get-item -Path "\\172.20.1.14\share\software\officescan_antivir"
        if (Test-Path -Path "d:\mis") {
            $software_copyto_path = "D:\mis"
        }
        else {
            $software_copyto_path = "C:\mis"
        }
        
    
        Copy-Item -Path $software_path -Destination $software_copyto_path -Recurse -Force

        ## �P�_OS�O32(x86)�άO64(AMD64), ��L��(ARM64)���w��  
        switch ($env:PROCESSOR_ARCHITECTURE) {
            "AMD64" { $software_exec = "agent_cloud_x64.msi" }
            "x86" { $software_exec = "agent_cloud_x86.msi" }
            default { Write-OutPut "Unsupport CPU or OS:"  $env:PROCESSOR_ARCHITECTURE; $software_exec = $null }
        }

        if ($software_exec -ne $null) {
            Start-Process -FilePath "msiexec.exe" -ArgumentList "/i $($software_copyto_path + "\" + $software_path.Name + "\" + $software_exec) /passive /log d:\mis\install_officescan.log" -Wait
            Start-Sleep -Seconds 5 
        }
        else {
            $warn_msg += "Software install fail: $software_name"
            Write-Warning $warn_msg[-1] 
        }
      
        #�w�˧�, �R���w���ɮ�
        Remove-Item -Path ($software_copyto_path + "\" + $software_path.Name) -Recurse -Force
      
        #�w�˧�, �A���s���o�w�˸�T
        $all_installed_program = get-installedprogramlist
        $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }
    
    }

    Write-Output ("Software has installed: " + $software_is_installed.DisplayName)
    Write-Output ("Version: " + $software_is_installed.DisplayVersion)

}



#�ɮ׿W�߰���ɷ|����禡, �p�G�O�Q?�J�ɤ��|����禡.
if ($run_main -eq $null) {

    #�ˬd�O�_�޲z��
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #�p�G�D�޲z��, �N�յ�run as admin, �öǤJrunadmin �Ѽ�1. �]���b����@��ϥΪ̥û������O�޲z���v��, �|�y���L�����]. ���ѼƥΨӻ��U�P�_�u�]�@��. 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    
    }

    if ($check_admin) { 
        install-officescan
    }
    else {
        Write-Warning "�L�k���o�޲z���v���Ӧw�˳n��, �ХH�޲z���b������."
    }
    pause
}