# �w�˴���Ū�d��,�p�վ� ezusb Driver
# /S /v /qn
# ��Jsetup.exe /?, �i�H�ݨ�䴩silent�w�˪����R, �o�w�ˬOmsi�ɬ���¦�b�]��.exe
# ���嬰0404
param($runadmin)

Import-Module ((Split-Path $PSCommandPath) + "\get-installedprogramlist.psm1")

function install-EZUSB {
    $software_name = "EZUSB Driver"
    $software_path = "\\172.20.5.187\mis\07-Ū�d��\Pisces-310PU"
    $software_msi = "\EZ100_Driver_64bit\setup.exe"
    $software_msi_x86 = "\EZ100_Driver_32bit\setup.exe"
    ## ��X�n��O�_�v�w��

    $all_installed_program = get-installedprogramlist
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }

    if ($null -eq $software_is_installed) {
    
        Write-Output "Start to insall: $software_name"

        #�_���ɮר쥻���Ȧs"
        $software_path = get-item -Path $software_path
        Copy-Item -Path $software_path -Destination $env:temp -Recurse -Force

        ## �P�_OS�O32(x86)�άO64(AMD64), ��L��(ARM64)���w��  
        switch ($env:PROCESSOR_ARCHITECTURE) {
            "AMD64" { $software_exec = $software_msi }
            "x86" { $software_exec = $software_msi_x86 }
            default { Write-Warning "Unsupport CPU or OS:"  $env:PROCESSOR_ARCHITECTURE; $software_exec = $null }
        }

        if ($null -ne $software_exec) {
            Start-Process -FilePath $($env:temp + "\" + $software_path.Name + "\" + $software_exec) -ArgumentList "/S /v/passive" -Wait
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
        install-EZUSB    
    }
    else {
        Write-Warning "�L�k���o�޲z���v���Ӧw�˳n��, �ХH�޲z���b������."
    }
    pause
}