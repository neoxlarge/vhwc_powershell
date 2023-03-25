# �w�˶��ݦw�����󰷫O�dŪ�d������(PCSC)
## 1. ���w��VC++�i��o�M��

param($runadmin)

Import-Module ((Split-Path $PSCommandPath) + "\get-installedprogramlist.psm1")

function install-PCSC {
    # �w�˶��ݦw�����󰷫O�dŪ�d������(PCSC)
    ## 1. ���w��VC++�i��o�M��

    $all_installed_program = get-installedprogramlist

    $software_name = "���O�dŪ�d������(PCSC)*"
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }

    

    if ($software_is_installed -eq $null) {
    
        Write-Output ("Start to install: " + $software_name)

        $software_path = get-item -Path "\\172.20.1.14\update\Vghtc_Update\00_mis\CMS_CS5.1.5.5-Ū�d������n��"
        $software_copyto_path = "C:\VGHTC\00_mis"

        #�_���ɮר�"C:\VGHTC\00_mis"
        Copy-Item -Path $software_path -Destination $software_copyto_path -Recurse -Force 

        Write-OutPut ("Start to install software: " + $software_name)
    
        ## 1. ���w��VC++�i��o�M��
        $software_exec = "CMS_CS5.1.5.5\CS5.1.5.5��\gCIE_Setup\vcredist_x86\vcredist_x86.exe"#2
        Start-Process -FilePath ($software_copyto_path + "\" + $software_path.Name + "\" + $software_exec) -ArgumentList "/passive" -Wait

        Start-Sleep -Seconds 5

        ## 2.�w�˶��ݤ���
        $software_exec = "CMS_CS5.1.5.5\CS5.1.5.5��\gCIE_Setup\gCIE_Setup.msi"
        Start-Process -FilePath "msiexec.exe" -ArgumentList "/i $($software_copyto_path + "\" + $software_path.Name + "\" + $software_exec) /passive /log d:\mis\install_pcsc.log" -Wait
    
        ## 3.�n�]�]�w�� bat��, ���F���Q�쥻bat����pause�d��, ���s�@��powershell����.

        Write-Output "�]Ū�d������n�骺�]�wbat"
        
        #1.����������Ū�d������

        $setup_file_ = @(
            "C:\VGHTC\ICCard\CsHis.dll",
            "C:\ICCARD_HIS\CsHis.dll",
            "C:\vhgp\ICCard\CsHis.dll"
        )

        foreach ($i in $setup_file_) {
            Copy-Item -Path "C:\VGHTC\00_mis\CMS_CS5.1.5.5-Ū�d������n��\CSHIS-ic20-����Ū�d������-v5155.dll" -Destination $i -Force
            $i_version = Get-ItemProperty -Path $i
            Write-Output ("Check dll: " + $i_version.FullName + " Version: " + $i_version.VersionInfo.ProductVersion )
        }

        #2.copy�W��SAM��-�ܫ��w��m
        Copy-Item -Path "C:\VGHTC\00_mis\CMS_CS5.1.5.5-Ū�d������n��\0640140012001000005984.SAM" -Destination "C:\NHI\SAM\COMX1\0640140012001000005984.SAM" -Force
        $i_version = Get-ItemProperty -Path "C:\NHI\SAM\COMX1\0640140012001000005984.SAM"
        Write-Output ("Check dll: " + $i_version.FullName + " Exists: " + $i_version.Exists )
    
        #3. ���ݦw���Ҳ�-���all-user�Ұ�
        Copy-Item -Path "C:\VGHTC\00_mis\CMS_CS5.1.5.5-Ū�d������n��\4-���ݦw���ҲեD���x-v5155.lnk" -Destination "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\���ݦw���ҲեD���x.lnk" -Force
        $i_version = Get-ItemProperty -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\���ݦw���ҲեD���x.lnk"
        Write-Output ("Check dll: " + $i_version.FullName + " Exists: " + $i_version.Exists )

        #5. 5-copy-dll-to-c-v5155 , �N�O�_��"C:\VGHTC\00_mis\CMS_CS5.1.5.5-Ū�d������n��\copy-to-C\ICCARD_HIS"�̩Ҧ�dll��3�Ӹ��?.
        $setup_file_ = Get-ChildItem -Path "C:\VGHTC\00_mis\CMS_CS5.1.5.5-Ū�d������n��\copy-to-C\ICCARD_HIS"
    
        $setup_file_target_path = @(
            "C:\ICCARD_HIS",
            "C:\Windows\System32",
            "C:\Windows\System"    
        )

        foreach ($i in $setup_file_) {
            Write-Output ("dll name: " + $i.Name + "dll versoin: " + $i.VersionInfo.ProductVersion    )

            foreach ($j in $setup_file_target_path) {
                copy-item -Path $i.FullName -Destination ($j + "\" + $i.Name)
                $j_version = Get-ItemProperty -Path ($j + "\" + $i.Name)
                Write-Output ("Check dll: " + $j_version.FullName + " Version: " + $j_version.VersionInfo.ProductVersion )
            }
            Write-Output "`n" #����@�U
        }
        
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
        install-PCSC
    }
    else {
        Write-Warning "�L�k���o�޲z���v���Ӧw�˳n��, �ХH�޲z���b������."
    }
    pause
}