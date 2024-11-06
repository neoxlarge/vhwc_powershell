# �w��CMS_CGServiSignAdapter
# 20240611 �]��CGServiSignAdapter����s��1.0.23
# �쥻��1.0.22�|�n�D�n��s�~��n�J���OVPN.

param($runadmin)

# �n�Dpowershell v5.1�H�W�~����, win7�w�]powershell v2.0.
if (!$PSVersionTable.PSCompatibleVersions -match "^5\.1") {
    Write-Output "powershell requires version 5.1, exit"
    Start-Sleep -Seconds 3
    exit
}

function import-vhwcmis_module {
    $moudle_paths = @(
        if ($script:MyInvocation.MyCommand.Path) {"$(Split-Path $script:MyInvocation.MyCommand.Path -ErrorAction SilentlyContinue)"},
        "d:\mis\vhwcmis",
        "c:\mis\vhwcmis",
        "\\172.20.5.185\powershell\vhwc_powershell",
        "\\172.20.1.14\share\00�s�q���w��\vhwc_powershell",
        "\\172.20.1.122\share\software\00newpc\vhwc_powershell",
        "\\172.19.1.229\cch-share\h040_�i�q��\vhwc_powershell"
    )

    $filename = "vhwcmis_module.psm1"

    foreach ($path in $moudle_paths) {
        
        if (Test-Path "$path\$filename") {
            write-output "$path\$filename"
            Import-Module "$path\$filename" -ErrorVariable $err_import_module
            if ($err_import_module -eq $null) {
                Write-Output "Imported module path successed: $path\$filename"
                break
            }
        }
    }

    $result = get-command -Name "get-installedprogramlist" -CommandType Function -ErrorAction SilentlyContinue
    if ($result -eq $null) {
        throw "�L�k���Jvhwcmis_module�Ҳ�, �{���L�k���`����. "
    }
}
import-vhwcmis_module


function install-CMS {
    $log_file = "\\172.20.1.14\update\0001-���a�t�����ҳ]�w\VHWC_logs\update-cms.log"

    ## �w��CMS_CGServiSignAdapter
    ### �̤��n�D,�w�˫e���������r�n��, �ҥH�񨾬r���w��

    $software_name = "NHIServiSignAdapterSetup"
    $software_path = "\\172.20.5.187\mis\05-CMS_CGServiSignAdapterSetup\CMS_CGServiSignAdapterSetup"
    $software_exec = "NHIServiSignAdapterSetup.exe"
    
    $all_installed_program = get-installedprogramlist
   
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like "$software_name*" }


    if ($software_is_installed) {
        $exe_version = (Get-ItemProperty -Path "$software_path\$software_exec").VersionInfo.FileversionRaw.toString()

        $result = Compare-Version -Version1 $exe_version -Version2 $software_is_installed.DisplayVersion

        if ($result) {
 
            $ipv4 = Get-IPv4Address 
            Write-Log -logfile $log_file -message "Find old CMS version:$($software_is_installed.DisplayVersion)"
  
            Write-Output "Find old CMS $software_name, version: $($software_is_installed.DisplayVersion)"
            Write-Output "Removing old version."
            Start-Process -FilePath $software_is_installed.UninstallString -ArgumentList "/S" -Wait
            $software_is_installed = $null
        }
    }

    if ($software_is_installed -eq $null ) {
        # �S�w��, �����w��.
        Write-Output "Start to install $software_name"

        #�ӷ����| ,�n�_����|,and �w�˰���{���W��
        $software_path = get-item -Path $software_path
                
        #�_���ɮר�temp
        Copy-Item -Path $software_path -Destination $env:temp -Recurse -Force -Verbose

        #installing...
        $process_id = Start-Process -FilePath "$env:temp\$($software_path.Name)\$software_exec" -PassThru

        #�̦w�ˤ��, CGServiSignMonitor�|�̫�Q�}��, �ҥH�ˬd��ӵ{�ǰ����, ��ܦw�˧���.
        $process_exist = $null
        while ($process_exist -eq $null) {
            $process_exist = Get-Process -Name CGServiSignMonitor -ErrorAction SilentlyContinue
            if ($process_exist -ne $null) { Stop-Process -Name $process_id.Name }
            write-output ($process_id.Name + "is installing...wait 5 seconds.")
            Start-Sleep -Seconds 5
        } 

        #�w�˧�, �A���s���o�w�˸�T
        $all_installed_program = get-installedprogramlist
        $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like "$software_name*" }


    }

    Write-output ("software has installed:" + $software_is_installed.DisplayName )
    Write-Output ("Version:" + $software_is_installed.DisplayVersion)


    

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
    install-CMS
    } else {
        Write-Warning "�L�k���o�޲z���v���Ӧw�˳n��, �ХH�޲z���b������." 
    }
    pause
    #Start-Sleep -Seconds 5
}