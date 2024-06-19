#�w��CMS_CGServiSignAdapter

param($runadmin)

Import-Module ((Split-Path $PSCommandPath) + "\get-installedprogramlist.psm1")

function install-CMS {
    ## �w��CMS_CGServiSignAdapter
    ### �̤��n�D,�w�˫e���������r�n��, �ҥH�񨾬r���w��

    $software_name = "NHIServiSignAdapterSetup"
    $software_path = "\\172.20.5.187\mis\05-CMS_CGServiSignAdapterSetup\CMS_CGServiSignAdapterSetup"
    $software_exec = "NHIServiSignAdapterSetup_1.0.22.0830.exe"
    
    $all_installed_program = get-installedprogramlist
   
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like "$software_name*" }

    if ($null -eq $software_is_installed) {
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
    }
    else {
        Write-Warning "�L�k���o�޲z���v���Ӧw�˳n��, �ХH�޲z���b������."
    }
    pause
}