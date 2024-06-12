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


function install-CMS {

    # ���ovhwcmis_module.psm1��3�ؤ覡:
    # 1.�{�������e���|, ���Group police����i��줣��.
    # 2.�`�Ϊ����|, d:\mis\vhwc_powershell, ���O�C�x������.
    # 3.�s��NAS�W���o. �D���쪺�q���|�S��NAS���v��, ����ʳs�WNAS.

    $pspaths = @()
    $pspaths += "$(Split-Path $PSCommandPath)\vhwcmis_module.psm1"
    $pspaths += "d:\mis\vhwc_powershell\vhwcmis_module.psm1"

    $path = "\\172.20.1.122\share\software\00newpc\vhwc_powershell"
    if (!(test-path $path)) {
        $Username = "software_download"
        $Password = "Us2791072"
        $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
        $credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)
        $nas_name = "nas122"

        New-PSDrive -Name $nas_name -Root "$path" -PSProvider FileSystem -Credential $credential | Out-Null
    }
    $pspaths += "$path\vhwcmis_module.psm1"

    foreach ($path in $pspaths) {
        Import-Module $path -ErrorAction SilentlyContinue
        if ((get-command -Name "get-installedprogramlist" -CommandType Function -ErrorAction SilentlyContinue)) {
            break
        }
    }


    ## �w��CMS_CGServiSignAdapter
    ### �̤��n�D,�w�˫e���������r�n��, �ҥH�񨾬r���w��

    $software_name = "NHIServiSignAdapterSetup"
    $software_path = "\\172.20.1.122\share\software\00newpc\05-CMS_CGServiSignAdapterSetup\CMS_CGServiSignAdapterSetup"
    $software_exec = "NHIServiSignAdapterSetup.exe"
    
    $all_installed_program = get-installedprogramlist
   
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like "$software_name*" }


    if ($software_is_installed) {
        $exe_version = (Get-ItemProperty -Path "$software_path\$software_exec").VersionInfo.FileversionRaw.toString()

        $result = Compare-Version -Version1 $exe_version -Version2 $software_is_installed.DisplayVersion

        if ($result) {
            Write-Output "Find old version $software_name : $($all_installed_program.DisplayVersion)"
            Write-Output "Removing old version."
            Start-Process -FilePath $software_is_installed.UninstallString -ArgumentList "/S" -Wait
            $software_is_installed = $null
        }
    }

    if ($software_is_installed -eq $null) {
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

    if (!$(Get-PSDrive -Name $nas_name -ErrorAction SilentlyContinue) -eq $null) {
        Remove-PSDrive -Name $nas_name
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
        install-CMS
    }
    else {
        Write-Warning "�L�k���o�޲z���v���Ӧw�˳n��, �ХH�޲z���b������."
    }
    #pause
    Start-Sleep -Seconds 5
}