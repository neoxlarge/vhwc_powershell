# �w��SmartIris

param($runadmin)

Import-Module ((Split-Path $PSCommandPath) + "\get-installedprogramlist.psm1")


function Parse-IniFile {
    <#
    code from chatgpt
    �o�Ө�ƨϥΤF PowerShell ���i���ѼƳB�z����ӳB�z���|�ѼơA�æbŪ�� .ini �ɮת��C�@��ɶi��A���ѪR�C����٨ϥΤF���Ѩӻ����C�ӳ������\��A���{���X�����z�ѩM���@�C
    #>
    
        # �ϥ� [CmdletBinding()] �ݩʨӶ}�Ҷi�����ѼƳB�z
        # �o���\�b��Ƥ��ϥζi���ѼơA�Ҧp Mandatory�BParameterSetName ��
        [CmdletBinding()]
        param (
            # �ϥ� [Parameter()] �ݩʨӫ��w���n�����|�Ѽ�
            [Parameter(Mandatory = $true)]
            [string]$Path
        )
     
        # �T�O�ɮצs�b
        if (-not (Test-Path $Path)) {
            throw "The file '$Path' does not exist."
        }
     
        # ��l�Ƥ@�ӪŪ������A�Ω�s�x�ѪR�᪺ .ini ���e
        $ini = @{}
     
        # ��l�Ƥ@�ӪŪ��`�I�W���ܼơA�Ω�ѪR�ثe���`�I
        $section = ""
     
        # Ū�� .ini �ɮפ����C�@��
        Get-Content $Path | ForEach-Object {
            # �h���C��e�᪺�Ů�
            $line = $_.Trim()
     
            # �p�G�Ӧ�O�`�I�W�١A�h�ѪR�X�`�I�W�٨ê�l�Ƥ@�ӷs�������
            if ($line -match "^\[.*\]$") {
                $section = $line.Substring(1, $line.Length - 2)
                $ini[$section] = @{}
            }
            # �p�G�Ӧ�O��ȹ�A�h�ѪR�X��M�ȡA�ñN��s�J�ثe�`�I�������
            elseif ($line -match "^([^=]+)=(.*)$") {
                $key = $matches[1].Trim()
                $value = $matches[2].Trim()
                $ini[$section][$key] = $value
            }
        }
     
        # ��^�ѪR�᪺ .ini ���e�����
        return $ini
    
    }
    

    function  save-iniFile {

        param (
            [CmdletBinding()]
            $ini,
            [CmdletBinding()]
            $path
        )
    
        $ini_content = ""
    
        foreach ($i in $ini.keys) {
            $ini_content += "[$i] `n"
            
            foreach ($j in $ini.$i.keys) {
                $ini_content += "$j=$($ini.$i.$j.tostring()) `n"
            }
        }
    
        Write-Output $ini_content
        Out-File -InputObject $ini_content -FilePath $path
    }
    
    
    

function check-SmartIris {

    $ini_path = "C:\temp\TEDPC\SmartIris\UltraQuery\SysIni\LocalSetting.ini"

    if (Test-Path -Path $ini_path) {
        $ini = Parse-IniFile -Path $ini_path

        $ini.LocalSetting.AETitle = $env:COMPUTERNAME
        
        save-iniFile -ini $ini -path $ini_path
    }

}

function install-SmartIris {

    
    $software_name = "SmartIris"
    $software_path = "\\172.20.5.187\mis\02-SmartIris\SmartIris_V1.3.6.4_Beta7_UQ-1.1.0.19_R2_Install_20200701"
    $software_exe = "setup.exe"
    ## ��X�n��O�_�v�w��

    $all_installed_program = get-installedprogramlist
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }

    if ($null -eq $software_is_installed) {
    
        Write-Output "Start to insall: $software_name"

        #�_���ɮר쥻���Ȧs"
        $software_path = get-item -Path $software_path
        Copy-Item -Path $software_path -Destination $env:temp -Recurse -Force -Verbose

        # �w��  
        $runid = Start-Process -FilePath $($env:temp + "\" + $software_path.Name + "\" + $software_exe) -ArgumentList "/s /f1$($env:temp + "\" + $software_path.Name + "\vhwc.iss")" -PassThru
        
        # �w�˹L�{��, �o2��{�|���X��, ����F, ���|�A�]�w�Y�i.
        while (!($runid.HasExited)) {
            get-process -Name MonitorCfg -ErrorAction SilentlyContinue | Stop-Process
            get-process -Name UQ_Setting -ErrorAction SilentlyContinue | Stop-Process
            Start-Sleep -Seconds 1
        }

        #�_��]�w�ɨ쥻��.
        copy-item -Path "\\172.20.5.187\mis\02-SmartIris\vhwc_UltraQuery_SysIni\*" -Destination "C:\TEDPC\SmartIris\UltraQuery\SysIni" -Force
 

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
        install-SmartIris    
    }
    else {
        Write-Warning "�L�k���o�޲z���v���Ӧw�˳n��, �ХH�޲z���b������."
    }
    pause
}