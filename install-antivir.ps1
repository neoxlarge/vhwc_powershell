# �w�˨��r Trend Micro Apex One Security Agent

param($runadmin)

$mymodule_path = "$(Split-Path $PSCommandPath)\"
Import-Module -name "$($mymodule_path)vhwcmis_module.psm1"


function install-AntiVir {
    # �w�˨��r Trend Micro Apex One Security Agent
    ## ��X�n��O�_�v�w��

    $software_name = "Trend Micro Apex One Security Agent"
    $software_path = "\\172.20.1.122\share\software\00newpc\officescan_antivir"
    $software_msi_x64 = "agent_cloud_x64.msi"
    $software_msi_x32 = "agent_cloud_x86.msi"

    $Username = "vhcy\vhwcmis"
    $Password = "Mis20190610"
    $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)

    $all_installed_program = get-installedprogramlist
      
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }

    if ($software_is_installed -eq $null) {
    
        Write-Output "Start to insall: $software_name"

        #�_���ɮר�temp
        #$software_path = get-item -Path $software_path
                
        #copy-item �L�k���{��, ���n�qpsdrive��, �ҥH�n��driver.
        $net_driver = "vhwcdrive" #�u�O����driver�W�r�Ӥv.
        New-PSDrive -Name $net_driver -Root $software_path -PSProvider FileSystem -Credential $credential
        Copy-Item -Path "$($net_driver):\" -Destination "$($env:TEMP)\$software_name" -Recurse -Force -Verbose 
        Remove-PSDrive -Name $net_driver


        ## �P�_OS�O32(x86)�άO64(AMD64), ��L��(ARM64)���w��  
        switch ($env:PROCESSOR_ARCHITECTURE) {
            "AMD64" { $software_exec = $software_msi_x64 }
            "x86" { $software_exec = $software_msi_x32 }
            default { Write-Warning "$software_name �L�k���`�w��: ���䴩���t��:  $($env:PROCESSOR_ARCHITECTURE)"; $software_exec = $null }
        }

        if ($software_exec -ne $null) {
            $argumentlist = "/i ""$($env:temp + "\" + $software_Name + "\" + $software_exec)"" /passive /log install_officescan.log"
            if ($check_admin) {
                $proc = Start-Process -FilePath "msiexec.exe" -ArgumentList $argumentlist -PassThru
            } else {
                $proc = Start-Process -FilePath "msiexec.exe" -ArgumentList $argumentlist -Credential $credential -PassThru
            }
            $proc.WaitForExit()
            Start-Sleep -Seconds 1 
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
    
    install-AntiVir
    pause
}