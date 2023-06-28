#install Powershell 5.1

param($runadmin)

function install-ps5 {
    if (($PSVersionTable.PSVersion.Major -lt 5) -and ($PSVersionTable.PSVersion.Minor -lt 1)) {
        Write-Output "Powershell �ثe������ $($PSVersionTable.PSVersion), �ɯŪ����5.1"

        $software_name = "Powershell"
        $software_path = "\\172.20.5.185\powershell\powershell5.1forWin7"
        $software_msi_x64 = "Win7-KB3191566-x64.zip"
        $software_msi_x32 = "Win7-KB3191566-x86.zip"
    
        #�ˬd�@�U�Ȧs�ؿ��O�_�s�b
        if (!(Test-Path -Path "$env:TEMP\$($software_path.Split("\")[-1])")) {
            switch ($env:PROCESSOR_ARCHITECTURE) {
                "AMD64" { $zip_path = "$software_path\$software_msi_x64"}
                "x86" { $zip_path = "$software_path\$software_msi_x32"}
            }

            New-Item -Path "$env:TEMP\$($software_path.Split("\")[-1])" -ItemType directory -Force
            
            Start-Process unzip.exe -ArgumentList "-o $zip_path -d $env:TEMP\$($software_path.Split("\")[-1])" -Wait -NoNewWindow

            Invoke-Expression "$env:TEMP\$($software_path.Split("\")[-1])\Install-WMF5.1.ps1 -AcceptEULA -AllowRestart" 
            
            write-output "Pause, please enter ..."
            
            $null = Read-Host
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
        install-ps5    
    }
    else {
        Write-Warning "�L�k���o�޲z���v���Ӧw�˳n��, �ХH�޲z���b������."
    }
    
}