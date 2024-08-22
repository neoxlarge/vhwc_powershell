function get-installedprogramlist {
    # ���o�Ҧ��w�˪��n��,���U�w�˳n��|�Ψ�.

    ### Win32_product���M��ä�����A Winnexus �ä��b�̭�.
    ### $all_installed_program = Get-WmiObject -Class Win32_Product

    ### �Ҧ����n��|�b���U�o�T�ӵn���ɸ��|��

    $software_reg_path = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    return (Get-ItemProperty -Path $software_reg_path -ErrorAction SilentlyContinue)
}



function  get-admin_cred {
    #���o�޲z���v��
    $Username = "vhcy\vhwcmis"
    $Password = "Mis20190610"
    $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)
    
    return $credential
}



function Uninstall-Software {
    <#
    .SYNOPSIS
        �������w�W�r���n��
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name 
    )

    $credential = Get-Admin_Cred

    $allInstalledPrograms = Get-InstalledProgramList
    $softwareToUninstall = $allInstalledPrograms | Where-Object { $_.DisplayName -like $Name }

    if ($null -eq $softwareToUninstall) {
        return "�䤣��n��: $Name"
    }

    foreach ($software in $softwareToUninstall) {
        if ($software.UninstallString -like "msiexec*") {
            # MSI ����
            $uninstallString = $software.UninstallString.Split(" ")[1].replace("I", "X")
            $process = Start-Process -FilePath "msiexec.exe" -ArgumentList "$uninstallString /passive" -Credential $credential -PassThru
        }
        elseif ($software.QuietUninstallString) {
            # �ϥΦw�R�����r�Ŧ�
            $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c $($software.QuietUninstallString)" -Credential $credential -PassThru
        }
        else {
            # �ϥΤ@������r�Ŧ�
            $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c $($software.UninstallString)" -Credential $credential -PassThru
        }

        $process.WaitForExit()
        
        # �ˬd�������G
        if ($process.ExitCode -eq 0) {
            Write-Output "���\���� $($software.DisplayName)"
        }
        else {
            Write-Output "���� $($software.DisplayName) ���ѡA�h�X�X: $($process.ExitCode)"
        }
    }
}


function get-msiversion {
    # �qmsi�ɮפ������n�骺����.
    # from https://joelitechlife.ca/2021/04/01/getting-version-information-from-windows-msi-installer/comment-page-1/#respond
    param (
        [parameter(Mandatory = $true)] 
        [ValidateNotNullOrEmpty()] 
        [System.IO.FileInfo] $MSIPATH
    ) 
    if (!(Test-Path $MSIPATH.FullName)) { 
        throw "File '{0}' does not exist" -f $MSIPATH.FullName 
    } 
    try { 
        $WindowsInstaller = New-Object -com WindowsInstaller.Installer 
        $Database = $WindowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $Null, $WindowsInstaller, @($MSIPATH.FullName, 0)) 
        $Query = "SELECT Value FROM Property WHERE Property = 'ProductVersion'"
        $View = $database.GetType().InvokeMember("OpenView", "InvokeMethod", $Null, $Database, ($Query)) 
        $View.GetType().InvokeMember("Execute", "InvokeMethod", $Null, $View, $Null) | Out-Null
        $Record = $View.GetType().InvokeMember( "Fetch", "InvokeMethod", $Null, $View, $Null ) 
        $Version = $Record.GetType().InvokeMember( "StringData", "GetProperty", $Null, $Record, 1 ) 
        return $Version
    }
    catch { 
        throw "Failed to get MSI file version: {0}." -f $_
    }

}




function Compare-Version {
    <#
    .SYNOPSIS
        ���2�Ӫ���, $version1 �j�� $version2 �^��$Ture , ����Τp��^��$False
    .DESCRIPTION
        ��ƪ��ԲӴy�z
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$Version1, # �Ĥ@�Ӫ���
    
        [Parameter(Mandatory = $true)]
        [string]$Version2     # �ĤG�Ӫ���
    )
    
    # �N������������}�C�A�H�K�v�Ӥ���U�ӳ���
    $version1Array = $Version1.Split('.')
    $version2Array = $Version2.Split('.')
    
    # �ϥ� foreach �j��M���C�ӳ����i����
    foreach ($i in 0..$version1Array.Count) {
        if ([int]$version1Array[$i] -gt [int]$version2Array[$i]) {
            return $true    # ��^ $true ��ܲĤ@�Ӫ������j��ĤG�Ӫ�����
        }
        elseif ([int]$version1Array[$i] -lt [int]$version2Array[$i]) {
            return $false   # ��^ $false ��ܲĤ@�Ӫ������p��ĤG�Ӫ�����
        }
        else {
            # �p�G��e�����۵��A�h�~�����U�@�ӳ���
            continue
        }
    }
    
    # �p�G�����ۦP�A�h��ܪ������ۦP
    return $false    # ��^ $true ��ܨ�Ӫ������ۦP
}
    
      
function Get-OSVersion {
    #���oOS������
    $os = (Get-WmiObject -Class Win32_OperatingSystem).Caption

    if ($os -like "*Windows 7*") {
        return "Windows 7"
    }
    elseif ($os -like "*Windows 10*") {
        return "Windows 10"
    }
    elseif ($os -like "*Windows 11*") {
        return "Windows 11"
    }
    else {
        return "Unknown OS"
    }
}         


function Get-IPv4Address {
    <#
    �^�ǧ�쪺IP,�u��b172.*�~���. 
    #>

    $ip = Get-WmiObject -Class Win32_NetworkAdapterConfiguration |
    Where-Object { $_.IPAddress -ne $null -and $_.IPAddress[0] -like "172.20.*" } |
    Select-Object -ExpandProperty IPAddress |
    Where-Object { $_ -match "\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}" } |
    Select-Object -First 1

    if ($ip -eq $null) {
        return $null
    }
    else {     
        return $ip
    }
}

function Write-Log {
    param(
        [string]$Message,
        [string]$LogFile = "C:\Logs\MyLog.txt"
    )
    $Log_Title = "$(Get-Date), $($env:COMPUTERNAME), $(Get-IPv4Address), $(Get-OSVersion)_$($env:PROCESSOR_ARCHITECTURE)"
    "$Log_Title - $Message" | Out-File -FilePath $LogFile -Append
}