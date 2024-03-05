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



function uninstall-software {
    <#
    .SYNOPSIS
        �������w�W�r���n��
    
    #>

    [Parameter(Mandatory = $true)]
    [string]$name 


    $mymodule_path = Split-Path $PSCommandPath + "\"
    Import-Module $mymodule_path + "get-installedprogramlist.psm1"
    Import-Module $mymodule_path + "get-admin_cred.psm1"

    $credential = get-admin_cred

    $all_installed_program = get-installedprogramlist
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $name }


    if ($software_is_installed -ne $null) {
        $uninstallstring = $software_is_installed.uninstallString.Split(" ")[1].replace("I", "X")

        $running_proc = Start-Process -FilePath "msiexec.exe" -ArgumentList "$uninstallstring /passive" -Credential $credential -PassThru
        $running_proc.WaitForExit()     
    }
    else {
        return "�䤣��n��: $name"
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



function  get-admin_cred {
    #���o�޲z���v��
    $Username = "vhcy\vhwcmis"
    $Password = "Mis20190610"
    $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)
    
    return $credential
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