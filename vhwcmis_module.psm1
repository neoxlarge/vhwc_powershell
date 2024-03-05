function get-installedprogramlist {
    # 取得所有安裝的軟體,底下安裝軟體會用到.

    ### Win32_product的清單並不完整， Winnexus 並不在裡面.
    ### $all_installed_program = Get-WmiObject -Class Win32_Product

    ### 所有的軟體會在底下這三個登錄檔路徑中

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
        移除指定名字的軟體
    
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
        return "找不到軟體: $name"
    }

}


function get-msiversion {
    # 從msi檔案中提取軟體的版本.
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
    #取得管理員權限
    $Username = "vhcy\vhwcmis"
    $Password = "Mis20190610"
    $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)
    
    return $credential
}




function Compare-Version {
    <#
    .SYNOPSIS
        比對2個版本, $version1 大於 $version2 回傳$Ture , 等於或小於回傳$False
    .DESCRIPTION
        函數的詳細描述
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$Version1, # 第一個版本
    
        [Parameter(Mandatory = $true)]
        [string]$Version2     # 第二個版本
    )
    
    # 將版本號拆分成陣列，以便逐個比較各個部分
    $version1Array = $Version1.Split('.')
    $version2Array = $Version2.Split('.')
    
    # 使用 foreach 迴圈遍歷每個部分進行比較
    foreach ($i in 0..$version1Array.Count) {
        if ([int]$version1Array[$i] -gt [int]$version2Array[$i]) {
            return $true    # 返回 $true 表示第一個版本號大於第二個版本號
        }
        elseif ([int]$version1Array[$i] -lt [int]$version2Array[$i]) {
            return $false   # 返回 $false 表示第一個版本號小於第二個版本號
        }
        else {
            # 如果當前部分相等，則繼續比較下一個部分
            continue
        }
    }
    
    # 如果完全相同，則表示版本號相同
    return $false    # 返回 $true 表示兩個版本號相同
}
    
      
function Get-OSVersion {
    #取得OS的版本
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