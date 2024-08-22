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



function  get-admin_cred {
    #取得管理員權限
    $Username = "vhcy\vhwcmis"
    $Password = "Mis20190610"
    $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)
    
    return $credential
}



function Uninstall-Software {
    <#
    .SYNOPSIS
        移除指定名字的軟體
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name 
    )

    $credential = Get-Admin_Cred

    $allInstalledPrograms = Get-InstalledProgramList
    $softwareToUninstall = $allInstalledPrograms | Where-Object { $_.DisplayName -like $Name }

    if ($null -eq $softwareToUninstall) {
        return "找不到軟體: $Name"
    }

    foreach ($software in $softwareToUninstall) {
        if ($software.UninstallString -like "msiexec*") {
            # MSI 卸載
            $uninstallString = $software.UninstallString.Split(" ")[1].replace("I", "X")
            $process = Start-Process -FilePath "msiexec.exe" -ArgumentList "$uninstallString /passive" -Credential $credential -PassThru
        }
        elseif ($software.QuietUninstallString) {
            # 使用安靜卸載字符串
            $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c $($software.QuietUninstallString)" -Credential $credential -PassThru
        }
        else {
            # 使用一般卸載字符串
            $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c $($software.UninstallString)" -Credential $credential -PassThru
        }

        $process.WaitForExit()
        
        # 檢查卸載結果
        if ($process.ExitCode -eq 0) {
            Write-Output "成功移除 $($software.DisplayName)"
        }
        else {
            Write-Output "移除 $($software.DisplayName) 失敗，退出碼: $($process.ExitCode)"
        }
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


function Get-IPv4Address {
    <#
    回傳找到的IP,只能在172.*才能用. 
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