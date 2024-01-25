#enable-ChangJieinput 倉頡輸入法

param($runadmin)


#取得OS的版本
function Get-OSVersion {
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


function import-module_func ($name) {
#此function會檢查本機上是否有要載入的模組. 如果沒有, 就連線到wcdc2.vhcy.gov上下載. 可能Win7沒有內建該模組. 
    $result = get-module -ListAvailable $name

    if ($result -ne $null) {

        Import-Module -Name $name -ErrorAction Stop

    } else {

        $Username = "vhwcmis"
        $Password = "Mis20190610"
        $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
        $credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)

        $rsession = New-PSSession -ComputerName wcdc2.vhcy.gov.tw -Credential $credential
        Import-Module $name -PSSession $rsession -ErrorAction Stop
        Disconnect-PSSession -Session $rsession | Out-Null
    }
    
}

function Enable-ChangJieinput {

    import-module_func International

    $OS = Get-OSVersion

    if ($os -in @("Windows 10","Windows 11")) {

        #安裝倉頡輪入法 windows 10
        $user_language = Get-WinUserLanguageList
        $GuidChangJie = "0404:{531FDEBF-9B4C-4A43-A2AA-960E8FCDC732}{4BDF9F03-C7D3-11D4-B2AB-0080C882687E}"
        #$GuidNewChangJie = "0404:{B115690A-EA02-48D5-A231-E3578D2FDF80}{F3BA907A-6C7E-11D4-97FA-0080C882687E}"


        foreach ($i in $user_language) { 
            if ($i.LanguageTag -eq "zh-Hant-TW") {
        
                if ($i.InputMethodTips -contains $GuidChangJie) {
                    write-output "倉頡輸入法($($OS))己安裝."
                    $i = $null
                    break
                }else {
                    Write-Output "倉頡輸入法($($OS))未安裝."
                    break 
                }
            }
        }


        if ($i -ne $null) {
        
            Write-Output "啟用倉頡輸入法($($OS))."
            $i.InputMethodTips.add($GuidChangJie)
        
            if ((Test-Path -Path "HKCU:\SOFTWARE\Microsoft\CTF\TIP\{531FDEBF-9B4C-4A43-A2AA-960E8FCDC732}") -ne $null) {
            
                New-Item -Path "HKCU:\SOFTWARE\Microsoft\CTF\TIP\{531FDEBF-9B4C-4A43-A2AA-960E8FCDC732}\LanguageProfile\0x00000404\{4BDF9F03-C7D3-11D4-B2AB-0080C882687E}" -Force -ErrorAction SilentlyContinue
                New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\CTF\TIP\{531FDEBF-9B4C-4A43-A2AA-960E8FCDC732}\LanguageProfile\0x00000404\{4BDF9F03-C7D3-11D4-B2AB-0080C882687E}" -Name "Enable" -Value 1 -PropertyType DWord 
    
                New-Item -Path "HKCU:\SOFTWARE\Microsoft\CTF\SortOrder\AssemblyItem\0x00000404\{34745C63-B2F0-4784-8B67-5E12C8701A31}\00000001" -Force
                New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\CTF\SortOrder\AssemblyItem\0x00000404\{34745C63-B2F0-4784-8B67-5E12C8701A31}\00000001" -Name "Profile" -Value "{4BDF9F03-C7D3-11D4-B2AB-0080C882687E}" -PropertyType string -Force
                New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\CTF\SortOrder\AssemblyItem\0x00000404\{34745C63-B2F0-4784-8B67-5E12C8701A31}\00000001" -Name "KeyboardLayout" -Value 0 -PropertyType DWORD -Force 
                New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\CTF\SortOrder\AssemblyItem\0x00000404\{34745C63-B2F0-4784-8B67-5E12C8701A31}\00000001" -Name "CLSID" -Value "{531FDEBF-9B4C-4A43-A2AA-960E8FCDC732}" -PropertyType string -Force
                New-ItemProperty -Path "HKCU:\Control Panel\International\User Profile\zh-Hant-TW" -Name "0404:{531FDEBF-9B4C-4A43-A2AA-960E8FCDC732}{4BDF9F03-C7D3-11D4-B2AB-0080C882687E}" -Value 2 -PropertyType string -Force
            }
    
            Set-WinUserLanguageList -LanguageList $user_language -Force
            #設定預設輸入模式為"英數字元"
            Set-ItemProperty -Path HKCU:\SOFTWARE\Microsoft\IME\15.0\IMETC -Name "Default Input Mode" -Value "0x00000001" -Type String -Force

            if (Test-Path -Path HKCU:\SOFTWARE\Microsoft\InputMethod\Settings\CHT) {
            Set-ItemProperty -Path HKCU:\SOFTWARE\Microsoft\InputMethod\Settings\CHT -name "Default Input Mode Changjie" -Value "0x00000000" -Type String -Force
            } else {
            New-Item -Path HKCU:\SOFTWARE\Microsoft\InputMethod\Settings\CHT -ItemType registry
            New-ItemProperty -Path HKCU:\SOFTWARE\Microsoft\InputMethod\Settings\CHT -Name "Default Input Mode Changjie" -Value "0x00000000" -PropertyType String -Force
            }
        }


    } elseif ($os -eq "Windows 7") {
        #安裝倉頡輸入法　windows 7

        $langueage_path = "HKCU:\Software\Microsoft\CTF\TIP\{531FDEBF-9B4C-4A43-A2AA-960E8FCDC732}\LanguageProfile\0x00000404\{4BDF9F03-C7D3-11D4-B2AB-0080C882687E}"

        if ((test-path -Path $langueage_path) -eq $false) {

            Write-Output "倉頡輸入法($($OS))未安裝."
            Write-Output "啟用倉頡輸入法$($OS))."
            New-Item -Path $langueage_path -Force
            New-ItemProperty -Path $langueage_path -Name "Enable" -Value 1 -PropertyType DWORD
    
        } else {
    
        Write-Output "倉頡輸入法($($OS))己安裝."
    
        }

    } else {

        Write-Output "非win7或win10,Win11,暫不安裝倉頡輸入法"

    }

}

#檔案獨立執行時會執行函式, 如果是被滙入時不會執行函式.
if ($run_main -eq $null) {

    #檢查是否管理員
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
    #如果非管理員, 就試著run as admin, 並傳入runadmin 參數1. 因為在網域一般使用者永遠拿不是管理員權限, 會造成無限重跑. 此參數用來輔助判斷只跑一次. 
    Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    
    }

    Enable-ChangJieinput
    
    pause
}