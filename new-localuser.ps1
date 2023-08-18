<#
這個函式 CreateLocalUser 接受幾個參數：帳號名稱 ($Username)、密碼 ($Password)、全名 ($FullName) 和描述 ($Description)。

在函式中，我們使用 Get-WMIObject 來檢查帳號是否存在。如果帳號已經存在，函式會顯示訊息並立即返回。

如果帳號不存在，我們則使用 New-LocalUser 創建新帳號。如果創建成功，函式會顯示成功訊息；如果發生錯誤，則會顯示錯誤訊息。

您可以根據需要調用此函式，並傳遞相應的參數值，例如：CreateLocalUser -Username "vhwc" -Password "Vh2791641" -FullName "VHWC User" -Description "Test user"。


# 使用範例：
CreateLocalUser -Username "vhwc" -Password "Vh2791641" -FullName "VHWC User" -Description "Test user"
#>



param($runadmin)

function CreateLocalUser {
    param(
        [Parameter(Mandatory=$true)] [string] $Username,
        [Parameter(Mandatory=$true)] [string] $Password,
        [Parameter(Mandatory=$true)] [string] $FullName,
        [Parameter(Mandatory=$false)] [string] $Description = ""
    )

    # 檢查帳號是否已經存在
    $userExists = Get-WMIObject -Class Win32_UserAccount | Where-Object { $_.Name -eq $Username }

    if ($userExists) {
        Write-Host "帳號 '$Username' 己存在。"
        return
    }

    # 創建新帳號
    $UserPrincipal = New-Object System.Security.Principal.NTAccount($env:COMPUTERNAME, $Username)
    $SID = $UserPrincipal.Translate([System.Security.Principal.SecurityIdentifier]).Value

    $PasswordSecure = ConvertTo-SecureString -String $Password -AsPlainText -Force
    $Credential = New-Object System.Management.Automation.PSCredential($SID, $PasswordSecure)
    
    try {
        New-LocalUser -Name $Username -Password $Credential -FullName $FullName -Description $Description
        Write-Host "帳號 '$Username' 己成功創建。"
    } catch {
        Write-Host "無法創建帳號 '$Username': $_.Exception.Message"
    }
}

function CheckAndAddUserToAdminGroup {
    param(
        [Parameter(Mandatory=$true)] [string] $Username
    )
    
    <#
    此函式 CheckAndAddUserToAdminGroup 接受一個參數：帳號名稱 ($Username)。

    在函式內部，我們使用 Get-LocalGroupMember 命令檢查 Administrators 群組的成員。然後，我們使用 Where-Object 過濾出指定帳號名稱的成員。如果找到了該使用者，則表示該使用者已經屬於 Administrators 群組。

    根據結果，函式會顯示相應的訊息。如果使用者不屬於 Administrators 群組，則函式會使用 Add-LocalGroupMember 命令將使用者添加到 Administrators 群組中。

    請確保執行此函式時具有足夠的特權以修改群組成員。

    # 使用範例：
    CheckAndAddUserToAdminGroup -Username "user"
    #>

    # 檢查使用者是否屬於 Administrators 群組
    $isAdmin = (Get-LocalGroupMember -Group "Administrators" | Where-Object { $_.Name -eq $Username })
    
    if ($isAdmin) {
        Write-Host "使用者 '$Username' 已屬於 Administrators 群組。"
    } else {
        Write-Host "使用者 '$Username' 不屬於 Administrators 群組。正在將其加入群組..."
        
        # 將使用者添加到 Administrators 群組
        Add-LocalGroupMember -Group "Administrators" -Member $Username
        
        if (!$?) {
            Write-Host "無法將使用者 '$Username' 添加到 Administrators 群組。請確保具有足夠的特權並重試。"
        } else {
            Write-Host "使用者 '$Username' 成功添加到 Administrators 群組。"
        }
    }
}

function CheckAndUpdatePassword {
    param(
        [Parameter(Mandatory=$true)] [string] $Username,
        [Parameter(Mandatory=$true)] [string] $NewPassword
    )
    <#
    此函式 CheckAndUpdatePassword 接受兩個參數：帳號名稱 ($Username) 和新密碼 ($NewPassword)。

    在函式中，我們使用 Get-WMIObject 來檢查帳號是否存在。如果帳號存在，我們則使用 Set-LocalUser 命令變更該帳號的密碼。如果密碼變更成功，函式會顯示成功訊息；如果發生錯誤，則會顯示錯誤訊息。

    如果帳號不存在，函式會顯示相應的訊息。

    您可以根據需要調用此函式，並傳遞相應的參數值，例如：CheckAndUpdatePassword -Username "user" -NewPassword "Us2791072"。
    
    # 使用範例：
    CheckAndUpdatePassword -Username "user" -NewPassword "Us2791072"
    #>


    # 檢查帳號是否存在
    $userExists = Get-WMIObject -Class Win32_UserAccount | Where-Object { $_.Name -eq $Username }

    if ($userExists) {
        # 變更密碼
        try {
            $UserPrincipal = New-Object System.Security.Principal.NTAccount($env:COMPUTERNAME, $Username)
            $SID = $UserPrincipal.Translate([System.Security.Principal.SecurityIdentifier]).Value

            $NewPasswordSecure = ConvertTo-SecureString -String $NewPassword -AsPlainText -Force
            $Credential = New-Object System.Management.Automation.PSCredential($SID, $NewPasswordSecure)

            Set-LocalUser -Name $Username -Password $Credential

            Write-Host "帳號 '$Username' 的密碼已成功變更。"
        } catch {
            Write-Host "無法變更帳號 '$Username' 的密碼: $_.Exception.Message"
        }
    } else {
        Write-Host "帳號 '$Username' 不存在。"
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

    CreateLocalUser -Username "vhwc" -Password "Vh2791641" -FullName "VHWC MIS" -Description "VHWC MIS Admin account" 
    CheckAndAddUserToAdminGroup -Username "vhwc"   
    CheckAndAddUserToAdminGroup -Username "user"
    CheckAndUpdatePassword -Username "user" -NewPassword "Cyat7322" 

    pause
}