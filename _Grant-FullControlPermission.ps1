#修改三個瀏覽器的預設首頁, IE, Edge, chrome.
#20230329, 網域一般使用者無法寫入HKCU:\SOFTWARE\Policies, 改成由管理者寫到HKLM:\SOFTWARE\Policies
param($runadmin)

function Grant-FullControlPermission {
    <#
    函數名稱為 Grant-FullControlPermission，有兩個變數：Folders（包含要授予完全控制權限的資料夾清單）和 UserName（要授權限的使用者名稱）。
    在函數內部，使用 foreach 迴圈遍歷資料夾清單，對每個資料夾執行相同的操作：取得 ACL、建立存取規則、新增規則至 ACL，然後將修改後的 ACL 套用至資料夾。
    最後，您可以呼叫該函數，填入資料夾清單和使用者名稱，以授予指定的使用者完全控制權限。
    #>

    $folders = "c:\2100", "C:\oracle", "c:\mis", "C:\cloudMED", "C:\ICCARD_HIS", "C:\IDMSClient45", "C:\NHI", "C:\TEDPC", "C:\VGHTC","C:\VghtcLogo", "C:\vhgp"
    $userName = "User"

    foreach ($folderPath in $Folders) {
        # 取得資料夾的 ACL
        $acl = Get-Acl -Path $folderPath
        
        # 建立一個新的存取規則，授予指定使用者完全控制權限
        $rule = New-Object System.Security.AccessControl.FileSystemAccessRule($UserName, "FullControl", "ContainerInherit, ObjectInherit", "None", "Allow")
        
        # 將存取規則新增至 ACL
        $acl.SetAccessRule($rule)
        
        # 將修改後的 ACL 套用至資料夾
        Set-Acl -Path $folderPath -AclObject $acl
    }
}



#檔案獨立執行時會執行函式, 如果是被?入時不會執行函式.
if ($run_main -eq $null) {

    #檢查是否管理員
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #如果非管理員, 就試著run as admin, 並傳入runadmin 參數1. 因為在網域一般使用者永遠拿不是管理員權限, 會造成無限重跑. 此參數用來輔助判斷只跑一次. 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    }

    Grant-FullControlPermission
    pause
}
