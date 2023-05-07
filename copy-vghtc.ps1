#復制vghtc
param($runadmin)

function copy-vghtc {
    
    $system_list = @(
        "\\172.20.5.187\mis\11-中榮系統\02-client-PC\cloudMED",
        "\\172.20.5.187\mis\11-中榮系統\02-client-PC\ICCARD_HIS",
        "\\172.20.5.187\mis\11-中榮系統\02-client-PC\IDMSClient45",
        "\\172.20.5.187\mis\11-中榮系統\02-client-PC\VGHTC",
        "\\172.20.5.187\mis\12-vhgp\vhgp"
    )
    
 
    if ($check_admin) {

        Write-Output "復制VGHTC到本機系統."
        foreach ($s in $system_list) {
            Write-Output "Copy $s"
            Copy-Item -Path $s -Destination "$env:HOMEDRIVE\" -Recurse -Force -Verbose
        }
    }
    else {
        Write-Warning "沒有系統管理員權限, 未復制VGHTC."
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

    copy-vghtc
    pause
}