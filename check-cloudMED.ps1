#雲端藥歷
#雲端藥歷有2個位置, 一個在c:\cloudMED, 一個在c:\vghtc\cloudMED (HIS call cloudMED)
#2030809, 112年新PC發生HIS Call cloudMED時, cloudMED會當掉. 原因為c:\vghtc\cloudMED\remoteserver.xml設定值有誤.


param($runadmin)

function check-cloudMED {

    $xmlfiles = "C:\cloudMED\RemoteServer.dat",
    "C:\vghtc\cloudMED\RemoteServer.dat"

    $remoteServervalue = "opd.vghb12.vhwc.gov.tw"            

    foreach ($xmlfile in $xmlfiles) {

        $is_fileexist = Test-Path -Path $xmlfile

        if ($is_fileexist) {
            $xml = [xml](Get-Content -Path $xmlfile)

            $check_value = $xml.NewDataSet.SYSASLT.SERVERNAME -eq $remoteServervalue

            if ($check_value -eq $false) {

                $xml.NewDataSet.SYSASLT.SERVERNAME = $remoteServervalue
                $xml.save($xmlfile)
                Write-Output "發現雲端藥歷remoteserver.dat設定值不是 $remoteServervalue, 己進行修改"
            }
            else {
                Write-Output "雲端藥歷設定值正確: $xmlfile : $($xml.NewDataSet.SYSASLT.SERVERNAME) "
            }

        }
        else {
            Write-Error "找不到底下檔案,醫療系統HIS可能會出錯, 請檢查 $xmlfile"
        }

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

    Check-cloudMED
    
    pause
}