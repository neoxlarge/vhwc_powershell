#設定adboe pdf不自動更新

param($runadmin)

<# 20230906
refer to https://www.adobe.com/devnet-docs/acrobatetk/tools/PrefRef/Windows/Updater-Win.html?zoom_highlight=updates#idkeyname_1_20396

禁止更新按鈕
[HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Adobe\Acrobat Reader\DC\FeatureLockDown]
"bUpdater"=dword:00000000

0: Disables and locks the Updater.
1: No effect.


禁用自動更新
電腦\HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Adobe\Adobe ARM\Legacy\Reader\{AC76BA86-7AD7-1028-7B44-AC0F074E4100}
"Mode"=0

0: Do not download or install updates automatically.
1: Do not download or install updates automatically. Same as 0.
2: Automatically download updates but let the user choose when to install them.
3: Automatically download and install updates.
4: Notify the user downloads are available but do not download them.
#>


<# 底下資料己不適用
Windows Registry Editor Version 5.00

[HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Adobe\Acrobat Reader\DC\FeatureLockDown]
"bUpdateToSingleApp"=dword:00000000

#新版adobe reader改名為adobe acrobat, registry應隨著變更為Adobe Arcobate.
#>

Import-Module ((Split-Path $PSCommandPath) + "\get-installedprogramlist.psm1")

function install-pdf {
    #安裝adobe pdf

    $software_name = "Adobe Acrobat*"
    $zip_path = "\\172.20.5.187\mis\16-PDF\adobe_pdf_離線安裝檔.zip"
    $software_path = $($zip_path.Split("\")[-1])
    $software_msi = "AcroRead.msi"
    $software_msi_update = "AcroRdrDCUpd2300120064.msp"

    ## 找出軟體是否己安裝

    $all_installed_program = get-installedprogramlist
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }

    if ($software_is_installed -eq $null) {

        if (Test-Path $zip_path) {
            Write-Output "Start to install $software_name"
            Write-Output "復制且解壓縮檔案: $zip_path"
            expand-archive -Path $zip_path -DestinationPath $env:TEMP\$software_path -Force
        
            Start-Process -FilePath "msiexec.exe" -ArgumentList "/i $env:TEMP\$software_path\$software_msi /passive /norestart" -wait
            Start-Sleep -Seconds 3
            Start-Process -FilePath "msiexec.exe" -ArgumentList "/update $env:TEMP\$software_path\$software_msi_update /passive /norestart" -wait
            Start-Sleep -Seconds 3

        }
        else {
            Write-Warning "安裝檔路徑不存在,請檢查,: $zip_path"
        }

        #安裝完, 再重新取得安裝資訊
        $all_installed_program = get-installedprogramlist
        $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }

    }

}

function check-pdf {

    $reg_path_disableUpdateButton= @("HKLM:\SOFTWARE\Policies\Adobe\Adobe Acrobat\DC\FeatureLockDown",
        "HKLM:\SOFTWARE\Policies\Adobe\Acrobat Reader\DC\FeatureLockDown")

    $reg_path_disableAutoupdate = "HKLM:\SOFTWARE\WOW6432Node\Adobe\Adobe ARM\Legacy\Reader\{AC76BA86-7AD7-1028-7B44-AC0F074E4100}"

    foreach ($r in $reg_path_disableUpdateButton) {
        if (Test-Path -Path $r) {
            $is_update = Get-ItemProperty -Path $r -Name "bUpdater" -ErrorAction SilentlyContinue

            if ($is_update.bUpdateToSingleApp -eq 0) {
                Write-Output "Adobe PDF Reader 己設定不自動更新."
            }
            else {
                if ($check_admin) {
                    Write-Output "正在設定Adobe PDF Reader 為不自動新更中."
                    Set-ItemProperty -Path $r -Name "bUpdater" -Value 00000000 -type DWord -Force
                    Set-ItemProperty -Path $reg_path_disableAutoupdate -Name "Mode" -Value 0 -Force -type Dword
                }
                else {
                    Write-Warning "沒有系統管理員權限,且Adobe PDF Reader未設定不自動新,請以系統管理員身分重新嘗試."
                }
            }    
        }
    }
}


 
#檔案獨立執行時會執行函式, 如果是被匯入時不會執行函式.
if ($run_main -eq $null) {

    #檢查是否管理員
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #如果非管理員, 就試著run as admin, 並傳入runadmin 參數1. 因為在網域一般使用者永遠拿不是管理員權限, 會造成無限重跑. 此參數用來輔助判斷只跑一次. 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    }
    else {
        install-pdf
    }
    
    check-pdf
    pause
}