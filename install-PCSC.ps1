# 安裝雲端安控元件健保卡讀卡機控制(PCSC)
# 此安裝必須在c:\vghtc安裝完成之後, 因為會復制檔案過去.
# 20230704 update to 5.1.5.7
param($runadmin)

Import-Module ((Split-Path $PSCommandPath) + "\get-installedprogramlist.psm1")

function install-PCSC {
    # 安裝雲端安控元件健保卡讀卡機控制(PCSC)

    $software_name = "健保卡讀卡機控制(PCSC)*"
    $software_path = get-item -Path "\\172.20.5.187\mis\23-讀卡機控制軟體\CMS_CS5.1.5.7_20220925\CS5.1.5.7版_20220925"
    #特別復制到c:\vghtc\00_mis保存.
    $software_copyto_path = "C:\VGHTC\00_mis" 
   
    $all_installed_program = get-installedprogramlist
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }
   

    if ($software_is_installed -eq $null) {
    
  
        #復制檔案到"C:\VGHTC\00_mis"
        Copy-Item -Path $software_path.FullName -Destination $software_copyto_path -Recurse -Force  -Verbose

        Write-OutPut ("Start to install software: " + $software_name)
    
        ## 1. 先安裝VC++可轉發套件
        $software_exec = "gCIE_Setup\vcredist_x86\vcredist_x86.exe"
        Start-Process -FilePath ($software_copyto_path + "\" + $software_path.Name + "\" + $software_exec) -ArgumentList "/passive" -Wait

        Start-Sleep -Seconds 3

        ## 2.安裝雲端元件
        $software_exec = "gCIE_Setup\gCIE_Setup.msi"
        Start-Process -FilePath "msiexec.exe" -ArgumentList "/i $($software_copyto_path + "\" + $software_path.Name + "\" + $software_exec) /passive " -Wait
    
        ## 3.要跑設定檔 bat檔, 為了不被原本bat內的pause卡住, 重寫一次powershell版本.

        Write-Output "執行讀卡機控制軟體的設定bat"
        
        #1.切換為晶片讀卡機版本

        $setup_file_ = @(
            "C:\VGHTC\ICCard\CsHis.dll",
            "C:\ICCARD_HIS\CsHis.dll",
            "C:\vhgp\ICCard\CsHis.dll"
        )

        foreach ($i in $setup_file_) {
            Copy-Item -Path "C:\NHI\LIB\CSHIS.dll" -Destination $i -Force
            $i_version = Get-ItemProperty -Path $i
            Write-Output ("Check dll: " + $i_version.FullName + " Version: " + $i_version.VersionInfo.ProductVersion )
        }

        #2.copy灣橋SAM檔-至指定位置
        $sam_path1 = "$software_copyto_path\$($software_path.Name)\0640140012001000005984.SAM"
        $sam_path2 = "C:\NHI\SAM\COMX1\0640140012001000005984.SAM"
        Copy-Item -Path $sam_path1 -Destination $sam_path2 -Force
    
        #3. 雲端安全模組-放到all-user啟動
        $link_path1 = "C:\Users\Public\Desktop\雲端安全模組主控台.lnk"
        $link_paht2 = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\雲端安全模組主控台.lnk"
        Copy-Item -Path $link_path1 -Destination $link_paht2 -Force

        #5. 5-copy-dll-to-c-v5155 , 就是復制"C:\VGHTC\00_mis\CMS_CS5.1.5.5-讀卡機控制軟體\copy-to-C\ICCARD_HIS"裡所有dll到3個資料?.
        $setup_file_ = Get-ChildItem -Path "C:\NHI\LIB"
    
        $setup_file_target_path = @(
            "C:\ICCARD_HIS",
            "C:\Windows\System32",
            "C:\Windows\System"    
        )

        foreach ($i in $setup_file_) {
            Write-Output ("dll name: " + $i.Name + "dll versoin: " + $i.VersionInfo.ProductVersion    )

            foreach ($j in $setup_file_target_path) {
                copy-item -Path $i.FullName -Destination ($j + "\" + $i.Name)
                $j_version = Get-ItemProperty -Path ($j + "\" + $i.Name)
                Write-Output ("Check dll: " + $j_version.FullName + " Version: " + $j_version.VersionInfo.ProductVersion )
            }
            Write-Output "`n" #跳行一下
        }
        
        #安裝完, 再重新取得安裝資訊
        $all_installed_program = get-installedprogramlist
        $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }
    }
 
    Write-Output ("Software has installed: " + $software_is_installed.DisplayName)
    Write-Output ("Version: " + $software_is_installed.DisplayVersion)


}



#檔案獨立執行時會執行函式, 如果是被匯入時不會執行函式.
if ($run_main -eq $null) {

    #檢查是否管理員
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #如果非管理員, 就試著run as admin, 並傳入runadmin 參數1. 因為在網域一般使用者永遠拿不是管理員權限, 會造成無限重跑. 此參數用來輔助判斷只跑一次. 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    
    }

    if ($check_admin) { 
        install-PCSC
    }
    else {
        Write-Warning "無法取得管理員權限來安裝軟體, 請以管理員帳號重試."
    }
    pause
}