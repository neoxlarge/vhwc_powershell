#Check-VGHTCenv
#此檔執行三個動作. 
#1.C:\VGHTC\00_mis\9-allinone-vghtc-env.reg
#2.C:\VGHTC\00_mis\DLL-for-vhwc.cmd
#3.C:\VGHTC\00_mis\中榮iccard環境變數設定.bat

param($runadmin)


function Check-EnvPathContains($path) {
    # 檢查系統環境變數$env:path是否包含特定路徑
    $envPathList = $env:path -split ";"
    foreach ($p in $envPathList) {
        if ($p -like "*$path*") {
            #系統環境變數中包含$path,不需執行C:\VGHTC\00_mis\中榮iccard環境變數設定.bat 
            return $true
        }
    }
    #系統環境變數中不包含$path,需執行C:\VGHTC\00_mis\中榮iccard環境變數設定.bat 
    return $false
}

function Check-VGHTCenv {
    
    #執行環境設定檔 1
    $setting_file = "C:\VGHTC\00_mis\9-allinone-vghtc-env.reg"
    Write-Output "執行環境設定: $setting_file" 

    if (Test-Path -Path $setting_file) {
        Start-Process -FilePath reg.exe -ArgumentList ("import " + $setting_file) -Wait
    }
    else {
        Write-Warning "設定檔不存在: $setting_file"
    } 

    #執行環境設定檔 2
    $setting_file = "C:\VGHTC\00_mis\DLL-for-vhwc.cmd"
    Write-Output "執行環境設定: $setting_file"

    if (Test-Path $setting_file) {
        $process_id = Start-Process -FilePath $setting_file -PassThru
        Start-Sleep -s 8
    }
    else {
        Write-Warning "設定檔不存在: $setting_file"
    }

    #檢查所有的DLL有無登錄

    #掛載hkey_classes_root 到 HKCR:
    New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT -ErrorAction SilentlyContinue  | Out-Null #-Credential $credential

    $dll_list = @(
        "C:\VGHTC\HISAssembly\FSHCAATL.dll",
        "C:\VGHTC\HISAssembly\FSHCACSAPIATL.dll",
        "C:\VGHTC\HISAssembly\FSVAATL.dll",
        "C:\VGHTC\HISAssembly\FSCAPIATL2.dll",
        "C:\VGHTC\HISAssembly\HCAAPIATL.dll",
        "C:\VGHTC\HISAssembly\HCACSAPIATL.dll"
    )

    $err = 0
    foreach ($i in $dll_list) {
        Write-OutPut ("檢查dll: " + $i )
        $result = Get-ChildItem -Path "HKCR:\TypeLib\*" -Recurse | Where-Object -FilterScript { $_.getvalue("") -eq $i }
        
        if ($result -ne $null) {
            Write-Host "Pass:  " -ForegroundColor Green -NoNewline
            Write-Output ($result.PSPath + "\" + $result.Property + " property: " + $result.GetValue("")) 
        }
        else {
            Write-Warning "dll登錄可能失敗,沒有找到登錄值,請檢查: $i"
            $err += 1
        }
    }
    #如果都沒有錯誤, 結束掉該程序
    if ($err -eq 0) {
        $process_id.Kill()
    }

    Remove-PSDrive -Name HKCR

    #檢查系統環境變數 3
    $setting_file = "C:\VGHTC\00_mis\中榮iccard環境變數設定.bat"
    Write-Output "執行環境設定: $setting_file"
    $path = "C:\VGHTC\ICCard"
    $result = Check-EnvPathContains $path

    if ($result -eq $false) {
    
        Write-Warning "系統環境變數中不包含 C:\VGHTC\ICCard"
        
        if ($check_admin) {
            if (Test-Path $setting_file) {
                Write-Output "執行設定檔: $setting_file"
                Start-Process -FilePath "cmd.exe" -ArgumentList "/c $setting_file"  -Wait
            }
            else {
                Write-Warning "設定檔不存在: $setting_file"
            }
        }
        else {
        
            Write-Warning "沒有系統管理員權限,無法執行$setting_file ,請以系統管理員身分重新嘗試."
        }
    
    }
    elseif ($result -eq $true) {

        Write-Output "系統環境變數中包含$path"
        Write-Output "不需執行 $setting_file "

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

    Check-VGHTCenv
    
    pause
}