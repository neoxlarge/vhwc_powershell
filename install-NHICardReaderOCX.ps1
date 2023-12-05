# 安裝上下班刷卡元件
param($runadmin)


function install-NHICardReaderOCX {
    
    #電腦\HKEY_CLASSE要自己掛上去.
    New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT


    $Username = "vhcy\vhwcmis"
    $Password = "Mis20190610"
    $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)

    $reg_classid_path = "HKCR:\WOW6432Node\CLSID\{1BFA1079-2761-4FF6-8499-5D886F7D972E}"
    $software_path = "\\172.20.5.187\mis\36-NHICardReaderOCX\NHICardReaderOCX.zip"
    
    
    if (!(Test-Path -path $reg_classid_path )) {
        #copy software to temp folder   
        Expand-Archive -Path $software_path -DestinationPath "$($env:temp)\ocx" -Force
        if ($check_admin) {
            $run_processor = Start-Process -FilePath "regsvr32.exe" -ArgumentList "/s $($env:temp)\ocx\NHICardReaderOCX.ocx" -NoNewWindow -PassThru
            $run_processor.WaitForExit()
        }
        else {
            $run_processor = Start-Process -FilePath "regsvr32.exe" -ArgumentList "/s $($env:temp)\ocx\NHICardReaderOCX.ocx" -NoNewWindow -Credential $credential -PassThru
            $run_processor.WaitForExit()
        }
        
        
    }
    else {
        Write-Output "NHICardReaderOCS 己經安裝了."
    }

    
}


#檔案獨立執行時會執行函式??如果是被載入時不會執行函式??J????|????禡??J????|????禡??J????|????禡.
if ($run_main -eq $null) {

    #檢查是否管理員??z????z????z??
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #如果非管理員??就試著, ?N???,N?N???並傳入 as ???Ji參數 ??因為在網域一般使用者永遠拿不是管理員權限??會造成無限重跑??此參數用來輔助判斷只跑一次|?y???L?????]. ?????Ψ???U?P?_?u?]?@??|?y???L?????]. ?????Ψ???U?P?_?u?]?@??|?y???L?????]. ?????Ψ???U?P?_?u?]?@??. 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    
    }
    else {
         
        install-NHICardReaderOCX
    }
    
    pause
}