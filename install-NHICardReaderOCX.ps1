# �w�ˤW�U�Z��d����
param($runadmin)


function install-NHICardReaderOCX {
    
    #�q��\HKEY_CLASSE�n�ۤv���W�h.
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
        Write-Output "NHICardReaderOCS �v�g�w�ˤF."
    }

    
}


#�ɮ׿W�߰���ɷ|����禡??�p�G�O�Q���J�ɤ��|����禡??J????|????��??J????|????��??J????|????��.
if ($run_main -eq $null) {

    #�ˬd�O�_�޲z��??z????z????z??
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #�p�G�D�޲z��??�N�յ�, ?N???,N?N???�öǤJ as ???Ji�Ѽ� ??�]���b����@��ϥΪ̥û������O�޲z���v��??�|�y���L�����]??���ѼƥΨӻ��U�P�_�u�]�@��|?y???L?????]. ?????�Z???U?P?_?u?]?@??|?y???L?????]. ?????�Z???U?P?_?u?]?@??|?y???L?????]. ?????�Z???U?P?_?u?]?@??. 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    
    }
    else {
         
        install-NHICardReaderOCX
    }
    
    pause
}