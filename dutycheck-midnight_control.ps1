# �إ�session��Ÿq���ݮୱ�D�� 172.19.1.24
# �b24����I�ϵ{���s�^���a


$Username = "vhcy\vhwcmis"
$Password = "Mis20190610"
$securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)

$remote_computer = 'remote_WIN2016.vhcy.gov.tw'


$script_block = {

    New-PSDrive -Name "x" -Root "\\172.20.5.185\mis" -PSProvider FileSystem -Persist
    #$proc = Start-Process D:\mis\webdriver\dutycheck-midnight_24.exe -ArgumentList "--png_foldername 'q:\'" -PassThru
   # $proc.WaitForExit()

    Remove-PSDrive -Name "x"
 }

 
 Invoke-Command -ComputerName $remote_computer -ScriptBlock $script_block -Credential $credential
