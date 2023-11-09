Write-Host "active the office 2020"
$Username = "vhcy\vhwcmis"
$Password = "Mis20190610"
$securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)

switch ($env:PROCESSOR_ARCHITECTURE) {
    "AMD64" { $officeospp_path = "C:\Program Files (x86)\Microsoft Office\Office14\OSPP.VBS"  }
    "x86" { $officeospp_path = "C:\Program Files\Microsoft Office\Office14\OSPP.VBS" }
    
}

$office_key = Read-Host "input office key"

Write-Host "input key to office"
$result = Start-Process -FilePath cscript.exe -ArgumentList """$officeospp_path"" /inpkey:""$office_key""" -NoNewWindow -Credential $credential -PassThru -RedirectStandardOutput "$($env:temp)\officeospp.txt"
#$result = Start-Process -FilePath cscript.exe -ArgumentList """$officeospp_path"" /?" -NoNewWindow -Credential $credential -PassThru -RedirectStandardOutput c:\officeospp.txt
$result.WaitForExit()
Get-Content -Path "$($env:temp)\officeospp.txt"
Start-Sleep -s 3

$result_act = Start-Process -FilePath cscript.exe -ArgumentList """$officeospp_path"" /act" -NoNewWindow -Credential $credential -PassThru -RedirectStandardOutput "$($env:temp)\officeact.txt"
$result_act.WaitForExit()
Get-Content -Path "$($env:temp)\officeact.txt"
Start-Sleep -s 3

pause