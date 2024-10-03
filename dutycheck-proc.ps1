$cred = Get-Credential
$option = New-CimSessionOption -Protocol DCOM
$session = New-CimSession -ComputerName ���ݥD��IP�ΦW�� -Credential $cred -SessionOption $option

Get-CimInstance -ClassName Win32_Process -Filter "Name like 'run1' or Name like 'run2' or Name like 'run3'" -CimSession $session | 
Select-Object Name, ProcessId, @{N='CPU';E={$_.UserModeTime}}, @{N='Memory';E={$_.WorkingSetSize}}

Remove-CimSession $session