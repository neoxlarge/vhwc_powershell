# �_����ӱ��|


function Get-IPv4Address {
    <#
    ���oip��function.
    �u��b172.*�~���.
    #>

    $ip = Get-WmiObject -Class Win32_NetworkAdapterConfiguration |
          Where-Object {$_.IPAddress -ne $null -and $_.IPAddress[0] -like "172.*" } |
          Select-Object -ExpandProperty IPAddress |
          Where-Object { $_ -match "\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}" } |
          Select-Object -First 1
    return $ip
}

#1. �ˬd���|�O�_�s�b c:\user\public\desktop

$result = Test-Path -Path C:\Users\Public\Desktop\abc.link

if (!$result) {
    #���|���s�b
    copy-item -Path "\\172.20.5.185\mis\abc.link" -Destination "C:\Users\Public\Desktop\abc.link"
    
    #�O���^��
    [string]$log = "���ӱ��|����: $(Get-Date)  `n"
    $log += "$env:COMPUTERNAME : $(Get-IPv4Address) `n"
    $log += "-------------------------------------- `n"
    Out-File -FilePath "\\172.20.5.185\mis\abc_log.txt" -Append

}