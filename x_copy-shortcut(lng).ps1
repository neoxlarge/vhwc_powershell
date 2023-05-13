# 復制長照捷徑


function Get-IPv4Address {
    <#
    取得ip的function.
    只能在172.*才能用.
    #>

    $ip = Get-WmiObject -Class Win32_NetworkAdapterConfiguration |
          Where-Object {$_.IPAddress -ne $null -and $_.IPAddress[0] -like "172.*" } |
          Select-Object -ExpandProperty IPAddress |
          Where-Object { $_ -match "\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}" } |
          Select-Object -First 1
    return $ip
}

#1. 檢查捷徑是否存在 c:\user\public\desktop

$result = Test-Path -Path C:\Users\Public\Desktop\abc.link

if (!$result) {
    #捷徑不存在
    copy-item -Path "\\172.20.5.185\mis\abc.link" -Destination "C:\Users\Public\Desktop\abc.link"
    
    #記錄回傳
    [string]$log = "長照捷徑完成: $(Get-Date)  `n"
    $log += "$env:COMPUTERNAME : $(Get-IPv4Address) `n"
    $log += "-------------------------------------- `n"
    Out-File -FilePath "\\172.20.5.185\mis\abc_log.txt" -Append

}