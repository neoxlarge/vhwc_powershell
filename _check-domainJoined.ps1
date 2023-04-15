function Check-DomainJoined {
    <#
    .SYNOPSIS
    檢查計算機是否已加入網域。

    .DESCRIPTION
    此函數使用 [System.DirectoryServices.ActiveDirectory.Domain] 的靜態屬性 "GetCurrentDomain" 和 "GetComputerDomain" 來判斷計算機是否加入了網域。如果計算機已加入網域，它會顯示一條綠色的消息，否則它會顯示一條紅色的消息。

    .EXAMPLE
    Check-DomainJoined

    這個命令將檢查計算機是否已加入網域。

    .NOTES
    作者：ChatGPT
    最後修改日期：2023年3月23日
    #>

    $domain = ([System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()).Name
    $computer = $env:COMPUTERNAME
    $computerDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain().Name
    
    if ($computerDomain -eq $domain) {
        Write-Host "$computer 已經加入網域 $domain。" -ForegroundColor Green
    }
    else {
        Write-Host "$computer 尚未加入網域。" -ForegroundColor Red
    }
}


Check-DomainJoined