function Update-Software {
    param (
        [string]$SoftwareName,   # 軟體名稱
        [string]$NewVersion,    # 新版本
        [string]$UninstallCommand,  # 卸載指令
        [string]$InstallCommand     # 安裝指令
    )
    
    # 取得目前軟體版本
    $currentVersion = (Get-CimInstance Win32_Product | Where-Object {$_.Name -eq $SoftwareName}).Version
    
    # 如果目前版本小於新版本，則執行卸載舊版本並安裝新版本
    if ($currentVersion -lt $NewVersion) {
        Write-Host "正在移除舊版本 $currentVersion ..."
        Invoke-Expression $UninstallCommand
        Write-Host "正在安裝新版本 $NewVersion ..."
        Invoke-Expression $InstallCommand
    }
    else {
        Write-Host "軟體已經是最新版本。"
    }
}


function compare-version {
    <#
    比對軟體版本用, 
    Version1 大於 Version2, 回傳True
    Version1 小於等於 Version2, 回傳True

    example:
    compare-version -Version1 2.0.4 -Version2 2.0.3
    #>


    param(
        [string]$Version1,
        [string]$Version2
    )

    $v1 = $Version1.Split('.')
    $v2 = $Version2.Split('.')

    foreach ($i in $v1) {
        $id = $v1.indexof($i)
         
        if ($i -gt $v2[$id]) {
            Write-Host $Version1 ">" $Version2
            return $true
        }
    }

    return $false
}



compare-version -Version1 2.2.5 -Version2 0







 # 卸?OneDrive?用程序
 $appPath = "$env:SystemRoot\SysWOW64\OneDriveSetup.exe"
 if (-not (Test-Path $appPath)) {
     $appPath = "$env:SystemRoot\System32\OneDriveSetup.exe"
    
 }
 
 Write-Host $appPath


 Function Disable-NewsAndInterests {
    $registryPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Feeds"
    Set-ItemProperty -Path $registryPath -Name "ShellFeedsTaskbarViewMode" -Value 2
    Stop-Process -Name WindowsShellExperienceHost
  }
  