function Update-Software {
    param (
        [string]$SoftwareName,   # �n��W��
        [string]$NewVersion,    # �s����
        [string]$UninstallCommand,  # �������O
        [string]$InstallCommand     # �w�˫��O
    )
    
    # ���o�ثe�n�骩��
    $currentVersion = (Get-CimInstance Win32_Product | Where-Object {$_.Name -eq $SoftwareName}).Version
    
    # �p�G�ثe�����p��s�����A�h��������ª����æw�˷s����
    if ($currentVersion -lt $NewVersion) {
        Write-Host "���b�����ª��� $currentVersion ..."
        Invoke-Expression $UninstallCommand
        Write-Host "���b�w�˷s���� $NewVersion ..."
        Invoke-Expression $InstallCommand
    }
    else {
        Write-Host "�n��w�g�O�̷s�����C"
    }
}


function compare-version {
    <#
    ���n�骩����, 
    Version1 �j�� Version2, �^��True
    Version1 �p�󵥩� Version2, �^��True

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







 # ��?OneDrive?�ε{��
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
  