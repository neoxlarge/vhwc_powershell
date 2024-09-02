# �]�wWin11 IE mode ���~�M��Ҧ�, �L�ɶ�����.

param($runadmin)

# ��ơG�ˬd�ó]�m���U���
function Set-RegistryValue {
    param (
        [string]$Path,
        [string]$Name,
        $Value,
        [string]$Type
    )

    $params = @{
        Path  = $Path
        Name  = $Name
        Value = $Value
        Type  = $Type
    }

    if (!(Test-Path $Path) ) {
        if ($check_admin) {
            New-Item -Path $Path -Force | Out-Null
        }
        else {
            Write-Output "���U����|���s�b: $path, �L�޲z���v���s�W."
        }
    } 

    $currentValue = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
    
    if ($null -eq $currentValue -or $currentValue.$Name -ne $Value) {
        if ($check_admin) {
            Set-ItemProperty @params
            Write-Output "���U����X�w�]�m $Path\$Name �� $Value"
        }
        else {
            Write-Output "���X�����T $path \ $Name : $($currentValue.Value)"
        }
    }
    else {
        Write-Output "���X���T $Path\$Name : $Value"
    }
}

function enable-iemode {

    if ((Get-CimInstance -ClassName Win32_OperatingSystem).Caption -match "Windows 11") {
    
        Write-Output "�]�wEdge�ҥ�IE�Ҧ�(���~�M��Ҧ�):"

        # �D�n�����U����|

        $edgePolicyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Edge"
        $popupsAllowedPath = "HKLM:\SOFTWARE\Policies\Microsoft\Edge\PopupsAllowedForUrls"

        # �]�m Edge ����
        Set-RegistryValue -Path $edgePolicyPath -Name "EnterpriseModeSiteListManagerAllowed" -Value 0 -Type DWord
        Set-RegistryValue -Path $edgePolicyPath -Name "InternetExplorerIntegrationLevel" -Value 1 -Type DWord
        Set-RegistryValue -Path $edgePolicyPath -Name "InternetExplorerIntegrationReloadInIEModeAllowed" -Value 1 -Type DWord
        Set-RegistryValue -Path $edgePolicyPath -Name "InternetExplorerIntegrationSiteList" -Value "\\172.20.1.14\update\0001-���a�t�����ҳ]�w\vhwc_win11_IEmode\SiteList.xml" -Type String

        # �]�m�u�X���f���\�C��
        $popupUrls = @(
            "[*.]vhcy.gov.tw",
            "[*.]vhwc.gov.tw",
            "172.19.[.*][.*]",
            "172.20.[.*][.*]",
            "[*.]vghtc.gov.tw",
            "172.19.[.*][.*]:9090",
            "172.19.[.*][.*]:8000",
            "172.20.[.*][.*]:9090",
            "172.20.[.*][.*]:8000"
        )

        for ($i = 0; $i -lt $popupUrls.Length; $i++) {
            Set-RegistryValue -Path $popupsAllowedPath -Name ($i + 1).ToString() -Value $popupUrls[$i] -Type String
        }

        Write-Output "���U��]�m����"
    } else {
        Write-Output "�t�Τ��Owin11, ���ϥ�EDGE IE MODE ���~�M��Ҧ�."
    }
}


#�ɮ׿W�߰���ɷ|����禡, �p�G�O�Q���J�ɤ��|����禡.
if ($run_main -eq $null) {

    #�ˬd�O�_�޲z��
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #�p�G�D�޲z��, �N�յ�run as admin, �öǤJrunadmin �Ѽ�1. �]���b����@��ϥΪ̥û������O�޲z���v��, �|�y���L�����]. ���ѼƥΨӻ��U�P�_�u�]�@��. 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    }

    enable-iemode 

    #pause
    Start-Sleep -Seconds 10
}