param($runadmin)


# ��ơG�ˬd�ó]�m DWORD ��
function Set-RegistryDWORD {
    param (
        [string]$Path,
        [string]$Name,
        [int]$Value
    )
    if (!(Test-Path $Path)) {
        New-Item -Path $Path -Force | Out-Null
    }
    $currentValue = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
    if ($currentValue -eq $null -or $currentValue.$Name -ne $Value) {
        Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type DWord
        Write-Output "�w�]�m $Path\$Name �� $Value"
    } else {
        Write-Output "$Path\$Name �w�s�b�B�ȥ��T"
    }
}

# ��ơG�ˬd�ó]�m�r�Ŧ��
function Set-RegistryString {
    param (
        [string]$Path,
        [string]$Name,
        [string]$Value
    )
    if (!(Test-Path $Path)) {
        New-Item -Path $Path -Force | Out-Null
    }
    $currentValue = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
    if ($currentValue -eq $null -or $currentValue.$Name -ne $Value) {
        Set-ItemProperty -Path $Path -Name $Name -Value $Value
        Write-Output "�w�]�m $Path\$Name �� $Value"
    } else {
        Write-Output "$Path\$Name �w�s�b�B�ȥ��T"
    }
}

function enable-iemode {

# �D�n�����U����|

$edgePolicyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Edge"
$popupsAllowedPath = "HKLM:\SOFTWARE\Policies\Microsoft\Edge\PopupsAllowedForUrls"

# �]�m Edge ����
Set-RegistryDWORD -Path $edgePolicyPath -Name "EnterpriseModeSiteListManagerAllowed" -Value 0
Set-RegistryDWORD -Path $edgePolicyPath -Name "InternetExplorerIntegrationLevel" -Value 1
Set-RegistryDWORD -Path $edgePolicyPath -Name "InternetExplorerIntegrationReloadInIEModeAllowed" -Value 1
Set-RegistryString -Path $edgePolicyPath -Name "InternetExplorerIntegrationSiteList" -Value "\\172.20.1.14\update\0001-���a�t�����ҳ]�w\vhwc_win11_IEmode\SiteList.xml"

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
    Set-RegistryString -Path $popupsAllowedPath -Name ($i + 1).ToString() -Value $popupUrls[$i]
}

Write-Output "���U��]�m����"
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