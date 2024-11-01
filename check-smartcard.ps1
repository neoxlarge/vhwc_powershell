# ���ɭ�smartcard �A�ȷ|����, ��smartcard�A�ȥ[�^�h.

function Check-SmartCard {
    $serviceName = "SCardSvr"

    Write-Host "�ˬd Smart Card �A�Ȫ��A..."

    # �ˬd�A�ȬO�_�s�b
    $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue

    if ($null -eq $service) {
        Write-Host "Smart Card �A�Ȥ��s�b�C���ղK�[..."
        try {
            Add-WindowsCapability -Online -Name "SmartCard.DiscreteSignalService~~~~0.0.1.0"
            Write-Host "Smart Card �A�Ȥw���\�K�[�C"
        }
        catch {
            Write-Host "�K�[ Smart Card �A�ȮɥX��: $_"
            return
        }
    }

    # �A������A�Ȫ��A�]�H�����K�[�^
    $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue

    if ($service.StartType -ne 'Automatic') {
        Write-Host "�]�m Smart Card �A�Ȭ��۰ʱҰ�..."
        Set-Service -Name $serviceName -StartupType Automatic
    }

    if ($service.Status -ne 'Running') {
        Write-Host "�Ұ� Smart Card �A��..."
        Start-Service -Name $serviceName
    }

    # �̲��ˬd
    $finalStatus = Get-Service -Name $serviceName
    Write-Host "Smart Card �A�ȷ�e���A:"
    Write-Host "  �Ұ�����: $($finalStatus.StartType)"
    Write-Host "  �B�檬�A: $($finalStatus.Status)"
}

Check-SmartCard