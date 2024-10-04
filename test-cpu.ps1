# �]�w�n�j�����ɶ��d��]�Ҧp�G�̪�1�p�ɡ^
$StartTime = (Get-Date).AddHours(-1)

# �q���ε{����x��������~�ƥ�
$Events = Get-WinEvent -FilterHashtable @{
    LogName = 'Application'
    Level = 2  # 2 �N����~�ŧO
    StartTime = $StartTime
} -ErrorAction SilentlyContinue

# �p�G�����~�ƥ�A�h��ܥ���
if ($Events) {
    Write-Host "�o�{�H�U���ε{�����~�G"
    foreach ($Event in $Events) {
        Write-Host "�ɶ�: $($Event.TimeCreated)"
        Write-Host "�ӷ�: $($Event.ProviderName)"
        Write-Host "�ƥ� ID: $($Event.Id)"
        Write-Host "����: $($Event.Message)"
        Write-Host "------------------------"
    }
} else {
    Write-Host "�b���w���ɶ��d�򤺨S���o�{���ε{�����~�C"
}

# �q�t�Τ�x��������~�ƥ�
$SysEvents = Get-WinEvent -FilterHashtable @{
    LogName = 'System'
    Level = 2  # 2 �N����~�ŧO
    StartTime = $StartTime
} -ErrorAction SilentlyContinue

# �p�G���t�ο��~�ƥ�A�h��ܥ���
if ($SysEvents) {
    Write-Host "�o�{�H�U�t�ο��~�G"
    foreach ($Event in $SysEvents) {
        Write-Host "�ɶ�: $($Event.TimeCreated)"
        Write-Host "�ӷ�: $($Event.ProviderName)"
        Write-Host "�ƥ� ID: $($Event.Id)"
        Write-Host "����: $($Event.Message)"
        Write-Host "------------------------"
    }
} else {
    Write-Host "�b���w���ɶ��d�򤺨S���o�{�t�ο��~�C"
}