# ���o�q���p�e GUID
$powerPlanGuid = Get-CimInstance -Namespace root\cimv2\power -ClassName Win32_PowerPlan -Filter "IsActive=true"

# �ܧ�ίv�]�w
$standbyTimeoutSettings = Get-WmiObject -Namespace root\cimv2\power -Class Win32_PowerSetting -Filter "InstanceID LIKE '*SLEEP*'"
$hibernateTimeoutSettings = Get-WmiObject -Namespace root\cimv2\power -Class Win32_PowerSetting -Filter "InstanceID LIKE '%HIBER%'"

foreach ($timeout in $standbyTimeoutSettings) {
    if ($timeout.DefaultValue > 0) {
        $args = @($powerPlanGuid, $timeout.InstanceID, 0)
        $result = (Get-WmiObject -Namespace root\cimv2\power -Class Win32_PowerSettingDataIndex).InvokeMethod('SetActivePowerScheme', $args)
        $args = @($null, [UInt32]$timeout.DefaultValue - 1, $null)
        $result = $timeout.InvokeMethod('SetPluggedIn', $args)
        $result = $timeout.InvokeMethod('SetBattery', $args)
    }
}

foreach ($timeout in $hibernateTimeoutSettings) {
    if ($timeout.DefaultValue > 0) {
        $args = @($powerPlanGuid, $timeout.InstanceID, 0)
        $result = (Get-WmiObject -Namespace root\cimv2\power -Class Win32_PowerSettingDataIndex).InvokeMethod('SetActivePowerScheme', $args)
        $args = @($timeout.DefaultValue - 1)
        $result = $timeout.InvokeMethod('HibernateTimeout', $args)
    }
}

# ���U�q����欰
$powerButtonActionSetting = Get-WmiObject -Namespace root\cimv2\power -Class Win32_PowerSetting -Filter "InstanceID LIKE '%BUTTON%' AND InstanceID LIKE '%POWER%'" | Where-Object {$_.ElementName -eq '���U�q�����s'}
if ($powerButtonActionSetting) {
    $args = @(4, 2, 0)
    $result = $powerButtonActionSetting.InvokeMethod('SetAcValueIndex', $args)
    $result = $powerButtonActionSetting.InvokeMethod('SetDcValueIndex', $args)
}

# �����V�X���ίv
$hibernateEnabledSetting = Get-WmiObject -Namespace root\cimv2\power -Class Win32_PowerSetting -Filter "InstanceID LIKE '%HIBERNATE%' AND InstanceID NOT LIKE '%SUB_SLEEP%'"
if ($hibernateEnabledSetting) {
    $args = @(0)
    $result = $hibernateEnabledSetting.InvokeMethod('SetAcValueIndex', $args)
    $result = $hibernateEnabledSetting.InvokeMethod('SetDcValueIndex', $args)
}

# �ίv�]�w���ä�
powercfg /change standby-timeout-ac 0
powercfg /change standby-timeout-dc 0
powercfg /change hibernate-timeout-ac 0
powercfg /change hibernate-timeout-dc 0




$activePowerPlan = Get-CimInstance -Namespace root\cimv2\power -ClassName Win32_PowerPlan -Filter "IsActive=true"
$activePowerPlan.ElementName


Get-WmiObject -Namespace root\cimv2\power -Class win32_powerplan