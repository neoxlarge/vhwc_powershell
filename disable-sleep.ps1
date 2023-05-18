# 眔筿方璸礶 GUID
$powerPlanGuid = Get-CimInstance -Namespace root\cimv2\power -ClassName Win32_PowerPlan -Filter "IsActive=true"

# 跑何痸砞﹚
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

# 筿方龄︽
$powerButtonActionSetting = Get-WmiObject -Namespace root\cimv2\power -Class Win32_PowerSetting -Filter "InstanceID LIKE '%BUTTON%' AND InstanceID LIKE '%POWER%'" | Where-Object {$_.ElementName -eq '筿方秙'}
if ($powerButtonActionSetting) {
    $args = @(4, 2, 0)
    $result = $powerButtonActionSetting.InvokeMethod('SetAcValueIndex', $args)
    $result = $powerButtonActionSetting.InvokeMethod('SetDcValueIndex', $args)
}

# 闽超睼Α何痸
$hibernateEnabledSetting = Get-WmiObject -Namespace root\cimv2\power -Class Win32_PowerSetting -Filter "InstanceID LIKE '%HIBERNATE%' AND InstanceID NOT LIKE '%SUB_SLEEP%'"
if ($hibernateEnabledSetting) {
    $args = @(0)
    $result = $hibernateEnabledSetting.InvokeMethod('SetAcValueIndex', $args)
    $result = $hibernateEnabledSetting.InvokeMethod('SetDcValueIndex', $args)
}

# 何痸砞﹚ッぃ
powercfg /change standby-timeout-ac 0
powercfg /change standby-timeout-dc 0
powercfg /change hibernate-timeout-ac 0
powercfg /change hibernate-timeout-dc 0




$activePowerPlan = Get-CimInstance -Namespace root\cimv2\power -ClassName Win32_PowerPlan -Filter "IsActive=true"
$activePowerPlan.ElementName


Get-WmiObject -Namespace root\cimv2\power -Class win32_powerplan