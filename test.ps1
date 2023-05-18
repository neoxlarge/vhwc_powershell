Function Prevent-SleepWmi {
    $Info = Get-CimInstance -ClassName CIM_SessionStatistics
    if (($Info | Select-Object -ExpandProperty IsActive) -contains $true) {
        Return;
    }

    $wmiParams = @{ 
        Namespace = 'root\cimv2\power' 
        Class = 'Win32_PowerSettingDataIndex'
    }
    $sleepSettings = Get-CimInstance @wmiParams -Filter "InstanceID='Microsoft:PowerSetting\{155AAB8B-23B9-4A29-87F0-CPF60A52EEFF}\SUB_NONE_7516b95f-008f-4cb4-9116-DC2EF728B5B6'"
    $displaySettings = Get-CimInstance @wmiParams -Filter "InstanceID='Microsoft:PowerSetting\{3C0BC021-C8A8-4E07-A973-6B14CBCB2B7E}\SUB_NONE_ultimate_power_plan_ac_dll_microsoft_powerplan_6d22571a_03e8_4eed_80b8_6bdd79498fc8'"
    Set-CimInstance -InputObject $sleepSettings[0] -Arguments @{ 
        AcValueIndex = 0
        DcValueIndex = 0
    }
    Set-CimInstance -InputObject $displaySettings[0] -Arguments @{$_.PropertyName = $_.PropertyValue for $_ in $displaySettings.Properties}
}
