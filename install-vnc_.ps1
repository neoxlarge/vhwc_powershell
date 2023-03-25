
function get-installedprogramlist {
    # ���o�Ҧ��w�˪��n��,���U�w�˳n��|�Ψ�.

    ### Win32_product���M��ä�����A Winnexus �ä��b�̭�.
    ### $all_installed_program = Get-WmiObject -Class Win32_Product

    ### �Ҧ����n��|�b���U�o�T�ӵn���ɸ��|��

    $software_reg_path = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    return (Get-ItemProperty -Path $software_reg_path -ErrorAction SilentlyContinue)
}

$all_installed_program = get-installedprogramlist



## �w��VNC
$software_name = "UltraVnc"
$software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like ($software_name + "*") }

if ($null -eq $software_is_installed) {
    Write-Output ("Start to install" + $software_name)

    #�ӷ����| ,�n�_����|,and �w�˰���{���W��
    $software_path = get-item -Path "\\172.20.5.187\mis\08-VNC\1_2_24"
    
    if (Test-Path -Path "d:\mis") {
        $software_copyto_path = "D:\mis"
    }
    else {
        $software_copyto_path = "C:\mis"
    }
    
    
    switch ($env:PROCESSOR_ARCHITECTURE) {
        "AMD64" { $software_exec = "UltraVNC_1_2_24_X64_Setup.exe" }
        "x86" { $software_exec = "UltraVNC_1_2_24_X86_Setup.exe" }
        default { Write-OutPut "Unsupport CPU or OS:"  $env:PROCESSOR_ARCHITECTURE; $software_exec = $null }
    }

    #�_���ɮר�D:\mis
    Copy-Item -Path $software_path -Destination $software_copyto_path -Recurse -Force 

    #installing...
    if ($software_exec) {
        Start-Process -FilePath ($software_copyto_path + "\" + $software_path.Name + "\" + $software_exec) -ArgumentList "/silent /loadinf=installvnc.inf /log=d:\mis\install_vnc.log" -Wait
        Start-Sleep -Seconds 5
    }
              
    #�_��]�w��vltravnc.ini ��C:\Program Files\uvnc bvba\UltraVNC
    Copy-Item -Path ($software_copyto_path + "\" + $software_path.Name + "\ultravnc.ini") -Destination ($env:ProgramFiles + "\uvnc bvba\UltraVNC") -Force

    #�w�˧�, �R���w���ɮ�
    remove-item -Path ($software_copyto_path + "\" + $software_path.Name) -Recurse -Force

    #�w�˧�, �A���s���o�w�˸�T
    $all_installed_program = get-installedprogramlist
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like ($software_name + "*") }

} 

Write-output ("software has installed:" + $software_is_installed.DisplayName )
Write-Output ("Version:" + $software_is_installed.DisplayVersion)

