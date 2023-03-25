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
