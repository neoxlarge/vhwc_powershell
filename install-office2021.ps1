# �w��office 2021
# office 2021 ���w�˥ثe���令�ݭn�� office-deployment-tool �w��.
# ²��ӻ��N�Osetup.exe �t�X�]�w�� .xml �Ӧw��.
# �Ӹ`�аѦҩ��U�s��
# https://learn.microsoft.com/zh-tw/microsoft-365-apps/deploy/overview-office-deployment-tool
# xml�̪��]�w�i�H:
# 1. �w�˨ӷ��ɮץi�H��������W�����|, ���ΤU���쥻��.
# 2. �۰ʿ�J�Ǹ�, �����ޱ��t��, ����ĳ�ϥ�.


# ���|�]�w
$office_source_path = "\\172.20.5.187\mis\02-Office\Office_LTSC_2021_std_64bit"
$odt_exe = "setup.exe"
$odt_xml = "install_config.xml"

# ���o�޲z���v��
$credential = Get-Credential -UserName "vhcy\vhwcmis" -Message "�п�J�޲z���b���K�X"

# ���JUI���ե�
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# �ˬd�v���O�_���T
try {
    $check_admin = new-PSDrive -Name check_admin -PSProvider FileSystem -Root "\\172.20.5.187\mis\02-office" -Credential $credential -ErrorAction stop
}
catch {
    # �u�X���~�T��
    
    [System.Windows.Forms.MessageBox]::Show(
        "�v������, �нT�{�޲z���v�����T", # �T�����e
        "���~", # �������D
        [System.Windows.Forms.MessageBoxButtons]::OK, # ���s
        [System.Windows.Forms.MessageBoxIcon]::Error # �ϥ�
    )

    write-output "�v������, �нT�{�޲z���v�����T"
    exit
}
finally {
    Remove-PSDrive -Name check_admin -ErrorAction SilentlyContinue
}



# �w��office 2021
Start-Process -FilePath "$office_source_path\$odt_exe" -ArgumentList "/configure $office_source_path\$odt_xml" -Credential $credential