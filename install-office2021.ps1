# �w��office 2021
# office 2021 ���w�˥ثe���令�ݭn�� office-deployment-tool �w��.
# ²��ӻ��N�Osetup.exe �t�X�]�w�� .xml �Ӧw��.
# �Ӹ`�аѦҩ��U�s��
# https://learn.microsoft.com/zh-tw/microsoft-365-apps/deploy/overview-office-deployment-tool
# xml�̪��]�w�i�H:
# 1. �w�˨ӷ��ɮץi�H��������W�����|, ���ΤU���쥻��.
# 2. �۰ʿ�J�Ǹ�, �����ޱ��t��, ����ĳ�ϥ�.

# 1. �|���ˬd���L�w��office 2003, �����ܲ���.
# 2. �|���ˬd���L�w��office 2021, �����ܲ���.

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


# �ˬd��LMS office �O�_�w��
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

$office_installed = (get-installedprogramlist) | Where-Object { $_.DisplayName -like "*Microsoft office*" }

if ($office_installed -ne $null) {
    
    $office_installed.displayname
    #$install_confirm = Read-Host "�H�WOffice�w�g�w��, �в�����A�w��, �H�K�X��(Y:��/N:����)"
    # �ЫؽT�{�������
    function Show-InstallConfirmDialog {
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "�w�˽T�{"
        $form.Size = New-Object System.Drawing.Size(450, 500)
        $form.StartPosition = "CenterScreen"
        $form.FormBorderStyle = "FixedDialog"
        $form.MaximizeBox = $false
        $form.MinimizeBox = $false

        # ĵ�i�ϥ�
        $pictureBox = New-Object Windows.Forms.PictureBox
        $pictureBox.Location = New-Object Drawing.Point(20, 20)
        $pictureBox.Size = New-Object Drawing.Size(48, 48)
        $pictureBox.Image = [System.Drawing.SystemIcons]::Warning.ToBitmap()
        $form.Controls.Add($pictureBox)

        # �T������
        $label = New-Object System.Windows.Forms.Label
        $label.Location = New-Object System.Drawing.Point(80, 20)
        $label.Size = New-Object System.Drawing.Size(340, 100)
        $text = $office_installed.displayname + "`n`n �H�WOffice�w�g�w��, �в�����A�w��, �H�K�X���C"
        $label.Text = $text
        #"�H�WOffice�w�g�w��,�в�����A�w��,�H�K�X���C`n`n�O�_�n�~��w��?"
        $label.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10)
        $form.Controls.Add($label)

        # �O���s
        $yesButton = New-Object System.Windows.Forms.Button
        $yesButton.Location = New-Object System.Drawing.Point(100, 420)
        $yesButton.Size = New-Object System.Drawing.Size(100, 30)
        $yesButton.Text = "�O(Y)"
        $yesButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
        $form.Controls.Add($yesButton)

        # �_���s
        $noButton = New-Object System.Windows.Forms.Button
        $noButton.Location = New-Object System.Drawing.Point(250, 420)
        $noButton.Size = New-Object System.Drawing.Size(100, 30)
        $noButton.Text = "�_(N)"
        $noButton.DialogResult = [System.Windows.Forms.DialogResult]::No
        $form.Controls.Add($noButton)

        # �]�w�w�]���s
        $form.AcceptButton = $yesButton
        $form.CancelButton = $noButton

        # �[�J����ƥ�
        $form.KeyPreview = $true
        $form.Add_KeyDown({
                if ($_.KeyCode -eq "Y") {
                    $form.DialogResult = [System.Windows.Forms.DialogResult]::Yes
                    $form.Close()
                }
                elseif ($_.KeyCode -eq "N") {
                    $form.DialogResult = [System.Windows.Forms.DialogResult]::No
                    $form.Close()
                }
            })

        # ��ܵ����è��o���G
        $result = $form.ShowDialog()

        # �^�ǵ��G
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            return "Y"
        }
        else {
            return "N"
        }
    }

    # �ϥγo�Ө�ƴ��N�쥻�� Read-Host
    $install_confirm = Show-InstallConfirmDialog
}   
else {
    $install_confirm = "Y"
}

if ($install_confirm -eq "N") { 
    exit
}

# �w��office 2021
Start-Process -FilePath "$office_source_path\$odt_exe" -ArgumentList "/configure $office_source_path\$odt_xml"