<# 20240108
��ݰ��O�p���f�A�]�����O�pvpn�d�����ľ��������A�p�G�����D�A�|�htry COMX1-COMX10�A
�üg�J��C:\NHI\INI\csdll.ini�A�ҥH�|�����W�����E�������D�A�����S���A�ҥH�p�G�G�|���O�T�{�bCOMX1�A
���O�p��ĳ�NC:\NHI\INI\csdll.ini�]����Ū�A�p�G�N���|�o�ͦ��W�����D�A�ҥH @�^�� @���� �T�{�@�U�A
�O�_���N���E���q�����ɮ׬ҳ]����Ū�A�קK�QVPN�ﱼ�C�t�~�p�]��Ū��b����{�����׭q�p�������ɮ��ݩʡA�n�`�N�O�o��^�ӡC
#>

function Parse-IniFile {
    <#
    code from chatgpt
    �o�Ө�ƨϥΤF PowerShell ���i���ѼƳB�z����ӳB�z���|�ѼơA�æbŪ�� .ini �ɮת��C�@��ɶi��A���ѪR�C����٨ϥΤF���Ѩӻ����C�ӳ������\��A���{���X�����z�ѩM���@�C
    #>
    
    # �ϥ� [CmdletBinding()] �ݩʨӶ}�Ҷi�����ѼƳB�z
    # �o���\�b��Ƥ��ϥζi���ѼơA�Ҧp Mandatory�BParameterSetName ��
    [CmdletBinding()]
    param (
        # �ϥ� [Parameter()] �ݩʨӫ��w���n�����|�Ѽ�
        [Parameter(Mandatory = $true)]
        [string]$Path
    )
     
    # �T�O�ɮצs�b
    if (-not (Test-Path $Path)) {
        throw "The file '$Path' does not exist."
    }
     
    # ��l�Ƥ@�ӪŪ������A�Ω�s�x�ѪR�᪺ .ini ���e
    $ini = @{}
     
    # ��l�Ƥ@�ӪŪ��`�I�W���ܼơA�Ω�ѪR�ثe���`�I
    $section = ""
     
    # Ū�� .ini �ɮפ����C�@��
    Get-Content $Path | ForEach-Object {
        # �h���C��e�᪺�Ů�
        $line = $_.Trim()
     
        # �p�G�Ӧ�O�`�I�W�١A�h�ѪR�X�`�I�W�٨ê�l�Ƥ@�ӷs�������
        if ($line -match "^\[.*\]$") {
            $section = $line.Substring(1, $line.Length - 2)
            $ini[$section] = @{}
        }
        # �p�G�Ӧ�O��ȹ�A�h�ѪR�X��M�ȡA�ñN��s�J�ثe�`�I�������
        elseif ($line -match "^([^=]+)=(.*)$") {
            $key = $matches[1].Trim()
            $value = $matches[2].Trim()
            $ini[$section][$key] = $value
        }
    }
     
    # ��^�ѪR�᪺ .ini ���e�����
    return $ini
    
}
    
    
function  save-iniFile {
    #�Nini�g�^�ɮ�
    param (
        [CmdletBinding()]
        $ini,
        [CmdletBinding()]
        $path
    )
    
    $ini_content = ""
    
    foreach ($i in $ini.keys) {
        $ini_content += "[$i] `n"
            
        foreach ($j in $ini.$i.keys) {
            $ini_content += "$j=$($ini.$i.$j.tostring()) `n"
        }
    }
    #Write-Output $ini_content
    Out-File -InputObject $ini_content -FilePath $path
}

function check_comx1 ($ini_path) {
    
    $log_file = "\\172.20.1.14\update\0001-���a�t�����ҳ]�w\set_csdllini.log"

    #�q���W�٭������
    $rule = "wmis-*"
    if ($env:COMPUTERNAME -like $rule) {
        $is_computername_rule = $true
    }
    else {
        $is_computername_rule = $false
    }

    #�ˬd�q���W�٬O�_�ŦXrule ���ˬdini���|, �p�Gini���s�b�]���ζ]
       
    if (Test-Path $ini_path -and $is_computername_rule) {

        #���oini���e
        $ini_content = Parse-IniFile -Path $ini_path 
        #���oini�ɮ��ݩ�
        $ini_fileproperty_readonly = Get-ItemPropertyValue -Path $ini_path -Name "IsReadOnly"

        Write-Host "COM value is ""$($ini_content.CS.COM)"""
    
        #�p�G���OCOMX1, �N�令COMX1
        if ($ini_content["CS"]["COM"] -ne "COMX1") {

            #�p�G����Ū���Ѷ}
            if ($ini_fileproperty_readonly -eq $true) {
                Set-ItemProperty -Path $ini_path -Name IsReadOnly -Value $false -Force
                #�粒�A�����s�oini�ɮ��ݩ�
                $ini_fileproperty_readonly = Get-ItemPropertyValue -Path $ini_path -Name "IsReadOnly"
            }

            Write-Host "Change COM value to ""COMX1""" -ForegroundColor Red
            $ini_content["CS"]["COM"] = "COMX1"

            #�s��
            save-iniFile -ini $ini_content -path $ini_path

            #�g�@�Ulog
            $log_string = "set csdll.ini COMX1: $env:COMPUTERNAME,$(Get-Date)" 
            $log_string | Add-Content -PassThru $log_file

        }

        #���O��Ū�������Ū
        if ($ini_fileproperty_readonly -eq $false) {
            Set-ItemProperty -Path $ini_path -Name IsReadOnly -Value $true -Force

            #�g�@�Ulog
            $log_string = "set csdll.ini readonly: $env:COMPUTERNAME,$(Get-Date)" 
            $log_string | Add-Content -PassThru $log_file

        }
    }
}

$ini_path = "C:\nhi\ini\csdll.ini"
check_comx1 -ini_path $ini_path