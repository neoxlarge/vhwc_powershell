function check-OPDList {

# �N�ܼ� $json_file �]�w�� JSON �ɮ� "opd_list.json" �����|
$json_file = "opd_list.json"

# Ū�� JSON �ɮת����e�ñN���x�s�b�ܼ� $json_content ��
$json_content = Get-Content -Path $json_file -Raw

# �N JSON ���e�ഫ�� PowerShell ����ñN��������ܼ� $opd_json
$opd_json = ConvertFrom-Json -InputObject $json_content

# ��l���ܼ� $opd�A�ñN��]�w�� null
$opd = $null

# ��X�ŦX�q���W�٪����.
foreach ($o in $opd_json.psobject.properties) {

    $result = $o.Value.name -eq "wnur-opd-pc02"
   
    if ($result) {
        $opd = $o.Value
        break
    }
 
}
}