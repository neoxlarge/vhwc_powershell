$url = "http://allapp.vhwc.gov.tw/LcsApiCGA/Api/LcsApi/GetLcsBaseData"

$header = @{
    "content-Type" = "application/x-www-form-urlencoded"
}

$body =  @{
    "Token"="nUqnViSJ+GXiq8YJvmIUuCPelGF2YfhUojEbH9SU0rk="
    "CaseIDNo"="T122432898"  # �����Ҧr��
    "BeginDate"="20221124"   #// �}�l���
    "EndDate"="20241109"      #// �������
}

$respone = Invoke-RestMethod -Uri $url -Method Post -Body $body -Headers $header