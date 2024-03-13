$url = "http://allapp.vhwc.gov.tw/LcsApiCGA/Api/LcsApi/GetLcsBaseData"

$header = @{
    "content-Type" = "application/x-www-form-urlencoded"
}

$body =  @{
    "Token"="nUqnViSJ+GXiq8YJvmIUuCPelGF2YfhUojEbH9SU0rk="
    "CaseIDNo"="T122432898"  # 身分證字號
    "BeginDate"="20221124"   #// 開始日期
    "EndDate"="20241109"      #// 結束日期
}

$respone = Invoke-RestMethod -Uri $url -Method Post -Body $body -Headers $header