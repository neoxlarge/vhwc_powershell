#載入模組
#部分模組Win10有內建, Win7沒有, 但可以其他電腦匯入.
#如果?入內建的失敗, 就從 wcdc2 (windows server 2012) ?入

function import-module_func ($name) {

    $result = get-module -ListAvailable $name

    if ($result -ne $null) {

        Import-Module -Name $name -ErrorAction Stop

    }
    else {

        $rsession = New-PSSession -ComputerName wcdc2.vhcy.gov.tw -Credential $credential
        Import-Module $name -PSSession $rsession -ErrorAction Stop
    }
    
}