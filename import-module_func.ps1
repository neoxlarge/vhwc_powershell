#���J�Ҳ�
#�����Ҳ�Win10������, Win7�S��, ���i�H��L�q���פJ.
#�p�G?�J���ت�����, �N�q wcdc2 (windows server 2012) ?�J

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