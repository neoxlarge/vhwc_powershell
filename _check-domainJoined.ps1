function Check-DomainJoined {
    <#
    .SYNOPSIS
    �ˬd�p����O�_�w�[�J����C

    .DESCRIPTION
    ����ƨϥ� [System.DirectoryServices.ActiveDirectory.Domain] ���R�A�ݩ� "GetCurrentDomain" �M "GetComputerDomain" �ӧP�_�p����O�_�[�J�F����C�p�G�p����w�[�J����A���|��ܤ@����⪺�����A�_�h���|��ܤ@�����⪺�����C

    .EXAMPLE
    Check-DomainJoined

    �o�өR�O�N�ˬd�p����O�_�w�[�J����C

    .NOTES
    �@�̡GChatGPT
    �̫�ק����G2023�~3��23��
    #>

    $domain = ([System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()).Name
    $computer = $env:COMPUTERNAME
    $computerDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain().Name
    
    if ($computerDomain -eq $domain) {
        Write-Host "$computer �w�g�[�J���� $domain�C" -ForegroundColor Green
    }
    else {
        Write-Host "$computer �|���[�J����C" -ForegroundColor Red
    }
}


Check-DomainJoined