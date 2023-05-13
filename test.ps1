function DisableJavaAutoUpdate {
    # �ˬd Java ���w�˸��|
    $javaPath = Get-Command java | Select-Object -ExpandProperty Source

    if ($javaPath) {
        # �c�ئ۰ʧ�s�]�w��󪺸��|
        $deploymentPropertiesPath = Join-Path -Path $javaPath -ChildPath 'lib\deployment.config'

        # �T�{�]�w���O�_�s�b
        if (Test-Path $deploymentPropertiesPath) {
            # Ū���]�w��󪺤��e
            $deploymentProperties = Get-Content $deploymentPropertiesPath -Raw

            # �ˬd�O�_�w�g�]�m�������۰ʧ�s
            if ($deploymentProperties -notmatch 'deployment\.expiration\.check\.enabled') {
                # �b�]�w��󪺥����K�[�T�Φ۰ʧ�s���]�w
                $newDeploymentProperties = $deploymentProperties + "`ndeployment.expiration.check.enabled=false"

                # �g�J��s�᪺�]�w��󤺮e
                $newDeploymentProperties | Set-Content $deploymentPropertiesPath

                Write-Output 'Java �۰ʧ�s�w�g�����C'
            } else {
                Write-Output 'Java �۰ʧ�s�w�g�Q�����C'
            }
        } else {
            Write-Output '�䤣�� Java �� deployment.config ���C'
        }
    } else {
        Write-Output '�䤣��w�w�˪� Java�C'
    }
}
