function DisableJavaAutoUpdate {
    # 檢查 Java 的安裝路徑
    $javaPath = Get-Command java | Select-Object -ExpandProperty Source

    if ($javaPath) {
        # 構建自動更新設定文件的路徑
        $deploymentPropertiesPath = Join-Path -Path $javaPath -ChildPath 'lib\deployment.config'

        # 確認設定文件是否存在
        if (Test-Path $deploymentPropertiesPath) {
            # 讀取設定文件的內容
            $deploymentProperties = Get-Content $deploymentPropertiesPath -Raw

            # 檢查是否已經設置為取消自動更新
            if ($deploymentProperties -notmatch 'deployment\.expiration\.check\.enabled') {
                # 在設定文件的末尾添加禁用自動更新的設定
                $newDeploymentProperties = $deploymentProperties + "`ndeployment.expiration.check.enabled=false"

                # 寫入更新後的設定文件內容
                $newDeploymentProperties | Set-Content $deploymentPropertiesPath

                Write-Output 'Java 自動更新已經取消。'
            } else {
                Write-Output 'Java 自動更新已經被取消。'
            }
        } else {
            Write-Output '找不到 Java 的 deployment.config 文件。'
        }
    } else {
        Write-Output '找不到已安裝的 Java。'
    }
}
