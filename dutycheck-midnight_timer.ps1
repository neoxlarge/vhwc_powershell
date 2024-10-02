# 建立session到嘉義遠端桌面主機 172.19.1.24
# line token(灣橋檢查群組): HdkeCg1k4nehNa8tEIrJKYrNOeNZMrs89LQTKbf1tbz

# line token(測試1): CclWwNgG6qbD5qx8eO3Oi4ii9azHfolj17SCzIE9UyI
# line token(測試2): AVt3SxMcHhatY2fuG2j6HzKGdb5BOTmrfAlEiBolQOO
# 定時每天晚上11:20分, 和早上0點20分執行.

$Username = "vhcy\vhwcmis"
$Password = "Mis20190610"
$securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)

$remote_computer = 'remote_WIN2016.vhcy.gov.tw'

# 對遠端電腦丟出要執行的指令區塊
Invoke-Command -ComputerName $remote_computer -FilePath .\dutycheck-midnight.ps1 -Credential $credential  

