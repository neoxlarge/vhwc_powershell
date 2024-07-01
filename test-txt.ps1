Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# 創建主窗口
$form = New-Object System.Windows.Forms.Form
$form.Text = "文件搜索工具"
$form.Size = New-Object System.Drawing.Size(800,600)
$form.StartPosition = "CenterScreen"

# 創建搜索字串標籤和輸入框
$labelSearch = New-Object System.Windows.Forms.Label
$labelSearch.Location = New-Object System.Drawing.Point(10,20)
$labelSearch.Size = New-Object System.Drawing.Size(100,20)
$labelSearch.Text = "搜索字串:"
$form.Controls.Add($labelSearch)

$textBoxSearch = New-Object System.Windows.Forms.TextBox
$textBoxSearch.Location = New-Object System.Drawing.Point(120,20)
$textBoxSearch.Size = New-Object System.Drawing.Size(400,20)
$form.Controls.Add($textBoxSearch)

# 創建目錄路徑標籤和輸入框
$labelPath = New-Object System.Windows.Forms.Label
$labelPath.Location = New-Object System.Drawing.Point(10,50)
$labelPath.Size = New-Object System.Drawing.Size(100,20)
$labelPath.Text = "目錄路徑:"
$form.Controls.Add($labelPath)

$textBoxPath = New-Object System.Windows.Forms.TextBox
$textBoxPath.Location = New-Object System.Drawing.Point(120,50)
$textBoxPath.Size = New-Object System.Drawing.Size(400,20)
$form.Controls.Add($textBoxPath)

# 創建瀏覽按鈕
$buttonBrowse = New-Object System.Windows.Forms.Button
$buttonBrowse.Location = New-Object System.Drawing.Point(530,50)
$buttonBrowse.Size = New-Object System.Drawing.Size(75,23)
$buttonBrowse.Text = "瀏覽"
$buttonBrowse.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "選擇搜索目錄"
    if ($folderBrowser.ShowDialog() -eq "OK") {
        $textBoxPath.Text = $folderBrowser.SelectedPath
    }
})
$form.Controls.Add($buttonBrowse)

# 創建搜索按鈕
$buttonSearch = New-Object System.Windows.Forms.Button
$buttonSearch.Location = New-Object System.Drawing.Point(120,80)
$buttonSearch.Size = New-Object System.Drawing.Size(75,23)
$buttonSearch.Text = "搜索"
$buttonSearch.Add_Click({
    $searchString = $textBoxSearch.Text
    $searchPath = $textBoxPath.Text

    if (-not $searchString -or -not $searchPath) {
        [System.Windows.Forms.MessageBox]::Show("請輸入搜索字串和目錄路徑", "錯誤", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    $listView.Items.Clear()

    Get-ChildItem -Path $searchPath -Recurse -Filter *.txt | ForEach-Object {
        $content = [System.IO.File]::ReadAllText($_.FullName, [System.Text.Encoding]::Default)
        if ($content -match $searchString) {
            $item = New-Object System.Windows.Forms.ListViewItem($_.Name)
            $item.SubItems.Add($_.FullName)
            $listView.Items.Add($item)
        }
    }
})
$form.Controls.Add($buttonSearch)

# 創建結果顯示列表
$listView = New-Object System.Windows.Forms.ListView
$listView.Location = New-Object System.Drawing.Point(10,120)
$listView.Size = New-Object System.Drawing.Size(760,430)
$listView.View = [System.Windows.Forms.View]::Details
$listView.FullRowSelect = $true
$listView.Columns.Add("文件名", 150)
$listView.Columns.Add("路徑", 600)
$form.Controls.Add($listView)

# 顯示窗口
$form.ShowDialog()