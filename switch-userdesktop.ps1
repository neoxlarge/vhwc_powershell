#因為醫師要求的桌面解析度不同,連帶要修改Pacs的螢幕配置.
#院長要求字體大, 所以解析度調低
#孫醫師要求解析度調高, 所以欄位才可見.

#AD上要建立2個帳號, 一個是最佳解析度:OPD-C09HighResolution,一個是大字體:OPD-PC09BigFont, 密碼設一樣
#電腦上要開powershel(admin mode), 執行指令,變更預設執行規則 Set-ExecutionPolicy -ExecutionPolicy RemoteSigned, 讓系統可以執行.ps1檔

#準備2個pacs的營幕設定檔,IRIS_Sys_Highresolution.ini 和 IRIS_Sys_BigFont.ini, 依使手者覆蓋底下設定檔.
#C:\TEDPC\SmartIris\SysIni\IRIS_Sys.ini

$ini_list = @{
    highresolution = "C:\TEDPC\SmartIris\SysIni\IRIS_Sys_Highresolution.ini";
    bigfont = "C:\TEDPC\SmartIris\SysIni\IRIS_Sys_BigFont.ini";
    original = "C:\TEDPC\SmartIris\SysIni\IRIS_Sys.ini"
}

switch ($env:USERNAME) {
    OPD-C09HighResolution { 
        Copy-Item -Path $ini_list.highresolution -Destination $ini_list.original -Force
     }
    OPD-PC09BigFont { 
        copy-item -Path $ini_list.bigfont -Destination $ini_list.original -Force
    }
    Default {}
}
