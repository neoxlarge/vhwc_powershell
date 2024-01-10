param($runadmin)

#診間的電腦不要進入screen save模式.

Import-Module ((Split-Path $PSCommandPath) + "\Set-ScreenSaver.ps1")
Set-ScreenSaver

function  alwayson-screen {
   
    write-output  "關閉顯示器: 不關閉"
    powercfg /change monitor-timeout-ac 0

}


Set-ScreenSaver -off
alwayson-screen