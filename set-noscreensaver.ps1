param($runadmin)

#�E�����q�����n�i�Jscreen save�Ҧ�.

Import-Module ((Split-Path $PSCommandPath) + "\Set-ScreenSaver.ps1")
Set-ScreenSaver

function  alwayson-screen {
   
    write-output  "������ܾ�: ������"
    powercfg /change monitor-timeout-ac 0

}


Set-ScreenSaver -off
alwayson-screen