
start-sleep -Seconds 120

& 'C:\Program Files\uvnc bvba\UltraVNC\vncviewer.exe' 172.20.5.185 -password sysc0012

Add-Type -AssemblyName System.Windows.Forms
#[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")

rundll32.exe user32.dll,LockWorkStation

start-sleep -Seconds 5
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")

start-sleep -Seconds 5

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class MouseOps {
    [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
    public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);

    private const int MOUSEEVENTF_LEFTDOWN = 0x02;
    private const int MOUSEEVENTF_LEFTUP = 0x04;

    public static void LeftClick()
    {
        mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
    }
}
"@

# °õ¦æ¹«¼Ð¥ªÁäÂIÀ»
[MouseOps]::LeftClick()