#將vhwc_powershell備份到NAS上
robocopy C:\users\73058\OneDrive\文件\vscode_workspace\vhwc_powershell \\172.19.1.229\cch-share\h040_張義明\vhwc_powershell /mir /XD ".git" ".vscode"
robocopy C:\users\73058\OneDrive\文件\vscode_workspace\vhwc_powershell \\172.20.1.122\share\software\00newpc\vhwc_powershell /mir /XD ".git" ".vscode"
