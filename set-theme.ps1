function Set-ISEDarkMode {
  $RegPath = "HKCU:\SOFTWARE\Microsoft\PowerShell\1\PowerShellEngine"
  $ValueName = "ApplicationPrivateData"
  $ISETheme = @{
    "ConsoleColor2" = "DarkCyan"
    "ConsoleColor3" = "White"
    "ConsoleColor4" = "Yellow"
    "ConsoleColor5" = "Green"
    "ConsoleColor6" = "Magenta"
    "ConsoleColor7" = "DarkGray"
    "ConsoleColor9" = "DarkGray"
    "ConsoleColor10" = "Cyan"
    "ConsoleColor11" = "Gray"
    "ConsoleColor12" = "Yellow"
    "ConsoleColor13" = "Green"
    "ConsoleColor14" = "Magenta"
    "ConsoleColor15" = "White"
    "CommentForegroundColor" = "#ffd700"
    "CommentBackgroundColor" = "#000000"
    "ErrorForegroundColor" = "Red"
    "ErrorBackgroundColor" = "Black"
    "WarningForegroundColor" = "#ff8c00"
    "WarningBackgroundColor" = "Black"
    "DebugForegroundColor" = "#4169E1"
    "DebugBackgroundColor" = "Black"
    "VerboseForegroundColor" = "#32CD32"
    "VerboseBackgroundColor" = "Black"
    "ProgressForegroundColor" = "#ffd700"
    "ProgressBackgroundColor" = "#000000"
    "OutDefaultForegroundColor" = "White"
    "OutDefaultBackgroundColor" = "Black"
    "PipelineVariableForegroundColor" = "#4169E1"
    "PipelineVariableBackgroundColor" = "Black"
    "CommandBorderColor" = "#32CD32"
    "CommandBackgroundColor" = "Black"
    "CommandForegroundColor" = "#32CD32"
    "ResultHighlightingForegroundColor" = "White"
    "ResultHighlightingBackgroundColor" = "#32CD32"
    "PopupBorderColor" = "#32CD32"
    "PopupBackgroundColor" = "#000000"
    "PopupForegroundColor" = "White"
    "RemoteHostBackgroundColor" = "#000000"
    "RemoteHostForegroundColor" = "#ffd700"
  }

  $ISEPath = "$env:UserProfile\Documents\WindowsPowerShell\Microsoft.PowerShellISE_profile.ps1"
  if (!(Test-Path $ISEPath)) { New-Item -Path $ISEPath -ItemType File }
  if (!(Get-ItemProperty -Path $RegPath -Name $ValueName -ErrorAction SilentlyContinue)) {
    New-ItemProperty -Path $RegPath -Name $ValueName -Value @{} | Out-Null
  }

  $ISEData = Get-ItemProperty -Path $RegPath -Name $ValueName
  $ISEData.ApplicationPrivateData.DefaultData.DefaultColors = $ISETheme
  Set-ItemProperty -Path $RegPath -Name $ValueName -Value $ISEData.ApplicationPrivateData
}