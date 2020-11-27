@PowerShell.exe -NoProfile -NoLogo -ExecutionPolicy Bypass -File "%~dp0Ophalen-MagisterData.ps1" -Inifile "teamsync.ini" 
@PowerShell.exe -NoProfile -NoLogo -ExecutionPolicy Bypass -File "%~dp0Transformeren-Naar-SchoolDataSync.ps1" -Inifile "teamsync.ini" 
