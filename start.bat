powershell.exe -Command "Set-ExecutionPolicy -ExecutionPolicy Bypass"
powershell.exe %~dp0systest.ps1
powershell.exe -Command "Set-ExecutionPolicy -ExecutionPolicy Restricted"