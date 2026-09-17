@echo off
setlocal
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='SilentlyContinue'; Remove-Item -LiteralPath (Join-Path $env:LOCALAPPDATA 'xExcel Viewer') -Recurse -Force; Remove-Item -LiteralPath (Join-Path ([Environment]::GetFolderPath('Desktop')) 'xExcel Viewer.lnk') -Force; Remove-Item -LiteralPath (Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\xExcel Viewer.lnk') -Force; Remove-Item -LiteralPath 'HKCU:\Software\Classes\xExcelViewer.Workbook' -Recurse -Force; foreach($ext in @('.xlsx','.xlsm','.xltx','.xltm','.xls')){ Remove-ItemProperty -LiteralPath ('HKCU:\Software\Classes\' + $ext + '\OpenWithProgids') -Name 'xExcelViewer.Workbook' -ErrorAction SilentlyContinue }"
echo xExcel Viewer was removed.
pause
