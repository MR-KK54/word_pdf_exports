@echo off
title Install Word ^& PDF Page Exporter Pro for Windows
cls
echo ============================================================
echo   Installing Word ^& PDF Page Exporter Pro on Windows...
echo ============================================================
echo.

set TARGET_DIR=%LOCALAPPDATA%\WordExporterPro
if not exist "%TARGET_DIR%" mkdir "%TARGET_DIR%"

echo Copying application files to %TARGET_DIR%...
xcopy /E /I /Y /Q "%~dp0dist\WordExporterPro.exe" "%TARGET_DIR%\"
if exist "%~dp0static\icons\app_icon.ico" copy /Y "%~dp0static\icons\app_icon.ico" "%TARGET_DIR%\"

echo Creating Desktop shortcut...
powershell -Command "$s=(New-Object -COM WScript.Shell).CreateShortcut('%USERPROFILE%\Desktop\Word & PDF Page Exporter Pro.lnk');$s.TargetPath='%TARGET_DIR%\WordExporterPro.exe';$s.IconLocation='%TARGET_DIR%\app_icon.ico';$s.Save()"

echo Creating Start Menu shortcut...
powershell -Command "$s=(New-Object -COM WScript.Shell).CreateShortcut('%APPDATA%\Microsoft\Windows\Start Menu\Programs\Word & PDF Page Exporter Pro.lnk');$s.TargetPath='%TARGET_DIR%\WordExporterPro.exe';$s.IconLocation='%TARGET_DIR%\app_icon.ico';$s.Save()"

echo.
echo ============================================================
echo   SUCCESS! Installation Complete.
echo   Desktop and Start Menu shortcuts created.
echo ============================================================
echo.
pause
