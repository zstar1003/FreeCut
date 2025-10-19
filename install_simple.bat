@echo off
echo ===============================================
echo FreeCut Plugin Installer
echo ===============================================

REM Set variables
set PLUGIN_NAME=FreeCut
set INSTALL_PATH=C:\FreeCut
set DLL_NAME=FreeCut.dll

echo.
echo 1. Checking build files...
if not exist "bin\Debug\%DLL_NAME%" (
    if not exist "bin\Release\%DLL_NAME%" (
        echo Error: %DLL_NAME% not found
        echo Please build the project first: dotnet build FreeCut.csproj
        pause
        exit /b 1
    ) else (
        set SOURCE_PATH=bin\Release
    )
) else (
    set SOURCE_PATH=bin\Debug
)

echo Build files found in %SOURCE_PATH%!

echo.
echo 2. Creating install directory...
if not exist "%INSTALL_PATH%" mkdir "%INSTALL_PATH%"

echo.
echo 3. Copying plugin files...
copy "%SOURCE_PATH%\*.dll" "%INSTALL_PATH%\" > nul
copy "%SOURCE_PATH%\*.pdb" "%INSTALL_PATH%\" > nul 2>nul
if exist "ribbon.xml" copy "ribbon.xml" "%INSTALL_PATH%\" > nul 2>nul

echo Files copied successfully!

echo.
echo 4. Registering plugin with PowerPoint...

REM Create registry file
echo Windows Registry Editor Version 5.00 > temp_register.reg
echo. >> temp_register.reg
echo [HKEY_CURRENT_USER\Software\Microsoft\Office\PowerPoint\Addins\FreeCut.ThisAddIn] >> temp_register.reg
echo "Description"="FreeCut - PPT Auto Crop PDF Export Plugin" >> temp_register.reg
echo "FriendlyName"="FreeCut" >> temp_register.reg
echo "LoadBehavior"=dword:00000003 >> temp_register.reg
echo "Manifest"="file:///%INSTALL_PATH:\=/%/FreeCut.dll.manifest" >> temp_register.reg

REM Import registry
regedit /s temp_register.reg
del temp_register.reg

echo Plugin registered successfully!

echo.
echo 5. Creating manifest file...
echo ^<?xml version="1.0" encoding="utf-8"?^> > "%INSTALL_PATH%\FreeCut.dll.manifest"
echo ^<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0"^> >> "%INSTALL_PATH%\FreeCut.dll.manifest"
echo   ^<assemblyIdentity name="FreeCut" version="1.0.0.0" type="win32" /^> >> "%INSTALL_PATH%\FreeCut.dll.manifest"
echo   ^<file name="FreeCut.dll" hashalg="SHA1"^> >> "%INSTALL_PATH%\FreeCut.dll.manifest"
echo     ^<comClass clsid="{12345678-1234-1234-1234-123456789ABC}" threadingModel="Both" /^> >> "%INSTALL_PATH%\FreeCut.dll.manifest"
echo   ^</file^> >> "%INSTALL_PATH%\FreeCut.dll.manifest"
echo ^</assembly^> >> "%INSTALL_PATH%\FreeCut.dll.manifest"

echo Manifest file created successfully!

echo.
echo ===============================================
echo Installation Complete!
echo ===============================================
echo.
echo Install location: %INSTALL_PATH%
echo.
echo Next steps:
echo 1. Close all PowerPoint windows
echo 2. Restart PowerPoint
echo 3. Look for "FreeCut" tab in the Ribbon
echo.
echo If plugin doesn't appear, check:
echo - PowerPoint Trust Center settings
echo - Macro security level settings
echo - Add-in loading behavior settings
echo.
echo Starting PowerPoint for testing...
start "" "powerpnt.exe"

echo.
echo Press any key to exit...
pause > nul