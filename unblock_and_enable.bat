@echo off
echo Checking FreeCut plugin status...
echo.

echo 1. Checking COM registration...
reg query "HKCR\CLSID\{12345678-1234-1234-1234-123456789ABC}\InprocServer32" /v CodeBase
echo.

echo 2. Checking PowerPoint add-in registration...
reg query "HKCU\Software\Microsoft\Office\PowerPoint\Addins\FreeCut.ThisAddIn"
echo.

echo 3. Checking DLL file...
dir C:\FreeCut\FreeCut.dll
echo.

echo 4. Checking if DLL is blocked by Windows...
powershell -Command "Get-Item 'C:\FreeCut\FreeCut.dll' | Get-ItemProperty -Name * | Select-Object PSPath, Length, CreationTime, LastWriteTime"
echo.

echo 5. Unblocking DLL (if needed)...
powershell -Command "Unblock-File -Path 'C:\FreeCut\FreeCut.dll' -ErrorAction SilentlyContinue"
powershell -Command "Unblock-File -Path 'C:\FreeCut\FreeCut.pdb' -ErrorAction SilentlyContinue"
echo Done!
echo.

echo 6. Re-enabling plugin...
powershell -Command "Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\PowerPoint\Addins\FreeCut.ThisAddIn' -Name 'LoadBehavior' -Value 3 -Type DWord"
echo Done!
echo.

echo ================================================
echo Plugin unblocked and re-enabled!
echo Please restart PowerPoint to test.
echo ================================================
pause
