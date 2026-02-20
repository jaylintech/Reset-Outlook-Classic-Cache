@echo off
setlocal enabledelayedexpansion
title Reset Outlook Classic Cache
echo ============================================
echo        Outlook Classic Cache Reset
echo ============================================
echo.

:: Close Outlook and related processes
echo [1/3] Closing Outlook if running...
taskkill /f /im outlook.exe >nul 2>&1
taskkill /f /im ucmapi.exe >nul 2>&1
taskkill /f /im lync.exe >nul 2>&1
echo     Waiting for processes to fully close...
timeout /t 5 /nobreak >nul

:: Show OST files
echo.
echo [2/3] OST (Offline Cache) files found:
set OUTLOOKPATH=%LOCALAPPDATA%\Microsoft\Outlook
dir /b "%OUTLOOKPATH%\*.ost" 2>nul
echo.

:: Prompt user
choice /c YN /m "Do you want to delete OST file(s)? Outlook will rebuild them"
if errorlevel 2 goto :skip_ost
if errorlevel 1 goto :delete_ost

:delete_ost
echo     Attempting to delete OST file(s)...
for %%f in ("%OUTLOOKPATH%\*.ost") do (
    del /f /q "%%f" >nul 2>&1
    if exist "%%f" (
        echo     %%~nxf is still locked. Scheduling for deletion on next reboot...
        powershell -NoProfile -Command "$p='%%f'; Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public class MFE { [DllImport(\"kernel32.dll\", CharSet=CharSet.Unicode)] public static extern bool MoveFileEx(string a, string b, int c); }'; [MFE]::MoveFileEx($p,$null,4)"
        echo     Reboot your PC and the file will be deleted automatically.
    ) else (
        echo     Deleted: %%~nxf
    )
)
goto :temp_files

:skip_ost
echo     Skipped: OST files kept.

:temp_files
:: Clear Outlook Temporary Files
echo.
echo [3/3] Clearing Outlook Temporary Files...
set OUTLOOKTEMP=%LOCALAPPDATA%\Microsoft\Windows\INetCache\Content.Outlook
if exist "%OUTLOOKTEMP%" (
    for /d %%i in ("%OUTLOOKTEMP%\*") do rd /s /q "%%i" >nul 2>&1
    del /q /f "%OUTLOOKTEMP%\*" >nul 2>&1
    echo     Done: Temporary files cleared.
) else (
    echo     Skipped: Temp folder not found.
)

echo.
echo ============================================
echo  Cache reset complete!
echo  Please restart Outlook.
echo ============================================
echo.
pause
endlocal
