@echo off
REM Change to script directory to ensure correct working directory
cd /d "%~dp0"

REM Function to find regasm.exe
call :FindRegasm
if "%REGASM_PATH%"=="" (
    echo ERROR: regasm.exe not found!
    echo Please ensure .NET Framework 4.0 or later is installed.
    echo.
    echo Common locations checked:
    echo - %WINDIR%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe
    echo - %WINDIR%\Microsoft.NET\Framework\v4.0.30319\regasm.exe
    pause
    exit /b 1
)

echo ========================================
echo  HFP Thumbnail Handler - Simple Build
echo ========================================
echo.
echo Current directory is %CD%
echo.
echo 1. Build only
echo 2. Build and attach (register COM)
echo 3. Remove (unregister COM)
echo.
set /p choice="Choose option (1-3): "

if "%choice%"=="1" goto BUILD
if "%choice%"=="2" goto BUILD_ATTACH
if "%choice%"=="3" goto REMOVE
echo Invalid choice. Exiting.
pause
exit /b 1

:BUILD
echo.
echo Building project...
REM Verify project file exists
if not exist "HfpThumbnailHandler.csproj" (
    echo ERROR: HfpThumbnailHandler.csproj not found in current directory!
    echo Current directory is %CD%
    echo Make sure you're running this script from the project folder.
    pause
    exit /b 1
)

dotnet clean HfpThumbnailHandler.csproj
dotnet restore HfpThumbnailHandler.csproj
dotnet build HfpThumbnailHandler.csproj --configuration Release /p:RegisterForComInterop=false
if %ERRORLEVEL% NEQ 0 (
    echo Build failed!
    pause
    exit /b %ERRORLEVEL%
)
echo Build completed successfully!
goto END

:BUILD_ATTACH
echo.
echo Building and registering COM component...
REM Check admin privileges
net session >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: COM registration requires Administrator privileges!
    echo Right-click this file and select "Run as administrator"
    pause
    exit /b 1
)

REM Verify project file exists
if not exist "HfpThumbnailHandler.csproj" (
    echo ERROR: HfpThumbnailHandler.csproj not found in current directory!
    echo Current directory is %CD%
    echo Make sure you're running this script from the project folder.
    pause
    exit /b 1
)

REM Build
dotnet clean HfpThumbnailHandler.csproj
dotnet restore HfpThumbnailHandler.csproj
dotnet build HfpThumbnailHandler.csproj --configuration Release /p:RegisterForComInterop=false
if %ERRORLEVEL% NEQ 0 (
    echo Build failed!
    pause
    exit /b %ERRORLEVEL%
)

REM Register COM
echo Registering COM component with %REGASM_PATH%
"%REGASM_PATH%" "bin\Release\net48\HfpThumbnailHandler.dll" /codebase
if %ERRORLEVEL% NEQ 0 (
    echo COM registration failed!
    echo Make sure you're running as Administrator and .NET Framework is properly installed.
    pause
    exit /b %ERRORLEVEL%
)

REM Register file association and shell extension
echo Registering file association and shell extension...
reg add "HKLM\SOFTWARE\Classes\.hfp" /ve /d "HfpFile" /f
reg add "HKLM\SOFTWARE\Classes\.hfp\ShellEx\{e357fccd-a995-4576-b01f-234630154e96}" /ve /d "{A2ECD1CB-B5D5-4136-82CA-9E4D994DABB3}" /f
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved" /v "{A2ECD1CB-B5D5-4136-82CA-9E4D994DABB3}" /d "HFP Thumbnail Provider" /f

REM Clear thumbnail cache
echo Clearing thumbnail cache...
del /q "%LOCALAPPDATA%\Microsoft\Windows\Explorer\thumbcache_*.db" 2>nul

REM Restart Explorer
call :RestartExplorer

echo Build and registration completed successfully!
goto END

:REMOVE
echo.
echo Removing COM registration...
REM Check admin privileges
net session >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: COM unregistration requires Administrator privileges!
    echo Right-click this file and select "Run as administrator"
    pause
    exit /b 1
)

REM Unregister COM
if exist "bin\Release\net48\HfpThumbnailHandler.dll" (
    echo Unregistering COM component with %REGASM_PATH%
    "%REGASM_PATH%" "bin\Release\net48\HfpThumbnailHandler.dll" /unregister
)

REM Remove registry entries
reg delete "HKLM\SOFTWARE\Classes\.hfp" /f 2>nul
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved" /v "{A2ECD1CB-B5D5-4136-82CA-9E4D994DABB3}" /f 2>nul

REM Clear thumbnail cache
echo Clearing thumbnail cache...
del /q "%LOCALAPPDATA%\Microsoft\Windows\Explorer\thumbcache_*.db" 2>nul

REM Restart Explorer
call :RestartExplorer

echo COM component removed successfully!
goto END

:END
echo.
pause
exit /b 0

:FindRegasm
REM Try to find regasm.exe in common locations
set "REGASM_PATH="

REM Try 64-bit first (preferred for 64-bit systems)
if exist "%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe" (
    set "REGASM_PATH=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe"
    echo Found regasm.exe 64-bit: %REGASM_PATH%
    goto :eof
)

REM Try 32-bit as fallback
if exist "%WINDIR%\Microsoft.NET\Framework\v4.0.30319\regasm.exe" (
    set "REGASM_PATH=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\regasm.exe"
    echo Found regasm.exe 32-bit: %REGASM_PATH%
    goto :eof
)

REM Could add more versions here if needed (v4.5, v4.6, etc.)
echo regasm.exe not found in standard locations
goto :eof

:RestartExplorer
echo Restarting Windows Explorer...

REM Check if Explorer is running
tasklist /FI "IMAGENAME eq explorer.exe" 2>nul | find /I "explorer.exe" >nul
if %ERRORLEVEL% EQU 0 (
    echo Stopping Explorer...
    taskkill /f /im explorer.exe >nul 2>&1
    if %ERRORLEVEL% EQU 0 (
        echo Explorer stopped successfully. Waiting for cleanup...
        timeout /t 3 /nobreak >nul
    ) else (
        echo Warning: Failed to stop Explorer gracefully
        timeout /t 1 /nobreak >nul
    )
) else (
    echo Explorer was not running
)

REM Start Explorer
echo Starting Explorer...
start explorer.exe

REM Wait and verify Explorer started (try multiple times)
set RETRY_COUNT=0
:CheckExplorer
timeout /t 1 /nobreak >nul
tasklist /FI "IMAGENAME eq explorer.exe" 2>nul | find /I "explorer.exe" >nul
if %ERRORLEVEL% EQU 0 (
    echo Explorer restarted successfully
    goto :eof
)

set /a RETRY_COUNT+=1
if %RETRY_COUNT% LSS 3 (
    echo Waiting for Explorer to start... (%RETRY_COUNT%/3)
    goto CheckExplorer
)

echo Warning: Explorer may not have started properly
echo If you don't see your desktop, try one of these:
echo   - Press Ctrl+Shift+Esc and click File - Run new task - explorer.exe
echo   - Press Win+R and type explorer.exe
echo   - Log off and back on
goto :eof