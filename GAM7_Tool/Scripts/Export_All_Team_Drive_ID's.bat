@echo off
setlocal

set "GAMDIR=C:\GAM7"
set "GAMEEXE=%GAMDIR%\gam.exe"
set "WORKDIR=C:\GAMWork"
set "OUTPUT=%WORKDIR%\GSuite_TeamDriveIDs.csv"

if not exist "%GAMDIR%" (
    echo ERROR: Required folder "%GAMDIR%" does not exist.
    exit /b 1
)

if not exist "%GAMEEXE%" (
    echo ERROR: "%GAMEEXE%" was not found.
    exit /b 1
)

if not exist "%WORKDIR%" (
    echo ERROR: The required folder "%WORKDIR%" does not exist.
    echo Please create the folder and try again.
    exit /b 1
)

echo Exporting Google Groups...
"%GAMEEXE%" print teamdrives fields id,name > "%OUTPUT%"

if errorlevel 1 (
    echo ERROR: GAM failed to export groups.
    exit /b 1
)

echo Export completed successfully.
exit /b 0

echo.
echo Closing window in 5 seconds...
timeout /t 5 /nobreak >nul
exit
