@echo off
color 1F
REM Developed by: Rhys Saddul

setlocal

set "GAMDIR=C:\GAM7"
set "GAMEEXE=%GAMDIR%\gam.exe"
set "WORKDIR=C:\GAMWork"
set "OUTPUT=%WORKDIR%\GSuite_TeamDriveIDs.csv"

if not exist "%GAMDIR%" (
    color 4F
    echo ERROR: Required folder "%GAMDIR%" does not exist.
    exit /b 1
)

if not exist "%GAMEEXE%" (
    color 4F
    echo ERROR: "%GAMEEXE%" was not found.
    exit /b 1
)

if not exist "%WORKDIR%" (
    color 4F
    echo ERROR: The required folder "%WORKDIR%" does not exist.
    echo Please create the folder and try again.
    exit /b 1
)

color 2F
echo Exporting Google Team Drives...
"%GAMEEXE%" print teamdrives fields id,name > "%OUTPUT%"

if errorlevel 1 (
    color 4F
    echo ERROR: GAM failed to export groups.
    exit /b 1
)

color 0A
echo Export completed successfully.
echo.
echo Closing window in 5 seconds...
timeout /t 5 /nobreak >nul
exit
