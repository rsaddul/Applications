@echo off
REM Developed by: Rhys Saddul
color 1F

setlocal EnableExtensions EnableDelayedExpansion

REM ==================================================
REM === FETCH EXISTING ORGANIZATIONAL UNITS (LIVE)
REM ==================================================
echo.
echo ================================================
echo Fetching existing OUs from Google Workspace...
echo (This may take a minute on large tenants)
echo ================================================
echo.

REM --- First run: live output
gam print orgs || ( 
    color 4F
    echo ERROR: Failed to fetch existing OUs
    exit /b 1
)

REM --- Second run: capture to file
gam print orgs > existing_ous.tmp || (
    color 4F
    echo ERROR: Failed to capture existing OUs
    exit /b 1
)

echo.
echo OU list retrieved.
echo.

REM ==================================================
REM === ADMIN ACCOUNTS
REM ==================================================
call :EnsureOU "/Admin Accounts" "Administrative user accounts"

REM ==================================================
REM === DEVICES
REM ==================================================
call :EnsureOU "/Devices" "All managed devices"
call :EnsureOU "/Devices/Governors" "Governor devices"
call :EnsureOU "/Devices/Guest" "Guest devices"
call :EnsureOU "/Devices/Staff" "Staff devices"
call :EnsureOU "/Devices/Students" "Student devices"
call :EnsureOU "/Devices/Status" "Device lifecycle status"
call :EnsureOU "/Devices/Status/Decommissioned" "Lost or stolen devices"
call :EnsureOU "/Devices/Status/Deprovisioned" "Retired or replaced devices"

REM ==================================================
REM === LEAVERS
REM ==================================================
call :EnsureOU "/Leavers" "Accounts for users who have left"
call :EnsureOU "/Leavers/Governors" "Governor leaver accounts"
call :EnsureOU "/Leavers/Staff" "Staff leaver accounts"
call :EnsureOU "/Leavers/Students" "Student leaver accounts"
call :EnsureOU "/Leavers/Guest" "Guest leaver accounts"

REM ==================================================
REM === USERS
REM ==================================================
call :EnsureOU "/Users" "Active user accounts"
call :EnsureOU "/Users/Governors" "Governor user accounts"
call :EnsureOU "/Users/Staff" "Staff user accounts"
call :EnsureOU "/Users/Service" "Service accounts"
call :EnsureOU "/Users/Students" "Student user accounts"
call :EnsureOU "/Users/Guest" "Guest user accounts"

REM ==================================================
REM === CLEANUP
REM ==================================================
del existing_ous.tmp

color 0A
echo.
echo ================================================
echo OU structure complete.
echo ================================================
echo.
echo Closing window in 5 seconds...
timeout /t 5 /nobreak >nul
exit


REM ==================================================
REM === FUNCTION: EnsureOU
REM ==================================================
:EnsureOU
setlocal EnableDelayedExpansion
set "OU=%~1"
set "DESC=%~2"

REM ---- Exact OU existence check
findstr /X /C:"%OU%" existing_ous.tmp >nul
if not errorlevel 1 (
    color 6F
    echo [SKIP] %OU% already exists
    endlocal
    exit /b
)

REM ---- Extract NAME (after last /)
set "NAME=%OU%"
set "NAME=%NAME:*/=%"

REM ---- Extract PARENT (everything before /NAME)
set "PARENT=%OU%"
set "PARENT=%PARENT:/%NAME%=%"

REM ---- Root-level OU fix
if "%PARENT%"=="%OU%" set "PARENT="

color 2F
echo.
echo [CREATE] %OU%

if defined PARENT (
    echo     GAM create org "%NAME%" parent "%PARENT%"
    gam create org "%NAME%" parent "%PARENT%" description "%DESC%"
) else (
    echo     GAM create org "%NAME%"
    gam create org "%NAME%" description "%DESC%"
)

REM ---- Refresh OU cache so children always see parents
gam print orgs > existing_ous.tmp

endlocal
exit /b