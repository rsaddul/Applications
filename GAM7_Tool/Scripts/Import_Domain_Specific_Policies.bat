@echo off
setlocal EnableExtensions EnableDelayedExpansion

rem ====================================================
rem ChromeOS Device Sign-In & Login Screen Configuration
rem ====================================================

set /p DOMAIN=Enter the domain (e.g. thecloudtrust.co.uk): 

if "%DOMAIN%"=="" (
    echo.
    echo ERROR: No domain entered. Exiting.
    exit /b 1
)

echo.
echo Applying ChromeOS device policies...
echo   Allowed sign-in : *@%DOMAIN%
echo   Autocomplete    : %DOMAIN%
echo   OU              : /Devices
echo.

REM ----------------------------------------------------
REM Apply device sign-in restriction (Restricted List)
REM ----------------------------------------------------
echo Applying device sign-in restriction...

gam update chromepolicy chrome.devices.SignInRestriction ^
  deviceAllowNewUsers ALLOW_NEW_USERS_ENUM_RESTRICTED_LIST ^
  userAllowlist *@%DOMAIN% ^
  orgunit /Devices

if errorlevel 1 (
    echo.
    echo ERROR: Failed to apply device sign-in restriction.
    exit /b 1
)

echo Sign-in restriction applied successfully.
echo.

REM ----------------------------------------------------
REM Apply login screen domain autocomplete
REM ----------------------------------------------------
echo Applying login screen domain autocomplete...

gam update chromepolicy chrome.devices.DeviceLoginScreenAutocompleteDomainGroup ^
  loginScreenDomainAutoComplete true ^
  loginScreenDomainAutoCompletePrefix %DOMAIN% ^
  orgunit /Devices

if errorlevel 1 (
    echo.
    echo ERROR: Failed to apply login screen domain autocomplete.
    exit /b 1
)

echo.
echo SUCCESS:
echo - Devices restricted to *@%DOMAIN%
echo - Login screen autocomplete enabled
echo - OU: /Devices
echo.

endlocal

echo.
echo Closing window in 5 seconds...
timeout /t 5 /nobreak >nul
exit

