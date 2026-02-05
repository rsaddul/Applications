@echo off
REM Developed by: Rhys Saddul
color 1F

setlocal EnableExtensions

set "OUTFILE=%~dp0audit_export_all_policies.csv"

REM --- Start fresh
> "%OUTFILE%" echo orgUnitPath,policyName,setting,value

echo Exporting Chrome policies across all OUs...

for /f "skip=1 tokens=1 delims=," %%o in ('gam print ous') do (
    echo Processing OU: %%o
    gam print chromepolicies orgunit "%%o" >> "%OUTFILE%"
)

color 0A
echo.
echo Export complete...
echo %OUTFILE%

echo.
echo Closing window in 5 seconds...
timeout /t 5 /nobreak >nul
exit