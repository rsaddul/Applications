@echo off
REM Developed by: Rhys Saddul

color 1F
echo Starting rollback of Standard Solution Design Policies...

color 2F
echo Rolling back policies...
gam csv "%~dp0Import_All_Policies.csv" ^
  gam delete chromepolicy "~policyName" orgunit "~orgUnitPath"

color 0A
echo Rollback completed successfully...
echo.
echo Closing window in 5 seconds...
timeout /t 5 /nobreak >nul
exit