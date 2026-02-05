@echo off
REM Developed by: Rhys Saddul

color 1F
echo Starting Standard Solution Design Policy import...

color 2F
echo Importing policies...
gam csv "%~dp0Import_All_Policies.csv" ^
  gam update chromepolicy "~policyName" "~fieldName" "~fieldValue" orgunit "~orgUnitPath"

color 0A
echo Import Complete...
echo.
echo Closing window in 5 seconds...
timeout /t 5 /nobreak >nul
exit