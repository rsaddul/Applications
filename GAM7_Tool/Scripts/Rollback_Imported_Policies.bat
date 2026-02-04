gam csv "%~dp0Import_All_Policies.csv" ^
  gam delete chromepolicy "~policyName" orgunit "~orgUnitPath"

echo.
echo Closing window in 5 seconds...
timeout /t 5 /nobreak >nul
exit