@echo off
:: Limit the download speed, example: 1M, 500K "empty means unlimited"
set speedLimit=

set "_work=%~dp0"
set "_work=%_work:~0,-1%"
set "_batn=%~nx0"
setlocal EnableDelayedExpansion
pushd "!_work!"
set exist=0
if exist "curl.exe" set exist=1
for %%i in (curl.exe) do @if not "%%~$PATH:i"=="" set exist=1
if %exist%==0 echo.&echo Error: curl.exe is not detected&echo.&popd&pause&exit /b
set "uri=temp_curl.txt"
if defined speedLimit set "speedLimit=--limit-rate %speedLimit%"
echo Downloading...
echo.
if exist "%uri%" del /f /q "%uri%"
call :GenTXT TXinfo > "%uri%"
curl.exe -q --create-dirs --retry 5 --retry-connrefused %speedLimit% -k -L -C - -K "%uri%"
if exist "%uri%" del /f /q "%uri%"
echo.
echo Done.
echo Press any key to exit.
popd
pause >nul
exit /b

:GenTXT
set [=&for /f "delims=:" %%s in ('findstr /nbrc:":%~1:\[" /c:":%~1:\]" "!_batn!"') do if defined [ (set /a ]=%%s-3) else set /a [=%%s-1
<"!_batn!" ((for /l %%i in (0 1 %[%) do set /p =)&for /l %%i in (%[% 1 %]%) do (set txt=&set /p txt=&echo(!txt!)) &exit/b

:TXinfo:[
url https://officecdn.microsoft.com/db/64256afe-f5d9-4f86-8936-8840a6a4f5be/Office/Data/v64_16.0.17029.20068.cab
-o C2R_MonthlyPreview\Office\Data\v64.cab
url https://officecdn.microsoft.com/db/64256afe-f5d9-4f86-8936-8840a6a4f5be/Office/Data/v64_16.0.17029.20068.cab
-o C2R_MonthlyPreview\Office\Data\v64_16.0.17029.20068.cab
url https://officecdn.microsoft.com/db/64256afe-f5d9-4f86-8936-8840a6a4f5be/Office/Data/16.0.17029.20068/i642052.cab
-o C2R_MonthlyPreview\Office\Data\16.0.17029.20068\i642052.cab
url https://officecdn.microsoft.com/db/64256afe-f5d9-4f86-8936-8840a6a4f5be/Office/Data/16.0.17029.20068/s642052.cab
-o C2R_MonthlyPreview\Office\Data\16.0.17029.20068\s642052.cab
url https://officecdn.microsoft.com/db/64256afe-f5d9-4f86-8936-8840a6a4f5be/Office/Data/16.0.17029.20068/i640.cab
-o C2R_MonthlyPreview\Office\Data\16.0.17029.20068\i640.cab
url https://officecdn.microsoft.com/db/64256afe-f5d9-4f86-8936-8840a6a4f5be/Office/Data/16.0.17029.20068/s640.cab
-o C2R_MonthlyPreview\Office\Data\16.0.17029.20068\s640.cab
url https://officecdn.microsoft.com/db/64256afe-f5d9-4f86-8936-8840a6a4f5be/Office/Data/16.0.17029.20068/stream.x64.zh-CN.dat
-o C2R_MonthlyPreview\Office\Data\16.0.17029.20068\stream.x64.zh-CN.dat
url https://officecdn.microsoft.com/db/64256afe-f5d9-4f86-8936-8840a6a4f5be/Office/Data/16.0.17029.20068/stream.x64.x-none.dat
-o C2R_MonthlyPreview\Office\Data\16.0.17029.20068\stream.x64.x-none.dat
:TXinfo:]
exit /b
