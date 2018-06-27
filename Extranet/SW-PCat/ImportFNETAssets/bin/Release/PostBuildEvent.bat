@echo off
copy C:\ImportFNETAssets\\*.config C:\ImportFNETAssets\bin\Release\
if errorlevel 1 goto CSharpReportError
goto CSharpEnd
:CSharpReportError
echo Project error: A tool returned an error code from the build event
exit 1
:CSharpEnd