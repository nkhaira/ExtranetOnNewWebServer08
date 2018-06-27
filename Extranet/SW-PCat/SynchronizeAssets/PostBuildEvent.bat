@echo off
copy C:\SynchronizeAssets\\*.config C:\SynchronizeAssets\bin\Debug\
if errorlevel 1 goto CSharpReportError
goto CSharpEnd
:CSharpReportError
echo Project error: A tool returned an error code from the build event
exit 1
:CSharpEnd