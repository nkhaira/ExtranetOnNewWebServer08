@echo off
copy C:\productengine_test\\*.config C:\productengine_test\bin\Debug\
if errorlevel 1 goto CSharpReportError
goto CSharpEnd
:CSharpReportError
echo Project error: A tool returned an error code from the build event
exit 1
:CSharpEnd