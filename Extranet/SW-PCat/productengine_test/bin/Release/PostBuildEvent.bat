@echo off
copy E:\Projects\Fluke\ReDesign\FNET_WWW\ProductEngine_TEST\\*.config E:\Projects\Fluke\ReDesign\FNET_WWW\ProductEngine_TEST\bin\Release\
if errorlevel 1 goto CSharpReportError
goto CSharpEnd
:CSharpReportError
echo Project error: A tool returned an error code from the build event
exit 1
:CSharpEnd