@echo off
copy C:\DTM_Web\AssetIndexing1\AssetIndexing\\*.config C:\DTM_Web\AssetIndexing1\AssetIndexing\bin\Debug\
if errorlevel 1 goto CSharpReportError
goto CSharpEnd
:CSharpReportError
echo Project error: A tool returned an error code from the build event
exit 1
:CSharpEnd