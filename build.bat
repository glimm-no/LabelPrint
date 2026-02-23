@echo off
setlocal

dotnet restore LabelPrint.csproj -r win-arm64
if errorlevel 1 exit /b %errorlevel%

dotnet clean LabelPrint.csproj -c Release -r win-arm64
if errorlevel 1 exit /b %errorlevel%

dotnet publish LabelPrint.csproj -c Release -r win-arm64 --self-contained false -p:EnableWindowsTargeting=true -o LabelPrint-arm64-framework
if errorlevel 1 exit /b %errorlevel%

REM dotnet publish LabelPrint.csproj -c Release -r win-arm64 --self-contained true -p:PublishTrimmed=true -p:EnableWindowsTargeting=true -o publish-arm64-small

REM dotnet publish LabelPrint.csproj -c Release -r win-arm64 --self-contained true -p:EnableWindowsTargeting=true -o publish-arm64-small

REM dotnet publish LabelPrint.csproj -c Release -r win-arm64 --self-contained true -p:PublishSingleFile=true -p:EnableWindowsTargeting=true -o publish
