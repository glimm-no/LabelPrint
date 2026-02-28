@echo off
setlocal

REM Usage: build.bat [arm64] [x64]
REM   arm64  - build for win-arm64 only
REM   x64    - build for win-x64 only
REM   (none) - build both (default)

set BUILD_ARM64=0
set BUILD_X64=0

:parse_args
if "%~1"=="" goto args_done
if /i "%~1"=="arm64" set BUILD_ARM64=1
if /i "%~1"=="x64"   set BUILD_X64=1
shift
goto parse_args
:args_done

REM Default: build both if no target was specified
if %BUILD_ARM64%==0 if %BUILD_X64%==0 (
    set BUILD_ARM64=1
    set BUILD_X64=1
)

REM --- win-arm64 ---
if %BUILD_ARM64%==1 (
    echo Building win-arm64...

    dotnet restore LabelPrint.csproj -r win-arm64
    if errorlevel 1 exit /b %errorlevel%

    dotnet clean LabelPrint.csproj -c Release -r win-arm64
    if errorlevel 1 exit /b %errorlevel%

    dotnet publish LabelPrint.csproj -c Release -r win-arm64 --self-contained false -p:EnableWindowsTargeting=true -o LabelPrint-arm64-framework
    if errorlevel 1 exit /b %errorlevel%

    REM dotnet publish LabelPrint.csproj -c Release -r win-arm64 --self-contained true -p:PublishTrimmed=true -p:EnableWindowsTargeting=true -o publish-arm64-small
    REM dotnet publish LabelPrint.csproj -c Release -r win-arm64 --self-contained true -p:EnableWindowsTargeting=true -o publish-arm64-small
    REM dotnet publish LabelPrint.csproj -c Release -r win-arm64 --self-contained true -p:PublishSingleFile=true -p:EnableWindowsTargeting=true -o publish
)

REM --- win-x64 ---
if %BUILD_X64%==1 (
    echo Building win-x64...

    dotnet restore LabelPrint.csproj -r win-x64
    if errorlevel 1 exit /b %errorlevel%

    dotnet clean LabelPrint.csproj -c Release -r win-x64
    if errorlevel 1 exit /b %errorlevel%

    dotnet publish LabelPrint.csproj -c Release -r win-x64 --self-contained false -p:EnableWindowsTargeting=true -o LabelPrint-x64-framework
    if errorlevel 1 exit /b %errorlevel%

    REM dotnet publish LabelPrint.csproj -c Release -r win-x64 --self-contained true -p:PublishTrimmed=true -p:EnableWindowsTargeting=true -o publish-x64-small
    REM dotnet publish LabelPrint.csproj -c Release -r win-x64 --self-contained true -p:EnableWindowsTargeting=true -o publish-x64-small
    REM dotnet publish LabelPrint.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:EnableWindowsTargeting=true -o publish-x64
)
