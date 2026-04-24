@echo off
setlocal EnableDelayedExpansion
set VSWHERE="C:\Program Files (x86)\Microsoft Visual Studio\Installer\vswhere.exe"
for /f "tokens=*" %%i in ('%VSWHERE% -latest -products * -requires Microsoft.Component.MSBuild -find MSBuild\**\Bin\MSBuild.exe') do set MSBUILD=%%i

echo Using MSBuild: %MSBUILD%

if not exist "%~dp0x64" mkdir "%~dp0x64"
if not exist "%~dp0x64\Debug" mkdir "%~dp0x64\Debug"
if not exist "%~dp0x64\Release" mkdir "%~dp0x64\Release"
if not exist "%~dp0x64\ReleaseNoDLL" mkdir "%~dp0x64\ReleaseNoDLL"

set BUILDNOFILE=%~dp0build-number.txt

set BUILDNO=0
if exist "%BUILDNOFILE%" (
    set /p BUILDNO=<"%BUILDNOFILE%"
)
set /a BUILDNO+=1
echo !BUILDNO!>"%BUILDNOFILE%"

echo Building version 1.0.15.!BUILDNO!...

powershell -ExecutionPolicy Bypass -File "%~dp0update-build-num.ps1" -BuildNo "!BUILDNO!" -RcFile "%~dp0Audio-Endpoint-Switcher\aes.rc"

"%MSBUILD%" "%~dp0Audio-Endpoint-Switcher.sln" /p:Configuration=Debug /p:Platform=x64 /p:OutDir="%~dp0x64\Debug\\" /p:IntDir="%~dp0x64\Debug\\" /p:GenerateFullPaths=true
"%MSBUILD%" "%~dp0Audio-Endpoint-Switcher.sln" /p:Configuration=Release /p:Platform=x64 /p:OutDir="%~dp0x64\Release\\" /p:IntDir="%~dp0x64\Release\\" /p:GenerateFullPaths=true
"%MSBUILD%" "%~dp0Audio-Endpoint-Switcher.sln" "/p:Configuration=Release No DLL" /p:Platform=x64 /p:OutDir="%~dp0x64\ReleaseNoDLL\\" /p:IntDir="%~dp0x64\ReleaseNoDLL\\" /p:GenerateFullPaths=true