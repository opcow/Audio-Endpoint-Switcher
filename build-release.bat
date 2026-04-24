@echo off
set VSWHERE="C:\Program Files (x86)\Microsoft Visual Studio\Installer\vswhere.exe"
for /f "tokens=*" %%i in ('%VSWHERE% -latest -products * -requires Microsoft.Component.MSBuild -find MSBuild\**\Bin\MSBuild.exe') do set MSBUILD=%%i

echo Using MSBuild: %MSBUILD%

if not exist "%~dp0x64" mkdir "%~dp0x64"
if not exist "%~dp0x64\Release" mkdir "%~dp0x64\Release"

"%MSBUILD%" "%~dp0Audio-Endpoint-Switcher.sln" /p:Configuration=Release /p:Platform=x64 /p:OutDir="%~dp0x64\Release\\" /p:IntDir="%~dp0x64\Release\\" /p:GenerateFullPaths=true