@echo off
set VSWHERE="C:\Program Files (x86)\Microsoft Visual Studio\Installer\vswhere.exe"
for /f "tokens=*" %%i in ('%VSWHERE% -latest -products * -requires Microsoft.Component.MSBuild -find MSBuild\**\Bin\MSBuild.exe') do set MSBUILD=%%i

echo Using MSBuild: %MSBUILD%

if not exist "%~dp0x64" mkdir "%~dp0x64"
if not exist "%~dp0x64\Debug" mkdir "%~dp0x64\Debug"

"%MSBUILD%" "%~dp0Audio-Endpoint-Switcher.sln" /p:Configuration=Debug /p:Platform=x64 /p:OutDir="%~dp0x64\Debug\\" /p:IntDir="%~dp0x64\Debug\\" /p:GenerateFullPaths=true