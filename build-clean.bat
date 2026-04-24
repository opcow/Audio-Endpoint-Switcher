@echo off
set VSWHERE="C:\Program Files (x86)\Microsoft Visual Studio\Installer\vswhere.exe"
for /f "tokens=*" %%i in ('%VSWHERE% -latest -products * -requires Microsoft.Component.MSBuild -find MSBuild\**\Bin\MSBuild.exe') do set MSBUILD=%%i

echo Using MSBuild: %MSBUILD%

if exist "%~dp0x64\Debug" rmdir /s /q "%~dp0x64\Debug"
if exist "%~dp0x64\Release" rmdir /s /q "%~dp0x64\Release"
if exist "%~dp0x64\ReleaseNoDLL" rmdir /s /q "%~dp0x64\ReleaseNoDLL"

"%MSBUILD%" "%~dp0Audio-Endpoint-Switcher.sln" /t:Clean /p:Configuration=Debug /p:Platform=x64 /p:GenerateFullPaths=true
"%MSBUILD%" "%~dp0Audio-Endpoint-Switcher.sln" /t:Clean /p:Configuration=Release /p:Platform=x64 /p:GenerateFullPaths=true
"%MSBUILD%" "%~dp0Audio-Endpoint-Switcher.sln" /t:Clean "/p:Configuration=Release No DLL" /p:Platform=x64 /p:GenerateFullPaths=true