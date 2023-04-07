@echo off

title Office Tool Plus - Setup

:: Get admin permissions.
cd /d "%~dp0" && ( if exist "%temp%\getadmin.vbs" del "%temp%\getadmin.vbs" ) && fsutil dirty query %systemdrive% 1>nul 2>nul || (  cmd /u /c echo Set UAC = CreateObject^("Shell.Application"^) : UAC.ShellExecute "cmd.exe", "/k cd ""%~sdp0"" && ""%~s0"" %Apply%", "", "runas", 1 >> "%temp%\getadmin.vbs" && "%temp%\getadmin.vbs" && exit /B )

:: Set the runtime path.
SET DOTNET_ROOT(x86)=%~dp0Runtime
SET DOTNET_ROOT=%~dp0Runtime
:: If Office Tool Plus still does not run, please download .NET 5 Desktop Runtime x86 from https://dotnet.microsoft.com/download/dotnet/current/runtime

:: Usage for silence install:
::	"Office Tool Plus" [/LoadConfig file-path] [/Edition value] [/SourcePath value]
:: Options:
::	/LoadConfig  file-path	Load configuration file.
::	/Edition value		Override the default OfficeClientEdition, value: 32 or 64.
::	/SourcePath value		Override the default SourcePath.
:: If edition or source path is not set, the value in the configuration will be used.
:: Put arguments after "Office Tool Plus.exe", such as "Office Tool Plus.exe" /loadConfig "%~dp0ConfigForISO.xml" /SourcePath "%~dp0"

start "" "Office Tool Plus.exe"
timeout 3
exit