@echo off

rem This batch file (re)installs the CoolTrayService, then starts it.

rem Change this path:
set SERVICEPATH="C:\V9\trunk\live Update\demos\CoolService"

%SERVICEPATH%\CoolService.exe /uninstall /silent
%SERVICEPATH%\CoolService.exe /install /silent
net start CoolTrayService

rem Start the Service Manager
rem %SystemRoot%\system32\services.msc /s

