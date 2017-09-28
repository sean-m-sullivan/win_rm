@echo off
REM Set Active Server Pages (ASP) configurations
%windir%\system32\inetsrv\appcmd set config -section:system.webServer/asp /limits.scriptTimeout:"00:00:30" /commit:apphost
%windir%\system32\inetsrv\appcmd set config -section:system.webServer/asp /session.allowSessionState:"false" /commit:apphost