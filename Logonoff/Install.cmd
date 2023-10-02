@echo off
set "appname=Logonoff"

echo %~dp0

sc stop %appname%

sc create %appname% start= auto binPath= "%~dp0bin\Release\Logonoff.exe" displayname= "Service %appname%" type= interact type= own
sc description %appname% "Micro-service : %appname%"
sc failure %appname% reset= 86400 actions= restart/1000/restart/1000/restart/1000

sc start %appname%