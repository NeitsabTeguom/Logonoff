@echo off
set "appname=Logonoff"

sc stop %appname%

sc create %appname% start= auto binPath= "C:\PERSO\DEV\Logonoff\Logonoff\Logonoff\bin\Release\Logonoff.exe" displayname= "Service %appname%"
sc description %appname% "Micro-service : %appname%"
sc failure %appname% reset= 86400 actions= restart/1000/restart/1000/restart/1000

sc start %appname%