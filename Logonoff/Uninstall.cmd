@echo off
set "appname=Logonoff"

sc stop %appname%
sc delete %appname%