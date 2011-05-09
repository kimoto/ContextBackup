@echo off
cd /d %~dp0
regsvr32.exe /u "..\x64\Release\ContextBackup.dll"
