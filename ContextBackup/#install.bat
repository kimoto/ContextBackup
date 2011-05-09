@echo off
cd /d %~dp0
regsvr32.exe "..\x64\Release\ContextBackup.dll"
