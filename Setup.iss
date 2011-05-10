[Setup]
AppName=ContextBackup
AppVerName=ContextBackup 1.0
OutputDir=.\
OutputBaseFilename=Setup
DefaultDirName={pf}\ContextBackup
SourceDir=C:\Users\kimoto\Documents\Visual Studio 2010\Projects\ContextBackup

[Files]
Source: readme.txt; DestDir: {app}
Source: x64\Release\ContextBackup.dll; DestDir: {app}\x64\Release
Source: ContextBackup\#install.bat; DestDir: {app}\ContextBackup;
Source: ContextBackup\#uninstall.bat; DestDir: {app}\ContextBackup
 
[Run]
Filename: "{app}\ContextBackup\#install.bat";

[UninstallRun]
Filename: "{app}\ContextBackup\#uninstall.bat";

[Languages]
Name: japanese; MessagesFile: compiler:Languages\Japanese.isl 
