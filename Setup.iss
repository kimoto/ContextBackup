[Setup]
AppName=ContextBackup
AppVerName=ContextBackup 1.0
OutputDir=.\
OutputBaseFilename=Setup(x64)
DefaultDirName={pf}\ContextBackup
SourceDir=.\
VersionInfoVersion=0.0.0.0
VersionInfoDescription=description
AppCopyright=kimoto
AppPublisher=kimoto
AppPublisherURL=
AppVersion=1.0
AppContact=peerler@gmail.com
AppSupportURL=
AppReadmeFile={app}\readme.txt
AppUpdatesURL=
AppComments=ContextBackup

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
