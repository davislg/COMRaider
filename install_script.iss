;InnoSetupVersion=4.2.6

[Setup]
AppName=COMRaider
AppVerName=COMRaider v0.132
DefaultDirName=c:\iDefense\COMRaider\
DefaultGroupName=ComRaider
OutputBaseFilename=COMRaider_Setup
OutputDir=./

[Files]
Source: "./\crashmon.dll"; DestDir: "{app}"; Flags: ignoreversion 
Source: "./\crashmon\crashmon.dsw"; DestDir: "{app}\crashmon"; 
Source: "./\crashmon\crashmon.def"; DestDir: "{app}\crashmon"; 
Source: "./\crashmon\crashmon.cpp"; DestDir: "{app}\crashmon"; 
Source: "./\crashmon\pebModules.cpp"; DestDir: "{app}\crashmon"; 
Source: "./\crashmon\stackwalk.cpp"; DestDir: "{app}\crashmon"; 
Source: "./\crashmon\crashmon.dsp"; DestDir: "{app}\crashmon"; 
Source: "./\crashmon\safeCoCreate.cpp"; DestDir: "{app}\crashmon"; 
Source: "./\crashmon\symload.cpp"; DestDir: "{app}\crashmon"; 
Source: "./\crashmon\TLBINF32.h"; DestDir: "{app}\crashmon"; 
Source: "./\vuln_test\vuln.opt"; DestDir: "{app}\vuln_test"; 
Source: "./\vuln_test\dlldata.c"; DestDir: "{app}\vuln_test"; 
Source: "./\vuln_test\dlldatax.c"; DestDir: "{app}\vuln_test"; 
Source: "./\vuln_test\dlldatax.h"; DestDir: "{app}\vuln_test"; 
Source: "./\vuln_test\Resource.h"; DestDir: "{app}\vuln_test"; 
Source: "./\vuln_test\server.cpp"; DestDir: "{app}\vuln_test"; 
Source: "./\vuln_test\server.h"; DestDir: "{app}\vuln_test"; 
Source: "./\vuln_test\server.rgs"; DestDir: "{app}\vuln_test"; 
Source: "./\vuln_test\StdAfx.cpp"; DestDir: "{app}\vuln_test"; 
Source: "./\vuln_test\StdAfx.h"; DestDir: "{app}\vuln_test"; 
Source: "./\vuln_test\vuln.aps"; DestDir: "{app}\vuln_test"; 
Source: "./\vuln_test\vuln.cpp"; DestDir: "{app}\vuln_test"; 
Source: "./\vuln_test\vuln.def"; DestDir: "{app}\vuln_test"; 
Source: "./\vuln_test\vuln.dsp"; DestDir: "{app}\vuln_test"; 
Source: "./\vuln_test\vuln.dsw"; DestDir: "{app}\vuln_test"; 
Source: "./\vuln_test\vuln.h"; DestDir: "{app}\vuln_test"; 
Source: "./\vuln_test\vuln.idl"; DestDir: "{app}\vuln_test"; 
Source: "./\vuln_test\vuln.ncb"; DestDir: "{app}\vuln_test"; 
Source: "./\vuln_test\vuln.plg"; DestDir: "{app}\vuln_test"; 
Source: "./\vuln_test\vuln.rc"; DestDir: "{app}\vuln_test"; 
Source: "./\vuln_test\vuln.tlb"; DestDir: "{app}\vuln_test"; 
Source: "./\vuln_test\vuln_i.c"; DestDir: "{app}\vuln_test"; 
Source: "./\vuln_test\vuln_p.c"; DestDir: "{app}\vuln_test"; 
Source: "./\vuln_test\vulnps.def"; DestDir: "{app}\vuln_test"; 
Source: "./\vuln_test\vulnps.mk"; DestDir: "{app}\vuln_test"; 
Source: "./\COMRaider.exe"; DestDir: "{app}"; Flags: ignoreversion 
Source: "./\BuildArgs.vbs"; DestDir: "{app}"; 
Source: "./\olly.dll"; DestDir: "{app}"; Flags: ignoreversion 
Source: "./\logger.dll"; DestDir: "{app}"; Flags: ignoreversion 
Source: "./\comraider.mdb"; DestDir: "{app}"; 
Source: "./\COMRaider.chm"; DestDir: "{app}"; 
Source: "./\vuln.dll"; DestDir: "{app}"; 
Source: "./dependancy\mscomctl.ocx"; DestDir: "{sys}"; Flags: uninsneveruninstall regserver 
Source: "./dependancy\msscript.ocx"; DestDir: "{sys}"; Flags: uninsneveruninstall regserver 
Source: "./dependancy\spSubclass.dll"; DestDir: "{win}"; Flags: regserver sharedfile 
Source: "./dependancy\TLBINF32.DLL"; DestDir: "{sys}"; Flags: regserver sharedfile 
Source: "./dependancy\vbDevKit.dll"; DestDir: "{win}"; Flags: regserver sharedfile 
Source: "./dependancy\MSWINSCK.OCX"; DestDir: "{sys}"; Flags: uninsneveruninstall regserver 
Source: "./\olly_dll\asmserv.c"; DestDir: "{app}\olly_dll"; 
Source: "./\olly_dll\assembl.c"; DestDir: "{app}\olly_dll"; 
Source: "./\olly_dll\disasm.c"; DestDir: "{app}\olly_dll"; 
Source: "./\olly_dll\disasm.h"; DestDir: "{app}\olly_dll"; 
Source: "./\olly_dll\gpl.wri"; DestDir: "{app}\olly_dll"; 
Source: "./\olly_dll\olly.def"; DestDir: "{app}\olly_dll"; 
Source: "./\olly_dll\olly.dsp"; DestDir: "{app}\olly_dll"; 
Source: "./\olly_dll\olly.dsw"; DestDir: "{app}\olly_dll"; 
Source: "./\olly_dll\readme.htm"; DestDir: "{app}\olly_dll"; 
Source: "./\olly_dll\Readme.txt"; DestDir: "{app}\olly_dll"; 
Source: "./\ComRaider\ComRaider.vbp"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\options.dat"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\CCrashMon.cls"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\frmCrashMon.frm"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\CMember.cls"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\CException.cls"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\CLogger.cls"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\frmMsg.frx"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\frmMsg.frm"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\modMain.bas"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\frmTlbViewer.frm"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\MDIForm1.frm"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\CArgument.cls"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\frmSafe4Scripting.frm"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\CTlbParse.cls"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\modCrashMon.bas"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\frmLoadFile.frm"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\MDIForm1.frx"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\Project1.vbw"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\CModule.cls"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\ComRaider.vbw"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\frmTlbViewer.frx"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\CClass.cls"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\CInterface.cls"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\frmScanDir.frm"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\FileVer.bas"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\frmOptions.frm"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\CAuditEntry.cls"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\frmOptions.frx"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\frmLoadFile.frx"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\CWindow.cls"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\modWinMonitor.bas"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\frmEditAuditNotes.frm"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\frmDownloadShared.frm"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\frmAuditLogs.frm"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\frmDistNotes.frm"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\frmAddDistNote.frm"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\modDllInject.bas"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\modPEParse.bas"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\frmAdvScan.frm"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\frmAbout.frm"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\frmAbout.frx"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\frmKillBits.frm"; DestDir: "{app}\ComRaider"; 
Source: "./\ComRaider\frmProgIDLookup.frm"; DestDir: "{app}\ComRaider"; 
Source: "./\buildDB\Project1.vbw"; DestDir: "{app}\buildDB"; 
Source: "./\buildDB\modIESafe.bas"; DestDir: "{app}\buildDB"; 
Source: "./\buildDB\Form1.frx"; DestDir: "{app}\buildDB"; 
Source: "./\buildDB\CClassSafety.cls"; DestDir: "{app}\buildDB"; 
Source: "./\buildDB\Form1.frm"; DestDir: "{app}\buildDB"; 
Source: "./\buildDB\Project1.vbp"; DestDir: "{app}\buildDB"; 
Source: "./\buildDB.exe"; DestDir: "{app}"; 
Source: "./\logger_dll\parse_h\parse_h.exe"; DestDir: "{app}\logger_dll\parse_h"; 
Source: "./\logger_dll\parse_h\Project1.vbw"; DestDir: "{app}\logger_dll\parse_h"; 
Source: "./\logger_dll\parse_h\Form1.frm"; DestDir: "{app}\logger_dll\parse_h"; 
Source: "./\logger_dll\parse_h\Project1.vbp"; DestDir: "{app}\logger_dll\parse_h"; 
Source: "./\logger_dll\dll.dsw"; DestDir: "{app}\logger_dll"; 
Source: "./\logger_dll\hooker.h"; DestDir: "{app}\logger_dll"; 
Source: "./\logger_dll\hooker.lib"; DestDir: "{app}\logger_dll"; 
Source: "./\logger_dll\main.cpp"; DestDir: "{app}\logger_dll"; 
Source: "./\logger_dll\ReadMe.txt"; DestDir: "{app}\logger_dll"; 
Source: "./\logger_dll\dll.dsp"; DestDir: "{app}\logger_dll"; 
Source: "./\logger_dll\main.h"; DestDir: "{app}\logger_dll"; 
Source: "./\logger_dll\dll.cpp"; DestDir: "{app}\logger_dll"; 

[Dirs]
Name: "{app}\crashmon"; 
Name: "{app}\vuln_test"; 
Name: "{app}\vuln_test\Debug"; 
Name: "{app}\olly_dll"; 
Name: "{app}\ComRaider"; 
Name: "{app}\buildDB"; 
Name: "{app}\logger_dll"; 
Name: "{app}\logger_dll\parse_h"; 
Name: "{app}\logger_dll\injector"; 

[Icons]
Name: "{group}\ComRaider"; Filename: "{app}\COMRaider.exe"; 
Name: "{group}\COMRaider.chm"; Filename: "{app}\COMRaider.chm"; 
Name: "{group}\Source\ComRaider.vbp"; Filename: "{app}\ComRaider\ComRaider.vbp"; 
Name: "{group}\Source\CrashMon.dsw"; Filename: "{app}\crashmon\crashmon.dsw"; 
Name: "{group}\Source\Olly_dll.dsw"; Filename: "{app}\olly_dll\olly.dsw"; 
Name: "{group}\Source\VulnTest_dll.dsw"; Filename: "{app}\vuln_test\Vuln.dsw"; 
Name: "{group}\Uninstall"; Filename: "{app}\unins000.exe"; 
Name: "{userdesktop}\COMRaider"; Filename: "{app}\COMRaider.exe"; WorkingDir: "{app}"; 
Name: "{group}\Source\BuildDB.exe"; Filename: "{app}\buildDB\Project1.vbp"; 
Name: "{group}\Source\logger.dsw"; Filename: "{app}\logger_dll\dll.dsw"; 

[CustomMessages]
NameAndVersion=%1 version %2
AdditionalIcons=Additional icons:
CreateDesktopIcon=Create a &desktop icon
CreateQuickLaunchIcon=Create a &Quick Launch icon
ProgramOnTheWeb=%1 on the Web
UninstallProgram=Uninstall %1
LaunchProgram=Launch %1
AssocFileExtension=&Associate %1 with the %2 file extension
AssocingFileExtension=Associating %1 with the %2 file extension...
