[Setup]
AppID=DVD Flick
AppName=DVD Flick
AppVerName=DVDFlick 1.3.0.7
AppVersion=1.3.0.7
AppPublisher=Dennis Meuwissen
AppCopyright=Copyright © 2006-2009, Dennis Meuwissen
AppPublisherURL=http://www.dvdflick.net
AppSupportURL=http://www.dvdflick.net
AppUpdatesURL=http://www.dvdflick.net
DefaultDirName={pf}\DVD Flick
DefaultGroupName=DVD Flick
PrivilegesRequired=admin
MinVersion=0,5.0.2195
Compression=lzma/ultra64
InternalCompressLevel=ultra64
SolidCompression=true
InfoAfterFile=readme.txt
LicenseFile=license.txt
AlwaysUsePersonalGroup=false
AppendDefaultGroupName=true
AllowNoIcons=true
ChangesAssociations=true
FlatComponentsList=true
UninstallLogMode=overwrite
ShowLanguageDialog=no
LanguageDetectionMethod=none
UninstallDisplayIcon={app}\dvdflick.exe
UninstallDisplayName=DVD Flick 1.3.0.7
VersionInfoVersion=1.3.0.7
OutputDir=.
SourceDir=.
OutputBaseFilename=dvdflick_setup_1.3.0.7
WizardImageStretch=false
WizardImageFile=resources\setupimage.bmp
WizardSmallImageFile=resources\setupimage_small.bmp
AppMutex=DVD Flick


[Languages]
Name: en; MessagesFile: compiler:Default.isl


[CustomMessages]
en.SetupIsRunningWarningInstall=%1's setup is already running!
en.SetupIsRunningWarningUninstall=%1's setup is already running!
en.DeleteSettings=Do you also want to delete %1's settings? If you plan on reinstalling %1, you don't have to delete them.


[Tasks]
Name: desktopicon; Description: Create a desktop icon
Name: associate; Description: Associate with DVD Flick Project files


[Files]
Source: shared\mscomctl.ocx; DestDir: {sys}; Flags: restartreplace sharedfile regserver uninsnosharedfileprompt
Source: shared\richtx32.ocx; DestDir: {sys}; Flags: restartreplace sharedfile regserver uninsnosharedfileprompt
Source: shared\mscomct2.ocx; DestDir: {sys}; Flags: restartreplace sharedfile regserver uninsnosharedfileprompt
Source: shared\comctl32.ocx; DestDir: {sys}; Flags: restartreplace sharedfile regserver uninsnosharedfileprompt
Source: shared\comct232.ocx; DestDir: {sys}; Flags: restartreplace sharedfile regserver uninsnosharedfileprompt
Source: shared\mousewheel.ocx; DestDir: {sys}; Flags: restartreplace sharedfile regserver uninsnosharedfileprompt

Source: shared\stdole2.tlb; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regtypelib uninsnosharedfileprompt
Source: shared\msvbvm60.dll; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver uninsnosharedfileprompt
Source: shared\oleaut32.dll; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver uninsnosharedfileprompt
Source: shared\olepro32.dll; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver uninsnosharedfileprompt
Source: shared\asycfilt.dll; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile uninsnosharedfileprompt
Source: shared\comcat.dll; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver uninsnosharedfileprompt

Source: shared\scrrun.dll; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver uninsnosharedfileprompt
Source: shared\trayicon_handler.ocx; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver uninsnosharedfileprompt
Source: shared\ssubtmr6.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver uninsnosharedfileprompt

Source: readme.txt; DestDir: {app}; Flags: ignoreversion
Source: changelog.txt; DestDir: {app}; Flags: ignoreversion
Source: license.txt; DestDir: {app}; Flags: ignoreversion
Source: dvdflick.exe; DestDir: {app}; Flags: ignoreversion
Source: dvdflick.dll; DestDir: {app}; Flags: ignoreversion
Source: rspcpu.dll; DestDir: {app}; Flags: ignoreversion

Source: ..\guide\*.*; DestDir: {app}\guide; Flags: recursesubdirs createallsubdirs ignoreversion
Source: data\*.*; DestDir: {app}\data; Flags: recursesubdirs createallsubdirs ignoreversion
Source: bin\*.*; DestDir: {app}\bin; Flags: recursesubdirs createallsubdirs ignoreversion
Source: delaycut\*.*; DestDir: {app}\delaycut; Flags: recursesubdirs createallsubdirs ignoreversion
Source: mkvextract\*.*; DestDir: {app}\mkvextract; Flags: recursesubdirs createallsubdirs ignoreversion
Source: templates\*.*; DestDir: {app}\templates; Flags: recursesubdirs createallsubdirs ignoreversion

Source: imgburn\imgburn.exe; DestDir: {app}\imgburn; Flags: ignoreversion
Source: imgburn\imgburnpreview.exe; DestDir: {app}\imgburn; Flags: ignoreversion
Source: imgburn\imgburn_bare.ini; DestDir: {app}\imgburn; DestName: imgburn.ini; Flags: ignoreversion

Source: resources\main icons\Document.ico; DestDir: {app}; DestName: document.ico; Flags: ignoreversion


[Dirs]
Name: {app}\imgburn


[Registry]
Root: HKCR; SubKey: .dfproj; ValueType: string; ValueData: DVDFlick; Flags: uninsdeletekey; Tasks: associate
Root: HKCR; SubKey: DVDFlick; ValueType: string; ValueData: DVD Flick Project; Flags: uninsdeletekey; Tasks: associate
Root: HKCR; SubKey: DVDFlick\Shell\Open\Command; ValueType: string; ValueData: """{app}\dvdflick.exe"" -load ""%1"""; Flags: uninsdeletevalue; Tasks: associate
Root: HKCR; Subkey: DVDFlick\DefaultIcon; ValueType: string; ValueData: {app}\document.ico,0; Flags: uninsdeletevalue; Tasks: associate


[Icons]
Name: {userdesktop}\DVD Flick; Filename: {app}\dvdflick.exe; WorkingDir: {app}; IconFilename: {app}\dvdflick.exe; IconIndex: 0; Tasks: " desktopicon"; Comment: DVD Flick
Name: {group}\DVD Flick; Filename: {app}\dvdflick.exe; WorkingDir: {app}; IconFilename: {app}\dvdflick.exe; IconIndex: 0; Comment: DVD Flick
Name: {group}\Help and Support\Guide; Filename: {app}\guide\index_en.html; Comment: Guide; WorkingDir: {app}\guide
Name: {group}\Help and Support\Readme; Filename: {app}\readme.txt; Comment: Readme; WorkingDir: {app}
Name: {group}\Help and Support\Changelog; Filename: {app}\changelog.txt; Comment: Changelog; WorkingDir: {app}
Name: {group}\Help and Support\GNU GPL License; Filename: {app}\license.txt; Comment: GNU GPL License; WorkingDir: {app}
Name: {group}\Help and Support\{cm:ProgramOnTheWeb,DVD Flick}; Filename: http://www.dvdflick.net; Comment: {cm:ProgramOnTheWeb,DVD Flick}; WorkingDir: {app}
Name: {group}\{cm:UninstallProgram, DVD Flick}; Filename: {uninstallexe}; Comment: {cm:UninstallProgram, DVD Flick}; WorkingDir: {app}


[InstallDelete]
; Delete on installation the entries of start menu so when you upgrade, to prevent having the same shortcut twice.
Type: files; Name: {group}\Guide.lnk
Type: files; Name: {group}\Readme.lnk
Type: files; Name: {group}\Changelog.lnk
Type: files; Name: {group}\GNU GPL License.lnk


[Run]
Filename: {app}\dvdflick.exe; WorkingDir: {app}; Description: Run DVD Flick; Flags: nowait postinstall unchecked hidewizard skipifsilent runascurrentuser


[UninstallDelete]
Name: {userappdata}\DVD Flick; Type: dirifempty
Name: {app}\imgburn; Type: dirifempty
Name: {app}\guide; Type: dirifempty
Name: {app}\data; Type: dirifempty
Name: {app}\bin; Type: dirifempty
Name: {app}\delaycut; Type: dirifempty
Name: {app}\mkvextract; Type: dirifempty
Name: {app}\templates; Type: dirifempty
Name: {app}; Type: dirifempty


[Code]
// Create a constant for the installer
const installer_mutex_name = 'dvdflick_setup_mutex';

// General functions
function IsInstalled( AppID: String ): Boolean;
var
	sPrevPath: String;
begin
	sPrevPath := '';
	if not RegQueryStringValue( HKLM, 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\'+AppID+'_is1', 'Inno Setup: App Path', sPrevpath) then
		RegQueryStringValue( HKCU, 'Software\Microsoft\Windows\CurrentVersion\Uninstall\'+AppID+'_is1', 'Inno Setup: App Path', sPrevpath);

  Result := sPrevPath<>'';
end;

// If this is an update then we use the same directories again
function ShouldSkipPage(PageID: Integer): Boolean;
begin
	Result := False;
	if (PageID = wpSelectDir) or (PageID = wpSelectProgramGroup) then begin
		Result := IsInstalled('{DVD Flick}');
	end;
end;

// When uninstalling the program, ask user to delete settings based on whether the file exists only
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
	if CurUninstallStep = usUninstall then begin
		if fileExists(ExpandConstant('{userappdata}\DVD Flick\dvdflick.cfg')) OR fileExists(ExpandConstant('{userappdata}\DVD Flick\dvdflick.log'))
		OR fileExists(ExpandConstant('{userappdata}\DVD Flick\report.txt')) OR fileExists(ExpandConstant('{userappdata}\DVD Flick\tetris.dat')) then begin
			if MsgBox(ExpandConstant('{cm:DeleteSettings,DVD Flick}'), mbConfirmation, MB_YESNO or MB_DEFBUTTON2) = IDYES then begin
				DeleteFile(ExpandConstant('{userappdata}\DVD Flick\dvdflick.cfg'));
				DeleteFile(ExpandConstant('{userappdata}\DVD Flick\dvdflick.log'));
				DeleteFile(ExpandConstant('{userappdata}\DVD Flick\report.txt'));
				DeleteFile(ExpandConstant('{userappdata}\DVD Flick\tetris.dat'));
			end;
		end;
	end;
end;

function InitializeSetup(): Boolean;
begin
	// Create a mutex for the installer and if it's already running then expose a message and stop installation
	Result := True;
	if CheckForMutexes(installer_mutex_name) then begin
		if not WizardSilent() then
			MsgBox(ExpandConstant('{cm:SetupIsRunningWarningInstall,DVD Flick}'), mbError, MB_OK);
			Result := False;
		end
		else begin
		CreateMutex(installer_mutex_name);
	end;
end;

function InitializeUninstall(): Boolean;
begin
	Result := True;
	if CheckForMutexes(installer_mutex_name) then begin
		if not WizardSilent() then
			MsgBox(ExpandConstant('{cm:SetupIsRunningWarningUninstall,DVD Flick}'), mbError, MB_OK);
		Result := False;
		end
		else begin
		CreateMutex(installer_mutex_name);
	end;
end;
