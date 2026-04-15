[Setup]
AppId={{A2F9E0FA-39C9-47F4-8F5F-7D49D5A8A501}
AppName=DetailExtract OCR
AppVersion=1.1.0
DefaultDirName={autopf}\DetailExtract OCR
DefaultGroupName=DetailExtract OCR
OutputDir=dist
OutputBaseFilename=DetailExtractOCR_Installer
Compression=lzma
SolidCompression=yes
WizardStyle=modern
SetupIconFile=assets\detailextract.ico

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a desktop icon"; GroupDescription: "Additional icons:"

[Files]
Source: "dist\DetailExtractOCRApp\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\DetailExtract OCR"; Filename: "{app}\DetailExtractOCRApp.exe"; IconFilename: "{app}\DetailExtractOCRApp.exe"
Name: "{commondesktop}\DetailExtract OCR"; Filename: "{app}\DetailExtractOCRApp.exe"; IconFilename: "{app}\DetailExtractOCRApp.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\DetailExtractOCRApp.exe"; Description: "Launch DetailExtract OCR"; Flags: nowait postinstall skipifsilent
