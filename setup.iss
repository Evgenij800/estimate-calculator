[Setup]
AppId={{E5A9C1B3-7D42-4F68-9B1E-3C2A6D8F0457}
AppName=Калькулятор каркасного дома
AppVersion=1.0
AppPublisher=Локальный калькулятор стоимости
DefaultDirName={localappdata}\Programs\KarkasCalculator
DefaultGroupName=Калькулятор каркасного дома
UninstallDisplayName=Калькулятор каркасного дома
UninstallDisplayIcon={app}\KarkasCalculator.exe
OutputDir=installer
OutputBaseFilename=Калькулятор_каркасного_дома_Установщик
SetupIconFile=icon.ico
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
DisableProgramGroupPage=yes
ShowLanguageDialog=no

[Languages]
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"

[Tasks]
Name: "desktopicon"; Description: "Создать ярлык на рабочем столе"; GroupDescription: "Дополнительно:"

[Files]
Source: "dist\KarkasCalculator\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion

[Icons]
Name: "{group}\Калькулятор каркасного дома"; Filename: "{app}\KarkasCalculator.exe"; WorkingDir: "{app}"
Name: "{group}\Удалить Калькулятор каркасного дома"; Filename: "{uninstallexe}"
Name: "{autodesktop}\Калькулятор каркасного дома"; Filename: "{app}\KarkasCalculator.exe"; WorkingDir: "{app}"; Tasks: desktopicon

[Run]
Filename: "{app}\KarkasCalculator.exe"; Description: "Запустить Калькулятор каркасного дома"; WorkingDir: "{app}"; Flags: nowait postinstall skipifsilent
