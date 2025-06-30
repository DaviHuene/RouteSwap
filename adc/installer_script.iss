[Setup]
AppName=Troca de Técnicos Claro
AppVersion=1.0
DefaultDirName={pf}\TrocaTecnicosClaro
DefaultGroupName=Troca de Técnicos Claro
AllowNoIcons=yes
OutputDir=dist
OutputBaseFilename=Instalador_TrocaTecnicos
Compression=lzma
SolidCompression=yes
SetupIconFile="C:\Users\Davi\Desktop\meuapp\Roteswap.ico"


[Files]
Source: "dist\Roteswap.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "logins.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "chromedriver.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Troca de Técnicos"; Filename: "{app}\Roteswap.exe"
Name: "{commondesktop}\Troca de Técnicos"; Filename: "{app}\Roteswap.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Criar atalho na Área de Trabalho"; GroupDescription: "Opções adicionais:"