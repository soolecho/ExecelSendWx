#define MyAppName "表格自动发送By春风予Lu"
#define MyAppVersion "1.4.23"
#define MyAppPublisher "春风予Lu"
#define MyAppURL "https://github.com/soolecho/ExecelSendWx"
#define MyAppExeName "表格自动发送By春风予Lu.exe"
#define MyShortcutName "自动微信发送"
#define MyAppSourceDir "dist_nuitka_config\main.dist"
#define MyAppIconFile "love.ico"

[Setup]
AppId={{60988979-D944-41FE-8DAC-28220E81FB69}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppComments=本地 Excel 数据整理与微信自动发送工具
AppContact={#MyAppPublisher}
AppCopyright=Copyright (C) 2026 {#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}/issues
AppUpdatesURL={#MyAppURL}/releases
DefaultDirName={localappdata}\Programs\ExcelSendWx
DefaultGroupName={#MyAppName}
DisableWelcomePage=no
DisableDirPage=no
DisableProgramGroupPage=yes
AllowNoIcons=yes
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
MinVersion=10.0
OutputDir=installer_output
OutputBaseFilename=表格自动发送安装包_v{#MyAppVersion}
SetupIconFile={#MyAppIconFile}
InfoBeforeFile=installer_info.txt
Uninstallable=yes
SignedUninstaller=yes
CreateUninstallRegKey=yes
UninstallDisplayName={#MyAppName}
UninstallDisplayIcon={app}\love.ico
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
CloseApplications=yes
RestartApplications=no
SetupLogging=yes
UsePreviousAppDir=yes
UsePreviousGroup=yes
VersionInfoVersion={#MyAppVersion}
VersionInfoCompany={#MyAppPublisher}
VersionInfoCopyright=Copyright (C) 2026 {#MyAppPublisher}
VersionInfoDescription={#MyAppName} 安装程序
VersionInfoProductName={#MyAppName}
VersionInfoProductVersion={#MyAppVersion}

[Languages]
Name: "chinesesimp"; MessagesFile: "installer_assets\ChineseSimplified.isl"

[Messages]
BeveledLabel=作者：{#MyAppPublisher}
WelcomeLabel1=欢迎使用 {#MyAppName} 安装向导
WelcomeLabel2=本向导将安装 [name/ver]。%n%n作者：{#MyAppPublisher}%n%n建议在继续安装前退出正在运行的软件。
InstallingLabel=正在安装 [name]，请稍候...
StatusExtractFiles=正在解压程序文件...
StatusCreateIcons=正在创建快捷方式...
StatusCreateRegistryEntries=正在注册安装和卸载信息...
StatusRunProgram=正在启动 [name]...

[Tasks]
Name: "startmenuicon"; Description: "创建开始菜单快捷方式"; GroupDescription: "快捷方式："; Flags: checkedonce
Name: "desktopicon"; Description: "创建桌面快捷方式"; GroupDescription: "快捷方式："; Flags: unchecked
Name: "autostart"; Description: "开机自动启动（当前用户）"; GroupDescription: "其他选项："; Flags: unchecked

[Files]
Source: "{#MyAppSourceDir}\*"; DestDir: "{app}"; Excludes: "wxauto_logs\*;*.log"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyShortcutName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\love.ico"; Tasks: startmenuicon
Name: "{group}\卸载 {#MyShortcutName}"; Filename: "{uninstallexe}"; IconFilename: "{app}\love.ico"; Tasks: startmenuicon
Name: "{userdesktop}\{#MyShortcutName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\love.ico"; Tasks: desktopicon

[Registry]
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "ExcelSendWx"; ValueData: """{app}\{#MyAppExeName}"""; Flags: uninsdeletevalue; Tasks: autostart
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: none; ValueName: "ExcelSendWx"; Flags: deletevalue; Tasks: not autostart

[Run]
; 交互式安装：最后一页勾选后启动（静默安装时跳过）
Filename: "{app}\{#MyAppExeName}"; Description: "运行 {#MyShortcutName}"; Flags: nowait postinstall skipifsilent
; 在线更新场景：程序以 /SILENT 启动安装器前会在 %%TEMP%% 写重启标记文件，
; 检测到标记时静默安装结束后自动启动新版（交互安装不受影响）
Filename: "{app}\{#MyAppExeName}"; Flags: nowait; Check: ShouldLaunchAfterSilent

[Code]
function ShouldLaunchAfterSilent: Boolean;
var
  FlagPath: String;
begin
  Result := False;
  if WizardSilent then
  begin
    FlagPath := AddBackslash(GetEnv('TEMP')) + 'excel_send_wx_restart.flag';
    if FileExists(FlagPath) then
    begin
      DeleteFile(FlagPath);
      Result := True;
    end;
  end;
end;
