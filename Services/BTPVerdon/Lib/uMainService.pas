unit uMainService;

interface

uses
  Windows
  , Messages
  , SysUtils
  , Classes
  , Graphics
  , Controls
  , SvcMgr
  , Dialogs
  , ConstServices
  , UtilBTPVerdon
  ;

type
  TSvcSyncBTPVerdon = class(TService)
    procedure ServiceAfterInstall(Sender: TService);
    procedure ServiceExecute(Sender: TService);
    procedure ServiceStop(Sender: TService; var Stopped: Boolean);
    procedure ServiceStart(Sender: TService; var Started: Boolean);
  private
    IniPath        : string;
    AppPath        : string;
    LogPath        : string;
    LogValues      : T_WSLogValues;
    TiersValues    : T_TiersValues;
    ChantierValues : T_ChantierValues;
    DevisValues    : T_DevisValues;
    LignesBRValues : T_LignesBRValues;

    procedure ClearTablesValues;
    function ReadSettings : boolean;
    function GetFilePath(Extension : string) : string;
  public
    function GetServiceController: TServiceController; override;
  end;

var
  SvcSyncBTPVerdon: TSvcSyncBTPVerdon;

implementation

uses
  Registry
  , CommonTools
  , ActiveX
  , WinSVC
  , ShellAPI
  , tThreadTiers
  , tThreadChantiers
  , IniFiles
  ;

{$R *.DFM}

procedure ServiceController(CtrlCode: DWord); stdcall;
begin
  SvcSyncBTPVerdon.Controller(CtrlCode);
end;

function TSvcSyncBTPVerdon.GetFilePath(Extension: string): string;
begin
  Result := Format('%s%s.%s', [ExtractFilePath(ParamStr(0)), ServiceName_BTPVerdon, Extension]);
end;

procedure TSvcSyncBTPVerdon.ClearTablesValues;
begin
  TiersValues.FirstExec    := True;
  TiersValues.Count        := 0;
  TiersValues.TimeOut      := 0;
  ChantierValues.FirstExec := True;
  ChantierValues.Count     := 0;
  ChantierValues.TimeOut   := 0;
  DevisValues.FirstExec    := True;
  DevisValues.Count        := 0;
  DevisValues.TimeOut      := 0;
  LignesBRValues.FirstExec := True;
  LignesBRValues.Count     := 0;
  LignesBRValues.TimeOut   := 0;
end;

function TSvcSyncBTPVerdon.ReadSettings : boolean;
var
  SettingFile : TInifile;
  Section     : string;
begin
  Result := True;
  SettingFile := TIniFile.Create(IniPath);
  try
    Section := 'GLOBALSETTINGS';
    LogValues.LogLevel     := SettingFile.ReadInteger(Section, 'LogLevel', 0);
    LogValues.LogMoMaxSize := SettingFile.ReadInteger(Section, 'LogMoMaxSize', 0);
    LogValues.DebugEvents  := SettingFile.ReadInteger(Section, 'DebugEvents', 0);
    LogValues.OneLogPerDay := (SettingFile.ReadInteger(Section, 'OneLogPerDay', 0) = 1);
    Section := 'TABLESTRIGGERTIME';
    TiersValues.TimeOut    := SettingFile.ReadInteger(Section, TUtilBTPVerdon.GetSynchroTableName(tnTiers)   , 0);
    ChantierValues.TimeOut := SettingFile.ReadInteger(Section, TUtilBTPVerdon.GetSynchroTableName(tnChantier), 0);
    DevisValues.TimeOut    := SettingFile.ReadInteger(Section, TUtilBTPVerdon.GetSynchroTableName(tnDevis), 0);
    LignesBRValues.TimeOut := SettingFile.ReadInteger(Section, TUtilBTPVerdon.GetSynchroTableName(tnLignesBR), 0);
  finally
    SettingFile.Free;
  end;
end;

function TSvcSyncBTPVerdon.GetServiceController: TServiceController;
begin
  Result := ServiceController;
end;

procedure TSvcSyncBTPVerdon.ServiceAfterInstall(Sender: TService);
var
  Reg : TRegistry;
begin                                                                               
  Reg := TRegistry.Create(KEY_READ or KEY_WRITE);
  try
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    if Reg.OpenKey('\SYSTEM\CurrentControlSet\Services\' + Sender.Name, false) then
    try
      Reg.WriteString('Description', 'LSE-Synchronisation des données entre BTP et VERDON.');
    finally
      Reg.CloseKey;
    end;
  finally
    Reg.Free;                                                                              
  end;
end;                                                                                    

procedure TSvcSyncBTPVerdon.ServiceExecute(Sender: TService);
var
  uThreadTiers    : ThreadTiers;
  uThreadChantier : ThreadChantiers;
begin
  IniPath := GetFilePath('ini');
  AppPath := GetFilePath('exe');
  LogPath := GetFilePath('log');
  if not FileExists(IniPath) then
  begin
    LogMessage(Format('Impossible d''initialiser le service %s. Le fichier de configuration "%s" est inexistant.', [ServiceName_BTPVerdon, IniPath]), EVENTLOG_ERROR_TYPE);
  end else
  begin
    ClearTablesValues;
    ReadSettings;
    while not Terminated do
    begin
      Inc(TiersValues.Count);
      Inc(ChantierValues.Count);
      Inc(DevisValues.Count);
      Inc(LignesBRValues.Count);
      { Traitement des tiers }
      if (TiersValues.Count >= TiersValues.TimeOut) or (TiersValues.FirstExec) then
      begin
        TiersValues.FirstExec := False;
        TiersValues.Count    := 0;
        try
          uThreadTiers := ThreadTiers.Create(True);
          uThreadTiers.TiersValues := TiersValues;
          uThreadTiers.LogValues   := LogValues;
          uThreadTiers.Resume;
        except
          on E: Exception do
            LogMessage(Format('Fin exécution du service avec erreur : %s', [E.Message]), EVENTLOG_ERROR_TYPE);
        end;
      end;
      { Traitement des chantiers }
      if (ChantierValues.Count >= ChantierValues.TimeOut) or (ChantierValues.FirstExec) then
      begin
        ChantierValues.FirstExec := False;
        ChantierValues.Count    := 0;
        try
          uThreadChantier                := ThreadChantiers.Create(True);
          uThreadChantier.ChantierValues := ChantierValues;
          uThreadChantier.LogValues      := LogValues;
          uThreadChantier.Resume;
        except
          on E: Exception do
            LogMessage(Format('Fin exécution du service avec erreur : %s', [E.Message]), EVENTLOG_ERROR_TYPE);
        end;
      end;

      Sleep(1000);
      ServiceThread.ProcessRequests(False);
    end;
  end;
end;

procedure TSvcSyncBTPVerdon.ServiceStop(Sender: TService; var Stopped: Boolean);
begin
  CoUnInitialize;
  LogMessage(Format('Arrêt de %s.', [ServiceName_BTPVerdon]), EVENTLOG_INFORMATION_TYPE);
end;

procedure TSvcSyncBTPVerdon.ServiceStart(Sender: TService; var Started: Boolean);
begin
  LogMessage(Format('Démarrage de de %s.', [ServiceName_BTPVerdon]), EVENTLOG_INFORMATION_TYPE);
  Coinitialize(nil);
end;


end.
