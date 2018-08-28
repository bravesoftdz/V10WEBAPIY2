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
  , tThreadTiers
  , tThreadChantiers
  , tThreadDevis
  , tThreadLignesBR
  ;

type
  TSvcSyncBTPVerdon = class(TService)
    procedure ServiceAfterInstall(Sender: TService);
    procedure ServiceExecute(Sender: TService);
    procedure ServiceStop(Sender: TService; var Stopped: Boolean);
    procedure ServiceStart(Sender: TService; var Started: Boolean);
  private
    uThreadTiers    : ThreadTiers;
    uThreadChantier : ThreadChantiers;
    uThreadDevis    : ThreadDevis;
    uThreadLignesBR : ThreadLignesBR;
    IniPath        : string;
    AppPath        : string;
    LogPath        : string;
    LogValues      : T_WSLogValues;
    TiersValues    : T_TiersValues;
    ChantierValues : T_ChantierValues;
    DevisValues    : T_DevisValues;
    LignesBRValues : T_LignesBRValues;
    FolderValues   : T_FolderValues;

    procedure ClearTablesValues;
    function ReadSettings : boolean;

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
  , IniFiles
  , StrUtils
  , AdoDB
  ;

{$R *.DFM}

procedure ServiceController(CtrlCode: DWord); stdcall;
begin
  SvcSyncBTPVerdon.Controller(CtrlCode);
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

  function IsActive(lTn : T_TablesName) : boolean;
  var
    TableName : string;
  begin
    TableName := TUtilBTPVerdon.GetTMPTableName(lTn);
    if SettingFile.ReadString(Section, TableName, 'na') = 'na' then
      Result := True
    else
      Result := (SettingFile.ReadInteger(Section, TableName, 0) = 1);
  end;

begin
  Result := True;
  SettingFile := TIniFile.Create(IniPath);
  try
    Section := 'GLOBALSETTINGS';
    LogValues.LogLevel     := SettingFile.ReadInteger(Section, 'LogLevel', 0);
    LogValues.LogMoMaxSize := SettingFile.ReadInteger(Section, 'LogMoMaxSize', 0);
    LogValues.DebugEvents  := SettingFile.ReadInteger(Section, 'DebugEvents', 0);
    LogValues.OneLogPerDay := (SettingFile.ReadInteger(Section, 'OneLogPerDay', 0) = 1);
    Section := 'FOLDER';
    FolderValues.BTPUserAdmin := SettingFile.ReadString(Section, 'BTPUser'     , '');
    FolderValues.BTPServer    := SettingFile.ReadString(Section, 'Server'      , '');
    FolderValues.BTPDataBase  := SettingFile.ReadString(Section, 'BTPFolder'   , '');
    FolderValues.TMPServer    := SettingFile.ReadString(Section, 'Server'      , '');
    FolderValues.TMPDataBase  := SettingFile.ReadString(Section, 'TMPBDDFolder', '');
    Section := 'TABLESTRIGGERTIME';
    TiersValues.TimeOut    := SettingFile.ReadInteger(Section, TUtilBTPVerdon.GetTMPTableName(tnTiers)   , 0);
    ChantierValues.TimeOut := SettingFile.ReadInteger(Section, TUtilBTPVerdon.GetTMPTableName(tnChantier), 0);
    DevisValues.TimeOut    := SettingFile.ReadInteger(Section, TUtilBTPVerdon.GetTMPTableName(tnDevis), 0);
    LignesBRValues.TimeOut := SettingFile.ReadInteger(Section, TUtilBTPVerdon.GetTMPTableName(tnLignesBR), 0);
    Section := 'TABLESLASTSYNCHRO';
    TiersValues.LastSynchro    := SettingFile.ReadString(Section, TUtilBTPVerdon.GetTMPTableName(tnTiers)   , '');
    ChantierValues.LastSynchro := SettingFile.ReadString(Section, TUtilBTPVerdon.GetTMPTableName(tnChantier), '');
    DevisValues.LastSynchro    := SettingFile.ReadString(Section, TUtilBTPVerdon.GetTMPTableName(tnDevis), '');
    LignesBRValues.LastSynchro := SettingFile.ReadString(Section, TUtilBTPVerdon.GetTMPTableName(tnLignesBR), '');
    Section := 'TABLESISACTIVE';
    TiersValues.IsActive    := IsActive(tnTiers);
    ChantierValues.IsActive := IsActive(tnChantier);
    DevisValues.IsActive    := IsActive(tnDevis);
    LignesBRValues.IsActive := IsActive(tnLignesBR);
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

  procedure StartLog(lTn : T_TablesName; LastSynchro : string);
  begin
    TUtilBTPVerdon.AddLog(lTn, '', LogValues, 0);
    TUtilBTPVerdon.AddLog(lTn, DupeString('*', 50), LogValues, 0);
    TUtilBTPVerdon.AddLog(lTn, TUtilBTPVerdon.GetMsgStartEnd(lTn, True, LastSynchro), LogValues, 0);
  end;

  procedure CallThreadTiers;
  begin
    if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(tnTiers, Format('%sWith Thread', [WSCDS_DebugMsg]), LogValues, 0);
    StartLog(tnTiers, TiersValues.LastSynchro);
    TiersValues.FirstExec := False;
    TiersValues.Count     := 0;
    uThreadTiers                 := ThreadTiers.Create(True);
    uThreadTiers.FreeOnTerminate := True;
    uThreadTiers.Priority        := tpNormal;
    uThreadTiers.lTn             := tnTiers;
    uThreadTiers.TableValues     := TiersValues;
    uThreadTiers.LogValues       := LogValues;
    uThreadTiers.FolderValues    := FolderValues;
    if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(tnTiers, Format('%sBefore Call uThreadTiers.Resume', [WSCDS_DebugMsg]), LogValues, 0);
    try
      uThreadTiers.Resume;
    except
      on E: Exception do
        LogMessage(Format('Fin exécution du service avec erreur : %s', [E.Message]), EVENTLOG_ERROR_TYPE);
    end;
  end;

  procedure CallThreadChantiers;
  begin
    if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(tnChantier, Format('%sWith Thread', [WSCDS_DebugMsg]), LogValues, 0);
    StartLog(tnChantier, ChantierValues.LastSynchro);
    ChantierValues.FirstExec := False;
    ChantierValues.Count     := 0;
    uThreadChantier                 := ThreadChantiers.Create(True);
    uThreadChantier.FreeOnTerminate := True;
    uThreadChantier.Priority        := tpNormal;
    uThreadChantier.lTn             := tnChantier;
    uThreadChantier.TableValues     := ChantierValues;
    uThreadChantier.LogValues       := LogValues;
    uThreadChantier.FolderValues    := FolderValues;
    if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(tnChantier, Format('%sBefore Call uThreadChantier.Resume', [WSCDS_DebugMsg]), LogValues, 0);
    try
      uThreadChantier.Resume;
    except
      on E: Exception do
        LogMessage(Format('Fin exécution du service avec erreur : %s', [E.Message]), EVENTLOG_ERROR_TYPE);
    end;
  end;

  procedure CallThreadDevis;
  begin
    if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(tnDevis, Format('%sWith Thread', [WSCDS_DebugMsg]), LogValues, 0);
    StartLog(tnDevis, DevisValues.LastSynchro);
    DevisValues.FirstExec := False;
    DevisValues.Count     := 0;
    uThreadDevis                 := ThreadDevis.Create(True);
    uThreadDevis.FreeOnTerminate := True;
    uThreadDevis.Priority        := tpNormal;
    uThreadDevis.lTn             := tnDevis;
    uThreadDevis.TableValues     := DevisValues;
    uThreadDevis.LogValues       := LogValues;
    uThreadDevis.FolderValues    := FolderValues;
    if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(tnDevis, Format('%sBefore Call uThreadDevis.Resume', [WSCDS_DebugMsg]), LogValues, 0);
    try
      uThreadDevis.Resume;
    except
      on E: Exception do
        LogMessage(Format('Fin exécution du service avec erreur : %s', [E.Message]), EVENTLOG_ERROR_TYPE);
    end;
  end;

  procedure CallThreadLignesBR;
  begin
    if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(tnLignesBR, Format('%sWith Thread', [WSCDS_DebugMsg]), LogValues, 0);
    StartLog(tnLignesBR, DevisValues.LastSynchro);
    LignesBRValues.FirstExec := False;
    LignesBRValues.Count     := 0;
    uThreadLignesBR                 := ThreadLignesBR.Create(True);
    uThreadLignesBR.FreeOnTerminate := True;
    uThreadLignesBR.Priority        := tpNormal;
    uThreadLignesBR.lTn             := tnLignesBR;
    uThreadLignesBR.TableValues     := LignesBRValues;
    uThreadLignesBR.LogValues       := LogValues;
    uThreadLignesBR.FolderValues    := FolderValues;
    if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(tnLignesBR, Format('%sBefore Call uThreadLignesBR.Resume', [WSCDS_DebugMsg]), LogValues, 0);
    try
      uThreadLignesBR.Resume;
    except
      on E: Exception do
        LogMessage(Format('Fin exécution du service avec erreur : %s', [E.Message]), EVENTLOG_ERROR_TYPE);
    end;
  end;

begin
  IniPath := TServicesLog.GetFilePath(ServiceName_BTPVerdon, 'ini');
  AppPath := TServicesLog.GetFilePath(ServiceName_BTPVerdon, 'exe');
  LogPath := TServicesLog.GetFilePath(ServiceName_BTPVerdon, 'log');
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
      if (TiersValues.IsActive)    and ((TiersValues.Count    >= TiersValues.TimeOut)    or (TiersValues.FirstExec))    then CallThreadTiers;
      if (ChantierValues.IsActive) and ((ChantierValues.Count >= ChantierValues.TimeOut) or (ChantierValues.FirstExec)) then CallThreadChantiers;
      if (DevisValues.IsActive)    and ((DevisValues.Count    >= DevisValues.TimeOut)    or (DevisValues.FirstExec))    then CallThreadDevis;
      if (LignesBRValues.IsActive) and ((LignesBRValues.Count >= LignesBRValues.TimeOut) or (LignesBRValues.FirstExec)) then CallThreadLignesBR;
      Sleep(1000);
      ServiceThread.ProcessRequests(False);
    end;
  end;
end;

procedure TSvcSyncBTPVerdon.ServiceStop(Sender: TService; var Stopped: Boolean);
begin
  FreeAndNil(uThreadTiers);
  FreeAndNil(uThreadChantier);
  FreeAndNil(uThreadDevis);
  FreeAndNil(uThreadLignesBR);
  LogMessage(Format('Arrêt de %s.', [ServiceName_BTPVerdon]), EVENTLOG_INFORMATION_TYPE);
end;

procedure TSvcSyncBTPVerdon.ServiceStart(Sender: TService; var Started: Boolean);
begin
  LogMessage(Format('Démarrage de de %s.', [ServiceName_BTPVerdon]), EVENTLOG_INFORMATION_TYPE);
end;

end.

