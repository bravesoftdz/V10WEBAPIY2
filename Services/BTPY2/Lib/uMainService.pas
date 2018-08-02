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
  ;

type
  TSvcSyncBTPY2 = class(TService)
    procedure ServiceAfterInstall(Sender: TService);
    procedure ServiceExecute(Sender: TService);
    procedure ServiceStop(Sender: TService; var Stopped: Boolean);
    procedure ServiceStart(Sender: TService; var Started: Boolean);
  private
    function GetFilePath(Extension : string) : string;
  public
    function GetServiceController: TServiceController; override;
  end;

var
  SvcSyncBTPY2: TSvcSyncBTPY2;

implementation

uses
  Registry
  , uThreadExecute
  , CommonTools
  , uExecuteService
  , ActiveX
  , WinSVC
  , UConnectWSConst
  , ShellAPI
  ;

{$R *.DFM}

procedure ServiceController(CtrlCode: DWord); stdcall;
begin
  SvcSyncBTPY2.Controller(CtrlCode);
end;

function TSvcSyncBTPY2.GetServiceController: TServiceController;
begin
  Result := ServiceController;
end;

procedure TSvcSyncBTPY2.ServiceAfterInstall(Sender: TService);
var
  Reg : TRegistry;
begin                                                                               
  Reg := TRegistry.Create(KEY_READ or KEY_WRITE);
  try
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    if Reg.OpenKey('\SYSTEM\CurrentControlSet\Services\' + Sender.Name, false) then
    try
      Reg.WriteString('Description', 'LSE-Synchronisation des données comptables entre BTP et Y2.');
    finally
      Reg.CloseKey;
    end;
  finally
    Reg.Free;                                                                              
  end;
end;                                                                                    

procedure TSvcSyncBTPY2.ServiceExecute(Sender: TService);
var
  Count     : Integer;
  BTPY2Exec : TSvcSyncBTPY2Execute;
  IniPath   : string;
  AppPath   : string;
  LogPath   : string;
  FirstExec : boolean;  
  //uExecute  : SynchroThread;
begin
  IniPath := GetFilePath('ini');
  AppPath := GetFilePath('exe');
  LogPath := GetFilePath('log');
  if not FileExists(IniPath) then
  begin
    LogMessage(Format('Impossible d''initialiser le service %s. Le fichier de configuration "%s" est inexistant.', [WSCDS_ServiceName, GetFilePath('ini')]), EVENTLOG_ERROR_TYPE);
  end else
  begin
    FirstExec := True;
    BTPY2Exec := TSvcSyncBTPY2Execute.Create;
    try
      BTPY2Exec.CreateObjects;
      try
        BTPY2Exec.IniFilePath := IniPath;
        BTPY2Exec.AppFilePath := AppPath;
        BTPY2Exec.LogFilePath := LogPath;
        BTPY2Exec.InitApplication;
        try
          Count := 0;
          while not Terminated do
          begin
            Inc(Count);
            if (Count >= BTPY2Exec.SecondTimeout) or (FirstExec) then
            begin
              FirstExec := False;
              Count     := 0;
              try
                LogMessage('Début d''exécution du service.', EVENTLOG_INFORMATION_TYPE);
                BTPY2Exec.ServiceExecute;
                try
                  BTPY2Exec.LogFilePath := LogPath;
                  //uExecute := SynchroThread.Create(True);
                  //uExecute.ServiceTreatment := BTPY2Exec;
                  //uExecute.Resume;
                finally
                  LogMessage('Fin d''exécution du service.', EVENTLOG_INFORMATION_TYPE);
                end;
              except
                on E: Exception do
                  LogMessage(Format('Fin exécution du service avec erreur : %s', [E.Message]), EVENTLOG_ERROR_TYPE);
              end;
            end;
            Sleep(1000);
            ServiceThread.ProcessRequests(False);
          end;
        finally
        end;
      finally
        BTPY2Exec.FreeObjects;
      end;
    finally
      BTPY2Exec.Free;
    end;
  end;
end;

procedure TSvcSyncBTPY2.ServiceStop(Sender: TService; var Stopped: Boolean);
begin
  CoUnInitialize;
  LogMessage(Format('Arrêt de %s.', [WSCDS_ServiceName]), EVENTLOG_INFORMATION_TYPE);
end;

procedure TSvcSyncBTPY2.ServiceStart(Sender: TService; var Started: Boolean);                                             
begin
  LogMessage(Format('Démarrage de de %s.', [WSCDS_ServiceName]), EVENTLOG_INFORMATION_TYPE);
  Coinitialize(nil);
end;

function TSvcSyncBTPY2.GetFilePath(Extension: string): string;
begin
  Result := Format('%s%s.%s', [ExtractFilePath(ParamStr(0)), WSCDS_ServiceName, Extension]);
end;

end.
