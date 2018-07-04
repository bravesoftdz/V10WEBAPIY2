unit uMainService;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, SvcMgr, Dialogs;

type
  TSvcSyncBTPY2 = class(TService)
    procedure ServiceAfterInstall(Sender: TService);
    procedure ServiceExecute(Sender: TService);
    procedure ServiceStop(Sender: TService; var Stopped: Boolean);
  public
    function GetServiceController: TServiceController; override;
  end;

var
  SvcSyncBTPY2: TSvcSyncBTPY2;

implementation

uses
  Registry
//  , uThreadExecute
  , CommonTools
  , uExecuteService
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
  Reg: TRegistry;
begin
  Reg := TRegistry.Create(KEY_READ or KEY_WRITE);
  try
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    if Reg.OpenKey('\SYSTEM\CurrentControlSet\Services\' + Sender.Name, false) then
    try
      Reg.WriteString('Description', 'Synchronisation des données comptables entre BTP et Y2.');
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
//uExecute : SynchroThread;
begin
//  uExecute := SynchroThread.Create(True);
//  uExecute.FreeOnTerminate := True;

  BTPY2Exec := TSvcSyncBTPY2Execute.Create;
  try
    BTPY2Exec.CreateObjects;
    try
      BTPY2Exec.InitApplication;
      Count := 0;
      while not Terminated do
      begin
        Inc(Count);
        if Count >= BTPY2Exec.SecondTimeout then
        begin
          Count := 0;
          BTPY2Exec.ServiceExecute;
        end;
        Sleep(1000);
        ServiceThread.ProcessRequests(False);
      end;
    finally
      BTPY2Exec.FreeObjects;
    end;
  finally
    BTPY2Exec.Free;
  end;
end;

procedure TSvcSyncBTPY2.ServiceStop(Sender: TService; var Stopped: Boolean);
var
  WindowsLog : TEventLogger;
begin
  WindowsLog := TEventLogger.Create(Application.Name);
  try
    WindowsLog.LogMessage(Format('Démarrage de %s.', [Application.Name]), EVENTLOG_INFORMATION_TYPE);
  finally
    WindowsLog.Free;
  end;
end;

end.
