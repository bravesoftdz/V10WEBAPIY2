unit uExecuteService;

interface

uses
  Windows
  , Classes
  , CommonTools
  ;

type

  TSvcSyncBTPVerdonExecute = class(TObject)
  private

  public
    AppFilePath   : string;
    IniFilePath   : string;
    LogFilePath   : string;
    TimeoutTiers : integer;

    procedure CreateObjects;
    procedure FreeObjects;
    function ServiceExecute: Boolean;
    procedure InitApplication;
  end;

implementation

uses
  Registry
  , Graphics
  , Controls
  , Dialogs
  , SysUtils
  , Messages
  , SvcMgr
  , StrUtils
  , IniFiles
  , DateUtils
  , uLog
  {$IF not defined(APPSRV)}
  , ParamSoc
  {$IFEND (APPSRV)}
  ;


{ TSvcSyncBTPVerdonExecute }

procedure TSvcSyncBTPVerdonExecute.CreateObjects;
begin

end;

procedure TSvcSyncBTPVerdonExecute.FreeObjects;
begin

end;

procedure TSvcSyncBTPVerdonExecute.InitApplication;
begin

end;

function TSvcSyncBTPVerdonExecute.ServiceExecute: Boolean;
begin
  Result := True;
end;

end.

