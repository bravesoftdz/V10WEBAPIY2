unit uThreadExecute;

interface

uses
  Classes
  , uExecuteService
  ;

type
  SynchroThread = class(TThread)
  public
    ServiceTreatment : TSvcSyncBTPY2Execute;

    constructor Create(CreateSuspended : boolean);
    destructor Destroy; override;

  protected
    procedure Execute; override;
  end;

(*
  R_Params = record
    ServerName  : string;
    DBName      : string;
    LastSynchro : string;
  end;
*)

implementation

{ SynchroThread }

constructor SynchroThread.Create(CreateSuspended : boolean);
begin
  inherited Create(CreateSuspended);
  FreeOnTerminate := True;
  Priority        := tpNormal;
end;

destructor SynchroThread.Destroy;
begin
  inherited;
end;

procedure SynchroThread.Execute;
begin
  ServiceTreatment.ServiceExecute;
end;

end.
