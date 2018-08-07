unit tThreadLignesBR;

interface

uses
  Classes
  , UtilBTPVerdon
  , ConstServices
  {$IFDEF MSWINDOWS}
  , Windows
  {$ENDIF}
  ;

type
  ThreadLignesBR = class(TThread)
  public
    LignesBRValues : T_LignesBRValues;
    LogValues      : T_WSLogValues;
    FolderValues   : T_FolderValues;

    constructor Create(CreateSuspended : boolean);
    destructor Destroy; override;
  private
    lTn : T_TablesName;
//    procedure SetName;
  protected
    procedure Execute; override;
  end;

implementation

{ Important : les méthodes et propriétés des objets de la VCL peuvent uniquement être
  utilisés dans une méthode appelée en utilisant Synchronize, comme : 

      Synchronize(UpdateCaption);

  où UpdateCaption serait de la forme

    procedure ThreadLignesBR.UpdateCaption;
    begin
      Form1.Caption := 'Mis à jour dans un thread';
    end; }

{$IFDEF MSWINDOWS}
type
  TThreadNameInfo = record
    FType: LongWord;     // doit être 0x1000
    FName: PChar;        // pointeur sur le nom (dans l'espace d'adresse de l'utilisateur)
    FThreadID: LongWord; // ID de thread (-1=thread de l'appelant)
    FFlags: LongWord;    // réservé pour une future utilisation, doit être zéro
  end;
{$ENDIF}

{ ThreadLignesBR }

(*
procedure ThreadLignesBR.SetName;
{$IFDEF MSWINDOWS}
var
  ThreadNameInfo: TThreadNameInfo;
{$ENDIF}
begin
{$IFDEF MSWINDOWS}
  ThreadNameInfo.FType := $1000;
  ThreadNameInfo.FName := 'ThreadNameLignesBR';
  ThreadNameInfo.FThreadID := $FFFFFFFF;
  ThreadNameInfo.FFlags := 0;

  try
    RaiseException( $406D1388, 0, sizeof(ThreadNameInfo) div sizeof(LongWord), @ThreadNameInfo );
  except
  end;
{$ENDIF}
end;
*)

constructor ThreadLignesBR.Create(CreateSuspended: boolean);
begin
  inherited Create(CreateSuspended);
  FreeOnTerminate := True;
  Priority        := tpNormal;
  lTn             := tnLignesBR;
end;

destructor ThreadLignesBR.Destroy;
begin
  inherited;
  TUtilBTPVerdon.AddLog(lTn, TUtilBTPVerdon.GetMsgStartEnd(lTn, False, LignesBRValues.LastSynchro), LogValues, 0);
end;

procedure ThreadLignesBR.Execute;
begin
//  SetName;
  try
    TUtilBTPVerdon.AddLog(lTn, TUtilBTPVerdon.GetMsgStartEnd(lTn, True, LignesBRValues.LastSynchro), LogValues, 0);
//    TUtilBTPVerdon.SetLastSynchro(lTn);
  except
  end;
end;

end.
