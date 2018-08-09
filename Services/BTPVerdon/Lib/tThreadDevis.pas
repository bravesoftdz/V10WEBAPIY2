unit tThreadDevis;

interface

uses
  Classes
  , UtilBTPVerdon
  , ConstServices
  , uTob
  , CbpMCD
  {$IFDEF MSWINDOWS}
  , Windows
  {$ENDIF}
  ;

type
  ThreadDevis = class(TThread)
  public
    DevisValues  : T_DevisValues;
    LogValues    : T_WSLogValues;
    FolderValues : T_FolderValues;

    constructor Create(CreateSuspended : boolean);
    destructor Destroy; override;

  private
    lTn : T_TablesName;
//    procedure SetName;
  protected
    procedure Execute; override;
  end;

implementation

uses
  CommonTools
  , SysUtils
  , hCtrls
  , hEnt1
  , StrUtils
  ;

{ Important : les méthodes et propriétés des objets de la VCL peuvent uniquement être
  utilisés dans une méthode appelée en utilisant Synchronize, comme : 

      Synchronize(UpdateCaption);

  où UpdateCaption serait de la forme

    procedure ThreadDevis.UpdateCaption;
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

{ ThreadDevis }

(*
procedure ThreadDevis.SetName;
{$IFDEF MSWINDOWS}
var
  ThreadNameInfo: TThreadNameInfo;
{$ENDIF}
begin
{$IFDEF MSWINDOWS}
  ThreadNameInfo.FType := $1000;
  ThreadNameInfo.FName := 'ThreadNameDevis';
  ThreadNameInfo.FThreadID := $FFFFFFFF;
  ThreadNameInfo.FFlags := 0;

  try
    RaiseException( $406D1388, 0, sizeof(ThreadNameInfo) div sizeof(LongWord), @ThreadNameInfo );
  except
  end;
{$ENDIF}
end;
*)

constructor ThreadDevis.Create(CreateSuspended: boolean);
begin
  inherited Create(CreateSuspended);
  FreeOnTerminate := True;
  Priority        := tpNormal;
  lTn             := tnDevis;
end;

destructor ThreadDevis.Destroy;
begin
  inherited;
  TUtilBTPVerdon.AddLog(lTn, TUtilBTPVerdon.GetMsgStartEnd(lTn, False, DevisValues.LastSynchro), LogValues, 0);
end;

procedure ThreadDevis.Execute;
var
  TobT                   : TOB;
  TobAdd                 : TOB;
  TobQry                 : TOB;
  AdoQryL                : AdoQry;
  BTPArrFields           : array of string;
  TMPArrFields           : array of string;
  BTPArrAdditionalFields : array of string;
  TMPArrAdditionalFields : array of string;
  Treatment              : TTnTreatment;
begin
//  SetName;
  TUtilBTPVerdon.AddLog(lTn, '', LogValues, 0);
  TUtilBTPVerdon.AddLog(lTn, DupeString('*', 50), LogValues, 0);
  if (LogValues.DebugEvents > 0) then
    TUtilBTPVerdon.AddLog(lTn, Format('%sThreadTiers.Execute / BTPSrv=%s, BTPFolder=%s, TMPSrv=%s, TMPFolder=%s', [WSCDS_DebugMsg, FolderValues.BTPServer, FolderValues.BTPDataBase, FolderValues.TMPServer, FolderValues.TMPDataBase]), LogValues, 0);
  TUtilBTPVerdon.AddLog(lTn, TUtilBTPVerdon.GetMsgStartEnd(lTn, True, DevisValues.LastSynchro), LogValues, 0);
  TobQry := TOB.Create('_QRY', nil, -1);
  try
    TobT := TOB.Create('_DEVIS', nil, -1);
    try
      TobAdd := TOB.Create('_ADDFIEDS', nil, -1);
      try
        AdoQryL := AdoQry.Create;
        try
          AdoQryL.ServerName  := FolderValues.TMPServer;
          AdoQryL.DBName      := FolderValues.TMPDataBase;
          AdoQryL.PgiDB       := '-';
          Treatment := TTnTreatment.Create;
          try
            Treatment.Tn           := lTn;
            Treatment.FolderValues := FolderValues;
            Treatment.LogValues    := LogValues;
            Treatment.LastSynchro  := DevisValues.LastSynchro;
            //Treatment.TnTreatment(TobT, TobAdd, TobQry, AdoQryL, BTPArrFields, TMPArrFields, BTPArrAdditionalFields,TMPArrAdditionalFields);
            Treatment.TnTreatment(TobT, TobAdd, TobQry, AdoQryL);
          finally
            Treatment.Free;
          end;
        finally
          AdoQryL.free;
        end;
      finally
        FreeAndNil(TobAdd);
      end;
    finally
      FreeAndNil(TobT);
    end;
  finally    
    FreeAndNil(TobQry);
  end;
end;

end.
