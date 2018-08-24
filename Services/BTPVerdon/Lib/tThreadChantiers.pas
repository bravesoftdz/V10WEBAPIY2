unit tThreadChantiers;

interface

uses
  Classes
  , UtilBTPVerdon
  , ConstServices
  , uTob
  , CbpMCD
  , AdoDB
  , CommonTools
  {$IFDEF MSWINDOWS}
  , Windows
  {$ENDIF}
  ;

type
  ThreadChantiers = class(TThread)
  public
    TableValues  : T_ChantierValues;
    LogValues    : T_WSLogValues;
    FolderValues : T_FolderValues;
    lTn          : T_TablesName;
    AdoQryBTP    : AdoQry;
    AdoQryTMP    : AdoQry;
    TobT         : TOB;
    TobAdd       : TOB;
    TobQry       : TOB;

    constructor Create(CreateSuspended : boolean);
    destructor Destroy; override;

  protected
    procedure Execute; override;
  end;

implementation

uses
  SysUtils
  , hCtrls
  , hEnt1
  , StrUtils
  , ActiveX
  ;

  { Important : les méthodes et propriétés des objets de la VCL peuvent uniquement être
  utilisés dans une méthode appelée en utilisant Synchronize, comme : 

      Synchronize(UpdateCaption);

  où UpdateCaption serait de la forme

    procedure ThreadChantiers.UpdateCaption;
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

{ ThreadChantiers }

constructor ThreadChantiers.Create(CreateSuspended: boolean);
begin
  inherited Create(CreateSuspended);
end;

destructor ThreadChantiers.Destroy;
begin
  inherited;
  TUtilBTPVerdon.AddLog(lTn, TUtilBTPVerdon.GetMsgStartEnd(lTn, False, TableValues.LastSynchro), LogValues, 0);
end;

procedure ThreadChantiers.Execute;
var
  TobT      : TOB;
  TobAdd    : TOB;
  TobQry    : TOB;
  AdoQryBTP : AdoQry;
  AdoQryTMP : AdoQry;
  Treatment : TTnTreatment;
begin
  Coinitialize(nil);
  try
    TobQry := TOB.Create('_QRY', nil, -1);
    try
      TobT := TOB.Create('_TABLE', nil, -1);
      try
        TobAdd := TOB.Create('_ADDFIEDS', nil, -1);
        try
          AdoQryBTP := AdoQry.Create;
          try
            AdoQryTMP := AdoQry.Create;
            try
              AdoQryBTP.Qry := TADOQuery.Create(nil);
              try
                AdoQryTMP.Qry := TADOQuery.Create(nil);
                try
                  TUtilBTPVerdon.AddFieldsTobAdd(tnChantier, TobAdd);
                  TUtilBTPVerdon.AssignAdoQry(AdoQryBTP, AdoQryTMP, FolderValues, LogValues);
                  Treatment := TTnTreatment.Create;
                  try
                    Treatment.Tn           := tnChantier;
                    Treatment.FolderValues := FolderValues;
                    Treatment.LogValues    := LogValues;
                    Treatment.LastSynchro  := TableValues.LastSynchro;
                    Treatment.AdoQryBTP    := AdoQryBTP;
                    Treatment.AdoQryTMP    := AdoQryTMP;
                    Treatment.TnTreatment(TobT, TobAdd, TobQry);
                  finally
                    Treatment.Free;
                  end;
                finally
                  AdoQryTMP.Qry.Free;
                end;
              finally
                AdoQryBTP.Qry.Free;
              end;
            finally
              AdoQryTMP.free;
            end;
          finally
            AdoQryBTP.Free;
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
  finally
    CoUninitialize();
  end;
end;

end.
