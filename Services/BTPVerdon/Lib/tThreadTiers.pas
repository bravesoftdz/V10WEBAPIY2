unit tThreadTiers;

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
  ThreadTiers = class(TThread)
  public
    TiersValues   : T_TiersValues;
    LogValues     : T_WSLogValues;
    FolderValues  : T_FolderValues;
    
    constructor Create(CreateSuspended : boolean);
    destructor Destroy; override;

  private
    lTn           : T_TablesName;
//    BTPArrFields  : array of string;
//    TMPArrFields  : array of string;

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

    procedure ThreadTiers.UpdateCaption;
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

{ ThreadTiers }

(*
procedure ThreadTiers.SetName;
{$IFDEF MSWINDOWS}
var
  ThreadNameInfo: TThreadNameInfo;
{$ENDIF}
begin
{$IFDEF MSWINDOWS}
  ThreadNameInfo.FType := $1000;
  ThreadNameInfo.FName := 'ThreadNameTiers';
  ThreadNameInfo.FThreadID := $FFFFFFFF;
  ThreadNameInfo.FFlags := 0;

  try
    RaiseException( $406D1388, 0, sizeof(ThreadNameInfo) div sizeof(LongWord), @ThreadNameInfo );
  except
  end;
{$ENDIF}
end;
*)

constructor ThreadTiers.Create(CreateSuspended : boolean);
begin
  inherited Create(CreateSuspended);
  FreeOnTerminate := True;
  Priority        := tpNormal;
  lTn             := tnTiers;
end;

destructor ThreadTiers.Destroy;
begin
  inherited;
  TUtilBTPVerdon.AddLog(lTn, TUtilBTPVerdon.GetMsgStartEnd(lTn, False, TiersValues.LastSynchro), LogValues, 0);
end;

procedure ThreadTiers.Execute;
var
  TobT      : TOB;
  TobQry    : TOB;
  AdoQryL   : AdoQry;
  Treatment : TTnTreatment;
(*
  TobL      : TOB;
  AdoQryL   : AdoQry;
  Cpt       : integer;
  InsertQty : integer;
  UpdateQty : integer;
  OtherQty  : integer;
  SqlUnlock : string;
  Sql       : string;
  Lock      : string;
  Treat     : string;
  KeyValues : string;
  KeyValue1 : string;
  KeyValue2 : string;
  Values    : string;
  FindData  : boolean;
*)
begin
//  SetName;
  TUtilBTPVerdon.AddLog(lTn, '', LogValues, 0);
  TUtilBTPVerdon.AddLog(lTn, DupeString('*', 50), LogValues, 0);
  if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(lTn, Format('%sThreadTiers.Execute / BTPSrv=%s, BTPFolder=%s, TMPSrv=%s, TMPFolder=%s'
                                                                        , [WSCDS_DebugMsg
                                                                           , FolderValues.BTPServer
                                                                           , FolderValues.BTPDataBase
                                                                           , FolderValues.TMPServer
                                                                           , FolderValues.TMPDataBase
                                                                          ]), LogValues, 0);
  TUtilBTPVerdon.AddLog(lTn, TUtilBTPVerdon.GetMsgStartEnd(lTn, True, TiersValues.LastSynchro), LogValues, 0);
  TobQry := TOB.Create('_QRY', nil, -1);
  try
    TobT := TOB.Create('_TIERS', nil, -1);
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
          Treatment.LastSynchro  := TiersValues.LastSynchro;
          Treatment.TnTreatment(TobT, TobQry, AdoQryL);
        finally
          Treatment.Free;
        end;
      finally
        AdoQryL.free;
      end;
(*
      InsertQty := 0;
      UpdateQty := 0;
      OtherQty  := 0;
      SetFieldsArray;
      Sql := TUtilBTPVerdon.GetDataSearchSql(tnTiers, BTPArrFields, TiersValues.LastSynchro);
      if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(lTn, Format('%sSql recherche Tiers = %s', [WSCDS_DebugMsg, Sql]), LogValues, 0);
      TobT.LoadDetailFromSQL(Sql);
      Sql := '';
      if TobT.Detail.Count > 0 then
      begin
        TUtilBTPVerdon.AddLog(lTn, Format('Recherche des données (%s enregistrement(s) de la table %s)', [IntToStr(TobT.Detail.Count), TUtilBTPVerdon.GetTMPTableName(lTn)]), LogValues, 1);
        AdoQryL := AdoQry.Create;
        try
          AdoQryL.ServerName  := FolderValues.TMPServer;
          AdoQryL.DBName      := FolderValues.TMPDataBase;
          AdoQryL.PgiDB       := '-';
          AdoQryL.FieldsList  := TUtilBTPVerdon.GetFieldsList(lTn); 
          AdoQryL.LogValues   := LogValues;
          AdoQryL.TSLResult.Clear;
          SqlUnlock := TUtilBTPVerdon.GetSqlUnlock(lTn);
          FindData  := False;
          for Cpt := 0 to pred(TobT.Detail.Count) do
          begin
            TobL      := TobT.Detail[Cpt];
            KeyValues := TUtilBTPVerdon.GetValueKey(lTn, TobL);
            KeyValue1 := Tools.ReadTokenSt_(KeyValues, ';');
            KeyValue2 := Tools.ReadTokenSt_(KeyValues, ';');
            SqlUnlock := Format('%s%s''%s''', [SqlUnlock, Tools.iif(Cpt = 0, '', ', '), KeyValue1]); // Prépare update de Unlock
            AdoQryL.Request := TUtilBTPVerdon.GetSqlDataExist(lTn, AdoQryL.FieldsList, KeyValue1); // Test si enregistrement existe
            AdoQryL.SingleTableSelect;
            if LogValues.LogLevel = 2 then TUtilBTPVerdon.AddLog(lTn, Format('%s - %s (%s)', [KeyValue1, KeyValue2 , Tools.iif(AdoQryL.RecordCount = 1, 'à modifier', 'à créer')]), LogValues, 3);
            if AdoQryL.RecordCount = 1 then // Update
            begin
              Values := AdoQryL.TSLResult[0];
              Lock   := Tools.ReadTokenSt_(Values, ToolsTobToTsl_Separator);
              Treat  := Tools.ReadTokenSt_(Values, ToolsTobToTsl_Separator);
              if Lock = 'N' then
              begin
                inc(UpdateQty);
                FindData  := True;
                Sql := TUtilBTPVerdon.GetSqlUpdate(lTn, FolderValues, TobL, TMPArrFields, BTPArrFields, KeyValue1); 
              end else
                Inc(OtherQty);
            end else
            begin
              Inc(InsertQty);
              FindData  := True;
              Sql := TUtilBTPVerdon.GetSqlInsert(lTn, FolderValues, TobL, TMPArrFields, BTPArrFields);
            end;
            if Sql <> '' then
            begin
              TobL := TOB.Create('_QRYL', TobQry, -1);
              TobL.AddChampSupValeur('SqlQry', Sql);
              Sql := '';
            end;
          end;
          if FindData then
          begin
            SqlUnlock := Format('%s)', [SqlUnlock]);
            TobL      := TOB.Create('_QRYL', TobQry, -1);
            TobL.AddChampSupValeur('SqlQry', SqlUnlock);
          end;
          TobQry.AddChampSupValeur('UPDATEQTY', IntToStr(UpdateQty));
          TobQry.AddChampSupValeur('INSERTQTY', IntToStr(InsertQty));
          TobQry.AddChampSupValeur('OTHERQTY', IntToStr(OtherQty));
          TUtilBTPVerdon.InsertUpdateData(lTn, AdoQryL, TobQry, LogValues);  
        finally
          AdoQryL.free;
        end;
      end else
        TUtilBTPVerdon.AddLog(lTn, Format('Aucun tiers n''a été trouvé.', []), LogValues, 1);
*)
    finally
      FreeAndNil(TobT);
    end;
  finally    
//    TUtilBTPVerdon.SetLastSynchro(lTn);
    FreeAndNil(TobQry);
  end;
end;

end.
