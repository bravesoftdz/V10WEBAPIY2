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
    BTPArrFields  : array of string;
    TMPArrFields  : array of string;
    TobFieldsType : TOB;

//    procedure SetName;
    procedure SetFieldsArray;

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

const
  ArrLen = 25;

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

procedure ThreadTiers.SetFieldsArray;
var
  Cpt : integer;

  procedure AddValues(SuffixBTP, SuffixTMP : string);
  var
    FieldType    : string;
    BTPFieldName : string;
    TMPFieldName : string;
  begin
    BTPFieldName := Format('T_%s'   , [SuffixBTP]);
    TMPFieldName := Format('TIE_%s' , [SuffixTMP]);
    if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(lTn, Format('%sThreadTiers.SetFieldsArray / AddValues / BTPFieldName=%s, TMPFieldName=%s', [WSCDS_DebugMsg, BTPFieldName, TMPFieldName]), LogValues, 0);
    FieldType := Tools.GetStFieldType(BTPFieldName{$IFDEF APPSRV}, FolderValues.BTPServer, FolderValues.BTPDataBase, LogValues.DebugEvents {$ENDIF APPSRV});
    if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(lTn, Format('%sThreadTiers.SetFieldsArray / AddValues / FieldType=%s', [WSCDS_DebugMsg, FieldType]), LogValues, 0);
    BTPArrFields[Cpt] := Format('%s;%s', [BTPFieldName, FieldType]);
    TMPArrFields[Cpt] := Format('%s;%s', [TMPFieldName, FieldType]);
  end;

begin
  if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(lTn, Format('%sThreadTiers.SetFieldsArray / Srv=%s, Folder=%s', [WSCDS_DebugMsg, FolderValues.BTPServer, FolderValues.BTPDataBase]), LogValues, 0);
  SetLength(BTPArrFields, ArrLen);
  SetLength(TMPArrFields, ArrLen);
  for Cpt :=  Low(BTPArrFields) to High(BTPArrFields) do
  begin
    case Cpt of
      0  : AddValues('AUXILIAIRE'  , 'AUXILIAIRE');
      1  : AddValues('COLLECTIF'   , 'COLLECTIF');
      2  : AddValues('NATUREAUXI'  , 'NATUREAUXI');
      3  : AddValues('LIBELLE'     , 'LIBELLE');
      4  : AddValues('DEVISE'      , 'DEVISE');
      5  : AddValues('ADRESSE1'    , 'ADRESSE1');
      6  : AddValues('ADRESSE2'    , 'ADRESSE2');
      7  : AddValues('ADRESSE3'    , 'ADRESSE3');
      8  : AddValues('CODEPOSTAL'  , 'CP');
      9  : AddValues('VILLE'       , 'VILLE');
      10 : AddValues('PAYS'        , 'PAYS');
      11 : AddValues('TELEPHONE'   , 'TELEPHONE');
      12 : AddValues('FAX'         , 'TELEPHONE2');
      13 : AddValues('TELEX'       , 'TELEPHONE3');
      14 : AddValues('EMAIL'       , 'EMAIL');
      15 : AddValues('RVA'         , 'WEBURL');
      16 : AddValues('COMPTATIERS' , 'COMPTA');
      17 : AddValues('FACTURE'     , 'TIERSFACTURE');
      18 : AddValues('PAYEUR'      , 'TIERSPAYEUR');
      19 : AddValues('BLOCNOTE'    , 'BLOCNOTE');
      20 : AddValues('SIRET'       , 'SIRET');
      21 : AddValues('NIF'         , 'NUMINTRACOMM');
      22 : AddValues('SOCIETE'     , 'CODESOCIETE');
      23 : AddValues('DATEMODIF'   , 'DATEMODIF');
      24 : AddValues('DATECREATION', 'DATECREATION');
    end;
  end;
end;

procedure ThreadTiers.Execute;
var
  TobT      : TOB;
  TobQry    : TOB;
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
  Auxiliary : string;
  Values    : string;
  FindData  : boolean;
begin
//  SetName;
  TUtilBTPVerdon.AddLog(lTn, '', LogValues, 0);
  TUtilBTPVerdon.AddLog(lTn, DupeString('*', 50), LogValues, 0);
  TobFieldsType := TOB.Create('_TYPEFIELDS', nil, -1);
  try
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
      if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(lTn, Format('%sTOB.Create(_QRY, nil, -1)', [WSCDS_DebugMsg]), LogValues, 0);
      TobT := TOB.Create('_TIERS', nil, -1);
      try
        if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(lTn, Format('%sTOB.Create(_TIERS, nil, -1)', [WSCDS_DebugMsg]), LogValues, 0);
        InsertQty := 0;
        UpdateQty := 0;
        OtherQty  := 0;
        if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(lTn, Format('%sBefore SetFieldsArray', [WSCDS_DebugMsg]), LogValues, 0);
        SetFieldsArray;
        if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(lTn, Format('%sAfter SetFieldsArray', [WSCDS_DebugMsg]), LogValues, 0);
        Sql := TUtilBTPVerdon.GetDataSearchSql(tnTiers, BTPArrFields, TiersValues.LastSynchro);
        if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(lTn, Format('%sSql : %s', [WSCDS_DebugMsg, Sql]), LogValues, 0);
        TobT.LoadDetailFromSQL(Sql);
        if TobT.Detail.Count > 0 then
        begin
          TUtilBTPVerdon.AddLog(lTn, Format('Recherche des données pour %s tiers', [IntToStr(TobT.Detail.Count)]), LogValues, 1);
          AdoQryL := AdoQry.Create;
          try
            AdoQryL.ServerName  := FolderValues.TMPServer;
            AdoQryL.DBName      := FolderValues.TMPDataBase;
            AdoQryL.PgiDB       := '-';
            AdoQryL.FieldsList  := 'TIE_LOCK,TIE_TRAITE';
            AdoQryL.LogValues   := LogValues;
            AdoQryL.TSLResult.Clear;
            SqlUnlock := 'UPDATE TIERS SET TIE_LOCK = ''N'' WHERE TIE_AUXILIAIRE IN(';
            FindData  := False;
            for Cpt := 0 to pred(TobT.Detail.Count) do
            begin
              TobL      := TobT.Detail[Cpt];
              Auxiliary := TobL.GetString('T_AUXILIAIRE');
              // Prépare update de Unlock
              SqlUnlock := Format('%s%s''%s''', [SqlUnlock, Tools.iif(Cpt = 0, '', ', '), Auxiliary]);
              // Test si enregistrement existe
              AdoQryL.Request := Format('SELECT %s FROM TIERS WHERE TIE_AUXILIAIRE = ''%s''', [AdoQryL.FieldsList, Auxiliary]);
              AdoQryL.SingleTableSelect;
              if LogValues.LogLevel = 2 then TUtilBTPVerdon.AddLog(lTn, Format('%s - %s (%s)', [Auxiliary, TobL.GetString('T_LIBELLE'), Tools.iif(AdoQryL.RecordCount = 1, 'modifié', 'ajouté')]), LogValues, 3);
              if AdoQryL.RecordCount = 1 then // Update
              begin
                Values := AdoQryL.TSLResult[0];
                Lock  := Tools.ReadTokenSt_(Values, ToolsTobToTsl_Separator);
                Treat := Tools.ReadTokenSt_(Values, ToolsTobToTsl_Separator);
                if Lock = 'N' then
                begin
                  inc(UpdateQty);
                  FindData  := True;
                  Sql := Format('UPDATE TIERS SET %s WHERE TIE_AUXILIAIRE = ''%s''', [TUtilBTPVerdon.GetSqlUpdate(tnTiers, FolderValues, TobL, TMPArrFields, BTPArrFields, TobFieldsType), Auxiliary]);
                end else
                  Inc(OtherQty);
              end else
              begin
                Inc(InsertQty);
                FindData  := True;
                Sql       := Format('INSERT INTO TIERS (%s) VALUES (%s)'
                                    , [  TUtilBTPVerdon.GetSqlInsertFields(tnTiers, TMPArrFields)
                                       , TUtilBTPVerdon.GetSqlInsertValues(tnTiers, FolderValues, TobL, BTPArrFields, TobFieldsType)
                                      ]);
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
            try
              BeginTrans;
              for Cpt := 0 to pred(TobQry.Detail.Count) do
              begin
                AdoQryL.RecordCount := 0;
                AdoQryL.Request     := TobQry.Detail[Cpt].GetString('SqlQry');
//*************************
                AdoQryL.InsertUpdate;
//*************************
              end;
              CommitTrans;
              if UpdateQty > 0 then
                TUtilBTPVerdon.AddLog(lTn, Format('%s tiers modifié(s)', [IntToStr(UpdateQty)]), LogValues, 1);
              if OtherQty > 0 then
                TUtilBTPVerdon.AddLog(lTn, Format('%s tiers non traité(s) car verrouillé(s)', [IntToStr(OtherQty)]), LogValues, 1);
              if InsertQty > 0 then
                TUtilBTPVerdon.AddLog(lTn, Format('%s tiers ajouté(s)', [IntToStr(InsertQty)]), LogValues, 1);
            except
              Rollback;
            end;
          finally
            AdoQryL.free;
          end;
        end else
          TUtilBTPVerdon.AddLog(lTn, Format('Aucun tiers n''a été trouvé.', []), LogValues, 1);
      finally
        FreeAndNil(TobT);
      end;
    finally
    
      TUtilBTPVerdon.SetLastSynchro(lTn);
      FreeAndNil(TobQry);
    end;
  finally
    FreeAndNil(TobFieldsType);
  end;
end;

end.
