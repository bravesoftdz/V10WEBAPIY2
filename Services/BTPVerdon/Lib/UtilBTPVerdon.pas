unit UtilBTPVerdon;

interface

uses
  ConstServices
  , uTob
  , CommonTools
  ;

type

  T_TablesName = (tnNone, tnChantier, tnDevis, tnLignesBR, tnTiers);

  T_TiersValues    = Record
                       FirstExec   : boolean;
                       Count       : integer;
                       TimeOut     : integer;
                       LastSynchro : string;
                     end;

  T_ChantierValues = Record
                       FirstExec   : boolean;
                       Count       : integer;
                       TimeOut     : integer;
                       LastSynchro : string;
                     end;

  T_DevisValues    = Record
                       FirstExec   : boolean;
                       Count       : integer;
                       TimeOut     : integer;
                       LastSynchro : string;
                     end;

  T_LignesBRValues = Record
                       FirstExec   : boolean;
                       Count       : integer;
                       TimeOut     : integer;
                       LastSynchro : string;
                     end;

  T_FolderValues   = Record
                       BTPConnectionName : string;
                       BTPUserAdmin      : string;
                       BTPServer         : string;
                       BTPDataBase       : string;
                       TMPServer         : string;
                       TMPDataBase       : string;
                     end;

  TUtilBTPVerdon = class (TObject)
    class function GetTMPTableName(Tn : T_TablesName) : string;
    class function GetMsgStartEnd(Tn : T_TablesName; Start : boolean; LastSynchro : string) : string;
    class function AddLog(Tn : T_TablesName; Msg : string; LogValues : T_WSLogValues; LineLevel : integer) : string;
  end;

  TTnTreatment = class (TObject)
  private
    BTPArrFields  : array of string;
    TMPArrFields  : array of string;

    function GetTMPPrefix : string;
    function ExtractFieldName(Value : string) : string;
    function ExtractFielType(Value : string) : string;
    procedure SetFieldsArray;
    function GetSqlUnlock : string;
    function GetValueKey(TobData : TOB) : string;
    function GetSqlDataExist(FieldsList, KeyValue : string) : string;
    function GetSystemFields : string;
    function GetFieldsListFromArray(ArrData : array of string) : string;
    function GetSqlInsertFields : string;
    function GetSqlInsertValues(TobData : TOB) : string;
    function GetSqlUpdate(TobData : TOB; KeyValue1 : string) : string;
    function GetSqlInsert(TobData : TOB) : string;
    function GetDataSearchSql : string;
    function InsertUpdateData(AdoQryL : AdoQry; TobData: TOB): boolean;
    function GetFieldsList : string;
    procedure SetLastSynchro;

  public
    Tn           : T_TablesName;
    LogValues    : T_WSLogValues;
    FolderValues : T_FolderValues;
    LastSynchro  : string;
    function TnTreatment(TobTable, TobQry: TOB; AdoQryL : AdoQry): boolean;
  end;

const
  DBSynchroName          = 'VERDON_TAMPON';
  LockDefaultValue       = 'O';
  TraiteDefaultValue     = 'N';
  DateTraiteDefaultValue = '19000101';

implementation

uses
  SysUtils
  , hEnt1
  , IniFiles
  , hCtrls
  ;

{ TUtilBTPVerdon }

function TTnTreatment.ExtractFieldName(Value : string) : string;
begin
  Result := copy(Value, 1, pos(';', Value) -1)
end;

function TTnTreatment.ExtractFielType(Value : string) : string;
begin
  Result := copy(Value, pos(';', Value) +1,length(Value));
end;

procedure TTnTreatment.SetFieldsArray;
var
  Cpt    : integer;
  ArrLen : integer;

  procedure AddValues(SuffixBTP, SuffixTMP : string);
  var
    FieldType    : string;
    BTPFieldName : string;
    TMPFieldName : string;
  begin
    BTPFieldName := Format('T_%s'   , [SuffixBTP]);
    TMPFieldName := Format('TIE_%s' , [SuffixTMP]);
    if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(Tn, Format('%sThreadTiers.SetFieldsArray / AddValues / BTPFieldName=%s, TMPFieldName=%s', [WSCDS_DebugMsg, BTPFieldName, TMPFieldName]), LogValues, 0);
    FieldType := Tools.GetStFieldType(BTPFieldName{$IFDEF APPSRV}, FolderValues.BTPServer, FolderValues.BTPDataBase, LogValues.DebugEvents {$ENDIF APPSRV});
    if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(Tn, Format('%sThreadTiers.SetFieldsArray / AddValues / FieldType=%s', [WSCDS_DebugMsg, FieldType]), LogValues, 0);
    BTPArrFields[Cpt] := Format('%s;%s', [BTPFieldName, FieldType]);
    TMPArrFields[Cpt] := Format('%s;%s', [TMPFieldName, FieldType]);
  end;

begin
  if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(Tn, Format('%sThreadTiers.SetFieldsArray / Srv=%s, Folder=%s', [WSCDS_DebugMsg, FolderValues.BTPServer, FolderValues.BTPDataBase]), LogValues, 0);
  case Tn of
    tnChantier : ArrLen := 0;
    tnDevis    : ArrLen := 0;
    tnLignesBR : ArrLen := 0;
    tnTiers    : ArrLen := 25;
  else
    Arrlen := 0;
  end;
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

class function TUtilBTPVerdon.GetTMPTableName(Tn : T_TablesName): string;
begin
  case Tn of
    tnChantier : Result := 'CHANTIER';
    tnDevis    : Result := 'DEVIS';
    tnLignesBR : Result := 'LIGNESBR';
    tnTiers    : Result := 'TIERS';
  else
    Result := '';
  end;
end;

function TTnTreatment.GetTMPPrefix : string;
begin
  case Tn of
    tnChantier : Result := 'CHA';
    tnDevis    : Result := 'DEV';
    tnLignesBR : Result := 'LBR';
    tnTiers    : Result := 'TIE';
  else
    Result := '';
  end;
end;

function TTnTreatment.GetFieldsList : string;
begin
  case Tn of
    tnChantier : Result := '';
    tnDevis    : Result := '';
    tnLignesBR : Result := '';
    tnTiers    : Result := 'TIE_LOCK,TIE_TRAITE';
  else
    Result := '';
  end;
end;

function TTnTreatment.GetSqlUnlock : string;
begin
  case Tn of
    tnChantier : Result := '';
    tnDevis    : Result := '';
    tnLignesBR : Result := '';
    tnTiers    : Result := 'UPDATE TIERS SET TIE_LOCK = ''N'' WHERE TIE_AUXILIAIRE IN(';
  else
    Result := '';
  end;
end;

function TTnTreatment.GetValueKey(TobData : TOB) : string;
begin
  case Tn of
    tnChantier : Result := '';
    tnDevis    : Result := '';
    tnLignesBR : Result := '';
    tnTiers    : Result := format('%s;%s', [TobData.GetString('T_AUXILIAIRE'), TobData.GetString('T_LIBELLE')]);
  else
    Result := '';
  end;
end;

function TTnTreatment.GetSqlDataExist(FieldsList, KeyValue : string) : string;
begin
  case Tn of
    tnChantier : Result := '';
    tnDevis    : Result := '';
    tnLignesBR : Result := '';
    tnTiers    : Result := Format('SELECT %s FROM TIERS WHERE TIE_AUXILIAIRE = ''%s''', [FieldsList, KeyValue]);
  else
    Result := '';
  end;
end;


class function TUtilBTPVerdon.GetMsgStartEnd(Tn : T_TablesName; Start : boolean; LastSynchro : string) : string;
begin
  Result := Format('%s de traitement de la table %s (données créées ou modifiées depuis le %s).', [Tools.iif(Start, 'Début', 'Fin'), TUtilBTPVerdon.GetTMPTableName(Tn), LastSynchro]);
end;

class function TUtilBTPVerdon.AddLog(Tn : T_TablesName; Msg : string; LogValues : T_WSLogValues; LineLevel : integer) : string;
begin
  TServicesLog.WriteLog(ssbylLog, Msg, ServiceName_BTPVerdon, LogValues, LineLevel, True, TUtilBTPVerdon.GetTMPTableName(Tn));
end;

procedure TTnTreatment.SetLastSynchro;
var
  SettingFile : TInifile;
  IniFilePath : string;
begin

//***************************
exit;
//***************************

  IniFilePath := TServicesLog.GetFilePath(ServiceName_BTPVerdon, 'ini');
  SettingFile := TIniFile.Create(IniFilePath);
  try
    SettingFile.WriteString('TABLESLASTSYNCHRO', TUtilBTPVerdon.GetTMPTableName(Tn), DateTimeToStr(Now));
    SettingFile.UpdateFile;
  finally
    SettingFile.Free;
  end;
end;

function TTnTreatment.GetSystemFields : string;
var
  Prefix : string;
begin
  Prefix := GetTMPPrefix;
  Result := Format('%s_LOCK;%s_TRAITE;%s_DATETRAITE', [Prefix, Prefix, Prefix]);
end;

function TTnTreatment.GetFieldsListFromArray(ArrData: array of string): string;
var
  Cpt    : integer;
begin
  Result := '';
  for Cpt := 0 to High(ArrData) do
    Result := Format('%s, %s', [Result, ExtractFieldName(ArrData[Cpt])]);
  Result := copy(Result, 2, length(Result));
end;

function TTnTreatment.GetSqlInsertFields : string;
var
  Cpt          : integer;
  SystemFields : string;
begin
  Result := '';
  for Cpt :=  Low(TMPArrFields) to High(TMPArrFields) do
    Result := Format('%s, %s', [Result, ExtractFieldName(TMPArrFields[Cpt])]); 
  SystemFields := GetSystemFields;
  while SystemFields <> '' do
    Result := Format('%s, %s', [Result, Tools.ReadTokenSt_(SystemFields, ';')]);
  Result := Copy(Result, 2, length(Result));
end;

function TTnTreatment.GetSqlInsertValues(TobData: TOB): string;
var
  Cpt        : integer;
  FieldName  : string;
  FieldValue : string;
  FieldType  : tTypeField;
begin
  Result := '';
  for Cpt := 0 to High(BTPArrFields) do
  begin
    FieldName := ExtractFieldName(BTPArrFields[Cpt]);
    FieldType := Tools.GetTypeFieldFromStringType(ExtractFielType(BTPArrFields[Cpt]));
    case FieldType of
      ttfNumeric, ttfInt                     : Result := Format('%s, %s'    , [Result, TobData.GetString(FieldName)]);
      ttfDate                                : Result := Format('%s, ''%s''', [Result, Tools.UsDateTime_(TobData.GetDateTime(FieldName))]);
      ttfCombo, ttfText, ttfBoolean, ttfMemo : begin
                                                 FieldValue := TobData.GetString(FieldName);
                                                 if pos('''', FieldValue) > 0 then
                                                   FieldValue := StringReplace(FieldValue, '''', '''''', [rfReplaceAll]);
                                                 Result := Format('%s, ''%s''', [Result, FieldValue]);
                                                end;
    end;
  end;
  Result := Format('%s, ''%s'', ''%s'', ''%s''', [Result, LockDefaultValue, TraiteDefaultValue, DateTraiteDefaultValue]);
  Result := Copy(Result, 2, length(Result));
end;

function TTnTreatment.GetSqlUpdate(TobData : TOB; KeyValue1 : string) : string;
var
  Cpt          : integer;
  FieldNameTMP : string;
  FieldNameBTP : string;
  Sql          : string;
  Prefix       : string;
  SystemFields : string;
  FieldValue   : string;
  FieldType    : tTypeField;
begin
  for Cpt := 0 to High(TMPArrFields) do
  begin
    FieldNameTMP := ExtractFieldName(TMPArrFields[Cpt]);
    FieldNameBTP := ExtractFieldName(BTPArrFields[Cpt]);
    FieldType    := Tools.GetTypeFieldFromStringType(ExtractFielType(TMPArrFields[Cpt]));
    case FieldType of
      ttfNumeric , ttfInt                    : Sql := Format('%s, %s=%s'    , [Sql, FieldNameTMP, TobData.GetString(FieldNameBTP)]);
      ttfDate                                : Sql := Format('%s, %s=''%s''', [Sql, FieldNameTMP, Tools.UsDateTime_(TobData.GetDateTime(FieldNameBTP))]);
      ttfCombo, ttfText, ttfBoolean, ttfMemo : begin
                                                 FieldValue := TobData.GetString(FieldNameBTP);
                                                 if pos('''', FieldValue) > 0 then
                                                   FieldValue := StringReplace(FieldValue, '''', '''''', [rfReplaceAll]);
                                                 Sql := Format('%s, %s=''%s''', [Sql, FieldNameTMP, FieldValue]);
                                                end;
    end;
  end;
  Prefix       := GetTMPPrefix;
  SystemFields := GetSystemFields;
  while SystemFields <> '' do
  begin
    FieldNameTMP := Tools.ReadTokenSt_(SystemFields, ';');
    case Tools.CaseFromString(FieldNameTMP, [Prefix + '_LOCK', Prefix + '_TRAITE', Prefix + '_DATETRAITE']) of
      {LOCK}       0 : Sql := Format('%s, %s=''%s''', [Sql, FieldNameTMP, LockDefaultValue]);
      {TRAITE}     1 : Sql := Format('%s, %s=''%s''', [Sql, FieldNameTMP, TraiteDefaultValue]);
      {DATETRAITE} 2 : Sql := Format('%s, %s=''%s''', [Sql, FieldNameTMP, DateTraiteDefaultValue]);
    end;
  end;
  Sql := Copy(Sql, 2, length(Sql));
  case Tn of
    tnChantier : Result := '';
    tnDevis    : Result := '';
    tnLignesBR : Result := '';
    tnTiers    : Result := Format('UPDATE TIERS SET %s WHERE TIE_AUXILIAIRE = ''%s''', [Sql, KeyValue1]);
  else
    Result := '';
  end;
end;

function TTnTreatment.GetSqlInsert(TobData : TOB) : string;
var
  Fields : string;
  Values : string;
begin
  Fields := GetSqlInsertFields;
  Values := GetSqlInsertValues(TobData);
  case Tn of
    tnChantier : Result := '';
    tnDevis    : Result := '';
    tnLignesBR : Result := '';
    tnTiers    : Result := Format('INSERT INTO TIERS (%s) VALUES (%s)', [Fields, Values]);
  else
    Result := '';
  end;
end;

function TTnTreatment.GetDataSearchSql : string;

  function GetLastSynchro : string;
  begin
    Result := Tools.CastDateTimeForQry(StrToDatetime(LastSynchro));
  end;

begin
  case Tn of
    tnTiers    : Result := Format('SELECT %s FROM TIERS WHERE T_DATEMODIF >= "%s" ORDER BY T_AUXILIAIRE'                                       , [GetFieldsListFromArray(BTPArrFields), GetLastSynchro]);
    tnChantier : Result := Format('SELECT %s FROM AFFAIRE WHERE AFF_DATEMODIF >= "%s" ORDER BY AFF_AFFAIRE'                                    , [GetFieldsListFromArray(BTPArrFields), GetLastSynchro]);
    tnDevis    : Result := Format('SELECT %s FROM PIECE WHERE GP_NATUREPIECEG = "DBT" AND GP_DATEMODIF >= "%s" ORDER BY GP_NUMERO'             , [GetFieldsListFromArray(BTPArrFields), GetLastSynchro]);
    tnLignesBR : Result := Format('SELECT %s FROM LIGNE WHERE GL_NATUREPIECEG = "BLF" AND GL_DATEMODIF >= "%s" ORDER BY GL_NUMERO, GL_NUMLIGNE', [GetFieldsListFromArray(BTPArrFields), GetLastSynchro]);
  else
    Result := '';
  end;
end;

function TTnTreatment.InsertUpdateData(AdoQryL : AdoQry; TobData: TOB): boolean;
var
  Cpt       : integer;
  UpdateQty : integer;
  InsertQty : integer;
  OtherQty  : integer;
begin
  Result := True;
  if (assigned(TobData)) then
  begin
    UpdateQty := TobData.GetInteger('UPDATEQTY');
    InsertQty := TobData.GetInteger('INSERTQTY');
    OtherQty  := TobData.GetInteger('OTHERQTY');
    try
      BeginTrans;
      for Cpt := 0 to pred(TobData.Detail.Count) do
      begin
        AdoQryL.RecordCount := 0;
        AdoQryL.Request     := TobData.Detail[Cpt].GetString('SqlQry');
        if LogValues.DebugEvents = 2 then TUtilBTPVerdon.AddLog(Tn, Format('%sExécution de %s ', [WSCDS_DebugMsg, AdoQryL.Request]), LogValues, 1);
        AdoQryL.InsertUpdate;
      end;
      CommitTrans;
      if UpdateQty > 0 then
        TUtilBTPVerdon.AddLog(Tn, Format('%s enregistrements de la table %s modifié(s)', [IntToStr(UpdateQty), TUtilBTPVerdon.GetTMPTableName(Tn)]), LogValues, 1);
      if InsertQty > 0 then
        TUtilBTPVerdon.AddLog(Tn, Format('%s enregistrements de la table %s créé(s)', [IntToStr(InsertQty), TUtilBTPVerdon.GetTMPTableName(Tn)]), LogValues, 1);
      if OtherQty > 0 then
        TUtilBTPVerdon.AddLog(Tn, Format('%s enregistrements de la table %s non traité(s) car verrouillé(s)', [IntToStr(OtherQty), TUtilBTPVerdon.GetTMPTableName(Tn)]), LogValues, 1);
      SetLastSynchro;
    except
      Result := False;
      Rollback;
    end;
  end;
end;

function TTnTreatment.TnTreatment(TobTable, TobQry: TOB; AdoQryL : AdoQry): boolean;
var
  TobL      : TOB;
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
begin
  Result    := True;
  InsertQty := 0;
  UpdateQty := 0;
  OtherQty  := 0;
  SetFieldsArray;
  Sql := GetDataSearchSql;
  if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(Tn, Format('%sSql recherche Tiers = %s', [WSCDS_DebugMsg, Sql]), LogValues, 0);
  TobTable.LoadDetailFromSQL(Sql);
  Sql := '';
  if TobTable.Detail.Count > 0 then
  begin
    TUtilBTPVerdon.AddLog(Tn, Format('Recherche des données (%s enregistrement(s) de la table %s)', [IntToStr(TobTable.Detail.Count), TUtilBTPVerdon.GetTMPTableName(Tn)]), LogValues, 1);
    AdoQryL.FieldsList := GetFieldsList;
    AdoQryL.LogValues  := LogValues;
    AdoQryL.TSLResult.Clear;
    SqlUnlock := GetSqlUnlock;
    FindData  := False;
    for Cpt := 0 to pred(TobTable.Detail.Count) do
    begin
      TobL      := TobTable.Detail[Cpt];
      KeyValues := GetValueKey(TobL);
      KeyValue1 := Tools.ReadTokenSt_(KeyValues, ';');
      KeyValue2 := Tools.ReadTokenSt_(KeyValues, ';');
      SqlUnlock := Format('%s%s''%s''', [SqlUnlock, Tools.iif(Cpt = 0, '', ', '), KeyValue1]); // Prépare update de Unlock
      AdoQryL.Request := GetSqlDataExist(AdoQryL.FieldsList, KeyValue1); // Test si enregistrement existe
      AdoQryL.SingleTableSelect;
      if LogValues.LogLevel = 2 then TUtilBTPVerdon.AddLog(Tn, Format('%s - %s (%s)', [KeyValue1, KeyValue2 , Tools.iif(AdoQryL.RecordCount = 1, 'à modifier', 'à créer')]), LogValues, 3);
      if AdoQryL.RecordCount = 1 then // Update
      begin
        Values := AdoQryL.TSLResult[0];
        Lock   := Tools.ReadTokenSt_(Values, ToolsTobToTsl_Separator);
        Treat  := Tools.ReadTokenSt_(Values, ToolsTobToTsl_Separator);
        if Lock = 'N' then
        begin
          inc(UpdateQty);
          FindData  := True;
          Sql := GetSqlUpdate(TobL, KeyValue1);
        end else
          Inc(OtherQty);
      end else
      begin
        Inc(InsertQty);
        FindData  := True;
        Sql := GetSqlInsert(TobL);
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
    InsertUpdateData(AdoQryL, TobQry);
  end else
    TUtilBTPVerdon.AddLog(Tn, Format('Aucun tiers n''a été trouvé.', []), LogValues, 1);
end;

end.
