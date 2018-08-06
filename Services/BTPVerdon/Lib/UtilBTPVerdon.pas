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
  private
    class function ExtractFieldName(Value : string) : string;
    class function ExtractFielType(Value : string) : string;
  public
    class function GetBTPTableName(Tn : T_TablesName) : string;
    class function GetTMPTableName(Tn : T_TablesName) : string;
    class function GetBTPPrefix(Tn : T_TablesName) : string;
    class function GetTMPPrefix(Tn : T_TablesName) : string;
    class function GetMsgStartEnd(Tn : T_TablesName; Start : boolean; LastSynchro : string) : string;
    class function AddLog(Tn : T_TablesName; Msg : string; LogValues : T_WSLogValues; LineLevel : integer) : string;
    class procedure SetLastSynchro(Tn : T_TablesName);
    class function GetSystemFields(Tn : T_TablesName) : string;
    class function GetFieldsListFromArray(ArrData : array of string) : string;
    class function GetSqlInsertFields(Tn : T_TablesName; TMPArrFields : array of string) : string;
    class function GetSqlInsertValues(Tn : T_TablesName; FolderValues : T_FolderValues; TobData : TOB; BTPArrFields : array of string; TobFieldsType : TOB) : string;
    class function GetSqlUpdate(Tn : T_TablesName; FolderValues : T_FolderValues; TobData : TOB; TMPArrFields, BTPArrFields : array of string; TobFieldsType : TOB) : string;
    class function GetDataSearchSql(Tn: T_TablesName; ArrFieldsList : array of string; LastSynchro : string): string;
    class function GetFieldType(TobFieldsType : TOB; FieldName: string{$IF defined(APPSRV)}; ServerName, DBName : string{$IFEND !APPSRV}): tTypeField;
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

class function TUtilBTPVerdon.ExtractFieldName(Value : string) : string;
begin
  Result := copy(Value, 1, pos(';', Value) -1)
end;

class function TUtilBTPVerdon.ExtractFielType(Value : string) : string;
begin
  Result := copy(Value, pos(';', Value) +1,length(Value));
end;


class function TUtilBTPVerdon.GetBTPTableName(Tn : T_TablesName): string;
begin
  case Tn of
    tnChantier : Result := 'AFFAIRE';
    tnDevis    : Result := 'PIECE';
    tnLignesBR : Result := 'LIGNE';
    tnTiers    : Result := 'TIERS';
  else
    Result := '';
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

class function TUtilBTPVerdon.GetBTPPrefix(Tn: T_TablesName): string;
begin
  case Tn of
    tnChantier : Result := 'AFF';
    tnDevis    : Result := 'GP';
    tnLignesBR : Result := 'GL';
    tnTiers    : Result := 'T';
  else
    Result := '';
  end;
end;

class function TUtilBTPVerdon.GetTMPPrefix(Tn: T_TablesName): string;
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

class function TUtilBTPVerdon.GetMsgStartEnd(Tn : T_TablesName; Start : boolean; LastSynchro : string) : string;
begin
  Result := Format('%s de traitement de la table %s (données créées ou modifiées depuis le %s).', [Tools.iif(Start, 'Début', 'Fin'), TUtilBTPVerdon.GetTMPTableName(Tn), LastSynchro]);
end;

class function TUtilBTPVerdon.AddLog(Tn : T_TablesName; Msg : string; LogValues : T_WSLogValues; LineLevel : integer) : string;
begin
  TServicesLog.WriteLog(ssbylLog, Msg, ServiceName_BTPVerdon, LogValues, LineLevel, True, TUtilBTPVerdon.GetTMPTableName(Tn));
end;

class procedure TUtilBTPVerdon.SetLastSynchro(Tn : T_TablesName);
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

class function TUtilBTPVerdon.GetSystemFields(Tn : T_TablesName) : string;
var
  Prefix : string;
begin
  Prefix := TUtilBTPVerdon.GetTMPPrefix(Tn);
  Result := Format('%s_LOCK;%s_TRAITE;%s_DATETRAITE', [Prefix, Prefix, Prefix]);
end;

class function TUtilBTPVerdon.GetFieldsListFromArray(ArrData: array of string): string;
var
  Cpt    : integer;
begin
  Result := '';
  for Cpt := 0 to High(ArrData) do
    Result := Format('%s, %s', [Result, ExtractFieldName(ArrData[Cpt])]);
  Result := copy(Result, 2, length(Result));
end;

class function TUtilBTPVerdon.GetSqlInsertFields(Tn : T_TablesName; TMPArrFields: array of string): string;
var
  Cpt          : integer;
  SystemFields : string;
begin
  Result := '';
  for Cpt :=  Low(TMPArrFields) to High(TMPArrFields) do
    Result := Format('%s, %s', [Result, ExtractFieldName(TMPArrFields[Cpt])]); 
  SystemFields := GetSystemFields(Tn);
  while SystemFields <> '' do
    Result := Format('%s, %s', [Result, Tools.ReadTokenSt_(SystemFields, ';')]);
  Result := Copy(Result, 2, length(Result));
end;

class function TUtilBTPVerdon.GetSqlInsertValues(Tn : T_TablesName; FolderValues: T_FolderValues; TobData: TOB; BTPArrFields: array of string; TobFieldsType : TOB): string;
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
    FieldType := Tools.GetTypeFieldFromStringType(TUtilBTPVerdon.ExtractFielType(BTPArrFields[Cpt]));
    case FieldType of
      ttfNumeric, ttfInt                     : Result := Format('%s, %s'    , [Result, TobData.GetString(FieldName)]);
      ttfDate                                : Result := Format('%s, ''%s''', [Result, UsDateTime(TobData.GetDateTime(FieldName))]);
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

class function TUtilBTPVerdon.GetSqlUpdate(Tn : T_TablesName; FolderValues : T_FolderValues; TobData : TOB; TMPArrFields, BTPArrFields : array of string; TobFieldsType : TOB) : string;
var
  Cpt          : integer;
  FieldNameTMP : string;
  FieldNameBTP : string;
  Prefix       : string;
  SystemFields : string;
  FieldValue   : string;
  FieldType    : tTypeField;
begin
  for Cpt := 0 to High(TMPArrFields) do
  begin
    FieldNameTMP := ExtractFieldName(TMPArrFields[Cpt]);
    FieldNameBTP := ExtractFieldName(BTPArrFields[Cpt]);
    FieldType    := Tools.GetTypeFieldFromStringType(TUtilBTPVerdon.ExtractFielType(TMPArrFields[Cpt]));
    case FieldType of
      ttfNumeric , ttfInt                    : Result := Format('%s, %s=%s'    , [Result, FieldNameTMP, TobData.GetString(FieldNameBTP)]);
      ttfDate                                : Result := Format('%s, %s=''%s''', [Result, FieldNameTMP, UsDateTime(TobData.GetDateTime(FieldNameBTP))]);
      ttfCombo, ttfText, ttfBoolean, ttfMemo : begin
                                                 FieldValue := TobData.GetString(FieldNameBTP);
                                                 if pos('''', FieldValue) > 0 then
                                                   FieldValue := StringReplace(FieldValue, '''', '''''', [rfReplaceAll]);
                                                 Result := Format('%s, %s=''%s''', [Result, FieldNameTMP, FieldValue]);
                                                end;
    end;
  end;
  Prefix       := TUtilBTPVerdon.GetTMPPrefix(Tn);
  SystemFields := GetSystemFields(Tn);
  while SystemFields <> '' do
  begin
    FieldNameTMP := Tools.ReadTokenSt_(SystemFields, ';');
    case Tools.CaseFromString(FieldNameTMP, [Prefix + '_LOCK', Prefix + '_TRAITE', Prefix + '_DATETRAITE']) of
      {LOCK}       0 : Result := Format('%s, %s=''%s''', [Result, FieldNameTMP, LockDefaultValue]);
      {TRAITE}     1 : Result := Format('%s, %s=''%s''', [Result, FieldNameTMP, TraiteDefaultValue]);
      {DATETRAITE} 2 : Result := Format('%s, %s=''%s''', [Result, FieldNameTMP, DateTraiteDefaultValue]);
    end;
  end;
  Result := Copy(Result, 2, length(Result));
end;
  
class function TUtilBTPVerdon.GetDataSearchSql(Tn: T_TablesName; ArrFieldsList : array of string; LastSynchro : string): string;

  function GetFieldsList : string;
  begin
    Result := TUtilBTPVerdon.GetFieldsListFromArray(ArrFieldsList);
  end;

  function GetLastSynchro : string;
  begin
    Result := Tools.CastDateTimeForQry(StrToDatetime(LastSynchro));
  end;

begin
  case Tn of
    tnTiers    : Result := Format('SELECT %s FROM TIERS WHERE T_DATEMODIF >= "%s" ORDER BY T_AUXILIAIRE'                                       , [GetFieldsList, GetLastSynchro]);
    tnChantier : Result := Format('SELECT %s FROM AFFAIRE WHERE AFF_DATEMODIF >= "%s" ORDER BY AFF_AFFAIRE'                                    , [GetFieldsList, GetLastSynchro]);
    tnDevis    : Result := Format('SELECT %s FROM PIECE WHERE GP_NATUREPIECEG = "DBT" AND GP_DATEMODIF >= "%s" ORDER BY GP_NUMERO'             , [GetFieldsList, GetLastSynchro]);
    tnLignesBR : Result := Format('SELECT %s FROM LIGNE WHERE GL_NATUREPIECEG = "BLF" AND GL_DATEMODIF >= "%s" ORDER BY GL_NUMERO, GL_NUMLIGNE', [GetFieldsList, GetLastSynchro]);
  else
    Result := '';
  end;
end;

class function TUtilBTPVerdon.GetFieldType(TobFieldsType : TOB; FieldName: string{$IF defined(APPSRV)}; ServerName, DBName : string{$IFEND !APPSRV}): tTypeField;
var
  TobL   : TOB;
  sValue : string;
  sType  : string;
begin
  TobL := TobFieldsType.FindFirst(['FIELDNAME'], [FieldName], True);
  if not Assigned(TobL) then
  begin
    sType := Tools.GetStFieldType(FieldName{$IF defined(APPSRV)}, ServerName, DBName{$IFEND !APPSRV});
    TobL  := TOB.Create('FIELDNAME', TobFieldsType, -1);
    TobL.AddChampSupValeur('FIELDNAME', Format('%s;%s', [FieldName, sType]));
  end else
  begin
    sValue := TobL.GetString('FIELDNAME');
    sType  := Copy(sValue, pos(';', sValue) + 1, length(sValue));
  end;
  Result := Tools.GetTypeFieldFromStringType(sType);
end;

end.
