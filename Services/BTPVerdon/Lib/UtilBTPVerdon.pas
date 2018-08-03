unit UtilBTPVerdon;

interface

uses
  ConstServices
  , uTob
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
    class function GetSqlInsertFields(Tn : T_TablesName; TmpArrFields : array of string) : string;
    class function GetSqlInsertValues(Tn : T_TablesName; FolderValues : T_FolderValues; TobData : TOB; BTPArrFields : array of string) : string;
  end;

const
  DBSynchroName = 'VERDON_TAMPON';

implementation

uses
  CommonTools
  , SysUtils
  , IniFiles
  , hCtrls
  , hent1
  ;

{ TUtilBTPVerdon }

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
  SettingFile: TInifile;
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
  Cpt : integer;
begin
  Result := '';
  for Cpt := 0 to High(ArrData) do
    Result := Format('%s, %s', [Result, ArrData[Cpt]]);
  Result := copy(Result, 2, length(Result));
end;

class function TUtilBTPVerdon.GetSqlInsertFields(Tn : T_TablesName; TMPArrFields: array of string): string;
var
  Cpt          : integer;
  SystemFields : string;
begin
  for Cpt :=  Low(TMPArrFields) to High(TMPArrFields) do
    Result := Format('%s, %s', [Result, TMPArrFields[Cpt]]);
  SystemFields := GetSystemFields(Tn);
  while SystemFields <> '' do
    Result := Format('%s, %s', [Result, Tools.ReadTokenSt_(SystemFields, ';')]);
  Result := Copy(Result, 2, length(Result));
end;

class function TUtilBTPVerdon.GetSqlInsertValues(Tn : T_TablesName; FolderValues: T_FolderValues; TobData: TOB; BTPArrFields: array of string): string;
var
  Cpt       : integer;
  FieldName : string;
  FieldType : tTypeField;
begin
  for Cpt := 0 to High(BTPArrFields) do
  begin
    FieldName := BTPArrFields[Cpt];
    FieldType := Tools.GetFieldType(FieldName{$IF defined(APPSRV)}, FolderValues.BTPServer, FolderValues.BTPDataBase {$IFEND !APPSRV});
    case FieldType of
      ttfNumeric
      , ttfInt     : Result := Format('%s, %s'    , [Result, TobData.GetString(FieldName)]);
      ttfDate      : Result := Format('%s, ''%s''', [Result, UsDateTime(TobData.GetDateTime(FieldName))]);
      ttfCombo
      , ttfText
      , ttfBoolean
      , ttfMemo    : Result := Format('%s, ''%s''', [Result, TobData.GetString(FieldName)]);
    end;
  end;
  Result := Format('%s, ''O'', ''N'', ''%s'' ', [Result, DateTimeToStr(iDate1900)]);
  Result := Copy(Result, 2, length(Result));
end;

end.
