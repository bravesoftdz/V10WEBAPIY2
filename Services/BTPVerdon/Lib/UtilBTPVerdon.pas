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
    BTPArrFields     : array of string;
    TMPArrFields     : array of string;
    BTPArrAdditionalFields : array of string;
    TMPArrAdditionalFields : array of string;

    function GetPrefix : string;
    function GetTMPFieldName(BTPFieldName : string) : string;
    function GetTMPIndexFieldName : string;
    function GetBTPIndexFieldName : string;
    function GetBTPLabelFieldName : string;
    function ExtractFieldName(Value : string) : string;
    function ExtractFieldType(Value : string) : string;
    //function SetFieldsArray(BTPArrFields, TMPArrFields, BTPArrAdditionalFields,TMPArrAdditionalFields : array of string) : boolean;
    function SetFieldsArray : boolean;
    function GetSqlDataExist(FieldsList, KeyValue1, KeyValue2 : string) : string;
    function GetSystemFields : string;
    function GetFieldsListFromArray(ArrData : array of string) : string;
    function GetValue(FieldNameBTP, FieldNameTMP : string; FieldType : tTypeField; TobData : TOB) : string;
    function GetSqlInsertFields : string;
    function GetSqlInsertAdditionalFields : string;
    function GetSqlInsertValues(TobData : TOB; IsAdditional : boolean=False) : string;
    function GetSqlUpdate(TobData, TobAdd : TOB; KeyValue1, KeyValue2 : string) : string;
    function GetSqlInsert(TobData, TobAdd : TOB) : string;
    function GetDataSearchSql : string;
    function GetTMPFieldSizeMax(FieldName : string) : integer;
    function InsertUpdateData(AdoQryL : AdoQry; TobData: TOB): boolean;
    procedure SetLinkedRecords(TobAdd, TobData : TOB);
    procedure SetLastSynchro;

  public
    Tn           : T_TablesName;
    LogValues    : T_WSLogValues;
    FolderValues : T_FolderValues;
    LastSynchro  : string;
    //function TnTreatment(TobTable, TobAdd, TobQry: TOB; AdoQryL : AdoQry; BTPArrFields, TMPArrFields, BTPArrAdditionalFields, TMPArrAdditionalFields : array of string): boolean;
    function TnTreatment(TobTable, TobAdd, TobQry: TOB; AdoQryL : AdoQry): boolean;
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
  , dpOutils
  {$IFNDEF DBXPRESS}
  , dbTables
  {$ELSE !DBXPRESS}
  , uDbxDataSet
  {$ENDIF !DBXPRESS}
  ;

{ TUtilBTPVerdon }

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

class function TUtilBTPVerdon.GetMsgStartEnd(Tn : T_TablesName; Start : boolean; LastSynchro : string) : string;
begin
  Result := Format('%s de traitement de la table %s (données créées ou modifiées depuis le %s).', [Tools.iif(Start, 'Début', 'Fin'), TUtilBTPVerdon.GetTMPTableName(Tn), LastSynchro]);
end;

class function TUtilBTPVerdon.AddLog(Tn : T_TablesName; Msg : string; LogValues : T_WSLogValues; LineLevel : integer) : string;
begin
  TServicesLog.WriteLog(ssbylLog, Msg, ServiceName_BTPVerdon, LogValues, LineLevel, True, TUtilBTPVerdon.GetTMPTableName(Tn));
end;

{ TTnTreatment }

function TTnTreatment.ExtractFieldName(Value : string) : string;
begin
  Result := copy(Value, 1, pos(';', Value) -1)
end;

function TTnTreatment.ExtractFieldType(Value : string) : string;
begin
  Result := copy(Value, pos(';', Value) +1,length(Value));
end;

//function TTnTreatment.SetFieldsArray(BTPArrFields, TMPArrFields, BTPArrAdditionalFields,TMPArrAdditionalFields : array of string) : boolean;
function TTnTreatment.SetFieldsArray : boolean;
var
  Cpt       : integer;
  ArrLen    : integer;

  procedure AddValues(BTPFieldName, TMPFieldName : string; IsAdditional : boolean=False);
  var
    FieldType : string;
  begin
    if not IsAdditional then
      FieldType := Tools.GetStFieldType(BTPFieldName{$IFDEF APPSRV}, FolderValues.BTPServer, FolderValues.BTPDataBase, LogValues.DebugEvents {$ENDIF APPSRV})
    else
      FieldType := 'VARCHAR(100)';
    if (LogValues.DebugEvents = 2) then TUtilBTPVerdon.AddLog(Tn, Format('%sSetFieldsArray / AddValues / BTPFieldName=%s, TMPFieldName=%s, FieldType=%s', [WSCDS_DebugMsg, BTPFieldName, TMPFieldName, FieldType]), LogValues, 0);
    if not IsAdditional then
    begin
      BTPArrFields[Cpt] := Format('%s;%s', [BTPFieldName, FieldType]);
      TMPArrFields[Cpt] := Format('%s;%s', [TMPFieldName, FieldType]);
    end else
    begin
      BTPArrAdditionalFields[Cpt] := Format('%s;%s', [BTPFieldName, FieldType]);
      TMPArrAdditionalFields[Cpt] := Format('%s;%s', [TMPFieldName, FieldType]);
    end;
  end;

begin
  Result := True;
  if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(Tn, Format('%sThreadTiers.SetFieldsArray / Srv=%s, Folder=%s', [WSCDS_DebugMsg, FolderValues.BTPServer, FolderValues.BTPDataBase]), LogValues, 0);
  case Tn of
    tnChantier : ArrLen := 11;
    tnDevis    : ArrLen := 12;
    tnLignesBR : ArrLen := 12;
    tnTiers    : ArrLen := 25;
  else
    Arrlen := 0;
  end;
  SetLength(BTPArrFields, ArrLen);
  SetLength(TMPArrFields, ArrLen);
  case Tn of
    tnChantier :
      begin
        for Cpt := Low(BTPArrFields) to High(BTPArrFields) do
        begin
          case Cpt of
            0  : AddValues('AFF_AFFAIRE'      , 'CHA_CODE');
            1  : AddValues('AFF_LIBELLE'      , 'CHA_LIBELLE');
            2  : AddValues('AFF_DESCRIPTIF'   , 'CHA_BLOCNOTE');
            3  : AddValues('AFF_TIERS'        , 'CHA_CODECLIENT');
            4  : AddValues('AFF_DATEDEBUT'    , 'CHA_DATEDEBUT');
            5  : AddValues('AFF_DATEFIN'      , 'CHA_DATEFIN');
            6  : AddValues('AFF_BOOLLIBRE1'   , 'CHA_VISA');
            7  : AddValues('AFF_STATUTAFFAIRE', 'CHA_STATUT');
            8  : AddValues('AFF_RESPONSABLE'  , 'CHA_RESPONSABLE');
            9  : AddValues('AFF_DATECREATION' , 'CHA_DATECREATION');
            10 : AddValues('AFF_DATEMODIF'    , 'CHA_DATEMODIF');
          end;
        end;
        ArrLen := 16;
        SetLength(BTPArrAdditionalFields, ArrLen);
        SetLength(TMPArrAdditionalFields, ArrLen);
        for Cpt :=  Low(BTPArrAdditionalFields) to High(BTPArrAdditionalFields) do
        begin
          case Cpt of
            0  : AddValues('LADR_LIBELLE'   , 'CHA_ADLLIBELLE' , True);
            1  : AddValues('LADR_ADRESSE1'  , 'CHA_ADLADRESSE1', True);
            2  : AddValues('LADR_ADRESSE2'  , 'CHA_ADLADRESSE2', True);
            3  : AddValues('LADR_ADRESSE3'  , 'CHA_ADLADRESSE3', True);
            4  : AddValues('LADR_CODEPOSTAL', 'CHA_ADLCP'      , True);
            5  : AddValues('LADR_VILLE'     , 'CHA_ADLVILLE'   , True);
            6  : AddValues('LADR_PAYS'      , 'CHA_ADLPAY'     , True);
            7  : AddValues('FADR_LIBELLE'   , 'CHA_ADFLIBELLE' , True);
            8  : AddValues('FADR_ADRESSE1'  , 'CHA_ADFADRESSE1', True);
            9  : AddValues('FADR_ADRESSE2'  , 'CHA_ADFADRESSE2', True);
            10 : AddValues('FADR_ADRESSE3'  , 'CHA_ADFADRESSE3', True);
            11 : AddValues('FADR_CODEPOSTAL', 'CHA_ADFCP'      , True);
            12 : AddValues('FADR_VILLE'     , 'CHA_ADFVILLE'   , True);
            13 : AddValues('FADR_PAYS'      , 'CHA_ADFPAY'     , True);
            14 : AddValues('LISTEDEVIS'     , 'CHA_LISTEDEVIS' , True);
            15 : AddValues('CODESOCIETE'    , 'CHA_CODESOCIETE', True);
          end;
        end;
      end;
    tnDevis :
      begin
        for Cpt :=  Low(BTPArrFields) to High(BTPArrFields) do
        begin
          case Cpt of
            0  : AddValues('GP_NUMERO'       , 'DEV_NUMDEVIS');
            1  : AddValues('GP_DATEPIECE'    , 'DEV_DATEDEVIS');
            2  : AddValues('GP_AFFAIRE'      , 'DEV_CODECHA');
            3  : AddValues('GP_TIERS'        , 'DEV_CODECLIENT');
            4  : AddValues('GP_REFINTERNE'   , 'DEV_LIBELLE');
            5  : AddValues('GP_REPRESENTANT' , 'DEV_RESPONSABLE');
            6  : AddValues('GP_TOTALHT'      , 'DEV_MONTANTHT');
            7  : AddValues('GP_BLOCNOTE'     , 'DEV_BLOCNOTE');
            8  : AddValues('GP_LIBREPIECE1'  , 'DEV_STATUS');
            9  : AddValues('GP_SOCIETE'      , 'DEV_CODESOCIETE');
            10 : AddValues('GP_DATECREATION' , 'DEV_DATECREATION');
            11 : AddValues('GP_DATEMODIF'    , 'DEV_DATEMODIF');
          end;
        end;
      end;
    tnLignesBR :
      begin
        for Cpt :=  Low(BTPArrFields) to High(BTPArrFields) do
        begin
          case Cpt of
            0  : AddValues('GL_NUMERO'       , 'LBR_NUMBR');
            1  : AddValues('GL_TIERS'        , 'LBR_FOURNISSEUR');
            2  : AddValues('GL_AFFAIRE'      , 'LBR_CODECHA');
            3  : AddValues('GL_NUMORDRE'     , 'LBR_NUMORDRE');
            4  : AddValues('GL_CODEARTICLE'  , 'LBR_CODEARTICLE');
            5  : AddValues('GL_LIBELLE'      , 'LBR_LIBELLE');
            6  : AddValues('GL_QTEFACT'      , 'LBR_QUANTITE');
            7  : AddValues('GL_PUHTDEV'      , 'LBR_PU');
            8  : AddValues('GL_UTILISATEUR'  , 'LBR_UTILISATEUR');
            9  : AddValues('GL_SOCIETE'      , 'LBR_CODESOCIETE');
            10 : AddValues('GL_DATECREATION', 'LBR_DATECREATION');
            11 : AddValues('GL_DATEMODIF'    , 'LBR_DATEMODIF');
          end;
        end;
      end;
    tnTiers :
      begin
        for Cpt :=  Low(BTPArrFields) to High(BTPArrFields) do
        begin
          case Cpt of
            0  : AddValues('T_AUXILIAIRE'  , 'TIE_AUXILIAIRE');
            1  : AddValues('T_COLLECTIF'   , 'TIE_COLLECTIF');
            2  : AddValues('T_NATUREAUXI'  , 'TIE_NATUREAUXI');
            3  : AddValues('T_LIBELLE'     , 'TIE_LIBELLE');
            4  : AddValues('T_DEVISE'      , 'TIE_DEVISE');
            5  : AddValues('T_ADRESSE1'    , 'TIE_ADRESSE1');
            6  : AddValues('T_ADRESSE2'    , 'TIE_ADRESSE2');
            7  : AddValues('T_ADRESSE3'    , 'TIE_ADRESSE3');
            8  : AddValues('T_CODEPOSTAL'  , 'TIE_CP');
            9  : AddValues('T_VILLE'       , 'TIE_VILLE');
            10 : AddValues('T_PAYS'        , 'TIE_PAYS');
            11 : AddValues('T_TELEPHONE'   , 'TIE_TELEPHONE');
            12 : AddValues('T_FAX'         , 'TIE_TELEPHONE2');
            13 : AddValues('T_TELEX'       , 'TIE_TELEPHONE3');
            14 : AddValues('T_EMAIL'       , 'TIE_EMAIL');
            15 : AddValues('T_RVA'         , 'TIE_WEBURL');
            16 : AddValues('T_COMPTATIERS' , 'TIE_COMPTA');
            17 : AddValues('T_FACTURE'     , 'TIE_TIERSFACTURE');
            18 : AddValues('T_PAYEUR'      , 'TIE_TIERSPAYEUR');
            19 : AddValues('T_BLOCNOTE'    , 'TIE_BLOCNOTE');
            20 : AddValues('T_SIRET'       , 'TIE_SIRET');
            21 : AddValues('T_NIF'         , 'TIE_NUMINTRACOMM');
            22 : AddValues('T_SOCIETE'     , 'TIE_CODESOCIETE');
            23 : AddValues('T_DATEMODIF'   , 'TIE_DATEMODIF');
            24 : AddValues('T_DATECREATION', 'TIE_DATECREATION');
          end;
        end;
      end;
  end;
end;

function TTnTreatment.GetPrefix : string;
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

function TTnTreatment.GetTMPFieldName(BTPFieldName : string) : string;
var
  Cpt : integer;
begin
  if BTPFieldName <> '' then
  begin
    for Cpt := Low(BTPArrFields) to High(BTPArrFields) do
    begin
     if ExtractFieldName(BTPArrFields[Cpt]) = BTPFieldName then
      begin
       Result := ExtractFieldName(TMPArrFields[Cpt]);
       break;
      end;
    end;
  end else
    Result := '';
end;
  
function TTnTreatment.GetTMPIndexFieldName : string;
begin
  case Tn of
    tnChantier : Result := 'CHA_CODE';
    tnDevis    : Result := 'DEV_NUMDEVIS';
    tnLignesBR : Result := 'LBR_NUMBR;LBR_NUMORDRE';
    tnTiers    : Result := 'TIE_AUXILIAIRE';
  else
    Result := '';
  end;
end;

function TTnTreatment.GetBTPIndexFieldName : string;
begin
  case Tn of
    tnChantier : Result := 'AFF_AFFAIRE';
    tnDevis    : Result := 'GP_NUMERO';
    tnLignesBR : Result := 'GL_NUMERO;GL_NUMORDRE';
    tnTiers    : Result := 'T_AUXILIAIRE';
  else
    Result := '';
  end;
end;
              
function TTnTreatment.GetBTPLabelFieldName : string;
begin
  case Tn of
    tnChantier : Result := 'AFF_LIBELLE';
    tnDevis    : Result := 'GP_REFINTERNE';
    tnLignesBR : Result := 'GL_LIBELLE';
    tnTiers    : Result := 'T_LIBELLE';
  else
    Result := '';
  end;
end;

function TTnTreatment.GetSqlDataExist(FieldsList, KeyValue1, KeyValue2 : string) : string;
begin
  case Tn of
    tnLignesBR : Result := Format('SELECT %s FROM LIGNESBR WHERE LBR_NUMBR = ''%s'' AND LBR_NUMORDRE = %s'  , [FieldsList, KeyValue1, KeyValue2]);
  else
    Result := Format('SELECT %s FROM %s WHERE %s = ''%s'''  , [FieldsList, TUtilBTPVerdon.GetTMPTableName(Tn), GetTMPIndexFieldName, KeyValue1]);
  end;
end;

procedure TTnTreatment.SetLinkedRecords(TobAdd, TobData : TOB);

  procedure ClearValues;
  begin
    TobAdd.SetString('LADR_LIBELLE'    , '');
    TobAdd.SetString('LADR_ADRESSE1'   , '');
    TobAdd.SetString('LADR_ADRESSE2'   , '');
    TobAdd.SetString('LADR_ADRESSE3'   , '');
    TobAdd.SetString('LADR_CODEPOSTAL' , '');
    TobAdd.SetString('LADR_VILLE'      , '');
    TobAdd.SetString('LADR_PAYS'       , '');
    TobAdd.SetString('FADR_LIBELLE'    , '');
    TobAdd.SetString('FADR_ADRESSE1'   , '');
    TobAdd.SetString('FADR_ADRESSE2'   , '');
    TobAdd.SetString('FADR_ADRESSE3'   , '');
    TobAdd.SetString('FADR_CODEPOSTAL' , '');
    TobAdd.SetString('FADR_VILLE'      , '');
    TobAdd.SetString('FADR_PAYS'       , '');
    TobAdd.SetString('LISTEDEVIS'      , '');
  end;

  procedure AddAdress;
  var
    Sql     : string;
    TobAdr  : TOB;
    TobAdrL : TOB;
    Cpt     : integer;

    procedure AddValues(Prefix : string);
    begin
      TobAdd.SetString(Format('%sADR_LIBELLE'    , [Prefix]), TobAdrL.GetString('ADR_LIBELLE'));
      TobAdd.SetString(Format('%sADR_ADRESSE1'   , [Prefix]), TobAdrL.GetString('ADR_ADRESSE1'));
      TobAdd.SetString(Format('%sADR_ADRESSE2'   , [Prefix]), TobAdrL.GetString('ADR_ADRESSE2'));
      TobAdd.SetString(Format('%sADR_ADRESSE3'   , [Prefix]), TobAdrL.GetString('ADR_ADRESSE3'));
      TobAdd.SetString(Format('%sADR_CODEPOSTAL' , [Prefix]), TobAdrL.GetString('ADR_CODEPOSTAL'));
      TobAdd.SetString(Format('%sADR_VILLE'      , [Prefix]), TobAdrL.GetString('ADR_VILLE'));
      TobAdd.SetString(Format('%sADR_PAYS'       , [Prefix]), TobAdrL.GetString('ADR_PAYS'));
    end;

  begin
    TobAdr := TOB.Create('_ADR', nil, -1);
    try
      Sql := Format('SELECT ADR_LIBELLE'
                  + '     , ADR_ADRESSE1'
                  + '     , ADR_ADRESSE2'
                  + '     , ADR_ADRESSE3'
                  + '     , ADR_CODEPOSTAL'
                  + '     , ADR_VILLE'
                  + '     , ADR_PAYS'
                  + '     , ADR_TYPEADRESSE'
                  + ' FROM ADRESSES'
                  + ' WHERE ADR_REFCODE     = "%s"'
                  + '   AND ADR_TYPEADRESSE IN ("INT", "AFA")'
                  , [TobData.GetString('AFF_AFFAIRE')]);
      TobAdr.LoadDetailFromSQL(Sql, True);
      for Cpt := 0 to pred(TobAdr.Detail.count) do
      begin
        TobAdrL := TobAdr.Detail[Cpt];
        if TobAdrL.GetString('ADR_TYPEADRESSE') = 'INT' then
          AddValues('L')
        else if TobAdrL.GetString('ADR_TYPEADRESSE') = 'AFA' then
          AddValues('F');
      end;
    finally
      FreeAndNil(TobAdr);
    end;
  end;

  procedure AddQuotationList;
  var
    Sql   : string;
    Value : string;
    Qry   : TQuery;
  begin
    Sql := Format('SELECT GP_NUMERO'
                + ' FROM PIECE'
                + ' WHERE GP_NATUREPIECEG = "DBT"'
                + '       AND GP_TIERS    = "%s"'
                + '       AND GP_AFFAIRE  = "%s"'
                  , [TobData.GetString('AFF_TIERS'), TobData.GetString('AFF_AFFAIRE')]);
    Qry := OpenSql(Sql, True);
    try
      while not Qry.Eof do
      begin
        Value := Format('%s;%s', [Value, Qry.Fields[0].AsString]);
        Qry.Next;
      end;
    finally
      Ferme(Qry);
    end;
    Value := Copy(Value, 2, length(Value));
    TobAdd.SetString('LISTEDEVIS' , Value);
  end;

begin
  TobAdd.ClearDetail;
  case Tn of
    tnChantier :
      begin
        ClearValues;
        AddAdress;
        AddQuotationList;
      end;
  end;
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
  Prefix := GetPrefix;
  Result := Format('%s_LOCK;%s_TRAITE;%s_DATETRAITE', [Prefix, Prefix, Prefix]);
end;

function TTnTreatment.GetFieldsListFromArray(ArrData: array of string): string;
var
  Cpt       : integer;
  FieldName : string;
begin
  Result := '';
  for Cpt := Low(ArrData) to High(ArrData) do
  begin
    FieldName := ExtractFieldName(ArrData[Cpt]);
    Result := Format('%s, %s', [Result, FieldName]);
  end;
  Result := copy(Result, 2, length(Result));
end;

function TTnTreatment.GetValue(FieldNameBTP, FieldNameTMP : string; FieldType : tTypeField; TobData : TOB) : string;
var
  FieldSize : integer;
  Value     : variant;     
begin
  Result := '';
  Value  := TobData.GetValue(FieldNameBTP);
  case FieldType of
    ttfNumeric, ttfInt                     : Result := StringReplace(Value, ',', '.', [rfReplaceAll]);
    ttfDate                                : Result := Format('''%s''', [Tools.UsDateTime_(Value)]); 
    ttfCombo, ttfText, ttfBoolean, ttfMemo : begin
                                               Result    := Value; 
                                               Result    := Tools.iif(FieldType = ttfMemo, BlobToString(Result), Result);
                                               FieldSize := GetTMPFieldSizeMax(FieldNameTMP);
                                               Result    := Tools.iif(FieldSize > -1, Trim(Copy(Result, 1, FieldSize)), Result);
                                               if pos('''', Result) > 0 then
                                                 Result := StringReplace(Result, '''', '''''', [rfReplaceAll]);
                                               Result := Format('''%s''', [Result]);
                                              end;
  end;    
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

function TTnTreatment.GetSqlInsertAdditionalFields : string;
var
  Cpt : integer;
begin
  Result := '';
  for Cpt :=  Low(TMPArrAdditionalFields) to High(TMPArrAdditionalFields) do
    Result := Format('%s, %s', [Result, ExtractFieldName(TMPArrAdditionalFields[Cpt])]);
  Result := Copy(Result, 2, length(Result));
end;

function TTnTreatment.GetSqlInsertValues(TobData: TOB; IsAdditional : boolean=False): string;
var
  Cpt          : integer;
  FieldNameTMP : string;
  FieldNameBTP : string;
  FieldType    : tTypeField;
begin
  Result := '';
  for Cpt := 0 to Tools.iif(not IsAdditional, High(BTPArrFields), High(BTPArrAdditionalFields)) do
  begin
    FieldNameBTP := ExtractFieldName(Tools.iif(not IsAdditional, BTPArrFields[Cpt], BTPArrAdditionalFields[Cpt]));
    FieldNameTMP := ExtractFieldName(Tools.iif(not IsAdditional, TMPArrFields[Cpt], TMPArrAdditionalFields[Cpt]));
    FieldType    := Tools.GetTypeFieldFromStringType(ExtractFieldType(Tools.iif(not IsAdditional, BTPArrFields[Cpt], BTPArrAdditionalFields[Cpt])));
    Result       := Format('%s, %s', [Result, GetValue(FieldNameBTP, FieldNameTMP, FieldType, TobData)]);
  end;
  if not IsAdditional then
    Result := Format('%s, ''%s'', ''%s'', ''%s''', [Result, LockDefaultValue, TraiteDefaultValue, DateTraiteDefaultValue]);
  Result := Copy(Result, 2, length(Result));
end;

function TTnTreatment.GetSqlUpdate(TobData, TobAdd : TOB; KeyValue1, KeyValue2 : string) : string;
var
  Cpt          : integer;
  FieldNameTMP : string;
  FieldNameBTP : string;
  Sql          : string;
  Prefix       : string;
  SystemFields : string;
  FieldType    : tTypeField;
begin
  for Cpt := 0 to High(TMPArrFields) do
  begin
    FieldNameTMP := ExtractFieldName(TMPArrFields[Cpt]);
    FieldNameBTP := ExtractFieldName(BTPArrFields[Cpt]);
    FieldType    := Tools.GetTypeFieldFromStringType(ExtractFieldType(TMPArrFields[Cpt]));
    Sql          := Format('%s, %s=%s', [Sql, FieldNameTMP, GetValue(FieldNameBTP, FieldNameTMP, FieldType, TobData)]);
  end;
  case Tn of
    tnChantier :
    begin
      for Cpt :=  Low(BTPArrAdditionalFields) to High(BTPArrAdditionalFields) do
      begin
        FieldNameTMP := ExtractFieldName(TMPArrAdditionalFields[Cpt]);
        FieldNameBTP := ExtractFieldName(BTPArrAdditionalFields[Cpt]);
        FieldType    := Tools.GetTypeFieldFromStringType(ExtractFieldType(TMPArrAdditionalFields[Cpt]));
        Sql          := Format('%s, %s=%s', [Sql, FieldNameTMP, GetValue(FieldNameBTP, FieldNameTMP, FieldType, TobAdd)]); //TobAdd.GetString(FieldNameBTP)]);
      end;
    end;
  end;
  Prefix       := GetPrefix;
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
    tnLignesBR : Result := Format('UPDATE LIGNESBR SET %s WHERE LBR_NUMBR = ''%s'' AND LBR_NUMORDRE = %s', [Sql, KeyValue1, KeyValue2]);
  else
    Result := Format('UPDATE %s SET %s WHERE %s = ''%s''', [TUtilBTPVerdon.GetTMPTableName(Tn), Sql, GetTMPIndexFieldName, KeyValue1]);
  end;
end;

function TTnTreatment.GetSqlInsert(TobData, TobAdd : TOB) : string;
var
  Fields    : string;
  Values    : string;
begin
  Fields := GetSqlInsertFields;
  Values := GetSqlInsertValues(TobData);
  case Tn of
    tnChantier :
      begin
        Fields := Format('%s, %s', [Fields, GetSqlInsertAdditionalFields]);
        Values := Format('%s, %s', [Values, GetSqlInsertValues(TobAdd, True)]);
      end;
  end;
  Result := Format('INSERT INTO %s (%s) VALUES (%s)', [TUtilBTPVerdon.GetTMPTableName(Tn), Fields, Values]);
end;

function TTnTreatment.GetDataSearchSql : string;

  function GetLastSynchro : string;
  begin
    Result := Tools.CastDateTimeForQry(StrToDatetime(LastSynchro));
  end;

begin
  case Tn of
    tnTiers    : Result := Format('SELECT %s FROM TIERS   WHERE T_DATEMODIF >= "%s" ORDER BY T_AUXILIAIRE'                                       , [GetFieldsListFromArray(BTPArrFields), GetLastSynchro]);
    tnDevis    : Result := Format('SELECT %s FROM PIECE   WHERE GP_NATUREPIECEG = "DBT" AND GP_DATEMODIF >= "%s" ORDER BY GP_NUMERO'             , [GetFieldsListFromArray(BTPArrFields), GetLastSynchro]);
    tnLignesBR : Result := Format('SELECT %s FROM LIGNE   WHERE GL_NATUREPIECEG = "BLF" AND GL_DATEMODIF >= "%s" ORDER BY GL_NUMERO, GL_NUMLIGNE', [GetFieldsListFromArray(BTPArrFields), GetLastSynchro]);
    tnChantier : Result := Format('SELECT %s FROM AFFAIRE WHERE AFF_AFFAIRE LIKE "%s" AND AFF_DATEMODIF >= "%s" ORDER BY AFF_AFFAIRE'            , [GetFieldsListFromArray(BTPArrFields), 'A%', GetLastSynchro]);
  else
    Result := '';
  end;
end;

function TTnTreatment.GetTMPFieldSizeMax(FieldName : string) : integer;
begin
  case Tools.CaseFromString(FieldName, [  'CHA_CODE'   , 'CHA_BLOCNOTE', 'CHA_ADLCP', 'CHA_ADFCP'
                                        , 'DEV_CODECHA', 'DEV_BLOCNOTE'
                                        , 'LBR_CODECHA'
                                        , 'TIE_CP'     , 'TIE_BLOCNOTE']) of
    {CHA_CODE}     0 : Result := 8;
    {CHA_BLOCNOTE} 1 : Result := 256;
    {CHA_ADLCP}    2 : Result := 5;
    {CHA_ADFCP}    3 : Result := 5;
    {DEV_CODECHA}  4 : Result := 8;
    {DEV_BLOCNOTE} 5 : Result := 256;
    {LBR_CODECHA}  6 : Result := 8;
    {TIE_CP}       7 : Result := 5;
    {TIE_BLOCNOTE} 8 : Result := 256;
  else
    Result := -1;
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

//function TTnTreatment.TnTreatment(TobTable, TobAdd, TobQry: TOB; AdoQryL : AdoQry; BTPArrFields, TMPArrFields, BTPArrAdditionalFields, TMPArrAdditionalFields : array of string): boolean;
function TTnTreatment.TnTreatment(TobTable, TobAdd, TobQry: TOB; AdoQryL : AdoQry): boolean;
var
  TobL          : TOB;
  Cpt           : integer;
  InsertQty     : integer;
  UpdateQty     : integer;
  OtherQty      : integer;
  FieldSize     : integer;
  SqlUnlock     : string;
  Sql           : string;
  Lock          : string;
  Treat         : string;
  KeyFieldsName : string;
  KeyField1     : string;
  KeyField2     : string;
  KeyValue1     : string;
  KeyValue2     : string;
  LabelValue    : string;
  Values        : string;
  FindData      : boolean;
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
    AdoQryL.FieldsList := Format('%s_LOCK,%s_TRAITE', [GetPrefix, GetPrefix]);
    AdoQryL.LogValues  := LogValues;
    AdoQryL.TSLResult.Clear;
    KeyFieldsName := GetTMPIndexFieldName;
    SqlUnlock     := Format('UPDATE %s SET %s_LOCK = ''N'' WHERE %s IN(', [TUtilBTPVerdon.GetTMPTableName(Tn), GetPrefix, Tools.ReadTokenSt_(KeyFieldsName, ';')]);
    FindData      := False;
    for Cpt := 0 to pred(TobTable.Detail.Count) do
    begin
      TobL          := TobTable.Detail[Cpt];
      KeyFieldsName := GetBTPIndexFieldName;
      KeyField1     := Tools.ReadTokenSt_(KeyFieldsName, ';'); //Tools.iif(pos(';', KeyFieldsName) > 0, Tools.ReadTokenSt_(KeyFieldsName, ';'), KeyFieldsName);
      KeyField2     := Tools.ReadTokenSt_(KeyFieldsName, ';'); //Tools.iif(pos(';', KeyFieldsName) > 0, Tools.ReadTokenSt_(KeyFieldsName, ';'), '');
      KeyValue1     := TobL.GetString(KeyField1);
      FieldSize     := GetTMPFieldSizeMax(GetTMPFieldName(KeyField1));
      KeyValue1     := Tools.iif(FieldSize > -1, Trim(Copy(KeyValue1, 1, FieldSize)), KeyValue1);
      if KeyField2 <> '' then
      begin
        KeyValue2 := TobL.GetString(KeyField2);
        FieldSize := GetTMPFieldSizeMax(GetTMPFieldName(KeyField2));
        KeyField2 := Tools.iif(FieldSize > -1, Trim(Copy(KeyField2, 1, FieldSize)), KeyField2);
      end;
      SetLinkedRecords(TobAdd, TobL);
      LabelValue := TobL.GetString(GetBTPLabelFieldName);
      SqlUnlock  := Format('%s%s''%s''', [SqlUnlock, Tools.iif(Cpt = 0, '', ', '), KeyValue1]); // Prépare update de Unlock
      AdoQryL.Request := GetSqlDataExist(AdoQryL.FieldsList, KeyValue1, KeyValue2); // Test si enregistrement existe
      AdoQryL.SingleTableSelect;
      if LogValues.LogLevel = 2 then TUtilBTPVerdon.AddLog(Tn, Format('%s%s%s%s - %s', [Tools.iif(AdoQryL.RecordCount = 1, 'Modification : ', 'Création : '), KeyValue1, Tools.iif(KeyValue2 <> '', ' ', ''), KeyValue2, LabelValue]), LogValues, 3);
      if AdoQryL.RecordCount = 1 then // Update
      begin
        Values := AdoQryL.TSLResult[0];
        Lock   := Tools.ReadTokenSt_(Values, ToolsTobToTsl_Separator);
        Treat  := Tools.ReadTokenSt_(Values, ToolsTobToTsl_Separator);
        if Lock = 'N' then
        begin
          inc(UpdateQty);
          FindData  := True;
          Sql := GetSqlUpdate(TobL, TobAdd, KeyValue1, KeyValue2);
        end else
          Inc(OtherQty);
      end else
      begin
        Inc(InsertQty);
        FindData  := True;
        Sql := GetSqlInsert(TobL, TobAdd);
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
