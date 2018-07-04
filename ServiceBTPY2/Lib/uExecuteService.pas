unit uExecuteService;

interface

uses
  Windows
  , Messages
  , SysUtils
  , Classes
  , Graphics
  , Controls
  , Dialogs
  , CommonTools
  , UConnectWSConst
  ;

type
  TSvcSyncBTPY2Log = (ssbylNone, ssbylLog, ssbylWindows);

  TSvcSyncBTPY2Execute = class(TObject)
  private
    TSlConnectionValues            : TStringList;
    TSlValues                      : TStringList;
    TSlFilter                      : TStringList;
    TSlIndice                      : TStringList;
    TSlCacheThirdBTP               : TStringList;
    TSlCacheSectionBTP             : TStringList;
    TSlCacheAcountBTP              : TStringList;
    TSlCachePaymentBTP             : TStringList;
    TSlCacheCorrespBTP             : TStringList;
    TSlCacheCurrencyBTP            : TStringList;
    TSlCacheCountryBTP             : TStringList;
    TSlCacheRecoveryBTP            : TStringList;
    TSlCacheCommonBTP              : TStringList;
    TSlCacheChoixCodBTP            : TStringList;
    TSlCacheJournalBTP             : TStringList;
    TSlCacheBankIdBTP              : TStringList;
    TSlCacheChangeRateBTP          : TStringList;
    TSlCacheFiscalYearBTP          : TStringList;
    TSlcacheSocietyParamBTP        : TStringList;
    TSlcacheEstablishmentBTP       : TStringList;
    TSlcachePaymentModeBTP         : TStringList;
    TSlcacheZipCodeBTP             : TStringList;
    TSlcacheContactBTP             : TStringList;
    TSlCacheThirdY2                : TStringList;
    TSlCacheSectionY2              : TStringList;
    TSlCacheAcountY2               : TStringList;
    TSlCachePaymentY2              : TStringList;
    TSlCacheCorrespY2              : TStringList;
    TSlCacheCurrencyY2             : TStringList;
    TSlCacheCountryY2              : TStringList;
    TSlCacheRecoveryY2             : TStringList;
    TSlCacheCommonY2               : TStringList;
    TSlCacheChoixCodY2             : TStringList;
    TSlCacheJournalY2              : TStringList;
    TSlCacheBankIdY2               : TStringList;
    TSlCacheChangeRateY2           : TStringList;
    TSlCacheFiscalYearY2           : TStringList;
    TSlcacheSocietyParamY2         : TStringList;
    TSlcacheEstablishmentY2        : TStringList;
    TSlcachePaymentModeY2          : TStringList;
    TSlcacheZipCodeY2              : TStringList;
    TSlcacheContactY2              : TStringList;
    TSlCacheWSFields               : TStringList;
    TSlCacheSendAccParam           : TStringList;
    TSlCacheGetY2Data              : TStringList;
    TSLCacheUpdateFrequencySetting : TStringList;
    TSLUpdateInsertData            : TStringList;
    TSLTRAFileQty                  : TStringList;
    AdoQryY2                       : AdoQry;
    AdoQryBTP                      : AdoQry;
    LogValues                      : T_WSLogValues;
    BTPValues                      : T_WSBTPValues;
    Y2Values                       : T_WSY2Values;
    LogFile                        : TextFile;
    SettingsFilePath               : string;

    procedure ClearValuesConnection;
    procedure SetLastSyncIniFile;
    function GetSettingsFileName(Extension: string): string;
    procedure WriteLog(TypeDebug: TSvcSyncBTPY2Log; Text: string; Level: integer);
    function GetValueFormSettingLine(LineValue: string): string;
    procedure SetFilterFromDSType_Get(DSType: T_WSDataService; TSl: TStringList);
    procedure SetFilterFromDSType_Set(DSType: T_WSDataService; Value, Value1, Value2: string; TSl: TStringList);
    function GetInfoFromDSType(InfoType: T_WSInfoFromDSType; DSType: T_WSDataService; FieldName: string = ''): string;
    procedure ExtractIndice(Cpt: integer; TSlOrig, TSlResult: TStringList);
    function AddData(wsAction: T_WSAction; DSType: T_WSDataService; lTSlValues: TStringList): boolean;
    function GetKeyValues(DSType: T_WSDataService; lTSlValues: TStringList): string;
    function GetIndiceY2DataCache(DSType: T_WSDataService; TSLValue: TStringList): string;
    function AddQuotes(DSType: T_WSDataService; FieldName, FieldValue: string): string;
    function GetY2Data(DSType: T_WSDataService): boolean;
    function GetData: boolean;
    function SendData: boolean;
    procedure ReadSettings;
    procedure LogsManagement;
    procedure CreateMemoryCache;
    procedure AddWindowsLog(Start: boolean);
    procedure LoadEcrFromBE0(AdoQryEcr, AdoQryAna: AdoQry; BE0Values: string);
    procedure SetSendY2TSl(AdoQryEcr, AdoQryAna: AdoQry; TSlResult: TStringList);
    procedure SearchRelatedParameters(TSlEntry: TStringList);
    procedure SearchOthersParameters;
    function GetFieldsList(lTSl: TStringList): string;
    procedure AddUpdateValues(AdoQryParam: AdoQry; DSType: T_WSDataService; FieldsList: string; KeyValue1, KeyValue2: string; IsRelatedParameters: boolean);
    function FieldExistInTSl(lTSl : TStringList; FieldName: string): boolean;
    function GetTSlFromDSType(DSType: T_WSDataService; IsBTP: boolean): TStringList;
    function GetTableNameFromDSType(DSType: T_WSDataService): string;
    function GetFieldTypeFromCache(DSType: T_WSDataService; FieldName: string): tTypeField;
    function SendY2Settings : boolean;

  public
    ApplicationName: string;
    SecondTimeout: integer;

    procedure CreateObjects;
    procedure FreeObjects;
    function ServiceExecute: Boolean;
    procedure InitApplication;
  end;

implementation

uses
  Registry
  , uWSDataService
  , SvcMgr
  , StrUtils
  , UConnectWSCEGID
  , IniFiles
  , DateUtils
  , TRAFileUtil
  , uLog
  ;

const
  BTPOrigin          = '#BTFIELDS#';
  Y2Origin           = '#Y2FIELDS#';
  IniFileLastSynchro = 'LastSynchro';
  IniFileBTPUser     = 'BTPUser';
  IniFileBTPServer   = 'BTPServer';
  IniFileBTPDataBase = 'BTPDataBase';
  IniFileY2Server    = 'Y2Server';
  IniFileY2DataBase  = 'Y2DataBase';

procedure TSvcSyncBTPY2Execute.CreateMemoryCache;
var
  lTslFieldsList: TStringList;

  procedure SetBTPDefaultValue(DSType: T_WSDataService);
  var
    CptCache     : integer;
    lTsl         : TStringList;
    Prefix       : string;
    Value        : string;
    FieldName    : string;
    FieldType    : string;
    DefaultValue : string;
  begin
    lTsl   := GetTSlFromDSType(DSType, True);
    Prefix := TReadWSDataService.GetPrefixFromDSType(DSType);
    AdoQryBTP.TSLResult.Clear;
    AdoQryBTP.FieldsList := 'DH_NOMCHAMP,DH_TYPECHAMP';
    AdoQryBTP.Request    := Format('SELECT %s FROM DECHAMPS WHERE DH_PREFIXE = ''%s'' ORDER BY DH_NOMCHAMP', [AdoQryBTP.FieldsList, Prefix]);
    AdoQryBTP.SingleTableSelect;
    if AdoQryBTP.TSLResult.Count > 0 then
    begin
      for CptCache := 0 to pred(AdoQryBTP.TSLResult.Count) do
      begin
        Value := AdoQryBTP.TSLResult[CptCache];
        FieldName := Tools.ReadTokenSt_(Value, ToolsTobToTsl_Separator);
        FieldType := Value;
        if Tools.GetTypeFieldFromStringType(FieldType) <> ttfMemo then // Exclusion des memo
        begin
          case Tools.GetTypeFieldFromStringType(Value) of
            ttfNumeric : DefaultValue := '0';
            ttfInt     : DefaultValue := '0';
            ttfBoolean : DefaultValue := '''-''';
            ttfDate    : DefaultValue := '''' + DateToStr(2) + '''';
            ttfMemo    : DefaultValue := '''''';
            ttfCombo   : DefaultValue := '''''';
            ttfText    : DefaultValue := '''''';
          end;
          lTsl.Add(FieldName + ToolsTobToTsl_Separator + FieldType + ToolsTobToTsl_Separator + DefaultValue);
        end;
      end;
    end;
    AdoQryBTP.TSLResult.Clear;
    AdoQryBTP.RecordCount := 0;
  end;

  procedure SetY2FieldsList(DSType: T_WSDataService);
  var
    lTsl   : TStringList;
    Cpt    : Integer;
    Prefix : string;
  begin
    lTsl := GetTSlFromDSType(DSType, False);
    for Cpt := 0 to pred(lTslFieldsList.Count) do
    begin
      Prefix := Copy(lTslFieldsList[Cpt], pos('=', lTslFieldsList[Cpt]) + 1, length(lTslFieldsList[Cpt]));
      Prefix := copy(Prefix, 1, pos('_', Prefix));
      if Prefix = TReadWSDataService.GetPrefixFromDSType(DSType) + '_' then
        lTsl.Add(copy(lTslFieldsList[Cpt], pos('=', lTslFieldsList[Cpt]) + 1, length(lTslFieldsList[Cpt])) + ToolsTobToTsl_Separator);
    end;
  end;

  procedure SetViewFields(DSType: T_WSDataService);
  begin
    TSlCacheWSFields.Add(TReadWSDataService.GetWSNameFromDSType(DSType) + '=' + TReadWSDataService.GetFiedsListFromDsType(DSType));
  end;

begin
  { Liste des champs BTP }
  SetBTPDefaultValue(wsdsThird);
  SetBTPDefaultValue(wsdsAnalyticalSection);
  SetBTPDefaultValue(wsdsAccount);
  SetBTPDefaultValue(wsdsPaymenChoice);
  SetBTPDefaultValue(wsdsCorrespondence);
  SetBTPDefaultValue(wsdsCurrency);
  SetBTPDefaultValue(wsdsCountry);
  SetBTPDefaultValue(wsdsRecovery);
  SetBTPDefaultValue(wsdsCommon);
  SetBTPDefaultValue(wsdsChoixCod);
  SetBTPDefaultValue(wsdsJournal);
  SetBTPDefaultValue(wsdsBankIdentification);
  SetBTPDefaultValue(wsdsChangeRate);
  SetBTPDefaultValue(wsdsFiscalYear);
  SetBTPDefaultValue(wsdsSocietyParameters);
  SetBTPDefaultValue(wsdsEstablishment);
  SetBTPDefaultValue(wsdsPaymentMode);
  SetBTPDefaultValue(wsdsZipCode);
  SetBTPDefaultValue(wsdsContact);
  { Liste des champs par vue }
  SetViewFields(wsdsThird);
  SetViewFields(wsdsAnalyticalSection);
  SetViewFields(wsdsAccount);
  SetViewFields(wsdsJournal);
  SetViewFields(wsdsBankIdentification);
  SetViewFields(wsdsChoixCod);
  SetViewFields(wsdsCommon);
  SetViewFields(wsdsRecovery);
  SetViewFields(wsdsCountry);
  SetViewFields(wsdsCurrency);
  SetViewFields(wsdsCorrespondence);
  SetViewFields(wsdsPaymenChoice);
  SetViewFields(wsdsChangeRate);
  SetViewFields(wsdsFiscalYear);
  SetViewFields(wsdsSocietyParameters);
  SetViewFields(wsdsEstablishment);
  SetViewFields(wsdsPaymentMode);
  SetViewFields(wsdsZipCode);
  SetViewFields(wsdsContact);
  SetViewFields(wsdsFieldsList);
  { Liste des champs Y2 }
  lTslFieldsList := TStringList.Create;
  try
    TReadWSDataService.GetData(wsdsFieldsList, AdoQryY2.ServerName, AdoQryY2.DBName, lTslFieldsList, TSlCacheWSFields);
    SetY2FieldsList(wsdsThird);
    SetY2FieldsList(wsdsAnalyticalSection);
    SetY2FieldsList(wsdsAccount);
    SetY2FieldsList(wsdsJournal);
    SetY2FieldsList(wsdsBankIdentification);
    SetY2FieldsList(wsdsChoixCod);
    SetY2FieldsList(wsdsCommon);
    SetY2FieldsList(wsdsRecovery);
    SetY2FieldsList(wsdsCountry);
    SetY2FieldsList(wsdsCurrency);
    SetY2FieldsList(wsdsCorrespondence);
    SetY2FieldsList(wsdsPaymenChoice);
    SetY2FieldsList(wsdsChangeRate);
    SetY2FieldsList(wsdsFiscalYear);
    SetY2FieldsList(wsdsSocietyParameters);
    SetY2FieldsList(wsdsEstablishment);
    SetY2FieldsList(wsdsPaymentMode);
    SetY2FieldsList(wsdsZipCode);
    SetY2FieldsList(wsdsContact);
  finally
    FreeAndNil(lTslFieldsList);
  end;
end;

procedure TSvcSyncBTPY2Execute.AddWindowsLog(Start: boolean);
var
  LogText        : string;
  ConnectionLine : string;
  Cpt            : Integer;
begin
  if Start then
  begin
    LogText := Format('Démarrage de %s.', [ApplicationName])
             + Format('%s> Délai d''exécution   : %s secondes.', [#13#10, IntToStr(SecondTimeout)])
             + Format('%s> Nombre de connexions : %s', [#13#10, IntToStr(TSlConnectionValues.Count)]);
    for Cpt := 0 to pred(TSlConnectionValues.Count) do
    begin
      ConnectionLine := TSlConnectionValues[Cpt];
      LogText := LogText
               + Format('%s   %s :', [#13#10, Tools.ReadTokenSt_(ConnectionLine, '=')])
               + Format('%s    Compte utilisateur : %s', [#13#10, Tools.ReadTokenSt_(ConnectionLine, ';')])
               + Format('%s    BTP-Serveur : %s', [#13#10, Tools.ReadTokenSt_(ConnectionLine, ';')])
               + Format('%s    BTP-Base de données : %s', [#13#10, Tools.ReadTokenSt_(ConnectionLine, ';')])
               + Format('%s    BTP-Dernière synchronisation : %s', [#13#10, Tools.ReadTokenSt_(ConnectionLine, ';')])
               + Format('%s    Y2-Serveur : %s', [#13#10, Tools.ReadTokenSt_(ConnectionLine, ';')])
               + Format('%s    Y2-Base de données  : %s', [#13#10, Tools.ReadTokenSt_(ConnectionLine, ';')])
               + Format('%s    Y2-Dernière synchronisation : %s', [#13#10, Tools.ReadTokenSt_(ConnectionLine, ';')]);
    end;
    LogText := LogText
             + Format('%s> Objets créés :'                 , [#13#10])
             + Format('%s   TSlCacheThird'                 , [#13#10])
             + Format('%s   TSlCacheSection'               , [#13#10])
             + Format('%s   TSlCacheAcount'                , [#13#10])
             + Format('%s   TSlCachePayment'               , [#13#10])
             + Format('%s   TSlCacheCorresp'               , [#13#10])
             + Format('%s   TSlCacheCurrency'              , [#13#10])
             + Format('%s   TSlCacheCountry'               , [#13#10])
             + Format('%s   TSlCacheRecovery'              , [#13#10])
             + Format('%s   TSlCacheCommon'                , [#13#10])
             + Format('%s   TSlCacheChoixCod'              , [#13#10])
             + Format('%s   TSlCacheJournal'               , [#13#10])
             + Format('%s   TSlCacheBankId'                , [#13#10])
             + Format('%s   TSlCacheChangeRate'            , [#13#10])
             + Format('%s   TSlCacheWSFields'              , [#13#10])
             + Format('%s   TSlCacheFiscalYear'            , [#13#10])
             + Format('%s   TSlcacheSocietyParam'          , [#13#10])
             + Format('%s   TSlcacheEstablishment'         , [#13#10])
             + Format('%s   TSlcachePaymentMode'           , [#13#10])
             + Format('%s   TSlcacheZipCodeBTP'            , [#13#10])
             + Format('%s   TSlcacheContactBTP'            , [#13#10])
             + Format('%s   TSlCacheSendAccParam'          , [#13#10])
             + Format('%s   TSlCacheGetY2Data'             , [#13#10])
             + Format('%s   TSLCacheUpdateFrequencySetting', [#13#10])
             + Format('%s   TSLUpdateInsertData'           , [#13#10])
             + Format('%s   TSLTRAFileQty      '           , [#13#10]);
             ;
  end else
  begin

  end;
  WriteLog(ssbylWindows, LogText, 0);
end;

procedure TSvcSyncBTPY2Execute.LoadEcrFromBE0(AdoQryEcr, AdoQryAna: AdoQry; BE0Values: string);
var
  Cpt         : integer;
  FieldIndex  : integer;
  lBE0Values  : string;
  BE0Entity   : string;
  BE0Exercice : string;
  BE0Journal  : string;
  BE0NumPce   : string;
begin
  lBE0Values  := BE0Values;
  BE0Entity   := Tools.ReadTokenSt_(lBE0Values, ToolsTobToTsl_Separator);
  BE0Exercice := Tools.ReadTokenSt_(lBE0Values, ToolsTobToTsl_Separator);
  BE0Journal  := Tools.ReadTokenSt_(lBE0Values, ToolsTobToTsl_Separator);
  BE0NumPce   := Tools.ReadTokenSt_(lBE0Values, ToolsTobToTsl_Separator);
  AdoQryEcr.TSLResult.Clear;
  AdoQryEcr.Request := Format('SELECT %s FROM ECRITURE WHERE E_ENTITY = ''%s'' AND E_EXERCICE = ''%s'' AND E_JOURNAL = ''%s'' AND E_NUMEROPIECE = %s'
                           , [AdoQryEcr.FieldsList, BE0Entity, BE0Exercice, BE0Journal, BE0NumPce]);
  AdoQryEcr.SingleTableSelect;
  { Charge l'éventuelle analytique }
  for Cpt := 0 to pred(AdoQryEcr.TSLResult.Count) do
  begin
    FieldIndex := Tools.GetTSlIndexFromFieldName(AdoQryEcr.FieldsList, 'E_ANA');                         // Recherche l'index dans la liste des champs
    if Tools.GetStValueFromTSl(AdoQryEcr.TSLResult[Cpt], FieldIndex, ToolsTobToTsl_Separator) = 'X' then // Recherche la valeur dans liste des valeurs
    begin
      AdoQryAna.TSLResult.Clear;
      AdoQryAna.Request := Format('SELECT %s FROM ANALYTIQ WHERE Y_ENTITY = ''%s'' AND Y_EXERCICE = ''%s'' AND Y_JOURNAL = ''%s'' AND Y_NUMEROPIECE = %s'
                               , [  AdoQryAna.FieldsList
                                  , Tools.GetStValueFromTSl(AdoQryEcr.TSLResult[Cpt], Tools.GetTSlIndexFromFieldName(AdoQryEcr.FieldsList, 'E_ENTITY')     , ToolsTobToTsl_Separator)
                                  , Tools.GetStValueFromTSl(AdoQryEcr.TSLResult[Cpt], Tools.GetTSlIndexFromFieldName(AdoQryEcr.FieldsList, 'E_EXERCICE')   , ToolsTobToTsl_Separator)
                                  , Tools.GetStValueFromTSl(AdoQryEcr.TSLResult[Cpt], Tools.GetTSlIndexFromFieldName(AdoQryEcr.FieldsList, 'E_JOURNAL')    , ToolsTobToTsl_Separator)
                                  , Tools.GetStValueFromTSl(AdoQryEcr.TSLResult[Cpt], Tools.GetTSlIndexFromFieldName(AdoQryEcr.FieldsList, 'E_NUMEROPIECE'), ToolsTobToTsl_Separator)
                                 ]);
      AdoQryAna.SingleTableSelect;
      Break;
    end;
  end;
end;

procedure TSvcSyncBTPY2Execute.SetSendY2TSl(AdoQryEcr, AdoQryAna: AdoQry; TSlResult: TStringList);
var
  CptE            : integer;
  CptY            : integer;
  AxisIndex       : integer;
  LineNumberIndex : integer;
  LineValue       : string;
  FieldName       : string;
  FieldValue      : string;
  LineNumber      : string;
  AxisNumber      : string;
  WithAna         : boolean;
  AddAxis         : boolean;

  function GetLevel(Level: integer): string;
  begin
    Result := ToolsTobToTsl_LevelName + IntToStr(Level);
  end;

  function GetLineValue(FieldsList, FieldsValue: string): string;
  begin
    Result := '';
    while FieldsList <> '' do
    begin
      FieldName  := Tools.ReadTokenSt_(FieldsList, ',');
      FieldValue := Tools.ReadTokenSt_(FieldsValue, ToolsTobToTsl_Separator);
      Result     := Result + ToolsTobToTsl_Separator + FieldName + '=' + FieldValue;
    end;
  end;

begin
  WithAna := (AdoQryAna.TSLResult.Count > 0);
  if WithAna then
  begin
    LineNumberIndex := Tools.GetTSlIndexFromFieldName(AdoQryAna.FieldsList, 'Y_NUMLIGNE');
    AxisIndex       := Tools.GetTSlIndexFromFieldName(AdoQryAna.FieldsList, 'Y_AXE');
  end else
  begin
    LineNumberIndex := -1;
    AxisIndex       := -1;
  end;
  TSlResult.Add(GetLevel(1) + '=COMPTABILITE');
  for CptE := 0 to pred(AdoQryEcr.TSLResult.count) do
  begin
    LineValue := GetLineValue(AdoQryEcr.FieldsList, AdoQryEcr.TSLResult[CptE]);
    TSlResult.Add(GetLevel(2) + '=ECRITURE' + LineValue);
    { La ligne est ventilable, ajout de l'analytique }
    if (WithAna) and (Tools.GetStValueFromTSl(TSlResult[pred(TSlResult.Count)], 'E_ANA') = 'X') then
    begin
      AddAxis := False;
      LineNumber := Tools.GetStValueFromTSl(TSlResult[pred(TSlResult.Count)], 'E_NUMLIGNE');
      for CptY := 0 to pred(AdoQryAna.TSLResult.count) do
      begin
        if Tools.GetStValueFromTSl(AdoQryAna.TSLResult[CptY], LineNumberIndex, ToolsTobToTsl_Separator) = LineNumber then
        begin
          AxisNumber := Tools.GetStValueFromTSl(AdoQryAna.TSLResult[CptY], AxisIndex, ToolsTobToTsl_Separator);
          if not AddAxis then
          begin
            AddAxis := True;
            TSlResult.Add(GetLevel(3) + '=' + AxisNumber);
          end;
          LineValue := GetLineValue(AdoQryAna.FieldsList, AdoQryAna.TSLResult[CptY]);
          TSlResult.Add(GetLevel(4) + '=ANALYTIQ' + LineValue);
        end;
      end;
    end;
  end;
end;

function TSvcSyncBTPY2Execute.GetFieldsList(lTSl: TStringList): string;
var
  Cpt: integer;
begin
  Result := '';
  for Cpt := 0 to pred(lTSl.Count) do
    Result := Result + ',' + Copy(lTSl[Cpt], 1, Pos(ToolsTobToTsl_Separator, lTSl[Cpt]) - 1);
  Result := Copy(Result, 2, Length(Result));
end;

procedure TSvcSyncBTPY2Execute.AddUpdateValues(AdoQryParam: AdoQry; DSType: T_WSDataService; FieldsList: string; KeyValue1, KeyValue2: string; IsRelatedParameters: boolean);
var
  sIndex         : string;
  Y2FieldValue   : string;
  TmpData        : string;
  TmpFieldName   : string;
  TmpFieldValue  : string;
  BtpSendData    : string;
  Y2SendData     : string;
  CommonSendData : string;
  AddData        : string;
  Y2FieldsList   : string;
  Y2StartPosData : integer;
  Cpt            : integer;
  SameValue      : boolean;

  function GetRequest(lAdoQry: AdoQry): string;
  var
      FormattedLastSynchro : string;
  begin
    FormattedLastSynchro := FormatDateTime('yyyymmdd hh:nn:ss', StrToDateTime(BTPValues.LastSynchro));
    case DSType of
      wsdsThird               : Result := Format('SELECT %s FROM TIERS WHERE T_AUXILIAIRE = ''%s'' AND T_DATEMODIF > ''%s'''                  , [lAdoQry.FieldsList, KeyValue1, FormattedLastSynchro]);
      wsdsBankIdentification  : Result := Format('SELECT %s FROM RIB WHERE R_AUXILIAIRE = ''%s'' AND R_DATEMODIF > ''%s'''                    , [lAdoQry.FieldsList, KeyValue1, FormattedLastSynchro]);
      wsdsAnalyticalSection   : Result := Format('SELECT %s FROM SECTION WHERE S_AXE = ''%s'' AND S_SECTION = ''%s'' AND S_DATEMODIF > ''%s''', [lAdoQry.FieldsList, KeyValue1, KeyValue2, FormattedLastSynchro]);
      wsdsContact             : Result := Format('SELECT %s FROM CONTACT WHERE C_AUXILIAIRE = ''%s'' AND C_DATEMODIF > ''%s'''                , [lAdoQry.FieldsList, KeyValue1, FormattedLastSynchro]);
      wsdsChoixCod            : Result := Format('SELECT %s FROM CHOIXCOD WHERE CC_TYPE IN (%s) ORDER BY CC_TYPE, CC_CODE'                    , [lAdoQry.FieldsList, KeyValue1]);
      wsdsCurrency            : Result := Format('SELECT %s FROM DEVISE ORDER BY D_DEVISE'                                                    , [lAdoQry.FieldsList]);
      wsdsPaymenChoice        : Result := Format('SELECT %s FROM MODEREGL ORDER BY MR_MODEREGLE'                                              , [lAdoQry.FieldsList]);
      wsdsChangeRate          : Result := Format('SELECT %s FROM CHANCELL ORDER BY H_DEVISE, H_DATECOURS'                                     , [lAdoQry.FieldsList]);
    end;
  end;

  function SetDataToSend(FieldsList, FieldsValues: string): string;
  var
    lFieldsList   : string;
    lFieldsValues : string;
  begin
    Result := '';
    if FieldsList <> '' then
    begin
      lFieldsList   := FieldsList;
      lFieldsValues := FieldsValues;
      while lFieldsList <> '' do
        Result := Result + ToolsTobToTsl_Separator + Tools.ReadTokenSt_(lFieldsList, ',') + '=' + Tools.ReadTokenSt_(lFieldsValues, ToolsTobToTsl_Separator);
      Result := Copy(Result, 2, Length(Result));
    end else
      Result := '';
  end;

  function GetY2DataFromCache(FieldsList, FieldsValue: string): string;
  var
    FieldsKey    : string;
    KeyField1    : string;
    KeyField2    : string;
    KeyValue1    : string;
    KeyValue2    : string;
    FieldName    : string;
    FieldValue   : string;
    IndexName    : string;
    lFieldsList  : string;
    lFieldsValue : string;
    CacheIndex   : integer;
  begin
    FieldsKey    := GetInfoFromDSType(wsidFieldsKey, DSType);
    KeyField1    := Tools.ReadTokenSt_(FieldsKey, ';');
    KeyField2    := Tools.ReadTokenSt_(FieldsKey, ';');
    lFieldsList  := FieldsList;
    lFieldsValue := FieldsValue;
    while lFieldsList <> '' do
    begin
      FieldName := Tools.ReadTokenSt_(lFieldsList, ',');
      FieldValue := Tools.ReadTokenSt_(lFieldsValue, ToolsTobToTsl_Separator);
      case Tools.CaseFromString(FieldName, [KeyField1, KeyField2]) of
        {KeyField1} 0: KeyValue1 := FieldValue;
        {KeyField2} 1: KeyValue2 := FieldValue;
      end;
      if (KeyValue1 <> '') and (KeyValue2 <> '') then
        Break;
    end;
    IndexName  := Format('%s_%s%s', [GetInfoFromDSType(wsidTableName, DSType), KeyValue1, KeyValue2]);
    CacheIndex := TSlCacheGetY2Data.IndexOfName(IndexName);
    if CacheIndex > -1 then
      Result := copy(TSlCacheGetY2Data[CacheIndex], pos('=', TSlCacheGetY2Data[CacheIndex]) + 1, length(TSlCacheGetY2Data[CacheIndex]))
    else
      Result := '';
  end;

begin
  sIndex := Tools.iif(IsRelatedParameters, Format('%s_%s_%s', [GetInfoFromDSType(wsidTableName, DSType), KeyValue1, KeyValue2]), GetInfoFromDSType(wsidTableName, DSType));
  if (Tools.CanInsertedInTable(GetInfoFromDSType(wsidTableName, DSType), AdoQryBTP.ServerName, AdoQryBTP.DBName)) // On peut exporter la table
    and ((TSlCacheSendAccParam.IndexOfName(sIndex) = -1)                                                          // L'enregistrement en cours n'a pas déjà été ajouté
    and (((IsRelatedParameters) and (KeyValue1 <> '')) or (not IsRelatedParameters))) then
  begin
    { Charge les enregistrements depuis la base BTP }
    AdoQryParam.TSLResult.Clear;
    AdoQryParam.FieldsList := FieldsList;
    AdoQryParam.Request    := GetRequest(AdoQryParam);
    AdoQryParam.SingleTableSelect;
    if AdoQryParam.RecordCount > 0 then
    begin
      Y2FieldsList := GetFieldsList(GetTSlFromDSType(DSType, False));
      { Boucle sur les enregistrements BTP trouvés et tests s'ils existent dans les datas provenant d'Y2}
      for Cpt := 0 to pred(AdoQryParam.TSLResult.count) do
      begin
        CommonSendData := '';
        BtpSendData := SetDataToSend(AdoQryParam.FieldsList, AdoQryParam.TSLResult[Cpt]);
        Y2SendData := GetY2DataFromCache(AdoQryParam.FieldsList, AdoQryParam.TSLResult[Cpt]);
        if BtpSendData <> Y2SendData then // Les données ou colonnes sont différentes, calcul des champs communs (ainsi que les données) et on renvoie
        begin
          SameValue := (Y2SendData <> '');
          while BtpSendData <> '' do
          begin
            TmpData       := Tools.ReadTokenSt_(BtpSendData, ToolsTobToTsl_Separator);
            TmpFieldName  := Copy(TmpData, 1, Pos('=', TmpData) - 1);
            TmpFieldValue := Copy(TmpData, Pos('=', TmpData) + 1, Length(TmpData));
            AddData       := TmpFieldName + '=' + TmpFieldValue;
            if Pos(',' + TmpFieldName + ',', ',' + Y2FieldsList + ',') > 0 then
            begin
              CommonSendData := CommonSendData + ToolsTobToTsl_Separator + AddData;
              if (Y2SendData <> '') and (SameValue) then // S'il existe des données issues de Y2, test si la valeur est la même
              begin
                Y2StartPosData := Pos(TmpFieldName + '=', Y2SendData) + length(TmpFieldName) + 1;
                while Y2SendData[Y2StartPosData] <> ToolsTobToTsl_Separator do
                begin
                  Y2FieldValue := Y2FieldValue + Copy(Y2SendData, Y2StartPosData, 1);
                  Inc(Y2StartPosData);
                end;
                SameValue := (TmpFieldValue = Y2FieldValue);
                Y2FieldValue := '';
              end;
            end;
          end;
          { Ajouter uniquement si pas les mêmes valeur dans les champs communs }
          if not SameValue then
          begin
            CommonSendData := Copy(CommonSendData, 2, Length(CommonSendData));
            TSlCacheSendAccParam.Add(sIndex + '=' + CommonSendData);
          end;
        end;
      end;
    end else
      TSlCacheSendAccParam.Add(sIndex + '=' + WSCDS_EmptyValue);
  end;
end;

function TSvcSyncBTPY2Execute.FieldExistInTSl(lTSl : TStringList; FieldName: string): boolean;
var
  Cpt : integer;
begin
  Result := False;
  if (lTSL.Count > 0) and (FieldName <> '') then
  begin
    for Cpt := 0 to pred(lTSL.Count) do
    begin
      Result := (Copy(lTSL[Cpt], 1, Pos(ToolsTobToTsl_Separator, lTSL[Cpt])-1) = FieldName);
      if Result then
        Break;
    end;
  end;
end;

function TSvcSyncBTPY2Execute.GetTSlFromDSType(DSType: T_WSDataService; IsBTP: boolean): TStringList;
begin
  case DSType of
    wsdsThird              : Result := Tools.iif(IsBTP, TSlCacheThirdBTP, TSlCacheThirdY2);
    wsdsAnalyticalSection  : Result := Tools.iif(IsBTP, TSlCacheSectionBTP, TSlCacheSectionY2);
    wsdsAccount            : Result := Tools.iif(IsBTP, TSlCacheAcountBTP, TSlCacheAcountY2);
    wsdsJournal            : Result := Tools.iif(IsBTP, TSlCacheJournalBTP, TSlCacheJournalY2);
    wsdsBankIdentification : Result := Tools.iif(IsBTP, TSlCacheBankIdBTP, TSlCacheBankIdY2);
    wsdsChoixCod           : Result := Tools.iif(IsBTP, TSlCacheChoixCodBTP, TSlCacheChoixCodY2);
    wsdsCommon             : Result := Tools.iif(IsBTP, TSlCacheCommonBTP, TSlCacheCommonY2);
    wsdsRecovery           : Result := Tools.iif(IsBTP, TSlCacheRecoveryBTP, TSlCacheRecoveryY2);
    wsdsCountry            : Result := Tools.iif(IsBTP, TSlCacheCountryBTP, TSlCacheCountryY2);
    wsdsCurrency           : Result := Tools.iif(IsBTP, TSlCacheCurrencyBTP, TSlCacheCurrencyY2);
    wsdsCorrespondence     : Result := Tools.iif(IsBTP, TSlCacheCorrespBTP, TSlCacheCorrespY2);
    wsdsPaymenChoice       : Result := Tools.iif(IsBTP, TSlCachePaymentBTP, TSlCachePaymentY2);
    wsdsChangeRate         : Result := Tools.iif(IsBTP, TSlCacheChangeRateBTP, TSlCacheChangeRateY2);
    wsdsFiscalYear         : Result := Tools.iif(IsBTP, TSlCacheFiscalYearBTP, TSlCacheFiscalYearY2);
    wsdsSocietyParameters  : Result := Tools.iif(IsBTP, TSlcacheSocietyParamBTP, TSlcacheSocietyParamY2);
    wsdsEstablishment      : Result := Tools.iif(IsBTP, TSlcacheEstablishmentBTP, TSlcacheEstablishmentY2);
    wsdsPaymentMode        : Result := Tools.iif(IsBTP, TSlcachePaymentModeBTP, TSlcachePaymentModeY2);
    wsdsZipCode            : Result := Tools.iif(IsBTP, TSlcacheZipCodeBTP, TSlcacheZipCodeY2);
    wsdsContact            : Result := Tools.iif(IsBTP, TSlcacheContactBTP, TSlcacheContactY2);
  else
    Result := nil;
  end;
end;

function TSvcSyncBTPY2Execute.GetTableNameFromDSType(DSType: T_WSDataService): string;
begin
  case DSType of
    wsdsThird              : Result := Tools.GetTableNameFromTtn(ttnTiers);
    wsdsAnalyticalSection  : Result := Tools.GetTableNameFromTtn(ttnSection);
    wsdsAccount            : Result := Tools.GetTableNameFromTtn(ttnGeneraux);
    wsdsJournal            : Result := Tools.GetTableNameFromTtn(ttnJournal);
    wsdsBankIdentification : Result := Tools.GetTableNameFromTtn(ttnRib);
    wsdsChoixCod           : Result := Tools.GetTableNameFromTtn(ttnChoixCod);
    wsdsCommon             : Result := Tools.GetTableNameFromTtn(ttnCommun);
    wsdsRecovery           : Result := Tools.GetTableNameFromTtn(ttnRelance);
    wsdsCountry            : Result := Tools.GetTableNameFromTtn(ttnPays);
    wsdsCurrency           : Result := Tools.GetTableNameFromTtn(ttnDevise);
    wsdsChangeRate         : Result := Tools.GetTableNameFromTtn(ttnChancell);
    wsdsCorrespondence     : Result := Tools.GetTableNameFromTtn(ttnCorresp);
    wsdsPaymenChoice       : Result := Tools.GetTableNameFromTtn(ttnModeRegl);
    wsdsFiscalYear         : Result := Tools.GetTableNameFromTtn(ttnExercice);
    wsdsSocietyParameters  : Result := Tools.GetTableNameFromTtn(ttnParamSoc);
    wsdsEstablishment      : Result := Tools.GetTableNameFromTtn(ttnEtabliss);
    wsdsPaymentMode        : Result := Tools.GetTableNameFromTtn(ttnModePaie);
    wsdsZipCode            : Result := Tools.GetTableNameFromTtn(ttnCodePostaux);
    wsdsContact            : Result := Tools.GetTableNameFromTtn(ttnContact);
  else
    Result := '';
  end;
end;

function TSvcSyncBTPY2Execute.GetFieldTypeFromCache(DSType: T_WSDataService; FieldName: string): tTypeField;
var
  lTsl      : TStringList;
  Cpt       : integer;
  LineValue : string;
  FieldType : string;
begin
  Result := ttfNone;
  lTsl   := GetTSlFromDSType(DSType, True);
  for Cpt := 0 to pred(lTsl.Count) do
  begin
    LineValue := lTsl[Cpt];
    if Copy(LineValue, 1, Pos(ToolsTobToTsl_Separator, LineValue) - 1) = FieldName then
    begin
      Tools.ReadTokenSt_(LineValue, ToolsTobToTsl_Separator);
      FieldType := Tools.ReadTokenSt_(LineValue, ToolsTobToTsl_Separator);
      Result    := Tools.GetTypeFieldFromStringType(FieldType);
      Break;
    end;
  end;
end;

(* Le principe de génération du fichier TRA :
   - Pour chaque ligne de TSlCacheSendAccParam qui doit être envoyée
     . appel de la création de la ligne TRA par rapport à la table à traiter
     . ajout de cette ligne dans un TSTringList au format TRA
   - Ecriture sur disque de ce fichier
   - Test de la taille (doit être inférieure à 4 Mo)
   - Si supérieur, découpe du fichier en x fichiers
   - Compression du/des fichiers
   - Envoie du/des fichiers
*)
function TSvcSyncBTPY2Execute.SendY2Settings : boolean;
var
  Cpt          : integer;
  LineValue    : string;
  TableName    : string;
  traLine      : string;
  TempPath     : string;
  TmpFileName  : string;
  PathFileName : string;
  traRCode     : T_TraRecordCode;
  TslTra       : TStringList;
  SendEntry    : TSendEntryY2;

  function GetAdditionalDataFromThird(DSType : T_WSDataService) : string;
  var
    Auxiliary     : string;
    AddValue      : string;
    IndexValue    : string;
    MainFieldName : string;
    Index         : Integer;
    Cpt           : Integer;
    IsMain        : Boolean;
  begin
    Result     := '';
    Auxiliary  := Tools.GetStValueFromTSl(LineValue, 'T_AUXILIAIRE');
    case DSType of
      wsdsBankIdentification :
        begin
          IndexValue := Format('RIB_%s_'    , [Auxiliary]);
          MainFieldName := 'R_PRINCIPAL';
        end;
      wsdsContact :
        begin
          IndexValue := Format('CONTACT_%s_', [Auxiliary]);
          MainFieldName := 'C_PRINCIPAL';
        end;
    end;
    Index      := TSlCacheSendAccParam.IndexOfName(IndexValue);
    if Index > -1  then
    begin
      AddValue := TSlCacheSendAccParam[Index];
      IsMain   := (Tools.GetStValueFromTSl(TSlCacheSendAccParam[Index], MainFieldName) = 'X');
      if IsMain then
        Result := TSlCacheSendAccParam[Index]
      else
      begin
        for Cpt := Index to pred(TSlCacheSendAccParam.count) do
        begin
          AddValue := TSlCacheSendAccParam[Index];
          if     (Copy(AddValue, 1, Pos('=', AddValue)-1) = IndexValue)
             and (Tools.GetStValueFromTSl(TSlCacheSendAccParam[Index], MainFieldName) = 'X')
          then
          begin
            Result := TSlCacheSendAccParam[Index];
            Exit;
          end;
        end;
      end;
    end;
  end;

begin
  Result := True;
  if TSlCacheSendAccParam.count > 0 then
  begin
    { Génération du TRA }
    TslTra := TStringList.Create;
    try
      traLine := TraUtil.GetFirstLine(trapS5, trafoCLI, trafcJRL, traffETE, BTPValues.UserAdmin);
      if traLine <> '' then
        TslTra.Add(traLine);
      for Cpt := 0 to pred(TSlCacheSendAccParam.count) do
      begin
        LineValue := TSlCacheSendAccParam[Cpt];
        if pos(WSCDS_EmptyValue, LineValue) = 0 then
        begin
          TableName := Copy(LineValue, 1, Pos('=', LineValue)-1);
          if pos('_', TableName) > 0 then
            TableName := Copy(LineValue, 1, Pos('_', LineValue)-1);
          traRCode  := TraUtil.GetTraRecordCodeFromTableName(TableName);
          case traRCode of
            {TIERS} trarcThird : LineValue := LineValue + '#§**§#' + GetAdditionalDataFromThird(wsdsBankIdentification) + '#§**§#' + GetAdditionalDataFromThird(wsdsContact);
          end;
          traLine := TraUtil.GetTraLine(traRCode, LineValue);
          if traLine <> '' then
            TslTra.Add(traLine);
        end;
      end;
    finally
      TempPath    := GetEnvironmentVariable('TEMP');
      TmpFileName := Format('%s%s%s%s%s%s.TRA', [  IntToStr(YearOf(Now))
                                                 , IntToStr(MonthOf(Now))
                                                 , IntToStr(DayOf(Now))
                                                 , IntToStr(HourOf(Now))
                                                 , IntToStr(MinuteOf(Now))
                                                 , IntToStr(MilliSecondOf(Now))
                                                ]);
      EcritLog(TempPath, TslTra.Text, TmpFileName);
      FreeAndNil(TslTra);
    end;
    PathFileName := IncludeTrailingPathDelimiter(TempPath) + TmpFileName;
    Tools.CompressFile(PathFileName);
    TSLTRAFileQty.Clear;
    Tools.FileCut(PathFileName, Tools.GetKoFromMo(3), TSLTRAFileQty);
    if TGetParamWSCEGID.ConnectToY2 then
    begin
      SendEntry := TSendEntryY2.Create;
      try
        SendEntry.ServerName := '';
        SendEntry.DBName     := '';
        SendEntry.SendAccountingParameters(TSLTRAFileQty);
      finally
        SendEntry.Free;
      end;
    end;
    DeleteFile(PathFileName);
  end;
end;


procedure TSvcSyncBTPY2Execute.SearchRelatedParameters(TSlEntry: TStringList);
var
  AdoQryParam : AdoQry;
  Cpt         : integer;
  Values      : string;
  KeyValue1   : string;
  KeyValue2   : string;
begin
  AdoQryParam := AdoQry.Create;
  try
    AdoQryParam.ServerName := AdoQryBTP.ServerName;
    AdoQryParam.DBName := AdoQryBTP.DBName;
    for Cpt := 0 to pred(TSlEntry.count) do
    begin
      Values := TSlEntry[Cpt];
      if Copy(Values, 1, Length(ToolsTobToTsl_LevelName) + 1) = ToolsTobToTsl_LevelName + '2' then // Export tiers et rib de l'écriture courante
      begin
        KeyValue1 := Tools.GetStValueFromTSl(Values, 'E_AUXILIAIRE');
        KeyValue2 := '';
        AddUpdateValues(AdoQryParam, wsdsThird, GetFieldsList(TSlCacheThirdBTP), KeyValue1, KeyValue2, True);
        AddUpdateValues(AdoQryParam, wsdsBankIdentification, GetFieldsList(TSlCacheBankIdBTP), KeyValue1, KeyValue2, True);
        AddUpdateValues(AdoQryParam, wsdsContact, GetFieldsList(TSlcacheContactBTP), KeyValue1, KeyValue2, True);
      end
      else if Copy(Values, 1, Length(ToolsTobToTsl_LevelName) + 1) = ToolsTobToTsl_LevelName + '4' then // Export de la section de l'analytique courante
      begin
        KeyValue1 := Tools.GetStValueFromTSl(Values, 'Y_AXE');
        KeyValue2 := Tools.GetStValueFromTSl(Values, 'Y_SECTION');
        AddUpdateValues(AdoQryParam, wsdsAnalyticalSection, GetFieldsList(TSlCacheSectionBTP), KeyValue1, KeyValue2, True);
      end;
    end;
  finally
    AdoQryParam.Free;
  end;
end;

procedure TSvcSyncBTPY2Execute.SearchOthersParameters;
var
  AdoQryParam: AdoQry;

  function GetKeyValue(DSType: T_WSDataService): string;
  var
    MultipleValues: string;
  begin
    if DSType = wsdsChoixCod then
    begin
      MultipleValues := Tools.GetExtractTypeFromTableName(GetTableNameFromDSType(DSType));
      while MultipleValues <> '' do
        Result := Result + ',''' + Tools.ReadTokenSt_(MultipleValues, ';') + '''';
      Result := copy(Result, 2, Length(Result));
    end
    else
      Result := '';
  end;

begin
  AdoQryParam := AdoQry.Create;
  try
    AdoQryParam.ServerName := AdoQryBTP.ServerName;
    AdoQryParam.DBName := AdoQryBTP.DBName;
    AddUpdateValues(AdoQryParam, wsdsChoixCod    , GetFieldsList(TSlCacheChoixCodBTP)  , GetKeyValue(wsdsChoixCod)    , '', False); // ChoixCod
    AddUpdateValues(AdoQryParam, wsdsCurrency    , GetFieldsList(TSlCacheCurrencyBTP)  , GetKeyValue(wsdsCurrency)    , '', False); // Devise
    AddUpdateValues(AdoQryParam, wsdsPaymenChoice, GetFieldsList(TSlCachePaymentBTP)   , GetKeyValue(wsdsPaymenChoice), '', False); // Mode de règlement
    AddUpdateValues(AdoQryParam, wsdsChangeRate  , GetFieldsList(TSlCacheChangeRateBTP), GetKeyValue(wsdsChangeRate)  , '', False); // Taux de change
  finally
    AdoQryParam.Free;
  end;
end;

procedure TSvcSyncBTPY2Execute.WriteLog(TypeDebug: TSvcSyncBTPY2Log; Text: string; Level: integer);
var
  LogText    : string;
  WindowsLog : TEventLogger;
begin
  case TypeDebug of
    ssbylLog:
      begin
        if LogValues.LogLevel > 0 then
        begin
          LogText := Format('%s : %s%s', [DateTimeToStr(Now), StringOfChar(' ', Level), Text]);
          Writeln(LogFile, LogText);
        end;
      end;
    ssbylWindows:
      begin
        WindowsLog := TEventLogger.Create(ApplicationName);
        try
          WindowsLog.LogMessage(Text, EVENTLOG_INFORMATION_TYPE);
        finally
          WindowsLog.Free;
        end;
      end;
  end;
end;

function TSvcSyncBTPY2Execute.GetValueFormSettingLine(LineValue: string): string;
begin
  Result := copy(LineValue, Pos('=', LineValue) + 1, Length(LineValue));
end;

procedure TSvcSyncBTPY2Execute.SetFilterFromDSType_Get(DSType: T_WSDataService; TSl: TStringList);
var
  sCode     : string;
  Prefix    : string;
  TableName : string;
  First     : Boolean;
begin
  TableName := GetTableNameFromDSType(DSType);
  case DSType of
    wsdsThird:
      begin
        sCode := Tools.GetExtractTypeFromTableName(TableName);
        First := True;
        TSlFilter.Add(';;;(');
        while sCode <> '' do
        begin
          TSl.Add(Tools.iif(First, '', 'OR') + ';T_NATUREAUXI;=;' + Tools.ReadTokenSt_(sCode, ';'));
          First := False;
        end;
        TSlFilter.Add(';;;)');
        TSl.Add('AND;T_DATEMODIF;>=;' + BTPValues.LastSynchro);
      end;
    wsdsAnalyticalSection:
      TSl.Add(';CSP_DATEMODIF;>=;' + BTPValues.LastSynchro);
    wsdsAccount:
      TSl.Add(';G_DATEMODIF;>=;' + BTPValues.LastSynchro);
    wsdsJournal:
      TSl.Add(';J_DATEMODIF;>=;' + BTPValues.LastSynchro);
    wsdsBankIdentification:
      TSl.Add(';R_DATEMODIF;>=;' + BTPValues.LastSynchro);
    wsdsChoixCod, wsdsCommon:
      begin
        Prefix := Tools.iif(DSType = wsdsChoixCod, 'CC', 'CO');
        sCode := Tools.iif(DSType = wsdsChoixCod, Tools.GetExtractTypeFromTableName(TableName), Tools.GetExtractTypeFromTableName(TableName));
        First := True;
        while sCode <> '' do
        begin
          TSl.Add(Tools.iif(First, '', 'OR') + ';' + Prefix + '_TYPE;=;' + Tools.ReadTokenSt_(sCode, ';'));
          First := False;
        end;
      end;
    wsdsRecovery:
      begin
        sCode := Tools.GetExtractTypeFromTableName(TableName);
        First := True;
        while sCode <> '' do
        begin
          TSl.Add(Tools.iif(First, '', 'OR') + ';RR_TYPERELANCE;=;' + Tools.ReadTokenSt_(sCode, ';'));
          First := False;
        end;
      end;
    wsdsCountry:
      TSl.Add(WSCDS_EmptyValue);
    wsdsCurrency:
      TSl.Add(WSCDS_EmptyValue);
    wsdsCorrespondence:
      TSl.Add(WSCDS_EmptyValue);
    wsdsPaymenChoice:
      TSl.Add(WSCDS_EmptyValue);
    wsdsChangeRate:
      TSl.Add(';H_DATECOURS;>=;' + BTPValues.LastSynchro);
    wsdsFiscalYear:
      TSl.Add(WSCDS_EmptyValue);
    wsdsSocietyParameters:
      TSl.Add(WSCDS_EmptyValue);
    wsdsEstablishment:
      TSl.Add(';ET_DATEMODIF;>=;' + BTPValues.LastSynchro);
    wsdsPaymentMode:
      TSl.Add(WSCDS_EmptyValue);
    wsdsZipCode:
      TSl.Add(WSCDS_EmptyValue);
    wsdsContact:
      begin
        TSl.Add(';C_DATEMODIF;>=;' + BTPValues.LastSynchro);
        TSl.Add('AND;C_NATUREAUXI;=;CLI');
      end;
  end;
end;

procedure TSvcSyncBTPY2Execute.SetFilterFromDSType_Set(DSType: T_WSDataService; Value, Value1, Value2: string; TSl: TStringList);
var
  sCode     : string;
  TableName : string;
  First     : Boolean;
begin
(*
      wsdsThird              : Result := 'WHERE T_AUXILIAIRE = ''%s''', [lAdoQry.FieldsList, KeyValue1]); // Format('SELECT %s FROM TIERS WHERE T_AUXILIAIRE = ''%s'' AND T_DATEMODIF > ''%s''', [lAdoQry.FieldsList, KeyValue1, BTPValues.LastSynchro]);
      wsdsBankIdentification : Result := 'WHERE R_AUXILIAIRE = ''%s''', [lAdoQry.FieldsList, KeyValue1]); // Format('SELECT %s FROM RIB WHERE R_AUXILIAIRE = ''%s'' AND R_DATEMODIF > ''%s''', [lAdoQry.FieldsList, KeyValue1, BTPValues.LastSynchro]);
      wsdsAnalyticalSection  : Result := 'WHERE S_AXE = ''%s'' AND S_SECTION = ''%s''', [lAdoQry.FieldsList, KeyValue1, KeyValue2]); // AdoQryParam.Request    := Format('SELECT %s FROM SECTION WHERE S_AXE = ''%s'' AND S_SECTION = ''%s'' AND S_DATEMODIF > ''%s''', [lAdoQry.FieldsList, KeyValue1, KeyValue2, BTPValues.LastSynchro]);
      wsdsChoixCod           : Result := 'WHERE CC_TYPE IN (%s) ORDER BY CC_TYPE, CC_CODE', [lAdoQry.FieldsList, KeyValue1]);
      wsdsCurrency           : Result := Format('SELECT %s FROM DEVISE ORDER BY D_DEVISE', [lAdoQry.FieldsList]);
      wsdsPaymenChoice       : Result := Format('SELECT %s FROM MODEREGL ORDER BY MR_MODEREGLE', [lAdoQry.FieldsList]);
      wsdsChangeRate         : Result := Format('SELECT %s FROM CHANCELL ORDER BY H_DEVISE, H_DATECOURS', [lAdoQry.FieldsList]);
*)
  TableName := GetTableNameFromDSType(DSType);
  case DSType of
    wsdsThird:
      TSl.Add(';T_AUXILIAIRE;=;' + Value);
    wsdsAnalyticalSection:
      begin
        TSl.Add(';S_AXE=;' + Value);
        TSl.Add('AND;S_SECTION;=;' + Value1);
      end;
    wsdsBankIdentification:
      begin
        TSl.Add(';R_AUXILIAIRE;=;' + Value);
        TSl.Add('AND;R_NUMERORIB;=;' + Value1);
      end;
    wsdsChoixCod:
      begin
        TSl.Add(';CC_TYPE;=;' + Value);
        TSl.Add('AND;CC_CODE;=;' + Value1);
      end;
    wsdsCommon:
      begin
        TSl.Add(';CO_TYPE;=;' + Value);
        TSl.Add('AND;CO_CODE;=;' + Value1);
      end;
    wsdsRecovery:
      begin
        sCode := Tools.GetExtractTypeFromTableName(TableName);
        First := True;
        while sCode <> '' do
        begin
          TSl.Add(Tools.iif(First, '', 'OR') + ';RR_TYPERELANCE;=;' + Tools.ReadTokenSt_(sCode, ';'));
          First := False;
        end;
      end;
    wsdsCountry:
      TSl.Add(WSCDS_EmptyValue);
    wsdsCurrency:
      TSl.Add(WSCDS_EmptyValue);
    wsdsCorrespondence:
      TSl.Add(WSCDS_EmptyValue);
    wsdsPaymenChoice:
      TSl.Add(WSCDS_EmptyValue);
    wsdsChangeRate:
      TSl.Add(';H_DATECOURS;>=;' + BTPValues.LastSynchro);
    wsdsSocietyParameters:
      TSl.Add(WSCDS_EmptyValue);
    wsdsEstablishment:
      TSl.Add(';ET_DATEMODIF;>=;' + BTPValues.LastSynchro);
    wsdsPaymentMode:
      TSl.Add(WSCDS_EmptyValue);
    wsdsZipCode:
      TSl.Add(WSCDS_EmptyValue);
    wsdsContact:
      begin
        TSl.Add(';C_DATEMODIF;>=;' + BTPValues.LastSynchro);
        TSl.Add('AND;C_NATUREAUXI;=;CLI');
      end;
  end;
end;

function TSvcSyncBTPY2Execute.GetInfoFromDSType(InfoType: T_WSInfoFromDSType; DSType: T_WSDataService; FieldName: string = ''): string;
var
  TableName : string;
  Value     : string;
  Value1    : string;
  Value2    : string;

  function GetValueFrom(lTSlIndice: TStringList; FieldName: string): string;
  var
    iCpt    : integer;
    lValues : string;
  begin
    Result := '';
    for iCpt := 0 to Pred(lTSlIndice.Count) do
    begin
      lValues := lTSlIndice[iCpt];
      if Copy(lValues, 1, Pos('=', lValues) - 1) = FieldName then
      begin
        Result := Copy(lValues, Pos('=', lValues) + 1, Length(lValues));
        Break;
      end;
    end;
  end;

begin
  TableName := GetTableNameFromDSType(DSType);
  case DSType of
    wsdsThird:
      begin
        case InfoType of
          wsidTableName     : Result := TableName;
          wsidFieldsKey     : Result := 'T_AUXILIAIRE';
          wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';T_NATUREAUXI;T_AUXILIAIRE;T_TIERS;') > 0));
          wsidFieldsList    : Result := 'T_AUXILIAIRE';
          wsidRequest       : Result := Format('SELECT %s FROM %s WHERE T_AUXILIAIRE = ''%s''', [AdoQryBTP.FieldsList, TableName, GetValueFrom(TSlIndice, AdoQryBTP.FieldsList)]);
        else
          Result := '';
        end;
      end;
    wsdsAnalyticalSection:
      begin
        case InfoType of
          wsidTableName     : Result := TableName;
          wsidFieldsKey     : Result := 'S_AXE;S_SECTION';
          wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';S_AXE;S_SECTION;') > 0));
          wsidFieldsList    : Result := 'S_SECTION';
          wsidRequest:
            begin
              Value := GetValueFrom(TSlIndice, 'S_AXE');
              Value1 := GetValueFrom(TSlIndice, 'S_SECTION');
              Result := Format('SELECT S_SECTION FROM %s WHERE S_AXE = ''%s'' AND S_SECTION = ''%s''', [TableName, Value, Value1]);
            end;
        else
          Result := '';
        end;
      end;
    wsdsAccount:
      begin
        case InfoType of
          wsidTableName     : Result := TableName;
          wsidFieldsKey     : Result := 'G_GENERAL';
          wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';G_GENERAL;') > 0));
          wsidFieldsList    : Result := 'G_GENERAL';
          wsidRequest       : Result := Format('SELECT %s FROM %s WHERE G_GENERAL = ''%s''', [AdoQryBTP.FieldsList, TableName, GetValueFrom(TSlIndice, AdoQryBTP.FieldsList)]);
        else
          Result := '';
        end;
      end;
    wsdsJournal:
      begin
        case InfoType of
          wsidTableName     : Result := TableName;
          wsidFieldsKey     : Result := 'J_JOURNAL';
          wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';J_JOURNAL;') > 0));
          wsidFieldsList    : Result := 'J_JOURNAL';
          wsidRequest       : Result := Format('SELECT J_JOURNAL FROM %s WHERE J_JOURNAL = ''%s''', [TableName, GetValueFrom(TSlIndice, 'J_JOURNAL')]);
        else
          Result := '';
        end;
      end;
    wsdsBankIdentification:
      begin
        case InfoType of
          wsidTableName     : Result := TableName;
          wsidFieldsKey     : Result := 'R_AUXILIAIRE;R_NUMERORIB';
          wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';R_AUXILIAIRE;R_NUMERORIB;') > 0));
          wsidFieldsList    : Result := 'R_AUXILIAIRE';
          wsidRequest:
            begin
              Value  := GetValueFrom(TSlIndice, 'R_AUXILIAIRE');
              Value1 := GetValueFrom(TSlIndice, 'R_NUMERORIB');
              Result := Format('SELECT R_AUXILIAIRE FROM %s WHERE R_AUXILIAIRE = ''%s'' AND R_NUMERORIB = ''%s''', [TableName, Value, Value1]);
            end;
        else
          Result := '';
        end;
      end;
    wsdsChoixCod:
      begin
        case InfoType of
          wsidTableName     : Result := TableName;
          wsidFieldsKey     : Result := 'CC_TYPE;CC_CODE';
          wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';CC_TYPE;CC_CODE;') > 0));
          wsidFieldsList    : Result := 'CC_TYPE';
          wsidRequest:
            begin
              Value  := GetValueFrom(TSlIndice, 'CC_TYPE');
              Value1 := GetValueFrom(TSlIndice, 'CC_CODE');
              Result := Format('SELECT CC_TYPE FROM %s WHERE CC_TYPE = ''%s'' AND CC_CODE = ''%s''', [TableName, Value, Value1]);
            end;
        else
          Result := '';
        end;
      end;
    wsdsCommon:
      begin
        case InfoType of
          wsidTableName     : Result := TableName;
          wsidFieldsKey     : Result := 'CO_TYPE;CO_CODE';
          wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';CO_TYPE;CO_CODE;') > 0));
          wsidFieldsList    : Result := 'CO_TYPE';
          wsidRequest:
            begin
              Value  := GetValueFrom(TSlIndice, 'CO_TYPE');
              Value1 := GetValueFrom(TSlIndice, 'CO_CODE');
              Result := Format('SELECT CO_TYPE FROM %s WHERE CO_TYPE = ''%s'' AND CO_CODE = ''%s''', [TableName, Value, Value1]);
            end;
        else
          Result := '';
        end;
      end;
    wsdsRecovery:
      begin
        case InfoType of
          wsidTableName     : Result := TableName;
          wsidFieldsKey     : Result := 'RR_TYPERELANCE;RR_FAMILLERELANCE';
          wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';RR_TYPERELANCE;RR_FAMILLERELANCE;') > 0));
          wsidFieldsList    : Result := 'RR_TYPERELANCE';
          wsidRequest:
            begin
              Value  := GetValueFrom(TSlIndice, 'RR_TYPERELANCE');
              Value1 := GetValueFrom(TSlIndice, 'RR_FAMILLERELANCE');
              Result := Format('SELECT RR_TYPERELANCE FROM %s WHERE RR_TYPERELANCE = ''%s'' AND RR_FAMILLERELANCE = ''%s''', [TableName, Value, Value1]);
            end;
        else
          Result := '';
        end;
      end;
    wsdsCountry:
      begin
        case InfoType of
          wsidTableName     : Result := TableName;
          wsidFieldsKey     : Result := 'PY_PAYS';
          wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';PY_PAYS;') > 0));
          wsidFieldsList    : Result := 'PY_PAYS';
          wsidRequest       : Result := Format('SELECT PY_PAYS FROM %s WHERE PY_PAYS = ''%s''', [TableName, GetValueFrom(TSlIndice, 'PY_PAYS')]);
        else
          Result := '';
        end;
      end;
    wsdsCurrency:
      begin
        case InfoType of
          wsidTableName     : Result := TableName;
          wsidFieldsKey     : Result := 'D_DEVISE';
          wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';D_DEVISE;') > 0));
          wsidFieldsList    : Result := 'D_DEVISE';
          wsidRequest       : Result := Format('SELECT D_DEVISE FROM %s WHERE D_DEVISE = ''%s''', [TableName, GetValueFrom(TSlIndice, 'D_DEVISE')]);
        else
          Result := '';
        end;
      end;
    wsdsCorrespondence:
      begin
        case InfoType of
          wsidTableName     : Result := TableName;
          wsidFieldsKey     : Result := 'CR_TYPE;CR_CORRESP';
          wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';CR_TYPE;CR_CORRESP;') > 0));
          wsidFieldsList    : Result := 'CR_CORRESP';
          wsidRequest:
            begin
              Value  := GetValueFrom(TSlIndice, 'CR_TYPE');
              Value1 := GetValueFrom(TSlIndice, 'CR_CORRESP');
              Result := Format('SELECT CR_CORRESP FROM %s WHERE CR_TYPE = ''%s'' AND CR_CORRESP = ''%s''', [TableName, Value, Value1]);
            end;
        else
          Result := '';
        end;
      end;
    wsdsPaymenChoice:
      begin
        case InfoType of
          wsidTableName     : Result := TableName;
          wsidFieldsKey     : Result := 'MR_MODEREGLE';
          wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';MR_MODEREGLE;') > 0));
          wsidFieldsList    : Result := 'MR_MODEREGLE';
          wsidRequest       : Result := Format('SELECT MR_MODEREGLE FROM %s WHERE MR_MODEREGLE = ''%s''', [TableName, GetValueFrom(TSlIndice, 'MR_MODEREGLE')]);
        else
          Result := '';
        end;
      end;
    wsdsChangeRate:
      begin
        case InfoType of
          wsidTableName     : Result := TableName;
          wsidFieldsKey     : Result := 'H_DEVISE;H_DATECOURS';
          wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';H_DEVISE;H_DATECOURS;') > 0));
          wsidFieldsList    : Result := 'H_DATECOURS';
          wsidRequest:
            begin
              Value  := GetValueFrom(TSlIndice, 'H_DEVISE');
              Value1 := GetValueFrom(TSlIndice, 'H_DATECOURS');
              Result := Format('SELECT H_DEVISE FROM %s WHERE H_DEVISE = ''%s'' AND H_DATECOURS = ''%s''', [TableName, Value, Value1]);
            end;
        else
          Result := '';
        end;
      end;
    wsdsFiscalYear:
      begin
        case InfoType of
          wsidTableName     : Result := TableName;
          wsidFieldsKey     : Result := 'EX_ENTITY;EX_EXERCICE';
          wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';EX_ENTITY;EX_EXERCICE;') > 0));
          wsidFieldsList    : Result := 'EX_EXERCICE';
          wsidRequest:
            begin
              Value  := GetValueFrom(TSlIndice, 'EX_ENTITY');
              Value1 := GetValueFrom(TSlIndice, 'EX_EXERCICE');
              Result := Format('SELECT EX_EXERCICE FROM %s WHERE EX_ENTITY = ''%s'' AND EX_EXERCICE = ''%s''', [TableName, Value, Value1]);
            end;
        else
          Result := '';
        end;
      end;
    wsdsSocietyParameters:
      begin
        case InfoType of
          wsidTableName     : Result := TableName;
          wsidFieldsKey     : Result := 'SOC_NOM';
          wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';SOC_NOM;') > 0));
          wsidFieldsList    : Result := 'SOC_NOM';
          wsidRequest       : Result := Format('SELECT SOC_NOM FROM %s WHERE SOC_NOM = ''%s''', [TableName, GetValueFrom(TSlIndice, 'SOC_NOM')]);
        else
          Result := '';
        end;
      end;
    wsdsEstablishment:
      begin
        case InfoType of
          wsidTableName     : Result := TableName;
          wsidFieldsKey     : Result := 'ET_ETABLISSEMENT';
          wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';ET_ETABLISSEMENT;') > 0));
          wsidFieldsList    : Result := 'ET_ETABLISSEMENT';
          wsidRequest       : Result := Format('SELECT ET_ETABLISSEMENT FROM %s WHERE ET_ETABLISSEMENT = ''%s''', [TableName, GetValueFrom(TSlIndice, 'ET_ETABLISSEMENT')]);
        else
          Result := '';
        end;
      end;
    wsdsPaymentMode:
      begin
        case InfoType of
          wsidTableName     : Result := TableName;
          wsidFieldsKey     : Result := 'MP_MODEPAIE';
          wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';MP_MODEPAIE;') > 0));
          wsidFieldsList    : Result := 'MP_MODEPAIE';
          wsidRequest       : Result := Format('SELECT MP_MODEPAIE FROM %s WHERE MP_MODEPAIE = ''%s''', [TableName, GetValueFrom(TSlIndice, 'MP_MODEPAIE')]);
        else
          Result := '';
        end;
      end;
    wsdsZipCode:
      begin
        case InfoType of
          wsidTableName     : Result := TableName;
          wsidFieldsKey     : Result := 'O_CODEPOSTAL;O_VILLE';
          wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';O_CODEPOSTAL;O_VILLE;') > 0));
          wsidFieldsList    : Result := 'O_CODEPOSTAL';
          wsidRequest:
            begin
              Value  := GetValueFrom(TSlIndice, 'O_CODEPOSTAL');
              Value1 := GetValueFrom(TSlIndice, 'O_VILLE');
              if Pos('''', Value1) > 0 then
                Value1 := StringReplace(Value1, '''', '''''', [rfReplaceAll]);
              Result := Format('SELECT O_CODEPOSTAL FROM %s WHERE O_CODEPOSTAL = ''%s'' AND O_VILLE = ''%s''', [TableName, Value, Value1]);
            end;
        else
          Result := '';
        end;
      end;
    wsdsContact:
      begin
        case InfoType of
          wsidTableName     : Result := TableName;
          wsidFieldsKey     : Result := 'C_TYPECONTACT;C_AUXILIAIRE;C_NUMEROCONTACT';
          wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';C_TYPECONTACT;C_AUXILIAIRE;C_NUMEROCONTACT;') > 0));
          wsidFieldsList    : Result := 'C_NUMEROCONTACT';
          wsidRequest:
            begin
              Value  := GetValueFrom(TSlIndice, 'C_TYPECONTACT');
              Value1 := GetValueFrom(TSlIndice, 'C_AUXILIAIRE');
              Value2 := GetValueFrom(TSlIndice, 'C_NUMEROCONTACT');
              Result := Format('SELECT C_NUMEROCONTACT FROM %s WHERE C_TYPECONTACT = ''%s'' AND C_AUXILIAIRE = ''%s'' AND C_NUMEROCONTACT = ''%s''', [TableName, Value, Value1, Value2]);
            end;
        else
          Result := '';
        end;
      end;
  else
    Result := '';
  end;
end;

procedure TSvcSyncBTPY2Execute.ExtractIndice(Cpt: integer; TSlOrig, TSlResult: TStringList);
var
  sIndex      : string;
  Value       : string;
  Pos         : integer;
  iCpt        : integer;
  IndexChange : Boolean;
begin
  sIndex := WSCDS_IndiceField + IntToStr(Cpt) + '#';
  Pos := TSlOrig.IndexOfName(sIndex);
  if Pos > -1 then
  begin
    sIndex := TSlOrig[Pos];
    for iCpt := Pos to Pred(TSlOrig.Count) do
    begin
      Value := TSlOrig[iCpt];
      IndexChange := ((Copy(Value, 1, 7) = WSCDS_IndiceField) and (Value <> sIndex));
      if not IndexChange then
      begin
        { Spécif pour SECTION, préfixe Y2=CSP_, préfixe BTP=S_ }
        if Copy(Value, 1, 4) = 'CSP_' then
          Value := StringReplace(Value, 'CSP_', 'S_', [rfReplaceAll]);
        TSlResult.Add(Value);
      end
      else
        Break;
    end;
  end;
end;

function TSvcSyncBTPY2Execute.AddData(wsAction: T_WSAction; DSType: T_WSDataService; lTSlValues: TStringList): boolean;
var
  CptUI             : integer;
  KeyValue          : string;
  KeyValuel         : string;
  FieldName         : string;
  FieldValue        : string;
  Sql               : string;
  Where             : string;
  InsertedFields    : string;
  InsertedValues    : string;
  CacheLine         : string;
  TSlInsertedFields : TStringList;
  FieldType         : tTypeField;
begin
  Result         := True;
  KeyValue       := GetKeyValues(DSType, lTSlValues);
  InsertedFields := '';
  InsertedValues := '';
  if LogValues.LogLevel = 2 then
    WriteLog(ssbylLog, Format('%s de "%s"', [Tools.iif(wsAction = wsacUpdate, 'Modification', 'Création'), KeyValue]), 4);
  for CptUI := 0 to pred(lTSlValues.Count) do
  begin
    if copy(lTSlValues[CptUI], 1, 7) <> WSCDS_IndiceField then
    begin
      FieldName := copy(lTSlValues[CptUI], 1, Pos('=', lTSlValues[CptUI]) - 1);
      FieldType := GetFieldTypeFromCache(DSType, FieldName);
      if     (FieldExistInTSl(GetTSlFromDSType(DSType, True), FieldName))              // Le champ existe dans BTP
         and (   (wsAction = wsacInsert)                                               // Insert
              or (    (wsAction = wsacUpdate)                                          // Update
                  and (GetInfoFromDSType(wsidExcludeFields, DSType, FieldName) = '0')) //  et le champs ne fait pas parti des champs à exclure
             )
      then
      begin
        if Pos('_DATEMODIF', FieldName) > 0 then
          FieldValue := DateTimeToStr(Now)
        else
        begin
          FieldValue := copy(lTSlValues[CptUI], Pos('=', lTSlValues[CptUI]) + 1, length(lTSlValues[CptUI]));
          case FieldType of
            ttfDate    : FieldValue := Tools.SetStrDateTimeFromStrUTCDateTime(FieldValue);
            ttfCombo   : FieldValue := Trim(FieldValue);
            ttfNumeric : FieldValue := StringReplace(FieldValue, ',', '.', [rfReplaceAll]);
          end;
        end;
        if FieldType = ttfDate then
          FieldValue := FormatDateTime('yyyymmdd hh:nn:ss', StrToDateTime(FieldValue));
        FieldValue     := AddQuotes(DSType, FieldName, FieldValue);
        Sql            := Sql + Format(', %s=%s', [FieldName, FieldValue]);
        InsertedFields := InsertedFields + ', ' + FieldName;
        InsertedValues := InsertedValues + ', ' + FieldValue;
      end;
    end;
  end;
  if wsAction = wsacInsert then  // Ajouter tous les autres champs
  begin
    TSlInsertedFields := GetTSlFromDSType(DSType, True);
    if Assigned(TSlInsertedFields) then
    begin
      for CptUI := 0 to pred(TSlInsertedFields.Count) do
      begin
        CacheLine := TSlInsertedFields[CptUI];
        FieldName := Tools.ReadTokenSt_(CacheLine, ToolsTobToTsl_Separator);
        Tools.ReadTokenSt_(CacheLine, ToolsTobToTsl_Separator);
        FieldValue := Tools.ReadTokenSt_(CacheLine, ToolsTobToTsl_Separator);
        if Pos(FieldName, InsertedFields) = 0 then
        begin
          InsertedFields := InsertedFields + ', ' + FieldName;
          InsertedValues := InsertedValues + ', ' + FieldValue;
        end;
      end;
    end;
  end;
  Sql := Copy(Sql, 2, Length(Sql));
  InsertedFields := Copy(InsertedFields, 2, Length(InsertedFields));
  InsertedValues := Copy(InsertedValues, 2, Length(InsertedValues));
  if wsAction = wsacUpdate then
  begin
    KeyValuel := KeyValue;
    while KeyValuel <> '' do
      Where := Where + ' AND ' + Tools.ReadTokenSt_(KeyValuel, '-');
    Where := Copy(Where, 5, length(Where));
    Sql := Format('UPDATE %s SET %s WHERE %s', [GetInfoFromDSType(wsidTableName, DSType), Sql, Where]);
  end else
    Sql := Format('INSERT INTO %s (%s) VALUES(%s)', [GetInfoFromDSType(wsidTableName, DSType), InsertedFields, InsertedValues]);
  TSLUpdateInsertData.Add(Sql);
end;

function TSvcSyncBTPY2Execute.GetKeyValues(DSType: T_WSDataService; lTSlValues: TStringList): string;
var
  FieldsKey  : string;
  Value      : string;
  FieldName  : string;
  FieldValue : string;
  Posit      : integer;
begin
  Result := '';
  FieldsKey := GetInfoFromDSType(wsidFieldsKey, DSType);
  while FieldsKey <> '' do
  begin
    Posit := lTSlValues.IndexOfName(Tools.ReadTokenSt_(FieldsKey, ';'));
    if Posit > -1 then
    begin
      FieldName  := copy(lTSlValues[Posit], 1, Pos('=', lTSlValues[Posit]) - 1);
      FieldValue := copy(lTSlValues[Posit], Pos('=', lTSlValues[Posit]) + 1, length(lTSlValues[Posit]));
      if GetFieldTypeFromCache(DSType, FieldName) = ttfCombo then
        FieldValue := Trim(FieldValue);
      FieldValue := AddQuotes(DSType, FieldName, FieldValue);
      Value      := Format('-%s=%s', [FieldName, FieldValue]);
      Result     := Result + Value;
    end;
  end;
  Result := Copy(Result, 2, Length(Result));
end;

function TSvcSyncBTPY2Execute.GetIndiceY2DataCache(DSType: T_WSDataService; TSLValue: TStringList): string;
var
  TableName : string;
  FieldsKey : string;
  FieldKey  : string;
  Index     : Integer;
begin
  TableName := GetInfoFromDSType(wsidTableName, DSType);
  FieldsKey := GetInfoFromDSType(wsidFieldsKey, DSType);
  Result    := TableName + '_';
  while FieldsKey <> '' do
  begin
    FieldKey := Tools.ReadTokenSt_(FieldsKey, ';');
    Index := TSLValue.IndexOfName(FieldKey);
    if Index > -1 then
      Result := Result + Copy(TSLValue[Index], Pos('=', TSLValue[Index]) + 1, Length(TSLValue[Index]));
  end;
end;

function TSvcSyncBTPY2Execute.AddQuotes(DSType: T_WSDataService; FieldName, FieldValue: string): string;
begin
  if Pos('''', FieldValue) > 0 then
    FieldValue := StringReplace(FieldValue, '''', '''''', [rfReplaceAll]);
  case GetFieldTypeFromCache(DSType, FieldName) of
    ttfBoolean
    , ttfCombo
    , ttfMemo
    , ttfText
    , ttfDate : Result := Format('''%s''', [FieldValue]);
  else
    Result := FieldValue;
  end;
end;

function TSvcSyncBTPY2Execute.GetY2Data(DSType: T_WSDataService): boolean;
var
  Cpt          : integer;
  Index        : Integer;
  ModifyQty    : integer;
  InsertQty    : integer;
  TableName    : string;
  GetDataState : string;

  procedure SetCacheY2Data;
  var
    IndexValue  : string;
    RecordValue : string;
    Index       : integer;
  begin
    if TSlIndice.Count > 0 then
    begin
      IndexValue := GetIndiceY2DataCache(DSType, TSlIndice) + '=';
      RecordValue := '';
      for Index := 0 to pred(TSlIndice.Count) do
      begin
        if Copy(TSlIndice[Index], 1, 7) <> WSCDS_IndiceField then
          RecordValue := RecordValue + ToolsTobToTsl_Separator + TSlIndice[Index]; //copy(TSlIndice[Index], pos('=', TSlIndice[Index]) +1, length(TSlIndice[Index]));
      end;
      RecordValue := Copy(RecordValue, 2, length(RecordValue));
      TSlCacheGetY2Data.Add(IndexValue + RecordValue);
    end;
  end;

begin
  TableName := GetInfoFromDSType(wsidTableName, DSType);
  ModifyQty := 0;
  InsertQty := 0;
  WriteLog(ssbylLog, TableName, 3);
  SetFilterFromDSType_Get(DSType, TSlFilter);
  GetDataState := TReadWSDataService.GetData(DSType, Y2Values.Server, Y2Values.DataBase, TSlValues, TSlCacheWSFields, TSlFilter);
  Result := (GetDataState = WSCDS_GetDataOk);
  if Result then
  begin
    if TSlValues.Count > 0 then
    begin
      {$IF defined(APPSRV)}
      for Cpt := 0 to Pred(TSlValues.Count) do
      begin
        if Copy(TSlValues[Cpt], 1, 7) = WSCDS_IndiceField then
        begin
          Index := StrToInt(Copy(TSlValues[Cpt], Pos('=', TSlValues[Cpt]) + 1, Length(TSlValues[Cpt])));
          TSlIndice.Clear;
          ExtractIndice(Index, TSlValues, TSlIndice);
          { Met la valeur trouvée en cache pour comparer si les valeurs ont changées lors du Send vers Y2 }
          SetCacheY2Data;         
          { Test si existe déjà }
          AdoQryBTP.TSLResult.Clear;
          AdoQryBTP.FieldsList := GetInfoFromDSType(wsidFieldsList, DSType);
          AdoQryBTP.Request := GetInfoFromDSType(wsidRequest, DSType);
          AdoQryBTP.SingleTableSelect;
          if AdoQryBTP.RecordCount > 0 then
          begin
            Inc(ModifyQty);
            Result := AddData(wsacUpdate, DSType, TSlIndice);
          end else
          begin
            Inc(InsertQty);
            Result := AddData(wsacInsert, DSType, TSlIndice);
          end;
          AdoQryBTP.TSLResult.Clear;
          AdoQryBTP.RecordCount := 0;
        end;
      end;
      {$ELSE  !APPSRV}
      {$IFEND !APPSRV}
    end;
    WriteLog(ssbylLog, Format('%s enregistrement(s) à modifier', [IntToStr(ModifyQty)]), 4);
    WriteLog(ssbylLog, Format('%s enregistrement(s) à créer', [IntToStr(InsertQty)]), 4);
  end
  else
    WriteLog(ssbylLog, Format('*** ERREUR : %s', [GetDataState]), 0);
  TSlFilter.Clear;
  TSlValues.Clear;
end;

procedure TSvcSyncBTPY2Execute.CreateObjects;
begin
  TSlConnectionValues := TstringList.Create;
  TSlValues                      := TstringList.Create;
  TSlFilter                      := TStringList.Create;
  TSlIndice                      := TStringList.Create;
  TSlCacheThirdBTP               := TStringList.Create;
  TSlCacheSectionBTP             := TStringList.Create;
  TSlCacheAcountBTP              := TStringList.Create;
  TSlCachePaymentBTP             := TStringList.Create;
  TSlCacheCorrespBTP             := TStringList.Create;
  TSlCacheCurrencyBTP            := TStringList.Create;
  TSlCacheCountryBTP             := TStringList.Create;
  TSlCacheRecoveryBTP            := TStringList.Create;
  TSlCacheCommonBTP              := TStringList.Create;
  TSlCacheChoixCodBTP            := TStringList.Create;
  TSlCacheJournalBTP             := TStringList.Create;
  TSlCacheBankIdBTP              := TStringList.Create;
  TSlCacheChangeRateBTP          := TStringList.Create;
  TSlCacheFiscalYearBTP          := TStringList.Create;
  TSlcacheSocietyParamBTP        := TStringList.Create;
  TSlcacheEstablishmentBTP       := TStringList.Create;
  TSlcachePaymentModeBTP         := TStringList.Create;
  TSlcacheZipCodeBTP             := TStringList.Create;
  TSlcacheContactBTP             := TStringList.Create;
  TSlCacheThirdY2                := TStringList.Create;
  TSlCacheSectionY2              := TStringList.Create;
  TSlCacheAcountY2               := TStringList.Create;
  TSlCachePaymentY2              := TStringList.Create;
  TSlCacheCorrespY2              := TStringList.Create;
  TSlCacheCurrencyY2             := TStringList.Create;
  TSlCacheCountryY2              := TStringList.Create;
  TSlCacheRecoveryY2             := TStringList.Create;
  TSlCacheCommonY2               := TStringList.Create;
  TSlCacheChoixCodY2             := TStringList.Create;
  TSlCacheJournalY2              := TStringList.Create;
  TSlCacheBankIdY2               := TStringList.Create;
  TSlCacheChangeRateY2           := TStringList.Create;
  TSlCacheFiscalYearY2           := TStringList.Create;
  TSlcacheSocietyParamY2         := TStringList.Create;
  TSlcacheEstablishmentY2        := TStringList.Create;
  TSlcachePaymentModeY2          := TStringList.Create;
  TSlcacheZipCodeY2              := TStringList.Create;
  TSlcacheContactY2              := TStringList.Create;
  TSlCacheWSFields               := TStringList.Create;
  TSlCacheSendAccParam           := TStringList.Create;
  TSlCacheGetY2Data              := TStringList.Create;
  TSLCacheUpdateFrequencySetting := TStringList.Create;
  TSLUpdateInsertData            := TStringList.Create;
  TSLTRAFileQty                  := TStringList.Create; 
  AdoQryBTP                      := AdoQry.Create;
  AdoQryY2                       := AdoQry.Create;
end;

procedure TSvcSyncBTPY2Execute.FreeObjects;
begin
  FreeAndNil(TSlConnectionValues);
  FreeAndNil(TSlValues);
  FreeAndNil(TSlFilter);
  FreeAndNil(TSlIndice);
  FreeAndNil(TSlCacheThirdBTP);
  FreeAndNil(TSlCacheSectionBTP);
  FreeAndNil(TSlCacheAcountBTP);
  FreeAndNil(TSlCachePaymentBTP);
  FreeAndNil(TSlCacheCorrespBTP);
  FreeAndNil(TSlCacheCurrencyBTP);
  FreeAndNil(TSlCacheCountryBTP);
  FreeAndNil(TSlCacheRecoveryBTP);
  FreeAndNil(TSlCacheCommonBTP);
  FreeAndNil(TSlCacheChoixCodBTP);
  FreeAndNil(TSlCacheJournalBTP);
  FreeAndNil(TSlCacheBankIdBTP);
  FreeAndNil(TSlCacheChangeRateBTP);
  FreeAndNil(TSlCacheFiscalYearBTP);
  FreeAndNil(TSlcacheSocietyParamBTP);
  FreeAndNil(TSlcacheEstablishmentBTP);
  FreeAndNil(TSlcachePaymentModeBTP);
  FreeAndNil(TSlcacheZipCodeBTP);
  FreeAndNil(TSlcacheContactBTP);
  FreeAndNil(TSlCacheThirdY2);
  FreeAndNil(TSlCacheSectionY2);
  FreeAndNil(TSlCacheAcountY2);
  FreeAndNil(TSlCachePaymentY2);
  FreeAndNil(TSlCacheCorrespY2);
  FreeAndNil(TSlCacheCurrencyY2);
  FreeAndNil(TSlCacheCountryY2);
  FreeAndNil(TSlCacheRecoveryY2);
  FreeAndNil(TSlCacheCommonY2);
  FreeAndNil(TSlCacheChoixCodY2);
  FreeAndNil(TSlCacheJournalY2);
  FreeAndNil(TSlCacheBankIdY2);
  FreeAndNil(TSlCacheChangeRateY2);
  FreeAndNil(TSlCacheFiscalYearY2);
  FreeAndNil(TSlcacheSocietyParamY2);
  FreeAndNil(TSlcacheEstablishmentY2);
  FreeAndNil(TSlcachePaymentModeY2);
  FreeAndNil(TSlcacheZipCodeY2);
  FreeAndNil(TSlcacheContactY2);
  FreeAndNil(TSlCacheWSFields);
  FreeAndNil(TSlCacheSendAccParam);
  FreeAndNil(TSlCacheGetY2Data);
  FreeAndNil(TSLCacheUpdateFrequencySetting);
  FreeAndNil(TSLUpdateInsertData);
  FreeAndNil(TSLTRAFileQty);
  AdoQryBTP.Free;
  AdoQryY2.Free;
  if LogValues.LogLevel > 0 then
    CloseFile(LogFile);
end;

procedure TSvcSyncBTPY2Execute.ClearValuesConnection;
begin
  TSlCacheGetY2Data.Clear;
  BTPValues.ConnectionName := '';
  BTPValues.UserAdmin      := '';
  BTPValues.Server         := '';
  BTPValues.DataBase       := '';
  BTPValues.LastSynchro    := '';
  Y2Values.ConnectionName  := '';
  Y2Values.Server          := '';
  Y2Values.DataBase        := '';
end;

procedure TSvcSyncBTPY2Execute.SetLastSyncIniFile;
var
  Section: string;
  SettingFile: TInifile;
begin
  Section    := copy(BTPValues.ConnectionName, 2, length(BTPValues.ConnectionName) - 2);
  SettingFile := TIniFile.Create(SettingsFilePath);
  try
    SettingFile.WriteString(Section, IniFileLastSynchro, DateTimeToStr(Now));
  finally
    SettingFile.Free;
  end;
end;

function TSvcSyncBTPY2Execute.GetSettingsFileName(Extension: string): string;
var
  Path: string;
begin
  if ApplicationName = '' then
    ApplicationName := Application.Name;
  Path   := ExtractFilePath(ApplicationName);
  Result := ExtractFileName(ApplicationName);
  Result := Path + Copy(Result, 1, Pos('.', Result)) + Extension;
end;

procedure TSvcSyncBTPY2Execute.ReadSettings;
var
  SettingFile       : TextFile;
  LineValue         : string;
  CurrentConnection : string;
  InSetting         : boolean;
  InConnection      : boolean;
  InFrequency       : boolean;

  procedure AddConnection;
  begin
    if BTPValues.UserAdmin <> '' then
    begin
      AdoQryBTP.ServerName := Tools.iif(AdoQryBTP.ServerName = '', BTPValues.Server  , AdoQryBTP.ServerName);
      AdoQryBTP.DBName     := Tools.iif(AdoQryBTP.DBName     = '', BTPValues.DataBase, AdoQryBTP.DBName);
      AdoQryY2.ServerName  := Tools.iif(AdoQryY2.ServerName  = '', Y2Values.Server   , AdoQryY2.ServerName);
      AdoQryY2.DBName      := Tools.iif(AdoQryY2.DBName      = '', Y2Values.DataBase , AdoQryY2.DBName);
      TSlConnectionValues.Add(Format('%s=%s;%s;%s;%s;%s;%s', [CurrentConnection, BTPValues.UserAdmin, BTPValues.Server, BTPValues.DataBase, BTPValues.LastSynchro, Y2Values.Server, Y2Values.DataBase]));
    end;
    ClearValuesConnection;
  end;

begin
  ClearValuesConnection;
  SettingsFilePath := GetSettingsFileName('ini');
  AssignFile(SettingFile, SettingsFilePath);
  Reset(SettingFile);
  InSetting := False;
  InConnection := False;
  InFrequency  := False;
  while not Eof(SettingFile) do
  begin
    Readln(SettingFile, LineValue);
    if (Copy(LineValue, 1, 1) = '[') then
    begin
      AddConnection;
      InSetting := (LineValue = '[' + WSCDS_SectionGlobalSettings + ']');
      InConnection := (Copy(LineValue, 1, 11) = '[' + WSCDS_SectionConnection);
      InFrequency := (LineValue = '[' + WSCDS_SectionUpdateFrequency + ']');
      if InConnection then
        CurrentConnection := LineValue;
    end;
    if InSetting then
    begin
      case Tools.CaseFromString(copy(LineValue, 1, Pos('=', LineValue) - 1), ['SecondTimeout'       // Délai de déclenchement
                                                                              , 'LogLevel'          // Niveau du log : 0=Pas de log, 1=Log activé, 2=Log détaillé
                                                                              , 'LogMoMaxSize'      // Taille maximum des logs en Mo: 0=illimitée
                                                                              , 'LogMaxQty'         // Nbre maximum de logs
                                                                              ]) of
        {SecondTimeout} 0: SecondTimeout := StrToInt(GetValueFormSettingLine(LineValue));
        {LogLevel}      1: LogValues.LogLevel := StrToInt(GetValueFormSettingLine(LineValue));
        {LogMoMaxSize}  2: LogValues.LogMoMaxSize := StrToFloat(GetValueFormSettingLine(LineValue));
        {LogMaxQty}     3: LogValues.LogMaxQty := StrToInt(GetValueFormSettingLine(LineValue));
      end;
    end
    else if InConnection then
    begin
      case Tools.CaseFromString(copy(LineValue, 1, Pos('=', LineValue) - 1), [IniFileBTPUser       // BTP : Utilisateur pour la mise à jour/création
                                                                              , IniFileBTPServer   // BTP : Serveur
                                                                              , IniFileBTPDataBase // BTP : Base
                                                                              , IniFileLastSynchro // Date de dernière synchronisation
                                                                              , IniFileY2Server    // Y2 : Serveur
                                                                              , IniFileY2DataBase  // Y2 : Base
                                                                              ]) of
        {BTPUser}        0: BTPValues.UserAdmin   := GetValueFormSettingLine(LineValue);
        {BTPServer}      1: BTPValues.Server      := GetValueFormSettingLine(LineValue);
        {BTPDataBase}    2: BTPValues.DataBase    := GetValueFormSettingLine(LineValue);
        {BTPLastSynchro} 3: BTPValues.LastSynchro := GetValueFormSettingLine(LineValue);
        {Y2Server}       4: Y2Values.Server       := GetValueFormSettingLine(LineValue);
        {Y2DataBase}     5: Y2Values.DataBase     := GetValueFormSettingLine(LineValue);
      end;
    end
    else if InFrequency then
      TSLCacheUpdateFrequencySetting.Add(LineValue);
  end;
  AddConnection;
  CloseFile(SettingFile);
end;

procedure TSvcSyncBTPY2Execute.LogsManagement;
var
  PathLog    : string;
  SizeFile   : Extended;
  SearchFile : TSearchRec;
  MaxSize    : double;
begin
  if LogValues.LogLevel > 0 then
  begin
    PathLog := GetSettingsFileName('log');
    MaxSize := LogValues.LogMoMaxSize;
    { Si dépasse la taille max, supprime puis créé un nouveau }
    if (MaxSize > 0) then
    begin
      if FindFirst(PathLog, faAnyFile, SearchFile) = 0 then
      try
        begin
          SizeFile := Tools.GetFileSize(PathLog, tssMo);
          if SizeFile > MaxSize then
            DeleteFile(PathLog);
        end;
      finally
        FindClose(SearchFile);
      end;
    end;
    AssignFile(LogFile, PathLog);
    if FileExists(PathLog) then
      Append(LogFile)
    else
      Rewrite(LogFile);
  end;
end;

function TSvcSyncBTPY2Execute.GetData: boolean;

  function GetPrefixSECTION(IsBTP: boolean): string;
  begin
    Result := Tools.iif(IsBTP, 'S', 'CSP');
  end;

  function GetTotalQty: integer;
  var
    CptQty: integer;
    Value: string;
  begin
    Result := 0;
    for CptQty := pred(TSlValues.Count) downto 0 do
    begin
      Value := TSlValues[CptQty];
      if Copy(Value, 1, 1) = '#' then
      begin
        Result := StrToInt(Copy(Value, Pos('=', Value) + 1, Length(Value))) + 1;
        Break;
      end;
    end;
  end;

  function GetFromDSType(DSType: T_WSDataService): boolean;
  var
    Index       : Integer;
    Execute     : boolean;
    LastSynchro : TDate;
    Action      : string;
    TableName   : string;
  begin
    Index := TSLCacheUpdateFrequencySetting.IndexOfName(TReadWSDataService.GetTableNameFromDSType(DSType));
    Execute := (Index = -1);
    if not Execute then
    begin
      LastSynchro := Int(StrToDateTime(BTPValues.LastSynchro));
      TableName   := Copy(TSLCacheUpdateFrequencySetting[Index], 1, Pos('=', TSLCacheUpdateFrequencySetting[Index])-1);
      Action := Trim(Copy(TSLCacheUpdateFrequencySetting[Index], Pos('=', TSLCacheUpdateFrequencySetting[Index]) + 1, Length(TSLCacheUpdateFrequencySetting[Index])));
      case Tools.CaseFromString(UpperCase(Action), ['DAILY', 'MONTHLY', 'ANNUAL', 'ONCE', 'EVERYTIME']) of
        {DAILY}     0 : Execute := (LastSynchro < Date);
        {MONTHLY}   1 : Execute := (MonthOf(LastSynchro) < MonthOf(Date));
        {ANNUAL}    2 : Execute := (YearOf(LastSynchro) < YearOf(Date));
        {ONCE}      3 : Execute := (LastSynchro <= 2);
        {EVERYTIME} 4 : Execute := True;                                   
      else
        begin
          WriteLog(ssbylLog, Format('ATTENTION, la fréquence "%s" associée à la table "%s" du fichier de configuration est inconnue. Le traitement n''a pas été effectué pour cette dernière.', [Action, TableName]), 3);
          Execute := False;
        end;
      end;
    end;
    if Execute then
      Result :=GetY2Data(DSType)
     else
      Result := True;
  end;

  function InsertOrUpdateData : Boolean;
  var
    Cpt    : integer;
    UpdQty : integer;
    InsQty : integer;
  begin
    Result := False;
    UpdQty := 0;
    InsQty := 0;
    AdoQryBTP.TSLResult.Clear;
    for Cpt := 0 to pred(TSLUpdateInsertData.Count) do
    begin
      AdoQryBTP.Request := TSLUpdateInsertData[Cpt];
      if Copy(AdoQryBTP.Request, 1, 6) = 'UPDATE' then
        Inc(UpdQty)
      else
        Inc(InsQty);
      AdoQryBTP.InsertUpdate;
      Result := (AdoQryBTP.RecordCount = 1);
      if not Result then
        Break;
    end;
    if Result then
    begin
      WriteLog(ssbylLog, Format('Modification de %s enregistrement(s)', [IntToStr(UpdQty)]), 2);
      WriteLog(ssbylLog, Format('Création de %s enregistrement(s)', [IntToStr(InsQty)]), 2);
    end;
  end;

begin
  WriteLog(ssbylLog, 'Récupération des données créées ou modifiées depuis le ' + BTPValues.LastSynchro, 2);
  Result := True;
  TSlCacheGetY2Data.Clear;
  TSLUpdateInsertData.Clear;
  if Result then Result := GetFromDSType(wsdsCommon);
  if Result then Result := GetFromDSType(wsdsChoixCod);
  if Result then Result := GetFromDSType(wsdsEstablishment);
  if Result then Result := GetFromDSType(wsdsSocietyParameters);
  if Result then Result := GetFromDSType(wsdsPaymentMode);
  if Result then Result := GetFromDSType(wsdsZipCode);
  if Result then Result := GetFromDSType(wsdsRecovery);
  if Result then Result := GetFromDSType(wsdsCountry);
  if Result then Result := GetFromDSType(wsdsCurrency);
  if Result then Result := GetFromDSType(wsdsCorrespondence);
  if Result then Result := GetFromDSType(wsdsFiscalYear);
  if Result then Result := GetFromDSType(wsdsJournal);
  if Result then Result := GetFromDSType(wsdsAccount);
  if Result then Result := GetFromDSType(wsdsPaymenChoice);
  if Result then Result := GetFromDSType(wsdsThird);
  if Result then Result := GetFromDSType(wsdsAnalyticalSection);
  if Result then Result := GetFromDSType(wsdsBankIdentification);
  if Result then Result := GetFromDSType(wsdsChangeRate);
  if Result then Result := GetFromDSType(wsdsContact);
  if Result then Result := InsertOrUpdateData;
  WriteLog(ssbylLog, 'Fin de la récupération des données créées ou modifiées depuis le ' + BTPValues.LastSynchro, 2);
end;

function TSvcSyncBTPY2Execute.SendData: boolean;
var
  Cpt           : integer;
  AdoQryEcr     : AdoQry;
  AdoQryAna     : AdoQry;
  TSlAccEntries : TStringList;
begin
  WriteLog(ssbylLog, 'Envoi des données créées ou modifiées depuis le ' + BTPValues.LastSynchro, 2);
  Result := True;
  { Recherche s'il existe des écritures non envoyées }
  AdoQryBTP.TSLResult.Clear;
  AdoQryBTP.FieldsList := 'BE0_ENTITY,BE0_EXERCICE,BE0_JOURNAL,BE0_NUMEROPIECE';
  AdoQryBTP.Request    := Format('SELECT %s FROM BTPECRITURE WHERE BE0_REFERENCEY2 = 0', [AdoQryBTP.FieldsList]);
  AdoQryBTP.SingleTableSelect;
  if AdoQryBTP.RecordCount > 0 then
  begin
    WriteLog(ssbylLog, Format('%s pièce(s) comptable trouvée(s).', [IntToStr(AdoQryBTP.RecordCount)]), 3);
    AdoQryEcr := AdoQry.Create;
    try
      AdoQryAna := AdoQry.Create;
      try
        TSlAccEntries := TStringList.create;
        try
          AdoQryEcr.ServerName := AdoQryBTP.ServerName;
          AdoQryEcr.DBName     := AdoQryBTP.DBName;
          AdoQryEcr.FieldsList := Tools.GetFieldsListFromTableName('E', AdoQryBTP.ServerName, AdoQryBTP.DBName);
          AdoQryAna.ServerName := AdoQryBTP.ServerName;
          AdoQryAna.DBName     := AdoQryBTP.DBName;
          AdoQryAna.FieldsList := Tools.GetFieldsListFromTableName('Y', AdoQryBTP.ServerName, AdoQryBTP.DBName);
          for Cpt := 0 to pred(AdoQryBTP.TSLResult.Count) do
          begin
            LoadEcrFromBE0(AdoQryEcr, AdoQryAna, AdoQryBTP.TSLResult[Cpt]); // Charge les écritures et analytique
            if AdoQryEcr.RecordCount > 0 then
            begin
              TSlAccEntries.Clear;
              SetSendY2TSl(AdoQryEcr, AdoQryAna, TSlAccEntries);
              SearchRelatedParameters(TSlAccEntries);
            end;
          end;
          SearchOthersParameters;
          SendY2Settings;
        finally
          FreeAndNil(TSlAccEntries);
        end;
      finally
        AdoQryAna.free;
      end;
    finally
      AdoQryEcr.free;
    end;
  end;
  AdoQryBTP.TSLResult.Clear;
  AdoQryBTP.FieldsList := '';
  AdoQryBTP.Request    := '';
  WriteLog(ssbylLog, 'Fin de l''envoi des données créées ou modifiées depuis le ' + BTPValues.LastSynchro, 2);
end;

function TSvcSyncBTPY2Execute.ServiceExecute: Boolean;
var
  Cpt: integer;

  procedure InitConnectionData(Values: string);
  begin
    BTPValues.ConnectionName := Tools.ReadTokenSt_(Values, '=');
    BTPValues.UserAdmin      := Tools.ReadTokenSt_(Values, ';');
    BTPValues.Server         := Tools.ReadTokenSt_(Values, ';');
    BTPValues.DataBase       := Tools.ReadTokenSt_(Values, ';');
    BTPValues.LastSynchro    := Tools.ReadTokenSt_(Values, ';');
    Y2Values.ConnectionName  := BTPValues.ConnectionName;
    Y2Values.Server          := Tools.ReadTokenSt_(Values, ';');
    Y2Values.DataBase        := Tools.ReadTokenSt_(Values, ';');
    AdoQryY2.ServerName      := Y2Values.Server;
    AdoQryY2.DBName          := Y2Values.DataBase;
    AdoQryBTP.ServerName     := BTPValues.Server;
    AdoQryBTP.DBName         := BTPValues.DataBase;
  end;

begin
  Result := True;
  LogsManagement;
  WriteLog(ssbylLog, 'Début d''exécution du service.', 0);
  for Cpt := 0 to pred(TSlConnectionValues.Count) do
  begin
    InitConnectionData(TSlConnectionValues[Cpt]);
    WriteLog(ssbylLog, Format('Traitement de la connexion %s', [BTPValues.ConnectionName]), 1);
    try
      Result := GetData;
      if Result then
        Result := SendData;
    finally
      SetLastSyncIniFile;
      ClearValuesConnection;
    end;
  end;
  WriteLog(ssbylLog, 'Fin d''exécution du service.', 0);
  WriteLog(ssbylLog, DupeString('*', 50), 0);
end;

procedure TSvcSyncBTPY2Execute.InitApplication;
begin
  ReadSettings;
  CreateMemoryCache;
  AddWindowsLog(True);
end;

end.

