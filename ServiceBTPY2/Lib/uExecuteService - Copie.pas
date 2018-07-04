unit uExecuteService;

interface

uses
  Windows
  , Messages
  , SysUtils
  , Classes
  , Graphics
  , Controls
  , SvcMgr
  , Dialogs
  ;

type

  TSvcSyncBTPY2Log = (ssbylNone, ssbylAll, ssbylLog, ssbylDebug);

  TSvcSyncBTPY2Execute = class (TObject)
  private
    procedure WriteInLog(TypeDebug : TSvcSyncBTPY2Log; Text : string);
    function GetValueFormSettingLine(LineValue : string) : string;

  public
    TSLY2Database   : TStringList;
    SecondTimeout   : integer;
    ApplicationName : string;
    BTPUserAdmin    : string;     
    BTPServer       : string;
    BTPDataBase     : string;
    BTPLastSynchro  : string;
    Y2Server        : string;
    Y2DataBase      : string;
    Y2LastSynchro   : string;
    LogDebug        : boolean;
    LogFile         : TextFile;
    DebugLogFile    : TextFile;

    Constructor Create;
    Destructor Destroy; override;

    procedure AssignObjects;
    function GetSettingsFileName(Extension : string) : string;
    procedure ReadSettings;
    procedure ReadDBData;
    function Y2DataRecovery : boolean;

  end;


implementation

uses
  Registry
  , CommonTools
  , UConnectWSConst
  , uWSDataService
  {$IF not defined(APPSRV)}
  , UConnectWSCEGID
  {$IFEND !APPSRV}
  ;

Constructor TSvcSyncBTPY2Execute.Create;
begin

end;

Destructor TSvcSyncBTPY2Execute.Destroy;
begin
  FreeAndNil(TSLY2Database);
  CloseFile(LogFile);
  if LogDebug then
    CloseFile(DebugLogFile);
  inherited;
end;

procedure TSvcSyncBTPY2Execute.WriteInLog(TypeDebug : TSvcSyncBTPY2Log; Text : string);
var
  LogText : string;
begin
  LogText := Format('%s : %s', [DateTimeToStr(Now), Text]);
  case TypeDebug of
    ssbylAll   :  begin
                    Writeln(LogFile, LogText);
                    if LogDebug then
                      Writeln(DebugLogFile, LogText);
                  end;
    ssbylLog   :  Writeln(LogFile, LogText);
    ssbylDebug :  if LogDebug then
                    Writeln(DebugLogFile, LogText);
  end;
end;

function TSvcSyncBTPY2Execute.GetValueFormSettingLine(LineValue : string) : string;
begin
  Result := copy(LineValue, Pos('=', LineValue) + 1, Length(LineValue));
end;

procedure TSvcSyncBTPY2Execute.AssignObjects;
var
  PathSettingFile : string;
begin
  TSLY2Database   := TStringList.Create;
  PathSettingFile := GetSettingsFileName('log');
  AssignFile(LogFile, PathSettingFile);
  Rewrite(LogFile);
end;

function TSvcSyncBTPY2Execute.GetSettingsFileName(Extension : string) : string;
var
  Path : string;
begin
  if ApplicationName = '' then
    ApplicationName := Application.Name;
  Path   := ExtractFilePath(ApplicationName);
  Result := ExtractFileName(ApplicationName);
  Result := Path + Copy(Result, 1, pos('.', Result)) + Extension;
end;
  
procedure TSvcSyncBTPY2Execute.ReadSettings;
var
  SettingFile     : TextFile;
  PathSettingFile : string;
  LineValue       : string;
  InSetting       : Boolean;
begin
  WriteInLog(ssbylLog, 'Lecture des paramètres :');
  TSLY2Database.Clear;
  PathSettingFile := GetSettingsFileName('ini');
  AssignFile(SettingFile, PathSettingFile);
  Reset(SettingFile);
  InSetting := False;
  while not Eof(SettingFile) do
  begin
    Readln(SettingFile, LineValue);
    if (Copy(LineValue, 1, 1) = '[') then
      InSetting := (LineValue = '[SETTINGS]');
    if InSetting then
    begin
      case Tools.CaseFromString(copy(LineValue, 1, Pos('=', LineValue)-1), ['SecondTimeout', 'Server', 'DataBase', 'LastSynchro', 'User', 'LogDebug']) of
        {SecondTimeout} 0 : SecondTimeout  := StrToInt(GetValueFormSettingLine(LineValue));
        {Server}        1 : BTPServer      := GetValueFormSettingLine(LineValue);
        {DataBase}      2 : BTPDataBase    := GetValueFormSettingLine(LineValue);
        {LastSynchro}   3 : BTPLastSynchro := GetValueFormSettingLine(LineValue);
        {User}          4 : BTPUserAdmin   := GetValueFormSettingLine(LineValue);
        {LogDebug}      5 : LogDebug       := (UpperCase(GetValueFormSettingLine(LineValue)) = 'TRUE');
      end;
    end;
  end;
  if LogDebug then
  begin
    PathSettingFile := Copy(PathSettingFile, 1, Pos('.ini', PathSettingFile)) + 'Debug.log';
    AssignFile(DebugLogFile, PathSettingFile);
    Rewrite(DebugLogFile);
    WriteInLog(ssbylDebug, '-> uExecuteService/TSvcSyncBTPY2Execute.ReadSettings');
  end;
  WriteInLog(ssbylDebug, Format(' Ouverture de "%s"', [GetSettingsFileName('ini')]));
  WriteInLog(ssbylAll  , Format('  Intervalle de déclenchement : %s seconde(s)', [IntToStr(SecondTimeout)]));
  WriteInLog(ssbylAll  , Format('  Serveur                     : %s', [BTPServer]));
  WriteInLog(ssbylAll  , Format('  Base de données             : %s', [BTPDataBase]));
  WriteInLog(ssbylAll  , Format('  Dernière synchronisation    : %s', [BTPLastSynchro]));
  WriteInLog(ssbylAll  , Format('  Utilisateur                 : %s', [BTPUserAdmin]));
  WriteInLog(ssbylDebug, Format(' Fermeture de "%s"', [GetSettingsFileName('ini')]));
  CloseFile(SettingFile);
end;

procedure TSvcSyncBTPY2Execute.ReadDBData;
var
  SettingFile     : TextFile;
  PathSettingFile : string;
  LineValue       : string;
  DBNumber        : string;
  Values          : string;
  InDatabase      : boolean;

  procedure AddLine;
  var
    LocValues : string;
  begin
    if (Values <> '') and (TSLY2Database.IndexOf(DBNumber) = -1) then
    begin
      Values := Copy(Values, 2, Length(Values));
      TSLY2Database.Add(DBNumber + '=' + Values);
      LocValues     := Values;
      Y2Server      := Tools.ReadTokenSt(LocValues, ';');
      Y2DataBase    := Tools.ReadTokenSt(LocValues, ';');
      Y2LastSynchro := Tools.ReadTokenSt(LocValues, ';');
      WriteInLog(ssbylLog, ' Base Y2 :');
      WriteInLog(ssbylLog, '  Serveur                  = ' + Y2Server);
      WriteInLog(ssbylLog, '  Nom                      = ' + Y2DataBase);
      WriteInLog(ssbylLog, '  Dernière synchronisation = ' + Y2LastSynchro);
      Values := '';
    end;
  end;

begin
  WriteInLog(ssbylDebug, '-> uExecuteService/TSvcSyncBTPY2Execute.ReadDBData');
  TSLY2Database.Clear;
  PathSettingFile := GetSettingsFileName('ini');
  WriteInLog(ssbylDebug, Format(' Ouverture de "%s"', [PathSettingFile]));
  AssignFile(SettingFile, PathSettingFile);
  Reset(SettingFile);
  InDatabase := False;
  Values     := '';
  while not Eof(SettingFile) do
  begin
    Readln(SettingFile, LineValue);
    if (Copy(LineValue, 1, 1) = '[') then
    begin
      InDatabase := (Copy(LineValue, 1, 5) = '[DBY2');
      if InDatabase then
      begin
        WriteInLog(ssbylDebug, Format('  Lecture "%s"', [LineValue]));
        AddLine;
        DBNumber := LineValue;
      end;
    end else
    if InDatabase then
    begin
      if Pos('=', LineValue) > 0 then
      begin
        WriteInLog(ssbylDebug, Format('  Lecture "%s"', [LineValue]));
        Values := Values + ';' + GetValueFormSettingLine(LineValue);
      end;
    end;
  end;
  AddLine;
  CloseFile(SettingFile);
  WriteInLog(ssbylDebug, Format(' Fermeture de "%s"', [PathSettingFile]));
end;

function TSvcSyncBTPY2Execute.Y2DataRecovery: boolean;
var
  TSlValues       : TStringList;
  TSlFilter       : TStringList;
  TSlIndice       : TStringList;
  TSlCacheThird   : TStringList;
  TSlCacheSection : TStringList;
  TSlCacheAcount  : TStringList;
  {$IF defined(APPSRV)}
  AdoQryY2        : AdoQry;
  AdoQryBTP       : AdoQry;
  {$IFEND !APPSRV}

  {$IF defined(APPSRV)}
  procedure InitObjects;
  begin
    WriteInLog(ssbylDebug, '-> uExecuteService/TSvcSyncBTPY2Execute.Y2DataRecovery.InitObjects');
    TSlValues       := TstringList.Create; WriteInLog(ssbylDebug, ' TSlValues');
    TSlFilter       := TStringList.Create; WriteInLog(ssbylDebug, ' TSlFilter');
    TSlIndice       := TStringList.Create; WriteInLog(ssbylDebug, ' TSlIndice');
    TSlCacheThird   := TStringList.Create; WriteInLog(ssbylDebug, ' TSlCacheThird');
    TSlCacheSection := TStringList.Create; WriteInLog(ssbylDebug, ' TSlCacheSection');
    TSlCacheAcount  := TStringList.Create; WriteInLog(ssbylDebug, ' TSlCacheAcount');
    AdoQryBTP       := AdoQry.Create;      WriteInLog(ssbylDebug, ' AdoQryBTP');
    AdoQryY2        := AdoQry.Create;      WriteInLog(ssbylDebug, ' AdoQryY2');
  end;

  procedure FreeObjects;
  begin
    WriteInLog(ssbylDebug, '-> uExecuteService/TSvcSyncBTPY2Execute.Y2DataRecovery.FreeObjects');
    FreeAndNil(TSlValues);        WriteInLog(ssbylDebug, ' TSlValues');
    FreeAndNil(TSlFilter);        WriteInLog(ssbylDebug, ' TSlFilter');
    FreeAndNil(TSlIndice);        WriteInLog(ssbylDebug, ' TSlIndice');
    FreeAndNil(TSlCacheThird);    WriteInLog(ssbylDebug, ' TSlCacheThird');
    FreeAndNil(TSlCacheSection);  WriteInLog(ssbylDebug, ' TSlCacheSection');
    FreeAndNil(TSlCacheAcount);   WriteInLog(ssbylDebug, ' TSlCacheAcount');
    AdoQryBTP.Free;               WriteInLog(ssbylDebug, ' AdoQryBTP');
    AdoQryY2.Free;                WriteInLog(ssbylDebug, ' AdoQryY2');
  end;

  procedure InitAdoQrys;
  begin
    WriteInLog(ssbylDebug, '-> uExecuteService/TSvcSyncBTPY2Execute.Y2DataRecovery.InitAdoQrys');
    AdoQryY2.ServerName  := Y2Server;    WriteInLog(ssbylDebug, ' AdoQryY2.ServerName');
    AdoQryY2.DBName      := Y2DataBase;  WriteInLog(ssbylDebug, ' AdoQryY2.DBName');
    AdoQryBTP.ServerName := BTPServer;   WriteInLog(ssbylDebug, ' AdoQryBTP.ServerName');
    AdoQryBTP.DBName     := BTPDataBase; WriteInLog(ssbylDebug, ' AdoQryBTP.DBName');
  end;

  procedure CreateMemoryCache;

    procedure SetDefaultValue(Prefix : string; lTsl : TStringList);
    var
      CptCache     : integer;
      Value        : string;
      FieldName    : string;
      DefaultValue : string;
    begin
      AdoQryBTP.TSLResult.Clear;
      AdoQryBTP.FieldsList := 'DH_NOMCHAMP,DH_TYPECHAMP';
      AdoQryBTP.Request    := Format('SELECT %s FROM DECHAMPS WHERE DH_PREFIXE = ''%s'' ORDER BY DH_NUMCHAMP', [AdoQryBTP.FieldsList, Prefix] );
      AdoQryBTP.SingleTableSelect;
      if AdoQryBTP.TSLResult.Count > 0 then
      begin
        for CptCache := 0 to pred(AdoQryBTP.TSLResult.Count) do
        begin
          Value     := AdoQryBTP.TSLResult[CptCache];
          FieldName := Tools.ReadTokenSt(Value, '^');
          case Tools.GetTypeFieldFromStringType(Value) of
            ttfNumeric : DefaultValue := '0';
            ttfInt     : DefaultValue := '0';
            ttfBoolean : DefaultValue := '-';
            ttfDate    : DefaultValue := DateToStr(2);
            ttfMemo    : DefaultValue := '';
            ttfCombo   : DefaultValue := '';
            ttfText    : DefaultValue := '';
          end;
          lTsl.Add(FieldName + '^' + DefaultValue);
        end;
      end;
      AdoQryBTP.TSLResult.Clear;
      AdoQryBTP.RecordCount := 0;
    end;

  begin
    WriteInLog(ssbylDebug, '-> uExecuteService/TSvcSyncBTPY2Execute.Y2DataRecovery.CreateMemoryCache');
    SetDefaultValue('T', TSlCacheThird);   WriteInLog(ssbylDebug, ' TSlCacheThird');
    SetDefaultValue('S', TSlCacheSection); WriteInLog(ssbylDebug, ' TSlCacheSection');
    SetDefaultValue('G', TSlCacheAcount);  WriteInLog(ssbylDebug, ' TSlCacheAcount');  
  end;                           

  function GetInfoFromDSType(InfoType : T_WSInfoFromDSType; DSType : T_WSDataService; FieldName : string='') : string;
  begin
    case DSType of
      wsdsCustomer :          begin
                                case InfoType of
                                  wsidTableName     : Result := 'TIERS';
                                  wsidFieldsKey     : Result := 'T_AUXILIAIRE';
                                  wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';T_NATUREAUXI;T_AUXILIAIRE;T_TIERS;') > 0));
                                else
                                  Result := '';
                                end;
                              end;
      wsdsAnalyticalSection : begin
                                case InfoType of
                                  wsidTableName     : Result := 'SECTION';
                                  wsidFieldsKey     : Result := 'S_AXE;S_SECTION';
                                  wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';S_AXE;S_SECTION;') > 0));
                                else
                                  Result := '';
                                end;
                              end;
      wsdsAccount :           begin
                                case InfoType of
                                  wsidTableName     : Result := 'GENERAUX';
                                  wsidFieldsKey     : Result := 'G_GENERAL';
                                  wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';G_GENERAL;') > 0));
                                else
                                  Result := '';
                                end;
                              end;
      wsdsJournal :           begin
                                case InfoType of
                                  wsidTableName     : Result := 'JOURNAL';
                                  wsidFieldsKey     : Result := 'J_JOURNAL';
                                  wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';J_JOURNAL;') > 0));
                                else
                                  Result := '';
                                end;
                              end;
    else
      Result := '';
    end;
  end;




  function GetPrefixSECTION(IsBTP : boolean) : string;
  begin
    Result := Tools.iif(IsBTP, 'S', 'CSP');
  end;

  procedure ExtractIndice(Cpt : integer; TSlOrig, TSlResult : TStringList);
  var
    sIndex : string;
    Value  : string;
    Pos    : integer;
    iCpt   : integer;
    IndexChange : Boolean;
  begin
    sIndex := '#INDICE' + IntToStr(Cpt) + '#';
    Pos    := TSlOrig.IndexOfName(sIndex);
    if Pos > -1 then
    begin
      sIndex := TSlOrig[Pos];
      for iCpt := Pos to Pred(TSlOrig.Count) do
      begin
        Value := TSlOrig[iCpt];
        IndexChange := ((Copy(Value, 1, 7) = '#INDICE') and (Value <> sIndex));
        if not IndexChange then
        begin
          { Spécif pour SECTION, préfixe Y2=CSP_, préfixe BTP=S_ }
          if Copy(Value, 1, 4) = 'CSP_' then
            Value := StringReplace(Value, 'CSP_', 'S_', [rfReplaceAll]);
          TSlResult.Add(Value);
        end else
          Break;
      end;
    end;
  end;

  function GetValueFrom(lTSlIndice : TStringList; FieldName : string) : string;
  var
    iCpt   : integer;
    lValues : string;
  begin
    Result := '';
    for iCpt := 0 to Pred(lTSlIndice.Count) do
    begin
      lValues := lTSlIndice[iCpt];
      if Copy(lValues, 1, Pos('=', lValues) -1) = FieldName then
      begin
        Result := Copy(lValues, Pos('=', lValues) + 1, Length(lValues));
        Break;
      end;
    end;
  end;
  {$IFEND !APPSRV}

  function AddQuotes(FieldName, FieldValue : string) : string;
  begin
    case Tools.GetFieldType(FieldName, BTPServer, BTPDataBase) of
      ttfBoolean
      , ttfCombo
      , ttfMemo
      , ttfText
      , ttfDate  : Result := Format('''%s''', [FieldValue]);
    else
      Result := FieldValue;
    end;
  end;

  function GetKeyValues(DSType : T_WSDataService; lTSlValues : TStringList) : string;
  var
    FieldsKey   : string;
    Value      : string;
    FieldName  : string;
    FieldValue : string;
    Posit      : integer;
  begin
    Result := '';
    FieldsKey := GetInfoFromDSType(wsidFieldsKey, DSType);
    while FieldsKey <> '' do
    begin
      Posit := lTSlValues.IndexOfName(Tools.ReadTokenSt(FieldsKey, ';'));
      if Posit > -1 then
      begin
        FieldName  := copy(lTSlValues[Posit], 1, pos('=', lTSlValues[Posit])-1);
        FieldValue := copy(lTSlValues[Posit], pos('=', lTSlValues[Posit])+1, length(lTSlValues[Posit]));
        if Tools.GetFieldType(FieldName, BTPServer, BTPDataBase) = ttfCombo then
          FieldValue := Trim(FieldValue);
        FieldValue := AddQuotes(FieldName, FieldValue);
        Value      := Format('-%s=%s', [FieldName, FieldValue]);
        Result     := Result + Value;
      end;
    end;
    Result := Copy(Result, 2, Length(Result));
  end;

  function AddData(wsAction : T_WSAction; DSType : T_WSDataService; lTSlValues : TStringList) : boolean;
  var
    CptUpd         : integer;
    KeyValue       : string;
    KeyValuel      : string;
    FieldName      : string;
    FieldValue     : string;
    Sql            : string;
    Where          : string;
    InsertedFields : string;
    InsertedValues : string;
  begin
    KeyValue       := GetKeyValues(DSType, lTSlValues);
    InsertedFields := '';
    InsertedValues := '';
    WriteInLog(ssbylLog  , Format('%s de "%s"', [Tools.iif(wsAction = wsacUpdate, ' Modification', ' Création'), KeyValue]));
    WriteInLog(ssbylDebug, Format(' -> uExecuteService/TSvcSyncBTPY2Execute.Y2DataRecovery.AddData (%s) - "%s"', [Tools.iif(wsAction = wsacUpdate, 'Update', 'Insert'), KeyValue]));
    for CptUpd := 0 to pred(lTSlValues.Count) do
    begin
      if copy(lTSlValues[CptUpd], 1, 7) <> '#INDICE' then
      begin
        FieldName  := copy(lTSlValues[CptUpd], 1, pos('=', lTSlValues[CptUpd]) - 1);
        if (wsAction = wsacInsert) or ((wsAction = wsacUpdate) and (GetInfoFromDSType(wsidExcludeFields, DSType, FieldName) = '0')) then
        begin
          if Pos('_DATEMODIF', FieldName) > 0 then
            FieldValue := DateTimeToStr(Now)
          else
          begin
            FieldValue := copy(lTSlValues[CptUpd], pos('=', lTSlValues[CptUpd]) + 1, length(lTSlValues[CptUpd]));
            case Tools.GetFieldType(FieldName, BTPServer, BTPDataBase) of
              ttfDate  : FieldValue := Tools.SetStrDateTimeFromStrUTCDateTime(FieldValue);
              ttfCombo : FieldValue := Trim(FieldValue);                                   
            end;
          end;
          if Tools.GetFieldType(FieldName, BTPServer, BTPDataBase) = ttfDate then
            FieldValue := FormatDateTime('yyyymmdd hh:nn:ss', StrToDateTime(FieldValue));
          FieldValue := StringReplace(FieldValue, '''', '''''', [rfReplaceAll]);
          FieldValue := AddQuotes(FieldName, FieldValue);
          Sql := Sql + Format(', %s=%s', [FieldName, FieldValue]);
          InsertedFields := InsertedFields + ', ' + FieldName;
          InsertedValues := InsertedValues + ', ' + FieldValue;
        end;
      end;
    end;
    Sql            := Copy(Sql, 2, Length(Sql));
    InsertedFields := Copy(InsertedFields, 2, Length(InsertedFields));
    InsertedValues := Copy(InsertedValues, 2, Length(InsertedValues));
    if wsAction = wsacUpdate then
    begin
      KeyValuel := KeyValue;
      while KeyValuel <> '' do
        Where := Where + ' AND ' + Tools.ReadTokenSt(KeyValuel, '-');
      Where := Copy(Where, 5, length(Where));
      Sql := Format('UPDATE %s SET %s WHERE %s', [GetInfoFromDSType(wsidTableName, DSType), Sql, Where]);
    end else
      Sql := Format('INSERT INTO %s (%s) VALUES(%s)', [GetInfoFromDSType(wsidTableName, DSType), InsertedFields, InsertedValues]);
    AdoQryBTP.TSLResult.Clear;
    AdoQryBTP.Request := Sql;
    AdoQryBTP.InsertUpdate;
    AdoQryBTP.TSLResult.Clear;
    Result := (AdoQryBTP.RecordCount = 1);
  end;

  function InsertData(DSType : T_WSDataService; lTSlValues : TStringList) : boolean;
  var
    CptInsert      : integer;
    KeyValue       : string;
    FieldName      : string;
    FieldValue     : string;
    Sql            : string;
    InsertedFields : string;
    InsertedValues : string;
  begin
    KeyValue       := GetKeyValues(DSType, lTSlValues);
    InsertedFields := '';
    InsertedValues := '';
    WriteInLog(ssbylLog  , Format(' Création de "%s"', [KeyValue]));
    WriteInLog(ssbylDebug, Format(' -> uExecuteService/TSvcSyncBTPY2Execute.Y2DataRecovery.InsertData - "%s"', [KeyValue]));
    for CptInsert := 0 to pred(lTSlValues.Count) do
    begin
      if copy(lTSlValues[CptInsert], 1, 7) <> '#INDICE' then
      begin
        FieldName  := copy(lTSlValues[CptInsert], 1, pos('=', lTSlValues[CptInsert]) - 1);
        if Pos('_DATEMODIF', FieldName) > 0 then
          FieldValue := DateTimeToStr(Now)
        else
        begin
          FieldValue := copy(lTSlValues[CptInsert], pos('=', lTSlValues[CptInsert]) + 1, length(lTSlValues[CptInsert]));
          case Tools.GetFieldType(FieldName, BTPServer, BTPDataBase) of
            ttfDate  : FieldValue := Tools.SetStrDateTimeFromStrUTCDateTime(FieldValue); // Conversion de la date
            ttfCombo : FieldValue := Trim(FieldValue);                                   // Suppression des espaces dans les combos
          end;
        end;
        if Tools.GetFieldType(FieldName, BTPServer, BTPDataBase) = ttfDate then
          FieldValue := FormatDateTime('yyyymmdd hh:nn:ss', StrToDateTime(FieldValue));
        FieldValue := StringReplace(FieldValue, '''', '''''', [rfReplaceAll]);
        FieldValue := AddQuotes(FieldName, FieldValue);
        InsertedFields := InsertedFields + ', ' + FieldName;
        InsertedValues := InsertedValues + ', ' + FieldValue;
      end;
    end;
    InsertedFields := Copy(InsertedFields, 2, Length(InsertedFields));
    InsertedValues := Copy(InsertedValues, 2, Length(InsertedValues));
    Sql := Format('INSERT INTO %s (%s) VALUES(%s)', [GetInfoFromDSType(wsidTableName, DSType), InsertedFields, InsertedValues]);
    AdoQryBTP.TSLResult.Clear;
    AdoQryBTP.Request := Sql;
    AdoQryBTP.InsertUpdate;
    AdoQryBTP.TSLResult.Clear;
    Result := (AdoQryBTP.RecordCount = 1);
  end;

  function GetTotalQty : integer;
  var
    CptQty : integer;
    Value  : string;
  begin
    Result := 0;
    for CptQty := pred(TSlValues.Count) Downto 0 do
    begin
      Value := TSlValues[CptQty];
      if Copy(Value, 1, 1) = '#' then
      begin
        Result := StrToInt(Copy(Value, pos('=', Value) +1, Length(Value))) + 1;
        Break;
      end;
    end;
  end;

//**************************
  function GetY2Data(InfoType : T_WSDataService) : boolean;
  var
    Cpt       : integer;
    Index     : Integer;
    TableName : string;
    Value     : string;
    Value1    : string;
  begin
    Result    := True;
    TableName := GetInfoFromDSType(wsidTableName, InfoType);
    WriteInLog(ssbylDebug, Format('-> uExecuteService/TSvcSyncBTPY2Execute.Y2DataRecovery.GetY2Data (%s)', [TableName]));
    WriteInLog(ssbylLog, ' ' + TableName);
    case InfoType of
      wsdsCustomer           : TSlFilter.Add(';T_DATEMODIF;>=;'   + Y2LastSynchro);
      wsdsAnalyticalSection  : TSlFilter.Add(';CSP_DATEMODIF;>=;' + Y2LastSynchro);
      wsdsAccount            : TSlFilter.Add(';J_DATEMODIF;>=;'   + Y2LastSynchro);
      wsdsJournal            : TSlFilter.Add(';G_DATEMODIF;>=;'   + Y2LastSynchro);
      wsdsBankIdentification : TSlFilter.Add(';R_DATEMODIF;>=;'   + Y2LastSynchro);
    end;
    TReadWSDataService.GetData(InfoType, Y2Server, Y2DataBase, TSlValues, TSlFilter);
    if TSlValues.Count > 0 then
    begin
      {$IF defined(APPSRV)}
      for Cpt := 0 to Pred(TSlValues.Count) do
      begin
        if Copy(TSlValues[Cpt], 1, 7) = '#INDICE' then
        begin
          Index := StrToInt(Copy(TSlValues[Cpt], Pos('=', TSlValues[Cpt]) + 1, Length(TSlValues[Cpt])));
          TSlIndice.Clear;
          ExtractIndice(Index, TSlValues, TSlIndice);
          { Test si la section existe déjà }
          AdoQryBTP.TSLResult.Clear;
          case InfoType of
            wsdsCustomer           :  begin
                                        AdoQryBTP.FieldsList := 'T_AUXILIAIRE';
                                        Value := GetValueFrom(TSlIndice, AdoQryBTP.FieldsList);
                                        AdoQryBTP.Request := Format('SELECT %s FROM TIERS WHERE T_AUXILIAIRE = ''%s''', [AdoQryBTP.FieldsList, Value]);
                                      end;
            wsdsAnalyticalSection  :  begin
                                        AdoQryBTP.FieldsList := 'S_SECTION';
                                        Value  := GetValueFrom(TSlIndice, 'S_AXE');
                                        Value1 := GetValueFrom(TSlIndice, 'S_SECTION');
                                        AdoQryBTP.Request := Format('SELECT S_SECTION FROM SECTION WHERE S_AXE = ''%s'' AND S_SECTION = ''%s''', [Value, Value1]);
                                      end;
            wsdsAccount            :  begin
                                        AdoQryBTP.FieldsList := 'G_GENERAL';
                                        Value := GetValueFrom(TSlIndice, AdoQryBTP.FieldsList);
                                        AdoQryBTP.Request := Format('SELECT %s FROM GENERAUX WHERE G_GENERAL = ''%s''', [AdoQryBTP.FieldsList, Value]);
                                      end;
            wsdsJournal            :  begin
                                        AdoQryBTP.FieldsList := 'J_JOURNAL';
                                        Value                := GetValueFrom(TSlIndice, 'J_JOURNAL');
                                        AdoQryBTP.Request    := Format('SELECT J_JOURNAL FROM JOURNAL WHERE J_JOURNAL = ''%s''', [Value]);
                                      end;
            wsdsBankIdentification :  begin
                                        AdoQryBTP.FieldsList := 'R_AUXILIAIRE';
                                        Value  := GetValueFrom(TSlIndice, 'R_AUXILIAIRE');
                                        Value1 := GetValueFrom(TSlIndice, 'R_NUMERORIB');
                                        AdoQryBTP.Request := Format('SELECT R_AUXILIAIRE FROM RIB WHERE R_AUXILIAIRE = ''%s'' AND R_NUMERORIB = ''%s''', [Value, Value1]);
                                      end;
          end;
          AdoQryBTP.SingleTableSelect;
          if AdoQryBTP.RecordCount > 0 then
            Result := AddData(wsacUpdate, InfoType, TSlIndice)
          else
            Result := AddData(wsacInsert, InfoType, TSlIndice);
          AdoQryBTP.TSLResult.Clear;
          AdoQryBTP.RecordCount := 0;
        end;
      end;
      {$ELSE  !APPSRV}
      {$IFEND !APPSRV}
    end;
    TSlFilter.Clear;
    TSlValues.Clear;
  end;
//**************************

  function GetJournal : boolean;
  var
    Cpt       : integer;
    QtyUpdate : integer;
    QtyInsert : integer;
    Index     : Integer;
    Value     : string;
  begin
    Result := True;
    WriteInLog(ssbylDebug, '-> uExecuteService/TSvcSyncBTPY2Execute.Y2DataRecovery.GetJournal');
    WriteInLog(ssbylLog, GetInfoFromDSType(wsidTableName, wsdsJournal));
    QtyUpdate := 0;
    QtyInsert := 0;
    TSlFilter.Add(';J_DATEMODIF;>=;' + Y2LastSynchro);
    TReadWSDataService.GetData(wsdsJournal, Y2Server, Y2DataBase, TSlValues, TSlFilter);
    if TSlValues.Count > 0 then
    begin
      {$IF defined(APPSRV)}
      for Cpt := 0 to Pred(TSlValues.Count) do
      begin
        if Copy(TSlValues[Cpt], 1, 7) = '#INDICE' then
        begin
          Index := StrToInt(Copy(TSlValues[Cpt], Pos('=', TSlValues[Cpt]) + 1, Length(TSlValues[Cpt])));
          TSlIndice.Clear;
          ExtractIndice(Index, TSlValues, TSlIndice);
          { Test si la section existe déjà }
          AdoQryBTP.TSLResult.Clear;
          AdoQryBTP.FieldsList := 'J_JOURNAL';
          Value := GetValueFrom(TSlIndice, 'J_JOURNAL');
          AdoQryBTP.Request := Format('SELECT J_JOURNAL FROM JOURNAL WHERE J_JOURNAL = ''%s''', [Value]);
          AdoQryBTP.SingleTableSelect;
          if AdoQryBTP.RecordCount > 0 then
            Result := AddData(wsacUpdate, wsdsJournal, TSlIndice)
          else
            AddData(wsacInsert, wsdsJournal, TSlIndice);
          AdoQryBTP.TSLResult.Clear;
          AdoQryBTP.RecordCount := 0;
        end;
      end;
      {$ELSE  !APPSRV}
      {$IFEND !APPSRV}
    end;
    TSlFilter.Clear;
    TSlValues.Clear;
  end;

  function GetThirds : Boolean;
  var
    Cpt       : integer;
    QtyUpdate : integer;
    QtyInsert : integer;
    Index     : Integer;
    Value     : string;
  begin
    Result := True;
    WriteInLog(ssbylDebug, '-> uExecuteService/TSvcSyncBTPY2Execute.Y2DataRecovery.GetThirds');
    WriteInLog(ssbylLog, GetInfoFromDSType(wsidTableName, wsdsCustomer));
    QtyUpdate := 0;
    QtyInsert := 0;
    TSlFilter.Add(';T_DATEMODIF;>=;' + Y2LastSynchro);
    TReadWSDataService.GetData(wsdsCustomer, Y2Server, Y2DataBase, TSlValues, TSlFilter);
    if TSlValues.Count > 0 then
    begin
      {$IF defined(APPSRV)}
      for Cpt := 0 to Pred(TSlValues.Count) do
      begin
        if Copy(TSlValues[Cpt], 1, 7) = '#INDICE' then
        begin
          Index := StrToInt(Copy(TSlValues[Cpt], Pos('=', TSlValues[Cpt]) + 1, Length(TSlValues[Cpt])));
          TSlIndice.Clear;
          ExtractIndice(Index, TSlValues, TSlIndice);
          { Test si le tiers existe déjà }
          AdoQryBTP.TSLResult.Clear;
          AdoQryBTP.FieldsList := 'T_AUXILIAIRE';
          Value := GetValueFrom(TSlIndice, AdoQryBTP.FieldsList);
          AdoQryBTP.Request := Format('SELECT %s FROM TIERS WHERE T_AUXILIAIRE = ''%s''', [AdoQryBTP.FieldsList, Value]);
          AdoQryBTP.SingleTableSelect;
          if AdoQryBTP.RecordCount > 0 then
            Result := AddData(wsacUpdate, wsdsCustomer, TSlIndice)
          else
            AddData(wsacInsert, wsdsCustomer, TSlIndice);
          AdoQryBTP.TSLResult.Clear;
          AdoQryBTP.RecordCount := 0;
        end;
      end;
      {$ELSE  !APPSRV}
      {$IFEND !APPSRV}
    end;
    TSlFilter.Clear;
    TSlValues.Clear;
    WriteInLog(ssbylAll, Format('  %s tiers modifié(s)', [IntToStr(QtyUpdate)]));
    WriteInLog(ssbylAll, Format('  %s tiers créé(s)', [IntToStr(QtyInsert)]));
  end;

  function GetAnalyticalSections : Boolean;
  var
    Cpt       : integer;
    QtyUpdate : integer;
    QtyInsert : integer;
    Index     : Integer;
    Value     : string;
    Value1    : string;
  begin
    Result := True;
    WriteInLog(ssbylDebug, '-> uExecuteService/TSvcSyncBTPY2Execute.Y2DataRecovery.GetAnalyticalSections');
    WriteInLog(ssbylLog, GetInfoFromDSType(wsidTableName, wsdsAnalyticalSection));
    QtyUpdate := 0;
    QtyInsert := 0;
    TSlFilter.Add(';CSP_DATEMODIF;>=;' + Y2LastSynchro);
    TReadWSDataService.GetData(wsdsAnalyticalSection, Y2Server, Y2DataBase, TSlValues, TSlFilter);
    if TSlValues.Count > 0 then
    begin
      {$IF defined(APPSRV)}
      for Cpt := 0 to Pred(TSlValues.Count) do
      begin
        if Copy(TSlValues[Cpt], 1, 7) = '#INDICE' then
        begin
          Index := StrToInt(Copy(TSlValues[Cpt], Pos('=', TSlValues[Cpt]) + 1, Length(TSlValues[Cpt])));
          TSlIndice.Clear;
          ExtractIndice(Index, TSlValues, TSlIndice);
          { Test si la section existe déjà }
          AdoQryBTP.TSLResult.Clear;
          AdoQryBTP.FieldsList := 'S_SECTION';
          Value  := GetValueFrom(TSlIndice, 'S_AXE');
          Value1 := GetValueFrom(TSlIndice, 'S_SECTION');
          AdoQryBTP.Request := Format('SELECT S_SECTION FROM SECTION WHERE S_AXE = ''%s'' AND S_SECTION = ''%s''', [Value, Value1]);
          AdoQryBTP.SingleTableSelect;
          if AdoQryBTP.RecordCount > 0 then
            Result := AddData(wsacUpdate, wsdsAnalyticalSection, TSlIndice)
          else
            AddData(wsacInsert, wsdsAnalyticalSection, TSlIndice);
          AdoQryBTP.TSLResult.Clear;
          AdoQryBTP.RecordCount := 0;
        end;
      end;
      {$ELSE  !APPSRV}
      {$IFEND !APPSRV}
    end;
    TSlFilter.Clear;
    TSlValues.Clear;
    WriteInLog(ssbylAll, Format('  %s Sections analytique modifiée(s)', [IntToStr(QtyUpdate)]));
    WriteInLog(ssbylAll, Format('  %s Sections analytique créée(s)', [IntToStr(QtyInsert)]));
  end;

  function GetAccount : boolean;
  begin
    WriteInLog(ssbylDebug, '-> uExecuteService/TSvcSyncBTPY2Execute.Y2DataRecovery.GetAccount');
    Result := True;

  end;

  function GetBankIdentification : boolean;
  begin
    WriteInLog(ssbylDebug, '-> uExecuteService/TSvcSyncBTPY2Execute.Y2DataRecovery.GetBankIdentification');
    Result := True;

  end;
  
begin
  WriteInLog(ssbylDebug, '-> uExecuteService/TSvcSyncBTPY2Execute.Y2DataRecovery');
  WriteInLog(ssbylAll, 'Récupération des données créées ou modifiées depuis le ' + Y2LastSynchro);
  InitObjects;
  try
    InitAdoQrys;
    CreateMemoryCache;
    Result := GetY2Data(wsdsAccount);
(*
    Result := GetY2Data(wsdsJournal);
    if Result then
      Result := GetY2Data(wsdsCustomer);
    if Result then
      Result := GetY2Data(wsdsAnalyticalSection);
    if Result then
      Result := GetY2Data(wsdsAccount);
*)
(*
    if Result then
      Result := GetBankIdentification;
*)
  finally
    FreeObjects;
  end;
  WriteInLog(ssbylDebug, '<- uExecuteService/TSvcSyncBTPY2Execute.Y2DataRecovery');
end;

end.
