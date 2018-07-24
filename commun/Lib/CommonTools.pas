unit CommonTools;

interface

uses
  Classes
  {$IF not defined(APPSRV)}
  , UTob
  , HEnt1
  {$IFEND !APPSRV}
  ;
const
  ToolsTobToTsl_LevelName = '^LEVEL';
  ToolsTobToTsl_Separator = '^';

type
  tTypeField = (ttfNone, ttfNumeric, ttfInt, ttfMemo, ttfBoolean, ttfDate, ttfCombo, ttfText);
  tScaleSize = (tssNone, tssKo, TssMo);
  tTableName = (  ttnNone
                , ttnChoixCod    // CHOIXCOD
                , ttnCommun      // COMMUN
                , ttnChoixExt    // CHOIXEXT
                , ttnDevise      // DEVISE
                , ttnModeRegl    // MODEREGL
                , ttnPays        // PAYS
                , ttnRib         // RIB
                , ttnSection     // SECTION
                , ttnTiers       // TIERS
                , ttnCodePostaux // CODEPOST
                , ttnContact     // CONTACT
                , ttnEtabliss    // ETABLISS
                , ttnModePaie    // MODEPAIE
                , ttnGeneraux    // GENERAUX
                , ttnJournal     // JOURNAL
                , ttnRelance     // RELANCE
                , ttnCorresp     // CORRESP
                , ttnChancell    // CHANCELL
                , ttnExercice    // EXERCICE
                , ttnParamSoc    // PARAMSOC
                , ttnEcriture    // ECRITURE
                , ttnAcomptes    // ACOMPTES
               );
  tTypeAlign = (traaNone, traaLeft, traaRigth);
  tFormatValueTypeDate = (tvtNone, tvtDate, tvtDateTime);

  AdoQry = class
  private
    function GetConnectionString : string;
  public
    ServerName  : string;
    DBName      : string;
    Request     : string;
    FieldsList  : string;                                           
    TSLResult   : TStringList;
    RecordCount : integer;

    Constructor Create;
    Destructor Destroy; override;
    procedure SingleTableSelect;
    procedure InsertUpdate;
  end;

  Tools = class
    class function CaseFromString(Value: string; Values: array of string): integer;
    class function GetTypeFieldFromStringType(TypeString : string) : tTypeField;
    class function GetFieldType(FieldName: string{$IF defined(APPSRV)}; ServerName, DBName : string{$IFEND !APPSRV}): tTypeField;
    class function GetDefaultValueFromtTypeField(FieldType : tTypeField) : string;
    class function iif(Const Expression, TruePart, FalsePart: Boolean): Boolean; overload;
    class function iif(Const Expression: Boolean; Const TruePart, FalsePart: Integer): Integer; overload;
    class function iif(Const Expression: Boolean; Const TruePart, FalsePart: Double): Double; overload;
    class function iif(Const Expression: Boolean; Const TruePart, FalsePart: String): String; overload;
    class function iif(Const Expression: Boolean; Const TruePart, FalsePart: char): char; overload;
    class function iif(Const Expression: Boolean; Const TruePart, FalsePart: TStringList): TStringList; overload;
    {$IF not defined(APPSRV)}
    class function iif(Const Expression: Boolean; Const TruePart, FalsePart: TActionFiche): TActionFiche; overload;
    {$IFEND !APPSRV}
    class function ReadTokenSt_(var S : string; Separator : string) : string;
    class function CountOccurenceString(const S : string; ToCount : string) : integer;
    class function GetArgumentValue(Argument: string; Const MyArg : String; Const WithUpperCase: Boolean = True; const Separator: String = ';'): String;
    class function GetArgumentString(Argument: string; Const MyArg : String; WithUpperCase: Boolean = True; const Separator: String = ';') : string;
    class function SetStrDateTimeFromStrUTCDateTime(UTCDateTime : string) : string;
    class function SetStrDateTimeToUTCDateTime(stDateTime : string) : string;
    class function SetStrUTCDateTimeToDateTime(stUTCDateTime : string) : string;
    class function GetFileSize(FilePath : string; Size : tScaleSize) : Extended;
    class function StrFPoint_(Value : Extended) : string;
    class function IsNumeric_(stValue : string) : boolean;
    class function GetFieldsListFromPrefix(TablePrefix : string{$IF defined(APPSRV)}; ServerName, DBName : string{$IFEND !APPSRV}; Separator : string=',') : string;
    class function GetStValueFromTSl(TSlLine, FieldName : string) : string; overload;
    class function GetStValueFromTSl(TSlLine : string; Index : integer; Separator : string=',') : string; overload;
    class function GetTSlIndexFromFieldName(TslLine, FieldName : string; Separator : string=',') : integer;
    class function GetTableNameFromTtn(Ttn : tTableName) : string;
    class function GetTtnFromTableName(TableName : string) : tTableName;
    class function CanInsertedInTable(TableName: string{$IF defined(APPSRV)}; ServerName, DBName : string{$IFEND APPSRV}): Boolean;
    class function GetPSocTreeToExport(OnlyAccounting : Boolean=False) : string;
    class function SetTRAFileFromTSl(lTSL : TStringList) : Boolean;
    class function FormatValue(Value : string; Align : tTypeAlign; iLength : integer; NbDec : Integer=0; TypeDate : tFormatValueTypeDate=tvtNone) : string;
    class function CompressFile(FullPath : string) : string;
    class function UnCompressFile(ZipPath, ZipFileName : string) : integer;
    class procedure FileCut(FullPath : string; MaxSizeBytes : integer; TSLResult : TStringList; KeepOriginFile : boolean=True);
    class function GetKoFromMo(Mo : integer) : integer;
    class function DeleteDirectroy(Path : string) : boolean;
    class function IsRecordableDocument(DocType, Establishment : string{$IF defined(APPSRV)}; ServerName, DBName : string{$IFEND !APPSRV}) : boolean;
    class function EvalueDateJJMMYYY(sDate : string) : TDateTime;
    class procedure DecodeAccDocReferency(DocReferency : string; var DocType : string; var Stump : string; var DocDate : TDateTime; var DocNumber : integer; var Index : integer);
    class function GetParamSocSecur_(PSocName : string; DefaultValue : string{$IFDEF APPSRV}; ServerName, FolderName : string{$ENDIF APPSRV}) : string;
    class function CastDateTimeForQry(lDate : TDateTime) : string;
    {$IF not defined(APPSRV)}
    class procedure TobToTStringList(TobOrig : TOB; TSlResult : TStringList; Level : Integer=1);
    {$IFEND !APPSRV}
  end;

implementation

uses                                                                  
  ADODB
  , Forms
  , SysUtils
  , DB
  , StrUtils
  , Variants
  , DateUtils
  , UConnectWSConst
  , Zip
  , UConnectWSCEGID
  {$IF not defined(APPSRV)}
  , hCtrls
  , EntGC
  , ParamSoc
  {$IFEND !APPSRV}
  ;

{ AdoQry }
function AdoQry.GetConnectionString : string;
begin
  Result := 'Provider=SQLOLEDB.1'
          + ';Password=ADMIN'
          + ';Persist Security Info=True'
          + ';User ID=ADMIN'
          + ';Initial Catalog=' + DBName
          + ';Data Source=' + ServerName
          + ';Use Procedure for Prepare=1'
          + ';Auto Translate=True'
          + ';Packet Size=4096'
          + ';Workstation ID=LOCALST'
          + ';Use Encryption for Data=False'
          + ';Tag with column collation when possible=False';
end;

constructor AdoQry.Create;
begin
  TSLResult           := TStringList.Create;
  TSLResult.Delimiter := ToolsTobToTsl_Separator;
end;

destructor AdoQry.Destroy;
begin
  FreeAndNil(TSLResult);
  inherited;
end;

{ Renvoie dans TSLResult le résultat du SELECT dont les valeurs sont séparées par des ^.
  Exemple de code pour appeler cette méthode :

    lAdoQry := AdoQry.Create;
    try
      lAdoQry.ServerName     := 'SRV-BDD-CLI';
      lAdoQry.DBName         := 'DEMOBTPV10_JT';
      lAdoQry.FieldsList     := 'SOC_DATA,SOC_TREE';
      lAdoQry.Request        := 'SELECT ' + lAdoQry.FieldsList + ' FROM PARAMSOC WHERE SOC_NOM IN (''SO_BTWSSERVEUR'', ''SO_BTWSCEGIDDOS'')';
      lAdoQry.SingleTableSelect;
      ServerName := lAdoQry.TSLResult[0]; // Exemple : MODELE_CEGID_JTR^001;016;022;003;
      FolderName := lAdoQry.TSLResult[1]; // Exemple : SRV-Y2-PHASE2^001;016;022;005;
    finally
      lAdoQry.Free;
    end;
}
procedure AdoQry.SingleTableSelect;
var
  Connect     : TADOConnection;
  Qry         : TADOQuery;
  Cpt         : integer;
  Sql         : string;
  Select      : string;
  ResultValue : string;
  lFieldsList : string;
  Start       : integer;
  FieldsArray : Array of string;
begin
  if     (ServerName <> '') // Nom du serveur
     and (DBName <> '')     // Nom de la BDD
     and (Request <> '')    // Requête
     and (FieldsList <> '') // Liste des champs
  then
  begin
    lFieldsList := FieldsList;
    SetLength(FieldsArray, Tools.CountOccurenceString(lFieldsList, ',') + 1);
    Cpt := 0;
    while lFieldsList <> '' do
    begin
      FieldsArray[Cpt] := Tools.ReadTokenSt_(lFieldsList, ',');
      Inc(Cpt);
    end;
    Sql := Request;
    { Si SELECT *, remplace par un SELECT des champs du tableau }
    if Pos('SELECT *', Sql) > 0 then
    begin
      Select := '';
      for Cpt := 0 to pred(Length(FieldsArray)) do
      begin
        if FieldsArray[Cpt] <> '' then
          Select := Select + ',' + FieldsArray[Cpt];
      end;
      Select := Copy(Select, 2, Length(Select));
      Start  := Pos('SELECT', Request);
      Sql := Copy(Request, Start, Start + 6)
           + Select
           + Copy(Request, pos('*', Request) + 1, Length(Request));
    end;
    Connect                  := TADOConnection.Create(application);
    Connect.ConnectionString := GetConnectionString;
    Connect.LoginPrompt      := False;
    try
      Connect.Connected := True;
      Connect.BeginTrans;
      try
        Qry            := TADOQuery.Create(Application);
        Qry.Connection := Connect;
        Qry.SQL.Text   := Sql;
        Qry.Prepared   := True;
        Qry.Open;
        try
          RecordCount := Qry.RecordCount;
          if not Qry.Eof then
          begin
            while not Qry.Eof do
            begin
              for Cpt := 0 to pred(Length(FieldsArray)) do
                ResultValue := ResultValue + TSLResult.Delimiter + VarToStr(Qry.FieldValues[FieldsArray[Cpt]]);
              ResultValue := Copy(ResultValue, 2, Length(ResultValue));
              TSLResult.Add(ResultValue);
              ResultValue := '';
              Qry.Next;
            end;
          end;
        finally
          Qry.active := False;
          Qry.Free;
        end;
        Connect.CommitTrans;
      except
        on E:Exception do
        begin
          Connect.RollbackTrans;
          Raise;
        end;
      end;
    finally
      Connect.Close;
      Connect.Free;
    end;
  end;
end;

procedure AdoQry.InsertUpdate;
var
  Connect : TADOConnection;
  Qry     : TADOQuery;
begin
  if     (ServerName <> '') // Nom du serveur
     and (DBName <> '')     // Nom de la BDD
     and (Request <> '')    // Requête
  then
  begin
    Connect                  := TADOConnection.Create(application);
    Connect.ConnectionString := GetConnectionString;
    Connect.LoginPrompt      := False;
    try
      Connect.Connected := True;
      Connect.BeginTrans;
      try
        Qry            := TADOQuery.Create(Application);
        Qry.Connection := Connect;
        Qry.SQL.Text   := Request;
        Qry.Prepared   := True;
        RecordCount    := Qry.ExecSQL;
        try

        finally
          Qry.active := False;
          Qry.Free;
        end;
        Connect.CommitTrans;
      except
        on E:Exception do
        begin
          Connect.RollbackTrans;
          Raise;
        end;
      end;
    finally
      Connect.Close;
      Connect.Free;
    end;
  end;
end;


{ Tools }

class function Tools.CaseFromString(Value: string; Values: array of string): integer;
var
  Cpt : Integer;
begin
  Result := -1;
  if (Value <> '') and (Length(Values) > -1) then
  begin
   for Cpt := Low(Values) to High(Values) do
    begin
      if Values[Cpt] = Value then
      begin
        Result := Cpt;
        Break;
      end;
    end;
  end;
end;

class function Tools.GetTypeFieldFromStringType(TypeString : string)  : tTypeField;
begin
  if TypeString <> '' then
  begin
    case Tools.CaseFromString(TypeString, ['INTEGER', 'SMALLINT', 'DOUBLE', 'RATE', 'EXTENDED', 'DATE', 'BLOB', 'DATA', 'COMBO', 'BOOLEAN']) of
      0..1 : Result := ttfInt;     {INTEGER , SMALLINT}
      2..4 : Result := ttfNumeric; {DOUBLE, RATE, EXTENDED}
      5    : Result := ttfDate;    {DATE}
      6..7 : Result := ttfMemo;    {BLOB, DATA}
      8    : Result := ttfCombo;   {COMBO}
      9    : Result := ttfBoolean; {BOOLEAN}
    else
       Result := ttfText;
    end;
  end else
    Result := ttfNone;
end;

class function Tools.GetFieldType(FieldName: string{$IF defined(APPSRV)}; ServerName, DBName : string{$IFEND !APPSRV}): tTypeField;
var
  FieldType : string;
  {$IF defined(APPSRV)}
  lAdoQry : AdoQry;
  {$IFEND !APPSRV}
begin
  if FieldName <> '' then
  begin
    {$IF not defined(APPSRV)}
    FieldType :=  ChampToType(FieldName);
    {$ELSE  !APPSRV}
    lAdoQry := AdoQry.Create;
    try
      lAdoQry.ServerName := ServerName;
      lAdoQry.DBName     := DBName;
      lAdoQry.FieldsList := 'DH_TYPECHAMP';
      lAdoQry.Request    := 'SELECT ' + lAdoQry.FieldsList + ' FROM DECHAMPS WHERE DH_NOMCHAMP =''' + FieldName + '''';
      lAdoQry.SingleTableSelect;
      FieldType := lAdoQry.TSLResult[0];
    finally
      lAdoQry.Free;
    end;
    {$IFEND !APPSRV}
    Result := Tools.GetTypeFieldFromStringType(FieldType);
  end else
    Result := ttfNone;
end;

class function Tools.GetDefaultValueFromtTypeField(FieldType : tTypeField) : string;
begin
  case FieldType of
    ttfNumeric : Result := '0';
    ttfInt     : Result := '0';
    ttfMemo    : Result := '';
    ttfBoolean : Result := '-';
    ttfDate    : Result := '2';
    ttfCombo   : Result := '';
    ttfText    : Result := '';
  end;
end;


class function Tools.iif(Const Expression, TruePart, FalsePart: Boolean): Boolean;
begin
	if Expression then
		Result := TruePart
	else
		Result := FalsePart;
end;

class function Tools.iif(Const Expression: Boolean; Const TruePart, FalsePart: Integer): Integer;
begin
	if Expression then
		Result := TruePart
	else
		Result := FalsePart;
end;

class function Tools.iif(Const Expression: Boolean; Const TruePart, FalsePart: Double): Double;
begin
	if Expression then
		Result := TruePart
	else
		Result := FalsePart;
end;

class function Tools.iif(Const Expression: Boolean; Const TruePart, FalsePart: String): String;
begin
	if Expression then
		Result := TruePart
	else
		Result := FalsePart;
end;

class function Tools.iif(Const Expression: Boolean; Const TruePart, FalsePart: Char): Char;
begin
	if Expression then
		Result := TruePart
	else
		Result := FalsePart;
end;

class function Tools.iif(Const Expression: Boolean; Const TruePart, FalsePart: TStringList): TStringList;
begin
	if Expression then
		Result := TruePart
	else
		Result := FalsePart;
end;

{$IF not defined(APPSRV)}
class function Tools.iif(Const Expression: Boolean; Const TruePart, FalsePart: TActionFiche): TActionFiche;
begin
	if Expression then
		Result := TruePart
	else
		Result := FalsePart;
end;
{$IFEND !APPSRV}

class function Tools.ReadTokenSt_(var S : string; Separator : string) : string;
var
  Cpt : integer;
begin
  Cpt:= Pos(Separator, S);
  if Cpt > 0 then
  begin
    Result := Copy(S, 1, Cpt-1);
    S      := Copy(S, Cpt + 1, Length(S)-Cpt);
  end else
  begin
    Result := S;
    S      := '';
  end;
end;

class function Tools.CountOccurenceString(const S: string; ToCount: string): integer;
var
  Pos : integer;
begin
  Result := 0;
  Pos    := PosEx(ToCount, S, 1);
  while Pos <> 0 do
  begin
    Inc(Result);
    Pos := PosEx(ToCount, S, Pos + 1);
  end;

end;

class function Tools.GetArgumentValue(Argument: string; const MyArg: String; const WithUpperCase: Boolean; const Separator: String): String;
var
	Critere	: String;
begin
	Result := '';
  while (Argument <> '') and (Result = '') do
  begin
    if WithUpperCase then
     	Critere := UpperCase(Tools.ReadTokenSt_(Argument, Separator))
    else
      Critere := Tools.ReadTokenSt_(Argument, Separator);
   	if (Pos(MyArg, Critere) > 0) and (Pos('=', Critere) <> 0) and (Trim(Copy(Critere, 1, Pos('=', Critere) - 1)) = MyArg) then
   	  Result := Trim(Copy(Critere, Pos('=', Critere) + 1, Length(Critere)));
	end;
end;

class function Tools.GetArgumentString(Argument: string; const MyArg: String; WithUpperCase: Boolean; const Separator: String): string;
begin
	if Pos(MyArg, Argument) > 0 then
		Result := VarToStr(GetArgumentValue(Argument, MyArg, WithUpperCase, Separator))
  else
   	Result := '';
end;

class function Tools.SetStrDateTimeFromStrUTCDateTime(UTCDateTime: string): string;
var
  Time       : TDateTime;
  Year       : integer;
  Month      : integer;
  Day        : integer;
  Hour       : integer;
  Minute     : integer;
  Second     : integer;
  Mlsecond   : integer;
  HourOffset : integer;
  MinOffset  : integer;
  WithMlSecond : Boolean;
  AddHour  : Boolean;
begin
  if UTCDateTime = '' then
    UTCDateTime := Tools.SetStrDateTimeToUTCDateTime('01/01/1900 00:00:00');
  WithMlSecond := UTCDateTime[20] = '.';
  Year     := StrToInt(copy(UTCDateTime,1 , 4));
  Month    := StrToInt(copy(UTCDateTime,6 , 2));
  Day      := StrToInt(copy(UTCDateTime,9 , 2));
  Hour     := StrToInt(copy(UTCDateTime,12, 2));
  Minute   := StrToInt(copy(UTCDateTime,15, 2));
  Second   := StrToInt(copy(UTCDateTime,18, 2));
  if not WithMlSecond then
  begin
    Mlsecond   := 0;
    HourOffset := StrToInt(copy(UTCDateTime,21, 2));
    MinOffset  := StrToInt(copy(UTCDateTime,24, 2));
    AddHour    := (UTCDateTime[20] = '+');
  end else
  begin
    Mlsecond   := StrToInt(copy(UTCDateTime,21, 3));
    HourOffset := 0;
    MinOffset  := 0;
    AddHour    := False;
  end;
  if AddHour then
  begin
    HourOffset := -1 * HourOffset;
    MinOffset  := -1 * MinOffset;
  end;
  Time := EncodeDateTime(Year, Month, Day, Hour, Minute, Second, Mlsecond);
  Time := IncHour(Time, hourOffset);
  Time := IncMinute(Time, minOffset);
  Result := DateTimeToStr(Time);
end;

class function Tools.SetStrDateTimeToUTCDateTime(stDateTime : string) : string;
begin
  if stDateTime <> '' then
    Result := FormatDateTime('yyyy-mm-dd', Int(StrToDateTime(stDateTime))) + 'T' +  FormatDateTime('hh:nn:ss.zzz', StrToDateTime(stDateTime)) + 'Z'
  else
    Result := '';
end;

class function Tools.SetStrUTCDateTimeToDateTime(stUTCDateTime : string) : string;
begin
  if stUTCDateTime <> '' then
  begin
    Result := copy(stUTCDateTime, 1, pos('T', stUTCDateTime) -1);
    Result := Format('%s/%s/%s', [copy(Result, 9, 2), copy(Result, 6, 2), copy(Result, 1, 4)]);
  end else
    Result := '';
end;


class function Tools.GetFileSize(FilePath: string; Size: tScaleSize): Extended;
var
  SearchFile : TSearchRec;
  FileSize   : Int64;
begin
  Result := 0;
  if FilePath <> '' then
  begin
    if (FindFirst(FilePath, faAnyFile, SearchFile) = 0) then
      FileSize := SearchFile.Size
    else
      FileSize := 0;
    case Size of
      tssKo : Result := (FileSize / 1024);
      TssMo : Result := (FileSize / 1048576);
    else
      Result := 0;
    end;
  end;
end;

class function Tools.StrFPoint_(Value : Extended) : string;
{$IF defined(APPSRV)}
var
  stValue : string;
{$IFEND !APPSRV}
begin
  {$IF defined(APPSRV)}
  stValue := FloatToStr(Value);
  Result  := StringReplace(stValue, ',', '.', [rfReplaceAll]);
  {$ELSE (APPSRV)}
  Result := StrFPoint(Value);
  {$IFEND !APPSRV}
end;

class function Tools.IsNumeric_(stValue : string) : boolean;
begin
  {$IF defined(APPSRV)}
  try
    Result := True;
    StrToFloat(stValue);
  except
    Result := False;
  end;
  {$ELSE (APPSRV)}
  Result := IsNumeric(stValue);
  {$IFEND !APPSRV}
end;

class function Tools.GetFieldsListFromPrefix(TablePrefix : string{$IF defined(APPSRV)}; ServerName, DBName : string{$IFEND !APPSRV}; Separator : string=',') : string;
var
  {$IF defined(APPSRV)}
  lAdoQry : AdoQry;
  Cpt     : integer;
  {$ELSE APPSRV}
  TobFieldsList : TOB;
  Cpt           : Integer;
  {$IFEND APPSRV}

  function GetSelect : string;
  begin
    Result := format('SELECT DH_NOMCHAMP FROM DECHAMPS WHERE DH_PREFIXE = ''%s'' ORDER BY DH_NUMCHAMP', [TablePrefix]);
  end;

begin
  if TablePrefix <> '' then
  begin
    {$IF defined(APPSRV)}
    lAdoQry := AdoQry.Create;
    try
      lAdoQry.ServerName := ServerName;
      lAdoQry.DBName     := DBName;
      lAdoQry.FieldsList := 'DH_NOMCHAMP';
      lAdoQry.Request    := GetSelect;
      lAdoQry.SingleTableSelect;
      for Cpt := 0 to pred(lAdoQry.TSLResult.Count) do
        Result := Result + Separator + lAdoQry.TSLResult[Cpt];
    finally
      lAdoQry.Free;
    end;
    {$ELSE !APPSRV}
    TobFieldsList := TOB.Create('_FIELD', nil, -1);
    try
      TobFieldsList.LoadDetailFromSQL(GetSelect);
      for Cpt := 0 to pred(TobFieldsList.Detail.count) do
        Result := Result + Separator + TobFieldsList.Detail[Cpt].GetString('DH_NOMCHAMP');
    finally
      FreeAndNil(TobFieldsList);
    end;
    {$IFEND !APPSRV}
    Result := Copy(Result, Length(Separator)+1, Length(Result));
  end else
    Result := '';
end;

class function Tools.GetTSlIndexFromFieldName(TslLine, FieldName : string; Separator : string=',') : integer;
var
  lTSlLine   : string;
  lFieldName : string;
begin
  Result := 0;
  if (TSlLine <> '') and (FieldName <> '') then
  begin
    lTSlLine := TslLine;
    while lTSlLine <> '' do
    begin
      inc(Result);
      lFieldName := Tools.ReadTokenSt_(lTSlLine, Separator);
      if lFieldName = FieldName then
        break;
    end;
  end;
end;

class function Tools.GetStValueFromTSl(TSlLine, FieldName : string) : string;
var
  Values : string;
  Value  : string;
  FindIt : boolean;
begin
  Result := '';
  if (TSlLine <> '') and (FieldName <> '') then
  begin
    Values := TSlLine;
    while Values <> '' do
    begin
      Value  := Tools.ReadTokenSt_(Values, ToolsTobToTsl_Separator);
      FindIt := Copy(Value, 1, Pos('=', Value) - 1) = FieldName;
      if FindIt then
      begin
        Result := Copy(Value, pos('=', Value) + 1, Length(Value));
        Break;
      end;
    end;
  end;
 end;

class function Tools.GetStValueFromTSl(TSlLine : string; Index : integer; Separator : string=',') : string;
var
  Values : string;
  Cpt    : integer;
begin
  Result := '';
  if (TSlLine <> '') and (Index >= 0) then
  begin
    Values := TSlLine;
    Cpt    := 0;
    while Cpt < Index do
    begin
      inc(Cpt);
      Result := Tools.ReadTokenSt_(Values, Separator);
    end;
  end;
end;

{$IF not defined(APPSRV)}
class procedure Tools.TobToTStringList(TobOrig: TOB; TSlResult: TStringList; Level : Integer=1);
var
  Cpt : Integer;

  procedure TSlAdd(TobOrigL : TOB; FirstLevelOnly : Boolean);
  var
    CptFields   : integer;
    CptSublevel : integer;
    NewLevel    : integer;
    FieldName   : string;
    FieldValue  : string;
    Fields      : string;
  begin
    Fields := ToolsTobToTsl_LevelName + IntToStr(Level) + '=' + TobOrigL.NomTable;
    for CptFields := 1 to TobOrigL.NombreChampReel do
    begin
      FieldName  := TobOrigL.GetNomChamp(CptFields);
      if FieldName <> '' then
      begin
        FieldValue := TobOrigL.GetString(FieldName);
        Fields := Fields + ToolsTobToTsl_Separator + FieldName + '=' + FieldValue;
      end;
    end;
    TSlResult.Add(Fields);
    if (TobOrigL.Detail.Count > 0) and (not FirstLevelOnly) then
    begin
      NewLevel := Level + 1;
      for CptSublevel := 0 to pred(TobOrigL.Detail.Count) do
        TobToTStringList(TobOrigL.Detail[CptSublevel], TSlResult, NewLevel);
    end;
  end;

begin
  if (assigned(TobOrig)) and (Assigned(TSlResult)) then
  begin
    TSlAdd(TobOrig, True);
    Inc(Level);
    for Cpt := 0 to pred(TobOrig.Detail.Count) do
      TSlAdd(TobOrig.detail[Cpt], False);
  end;
end;
{$IFEND !APPSRV}

class function Tools.GetTableNameFromTtn(Ttn: tTableName): string;
begin
  case Ttn of
    ttnChoixCod    : Result := 'CHOIXCOD';
    ttnCommun      : Result := 'COMMUN';
    ttnChoixExt    : Result := 'CHOIXEXT';
    ttnDevise      : Result := 'DEVISE';
    ttnModeRegl    : Result := 'MODEREGL';
    ttnPays        : Result := 'PAYS';
    ttnRib         : Result := 'RIB';
    ttnSection     : Result := 'SECTION';
    ttnTiers       : Result := 'TIERS';
    ttnCodePostaux : Result := 'CODEPOST';
    ttnContact     : Result := 'CONTACT';
    ttnEtabliss    : Result := 'ETABLISS';
    ttnModePaie    : Result := 'MODEPAIE';
    ttnGeneraux    : Result := 'GENERAUX';
    ttnJournal     : Result := 'JOURNAL';
    ttnRelance     : Result := 'RELANCE';
    ttnCorresp     : Result := 'CORRESP';
    ttnChancell    : Result := 'CHANCELL';
    ttnExercice    : Result := 'EXERCICE';
    ttnParamSoc    : Result := 'PARAMSOC';
    ttnEcriture    : Result := 'ECRITURE';
  else
    Result := '';
  end;
end;


class function Tools.GetTtnFromTableName(TableName: string): tTableName;
begin
  case CaseFromString(TableName, [  'CHOIXCOD', 'COMMUN'  , 'DEVISE'  , 'MODEREGL', 'PAYS'    , 'RIB'
                                  , 'SECTION' , 'TIERS'   , 'CODEPOST', 'CONTACT' , 'ETABLISS', 'MODEPAIE'
                                  , 'GENERAUX', 'JOURNAL' , 'RELANCE' , 'CORRESP' , 'CHANCELL', 'EXERCICE'
                                  , 'PARAMSOC', 'CHOIXEXT', 'ECRITURE'
                                 ]) of
    {CHOIXCOD} 0  : Result := ttnChoixCod;
    {COMMUN}   1  : Result := ttnCommun;
    {DEVISE}   2  : Result := ttnDevise;
    {MODEREGL} 3  : Result := ttnModeRegl;
    {PAYS}     4  : Result := ttnPays;
    {RIB}      5  : Result := ttnRib;
    {SECTION}  6  : Result := ttnSection;
    {TIERS}    7  : Result := ttnTiers;
    {CODEPOST} 8  : Result := ttnCodePostaux;
    {CONTACT}  9  : Result := ttnContact;
    {ETABLISS} 10 : Result := ttnEtabliss;
    {MODEPAIE} 11 : Result := ttnModePaie;
    {GENERAUX} 12 : Result := ttnGeneraux;
    {JOURNAL}  13 : Result := ttnJournal;
    {RELANCE}  14 : Result := ttnRelance;
    {CORRESP}  15 : Result := ttnCorresp;
    {CHANCELL} 16 : Result := ttnChancell;
    {EXERCICE} 17 : Result := ttnExercice;
    {PARAMSOC} 18 : Result := ttnParamSoc;
    {CHOIXEXT} 19 : Result := ttnChoixExt;
    {ECRITURE} 20 : Result := ttnEcriture;
  else
    Result := ttnNone;
  end;
end;

class function Tools.CanInsertedInTable(TableName: string{$IF defined(APPSRV)}; ServerName, DBName : string{$IFEND APPSRV}): Boolean;
{$IF defined(APPSRV)}
var
  AdoQryAut : AdoQry;
{$IFEND APPSRV}
begin
  if (TableName <> '') and (GetTtnFromTableName(TableName) <> ttnNone) then
  begin
    {$IF not defined(APPSRV)}
    Result := ExisteSql(Format('SELECT 1 FROM BTWSTABLEAUTO WHERE BWT_NOMTABLE = "%s" AND BWT_AUTORISEE = "X"', [TableName]))
    {$ELSE !APPSRV}
    AdoQryAut := AdoQry.create;
    try
      AdoQryAut.ServerName := ServerName;
      AdoQryAut.DBName     := DBName;
      AdoQryAut.FieldsList := 'BWT_NOMTABLE';
      AdoQryAut.Request    := Format('SELECT %s FROM BTWSTABLEAUTO WHERE BWT_NOMTABLE = ''%s'' AND BWT_AUTORISEE = ''X''', [AdoQryAut.FieldsList, TableName]);
      AdoQryAut.SingleTableSelect;
      Result := AdoQryAut.RecordCount = 1;
    finally
      AdoQryAut.Free;
    end;
    {$IFEND !APPSRV}
  end else
    Result := True;
end;

class function Tools.GetPSocTreeToExport(OnlyAccounting : Boolean=False) : string;
var
  Sep : string;
begin
  {$IF not defined(APPSRV)}
  Sep := '"';
  {$ELSE !APPSRV}
  Sep := '''';
  {$IFEND !APPSRV}
  Result := Format('(soc_tree like %s001;001;%%s)', [Sep, Sep]);
  if not OnlyAccounting then
    Result := Result + Format('(soc_tree like %s001;027;%%s)', [Sep, Sep])
                     + Format('(soc_tree like %s001;035;%%s)', [Sep, Sep])
                     + Format('(soc_tree like %s001;012;%%s)', [Sep, Sep])
                     + Format('(soc_tree like %s001;002;%%s)', [Sep, Sep])
                     + Format('(soc_tree like %s001;023;%%s)', [Sep, Sep])
                     + Format('(soc_tree like %s001;006;%%s)', [Sep, Sep])
                     + Format('(soc_tree like %s001;005;%%s)', [Sep, Sep])
                     + Format('(soc_tree like %s001;014;%%s)', [Sep, Sep])
                     + Format('(soc_tree like %s001;031;%%s)', [Sep, Sep])
                     + Format('(soc_tree like %s001;013;%%s)', [Sep, Sep])
                     + Format('(soc_tree like %s001;018;%%s)', [Sep, Sep])
                     + Format('(soc_tree like %s001;021;%%s)', [Sep, Sep])
                     ;
end;

class function Tools.SetTRAFileFromTSl(lTSL : TStringList) : Boolean;
begin
  Result := False;
  if Assigned(lTSL) and (lTSL.Count > 0) then
  begin

  end;
end;

class function Tools.FormatValue(Value : string; Align : tTypeAlign; iLength : integer; NbDec : Integer=0; TypeDate : tFormatValueTypeDate=tvtNone) : string;
var
  sIntValue : string;
  sDecValue : string;
  Cpt       : integer;
begin
  if iLength > 0 then
  begin
    Result := Value;
    if NbDec > 0 then
    begin
      if Pos(DecimalSeparator, Result) = 0 then
      begin
        Result := Result + ',';
        Cpt    := 0;
        while Cpt < NbDec do
        begin
          inc(Cpt);
          Result := Result + '0';
        end;
      end else
      begin
        sIntValue := Copy(Result, 1, Pos(DecimalSeparator, Result) -1);
        sDecValue := Copy(Result, Pos(DecimalSeparator, Result) +1 , NbDec);
        while Length(sDecValue) < NbDec do
          sDecValue := sDecValue + '0';
        Result := sIntValue + DecimalSeparator + sDecValue;
      end;
    end else
    if TypeDate = tvtDate then
      Result := FormatDateTime('ddmmyyyy', StrToDate(Value))
    else if TypeDate = tvtDateTime then
      Result := FormatDateTime('ddmmyyyyhhnn', StrToDateTime(Value));

    while Length(Result) < iLength do
    begin
      case Align of
        traaLeft  : Result := Result + ' ';
        traaRigth : Result := ' ' + Result;
      end;
    end;
  end else
    Result := '';
end;

class function Tools.CompressFile(FullPath : string) : string;
var
  Zip : TZip;
  FileList : TStrings;
begin
  Result := '';
  if FullPath <> '' then
  begin
    Zip := TZip.create(nil);
    try
      FileList := TStringList.Create ;
      try
        FileList.Add(FullPath);
        Zip.FileSpecList := FileList;
        Zip.Filename     := FullPath + '.zip';
        if Zip.Add > 0 then
          Result := Zip.Filename;
      finally
        FreeAndNil(FileList);
      end;
    finally
      FreeAndNil(Zip);
    end;
  end;
end;

class function Tools.UnCompressFile(ZipPath, ZipFileName : string) : integer;
var
  Zip : TZip;
  Cpt : integer;
begin
  Result := 0;
  if (ZipPath <> '') and (ZipFileName <> '') then
  begin
    Zip := TZip.create(nil);
    try
      Zip.FileSpecList.Clear;
      //Zip.ExtractOptions := [oeUpdate];
      Zip.ExtractPath := ZipPath;
      Zip.Filename    := ZipPath + ZipFileName;
      if Zip.Count >= 0 then
      begin
        for Cpt := 0 to pred(Zip.Count) do
          Zip.FileSpecList.Add(Zip.FileInfos[Cpt].Filename);
      Result := Zip.Extract;
      end;
    finally
      FreeAndNil(Zip);
    end;

  end;
end;
  
class procedure Tools.FileCut(FullPath : string; MaxSizeBytes : integer; TSLResult : TStringList; KeepOriginFile : boolean=True);
var
  FileStream : TFileStream;
  Cpt        : integer;
  Qty        : Integer;

  procedure DoCut(Index : Integer);
  var
    FileExtension : string;
    SplitFileName : string;
    sStream       : TFileStream;
    DelPos        : Integer;
  begin
    DelPos        := LastDelimiter('.', FullPath);
    FileExtension := Copy(FullPath, DelPos, Length(FullPath));
    SplitFileName := Copy(FullPath, 1, DelPos -1) + '_' + FormatFloat('000', Index) + FileExtension;
    sStream       := TFileStream.Create(SplitFileName, fmCreate);
    try
      if FileStream.Size - FileStream.Position < MaxSizeBytes then
        MaxSizeBytes := FileStream.Size - FileStream.Position;
      sStream.CopyFrom(FileStream, MaxSizeBytes);
    finally
      TSLResult.Add(SplitFileName);
      sStream.Free;
    end;
  end;

begin
  if (FullPath <> '') and (MaxSizeBytes > 0 ) then
  begin
    FileStream := TFileStream.Create(FullPath, fmOpenRead);
    try
      if FileStream.Size > MaxSizeBytes then
      begin
        Qty := 0;
        for Cpt := 0 to pred(Trunc(FileStream.Size / MaxSizeBytes)) do
        begin
          inc(Qty);
          DoCut(Cpt);
        end;
        DoCut(Qty);
        if not KeepOriginFile then
          DeleteFile(FullPath);
      end else
        TSLResult.Add(FullPath);
    finally
      FileStream.free;
    end;
  end;
end;

class function Tools.GetKoFromMo(Mo : integer) : Integer;
begin
  Result := Mo * 1048576;
end;

class function Tools.DeleteDirectroy(Path : string) : boolean;
var
  iIndex    : Integer;
  SearchRec : TSearchRec;
  LocalPath : string;
  sFileName : string;
begin
  if Path <> '' then
  begin
    if Copy(Path, Length(Path), 1) <> '\' then
      LocalPath := Path + '\*.*'
    else
      LocalPath := Path + '*.*';
    iIndex := FindFirst(LocalPath, faAnyFile, SearchRec);
    while iIndex = 0 do
    begin
      sFileName := ExtractFileDir(LocalPath) + '\' + SearchRec.Name;
      if SearchRec.Attr = faDirectory then
      begin
      if     (SearchRec.Name <> '' )
         and (SearchRec.Name <> '.')
         and (SearchRec.Name <> '..')
      then
         DeleteFile(sFileName);
      end else
      begin
        if SearchRec.Attr <> faArchive then
          FileSetAttr(sFileName, faArchive);
        DeleteFile(sFileName);
      end;
      iIndex := FindNext(SearchRec);
    end;
    FindClose(SearchRec);
    Result := RemoveDir(LocalPath);
  end else
    Result := False;
end;

class function Tools.IsRecordableDocument(DocType, Establishment : string{$IF defined(APPSRV)}; ServerName, DBName : string{$IFEND !APPSRV}) : boolean;
var
  AccState : string;
  {$IF defined(APPSRV)}
  lAdoQry  : AdoQry;
  {$IFEND !APPSRV}

  {$IF defined(APPSRV)}
  function GetAccType(Prefix : string) : string;
  var
    Sql : string;
  begin
    lAdoQry.TSLResult.Clear;
    lAdoQry.RecordCount := 0;
    lAdoQry.ServerName  := ''; //ServerName;
    lAdoQry.DBName      := ''; //DBName;
    lAdoQry.FieldsList  := Prefix + '_TYPEECRCPTA';
    Sql                 := Format('SELECT %s FROM %s WHERE %s_NATUREPIECEG = ''%s''', [lAdoQry.FieldsList, Tools.iif(Prefix = 'GPC', 'PARPIECECOMPL', 'PARPIECE'), Prefix, DocType]);
    if Prefix = 'GPC' then
      Sql := Sql + Format(' AND GPC_ETABLISSEMENT = ''%s''', [Establishment]);
    lAdoQry.Request    := Sql;
    lAdoQry.SingleTableSelect;
    if lAdoQry.RecordCount > 0 then
      Result := lAdoQry.TSLResult[0]
    else
      Result := '';
  end;
  {$IFEND !APPSRV}

begin
  if DocType <> '' then
  begin
    {$IF defined(APPSRV)}
    lAdoQry := AdoQry.Create;
    try
      AccState := GetAccType('GPC');
      if AccState = '' then
        AccState := GetAccType('GPP');
    finally
      lAdoQry.Free;
    end;
    {$ELSE !APPSRV}
    AccState := GetInfoParPieceCompl(DocType, Establishment, 'GPC_TYPEECRCPTA');
    if AccState = '' then
      AccState := GetInfoParPiece(DocType, 'GPP_TYPEECRCPTA');
    {$IFEND !APPSRV}
    Result := ((AccState <> '') and (AccState <> 'RIE'));
  end else
    Result := False;
end;

class function Tools.EvalueDateJJMMYYY(sDate : string) : TDateTime;
var
  dd : word;
  mm : Word;
  yy : Word ;
begin
  if sDate <> '' then
  begin
    dd     := StrToInt(Copy(sDate,1,2));
    mm     := StrToInt(Copy(sDate,3,2));
    yy     := StrToInt(Copy(sDate,5,4));
    Result := Encodedate(yy,mm,dd);
  end else
    Result := 2;
end;

class procedure Tools.DecodeAccDocReferency(DocReferency : string; var DocType : string; var Stump : string; var DocDate : TDateTime; var DocNumber : integer; var Index : integer);
begin
  if DocReferency <> '' then
  begin
    DocType   := Tools.ReadTokenSt_(DocReferency, ';');
    Stump     := Tools.ReadTokenSt_(DocReferency, ';');
    DocDate   := Tools.EvalueDateJJMMYYY(Tools.ReadTokenSt_(DocReferency, ';'));
    DocNumber := StrToInt(Tools.ReadTokenSt_(DocReferency, ';'));
    Index     := StrToInt(Tools.ReadTokenSt_(DocReferency, ';'));
  end;
end;

class function Tools.GetParamSocSecur_(PSocName : string; DefaultValue : string{$IFDEF APPSRV}; ServerName, FolderName : string{$ENDIF APPSRV}) : string;
{$IFDEF APPSRV}
var
  AdoQryL : AdoQry;
{$ENDIF APPSRV}
begin
  Result := DefaultValue;
  {$IFDEF APPSRV}
  if (ServerName <> '') and (FolderName <> '') and (PSocName <> '') then
  begin
    AdoQryL := AdoQry.Create;
    try
      AdoQryL.ServerName  := ServerName;
      AdoQryL.DBName      := FolderName;
      AdoQryL.FieldsList  := 'SOC_DATA';
      AdoQryL.Request     := Format('SELECT %s FROM PARAMSOC WHERE SOC_NOM = ''%s''', [AdoQryL.FieldsList, PSocName]);
      AdoQryL.SingleTableSelect;
      if AdoQryL.RecordCount > 0 then
        Result := AdoQryL.TSLResult[0];
    finally
      AdoQryL.free;
    end;
  end;
  {$ELSE APPSRV}
  Result := GetParamSocSecur(PSocName, DefaultValue);
  {$ENDIF APPSRV}
end;

class function Tools.CastDateTimeForQry(lDate: TDateTime): string;
begin
  Result := FormatDateTime('yyyymmdd', lDate);                    
end;

end.

