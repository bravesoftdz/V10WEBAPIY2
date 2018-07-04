unit UConnectWSCEGID;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, WinHttp_TLB, XMLDoc, xmlintf, DateUtils,
  UConnectWSConst
  {$IF not defined(APPSRV)}
    , UTOB, HMsgBox, ParamSoc
  {$IFEND !APPSRV}
;

type
  TconnectCEGID = class(TObject)
  private
    factive: Boolean;
    fServer: string;
    fport: integer;
    fDossier: string;
    function GetPort: string;
    procedure SetPort(const Value: string);
    procedure SetDossier(const Value: string);
    procedure SetServer(const Value: string);
    function AppelEntriesWS(DocType: string; DocNumber: integer; TheXml: WideString; var NumDocOut: Integer): Boolean;
    {$IF not defined(APPSRV)}
    procedure RemplitTOBDossiers(ListeDoss: TOB; HTTPResponse: WideString);
    procedure RemplitTOBExercices(TOBexer: TOB; HTTPResponse: WideString);
    {$IFEND !APPSRV}
    function GetStartUrl: string;
    
  public
    constructor create;
    destructor destroy; override;
    {$IF not defined(APPSRV)}
    procedure GetDossiers(var ListeDoss: TOB; var TheResponse: WideString);
    procedure GetExCpta(TOBexer: TOB);
    {$IFEND !APPSRV}

    //
    property CEGIDServer: string read fServer write SetServer;
    property CEGIDPORT: string read GetPort write SetPort;
    property DOSSIER: string read fDossier write SetDossier;
    property IsActive: boolean read factive;
  end;

  TGetParamWSCEGID = class(TObject)
  public
    class function GetCodeFromWsEt(WsEt: T_WSEntryType): string;
    class function ConnectToY2: Boolean;
    class function GetPSoc(PSocType: T_WSPSocType): string;
  end;

  TSendEntryY2 = class(Tobject)
  private
    function EncodeDocType(TypePiece: string): string;
    function EncodeEntryType(EntryType: string): string;
    function EnregistreInfoCptaY2(WsEt: T_WSEntryType; TslLine: string; TheNumDoc: Integer): boolean;
    function EncodeAmount(TslLine: string): string;
    function EncodeDate(TslLine, DateFieldName: string): string;
    function GetFirstIndiceEcr(TSlEcr: TStringList): Integer;
    function ConstitueEntries(TSlEcr: TStringList): WideString;
//    function GetStValueFromTSl(TSlLine, FieldName : string) : string;
    function EncodeAxis(AxisCode: string): string;
    procedure SetCegidConnectParameters(CegidConnect : TconnectCEGID);
  public
    ServerName: string;
    DBName: string;
    SendCegid: boolean;

    {$IF not defined(APPSRV)}
    function SendEntryCEGID(WsEt: T_WSEntryType; TOBecr: TOB; DocType: string; DocNumber: integer): boolean; overload;
  	{$IFEND !APPSRV}
    function SendEntryCEGID(WsEt: T_WSEntryType; TSlEcr: TStringList; DocType: string; DocNumber: integer): boolean; overload;
    function SendAccountingParameters(TSLFullPathFile : TStringList): boolean;
  end;

implementation

uses
  db, uLkJSON, CommonTools
  {$IFNDEF DBXPRESS}
    , dbtables
  {$ELSE DBXPRESS}
    , uDbxDataSet
  {$ENDIF DBXPRESS}
  {$IF not defined(APPSRV)}
    , Hctrls, Aglinit, Hent1, wCommuns, UtilPGI
  {$IFEND !APPSRV}
;

function StringToStream(const AString: string): Tstream;
var
  SS: TStringStream;
begin
  Result := nil;
  SS := TStringStream.Create(AString);
  try
    SS.Position := 0;
    result.CopyFrom(SS, SS.Size);  //This is where the "Abstract Error" gets thrown
  finally
    SS.Free;
  end;
end;

function DateTime2Tdate(TheDateTime: TdateTime): string;
var
  YY, MM, DD, Hours, Mins, secs, milli: Word;
begin
  DecodeDateTime(TheDateTime, YY, MM, DD, Hours, Mins, secs, milli);
  Result := Format('%4d-%0.2d-%0.2dT%0.2d:%0.2d:%0.2d', [YY, MM, DD, Hours, Mins, secs]);
end;

function TDate2DateTime(OneTdate: string): TDateTime;
var
  TheTDate: string;
  YYYY, MM, DD: string;
  PDATE: string;
  TheTime: string;
  IposT, II: Integer;
begin
  {$IF not defined(APPSRV)}
  Result := iDate1900;
  {$ELSE !APPSRV}
  Result := 2;
  {$IFEND !APPSRV}
  IposT := Pos('T', OneTdate);
  if IposT > 0 then
  begin
    II := 0;
    TheTDate := Copy(OneTdate, 1, IposT - 1);
    TheTime := Copy(OneTdate, IposT + 1, Length(OneTdate) - 1);
    repeat
      PDATE := Tools.ReadTokenSt_(TheTDate, '-');
      if PDATE <> '' then
      begin
        if II = 0 then
        begin
          YYYY := PDATE;
          Inc(II);
        end
        else if II = 1 then
        begin
          MM := PDATE;
          Inc(II);
        end
        else
        begin
          DD := PDATE;
          Inc(II);
        end;
      end;
    until PDATE = '';
    if (YYYY <> '') and (MM <> '') and (DD <> '') then
    begin
      Result := StrToDateTime(DD + '/' + MM + '/' + YYYY + ' ' + TheTime);
    end;
  end;
end;

function TSendEntryY2.EncodeDocType(TypePiece: string): string;
begin
  case Tools.CaseFromString(TypePiece, ['N', 'S']) of
    {N}         0:
      Result := 'Normal';
    {S}         1:
      Result := 'Simulation';
  end;
end;

function TSendEntryY2.EncodeEntryType(EntryType: string): string;
begin
  case Tools.CaseFromString(EntryType, ['AC', 'AF', 'ECC', 'FC', 'FF', 'OC', 'OD', 'OF', 'RC', 'RF']) of
    {AC }         0:
      Result := 'CustomerCredit';
    {AF }         1:
      Result := 'ProviderCredit';
    {ECC}         2:
      Result := 'ExchangeDifference';
    {FC }         3:
      Result := 'CustomerInvoice';
    {FF }         4:
      Result := 'ProviderInvoice';
    {OC }         5:
      Result := 'CustomerDeposite';
    {OD }         6:
      Result := 'SpecificOperation';
    {OF }         7:
      Result := 'ProviderDeposite';
    {RC }         8:
      Result := 'CustomerPayment';
    {RF }         9:
      Result := 'ProviderPayment';
  end;
end;

function TSendEntryY2.EnregistreInfoCptaY2(WsEt: T_WSEntryType; TslLine: string; TheNumDoc: Integer): boolean;
var
  eJournal: string;
  eExercice: string;
  eDateComptable: TDateTime;
  eEntity: integer;
  eNumeroPiece: integer;
  {$IF defined(APPSRV)}
  AdoQryL: AdoQry;
  {$IFEND APPSRV}

  function GetSep: string;
  begin
  {$IF defined(APPSRV)}
    Result := '''';
  {$ELSE APPSRV}
    Result := '"';
  {$IFEND APPSRV}
  end;

  function GetWhere: string;
  begin
    Result := ' WHERE BE0_ENTITY      = ' + IntToSTr(eEntity) + '   AND BE0_JOURNAL     = ' + GetSep + eJournal + GetSep + '   AND BE0_EXERCICE    = ' + GetSep + eExercice + GetSep + '   AND BE0_NUMEROPIECE = ' + IntToStr(eNumeroPiece);
  end;

  function GetSqlExist: string;
  begin
    Result := 'SELECT 1 FROM BTPECRITURE' + GetWhere;
  end;

  function GetSqlUpdate: string;
  begin
    Result := 'UPDATE BTPECRITURE SET BE0_REFERENCEY2 = ' + IntToStr(TheNumDoc) + GetWhere;
  end;

  function GetSqlInsert: string;
  begin
    Result := 'INSERT INTO BTPECRITURE' + ' (  BE0_ENTITY' + '  , BE0_JOURNAL' + '  , BE0_EXERCICE' + '  , BE0_NUMEROPIECE' + '  , BE0_DATECOMPTABLE' + '  , BE0_REFERENCEY2' + '  , BE0_TYPE' + ' )  VALUES' + ' (' + IntToStr(eEntity) + ', ' + GetSep + eJournal + GetSep + ', ' + GetSep + eExercice + GetSep + ', ' + IntToStr(eNumeroPiece) + ', ' + GetSep + DateTimeToStr(eDateComptable) + GetSep + ', ' + IntToStr(TheNumDoc) + ', ' + GetSep + TGetParamWSCEGID.GetCodeFromWsEt(WsEt) + GetSep + ' )';
  end;

begin
  Result := False;
  eEntity := StrToInt(Tools.GetStValueFromTSl(TslLine, 'E_ENTITY'));
  eJournal := Tools.GetStValueFromTSl(TslLine, 'E_JOURNAL');
  eExercice := Tools.GetStValueFromTSl(TslLine, 'E_EXERCICE');
  eDateComptable := StrToDateTime(Tools.GetStValueFromTSl(TslLine, 'E_DATECOMPTABLE'));
  eNumeroPiece := StrToInt(Tools.GetStValueFromTSl(TslLine, 'E_NUMEROPIECE'));
  if eJournal <> '' then
  begin
    {$IF defined(APPSRV)}
    AdoQryL := AdoQry.Create;
    try
      { Test si existe }
      AdoQryL.ServerName := ServerName;
      AdoQryL.DBName := DBName;
      AdoQryL.FieldsList := 'BE0_ENTITY';
      AdoQryL.Request := GetSqlExist;
      AdoQryL.SingleTableSelect;
      { Exécute l'Update ou l'Insert }
      AdoQryL.RecordCount := 0;
      AdoQryL.TSLResult.Clear;
      AdoQryL.FieldsList := '';
      AdoQryL.Request := Tools.iif(AdoQryL.RecordCount = 1, GetSqlUpdate, GetSqlInsert);
      AdoQryL.InsertUpdate;
      Result := (AdoQryL.RecordCount = 1);
    finally
      AdoQryL.free;
    end;
    {$ELSE APPSRV}
    if ExisteSQL(GetSqlExist) then
      Result := (ExecuteSql(GetSqlUpdate) = 1)
    else
      Result := (ExecuteSql(GetSqlInsert) = 1);
    {$IFEND APPSRV}
  end;
end;

function TSendEntryY2.GetFirstIndiceEcr(TSlEcr: TStringList): Integer;
var
  CptIndice: integer;
begin
  Result := -1;
  for CptIndice := 0 to pred(TSlEcr.count) do
  begin
    if Pos('=ECRITURE' + ToolsTobToTsl_Separator, TSlEcr[CptIndice]) > 0 then
    begin
      Result := CptIndice;
      Break;
    end;
  end;
end;

function TSendEntryY2.EncodeAmount(TslLine: string): string;
var
  DebAmount: Double;
  CredAmount: double;
begin
  DebAmount := StrToFloat(Tools.GetStValueFromTSl(TslLine, 'E_DEBITDEV'));
  CredAmount := StrToFloat(Tools.GetStValueFromTSl(TslLine, 'E_CREDITDEV'));
  Result := Tools.StrFPoint_(Tools.iif(DebAmount <> 0, DebAmount, CredAmount));
end;

function TSendEntryY2.EncodeDate(TslLine, DateFieldName: string): string;
var
  DateValue: TDateTime;
begin
  DateValue := StrToDateTime(Tools.GetStValueFromTSl(TslLine, DateFieldName));
  Result := DateTime2Tdate(Tools.iif((DateValue < 2), 2, DateValue));
end;

function TSendEntryY2.EncodeAxis(AxisCode: string): string;
begin
  case Tools.CaseFromString(AxisCode, ['A1', 'A2', 'A3', 'A4', 'A5']) of
    {A1}         0:
      Result := 'One';
    {A2}         1:
      Result := 'Two';
    {A3}         2:
      Result := 'Three';
    {A4}         3:
      Result := 'Four';
    {A5}         4:
      Result := 'Five';
  end;
end;

procedure TSendEntryY2.SetCegidConnectParameters(CegidConnect : TconnectCEGID);
  {$IF defined(APPSRV)}
var
  AdoQryL: AdoQry;
  {$IFEND APPSRV}
begin
  {$IF not defined(APPSRV)}
  CegidConnect.CEGIDServer := TGetParamWSCEGID.GetPSoc(wspsServer);
  CegidConnect.CEGIDPORT   := TGetParamWSCEGID.GetPSoc(wspsPort);
  CegidConnect.DOSSIER     := TGetParamWSCEGID.GetPSoc(wspsFolder);
  {$ELSE APPSRV}
  AdoQryL := AdoQry.Create;
  try
    AdoQryL.ServerName := ServerName;
    AdoQryL.DBName     := DBName;
    AdoQryL.FieldsList := 'SOC_DATA';
    AdoQryL.Request    := Format('SELECT %s FROM PARAMSOC WHERE SOC_NOM IN (''%s'', ''%s'', ''%s'') ORDER BY SOC_NOM DESC', [AdoQryL.FieldsList, WSCDS_SocServer, WSCDS_SocNumPort, WSCDS_SocCegidDos]);
    AdoQryL.SingleTableSelect;
    if AdoQryL.RecordCount = 3 then
    begin
      CegidConnect.CEGIDServer := AdoQryL.TSLResult[0];
      CegidConnect.CEGIDPORT   := AdoQryL.TSLResult[1];
      CegidConnect.DOSSIER     := AdoQryL.TSLResult[2];
    end;
  finally
    AdoQryL.free;
  end;
  {$IFEND APPSRV}
end;
  
{ TSLecr est une structure issue de la TobEcr à plat.
  Exemple pour une tob avec 3 lignes d'écriture et de l'analytique sur la 2ème ligne :
    ^LEVEL1=COMPTABILITE
    ^LEVEL2=ECRITURE^E_AFFAIRE=^E_ANA=-^E_AUXILIAIRE=CJT0100000^...
    ^LEVEL2=ECRITURE^E_AFFAIRE=^E_ANA=-^E_AUXILIAIRE=CJT0100000^...
    ^LEVEL3=A1'
    ^LEVEL4=ANALYTIQ^Y_AFFAIRE=^Y_AUXILIAIRE=^Y_AXE=A1^...
    ^LEVEL4=ANALYTIQ^Y_AFFAIRE=^Y_AUXILIAIRE=^Y_AXE=A1^...
    ^LEVEL3=A2'
    ^LEVEL3=A3'
    ^LEVEL3=A4'
    ^LEVEL3=A5'
    ^LEVEL2=ECRITURE^E_AFFAIRE=^E_ANA=-^E_AUXILIAIRE=^..
}
function TSendEntryY2.ConstitueEntries(TSlEcr: TStringList): WideString;
var
  XmlDoc: IXMLDocument;
  Root: IXMLNode;
  NodeDoc: IXMLNode;
  Entries: IXMLNode;
  Entry: IXMLNode;
  Amounts: IXMLNode;
  EntryAmount: IXMLNode;
  Analytics: IXMLNode;
  Analytic: IXMLNode;
  Sections: IXMLNode;
  Setting: IXMLNode;
  N1: IXMLNode;
  FirstIndice: integer;
  Cpt: integer;
  EcrLevelName: string;
  AnaLevelName: string;

  function GetLevelName(TableName: string): string;
  var
    CptIndice: integer;
  begin
    for CptIndice := 0 to pred(TSlEcr.count) do
    begin
      if Pos('=' + TableName + ToolsTobToTsl_Separator, TSlEcr[CptIndice]) > 0 then
      begin
        Result := Copy(TSlEcr[CptIndice], 1, Pos('=', TSlEcr[CptIndice]) - 1);
        Break;
      end;
    end;
  end;

  procedure AddAnalytics(ParentNode: IXMLNode; CurrentIndex: integer);
  var
    CptE: integer;
    StartAna: Boolean;
    LevelName: string;
  begin
    StartAna := False;
    Analytics := Amounts.AddChild('Analytics');    // <- Noeud Analytics> ->
    for CptE := CurrentIndex to pred(TSlEcr.Count) do
    begin
      LevelName := copy(TSlEcr[CptE], 1, Pos('=', TSlEcr[CptE]) - 1);
      if (not StartAna) and (LevelName = AnaLevelName) then
        StartAna := True
      else if (StartAna) and (LevelName = EcrLevelName) then
        Break;
      if (StartAna) and (LevelName = AnaLevelName) then
      begin
        Analytic := Analytics.AddChild('EntryAnalytic');    // <- Noeud EntryAnalytic>> ->
        N1 := Analytic.AddChild('Amount');
        N1.Text := Tools.StrFPoint_(StrTofloat(Tools.GetStValueFromTSl(TSlEcr[CptE], 'Y_DEBIT')) + StrTofloat(Tools.GetStValueFromTSl(TSlEcr[CptE], 'Y_CREDIT')));
        N1 := Analytic.AddChild('Axis');
        N1.Text := EncodeAxis(Tools.GetStValueFromTSl(TSlEcr[CptE], 'Y_AXE'));
        N1 := Analytic.AddChild('Percent');
        N1.Text := '0';
        Sections := Analytic.AddChild('Sections');    // <- Noeud EntryAnalytic>> ->
        Sections.Attributes['xmlns:a'] := 'http://schemas.microsoft.com/2003/10/Serialization/Arrays';
        N1 := Sections.AddChild('a:string');
        N1.Text := Tools.GetStValueFromTSl(TSlEcr[CptE], 'Y_SECTION');
      end;
    end;
    if not StartAna then
      Analytics.Attributes['i:nil'] := 'true';
  end;

begin
  Result := '';
  FirstIndice := GetFirstIndiceEcr(TSlEcr);
  EcrLevelName := GetLevelName('ECRITURE');
  AnaLevelName := GetLevelName('ANALYTIQ');
  XmlDoc := NewXMLDocument();
  XmlDoc.Options := [doNodeAutoIndent];
  try
    Root := XmlDoc.AddChild('EntryParameter');
    NodeDoc := Root.AddChild('Document');    // <- Noeud Document ->
    // --- Sur le <document>
    N1 := NodeDoc.Addchild('AccountingDate');
    N1.Text := DateTime2Tdate(StrToDateTime(Tools.GetStValueFromTSl(TSlEcr[FirstIndice], 'E_DATECOMPTABLE')));
    N1 := NodeDoc.Addchild('BusinessCenter');
    N1.Text := Tools.GetStValueFromTSl(TSlEcr[FirstIndice], 'E_ETABLISSEMENT');
    N1 := NodeDoc.Addchild('Currency');
    N1.Text := Tools.GetStValueFromTSl(TSlEcr[FirstIndice], 'E_DEVISE');
    N1 := NodeDoc.Addchild('CurrencyRate');
    N1.Text := Tools.StrFPoint_(StrToFloat(Tools.GetStValueFromTSl(TSlEcr[FirstIndice], 'E_TAUXDEV')));
    N1 := NodeDoc.Addchild('DocumentType');
    N1.Text := EncodeDocType(Tools.GetStValueFromTSl(TSlEcr[FirstIndice], 'E_QUALIFPIECE'));
    Entries := NodeDoc.Addchild('Entries'); // <- Noeud Entries ->
    for Cpt := FirstIndice to pred(TSlEcr.Count) do
    begin
      if copy(TSlEcr[Cpt], 1, Pos('=', TSlEcr[Cpt]) - 1) = EcrLevelName then
      begin
        Entry := Entries.Addchild('Entry');   // <- Noeud Entry ->
        N1 := Entry.AddChild('AmountDirection');
        N1.Text := Tools.iif(StrToFloat(Tools.GetStValueFromTSl(TSlEcr[Cpt], 'E_DEBITDEV')) <> 0, 'Debit', 'Credit');
        Amounts := Entry.AddChild('Amounts');    // <- Noeud Amounts ->
        EntryAmount := Amounts.AddChild('EntryAmount');    // <- Noeud EntryAmount ->
        N1 := EntryAmount.AddChild('Amount');
        N1.Text := EncodeAmount(TSlEcr[Cpt]);
        N1 := EntryAmount.AddChild('DueDate');
        N1.Text := EncodeDate(TSlEcr[Cpt], 'E_DATEECHEANCE');
        N1 := EntryAmount.AddChild('Iban');
        N1.Text := Tools.GetStValueFromTSl(TSlEcr[Cpt], 'E_RIB');
        N1 := EntryAmount.AddChild('PaymentMode');
        N1.Text := Tools.GetStValueFromTSl(TSlEcr[Cpt], 'E_MODEPAIE');
        if Tools.GetStValueFromTSl(TSlEcr[Cpt], 'E_MODEPAIE') <> '' then
          N1 := EntryAmount.AddChild('SepaCreditorIdentifier');
        N1.Attributes['i:nil'] := 'true';
        N1 := EntryAmount.AddChild('UniqueMandateReference');
        if Tools.GetStValueFromTSl(TSlEcr[Cpt], 'E_ANA') = 'X' then
          AddAnalytics(Entry, Cpt)
        else
        begin
          Analytics := Amounts.AddChild('Analytics');    // <- Noeud Analytics> ->
          Analytics.Attributes['i:nil'] := 'true';
        end;
        N1 := Entry.AddChild('Description');
        N1.Text := Tools.GetStValueFromTSl(TSlEcr[Cpt], 'E_LIBELLE');
        N1 := Entry.AddChild('ExternalDateReference');
        N1.Text := EncodeDate(TSlEcr[Cpt], 'E_DATEREFEXTERNE');
        N1 := Entry.AddChild('ExternalReference');
        N1.Text := Tools.GetStValueFromTSl(TSlEcr[Cpt], 'E_REFEXTERNE');
        N1 := Entry.AddChild('GeneralAccount');
        N1.Text := Tools.GetStValueFromTSl(TSlEcr[Cpt], 'E_GENERAL');
        N1 := Entry.AddChild('InternalReference');
        N1.Text := Tools.GetStValueFromTSl(TSlEcr[Cpt], 'E_REFINTERNE');
        N1 := Entry.AddChild('SubsidiaryAccount');
        N1.Text := Tools.GetStValueFromTSl(TSlEcr[Cpt], 'E_AUXILIAIRE');
      end;
    end;
    N1 := NodeDoc.Addchild('EntryType');
    N1.Text := EncodeEntryType(Tools.GetStValueFromTSl(TSlEcr[FirstIndice], 'E_NATUREPIECE'));
    N1 := NodeDoc.Addchild('Journal');
    N1.Text := Tools.GetStValueFromTSl(TSlEcr[FirstIndice], 'E_JOURNAL');
    Setting := Root.AddChild('Setting');
    N1 := Setting.AddChild('AnalyticBehavior');
    N1.Text := 'Amount';
    N1 := Setting.AddChild('CurrencyBehavior');
    N1.Text := 'Known';
    N1 := Setting.AddChild('ValidationBehavior');
    N1.Text := 'Enabled';
    Root.Attributes['xmlns'] := 'http://schemas.datacontract.org/2004/07/Cegid.Finance.Services.WebPortal';
    Root.Attributes['xmlns:i'] := 'http://www.w3.org/2001/XMLSchema-instance';
    Result := UTF8Encode(Root.XML);
    XmlDoc.SaveToFile('C:\pgi01\XMLOUT\out.xml');
  finally
    XmlDoc := nil;
  end;
end;

{$IF not defined(APPSRV)}
function TSendEntryY2.SendEntryCEGID(WsEt: T_WSEntryType; TOBecr: TOB; DocType: string; DocNumber: integer): boolean;
var
  TSlEcr: TStringList;
begin
  { Transformation de la TOB en TStringList }
  TSlEcr := TStringList.Create;
  try
    Tools.TobToTStringList(TOBecr, TSlEcr);
    Result := SendEntryCEGID(WsEt, TSlEcr, DocType, DocNumber);
  finally
    FreeAndNil(TSlEcr);
  end;
end;
{$IFEND !APPSRV}

function TSendEntryY2.SendEntryCEGID(WsEt: T_WSEntryType; TSlEcr: TStringList; DocType: string; DocNumber: integer): boolean;
var
  OneConnectCEGID: TconnectCEGID;
  TheXml: WideString;
  TheRealNumDoc: Integer;
begin
  Result := False;
  OneConnectCEGID := TconnectCEGID.create;
  try
    SetCegidConnectParameters(OneConnectCEGID);
    if OneConnectCEGID.IsActive then
    begin
      TheXml := ConstitueEntries(TSlEcr);
      if (TheXml <> '') then
      begin
        OneConnectCEGID.AppelEntriesWS(DocType, DocNumber, TheXml, TheRealNumDoc);
        Result := (EnregistreInfoCptaY2(WsEt, TSlEcr[GetFirstIndiceEcr(TSlEcr)], TheRealNumDoc)); // Ecriture dans BTPECRITURE
      end;
    end;
  finally
    OneConnectCEGID.free;
  end;
end;

function TSendEntryY2.SendAccountingParameters(TSLFullPathFile : TStringList): boolean;
var
  MemFile      : TmemoryStream;
  CegidConnect : TconnectCEGID;
  http         : IWinHttpRequest;
  url          : string;
  sFile        : WideString;
  Cpt          : integer;
begin
  Result := False;
  if TSLFullPathFile.Count > 0 then
  begin
    CegidConnect := TconnectCEGID.create;
    try
      SetCegidConnectParameters(CegidConnect);
      if CegidConnect.IsActive then
      begin
        url  := Format('%s/%s/%s', [CegidConnect.GetStartUrl, CegidConnect.fDossier, WSCDS_EndUrlUploadBytes]);
        http := CoWinHttpRequest.Create;
        try
          { 1ère étape, upload du fichier }
          http.SetAutoLogonPolicy(0);
          http.Open('POST', url, False);
          http.SetRequestHeader('Content-Type', 'text/xml');
          http.SetRequestHeader('Accept', 'application/xml');
          try
            for Cpt := 0 to pred(TSLFullPathFile.Count) do
            begin
              MemFile := TmemoryStream.Create;
              try
                MemFile.LoadFromFile(TSLFullPathFile[Cpt]);
                SetString(sFile, PChar(MemFile.Memory), MemFile.Size div SizeOf(Char));
                http.Send(sFile);
              finally
                Result := (http.status = 200);
                MemFile.free;
              end;
            end;
          except
            on E: Exception do                                 
            begin
//              EnregistreResponse(http.ResponseText, NumDocOut);
              ShowMessage(E.Message);
              exit;
            end;
          end;
        finally
          http := nil;
        end;
        { 2ème étape, déclencher l'import }
        if Result then
        begin

        end;
      end;
    finally
      CegidConnect.Free;
    end;
  end;
end;

{ TconnectCEGID }
function TconnectCEGID.AppelEntriesWS(DocType: string; DocNumber: integer; TheXml: WideString; var NumDocOut: Integer): Boolean;

  {$IF not defined(APPSRV)}
  procedure EnregistreEVT(NumDocOut: Integer; MessageOut: widestring);
  var
    TobJnal: TOB;
    Nature: string;
    BlocNote: TStringList;
    QQ: TQuery;
    NumEvt: Integer;
  begin
    Nature := RechDom('GCNATUREPIECEG', DocType, False);
    BlocNote := TStringList.Create;
    try
      if NumDocOut <> 0 then
      begin
        BlocNote.Add(Nature + TraduireMemoire(' numéro ') + IntToStr(DocNumber));
        BlocNote.Add(TraduireMemoire(format('L''écriture comptable %d à été créé en comptabilité', [NumDocOut])));
      end
      else
      begin
        BlocNote.Add(Nature + TraduireMemoire(' numéro ') + IntToStr(DocNumber));
        BlocNote.Add('Annomalie lors du transfert');
        BlocNote.Add(TraduireMemoire('Message : ') + MessageOut);
      end;
      TobJnal := TOB.Create('JNALEVENT', nil, -1);
      try
        TobJnal.SetString('GEV_TYPEEVENT', 'WS');
        TobJnal.SetString('GEV_LIBELLE', 'Liaison WebApi Fiscalité');
        TobJnal.SetDateTime('GEV_DATEEVENT', Date);
        TobJnal.SetString('GEV_UTILISATEUR', V_PGI.User);
        if NumDocOut <> 0 then
          TobJnal.SetString('GEV_ETATEVENT', 'OK')
        else
          TobJnal.SetString('GEV_ETATEVENT', 'ERR');
        TobJnal.PutValue('GEV_BLOCNOTE', BlocNote.Text);
        QQ := OpenSQL('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT', True, -1, '', True);
        if not QQ.EOF then
          NumEvt := QQ.Fields[0].AsInteger
        else
          NumEvt := 0;
        Inc(NumEvt);
        Ferme(QQ);
        TobJnal.PutValue('GEV_NUMEVENT', NumEvt);
        TobJnal.InsertDB(nil);
      finally
        TobJnal.Free;
      end;
    finally
      BlocNote.Free;
    end;
  end;
  {$ELSE !APPSRV}

  procedure EnregistreEVT(NumDocOut: Integer; MessageOut: widestring);
  begin

  end;
  {$IFEND !APPSRV}

  procedure EnregistreResponse(HTTPResponse: Widestring; var NumDocOut: integer);
  var
    XmlDoc: IXMLDocument;
    NodeFolder: IXMLNode;
    II: Integer;
    JJ: Integer;
    MessageOut: string;
  begin
    NumDocOut := 0;
    XmlDoc := NewXMLDocument();
    try
      try
        XmlDoc.LoadFromXML(HTTPResponse);
      except
        {$IF not defined(APPSRV)}
        on E: Exception do
          PgiError('Erreur durant Chargement XML : ' + E.Message);
        {$IFEND !APPSRV}
      end;
      if not XmlDoc.IsEmptyDoc then
      begin
        MessageOut := '';
        for II := 0 to XmlDoc.DocumentElement.ChildNodes.Count - 1 do
        begin
          NodeFolder := XmlDoc.DocumentElement.ChildNodes[II]; // Liste des <Folder>
          case Tools.CaseFromString(NodeFolder.NodeName, ['DocumentNumber', 'Errors']) of
            {DocumentNumber}                         0:
              NumDocOut := StrToInt(NodeFolder.NodeValue);
            {Errors}                                 1:
              begin
                for JJ := 0 to NodeFolder.ChildNodes.Count - 1 do
                  MessageOut := MessageOut + Tools.iif(MessageOut <> '', '#13#10', '') + NodeFolder.ChildNodes[JJ].NodeValue;
                if MessageOut <> '' then
                  EnregistreEVT(NumDocOut, MessageOut);
              end;
          end;
        end;
      end;
    finally
      XmlDoc := nil;
    end;
  end;

var
  http: IWinHttpRequest;
  url: string;
begin
  Result := false;
  url := Format('%s/%s/%s', [GetStartUrl, fDossier, WSCDS_EndUrlEntries]);
  http := CoWinHttpRequest.Create;
  try
    http.SetAutoLogonPolicy(0); // Enable SSO
    http.Open('POST', url, False);
    http.SetRequestHeader('Content-Type', 'text/xml');
    http.SetRequestHeader('Accept', 'application/xml');
    try
      http.Send(TheXml);
    except
      on E: Exception do
      begin
        EnregistreResponse(http.ResponseText, NumDocOut);
        ShowMessage(E.Message);
        exit;
      end;
    end;
    if http.status = 200 then
    begin
      EnregistreResponse(http.ResponseText, NumDocOut);
      Result := (NumDocOut <> 0);
    end
    else
    begin
      EnregistreResponse(http.ResponseText, NumDocOut);
    end;
  finally
    http := nil;
  end;
end;

constructor TconnectCEGID.create;
begin
  factive := false;
  fServer := '';
  fport := 80;
  fDossier := '';
end;

destructor TconnectCEGID.destroy;
begin

  inherited;
end;

{$IF not defined(APPSRV)}
procedure TconnectCEGID.GetDossiers(var ListeDoss: TOB; var TheResponse: WideString);
var
  http: IWinHttpRequest;
  url: string;
begin
  if fServer = '' then
  begin
    PgiInfo('LE Serveur CEGID Y2 n''est pas défini');
    Exit;
  end;
  url := Format('%s/folders', [GetStartUrl]);
  http := CoWinHttpRequest.Create;
  try
    http.SetAutoLogonPolicy(0); // Enable SSO
    http.Open('GET', url, False);
    http.SetRequestHeader('Content-Type', 'text/xml');
    http.SetRequestHeader('Accept', 'application/xml,*/*');
    try
      http.Send(EmptyParam);
    except
      on E: Exception do
      begin
        ShowMessage(E.Message);
        exit;
      end;
    end;
    if http.status = 200 then
    begin
      TheResponse := http.ResponseText;
      RemplitTOBDossiers(ListeDoss, http.ResponseText);
    end;
  finally
    http := nil;
  end;
end;
{$IFEND !APPSRV}

{$IF not defined(APPSRV)}

procedure TconnectCEGID.GetExCpta(TOBexer: TOB);
var
  http: IWinHttpRequest;
  url: string;
  TheResponse: WideString;
begin
  url := Format('%s/%s/fiscalYears', [GetStartUrl, fDossier]);
  http := CoWinHttpRequest.Create;
  try
    http.SetAutoLogonPolicy(0); // Enable SSO
    http.Open('GET', url, False);
    http.SetRequestHeader('Content-Type', 'text/xml');
    http.SetRequestHeader('Accept', 'application/xml,*/*');
    http.Send(EmptyParam);
    if http.status = 200 then
    begin
      TheResponse := http.ResponseText;
      RemplitTOBExercices(TOBexer, http.ResponseText);
    end;
  finally
    http := nil;
  end;
end;
{$IFEND !APPSRV}

function TconnectCEGID.GetPort: string;
begin
  Result := IntToStr(fPort);
end;

{$IF not defined(APPSRV)}
procedure TconnectCEGID.RemplitTOBDossiers(ListeDoss: TOB; HTTPResponse: WideString);
var
  XmlDoc: IXMLDocument;
  NodeFolder, OneStep: IXMLNode;
  II, JJ: Integer;
  TOBL: TOB;
begin
  XmlDoc := NewXMLDocument();
  try
    try
      XmlDoc.LoadFromXML(HTTPResponse);
    except
      on E: Exception do
      begin
        PgiError('Erreur durant Chargement XML : ' + E.Message);
      end;
    end;
    if not XmlDoc.IsEmptyDoc then
    begin
      for II := 0 to XmlDoc.DocumentElement.ChildNodes.Count - 1 do
      begin
        NodeFolder := XmlDoc.DocumentElement.ChildNodes[II]; // Liste des <Folder>
        TOBL := TOB.Create('UN DOSSIER', ListeDoss, -1);
        for JJ := 0 to NodeFolder.ChildNodes.Count - 1 do
        begin
          OneStep := NodeFolder.ChildNodes.Nodes[JJ];
          TOBL.AddChampSupValeur(OneStep.NodeName, OneStep.NodeValue);
        end;
      end;
    end;
  finally
    XmlDoc := nil;
  end;
end;
{$IFEND !APPSRV}

{$IF not defined(APPSRV)}

procedure TconnectCEGID.RemplitTOBExercices(TOBexer: TOB; HTTPResponse: WideString);
var
  XmlDoc: IXMLDocument;
  NodeFolder, OneStep: IXMLNode;
  II, JJ: Integer;
  TOBL: TOB;
begin
  XmlDoc := NewXMLDocument();
  try
    try
      XmlDoc.LoadFromXML(HTTPResponse);
    except
      on E: Exception do
      begin
        PgiError('Erreur durant Chargement XML : ' + E.Message);
      end;
    end;
    if not XmlDoc.IsEmptyDoc then
    begin
      for II := 0 to XmlDoc.DocumentElement.ChildNodes.Count - 1 do
      begin
        NodeFolder := XmlDoc.DocumentElement.ChildNodes[II]; // Liste des <Folder>
        TOBL := TOB.Create('UN EXERCICE', TOBexer, -1);
        for JJ := 0 to NodeFolder.ChildNodes.Count - 1 do
        begin
          OneStep := NodeFolder.ChildNodes.Nodes[JJ];
          TOBL.AddChampSupValeur(OneStep.NodeName, OneStep.NodeValue);
        end;
      end;
    end;
  finally
    XmlDoc := nil;
  end;
end;
{$IFEND !APPSRV}

function TconnectCEGID.GetStartUrl: string;
begin
  Result := Format('http://%s:%d/CegidFinanceWebApi/api/v1', [fServer, fport]);
end;

procedure TconnectCEGID.SetDossier(const Value: string);
begin
  fDossier := Value;
  factive := (fServer <> '') and (fDossier <> '');
end;

procedure TconnectCEGID.SetPort(const Value: string);
begin
  if Tools.IsNumeric_(Value) then
    fport := strtoint(Value);
  if fport = 0 then
    fPort := 80;
  factive := (fServer <> '') and (fDossier <> '');
end;

procedure TconnectCEGID.SetServer(const Value: string);
begin
  fServer := Value;
  factive := (fServer <> '') and (fDossier <> '');
end;

{ TGetParamWSCEGID }
class function TGetParamWSCEGID.GetCodeFromWsEt(WsEt: T_WSEntryType): string;
begin
  case WsEt of
    wsetDocument:
      Result := 'DOC'; // Ecriture de pièce
    wsetPayment:
      Result := 'RGT'; // Ecriture de règlement
    wsetPayer:
      Result := 'PAY'; // Ecriture de tiers payeur
    wsetExtourne:
      Result := 'EXT'; // Ecriture d'extourne
    wsetSubContractPayment:
      Result := 'SCP'; // Ecriture de règlement de sous-traitance
    wsetStock:
      Result := 'STK'; // Ecriture de stock
  else
    Result := '';
  end;
end;

class function TGetParamWSCEGID.ConnectToY2: Boolean;
begin
  {$IF not defined(APPSRV)}
  Result := (GetPSoc(wspsFolder) <> '');
  {$ELSE !APPSRV}
  Result := True;
  {$IFEND !APPSRV}
end;

class function TGetParamWSCEGID.GetPSoc(PSocType: T_WSPSocType): string;
begin
  {$IF not defined(APPSRV)}
  case PSocType of
    wspsServer:
      Result := GetParamSocSecur(WSCDS_SocServer, '');
    wspsPort:
      Result := GetParamSocSecur(WSCDS_SocNumPort, '');
    wspsFolder:
      Result := GetParamSocSecur(WSCDS_SocCegidDos, '');
    wspsLastSynchro:
      Result := GetParamSocSecur(WSCDS_SocLastSync, '31/12/2099 23:59:59');
  else
    Result := '';
  end;
  {$ELSE !APPSRV}
  Result := '';
  {$IFEND !APPSRV}
end;

end.

