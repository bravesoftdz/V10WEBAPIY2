unit uSendEntryY2;

interface

uses
  Windows
  , Messages
  , SysUtils
  , Variants
  , Classes
  , Graphics
  , Controls
  , Forms
  , Dialogs
  , StdCtrls
  , ComCtrls
  , WinHttp_TLB
  , XMLDoc
  , xmlintf
  , DateUtils
  , UConnectWSConst
  {$IF not defined(APPSRV)}
  , UTOB
  , HMsgBox
  , ParamSoc
  {$IFEND !APPSRV}
  ;

type
  TSendEntryY2 = class (Tobject)
  private
    class function EncodeDocType (TypePiece : string) : string;
    class function EncodeEntryType (EntryType : string) : string;
    class function EnregistreInfoCptaY2(WsEt : T_WSEntryType; TslLine : string; TheNumDoc : Integer) : boolean; //class function EnregistreInfoCptaY2(WsEt : T_WSEntryType; TOBECr : TOB; TheNumDoc : Integer) : boolean;s
    class function EncodeAmount(TslLine : String) : string; //class function EncodeAmount (TobE : TOB) : string;
    class function EncodeDate(TslLine, DateFieldName : string) : string; //class function EncodeDueDate(TOBE : TOB) : string;
    class function ConstitueEntries(TSlEcr : TStringList) : WideString; //class function ConstitueEntries(TOBecr : TOB) : WideString;
    class function GetStValueFromTSl(TSlLine, FieldName : string) : string;
    {$IF not defined(APPSRV)}
    class function EncodeSens(TOBE : TOB) : string;
    class function EncodeAxis(TOBAN : TOB) : string;
  	{$IFEND !APPSRV}

  public
//    class procedure RecupParamCptaFromWS;
    {$IF not defined(APPSRV)}
    class procedure SendEntryCEGID(WsEt : T_WSEntryType; TOBecr : TOB; DocType : string; DocNumber : integer; Var SendCegid : boolean); overload;
  	{$IFEND !APPSRV}
    class procedure SendEntryCEGID(WsEt : T_WSEntryType; TSlEcr : TStringList; DocType : string; DocNumber : integer; Var SendCegid : boolean); overload;
  end;

implementation

uses
  db
  , uLkJSON
  , CommonTools
  {$IFNDEF DBXPRESS}
  , dbtables
  {$ELSE DBXPRESS}
  , uDbxDataSet
  {$ENDIF DBXPRESS}
  {$IF not defined(APPSRV)}
  , Hctrls
  , Aglinit
  , Hent1
  , wCommuns
  , UtilPGI
  {$IFEND !APPSRV}
  ;

function StringToStream(const AString: string) : Tstream;
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

function DateTime2Tdate (TheDateTime : TdateTime ) : String;
var
  YY,MM,DD,Hours,Mins,secs,milli : Word;
begin
  DecodeDateTime(TheDateTime,YY,MM,DD,Hours,Mins,secs,Milli);
  Result := Format ('%4d-%0.2d-%0.2dT%0.2d:%0.2d:%0.2d',[YY,MM,DD,Hours,Mins,secs]);
end;

function TDate2DateTime (OneTdate : String ) : TDateTime;
var
  TheTDate : string;
  YYYY,MM,DD : string;
  PDATE : string;
  TheTime : string;
  IposT,II : Integer;
begin
  {$IF not defined(APPSRV)}
  Result := iDate1900;
  {$ELSE !APPSRV}
  Result := 2;
  {$IFEND !APPSRV}
  IposT := Pos('T',OneTdate);
  if IposT>0 then
  begin
    II := 0;
    TheTDate := Copy(OneTdate,1,IPosT-1);
    TheTime := Copy(OneTdate,IposT+1,Length(OneTdate)-1);
    repeat
      PDATE := Tools.ReadTokenSt_(TheTDate, '-');
      if PDATE <> '' then
      begin
        if II = 0 then
        begin
          YYYY := PDATE; Inc(II);
        end else if II = 1 then
        begin
          MM := PDATE; Inc(II);
        end else
        begin
          DD := PDATE; Inc(II);
        end;
      end;
    until PDATE='';
    if (YYYY <> '') and (MM <> '') and (DD <> '') then
    begin
      Result := StrToDateTime(DD+'/'+MM+'/'+YYYY+' '+TheTime);
    end;
  end;
end;

class function TSendEntryY2.EncodeDocType (TypePiece : string) : string;
begin
  case Tools.CaseFromString(TypePiece, ['N', 'S']) of
    {N} 0 : Result := 'Normal';
    {S} 1 : Result := 'Simulation';
  end;
end;

class function TSendEntryY2.EncodeEntryType (EntryType : string) : string;
begin
  case Tools.CaseFromString(EntryType, ['AC', 'AF', 'ECC', 'FC', 'FF', 'OC', 'OD', 'OF', 'RC', 'RF']) of
    {AC } 0 : Result := 'CustomerCredit';
    {AF } 1 : Result := 'ProviderCredit';
    {ECC} 2 : Result := 'ExchangeDifference';
    {FC } 3 : Result := 'CustomerInvoice';
    {FF } 4 : Result := 'ProviderInvoice';
    {OC } 5 : Result := 'CustomerDeposite';
    {OD } 6 : Result := 'SpecificOperation';
    {OF } 7 : Result := 'ProviderDeposite';
    {RC } 8 : Result := 'CustomerPayment';
    {RF } 9 : Result := 'ProviderPayment';
  end;
end;

{$IF not defined(APPSRV)}
class function TSendEntryY2.EncodeSens(TOBE : TOB) : string;
begin
  Result := iif(TOBE.GetDouble('E_DEBITDEV') <> 0, 'Debit', 'Credit');
end;
{$IFEND !APPSRV}

class function TSendEntryY2.EnregistreInfoCptaY2(WsEt : T_WSEntryType; TslLine : string; TheNumDoc : Integer) : boolean;
var
  eJournal       : string;
  eExercice      : string;
  eDateComptable : TDateTime;
  eEntity        : integer;
  eNumeroPiece   : integer;
  Exist          : boolean;
begin
  eEntity        := StrToInt(GetStValueFromTSl(TslLine, 'E_ENTITY'));
  eJournal       := GetStValueFromTSl(TslLine, 'E_JOURNAL');
  eExercice      := GetStValueFromTSl(TslLine, 'E_EXERCICE');
  eDateComptable := StrToDateTime(GetStValueFromTSl(TslLine, 'E_DATECOMPTABLE'));
  eNumeroPiece   := StrToInt(GetStValueFromTSl(TslLine, 'E_NUMEROPIECE'));
  if eJournal <> '' then
  begin
    
  end else
    Result := False;

(*
  if TOBECr.Detail.Count > 0 then
  begin
    TobEcrL := TOBECr.detail[0];
    TobLBTP := TOB.Create('BTPECRITURE',nil,-1);
    try
      TobLBTP.SetInteger('BE0_ENTITY'        , TobEcrL.GetInteger('E_ENTITY'));
      TobLBTP.SetString('BE0_JOURNAL'        , TobEcrL.GetString('E_JOURNAL'));
      TobLBTP.SetString('BE0_EXERCICE'       , TobEcrL.GetString('E_EXERCICE'));
      TobLBTP.SetDateTime('BE0_DATECOMPTABLE', TobEcrL.GetDateTime('E_DATECOMPTABLE'));
      TobLBTP.SetInteger('BE0_NUMEROPIECE'   , TobEcrL.GetInteger('E_NUMEROPIECE'));
      TOBLBTP.SetString('BE0_REFERENCEY2'    , IntToStr(TheNumDoc));
      TobLBTP.SetString('BE0_TYPE'           , TGetParamWSCEGID.GetCodeFromWsEt(WsEt));
      Result := (TobLBTP.InsertOrUpdateDB);
    finally
      FreeAndNil(TobLBTP);
    end;
  end else
    Result := False;
*)
end;

class function TSendEntryY2.EncodeAmount(TslLine : String) : string;
var
  DebAmount  : Double;
  CredAmount : double;
  TheMontant : double;
begin
  DebAmount  := StrToFloat(GetStValueFromTSl(TslLine, 'E_DEBITDEV'));
  CredAmount := StrToFloat(GetStValueFromTSl(TslLine, 'E_CREDITDEV'));
  Result := Tools.StrFPoint_(Tools.iif(DebAmount <> 0, DebAmount, CredAmount));
//  TheMontant := iif(TOBE.GetDouble('E_DEBITDEV') <> 0, TOBE.GetDouble('E_DEBITDEV'), TOBE.GetDouble('E_CREDITDEV'));
//  Result := STRFPOINT(TheMontant);
end;

class function TSendEntryY2.EncodeDate(TslLine, DateFieldName : string) : string;
var
//  TheDate : Tdatetime;
  DateValue : TDateTime;
begin
  DateValue := StrToDateTime(GetStValueFromTSl(TslLine, DateFieldName));
  Result  := DateTime2Tdate(Tools.iif((DateValue < 2), 2, DateValue));
//  TheDate := iif(TOBE.GetDateTime('E_DATEECHEANCE') < IDate1900, IDate1900, TOBE.GetDateTime('E_DATEECHEANCE'));
//  Result  := DateTime2Tdate(TheDate);
end;

{$IF not defined(APPSRV)}
class function TSendEntryY2.EncodeAxis(TOBAN : TOB) : string;
begin
  case Tools.CaseFromString(TOBAN.GetString('Y_AXE'), ['A1', 'A2', 'A3', 'A4', 'A5']) of
    {A1} 0 : Result := 'One';
    {A2} 1 : Result := 'Two';
    {A3} 2 : Result := 'Three';
    {A4} 3 : Result := 'Four';
    {A5} 4 : Result := 'Five';
  end;
end;
{$IFEND !APPSRV}

//class function TSendEntryY2.ConstitueEntries(TOBecr : TOB) : WideString;
class function TSendEntryY2.ConstitueEntries(TSlEcr : TStringList) : WideString;
var
  XmlDoc      : IXMLDocument;
  Root        : IXMLNode;
  NodeDoc     : IXMLNode;
  Entries     : IXMLNode;
  Entry       : IXMLNode;
  Amounts     : IXMLNode;
  EntryAmount : IXMLNode;
  Analytics   : IXMLNode;
  Analytic    : IXMLNode;
  Sections    : IXMLNode;
  Setting     : IXMLNode;
  N1          : IXMLNode;
  Cpt         : Integer;
//  TOBE : TOB;

(*
  procedure AddAnalytics(ParentNode : IXMLNode);
  var
    CptTobE  : integer;
    CptTobAn : integer;
    TobAn    : TOB;
  begin
    Analytics := Amounts.AddChild('Analytics');    // <- Noeud Analytics> ->
    if TOBE.detail.count > 0 then
    begin
      for CptTobE := 0 to TOBE.Detail.count -1 do
      begin
        TobAn := TOBE.detail[CptTobE];
        for CptTobAn := 0 to pred(TobAn.detail.count) do
        begin
          Analytic := Analytics.AddChild('EntryAnalytic');    // <- Noeud EntryAnalytic>> ->
          N1 := Analytic.AddChild('Amount');  N1.Text := StrfPoint(TobAn.detail[CptTobAn].GetDouble('Y_DEBIT') + TobAn.detail[CptTobAn].GetDouble('Y_CREDIT'));
          N1 := Analytic.AddChild('Axis');    N1.Text := EncodeAxis(TobAn.detail[CptTobAn]);
          N1 := Analytic.AddChild('Percent'); N1.Text := '0';
          Sections := Analytic.AddChild('Sections');    // <- Noeud EntryAnalytic>> ->
          Sections.Attributes['xmlns:a'] := 'http://schemas.microsoft.com/2003/10/Serialization/Arrays';
          N1 := Sections.AddChild('a:string'); N1.Text := TobAn.detail[CptTobAn].GetString('Y_SECTION');
        end;
      end;
    end else
      Analytics.Attributes ['i:nil'] := 'true';
  end;
*)

begin
  Result         := '';
  XmlDoc         := NewXMLDocument();
  XmlDoc.Options := [doNodeAutoIndent];
  try
    Root := XmlDoc.AddChild('EntryParameter');
    NodeDoc := Root.AddChild('Document');    // <- Noeud Document ->
    // --- Sur le <document>
    N1 := NodeDoc.Addchild ('AccountingDate'); N1.Text := DateTime2Tdate(StrToDateTime(GetStValueFromTSl(TSlEcr[0], 'E_DATECOMPTABLE')));  //DateTime2Tdate(TOBecr.detail[0].GetDateTime('E_DATECOMPTABLE'));
    N1 := NodeDoc.Addchild ('BusinessCenter'); N1.Text := GetStValueFromTSl(TSlEcr[0], 'E_ETABLISSEMENT');  //TOBecr.detail[0].GetString('E_ETABLISSEMENT');
    N1 := NodeDoc.Addchild ('Currency');       N1.Text := GetStValueFromTSl(TSlEcr[0], 'E_DEVISE');  //TOBecr.detail[0].GetString('E_DEVISE');
    N1 := NodeDoc.Addchild ('CurrencyRate');   N1.Text := Tools.StrFPoint_(StrToFloat(GetStValueFromTSl(TSlEcr[0], 'E_TAUXDEV')));  //STRFPOINT(TOBecr.detail[0].Getdouble('E_TAUXDEV'));
    N1 := NodeDoc.Addchild ('DocumentType');   N1.Text := EncodeDocType(GetStValueFromTSl(TSlEcr[0], 'E_DEVISE'));  //EncodeDocType(TOBecr.detail[0].GetString('E_QUALIFPIECE'));
    Entries := NodeDoc.Addchild('Entries'); // <- Noeud Entries ->
    for Cpt := 0 to pred(TSlEcr.Count) do
    begin
      Entry := Entries.Addchild('Entry');   // <- Noeud Entry ->
      N1          := Entry.AddChild('AmountDirection'); N1.Text := Tools.iif(StrToFloat(GetStValueFromTSl(TSlEcr[Cpt], 'E_DEBITDEV')) <> 0, 'Debit', 'Credit');   //EncodeSens(TOBE);
      Amounts     := Entry.AddChild('Amounts');    // <- Noeud Amounts ->
      EntryAmount := Amounts.AddChild('EntryAmount');    // <- Noeud EntryAmount ->
      N1          := EntryAmount.AddChild('Amount');      N1.Text := EncodeAmount(TSlEcr[Cpt]);
      N1          := EntryAmount.AddChild('DueDate');     N1.Text := EncodeDate(TSlEcr[Cpt], 'E_DATEECHEANCE');
      N1          := EntryAmount.AddChild('Iban');        N1.Text := GetStValueFromTSl(TSlEcr[Cpt], 'E_RIB');  //if TOBE.GetString('E_RIB') <> '' then N1.Text := TOBE.GetString('E_RIB');
      N1          := EntryAmount.AddChild('PaymentMode'); N1.Text := GetStValueFromTSl(TSlEcr[Cpt], 'E_MODEPAIE');  //iTOBE.GetString('E_MODEPAIE');
      if GetStValueFromTSl(TSlEcr[Cpt], 'E_MODEPAIE') <> '' then N1 := EntryAmount.AddChild('SepaCreditorIdentifier');
      N1.Attributes ['i:nil'] := 'true';
      N1 := EntryAmount.AddChild('UniqueMandateReference');

//      AddAnalytics(Entry);

      N1 := Entry.AddChild('Description');           N1.Text := GetStValueFromTSl(TSlEcr[Cpt], 'E_LIBELLE');  //TOBE.GetString('E_LIBELLE');
      N1 := Entry.AddChild('ExternalDateReference'); N1.Text := EncodeDate(TSlEcr[Cpt], 'E_DATEREFEXTERNE');// EncodeExternalDateReference(TOBE); 
      N1 := Entry.AddChild('ExternalReference');     N1.Text := GetStValueFromTSl(TSlEcr[Cpt], 'E_REFEXTERNE');  //if TOBE.GetString('E_REFEXTERNE') <> '' then N1.Text := TOBE.GetString('E_REFEXTERNE');
      N1 := Entry.AddChild('GeneralAccount');        N1.Text := GetStValueFromTSl(TSlEcr[Cpt], 'E_GENERAL');  //TOBE.GetString('E_GENERAL');
      N1 := Entry.AddChild('InternalReference');     N1.Text := GetStValueFromTSl(TSlEcr[Cpt], 'E_REFINTERNE');  //if TOBE.GetString('E_REFINTERNE') <> '' then N1.Text := TOBE.GetString('E_REFINTERNE');
      N1 := Entry.AddChild('SubsidiaryAccount');     N1.Text := GetStValueFromTSl(TSlEcr[Cpt], 'E_AUXILIAIRE');  //if TOBE.GetString('E_AUXILIAIRE') <> '' then N1.Text := TOBE.GetString('E_AUXILIAIRE');
    end;
    N1 := NodeDoc.Addchild ('EntryType'); N1.Text := EncodeEntryType(GetStValueFromTSl(TSlEcr[Cpt], 'E_NATUREPIECE'));  //EncodeEntryType(TOBecr.detail[0].GetString('E_NATUREPIECE'));
    N1 := NodeDoc.Addchild ('Journal');   N1.Text := GetStValueFromTSl(TSlEcr[Cpt], 'E_JOURNAL');  //TOBecr.detail[0].GetString('E_JOURNAL');
    Setting := Root.AddChild('Setting');
    N1 := Setting.AddChild('AnalyticBehavior');   N1.Text := 'Amount';
    N1 := Setting.AddChild('CurrencyBehavior');   N1.Text := 'Known';
    N1 := Setting.AddChild('ValidationBehavior'); N1.Text := 'Enabled';
    Root.Attributes['xmlns']  := 'http://schemas.datacontract.org/2004/07/Cegid.Finance.Services.WebPortal';
    Root.Attributes['xmlns:i'] := 'http://www.w3.org/2001/XMLSchema-instance';
    Result := UTF8Encode(Root.XML);
    XmlDoc.SaveToFile('C:\pgi01\XMLOUT\out.xml');
  finally
    XmlDoc:= nil;
  end;
end;

class function TSendEntryY2.GetStValueFromTSl(TSlLine, FieldName : string) : string;
var
  Values : string;
  Value  : string;
  FindIt : boolean;
begin
  if (TSlLine <> '') and (FieldName <> '') then
  begin
    FindIt := False;
    Values := TSlLine;
    while (Values <> '') or (not FindIt) do
    begin
      Value := Tools.ReadTokenSt_(Values, '^');
      if Copy(Value, 1, Pos('=', Value)) = FieldName then
      begin
        Result := Copy(Value, pos('=', Value) + 1, Length(Value));
        FindIt := True;
      end;
    end;
  end else
    Result := '';
end;



{$IF not defined(APPSRV)}
class procedure TSendEntryY2.SendEntryCEGID(WsEt : T_WSEntryType; TOBecr : TOB; DocType : string; DocNumber : integer; Var SendCegid : boolean);
var
  TSlEcr : TStringList;
begin
  { Transformation de la TOB en TStringList }
  TSlEcr := TStringList.Create;
  try
    Tools.TobToTStringList(Tobecr, TSlEcr);
    SendEntryCEGID(WsEt, TSlEcr, DocType, DocNumber, SendCegid);
  finally
    FreeAndNil(TSlEcr);
  end;
end;
{$IFEND !APPSRV}

class procedure TSendEntryY2.SendEntryCEGID(WsEt : T_WSEntryType; TSlEcr : TStringList; DocType : string; DocNumber : integer; Var SendCegid : boolean);
var
  OneConnectCEGID : TconnectCEGID;
  TheXml          : WideString;
  TheRealNumDoc   : Integer;
begin
  SendCegid := False;
  OneConnectCEGID := TconnectCEGID.create;
  try
    OneConnectCEGID.CEGIDServer := TGetParamWSCEGID.GetPSoc(wspsServer);
    OneConnectCEGID.CEGIDPORT   := TGetParamWSCEGID.GetPSoc(wspsPort);
    OneConnectCEGID.DOSSIER     := TGetParamWSCEGID.GetPSoc(wspsFolder);
    if OneConnectCEGID.IsActive then
    begin
      TheXml := ConstitueEntries(TSlEcr);
      if (TheXml <> '') then
      begin
       OneConnectCEGID.AppelEntriesWS(DocType, DocNumber, TheXml, TheRealNumDoc);
       //SendCegid := (EnregistreInfoCptaY2(WsEt, TOBecr, TheRealNumDoc)); // Ecriture dans BTPECRITURE
       SendCegid := (EnregistreInfoCptaY2(WsEt, TSlEcr[0], TheRealNumDoc)); // Ecriture dans BTPECRITURE
      end;
    end;
  finally
    OneConnectCEGID.free;
  end;
end;

(*
class procedure TSendEntryY2.RecupParamCptaFromWS;

  procedure Majexercices (TOBEx : TOB);
  var
    II : Integer;
    TOBE,TOBEXER,TOBEE : TOB;
  begin
    TOBEXER := TOB.Create('LES EXERCICES',nil,-1);
    TRY
      BEGINTRANS;
      TRY
        ExecuteSQL('DELETE FROM EXERCICE');
        for II := 0 to TOBEx.detail.count -1 do
        begin
          TOBE := TOBEx.detail[II];
          TOBEE := TOB.create('EXERCICE', TOBEXER, -1);
          TOBEE.SetString('EX_EXERCICE'       , TOBE.GetString('Id'));
          TOBEE.SetString('EX_LIBELLE'        , TOBE.GetString('Description'));
          TOBEE.SetString('EX_ABREGE'         , TOBE.GetString('ShortName'));
          TOBEE.SetDateTime('EX_DATEDEBUT'    , TDate2DateTime(TOBE.GetString('BeginDate')));
          TOBEE.SetDateTime('EX_DATEFIN'      , TDate2DateTime(TOBE.GetString('EndDate')));
          TOBEE.SetString('EX_ETATCPTA'       , iif(TOBE.GetString('State')='FinalClosing', 'CDE', 'OUV'));
          TOBEE.SetString('EX_ETATBUDGET'     , 'OUV');
          TOBEE.SetString('EX_SOCIETE'        , V_PGI.CodeSociete);
          TOBEE.SetString('EX_VALIDEE'        , '------------------------');
          TOBEE.SetDateTime('EX_DATECUM'      , Idate1900);
          TOBEE.SetDateTime('EX_DATECUMRUB'   , Idate1900);
          TOBEE.SetDateTime('EX_DATECUMBUD'   , Idate1900);
          TOBEE.SetDateTime('EX_DATECUMBUDGET', Idate1900);
          TOBEE.SetInteger('EX_ENTITY'        , 0);
        end;
        if TOBEXER.DETAIL.Count > 0 then
        begin
          TOBEXER.InsertDB(nil);
        end;
        COMMITTRANS;
      EXCEPT
        ROLLBACK;
      END;
    FINALLY
      TOBEXER.Free;
    END;
  end;

  procedure GetInfoFromCegid (OneConnectCEGID : TConnectCEGID;  DateDay : Tdatetime);
  var
    TOBEx : TOB;
  begin
    TOBEx := TOB.Create ('LES EX CPTA',nil,-1);
    TRY
      TRY
        ExecuteSql ('INSERT INTO BTBLOCAGE '+
                    '(BTB_GUID, BTB_TYPE, BTB_IDDOC, BTB_USER, BTB_DATECREATION, BTB_HEURECREATION) '+
                    'VALUES '+
                    '("RECUPINFOCPTAWS","","","'+V_PGI.User+'","'+UsDateTime(DateDay) + '","' + USDateTime(DateDay) + '")');
        OneConnectCEGID.GetExCpta (TOBEx);
        Majexercices (TOBEx);
        ExecuteSQL('DELETE FROM BTBLOCAGE WHERE BTB_GUID="RECUPINFOCPTAWS"');
        SetParamSoc(WSCDS_SocLastSync, DateTimeToStr(DateDay));
      EXCEPT
      end;
    FINALLY
      TOBEx.Free;
    end;
  end;

var OneConnectCEGID : TconnectCEGID;
    finTrait : Boolean;
    NbTry : Integer;
    DateDay,LastSync : TDateTime;
begin
  finTrait := false;
  NbTry := 1;
  DateDay := Date;
  LastSync := StrToDate(DateToStr(StrToDateTime(TGetParamWSCEGID.GetPSoc(wspsLastSynchro))));
  OneConnectCEGID := TconnectCEGID.create;
  TRY
    OneConnectCEGID.CEGIDServer := TGetParamWSCEGID.GetPSoc(wspsServer);
    OneConnectCEGID.CEGIDPORT := TGetParamWSCEGID.GetPSoc(wspsPort);
    OneConnectCEGID.DOSSIER := TGetParamWSCEGID.GetPSoc(wspsFolder);
    if OneConnectCEGID.IsActive then
    begin
      repeat
        TRY
          if DateDay > LastSync then
          begin
            GetInfoFromCegid (OneConnectCEGID,DateDay);
            finTrait := True;
          end else
          begin
            fintrait := true;
          end;
        except
          Sleep(1000); inc(NbTry);
        end;
      until (finTrait) or (NbTry > 30);
    end else
    begin
      ;
    end;
  FINALLY
    OneConnectCEGID.Free;
  end;
end;
*)

{ TconnectCEGID }
function TconnectCEGID.AppelEntriesWS(DocType : string; DocNumber : integer;TheXml : WideString; var NumDocOut : Integer) : Boolean;

  {$IF not defined(APPSRV)}
  procedure EnregistreEVT (NumDocOut : Integer ; MessageOut : widestring);
  var
    TobJnal : TOB;
    Nature : string;
    BlocNote : TStringList;
    QQ : TQuery;
    NumEvt : Integer;
  begin
    Nature := RechDom('GCNATUREPIECEG', DocType, False);
    BlocNote := TStringList.Create;
    try
      if NumDocOut <> 0 then
      begin
        BlocNote.Add(Nature + TraduireMemoire(' numéro ') + IntToStr(DocNumber));
        BlocNote.Add(TraduireMemoire( format('L''écriture comptable %d à été créé en comptabilité',[NumDocOut])));
      end else
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
        if NumDocOut <> 0 then TobJnal.SetString('GEV_ETATEVENT', 'OK')
                          else TobJnal.SetString('GEV_ETATEVENT', 'ERR');
        TobJnal.PutValue('GEV_BLOCNOTE', BlocNote.Text);
        QQ := OpenSQL('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT', True,-1, '', True);
        if not QQ.EOF then
          NumEvt := QQ.Fields[0].AsInteger
        else
          NumEvt := 0;
        Inc(NumEvt);
        Ferme(QQ);
        TOBJnal.PutValue('GEV_NUMEVENT', NumEvt);
        TobJnal.InsertDB(nil);
      finally
        TobJnal.Free;
      end;
    finally
      BlocNote.Free;
    end;
  end;
  {$ELSE !APPSRV}
  procedure EnregistreEVT (NumDocOut : Integer ; MessageOut : widestring);
  begin
  
  end;
  {$IFEND !APPSRV}
  
  procedure  EnregistreResponse (HTTPResponse: Widestring; var NumDocOut : integer);
  var
    XmlDoc     : IXMLDocument ;
    NodeFolder : IXMLNode;
    II         : Integer;
    JJ         : Integer;
    MessageOut : string;
  begin
    NumDocOut := 0;
    XmlDoc := NewXMLDocument();
    TRY
      TRY
        XmlDoc.LoadFromXML(HTTPResponse);
      EXCEPT
        {$IF not defined(APPSRV)}
        On E: Exception do
          PgiError('Erreur durant Chargement XML : ' + E.Message );
        {$IFEND !APPSRV}
      end;
      if not XmlDoc.IsEmptyDoc then
      begin
        MessageOut := '';
        For II := 0 to Xmldoc.DocumentElement.ChildNodes.Count -1 do
        begin
          NodeFolder := XmlDoc.DocumentElement.ChildNodes[II]; // Liste des <Folder>
          case Tools.CaseFromString(NodeFolder.NodeName, ['DocumentNumber', 'Errors']) of
            {DocumentNumber} 0 : NumDocOut := StrToInt(NodeFolder.NodeValue);
            {Errors}         1 : begin
                                   for JJ := 0 to NodeFolder.ChildNodes.Count -1 do
                                     MessageOut := MessageOut +  Tools.iif(MessageOut <> '', '#13#10', '') + NodeFolder.ChildNodes [JJ].NodeValue;
                                   if MessageOut <> '' then
                                     EnregistreEVT (NumDocOut,MessageOut);
                                 end;
          end;
        end;
      end;
    FINALLY
      XmlDoc:= nil;
    end;
  end;

var
  http: IWinHttpRequest;
  url : string;

begin
  Result := false;
  url := Format('%s/%s/entries',[GetStartUrl, fDossier]);
  http := CoWinHttpRequest.Create;
  try
    http.SetAutoLogonPolicy(0); // Enable SSO
    http.Open('POST', url, False);
    http.SetRequestHeader('Content-Type', 'text/xml');
    http.SetRequestHeader('Accept', 'application/xml');
    TRY
      http.Send(TheXml);
    EXCEPT
      on E: Exception do
      begin
        EnregistreResponse (http.ResponseText,NumDocOut);
        ShowMessage(E.Message);
        exit;
      end;
    end;
    if http.status = 200 then
    begin
      EnregistreResponse (http.ResponseText,NumDocOut);
      Result := (NumDocOut <> 0);
    end else
    begin
      EnregistreResponse (http.ResponseText,NumDocOut);
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
procedure TconnectCEGID.GetDossiers(var ListeDoss: TOB; var TheResponse : WideString);
var
  http: IWinHttpRequest;
  url : string;
begin
  if fServer = '' then
  begin
    PgiInfo('LE Serveur CEGID Y2 n''est pas défini');
    Exit;
  end;
  url := Format('%s/folders',[GetStartUrl]);
  http := CoWinHttpRequest.Create;
  try
    http.SetAutoLogonPolicy(0); // Enable SSO
    http.Open('GET', url, False);
    http.SetRequestHeader('Content-Type', 'text/xml');
    http.SetRequestHeader('Accept', 'application/xml,*/*');
    TRY
      http.Send(EmptyParam);
    EXCEPT
      on E: Exception do
      begin
        ShowMessage(E.Message);
        exit;
      end;
    END;
    if http.status = 200 then
    begin
      TheResponse := http.ResponseText;
      RemplitTOBDossiers(ListeDoss,http.ResponseText);
    end;
  finally
    http := nil;
  end;
end;
{$IFEND !APPSRV}

{$IF not defined(APPSRV)}
procedure TconnectCEGID.GetExCpta (TOBexer : TOB);
var
  http: IWinHttpRequest;
  url : string;
  TheResponse : WideString;
begin
  url := Format('%s/%s/fiscalYears',[GetStartUrl, fDossier]);
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
      RemplitTOBExercices(TOBexer,http.ResponseText);
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
var XmlDoc : IXMLDocument ;
    NodeFolder,OneStep : IXMLNode;
    II,JJ : Integer;
    TOBL : TOB;
begin
  XmlDoc := NewXMLDocument();
  TRY
    TRY
      XmlDoc.LoadFromXML(HTTPResponse);
    EXCEPT
      On E: Exception do
      begin
        PgiError('Erreur durant Chargement XML : ' + E.Message );
      end;
    end;
    if not XmlDoc.IsEmptyDoc then
    begin
      For II := 0 to Xmldoc.DocumentElement.ChildNodes.Count -1 do
      begin
        NodeFolder := XmlDoc.DocumentElement.ChildNodes[II]; // Liste des <Folder>
        TOBL := TOB.Create('UN DOSSIER',ListeDoss,-1);
        for JJ := 0 to NodeFolder.ChildNodes.Count -1 do
        begin
          OneStep := NodeFolder.ChildNodes.Nodes[JJ];
          TOBL.AddChampSupValeur(OneStep.NodeName,OneStep.NodeValue);
        end;
      END;
    end;
  FINALLY
  	XmlDoc:= nil;
  end;
end;
{$IFEND !APPSRV}

{$IF not defined(APPSRV)}
procedure TconnectCEGID.RemplitTOBExercices(TOBexer: TOB; HTTPResponse: WideString);
var XmlDoc : IXMLDocument ;
    NodeFolder,OneStep : IXMLNode;
    II,JJ : Integer;
    TOBL : TOB;
begin
  XmlDoc := NewXMLDocument();
  TRY
    TRY
      XmlDoc.LoadFromXML(HTTPResponse);
    EXCEPT
      On E: Exception do
      begin
        PgiError('Erreur durant Chargement XML : ' + E.Message );
      end;
    end;
    if not XmlDoc.IsEmptyDoc then
    begin
      For II := 0 to Xmldoc.DocumentElement.ChildNodes.Count -1 do
      begin
        NodeFolder := XmlDoc.DocumentElement.ChildNodes[II]; // Liste des <Folder>
        TOBL := TOB.Create('UN EXERCICE',TOBexer,-1);
        for JJ := 0 to NodeFolder.ChildNodes.Count -1 do
        begin
          OneStep := NodeFolder.ChildNodes.Nodes[JJ];
          TOBL.AddChampSupValeur(OneStep.NodeName,OneStep.NodeValue);
        end;
      END;
    end;
  FINALLY
  	XmlDoc:= nil;
  end;
end;
{$IFEND !APPSRV}

function TconnectCEGID.GetStartUrl : string;
begin
  Result := Format('http://%s:%d/CegidFinanceWebApi/api/v1',[fServer,fport]);
end;

procedure TconnectCEGID.SetDossier(const Value: string);
begin
  fDossier := Value;
  factive :=  (fServer <> '') and (fDossier <> ''); 
end;

procedure TconnectCEGID.SetPort(const Value: string);
begin
  if Tools.IsNumeric_(Value) then fport := strtoint(Value);
  if fport = 0 then fPort := 80;
  factive :=  (fServer <> '') and (fDossier <> ''); 
end;

procedure TconnectCEGID.SetServer(const Value: string);
begin
  fServer := Value;
  factive :=  (fServer <> '') and (fDossier <> ''); 
end;

{ TGetParamWSCEGID }
class function TGetParamWSCEGID.ConnectToY2: Boolean;
begin
  Result := (GetPSoc(wspsFolder) <> '');
end;

class function TGetParamWSCEGID.GetCodeFromWsEt(WsEt : T_WSEntryType) : string;
begin
  case WsEt of
    wsetDocument           : Result := 'DOC'; // Ecriture de pièce
    wsetPayment            : Result := 'RGT'; // Ecriture de règlement
    wsetPayer              : Result := 'PAY'; // Ecriture de tiers payeur
    wsetExtourne           : Result := 'EXT'; // Ecriture d'extourne
    wsetSubContractPayment : Result := 'SCP'; // Ecriture de règlement de sous-traitance
    wsetStock              : Result := 'STK'; // Ecriture de stock
  else
    Result := '';
  end;
end;

class function TGetParamWSCEGID.GetPSoc(PSocType : T_WSPSocType) : string;
begin
  case PSocType of
    wspsServer      : Result := GetParamSocSecur(WSCDS_SocServer, '');
    wspsPort        : Result := GetParamSocSecur(WSCDS_SocNumPort, '');
    wspsFolder      : Result := GetParamSocSecur(WSCDS_SocCegidDos, '');
    wspsLastSynchro : Result := GetParamSocSecur(WSCDS_SocLastSync, '31/12/2099 23:59:59');
  else
    Result := '';
  end;
end;

end.
