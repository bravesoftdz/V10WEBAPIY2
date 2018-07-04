// ************************************************************************ //
// The types declared in this file were generated from data read from the
// WSDL File described below:
// WSDL     : D:\Download\Incoming\Esker\Esker\QueryService.wsdl
// Version  : 1.0
// (13/12/2006 17:16:36 - 1.33.2.6)
// ************************************************************************ //

unit QueryService;

interface

uses InvokeRegistry, SOAPHTTPClient, Types, XSBuiltIns, WSEskerUtil;

type

  // ************************************************************************ //
  // The following types, referred to in the WSDL document are not being represented
  // in this file. They are either aliases[@] of other types represented or were referred
  // to but never[!] declared in the document. The types from the latter category
  // typically map to predefined/known XML or Borland types; however, they could also
  // indicate incorrect WSDL documents that failed to declare or import a schema type.
  // ************************************************************************ //
  // !:string          - "http://www.w3.org/2001/XMLSchema"
  // !:int             - "http://www.w3.org/2001/XMLSchema"
  // !:boolean         - "http://www.w3.org/2001/XMLSchema"
  // !:double          - "http://www.w3.org/2001/XMLSchema"
  // !:base64Binary    - "http://www.w3.org/2001/XMLSchema"

  QueryHeaderValue     = class;                 { "urn:QueryService" }
  QueryRequest         = class;                 { "urn:QueryService" }
  Transport            = class;                 { "urn:QueryService" }
  QueryResult          = class;                 { "urn:QueryService" }
  QueryAttributesResult = class;                { "urn:QueryService" }
  Attachments2         = class;                 { "urn:QueryService" }
  StatisticsResult     = class;                 { "urn:QueryService" }
  StatisticsLine       = class;                 { "urn:QueryService" }
  ActionResult         = class;                 { "urn:QueryService" }
  ResubmitParameters   = class;                 { "urn:QueryService" }



  // ************************************************************************ //
  // Namespace : urn:QueryService
  // ************************************************************************ //
  QueryHeaderValue = class(TRemotable)
  private
    FqueryID: WideString;
  published
    property queryID: WideString read FqueryID write FqueryID;
  end;

//  QueryHeaderValue = QueryHeader;      { "urn:QueryService"[H] }


  // ************************************************************************ //
  // Namespace : urn:QueryService
  // ************************************************************************ //
  QueryRequest = class(TRemotable)
  private
    Ffilter: WideString;
    FsortOrder: WideString;
    Fattributes: WideString;
    FnItems: Integer;
    FincludeSubNodes: Boolean;
  published
    property filter: WideString read Ffilter write Ffilter;
    property sortOrder: WideString read FsortOrder write FsortOrder;
    property attributes: WideString read Fattributes write Fattributes;
    property nItems: Integer read FnItems write FnItems;
    property includeSubNodes: Boolean read FincludeSubNodes write FincludeSubNodes;
  end;


  // ************************************************************************ //
  // Namespace : urn:QueryService
  // ************************************************************************ //
  Transport = class(TRemotable)
  private
    FtransportID: Integer;
    FtransportName: WideString;
    FrecipientType: WideString;
    Fstate: Integer;
    FnVars: Integer;
    Fvars: TVars;
    FnSubnodes: Integer;
    Fsubnodes: TSubnodes;
  public
    destructor Destroy; override;
  published
    property transportID: Integer read FtransportID write FtransportID;
    property transportName: WideString read FtransportName write FtransportName;
    property recipientType: WideString read FrecipientType write FrecipientType;
    property state: Integer read Fstate write Fstate;
    property nVars: Integer read FnVars write FnVars;
    property vars: TVars read Fvars write Fvars;
    property nSubnodes: Integer read FnSubnodes write FnSubnodes;
    property subnodes: TSubnodes read Fsubnodes write Fsubnodes;
  end;

  TTransports = array of Transport;              { "urn:QueryService" }

  // ************************************************************************ //
  // Namespace : urn:QueryService
  // ************************************************************************ //
  QueryResult = class(TRemotable)
  private
    FnoMoreItems: Boolean;
    FnTransports: Integer;
    Ftransports: TTransports;
  public
    destructor Destroy; override;
  published
    property noMoreItems: Boolean read FnoMoreItems write FnoMoreItems;
    property nTransports: Integer read FnTransports write FnTransports;
    property transports: TTransports read Ftransports write Ftransports;
  end;


  TAttributes = array of WideString;             { "urn:QueryService" }


  // ************************************************************************ //
  // Namespace : urn:QueryService
  // ************************************************************************ //
  QueryAttributesResult = class(TRemotable)
  private
    FnAttributes: Integer;
    Fattributes: TAttributes;
  published
    property nAttributes: Integer read FnAttributes write FnAttributes;
    property attributes: TAttributes read Fattributes write Fattributes;
  end;


  // ************************************************************************ //
  // Namespace : urn:QueryService
  // ************************************************************************ //
  Attachments2 = class(TRemotable)
  private
    FnAttachments: Integer;
    Fattachments: TAttachments;
  public
    destructor Destroy; override;
  published
    property nAttachments: Integer read FnAttachments write FnAttachments;
    property attachments: TAttachments read Fattachments write Fattachments;
  end;



  TTypeName   = array of WideString;             { "urn:QueryService" }
  TTypeContent = array of StatisticsLine;        { "urn:QueryService" }


  // ************************************************************************ //
  // Namespace : urn:QueryService
  // ************************************************************************ //
  StatisticsResult = class(TRemotable)
  private
    FnTypes: Integer;
    FtypeName: TTypeName;
    FtypeContent: TTypeContent;
  public
    destructor Destroy; override;
  published
    property nTypes: Integer read FnTypes write FnTypes;
    property typeName: TTypeName read FtypeName write FtypeName;
    property typeContent: TTypeContent read FtypeContent write FtypeContent;
  end;

  TStates     = array of Integer;                { "urn:QueryService" }
  TCounts     = array of Integer;                { "urn:QueryService" }


  // ************************************************************************ //
  // Namespace : urn:QueryService
  // ************************************************************************ //
  StatisticsLine = class(TRemotable)
  private
    FnStates: Integer;
    Fstates: TStates;
    Fcounts: TCounts;
  published
    property nStates: Integer read FnStates write FnStates;
    property states: TStates read Fstates write Fstates;
    property counts: TCounts read Fcounts write Fcounts;
  end;

  TErrorReason = array of WideString;            { "urn:QueryService" }


  // ************************************************************************ //
  // Namespace : urn:QueryService
  // ************************************************************************ //
  ActionResult = class(TRemotable)
  private
    FnSucceeded: Integer;
    FnFailed: Integer;
    FnItem: Integer;
    FtransportIDs: TTransportIDs;
    FerrorReason: TErrorReason;
  published
    property nSucceeded: Integer read FnSucceeded write FnSucceeded;
    property nFailed: Integer read FnFailed write FnFailed;
    property nItem: Integer read FnItem write FnItem;
    property transportIDs: TTransportIDs read FtransportIDs write FtransportIDs;
    property errorReason: TErrorReason read FerrorReason write FerrorReason;
  end;



  // ************************************************************************ //
  // Namespace : urn:QueryService
  // ************************************************************************ //
  ResubmitParameters = class(TRemotable)
  private
    FnVars: Integer;
    Fvars: TVars;
  public
    destructor Destroy; override;
  published
    property nVars: Integer read FnVars write FnVars;
    property vars: TVars read Fvars write Fvars;
  end;


  // ************************************************************************ //
  // Namespace : urn:QueryService
  // soapAction: #%operationName%
  // transport : http://schemas.xmlsoap.org/soap/http
  // style     : document
  // binding   : QueryServiceSoap
  // service   : QueryService
  // port      : QueryServiceSoap
  // URL       : https://as1.ondemand.esker.com:443/EDPWS/EDPWS.dll?Handler=QueryHandler&Version=1.0
  // ************************************************************************ //
  QueryServiceSoap = interface(IInvokable)
  ['{029D17D6-9165-2E1D-8D61-59E099F0CC5B}']
    function  QueryFirst(const request: QueryRequest): QueryResult; stdcall;
    function  QueryLast(const request: QueryRequest): QueryResult; stdcall;
    function  QueryNext(const request: QueryRequest): QueryResult; stdcall;
    function  QueryPrevious(const request: QueryRequest): QueryResult; stdcall;
    function  QueryAttributes(const request: QueryRequest): QueryAttributesResult; stdcall;
    function  QueryAttachments(const transportID: Integer; const eFilter: ATTACHMENTS_FILTER; const eMode: WSFILE_MODE): Attachments2; stdcall;
    function  QueryStatistics(const filter: WideString): StatisticsResult; stdcall;
    function  Delete(const identifier: WideString): ActionResult; stdcall;
    function  Cancel(const identifier: WideString): ActionResult; stdcall;
    function  Resubmit(const identifier: WideString; const params: ResubmitParameters): ActionResult; stdcall;
    function  Approve(const identifier: WideString; const reason: WideString): ActionResult; stdcall;
    function  Reject(const identifier: WideString; const reason: WideString): ActionResult; stdcall;
    function  DownloadFile(const wsFile: WSFile): TByteDynArray; stdcall;
  end;

function GetQueryServiceSoap(UseWSDL: Boolean=System.False; Addr: string=''; HTTPRIO: THTTPRIO = nil): QueryServiceSoap;


implementation

function GetQueryServiceSoap(UseWSDL: Boolean; Addr: string; HTTPRIO: THTTPRIO): QueryServiceSoap;
const
  defWSDL = 'D:\Download\Incoming\Esker\Esker\QueryService.wsdl';
  defURL  = 'https://as1.ondemand.esker.com:443/EDPWS/EDPWS.dll?Handler=QueryHandler&Version=1.0';
  defSvc  = 'QueryService';
  defPrt  = 'QueryServiceSoap';
var
  RIO: THTTPRIO;
begin
  Result := nil;
  if (Addr = '') then
  begin
    if UseWSDL then
      Addr := defWSDL
    else
      Addr := defURL;
  end;
  if HTTPRIO = nil then
    RIO := THTTPRIO.Create(nil)
  else
    RIO := HTTPRIO;
  try
    Result := (RIO as QueryServiceSoap);
    if UseWSDL then
    begin
      RIO.WSDLLocation := Addr;
      RIO.Service := defSvc;
      RIO.Port := defPrt;
    end else
      RIO.URL := Addr;
  finally
    if (Result = nil) and (HTTPRIO = nil) then
      RIO.Free;
  end;
end;

destructor Transport.Destroy;
var
  I: Integer;
begin
  for I := 0 to Length(Fvars)-1 do
    if Assigned(Fvars[I]) then
      Fvars[I].Free;
  SetLength(Fvars, 0);
  for I := 0 to Length(Fsubnodes)-1 do
    if Assigned(Fsubnodes[I]) then
      Fsubnodes[I].Free;
  SetLength(Fsubnodes, 0);
  inherited Destroy;
end;

destructor QueryResult.Destroy;
var
  I: Integer;
begin
  for I := 0 to Length(Ftransports)-1 do
    if Assigned(Ftransports[I]) then
      Ftransports[I].Free;
  SetLength(Ftransports, 0);
  inherited Destroy;
end;

destructor Attachments2.Destroy;
var
  I: Integer;
begin
  for I := 0 to Length(Fattachments)-1 do
    if Assigned(Fattachments[I]) then
      Fattachments[I].Free;
  SetLength(Fattachments, 0);
  inherited Destroy;
end;

destructor StatisticsResult.Destroy;
var
  I: Integer;
begin
  for I := 0 to Length(FtypeContent)-1 do
    if Assigned(FtypeContent[I]) then
      FtypeContent[I].Free;
  SetLength(FtypeContent, 0);
  inherited Destroy;
end;

destructor ResubmitParameters.Destroy;
var
  I: Integer;
begin
  for I := 0 to Length(Fvars)-1 do
    if Assigned(Fvars[I]) then
      Fvars[I].Free;
  SetLength(Fvars, 0);
  inherited Destroy;
end;

initialization
  InvRegistry.RegisterInterface(TypeInfo(QueryServiceSoap), 'urn:QueryService', '');
  InvRegistry.RegisterDefaultSOAPAction(TypeInfo(QueryServiceSoap), '#%operationName%');
  InvRegistry.RegisterInvokeOptions(TypeInfo(QueryServiceSoap), ioDocument);
  InvRegistry.RegisterHeaderClass(TypeInfo(QueryServiceSoap), SessionHeaderValue, 'SessionHeaderValue', '');
  RemClassRegistry.RegisterXSInfo(TypeInfo(ATTACHMENTS_FILTER), 'urn:QueryService', 'ATTACHMENTS_FILTER');
  RemClassRegistry.RegisterXSInfo(TypeInfo(WSFILE_MODE), 'urn:QueryService', 'WSFILE_MODE');
  RemClassRegistry.RegisterXSInfo(TypeInfo(VAR_TYPE), 'urn:QueryService', 'VAR_TYPE');
  RemClassRegistry.RegisterXSClass(QueryHeaderValue, 'urn:QueryService', 'QueryHeader');
  RemClassRegistry.RegisterXSInfo(TypeInfo(QueryHeaderValue), 'urn:QueryService', 'QueryHeaderValue');
  RemClassRegistry.RegisterXSClass(SessionHeaderValue, 'urn:QueryService', 'SessionHeader');
  RemClassRegistry.RegisterXSInfo(TypeInfo(SessionHeaderValue), 'urn:QueryService', 'SessionHeaderValue');
  RemClassRegistry.RegisterXSClass(QueryRequest, 'urn:QueryService', 'QueryRequest');
  RemClassRegistry.RegisterXSInfo(TypeInfo(TTransports), 'urn:QueryService', 'transports');
  RemClassRegistry.RegisterXSClass(QueryResult, 'urn:QueryService', 'QueryResult');
  RemClassRegistry.RegisterXSInfo(TypeInfo(TVars), 'urn:QueryService', 'vars');
  RemClassRegistry.RegisterXSInfo(TypeInfo(TSubnodes), 'urn:QueryService', 'subnodes');
  RemClassRegistry.RegisterXSClass(Transport, 'urn:QueryService', 'Transport');
  RemClassRegistry.RegisterXSInfo(TypeInfo(TMultipleStringValues), 'urn:QueryService', 'multipleStringValues');
  RemClassRegistry.RegisterXSInfo(TypeInfo(TMultipleLongValues), 'urn:QueryService', 'multipleLongValues');
  RemClassRegistry.RegisterXSInfo(TypeInfo(TMultipleDoubleValues), 'urn:QueryService', 'multipleDoubleValues');
  RemClassRegistry.RegisterXSClass(Var_, 'urn:QueryService', 'Var_', 'Var');
  RemClassRegistry.RegisterExternalPropName(TypeInfo(Var_), 'type_', 'type');
  RemClassRegistry.RegisterXSInfo(TypeInfo(TSubNodes2), 'urn:QueryService', 'subNodes2', 'subNodes');
  RemClassRegistry.RegisterXSClass(TSubNode, 'urn:QueryService', 'SubNode');
  RemClassRegistry.RegisterExternalPropName(TypeInfo(TSubNode), 'name_', 'name');
  RemClassRegistry.RegisterXSInfo(TypeInfo(TAttributes), 'urn:QueryService', 'attributes');
  RemClassRegistry.RegisterXSClass(QueryAttributesResult, 'urn:QueryService', 'QueryAttributesResult');
  RemClassRegistry.RegisterXSInfo(TypeInfo(TAttachments), 'urn:QueryService', 'attachments');
  RemClassRegistry.RegisterXSClass(Attachments2, 'urn:QueryService', 'Attachments2', 'Attachments');
  RemClassRegistry.RegisterXSInfo(TypeInfo(TConvertedAttachments), 'urn:QueryService', 'convertedAttachments');
  RemClassRegistry.RegisterXSClass(WSFile, 'urn:QueryService', 'WSFile');
  RemClassRegistry.RegisterExternalPropName(TypeInfo(WSFile), 'name_', 'name');
  RemClassRegistry.RegisterXSClass(TAttachment, 'urn:QueryService', 'Attachment');
  RemClassRegistry.RegisterXSInfo(TypeInfo(TTypeName), 'urn:QueryService', 'typeName');
  RemClassRegistry.RegisterXSInfo(TypeInfo(TTypeContent), 'urn:QueryService', 'typeContent');
  RemClassRegistry.RegisterXSClass(StatisticsResult, 'urn:QueryService', 'StatisticsResult');
  RemClassRegistry.RegisterXSInfo(TypeInfo(TStates), 'urn:QueryService', 'states');
  RemClassRegistry.RegisterXSInfo(TypeInfo(TCounts), 'urn:QueryService', 'counts');
  RemClassRegistry.RegisterXSClass(StatisticsLine, 'urn:QueryService', 'StatisticsLine');
  RemClassRegistry.RegisterXSInfo(TypeInfo(TTransportIDs), 'urn:QueryService', 'transportIDs');
  RemClassRegistry.RegisterXSInfo(TypeInfo(TErrorReason), 'urn:QueryService', 'errorReason');
  RemClassRegistry.RegisterXSClass(ActionResult, 'urn:QueryService', 'ActionResult');
  RemClassRegistry.RegisterXSClass(ResubmitParameters, 'urn:QueryService', 'ResubmitParameters');

end.