// ************************************************************************ //
// The types declared in this file were generated from data read from the
// WSDL File described below:
// WSDL     : D:\Download\Incoming\Esker\Esker\SubmissionService.wsdl
// Version  : 1.0
// (13/12/2006 17:17:21 - 1.33.2.6)
// ************************************************************************ //

unit SubmissionService;

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
  // !:base64Binary    - "http://www.w3.org/2001/XMLSchema"
  // !:double          - "http://www.w3.org/2001/XMLSchema"
  // !:boolean         - "http://www.w3.org/2001/XMLSchema"

  BusinessRules        = class;                 { "urn:SubmissionService" }
  SubmissionTransport  = class;                 { "urn:SubmissionService" }
  ExtractionHeader     = class;                 { "urn:SubmissionService" }
  BusinessData         = class;                 { "urn:SubmissionService" }
  SubmissionResult     = class;                 { "urn:SubmissionService" }
  XMLDescription       = class;                 { "urn:SubmissionService" }
  SubmissionResults    = class;                 { "urn:SubmissionService" }
  ExtractionParameters = class;                 { "urn:SubmissionService" }
  ExtractionResult     = class;                 { "urn:SubmissionService" }
  ConversionParameters = class;                 { "urn:SubmissionService" }
  ConversionResult     = class;                 { "urn:SubmissionService" }
  Resources2           = class;                 { "urn:SubmissionService" }



  // ************************************************************************ //
  // Namespace : urn:SubmissionService
  // ************************************************************************ //
  BusinessRules = class(TRemotable)
  private
    FconfigurationName: WideString;
    FruleName: WideString;
  published
    property configurationName: WideString read FconfigurationName write FconfigurationName;
    property ruleName: WideString read FruleName write FruleName;
  end;


  // ************************************************************************ //
  // Namespace : urn:SubmissionService
  // ************************************************************************ //
  SubmissionTransport = class(TRemotable)
  private
    FtransportName: WideString;
    FrecipientType: WideString;
    FtransportIndex: WideString;
    FnVars: Integer;
    Fvars: TVars;
    FnSubnodes: Integer;
    Fsubnodes: TSubnodes;
    FnAttachments: Integer;
    Fattachments: TAttachments;
  public
    destructor Destroy; override;
  published
    property transportName: WideString read FtransportName write FtransportName;
    property recipientType: WideString read FrecipientType write FrecipientType;
    property transportIndex: WideString read FtransportIndex write FtransportIndex;
    property nVars: Integer read FnVars write FnVars;
    property vars: TVars read Fvars write Fvars;
    property nSubnodes: Integer read FnSubnodes write FnSubnodes;
    property subnodes: TSubnodes read Fsubnodes write Fsubnodes;
    property nAttachments: Integer read FnAttachments write FnAttachments;
    property attachments: TAttachments read Fattachments write Fattachments;
  end;


  // ************************************************************************ //
  // Namespace : urn:SubmissionService
  // ************************************************************************ //
  ExtractionHeader = class(TRemotable)
  private
    FExtractionJobID: WideString;
    FExtractionDocID: WideString;
    Foffset: Integer;
    FtransportIndex: Integer;
  published
    property ExtractionJobID: WideString read FExtractionJobID write FExtractionJobID;
    property ExtractionDocID: WideString read FExtractionDocID write FExtractionDocID;
    property offset: Integer read Foffset write Foffset;
    property transportIndex: Integer read FtransportIndex write FtransportIndex;
  end;

  ExtractionHeaderValue = ExtractionHeader;      { "urn:SubmissionService"[H] }
  TExternalVars = array of Var_;                 { "urn:SubmissionService" }


  // ************************************************************************ //
  // Namespace : urn:SubmissionService
  // ************************************************************************ //
  BusinessData = class(TRemotable)
  private
    Ffile_: WSFile;
    FnExternalVars: Integer;
    FexternalVars: TExternalVars;
    FnAttachments: Integer;
    Fattachments: TAttachments;
  public
    destructor Destroy; override;
  published
    property file_: WSFile read Ffile_ write Ffile_;
    property nExternalVars: Integer read FnExternalVars write FnExternalVars;
    property externalVars: TExternalVars read FexternalVars write FexternalVars;
    property nAttachments: Integer read FnAttachments write FnAttachments;
    property attachments: TAttachments read Fattachments write Fattachments;
  end;


  // ************************************************************************ //
  // Namespace : urn:SubmissionService
  // ************************************************************************ //
  SubmissionResult = class(TRemotable)
  private
    FsubmissionID: WideString;
    FtransportID: Integer;
  published
    property submissionID: WideString read FsubmissionID write FsubmissionID;
    property transportID: Integer read FtransportID write FtransportID;
  end;



  // ************************************************************************ //
  // Namespace : urn:SubmissionService
  // ************************************************************************ //
  XMLDescription = class(TRemotable)
  private
    FxmlFile: WSFile;
    FnAttachments: Integer;
    Fattachments: TAttachments;
  public
    destructor Destroy; override;
  published
    property xmlFile: WSFile read FxmlFile write FxmlFile;
    property nAttachments: Integer read FnAttachments write FnAttachments;
    property attachments: TAttachments read Fattachments write Fattachments;
  end;



  // ************************************************************************ //
  // Namespace : urn:SubmissionService
  // ************************************************************************ //
  SubmissionResults = class(TRemotable)
  private
    FsubmissionID: WideString;
    FnTransport: Integer;
    FtransportIDs: TTransportIDs;
  published
    property submissionID: WideString read FsubmissionID write FsubmissionID;
    property nTransport: Integer read FnTransport write FnTransport;
    property transportIDs: TTransportIDs read FtransportIDs write FtransportIDs;
  end;



  // ************************************************************************ //
  // Namespace : urn:SubmissionService
  // ************************************************************************ //
  ExtractionParameters = class(TRemotable)
  private
    FnItems: Integer;
    FfullPreviewMode: Boolean;
    FattachmentFilter: ATTACHMENTS_FILTER;
    FoutputFileMode: WSFILE_MODE;
    FincludeSubNodes: Boolean;
    FstartIndex: WideString;
  published
    property nItems: Integer read FnItems write FnItems;
    property fullPreviewMode: Boolean read FfullPreviewMode write FfullPreviewMode;
    property attachmentFilter: ATTACHMENTS_FILTER read FattachmentFilter write FattachmentFilter;
    property outputFileMode: WSFILE_MODE read FoutputFileMode write FoutputFileMode;
    property includeSubNodes: Boolean read FincludeSubNodes write FincludeSubNodes;
    property startIndex: WideString read FstartIndex write FstartIndex;
  end;

  TSubmissionTransports = array of SubmissionTransport;              { "urn:QueryService" }

  // ************************************************************************ //
  // Namespace : urn:SubmissionService
  // ************************************************************************ //
  ExtractionResult = class(TRemotable)
  private
    FnoMoreItems: Boolean;
    FnTransports: Integer;
    Ftransports: TSubmissionTransports;
  public
    destructor Destroy; override;
  published
    property noMoreItems: Boolean read FnoMoreItems write FnoMoreItems;
    property nTransports: Integer read FnTransports write FnTransports;
    property transports: TSubmissionTransports read Ftransports write Ftransports;
  end;



  // ************************************************************************ //
  // Namespace : urn:SubmissionService
  // ************************************************************************ //
  ConversionParameters = class(TRemotable)
  private
    FinputType: WideString;
    FoutputType: WideString;
    FcustomParameters: WideString;
    FoutputFileMode: WSFILE_MODE;
  published
    property inputType: WideString read FinputType write FinputType;
    property outputType: WideString read FoutputType write FoutputType;
    property customParameters: WideString read FcustomParameters write FcustomParameters;
    property outputFileMode: WSFILE_MODE read FoutputFileMode write FoutputFileMode;
  end;



  // ************************************************************************ //
  // Namespace : urn:SubmissionService
  // ************************************************************************ //
  ConversionResult = class(TRemotable)
  private
    FconvertedFile: WSFile;
  public
    destructor Destroy; override;
  published
    property convertedFile: WSFile read FconvertedFile write FconvertedFile;
  end;

  TResources  = array of WideString;             { "urn:SubmissionService" }


  // ************************************************************************ //
  // Namespace : urn:SubmissionService
  // ************************************************************************ //
  Resources2 = class(TRemotable)
  private
    FnResources: Integer;
    Fresources: TResources;
  published
    property nResources: Integer read FnResources write FnResources;
    property resources: TResources read Fresources write Fresources;
  end;


  // ************************************************************************ //
  // Namespace : urn:SubmissionService
  // soapAction: #%operationName%
  // transport : http://schemas.xmlsoap.org/soap/http
  // style     : document
  // binding   : SubmissionServiceSoap
  // service   : SubmissionService
  // port      : SubmissionServiceSoap
  // URL       : https://as1.ondemand.esker.com:443/EDPWS/EDPWS.dll?Handler=SubmissionHandler&Version=1.0
  // ************************************************************************ //
  SubmissionServiceSoap = interface(IInvokable)
  ['{95CC28FE-F277-C7A9-5D94-6ACAB93080CB}']
    function  Submit(const subject: WideString; const document: BusinessData; const rules: BusinessRules): SubmissionResult; stdcall;
    function  SubmitTransport(const transport: SubmissionTransport): SubmissionResult; stdcall;
    function  SubmitXML(const xml: XMLDescription): SubmissionResults; stdcall;
    function  ExtractFirst(const document: BusinessData; const rules: BusinessRules; const params: ExtractionParameters): ExtractionResult; stdcall;
    function  ExtractNext(const params: ExtractionParameters): ExtractionResult; stdcall;
    function  ConvertFile(const inputFile: WSFile; const params: ConversionParameters): ConversionResult; stdcall;
    function  UploadFile(const fileContent: TByteDynArray; const name_: WideString): WSFile; stdcall;
    function  DownloadFile(const wsFile: WSFile): TByteDynArray; stdcall;
    procedure RegisterResource(const resource: WSFile; const type_: RESOURCE_TYPE; const published_: Boolean; const overwritePrevious: Boolean); stdcall;
    function  ListResources(const type_: RESOURCE_TYPE; const published_: Boolean): Resources2; stdcall;
    procedure DeleteResource(const resourceName: WideString; const type_: RESOURCE_TYPE; const published_: Boolean); stdcall;
    function  RetrieveResource(const resourceName: WideString; const type_: RESOURCE_TYPE; const published_: Boolean; const eMode: WSFILE_MODE): WSFile; stdcall;
  end;

function GetSubmissionServiceSoap(UseWSDL: Boolean=System.False; Addr: string=''; HTTPRIO: THTTPRIO = nil): SubmissionServiceSoap;


implementation

function GetSubmissionServiceSoap(UseWSDL: Boolean; Addr: string; HTTPRIO: THTTPRIO): SubmissionServiceSoap;
const
  defWSDL = 'D:\Download\Incoming\Esker\Esker\SubmissionService.wsdl';
  defURL  = 'https://as1.ondemand.esker.com:443/EDPWS/EDPWS.dll?Handler=SubmissionHandler&Version=1.0';
  defSvc  = 'SubmissionService';
  defPrt  = 'SubmissionServiceSoap';
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
    Result := (RIO as SubmissionServiceSoap);
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


destructor BusinessData.Destroy;
var
  I: Integer;
begin
  for I := 0 to Length(FexternalVars)-1 do
    if Assigned(FexternalVars[I]) then
      FexternalVars[I].Free;
  SetLength(FexternalVars, 0);
  for I := 0 to Length(Fattachments)-1 do
    if Assigned(Fattachments[I]) then
      Fattachments[I].Free;
  SetLength(Fattachments, 0);
  if Assigned(Ffile_) then
    Ffile_.Free;
  inherited Destroy;
end;

destructor SubmissionTransport.Destroy;
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
  for I := 0 to Length(Fattachments)-1 do
    if Assigned(Fattachments[I]) then
      Fattachments[I].Free;
  SetLength(Fattachments, 0);
  inherited Destroy;
end;

destructor XMLDescription.Destroy;
var
  I: Integer;
begin
  for I := 0 to Length(Fattachments)-1 do
    if Assigned(Fattachments[I]) then
      Fattachments[I].Free;
  SetLength(Fattachments, 0);
  if Assigned(FxmlFile) then
    FxmlFile.Free;
  inherited Destroy;
end;

destructor ExtractionResult.Destroy;
var
  I: Integer;
begin
  for I := 0 to Length(Ftransports)-1 do
    if Assigned(Ftransports[I]) then
      Ftransports[I].Free;
  SetLength(Ftransports, 0);
  inherited Destroy;
end;

destructor ConversionResult.Destroy;
begin
  if Assigned(FconvertedFile) then
    FconvertedFile.Free;
  inherited Destroy;
end;

initialization
  InvRegistry.RegisterInterface(TypeInfo(SubmissionServiceSoap), 'urn:SubmissionService', '');
  InvRegistry.RegisterDefaultSOAPAction(TypeInfo(SubmissionServiceSoap), '#%operationName%');
  InvRegistry.RegisterInvokeOptions(TypeInfo(SubmissionServiceSoap), ioDocument);
  InvRegistry.RegisterExternalParamName(TypeInfo(SubmissionServiceSoap), 'UploadFile', 'name_', 'name');
  InvRegistry.RegisterExternalParamName(TypeInfo(SubmissionServiceSoap), 'RegisterResource', 'type_', 'type');
  InvRegistry.RegisterExternalParamName(TypeInfo(SubmissionServiceSoap), 'RegisterResource', 'published_', 'published');
  InvRegistry.RegisterExternalParamName(TypeInfo(SubmissionServiceSoap), 'ListResources', 'type_', 'type');
  InvRegistry.RegisterExternalParamName(TypeInfo(SubmissionServiceSoap), 'ListResources', 'published_', 'published');
  InvRegistry.RegisterExternalParamName(TypeInfo(SubmissionServiceSoap), 'DeleteResource', 'type_', 'type');
  InvRegistry.RegisterExternalParamName(TypeInfo(SubmissionServiceSoap), 'DeleteResource', 'published_', 'published');
  InvRegistry.RegisterExternalParamName(TypeInfo(SubmissionServiceSoap), 'RetrieveResource', 'type_', 'type');
  InvRegistry.RegisterExternalParamName(TypeInfo(SubmissionServiceSoap), 'RetrieveResource', 'published_', 'published');
  InvRegistry.RegisterHeaderClass(TypeInfo(SubmissionServiceSoap), SessionHeaderValue, 'SessionHeaderValue', '');
  RemClassRegistry.RegisterXSInfo(TypeInfo(ATTACHMENTS_FILTER), 'urn:SubmissionService', 'ATTACHMENTS_FILTER');
  RemClassRegistry.RegisterXSInfo(TypeInfo(RESOURCE_TYPE), 'urn:SubmissionService', 'RESOURCE_TYPE');
  RemClassRegistry.RegisterXSInfo(TypeInfo(WSFILE_MODE), 'urn:SubmissionService', 'WSFILE_MODE');
  RemClassRegistry.RegisterXSInfo(TypeInfo(VAR_TYPE), 'urn:SubmissionService', 'VAR_TYPE');
  RemClassRegistry.RegisterXSClass(SessionHeaderValue, 'urn:SubmissionService', 'SessionHeader');
  RemClassRegistry.RegisterXSInfo(TypeInfo(SessionHeaderValue), 'urn:SubmissionService', 'SessionHeaderValue');
  RemClassRegistry.RegisterXSClass(BusinessRules, 'urn:SubmissionService', 'BusinessRules');
  RemClassRegistry.RegisterXSClass(ExtractionHeader, 'urn:SubmissionService', 'ExtractionHeader');
  RemClassRegistry.RegisterXSInfo(TypeInfo(ExtractionHeaderValue), 'urn:SubmissionService', 'ExtractionHeaderValue');
  RemClassRegistry.RegisterXSInfo(TypeInfo(TExternalVars), 'urn:SubmissionService', 'externalVars');
  RemClassRegistry.RegisterXSInfo(TypeInfo(TAttachments), 'urn:SubmissionService', 'attachments');
  RemClassRegistry.RegisterXSClass(WSFile, 'urn:SubmissionService', 'WSFile');
  RemClassRegistry.RegisterExternalPropName(TypeInfo(WSFile), 'name_', 'name');
  RemClassRegistry.RegisterXSClass(BusinessData, 'urn:SubmissionService', 'BusinessData');
  RemClassRegistry.RegisterExternalPropName(TypeInfo(BusinessData), 'file_', 'file');
  RemClassRegistry.RegisterXSInfo(TypeInfo(TMultipleStringValues), 'urn:SubmissionService', 'multipleStringValues');
  RemClassRegistry.RegisterXSInfo(TypeInfo(TMultipleLongValues), 'urn:SubmissionService', 'multipleLongValues');
  RemClassRegistry.RegisterXSInfo(TypeInfo(TMultipleDoubleValues), 'urn:SubmissionService', 'multipleDoubleValues');
  RemClassRegistry.RegisterXSClass(Var_, 'urn:SubmissionService', 'Var_', 'Var');
  RemClassRegistry.RegisterExternalPropName(TypeInfo(Var_), 'type_', 'type');
  RemClassRegistry.RegisterXSInfo(TypeInfo(TConvertedAttachments), 'urn:SubmissionService', 'convertedAttachments');
  RemClassRegistry.RegisterXSClass(TAttachment, 'urn:SubmissionService', 'Attachment');
  RemClassRegistry.RegisterXSClass(SubmissionResult, 'urn:SubmissionService', 'SubmissionResult');
  RemClassRegistry.RegisterXSInfo(TypeInfo(TVars), 'urn:SubmissionService', 'vars');
  RemClassRegistry.RegisterXSInfo(TypeInfo(TSubnodes), 'urn:SubmissionService', 'subnodes');
  RemClassRegistry.RegisterXSClass(SubmissionTransport, 'urn:SubmissionService', 'Transport');
  RemClassRegistry.RegisterXSInfo(TypeInfo(TSubNodes2), 'urn:SubmissionService', 'subNodes2', 'subNodes');
  RemClassRegistry.RegisterXSClass(TSubNode, 'urn:SubmissionService', 'SubNode');
  RemClassRegistry.RegisterExternalPropName(TypeInfo(TSubNode), 'name_', 'name');
  RemClassRegistry.RegisterXSClass(XMLDescription, 'urn:SubmissionService', 'XMLDescription');
  RemClassRegistry.RegisterXSInfo(TypeInfo(TTransportIDs), 'urn:SubmissionService', 'transportIDs');
  RemClassRegistry.RegisterXSClass(SubmissionResults, 'urn:SubmissionService', 'SubmissionResults');
  RemClassRegistry.RegisterXSClass(ExtractionParameters, 'urn:SubmissionService', 'ExtractionParameters');
  RemClassRegistry.RegisterXSInfo(TypeInfo(TSubmissionTransports), 'urn:SubmissionService', 'transports');
  RemClassRegistry.RegisterXSClass(ExtractionResult, 'urn:SubmissionService', 'ExtractionResult');
  RemClassRegistry.RegisterXSClass(ConversionParameters, 'urn:SubmissionService', 'ConversionParameters');
  RemClassRegistry.RegisterXSClass(ConversionResult, 'urn:SubmissionService', 'ConversionResult');
  RemClassRegistry.RegisterXSInfo(TypeInfo(TResources), 'urn:SubmissionService', 'resources');
  RemClassRegistry.RegisterXSClass(Resources2, 'urn:SubmissionService', 'Resources2', 'Resources');

end.