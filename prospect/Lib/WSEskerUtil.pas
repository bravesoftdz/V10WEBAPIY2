unit WSEskerUtil;

interface

uses InvokeRegistry, SOAPHTTPClient, Types, XSBuiltIns;

type

  SessionHeaderValue   = class;                 { "urn:SessionService" }
  WSFile               = class;                 { "urn:SubmissionService" }
  Var_                 = class;                 { "urn:SubmissionService" }
  TAttachment           = class;                 { "urn:SubmissionService" }
  TSubNode              = class;                 { "urn:SubmissionService" }

  TVars       = array of Var_;                   { "urn:QueryService" }
  TSubnodes   = array of TSubNode;                { "urn:QueryService" }
  TAttachments = array of TAttachment;            { "urn:QueryService" }
  TMultipleStringValues = array of WideString;   { "urn:QueryService" }
  TMultipleLongValues = array of Integer;        { "urn:QueryService" }
  TMultipleDoubleValues = array of Double;       { "urn:QueryService" }
  TSubNodes2  = array of WideString;             { "urn:QueryService" }
  TTransportIDs = array of Integer;              { "urn:QueryService" }
  TConvertedAttachments = array of WSFile;       { "urn:QueryService" }


  // ************************************************************************ //
  // Namespace : urn:SessionService
  // ************************************************************************ //
  SessionHeaderValue = class(TRemotable)
  private
    FsessionID: WideString;
  published
    property sessionID: WideString read FsessionID write FsessionID;
  end;

//  SessionHeaderValue = SessionHeader;      { "urn:SessionService"[H] }


  { "urn:SubmissionService" }
  ATTACHMENTS_FILTER = (FILTER_NONE, FILTER_ALL, FILTER_CONVERTED, FILTER_SOURCE);

  { "urn:SubmissionService" }
//  RESOURCE_TYPE = (TYPE_SALESFORCE_LEAD_EMAILBODY, TYPE_SALESFORCE_LEAD_FAXCOVER, TYPE_SALESFORCE_CONTACT_EMAILBODY, TYPE_SALESFORCE_CONTACT_FAXCOVER, TYPE_SALESFORCE_SETTINGS, TYPE_STYLESHEET, TYPE_IMAGE, TYPE_COVER);

  { "urn:SubmissionService" }
  WSFILE_MODE = (MODE_UNDEFINED, MODE_ON_SERVER, MODE_INLINED);

  { "urn:SubmissionService" }
  VAR_TYPE = (TYPE_STRING, TYPE_DATETIME, TYPE_DOUBLE, TYPE_INTEGER);

  RESOURCE_TYPE = (TYPE_COVER, TYPE_IMAGE, TYPE_STYLESHEET);

  // ************************************************************************ //
  // Namespace : urn:QueryService
  // ************************************************************************ //
  Var_ = class(TRemotable)
  private
    Fattribute: WideString;
    Ftype_: VAR_TYPE;
    FsimpleValue: WideString;
    FnValues: Integer;
    FmultipleStringValues: TMultipleStringValues;
    FmultipleLongValues: TMultipleLongValues;
    FmultipleDoubleValues: TMultipleDoubleValues;
  published
    property attribute: WideString read Fattribute write Fattribute;
    property type_: VAR_TYPE read Ftype_ write Ftype_;
    property simpleValue: WideString read FsimpleValue write FsimpleValue;
    property nValues: Integer read FnValues write FnValues;
    property multipleStringValues: TMultipleStringValues read FmultipleStringValues write FmultipleStringValues;
    property multipleLongValues: TMultipleLongValues read FmultipleLongValues write FmultipleLongValues;
    property multipleDoubleValues: TMultipleDoubleValues read FmultipleDoubleValues write FmultipleDoubleValues;
  end;


  // ************************************************************************ //
  // Namespace : urn:QueryService
  // ************************************************************************ //
  TSubNode = class(TRemotable)
  private
    Fname_: WideString;
    FrelativeName: WideString;
    FnSubnodes: Integer;
    FsubNodes: TSubNodes2;
    FnVars: Integer;
    Fvars: TVars;
  public
    destructor Destroy; override;
  published
    property name_: WideString read Fname_ write Fname_;
    property relativeName: WideString read FrelativeName write FrelativeName;
    property nSubnodes: Integer read FnSubnodes write FnSubnodes;
    property subNodes: TSubNodes2 read FsubNodes write FsubNodes;
    property nVars: Integer read FnVars write FnVars;
    property vars: TVars read Fvars write Fvars;
  end;

  // ************************************************************************ //
  // Namespace : urn:QueryService
  // ************************************************************************ //
  WSFile = class(TRemotable)
  private
    Fname_: WideString;
    Fmode: WSFILE_MODE;
    Fcontent: TByteDynArray;
    Furl: WideString;
    FstorageID: WideString;
  published
    property name_: WideString read Fname_ write Fname_;
    property mode: WSFILE_MODE read Fmode write Fmode;
    property content: TByteDynArray read Fcontent write Fcontent;
    property url: WideString read Furl write Furl;
    property storageID: WideString read FstorageID write FstorageID;
  end;



  // ************************************************************************ //
  // Namespace : urn:QueryService
  // ************************************************************************ //
  TAttachment = class(TRemotable)
  private
    FinputFormat: WideString;
    FoutputFormat: WideString;
    Fstylesheet: WideString;
    FoutputName: WideString;
    FsourceAttachment: WSFile;
    FnConvertedAttachments: Integer;
    FconvertedAttachments: TConvertedAttachments;
  public
    destructor Destroy; override;
  published
    property inputFormat: WideString read FinputFormat write FinputFormat;
    property outputFormat: WideString read FoutputFormat write FoutputFormat;
    property stylesheet: WideString read Fstylesheet write Fstylesheet;
    property outputName: WideString read FoutputName write FoutputName;
    property sourceAttachment: WSFile read FsourceAttachment write FsourceAttachment;
    property nConvertedAttachments: Integer read FnConvertedAttachments write FnConvertedAttachments;
    property convertedAttachments: TConvertedAttachments read FconvertedAttachments write FconvertedAttachments;
  end;

implementation

destructor TSubNode.Destroy;
var
  I: Integer;
begin
  for I := 0 to Length(Fvars)-1 do
    if Assigned(Fvars[I]) then
      Fvars[I].Free;
  SetLength(Fvars, 0);
  inherited Destroy;
end;

destructor TAttachment.Destroy;
var
  I: Integer;
begin
  for I := 0 to Length(FconvertedAttachments)-1 do
    if Assigned(FconvertedAttachments[I]) then
      FconvertedAttachments[I].Free;
  SetLength(FconvertedAttachments, 0);
  if Assigned(FsourceAttachment) then
    FsourceAttachment.Free;
  inherited Destroy;
end;

end.
