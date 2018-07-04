// ************************************************************************ //
// The types declared in this file were generated from data read from the
// WSDL File described below:
// WSDL     : D:\Download\Incoming\Esker\Esker\SessionService.wsdl
// Version  : 1.0
// (13/12/2006 17:15:05 - 1.33.2.6)
// ************************************************************************ //

unit SessionService;

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

  BindingResult        = class;                 { "urn:SessionService" }
  LoginResult          = class;                 { "urn:SessionService" }
  ServiceInformation   = class;                 { "urn:SessionService" }
  SessionInformation   = class;                 { "urn:SessionService" }



  // ************************************************************************ //
  // Namespace : urn:SessionService
  // ************************************************************************ //
  BindingResult = class(TRemotable)
  private
    FsessionServiceLocation: WideString;
    FsubmissionServiceLocation: WideString;
    FqueryServiceLocation: WideString;
    FsessionServiceWSDL: WideString;
    FsubmissionServiceWSDL: WideString;
    FqueryServiceWSDL: WideString;
  published
    property sessionServiceLocation: WideString read FsessionServiceLocation write FsessionServiceLocation;
    property submissionServiceLocation: WideString read FsubmissionServiceLocation write FsubmissionServiceLocation;
    property queryServiceLocation: WideString read FqueryServiceLocation write FqueryServiceLocation;
    property sessionServiceWSDL: WideString read FsessionServiceWSDL write FsessionServiceWSDL;
    property submissionServiceWSDL: WideString read FsubmissionServiceWSDL write FsubmissionServiceWSDL;
    property queryServiceWSDL: WideString read FqueryServiceWSDL write FqueryServiceWSDL;
  end;



  // ************************************************************************ //
  // Namespace : urn:SessionService
  // ************************************************************************ //
  LoginResult = class(TRemotable)
  private
    FsessionID: WideString;
  published
    property sessionID: WideString read FsessionID write FsessionID;
  end;



  // ************************************************************************ //
  // Namespace : urn:SessionService
  // ************************************************************************ //
  ServiceInformation = class(TRemotable)
  private
    Fmessage_: WideString;
    Fdetails: WideString;
  published
    property message_: WideString read Fmessage_ write Fmessage_;
    property details: WideString read Fdetails write Fdetails;
  end;



  // ************************************************************************ //
  // Namespace : urn:SessionService
  // ************************************************************************ //
  SessionInformation = class(TRemotable)
  private
    Flogin: WideString;
    Fidentifier: WideString;
    Faccount: WideString;
    Fname_: WideString;
    Fcompany: WideString;
    Femail: WideString;
    Fculture: WideString;
    FtimeZone: WideString;
    Flanguage: WideString;
    FfilesPath: WideString;
  published
    property login: WideString read Flogin write Flogin;
    property identifier: WideString read Fidentifier write Fidentifier;
    property account: WideString read Faccount write Faccount;
    property name_: WideString read Fname_ write Fname_;
    property company: WideString read Fcompany write Fcompany;
    property email: WideString read Femail write Femail;
    property culture: WideString read Fculture write Fculture;
    property timeZone: WideString read FtimeZone write FtimeZone;
    property language: WideString read Flanguage write Flanguage;
    property filesPath: WideString read FfilesPath write FfilesPath;
  end;


  // ************************************************************************ //
  // Namespace : urn:SessionService
  // soapAction: #%operationName%
  // transport : http://schemas.xmlsoap.org/soap/http
  // style     : document
  // binding   : SessionServiceSoap
  // service   : SessionService
  // port      : SessionServiceSoap
  // URL       : https://as1.ondemand.esker.com:443/EDPWS/EDPWS.dll?Handler=Default&Version=1.0
  // ************************************************************************ //
  SessionServiceSoap = interface(IInvokable)
  ['{C6E39D8C-5C1F-3B3F-6657-386408415D8B}']
    function  GetBindings(const reserved: WideString): BindingResult; stdcall;
    function  Login(const userName: WideString; const password: WideString): LoginResult; stdcall;
    procedure Logout; stdcall;
    function  GetServiceInformation(const language: WideString): ServiceInformation; stdcall;
    function  GetSessionInformation: SessionInformation; stdcall;
  end;

function GetSessionServiceSoap(UseWSDL: Boolean=System.False; Addr: string=''; HTTPRIO: THTTPRIO = nil): SessionServiceSoap;


implementation

function GetSessionServiceSoap(UseWSDL: Boolean; Addr: string; HTTPRIO: THTTPRIO): SessionServiceSoap;
const
  defWSDL = 'D:\Download\Incoming\Esker\Esker\SessionService.wsdl';
  defURL  = 'https://as1.ondemand.esker.com:443/EDPWS/EDPWS.dll?Handler=Default&Version=1.0';
  defSvc  = 'SessionService';
  defPrt  = 'SessionServiceSoap';
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
    Result := (RIO as SessionServiceSoap);
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


initialization
  InvRegistry.RegisterInterface(TypeInfo(SessionServiceSoap), 'urn:SessionService', '');
  InvRegistry.RegisterDefaultSOAPAction(TypeInfo(SessionServiceSoap), '#%operationName%');
  InvRegistry.RegisterInvokeOptions(TypeInfo(SessionServiceSoap), ioDocument);
  InvRegistry.RegisterHeaderClass(TypeInfo(SessionServiceSoap), SessionHeaderValue, 'SessionHeaderValue', '');
  RemClassRegistry.RegisterXSClass(SessionHeaderValue, 'urn:SessionService', 'SessionHeader');
  RemClassRegistry.RegisterXSInfo(TypeInfo(SessionHeaderValue), 'urn:SessionService', 'SessionHeaderValue');
  RemClassRegistry.RegisterXSClass(BindingResult, 'urn:SessionService', 'BindingResult');
  RemClassRegistry.RegisterXSClass(LoginResult, 'urn:SessionService', 'LoginResult');
  RemClassRegistry.RegisterXSClass(ServiceInformation, 'urn:SessionService', 'ServiceInformation');
  RemClassRegistry.RegisterExternalPropName(TypeInfo(ServiceInformation), 'message_', 'message');
  RemClassRegistry.RegisterXSClass(SessionInformation, 'urn:SessionService', 'SessionInformation');
  RemClassRegistry.RegisterExternalPropName(TypeInfo(SessionInformation), 'name_', 'name');

end.