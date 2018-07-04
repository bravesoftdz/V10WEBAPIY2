// ************************************************************************ //
// Les types déclaré dans ce fichier ont été générés à partir de données lues dans le fichier
// WSDL décrit ci-dessous :
// WSDL     : C:\V9\branches\evol\Liens VP\WebServices\gestion Client\WSDL\WSDL.wsdl
// Encodage : utf-8
// Version  : 1.0
// (17/05/2013 13:43:11 - 1.33.2.5)
// ************************************************************************ //

unit WSDL;

interface

uses InvokeRegistry, SOAPHTTPClient, Types, XSBuiltIns;

type

  // ************************************************************************ //
  // Les types suivants mentionnés dans le document WSDL ne sont pas représentés
  // dans ce fichier. Ce sont soit des alias[@] de types représentés ou alors ils sont 
  // référencés mais jamais[!] déclarés dans ce document. Les types de la dernière catégorie
  // sont en principe mappés à des types Borland ou XML prédéfinis/connus. Toutefois, ils peuvent aussi
  // signaler des documents WSDL incorrects n'ayant pas réussi à déclarer ou importer un type de schéma.
  // ************************************************************************ //
  // !:string          - "http://www.w3.org/2001/XMLSchema"
  // !:int             - "http://www.w3.org/2001/XMLSchema"

  TWSUserParam         = class;                 { "urn:WSNomadeDecl" }
  TWSUser              = class;                 { "urn:WSNomadeDecl" }
  TWSAppel             = class;                 { "urn:WSNomadeDecl" }



  // ************************************************************************ //
  // Espace de nommage : urn:WSNomadeDecl
  // ************************************************************************ //
  TWSUserParam = class(TRemotable)
  private
    FUniqueID: WideString;
    FUser: WideString;
    FPassword: WideString;
  published
    property UniqueID: WideString read FUniqueID write FUniqueID;
    property User: WideString read FUser write FUser;
    property Password: WideString read FPassword write FPassword;
  end;



  // ************************************************************************ //
  // Espace de nommage : urn:WSNomadeDecl
  // ************************************************************************ //
  TWSUser = class(TRemotable)
  private
    FInternalID: WideString;
    FRessource: WideString;
    FCodeuser: WideString;
    FDataBaseServer: WideString;
    FDatabase: WideString;
    FCodeErreur: Integer;
    FLibErreur: WideString;
  published
    property InternalID: WideString read FInternalID write FInternalID;
    property Ressource: WideString read FRessource write FRessource;
    property Codeuser: WideString read FCodeuser write FCodeuser;
    property DataBaseServer: WideString read FDataBaseServer write FDataBaseServer;
    property Database: WideString read FDatabase write FDatabase;
    property CodeErreur: Integer read FCodeErreur write FCodeErreur;
    property LibErreur: WideString read FLibErreur write FLibErreur;
  end;



  // ************************************************************************ //
  // Espace de nommage : urn:WSNomadeDecl
  // ************************************************************************ //
  TWSAppel = class(TRemotable)
  private
    FCodeAppel: WideString;
    FEtatAppel: WideString;
    FPriorite: Integer;
    FDesignation: WideString;
    FDateAppel: WideString;
    FNomClient: WideString;
    FNomContact: WideString;
    FTelContact: WideString;
    FAdresse1: WideString;
    FAdresse2: WideString;
    FCodePostal: WideString;
    FVille: WideString;
  published
    property CodeAppel: WideString read FCodeAppel write FCodeAppel;
    property EtatAppel: WideString read FEtatAppel write FEtatAppel;
    property Priorite: Integer read FPriorite write FPriorite;
    property Designation: WideString read FDesignation write FDesignation;
    property DateAppel: WideString read FDateAppel write FDateAppel;
    property NomClient: WideString read FNomClient write FNomClient;
    property NomContact: WideString read FNomContact write FNomContact;
    property TelContact: WideString read FTelContact write FTelContact;
    property Adresse1: WideString read FAdresse1 write FAdresse1;
    property Adresse2: WideString read FAdresse2 write FAdresse2;
    property CodePostal: WideString read FCodePostal write FCodePostal;
    property Ville: WideString read FVille write FVille;
  end;

  TWSAppels  = array of TWSAppel;               { "urn:WSNomadeDecl" }

  // ************************************************************************ //
  // Espace de nommage : urn:WSIISLSENomadeIntf-IWSIISLSENomade
  // soapAction : urn:WSIISLSENomadeIntf-IWSIISLSENomade#%operationName%
  // transport : http://schemas.xmlsoap.org/soap/http
  // style     : rpc
  // liaison   : IWSIISLSENomadebinding
  // service   : IWSIISLSENomadeservice
  // port      : IWSIISLSENomadePort
  // URL       : http://vm-2008/WebServices/WSIISNomade.dll/soap/IWSIISLSENomade
  // ************************************************************************ //
  IWSIISLSENomade = interface(IInvokable)
  ['{D125CAEE-5C59-3C74-7026-4D446086C3F9}']
    function  IsValideConnect(const ParamIn: TWSUserParam): TWSUser; stdcall;
    function  GetAppels(const TheUSer: TWSUser; const Depuis: WideString): TWSAppels; stdcall;
  end;

function GetIWSIISLSENomade(UseWSDL: Boolean=System.False; Addr: string=''; HTTPRIO: THTTPRIO = nil): IWSIISLSENomade;


implementation

function GetIWSIISLSENomade(UseWSDL: Boolean; Addr: string; HTTPRIO: THTTPRIO): IWSIISLSENomade;
const
  defWSDL = 'C:\V9\branches\evol\Liens VP\WebServices\gestion Client\WSDL\WSDL.wsdl';
  defURL  = 'http://vm-2008/WebServices/WSIISNomade.dll/soap/IWSIISLSENomade';
  defSvc  = 'IWSIISLSENomadeservice';
  defPrt  = 'IWSIISLSENomadePort';
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
    Result := (RIO as IWSIISLSENomade);
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
  InvRegistry.RegisterInterface(TypeInfo(IWSIISLSENomade), 'urn:WSIISLSENomadeIntf-IWSIISLSENomade', 'utf-8');
  InvRegistry.RegisterDefaultSOAPAction(TypeInfo(IWSIISLSENomade), 'urn:WSIISLSENomadeIntf-IWSIISLSENomade#%operationName%');
  RemClassRegistry.RegisterXSClass(TWSUserParam, 'urn:WSNomadeDecl', 'TWSUserParam');
  RemClassRegistry.RegisterXSClass(TWSUser, 'urn:WSNomadeDecl', 'TWSUser');
  RemClassRegistry.RegisterXSClass(TWSAppel, 'urn:WSNomadeDecl', 'TWSAppel');
  RemClassRegistry.RegisterXSInfo(TypeInfo(TWSAppels), 'urn:WSNomadeDecl', 'TWSAppels');

end. 