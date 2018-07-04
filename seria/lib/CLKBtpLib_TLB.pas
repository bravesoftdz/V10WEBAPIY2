unit CLKBtpLib_TLB;

// ************************************************************************ //
// AVERTISSEMENT                                                                 
// -------                                                                    
// Les types déclarés dans ce fichier ont été générés à partir de données lues 
// depuis la bibliothèque de types. Si cette dernière (via une autre bibliothèque de types 
// s'y référant) est explicitement ou indirectement ré-importée, ou la commande "Rafraîchir"  
// de l'éditeur de bibliothèque de types est activée lors de la modification de la bibliothèque 
// de types, le contenu de ce fichier sera régénéré et toutes les modifications      
// manuellement apportées seront perdues.                                     
// ************************************************************************ //

// PASTLWTR : 1.2
// Fichier généré le 17/12/2009 13:16:06 depuis la bibliothèque de types ci-dessous.

// ************************************************************************  //
// Bibl. types : C:\WINDOWS\system32\CLKBtp.dll (1)
// LIBID: {03E4880B-E1F7-4F3A-AC96-0F1E227D92D7}
// LCID: 0
// Fichier d'aide : 
// Chaîne d'aide : Bibliothèque de types CLKBtp 1.0
// DepndLst: 
//   (1) v2.0 stdole, (C:\WINDOWS\System32\stdole2.tlb)
// Erreurs :
//   Erreur à la création du bitmap de palette de (TBuilder) : Le serveur C:\WINDOWS\system32\CLKBtp.dll ne contient pas d'icônes
// ************************************************************************ //
// *************************************************************************//
// REMARQUE :                                                                      
// Les éléments gardés par $IFDEF_LIVE_SERVER_AT_DESIGN_TIME sont  utilisés 
// par des propriétés qui renvoient des objets pouvant nécessiter d'être  
// explicitement créés par un appel de fonction avant tout accès via la   
// propriété. Ces éléments ont été désactivés afin de prévenir une utilisation  
// accidentelle depuis l'inspecteur d'objets. Vous pouvez les activer en définissant  
// LIVE_SERVER_AT_DESIGN_TIME ou en les retirant sélectivement des blocs 
//  $IFDEF. Cependant, ces éléments doivent toujours être créés par programmation
//  via une méthode de la CoClasse appropriée avant d'être utilisés.                          
{$TYPEDADDRESS OFF} // L'unité doit être compilée sans pointeur à type contrôlé. 
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}
{$VARPROPSETTER ON}
interface

uses Windows, ActiveX, Classes, Graphics, OleServer, StdVCL, Variants;
  

// *********************************************************************//
// GUIDS déclarés dans la bibliothèque de types. Préfixes utilisés :    
//   Bibliothèques de types : LIBID_xxxx                                      
//   CoClasses              : CLASS_xxxx                                      
//   DISPInterfaces         : DIID_xxxx                                       
//   Non-DISP interfaces    : IID_xxxx                                        
// *********************************************************************//
const
  // Versions majeure et mineure de la bibliothèque de types
  CLKBtpLibMajorVersion = 1;
  CLKBtpLibMinorVersion = 0;

  LIBID_CLKBtpLib: TGUID = '{03E4880B-E1F7-4F3A-AC96-0F1E227D92D7}';

  IID_IBuilder: TGUID = '{B0D592CB-7D9A-4968-A897-C07A62512867}';
  CLASS_Builder: TGUID = '{746E59FD-BB06-41E8-A0DB-4A67E4D9EC93}';
type

// *********************************************************************//
// Déclaration Forward des types définis dans la bibliothèque de types    
// *********************************************************************//
  IBuilder = interface;
  IBuilderDisp = dispinterface;

// *********************************************************************//
// Déclaration de CoClasses définies dans la bibliothèque de types 
// (REMARQUE: On affecte chaque CoClasse à son Interface par défaut)              
// *********************************************************************//
  Builder = IBuilder;


// *********************************************************************//
// Interface   : IBuilder
// Indicateurs : (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID        : {B0D592CB-7D9A-4968-A897-C07A62512867}
// *********************************************************************//
  IBuilder = interface(IDispatch)
    ['{B0D592CB-7D9A-4968-A897-C07A62512867}']
    function GetKey(HdwId: Integer; const ProductCodeVer: WideString; LicenseNum: Integer): WideString; safecall;
    function GetTempKey(HdwId: Integer; const ProductCodeVer: WideString; LicenseNum: Integer; 
                        LastMonth: Integer; LastYear: Integer): WideString; safecall;
  end;

// *********************************************************************//
// DispIntf :  IBuilderDisp
// Flags :     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID :      {B0D592CB-7D9A-4968-A897-C07A62512867}
// *********************************************************************//
  IBuilderDisp = dispinterface
    ['{B0D592CB-7D9A-4968-A897-C07A62512867}']
    function GetKey(HdwId: Integer; const ProductCodeVer: WideString; LicenseNum: Integer): WideString; dispid 1;
    function GetTempKey(HdwId: Integer; const ProductCodeVer: WideString; LicenseNum: Integer; 
                        LastMonth: Integer; LastYear: Integer): WideString; dispid 3;
  end;

// *********************************************************************//
// La classe CoBuilder fournit une méthode Create et CreateRemote pour          
// créer des instances de l'interface par défaut IBuilder exposée             
// par la CoClasse Builder. Les fonctions sont destinées à être utilisées par            
// les clients désirant automatiser les objets CoClasse exposés par       
// le serveur de cette bibliothèque de types.                                            
// *********************************************************************//
  CoBuilder = class
    class function Create: IBuilder;
    class function CreateRemote(const MachineName: string): IBuilder;
  end;


// *********************************************************************//
// Déclaration de classe proxy de serveur OLE
// Objet serveur        : TBuilder
// Chaîne d'aide        : Builder Class
// Interface par défaut : IBuilder
// DISP Int. Déf. ?     : No
// Interface événements : 
// TypeFlags            : (2) CanCreate
// *********************************************************************//
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  TBuilderProperties= class;
{$ENDIF}
  TBuilder = class(TOleServer)
  private
    FIntf:        IBuilder;
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    FProps:       TBuilderProperties;
    function      GetServerProperties: TBuilderProperties;
{$ENDIF}
    function      GetDefaultInterface: IBuilder;
  protected
    procedure InitServerData; override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: IBuilder);
    procedure Disconnect; override;
    function GetKey(HdwId: Integer; const ProductCodeVer: WideString; LicenseNum: Integer): WideString;
    function GetTempKey(HdwId: Integer; const ProductCodeVer: WideString; LicenseNum: Integer; 
                        LastMonth: Integer; LastYear: Integer): WideString;
    property DefaultInterface: IBuilder read GetDefaultInterface;
  published
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    property Server: TBuilderProperties read GetServerProperties;
{$ENDIF}
  end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
// *********************************************************************//
// Classe proxy Propriétés Serveur OLE
// Objet serveur    : TBuilder
// (Cet objet est utilisé pas l'inspecteur de propriété de l'EDI pour la 
//  modification des propriétés de ce serveur)
// *********************************************************************//
 TBuilderProperties = class(TPersistent)
  private
    FServer:    TBuilder;
    function    GetDefaultInterface: IBuilder;
    constructor Create(AServer: TBuilder);
  protected
  public
    property DefaultInterface: IBuilder read GetDefaultInterface;
  published
  end;
{$ENDIF}


procedure Register;

resourcestring
  dtlServerPage = 'ActiveX';

  dtlOcxPage = 'ActiveX';

implementation

uses ComObj;

class function CoBuilder.Create: IBuilder;
begin
  Result := CreateComObject(CLASS_Builder) as IBuilder;
end;

class function CoBuilder.CreateRemote(const MachineName: string): IBuilder;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Builder) as IBuilder;
end;

procedure TBuilder.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{746E59FD-BB06-41E8-A0DB-4A67E4D9EC93}';
    IntfIID:   '{B0D592CB-7D9A-4968-A897-C07A62512867}';
    EventIID:  '';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;

procedure TBuilder.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk := GetServer;
    Fintf:= punk as IBuilder;
  end;
end;

procedure TBuilder.ConnectTo(svrIntf: IBuilder);
begin
  Disconnect;
  FIntf := svrIntf;
end;

procedure TBuilder.DisConnect;
begin
  if Fintf <> nil then
  begin
    FIntf := nil;
  end;
end;

function TBuilder.GetDefaultInterface: IBuilder;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface est NULL. Le composant n''est pas connecté au serveur. Vous devez appeler ''Connect'' ou ''ConnectTo'' avant cette opération');
  Result := FIntf;
end;

constructor TBuilder.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps := TBuilderProperties.Create(Self);
{$ENDIF}
end;

destructor TBuilder.Destroy;
begin
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps.Free;
{$ENDIF}
  inherited Destroy;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
function TBuilder.GetServerProperties: TBuilderProperties;
begin
  Result := FProps;
end;
{$ENDIF}

function TBuilder.GetKey(HdwId: Integer; const ProductCodeVer: WideString; LicenseNum: Integer): WideString;
begin
  Result := DefaultInterface.GetKey(HdwId, ProductCodeVer, LicenseNum);
end;

function TBuilder.GetTempKey(HdwId: Integer; const ProductCodeVer: WideString; LicenseNum: Integer; 
                             LastMonth: Integer; LastYear: Integer): WideString;
begin
  Result := DefaultInterface.GetTempKey(HdwId, ProductCodeVer, LicenseNum, LastMonth, LastYear);
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
constructor TBuilderProperties.Create(AServer: TBuilder);
begin
  inherited Create;
  FServer := AServer;
end;

function TBuilderProperties.GetDefaultInterface: IBuilder;
begin
  Result := FServer.DefaultInterface;
end;

{$ENDIF}

procedure Register;
begin
  RegisterComponents(dtlServerPage, [TBuilder]);
end;

end.

