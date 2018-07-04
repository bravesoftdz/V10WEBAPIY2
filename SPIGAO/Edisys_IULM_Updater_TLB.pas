unit Edisys_IULM_Updater_TLB;

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
// Fichier généré le 27/06/2012 15:25:58 depuis la bibliothèque de types ci-dessous.

// ************************************************************************  //
// Bibl. types : C:\Program Files (x86)\Edisys\IULM\x86\Edisys.IULM.Updater.tlb (1)
// LIBID: {1D17E365-4FE8-494D-B18B-733420A26600}
// LCID: 0
// Fichier d'aide : 
// Chaîne d'aide : 
// DepndLst: 
//   (1) v2.0 stdole, (C:\Windows\SysWOW64\stdole2.tlb)
//   (2) v2.0 mscorlib, (C:\Windows\Microsoft.NET\Framework\v2.0.50727\mscorlib.tlb)
// Erreurs :
//   Erreur à la création du bitmap de palette de (TSpigaoUpdater) : Le serveur mscoree.dll ne contient pas d'icônes
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

uses Windows, ActiveX, Classes, Graphics, mscorlib_TLB, OleServer, StdVCL, Variants;
  


// *********************************************************************//
// GUIDS déclarés dans la bibliothèque de types. Préfixes utilisés :    
//   Bibliothèques de types : LIBID_xxxx                                      
//   CoClasses              : CLASS_xxxx                                      
//   DISPInterfaces         : DIID_xxxx                                       
//   Non-DISP interfaces    : IID_xxxx                                        
// *********************************************************************//
const
  // Versions majeure et mineure de la bibliothèque de types
  Edisys_IULM_UpdaterMajorVersion = 1;
  Edisys_IULM_UpdaterMinorVersion = 0;

  LIBID_Edisys_IULM_Updater: TGUID = '{1D17E365-4FE8-494D-B18B-733420A26600}';

  DIID_ISpigaoUpdater: TGUID = '{6AA6C6F4-8938-4049-A1DE-CD25C069D3B0}';
  CLASS_SpigaoUpdater: TGUID = '{D03D750A-7FA8-45C0-8212-BA2BC11D8415}';
type

// *********************************************************************//
// Déclaration Forward des types définis dans la bibliothèque de types    
// *********************************************************************//
  ISpigaoUpdater = dispinterface;

// *********************************************************************//
// Déclaration de CoClasses définies dans la bibliothèque de types 
// (REMARQUE: On affecte chaque CoClasse à son Interface par défaut)              
// *********************************************************************//
  SpigaoUpdater = ISpigaoUpdater;


// *********************************************************************//
// DispIntf :  ISpigaoUpdater
// Flags :     (4096) Dispatchable
// GUID :      {6AA6C6F4-8938-4049-A1DE-CD25C069D3B0}
// *********************************************************************//
  ISpigaoUpdater = dispinterface
    ['{6AA6C6F4-8938-4049-A1DE-CD25C069D3B0}']
    property LastErrorMessage: WideString readonly dispid 1;
    function IsUpdateAvailable: WordBool; dispid 128;
    function Update: WordBool; dispid 129;
  end;

// *********************************************************************//
// La classe CoSpigaoUpdater fournit une méthode Create et CreateRemote pour          
// créer des instances de l'interface par défaut ISpigaoUpdater exposée             
// par la CoClasse SpigaoUpdater. Les fonctions sont destinées à être utilisées par            
// les clients désirant automatiser les objets CoClasse exposés par       
// le serveur de cette bibliothèque de types.                                            
// *********************************************************************//
  CoSpigaoUpdater = class
    class function Create: ISpigaoUpdater;
    class function CreateRemote(const MachineName: string): ISpigaoUpdater;
  end;


// *********************************************************************//
// Déclaration de classe proxy de serveur OLE
// Objet serveur        : TSpigaoUpdater
// Chaîne d'aide        : 
// Interface par défaut : ISpigaoUpdater
// DISP Int. Déf. ?     : Yes
// Interface événements : 
// TypeFlags            : (2) CanCreate
// *********************************************************************//
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  TSpigaoUpdaterProperties= class;
{$ENDIF}
  TSpigaoUpdater = class(TOleServer)
  private
    FIntf:        ISpigaoUpdater;
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    FProps:       TSpigaoUpdaterProperties;
    function      GetServerProperties: TSpigaoUpdaterProperties;
{$ENDIF}
    function      GetDefaultInterface: ISpigaoUpdater;
  protected
    procedure InitServerData; override;
    function Get_LastErrorMessage: WideString;
  public
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: ISpigaoUpdater);
    procedure Disconnect; override;
    function IsUpdateAvailable: WordBool;
    function Update: WordBool;
    property DefaultInterface: ISpigaoUpdater read GetDefaultInterface;
    property LastErrorMessage: WideString read Get_LastErrorMessage;
  published
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    property Server: TSpigaoUpdaterProperties read GetServerProperties;
{$ENDIF}
  end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
// *********************************************************************//
// Classe proxy Propriétés Serveur OLE
// Objet serveur    : TSpigaoUpdater
// (Cet objet est utilisé pas l'inspecteur de propriété de l'EDI pour la 
//  modification des propriétés de ce serveur)
// *********************************************************************//
 TSpigaoUpdaterProperties = class(TPersistent)
  private
    FServer:    TSpigaoUpdater;
    function    GetDefaultInterface: ISpigaoUpdater;
    constructor Create(AServer: TSpigaoUpdater);
  protected
    function Get_LastErrorMessage: WideString;
  public
    property DefaultInterface: ISpigaoUpdater read GetDefaultInterface;
  published
  end;
{$ENDIF}


procedure Register;

resourcestring
  dtlServerPage = 'ActiveX';

  dtlOcxPage = 'ActiveX';

implementation

uses ComObj;

class function CoSpigaoUpdater.Create: ISpigaoUpdater;
begin
  Result := CreateComObject(CLASS_SpigaoUpdater) as ISpigaoUpdater;
end;

class function CoSpigaoUpdater.CreateRemote(const MachineName: string): ISpigaoUpdater;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_SpigaoUpdater) as ISpigaoUpdater;
end;

procedure TSpigaoUpdater.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{D03D750A-7FA8-45C0-8212-BA2BC11D8415}';
    IntfIID:   '{6AA6C6F4-8938-4049-A1DE-CD25C069D3B0}';
    EventIID:  '';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;

procedure TSpigaoUpdater.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk := GetServer;
    Fintf:= punk as ISpigaoUpdater;
  end;
end;

procedure TSpigaoUpdater.ConnectTo(svrIntf: ISpigaoUpdater);
begin
  Disconnect;
  FIntf := svrIntf;
end;

procedure TSpigaoUpdater.DisConnect;
begin
  if Fintf <> nil then
  begin
    FIntf := nil;
  end;
end;

function TSpigaoUpdater.GetDefaultInterface: ISpigaoUpdater;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface est NULL. Le composant n''est pas connecté au serveur. Vous devez appeler ''Connect'' ou ''ConnectTo'' avant cette opération');
  Result := FIntf;
end;

constructor TSpigaoUpdater.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps := TSpigaoUpdaterProperties.Create(Self);
{$ENDIF}
end;

destructor TSpigaoUpdater.Destroy;
begin
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps.Free;
{$ENDIF}
  inherited Destroy;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
function TSpigaoUpdater.GetServerProperties: TSpigaoUpdaterProperties;
begin
  Result := FProps;
end;
{$ENDIF}

function TSpigaoUpdater.Get_LastErrorMessage: WideString;
begin
    Result := DefaultInterface.LastErrorMessage;
end;

function TSpigaoUpdater.IsUpdateAvailable: WordBool;
begin
  Result := DefaultInterface.IsUpdateAvailable;
end;

function TSpigaoUpdater.Update: WordBool;
begin
  Result := DefaultInterface.Update;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
constructor TSpigaoUpdaterProperties.Create(AServer: TSpigaoUpdater);
begin
  inherited Create;
  FServer := AServer;
end;

function TSpigaoUpdaterProperties.GetDefaultInterface: ISpigaoUpdater;
begin
  Result := FServer.DefaultInterface;
end;

function TSpigaoUpdaterProperties.Get_LastErrorMessage: WideString;
begin
    Result := DefaultInterface.LastErrorMessage;
end;

{$ENDIF}

procedure Register;
begin
  RegisterComponents(dtlServerPage, [TSpigaoUpdater]);
end;

end.
