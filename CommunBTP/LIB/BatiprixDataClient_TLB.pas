unit BatiprixDataClient_TLB;

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

// PASTLWTR : $Revision:   1.88.1.0.1.0  $
// Fichier généré le 12/12/2007 13:04:37 depuis la bibliothèque de types ci-dessous.

// *************************************************************************//
// REMARQUE :                                                                   
// Les éléments gardés par $IFDEF_LIVE_SERVER_AT_DESIGN_TIME sont utilisés par  
// des propriétés qui renvoient des objets pouvant nécessiter d'être créés par  
// un appel de fonction avant tout accès via la propriété. Ces éléments ont été 
// désactivés pour prévenir une utilisation accidentelle depuis l'Inspecteur
// d'objets. Vous pouvez les activer en définissant LIVE_SERVER_AT_DESIGN_TIME  
// ou en les retirant sélectivement des blocs $IFDEF. Cependant, ces éléments   
// doivent toujours être créés par programmation via une méthode da la CoClass  
// appropriée avant d'être utilisés.                                            
// ************************************************************************ //
// Bibl.Types     : \\srv-btp\Src\PGIS5\COMMUN\Batiprix\BatiprixDataClient.dll (1)
// IID\LCID       : {FC804AA5-6A82-426E-BCEA-0545FC319AD6}\0
// Fichier d'aide :
// DepndLst       :
//   (1) v2.0 stdole, (C:\WINDOWS\system32\stdole2.tlb)
//   (2) v4.0 StdVCL, (C:\WINDOWS\system32\STDVCL40.DLL)
// Erreurs :
//   Erreur à la création du bitmap de palette de (TCommandeMaster) : Format GUID incorrect
// ************************************************************************ //
{$TYPEDADDRESS OFF} // L'unité doit être compilée sans vérification de type des pointeurs.
interface

uses Windows, ActiveX, Classes, Graphics, OleServer, OleCtrls, StdVCL,SysUtils;

// *********************************************************************//
// GUIDS déclarés dans la bibliothèque de types. Préfixes utilisés :
//   Bibliothèques de types : LIBID_xxxx
//   CoClasses              : CLASS_xxxx
//   DISPInterfaces         : DIID_xxxx
//   Non-DISP interfaces    : IID_xxxx
// *********************************************************************//
const
  // Version majeure et mineure de la bibliothèque de types
  BatiprixDataClientMajorVersion = 1;
  BatiprixDataClientMinorVersion = 0;

  LIBID_BatiprixDataClient: TGUID = '{FC804AA5-6A82-426E-BCEA-0545FC319AD6}';

  IID_ICommandeMaster: TGUID = '{C82FE647-0682-4DB3-AFB6-406E4B9EB1C1}';
  DIID_ICommandeMasterEvents: TGUID = '{09579A3B-FD5F-4EA1-97B9-2B1EAE5D6657}';
  CLASS_CommandeMaster: TGUID = '{ED55839D-D452-4214-80E0-8EAA702D760C}';
type

  Tcommande = class
    private
      flibelle : Widestring;
      fprix : double;
    public
      property Libelle : widestring read flibelle write flibelle;
      property Prix : double read fprix write fprix;
  end;

  TdetailCommande = class
    private
      fLibelle : WideString;
      fprix : double;
    public
      property Libelle : WideString read fLibelle write Flibelle;
      property prix : double read fprix Write fprix;
  end;

  TListDetailCommande = class (Tlist)
    private
      function GetItems(Indice : integer): TDetailCommande;
      procedure SetItems(Indice : integer; const Value: TDetailCommande);
      function Add(AObject: TDetailCommande): Integer;
    public
      property Items [Indice : integer] : TDetailCommande read GetItems write SetItems;
      procedure clear; override;
      destructor destroy ; override;
  end;

// *********************************************************************//
// Déclaration Forward des types définis dans la bibliothèque de types
// *********************************************************************//
  ICommandeMaster = interface;
  ICommandeMasterEvents = dispinterface;

// *********************************************************************//
// Déclaration de CoClasses définies dans la bibliothèque de types
// (REMARQUE: On affecte chaque CoClass à son Interface par défaut)
// *********************************************************************//
  CommandeMaster = ICommandeMaster;


// *********************************************************************//
// Interface   : ICommandeMaster
// Indicateurs : (256) OleAutomation
// GUID        : {C82FE647-0682-4DB3-AFB6-406E4B9EB1C1}
// *********************************************************************//
  ICommandeMaster = interface(IUnknown)
    ['{C82FE647-0682-4DB3-AFB6-406E4B9EB1C1}']
    function  InstallEdition(const RepInstall: WideString; const RepSource: WideString;
                             out ErreurCode: SYSINT; out ErreurMessage: WideString): HResult; stdcall;
    function  GenereCommande(const RepInstall: WideString; out ErreurCode: SYSINT;
                             out ErreurMessage: WideString): HResult; stdcall;
    function  InstallCommande(const RepSource: WideString; const RepDestination: WideString;
                              out ErreurCode: SYSINT; out ErreurMessage: WideString): HResult; stdcall;
    function  Get_NomCommercial(out Value: WideString): HResult; stdcall;
    function  Set_NomCommercial(const Value: WideString): HResult; stdcall;
    function  Get_NoSerie(out Value: WideString): HResult; stdcall;
    function  Set_NoSerie(const Value: WideString): HResult; stdcall;
    function  Get_RaisonSociale(out Value: WideString): HResult; stdcall;
    function  Set_RaisonSociale(const Value: WideString): HResult; stdcall;
    function  Get_Nom(out Value: WideString): HResult; stdcall;
    function  Set_Nom(const Value: WideString): HResult; stdcall;
    function  Get_Prenom(out Value: WideString): HResult; stdcall;
    function  Set_Prenom(const Value: WideString): HResult; stdcall;
    function  Get_Adresse1(out Value: WideString): HResult; stdcall;
    function  Set_Adresse1(const Value: WideString): HResult; stdcall;
    function  Get_Adresse2(out Value: WideString): HResult; stdcall;
    function  Set_Adresse2(const Value: WideString): HResult; stdcall;
    function  Get_CodePostal(out Value: WideString): HResult; stdcall;
    function  Set_CodePostal(const Value: WideString): HResult; stdcall;
    function  Get_Ville(out Value: WideString): HResult; stdcall;
    function  Set_Ville(const Value: WideString): HResult; stdcall;
    function  Get_Pays(out Value: WideString): HResult; stdcall;
    function  Set_Pays(const Value: WideString): HResult; stdcall;
    function  Get_Tel(out Value: WideString): HResult; stdcall;
    function  Set_Tel(const Value: WideString): HResult; stdcall;
    function  Get_Fax(out Value: WideString): HResult; stdcall;
    function  Set_Fax(const Value: WideString): HResult; stdcall;
    function  Get_Email(out Value: WideString): HResult; stdcall;
    function  Set_Email(const Value: WideString): HResult; stdcall;
    function  Get_RepertoireInfoMaster(out Value: WideString): HResult; stdcall;
    function  Get_CommandeLibelle(out Value: WideString): HResult; stdcall;
    function  Get_CommandePrix(out Value: Currency): HResult; stdcall;
    function  Get_ListeCommandeLibelle(out Value: PSafeArray): HResult; stdcall;
    function  Get_ListeCommandePrix(out Value: PSafeArray): HResult; stdcall;
    function  Get_ClefCommande(out Value: WideString): HResult; stdcall;
    function  Get_SommeRaisonSociale(out Value: WideString): HResult; stdcall;
    function  Get_SommeNom(out Value: WideString): HResult; stdcall;
    function  Get_SommeClefCommande(out Value: WideString): HResult; stdcall;
    function  Get_Version(out Value: WideString): HResult; stdcall;
    function  Get_FichierComposant(out Value: WideString): HResult; stdcall;
    function  Get_InstalleMultiposte(out Value: Integer): HResult; stdcall;
    function  Set_InstalleMultiposte(Value: Integer): HResult; stdcall;
    function  Get_NbPostes(out Value: Integer): HResult; stdcall;
    function  Set_NbPostes(Value: Integer): HResult; stdcall;
    function  Get_SommeNomCommercial(out Value: WideString): HResult; stdcall;
    function  Get_SommeNoSerie(out Value: WideString): HResult; stdcall;
  end;

// *********************************************************************//
// DispIntf:  ICommandeMasterEvents
// Flags:     (4096) Dispatchable
// GUID:      {09579A3B-FD5F-4EA1-97B9-2B1EAE5D6657}
// *********************************************************************//
  ICommandeMasterEvents = dispinterface
    ['{09579A3B-FD5F-4EA1-97B9-2B1EAE5D6657}']
    procedure onErreur(ErreurCode: SYSINT; const ErreurMessage: WideString); dispid 1;
  end;

// *********************************************************************//
// La classe CoCommandeMaster fournit une méthode Create et CreateRemote pour
// créer des instances de l'interface par défaut ICommandeMaster exposée
// pas la CoClass CommandeMaster. Les fonctions sont destinées à être utilisées par
// les clients désirant automatiser les objets CoClass exposés par
// le serveur de cette bibliothèque de types.
// *********************************************************************//
  CoCommandeMaster = class
    class function Create: ICommandeMaster;
    class function CreateRemote(const MachineName: string): ICommandeMaster;
  end;

  TCommandeMasteronErreur = procedure(Sender: TObject; ErreurCode: SYSINT;
                                                       var ErreurMessage: OleVariant) of object;

// *********************************************************************//
// Déclaration de classe proxy de serveur OLE
// Objet serveur        : TCommandeMaster
// Chaîne d'aide        : CommandeMaster
// Interface par défaut : ICommandeMaster
// DISP Int. Déf. ?     : No
// Interface événements : ICommandeMasterEvents
// TypeFlags            : (2) CanCreate
// *********************************************************************//
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  TCommandeMasterProperties= class;
{$ENDIF}
  TCommandeMaster = class(TOleServer)
  private
    FOnonErreur: TCommandeMasteronErreur;
    FIntf:        ICommandeMaster;
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    FProps:       TCommandeMasterProperties;
    function      GetServerProperties: TCommandeMasterProperties;
{$ENDIF}
    function      GetDefaultInterface: ICommandeMaster;
    function  GetNomCommercial: widestring;
    procedure SetNomCommercial(const Value: widestring);
    function  GetNoSerie: WideString;
    procedure SetNoSerie(const Value: WideString);
    function  GetRaisonSociale: WideString;
    procedure SetRaisonSociale(const Value: WideString);
    function  getNom: WideString;
    procedure SetNom(const Value: WideString);
    function  GetPrenom: WideString;
    procedure SetPrenom(const Value: WideString);
    function  GetAdresse1: WideString;
    procedure SetAdresse1(const Value: WideString);
    function  GetAdresse2: WideString;
    procedure SetAdresse2(const Value: WideString);
    function  GetCodePostal: WideString;
    procedure SetCodePostal(const Value: WideString);
    function  GetVille: WideString;
    procedure SetVille(const Value: WideString);
    function  Getpays: WideString;
    procedure SetPays(const Value: WideString);
    function  GetTelephone: WideString;
    procedure SetTelephone(const Value: WideString);
    function  getFax: WideString;
    procedure SetFax(const Value: WideString);
    function  GetEmail: WideString;
    procedure SetEmail(const Value: WideString);
    function  GetSommeNom: WideString;
    function  GetSommeRaisonSociale: WideString;
    function  GetInstallMultiPoste: integer;
    function  GetNbPostes: integer;
    procedure SetNbPostes(const Value: integer);
    function  GetSommeClefCommande: WideString;
    function  GetClefCommande: WideString;
    function  GetDetailCommande: TListDetailCommande;
    function  GetCommande: Tcommande ;
    function GetSommeNoSerie: WideString;
    function GetSommeNomCommercial: WideString;
    function GetVersion: WideString;
  protected
    procedure InitServerData; override;
    procedure InvokeEvent(DispID: TDispID; var Params: TVariantArray); override;
    function  Get_NomCommercial(out Value: WideString): HResult;
    function  Set_NomCommercial(const Value: WideString): HResult;
    function  Get_NoSerie(out Value: WideString): HResult;
    function  Set_NoSerie(const Value: WideString): HResult;
    function  Get_RaisonSociale(out Value: WideString): HResult;
    function  Set_RaisonSociale(const Value: WideString): HResult;
    function  Get_Nom(out Value: WideString): HResult;
    function  Set_Nom(const Value: WideString): HResult;
    function  Get_Prenom(out Value: WideString): HResult;
    function  Set_Prenom(const Value: WideString): HResult;
    function  Get_Adresse1(out Value: WideString): HResult;
    function  Set_Adresse1(const Value: WideString): HResult;
    function  Get_Adresse2(out Value: WideString): HResult;
    function  Set_Adresse2(const Value: WideString): HResult;
    function  Get_CodePostal(out Value: WideString): HResult;
    function  Set_CodePostal(const Value: WideString): HResult;
    function  Get_Ville(out Value: WideString): HResult;
    function  Set_Ville(const Value: WideString): HResult;
    function  Get_Pays(out Value: WideString): HResult;
    function  Set_Pays(const Value: WideString): HResult;
    function  Get_Tel(out Value: WideString): HResult;
    function  Set_Tel(const Value: WideString): HResult;
    function  Get_Fax(out Value: WideString): HResult;
    function  Set_Fax(const Value: WideString): HResult;
    function  Get_Email(out Value: WideString): HResult;
    function  Set_Email(const Value: WideString): HResult;
    function  Get_RepertoireInfoMaster(out Value: WideString): HResult;
    function  Get_CommandeLibelle(out Value: WideString): HResult;
    function  Get_CommandePrix(out Value: Currency): HResult;
    function  Get_ListeCommandeLibelle(out Value: PSafeArray): HResult;
    function  Get_ListeCommandePrix(out Value: PSafeArray): HResult;
    function  Get_ClefCommande(out Value: WideString): HResult;
    function  Get_SommeRaisonSociale(out Value: WideString): HResult;
    function  Get_SommeNom(out Value: WideString): HResult;
    function  Get_SommeClefCommande(out Value: WideString): HResult;
    function  Get_Version(out Value: WideString): HResult;
    function  Get_FichierComposant(out Value: WideString): HResult;
    function  Get_InstalleMultiposte(out Value: Integer): HResult;
    function  Set_InstalleMultiposte(Value: Integer): HResult;
    function  Get_NbPostes(out Value: Integer): HResult;
    function  Set_NbPostes(Value: Integer): HResult;
    function  Get_SommeNomCommercial(out Value: WideString): HResult;
    function  Get_SommeNoSerie(out Value: WideString): HResult;
  public
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: ICommandeMaster);
    procedure Disconnect; override;
    function  InstallEdition(const RepInstall: WideString; const RepSource: WideString;
                             out ErreurCode: SYSINT; out ErreurMessage: WideString): HResult;
    function  GenereCommande(const RepInstall: WideString; out ErreurCode: SYSINT;
                             out ErreurMessage: WideString): HResult;
    function  InstallCommande(const RepSource: WideString; const RepDestination: WideString;
                              out ErreurCode: SYSINT; out ErreurMessage: WideString): HResult;
    property  DefaultInterface: ICommandeMaster read GetDefaultInterface;
    // MODIF LS
    property  NomCommercial : widestring read GetNomCommercial write SetNomCommercial;
    property  NoSerie : WideString read GetNoSerie Write SetNoSerie;
    property  RaisonSociale : WideString read GetRaisonSociale Write SetRaisonSociale;
    property  Nom : WideString read GetNom write SetNom;
    property  Prenom : WideString read GetPrenom Write SetPrenom;
    property  Adresse1 : WideString read GetAdresse1 Write SetAdresse1;
    property  Adresse2 : WideString read GetAdresse2 Write SetAdresse2;
    property  CodePostal : WideString read GetCodePostal Write SetCodePostal;
    property  Ville : WideString Read GetVille write SetVille;
    property  Pays : WideString read Getpays write SetPays;
    property  Telephone : WideString read GetTelephone write SetTelephone;
    property  Fax : WideString read getFax write SetFax;
    property  Email : WideString read GetEmail write SetEmail;
    property  SommeNom : WideString read GetSommeNom;
    property  SommeRaisonSociale : WideString read GetSommeRaisonSociale;
    property  SommeClefCommande : WideString read GetSommeClefCommande;
    property  InstalleMultiPoste : integer read GetInstallMultiPoste;
    property  NbPostes : integer read GetNbPostes write SetNbPostes;
    property  ClefCommande : WideString  read GetClefCommande;
    property  DetailCommande : TListDetailCommande read GetDetailCommande;
    property  InfoCommande : Tcommande read GetCommande;
    property  SommeNumeroSerie : WideString read GetSommeNoSerie;
    property  SommeNomCommercial : WideString read GetSommeNomCommercial;
    property  VersionDll : WideString read GetVersion;
  published
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    property Server: TCommandeMasterProperties read GetServerProperties;
{$ENDIF}
    property OnonErreur: TCommandeMasteronErreur read FOnonErreur write FOnonErreur;
  end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
// *********************************************************************//
// Classe proxy Propriétés Serveur OLE
// Objet serveur    : TCommandeMaster
// (Cet objet est utilisé pas l'inspecteur de propriété de l'EDI pour la
//  modification des propriétés de ce serveur)
// *********************************************************************//
 TCommandeMasterProperties = class(TPersistent)
  private
    FServer:    TCommandeMaster;
    function    GetDefaultInterface: ICommandeMaster;
    constructor Create(AServer: TCommandeMaster);
  protected
    function  Get_NomCommercial(out Value: WideString): HResult;
    function  Set_NomCommercial(const Value: WideString): HResult;
    function  Get_NoSerie(out Value: WideString): HResult;
    function  Set_NoSerie(const Value: WideString): HResult;
    function  Get_RaisonSociale(out Value: WideString): HResult;
    function  Set_RaisonSociale(const Value: WideString): HResult;
    function  Get_Nom(out Value: WideString): HResult;
    function  Set_Nom(const Value: WideString): HResult;
    function  Get_Prenom(out Value: WideString): HResult;
    function  Set_Prenom(const Value: WideString): HResult;
    function  Get_Adresse1(out Value: WideString): HResult;
    function  Set_Adresse1(const Value: WideString): HResult;
    function  Get_Adresse2(out Value: WideString): HResult;
    function  Set_Adresse2(const Value: WideString): HResult;
    function  Get_CodePostal(out Value: WideString): HResult;
    function  Set_CodePostal(const Value: WideString): HResult;
    function  Get_Ville(out Value: WideString): HResult;
    function  Set_Ville(const Value: WideString): HResult;
    function  Get_Pays(out Value: WideString): HResult;
    function  Set_Pays(const Value: WideString): HResult;
    function  Get_Tel(out Value: WideString): HResult;
    function  Set_Tel(const Value: WideString): HResult;
    function  Get_Fax(out Value: WideString): HResult;
    function  Set_Fax(const Value: WideString): HResult;
    function  Get_Email(out Value: WideString): HResult;
    function  Set_Email(const Value: WideString): HResult;
    function  Get_RepertoireInfoMaster(out Value: WideString): HResult;
    function  Get_CommandeLibelle(out Value: WideString): HResult;
    function  Get_CommandePrix(out Value: Currency): HResult;
    function  Get_ListeCommandeLibelle(out Value: PSafeArray): HResult;
    function  Get_ListeCommandePrix(out Value: PSafeArray): HResult;
    function  Get_ClefCommande(out Value: WideString): HResult;
    function  Get_SommeRaisonSociale(out Value: WideString): HResult;
    function  Get_SommeNom(out Value: WideString): HResult;
    function  Get_SommeClefCommande(out Value: WideString): HResult;
    function  Get_Version(out Value: WideString): HResult;
    function  Get_FichierComposant(out Value: WideString): HResult;
    function  Get_InstalleMultiposte(out Value: Integer): HResult;
    function  Set_InstalleMultiposte(Value: Integer): HResult;
    function  Get_NbPostes(out Value: Integer): HResult;
    function  Set_NbPostes(Value: Integer): HResult;
    function  Get_SommeNomCommercial(out Value: WideString): HResult;
    function  Get_SommeNoSerie(out Value: WideString): HResult;
  public
    property DefaultInterface: ICommandeMaster read GetDefaultInterface;
  published
  end;
{$ENDIF}


procedure Register;

implementation

uses ComObj;

class function CoCommandeMaster.Create: ICommandeMaster;
begin
  Result := CreateComObject(CLASS_CommandeMaster) as ICommandeMaster;
end;

class function CoCommandeMaster.CreateRemote(const MachineName: string): ICommandeMaster;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_CommandeMaster) as ICommandeMaster;
end;

procedure TCommandeMaster.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{ED55839D-D452-4214-80E0-8EAA702D760C}';
    IntfIID:   '{C82FE647-0682-4DB3-AFB6-406E4B9EB1C1}';
    EventIID:  '{09579A3B-FD5F-4EA1-97B9-2B1EAE5D6657}';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;

procedure TCommandeMaster.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk := GetServer;
    ConnectEvents(punk);
    Fintf:= punk as ICommandeMaster;
  end;
end;

procedure TCommandeMaster.ConnectTo(svrIntf: ICommandeMaster);
begin
  Disconnect;
  FIntf := svrIntf;
  ConnectEvents(FIntf);
end;

procedure TCommandeMaster.DisConnect;
begin
  if Fintf <> nil then
  begin
    DisconnectEvents(FIntf);
    FIntf := nil;
  end;
end;

function TCommandeMaster.GetDefaultInterface: ICommandeMaster;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface est NULL. Le composant n''est pas connecté au serveur. Vous devez appeler ''Connect'' ou ''ConnectTo'' avant cette opération');
  Result := FIntf;
end;

constructor TCommandeMaster.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps := TCommandeMasterProperties.Create(Self);
{$ENDIF}
end;

destructor TCommandeMaster.Destroy;
begin
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps.Free;
{$ENDIF}
  inherited Destroy;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
function TCommandeMaster.GetServerProperties: TCommandeMasterProperties;
begin
  Result := FProps;
end;
{$ENDIF}

procedure TCommandeMaster.InvokeEvent(DispID: TDispID; var Params: TVariantArray);
begin
  case DispID of
    -1: Exit;  // DISPID_UNKNOWN
   1: if Assigned(FOnonErreur) then
            FOnonErreur(Self, Params[0] {SYSINT}, Params[1] {const WideString});
  end; {case DispID}
end;

function  TCommandeMaster.Get_NomCommercial(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_NomCommercial(Value);
end;

function  TCommandeMaster.Set_NomCommercial(const Value: WideString): HResult;
begin
  Result := DefaultInterface.Set_NomCommercial(Value);
end;

function  TCommandeMaster.Get_NoSerie(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_NoSerie(Value);
end;

function  TCommandeMaster.Set_NoSerie(const Value: WideString): HResult;
begin
  Result := DefaultInterface.Set_NoSerie(Value);
end;

function  TCommandeMaster.Get_RaisonSociale(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_RaisonSociale(Value);
end;

function  TCommandeMaster.Set_RaisonSociale(const Value: WideString): HResult;
begin
  Result := DefaultInterface.Set_RaisonSociale(Value);
end;

function  TCommandeMaster.Get_Nom(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_Nom(Value);
end;

function  TCommandeMaster.Set_Nom(const Value: WideString): HResult;
begin
  Result := DefaultInterface.Set_Nom(Value);
end;

function  TCommandeMaster.Get_Prenom(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_Prenom(Value);
end;

function  TCommandeMaster.Set_Prenom(const Value: WideString): HResult;
begin
  Result := DefaultInterface.Set_Prenom(Value);
end;

function  TCommandeMaster.Get_Adresse1(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_Adresse1(Value);
end;

function  TCommandeMaster.Set_Adresse1(const Value: WideString): HResult;
begin
  Result := DefaultInterface.Set_Adresse1(Value);
end;

function  TCommandeMaster.Get_Adresse2(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_Adresse2(Value);
end;

function  TCommandeMaster.Set_Adresse2(const Value: WideString): HResult;
begin
  Result := DefaultInterface.Set_Adresse2(Value);
end;

function  TCommandeMaster.Get_CodePostal(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_CodePostal(Value);
end;

function  TCommandeMaster.Set_CodePostal(const Value: WideString): HResult;
begin
  Result := DefaultInterface.Set_CodePostal(Value);
end;

function  TCommandeMaster.Get_Ville(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_Ville(Value);
end;

function  TCommandeMaster.Set_Ville(const Value: WideString): HResult;
begin
  Result := DefaultInterface.Set_Ville(Value);
end;

function  TCommandeMaster.Get_Pays(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_Pays(Value);
end;

function  TCommandeMaster.Set_Pays(const Value: WideString): HResult;
begin
  Result := DefaultInterface.Set_Pays(Value);
end;

function  TCommandeMaster.Get_Tel(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_Tel(Value);
end;

function  TCommandeMaster.Set_Tel(const Value: WideString): HResult;
begin
  Result := DefaultInterface.Set_Tel(Value);
end;

function  TCommandeMaster.Get_Fax(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_Fax(Value);
end;

function  TCommandeMaster.Set_Fax(const Value: WideString): HResult;
begin
  Result := DefaultInterface.Set_Fax(Value);
end;

function  TCommandeMaster.Get_Email(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_Email(Value);
end;

function  TCommandeMaster.Set_Email(const Value: WideString): HResult;
begin
  Result := DefaultInterface.Set_Email(Value);
end;

function  TCommandeMaster.Get_RepertoireInfoMaster(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_RepertoireInfoMaster(Value);
end;

function  TCommandeMaster.Get_CommandeLibelle(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_CommandeLibelle(Value);
end;

function  TCommandeMaster.Get_CommandePrix(out Value: Currency): HResult;
begin
  Result := DefaultInterface.Get_CommandePrix(Value);
end;

function  TCommandeMaster.Get_ListeCommandeLibelle(out Value: PSafeArray): HResult;
begin
  Result := DefaultInterface.Get_ListeCommandeLibelle(Value);
end;

function  TCommandeMaster.Get_ListeCommandePrix(out Value: PSafeArray): HResult;
begin
  Result := DefaultInterface.Get_ListeCommandePrix(Value);
end;

function  TCommandeMaster.Get_ClefCommande(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_ClefCommande(Value);
end;

function  TCommandeMaster.Get_SommeRaisonSociale(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_SommeRaisonSociale(Value);
end;

function  TCommandeMaster.Get_SommeNom(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_SommeNom(Value);
end;

function  TCommandeMaster.Get_SommeClefCommande(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_SommeClefCommande(Value);
end;

function  TCommandeMaster.Get_Version(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_Version(Value);
end;

function  TCommandeMaster.Get_FichierComposant(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_FichierComposant(Value);
end;

function  TCommandeMaster.Get_InstalleMultiposte(out Value: Integer): HResult;
begin
  Result := DefaultInterface.Get_InstalleMultiposte(Value);
end;

function  TCommandeMaster.Set_InstalleMultiposte(Value: Integer): HResult;
begin
  Result := DefaultInterface.Set_InstalleMultiposte(Value);
end;

function  TCommandeMaster.Get_NbPostes(out Value: Integer): HResult;
begin
  Result := DefaultInterface.Get_NbPostes(Value);
end;

function  TCommandeMaster.Set_NbPostes(Value: Integer): HResult;
begin
  Result := DefaultInterface.Set_NbPostes(Value);
end;

function  TCommandeMaster.Get_SommeNomCommercial(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_SommeNomCommercial(Value);
end;

function  TCommandeMaster.Get_SommeNoSerie(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_SommeNoSerie(Value);
end;

function  TCommandeMaster.InstallEdition(const RepInstall: WideString; const RepSource: WideString; 
                                         out ErreurCode: SYSINT; out ErreurMessage: WideString): HResult;
begin
  Result := DefaultInterface.InstallEdition(RepInstall, RepSource, ErreurCode, ErreurMessage);
end;

function  TCommandeMaster.GenereCommande(const RepInstall: WideString; out ErreurCode: SYSINT; 
                                         out ErreurMessage: WideString): HResult;
begin
  Result := DefaultInterface.GenereCommande(RepInstall, ErreurCode, ErreurMessage);
end;

function  TCommandeMaster.InstallCommande(const RepSource: WideString; 
                                          const RepDestination: WideString; out ErreurCode: SYSINT; 
                                          out ErreurMessage: WideString): HResult;
begin
  Result := DefaultInterface.InstallCommande(RepSource, RepDestination, ErreurCode, ErreurMessage);
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
constructor TCommandeMasterProperties.Create(AServer: TCommandeMaster);
begin
  inherited Create;
  FServer := AServer;
end;

function TCommandeMasterProperties.GetDefaultInterface: ICommandeMaster;
begin
  Result := FServer.DefaultInterface;
end;

function  TCommandeMasterProperties.Get_NomCommercial(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_NomCommercial(Value);
end;

function  TCommandeMasterProperties.Set_NomCommercial(const Value: WideString): HResult;
begin
  Result := DefaultInterface.Set_NomCommercial(Value);
end;

function  TCommandeMasterProperties.Get_NoSerie(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_NoSerie(Value);
end;

function  TCommandeMasterProperties.Set_NoSerie(const Value: WideString): HResult;
begin
  Result := DefaultInterface.Set_NoSerie(Value);
end;

function  TCommandeMasterProperties.Get_RaisonSociale(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_RaisonSociale(Value);
end;

function  TCommandeMasterProperties.Set_RaisonSociale(const Value: WideString): HResult;
begin
  Result := DefaultInterface.Set_RaisonSociale(Value);
end;

function  TCommandeMasterProperties.Get_Nom(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_Nom(Value);
end;

function  TCommandeMasterProperties.Set_Nom(const Value: WideString): HResult;
begin
  Result := DefaultInterface.Set_Nom(Value);
end;

function  TCommandeMasterProperties.Get_Prenom(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_Prenom(Value);
end;

function  TCommandeMasterProperties.Set_Prenom(const Value: WideString): HResult;
begin
  Result := DefaultInterface.Set_Prenom(Value);
end;

function  TCommandeMasterProperties.Get_Adresse1(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_Adresse1(Value);
end;

function  TCommandeMasterProperties.Set_Adresse1(const Value: WideString): HResult;
begin
  Result := DefaultInterface.Set_Adresse1(Value);
end;

function  TCommandeMasterProperties.Get_Adresse2(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_Adresse2(Value);
end;

function  TCommandeMasterProperties.Set_Adresse2(const Value: WideString): HResult;
begin
  Result := DefaultInterface.Set_Adresse2(Value);
end;

function  TCommandeMasterProperties.Get_CodePostal(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_CodePostal(Value);
end;

function  TCommandeMasterProperties.Set_CodePostal(const Value: WideString): HResult;
begin
  Result := DefaultInterface.Set_CodePostal(Value);
end;

function  TCommandeMasterProperties.Get_Ville(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_Ville(Value);
end;

function  TCommandeMasterProperties.Set_Ville(const Value: WideString): HResult;
begin
  Result := DefaultInterface.Set_Ville(Value);
end;

function  TCommandeMasterProperties.Get_Pays(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_Pays(Value);
end;

function  TCommandeMasterProperties.Set_Pays(const Value: WideString): HResult;
begin
  Result := DefaultInterface.Set_Pays(Value);
end;

function  TCommandeMasterProperties.Get_Tel(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_Tel(Value);
end;

function  TCommandeMasterProperties.Set_Tel(const Value: WideString): HResult;
begin
  Result := DefaultInterface.Set_Tel(Value);
end;

function  TCommandeMasterProperties.Get_Fax(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_Fax(Value);
end;

function  TCommandeMasterProperties.Set_Fax(const Value: WideString): HResult;
begin
  Result := DefaultInterface.Set_Fax(Value);
end;

function  TCommandeMasterProperties.Get_Email(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_Email(Value);
end;

function  TCommandeMasterProperties.Set_Email(const Value: WideString): HResult;
begin
  Result := DefaultInterface.Set_Email(Value);
end;

function  TCommandeMasterProperties.Get_RepertoireInfoMaster(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_RepertoireInfoMaster(Value);
end;

function  TCommandeMasterProperties.Get_CommandeLibelle(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_CommandeLibelle(Value);
end;

function  TCommandeMasterProperties.Get_CommandePrix(out Value: Currency): HResult;
begin
  Result := DefaultInterface.Get_CommandePrix(Value);
end;

function  TCommandeMasterProperties.Get_ListeCommandeLibelle(out Value: PSafeArray): HResult;
begin
  Result := DefaultInterface.Get_ListeCommandeLibelle(Value);
end;

function  TCommandeMasterProperties.Get_ListeCommandePrix(out Value: PSafeArray): HResult;
begin
  Result := DefaultInterface.Get_ListeCommandePrix(Value);
end;

function  TCommandeMasterProperties.Get_ClefCommande(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_ClefCommande(Value);
end;

function  TCommandeMasterProperties.Get_SommeRaisonSociale(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_SommeRaisonSociale(Value);
end;

function  TCommandeMasterProperties.Get_SommeNom(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_SommeNom(Value);
end;

function  TCommandeMasterProperties.Get_SommeClefCommande(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_SommeClefCommande(Value);
end;

function  TCommandeMasterProperties.Get_Version(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_Version(Value);
end;

function  TCommandeMasterProperties.Get_FichierComposant(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_FichierComposant(Value);
end;

function  TCommandeMasterProperties.Get_InstalleMultiposte(out Value: Integer): HResult;
begin
  Result := DefaultInterface.Get_InstalleMultiposte(Value);
end;

function  TCommandeMasterProperties.Set_InstalleMultiposte(Value: Integer): HResult;
begin
  Result := DefaultInterface.Set_InstalleMultiposte(Value);
end;

function  TCommandeMasterProperties.Get_NbPostes(out Value: Integer): HResult;
begin
  Result := DefaultInterface.Get_NbPostes(Value);
end;

function  TCommandeMasterProperties.Set_NbPostes(Value: Integer): HResult;
begin
  Result := DefaultInterface.Set_NbPostes(Value);
end;

function  TCommandeMasterProperties.Get_SommeNomCommercial(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_SommeNomCommercial(Value);
end;

function  TCommandeMasterProperties.Get_SommeNoSerie(out Value: WideString): HResult;
begin
  Result := DefaultInterface.Get_SommeNoSerie(Value);
end;

{$ENDIF}

procedure Register;
begin
  RegisterComponents('Batiprix',[TCommandeMaster]);
end;

function TCommandeMaster.GetNomCommercial: widestring;
var valeur : WideString;
begin
  result := '';
  if Get_NomCommercial(valeur)<> 0 then exit;
  result := Valeur;
end;

procedure TCommandeMaster.SetNomCommercial(const Value: widestring);
begin
  Set_NomCommercial(Value);
end;

function TCommandeMaster.GetNoSerie: WideString;
var valeur : WideString;
begin
  result := '';
  if Get_NoSerie(Valeur) <> 0 then exit;
  result := valeur;
end;

procedure TCommandeMaster.SetNoSerie(const Value: WideString);
begin
  Set_NoSerie(Value);
end;

function TCommandeMaster.GetRaisonSociale: WideString;
var Valeur : WideString;
begin
  result := '';
  if Get_RaisonSociale(valeur) <> 0 then exit;
  result := Valeur;
end;

procedure TCommandeMaster.SetRaisonSociale(const Value: WideString);
begin
  Set_RaisonSociale(Value);
end;

function TCommandeMaster.getNom: WideString;
var  Valeur : WideString;
begin
  result := '';
  if Get_Nom(valeur) <> 0 then exit;
  result:= Valeur;
end;

procedure TCommandeMaster.SetNom(const Value: WideString);
begin
  Set_Nom(Value);
end;

function TCommandeMaster.GetPrenom: WideString;
var Valeur : WideString;
begin
  result := '';
  if Get_Prenom(Valeur) <> 0 then exit;
  result := valeur;
end;

procedure TCommandeMaster.SetPrenom(const Value: WideString);
begin
  Set_Prenom(Value);
end;

function TCommandeMaster.GetAdresse1: WideString;
var Valeur : WideString;
begin
  result := '';
  if Get_Adresse1(valeur) <> 0 then exit;
  result := valeur;
end;

procedure TCommandeMaster.SetAdresse1(const Value: WideString);
begin
  Set_Adresse1(Value);
end;

function TCommandeMaster.GetAdresse2: WideString;
var Valeur : WideString;
begin
  result := '';
  if Get_Adresse2(valeur) <> 0 then exit;
  result := valeur;
end;

procedure TCommandeMaster.SetAdresse2(const Value: WideString);
begin
  Set_Adresse2(Value);
end;

function TCommandeMaster.GetCodePostal: WideString;
var  Valeur : WideString;
begin
  result := '';
  if Get_CodePostal(valeur) <> 0 then exit;
  result := valeur;
end;

procedure TCommandeMaster.SetCodePostal(const Value: WideString);
begin
  Set_CodePostal(Value);
end;

function TCommandeMaster.GetVille: WideString;
var  Valeur : WideString;
begin
  result := '';
  if Get_Ville(Valeur) <> 0 then exit;
  result := valeur;
end;

procedure TCommandeMaster.SetVille(const Value: WideString);
begin
  Set_Ville(Value);
end;

function TCommandeMaster.Getpays: WideString;
var Valeur : WideString;
begin
  result := '';
  if Get_Pays(Valeur) <> 0 then exit;
  result := Valeur;
end;

procedure TCommandeMaster.SetPays(const Value: WideString);
begin
  Set_Pays(Value);
end;

function TCommandeMaster.GetTelephone: WideString;
var  Valeur : WideString;
begin
  result := '';
  if Get_Tel(Valeur) <> 0 then exit;
  result := valeur;
end;

procedure TCommandeMaster.SetTelephone(const Value: WideString);
begin
  Set_Tel(Value);
end;

function TCommandeMaster.getFax: WideString;
var Valeur : WideString;
begin
  result :='';
  if Get_Fax(valeur) <> 0 then exit;
  result := Valeur;
end;

procedure TCommandeMaster.SetFax(const Value: WideString);
begin
  Set_Fax(Value);
end;

function TCommandeMaster.GetEmail: WideString;
var  Valeur : WideString;
begin
  result := '';
  if Get_Email(Valeur) <> 0 then exit;
  result := Valeur;
end;

procedure TCommandeMaster.SetEmail(const Value: WideString);
begin
  Set_Email(Value);
end;

function TCommandeMaster.GetSommeNom: WideString;
var Valeur : WideString;
begin
  result := '';
  if Get_SommeNom(Valeur) <> 0 then exit;
  result := Valeur;
end;

function TCommandeMaster.GetSommeRaisonSociale: WideString;
var  Valeur : WideString;
begin
  result := '';
  if Get_SommeRaisonSociale(Valeur) <> 0 then exit;
  result := Valeur;
end;

function TCommandeMaster.GetInstallMultiPoste: integer;
var Valeur : integer;
begin
  result := 1;
  IF Get_InstalleMultiposte(valeur) <> 0 THEN exit;
  result := Valeur;
end;

function TCommandeMaster.getNbPostes: integer;
var Valeur : integer;
begin
  result := 1;
  if Get_NbPostes(valeur) <> 0 then exit;
  result := valeur;
end;

procedure TCommandeMaster.SetNbPostes(const Value: integer);
begin
  Set_NbPostes(Value);
end;

function TCommandeMaster.GetSommeClefCommande: WideString;
var Valeur : WideString;
begin
  result := '';
  if Get_SommeClefCommande(valeur) <> 0 then exit;
  result := valeur;
end;

function TCommandeMaster.GetClefCommande: WideString;
var Valeur : WideString;
begin
  result := '';
  if Get_ClefCommande(valeur) <> 0 then exit;
  result := Valeur;
end;

function TCommandeMaster.GetDetailCommande: TListDetailCommande;
var
    TLibelle,Tvaleur : PSafeArray;
    Min,Max,Indice : integer;
    TheListDetailCmd : TlistDetailCommande;
    UneLigneCommande :TdetailCommande;
    Libelle : WideString;
    valeur : Currency;
    indiceT : integer;
begin
  TheListDetailCmd := TListDetailCommande.create;
  result := TheListDetailCmd;
  IndiceT := 1;
  //
  if Get_ListeCommandeLibelle(Tlibelle) <> 0 then exit;
  if Get_ListeCommandePrix(Tvaleur) <> 0 then exit;
  //
  if SafeArrayGetUBound (Tvaleur,IndiceT,Max) <> 0 then exit;
  if SafeArrayGetlBound (Tvaleur,IndiceT,Min) <> 0 then exit;

  for Indice := min to max do
  begin
    if SafeArrayGetElement (TLibelle,Indice,Libelle) <> 0 then break;
    if SafeArrayGetElement (TValeur,Indice,valeur) <> 0 then break;
    UneLigneCommande := TdetailCommande.create;
    UneLigneCommande.Libelle := Libelle;
    UneLigneCommande.prix  := Valeur ;
    TheListDetailCmd.Add (UneLigneCommande);
  end;
end;

function TCommandeMaster.GetCommande: Tcommande;
var LaCommande : Tcommande;
    Libelle : WideString;
    Prix : Currency;
begin
  result := nil;
  if Get_CommandeLibelle(Libelle)<> 0 then exit;
  if Get_CommandePrix(prix)<>0 then exit;
  LaCommande := TCommande.create;
  laCOmmande.Libelle := Libelle;
  LaCommande.Prix := Prix;
  result := LaCommande;
end;


{ TListDetailCommande }

function TListDetailCommande.Add(AObject: TDetailCommande): Integer;
begin
  result := Inherited Add(AObject);
end;

procedure TListDetailCommande.clear;
var Indice : integer;
begin
  for Indice := 0 to Count -1 do
  begin
    TDetailCommande(Items [Indice]).free;
  end;
  inherited;
end;

destructor TListDetailCommande.destroy;
begin
  clear;
  inherited;
end;

function TListDetailCommande.GetItems(Indice: integer): TDetailCommande;
begin
  result := TDetailCommande (Inherited Items[Indice]);
end;

procedure TListDetailCommande.SetItems(Indice: integer; const Value: TDetailCommande);
begin
  Inherited Items[Indice]:= Value;
end;

function TCommandeMaster.GetSommeNoSerie: WideString;
var Valeur : WideString;
begin
  result := '';
  if Get_SommeNoSerie (valeur) <> 0 then exit;
  result := Valeur;
end;

function TCommandeMaster.GetSommeNomCommercial: WideString;
var Valeur : WideString;
begin
  result := '';
  if Get_SommeNomCommercial (valeur) <> 0 then exit;
  result := Valeur;
end;

function TCommandeMaster.GetVersion: WideString;
var Valeur : WideString;
begin
  result := '';
  if Get_Version (valeur) <> 0 then exit;
  result := Valeur;
end;

end.
