unit ETOBACTIVE_TLB;

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
// Fichier généré le 31/07/2001 16:49:35 depuis la bibliothèque de types ci-dessous.

// ************************************************************************ //
// Bibl.Types     : Y:\PgiS5\VDEV\Gescom\LibEC\ETOBACTIVE.tlb (1)
// IID\LCID       : {E86A1F53-9EB6-4D78-8463-A42EB3D22C16}\0
// Fichier d'aide : 
// DepndLst       : 
//   (1) v2.0 stdole, (C:\WINDOWS\System32\stdole2.tlb)
//   (2) v4.0 StdVCL, (C:\WINDOWS\System32\STDVCL40.DLL)
// ************************************************************************ //
{$TYPEDADDRESS OFF} // L'unité doit être compilée sans vérification de type des pointeurs. 
interface

uses Windows, ActiveX, Classes, Graphics, OleServer, OleCtrls, StdVCL;

// *********************************************************************//
// GUIDS déclarés dans la bibliothèque de types. Préfixes utilisés :    
//   Bibliothèques de types : LIBID_xxxx                                      
//   CoClasses              : CLASS_xxxx                                      
//   DISPInterfaces         : DIID_xxxx                                       
//   Non-DISP interfaces    : IID_xxxx                                        
// *********************************************************************//
const
  // Version majeure et mineure de la bibliothèque de types
  ETOBACTIVEMajorVersion = 1;
  ETOBACTIVEMinorVersion = 0;

  LIBID_ETOBACTIVE: TGUID = '{E86A1F53-9EB6-4D78-8463-A42EB3D22C16}';

  IID_ITOBActive: TGUID = '{C66A60D3-522F-46D2-864D-52DCFD9C3B8A}';
  CLASS_TOBActive: TGUID = '{F57812CB-8A8A-485E-B8EE-E9C274A79D25}';
type

// *********************************************************************//
// Déclaration Forward des types définis dans la bibliothèque de types    
// *********************************************************************//
  ITOBActive = interface;
  ITOBActiveDisp = dispinterface;

// *********************************************************************//
// Déclaration de CoClasses définies dans la bibliothèque de types 
// (REMARQUE: On affecte chaque CoClass à son Interface par défaut)              
// *********************************************************************//
  TOBActive = ITOBActive;


// *********************************************************************//
// Interface   : ITOBActive
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID        : {C66A60D3-522F-46D2-864D-52DCFD9C3B8A}
// *********************************************************************//
  ITOBActive = interface(IDispatch)
    ['{C66A60D3-522F-46D2-864D-52DCFD9C3B8A}']
    procedure OnStartPage(const AScriptingContext: IUnknown); safecall;
    procedure OnEndPage; safecall;
    function  EGETTARIF(FichierIni: OleVariant; NomSoc: OleVariant; CodeArt: OleVariant; 
                        TarifArticle: OleVariant; CodeTiers: OleVariant; TarifTiers: OleVariant; 
                        CodeDevise: OleVariant; CodeDepot: OleVariant; LaDate: OleVariant; 
                        Qte: OleVariant; RemiseTiers: OleVariant; var LePrix: OleVariant; 
                        var LaRemise: OleVariant): OleVariant; safecall;
    function  EGETTAUXTVA(FichierIni: OleVariant; NomSoc: OleVariant; Regime: OleVariant; 
                          Code: OleVariant): OleVariant; safecall;
    procedure EGetDBParams(FichierIni: OleVariant; NomSoc: OleVariant; var StDriver: OleVariant; 
                           var StServer: OleVariant; var StPath: OleVariant; 
                           var StODBC: OleVariant; var StBase: OleVariant; var StUser: OleVariant; 
                           var StPassWord: OleVariant; var StOptions: OleVariant); safecall;
    function  EGETTARIF_WP(CodeArt: OleVariant; TarifArticle: OleVariant; CodeTiers: OleVariant; 
                           TarifTiers: OleVariant; CodeDevise: OleVariant; CodeDepot: OleVariant; 
                           LaDate: OleVariant; Qte: OleVariant; RemiseTiers: OleVariant; 
                           var LePrix: OleVariant; var LaRemise: OleVariant; StDriver: OleVariant; 
                           StServer: OleVariant; StPath: OleVariant; StODBC: OleVariant; 
                           StBase: OleVariant; StUser: OleVariant; StPassWord: OleVariant; 
                           StOptions: OleVariant): OleVariant; safecall;
    procedure EGetInfosDossier(FichierIni: OleVariant; NomSoc: OleVariant; 
                               var DevisePivot: OleVariant; var TenueEuro: OleVariant; 
                               var DepotDefaut: OleVariant); safecall;
  end;

// *********************************************************************//
// DispIntf:  ITOBActiveDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {C66A60D3-522F-46D2-864D-52DCFD9C3B8A}
// *********************************************************************//
  ITOBActiveDisp = dispinterface
    ['{C66A60D3-522F-46D2-864D-52DCFD9C3B8A}']
    procedure OnStartPage(const AScriptingContext: IUnknown); dispid 1;
    procedure OnEndPage; dispid 2;
    function  EGETTARIF(FichierIni: OleVariant; NomSoc: OleVariant; CodeArt: OleVariant; 
                        TarifArticle: OleVariant; CodeTiers: OleVariant; TarifTiers: OleVariant; 
                        CodeDevise: OleVariant; CodeDepot: OleVariant; LaDate: OleVariant; 
                        Qte: OleVariant; RemiseTiers: OleVariant; var LePrix: OleVariant; 
                        var LaRemise: OleVariant): OleVariant; dispid 8;
    function  EGETTAUXTVA(FichierIni: OleVariant; NomSoc: OleVariant; Regime: OleVariant; 
                          Code: OleVariant): OleVariant; dispid 9;
    procedure EGetDBParams(FichierIni: OleVariant; NomSoc: OleVariant; var StDriver: OleVariant; 
                           var StServer: OleVariant; var StPath: OleVariant; 
                           var StODBC: OleVariant; var StBase: OleVariant; var StUser: OleVariant; 
                           var StPassWord: OleVariant; var StOptions: OleVariant); dispid 3;
    function  EGETTARIF_WP(CodeArt: OleVariant; TarifArticle: OleVariant; CodeTiers: OleVariant; 
                           TarifTiers: OleVariant; CodeDevise: OleVariant; CodeDepot: OleVariant; 
                           LaDate: OleVariant; Qte: OleVariant; RemiseTiers: OleVariant; 
                           var LePrix: OleVariant; var LaRemise: OleVariant; StDriver: OleVariant; 
                           StServer: OleVariant; StPath: OleVariant; StODBC: OleVariant; 
                           StBase: OleVariant; StUser: OleVariant; StPassWord: OleVariant; 
                           StOptions: OleVariant): OleVariant; dispid 4;
    procedure EGetInfosDossier(FichierIni: OleVariant; NomSoc: OleVariant; 
                               var DevisePivot: OleVariant; var TenueEuro: OleVariant; 
                               var DepotDefaut: OleVariant); dispid 5;
  end;

// *********************************************************************//
// La classe CoTOBActive fournit une méthode Create et CreateRemote pour          
// créer des instances de l'interface par défaut ITOBActive exposée             
// pas la CoClass TOBActive. Les fonctions sont destinées à être utilisées par            
// les clients désirant automatiser les objets CoClass exposés par       
// le serveur de cette bibliothèque de types.                                            
// *********************************************************************//
  CoTOBActive = class
    class function Create: ITOBActive;
    class function CreateRemote(const MachineName: string): ITOBActive;
  end;

implementation

uses ComObj;

class function CoTOBActive.Create: ITOBActive;
begin
  Result := CreateComObject(CLASS_TOBActive) as ITOBActive;
end;

class function CoTOBActive.CreateRemote(const MachineName: string): ITOBActive;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_TOBActive) as ITOBActive;
end;

end.
