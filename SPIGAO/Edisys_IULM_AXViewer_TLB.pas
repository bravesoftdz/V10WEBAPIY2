unit Edisys_IULM_AXViewer_TLB;

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
// Fichier généré le 27/06/2012 15:25:09 depuis la bibliothèque de types ci-dessous.

// ************************************************************************  //
// Bibl. types : C:\Program Files (x86)\Edisys\IULM\x86\Edisys.IULM.AXViewer.tlb (1)
// LIBID: {E4442298-F916-4372-804D-B3B01B261217}
// LCID: 0
// Fichier d'aide : 
// Chaîne d'aide : 
// DepndLst: 
//   (1) v2.0 mscorlib, (C:\Windows\Microsoft.NET\Framework\v2.0.50727\mscorlib.tlb)
//   (2) v2.0 System, (C:\Windows\Microsoft.NET\Framework\v2.0.50727\System.tlb)
//   (3) v2.0 System_Windows_Forms, (C:\Windows\Microsoft.NET\Framework\v2.0.50727\System.Windows.Forms.tlb)
//   (4) v2.2 Edisys_IULM_Core, (C:\Program Files (x86)\Edisys\IULM\x86\Edisys.IULM.Core.tlb)
//   (5) v2.0 stdole, (C:\Windows\SysWOW64\stdole2.tlb)
// Erreurs :
//   Erreur à la création du bitmap de palette de (TDocumentViewer) : Clé de registre CLSID\{713C7960-0964-4D42-B771-D1AF99C0EBC7}\ToolboxBitmap32 non trouvée
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

uses Windows, ActiveX, Classes, Edisys_IULM_Core_TLB, Graphics, mscorlib_TLB, OleCtrls, OleServer, 
StdVCL, System_TLB, System_Windows_Forms_TLB, Variants;
  

// *********************************************************************//
// GUIDS déclarés dans la bibliothèque de types. Préfixes utilisés :    
//   Bibliothèques de types : LIBID_xxxx                                      
//   CoClasses              : CLASS_xxxx                                      
//   DISPInterfaces         : DIID_xxxx                                       
//   Non-DISP interfaces    : IID_xxxx                                        
// *********************************************************************//
const
  // Versions majeure et mineure de la bibliothèque de types
  Edisys_IULM_AXViewerMajorVersion = 2;
  Edisys_IULM_AXViewerMinorVersion = 2;

  LIBID_Edisys_IULM_AXViewer: TGUID = '{E4442298-F916-4372-804D-B3B01B261217}';

  CLASS_DocumentViewer: TGUID = '{713C7960-0964-4D42-B771-D1AF99C0EBC7}';
type

// *********************************************************************//
// Déclaration de CoClasses définies dans la bibliothèque de types 
// (REMARQUE: On affecte chaque CoClasse à son Interface par défaut)              
// *********************************************************************//
  DocumentViewer = IDocumentViewer;



// *********************************************************************//
// Déclaration de classe proxy de contrôle OLE
// Nom du contrôle      : TDocumentViewer
// Chaîne d'aide        : 
// Interface par défaut : IDocumentViewer
// DISP Int. Déf. ?     : Yes
// Interface événements : 
// TypeFlags            : (2) CanCreate
// *********************************************************************//
  TDocumentViewer = class(TOleControl)
  private
    FIntf: IDocumentViewer;
    function  GetControlInterface: IDocumentViewer;
  protected
    procedure CreateControl;
    procedure InitControlData; override;
  public
    procedure Show(const Project: IProject);
    function NavigateTo(const Project: IProject; const workItemId: WideString; 
                        documentType: ComDocumentType): WordBool;
    property  ControlInterface: IDocumentViewer read GetControlInterface;
    property  DefaultInterface: IDocumentViewer read GetControlInterface;
  published
    property Anchors;
  end;

procedure Register;

resourcestring
  dtlServerPage = 'ActiveX';

  dtlOcxPage = 'ActiveX';

implementation

uses ComObj;

procedure TDocumentViewer.InitControlData;
const
  CControlData: TControlData2 = (
    ClassID: '{713C7960-0964-4D42-B771-D1AF99C0EBC7}';
    EventIID: '';
    EventCount: 0;
    EventDispIDs: nil;
    LicenseKey: nil (*HR:$80004002*);
    Flags: $00000000;
    Version: 401);
begin
  ControlData := @CControlData;
end;

procedure TDocumentViewer.CreateControl;

  procedure DoCreate;
  begin
    FIntf := IUnknown(OleObject) as IDocumentViewer;
  end;

begin
  if FIntf = nil then DoCreate;
end;

function TDocumentViewer.GetControlInterface: IDocumentViewer;
begin
  CreateControl;
  Result := FIntf;
end;

procedure TDocumentViewer.Show(const Project: IProject);
begin
  DefaultInterface.Show(Project);
end;

function TDocumentViewer.NavigateTo(const Project: IProject; const workItemId: WideString; 
                                    documentType: ComDocumentType): WordBool;
begin
  Result := DefaultInterface.NavigateTo(Project, workItemId, documentType);
end;

procedure Register;
begin
  RegisterComponents(dtlOcxPage, [TDocumentViewer]);
end;

end.
