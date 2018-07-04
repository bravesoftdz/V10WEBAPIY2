unit KB_Param;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, UiUtil,
  HSysMenu, HColComb, StdCtrls, Hctrls, ExtCtrls, HPanel, HTB97, Mask, Utob, HEnt1,
  KB_Ecran, HMsgBox, ComCtrls, Math, Spin, M3FP, HColor, Buttons, Grids,
  Hgauge,
{$IFDEF EAGLCLIENT}                    //MR debut 23/10/02
  MaineAgl,
{$ELSE}
  dbTables, FE_Main, AssistSQL,
{$ENDIF}                               //MR fin 23/10/02
{$IFDEF PGIS5}
  MC_Lib, FOUtil, FODefi, TickUtilFO ;  //NA 17/11/03
{$ELSE}
  Ut, uterreur2, utpv, S1Util, utSaisie ;
{$ENDIF}

Procedure Param_ClavierEcran ( Caisse : String ; var Params : String ; action : TActionFiche ) ;

type
  TFParamKB = class(TForm)
    BottomDock: TDock97;
    BarButtons: TToolWindow97;
    PMain: THPanel;
    HMTrad: THSystemMenu;
    PnlBtns: THPanel;
    BAide: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BValider: TToolbarButton97;
    BRecupere: TToolbarButton97;
    TimerMax: TTimer;
    PnlAvis: THPanel;
    pnlParam: THPanel;
    pages: TPageControl;
    tsArticle: TTabSheet;
    grArticle: TGroupBox;
    tArtCode: THLabel;
    TCE_QTE: THLabel;
    TCE_PRIX: THLabel;
    ArtLabel: THLabel;
    ArtCode: THCritMaskEdit;
    ArtQte: THNumEdit;
    ArtPrix: THNumEdit;
    tsCommentaire: TTabSheet;
    grCommentaire: TGroupBox;
    tcomCode: THLabel;
    ComLabel: THLabel;
    ComCode: THCritMaskEdit;
    tsRemise: TTabSheet;
    grRemise: TGroupBox;
    tremmotif: THLabel;
    tRempourcent: THLabel;
    rempourcent: THNumEdit;
    tsVendeur: TTabSheet;
    tsFonction: TTabSheet;
    grFonction: TGroupBox;
    tFonCode: THLabel;
    FonCode: THValComboBox;
    tsReglement: TTabSheet;
    grReglement: TGroupBox;
    tregCode: THLabel;
    RegCode: THValComboBox;
    tsOCT: TTabSheet;
    grOCT: TGroupBox;
    lbloctcode: THLabel;
    lbloctprix: THLabel;
    octCode: THValComboBox;
    OCTPrix: THNumEdit;
    tsOFC: TTabSheet;
    grOFC: TGroupBox;
    lblOFCCode: THLabel;
    lblOFCPrix: THLabel;
    OFCCode: THValComboBox;
    OFCPrix: THNumEdit;
    grCouleur: TGroupBox;
    rbNonCouleur: TRadioButton;
    RbCouleur: TRadioButton;
    FCouleur: THComboColor;
    grRacourci: TGroupBox;
    TCE_TOUCHE: THLabel;
    FTouche: THValComboBox;
    cbCtrl: TCheckBox;
    cbShift: TCheckBox;
    CE_CAISSE: TEdit;
    CE_NUMERO: TEdit;
    grImage: TGroupBox;
    FImageLibre: THCritMaskEdit;
    rbPhoto: TRadioButton;
    rbImageLibre: TRadioButton;
    CE_PAGE: TEdit;
    grLibelle: TGroupBox;
    FLibLIbre: TEdit;
    rbCode: TRadioButton;
    rbCodeBarre: TRadioButton;
    rbLibelle: TRadioButton;
    rbLibelleLibre: TRadioButton;
    FBlocage: TCheckBox;
    rdImage: TRadioButton;
    FImage: TSpinEdit;
    tsRetour: TTabSheet;
    grRetourArticle: TGroupBox;
    tRETArt: THLabel;
    tRETQTE: THLabel;
    tRETPRIX: THLabel;
    RETLabel: THLabel;
    RETCode: THCritMaskEdit;
    RETQte: THNumEdit;
    RETPRIX: THNumEdit;
    tsModele: TTabSheet;
    tsMIP: TTabSheet;
    grModele: TGroupBox;
    lblModNature: THLabel;
    MODNATURE: THValComboBox;
    TMODMODELE: THLabel;
    MODMODELE: THValComboBox;
    grMIP: TGroupBox;
    TMIPMODE: THLabel;
    MIPMODE: THValComboBox;
    cbConcept: THValComboBox;
    TcbConcept: THLabel;
    grVendeur: TGroupBox;
    TVencode: THLabel;
    VenCode: THValComboBox;
    tsVide: TTabSheet;
    cbAlt: TCheckBox;
    BFont: TPaletteButton97;
    ArtWhere: TEdit;
    bAssistant: TToolbarButton97;
    remmotif: THValComboBox;
    BCopier: TToolbarButton97;
    TWPropietes: TToolWindow97;
    FBtnsLarge: TSpinEdit;
    FBtnsHaut: TSpinEdit;
    BValideTW: TToolbarButton97;
    BAnnulerTW: TToolbarButton97;
    lblFBtnsHaut: THLabel;
    lblFBtnsLarge: THLabel;
    RGCalculettea: TRadioGroup;
    BPropietes: TToolbarButton97;
    tsFPV: TTabSheet;
    grFPV: TGroupBox;
    tFPVFON: THLabel;
    FPVFON: THValComboBox;
    BFirst: TToolbarButton97;
    BPrev: TToolbarButton97;
    BNext: TToolbarButton97;
    BLast: TToolbarButton97;
    BApercuIm: TToolbarButton97;
    TWImages: TToolWindow97;
    LwImages: TListView;
    BBoutonPos: TToolbarButton97;
    TWBouton: TToolWindow97;
    BValideBtn: TToolbarButton97;
    BCancelBtn: TToolbarButton97;
    TFPage: THLabel;
    TFLigne: THLabel;
    FLigne: TSpinEdit;
    FPage: TSpinEdit;
    TFCol: THLabel;
    FCol: TSpinEdit;
    BCopieSoc: TToolbarButton97;
    Localgauge: TEnhancedGauge;
    FAllera: TSpinEdit;
    tfAllera: THLabel;
    tsFSF: TTabSheet;
    grFSF: TGroupBox;
    TFSFFSF: THLabel;
    FSFFSF: THValComboBox;
    FSFSousFam: TCheckBox;
    BDynamique: TToolbarButton97;
    lblavis: TLabel;
    FPVPage: TSpinEdit;
    tFPVPage: THLabel;
    tsRes: TTabSheet;
    grRess: TGroupBox;
    ResWhere: TEdit;
    Resscode: THCritMaskEdit;
    cbClcVisible: TCheckBox;
    BReproForme: TToolbarButton97;
    TWListe: TToolWindow97;
    GS: THGrid;
    BListe: TToolbarButton97;
    PListePied: THPanel;
    BInsereListe: TToolbarButton97;
    BCancelListe: TToolbarButton97;
    PListeTete: THPanel;
    TcbConceptListe: THLabel;
    cbConceptListe: THValComboBox;
    BValideListe: TToolbarButton97;
    cbChampListe: THValComboBox;
    cbValeurListe: THValComboBox; //NA fin 17/11/03
    LNoPage: THLabel; //NA 11/02/04
    rbImageLibreBase: TRadioButton;
    FImageLibreBase: THCritMaskEdit; // JTR - Image libre stockées
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure BRecupereClick(Sender: TObject);
    procedure FLibelleClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FIMageClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure rbNonCouleurClick(Sender: TObject);
    procedure FToucheExit(Sender: TObject);
    procedure TimerMaxTimer(Sender: TObject);
    procedure cbConceptChange(Sender: TObject);
    procedure FormResize(Sender: TObject);
    Procedure Gestionpardefautcontrol(Sender:Tobject) ;
    procedure FBlocageClick(Sender: TObject);
    procedure BAbandonClick(Sender: TObject);
    procedure MODNATUREChange(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure BFontChange(Sender: TObject);
    procedure FToucheChange(Sender: TObject);
    procedure bAssistantClick(Sender: TObject);
    procedure BCopierClick(Sender: TObject);
    procedure BPropietesClick(Sender: TObject);
    procedure BValideTWClick(Sender: TObject);
    procedure BAnnulerTWClick(Sender: TObject);
    procedure BFirstClick(Sender: TObject);
    procedure BLastClick(Sender: TObject);
    procedure BPrevClick(Sender: TObject);
    procedure BNextClick(Sender: TObject);
    procedure FImageChange(Sender: TObject);
    procedure LwImagesDblClick(Sender: TObject);
    procedure BApercuImClick(Sender: TObject);
    procedure BBoutonPosClick(Sender: TObject);
    procedure BValideBtnClick(Sender: TObject);
    procedure FImageExit(Sender: TObject);
    procedure BCopieSocClick(Sender: TObject);
    procedure FSFSousFamClick(Sender: TObject);
    procedure BDynamiqueClick(Sender: TObject);
    procedure RechercheArtCode(Sender: TObject);
    procedure FPVFONChange(Sender: TObject);
    procedure BloqueFiche; //NA 10/05/04
    procedure pagesChange(Sender: TObject);
    procedure BAideClick(Sender: TObject);        //NA 13/03/02
    procedure BReproFormeClick(Sender: TObject);  //NA debut 17/11/03
    procedure BValideListeClick(Sender: TObject);
    procedure BListeClick(Sender: TObject);
    procedure BCancelListeClick(Sender: TObject);
    procedure BInsereListeClick(Sender: TObject);
    procedure cbChampListeChange(Sender: TObject);
    procedure cbConceptListeChange(Sender: TObject); //NA fin 17/11/02
    // JTR - Image libre stockées
    procedure FImageLibreBaseChange(Sender: TObject);
    procedure FImageLibreBaseClick(Sender: TObject);
    procedure FImageLibreBaseElipsisClick(Sender: TObject);
    procedure TWListeResize(Sender: TObject);
    // Fin JTR - Image libre stockées
  private
    PnlBoutons   : TClavierEcran ;
    BtnCourant   : TOB ;
    BtnReproForme: TOB ; //NA 17/11/03
    EnChargement : Boolean ;
    FParamsCE    : String ;
    FOldParamsCE : String ;
    TouchesFonct : TStringList ;  //XMG 29/11/01
    Procedure  ParamBouton ( UnBouton : TOB) ;
    Procedure  ChargeDataType( Ctrl : TWinControl ; Concept : String ) ;
    Procedure  ChargeBtnCourant ;
    Procedure  EtablirFAllerA ( isEnabled : Boolean ) ;  //XMG 05/10/01
    Procedure  EnregBtnCourant ;
    procedure  ChercheLibelle (Concept : String) ;
    procedure  enablelesrb (Concept : String) ;
    Procedure  CacheTW(TW : TToolWindow97 )  ;
    Procedure  MajNav ;
    Procedure  ChargePosBouton ;
    Procedure  ChangePage(NewPage : Integer ) ;
    Procedure  SetParamsCE ( AParams : String  ) ;
{$IFDEF PGIS5}                                     //MR debut 10/09/2002
    procedure RechercheCommentaire(Sender: TObject);
{$ENDIF}                                          //MR fin 10/09/2002
    procedure ReproduitMiseEnForme(UnBouton: TOB); //NA 17/11/03
    procedure RecupImageLibreBase; // JTR - Image libre stockées
  Protected
  public
    Caisse   : String ;
    Action   : TActionFiche ;
    Property ParamsCE : String read FParamsCE write SetParamsCE ;
  end;


implementation

uses LC_copie,                            //NA debut 25/11/02
{$IFNDEF EAGLCLIENT}
     kb_assis,                            //NA fin 25/11/02
{$ENDIF}
     ParamSoc;  //XMG 09/10/01

{$R *.DFM}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure Param_ClavierEcran ( Caisse : String ; var Params : String ; action : TActionFiche ) ;
Var X : TFParamKB ;
Begin
X:=TFParamKB.Create(Application) ;
X.Caisse:=Caisse ;
X.ParamsCE:=Params ;
X.Action:=Action ;
try
  X.ShowModal ;
  Params:=X.ParamsCE ;
 Finally
  X.Free ;
 End ;
SourisNormale ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////
//XMG debut 29/11/01
Procedure TFParamKB.ChargeDataType( Ctrl : TWinControl ; Concept : String ) ;
Var tt,Pl : String ;
    ex    : TExhaustif  ;
Begin
if (Ctrl is THValCombobox) or (Ctrl is THCritMaskEdit) then
   Begin
   DonneTablette(Concept,Caisse,tt,Pl) ;         //NA 13/03/02
   Ex:=exOui ; if trim(pl)<>'' then Ex:=exPLus ;
   if Ctrl is THValCombobox then
      Begin
      with THValCombobox(Ctrl) do
        Begin
        Datatype:=tt ;
        Plus:=pl ;
        Exhaustif:=ex ;
        End ;
      End else
   if Ctrl is THCritMaskEdit then
      Begin
      with THCritMaskEdit(Ctrl) do
        Begin
        Datatype:=tt ;
        Plus:=pl ;
        End ;
      End else
   End ;
End ;
{$IFDEF PGIS5}                                     //NA debut 25/11/02
//////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.RechercheCommentaire(Sender: TObject);
var QQ : TQuery;
begin
ComCode.Text:=AGLLanceFiche('GC','GCMULCOMMENTAIRE','','','');
if ComCode.Text='' then exit;
QQ:=OpenSql('Select GCT_LIBELLE From COMMENTAIRE WHERE GCT_CODE="'+ComCode.Text+'"', True);
if Not QQ.Eof then
   BEGIN
   ComLabel.Caption:=QQ.FindField('GCT_LIBELLE').AsString ;
   END;
Ferme(QQ);
end;
{$ENDIF}                                          //NA fin 25/11/02
//XMG 14/03/02 debut
//////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.RechercheArtCode(Sender: TObject);
begin
{$IFDEF PGIS5}
FOGetArticleRecherche(TControl(Sender), '', 'FFO') ;
{$ENDIF}
end;
//XMG 14/03/02 fin
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.FormShow(Sender: TObject);
var i  : integer ;
    //QQ : TQuery ;            //NA 25/11/02
begin
for i:=0 to pages.PageCount-1 do Pages.pages[i].TabVisible:=FALSE ; //XMG 20/07/01
{$IFDEF EAGLCLIENT}                    //NA debut 04/12/02
BAssistant.Visible:=False ; 
{$ELSE}
BAssistant.Visible:=not (V_PGI.Laserie in [S1,S2,S3]) ; //XMG 20/07/01
{$ENDIF}                               //NA fin 04/12/02
EnableControls(Self,FALSE,TRUE) ;
LWImages.Items.Clear ;
lwImages.LargeImages:=PnlBoutons.Images ;
For i:=0 to PnlBoutons.Images.Count-1 do
    with LWImages.items.Add do
      Begin
      Caption:=inttostr(i+1) ;
      ImageIndex:=i ;
      End ;
BBoutonPos.Enabled:=FALSE ;
bassistant.Enabled:=FALSE ;
PnlParam.Enabled:=(Action in [taCreat,taModif]) ;
EnableControls(TWPropietes,Action in [taCreat,taModif]) ;
if not BAnnulerTW.Enabled then BAnnulerTW.Enabled:=TRUE ;
BValideTW.Visible:=(Action in [taCreat,taModif]) ;

{$IFDEF PGIS5}             //NA debut 06/11/02
BDynamique.Visible:=FALSE ; //XMG 09/10/01
{$ELSE}
BDynamique.Visible:=TRUE; //XMG 09/10/01
{$ENDIF}                  //NA fin 06/11/02
{$IFDEF EAGLCLIENT}       //MR debut 23/10/02
BCopieSoc.Visible:=False ;
{$ELSE}
BCopieSoc.Visible:=TRUE ;
{$ENDIF}                  //MR fin 23/10/02
cbConcept.Values.Clear ;
cbConcept.Items.Clear ;
cbConcept.Values.Add('0') ;
cbConcept.Items.Add(Traduirememoire('Vide')) ;
cbConcept.Values.Add('1') ;
cbConcept.Items.Add(Traduirememoire('Article')) ;
{$IFDEF PGIS5}             //NA debut 06/11/02
{$ELSE}
cbConcept.Values.Add('13') ;
cbConcept.Items.Add(Traduirememoire('Famille/SousFamille')) ;
{$ENDIF}                  //NA fin 06/11/02
cbConcept.Values.Add('2') ;
cbConcept.Items.Add(Traduirememoire('Commentaire')) ;
cbConcept.Values.Add('3') ;
cbConcept.Items.Add(Traduirememoire('Remise')) ;
cbConcept.Values.Add('4') ;
cbConcept.Items.Add(Traduirememoire('Vendeur')) ; 
cbConcept.Values.Add('5') ;
cbConcept.Items.Add(Traduirememoire('Fonction')) ;
cbConcept.Values.Add('6') ;
cbConcept.Items.Add(Traduirememoire('Règlement')) ;
{$IFDEF PGIS5}             //NA debut 06/11/02
cbConcept.Values.Add('7') ;
cbConcept.Items.Add(Traduirememoire('Opérations de caisse')) ; //NA 13/03/02
{$ELSE}
cbConcept.Values.Add('7') ;
cbConcept.Items.Add(Traduirememoire('Opération caisse')) ;
cbConcept.Values.Add('8') ;
cbConcept.Items.Add(Traduirememoire('Opération financière')) ;
cbConcept.Values.Add('9') ;
cbConcept.Items.Add(Traduirememoire('Retour article')) ;
{$ENDIF}                  //NA fin 06/11/02
cbConcept.Values.Add('10') ;
cbConcept.Items.Add(Traduirememoire('Modèle d''impression')) ;
{$IFDEF PGIS5}             //NA debut 06/11/02
{$ELSE}
cbConcept.Values.Add('11') ;
cbConcept.Items.Add(Traduirememoire('Mode d''impression')) ;
{$ENDIF}                  //NA fin 06/11/02
cbConcept.Values.Add('12') ;
cbConcept.Items.Add(Traduirememoire('Fonctions spéciales')) ;
{$IFDEF CHR} //NA debut 17/11/03
cbConcept.Values.Add('14') ;
cbConcept.Items.Add(Traduirememoire('Ressources')) ;
{$ENDIF} //NA fin 17/11/03

{$IFDEF PGIS5}             //NA debut 06/11/02
BCopier.Visible:=(not FOMonoCaisse) ;
{$ELSE}
BCopier.Visible:=not VS1^.Monocaisse ;
{$ENDIF}                  //NA fin 06/11/02
ChargeDataType(ArtCode,'ART') ;
{$IFDEF PGIS5}             //NA debut 06/11/02
ArtCode.OnElipsisClick:=RechercheArtCode ;         //NA debut 13/03/02
ArtCode.OnDblClick:=RechercheArtCode ;             //NA fin 13/03/02
{$ENDIF}                  //NA fin 06/11/02

ChargeDataType(ComCode,'COM') ;
{$IFDEF PGIS5}             //NA debut 06/11/02
if GetParamSoc('SO_GCCOMMENTAIRE') then            //MR debut 10/09/2002
   begin
   ComCode.DataType:='' ;
   ComCode.Plus:='' ;
   ComCode.Libelle:=nil ;
   ComCode.OnElipsisClick:=RechercheCommentaire;
   ComCode.OnDblClick:=RechercheCommentaire;
   end;                                            //MR fin 10/09/2002
{$ENDIF}                  //NA fin 06/11/02
ChargeDataType(RemMotif,'REM') ;

ChargeDataType(VenCode,'VEN') ;

ChargeDataType(FonCode,'FON') ;

ChargeDatatype(RegCode,'REG') ;


{$IFDEF PGIS5}             //NA debut 06/11/02
OctCode.Vide:=True ;                //NA debut 25/11/02
OctCode.VideString:='<<Tous>>' ;
OctCode.Plus:='' ;                  //NA fin 25/11/02
lblOctPrix.Caption:=TraduireMemoire('Prix') ;  //NA fin 03/07/02
{$ENDIF}                  //NA fin 06/11/02
ChargeDatatype(OctCode,'OCT') ;

ChargeDatatype(OfcCode,'OFC') ;

ChargeDatatype(RetCode,'ART') ;

ChargeDatatype(MODNature,'MODNAT') ;
if ModNature.Values.count>0 then ModNature.Value:=ModNature.Values[ModNature.Values.count-1] ;  //NA 06/11/02

ChargeDatatype(MIPMode,'MIP') ;

ChargeDatatype(FPVFON,'FPV;X;') ;

FSFSousFam.Checked:=FALSE ;

PnlAvis.BringtoFront ;
PnlParam.SendToBack ;
Pages.ActivePage:=Pages.FindNextPage(nil,TRUE,FALSE) ;
PnlBoutons.Caisse:=Caisse ;
MajNav ;

TimerMax.Enabled:=TRUE ;
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
FCouleur.AddStdColors(FALSE) ;
{$IFDEF PGIS5}
ArtPrix.Masks.PositiveMask:=StrfMask(V_PGI.OkDecV,'',TRUE,FALSE) ;
OCTPrix.Masks.PositiveMask:=StrfMask(V_PGI.OkDecV,'',TRUE,FALSE) ;
OFCPrix.Masks.PositiveMask:=StrfMask(V_PGI.OkDecV,'',TRUE,FALSE) ;
RETPrix.Masks.PositiveMask:=StrfMask(V_PGI.OkDecV,'',TRUE,FALSE) ;
ArtQTE.Masks.PositiveMask:=StrfMask(V_PGI.OkDecQ,'',TRUE,FALSE) ;
{$ELSE}
ArtPrix.Masks.PositiveMask:=StrfMask(OkDecPV('P'),'',TRUE,FALSE) ;
OCTPrix.Masks.PositiveMask:=StrfMask(V_PGI.OkDecV,'',TRUE,FALSE) ;
OFCPrix.Masks.PositiveMask:=StrfMask(V_PGI.OkDecV,'',TRUE,FALSE) ;
RETPrix.Masks.PositiveMask:=StrfMask(OkDecPV('P'),'',TRUE,FALSE) ;
ArtQTE.Masks.PositiveMask:=StrfMask(OkDecPV('Q'),'',TRUE,FALSE) ;
{$ENDIF}

FLibelleClick(nil) ;
FImageClick(nil) ;
PnlParam.Visible:=FALSE ;
PnlAvis.Visible:=TRUE ;
EnableControls(Self,TRUE,TRUE) ;
{$IFDEF PGIS5}                                                //NA debut 20/09/2002
HelpContext:=301300191 ;
{$ENDIF}
// Appel de la fonction d'empilage dans la liste des fiches
AglEmpileFiche(Self) ;                                       //NA fin 20/09/2002
end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.FormCreate(Sender: TObject);
var kk    : TStringList ;  //XMG debut 29/11/01
    St,TS : String ;
    i     : Integer ;      //XMG fin 29/11/01
begin
PnlParam.top:=1 ; PnlParam.Left:=1 ; PnlParam.Width:=pMain.Width ;
PnlAvis.top:=1  ; PnlAvis.left:=1  ; PnlAvis.Width:=pMain.Width ;
PnlAvis.Height:=PnlParam.Height ;
PMain.Tag:=PMain.Height-PnlParam.Height ;

PnlBoutons:=TClavierEcran.Create(Self) ;
PnlBoutons.Parent:=PMain ;
PNLBoutons.Align:=alBottom ;
PnlBoutons.Height:=Pmain.Tag ;
PnlBoutons.colornb:=6 ;
PnlBoutons.Parametrage:=TRUE ;
PnlBoutons.ParamBouton:=ParamBouton ;
PnlBoutons.Gauge:=LocalGauge ;
BtnCourant:=nil ;
BtnReproForme:=nil ; //NA 17/11/03
EnChargement:=FALSE ;

FImage.MaxValue:=PnlBoutons.Images.Count ;
FOldPAramsCE:='' ;
//XMG debut 29/11/01
TouchesFonct:=TStringlist.Create ;
kk:=TStringlist.Create ;
{$IFDEF PGIS5}
RemplirValCombo('GCFONCTIONSPVFO',' and CO_ABREGE<>""','',TouchesFonct,kk,FALSE,TRUE) ;
{$ELSE}
RemplirValCombo('TTFONCTIONSPV',' and CO_ABREGE<>""','',TouchesFonct,kk,FALSE,TRUE) ;
{$ENDIF}       //NA fin 13/03/02
kk.free ;
For i:=0 to TouchesFonct.Count-1 do
  Begin
  St:=TouchesFonct[i] ;
  if pos(';TS_',St)>0 then
     Begin
     if sright(St,1)<>';' then St:=St+';' ;
     TS:=gtfs(St,';',4) ; St:=Copy(St,1,pos(TS,St)-1) ;
     if TS='TS_CHERCHE'  then TS:=inttostr(VK_CHERCHE)  else
{$IFDEF PGIS5} //NA debut 06/11/02
     if TS='TS_DELLIGNE' then TS:=inttostr(VK_DELETE)   else
     if TS='TS_INSLIGNE' then TS:=inttostr(VK_INSERT)   else
{$ELSE}
     if TS='TS_DELLIGNE' then TS:=inttostr(VK_DELLIGNE) else
     if TS='TS_INSLIGNE' then TS:=inttostr(VK_INSLIGNE) else
{$ENDIF}       //NA fin 06/11/02
     if TS='TS_LISTE'    then TS:=inttostr(VK_LISTE)    else
     if TS='TS_HOME'     then TS:=inttostr(VK_HOME)     else
     if TS='TS_LEFT'     then TS:=inttostr(VK_LEFT)     else
     if TS='TS_CHOIX'    then TS:=inttostr(VK_CHOIX)    else
     if TS='TS_RIGHT'    then TS:=inttostr(VK_RIGHT)    else
     if TS='TS_FIN'      then TS:=inttostr(VK_END)      else
{$IFDEF PGIS5} //NA debut 13/03/02
     if TS='TS_APPELDIM' then TS:='D'                   else   //NA 25/11/02
     if TS='TS_ECHEANCE' then TS:='E'                   else   //NA 25/11/02
{$ELSE}
     if TS='TS_ATTENTE'  then TS:=inttostr(VK_ATTENTE)  else
     if TS='TS_INFOCOMP' then TS:=inttostr(VK_INFOCOMP) else
     if TS='TS_RECALCUL' then TS:=inttostr(VK_RECALCUL) else
     if TS='TS_PRIX'     then TS:=VK_PRIX               else
     if TS='TS_QUANTITE' then TS:=VK_QUANTITE           else
     if TS='TS_REMISE'   then TS:=gtfs(VK_REMISE,';',1) else
     if TS='TS_PLUS'     then TS:=VK_PLUS               else
     if TS='TS_MOINS'    then TS:=VK_MOINS              else
{$ENDIF}       //NA fin 13/03/02
     ;
     if St+TS+';'<>TouchesFonct[i] then TouchesFonct[i]:=St+TS+';' ;
     End ;
  End ;
//XMG fin 29/11/01
end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.FormDestroy(Sender: TObject);
begin
TouchesFonct.Free ;  //XMG 29/11/01
//if Assigned(BtnCourant) then BtnCourant.Free ;
PnlBoutons.Free ;
// Appel de la fonction de dépilage dans la liste des fiches
AglDepileFiche ;                                                   //NA 20/09/2002
end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure  TFParamKB.ParamBouton ( UnBouton : TOB) ;
Begin
EnregBtnCourant ;
PnlAvis.Visible:=FALSE ;
PnlParam.Visible:=TRUE ;
//if Assigned(BtnCourant) then begin BtnCourant.Free ;BtnCourant:=Nil ; End ;
//BtnCourant:=TOB.Create('CLAVIERECRAN',nil,-1) ;
//BtnCourant.Assign(UnBouton) ;
BtnCourant:=UnBouton ;
ReproduitMiseEnForme(UnBouton); //NA 17/11/03
ChargeBtnCourant ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.BRecupereClick(Sender: TObject);
begin
{if assigned(BtnCourant) then
   Begin
   BtnCourant.Free ;
   BtnCourant:=nil ;
   End ;}
PnlParam.Visible:=FALSE ;
PnlAvis.Visible:=TRUE ;
BtnCourant:=nil ;
BBoutonPos.Enabled:=FALSE ;
PnlBoutons.ChargeClavier(Caisse) ;
ParamsCE:=FOldParamsCE ;
ChangePage(0) ;
end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.ChercheLibelle (Concept : String) ;
var code             : THCritMaskEdit ;
    libe             : TLabel ;
    Libelle,Cle1,SQL : String ;
    Q                : TQuery ;
begin
if pos(';'+concept+';',';ART;COM;REM;RET;')<=0 then exit ;
code:=THCritmaskedit(FindComponent(Concept+'Code')) ;
Libe:=TLabel(FindComponent(Concept+'Label')) ;
if (assigned(code)) and (assigned(libe)) then
   begin
   SQL:=KBMakeSQL(Concept, Code.Text, Cle1, Libelle) ; if SQL='' then exit ;
   Q:=OpenSQL(SQL,TRUE) ;
   if not Q.Eof then libe.Caption:=Q.FindField(Libelle).AsString ;
   Ferme(Q) ;
   End ;
end ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.enablelesrb (Concept : String) ;
var okArt : boolean ;
begin
{$IFDEF CHR} //NA debut 17/11/03
OkArt:=(pos(';'+Concept+';',';ART;RET;RES;')>0) ;
{$ELSE}
OkArt:=(pos(';'+Concept+';',';ART;RET;')>0) ;
{$ENDIF} //NA fin 17/11/03
rbPhoto.Enabled:=OkArt ; FIMageClick(nil) ;
rbCodeBarre.Enabled:=OkArt ; FLibelleClick(nil) ;
bassistant.Enabled:=OkArt ;  
end ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure KBParamToNumEdit(Data: string; NumEdit: THNumEdit); //NA debut 03/03/04
var
  dVal: double;
  OldSepDecimal, OldSepMillier: char;
begin
  // séparateurs French pour la fonction Valeur()
  OldSepDecimal := V_PGI.SepDecimal;
  OldSepMillier := V_PGI.SepMillier;
  V_PGI.SepDecimal := ',';
  V_PGI.SepMillier := ' ';
  try
    dVal := Valeur(Data);
  finally
    V_PGI.SepDecimal := OldSepDecimal;
    V_PGI.SepMillier := OldSepMillier;
  end;
  NumEdit.Value := dVal;
end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function NumEditTOKBParam(NumEdit: THNumEdit): string;
var
  dVal: double;
  OldDecimalSeparator, OldThousandSeparator: char;
begin
  dVal := NumEdit.Value;
  // séparateurs French pour la conversion des montants
  OldDecimalSeparator := DecimalSeparator;
  OldThousandSeparator := ThousandSeparator;
  DecimalSeparator := ',';
  ThousandSeparator := ' ';
  try
    Result := STRFMONTANT(dVal, NumEdit.Digits, NumEdit.Decimals, NumEdit.CurrencySymbol, True);
  finally
    DecimalSeparator := OldDecimalSeparator;
    ThousandSeparator := OldThousandSeparator;
  end;
end;  //NA fin 03/03/04
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TFParamKB.ChargeBtnCourant ;
         //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
         procedure razcontrols (Ctrl : TControl) ;
         var j    : integer ;
         begin
         if Ctrl is TCombobox      then TCombobox(CTRL).itemindex:=-1 else
         if Ctrl is THNumEdit      then THNumEdit(Ctrl).value:=0 else
         if Ctrl is TEdit          then TEdit(Ctrl).Text:='' else
         if Ctrl is THCritMaskEdit then THCritMaskEdit(Ctrl).Text:='' ;
         if (Ctrl is TLabel) and
            (TLabel(CTRL).tag=1)   then TLabel(Ctrl).Caption:='' ;
         if Ctrl is TWinControl    then
            for j:=0 to twincontrol(Ctrl).ControlCount-1 do RazControls(tWinControl(Ctrl).Controls[j]) ;
         end ;
         //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Var St1,st2,st3,val : String ;  //XMG 05/10/01
    i,p             : integer ;
    st0, st4, st5   : string; // Fin JTR - Image libre stockées
Begin
EnChargement:=TRUE ;
for i:=0 to Pages.PageCount-1 do RazControls(Pages.Pages[i]) ;
EnChargement:=FALSE ;

if not assigned(BtnCourant) then exit ;
EnChargement:=TRUE ;
BtnCourant.PutEcran(Self,nil) ;
FBlocage.Checked:=(VInteger(BtnCourant.GetValue('CE_PAGE'))=-1) ;
FAllera.Value:=VInteger(BtnCourant.GetValue('CE_ALLERA')) ;
St1:=VString(BtnCourant.GetValue('CE_TOUCHE')) ;
St2:=Readtokenst(ST1) ;
cbCtrl.Checked:=(st2='X') ;
St2:=Readtokenst(ST1) ;
cbShift.Checked:=(st2='X') ;
St2:=Readtokenst(ST1) ;
cbALT.Checked:=(st2='X') ;
if trim(St1)<>'' then FTouche.Value:=readtokenSt(St1) else FTouche.ItemIndex:=0 ;
FToucheExit(nil) ;
if VInteger(BtnCourant.GetValue('CE_COULEUR')) and $80000000<>0 then rbNonCouleur.Checked:=TRUE else
   begin
   rbCouleur.Checked:=TRUE ;
   FCouleur.SelectedColor:=TColor(VInteger(BtnCourant.GetValue('CE_COULEUR'))) ;
   End ;
rbNoncouleurclick(nil) ;
FLibLibre.Text:='' ;
// JTR - Image libre stockées
//St2:=VString(BtnCourant.GetValue('CE_IMAGE')) ; St1:=ReadTokenSt(St2) ; st2:=ReadTokenSt(St2) ;
St0    := VString(BtnCourant.GetValue('CE_IMAGE')) ;
St1    := ReadTokenSt(St0) ;
St2    := Readtokenst(St0) ;
St3    := Readtokenst(St0) ;
St4    := Readtokenst(St0) ;
St5    := Readtokenst(St0) ;
// Fin JTR - Image libre stockées

if St1='**PHOTO**' then
  rbPhoto.Checked:=TRUE
else if St1='**INTERNE**' then
Begin
   rdImage.Checked:=TRUE ;
   FImage.Value:=Valeuri(St2)+1 ;
End else if St1='**LIENSOLE**' then // JTR - Image libre stockées
begin
   rbImageLibreBase.Checked := TRUE ;
   FImageLibreBase.Text := St5 ;
   End else
   Begin
   rbImageLibre.Checked:=TRUE ;
   FImageLibre.Text:=St1 ;
   End ;
FImageClick(nil) ;
St1:=VString(BtnCourant.GetValue('CE_CAPTION')) ;
//if Copy(St1,1,2)<>#255+#255 then //NA 06/01/04
if not IsKBIDCaptionStd(St1) then //NA 06/01/04
   begin
   FLibLibre.Text:=St1 ;
   rblibellelibre.checked:=TRUE ;
   End else
   Begin
   //delete(St1,1,2) ; //NA 06/01/04
   if St1='ART' then rbCode.Checked:=TRUE else
   if St1='BAR' then rbCodeBarre.Checked:=TRUE else rbLibelle.Checked:=TRUE ;
   End ;
FLibelleCLick(nil) ;
st1:=VString(BtnCourant.GetValue('CE_TEXTE')) ;
St2:=Readtokenst(ST1) ; p:=0 ;
bassistant.Enabled:=FALSE ;
//XMG debut 05/10/01
//Val:=#255+#255 ; //NA 06/01/04
Val := GetKBIDCaptionStd; //NA 06/01/04
if st2='VIDE' then
   begin
   p:=0 ;
   end else
if st2='ART' then
   begin
   artcode.text:=ReadTokenst(St1) ; Val:=ArtCode.Text ;
   KBParamToNumEdit(ReadTokenST(st1), artqte); //NA 03/03/04
   KBParamToNumEdit(ReadTokenST(st1), artprix); //NA 03/03/04
   ArtWhere.text:=ReadTokenst(St1) ;
   bassistant.Enabled:=TRUE ;
   p:=1 ;
   end else
if st2='COM' then
   begin
   comcode.text:=readtokenst(st1) ; Val:=Comcode.text ;
   p:=2 ;
   end else
if st2='REM' then
   begin
   KBParamToNumEdit(ReadTokenST(st1), rempourcent); //NA 03/03/04
   remmotif.Value:=readtokenst(st1) ;  Val:=RemMotif.Value ;
   p:=3 ;
   end else
if st2='VEN' then
   begin
   vencode.Value:=Readtokenst(st1) ; Val:=Vencode.Value ;
   p:=4 ;
   end else
if st2='FON' then
   begin
   FonCode.Value:=Readtokenst(St1) ; Val:=FonCode.Value ;
   p:=5 ;
   end else
if st2='REG' then
   begin
   RegCode.value:=readtokenst(St1) ; Val:=RegCode.Value ;
   p:=6 ;
   end else
if st2='OCT' then
   begin
   OCTCode.value:=readtokenst(St1) ;  Val:=OCTCode.Value ;
   KBParamToNumEdit(ReadTokenST(st1), OCTPrix); //NA 03/03/04
   p:=7 ;
   end else
if st2='OFC' then
   begin
   OFCCode.value:=readtokenst(St1) ;  Val:=OFCCode.Value ;
   KBParamToNumEdit(ReadTokenST(st1), OFCPrix); //NA 03/03/04
   p:=8 ;
   end else
if st2='RET' then
   begin
   Retcode.text:=ReadTokenst(St1) ; Val:=RetCOde.Text ;
   KBParamToNumEdit(ReadTokenST(st1), retqte); //NA 03/03/04
   KBParamToNumEdit(ReadTokenST(st1), retprix); //NA 03/03/04
   p:=9 ;
   end else
if st2='MOD' then
   begin
   st3:=ReadTokenst(St1) ;
   MODNature.Value:=ReadTokenSt(St1) ;
   MODModele.value:=St3; Val:=MODModele.Value ;
   p:=10 ;
   end else
if st2='MIP' then
   begin
   MIPMode.Value:=ReadTokenst(St1) ; Val:=MipMode.Value ;
   p:=11 ;
   end else
if st2='FPV' then
   begin
   FPVFon.Value:=ReadTokenst(St1) ; Val:=FPVFon.Value ;
   p:=12 ;
//XMG 15/04/02 début
   FPVPage.Enabled:=FALSE ;
   if FPVFon.Value='GOT' then
      Begin
      FPVPage.Enabled:=TRUE ;
      FPVPage.Value:=Valeuri(ReadTokenSt(St1)) ;
      End ;
//XMG 15/04/02 fin
//XMG debut 20/07/01
   end else
if st2='FSF' then
   begin
   St3:=ReadTokenSt(St1) ;
   FSFSousFam.Checked:=TRUEFALSEBo(ReadTokenst(St1)) ;
   FSFFSF.Value:=St3 ; Val:=FSFFSF.Value ;
   p:=13 ;
//XMG fin 20/07/01
  {$IFDEF CHR} //NA debut 17/11/03
   end else
   if st2='RES' then
   begin
   resscode.text:=ReadTokenst(St1) ; Val:=ressCode.Text ;
   bassistant.Enabled:=TRUE ;
   ResWhere.text:=ReadTokenst(St1) ;
   bassistant.Enabled:=TRUE ;
   p:=14 ;
  {$ENDIF} // FIN CHR
   End ; //NA fin 17/11/03
///cbConcept.itemindex:=p ;
cbConcept.itemindex:=cbConcept.Values.IndexOf(IntToStr(p)) ;
EnablelesRB(St2) ;
cherchelibelle(St2) ;
AppliquePolice(BtnCourant.GetValue('CE_POLICE'), BFont.Font, BFont.Alignment)  ;
BloqueFiche; //NA 10/05/04
{$IFDEF PGIS5}  //NA debut 17/11/03
EtablirFAllerA((not (p in [0,12]))) ;
{$ELSE}
EtablirFAllerA((not (p in [0,12]) and (Trim(Val)<>''))) ;
{$ENDIF}        //NA fin 17/11/03
pages.activepage:=Pages.Pages[p] ;
BBoutonPos.Enabled:=(VString(BtnCourant.GetValue('CE_TEXTE'))<>ConstVide) ;  //XMG 09/10/01
EnChargement:=FALSE ;
if cbConcept.canfocus then cbConcept.setfocus ;
EnregBtnCourant ;
End ;
//XMG debut 05/10/01
//////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TFParamKB.EtablirFAllerA ( isEnabled : Boolean ) ;
Begin
FAllerA.enabled:=IsEnabled ;
FAllerA.Value:=FAllerA.Value*ord(IsEnabled)
End ;
//XMG fin 05/10/01
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TFParamKB.EnregBtnCourant ;
Var St     : String ;
    Val    : Integer ;
    Ok     : Boolean ;
    BTNCol : TOB ; //XMG 29/11/01
Begin
if (assigned(BtnCourant) and (not EnChargement)) then
   Begin
   if cbConcept.ItemIndex=0 then
      Begin
      BtnCourant.Initvaleurs ;
      BtnCourant.GetEcran(Self) ;
      BtnCourant.PutValue('CE_ALLERA',0) ;  //XMG 11/05/01
      BtnCourant.PutValue('CE_TEXTE',ConstVide)  ;  //XMG 09/10/01
      BtnCourant.PutValue('CE_CAPTION',TraduireMemoire('Libre')) ;
      BtnCourant.PutValue('CE_COULEUR',integer($80000000)) ;
      BtnCourant.PutValue('VALCAPTION',BtnCourant.GetValue('CE_CAPTION')) ;
      BtnCourant.PutValue('IndexImage',0) ;
      End else
      Begin
      BtnCourant.GetEcran(Self,nil) ;
      BtnCourant.PutValue('CE_ALLERA',FAllera.value) ;  //XMG 11/05/01
      St:=FTouche.Value+';' ;
      if trim(St)<>'' then
         Begin
         St:=TrueFalseSt(cbALT.State=cbChecked)+';'+St ;
         St:=TrueFalseSt(cbShift.State=cbChecked)+';'+St ;
         St:=TrueFalseSt(cbCtrl.State=cbChecked)+';'+St ;
         End ;
//XMG debut 29/11/01
      BtnCol:=PnlBoutons.ChercheBoutonKey(St) ;
      if ((Assigned(BtnCol)) and (BtnCol<>BtnCourant)) or
         ((trim(St)<>'') and (TouchesFonct.IndexOf(St)>-1)) then PGIBox(MsgErrDefautS1(752),Caption) ; //XMG 07/01/04
//XMG fin 29/11/01
      BtnCourant.PutValue('CE_TOUCHE',St) ;
      if rbnoncouleur.checked then val:=integer($80000000) else val:=ord(FCouleur.SelectedColor) ;
      BtnCourant.PutValue('CE_COULEUR',Val) ;
      St:='' ;
// JTR - Image libre stockées
      if rbPhoto.Checked then
        St := '**PHOTO**'
      else if rdImage.Checked then
        St := '**INTERNE**;'+inttostr(FImage.Value-1) +';'
      else if rbImageLibre.Checked then
        St := FImageLibre.Text
      else if rbImageLibreBase.Checked then
        St := BtnCourant.GetValue('CE_IMAGE'); // La valeur est déjà affectée
// Fin JTR
      BtnCourant.PutValue('CE_IMAGE',St) ;
      St:='' ;
      if rbLibelleLibre.checked then St:=FLibLibre.Text else
         Begin
         //St:=#255+#255 ; //NA 06/01/04
         St:=GetKBIDCaptionStd; //NA 06/01/04
         if RbCode.Checked then St:=St+'COD' else
         if RbCodeBarre.Checked then St:=St+'BAR' else St:=St+'LIB' ;
         End ;
      BtnCourant.PutValue('CE_CAPTION',St) ;
      ////case cbconcept.itemindex of
      case StrToInt(cbconcept.Value) of  //NA debut 12/03/03
        0  : St:='VIDE;' ; //on ne devrait jamais arriver ici.
        1  : St:='ART;'+artcode.Text+';'+NumEditTOKBParam(artqte)+';'+NumEditTOKBParam(artprix)+';'+trim(ArtWhere.text)+';' ; //NA 03/03/04
        2  : St:='COM;'+comcode.text+';' ;
        3  : st:='REM;'+NumEditTOKBParam(rempourcent)+';'+remMotif.Value+';' ; //NA 03/03/04
        4  : st:='VEN;'+vencode.value+';' ;
        5  : st:='FON;'+fonCode.Value+';' ;
        6  : st:='REG;'+RegCode.Value+';' ;
        7  : st:='OCT;'+OCTCode.Value+';'+NumEditTOKBParam(OCTprix)+';' ; //NA 03/03/04
        8  : st:='OFC;'+OFCCode.Value+';'+NumEditTOKBParam(OFCprix)+';' ; //NA 03/03/04
        9  : st:='RET;'+RETcode.Text+';'+NumEditTOKBParam(RETqte)+';'+NumEditTOKBParam(RETprix)+';' ; //NA 03/03/04
        10 : st:='MOD;'+MODModele.Value+';'+ModNature.Value+';' ;
        11 : st:='MIP;'+MIPMODE.Value+';' ;
        12 : st:='FPV;'+FPVFON.Value+';'+inttostr(FPVPage.Value)+';' ;  //XMG 15/04/02
        13 : st:='FSF;'+FSFFSF.Value+';'+TRUEFALSESt(FSFSOusFam.Checked)+';' ;  //XMG 20/07/01
{$IFDEF CHR} //NA 17/11/03
        14 : st:='RES;'+resscode.Text+';'+trim(ResWhere.text)+';' ;
{$ENDIF} //NA 17/11/03
        end ;     //NA fin 12/03/03
      BtnCourant.PutValue('CE_TEXTE',st) ;
      Ok:=(trim(readtokenst(St))<>'') and (not (Valeuri(cbconcept.Value) in [0,12])) ;  //XMG 15/04/02
      if StrToInt(cbconcept.Value)=3 then Ok:=trim(readtokenst(St))<>'' ;
      EtablirFAllerA(Ok) ;
      End ;
   PnlBoutons.UpdateBouton(BtnCourant) ;
   End ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.FLibelleClick(Sender: TObject);
begin
rbCode.Checked:=((not RbCodeBarre.Enabled) and (rbCodeBarre.Checked)) or (rbCode.Checked) ;
FLibLibre.Enabled:=(rbLibelleLibre.Checked) ;
if (Not FLibLibre.Enabled) and (trim(FLibLibre.Text)<>'') then FLibLibre.Text:='' ;
EnregBtnCourant ;
end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if IsInside(Self) then Action:=caFree ;
end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.FIMageClick(Sender: TObject);
begin
RbImageLibre.checked:=((not rbPhoto.enabled) or (RbImageLibre.checked)) and (not rdImage.Checked) and (not rbImageLibreBase.Checked) ;
FImage.Enabled:=(RdImage.checked) ;
BApercuIm.Visible:=FImage.Enabled ;
if Not FIMage.Enabled then FIMage.Value:=0 ;
FImageLibre.Enabled:=(RbImageLibre.checked) ;
if Not FIMageLibre.Enabled then FIMageLibre.Text:='' ;
// JTR - Image libre stockées
FIMageLibreBase.Enabled := rbImageLibreBase.Checked;
if Not FIMageLibreBase.Enabled then FIMageLibreBase.Text:='' ;
// Fin JTR - Image libre stockées
EnregBtnCourant ;
end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.BValiderClick(Sender: TObject);
begin
//XMG debut 30/11/01
NextControl(Self,TRUE) ;
PnlParam.Visible:=FALSE ;
lblAvis.Visible:=FALSE ;
PnlAvis.Visible:=TRUE ;
Application.ProcessMessages ; 
EnregBtnCourant ;
PnlBoutons.EnregClavier ;
ParamsCE:=inttostr(PnlBoutons.NbrBtnWidth)+';'+inttostr(PnlBoutons.NbrBtnHeight)+';'+inttostr(ord(PnlBoutons.ClcPosition))+';' //NA debut 17/11/03
  +TrueFalseSt(PnlBoutons.ClcVisible)+';' ;
//ParamsCE:=inttostr(PnlBoutons.NbrBtnWidth)+';'+inttostr(PnlBoutons.NbrBtnHeight)+';'+inttostr(ord(PnlBoutons.ClcPosition))+';' ; //NA fin 17/11/03
lblAvis.Visible:=TRUE ;
//XMG fin 30/11/01
end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
Canclose:=(not PnlBoutons.isModified) or (PGIAsk(MsgErrDefautS1(744), caption)=mrYes) ; //XMG 07/01/04
end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.rbNonCouleurClick(Sender: TObject);
begin
FCouleur.Enabled:=(not rbNonCouleur.checked) ;
if Not FCouleur.Enabled then FCouleur.itemindex:=-1 ;
EnregBtnCourant ;
end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.FToucheExit(Sender: TObject);
Var WCtrl : TWinControl ;
begin
cbCtrl.Enabled:=(FTouche.ItemIndex>0) ;
cbShift.Enabled:=(FTouche.ItemIndex>0) ;
cbALT.Enabled:=(FTouche.ItemIndex>0) ;
if FTouche.ItemIndex<=0 then
   Begin
   cbCtrl.State:=cbUnchecked ;
   cbShift.State:=cbUnchecked ;
   cbALT.State:=cbUnchecked ;
   End ;
WCtrl:=FindNextControl(FTouche,AbsoluteGoForward(Ftouche,ActiveControl),TRUE,FALSE) ;
if (Assigned(WCtrl)) and (WCtrl.Canfocus) then WCtrl.Setfocus ;
EnregBtnCourant ;
end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.FToucheChange(Sender: TObject);
begin
cbCtrl.Enabled:=(FTouche.ItemIndex>0) ;
cbShift.Enabled:=(FTouche.ItemIndex>0) ;
cbALT.Enabled:=(FTouche.ItemIndex>0) ;
if FTouche.ItemIndex<=0 then
   Begin
   cbCtrl.State:=cbUnchecked ;
   cbShift.State:=cbUnchecked ;
   cbALT.State:=cbUnchecked ;                               
   End ;
end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.TimerMaxTimer(Sender: TObject);
begin
TimerMax.Enabled:=FALSE ;
WindowState:=wsMaximized ;
end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.cbConceptChange(Sender: TObject);
//XMG Debut 15/04/02
begin
////pages.ActivePage:=pages.Pages[cbConcept.itemindex] ;
pages.ActivePage:=pages.Pages[StrToInt(cbConcept.Value)] ;
PagesChange(Pages) ; //XMG on doit appeler directement au OnCHange, car le control ne le fait que pour Click ou SlectNext (bug?)
end;
//XMG 15/04/02 fin
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.FormResize(Sender: TObject);
begin
PMain.Tag:=PMain.Height-PnlParam.Height ;
if assigned(PnlBoutons) then PnlBoutons.ResizeClavier(nil) ;
end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TfParamKb.Gestionpardefautcontrol(Sender:Tobject) ;
Begin
EnregBtnCourant ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.FBlocageClick(Sender: TObject);
begin
if FBlocage.state=cbChecked then
   Begin
   if Valeuri(CE_PAGE.Text)<>-1 then
      Begin
      if PnlBoutons.BoutonsurAutrespages(BtnCourant) then
         Begin
         PGIBox(MsgErrDefautS1(787),Caption) ; //XMG 07/01/04
         FBlocage.State:=cbUnChecked ;
         End else CE_PAGE.Text:='-1' ;
      End ;
   End else if valeuri(CE_PAGE.Text)<>PnlBoutons.PageCourante then CE_PAGE.Text:=inttostr(PnlBoutons.PageCourante) ;
EnregBtnCourant ;
end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.BAbandonClick(Sender: TObject);
begin
Modalresult:=mrCancel ;
end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.MODNATUREChange(Sender: TObject);
Var sPlus : String ;  //NA debut 06/11/02
{$IFNDEF PGIS5}
    tt    : String ;  //XMG 05/10/01
{$ENDIF}              //NA fin 06/11/02
begin
{$IFDEF PGIS5}
ModMODELE.DataType:='GCIMPMODELEFO' ;
sPlus := 'MO_NATURE="' ;
if MODNature.Value = 'T' then sPlus := sPlus + NatureModTicket else
 if MODNature.Value = 'E' then sPlus := sPlus + NatureModTicEtat else
  if MODNature.Value = 'L' then sPlus := sPlus + NatureModTicDoc ;
sPlus := sPlus + '" AND MO_TYPE="' + MODNature.Value + '"' ;
ModMODELE.Plus := sPlus ;
{$ELSE}
//XMG debut 05/10/01
DonneTablette('MODMOD;'+MODNATURE.Value+';',Caisse,tt,splus) ;        //NA 13/03/02
ModMODELE.DataType:=tt ; //'TTMODELEETAT' ;
ModMODELE.plus:=splus ; //'MO_TYPE="T" AND MO_NATURE="'+MODNature.Value+'"' ;
//XMG fin 05/10/01
{$ENDIF}
EnregBtnCourant ;
end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
if Key = VK_VALIDE   then
   Begin
   BValider.Click ;
   Key:=0 ;
   End ;
if (Key=VK_ESCAPE) then
   Begin
   BAbandon.click ;
   Key:=0 ;
   End ;
end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.BFontChange(Sender: TObject);
Var St : String ;
begin
St:=BFont.Font.Name+';' ;
if fsBold in BFont.Font.Style then St:=St+'B' ;
if fsItalic in BFont.Font.Style then St:=St+'I' ;
if fsUnderline in BFont.Font.Style then St:=St+'U' ;
if fsStrikeOut in BFont.Font.Style then St:=St+'S' ;
St:=St+';'+IntToStr(BFont.Font.Size)+';' ;
St:=St+IntToStr(ord(BFont.Font.Color))+';' ;
if BFont.Alignment=taLeftjustify then St:=St+'L;' ;
if BFont.Alignment=taRightjustify then St:=St+'R;' ;
if BFont.Alignment=taCenter then St:=St+'C;' ;
BtnCourant.PutValue('CE_POLICE', St) ;
PnlBoutons.UpdateBouton(BtnCourant) ;
end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.bAssistantClick(Sender: TObject);
Var sSQL    : string ;
begin
////if (cbConcept.itemindex=1) and (ArtCode.Text='') then
if (cbConcept.Value='1') and (ArtCode.Text='') then
   begin
   sSQL:=AGLLanceFiche('MFO','ASSISTCLAVECR', '', '', 'TABLE=ARTICLE;'+ArtWhere.Text) ;   //NA : 30/07/01
   if sSql <> '' then ArtWhere.Text:=sSQL ;
   end ;
{$IFDEF CHR} //debut 17/11/03
if (cbConcept.Value='14') then
   begin
   sSQL:=AGLLanceFiche('MFO','ASSISTCLAVECR', '', '', 'TABLE=HRDOSRES;'+ResWhere.Text) ;   //NA : 30/07/01
   if sSql <> '' then ResWhere.Text:=sSQL ;
   end;
{$ENDIF} //NA fin 17/11/03
end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.BCopierClick(Sender: TObject);  
Var FromCaisse : String ;
begin
FromCaisse:='' ;
{$IFDEF PGIS5}
FromCaisse:=OkRecopie(Caisse) ; 
{$ELSE}
FromCaisse:=OkRecopie(CE_CAISSE.Text) ;
{$ENDIF}
if trim(FromCaisse)<>'' then PnlBoutons.RecopieClavier(FromCaisse) ;
end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.BPropietesClick(Sender: TObject);
begin
CentreTW(self,TWPropietes) ;
enabled:=FALSE ;
cbClcVisible.Visible:=True; //NA 17/11/03
cbClcVisible.Checked:=PnlBoutons.ClcVisible ; //NA 17/11/03
RgCalculettea.ItemIndex:=ord(PnlBoutons.ClcPosition) ;
FBtnsLarge.value:=PnlBoutons.NbrBtnWidth ;
FBtnsHaut.value:=PnlBoutons.NbrBtnHeight ;
TWPropietes.visible:=TRUE ;
end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.BValideTWClick(Sender: TObject);
begin
CacheTW(TWPropietes) ;
PnlBoutons.NbrBtnWidth:=FBtnsLarge.Value ;
PnlBoutons.NbrBtnHeight:=FBtnsHaut.Value ;
PnlBoutons.ClcPosition:=TPosClc(RGCalculettea.ItemIndex) ;
PnlBoutons.ClcVisible:=cbClcVisible.Checked ; //NA 17/11/03
end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TFParamKB.CacheTW(TW : TToolWindow97 )  ;
Begin
if assigned(TW) then TW.visible:=FALSE ;
Enabled:=TRUE ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.BAnnulerTWClick(Sender: TObject);
begin
if (Sender is TToolbarbutton97) and (TToolBarButton97(Sender).parent is TToolWindow97) then CacheTW(TToolWindow97(TToolBarButton97(Sender).parent)) ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.BFirstClick(Sender: TObject);
begin
ChangePage(0) ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.BLastClick(Sender: TObject);
begin
ChangePage(PnlBoutons.MaxPages) ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.BPrevClick(Sender: TObject);
begin
ChangePage(PnlBoutons.PageCourante-1) ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.BNextClick(Sender: TObject);
begin
ChangePage(PnlBoutons.PageCourante+1) ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TFParamKB.MajNav ;
Begin
BFirst.Enabled:=PnlBoutons.PageCourante>0 ;
BPrev.Enabled:=PnlBoutons.PageCourante>0 ;
BNext.Enabled:=PnlBoutons.PageCourante<PnlBoutons.MaxPages+1 ;  //Creation de nouvelle page.
BLast.Enabled:=PnlBoutons.PageCourante<PnlBoutons.MaxPages ;
LNoPage.Caption:=TraduireMemoire('Page ')+IntToStr(PnlBoutons.PageCourante+1);  //NA 11/02/04
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.FImageChange(Sender: TObject);
var pp : TPicture ;
begin
if TWImages.visible then BApercuimClick(nil) ;
if assigned(BApercuIm.Glyph) then BApercuIm.Glyph.ReleaseHandle ;
if (FImage.Text='') or (FImage.Value < FImage.MinValue) or (FImage.Value > FImage.MaxValue) then Exit ;
pp:=TPicture.Create ;
try
  PnlBoutons.Images.GetBitmap(FImage.Value-1,pp.Bitmap) ;
  BApercuIm.Glyph.assign(pp.Bitmap) ; 
 finally
  pp.free ;
 End ;
Gestionpardefautcontrol(Sender) ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.FImageExit(Sender: TObject);
begin
if (FImage.Text='') or (FImage.Value < FImage.MinValue) or (FImage.Value > FImage.MaxValue) then
   Begin
   if (FImage.Text='') or (FImage.Value < FImage.MinValue) then FImage.Value:=FImage.MinValue
                                                           else FImage.Value:=FImage.MaxValue ;
   if FImage.CanFocus then FImage.SetFocus ;
   Exit ;
   End ;
Gestionpardefautcontrol(Sender) ;
end;
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.LwImagesDblClick(Sender: TObject);
begin
FImage.Value:=Valeuri(LWImages.Selected.Caption) ;
TWImages.visible:=FALSE ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.BApercuImClick(Sender: TObject);
begin
LWImages.items[FImage.Value-1].Selected:=TRUE ;
LWImages.items[FImage.Value-1].MakeVisible(FALSE) ;
centreTW(Self,TWImages) ;
TWImages.Visible:=TRUE ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.BBoutonPosClick(Sender: TObject);
begin
CentreTW(self,TWBouton) ;
enabled:=FALSE ;
ChargePosBouton ;
TWBouton.visible:=TRUE ;
if FPage.Canfocus then FPage.SetFocus else FLigne.SetFocus ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TFParamKB.ChangePage(NewPage : Integer ) ;
Var Btn : TOB ;  //NA 11/02/04
Begin
If Assigned(PnlBoutons) then PnlBoutons.PageCourante:=NewPage ;
MajNav ;
If Assigned(BtnCourant) then //NA debut 11/02/04
begin
  Btn:=BtnCourant.Parent.FindFirst(['CE_PAGE','CE_NUMERO'],[NewPage,BtnCourant.GetValue('CE_NUMERO')],FALSE) ;
  if Assigned(Btn) then ParamBouton(Btn) ;
end;  //NA fin 11/02/04
End ;
/////////////////////////////////////////////////////////////////////////////////////////
Procedure  TFParamKB.SetParamsCE ( AParams : String  ) ;
Var St : String ;
Begin
if AParams=ParamsCE then exit ;
FParamsCE:=AParams ;
if Trim(FOldParamsCE)='' then FOldParamsCE:=ParamsCE ;
St:=ParamsCE ;
if (trim(St)='') or (NbCarInString(St,';')<3) then St:=inttostr(CENbrBtnWidthDef)+';'+inttostr(CENbrBtnHeightDef)+';'+inttostr(ord(pcLeft))+';X;' ; //NA debut 17/11/03
PnlBoutons.NbrBtnWidth:=valeuri(readtokenst(St))  ;
PnlBoutons.NbrBtnHeight:=Valeuri(ReadTokenst(St)) ;
PnlBoutons.ClcPosition:=tPosClc(Valeuri(readtokenst(St))) ;
PnlBoutons.ClcVisible:=(ReadTokenst(St)<>'-') ; //NA fin 17/11/03
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TFParamKB.ChargePosBouton ;
var Pg : Integer ;
Begin
FPage.MaxValue:=PnlBoutons.MaxPages+2 ;
FLigne.MaxValue:=PnlBoutons.NbrBtnHeight ;
FCol.Maxvalue:=PnlBoutons.NbrBtnWidth ;
Pg:=VInteger(BtnCourant.GetValue('CE_PAGE')) ;
TFPage.Visible:=(pg>-1) ;
TFPage.FocusCOntrol.visible:=TFPage.Visible ;
if pg<>-1 then FPage.Value:=pg+1 ;
FLigne.Value:=VInteger(BtnCourant.GetValue('Ligne'))+1 ;
FCol.Value:=VInteger(BtnCourant.GetValue('Colonne'))+1 ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.BValideBtnClick(Sender: TObject);
var Err  : TPGIErrS1 ; //XMG 07/01/04
    page : Integer ;
    Btn  : TOB ;
begin
CacheTW(TWBouton) ;
if ((FPage.Value-1<>VInteger(BtnCourant.GetValue('CE_PAGE'))) and (VInteger(BtnCourant.GetValue('CE_PAGE'))<>-1)) or
   (FLigne.Value-1<>VInteger(BtnCourant.GetValue('Ligne')))  or
   (FCol.Value-1<>VInteger(BtnCourant.GetValue('Colonne')))    then
   Begin
   Page:=FPage.Value-1 ;
   if VInteger(BtnCourant.GetValue('CE_PAGE'))=-1 then Page:=-1 ;
   If (not PnlBoutons.ExisteBouton(Page,FLigne.Value-1,FCOl.Value-1)) or (PGIAsk(MsgErrDefautS1(788),Caption)=mrYes) then //XMG 07/01/04
      Begin
      Btn:=BtnCourant ; //BtnCourant sera recharge pendant le deplacement du bouton.
      BtnCourant:=nil ;
      PnlBoutons.MoveBouton(VInteger(Btn.GetValue('CE_PAGE')),VInteger(Btn.GetValue('CE_NUMERO')),page,FLigne.value-1,FCol.value-1,Err) ;
      if err.Code<>0 then PGIInfo(Err.Libelle,Caption) ;
      MajNav ;
      End ;
   End ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.BCopieSocClick(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}          //MR debut 23/10/02
exit ;
{$ELSE}
if CopieClavierFromSoc(Self,'2;3;4;5;6;10;11;12;') then
   Begin
   //ParamsCE:=inttostr(PnlBoutons.NbrBtnWidth)+';'+inttostr(PnlBoutons.NbrBtnHeight)+';'+inttostr(ord(PnlBoutons.ClcPosition))+';' ; //NA debut 17/11/03
   ParamsCE:=inttostr(PnlBoutons.NbrBtnWidth)+';'+inttostr(PnlBoutons.NbrBtnHeight)+';'+inttostr(ord(PnlBoutons.ClcPosition))+';'
     +TrueFalseSt(PnlBoutons.ClcVisible)+';' ; //NA fin 17/11/03
   ChangePage(0) ;
   End ;
{$ENDIF}                     //MR fin 23/10/02
end;
/////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.FSFSousFamClick(Sender: TObject);
var Tt,Pl : String ;
begin
{$IFDEF PGIS5}
{$ELSE}
DonneTablette('FSF;'+TRUEFALSESt(FSFSousFam.Checked)+';',Caisse,tt,Pl) ;     //NA 13/03/02
{$ENDIF}
if (tt+pl<>'') and ((tt<>FSFFSF.DataType) or (pl<>FSFFSF.plus)) then
   Begin
   FSFFSF.DataType:=tt ;
   FSFFSF.Plus:=pl ;
   EnregBtnCourant ;
   End ;
end;
/////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.BDynamiqueClick(Sender: TObject);
var Params : String ;
    Err    : TPGIErrS1 ; //XMG 07/01/04
    tpv    : TSetTypePave ;
begin
if ConfirmeCreationPave('',TRUE) then
   Begin
   Params:=inttostr(PnlBoutons.NbrBtnWidth)+';'+inttostr(PnlBoutons.NbrBtnHeight)+';'+inttostr(ord(PnlBoutons.ClcPosition))+';' ;
{$IFDEF PGIS5} //NA debut 17/11/03
   tpv:=[tpEncaissement, tpFamilles] ;
{$ELSE}
   tpv:=[tpEncaissement] ;
   if VString(GetParamSoc('SO_TYPEPAVEDEFAUT'))<>'000' then tpv:=tpv+[tpFamilles] ;
{$ENDIF}  //NA fin 17/11/03
   if not CreePaveDynamique(PnlBoutons,Caisse,Params,tpv,TRUE,Err) then PGIBox(Err.Libelle,Caption) else
      Begin // on recharge la page courante
      BtnCourant:=nil ;
      PnlBoutons.PageCourante:=PnlBoutons.PageCourante ;
      End ;
   End ;
end;
//////////////////////////////////////////////////////////////////////////////////////////////////
//XMG 15/04/02 début
procedure TFParamKB.FPVFONChange(Sender: TObject);
begin
FPVPage.Enabled:=(FPVFon.Value='GOT') ;
FPVPage.Value:=FPVPage.Value*ord(FPVPage.Enabled) ;
Gestionpardefautcontrol(Sender) ;
end;
//////////////////////////////////////////////////////////////////////////////////////////////////
//NA 10/05/04 début
procedure TFParamKB.BloqueFiche;
begin
DisableControlsFiche(Self,cbConcept,(cbConcept.Itemindex=0),FALSE) ;
GS.Enabled := True;
cbConceptListe.Enabled := True;
end;
//NA 10/05/04 début
//////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.pagesChange(Sender: TObject);
var concept : String ;
    Ok      : Boolean ;
begin
//concept:=Pages.ActivePage.Hint ; Ok:=FALSE ; //NA debut 12/03/03
case Pages.ActivePageIndex of
  0  : concept:='VIDE' ;
  1  : concept:='ART' ;
  2  : concept:='COM' ;
  3  : concept:='REM' ;
  4  : concept:='VEN' ;
  5  : concept:='FON' ;
  6  : concept:='REG' ;
  7  : concept:='OCT' ;
  8  : concept:='OFC' ;
  9  : concept:='RET' ;
  10 : concept:='MOD' ;
  11 : concept:='MIP' ;
  12 : concept:='FPV' ;
  13 : concept:='FSF' ;
{$IFDEF CHR}
  14 : concept:='RES' ; //NA 17/11/03
{$ENDIF}
  end ;
Ok:=FALSE ; //NA fin 12/03/03
if Concept='MOD' then ModNature.Value:=ModNature.Values[ModNature.Values.count-1] else
if Concept='FPV' then Begin if FPVFon.ItemIndex<>0 then FPVFon.ItemIndex:=0 ; FPVFonChange(FPVFon) ; Ok:=TRUE ; end else

FBlocage.Checked:=Ok ;
rbLibelle.Checked:=TRUE ;
EnablelesRB(Concept) ;
BloqueFiche; //NA 10/05/04
etablirFAllera(not Ok) ;
EnregBtnCourant ;
end;
//////////////////////////////////////////////////////////////////////////////////
//XMG 15/04/02 fin

//NA debut 20/09/2002
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.BReproFormeClick(Sender: TObject); //NA debut 17/11/03
begin
  if (Assigned(BtnCourant)) and (BReproForme.Down) then
  begin
    BtnReproForme := TOB.Create('CLAVIERECRAN', nil, -1);
    BtnReproForme.Dupliquer(BtnCourant, True, True);
  end else
  begin
    BtnReproForme.Free;
    BtnReproForme := nil;
    BReproForme.Down := False;
  end;
end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFParamKB.ReproduitMiseEnForme(UnBouton: TOB);
begin
  if (Assigned(UnBouton)) and (Assigned(BtnReproForme)) then
  begin
    BtnCourant.PutValue('CE_COULEUR', VInteger(BtnReproForme.GetValue('CE_COULEUR')));
    BtnCourant.PutValue('CE_POLICE', BtnReproForme.GetValue('CE_POLICE'));
  end;
end;

procedure TFParamKB.BValideListeClick(Sender: TObject);
var
  Ind: integer;
  sConcept, sTablette, sPlus, sTable, sCode, sWhere, sPrfx, sLib, sSql, sColName: string;
  TOBMain: TOB;
begin
  Ind := 0;
  if IsNumeric(cbConceptListe.Value) then Ind := StrToInt(cbConceptListe.Value);
  case Ind of
    1  : sConcept:='ART' ;
    2  : sConcept:='COM' ;
    3  : sConcept:='REM' ;
    4  : sConcept:='VEN' ;
    5  : sConcept:='FON' ;
    6  : sConcept:='REG' ;
    7  : sConcept:='OCT' ;
    8  : sConcept:='OFC' ;
    9  : sConcept:='RET' ;
    10 : sConcept:='MOD' ;
    11 : sConcept:='MIP' ;
    12 : sConcept:='FPV' ;
    13 : sConcept:='FSF' ;
{$IFDEF CHR}
    14 : sConcept:='RES' ;
{$ENDIF}
    else sConcept:='VIDE' ;
  end ;
  DonneTablette(sConcept, Caisse, sTablette, sPlus);
  GetCorrespType(sTablette, sTable, sCode, sWhere, sPrfx, sLib);
  if sTable = '' then Exit;
  sColName := sCode + ',' + sLib;
  if sConcept = 'REG' then sColName := sColName + ',' + 'MP_TYPEMODEPAIE';
 {$IFDEF PGIS5} //NA debut 10/05/04
  if sConcept = 'ART' then
  begin
    if sWhere <> '' then sWhere := sWhere + ' AND ';
    sWhere := sWhere + 'GA_STATUTART<>"DIM"';
  end;
  {$ENDIF} //NA fin 10/05/04
  sSql := 'SELECT ' + sColName + ' FROM ' + sTable;
  if sWhere <> '' then
  begin
    ChangeWhereTT(sWhere, sPlus, False);
    sSql := sSql + ' WHERE ' + sWhere;
  end;
  if (sConcept = 'ART') and (cbChampListe.Value <> '') and (cbValeurListe.Value <> '') then //NA 10/05/04
  begin
    if sWhere = '' then sSql := sSql + ' WHERE ' else sSql := sSql + ' AND ';
    sSql := sSql + cbChampListe.Value + '="' + cbValeurListe.Value + '"';
  end;
  TOBMain := TOB.Create('', nil, -1);
  TOBMain.LoadDetailFromSQL(sSql);
  sColName := sCode + ';' + sLib;
  if sConcept = 'REG' then sColName := sColName + ';' + 'MP_TYPEMODEPAIE';
  TOBMain.PutGridDetail(GS, False, False, sColName, True);
  TOBMain.Free;
  if (sConcept <> 'ART') and (sConcept <> 'COM') and (sConcept <> 'VEN') then
  begin
    GS.ColWidths[0] := 0;
    if (sConcept = 'REG') and (GS.ColCount >= 2) then GS.ColWidths[2] := 0;
  end else
  begin
    GS.ColWidths[0] := 70;
  end;
  HMTrad.ResizeGridColumns(GS);
end;

procedure TFParamKB.BListeClick(Sender: TObject);
begin
  cbConceptListe.Values.Clear ;
  cbConceptListe.Items.Clear ;
  cbConceptListe.Values.Add('1') ;
  cbConceptListe.Items.Add(Traduirememoire('Article')) ;
  cbConceptListe.Values.Add('2') ;
  cbConceptListe.Items.Add(Traduirememoire('Commentaire')) ;
  cbConceptListe.Values.Add('3') ;
  cbConceptListe.Items.Add(Traduirememoire('Remise')) ;
  cbConceptListe.Values.Add('4') ;
  cbConceptListe.Items.Add(Traduirememoire('Vendeur')) ;
  cbConceptListe.Values.Add('5') ;
  cbConceptListe.Items.Add(Traduirememoire('Fonction')) ;
  cbConceptListe.Values.Add('6') ;
  cbConceptListe.Items.Add(Traduirememoire('Règlement')) ;
  cbConceptListe.Values.Add('7') ;
  cbConceptListe.Items.Add(Traduirememoire('Opérations de caisse')) ;
  cbConceptListe.Values.Add('12') ;
  cbConceptListe.Items.Add(Traduirememoire('Fonctions spéciales')) ;
{$IFDEF PGIS5}
{$ELSE}
  cbConceptListe.Values.Add('8') ;
  cbConceptListe.Items.Add(Traduirememoire('Opération financière')) ;
  cbConceptListe.Values.Add('9') ;
  cbConceptListe.Items.Add(Traduirememoire('Retour article')) ;
  cbConceptListe.Values.Add('10') ;
  cbConceptListe.Items.Add(Traduirememoire('Modèle d''impression')) ;
  cbConceptListe.Values.Add('11') ;
  cbConceptListe.Items.Add(Traduirememoire('Mode d''impression')) ;
  cbConceptListe.Values.Add('13') ;
  cbConceptListe.Items.Add(Traduirememoire('Famille/SousFamille')) ;
{$ENDIF}
{$IFDEF CHR}
  cbConceptListe.Values.Add('14') ;
  cbConceptListe.Items.Add(Traduirememoire('Ressources')) ;
{$ENDIF}
  Enabled := True;
  TWListe.Visible := True;
  if cbConceptListe.Canfocus then cbConceptListe.SetFocus;
end;

procedure TFParamKB.BCancelListeClick(Sender: TObject);
begin
  CacheTW(TWListe);
end;

procedure TFParamKB.BInsereListeClick(Sender: TObject);
var
  Stg, sCode, sType: string;
  Im: integer;
begin
  if not Assigned(BtnCourant) then Exit;
  if cbConceptListe.ItemIndex < 0 then Exit;
  if (GS.Col < GS.FixedCols) and (GS.Col >= GS.ColCount) then Exit;
  if (GS.Row < GS.FixedRows) and (GS.Row >= GS.RowCount) then Exit;
  sCode := GS.Cells[0, GS.Row];
  case StrToInt(cbConceptListe.Value) of
    0  : Stg:='VIDE;' ; //on ne devrait jamais arriver ici.
    1  : Stg:='ART;'+sCode+';;;;' ;
    2  : Stg:='COM;'+sCode+';' ;
    3  : Stg:='REM;;'+sCode+';' ;
    4  : Stg:='VEN;'+sCode+';' ;
    5  : Stg:='FON;'+sCode+';' ;
    6  : Stg:='REG;'+sCode+';' ;
    7  : Stg:='OCT;'+sCode+';;' ;
    8  : Stg:='OFC;'+sCode+';;' ;
    9  : Stg:='RET;'+sCode+';;;' ;
    10 : Stg:='MOD;'+sCode+';;' ;
    11 : Stg:='MIP;'+sCode+';' ;
    12 : Stg:='FPV;'+sCode+';;' ;
    13 : Stg:='FSF;'+sCode+';-;' ;
{$IFDEF CHR}
    14 : Stg:='RES;'+sCode+';;' ;
{$ENDIF}
  end;
  BtnCourant.PutValue('CE_TEXTE', Stg);
  //Stg := #255+#255; //NA 06/01/04
  Stg := GetKBIDCaptionStd; //NA 06/01/04
  BtnCourant.PutValue('CE_CAPTION', Stg) ;
  if cbConceptListe.Value = '6' then
  begin
    Im := 0;
    if GS.ColCount >= 2 then sType := GS.Cells[2, GS.Row];
{$IFDEF PGIS5}
    if sType = TYPEPAIEESPECE   then Im := 26 else // Espéce
    if sType = TYPEPAIECHEQUE   then Im := 29 else // Cheque
    if sType = TYPEPAIECHQDIFF  then Im := 29 else // Cheque differe
    if sType = TYPEPAIECB       then Im := 42 else // Carte blue
    if sType = TYPEPAIEBONACHAT then Im := 12 else // Bon d'achat
    if sType = TYPEPAIEARRHES   then Im := 38 else // Arrhes
    if sType = TYPEPAIEAVOIR    then Im := 8  else // Avoirs
    if sType = TYPEPAIERESTEDU  then Im := 41 ;    // Reste dû
{$ELSE}
{$ENDIF} //XMG 07/01/04
    if Im > 0 then
    begin
      Stg := '**INTERNE**;' + IntToStr(Im) + ';';
      BtnCourant.PutValue('CE_IMAGE', Stg);
    end;
  end;
  ReproduitMiseEnForme(BtnCourant);
  ChargeBtnCourant;
end;

procedure TFParamKB.cbChampListeChange(Sender: TObject);
var
  sChamp: string;
begin
  sChamp := cbChampListe.Value;
  if sChamp <> '' then cbValeurListe.DataType := Get_Join(sChamp);
end;

procedure TFParamKB.cbConceptListeChange(Sender: TObject);
begin
  if cbConceptListe.Value = '1' then
  begin
    cbChampListe.Visible := True;
    cbChampListe.Enabled := True;
    cbValeurListe.Visible := True;
    cbValeurListe.Enabled := True;
    PListeTete.Height := cbChampListe.Top + cbChampListe.Height + 5;
  end else
  begin
    cbChampListe.Visible := False;
    cbChampListe.Enabled := False;
    cbValeurListe.Visible := False;
    cbValeurListe.Enabled := False;
    GS.Top := cbChampListe.Top;
    PListeTete.Height := BValideListe.Top + BValideListe.Height + 5;
  end;
end; //NA fin 17/11/03

// JTR - Image libre stockées
procedure TFParamKB.FImageLibreBaseChange(Sender: TObject); 
begin
//  RecupImageLibreBase;
end;

procedure TFParamKB.FImageLibreBaseClick(Sender: TObject);
begin
  RecupImageLibreBase;
end;

procedure TFParamKB.FImageLibreBaseElipsisClick(Sender: TObject);
begin
  RecupImageLibreBase;
end;

procedure TFParamKB.RecupImageLibreBase;
var LeChoix : string;
begin
{$IFDEF PGIS5}
  LeChoix := AGLLanceFiche('YY','YYLIENSOLE','CE','','DEFAUT=PHJ;SELECTBYDBCLICK=TRUE;PHOTOBTN='+FImageLibreBase.Text+';FORCE=TRUE');
  if Lechoix <> '' then
  begin
    BtnCourant.PutValue('CE_IMAGE','**LIENSOLE**;'+ReadTokenSt(LeChoix));                     //LO_TABLEBLOB
    BtnCourant.PutValue('CE_IMAGE',BtnCourant.GetValue('CE_IMAGE')+';'+ReadTokenSt(LeChoix)); //LO_IDENTIFIANT
    BtnCourant.PutValue('CE_IMAGE',BtnCourant.GetValue('CE_IMAGE')+';'+ReadTokenSt(LeChoix)); //LO_RANGBLOB
    FImageLibreBase.Text := ReadTokenSt(LeChoix); //LO_LIBELLE
    BtnCourant.PutValue('CE_IMAGE',BtnCourant.GetValue('CE_IMAGE')+';'+FImageLibreBase.Text); //LO_LIBELLE
    EnregBtnCourant
  end;
{$ENDIF PGIS5}
end;
// Fin JTR - Image libre stockées

{$IFDEF PGIS5}
function FOAGLParam_ClavierEcran ( parms: array of variant; nb: integer ) : variant ; //NA debut 17/11/03
Var Action   : TActionFiche ;
    ParamsCE : String ;
    {$IFNDEF EAGL}
    DrawXP: boolean;
    {$ENDIF}
BEGIN
{$IFNDEF EAGL}
  DrawXP:=V_PGI.DrawXP;
  V_PGI.DrawXP:=False;
{$ENDIF}
Result:='' ;
ParamsCE:=vString(Parms[1]) ;
Action:=TActionFiche(vInteger(Parms[2])) ;
Param_ClavierEcran(Parms[0], ParamsCE, Action) ;
if Action in [taCreat,taModif] then Result:=ParamsCE ;
{$IFNDEF EAGL}
  V_PGI.DrawXP:=DrawXP;
{$ENDIF}
END ; //NA fin 17/11/03
//NA fin 20/09/2002
procedure TFParamKB.TWListeResize(Sender: TObject); //NA debut 01/06/04
begin
HMTrad.ResizeGridColumns(GS);
end; //NA fin 01/06/04

initialization
RegisterAglFunc('FOParam_ClavierEcran', FALSE, 3, FOAGLParam_ClavierEcran) ;
{$ENDIF}

end.

