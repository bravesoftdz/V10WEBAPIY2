Unit Lettrage;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, Hctrls, StdCtrls, Buttons, ExtCtrls, SaisUtil, LettUtil, HTB97,
  Menus, HSysMenu, hmsgbox, Mask, ComCtrls,
  Ent1,SaisComm, LettVisu, LettRegu,DelVisuE,
  UtilSoc,  // pour le ParamSociete
  CPCHANCELL_TOF,
  DEVISE_TOM,
{$IFDEF EAGLCLIENT}
  MaineAGL,
  CPJUSTISOL_TOF,
{$ELSE}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  DB,
  FE_Main,
  Filtre,
  JustiSol,
{$ENDIF}
{$IFDEF VER150}
   Variants,
 {$ENDIF}
{$IFNDEF IMP}
  Saisie,
  SaisBor,
  UTOFCPMULMVT,
{$ENDIF}
  LetBatch,
  UtilPGI,
  ed_tools,LettAuto,
  //LG*
  ULibEcriture,
  UTOB,
  ULibWindows,
  UtilSais,
  HEnt1,
  ParamSoc,
  HPanel, // pour le THPanel
  UIUtil, // pour le findInsidePanel
  //AGLInit, // pour theTOB
  CPJournal_TOM,
  CPGeneraux_TOM,
  CPTiers_TOM,
  UTOFECRLET, TntStdCtrls, TntButtons, TntGrids
  ;

Procedure LettrageManuel ( RL : RLETTR ; OkLettrage : boolean ; Action : TActionFiche ; vQ : TQuery = nil ; vStNatAux : string = '' ; vInTypeMvt : integer = 0 ; vBoFromConsult : boolean = false) ;
Procedure LettrageEnSaisieTOB ( RL : RLETTR ; vTOB : TOB ; OkLettrage : boolean  ; Action : TActionFiche) ;
{$IFDEF CCMP}
Procedure LettrageManuelMP ( RL : RLETTR ; OkLettrage : boolean ; Action : TActionFiche ; Qui : tProfilTraitement) ;
{$ENDIF}

type
  TFLettrage = class(TForm)
    DockTop: TDock97;
    DockRight: TDock97;
    DockLeft: TDock97;
    DockBottom: TDock97;
   PCouverture     : TPanel;
    PFenCouverture : TPanel;
    GCouvDev       : TGroupBox;
    H_TitreCouv    : TLabel;
    H_MontantD     : TLabel;
    H_COUVERTURED  : TLabel;
    H_RESTED       : TLabel;
    H_TOTCOUVDD    : TLabel;
    E_TAUXDEV      : THNumEdit;
    E_MONTANTS: THNumEdit;
    E_COUVERTURES: THNumEdit;
    E_RESTES: THNumEdit;
    E_TOTCOUVDEBS: THNumEdit;
    H_TOTCOUVDC    : TLabel;
    E_TOTCOUVCREDS: THNumEdit;
    BCValide: THBitBtn;
    BCAbandon: THBitBtn;
    GD             : THGrid;
    GC             : THGrid;
    HLettre        : THMsgBox;
    HDivL          : THMsgBox;
    CJOURNAL       : THValComboBox;
    CNATUREPIECE   : THValComboBox;
    CMODEPAIE      : THValComboBox;
    CDevise        : THValComboBox;
    POPS           : TPopupMenu;
    FindLettre     : TFindDialog;
    POPZ: TPopupMenu;
    BZoomEcriture: THBitBtn;
    BZoomCpte: THBitBtn;
    BChancel: THBitBtn;
    HMTrad: THSystemMenu;
    BJustif: THBitBtn;
    Valide97: TToolbar97;
    POutils: TToolbar97;
    BAnnulerSel: TToolbarButton97;
    BRajoute: TToolbarButton97;
    BCombi: TToolbarButton97;
    BUnique: TToolbarButton97;
    BChercher: TToolbarButton97;
    BMenuZoom: TToolbarButton97;
    BEnleve: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BAide: TToolbarButton97;
    ToolbarSep972: TToolbarSep97;
    GeneValide: TToolbarButton97;
    BAjouteEnleve: TToolbarButton97;
    BValide: TToolbarButton97;
    BValideSelect: TToolbarButton97;
    BTF: TToolbarButton97;
    Pages: TPageControl;
    TComptes: TTabSheet;
    TParams: TTabSheet;
    ISigneEuro: TImage;
    CParCouverture: TCheckBox;
    H_REGUL: THLabel;
    E_REGUL: THValComboBox;
    H_ECART: THLabel;
    E_ECART: THValComboBox;
    CParEnsemble: TCheckBox;
    TG_GENERAL: THLabel;
    TT_AUXILIAIRE: THLabel;
    HG_GENERAL: THLabel;
    HT_AUXILIAIRE: THLabel;
    TSoldes: TLabel;
    LESOLDE: THNumEdit;
    BvSolde: TBevel;
    CParExcept: TCheckBox;
    HTitreL: TLabel;
    H_CODEL: TLabel;
    TE_REFINTERNE: THLabel;
    E_REFLETTRAGE: TEdit;
    BDelettre: TToolbarButton97;
    TDATEGENERATION: THLabel;
    DATEGENERATION: THCritMaskEdit;
    BAgrandir: TToolbarButton97;
    BReduire: TToolbarButton97;
    PCDetail: TPageControl;
    TSDetail: TTabSheet;
    TSDetail1: TTabSheet;
    PDetail: TPanel;
    BSepare: TBevel;
    H_JOURNAL: TLabel;
    E_JOURNAL: TLabel;
    H_NATUREPIECE: TLabel;
    E_NATUREPIECE: TLabel;
    H_DATEECHEANCE: TLabel;
    E_DATEECHEANCE: TLabel;
    H_MODEPAIE: TLabel;
    E_MODEPAIE: TLabel;
    H_LIBELLE: TLabel;
    E_LIBELLE: TLabel;
    H_REFLIBRE: TLabel;
    E_REFRELEVE: TLabel;
    H_TOTAL: TLabel;
    H_JOURNAL_: TLabel;
    E_JOURNAL_: TLabel;
    H_NATUREPIECE_: TLabel;
    E_NATUREPIECE_: TLabel;
    H_LIBELLE_: TLabel;
    E_LIBELLE_: TLabel;
    H_DATEECHEANCE_: TLabel;
    E_DATEECHEANCE_: TLabel;
    H_MODEPAIE_: TLabel;
    E_MODEPAIE_: TLabel;
    H_REFLIBRE_: TLabel;
    E_REFRELEVE_: TLabel;
    H_TOTAL_: TLabel;
    E_TOTAL: THNumEdit;
    E_TOTAL_: THNumEdit;
    BCalcul: TToolbarButton97;
    PopCalcul: TPopupMenu;
    Lettragesurpassagedesolde1: TMenuItem;
    Rapprochementcombinatoire1: TMenuItem;
    Rapprochementsimple1: TMenuItem;
    Rapprochementsurrfrence1: TMenuItem;
    Lettragesurpassagesoldenul1: TMenuItem;
    POPDeLett: TPopupMenu;
    DeLettPartiel: TMenuItem;
    DELettTotal: TMenuItem;
    PCSoldes: TPageControl;
    TSSoldePME: TTabSheet;
    TSSoldePCL: TTabSheet;
    SL_SOLDEDEBIT: THNumEdit;
    H_INFODEBIT: THLabel;
    SL_TOTALDEBIT: THNumEdit;
    SL_TOTALCREDIT: THNumEdit;
    H_INFOCREDIT: THLabel;
    SL_SOLDECREDIT: THNumEdit;
    SL_SOLDEDEBITPCL: THNumEdit;
    SL_CUMULDEBIT: THNumEdit;
    SL_CUMULCREDIT: THNumEdit;
    SL_SOLDECREDITPCL: THNumEdit;
    SL_SOLDE: THNumEdit;
    h1: THLabel;
    h2: THLabel;
    Label1: TLabel;
    e_refrelevepcl: TLabel;
    e_naturepcl: TLabel;
    Label3: TLabel;
    Label7: TLabel;
    e_affairepcl: TLabel;
    Label8: TLabel;
    Label4: TLabel;
    Label2: TLabel;
    e_dateecheancepcl: TLabel;
    e_reflibrepcl: TLabel;
    e_refexternepcl: TLabel;
    Label6: TLabel;
    Label5: TLabel;
    E_ETABLISSEMENTPCL: TLabel;
    E_CONTREPARTIEGENPCL: TLabel;
    BForceVadide: TToolbarButton97;
    BTri: TToolbarButton97;
    PopForceValide: TPopupMenu;
    Soldeautomatique1: TMenuItem;
    Lettragepartiel1: TMenuItem;
    Passagedunecrituresimplifi1: TMenuItem;
    POPPCL: TPopupMenu;
    Validerlelettrage1: TMenuItem;
    Prlettrerlaslection1: TMenuItem;
    N1: TMenuItem;
    Lettragepartiel2: TMenuItem;
    Passagedunecrituresimplifie1: TMenuItem;
    Soldeautomatique2: TMenuItem;
    N2: TMenuItem;
    Dfairelaselectionencours1: TMenuItem;
    Rechercherdanslaliste1: TMenuItem;
    Enleverlestotalementslettrs1: TMenuItem;
    Remettrelestotalementslettres1: TMenuItem;
    Agrandirlaliste1: TMenuItem;
    Rduirelaliste1: TMenuItem;
    POPAnnulSelect: TMenuItem;
    POPAnnulTotal: TMenuItem;
    Zoom1: TMenuItem;
    Lettragesurpassagedesolde2: TMenuItem;
    Rapprochementcombinatoire2: TMenuItem;
    Rapprochementsurlesmontants1: TMenuItem;
    Rapprochementssurlesrfrences1: TMenuItem;
    Lettragesurpassagedesoldenul1: TMenuItem;
    Zoom2: TMenuItem;
    Voirlcriture1: TMenuItem;
    Voirlecompte1: TMenuItem;
    Tabledechancellerie1: TMenuItem;
    Dtailsdeschances1: TMenuItem;
    Lettrageparexception1: TMenuItem;
    Recherchedanslacomptabilit1: TMenuItem;
    PMessLettAuto: TPanel;
    Label9: TLabel;
    Panel2: TPanel;
    GroupBox1: TGroupBox;
    Label12: TLabel;
    Label13: TLabel;
    THNBP: TLabel;
    HNBS: THNumEdit;
    HNBN: THNumEdit;
    HNBP: THNumEdit;
    EBStop: TToolbarButton97;
    SL_SOLDEPCL: THNumEdit;
    Label11: TLabel;
    e_modepaiepcl: TLabel;
    BPresentationPGE: TToolbarButton97;
    BPresentationPCL: TToolbarButton97;
    BTriDate: TToolbarButton97;
    BTriCode: TToolbarButton97;
    POPAUX: TMenuItem;
    POPAuxSuiv: TMenuItem;
    POPAUXPREC: TMenuItem;
    BDown: TToolbarButton97;
    BUP: TToolbarButton97;
    LEX: TMenuItem;
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure MontreDetail(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
    procedure SauteZero(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
    procedure ClickSelect(Sender: TObject);
    procedure GDMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure ClickActive(Sender: TObject);
    procedure BValideClick(Sender: TObject);
    procedure BAbandonClick(Sender: TObject);
    procedure POPSPopup(Sender: TObject);
    procedure BZoomCpteClick(Sender: TObject);
    procedure BZoomEcritureClick(Sender: TObject);
    procedure BAnnulerSelClick(Sender: TObject);
    procedure BEnleveClick(Sender: TObject);
    procedure BRajouteClick(Sender: TObject);
    procedure BChercherClick(Sender: TObject);
    procedure BChancelClick(Sender: TObject);
    procedure BCombiClick(Sender: TObject);
    procedure BValideSelectClick(Sender: TObject);
    procedure E_REGULChange(Sender: TObject);
    procedure H_REGULDblClick(Sender: TObject);
    procedure H_ECARTDblClick(Sender: TObject);
    procedure E_ECARTChange(Sender: TObject);
    procedure CParCouvertureClick(Sender: TObject);
    procedure CParEnsembleClick(Sender: TObject);
    procedure E_COUVERTURESChange(Sender: TObject);
    procedure BCAbandonClick(Sender: TObject);
    procedure BCValideClick(Sender: TObject);
    procedure FindLettreFind(Sender: TObject);
    procedure BUniqueClick(Sender: TObject);
    procedure BJustifClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure BMenuZoomclick(Sender: TObject);
    procedure BMenuZoomMouseEnter(Sender: TObject);
    procedure GeneValideClick(Sender: TObject);
    procedure BAjouteEnleveClick(Sender: TObject);
    procedure BTFClick(Sender: TObject);
    procedure BDelettreClick(Sender: TObject);
    procedure DATEGENERATIONExit(Sender: TObject);
    procedure BAgrandirClick(Sender: TObject);
    procedure BReduireClick(Sender: TObject);
    procedure Rapprochementsurrfrence1Click(Sender: TObject);
    procedure Lettragesurpassagedesolde1Click(Sender: TObject);
    procedure DELettTotalClick(Sender: TObject);
    procedure DeLettPartielClick(Sender: TObject);
    procedure BCalculMouseEnter(Sender: TObject);
    procedure BForceVadideClick(Sender: TObject);
    procedure Soldeautomatique1Click(Sender: TObject);
    procedure Lettragepartiel1Click(Sender: TObject);
    procedure Rapprochementsimple1Click(Sender: TObject);
    procedure Passagedunecrituresimplifi1Click(Sender: TObject);
    procedure Rapprochementcombinatoire1Click(Sender: TObject);
    procedure Lettragesurpassagesoldenul1Click(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure POPPCLPopup(Sender: TObject);
    procedure Validerlelettrage1Click(Sender: TObject);
    procedure Lettrageparexception1Click(Sender: TObject);
    procedure Recherchedanslacomptabilit1Click(Sender: TObject);
    procedure EBStopClick(Sender: TObject);
    procedure Rapprochementsurlesmontants1Click(Sender: TObject);
    procedure BPresentationPGEClick(Sender: TObject);
    procedure BPresentationPCLClick(Sender: TObject);
    procedure POPAUXSuivClick(Sender: TObject);
    procedure POPAUXPRECClick(Sender: TObject);
    procedure BUPClick(Sender: TObject);
    procedure BDownClick(Sender: TObject);
  private
    Client,ModeSelection,LettrageTotP,LettrageTotCur : boolean ;
    PieceModifiee,EcartChange,PasEcart,AvecRegul,ErreurReseau,AvecConvert : boolean ;
    LTE_TotDebP,LTE_TotCredP,LTE_TotDebD,LTE_TotCredD,LTE_TotDebCur,LTE_TotCredCur : Double ;
    LTE_CouvDebP,LTE_CouvCredP,LTE_CouvDebD,LTE_CouvCredD,LTE_CouvDebCur,LTE_CouvCredCur : Double ;
    EcartRegulDevise,EcartChangeFranc,EcartChangeEuro,ConvertFranc,ConvertEuro : Double ;
    ModeCouverture,FindFirst,DelettrageGlobal,AvecTL,ExistTvaEnc,IncoherECC,AvantEuro,TraiteReturn : boolean ;
    DatePaquetMax,DatePaquetMin,GeneNow,LaDateGenere : TDateTime ;
    JREGUL,JECART                     : REGJAL ;
    DEV,DEVF                          : RDEVISE ;
    RL                                : RLETTR ;
    GSais                             : THGrid ;
    LigSais                           : integer ;
    TLAD,TLAC,TTSOLU                  : TLIST ;
    DernLettrage,CodeL,RefReleve,SauveCode,LeEtab : String ;
    NbSelD,NbSelC,CurGrid,IndiceRegul,CurLig,NbReleve,GridRappro,CurDec : integer ;
    ColDebit,ColCredit,ColLetD,ColLetC,MaxNumSol : integer ;
    WMinX,WMinY : Integer ;
    Qui : tProfilTraitement ;
    //LG*
    FBoUneGrille            : boolean;  // si True -> presentation du lettrage dans une seule grille
    FBoAucunJournal         : boolean; // boolean indiquant l'existance de journaux de regul
    FStSensCompte           : string; // sens du compte D : debiteur   C : crediteur M : mixte
    ColSolde                : integer; // indice de la colonnne solde progrssif
    FIFL                    : RINFOSLETT; // info de lettrage
    FLDAnnul                : TList; // Liste d'annulation pour la grille des debits
    FLCAnnul                : TList; // Liste d'annulation pour la grille des credits
    FStLastError            : string; // Dernier message d'erreur delphi
    FBoOter                 : boolean; // si True on enleve les totalements lettres de l'affichage
    FEcrRegul               : TEcrRegul; // TOB stockant les ecritures de regul
    FInLastNumero           : integer;
    FStLastDateComptable    : TDateTime;
    FStLastJournal          : string;
    FStLastLibelle          : string;
    FStLastGeneral          : string;
    FTOBEcr                 : TOB;
    FInOrdreTri             : integer; // 1 : tri par date comptable  2 : tri par code lettre
    FBoFermer               : boolean; // true si une erreur c'est produite dans le create de la fiche -> la fenetre est ferme
    //FZListJournal           : TZListJournal; // objet de gestion des journaux
    FInfoEcr                : TInfoEcriture;
    FBoChangementMode       : boolean ;
    FBoAuxSuiv              : boolean ;
    FBoDessiner             : boolean ;
    FBoAucunEnr             : boolean ;
    FBoFromSaisie           : boolean ;
    FStAnnulLettrage        : HTStringList ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    Procedure GetCellCanvasD(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
    Procedure GetCellCanvasC(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
    procedure HGPostDrawCell(ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
// Requêtes, validations fichier
    function  WhereCptes : String ;
    procedure LettrageFichier ;
    procedure GenerePiecesRegul ;
    procedure RemplirLeM ( Var M : RMVT ; Quoi : PG_LETTRAGE) ;
    procedure GetCodeSuivant ;
    procedure DevalideCode ;
    procedure UpdateDevalide ;
    procedure DelettrageCourant ;
// Lettrage mémoire
    function  LettrageMemoire : boolean;
    procedure LettrageEpuise ;
    procedure ClickValide ;
    procedure ProposeValide ;
    procedure FinirLettrage ( G : THgrid ) ;
    procedure FinirDeLettrage ( G : THgrid ) ;
    procedure TraiteLeEtat ( O : TOBM ) ;
    function  Validation(DMin,DMax,DEche : tDateTime ; ModePaie : string) : boolean;
    function  SoldeAutomatique(ChoixValid : string ; DMin,DMax,DEche : tDateTime ; ModePaie : string): boolean;
// Init et defauts
    procedure InitEcran (Tout : boolean ) ;
    procedure GereNature ;
    procedure InitialiseLettrage ;
    procedure ShowLettrage ( Tout : boolean ) ;
    function AttribTitre : string;
//    procedure AvertDevise ;
    procedure GereEnabled ;
// Lettrage auto
    procedure EnvoieLaSauce ( Solde : double ; Unique : boolean ) ;
    function  MontantLigne ( G : THGrid ; i : integer ) : double ;
    procedure ChargeElimine ( Var LM : T_D ; Var LP,LG : T_I ; Var NbD,NoSol : integer ) ;
    procedure ArchiveSolution ( LP,LG : T_I ; NoSol : integer ; Var Multi : boolean ) ;
    procedure AjouteSol ( TT : TList ; PosG,NoSol : integer ; Ident : TSolLet ) ;
    function  JeGardeElement ( G : THGrid ; Lig,NoSol : integer ) : boolean ;
    Procedure AfficheLettAutoEnCours(HNM : THNumedit ; Val : Integer) ;
    procedure ExploreUnivers ( Solde : double ; NoSol : integer ; Unique : boolean ) ;
    procedure CompleteListe ( Var LM : T_D ; Var LP,LG : T_I ; Var NbD : integer ; NoSol,ii : integer ) ;
    function  NouvelleDonne ( LG : T_I ; NbD : integer ) : boolean ;
    function  Pareil ( Var LG,LT : T_I ; NbD : integer ) : boolean ;
    procedure InitLesTLA ;
    procedure Calcul(vInTypeCalcul : integer) ;
// Entête
    function  ChargeInfosRegul ( Var J : REGJAL ; Value : String ; Controle : boolean  ) : boolean ;
// Couverture
    procedure SaisieCouverture ( O : TOBM ) ;
    procedure CacheCouverture ;
    function  EquilibreCouverture : boolean ;
// Affichages
    procedure ActiveGrid ;
    procedure GerelesValide ;
    procedure CentreGrids ;
    procedure SwapGrids ;
    procedure EnleveOuRajoute( Enlev : boolean ) ;
    procedure SwapPanels ;
    procedure DetailLigne ( Lig : integer ; Debit : boolean ) ;
    procedure AfficheLesSoldes ;
//    procedure VoirPiecesGenerees ;
    procedure RemettreGrilleAJour(G : THGrid ; Oter : boolean );
    procedure Desectionner(G : THGrid);
// Calculs
    procedure CalculDebitCredit ( Propose : boolean ) ;
    procedure CodeDateLettre(Var DMin,DMax, DEche : tDateTime ; var ModePaie : string ) ;
    procedure RazValeurs ;
    Function  GetCouvCur ( O : TOBM ) : Double ;
    Function  GetDebitCur ( O : TOBM ) : Double ;
    Function  GetCreditCur ( O : TOBM ) : Double ;
// Selections
    //LG* 05/02/2002 nouveau param
    procedure Selection ( Next : boolean; Parle : boolean=true; LigneAvecRegul : boolean=false) ;
    procedure PrepareSelect ( G : THGrid ) ;
    procedure Inverseselection ( G : THGrid ; Lig : integer ) ;
    function  EstLettre ( G : THGrid ; Lig : integer ; Totale : boolean ) : boolean ;
    function  JePrends ( G : THGrid ; Lig : integer ) : boolean ;
    procedure PrepareLettrageSauf ;
    //LG* 05/01/2001
    procedure PasseMoisSuivant;
    procedure PasseLibelleSuivant( vStLibelle : string = '');
    procedure PasseJourSuivant;
// Chargements
    procedure AttribColonnes ( G : THGrid ) ;
    function  ChargeMouvements : boolean ;
    procedure RempliGrids ( Q : TQuery ) ;
    function  EcheSaisie ( O : TOBM ) : boolean;
// Algo lettrage
    procedure PreLettreSolution ( NumSol : integer ) ;
    procedure VireCouverture ;
// Ecarts de change
    procedure GereEcartDeChange ;
    procedure GereRegul ;
//    procedure GereConvert(DMin,DMax : tDateTime) ;
    procedure CalculEcartRegulDevise ;
// Barre outils
    procedure ClickZoom ;
    procedure ClickZoomPiece ;
    procedure ClickAnnulerSel ;
    procedure ClickEnleve(Oter : boolean) ;
    procedure ClickEnlevePME(Oter : boolean) ;
    procedure ClickCherche ;
    procedure ClickChancel ;
    procedure ClickCombi ( Unique : boolean ) ;
    procedure ClickToutSel(AvecLet : boolean = false; MemeMvt : boolean = false; MemeMois : boolean = false);
    procedure AjouteEcrGrille(Sender: TOB);
    // gestion des erreurs
    procedure AnnulationSurErreur( vStFonction : string = '');
    procedure AnnulLettrage;
    procedure AnnulLigneLettrage;
    procedure InitLettrage;
    function  ControleLigne: boolean;
    procedure ChangePresentation( vModeUneGrille : boolean );
    procedure ChangerAux(vBoSuiv: boolean);
  public
    FQ           : TQuery ;
    FBoFromTOB   : boolean ;
  //  FTOB         : TOB ;
    FStNatAux    : string ;
    FInTypeMvt   : integer ;
    OkLettrage   : boolean ;
    Action       : TActionFiche ;

    property TOBEcr : TOB read FTOBEcr write FTOBEcr ;

  end;


implementation

uses
  {$IFDEF MODENT1}
  CPProcGen,
  CPProcMetier,
  CPTypeCons,
  {$ENDIF MODENT1}
  uTofRechercheEcr,ulibexercice; // CPLanceFiche_CPRechercheEcr(True)

{$R *.DFM}
const
 cDeuxLigne=#10#13#10#13;


{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 13/12/2001
Modifié le ... : 30/06/2004
Description .. : LG Modif : utilise le smode inside
Suite ........ : - LG - 30/06/2004 - FB 12989 12960 - on utilise la meme 
Suite ........ : requet épour la consultation et le lettrage
Mots clefs ... : 
*****************************************************************}
Procedure LettrageManuel ( RL : RLETTR ; OkLettrage : boolean ; Action : TActionFiche ; vQ : TQuery = nil ; vStNatAux : string = '' ; vInTypeMvt : integer = 0 ; vBoFromConsult : boolean = false) ;
Var X : TFLettrage ;
    OkRes : boolean ;
    PP : THPanel;
BEGIN
OkRes:=True ; if RL.Appel=tlSaisieTreso then OkRes:=False ;
if ((Action<>taConsult) and (OkRes) and (OkLettrage)) then
   BEGIN
   if _Blocage(['nrCloture','nrBatch','nrEnca','nrDeca'],True,'nrLettrage') then Exit ;
   END ;
X:=TFLettrage.Create(Application) ;
PP:=FindInsidePanel;
 Try
  X.Qui:=prAucun ; X.FBoFromTOB := false ;
  X.RL:=RL  ; X.OkLettrage:=OkLettrage ; X.Action:=Action ; X.FQ := vQ ; X.FStNatAux := vStNatAux ; X.FInTypeMvt := vInTypeMvt ;
  X.FBoFromSaisie := vBoFromConsult ;
  if X.RL.LettrageDevise then X.DEV.Code:=X.RL.DeviseMvt else X.DEV.Code:=V_PGI.DevisePivot ;
  if PP=nil then
  begin
    try
      X.ShowModal ;
    finally
      X.Free ;
    end ;
  end else
  begin
    InitInside(X,PP);
    X.Show;
  end;

 Finally
    if ((Action<>taConsult) and (OkRes) and (OkLettrage)) then _Bloqueur('nrLettrage',False) ;
 End ;

Screen.Cursor:=SyncrDefault ;
END ;


Procedure LettrageEnSaisieTOB ( RL : RLETTR ; vTOB : TOB ; OkLettrage : boolean  ; Action : TActionFiche) ;
var
 X : TFLettrage ;
 OkRes : boolean ;
 PP : THPanel;
BEGIN
OkRes:=True ; if RL.Appel=tlSaisieTreso then OkRes:=False ;
if ((Action<>taConsult) and (OkRes) and (OkLettrage)) then
   BEGIN
   if _Blocage(['nrCloture','nrBatch','nrEnca','nrDeca'],True,'nrLettrage') then Exit ;
   END ;

if not CControlePresenceLettrage(vTOB) then
 begin
  PGIInfo('Impossible de lettrer, la totalité des mouvements n''est pas incluse dans la sélection !','Lettrage');
  exit ;
 end ;// if

X:=TFLettrage.Create(Application) ;
PP:=FindInsidePanel;
 Try
  X.Qui:=prAucun ; X.FBoFromTOB := true ; X.TOBEcr := vTOB ;
  X.RL:=RL  ; X.OkLettrage:=OkLettrage ; X.Action:=Action ;
  if X.RL.LettrageDevise then X.DEV.Code:=X.RL.DeviseMvt else X.DEV.Code:=V_PGI.DevisePivot ;
  if PP=nil then
  begin
    try
      X.ShowModal ;
    finally
      X.Free ;
    end ;
  end else
  begin
    InitInside(X,PP);
    X.Show;
  end;

 Finally
    if ((Action<>taConsult) and (OkRes) and (OkLettrage)) then _Bloqueur('nrLettrage',False) ;
 End ;

Screen.Cursor:=SyncrDefault ;
END ;




{$IFDEF CCMP}
Procedure LettrageManuelMP ( RL : RLETTR ; OkLettrage : boolean ; Action : TActionFiche ; Qui : tProfilTraitement) ;
Var X : TFLettrage ;
    OkRes : boolean ;
BEGIN
If Not CptDansProfil(RL.General,RL.Auxiliaire,Qui) Then Exit ;
OkRes:=True ; if RL.Appel=tlSaisieTreso then OkRes:=False ;
If Not MonProfilOk(Qui) Then
  BEGIN
  if ((Action<>taConsult) and (OkRes) and (OkLettrage)) then
     BEGIN
     if _Blocage(['nrCloture','nrBatch','nrEnca','nrDeca'],True,'nrLettrage') then Exit ;
     END ;
  END ;
X:=TFLettrage.Create(Application) ;
 Try
  X.Qui:=Qui ;
  X.RL:=RL  ; X.OkLettrage:=OkLettrage ; X.Action:=Action ;
  if X.RL.LettrageDevise then X.DEV.Code:=X.RL.DeviseMvt else X.DEV.Code:=V_PGI.DevisePivot ;
  X.ShowModal ;
 Finally
  X.Free ;
  if ((Action<>taConsult) and (OkRes) and (OkLettrage)) then
    If Not MonProfilOk(Qui) Then _Bloqueur('nrLettrage',False) ;
 End ;
Screen.Cursor:=SyncrDefault ;
END ;
{$ENDIF}

Function TFLettrage.GetCouvCur ( O : TOBM ) : Double ;
BEGIN
if RL.LettrageDevise then Result:=O.GetMvt('E_COUVERTUREDEV') else
Result:=O.GetMvt('E_COUVERTURE') ;
END ;

Function TFLettrage.GetDebitCur ( O : TOBM ) : Double ;
BEGIN
if RL.LettrageDevise then Result:=O.GetMvt('E_DEBITDEV') else
Result:=O.GetMvt('E_DEBIT') ;
END ;

Function TFLettrage.GetCreditCur ( O : TOBM ) : Double ;
BEGIN
if RL.LettrageDevise then Result:=O.GetMvt('E_CREDITDEV') else
Result:=O.GetMvt('E_CREDIT') ;
END ;

{================================== REQUETES ===========================================}
{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 27/11/2002
Modifié le ... : 29/11/2002
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TFLettrage.WhereCptes : String ;
BEGIN
{OKOK}
Result:='E_GENERAL="'+RL.General+'" AND E_AUXILIAIRE="'+RL.Auxiliaire+'"' ;
if RL.CritDev<>'' then Result:=Result+' AND E_DEVISE="'+RL.CritDev+'"' ;
if Not OkLettrage then Result:=Result+' AND E_LETTRAGE="'+RL.CodeLettre+'"' ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 26/04/2002
Modifié le ... :   /  /
Description .. : reaffectation des couvertures pour un lettrage total pour 
Suite ........ : supprimer les pb de conversion.
Mots clefs ... : 
*****************************************************************}
procedure TFLettrage.LettrageFichier ;
Type TCODEETAT = RECORD
                 CodeL,EtatL : String ;
                 END ;
Var O : TOBM ;
    G : THgrid ;
    i,k: integer ;
    StEtat,LCode,EtatL : String ;
    GenerePieces : boolean ;
//    lTOB : TOB ;
    lDateModif : TDateTime ;
BEGIN
{OKOK}
if Action=taConsult then Exit ; GenerePieces:=False ; FStLastError:='';
if FBoUneGrille then
 BEGIN
  FBoOter:=false ; GD.BeginUpdate ; RemettreGrilleAJour(GD,FBoOter) ; GD.EndUpdate ;
 END;
for k:=1 to 2 do
 BEGIN
    if k=1 then G:=GD else G:=GC ;
    for i:=1 to G.RowCount-1 do
        BEGIN
        O:=GetO(G,i) ; if O=Nil then Break ;
        if (O.Etat<>Modifie) or (O.GetMvt('PASENREGISTRER')='X') or (O.GetMvt('E_CONTROLEUR')='X') then Continue ;
        LCode:=O.GetMvt('E_LETTRAGE') ; if Length(LCode)>4 then BEGIN LCode:=Copy(LCode,1,4) ; O.PutMvt('E_LETTRAGE',LCode) ; END ;
        StEtat:=O.GetMvt('E_ETAT') ;
        // reaffectation de la couverture pour un lettrage total pour supprimer les ecartde de conversion en contrevaleur, en change en tout...
        if O.GetValue('E_ETATLETTRAGE')='TL' then
         begin
          O.PutValue('E_COUVERTURE',Arrondi(Abs(O.GetValue('E_DEBIT')-O.GetValue('E_CREDIT')),V_PGI.OkDecV)) ;
          O.PutValue('E_COUVERTUREDEV',Arrondi(Abs(O.GetValue('E_DEBITDEV')-O.GetValue('E_CREDITDEV')),V_PGI.OkDecV)) ;
         end; // if
        if ((OkLettrage) and (PasTouche(StEtat))) then
         GenerePieces:=True ;
        EtatL:=O.GetMvt('E_ETATLETTRAGE') ; lDateModif:=RL.NowStamp ;
        if Not GoReqMajLet(O,RL.General,RL.Auxiliaire,lDateModif,OkLettrage) then BEGIN V_PGI.IOError:=oeLettrage ; Exit ; END ;
      {  if O.GetValue('E_QUALIFPIECE') = 'N' then
         begin
          lTOB:=TOB.Create('',RL.TOBResult,-1) ;
          lTOB.Dupliquer(TOB(O),true,true) ;
          lTOB.PutValue('E_DATEMODIF',lDateModif) ;
         end ;  }
        END ;
 END ;
//LG* 03/01/2002
if ((OkLettrage) and (GenerePieces)) then GenerePiecesRegul ;
try
 if FBoUneGrille and (FIFL.ChoixValid<>'AL5') then
  begin
   if not FEcrRegul.Save then
    V_PGI.IOError:=oeLettrage ;
  end;
except
 on E:Exception do BEGIN FStLastError:=E.Message ; V_PGI.IOError:=oeSaisie END ;
end;
PieceModifiee:=False ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 13/12/2001
Modifié le ... : 23/09/2002
Description .. : Utilisation du FZListJournal pour recherche les info sur le 
Suite ........ : journal
Suite ........ : 
Suite ........ : -LG - 22/04/2002 - la date comptable est egale à la date 
Suite ........ : paquet max
Suite ........ : 
Suite ........ : -LG-23/09/2002-le test sur l'existance des journaux de regul 
Suite ........ : etaient inverse
Mots clefs ... : 
*****************************************************************}
procedure TFLettrage.RemplirLeM ( Var M : RMVT ; Quoi : PG_LETTRAGE) ;
Var Facturier : String3 ;
BEGIN
if Action=taConsult then Exit ;
FillChar(M,Sizeof(M),#0) ;
Case Quoi of
   pglRegul   : BEGIN
                M.Jal:=E_REGUL.Value ; Facturier:=JREGUL.Facturier ; M.ModeSaisieJal:=JREGUL.ModeSaisie ;
                if RL.LettrageDevise then BEGIN M.CodeD:=DEV.Code ; M.TauxD:=DEV.Taux ; END
                                     else BEGIN M.CodeD:=V_PGI.DevisePivot ; M.TauxD:=1 ; END ;
                M.Nature:='OD' ;
                END ;
   pglEcart   : BEGIN
                M.Jal:=E_ECART.Value ; Facturier:=JECART.Facturier ;
                M.CodeD:=V_PGI.DevisePivot ; M.TauxD:=1 ;
                M.Nature:='ECC' ;
                END ;
   pglConvert,
   pglInverse :
               BEGIN
             {   M.Jal:=VH^.JalEcartEuro ;
               if FInfoEcr.Journal.Load([VH^.JalEcartEuro])=-1 then
                 BEGIN
                  PGIInfo('Le journal d''écart de conversion euro ' + VH^.JalEcartEuro + ' n''existe pas en base !',Caption) ;
                  V_PGI.IoError:=oeSaisie ; exit ;
                 END; // if }
                Facturier:=FInfoEcr.Journal.GetValue('J_COMPTEURNORMAL') ; M.ModeSaisieJal:=FInfoEcr.Journal.GetValue('J_MODESAISIE') ;
                if RL.LettrageDevise then BEGIN M.CodeD:=DEV.Code ; M.TauxD:=DEV.Taux ; END
                                     else BEGIN M.CodeD:=V_PGI.DevisePivot ; M.TauxD:=1 ; END ;
                M.Nature:='OD' ;
                END ;
 END;
//SetIncNum(EcrGen,Facturier,M.Num) ;
M.Axe:='' ;
if LeEtab<>'' then M.Etabl:=LeEtab else M.Etabl:=VH^.EtablisDefaut ;
M.DateTaux:=LaDategenere ; M.Simul:='N' ; M.Valide:=False ;
M.ANouveau:=False ;
if FBoUneGrille then BEGIN M.DateC:=DatePaquetMax ; M.Exo:=QuelExo(DateToStr(DatePaquetMax)) ; END  
else BEGIN M.DateC:=LaDateGenere ; M.Exo:=QuelExo(DateToStr(M.DateC)) ; END ;
if ctxExercice.EstExoClos(M.Exo) then
 begin
  M.DateC := ctxExercice.EnCours.Deb ;
  M.Exo   := ctxExercice.EnCours.Code ;
 end ;
M.General:='' ;
SetIncNum(EcrGen,Facturier,M.Num,M.DateC) ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 03/01/2002
Modifié le ... : 27/11/2002
Description .. : -Message delphi en cas d'erreur
Suite ........ : -Remplacement de OLDDEBIT par ECARTREGUL
Suite ........ : 
Suite ........ : -27/11/2002 - affectation datepaquet min et max, les ecart
Suite ........ : de conversion etaient genere sans exercice
Mots clefs ... : 
*****************************************************************}
procedure TFLettrage.GenerePiecesRegul ;
Var CodesUtil  : Array[1..100] of String[5] ;
    i,NbUtil,j,k,Col,Ind : integer ;
    StL        : String[5] ;
    O          : TOBM ;
    OOO        : TOBM ;
    G          : THGrid ;
    StEtat     : String ;
    M          : RMVT ;
    Cpte       : String ;
    Montant    : double ;
    Okok       : boolean ;
    NumPiece   : Longint ;
    XX         : T_ECARTCHANG ;
BEGIN
if Action=taConsult then Exit ;
if Not OkLettrage then Exit ;
if RevisionActive(LaDateGenere) then Exit ;
FillChar(CodesUtil,Sizeof(Codesutil),#0) ; NbUtil:=0 ; OOO := nil ;
//LG* 03/01/2002 mettre gestion des erreurs
for k:=1 to 2 do
    BEGIN
    if k=1 then BEGIN G:=GD ; Col:=ColLetD ; END else BEGIN G:=GC ; Col:=ColLetC ; END ;
    for i:=1 to G.RowCount-1 do
        BEGIN
        FillChar(XX,Sizeof(XX),#0) ;
        StL:=Copy(G.Cells[Col,i],1,4) ; if StL='' then Continue ;
        Okok:=False ; for j:=1 to NbUtil do if StL=CodesUtil[j] then Okok:=True ; if Okok then Continue ;
        O:=GetO(G,i) ; if O=Nil then Break ; StEtat:=O.GetMvt('E_ETAT') ; if Not PasTouche(StEtat) then Continue ;
        DatePaquetMin:=O.GetMvt('E_DATEPAQUETMIN') ; DatePaquetMax:=O.GetMvt('E_DATEPAQUETMAX') ;
        NumPiece:=O.GetMvt('E_NUMEROPIECE') ;
        XX.RefLettrage:=E_REFLETTRAGE.Text ;
        //LG* on ne gere plus les ecritures de cette façon
        if ((StEtat[1]='X') and not FBoUneGrille) or ( (StEtat[1]='X') and (FIFL.ChoixValid='AL5') and FBoUneGrille ) then {regul}
           BEGIN
           RemplirLeM(M,pglRegul) ; Ind:=StrToInt(StEtat[4]) ;
           Montant:=O.GetMvt('OLDDEBIT') ;
           if Montant>0 then Cpte:=JREGUL.D[Ind] else Cpte:=JREGUL.C[Ind] ;
           XX.Cpte:=Cpte ; XX.Montant1:=Montant ; XX.Regul:=pglRegul ; XX.Quotite:=DEV.Quotite ;
           XX.Ref:=Copy(HDivL.Mess[3]+' N° '+Inttostr(NumPiece),1,35) ;
           XX.Lib:=Copy(HDivL.Mess[4]+' N° '+Inttostr(NumPiece),1,35) ;
           XX.DPMin:=DatePaquetMin ; XX.DPMax:=DatePaquetMax ; XX.Decimale:=DEV.Decimale ;

           OOO := CreerPartieDoubleLett(M,RL,XX,O) ;

           END ;
        if StEtat[2]='X' then {ecart change}
           BEGIN
           RemplirLeM(M,pglEcart) ;
           Montant:=O.GetMvt('OLDCREDIT') ;
           if Montant>0 then Cpte:=DEV.CptDebit else Cpte:=DEV.CptCredit ;
           XX.Cpte:=Cpte ; XX.Montant1:=Montant ; XX.Regul:=pglEcart ;
           XX.Montant2:=O.GetMvt('RATIO') ;
           XX.Decimale:=DEV.Decimale ; XX.Quotite:=DEV.Quotite ;
           XX.Ref:=Copy(HDivL.Mess[5]+' N° '+Inttostr(NumPiece),1,35) ;
           XX.Lib:=Copy(HDivL.Mess[6]+' N° '+Inttostr(NumPiece),1,35) ;
           XX.DPMin:=DatePaquetMin ; XX.DPMax:=DatePaquetMax ;

           OOO := CreerPartieDoubleLett(M,RL,XX,O) ;
   
           END ;
        if StEtat[2]='#' then {ecart conversion}
           BEGIN
           RemplirLeM(M,pglConvert) ;
           XX.Montant1:=O.GetMvt('CONVERTEURO') ; XX.Montant2:=O.GetMvt('CONVERTFRANC') ;
        //   if (XX.Montant1>0) or ((XX.Montant1=0) and (XX.Montant2>0)) then XX.Cpte:=VH^.EccEuroDebit else XX.Cpte:=VH^.EccEuroCredit ;
           XX.Regul:=pglConvert ;
           XX.Decimale:=DEV.Decimale ; XX.Quotite:=DEV.Quotite ;
           XX.Ref:=Copy(HDivL.Mess[15]+' N° '+Inttostr(NumPiece),1,35) ;
           XX.Lib:=Copy(HDivL.Mess[16]+' N° '+Inttostr(NumPiece),1,35) ;
           XX.DPMin:=DatePaquetMin ; XX.DPMax:=DatePaquetMax ;

           OOO := CreerPartieDoubleLett(M,RL,XX,O) ;
  
           END ;
        if StEtat[2]='&' then {ecart inverse}
           BEGIN
           RemplirLeM(M,pglInverse) ;
           XX.Montant1:=O.GetMvt('CONVERTEURO') ; XX.Montant2:=O.GetMvt('CONVERTFRANC') ;
//           if (XX.Montant1>0) or ((XX.Montant1=0) and (XX.Montant2>0)) then XX.Cpte:=VH^.EccEuroDebit else XX.Cpte:=VH^.EccEuroCredit ;
           XX.Regul:=pglInverse ;
           XX.Decimale:=DEV.Decimale ; XX.Quotite:=DEV.Quotite ;
           XX.Ref:=Copy(HDivL.Mess[15]+' N° '+Inttostr(NumPiece),1,35) ;
           XX.Lib:=Copy(HDivL.Mess[16]+' N° '+Inttostr(NumPiece),1,35) ;
           XX.DPMin:=DatePaquetMin ; XX.DPMax:=DatePaquetMax ;

           OOO := CreerPartieDoubleLett(M,RL,XX,O) ;
  
           END ;
        Inc(NbUtil) ; CodesUtil[NbUtil]:=StL ;
        END ;
    END ;
 FreeAndNil(OOO) ;
END ;


{=============================== PRELETTRE SOL =========================================}
{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 17/01/2002
Modifié le ... :   /  /
Description .. : Modif pour la gestion d'une seule grille
Mots clefs ... :
*****************************************************************}
procedure TFLettrage.PreLettreSolution ( NumSol : integer ) ;
Var i : integer ;
BEGIN
{OKOK}
if Action=taConsult then Exit ;
if NumSol<=0 then Exit ;
for i:=0 to TLAD.Count-1 do
    BEGIN
    if ((P_LAS(TLAD[i]).S[NumSol]<>slPasSolution) and (Not EstSelect(GD,i+1))) then InverseSelection(GD,i+1) ;
    END ;
for i:=0 to TLAC.Count-1 do
    BEGIN
    if not FBoUneGrille then
    BEGIN if ((P_LAS(TLAC[i]).S[NumSol]<>slPasSolution) and (Not EstSelect(GC,i+1))) then InverseSelection(GC,i+1) ; END
    else BEGIN if ((P_LAS(TLAC[i]).S[NumSol]<>slPasSolution) and (Not EstSelect(GD,i+1))) then InverseSelection(GD,i+1) ; END
    END ;
CalculDebitCredit(False) ;
if ModeSelection then ClickValide ;
END ;

{================================== ENTETE =============================================}
{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 27/12/2001
Modifié le ... :   /  /
Description .. : Suppression d'un warning
Mots clefs ... :
*****************************************************************}
function TFLettrage.MontantLigne ( G : THGrid ; i : integer ) : double ;
Var O : TOBM ;
BEGIN
{OKOK}
Result:=0 ; O:=GetO(G,i) ; if O=Nil then Exit ;
if FBoUneGrille then BEGIN if GridRappro=2 then Result:=Arrondi(GetDebitCur(O),CurDec) else Result:=Arrondi(GetCreditCur(O),CurDec) ; END
else if G=GD then Result:=Arrondi(GetDebitCur(O),CurDec) else Result:=Arrondi(GetCreditCur(O),CurDec) ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 16/01/2002
Modifié le ... : 17/01/2002
Description .. : modif de la gestion pour l'utilisation d'une seule grille qui 
Suite ........ : stocke les debits et les credits.
Suite ........ : En mode pme la grille passé en parametre ne contient que 
Suite ........ : des debit ou des credit.
Suite ........ : On parcourd la grille et si GridRappro=2 on ne s'intèresse 
Suite ........ : qu'un Debit et si GridRappro=1 on s'interesse au credit
Mots clefs ... : 
*****************************************************************}
function TFLettrage.JeGardeElement ( G : THGrid ; Lig,NoSol : integer ) : boolean ;
Var TT   : TList ;
    Sol  : integer ;
    Okok : boolean ;
    ii   : TSolLet ;
    O    : TOBM;
BEGIN
{OKOK}
//LG*
if FBoUneGrille then
BEGIN
JeGardeElement:=false ;
O:=GetO(G,Lig) ; if O=Nil then exit ; if (GridRappro=2) and (O.GetMvt('E_CREDIT')=0) then exit ;
if (GridRappro=1) and (O.GetMvt('E_DEBIT')=0) then exit ;
END;
JeGardeElement:=True ; if NoSol<=0 then Exit ;
if FBoUneGrille then BEGIN if GridRappro=2 then TT:=TLAC else TT:=TLAD ; END
else if G=GD then TT:=TLAD else TT:=TLAC ;
Okok:=True ;
for Sol:=1 to NoSol do
    BEGIN
     ii:=P_LAS(TT[Lig-1]).S[Sol] ;
     if ii<>slPasSolution then BEGIN Okok:=False ; Break ; END ;
    END ;
if Okok then Exit ;
JeGardeElement:=False ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 16/01/2002
Modifié le ... :   /  /
Description .. : Modif pour la gestion d'une seule grille -> on se sert
Suite ........ : uniquement de GD
Mots clefs ... :
*****************************************************************}
procedure TFLettrage.ChargeElimine ( Var LM : T_D ; Var LP,LG : T_I ; Var NbD,NoSol : integer ) ;
Var i : integer ;
    G : THGrid ;
    O : TOBM ;
BEGIN
{OKOK}
//LG* 16/01/2002
FillChar(LM,Sizeof(LM),#0) ; FillChar(LP,Sizeof(LP),#0) ; FillChar(LG,Sizeof(LG),#0) ;
if FBoUneGrille then G:=GD else if GridRappro=2 then G:=GC else G:=GD ;
NbD:=0 ;
for i:=1 to G.RowCount-2 do if ((Not EstSelect(G,i)) and (Not EstLettre(G,i,True))) then
 if i<=MaxDroite then
    if JeGardeElement(G,i,NoSol) then
    BEGIN
    if FBoUneGrille then
     BEGIN
      O:=GetO(G,i) ; if O=Nil then Exit ;
      if GridRappro=1 then LM[NbD]:=Arrondi(GetDebitCur(O),CurDec) else LM[NbD]:=Arrondi(GetCreditCur(O),CurDec);
     END
      else LM[NbD]:=MontantLigne(G,i) ;
    LP[NbD]:=0 ; LG[NbD]:=i ; Inc(NbD) ;
    if NbD>MaxDroite-1 then Break ;
    END ;
if NbD>MaxDroite-1 then NbD:=MaxDroite-1 ;
END ;

procedure TFLettrage.AjouteSol ( TT : TList ; PosG,NoSol : integer ; Ident : TSolLet ) ;
BEGIN
{OKOK}
if PosG>TT.Count then Exit ; if PosG<=0 then Exit ;
P_LAS(TT[PosG-1]).S[NoSol]:=Ident ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 17/01/2002
Modifié le ... :   /  /
Description .. : modif pour la gestion d'une seule grille
Mots clefs ... :
*****************************************************************}
procedure TFLettrage.ArchiveSolution ( LP,LG : T_I ; NoSol : integer ; Var Multi : boolean ) ;
Var i,NbItems : integer ;
    ii : TSolLet ;
BEGIN
{OKOK}
//LG* 17/01/2002
if Not ModeSelection then Exit ;
NbItems:=0 ; for i:=0 to MaxDroite-1 do if LP[i]<>0 then Inc(NbItems) ; Multi:=(NbItems>1) ;
if Multi then ii:=slSolutionMultiple else ii:=slSolutionUnique ;
for i:=0 to MaxDroite-1 do if LP[i]<>0 then
    BEGIN
    if LG[i]>0 then
       BEGIN
       if GridRappro=2 then AjouteSol(TLAC,LG[i],NoSol,ii) else AjouteSol(TLAD,LG[i],NoSol,ii) ;
       END else
       BEGIN
       if GridRappro=2 then AjouteSol(TLAD,-LG[i],NoSol,ii) else AjouteSol(TLAC,-LG[i],NoSol,ii) ;
       END ;
    END ;
for i:=1 to GD.RowCount-1 do if EstSelect(GD,i) then AjouteSol(TLAD,i,NoSol,slPaquet) ;
// en mode pcl on utilise uniquement la grille GD donc les credits sont recuperes de GD
if FBoUneGrille then
 BEGIN
  if GridRappro=2 then
   for i:=1 to GD.RowCount-1 do if EstSelect(GD,i) then AjouteSol(TLAC,i,NoSol,slPaquet)
 END
else BEGIN for i:=1 to GC.RowCount-1 do if EstSelect(GC,i) then AjouteSol(TLAC,i,NoSol,slPaquet) ; END ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 24/09/2002
Modifié le ... : 24/09/2002
Description .. : - 24/09/2002 - on en recuperait pas le bon montant dans la 
Suite ........ : grille ( bug en lettrage combinatoire pour des solutions 
Suite ........ : multiples )
Mots clefs ... : 
*****************************************************************}
procedure TFLettrage.CompleteListe ( Var LM : T_D ; Var LP,LG : T_I ; Var NbD : integer ; NoSol,ii : integer ) ;
Var NbP,i : integer ;
    G   : THGrid ;
    TT  : TLIST ;
    O : tobm;
BEGIN
{OKOK}
//LG* 17/01/2002
NbP:=0 ;
if GridRappro=2 then BEGIN G:=GC ; TT:=TLAC ; END else BEGIN G:=GD ; TT:=TLAD ; END ;
if FBoUneGrille then G:=GD ;
for i:=1 to G.RowCount-2 do if ((Not EstSelect(G,i)) and (Not EstLettre(G,i,True))) then
 if i<=MaxDroite then
    BEGIN
    if P_LAS(TT[i-1]).S[NoSol]=slSolutionMultiple then
       BEGIN
       Inc(NbP) ;
       if NbP<>ii then
        BEGIN // rajout de la solution multiple
         O:=GetO(G,i) ; if O=Nil then Exit ;
         LM[NbD]:=Abs(Arrondi(GetDebitCur(O),CurDec)-Arrondi(GetCreditCur(O),CurDec)) ; // on recupere forcement le montant
         LP[NbD]:=0 ; LG[NbD]:=i ; Inc(NbD) ;
        END ;
       END ;
    if NbD>MaxDroite-1 then Break ;
    END ;
if NbD>MaxDroite-1 then NbD:=MaxDroite-1 ;
END ;

function TFLettrage.Pareil ( Var LG,LT : T_I ; NbD : integer ) : boolean ;
Var i,k,Max,NbG,NbT : integer ;
    Okok : boolean ;
BEGIN
{OKOK}
Pareil:=False ;
NbG:=0 ; NbT:=0 ; Max:=GD.RowCount+GC.RowCount-2 ;
if Max>MaxDroite-1 then Max:=MaxDroite-1 ;
for i:=0 to Max do if LG[i]<>0 then Inc(NbG) ;
for i:=0 to Max do if LT[i]<>0 then Inc(NbT) ;
if NBG<>NbT then Exit ;
for i:=0 to NbD do if ((i<MaxDroite) and (LG[i]<>0)) then
    BEGIN
    Okok:=False ; for k:=0 to Max do if ((LT[k]<>0) and (LT[k]=LG[i])) then BEGIN Okok:=True ; Break ; END ;
    if Not Okok then BEGIN Pareil:=False ; Exit ; END ;
    END ;
Pareil:=True ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 14/12/2001
Modifié le ... :   /  /
Description .. : LG Modif : suppression d'un Warning
Mots clefs ... :
*****************************************************************}
function TFLettrage.NouvelleDonne ( LG : T_I ; NbD : integer ) : boolean ;
Var i     : integer ;
    Idem  : boolean ;
    LT    : T_I ;
BEGIN
{OKOK}
Idem:=False ;
for i:=0 to TTSOLU.Count-1 do
    BEGIN
    LT:=P_I(TTSOLU[i]).I ; if Pareil(LG,LT,NbD) then BEGIN Idem:=True ; Break ; END ;
    END ;
result:=Not Idem ;
END ;

Procedure TFLettrage.AfficheLettAutoEnCours(HNM : THNumedit ; Val : Integer) ;
BEGIN
HNM.Value:=Val ; Application.ProcessMessages ;
END ;


{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 14/12/2001
Modifié le ... : 24/09/2002
Description .. : LG Modif : suppression de warning
Suite ........ : 
Suite ........ : -LG-24/09/2002-correction du comptage du nombre de 
Suite ........ : solution multiple quand on dispose d'une seule grille
Mots clefs ... : 
*****************************************************************}
procedure TFLettrage.ExploreUnivers ( Solde : double ; NoSol : integer ; Unique : boolean ) ;
Var Niveau,i,NbD,NbPlus,OldD : integer ;
    ListeM,OldM : ^T_D ;
    Multi       : boolean ;
    ListeP,ListeG,OldP,OldG : ^T_I ;
    P                       : P_I ;
    Infos                   : REC_AUTO ;
BEGIN
{OKOK}
If HalteAuFeu Then Exit ;
if NoSol>=MaxSolRech then Exit ;
AfficheLettAutoEnCours(HNBP,NoSol) ;
New(ListeM) ; New(ListeP) ;   New(ListeG) ;
New(OldM) ;  New(OldP) ;  New(OldG) ;
ChargeElimine(ListeM^,ListeP^,ListeG^,NbD,NoSol) ;
if ListeM^[0]=0 then Exit ; NbPlus:=0 ;
// on compte le nombre de solution multiple
// quand on a une seule grille on recupere les valeurs dans GD
if FBoUneGrille then
BEGIN
 for i:=1 to GD.RowCount-2 do if ((Not EstSelect(GD,i)) and (Not EstLettre(GD,i,True))) then
 if i<=MaxDroite then
    BEGIN
    if P_LAS(TLAD[i-1]).S[NoSol]=slSolutionMultiple then Inc(NbPlus) ;
    if P_LAS(TLAC[i-1]).S[NoSol]=slSolutionMultiple then Inc(NbPlus) ;
    END ;
END
else
BEGIN
for i:=1 to GD.RowCount-2 do if ((Not EstSelect(GD,i)) and (Not EstLettre(GD,i,True))) then
 if i<=MaxDroite then
    BEGIN
    if P_LAS(TLAD[i-1]).S[NoSol]=slSolutionMultiple then Inc(NbPlus) ;
    END ;
for i:=1 to GC.RowCount-2 do if ((Not EstSelect(GC,i)) and (Not EstLettre(GC,i,True))) then
 if i<=MaxDroite then
    BEGIN
    if P_LAS(TLAC[i-1]).S[NoSol]=slSolutionMultiple then Inc(NbPlus) ;
    END ;
END;
if NbPlus<=0 then Exit ;
OldM^:=ListeM^ ; OldP^:=ListeP^ ; OldG^:=ListeG^ ; OldD:=NbD ; Niveau:=0 ;
for i:=1 to NbPlus do
    BEGIN
    AfficheLettAutoEnCours(HNBN,i) ;
    ListeM^:=OldM^ ; ListeP^:=OldP^ ; ListeG^:=OldG^ ; NbD:=OldD ;
    CompleteListe(ListeM^,ListeP^,ListeG^,NbD,NoSol,i) ;
    Infos.Nival:=Niveau ; Infos.NbD:=NbD ; Infos.Decim:=DEV.Decimale ;
    Infos.Temps:=MaxTempo ; Infos.Unique:=Unique ;
    if Unique then Niveau:=0 else Niveau:=MaxNivProf-1 ;
    if ((NoSol<MaxSolRech) and (MaxNumSol<MaxSolRech) And (Not HalteAuFeu) and
        (NouvelleDonne(ListeG^,NbD)) and (LettrageAuto(Solde,ListeM^,ListeP^,Infos)=1)) then
       BEGIN
       Inc(MaxNumSol) ;
       AfficheLettAutoEnCours(HNBS,MaxNumSol) ;
       ArchiveSolution(ListeP^,ListeG^,MaxNumSol,Multi) ;
       if Multi then ExploreUnivers(Solde,MaxNumSol,Unique) ;
       END ;
    P:=P_I.Create ; P.I:=ListeG^ ; TTSOLU.Add(P) ;
    END ;
Dispose(ListeM) ; Dispose(ListeP) ; Dispose(ListeG) ;
Dispose(OldM) ; Dispose(OldP) ; Dispose(OldG) ;
END ;

procedure TFLettrage.InitLesTLA ;
Var i : integer ;
    X : T_LAS ;
BEGIN
{OKOK}
for i:=0 to TLAD.Count-1 do BEGIN X:=P_LAS(TLAD[i]).S ; FillChar(X,Sizeof(X),slPasSolution) ; P_LAS(TLAD[i]).S:=X ; END ;
for i:=0 to TLAC.Count-1 do BEGIN X:=P_LAS(TLAC[i]).S ; FillChar(X,Sizeof(X),slPasSolution) ; P_LAS(TLAC[i]).S:=X ; END ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 29/11/2001
Modifié le ... : 10/06/2002
Description .. : LG modif :
Suite ........ : -Ajout du parametre d'affichage de la grille des
Suite ........ : credits
Suite ........ : -Suppresson de warning
Suite ........ : 
Suite ........ : - 10/06/2002 - GP - ajout d'une fenetre d'annulation
Suite ........ : 
Mots clefs ... : 
*****************************************************************}
procedure TFLettrage.EnvoieLaSauce ( Solde : double ; Unique : boolean ) ;
Var ListeM : ^T_D ;
    ListeP,ListeG : ^T_I ;
    P             : P_I ;
    NbD,NoSol,ii,Niveau : integer ;
    Multi           : boolean ;
    XX              : T_LETTECHANGE ;
    Infos           : REC_AUTO ;
BEGIN
{OKOK}
VH^.STOPRSP:=FALSE ;
if Action=taConsult then Exit ;
SourisSablier ;
try
NoSol:=0 ; VideListe(TTSOLU) ; InitLesTLA ;
If Multi Then
  BEGIN
  PMessLettAuto.Visible:=TRUE ; HNBS.Value:=0 ; HNBN.Value:=0 ; HNBP.Value:=0 ;
  Application.ProcessMessages ;
  END ;
Repeat
 New(ListeM) ;  New(ListeP) ;  New(ListeG) ;
 ChargeElimine(ListeM^,ListeP^,ListeG^,NbD,NoSol) ;
 if Unique then Niveau:=0 else Niveau:=MaxNivProf-1 ;
 Infos.Nival:=Niveau ; Infos.NbD:=NbD ; Infos.Decim:=DEV.Decimale ;
 Infos.Temps:=MaxTempo ; Infos.Unique:=Unique ;
 if ((NoSol<MaxSolRech) and (NouvelleDonne(ListeG^,NbD)) and
     (LettrageAuto(Solde,ListeM^,ListeP^,Infos)=1)) then
    BEGIN
    P:=P_I.Create ; P.I:=ListeG^ ; TTSOLU.Add(P) ;
    Inc(NoSol) ; ArchiveSolution(ListeP^,ListeG^,NoSol,Multi) ;
    AfficheLettAutoEnCours(HNBS,1) ;
    Dispose(ListeM) ; Dispose(ListeP) ; Dispose(ListeG) ;
    If HalteAuFeu Then break ;
    if Multi then BEGIN MaxNumSol:=NoSol ; ExploreUnivers(Solde,NoSol,Unique) ; NoSol:=MaxNumSol ; END
             else BEGIN if NoSol>MaxNumSol then MaxNumSol:=NoSol ; END ;
    END else
    BEGIN
    Dispose(ListeM) ; Dispose(ListeP) ; Dispose(ListeG) ;
    Break ;
    END ;
Until ((NbD<=0) or (MaxNumSol>=MaxSolRech)) ;
If PMessLettAuto.Visible Then
  BEGIN
  PMessLettAuto.Visible:=FALSE ; Application.ProcessMessages ;
  END ;
if MaxNumSol<=0 then HLettre.Execute(17,'','') else
   BEGIN
   FillChar(XX,Sizeof(XX),#0) ;
   if HT_AUXILIAIRE.Caption<>'' then XX.Titre:=HT_AUXILIAIRE.Caption else XX.Titre:=HG_GENERAL.Caption ;
   XX.RLL:=RL ; XX.DEV:=DEV ; XX.Client:=Client ; XX.NbSol:=MaxNumSol ;
   //LG* 29/11/2001 Ajout du parametre d'affichage de la grille des credits

   ii:=VisuSolutionsLettrage(GD,GC,TLAD,TLAC,XX,FBoUneGrille) ;

   if ii>0 then PreLettreSolution(ii) ;
   END ;
finally
 SourisNormale ;
end;
VH^.STOPRSP:=FALSE ;
END ;

{================================== ENTETE =============================================}
procedure TFLettrage.CParCouvertureClick(Sender: TObject);
begin
ModeCouverture:=CParCouverture.Checked ;
CPArExcept.Enabled:=(Not ModeCouverture) ;
if ModeCouverture then CParExcept.Checked:=False ;
if RL.Appel in [tlSaisieTreso,tlSaisieCour] then BEGIN CParExcept.Enabled:=False ; CParExcept.Checked:=False ; END ;
end;

procedure TFLettrage.CParEnsembleClick(Sender: TObject);
begin
{OKOK}
DelettrageGlobal:=CParEnsemble.Checked ;
end;

procedure TFLettrage.H_REGULDblClick(Sender: TObject);
Var A : TActionFiche ;
begin
{OKOK}
if Not ExJaiLeDroitConcept(TConcept(ccJalModif),False) then A:=taConsult else A:=taModif ;
if E_REGUL.Value<>'' then FicheJournal(Nil,'',E_REGUL.Value,A,0) ;
end;

procedure TFLettrage.H_ECARTDblClick(Sender: TObject);
Var A : TActionFiche ;
begin
{OKOK}
if Not ExJaiLeDroitConcept(TConcept(ccJalModif),False) then A:=taConsult else A:=taModif ;
if E_ECART.Value<>'' then FicheJournal(Nil,'',E_ECART.Value,A,0) ;
end;

Function TFLettrage.ChargeInfosRegul ( Var J : REGJAL ; Value : String ; Controle : boolean ) : boolean ;
BEGIN
Result:=True ;
if Not OkLettrage then Exit ; if Value=J.Journal then Exit ;
RemplirInfosRegul(J,Value) ;
if ((Value<>'') and (J.Facturier='')) then BEGIN HLettre.Execute(28,'','') ; Result:=False ; Exit ; END ;
if Controle then
   BEGIN
   if J.Journal<>'' then if ((J.D[1]='') and (J.D[2]='') and (J.D[3]='')) or ((J.C[1]='') and (J.C[2]='') and (J.C[3]=''))
      then HLettre.Execute(6,'','') ;
   END ;
END ;

{procedure TFLettrage.AvertDevise ; // 14476
BEGIN
if FBoUneGrille then exit;
HLettre.Execute(18,'','') ;
END ;}

procedure TFLettrage.E_REGULChange(Sender: TObject);
begin
if Not ChargeInfosRegul(JREGUL,E_REGUL.Value,True) then E_REGUL.Value:='' ;
// S'assure lors d'un lettrage en devise que le taux de change du jour est renseigné, sinon affiche un message
// if ((E_REGUL.Value<>'') and (DEV.DateTaux<>LaDateGenere) and (DEV.Code<>V_PGI.DevisePivot)) then AvertDevise ; // 14476
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 26/04/2002
Modifié le ... : 18/06/2002
Description .. : - on rend actif le bouton de lettrage auto apres un choix
Suite ........ : valide de journal ( il n'est pas actif par defaut pour un
Suite ........ : lettrage en devise )
Suite ........ : 
Suite ........ : - FB 10144 - le popup de methode le lettrage n'est aps 
Suite ........ : accessible en delettrage
Mots clefs ... : 
*****************************************************************}
procedure TFLettrage.E_ECARTChange(Sender: TObject) ;
begin  //LG*
if Not ChargeInfosRegul(JECART,E_ECART.Value,False) then E_ECART.Value:='' ;
BCalcul.Enabled:= FBoUneGrille and (E_ECART.Value<> '') ; Zoom1.Enabled:= BCalcul.Enabled and OkLettrage ;
end;

procedure TFLettrage.VireCouverture ;
Var O : TOBM ;
    i : integer ;
BEGIN
{OKOK}
for i:=1 to GD.RowCount-1 do
    BEGIN
    O:=GetO(GD,i) ; if O=Nil then Break ;
    if JePrends(GD,i) then
       BEGIN
       O.PutMvt('E_COUVERTURE',0) ; O.PutMvt('E_COUVERTUREDEV',0) ;
       END ;
    END ;
for i:=1 to GC.RowCount-1 do
    BEGIN
    O:=GetO(GC,i) ; if O=Nil then Break ;
    if JePrends(GC,i) then
       BEGIN
       O.PutMvt('E_COUVERTURE',0) ; O.PutMvt('E_COUVERTUREDEV',0) ; 
       END ;
    END ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 26/12/2001
Modifié le ... : 03/11/2004
Description .. : utilisation des fonctions du lettrage automatique
Suite ........ : - LG - 03/11/2004 - FB 14765 - correction pour le lettrage 
Suite ........ : en devise
Mots clefs ... : 
*****************************************************************}
procedure TFLettrage.LettrageEpuise;
Var
LD,LC : TList;
BEGIN
{OKOK}
if Action=taConsult then Exit ;
{si Delettrage --> virer les couvertures}
if Not OkLettrage then VireCouverture ;
//LG*
LD:=TList.Create;LC:=TList.Create;
try
 CAssignListRappro(GD,LD,LC,'M',RL.LettrageDevise);
 VideListe(FLDAnnul) ; VideListe(FLCAnnul) ; CAssignListRappro(GD,FLDAnnul,FLCAnnul,'M',RL.LettrageDevise);
 if not FBoUneGrille then CAssignListRappro(GC,LD,LC,'M',RL.LettrageDevise);
 RefEpuise(LD,LC);
 //recuperation des valeurs des TList
 if FBoUneGrille then
 begin
   CTLettVersTHGrid(GD,LD,RL,true) ;
   CTLettVersTHGrid(GD,LC,RL,false) ;
  end
 else begin CTLettVersTHGrid(GD,LD,RL,true); CTLettVersTHGrid(GC,LC,RL,false); end;
finally
 VideListe(LD) ; VideListe(LC) ;
 LD.Free;LC.Free;
end; // try
END ;


{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 07/04/2006
Modifié le ... :   /  /    
Description .. : - LG - 07/03/2006 - FB 13175 - paramsoc qui conditionne 
Suite ........ : l'ouverture
Mots clefs ... :
*****************************************************************}
procedure TFLettrage.ProposeValide ;
Var i : integer ;
BEGIN
{OKOK}
if Action=taConsult then Exit ;
if Not ModeSelection then Exit ; if ((NbSelD<=0) and (NbSelC<=0)) then Exit ;
if FBoUneGrille and not GetParamSocSecur('SO_LETAVER',true) then
 ClickValide
else
 begin
  i:=HLettre.Execute(1,'','') ; if i=mrYes then ClickValide ;
 end ;
//END ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 27/12/2001
Modifié le ... :   /  /
Description .. : utilisation d'une fonction générique
Mots clefs ... :
*****************************************************************}
procedure TFLettrage.TraiteLeEtat ( O : TOBM ) ;
BEGIN
//LG* 19/12/2001
O.TraiteLeEtat(PasEcart,AvecRegul,EcartChange,AvecConvert,false,IndiceRegul,'M');
END ;

procedure TFLettrage.FinirLettrage ( G : THgrid ) ;
Var i : integer ;
    O : TOBM ;
    St : string;
BEGIN
if Action=taConsult then Exit ;
if Not LettrageTotCur then CodeL:=LOWERCASE(CodeL) else CodeL:=Trim(uppercase(CodeL)+CSuppLett(PasEcart,AvecRegul,EcartChange,AvecConvert)) ;
H_CodeL.Caption:=CodeL ;
for i:=1 to G.RowCount-1 do if (JePrends(G,i)) then
    BEGIN
    O:=GetO(G,i) ; if O=Nil then Break ;
    //LG* 26/11/2001 on test si on recupere le code lettrage du TL_RAPPRO
    St:=O.GetMvt('E_ETAT') ;
    //if St='xxxxxxxxx' then St='00000000' ;
    TraiteLeEtat(O) ;
    O.PutMvt('E_LETTRAGE',CodeL) ;
    O.PutMvt('E_REFLETTRAGE',E_REFLETTRAGE.Text) ;
    if RL.Appel=tlSaisieTreso then O.PutMvt('E_TRESOLETTRE','X') else O.PutMvt('E_TRESOLETTRE','-') ;
    O.PutMvt('E_DATEPAQUETMAX',DatePaquetMax) ; O.PutMvt('E_DATEPAQUETMIN',DatePaquetMin) ;
    if LettrageTotCur then O.PutMvt('E_ETATLETTRAGE','TL') else O.PutMvt('E_ETATLETTRAGE','PL') ;
    if RL.LettrageDevise then
       BEGIN
       O.PutMvt('E_LETTRAGEDEV','X') ;
       END else
       BEGIN
        O.PutMvt('E_LETTRAGEDEV','-') ;
       END ;
    {Ruses pour gestion interne : les ecarts de regul, conversion sont mis dans des champs inutilise de la TOBM}
    O.PutMvt('OLDDEBIT',EcartRegulDevise) ;
    O.PutMvt('OLDCREDIT',EcartChangeEuro) ;
    O.PutMvt('RATIO',EcartChangeFranc) ;
    O.PutMvt('CONVERTEURO',ConvertEuro) ;
    O.PutMvt('CONVERTFRANC',ConvertFranc) ;
    END ;
END ;


procedure TFLettrage.FinirDeLettrage ( G : THgrid ) ;
Var i,Col : integer ;
    O     : TOBM ;
BEGIN
{OKOK}
if Action=taConsult then Exit ;
if OkLettrage then Exit ;
if G=GD then Col:=ColLetD else Col:=ColLetC ;
for i:=1 to G.RowCount-1 do if EstSelect(G,i) then
    BEGIN
    O:=GetO(G,i) ; if O=Nil then Break ;
    G.Cells[Col,i]:='' ; G.Cells[G.ColCount-1,i]:=' ' ; O.PutMvt('E_LETTRAGE','') ;
    O.PutMvt('E_DATEPAQUETMAX',IDate1900) ; O.PutMvt('E_DATEPAQUETMIN',IDate1900) ;
    O.PutMvt('E_ETATLETTRAGE','AL') ; O.PutMvt('E_LETTRAGEDEV','-') ;
    O.PutMvt('E_COUVERTURE',0)      ; O.PutMvt('E_COUVERTUREDEV',0) ;
    O.PutMvt('E_REFLETTRAGE','') ;
    TraiteLeEtat(O) ;
    { GP le 18/11/97 : Pour appel délettrage à partir de la saisie }
    if (RL.Appel=tlSaisieTreso) then O.PutMvt('E_TRESOLETTRE','-') ;
    END ;
END ;

function TFLettrage.EquilibreCouverture : boolean ;
BEGIN
{OKOK}
Result:=Arrondi(LTE_CouvDebCur-LTE_CouvCredCur,CurDec)=0 ;
if Not Result then HLettre.Execute(11,'','') ;
END ;

procedure TFLettrage.UpdateDevalide ;
Var SQL : String ;
BEGIN
{OKOK}
if RL.Auxiliaire<>'' then
   BEGIN
   SQL:='UPDATE TIERS SET T_DERNLETTRAGE="'+SauveCode+'" Where T_AUXILIAIRE="'+RL.Auxiliaire+'" AND T_DERNLETTRAGE="'+DernLettrage+'"' ;
   END else
   BEGIN
   SQL:='UPDATE GENERAUX SET G_DERNLETTRAGE="'+SauveCode+'" Where G_GENERAL="'+RL.General+'" AND G_DERNLETTRAGE="'+DernLettrage+'"' ;
   END ;
ExecuteSQL(SQL) ;
END ;


procedure TFLettrage.DevalideCode ;
BEGIN
{OKOK}
if DernLettrage=SauveCode then Exit ;
if ErreurReseau then Exit ;
if Transactions(UpdateDevalide,3)<>oeOk then MessageAlerte(HDivL.Mess[13]) ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : LG* Laurent GENDREAU
Créé le ...... : 13/12/2001
Modifié le ... : 08/04/2002
Description .. : Remplacement de la focntion par GetSetCodeLettre de
Suite ........ : LettUtil
Mots clefs ... : 
*****************************************************************}
procedure TFLettrage.GetCodeSuivant ;
var
 X : TL_Rappro;
BEGIN
 X:=TL_Rappro.Create;
 try
  X.General:=RL.General ; X.Auxiliaire:=RL.Auxiliaire ;
  DernLettrage:=GetSetCodeLettre(X.General,X.Auxiliaire,DernLettrage);
 finally
  X.Free;
 end;
END ;

procedure TFLettrage.PrepareLettrageSauf ;
Var i : integer ;
BEGIN
for i:=1 to GD.RowCount-2 do InverseSelection(GD,i) ;
for i:=1 to GC.RowCount-2 do InverseSelection(GC,i) ;
GD.Invalidate ; GC.Invalidate ; CalculDebitCredit(False) ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 10/06/2002
Modifié le ... : 04/10/2007
Description .. : - 10/06/2002 - on ne se replace plus au debut de la grille
Suite ........ : - FB 10140 - 17/06/2002 - pour un lettrage en devise, le 
Suite ........ : rafraissiment d'ecran n'etait pas correcte
Suite ........ : - LG - 07/06/2006 - FB 18314 - correction de la gestion de 
Suite ........ : l'affichage des totalement lettres
Suite ........ : - LG - FB 21517 - 04/10/2007 - on re selectionnait la ligne 
Suite ........ : de saisie apres une mise a jour de la grille
Mots clefs ... :  
*****************************************************************}
procedure TFLettrage.RemettreGrilleAJour(G : THGrid ; Oter : boolean);
var
  i,LastRow, LastRow2 : integer;
lStDebit, lStCredit : string ;
lRdSolde : double ;
lStCode : string ;
O : TOBM ;
begin
if FTOBEcr.Detail = nil then exit ;
if FBoUneGrille then
 BEGIN
  if FInOrdreTri=1 then BEGIN FTOBEcr.Detail.Sort('E_LETTRAGE;E_AUXILIAIRE;E_GENERAL;E_DATECOMPTABLE;E_NUMEROPIECE;E_NUMLIGNE;E_NUMECHE') END
  else BEGIN FTOBEcr.Detail.Sort('E_AUXILIAIRE;E_GENERAL;E_DATECOMPTABLE;E_NUMEROPIECE;E_NUMLIGNE;E_NUMECHE') END ;
  AnalyseTitreMontant(lStDebit,RL.LettrageDevise) ; AnalyseTitreMontant(lStCredit,RL.LettrageDevise) ;
  lStDebit:='E_DEBIT' ; lStCredit:='E_CREDIT' ; lRdSolde:=0 ;
  G.VidePile(false);
  for i:=0 to FTOBEcr.Detail.Count-1 do
   BEGIN
    O:=TOBM(FTOBEcr.Detail[i]) ;
    if Oter then
     begin
      lStCode:=trim(O.GetValue('E_LETTRAGE')) ;
      if ((lStCode<>'') and (lStCode=uppercase(lStCode))) then
       continue ;
     end;
    O.PutValue('SELECTIONNER','-') ; // qd on remet la grille a jour -> plus aucune ecriture n'est selectionner
    G.Row:=G.RowCount-1 ; G.Objects[GD.ColCount-1,GD.Row]:=FTOBEcr.Detail[i] ;
    lRdSolde:=lRdSolde+FTOBEcr.Detail[i].GetValue(lStDebit)-FTOBEcr.Detail[i].GetValue(lStCredit) ; FTOBEcr.Detail[i].PutValue('SOLDEPRO',lRdSolde) ;
    CTOBVersTHGrid(G,GetO(G,G.Row),RL.LettrageDevise) ;
    if ((RL.Appel<>tlMenu) and (EcheSaisie(O))) then // Fb 21571 - 
     begin
        GSais:=G ; LigSais:=G.RowCount-1 ;
     end ;
    G.RowCount:=G.RowCount+1 ;
   END; // for
 END
else BEGIN
  LastRow := G.Row;
  LastRow2 := G.TopRow;

  for i:=1 to G.RowCount-1 do
   begin
    G.Row:=i ;
    CTOBVersTHGrid(G,GetO(G,i),RL.LettrageDevise);
   end ;
  G.TopRow := LastRow2;
  G.Row := LastRow;
END; // if
 // changement de bouton enlever ou remettre les ecritures
 EnleveOuRajoute(not Oter) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 20/11/2001
Modifié le ... : 07/06/2006
Description .. : LG Modif :
Suite ........ : - ouverture des fenetres de saisie d'ecritures
Suite ........ : 
Suite ........ : - 10/06/2002 - Avec une seule grille on n'affiche plus la
Suite ........ : message d'avertissement sur le lettrage partiel ici, mais dans
Suite ........ : la fct Validation
Suite ........ : - 07/04/2006 - FB 16322 - le curseur se palcait une ligne 
Suite ........ : trop bas qd on affichait pas les ecritures totalement lettrés
Suite ........ : - 07/06/2006 - FB 13402 - on passe la piece en modif qd 
Suite ........ : lance le delettrage ( une question es t pose si on 
Suite ........ : abandonne )
Mots clefs ... : 
*****************************************************************}
function TFLettrage.LettrageMemoire : boolean;
Var Okok : boolean ;
    i,index : integer;
    O : TOBM;
    lTOBSauve : TOBM;
    CurrentRow : integer ;
    DMin,DMax,DEche : tDateTime ;
    ModePaie : string ;
BEGIN
result:=true ;
if Action=taConsult then Exit ;
if Not ModeSelection then Exit ;
if ((ModeCouverture) and (Not EquilibreCouverture)) then Exit ;
result:=false ; GD.SynEnabled:=false ; GC.SynEnabled:=false ; CurrentRow:=GD.Row ;
lTOBSauve:=CGetTOBSelectFromGrid(GD);
try
if ((OkLettrage) and (Not ModeCouverture) and (CParExcept.Checked)) then
   BEGIN
   if LeSolde.Debit then Okok:=Arrondi(LTE_TotDebP-LTE_TotCredP-LeSolde.Value,V_PGI.OkDecV)=0
                    else Okok:=Arrondi(LTE_TotCredP-LTE_TotDebP-LeSolde.Value,V_PGI.OkDecV)=0 ;
   if Not Okok then
      BEGIN
      HLettre.Execute(33,'','') ; Exit ;
      END else
      BEGIN
      PrepareLettrageSauf ;
      END ;
   END ;
{Récup code lettre, date paquet ...}
LettrageTotP:=Arrondi(LTE_TotDebP-LTE_TotCredP,V_PGI.OkDecV)=0 ;
LettrageTotCur:=Arrondi(LTE_TotDebCur-LTE_TotCredCur,CurDec)=0 ;
if ((OkLettrage) and (ModeCouverture)) then
   BEGIN
   LettrageTotP:=LettrageTotP and (Arrondi(LTE_CouvDebP-LTE_TotDebP,V_PGI.OkDecV)=0) and (Arrondi(LTE_CouvCredP-LTE_TotCredP,V_PGI.OkDecV)=0) ;
  // LettrageTotE:=LettrageTotE and (Arrondi(LTE_CouvDebE-LTE_TotDebE,V_PGI.OkDecE)=0) and (Arrondi(LTE_CouvCredE-LTE_TotCredE,V_PGI.OkDecE)=0) ;
   LettrageTotCur:=LettrageTotCur and (Arrondi(LTE_CouvDebCur-LTE_TotDebCur,CurDec)=0) and (Arrondi(LTE_CouvCredCur-LTE_TotCredCur,CurDec)=0) ;
   END ;
if ((OkLettrage) and (Not LettrageTotCur)) then
   BEGIN
   if NbReleve=1 then
      BEGIN
      {Alerte sur lettrage partiel d'un relevé}
      if HLettre.Execute(19,'','')<>mrYes then Exit ;
      END else if not FBoUneGrille and VH^.AlertePartiel then //LG* 10/06/2002
      BEGIN
      if HLettre.Execute(31,'','')<>mrYes then Exit ;
      END ;
   END ;
CodeDateLettre(DMin,DMax,DEche,ModePaie) ;
if ((CodeL='') and (OkLettrage)) then
   BEGIN
   if Transactions(GetCodeSuivant,3)=oeOk then
   CodeL:=DernLettrage
   else
      BEGIN
      MessageAlerte(HDivL.Mess[12]) ; ErreurReseau:=True ;
      Exit ;
      END ;
   END ;
{Algo complexe}
if Not ModeCouverture then
   BEGIN
   if not LettrageTotCur then LettrageEpuise ;
   END ;
if OkLettrage then
   BEGIN
   GereRegul ; GereEcartDeChange ; //GereConvert(DMin,DMax) ;
   if not FBoUneGrille and ((AvecRegul) or (AvecConvert) or (EcartChange)) then
      if ((LettrageTotCur) and (VH^.AlerteRegul)) then
      BEGIN
      if HLettre.Execute(32,'','')<>mrYes then Exit ;
      END ;
   END ;
//LG*2 15/ 11/2001
if (not FBoUneGrille) or (LettrageTotCur and not AvecRegul) then
BEGIN
 FinirLettrage(GD) ; FinirDeLettrage(GD) ;
 FinirLettrage(GC) ; FinirDeLettrage(GC) ;
END
else
BEGIN
if FBoUneGrille and (FIFL.ChoixValid<>'AL5') then
 BEGIN
  index:=0 ;
  if ( (not LettrageTotCur) or AvecRegul ) and not Validation(DMin,DMax,DEche,ModePaie) then
  BEGIN
   for i:=1 to GD.RowCount - 1 do
    if EstSelect(GD,i) then
     BEGIN
      O:=GetO(GD,i) ; if O=nil then continue ;
      O.Dupliquer(lTOBSauve.Detail[index],false,true) ; Inc(index) ;
     END;
    //action:=taModif;
    exit;
  END;
 END; // if
END;
//Remplissage final de les grids et affectation du code lettre
// bloque le rafraichissement de la grille
if FBoUneGrille then GD.BeginUpdate;
RemettreGrilleAJour(GD,FBoOter) ; if not FBoUneGrille then RemettreGrilleAJour(GC,FBoOter) ;
GD.Invalidate ; GC.Invalidate ;
Desectionner(GD) ; Desectionner(GC) ;
{Affichages et re-init}
ModeSelection:=False ; NbSelD:=0 ; NbSelC:=0 ; IndiceRegul:=0 ;
CodeL:='' ; PieceModifiee:=True ;
LettrageTotCur:=False ; LettrageTotP:=False ; //LettrageTotE:=False ;
EcartChangeFranc:=0 ; EcartChangeEuro:=0 ; EcartRegulDevise:=0 ; ConvertFranc:=0 ; ConvertEuro:=0 ;
CParCouverture.Enabled:=True ; if ((LienS1S3) or (EstSerie(S3))) then CParCouverture.Enabled:=False ;
CParEnsemble.Enabled:=True ; CParExcept.Enabled:=True ;
if RL.Appel in [tlSaisieTreso,tlSaisieCour] then BEGIN CParExcept.Enabled:=False ; CParExcept.Checked:=False ; END ;
POPAnnulTotal.Enabled:=true ; DELettTotal.Enabled:=true ;// on rend actif ce menu une fois que l'on a lettré un elements
GereLesValide ;
if FBoUneGrille then
 begin
  if FboOter then
   begin
    if ( CurrentRow > 1 ) and ( CurrentRow < GD.RowCount) then
     GD.Row:=CurrentRow - 1;
   end
    else
     if (CurrentRow<GD.RowCount) then GD.Row:=CurrentRow ;
 end ;
result:=true ;
finally
 if FBoUneGrille then GD.EndUpdate; GD.SynEnabled:=true ; GC.SynEnabled:=true ; lTOBSauve.Free ;
 if FBoUneGrille then GD.Setfocus ; GD.Refresh ;
end;
END ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 10/12/2001
Modifié le ... : 13/09/2005
Description .. : AGLLanceFichie retourne l'item du mode de calcul
Suite ........ : selectionné
Suite ........ :   AL1 : solde automatique
Suite ........ :   AL2 : lettrage partiel
Suite ........ :   Al3 : ecriture simplifie
Suite ........ :   AL4 : annulation
Suite ........ :
Suite ........ : - 10/06/2002 - gestion du paramtre d'avertissement sur le
Suite ........ : lettrage partiel
Suite ........ : - 19/ 01/2004 - FB 13106 - en lettrage en saisie, on
Suite ........ : supprimer les optopns de lettrage automatiques des
Suite ........ : parametres societes
Suite ........ : - 13/09/2005 - FB 13346 - qd on fait echap sur la fenetre
Suite ........ : de validation, on ferme la fenetre
Mots clefs ... :
*****************************************************************}
function TFLettrage.Validation(DMin,DMax,DEche : tDateTime ; ModePaie : string ) : boolean;
var
 AGLResult : string;
begin
//LG*2
 result := true;

 // en lettrage en saisie on supprime les options de validation automatiques des parametres generaux
 if RL.LettrageEnSaisie then FIFL.ChoixValid := 'AL4' ;

 //Lettrage partiel
 if ( FIFL.ChoixValid='AL2') then
 BEGIN
  if VH^.AlertePartiel then
   BEGIN
    if HLettre.Execute(31,'','')<>mrYes then BEGIN result:=false ; Exit ; END ;
   END ; // if
  AvecRegul:=false ; AvecConvert:=false ; LettrageTotCur:=false ;
  FinirLettrage(GD) ; FinirDeLettrage(GD) ; exit ;
 END ;

 //Ecriture de regul
 if AvecRegul and ( FIFL.ChoixValid='AL1') and ( not FBoAucunJournal ) then
 BEGIN result := SoldeAutomatique('AL1',DMin,DMax,DEche,ModePaie) ; exit ; END;

 //Ecriture de regul
 if ( FIFL.ChoixValid='AL3') then BEGIN result := SoldeAutomatique('AL3',DMin,DMax,DEche,MOdePaie) ;  exit ; END;

 // on envoie '1' a la fenetre pour interdire la saisie de regul
 if RL.LettrageEnSaisie then // on appelle le lettrage depuis la saisie bordereau
  AGLResult := '2'           // on n'a pas acces au ecriture simplifies
   else
    if AvecRegul and ( not FBoAucunJournal ) then
     AGLResult := '1'
      else
        AGLresult := '';
 AGLResult := AGLLanceFiche('CP','CPARAMLET','','',AGLResult); // ouverture de la fenêtre du choix de traitement
 if (AGLResult='AL4') or (AGLResult='') then
 begin
   result    := false ;
   FBoFermer := true ;

   // GCO - 30/05/2006 - FQ 17330
   if FBoFromTOB then
     Close;
 end
 else
  BEGIN
    // l'utilisateur n'a pas choisit de générer une ecriture de regul -> on repasse les flag a false
   //if AvecRegul and (AGLResult<>'AL1') then BEGIN AvecRegul:=false ; AvecConvert:=false ; LettrageTotCur:=false ; END;
   if (AGLResult='AL1') or (AGLResult='AL3') then
    result := SoldeAutomatique(AGLResult,DMin,DMax,DEche,ModePaie) // passage d'une ecriture de regul ou simplifie
     else
      if (AGLResult='AL2') then // lettrage partiel
       begin
        if VH^.AlertePartiel then
         BEGIN
          if HLettre.Execute(31,'','')<>mrYes then BEGIN result:=false ; Exit ; END ;
         END ; // if
        AvecRegul:=false ; AvecConvert:=false ; LettrageTotCur:=false ;
        FinirLettrage(GD) ; FinirDeLettrage(GD) ; result := true ;
       end;
  END;

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 21/03/2002
Modifié le ... : 02/07/2003
Description .. : LG* fonction d'ouverture de la fenetre de choix des ecritures
Suite ........ : de regularisation
Suite ........ : 
Suite ........ : - 25/04/2002 - recuperation du dernier general selectionne
Suite ........ : 
Suite ........ : -07/06/2002 - on test la presence du numero de piece
Suite ........ : avant de la recupere
Suite ........ : - bug 10134 - affectation de la reference lettrage avnat
Suite ........ : l'envoi vers la fenetre de saisie des ecritures reguk/simplifie
Suite ........ : 
Suite ........ : - 10/06/2002 - affectation du champs
Suite ........ : E_PAQUETREVISION
Suite ........ : 
Suite ........ : - 17/09/2002 - suppression du record FIFL
Suite ........ : -01/04/2003 - FB 12079 - on reaffecte le code lettrage 
Suite ........ : apres son recalcul
Suite ........ : -02/05/2003 - FB 12234 - on recupere la date paquet min 
Suite ........ : et max en provenance des ecritures simplifies
Suite ........ : -02/07/2003 - on enregistrait le code lettre avec le signe
Suite ........ : du type de lettrage ( # @ etc etc )
Mots clefs ... :
*****************************************************************}
function TFLettrage.SoldeAutomatique( ChoixValid : string ; DMin,DMax,DEche : tDateTime ; ModePaie : string ) : boolean;
var
 lStRetour,S,l : string ;
 i : integer ;
 lTOBEcr, lTOBResult : TOB;
begin
//LG*2
result:= false ; FStLastError:='' ; lTOBResult:=TOB.Create('E_ECRITURE',nil,-1);
try
 lTOBEcr:=CGetTOBSelectFromGrid(GD) ; // recuperation des elements selectionné de la grille
 if (lTOBEcr.Detail=nil) or (lTOBEcr.Detail.Count=0) then exit ;
 try
  if RL.LettrageDevise then lTOBEcr.Detail[0].PutValue('E_LETTRAGEDEV','X') else lTOBEcr.Detail[0].PutValue('E_LETTRAGEDEV','-') ;
  lTOBEcr.Detail[0].PutValue('E_REFLETTRAGE',rl.RefLettrage) ; lTOBEcr.Detail[0].AddChampSupValeur('ECARTREGUL',EcartRegulDevise,true) ;
  S:=CPLanceFiche_ECRLET(lTOBEcr,lTOBResult,ChoixValid+';'+intToStr(FInLastNumero)+';'+DateTimeToStr(FStLastDateComptable)+';'+FStLastJournal+';'+FStLastLibelle+';'+FStLastGeneral, FInfoEcr) ;
  lStRetour:=ReadTokenST(S) ; result:=lStRetour='1' ; if not result then exit ; l:=ReadTokenST(S);
  // recuperation des valeurs choisit
  if (trim(l)<>'') then FInLastNumero        := StrToInt(l);
  if (S<>'')       then FStLastDateComptable := StrToDateTime(ReadTokenST(S)) else FStLastDateComptable:=iDate1900;
  if (S<>'')       then FStLastJournal       := ReadTokenST(S);
  if (S<>'')       then FStLastLibelle       := ReadTokenST(S);
  if (S<>'')       then FStLastGeneral       := ReadTokenST(S);
  if (S<>'')       then DatePaquetMax        := StrToDateTime(ReadTokenST(S));
  if (S<>'')       then DatePaquetMin        := StrToDateTime(ReadTokenST(S));
  // recuperation de la datepaquet max et min
  AvecRegul:=false ; AvecConvert:=false ; LettrageTotCur:=false ;
  for i:=0 to ( lTOBResult.Detail.Count - 1 ) do
   begin
    if lTOBResult.Detail[i].GetValue('SELECTIONNER')='X' then
     begin
      lTOBResult.Detail[i].PutValue('E_LETTRAGE',Copy(CodeL,1,4)) ; lTOBResult.Detail[i].PutValue('E_REFLETTRAGE',E_REFLETTRAGE.Text) ;
      if rl.Appel=tlSaisieTreso then lTOBResult.Detail[i].PutValue('E_TRESOLETTRE','X') else lTOBResult.Detail[i].PutValue('E_TRESOLETTRE','-') ;
      lTOBResult.Detail[i].PutValue('E_ETATLETTRAGE','TL') ;
      if RL.LettrageDevise then
         BEGIN
         lTOBResult.Detail[i].PutValue('E_LETTRAGEDEV','X') ;
         END else 
         BEGIN
         lTOBResult.Detail[i].PutValue('E_LETTRAGEDEV','-') ;
         END ;
      TraiteLeEtat(TOBM(lTOBResult.Detail[i])) ; lTOBResult.Detail[i].PutValue('E_PAQUETREVISION',1);
     end; // if
   end; // for
  AjouteEcrGrille(lTOBResult) ;
  EcartChangeFranc:=0 ; EcartChangeEuro:=0 ; EcartRegulDevise:=0 ; ConvertFranc:=0 ; ConvertEuro:=0 ;
  CalculDebitCredit(false) ;
  LettrageTotP:=Arrondi(LTE_TotDebP-LTE_TotCredP,V_PGI.OkDecV)=0 ;
  LettrageTotCur:=Arrondi(LTE_TotDebCur-LTE_TotCredCur,CurDec)=0 ;
  GereRegul ; CalculEcartRegulDevise ; GereEcartDeChange ; //GereConvert(DMin,DMax) ;
  FinirLettrage(GD) ; FinirDeLettrage(GD) ;
  // on reaffecte le code lettrage apres le calcul total ou partiel
  for i:=0 to ( lTOBResult.Detail.Count - 1 ) do
  begin
   if lTOBResult.Detail[i].GetValue('SELECTIONNER')='X' then
    begin
     lTOBResult.Detail[i].PutValue('E_DATEECHEANCE',DEche) ;
     lTOBResult.Detail[i].PutValue('E_MODEPAIE',ModePaie) ;
     lTOBResult.Detail[i].PutValue('E_LETTRAGE',Copy(CodeL,1,4)) ;
    end ;
    // YMO 31/07/2006 FQ17589 Renseignement du régime de tva (nécessaire pour l'état de contrôle)
    lTOBResult.Detail[i].PutValue('E_TVA',lTOBEcr.Detail[0].GetValue('E_TVA')) ; 
  end;  
  FEcrRegul.RecupereLigne(lTOBResult) ; if GD.CanFocus then GD.SetFocus ;
 finally
  FreeAndNil(lTOBResult) ; FreeAndNil(lTOBEcr) ;
 end; //try
except
 on E:Exception do
  BEGIN Showmessage('Erreur lors de la génération des écritures de régularisation'+cDeuxLigne+E.Message) ; FStLastError:=E.Message ; result:=false END ;
end; // try
end;



{================================== REGULS, ECARTS ==========================================}
procedure TFLettrage.CalculEcartRegulDevise ;
BEGIN
{OKOK}
if Not ModeCouverture
   then EcartRegulDevise:=Arrondi(LTE_TotDebCur-LTE_TotCredCur,CurDec)
   else EcartRegulDevise:=Arrondi((LTE_TotDebCur-LTE_CouvDebCur)-(LTE_TotCredCur-LTE_CouvCredCur),CurDec) ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 19/12/2001
Modifié le ... : 19/12/2001
Description .. : -en mode PCL on n'est pas obliger de choisir le journal de
Suite ........ : regul au depart
Suite ........ : -utilisation de parametres generaux pour les ecarts de regul
Mots clefs ... :
*****************************************************************}
procedure TFLettrage.GereRegul ;
Var Cpte  : String ;
    Debit : boolean ;
    EqRegFranc,EqRegEuro,LimiteF,LimiteE,SeuilD,SeuilC : double ;
    DecF : integer ;
BEGIN
if Not OkLettrage then Exit ;
//LG* 30/11/2001 en mode PCL on n'est pas obliger de choisir le journal de regul au depart
if FBoUneGrille and (FIFL.ChoixValid='AL5') and (E_REGUL.Value='') then Exit ;
if not FBoUneGrille and (E_REGUL.Value='') then Exit ;
Debit:=False ;
if Not ModeCouverture then
   BEGIN
   if LettrageTotCur then Exit ;
   END else
   BEGIN
   if ((Arrondi(LTE_TotDebCur-LTE_CouvDebCur,CurDec)<>0) and (Arrondi(LTE_TotCredCur-LTE_CouvCredCur,CurDec)<>0)) then Exit ;
   END ;
CalculEcartRegulDevise ;
if EcartRegulDevise=0 then BEGIN LettrageTotCur:=True ; Exit ; END ;
//LG* 19/12/2001 modif de la gestion des ecarts sur les devises -> utilisation de parametres geenraux pour les ecarts de regul
if FBoUneGrille then
 BEGIN
  if (DEV.Code<>V_PGI.DevisePivot) then BEGIN SeuilD:=DEV.MaxDebit ; SeuilC:=DEV.MaxCredit ; END
  else BEGIN SeuilD:=FIFL.EcartDebit ; SeuilC:=FIFL.EcartCredit ; END ;
  DEV.MaxDebit:=SeuilD ; DEV.MaxCredit:=SeuilC ;
 END
  else
   BEGIN
    SeuilD:=DEV.MaxDebit ; SeuilC:=DEV.MaxCredit ;
   END ;
if EcartRegulDevise>0 then BEGIN if EcartRegulDevise>SeuilD then Exit else Debit:=True ; END ;
if EcartRegulDevise<0 then BEGIN if Abs(EcartRegulDevise)>SeuilC then Exit else Debit:=False ; END ;
if Not VH^.TenueEuro then
   BEGIN
   EqRegFranc:=DeviseToPivot(EcartRegulDevise,DEV.Taux,DEV.Quotite) ;
   EqRegEuro:=DeviseToEuro(EcartRegulDevise,DEV.Taux,DEV.Quotite) ;
   END else
   BEGIN
   EqRegEuro:=DeviseToPivot(EcartRegulDevise,DEV.Taux,DEV.Quotite) ;
   EqRegFranc:=DeviseToEuro(EcartRegulDevise,DEV.Taux,DEV.Quotite) ;  // LG* 26/04/2002 DeviseToFranc -> DeviseToEuro
   END ;
if VH^.TenueEuro then
   BEGIN
   LimiteF:=Resolution(V_PGI.OkDecE) ; LimiteE:=Resolution(V_PGI.OkDecV) ;
   DecF:=V_PGI.OkDecE ; 
   END else
   BEGIN
   LimiteF:=Resolution(V_PGI.OkDecV) ; LimiteE:=Resolution(V_PGI.OkDecE) ;
   DecF:=V_PGI.OkDecV ; 
   END ;
// if DEV.Code<>V_PGI.DevisePivot then BEGIN LimiteF:=2*LimiteF ; LimiteE:=2*LimiteE ; END ;
if EcartRegulDevise<>0 then
   BEGIN
   if VH^.TenueEuro then
      BEGIN
      if Abs(EqRegEuro)<LimiteE then BEGIN HLettre.Execute(25,'','') ; Exit ; END ;
      END else
      BEGIN
      if Abs(EqRegFranc)<LimiteF then BEGIN HLettre.Execute(25,'','') ; Exit ; END ;
      END ;
   END ;
//LG* 30/11/2001   modif pour eAGL

if not FBoUneGrille then
 Cpte:=ChoixCpteRegul(JREGUL,Debit,Abs(EcartRegulDevise),DEV.Decimale,DEV.Symbole,IndiceRegul) ;

if not FBoUneGrille and (Cpte='') then BEGIN IndiceRegul:=0 ; Exit ; END ;
LettrageTotCur:=True ; AvecRegul:=True ;
if not FBoUneGrille then
BEGIN
 if DatePaquetMax<LaDateGenere then DatePaquetMax:=LaDateGenere ;
 if DatePaquetMin>LaDateGenere then DatePaquetMin:=LaDateGenere ;
END;
EcartChangeFranc:=EqRegFranc ; EcartChangeEuro:=EqRegEuro ;
if Not ModeCouverture then
   BEGIN
   if Arrondi(LTE_TotDebP-LTE_TotCredP-EcartChangeFranc,DecF)=0 then LettrageTotP:=True ;
 //  if Arrondi(LTE_TotDebE-LTE_TotCredE-EcartChangeEuro,DecE)=0 then LettrageTotE:=True ;
   END else
   BEGIN
   if Arrondi(LTE_CouvDebP-LTE_CouvCredP-EcartChangeFranc,DecF)=0 then LettrageTotP:=True ;
  // if Arrondi(LTE_CouvDebE-LTE_CouvCredE-EcartChangeEuro,DecE)=0 then LettrageTotE:=True ;
   END ;
END ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 26/04/2002
Modifié le ... :   /  /
Description .. : Correction du calcul des ecart de change en euro
Mots clefs ... :
*****************************************************************}
procedure TFLettrage.GereEcartDeChange ;
Var Debit : boolean ;
    Cpte  : String ;
    lRdEcartPivot : double ;
BEGIN
if Not OkLettrage then Exit ;
{pasecart signifie, totalemenft lettré devise mais refus de procéder à écart de change ou journal/comptes non renseignés}
EcartChange:=False ; PasEcart:=False ;
if Not RL.LettrageDevise then Exit ;
if Not LettrageTotCur then Exit ;

if V_PGI.DevisePivot = DEV.Code then Exit;
{if EstMonnaieIN(DEV.Code) then Exit ;}

if IncoherECC then Exit ;
if LaDateGenere<V_PGI.DateDebutEuro then
   BEGIN
   {JLDEURO}
   END else
   BEGIN
   {Apres entrée en vigueur, cotation / Euro --> EcartChange <-> Diff sur les Euros}
   if ((VH^.TenueEuro) and (LettrageTotP)) then Exit ;
   END ;
if ((E_ECART.Value='') or (Not E_ECART.Enabled)) then BEGIN PasEcart:=True ; Exit ; END ;
if Not ModeCouverture then
   BEGIN
        // LG* 26/04/2002 correction des ecarts de change pour l'euro
        lRdEcartPivot:=EcartChangeFranc ;
        EcartChangeEuro:=Arrondi(LTE_TotDebP-LTE_TotCredP-lRdEcartPivot,V_PGI.OkDecV) ;
   END else
      EcartChangeEuro:=Arrondi(LTE_CouvDebP-LTE_CouvCredP-EcartChangeEuro,V_PGI.OkDecV) ;

Debit:=(EcartChangeEuro>0) ;
if Debit then Cpte:=DEV.CptDebit else Cpte:=DEV.CptCredit ;
if Cpte='' then BEGIN PasEcart:=True ; Exit ; END ;
if not FBoUneGrille then
BEGIN
 if DatePaquetMax<LaDateGenere then DatePaquetMax:=LaDateGenere ;
 if DatePaquetMin>LaDateGenere then DatePaquetMin:=LaDateGenere ;
END;
LettrageTotP:=True ;
EcartChange:=True ;
END ;

{==================================== CALCULS ==========================================}
procedure TFLettrage.RazValeurs ;
BEGIN
{OKOK}
LTE_TotDebD:=0 ; LTE_TotCredD:=0 ; LTE_TotDebP:=0 ; LTE_TotCredP:=0 ;
LTE_TotDebCur:=0 ; LTE_TotCredCur:=0 ;
LTE_CouvDebD:=0 ; LTE_CouvCredD:=0 ; LTE_CouvDebP:=0 ; LTE_CouvCredP:=0 ;
LTE_CouvDebCur:=0 ; LTE_CouvCredCur:=0 ;
RefReleve:='' ; NbReleve:=0 ; ExistTvaEnc:=False ; IncoherECC:=False ; AvantEuro:=False ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 24/12/2001
Modifié le ... : 10/10/2005
Description .. : Modif pour l'affichage des solde en mode pcl
Suite ........ : - 29/11/2002 - on n'affiche plus le sigle des devises dasn la
Suite ........ : case des soldes
Suite ........ : - 10/10/2005 - LG - FB 15690 - correction de l'afffichage du 
Suite ........ : solde s'il etait crediteur
Mots clefs ... : 
*****************************************************************}
procedure TFLettrage.AfficheLesSoldes ;
BEGIN
{OKOK}
AfficheLeSolde(SL_SOLDEDEBIT,0,0) ; AfficheLeSolde(SL_SOLDECREDIT,0,0) ;
AfficheLeSolde(SL_TOTALDEBIT,LTE_TotDebCur,0) ; AfficheLeSolde(SL_TOTALCREDIT,0,LTE_TotCredCur) ;
if Abs(LTE_TotDebCur)>Abs(LTE_TotCredCur) then
 begin
  AfficheLeSolde(SL_SOLDEDEBIT,LTE_TotDebCur-LTE_TotCredCur,0) ;
  AfficheLeSolde(SL_SOLDE,LTE_TotDebCur-LTE_TotCredCur,0) ;
 end
  else
   begin
    AfficheLeSolde(SL_SOLDECREDIT,0,LTE_TotCredCur-LTE_TotDebCur) ;
    AfficheLeSolde(SL_SOLDE,0,LTE_TotCredCur-LTE_TotDebCur) ;
   end ;// AfficheLeSolde(SL_SOLDE,0,LTE_TotCredCur-LTE_TotDebCur) ; END ;
//LG* 24/12/2001 affichage des solde en mode pcl
SL_SOLDEDEBITPCL.Value:=LTE_TotDebCur ; SL_SOLDECREDITPCL.Value:=LTE_TotCredCur ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 07/12/2001
Modifié le ... : 07/12/2001
Description .. : LG modif :
Suite ........ : -test les soldes ne sont pas au null pour proposé
Suite ........ : de validé le lettrage
Suite ........ : -gestion d'une solde grille pour le calcul des soldes
Mots clefs ... :
*************************************************** **************}
procedure TFLettrage.CalculDebitCredit ( Propose : boolean );
Var i : integer ;
    O : TOBM ;
    Premier : boolean ;
    StR  : String ;
    DMoins,DPlus,DD : TDateTime ;
    lBoTousLettre : boolean ;
BEGIN
{OKOK}
RazValeurs ; if Not OkLettrage then Propose:=False ;
DMoins:=V_PGI.DateEntree ; DPlus:=V_PGI.DateEntree ;
LeEtab:='' ; Premier:=True ; lBoTousLettre:=true ;
for i:=1 to GD.RowCount-1 do if JePrends(GD,i) then
    BEGIN
    O:=GetO(GD,i) ; if O=Nil then Break ;
    if Premier then BEGIN Premier:=False ; LeEtab:=O.GetMvt('E_ETABLISSEMENT') ; END
               else BEGIN if LeEtab<>O.GetMvt('E_ETABLISSEMENT') then LeEtab:='' ; END ;
    {Sommations}
    LTE_TotDebP:=LTE_TotDebP+O.GetMvt('E_DEBIT') ;
    LTE_TotDebD:=LTE_TotDebD+O.GetMvt('E_DEBITDEV') ;
    LTE_CouvDebP:=LTE_CouvDebP+O.GetMvt('E_COUVERTURE') ;
    LTE_CouvDebD:=LTE_CouvDebD+O.GetMvt('E_COUVERTUREDEV') ;
    LTE_TotDebCur:=LTE_TotDebCur+GetDebitCur(O) ;
    LTE_CouvDebCur:=LTE_CouvDebCur+GetCouvCur(O) ;
    lBoTousLettre:=lBoTousLettre and (O.GetMvt('E_ETATLETTRAGE') = 'TL') ;
    if FBoUneGrille then
     begin
      LTE_TotCredP:=LTE_TotCredP+O.GetMvt('E_CREDIT') ;
      LTE_TotCredD:=LTE_TotCredD+O.GetMvt('E_CREDITDEV') ;
      LTE_CouvCredP:=LTE_CouvCredP+O.GetMvt('E_COUVERTURE') ;
      LTE_CouvCredD:=LTE_CouvCredD+O.GetMvt('E_COUVERTUREDEV') ;
      LTE_TotCredCur:=LTE_TotCredCur+GetCreditCur(O) ;
      LTE_CouvCredCur:=LTE_CouvCredCur+GetCouvCur(O) ;
      {#TVAENC}
      if ( O.GetMvt('E_CREDIT') > 0 ) and ( O.GetMvt('E_EDITEETATTVA')<>'-') then ExistTvaEnc:=True ;
     end; // if
    {Relevé}
    StR:=O.GetMvt('E_REFRELEVE') ;
    if StR<>'' then
       BEGIN
       if ((RefReleve='') or (NbReleve<=0)) then BEGIN NbReleve:=1 ; RefReleve:=StR ; END
                                            else if StR<>RefReleve then NbReleve:=2 ;
       END ;
    DD:=O.GetMvt('E_DATECOMPTABLE') ;
    if DD>DPlus then DPlus:=DD ; if DD<DMoins then DMoins:=DD ;
    END else if Not OkLettrage then
    BEGIN
    {#TVAENC}
    O:=GetO(GD,i) ; if O=Nil then Break ;
    if O.GetMvt('E_EDITEETATTVA')<>'-' then ExistTvaEnc:=True ;
    END ;
for i:=1 to GC.RowCount-1 do if JePrends(GC,i) then
    BEGIN
    O:=GetO(GC,i) ; if O=Nil then Break ;
    if Premier then BEGIN Premier:=False ; LeEtab:=O.GetMvt('E_ETABLISSEMENT') ; END
               else BEGIN if LeEtab<>O.GetMvt('E_ETABLISSEMENT') then LeEtab:='' ; END ;
    {Sommations}
    LTE_TotCredP:=LTE_TotCredP+O.GetMvt('E_CREDIT') ;
    LTE_TotCredD:=LTE_TotCredD+O.GetMvt('E_CREDITDEV') ;
    LTE_CouvCredP:=LTE_CouvCredP+O.GetMvt('E_COUVERTURE') ;
    LTE_CouvCredD:=LTE_CouvCredD+O.GetMvt('E_COUVERTUREDEV') ;
    LTE_TotCredCur:=LTE_TotCredCur+GetCreditCur(O) ;
    LTE_CouvCredCur:=LTE_CouvCredCur+GetCouvCur(O) ;
    {#TVAENC}
    if O.GetMvt('E_EDITEETATTVA')<>'-' then ExistTvaEnc:=True ;
    {Relevé}
    StR:=O.GetMvt('E_REFRELEVE') ;
    if StR<>'' then
       BEGIN
       if ((RefReleve='') or (NbReleve<=0)) then BEGIN NbReleve:=1 ; RefReleve:=StR ; END
                                            else if StR<>RefReleve then NbReleve:=2 ;
       END ;
    DD:=O.GetMvt('E_DATECOMPTABLE') ;
    if DD>DPlus then DPlus:=DD ; if DD<DMoins then DMoins:=DD ;
    END else if Not OkLettrage then
    BEGIN
    {#TVAENC}
    O:=GetO(GC,i) ; if O=Nil then Break ;
    if O.GetMvt('E_EDITEETATTVA')<>'-' then ExistTvaEnc:=True ;
    END ;
AfficheLesSoldes ;
if ((DMoins<V_PGI.DateDebutEuro) and (DPlus>=V_PGI.DateDebutEuro)) then IncoherECC:=True ;
if ((DMoins<V_PGI.DateDebutEuro) or (V_PGI.DateEntree<V_PGI.DateDebutEuro)) then AvantEuro:=True ;
if Not ModeCouverture then
   BEGIN
   //LG* 07/12/2001 on ne propose pas la fenetre de validation des ecritures si elles sont toutes lettres
   if FBoUneGrille then
   begin
    if not lBoTousLettre and ((Propose) and (Arrondi(LTE_TotDebCur-LTE_TotCredCur,CurDec)=0)) and (LTE_TotDebCur<>0) then ProposeValide ;
   end
    else
      if ((Propose) and (Arrondi(LTE_TotDebCur-LTE_TotCredCur,CurDec)=0)) and (LTE_TotDebCur<>0) then ProposeValide ;
   END else
   BEGIN
   if ((Propose) and (Arrondi(LTE_CouvDebCur-LTE_CouvCredCur,CurDec)=0))and (LTE_TotDebCur<>0) then ProposeValide ;
   END ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 05/05/2006
Modifié le ... :   /  /    
Description .. : - 05/05/2006 - LG - FB 17712 - calcul de la date echeance 
Suite ........ : mx pour l'affecte a l'ecriture simplifie
Mots clefs ... : 
*****************************************************************}
Procedure TFLettrage.CodeDateLettre(var DMin,DMax, DEche : tDateTime ; var ModePaie : string ) ;
Var i  : integer ;
    O  : TOBM ;
    DD,DE : TDateTime ;
BEGIN
{OKOK}
CodeL:='' ;
DatePaquetMax:=IDate1900; DatePaquetMin:=IDate2099; DEche := iDate1900 ; ModePaie := 'DIV' ;
for i:=1 to GD.RowCount-1 do if JePrends(GD,i) then
    BEGIN
    O:=GetO(GD,i) ; if O=Nil then Break ;
    if GD.Cells[ColLetD,i]<>'' then CodeL:=Copy(GD.Cells[ColLetD,i],1,4) ;
    DD:=O.GetMvt('E_DATECOMPTABLE') ;
    DE:=O.GetMvt('E_DATEECHEANCE') ;
    if DD>DatePaquetMax then DatePaquetMax:=DD ; if DD<DatePaquetMin then DatePaquetMin:=DD ;
    if DE>DEche then
     begin
      DEche:=DE ;
      ModePaie:=O.GetMvt('E_MODEPAIE') ;
     end ;// if
    END ;
for i:=1 to GC.RowCount-1 do if JePrends(GC,i) then
    BEGIN
    O:=GetO(GC,i) ; if O=Nil then Break ;
    if GC.Cells[ColLetC,i]<>'' then CodeL:=Copy(GC.Cells[ColLetC,i],1,4) ;
    DD:=O.GetMvt('E_DATECOMPTABLE') ;
    if DD>DatePaquetMax then DatePaquetMax:=DD ; if DD<DatePaquetMin then DatePaquetMin:=DD ;
    END ;
DMin:=DatePaquetMin ; DMax:=DatePaquetMax ;
END ;

{================================ INIT ET DEFAUTS ======================================}
procedure TFLettrage.InitialiseLettrage ;
BEGIN
{OKOK}
GD.VidePile(True) ; GC.VidePile(True) ; PieceModifiee:=False ;
GD.RowCount:=2 ; GC.RowCount:=2 ; GD.TypeSais:=tsLettrage ; GC.TypeSais:=tsLettrage ;
LettrageTotP:=False ; LettrageTotCur:=False ;
ModeSelection:=False ; GereLesValide ;
Coldebit:=-1 ; ColCredit:=-1 ; NbSelD:=0 ; NbSelC:=0 ;
PasEcart:=False ; AvecRegul:=False ; EcartChange:=False ;
EcartChangeFranc:=0 ; EcartChangeEuro:=0 ; EcartRegulDevise:=0 ;
ConvertFranc:=0 ; ConvertEuro:=0 ; IndiceRegul:=0 ;
CParCouverture.Enabled:=True ; if ((LienS1S3) or (EstSerie(S3))) then CParCouverture.Enabled:=False ;
CParEnsemble.Enabled:=True ; CParExcept.Enabled:=True ;
RazValeurs ; VideListe(TLAD) ; VideListe(TLAC) ; RefReleve:='' ;
NbReleve:=0 ; ExistTvaEnc:=False ;
GSais:=Nil ; LigSais:=0 ; IncoherECC:=False ; AvantEuro:=False ;
END ;

procedure TFLettrage.EnleveOuRajoute( Enlev : boolean ) ;
BEGIN
{OKOK}
if Enlev then
   BEGIN
   BEnleve.Enabled:=True ; BRajoute.Enabled:=False ;
   BAjouteEnleve.Glyph:=BEnleve.Glyph ;
   BAjouteEnleve.Hint:=BEnleve.Hint ;
   END else
   BEGIN
   BEnleve.Enabled:=False ; BRajoute.Enabled:=True ;
   BAjouteEnleve.Glyph:=BRajoute.Glyph ;
   BAjouteEnleve.Hint:=BRajoute.Hint ;
   END ;
END ;


{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 18/12/2001
Modifié le ... : 30/03/2006
Description .. : LG Modif : avertissement s'il n'existe aucun journal de regul
Suite ........ : 
Suite ........ : LG* - 09/04/2002 - test des param de regul de lettrage si le 
Suite ........ : choix de validation par defaut est génération automatique et 
Suite ........ : affichage de la fenetre des param
Suite ........ : 
Suite ........ : - LG - 12/09/2002 - suppression du controle des paramsoc
Suite ........ : en eagl
Suite ........ : - LG - 16/06/2004 - si on a demande l'auxi suivant et qu'il 
Suite ........ : n'y en a pas, on ne ferme pas la fenetre
Suite ........ : - LG - 19/09/2005 - FB 16346 - on ferme la fenetre qd on 
Suite ........ : ferme on echap
Suite ........ : - LG - 29/3/2006 - FB 17619 - proccessmessage pour 
Suite ........ : couper les touches ( double f11 )
Suite ........ : - LG - 30/03/2006 - FB 16352 - on force la selection des 
Suite ........ : ecriture lettre
Mots clefs ... : 
*****************************************************************}
procedure TFLettrage.ShowLettrage ( Tout : boolean ) ;
Var i : integer ;
    OldSel : Boolean ;
BEGIN
{OKOK}
//CentreGrids ;
InitEcran(Tout) ; ErreurReseau:=False ; FBoAucunEnr := false ;
AvecTL:=True ; EnleveOuRajoute(True) ; 
if RL.Appel in [tlSaisieTreso,tlSaisieCour] then BEGIN CParExcept.Enabled:=False ; CParExcept.Checked:=False ; END ;
{POPAux.Enabled:= RL.Appel<>tlSaisieCour ;
BUP.Enabled:= POPAux.Enabled ;
BDown.Enabled:= POPAux.Enabled ; }
ModeCouverture:=(CParCouverture.Checked) ; DelettrageGlobal:=(CParEnsemble.Checked) ;
AjouteDevise(CDevise,V_PGI.DevisePivot) ; AjouteDevise(CDevise,V_PGI.DeviseFongible) ;
if Not ChargeMouvements then
   BEGIN
   if Action<>taConsult then
     BEGIN
     if OkLettrage then HLettre.Execute(13,'','') else HLettre.Execute(27,'','') ;
     END else HLettre.Execute(20,'','') ;
   if FBoAuxSuiv then
    exit
    else
     begin PostMessage(Handle,WM_CLOSE,0,0) ; Exit ; end ;
   END ;
if ((ColDebit<0) or (ColCredit<0)) then BEGIN HLettre.Execute(0,'','') ; PostMessage(Handle,WM_CLOSE,0,0) ; END ;
if ((RL.Appel=tlMenu) and (Not OkLettrage) and (RL.ToutSelDel)) then
   BEGIN
   OldSel:=DelettrageGlobal ; DelettrageGlobal:=False ;
   GSais:=GD ; GSais.SetFocus ; for i:=1 to GD.RowCount-2 do BEGIN GD.Row:=i ; Selection(True) ; END ;
   if not FBoUneGrille then
    BEGIN
     GSais:=GC ; GSais.SetFocus ; for i:=1 to GC.RowCount-2 do BEGIN GC.Row:=i ; Selection(True) ; END ;
    END;
   GD.SetFocus ; GD.Row:=1 ; GC.Row:=1 ; DetailLigne(1,True) ; DetailLigne(1,False) ;
   DelettrageGlobal:=OldSel ;
   END else if ((RL.Appel<>tlMenu) and (GSais<>Nil) and (LigSais>0)) then
   BEGIN
   if GSais=GD then
      BEGIN
      GD.SetFocus ;
      if LigSais < GD.RowCount then
       begin
        GD.Row:=LigSais ;
        DetailLigne(LigSais,True) ;
       end ;
       DetailLigne(1,False) ;
      END else
      if not FBoUneGrille then
      BEGIN
      //LG* 27/11/2001 modif quand on une seule grille de selection
      GC.SetFocus ; GC.Row:=LigSais ;
      DetailLigne(1,True) ; DetailLigne(LigSais,False) ;
      END ;
   Selection(True) ;
   //LG* 09/01/2002 modif qd qd une seule grille
   if not FBoUneGrille then BEGIN if (GSais=GD) then GC.SetFocus else GD.SetFocus ; END ;
   END else
   BEGIN
   GD.SetFocus ;
   DetailLigne(1,True) ; DetailLigne(1,False) ;
   END ;
ActiveGrid ; CalculDebitCredit(False) ;
//LG* 18/12/2001
if not ExisteSQL('select j_journal from journal where J_FERME="-" AND J_NATUREJAL="REG" AND J_MODESAISIE<>"LIB"') then
 BEGIN
  PGIInfo('Il n''existe aucun journal de régularisation, vous ne pourrez pas générer d''écriture de regul',Caption);
  FBoAucunJournal:=true;
 END
  else
   FBoAucunJournal:=false;

{$IFNDEF EAGLCLIENT}
if (FIFL.ChoixValid='AL5') and
   ( (trim(FIFL.Journal)='') or (trim(FIFL.Libelle)='') or (trim(FIFL.GeneralDebit)='') or (trim(FIFL.GeneralCredit)='') )then
 begin
  PGIInfo('Les paramètres d''écriture de régularisation ne sont pas correctement renseignés',Caption);
  ParamSociete(False,'','SCO_LETTRAGEDETAIL','',RechargeParamSoc,ChargePageSoc,SauvePageSoc,InterfaceSoc,1105000) ;
  RechargeParamSoc ; FIFL:=ChargeInfosLett ; if RL.LettrageEnSaisie then FIFL.ChoixValid := 'AL4' ;
  ShowLettrage(Tout) ;
 end; // if
{$ENDIF}
if FBoFromTOB then
 begin
  Application.ProcessMessages ;
  ClickToutSel(true) ;
  ClickValide ; 
  if not FBoFermer then ClickValide ;
  PostMessage(self.Handle, WM_CLOSE, 0, 0);
 end ;


END ;

function TFLettrage.AttribTitre : string ;
Var St : String ;
BEGIN
{OKOK}
if OkLettrage then
   BEGIN
   if RL.Appel=tlMenu then St:=HDivL.Mess[0] else St:=HDivL.Mess[1] ;
   END else
   BEGIN
   if Action<>taConsult then St:=HDivL.Mess[2] else St:=HDivL.Mess[9] ;
   BValide.Hint:=HDivL.Mess[10] ; BValideSelect.Hint:=HDivL.Mess[11] ;
   END ;
result:=St
//UpdateCaption(Self) ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 20/11/2001
Modifié le ... : 26/06/2006
Description .. : LG Modif :
Suite ........ : - prise en compte de la nouvelle ListeParam LETTPCL
Suite ........ : - Gestion des panels details
Suite ........ : - gestion de l'affichage de la grille credit
Suite ........ : 
Suite ........ : - LG - 26/04/2002 - le bouton de lancement de lettrage
Suite ........ : auto n'est plus valide pour un lettrage en devise ( il faut
Suite ........ : choisir un journal pour le rendre actif )
Suite ........ : 
Suite ........ : - 10/06/2002 - affectation journal de change avec le journal
Suite ........ :  definie dans les parametres pour un lettrage en devise
Suite ........ : - on n'affiche pas l'info sur le solde de selection
Suite ........ : -01/04/2003 - LG - suppression de l'affichage de la date de
Suite ........ : generation
Suite ........ : - 14/10/2003 - LG - on n'a pas acces au ecriture simplifie
Suite ........ : en lettrage en saisie
Suite ........ : - 15/09/2004 - LG - FB 14568 - blocage du defilement
Suite ........ : des comptes pour els TIC ou TID
Suite ........ : - 03/04/2006 - LG  - FB 10757 - param sur l'ordre de tri
Suite ........ : - 07/03/2006 - LG - FB 10758 - Masquer par défaut les 
Suite ........ : écritures totalement lettrées
Suite ........ : - 14/06/2006 - LG - FB 18375 - l'option "masquer les 
Suite ........ : écritures..." ne marchait aps avec le delettrage
Suite ........ : - 26/06/2006 - LG - FB 16321 - on bloque le defilment des 
Suite ........ : auxi pour un tic ou tid
Mots clefs ... : 
*****************************************************************}
procedure TFLettrage.InitEcran ( Tout : boolean ) ;
BEGIN
GereNature ; Caption:=AttribTitre ;
if RL.Appel<>tlMenu then CParCouverture.Enabled:=False ;
//LG*
BReduire.Visible:=false;
if FBoUneGrille then BEGIN GD.ListeParam:='LETTPCL'; GC.ListeParam:='LETCLICRE' ; END
else if Client then BEGIN GD.ListeParam:='LETCLIDEB' ; GC.ListeParam:='LETCLICRE' ; END
else BEGIN GD.ListeParam:='LETFOUDEB' ; GC.ListeParam:='LETFOUCRE' ; END ;
PrepareSelect(GD) ; PrepareSelect(GC) ;
GetInfosDevise(DEV) ; DEV.Taux:=GetTaux(DEV.Code,DEV.DateTaux,V_PGI.DateEntree) ;
DEVF:=DEV ;
if RL.LettrageDevise then
   BEGIN
   E_REGUL.DataType:='ttJalRegulDevise' ;
   ChangeFormatDevise(Self,DEV.Decimale,DEV.Symbole) ;
   ChangeMask(E_TOTAL,V_PGI.OkDecV,V_PGI.SymbolePivot) ; ChangeMask(E_TOTAL_,V_PGI.OkDecV,V_PGI.SymbolePivot) ;

   {if EstMonnaieIn(DEV.Code) then BEGIN H_ECART.Enabled:=False ; E_ECART.Enabled:=False ; END ;}
   if V_PGI.DevisePivot = DEV.Code then BEGIN H_ECART.Enabled:=False ; E_ECART.Enabled:=False ; END ;
   END else
   BEGIN
   H_ECART.Enabled:=False ; E_ECART.Enabled:=False ;
   ChangeFormatPivot(Self) ;
   ChangeMask(E_TOTAL,DEV.Decimale,DEV.Symbole) ; ChangeMask(E_TOTAL_,DEV.Decimale,DEV.Symbole) ;
   END ;
if Tout then
   BEGIN
   if OkLettrage then
      BEGIN
      CParEnsemble.Visible:=False ;
      if (RL.LettrageDevise)  then CParExcept.Visible:=False ;
      END else
      BEGIN
      CParCouverture.Visible:=False ; CParExcept.Visible:=False ;
      CParEnsemble.Top:=CParCouverture.Top ; CParEnsemble.Left:=CParCouverture.Left ;
      LeSolde.Visible:=False ; BVSolde.Visible:=False ; TSoldes.Visible:=False ;
      E_ECART.Visible:=False ; E_REGUL.Visible:=False ; H_ECART.Visible:=False ; H_REGUL.Visible:=False ;
      DateGeneration.Enabled:=False ;
      BDelettre.Visible:=False ;
      END ;
   END ;
if RL.LettrageDevise then CurDec:=DEV.Decimale else CurDec:=V_PGI.OkDecV ;
H_TOTAL.Caption:=RechDom('TTDEVISETOUTES',V_PGI.DevisePivot,False) ;
H_TOTAL_.Caption:=H_TOTAL.Caption ;
ChangeMask(E_TOTAL,V_PGI.OkDecV,'') ; ChangeMask(E_TOTAL_,V_PGI.OkDecV,'') ;
//LG*
GC.Visible:=not FBoUneGrille ; // on utilise une seule grille
BTri.Visible:=FBoUneGrille ;
DeLettPartiel.Visible:=FBoUneGrille ;
DATEGENERATION.Visible:=not FBoUneGrille ;
TDATEGENERATION.Visible:=not FBoUneGrille ;
DateGeneration.Enabled:=not FBoUneGrille ;
BForceVadide.Visible := FBoUneGrille ;
// gestion du bon panel detail affiché
TSDetail.TabVisible:=false ; TSDetail1.TabVisible:=false ; BCValide.Visible:=FBoUneGrille ;
TSSoldePME.TabVisible:=false ; TSSoldePCL.TabVisible:=false ;
BCombi.Visible:=not FBoUneGrille ;
//LG* les methode de calcul ne sont pas active pour un lettrage en regul sans le choix du journal
BCalcul.Visible:= FBoUneGrille ;
// YMO FQ16966 07/11/05 si pas visible, pas enabled
BCalcul.Enabled:= (FBoUneGrille) and not (RL.LettrageDevise ) ;
Zoom1.Enabled:=BCalcul.Enabled ;
BUnique.Visible:=not FBoUneGrille ;
// initialisation des combos de l'onglet parametre
E_REGUL.Value:=FIFL.Journal;
if FBoUneGrille then
 BEGIN // positionnement des panels
  PCDetail.activepage:=TSDetail1 ; PCDetail.Height:=60 ;
  PCSoldeS.Height:=65 ; PCSoldes.Activepage:=TSSoldePCL;
  // en mode pcl on debranche l'evenement qui saute les lignes de hauteur 0
  GD.OnCellEnter:=nil ;
  // initialisation des combos de l'onglet parametre
  BvSolde.Visible:=false ;
  TSoldes.Visible:=false ; LESOLDE.Visible:=false ;
  if RL.LettrageDevise and (trim(FIFL.JournalChange)<>'') then E_ECART.Value:=FIFL.JournalChange ;
  Passagedunecrituresimplifi1.Enabled:=not RL.LettrageEnSaisie ;
  Soldeautomatique1.Enabled:=not RL.LettrageEnSaisie ;
 END
  else BEGIN PCDetail.activepage:=TSDetail ; PCSoldes.Height:=40 ; PCSoldes.Activepage:=TSSoldePME ; END ;
ChangeMask(HNBS,0,'') ; ChangeMask(HNBN,0,'') ; ChangeMask(HNBP,0,'') ;
// BPY le 08/10/2003 : Fiche de com FFF n°68
if (not OkLettrage) then
begin
    BForceVadide.Visible := false;
    BCalcul.Visible := false;
    BAjouteEnleve.Visible := false;
end;
// BPY
{$IFNDEF EAGLCLIENT}
POPAux.Enabled:= ( RL.Appel<>tlSaisieCour ) and ( HT_AUXILIAIRE.Caption <> '' ) ;
BUP.Enabled:= POPAux.Enabled ;
BDown.Enabled:= POPAux.Enabled ;
BUp.Visible := POPAux.Enabled ;
BDown.Visible := POPAux.Enabled ;
{$ELSE}                  //SG6 08/11/2004
BUp.Visible := false;
BDown.Visible := false;
{$ENDIF}
if FBoFromTOB then exit ;
if GetParamSocSecur('SO_LETPREF','') = 'L1' then
 BTri.Tag := 2
  else
   BTri.Tag := 1 ;
//if not RL.LettrageEnSaisie then
// begin
  BForceVadideClick(nil) ;
  FBoOter := okLettrage and GetParamSocSecur('SO_LETPREFTOT',false) ;
{ end
  else
   FBoOter := false ;  }
END ;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 29/11/2001
Modifié le ... :   /  /
Description .. : LG : Modif
Suite ........ :  - Ajout de la recherche su sens du compte ( pour l'affichage
Suite ........ : des couleurs sur le solde progressif )
Mots clefs ... :
*****************************************************************}
procedure TFLettrage.GereNature ;
Var Q : TQuery ;
    StN : String ;
    IsFerme : boolean ;
BEGIN
{OKOK}
StN:='' ; HG_GENERAL.Caption:='' ; HT_AUXILIAIRE.Caption:='' ;
IsFerme:=False ;
BVSolde.SetBounds(LeSolde.Left-1,LeSolde.Top-1,LeSolde.Width+2,LeSolde.Height+2) ;
if (trim(RL.Auxiliaire)<>'') then
   BEGIN
   Q:=OpenSQL('Select T_LIBELLE, T_NATUREAUXI, T_DERNLETTRAGE, T_FERME, T_ISPAYEUR from TIERS Where T_AUXILIAIRE="'+RL.Auxiliaire+'"',True) ;
   if Not Q.EOF then
      BEGIN
      HT_AUXILIAIRE.Caption:=RL.Auxiliaire+'  '+Q.FindField('T_LIBELLE').AsString ;
      StN:=Q.FindField('T_NATUREAUXI').AsString ; Client:=((StN<>'FOU') and (StN<>'AUC')) ;
      DernLettrage:=Q.FindField('T_DERNLETTRAGE').AsString ; SauveCode:=Dernlettrage ;
      IsFerme:=(Q.FindField('T_FERME').AsString='X') ; H_CodeL.Caption:=DernLettrage ;
      if Q.FindField('T_ISPAYEUR').AsString='X' then BEGIN BTF.Down:=True ; BTFClick(Nil) ; END ;
      END ;
   Ferme(Q) ;
   END ;
if RL.General<>'' then
   BEGIN
   //LG* Recherche du sens du compte
   Q:=OpenSQL('Select G_LIBELLE, G_NATUREGENE, G_DERNLETTRAGE, G_FERME, G_TOTALDEBIT, G_TOTALCREDIT,G_SENS from GENERAUX Where G_GENERAL="'+RL.General+'"',True) ;
   if Not Q.EOF then
      BEGIN
      HG_GENERAL.Caption:=RL.General+'  '+Q.FindField('G_LIBELLE').AsString ;
      FStSensCompte:=UpperCase(Q.FindField('G_SENS').AsString);
      if StN='' then
         BEGIN
         StN:=Q.FindField('G_NATUREGENE').AsString ; Client:=(StN<>'TIC') ;
         DernLettrage:=Q.FindField('G_DERNLETTRAGE').AsString ; SauveCode:=DernLettrage ;
         IsFerme:=(Q.FindField('G_FERME').AsString='X') ; H_CodeL.Caption:=DernLettrage ;
         END ;
      END ;
   Ferme(Q) ;
   END ;
if ((IsFerme) and (Action<>taConsult)) then HLettre.Execute(26,caption,'') ;
if Not OkLettrage then
   BEGIN
   if RL.CodeLettre<>'' then H_CodeL.Caption:=RL.CodeLettre ;
   E_REFLETTRAGE.Enabled:=False ;
   END ;
END ;

{================================= CHARGEMENTS =========================================}

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 07/12/2001
Modifié le ... :   /  /
Description .. : LG modif : ajout de ColSolde
Mots clefs ... :
*****************************************************************}
procedure TFLettrage.AttribColonnes ( G : THGrid ) ;
Var StC,STitreCol : String ;
    C  : integer ;
BEGIN
{OKOK}
//LG* 07/12/2001
StC:=G.Titres[0] ; C:=0 ;
Repeat
 STitreCol:=ReadTokenSt(StC) ;
 if Pos('E_DEBIT',STitreCol)>0 then
  ColDebit:=C ;
 if Pos('E_CREDIT',STitreCol)>0 then
  ColCredit:=C ;
 if Pos('SOLDEPRO',STitreCol)>0 then
  ColSolde:=C ;
 if Pos('E_LETTRAGE',STitreCol)>0 then BEGIN if G.Name='GD' then ColLetD:=C else ColLetC:=C ; END ;
 inc(C) ;
Until ((StC='') or (STitreCol='') or (C>=G.ColCount)) ;
END ;

function TFLettrage.EcheSaisie ( O : TOBM ) : boolean;
BEGIN
{OKOK}
Result:=False ;
if Action=taConsult then Exit ;
if RL.Appel=tlMenu then Exit ;
if ((GSais<>Nil) and (LigSais>0)) then Exit ;
if RL.Ident.Jal<>O.GetMvt('E_JOURNAL') then Exit ;
if RL.Ident.DateC<>O.GetMvt('E_DATECOMPTABLE') then Exit ;
if RL.Ident.Num<>O.GetMvt('E_NUMEROPIECE') then Exit ;
if RL.Ident.Nature<>O.GetMvt('E_NATUREPIECE') then Exit ;
if RL.Ident.NumLigne<>O.GetMvt('E_NUMLIGNE') then Exit ;
if RL.Ident.NumEche<>O.GetMvt('E_NUMECHE') then Exit ;
Result:=True ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Lauretn GENDREAU
Créé le ...... : 20/11/2001
Modifié le ... : 07/04/2006
Description .. : LG Modif :
Suite ........ : -si une seule grille, on ne relance pas attribColonnes pour ne
Suite ........ : pas perdre les indices des colonnes debit et credit (les deux
Suite ........ : grilles n'ont pas les memes colonnes dans ce cas là )
Suite ........ : -modif de la creation des deux TList TLAD,TLAC servant
Suite ........ : aux lettrages combinatoires. En mode PCL les deux listes on
Suite ........ : la taille de la grille GD
Suite ........ : rem : ils sont pre-remplie de  structure vide ayant les
Suite ........ : dimension des grilles
Suite ........ : 
Suite ........ : - LG - 09/04/2002 - utilisation de la fct
Suite ........ : CAjouteChampsSuppLett
Suite ........ : 
Suite ........ : - LG - 23/04/2002 - les lignes d'ecart de conversion ne sont
Suite ........ : plus affichés dans la grille
Suite ........ : 
Suite ........ : - LG - 10/06/2002 - on affiche le cumul de la selection dans
Suite ........ : la monnaie de lettrage
Suite ........ : 
Suite ........ : - LG - 15/10/2002 - le calcul des soldes etaient faux ( on
Suite ........ : additionnait des motants en devise et en euro
Suite ........ : 
Suite ........ : - LG - 13/02/2003 - FB 10219 - on marque les
Suite ........ : enregistremens utilises dasn un bordereau
Suite ........ : -LG  - 01/02/2005 - FB 17415 - le test sur la table courrier 
Suite ........ : etait fait sur chq ligne
Suite ........ : - 07/03/2006 - LG - FB 10758 - Masquer par défaut les 
Suite ........ : écritures totalement lettrées
Mots clefs ... : 
*****************************************************************}
procedure TFLettrage.RempliGrids ( Q : TQuery ) ;
Var OBM : TOBM ;
    G   : THGrid ;
    CD,Etab,StRef,SDebit,SCredit : String ;
    P   : P_LAS ;// definie dans LettAuto
    DD  : TDateTime ;
    Avant,Apres,Premier : Boolean ;
    TotDebit,TotCredit : Double ;
    s : double ;
    lTOBListeBor : TOB ;
    i : integer ;
BEGIN
{OKOK}
Avant:=False ; Apres:=False ; Premier:=True ;
if V_PGI.DateEntree<V_PGI.DateDebutEuro then Avant:=True else Apres:=True ;
AttribColonnes(GD) ;            
//LG* 24/12/2001 on relance le centre grid pour reposition les controle en fonction des colonnes debit/credit
CentreGrids ;
SDebit:='E_DEBIT' ; SCredit:='E_CREDIT' ;
if not FBoUneGrille then AttribColonnes(GC) ;
TotDebit:=0 ; TotCredit:=0 ; s:=0 ;
lTOBListeBor:=CGetListeBordereauBloque ;
if FBoFromTOB then
 for i := 0 to FTOBEcr.Detail.Count - 1 do
  begin
    OBM := TOBM(FTOBEcr.Detail[i]) ;
    if FBoUneGrille then G:=GD
    else if OBM.GetValue('E_DEBIT')<>0 then G:=GD else G:=GC ;
    CAjouteChampsSuppLett(OBM) ;
   if CEstBloqueEcritureBor(TOB(OBM),lTOBListeBor) then
    begin
     OBM.PutMvt('E_CONTROLEUR','X') ;
     OBM.Etat:=Inchange ;
    end ; // if
   OBM.PutMvt('ENBASE','X') ; // ecriture presente en base, sert pour les ecriture de simul on ne le supprime pas d e la grill ds ce cas la
   POPAnnulTotal.Enabled:=(POPAnnulTotal.Enabled or (OBM.GetValue('E_LETTRAGE')<>'')) and OkLettrage ;
   TotDebit:=Arrondi(TotDebit+GetDebitCur(OBM),V_PGI.OkDecV) ; // LG - 15/10/2002 - correction du calcul du solde
   TotCredit:=Arrondi(TotCredit+GetCreditCur(OBM),V_PGI.OkDecV) ;
   AnalyseTitreMontant(SDebit,RL.LettrageDevise) ; AnalyseTitreMontant(SCredit,RL.LettrageDevise) ;
   s:=s+OBM.GetMvt(SDebit)-OBM.GetMvt(SCredit) ;
   OBM.PutValue('SOLDEPRO',s);
   Etab:=OBM.GetMvt('E_ETABLISSEMENT') ;
   if Premier then LeEtab:=Etab else if ((LeEtab<>'') and (LeEtab<>Etab)) then LeEtab:='' ;
   DD:=OBM.GetMvt('E_DATECOMPTABLE') ; Premier:=False ;
   if DD<V_PGI.DateDebutEuro then Avant:=True else Apres:=True ;
   //LG*
   G.Row:=G.RowCount-1;
   G.Objects[G.ColCount-1,G.RowCount-1]:=OBM ;
   CTOBVersTHGrid(G,OBM,RL.LettrageDevise) ;
   if Not RL.LettrageDevise then
      BEGIN
      CD:=OBM.GetMvt('E_DEVISE') ;
      if ((CD<>V_PGI.DevisePivot) and (CD<>V_PGI.DeviseFongible)) then AjouteDevise(CDevise,CD) ;
      END ;
   if OkLettrage then
      BEGIN
      P:=P_LAS.Create ; FillChar(P.S,Sizeof(P.S),slPasSolution) ;
      //LG* 16/01/2002 les liste ont la taille des la grille GD
      if FBoUneGrille then BEGIN TLAD.Add(P) ; P:=P_LAS.Create ; FillChar(P.S,Sizeof(P.S),slPasSolution) ; TLAC.Add(P) ; END
      else BEGIN if G=GD then TLAD.Add(P) else TLAC.Add(P) ; END ;
      if ((RL.Appel<>tlMenu) and (EcheSaisie(OBM))) then
       BEGIN
        GSais:=G ; LigSais:=G.RowCount-1 ;
       END ;
      END else
      BEGIN
      StRef:=OBM.GetMvt('E_REFLETTRAGE') ;
      if StRef<>'' then E_REFLETTRAGE.Text:=StRef ;
      END ;
    G.RowCount:=G.RowCount+1 ;
  end
  else
While Not Q.EOF do
   BEGIN
   if FBoUneGrille then G:=GD
   else if Q.FindField('E_DEBIT').AsFloat<>0 then G:=GD else G:=GC ;
   OBM:=TOBM.Create(EcrGen,'',false,FTOBEcr) ; CAjouteChampsSuppLett(OBM) ;
   OBM.ChargeMvt(Q) ;
   if CEstBloqueEcritureBor(TOB(OBM),lTOBListeBor) then
    begin
     OBM.PutMvt('E_CONTROLEUR','X') ;
     OBM.Etat:=Inchange ;
    end ; // if
   OBM.PutMvt('ENBASE','X') ; // ecriture presente en base, sert pour les ecriture de simul on ne le supprime pas d e la grill ds ce cas la
   POPAnnulTotal.Enabled:=(POPAnnulTotal.Enabled or (OBM.GetValue('E_LETTRAGE')<>'')) and OkLettrage ;
   TotDebit:=Arrondi(TotDebit+GetDebitCur(OBM),V_PGI.OkDecV) ; // LG - 15/10/2002 - correction du calcul du solde
   TotCredit:=Arrondi(TotCredit+GetCreditCur(OBM),V_PGI.OkDecV) ;
   AnalyseTitreMontant(SDebit,RL.LettrageDevise) ; AnalyseTitreMontant(SCredit,RL.LettrageDevise) ;
   s:=s+OBM.GetMvt(SDebit)-OBM.GetMvt(SCredit) ;
   OBM.PutValue('SOLDEPRO',s);
   Etab:=OBM.GetMvt('E_ETABLISSEMENT') ;
   if Premier then LeEtab:=Etab else if ((LeEtab<>'') and (LeEtab<>Etab)) then LeEtab:='' ;
   DD:=OBM.GetMvt('E_DATECOMPTABLE') ; Premier:=False ;
   if DD<V_PGI.DateDebutEuro then Avant:=True else Apres:=True ;
   //LG*
   G.Row:=G.RowCount-1;
   G.Objects[G.ColCount-1,G.RowCount-1]:=OBM ;
   CTOBVersTHGrid(G,OBM,RL.LettrageDevise) ;
   if Not RL.LettrageDevise then
      BEGIN
      CD:=OBM.GetMvt('E_DEVISE') ;
      if ((CD<>V_PGI.DevisePivot) and (CD<>V_PGI.DeviseFongible)) then AjouteDevise(CDevise,CD) ;
      END ;
   if OkLettrage then
      BEGIN
      P:=P_LAS.Create ; FillChar(P.S,Sizeof(P.S),slPasSolution) ;
      //LG* 16/01/2002 les liste ont la taille des la grille GD
      if FBoUneGrille then BEGIN TLAD.Add(P) ; P:=P_LAS.Create ; FillChar(P.S,Sizeof(P.S),slPasSolution) ; TLAC.Add(P) ; END
      else BEGIN if G=GD then TLAD.Add(P) else TLAC.Add(P) ; END ;
      if ((RL.Appel<>tlMenu) and not FBoOter and (EcheSaisie(OBM))) then // Fb 21571 - on ne pre selectionne pas la ligne ici ca serai fait ds le RemettreGrilleAJOur
       BEGIN
        GSais:=G ; LigSais:=G.RowCount-1 ;
       END ;
      END else
      BEGIN
      StRef:=OBM.GetMvt('E_REFLETTRAGE') ;
      if StRef<>'' then E_REFLETTRAGE.Text:=StRef ;
      END ;
    G.RowCount:=G.RowCount+1 ; Q.Next ;
   END ;
   
if ((Avant) and (Apres)) then BEGIN E_ECART.Value:='' ; E_ECART.Enabled:=False ; END ;
//LG* 24/12/2001 affiche des solde en mode pcl
SL_CUMULDEBIT.Value:=TotDebit ; SL_CUMULCREDIT.Value:=TotCredit ;
AfficheLeSolde(LeSolde,TotDebit,TotCredit) ; SL_SOLDEPCL.Value:=LeSolde.Value ;
if FBoUneGrille then
 BEGIN
  PostMessage(GD.Handle, WM_KEYDOWN, VK_HOME, 0) ;
  if FBoOter then RemettreGrilleAJour(GD,FBoOter) ;
 END;
lTOBListeBor.Free ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 20/11/2001
Modifié le ... : 21/06/2006
Description .. : LG Modif de l'ordre de l'order by pour tenir compte de l'ordre
Suite ........ : chronologique
Suite ........ : - 30/06/2003 - FB 12342 - deux new param ds 
Suite ........ : LWhereBase : si on lettre en saisie, et le types d'ecritures a 
Suite ........ : lettres ( TL, PL, AL )
Suite ........ : - LG - 30/06/2004 - FB 12989 12960 - on utilise la meme 
Suite ........ : requet pour la consultation et le lettrage
Suite ........ : - LG - 21/06/2006 - FB 18428 - on ne presente plus les 
Suite ........ : ecritures anterieur a l'exo V8 qd on appelle le lettrage depuis 
Suite ........ : la consultation des ecritures
Mots clefs ... : 
*****************************************************************}
function TFLettrage.ChargeMouvements : boolean ;
Var St : String ;
    Q  : TQuery ;
BEGIN
if FBoFromTOB then begin RempliGrids(nil) ; result := true ; exit ; end ;
BeginTrans ;
try
Result:=True ;
if FBoFromSaisie then
 begin
  if VH^.ExoV8.Code<>'' then
   RL.CritMvt := RL.CritMvt + ' AND E_DATECOMPTABLE>="'+UsDateTime(VH^.ExoV8.Deb)+'" ' ;
  if FInOrdreTri=1 then
   St:='Select * from Ecriture' + RL.CritMvt + ' order by E_LETTRAGE,E_AUXILIAIRE,E_GENERAL,E_DATECOMPTABLE,E_NUMEROPIECE,E_NUMLIGNE,E_NUMECHE '
    else
     St:='Select * from Ecriture' + RL.CritMvt +  'order by E_AUXILIAIRE,E_GENERAL,E_DATECOMPTABLE,E_NUMEROPIECE,E_NUMLIGNE,E_NUMECHE ' ;
 end
  else
   begin
    St:='Select * from Ecriture Where '+WhereCptes
        +' AND '+LWhereBase(OkLettrage,False,RL.Appel=tlSaisieTreso,RL.SansLesPartiel,RL.LettrageEnSaisie,RL.CritEtatLettrage) ;
    //LG* modif de l'ordre de tri
    if FBoUneGrille then
     St:=St+RL.CritMvt+' order by E_AUXILIAIRE, E_GENERAL,E_DATECOMPTABLE ,E_NUMEROPIECE,E_NUMLIGNE,E_NUMECHE '
     else
     St:=St+RL.CritMvt+LTri ;
   end ;
Q:=OpenSQL(St,True) ;
if Not Q.EOF then
 RempliGrids(Q)
  else
   begin
    Result:=False ;
    FBoAucunEnr := true ;
   end ;
Ferme(Q) ;

finally
CommitTrans ;
end;
END ;

{================================ BARRE OUTILS =========================================}
{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 10/04/2002
Modifié le ... : 19/06/2006
Description .. : On visualise pas les ecritures simplifiees, car elles ne
Suite ........ : sont pas forcement enregistees
Suite ........ : - LG - 18/08/2004 - Suppression de la fct debutdemois pour 
Suite ........ : l'appel de la saisie bor, ne fct pas avec les exercices 
Suite ........ : decalees
Suite ........ : - LG - 19/6/2006 - FB 15992 - en deletrage on active le 
Suite ........ : zoom sur les ecritures simplifies
Mots clefs ... : 
*****************************************************************}
procedure TFLettrage.ClickZoomPiece ;
Var O  : TOBM ;
    M  : RMVT ;
{$IFNDEF IMP}
    P  : RParFolio ;
{$ENDIF}
BEGIN
{OKOK}
{$IFNDEF IMP}
if ((GD.Focused) or (CurGrid=1)) then O:=GetO(GD,GD.Row) else
 if ((GC.Focused) or (CurGrid=2)) then O:=GetO(GC,GC.Row) else Exit ;
if O=Nil then Exit ;
if ( O.GetMvt('E_QUALIFORIGINE')='SIM' ) and ( OkLettrage ) then exit;
M:=OBMToIdent(O,False) ;
if ((M.ModeSaisieJal<>'-') and (M.ModeSaisieJal<>'')) then
   BEGIN
   FillChar(P, Sizeof(P), #0) ;
   P.ParPeriode:=DateToStr(O.GetMvt('E_DATECOMPTABLE')) ;
   P.ParCodeJal:=O.GetMvt('E_JOURNAL') ;
   P.ParNumFolio:=IntToStr(O.GetMvt('E_NUMEROPIECE')) ;
   P.ParNumLigne:=O.GetMvt('E_NUMLIGNE') ;
   ChargeSaisieFolio(P, taConsult) ;
   END else
   BEGIN
   LanceSaisie(Nil,taConsult,M) ;
   END ;
{$ENDIF}
END ;

procedure TFLettrage.ClickZoom ;
BEGIN
{OKOK}
if RL.Auxiliaire<>'' then
 FicheTiers(Nil,'',RL.Auxiliaire,taConsult,1)
 else
  FicheGene(Nil,'',RL.General,taConsult,0)
END ;


{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 10/01/2002
Modifié le ... : 26/06/2006
Description .. : Remettre la grille à jour pour recalculer le solde de progressif
Suite ........ : - LG - 07/06/2006 - FB 18314 - correction de la gestion de 
Suite ........ : l'affichage des totalement lettres
Suite ........ : - LG - 26/06/2006 - FB 18464 - le param FBOter est pri en 
Suite ........ : compte a l'ouverture du lettrage, apres on utilise la veur du 
Suite ........ : bouton 
Mots clefs ... : 
*****************************************************************}
procedure TFLettrage.ClickEnleve(Oter : boolean) ;
var
 G : THGrid;
begin
if Action=taConsult then Exit ;
if not FBoUneGrille then ClickEnlevePME(Oter)
else
 BEGIN
  if GD.Focused then G:=GD else if GC.Focused then G:=GC else Exit ;
  G.SynEnabled:=false ;
  try
  if Not OkLettrage then Exit ; if ModeSelection then Exit ;
  G.BeginUpdate ; Enleverlestotalementslettrs1.Enabled:=not Oter ; Remettrelestotalementslettres1.Enabled:=Oter ;
  FBoOter:=Oter ;
  RemettreGrilleAJour(G,Oter) ;
  finally
   G.SynEnabled:=true ; G.EndUpdate ;
  end;
 END; // if
end;

procedure TFLettrage.ClickEnlevePME(Oter : boolean) ;
Var St : String ;
    i  : integer ;
begin
{OKOK}
if Action=taConsult then Exit ;
if Not OkLettrage then Exit ; if ModeSelection then Exit ; FBoOter:=Oter ;
if Oter then
   BEGIN
   for i:=1 to GD.RowCount-1 do BEGIN St:=Copy(GD.Cells[ColLetD,i],1,4) ; if ((St<>'') and (St=uppercase(St))) then GD.RowHeights[i]:=0 ; END ;
   for i:=1 to GC.RowCount-1 do BEGIN St:=Copy(GC.Cells[ColLetC,i],1,4) ; if ((St<>'') and (St=uppercase(St))) then GC.RowHeights[i]:=0 ; END ;
   AvecTL:=False ; EnleveOuRajoute(False) ; GD.SetFocus ;
   END else
   BEGIN
   for i:=1 to GD.RowCount-1 do if GD.RowHeights[i]=0 then GD.RowHeights[i]:=GD.DefaultRowHeight ;
   for i:=1 to GC.RowCount-1 do if GC.RowHeights[i]=0 then GC.RowHeights[i]:=GC.DefaultRowHeight ;
   AvecTL:=True ; EnleveOuRajoute(True) ; GD.SetFocus ;
   END ;
 // LG* 10/01/2002
if FBoUneGrille then RemettreGrilleAJour(GD,FBoOter);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 21/03/2002
Modifié le ... :   /  /    
Description .. : LG* on ne pose plus la question en mode PCL
Mots clefs ... : 
*****************************************************************}
{procedure TFLettrage.VoirPiecesGenerees ;
BEGIN
if Not OkLettrage then Exit ;
if ctxPCL in V_PGI.PGIContexte then exit;
if TPIECES.Count<=0 then Exit ;
if HLettre.Execute(34,'','')<>mrYes then Exit ;
VisuPiecesGenere(TPieces,EcrGen,15) ;
END ;}

procedure TFLettrage.BAbandonClick(Sender: TObject);
begin
{OKOK}
Close ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 04/01/2002
Modifié le ... : 16/06/2004
Description .. : LG Modif :
Suite ........ : -Nouveaux message d'erreur
Suite ........ : 
Suite ........ : - LG - 12/04/2002 - gestion de l'etat des popup menu 
Suite ........ : validation lettrage te pre-lettre solution
Suite ........ : - LG -20/01/2004 - FB 13106 - en lettrage en saisie la 
Suite ........ : fenetre de la fenetre du choix du type de  validatnio fermait 
Suite ........ : la fenetre de lettrage
Suite ........ : - LG - 16/06/25004 - FB 10647 - qd on change d'auxi, on 
Suite ........ : ne ferme pas la fentre a la validation
Mots clefs ... : 
*****************************************************************}
procedure TFLettrage.ClickValide ;
Var io : TIOErr ;
BEGIN
//LG*
if Action=taConsult then Exit ;
if Not GD.SynEnabled then Exit ; GridEna(GD,False) ;
if Not GC.SynEnabled then Exit ; GridEna(GC,False) ;
if PieceModifiee and not ControleLigne then exit ;
GeneNow:=NowH ; RL.NowStamp:=GeneNow ;
//LG*
Validerlelettrage1.Enabled:=((ModeSelection) and (RL.Appel=tlMenu) and (OkLettrage));
Prlettrerlaslection1.Enabled:=ModeSelection and ((RL.Appel<>tlMenu) or (Not OkLettrage));
if ((ModeSelection) and (RL.Appel=tlMenu) and (OkLettrage)) then LettrageMemoire else
   BEGIN
   if ModeSelection and ((RL.Appel<>tlMenu) or (Not OkLettrage)) then
    begin
      if not LettrageMemoire then exit ;
    end ;
   if PieceModifiee then
      BEGIN
      GD.Enabled:=False ; GC.Enabled:=False ;
      io:=Transactions(LettrageFichier,1) ;
      GD.Enabled:=True ; GC.Enabled:=True ;
      if io=oeUnknown then MessageAlerte(HDivL.Mess[7]) else
       if io=oeLettrage then
        BEGIN
         if FBoUneGrille then MessageAlerte('Erreur lors de la mise à jour des informations de lettrage'+cDeuxLigne+FStLastError)
         else MessageAlerte(HDivL.Mess[8]) ;
        END
         else
          if io=oeSaisie then MessageAlerte('Erreur lors de l''enregistrement des écritures de régularisation'+cDeuxLigne+FStLastError)
           else
            if io=oeStock then MessageAlerte('Erreur lors de l''enregistrement des écritures de régularisation');
      END ;
// FQ 12785 (SBO) : Plus de visu des pièces générées, en PCL comme en PGE.
// if ((io=oeOk) and (OkLettrage)) then VoirPiecesGenerees ;
   ModeSelection:=False ; PieceModifiee:=False ; // Ne pas provoquer question sur close si erreur valide
   if not FBoAuxSuiv then Close ;
   END ;
EcartChange:=False ; PasEcart:=False ; AvecRegul:=False ;
GridEna(GD,True) ; GridEna(GC,True) ;
if FBoUneGrille then GD.Setfocus ;

// FQ 13325
GC.EndUpdate;
GD.EndUpdate;
END ;


procedure TFLettrage.ClickAnnulerSel ;
Var i : integer ;
BEGIN
{OKOK}
if Action=taConsult then Exit ;
if Not OkLettrage then Exit ;
if Not ModeSelection then Exit ;
for i:=1 to GD.RowCount-1 do if EstSelect(GD,i) then InverseSelection(GD,i) ;
for i:=1 to GC.RowCount-1 do if EstSelect(GC,i) then InverseSelection(GC,i) ;
GD.Invalidate ; GC.Invalidate ; CalculDebitCredit(False) ;
END ;

procedure TFLettrage.ClickChancel ;
BEGIN
if ((DEV.Code=V_PGI.DevisePivot) or (DEV.Code=V_PGI.DeviseFongible)) then Exit ;
FicheChancel(DEV.Code,True,V_PGI.DateEntree,taModif,(V_PGI.DateEntree>=V_PGI.DateDebutEuro)) ;
(*
if Not EstMonnaieIN(DEV.Code) then
   BEGIN
   FicheChancel(DEV.Code,True,V_PGI.DateEntree,taModif,(V_PGI.DateEntree>=V_PGI.DateDebutEuro)) ;
   END else
   BEGIN
   FicheDevise(DEV.Code,taModif,False) ;
   END ;
*)
DEV.Taux:=GetTaux(DEV.Code,DEV.DateTaux,LaDateGenere) ;
END ;

procedure TFLettrage.ClickCherche ;
BEGIN
{OKOK}
FindFirst:=True ; FindLettre.Execute ;
END ;

procedure TFLettrage.ClickCombi ( Unique : boolean ) ;
Var Solde  : double ;
BEGIN
{OKOK}
if Action=taConsult then Exit ;
if Not OkLettrage then Exit ;
if ((ModeSelection) and (Arrondi(LTE_TotDebCur-LTE_TotCredCur,CurDec)=0)) then BEGIN HLettre.Execute(14,'','') ; Exit ; END ;
if ModeSelection then
   BEGIN
   Solde:=LTE_TotDebCur-LTE_TotCredCur ;
   if Solde>0 then GridRappro:=2 else BEGIN GridRappro:=1 ; Solde:=-Solde ; END ;
   END else
   BEGIN
   HLettre.Execute(24,'','') ; Exit ;
   END ;
MaxNumSol:=0 ; EnvoieLaSauce(Solde,Unique) ;
END ;

procedure TFLettrage.BZoomCpteClick(Sender: TObject);
begin
{OKOK}
ClickZoom ;
end;

procedure TFLettrage.BZoomEcritureClick(Sender: TObject);
begin
{OKOK}
ClickZoomPiece ;
end;

procedure TFLettrage.BValideClick(Sender: TObject);
begin
{OKOK}
if Action=taConsult then Exit ;
ClickValide ;
end;

procedure TFLettrage.POPSPopup(Sender: TObject);
begin
{OKOK}
if Action=taConsult then Exit ;
InitPopUp(Self) ;
end;

procedure TFLettrage.BAnnulerSelClick(Sender: TObject);
begin
{OKOK}
if Action=taConsult then Exit ;
ClickAnnulerSel ;
end;

procedure TFLettrage.BEnleveClick(Sender: TObject);
begin
{OKOK}
if Action=taConsult then Exit ;
ClickEnleve(True) ;
end;

procedure TFLettrage.BRajouteClick(Sender: TObject);
begin
{OKOK}
if Action=taConsult then Exit ;
ClickEnleve(False) ;
end;

procedure TFLettrage.BChercherClick(Sender: TObject);
begin
{OKOK}
ClickCherche ;
end;

procedure TFLettrage.BChancelClick(Sender: TObject);
begin
{OKOK}
ClickChancel ;
end;

procedure TFLettrage.BCombiClick(Sender: TObject);
begin
{OKOK}
if Action=taConsult then Exit ;
ClickCombi(False) ;
end;

procedure TFLettrage.BUniqueClick(Sender: TObject);
begin
{OKOK}
if Action=taConsult then Exit ;
ClickCombi(True) ;
end;

procedure TFLettrage.BValideSelectClick(Sender: TObject);
begin
{OKOK}
if Action=taConsult then Exit ;
ClickValide ;
end;

{================================= SELECTIONS ==========================================}
function TFLettrage.EstLettre ( G : THGrid ; Lig : integer ; Totale : boolean ) : boolean ;
Var Col  : integer ;
    Code : String ;
BEGIN
{OKOK}
Result:=True ;
if G=GD then Col:=ColLetD else Col:=ColLetC ;
Code:=Trim(Copy(G.Cells[Col,Lig],1,4)) ;
if Code='' then Result:=False else if ((Totale) and (Code<>uppercase(Code))) then Result:=False ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 10/04/2002
Modifié le ... : 14/06/2002
Description .. : - LG - 09/04/2002 - 
Suite ........ : 
Suite ........ : - 14/06/2002 - bug su rle mode deux grilles
Mots clefs ... : 
*****************************************************************}
function TFLettrage.JePrends ( G : THGrid ; Lig : integer ) : boolean ;
var
 O : TOBM ;
BEGIN
{OKOK}
if OkLettrage then
 BEGIN
  if FBoUneGrille then
   BEGIN
    O:=GetO(G,Lig) ; if O=Nil then BEGIN JePrends:=EstSelect(G,Lig) ; exit ; END ;
    if (O.GetValue('SELECTIONNER')='X') then result:=true
    else JePrends:=EstSelect(G,Lig)
   END
    else JePrends:=EstSelect(G,Lig) ;
 END
 else JePrends:=Not EstSelect(G,Lig) ;
END ;

procedure TFLettrage.PrepareSelect ( G : THGrid ) ;
BEGIN
{OKOK}
G.ColWidths[G.ColCount-1]:=0 ;
END ;

procedure TFLettrage.Inverseselection ( G : THGrid ; Lig : integer ) ;
Var CL    : String ;
    i,Col : integer ;
    O : TOBM;
BEGIN
{OKOK}
if Action=taConsult then Exit ;
if G=GD then Col:=ColLetD else Col:=ColLetC ; CL:=G.Cells[Col,Lig] ;
if EstSelect(G,Lig) then
   BEGIN
   G.Cells[G.ColCount-1,Lig]:=' ' ; O:=GetO(G,Lig) ; O.PutMvt('SELECTIONNER','-') ; if G=GD then Dec(NbSelD) else Dec(NbSelC) ;
   if ((Not OkLettrage) and (DelettrageGlobal)) then
      BEGIN
      for i:=1 to GD.RowCount-1 do if GD.Cells[ColLetD,i]=CL then if ((GD<>G) or (i<>Lig)) then
          BEGIN GD.Cells[GD.ColCount-1,i]:=' ' ; O:=GetO(GD,i) ; O.PutMvt('SELECTIONNER','-') ; Dec(NbSelD) ; END ;
      for i:=1 to GC.RowCount-1 do if GC.Cells[ColLetC,i]=CL then if ((GC<>G) or (i<>Lig)) then
          BEGIN GC.Cells[GC.ColCount-1,i]:=' ' ; O:=GetO(GC,i) ; O.PutMvt('SELECTIONNER','-') ; Dec(NbSelC) ; END ;
      END else if ((OkLettrage) and (CL<>'')) then
      BEGIN
      for i:=1 to GD.RowCount-1 do if EstSelect(GD,i) then
          if GD.Cells[ColLetD,i]=CL then if ((GD<>G) or (i<>Lig)) then
             BEGIN
             GD.Cells[GD.ColCount-1,i]:=' ' ; O:=GetO(GD,i) ; O.PutMvt('SELECTIONNER','-') ; Dec(NbSelD) ;
             END ;
      for i:=1 to GC.RowCount-1 do if EstSelect(GC,i) then
          if GC.Cells[ColLetC,i]=CL then if ((GC<>G) or (i<>Lig)) then
             BEGIN
             GC.Cells[GC.ColCount-1,i]:=' ' ; O:=GetO(GC,i) ; O.PutMvt('SELECTIONNER','-') ; Dec(NbSelC) ;
             END ;
      END ;
   END else
   BEGIN
   G.Cells[G.ColCount-1,Lig]:='+' ; O:=GetO(G,Lig) ; O.PutMvt('SELECTIONNER','X') ; if G=GD then Inc(NbSelD) else Inc(NbSelC) ;
   if CL<>'' then if ((OkLettrage) or ((Not OkLettrage) and (DelettrageGlobal))) then
      BEGIN
      for i:=1 to GD.RowCount-1 do
       if GD.Cells[ColLetD,i]=CL then
        if ((GD<>G) or (i<>Lig)) then
          BEGIN
           GD.Cells[GD.ColCount-1,i]:='+' ; O:=GetO(GD,i) ; O.PutMvt('SELECTIONNER','X') ;
           Inc(NbSelD) ;
           END ;
      for i:=1 to GC.RowCount-1 do if GC.Cells[ColLetC,i]=CL then if ((GC<>G) or (i<>Lig)) then
          BEGIN GC.Cells[GC.ColCount-1,i]:='+' ; O:=GetO(GC,i) ; O.PutMvt('SELECTIONNER','X') ; Inc(NbSelC) ; END ;
      END ;
   END ;
GD.Invalidate ; GC.Invalidate ;
ModeSelection:=((NbSelD>0) or (NbSelC>0)) ; GereLesValide ;
CParCouverture.Enabled:=Not ModeSelection ; if ((LienS1S3) or (EstSerie(S3))) then CParCouverture.Enabled:=False ;
CParEnsemble.Enabled:=Not ModeSelection ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 05/03/2002
Modifié le ... : 01/02/2006
Description .. : Ajout d'un parametre Parle : si false, la fenetre posant la
Suite ........ : question 'voulez vous lettre...' n'apparait pas
Suite ........ : 
Suite ........ : - 03/04/2002 modif pour permettre la selection de ligne
Suite ........ : avec ecart de regul... avec la barre d'espace
Suite ........ : Ajout d'un nouveau param LigneAvecRegul : si true autorise
Suite ........ : la selection des lignes avec ecart de regul
Suite ........ : - 01/02/2005 - LG - FB 17415 - qd on selectionne une ligne 
Suite ........ : on verifie systematiquement en base
Mots clefs ... : 
*****************************************************************}
procedure TFLettrage.Selection ( Next : boolean ; Parle : boolean=true ; LigneAvecRegul : boolean=false) ;
Var G : THGrid ;
    O : TOBM ;
    lTOBListeBor : TOB ;
BEGIN
if Action=taConsult then Exit ;
if GD.Focused then BEGIN G:=GD ; CurGrid:=1 ; END else if GC.Focused then BEGIN G:=GC ; CurGrid:=2 ; END else Exit ;
if G.Row<=0 then Exit ; O:=GetO(G,G.Row) ; if O=Nil then Exit ;
// on rafraichissit la liste des bordereaux bloques
lTOBListeBor:=CGetListeBordereauBloque ;
if CEstBloqueEcritureBor(O,lTOBListeBor) then
 begin
  O.PutMvt('E_CONTROLEUR','X') ;
  lTOBListeBor.Free ;
  exit ;
 end
  else
   begin
    O.PutMvt('E_CONTROLEUR','-') ;
    lTOBListeBor.Free ;
   end ;
//LG* 03/04/2002 modif pour permettre la selection de ligne avec ecart de regul... avec la barre d'espace
if (not LigneAvecRegul) and ((OkLettrage) and (PasTouche(O.GetMvt('E_ETAT')))) then
 BEGIN if Parle then HLettre.Execute(8,'','') ; Exit ; END ;
// Depuis saisie tréso, ne pas pouvoir déselectionner la ligne d'où l'on vient
if ((RL.Appel=tlSaisieTreso) and (GSais=G) and (LigSais>0) and
    (G.Row=LigSais) and (EstSelect(G,G.Row))) then Exit ;
InverseSelection(G,G.Row) ;
//LG*
if FBoUneGrille then CalculDebitCredit(Parle) else CalculDebitCredit(Not ModeCouverture) ;
if ((EstSelect(G,G.Row)) and (ModeCouverture)) then BEGIN CurLig:=G.Row ; SaisieCouverture(O) ; END ;
if ((Next) and (G.Row<G.RowCount-1)) then G.Row:=G.Row+1 ;
END ;

{================================ AFFICHAGES ===========================================}
{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 29/10/2002
Modifié le ... :   /  /    
Description .. : -29/10/2002- les menus de lettrage etaient toujours grisés
Mots clefs ... : 
*****************************************************************}
procedure TFLettrage.GereLesValide ;
BEGIN
{OKOK}
if ModeSelection then
   BEGIN
   BValideSelect.Enabled:=True ;
   GeneValide.Glyph:=BValideSelect.Glyph ; GeneValide.NumGlyphs:=2 ;
   GeneValide.Hint:=BValideSelect.Hint ;
   BValide.Enabled:=False ;
   END else
   BEGIN
   BValide.Enabled:=True ;
   GeneValide.Glyph:=BValide.Glyph ; GeneValide.NumGlyphs:=2 ;
   GeneValide.Hint:=BValide.Hint ;
   BValideSelect.Enabled:=False ;
   END ;
Lettragepartiel2.Enabled:=BValideSelect.Enabled ; Passagedunecrituresimplifie1.Enabled:=BValideSelect.Enabled ;
Soldeautomatique2.Enabled:=BValideSelect.Enabled ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 26/12/2001
Modifié le ... :   /  /
Description .. : LG modif :
Suite ........ : -gestion des nouveaux controles pour le mode pcl
Mots clefs ... :
*****************************************************************}
procedure TFLettrage.DetailLigne ( Lig : integer ; Debit : boolean ) ;
Var i  : integer ;
    O  : TOBM ;
    VV : Variant ;
    C  : TControl ;
    XX,YY : double ;
    Nom,Jal   : String ;
BEGIN
{OKOK}
// on ne met pas a jour la grille des credit pour le mode pcl
if not Debit and FBoUneGrille then exit;
if Debit then O:=GetO(GD,Lig) else O:=GetO(GC,Lig) ;
if O = nil then exit ;
for i:=0 to PDetail.ControlCount-1 do
    BEGIN
    C:=PDetail.Controls[i] ; if C.Tag=0 then Continue ;
    if ((Debit) and (C.Name[Length(C.Name)]='_')) then Continue ;
    if ((Not Debit) and (C.Name[Length(C.Name)]<>'_')) then Continue ;
    if O=Nil then
       BEGIN
       if C is THNumEdit then
          BEGIN
          THNumEdit(C).Value:=0 ;
          END else if C is TLabel then TLabel(C).Caption:='' ;
       END else
       BEGIN
       Nom:=C.Name ; if Not Debit then Nom:=Copy(Nom,1,Length(Nom)-1) ;
       if ((Nom<>'E_TOTAL') and (Nom<>'E_TOTALEURO')) then VV:=O.GetMvt(Nom) ;
       if Nom='E_JOURNAL' then BEGIN CJOURNAL.Value:=VV ; TLabel(C).Caption:=CJOURNAL.Text ; END else
       if Nom='E_NATUREPIECE' then BEGIN CNATUREPIECE.Value:=VV ; TLabel(C).Caption:=CNATUREPIECE.Text ; END else
       if Nom='E_LIBELLE' then TLabel(C).Caption:=VarAsType(VV,VarString) else
       if Nom='E_DATEECHEANCE' then TLabel(C).Caption:=DateToStr(VarAsType(VV,VarDate)) else
       if Nom='E_MODEPAIE' then
          BEGIN
          if BTF.Down then
             BEGIN
             Jal:=O.GetMvt('E_JOURNAL') ;
             if ((Jal=VH^.JalVTP) or (Jal=VH^.JalATP)) then TLabel(C).Caption:=O.GetMvt('E_CONTREPARTIEAUX')
                                                       else TLabel(C).Caption:='' ;
             END else
             BEGIN
             CMODEPAIE.Value:=VV ; TLabel(C).Caption:=CMODEPAIE.Text ;
             END ;
          END else
       if Nom='E_REFRELEVE' then TLabel(C).Caption:=VarAsType(VV,VarString) else
       if Nom='E_TOTAL' then
          BEGIN
          if Debit then BEGIN XX:=O.GetMvt('E_DEBIT') ; YY:=0 END else BEGIN XX:=0 ; YY:=O.GetMvt('E_CREDIT') ; END ;
          AfficheLeSolde(THNumEdit(C),XX,YY) ;
          END else
            TLabel(C).Caption:=VarAsType(VV,VarString) ;
       END ;
    END ;
//LG* 26/12/2001 gestion des nouveaux controles
if FBoUneGrille and (O<>nil) then
 BEGIN
  e_naturepcl.Caption:=RechDom('ttNaturePiece', O.GetMvt('E_NATUREPIECE') ,false) ; e_dateecheancepcl.Caption:=DateToStr(O.GetMvt('E_DATEECHEANCE')) ;
  e_refrelevepcl.Caption:=O.GetMvt('E_REFRELEVE') ; e_contrepartiegenpcl.Caption:=O.GetMvt('E_CONTREPARTIEGEN') ;
  e_etablissementpcl.Caption:=RechDom('ttEtablissement', O.GetMvt('E_ETABLISSEMENT') ,false) ; e_reflibrepcl.Caption:=O.GetMvt('E_REFLIBRE') ;
  e_refexternepcl.Caption:=O.GetMvt('E_REFEXTERNE') ; e_affairepcl.Caption:=O.GetMvt('E_AFFAIRE') ; e_dateecheancepcl.Caption:=O.GetMvt('E_DATEECHEANCE') ;
  if BTF.Down then
   BEGIN
    Jal:=O.GetMvt('E_JOURNAL') ;
    if ((Jal=VH^.JalVTP) or (Jal=VH^.JalATP)) then e_modepaiepcl.Caption:=O.GetMvt('E_CONTREPARTIEAUX') else e_modepaiepcl.Caption:='' ;
   END
    else
     BEGIN
      CMODEPAIE.Value:=O.GetMvt('E_MODEPAIE') ; e_modepaiepcl.Caption:=CMODEPAIE.Text ;
     END ;
 END;
 if O=nil then exit ;
 if O.GetValue('E_CONTROLEUR')='X' then
  begin
   caption:=AttribTitre + ' : ligne en cours de modification dans le folio n°' + intToStr(O.GetValue('E_NUMEROPIECE')) + ' du ' + dateToStr(O.GetValue('E_DATECOMPTABLE')) ;
  end
   else
    caption:=AttribTitre ; ;
END ;

procedure TFLettrage.ActiveGrid ;
BEGIN
{OKOK}
if GD.Focused then BEGIN CurGrid:=1 ; GD.Font.Color:=clWindowText ; GC.Font.Color:=clGray ; END ;
if GC.Focused then BEGIN CurGrid:=2 ; GC.Font.Color:=clWindowText ; GD.Font.Color:=clGray ; END ;
END ;

procedure TFLettrage.SwapPanels ;
BEGIN
{OKOK}
PDetail.Visible:=Not PDetail.Visible ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 22/11/2001
Modifié le ... : 24/12/2001
Description .. : LG Modif :
Suite ........ : -redim de la grille GD
Suite ........ : -repositionnement des controles soldes
Mots clefs ... : 
*****************************************************************}
procedure TFLettrage.CentreGrids ;
var
  lInPosCol,i : integer;
BEGIN
{OKOK}
//LG*
if FBoUneGrille then
 BEGIN
  if FBoDessiner then begin HMTrad.ResizeGridColumns(GD); FBoDessiner := false ; end ;
  lInPosCol:=0 ;
  for i:=0 to ColDebit-1 do lInPosCol:=lInPosCol+GD.ColWidths[i];
  SL_CUMULDEBIT.Left:=lInPosCol ; SL_CUMULDEBIT.Width:=GD.ColWidths[ColDebit] ;
  SL_SOLDEDEBITPCL.Left:=lInPosCol ; SL_SOLDEDEBITPCL.Width:=SL_CUMULDEBIT.Width ;
  SL_SOLDEDEBITPCL.Top:=SL_CUMULDEBIT.Top+SL_CUMULDEBIT.Height+2 ;
  h1.Left:=SL_CUMULDEBIT.Left-h1.Width-3 ; h1.top:=SL_CUMULDEBIT.top+2 ;
  lInPosCol:=0 ;
  for i:=0 to ColCredit-1 do lInPosCol:=lInPosCol+GD.ColWidths[i];
  SL_CUMULCREDIT.Left:=lInPosCol ; SL_CUMULCREDIT.Width:=GD.ColWidths[ColDebit] ;
  SL_SOLDECREDITPCL.Left:=lInPosCol ; SL_SOLDECREDITPCL.Width:=SL_CUMULCREDIT.Width ;
  SL_SOLDECREDITPCL.Top:=SL_CUMULCREDIT.Top+SL_CUMULCREDIT.Height+2 ;
  SL_SOLDE.Left:=SL_SOLDECREDITPCL.Left+SL_SOLDECREDITPCL.Width+2 ;  SL_SOLDE.Top:=SL_SOLDECREDITPCL.Top ; SL_SOLDE.Width:=SL_SOLDECREDITPCL.Width ;
  h2.Left:=SL_CUMULDEBIT.Left-h2.Width-3  ; h2.top:=SL_SOLDECREDITPCL.Top+2 ;
  SL_SOLDEPCL.Top:=SL_CUMULCREDIT.Top ; SL_SOLDEPCL.Left:=SL_SOLDE.Left ; SL_SOLDEPCL.Width:=SL_SOLDE.Width ;
  exit;
 END;
GD.Align:=alNone ; GC.Align:=alNone ;
GD.Width:=(Pages.Width-10) Div 2+1 ; GC.Width:=GD.Width;
GC.Align:=alRight ; GD.Align:=alClient ;
BSepare.Height:=PDetail.Height ; BSepare.Left:=10 + PDetail.Width div 2 ;
PCouverture.Left:=(Width-PCouverture.Width) div 2 ;
PCouverture.Top:=GD.Top+(GD.Height-PCouverture.Height) div 2 ;
if not FBoUneGrille then
begin
 HMTrad.ResizeGridColumns(GC) ; HMTrad.ResizeGridColumns(GD) ;
 if CurGrid=1 then BEGIN GD.Font.Color:=clWindowText ; GC.Font.Color:=clGray ; END ;
 if CurGrid=2 then BEGIN GC.Font.Color:=clWindowText ; GD.Font.Color:=clGray ; END ;
end ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 14/01/2002
Modifié le ... :   /  /
Description .. : LG Modif :
Suite ........ : - inutilise en mode pcl
Mots clefs ... : 
*****************************************************************}
procedure TFLettrage.SwapGrids ;
BEGIN
{OKOK}
//LG*
if FBoUneGrille then exit;
if GD.Focused then GC.SetFocus else if GC.Focused then GD.SetFocus ;
END ;

{============================ LETTRAGE PAR COUVERTURE ==================================}
procedure TFLettrage.BCAbandonClick(Sender: TObject);
begin
{OKOK}
CacheCouverture ;
end;

procedure TFLettrage.BCValideClick(Sender: TObject);
Var O : TOBM ;
    G : THGrid ;
    X,Y : Double ;
begin
X:=E_COUVERTURES.Value ;
if X<0 then BEGIN HLettre.Execute(9,'','') ; Exit ; END ;
if E_RESTES.Value<0 then BEGIN HLettre.Execute(10,'','') ; Exit ; END ;
if CurGrid=1 then G:=GD else if CurGrid=2 then G:=GC else Exit ;
O:=GetO(G,CurLig) ;
if RL.LettrageDevise then
   BEGIN
   O.PutMvt('E_COUVERTUREDEV',X) ;
   Y:=DeviseToEuro(X,E_TAUXDEV.Value,DEV.Quotite)  ; O.PutMvt('E_COUVERTURE',Y) ;
   END else
   BEGIN
   O.PutMvt('E_COUVERTURE',X) ;
   ConvertCouverture(O,tsmPivot) ;
   END ;
CacheCouverture ; CalculDebitCredit(True) ;
end;

procedure TFLettrage.E_COUVERTURESChange(Sender: TObject);
Var C : Double ;
begin
{OKOK}
C:=Valeur(E_COUVERTURES.Text) ; E_RESTES.Value:=Abs(E_MONTANTS.Value)-C ;
end;

procedure TFLettrage.CacheCouverture ;
BEGIN
{OKOK}
Pages.Enabled:=True ; POutils.Enabled:=True ;  Valide97.Enabled:=True ;
GD.Enabled:=True ; GC.Enabled:=True ;
if CurGrid=1 then GD.SetFocus else GC.SetFocus ;
SetCaptureControl(Nil) ; PCouverture.Visible:=False ;
BCValide.Visible:=False ; BCAbandon.Visible:=False ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 27/12/2001
Modifié le ... :   /  /    
Description .. : Suppression des warnings
Mots clefs ... : 
*****************************************************************}
procedure TFLettrage.SaisieCouverture ( O : TOBM ) ;
Var Debit,Credit,Delta : Double ;
BEGIN
if Action=taConsult then Exit ;
Debit:=GetDebitCur(O) ; Credit:=GetCreditCur(O) ; Delta:=GetCouvCur(O) ;
if CurGrid=1 then LTE_CouvDebCur:=LTE_CouvDebCur-Delta
             else LTE_CouvCredCur:=LTE_CouvCredCur-Delta ;
E_TOTCOUVDebS.Value:=LTE_CouvDebCur ; E_TOTCOUVCredS.Value:=LTE_CouvCredCur ;
E_TAUXDEV.Visible:=False ; E_TAUXDEV.Value:=O.GetMvt('E_TAUXDEV') ;
{Affectation des masques et formats}
ChangeMask(E_TOTCOUVDebS,CurDec,DEV.Symbole) ; ChangeMask(E_TOTCOUVCredS,CurDec,DEV.Symbole) ;
ChangeMask(E_MONTANTS,CurDec,DEV.Symbole) ; ChangeMask(E_COUVERTURES,CurDec,DEV.Symbole) ;
ChangeMask(E_RESTES,CurDec,DEV.Symbole) ;
{Affichages}
AfficheLeSolde(E_MONTANTS,Debit,Credit) ; E_COUVERTURES.Value:=Abs(E_MONTANTS.Value) ;
//CentreGrids ;
PCouverture.Visible:=True ;
BCValide.Visible:=True ; BCAbandon.Visible:=True ;
SetCaptureControl(Pages) ;
E_COUVERTURES.SetFocus ; E_COUVERTURES.SelectAll ;
Pages.Enabled:=False ; POutils.Enabled:=False ; Valide97.Enabled:=False ;
GD.Enabled:=False ; GC.Enabled:=False ;
END ;

{============================= METHODES DE LA FORM =====================================}
procedure TFLettrage.FindLettreFind(Sender: TObject);
begin
{OKOK}
if CurGrid=1 then Rechercher(GD,FindLettre,FindFirst) else Rechercher(GC,FindLettre,FindFirst) ;
end;

procedure TFLettrage.GereEnabled ;
BEGIN
if Action=taConsult then
   BEGIN
   BValide.Enabled:=False        ; BAnnulerSel.Enabled:=False ;
   BRajoute.Enabled:=False       ; BEnleve.Enabled:=False ; BAjouteEnleve.Enabled:=False ;
   BCombi.Enabled:=False         ; BUnique.Enabled:=False ;
   CParCouverture.Enabled:=False ; CParEnsemble.Enabled:=False ; CParExcept.Enabled:=False ; CParExcept.Checked:=False ;
   E_REGUL.Enabled:=False        ; E_ECART.Enabled:=False ; DateGeneration.Enabled:=False ; 
   END else
   BEGIN
   if OkLettrage then
      BEGIN
      E_REGUL.Enabled:=True ;
//      if LienS1S3 then E_REGUL.Enabled:=False ; // 14350
      END else
      BEGIN
      if RL.Appel<>tlMenu then BEGIN E_REGUL.Enabled:=False ; E_ECART.Enabled:=False ; DateGeneration.Enabled:=False ; END ;
      END ;
   END ;
if Not OkLettrage then
   BEGIN
   BAnnulerSel.Enabled:=False ; BCombi.Enabled:=False ; BUnique.Enabled:=False ;
   END ;
if ((LaDateGenere<VH^.Encours.Deb) {or (LienS1S3)}) then // 14350
   BEGIN
   E_REGUL.Enabled:=False ; E_ECART.Enabled:=False ; DateGeneration.Enabled:=False ;
   END ;
if Not VH^.OuiTP then BTF.Visible:=False ;
If (RL.Appel=tlSaisieTreso) Then BAnnulerSel.Enabled:=FALSE ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 13/04/2006
Modifié le ... :   /  /    
Description .. : - 13/04/2006 - LG - deplacement du initlettrage du 
Suite ........ : formCreate vers le formshow ( le FboFromTOB etait tout le 
Suite ........ : temps faux ds le formCreate )
Mots clefs ... : 
*****************************************************************}
procedure TFLettrage.FormShow(Sender: TObject);
begin
{OKOK}
InitLettrage ;
if FBoFermer then begin modalresult:=mrOk ; close; exit ; end ;
if OkLettrage then HelpContext:=7508100 else HelpContext:=7511100 ;
if RL.Appel=tlSaisieCour then HelpContext:=7244400 ;
LaDateGenere:=V_PGI.DateEntree ;
if VH^.CPExoRef.Code<>'' then
  LaDateGenere:=VH^.CPExoRef.Fin ;
DateGeneration.Text:=DateToStr(LaDateGenere) ;
FBoOter := GetParamSocSecur('SO_LETPREFTOT',false) and not FBoFromTOB ;
ShowLettrage(True) ;
GereEnabled ;

end;                                          

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 20/11/2001
Modifié le ... : 26/04/2004
Description .. : LG Modif :
Suite ........ : -Ajout de nouveau evenement pour grisé les colonnes
Suite ........ : debit/credit et colorisé la case solde en fonction du solde
Suite ........ : -Chargement des info du lettrage
Suite ........ : -Création des listes servant à l'annulation et de la tob de 
Suite ........ : gestion des ecr. de regul
Suite ........ : -Gestion du choix du lettrage à l'ouverture
Suite ........ : 
Suite ........ : - FB 10144 - cache les fonctionnalite ne servant pas en
Suite ........ : delettrage
Suite ........ : 
Suite ........ : - 12/02/2003 - deplacement du code dasn InitLettrage
Suite ........ : 
Suite ........ : - 26/04/2004 : Init popmenu PCL placé dans form create
Mots clefs ... : 
*****************************************************************}
procedure TFLettrage.FormCreate(Sender: TObject);
begin
 FBoChangementMode:=false ; FBoDessiner := true ;

// SBO 26/04/2004 : Correction menu pop, l'initialisation doti se faire une seule fois, reporté dans le FormCreate
  // Initialisation du popup PCL
  AddMenuPop(POPPCL,'','') ;
  // affectation des raccourcis par code car des vrais raccourcis ( proprietes shortcut du pop ) ne fct pas avec deux valeurs
  // identiques.
  Validerlelettrage1.Caption             := Validerlelettrage1.Caption+#9+ShortCutToText(ShortCut(VK_F10,[])) ;
  Prlettrerlaslection1.Caption           := Prlettrerlaslection1.Caption+#9+ShortCutToText(ShortCut(VK_F10,[])) ;
  Enleverlestotalementslettrs1.Caption   := Enleverlestotalementslettrs1.Caption+#9+ShortCutToText(ShortCut(VK_DELETE,[ssShift])) ;
  Remettrelestotalementslettres1.Caption := Remettrelestotalementslettres1.Caption+#9+ShortCutToText(ShortCut(VK_DELETE,[ssShift])) ;
  Voirlcriture1.Caption                  := Voirlcriture1.Caption+#9+ShortCutToText(ShortCut(VK_F5,[])) ;
  Voirlecompte1.Caption                  := Voirlecompte1.Caption+#9+ShortCutToText(ShortCut(VK_F5,[ssShift])) ;
  Tabledechancellerie1.Caption           := Tabledechancellerie1.Caption+#9+ShortCutToText(ShortCut(ord('H'),[ssAlt])) ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 20/11/2001
Modifié le ... : 29/11/2001
Description .. : LG Modif :Affiche le solde progressif en vert si le solde
Suite ........ : progressif est de meme sens que le compte et en rouge
Suite ........ : dans le cas contraire
Mots clefs ... :
*****************************************************************}
procedure TFLettrage.GetCellCanvasD(ACol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
BEGIN
{OKOK}
//LG*
  if ARow <= 0 then
    Exit;
  if EstSelect ( GD, ARow ) then
    GD.Canvas.Font.Style := GD.Canvas.Font.Style + [fsItalic]
  else
  begin
    GD.Canvas.Font.Style := GD.Canvas.Font.Style - [fsItalic];
    if FBoUneGrille and ( ACol = ( ColSolde ) ) and ( ARow <> GD.Row ) then
    begin
      if FStSensCompte = 'D' then
        Canvas.Font.Color := IIF( Pos('D', GD.Cells[ACol, Arow]) > 0, ClGreen, ClRed)
      else
        if FStSensCompte = 'C' then
          Canvas.Font.Color := IIF( Pos('C', GD.Cells[ACol, Arow]) > 0, ClGreen, ClRed);
    end;
  end;
END ;

procedure TFLettrage.GetCellCanvasC(ACol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
BEGIN
{OKOK}
if ARow<=0 then Exit ;
if EstSelect(GC,ARow) then GC.Canvas.Font.Style:=GC.Canvas.Font.Style+[fsItalic]
                      else GC.Canvas.Font.Style:=GC.Canvas.Font.Style-[fsItalic] ;
END ;

procedure TFLettrage.FormResize(Sender: TObject);
begin
{OKOK}
CentreGrids ; // FQ13904 SBO 04/08/2003 : pb d'affichage en mode débit/crédit sur resize de la fenêtre
end;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 27/11/2001
Modifié le ... : 02/07/2004
Description .. : Ajout des nouveaux raccourcis
Suite ........ : 
Suite ........ : - 10/06/2002 - ajout d'un effet interrupteur sur la touche de 
Suite ........ : raccourci alt+z
Suite ........ : -17/12/2002 - suprression du alt+Z integrre dans l'agl
Suite ........ : - 02/07/2004 - FB 10647 - LG - Ajout des raccourci F3 et 
Suite ........ : F4
Mots clefs ... : 
*****************************************************************}
procedure TFLettrage.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var Vide,OkG,EnCouv : boolean ;
begin
if DATEGENERATION.Focused then Exit ;
OkG:=((GD.Focused) or (GC.Focused)) ; Vide:=(Shift=[]) ; EnCouv:=(Not POutils.Enabled) ;
if ((EnCouv) and (Key<>VK_ESCAPE) and (Key<>VK_RETURN) and (Key<>VK_F10)) then Exit ;
Case Key of
   VK_TAB           : if OKG then SwapGrids ;
   VK_ESCAPE        : if Vide then Close ;
   VK_INSERT        : if Shift=[ssShift] then ClickEnleve(False) ;
   //LG* 20/12/2001 ajout d'une fonction de delettrage
   VK_DELETE        : if Shift=[ssShift] then ClickEnleve(not FBoOter) else if Shift=[ssCtrl] then DeLettPartielClick(nil);
   VK_RIGHT,VK_LEFT : BEGIN Key:=0 ; SwapGrids ; END ;
   VK_RETURN        : if EnCouv then BEGIN NextControl(Self) ; BCValideClick(Nil) ; END else
                         BEGIN
                         if TraiteReturn then Key:=0 else BEGIN TraiteReturn:=True ; Selection(True) ; TraiteReturn:=False ; END ;
                         END ;
   VK_SPACE         : if ((Not EnCouv) and (Vide) and (OkG)) then Selection(False,true,true) ;
   VK_F3            : if Vide then
                       BUPClick(nil)
                      else
                       BEGIN Key:=0 ; if CParExcept.Enabled then CParExcept.Checked:=True ; END ;
   VK_F4            : if Vide then
                       BDownClick(nil) ;

   VK_F5            : BEGIN
                       if Vide then BEGIN Key:=0 ; ClickZoomPiece ; END ;
                       if Shift=[ssShift] then BEGIN Key:=0 ; ClickZoom ; END ;
                      END ;
//   VK_F8            : BEGIN Key:=0 ; DelettrageCourant ; END ;
   VK_F10           : BEGIN
                      Key:=0 ;
                      if EnCouv then
                      BEGIN NextControl(Self) ; BCValideClick(Nil) ; END
                      else if Vide then ClickValide ;
                      END ;
   VK_F6            : SwapPanels ;
   VK_F11           : BEGIN
                       Key:=0 ;
                       if FBoUneGrille then // le menu popup est diffrent en mode pcl
                        POPPCL.PopUp(Mouse.CursorPos.X,Mouse.CursorPos.Y)
                       else
                        PopS.PopUp(Mouse.CursorPos.X,Mouse.CursorPos.Y) ;
                      END ;
   VK_F12           : begin
                       Key := 0 ;
                       if Pages.Focused then
                        GD.SetFocus
                         else
                          Pages.SetFocus ;
                      end ;
{AC}             67 : if (Shift=[ssAlt]) and (not FBoUneGrille) then BEGIN Key:=0 ; ClickCombi(False) ; END ;
{AH}             72 : if Shift=[ssAlt] then BEGIN Key:=0 ; ClickChancel ; END ;
{^Z}             90 : if Shift=[ssCtrl] then BEGIN Key:=0 ; ClickAnnulerSel ; END ;
(*                      else if Shift=[ssAlt] then
                       BEGIN Key:=0 ; if WindowState=wsMaximized then WindowState:=wsNormal else WindowState:=wsMaximized END ; //LG* Ajout d'un zoom sur la fenêtre*)
//LG* nouveaux raccourci
{Alt+A}     Ord('A'): if (BCalcul.Visible) and (Shift=[ssCtrl]) then BEGIN Key:=0 ; ClickToutSel ; END ;
{Alt+1}     Ord('1'): if (BCalcul.Visible) and (Shift=[ssAlt]) then BEGIN Key:=0 ; Calcul(1)  ; END ;  // lettrage sur passage de solde
{Alt+2}     Ord('2'): if (BCalcul.Visible) and (Shift=[ssAlt]) then BEGIN Key:=0 ; ClickCombi(False) ; END ;
{Alt+3}     Ord('3'): if (BCalcul.Visible) and (Shift=[ssAlt]) then BEGIN Key:=0 ; ClickCombi(True) ; END ;
{Alt+4}     Ord('4'): if (BCalcul.Visible) and (Shift=[ssAlt]) then BEGIN Key:=0 ; Calcul(0) ; END ; // lettrage sur refrence
{Alt+5}     Ord('5'): if (BCalcul.Visible) and (Shift=[ssAlt]) then BEGIN Key:=0 ; Calcul(2) ;  END ;  // lettrage sur passage de solde à nulle
{Alt+M}     Ord('M'): if FBoUneGrille and Vide then BEGIN Key:=0 ; PasseMoisSuivant ; END
                      else if (Shift=[ssCtrl]) then BEGIN Key:=0 ; ClickToutSel(false,false,true) ; END ;
{Alt+J}     Ord('J'): if FBoUneGrille and Vide then BEGIN Key:=0 ; PasseJourSuivant ; END
                      else if (Shift=[ssCtrl]) then BEGIN Key:=0 ; ClickToutSel(false,true) ; END ;
{Alt+L}     Ord('L'): if FBoUneGrille and Vide then BEGIN Key:=0 ; PasseLibelleSuivant ; END ;
            // GCO - 30/05/2006 - FQ 18187
{Alt+F}     Ord('F'): if (Shift=[ssCtrl]) then
                      begin
                        Key:=0 ;
                        ClickCherche ;
                      end;
{ALT+C}    // 67 : if Shift=[ssAlt] then begin Key:=0 ; ComplClick ; end ;
{CTRL+D}    Ord('D'): if (Shift=[ssCtrl]) then
                          begin Key:=0 ; DelettrageCourant ; END ;
   END ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 28/06/2006
Modifié le ... :   /  /    
Description .. : - LG - 28/06/2006 - FB 13402 - remplacment apr un PGIAsk 
Suite ........ : pour gerer correctemetn le echap sur la fenetre de 
Suite ........ : confirmation
Mots clefs ... : 
*****************************************************************}
procedure TFLettrage.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
Var ii : integer ;
begin
{mode selection}
//LG*
if FBoFromTOB and FBoFermer then
 BEGIN
 CanClose:=true ;
  exit;
  END;
if ModeSelection then
   BEGIN
   if OkLettrage then ii:=HLettre.Execute(4,'','') else ii:=HLettre.Execute(23,'','') ;
   if ii<>mrYes then CanClose:=False else BEGIN if OkLettrage then DevalideCode ; END ;
   Exit ;
   END ;
{mode déplacement}
if PieceModifiee then
   BEGIN
   if RL.Appel<>tlMenu then ii:=mrYes else
      BEGIN
      if OkLettrage then
       ii := PGIAskCancel('Vous abandonnez un traitement de lettrage en cours. Voulez-vous valider ce lettrage ?')
        else
          ii := PGIAskCancel('Vous abandonnez un traitement de délettrage en cours. Voulez-vous valider ce délettrage ?') ;
      END ;
   Case ii of
      mrYes    : ClickValide ;
      mrCancel : CanClose:=False ;
      mrNo     : if OkLettrage then DevalideCode ;
      END ;
   END ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 10/12/2001
Modifié le ... : 10/03/2006
Description .. : LG Modif :
Suite ........ : -Suppression de TList d'annulation
Suite ........ : -Ajout de la tob stockant toutes les TOBM
Suite ........ : 
Suite ........ : LG - 15/07/2002 - reecriture du code, la fenetre se ferme 
Suite ........ : malgres les violations d'acces
Suite ........ : 
Suite ........ : - 13/02/2003 - suppresion de FTOBListeBor, liste des 
Suite ........ : bordereaux bloques
Suite ........ : - 10/03/2006 - suprression d'une fuite memoire
Mots clefs ... : 
*****************************************************************}
procedure TFLettrage.FormClose(Sender: TObject; var Action: TCloseAction);
begin
{OKOK}
//LG*
try
 if FStAnnulLettrage<>nil then FStAnnulLettrage.Free ;
 if not FBoFromTOB  and (FTOBEcr<>nil) then FTOBEcr.Free ;
 if FInfoEcr<>nil then FInfoEcr.Free ;
 if FLCAnnul<>nil then BEGIN VideListe(FLDAnnul) ; FLDAnnul.Free; END ;
 if FLCAnnul<>nil then BEGIN VideListe(FLCAnnul) ; FLCAnnul.Free; END ;
 if TLAD<>nil then BEGIN VideListe(TLAD)    ; TLAD.Free    ; END ;
 if TLAC<>nil then BEGIN VideListe(TLAC)    ; TLAC.Free    ; END ;
 if TTSOLU<>nil then BEGIN VideListe(TTSOLU)  ; TTSOLU.Free  ; END ;
 if FEcrRegul<>nil then FEcrRegul.Free ;
 PurgePopup(POPS) ;
 RegSaveToolbarPos(Self,'Lettrage') ;
finally
 FInfoEcr:=nil ; if not FBoFromTOB  then FTOBEcr:=nil ; FLDAnnul:=nil ; FLCAnnul:=nil ; FStAnnulLettrage := nil ;
 TLAD:=Nil ;  TLAC:=Nil ; TTSOLU:=Nil ;
end;
end;

{============================= METHODES DE LES GRIDS ===================================}
procedure TFLettrage.MontreDetail(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
Var G : THGrid ;
begin
{OKOK}
G:=THGrid(Sender) ; DetailLigne(Ou,G.Name='GD') ;
end;

procedure TFLettrage.ClickSelect(Sender: TObject);
begin
{OKOK}
if Action=taConsult then Exit ;
Selection(True) ;
end;

procedure TFLettrage.ClickActive(Sender: TObject);
begin
{OKOK}
ActiveGrid ;
end;

procedure TFLettrage.GDMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
{OKOK}
if ((ssCtrl in Shift) and (Button=mbLeft)) then Selection(False) ;
end;

procedure TFLettrage.SauteZero(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
Var G : THGrid ;
begin
{OKOK}
G:=THGrid(Sender) ;
if G.RowHeights[G.Row]=0 then
   BEGIN
   if ((ARow<G.Row) and (G.Row<G.RowCount-1)) then ARow:=G.Row+1 else
    if ((ARow>G.Row) and (G.Row>1)) then ARow:=G.Row-1 ;
   Cancel:=True ;
   END ;
end;

procedure TFLettrage.WMGetMinMaxInfo(var MSG: Tmessage);
begin
{OKOK}
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFLettrage.BJustifClick(Sender: TObject);
begin
{OKOK}
if RL.Auxiliaire='' then JustifSolde(RL.General,fbGene) else JustifSolde(RL.Auxiliaire,fbAux) ;
end;

procedure TFLettrage.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFLettrage.BMenuZoomclick(Sender: TObject);
begin
PopZoom97(BMenuZoom,POPZ) ;
end;

procedure TFLettrage.BMenuZoomMouseEnter(Sender: TObject);
begin
PopZoom97(BMenuZoom,POPZ) ;
end;

procedure TFLettrage.GeneValideClick(Sender: TObject);
begin
DATEGENERATIONExit(nil); // 14350
if BValide.Enabled then BValideClick(Nil) else
 if BValideSelect.Enabled then BValideSelectClick(Nil) ;
end;

procedure TFLettrage.BAjouteEnleveClick(Sender: TObject);
begin
{OKOK}
if BRajoute.Enabled then BRajouteClick(Nil) else
 if BEnleve.Enabled then BEnleveClick(Nil) ;
end;

procedure TFLettrage.BTFClick(Sender: TObject);
begin
if BTF.Down then
   BEGIN
   H_MODEPAIE.Caption:=HDivL.Mess[18] ;
   H_MODEPAIE_.Caption:=HDivL.Mess[18] ;
   END else
   BEGIN
   H_MODEPAIE.Caption:=HDivL.Mess[17] ;
   H_MODEPAIE_.Caption:=HDivL.Mess[17] ;
   END ;
DetailLigne(GD.Row,True) ; DetailLigne(GC.Row,False) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 15/05/2006
Modifié le ... : 24/08/2007
Description .. : - FB 13402 - LG - 15/05/2006 - on mets PieceModifie a 
Suite ........ : false pour afficher un message qd on annule le lettrage
Suite ........ : - FB 13402 - LG - 24/08/2007 - suppression d ela correction 
Suite ........ : precedente, le delettrage est fait de suite et non pas  la 
Suite ........ : sortie de la fenetre
Mots clefs ... : 
*****************************************************************}
procedure TFLettrage.DelettrageCourant ;
Var CodeL,SQL,sCode : String ;
    ii,i,k    : integer ;
    TT        : HTStrings ;
    O         : TOBM ;
    G         : THGrid ;
BEGIN
if ModeSelection then CodeL:=H_CodeL.Caption else CodeL:='' ; PieceModifiee := true ;
TT:=HTStringList.Create ;
try
for i:=1 to GD.RowCount-1 do
    BEGIN
    CodeL:=GD.Cells[ColLetD,i] ;
    if ((CodeL<>'') and (TT.IndexOf(CodeL)<0)) then TT.Add(CodeL) ;
    END ;
for i:=1 to GC.RowCount-1 do
    BEGIN
    CodeL:=GC.Cells[ColLetC,i] ;
    if ((CodeL<>'') and (TT.IndexOf(CodeL)<0)) then TT.Add(CodeL) ;
    END ;
if TT.Count<=0 then Exit ;
if TT.Count>1 then ii:=Hlettre.Execute(35,'','') else ii:=HLettre.Execute(36,'',TT[0]+' ?') ;
if ii<>mrYes then Exit ;
SQL:='UPDATE ECRITURE SET E_LETTRAGE="", E_ETATLETTRAGE="AL", E_ETAT="0000000000", E_DATEPAQUETMIN="'+UsDateTime(iDate1900)+'", E_DATEPAQUETMAX="'+UsDateTime(iDate1900)+'", '
    +'E_REFLETTRAGE="", E_COUVERTURE=0, E_COUVERTUREDEV=0, E_LETTRAGEDEV="-", E_TRESOSYNCHRO = "LET" ';{JP 26/04/04 : pour l'échéancier de la Tréso}
SQL:=SQL+' WHERE E_GENERAL="'+RL.General+'" AND E_AUXILIAIRE="'+RL.Auxiliaire+'" ' ;
sCode:='' ;
for i:=0 to TT.Count-1 do sCode:=sCode+' E_LETTRAGE="'+TT[i]+'" OR ' ;
Delete(sCode,Length(sCode)-3,4) ;
SQL:=SQL+' AND ('+sCode+')' ;
ExecuteSQL(SQL) ;
for k:=1 to 2 do
    BEGIN
    if k=1 then G:=GD else G:=GC ;
    for i:=1 to G.RowCount-1 do
        BEGIN
        if k=1 then CodeL:=G.Cells[ColLetD,i] else CodeL:=G.Cells[ColLetC,i] ;
        if TT.IndexOf(CodeL)>=0 then
           BEGIN
           //LG* 28/02/2002 utilisation de la fct CRemplirInfoLettrage
           O:=GetO(G,i) ; if O=Nil then Break ; CRemplirInfoLettrage(O) ; TraiteLeEtat(O) ;
           END ;
        if EstSelect(G,i) then InverseSelection(G,i) ;
        if k=1 then G.Cells[ColLetD,i]:='' else G.Cells[ColLetC,i]:='' ;
        END ;
    END ;
GD.Invalidate ; GC.Invalidate ; CalculDebitCredit(False) ;
ModeSelection:=False ;
PieceModifiee:=false ;
ClickEnleve(not FBoOter) ;
POPAnnulTotal.Enabled:=false ; DELettTotal.Enabled:=false ;// on rend actif ce menu une fois que l'on a lettrer un elements
finally
 TT.Free ;
end ;
END ;

procedure TFLettrage.BDelettreClick(Sender: TObject);
begin
DelettrageCourant ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 26/04/2002
Modifié le ... :   /  /    
Description .. : On recalcul la taux de la devise quand on change la date 
Suite ........ : de generation
Mots clefs ... : 
*****************************************************************}
procedure TFLettrage.DATEGENERATIONExit(Sender: TObject);
Var Err : integer ;
begin
if csDestroying in ComponentState then Exit ;
if Not DateGeneration.Enabled then Exit ;
Err:=ControleDate(DATEGENERATION.Text) ;
if (Err > 0) and (Err<3) then
   BEGIN
   HLettre.Execute(37,caption,'') ;
   if DateGeneration.CanFocus then DateGeneration.SetFocus ;
   Exit ;
   END ;
LaDateGenere:=StrTodate(DATEGENERATION.Text) ;
DEV.Taux:=GetTaux(DEV.Code,DEV.DateTaux,LaDateGenere) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 20/11/2001
Modifié le ... :   /  /
Description .. : Grise les case debit et credit si elles sont vides
Mots clefs ... :
*****************************************************************}
procedure TFLettrage.HGPostDrawCell(ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
var
 lBoGrise : boolean;
begin
 lBoGrise := ( ( ACol = ColDebit )  and ( GD.Cells[ColCredit, ARow] <>'' ) and  ( ARow > 0 ) ) or
             ( ( ACol = ColCredit ) and ( GD.Cells[ColDebit, ARow]  <>'' ) and  ( ARow > 0 ) );
 if lBoGrise then
  begin
   GD.PostDrawCell  := nil; // on debranche l'évènement lors du dessin de la grille
   SetGridGrise(ACol, ARow, GD);  // fonction de ULibWindows
   GD.PostDrawCell  := HGPostDrawCell;
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 20/11/2001
Modifié le ... :   /  /
Description .. : Gestion de l'affichage des boutons agrandir et reduire
Mots clefs ... :
*****************************************************************}
procedure TFLettrage.BReduireClick(Sender: TObject);
begin
 PCDetail.Visible  := true;
 BAgrandir.Visible := true;
 BReduire.Visible  := false;
 Rduirelaliste1.Enabled:=false ; // gestion de l'etat des menu pop
 Agrandirlaliste1.Enabled:=true ;
 // on réaffecte les align pour que les grille reprenne leur position initiale
 PCSoldes.Align:=alNone ; PCDetail.Align:=alBottom ; PCSoldes.Align:=alBottom ;
end;

procedure TFLettrage.BAgrandirClick(Sender: TObject);
begin
 PCDetail.Visible  := not PCDetail.Visible;
 BAgrandir.Visible := false;
 BReduire.Visible  := true;
 Agrandirlaliste1.Enabled:=false ; // gestion de l'etat des menu pop
 Rduirelaliste1.Enabled:=true ;
 // on réaffecte les align pour que les grille reprenne leur position initiale
 PCSoldes.Align:=alNone ; PCDetail.Align:=alBottom ; PCSoldes.Align:=alBottom ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 05/03/2002
Modifié le ... :   /  /    
Description .. : Procedure de lettrage automatique 
Suite ........ : 0 : lettrage sur reference
Suite ........ : 1 : lettrage sur passage de solde (lettrage entre deux soldes 
Suite ........ : progressif identiques )
Suite ........ : 2 : lettrage sur passage de solde à nulle
Mots clefs ... : 
*****************************************************************}
{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 22/04/2002
Modifié le ... : 13/02/2003
Description .. : -LG - 22/04/2002 - on ne rafraichit pas la grille quand 
Suite ........ : aucune solution n'est trouve
Suite ........ : 
Suite ........ : - LG - 26/04/2002 - le choix d'un journal d'ecart de change 
Suite ........ : est obligatoire pour un lettrage en devise
Suite ........ : 
Suite ........ : - 10/06/2002 - Ajout d'une nouvelle methode de lettrage 
Suite ........ : auto : lettrage combinatoire
Suite ........ :
Suite ........ : -13/12/2002 - fiche 10763 - prise en compte du parametre
Suite ........ : avec nature OD pour le lettrage par reference
Suite ........ :
Suite ........ : - 13/02/2003 - correction fuite memoire
Mots clefs ... :
*****************************************************************}
procedure TFLettrage.Calcul(vInTypeCalcul : integer) ;
var
 lTLett : TList ;
 lStDernLettrage : string;
begin
{$IFDEF COMPTA}
 if RL.LettrageDevise and (E_ECART.Value='') then
  begin
   PGIInfo('Vous devez renseigner le journal d''écart de change',Caption) ; exit ;
  end;
 lTLett:=nil ; GD.SynEnabled:= false ; GD.BeginUpdate ; lStDernLettrage:=DernLettrage ;
 try
 lTLett:=CTHGridVersTLett(GD,RL.LettrageDevise,true) ;
 case vInTypeCalcul of
  0 : LettrerReference( RL ,lTLett,FIFL.CAvecOD,true,FIFL,1,DernLettrage) ;
  1 : begin LettrerSoldeProgressif( RL , lTLett,DernLettrage) ; LettrageTotCur:=true ; end ;
  2 : begin LettrerSolde( RL , lTLett,DernLettrage) ; LettrageTotCur:=true ; end ;
  3 : begin CLettrerMontant(RL,lTLett,true,FIFL.MaxProf,FIFL.MaxDuree,DernLettrage) ; LettrageTotCur:=true ; end ;
 end; // end
 if lStDernLettrage=DernLettrage then exit ;
 CTLettVersTHGrid( GD, lTLett , RL , true ) ;
 CTLettVersTHGrid( GD, lTLett , RL , false) ;
 RemettreGrilleAJour(GD,FBoOter) ; Desectionner(GD) ;
{Affichages et re-ini}
 ModeSelection:=False ; NbSelD:=0 ; NbSelC:=0 ; IndiceRegul:=0 ;
 CodeL:='' ; PieceModifiee:=True ;
 LettrageTotCur:=False ; LettrageTotP:=False ;
 EcartChangeFranc:=0 ; EcartChangeEuro:=0 ; EcartRegulDevise:=0 ; ConvertFranc:=0 ; ConvertEuro:=0 ;
 CParCouverture.Enabled:=True ; if ((LienS1S3) or (EstSerie(S3))) then CParCouverture.Enabled:=False ;
 CParEnsemble.Enabled:=True ; CParExcept.Enabled:=True ;
 if RL.Appel in [tlSaisieTreso,tlSaisieCour] then BEGIN CParExcept.Enabled:=False ; CParExcept.Checked:=False ; END ;
 GereLesValide ;
  finally
    if lTLett<>nil then begin VideListe(lTLett) ; lTLett.free ; end ;
    GD.SynEnabled:= true ; GD.EndUpdate ;
 end;
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 26/12/2001
Modifié le ... :   /  /
Description .. : Rapprochement par référence
Mots clefs ... :
*****************************************************************}
procedure TFLettrage.Rapprochementsurrfrence1Click(Sender: TObject);
begin
 Calcul(0);
end;

//LG*
procedure TFLettrage.Lettragesurpassagedesolde1Click(Sender: TObject);
begin
 Calcul(1);
end;

procedure TFLettrage.Lettragesurpassagesoldenul1Click(Sender: TObject);
begin
 Calcul(2);
end;

procedure TFLettrage.Rapprochementsurlesmontants1Click(Sender: TObject);
begin
 Calcul(3) ;
end;

procedure TFLettrage.DELettTotalClick(Sender: TObject);
begin
  BDelettreClick ( nil ) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 23/05/2002
Modifié le ... : 17/05/2005
Description .. : Annulation total d'un ou plusieur lettrage
Suite ........ : LG - 03/12/2003 - FB 12969 - on ne supprime pas les 
Suite ........ : ecritures simplifies enregsitrees en base lors du reaffichage
Suite ........ : - LG  - 06/01/2005 - FB 14567 - on ne pouvait pas annuler 
Suite ........ : deux fois de suite le meme lettrage
Suite ........ : - LG - 17/05/2005 - FB 14567 - on ne fait pas le test 
Suite ........ : precedent si  les mouvements sont uniquement en memoire
Suite ........ : 
Mots clefs ... : 
*****************************************************************}
procedure TFLettrage.AnnulLettrage;
var
 lStCodeL : string;
 i : integer;
 lTOBligneEcr : TOB;
 lRdSolde : double ;
 lStDebit, lStCredit : string ;
 lLett : HTStringList ;
 lSql : string ;
 lQ : TQuery ;
 lInNbMem,lInNbBase : integer ;
begin
lStDebit:='E_DEBIT' ; lStCredit:='E_CREDIT' ; lRdSolde:=0 ; lInNbMem:=0 ; lInNbBase:=0 ;
// controle le solde du lettrage
i:=0 ; lLett:=HTStringList.Create ;
try
// on teste pour savoir s'il on pas deja fait une annulatin de lettrage pour ce code
if FStAnnulLettrage.IndexOf(copy(GD.Cells[ColLetD,GD.Row],1,4)) = - 1 then
begin
while i<FTOBEcr.Detail.Count do
 begin
   lTOBligneEcr:=FTOBEcr.Detail[i] ; if lTOBLigneEcr=nil then exit ;
   if (lTOBligneEcr.GetValue('SELECTIONNER')='X') and (lTOBligneEcr.GetValue('E_ETATLETTRAGE')='TL') then
     lRdSolde:=lRdSolde+lTOBligneEcr.GetValue(lStDebit)-lTOBligneEcr.GetValue(lStCredit) // pour les ecritures totalement lettre le solde doit etre egale a zero
      else
       if (lTOBligneEcr.GetValue('SELECTIONNER')='X') and (lTOBligneEcr.GetValue('E_ETATLETTRAGE')='PL') then
        begin
         // pour les partiellemnt lettre on compte les mouvements ds la grille pour chaque lettrage différents
         if lLett.IndexOf(lTOBligneEcr.GetValue('E_LETTRAGE'))=-1 then lLett.add(lTOBligneEcr.GetValue('E_LETTRAGE')) ;
         inc(lInNbMem) ;
        end;
   inc(i) ;
 end ;
end ; // if
if Arrondi(lRdSolde,V_PGI.OkDecV)<>0 then
 begin
  // la totalite des ecritures d'un mouvments totalement lettres n'est pas present ds la grille
  PGIInfo('Impossible d''annuler le lettrage, la totalité des mouvements n''est pas incluse dans la sélection !',Caption);
  exit ;
 end;
// on ajoute ce code lettre a la liste des codes annules
// on ne testerat plus l'egalite entre le nombre de mouvements present en base et le nombre de mvt present ds la grille
FStAnnulLettrage.Add(copy(GD.Cells[ColLetD,GD.Row],1,4)) ;
for i:=0 to lLett.Count-1 do
begin
 lSql:='select count(*) N from ECRITURE where E_GENERAL="'+RL.General+'" AND E_AUXILIAIRE="'+RL.Auxiliaire+'"' ;
 lSql:=lSql+' and E_LETTRAGE="'+lLett[i]+'" ' ;
 lQ:=OpenSql(lSql,true) ;
 // on compte les mouvements partiellment lettre en base
 lInNbBase:=lInNbBase+lQ.findField('N').asInteger ;
 ferme(lQ) ;
end;
if ( lInNbBase <> 0) and (Arrondi(Abs(lInNbBase-lInNbMem),V_PGI.OkDecV)<>0) then
 begin
  // le nbr de mouvements partiellement lettre n'est pas egale au nbr de mouvements en memoire
  PGIInfo('Impossible d''annuler le lettrage, la totalité des mouvements n''est pas incluse dans la sélection !',Caption);
  exit ;
 end;
i:=0 ;
while i<FTOBEcr.Detail.Count do
  BEGIN
   lTOBligneEcr:=FTOBEcr.Detail[i] ; if lTOBLigneEcr=nil then exit ;
   if lTOBligneEcr.GetValue('SELECTIONNER')='X' then
    begin
     lStCodeL:=Copy(lTOBligneEcr.GetValue('E_LETTRAGE'),1,4) ;
     // on supprime les ecritures de regul que l'on a asjoute à la grille rem on incremente pas i car on suprime la ligne courante -> i:=i-1
     if (lTOBligneEcr.GetValue('ENBASE')<>'X') and
        ( (TOBM(lTOBligneEcr).TypeEcr=TERegulLettrage) or (TOBM(lTOBligneEcr).TypeEcr=TESimplifie) ) then // FB 12969 les ecritures presente en base ne doivant pas etre supprimer sinon elles ne seraont pas enregsitrees
      begin lTOBligneEcr.Free ; continue ; end
     else
      BEGIN
       CRemplirInfoLettrage(lTOBligneEcr) ; TraiteLeEtat(TOBM(lTOBligneEcr)) ;
       { GP le 18/11/97 : Pour appel délettrage à partir de la saisie }
       if (RL.Appel=tlSaisieTreso) then lTOBligneEcr.PutValue('E_TRESOLETTRE','-') ;
      END;
    end; // if
   Inc(i) ;
   // on supprime les ecritures de regul associée au lettrage
   FEcrRegul.Delete(lStCodeL) ;
  END; // while
 GD.BeginUpdate ; RemettreGrilleAJour(GD,FBoOter) ; CalculDebitCredit(False) ;
 ModeSelection:=False ; PieceModifiee:=true ; GD.EndUpdate ;
finally
 lLett.Free ;
end ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 23/05/2002
Modifié le ... : 05/01/2005
Description .. : Annulation d'une ligne de lettrage
Suite ........ : 
Suite ........ : LG - 30/11/2003 - FB 12698 - on ne pouvait plus annuler
Suite ........ : un lettrage partiel
Suite ........ : LG - 30/09/2004 - FB 14567 - on ne pouvait pas nnuler 
Suite ........ : deux fois de suite un meme mvt
Suite ........ : LG - 05/01/2005 - FB 15207 - qd on annul un lettrage, on 
Suite ........ : supprime les ecritures de regul associées
Mots clefs ... : 
*****************************************************************}
procedure TFLettrage.AnnulLigneLettrage;
var
 lStCodeL : string;
 O : TOBM;
 i : integer;
 lBoTrouver : boolean;
 lTOBSauve : TOB;
 lStOldChoixValid : string ;
 lTOBligneEcr : TOB;
 lQ : TQuery ;
 lInNbMem : integer ;
 lSql : string ;
begin
  // controle que tous les mouvements sont presnets ds la grille
  i:=0 ; lInNbMem:=0 ;
  // on teste pour savoir s'il on pas deja fait une annulatin de lettrage pour ce code
  if FStAnnulLettrage.IndexOf(copy(GD.Cells[ColLetD,GD.Row],1,4)) = - 1 then
   begin
    while i<FTOBEcr.Detail.Count do
     begin
       lTOBligneEcr:=FTOBEcr.Detail[i] ; if lTOBLigneEcr=nil then exit ;
       if (Copy(lTOBligneEcr.GetValue('E_LETTRAGE'),1,4)=copy(GD.Cells[ColLetD,GD.Row],1,4)) then inc(lInNbMem) ;
       inc(i) ;
     end ;
    lSql:='select count(*) N from ECRITURE where E_GENERAL="'+RL.General+'" AND E_AUXILIAIRE="'+RL.Auxiliaire+'"' ;
    lSql:=lSql+' and E_LETTRAGE="'+Copy(GD.Cells[ColLetD,GD.Row],1,4)+'" ' ;
    lQ:=OpenSql(lSql,true) ;
    if (lQ.findField('N').asInteger<>0) and (lQ.findField('N').asInteger<>lInNbMem) then
     begin
      // le nbr de mouvements partiellement lettre n'est pas egale au nbr de mouvements en memoire
      PGIInfo('Impossible d''annuler le lettrage, la totalité des mouvements n''est pas incluse dans la sélection !',Caption);
      exit ;
     end;
    ferme(lQ) ;
    // on ajoute ce code lettre a la liste des codes annules
    // on ne testerat plus l'egalite entre le nombre de mouvements present en base et le nombre de mvt present ds la grille
    FStAnnulLettrage.Add(copy(GD.Cells[ColLetD,GD.Row],1,4)) ;
   end ;// if

 lTOBSauve:=TOB.Create('',nil,-1) ; lBoTrouver:=false ;
 // delettrage d'un ligne en particulier
 GD.Cells[ColLetD,GD.Row]:='' ; GD.Cells[GD.ColCount-1,GD.Row]:=' ' ; O:=GetO(GD,GD.Row) ;
 GD.BeginUpdate ;
 try
 try
  lTOBSauve.Dupliquer(FTOBEcr,true,true) ; lStCodeL:=Copy(O.GetValue('E_LETTRAGE'),1,4) ;
  // on supprime les ecritures de regul associée au lettrage
  FEcrRegul.Delete(lStCodeL) ;
  CRemplirInfoLettrage(O) ; TraiteLeEtat(O) ;
  i := 0 ;
  while i<FTOBEcr.Detail.Count do
   begin
    lTOBligneEcr:=FTOBEcr.Detail[i] ; if lTOBLigneEcr=nil then exit ;
    // on supprime les ecritures de regul que l'on a asjoute à la grille rem on incremente pas i car on suprime la ligne courante -> i:=i-1
     if (lTOBligneEcr.GetValue('ENBASE')<>'X') and
        ( (TOBM(lTOBligneEcr).TypeEcr=TERegulLettrage) or (TOBM(lTOBligneEcr).TypeEcr=TESimplifie) ) then
      begin lTOBligneEcr.Free ; continue ; end ;
    Inc(i) ;
   end ;// while
 RemettreGrilleAJour(GD,FBoOter) ;
 // recherche une ligne correspond a l'ancien lettrage
 for i:=1 to GD.RowCount-1 do
  BEGIN
   O:=GetO(GD,i) ;
   if O=Nil then
   begin
    ModeSelection:=false ; NbSelD:=0 ; NbSelC:=0 ; IndiceRegul:=0 ;
    CodeL:='' ; PieceModifiee:=false ;
    LettrageTotCur:=False ; LettrageTotP:=False ;
    EcartChangeFranc:=0 ; EcartChangeEuro:=0 ; EcartRegulDevise:=0 ; ConvertFranc:=0 ; ConvertEuro:=0 ;
    GereLesValide ; RemettreGrilleAJour(GD,FBoOter) ; GD.EndUpdate ;
    exit;
   end; // if
   if Copy(O.GetMvt('E_LETTRAGE'),1,4)=lStCodeL then BEGIN lBoTrouver:=true ; break ; END;
  END;
 if lBoTrouver then GD.Row:=i ; Selection(False,true,true) ;
 GD.EndUpdate ; GD.Refresh ;  CalculDebitCredit(false) ;
 lStOldChoixValid:=FIFL.ChoixValid ;
 FIFL.ChoixValid:='' ;
 if not LettrageMemoire then
  BEGIN
   FTOBEcr.Dupliquer(lTOBSauve,true,true) ;
   // remise a jour des parametre de lettrage
   RemettreGrilleAJour(GD,FBoOter) ;
   ModeSelection:=False ; NbSelD:=0 ; NbSelC:=0 ; IndiceRegul:=0 ;
   CodeL:='' ; PieceModifiee:=True ;
   LettrageTotCur:=False ; LettrageTotP:=False ;
   EcartChangeFranc:=0 ; EcartChangeEuro:=0 ; EcartRegulDevise:=0 ; ConvertFranc:=0 ; ConvertEuro:=0 ;
  END;
  FIFL.ChoixValid:=lStOldChoixValid ;
 finally
  lTOBSauve.Free ; GD.EndUpdate ;
 end;

 except
  on E : Exception do begin FStLastError:=E.Message ; AnnulationSurErreur('Délettrage partiel') ; end;
 end;

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 19/12/2001
Modifié le ... : 23/05/2002
Description .. : Annulation de lettrage de l'ecriture courante
Suite ........ : 
Suite ........ : 03/04/2002 : on peut deselectionne une ligne ayant des
Suite ........ : ecritures de regularisation
Suite ........ : 
Suite ........ : - 17/04/2002 - quand on delettre un code lettrage on
Suite ........ : supprime les ecritures de regul associés
Suite ........ : 
Suite ........ : - 25/04/2002 - Passer le flag PieceModifie a true lors d'une
Suite ........ : annulation de lettrage
Suite ........ : 
Suite ........ : 
Suite ........ : version 582
Suite ........ : - 23/05/2002 - la fct est scinde en deux
Mots clefs ... : 
*****************************************************************}
procedure TFLettrage.DeLettPartielClick(Sender: TObject);
var
 O:TOBM;
begin
// securite
if not FBoUneGrille then exit ;
O:=GetO(GD,GD.Row) ; if O=Nil then exit ; if O.GetMvt('E_LETTRAGE')='' then exit; // ecriture non lettre -> on sort

if (EstSelect(GD,GD.Row)) then AnnulLettrage
else AnnulLigneLettrage;

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 28/11/2001
Modifié le ... : 18/12/2001
Description .. : Fonction de selection/desselection de l'ensemble des enr.
Suite ........ : Param :
Suite ........ : -AvecLet : si true on selectionne les mouvements lettres
Suite ........ : -MemeMvt : si true on ne selectionne que les mvt ayant le
Suite ........ : m. journal et la m. date comptable que l'enregistrement
Suite ........ : courant. Sert pour le Alt+M
Suite ........ : -MemeMois : si true on selectionne que les mvt du m. journal
Suite ........ : et m. mois que l'enregistrement courant
Mots clefs ... :
*****************************************************************}
procedure TFLettrage.ClickToutSel(AvecLet : boolean = false; MemeMvt : boolean = false; MemeMois : boolean = false);
var
 i,Col:integer;
 G:THGrid;
 CritD,CritM :string;
 O:TOBM;
begin
 if GD.Focused then BEGIN G:=GD ; Col:=ColLetD ; END else if GC.Focused then BEGIN Col:=ColLetC ; G:=GC END else Exit ;
 if G.Row<=0 then Exit; NbSelD:=0 ; NbSelC:=0 ;
 O:=GetO(G,G.Row); if O=nil then exit; // on recupere les info de l'enregsitrement courant pour la selection des m. mvt
 G.SynEnabled:=false;
 CritD:=VarToStr(O.GetMvt('E_DATECOMPTABLE'))+O.GetMvt('E_JOURNAL');
 CritM:=VarToStr(GetPeriode(O.GetMvt('E_DATECOMPTABLE')))+O.GetMvt('E_JOURNAL');
 for i:=1 to G.RowCount-2 do // RowCount-2 car la grille a tj. une ligne vide
  BEGIN
   if (G.Cells[Col,i]<>'') and not AvecLet then continue; // on passe les mvt. lettres
   if MemeMois then
    BEGIN
     O:=GetO(G,i) ; if O=nil then exit ; if (VarToStr(GetPeriode(O.GetMvt('E_DATECOMPTABLE')))+O.GetMvt('E_JOURNAL')) <> CritM then continue;
    END
   else
   if MemeMvt then
    BEGIN
     O:=GetO(G,i) ; if O=nil then exit ; if (VarToStr(O.GetMvt('E_DATECOMPTABLE'))+O.GetMvt('E_JOURNAL')) <> CritD then continue;
    END;
   if ( G.Cells[G.ColCount-1,i]='+') then
    BEGIN
     if G.Cells[ColDebit,i]<>'' then Dec(NbSelD) else Dec(NbSelC);
     G.Cells[G.ColCount-1,i]:=' '
    END
   else
    BEGIN
     if G.Cells[ColDebit,i]<>'' then Inc(NbSelD) else Inc(NbSelC);
     G.Cells[G.ColCount-1,i]:='+';
    END;
  END;
 G.SynEnabled:=true ; G.Refresh;
 ModeSelection:=((NbSelD>0) or (NbSelC>0)) ; GereLesValide ;
 CParCouverture.Enabled:=Not ModeSelection ; if ((LienS1S3) or (EstSerie(S3))) then CParCouverture.Enabled:=False ;
 CParEnsemble.Enabled:=Not ModeSelection ; CalculDebitCredit(False) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 24/12/2001
Modifié le ... : 24/12/2001
Description .. : Le cursor se place automatiquement sur la premiere ligne du
Suite ........ : mois suivant
Mots clefs ... :
*****************************************************************}
procedure TFLettrage.PasseMoisSuivant;
var
 i : integer; G : THGrid;
 Periode : integer; O : TOBM;
begin
 if GD.Focused then BEGIN G:=GD  END else if GC.Focused then BEGIN  G:=GC END else Exit ;
 if G.Row<=0 then Exit;
 O:=GetO(G,G.Row) ; if O=nil then exit ;
 G.SynEnabled:=false ; Periode:=O.GetMvt('E_PERIODE') ;
 for i:=G.Row to G.RowCount-2 do
  BEGIN
    O:=GetO(G,i) ; if O=nil then exit ; if O.GetMvt('E_PERIODE')<>Periode then break;
  END; // for
 if i<(G.RowCount-1) then G.Row:=i else if i=(G.RowCount-1) then
 BEGIN G.Row:=1 ; end;//PasseMoisSuivant; END;
 G.SynEnabled:=true ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 15/01/2002
Modifié le ... : 02/05/2002
Description .. : Le curseur se deplace se le prochain libelle identique. On 
Suite ........ : utilise le parametrage du lettrage automatique : egalite strict 
Suite ........ : des ref pour choisir la prochaine refrence
Suite ........ : 
Suite ........ : - LG - 02/04/2002 - correction si le libelle n'existait pas onn 
Suite ........ : bouclait indéfiniment
Mots clefs ... :
*****************************************************************}
procedure TFLettrage.PasseLibelleSuivant( vStLibelle : string = '');
var
 G : THGrid ;
 Crit : string;
 O : TOBM ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 06/07/2004
Modifié le ... : 03/11/2004
Description .. : - FB 11908 - correction du deplacement sur la der cellule
Suite ........ : 
Mots clefs ... : 
*****************************************************************}
 procedure RechercheLib( vNbBoucle : integer ; vStLibelle : string = '' ; vInIndex : integer = -1 );
 var
  i :integer ;
 begin
  if vNbBoucle > 1 then exit ; // on a deja fait une fois le tour des enregistrements de la grille on sort
  if vInIndex = -1 then
   vInIndex := (G.Row+1) ;
  i := vInIndex ; // ne pas enlever...
  for i:=vInIndex to (G.RowCount-2) do
  begin
   O:=GetO(G,i) ; if O=nil then continue ;
   if MemeRef(Crit,O.GetMvt('E_LIBELLE'),Client,FIFL) then break ;
  end; // for
  if i=0 then
   begin
    G.Row:=1 ;
    RechercheLib(vNbBoucle,Crit,0) ;
   end
    else
     if i<(G.RowCount-1) then
      G.Row:=i
       else
       if i=(G.RowCount-1) then
        begin
         G.Row:=1 ;
         Inc(vNbBoucle) ;
         RechercheLib(vNbBoucle,Crit,1) ;
       end;
 end;

begin
 if GD.Focused then G:=GD else if GC.Focused then G:=GC else Exit ;
 if G.Row<=0 then Exit ;
 O:=GetO(G,G.Row) ; if O=nil then exit ;
 G.SynEnabled:=false ; G.BeginUpdate ;
 if vStLibelle='' then Crit:=O.GetMvt('E_LIBELLE') else Crit:=vStLibelle;
 if Crit<>'' then RechercheLib(0,Crit) ;
 G.SynEnabled:=true ; G.EndUpdate ; G.Refresh ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 15/01/2002
Modifié le ... : 25/01/2002
Description .. : se deplace sur le jour suivant
Mots clefs ... :
*****************************************************************}
procedure TFLettrage.PasseJourSuivant;
var
 i : integer ; G : THGrid;
 Crit : TDateTime ; O : TOBM;
begin
 if GD.Focused then G:=GD else if GC.Focused then G:=GC else Exit ;
 if G.Row<=0 then Exit ; O:=GetO(G,G.Row) ; if O=nil then exit ;
 G.SynEnabled:=false ; Crit:=O.GetMvt('E_DATECOMPTABLE')+1 ;
 for i:=(G.Row+1) to G.RowCount-2 do
  BEGIN
    O:=GetO(G,i) ; if O=nil then exit ; if Crit<(VarToDateTime(O.GetMvt('E_DATECOMPTABLE'))) then break;
  END; // for
 if i=0 then G.Row:=1 else if i<(G.RowCount-1) then G.Row:=i else if i=(G.RowCount-1) then G.Row:=1 ;
 G.SynEnabled:=true ;
end;

procedure TFLettrage.Desectionner(G : THGrid);
var
 i : integer;
begin
for i:=1 to G.RowCount-1 do if JePrends(G,i) then
 G.Cells[G.ColCount-1,i]:=' ';
end;

procedure TFLettrage.BCalculMouseEnter(Sender: TObject);
begin
// PopZoom97(BCalcul,PopCalcul) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 25/01/2002
Modifié le ... :   /  /
Description .. : INsert les ecriture de regularisation dans la grille
Mots clefs ... :
*****************************************************************}
procedure TFLettrage.AjouteEcrGrille(Sender: TOB);
var
 i : integer ;
 lTOBMEcrRegul : TOBM ;
begin
 for i:=0 to Sender.Detail.Count - 1 do
   if Sender.Detail[i].GetValue('SELECTIONNER')='X' then
    BEGIN
     // on ajoute la nouvelle tob à la liste des ecriture
     lTOBMEcrRegul:=TOBM.Create(EcrGen,'',False,FTOBEcr) ;
     lTOBMEcrRegul.Dupliquer(Sender.Detail[i],false,true) ;  CAjouteChampsSuppLett(lTOBMEcrRegul) ;
     lTOBMEcrRegul.PutValue('SELECTIONNER','X') ; GD.InsertRow(GD.RowCount) ;
     GD.Row:=GD.RowCount-1 ; GD.Objects[GD.ColCount-1,GD.Row]:=lTOBMEcrRegul ;
    END; // if
end;


procedure TFLettrage.BForceVadideClick(Sender: TObject);
begin
 FInOrdreTri:=BTri.Tag ; RemettreGrilleAJour(GD,FBoOter) ;
 // on inverse l'option de tri
  if BTri.Tag=1 then
    begin
    BTri.Tag   := 2 ;
    BTri.Glyph := BTriDate.Glyph ;
    BTri.Hint  := BTriDate.Hint ;
    end
  else
    begin
    BTri.Tag   := 1 ;
    BTri.Glyph := BTriCode.Glyph ;
    BTri.Hint  := BTriCode.Hint ;
    end ;
end;


procedure TFLettrage.Soldeautomatique1Click(Sender: TObject);
var
 lIFL : RINFOSLETT; // info de lettrage
begin
 lIFL:= FIFL ; FIFL.ChoixValid:='AL1' ;
 try
  LettrageMemoire ;
 finally
  FIFL:=lIFL ;
 end; // try
end;

procedure TFLettrage.Lettragepartiel1Click(Sender: TObject);
var
 lIFL : RINFOSLETT; // info de lettrage
begin
 lIFL:= FIFL ; FIFL.ChoixValid:='AL2' ;
 try
  LettrageMemoire ;
 finally
  FIFL:=lIFL ;
 end; // try
end;

procedure TFLettrage.Passagedunecrituresimplifi1Click(Sender: TObject);
var
 lIFL : RINFOSLETT; // info de lettrage
begin
 lIFL:= FIFL ; FIFL.ChoixValid:='AL3' ;
 try
  LettrageMemoire ;
 finally
  FIFL:=lIFL ;
 end; // try
end;

procedure TFLettrage.Rapprochementsimple1Click(Sender: TObject);
begin
if Action=taConsult then Exit ;
Calcul(3) ;
end;

procedure TFLettrage.Rapprochementcombinatoire1Click(Sender: TObject);
begin
if Action=taConsult then Exit ;
ClickCombi(False);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 03/04/2002
Modifié le ... :   /  /    
Description .. : En cas d'erreur grave, on affiche un message et on sort 
Suite ........ : directement du lettrage
Mots clefs ... : 
*****************************************************************}
procedure TFLettrage.AnnulationSurErreur( vStFonction : string = '');
var
 s : string;
begin
 s:='Erreur dans le lettrage, toutes les opérations vont être annulées !';
 if vStFonction='' then vStFonction:='Lettrage';
 if FStLastError<>'' then
  s:=s+#10#13+FStLastError ;  
 PGIError(s,vStFonction) ; ModeSelection:=false ;  PieceModifiee:=false ; FBoFermer:=true ;
 ModalResult:=mrAbort ; Close ;
end;

procedure TFLettrage.FormActivate(Sender: TObject);
begin
 //if FBoFermer then begin modalresult:=mrok; close; end ;
end;

procedure TFLettrage.POPPCLPopup(Sender: TObject);
var
 O : TOBM ;
begin
if Action=taConsult then Exit ;
// test des lignes
O:=GetO(GD,GD.Row) ; if O=nil then exit ;
POPAnnulSelect.Enabled:= (O.GetValue('E_LETTRAGE')<>'') ;
Passagedunecrituresimplifie1.Enabled:= not RL.LettrageEnSaisie ;
end;

procedure TFLettrage.Validerlelettrage1Click(Sender: TObject);
var
 EnCouv:boolean;
begin
 EnCouv:=(Not POutils.Enabled) ;
 if EnCouv then
 BEGIN NextControl(Self) ; BCValideClick(Nil) ; END
 else ClickValide ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 10/11/2004
Modifié le ... :   /  /
Description .. : - LG - 10/11/2004 - FB 14880 - ajout d'un menu sur le F11, 
Suite ........ : lettrage sur exception
Mots clefs ... : 
*****************************************************************}
procedure TFLettrage.Lettrageparexception1Click(Sender: TObject);
begin
 Lettrageparexception1.Checked:=not  Lettrageparexception1.Checked ;
 LEX.Checked := Lettrageparexception1.Checked ;
 CParExcept.Checked:=Lettrageparexception1.Checked ; CParCouvertureClick(nil) ;
end;

procedure TFLettrage.Recherchedanslacomptabilit1Click(Sender: TObject);
begin
{$IFNDEF IMP}
  if ctxPcl in V_Pgi.PgiContexte then
    CPLanceFiche_CPRechercheEcr(True)
  else
    MultiCritereMvt(taConsult,'N',False) ;
{$ENDIF}
end;

procedure TFLettrage.EBStopClick(Sender: TObject);
begin
DemandeStop(Caption) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 13/04/2006
Modifié le ... :   /  /    
Description .. : - LG - 13/04/2006 - qd on fait un lettrage depuis la 
Suite ........ : consultation des comptes, on est force le mode 'solde 
Suite ........ : progressif'
Mots clefs ... : 
*****************************************************************}
procedure TFLettrage.InitLettrage ;
begin
//LG*
try
// chargement des parametre de lettrage
FBoFermer:=false ;  FStAnnulLettrage := HTStringList.Create ;
FIFL:=ChargeInfosLett ; FInfoEcr:=TInfoEcriture.Create ; FEcrRegul:=TEcrRegul.Create(FInfoEcr) ; FInOrdreTri:=0 ;
if not FBoChangementMode then
 BEGIN
  if FBoFromTOB then FBoUneGrille:= true
   else
  if FIFL.ModeLettrage<>'' then FBoUneGrille:=(FIFL.ModeLettrage = 'PCL')
  else FBoUneGrille:= ( ctxPCL in V_PGI.PGIContexte ) ;
  BPresentationPCL.Visible:= not FBoUneGrille ;
  BPresentationPGE.Visible:= not BPresentationPCL.Visible ;
 END ;
if FBoUneGrille then
 BEGIN
 self.PopupMenu:=POPPCL ;
// SBO 26/04/2004 : Correction menu pop, l'initialisation doti se faire une seule fois, reporté dans le FormCreate
  Lettragepartiel2.Enabled:=OkLettrage ; Passagedunecrituresimplifie1.Enabled:=OkLettrage ;
  Soldeautomatique2.Enabled:=OkLettrage ; Enleverlestotalementslettrs1.Enabled:=OkLettrage ;
  POPAnnulSelect.Enabled:=OkLettrage ; POPAnnulTotal.Enabled:=OkLettrage ; Zoom1.enabled:=OkLettrage ;
  POPAux.Enabled:=(FQ <> nil) ;
 END
 else
 begin
// SBO 26/04/2004 : Correction menu pop, on repasse en menu de base quand on change d'affichage
   self.PopupMenu:=POPS ;
 end ;

FInLastNumero:=0 ; FStLastDateComptable:=iDate1900 ; if not FBoFromTOB  then FTOBEcr:=TOB.Create('',nil,-1) ;
FLDAnnul:=TList.Create ; FLCAnnul:=TList. Create ;
if FBoUneGrille then GD.PostDrawCell:=HGPostDrawCell else GD.PostDrawCell:=nil ;
GD.GetCellCanvas:=GetCellCanvasD ; GC.GetCellCanvas:=GetCellCanvasC ;
WMinX:=Width ; WMinY:=Height ; TraiteReturn:=False ;
TLAD:=TLIST.Create ; TLAC:=TList.Create ; TTSOLU:=TList.Create ; 
InitialiseLettrage ; CurDec:=V_PGI.OkDecV ; LeEtab:='' ;
HG_GENERAL.Caption:='' ; HT_AUXILIAIRE.Caption:='' ; H_CODEL.Caption:='' ;
Pages.ActivePageIndex:=0;
RegLoadToolbarPos(Self,'Lettrage') ;
except
 on E : Exception do begin FStLastError:=E.Message ; AnnulationSurErreur('Initialisation des paramètres de lettrage') ; end;
end;
end;

procedure TFLettrage.BPresentationPGEClick(Sender: TObject);
begin
 ChangePresentation(false) ;
end;

procedure TFLettrage.BPresentationPCLClick(Sender: TObject);
begin
 ChangePresentation(true) ;
end;

procedure TFLettrage.ChangePresentation( vModeUneGrille : boolean );
var
 Action : TCloseAction ;
begin
 GD.VidePile(false) ; GC.VidePile(false) ;
 Action:=caNone ; FBoChangementMode:=true ; FBoUneGrille:= vModeUneGrille ;
 try
  BPresentationPCL.Visible:= not vModeUneGrille ;
  BPresentationPGE.Visible:= not BPresentationPCL.Visible ;
  FormClose(nil,Action) ;
  InitLettrage ;
  ShowLettrage(True) ;
 finally
  FBoChangementMode:=false ;
 end;
end;

function TFLettrage.ControleLigne : boolean ;
var
 lTOBListeBor : TOB ;
begin
 // on rafraichissit la liste des bordereaux bloques
 lTOBListeBor:=CGetListeBordereauBloque ;
 result := CControleLigneBor(GD,lTOBListeBor) and CControleLigneBor(GC,lTOBListeBor) ;
 lTOBListeBor.Free ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 14/06/2004
Modifié le ... : 19/09/2006
Description .. : - FB 10647 - LG - on peu changer d'auxi directement ds la
Suite ........ : fenetre
Suite ........ : - LG - 20/08/2004 - FB 14018 - Passage au compte suivant 
Suite ........ : reste avec les montant du compte précédent mais l'intitulé 
Suite ........ : du nouveau compte
Suite ........ : - LG - 20/08/2004 - FB 14019 - message de confirmation qd 
Suite ........ : on change d'auxi
Suite ........ : - LG - 15/09/2004 - FB 14564 - le lettrage n'eatit pas 
Suite ........ : enregistrer en base qd on changeait d'auxi
Suite ........ : - LG - 22/06/2006 - FB 18365 - violation d'acces qd on 
Suite ........ : passait sur l'auxi suivant
Suite ........ : - LG - 22/06/2006 - FB 16321 - violation d'acces qd on 
Suite ........ : passait a au compte suivant pour les tic ou tid
Suite ........ : - LG - 19/09/2006 - FB 18575 - on ne pouvait pas passer a 
Suite ........ : l'auxiliare suivant s'il contenait AND ds son nom
Mots clefs ... : 
*****************************************************************}
procedure TFLettrage.ChangerAux( vBoSuiv : boolean );
 {$IFNDEF EAGLCLIENT}
var
 Action : TCloseAction ;
 lSQL : string ;
 lQ : TQuery ;
 lSt1,lSt2 : string ;
 lInIndex : integer ;
 FOldValue : integer ;
 {$ENDIF}
begin
 {$IFNDEF EAGLCLIENT}
 if FBoAuxSuiv or ( RL.Appel=tlSaisieCour ) or ( not BDown.Visible ) then exit ;
 if not FBoAucunEnr then
  begin
  if PieceModifiee and ( PGIAsk('Valider le lettrage en cours ?') <> mrYes ) then exit ;
  FBoAuxSuiv := true ;
  GeneValideClick(nil) ; // appel lettrageMemoire
  GeneValideClick(nil) ; // appel lettragefichier
  end ;

 FBoAuxSuiv := true ;

 if (FQ = nil)  then
  begin

  if vBoSuiv then
   lSQL := 'SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE>"' + RL.Auxiliaire + '" '
    else
     lSQL := 'SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE<"' + RL.Auxiliaire + '" ';

  // Nature des comptes
  if FStNatAux <> '' then
    lSQL := lSQL + ' AND T_NATUREAUXI = "' + FStNatAux + '"';

  // Défilement des comptes
  case FInTypeMvt of
    0: ; // Tous les Comptes
    1: lSQL := lSQL + ' AND (T_TOTALDEBIT - T_TOTALCREDIT) <> 0 ';
    2: lSQL := lSQL + ' AND (T_TOTALDEBIT <> 0 OR T_TOTALCREDIT <> 0) ';
  end;

  if vBoSuiv then
   lSQL := lSQL + ' ORDER BY T_AUXILIAIRE '
    else
     lSQL := lSQL + ' ORDER BY T_AUXILIAIRE DESC ';

  lQ := OpenSQL(lSQL, True);

  if not lQ.EOF then
   RL.Auxiliaire := lQ.FindField('T_AUXILIAIRE').AsString ;

  Ferme(lQ);
  end
   else
    begin
     if vBoSuiv then
      begin
       FQ.Next ;
       if FQ.EOF then FQ.First
      end
       else
        begin
         FQ.Prior ;
         if FQ.BOF then FQ.Last ;
        end ;
     RL.General := FQ.FindField('e_general').asString ;
     RL.Auxiliaire := FQ.FindField('e_auxiliaire').asString ;
    end ;

 if FBoFromSaisie then
  begin
   lInIndex := Pos('E_GENERAL' , RL.CritMvt ) ;
   lSt1     := Copy( RL.CritMvt , 1 , lInIndex - 1 ) ;
   lSt1     := lSt1 + 'E_GENERAL="' + RL.General + '" ' ;
   lInIndex := Pos('E_AUXILIAIRE' , RL.CritMvt ) ;
   lSt2     := Copy( RL.CritMvt , lInIndex + 12 , length( RL.CritMvt ) ) ;
   lInIndex := Pos(' AND ' , lSt2 ) ;
   if lInIndex > 0 then
    lSt2     := Copy( lSt2 , lInIndex , length( lSt2 ) )
     else
      lSt2 := '' ;
   lSt1     := lSt1 + ' AND E_AUXILIAIRE="' + RL.Auxiliaire + '" ' + lSt2 ;
  end ;

 RL.CritMvt := lSt1 ;
 Action:=caNone ;
 GD.VidePile(false) ; GC.VidePile(false) ;
 FormClose(nil,Action) ;
//initEcran(true) ;
 FOldValue:=FInOrdreTri ;
 InitLettrage ;
 FInOrdreTri:=FOldValue ;
 ShowLettrage(True) ;
 FBoAuxSuiv := false ;
 {$ENDIF}
end;

procedure TFLettrage.POPAUXSuivClick(Sender: TObject);
begin
 ChangerAux(True) ;
end;

procedure TFLettrage.POPAUXPRECClick(Sender: TObject);
begin
 ChangerAux(False) ;
end;

procedure TFLettrage.BUPClick(Sender: TObject);
begin
 ChangerAux(False) ;
end;

procedure TFLettrage.BDownClick(Sender: TObject);
begin
 ChangerAux(True) ;
end;

end.


