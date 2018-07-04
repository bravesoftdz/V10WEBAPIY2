UNIT Saisie ;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Grids, Mask, ExtCtrls, ComCtrls, Buttons, Hspliter, Ent1, Menus,
  HEnt1, UtilPGI, HFLabel, About, HStatus, hmsgbox, HQry, Hctrls,
  HDebug, HLines, HSysMenu, HPop97, HTB97, ed_tools, HPanel,ULibAnalytique,
  Math, // pour la fonction Min
  LookUp, UTobDebug,

{$IFDEF VER150}
 Variants,
{$ENDIF}

 // SelGuide ,
{$IFNDEF CMPGIS35}
  uTOFCPMulGuide,     // pour LanceGuide, SelectGuide
{$ENDIF}
{$IFDEF AMORTISSEMENT}
  IMMO_TOM,	// pour FicheImmobilisation
{$ENDIF}
{$IFDEF EAGLCLIENT} //===========================================> EAGL
  Maineagl,
  DEVISE_TOM,         // pour Zoom sur fiche devise
  SUIVCPTA_TOM,       // pour Zoom sur fiche scénario
  CPMULSAISLETT_TOF,  // pour LettrerEnSaisie
  CPCHANCELL_TOF,     // pour FicheChancel
{$ELSE}	//=======================================================> NON EAGL
  MenuOLG,
  Fe_Main,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  DB,
  {$IFNDEF GCGC}
    //SelGuide,		// pour LanceGuide, SelectGuide
    { FQ 19684 BVE 21.05.07 }
    // Scenario,
    SUIVCPTA_TOM, // pour Zoom sur fiche scénario
    ValPerio,
    { FQ 19816 BVE 27.04.07 }
    //SaisLett,		// pour LettrerEnSaisie
    CPMULSAISLETT_TOF,  // pour LettrerEnSaisie
    { END FQ 19816 }
  {$ENDIF}
  CPChancell_TOF,	// pour FicheChancel,
  Devise_TOM,	        // pour FicheDevise,
  UtilSoc,  	        // EAGL : EN ATTENTE --> MarquerPublifi...
{$ENDIF} //=======================================================> FIN NON EAGL

//=======================================================> COMMUN EAGL / NON-AGL
  SaisTVA,			// pour TFSaisTVA
{$IFNDEF CCS3}
    TiersPayeur,   // pour GenerePiecesPayeur, SupprimePieceTP, ExisteLettrageSurTP, ZoomPieceTP
  {$IFNDEF GCGC}
  //SaisTVA,			// pour TFSaisTVA
  DocRegl,
  {$ENDIF}
{$ENDIF}
{$IFNDEF GCGC}
  SaisComp,
  SaisEnc,    // pour InfosTvaEnc
  SaisBase,   // pour Enr_Base et SaisieBasesHT
  SaisVisu,
  CPJOURNAL_TOM,  // pour Zoom sur fiche journal
{$ELSE !GCGC}
{$IFDEF CMPGIS35}
  SaisComp,
{$ENDIF CMPGIS35}
{$ENDIF}
  eSaisAnal, 
  UTOB, Filtre, Choix, Formule,
  SaisUtil,
  SaisComm,     // RecupTotalPivot
  LettUtil, FichComm, HCompte, UtilSais,
  EcrPiece_TOF,	// pour LanceZoomPieceGC
{$IFNDEF CMPGIS35}
  ECRLIGNES_TOF, // pour LanceZoomLignesGC
{$ENDIF}
  EcheMono,			// pour SaisirMonoEcheance,
  SaisTaux1,		// pour SaisieNewTaux2000,
  Echeance,			// pour CalculModeRegle et	SaisirEcheance,
  ParamSoc,			// pour GetParamSoc et SetParamSoc,
  UiUtil,
  ImgList,
{$IFDEF SCANGED}
  UtilGed,
  UGedFiles ,
  UGedViewer,
  AnnOutils,
{$ENDIF}
{$IFDEF MODENT1}
  CPTypeCons,
{$ENDIF MODENT1}

{$IFDEF COMPTAAVECSERVANT}
  EntImo, FicheImo,
{$ENDIF}
  AGLInit,
  UlibEcriture, HImgList, TntButtons, TntStdCtrls, TntExtCtrls, TntGrids


  ;


Function  TrouveSaisie(Q : TQuery ; Var M : RMVT ; Simul : String3) : Boolean ;
Function  TrouveEtLanceSaisie(Q : TQuery ; TypeAction : TActionFiche ; Simul : String3 ; ModLess : Boolean = FALSE) : boolean ;
procedure SaisieGuidee (StG : String3) ;
Function  LanceSaisie ( QListe : TQuery ; Action : TActionFiche ; Var M : RMVT ; ModLess : Boolean = FALSE) : boolean ;
Function  LanceSaisieMP ( M : RMVT ; TOBMPOrig,TOBMPEsc : TOB ; GenereAuto : boolean ; LettrageMP : Boolean ; TOBParamEsc : TOB = nil; Reedition : Boolean = False; ModeleBordereau : String = '') : boolean ;
Procedure LanceMultiSaisie ( LesM : TList ) ;
{$IFDEF SCANGED}
procedure SaisieMyAfterImport (Sender : TObject; FileGuid: string; var Cancel: Boolean) ;
{$ENDIF}

type
  TFSaisie = class(TForm)
    GS              : THGrid;
    PEntete: THPanel;
    E_JOURNAL       : THValComboBox;
    H_JOURNAL       : THLabel;
    H_DATECOMPTABLE : THLabel;
    E_NATUREPIECE   : THValComboBox;
    H_NATUREPIECE   : THLabel;
    E_ETABLISSEMENT : THValComboBox;
    H_ETABLISSEMENT : THLabel;
    H_NUMEROPIECE   : THLabel;
    E_DEVISE        : THValComboBox;
    E_DEVISE_       : THValComboBox;
    H_DEVISE        : THLabel;
    H_EXERCICE      : THLabel;
    PPied: THPanel;
    G_LIBELLE: THLabel;
    H_SOLDE         : THLabel;
    T_LIBELLE: THLabel;
    H_NUMEROPIECE_  : THLabel;
    SA_SOLDEG       : THNumEdit;
    SA_SOLDET       : THNumEdit;
    SA_TOTALDEBIT   : THNumEdit;
    SA_TOTALCREDIT  : THNumEdit;
    SA_SOLDE        : THNumEdit;
    E_NUMEROPIECE   : THLabel;
    HLigne          : THMsgBox;
    HPiece          : THMsgBox;
    HTitres         : THMsgBox;
    HDiv            : THMsgBox;
    HTVA            : THMsgBox;
    POPS            : TPopupMenu;
    POPZ            : TPopupMenu;
    BZoom: THBitBtn;
    BSaisTaux: THBitBtn;
    BZoomJournal: THBitBtn;
    BZoomDevise: THBitBtn;
    BZoomEtabl: THBitBtn;
    BScenario: THBitBtn;
    BDernPieces: THBitBtn;
    BChancel: THBitBtn;
    BSwapPivot: THBitBtn;
    BModifRIB: THBitBtn;
    BInsert: THBitBtn;
    BSDel: THBitBtn;
    PForce          : TPanel;
    Label2          : TLabel;
    Label3          : TLabel;
    HForce          : THNumEdit;
    Image1          : TImage;
    FindSais        : TFindDialog;
    FlashGuide      : TFlashingLabel;
   PCreerGuide      : TPanel;
    H_TitreGuide    : TLabel;
    PFenGuide       : TPanel;
    BCValide: THBitBtn;
    BCAbandon: THBitBtn;
    BProrata: THBitBtn;
    BGuide: THBitBtn;
    BCreerGuide: THBitBtn;
    BMenuGuide: THBitBtn;
    BModifSerie: THBitBtn;
    CMontantsGuide  : TCheckBox;
    FNOMGUIDE       : TEdit;
    H_NOMGUIDE      : THLabel;
    Bevel1          : TBevel;
    Bevel2          : TBevel;
    Bevel3          : TBevel;
    Bevel4          : TBevel;
    HMTrad: THSystemMenu;
    EURO: TEdit;
    DockTop: TDock97;
    DockRight: TDock97;
    DockLeft: TDock97;
    DockBottom: TDock97;
    Valide97: TToolbar97;
    BValide: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BAide: TToolbarButton97;
    Outils: TToolbar97;
    BSolde: TToolbarButton97;
    BEche: TToolbarButton97;
    BVentil: TToolbarButton97;
    BGenereTva: TToolbarButton97;
    BControleTVA: TToolbarButton97;
    BComplement: TToolbarButton97;
    BChercher: TToolbarButton97;
    BMenuZoom: TToolbarButton97;
    BPrev: TToolbarButton97;
    BNext: TToolbarButton97;
    ToolbarSep971: TToolbarSep97;
    ToolbarSep972: TToolbarSep97;
    Sep97: TToolbarSep97;
    BModifTva: TToolbarButton97;
    LeTimer: TTimer;
    BCheque: THBitBtn;
    TVAImages: THImageList;
    BPopTva: TPopButton97;
    HConf: TToolbarButton97;
    ISigneEuro: TImage;
    LSA_SoldeG: THLabel;
    LSA_SoldeT: THLabel;
    LSA_Solde: THLabel;
    LSA_TotalCredit: THLabel;
    LSA_TotalDebit: THLabel;
    BChoixRegime: THBitBtn;
    BZoomImmo: THBitBtn;
    BZoomTP: THBitBtn;
    BMenuAction: THBitBtn;
    BModifs: THBitBtn;
    BLibAuto: THBitBtn;
    BVidePiece: THBitBtn;
    BPieceGC: THBitBtn;
    BScan: TToolbarButton97;
    compte: THEdit;
    bLignesGC: THBitBtn;
    E_DATECOMPTABLE: THCritMaskEdit;
    BSCANPDF: THBitBtn;
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure E_JOURNALChange(Sender: TObject);
    procedure E_DEVISEChange(Sender: TObject);
    procedure E_NATUREPIECEChange(Sender: TObject);
    procedure E_DATECOMPTABLEChange(Sender: TObject);
    procedure E_ETABLISSEMENTChange(Sender: TObject);
    procedure E_DATECOMPTABLEExit(Sender: TObject);
    procedure H_JOURNALDblClick(Sender: TObject);
    procedure H_ETABLISSEMENTDblClick(Sender: TObject);
    procedure GSSetEditText(Sender: TObject; ACol, ARow: Longint; const Value: string);
    procedure GSExit(Sender: TObject);
    procedure GSMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure GSDblClick(Sender: TObject);
    procedure GSCellExit(Sender: TObject; var ACol, ARow: Longint;  var Cancel: Boolean);
    procedure GSCellEnter(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
    procedure GSRowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
    procedure GSRowExit(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
    procedure GSEnter(Sender: TObject);
    procedure BAbandonClick(Sender: TObject);
    procedure BInsertClick(Sender: TObject);
    procedure BValideClick(Sender: TObject);
    procedure BSoldeClick(Sender: TObject);
    procedure BSDelClick(Sender: TObject);
    procedure BVentilClick(Sender: TObject);
    procedure BEcheClick(Sender: TObject);
    procedure BZoomClick(Sender: TObject);
    procedure BGenereTvaClick(Sender: TObject);
    procedure BControleTVAClick(Sender: TObject);
    procedure BProrataClick(Sender: TObject);
    procedure POPSPopup(Sender: TObject);
    procedure BModifRIBClick(Sender: TObject);
    procedure BSwapPivotClick(Sender: TObject);
    procedure BComplementClick(Sender: TObject);
    procedure BChercherClick(Sender: TObject);
    procedure BChequeClick(Sender: TObject);
    procedure FindSaisFind(Sender: TObject);
    procedure BModifTvaClick(Sender: TObject);
    procedure BGuideClick(Sender: TObject);
    procedure BLibAutoClick(Sender: TObject);
    procedure BChancelClick(Sender: TObject);
    procedure BDernPiecesClick(Sender: TObject);
    procedure BZoomDeviseClick(Sender: TObject);
    procedure BZoomEtablClick(Sender: TObject);
    procedure BCreerGuideClick(Sender: TObject);
    procedure BSaisTauxClick(Sender: TObject);
    procedure BCAbandonClick(Sender: TObject);
    procedure BCValideClick(Sender: TObject);
    procedure BZoomJournalClick(Sender: TObject);
    procedure BScenarioClick(Sender: TObject);
    procedure BModifSerieClick(Sender: TObject);
    procedure BPrevClick(Sender: TObject);
    procedure BNextClick(Sender: TObject);
    procedure GSKeyPress(Sender: TObject; var Key: Char);
    procedure BAideClick(Sender: TObject);
    procedure BMenuZoomMouseEnter(Sender: TObject);
    procedure BValideMouseEnter(Sender: TObject);
    procedure BValideMouseExit(Sender: TObject);
    procedure GereTagEnter(Sender: TObject);
    procedure GereTagExit(Sender: TObject);
    procedure LeTimerTimer(Sender: TObject);
    procedure PEnteteEnter(Sender: TObject);
    procedure BPopTvaChange(Sender: TObject);
    procedure GereAffSolde(Sender: TObject);
    procedure E_DATECOMPTABLEEnter(Sender: TObject);
    procedure BChoixRegimeClick(Sender: TObject);
    procedure BZoomImmoClick(Sender: TObject);
    procedure BZoomTPClick(Sender: TObject);
    procedure BVidePieceClick(Sender: TObject);
    procedure BPieceGCClick(Sender: TObject);
    procedure BLignesGCClick(Sender : TObject);
    procedure BScanClick(Sender: TObject);
    procedure BSCANPDFClick(Sender: TObject);
  private
    QListe                     : TQuery ;
    SAJAL                      : TSAJAL ;
    DEV                        : RDEVISE ;
    NumPieceInt,NbLigEcr,NbLigAna : Longint ;
    CpteAuto,OkScenario,OkBQE,OkJalEffet,OkLC,OkPremShow,ModifNext,NeedEdition,RegimeForce : boolean ;
    CurAuto,CurLig             : integer ;
    TS,TGUI,TDELECR,TDELANA,TECHEORIG : TList ;
    TSA : TList ;
    M                          : RMVT ;
//    TOBLumiere                 : TOB ;
    Scenario,Entete,ModifSerie : TOBM ;
    GX,GY                      : integer ;
    PieceModifiee,ModeCG,TresoLettre,PieceConf,ValideEcriture : boolean ;
    GeneRegTVA,GeneTVA,GeneTPF,GuideAnal,LastPaie,GestionRIB : String3 ;
    GeneSoumisTPF,CoherTVA,ChoixExige : boolean ;
    Revision,UnChange,Guide,DejaRentre,ToutDebit,ToutEncais  : Boolean ;
    GuideAEnregistrer          : Boolean ;
    CodeGuideFromSaisie        : String ;
    appeldatepourguide               : Boolean ;
    SorteTva                   : TSorteTva ;
    ExigeTva,ExigeEntete       : TExigeTva ;
 //   GeneTypeExo                : TTypeExo ;
    YATP                       : T_YATP ;
    ListeTP                    : HTStrings ;
    OkMessAnal,AutoCharged,OkSuite,UnSeulTiers,ModeLC,RentreDate : boolean ;
    ModeDevise,FindFirst,Volatile,OuiTvaLoc,GeneCharge : Boolean ;
    NbLOrig,NbLG,DecDev : integer ;
    ModeForce                  : tmSais ;
    NowFutur,DatePiece,LastEche,LastVal : TDateTime ;
    GeneTreso,LastTraiteCHQ    : String17 ;
    ModeConf                   : String[1] ;
    OldDevise,DernVentiltype,EtatSaisie,StCellCur : String ;
    OldTauxDev                 : Double ;
    MemoComp                   : HTStrings ;
    WMinX,WMinY                : Integer ;
    TabTvaEnc                  : Array[1..5] of Double ;
    ModLess : Boolean ;
    TOBNumChq : TOB ; // Modif Fiche 12017 // Enca Deca : Gestion des numéro de traite
    bReedition : Boolean;
    szModeleBordereau          : String; // 12339
    FEcrCompl                  : TOB ;
    FInfo                      : TInfoEcriture ;
    Sanslettrage               : Boolean ;
    FBoInsertSpecif            : Boolean ;    // SBO 01/07/2007 : enregistrement spécifique pour pb CWAS
    FNeFermePlus               : Boolean; {JP 05/07/07 : FQ 19022 : Si on répond non à la confirmation fermeture}
// TVA, TPF
    procedure AlimHTByTVA ( Lig : integer ; RegTVA,TVAENC : String3 ; SoumisTpf,AChat : boolean ) ;
    function  InsereLigneTVA ( LigHT : integer ; RegTVA : String3 ; SoumisTpf,Achat : boolean ) : boolean ;
    procedure RemplirEcheAuto ( Lig : integer ) ;
    procedure RemplirAnalAuto ( Lig : integer ) ;
    procedure RegroupeTVA ;
    procedure GetSorteTVA ;
    procedure AfficheExigeTVA ;
    procedure GereTvaEncNonFact ( Lig : integer ) ;
    procedure GereLesImmos ( Lig : integer ) ;
    procedure TvaTauxDirecteur ( Lig : integer ; Client,ToutDebit : Boolean ; Regime : String3 ) ;
    procedure TabTvaDansModR ( OMODR : TMOD ) ;
// Scénario de saisie
    procedure ChargeScenario ;
    procedure LettrageEnSaisie ;
// Init et defaut
    Function  Recharge(QM : RMVT ) : Boolean ;
    procedure LectureSeule ;
    procedure RenseignePopTva ;
    procedure GereANouveau ;
    procedure InitEnteteJal(Totale,Zoom,ReInit : boolean) ;
    procedure InitModifJal(Action : TActionFiche) ;
    procedure InitGrid ;
    procedure DefautEntete ;
    procedure DefautPied ;
    procedure InsereLigne(Lig : integer) ;
    procedure AttribNumeroDef ;
    procedure AttribLeTitre ;
    procedure ReInitPiece ( Force : boolean ) ;
    procedure GereDeviseEcart ;
    procedure GereEnabled ( Lig : integer ) ;
    procedure GereOptionsGrid ( Lig : integer ) ;
    procedure PosLesSoldes ;
// Chargements
    procedure ChargeLignes ;
    procedure ChargeEcritures ;
    procedure ChargeComptes ;
    procedure ChargeSoldesComptes ;
    procedure ChargeAnals ;
    procedure RuptureEche( Lig : integer ; TDD,TCD,TDP,TCP : Double ) ;
    procedure AlimNoyauEche ( Lig : integer ; Premier : boolean ; LDD,LCD,LDP,LCP : Double ; QEcr : TDataSet ) ;
// Allocation et Désallocation
    function  DateToJour ( DD : TDateTime ) : String ;
    procedure AlloueEcheAnal(C,L : LongInt) ;
    procedure AlloueEche(LaLig : LongInt) ;
    procedure DesalloueEche(LaLig : LongInt) ;
    procedure AlloueAnal(LaLig : LongInt) ;
    procedure DesalloueAnal(LaLig : LongInt); overload;
    procedure DesalloueAnal(LaLig : LongInt; Ventilable : T5B); overload;
// Object Mouvement, Divers
    procedure AlimObjetMvt(Lig : integer ; AvecM : boolean ) ;
    procedure GereComplements(Lig : integer) ;
    function  FermerSaisie : boolean ;
    Procedure EditionSaisie ;
    procedure DecrementeNumero ;
    procedure GotoEntete ;
    procedure NumeroteVentils ;
    procedure StockeLaPiece ;
    Function  FabricFromM : RMVT ;
    procedure QuelNExo ;
    Procedure GetCptsContreP ( Lig : Integer ; Var LeGene,LeAux,TL0 : String17 ) ;
    Function  TraiteDoublon ( Lig : integer ) : Boolean ;
    Procedure FocusSurNatP ;
// Tiers payeur
    Procedure YaBienTP ( QEcr : TDataSet ) ;
    Function  ExisteTP : boolean ;
    Procedure TripoteYATP ;
    Procedure GereTiersPayeur ;
    Procedure DetruitOldPayeurs ;
    Procedure DetruitPiecesTP ;
// Automates de ref et lib
    procedure ChargeLibelleAuto ;
    procedure GereRefLib(Lig : integer) ;
    procedure RenseigneLibelleAuto (Col,Lig : integer ; Ref : boolean ) ;
    Function  CommeUneBQE ( Lig : integer ) : Boolean ;
    Function  IncNumLotSaisie ( Lig : integer ) : String ;
    Procedure RemonteValeurTraCHQ ( sTraCHQ : String ; Lig : integer ) ;
// Guides de saisie
    function  ProchainArret ( Col,Lig : integer ) : integer ;
    function  ZoneObli ( Col,Lig : integer ; St : String ) : Boolean ;
    procedure SuiteLigneColonne ( Var Col,Lig : longint ) ;
    procedure ValideCeQuePeux( Lig : longint ) ;
    procedure EtudieSiGuide(ACol,ARow : longint) ;
    procedure AlimGuideLettrage ( Lig : integer ; Q : TQuery ) ;
    procedure AlimGuideTVA ( Lig : integer ; Q : TQuery ) ;
    procedure AlimGuideVentil ( Lig : integer ) ;
    procedure RecalcGuide ;
    procedure TraiteLigneAZero ;
    procedure ResaisirLaDatePourLeGuide(OkOk : Boolean) ;
    procedure FinGuide ;
    procedure LanceGuide ( TypeG,CodeG : String3 ) ;
    function  Load_Sais ( St : hString ) : Variant ;
    function  Load_TVATPF ( CGen : TGGeneral ; OKTVA : boolean ; TVA : String ) : Variant ;
// Création du guide en saisie
    procedure CreerLignesGuide ( CodeG : String ) ;
    procedure CreerEnteteGuide ( CodeG : String ) ;
    function  BonRegle ( Lig,NbEche : integer ; LeMode : String ; T : T_TabEche ) : Boolean ;
// Lettres chèque, Encadeca
    Function  ChoixFromTenueSaisie ( LeMontantPIVOT,LeMontantEURO,LeMontantDEV : Double ) : Double ;
    Function  EcheancesLC ( NbC : integer ) : boolean ;

    // Modif fiche 12017 : pb maj E_NUMTRAITECHQ
    procedure SetNumTraiteDOCREGLE ;
    procedure StockNumChqDepuisEche   ( LL : TList ) ;
    // Fin modif 12017

    procedure MajMPSansLettrage ;

    procedure SuiteEncaDeca ;
    procedure NormaliserEscomptes ;
    procedure LettrerReglement ;
    procedure EnvoiCFONB ;
    procedure EnvoiBordereau ;
    procedure CloseSaisie ;
    procedure CreateSaisie ;
    procedure TraiteFormuleMP(St : String ; O : TOBM ; Lig,Col : Integer ; FormuleEnDur : String = '') ;
    procedure AffecteRef(FormuleOrigine : Boolean) ;
    procedure ShowSaisie ;
    procedure TraiteAffichePCL ;
    procedure TraiteAfficheGescom ;
// Barre Outils
    procedure ClickDel ;
    procedure ClickInsert ;
    procedure ClickSolde(Egal,RowEx : boolean) ;
    procedure ClickVentil ;
    procedure ClickEche ;
    procedure ClickZoom(DoPopup : boolean = false);
    procedure ClickConsult ;
    procedure ClickAbandon ;
    procedure ClickGenereTVA ;
    function  ClickControleTVA ( Bloque : boolean ; LigHT : integer ) : boolean ;
    procedure ClickProrata ;
    procedure ClickSwapPivot ;
    procedure ClickModifRIB ;
    procedure ClickComplement ;
    procedure ClickCherche ;
    procedure ClickCheque ;
    procedure ClickGuide ;
    procedure ClickVidePiece ( Parle : boolean ) ;
    procedure ClickLibelleAuto ;
    procedure ClickChancel ;
    procedure ClickDernPieces ;
    procedure ClickCreerGuide ;
    procedure ClickEtabl ;
    procedure ClickDevise ;
    procedure ClickSaisTaux ;
    procedure ClickChoixRegime ;
    procedure ClickScenario ;
    procedure ClickJournal ;
    procedure ClickPieceGC(FromPce : boolean=true);
    procedure ClickImmo ;
    procedure ClickTP ;
    procedure ClickModifTva ;
    procedure ForcerMode(Cons : boolean ; Key : word) ;
// Calculs lignes
    procedure InfoLigne( Lig : integer ; Var D,C,TD,TC : Double ) ;
    procedure DetruitLigne( Lig : integer ; Totale : boolean) ;
    procedure DetruitLigneGuide(Lig : integer) ;
    procedure CalculDebitCredit ;
    procedure AfficheConf ;
    procedure CalculSoldeCompte ( CpteG,CpteA : String ; DIG,CIG,DIA,CIA : Double ) ;
    procedure SommeSoldeCompte ( Col : integer ; Cpte : String ; Var TD,TC : Double ; Old,OkDev : Boolean) ;
    procedure DefautLigne ( Lig : integer ; Init : boolean ) ;
    procedure PutScenar ( O : TOBM ; Lig : integer ) ;
    procedure TraiteMontant( Lig : integer ; Calc : boolean ) ;
    procedure CaptureRatio( Lig : integer ; Force : boolean ) ;
// Analytiques
    Procedure OuvreAnal ( Lig : integer ; Scena : boolean ; vNumAxe : integer = 0 ) ;
    Function  AOuvrirAnal ( Lig : integer ; Auto : boolean ) : Boolean ;
    Procedure GereAnal ( Lig : integer ; Auto,Scena : boolean ) ;
    Procedure RecupAnal(Lig : integer ; Var OBA : TOB ; NumAxe,NumV : integer ) ;
// Echéances
    Function  AOuvrirEche ( Lig : integer ; Var Cpte : String17 ; Var MR : String3 ;
                            Var OuvreAuto,RempliAuto : boolean ; Var t : TLettre ) : Boolean ;
    Procedure OuvreEche ( Lig : integer ; Cpte : String17 ; MR : String3 ; RempliAuto : boolean ; t : TLettre ) ;
    Procedure GereEche ( Lig : integer ; OuvreAuto,RempliAuto : boolean ) ;
    Procedure RecupEche( Lig : integer ; R : T_ModeRegl ; NumEche : integer; var OBM : TOBM ; vBoMonoEche : boolean = False ) ;
    Procedure EcheGuideRegle ( Lig : integer ; Var T : T_MODEREGL ) ;
    Procedure EcheLettreCheque ( Lig : integer ; Var T : T_MODEREGL ) ;
    Procedure GereMonoEche ( Var OO : T_ModeRegl ; OkInit,AutoPourEncaDeca : boolean ; Lig : integer ; ActionEche : TActionFiche ) ;
// Contrôles
    function  DateEchesOk ( Lig : integer ) : boolean ;
    Function  LigneCorrecte ( Lig : integer ; Totale,Alim : boolean ) : boolean ;
    Function  PieceCorrecte : boolean ;
    Function  EquilEuroValide : Boolean ;
    Function  ControleRIB : boolean ;
    Function  PasModif : boolean ;
    Function  Equilibre : boolean ;
    procedure ErreurSaisie ( Err : integer ) ;
    procedure AjusteLigne ( Lig : integer ; AvecM : Boolean ) ;
    Function  PossibleRecupNum : Boolean ;
    procedure ControleLignes ;
    function  ControleTVAOK : boolean ;
    function  PasToucheLigne ( Lig : integer ) : boolean ;
// Appels comptes
    Function  ChercheGen(C,L : integer ; Force : boolean ; DoPopup : boolean = false) : byte;
    Function  ChercheAux(C,L : integer ; Force : boolean ; DoPopup : boolean = false) : byte;
    procedure ChercheJAL ( Jal : String3 ; Zoom : boolean ) ;
    procedure ChercheMontant(Acol,Arow : longint) ;
    procedure AppelAuto ( indice : integer ) ;
    Procedure AffecteLeRib(O : TOBM ; Cpt : String ; ForceRAZ : Boolean = FALSE) ;
    Function  AffecteRIB ( C,L : integer ; IsAux : Boolean ; ForceRAZ : Boolean = FALSE) : Boolean ;
    Procedure AffecteConso ( C,L : integer ) ;
    Procedure AffecteTva ( C,L : integer ) ;
    Function  SortePiece : boolean ;
// MAJ Fichier
    Procedure GetEcr(Lig : Integer) ;
    Procedure GetEcrGrid(Lig : Integer) ;
    Procedure GetAna(Lig : Integer) ;
    Procedure RecupTronc(Lig : Integer ; MO : TMOD; Var OBM : TOBM );
    procedure RecupFromGrid( Lig : integer ; MO : TMOD ; NumE : integer ) ;
    Procedure ClickValide ;
    Procedure DetruitPieceMP ;
    procedure TraiteMP ;
    procedure TraiteOrigMP ;
    procedure ValideLaPiece ;
    procedure ValideLeReste ;
    procedure ValideLesAno ;
    Procedure ValideLignesGene ;
    Procedure ValideLesComptes ;
    procedure ValideLeJournalNew ;
    Procedure AttribRegimeEtTVA ;
    Procedure RenseigneRegime ( Lig : integer ; Recharge : boolean ) ;
    procedure IntoucheGene ;
    procedure IntoucheAux ;
    procedure MajCptesGeneNew ;
    procedure MajCptesAuxNew  ;
 //   procedure MajCptesSectionNew ;
    procedure DetruitAncien ;
    procedure InverseSoldeNew ;
    Procedure EcartChange(Obj : TOB) ;
    function  QuelEnc ( Lig : Integer ) : String3 ;
// Affichages, Positionnements
    Procedure AutoSuivant( Suiv : boolean ) ;
    Procedure SoldelaLigne ( Lig : integer ) ;
    procedure AffichePied ;
    procedure PositionneDevise ( ReInit : boolean ) ;
    procedure GereZoom ;
    procedure AvertirPbTaux(Code : String3 ; DateTaux,DateCpt : TDateTime) ;
    Function  PbTaux ( DEV : RDevise ; DateCpt : TDateTime ) : boolean ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
// Conversions, Caluls
    Function  ArrS( X : Double ) : Double ;
    Function  DateMvt ( Lig : integer ) : TDateTime ;
   {$IFDEF SCANGED}
    procedure SetGuidId ( vGuidId : string ) ;
    function  RechGuidId : string ;
    procedure AjouteGuidId( vGuidId : string ) ;
    procedure SupprimeLeDocGuid ;
    {$ENDIF}
    procedure ChargeCompl ;
    procedure GereCutoff(Lig : integer) ;
    procedure SupprimeEcrCompl ( vTOB : TOB ) ;
    procedure PostDrawCell(ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
    function  GereAcc (Ou : integer = - 1) : integer ;
    function  GetCompteAcc( Ou : integer ; E_AUXILIAIRE : string ) : string ;
    procedure SetTypeExo ;   // FQ 13246 : SBO 30/03/2005
    {JP 29/07/05 : FQ 16032 : Mise à jour de TEcheOrig avec les écritures qui ont réellement été traitées}
    procedure MajTEcheOrig;
    // DEV 3216 10/04/2006 SBO : récupération des fonctionnalités de saisBor.pas
    function  GetFormulePourCalc(Formule: hstring): Variant;
    function  GetFormatPourCalc : string ;
    Function YAFormule(St : string) : Boolean ;
    function TestQteAna : boolean ;
  public
    Action,OldAction : TActionFiche ;
    TPIECE : TList ;
    SI_TotDS,SI_TotCS,SI_TotDP,SI_TotCP,SI_TotDD,SI_TotCD : double ;
    SI_FormuleRef,SI_FormuleLib : String ;
    CEstGood,SuiviMP,GenereAuto : boolean ;
    LesM               : TList ;
    CurM               : integer ;
    TOBMPOrig,TOBMPEsc,TOBParamEsc : TOB ;
    DateSais           : TDateTime ;
   {$IFDEF SCANGED}
    function GetInfoLigne : string ;
    property GuidID : string  write SetGuidID ;
    {$ENDIF}
  protected
    FClosing: boolean ;
    procedure Fermeture ;
    function  QuelleSaisieCascade : string ;
  end;

implementation

{$R *.DFM}

Uses
 {$IFDEF MODENT1}
 CPProcGen,
 CPProcMetier,
 CPVersion,
 {$ENDIF MODENT1}

{$IFDEF EAGLCLIENT}
	UtileAGL,
{$ELSE}
  {$IFNDEF IMP}
   {$IFNDEF GCGC}
    EdtREtat,
   {$ENDIF}
  {$ENDIF}
{$ENDIF}
  {$IFNDEF IMP}
   {$IFNDEF GCGC}
    {$IFNDEF CCS3}
      {$IFNDEF VEGA}
      CFONB,			// pour ExportCFONB
     {$ENDIF}
    {$ENDIF}
   {$ENDIF}
  {$ENDIF}
   LetBatch,
   SaisBor,
   GuiDate,
{$IFNDEF CMPGIS35}
   UTofConsEcr,
{$ENDIF}
  {$IFNDEF VEGA}
{$IFNDEF CMPGIS35}
  MulSmpUtil,
{$ENDIF}
  {$ENDIF}
  Constantes, {Maj E_TRESOSYNCHRO}
  {$IFDEF COMPTA}
  UProcGen, {RPos, StrLeft}
  {$ENDIF COMPTA}
  {$IFDEF SCANGED}
  cbpPath,
  {$ENDIF}
   CPGeneraux_TOM, // pour Zoom sur fiche Gene
   CPTiers_TOM,     // pour Zoom sur fiche tiers
   GuidUtil        // N°23437 : appel CCalculeCodeGuide
   ;

procedure SaisieGuidee ( StG : String3) ;
Var R : RMVT ;
    FromMenu : Boolean ;
    Jal,Nat,Devi,Etab,OldCode : String3 ;
    Q                 : TQuery ;
BEGIN
{$IFNDEF GCGC}
FromMenu:=False ; Jal:='' ; Nat:='' ; Devi:='' ; Etab:='' ; OldCode:='' ;
Repeat
  if StG='' then BEGIN
    FromMenu:=True ;
    // StG := SelectGuide(Jal,Nat,Devi,Etab,(V_PGI.DateEntree>=V_PGI.DateDebutEuro),OldCode) ;
  END ;
  if StG='' then Exit ;
  FillChar(R,Sizeof(R),#0) ; R.DateC:=V_PGI.DateEntree ; R.Exo:=QuelExoDT(R.DateC) ; R.Simul:='N' ;
  R.LeGuide:=StG ; R.TypeGuide:='NOR' ; R.ANouveau:=False ; R.SaisieGuidee:=True ;
  LanceSaisie(Nil,taCreat,R) ;
  Screen.Cursor:=SyncrDefault ;
  if FromMenu then
     BEGIN
     Q:=OpenSQL('Select GU_JOURNAL, GU_NATUREPIECE, GU_DEVISE, GU_ETABLISSEMENT from GUIDE Where GU_TYPE="NOR" AND GU_GUIDE="'+StG+'"',True) ;
     if Not Q.EOF then
        BEGIN
        Jal:=Q.Fields[0].AsString  ; Nat:=Q.Fields[1].AsString ;
        Devi:=Q.Fields[2].AsString ; Etab:=Q.Fields[3].AsString ;
        OldCode:=StG ; StG:='' ;
        END ;
     Ferme(Q) ;
     END ;
Until Not FromMenu ;
{$ENDIF}
END ;

Procedure LanceMultiSaisie ( LesM : TList ) ;
Var X : TFSaisie ;
    M : RMVT ;
BEGIN
  M := P_MV(LesM[0]).R ;
  X := TFSaisie.Create(Application) ;
 Try
  X.Action:=taCreat ; X.OldAction:=taCreat ; X.M:=M ;
  X.LesM:=LesM ; X.CurM:=0 ; X.SuiviMP:=False ;
  X.ShowModal ;
 Finally
  X.Free ;
 End ;
Screen.Cursor:=SyncrDefault ;
END ;

Function LanceSaisie ( QListe : TQuery ; Action : TActionFiche ; Var M : RMVT  ; ModLess : Boolean = FALSE) : boolean ;
Var X  : TFSaisie ;
    OA : TActionFiche ;
    PP : THPanel ;
    DateSais : TDateTime ;
BEGIN
{OA:=Action ;} Result:=False ; M.LastNumCreat:=-1 ;
DateSais:=V_PGI.DateEntree ;
(* PFU : DEBUT AJOUT 03/07/2000 *)
if (Action=taModif) and (SaisieLancer) then Action:=taConsult ;
(* PFU : FIN AJOUT 03/07/2000 *)
{Gestion exo de référence}
if ((Action=taCreat) and (VH^.CPExoRef.Code<>'') and (VH^.CPLastSaisie>0) and
    (Not M.SaisieGuidee) and (M.TypeGuide='') and (Not M.ANouveau) and
    (VH^.CPLastSaisie>=VH^.CPExoRef.Deb) and (VH^.CPLastSaisie<=VH^.CPExoRef.Fin) and
    (ctxPCL in V_PGI.PGIContexte)) then DateSais:=VH^.CPLastSaisie ;
Case Action of
   taCreat : BEGIN
             if PasCreerDate(DateSais) then Exit ;
             {$IFNDEF TT}
             if DepasseLimiteDemo then Exit ;
             {$ENDIF}
             if _Blocage(['nrCloture'],True,'nrSaisieCreat') then Exit ;
             END ;
   taModif : BEGIN
             if RevisionActive(M.DateC) then Exit ;
             if _Blocage(['nrCloture','nrBatch'],True,'nrSaisieModif') then Exit ;
             END ;
   END ;
PP:=FindInsidePanel ;
X:=TFSaisie.Create(Application) ;
X.QListe:=QListe ; X.Action:=Action ; X.OldAction:=Action ;
OA:=Action ; X.M:=M ; X.SuiviMP:=False ; X.DateSais:=DateSais ;
X.TOBMPOrig:=Nil ; X.TOBMPEsc:=Nil ;
X.TOBParamEsc := nil ;
X.ModLess:=ModLess ;
If ModLess Then
  BEGIN
  X.Show ;
  END Else
if PP=Nil then
   BEGIN
    Try
     X.ShowModal ;
     Result:=X.CEstGood ;
     M.LastNumCreat:=X.M.LastNumCreat ;
     M.MSED.SoucheSpooler:=X.M.MSED.SoucheSpooler ;
     M.MSED.ModeleMultiSession:=X.M.MSED.ModeleMultiSession ;
    Finally
     X.Free ;
     Case OA of
        taCreat : _Bloqueur('nrSaisieCreat',False) ;
        taModif : _Bloqueur('nrSaisieModif',False) ;
        END ;
    End ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(X,PP) ;
   X.Show ;
   END ;
END ;

Function LanceSaisieMP ( M : RMVT ; TOBMPOrig,TOBMPEsc : TOB ; GenereAuto : Boolean ; LettrageMP : Boolean ; TOBParamEsc : TOB = nil; Reedition : Boolean = False; ModeleBordereau : String = '') : boolean ;
Var X  : TFSaisie ;
    PP : THPanel ;
BEGIN
Result:=False ; M.LastNumCreat:=-1 ;
if RevisionActive(M.DateC) then Exit ;
PP:=FindInsidePanel ;
X:=TFSaisie.Create(Application) ;
X.QListe:=Nil ; X.Action:=taModif ; X.OldAction:=taModif ; //OA:=taModif ;
X.M:=M ; X.SuiviMP:=True ; X.GenereAuto:=GenereAuto ;
X.TOBMPOrig:=TOBMPOrig ;
X.TOBMPEsc:=TOBMPEsc ;
X.bReedition := Reedition;
X.szModeleBordereau := ModeleBordereau;
// Prise en compte des paramètres d'escomptes par lignes
X.TOBParamEsc := TOBParamEsc ;
//Récupération du choix de lettrer ou pas pour les pièces en devises
X.Sanslettrage:= not LettrageMP;
If X.SuiviMP And (X.M.Smp in [smpEncTraEdtNC,smpDecBorEdtNC,smpDecVirEdtNC,smpDecVirInEdtNC]) Then
  BEGIN
  X.BorderIcons:=X.BorderIcons+[biMinimize] ; X.WindowState:=wsMinimized ;
  END ;
if PP=Nil then
   BEGIN
    Try
     X.ShowModal ;
     Result:=X.CEstGood ;
     M.LastNumCreat:=X.M.LastNumCreat ;
    Finally
     X.Free ;
    End ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(X,PP) ;
   X.Show ;
   END ;

END ;

{$IFDEF SCANGED}
procedure SaisieMyAfterImport (Sender: TObject; FileGuid: string; var Cancel: Boolean) ;
var LastError : String;
    lTobDoc, lTobDocGed : TOB;
    lStDocId : string;
begin

 if  (ctxPCL in V_PGI.PGIContexte) and not JaileDroitConceptBureau (187315) then exit ;

 lTobDoc := Tob.Create('YDOCUMENTS', nil, -1) ;
 lTobDoc.LoadDb ; // Pas de clé = Charge une clé à 0
 lTobDocGed := Tob.Create('DPDOCUMENT', nil, -1) ;
 lTobDocGed.LoadDb ;

 try

  // TobDoc :    principaux champs à renseigner
  //             DOCID, ANNEE, MOIS, DATEDEB, DATEFIN, LIBELLEDOC,
  //             NATDOCUMENT, AUTEUR, MOTSCLES, BLOCNOTE
  lTobDoc.PutValue('YDO_LIBELLEDOC'    , TFSaisie(Sender).GetInfoLigne) ; // PAR EXEMPLE !!
  lTobDoc.PutValue('YDO_NATDOC'        , 'Documents scanné') ; // Libre !
  lTobDoc.PutValue('YDO_MOTSCLES'      , 'ECRITURE') ; // En majuscules, séparés par des ;
  lTobDoc.PutValue('YDO_ANNEE'         , FormatDateTime('yyyy', Date)) ;

  // TobDocGed : NODOSSIER
  lTobDocGed.PutValue('DPD_NODOSSIER', V_PGI.NoDossier) ;

  // Insertion effective du document avec update sur les TOB
  lStDocId := InsertDocumentGed(lTobDoc, lTobDocGed, FileGuid, LastError) ;

 finally
  lTobDoc.Free ;
  lTobDocGed.Free ;
 end ;

 // Fichier refusé, suppression dans la GED
 if lStDocId = '' then // consulter éventuellement LastError
  V_GedFiles.Erase(FileGuid)
   else
    TFSaisie(Sender).GuidID := lStDocId ; // pour traitement par l'intéressé

end;


procedure AjouterFichierDansGed ( Sender : TObject ) ;
var
 lBoCancel       : boolean ;
 lStFileGuid     : string ;
 lDialog         : TOpendialog ;
 lStFichier      : string ;
begin

 lDialog            := TOpendialog.Create(nil) ;

 try

 lDialog.InitialDir := TcbpPath.GetMyDocuments ;
 lDialog.Options    := [ofFileMustExist, ofHideReadOnly];
 lDialog.Filter     := 'Fichiers pdf (*.pdf)|*.pdf|';

 if lDialog.Execute then
  begin
   lStFichier := lDialog.FileName ;
   lStFileGuid := V_GedFiles.Import (lStFichier);
  end ;

 finally
  lDialog.Free ;
 end ;

 SaisieMyAfterImport(Sender,lStFileGuid,lBoCancel) ;

end ;

{$ENDIF}

{==========================================================================================}
{================================= Libellés auto ==========================================}
{==========================================================================================}
procedure TFSaisie.RenseigneLibelleAuto (Col,Lig : integer ; Ref : Boolean );
BEGIN
if GS.Cells[Col,Lig]<>'' then Exit ; if M.ANouveau then Exit ;
if Ref then GS.Cells[Col,Lig]:=GFormule(SI_FormuleRef,Load_Sais,Nil,1)
       else GS.Cells[Col,Lig]:=GFormule(SI_FormuleLib,Load_Sais,Nil,1) ;
END ;

procedure TFSaisie.GereRefLib(Lig : integer) ;
BEGIN
if PasModif then Exit ; if Guide then Exit ; if M.ANouveau then Exit ;
RenseigneLibelleAuto(SA_RefI,Lig,True) ;
RenseigneLibelleAuto(SA_Lib,Lig,False) ;
END ;

procedure TFSaisie.ChargeLibelleAuto ;
Var Q : TQuery ;
BEGIN
if E_JOURNAL.Value='' then Exit ; if E_NATUREPIECE.Value='' then Exit ;
if PasModif then Exit ;
Q:=OpenSQL('Select RA_FORMULEREF, RA_FORMULELIB from REFAUTO Where RA_JOURNAL="'+E_JOURNAL.Value+'" AND RA_NATUREPIECE="'+E_NATUREPIECE.Value+'"',True) ;
if Not Q.EOF then
   BEGIN
   SI_FormuleRef:=Q.FindField('RA_FORMULEREF').AsString ;
   SI_FormuleLib:=Q.FindField('RA_FORMULELIB').AsString ;
   AutoCharged:=True ; Q.Next ;
   // N° 10129 le 13/06/2002
   If Not Q.Eof Then
     BEGIN
     SI_FormuleRef:='' ;
     SI_FormuleLib:='' ;
     AutoCharged:=FALSE ;
     END ;
   END ;
Ferme(Q) ;
END ;

{==========================================================================================}
{================================== Chargements ===========================================}
{==========================================================================================}
procedure TFSaisie.AlimNoyauEche ( Lig : integer ; Premier : boolean ; LDD,LCD,LDP,LCP : Double ; QEcr : TDataSet ) ;
Var TM : TMOD ;
    k  : integer ;
BEGIN
TM:=TMOD(GS.Objects[SA_NumP,Lig]) ; if TM=Nil then Exit ;
With TM.MODR do
     BEGIN
     if Premier then
        BEGIN
        Action:=taModif ; NbEche:=0 ; CodeDevise:=DEV.Code ; Symbole:=DEV.Symbole ;
        TauxDevise:=DEV.Taux ; Quotite:=DEV.Quotite ; Decimale:=DecDev ;
        DateFact:=StrToDate(E_DATECOMPTABLE.Text) ; DateBl:=DateFact ; DateFactExt:=DateFact ;
// GP REGL
        DateFactExt:=QEcr.FindField('E_DATEREFEXTERNE').AsDateTime ;
        If DateFactExt=IDate1900 Then DateFactExt:=DateFact ;
        Aux:=GS.Cells[SA_Aux,Lig] ; if Aux='' then Aux:=GS.Cells[SA_Gen,Lig] ;
        TotalAPayerP:=0 ; TotalAPayerD:=0 ;
        END ;
     TotalAPayerP:=TotalAPayerP+LDP+LCP ;
     TotalAPayerD:=TotalAPayerD+LDD+LCD ;
     Inc(NbEche) ;
     With TabEche[NbEche] do
        BEGIN
        MontantD:=LDD+LCD ; MontantP:=LDP+LCP ;
        ReadOnly:=(Trim(QEcr.FindField('E_LETTRAGE').AsString)<>'') ;
        CodeLettre:=QEcr.FindField('E_LETTRAGE').AsString ;
        ModePaie:=QEcr.FindField('E_MODEPAIE').AsString ;
        DateEche:=QEcr.FindField('E_DATEECHEANCE').AsDateTime ;
        DateValeur:=QEcr.FindField('E_DATEVALEUR').AsDateTime ;
        DateRelance:=QEcr.FindField('E_DATERELANCE').AsDateTime ; NiveauRelance:=QEcr.FindField('E_NIVEAURELANCE').AsInteger ;
        Couverture:=QEcr.FindField('E_COUVERTURE').AsFloat ;
        CouvertureDev:=QEcr.FindField('E_COUVERTUREDEV').AsFloat ;
        LettrageDev:=QEcr.FindField('E_LETTRAGEDEV').AsString ;
        DatePaquetMax:=QEcr.FindField('E_DATEPAQUETMAX').AsDateTime ;
        DatePaquetMin:=QEcr.FindField('E_DATEPAQUETMIN').AsDateTime ;
        EtatLettrage:=QEcr.FindField('E_ETATLETTRAGE').AsString ;
        {#TVAENC}
        if OuiTvaLoc then
           BEGIN
           for k:=1 to 4 do
               BEGIN
               TAV[k]:=QEcr.FindField('E_ECHEENC'+IntToStr(k)).AsFloat ;
               if TAV[k]<>0 then ToutDebit:=False ;
               END ;
           TAV[5]:=QEcr.FindField('E_ECHEDEBIT').AsFloat ;
           if TAV[5]<>0 then ToutEncais:=False ;
           END ;
        END ;
     END ;
END ;

procedure TFSaisie.RuptureEche( Lig : integer ; TDD,TCD,TDP,TCP : Double ) ;
Var OBM : TOBM ;
    TM  : TMOD ;
    i   : integer ;
BEGIN
 GS.Cells[SA_Debit,Lig]:=StrS(TDD,DECDEV) ; GS.Cells[SA_Credit,Lig]:=StrS(TCD,DECDEV) ;
FormatMontant(GS,SA_Debit,Lig,DECDEV) ; FormatMontant(GS,SA_Credit,Lig,DECDEV) ;
OBM:=GetO(GS,Lig) ;
if OBM<>Nil then
   BEGIN
   OBM.PutMvt('E_DEBIT',TDP)     ; OBM.PutMvt('E_CREDIT',TCP) ;
   OBM.PutMvt('E_DEBITDEV',TDD)  ; OBM.PutMvt('E_CREDITDEV',TCD) ;
   OBM.HistoMontants ;
   END ;
TM:=TMOD(GS.Objects[SA_NumP,Lig]) ;
if TM<>Nil then if TM.ModR.TotalAPayerP<>0 then
   BEGIN
   for i:=1 to TM.ModR.NbEche do TM.ModR.TabEche[i].Pourc:=Arrondi(100.0*TM.ModR.TabEche[i].MontantP/TM.ModR.TotalAPayerP,ADecimP) ;
   END ;
END ;

procedure TFSaisie.ChargeComptes ;
Var X_Query : TQuery ;
    i       : integer ;
    CGen    : TGGeneral ;
    CAux    : TGTiers ;
    LigC    : integer ;
    stReq		: String;
BEGIN
  if Action=taCreat then Exit ;
  InitMove(2*GS.RowCount,'') ;
  //BeginTrans ;
  // Comptes généraux
  for i:=1 to GS.RowCount-2 do
  	if GS.Cells[SA_Gen,i]<>'' then
      BEGIN
      if MoveCur(FALSE) then ;
      LigC := TrouveLigCompte(GS,SA_Gen,0,GS.Cells[SA_Gen,i]) ;
      CGen := Nil ;
      if LigC>0 then AttribGen(GS,LigC,i) else
        BEGIN
        stReq := SelectGene + ' FROM GENERAUX WHERE G_GENERAL="' + GS.Cells[SA_Gen,i] + '"' ;
        X_Query := OpenSQL(stReq, True) ;
        if Not X_Query.EOF then
          BEGIN
          CGen:=TGGeneral.Create('') ;
          CGen.QueryToGeneral(X_Query) ;
          END ;
        Ferme(X_Query);
        GS.Objects[SA_Gen,i]:=TObject(CGen) ;
        END ;
      AlloueEcheAnal(SA_Gen,i) ;
      if ((GeneRegTva='') and (isTiers(GS,i))) then AffecteTva(SA_Gen,i) ;
      END ;

  // Auxiliaires
  for i:=1 to GS.RowCount-2 do
  	if GS.Cells[SA_Aux,i]<>'' then
      BEGIN
      if MoveCur(FALSE) then ;
      LigC:=TrouveLigCompte(GS,SA_Aux,0,GS.Cells[SA_Aux,i]) ;
      CAux:=Nil ;
      if LigC>0 then AttribAux(GS,LigC,i)
      	else
          BEGIN
          stReq := SelectAuxi + ' FROM TIERS WHERE T_AUXILIAIRE="' + GS.Cells[SA_Aux,i] + '"';
          X_Query := OpenSQL(stReq, True) ;
          if Not X_Query.EOF then
            BEGIN
            CAux:=TGTiers.Create('') ;
            CAux.QueryToTiers(X_Query) ;
            END ;
          Ferme(X_Query);
          GS.Objects[SA_Aux,i]:=TObject(CAux) ;
          END ;
      AlloueEcheAnal(SA_Aux,i) ;
      if ((GeneRegTva='') and (isTiers(GS,i))) then AffecteTva(SA_Aux,i) ;
      END ;

  // TInfoECriture   // FQ 13246 : SBO 30/03/2005
  for i := 1 to GS.RowCount - 2 do
    FInfo.Load( GS.Cells[SA_Gen,i], GS.Cells[SA_Aux,i], '' ) ;

  // Compte de Trésorerie FQ 15190 SBO 25/03/2005
  if ( SAJAL<>Nil ) then
    if (OkBQE or OkJALEFFET) then GeneTreso := SAJAL.TRESO.Cpt ;

  FiniMove ;
END ;

procedure TFSaisie.ChargeEcritures ;
Var Lig  : integer ;
    NumE,NbEche : integer ;
    OkE         : Boolean ;
    OBM         : TOBM ;
    TDD,TCD,TDP,TCP : Double ;
    LDD,LCD,LDP,LCP : Double ;
    TDJ,TCJ         : Double ;
    XDEL            : TDELMODIF ;
    stReq           : String;
    QEcr	    : TQuery;
BEGIN
if Action=taCreat then Exit ;
 // SBO 05/07/2007 : gestion pb DB2 avec champ blocnote à positionner à la fin
stReq := 'SELECT ' + GetSelectAll('E', True) + ', E_BLOCNOTE FROM ECRITURE WHERE ' + WhereEcriture(tsGene,M,False) ;
stReq := stReq + ' ORDER BY E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE' ;
QEcr := OpenSQL(stReq, True);
Lig:=0 ; TDJ:=0 ; TCJ:=0 ;
InitMove(50,'') ; NbLigEcr:=0 ; NbLigAna:=0 ; OkE:=False ;
TDD:=0 ; TCD:=0 ; TDP:=0 ; TCP:=0 ; NbEche:=0 ;
While Not QEcr.EOF do
   BEGIN
   {Sauvegarde des dates modif}
   XDEL:=TDELMODIF.Create ;
   XDEL.NumLigne:=QEcr.FindField('E_NUMLIGNE').AsInteger ;
   XDEL.NumEcheVent:=QEcr.FindField('E_NUMECHE').AsInteger ;
   XDEL.DateModification:=QEcr.FindField('E_DATEMODIF').AsDateTime ;
   TDELECR.Add(XDEL) ;
   {Lectures des infos}
   if QEcr.FindField('E_CONFIDENTIEL').AsString>'0' then PieceConf:=True ;
   MoveCur(False) ; Inc(NbLigEcr) ;
//   LDD:=0 ; LCD:=0 ; LDP:=0 ; LCP:=0 ; LDE:=0 ; LCE:=0 ;
   NumE:=XDEL.NumEcheVent ;
   if (NumE<=0) or ((NumE=1) and (OkE)) then
      BEGIN
      if OkE then RuptureEche(Lig,TDD,TCD,TDP,TCP) ; OkE:=FALSE ; NbEche:=0 ;
      END ;
   if NumE=1 then
      BEGIN
      OkE:=True ; NbEche:=1 ;
      END else if NumE>1 then
      BEGIN
      OkE:=True ; Inc(NbEche) ;
      END ;
   if NbEche<=1 then
      BEGIN
      Inc(Lig) ; DefautLigne(Lig,False) ;
      OBM:=GetO(GS,Lig) ; OBM.ChargeMvt(QEcr) ; OBM.HistoMontants ;
      OBM.CompS:=True ;
      TDD:=0 ; TCD:=0 ; TDP:=0 ; TCP:=0 ;
      GS.Cells[SA_Gen,Lig]:=QEcr.FindField('E_GENERAL').AsString ;
      GS.Cells[SA_Aux,Lig]:=QEcr.FindField('E_AUXILIAIRE').AsString ;
      GS.Cells[SA_RefI,Lig]:=QEcr.FindField('E_REFINTERNE').AsString ;
      GS.Cells[SA_Lib,Lig]:=QEcr.FindField('E_LIBELLE').AsString ;
      if EcheMvt(QEcr) then AlloueEche(Lig) ;
      END ;
   if ((Action=taModif) and (M.Simul='N') and (Not M.Valide) and (QEcr.FindField('E_TRESOLETTRE').AsString='X')) then TresoLettre:=True ;
   LDD:=QEcr.FindField('E_DEBITDEV').AsFloat  ; LCD:=QEcr.FindField('E_CREDITDEV').AsFloat ;
   LDP:=QEcr.FindField('E_DEBIT').AsFloat     ; LCP:=QEcr.FindField('E_CREDIT').AsFloat ;
   TDD:=TDD+LDD ; TCD:=TCD+LCD ; TDP:=TDP+LDP ; TCP:=TCP+LCP ;
   TDJ:=TDJ+LDP ; TCJ:=TCJ+LCP ;
   if NbEche>=1 then AlimNoyauEche(Lig,(NbEche=1),LDD,LCD,LDP,LCP, QEcr) ;
      GS.Cells[SA_Debit,Lig]:=StrS(LDD,DECDEV) ; GS.Cells[SA_Credit,Lig]:=StrS(LCD,DECDEV) ;
   FormatMontant(GS,SA_Debit,Lig,DECDEV) ; FormatMontant(GS,SA_Credit,Lig,DECDEV) ;
   CaptureRatio(Lig,True) ; GereNewLigne(GS) ;
   YABienTP(QEcr) ;
   QEcr.Next ;
   END ;
Ferme(QEcr);
FiniMove ;
if NbEche>=1 then RuptureEche(Lig,TDD,TCD,TDP,TCP) ;
SAJAL.OldDebit:=TDJ ; SAJAL.OldCredit:=TCJ ;
//CommitTrans ;
END ;

procedure TFSaisie.ChargeLignes ;
BEGIN
ChargeEcritures ;
ChargeComptes ;
TripoteYATP ;
if Action<>taCreat then BEGIN GS.Row:=1 ; GS.Col:=SA_Gen ; GS.SetFocus ; END ;
AffichePied ; NbLOrig:=GS.RowCount ;

END ;

procedure TFSaisie.ChargeSoldesComptes ;
Var i,j,k, lInCpt   : integer ;
    OBM 	    : TOBM ;
    Gen,Aux,Sect,Ax : String ;
    D,C             : Double ;
    lStNat          : string ;
    lIndex          : integer ;
BEGIN
if PasModif then Exit ;
  for i:=1 to GS.RowCount-2 do
    BEGIN
    OBM:=GetO(GS,i) ;
    if OBM=Nil then Continue ;
    Gen	:= OBM.GetMvt('E_GENERAL') ;
    lIndex:=FInfo.Compte.GetCompte(Gen) ;
    if lIndex>-1 then
     lStNat:= FInfo.Compte_GetValue('G_NATUREGENE') ;
    Aux	:= OBM.GetMvt('E_AUXILIAIRE') ;
    D		:= OBM.GetMvt('E_DEBIT') ;
    C		:= OBM.GetMvt('E_CREDIT') ;
    Ajoute(TS,'G',Gen,D,C,-1,-1,iDate1900) ;
    Ajoute(TS,'T',Aux,D,C,-1,-1,iDate1900) ;
    if M.Simul = 'N' then
      AjouteAno(TSA,OBM,lStNat,true) ;
    for j:=1 to Min( MaxAxe, OBM.Detail.Count ) do
      for k:=0 to OBM.Detail[j-1].Detail.Count-1 do
        BEGIN
          D := TOBM(OBM.Detail[j-1].Detail[k]).GetMvt('Y_DEBIT') ;
          C := TOBM(OBM.Detail[j-1].Detail[k]).GetMvt('Y_CREDIT') ;

          //SG6 27.01.05 Gestion de mode croisaxe
          if not (VH^.AnaCroisaxe) then
          begin
            Ax	    := TOBM(OBM.Detail[j-1].Detail[k]).GetMvt('Y_AXE')  ;
            Sect    := TOBM(OBM.Detail[j-1].Detail[k]).GetMvt('Y_SECTION') ;
            Ajoute(TS,Ax,Sect,D,C,-1,-1,iDate1900) ;
          end
          else
          begin
            for lInCpt := 1 to MaxAxe do
            begin
              Ax    := 'A' + IntToStr(lInCpt) ;
              Sect  := TOBM(OBM.Detail[j-1].Detail[k]).GetMvt('Y_SOUSPLAN' + IntToStr( lInCpt) ) ;
              if Sect <> '' then
              begin
                Ajoute(TS,Ax,Sect,D,C,-1,-1,iDate1900) ;
              end;
            end;
          end;
        END ;
    END ;
END ;


procedure TFSaisie.ChargeAnals ;
Var NumL,NumV,NumA,OldL : integer ;
    OBA            : TOBM ;
    OBM			       : TOBM ;
    Ax             : String3 ;
    QAna           : TQuery ;
    X              : TDELMODIF ;
BEGIN
	if Action=taCreat then Exit ;

  // SBO 05/07/2007 : gestion pb DB2 avec champ blocnote à positionner à la fin
  QAna := OpenSQL('SELECT ' + GetSelectAll('Y', True) + ', Y_BLOCNOTE FROM ANALYTIQ'
                   + ' WHERE ' + WhereEcriture(tsAnal,M,False)
  							   + ' ORDER BY Y_JOURNAL, Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE, Y_AXE, Y_NUMVENTIL', True) ;
	OldL     := -1 ;
  NbLigAna := 0 ;
  OBM      := nil ;

  While Not QAna.EOF do
    BEGIN
    NumL := QAna.FindField('Y_NUMLIGNE').AsInteger ;
    NumV := QAna.FindField('Y_NUMVENTIL').AsInteger ;
    NumA := StrToInt(Copy(QAna.FindField('Y_AXE').AsString,2,1)) ;

		if OldL<>NumL then
      BEGIN
      OBM := GetO(GS,NumL) ;
      if OBM = nil then break;
      if OBM.Detail.Count<MaxAxe then AlloueAnal(NumL) ;
//      OBA.ModeConf := QAna.FindField('Y_CONFIDENTIEL').AsString ;
      Revision:=False ;
      END ;

    Ax:='A'+Chr(48+NumA) ;
    OBA:=TOBM.Create(EcrAna,Ax,False,OBM.Detail[NumA-1]) ;
    OBA.ChargeMvt(QAna) ;
    OBA.HistoMontants ;

    OldL := NumL ;

    {Réseau}
    X:=TDELMODIF.Create ;
    X.NumLigne					:= NumL ;
    X.NumEcheVent				:= NumV ;
    X.Ax								:= Ax ;
    X.DateModification	:= QAna.FindField('Y_DATEMODIF').AsDateTime ;
    TDELANA.Add(X) ;
    Inc(NbLigAna) ;
    QAna.Next ;
    END ;

	Ferme(QAna) ;
END ;


procedure TFSaisie.ChargeCompl ;
var
 lQ : TQuery ;
 OBM : TOBM ;
 lInIndex : integer ;
begin
if Action=taCreat then Exit ;
lQ := OpenSQL('Select * from ECRCOMPL where EC_JOURNAL="'+M.JAL+'" AND EC_EXERCICE="'+M.EXO+'"'
              +' AND EC_DATECOMPTABLE="'+UsDateTime(M.DateC)+'" AND EC_NUMEROPIECE='+InttoStr(M.Num)
              +' AND EC_QUALIFPIECE="'+M.Simul+'" ' 
              + ' ORDER BY EC_JOURNAL, EC_EXERCICE, EC_DATECOMPTABLE, EC_NUMEROPIECE, EC_NUMLIGNE, EC_NUMECHE', true ) ;
while not lQ.EOF do
 begin
  lInIndex := lQ.FindField('EC_NUMLIGNE').AsInteger ;
  OBM := GetO(GS,lInIndex) ;
  if OBM = nil then break;
  OBM.CompE := true ;
  CCreateDBTOBCompl(OBM,FEcrCompl, lQ) ;
  lQ.Next ;
 end ; // if
ferme(lQ) ;
end ;


{==========================================================================================}
{=========================== Allocations et Désallocations ================================}
{==========================================================================================}
function  TFSaisie.DateToJour ( DD : TDateTime ) : String ;
Var y,m,d : Word ;
    StD   : String ;
BEGIN
DecodeDate(DD,y,m,d) ; StD:=IntToStr(d) ; if d<10 then StD:='0'+StD ;
Result:=StD ;
END ;

procedure TFSaisie.AlloueEcheAnal(C,L : LongInt) ;
Var CGen : TGGeneral ;
    CAux : TGTiers ;
BEGIN
if C=SA_Gen then
   BEGIN
   CGen:=GetGGeneral(GS,L) ; if CGen=Nil then Exit ;
   if Lettrable(CGen)<>NonEche then AlloueEche(L) ;
   if Ventilable(CGen,0) then AlloueAnal(L) ;
   END else
   BEGIN
   CAux:=GetGTiers(GS,L) ; if CAux=Nil then Exit ;
   if CAux.Lettrable then AlloueEche(L) ;
   END ;
END ;

procedure TFSaisie.AlloueEche(LaLig : LongInt) ;
Var TTM : TMOD ;
begin
if GS.Objects[SA_NumP,LaLig]<>Nil then Exit ;
TTM:=TMOD.Create ;
GS.Objects[SA_NumP,LaLig]:=TObject(TTM) ;
Revision:=False ;
end ;

procedure TFSaisie.DesalloueEche(LaLig : LongInt) ;
Var OBM : TOBM ;
begin
if GS.Objects[SA_NumP,LaLig]=Nil then Exit ;
TMOD(GS.Objects[SA_NumP,LaLig]).Free ; GS.Objects[SA_NumP,LaLig]:=Nil ;
OBM:=GetO(GS,LaLig) ; if OBM=Nil then Exit ;
// Re-init des infos echéances
OBM.PutMvt('E_LETTRAGE','') ; OBM.PutMvt('E_LETTRAGEDEV','-') ;
OBM.PutMvt('E_COUVERTURE',0) ; OBM.PutMvt('E_COUVERTUREDEV',0) ; //OBM.PutMvt('E_COUVERTUREEURO',0) ;
OBM.PutMvt('E_MODEPAIE','') ; OBM.PutMvt('E_NIVEAURELANCE',0) ;
OBM.PutMvt('E_ETATLETTRAGE','RI') ; OBM.PutMvt('E_NUMECHE',0) ;
OBM.PutMvt('E_DATEECHEANCE',IDate1900) ; OBM.PutMvt('E_DATERELANCE',IDate1900) ;
OBM.PutMvt('E_DATEVALEUR',IDate1900) ;
Revision:=False ;
end ;

procedure TFSaisie.AlloueAnal(LaLig : LongInt) ;
var OBM : TOBM ;
begin
  OBM := GetO(GS,LaLig) ;
  if OBM = nil then Exit;
  AlloueAxe(OBM) ;
  Revision:=False ;
end ;

procedure TFSaisie.DesalloueAnal(LaLig : LongInt) ;
var OBM : TOBM ;
begin
  OBM := GetO(GS,LaLig) ;
  if OBM.Detail.Count=0 then Exit ;
  OBM.ClearDetail;
  Revision:=False ;
end ;

procedure TFSaisie.DesalloueAnal(LaLig : LongInt; Ventilable : T5B);
var
  OBM : TOBM ;
  i : Integer;
  bClear : Boolean;
begin
  OBM := GetO(GS,LaLig) ;
  bClear := True;

  for i := 0 to OBM.Detail.Count-1 do
    begin
    if not Ventilable[i+1]
      then OBM.Detail[i].ClearDetail
      else bClear := False;
    end;

  if bClear then OBM.ClearDetail;

  Revision:=False ;
end ;

{==========================================================================================}
{=========================== Initialisations et valeurs par défaut ========================}
{==========================================================================================}
procedure TFSaisie.GereEnabled ( Lig : integer ) ;
Var OKL,Remp,LectSeul,Visu,OkRIB : Boolean ;
    O   : TOBM ;
    S : string;
BEGIN
Remp:=EstRempli(GS,Lig) ; O:=GetO(GS,Lig) ; Visu:=(Action=taConsult) ;
OkL:=((Remp) and (O<>Nil)) ; LectSeul:=PasToucheLigne(Lig) ;
BSolde.Enabled:=((GS.RowCount>2) and (Not LectSeul) and (Not Visu) and (Not ModeLC)) ;
BInsert.Enabled:=((Estrempli(GS,1)) and (Not Visu) and (Not ModeLC)) ;
BSDel.Enabled:=BSolde.Enabled ;
BEche.Enabled:=((GS.Objects[SA_NumP,Lig]<>Nil) or (LigneEche(GS,Lig))) ;
BVentil.Enabled:=( (O<>Nil) and VentilationExiste(O) ) ;
if Not M.ANouveau then BControleTVA.Enabled:=((isHT(GS,Lig,True)) and (Not Visu) and (Not ModeLC)) ;
BProrata.Enabled:=((BSolde.Enabled) and (Not M.ANouveau) and (Not LectSeul) and (Not ModeLC)) ;
BComplement.Enabled:=((OkL) and (Not M.ANouveau)) ;
BModifSerie.Enabled:=((GS.RowCount>2) and (Not Visu)) ;
{$IFDEF CCS3}
BChoixRegime.Enabled:=False ;
BVidePiece.Enabled:=False ;
{$ELSE}
BChoixRegime.Enabled:=((Not Visu) and (Not ModeLC)) ;
BVidePiece.Enabled:=BSDel.Enabled ;
{$ENDIF}
if OkL then
   BEGIN
   BPieceGC.Enabled:=((Not ModeLC) and (O.GetMvt('E_REFGESCOM')<>'')) ;
   bLignesGC.Enabled := ((GetParamSocSecur('SO_OKLIENLIGECR', false)) and (BPieceGC.Enabled) and (O.GetString('E_TYPEMVT') = 'HT'));
   // BDU - 16/11/06, Modifie le libellé du menu selon que l'écriture
   // pointe sur une pièce commerciale ou sur des frais
   S := Copy(O.GetMvt('E_REFGESCOM'), 1, 3);
   if (S = 'AA;') or (S = 'AD;') or (S = 'AR;') then
     BPieceGC.Hint := TraduireMemoire('Voir les frais')
   else
     BPieceGC.Hint := TraduireMemoire('Voir pièce commerciale');
   // fin
{$IFDEF SANSIMMO}
   BZoomImmo.Enabled:=FALSE ;
{$ELSE}
   BZoomImmo.Enabled:=((Not ModeLC) and (O.GetMvt('E_IMMO')<>'') and ((VH^.OkModImmo) or (V_PGI.VersionDemo))) ;
{$ENDIF}
{$IFDEF CCS3}
   BZoomTP.Enabled:=False ;
{$ELSE}
   BZoomTP.Enabled:=((Not ModeLC) and (O.GetMvt('E_PIECETP')<>'') and (VH^.OuiTP)) ;
{$ENDIF}
   END ;

// FQ 18704 : On Autorise la vision en devise pivot des pièces d'ECC :
BSwapPivot.Enabled:=(GS.RowCount>2) and ((M.Nature='ECC') or ModeDevise or (ModeForce=tmPivot) and (Not ModeLC)) ;
{$IFDEF MODENT1}
BChancel.Enabled:=((GS.RowCount<=2) and (Not EstRempli(GS,1)) and (Action=taCreat) and (Not ModeLC) and (V_PGI.DevisePivot <> DEV.Code)) ;
BSaisTaux.Enabled:=((GS.RowCount<=2) and (ModeDevise) and (Action=taCreat) and (Not ModeLC) and (V_PGI.DevisePivot <> DEV.Code));
{$ELSE}
BChancel.Enabled:=((GS.RowCount<=2) and (Not EstRempli(GS,1)) and (Action=taCreat) and (Not ModeLC) and (Not EstMonnaieIN(DEV.Code))) ;
BSaisTaux.Enabled:=((GS.RowCount<=2) and (ModeDevise) and (Action=taCreat) and (Not ModeLC) and (Not EstMonnaieIN(DEV.Code))) ;
{$ENDIF MODENT1}
BModifTva.Enabled:=((OkL) and (OuiTvaLoc) and ((isHT(GS,Lig,True)) or (isTiers(GS,Lig)))) ;
if VH^.PaysLocalisation=CodeISOES then
   BGenereTVA.Enabled:=(BModifTva.Enabled) and (Assigned(SAJAL)) and (pos(';'+SAJAL.NATUREJAL+';',';ACH;VTE;')>0) ; //XVI 24/02/2005
{Gestion du RIB}
OkRIB:=OkScenario and ((GestionRIB='MAN') or (GestionRIB='PRI')) ;
OkRIB:=OkRIB or ( ((M.TypeGuide='DEC') or (M.TypeGuide='ENC') or (SuiviMP)) and (M.ExportCFONB) ) ;
BModifRIB.Enabled:=((OkL) and (OkRIB)) ;
if ModeLC then
   BEGIN
   BLibAuto.Enabled:=False ;
   BMenuGuide.Enabled:=False ; BGuide.Enabled:=False ; BCreerGuide.Enabled:=False ;
   END ;
if Action=taConsult then
   BEGIN
   BMenuAction.Enabled:=False ;
   BModifs.Enabled:=False ;
   END ;
{$IFDEF SCANGED}
BScan.Enabled:= OkL and (RechGuidId<>'') ;
bScanPdf.Enabled := OkL ;
{$ENDIF}
{$IFDEF TT}
BScan.Enabled:=true ;
BScan.Visible:=true ;
{$ENDIF}

END ;

{***********A.G.L.***********************************************
Auteur  ...... : Piot
Créé le ...... : 22/02/2005
Modifié le ... : 09/11/2006
Description .. : - LG  - 22/02/2005 - Suppression de la TOB des ecritures
Suite ........ : complementaire
Suite ........ : - LG - 09/11/2006 - chg des fct de raz des compte
Mots clefs ... : 
*****************************************************************}
procedure TFSaisie.ReInitPiece ( Force : boolean ) ;
Var bc : boolean ;
 i : integer ;
BEGIN
V_PGI.IOError:=oeOk ; CEstGood:=False ; GuideAEnregistrer:=FALSE ; appeldatepourguide:=FALSE ; CodeGuideFromSaisie:='' ;
{Grid}
GS.VidePile(True) ; GS.RowCount:=2 ; FinGuide ;
if ((Action<>taCreat) and (Not Force)) then Exit ;
{Pointeurs}
// Ré-init des infos comptes / auxi pour rechargement des totaux...
if Assigned( FInfo )  then // FQ 13246 SBO 04/08/2005
  begin
  FInfo.Compte.Clear ;
  FInfo.Aux.Clear ;
  end ;
// GG MEMCHECK
if TSA <> nil then
begin
for i:=0 to TSA.Count-1 do
 if assigned(TSA[i]) then Dispose(TSA[i]);
TSA.Free ;
end ;
VideListe(TS) ; TS.Free ; TS:=Nil ; TS:=TList.Create ; TSA:=TList.Create ; VideListe(TECHEORIG) ;
// GG MEMCHECK
VideListe(TGUI) ; TGUI.Free ; TGUI:=Nil ; TGUI:=TList.Create ; GestionRIB:='...' ;
Scenario.Free ; Scenario:=Nil ; Scenario:=TOBM.Create(EcrSais,'',True) ;
MemoComp.Free ; MemoComp:=Nil ; MemoComp:=HTStringList.Create ;
Entete.Free ; Entete:=Nil ; Entete:=TOBM.Create(EcrSais,'',True) ;
ModifSerie.Free ; ModifSerie:=Nil ; ModifSerie:=TOBM.Create(EcrGen,'',True) ;
OkScenario:=False ; if Action=taCreat then BEGIN ChargeScenario ; ChargeLibelleAuto ; END ;
FreeAndNil(FEcrCompl) ; //FreeAndNil(FInfo) ;
{Inits}
BValide.Enabled:=True ; DefautPied ;
H_NUMEROPIECE.Visible:=False ; H_NUMEROPIECE_.Visible:=True ;
E_JOURNAL.Enabled:=True ; E_DATECOMPTABLE.Enabled:=True ; InitEnteteJal(False,False,True) ;
PieceModifiee:=False ; Revision:=False ; OkMessAnal:=False ; DejaRentre:=False ; UnChange:=False ;
GS.SetFocus ; GS.Col:=SA_Gen ; GS.Row:=1 ; ModifNext:=False ; Volatile:=False ;
VideListe(TDELECR) ; VideListe(TDELANA) ; ModeConf:='0' ; HConf.Visible:=False ; PieceConf:=False ;
YATP:=Saiscomm.yaRien ; ListeTP.Clear ; RentreDate:=False ; RegimeForce:=False ;
LastTraiteCHQ:='' ; LastEche:=0 ;
{#TVAENC}
if ((M.Simul='N') or (M.Simul='S') or (M.Simul='R')) then OuiTvaLoc:=VH^.OuiTvaEnc ;
ExigeTva:=tvaMixte ; ExigeEntete:=tvaMixte ; ChoixExige:=False ;
BPopTva.ItemIndex:=Ord(ExigeEntete)+1 ; 
if VH^.PaysLocalisation=CodeISOES then
   BpopTVAChange(bpopTVA) ; //XVI 24/02/2005
ToutDebit:=True ; ToutEncais:=True ; GeneRegTva:='' ;
{Affichage}
if ((VH^.CPDateObli) and (Action=taCreat) and (E_DATECOMPTABLE.CanFocus)) then E_DATECOMPTABLE.SetFocus else
 if ((OkScenario) and (Scenario<>Nil) and (Action=taCreat) and ( Scenario.GetString('SC_DATEOBLIGEE')='X' ) and (E_DATECOMPTABLE.CanFocus)) then E_DATECOMPTABLE.SetFocus else
   BEGIN
   if Action=taCreat then GSEnter(Nil) ;
   GSRowEnter(Nil,1,bc,False) ;
   END ;
END ;

procedure TFSaisie.InitModifJal(Action : TActionFiche) ;
BEGIN
if Action=taCreat then Exit ;
Case CaseNatJal(SAJAL.NatureJal) of
   tzJVente       : E_NaturePiece.DataType:='ttNatPieceVente'  ;
   tzJAchat       : E_NaturePiece.DataType:='ttNatPieceAchat'  ;
   tzJBanque      : E_NaturePiece.DataType:='ttNatPieceBanque' ;
   tzJEcartChange : E_NaturePiece.DataType:='ttNatPieceEcartChange' ;
   tzJOD          : E_NaturePiece.DataType:='ttNaturePiece'    ;
   END ;
END ;

procedure TFSaisie.InitEnteteJal(Totale,Zoom,ReInit : boolean) ;
Var MM  : String17 ;
    DD : TDateTime ;
BEGIN
if Action<>taCreat then Exit ; if SAJAL=Nil then Exit ;
PositionneDevise(ReInit) ;
if Totale then
   BEGIN
   if Action=taCreat then
      BEGIN
      BMenuGuide.Enabled:=True ; BGuide.Enabled:=True ; BCreerGuide.Enabled:=True ;
      END ;
   if E_JOURNAL.DataType='TTJALANOUVEAU' then
      BEGIN
      BMenuGuide.Enabled:=False ; BGuide.Enabled:=False ; BCreerGuide.Enabled:=False ;
      END ;
   Case CaseNatJal(SAJAL.NatureJal) of
      tzJVente  : BEGIN E_NaturePiece.DataType:='ttNatPieceVente'  ; E_NaturePiece.Value:='FC' ; END ;
      tzJAchat  : BEGIN E_NaturePiece.DataType:='ttNatPieceAchat'  ; E_NaturePiece.Value:='FF' ; END ;
      tzJBanque : BEGIN
                    // BPY le 03/10/2003 : fiche de bug 12620
                    E_NaturePiece.DataType:='ttNatPieceBanque' ;
                    {$IFDEF CCMP}
                    if not (VH^.CCMP.LotCli) then E_NaturePiece.Value:='RF'
                    else
                    {$ENDIF}
                        E_NaturePiece.Value:='RC' ;
                    // fin BPY

                  END ;
 tzJEcartChange : BEGIN
                  E_NaturePiece.DataType:='ttNatPieceEcartChange' ;
                  if ((Action<>taCreat) or (SAJAL.MultiDevise))
                    then E_NaturePiece.Value:='ECC'
                    else E_NATUREPIECE.Value:='OD' ;
                  END ;
      tzJOD     : BEGIN E_NaturePiece.DataType:='ttNaturePiece' ; E_NaturePiece.Value:='OD' ; END ;
      END ;
   END ;
if Not Zoom then
   BEGIN
   DD:=StrToDate(E_DATECOMPTABLE.Text) ;
   if ( M.Simul <> 'N' ) and ( M.Simul <> 'I' ) // Modif IFRS
     then NumPieceInt:=GetNum(EcrGen,SAJAL.COMPTEURSIMUL,MM,DD)
     else NumPieceInt:=GetNum(EcrGen,SAJAL.COMPTEURNORMAL,MM,DD) ;
   InitGrid ;
   END ;
if NumPieceInt>0 then E_NUMEROPIECE.Caption:=IntToStr(NumPieceInt) ;
END ;

procedure TFSaisie.GereDeviseEcart ;
BEGIN
if ((E_NaturePiece.Value='') or (E_Journal.Value='')) then Exit ;

// FQ 18704
  E_DEVISE.Visible:=True ;
  E_DEVISE_.Visible:=False ;
  E_DEVISE_.Value:=V_PGI.DevisePivot ;
{
if E_NATUREPIECE.Value='ECC' then
   BEGIN
   E_DEVISE.Value:=V_PGI.DevisePivot ;
   E_DEVISEChange(Nil) ;
   E_DEVISE_.Visible:=True ;
   E_DEVISE.Visible:=False ;
   if ((E_DEVISE_.Value='') and (E_DEVISE_.Items.Count>0))
     then E_DEVISE_.ItemIndex:=0 ;
   END else
   BEGIN
   E_DEVISE.Visible:=True ;
   E_DEVISE_.Visible:=False ;
   E_DEVISE_.Value:=V_PGI.DevisePivot ;
   END ;
}

END ;

procedure TFSaisie.DefautEntete ;
BEGIN
if not M.Parcentral and ((VH^.CPExoRef.Code<>'') and (DateSais>0) and (Action=taCreat)
and (ctxPCL in V_PGI.PGIContexte)) and (M.TypeGuide<>'ENC') and (M.TypeGuide<>'DEC') then M.DateC:=DateSais ;
E_DATECOMPTABLE.Text:=DateToStr(M.DateC) ;
E_JOURNAL.Value:=M.JAL ; E_NATUREPIECE.Value:=M.Nature ;
E_ETABLISSEMENT.Value:=M.Etabl ; E_ETABLISSEMENT.Enabled:=VH^.EtablisCpta ;
if Action=taCreat then PositionneEtabUser(E_ETABLISSEMENT) ;
DatePiece:=M.DateC ;
SetTypeExo ;   // FQ 13246 : SBO 30/03/2005
// FQ 18704
  E_DEVISE.Value:=M.CodeD ;
  E_DEVISE_.Value:=V_PGI.DevisePivot ;
{
if M.Nature='ECC' then
   BEGIN
   E_DEVISE.Value:=V_PGI.DevisePivot ;
   E_DEVISE_.Value:=M.CodeD ;
   END else
   BEGIN
    E_DEVISE.Value:=M.CodeD ;
    E_DEVISE_.Value:=V_PGI.DevisePivot ;
   END ;
}
Case Action of
   taCreat : BEGIN
             E_DEVISE.Enabled:=False ; E_NumeroPiece.Caption:='' ;
             H_NUMEROPIECE.Visible:=False ; H_NUMEROPIECE_.Visible:=True ;
             GS.Enabled:=False ;
             END ;
   taModif : BEGIN
             E_JOURNAL.Enabled:=False ; E_DEVISE.Enabled:=False ; E_DEVISE_.Enabled:=False ;
             E_DATECOMPTABLE.Enabled:=False ; E_NATUREPIECE.Enabled:=False ;
             E_ETABLISSEMENT.Enabled := False; {JP 04/07/07 : FQ 18070}
             H_NUMEROPIECE.Visible:=True ; H_NUMEROPIECE_.Visible:=False ; NumPieceInt:=M.Num ;
             if Not SuiviMP then E_NumeroPiece.Caption:=InttoStr(M.Num)
                            else E_NUMEROPIECE.Caption:='' ; 
             GS.Enabled:=True ;
             GS.SetFocus ;
             END ;
 taConsult : BEGIN
             E_NumeroPiece.Caption:=InttoStr(M.Num) ; NumPieceInt:=M.Num ;
             H_NUMEROPIECE.Visible:=True ; H_NUMEROPIECE_.Visible:=False ;
             GS.Enabled:=True ; PEntete.Enabled:=False ;
             END ;
   END ;
 GereDeviseEcart ;
 END ;

procedure TFSaisie.DefautPied ;
BEGIN
SI_TotDS:=0 ; SI_TotCS:=0 ;
SI_TotDD:=0 ; SI_TotCD:=0 ;
SI_TotDP:=0 ; SI_TotCP:=0 ;
ZeroBlanc(PPied) ;
END ;

procedure TFSaisie.InitGrid ;
BEGIN
if Action<>taCreat then Exit ;
GS.Enabled:=True ; DefautLigne(GS.RowCount-1,True) ;
GS.Col:=SA_Gen ; GS.Row:=1 ;
END ;

procedure TFSaisie.AttribLeTitre ;
Var i,j : integer ;
    C   : Char ;
BEGIN
i:=1 ; j:=1 ;
Case Action of taConsult : i:=1 ; taModif : i:=2 ; taCreat : i:=3 ; END ;
C:=M.Simul[1] ; Case C of 'N' : j:=1 ; 'S' : j:=2 ; 'P' : j:=3 ; 'U' : j:=4 ; 'R' : j:=5 ; 'I' : j:=6 END ;
if M.ANouveau then j:=7 ;
Caption:=HTitres.Mess[10*(i-1)+(j-1)] ;
{Help Contexte}
Case Action of
   taConsult : BEGIN
               Case C of
                  'N' : HelpContext:=7256200 ;
                  'S' : HelpContext:=7277200 ;
                  'U' : HelpContext:=7298200 ;
                  'P' : HelpContext:=7313200 ;
                  'R' : HelpContext:=7651100 ;
                  'I' : HelpContext:=7249100 ; // Modif IFRS
                  END ;
               if M.Anouveau then HelpContext:=7728000 ;
               END ;
     taCreat : BEGIN
               Case C of
                  'N' : HelpContext:=7244000 ;
                  'S' : HelpContext:=7277000 ;
                  'U' : HelpContext:=7298000 ;
                  'P' : HelpContext:=7313000 ;
                  'R' : HelpContext:=7649000 ;
                  'I' : HelpContext:=7249100 ; // Modif IFRS
                  END ;
               if M.Anouveau then HelpContext:=7724000 ;
               END ;
     taModif : BEGIN
               Case C of
                  'N' : HelpContext:=7256300 ;
                  'S' : HelpContext:=7277300 ;
                  'U' : HelpContext:=7298300 ;
                  'P' : HelpContext:=7313300 ;
                  'R' : HelpContext:=7651200 ;
                  'I' : HelpContext:=7249100 ; // Modif IFRS
                  END ;
               END ;
   END ;
if SuiviMP then Caption:=HTitres.Mess[27] ;
UpdateCaption(Self) ;
END ;

procedure TFSaisie.AttribNumeroDef ;
Var Facturier  : String3 ;
    Lig,NumAxe,NumV : integer ;
    O          : TOBM ;
    DD         : TDateTime ;
BEGIN
if ((Action<>taCreat) and (Not SuiviMP)) then Exit ;
if ( M.Simul <> 'N' ) and ( M.Simul <> 'I' ) and ( Not SuiviMP ) // Modif IFRS
  then Facturier:=SAJAL.CompteurSimul
  else Facturier:=SAJAL.CompteurNormal ;
DD:=StrToDate(E_DATECOMPTABLE.Text) ;
SetIncNum(EcrGen,Facturier,NumPieceInt,DD) ;
H_NUMEROPIECE.Visible:=True ; H_NUMEROPIECE_.Visible:=False ;
E_NUMEROPIECE.Caption:=IntToStr(NumPieceInt) ;
M.LastNumCreat:=NumPieceInt ;
{Attribution aux objets du nouveau numéro}
	try
    for Lig:=1 to GS.RowCount-1 do
      BEGIN
      GS.Cells[SA_NumP,Lig]:=IntToStr(NumPieceInt) ;
      O:=GetO(GS,Lig) ;
      if O=Nil then Break ;
      O.PutMvt('E_NUMEROPIECE',NumPieceInt) ;
      for NumAxe:=1 to O.Detail.Count do
        for NumV:=0 to O.Detail[NumAxe-1].Detail.Count-1 do
          TOBM(O.Detail[NumAxe-1].Detail[NumV]).PutMvt('Y_NUMEROPIECE',NumPieceInt) ;
      END ;
	except
 		on e:exception do MessageAlerte( 'Erreur lors de l''affectation du numéro définitif'+#10#13+e.Message );
end;
END ;

{==========================================================================================}
{================================== Scenario de saisie ====================================}
{==========================================================================================}
procedure TFSaisie.LettrageEnSaisie ;
Var X : RMVT ;
    Q : TQuery ;
    Nb  : integer ;
BEGIN
{$IFNDEF GCGC}
if Action=taConsult then Exit ;
if M.Simul<>'N' then Exit ;
X.Axe:='' ; X.Etabl:=E_ETABLISSEMENT.Value ;
X.Jal:=E_JOURNAL.Value ; X.Exo:=QuelExoDT(StrToDate(E_DATECOMPTABLE.Text)) ;
X.CodeD:=E_DEVISE.Value ; X.Simul:=M.Simul ; X.Nature:=E_NATUREPIECE.Value ;
X.DateC:=StrToDate(E_DATECOMPTABLE.Text) ; X.DateTaux:=DEV.DateTaux ;
X.Num:=NumPieceInt ; X.TauxD:=DEV.Taux ; X.Valide:=False ; X.ANouveau:=False ;
Q:=OpenSQL('Select Count(*) from ECRITURE Where '+WhereEcriture(tsGene,X,False)+' AND E_ECHE="X" AND E_NUMECHE>0 AND E_ETATLETTRAGE="AL"',True) ;
if Not Q.EOF then
   BEGIN
   Nb:=Q.Fields[0].AsInteger ;
   if Nb>0 then
      BEGIN
      if HPiece.Execute(24,caption,'')=mrYes then LettrerEnSaisie(X,Nb) ;
      END ;
   END ;
Ferme(Q) ;
{$ENDIF}
END ;

procedure TFSaisie.ChargeScenario ;
Var Q      : TQuery ;
    STVA,SQL,SS   : String ;
    StComp : String ;
BEGIN
BGenereTVA.Enabled:=False ;
Scenario.PutDefautDivers ; OkScenario:=False ; DejaRentre:=False ;
Entete.Free ; Entete:=TOBM.Create(EcrGen,'',True) ;
if PasModif then Exit ; if M.ANouveau then Exit ;
if ((E_JOURNAL.Value='') or (E_NATUREPIECE.Value='')) then Exit ;
SS:=M.Simul ; if SuiviMP then SS:='N' ;
SQL:='Select * from SUIVCPTA Where SC_JOURNAL="'+E_JOURNAL.Value+'" AND SC_NATUREPIECE="'+E_NATUREPIECE.Value+'"'
    +' AND (SC_QUALIFPIECE="'+SS+'" OR SC_QUALIFPIECE="" OR SC_QUALIFPIECE="...") ' ;
SQL:=SQL+' AND (SC_ETABLISSEMENT="" OR SC_ETABLISSEMENT="'+E_ETABLISSEMENT.Value+'" OR SC_ETABLISSEMENT="...")'
    +' AND (SC_USERGRP="" OR SC_USERGRP="'+V_PGI.Groupe+'" OR SC_USERGRP="...")'
    +' order by SC_USERGRP, SC_ETABLISSEMENT, SC_QUALIFPIECE ' ;
Q:=OpenSQL(SQL,True) ;
// SBO 17/08/2004 : FQ 13892 : pb de réinitialisation du paramétrage lors du changement des scénarii 
MemoComp.Free ; MemoComp:=Nil ; MemoComp:=HTStringList.Create ;
// Fin FQ 13892
if Not Q.EOF then
   BEGIN
   Scenario.ChargeMvt(Q) ;
   if Not TMemoField(Q.FindField('SC_ATTRCOMP')).IsNull then
{$IFDEF EAGLCLIENT}
   	 MemoComp.Text := (TMemoField(Q.FindField('SC_ATTRCOMP')).AsString) ;
{$ELSE}
   	 MemoComp.Assign(TMemoField(Q.FindField('SC_ATTRCOMP'))) ;
{$ENDIF}
   OkScenario:=True ;
   AttribLeTitre ;
   END ;
Ferme(Q) ;
if OkScenario then
   BEGIN
   StComp:=Scenario2Comp(Scenario);
   Entete.PutMvt('E_ETAT',StComp);
   STVA:=Scenario.GetMvt('SC_CONTROLETVA') ; BGenereTVA.Enabled:=((STVA='GTD') or (STVA='GTG')) ;
   ValideEcriture:=((Action=taCreat) and (M.Simul='N') and (Scenario.GetMvt('SC_VALIDE')='X')) ;
   GestionRIB:=Scenario.GetMvt('SC_RIB') ;
   NeedEdition:=False ; EtatSaisie:='' ;
   if Scenario.GetMvt('SC_BROUILLARD')='X' then if (Not M.MajDirecte) And (Not M.MSED.MultiSessionEncaDeca) then BEGIN NeedEdition:=True ; EtatSaisie:='' ; END ;
   if Scenario.GetMvt('SC_DOCUMENT')<>'' then if (Not M.MajDirecte) And (Not M.MSED.MultiSessionEncaDeca) then BEGIN NeedEdition:=True ; EtatSaisie:=Scenario.GetMvt('SC_DOCUMENT') ; END ;
   END else
   BEGIN
   ValideEcriture:=False ; GestionRIB:='...' ;
   NeedEdition:=False ; EtatSaisie:='' ;
   END ;
END ;

{==========================================================================================}
{============================== Traitements liés à l'entête ===============================}
{==========================================================================================}
procedure TFSaisie.E_NATUREPIECEChange(Sender: TObject);
begin
if E_NaturePiece.Value='' then Exit ;
if ((SAJAL<>Nil) and (Action=taCreat) and (E_NATUREPIECE.Value='ECC')) then
  if Not SAJAL.MultiDevise then
    BEGIN
    E_NATUREPIECE.Value:='' ;
    HPiece.Execute(33,caption,'') ;
    Exit ;
    END ;
PieceModifiee:=True ; AutoCharged:=False ; Revision:=False ;
ChargeLibelleAuto ; GereDeviseEcart ;
ChargeScenario ;
GetSorteTva ; AfficheExigeTva ;
end;

procedure TFSaisie.E_DATECOMPTABLEChange(Sender: TObject);
begin
PieceModifiee:=True ;
end;

procedure TFSaisie.E_ETABLISSEMENTChange(Sender: TObject);
begin
PieceModifiee:=True ;
if ((E_ETABLISSEMENT.Value='') or (E_ETABLISSEMENT.Value=M.Etabl)) then Exit ;
ChargeScenario ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 25/07/2005
Modifié le ... :   /  /    
Description .. : - LG - 25/07/2005 - FB 15908 - afiichage d'un msg pour un 
Suite ........ : compte accelere ds l'entete de le fenetre
Mots clefs ... : 
*****************************************************************}
procedure TFSaisie.E_JOURNALChange(Sender: TObject);
Var Jal : String ;
begin
Jal:=E_JOURNAL.Value ;
if (Jal='') then BEGIN SAJAL.Free ; SAJAL:=Nil ; Exit ; END ;
if ( not M.ParCentral) and (SAJAL<>Nil) then if SAJAL.JOURNAL=Jal then Exit ;
RentreDate := False; // FQ 10225
try
if ((VH^.JalAutorises<>'') and (Pos(';'+Jal+';',VH^.JalAutorises)<=0) and (Action=taCreat)) then
   BEGIN
   HPiece.Execute(26,caption,'') ;
   if SAJAL<>Nil then BEGIN SAJAL.Free ; SAJAL:=Nil ; END ;
   E_JOURNAL.Value:='' ;
   if Action=taCreat then BEGIN E_JOURNAL.SetFocus ; GS.Enabled:=False ; END ;
   Exit ;
   END ;
if (Action=taCreat) and ((Jal=VH^.JalATP) or (Jal=VH^.JalVTP)) then
   BEGIN
   HPiece.Execute(43,caption,'') ;
   if SAJAL<>Nil then BEGIN SAJAL.Free ; SAJAL:=Nil ; END ;
   E_JOURNAL.Value:='' ;
   if Action=taCreat then BEGIN E_JOURNAL.SetFocus ; GS.Enabled:=False ; END ;
   Exit ;
   END ;
Volatile:=False ; ChercheJal(Jal,False) ;
if ((Action=taCreat) and (NumPieceInt=0)) then
   BEGIN
   HPiece.Execute(27,caption,'') ;
   if SAJAL<>Nil then BEGIN SAJAL.Free ; SAJAL:=Nil ; END ;
   E_JOURNAL.Value:='' ; E_JOURNAL.SetFocus ; GS.Enabled:=False ;
   Exit ;
   END ;
GereDeviseEcart ;
finally
AttribLeTitre ;
if (SAJAL<>nil) and (SAJAL.J_ACCELERATEUR='X') and not OkScenario then
 Caption:= Caption + TraduireMemoire(' (accélérateur de saisie activé)') ;
UpdateCaption(Self) ;
end ;
END ;

procedure TFSaisie.AvertirPbTaux(Code : String3 ; DateTaux,DateCpt : TDateTime) ;
Var ii : integer ;
    OkTaux1 : boolean ;
begin
if Action<>taCreat then Exit ;
OkTaux1:=(Arrondi(DEV.Taux-1,ADecimP)=0) ;
if ((OkScenario) and (Not OkTaux1)) then
   BEGIN
   if Scenario.GetMvt('SC_ALERTEDEV')='AUC' then Exit else
    if Scenario.GetMvt('SC_ALERTEDEV')='MOI' then
       BEGIN
       if GetPeriode(DateTaux)=GetPeriode(DateCpt) then Exit ;
       END else
     if Scenario.GetMvt('SC_ALERTEDEV')='SEM' then
        BEGIN
        if NumSemaine(DateTaux)=NumSemaine(DateCpt) then Exit ;
        END ;
   END ;
{$IFNDEF MODENT1}
if ((EstMonnaieIN(Code)) and (DateCpt>=V_PGI.DateDebutEuro)) then
   BEGIN
   ii:=HDiv.Execute(14,Caption,'') ;
   if ii<>mrYes then ii:=HDiv.Execute(15,caption,'') ;
   if ii=mrYes then
      BEGIN
      FicheDevise(Code,taModif,False) ;
      DEV.Taux:=GetTaux(DEV.Code,DEV.DateTaux,DateCpt) ;
      END ;
   END else
{$ENDIF MODENT1}
   BEGIN
   if OkTaux1 then ii:=HDiv.Execute(8,caption,'') else ii:=HDiv.Execute(2,caption,'') ;
   if ii=mrYes then
      BEGIN
      FicheChancel(E_DEVISE.VALUE,True,DateCpt,taCreat,(DateCpt>=V_PGI.DateDebutEuro)) ;
      DEV.Taux:=GetTaux(DEV.Code,DEV.DateTaux,DateCpt) ;
      END ;
   END ;
end ;

Function TFSaisie.PbTaux ( DEV : RDevise ; DateCpt : TDateTime ) : boolean ;
Var Code : string3 ;
BEGIN
Result:=False ;
Code:=DEV.Code ;
if ((Code=V_PGI.DevisePivot) or (Code=V_PGI.DeviseFongible)) then Exit ;
{$IFDEF MODENT1}
Result:=(DEV.Taux=1);
{$ELSE}
if ((DateCpt<V_PGI.DateDebutEuro) or (Not EstMonnaieIn(Code))) then Result:=(DEV.Taux=1) else
   if EstMonnaieIn(Code) then Result:=(DEV.Taux=V_PGI.TauxEuro) ;
{$ENDIF MODENT1}
END ;

procedure TFSaisie.E_DEVISEChange(Sender: TObject);
Var DateCpt : TDateTime ;
//    Okok    : boolean ;
begin
if E_DEVISE.Value='' then BEGIN FillChar(DEV,Sizeof(DEV),#0) ; Exit ; END ;
DEV.Code:=E_DEVISE.Value ; GetInfosDevise(DEV) ;
//Okok:=True ;
Case Action of
   taCreat   : BEGIN
               if (M.CodeD<>V_PGI.DevisePivot) and (not Volatile) and (ModeForce=tmNormal) and
                  ((M.TypeGuide='ENC') or (M.TypeGuide='DEC')) then
                  BEGIN
                  DEV.Taux:=M.TauxD ; DEV.DateTaux:=M.DateTaux ;
                  if SaisieNewTaux2000(DEV,DatePiece) then BEGIN M.TauxD:=DEV.Taux ; Volatile:=True ; END ;
                  END else
                  BEGIN
                  if ModeForce=tmNormal then
                     BEGIN
                     if ((DEV.Code=V_PGI.DevisePivot) {or (EstMonnaieIN(DEV.Code))}) then Volatile:=False ;
                     DateCpt:=StrToDate(E_DATECOMPTABLE.Text) ;
                     if Not Volatile then
                        BEGIN
                        DEV.Taux:=GetTaux(DEV.Code,DEV.DateTaux,DateCpt) ;
                        if (E_DEVISE.Enabled) and (E_DEVISE.Value<>'') and
                           (E_DEVISE.Value<>V_PGI.DevisePivot) and
                           ((DEV.DateTaux<>DateCpt) or (PbTaux(DEV,DateCpt))) then AvertirPbTaux(DEV.Code,DEV.DateTaux,DateCpt) ;
                        END ;
                     END ;
                  END ;
               BSaisTaux.Enabled:=((GS.RowCount<=2) and (E_DEVISE.Value<>V_PGI.DevisePivot)) ;
               END ;
   taModif   : BEGIN DEV.Taux:=M.TauxD ; DEV.DateTaux:=M.DateTaux ; END ;
   taConsult : DEV.Taux:=M.TauxD ;
   END ;
PieceModifiee:=True ; ModeDevise:=EnDevise(DEV.Code) ;
if (DEV.Code=V_PGI.DevisePivot) then DEV.Symbole:=V_PGI.SymbolePivot ;
DecDev:=DEV.Decimale ;
ChangeFormatDevise(Self,DECDEV,DEV.Symbole) ;
end;

procedure TFSaisie.H_JOURNALDblClick(Sender: TObject);
begin ClickJournal ; end;

procedure TFSaisie.H_ETABLISSEMENTDblClick(Sender: TObject);
begin ClickEtabl ; end;

procedure TFSaisie.QuelNExo ;
Var EXO : String ;
BEGIN
if Action=taCreat then EXO:=QuelExoDT(StrToDate(E_DATECOMPTABLE.Text)) else EXO:=M.EXO ;
if EXO=VH^.Encours.Code then H_EXERCICE.Caption:='(N)' else
 if EXO=VH^.Suivant.Code then H_EXERCICE.Caption:='(N+1)' else
  H_EXERCICE.Caption:='(N-1)' ;
END ;

procedure TFSaisie.E_DATECOMPTABLEEnter(Sender: TObject);
begin
RentreDate:=True ;
end;

procedure TFSaisie.E_DATECOMPTABLEExit(Sender: TObject);
Var Err,i   : integer ;
    DateCpt : TDateTime ;
begin
if csDestroying in ComponentState then Exit ;
Err:=ControleDate(E_DATECOMPTABLE.Text) ;
if Err>0 then BEGIN HPiece.Execute(15+Err,caption,'') ; E_DATECOMPTABLE.SetFocus ; Exit ; END ;
if ((Action=taCreat) and (RevisionActive(StrToDate(E_DATECOMPTABLE.Text)))) then
   BEGIN
   E_DATECOMPTABLE.SetFocus ;
   Exit ;
   END ;
QuelNExo ; DatePiece:=StrToDate(E_DATECOMPTABLE.Text) ;
// détermination du type d'exo
SetTypeExo ;   // FQ 13246 : SBO 30/03/2005

if vh^.CPIFDEFCEGID then          // verif ok
begin
If MethodeGuideDate1 Then
  BEGIN
  if ( appeldatepourguide And (Not M.FromGuide) And (((M.TypeGuide='NOR') and (M.SaisieGuidee)) Or (CodeGuideFromSaisie<>''))) then
    BEGIN
    GS.Enabled:=TRUE ; GS.SynEnabled:=TRUE ; GS.SetFocus ; Exit ;
    END ;
  END ;
end ;
if ((DEV.Code<>V_PGI.DevisePivot) and (Action=taCreat) and (Not Volatile)) Then
   BEGIN
   DateCpt:=DatePiece ;
   DEV.Taux:=GetTaux(DEV.Code,DEV.DateTaux,DateCpt) ;
   If ((DEV.DateTaux<>DateCpt) or (PbTaux(DEV,DateCpt))) Then AvertirPbTaux(DEV.Code,DEV.DateTaux,DateCpt) ;
   END ;
for i:=1 to GS.RowCount-1 do GS.Cells[SA_DateC,i]:=DateToJour(StrToDate(E_DATECOMPTABLE.EditText)) ;
// Recalcul du pied pour solde des comptes à dates FQ13246 SBO 03/08/2005
AffichePied ;

end;

{==========================================================================================}
{=================================== ANALYTIQUES ==========================================}
{==========================================================================================}
Function TFSaisie.AOuvrirAnal ( Lig : integer ; Auto : boolean ) : Boolean ;
Var CGen  : TGGeneral ;
    OBM   : TOBM ;
    XD,XP,Deb{,Cred} : Double ;
    Sens     : integer ;
BEGIN
  Result:=False ;
  CGen:=GetGGeneral(GS,Lig) ;
  if CGen=Nil then exit ;
  if not Ventilable(CGen,0) then Exit ;
  if Auto then
    BEGIN
    OBM:=GetO(GS,Lig) ;
    if OBM=Nil then Exit ;
    Deb:=ValD(GS,Lig) ;
//    Cred:=ValC(GS,Lig) ;
    if Deb<>0
      then Sens:=1
      else Sens:=2 ;
    XD:=MontantLigne(GS,Lig,tsmDevise) ;
    XP:=MontantLigne(GS,Lig,tsmPivot) ;
    if ((GetTotalVentilDev(OBM)=XD) and (GetTotalVentil(OBM)=XP) and (Sens=GetSensVentil(OBM))) then Exit ;
    if GetVentilGeneral(OBM)=CGen.General then
      BEGIN
      RecalculProrataAnalNEW('Y',OBM,0,DEV) ;
      Exit ;
      END ;
    END ;
  Result:=True ;
END ;

procedure TFSaisie.GereAnal(Lig: integer; Auto, Scena: Boolean);
var
  bOuvreAnal: Boolean;
  OBM: TOBM;
  lNumAxe : integer ;
begin
  if not (ctxPCL in V_PGI.PGIContexte) then // verif ok // FQ 15571 : SBO 20/04/2005 Anciennement vh^.CPIFDEFCEGID
  begin
    if Guide and ((ValD(GS, Lig) = 0) and (ValC(GS, Lig) = 0)) then Exit;
  end;
  if AOuvrirAnal(Lig, Auto) then
  begin
    (* HISTORIQUE DES FICHES :
    // FQ 10509
    //SG6 07.02.05 correction FQ 15361
    {JP 12/10/05 : FQ 16837 : Récupération par défaut de la préventilation ou des sections d'attente
                   lors de l'ouverture de la fiche analytique
     JP 25/01/06 : 17303 : ajout de "not Guide" : Le deuxième paramètre '' est en fait le code guide. Mais
                   comme GuideAna n'est pas bétonné, en mettant not Guide, on est sûr de passer dans OuvreAnal}
     // 07/02/2006 SBO : DEV 3946
     *)
    lNumAxe := 0 ;
    if not Guide and Auto then
       begin
       OBM := GetO(GS, Lig);
       if okScenario
         then bOuvreAnal := CAOuvrirVentil( OBM, Scenario, FInfo, lNumAxe )
         else bOuvreAnal := CAOuvrirVentil( OBM, nil, FInfo, lNumAxe ) ;
       end
     else bOuvreAnal := True;

    // Ouvre la fenêtre d'analytique
    if bOuvreAnal then OuvreAnal(Lig, Scena,lNumAxe);
  end;
  CalculDebitCredit;
end;


Procedure TFSaisie.OuvreAnal ( Lig : integer ; Scena : boolean ; vNumAxe : integer ) ;
Var XD,XP,VV 	: Double ;
    C           : integer ;
    Ax       	: String ;
    ArgAna 	: ARG_ANA;
    OBM 	: TOBM;
    lMulQSav    : TQuery ;
BEGIN
  OBM := GetO(GS,Lig);

  // Mise en place des arguments de lancement de la saisie analytique
  ArgAna.QuelEcr        := EcrGen ;
  if PasToucheLigne(Lig) or ( Action = taConsult )
    then ArgAna.Action	:= taConsult
    else ArgAna.Action := Action ;
  if Guide
    then ArgAna.GuideA	:= GuideAnal
    else ArgAna.GuideA	:= '' ;
  ArgAna.DernVentilType	:= DernVentilType ;
  ArgAna.ControleBudget := ((OkScenario) and (Action<>taConsult) and (Scenario.GetValue('SC_CONTROLEBUDGET')='X'));
  ArgAna.ModeConf	:= ModeConf ;
  ArgAna.CC		:= GetGGeneral(GS,Lig) ;
  ArgAna.DEV		:= DEV ;
{$IFDEF OLDVENTIL}
  ArgAna.Verifventil	:= False ;
{$ELSE}
  ArgAna.Verifventil	:= True ;
{$ENDIF OLDVENTIL}
  ArgAna.VerifQte       := ((OkScenario) and (Action<>taConsult) and (Scenario.GetValue('SC_CONTROLEQTE')='X'));
  ArgAna.Info           := FInfo ;
  ArgAna.MessageCompta  := nil ;
  ArgAna.NumLigneDecal  := 0;
  // - Pour les scénario : recherche de l'axe
{$IFDEF OLDVENTIL}
  vNumAxe:=0 ;
{$ELSE  OLDVENTIL}
  if vNumAxe = 0 then
{$ENDIF OLDVENTIL}
    if ((OkScenario) and (Scena) and (Scenario.GetMvt('SC_OUVREANAL')='X')) then    //SG6 FQ 14920 15/11/2004 = et pas <>
      begin
      Ax:=Trim(Scenario.GetMvt('SC_NUMAXE')) ;
      if Length(Ax)<2 then Exit ;
      vNumAxe:=Ord(Ax[2])-48 ;
      end;
  ArgAna.NumGeneAxe			:= vNumAxe ;

  // Backup des montants
  XD:=MontantLigne(GS,Lig,tsmDevise) ;
  XP:=MontantLigne(GS,Lig,tsmPivot) ;

  // Mise en place des Axes dans l'OBM si besoin
  AlloueAxe( OBM ) ;

// Patch AGL qui ajoute l'order by du theMulQ si celui renseigné à l'entrée dans une fiche type TOM...
// FQ 15989 SBO 02/08/2005
lMulQSav := TheMulQ ;
TheMulQ  := nil;

  // Appel de la saisie analytique
  eSaisieAnal(TOB(OBM),ArgAna) ;

// FQ 15989 SBO 02/08/2005
TheMulQ  := lMulQSav ;

  if ArgAna.DernVentilType<>'' then DernVentilType:=ArgAna.DernVentilType ;
  if ArgAna.ModeConf<>ModeConf then
    BEGIN
    if ArgAna.ModeConf>ModeConf then ModeConf:=ArgAna.ModeConf ;
    CalculDebitCredit ;
    END ;
  if Action=taConsult then Exit ;

  // Test montant écriture modifiée en saisie analytique
  if ((XD=GetMontantDev(OBM)) and (XP=GetMontant(OBM))) then Exit ;
  VV:=GetMontantDev(OBM) ;
  if GetSens(OBM)=1
    then C:=SA_Debit
    else C:=SA_Credit ;
  GS.Cells[C,Lig]:=StrS(VV,DECDEV) ;
  CaptureRatio(Lig,False) ;
  TraiteMontant(Lig,True) ;
END ;

{==========================================================================================}
{=================================== ECHEANCES ============================================}
{==========================================================================================}
Procedure TFSaisie.GereEche ( Lig : integer ; OuvreAuto,RempliAuto : Boolean ) ;
Var Cpte : String17 ;
    MR   : String3 ;
    t    : TLettre ;
    FromLigne : Boolean ;
BEGIN
FromLigne:=((OuvreAuto) and (Not RempliAuto)) ;
if AOuvrirEche(Lig,Cpte,MR,OuvreAuto,RempliAuto,t) then
   BEGIN
   if ((OkScenario) and (FromLigne) and (t=MultiEche)) then
      if Scenario.GetMvt('SC_OUVREECHE')<>'X' then RempliAuto:=True ;
   OuvreEche(Lig,Cpte,MR,RempliAuto,t) ;
   END ;
{ GP le 25/08/98 : }
CalculDebitCredit ;
END ;

Procedure TFSaisie.EcheGuideRegle ( Lig : integer ; Var T : T_MODEREGL ) ;
Var O,OEcr : TOBM ;
BEGIN
if Lig>TGUI.Count then Exit ;
O:=TOBM(TGUI[Lig-1]) ; if O=Nil then Exit ;
OEcr:=GetO(GS,Lig) ; if OEcr=Nil then Exit ;
T.TabEche[1].ModePaie:=O.GetMvt('EG_MODEPAIE') ; LastPaie:=T.TabEche[1].ModePaie ;
T.TabEche[1].DateEche:=VarAsType(O.GetMvt('EG_DATEECHEANCE'),VarDate) ;
OEcr.PutMvt('E_DATEECHEANCE',T.TabEche[1].DateEche) ;
T.TabEche[1].DateValeur:=VarAsType(O.GetMvt('EG_DATEECHEANCE'),VarDate) ;
T.TabEche[1].MontantD:=MontantLigne(GS,Lig,tsmDevise) ;
T.TabEche[1].MontantP:=MontantLigne(GS,Lig,tsmPivot) ;
T.TabEche[1].Pourc:=100.0 ;
T.NbEche:=1 ;
{#TVAENC}
if OuiTvaLoc then GereTvaEncNonFact(Lig) ;
END ;

Procedure TFSaisie.EcheLettreCheque ( Lig : integer ; Var T : T_MODEREGL ) ;
Var O : TOBM ;
BEGIN
O:=GetO(GS,Lig) ; if O=Nil then Exit ;
T.TabEche[1].ModePaie:=O.GetMvt('E_MODEPAIE') ; LastPaie:=T.TabEche[1].ModePaie ;
T.TabEche[1].DateEche:=O.GetMvt('E_DATEECHEANCE') ;
T.TabEche[1].DateValeur:=T.TabEche[1].DateEche ;
T.TabEche[1].MontantD:=MontantLigne(GS,Lig,tsmDevise) ;
T.TabEche[1].MontantP:=MontantLigne(GS,Lig,tsmPivot) ;
T.TabEche[1].Pourc:=100.0 ;
T.NbEche:=1 ;
{#TVAENC}
if OuiTvaLoc then GereTvaEncNonFact(Lig) ;
END ;

Function TFSaisie.IncNumLotSaisie ( Lig : integer ) : String ;
Var O : TOBM ;
    CGen : TGGeneral ;
    NatG : String ;
    Cli  : Boolean ;
    sRac,sNum,sNewNum : String ;
BEGIN
Result:='' ;
{$IFDEF SPEC350}
Exit ;
{$ENDIF}
if Not OkJalEffet then Exit ;
O:=GetO(GS,Lig) ; if O=Nil then Exit ;
CGen:=GetGGeneral(GS,Lig) ; if CGen=Nil then Exit ;
NatG:=CGen.NatureGene ;
Cli:=((NatG='COC') or (NatG='TID')) ;
if Not Cli then if ((NatG<>'COF') and (NatG='TIC')) then Exit ;
if Cli then
   BEGIN
   sRac:=GetParamSoc('SO_CPSOUCHETRACLI') ;
   sNum:=IntToStr(GetParamSoc('SO_CPNUMTRACLI')-1) ;
   sNewNum:=IncNumLotTraCHQ(sNum) ;
   SetParamsoc('SO_CPNUMTRACLI',ValeurI(sNewNum)+1) ;
   Result:=sRac+' '+sNewNum ;
   END else
   BEGIN
   sRac:=GetParamSoc('SO_CPSOUCHETRAFOU') ;
   sNum:=IntToStr(GetParamSoc('SO_CPNUMTRAFOU')-1) ;
   sNewNum:=IncNumLotTraCHQ(sNum) ;
   SetParamsoc('SO_CPNUMTRAFOU',ValeurI(sNewNum)+1) ;
   Result:=sRac+' '+sNewNum ;
   END ;
END ;

Procedure TFSaisie.RemonteValeurTraCHQ ( sTraCHQ : String ; Lig : integer ) ;
Var O : TOBM ;
    CGen : TGGeneral ;
    i    : integer ;
BEGIN
{$IFDEF SPEC350}
Exit ;
{$ENDIF}
if Not OkJalEffet then Exit ;
if sTraCHQ='' then Exit ;
O:=GetO(GS,Lig) ; if O=Nil then Exit ;
CGen:=GetGGeneral(GS,Lig) ; if CGen=Nil then Exit ;
if CGen.General<>SAJAL.Treso.Cpt then Exit ;
for i:=Lig-1 downto 1 do
    BEGIN
    CGen:=GetGGeneral(GS,i) ; if CGen=Nil then Continue ;
    O:=GetO(GS,i) ; if O=Nil then Continue ;
    if CGen.General=SAJAL.Treso.Cpt then Break ;
    if ((CGen.NatureGene<>'COC') and (CGen.NatureGene<>'COF') and
        (CGen.NatureGene<>'TID') and (CGen.NatureGene<>'TIC')) then Continue ;
    if O.GetMvt('E_NUMTRAITECHQ')='' then O.PutMvt('E_NUMTRAITECHQ',sTraCHQ) ;
    END ;
END ;

Procedure TFSaisie.GereMonoEche ( Var OO : T_ModeRegl ; OkInit,AutoPourEncaDeca : boolean ; Lig : integer ; ActionEche : TActionFiche ) ;
Var X : T_MONOECH ;
    OkOk : Boolean ;
    O    : TOBM ;
BEGIN
//OkOk:=FALSE ;
O:=GetO(GS,Lig) ;
With OO do
   BEGIN
   X.DateMvt:=DateMvt(Lig) ;
   {b FP 30/01/2006 FQ17278 Reprend les infos du compte tiers}
   if OkInit then
     begin
     CalculModeRegle(OO,False);
     X.ModePaie:=OO.TabEche[1].ModePaie ;
     X.DateEche:=OO.TabEche[1].DateEche ;
     end
   else
     begin
   {e FP 30/01/2006 FQ17278}
     if (LastPaie<>'') and (ActionEche = taCreat)
       then X.ModePaie:=LastPaie
       else X.ModePaie:=TabEche[1].ModePaie ;
     if (LastEche>0) and (ActionEche = taCreat)
       then X.DateEche:=LastEche
       else X.DateEche:=TabEche[1].DateEche ;
     end;
   X.Cat:='' ; X.Treso:=False ; X.OkInit:=OkInit ;
   if OkInit
     then TabEche[1].DateValeur:=CalculDateValeur(DateMvt(Lig),MonoEche,GS.Cells[SA_Gen,Lig]) ;
   if (LastVal>0) and (ActionEche = taCreat)
     then X.DateValeur:=LastVal
     else X.DateValeur:=TabEche[1].DateValeur ;
   X.NumTraiteCHQ:='' ;
{$IFNDEF SPEC350}
   if O<>Nil then X.NumTraiteCHQ:=O.GetMvt('E_NUMTRAITECHQ') ;
   if ((X.NumTraiteCHQ='') and (LastTraiteCHQ<>'') and (OkBQE))
     then X.NumTraiteCHQ:=LastTraiteCHQ ;
   if ((X.NumTraiteCHQ='') and (OkJalEffet) and (O<>Nil) and (O.GetMvt('E_GENERAL')=SAJAL.Treso.Cpt))
     then X.NumTraiteCHQ:=IncNumLotSaisie(Lig) ;
{$ENDIF}
   X.OkVal:=True ; X.Action:=Action ;

   If AutoPourEncaDeca Then
      If (X.ModePaie='') And (Not Guide)
        Then OkOk:=SaisirMonoEcheance(X)
        Else OkOk:=TRUE
   Else
     OkOk:=SaisirMonoEcheance(X) ;

   if OkOk then
      BEGIN
      TabEche[1].DateEche:=X.DateEche ;
      TabEche[1].ModePaie:=X.ModePaie ;
      NbEche:=1 ;
      TabEche[1].Pourc:=100.0 ;
      TabEche[1].MontantP:=TotalAPayerP ;
      TabEche[1].MontantD:=TotalAPayerD ;
      TabEche[1].DateValeur:=X.DateValeur ;
      if TabEche[1].EtatLettrage=''
        then TabEche[1].EtatLettrage:='AL' ;
      LastPaie:=X.ModePaie ; LastEche:=X.DateEche ; LastVal:=X.DateValeur ;
{$IFNDEF SPEC350}
      if O<>Nil
        then O.PutMvt('E_NUMTRAITECHQ',X.NumTraiteCHQ) ;
      if X.NumTraiteCHQ<>''
        then LastTraiteCHQ:=X.NumTraiteCHQ ;
      RemonteValeurTraCHQ(X.NumTraiteCHQ,Lig) ;
{$ENDIF}
      END ;
   END ;
END ;

Procedure TFSaisie.OuvreEche ( Lig : integer ; Cpte : String17 ; MR : String3 ; RempliAuto : boolean ; t : TLettre ) ;
Var OMODR  : TMOD ;
    i      : integer ;
    CGen   : TGGeneral ;
    OBM    : TOBM ;
    OkInit,AutoPourEncaDeca,OkMultiEche : Boolean ;
BEGIN
OMODR:=TMOD(GS.Objects[SA_NumP,Lig]) ; if OMODR=Nil then Exit ;
CGen:=GetGGeneral(GS,Lig) ; if CGen=Nil then Exit ;
AutoPourEncaDeca:=FALSE ;
{$IFDEF SPEC350}
OkMultiEche:=((Not OkBqe)) ;
{$ELSE}
OkMultiEche:=((Not OkBqe) and (Not OkJalEffet)) ;
{$ENDIF}
With OMODR.MODR do
   BEGIN
   // Modif 02/11/2004 SBO : Pb d'affectation de l'état de lettrage en modeLC si l'on passe l'action du MODR en taConsult alors qu'on est en mode taCreat... FQ14892
   if ((Self.Action=taConsult) (*or (ModeLC)*)) then OMODR.MODR.Action:=taConsult else
    if ((Self.Action=taModif) and (PasToucheLigne(Lig))) then OMODR.MODR.Action:=taConsult else
      BEGIN
      if Not ((TotalAPayerP=0) or (NbEche=0)) then OMODR.MODR.Action:=taModif else OMODR.MODR.Action:=taCreat;
      END ;
   OkInit:=(OMODR.MODR.Action=taCreat) ;
   TotalAPayerD:=MontantLigne(GS,Lig,tsmDevise) ;
   TotalAPayerP:=MontantLigne(GS,Lig,tsmPivot) ;
   CodeDevise:=DEV.Code ; TauxDevise:=DEV.Taux ;
   Quotite:=DEV.Quotite ; Decimale:=DECDEV ;
   Symbole:=DEV.Symbole ; DateFact:=DateMvt(Lig) ;
// GP REGL
   DateFactExt:=DateFact ;
   OBM:=GetO(Gs,Lig) ;
   If OBM<>NIL Then
      BEGIN
      DateFactExt:=OBM.GetMvt('E_DATEREFEXTERNE') ;
      If DateFactExt=IDate1900 Then DateFactExt:=DateFact ;
      END ;
   DateBL:=DateFact     ; Aux:=Cpte ;
   ModeInitial:=MR      ;
   if OMODR.MODR.Action=taCreat then TabEche[1].DateEche:=DateMvt(Lig) ;
   END ;
if RempliAuto then
   BEGIN
   if (M.LeGuide='') or ((M.TypeGuide<>'DEC') and (M.TypeGuide<>'ENC')) then
      BEGIN
      if ((CGen.Collectif) and (Ventilable(CGen,0))) then OkMultiEche:=False ;
      if LienS1S3 then OkMultiEche:=False ;
      CalculModeRegle(OMODR.MODR,OkMultiEche) ;
      END
   else
      BEGIN
      if ((M.TypeGuide='DEC') Or (M.TypeGuide='ENC')) and (Not Guide) then AutoPourEncaDeca:=True else EcheGuideRegle(Lig,OMODR.MODR) ;
      END ;
   END else
   BEGIN
   Case t of
      MultiEche : if ((Not OkBQE) and (Not OkJalEffet)) // and (Not LienS1S3)) FQ 13099
                     then SaisirEcheance(OMODR.MODR)
                     else GereMonoEche(OMODR.MODR,OkInit,AutoPourEncaDeca,Lig,OMODR.MODR.Action) ;
      MonoEche  : GereMonoEche(OMODR.MODR,OkInit,AutoPourEncaDeca,Lig,OMODR.MODR.Action) ;
      END ;
   Revision:=False ;
   END ;
if AutoPourEncaDeca then GereMonoEche(OMODR.MODR,OkInit,AutoPourEncaDeca,Lig,OMODR.MODR.Action) ;
With OMODR.MODR do if OkInit then for i:=1 to MaxEche do With TabEche[i] do
     BEGIN
     DateRelance:=IDate1900 ; NiveauRelance:=0 ;
     if t=MultiEche then
        BEGIN
        if DateValeur<=2 then DateValeur:=CalculDateValeur(DateFact,MultiEche,GS.Cells[SA_Gen,Lig]) ;
        if EtatLettrage='' then EtatLettrage:='AL' ;
        if LettrageDEV='' then LettrageDEV:='-' ;
        END ;
     {Reste géré par Echeance.pas}
     END ;
END ;

Function TFSaisie.AOuvrirEche ( Lig : integer ; Var Cpte : String17 ; Var MR : String3 ;
                                Var OuvreAuto,RempliAuto : boolean ; Var t : TLettre) : Boolean ;
Var CGen     : TGGeneral ;
    CAux     : TGTiers ;
    OMODR    : TMOD ;
    XD,XP : Double ;
    OG       : TOBM ;
    St,StR   : String ;
BEGIN
  Result := False ;
  Cpte   := '' ;
  MR     := '' ;
  t      := NonEche ;

  CGen   := GetGGeneral(GS,Lig) ;
  if CGen=Nil then exit ;

  CAux   := GetGTiers(GS,Lig) ;
  if CGen.Collectif then
    begin
    if CAux=Nil then Exit ;
    if Not CAux.Lettrable then Exit ;
    Cpte:=CAux.Auxi ;
    MR:=CAux.ModeRegle ;
    if Ventilable(CGen,0)
      then t:=MonoEche
      else t:=MultiEche ;
    end
  else
    begin
    t:=Lettrable(CGen) ;
    if t=NonEche then Exit ;
    if Ventilable(CGen,0)
      then t:=MonoEche ;
    Cpte := CGen.General ;
    MR   := CGen.ModeRegle ;
    // Ajout test pour comptes divers lettrables... FQ 16136 SBO 09/08/2005
    if CGen.NatureGene = 'DIV' then
      begin
      OuvreAuto  := False ;
      RempliAuto := True ;
  {    OMODR := TMOD(GS.Objects[SA_NumP,Lig]) ;
      if OMODR<>Nil then
          if OMODR.MODR.Action<>taCreat then
          RempliAuto := False ;}
      MR         := GetParamSoc('SO_GCMODEREGLEDEFAUT') ;
      end ;

    end ;

  if Guide then
    begin
    OG := TOBM(TGUI[Lig-1]) ;
    if OG = Nil then Exit ;
    StR  := OG.GetMvt('EG_MODEREGLE') ;
    St   := OG.GetMvt('EG_ARRET') ;
    if ( (M.LeGuide='') or ((M.TypeGuide<>'DEC') and (M.TypeGuide<>'ENC')) )
       and ( CGen.NatureGene <> 'DIV' ) then     // Ajout test pour comptes divers lettrables... FQ 16136 SBO 09/08/2005
      begin
      if StR<>'' then
        begin
        MR:=StR ;
        if St[7]='X'
          then OuvreAuto  := True
          else RempliAuto := True ;
        end ;
      end
    else
      begin
      // guides de règlement
      OuvreAuto  := False ;
      RempliAuto := True ;
      end ;
    end ;

  if OuvreAuto then
    begin
    OMODR := TMOD(GS.Objects[SA_NumP,Lig]) ;
    if OMODR=Nil then Exit ;
    XD := MontantLigne(GS,Lig,tsmDevise) ;
    XP := MontantLigne(GS,Lig,tsmPivot) ;
    if (OMODR.MODR.TotalAPayerD=XD) and (OMODR.MODR.TotalAPayerP=XP) then Exit ;
    if OMODR.MODR.Aux=Cpte then
      begin
      RecalculProrataEche(OMODR.MODR,XP,XD) ;
      Exit ;
      end ;
    end ;

  Result := True ;

END ;

{==========================================================================================}
{=================================== Contrôles ============================================}
{==========================================================================================}
Procedure TFSaisie.TripoteYATP ;
Var OBM : TOBM ;
    Lig,k : integer ;
    TM  : TMOD ;
//    EtatMod : String ;
BEGIN
if YATP<>yaNL then Exit ;
for Lig:=1 to GS.RowCount-2 do
    BEGIN
    OBM:=GetO(GS,Lig) ; if OBM=Nil then Break ;
    if OBM.GetMvt('E_AUXILIAIRE')='' then Continue ;
    if OBM.GetMvt('E_ECHE')<>'X' then Continue ;
    if OBM.GetMvt('E_TIERSPAYEUR')='' then Continue ;
    if OBM.GetMvt('E_PIECETP')='' then Continue ;
    OBM.PutMvt('E_LETTRAGE','') ; OBM.PutMvt('E_ETATLETTRAGE','AL') ;
    TM:=TMOD(GS.Objects[SA_NumP,Lig]) ;
    if TM<>Nil then
       BEGIN
       for k:=1 to TM.ModR.NbEche do
           BEGIN
           TM.ModR.TabEche[k].Couverture    :=0 ;
           TM.ModR.TabEche[k].CouvertureDev :=0 ;
           TM.ModR.TabEche[k].CodeLettre    :='' ;
           TM.ModR.TabEche[k].LettrageDev   :='-' ;
           TM.ModR.TabEche[k].DatePaquetMax :=OBM.GetMvt('E_DATECOMPTABLE') ;
           TM.ModR.TabEche[k].DatePaquetMin :=OBM.GetMvt('E_DATECOMPTABLE') ;
           TM.ModR.TabEche[k].EtatLettrage  :='AL' ;
           END ;
       END ;
    END ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 25/08/2005
Modifié le ... :   /  /    
Description .. : - FB 16397 - LG  - 25/08/2005 - test de la revision 
Suite ........ : uniquement pour l'exercice en cours
Mots clefs ... : 
*****************************************************************}
function TFSaisie.PasToucheLigne ( Lig : integer ) : boolean ;
Var OBM : TOBM ;
    TM  : TMOD ;
    k   : integer ;
    EtatLig,RefP,EtatMod,StEtatTva : String ;
    ExisteL,XEditeTva,ExportePCL : boolean ;
    CAux     : TGTiers ;
    CGen     : TGGeneral ;
    CptFerme : Boolean ;
    lBoVisa  : boolean ;
BEGIN
Result:=False ;
OBM:=GetO(GS,Lig) ; ExisteL:=False ; XEditeTva:=False ; if OBM=Nil then Exit ;
EtatLig:=OBM.GetMvt('E_ETATLETTRAGE') ;
RefP:=OBM.GetMvt('E_REFPOINTAGE') ;
ExportePCL:=False ;
// ExportePCL:=(OBM.GetMvt('E_XXXX')='ZZZZ')  ;
if OuiTvaLoc then
   BEGIN
   StEtatTva:=OBM.GetMvt('E_EDITEETATTVA') ; if StEtatTva='#' then XEditeTva:=True ;
   END ;
TM:=TMOD(GS.Objects[SA_NumP,Lig]) ;
if TM<>Nil then
   BEGIN
   for k:=1 to TM.ModR.NbEche do
       BEGIN
       EtatMod:=TM.ModR.TabEche[k].EtatLettrage ;
       if ((EtatMod='PL') or (EtatMod='TL')) then BEGIN ExisteL:=True ; Break ; END ;
       END ;
   END ;

  // SBO 19/08/2004 FQ 13290 modif interdite sur les comptes, journaux fermés
  CptFerme := False ;
  CGen     := GetGGeneral(GS,Lig) ;
  if CGen <> Nil then CptFerme := CGen.Ferme ;
  CAux     := GetGTiers(GS,Lig) ;
  if CAux <> Nil then CptFerme := CptFerme or CAux.Ferme ;
  if (SAJAL<>nil) then CptFerme := CptFerme or SAJAL.OkFerme ;

  lBoVisa := false ;
  if FInfo.LoadCompte( GS.Cells[SA_Gen,Lig] ) then
   lBoVisa := ( FInfo.Compte_GetValue('G_VISAREVISION') = 'X' ) and ( FInfo.TypeExo = teEnCours ) ;

  {JP 28/07/05 : FQ 15124 : ON ne peut pas modifier les lignes TTC des écritures de GESCOM}
  Result := (OBM.GetMvt('E_REFGESCOM') <> '') and (OBM.GetMvt('E_TYPEMVT') = 'TTC');
  {JP 26/06/07 : FQ TRESO 10491 : on vérouille les flux originaires de la Tréso}
  Result := Result and (OBM.GetMvt('E_QUALIFORIGINE') <> QUALIFTRESO);

  Result := Result or
            (EtatLig='PL') or (EtatLig = 'TL')
              or (RefP<>'') or (ExisteL)
              or (XEditeTva) or (ExportePCL)
              or CptFerme or lBoVisa  ;
  // SBO 19/08/2004 FQ 13290 modif interdite sur les comptes, journaux fermés

END ;

function TFSaisie.ControleTVAOK : boolean ;
Var LigHT,NbH : integer ;
    Okok : boolean ;
BEGIN
Result:=True ;
if Not OkScenario then Exit ; if PasModif then Exit ;
LigHT:=0 ; NbH:=0 ; //Okok:=True ;
Repeat
 LigHT:=TrouveLigHT(GS,LigHT,True) ; Okok:=True ;
 if LigHT>0 then BEGIN Inc(NbH) ; Okok:=ClickControleTVA(True,LigHT) ; END ;
Until ((LigHT<=0) or (Not Okok)) ;
if ((NbH<=0) or (Not Okok)) then Result:=(HTVA.Execute(11,caption,'')=mrYes) ;
END ;

procedure TFSaisie.ControleLignes ;
Var Lig : integer ;
BEGIN
for Lig:=1 to GS.RowCount-1 do if Not EstRempli(GS,Lig) then DefautLigne(Lig,True) ;
END ;

procedure TFSaisie.AjusteLigne ( Lig : integer ; AvecM : boolean ) ;
Var Cpte              : String17 ;
    MR                : String3 ;
    t                 : TLettre ;
    Ouv,Remp          : boolean ;
BEGIN
AlimObjetMvt(Lig,AvecM) ; Remp:=False ; Ouv:=True ;
if AOuvrirEche(Lig,Cpte,MR,Ouv,Remp,t) then ;
if AOuvrirAnal(Lig,True) then ;
END ;

procedure TFSaisie.ErreurSaisie ( Err : integer ) ;
BEGIN
if Err<100 then
   BEGIN
   HLigne.Execute(Err-1,caption,'') ;
   END else if Err<200 then
   BEGIN
   HPiece.Execute(Err-101,caption,'') ;
   END ;
END ;

Function TFSaisie.PossibleRecupNum : Boolean ;
Var MM : String17 ;
    Facturier : String3 ;
    DD : TDateTime ;
BEGIN
Result:=True ;
if H_NUMEROPIECE_.Visible then Exit ;
if ( M.Simul <> 'N' ) and ( M.Simul <> 'I' ) // Modif IFRS
  then Facturier:=SAJAL.CompteurSimul
  else Facturier:=SAJAL.CompteurNormal ;
DD:=StrToDate(E_DATECOMPTABLE.Text) ;
if GetNum(EcrGen,Facturier,MM,DD)=NumPieceInt+1 then Exit ;
Result:=False ;
END ;

Function TFSaisie.PasModif : Boolean ;
BEGIN
PasModif:=((Action=taConsult) or (ModeForce<>tmNormal)) ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 31/03/2003
Modifié le ... :   /  /    
Description .. : CA - 31/03/2003 - controle des doublons sur le dossiers 
Suite ........ : complet et non plus l'exercice
Mots clefs ... : 
*****************************************************************}
Function TFSaisie.TraiteDoublon ( Lig : integer ) : Boolean ;
Var SQL,Gene,Aux,Champ,Val : String ;
    O,OZoom : TOBM ;
    QQ : TQuery ;
    XX : RMVT ;
BEGIN
Result:=True ; OZoom:=Nil ;
if M.Simul<>'N' then Exit ;
O:=GetO(GS,Lig) ; if O=Nil then Exit ;
Aux:=O.GetMvt('E_AUXILIAIRE') ; if Aux='' then Exit ;
Gene:=O.GetMvt('E_GENERAL') ; if Gene='' then Exit ;
// SBO 05/07/2007 : gestion pb DB2 avec champ blocnote à positionner à la fin
SQL:='SELECT ' + GetSelectAll('E', True) + ', E_BLOCNOTE FROM ECRITURE WHERE E_GENERAL="'+Gene+'" AND E_AUXILIAIRE="'+Aux+'" AND E_QUALIFPIECE="N" ' ;
// BPY le 11/03/2004 => fiche de bug n°13213 : verification sur uniquement les piece de typre facture.
SQL := SQL + 'AND E_NATUREPIECE in ("AC","AF","FC","FF")';
// Fin BPY
Champ:='' ; Val:='' ;
if VH^.CPChampDoublon='RFI' then BEGIN Champ:='E_REFINTERNE' ; Val:=O.GetMvt('E_REFINTERNE') ; END else
if VH^.CPChampDoublon='LIB' then BEGIN Champ:='E_LIBELLE'    ; Val:=O.GetMvt('E_LIBELLE')    ; END else
if VH^.CPChampDoublon='RFX' then BEGIN Champ:='E_REFEXTERNE' ; Val:=O.GetMvt('E_REFEXTERNE') ; END else
if VH^.CPChampDoublon='RFL' then BEGIN Champ:='E_REFLIBRE'   ; Val:=O.GetMvt('E_REFLIBRE')   ; END ;
if ((Champ='') or (Val='')) then Exit ;
// BPY le 15/10/2004 => pb de double quote !
SQL:=SQL+' AND '+Champ+'="'+CheckdblQuote(Val)+'"' ;
// Fin BPY !
QQ:=OpenSQL(SQL,True) ;
if Not QQ.EOF then BEGIN OZoom:=TOBM.Create(EcrGen,'',False) ; OZoom.ChargeMvt(QQ) ; END ;
Ferme(QQ) ;
if OZoom<>Nil then
   BEGIN
   if HPiece.Execute(44,Caption,'')<>mrYes then
      BEGIN
      Result:=False ;
      if HPiece.Execute(45,Caption,'')=mrYes then BEGIN XX:=OBMToIdent(OZoom,False) ; LanceSaisie(Nil,taConsult,XX) ; END ;
      END ;
   OZoom.Free ;
   END ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Piot
Créé le ...... : 21/02/2005
Modifié le ... :   /  /    
Description .. : - LG - 21/02/2005 - affectation de NumA
Mots clefs ... : 
*****************************************************************}
Function TFSaisie.LigneCorrecte ( Lig : integer ; Totale,Alim : boolean ) : Boolean ;
Var CGen  : TGGeneral ;
    CAux  : TGTiers ;
    Err,NumA,NumL : integer ;
    OMOD  : TMOD ;
    OBM   : TOBM ;
    OBA   : TOBM ;
    TotD,TotP{,TotE} : Double ;
    Sens,ii{,SensC}  : byte ;
    Col,premier_axe {SG6 16/12/04}      : integer ;
BEGIN
NumA := 0 ;
if PasModif then BEGIN Result:=True ; Exit ; END ;
{Result:=False ;} Err:=0 ; Sens:=1 ; if ValC(GS,Lig)<>0 then Sens:=2 ;
CGen:=GetGGeneral(GS,Lig) ; Col:=-1 ;
if ((Err=0) and (CGen=Nil)) then BEGIN Err:=1 ; Col:=SA_Gen ; END ;
if ((Err=0) and (CGen.General<>GS.Cells[SA_Gen,Lig])) then BEGIN Err:=1 ; Col:=SA_Gen ; END ;
CAux:=GetGTiers(GS,Lig) ;
if ((Err=0) and (Not CGen.Collectif) and (CAux<>Nil)) then BEGIN Err:=2 ; Col:=SA_Gen ; END ;
if ((Err=0) and (CAux<>Nil)) then if CAux.Auxi<>GS.Cells[SA_Aux,Lig] then
   BEGIN
   if GS.Cells[SA_Aux,Lig]<>'' then Err:=2 else Err:=3 ;
   Col:=SA_Aux ;
   END ;
if ((Err=0) and (CGen.Collectif) and (CAux=Nil)) then BEGIN Err:=3 ; Col:=SA_Aux ; END ;
if Err=0 then
   BEGIN
   ii:=EstInterdit(SAJAL.CompteInterdit,CGen.General,Sens) ;
   Case ii of 1 : Err:=18 ; 2 : Err:=19 ; 3 : Err:=4 ; END ;
   if Err>0 then Col:=SA_Gen ;
   END ;
if ((Err=0) and (Not M.ANouveau) and (EstOuvreBil(CGen.General))) then Err:=20 ;
if ((Err=0) and (M.ANouveau) and (EstChaPro(CGen.NatureGene))) then Err:=21 ;
if not (ctxPCL in V_PGI.PGIContexte) then // verif ok // FQ 15571 : SBO 20/04/2005 Anciennement vh^.CPIFDEFCEGID
  begin
  If Not Guide Then
    BEGIN
    if ((Err=0) and (ValD(GS,Lig)=0) and (ValC(GS,Lig)=0)) then
       if Not EstCptEcart(CGen.General) then
       BEGIN
       Err:=5 ; if QuelSens(E_NATUREPIECE.Value,CGen.NatureGene,CGen.Sens)=1 then Col:=SA_Debit else Col:=SA_Credit ;
       END ;
    END
  end
else
  if ((Err=0) and (ValD(GS,Lig)=0) and (ValC(GS,Lig)=0)) then
     if Not (EstCptEcart(CGen.General) or
             (EstChaPro(CGen.NatureGene)) and (VH^.PaysLocalisation=CodeISOES)) then //XVI 24/02/2005
     BEGIN
     Err:=5 ; if QuelSens(E_NATUREPIECE.Value,CGen.NatureGene,CGen.Sens)=1 then Col:=SA_Debit else Col:=SA_Credit ;
     END ;

if not (ctxPCL in V_PGI.PGIContexte) then // verif ok // FQ 15571 : SBO 20/04/2005 Anciennement vh^.CPIFDEFCEGID
  begin
  If Not Guide Then
    BEGIN
    if ((Err=0) and (ModeDevise) and (ValeurPivot(GS.Cells[SA_Debit,Lig],DEV.Taux,Dev.Quotite)=0)
                and (ValeurPivot(GS.Cells[SA_Credit,Lig],DEV.Taux,Dev.Quotite)=0)) then
        BEGIN
        Err:=6 ; if Sens=1 then Col:=SA_Debit else Col:=SA_Credit ;
        END ;
    END
  end
else
  if ((Err=0) and (ModeDevise) and (ValeurPivot(GS.Cells[SA_Debit,Lig],DEV.Taux,Dev.Quotite)=0)
              and (ValeurPivot(GS.Cells[SA_Credit,Lig],DEV.Taux,Dev.Quotite)=0)) then
      BEGIN
      Err:=6 ; if Sens=1 then Col:=SA_Debit else Col:=SA_Credit ;
      END ;

if ((Err=0) and
    (MontantLigne(GS,Lig,tsmDevise)<>0) and
    (MontantLigne(GS,Lig,tsmPivot)=0) and
    (Not EstCptEcart(CGen.General))) then
    BEGIN
    Err:=6 ;
    if Sens=1
      then Col:=SA_Debit
      else Col:=SA_Credit ;
    END ;

if ((Err=0) And (CGen.Collectif) and (Not NatureGenAuxOk(CGen.NatureGene,CAux.NatureAux))) Then BEGIN Err:=7 ; Col:=SA_Aux ; END ;
if ((Err=0) and (not NaturePieceCompteOk(E_NATUREPIECE.Value,CGen.NatureGene))) then BEGIN Err:=8 ; Col:=SA_Gen ; END ;
if ((Err=0) and (Not VH^.MontantNegatif)) then
   if Not EstCptEcart(CGen.General) then
   BEGIN
   if ((ValD(GS,Lig)<0) or (ValC(GS,Lig)<0)) then BEGIN Err:=9 ; Col:=SA_Debit ; if Sens=2 then Col:=SA_Credit ; END ;
   END ;
if Err=0 then
   BEGIN
   if ((Abs(MontantLigne(GS,Lig,tsmPivot))<VH^.GrpMontantMin) or (Abs(MontantLigne(GS,Lig,tsmPivot))>VH^.GrpMontantMax))
      then BEGIN Err:=23 ; Col:=SA_Debit ; if Sens=2 then Col:=SA_Credit ; END ;
   END ;
// Verif date échéances
if (Err=0) and Totale and not DateEchesOk( Lig ) then
  begin
  Err := 32 ;
  end ;
if ((Err=0) and (Totale)) then
   BEGIN
   if TMOD(GS.Objects[SA_NumP,Lig])<>Nil then
      BEGIN
      OMOD:=TMOD(GS.Objects[SA_NumP,Lig]) ;
      if TotDifferent(MontantLigne(GS,Lig,tsmDevise),OMOD.MODR.TotalAPayerD) then BEGIN Err:=10 ; Col:=SA_Debit ; if Sens=2 then Col:=SA_Credit ; END ;
      END ;
   if ((Err=0) and (GetO(GS,Lig).Detail.Count>0)) then
      BEGIN
      OBM := GetO(GS,Lig) ;
      if Err=0 then
      //Mode Croisaxe  - SG6
        if VH^.AnaCroisaxe then
          begin
          if Ventilable(CGen,0) then   // FQ 16433 : SBO 04/08/2005
            begin
              premier_axe := RecherchePremDerAxeVentil.premier_axe;
              if premier_axe <> 0 then
                if OBM.Detail[premier_axe - 1].Detail.Count = 0
                  then Err := 11 + NumA;
            end
          end
      //Mode classique de ventilation
      else

      for NumA:=1 to MaxAxe do
      	if Ventilable(CGen,NumA) then
          BEGIN
          if OBM.Detail[NumA-1].Detail.Count=0 then Err:=11+NumA
          else
            BEGIN
            TotD:=0 ; TotP:=0 ; {TotE:=0} ;
            for NumL:=0 to OBM.Detail[NumA-1].Detail.Count-1 do
              BEGIN
              OBA := TOBM(OBM.Detail[NumA-1].Detail[NumL]);
              TotD:=TotD + GetMontantDev(OBA) ;
              TotP:=TotP + GetMontant(OBA) ;
              END ;
            if Arrondi(TotD-MontantLigne(GS,Lig,tsmDevise),DEV.Decimale)<>0 then BEGIN Err:=11+NumA ; Break ; END ;
            if Arrondi(TotP-MontantLigne(GS,Lig,tsmPivot),V_PGI.OkDecV)<>0 then BEGIN Err:=11+NumA ; Break ; END ;
            END ;
         if Err<>0 then BEGIN Col:=SA_Debit ; if Sens=2 then Col:=SA_Credit ; END ;
         END ;
      END ;
   END ;
// controles de ventilation Err=12..16
Result:=(Err=0) ;
if ( (Err=0) and (VH^.CPControleDoublon) and
     (Totale) and (Action=taCreat) and (EstJalFact(E_JOURNAL.Value)) and
//     (M.TypeGuide='') and (M.LeGuide='') and  // Modif SBO : Fiche 10435 : Active le contrôle des doublons en saisie guidée
     (Not ModeLC) and (VH^.CPChampDoublon<>''))
     // BPY le 11/03/2004 => Fiche 13213
     and ((E_NATUREPIECE.value = 'AC') or (E_NATUREPIECE.value = 'AF') or (E_NATUREPIECE.value = 'FC') or (E_NATUREPIECE.value = 'FF')) then
     // fin BPY
     BEGIN
     Result:=TraiteDoublon(Lig) ;
     if Not Result then
        BEGIN
        if Col>0 then BEGIN if GS.CanFocus then GS.SetFocus ; GS.Col:=Col ; END ;
        Exit ;
        END ;
     END ;
if Not Result then
   BEGIN
   ErreurSaisie(Err) ;
   if Col>0 then BEGIN if GS.CanFocus then GS.SetFocus ; GS.Col:=Col ; END ;
   END else
   BEGIN
   if Alim then AlimObjetMvt(Lig,TRUE) ;
   END ;
END ;

Function TFSaisie.Equilibre : Boolean ;
BEGIN
Result:=ArrS(SI_TotDS-SI_TotCS)=0 ;
END ;

Function TFSaisie.ControleRIB : Boolean ;
Var Lig  : integer ;
    Okok : boolean ;
    O    : TOBM ;
    RIB  : String ;
BEGIN
Result:=True ; Okok:=True ;
if Not SuiviMP then
   BEGIN
   if Action<>taCreat then Exit ;
   if ((M.TypeGuide<>'ENC') and (M.TypeGuide<>'DEC')) then Exit ;
   END ;
if Not M.ExportCFONB then Exit ;
for Lig:=1 to GS.RowCount-2 do
    BEGIN
    O:=GetO(GS,Lig) ; if O=Nil then Break ;
    if O.GetMvt('E_AUXILIAIRE')='' then
      If (O.GetMvt('E_GENERAL')<>'TIC') And (O.GetMvt('E_GENERAL')<>'TID') Then Continue ;
    if Not SuiviMP then
       BEGIN
       if O.LC.Count<=0 then Continue ;
       END else
       BEGIN
       if O.GetMvt('E_TRACE')='' then Continue ;
       END ;
    RIB:=Trim(O.GetMvt('E_RIB')) ; if RIB='' then Okok:=False ;
    if Not Okok then Break ;
    END ;
if Not Okok then BEGIN if HPiece.Execute(36,Caption,'')=mrYes then Okok:=True ; END ;
Result:=Okok ;
END ;



Function TFSaisie.PieceCorrecte : Boolean ;
Var i,Err : integer ;
    OkokT : boolean ;
BEGIN
if PasModif then BEGIN Result:=True ; Exit ; END ;
Err:=0 ; OkokT:=False ;
if ((OkBQE) or (OkJalEffet)) then
   BEGIN
   if Action=taCreat then GeneTreso:=TrouveAuto(SAJAL.COMPTEAUTOMAT,1) else
    if ((Action=taModif) and (GeneTreso='')) then GeneTreso:=TrouveAuto(SAJAL.COMPTEAUTOMAT,1) ;
   END ;
if ((EstRempli(GS,GS.RowCount-1)) or (GS.Objects[SA_Gen,GS.RowCount-1]<>Nil) or (GS.Objects[SA_Aux,GS.RowCount-1]<>Nil)) then Err:=22 ;
if Not Equilibre then BEGIN if ModeDevise then Err:=101 else Err:=103 ; END ;
if ((Err=0) and (GS.RowCount<3)) then Err:=102 ;
if Err=0 then for i:=1 to GS.RowCount-2 do
   BEGIN
   if Not LigneCorrecte(i,True,False) then BEGIN Err:=1000 ; Break ; END ;
   if GeneTreso<>'' then
     OkokT := OkokT or ( ( OkBQE or OkJalEffet ) and ( GS.Cells[SA_Gen,i] = GeneTreso ) ) ;
   END ;
if ((Err=0) and (OkBQE) and (Not OkokT)) then Err:=126 ;
If Not SuiviMP Then BEGIN if ((Err=0) and (OkJalEffet) and (Not OkokT)) then Err:=151 ; END ;
Result:=(Err=0) ;
if not Result then
   BEGIN
   if ((ModeDevise) and (Err=103)) then ForcerMode(True,0) else ErreurSaisie(Err) ;
   END ;
END ;

{==========================================================================================}
{========================= Traitements de calcul liés aux lignes ==========================}
{==========================================================================================}
procedure TFSaisie.PutScenar ( O : TOBM ; Lig : integer ) ;
Var St : String ;
    DD : TDateTime ;
    i  : integer ;
    XX : double ;
BEGIN
{Zones extra-comptables}
St:=Entete.GetMvt('E_REFINTERNE') ; if St<>'' then BEGIN O.PutMvt('E_REFINTERNE',St) ; GS.Cells[SA_RefI,Lig]:=St ; END ;
St:=Entete.GetMvt('E_LIBELLE') ; if St<>'' then BEGIN O.PutMvt('E_LIBELLE',St) ; GS.Cells[SA_Lib,Lig]:=St ; END ;
St:=Entete.GetMvt('E_REFEXTERNE') ; if St<>'' then O.PutMvt('E_REFEXTERNE',St) ;
DD:=Entete.GetMvt('E_DATEREFEXTERNE') ; if DD>Encodedate(1901,01,01)  then O.PutMvt('E_DATEREFEXTERNE',DD) ;
St:=Entete.GetMvt('E_AFFAIRE') ; if St<>'' then O.PutMvt('E_AFFAIRE',St) ;
St:=Entete.GetMvt('E_REFLIBRE') ; if St<>'' then O.PutMvt('E_REFLIBRE',St) ;
{Zones libres}
for i:=0 to 9 do
    BEGIN
    St:=Entete.GetMvt('E_LIBRETEXTE'+IntToStr(i)) ; if St<>'' then O.PutMvt('E_LIBRETEXTE'+IntToStr(i),St) ;
    if i<=3 then
       BEGIN
       XX:=Entete.GetMvt('E_LIBREMONTANT'+IntToStr(i)) ; if XX<>0 then O.PutMvt('E_LIBREMONTANT'+IntToStr(i),XX) ;
       St:=Entete.GetMvt('E_TABLE'+IntToStr(i)) ; if St<>'' then O.PutMvt('E_TABLE'+IntToStr(i),St) ;
       if i<=1 then BEGIN St:=Entete.GetMvt('E_LIBREBOOL'+IntToStr(i)) ; if St<>'' then O.PutMvt('E_LIBREBOOL'+IntToStr(i),St) ; END ;
       END ;
    END ;
DD:=Entete.GetMvt('E_LIBREDATE') ; if DD>Encodedate(1901,01,01) then O.PutMvt('E_LIBREDATE',DD) ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 03/11/2004
Modifié le ... :   /  /
Description .. : - 03/11/2004 - LG - FB 14816 - affection du e_docid qd on
Suite ........ : insert ou ajoute une ligne a une piece existante
Mots clefs ... :
*****************************************************************}
procedure TFSaisie.DefautLigne ( Lig : integer ; Init : boolean ) ;
Var O : TOBM ;
BEGIN
GS.Cells[SA_NumP,Lig]:=IntToStr(NumPieceInt) ;
if Lig>1 then
   BEGIN
   GS.Cells[SA_DateC,Lig]:=GS.Cells[SA_DateC,Lig-1] ;
   GS.Cells[SA_RefI,Lig]:=GS.Cells[SA_RefI,Lig-1] ;
   GS.Cells[SA_Lib,Lig]:=GS.Cells[SA_Lib,Lig-1] ;
   END else
   BEGIN
   GS.Cells[SA_DateC,Lig]:=DateToJour(StrToDate(E_DATECOMPTABLE.Text)) ;
   END ;
GS.Cells[SA_NumL,Lig]:=IntToStr(Lig) ;
GS.Cells[SA_Exo,Lig]:=QuelExoDT(StrToDate(E_DATECOMPTABLE.Text)) ;
AlloueMvt(GS,EcrGen,Lig,Init) ; O:=GetO(GS,Lig) ; if O=Nil then Exit ;
if ((OkScenario) and (Entete<>Nil) and (Init)) then PutScenar(O,Lig) ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 03/10/2003
Modifié le ... : 21/02/2006
Description .. : - LG - 03/10/2003 - FB 12786 - on supprime le document
Suite ........ : associé
Suite ........ : - LG - 21/02/2205 - suppression de fuite memoire
Mots clefs ... : 
*****************************************************************}
procedure TFSaisie.DetruitLigne( Lig : integer ; Totale : boolean ) ;
Var R : integer ;
OBM : TOB ;
BEGIN
GS.CacheEdit ; OBM := GetO(GS,Lig) ;
{$IFDEF SCANGED}
 if ( RechGuidId <> '' ) and ( PGIAsk('Il existe une image ou un fichier associé , voulez-vous le remplacer ?') = mrNo ) then exit ;
 SupprimeLeDocGuid ;
{$ENDIF}
DesalloueEche(Lig);
DesalloueAnal(Lig);
Desalloue(GS,SA_GEN,Lig);
Desalloue(GS,SA_Aux,Lig);
OBM.Free ;
if Totale then R:=GS.Row else R:=1 ; GS.DeleteRow(Lig) ;
CalculDebitCredit ; NumeroteLignes(GS) ; NumeroteVentils ;
GS.SetFocus ; if R>=GS.RowCount then R:=GS.RowCount-1 ;
if R>1 then GS.Row:=R else GS.Row:=1 ;
GS.Col:=SA_Gen ; GereNewLigne(GS) ; GS.Invalidate ;
GS.MontreEdit ;
END ;

procedure TFSaisie.TraiteMontant ( Lig : integer ; Calc : boolean ) ;
Var XC,XD{,SD,SC} : Double ;
    OBM         : TOBM ;
BEGIN
OBM:=GetO(GS,Lig) ; if OBM=Nil then Exit ;
XD:=ValD(GS,Lig) ; XC:=ValC(GS,Lig) ;
OBM.SetMontants(XD,XC,DEV,False) ;
if Calc then CalculDebitCredit ;
END ;

procedure TFSaisie.AfficheConf ;
BEGIN
HConf.Visible:=(ModeConf>'0') or ((PieceConf) and (Action=taConsult)) ;
END ;

procedure TFSaisie.CalculDebitCredit ;
Var i,NumBase,NbTiers : integer ;
    OBM : TOBM ;
    CGen : TGGeneral ;
    CAux : TGTiers ;
//    OBA  : TObjAna ;
    XD,XC,Solde : Double ;
BEGIN
ModeConf:='0' ; UnSeulTiers:=False ; NbTiers:=0 ;
//XD:=0 ; XC:=0 ; Solde:=0 ;
SI_TotDS:=0 ; SI_TotCS:=0 ; SI_TotDP:=0 ; SI_TotCP:=0 ;
SI_TotDD:=0 ; SI_TotCD:=0 ;
FillChar(TabTvaEnc,Sizeof(TabTvaEnc),#0) ;
for i:=1 to GS.RowCount-1 do if ((Trim(GS.Cells[SA_Gen,i])<>'') or (Guide)) then
    BEGIN
    SI_TotDS:=SI_TotDS+ValD(GS,i) ; SI_TotCS:=SI_TotCS+ValC(GS,i) ;
    OBM:=GetO(GS,i) ;
    if OBM<>Nil then
       BEGIN
       XD:=OBM.GetMvt('E_DEBIT') ; XC:=OBM.GetMvt('E_CREDIT') ;
       SI_TotDP:=SI_TotDP+OBM.GetMvt('E_DEBIT')     ; SI_TotCP:=SI_TotCP+OBM.GetMvt('E_CREDIT') ;
       SI_TotDD:=SI_TotDD+OBM.GetMvt('E_DEBITDEV')  ; SI_TotCD:=SI_TotCD+OBM.GetMvt('E_CREDITDEV') ;
      {#TVAENC}
       if OuiTvaLoc then
          BEGIN
          if isTiers(GS,i) then Inc(NbTiers) ;
           if ((SorteTva<>stvDivers) and (isHT(GS,i,True))) then
              BEGIN
              if SorteTva=stvVente then Solde:=XC-XD else Solde:=XD-XC ;
              if OBM.GetMvt('E_TVAENCAISSEMENT')='X' then
                 BEGIN
                 NumBase:=Tva2NumBase(OBM.GetMvt('E_TVA')) ;
                 if NumBase>0 then TabTvaEnc[NumBase]:=TabTvaEnc[NumBase]+Solde ;
                 END else
                 BEGIN
                 TabTvaEnc[5]:=TabTvaEnc[5]+Solde ;
                 END ;
              END ;
          END ;
       END ;
    CGen:=GetGGeneral(GS,i) ; if ((CGen<>Nil) and (CGen.Confidentiel>ModeConf)) then ModeConf:=CGen.Confidentiel ;
    CAux:=GetGTiers(GS,i) ; if CAux<>Nil then if CAux.Confidentiel>ModeConf then ModeConf:=CAux.Confidentiel ;
//    OBA:=TObjAna(GS.Objects[SA_DateC,i]) ; if OBA<>Nil then if OBA.ModeConf>ModeConf then ModeConf:=OBA.ModeConf ;	// EAGL : TObjAna plus utilisé
    END ;
SA_TotalDebit.Value:=SI_TotDS ; SA_TotalCredit.Value:=SI_TotCS ;
AfficheLeSolde(SA_Solde,SI_TotDS,SI_TotCS) ;
AfficheConf ; UnSeulTiers:=(NbTiers=1) ;
END ;

procedure TFSaisie.InfoLigne ( Lig : integer ; Var D,C,TD,TC : Double ) ;
Var Cpte : String17 ;
BEGIN
  D:=0 ; C:=0 ; TD:=0 ; TC:=0 ;
  Cpte:=GS.Cells[SA_Gen,Lig] ;
  if Cpte='' then Exit ;
  D:=ValD(GS,Lig) ;
  C:=ValC(GS,Lig) ;
  SommeSoldeCompte(SA_Gen,Cpte,TD,TC,FALSE,True) ;
END ;

procedure TFSaisie.SommeSoldeCompte ( Col : integer ; Cpte : String ; Var TD,TC : Double ; Old,OkDev : Boolean) ;
Var i : integer ;
BEGIN
TD:=0 ; TC:=0 ; if Cpte='' then Exit ;
for i:=1 to GS.RowCount-2 do if GS.Cells[Col,i]=Cpte then
    BEGIN
    if OkDev then
       BEGIN
       TD:=TD+ValD(GS,i) ; TC:=TC+ValC(GS,i) ;
       END else
       BEGIN
       TD:=TD+ValeurPivot(GS.Cells[SA_Debit,i],DEV.Taux,DEV.Quotite)-Ord(Old)*GetDouble(GS,i,'OLDDEBIT') ;
       TC:=TC+ValeurPivot(GS.Cells[SA_Credit,i],DEV.Taux,DEV.Quotite)-Ord(Old)*GetDouble(GS,i,'OLDCREDIT');
       END ;
    END ;
END ;

procedure TFSaisie.CalculSoldeCompte ( CpteG,CpteA : String ; DIG,CIG,DIA,CIA : Double ) ;
Var TDG : Double ;
    TCG : Double ;
    TDA : Double ;
    TCA : Double ;
begin
  TDG := 0 ; TCG := 0 ;
  TDA := 0 ; TCA := 0 ;
  // Affichage solde des généraux
  if ( Action <> taConsult ) and (M.Simul='N') then
    SommeSoldeCompte(SA_Gen,CpteG,TDG,TCG,TRUE,FALSE) ;
  AfficheLeSolde(SA_SoldeG,TDG+DIG,TCG+CIG) ;
  // Affichage solde des tiers
  if ( Action <> taConsult ) and (M.Simul='N') then
    SommeSoldeCompte(SA_Aux,CpteA,TDA,TCA,TRUE,FALSE) ;
  AfficheLeSolde(SA_SoldeT,TDA+DIA,TCA+CIA) ;
end ;

procedure TFSaisie.CaptureRatio ( Lig : integer ; Force : boolean ) ;
Var ND,NC,OD,OC,Rat,OT,NT : Double ;
    O   : TOBM ;
BEGIN
if PasModif then exit ;
ND:=ValD(GS,Lig) ; NC:=ValC(GS,Lig); O:=GetO(GS,Lig) ; if O=Nil then Exit ;
OD:=O.GetDebitSaisi ; OC:=O.GetCreditSaisi ;
OT:=OD+OC ; NT:=ND+NC ;
if ((OT=NT) and (OT<>0) and (Not Force)) then Exit ;
if ((OT=0) or (NT=0) or (Force)) then Rat:=1 else Rat:=Abs(NT/OT) ;
O.PutMvt('RATIO',Rat) ;
END ;

procedure TFSaisie.InsereLigne(Lig : integer) ;
BEGIN
if PasModif then Exit ;
if Lig<GS.RowCount-1 then BEGIN GS.InsertRow(Lig) ; DefautLigne(Lig,True) ; END ;
NumeroteLignes(GS) ; NumeroteVentils ; AffichePied ;
END ;

{==========================================================================================}
{================================ Appels des comptes en saisie=== =========================}
{==========================================================================================}
procedure TFSaisie.ChercheJAL ( Jal : String3 ; Zoom : boolean ) ;
BEGIN
SAJAL.Free ; SAJAL:=TSAJAL.Create(Jal,False) ;
Case Action of
   taCreat   : InitEnteteJal(True,Zoom,False) ;
   taModif   : InitModifJal(Action) ;
   taConsult : InitModifJal(Action) ;
   END ;
PieceModifiee:=True ;
OkBQE:=((SAJAL.NatureJal='BQE') or (SAJAL.NatureJal='CAI')) ;
{$IFDEF SPEC350}
OKJalEffet:=FALSE ;
{$ELSE}
OkJalEffet:=SAJAL.JalEffet ;
{$ENDIF}
OkLC:=((SAJAL.NatureJal='BQE') or (SAJAL.NatureJal='OD')) and (M.SorteLettre in [tslCheque,tslTraite,tslBOR,tslVir,tslPre]) ;
if ((OkLC) and (Action<>taConsult)) then BCheque.Enabled:=True ;
if SAJAL.OkFerme then
  begin
  HPiece.Execute(30,caption,SAJAL.JOURNAL+'  '+SAJAL.LibelleJournal+')') ;
  { FQ 18533 BVE 14.05.07 }
  if ( Action <> taConsult ) then
     E_JOURNAL.ItemIndex := 0 ;
  { END FQ 18533 }
  end ;
end;

procedure TFSaisie.AppelAuto ( indice : integer ) ;
Var Cpte : String ;
BEGIN
if PasModif then Exit ; if GS.Col<>SA_Gen then Exit ;
Cpte:=TrouveAuto(SAJAL.COMPTEAUTOMAT,indice) ; if Cpte='' then Exit ;
GS.Cells[SA_Gen,GS.Row]:=Cpte ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 05/10/2004
Modifié le ... : 25/08/2005
Description .. : - BPY le 05/10/2004 - Gestion des lookup sur compte tiers
Suite ........ : et generaux par l'intermediaire d'un THEdit et des lookup 
Suite ........ : standard AGL plutot que par l'intermediaire d'un hcompte
Suite ........ : - FB 16397 - LG  - 25/08/2005 - test de la revision 
Suite ........ : uniquement pour l'exercice en cours
Mots clefs ... : 
*****************************************************************}
function TFSaisie.ChercheGen(C,L : integer ; Force : boolean ; DoPopup : boolean) : byte;
var
    St : String;
    CGen,CGenAvant : TGGeneral;
    CAux : TGTiers;
    Idem,Changed,CollAvant,HavePopup,found : boolean;
    q : tquery;
begin
    result := 0;

    // sauvegarde de la valeutr saisie
    St := uppercase(ConvertJoker(GS.Cells[C,L]));

    // compte collectif ??
    CGenAvant := GetGGeneral(GS,L);
    if (CGenAvant <> nil) then CollAvant := CGenAvant.Collectif
    else CollAvant := false;

// BPY le 15/11/2004 => Fiche n° 14948 : intitulé du llokup ne doit pas changé ... modification du plus a la place de la tablette ...
    // set de la tablette
    compte.DataType := 'TZGENERAL';
    // set de la clause plus en fonction de ??? ...
    if (M.ANouveau) then compte.Plus := 'G_NATUREGENE<>"CHA" AND G_NATUREGENE<>"PRO"'
    else compte.Plus := '';
    // ... ou du compte auxiliaire
    CAux := GetGTiers(GS,L);
    if ((CAux <> Nil) and (GS.Cells[SA_Aux,L] <> '')) then
    begin
        if ((CAux.NatureAux = 'AUD') or (CAux.NatureAux = 'CLI')) then compte.Plus := 'G_COLLECTIF="X" AND (G_NATUREGENE="COC" OR G_NATUREGENE="COD")'
        else if ((CAux.NatureAux = 'AUC') or (CAux.NatureAux = 'FOU')) then compte.Plus := 'G_COLLECTIF="X" AND (G_NATUREGENE="COF" OR G_NATUREGENE="COD")'
        else if (CAux.NatureAux = 'SAL') then compte.Plus := 'G_COLLECTIF="X" AND G_NATUREGENE="COS"'
        else compte.Plus := 'G_COLLECTIF="X"';
    end;

  compte.Plus := compte.Plus + ' AND G_FERME<>"X" ' ;
    
// Fin BPY

    // set des coordonnées
    compte.Top := GS.Top + GS.CellRect(C,L).Top;
    compte.Left := Gs.Left + GS.CellRect(C,L).Left;
    compte.Height := GS.CellRect(C,L).Bottom - GS.CellRect(C,L).top;
    compte.Width := GS.CellRect(C,L).Right - GS.CellRect(C,L).Left;

    // recherche d'un possible compte acev le caractere de bourage !
    if GetKeyState(VK_CONTROL) < 0    // FQ 16060 : SBO 03/08/2005
      then Compte.text := st // pas de bourage en recherche sur libelle ...
      else compte.text := BourreLaDonc(st,fbGene);
    // recherche du code dans le lookup
    found := LookupValueExist(compte);
    // recherche du libellé unique
    if (not found) then
    begin
        if (not DoPopup) then
        begin
            q := OpenSQL('SELECT G_GENERAL FROM GENERAUX WHERE UPPER(G_LIBELLE) LIKE UPPER("' + st + '%") AND ' + compte.Plus,true);
            if (not(q.EOF) and (Q.RecordCount = 1)) then
            begin
                compte.text := q.Fields[0].AsString;
                found := true;
            end;
            Ferme(q);
        end;
    end;
    // recheche manuel dans le lookup
    if (not found) then
    begin
        // set du text et de la visibilité
        compte.text := st;
        compte.visible := true;
        // click ;)
        compte.ElipsisClick(compte);
        HavePopup := true;
    end
    else HavePopup := false;

    // focus !
    if ((GS.Visible) and (GS.Enabled)) then GS.SetFocus;
    // cache le ctrl
    compte.visible := false;

    // si on selectionne qqchose dans le lookup !
    if (LookupValueExist(compte)) then
    begin
        // mais si on a rien changé mais que c'est valide quand meme ....
        result := 1;

        // si il y avais un objet attaché !
        if (GS.Objects[C, L] <> nil) then
        begin
            {JP 20/02/04 : Sur le premier Focus, en ouverture de fiche, le CellsEnter n'est pas exécuté}
            if ((TGGeneral(GS.Objects[C,L]).General <> '') and (StCellCur = '')) then StCellCur := TGGeneral(GS.Objects[C,L]).General;
            if (TGGeneral(GS.Objects[C,L]).General <> St) then Changed := True
            else changed := false;
        end
        else changed := false;
        // a t'il changer ??
        Idem := ((St = compte.Text) and (GS.Objects[C, L] <> nil) and (not Changed));

        if ((not Idem) or (Force)) then
        begin
            // set du nouveau compte !
            GS.Cells[C,L] := compte.Text;
            // vire l'ancien compte dans la grille
            Desalloue(GS,C,L);
            // recup du compte et mise en objet de la grille !
            CGen := TGGeneral.Create(compte.Text);
            GS.Objects[C,L] := TObject(CGen);

            // si ct pas la meme on desaloue !
            if (not Idem) then
            begin
                if ((not CollAvant) or (not CGen.Collectif)) then DesalloueEche(L);
                DesalloueAnal(L,CGEN.Ventilable); // FQ 10296
            end;

            // compte interdit ??
            if EstInterdit(SAJAL.COMPTEINTERDIT,compte.Text, 0) > 0 then
            begin
                DesalloueEche(L);
                DesalloueAnal(L);
                Desalloue(GS,C,L);
                GS.Cells[C,L] := '';
                ErreurSaisie(4);
                Exit;
            end;

            // compte de bilan ???
            if ((not M.ANouveau) and (EstOuvreBil(compte.Text))) then
            begin
                DesalloueEche(L);
                DesalloueAnal(L);
                Desalloue(GS,C,L);
                GS.Cells[C,L] := '';
                ErreurSaisie(20);
                Exit;
            end;

            // compte de charge
            if ((M.ANouveau) and (EstChaPro(CGen.NatureGene))) then
            begin
                DesalloueEche(L);
                DesalloueAnal(L);
                Desalloue(GS,C,L);
                GS.Cells[C,L] := '';
                ErreurSaisie(21);
                Exit;
            end;

            if ((CGen.NatureGene = 'BQE') and (E_NATUREPIECE.Value <> 'ECC')) then if (PasDeviseBanque(CGen.General,E_DEVISE.Value)) then
            begin
                DesalloueEche(L);
                DesalloueAnal(L);
                Desalloue(GS,C,L);
                GS.Cells[C,L] := '';
                ErreurSaisie(25);
                Exit;
            end;

            // SBO 19/08/2004 FQ 13290 Saisie interdite sur les comptes fermés
            if CGen.Ferme then
            begin
                DesalloueEche(L);
                DesalloueAnal(L);
                Desalloue(GS,C,L);
                GS.Cells[C,L] := '';
                ErreurSaisie(30);
                Exit;
            end;

            if FInfo.LoadCompte( GS.Cells[SA_Gen,L] ) then
             begin
              if (FInfo.Compte_GetValue('G_VISAREVISION') = 'X') and ( FInfo.TypeExo = teEnCours ) then
              begin
                  DesalloueEche(L);
                  DesalloueAnal(L);
                  Desalloue(GS,C,L);
                  GS.Cells[C,L] := '';
                  ErreurSaisie(31);
                  Exit;
              end;
             end ;

            // recup l'analytique
            if (not Idem) then AlloueEcheAnal(C,L);

            Result := 1;

            // affecte les truc
            AffecteConso(C,L);
            AffecteTva(C,L);
            if ((CGen.NatureGene = 'TIC') or (CGen.NatureGene = 'TID')) then
              if AffecteRIB( C, L, False, (not Idem) { à la place de FALSE par défaut } ) then // FQ 17062 SBO 23/11/2005 Pb changement de compte de Tic / Tid
                BModifRIB.Enabled := true;
        end
        else
        begin
            CGen := TGGeneral(GS.Objects[C,L]);
            if (HavePopup) and ((not HavePopup) and (DoPopup)) then result := 1
            else result := 2;
        end;

        if (CGen <> nil) then
            if (not CGen.Collectif) then
            begin
                Desalloue(GS,SA_Aux,L);
                GS.Cells[SA_Aux,L] := '';
                if (Lettrable(CGen) = NonEche) then DesalloueEche(L);
                if (not Ventilable(CGen,0)) then DesalloueAnal(L);
            end;

        // Chargement TInfoECriture   // FQ 13246 : SBO 30/03/2005
        if GS.Cells[C,L] <> '' then
          FInfo.LoadCompte( GS.Cells[C,L] ) ;

        AffichePied;
    end;
end;

(*
Function TFSaisie.ChercheGen ( C,L : integer ; Force : boolean ) : byte ;
Var St   : String ;
    CGen,CGenAvant : TGGeneral ;
    CAux : TGTiers ;
    Idem,CollAvant,Changed : boolean ;
BEGIN
Result:=0 ; {CGen:=Nil ;} Changed:=False ; CollAvant:=False ;
if M.ANouveau then Cache.ZoomTable:=tzGBilan else Cache.ZoomTable:=tzGeneral ;
St:=uppercase(ConvertJoker(GS.Cells[C,L])) ; Cache.Text:=St ;
CGenAvant:=GetGGeneral(GS,L) ; if CGenAvant<>Nil then CollAvant:=CGenAvant.Collectif ;
CAux:=GetGTiers(GS,L) ; if ((CAux<>Nil) and (GS.Cells[SA_Aux,L]<>'')) then QuelZoomTableG(Cache,CAux) ;
if GChercheCompte(Cache,Nil) then
   BEGIN
    if GS.Objects[C,L]<>Nil then begin
      {JP 20/02/04 : Sur le premier Focus, en ouverture de fiche, le CellsEnter n'est pas exécuté}
      if (TGGeneral(GS.Objects[C,L]).General <> '') and (StCellCur = '') then
        StCellCur := TGGeneral(GS.Objects[C,L]).General;
      if TGGeneral(GS.Objects[C,L]).General<>St then Changed:=True ;
    end;
    Idem:=((St=Cache.Text) and (GS.Objects[C,L]<>Nil) and (Not Changed)) ;
    if ((Not Idem) or (Force)) then
      BEGIN
      if EstInterdit(SAJAL.COMPTEINTERDIT,Cache.Text,0)>0 then BEGIN ErreurSaisie(4) ; Exit ; END ;
      if ((Not M.ANouveau) and (EstOuvreBil(Cache.Text))) then BEGIN ErreurSaisie(20) ; Exit ; END ;
      GS.Cells[C,L]:=Cache.Text ; Desalloue(GS,C,L) ;
      CGen:=TGGeneral.Create(Cache.Text) ; GS.Objects[C,L]:=TObject(CGen) ;
      if Not Idem then
         BEGIN
         if ((Not CollAvant) or (Not CGen.Collectif)) then DesalloueEche(L) ;
         DesalloueAnal(L, CGEN.Ventilable) ; // FQ 10296
         END ;
      if ((M.ANouveau) and (EstChaPro(CGen.NatureGene))) then
         BEGIN
         DesalloueEche(L) ; DesalloueAnal(L) ; Desalloue(GS,C,L) ; GS.Cells[C,L]:='' ;
         ErreurSaisie(21) ; Exit ;
         END ;
      if ((CGen.NatureGene='BQE') and (E_NATUREPIECE.Value<>'ECC')) then if PasDeviseBanque(CGen.General,E_DEVISE.Value) then
         BEGIN
         DesalloueEche(L) ; DesalloueAnal(L) ; Desalloue(GS,C,L) ; GS.Cells[C,L]:='' ;
         ErreurSaisie(25) ; Exit ;
         END ;
      if CGen.Ferme then
         BEGIN
         // SBO 19/08/2004 FQ 13290 Saisie interdite sur les comptes fermés
         DesalloueEche(L) ;
         DesalloueAnal(L) ;
         Desalloue(GS,C,L) ;
         GS.Cells[C,L]:='' ;
         ErreurSaisie(30) ;
         Exit ;
         // FIN SBO 19/08/2004 FQ 13290 Saisie interdite sur les comptes fermés
         END ;
      if Not Idem then AlloueEcheAnal(C,L) ;
      Result:=1 ;
      AffecteConso(C,L) ; AffecteTva(C,L) ;
      If (CGen.NatureGene='TIC') Or (CGen.NatureGene='TID') Then If AffecteRIB(C,L,FALSE) Then BModifRIB.Enabled:=TRUE ;
      END else
      BEGIN
      CGen:=TGGeneral(GS.Objects[C,L]) ; Result:=2 ;
      END ;
   if ((Result>0) and (CGen<>Nil)) then if not CGen.Collectif then
      BEGIN
      Desalloue(GS,SA_Aux,L) ; GS.Cells[SA_Aux,L]:='' ;
      if Lettrable(CGen)=NonEche then DesalloueEche(L) ;
      if Not Ventilable(CGen,0) then DesalloueAnal(L) ;
      END ;
   AffichePied ;
   END ;
END ;
*)

Procedure TFSaisie.AffecteTva ( C,L : integer ) ;
Var O : TOBM ;
    CAux  : TGTiers ;
    CGen  : TGGeneral ;
    StE   : String ;
    Recharge : boolean ;
BEGIN
{#TVAENC}
if Action=taConsult then Exit ;
O:=GetO(GS,L) ; if O=Nil then Exit ;
if isTiers(GS,L) then
   BEGIN
   Recharge:=((GS.RowCount<=3) and (Not EstRempli(GS,GS.RowCount-1)) and (Not Guide) and (Not RegimeForce)) ;
   RenseigneRegime(L,Recharge) ;
   if GS.Cells[SA_Aux,L]<>'' then
      BEGIN
      CAux:=GetGTiers(GS,L) ; if CAux=Nil then Exit ;
      if ((Not ChoixExige) or (Recharge)) then
         BEGIN
         if ((SorteTva=stvAchat) or (E_NATUREPIECE.Value='OF')) then ExigeTva:=Code2Exige(CAux.Tva_Encaissement) else
          if ((SorteTva=stvDivers) and (E_NATUREPIECE.Value='OC') and (VH^.TvaEncSociete='TE')) then ExigeTva:=tvaEncais
            else ExigeTva:=tvaMixte ;
         ExigeEntete:=ExigeTva ;
         BPopTva.ItemIndex:=Ord(ExigeEntete)+1 ; BPopTvaChange(Nil) ;
         END ;
      END else
      BEGIN
      CGen:=GetGGeneral(GS,L) ; if CGen=Nil then Exit ;
      if Not ChoixExige then
         BEGIN
         if ((SorteTva=stvAchat) or (E_NATUREPIECE.Value='OF')) then ExigeTva:=Code2Exige(CGen.Tva_Encaissement) else
          if ((SorteTva=stvDivers) and (E_NATUREPIECE.Value='OC') and (VH^.TvaEncSociete='TE')) then ExigeTva:=tvaEncais
                                                                else ExigeTva:=tvaMixte ;
         ExigeEntete:=ExigeTva ; BPopTva.ItemIndex:=Ord(ExigeEntete)+1 ;
         if VH^.PaysLocalisation=CodeISOES then
            BpopTVAChange(bpopTVA) ; //XVI 24/02/2005
         END ;
      END ;
   O.PutMvt('E_REGIMETVA',GeneRegTva) ;
   END else if isHT(GS,L,True) then
   BEGIN
   CGen:=GetGGeneral(GS,L) ; if CGen=Nil then Exit ;
   StE:=FlagEncais(ExigeEntete,CGen.TvaSurEncaissement) ;
   O.PutMvt('E_TVAENCAISSEMENT',StE) ; O.PutMvt('E_REGIMETVA',GeneRegTva) ;
   O.PutMvt('E_TVA',CGen.Tva)        ;
{$IFDEF COMPTA}
   if not IsTiersSoumisTPF(GS) then
      O.PutMvt('E_TPF','')
   else
      O.PutMvt('E_TPF',CGen.Tpf) ;
{$ENDIF}
   END ;

END ;

Procedure TFSaisie.AffecteConso ( C,L : integer ) ;
Var O : TOBM ;
    CAux : TGTiers ;
    CGen : TGGeneral ;
    Code : String ;
BEGIN
if Action=taConsult then Exit ;
O:=GetO(GS,L) ; if O=Nil then Exit ;
if C=SA_AUX then
   BEGIN
   CAux:=GetGTiers(GS,L) ; if CAux=Nil then Exit ;
   Code:=CAux.CodeConso ;
   END else
   BEGIN
   CGen:=GetGGeneral(GS,L) ; if CGen=Nil then Exit ;
   Code:=CGen.CodeConso ;
   END ;
if ((C=SA_Aux) or (Code<>'')) then O.PutMvt('E_CONSO',Code) ;
END ;

Procedure TFSaisie.AffecteLeRib(O : TOBM ; Cpt : String ; ForceRAZ : Boolean = FALSE) ;
Var St,SRib : String ;
BEGIN
If O=Nil Then Exit ; If Cpt='' Then Exit ;
sRIB:=O.GetMvt('E_RIB') ;
If (SRib<>'') And ForceRAZ Then // FQ 17062 SBO 23/11/2005 Pb changement de compte de Tic / Tid
  begin
  SRib := '' ;
  O.PutMvt('E_RIB', '' ) ;
  end ;
if OkScenario then
   BEGIN
   if M.SorteLettre=tslTraite then Exit ;
   if GestionRIB<>'PRI' then Exit ;
   END else
   BEGIN
   if Not VH^.AttribRIBAuto then Exit ;
   if sRIB<>'' then Exit ;
   END ;
St:=GetRIBPrincipal(Cpt) ;
O.PutMvt('E_RIB',St) ;
END ;

Function  TFSaisie.AffecteRIB ( C,L : integer ; IsAux : Boolean ; ForceRAZ : Boolean = FALSE) : Boolean ;
Var O : TOBM ;
    Cpt{,St} : String ;
BEGIN
Result:=FALSE ;
if Action=taConsult then Exit ;
if (IsAux And (C<>SA_Aux)) Or ((Not IsAux) And (C<>SA_Gen)) then Exit ;
Cpt:=GS.Cells[C,L] ; if Cpt='' then Exit ;
O:=GetO(GS,L) ; if O=Nil then Exit ;
if ((M.TypeGuide='DEC') or (M.TypeGuide='ENC')) And Guide Then Exit ;
AffecteLeRib(O,Cpt,ForceRAZ) ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 05/10/2004
Modifié le ... : 06/10/2004
Description .. : - BPY le 05/10/2004 - Gestion des lookup sur compte tiers
Suite ........ : et generaux par l'intermediaire d'un THEdit et des lookup
Suite ........ : standard AGL plutot que par l'intermediaire d'un hcompte
Mots clefs ... :
*****************************************************************}
function TFSaisie.ChercheAux(C,L : integer ; Force : boolean ; DoPopup : boolean) : byte;
var
    st : string;
    CAux : TGTiers ;
    CGen,GenColl : TGGeneral;
    Idem,Changed,HavePopup,found : boolean;
    Err : integer;
    q : tquery;

 lStSelect     : string ;
 lStColonne    : string ;
 lStOrder      : string ;
 lStWhere      : string ;
 lStNatureGene : string ;
 lStCarLookUp  : string ;

begin
    result := 0;

    // sauvegarde de la valeutr saisie
    St := uppercase(ConvertJoker(GS.Cells[C,L]));

// BPY le 15/11/2004 => Fiche n° 14948 : intitulé du llokup ne doit pas changé ... modification du plus a la place de la tablette ...
    // set de la tablette
    compte.DataType := 'TZTTOUS';
    // FQ 19470 : SBO 10/01/2007 Uniformisation des conditions de recherche avec les autres saisie via la fonction du noyau
{
    // set de la clause plus en fonction du journal ...
    if (SAJAL.NATUREJAL = 'ACH') then compte.Plus := '(T_NATUREAUXI="AUC" OR T_NATUREAUXI="FOU" OR T_NATUREAUXI="DIV") ' //  OR T_NATUREAUXI="SAL")' // 13988 Enlévé les salariés à la demande de la qualité
    else if (SAJAL.NATUREJAL = 'VTE') then compte.Plus := '(T_NATUREAUXI="AUD" OR T_NATUREAUXI="CLI" OR T_NATUREAUXI="DIV")'    // 13988
    else compte.Plus := '(T_NATUREAUXI<>"NCP" AND T_NATUREAUXI<>"CON" AND T_NATUREAUXI<>"PRO" AND T_NATUREAUXI<>"SUS")' ;  // FQ 15838 : SBO 03/08/2005
    // ... ou en fonction du compte general !!
    CGen := GetGGeneral(GS,L);
    if ((CGen <> Nil) and (GS.Cells[SA_Gen,L] <> '')) then
    begin
        if (CGen.NatureGene = 'COC') then compte.Plus := '(T_NATUREAUXI="AUD" OR T_NATUREAUXI="CLI" OR T_NATUREAUXI="DIV")'
        else if (CGen.NatureGene = 'COF') then compte.Plus := '(T_NATUREAUXI="AUC" OR T_NATUREAUXI="FOU" OR T_NATUREAUXI="DIV")' // OR T_NATUREAUXI="SAL")' // FQ13988 VLA 02/02/2005
        else if (CGen.NatureGene = 'COS') then compte.Plus := 'T_NATUREAUXI="SAL"';
    end;
    compte.Plus := compte.Plus + ' AND T_FERME<>"X" ' ;
}// Fin BPY

    // Mise en place de la condition
   lStWhere     := '' ;
   lStColonne   := '' ;
   lStOrder     := '' ;
   lStSelect    := '' ;
   lStCarLookUp := '' ;
    if Trim(GS.Cells[ SA_Aux, L])<>''
      then lStCarLookUp := Copy(GS.Cells[SA_Aux, L],1,1) // FB 11938
      else lStCarLookUp := '' ;
    CGen := GetGGeneral(GS,L);
    if ((CGen <> Nil) and (GS.Cells[SA_Gen,L] <> ''))
      then lStNatureGene := CGen.NatureGene
      else lStNatureGene := '' ;

    CMakeSQLLookupAuxGrid( lStWhere,
                           lStColonne,
                           lStOrder,
                           lStSelect,
                           lStCarLookUp,
                           E_NaturePiece.Value,
                           lStNatureGene   ) ;
    Compte.plus := lStWhere ;
    // FIN FQ 19470 : SBO 10/01/2007 Uniformisation des conditions de recherche avec les autres saisie via la fonction du noyau

    // set des coordonnées
    compte.Top    := GS.Top + GS.CellRect(C,L).Top;
    compte.Left   := Gs.Left + GS.CellRect(C,L).Left;
    compte.Height := GS.CellRect(C,L).Bottom - GS.CellRect(C,L).top;
    compte.Width  := GS.CellRect(C,L).Right - GS.CellRect(C,L).Left;

    // recherche d'un possible compte acev le caractere de bourage !
    if GetKeyState(VK_CONTROL) < 0    // FQ 16060 : SBO 03/08/2005
      then Compte.text := st // pas de bourage en recherche sur libelle ...
      else compte.text := BourreLaDonc(st,fbAux);
    // recherche du code dans le lookup
    found := LookupValueExist(compte);
    // recherche du libellé unique
    if (not found) then
    begin
        if (not ((DoPopup) or (GetParamSoc('SO_CPCODEAUXIONLY',false)))) then
        begin
            q := OpenSQL('SELECT T_AUXILIAIRE FROM TIERS WHERE UPPER(T_LIBELLE) LIKE UPPER("' + st + '%") AND ' + compte.Plus,true);
            if (not(q.EOF) and (Q.RecordCount = 1)) then
            begin
                compte.text := q.Fields[0].AsString;
                found := true;
            end;
            Ferme(q);
        end;
    end;
    // recheche manuel dans le lookup
    if (not found) then
    begin
        // set du text et de la visibilité
        compte.text := st;
        compte.visible := true;
        // click ;)
        if GetParamSocSecur('SO_CPMULTIERS', false) then
         begin
          compte.text:= AGLLanceFiche('CP','MULTIERS','','','M;' + compte.text + ';' + compte.Plus + ';' ) ;
         end
          else
           compte.ElipsisClick(compte);
        HavePopup := true;
    end
    else HavePopup := false;

    // focus !
    if ((GS.Visible) and (GS.Enabled)) then GS.SetFocus;
    // cache le ctrl
    compte.visible := false;

    // si on selectionne qqchose dans le lookup !
    if (LookupValueExist(compte)) then
    begin
        // mais si on a rien changé mais que c'est valide quand meme ....
//        result := 1;

        // a t'il changer ??
        if (GS.Objects[C,L] <> nil) then
        begin
            if (TGTiers(GS.Objects[C,L]).Auxi <> st) then Changed := true
            else changed := false;
        end
        else changed := false;
        Idem := ((st = GS.Cells[C,L]) and (GS.Objects[C,L] <> nil) and (not Changed));

        // si oui  !
        if ((not Idem) or (Force)) then
        begin
            // set du nouveau compte
            GS.Cells[C,L] := compte.Text;
            // vire l'ancien compte dans la grille
            Desalloue(GS,C,L);
            // recup du compte et mise en objet de la grille !
            CAux := TGTiers.Create(GS.Cells[C,L]);
            GS.Objects[C,L] := TObject(CAux);

            // si c pas le meme on desaloue
            if (not Idem) then
            begin
                DesalloueEche(L);
                if ((CGen = nil) or (GS.Cells[SA_Gen,L] = '')) then DesalloueAnal(L)
                else if (CGen <> Nil) then if Not Ventilable(CGen,0) then DesalloueAnal(L);
            end;

            // Refuser tiers en devise si pas celle de saisie et pas sur devisepivot ni opposée
            if ((Not CAux.MultiDevise) and (DEV.Code <> CAux.Devise) and (DEV.Code <> V_PGI.DevisePivot) and (CAux.Devise <> '') and (E_NATUREPIECE.Value <> 'ECC')) then
            begin
                DesalloueEche(L);
                Desalloue(GS,C,L);
                GS.Cells[C,L] := '';
                ErreurSaisie(28);
                result := 0;
                Exit;
            end
            {JP 04/07/07 : FQ 18753 : message d'avertissement si le tiers est mono devise, dans une devise autre
                           que la devise pivot et que l'on saisit en devise pivot}
            else if not CAux.MultiDevise and (DEV.Code <> CAux.Devise) and (CAux.Devise <> '') then
              ErreurSaisie(28);


            // SBO 19/08/2004 FQ 13290 Saisie interdite sur les comptes fermés
            if CAux.Ferme then
            begin
                DesalloueEche(L);
                Desalloue(GS,C,L);
                GS.Cells[C,L] := '';
                ErreurSaisie(30);
                result := 0;
                Exit;
            end;

            // recup l'analytique
            if (not Idem) then AlloueEcheAnal(C,L);

            Result := 1;

            // affecte les truc
            AffecteRIB(C,L,TRUE,(Not Idem));
            AffecteConso(C,L);
            AffecteTva(C,L);
        end
        else    // si c'est le meme !
        begin
            if (HavePopup) and ((not HavePopup) and (DoPopup)) then result := 1
            else result := 2;
        end;

        // si on a trouvé un auxiliaire mais que l'on a pas de general
        if (GS.Cells[SA_Gen,L] = '') then
        begin
            // recup du collectif de l'auxiliaire
            GS.Cells[SA_Gen,L] := GetGTiers(GS,L).Collectif;
            Desalloue(GS,SA_Gen,L);

            Err := 0;
            GenColl := TGGeneral.Create(GS.Cells[SA_Gen,L]);
            GS.Objects[SA_Gen,L] := GenColl;

            // general interdit pour ce journal ?
            if (EstInterdit(SAJAL.COMPTEINTERDIT,GS.Cells[SA_Gen,L],0) > 0) then Err := 4;
            // confidentialité ??
            if (EstConfidentiel(GenColl.Confidentiel)) then Err := 24;
            // SBO 19/08/2004 FQ 13290 Saisie interdite sur les comptes fermés
            if (GenColl.Ferme) then Err := 24;

            // si on a une erreur on bloque ..
            if (Err > 0) then
            begin
                DesalloueEche(L);
                Desalloue(GS,C,L);
                Desalloue(GS,SA_Gen,L);
                GS.Cells[SA_Gen,L] := '';
                GS.Cells[SA_Aux,L] := '';
                ErreurSaisie(Err);
            end
            else AlloueEcheAnal(SA_Gen,L);
        end;

        // Chargement TInfoECriture   // FQ 13246 : SBO 30/03/2005
        if GS.Cells[C,L] <> '' then
          FInfo.LoadAux( GS.Cells[C,L] ) ;

        AffichePied;
    end;
end;

(*
Function TFSaisie.ChercheAux ( C,L : integer ; Force : boolean ) : byte ;
Var St   : String ;
    CAux : TGTiers ;
    GenColl : TGGeneral ;
    Idem,Changed : boolean ;
    Err  : integer ;
BEGIN
Result:=0 ;{ Err:=0 ;} Changed:=False ;
Cache.ZoomTable:=tzTiers ;
if SAJAL.NATUREJAL = 'ACH' then Cache.ZoomTable:=tzTToutCredit else // 13988
if SAJAL.NATUREJAL = 'VTE' then Cache.ZoomTable:=tzTToutDebit;      // 13988
St:=uppercase(ConvertJoker(GS.Cells[C,L])) ; Cache.Text:=St ;
CGen:=GetGGeneral(GS,L) ; if ((CGen<>Nil) and (GS.Cells[SA_Gen,L]<>'')) then QuelZoomTableT(Cache,CGen) ;
  if GChercheCompte(Cache,Nil) then
   BEGIN
   if GS.Objects[C,L]<>Nil then if TGTiers(GS.Objects[C,L]).Auxi<>St then Changed:=True ;
    Idem:=((St=Cache.Text) and (GS.Objects[C,L]<>Nil) and (Not Changed)) ;
    if ((Not Idem) or (Force)) then
      BEGIN
      GS.Cells[C,L]:=Cache.Text ; Desalloue(GS,C,L) ;
      if Not Idem then
        BEGIN
        DesalloueEche(L) ;
        if ((CGen=Nil) or (GS.Cells[SA_Gen,L]='')) then DesalloueAnal(L) else
           if CGen<>Nil then
           BEGIN
           if Not Ventilable(CGen,0) then DesalloueAnal(L) ;
           END ;
        END ;
      CAux:=TGTiers.Create(Cache.Text) ; GS.Objects[C,L]:=TObject(CAux) ;
      // Refuser tiers en devise si pas celle de saisie et pas sur devisepivot ni opposée
      if ((Not CAux.MultiDevise) and (DEV.Code<>CAux.Devise) and (DEV.Code<>V_PGI.DevisePivot) and
          (CAux.Devise<>'') and (E_NATUREPIECE.Value<>'ECC')) then
         BEGIN
         DesalloueEche(L) ; Desalloue(GS,C,L) ; GS.Cells[C,L]:='' ;
         ErreurSaisie(28) ; Exit ;
         END ;
      if CAux.Ferme then
         BEGIN
         // SBO 19/08/2004 FQ 13290 Saisie interdite sur les comptes fermés
         DesalloueEche(L) ;
         Desalloue(GS,C,L) ;
         GS.Cells[C,L]:='' ;
         ErreurSaisie(30) ;
         Exit ;
         // SBO 19/08/2004 FQ 13290 Saisie interdite sur les comptes fermés
         END ;
      if Not Idem then AlloueEcheAnal(C,L) ;
      Result:=1 ;
      AffecteRIB(C,L,TRUE,(Not Idem)) ;
      AffecteConso(C,L) ; AffecteTva(C,L) ;
      END else
      BEGIN
      Result:=2 ;
      END ;
   if Result>0 then if GS.Cells[SA_Gen,L]='' then
      BEGIN
      GS.Cells[SA_Gen,L]:=GetGTiers(GS,L).Collectif ;
      Desalloue(GS,SA_Gen,L) ; Err:=0 ;
      GenColl:=TGGeneral.Create(GS.Cells[SA_Gen,L]) ; GS.Objects[SA_Gen,L]:=GenColl ;
      if EstInterdit(SAJAL.COMPTEINTERDIT,GS.Cells[SA_Gen,L],0)>0 then Err:=4 ;
      if EstConfidentiel(GenColl.Confidentiel) then Err:=24 ;
      if GenColl.Ferme then
         BEGIN
         // SBO 19/08/2004 FQ 13290 Saisie interdite sur les comptes fermés
         Err:=24 ;
         // SBO 19/08/2004 FQ 13290 Saisie interdite sur les comptes fermés
         END ;
      if Err>0 then
         BEGIN
         DesalloueEche(L) ; Desalloue(GS,C,L) ; Desalloue(GS,SA_Gen,L) ;
         GS.Cells[SA_Gen,L]:='' ; GS.Cells[SA_Aux,L]:='' ;
         ErreurSaisie(Err) ;
         END else
         BEGIN
         AlloueEcheAnal(SA_Gen,L) ;
         END ;
      END ;
   AffichePied ;
   END ;
END ;
*)

procedure TFSaisie.ChercheMontant(Acol,Arow : longint) ;
BEGIN
CaptureRatio(ARow,False) ;
if ACol=SA_Debit then BEGIN if ValD(GS,ARow)<>0 then GS.Cells[SA_Credit,ARow]:='' else GS.Cells[ACol,ARow]:='' ; END
                 else BEGIN if ValC(GS,ARow)<>0 then GS.Cells[SA_Debit,ARow]:='' else GS.Cells[ACol,ARow]:='' ; END ;
FormatMontant(GS,ACol,ARow,DECDEV) ; TraiteMontant(ARow,True) ;
END ;

{==========================================================================================}
{================================= Conversions, Caluls ====================================}
{==========================================================================================}
Function  TFSaisie.DateMvt ( Lig : integer ) : TDateTime ;
BEGIN
Result:=StrToDate(E_Datecomptable.Text) ;
END ;

Function  TFSaisie.ArrS( X : Double ) : Double ;
BEGIN
Result:=Arrondi(X,DECDEV) ;
END ;

{==========================================================================================}
{========================= Affichages, Positionnements ====================================}
{==========================================================================================}
Procedure TFSaisie.AutoSuivant( Suiv : boolean ) ;
Var Cpte : String17 ;
BEGIN
if SAJAL.NbAuto<=0 then Exit ;
if ((GS.Cells[SA_Gen,GS.Row]<>'') and (Not CpteAuto)) then Exit ;
if GS.Cells[SA_Gen,GS.Row]='' then CpteAuto:=True ;
if Suiv then BEGIN if CurAuto>=SAJAL.NbAuto then CurAuto:=1 else Inc(CurAuto) ; END
        else BEGIN if CurAuto<=1 then CurAuto:=SAJAL.NbAuto else Dec(CurAuto) ; END ;
Cpte:=TrouveAuto(SAJAL.COMPTEAUTOMAT,CurAuto) ; if Cpte='' then Exit ;
GS.Cells[SA_Gen,GS.Row]:=Cpte ;
if ((OkJalEffet) and (CurAuto=1) and (Cpte<>'') and (GS.Row>1)) then GS.Cells[SA_Aux,GS.Row]:=GS.Cells[SA_Aux,GS.Row-1] ;
END ;

Procedure TFSaisie.SoldelaLigne ( Lig : integer ) ;
Var Diff,X : Double ;
    OBM    : TOBM ;
    Col    : integer ;
    Okok   : boolean ;
BEGIN
OBM:=GetO(GS,Lig) ; if OBM=Nil then Exit ;
Diff:=SI_TotDS-SI_TotCS ; Okok:=False ;
X:=-1*ValD(GS,Lig) ; if X=0 then X:=ValC(GS,Lig) ; Diff:=Diff+X ;
GS.Cells[SA_Debit,Lig]:='' ; GS.Cells[SA_Credit,Lig]:='' ;
if Diff>0 then BEGIN Col:=SA_Credit ; GS.Cells[Col,Lig]:=StrS(Diff,DECDEV) ; END
          else BEGIN Col:=SA_Debit  ; GS.Cells[Col,Lig]:=StrS(-Diff,DECDEV) ; END ;
ChercheMontant(Col,Lig) ;
// Pivot
Diff:=SI_TotDP-SI_TotCP ;
if Diff<>0 then
   BEGIN
   X:=-1*OBM.GetMvt('E_DEBIT') ; if X=0 then X:=OBM.GetMvt('E_CREDIT') ;
   Diff:=Diff+X ;
   if Diff>0 then OBM.PutMvt('E_CREDIT',Diff) else OBM.PutMvt('E_DEBIT',-Diff) ;
   Okok:=True ;
   END ;
// Devise
Diff:=SI_TotDD-SI_TotCD ;
if Diff<>0 then
   BEGIN
   X:=-1*OBM.GetMvt('E_DEBITDEV') ; if X=0 then X:=OBM.GetMvt('E_CREDITDEV') ; Diff:=Diff+X ;
   if Diff>0 then OBM.PutMvt('E_CREDITDEV',Diff) else OBM.PutMvt('E_DEBITDEV',-Diff) ;
   Okok:=True ;
   END ;
if Okok then CalculDebitCredit ;
AjusteLigne(Lig,False) ;
END ;

Procedure TFSaisie.PositionneDevise ( ReInit : boolean ) ;
Var OldIndex : integer ;
BEGIN
if ((Not E_DEVISE.Enabled) and (E_DEVISE.Value<>'') and (GS.RowCount>2)) then Exit ;
OldIndex:=-2 ;
if SAJAL.MultiDevise then
   BEGIN
   E_DEVISE_.Enabled:=True ;
   E_DEVISE.Enabled:=True ;
   END else
   BEGIN
   E_DEVISE_.Value:=V_PGI.DevisePivot ;
   E_DEVISE_.Enabled:=False ;
   E_DEVISE.Value:=V_PGI.DevisePivot ;              // FQ 15790 SBO 11/08/2005
   E_DEVISE.Enabled:=False ;                        // FQ 15790 SBO 11/08/2005
   END ;
if ((ReInit) and (OldIndex<>-2)) then BEGIN E_DEVISE.ItemIndex:=OldIndex ; E_DEVISEChange(Nil) ; END ;
END ;

procedure TFSaisie.AffichePied ;
Var CpteG : String ;
    CpteA : String ;
    DIG   : double ;
    CIG   : double ;
    DIA   : double ;
    CIA   : double ;
begin
  // Init. zones et variables
  G_LIBELLE.Font.Color:=clBlack ;
  G_LIBELLE.Caption:='' ;
  T_LIBELLE.Font.Color:=clBlack ;
  T_LIBELLE.Caption:='' ; // Modif SBO 18/08/2004
  CpteG:='' ;
  CpteA:='' ;
  DIG:=0 ;
  CIG:=0 ;
  DIA:=0 ;
  CIA:=0 ;

  // Infos Généraux
  if FInfo.LoadCompte( GS.Cells[SA_Gen,GS.Row] ) then
    begin
    G_LIBELLE.Caption := FInfo.Compte_GetValue('G_LIBELLE') ;
    CpteG             := FInfo.Compte_GetValue('G_GENERAL') ;
    FInfo.Compte.Solde( DIG, CIG, FInfo.TypeExo ) ;   // Calcul du solde en fonction de l'année de référence   // FQ 13246 : SBO 30/03/2005
    LSA_SoldeG.Visible:=True ;
    end
  else LSA_SoldeG.Visible:=False ;

  // Infos Tiers
  if FInfo.LoadAux( GS.Cells[SA_Aux,GS.Row] ) then
    begin
    T_LIBELLE.Caption := FInfo.Aux_GetValue('T_LIBELLE') ;
    CpteA := FInfo.Aux_GetValue('T_AUXILIAIRE') ;
    FInfo.Aux.Solde( DIA, CIA, FInfo.TypeExo ) ; // Calcul du solde en fonction de l'année de référence   // FQ 13246 : SBO 30/03/2005
    LSA_SoldeT.Visible:=True ;
    if (FInfo.Aux_GetValue('YTC_ACCELERATEUR')='X') then
      T_LIBELLE.Caption := T_LIBELLE.Caption + ' ' + TraduireMemoire('Accélérateur') + ' ' + VarToStr(FInfo.Aux_GetValue('YTC_SCHEMAGEN')) + ' (AUTOMATIQUE) ' ;
    end
  else LSA_SoldeT.Visible:=False ;

  // Affichage des soldes cumulés
  CalculSoldeCompte( CpteG, CpteA, DIG, CIG, DIA, CIA ) ;

end ;

procedure TFSaisie.GereZoom ;
Var Okok : boolean ;
BEGIN
Okok:=False ;
if ((GS.Col=SA_Gen) and (GS.Cells[SA_Gen,GS.Row]<>'')) then Okok:=True ;
if ((GS.Col=SA_Aux) and (GS.Cells[SA_Aux,GS.Row]<>'')) then Okok:=True ;
BZoom.Enabled:=Okok ;
END ;

{==========================================================================================}
{================================ Methodes de la form =====================================}
{==========================================================================================}

Procedure TFSaisie.FocusSurNatP ;
BEGIN
If E_JOURNAL.Focused And VH^.JalLookUp Then
  BEGIN
  If E_NATUREPIECE.CanFocus Then E_NATUREPIECE.SETFOCUS Else If E_DateComptable.CanFocus Then E_DateComptable.SetFocus ;
  END ;
(*NextControl(Self,TRUE) ;*)
END ;

procedure TFSaisie.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var
    OkG,Vide : boolean ;
    keys : TKeyboardState;
    MajOk : Boolean; {JP 05/07/07 : FQ 19078}
begin
if Not GS.SynEnabled then BEGIN Key:=0 ; Exit ; END ;
OkSuite:=False ;
{Panel de création du Guide}
if ModeCG then
   BEGIN
   Case Key of
      VK_ESCAPE : BCAbandonClick(Nil) ;
      VK_F10    : BCValideClick(Nil) ;
//      VK_F1     : BCAideClick(Nil) ;
      END ;
   Exit ;
   END ;
{Panel de Swap devise}
if ((Action<>taConsult) and (ModeForce<>tmNormal)) then
   BEGIN
   Case ModeForce of
     tmDevise : BEGIN
                if Shift<>[] then Exit ;
                if (Key in [VK_RETURN,VK_ESCAPE,VK_PRIOR,VK_NEXT,VK_END,VK_HOME,VK_UP,VK_DOWN])=False then Exit ;
                if Key in [VK_RETURN,VK_ESCAPE] then ForcerMode(False,Key) ;
                Exit ;
                END ;
      tmPivot : if ((Shift<>[ssAlt]) or (Key<>VK_F8)) then
                   BEGIN
                   if (Key in [VK_RETURN,VK_ESCAPE,VK_PRIOR,VK_NEXT,VK_END,VK_HOME,VK_UP,VK_DOWN])=False then Exit ;
                   END ;
               End ;
   END ;
{Saisie normale}
OkG:=(Screen.ActiveControl=GS) ; Vide:=(Shift=[]) ;
{JP 05/07/07 : FQ 19078 : Pour gérer le vérouillage Majuscule pour distinguer le + du = (Key = 187)
               à True si le clavier est vérouillé ; Remarque le '+' est géré dans le KeyPress}
MajOk := (GetKeyState(VK_CAPITAL) > 0);

Case Key of
  VK_F3     : BEGIN
              if ((OkG) and (Vide) and (GS.Col=SA_Gen)) then BEGIN FocusSurNatP ; Key:=0 ; AutoSuivant(False) ; END else
               if Shift=[ssShift] then BEGIN FocusSurNatP ; Key:=0 ; BPrevClick(Nil) ; END ;
              END ;
  VK_F4     : BEGIN
              if ((OkG) and (Vide) and (GS.Col=SA_Gen)) then BEGIN Key:=0 ; AutoSuivant(True) ; END else
               if Shift=[ssShift] then BEGIN Key:=0 ; BNextClick(Nil) ; END ;
              END ;
  VK_F5     :   // BPY le 05/10/2004 et le 29/10/2004 : gestion du lookup std a la place du thcompte
                if ((OkG) and ((Vide) or (Shift=[ssCtrl]))) then
                begin
                    Key := 0;
                    if ((GetParamSoc('SO_CPCODEAUXIONLY',false) = true) and (GS.Col = SA_Aux) and (Shift = [ssCtrl])) then
                    begin
                        GetKeyboardState(Keys);
                        keys[VK_CONTROL] := 0;
                        keys[VK_LCONTROL] := 0;
                        keys[VK_RCONTROL] := 0;
                        SetKeyboardState(Keys);
                    end;
                    ClickZoom(true);
                end
                else if ((OkG) and (Shift=[ssAlt])) then
                begin
                    Key := 0;
                    ClickConsult;
                end;
                // Fin BPY
  VK_F6     : if ((OkG) and (Vide)) then BEGIN Key:=0 ; ClickSolde(False,True) ; END ;
  VK_F8     : if Shift=[ssAlt] then BEGIN Key:=0 ; ClickSwapPivot ; END
              else if Shift=[ssCtrl] then
               begin
                Key := 0 ;
                {$IFDEF SCANGED}
                //if ( RechGuidId <> '' ) and ( PGIAsk('Il existe un fichier associé , voulez-vous le supprimer ?') = mrNo ) then exit ; 
                AjouterFichierDansGed(Self) ;
                {$ENDIF}
               end ;
  VK_F9     : if Vide then BEGIN Key:=0 ; ClickGenereTva ; END ;
  VK_F10    : if Vide then BEGIN Key:=0 ; FocusSurNatP ; ClickValide ; END ;
  VK_ESCAPE : if Vide then begin key := 0 ; ClickAbandon ; end ;
  VK_RETURN : if ((OkG) and (Vide)) then KEY:=VK_TAB ;
  VK_INSERT : if ((OkG) and (Vide)) then BEGIN Key:=0 ; ClickInsert ; END ;
  VK_DELETE : if ((OkG) and (Shift=[ssCtrl])) then BEGIN Key:=0 ; ClickDel ; END ;
  VK_BACK   : if ((OkG) and (Shift=[ssCtrl])) then BEGIN Key:=0 ; VideZone(GS) ; END ;
{^Home}  36 : if Shift=[ssCtrl] then BEGIN Key:=0 ; GotoEntete ; END ;
{A1.9}49..57: if Shift=[ssAlt] then BEGIN AppelAuto(Key-48) ; Key:=0 ; END ;
//{AA}     65 : if ((OkG) and (Shift=[ssAlt])) then BEGIN Key:=0 ; ClickVentil ; END ;
{^A}     65 : if ((OkG) and (Shift=[ssCtrl])) then BEGIN Key:=0 ; ClickVentil ; END ;
{AC}     67 : if Shift=[ssAlt] then BEGIN Key:=0 ; ClickComplement ; END ;
{AE}     69 : if ((OkG) and (Shift=[ssAlt])) then BEGIN Key:=0 ; ClickEche ; END ;
{^F}     70 : if Shift=[ssCtrl] then BEGIN Key:=0 ; ClickCherche ; END ;
{^G}     71 : if Shift=[ssctrl] then BEGIN Key:=0 ; ClickCreerGuide ; END else
{AG}           if Shift=[ssAlt] then BEGIN Key:=0 ; ClickGuide ; END ;
{AH}     72 : if Shift=[ssAlt]  then BEGIN Key:=0 ; ClickChancel ; END ;
{AI}     73 : if Shift=[ssAlt]  then BEGIN Key:=0 ; ClickDevise ; END ;
{AJ}     74 : if Shift=[ssAlt]  then BEGIN FocusSurNatP ; Key:=0 ; ClickJournal ; END ;
{^K}     75 : if Shift=[ssCtrl] then BEGIN Key:=0 ; ClickProrata ; END ;
{AL}     76 : if Shift=[ssAlt]  then BEGIN Key:=0 ; ClickLibelleAuto ; END ;
{AM}     77 : if Shift=[ssAlt]  then BEGIN Key:=0 ; ClickPieceGC ; END ;
{^O}     79 : if Shift=[ssctrl] then BEGIN Key:=0 ; ClickImmo ; END ;
{AP}     80 : if Shift=[ssAlt]  then BEGIN Key:=0 ; ClickDernPieces ; END ;
{^R}     82 : if Shift=[ssCtrl]  then BEGIN Key:=0 ; ClickChoixRegime ; END else
{AR}             if Shift=[ssAlt]  then BEGIN Key:=0 ; ClickModifRIB ; END ;
{AS}     83 : if Shift=[ssAlt]  then BEGIN FocusSurNatP ; Key:=0 ; ClickScenario ; END ;
{^T}     84 : if Shift=[ssctrl] then BEGIN Key:=0 ; ClickControleTva(False,0) ; END else
{AT}               if Shift=[ssAlt]  then BEGIN Key:=0 ; ClickEtabl ; END ;
{AV}     86 : if Shift=[ssAlt]  then BEGIN Key:=0 ; ClickModifTva ; END ;
{^Y}     89 : if Shift=[ssCtrl] then BEGIN Key:=0 ; ClickVidePiece(True) ; END ;
             {JP 05/07/07 : FQ 19078 : ajout du not MajOk}
{=}     187 : if ((OkG) and (Vide) and not MajOk) then BEGIN Key:=0 ; ClickSolde(True,True) ; END ;
  END ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Piot
Créé le ...... : 22/02/2005
Modifié le ... : 23/02/2005
Description .. : - LG  - 22/02/2005 - Creationde la TOB des ecritures
Suite ........ : complementaire
Suite ........ : - LG - 23/02/2005 - Ajout de la creation du TInfoEcr pour
Suite ........ : l'accelerateur de saisie
Mots clefs ... : 
*****************************************************************}
procedure TFSaisie.CreateSaisie ;
BEGIN
GS.VidePile(True) ; GS.RowCount:=2 ; GS.TypeSais:=tsGene ; GS.ListeParam:='SSAISIE1' ;
SAJAL:=Nil ; DecDev:=V_PGI.OkdecV ; PieceModifiee:=False ; Revision:=False ; UnChange:=False ;
TS:=TList.Create ; TGUI:=TList.Create ; TPIECE:=TList.Create ; TECHEORIG:=TList.Create ;
TDELECR:=TList.Create ; TDELANA:=TList.Create ; TSA := TList.Create ;
Scenario:=TOBM.Create(EcrSais,'',True) ; Entete:=TOBM.Create(EcrSais,'',True) ;
MemoComp:=HTStringList.Create ; GeneRegTva:='' ; StCellCur:='' ;
ModifSerie:=TOBM.Create(EcrGen,'',True) ; ModifNext:=False ; GeneCharge:=False ;
DejaRentre:=False ; OkScenario:=False ; ModeConf:='0' ; PieceConf:=False ;
Volatile:=False ; CEstGood:=False ; ModeLC:=False ;
ValideEcriture:=False ; GestionRIB:='...' ; NeedEdition:=False ; EtatSaisie:='' ;
YATP:=Saiscomm.yaRien ; ListeTP:=HTStringList.Create ; RentreDate:=False ; RegimeForce:=False ;
{#TVAENC}
SorteTva:=stvDivers ; OuiTvaLoc:=False ;
ExigeTva:=tvaMixte ; ExigeEntete:=tvaMixte ; ChoixExige:=False ;
BPopTva.ItemIndex:=Ord(ExigeEntete)+1 ;
if VH^.PaysLocalisation=CodeISOES then
   BpopTVAChange(bpopTVA) ; //XVI 24/02/2005
ToutDebit:=True ; ToutEncais:=True ;
{Param Liste}
GS.ColLengths[SA_Gen]:=VH^.Cpta[fbGene].Lg ;
if VH^.CPCodeAuxiOnly then GS.ColLengths[SA_Aux]:=VH^.Cpta[fbAux].Lg else GS.ColLengths[SA_Aux]:=35 ;
GS.ColLengths[SA_RefI]:=35 ; GS.ColLengths[SA_Lib]:=35 ;
FEcrCompl := TOB.Create('',nil,-1) ;
FInfo := TInfoEcriture.Create ;
END ;

procedure TFSaisie.PosLesSoldes ;
BEGIN
LSA_SoldeG.SetBounds(SA_SoldeG.Left,SA_SoldeG.Top,SA_SoldeG.Width,SA_SoldeG.Height) ;
LSA_SoldeT.SetBounds(SA_SoldeT.Left,SA_SoldeT.Top,SA_SoldeT.Width,SA_SoldeT.Height) ;
LSA_Solde.SetBounds(SA_Solde.Left,SA_Solde.Top,SA_Solde.Width,SA_Solde.Height) ;
LSA_TotalDebit.SetBounds(SA_TotalDebit.Left,SA_TotalDebit.Top,SA_TotalDebit.Width,SA_TotalDebit.Height) ;
LSA_TotalCredit.SetBounds(SA_TotalCredit.Left,SA_TotalCredit.Top,SA_TotalCredit.Width,SA_TotalCredit.Height) ;
SA_SoldeG.Visible:=False ; SA_SoldeT.Visible:=False ; SA_Solde.Visible:=False ;
SA_TotalDebit.Visible:=False ; SA_TotalCredit.Visible:=False ;
END ;

procedure TFSaisie.FormCreate(Sender: TObject);
//Var PP : thPanel ;
begin
GS.PostDrawCell:=PostDrawCell ;
if (VH^.PaysLocalisation=CodeISOES) then
 begin
  G_LIBELLE.AutoSize := False ;
  G_LIBELLE.Width := 224 ;
  T_LIBELLE.Autosize := false ;
  T_LIBELLE.Width := 224 ;
 end ; //XVI 24/02/2005
CreateSaisie ;
WMinX:=Width ; WMinY:=Height_Ecr ;
RegLoadToolbarPos(Self,'Saisie') ;
PosLesSoldes ;
{$IFDEF CCMP}
// CA- 16/10/2006 - Remise en place des index de colonne en saisie normale
initSA_SaisieNormale ;
SA_PILESAI := 'SAI;' + SA_PILESAI;
{$ENDIF}
{$IFDEF CCS3}     //SG6 01/12/2004 Bouton TVA pas linké en 6.01 en S3
if VH^.PaysLocalisation<>CodeISOEs then //XVI 24/02/2005
  BModifTVA.Visible:=false;
{$ENDIF}

  FClosing := False ;

  // SBO 01/07/2007 : enregistrement spécifique pour pb CWAS
  FBoInsertSpecif := EstSpecif('51215') ;

end;

procedure TFSaisie.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  CanClose:=FermerSaisie ;
  FClosing := True ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Piot
Créé le ...... : 22/02/2005
Modifié le ... :   /  /    
Description .. : - LG  - 22/02/2005 - Suppression de la TOB des ecritures 
Suite ........ : complementaire
Mots clefs ... : 
*****************************************************************}
procedure TFSaisie.CloseSaisie ;
var
 i : integer ;
BEGIN
  GS.VidePile(True) ;

  FreeAndNil(SAJAL) ;
  FreeAndNil(FInfo) ;

  VideListe(TS) ;        FreeAndNil( TS ) ;
  VideListe(TECHEORIG) ; FreeAndNil( TECHEORIG ) ;
  VideListe(TGUI) ;      FreeAndNil( TGUI ) ;
  VideListe(TPIECE) ;    FreeAndNil( TPIECE ) ;
  VideListe(TDELECR) ;   FreeAndNil( TDELECR ) ;
  VideListe(TDELANA) ;   FreeAndNil( TDELANA ) ;

  FreeAndNil( MemoComp ) ;
  FreeAndNil( Scenario ) ;
  FreeAndNil( Entete ) ;
  FreeAndNil( TOBNumChq ) ;

  if ListeTP <> nil then ListeTP.Clear ;
  FreeAndNil( ListeTP ) ;
  FreeAndNil( ModifSerie ) ;

  FreeAndNil(FEcrCompl) ;

  if TSA <> nil then
   begin
    for i:=0 to TSA.Count-1 do
     if assigned(TSA[i]) then Dispose(TSA[i]);
    TSA.Free ;
   end ;

  PurgePopup(POPS) ; PurgePopup(POPZ) ;

END ;

procedure TFSaisie.FormClose(Sender: TObject; var Action: TCloseAction);
begin
{$IFDEF SCANGED}
  // Ajout CA - 20/02/2007 - Dans le cas où la GED n'est pas sérialisée
  if not VH^.OkModGed then FinalizeGedFiles;
{$ENDIF}
CloseSaisie ;
RegSaveToolbarPos(Self,'Saisie') ;
if Parent is THPanel then
   BEGIN
   Case Self.Action of
     taCreat : _Bloqueur('nrSaisieCreat',False) ;
     taModif : _Bloqueur('nrSaisieModif',False) ;
     END ;
   Action:=caFree ;
   END ;
If ModLess Then Action:=caFree ;

{$IFDEF CCMP}
// CA- 16/10/2006 - Remise en place des index de colonne en saisie normale
if QuelleSaisieCascade = 'EFF' then
  RestoreSAEff ;
{$ENDIF}

  // précédemment dans le closeFen ( Modification par SG6 01/12/2004 FQ 15046 )
  // SBO 05/06/2005 : Ne tenait pas compte de l'annulation de fermeture
(*** annulé par XP le 3-11-2005
  if IsInside(Self) then
  begin
    CloseInsidePanel(Self) ;
    THPanel(Self.parent).FInsideForm := nil;
    THPanel(Self.parent).VideToolBar;
  end;
*)
end;

procedure TFSaisie.GereANouveau ;
BEGIN
if Action<>taCreat then Exit ; if Not M.ANouveau then Exit ;
E_DATECOMPTABLE.EditText:=DateToStr(VH^.Encours.Deb) ;
E_DATECOMPTABLE.Enabled:=False ;
E_DEVISE.Value:=V_PGI.DevisePivot ; E_DEVISE.Enabled:=False ;
E_NATUREPIECE.Value:='OD' ; E_NATUREPIECE.Enabled:=False ; E_NATUREPIECE.Style:=csSimple ;
BGenereTVA.Enabled:=False  ; BControleTVA.Enabled:=False ;
BProrata.Enabled:=False    ; BGuide.Enabled:=False ; BMenuGuide.Enabled:=False ;
BCreerGuide.Enabled:=False ; BLibAuto.Enabled:=False ;
END ;

procedure TFSaisie.LectureSeule ;
BEGIN
BInsert.Enabled:=False ; BSDel.Enabled:=False ; BGuide.Enabled:=False ;
BGenereTVA.Enabled:=False ; BProrata.Enabled:=False ; BLibAuto.Enabled:=False ;
BSolde.Enabled:=False ; BDernPieces.Enabled:=False ; BCreerGuide.Enabled:=False ;
BMenuGuide.Enabled:=False ; BModifSerie.Enabled:=False ;
END ;

procedure TFSaisie.RenseignePopTva ;
BEGIN
if Action=taCreat then Exit ;
if Not BPopTva.Visible then Exit ;
if ToutDebit then ExigeTva:=tvaDebit else
 if ToutEncais then ExigeTva:=tvaEncais else
    ExigeTva:=tvaMixte ;
ExigeEntete:=ExigeTva ;
BPopTva.ItemIndex:=Ord(ExigeEntete)+1 ;
if VH^.PaysLocalisation=CodeISOES then
   BpopTVAChange(bpopTVA) ; //XVI 24/02/2005
END ;

procedure TFSaisie.TraiteOrigMP ;
Var TOBL : TOB ;
    i    : integer ;
    St   : String ;
    T    : HTStrings ;
    YY   : RMVT ;
BEGIN
if Not SuiviMP then Exit ;
T:=HTStringList.Create ;
for i:=0 to TOBMPOrig.Detail.Count-1 do
    BEGIN
    TOBL:=TOBMPOrig.Detail[i] ;
    YY:=TOBToIdent(TOBL,True) ; St:=EncodeLC(YY) ;
    T.Add(St) ;
    END ;
TECHEORIG.Add(T) ;
if ((Action=taModif) and ((M.SorteLettre in [tslCheque,tslTraite,tslBOR,tslVir,tslPre]))) then
   BEGIN
   BlocageEntete(Self,False) ;
   LeTimer.Enabled:=True ; ModeLC:=True ; GridEna(GS,False) ;
   END ;
if ((Action=taModif) and (M.smp in [smpEncTous,smpDecTous]) and (SuiviMP) and (GenereAuto)) then
   BEGIN
   BlocageEntete(Self,False) ;
   LeTimer.Enabled:=True ; ModeLC:=True ; GridEna(GS,False) ;
   END ;
{b FP 21/02/2006}
if (Action=taModif) and (M.smp in [smpCompenCli, smpCompenFou]) and (SuiviMP) then
   BEGIN
   BlocageEntete(Self,False) ;
   ModeLC:=True ; GridEna(GS,True) ;
   GereOptionsGrid(1) ;
   LeTimer.Enabled:= GenereAuto;      {FP 25/04/2006 FQ17966}
   END ;
{e FP 21/02/2006}
END ;

procedure TFSaisie.TraiteFormuleMP(St : String ; O : TOBM ; Lig,Col : Integer ; FormuleEnDur : String = '') ;
Var St1,St2 : String ;
BEGIN
If FormuleEnDur='' Then St2:=O.GetMvt(St) Else St2:=FormuleEnDur ;
If St2='' Then Exit ;
If YAFormule(St2) Then
  BEGIN
  CurLig:=Lig ;
  St1:=GFormule(St2,Load_Sais,Nil,1) ;
  St1 := Copy( St1 ,1 ,35 ) ; // FQ 14962 Modif SBO 16/11/2004 pour éviter allocation de chaîne trop longue par rapport type de la colonne ( VARCHAR2(35) )
  O.PutMvt(St,St1) ;
  If Col>0 Then GS.Cells[Col,Lig]:=St1 ;
  END ;
END ;

procedure TFSaisie.AffecteRef(FormuleOrigine : Boolean) ;
Var O : TOBM ;
    i : Integer ;
    CptG : String ;
  Label 0 ;
BEGIN
If (m.SorteLettre<>tslAucun) And FormuleOrigine Then
  BEGIN
  Case M.smp of
    smpEncTraEdt,smpEncTraEdtNC     : CptG:=M.Section ;
    smpEncPreEdt,smpEncPreEdtNC     : CptG:=M.General ;
    smpDecChqEdt,smpDecChqEdtNC     : CptG:=M.General ;
    smpDecVirEdt,smpDecVirEdtNC     : CptG:=M.General ;
    smpDecVirInEdt,smpDecVirInEdtNC : CptG:=M.General ;
    smpDecBorEdt,smpDecBorEdtNC     : CptG:=M.Section ;
    Else Goto 0 ;
    END ;
  for i:=1 to GS.RowCount-2 do
      BEGIN
      O:=GetO(GS,i) ; if O=Nil then Continue ;
      If O.GetMvt('E_GENERAL')=CptG Then
        BEGIN
        // Si on avait un champs avec [..] par exemple, il ne passait pas dans traite formule
        TraiteFormuleMP('E_REFINTERNE',O,i,SA_RefI,M.FormuleRefCCMP.Ref2) ;
        TraiteFormuleMP('E_LIBELLE',O,i,SA_Lib,M.FormuleRefCCMP.Lib2) ;
        TraiteFormuleMP('E_REFEXTERNE',O,i,0,M.FormuleRefCCMP.RefExt2) ;
        TraiteFormuleMP('E_REFLIBRE',O,i,0,M.FormuleRefCCMP.RefLib2) ;
        END Else
        BEGIN
        TraiteFormuleMP('E_REFINTERNE',O,i,SA_RefI,M.FormuleRefCCMP.Ref1) ;
        TraiteFormuleMP('E_LIBELLE',O,i,SA_Lib,M.FormuleRefCCMP.Lib1) ;
        TraiteFormuleMP('E_REFEXTERNE',O,i,0,M.FormuleRefCCMP.RefExt1) ;
        TraiteFormuleMP('E_REFLIBRE',O,i,0,M.FormuleRefCCMP.RefLib1) ;
        END ;
      END ;
  END Else
0:for i:=1 to GS.RowCount-2 do
    BEGIN
    O:=GetO(GS,i) ; if O=Nil then Continue ;
    TraiteFormuleMP('E_REFINTERNE',O,i,SA_RefI) ;
    TraiteFormuleMP('E_LIBELLE',O,i,SA_Lib) ;
    TraiteFormuleMP('E_REFEXTERNE',O,i,0) ;
    TraiteFormuleMP('E_REFLIBRE',O,i,0) ;
    END ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. : - 24/04/2003 - reaffectation des numero de colonne. Ils
Suite ........ : sont commun avec la saisie en vrac.
Mots clefs ... :
*****************************************************************}
procedure TFSaisie.ShowSaisie ;
Var OkM,bc,OkClo,OkPasJal,OkPasModif,OkBloquePieceImport : boolean ;
    tmp : integer;
BEGIN
InitDefautColSaisie ; // LG 24/04/2003
GeneCharge:=True ; GuideAEnregistrer:=FALSE ; CodeGuideFromSaisie:='' ; appeldatepourguide:=FALSE ;
if ((M.Simul='N') or (M.Simul='S') or (M.Simul='R')) then OuiTvaLoc:=VH^.OuiTvaEnc ;
LastPaie:='' ; LastEche:=0 ; LastVal:=0 ; LastTraiteCHQ:='' ;
TresoLettre:=False ; PieceConf:=False ;

{$IFDEF EAGLCLIENT}
  BPrev.Visible := False ;
  BNext.Visible := False ;
  Sep97.Visible := False ;
{$ELSE}
  BPrev.Visible := (QListe<>NIL) ;
  BNext.Visible := (QListe<>NIL) ;
  Sep97.Visible := (QListe<>Nil) ;
{$ENDIF}

	if Action<>taCreat
  	then E_JOURNAL.DataType:='ttJournal'
    else
   		BEGIN
      {$IFDEF CCMP}
        E_JOURNAL.DataType := 'ttJalBanque';
        if (VH^.CCMP.LotCli)
          then E_NATUREPIECE.Plus := 'AND CO_CODE<>"OD" AND CO_CODE<>"OF" AND CO_CODE<>"RF"'
          else E_NATUREPIECE.Plus := 'AND CO_CODE<>"OD" AND CO_CODE<>"OC" AND CO_CODE<>"RC"';
      {$ELSE}

   		if ((Action=taCreat) and (M.FromGuide) and (M.LeGuide<>'') and (M.Simul='S'))
      	then E_JOURNAL.DataType:='TTJOURNAUX'
        else if ((Action=taCreat) and (M.Simul='N') and (M.ANouveau)) then
      		BEGIN
		      E_JOURNAL.DataType:='ttJalANouveau' ;
    		  BScenario.Enabled:=False ; BMenuGuide.Enabled:=False ;
      		END
          else
      			BEGIN
            if M.Simul='I'
              then E_JOURNAL.DataType:='CPJOURNALIFRS'
              else if M.Simul<>'N'
            	  then E_JOURNAL.DataType:='ttJalSansEcart'
                else BEGIN
                     E_JOURNAL.DataType:='ttJalSaisie' ;
                     END ;
      			END ;
      {$ENDIF}
      If Action=taCreat Then
        BEGIN
        // GP le 20/08/2008 21496
        RetoucheHVCPoursaisie(E_JOURNAL) ;
        RetoucheHVCPoursaisie(E_ETABLISSEMENT) ;
        END ;
   		END ;

ModeCG:=False ; Revision:=False ; UnSeulTiers:=False ; UnChange:=False ;
OkM:=((M.Valide) and (Action=taModif)) ;
OkPasModif:=((Action=taModif) and (Not ExJaiLeDroitConcept(TConcept(ccSaisEcritures),False))) ;
// GG COM
//OkBloquePieceImport:=FALSE ;
OkBloquePieceImport:=M.BloquePieceImport ;
If OkBloquePieceImport Then Action:=taConsult ;
OkPasJal:=((VH^.JalAutorises<>'') and (Pos(';'+M.Jal+';',VH^.JalAutorises)<=0) and (Action=taModif)) ;
OkClo:=((Action=taModif) and (ControleDate(DateToStr(M.DateC)) in [2,3])) ;
if ((OkM) or (OkClo) or (OkPasJal) or (OkPasModif)) then Action:=taConsult ;
if ((Action=taCreat) and (M.Etabl='')) then M.Etabl:=VH^.EtablisDefaut ;
DefautEntete ; DefautPied ; NbLigEcr:=0 ; NbLigAna:=0 ; GeneTreso:='' ;
ChargeLignes ;
ChargeAnals ;
ChargeCompl ;
CalculDebitCredit ;
OkPremShow:=True ; OkSuite:=False ;
if TresoLettre then Action:=taConsult ;
Case Action of
   taConsult : BEGIN
               LectureSeule ;
               if SAJAL.NatureJal='ANO' then BEGIN M.ANouveau:=True ; GereANouveau ; END ;
               END ;
   taModif   : BEGIN
               ChargeSoldesComptes ; Revision:=True ; If M.ForceModif Then Revision:=FALSE ;
               if SuiviMP then Revision:=False ;
               BDernPieces.Enabled:=False ;
               if SAJAL.NatureJal='ANO' then BEGIN M.ANouveau:=True ; GereANouveau ; END ;
               END ;
   taCreat   : BEGIN
               BDernPieces.Enabled:=True ;
               if M.ANouveau then GereANouveau ;
               END ;
   END ;
if Action<>taModif then AffecteGrid(GS,Action) ; PieceModifiee:=False ;
AfficheLeSolde(SA_Solde,0,0) ; AttribLeTitre ; QuelNExo ;
RenseignePopTva ;
if OkM then HPiece.Execute(11,caption,'') else
 if OkClo then HPiece.Execute(38,caption,'') else
  if OkPasJal then HPiece.Execute(40,caption,'') else
   if OkPasModif then HPiece.Execute(48,caption,'') else
    if TresoLettre then HPiece.Execute(32,caption,'') Else
     if OkBloquePieceImport then HPiece.Execute(52,caption,'') ;
if Action<>taCreat then
   BEGIN
   if ((M.NumLigVisu<>0) and (M.NumLigVisu<=GS.RowCount-2)) then
      BEGIN
      GS.Row:=M.NumLigVisu ;
      if Action=taConsult then GS.Cells[SA_NumL,GS.Row]:='**'+GS.Cells[SA_NumL,GS.Row]+'**' ;
      END else if M.General<>'' then
      begin
        tmp := TrouveLigCompte(GS,SA_Gen,0,M.General) ;
        if (tmp <> -1) then GS.Row := tmp;
      end;
   END ;
GSRowEnter(Nil,GS.Row,bc,False) ; OkPremShow:=False ; Screen.Cursor:=SynCrDefault ;
if ((Action=taCreat) and (M.LeGuide<>'') and (Not M.ANouveau)) then
   BEGIN
   {Lancements du guide depuis Encaissements, Guides pour Test, Menu}
   if ((M.TypeGuide='DEC') or (M.TypeGuide='ENC')) then LanceGuide(V_PGI.User,M.LeGuide) else
     if ((M.TypeGuide='NOR') and (M.FromGuide)) then LanceGuide(M.TypeGuide,M.LeGuide) else
       if ((M.TypeGuide='ABO') and (M.FromGuide)) then LanceGuide(M.TypeGuide,M.LeGuide) else
         if ((M.TypeGuide='NOR') and (M.SaisieGuidee)) then LanceGuide(M.TypeGuide,M.LeGuide) ;
   END ;
GeneCharge:=False ;
if SuiviMP then
   BEGIN
   PieceModifiee:=True ;
   TraiteOrigMP ;
   END ;
If SuiviMP Then AffecteRef(FALSE) ;
//GP - 07/03/2002  force le redimensionnement de la grille
//HMTrad.ResizeGridColumns(GS);
//HMTrad.Resize(Self);
END ;

procedure TFSaisie.TraiteAffichePCL ;
BEGIN
//RR préparation version S3 Premium.
if (ctxPCL in V_PGI.PGIContexte) {or EstComptaSansAna} then
  if VH^.CPPCLSansAna then BVentil.Visible:=False ;
END ;

procedure TFSaisie.TraiteAfficheGescom ;
BEGIN
if Not (ctxGescom in V_PGI.PGIContexte) then Exit ;
{$IFDEF GCGC}
BGenereTVA.Visible:=False   ; BGenereTVA.Enabled:=False ;
BControleTVA.Visible:=False ; BControleTVA.Enabled:=False ;
BModifTVA.Visible:=False    ; BModifTVA.Enabled:=False ;
BComplement.Visible:=False  ; BComplement.Enabled:=False ;
BZoomJournal.Visible:=False ; BZoomJournal.Enabled:=False ;
BPieceGC.Visible:=False     ; BPieceGC.Enabled:=False ;
bLignesGC.Visible:=False    ; bLignesGC.Enabled:=False ;
BScenario.Visible:=False    ; BScenario.Enabled:=False ;
BDernPieces.Visible:=False  ; BDernPieces.Enabled:=False ;
BMenuGuide.Visible:=False   ; BMenuGuide.Enabled:=False ;
BModifS.Visible:=False      ; BModifS.Enabled:=False ;
BZoomImmo.Visible:=False    ; BZoomImmo.Enabled:=False ;
BZoomTP.Visible:=False      ; BZoomTP.Enabled:=False ;
BSaisTaux.Visible:=False    ; BSaisTaux.Enabled:=False ;
BSwapPivot.Visible:=False   ; BSwapPivot.Enabled:=False ;
{$ENDIF}
END ;

procedure TFSaisie.FormShow(Sender: TObject);
begin
LookLesDocks(Self) ;
TraiteAffichePCL ;
TraiteAfficheGescom ;
ShowSaisie ;
// FQ 13082 : placement du focus à l'ouverture // SBO 16/08/2004
  if ( Action = taCreat ) and                                       // mode création
     not ( (M.TypeGuide='DEC') or (M.TypeGuide='ENC') or SuiviMP )  // ne vient pas de la génération d'encadeca
    then E_JOURNAL.setFocus ;
// Fin FQ 13082
if M.ParCentral then
 E_JOURNALChange(nil) ;
{$IFDEF SCANGED}
{
  CA - 20/02/2007 - Si la GED n'est pas sérialisée, on autorise quand même
                    l'insertion de documents en saisie
}
  if not VH^.OkModGed then
  begin
    if V_PGI.RunFromLanceur then
      InitializeGedFiles(V_PGI.DefaultSectionDbName, SaisieMyAfterImport)
    else
      InitializeGedFiles(V_PGI.DbName, SaisieMyAfterImport);
  end;
{$ENDIF}
end;

{==========================================================================================}
{================================ Ecriture Fichier=========================================}
{==========================================================================================}
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 09/03/2004
Modifié le ... : 09/03/2004
Description .. : TEST si le paramètrage du compte a été modifié depuis son
Suite ........ : chargement dans la grille.
Mots clefs ... :
*****************************************************************}
procedure TFSaisie.IntoucheGene ;
Var lStCpt : String ;
    lInLig : integer ;
    CGen   : TGGeneral ;
    lStSQL : String ;
    lInAxe : Integer ;
begin

  if PasModif then Exit ;
  if M.Simul='N' then Exit ;

  for lInLig := 1 to GS.RowCount - 2 do
    if ( (GS.Cells[ SA_Gen, lInLig ] <> '') and
         (GS.Objects[ SA_Gen, lInLig] <> Nil) ) then
      begin

      // Si compte non crédité, on ne fait pas le test
      CGen := TGGeneral(GS.Objects[SA_Gen,lInLig]) ;
      if ((CGen.TotalDebit<>0) or (CGen.TotalCredit<>0)) then Continue ;

      // Conctruction requête
      lStCpt:=GS.Cells[SA_Gen,lInLig] ;
      lStSQL := 'Select G_GENERAL from GENERAUX Where G_GENERAL="' + lStCpt + '" AND G_NATUREGENE="' + CGen.NatureGene + '"' ;
      if CGen.Lettrable
         then lStSQL := lStSQL + ' AND G_LETTRABLE="X"'
         else lStSQL := lStSQL + ' AND G_LETTRABLE="-"' ;
      if CGen.Collectif
         then lStSQL := lStSQL + ' AND G_COLLECTIF="X"'
         else lStSQL := lStSQL + ' AND G_COLLECTIF="-"' ;
      for lInAxe:=1 to MaxAxe do
        if CGen.Ventilable[lInAxe]
          then lStSQL := lStSQL + ' AND G_VENTILABLE' + IntToStr(lInAxe) + '="X"'
          else lStSQL := lStSQL + ' AND G_VENTILABLE' + IntToStr(lInAxe) + '="-"' ;

      // Test la présence du compte avec les mêmes paramètres qu'au chargement
      if not ExisteSQL( lStSQL ) then
        begin
        // Si non trouvé, le compte a été modifié, déclaration d'une error
        V_PGI.IoError:=oeSaisie ;
        Break ;
        end ;
      end ;

end ;

procedure TFSaisie.IntoucheAux ;
Var lStCpt : String ;
    lInLig : integer ;
    CAux  : TGTiers ;
    lStSQL : String ;
begin

  if PasModif then Exit ;
  if M.Simul='N' then Exit ;

  for lInLig := 1 to GS.RowCount - 2 do
    if ( (GS.Cells[ SA_Aux, lInLig ] <> '') and
         (GS.Objects[ SA_Aux, lInLig] <> Nil) ) then
      begin

      // Si compte non crédité, on ne fait pas le test
      CAux := TGTiers( GS.Objects[ SA_Aux, lInLig ] ) ;
      if ((CAux.TotalDebit<>0) or (CAux.TotalCredit<>0)) then Continue ;

      // Conctruction requête
      lStCpt := GS.Cells[SA_Aux,lInLig] ;
      lStSQL := 'Select T_AUXILIAIRE from TIERS Where T_AUXILIAIRE="' + lStCpt + '" AND T_NATUREAUXI="' + CAux.NatureAux + '"' ;
      if CAux.Lettrable
         then lStSQL := lStSQL + ' AND T_LETTRABLE="X"'
         else lStSQL := lStSQL + ' AND T_LETTRABLE="-"' ;

      // Test la présence du compte avec les mêmes paramètres qu'au chargement
      if not ExisteSQL( lStSQL ) then
        begin
        // Si non trouvé, le compte a été modifié, déclaration d'une error
        V_PGI.IoError:=oeSaisie ;
        Break ;
        end ;
      end ;

end ;

procedure TFSaisie.MajCptesGeneNew ;
Var XD,XC : Double ;
    Lig   : integer ;
    FRM   : TFRM ;
    ll    : LongInt ;
BEGIN
if PasModif then Exit ;
if M.Simul<>'N' then Exit ;
for Lig:=1 to GS.RowCount-2 do
  if ((GS.Cells[SA_Gen,Lig]<>'') and (GS.Objects[SA_Gen,Lig]<>Nil)) then
    BEGIN
    VideTFRM(FRM) ;  //   Fillchar(FRM,SizeOf(FRM),#0) ;
    FRM.Cpt := GS.Cells[SA_Gen,Lig] ;
    RecupTotalPivot( GS, Lig, XD, XC ) ;// SBO 09/08/2007 FQ 20910 maj solde des comptes erroné en devise
    if Not M.ANouveau then
       BEGIN
       FRM.NumD:=NumPieceInt ;
       FRM.DateD:=DateMvt(Lig) ;
       FRM.LigD:=StrToInt(GS.Cells[SA_NumL,Lig]) ;
       AttribParamsNew(FRM,XD,XC,FInfo.TypeExo) ;
       END else
       BEGIN
       FRM.Deb:=XD ; FRM.Cre:=XC ;
       END ;
    if Action<>taConsult then AttribParamsComp(FRM,GS.Objects[SA_Gen,Lig]) ;
    LL:=ExecReqMAJ(fbGene,M.ANouveau,Action<>taConsult,FRM) ;
    If ll<>1 then V_PGI.IoError:=oeSaisie ;
    END ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 21/06/2007
Modifié le ... :   /  /    
Description .. : - LG - 21/06/2007 - le fillchar provoque des fuite memoires 
Suite ........ : ds une boucle
Mots clefs ... : 
*****************************************************************}
procedure TFSaisie.MajCptesAuxNew ;
Var //Cpte  : String ;
    XD,XC : Double ;
    Lig   : integer ;
    Trouv : Boolean ;
    FRM : TFRM ;
    ll : LongInt ;
BEGIN
if PasModif then Exit ;
if M.Simul<>'N' then Exit ;
Trouv:=False ;
for Lig:=1 to GS.RowCount-2 do if ((GS.Cells[SA_Aux,Lig]<>'') and (GS.Objects[SA_Aux,Lig]<>Nil))
    then BEGIN Trouv:=True ; Break ; END ;
if Not Trouv then Exit ;

for Lig:=1 to GS.RowCount-2 do if ((GS.Cells[SA_Aux,Lig]<>'') and (GS.Objects[SA_Aux,Lig]<>Nil)) then
    BEGIN
    VideTFRM(FRM) ; // Fillchar(FRM,SizeOf(FRM),#0) ;
    FRM.Cpt:=GS.Cells[SA_Aux,Lig] ;
    RecupTotalPivot( GS, Lig, XD, XC ) ;// SBO 09/08/2007 FQ 20910 maj solde des comptes erroné en devise
    if Not M.ANouveau then
       BEGIN
       FRM.NumD:=NumPieceInt ;
       FRM.DateD:=DateMvt(Lig) ;
       FRM.LigD:=StrToInt(GS.Cells[SA_NumL,Lig]) ;
       AttribParamsNew(FRM,XD,XC,FInfo.TypeExo) ;
       END else
       BEGIN
       FRM.Deb:=XD ; FRM.Cre:=XC ;
       END ;
    if Action<>taConsult then AttribParamsComp(FRM,GS.Objects[SA_Aux,Lig]) ;
    ll:=ExecReqMAJ(fbAux,M.ANouveau,Action<>taConsult,FRM) ;
    If ll<>1 then V_PGI.IoError:=oeSaisie ;
    END ;
END ;

{procedure TFSaisie.MajCptesSectionNew ;
Var Sect   					: String ;
    NumAxe,NumV,Lig : integer ;
    OBM							: TOBM ;
    OBA							: TOB ;
    XC,XD  					: Double ;
    FRM 						: TFRM ;
    ll 							: LongInt ;
    Trouv						: Boolean ;
BEGIN
	if PasModif then Exit ;
	if M.Simul<>'N' then Exit ;
	// Y'a-t-il eu ventilation ?
  Trouv:=False ;
  for Lig:=1 to GS.RowCount-2 do
  	if GS.Cells[SA_Gen,Lig]<>'' then
      BEGIN
      OBM:=GetO(GS,Lig) ;
      if OBM<>Nil then
        BEGIN
        Trouv := VentilationExiste(OBM) ;
        Break ;
        END ;
      END ;
  if Not Trouv then Exit ;

  // Parcours des écritures
	for Lig:=1 to GS.RowCount-2 do
  	if GS.Cells[SA_Gen,Lig]<>'' then
    	BEGIN
    	OBM:=GetO(GS,Lig) ;
      if OBM=Nil then Continue ;
      // Parcours des axes
    	for NumAxe:=1 to OBM.Detail.Count do
        for NumV:=0 to OBM.Detail[NumAxe-1].Detail.Count-1 do
          BEGIN
          OBA 	:= OBM.Detail[NumAxe-1].Detail[NumV];
          Sect	:= OBA.GetValue('Y_SECTION') ;
          XD		:= OBA.GetValue('Y_DEBIT') ;
          XC		:= OBA.GetValue('Y_CREDIT') ;
          Fillchar(FRM,SizeOf(FRM),#0) ;
          FRM.Cpt:= Sect ;
          FRM.Axe:= 'A'+Chr(48+NumAxe) ;
          if Not M.ANouveau then
             BEGIN
             FRM.NumD:=NumPieceInt ;
             FRM.DateD:=DateMvt(Lig) ;
             FRM.LigD:=StrToInt(GS.Cells[SA_NumL,Lig]) ;
             AttribParamsNew(FRM,XD,XC,GeneTypeExo) ;
             END else
             BEGIN
             FRM.Deb:=XD ; FRM.Cre:=XC ;
             END ;
          LL:=ExecReqMAJNew(fbSect,M.ANouveau,False,FRM) ;
          If ll<>1 then V_PGI.IoError:=oeSaisie ;
          END ;
    	END ;
END ;  }

function TFSaisie.QuelEnc ( Lig : Integer ) : String3 ;
var CGen : TGGeneral ;
begin
  CGen:=GetGGeneral(GS,Lig) ;
  // Pour les comptes lettrables divers... FQ 16136 SBO 09/08/2005
  if ( CGen <> Nil ) and CGen.Lettrable and ( CGen.NatureGene = 'DIV' )
    then result := 'RIE'
    // Sinon récultat habituel
    else Result := SensEnc(ValD(GS,Lig),ValC(GS,Lig)) ;
END ;

Procedure TFSaisie.GetCptsContreP ( Lig : Integer ; Var LeGene,LeAux,TL0 : String17 ) ;
Var k : integer ;
    CAux : TGTiers ;
BEGIN
TL0:='' ;
LeGene:='' ; LeAux:='' ;
if OkBqe then
   BEGIN
   {Pièce de banque}
   if GS.Cells[SA_Gen,Lig]=GeneTreso then
      BEGIN
      {Sur cpte de banque, contrep=première ligne non banque au dessus}
      for k:=Lig-1 downto 1 do if GS.Cells[SA_Gen,k]<>GeneTreso then
          BEGIN
          LeGene:=GS.Cells[SA_Gen,k] ;
          LeAux:=GS.Cells[SA_Aux,k] ; // Modif SBO 06/10/2004 : E_CONTREPARTIEAUX pas toujours renseigné !
          Break ;
          END ;
      END else
      BEGIN
      {Sur cpte non banque, contrep=première ligne banque au dessous}
      for k:=Lig+1 to GS.RowCount-1 do if GS.Cells[SA_Gen,k]=GeneTreso then
          BEGIN
          LeGene:=GS.Cells[SA_Gen,k] ; LeAux:=GS.Cells[SA_Aux,k] ;
          Break ;
          END ;
      END ;
   END else if OkJalEffet then
   BEGIN
   {Pièce sur Journal d'effet}
   if GS.Cells[SA_Gen,Lig]=GeneTreso then
      BEGIN
      {Sur cpte de banque, contrep=première ligne non banque au dessus}
      for k:=Lig-1 downto 1 do if GS.Cells[SA_Gen,k]<>GeneTreso then
          BEGIN
          LeGene:=GS.Cells[SA_Gen,k] ; LeAux:=GS.Cells[SA_Aux,k] ;
          Break ;
          END ;
      END else
      BEGIN
      {Sur cpte non banque, contrep=première ligne effet au dessous}
      for k:=Lig+1 to GS.RowCount-1 do if GS.Cells[SA_Gen,k]=GeneTreso then
          BEGIN
          LeGene:=GS.Cells[SA_Gen,k] ; LeAux:=GS.Cells[SA_Aux,k] ;
          Break ;
          END ;
      END ;
   END else
   BEGIN
   {Pièce non banque}
   if ((M.Section<>'') and (M.General='') and ((M.TypeGuide='ENC') or (M.TypeGuide='DEC'))) then
      BEGIN
      {Génération Encadéca non banque type "403"}
      if GS.Cells[SA_Gen,Lig]=M.Section then
         BEGIN
         {Lecture arrière pour trouver "401"}
         for k:=Lig-1 downto 1 do if GS.Cells[SA_Gen,k]<>M.Section then
             BEGIN
             LeGene:=GS.Cells[SA_Gen,k] ; LeAux:=GS.Cells[SA_Aux,k] ;
             Break ;
             END ;
         if  LeGene='' then
             BEGIN
             {Si pas trouvé, lecture avant pour trouver "401"}
             for k:=Lig+1 to GS.RowCount-2 do if GS.Cells[SA_Gen,k]<>M.Section then
                 BEGIN
                 LeGene:=GS.Cells[SA_Gen,k] ; LeAux:=GS.Cells[SA_Aux,k] ;
                 Break ;
                 END ;
             END ;
         END else
         BEGIN
         {Lecture avant pour trouver "403"}
         for k:=Lig+1 to GS.RowCount-2 do if GS.Cells[SA_Gen,k]=M.Section then
             BEGIN
             LeGene:=GS.Cells[SA_Gen,k] ; LeAux:=GS.Cells[SA_Aux,k] ;
             Break ;
             END ;
         if  LeGene='' then
             BEGIN
             {Lecture arrière pour trouver "403"}
             for k:=Lig-1 downto 1 do if GS.Cells[SA_Gen,k]=M.Section then
                 BEGIN
                 LeGene:=GS.Cells[SA_Gen,k] ; LeAux:=GS.Cells[SA_Aux,k] ;
                 Break ;
                 END ;
             END ;
         END ;
      END else
      BEGIN
      {Cas normal}
      if isTiers(GS,Lig) then
         BEGIN
         {Lecture avant pour trouver "HT"}
         if GS.Objects[SA_Aux,Lig]<>Nil then BEGIN CAux:=TGTiers(GS.Objects[SA_Aux,Lig]) ; TL0:=CAux.TL0 ; END ;
         for k:=Lig+1 to GS.RowCount-2 do if isHT(GS,k,True) then
             BEGIN
             LeGene:=GS.Cells[SA_Gen,k] ;
             Break ;
             END ;
         if  LeGene='' then
             BEGIN
             {Lecture arrière pour trouver "HT"}
             for k:=Lig-1 downto 1 do if isHT(GS,k,True) then
                 BEGIN
                 LeGene:=GS.Cells[SA_Gen,k] ;
                 Break ;
                 END ;
             END ;
         END else
         BEGIN
         {Lecture arrière pour trouver "Tiers"}
         for k:=Lig-1 downto 1 do if isTiers(GS,k) then
             BEGIN
             LeGene:=GS.Cells[SA_Gen,k] ; LeAux:=GS.Cells[SA_Aux,k] ;
             if GS.Objects[SA_Aux,k]<>Nil then BEGIN CAux:=TGTiers(GS.Objects[SA_Aux,k]) ; TL0:=CAux.TL0 ; END ;
             Break ;
             END ;
         if  LeGene='' then
             BEGIN
             {Lecture avant pour trouver "Tiers"}
             for k:=Lig+1 to GS.RowCount-2 do if isTiers(GS,k) then
                 BEGIN
                 LeGene:=GS.Cells[SA_Gen,k] ; LeAux:=GS.Cells[SA_Aux,k] ;
                 if GS.Objects[SA_Aux,k]<>Nil then BEGIN CAux:=TGTiers(GS.Objects[SA_Aux,k]) ; TL0:=CAux.TL0 ; END ;
                 Break ;
                 END ;
             END ;
         END ;
      END ;
   END ;
{Cas particuliers}
if LeGene<>'' then Exit ;
if ((OkBqe) and (GS.Cells[SA_Gen,Lig]<>GeneTreso)) then LeGene:=GeneTreso else
 if ((M.General<>'') and (GS.Cells[SA_Gen,Lig]<>M.General)) then LeGene:=M.General else
  if Not IsBQE(GS,Lig) then
     BEGIN
     for k:=Lig+1 to GS.RowCount-2 do if isBQE(GS,k) then BEGIN LeGene:=GS.Cells[SA_Gen,k] ; Break ; END ;
     if LeGene='' then for k:=Lig-1 downto 1 do if isBQE(GS,k) then BEGIN LeGene:=GS.Cells[SA_Gen,k] ; Break ; END ;
     END else
     BEGIN
     for k:=Lig-1 downto 1 do if isTiers(GS,k) then BEGIN LeGene:=GS.Cells[SA_Gen,k] ; LeAux:=GS.Cells[SA_Aux,k] ; Break ; END ;
     if LeGene='' then for k:=Lig+1 to GS.RowCount-2 do if isTiers(GS,k) then BEGIN LeGene:=GS.Cells[SA_Gen,k] ; LeAux:=GS.Cells[SA_Aux,k] ; Break ; END ;
     END ;
{Si rien trouvé, swaper les lignes 1 et 2}
if LeGene<>'' then Exit ;
if Lig>1 then BEGIN LeGene:=GS.Cells[SA_Gen,1] ; LeAux:=GS.Cells[SA_Aux,1] ; END else
 if GS.RowCount>=4 then BEGIN LeGene:=GS.Cells[SA_Gen,2] ; LeAux:=GS.Cells[SA_Aux,2] ; END ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 12/03/2004
Modifié le ... : 12/03/2004
Description .. : - LG - 12/03/2004 - FB 13341 - affectation de E_IO a X
Mots clefs ... :
*****************************************************************}
Procedure TFSaisie.RecupTronc ( Lig : Integer ; MO : TMOD ; Var OBM : TOBM);
Var CGen : TGGeneral ;
    CAux : TGTiers ;
    OBA  : TOB ;
    TVA,TPF  : String3 ;
    LeGen,LeAux,TL0 : String17 ;
    NumAxe,NumV{,iEche} : integer ;
BEGIN
  if PasModif then exit ;
  CGen:=NIL ; CAux:=NIL ; OBM:=Nil ;
  if GS.Objects[SA_Gen,Lig]<>Nil then BEGIN CGen:=TGGeneral(GS.Objects[SA_Gen,Lig]) ; END ;
  if GS.Objects[SA_Aux,Lig]<>Nil then BEGIN CAux:=TGTiers(GS.Objects[SA_Aux,Lig]) ; END ;
  if GS.Objects[SA_Exo,Lig]<>Nil then OBM:=GetO(GS,Lig) ;
  if OBM = nil then Exit;

  {Teste CFONB}
  if Not SuiviMP then
    BEGIN
    if ((M.TypeGuide='ENC') or (M.TypeGuide='DEC')) and
       ((M.ExportCFONB) or (M.Bordereau)) and (OBM.LC.Count>0)
      then OBM.PutMvt('E_CFONBOK','#') ;
    END
  else
    if ((M.ExportCFONB) or (M.Bordereau)) and (OBM.GetMvt('E_TRACE')<>'')
      then OBM.PutMvt('E_CFONBOK','#') ;

  OBM.SetCotation(DatePiece) ;
  if (not SuiviMP) or (M.MPGUnique <> '') then OBM.SetMPACC ; // FQ 12438 //SG6

  OBM.PutMvt('E_DATECOMPTABLE',DatePiece) ;
  OBM.PutMvt('E_EXERCICE',QuelExoDT(DatePiece)) ;
  {$IFNDEF SPEC302}
    OBM.PutMvt('E_PERIODE',GetPeriode(DatePiece)) ;
    OBM.PutMvt('E_SEMAINE',NumSemaine(DatePiece)) ;
  {$ENDIF}
  OBM.PutMvt('E_QUALIFPIECE',M.Simul) ;
  OBM.PutMvt('E_NATUREPIECE',E_NATUREPIECE.Value) ;
  OBM.PutMvt('E_CONTROLETVA','RIE') ;
  OBM.PutMvt('E_ETABLISSEMENT',E_ETABLISSEMENT.Value) ;

  {JP 06/10/05 : Repris de ZoneObli dans le cadre de la FQ 13250}
  OBM.PutMvt('E_REFINTERNE',Copy(GS.Cells[SA_RefI,Lig],1,35)) ;
  OBM.PutMvt('E_LIBELLE',Copy(GS.Cells[SA_Lib,Lig],1,35)) ;

  // Modif Lettrage des comptes divers  FQ 16136 SBO 09/08/2005
  CGetTypeMvt( OBM, FInfo ) ;

  if (OBM.GetMvt('E_REGIMETVA')='') or (VH^.PaysLocalisation=CodeISOES) then
     OBM.PutMvt('E_REGIMETVA',GeneRegTVA) ; //XVI 24/02/2005

  if Not M.ANouveau then
     BEGIN
     OBM.PutMvt('E_ECRANOUVEAU','N') ;
     if ValideEcriture then OBM.PutMvt('E_VALIDE','X')
                       else OBM.PutMvt('E_VALIDE','-') ;
     END
  else
     BEGIN
     OBM.PutMvt('E_ECRANOUVEAU','H') ;
     OBM.PutMvt('E_VALIDE','X') ;
     END ;

  OBM.PutMvt('E_BUDGET','') ;
  OBM.PutMvt('E_ECHE','-') ;

  if CAux<>Nil then
     BEGIN
     if GeneRegTVA<>'' then OBM.PutMvt('E_REGIMETVA',GeneregTVA)
                       else OBM.PutMvt('E_REGIMETVA',CAux.RegimeTVA) ;
     END ;

  if CGen<>Nil then
     BEGIN
     TVA:=CGen.Tva ; if TVA='' then TVA:=GeneTVA ;
     TPF:=CGen.Tpf ; if TPF='' then TPF:=GeneTPF ;
     if (OBM.GetMvt('E_TVA')='') or (VH^.PaysLocalisation=CodeISOES) then OBM.PutMvt('E_TVA',GeneTva) ; //XVI 24/02/2005
     if (OBM.GetMvt('E_TPF')='') or (VH^.PaysLocalisation=CodeISOES) then OBM.PutMvt('E_TPF',GeneTpf) ;
     OBM.PutMvt('E_BUDGET',CGen.Budgene) ;

     // compte lettrable
     if ((Lettrable(CGen)=MonoEche) and (MO<>Nil)) then
        BEGIN
        // récupération de la 1ère échéance
        RecupEche( Lig, MO.MODR , 1 , OBM, True ) ; // FQ 18736 SBO 06/09/2006 : pb d'affectation des zones de lettrage pour les comptes divers lettrables 
        // Ajout Sinon bug LCD FQ 16584 SBO 06/08/2005
        if CGen.pointable then
          OBM.PutMvt('E_ETATLETTRAGE',	'RI') ;  // FQ 16761 SBO 30/09/2005
        END ;
     END ;

  GetCptsContreP(Lig,LeGen,LeAux,TL0) ;
  If (TL0<>'') And EstSpecif('51196') Then OBM.PutMvt('E_TABLE0',TL0) ;
  OBM.PutMvt('E_TRACE','') ;
  OBM.PutMvt('E_CONTREPARTIEGEN',LeGen) ;
  OBM.PutMvt('E_CONTREPARTIEAUX',LeAux) ;
  OBM.PutMvt('E_DATEMODIF',NowFutur) ;
  OBM.PutMvt('E_CONFIDENTIEL',ModeConf) ;
  OBM.PutMvt('E_IO','X') ;
  {Mise à jour contreparties analytiques}
  if OBM.Detail.Count>0 then
  	for NumAxe:=0 to OBM.Detail.Count-1 do
      for NumV:=0 to OBM.Detail[NumAxe].Detail.Count-1 do
        BEGIN
        OBA := OBM.Detail[NumAxe].Detail[NumV];
        OBA.PutValue('Y_CONTREPARTIEGEN',OBM.GetMvt('E_CONTREPARTIEGEN')) ;
        OBA.PutValue('Y_CONTREPARTIEAUX',OBM.GetMvt('E_CONTREPARTIEAUX')) ;
        {JP 09/11/07 : FQ 21812 : mise à jour de l'auxiliaire}
        OBA.PutValue('Y_AUXILIAIRE', OBM.GetMvt('E_AUXILIAIRE'));
        END ;
  //YMO 17/01/2006 FQ17245
  OBM. PutMvt ('E_EXPORTE','---') ;

END ;

Procedure TFSaisie.RecupAnal(Lig : integer ; var OBA : TOB ; NumAxe,NumV : integer ) ;
BEGIN
  if PasModif then Exit ;
  // Attention NumV démarre à 0
  if NumV=0
    then OBA.PutValue('Y_TYPEMVT', 'AE')
    else OBA.PutValue('Y_TYPEMVT', 'AL') ;

  OBA.PutValue('Y_TYPEANALYTIQUE',  '-') ;
  OBA.PutValue('Y_DATEMODIF',	    NowFutur) ;
  OBA.PutValue('Y_DATECOMPTABLE',   DatePiece) ;
  OBA.PutValue('Y_EXERCICE',	    QuelExoDT(DatePiece)) ;
{$IFNDEF SPEC302}
  OBA.PutValue('Y_PERIODE',         GetPeriode(DatePiece)) ;
  OBA.PutValue('Y_SEMAINE',         NumSemaine(DatePiece)) ;
{$ENDIF}
  OBA.PutValue('Y_NATUREPIECE',     E_NATUREPIECE.Value) ;
  OBA.PutValue('Y_ETABLISSEMENT',   E_ETABLISSEMENT.Value) ;
  OBA.PutValue('Y_CONFIDENTIEL',    ModeConf) ;
  if Not M.ANouveau then OBA.PutValue('Y_ECRANOUVEAU','N')
                    else OBA.PutValue('Y_ECRANOUVEAU','H') ;

  // FQ 16867 : SBO 13/10/2005 report des infos sur analytique
  OBA.PutValue('Y_REFINTERNE',  GS.Cells[SA_RefI,Lig] ) ;
  OBA.PutValue('Y_LIBELLE',     GS.Cells[SA_Lib,Lig]  ) ;
  OBA.PutValue('Y_GENERAL',     GS.Cells[SA_Gen,Lig]  ) ; // FQ 10296
END ;


Procedure TFSaisie.RecupEche(Lig : integer ; R : T_ModeRegl ; NumEche : integer; Var OBM : TOBM ; vBoMonoEche : boolean ) ;
Var Deb  : boolean ;
    i    : integer ;
    Coef : Double ;
    Coll : String ;
BEGIN
  if PasModif then Exit ; //Coef:=1.0 ;
  OBM.PutMvt('E_NUMECHE',       NumEche) ;
  OBM.PutMvt('E_ECHE',          'X') ;
  OBM.PutMvt('E_NIVEAURELANCE',	R.TabEche[NumEche].NiveauRelance) ;
  OBM.PutMvt('E_MODEPAIE',	R.TabEche[NumEche].ModePaie) ;
  OBM.PutMvt('E_DATEECHEANCE',	R.TabEche[NumEche].DateEche) ;
  OBM.PutMvt('E_DATEVALEUR',	R.TabEche[NumEche].DateValeur) ;
  OBM.PutMvt('E_COUVERTURE',	R.TabEche[NumEche].Couverture) ;
  OBM.PutMvt('E_COUVERTUREDEV',	R.TabEche[NumEche].CouvertureDev) ;
  OBM.PutMvt('E_ETATLETTRAGE',	R.TabEche[NumEche].EtatLettrage) ;
  OBM.PutMvt('E_LETTRAGE',	R.TabEche[NumEche].CodeLettre) ;
  OBM.PutMvt('E_LETTRAGEDEV',	R.TabEche[NumEche].LettrageDev) ;
  OBM.PutMvt('E_DATEPAQUETMAX',	R.TabEche[NumEche].DatePaquetMax) ;
  OBM.PutMvt('E_DATEPAQUETMIN',	R.TabEche[NumEche].DatePaquetMin) ;
  OBM.PutMvt('E_DATERELANCE',	R.TabEche[NumEche].DateRelance) ;

  // SBO 03/12/2006 : Pb affectation des montants si utilisation d'un MR par défaut multi-échéance
  if not vBoMonoEche then
    begin
    OBM.PutMvt('E_DEBIT',0) ;  	OBM.PutMvt('E_CREDIT',0) ;
    OBM.PutMvt('E_DEBITDEV',0) ;	OBM.PutMvt('E_CREDITDEV',0) ;
    Deb:=(ValD(GS,Lig)<>0) ;
    if Deb then
      BEGIN
      OBM.PutMvt('E_DEBIT',       R.TabEche[NumEche].MontantP) ;
      OBM.PutMvt('E_DEBITDEV',    R.TabEche[NumEche].MontantD) ;
      END
    else
      BEGIN
      OBM.PutMvt('E_CREDIT',      R.TabEche[NumEche].MontantP) ;
      OBM.PutMvt('E_CREDITDEV',   R.TabEche[NumEche].MontantD) ;
      END ;
    if ModeDevise then
      BEGIN
      if Deb
        then OBM.PutMvt('E_DEBITDEV',  R.TabEche[NumEche].MontantD)
        else OBM.PutMvt('E_CREDITDEV', R.TabEche[NumEche].MontantD) ;
      END ;
    end ;

  OBM.PutMvt('E_ENCAISSEMENT',       QuelEnc(Lig)) ;
  OBM.PutMvt('E_ORIGINEPAIEMENT',    R.TabEche[NumEche].DateEche) ;

  if (not SuiviMP) or (M.MPGUnique <> '') then OBM.SetMPACC; // FQ 12438   //SG6 30/11/04 FQ 15030

  {#TVAENC}
  if OuiTvaLoc then
    BEGIN
    Coll:=OBM.GetMvt('E_GENERAL') ;
    if EstCollFact(Coll) then
      BEGIN
      if ((SorteTva<>stvDivers) and (UnSeulTiers) and (R.TotalAPayerP<>0)) then
        BEGIN
        Coef:=R.TabEche[NumEche].MontantP/R.TotalAPayerP ;
        for i:=1 to 4 do OBM.PutMvt('E_ECHEENC'+IntToStr(i),Arrondi(TabTvaEnc[i]*Coef,V_PGI.OkDecV)) ;
        OBM.PutMvt('E_ECHEDEBIT',Arrondi(TabTvaEnc[5]*Coef,V_PGI.OkDecV)) ;
        OBM.PutMvt('E_EMETTEURTVA','X') ;
        END
      else if ((SorteTva=stvDivers) and (R.ModifTva)) then
        BEGIN
        for i:=1 to 4 do OBM.PutMvt('E_ECHEENC'+IntToStr(i),Arrondi(R.TabEche[NumEche].TAV[i],V_PGI.OkDecV)) ;
        OBM.PutMvt('E_ECHEDEBIT',Arrondi(R.TabEche[NumEche].TAV[5],V_PGI.OkDecV)) ;
        END ;
      END ;
    END ;

END ;

Procedure TFSaisie.GetAna(Lig : Integer) ;
Var i,j : integer ;
    OBM : TOBM ;
    OBA : TOB ;
BEGIN
  if PasModif then exit ;
  OBM := GetO(GS,Lig) ;
  if OBM = Nil then Exit;
  for i:=0 to OBM.Detail.Count-1 do
    for j:=0 to OBM.Detail[i].Detail.Count-1 do
      BEGIN
      OBA := OBM.Detail[i].Detail[j];
      RecupAnal(Lig,OBA,i,j) ;
      EcartChange(OBA) ;
      if revision then
        // SBO 01/07/2007 : enregistrement spécifique pour pb CWAS
        if FBoInsertSpecif
          then begin
               if not CTobInsertDB( OBA ) then
                 V_PGI.IOError := oeSaisie ;
               end
          else OBA.InsertDB(nil);
      END ;
END ;


Procedure TFSaisie.EcartChange(Obj : TOB) ;
Var stPref : String;
begin
  if PasModif then Exit ;
  if E_NATUREPIECE.Value<>'ECC' then Exit ;

  if Obj.NomTable = 'ECRITURE'
    then stPref := 'E'
    else if Obj.NomTable = 'ANALYTIQ'
      then stPref := 'Y'
      else Exit;

  Obj.PutValue(stPref+'_DEBITDEV',  0) ;
  Obj.PutValue(stPref+'_CREDITDEV', 0) ;
// FQ 18704
  Obj.PutValue(stPref+'_DEVISE',    E_DEVISE.Value) ;
{
  Obj.PutValue(stPref+'_DEVISE',    E_DEVISE_.Value) ;
}
end ;

Procedure TFSaisie.GetEcr(Lig : Integer) ;
Var M    : TMOD ;
    R    : T_ModeRegl ;
    i    : integer ;
    CGen : TGGeneral ;
    OBM  : TOBM;
BEGIN
  if PasModif then Exit ;
  M := TMOD(GS.Objects[SA_NumP,Lig]) ;
  CGen := GetGGeneral(GS,Lig) ;
  if CGen=Nil then Exit ;

  if ((M=Nil) or (Lettrable(CGen)=MonoEche)) then
    BEGIN
    // Récupération et remplissage TOB Ecritures générales
    RecupTronc(Lig,M,OBM);
    EcartChange(OBM);
    // Finalisation des TOBs Analytiques
    GetAna(Lig) ;
    OnUpdateEcritureTOB( OBM , Action ,[cEcrCompl] ) ;
    // Enregistrement Base
    {$IFNDEF EAGLCLIENT}
      OBM.SetAllModifie(TRUE) ; // FQ 18712 : SBO 06/09/2006 Maj du statut des champs sinon pb de BlocNote
    {$ENDIF EAGLCLIENT}
    // SBO 01/07/2007 : enregistrement spécifique pour pb CWAS
    if FBoInsertSpecif
      then begin
           if not CTobInsertDB( OBM ) then
             V_PGI.IOError := oeSaisie ;
           end
      else OBM.InsertDB(nil);

    END
  else
    BEGIN
    R := M.MODR ;
    for i:=1 to R.NbEche do
      BEGIN
      // Récupération et remplissage TOB Ecritures générales
      RecupTronc(Lig,M,OBM);
      RecupEche(Lig,R,i,OBM);
      EcartChange(OBM);
      // Finalisation des TOBs Analytiques
      if i=1 then
        GetAna(Lig) ;
      OnUpdateEcritureTOB( OBM , Action ,[cEcrCompl]) ;
      // Enregistrement Base
      {$IFNDEF EAGLCLIENT}
        OBM.SetAllModifie(TRUE) ; // FQ 18712 : SBO 06/09/2006 Maj du statut des champs sinon pb de BlocNote
      {$ENDIF EAGLCLIENT}
      // SBO 01/07/2007 : enregistrement spécifique pour pb CWAS
      if FBoInsertSpecif
        then begin
             if not CTobInsertDB( OBM ) then
               V_PGI.IOError := oeSaisie ;
             end
        else OBM.InsertDB(nil);
      END ;
    END ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/02/2004
Modifié le ... :   /  /
Description .. : -09/02/2004 - affectation de E_IO a X en modification
Mots clefs ... :
*****************************************************************}
procedure TFSaisie.RecupFromGrid( Lig : integer ; MO : TMOD ; NumE : integer ) ;
Var OBM  : TOBM ;
    St   : String ;
    R    : T_MODEREGL ;
    ie,k{,iEche} : integer ;
    Sens : integer ;
    SQL  : String ;
    LeGen,LeAux,TL0 : String17 ;
    Stamp  : TDateTime ;
    XStamp : TDELMODIF ;
BEGIN
OBM:=GetO(GS,Lig) ;
if ValD(GS,Lig)<>0 then Sens:=1 else Sens:=2 ;
if MO<>Nil then
   BEGIN
   R:=MO.MODR ; ie:=NumE ; if ie<=0 then ie:=1 ;
   if Sens=1 then
      BEGIN
      OBM.PutMvt('E_DEBIT',R.TabEche[ie].MontantP) ; OBM.PutMvt('E_CREDIT',0) ;
      OBM.PutMvt('E_DEBITDEV',R.TabEche[ie].MontantD) ; OBM.PutMvt('E_CREDITDEV',0) ;
      END else
      BEGIN
      OBM.PutMvt('E_CREDIT',R.TabEche[ie].MontantP) ; OBM.PutMvt('E_DEBIT',0) ;
      OBM.PutMvt('E_CREDITDEV',R.TabEche[ie].MontantD) ; OBM.PutMvt('E_DEBITDEV',0) ;
      END ;
   OBM.PutMvt('E_NUMECHE',ie) ; OBM.PutMvt('E_MODEPAIE',R.TabEche[ie].ModePaie) ;
   OBM.PutMvt('E_DATEECHEANCE',R.TabEche[ie].DateEche) ;
   OBM.PutMvt('E_DATEVALEUR',R.TabEche[ie].DateValeur) ;
   OBM.PutMvt('E_ECHE','X') ;
   OBM.PutMvt('E_ENCAISSEMENT',QuelEnc(Lig)) ;
   OBM.PutMvt('E_ORIGINEPAIEMENT',R.TabEche[ie].DateEche) ;
   END else
   BEGIN
   OBM.PutMvt('E_ECHE','-') ; //ie:=0 ;
   END ;

  // Modif Lettrage des comptes divers SBO 09/09/2005  FQ 16136 SBO 09/08/2005
  CGetTypeMvt( OBM, FInfo ) ;

OBM.PutMvt('E_VISION','DEM') ;
OBM.PutMvt('E_NATUREPIECE',E_NATUREPIECE.Value) ;
OBM.PutMvt('E_CONTROLETVA','RIE') ;
OBM.PutMvt('E_ETABLISSEMENT',E_ETABLISSEMENT.Value) ;
OBM.PutMvt('E_DATECOMPTABLE',DatePiece) ;
{$IFNDEF SPEC302}
OBM.PutMvt('E_PERIODE',GetPeriode(DatePiece)) ;
OBM.PutMvt('E_SEMAINE',NumSemaine(DatePiece)) ;
{$ENDIF}
OBM.PutMvt('E_CONFIDENTIEL',ModeConf) ;
if ((M.ANouveau) or (ValideEcriture)) then OBM.PutMvt('E_VALIDE','X') else OBM.PutMvt('E_VALIDE','-') ;
GetCptsContreP(Lig,LeGen,LeAux,TL0) ;
OBM.PutMvt('E_CONTREPARTIEGEN',LeGen) ; OBM.PutMvt('E_CONTREPARTIEAUX',LeAux) ;
If (TL0<>'') And EstSpecif('51196') Then OBM.PutMvt('E_TABLE0',TL0) ;
EcartChange(OBM) ;
OnUpdateEcritureTOB( OBM , Action , [cEcrCompl]) ;
OBM.PutMvt('E_IO','X') ;
St:=OBM.StPourUpdate ; if St='' then Exit ;
SQL:='UPDATE ECRITURE SET '+St+', E_DATEMODIF="'+UsTime(NowFutur)+'"'
   +' Where  '+WhereEcriture(tsGene,M,False)+' AND E_NUMLIGNE='+IntToStr(Lig) ;
if NumE>0 then SQL:=SQL+' AND E_NUMECHE='+IntToStr(NumE) ;
{Réseau}
Stamp:=0 ;
for k:=Lig-1 to TDELECR.Count-1 do
    BEGIN
    XStamp:=TDELMODIF(TDELECR[k]) ;
    if ((XStamp.NumLigne=Lig) and ((XStamp.NumEcheVent=NumE) or (NumE=0))) then
       BEGIN  ;
       Stamp:=XStamp.DateModification ; Break ;
       END ;
    END ;
if Stamp>0 then SQL:=SQL+' AND E_DATEMODIF="'+UsTime(Stamp)+'"' ;
{Lancement}
if ExecuteSQL(SQL)<=0 then V_PGI.IOError:=oeSaisie ;
END ;

Procedure TFSaisie.GetEcrGrid(Lig : Integer) ;
Var M    : TMOD ;
    R    : T_ModeRegl ;
    i    : integer ;
    CGen : TGGeneral ;
    t    : TLettre ;
BEGIN
if PasModif then exit ;
M:=TMOD(GS.Objects[SA_NumP,Lig]) ; CGen:=GetGGeneral(GS,Lig) ; if CGen=Nil then Exit ;
t:=Lettrable(CGen) ;
if ((M=Nil) or (t=MonoEche)) then RecupFromGrid(Lig,M,0) else
   BEGIN
   R:=M.MODR ; for i:=1 to R.NbEche do RecupFromGrid(Lig,M,i) ;
   END ;
END ;

procedure TFSaisie.ValideLignesGene ;
Var i : integer ;
BEGIN
if PasModif then Exit ;
InitMove(GS.RowCount-2,'') ;
if Revision then
   BEGIN
   for i:=1 to GS.RowCount-2 do
       BEGIN
       MoveCur(FALSE) ;
      GetEcrGrid(i) ;
      GetAna(i) ;
      if V_PGI.IoError<>oeOk then Break ;
      END ;
   END else
   BEGIN
   for i:=1 to GS.RowCount-2 do
       BEGIN
       MoveCur(FALSE) ;
       GetEcr(i) ;
       if V_PGI.IoError<>oeOk then Break ;
       END ;
   END ;
FiniMove ;
END ;

procedure TFSaisie.ValideLeJournalNew ;
Var Per  : Byte ; // 14047
    DD   : TDateTime ;
    FRM : TFRM ;
    ll : LongInt ;
BEGIN
if PasModif then Exit ;
if M.Simul<>'N' then Exit ;
Fillchar(FRM,SizeOf(FRM),#0) ;
FRM.CPT:=SAJAL.Journal ;
FRM.NumD:=NumPieceInt ;
DD:=DateMvt(1) ; FRM.DateD:=DD ;
if Not M.ANouveau then
   BEGIN
   AttribParamsNew(FRM,SI_TotDP,SI_TotCP,FInfo.TypeExo) ;
   END else
   BEGIN
   FRM.DE:=SI_TotDP ; FRM.CE:=SI_TotCP ;
   FRM.Deb:=0 ; FRM.Cre:=0 ;
   FRM.DS:=0 ; FRM.CS:=0 ;
   FRM.DP:=0 ; FRM.CP:=0 ;
   END ;
ll:=ExecReqMAJ(fbJal,M.ANouveau,False,FRM) ;
If ll<>1 then V_PGI.IoError:=oeSaisie ;
{Dévalidation éventuelle période+jal}
if FInfo.TypeExo=teEncours then // Ne doit pas être dévalidé FQ 14047
//Lek 27/09/05 FQ16116 c'été une erreur il faut remettre
   BEGIN
   Per:=QuellePeriode(DD,VH^.Encours.Deb) ;
   if SAJAL.ValideEN[Per]='X' then ADevalider(E_JOURNAL.Value,DD) ;
   END else
   BEGIN
   Per:=QuellePeriode(DD,VH^.Suivant.Deb) ;
   if SAJAL.ValideEN1[Per]='X' then ADevalider(E_JOURNAL.Value,DD) ;
   END ;
{$IFDEF EAGLCLIENT}
{$ELSE}
	MarquerPublifi(True) ; //EAGL : EN ATTENTE
{$ENDIF}
END ;

procedure TFSaisie.ValideLesComptes ;
var
  //SG6 27.01.05 Gestioon mode croisaxe
  OBM : TOBM ;
  lInCpt , lInCptAxe, lInCptVentil  : integer ;
BEGIN
  if PasModif then exit ;
  if M.Simul='N' then
    BEGIN
    MajCptesGeneNew ;
    MajCptesAuxNew ;

    //SG6 27.01.05 Gestion mode croisaxe
//    MajCptesSectionNew ;
    for lInCpt := 1 to GS.RowCount - 2 do
    begin
      if GS.Cells[SA_Gen,lInCpt] <> '' then
      begin
        OBM := GetO(GS,lInCpt) ;
        if OBM=Nil then Continue ;
        for lInCptAxe := 0 to OBM.Detail.Count - 1 do
        begin
          for lInCptVentil := 0 to OBM.Detail[lInCptAxe].Detail.Count -1 do
          begin
            MajSoldeSectionTOB(OBM.Detail[lInCptAxe].Detail[lInCptVentil],true);
          end;
        end;

      end;
    end;

    END
  else
    BEGIN
    IntoucheGene ;
    IntoucheAux ;
    END ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 21/06/2007
Modifié le ... :   /  /    
Description .. : - LG - 21/06/2007 - le filchar provoque des fuites ds une 
Suite ........ : boucle
Mots clefs ... : 
*****************************************************************}
procedure TFSaisie.InverseSoldeNew ;
Var i    : integer ;
    T    : T_SC ;
    FRM : TFRM ;
    ll : LongInt ;
BEGIN
if PasModif then Exit ;
if M.Simul<>'N' then Exit ;
// Généraux
for i:=0 to TS.Count-1 do
    BEGIN
    VideTFRM(FRM) ;
    //Fillchar(FRM,SizeOf(FRM),#0) ;
    T:=T_SC(TS[i]) ;
    FRM.Cpt:=T.Cpte ; FRM.Deb:=T.Debit ; FRM.Cre:=T.Credit ; FRM.Axe:=T.Identi ;
    if FRM.Axe='G' then
       BEGIN
       AttribParamsNew(FRM,T.Debit,T.Credit,FInfo.TypeExo) ; ll:=ExecReqINVNew(fbGene,FRM) ;
       If ll<>1 then V_PGI.IoError:=oeSaisie ;
       END  ;
    END ;
// Tiers
for i:=0 to TS.Count-1 do
    BEGIN
    VideTFRM(FRM) ;
    //Fillchar(FRM,SizeOf(FRM),#0) ;
    T:=T_SC(TS[i]) ;
    FRM.Cpt:=T.Cpte ; FRM.Deb:=T.Debit ; FRM.Cre:=T.Credit ; FRM.Axe:=T.Identi ;
    if FRM.Axe='T' then
       BEGIN
       AttribParamsNew(FRM,T.Debit,T.Credit,FInfo.TypeExo) ; ll:=ExecReqINVNew(fbAux,FRM) ;
       If ll<>1 then V_PGI.IoError:=oeSaisie ;
       END ;
    END ;
// Sections
for i:=0 to TS.Count-1 do
    BEGIN
    VideTFRM(FRM) ;
    //Fillchar(FRM,SizeOf(FRM),#0) ;
    T:=T_SC(TS[i]) ;
    FRM.Cpt:=T.Cpte ; FRM.Deb:=T.Debit ; FRM.Cre:=T.Credit ; FRM.Axe:=T.Identi ;
    if ((FRM.Axe<>'G') and (FRM.Axe<>'T')) then
       BEGIN
       AttribParamsNew(FRM,T.Debit,T.Credit,FInfo.TypeExo) ; ll:=ExecReqINVNew(fbSect,FRM) ;
       If ll<>1 then V_PGI.IoError:=oeSaisie ;
       END ;
    END ;
// Journal
Fillchar(FRM,SizeOf(FRM),#0) ;
FRM.Cpt:=SAJAL.JOURNAL ; FRM.Deb:=SAJAL.OldDebit ; FRM.Cre:=SAJAL.OldCredit ;
AttribParamsNew(FRM,SAJAL.OldDebit,SAJAL.OldCredit,FInfo.TypeExo) ;
ll:=ExecReqINVNew(fbJal,FRM) ; If LL<>1 then V_PGI.IoError:=oeSaisie ;
END ;

procedure TFSaisie.TvaTauxDirecteur ( Lig : integer ; Client,ToutDebit : boolean ; Regime : String3 ) ;
Var OMODR   : TMOD ;
    TTC,HT  : Double ;
    CodeTva : String3 ;
    Taux,XX  : Double ;
    i        : integer ;
    O        : TOBM ;
BEGIN
{#TVAENC}
if Not OuiTvaLoc then Exit ;
OMODR:=TMOD(GS.Objects[SA_NumP,Lig]) ; if OMODR=Nil then Exit ;
{Si modifié manuellement ou recalculé prorata --> pas de traitement}
if ((OMODR.ModR.ModifTva) or (OMODR.ModR.TotalAPayerP=0)) then Exit ;
O:=GetO(GS,Lig) ; if O=Nil then Exit ;
TTC:=O.GetMvt('E_CREDIT')-O.GetMvt('E_DEBIT') ; if Not Client then TTC:=-TTC ; 
CodeTva:=VH^.NumCodeBase[1] ; if CodeTva='' then Exit ;
Taux:=Tva2Taux(Regime,CodeTva,Not Client) ; if Taux=-1 then Exit ;
{Prorater le HT, calculé sur le taux directeur, sur les échéances}
HT:=Arrondi(TTC/(1.0+Taux),V_PGI.OkDecV) ;
for i:=1 to MaxEche do
    BEGIN
    FillChar(OMODR.ModR.TabEche[i].TAV,Sizeof(OMODR.ModR.TabEche[i].TAV),#0) ;
    XX:=Arrondi(HT*OMODR.ModR.TabEche[i].MontantP/OMODR.ModR.TotalAPayerP,V_PGI.OkDecV) ;
    if ToutDebit then OMODR.ModR.TabEche[i].TAV[5]:=XX else OMODR.ModR.TabEche[i].TAV[1]:=XX ;
    END ;
{Eviter la ré-initialisation ultérieure, recalcul par prorata ou manuel}
OMODR.ModR.ModifTva:=True ; Revision:=False ;
END ;


{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 19/07/2005
Modifié le ... : 19/07/2005
Description .. : - LG - 19/07/2005 - FB 10054 - on ne saisie plus d'immo sur 
Suite ........ : une ligne crediteur
Mots clefs ... : 
*****************************************************************}
procedure TFSaisie.GereLesImmos ( Lig : integer ) ;
{$IFDEF AMORTISSEMENT}
Var O : TOBM ;
    CGen   : TGGeneral ;
    CAux   : TGTiers ;
    X      : Double ;
    LaLig  : integer ;
    TImmo : TOB ;
{$IFDEF COMPTAAVECSERVANT}
    RecImmo1 : TBaseImmo1 ;
{$ENDIF}
    CodeImmo, LibAux, NumAux : String ;
{$ENDIF}
BEGIN
{$IFDEF AMORTISSEMENT}
//if ((Not VH^.OkModImmo) and (Not V_PGI.VersionDemo)) then Exit ;
{$IFNDEF TT}
if V_PGI.VersionDemo then Exit ;
{$ENDIF}
if Action<>taCreat then Exit ;
// if M.TypeGuide<>'' then Exit ;
if ModeLC then Exit ;
if DEV.Code<>V_PGI.DevisePivot then Exit ;
if SAJAL=Nil then Exit ;
if SAJAL.NatureJal<>'ACH' then
   BEGIN
   if ((SAJAL.NatureJal<>'OD') or (E_NATUREPIECE.Value<>'FF')) then Exit ;
   END ;
if E_NATUREPIECE.Value='AF' then Exit ;
CGen:=GetGGeneral(GS,Lig) ; if CGen=Nil then Exit ; if CGen.NatureGene<>'IMO' then Exit ;
O:=GetO(GS,Lig) ; if O=Nil then Exit ; if O.GetMvt('E_DEBIT') = 0 then exit ;
if O.GetMvt('E_NUMEROIMMO')=666 then Exit else O.PutMvt('E_NUMEROIMMO',666) ;
if O.GetMvt('E_IMMO')<>'' then Exit ;
if HPiece.Execute(46,Caption,'')<>mrYes then Exit ;
{$IFDEF AMORTISSEMENT}
X:=O.GetMvt('E_DEBIT')+O.GetMvt('E_CREDIT') ;
LaLig:=TrouveLigTiers(GS,0) ; LibAux:='' ; NumAux:='' ;
if LaLig>0 then BEGIN CAux:=GetGTiers(GS,LaLig) ; if CAux<>Nil then BEGIN LibAux:=CAux.Libelle ; NumAux:=CAux.Auxi ; END ; END ;
TImmo := TOB.Create('IMMO',nil,-1);
TImmo.PutValue('I_DATEAMORT',DatePiece);
TImmo.PutValue('I_DATEPIECEA',DatePiece);
TImmo.PutValue('I_COMPTEIMMO',CGen.General);
TImmo.PutValue('I_MONTANTHT', X);
TImmo.PutValue('I_TVARECUPERABLE',HT2TVA(X,GeneRegTVA,GeneSoumisTPF,O.GetMvt('E_TVA'),O.GetMvt('E_TPF'),True,DECDEV));
TImmo.PutValue('I_TIERSA',LibAux);
TImmo.PutValue('I_REFINTERNEA',GS.Cells[SA_RefI,Lig]);
TImmo.PutValue('I_VALEURACHAT',X);
TImmo.PutValue('I_QUANTITE',O.GetMvt('E_QTE1'));
if (TImmo.GetValue('I_QUANTITE')=0) then TImmo.PutValue('I_QUANTITE',1);
TImmo.PutValue('I_ETABLISSEMENT',E_ETABLISSEMENT.Value);
TImmo.PutValue('I_LIBELLE',GS.Cells[SA_Lib,Lig] );
{$ENDIF}
{$IFDEF COMPTAAVECSERVANT}
IF VHIM^.ServantPGI Then
  BEGIN
  RecImmo1.CompteImmo:=TImmo.GetValue('I_COMPTEIMMO') ; RecImmo1.Fournisseur:=NumAux ;
  RecImmo1.Reference:=TImmo.GetValue('I_REFINTERNEA') ;
  RecImmo1.Libelle:=TImmo.GetValue('I_LIBELLE') ;
  RecImmo1.CodeEtab:=TImmo.GetValue('I_ETABLISSEMENT') ;
  RecImmo1.DateAchat:=TImmo.GetValue('I_DATEPIECEA') ;
  RecImmo1.MontantHT:=TImmo.GetValue('I_MONTANTHT') ;
  RecImmo1.MontantTVA:=TImmo.GetValue('I_TVARECUPERABLE') ;
  RecImmo1.Quantite:=TImmo.GetValue('I_QUANTITE') ;
  Codeimmo:=SaisieExterne(taCreat,0,RecImmo1) ;
  END Else
  if VH^.OkModImmo then Codeimmo:=AMLanceFiche_FicheImmobilisationEnSaisie ( TImmo);
  TImmo.Free;
{$ELSE}
{$IFDEF AMORTISSEMENT}
  if VH^.OkModImmo then CodeImmo:=AMLanceFiche_FicheImmobilisationEnSaisie ( TImmo);
  TImmo.Free;
{$ENDIF}
{$ENDIF}
O.PutMvt('E_IMMO',CodeImmo) ;
{$ENDIF}
END ;

procedure TFSaisie.GereTvaEncNonFact ( Lig : integer ) ;
Var CGen : TGGeneral ;
    CAux : TGTiers ;
    Client : boolean ;
    RegEnc,StEnc,Regime : String3 ;
    Exi          : TExigeTva ;
BEGIN
{#TVAENC}
if PasModif then Exit ;
if Not OuiTvaLoc then Exit ;
{Si facture --> géré ailleurs (Calculdébitcrédit)}
if SorteTva<>stvDivers then Exit ;
if Not isTiers(GS,Lig) then Exit ;
{Client:=False ;} Regime:='' ;
if GS.Cells[SA_Aux,Lig]<>'' then
   BEGIN
   CAux:=GetGTiers(GS,Lig) ; if CAux=Nil then Exit ;
   if ((CAux.NatureAux='CLI') or (CAux.NatureAux='AUD')) then Client:=True else
    if ((CAux.NatureAux='FOU') or (CAux.NatureAux='AUC')) then Client:=False else Exit ;
   StEnc:=CAux.Tva_Encaissement ;
   if GeneRegTVA<>'' then Regime:=GeneRegTVA else Regime:=CAux.RegimeTva ;
   END else if GS.Cells[SA_Gen,Lig]<>'' then
   BEGIN
   CGen:=GetGGeneral(GS,Lig) ; if CGen=Nil then Exit ;
   if CGen.NatureGene='TID' then Client:=True else
    if CGen.NatureGene='TIC' then Client:=False else Exit ;
   StEnc:=CGen.Tva_Encaissement ;
   if GeneRegTVA<>'' then Regime:=GeneRegTVA else Regime:=CGen.RegimeTva ;
   END else Exit ;
CGen:=GetGGeneral(GS,Lig) ; if CGen=Nil then Exit ;
if Not EstCollFact(CGen.General) then Exit ;
{Si client --> param soc, sinon exige du fourn. Si choix explicite, forcer. Débit --> Aucun traitement}
if Client then RegEnc:=VH^.TvaEncSociete else RegEnc:=StEnc ;
if ChoixExige then Exi:=TExigeTva(BPopTva.ItemIndex-1) else Exi:=Code2Exige(RegEnc) ;
Case Exi of
   tvaMixte  : Exit ;
   tvaEncais : TvaTauxDirecteur(Lig,Client,False,Regime) ;
   tvaDebit  : TvaTauxDirecteur(Lig,Client,True,Regime) ;
   END ;
END ;

procedure TFSaisie.AfficheExigeTva ;
Var Nat : String3 ;
BEGIN
if (VH^.PaysLocalisation=CodeISOES) or (Not OuiTvaLoc) then BEGIN BPopTva.Visible:=False ; Exit ; END ; //XVI 24/02/2005
Nat:=E_NATUREPIECE.Value ;
if ((Nat<>'FC') and (Nat<>'AC') and (Nat<>'OC') and
    (Nat<>'FF') and (Nat<>'AF') and (Nat<>'OF')) then BEGIN BPopTva.Visible:=False ; Exit ; END ;
if (((Nat='OC') or (Nat='FC') or (Nat='AC')) and (VH^.TvaEncSociete='TD')) then BEGIN BPopTva.Visible:=False ; Exit ; END ;
if (((Nat='FC') or (Nat='AC')) and (SorteTva<>stvVente)) then BEGIN BPopTva.Visible:=False ; Exit ; END ;
if (((Nat='FF') or (Nat='AF')) and (SorteTva<>stvAchat)) then BEGIN BPopTva.Visible:=False ; Exit ; END ;
BPopTva.Visible:=True ;
END ;

procedure TFSaisie.GetSorteTva ;
BEGIN
SorteTva:=stvDivers ; if SAJAL=Nil then Exit ;
if ((SAJAL.NatureJal='VTE') and ((E_NATUREPIECE.Value='FC') or (E_NATUREPIECE.Value='AC'))) then SorteTva:=stvVente else
 if ((SAJAL.NatureJal='ACH') and ((E_NATUREPIECE.Value='FF') or (E_NATUREPIECE.Value='AF'))) then SorteTva:=stvAchat ;
END ;

procedure TFSaisie.RenseigneRegime ( Lig : integer ; Recharge : boolean ) ;
Var CGen : TGGeneral ;
    CAux : TGTiers ;
    TlT : integer;
BEGIN
if Lig<=0 then Exit ;
if GS.Cells[SA_Aux,Lig]<>'' then
   BEGIN
   // 17/11/2006 FQ18787 YMO On applique le régime du PREMIER AUXILIAIRE pour alimenter E_ECHEENC1,
   // quelque soient les taux correpondant aux autres tiers dans la pièce
   TlT :=  TrouveLigTiers(GS,0);
   if TlT >0 then
       CAux:=GetGTiers(GS,TlT)
   else CAux:=GetGTiers(GS,Lig) ;
   if CAux<>Nil then
      BEGIN
      if ((GeneRegTVA='') or (Not RegimeForce) or (Recharge)) then GeneRegTVA:=CAux.RegimeTVA ;
      if ((Action=taCreat) and (GS.RowCount<=3) and (Not EstRempli(GS,GS.RowCount-1))) then GeneRegTVA:=CAux.RegimeTVA ;
      GeneSoumisTPF:=CAux.SoumisTPF ;
      END ;

   END else
   BEGIN
   CGen:=GetGGeneral(GS,Lig) ;
   if CGen<>Nil then
      BEGIN
      if ((GeneRegTVA='') or (Not RegimeForce) or (Recharge)) then GeneRegTVA:=CGen.RegimeTVA ;
{$IFDEF COMPTA}
      if VH^.PaysLocalisation=CodeISOES then
         GeneSoumisTPF:=IsTiersSoumisTPF(GS)
      else
         GeneSoumisTPF:=CGen.SoumisTPF ;
{$ENDIF}
      END ;
   END ;
END ;

procedure TFSaisie.AttribRegimeEtTVA ;
Var ia : integer ;
    CGen : TGGeneral ;
    O : TOBM ;
BEGIN
ia:=TrouveLigTiers(GS,0) ;
if ia>0 then RenseigneRegime(ia,False) ;
ia:=TrouveLigHT(GS,0,False) ; if ia<=0 then exit ;
if GS.Cells[SA_Gen,ia]<>'' then
   BEGIN
   CGen:=GetGGeneral(GS,ia) ;
   if CGen<>Nil then
      BEGIN
      // FQ 21143 SBO 06/08/2007 : si TVA renseigné dans la ligne, il faut la récupérer
      O := GetO( GS, ia ) ;
      if Assigned( O ) and (O.GetString('E_TVA') <> '' )
        then GeneTVA := O.GetString('E_TVA')
        else GeneTVA := CGen.TVA ;
{$IFDEF COMPTA}
      if not IsTiersSoumisTPF(GS) then
         GeneTPF:=''
      else
         GeneTPF:=CGen.TPF ;
{$ENDIF}
      END ;
   END ;
END ;

procedure TFSaisie.DetruitAncien ;
Var iecr       : integer ;
    StW,StLig  : String ;
    X          : TDELMODIF ;
    DateUnique : TDateTime ;
    MonoDate   : boolean ;
BEGIN
if PasModif then Exit ;
MonoDate:=True ; DateUnique:=0 ;
if Not Revision then
   BEGIN
   {Destruction Ecritures}
   StW:=WhereEcriture(tsGene,M,False) ;
   for iecr:=1 to NbLigEcr do
       BEGIN
       X:=TDELMODIF(TDELECR[iecr-1]) ;
       if iEcr=1 then Dateunique:=X.DateModification else
          if X.DateModification<>DateUnique then BEGIN MonoDate:=False ; Break ; END ;
       END ;
   if ((MonoDate) and (DateUnique>0)) then
      BEGIN
      StLig:=' AND E_DATEMODIF="'+UsTime(DateUnique)+'"' ;
      if ExecuteSQL('Delete from Ecriture where '+StW+StLig)<>NbLigEcr then BEGIN V_PGI.IOError:=oeSaisie ; Exit ; END ;
      END else
      BEGIN
      for iecr:=1 to NbLigEcr do
          BEGIN
          X:=TDELMODIF(TDELECR[iecr-1]) ;
          StLig:=' AND E_NUMLIGNE='+IntToStr(X.NumLigne)+' AND E_NUMECHE='+InttoStr(X.NumEcheVent)+' AND E_DATEMODIF="'+UsTime(X.DateModification)+'"' ;
          if ExecuteSQL('Delete from Ecriture where '+StW+StLig)<=0 then BEGIN V_PGI.IOError:=oeSaisie ; Exit ; END ;
          END ;
      END ;
   END ;
{Destruction Analytiques}
StW:=WhereEcriture(tsAnal,M,False) ;
MonoDate:=True ; DateUnique:=0 ;
for iecr:=1 to NbLigAna do
    BEGIN
    X:=TDELMODIF(TDELANA[iecr-1]) ;
    if iEcr=1 then Dateunique:=X.DateModification else
       if X.DateModification<>DateUnique then BEGIN MonoDate:=False ; Break ; END ;
    END ;
if ((MonoDate) and (DateUnique>0)) then
   BEGIN
   StLig:=' AND Y_DATEMODIF="'+UsTime(DateUnique)+'"' ;
   if ExecuteSQL('Delete from Analytiq where '+StW)<>NbLigAna then BEGIN V_PGI.IOError:=oeSaisie ; Exit ; END ;
   END else
   BEGIN
   for iecr:=1 to NbLigAna do
       BEGIN
       X:=TDELMODIF(TDELANA[iecr-1]) ;
       StLig:=' AND Y_NUMLIGNE='+IntToStr(X.NumLigne)+' AND Y_AXE="'+X.Ax+'"'
             +' AND Y_NUMVENTIL='+InttoStr(X.NumEcheVent)+' AND Y_DATEMODIF="'+UsTime(X.DateModification)+'"' ;
       if ExecuteSQL('Delete from Analytiq where '+StW+StLig)<=0 then BEGIN V_PGI.IOError:=oeSaisie ; Exit ; END ;
       END ;
   END ;
END ;

procedure TFSaisie.EnvoiCFONB ;
Var RR    : RMVT ;
    Q     : TQuery ;
    SQL, VerifAuxi   : String ;
    O     : TOBM ;
    TEche : TList ;
    OkEnc : boolean ;
BEGIN
if Not M.ExportCFONB then Exit ;
if TPIECE.Count<=0 then Exit ;
if Not SuiviMP then
   BEGIN
   if ((M.TypeGuide<>'ENC') and (M.TypeGuide<>'DEC')) then Exit ;
   OkEnc:=(M.TypeGuide='ENC') ;
   END else
   BEGIN
   OkEnc:=isEncMP(M.smp) ;
   END ;
if HPiece.Execute(35,Caption,'')<>mrYes then Exit ;
RR:=P_MV(TPIECE[TPIECE.Count-1]).R ;
TEche:=TList.Create ;
{YMO 01/02/2006 FQ 17308, JP 10/05/07 pour le not null pour Oracle}
If Not M.TIDTIC then VerifAuxi := ' AND (E_AUXILIAIRE <> "" AND E_AUXILIAIRE IS NOT NULL)';

// SBO 05/07/2007 : gestion pb DB2 avec champ blocnote à positionner à la fin
SQL:='SELECT ' + GetSelectAll('E', True) + ', E_BLOCNOTE FROM ECRITURE WHERE '+WhereEcriture(tsGene,RR,False)+VerifAuxi+' AND E_ECHE="X" AND E_NUMECHE>0 AND E_CFONBOK="#"';  // FQ 12393
Q:=OpenSQL(SQL,True) ;
While Not Q.EOF do
   BEGIN
   O:=TOBM.Create(EcrGen,'',False) ; O.ChargeMvt(Q) ; TEche.Add(O) ;
   Q.Next ;
   END ;
Ferme(Q) ;
{$IFNDEF IMP}
{$IFNDEF GCGC}
{$IFNDEF CCS3}
{$IFNDEF VEGA}
If Not VH^.OldTeleTrans Then M.EnvoiTrans:='' ;
if TEche.Count>0 then begin

  {JP 12/08/05 : FQ 12339 : en suivi de tiers, il faut aussi renseigné le modèle de bordereau}
  if szModeleBordereau = '' then szModeleBordereau := M.Document;
  ExportCFONB(OkEnc,M.General,M.FormatCFONB,M.EnvoiTrans,TECHE,M.smp, szModeleBordereau);

  {VL 30102003 FQ 12958, JP 10/05/07 pour le VerifAuxi}
  SQL:='Update Ecriture Set E_CFONBOK="X" where '+WhereEcriture(tsGene,RR,False)+ VerifAuxi + ' AND E_ECHE="X" AND E_NUMECHE>0 AND E_CFONBOK="#"' ;
  ExecuteSQL(SQL) ;
  {JP 05/08/05 : FQ 14853 : le bordereau est géré dans CFONB.EditeCFONB ; en passant à False, cela évite qu'il
                 repose la questionTFSaisie.EnvoiBordereau}
  M.Bordereau := False;
  // FIN VL
end;
{$ENDIF}
{$ENDIF}
{$ENDIF}
{$ENDIF}
VideListe(TEche) ; TEche.Free ;
END ;

procedure TFSaisie.EnvoiBordereau ;
Var RR    : RMVT ;
    Q     : TQuery ;
    SQL,SWhere   : String ;
    O     : TOBM ;
    TEche : TList ;
    VerifAuxi : string; {JP 10/05/07}
BEGIN
if Not M.Bordereau then Exit ;
if M.Document='' then Exit ;
if TPIECE.Count<=0 then Exit ;
if Not SuiviMP then
   BEGIN
   if ((M.TypeGuide<>'ENC') and (M.TypeGuide<>'DEC')) then Exit ;
   END ;
if HPiece.Execute(37,Caption,'')<>mrYes then Exit ;
RR:=P_MV(TPIECE[TPIECE.Count-1]).R ;
TEche:=TList.Create ;
{JP 10/05/07 : Gestion des TIC / TID}
if not M.TIDTIC then VerifAuxi := ' AND (E_AUXILIAIRE <> "" AND E_AUXILIAIRE IS NOT NULL)';

// SBO 05/07/2007 : gestion pb DB2 avec champ blocnote à positionner à la fin
SQL:='SELECT ' + GetSelectAll('E', True) + ', E_BLOCNOTE FROM ECRITURE WHERE '+WhereEcriture(tsGene,RR,False)+ VerifAuxi + ' AND E_ECHE="X"  AND E_NUMECHE>0 AND E_CFONBOK="#"';
Q:=OpenSQL(SQL,True) ;
if Not Q.EOF then
   BEGIN
   O:=TOBM.Create(EcrGen,'',False) ; O.ChargeMvt(Q) ; TEche.Add(O) ;
   END ;
Ferme(Q) ;
if TEche.Count>0 then
   BEGIN
   O:=TOBM(TECHE[0]) ; RR:=OBMToIdent(O,False) ; SWhere:=WhereEcriture(tsGene,RR,False) ;
{$IFNDEF IMP}
{$IFNDEF GCGC}
   LanceEtat('E','BOR',M.Document,True,False,False,Nil,SWhere,'',False) ;
{$ENDIF}
{$ENDIF}
   {JP 05/08/05 : FQ 14853 : si dans l'on est dans le cas d'éditions seules, on met E_CFONBOK = "-"
    JP 10/05/07 : Ajout du VerifAuxi pour la gestion des TIC ou TID}
   if (M.Smp = smpDecVirEdt) or (M.Smp = smpDecVirInEdt) or
      {JP 19/08/05 : FQ 10094 : Gestion des lettres-chèques}
      ((M.Smp = smpAucun) and (M.SorteLettre = tslCheque)) then
     SQL:='UPDATE ECRITURE SET E_CFONBOK = "-" WHERE ' + WhereEcriture(tsGene,RR,False) + VerifAuxi + ' AND E_ECHE="X" AND E_NUMECHE>0 AND E_CFONBOK = "#"'
   else
     SQL:='UPDATE ECRITURE SET E_CFONBOK = "X" WHERE ' + WhereEcriture(tsGene,RR,False) + VerifAuxi + ' AND E_ECHE="X" AND E_NUMECHE>0 AND E_CFONBOK = "#"';
   ExecuteSQL(SQL) ;
   END ;
VideListe(TEche) ; TEche.Free ;
END ;

procedure TFSaisie.LettrerReglement ;
Var RR  : RMVT ;
BEGIN
RR:=P_MV(TPIECE[TPIECE.Count-1]).R ;
// Modif Fiche 12017
  LettrageEncaDeca(TECHEORIG,RR,DEV,TOBNumChq) ;  // ajout tobNumChq SBO maj n°chq depuis CCMP
  FreeAndNil(TOBNumChq) ;                         // libération TobNumChq
// Fin Modif Fiche 12017
END ;

procedure TFSaisie.TraiteMP ;
Var i    : integer ;
    j    : integer ;    // FQ 15127 SBO 04/01/2004
    k    : integer ;    // FQ 15127 SBO 04/01/2004
    O    : TOBM ;
    lTob : TOB ;        // FQ 15127 SBO 04/01/2004
BEGIN
M.Simul:='N' ;
for i:=1 to GS.RowCount-1 do
    BEGIN
    O:=GetO(GS,i) ; if O=Nil then Break ;
    O.PutMvt('E_QUALIFPIECE','N') ;
    O.PutMvt('E_NUMENCADECA',M.NumEncaDeca); // VL 17122003 FFF
    // FQ 15127 SBO 04/01/2004 Analytique pas mis à jour...
    for j := 0 to O.Detail.Count - 1 do // Pour tous les axes
      begin
      lTob := O.Detail[ j ] ;
      for k := 0 to lTob.Detail.Count - 1 do // Pour toutes les ventilations
        lTob.Detail[ k ].PutValue('Y_QUALIFPIECE', 'N') ;
      end ; // Fin FQ 15127
    END ;
END ;

procedure TFSaisie.ValideLaPiece ;
BEGIN
if Not GS.SynEnabled then Exit ;
GridEna(GS,False) ;
if ((OuiTvaLoc) and (Not Revision)) then CalculDebitCredit ;
if Action=taModif then
  BEGIN
  DetruitAncien ;
  END ;
if SuiviMP then TraiteMP ;
if V_PGI.IOError=oeOK then ValideLignesGene ;
END ;

procedure TFSaisie.ValideLeReste ;
BEGIN
V_PGI.IOError:=oeOK ;
if (Action=taModif) and (not SUIVIMP) then InverseSoldeNew ;
if V_PGI.IOError=oeOK then ValideLesComptes ;
if V_PGI.IOError=oeOK then
  BEGIN
  ValideLeJournalNew ;
  END ;
GridEna(GS,True) ;
END ;

procedure TFSaisie.ValideLesAno ;
var
 i,j     : integer ;
 OBM     : TOBM ;
 lIndex  : integer ;
 lStNat  : string ;
 Gen     : string ;
 MD      : TMOD ;
 R       : T_ModeRegl ;
 CGen    : TGGeneral ;

 procedure _AjouteAno ;
 begin
  Gen	:= OBM.GetMvt('E_GENERAL') ;
  lIndex:=FInfo.Compte.GetCompte(Gen) ;
  if lIndex>-1 then
   lStNat:= FInfo.Compte_GetValue('G_NATUREGENE') ;
  AjouteAno(TSA,OBM,lStNat,false) ;
 end ;


begin
V_PGI.IOError:=oeOK ;
if PasModif then Exit ;
if M.Simul<>'N' then Exit ;

for i:=1 to GS.RowCount-2 do
 begin
  MD := TMOD(GS.Objects[SA_NumP,i]) ;
  CGen := GetGGeneral(GS,i) ;
  if CGen=Nil then Exit ;
  if ((MD=Nil) or (Lettrable(CGen)=MonoEche)) then
   begin
    // Récupération et remplissage TOB Ecritures générales
    RecupTronc(i,MD,OBM);
    EcartChange(OBM);
    _AjouteAno ;
   end
    else
     begin
      R := MD.MODR ;
      for j:=1 to R.NbEche do
       begin
        // Récupération et remplissage TOB Ecritures générales
        RecupTronc(i,MD,OBM);
        RecupEche(i,R,j,OBM);
        EcartChange(OBM);
        _AjouteAno ;
       end ;
    end ;
 end ; // for

 if not ExecReqMAJAno(TSA) then
  begin
   V_PGI.IoError := oeSaisie ;
   exit ;
  end ;

GridEna(GS,True) ;
end ;

{==========================================================================================}
{================================ Barre d'outils ==========================================}
{==========================================================================================}
procedure TFSaisie.ForcerMode ( Cons : boolean ; Key : Word ) ;
BEGIN
if ((Cons) and (Modeforce=tmDevise)) then Exit ;
if ((Not Cons) and (Modeforce=tmNormal)) then Exit ;
if Cons then
   BEGIN
   AffecteGrid(GS,taConsult) ; ModeForce:=tmDevise ; GS.SetFocus ;
   PEntete.Enabled:=False ; Outils.Enabled:=False ; Valide97.Enabled:=False ;
   PForce.Align:=alClient ; PForce.Parent:=PPied ; PForce.Visible:=True ;
   AfficheLeSolde(HForce,SI_TotDP,SI_TotCP) ;
   END else
   BEGIN
   if Key=VK_RETURN then SoldeLaligne(GS.Row) ;
   PEntete.Enabled:=True ; Outils.Enabled:=True ;
   if Outils.CanFocus then Outils.SetFocus ; Valide97.Enabled:=True ;
   GS.CacheEdit ; AffecteGrid(GS,Action) ;
   GS.Col:=SA_Gen ; if GS.Row=1 then GS.Row:=2 else GS.Row:=1 ;
   GS.SetFocus ; GS.MontreEdit ;
   ModeForce:=tmNormal ; PForce.Visible:=False ;
   END ;
END ;

procedure TFSaisie.NormaliserEscomptes ;
Var i : integer ;
    TOBL : TOB ;
    YY   : RMVT ;
    T    : HTStrings ;
    St   : String ;
BEGIN
if Not SuiviMP then Exit ;
if TOBMPEsc=Nil then Exit ;
if TOBMPEsc.Detail.Count<=0 then Exit ;
for i:=0 to TOBMPEsc.Detail.Count-1 do
    BEGIN
    TOBL:=TOBMPEsc.Detail[i] ;
{$IFNDEF VEGA}
{$IFNDEF CMPGIS35}
    if Not NormaliserPieceSimuTOB(TOBL,True) then BEGIN V_PGI.IoError:=oeUnknown ; Break ; END ;
{$ENDIF}
{$ENDIF}
    T:=HTStringList.Create ;
    YY:=TOBToIdent(TOBL,True) ; St:=EncodeLC(YY) ; T.Add(St) ;
    TECHEORIG.Add(T) ;
    END ;
END ;

procedure TFSaisie.SuiteEncaDeca ;
BEGIN
GS.VidePile(True) ;

if SuiviMP then
begin
 if Transactions(NormaliserEscomptes,1)<>oeOk then MessageAlerte(HDiv.Mess[25]) ;
 //YMO 12/2005 Lettrage facultatif pour les ecritures en devise
 {YMO 16/06/2006 FQ18245 Si l'on est dans 'Encaissements divers', il n'y a pas
 de document généré, il n'est donc pas nécessaire de mettre à jour E_NUMTRAITECHQ
 dans la pièce d'origine, d'où le test sur m.Sortelettre}
 if Sanslettrage and (m.SorteLettre<>tslAucun) then MajMPSansLettrage; // Maj diverses sans lettrage
end;

if (not Sanslettrage) and (Transactions(LettrerReglement,1)<>oeOk) then MessageAlerte(HDiv.Mess[19]) ;

EnvoiCFONB ;
EnvoiBordereau ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 25/09/2002
Modifié le ... : 29/09/2003
Description .. : - LG - 25/09/2002 - l'appel de MajIOCom est conditionne
Suite ........ : par par le test OkSynchro
Suite ........ : - LG - 29/09/2003 - FB 12694 - l'affectation de e_refrevision
Mots clefs ... :
*****************************************************************}
procedure TFSaisie.ClickValide ;
Var StE : String3 ;
    io,io1  : TIOErr ;
//    TesterLig : boolean ;
    i : integer; //SG6 21/12/2004 FQ 14731
    lTobEcr : Tob ;
    Cancel : boolean ;
    Chg    : boolean ;
    ACol   : integer ;
    ARow   : integer ;
BEGIN
if Not GS.SynEnabled then Exit ;
if ((M.TypeGuide='NOR') and (M.FromGuide)) then BEGIN ClickAbandon ; Exit ; END ;
If Guide And (M.TypeGuide<>'DEC') and (M.TypeGuide<>'ENC') (* And FlashGuide.Visible *) Then // FQ 21700 SBO 22/10/2007 test fin guide sur flashing label ! (du coup, possible de valide quand il était invisible...)
  BEGIN
  HDiv.Execute(26,Caption,'') ; Exit ;
  END ;
if vh^.CPIFDEFCEGID then // modif ok
//if (appeldatepourguide And (M.TypeGuide='NOR') and (M.SaisieGuidee)) then
If MethodeGuideDate1 Then
  BEGIN
  if ( appeldatepourguide And (Not M.FromGuide) And (((M.TypeGuide='NOR') and (M.SaisieGuidee)) Or (CodeGuideFromSaisie<>''))) then
    BEGIN
    GS.Enabled:=TRUE ;
    if ((GS.RowCount<=2) and (GS.CanFocus) and
        (Screen.ActiveControl<>GS) and (Screen.ActiveControl<>Valide97)) then BEGIN GS.SetFocus ; Exit ; END ;
    END ;
  END ;

if ((E_JOURNAL.Value<>'') and (GS.RowCount<=2) and (GS.CanFocus) and
    (Screen.ActiveControl<>GS) and (Screen.ActiveControl<>Valide97)) then BEGIN GS.SetFocus ; Exit ; END ;
{Contrôles avant validation}
if Action=taConsult then BEGIN Fermeture() ; Exit ; END ;
if ((Not BValide.Enabled) and (CurM<=0)) then exit ;
if not (ctxPCL in V_PGI.PGIContexte) then // FQ 15571 : SBO 20/04/2005 Anciennement vh^.CPIFDEFCEGID
 TraiteLigneAZero ;

// === Validation de la ligne en cours de saisie ===  // FQ 17564
  // Gestion des evènements : Sortie de cellule
  ACol   := GS.Col ;
  ARow   := GS.Row ;
  Cancel := False  ;
  GSCellExit( GS, ACol, ARow, Cancel ) ;
  // Gestion des evènements : Sortie de ligne
  if not Cancel then
    begin
    Chg    := False ;
    GSRowExit( GS, GS.Row, Cancel, Chg ) ;
    if Cancel then exit ;
    end
  else Exit ;
{
ValideCeQuePeux(GS.Row) ;
TesterLig:=((EstRempli(GS,GS.Row)) or (Guide)) ;
if ((TesterLig) and (Not LigneCorrecte(GS.Row,False,(Action=taCreat) or (Not Revision) or (UnChange)))) then Exit ;
if ((TesterLig) and (DEV.Code<>V_PGI.DevisePivot) and (Action<>taConsult) and (Not Revision) and (EstRempli(GS,GS.Row)))
   then CalculDebitCredit ;
}
// === Fin de Validation de la ligne en cours de saisie === // Fin FQ 17564

if not PieceCorrecte then Exit ;
if not EquilEuroValide then Exit ;
if Not ControleRIB then Exit ;
MajIoCom(GS) ;
{$IFNDEF FFF}
if ((Action=taCreat) and (H_NUMEROPIECE_.Visible) and (GS.RowCount>2)) or (SuiviMP) then AttribNumeroDef ;
{$ENDIF FFF}
if GS.RowCount<>NbLOrig then Revision:=False ;
GS.Row:=1 ; GS.Col:=SA_Gen ; GS.SetFocus ; BValide.Enabled:=False ;
AttribRegimeEtTVA ;
{Scenario avant validation}
if ((OkScenario) and (Scenario.GetMvt('SC_CONTROLETVA')='OUI')) then if Not ControleTVAOK then Exit ;
{Validation}
DatePiece:=StrToDate(E_DATECOMPTABLE.Text) ;
SetTypeExo ;   // FQ 13246 : SBO 30/03/2005
NowFutur:=NowH ; GS.Enabled:=False ;

// Essai Pb FFF validation de pièce SBO 29/16/2006
{$IFDEF FFF}
if ((Action=taCreat) and (H_NUMEROPIECE_.Visible) and (GS.RowCount>2)) or (SuiviMP) then
   if Transactions(AttribNumeroDef,1)<>oeOK then
     begin
     MessageAlerte(HDiv.Mess[5]) ;
     Exit ;
     end ;
{$ENDIF FFF}
// Fin Essai Pb FFF validation de pièce SBO 29/16/2006

io:=Transactions(ValideLaPiece,5) ;
If io=oeOK then
 begin
  io1:=Transactions(ValideLeReste,10) ;
  If io1<>oeOK Then MessageAlerte(HDiv.Mess[22]+HDiv.Mess[23])
   else
    Transactions(ValideLesAno,1) ;
 end ;
GridEna(GS,True) ;
GS.Enabled:=True ;
Case io of
    oeOK : BEGIN
           StockeLaPiece ;
           if SuiviMP then
              BEGIN
              SuiteEncaDeca ; CEstGood:=True ;
              END else
              BEGIN
              if (Action=taCreat) and ((M.TypeGuide='ENC') or (M.TypeGuide='DEC')) then
                 BEGIN
                 {Encaissements/Décaissements}
                 SuiteEncaDeca ; CEstGood:=True ;
                 END else
                 BEGIN
                 {Scénario après validation ou Tiers payeur}
                 if ExisteTP then
                    BEGIN
                    if Transactions(GereTiersPayeur,2)<>oeOk then MessageAlerte(HDiv.Mess[17]) ;
                    END else
                    BEGIN
                    if Transactions(DetruitOldPayeurs,2)<>oeOk then MessageAlerte(HDiv.Mess[24]) ;
                    if OkScenario then BEGIN if ((Scenario.GetMvt('SC_LETTRAGESAISIE')='X') and (Not SuiviMP)) then LettrageEnSaisie ; END ;
                    END ;
                 END ;
              END ;
           if ((Action=taCreat) and (VH^.CPExoRef.Code<>'')) then
             if ((DatePiece>=VH^.CPExoRef.Deb) and (DatePiece<=VH^.CPExoRef.Fin)) then
                if ctxPCL in V_PGI.PGIContexte then
                BEGIN
                VH^.CPLastSaisie:=DatePiece ;
                {$IFNDEF SPEC350}
                SetParamSoc('SO_CPLASTSAISIE',VH^.CPLastSaisie) ;
                {$ENDIF}
                END ;
           if (SAJAL.NATUREJAL = 'ANO') then
             begin
             StE:=QuelExoDT(StrToDate(E_DATECOMPTABLE.Text)) ;
             SetParamSoc ('SO_EXOV8',StE);
             if StE = VH^.Encours.Code then
              begin
              VH^.ExoV8.Code := VH^.Encours.Code;
              VH^.ExoV8.Deb := VH^.Encours.Deb;
              VH^.ExoV8.Fin := VH^.Encours.Fin;
              end ;
             end;
           END ;
      else BEGIN
            if ((Action=taCreat) and (H_NUMEROPIECE_.Visible=False) and (PossibleRecupNum)) then
               if Transactions(DecrementeNumero,10)<>oeOk then MessageAlerte(HDiv.Mess[6]) ;
           if io=oeSaisie then
              BEGIN
              MessageAlerte(HDiv.Mess[4]) ;
              END else
              BEGIN
              if io=oeUnknown then MessageAlerte(HDiv.Mess[3]) ;
              END ;
           CEstGood:=False ;
           END ;
   END ;
{Ré-initialisation}
//SG6 21/12/2004 FQ 14731 Vérification quantité saisie
if TestQteAna then // DEV3946
  for i:=1 to GS.RowCount-1 do
    begin
    lTobEcr := Tob(GetO(GS,i)) ;
    if ( lTobEcr <> nil ) and ( ( lTobEcr.GetNumChamp('CHECKQTE') < 0 ) or
                                ( lTobEcr.GetString('CHECKQTE')<>'X'  ) ) then
      CheckQuantite( lTobEcr );
    end ;


PieceModifiee:=False ; ModeLC:=False ;
if ((Action=taCreat) and (M.LeGuide='') and (Not M.ANouveau)) then
   BEGIN
   if VH^.BouclerSaisieCreat then ReInitPiece(False) else
      BEGIN
      Fermeture() ;
      END ;
   END else
 if ((Action=taCreat) and (M.LeGuide='') and (M.ANouveau) and (IsInside(Self))) then ReInitPiece(False) else
   BEGIN
   if ((Action=taCreat) and (LesM<>Nil) and (CEstGood)) then
      BEGIN
      if LesM.Count>1 then
         BEGIN
         Inc(CurM) ;
         if LesM.Count-1>=CurM then
            BEGIN
            CloseSaisie ;
            CreateSaisie ; M:=P_MV(LesM[CurM]).R ;
            ShowSaisie ; BValide.Enabled:=True ;
            END else Fermeture();
         END else Fermeture();
      END else If M.MSED.MultiSessionEncaDeca Then Fermeture()
            Else if Not ModifNext then Fermeture();
   END ;
END ;

function TFSaisie.EquilEuroValide : Boolean;
Var EcartP: Double ;
BEGIN
Result:=True ;
//if Not EuroOK then Exit ;
if PasModif then Exit ;
{$IFNDEF GCGC}
EcartP:=Arrondi(SI_TotDP-SI_TotCP,V_PGI.OkDecV) ;
if (EcartP = 0) then Exit ;
if DatePiece<V_PGI.DateDebutEuro then BEGIN SoldeLaLigne(1) ; Exit ; END ;
if DEV.Code<>V_PGI.DevisePivot then BEGIN SoldeLaLigne(1) ; Exit ; END ;
{$ENDIF}
END ;

procedure TFSaisie.ClickVentil ;
Var OkL : Boolean ;
BEGIN
if GS.Row>=GS.RowCount-1 then Exit ;
if Not EstRempli(GS,GS.Row) then Exit ;
OkL:=LigneCorrecte(GS.Row,False,True) ;
if OkL then GereAnal(GS.Row,False,False) ;
END ;

procedure TFSaisie.ClickEche ;
Var OkL : Boolean ;
BEGIN
if GS.Row>=GS.RowCount-1 then Exit ;
if Not EstRempli(GS,GS.Row) then Exit ;
OkL:=LigneCorrecte(GS.Row,False,True) ;
if OkL then BEGIN GereEche(GS.Row,False,False) ; Revision:=False ; END ;
END ;

procedure TFSaisie.ClickDel ;
BEGIN
if PasModif then Exit ;
if ModeLC then Exit ;
if GS.RowCount<=3 then Exit ;
if ((GS.Row=GS.RowCount-1) and (Not EstRempli(GS,GS.Row))) then BEGIN Gs.SetFocus ; Exit ; END ;
if PasToucheLigne(GS.Row) then Exit ;
GS.SynEnabled:=False ;
if Guide then DetruitLigneGuide(GS.Row);
DetruitLigne(GS.Row,True) ; Revision:=False ;
ControleLignes ;
GereOptionsGrid(GS.Row) ;
// Ajout SBO 22/07/2004 ( FQ 10233 )
if Guide and ( GS.Row > NbLg ) then // si on est sur la dernière ligne en mode guide...
  FinGuide ;                      //  ... alors le guide est fini
// Ajout SBO 22/07/2004 ( FQ 10233 )
GS.SynEnabled:=True ;
END ;

procedure TFSaisie.ClickInsert ;
Var R : integer ;
BEGIN
if PasModif then exit ;
if ModeLC then Exit ;
if Not EstRempli(GS,GS.Row) then Exit ;
if Not LigneCorrecte(GS.Row,False,True) then Exit ;
GS.SynEnabled:=False ;
GS.CacheEdit ;
R:=GS.Row ; InsereLigne(GS.Row) ;
GS.SetFocus ; GS.Row:=R ; GS.Col:=SA_Gen ; GS.Cells[Gs.Col,Gs.Row]:='' ;
GereNewLigne(GS) ;
GS.MontreEdit ; Revision:=False ;
GereOptionsGrid(GS.Row) ;
GS.SynEnabled:=True ;
END ;

procedure TFSaisie.ClickAbandon ;
BEGIN
if ModeForce=tmPivot then ClickSwapPivot else BEGIN ModeLC:=False ; Fermeture() ; END ;
END ;

procedure TFSaisie.ClickSolde (Egal,RowEx : boolean) ;
Var b,bb : boolean ;
BEGIN
if PasModif then Exit ;
if ModeLC then Exit ;
if Not BSolde.Enabled then Exit ;
if Not ExJaiLeDroitConcept(TConcept(ccSaisSolde),True) then Exit ;
if Not EstRempli(GS,GS.Row) then Exit ;
if ((Egal) and (GS.Col<>SA_Gen) and (GS.Col<>SA_Aux) and (GS.Col<>SA_Debit) and (GS.Col<>SA_Credit)) then Exit ;
if PasToucheLigne(GS.Row) then Exit ;
GereNewLigne(GS) ;
if ((GS.Col=SA_Debit) or (GS.Col=SA_Credit)) then ChercheMontant(GS.Col,GS.Row) ;
SoldeLaLigne(GS.Row) ;
if ((VH^.AutoExtraSolde) or (RowEx)) then BEGIN b:=False ; bb:=False ; GSRowExit(GS,GS.Row,b,bb) ; END ;
{$IFDEF FFF}
// Essai Pb FFF validation de pièce SBO 29/16/2006
BValide.Enabled := (GS.RowCount>2) and Equilibre;
{$ELSE FFF}
BValide.Enabled:= ((H_NumeroPiece.Visible) and (TRUE)) ;
{$ENDIF FFF}
END ;

Function TFSaisie.SortePiece : boolean ;
Var Achat : boolean ;
BEGIN
Achat:=False ;
Case SorteTva of
   stvAchat : Achat:=True ;
   stvVente : Achat:=False ;
   else BEGIN
        if ((E_NATUREPIECE.Value='FF') or (E_NATUREPIECE.Value='AF') or (E_NATUREPIECE.Value='OF') or (E_NATUREPIECE.Value='RF')) then Achat:=True else
        if SAJAL.NatureJAL='ACH' then Achat:=True else

        if ((E_NATUREPIECE.Value='FC') or (E_NATUREPIECE.Value='AC') or (E_NATUREPIECE.Value='OC') or (E_NATUREPIECE.Value='RC')) then Achat:=False else
        if SAJAL.NatureJAL='VTE' then Achat:=False ;
        END ;
   END ;
Result:=Achat ;
END ;

function TFSaisie.ClickControleTVA ( Bloque : boolean ; LigHT : integer ) : boolean ;
{$IFNDEF GCGC}
{ dolar IFNDEF CCS3}
Var ia,LigTVA,LigTPF : integer ;
    X   : TFSaisTVA ;
    SDH,SCH,STDH,STCH,SDA,SCA,STDA,STCA,SDF,SCF,STDF,STCF : Double ;
    CDA,CCA,CTDA,CTCA,CDF,CCF,CTDF,CTCF : Double ;
    EcartTVALig,EcartTPFLig,EcartTVAPiece,EcartTPFPiece : double ;
    CHT,CTVA,CTPF : String17 ;
    TauxA,TauxF,Resol : Double ;
    Achat,EncON : boolean ;
    OHT   : TOBM ;
    CodeTva,CodeTpf : String3 ;
{ dolar ENDIF}
{$ENDIF}
BEGIN
{ dolar IFNDEF CCS3}
{$IFNDEF GCGC}
Result:=(Not Bloque) or ((EstSerie(S3)) and (VH^.PaysLocalisation<>CodeISOES)) ; //XVI 24/02/2005
if ((Not BControleTVA.Enabled) and (Not Bloque)) or (M.ANouveau) or ((EstSerie(S3)) and (VH^.PaysLocalisation<>CodeISOES)) then Exit ;
if GS.RowCount<=2 then Exit ; if ((Not Bloque) or (LigHT<=0)) then LigHT:=GS.Row ;
if Not EstRempli(GS,LigHT) then Exit ;
if ModeLC then Exit ;
if Not isHT(GS,LigHT,True) then BEGIN if not Bloque then HTVA.Execute(8,caption,'') ; Exit ; END ;
ia:=TrouveLigTiers(GS,0) ; if ia<=0 then BEGIN if not Bloque then HTVA.Execute(9,caption,'') ; Exit ; END ;
CHT:=GS.Cells[SA_Gen,LigHT] ; if CHT='' then Exit ;
OHT:=GetO(GS,LigHT) ; if OHT=Nil then Exit ;
CodeTva:=OHT.GetMvt('E_TVA') ; CodeTpf:=OHT.GetMvt('E_TPF') ;
EncON:=(OHT.GetMvt('E_TVAENCAISSEMENT')='X') ;
// Init des variables
SDH:=0 ; SCH:=0 ; STDH:=0 ; STCH:=0 ; SDA:=0 ; SCA:=0 ; STDA:=0 ; STCA:=0 ; SDF:=0 ; SCF:=0 ; STDF:=0 ; STCF:=0 ;
// Recup infos régime, Codes TVA et TPF du compte HT
Achat:=SortePiece ;
AttribRegimeEtTva ;
TauxA:=Tva2Taux(GeneRegTVA,CodeTva,Achat) ; TauxF:=Tpf2Taux(GeneRegTVA,CodeTpf,Achat) ;
if EncON then BEGIN CTVA:=Tva2Encais(GeneRegTVA,CodeTva,Achat) ; CTPF:=Tpf2Encais(GeneRegTVA,CodeTpf,Achat) ; END
         else BEGIN CTVA:=Tva2Cpte(GeneRegTVA,CodeTva,Achat) ; CTPF:=Tpf2Cpte(GeneRegTVA,CodeTpf,Achat) ; END ;
if ((CTVA='') and (TauxA<>0)) or ((CTPF='') and (TauxF<>0)) then
   BEGIN
   if Not Bloque then HTVA.Execute(10,caption,'') ; Exit ;
   END ;
// Calculs et recup de la ligne HT
InfoLigne(LigHT,SDH,SCH,STDH,STCH) ;
// Calculs et recup de la ligne TVA
if CTVA<>'' then BEGIN LigTVA:=TrouveLigCompte(GS,SA_Gen,0,CTVA) ;
if LigTVA>0 then InfoLigne(LigTVA,SDA,SCA,STDA,STCA) ; END ;
CDA:=HT2TVA(SDH,GeneRegTVA,GeneSoumisTPF,CodeTva,CodeTpf,Achat,DECDEV) ;
CCA:=HT2TVA(SCH,GeneRegTVA,GeneSoumisTPF,CodeTva,CodeTpf,Achat,DECDEV) ;
CTDA:=HT2TVA(STDH,GeneRegTVA,GeneSoumisTPF,CodeTva,CodeTpf,Achat,DECDEV) ;
CTCA:=HT2TVA(STCH,GeneRegTVA,GeneSoumisTPF,CodeTva,CodeTpf,Achat,DECDEV) ;
// Calculs et recup de la ligne TPF
if CTPF<>'' then BEGIN LigTPF:=TrouveLigCompte(GS,SA_Gen,0,CTPF) ;
if LigTPF>0 then InfoLigne(LigTPF,SDF,SCF,STDF,STCF) ; END ;
CDF:=HT2TPF(SDH,GeneRegTVA,GeneSoumisTPF,CodeTva,CodeTpf,Achat,DECDEV) ;
CCF:=HT2TPF(SCH,GeneRegTVA,GeneSoumisTPF,CodeTva,CodeTpf,Achat,DECDEV) ;
CTDF:=HT2TPF(STDH,GeneRegTVA,GeneSoumisTPF,CodeTva,CodeTpf,Achat,DECDEV) ;
CTCF:=HT2TPF(STCH,GeneRegTVA,GeneSoumisTPF,CodeTva,CodeTpf,Achat,DECDEV) ;
// Calculs des écarts
EcartTVALig:=Arrondi((SDA-CDA)-(SCA-CCA),DEV.Decimale) ; EcartTPFLig:=Arrondi((SDF-CDF)-(SCF-CCF),DEV.Decimale) ;
EcartTVAPiece:=Arrondi((STDA-CTDA)-(STCA-CTCA),DEV.Decimale) ; EcartTPFPiece:=Arrondi((STDF-CTDF)-(STCF-CTCF),DEV.Decimale) ;
//Scenario de saisie
if ((OkScenario) and (Bloque)) then
   BEGIN
   Resol:=Resolution(DEV.Decimale) ;
   Result:=((Abs(EcartTVALig)<=Resol) and (Abs(EcartTPFLig)<=Resol) and
            (Abs(EcartTVAPiece)<=Resol) and (Abs(EcartTPFPiece)<=Resol)) ;
   Exit ;
   END ;
// Appel de la form VisuTVA
X:=TFSaisTVA.Create(Application) ;
 Try
  ChangeFormatDevise(X,DECDEV,DEV.Symbole) ; X.Decim:=DECDEV ;
  With X do
    BEGIN
    // Comptes
    SGC.Text:=CHT  ; SAC.Text:=CTVA ; CAC.Text:=CTVA ; EAC.Text:=CTVA ;
    SFC.Text:=CTPF ; CFC.Text:=CTPF ; EFC.Text:=CTPF ;
    // Saisie
    AfficheLeSolde(SGL,SDH,SCH) ; AfficheLeSolde(SGP,STDH,STCH) ;
    AfficheLeSolde(SAL,SDA,SCA) ; AfficheLeSolde(SAP,STDA,STCA) ;
    AfficheLeSolde(SFL,SDF,SCF) ; AfficheLeSolde(SFP,STDF,STCF) ;
    // Calculs
    AfficheLeSolde(CAL,CDA,CCA) ; AfficheLeSolde(CAP,CTDA,CTCA) ;
    AfficheLeSolde(CFL,CDF,CCF) ; AfficheLeSolde(CFP,CTDF,CTCF) ;
    // Ecarts
    AfficheLeSolde(EAL,SDA-CDA,SCA-CCA) ; AfficheLeSolde(EAP,STDA-CTDA,STCA-CTCA) ;
    AfficheLeSolde(EFL,SDF-CDF,SCF-CCF) ; AfficheLeSolde(EFP,STDF-CTDF,STCF-CTCF) ;
    // Info Ligne
    HInfo.Caption:=HInfo.Caption+IntToStr(LigHT)+'   '+GS.Cells[SA_Gen,LigHT]+'   '+GS.Cells[SA_RefI,LigHT] ;
    END ;
  X.ShowModal ;
 Finally
  X.Free ;
 End ;
{$ENDIF}
{ dolar ENDIF}
END ;

procedure TFSaisie.RegroupeTVA ;
Var LigTVA,LigNext,NbT,NbN : integer ;
    CpteTVA : String17 ;
    DTVA,CTVA,DNext,CNext : Double ;
    lMOD   : TMOD ;         // FQ 21630 15/10/2007 : Gestion des comptes lettrables...
    TvaGen : TGGeneral ;    // FQ 21630 15/10/2007 : Gestion des comptes lettrables...
    XD,XP  : Double ;       // FQ 21630 15/10/2007 : Gestion des comptes lettrables...
BEGIN
LigTVA:=0 ; NbT:=0 ; NbN:=0 ;
Repeat
 LigTVA:=TrouveLigTVA(GS,LigTVA) ; Inc(NbT) ;
 if ((LigTVA>0) and (LigTVA<GS.RowCount-1)) then
    BEGIN
    CpteTVA:=GS.Cells[SA_Gen,LigTVA] ; DTVA:=ValD(GS,LigTVA) ; CTVA:=ValC(GS,LigTVA) ;
    Repeat
     LigNext:=TrouveLigCompte(GS,SA_Gen,LigTVA,CpteTVA) ;
     if ((LigNext>0) and (LigNext<GS.RowCount-1)) then
        BEGIN
        DNext:=ValD(GS,LigNext) ; CNext:=ValC(GS,LigNext) ;
        DTVA:=DTVA+DNext ; CTVA:=CTVA+CNext ;
        DetruitLigne(LigNext,False) ; Inc(NbN) ;
        END ;
    Until ((LigNext<=0) or (LigNext>=GS.RowCount) or (NbN>50)) ;
    if Abs(DTVA)>Abs(CTVA) then BEGIN DTVA:=DTVA-CTVA ; CTVA:=0 ; END else BEGIN CTVA:=CTVA-DTVA ; DTVA:=0 ; END ;
    GS.Cells[SA_Debit,LigTVA]:=StrS(DTVA,DECDEV) ; GS.Cells[SA_Credit,LigTVA]:=StrS(CTVA,DECDEV) ;
    FormatMontant(GS,SA_Debit,LigTVA,DECDEV) ; FormatMontant(GS,SA_Credit,LigTVA,DECDEV) ;
    TraiteMontant(LigTVA,False) ; AjusteLigne(LigTVA,True) ;
    // FQ 21630 15/10/2007 : Gestion des comptes lettrables...
    TvaGen := GetGGeneral(GS,LigTVA) ;
    if Lettrable(TvaGen)<>NonEche then
      begin
      lMOD := TMOD(GS.Objects[SA_NumP,LigTVA]) ;
      if Assigned(lMOD) then
        begin
        XD := MontantLigne(GS,LigTVA,tsmDevise) ;
        XP := MontantLigne(GS,LigTVA,tsmPivot) ;
        if (lMOD.MODR.TotalAPayerD<>XD) or (lMOD.MODR.TotalAPayerP<>XP) then
          RecalculProrataEche(lMOD.MODR,XP,XD) ;
        end ;
      end ;
    // Fin FQ 21630 15/10/2007
    END ;
Until ((LigTVA<=0) or (LigTVA>=GS.RowCount-1) or (NbT>50)) ;
END ;

procedure TFSaisie.ClickGenereTVA ;
Var Nb,i,iAux,iHT,iTva : integer ;
    RegTVA,TvaEnc : String3 ;
    SoumisTPF,Achat  : Boolean ;
    CAux       : TGTiers ;
    CGen       : TGGeneral ;
    OBM        : TOBM ;
    OkCpteTVA : boolean ;
    Cpte       : String17 ;
    TotTTC,Ht       : Double ;
    Sens,jj,oldNb   : Integer ;
    COdeTVA,COdeTPF : String ;
BEGIN
TotTTC := 0;
Sens   := 1;
if PasModif then Exit ;
if ModeLC then Exit ;
if ((M.ANouveau) or (Not BGenereTVA.Enabled)) then Exit ;
OkMessAnal:=False ; CoherTVA:=True ; OkCpteTVA:=True ;
// Contrôles de présence 1 seul TTC et au moins 1 HT
// Essai sur FQ 17564 (ajout "and IsHT(GS,GS.Row,True)" ): finalement mis en place dans clickvalide
//if EstRempli(GS,GS.Row) and IsHT(GS,GS.Row,True) and Not LigneCorrecte(GS.Row,True,True) then exit ;  //SG6 03/12/2004 FQ 15048
if ((SAJAL.NATUREJAL<>'ACH') and (SAJAL.NATUREJAL<>'VTE')) then BEGIN HTva.Execute(0,caption,'') ; Exit ; END ;
Achat:=SortePiece ; oldNB := 0 ;
iAux:=TrouveLigTiers(GS,0) ; if iAux<=0 then BEGIN HTVA.Execute(1,caption,'') ; Exit ; END ;
if TrouveLigTiers(GS,iAux)>0 then BEGIN HTVA.Execute(2,caption,'') ; Exit ; END ;
iHT:=TrouveLigHT(GS,0,TRUE) ; if iHT<=0 then BEGIN HTVA.Execute(3,caption,'') ; Exit ; END ;

  if GS.Cells[SA_Aux,iAux]<>'' then
    BEGIN
    CAux:=GetGTiers(GS,iAux) ; RegTVA:=GeneRegTVA ; SoumisTpf:=CAux.SoumisTPF ;
    if Achat then TvaEnc:=CAux.TVA_Encaissement else TvaEnc:=VH^.TvaEncSociete ;
    Cpte:=GS.Cells[SA_Aux,iAux] ;
    END
  else
    BEGIN
    CGen:=GetGGeneral(GS,iAux) ; RegTVA:=GeneRegTVA ; SoumisTpf:=CGen.SoumisTPF ;
    if Achat then TvaEnc:=CGen.TVA_Encaissement else TvaEnc:=VH^.TvaEncSociete ;
    Cpte:=GS.Cells[SA_Gen,iAux] ;
    END ;

  if RegTVA='' then BEGIN HTVA.Execute(6,caption,Cpte) ; Exit ; END ;
  iHT:=0 ; Nb:=0 ;

  //XVI 24/02/2005
  if VH^.PaysLocalisation=CodeISOES then //XVI 24/02/2005
    begin
    // Destruction des lignes TVA/TPF concernées déja présentes
    Repeat
      iTva:=TrouveLigTVA(GS,0) ; Inc(Nb) ;
      if iTva>0 then DetruitLigne(iTva,False) ;
      Until ((iTva<=0) or (Nb>100)) ;
    End ;

  // Remplissage des comptes HT avec infos de TVA/TPF associées
  Repeat
    iHT:=TrouveLigHT(GS,iHT,True) ; Inc(Nb) ;
    if iHT>0 then
      AlimHTByTVA(iHT,RegTVA,TvaEnc,SoumisTpf,Achat) ;
    Until ((iHT<=0) or (Nb>100)) ;
  GS.CacheEdit ;

  //XVI 24/02/2005
  if VH^.PaysLocalisation<>CodeISOES then //XVI 24/02/2005
    begin
    // Destruction des lignes TVA/TPF concernées déja présentes
    Nb:=0 ;
    Repeat
      iTva:=TrouveLigTVA(GS,0) ; Inc(Nb) ;
      if iTva>0 then DetruitLigne(iTva,False) ;
      Until ((iTva<=0) or (Nb>100)) ;
    End ;
  //XVI 24/02/2005

  Nb:=GS.RowCount-2 ;

  // Génération des lignes TVA/TPF
if VH^.PaysLocalisation=CodeISOES then //XVI 24/02/2005
 begin
  TotTTC:=Arrondi(Valeur(GS.Cells[SA_Debit,iAux]),DecDev) ;
  Sens:=1 ;
  if TotTTC=0 then
    begin
    Sens:=2 ;
    TotTTC:=Arrondi(Valeur(GS.Cells[SA_Credit,iAux]),DecDev) ;
    end ;
 End ;
  for i:=1 to Nb do if isHT(GS,i,True) then
    BEGIN
    OBM:=GetO(GS,i) ;
    OBM.PutMvt('E_REGIMETVA',RegTVA) ;

    //XVI 24/02/2005
    if VH^.PaysLocalisation=CodeISOES then
      begin
      HT:=Arrondi(Valeur(GS.Cells[SA_Debit,i])+Valeur(GS.Cells[SA_Credit,i]),DecDev) ;
      if HT=0 then
        Begin
        OBM:=GetO(GS,i) ;
        CodeTva:=OBM.GetMvt('E_TVA') ;
        CodeTpf:=OBM.GetMvt('E_TPF') ;
        HT:=TTC2HT(TotTTC,RegTVA,SoumisTPF,CodeTVA,CodeTPF,Achat,DecDev) ;
        if Sens=2
          then GS.Cells[SA_Debit,i]:=StrS(HT,DECDEV)
          else GS.Cells[SA_Credit,i]:=StrS(HT,DECDEV) ;
        FormatMontant(GS,SA_Debit,i,DECDEV) ;
        FormatMontant(GS,SA_Credit,i,DECDEV) ;
        TraiteMontant(i,False) ;
        End ;
      TotTTC:=TotTTC-HT ;
      oldNB:=GS.RowCount-2 ;
      End ; //XVI 24/02/2005

    if Not InsereLigneTVA(i,RegTVA,SoumisTpf,Achat) then OkCpteTVA:=False ;

    //XVI 24/02/2005
    if VH^.PaysLocalisation=CodeISOES then
       For jj:=OldNB to GS.RowCount-2 do
           TotTTC:=TotTTC-arrondi(Valeur(GS.Cells[SA_Debit,jj])+Valeur(GS.Cells[SA_Credit,jj]),DecDev) ;
    //XVI 214/02/2005

    END ;

  // Regroupement selon scénario
  if ((OkScenario) and (Scenario.GetMvt('SC_CONTROLETVA')='GTG')) then RegroupeTVA ;

  // Fin et ré-init du traitement
  CalculDebitCredit ;
  GS.MontreEdit ;

{$IFDEF FFF}
// Essai Pb FFF validation de pièce SBO 29/16/2006
  BValide.Enabled := (GS.RowCount>2) and Equilibre;
{$ELSE FFF}
  BValide.Enabled:=((H_NumeroPiece.Visible) and (Equilibre)) ;
{$ENDIF FFF}
  if OkMessAnal then HTVA.Execute(4,caption,'')  ;
  if Not OkCpteTVA
    then HTVA.Execute(7,caption,'')
    else if Not CoherTVA
      then HTVA.Execute(5,caption,'')  ;
  Revision:=False ; OkMessAnal:=False ; CoherTVA:=True ;
END ;

procedure TFSaisie.ClickZoom(DoPopup : boolean);
Var b : byte ;
    C,R : longint ;
    A   : TActionFiche ;
    lMulQSav : TQuery ;
BEGIN
if ((E_JOURNAL.Value='') or (Not GS.Enabled)) then Exit ;
if ((Action=taConsult) or (ModeForce<>tmNormal))
    then BEGIN GS.MouseToCell(GX,GY,C,R) ; A:=taConsult ; END
    else BEGIN R:=GS.Row ; C:=GS.Col ; A:=taModif ; END ;
if R<=0 then BEGIN R:=GS.Row ; C:=SA_Gen ; END ;
if R<1 then Exit ; if C<SA_Gen then Exit ;
if ((Action=taConsult) or (ModeLC)) and (GS.Cells[C,R]='') then Exit ;
if ModeForce<>tmNormal then Exit ;

// Patch AGL qui ajoute l'order by du theMulQ si celui renseigné à l'entrée dans une fiche type TOM...
// FQ 15989 SBO 02/08/2005
lMulQSav := TheMulQ ;
TheMulQ  := nil;

if C=SA_Gen then
   BEGIN
   {$IFNDEF GCGC}
   if ((Not ExJaiLeDroitConcept(TConcept(ccGenModif),False)) or (ModeLC)) then A:=taConsult ;
   b := ChercheGen(C,R,False,DoPopup);
   if b=2 then
      BEGIN
      FicheGene(Nil,'',GS.Cells[C,R],A,0) ;
      if Action<>taConsult then ChercheGen(C,R,True) ;
      END ;
   {$ELSE}
   Exit ;
   {$ENDIF}
   END else if C=SA_Aux then
   BEGIN
   if ((Not ExJaiLeDroitConcept(TConcept(ccAuxModif),False)) or (ModeLC)) then A:=taConsult ;
   b:=ChercheAux(C,R,False,DoPopup) ;
   {$IFNDEF GCGC}
   if b=2 then
     BEGIN
       FicheTiers(Nil,'',GS.Cells[C,R],A,1) ;
       if Action<>taConsult then ChercheAux(C,R,True) ;
     END ;
   {$ELSE}
   if b=2 then AGLLanceFiche('GC','GCTIERS','',GS.Cells[C,R],ActionToString(taConsult)+';MONOFICHE;T_NATUREAUXI=CLI') ;
   {$ENDIF}
   END ;

// FQ 15989 SBO 02/08/2005
TheMulQ  := lMulQSav ;

END ;

procedure TFSaisie.ClickConsult ;
{$IFDEF COMPTA}
var
    SG, LeExo : String;
{$ENDIF}
BEGIN
{$IFDEF COMPTA}
if ModeForce<>tmNormal then Exit ;
if ModeLC then Exit ;
if ((E_JOURNAL.Value='') or (Not GS.Enabled)) then Exit ;

// SBO 23/04/2004 : correction appel consultation avec recherche sur auxiliaire
  if GS.Cells[SA_Aux,GS.Row]<>'' then
    // Recherche du tiers par Hcompte, car recherche consultation des comptes insuffisantes :
    if ChercheAux(SA_Aux,GS.Row,False)<=0 then Exit ;
// Fin correction appel consultation avec recherche sur auxiliaire

SG:=GS.Cells[SA_Gen,GS.Row] ; if SG='' then Exit ;
if QuelExoDT(StrToDate(E_DATECOMPTABLE.Text))=VH^.Encours.Code then LeExo:='0' else LeExo:='1' ;
{$IFNDEF CMPGIS35}
OperationsSurComptes(SG,LeExo,'',GS.Cells[SA_Aux,GS.Row],True) ;
{$ENDIF}
{$ENDIF}
END ;

procedure TFSaisie.ClickSwapPivot ;
Var Cancel,Chg : boolean ;
BEGIN
// FQ 18704 : On Autorise la vision en devise pivot des pièces d'ECC
if (Not ((M.Nature='ECC') or ModeDevise) and (ModeForce=tmNormal)) then Exit ;
if ModeLC then Exit ;
if Not BSwapPivot.Enabled then Exit ;
if ((EstRempli(GS,GS.Row)) and (Not LigneCorrecte(GS.Row,False,True))) then Exit ;
if Not EstRempli(GS,1) then Exit ;
if ModeForce=tmNormal then
   BEGIN
   ModeForce:=tmPivot ;
   OldDevise:=DEV.Code ;
   OldTauxDEV:=DEV.Taux ;
   DEV.Decimale:=V_PGI.OkDecV ;
   DECDEV:=V_PGI.OkDecV ;
   E_DEVISE.Value:=V_PGI.DevisePivot ;
   E_DEVISE.Enabled:=True ;
   E_Devise.Font.Color:=clRed ;
   AffecteGrid(GS,taConsult) ;
   GS.SetFocus ;
   PEntete.Enabled:=False ;
   Outils.Enabled:=False ;
   BValide.Enabled:=False ; // FQ 15423 SBO 24/03/2005
   END else
   BEGIN
   E_DEVISE.Enabled:=False ;
   E_DEVISE.Value:=OldDevise ;
   E_Devise.Font.Color:=clBlack ;
   DEV.Taux:=OldTauxDEV ;
   ModeForce:=tmNormal ;
   PEntete.Enabled:=True ;
   Outils.Enabled:=True ;
   if Outils.CanFocus then Outils.SetFocus ;
   BValide.Enabled:=True ; // FQ 15423 SBO 24/03/2005
   GS.Col:=SA_Gen ;
   GS.Row:=1 ;
   GS.CacheEdit ;
   AffecteGrid(GS,Action) ;
   GS.Col:=SA_Gen ;
   GS.Row:=2 ;
   GS.SetFocus ;
   GS.MontreEdit ;
   GS.Row:=1 ;
   Cancel:=False ;
   Chg:=False ;
   GSRowEnter(Nil,GS.Row,Cancel,Chg) ;
   END ;
ChangeAffGrid(GS,ModeForce,DECDEV) ;
CalculDebitCredit ;
END ;

procedure TFSaisie.ClickModifRIB ;
Var O : TOBM ;
    IsAux : Boolean ;
BEGIN
if Not BModifRIB.Enabled then Exit ;
O:=GetO(GS,GS.Row) ; if O=Nil then Exit ;
IsAux:=O.GetMvt('E_AUXILIAIRE')<>'' ;
ModifRIBOBM(O,False,FALSE,'',IsAux) ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 11/03/2005
Modifié le ... : 09/09/2005
Description .. : - LG - 11/03/2005 - recherche des info de cutoff pour les
Suite ........ : affe ds la fenetre de saisie des info comp.
Suite ........ : - LG - 09/09/2005 - rajout des new apram de cutoff
Mots clefs ... : 
*****************************************************************}
procedure TFSaisie.ClickComplement ;
{$IFNDEF GCGC}
Var ModBN : boolean ;
    O     : TOBM ;
    Lig   : integer ;
    RC    : R_COMP ;
    OQ1,OQ2,NQ1,NQ2 : Double ;
{$ENDIF}
begin
{$IFNDEF GCGC}
Lig:=GS.Row ;
if Not EstRempli(GS,Lig) then Exit ;
if Not LigneCorrecte(Lig,False,True) then Exit ;
O:=GetO(GS,Lig) ; OQ1:=O.GetMvt('E_QTE1') ; OQ2:=O.GetMvt('E_QTE2') ;
RC.StLibre:='XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX' ;
if ModeLC then RC.StComporte:='XXXXXXXXXX' else RC.StComporte:='--XXXXXXXX' ;
RC.Conso:=True ; RC.DateC:=DatePiece ;
O:=GetO(GS,Lig) ; if O=Nil then Exit ;
if not FInfo.LoadCompte( GS.Cells[SA_Gen,Lig] ) then exit ;
RC.CutOffPer := FInfo.Compte.GetValue('G_CUTOFFPERIODE') ;
RC.CutOffEchue := FInfo.Compte.GetValue('G_CUTOFFECHUE') ;
RC.Attributs:=False ; RC.MemoComp:=Nil ; RC.Origine:=-1 ; RC.TOBCompl := CGetTOBCompl(O) ;
if Not SaisieComplement(O,EcrGen,Action,ModBN,RC) then Exit ;
if GS.CanFocus then GS.SetFocus ;
if MODBN then Revision:=False ;
NQ1:=O.GetMvt('E_QTE1') ; NQ2:=O.GetMvt('E_QTE2') ;
// Modif SBO 08/03/04 : Utiliser la fonction ProrateQteAnalTOB à la place de ProrateQteAnal : 
if ((NQ1<>OQ1) and (OQ1<>0)) then ProrateQteAnalTOB(O,OQ1,NQ1,'1') ;
if ((NQ2<>OQ2) and (OQ2<>0)) then ProrateQteAnalTOB(O,OQ2,NQ2,'2') ;

{$ENDIF}
end;

Function TFSaisie.ChoixFromTenueSaisie ( LeMontantPIVOT,LeMontantEURO,LeMontantDEV : Double ) : Double ;
BEGIN
if DEV.Code<>V_PGI.DevisePivot then Result:=LeMontantDEV
else Result:=LeMontantPIVOT ;
END ;

Function TFSaisie.EcheancesLC ( NbC : integer ) : boolean ;
Var Q : TQuery ;
    Lig  : integer ;
    OBM  : TOBM ;
    TM   : TMOD ;
    TotC,LeMontant : Double ;
    sNumTraChq : String ;
BEGIN
Result:=False ; Lig:=0 ; if ModeLC then Exit ;
{Tests préalables}
if ((Action<>taCreat) and (Not SuiviMP)) then Exit ;
if M.SorteLettre=tslAucun then Exit ;
if SuiviMP then
   BEGIN
   END else
   BEGIN
   if ((M.SorteLettre in [tslBOR,tslCheque,tslVir]) and (M.TypeGuide<>'DEC')) then Exit ;
   if ((M.SorteLettre in [tslTraite,tslPre]) and (M.TypeGuide<>'ENC')) then Exit ;
   END ;
{Vider toutes les lignes}
ClickVidePiece(False) ;
if Guide then FinGuide ;  // FQ 15683 : SBO 21/04/2005
{Lire le fichiers des échéances éditées}
InitMove(NbC,'') ; TotC:=0 ; sNumTraChq:='' ;
Q:=OpenSQL('Select * from DOCREGLE Where DR_USER="'+V_PGI.User+'" AND DR_EDITE="X" ORDER BY DR_NUMCHEQUE',True) ;
if Not Q.EOF then
   BEGIN
   ModeLC:=True ;
   While Not Q.EOF do
      BEGIN
      {Création des lignes de tiers}
      Inc(Lig) ; DefautLigne(Lig,True) ; MoveCur(False) ;
      OBM:=GetO(GS,Lig) ; GereNewLigne(GS) ;
      GS.Cells[SA_Gen,Lig]:=Q.FindField('DR_GENERAL').AsString ;
      GS.Cells[SA_Aux,Lig]:=Q.FindField('DR_AUXILIAIRE').AsString ;
      LeMontant:=ChoixFromTenueSaisie(Q.FindField('DR_MONTANT').AsFloat,0,Q.FindField('DR_MONTANTDEV').AsFloat) ;
      GS.Cells[SA_RefI,Lig]:=Q.FindField('DR_NUMCHEQUE').AsString ;
      GS.Cells[SA_Lib,Lig]:=Q.FindField('DR_LIBELLE').AsString ;
      // Modif SBO prélèvements
      if M.SorteLettre in [tslTraite,tslPre]
        then GS.Cells[SA_Credit,Lig]:=StrS(LeMontant,DECDEV)
        else GS.Cells[SA_Debit,Lig]:=StrS(LeMontant,DECDEV) ;
      // Fin Modif SBO prélèvements
      ValideCeQuePeux(Lig) ;
      AlimObjetMvt(Lig,False) ;
      RemplirEcheAuto(Lig) ;
      RemplirAnalAuto(Lig) ;
      OBM.PutMvt('E_RIB',Q.FindField('DR_RIB').AsString) ;
      OBM.PutMvt('E_MODEPAIE',Q.FindField('DR_MODEPAIE').AsString) ;
      OBM.PutMvt('E_DATEECHEANCE',Q.FindField('DR_DATEECHEANCE').AsDateTime) ;
      sNumTraChq:=GetNumLotTraChq(GS.Cells[SA_RefI,Lig]) ;
      OBM.PutMvt('E_NUMTRAITECHQ',sNumTraChq) ;
      // VL 301003 FQ 12958 et JP 19/08/05 FQ 10094 et 14853 pour M.Bordereau
      if M.ExportCFONB or M.Bordereau then
        OBM.PutMVT('E_CFONBOK', '#');
      // FIN VL

      //VT 06/02/2003 if M.SorteLettre<>tslCheque then OBM.PutMvt('E_CFONBOK','X') ;
      TM:=TMOD(GS.Objects[SA_NumP,Lig]) ; if TM<>Nil then EcheLettreCheque(Lig,TM.MODR) ;
      AlimGuideVentil(Lig) ;
      if SuiviMP then
        AlimTableTemporaire(OBM,'',M.smp,'','') ;
      {Création des lignes de banque}
      if (Not M.Globalise) { FICHE 12351 or (M.SorteLettre in [tslTraite,tslBOR]) } then
         BEGIN
         Inc(Lig) ; DefautLigne(Lig,True) ;
         OBM:=GetO(GS,Lig) ; GereNewLigne(GS) ;
         if M.SorteLettre in [tslTraite,tslBOR] then GS.Cells[SA_Gen,Lig]:=M.Section else GS.Cells[SA_Gen,Lig]:=M.General ;
         GS.Cells[SA_Aux,Lig]:=Q.FindField('DR_AUXILIAIRE').AsString ;
         // Modif SBO prélèvements
         if M.SorteLettre in [tslTraite,tslPre]
           then GS.Cells[SA_Debit,Lig]:=StrS(LeMontant,DECDEV)
           else GS.Cells[SA_Credit,Lig]:=StrS(LeMontant,DECDEV) ;
         // Fin Modif SBO prélèvements
         if SuiviMP then
         begin
          if (M.FormuleRefCCMP.Ref2 <> '') then GS.Cells[SA_RefI,Lig]:= M.FormuleRefCCMP.Ref2; // Référence génération
          if (M.FormuleRefCCMP.Lib2 <> '') then GS.Cells[SA_Lib,Lig] := M.FormuleRefCCMP.Lib2; // Libellé génération
         end;
         ValideCeQuePeux(Lig) ; AlimObjetMvt(Lig,False) ;
         RemplirEcheAuto(Lig) ; RemplirAnalAuto(Lig) ;
         OBM.PutMvt('E_MODEPAIE',Q.FindField('DR_MODEPAIE').AsString) ;
         OBM.PutMvt('E_BANQUEPREVI',Q.FindField('DR_BANQUEPREVI').AsString) ;
         OBM.PutMvt('E_DATEECHEANCE',Q.FindField('DR_DATEECHEANCE').AsDateTime) ;
         OBM.PutMvt('E_RIB',Q.FindField('DR_RIB').AsString) ;
         OBM.PutMvt('E_NUMTRAITECHQ',sNumTraChq) ;

         TM:=TMOD(GS.Objects[SA_NumP,Lig]) ; if TM<>Nil then EcheLettreCheque(Lig,TM.MODR) ;
         END else
         BEGIN
         TotC:=TotC+LeMontant ;
         END ;
      Q.Next ;
      END ;
   END ;
Ferme(Q) ;
{Dernière ligne globalisée}
if M.Globalise then
   BEGIN
   Inc(Lig) ; DefautLigne(Lig,True) ;
   OBM:=GetO(GS,Lig) ; GereNewLigne(GS) ;
   // Modif Fiche 12351
   if M.SorteLettre in [tslTraite,tslBOR]
    then GS.Cells[SA_Gen,Lig]:=M.Section
    else GS.Cells[SA_Gen,Lig]:=M.General ;
   if M.SorteLettre in [tslTraite,tslPre]
    then GS.Cells[SA_Debit,Lig]  := StrS(TotC,DECDEV)
    else GS.Cells[SA_Credit,Lig] := StrS(TotC,DECDEV) ;
   // Fin Modif Fiche 12351
   if SuiviMP then
   begin
    GS.Cells[SA_RefI,Lig]:= M.FormuleRefCCMP.Ref2; // Référence génération
    GS.Cells[SA_Lib,Lig] := M.FormuleRefCCMP.Lib2; // Libellé génération
   end;
   ValideCeQuePeux(Lig) ; AlimObjetMvt(Lig,False) ;
   RemplirEcheAuto(Lig) ; RemplirAnalAuto(Lig) ;
   OBM.PutMvt('E_MODEPAIE',LastPaie) ;
   OBM.PutMvt('E_DATEECHEANCE',OBM.GetMvt('E_DATECOMPTABLE')) ;
   TM:=TMOD(GS.Objects[SA_NumP,Lig]) ; if TM<>Nil then EcheLettreCheque(Lig,TM.MODR) ;
   END ;
{Ajustements}
GereNewLigne(GS) ;
CalculDebitCredit ;
GereOptionsGrid(GS.Row) ;

{JP 29/07/05 : FQ 16032 : Le lettrage se fait sur toutes les écritures sélectionnées dans le mul
               et non pas sur celles qui ont réellement été traitées dans DocRegl}
MajTEcheOrig;

{Supprimer les enregistrements temporaires}
ExecuteSQL('DELETE from DOCREGLE Where DR_USER="'+V_PGI.USER+'"') ;
ExecuteSQL('DELETE from DOCFACT Where DF_USER="'+V_PGI.USER+'"') ;
FiniMove ;
If SuiviMP Then AffecteRef(TRUE) ;
Result:=True ;
END ;

procedure TFSaisie.ClickCheque ;
{$IFNDEF GCGC}
Var LL        : TList ;
    i,Err,NbC : integer ;
    O         : TOBM ;
    MM        : TMOD ;
    Premier   : Boolean ;
    CurMP     : String3 ;
    OkPrintDialog : Boolean ;
    {$IFNDEF CCS3}
    DevDocRegl : TDevDocRegl ;
    {$ENDIF}
{$ENDIF}
BEGIN
if Action=taConsult then Exit ;
{$IFNDEF GCGC}
{$IFNDEF CCS3}
if Not BCheque.Enabled then Exit ;
if Not OkLC then Exit ;
if Not PieceCorrecte then Exit ;
LL:=TList.Create ; Err:=0 ; CurMP:='' ; Premier:=True ; DevDocRegl:=OnLocale ;
for i:=1 to GS.RowCount-1 do
    BEGIN
    O:=GetO(GS,i) ; if O=Nil then Break ;
    If i=1 Then
      BEGIN
      If O.GetMvt('E_DEVISE')<>V_PGI.DevisePivot Then DevDocRegl:=OnDevise Else
        BEGIN
          DevDocRegl:=OnLocale ;
        END ;
      END ;
    MM:=TMOD(GS.Objects[SA_NumP,i]) ;
    if MM=Nil
      then Continue
      else if MM.ModR.NbEche>1
           then BEGIN Err:=9 ; Break ; END ;
    If M.TIDTIC then
      begin
      // FQ 15590 SBO 04/04/20005
      if ( not FInfo.LoadCompte( O.GetMvt('E_GENERAL') ) ) or
         ( ( FInfo.Compte_GetValue('G_NATUREGENE') <> 'TIC' ) and
           ( FInfo.Compte_GetValue('G_NATUREGENE') <> 'TID' )
         ) then Continue ;
      end
      Else if O.GetMvt('E_AUXILIAIRE')=''
           then Continue ;
    if ((M.SorteLettre in [tslTraite,tslBOR]) and (O.GetMvt('E_GENERAL')=M.Section)) then
    // EN ATTENTE MODIF LETTRAGE : SBO
{      if SuiviMP then
        begin
        // indicateur bidon pour repérer les lignes de comptes de génération
        if O.GetValue('E_NUMCFONB')='CCMP' then
          begin
          O.PutValue('E_NUMCFONB', '' );
          Continue ;
          end ;
        end
      else} continue ; // cas standard non CCMP
    if ((SuiviMP) and (M.MPGUnique<>'')) then
       BEGIN
       CurMP:=M.MPGUnique ;
       END else
       BEGIN
       if Premier then CurMP:=MM.ModR.TabEche[1].ModePaie else
          if MM.ModR.TabEche[1].ModePaie<>CurMP then BEGIN Err:=10 ; Break ; END ;
       END ;
    {Préparation OBM échéance}
    O.PutMvt('E_DATEECHEANCE',MM.ModR.TabEche[1].DateEche) ;
    O.PutMvt('E_MODEPAIE',CurMP) ;
    if (M.SorteLettre=tslCheque) Or (M.SorteLettre=tslVir) Or (M.SorteLettre=tslPre) then O.PutMvt('E_CONTREPARTIEGEN',M.General) ;
    LL.Add(O) ;
    Premier:=False ;
    END ;
if Err<=0 then
   BEGIN
   OkPrintDialog:=TRUE ;
   If M.MSED.MultiSessionEncaDeca And M.MSED.SessionFaite Then OkPrintDialog:=FALSE ;
   If M.MSED.MultiSessionEncaDeca Then BAbandon.Enabled:=FALSE ;
   Err:=12 ;
   If M.TIDTIC
     Then NbC:=LanceDocReglTID(LL, M.SorteLettre, '', M.GroupeEncadeca, SuiviMP,
                              M.MSED, OkPrintDialog, DevDocRegl, M.smp, TOBParamEsc)
     Else NbC:=LanceDocRegl(LL, M.SorteLettre, '', M.GroupeEncadeca, SuiviMP,
                            M.MSED, OkPrintDialog, DevDocRegl, M.smp, TOBParamEsc) ;
   // Attention bloc modifié 08/2003 suite pb maj E_NUMTRAITECHQ // FICHE 12017
   if NbC=-1
     then Err:=123456
     else if NbC>0 then
       begin
       // stoker les numéro de traite / chq par échéance d'origine
       stockNumChqDepuisEche( LL ) ;
       if (M.smp=smpEncTraEdtNC) Or (M.smp=smpDecBorEdtNC) then
         begin
         if Transactions(SetNumTraiteDOCREGLE,1)<>oeOk then
           MessageAlerte(HDiv.Mess[21]) ;
         Err:=0 ;
         end
       else
         if EcheancesLC(NbC) then
          Err:=0 ;
       end ;
   END ;
LL.Free ;
if Err>0 then
   BEGIN
   if Err<>123456 then HDiv.Execute(Err,'','') ;
   CEstGood:=False ;
   PieceModifiee:=False ;
   Fermeture();
   END ;
{$ENDIF}
{$ENDIF}
END ;

procedure TFSaisie.ClickCherche ;
BEGIN
if GS.RowCount<=2 then Exit ;
FindFirst:=True ; FindSais.Execute ;
END ;

procedure TFSaisie.ClickLibelleAuto ;
Var StL,Swhere1,SWhere2 : String ;
    Q   : TQuery ;
    i,OldL : integer ;
BEGIN
if ModeLC then Exit ;
if ((M.ANouveau) or (Not BLibAuto.Enabled) or (Action=taConsult) or (Guide)) then Exit ;
if Not AutoCharged then
   BEGIN
   SWhere1:='' ; SWhere2:='' ;
   If SAJAL<>NIL Then SWhere1:=' AND RA_JOURNAL="'+SAJAL.Journal+'" ' ;
   SWhere2:=' RA_NATUREPIECE="'+E_NATUREPIECE.Value+'" ' ;
   If SWhere1<>'' Then SWhere2:=SWhere2+SWhere1 ;
   Q:=OpenSQL('SELECT RA_FORMULEREF FROM REFAUTO WHERE '+SWhere2,True);
   if Q.EOF then
     StL:=Choisir(HDiv.Mess[1],'REFAUTO','RA_LIBELLE ||"(" || RA_JOURNAL || " " || RA_NATUREPIECE || ")"','RA_CODE','','', False, False, 7244310) // FQ 16095 SBO 11/10/2005
   else
     StL:=Choisir(HDiv.Mess[1],'REFAUTO','RA_LIBELLE ||"(" || RA_JOURNAL || " " || RA_NATUREPIECE || ")"','RA_CODE',SWhere2,'', False, False, 7244310) ;  // FQ 16095 SBO 11/10/2005
   Ferme(Q) ;
   if StL='' then BEGIN HPiece.Execute(15,caption,'') ; Exit ; END ;
   Q:=OpenSQL('Select RA_FORMULEREF, RA_FORMULELIB from REFAUTO where RA_CODE="'+StL+'"',True) ;
   SI_FormuleRef:=Q.FindField('RA_FORMULEREF').AsString ; SI_FormuleLib:=Q.FindField('RA_FORMULELIB').AsString ;
   Ferme(Q) ;
   END ;
OldL:=CurLig ;
for i:=1 to GS.RowCount-2 do BEGIN CurLig:=i ; GereRefLib(i) ; END ;
GS.EditorMode:=True ; CurLig:=OldL ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 18/03/2003
Modifié le ... : 24/09/2007
Description .. : - LG - 18/03/2003 - blocage du rafraichissement de la grille
Suite ........ : - LG - 12/09/2007 - suppression dela fct SelectGuide
Suite ........ : - LG - 24/09/2007 - FB 21495 - correction de l'appel des 
Suite ........ : guides
Mots clefs ... : 
*****************************************************************}
procedure TFSaisie.ClickGuide ;
Var StG         : String3 ;
BEGIN
{$IFNDEF GCGC}
GS.BeginUpdate ;
try
if ModeLC then Exit ;
if PasModif then Exit ; if ((M.ANouveau) or (Not BGuide.Enabled)) then Exit ;
if Guide then BEGIN FinGuide ; Exit ; END ;
if GS.RowCount>2 then Exit ; if EstRempli(GS,1) then Exit ;
if ((EstRempli(GS,GS.Row)) and (Not LigneCorrecte(GS.Row,False,True))) then Exit ;
StG := CPLanceFiche_MulGuide('','','SELECT;'+E_JOURNAL.Value+';'+E_NATUREPIECE.Value+';'+E_DEVISE.Value+';'+E_ETABLISSEMENT.Value) ;
if StG='' then Exit ;
LanceGuide('NOR',StG) ;
finally
GS.EndUpdate ;
end;
{$ENDIF}
END ;

procedure TFSaisie.ClickProrata ;
Var Rat : double ;
    i   : integer ;
    D,C : Double ;
    O   : TOBM ;
    Okok : boolean ;
BEGIN
if ModeLC then Exit ;
if PasModif then Exit ; if ((M.ANouveau) or (Not BProrata.Enabled)) then Exit ;
if Not ExJaiLeDroitConcept(TConcept(ccSaisProrata),True) then Exit ;
if Not EstRempli(GS,GS.Row) then Exit ;
if Not LigneCorrecte(GS.Row,False,False) then Exit ;
if ((GS.Col=SA_Debit) or (GS.Col=SA_Credit)) then CaptureRatio(GS.Row,False) ;
O:=GetO(GS,GS.Row) ; if O=Nil then Exit ; Rat:=O.GetMvt('RATIO') ;
if ((Rat=0) or (Rat=1)) then Exit ;
Okok:=True ; for i:=1 to GS.RowCount-2 do if PasToucheLigne(i) then BEGIN Okok:=False ; Break ; END ;
if Not Okok then BEGIN HPiece.Execute(31,caption,'') ; Exit ; END ;
for i:=1 to GS.RowCount-2 do if i<>GS.Row then
    BEGIN
    D:=ValD(GS,i) ; C:=ValC(GS,i) ;
    GS.Cells[SA_Debit,i]:=StrS(ArrS(D*Rat),DECDEV) ; GS.Cells[SA_Credit,i]:=StrS(ArrS(C*Rat),DECDEV) ;
    FormatMontant(GS,SA_Debit,i,DECDEV) ; FormatMontant(GS,SA_Credit,i,DECDEV) ;
    TraiteMontant(i,False) ; AjusteLigne(i,True) ;
    END ;
AlimObjetMvt(GS.Row,True) ;
CalculDebitCredit ;
CaptureRatio(GS.Row,True) ;
END ;

procedure TFSaisie.ClickDernPieces ;
BEGIN
{$IFNDEF GCGC}
VisuDernieresPieces(TPIECE) ;
{$ENDIF}
END ;

procedure TFSaisie.BCValideClick(Sender: TObject);
Var Q     : TQuery ;
//    i     : integer ;
    ValSt,CodeG : String ;
begin
if FNomGuide.Text='' then BEGIN HPiece.Execute(21,caption,'') ; Exit ; END ;
if HPiece.Execute(22,caption,'')<>mrYes then BEGIN BCAbandonClick(Nil) ; Exit ; END ;
Q:=OpenSQL('Select GU_GUIDE from GUIDE Where GU_TYPE="NOR" AND UPPER(GU_LIBELLE)="'+uppercase(FNomGuide.Text)+'"',True) ;
if Not Q.EOF then
   BEGIN
   if HPiece.Execute(23,caption,'')<>mrYes then BEGIN Ferme(Q) ; Exit ; END ;
   CodeG:=Q.Fields[0].AsString ;
   EXECUTESQL('DELETE FROM GUIDE WHERE GU_TYPE="NOR" AND GU_GUIDE="'+CodeG+'"') ;
   EXECUTESQL('DELETE FROM ECRGUI WHERE EG_TYPE="NOR" AND EG_GUIDE="'+CodeG+'"') ;
   EXECUTESQL('DELETE FROM ANAGUI WHERE AG_TYPE="NOR" AND AG_GUIDE="'+CodeG+'"') ;
   END ;
Ferme(Q) ;
(* GP le 03/09/2008 : 23437
Q:=OpenSQL('Select MAX(GU_GUIDE) from GUIDE Where GU_TYPE="NOR"',True) ;
if Not Q.EOF then
   BEGIN
   ValSt:=Q.Fields[0].AsString ;
   if ValSt<>'' then BEGIN i:=StrToInt(ValSt) ; ValSt:=IntToStr(i+1) ; END else ValSt:='001' ;
   While Length(ValSt)<3 do ValSt:='0'+ValSt ;
   END else ValSt:='001' ;
Ferme(Q) ;
*)

// N°23437 :
ValSt:=CCalculeCodeGuide('NOR') ;
// Fin N°23437 :
Application.ProcessMessages ;
BeginTrans ;
CreerEnteteGuide(ValSt) ;
CreerLignesGuide(ValSt) ;
CommitTrans ;
AvertirTable('ttGuideEcr') ; 
BCAbandonClick(Nil) ;
end;

procedure TFSaisie.BCAbandonClick(Sender: TObject);
begin
PEntete.Enabled:=True ; Outils.Enabled:=True ; Valide97.Enabled:=True ;
GS.Enabled:=True ; PPied.Enabled:=True ;
GS.SetFocus ; PCreerGuide.Visible:=False ;
ModeCG:=False ; FNomGuide.Text:='' ;
end;

procedure TFSaisie.ClickCreerGuide ;
Var NatJ : String3 ;
BEGIN
if ModeLC then Exit ;
if Not GS.Enabled then Exit ; if GS.RowCount<=2 then Exit ;
if Not ExJaiLeDroitConcept(TConcept(ccSaisCreatGuide),True) then Exit ;
if ((M.Anouveau) or (Not BCreerGuide.Enabled)) then Exit ;
if ((EstRempli(GS,GS.Row)) and (Not LigneCorrecte(GS.Row,False,True))) then Exit ;
if SAJAL=Nil then Exit ; NatJ:=SAJAL.NatureJal ;
// Modif SBO : On autorise la création pour les journaux de caisse
if ((NATJ='ANO') or (NATJ='CLO') or (NATJ='ODA') or (NATJ='ANA') or (NATJ='ECC')) then
  BEGIN HPiece.Execute(29,caption,'') ; Exit ; END ;
PCreerGuide.Left:=GS.Left+(GS.Width-PCreerGuide.Width) div 2 ;
PCreerGuide.Top:=GS.Top+(GS.Height-PCreerGuide.Height) div 2 ;
PCreerGuide.Visible:=True ; FNomGuide.SetFocus ;
PEntete.Enabled:=False ; Outils.Enabled:=False ; Valide97.Enabled:=False ;
GS.Enabled:=False ; PPied.Enabled:=False ;
ModeCG:=True ;
END ;

procedure TFSaisie.ClickVidePiece ( Parle : boolean ) ;
BEGIN
if PasModif then Exit ;
if Not PieceModifiee then Exit ;
if Parle then
   BEGIN
   if ModeLC then Exit ;
   if HPiece.Execute(14,caption,'')<>mrYes then Exit ;
   END ;
GS.Col:=SA_Gen ; GS.CacheEdit ;
GS.VidePile(True) ; GS.Row:=1 ; GS.MontreEdit ;
DefautLigne(1,True) ;
END ;

procedure TFSaisie.ClickChancel ;
Var DD : TDateTime ;
    AA : TActionFiche ;
BEGIN
if ModeLC then Exit ;
if Action<>taCreat then Exit ;
if DEV.Code=V_PGI.DevisePivot then Exit ;
if ModeForce<>tmNormal then Exit ;
if Not BChancel.Enabled then Exit ;
DD:=StrToDate(E_DATECOMPTABLE.Text) ;
AA:=taModif ; if Action=taConsult then AA:=taConsult ;
FicheChancel(DEV.Code,True,DD,AA,(DD>=V_PGI.DateDebutEuro)) ;
if AA<>taConsult then DEV.Taux:=GetTaux(DEV.Code,DEV.DateTaux,DD) ;
END ;

procedure TFSaisie.TabTvaDansModR ( OMODR : TMOD ) ;
Var Coeff,Tot : Double ;
    i,k   : integer ;
BEGIN
Tot:=OMODR.MODR.TotalAPayerP ; if Tot=0 then Exit ;
for i:=1 to OMODR.MODR.NbEche do
    BEGIN
    Coeff:=OMODR.MODR.TabEche[i].MontantP/Tot ;
    for k:=1 to 5 do OMODR.MODR.TabEche[i].TAV[k]:=Arrondi(TabTvaEnc[k]*Coeff,V_PGI.OkDecV) ;
    END ;
END ;

procedure TFSaisie.ClickModifTva ;
{$IFNDEF GCGC}
Var Lig : integer ;
    O   : TOBM ;
    OMODR : TMOD ;
    Client : Boolean ;
    CAux   : TGTiers ;
    CGen   : TGGeneral ;
    AA     : TActionFiche ;
    RTVA   : Enr_Base ;
{$ENDIF}
BEGIN
{#TVAENC}
{$IFNDEF GCGC}
if Not OuiTvaLoc then Exit ;
if Not BModifTva.Enabled then Exit ;
Lig:=GS.Row ; O:=GetO(GS,Lig) ; if O=Nil then Exit ;
AA:=Action ;
if PasToucheLigne(Lig) then AA:=taConsult ;
if IsHT(GS,Lig,True) then
   BEGIN
   if O.GetMvt('E_REGIMETVA')='' then O.PutMvt('E_REGIMETVA',GeneRegTva) ;
{$IFDEF COMPTA}
   if (VH^.PaysLocalisation=CodeISOES) and
      (InfosTvaEnc(O,AA)) and (isTiersSoumisTPF(GS)) then
      Begin
      O.PutMvt('E_TPF',O.GetValue('E_TVA')) ;
      CalculDebitCredit ;
      End else
{$ENDIF}
   if InfosTvaEnc(O,AA) then CalculDebitCredit ;
   END else if IsTiers(GS,Lig) then
   BEGIN
   if SorteTva<>stvDivers then BEGIN AA:=taConsult ; if Not UnSeulTiers then Exit ; END ;
   if GS.Cells[SA_Aux,Lig]<>'' then
      BEGIN
      CAux:=GetGTiers(GS,Lig) ; if CAux=Nil then Exit ;
      if ((CAux.NatureAux='CLI') or (CAux.NatureAux='AUD')) then Client:=True else
       if ((CAux.NatureAux='FOU') or (CAux.NatureAux='AUC')) then Client:=False else Exit ;
      END else if GS.Cells[SA_Gen,Lig]<>'' then
      BEGIN
      CGen:=GetGGeneral(GS,Lig) ; if CGen=Nil then Exit ;
      if CGen.NatureGene='TID' then Client:=True else
       if CGen.NatureGene='TIC' then Client:=False else Exit ;
      END else Exit ;
   CGen:=GetGGeneral(GS,Lig) ; if CGen=Nil then Exit ;
   if Not EstCollFact(CGen.General) then Exit ;
   OMODR:=TMOD(GS.Objects[SA_NumP,Lig]) ; if OMODR=Nil then Exit ;
   if ((SorteTva<>stvDivers) and (AA=taConsult)) then TabTvaDansModR(OMODR) ;
   RTVA.Regime:=O.GetMvt('E_REGIMETVA') ; RTVA.Nature:=E_NATUREPIECE.Value ;
   RTVA.Client:=Client ; RTVA.Action:=AA ; RTVA.CodeTva:=O.GetMvt('E_TVA') ;
   if SaisieBasesHT(OMODR,RTVA) then O.PutMvt('E_TVA',RTVA.CodeTva) ;
   END else Exit ;
Revision:=False ;
{$ENDIF}
END ;

procedure TFSaisie.ClickScenario ;
{$IFNDEF GCGC}
var lStLequel : string ;
{$ENDIF}
BEGIN
{$IFNDEF GCGC}
  // recherche clé de l'éventuel scénario chargé
  lStLequel := '' ;
  if OkScenario and Assigned( Scenario ) then
    // clé : SC_USERGRP, SC_JOURNAL, SC_NATUREPIECE, SC_QUALIFPIECE, SC_ETABLISSEMENT
    lStLequel := Scenario.GetString('SC_USERGRP') + ';' +
                 Scenario.GetString('SC_JOURNAL') + ';' +
                 Scenario.GetString('SC_NATUREPIECE') + ';' +
                 Scenario.GetString('SC_QUALIFPIECE') + ';' +
                 Scenario.GetString('SC_ETABLISSEMENT') ;
  CPLanceFiche_Scenario( lStLequel, '', False ) ;
{$ENDIF}
END ;

procedure TFSaisie.ClickJournal ;
{$IFNDEF GCGC}
Var a : TActionFiche ;
    lMulQSav : TQuery ;
{$ENDIF}
begin
{$IFNDEF GCGC}

if Action=taConsult then a:=taConsult else a:=taModif ;
if Not ExJaiLeDroitConcept(TConcept(ccJalModif),False) then a:=taConsult ;
if ((E_JOURNAL.Value='') or (SAJAL=Nil)) then Exit ;

// Patch AGL qui ajoute l'order by du theMulQ si celui renseigné à l'entrée dans une fiche type TOM...
// FQ 15989 SBO 02/08/2005
lMulQSav := TheMulQ ;
TheMulQ  := nil;

if ModeLC then a:=taConsult ;
FicheJournal(Nil,'',E_JOURNAL.Value,a,0) ;
if a<>taConsult then ChercheJal(E_JOURNAL.Value,True) ;

// FQ 15989 SBO 02/08/2005
TheMulQ  := lMulQSav ;

{$ENDIF}
end;

procedure TFSaisie.ClickPieceGC(FromPce : boolean=true);
Var RefGC : String ;
    O     : TOBM ;
begin
  if (FromPce) and (Not BPieceGC.Enabled) then Exit;
  if (not FromPce) and (Not bLignesGC.Enabled) then Exit;
  O := GetO(GS, GS.Row) ;
  if O = Nil then Exit ;
  RefGC := O.GetMvt('E_REFGESCOM') ;
  if RefGC = '' then Exit ;
  // BDU - 16/11/06, si on trouve un préfixe AA, AD ou AR, on visualise les frais
  if (Copy(RefGC, 1, 3) = 'AA;') or (Copy(RefGC, 1, 3) = 'AD;') or (Copy(RefGC, 1, 3) = 'AR;') then
    LanceZoomFrais(RefGC, O.GetMvt('E_AFFAIRE'))
  else
  begin
{$IFNDEF CMPGIS35}
    if FromPce then
      LanceZoomPieceGC(RefGC)
    else if GetParamSocSecur('SO_OKLIENLIGECR', false) then
      LanceZoomLignesGC(O);
{$ENDIF}
  end;
end;

procedure TFSaisie.ClickImmo ;
{$IFDEF COMPTAAVECSERVANT}
Var
    i : Integer ;
    CodeI : String ;
    O : TOBM ;
{$ELSE}
{$IFDEF AMORTISSEMENT}
Var
    CodeI : String ;
    O : TOBM ;
{$ENDIF}
{$ENDIF}
begin
  {$IFNDEF CCMP}
    {$IFNDEF GCGC}
      {$IFNDEF IMP}
if Not BZoomImmo.Enabled then Exit ;
{$IFDEF COMPTAAVECSERVANT}
IF VHIM^.ServantPGI Then
  BEGIN
  O:=GetO(GS,GS.Row) ; CodeI:=Trim(O.GetMvt('E_IMMO')) ; if CodeI='' then Exit ;
  For i:=1 To Length(CodeI) Do If CodeI[i] in ['0'..'9']=FALSE Then Exit ;
  if Not Presence('IMOREF','IRF_NUM',CodeI) then HDiv.Execute(20,Caption,'')
                                       else VoirFicheImo(taConsult,StrToInt(CodeI)) ;
  END Else
  BEGIN
  O:=GetO(GS,GS.Row) ; CodeI:=O.GetMvt('E_IMMO') ; if CodeI='' then Exit ;
  if Not Presence('IMMO','I_IMMO',CodeI) then HDiv.Execute(20,Caption,'')
                                         else AMLanceFiche_FicheImmobilisation(CodeI, taConsult, '');
  END ;
{$ELSE}
{$IFDEF AMORTISSEMENT}
O:=GetO(GS,GS.Row) ; CodeI:=O.GetMvt('E_IMMO') ; if CodeI='' then Exit ;
if Not Presence('IMMO','I_IMMO',CodeI) then HDiv.Execute(20,Caption,'')
                                       else     AMLanceFiche_FicheImmobilisation(CodeI, taConsult, '');
{$ENDIF}
{$ENDIF}
      {$ENDIF}
    {$ENDIF}
  {$ENDIF}
end;

procedure TFSaisie.ClickTP ;
Var O : TOBM ;
    PieceTP : String ;
begin
if Not BZoomTP.Enabled then Exit ;
{$IFNDEF CCS3}
O:=GetO(GS,GS.Row) ; PieceTP:=O.GetMvt('E_PIECETP') ; if PieceTP='' then Exit ;
ZoomPieceTP(PieceTP) ;
{$ENDIF}
end;


procedure TFSaisie.BModifSerieClick(Sender: TObject);
{$IFNDEF GCGC}
Var b : boolean ;
    St : String ;
    Lig,k,i   : integer ;
    O         : TOBM ;
    DD        : TDateTime ;
    Okok      : boolean ;
    XX        : Double ;
    TM        : TMOD ;
    RC        : R_COMP ;
{$ENDIF}
begin
{$IFNDEF GCGC}
if Action=taConsult then Exit ;
ModifSerie.Free ; ModifSerie:=TOBM.Create(EcrGen,'',True) ;
RC.StLibre:='XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX' ;
RC.StComporte:='XXXXXXXXXX' ; RC.Conso:=True ; RC.DateC:=DatePiece ;
RC.Attributs:=False ; RC.MemoComp:=Nil ; RC.Origine:=-1 ; 
if Not SaisieComplement(ModifSerie,EcrGen,taCreat,b,RC) then Exit ;
if HPiece.Execute(4,caption,'')<>mrYes then Exit ; Okok:=False ;
for Lig:=1 to GS.RowCount-1 do
    BEGIN
    O:=GetO(GS,Lig) ; if O=Nil then Break ;
    if O.GetMvt('E_LETTRAGE')<>'' then BEGIN Okok:=True ; Continue ; END ;
    if O.GetMvt('E_REFPOINTAGE')<>'' then BEGIN Okok:=True ; Continue ; END ;
    TM:=TMOD(GS.Objects[SA_NumP,Lig]) ;
    if TM<>Nil then
       BEGIN
       for k:=1 to TM.MODR.NbEche do if TM.MODR.TabEche[k].ReadOnly then BEGIN Okok:=True ; Break ; END ;
       if Okok then Continue ;
       END ;
    St:=ModifSerie.GetMvt('E_REFINTERNE') ; if St<>'' then BEGIN GS.Cells[SA_RefI,Lig]:=St ; O.PutMvt('E_REFINTERNE',St) ; END ;
    St:=ModifSerie.GetMvt('E_LIBELLE') ; if St<>'' then BEGIN GS.Cells[SA_Lib,Lig]:=St ; O.PutMvt('E_LIBELLE',St) ; END ;
    St:=ModifSerie.GetMvt('E_REFEXTERNE') ; if St<>'' then O.PutMvt('E_REFEXTERNE',St) ;
    St:=ModifSerie.GetMvt('E_REFLIBRE') ; if St<>'' then O.PutMvt('E_REFLIBRE',St) ;
    St:=ModifSerie.GetMvt('E_AFFAIRE') ; if St<>'' then O.PutMvt('E_AFFAIRE',St) ;
    St:=ModifSerie.GetMvt('E_CONSO') ; if St<>'' then O.PutMvt('E_CONSO',St) ;
    DD:=ModifSerie.GetMvt('E_DATEREFEXTERNE') ; if DD>Encodedate(1901,01,01) then O.PutMvt('E_DATEREFEXTERNE',DD) ;
    {Zones libres}
    for i:=0 to 9 do
        BEGIN
        St:=ModifSerie.GetMvt('E_LIBRETEXTE'+IntToStr(i)) ; if St<>'' then O.PutMvt('E_LIBRETEXTE'+IntToStr(i),St) ;
        if i<=3 then
           BEGIN
           St:=ModifSerie.GetMvt('E_TABLE'+IntToStr(i)) ; if St<>'' then O.PutMvt('E_TABLE'+IntToStr(i),St) ;
           XX:=ModifSerie.GetMvt('E_LIBREMONTANT'+IntToStr(i)) ; if XX<>0 then O.PutMvt('E_LIBREMONTANT'+IntToStr(i),XX) ;
           if i<=1 then BEGIN St:=ModifSerie.GetMvt('E_LIBREBOOL'+IntToStr(i)) ; if St<>'' then O.PutMvt('E_LIBREBOOL'+IntToStr(i),St) ; END ;
           END ;
        END ;
    DD:=ModifSerie.GetMvt('E_LIBREDATE') ; if DD>Encodedate(1901,01,01) then O.PutMvt('E_LIBREDATE',DD) ;
    END ;
if Okok then HPiece.Execute(13,caption,'') ;
{$ENDIF}
end;

procedure TFSaisie.ClickDevise ;
BEGIN
FicheDevise(E_DEVISE.Value,taConsult,False) ;
END ;

procedure TFSaisie.ClickSaisTaux ;
BEGIN
if Not BSaisTaux.Enabled then Exit ;
if DEV.Code=V_PGI.DevisePivot then Exit ;
if DEV.Code='' then Exit ;
if Action<>taCreat then Exit ;
if ModeLC then Exit ;
if SaisieNewTaux2000(DEV,DatePiece) then BEGIN Volatile:=True ; M.TauxD:=DEV.Taux ; END ;
END ;

procedure TFSaisie.ClickChoixRegime ;
Var NewReg,Titre : String ;
    Lig : integer ;
    O   : TOBM ;
BEGIN
if Not BChoixRegime.Enabled then Exit ;
if Action=taConsult then Exit ;
if ModeLC then Exit ;
if ModeForce<>tmNormal then Exit ;
if GeneRegTva='' then GeneRegTVA:=VH^.RegimeDefaut ;
Titre:=HDiv.Mess[18]+' '+RechDom('TTREGIMETVA',GeneRegTVA,False)+')' ;
NewReg:=Choisir(Titre,'CHOIXCOD','CC_LIBELLE','CC_CODE','CC_TYPE="RTV"','') ;
if ((NewReg='') or (NewReg=GeneRegTVA)) then Exit ;
if HPiece.Execute(42,Caption,'')<>mrYes then Exit ;
GeneRegTVA:=NewReg ; RegimeForce:=True ;
for Lig:=1 to GS.RowCount-2 do
    BEGIN
    O:=GetO(GS,Lig) ; if O=Nil then Break ;
    O.PutMvt('E_REGIMETVA',GeneRegTVA) ;
    END ; 
END ;

procedure TFSaisie.ClickEtabl ;
BEGIN
FicheEtablissement_AGL(taConsult) ;
END ;

procedure TFSaisie.BVentilClick(Sender: TObject);
begin ClickVentil ; if GS.Enabled then GS.SetFocus ; end;

procedure TFSaisie.BEcheClick(Sender: TObject);
begin ClickEche ; if GS.Enabled then GS.SetFocus ; end;

procedure TFSaisie.BSDelClick(Sender: TObject);
begin ClickDel ; if GS.Enabled then GS.SetFocus ; end;

procedure TFSaisie.BVidePieceClick(Sender: TObject);
begin ClickVidePiece(True) ; end;

procedure TFSaisie.BInsertClick(Sender: TObject);
begin ClickInsert ; if GS.Enabled then GS.SetFocus ; end;

procedure TFSaisie.BValideClick(Sender: TObject);
begin
if ModeForce<>tmNormal then Exit ;
if ((M.TypeGuide<>'NOR') or (Not M.FromGuide)) then ClickValide else ClickAbandon ;
end;

procedure TFSaisie.BZoomDeviseClick(Sender: TObject);
begin ClickDevise ; end;

procedure TFSaisie.BModifTvaClick(Sender: TObject);
begin ClickModifTva ; end;

procedure TFSaisie.BSaisTauxClick(Sender: TObject);
begin ClickSaisTaux ; end;

procedure TFSaisie.BChoixRegimeClick(Sender: TObject);
begin ClickChoixRegime ; end;

procedure TFSaisie.BZoomEtablClick(Sender: TObject);
begin ClickEtabl ; end;

procedure TFSaisie.BSoldeClick(Sender: TObject) ;
BEGIN ClickSolde(False,False) ; if GS.Enabled then GS.SetFocus ; end;

procedure TFSaisie.BAbandonClick(Sender: TObject);
begin ClickAbandon ; end;

procedure TFSaisie.BZoomClick(Sender: TObject);
begin
ClickZoom ; if GS.Enabled then GS.SetFocus ;
end;

procedure TFSaisie.BScenarioClick(Sender: TObject);
begin ClickScenario ; end;

procedure TFSaisie.BZoomJournalClick(Sender: TObject);
begin ClickJournal ; end;

procedure TFSaisie.BPieceGCClick(Sender: TObject);
begin ClickPieceGC ; end;

procedure TFSaisie.BLignesGCClick(Sender : TObject);
begin
  ClickPieceGC(False);
end;

procedure TFSaisie.BZoomImmoClick(Sender: TObject);
begin ClickImmo ; end;

procedure TFSaisie.BZoomTPClick(Sender: TObject);
begin ClickTP ; end;

procedure TFSaisie.BGenereTvaClick(Sender: TObject);
begin ClickGenereTVA ; if GS.Enabled then GS.SetFocus ; end;

procedure TFSaisie.BControleTVAClick(Sender: TObject);
begin ClickControleTVA(False,0) ; if GS.Enabled then GS.SetFocus ; end;

procedure TFSaisie.BProrataClick(Sender: TObject);
begin ClickProrata ; if GS.Enabled then GS.SetFocus ; end;

procedure TFSaisie.BSwapPivotClick(Sender: TObject);
begin ClickSwapPivot ; end;

procedure TFSaisie.BModifRIBClick(Sender: TObject);
begin ClickModifRIB ; end;

procedure TFSaisie.BComplementClick(Sender: TObject);
begin ClickComplement ; end;

procedure TFSaisie.BChercherClick(Sender: TObject);
begin ClickCherche ; end;

procedure TFSaisie.BChequeClick(Sender: TObject);
begin ClickCheque ; end;

procedure TFSaisie.BGuideClick(Sender: TObject);
begin ClickGuide ; end;

procedure TFSaisie.BLibAutoClick(Sender: TObject);
begin ClickLibelleAuto ; if GS.Enabled then GS.SetFocus ; end;

procedure TFSaisie.BChancelClick(Sender: TObject);
begin ClickChancel ; end;

procedure TFSaisie.BDernPiecesClick(Sender: TObject);
begin ClickDernPieces ; if GS.Enabled then GS.SetFocus ; end;

procedure TFSaisie.BCreerGuideClick(Sender: TObject);
begin ClickCreerGuide ; if GS.Enabled then GS.SetFocus ; end;

{==========================================================================================}
{============================ Création Guide depuis saisie ================================}
{==========================================================================================}
function TFSaisie.BonRegle ( Lig,NbEche : integer ; LeMode : String ; T : T_TabEche ) : Boolean ;
Var MM : T_ModeRegl ;
    i  : integer ;
BEGIN
BonRegle:=False ; if LeMode='' then Exit ;
FillChar(MM,Sizeof(MM),#0) ;
MM.Action:=taCreat ;
MM.ModeInitial:=LeMode ; MM.ModeFinal:=LeMode ;
MM.TotalAPayerP:=MontantLigne(GS,Lig,tsmPivot) ;
MM.TotalAPayerD:=MontantLigne(GS,Lig,tsmDevise) ;                            
MM.CodeDevise:=DEV.CODE ; MM.Symbole:=DEV.Symbole ; MM.TauxDevise:=DEV.Taux ;
MM.Quotite:=DEV.Quotite ; MM.Decimale:=DEV.Decimale ;
MM.DateFact:=DateMvt(Lig) ; MM.DateBL:=DateMvt(Lig) ;
MM.Aux:=GS.Cells[SA_Aux,Lig] ;
CalculModeRegle(MM,True) ; // JLD ultérieurement, étudier le cas collectif+ventilable
if MM.NbEche<>NbEche then Exit ;
for i:=1 to MM.NbEche do
    BEGIN
    if MM.TabEche[i].ModePaie<>T[i].ModePaie then Exit ;
    if MM.TabEche[i].DateEche<>T[i].DateEche then Exit ;
    if Arrondi(MM.TabEche[i].MontantP-T[i].MontantP,V_PGI.OkDecV)>Resolution(V_PGI.OkDecV) then Exit ;
    if Arrondi(MM.TabEche[i].MontantD-T[i].MontantD,DEV.Decimale)>Resolution(DEV.Decimale) then Exit ;
    END ;
BonRegle:=True ;
END ;

procedure TFSaisie.CreerLignesGuide ( CodeG : String ) ;
Var
    TEcrG,TAnaG         : TOB ;
    Q                   : TQuery ;
    Lig, k, NbE         : integer ;
    NumAxe, Ind, NumV   : integer ;
    O                   : TOBM ;
    OBA                 : TOB ;
    OMR                 : TMOD ;
    Nom, NomG, LeMode   : String ;
    StArret, StEG       : String ;
begin
  for Lig:=1 to GS.RowCount-2 do
    begin
    if not EstRempli(GS,Lig) then Break ;
    // Recup données écritures
    O := GetO(GS,Lig) ;
    if O = Nil then Break ;
    // Allocation TOB sur ECRGUI
    TEcrG := TOB.Create('ECRGUI',nil,-1) ;
    TEcrG.InitValeurs ;
    // Remplissage des champs de ECRGUI
    for k:=0 to TEcrG.nbChamps-1 do
      begin
      // Recup nom du champ guide
      NomG := TEcrG.GetNomChamp(k+1) ;
      // Construction du nom du champ Ecriture équivalent
      Nom := NomG ;
      Delete(Nom,2,1) ;
      // Récupération de la valeur du TOBM
      Ind:=TrouveIndice(Nom,False) ;
      if Ind>0 then
        begin
        if VarType(O.T[Ind].V)=VarNull
          then StEG := ''
          else StEG := VarAsType(O.T[Ind].V,VarString) ;
        // Affectation valeur au guide
        TEcrG.PutValue( NomG , StEG) ;
        end ;
      end ;

    // Clé
    TEcrG.PutValue('EG_TYPE',     'NOR') ;
    TEcrG.PutValue('EG_GUIDE',    CodeG) ;
    TEcrG.PutValue('EG_NUMLIGNE', Lig) ;
    StArret:='------------------------------------------------' ;

    // Montants
    if CMontantsGuide.Checked then
      begin
      TEcrG.PutValue('EG_DEBITDEV' ,  GS.Cells[SA_Debit, Lig] ) ;
      TEcrG.PutValue('EG_CREDITDEV' , GS.Cells[SA_Credit,Lig] ) ;
      end
    else
      begin
      TEcrG.PutValue('EG_DEBITDEV'  , '') ;
      TEcrG.PutValue('EG_CREDITDEV' , '') ;
      if GS.Cells[SA_Debit,Lig]<>''
        then StArret[5] := 'X'
        else StArret[6] := 'X' ;
      end ;

    // Point d'arrêt pour la saisie guidé
    TEcrG.PutValue('EG_ARRET' , StArret) ;

    // Echéances
    OMR:=TMOD(GS.Objects[SA_NumP,Lig]) ;
    LeMode:='' ;
    if OMR<>Nil then
      begin
      NbE:=OMR.ModR.NbEche ;
      if BonRegle(Lig,NbE,OMR.ModR.ModeInitial,OMR.ModR.TabEche)
        then LeMode:=OMR.ModR.ModeInitial
        else
          if BonRegle(Lig,NbE,OMR.ModR.ModeFinal,OMR.ModR.TabEche)
            then LeMode:=OMR.ModR.ModeFinal
            else
              begin
              Q := OpenSQL('Select MR_MODEREGLE from MODEREGL Where MR_MODEREGLE<>"'+OMR.ModR.ModeInitial+'" '
                     +'AND MR_MODEREGLE<>"'+OMR.ModR.ModeFinal+'"',True) ;
              While Not Q.EOF do
                begin
                if BonRegle(Lig,NbE,Q.Fields[0].AsString,OMR.ModR.TabEche) then
                  begin
                  LeMode:=Q.Fields[0].AsString ;
                  Break ;
                  end ;
                Q.Next ;
                end ;
              Ferme(Q) ;
              end ;
      if LeMode='' then LeMode:=CreerModeRegle(OMR.ModR) ;
      end ;
    TEcrG.PutValue('EG_MODEREGLE' , LeMode ) ;

    // MAJ BASE
    TEcrG.InsertDB(nil) ;
    TEcrG.Free ;

    // Analytiques
   	for NumAxe := 1 to Min( MaxAxe, O.Detail.Count ) do
      begin
      for NumV := 0 to (O.Detail[NumAxe-1].Detail.Count - 1) do
        begin
        OBA := O.Detail[NumAxe-1].Detail[NumV] ;
        // Allocation tob ANAGUI
        TAnaG := TOB.Create('ANAGUI',nil,-1) ;
        TAnaG.InitValeurs ;
        // Recup valeur ventilations
        TAnaG.PutValue('AG_TYPE',         'NOR'                           ) ;
        TAnaG.PutValue('AG_GUIDE',        CodeG                           ) ;
        TAnaG.PutValue('AG_NUMLIGNE',     Lig                             ) ;
        TAnaG.PutValue('AG_NUMVENTIL',    NumV + 1                        ) ;
        TAnaG.PutValue('AG_SECTION',      OBA.GetValue('Y_SECTION')       ) ;
        TAnaG.PutValue('AG_AXE',          OBA.GetValue('Y_AXE')           ) ;
        TAnaG.PutValue('AG_ARRET',        '----'                          ) ;
        // montants
        if OBA.GetValue('Y_POURCENTAGE')<>0
          then TAnaG.PutValue('AG_POURCENTAGE',  FloatToStr(OBA.GetValue('Y_POURCENTAGE'))  )
          else TAnaG.PutValue('AG_POURCENTAGE',  '' ) ;
        if OBA.GetValue('Y_POURCENTQTE1')<>0
          then TAnaG.PutValue('AG_POURCENTQTE1', FloatToStr(OBA.GetValue('Y_POURCENTQTE1')) )
          else TAnaG.PutValue('AG_POURCENTQTE1', '' ) ;
        if OBA.GetValue('Y_POURCENTQTE2')<>0
          then TAnaG.PutValue('AG_POURCENTQTE2', FloatToStr(OBA.GetValue('Y_POURCENTQTE2')) )
          else TAnaG.PutValue('AG_POURCENTQTE2', '' ) ;
        // insertion base et libération
        TAnaG.InsertDB(nil) ;
        TAnaG.Free ;
        end ;
      end ;
    end ;
end ;


procedure TFSaisie.CreerEnteteGuide ( CodeG : String ) ;
Var
  TGuide : TOB ;
BEGIN
  ExecuteSQL('DELETE FROM GUIDE WHERE GU_TYPE="NOR" AND GU_GUIDE="'  + CodeG + '"' ) ;
  ExecuteSQL('DELETE FROM ECRGUI WHERE EG_TYPE="NOR" AND EG_GUIDE="' + CodeG + '"' ) ;
  ExecuteSQL('DELETE FROM ANAGUI WHERE AG_TYPE="NOR" AND AG_GUIDE="' + CodeG + '"' ) ;
  TGuide := TOB.Create('GUIDE', nil, -1) ;
  TGuide.InitValeurs ;
  TGuide.PutValue('GU_TYPE', 'NOR') ;
  TGuide.PutValue('GU_GUIDE', CodeG) ;
  TGuide.PutValue('GU_LIBELLE', FNomGuide.Text) ;
  TGuide.PutValue('GU_ABREGE', Copy(FNomGuide.Text,1,17)) ;
  TGuide.PutValue('GU_DATECREATION', Date) ;
  TGuide.PutValue('GU_DATEMODIF', Date) ;
  TGuide.PutValue('GU_JOURNAL', E_Journal.Value) ;
  TGuide.PutValue('GU_NATUREPIECE', E_NaturePiece.Value) ;
  TGuide.PutValue('GU_ETABLISSEMENT', E_Etablissement.Value) ;
  TGuide.PutValue('GU_DEVISE', E_Devise.Value) ;
  TGuide.PutValue('GU_UTILISATEUR', V_PGI.User) ;
  TGuide.PutValue('GU_SOCIETE', V_PGI.CodeSociete) ;
  TGuide.PutValue('GU_TRESORERIE', '-') ;
  TGuide.InsertDB(nil) ;
  TGuide.Free ;
END ;


{==========================================================================================}
{================================ Guide de saisie =========================================}
{==========================================================================================}
procedure TFSaisie.EtudieSiGuide(ACol,ARow : longint);
Var St : String ;
BEGIN
if PasModif then Exit ; if M.ANouveau then Exit ; 
St:=GS.Cells[ACol,ARow] ; if St='' then Exit ; if St[1]<>'*' then Exit ; if GS.Objects[ACol,ARow]<>Nil then Exit ;
System.Delete(St,1,1) ; if St<>'' then LanceGuide('NOR',St) ;
END ;

procedure TFSaisie.AlimGuideTVA ( Lig : integer ; Q : TQuery ) ;
Var O : TOBM ;
    St : String ;
BEGIN
{#TVAENC}
if Not OuiTvaLoc then Exit ;
if Action<>taCreat then Exit ;
if ((M.LeGuide=V_PGI.User) or (M.TypeGuide='ENC') or (M.TypeGuide='DEC')) then Exit ;
O:=GetO(GS,Lig) ; if O=Nil then Exit ;
St:=Q.FindField('EG_TVAENCAIS').AsString ; if St<>'' then O.PutMvt('E_TVAENCAISSEMENT',St) ;
St:=Q.FindField('EG_TVA').AsString ; if St<>'' then O.PutMvt('E_TVA',St) ;
END ;

procedure TFSaisie.AlimGuideVentil ( Lig : integer ) ;
var CGen : TGGeneral ;
BEGIN
if Action<>taCreat then Exit ;
if ((M.TypeGuide<>'ENC') and (M.TypeGuide<>'DEC')) then Exit ;
CGen:=GetGGeneral(GS,Lig) ; if CGen=Nil then Exit ;
if Ventilable(CGen,0) then RemplirAnalAuto(Lig) ;
END ;

procedure TFSaisie.AlimGuideLettrage ( Lig : integer ; Q : TQuery ) ;
Var O  : TOBM ;
    T  : HTStrings ;
    St : String ;
BEGIN
if Action<>taCreat then Exit ;
if ((M.TypeGuide<>'ENC') and (M.TypeGuide<>'DEC')) then Exit ;
O:=GetO(GS,Lig) ;
St:=Q.FindField('EG_RIB').AsString ; if ((St<>'') and (O.GetMvt('E_RIB')='')) then O.PutMvt('E_RIB',St) ;
St:=O.GetMvt('E_RIB') ;
If St='' Then AffecteLeRIB(O,O.GetMvt('E_AUXILIAIRE')) ;
St:=Q.FindField('EG_REFEXTERNE').AsString ; if ((St<>'') and (O.GetMvt('E_REFEXTERNE')='')) then O.PutMvt('E_REFEXTERNE',St) ;
St:=Q.FindField('EG_BANQUEPREVI').AsString ; O.PutMvt('E_BANQUEPREVI',St) ;
{$IFDEF EAGLCLIENT}
	if Q.FindField('EG_ECHEANCES').asString='' then Exit ;
{$ELSE}
	if IsFieldNull(Q,'EG_ECHEANCES') then Exit ;
{$ENDIF}
O.LC.Assign(TMemoField(Q.FindField('EG_ECHEANCES'))) ;

  {$IFDEF COMPTA}
  {JP 01/03/2005 : en complément de FQ 12017 : E_NUMTRAITECHQ n'est pas correctement renseigné
                   si deux écritures s'annulent}
  st := O.LC.Text;
  st := StrLeft(st, RPosAbs(';', st));
  O.SetString('E_TRACE', st);

  {$ENDIF COMPTA}

{Mémoriser les échéances d'origine}
T:=HTStringList.Create ; T.Assign(O.LC) ; TECHEORIG.Add(T) ;
END ;

procedure TFSaisie.ResaisirLaDatePourLeGuide(OkOk : Boolean) ;
{$IFNDEF GCGC}
{$IFNDEF IMP}
Var DD : TDateTime ;
    OBA   : TOB ;
    NumV,i,NumAxe : Integer ;
    OBM : TOBM ;
{$ENDIF}
{$ENDIF}
BEGIN
if not vh^.CPIFDEFCEGID then exit ;
{$IFNDEF GCGC}
{$IFNDEF IMP}
If Not MethodeGuideDate1 Then
  BEGIN
  if OkOk then
    BEGIN
    DD:=V_PGI.DateEntree ;
    SaisirDateGuide(DD) ;
    E_DATECOMPTABLE.Text:=DateToStr(DD) ;
    for i:=1 to GS.RowCount-2 do
       BEGIN
       OBM:=GetO(GS,i) ;
       if OBM<>NIL then OBM.PutMvt('E_DATECOMPTABLE',DD) ;

       for NumAxe:=1 to Min(MaxAxe, OBM.Detail.Count) do
        for NumV:=0 to OBM.Detail[NumAxe-1].Detail.Count-1 do
          begin
          OBA := OBM.Detail[NumAxe-1].Detail[NumV] ;
        	OBA.PutValue('Y_DATECOMPTABLE', DD) ;
          END ;
       END ;
    END ;
  END ;
{$ENDIF}
{$ENDIF}
END ;

procedure TFSaisie.LanceGuide ( TypeG,CodeG : String3 ) ;
Var OG : TOBM ;
    Premier : boolean ;
    St        : String ;
    Col,Lig   : integer ;
    Q,QE      : TQuery ;
BEGIN
if M.ANouveau then Exit ;
if vh^.CPIFDEFCEGID then  // verif ok
begin
If MethodeGuideDate1 Then
  BEGIN
  If (TypeG='NOR') And (Not appeldatepourguide) And (Not M.FromGuide) Then
    BEGIN
    CodeGuideFromSaisie:=CodeG ;
    appeldatepourguide:=TRUE ;
    E_DATECOMPTABLE.SetFocus ;
    Exit ;
    END ;
  END ;
appeldatepourguide:=FALSE ;
CodeGuideFromSaisie:='' ;
end ;
Q:=OpenSQL('Select * from ECRGUI Where EG_TYPE="'+TypeG+'" AND EG_GUIDE="'+CodeG+'" ORDER BY EG_GUIDE, EG_NUMLIGNE',True) ;
VideListe(TGUI) ; NbLG:=0 ; Premier:=True ; Guide:=False ; GuideAEnregistrer:=FALSE ; CodeGuideFromSaisie:='' ;
InitMove(RecordsCount(Q),'') ;
While Not Q.EOF do
   BEGIN
   if MoveCur(FALSE) then ;
   if Premier then
      BEGIN
      Guide:=True ; GuideAEnregistrer:=TRUE ;
      GS.Cells[SA_Gen,1]:='' ;
      if E_JOURNAL.Value='' then
         BEGIN
         QE:=OpenSQL('Select * from Guide where GU_TYPE="'+TypeG+'" AND GU_GUIDE="'+CodeG+'"',True) ;
         if Not QE.EOF then
            BEGIN
            E_JOURNAL.Value:=QE.FindField('GU_JOURNAL').AsString ;
            if QE.FindField('GU_NATUREPIECE').AsString<>'' then E_NATUREPIECE.Value:=QE.FindField('GU_NATUREPIECE').AsString ;
            if ((QE.FindField('GU_DEVISE').AsString<>'') and (QE.FindField('GU_DEVISE').AsString<>E_DEVISE.Value))
               then E_DEVISE.Value:=QE.FindField('GU_DEVISE').AsString
               else if E_DEVISE.Value='' then E_DEVISE.Value:=V_PGI.DevisePivot ;
            if QE.FindField('GU_ETABLISSEMENT').AsString<>''
               then E_ETABLISSEMENT.Value:=QE.FindField('GU_ETABLISSEMENT').AsString
               else E_ETABLISSEMENT.Value:=VH^.EtablisDefaut ;
//            if (E_DEVISE.Value=V_PGI.DevisePivot) then
//               BEGIN
//               E_DEVISE.ItemIndex:=E_DEVISE.Values.Count-1 ; E_DEVISEChange(Nil) ;
//               END ;
            END ;
         Ferme(QE) ;
         END ;
      END ;
   if E_JOURNAL.Value='' then BEGIN Guide:=False ; GuideAEnregistrer:=FALSE ; Break ; END ;
   OG:=TOBM.Create(EcrGui,'',True) ; OG.ChargeMvt(Q) ; Lig:=GS.RowCount-1 ;
   St:=Q.FindField('EG_GENERAL').AsString ; OG.PutMvt('EG_GENERAL',St) ; if St<>'' then GS.Cells[SA_Gen,Lig]:=GFormule(St,Load_Sais,Nil,1) ;
   GS.Cells[SA_Aux,Lig]:=Q.FindField('EG_AUXILIAIRE').AsString ;
   St:=Q.FindField('EG_REFINTERNE').AsString ; OG.PutMvt('EG_REFINTERNE',St) ;
   St:=Q.FindField('EG_LIBELLE').AsString ; OG.PutMvt('EG_LIBELLE',St) ;
   St:=Q.FindField('EG_DEBITDEV').AsString ; OG.PutMvt('EG_DEBITDEV',St) ;
   St:=Q.FindField('EG_CREDITDEV').AsString ; OG.PutMvt('EG_CREDITDEV',St) ;
   ValideCeQuePeux(Lig) ; AlimObjetMvt(Lig,True) ;
   AlimGuideLettrage(Lig,Q) ;
   AlimGuideTva(Lig,Q) ;
   TGUI.Add(OG) ;
   NewLigne(GS) ; DefautLigne(GS.RowCount-1,True) ;
   Q.Next ; Inc(NbLG) ; Premier:=False ;
   END ;
Ferme(Q) ; FiniMove ;
if Not Guide then Exit ;
RecalcGuide ; GS.Row:=1 ;
if ((Action=taCreat) and (TypeG='NOR') and (CodeG<>'')) then GuideAnal:=CodeG ;
Col:=ProchainArret(GS.Col,GS.Row) ;
GS.Enabled:=True ;
if (Col<=0) and ((M.TypeGuide='ENC') or (M.TypeGuide='DEC')) then
   BEGIN
   GS.Col:=SA_Gen ; GS.Row:=1 ; FinGuide ;  GS.SetFocus ;
   END else
   BEGIN
   FlashGuide.Visible   := True ;
   FlashGuide.Flashing  := True ;
   // Gestion de l'arrêt sur la date comptable // SBO FQ 14730
   if ( Col = SA_DATEC ) then
     begin
     E_DATECOMPTABLE.SetFocus ;
     end
   else
     begin
     GS.SetFocus ;
     if Col > 0
       then GS.Col := Col ;
     end ;
   END ;
if ((Action=taCreat) and ((M.TypeGuide='ENC') or (M.TypeGuide='DEC'))) then
   BEGIN
   BlocageEntete(Self,False) ;
   if TypeG=V_PGI.User then BEGIN LeTimer.Enabled:=True ; ModeLC:=True ; GridEna(GS,False) ; END ;
   END ;
if vh^.CPIFDEFCEGID then  // verif ok
   ResaisirLaDatePourLeGuide((Not M.FromGuide) And (TypeG='NOR') And (M.TypeGuide='')) ; // a partir de Alt G
END ;

procedure TFSaisie.RecalcGuide ;
Var OG  : TOBM ;
    Lig : integer ;
    St  : String ;
BEGIN
if Not Guide then Exit ;
for Lig:=1 to NbLG do
    BEGIN
    OG:=TOBM(TGUI[Lig-1]) ; if OG=Nil then Continue ; CurLig:=Lig ;
    St:=OG.GetMvt('EG_GENERAL') ; if ((St<>'') and (St[1] in ['[','{','('])) then GS.Cells[SA_Gen,Lig]:=GFormule(St,Load_Sais,Nil,1) ;
    St:=OG.GetMvt('EG_REFINTERNE') ;
    if St<>'' then
       BEGIN
       if ((EstFormule(St)) or (GS.Cells[SA_RefI,Lig]='')) then
          BEGIN
          if Trim(St)<>'' then GS.Cells[SA_RefI,Lig]:=GFormule(St,Load_Sais,Nil,1) else GS.Cells[SA_RefI,Lig]:=St ;
          END ;
       END ;
    St:=OG.GetMvt('EG_LIBELLE') ;
    if St<>'' then
       BEGIN
       {JP 27/07/05 : FQ 16216 : Le or est mieux que le and si le libellé n'est pas une formule}
       if ((EstFormule(St)) or (GS.Cells[SA_Lib,Lig]='')) then //SG6 01/12/2004 FQ 15000
          BEGIN
          if Trim(St)<>'' then GS.Cells[SA_Lib,Lig]:=GFormule(St,Load_Sais,Nil,1) else GS.Cells[SA_Lib,Lig]:=St ;
          END ;
       END ;
    St:=OG.GetMvt('EG_DEBITDEV') ;
    if St<>'' then
       BEGIN
       GS.Cells[SA_Debit,Lig]:=GFormule(St,Load_Sais,Nil,1) ;
       if GS.Cells[SA_Debit,Lig]='SOLDE' then
          BEGIN
          GS.Cells[SA_Debit,Lig]:='' ; CalculDebitCredit ;
          SoldeLaLigne(Lig) ;
          END ;
       END ;
    St:=OG.GetMvt('EG_CREDITDEV') ;
    if St<>'' then
       BEGIN
       GS.Cells[SA_Credit,Lig]:=GFormule(St,Load_Sais,Nil,1) ;
       if GS.Cells[SA_Credit,Lig]='SOLDE' then
          BEGIN
          GS.Cells[SA_Credit,Lig]:='' ; CalculDebitCredit ;
          SoldeLaLigne(Lig) ;
          END ;
       END ;
    FormatMontant(GS,SA_Debit,Lig,DECDEV) ; FormatMontant(GS,SA_Credit,Lig,DECDEV) ;
    AlimObjetMvt(Lig,True) ;
    if True then
       BEGIN
       if ((M.TypeGuide='DEC') or (M.TypeGuide='ENC')) then GereEche(Lig,False,True) ;
       AlimGuideVentil(Lig) ;
       END ;
    END ;
CalculDebitCredit ;
END ;

function TFSaisie.Load_TVATPF ( CGen : TGGeneral ; OKTVA : boolean ; TVA : String ) : Variant ;
BEGIN
Result:=0 ; if CGen=Nil then Exit ;
AttribRegimeEtTva ;
if OKTVA then
   BEGIN
   if TVA='' then TVA:=CGen.TVA ;
   Result:=Tva2Taux(GeneRegTVA,TVA,(SAJAL.NatureJAL='ACH')) ;
   END else Result:=Tpf2Taux(GeneRegTVA,CGen.TPF,(SAJAL.NatureJAL='ACH')) ;
END ;

function TFSaisie.Load_Sais ( St : hString ) : Variant ;
Var V    : Variant ;
    Lig  : integer ;
    OBM  : TOBM ;
    CGen : TGGeneral ;
    CAux : TGTiers ;
BEGIN
V:=#0 ; Result:=V ;
St:=uppercase(Trim(St)) ; if St='' then Exit ;
Lig:=QuelleLigne(St) ;
if ((Lig=-1) and (CurLig>1)) then Lig:=CurLig-1 else if Lig<=0 then Lig:=CurLig ;
if Lig<=0 then Lig:=GS.Row ;
OBM:=GetO(GS,Lig) ; CGen:=GetGGeneral(GS,Lig) ; CAux:=GetGTiers(GS,Lig) ;
{Zones entête}
if ((St='JOURNAL') or (St='E_JOURNAL')) then V:=SAJAL.JOURNAL else
if ((St='DATECOMPTABLE') or (St='E_DATECOMPTABLE')) then V:=StrToDate(E_DATECOMPTABLE.Text) else
if ((St='NATUREPIECE') or (St='E_NATUREPIECE')) then V:=E_NATUREPIECE.Text else
if ((St='NUMEROPIECE') or (St='E_NUMEROPIECE')) then V:=NumPieceInt else
if ((St='DEVISE') or (St='E_DEVISE')) then V:=E_DEVISE.Text else
if ((St='ETABLISSEMENT') or (St='E_ETABLISSEMENT')) then V:=E_ETABLISSEMENT.Text else
if St='E_GENERAL' then BEGIN if CGen<>Nil then V:=CGen.General ; END else
if St='E_AUXILIAIRE' then BEGIN if CAux<>Nil then V:=CAux.auxi ; END else
if St='TVA' then V:=Load_TVATPF(CGen,TRUE,'') else
if St='TVANOR' then V:=Load_TVATPF(CGen,TRUE,'NOR') else
if St='TVARED' then V:=Load_TVATPF(CGen,TRUE,'RED') else
if St='TPF' then V:=Load_TVATPF(CGen,False,'') else
if St='E_REFINTERNE' then V:=GS.Cells[SA_RefI,Lig] else
if St='E_LIBELLE' then V:=GS.Cells[SA_Lib,Lig] else
if St='SOLDE' then V:='SOLDE' else
{Zones Général}
if Copy(St,1,2)='G_' then
   BEGIN
   if CGen=Nil then Exit ; System.Delete(St,1,2) ;
   if St='GENERAL' then V:=CGen.General else
   if St='LIBELLE' then V:=CGen.Libelle else
   if St='ABREGE' then V:=CGen.Abrege else
      V:=RechercheLente('G_'+St,CGen.General) ;
   END else
{Zones Auxiliaire}
if Copy(St,1,2)='T_' then
   BEGIN
   if CAux=Nil then Exit ;  System.Delete(St,1,2) ;
   if St='AUXILIAIRE' then V:=CAux.Auxi else
   if St='LIBELLE' then V:=CAux.Libelle else
   if St='ABREGE' then V:=CAux.Abrege else
      V:=RechercheLente('T_'+St,CAux.Auxi) ;
   END else
{Zones Journal}
if Copy(St,1,2)='J_' then
   BEGIN
   if St='J_JOURNAL' then V:=SAJAL.Journal else
   if St='J_LIBELLE' then V:=SAJAL.LibelleJournal else
   if St='J_ABREGE' then V:=SAJAL.AbregeJournal else
      V:=RechercheLente('J_'+St,SAJAL.JOURNAL) ;
   END else
{Comptes auto}
if Copy(St,1,4)='AUTO' then V:=TrouveAuto(SAJAL.COMPTEAUTOMAT,Ord(St[Length(St)])-48) else
 if St='INTITULE' then
   BEGIN
   if CAux<>Nil then V:=CAux.Libelle else if CGen<>Nil then V:=CGen.Libelle ;
   END else
{Zones Ecriture}
   BEGIN
   if OBM=Nil then Exit ;
   if Copy(St,1,2)='E_' then System.Delete(St,1,2) ; if St='REFERENCE' then St:='REFINTERNE' ;
   V:=OBM.GetMvt('E_'+St) ;
   // Point 90 FFF
   if (VarToStr(V) = '') and (St = 'NUMENCADECA') then V := M.NumEncaDeca
   END ;
Load_Sais:=V ;
END ;


Function TrouveNextLigneAZero(GS : thGrid) : Integer ;
Var i : Integer ;
BEGIN
Result:=0 ;
if (ctxPCL in V_PGI.PGIContexte) then // FQ 15571 : SBO 20/04/2005 Anciennement vh^.CPIFDEFCEGID
  Exit ;
for i:=1 to GS.RowCount-2 do
  BEGIN
  If ((ValD(GS,i)=0) and (ValC(GS,i)=0)) then
    BEGIN
    if GS.RowCount>3 then BEGIN Result:=i ; Break ; END ;
    END ;
  END ;
END ;

procedure TFSaisie.TraiteLigneAZero ;
Var LigneAZero : Boolean ;
    LigNext : Integer ;
    i : Integer ;
BEGIN
if (ctxPCL in V_PGI.PGIContexte) then // FQ 15571 : SBO 20/04/2005 Anciennement vh^.CPIFDEFCEGID
  Exit ;
If Not GuideAEnregistrer Then Exit ; GuideAEnregistrer:=FALSE ;
LigneAZero:=FALSE ;
for i:=1 to GS.RowCount-2 do
  BEGIN
  If ((ValD(GS,i)=0) and (ValC(GS,i)=0)) then LigneAZero:=TRUE ;
  END ;
If LigneAZero And (HPiece.Execute(51,Caption,'')=mrYes) Then
  BEGIN
  Repeat
   LigNext:=TrouveNextLigneAZero(GS) ;
   if ((LigNext>0) and (LigNext<GS.RowCount-1)) then DetruitLigne(LigNext,TRUE) ;
  Until ((LigNext<=0) or (LigNext>=GS.RowCount)) ;
  END ;
END ;

procedure TFSaisie.FinGuide ;
Var Cancel : boolean ;
    i : integer ;
BEGIN
if Guide then CalculDebitCredit ;
for i:=1 to GS.RowCount-2 do
    BEGIN
    Cancel:=False ;
    GSRowExit(Nil,i,Cancel,True) ;
    END ;
Guide:=False ; FlashGuide.Flashing:=False ; FlashGuide.Visible:=False ;
END ;

Function TFSaisie.CommeUneBQE ( Lig : integer ) : Boolean ;
Var O : TOBM ;
BEGIN
Result:=False ;
if Not OkJalEffet then Exit ;
if SAJAL=Nil then Exit ;
if GS.Cells[Sa_Gen,Lig]='' then Exit ;
if GS.Cells[Sa_Gen,Lig]=SAJAL.Treso.Cpt then Exit ;
O:=GetO(GS,Lig) ; if O=Nil then Exit ;
Result:=True ;
END ;

function TFSaisie.ZoneObli ( Col,Lig : integer ; St : String ) : Boolean ;
Var CGen : TGGeneral ;
BEGIN
ZoneObli:=True ;
if Col=SA_Gen then
   BEGIN
   if ((Not ArretZone(SA_Aux,St,False)) and (Not Presence('GENERAUX','G_GENERAL',GS.Cells[SA_Gen,Lig]))) then Exit ;
   END ;
if Col=SA_Aux then
   BEGIN
   if Not Presence('TIERS','T_AUXILIAIRE',GS.Cells[SA_Aux,Lig]) then
      BEGIN
      CGen:=GetGGeneral(GS,Lig) ; if GS.Cells[SA_Gen,Lig]='' then Exit ;
      if CGen=Nil then Exit ; if CGen.Collectif then Exit ;
      END ;
   END ;
if Col=SA_Debit then
   BEGIN
if not vh^.CPIFDEFCEGID then // verif ok
   if ((TrouveSens(GS,E_NATUREPIECE.Value,Lig,CommeUneBQE(Lig))=1) and (GS.Cells[SA_Credit,Lig]='') and
       (GS.Cells[SA_Debit,Lig]='') and (Not ArretZone(SA_Credit,St,False))) then Exit ;
   END ;
if Col=SA_Credit then
   BEGIN
if not vh^.CPIFDEFCEGID then // verif ok
   if ((TrouveSens(GS,E_NATUREPIECE.Value,Lig,CommeUneBQE(Lig))=2) and (GS.Cells[SA_Debit,Lig]='') and
       (GS.Cells[SA_Credit,Lig]='') and (Not ArretZone(SA_Debit,St,False))) then Exit ;
   END ;
ZoneObli:=False ;
END ;

function TFSaisie.ProchainArret ( Col,Lig : integer ) : integer ;
Var OBM : TOBM ;
    St  : String ;
BEGIN
Result:=-1 ; if Not Guide then Exit ; if Lig>TGUI.Count then Exit ;
OBM:=TOBM(TGUI[Lig-1]) ; if OBM=Nil then BEGIN FinGuide ; Result:=-1 ; Exit ; END ;

  // Arrêt programmé dans les paramètres sociétés ou dans le scénario
  if (Action=taCreat) and (E_DATECOMPTABLE.CanFocus) and (not RentreDate) and
        ( (VH^.CPDateObli) or
          ( (OkScenario) and (Scenario<>Nil) and (Scenario.GetMvt('SC_DATEOBLIGEE')='X') )
        ) then
    begin
    Result := SA_DateC ;
    exit ;
    end ;  // SBO FQ 14730

St:=OBM.GetMvt('EG_ARRET') ;
for Col:=SA_Gen to SA_Credit do
    BEGIN
    if ArretZone(Col,St,True) then BEGIN OBM.PutMvt('EG_ARRET',St) ; Result:=Col ; Exit ; END ;
    if ZoneObli(Col,Lig,St) then BEGIN Result:=Col ; Exit ; END ;
    END ;
END ;

procedure TFSaisie.ValideCeQuePeux( Lig : longint ) ;
Var Cpte : String ;
    OBM  : TOBM ;
BEGIN
OBM:=GetO(GS,Lig) ; if OBM=Nil then Exit ;
if ((GS.Cells[SA_Gen,Lig]<>'') and (GS.Objects[SA_Gen,Lig]=Nil)) then
   BEGIN
   Cpte:=uppercase(GS.Cells[SA_Gen,Lig]) ;
   if Presence('GENERAUX','G_GENERAL',BourreLaDonc(Cpte,fbGene)) then
      BEGIN
      ChercheGen(SA_Gen,Lig,False) ; OBM.PutMvt('E_GENERAL',GS.Cells[SA_Gen,Lig]) ;
      END ;
   END ;
if ((GS.Cells[SA_Aux,Lig]<>'') and (GS.Objects[SA_Aux,Lig]=Nil)) then
   BEGIN
   Cpte:=uppercase(GS.Cells[SA_Aux,Lig]) ;
   if Presence('TIERS','T_AUXILIAIRE',BourreLaDonc(Cpte,fbAux)) then
      BEGIN
      ChercheAux(SA_Aux,Lig,False) ; OBM.PutMvt('E_AUXILIAIRE',GS.Cells[SA_Aux,Lig]) ;
      END ;
   END ;
{JP 06/10/05 : Déplacé dans RecupTronc dans le cadre de la FQ 13250
OBM.PutMvt('E_REFINTERNE',Copy(GS.Cells[SA_RefI,Lig],1,35)) ;
OBM.PutMvt('E_LIBELLE',Copy(GS.Cells[SA_Lib,Lig],1,35)) ;}
if GS.Cells[SA_Debit,Lig]<>'' then ChercheMontant(SA_Debit,Lig) ;
if GS.Cells[SA_Credit,Lig]<>'' then ChercheMontant(SA_Credit,Lig) ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 18/03/2003
Modifié le ... :   /  /    
Description .. : - LG - 18/03/2003 - reecriture de la fct pour n'appeler 
Suite ........ : finGuide qu'une seule fois
Mots clefs ... : 
*****************************************************************}
procedure TFSaisie.SuiteLigneColonne ( Var Col,Lig : longint ) ;
Var CN: integer ;

  procedure _SuiteLigneColonne ( Var Col,Lig : longint ) ;
   begin
    OkSuite:=False ;
    if Not Guide then Exit ;
     if Lig>NbLG then BEGIN FinGuide ; Exit ; END ;
    ValideCeQuePeux(Lig) ; CN:=ProchainArret(Col,Lig) ;
    if CN>0 then
     BEGIN OkSuite:=True ; Col:=CN ; END
      else
       BEGIN
       if LigneCorrecte(Lig,False,True) then
          BEGIN
          //Cancel:=False ; Chg:=True ;
          Col:=SA_Gen ; Lig:=Lig+1 ;
          _SuiteLigneColonne(Col,Lig) ;
          END ;
        //if Lig>NbLG then FinGuide ;
       END ;
    end ;

BEGIN
 GS.BeginUpdate ;
 try
  _SuiteLigneColonne ( Col,Lig ) ;
  if (CN<0) and (Lig>NbLG) then FinGuide ;
 finally
  GS.EndUpdate ;
 end ;
END ;   


{==========================================================================================}
{============================ Objet MOUVEMENT, Divers =====================================}
{==========================================================================================}
Function TFSaisie.FabricFromM : RMVT ;
Var X : RMVT ;
BEGIN
X.Axe:='' ; X.Etabl:=E_ETABLISSEMENT.Value ;
X.Jal:=E_JOURNAL.Value ; X.Exo:=QuelExoDT(StrToDate(E_DATECOMPTABLE.Text)) ;
X.CodeD:=E_DEVISE.Value ; X.Simul:=M.Simul ; X.Nature:=E_NATUREPIECE.Value ;
X.DateC:=StrToDate(E_DATECOMPTABLE.Text) ; X.DateTaux:=DEV.DateTaux ;
X.Num:=NumPieceInt ; X.TauxD:=DEV.Taux ; X.Valide:=False ; X.ANouveau:=M.ANouveau ;
Result:=X ;
END ;

procedure TFSaisie.StockeLaPiece ;
Var X : RMVT ;
    P : P_MV ;
BEGIN
if ((Action<>taCreat) and (Not SuiviMP)) then Exit ; FillChar(X,Sizeof(X),#0) ;
X:=FabricFromM ; P:=P_MV.Create ; P.R:=X ; TPIECE.Add(P) ;
END ;

procedure TFSaisie.NumeroteVentils ;
Var OBM	 : TOBM;
		OBA	 : TOB ;
    NumL,NumAxe,NumV : integer ;
BEGIN
	for NumL:=1 to GS.RowCount-2 do
    BEGIN
    OBM := GetO(GS,NumL) ;
    if OBM=Nil then Continue ;
    for NumAxe:=0 to OBM.Detail.Count-1 do
    	if OBM.Detail[NumAxe].Detail.Count>0 then
        BEGIN
        for NumV:=0 to OBM.Detail[NumAxe].Detail.Count-1 do
        	begin
          OBA := OBM.Detail[NumAxe].Detail[NumV];
        	OBA.PutValue('Y_NUMLIGNE',NumL) ;
          end;
        END ;
    END ;
END ;

procedure TFSaisie.GotoEntete ;
Var i : integer ;
    T : TWinControl ;
BEGIN
if PasModif then Exit ; if Not PEntete.Enabled then Exit ;
for i:=0 to PEntete.ControlCount-1 do if PEntete.Controls[i] is TWinControl then
    BEGIN
    T:=TWinControl(PEntete.Controls[i]) ;
    if ((T.CanFocus) and (Copy(T.Name,1,2)='E_')) then BEGIN T.SetFocus ; Break ; END ;
    END ;
END ;

procedure TFSaisie.GereComplements(Lig : integer) ;
{$IFNDEF GCGC}
Var O      : TOBM ;
    CGen   : TGGeneral ;
    ModBN  : boolean ;
    StComp,StLibre : String ;
    RC             : R_COMP ;
    NumRad         : integer ;
{$ENDIF}
BEGIN
{$IFNDEF GCGC}
if PasModif then Exit ;
O:=GetO(GS,Lig) ; if O=Nil then Exit ; if O.CompS then Exit ;
CGen:=GetGGeneral(GS,Lig) ; if CGen=Nil then Exit ;
NumRad:=ComporteLigne(Scenario,CGen.General,StComp,StLibre) ; if NumRad<=0 then Exit ;
if ((Pos('X',StComp)<=0) and (Pos('X',StLibre)<=0)) then Exit ;
RC.StComporte:=StComp ; RC.StLibre:=StLibre ; RC.Conso:=False ; RC.DateC:=DatePiece ;
RC.Attributs:=(Action=taCreat) ; RC.MemoComp:=MemoComp ; RC.Origine:=NumRad ;
O.CompS:=True ; if Not SaisieComplement(O,EcrGen,Action,ModBN,RC) then Exit ;
if GS.CanFocus then GS.SetFocus ;
if ModBN then Revision:=False ;
{$ENDIF}
END ;

procedure TFSaisie.DecrementeNumero ;
Var Facturier : String3 ;
    DD : TDateTime ;
BEGIN
if ( M.Simul <> 'N' ) and ( M.Simul <> 'I' ) // Modif IFRS
  then Facturier:=SAJAL.CompteurSimul
  else Facturier:=SAJAL.CompteurNormal ;
DD:=StrToDate(E_DATECOMPTABLE.Text) ;
SetDecNum(EcrGen,Facturier,DD) ;
END ;

Procedure TFSaisie.EditionSaisie ;
Var SWhere : String ;
    i      : integer ;
    RR     : RMVT ;
BEGIN
if Not NeedEdition then Exit ;
if M.MajDirecte then Exit ;
If M.MSED.MultiSessionEncaDeca Then Exit ;
if EtatSaisie='' then Exit ;
if HDiv.Execute(11,Caption,'')<>mrYes then Exit ;
NeedEdition:=False ; SWhere:='' ;
for i:=0 to TPIECE.Count-1 do
    BEGIN
    RR:=P_MV(TPIECE[i]).R ;
    SWhere:=SWhere+'(E_JOURNAL="'+RR.Jal+'" AND E_NUMEROPIECE='+IntToStr(RR.Num)+' AND E_QUALIFPIECE="'+RR.Simul+'" AND E_EXERCICE="'+RR.Exo+'")' ;
    if i<TPIECE.Count-1 then SWhere:=SWhere+' OR ' ;
    END ;
if TPIECE.Count>1 then SWhere:='('+SWhere+')' ;
{$IFNDEF IMP}
{$IFNDEF GCGC}
LanceEtat('E','SAI',EtatSaisie,True,False,False,Nil,SWhere,'',False) ;
{$ENDIF}
{$ENDIF}
END ;

Procedure TFSaisie.DetruitPieceMP ;
Var XX : RMVT ;
    i  : integer ;
    TOBL : TOB ;
BEGIN
XX:=M ; XX.Simul:='S' ;
ExecuteSQL('DELETE FROM ECRITURE WHERE '+WhereEcriture(tsGene,XX,False)) ;
ExecuteSQL('DELETE FROM ANALYTIQ WHERE '+WhereEcriture(tsAnal,XX,False)) ;
if TOBMPEsc<>Nil then
   BEGIN
   for i:=0 to TOBMPEsc.Detail.Count-1 do
       BEGIN
       TOBL:=TOBMPEsc.Detail[i] ;
       XX:=TOBToIdent(TOBL,True) ; XX.Simul:='S' ;
       ExecuteSQL('DELETE FROM ECRITURE WHERE '+WhereEcriture(tsGene,XX,False)) ;
       ExecuteSQL('DELETE FROM ANALYTIQ WHERE '+WhereEcriture(tsAnal,XX,False)) ;
       END ;
   END ;
END ;

Function TFSaisie.FermerSaisie : boolean  ;
Var i : integer ;
    Okok : boolean ;
BEGIN
  {JP 05/07/07 : FQ 19022 : Si on répond non à la confirmation fermeture}
  FNeFermePlus := False;

Result:=True ;
if ((Action<>taConsult) and (ModeForce<>tmNormal)) then BEGIN Result:=False ; Exit ; END ;
if ((NeedEdition) and (Action=taCreat) and (TPIECE.Count>0) and
    (Not M.MajDirecte) And (Not M.MSED.MultiSessionEncaDeca)) then EditionSaisie ;
if ((Action=taConsult) or (Not PieceModifiee) or (M.smp=smpEncTraEdtNC) Or (M.smp=smpDecBorEdtNC) Or (M.smp=smpDecChqEdtNC) Or (M.smp=smpDecVirEdtNC) Or (M.smp=smpDecVirInEdtNC) Or (M.smp=smpEncPreEdtNC) or
     (M.smp in [smpCompenFou, smpCompenCli])) then   {FP 21/02/2006}
   BEGIN
   if SuiviMP then DetruitPieceMP ;
   Exit ;
   END ;
if ((Action=taCreat) and (M.TypeGuide='NOR') and (M.FromGuide)) then Exit ;
if Action=taModif then
   BEGIN
     if HPiece.Execute(3,caption,'')<>mrYes then begin
        Result:=False;
        {JP 05/07/07 : FQ 19022 : Si on répond non à la confirmation fermeture}
        FNeFermePlus := True;
     end
     else if SuiviMP then DetruitPieceMP ;
   END else
   BEGIN
   if PossibleRecupNum then
      BEGIN
      Okok:=True ; i:=HPiece.Execute(3,caption,'') ;
      {JP 05/07/07 : FQ 19022 : Si on répond non à la confirmation fermeture}
      FNeFermePlus := i <> mrYes;

      if i<>mrYes then begin Result:=False ; Exit ; END ;
      END else Okok:=False ;
   if ((Okok) and (PossibleRecupNum)) then {Revérifier, le numéro a pu être attribué entre temps!}
      BEGIN
      if H_NUMEROPIECE_.Visible=False then
         BEGIN
         if Transactions(DecrementeNumero,10)<>oeOk then MessageAlerte(HDiv.Mess[6]) else M.LastNumCreat:=-1 ;
         END ;
      END else
      BEGIN
      HPiece.Execute(12,caption,'') ; Result:=False ;
      END ;
   END ;
END ;

procedure TFSaisie.AlimObjetMvt ( Lig : integer ; AvecM : boolean ) ;
Var OBM   : TOBM ;
    SD,SC : Double ;
    CGen  : TGGeneral ;
    LaDate      : TDateTime ;
    St : String ;
BEGIN
if PasModif then Exit ;
OBM:=GetO(GS,Lig) ; if OBM=Nil then BEGIN AlloueMvt(GS,EcrGen,Lig,True) ; OBM:=GetO(GS,Lig) ; END ;
LaDate:=StrToDate(E_DateComptable.Text) ;
OBM.PutMvt('E_EXERCICE',QuelExoDT(LaDate)) ;
OBM.PutMvt('E_JOURNAL',SAJAL.JOURNAL)           ; OBM.PutMvt('E_DATECOMPTABLE',LaDate) ;
{$IFNDEF SPEC302}
OBM.PutMvt('E_PERIODE',GetPeriode(LaDate))      ; OBM.PutMvt('E_SEMAINE',NumSemaine(LaDate)) ;
{$ENDIF}
OBM.PutMvt('E_NUMEROPIECE',NumPieceInt)         ; OBM.PutMvt('E_NUMLIGNE',Lig) ;
OBM.PutMvt('E_GENERAL',GS.Cells[SA_Gen,Lig])    ; OBM.PutMvt('E_AUXILIAIRE',GS.Cells[SA_Aux,Lig]) ;
OBM.PutMvt('E_QUALIFPIECE',M.Simul)             ; OBM.PutMvt('E_DEVISE',DEV.Code) ;
OBM.PutMvt('E_TAUXDEV',DEV.Taux)                ; OBM.PutMvt('E_DATETAUXDEV',DEV.DateTaux) ;
OBM.PutMvt('E_NATUREPIECE',E_NATUREPIECE.Value) ; OBM.PutMvt('E_ETABLISSEMENT',E_ETABLISSEMENT.Value) ;
OBM.PutMvt('E_CONTROLETVA','RIE') ;
OBM.PutMvt('E_LIBELLE',Copy(GS.Cells[SA_Lib,Lig],1,35)) ;
OBM.PutMvt('E_REFINTERNE',Copy(GS.Cells[SA_RefI,Lig],1,35)) ;
OBM.PutMvt('E_MODESAISIE','-') ;
OBM.MajLesDates(Action) ;
if AvecM then
   BEGIN
   SD:=ValD(GS,Lig) ; SC:=ValC(GS,Lig) ;
   OBM.SetMontants(SD,SC,DEV,True) ;
   END ;
CGen:=GetGGeneral(GS,Lig) ;
if CGen<>Nil then
   BEGIN
   if Ventilable(CGen,0) Then OBM.PutMvt('E_ANA','X') Else OBM.PutMvt('E_ANA','-') ;
   St:=OBM.GetMvt('E_QUALIFQTE1') ; if ((St='') or (St='...')) then OBM.PutMvt('E_QUALIFQTE1',CGen.QQ1) ;
   St:=OBM.GetMvt('E_QUALIFQTE2') ; if ((St='') or (St='...')) then OBM.PutMvt('E_QUALIFQTE2',CGen.QQ2) ;
   END ;
END ;

{==========================================================================================}
{================================ Génération, Calcul TVA ==================================}
{==========================================================================================}
procedure TFSaisie.RemplirEcheAuto ( Lig : integer ) ;
BEGIN
GereEche(Lig,False,True) ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Compta
Créé le ...... : 01/01/1900
Modifié le ... : 21/07/2006
Description .. : 
Suite ........ : SBO 02/02/2005 : FQ 15350 L'affectation automatique de
Suite ........ : l'analytique ne fonctionnait pas : les tob "Analytiq" n'étaient
Suite ........ : pas rattaché à l'arborescence :
Suite ........ : TOB Ecriture > Tob Axe > Tob Analytique
Suite ........ : 
Suite ........ : SBO 30/06/2006 : FQ18480  : utilisation de la fonction 
Suite ........ : standard pour gestion du CroiseAxe
Mots clefs ... : 
*****************************************************************}
procedure TFSaisie.RemplirAnalAuto ( Lig : integer ) ;
Var CGen   : TGGeneral ;
    OBM    : TOBM ;
begin

  CGen := GetGGeneral(GS,Lig) ;
  if not Ventilable( CGen, 0 ) then Exit ;

  OBM	:= GetO(GS,Lig) ;
  if OBM=Nil then Exit ;

  if OBM.Detail.Count<MaxAxe then
    AlloueAxe(OBM);

  VentilerTob( OBM, '', 0, DecDev, True ) ;
  OkMessAnal:=True ;

end ;

function TFSaisie.InsereLigneTVA ( LigHT : integer ; RegTVA : String3 ; SoumisTpf,Achat : boolean ) : boolean ;
Var CGen,TvaGen,TpfGen : TGGeneral ;
    XHT,XTVA,XTPF   : Double ;
    Sens,LigA,LigF  : integer ;
    OBM             : TOBM ;
    CodeTva,CodeTpf : String3 ;
BEGIN
Result:=True ;
Sens:=1 ; XHT:=ValD(GS,LigHT) ; if XHT=0 then BEGIN Sens:=2 ; XHT:=ValC(GS,LigHT) ; END ;
CGen:=GetGGeneral(GS,LigHT) ;
OBM:=GetO(GS,LigHT) ; CodeTva:=OBM.GetMvt('E_TVA') ; CodeTpf:=OBM.GetMvt('E_TPF') ;
if ((SoumisTpf) and (CGen.TPF<>'') and (CGen.CpteTPF<>'') and (CGen.TauxTPF<>0)) then
   BEGIN
   XTPF:=HT2TPF(XHT,RegTVA,SoumisTPF,CodeTVA,CodeTPF,Achat,DECDEV) ;
   if XTPF<>0 then
      BEGIN
      LigF:=GS.RowCount-1 ; DefautLigne(LigF,True) ;
      GS.Cells[SA_Gen,LigF]:=CGen.CpteTPF ; ChercheGen(SA_Gen,LigF,False) ;
      if Sens=1 then GS.Cells[SA_Debit,LigF]:=StrS(XTPF,DECDEV) else GS.Cells[SA_Credit,LigF]:=StrS(XTPF,DECDEV) ;
      FormatMontant(GS,SA_Debit,LigF,DECDEV) ; FormatMontant(GS,SA_Credit,LigF,DECDEV) ; TraiteMontant(LigF,False) ;
      AlimObjetMvt(LigF,True) ; GereNewLigne(GS) ;
      TpfGen:=GetGGeneral(GS,LigF) ;
      if Lettrable(TpfGen)<>NonEche then RemplirEcheAuto(LigF) else
         if Ventilable(TpfGen,0) then RemplirAnalAuto(LigF) ;
      OBM:=GetO(GS,LigF) ; OBM.PutMvt('E_REGIMETVA',RegTVA) ; OBM.PutMvt('E_TVA',CodeTVA) ; OBM.PutMvt('E_TPF',CodeTPF) ;
      END ;
   END else if ((SoumisTPF) and (CGen.TauxTpf<>0)) then Result:=False ;
if ((CGen.TVA<>'') and (CGen.CpteTVA<>'') and (CGen.TauxTVA<>0)) then
   BEGIN
   XTVA:=HT2TVA(XHT,RegTVA,SoumisTPF,CodeTVA,CodeTPF,Achat,DECDEV) ;
   if XTVA<>0 then
      BEGIN
      LigA:=GS.RowCount-1 ; DefautLigne(LigA,True) ;
      GS.Cells[SA_Gen,LigA]:=CGen.CpteTVA ; ChercheGen(SA_Gen,LigA,False) ;
      if Sens=1 then GS.Cells[SA_Debit,LigA]:=StrS(XTVA,DECDEV) else GS.Cells[SA_Credit,LigA]:=StrS(XTVA,DECDEV) ;
      FormatMontant(GS,SA_Debit,LigA,DECDEV) ; FormatMontant(GS,SA_Credit,LigA,DECDEV) ; TraiteMontant(LigA,False) ;
      AlimObjetMvt(LigA,True) ; GereNewLigne(GS) ;
      TvaGen:=GetGGeneral(GS,LigA) ;
      if Lettrable(TvaGen)<>NonEche then RemplirEcheAuto(LigA) else
         if Ventilable(TvaGen,0) then RemplirAnalAuto(LigA) ;
      OBM:=GetO(GS,LigA) ; OBM.PutMvt('E_REGIMETVA',RegTVA) ; OBM.PutMvt('E_TVA',CodeTVA) ; OBM.PutMvt('E_TPF',CodeTPF) ;
      END ;
   END else if CGen.TauxTva<>0 then Result:=False ;
END ;

procedure TFSaisie.AlimHTByTVA ( Lig : integer ; RegTVA,TVAENC : String3 ; SoumisTpf,Achat : boolean ) ;
Var CGen : TGGeneral ;
    O    : TOBM ;
    EncON : boolean ;
    CodeTva,CodeTpf : String3 ;
BEGIN
CGen:=GetGGeneral(GS,Lig) ; if CGen=Nil then Exit ;
O:=GetO(GS,Lig) ; if O=Nil then Exit ;
CodeTva:=O.GetMvt('E_TVA') ; CodeTpf:=O.GetMvt('E_TPF') ;
EncON:=O.GetMvt('E_TVAENCAISSEMENT')='X' ;
RenseigneHTByTva(CGen,RegTVA,CodeTva,CodeTpf,SoumisTPF,Achat,EncON,CoherTva) ;
END ;

{==========================================================================================}
{================================ Methodes du HGrid =======================================}
{==========================================================================================}
procedure TFSaisie.GSDblClick(Sender: TObject);
begin
if ((Action<>taConsult) and (ModeForce=tmDevise)) then ForcerMode(False,VK_RETURN) else ClickZoom ;
end;

procedure TFSaisie.GSCellExit(Sender: TObject; var ACol,ARow : Longint ; var Cancel : Boolean );
begin
if PasModif then Exit ;
if GS.Row=ARow then
   BEGIN
   // Généraux
   if ACol=SA_Gen then
      BEGIN
      if ((GS.Cells[ACol,ARow]='') and ((GS.Cells[SA_Aux,ARow]='') or (GS.Objects[SA_Aux,ARow]=Nil))) then Exit ;
      {JP 20/02/04 :
        Ancien Code : if LeMeme(GS,ACol,ARow) then Exit;
        Le problème est que si l'on passe par F5, ChercheGen(ACol,ARow,False) a déjà été exécuté, et
        donc l'objet derrière la cellule en cours a déjà été modifié => LeMeme renvoie True => on sort
        et la variable UnChange reste à False et dans ClickValide, l'appel à LigneCorrecte se fait avec
        un paramétre à False}
      if GS.Cells[ACol,ARow] = StCellCur then Exit;

      if ((Not Guide) and (ARow<=1) and (GS.Cells[ACol,ARow]<>'')) then EtudieSiGuide(ACol,ARow) ;
      if ChercheGen(ACol,ARow,False)<=0 then BEGIN Cancel:=True ; Exit ; END ;
      if Not Cancel then CpteAuto:=False ;
       GereCutoff(ARow) ;
      END
   else
    // Auxiliaire
     if ACol=SA_Aux then
      BEGIN
      if ((GS.Cells[ACol,ARow]='') and (GS.Col=SA_Gen)) then Exit ;
      if LeMeme(GS,ACol,ARow) then Exit ;
      if ChercheAux(Acol,ARow,False)<=0 then BEGIN Cancel:=True ; Exit ; END ;
      END ;
   END ;
if ((ACol=SA_Debit) or (ACol=SA_Credit)) then
  begin
  // DEV 3216 10/04/2006 SBO : récupération des fonctionnalités de saisBor.pas
  if (pos('+',GS.Cells[ACol, ARow])>0) or (pos('-',GS.Cells[ACol, ARow])>0) or (pos('/',GS.Cells[ACol, ARow])>0) or (pos('*',GS.Cells[ACol, ARow])>0) then
    begin
    if (not VH^.MontantNegatif) and ( pos( Copy(GS.Cells[ACol, ARow], 1, 1), '+;-;/;*' ) > 0 ) then
      begin
      if ACol = SA_DEBIT
        then GS.Cells[ACol, ARow]:=GFormule('{[E_DEBIT]'+GS.Cells[ACol, ARow]+'}',  GetFormulePourCalc, nil, 1)
        else GS.Cells[ACol, ARow]:=GFormule('{[E_CREDIT]'+GS.Cells[ACol, ARow]+'}', GetFormulePourCalc, nil, 1);
      end
    else
      // sinon on effectue l'opération simplement
      begin
      if Pos(GS.Cells[ACol, ARow][1],'/')>-1
          then GS.Cells[ACol, ARow]:='{' + GetFormatPourCalc + GS.Cells[ACol, ARow]+'}'; // FQ18371 et FQ18080
      GS.Cells[ACol, ARow] := GFormule( '{' + GS.Cells[ACol, ARow] + '}', GetFormulePourCalc, nil, 1) ;
      end ;
   end;
   FormatMontant(GS, ACol, ARow, DEV.Decimale) ;
  // FIN DEV 3216 10/04/2006 SBO
  ChercheMontant(Acol,ARow) ;
  end ;
if Guide then RecalcGuide ;
if (VH^.PaysLocalisation=CodeISOES) then  //XVI 24/02/2005
   GereEnabled(ARow) ;
if ((Not Cancel) and (GS.Cells[ACol,ARow]<>StCellCur)) then UnChange:=True ;
end;

procedure TFSaisie.GereOptionsGrid ( Lig : integer ) ;
Var Okok : boolean ;
BEGIN
Okok:=(GoRowSelect in GS.Options) ;
if ModeLC then
   BEGIN
   GS.CacheEdit ;
   GS.Options:=GS.Options-[GoEditing,GoTabs,GoAlwaysShowEditor] ;
   GS.Options:=GS.Options+[GoRowSelect] ;
   END else
   BEGIN
   if ((Action=taModif) and (ModeForce=tmNormal)) then
      BEGIN
      if PasToucheLigne(Lig) then
         BEGIN
         GS.CacheEdit ;
         GS.Options:=GS.Options-[GoEditing,GoTabs,GoAlwaysShowEditor] ;
         GS.Options:=GS.Options+[GoRowSelect] ;
         G_LIBELLE.Font.Color:=clRed ;
         T_LIBELLE.Font.Color:=clRed ;
         G_LIBELLE.Caption:=HDiv.Mess[7] ;
         if GetO(GS,Lig).GetMvt('E_REFPOINTAGE')<>''
           then T_LIBELLE.Caption:=' ( lettré, pointé ou compte ou journal fermé ). [' + GetO(GS,Lig).GetMvt('E_REFPOINTAGE') + ']' // Modif BPY 08/08/2003
           else T_LIBELLE.Caption:= ' ( lettré, pointé, compte ou journal fermé ).' ;
         END else
         BEGIN
         GS.Options:=GS.Options-[GoRowSelect] ;
         GS.Options:=GS.Options+[GoEditing,GoTabs,GoAlwaysShowEditor] ;
         GS.MontreEdit ;
         END ;
      END ;
   END ;
if Okok<>(GoRowSelect in GS.Options) then GS.Refresh ;
END ;

procedure TFSaisie.GSCellEnter(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
Var CGen : TGGeneral ;
begin
if ARow<>GS.Row then GereOptionsGrid(GS.Row) ;
if ((Action<>taConsult) and (Not EstRempli(GS,GS.Row)) and (GS.Col<>SA_Gen) and (GS.Col<>SA_Aux))
   then BEGIN ACol:=SA_Gen ; ARow:=GS.Row ; Cancel:=True ; Exit ; END ;

// Essai Pb FFF validation de pièce SBO 29/16/2006
if ((H_NUMEROPIECE_.Visible) and (GS.Row>1)) then
   BEGIN
{$IFNDEF FFF}
   if Transactions(AttribNumeroDef,10)<>oeOK then MessageAlerte(HDiv.Mess[5]) ;
{$ENDIF FFF}
   BlocageEntete(Self,FALSE) ;
   END ;
{$IFDEF FFF}
BValide.Enabled := (GS.RowCount>2) and Equilibre; //((H_NumeroPiece.Visible) and (Equilibre)) ;
{$ELSE FFF}
BValide.Enabled := ((H_NumeroPiece.Visible) and (Equilibre)) ;
{$ENDIF FFF}
// Fin Essai Pb FFF validation de pièce SBO 29/16/2006

GereNewLigne(GS) ; GereZoom ;
if PasModif then exit ;
if GS.Col=SA_Aux then
   BEGIN
   // Shunt de la zone aux si compte pas collectif
   CGen:=GetGGeneral(GS,GS.Row) ;
   if ((CGen<>Nil) and (GS.Cells[SA_Gen,GS.Row]<>'') and (Not CGen.Collectif)) then
      BEGIN
      if ACol=SA_Gen then ACol:=SA_RefI else ACol:=SA_Gen ;
      ARow:=GS.Row ; Cancel:=TRUE ;
      END ;
   END else if GS.Col=SA_RefI then
   BEGIN
//   if AutoCharged then RenseigneLibelleAuto(GS.Col,GS.Row,True) ;
   END else if GS.Col=SA_Lib then
   BEGIN
//   if AutoCharged then RenseigneLibelleAuto(GS.Col,GS.Row,False) ;
   END else if GS.Col=SA_Credit then
   BEGIN
   if GS.Cells[SA_Debit,GS.Row]<>'' then
      BEGIN
      if GS.Row=ARow then
         BEGIN
         if ACol=SA_Debit then BEGIN ARow:=ARow+1 ; ACol:=SA_Gen ; END else BEGIN ACol:=SA_Debit ; ARow:=GS.Row ; END ;
         END else
         BEGIN
         ACol:=SA_Debit ; ARow:=GS.Row ;
         END ;
      Cancel:=True ;
      END ;
   END else if GS.Col=SA_Debit then
   BEGIN
   if GS.Cells[SA_Credit,GS.Row]<>'' then
      BEGIN
      if GS.Row=ARow then
         BEGIN
         if ACol<=SA_Lib then ACol:=SA_Credit else if ACol>=SA_Credit then ACol:=SA_Lib ;
         END else
         BEGIN
         ACol:=SA_Credit ; ARow:=GS.Row ;
         END ;
      Cancel:=TRUE ;
      END else
      BEGIN
      if ((GS.Row=ARow) and (TrouveSens(GS,E_NATUREPIECE.Value,GS.Row,CommeUneBQE(GS.Row))=2) and (Acol<GS.Col) and (GS.Cells[SA_Debit,GS.Row]=''))
         then BEGIN ACol:=SA_Credit ; Cancel:=True ; END ;
      END ;
   END ;
if ((Guide) and (Not Cancel)) then
    BEGIN
    if Not OkSuite then
       BEGIN
       ACol:=GS.Col ; ARow:=GS.Row ; SuiteLigneColonne(ACol,ARow) ;
       if ((ACol<>GS.Col) or (ARow<>GS.Row)) then Cancel:=True ;
       END ;
    END ;
if Not Cancel then StCellCur:=GS.Cells[GS.Col,GS.Row] ;
end;

procedure TFSaisie.GSRowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
begin
if ((Not EstRempli(GS,Ou)) and (Not PasModif)) then DefautLigne(Ou,True) ;
CurLig:=Ou ; AffichePied ;
GereEnabled(Ou) ; GereOptionsGrid(Ou) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Piot
Créé le ...... : 23/02/2005
Modifié le ... : 09/09/2005
Description .. : 23/02/2005 - LG - Gestion de l'accelerateur de saisie
Suite ........ : - 09/09/2005 -  LG - FB 16363 - appel du cutoff en sortie de 
Suite ........ : ligne pour la saisie guide
Mots clefs ... : 
*****************************************************************}
procedure TFSaisie.GSRowExit(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
Var OkL  : Boolean ;
    lTobEcr : Tob ;
begin
if PasModif then Exit ;
if Ou>=GS.RowCount-1 then Exit ;
GridEna(GS,False) ;
OkL:=LigneCorrecte(Ou,False,(Action=taCreat) or (Not Revision) or (UnChange)) ;
if OkL then
   BEGIN {YMO FQ18125 13/02/2007 Pas de déclenchement des vérifications sur les lignes non modifiables}
   if not PasToucheLigne(Ou) then
     begin
     GereComplements(Ou) ;
     GereEche(Ou,True,False) ;
     GereAnal(Ou,True,True) ;
     GereTvaEncNonFact(Ou) ;
     GereLesImmos(Ou) ;
     GereAcc(Ou) ;
     GereCutoff(Ou) ;
     if TestQteAna then // DEV3946
       begin
       lTobEcr := Tob(GetO(GS,Ou)) ;
       if ( lTobEcr <> nil ) and ( ( lTobEcr.GetNumChamp('CHECKQTE') < 0 ) or
                                   ( lTobEcr.GetString('CHECKQTE')<>'X'  ) ) then
          CheckQuantite( lTobEcr );
       end ;
     end ;
   END
  else Cancel:=True ;

GridEna(GS,True) ;
end;

procedure TFSaisie.GSEnter(Sender: TObject);
Var StComp,StLibreEnt : String ;
    b,OkD  : boolean ;
{$IFNDEF GCGC}
    RC     : R_COMP ;
{$ENDIF}
    i      : integer ;
    O      : TOBM ;
begin
GereEnabled(GS.Row) ;
if Action<>taCreat then Exit ;
if vh^.CPIFDEFCEGID then  // verif ok
If MethodeGuideDate1 Then
  BEGIN
  if (appeldatepourguide And (Not M.FromGuide) And (((M.TypeGuide='NOR') and (M.SaisieGuidee)) Or (CodeGuideFromSaisie<>''))) then
    BEGIN
    If M.SaisieGuidee Then LanceGuide(M.TypeGuide,M.LeGuide) Else LanceGuide('NOR',CodeGuideFromSaisie) ;
    END ;
  END ;

if ((Not OkScenario) and (Not VH^.CPDateObli)) then Exit ;
if DejaRentre then Exit ;
if PasModif then Exit ; if Entete=Nil then Exit ;
if ((GS.RowCount>2) and (Not M.SaisieGuidee)) then Exit ;
if ((EstRempli(GS,1)) and (Not M.SaisieGuidee)) then Exit ;
{$IFNDEF SPEC302}
OkD:=False ;
if VH^.CPDateObli then OkD:=True else
   BEGIN
   if ((OkScenario) and (Scenario<>Nil)) then OkD:=(Scenario.GetMvt('SC_DATEOBLIGEE')='X') ;
   END ;
if ((Action=taCreat) and (Not EstRempli(GS,1)) and
   (Not M.FromGuide) and (Not M.SaisieGuidee) and
   (OkD) and (Not DejaRentre) and (Not RentreDate)) then
   BEGIN
   if E_DATECOMPTABLE.CanFocus then E_DATECOMPTABLE.SetFocus ; Exit ;
   END ;
{$ENDIF}
if vh^.CPIFDEFCEGID then  // verif ok
ResaisirLaDatePourLeGuide((Not M.FromGuide) And (M.TypeGuide='NOR') And M.SaisieGuidee And (Not RentreDate)) ; // a prtie de saisie guidée

RentreDate:=True ; DejaRentre:=True ;
{$IFNDEF GCGC}
if ((OkScenario) and (Entete<>Nil)) then
   BEGIN
   StComp:=Entete.GetMvt('E_ETAT') ; StLibreEnt:=Scenario.GetMvt('SC_LIBREENTETE') ;
   if ((Pos('X',StComp)<=0) and (Pos('X',StLibreEnt)<=0)) then Exit ;
   RC.StComporte:=StComp ; RC.StLibre:=StLibreEnt ; RC.Conso:=False ; RC.DateC:=DatePiece ;
   RC.Attributs:=(Action=taCreat) ; RC.MemoComp:=MemoComp ; RC.Origine:=0 ;
   SaisieComplement(Entete,EcrGen,Action,b,RC) ;
   DefautLigne(1,True) ; GereEnabled(1) ;
   for i:=2 to GS.RowCount-2 do
       BEGIN
       O:=GetO(GS,i) ; if O=Nil then Break ;
       PutScenar(O,i) ;
       END ;
   END ;
{$ENDIF}
end;

procedure TFSaisie.GSKeyPress(Sender: TObject; var Key: Char);
begin
if Not GS.SynEnabled then Key:=#0
else if Key=#127
     then Key:=#0
   else if ((Key='=') and ((GS.Col=SA_Debit) or (GS.Col=SA_Credit) or (GS.Col=SA_Gen) or (GS.Col=SA_Aux)))
      then Key:=#0
   else if Key=#17
     then Key:=#0 ;
OkSuite:=False ;

end;

procedure TFSaisie.GSExit(Sender: TObject);
Var b,bb : boolean ;
    C,R  : longint ;
begin
if PasModif then Exit ;
if Valide97.Tag<>1 then Exit ;
if Outils.Tag<>1 then Exit ;
C:=GS.Col ; R:=GS.Row ; b:=False ; bb:=False ;
GSCellExit(GS,C,R,b) ; GSRowExit(GS,R,b,bb) ;
if ((Action=taCreat) and (GS.RowCount<=3) and (R=1) and (EstRempli(GS,R))) then
   BEGIN
   if ((Screen.ActiveControl<>Nil) and (Screen.ActiveControl.Parent=PEntete)) then GS.SetFocus ;
   BlocageEntete(Self,False) ;
   END ;
end;

procedure TFSaisie.GSMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin GX:=X ; GY:=Y ; end;

procedure TFSaisie.POPSPopup(Sender: TObject);
begin InitPopUp(Self) ; end;

procedure TFSaisie.FindSaisFind(Sender: TObject);
Var bc : boolean ;
begin
Rechercher(GS,FindSais,FindFirst) ;
bc:=False ; GSRowEnter(Nil,GS.Row,bc,False) ;
end;

procedure TFSaisie.GSSetEditText(Sender: TObject; ACol, ARow: Longint; const Value: string);
begin PieceModifiee:=True ; end;

Function QOK(Q : tQuery ; Simul : String3)  : Boolean ;
Var St : String ;
BEGIN
Result:=TRUE ;
If (ctxPCL in V_PGI.PGIContexte) Then Exit ;
St:='' ;
If Q.FindField('E_JOURNAL')=NIL Then St:='E_JOURNAL' ;
If Q.FindField('E_DATECOMPTABLE')=NIL Then St:='E_DATECOMPTABLE' ;
If Q.FindField('E_NUMEROPIECE')=NIL Then St:='E_NUMEROPIECE' ;
If St<>'' Then
  BEGIN
  Result:=FALSE ;
  HShowMessage('0;Opération impossible !;Le champ '+ST+' doit faire partie des colonnes affichées;E;O;O;O;','','') ;
  HShowMessage('0;Opération impossible !;Vous devez le rajouter dans le paramétrage de cette liste;E;O;O;O;','','') ;
  END ;
END ;

Function TrouveSaisie(Q : TQuery ; Var M : RMVT ; Simul : String3) : Boolean ;
Var Q1 : TQuery ;
    Trouv : boolean ;
begin
TrouveSaisie:=FALSE ;
if (Q.EOF) And (Q.Bof) then Exit ; //n°1825
If Not QOK(Q,Simul) Then BEGIN Result:=FALSE ; Exit ; END ;
If Simul='' Then
  Q1:=OpenSQL('SELECT ' + SQLForIdent( fbGene ) + ' FROM ECRITURE WHERE E_JOURNAL="'+Q.FindField('E_JOURNAL').AsString+'"'
            +' AND E_EXERCICE="'+QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime))+'"'
            +' AND E_DATECOMPTABLE="'+USDATETIME(Q.FindField('E_DATECOMPTABLE').AsDateTime)+'"'
            +' AND E_NUMEROPIECE='+Q.FindField('E_NUMEROPIECE').AsString,True)
            Else
  Q1:=OpenSQL('SELECT ' + SQLForIdent( fbGene ) + ' FROM ECRITURE WHERE E_JOURNAL="'+Q.FindField('E_JOURNAL').AsString+'"'
            +' AND E_EXERCICE="'+QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime))+'"'
            +' AND E_DATECOMPTABLE="'+USDATETIME(Q.FindField('E_DATECOMPTABLE').AsDateTime)+'"'
            +' AND E_NUMEROPIECE='+Q.FindField('E_NUMEROPIECE').AsString
            +' AND E_QUALIFPIECE="'+Simul+'" ',True) ;
Trouv:=Not Q1.EOF ; if Trouv then M:=MvtToIdent(Q1,fbGene,False) ; Ferme(Q1) ;
TrouveSaisie:=Trouv ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 25/01/2005
Modifié le ... :   /  /    
Description .. : - FQ 15077 - CA - 25/01/2005 - Correction plantage appel 
Suite ........ : saisie depuis recherche d'écritures en Web Access
Mots clefs ... :
*****************************************************************}
Function TrouveEtLanceSaisie(Q : TQuery ; TypeAction : TActionFiche ; Simul : String3 ; ModLess : Boolean = FALSE) : Boolean ;
Var M : RMVT ;
begin
Result:=TrouveSaisie(Q,M,Simul) ;
if Result then
   BEGIN
  {$IFDEF RECUPSISCOPGI}
   if ((M.ModeSaisieJal<>'-') and (M.ModeSaisieJal<>''))
      then LanceSaisieFolio(Q,TypeAction) else
  {$ELSE}
    {$IFNDEF GCGC}
   if ((M.ModeSaisieJal<>'-') and (M.ModeSaisieJal<>''))
      then LanceSaisieFolio(Q,TypeAction) else
    {$ENDIF}  // GCGC
  {$ENDIF}    // RECUPSISCOPGI
  LanceSaisie(Q,TypeAction,M,ModLess) ;
   END ;
END ;

Function TFSaisie.Recharge(QM : RMVT ) : Boolean ;
Var ii : integer ;
BEGIN
Result:=TRUE ; ModifNext:=False ;
if ((Action=taModif) and (PieceModifiee)) then
   BEGIN
   ii:=HPiece.Execute(28,caption,'') ;
   Case ii of
      mrYes : BEGIN
              ModifNext:=True ; ClickValide ; ModifNext:=False ;
              if PieceModifiee then Exit ;
              END ;
      mrCancel : Exit ;
      End ;
   END ;
Action:=OldAction ; M:=QM ; ReinitPiece(True) ; FormShow(Nil) ;
END ;

procedure TFSaisie.BPrevClick(Sender: TObject);
Var QM : RMVT ;
begin
if ((QListe=Nil) or (Not BPrev.Enabled) or (Not BPrev.Visible) or (Action=taCreat)) then Exit ;
GridEna(GS,False) ;
BNext.Enabled:=TRUE ;
If QListe.BOF then BPrev.Enabled:=FALSE else
   BEGIN
   QListe.Prior ;
   if TrouveSaisie(QListe,QM,M.Simul) then
      BEGIN
      if ((QM.ModeSaisieJal<>'') and (QM.ModeSaisieJal<>'-')) then QListe.Next else
         if Not Recharge(QM) then QListe.Next ;
      END ;
   END ;
GridEna(GS,True) ;
end;

procedure TFSaisie.BNextClick(Sender: TObject);
Var QM : RMVT ;
begin
if ((QListe=Nil) or (Not BNext.Enabled) or (Not BNext.Visible) or (Action=taCreat)) then Exit ;
GridEna(GS,False) ;
BPrev.Enabled:=TRUE ;
if QListe.EOF then BNext.Enabled:=FALSE else
   BEGIN
   QListe.Next ;
   if TrouveSaisie(QListe,QM,M.Simul) then
      BEGIN
      if ((QM.ModeSaisieJal<>'') and (QM.ModeSaisieJal<>'-')) then QListe.Prior else
         if Not Recharge(QM) then QListe.Prior ;
      END ;
   END ;
GridEna(GS,True) ;
end;

procedure TFSaisie.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFSaisie.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFSaisie.BMenuZoomMouseEnter(Sender: TObject);
begin
PopZoom97(BMenuZoom,POPZ) ;
end;

procedure TFSaisie.BValideMouseEnter(Sender: TObject);
begin
Valide97.Tag:=1 ;
end;

procedure TFSaisie.BValideMouseExit(Sender: TObject);
begin
Valide97.Tag:=0 ;
end;

procedure TFSaisie.GereTagEnter(Sender: TObject);
Var B : TToolbarButton97 ;
begin
B:=TToolbarButton97(Sender) ; if B=Nil then Exit ;
if B.Parent=Outils then Outils.Tag:=1 ;
end;

procedure TFSaisie.GereTagExit(Sender: TObject);
Var B : TToolbarButton97 ;
begin
B:=TToolbarButton97(Sender) ; if B=Nil then Exit ;
if B.Parent=Outils then Outils.Tag:=0 ;
end;

procedure TFSaisie.LeTimerTimer(Sender: TObject);
begin
LeTimer.Enabled:=False ; ModeLC:=False ; GridEna(GS,True) ;
if ((GenereAuto) and (SuiviMP) and (M.smp in [smpEncTous,smpDecTous,smpCompenCli,smpCompenFou])) then BValideClick(Nil) else    {FP 25/04/2006 FQ17966}
 if ((M.MajDirecte) and (M.SorteLettre=tslAucun)) then BValideClick(Nil) else
  if ((Not M.MajDirecte) and (M.SorteLettre<>tslAucun)) then
    BEGIN
    ClickCheque ;
    if M.smp=smpEncTraEdtNC then BAbandonClick(Nil) else
     if M.smp=smpDecBorEdtNC then BAbandonClick(Nil) else
      if M.smp=smpDecChqEdtNC then BAbandonClick(Nil) else
       if M.smp=smpDecVirEdtNC then BAbandonClick(Nil) else
        if M.smp=smpDecVirInEdtNC then BAbandonClick(Nil) else
          if M.smp=smpEncPreEdtNC then BAbandonClick(Nil) else
         If M.MSED.MultiSessionEncaDeca Then BValideClick(Nil) ;
    END ;
end;

procedure TFSaisie.PEnteteEnter(Sender: TObject);
begin
if ModeLC then
   BEGIN
   if GS.CanFocus then GS.SetFocus else if Valide97.CanFocus then Valide97.SetFocus ;
   END ;
end;

procedure TFSaisie.BPopTvaChange(Sender: TObject);
Var O : TOBM ;
    i : integer ;
    CGen : TGGeneral ;
    Ste  : String ;
    OMODR : TMOD ;
begin
if GeneCharge then Exit ;
if VH^.PaysLocalisation=CodeISOES then BPopTva.ItemIndex:=ord(tvaDebit)+1 ; //En la version Espagnole on ne supporte que la TVA débit //XVI 24/02/2005
ChoixExige:=True ; GridEna(GS,False) ;
ExigeEntete:=TExigeTva(BPopTva.ItemIndex-1) ;
if SorteTva<>stvDivers then
   BEGIN
   for i:=1 to GS.RowCount-2 do if isHT(GS,i,True) then
       BEGIN
       O:=GetO(GS,i) ; if O=Nil then Break ;
       CGen:=GetGGeneral(GS,i) ; if CGen=Nil then Break ;
       StE:=FlagEncais(ExigeEntete,CGen.TvaSurEncaissement) ;
       O.PutMvt('E_TVAENCAISSEMENT',StE) ;
       END ;
   CalculDebitCredit ;
   END else
   BEGIN
   for i:=1 to GS.RowCount-2 do if isTiers(GS,i) then
       BEGIN
       OMODR:=TMOD(GS.Objects[SA_NumP,i]) ; if OMODR=Nil then Continue ;
       OMODR.ModR.ModifTva:=False ; GereTvaEncNonFact(i) ;
       END ;
   END ;
GridEna(GS,True) ;
end;

procedure TFSaisie.GereAffSolde(Sender: TObject);
Var Nam : String ;
    C   : THLabel ;
begin
Nam:=THNumEdit(Sender).Name ; Nam:='L'+Nam ;
C:=THLabel(FindComponent(Nam)) ;
if C<>Nil then C.Caption:=THNumEdit(Sender).Text ;
end;

{================================ Tiers payeur ==================================}
Procedure TFSaisie.DetruitPiecesTP ;
Var i : integer ;
    PieceTP : String ;
BEGIN
{$IFNDEF GCGC}
{$IFNDEF CCS3}
for i:=0 to ListeTP.Count-1 do
    BEGIN
    PieceTP:=ListeTP[i] ;
    SupprimePieceTP(PieceTP) ;
    END ;
{$ENDIF}
{$ENDIF}
END ;

Procedure TFSaisie.DetruitOldPayeurs ;
BEGIN
{$IFNDEF GCGC}
if Action<>taModif then Exit ;
if Not VH^.OuiTP then Exit ;
if M.Simul<>'N' then Exit ;
if M.ANouveau then Exit ;
if Not EstJalFact(M.Jal) then Exit ;
if YATP=yaNL then DetruitPiecesTP ;
{$ENDIF}
END ;

Procedure TFSaisie.GereTiersPayeur ;
Var X : RMVT ;
BEGIN
{$IFNDEF GCGC}
{$IFNDEF CCS3}
if YATP=yaNL then DetruitPiecesTP ;
FillChar(X,Sizeof(X),#0) ;
X.Axe:='' ; X.Etabl:=E_ETABLISSEMENT.Value ;
X.Jal:=E_JOURNAL.Value ; X.Exo:=QuelExoDT(StrToDate(E_DATECOMPTABLE.Text)) ;
X.CodeD:=E_DEVISE.Value ; X.Simul:=M.Simul ; X.Nature:=E_NATUREPIECE.Value ;
X.DateC:=StrToDate(E_DATECOMPTABLE.Text) ; X.DateTaux:=DEV.DateTaux ;
X.Num:=NumPieceInt ; X.TauxD:=DEV.Taux ; X.Valide:=False ; X.ANouveau:=False ;
GenerePiecesPayeur(X) ;
{$ENDIF}
{$ENDIF}
END ;

Procedure TFSaisie.YaBienTP ( QEcr : TDataSet ) ;
Var PieceTP,TiersP,Nat,TestJ : String ;
BEGIN
{$IFNDEF GCGC}
{$IFNDEF CCS3}
if YATP=yaL then Exit ;
if Action<>taModif then Exit ;
if Not VH^.OuiTP then Exit ;
if M.Simul<>'N' then Exit ;
if M.ANouveau then Exit ;
if Not EstJalFact(M.Jal) then Exit ;
Nat:=M.Nature ;
if ((NAT<>'FC') and (NAT<>'FF') and (NAT<>'AC') and (NAT<>'AF')) then Exit ;
if ((Nat='FC') or (Nat='AC')) then
   BEGIN
   if VH^.JalVTP='' then Exit ;
   TestJ:=RechDom('TTJOURNAUX',VH^.JalVTP,False) ; if ((TestJ='') or (TestJ='Error')) then Exit ;
   END ;
if ((Nat='FF') or (Nat='AF')) then
   BEGIN
   if VH^.JalATP='' then Exit ;
   TestJ:=RechDom('TTJOURNAUX',VH^.JalATP,False) ; if ((TestJ='') or (TestJ='Error')) then Exit ;
   END ;
TiersP:=QEcr.FindField('E_TIERSPAYEUR').AsString ; if TiersP='' then Exit ;
PieceTP:=QEcr.FindField('E_PIECETP').AsString ; if PieceTP='' then Exit ;
ListeTP.Add(PieceTP) ;
YATP:=yaNL ; if ExisteLettrageSurTP(TiersP,PieceTP) then YATP:=yaL ;
{$ENDIF}
{$ENDIF}
END ;

Function TFSaisie.ExisteTP : boolean ;
Var i : integer ;
    OKP,LaQuestion,AvoirRbt : boolean ;
    Nat : String ;
    GAux : TGTiers ;
BEGIN
Result:=False ; OkP:=False ; LaQuestion:=False ; AvoirRbt:=False ;
if Action=taConsult then Exit ;
if Not VH^.OuiTP then Exit ;
If not GetParamSocSecur('SO_AUTOTP',False) then exit; {FQ20760 20.06.07 YMO On court-circuite le traitement TP plus tôt}
if M.Simul<>'N' then Exit ;
if M.ANouveau then Exit ;
if Not EstJalFact(E_JOURNAL.Value) then Exit ;
Nat:=E_NATUREPIECE.Value ;
if ((NAT<>'FC') and (NAT<>'FF') and (NAT<>'AC') and (NAT<>'AF')) then Exit ;
if ((Nat='FC') or (Nat='AC')) then if VH^.JalVTP='' then
   BEGIN
   HPiece.Execute(49,Caption,'') ;
   Exit ;
   END ;
if ((Nat='FF') or (Nat='AF')) then if VH^.JalATP='' then
   BEGIN
   HPiece.Execute(49,Caption,'') ;
   Exit ;
   END ;
for i:=1 to GS.RowCount-1 do
    BEGIN
    GAux:=GetGTiers(GS,i) ; if GAux=Nil then Continue ;
    if ((GAux.AuxiPayeur<>'') and (Not GAux.IsPayeur)) then
       BEGIN
       OkP:=True ;
       if GAux.DebrayePayeur then LaQuestion:=True ;
       AvoirRbt:=GAux.AvoirRbt ;
       Break ;
       END ;
    END ;
if Not OKP then Exit ;
if ((Nat='AC') or (Nat='AF')) then if AvoirRbt then Exit ;
if LaQuestion then if HPiece.Execute(47,Caption,'')<>mrYes then Exit ;
Result:=True ;
END ;


{***********A.G.L.***********************************************
Auteur  ...... : COMPTA
Créé le ...... : 01/01/2003
Modifié le ... : 22/07/2004
Description .. : Détruit la ligne d'un guide lors d'un CTRL+SUPP 
Suite ........ : ( VL 01/2003 : FQ10809 )
Suite ........ : 
Suite ........ : 22/07/2004 : SBO, Correction algo + FQ10233
Mots clefs ... : 
*****************************************************************}
procedure TFSaisie.DetruitLigneGuide(Lig: integer);
var iLigne    : integer ;
    OG  : TOBM ;
begin

  // Renumérotation des lignes du guide qui se trouvent
  //  au-delà de la ligne en cours de suppression...
  for iLigne := NbLG downto Lig + 1 do
    begin
    if iLigne > TGUI.Count then continue ;
    if iLigne <= 1 then break ;

    OG  := TOBM( TGUI[ iLigne - 1 ] ) ;
    if OG<>nil then
      OG.PutValue( 'EG_NUMLIGNE', iLigne - 1 ) ;
    end;

  // Suppression de la ligne de guide
  TGUI.Delete( Lig - 1 ); // Correction SBO 22/07/2003 ( FQ 10233 )
  TGUI.Capacity := TGUI.Count;
  NbLG          := TGUI.Count; // ??

end;

{$IFDEF SCANGED}
procedure TFSaisie.SetGuidId ( vGuidId : string ) ;
begin

 if ( RechGuidId <> '' ) and ( PGIAsk('Il existe une image ou un fichier associé , voulez-vous le supprimer ?') = mrNo ) then exit ;

 AjouteGuidId(vGuidId) ;

 GereEnabled(GS.Row) ;

end ;

function TFSaisie.RechGuidId : string ;
var
 lTOBEcr   : TOB ;
begin

 result := '' ;

 lTOBEcr := GetO(GS,1) ;
 if lTOBEcr = nil then exit ;

 result := VarAsType(CGetValueTOBCompl(lTOBEcr,'EC_DOCGUID'),varString) ;

 if result = #0 then result := '' ;

end ;

procedure TFSaisie.AjouteGuidId( vGuidId : string ) ;
var
 lTOBCompl : TOB ;
 lTOBEcr   : TOB ;
begin

 lTOBEcr := GetO(GS,1) ;
 if lTOBEcr = nil then exit ;

 lTOBCompl := CGetTOBCompl(lTOBEcr) ;

 if lTOBCompl = nil then
  lTOBCompl := CCreateTOBCompl(lTOBEcr,FEcrCompl)
   else
    if Length(lTOBCompl.GetValue('EC_DOCGUID')) > 1 then
     SupprimeDocumentGed(lTOBCompl.GetValue('EC_DOCGUID')) ;

 lTOBCompl.PutValue('EC_DOCGUID', vGuidId) ;

end ;

procedure TFSaisie.SupprimeLeDocGuid ;
var
 lTOBCompl : TOB ;
 lTOBEcr   : TOB ;
begin

 lTOBEcr := GetO(GS,1) ;
 if lTOBEcr = nil then exit ;

 lTOBCompl := CGetTOBCompl(lTOBEcr) ;

 if lTOBCompl <> nil then
  begin
   if Length(lTOBCompl.GetValue('EC_DOCGUID')) > 1 then
    SupprimeDocumentGed(lTOBCompl.GetValue('EC_DOCGUID')) ;
   lTOBCompl.PutValue('EC_DOCGUID', '' ) ;
  end ;

end ;

function TFSaisie.GetInfoLigne : string;
begin
 result:='Document scanné Période : '+E_DATECOMPTABLE.Text+' Journal : '+E_JOURNAL.Text+' Folio : '+E_NUMEROPIECE.Caption ;
end;

{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 29/09/2003
Modifié le ... : 05/02/2007
Description .. : LG - 29/09/2003 - FB 12777 - ajout du bp scan
Suite ........ : - LG - 05/02/2007 - ajout du bForceModale a true, sion 
Suite ........ : plantage depuis el journal centralisateur
Mots clefs ... : 
*****************************************************************}
procedure TFSaisie.BScanClick(Sender: TObject);
begin
{$IFDEF SCANGED}
ShowGedViewer(RechGuidId,true) ;
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 25/08/2003
Modifié le ... : 23/09/2003
Description .. : Uniquement pour CCMP...
Suite ........ :
Suite ........ : Réécriture de SetNumTraiteDocRegle, en remprenant le
Suite ........ : numéro de traite/Chèque directement dans la tob mise à
Suite ........ : jour depuis DocRegl lors de l'édition.
Suite ........ : Evite de refaire requête et d'oublier des maj.
Suite ........ :
Suite ........ : FICHE 12017
Mots clefs ... :
*****************************************************************}
Procedure TFSaisie.SetNumTraiteDOCREGLE ;
Var i     : integer ;
    RR    : RMVT ;
    TobD  : TOB ;
    sTra  : String ;
    SQL   : string;
begin
  // Tests
  if bReedition then exit;
  if (M.smp<>smpEncTraEdtNC) And (M.smp<>smpDecBorEdtNC) then Exit ;
  if ((Action<>taCreat) and (Not SuiviMP)) then Exit ;
  if M.SorteLettre=tslAucun then Exit ;
  if ModeLC then Exit ;

  // Parcours des écheances d'origines
  For i := 0 to (TOBNumChq.Detail.Count - 1) do
    begin
    SQL:=''; // Réinitialisation  YMO  05/12/2005  FQ 17032
    TobD  := TOBNumChq.Detail[i] ;
    if TobD.GetValue('CODELC') <> '' then
      begin
      RR   := DecodeLC( TobD.GetValue('CODELC') ) ;
      //test pour éviter msg SAV de non existence du champ dans la TOB
      if TobD.GetNumChamp('NUMTRAITECHQ') > 0 then
        sTra := TobD.GetString('NUMTRAITECHQ')
      else
        sTra:='';

      {MAJ numéro de chèque / traite}
      if sTra <> '' then SQL := 'E_NUMTRAITECHQ = "' + sTra + '"';
      if SuiviMP then
      begin
        {JP 28/01/05 : Mise à jour de E_TRESOSYNCHRO et de E_BANQUEPREVI}
        if (TobD.GetNumChamp('BANQUEPREVI') > 0) and (TobD.GetString('BANQUEPREVI') <> '') and
           {JP 20/05/05 : FQ 15921 : le TobD.GetString('BANQUEPREVI') renoie une espèce de #0 qui n'est pas
                          tout à fait équivalent à '' => donc le test ci-dessous est nécessaire faute de pouvoir
                          suggérer une modification à l'agl}
           (Length(Trim(TobD.GetValue('BANQUEPREVI'))) > 0) then begin
          if sTra <> '' then SQL := SQL + ', E_BANQUEPREVI = "' + TobD.GetString('BANQUEPREVI') + '", E_TRESOSYNCHRO = "' + ets_BqPrevi + '" '
                        else SQL := 'E_BANQUEPREVI = "' + TobD.GetString('BANQUEPREVI') + '", E_TRESOSYNCHRO = "' + ets_BqPrevi + '" ';
        end;
      end;

      if (SQL <> '') then
      begin
        SQL := 'UPDATE ECRITURE SET ' + SQL;
        SQL := SQL + ' WHERE ' + WhereEcriture(tsGene, RR, True);
        if ExecuteSQL(SQL) <> 1 then begin
          V_PGI.IoError:=oeUnknown;
          Break ;
        end;
      end ;
    end ;
  end;

  // Supprimer les enregistrements temporaires
  ExecuteSQL('DELETE from DOCREGLE Where DR_USER="'+V_PGI.USER+'"') ;
  ExecuteSQL('DELETE from DOCFACT Where DF_USER="'+V_PGI.USER+'"') ;
  FiniMove ;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 25/08/2003
Modifié le ... : 23/09/2003
Description .. : Pour utilisation dans CCMP uniqement...
Suite ........ :
Suite ........ : Stocke dans TobNumChq la liste des numéro de traite/chq
Suite ........ : pour chaque pièce d'origine, afin de pouvoir la récupérer
Suite ........ : lors de la maj pendant le lettrage
Suite ........ :
Suite ........ : // Modif fiche 12017 : pb maj E_NUMTRAITECHQ
Suite ........ :
Mots clefs ... :
*****************************************************************}
procedure TFSaisie.StockNumChqDepuisEche(LL: TList);
Var i         : integer ;
    O         : TOBM ;
    TobDetail : TOB ;
begin

  FreeAndNil(TOBNumChq) ;

  TOBNumChq := TOB.Create('VNUMCHQ', nil, -1);

  For i := 0 to ( LL.Count - 1 ) do
    begin
    O     := TOBM( LL[i] ) ;
    {JP 18/08/05 : FQ 16486 : Pour que la correction effectuée dans MajTEcheOrig soit pertinente
                      il faut que TOBNumChq contienne toutes les écritures traitées
    if (O.FieldExists('NUMTRAITECHQ')) and (O.GetValue('NUMTRAITECHQ') <> '')
                                       and (O.GetValue('E_TRACE') <> '' ) then}
    if O.GetValue('E_TRACE') <> '' then
    begin
      TobDetail := TOB.Create('VNUMCHQDETAIL', TOBNumChq, - 1) ;
      if (O.FieldExists('NUMTRAITECHQ')) and (O.GetValue('NUMTRAITECHQ') <> '') then
        TobDetail.AddChampsupValeur('NUMTRAITECHQ', O.GetValue('NUMTRAITECHQ')) ;

      TobDetail.AddChampsupValeur('CODELC',       O.GetValue('E_TRACE')) ;
      if SuiviMP then
      begin
        //test pour éviter msg SAV de non existence du champ dans la TOB
        IF O.GetNumChamp('BANQUEPREVI') > 0 then
          TobDetail.AddChampsupValeur('BANQUEPREVI',  O.GetValue('BANQUEPREVI'))
        else
          TobDetail.AddChampsupValeur('BANQUEPREVI',  '')
      end;
      end ;
    end ;

end;


procedure TFSaisie.SupprimeEcrCompl ( vTOB : TOB ) ;
var
 lStGuid : string ;
begin

 lStGuid := CGetValueTOBCompl(vTOB,'EC_DOCGUID') ;
 if lStGuid = #0 then lStGuid := '' ;

 if ( lStGuid = '' ) then
  CFreeTOBCompl(vTOB)
   else
    begin
     CPutValueTOBCompl(vTOB,'EC_CUTOFFDEB',iDate1900) ;
     CPutValueTOBCompl(vTOB,'EC_CUTOFFFIN',iDate1900) ;
    end ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Piot
Créé le ...... : 21/02/2005
Modifié le ... : 19/06/2007
Description .. : - LG 21/05/2005 - initialisation de NumRad
Suite ........ : - LG - 25/08/2005 - LG  - modif de la gestion du cutoff
Suite ........ : - LG - 19/06/2007 - FB 20222 - on prends la date 
Suite ........ : comptabel, la valeur n'etais aps correct ds la variable O
Mots clefs ... : 
*****************************************************************}
procedure TFSaisie.GereCutoff(Lig : integer) ;
Var O      : TOBM ;
    ModBN  : boolean ;
    StComp,StLibre : String ;
    RC             : R_COMP ;
    lTOBCompl      : TOB ;
begin
if PasModif then Exit ;
O:=GetO(GS,Lig) ; if O=Nil then Exit ;
if O.GetValue('E_QUALIFORIGINE') = 'CUT' then exit ; // ecriture genere par la traitement de cutoff
if not FInfo.LoadCompte( GS.Cells[SA_Gen,Lig] ) then exit ;
if FInfo.Compte_GetValue('G_CUTOFF') <> 'X' then
 begin
   if O.CompE then begin SupprimeEcrCompl(O) ; O.CompE := false ; end ;
  exit ;
 end ;
if O.CompE then exit ;
RC.StComporte:=StComp ; RC.StLibre:=StLibre ; RC.Conso:=False ; RC.DateC:=DatePiece ;
RC.Attributs:=(Action=taCreat) ; RC.MemoComp:=MemoComp ; RC.Origine:=0 ;
RC.StLibre:='---CUTXXXXXXXXXXXXXXXXXXXXXXXX' ;
RC.StComporte:='--XXXXXXXX' ; RC.AvecCalcul := true ;
RC.CutOffPer := FInfo.Compte.GetValue('G_CUTOFFPERIODE') ;
RC.CutOffEchue := FInfo.Compte.GetValue('G_CUTOFFECHUE') ;
O.CompE:=True ;
lTOBCompl := CGetTOBCompl(O) ;
if lTOBCompl = nil then
 lTOBCompl := CCreateTOBCompl(O,FEcrCompl) ;
lTOBCompl.PutValue('EC_DATECOMPTABLE', StrToDate(E_DATECOMPTABLE.Text)) ;
CCalculDateCutOff( lTOBCompl, FInfo.Compte_GetValue('G_CUTOFFPERIODE') , FInfo.Compte_GetValue('G_CUTOFFECHUE') ) ;
RC.TOBCompl := lTOBCompl ;
if not SaisieComplement(O,EcrGen,Action,ModBN,RC) then
 begin
  O.CompE:=True ;
  SupprimeEcrCompl(O) ;
 end ;
if GS.CanFocus then GS.SetFocus ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 22/02/2005
Modifié le ... : 27/08/2007
Description .. : - Ajout un triangle ds la Cellule des generaux si une ecriture 
Suite ........ : complementaire est associee a la ligne
Suite ........ : - LG - 27/08/2007 - FB 23230 - coorection de l'affichage 
Suite ........ : du triangle indiquant une ecriture de cutoff
Mots clefs ... : 
*****************************************************************}
procedure TFSaisie.PostDrawCell(ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
var
 O : TOBM ;
 Rect : TRect ;
 T1,T2,T3 : TPoint ;
 lBoCutOff : boolean ;
 lBoBlocNote : boolean ;
begin
 if (ARow<GS.FixedRows) then Exit ;
 if ACol<>SA_GEN then Exit ;
 O:=GetO(GS,ARow) ; if O=Nil then Exit ;
 lBoBlocNote := trim(O.GetString('E_BLOCNOTE')) <> '' ;
 lBoCutOff   := VarAsType ( CGetValueTOBCompl(O,'EC_CUTOFFDEB') , varstring )  <> #0 ;
 if lBoBlocNote or lBoCutOff then
  begin
   if lBoCutOff then begin Canvas.Brush.Color:= clBlue ; Canvas.Pen.Color:=clBlue ; end
   else begin Canvas.Brush.Color:= clRed ; Canvas.Pen.Color:=clRed ; end ;
   Rect:=GS.CellRect(SA_GEN,ARow) ;
   Canvas.Brush.Style:=bsSolid ;
   Canvas.Pen.Mode:=pmCopy ; Canvas.Pen.Width:= 1 ;
   T1.X:=Rect.Right-5 ; T1.Y:=Rect.Top+1 ;
   T2.X:=T1.X+4       ; T2.Y:=T1.Y ;
   T3.X:=T2.X         ; T3.Y:=T2.Y+4 ;
   Canvas.Polygon([T1,T2,T3]) ;
  end ; // if
end ;

function TFSaisie.GetCompteAcc( Ou : integer ; E_AUXILIAIRE : string ) : string ;
var
 i : integer ;
 lIndex : integer ;

 function _GetContre : string ;
 var
  lEcr : TOBM ;
  i : integer ;
 begin
   for i:=Ou-1 downto 1 do
      begin
       if E_AUXILIAIRE = GS.Cells[SA_AUX, i]  then
        begin
         lEcr:=GetO(GS,i) ; if lEcr=nil then Exit ;
         result:=lEcr.GetValue('E_CONTREPARTIEGEN') ;
         Break ;
       end;
      end ; // for
 end ;


begin
 result := '' ;
 if (FInfo.Aux_GetValue('YTC_ACCELERATEUR')='X') then
  begin
   result:= VarToStr(FInfo.Aux_GetValue('YTC_SCHEMAGEN')) ;
   if result = '' then result := _GetContre ;
  end  // if
   else
    begin
     result := _GetContre ;
     if result = '' then
      result:= VarToStr(FInfo.Aux_GetValue('YTC_SCHEMAGEN')) ;
    end ;

 if result = '' then
  begin
   for i:=Ou-1 downto 1 do
    if isHT(GS,i,True) then
     begin
      result:=GS.Cells[SA_GEN, i] ;
      Break ;
     end;
  end ;

 if result <> '' then
  begin
   lIndex:=FInfo.Compte.GetCompte(result) ;
   if (lIndex=-1) or FInfo.Compte.IsCollectif(lIndex)  then
    result := '' ;
  end ;

end;


{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 26/07/2005
Modifié le ... : 02/09/2005
Description .. : - FB 15909 - LG - on désactive l'accélérateure de saisie si 
Suite ........ : un scenario existe sur le journal
Suite ........ : - FB 15959 - ajout de l'affection du e_modesaisie
Mots clefs ... : 
*****************************************************************}
function TFSaisie.GereAcc ( Ou : integer = - 1 ) : integer ;
var
 EcrRef,Ecr : TOBM ;
 lStGen : string ;
 lStCompteTVA : string ;
 lRdTauxTva   : double ;
 lStNatureAuxi : string ;
 lStRegimeTVA  : string ;
 lBoMvtDebit : boolean ;
 C             : double ;
 lRDVal,lRdTva : double ;
 lBoTvaEnc     : Boolean ;
 lStNatPiece   : string ;
 ACol,ARow               : integer ;
 Cancel                  : boolean ;
 Chg                     : boolean ;
begin
 result := - 1 ;

 if (SAJAL.J_ACCELERATEUR = '-') or OkScenario then exit ;
 if Modedevise then exit ;
 if Action <> taCreat then Exit ; // Uniquement en mode création pour le moment...
 if Ou = - 1 then Ou := GS.Row ;
 FInfo.Load(GS.Cells[SA_Gen,Ou],GS.Cells[SA_Aux,Ou],'') ;

 if FInfo.Aux_GetValue('YTC_ACCELERATEUR') <> 'X' then exit ;

 EcrRef := GetO(GS,Ou) ;
 lStGen := GetCompteAcc(Ou,GS.Cells[SA_Aux,Ou]) ;
 if lStGen='' then exit ;

  // Si achat
  FInfo.Compte.GetCompte( lStGen ) ;
  lStNatureAuxi := FInfo.Aux_GetValue('T_NATUREAUXI') ;
  lStRegimeTVA  := FInfo.Aux_GetValue('T_REGIMETVA') ;
  lStNatPiece   := EcrRef.GetValue('E_NATUREPIECE') ;
  // détermination du type de tva (débit / encaiss) // FQ 18525 : gestion de la tva/encaissement avec Accélérateur de saisie
  if (SAJAL.NATUREJAL='ACH') or ( CaseNatP( lStNatPiece ) in [4,5,6] ) or ( CASENATA( lStNatureAuxi ) = 2 )
    then lBoTvaEnc := VH^.OuiTvaEnc and ( ( FInfo.Aux_GetValue('T_TVAENCAISSEMENT')='TE' ) or
                                          ( (FInfo.Aux_GetValue('T_TVAENCAISSEMENT')='TM') and
                                            (FInfo.Compte.GetValue('G_TVASURENCAISS')='X' ) ) )
    else lBoTvaEnc := VH^.OuiTvaEnc and ( VH^.TvaEncSociete='TE' ) ;


 FInfo.Compte.RecupInfoTVA(lStGen,
                           EcrRef.GetValue('E_DEBIT') ,
                           lStNatPiece ,
                           SAJAL.NATUREJAL ,
                           lStNatureAuxi ,
                           lStCompteTVA,
                           lRdTauxTva,
                           lStRegimeTVA,
                           FAlse,
                           lBoTvaEnc );

 lBoMvtDebit := EcrRef.GetValue('E_DEBIT') <> 0 ;
 if lBoMvtDebit
   then C := EcrRef.GetValue('E_DEBIT')
   else C := EcrRef.GetValue('E_CREDIT') ;

 if lStCompteTVA <> '' then
  begin
   lRdVal := Arrondi(C / ( 1 + ( lRdTauxTva / 100 ) ),DecDev) ;
   lRdTva := Arrondi(C - lRdVal,DecDev) ;
  end
   else
    begin
     lRdVal := Arrondi(C ,DecDev) ;
     lRdTva := 0 ;
    end;

  // désactivation des evts
  GS.OnCellExit := nil ;
  GS.OnRowExit := nil ;

  // Calcul ligne HT
  Inc(Ou) ;
  GereNewLigne(GS) ;
  DefautLigne(Ou,True) ;
  GS.Row := Ou ;
  Ecr := GetO(GS,Ou) ;
  if Ecr = nil then exit ;
  GS.Cells[SA_Gen,Ou]  := lStGen ;
  GS.Cells[SA_Aux,Ou]  := '' ;
  GS.Cells[SA_RefI,Ou] := EcrRef.GetValue('E_REFINTERNE') ;      // SBO 18/04/2006 : FQ17853
  GS.Cells[SA_Lib,Ou]  := EcrRef.GetValue('E_LIBELLE') ;
  if lBoMvtDebit
    then GS.Cells[SA_Credit,Ou]:=StrS(lRdVal,DECDEV)
    else GS.Cells[SA_DEbit,ou]:=StrS(lRdVal,DECDEV) ;
{  ValideCeQuePeux(Ou) ;
  AlimObjetMvt(Ou,False) ;
  GereCutoff(Ou) ;
}
 ACol:=GS.Col ; ARow:=ou ; Cancel:=FALSE ;
 GSCellExit(GS, ACol, ARow, Cancel) ; if Cancel then exit ; Chg := false ;
 GSRowExit(GS, ARow, Cancel, Chg ); if Cancel then exit ;

 // Calcul ventilation non complète // FQ 18359
// GereAnal( Ou, True, False ) ;
 if Ecr.GetValue('E_ANA')='X' then
   VentilerTob( Ecr, '', 0, DecDev, False) ;

  if lStCompteTVA <> '' then
    begin
    Inc(Ou) ;
    GereNewLigne(GS) ;
    DefautLigne(Ou,True) ;
    GS.Row := Ou ;
    Ecr:=GetO(GS,Ou) ;
    if Ecr = nil then exit ;
    GS.Cells[SA_Gen,Ou]:=lStCompteTVA ;
    GS.Cells[SA_Aux,Ou]:='' ;
    GS.Cells[SA_RefI,Ou]:=EcrRef.GetValue('E_REFINTERNE') ;       // SBO 18/04/2006 : FQ17853
    GS.Cells[SA_Lib,Ou]:=EcrRef.GetValue('E_LIBELLE') ;
    if lBoMvtDebit
      then GS.Cells[SA_Credit,Ou]:=StrS(lRdTva,DECDEV)
      else GS.Cells[SA_DEbit,ou]:=StrS(lRdTva,DECDEV) ;
{    ValideCeQuePeux(Ou) ;
    AlimObjetMvt(Ou,False) ;
    // Calcul ventilation non complète // FQ 18359
//   GereAnal( Ou, True, False ) ;
    if Ecr.GetValue('E_ANA')='X' then
      VentilerTob( Ecr, '', 0, DecDev, False) ;
    AjusteLigne(Ou,True) ;  // FQ 19752 : pb affectation des échéances sur lignes tva divers lettrable
}
    ACol:=GS.Col ; ARow:=Ou ; Cancel:=FALSE ;
    GSCellExit(GS, ACol, ARow, Cancel) ; if Cancel then exit ; Chg := false ;
    GSRowExit(GS, ARow, Cancel, Chg ); if Cancel then exit ;
    end ; // tva

  GereNewLigne(GS) ;
  GS.OnCellExit := GSCellExit ;
  GS.OnRowExit := GSRowExit ;
  GS.Row := GS.RowCount - 1 ;

end ;

procedure TFSaisie.SetTypeExo;
begin
  // FQ 13246 : SBO 30/03/2005
  FInfo.TypeExo   := CGetTypeExo( StrToDate(E_DATECOMPTABLE.Text) ) ;
end;

{JP 29/07/05 : FQ 16032 : Mise à jour de TEcheOrig avec les écritures qui ont réellement
               été traitées
{---------------------------------------------------------------------------------------}
procedure TFSaisie.MajTEcheOrig;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  p : Integer;
  s : string;
  T : HTStrings;
  Q : TOB;
begin
  {On balaie les écritures}
  for n := TEcheOrig.Count - 1 downto 0 do begin
    T := HTStrings(TEcheOrig[n]);
    {On balaie les échéances}
    for p := T.Count - 1 downto 0 do begin
      s := T[p];
      {On regarde si l'échéance figure parmi celles qui ont été traitées dans DocRegl ...}
      Q := TOBNumChq.FindFirst(['CODELC'], [s], True);
      {... Si ce n'est pas le cas, on supprime l'échénce de la liste d'origine}
      if not Assigned(Q) then T.Delete(p);
    end;
    {Si toutes les échéances d'une écriture ont été supprimées, on supprime l'écriture}
    if T.Count = 0 then TEcheOrig.Delete(n);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann MORENO
Créé le ...... : 03/11/2005
Modifié le ... :   /  /
Description .. : Procédure de fermeture de l'appli avec prise en compte de
Suite ........ : fermture du panel inside
Mots clefs ... :
*****************************************************************}
procedure TFSaisie.Fermeture() ;
begin
  Close () ;
  {JP 05/07/07 : FQ 19022 : Si on a répondu non à la confirmation fermeture, on évite de vider le panel !}
  if FNeFermePlus then Exit;

  if IsInside(Self) then
  begin
    CloseInsidePanel(Self) ;
    THPanel(Self.parent).InsideForm := nil;
    THPanel(Self.parent).VideToolBar;
  end;

//  if FClosing and IsInside(Self) then THPanel(parent).CloseInside;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Yann MORENO
Créé le ...... :
Modifié le ... :
Description .. : Uniquement pour CCMP...
Suite ........ :
Suite ........ : Réécriture de SetNumTraiteDocRegle,
Suite ........ : On passe par là si on ne lettre pas
Suite ........ : (possibilité dans le cas des ecritures en devise)
Suite ........ : mise à jour de paramètres
Suite ........ :
Suite ........ :
Mots clefs ... : YMO
*****************************************************************}
Procedure TFSaisie.MajMPSansLettrage ;
Var i     : integer ;
    RR    : RMVT ;
    TobD  : TOB ;
    sTra  : String ;
    SQL   : string;
begin
  // Parcours des écheances d'origines
  For i := 0 to (TOBNumChq.Detail.Count - 1) do
    begin
    SQL:=''; // Réinitialisation  YMO  05/12/2005  FQ 17032
    TobD  := TOBNumChq.Detail[i] ;
    if TobD.GetValue('CODELC') <> '' then
      begin
      RR   := DecodeLC( TobD.GetValue('CODELC') ) ;
      //test pour éviter msg SAV de non existence du champ dans la TOB
      if TobD.GetNumChamp('NUMTRAITECHQ') > 0 then
        sTra := TobD.GetString('NUMTRAITECHQ')
      else
        sTra:='';

      SQL:=SQL+'E_DATEMODIF="'+USTime(NowFutur)+'", E_SUIVDEC="", E_SAISIMP=0';

      {MAJ numéro de chèque / traite}
      if sTra <> '' then SQL := SQL + ', E_NUMTRAITECHQ = "' + sTra + '"';

      //E_ETAT ?
      if SuiviMP then
      begin
        if (TobD.GetNumChamp('BANQUEPREVI') > 0)
        and (TobD.GetString('BANQUEPREVI') <> '')
        and (Length(Trim(TobD.GetValue('BANQUEPREVI'))) > 0) then
          SQL := SQL + ', E_BANQUEPREVI = "' + TobD.GetString('BANQUEPREVI') + '", E_TRESOSYNCHRO = "' + ets_BqPrevi + '" ';
      end;

      SQL := 'UPDATE ECRITURE SET ' + SQL;
      SQL := SQL + ' WHERE ' + WhereEcriture(tsGene, RR, True);
      if ExecuteSQL(SQL) <> 1 then
      begin
        V_PGI.IoError:=oeUnknown;
        Break ;
      end;

    end ;
  end;

end ;


{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 10/04/2006
Modifié le ... :   /  /    
Description .. : récupération des fonctionnalités de saisBor.pas
Suite ........ : ( Calcul dans les champs montants )
Mots clefs ... : 
*****************************************************************}
function TFSaisie.GetFormulePourCalc(Formule: hstring): Variant;
var lCalcRow : integer ;
    lEcr     : TobM ;
begin
  Result  := #0 ;
  Formule := AnsiUpperCase(Trim(Formule)) ;
  if Formule='' then Exit ;
  if pos('E_DEBIT',Formule)>0
    then lCalcRow:=GS.Row
    else lCalcRow:=GS.Row-1 ;

  if not EstRempli( GS, lCalcRow ) then Exit ;
  lEcr := GetO( GS, lCalcRow ) ;
  if lEcr=nil then Exit ;
  if Pos('E_', Formule)>0 then
    Result := lEcr.GetValue(Formule) ;
end;

function TFSaisie.DateEchesOk(Lig: integer): boolean;
var lModR : TMod ;
    i     : integer ;
begin

  result := True ;

  lModR := TMOD( GS.Objects[SA_NumP,Lig] ) ;
  if lModR = Nil then Exit ;

  With lModR.ModR do
    For i:=1 to NbEche do
      if TabEche[i].Pourc > 0 then
        if not NbJoursOk( StrToDate(E_DATECOMPTABLE.Text), TabEche[i].DateEche ) then
          begin
          result := False ;
          break ;
          end ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 14/06/2006
Modifié le ... :   /  /    
Description .. : // FQ18371 et FQ18080
Suite ........ : 
Suite ........ : Permet un calcul arithmétique tenant compte du nombre de 
Suite ........ : décimales de la devise
Mots clefs ... : 
*****************************************************************}
function TFSaisie.GetFormatPourCalc: string;
begin
  // FQ18371 et FQ18080
  result := '"#.###' ;
  if DEV.Decimale > 0 then
    result := result + ',' + Copy('000000000', 1, Dev.Decimale ) ;
  result := result + '"' ;
end;

function TFSaisie.YAFormule(St: string): Boolean;
begin
  if SuiviMP then
  Result:=((Pos('[',St)>0) And (Pos(']',St)>0)) Or (St<>'')
  else Result:=((Pos('[',St)>0) And (Pos(']',St)>0)) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Compta
Créé le ...... : 25/05/2007
Modifié le ... :   /  /    
Description .. : SBO FQ20159 : détermine l'éventuelle fenêtre de saisie 
Suite ........ : appelante
Mots clefs ... : 
*****************************************************************}
function TFSaisie.QuelleSaisieCascade : string ;
var lStcopy : string ;
begin
  lStcopy := SA_PILESAI ;

  ReadTokenSt( lStcopy ) ;            // fenêtre en cours
  result := ReadTokenSt( lStcopy ) ;  // fenêtre précédente

  ReadTokenSt( SA_PILESAI ) ;         // on ôte la fenêtre en cours
  
end;

procedure TFSaisie.BSCANPDFClick(Sender: TObject);
begin
{$IFDEF SCANGED}
 //if ( RechGuidId <> '' ) and ( PGIAsk('Il existe un fichier associé , voulez-vous le supprimer ?') = mrNo ) then exit ;
 AjouterFichierDansGed(Self) ;
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 25/09/2007
Modifié le ... : 25/09/2007
Description .. : Le paramètre Société de controles des QTE Analytique doit
Suite ........ : intervenir comme une valeur par défaut en l'absence de
Suite ........ : scénario
Suite ........ : Jamais effectué en mode PCL
Mots clefs ... :
*****************************************************************}
function TFSaisie.TestQteAna: boolean;
begin
  if ctxPCL in V_PGI.PGIContexte then
    result := False
  else if OkScenario then
    result := Scenario.GetValue('SC_CONTROLEQTE') = 'X'
  else
    result := GetParamSocSecur('SO_ZCTRLQTE', False) ;
end;

end.










