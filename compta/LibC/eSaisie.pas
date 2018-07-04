UNIT eSaisie ;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Grids, Mask, ExtCtrls, ComCtrls, Buttons, Hspliter, Ent1, Menus,
  HEnt1, UtilPGI, HFLabel, About, HStatus, hmsgbox, HQry, Hctrls,
  HDebug, HLines, HSysMenu, HPop97, HTB97, ed_tools, HPanel,
  Math, // pour la fonction Min
{$IFDEF EAGLCLIENT} //===========================================> EAGL
  Maineagl,
  uTOFCPMulGuide, 		// pour LanceGuide, SelectGuide
  CPGENERAUX_TOM,     // pour Zoom sur fiche Gene
  CPTIERS_TOM,        // pour Zoom sur fiche tiers
  CPJOURNAL_TOM,      // pour Zoom sur fiche journal
  DEVISE_TOM,         // pour Zoom sur fiche devise
  SUIVCPTA_TOM,       // pour Zoom sur fiche scénario
  CPMULSAISLETT_TOF,  // pour LettrerEnSaisie
  CPCHANCELL_TOF,     // pour FicheChancel

{$ELSE}	//=======================================================> NON EAGL
  Fe_Main,
  DBTables,
  DB,
	{$IFNDEF GCGC}
	  General,		// pour FicheGene
    Tiers,
    Journal,
    SelGuide,		// pour LanceGuide, SelectGuide
    Scenario,
//    EncUtil,
    ValPerio,
  	SaisLett,		// pour LettrerEnSaisie
  	{$IFNDEF IMP}
  		MulGene,
      MulTiers,
			{$IFNDEF CCMP}
  			ImoGen,	// pour FicheImmobilisation
			{$ENDIF}
  	{$ENDIF}
	{$ENDIF}
  Chancel,			// pour FicheChancel,
  Devise,				// pour FicheDevise,
	UtilSoc,  		// EAGL : EN ATTENTE --> MarquerPublifi...
{$ENDIF} //=======================================================> FIN NON EAGL

//=======================================================> COMMUN EAGL / NON-AGL
{$IFNDEF CCS3}
    TiersPayeur,   // pour GenerePiecesPayeur, SupprimePieceTP, ExisteLettrageSurTP, ZoomPieceTP
  {$IFNDEF GCGC}
  SaisTVA,			// pour TFSaisTVA
  DocRegl,
  {$ENDIF}
{$ENDIF}
{$IFNDEF GCGC}
  //SaisEcar,			// pour CorrigeEcartConvert
  SaisComp,
  SaisEnc,    // pour InfosTvaEnc
	SaisBase,   // pour Enr_Base et SaisieBasesHT
    SaisVisu,
{$ENDIF}
  eSaisAnal, ULibAnalytique,
	UTOB, Filtre, Choix, Formule,
	SaisUtil, SaisComm, LettUtil, FichComm, HCompte, UtilSais,
  EcrPiece_TOF,	// pour LanceZoomPieceGC
  EcheMono,			// pour SaisirMonoEcheance,
  SaisTaux1,		// pour SaisieNewTaux2000,
  Echeance,			// pour CalculModeRegle et	SaisirEcheance,
  ParamSoc,			// pour GetParamSoc et SetParamSoc,
  AGLInit,
  UiUtil,
  ImgList, TntStdCtrls, TntExtCtrls, TntGrids;


{$IFDEF EAGLCLIENT}
Function  LanceSaisieMP ( M : RMVT ; TOBMPOrig,TOBMPEsc : TOB ; GenereAuto : boolean ; TOBParamEsc : TOB = nil ) : boolean ;
{$ELSE}
procedure SaisieGuidee (StG : String3) ;
{$ENDIF}
Procedure LanceMultiSaisie(LesM : TList);
Function  LanceSaisie ( QListe : TQuery ; Action : TActionFiche ; Var M : RMVT ; ModLess : Boolean = FALSE) : boolean ;
Procedure LanceSaisieDirecte;
Function  TrouveSaisie(Q : TQuery ; Var M : RMVT ; Simul : String3) : Boolean ;
Function  TrouveEtLanceSaisie(Q : TQuery ; TypeAction : TActionFiche ; Simul : String3 ; ModLess : Boolean = FALSE) : boolean ;

type
  TFeSaisie = class(TForm)
    GS              : THGrid;
    PEntete: THPanel;
    E_JOURNAL       : THValComboBox;
    H_JOURNAL       : THLabel;
    E_DATECOMPTABLE : TMaskEdit;
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
    BZoom           : TBitBtn;
    BSaisTaux       : TBitBtn;
    BZoomJournal    : TBitBtn;
    BZoomDevise     : TBitBtn;
    BZoomEtabl      : TBitBtn;
    BScenario       : TBitBtn;
    BDernPieces     : TBitBtn;
    BChancel        : TBitBtn;
    BSwapPivot      : TBitBtn;
    BModifRIB       : TBitBtn;
    BInsert         : TBitBtn;
    BSDel           : TBitBtn;
    Cache           : THCpteEdit;
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
    BCValide        : TBitBtn;
    BCAbandon       : TBitBtn;
    BProrata        : TBitBtn;
    BGuide          : TBitBtn;
    BCreerGuide     : TBitBtn;
    BMenuGuide      : TBitBtn;
    BModifSerie     : TBitBtn;
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
    BCheque: TBitBtn;
    TVAImages: TImageList;
    BPopTva: TPopButton97;
    HConf: TToolbarButton97;
    ISigneEuro: TImage;
    LSA_SoldeG: THLabel;
    LSA_SoldeT: THLabel;
    LSA_Solde: THLabel;
    LSA_TotalCredit: THLabel;
    LSA_TotalDebit: THLabel;
    BChoixRegime: TBitBtn;
    BZoomImmo: TBitBtn;
    BZoomTP: TBitBtn;
    BMenuAction: TBitBtn;
    BModifs: TBitBtn;
    BLibAuto: TBitBtn;
    BVidePiece: TBitBtn;
    BPieceGC: TBitBtn;
    procedure FormShow(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
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
//    procedure BSwapEuroClick(Sender: TObject);
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
    procedure E_DEVISEExit(Sender: TObject);
    procedure GereAffSolde(Sender: TObject);
    procedure E_DATECOMPTABLEEnter(Sender: TObject);
    procedure BChoixRegimeClick(Sender: TObject);
    procedure BZoomImmoClick(Sender: TObject);
    procedure BZoomTPClick(Sender: TObject);
    procedure BVidePieceClick(Sender: TObject);
    procedure BPieceGCClick(Sender: TObject);
    procedure BMenuZoomClick(Sender: TObject);
  private
    QListe                     : TQuery ;
    SAJAL                      : TSAJAL ;
    DEV                        : RDEVISE ;
    NumPieceInt,NbLigEcr,NbLigAna : Longint ;
    CpteAuto,OkScenario,OkBQE,OkJalEffet,OkLC,OkPremShow,ModifNext,NeedEdition,RegimeForce : boolean ;
    CurAuto,CurLig             : integer ;
    TS,TGUI,TDELECR,TDELANA,TECHEORIG : TList ;
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
    GeneTypeExo                : TTypeExo ;
    YATP                       : T_YATP ;
    ListeTP                    : TStrings ;
    OkMessAnal,AutoCharged,OkSuite,UnSeulTiers,ModeLC,RentreDate : boolean ;
    ModeDevise,FindFirst,Volatile,{ModeOppose,}OuiTvaLoc,GeneCharge : Boolean ;
    NbLOrig,NbLG,DecDev : integer ;
    ModeForce                  : tmSais ;
    NowFutur,DatePiece,LastEche,LastVal : TDateTime ;
    GeneTreso,LastTraiteCHQ    : String17 ;
    ModeConf                   : String[1] ;
    OldDevise,DernVentiltype,EtatSaisie,StCellCur : String ;
    OldTauxDev                 : Double ;
    MemoComp                   : TStrings ;
    WMinX,WMinY                : Integer ;
    TabTvaEnc                  : Array[1..5] of Double ;
    ModLess : Boolean ;
    TOBNumChq : TOB ; // Modif Fiche 12017 // Enca Deca : Gestion des numéro de traite
// TVA, TPF
    procedure AlimHTByTVA ( Lig : integer ; RegTVA,TVAENC : String3 ; SoumisTpf,AChat : boolean ) ;
    function  InsereLigneTVA ( LigHT : integer ; RegTVA : String3 ; SoumisTpf,Achat : boolean ) : boolean ;
    procedure RemplirEcheAuto ( Lig : integer ) ;
    procedure RemplirAnalAuto ( Lig : integer ) ; 											// OK
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
    procedure AttribNumeroDef ;																				// OK
    procedure AttribLeTitre ;
    procedure ReInitPiece ( Force : boolean ) ;
    procedure GereDeviseEcart ;
    procedure GereEnabled ( Lig : integer ) ;
    procedure GereOptionsGrid ( Lig : integer ) ;
    procedure PosLesSoldes ;
// Euro
//    procedure AddEuroCombo ;
  //  procedure AddEuroFranc ;
    //procedure RestoreDeviseEuro ;
//    procedure SetOpposeEuro ;
// Chargements
    procedure ChargeLignes ;
    procedure ChargeEcritures ;
    procedure ChargeComptes ;
    procedure ChargeSoldesComptes ;																		// OK
    procedure ChargeAnals ;
    procedure RuptureEche( Lig : integer ; TDD,TCD,TDP,TCP,TDE,TCE : Double ) ;
    procedure AlimNoyauEche ( Lig : integer ; Premier : boolean ; LDD,LCD,LDP,LCP,LDE,LCE : Double ; QEcr : TDataSet ) ;
// Allocation et Désallocation
    function  DateToJour ( DD : TDateTime ) : String ;
    procedure AlloueEcheAnal(C,L : LongInt) ;
    procedure AlloueEche(LaLig : LongInt) ;
    procedure DesalloueEche(LaLig : LongInt) ;
    procedure AlloueAnal(LaLig : LongInt) ;															// OK
    procedure DesalloueAnal(LaLig : LongInt) ;													// OK
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
{$IFDEF CEGID}
    procedure TraiteLigneAZero ;
    procedure ResaisirLaDatePourLeGuide(OkOk : Boolean) ;
{$ENDIF}
    procedure FinGuide ;
    procedure LanceGuide ( TypeG,CodeG : String3 ) ;
    function  Load_Sais ( St : String ) : Variant ;
    function  Load_TVATPF ( CGen : TGGeneral ; OKTVA : boolean ; TVA : String ) : Variant ;
// Création du guide en saisie
    procedure CreerLignesGuide ( CodeG : String ) ;		// EAGL : EN ATTENTE
    procedure CreerEnteteGuide ( CodeG : String ) ;		// EAGL : EN ATTENTE
    function  BonRegle ( Lig,NbEche : integer ; LeMode : String ; T : T_TabEche ) : Boolean ;
// Lettres chèque, Encadeca
    Function  ChoixFromTenueSaisie ( LeMontantPIVOT,LeMontantEURO,LeMontantDEV : Double ) : Double ;
    Function  EcheancesLC ( NbC : integer ) : boolean ;

    // Nouvelle procedures suite fiche 12017 : pb maj E_NUMTRAITECHQ
    procedure SetNumTraiteDOCREGLE ;
    procedure StockNumChqDepuisEche   ( LL : TList ) ;
    // Fin modif 12017

    procedure SuiteEncaDeca ;
    procedure NormaliserEscomptes ;
    procedure LettrerReglement ;
    procedure EnvoiCFONB ;
    procedure EnvoiBordereau ;
    procedure CloseFen ;
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
    procedure ClickZoom ;
    procedure ClickConsult ;
    procedure ClickAbandon ;
    procedure ClickGenereTVA ;
    function  ClickControleTVA ( Bloque : boolean ; LigHT : integer ) : boolean ;
    procedure ClickProrata ;
    procedure ClickSwapPivot ;
//    procedure ClickSwapEuro ;
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
    procedure ClickPieceGC ;
    procedure ClickImmo ;
    procedure ClickTP ;
    procedure ClickModifTva ;
    procedure ForcerMode(Cons : boolean ; Key : word) ;
// Calculs lignes
    procedure InfoLigne( Lig : integer ; Var D,C,TD,TC : Double ) ;
    procedure DetruitLigne( Lig : integer ; Totale : boolean) ;
    procedure DetruitLigneGuide( Lig : integer) ;
    procedure CalculDebitCredit ;
    procedure AfficheConf ;
    procedure CalculSoldeCompte ( CpteG,CpteA : String ; DIG,CIG,DIA,CIA : Double ) ;
    procedure SommeSoldeCompte ( Col : integer ; Cpte : String ; Var TD,TC : Double ; Old,OkDev : Boolean) ;
    procedure DefautLigne ( Lig : integer ; Init : boolean ) ;
    procedure PutScenar ( O : TOBM ; Lig : integer ) ;
    procedure TraiteMontant( Lig : integer ; Calc : boolean ) ;
    procedure CaptureRatio( Lig : integer ; Force : boolean ) ;
// Analytiques
//    Procedure AlimAnalDef ( Lig : integer ; OBA : TOB ) ;
    Procedure OuvreAnal ( Lig : integer ; Scena : boolean ) ;
    Function  AOuvrirAnal ( Lig : integer ; Auto : boolean ) : Boolean ;					// OK
    Procedure GereAnal ( Lig : integer ; Auto,Scena : boolean ) ;									// OK
    Procedure RecupAnal(Lig : integer ; Var OBA : TOB ; NumAxe,NumV : integer ) ;			// OK
// Echéances
    Function  AOuvrirEche ( Lig : integer ; Var Cpte : String17 ; Var MR : String3 ;
                            Var OuvreAuto,RempliAuto : boolean ; Var t : TLettre ) : Boolean ;
    Procedure OuvreEche ( Lig : integer ; Cpte : String17 ; MR : String3 ; RempliAuto : boolean ; t : TLettre ) ;
    Procedure GereEche ( Lig : integer ; OuvreAuto,RempliAuto : boolean ) ;
    Procedure RecupEche( Lig : integer ; R : T_ModeRegl ; NumEche : integer; var OBM : TOBM) ;
    Procedure EcheGuideRegle ( Lig : integer ; Var T : T_MODEREGL ) ;
    Procedure EcheLettreCheque ( Lig : integer ; Var T : T_MODEREGL ) ;
    Procedure GereMonoEche ( Var OO : T_ModeRegl ; OkInit,AutoPourEncaDeca : boolean ; Lig : integer ; ActionEche : TActionFiche ) ;
// Contrôles
    Function  LigneCorrecte ( Lig : integer ; Totale,Alim : boolean ) : boolean ;		// OK pour Anal
    Function  PieceCorrecte : boolean ;
    Function  EquilEuroValide : Boolean ;
//    Function  AjouteLigneEcart ( EcartP,EcartE : double ) : boolean ;
    Function  ControleRIB : boolean ;
    Function  PasModif : boolean ;
    Function  Equilibre : boolean ;
    procedure ErreurSaisie ( Err : integer ) ;
    procedure AjusteLigne ( Lig : integer ; AvecM : Boolean ) ;
    Procedure AjusteLigneEuro ( GS : THGrid ; k : integer ) ;
    Function  PossibleRecupNum : Boolean ;
    procedure ControleLignes ;
    function  ControleTVAOK : boolean ;
    function  PasToucheLigne ( Lig : integer ) : boolean ;
// Appels comptes
    Function  ChercheGen ( C,L : integer ; Force : boolean ) : byte ;
    Function  ChercheAux ( C,L : integer ; Force : boolean ) : byte ;
    procedure ChercheJAL ( Jal : String3 ; Zoom : boolean ) ;
    procedure ChercheMontant(Acol,Arow : longint) ;
    procedure AppelAuto ( indice : integer ) ;
    Procedure AffecteLeRib(O : TOBM ; Cpt : String ; ForceRAZ : Boolean = FALSE) ;
    Function  AffecteRIB ( C,L : integer ; IsAux : Boolean ; ForceRAZ : Boolean = FALSE) : Boolean ;
    Procedure AffecteConso ( C,L : integer ) ;
    Procedure AffecteTva ( C,L : integer ) ;
    Function  SortePiece : boolean ;
// MAJ Fichier
    Procedure GetEcr(Lig : Integer) ;																		// Modif EAGL
    Procedure GetEcrGrid(Lig : Integer) ;
    Procedure GetAna(Lig : Integer) ;																		// OK
    Procedure RecupTronc(Lig : Integer ; MO : TMOD; Var OBM : TOBM );		// Modif EAGL
    procedure RecupFromGrid( Lig : integer ; MO : TMOD ; NumE : integer ) ;
    Procedure ClickValide ;
    Procedure DetruitPieceMP ;
    procedure TraiteMP ;
    procedure TraiteOrigMP ;
    procedure ValideLaPiece ;
    procedure ValideLeReste ;
    Procedure ValideLignesGene ;
    Procedure ValideLesComptes ;  		// EAGL : Modifié
    procedure ValideLeJournalNew ;
    Procedure AttribRegimeEtTVA ;
    Procedure RenseigneRegime ( Lig : integer ; Recharge : boolean ) ;
{$IFDEF EAGLCLIENT}
{$ELSE}
    procedure IntoucheGene ; 		// EAGL : En ATTENTE
    procedure IntoucheAux ; 		// EAGL : En ATTENTE
{$ENDIF}
    procedure MajCptesGeneNew ;
    procedure MajCptesAuxNew  ;
    procedure MajCptesSectionNew ;																					// OK
    procedure DetruitAncien ;
    procedure InverseSoldeNew ;
    Procedure EcartChange(Obj : TOB) ;																			// OK
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
//  Procedure GetCellCanvas ( ACol,ARow : Longint; Canvas : TCanvas ; AState: TGridDrawState) ;
  public
    Action,OldAction : TActionFiche ;
    TPIECE : TList ;
    SI_TotDS,SI_TotCS,SI_TotDP,SI_TotCP,SI_TotDD,SI_TotCD,SI_TotDE,SI_TotCE : double ;
    SI_FormuleRef,SI_FormuleLib : String ;
    CEstGood,SuiviMP,GenereAuto : boolean ;
    LesM               : TList ;
    CurM               : integer ;
    TOBMPOrig,TOBMPEsc,
    TOBParamEsc 			 : TOB ;
    DateSais           : TDateTime ;
  end;

implementation

{$R *.DFM}

Uses
{$IFDEF EAGLCLIENT}
	UtileAGL,
{$ELSE}
  {$IFNDEF IMP}
   {$IFNDEF GCGC}
     {$IFDEF V530}
           EdtEtat,
     {$ELSE}
           EdtREtat,
     {$ENDIF}
    ConsEcr,GuiDate,SaisBor,
   {$ELSE}
    GuiDate,
   {$ENDIF}
  {$ELSE}
    SaisBor,
   {$IFDEF CEGID}
    GuiDate,
   {$ENDIF}
  {$ENDIF}
{$ENDIF}
  {$IFNDEF IMP}
   {$IFNDEF GCGC}
    {$IFNDEF CCS3}
      CFONB,			// pour ExportCFONB
    {$ENDIF}
   {$ENDIF}
  {$ENDIF}
   LetBatch,
   UTofConsEcr
	;

Procedure LanceSaisieDirecte;
Var M : RMVT ;
Begin
  FillChar(M,Sizeof(M),#0) ;
  M.Simul			:= 'N' ;
  M.CodeD			:= V_PGI.DevisePivot ;
  M.DateC			:= V_PGI.DateEntree ;
  M.TauxD			:= 1 ;
  M.DateTaux	:= M.DateC ;
  M.Valide		:= False ;
  M.Etabl			:= VH^.ETABLISDEFAUT ;
  M.ANouveau	:= False ;
  LanceSaisie(Nil,taCreat,M) ;
end;

procedure SaisieGuidee ( StG : String3) ;
Var R : RMVT ;
    FromMenu : Boolean ;
    Jal,Nat,Devi,Etab,OldCode : String3 ;
    Q                 : TQuery ;
BEGIN
{$IFNDEF GCGC}
  FromMenu  := False ;
  Jal       := '' ;
  Nat       := '' ;
  Devi      := '' ;
  Etab      := '' ;
  OldCode   := '' ;
  repeat
    if StG='' then
      BEGIN
      FromMenu := True ;
      StG := SelectGuide(Jal,Nat,Devi,Etab,(V_PGI.DateEntree>=V_PGI.DateDebutEuro),OldCode) ;
      END ;
    if StG='' then Exit ;
    FillChar(R,Sizeof(R),#0) ;
    R.DateC         := V_PGI.DateEntree ;
    R.Exo           := QuelExoDT(R.DateC) ;
    R.Simul         := 'N' ;
    R.LeGuide       := StG ;
    R.TypeGuide     := 'NOR' ;
    R.ANouveau      := False ;
    R.SaisieGuidee  := True ;
    LanceSaisie(Nil,taCreat,R) ;
    Screen.Cursor:=SyncrDefault ;
    if FromMenu then
      begin
      Q:=OpenSQL('Select GU_JOURNAL, GU_NATUREPIECE, GU_DEVISE, GU_ETABLISSEMENT from GUIDE Where GU_TYPE="NOR" AND GU_GUIDE="'+StG+'"',True) ;
      if not Q.EOF then
        begin
        Jal     :=Q.Fields[0].AsString  ;
        Nat     :=Q.Fields[1].AsString ;
        Devi    :=Q.Fields[2].AsString ;
        Etab    :=Q.Fields[3].AsString ;
        OldCode :=StG ;
        StG     :='' ;
        end ;
     Ferme(Q) ;
     end ;
  until not FromMenu ;
{$ENDIF}
end ;

Procedure LanceMultiSaisie ( LesM : TList ) ;
var
  {$IFDEF EAGLCLIENT}
    X : TFeSaisie ;
  {$ELSE}
    X : TFSaisie ;
  {$ENDIF}
    M : RMVT ;
begin
  M := P_MV(LesM[0]).R ;
  {$IFDEF EAGLCLIENT}
  X := TFeSaisie.Create(Application) ;
  {$ELSE}
  X := TFSaisie.Create(Application) ;
  {$ENDIF}
  try
    X.Action    := taCreat ;
    X.OldAction := taCreat ;
    X.M         := M ;
    X.LesM      := LesM ;
    X.CurM      := 0 ;
    X.SuiviMP   := False ;
    X.ShowModal ;
    finally
      X.Free ;
    end ;
  Screen.Cursor:=SyncrDefault ;
end ;

Function LanceSaisie ( QListe : TQuery ; Action : TActionFiche ; Var M : RMVT  ; ModLess : Boolean = FALSE) : boolean ;
Var X  : TFeSaisie ;
    OA : TActionFiche ;
    PP : THPanel ;
    DateSais : TDateTime ;
BEGIN
  Result          := False ;
  M.LastNumCreat  := -1 ;
  DateSais        := V_PGI.DateEntree ;
// PFU : DEBUT AJOUT 03/07/2000
  if (Action=taModif) and (SaisieLancer)
    then Action := taConsult ;
// PFU : FIN AJOUT 03/07/2000

  {Gestion exo de référence}
  if ((Action=taCreat) and (VH^.CPExoRef.Code<>'') and (VH^.CPLastSaisie>0) and
    (Not M.SaisieGuidee) and (M.TypeGuide='') and (Not M.ANouveau) and
    (VH^.CPLastSaisie>=VH^.CPExoRef.Deb) and (VH^.CPLastSaisie<=VH^.CPExoRef.Fin) and
    (ctxPCL in V_PGI.PGIContexte))
    then DateSais := VH^.CPLastSaisie ;

  Case Action of
   taCreat : BEGIN
             if PasCreerDate(DateSais) then Exit ;
             if DepasseLimiteDemo then Exit ;
             if _Blocage(['nrCloture'],True,'nrSaisieCreat') then Exit ;
             END ;
   taModif : BEGIN
             if RevisionActive(M.DateC) then Exit ;
             if _Blocage(['nrCloture','nrBatch'],True,'nrSaisieModif') then Exit ;
             END ;
   END ;

  OA  :=Action ;
  PP  := FindInsidePanel ;
  X   := TFeSaisie.Create(Application) ;
  X.QListe    := QListe ;
  X.Action    := Action ;
  X.OldAction := Action ;
  X.M         := M ;
  X.SuiviMP   := False ;
  X.DateSais  := DateSais ;
  X.TOBMPOrig := Nil ;
  X.TOBMPEsc  := Nil ;
  X.TOBParamEsc := nil ;
  X.ModLess   := ModLess ;

  if ModLess
    then X.Show
    else
      if PP=Nil then
        begin
        try
          X.ShowModal ;
          Result := X.CEstGood ;
          M.LastNumCreat := X.M.LastNumCreat ;
          M.MSED.SoucheSpooler := X.M.MSED.SoucheSpooler ;
          M.MSED.ModeleMultiSession := X.M.MSED.ModeleMultiSession ;
          finally
            X.Free ;
            case OA of
              taCreat : _Bloqueur('nrSaisieCreat',False) ;
              taModif : _Bloqueur('nrSaisieModif',False) ;
              end ;
          end ;
        Screen.Cursor:=SyncrDefault ;
        end
    else
      begin
      InitInside(X,PP) ;
      X.Show ;
      end ;
END ;

{$IFDEF EAGLCLIENT}
Function LanceSaisieMP ( M : RMVT ; TOBMPOrig,TOBMPEsc : TOB ; GenereAuto : Boolean ; TOBParamEsc : TOB ) : boolean ;
Var X  : TFeSaisie ;
//    OA : TActionFiche ;
    PP : THPanel ;
BEGIN
{OA:=taModif ;} Result:=False ; M.LastNumCreat:=-1 ;
if RevisionActive(M.DateC) then Exit ;
PP:=FindInsidePanel ;
X:=TFeSaisie.Create(Application) ;
X.QListe:=Nil ; X.Action:=taModif ; X.OldAction:=taModif ; //OA:=taModif ;
X.M:=M ; X.SuiviMP:=True ; X.GenereAuto:=GenereAuto ;
X.TOBMPOrig:=TOBMPOrig ; X.TOBMPEsc:=TOBMPEsc ;
// Prise en compte des paramètres d'escomptes par lignes
	X.TOBParamEsc := TOBParamEsc ;
{$IFDEF CCMP}
If X.SuiviMP And (X.M.Smp in [smpEncTraEdtNC,smpDecBorEdtNC,smpDecVirEdtNC,smpDecVirInEdtNC]) Then
  BEGIN
  X.BorderIcons:=X.BorderIcons+[biMinimize] ; X.WindowState:=wsMinimized ;
  END ;
{$ENDIF}
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
{$ENDIF}


{==========================================================================================}
{================================= Libellés auto ==========================================}
{==========================================================================================}
procedure TFeSaisie.RenseigneLibelleAuto (Col,Lig : integer ; Ref : Boolean );
BEGIN
if GS.Cells[Col,Lig]<>'' then Exit ; if M.ANouveau then Exit ;
if Ref then GS.Cells[Col,Lig]:=GFormule(SI_FormuleRef,Load_Sais,Nil,1)
       else GS.Cells[Col,Lig]:=GFormule(SI_FormuleLib,Load_Sais,Nil,1) ;
END ;

procedure TFeSaisie.GereRefLib(Lig : integer) ;
BEGIN
if PasModif then Exit ; if Guide then Exit ; if M.ANouveau then Exit ;
RenseigneLibelleAuto(SA_RefI,Lig,True) ;
RenseigneLibelleAuto(SA_Lib,Lig,False) ;
END ;

procedure TFeSaisie.ChargeLibelleAuto ;
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
procedure TFeSaisie.AlimNoyauEche ( Lig : integer ; Premier : boolean ; LDD,LCD,LDP,LCP,LDE,LCE : Double ; QEcr : TDataSet ) ;
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

procedure TFeSaisie.RuptureEche( Lig : integer ; TDD,TCD,TDP,TCP,TDE,TCE : Double ) ;
Var OBM : TOBM ;
    TM  : TMOD ;
    i   : integer ;
BEGIN
{
if ModeOppose then
   BEGIN
   GS.Cells[SA_Debit,Lig]:=StrS(TDE,DECDEV) ; GS.Cells[SA_Credit,Lig]:=StrS(TCE,DECDEV) ;
   END else
   BEGIN
   }
   GS.Cells[SA_Debit,Lig]:=StrS(TDD,DECDEV) ; GS.Cells[SA_Credit,Lig]:=StrS(TCD,DECDEV) ;
//   END ;
FormatMontant(GS,SA_Debit,Lig,DECDEV) ; FormatMontant(GS,SA_Credit,Lig,DECDEV) ;
OBM:=GetO(GS,Lig) ;
if OBM<>Nil then
   BEGIN
   OBM.PutMvt('E_DEBIT',TDP)     ; OBM.PutMvt('E_CREDIT',TCP) ;
   OBM.PutMvt('E_DEBITDEV',TDD)  ; OBM.PutMvt('E_CREDITDEV',TCD) ;
//   OBM.PutMvt('E_DEBITEURO',TDE) ; OBM.PutMvt('E_CREDITEURO',TCE) ;
   OBM.HistoMontants ;
   END ;
TM:=TMOD(GS.Objects[SA_NumP,Lig]) ;
if TM<>Nil then if TM.ModR.TotalAPayerP<>0 then
   BEGIN
   for i:=1 to TM.ModR.NbEche do TM.ModR.TabEche[i].Pourc:=Arrondi(100.0*TM.ModR.TabEche[i].MontantP/TM.ModR.TotalAPayerP,ADecimP) ;
   END ;
END ;

procedure TFeSaisie.ChargeComptes ;
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
          if ((OkBQE) and ((CGen.NatureGene='BQE') or (CGen.NatureGene='CAI'))) then GeneTreso:=CGen.General ;
        {$IFNDEF SPEC350}
          if SAJAL<>Nil then if ((OkJALEFFET) and (CGen.General=SAJAL.Treso.Cpt)) then GeneTreso:=CGen.General ;
        {$ENDIF}
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

  FiniMove ;
END ;

procedure TFeSaisie.ChargeEcritures ;
Var Lig  : integer ;
    NumE,NbEche : integer ;
    OkE         : Boolean ;
    OBM         : TOBM ;
    TDD,TCD,TDP,TCP,TDE,TCE : Double ;
    LDD,LCD,LDP,LCP,LDE,LCE : Double ;
    TDJ,TCJ         : Double ;
    XDEL            : TDELMODIF ;
    stReq 					: String;
    QEcr						: TQuery;
BEGIN
if Action=taCreat then Exit ;
	stReq := 'Select * from Ecriture where ' + WhereEcriture(tsGene,M,False) ;
	stReq := stReq + ' ORDER BY E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE' ;
	QEcr := OpenSQL(stReq, True);
Lig:=0 ; TDJ:=0 ; TCJ:=0 ;
InitMove(50,'') ; NbLigEcr:=0 ; NbLigAna:=0 ; OkE:=False ;
TDD:=0 ; TCD:=0 ; TDP:=0 ; TCP:=0 ; TDE:=0 ; TCE:=0 ; NbEche:=0 ;
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
      if OkE then RuptureEche(Lig,TDD,TCD,TDP,TCP,TDE,TCE) ; OkE:=FALSE ; NbEche:=0 ;
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
{$IFDEF EAGLCLIENT}
  // ????
{$ELSE}
      if Not TMemoField(QEcr.FindField('E_BLOCNOTE')).IsNull
        then OBM.M.Assign(TMemoField(QEcr.FindField('E_BLOCNOTE'))) ;
{$ENDIF}
      OBM.CompS:=True ;
      TDD:=0 ; TCD:=0 ; TDP:=0 ; TCP:=0 ; TDE:=0 ; TCE:=0 ;
      GS.Cells[SA_Gen,Lig]:=QEcr.FindField('E_GENERAL').AsString ;
      GS.Cells[SA_Aux,Lig]:=QEcr.FindField('E_AUXILIAIRE').AsString ;
      GS.Cells[SA_RefI,Lig]:=QEcr.FindField('E_REFINTERNE').AsString ;
      GS.Cells[SA_Lib,Lig]:=QEcr.FindField('E_LIBELLE').AsString ;
      if EcheMvt(QEcr) then AlloueEche(Lig) ;
      END ;
   if ((Action=taModif) and (M.Simul='N') and (Not M.Valide) and (QEcr.FindField('E_TRESOLETTRE').AsString='X')) then TresoLettre:=True ;
   LDD:=QEcr.FindField('E_DEBITDEV').AsFloat  ; LCD:=QEcr.FindField('E_CREDITDEV').AsFloat ;
   LDP:=QEcr.FindField('E_DEBIT').AsFloat     ; LCP:=QEcr.FindField('E_CREDIT').AsFloat ;
//   LDE:=QEcr.FindField('E_DEBITEURO').AsFloat ; LCE:=QEcr.FindField('E_CREDITEURO').AsFloat ;
   TDD:=TDD+LDD ; TCD:=TCD+LCD ; TDP:=TDP+LDP ; TCP:=TCP+LCP ; TDE:=TDE+LDE ; TCE:=TCE+LCE ;
   TDJ:=TDJ+LDP ; TCJ:=TCJ+LCP ;
   if NbEche>=1 then AlimNoyauEche(Lig,(NbEche=1),LDD,LCD,LDP,LCP,LDE,LCE, QEcr) ;
   {
   if ModeOppose then
      BEGIN
      GS.Cells[SA_Debit,Lig]:=StrS(LDE,DECDEV) ; GS.Cells[SA_Credit,Lig]:=StrS(LCE,DECDEV) ;
      END else
      BEGIN
      }
      GS.Cells[SA_Debit,Lig]:=StrS(LDD,DECDEV) ; GS.Cells[SA_Credit,Lig]:=StrS(LCD,DECDEV) ;
//      END ;
   FormatMontant(GS,SA_Debit,Lig,DECDEV) ; FormatMontant(GS,SA_Credit,Lig,DECDEV) ;
   CaptureRatio(Lig,True) ; GereNewLigne(GS) ;
   YABienTP(QEcr) ;
   QEcr.Next ;
   END ;
Ferme(QEcr);
FiniMove ;
if NbEche>=1 then RuptureEche(Lig,TDD,TCD,TDP,TCP,TDE,TCE) ;
SAJAL.OldDebit:=TDJ ; SAJAL.OldCredit:=TCJ ;
//CommitTrans ;
END ;

procedure TFeSaisie.ChargeLignes ;
BEGIN
ChargeEcritures ;
ChargeComptes ;
TripoteYATP ;
if Action<>taCreat then BEGIN GS.Row:=1 ; GS.Col:=SA_Gen ; GS.SetFocus ; END ;
AffichePied ; NbLOrig:=GS.RowCount ;

END ;

procedure TFeSaisie.ChargeSoldesComptes ;
Var i,j,k 					: integer ;
    OBM 						: TOBM ;
    Gen,Aux,Sect,Ax : String17 ;
    D,C             : Double ;
BEGIN
	if PasModif then Exit ;
	for i:=1 to GS.RowCount-2 do
    BEGIN
    OBM:=GetO(GS,i) ;
    if OBM=Nil then Continue ;
    Gen	:= OBM.GetMvt('E_GENERAL') ;
    Aux	:= OBM.GetMvt('E_AUXILIAIRE') ;
    D		:= OBM.GetMvt('E_DEBIT') ;
    C		:= OBM.GetMvt('E_CREDIT') ;
    Ajoute(TS,'G',Gen,D,C) ;
    Ajoute(TS,'T',Aux,D,C) ;
    for j:=1 to Min( MaxAxe, OBM.Detail.Count ) do
      for k:=0 to OBM.Detail[j-1].Detail.Count-1 do
        BEGIN
        Ax		:= TOBM(OBM.Detail[j-1].Detail[k]).GetMvt('Y_AXE')  ;
        Sect	:= TOBM(OBM.Detail[j-1].Detail[k]).GetMvt('Y_SECTION') ;
        D			:= TOBM(OBM.Detail[j-1].Detail[k]).GetMvt('Y_DEBIT') ;
        C			:= TOBM(OBM.Detail[j-1].Detail[k]).GetMvt('Y_CREDIT') ;
        Ajoute(TS,Ax,Sect,D,C) ;
        END ;
    END ;
END ;


procedure TFeSaisie.ChargeAnals ;
Var NumL,NumV,NumA,OldL : integer ;
    OBA            : TOBM ;
    OBM			       : TOBM ;
    Ax             : String3 ;
    QAna           : TQuery ;
    X              : TDELMODIF ;
{$IFDEF EAGLCLIENT}
{$ELSE}
    T              : TStrings ;
{$ENDIF}
BEGIN
	if Action=taCreat then Exit ;
	QAna := OpenSQL('Select * from Analytiq where ' + WhereEcriture(tsAnal,M,False)
  								+ ' ORDER BY Y_JOURNAL, Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE, Y_AXE, Y_NUMVENTIL',True) ;

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
      if OBM.Detail.Count<MaxAxe then AlloueAnal(NumL) ;
//      AlimAnalDef(NumL,OBA) ;
//      OBA.ModeConf := QAna.FindField('Y_CONFIDENTIEL').AsString ;
      Revision:=False ;
      END ;

    Ax:='A'+Chr(48+NumA) ;
    OBA:=TOBM.Create(EcrAna,Ax,False,OBM.Detail[NumA-1]) ;
    OBA.ChargeMvt(QAna) ;
    OBA.HistoMontants ;

{$IFDEF EAGLCLIENT}
  // ????
{$ELSE}
    T:=TStringList.Create ;
    if Not TMemoField(QAna.FindField('Y_BLOCNOTE')).IsNull
    	then T.Assign(TMemoField(QAna.FindField('Y_BLOCNOTE'))) ;
{$ENDIF}

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

{==========================================================================================}
{=========================== Allocations et Désallocations ================================}
{==========================================================================================}
function  TFeSaisie.DateToJour ( DD : TDateTime ) : String ;
Var y,m,d : Word ;
    StD   : String ;
BEGIN
DecodeDate(DD,y,m,d) ; StD:=IntToStr(d) ; if d<10 then StD:='0'+StD ;
Result:=StD ;
END ;

procedure TFeSaisie.AlloueEcheAnal(C,L : LongInt) ;
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

procedure TFeSaisie.AlloueEche(LaLig : LongInt) ;
Var TTM : TMOD ;
begin
if GS.Objects[SA_NumP,LaLig]<>Nil then Exit ;
TTM:=TMOD.Create ; 
GS.Objects[SA_NumP,LaLig]:=TObject(TTM) ;
Revision:=False ;
end ;

procedure TFeSaisie.DesalloueEche(LaLig : LongInt) ;
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

procedure TFeSaisie.AlloueAnal(LaLig : LongInt) ;
var OBM : TOBM ;
begin
	OBM := GetO(GS,LaLig) ;
  if OBM = nil then Exit;
	AlloueAxe(OBM) ;
	Revision:=False ;
end ;

procedure TFeSaisie.DesalloueAnal(LaLig : LongInt) ;
var OBM : TOBM ;
begin
	OBM := GetO(GS,LaLig) ;
  if OBM.Detail.Count=0 then Exit ;
  OBM.ClearDetail;
	Revision:=False ;
end ;

{==========================================================================================}
{=========================== Initialisations et valeurs par défaut ========================}
{==========================================================================================}
procedure TFeSaisie.GereEnabled ( Lig : integer ) ;
Var OKL,Remp,LectSeul,Visu,OkRIB : Boolean ;
    O   : TOBM ;
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
BChancel.Enabled:=((GS.RowCount<=2) and (Not EstRempli(GS,1)) and (Action=taCreat) and (Not ModeLC) and (Not EstMonnaieIN(DEV.Code))) ;
BSwapPivot.Enabled:=(GS.RowCount>2) and ((ModeDevise) or (ModeForce=tmPivot) and (Not ModeLC)) ;
//BSwapEuro.Enabled:=((GS.RowCount>2) and (Not ModeLC)) ;
BSaisTaux.Enabled:=((GS.RowCount<=2) and (ModeDevise) and (Action=taCreat) and (Not ModeLC) and (Not EstMonnaieIN(DEV.Code))) ;
BModifTva.Enabled:=((OkL) and (OuiTvaLoc) and ((isHT(GS,Lig,True)) or (isTiers(GS,Lig)))) ;
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
END ;

procedure TFeSaisie.ReInitPiece ( Force : boolean ) ;
Var bc : boolean ;
BEGIN
V_PGI.IOError:=oeOk ; CEstGood:=False ; GuideAEnregistrer:=FALSE ; appeldatepourguide:=FALSE ; CodeGuideFromSaisie:='' ;
{Grid}
GS.VidePile(True) ; GS.RowCount:=2 ; FinGuide ;
if ((Action<>taCreat) and (Not Force)) then Exit ;
{Pointeurs}
// GG MEMCHECK
VideListe(TS) ; TS.Free ; TS:=Nil ; TS:=TList.Create ; VideListe(TECHEORIG) ;
// GG MEMCHECK
VideListe(TGUI) ; TGUI.Free ; TGUI:=Nil ; TGUI:=TList.Create ; GestionRIB:='...' ;
Scenario.Free ; Scenario:=Nil ; Scenario:=TOBM.Create(EcrSais,'',True) ;
MemoComp.Free ; MemoComp:=Nil ; MemoComp:=TStringList.Create ;
Entete.Free ; Entete:=Nil ; Entete:=TOBM.Create(EcrSais,'',True) ;
ModifSerie.Free ; ModifSerie:=Nil ; ModifSerie:=TOBM.Create(EcrGen,'',True) ;
OkScenario:=False ; if Action=taCreat then BEGIN ChargeScenario ; ChargeLibelleAuto ; END ;
{Inits}
BValide.Enabled:=True ; DefautPied ;
H_NUMEROPIECE.Visible:=False ; H_NUMEROPIECE_.Visible:=True ;
E_JOURNAL.Enabled:=True ; E_DATECOMPTABLE.Enabled:=True ; InitEnteteJal(False,False,True) ;
PieceModifiee:=False ; Revision:=False ; OkMessAnal:=False ; DejaRentre:=False ; UnChange:=False ;
GS.SetFocus ; GS.Col:=SA_Gen ; GS.Row:=1 ; ModifNext:=False ; Volatile:=False ;
VideListe(TDELECR) ; VideListe(TDELANA) ; ModeConf:='0' ; HConf.Visible:=False ; PieceConf:=False ;
YATP:=yaRien ; ListeTP.Clear ; RentreDate:=False ; RegimeForce:=False ;
LastTraiteCHQ:='' ; LastEche:=0 ;
{#TVAENC}
if ((M.Simul='N') or (M.Simul='S') or (M.Simul='R')) then OuiTvaLoc:=VH^.OuiTvaEnc ;
ExigeTva:=tvaMixte ; ExigeEntete:=tvaMixte ; ChoixExige:=False ;
BPopTva.ItemIndex:=Ord(ExigeEntete)+1 ;
ToutDebit:=True ; ToutEncais:=True ; GeneRegTva:='' ;
{Affichage}
if ((VH^.CPDateObli) and (Action=taCreat) and (E_DATECOMPTABLE.CanFocus)) then E_DATECOMPTABLE.SetFocus else
 if ((OkScenario) and (Scenario<>Nil) and (Action=taCreat) and (E_DATECOMPTABLE.CanFocus)) then E_DATECOMPTABLE.SetFocus else
   BEGIN
   if Action=taCreat then GSEnter(Nil) ;
   GSRowEnter(Nil,1,bc,False) ;
   END ;
END ;

procedure TFeSaisie.InitModifJal(Action : TActionFiche) ;
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

procedure TFeSaisie.InitEnteteJal(Totale,Zoom,ReInit : boolean) ;
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
      tzJVente  : BEGIN
      						E_NaturePiece.DataType := 'ttNatPieceVente'  ;
      						E_NaturePiece.Value := 'FC' ;
      						END ;
      tzJAchat  : BEGIN E_NaturePiece.DataType:='ttNatPieceAchat'  ; E_NaturePiece.Value:='FF' ; END ;
      tzJBanque : BEGIN
                    // BPY le 03/10/2003 : fiche de bug 12620
                    E_NaturePiece.DataType:='ttNatPieceBanque' ;
{$IFDEF CCMP}
                    if (not (VH^.CCMP.LotCli)) then E_NaturePiece.Value:='RF'
                    else
{$ENDIF}
                        E_NaturePiece.Value:='RC' ;
                    // fin BPY
                  END ;
 tzJEcartChange : BEGIN
                  E_NaturePiece.DataType:='ttNatPieceEcartChange' ;
                  if ((Action<>taCreat) or (SAJAL.MultiDevise)) then E_NaturePiece.Value:='ECC' else E_NATUREPIECE.Value:='OD' ;
                  END ;
      tzJOD     : BEGIN E_NaturePiece.DataType:='ttNaturePiece' ; E_NaturePiece.Value:='OD' ; END ;
      END ;
   END ;
if Not Zoom then
   BEGIN
   DD:=StrToDate(E_DATECOMPTABLE.Text) ;
   if M.Simul<>'N' then NumPieceInt:=GetNum(EcrGen,SAJAL.COMPTEURSIMUL,MM,DD)
                   else NumPieceInt:=GetNum(EcrGen,SAJAL.COMPTEURNORMAL,MM,DD) ;
   InitGrid ;
   END ;
if NumPieceInt>0 then E_NUMEROPIECE.Caption:=IntToStr(NumPieceInt) ;
END ;

procedure TFeSaisie.GereDeviseEcart ;
BEGIN
if ((E_NaturePiece.Value='') or (E_Journal.Value='')) then Exit ;
if E_NATUREPIECE.Value='ECC' then
   BEGIN
   E_DEVISE.Value:=V_PGI.DevisePivot ; E_DEVISEChange(Nil) ;
   E_DEVISE_.Visible:=True ; E_DEVISE.Visible:=False ;
   if ((E_DEVISE_.Value='') and (E_DEVISE_.Items.Count>0)) then E_DEVISE_.ItemIndex:=0 ;
   END else
   BEGIN
   E_DEVISE.Visible:=True ; E_DEVISE_.Visible:=False ; E_DEVISE_.Value:=V_PGI.DevisePivot ;
   END ;
END ;

procedure TFeSaisie.DefautEntete ;
BEGIN
if ((VH^.CPExoRef.Code<>'') and (DateSais>0) and (Action=taCreat)
and (ctxPCL in V_PGI.PGIContexte)) and (M.TypeGuide<>'ENC') and (M.TypeGuide<>'DEC') then M.DateC:=DateSais ;
E_DATECOMPTABLE.Text:=DateToStr(M.DateC) ;
E_JOURNAL.Value:=M.JAL ; E_NATUREPIECE.Value:=M.Nature ;
E_ETABLISSEMENT.Value:=M.Etabl ; E_ETABLISSEMENT.Enabled:=VH^.EtablisCpta ;
if Action=taCreat then PositionneEtabUser(E_ETABLISSEMENT) ;
//ColorOpposeEuro(GS,ModeOppose,ModeDevise) ;
DatePiece:=M.DateC ;
if M.Nature='ECC' then
   BEGIN
   E_DEVISE.Value:=V_PGI.DevisePivot ;
   E_DEVISE_.Value:=M.CodeD ;
   END else
   BEGIN
   {
   if ((ModeOppose) and (M.CodeD=V_PGI.DevisePivot)) then
      BEGIN
      E_DEVISE.ItemIndex:=E_DEVISE.Values.Count-1 ; E_DEVISEChange(Nil) ;
      END else
      BEGIN
      }
      E_DEVISE.Value:=M.CodeD ;
      //END ;
   E_DEVISE_.Value:=V_PGI.DevisePivot ;
   END ;
Case Action of
   taCreat : BEGIN
             E_DEVISE.Enabled:=False ; E_NumeroPiece.Caption:='' ;
             H_NUMEROPIECE.Visible:=False ; H_NUMEROPIECE_.Visible:=True ;
             GS.Enabled:=False ;
             END ;
   taModif : BEGIN
             E_JOURNAL.Enabled:=False ; E_DEVISE.Enabled:=False ; E_DEVISE_.Enabled:=False ;
             E_DATECOMPTABLE.Enabled:=False ; E_NATUREPIECE.Enabled:=False ;
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

procedure TFeSaisie.DefautPied ;
BEGIN
SI_TotDS:=0 ; SI_TotCS:=0 ;
SI_TotDD:=0 ; SI_TotCD:=0 ;
SI_TotDP:=0 ; SI_TotCP:=0 ;
SI_TotDE:=0 ; SI_TotCE:=0 ;
ZeroBlanc(PPied) ;
END ;

procedure TFeSaisie.InitGrid ;
BEGIN
if Action<>taCreat then Exit ;
GS.Enabled:=True ; DefautLigne(GS.RowCount-1,True) ;
GS.Col:=SA_Gen ; GS.Row:=1 ;
END ;

procedure TFeSaisie.AttribLeTitre ;
Var i,j : integer ;
    C   : Char ;
BEGIN
i:=1 ; j:=1 ;
Case Action of taConsult : i:=1 ; taModif : i:=2 ; taCreat : i:=3 ; END ;
C:=M.Simul[1] ; Case C of 'N' : j:=1 ; 'S' : j:=2 ; 'P' : j:=3 ; 'U' : j:=4 ; 'R' : j:=5 ; END ;
if M.ANouveau then j:=6 ;
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
                  END ;
               END ;
   END ;
if SuiviMP then Caption:=HTitres.Mess[26] ;
UpdateCaption(Self) ;
END ;

procedure TFeSaisie.AttribNumeroDef ;
Var Facturier  : String3 ;
    Lig,NumAxe,NumV : integer ;
    O          : TOBM ;
    DD         : TDateTime ;
BEGIN
	if ((Action<>taCreat) and (Not SuiviMP)) then Exit ;
	if ((M.Simul<>'N') and (Not SuiviMP))
  	then Facturier:=SAJAL.CompteurSimul
    else Facturier:=SAJAL.CompteurNormal ;
  DD:=StrToDate(E_DATECOMPTABLE.Text) ;
	SetIncNum(EcrGen,Facturier,NumPieceInt,DD) ;
	H_NUMEROPIECE.Visible		:=True ;
  H_NUMEROPIECE_.Visible	:=False ;
	E_NUMEROPIECE.Caption		:=IntToStr(NumPieceInt) ;
	M.LastNumCreat					:=NumPieceInt ;
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
procedure TFeSaisie.LettrageEnSaisie ;
var X : RMVT ;
    Q : TQuery ;
    Nb  : integer ;
begin
{$IFNDEF GCGC}
  if Action=taConsult then Exit ;
  if M.Simul<>'N' then Exit ;
  X.Axe:='' ;
  X.Etabl:=E_ETABLISSEMENT.Value ;
  X.Jal:=E_JOURNAL.Value ;
  X.Exo:=QuelExoDT(StrToDate(E_DATECOMPTABLE.Text)) ;
  X.CodeD:=E_DEVISE.Value ;
  X.Simul:=M.Simul ;
  X.Nature:=E_NATUREPIECE.Value ;
  X.DateC:=StrToDate(E_DATECOMPTABLE.Text) ;
  X.DateTaux:=DEV.DateTaux ;
  X.Num:=NumPieceInt ;
  X.TauxD:=DEV.Taux ;
  X.Valide:=False ;
  X.ANouveau:=False ;
  Q:=OpenSQL('Select Count(*) from ECRITURE Where '+WhereEcriture(tsGene,X,False)+' AND E_ECHE="X" AND E_NUMECHE>0 AND E_ETATLETTRAGE="AL"',True) ;
  if Not Q.EOF then
    begin
    Nb:=Q.Fields[0].AsInteger ;
    if Nb>0 then
      begin
      if HPiece.Execute(24,caption,'')=mrYes then
        LettrerEnSaisie(X,Nb) ;
      end ;
    end ;
  Ferme(Q) ;
{$ENDIF}
END ;

procedure TFeSaisie.ChargeScenario ;
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
if Not Q.EOF then
   BEGIN
   Scenario.ChargeMvt(Q) ;
   if Not TMemoField(Q.FindField('SC_ATTRCOMP')).IsNull then
{$IFDEF EAGLCLIENT}
   	 MemoComp.Text := (TMemoField(Q.FindField('SC_ATTRCOMP')).AsString) ; // EAGL : CELA CONVIENT-IL ?
{$ELSE}
   	 MemoComp.Assign(TMemoField(Q.FindField('SC_ATTRCOMP'))) ;
{$ENDIF}
   OkScenario:=True ;
   END ;
Ferme(Q) ;
if OkScenario then
   BEGIN
   StComp:=Scenario2Comp(Scenario) ; Entete.PutMvt('E_ETAT',StComp) ;
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
procedure TFeSaisie.E_NATUREPIECEChange(Sender: TObject);
begin
if E_NaturePiece.Value='' then Exit ;
if ((SAJAL<>Nil) and (Action=taCreat) and (E_NATUREPIECE.Value='ECC')) then if Not SAJAL.MultiDevise
   then BEGIN E_NATUREPIECE.Value:='' ; HPiece.Execute(33,caption,'') ; Exit ; END ;
PieceModifiee:=True ; AutoCharged:=False ; Revision:=False ;
ChargeLibelleAuto ; GereDeviseEcart ;
ChargeScenario ;
GetSorteTva ; AfficheExigeTva ;
end;

procedure TFeSaisie.E_DATECOMPTABLEChange(Sender: TObject);
begin
PieceModifiee:=True ;
end;

procedure TFeSaisie.E_ETABLISSEMENTChange(Sender: TObject);
begin
PieceModifiee:=True ;
if ((E_ETABLISSEMENT.Value='') or (E_ETABLISSEMENT.Value=M.Etabl)) then Exit ;
ChargeScenario ;
end;

procedure TFeSaisie.E_JOURNALChange(Sender: TObject);
Var Jal : String ;
begin
Jal:=E_JOURNAL.Value ;
if Jal='' then BEGIN SAJAL.Free ; SAJAL:=Nil ; Exit ; END ;
if SAJAL<>Nil then if SAJAL.JOURNAL=Jal then Exit ;
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
END ;

procedure TFeSaisie.AvertirPbTaux(Code : String3 ; DateTaux,DateCpt : TDateTime) ;
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
if ((EstMonnaieIN(Code)) and (DateCpt>=V_PGI.DateDebutEuro)) then
   BEGIN
   ii:=HDiv.Execute(14,Caption,'') ;
   if ii<>mrYes then ii:=HDiv.Execute(15,caption,'') ;
   if ii=mrYes then
      BEGIN
      FicheDevise( Code, taModif, False ) ;
      DEV.Taux:=GetTaux(DEV.Code,DEV.DateTaux,DateCpt) ;
      END ;
   END else
   BEGIN
   if OkTaux1 then ii:=HDiv.Execute(8,caption,'') else ii:=HDiv.Execute(2,caption,'') ;
   if ii=mrYes then
      BEGIN
      FicheChancel(E_DEVISE.VALUE,True,DateCpt,taCreat,(DateCpt>=V_PGI.DateDebutEuro)) ; 	// EAGL : EN ATTENTE
      DEV.Taux:=GetTaux(DEV.Code,DEV.DateTaux,DateCpt) ;
      END ;
   END ;
end ;

Function TFeSaisie.PbTaux ( DEV : RDevise ; DateCpt : TDateTime ) : boolean ;
Var Code : string3 ;
BEGIN
Result:=False ;
Code:=DEV.Code ;
if ((Code=V_PGI.DevisePivot) or (Code=V_PGI.DeviseFongible)) then Exit ;
if ((DateCpt<V_PGI.DateDebutEuro) or (Not EstMonnaieIn(Code))) then Result:=(DEV.Taux=1) else
   if EstMonnaieIn(Code) then Result:=(DEV.Taux=V_PGI.TauxEuro) ;
END ;

procedure TFeSaisie.E_DEVISEChange(Sender: TObject);
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
                     if ((DEV.Code=V_PGI.DevisePivot) or (EstMonnaieIN(DEV.Code))) then Volatile:=False ;
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
{
SetOpposeEuro ;
if ModeOppose then BEGIN DEV.Decimale:=V_PGI.OkDecE ; DEV.Symbole:='' ; END else
 if ((Not ModeOppose) and (DEV.Code=V_PGI.DevisePivot)) then DEV.Symbole:=V_PGI.SymbolePivot ;
 }
DecDev:=DEV.Decimale ;
ChangeFormatDevise(Self,DECDEV,DEV.Symbole) ;
end;

procedure TFeSaisie.H_JOURNALDblClick(Sender: TObject);
begin ClickJournal ; end;

procedure TFeSaisie.H_ETABLISSEMENTDblClick(Sender: TObject);
begin ClickEtabl ; end;

procedure TFeSaisie.QuelNExo ;
Var EXO : String ;
BEGIN
if Action=taCreat then EXO:=QuelExoDT(StrToDate(E_DATECOMPTABLE.Text)) else EXO:=M.EXO ;
if EXO=VH^.Encours.Code then H_EXERCICE.Caption:='(N)' else
 if EXO=VH^.Suivant.Code then H_EXERCICE.Caption:='(N+1)' else
  H_EXERCICE.Caption:='(N-1)' ;
END ;

procedure TFeSaisie.E_DATECOMPTABLEEnter(Sender: TObject);
begin
RentreDate:=True ;
end;

procedure TFeSaisie.E_DATECOMPTABLEExit(Sender: TObject);
Var Err,i   : integer ;
    DateCpt : TDateTime ;
//    NewNumPieceInt : Integer ;
//    MM : String17 ;
begin
if csDestroying in ComponentState then Exit ;
Err:=ControleDate(E_DATECOMPTABLE.Text) ;
if Err>0 then BEGIN HPiece.Execute(15+Err,caption,'') ; E_DATECOMPTABLE.SetFocus ; Exit ; END ;
if ((Action=taCreat) and (RevisionActive(StrToDate(E_DATECOMPTABLE.Text)))) then
   BEGIN
   E_DATECOMPTABLE.SetFocus ;
   Exit ;
   END ;
(*
if ((Action=taCreat) and (ModeOppose) and (Not VH^.TenueEuro) and
    (StrToDate(E_DATECOMPTABLE.Text)<V_PGI.DateDebutEuro)) then
   BEGIN
   HPiece.Execute(39,Caption,'') ;
   E_DATECOMPTABLE.SetFocus ;
   Exit ;
   END ;
if EstSpecif('51188') Then
  BEGIN
  If (Action=taCreat) And StopDevise(StrToDate(E_DATECOMPTABLE.Text),DEV.Code,ModeOppose) Then
    BEGIN
     HPiece.Execute(53,Caption,'') ;
     E_DATECOMPTABLE.SetFocus ;
     Exit ;
    END ;
  END ;
*)
QuelNExo ; DatePiece:=StrToDate(E_DATECOMPTABLE.Text) ;
{$IFDEF CEGID}
If MethodeGuideDate1 Then
  BEGIN
  if ( appeldatepourguide And (Not M.FromGuide) And (((M.TypeGuide='NOR') and (M.SaisieGuidee)) Or (CodeGuideFromSaisie<>''))) then
    BEGIN
    GS.Enabled:=TRUE ; GS.SynEnabled:=TRUE ; GS.SetFocus ; Exit ;
    END ;
  END ;
{$ENDIF}
if ((DEV.Code<>V_PGI.DevisePivot) and (Action=taCreat) and (Not Volatile)) Then
   BEGIN
   DateCpt:=DatePiece ;
   DEV.Taux:=GetTaux(DEV.Code,DEV.DateTaux,DateCpt) ;
   If ((DEV.DateTaux<>DateCpt) or (PbTaux(DEV,DateCpt))) Then AvertirPbTaux(DEV.Code,DEV.DateTaux,DateCpt) ;
   END ;
for i:=1 to GS.RowCount-1 do GS.Cells[SA_DateC,i]:=DateToJour(StrToDate(E_DATECOMPTABLE.EditText)) ;
(** JLD 04/06/99
if SAJAL<>Nil then
   BEGIN
   if M.Simul<>'N' then NewNumPieceInt:=GetNum(EcrGen,SAJAL.COMPTEURSIMUL,MM,DatePiece)
                   else NewNumPieceInt:=GetNum(EcrGen,SAJAL.COMPTEURNORMAL,MM,DatePiece) ;
   if (NewNumPieceInt>0) And (NewNumPieceInt<>NumPieceInt) then
     BEGIN
     NumPieceInt:=NewNumPieceInt ;
     E_NUMEROPIECE.Caption:=IntToStr(NumPieceInt) ;
     END ;
   END ;
**)
end;

{==========================================================================================}
{=================================== ANALYTIQUES ==========================================}
{==========================================================================================}
Function TFeSaisie.AOuvrirAnal ( Lig : integer ; Auto : boolean ) : Boolean ;
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

Procedure TFeSaisie.GereAnal ( Lig : integer ; Auto,Scena : Boolean ) ;
BEGIN
{$IFDEF CEGID}
	if Guide And ((ValD(GS,Lig)=0) and (ValC(GS,Lig)=0)) then Exit ;
{$ENDIF}
	if AOuvrirAnal(Lig,Auto) then OuvreAnal(Lig,Scena) ;
	CalculDebitCredit ;
END ;

Procedure TFeSaisie.OuvreAnal ( Lig : integer ; Scena : boolean ) ;
Var XD,XP,VV 	: Double ;
    C,NumAxe 			: integer ;
    Ax       			: String ;
		ArgAna 				: ARG_ANA;
    OBM 					: TOBM;
    i							:	Integer;
BEGIN
	// Mise en place des arguments de lancement de la saisie analytique
  ArgAna.QuelEcr        := EcrGen ;
  if PasToucheLigne(Lig)
  	then ArgAna.Action	:= taConsult
    else ArgAna.Action	:= Action ;
  if Guide
  	then ArgAna.GuideA	:= GuideAnal
    else ArgAna.GuideA	:= '' ;
  ArgAna.DernVentilType	:= DernVentilType ;
	ArgAna.ControleBudget := ((OkScenario) and (Action<>taConsult) and (Scenario.GetMvt('SC_CONTROLEBUDGET')='X'));
  ArgAna.ModeConf				:= ModeConf ;
  ArgAna.CC							:= GetGGeneral(GS,Lig) ;
  ArgAna.DEV						:= DEV ;
  ArgAna.Verifventil		:= False ;
	ArgAna.NumLigneDecal  := 0;
  // - Pour les scénario : recherche de l'axe
	NumAxe:=0 ;
	if ((OkScenario) and (Scena) and (Scenario.GetMvt('SC_OUVREANAL')<>'X')) then
    BEGIN
    Ax:=Trim(Scenario.GetMvt('SC_NUMAXE')) ;
    if Length(Ax)<2 then Exit ;
    NumAxe:=Ord(Ax[2])-48 ;
    END ;
  ArgAna.NumGeneAxe			:= NumAxe ;

	// Backup des montants
  XD:=MontantLigne(GS,Lig,tsmDevise) ;
  XP:=MontantLigne(GS,Lig,tsmPivot) ;

	// Mise en place des Axes dans l'OBM si besoin
  OBM := GetO(GS,Lig);
  for i:=(OBM.Detail.Count+1) to MaxAxe do
  	begin
    TOB.Create('A'+IntToStr(i),OBM,-1);
    end;

	// Appel de la saisie analytique
	eSaisieAnal(TOB(OBM),ArgAna) ;

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
Procedure TFeSaisie.GereEche ( Lig : integer ; OuvreAuto,RempliAuto : Boolean ) ;
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

Procedure TFeSaisie.EcheGuideRegle ( Lig : integer ; Var T : T_MODEREGL ) ;
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

Procedure TFeSaisie.EcheLettreCheque ( Lig : integer ; Var T : T_MODEREGL ) ;
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

Function TFeSaisie.IncNumLotSaisie ( Lig : integer ) : String ;
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

Procedure TFeSaisie.RemonteValeurTraCHQ ( sTraCHQ : String ; Lig : integer ) ;
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

Procedure TFeSaisie.GereMonoEche ( Var OO : T_ModeRegl ; OkInit,AutoPourEncaDeca : boolean ; Lig : integer ; ActionEche : TActionFiche ) ;
Var X : T_MONOECH ;
    OkOk : Boolean ;
    O    : TOBM ;
BEGIN
//OkOk:=FALSE ;
O:=GetO(GS,Lig) ;
With OO do
   BEGIN
   X.DateMvt:=DateMvt(Lig) ;
   if LastPaie<>'' then X.ModePaie:=LastPaie else X.ModePaie:=TabEche[1].ModePaie ;
   if LastEche>0 then X.DateEche:=LastEche else X.DateEche:=TabEche[1].DateEche ;
   X.Cat:='' ; X.Treso:=False ; X.OkInit:=OkInit ;
   if OkInit then TabEche[1].DateValeur:=CalculDateValeur(DateMvt(Lig),MonoEche,GS.Cells[SA_Gen,Lig]) ;
   if LastVal>0 then X.DateValeur:=LastVal else X.DateValeur:=TabEche[1].DateValeur ;
   X.NumTraiteCHQ:='' ;
{$IFNDEF SPEC350}
   if O<>Nil then X.NumTraiteCHQ:=O.GetMvt('E_NUMTRAITECHQ') ;
   if ((X.NumTraiteCHQ='') and (LastTraiteCHQ<>'') and (OkBQE)) then X.NumTraiteCHQ:=LastTraiteCHQ ;
   if ((X.NumTraiteCHQ='') and (OkJalEffet) and (O<>Nil) and (O.GetMvt('E_GENERAL')=SAJAL.Treso.Cpt)) then X.NumTraiteCHQ:=IncNumLotSaisie(Lig) ;
{$ENDIF}
   X.OkVal:=True ; X.Action:=Action ;
   If AutoPourEncaDeca Then
    BEGIN
    If (X.ModePaie='') And (Not Guide)
      Then OkOk:=SaisirMonoEcheance(X)
      Else OkOk:=TRUE ;
    END
   Else OkOk:=SaisirMonoEcheance(X) ;
   if OkOk then
      BEGIN
      TabEche[1].DateEche:=X.DateEche ; TabEche[1].ModePaie:=X.ModePaie ;
      NbEche:=1 ; TabEche[1].Pourc:=100.0 ;
      TabEche[1].MontantP:=TotalAPayerP ;
      TabEche[1].MontantD:=TotalAPayerD ;
      TabEche[1].DateValeur:=X.DateValeur ;
      if TabEche[1].EtatLettrage='' then TabEche[1].EtatLettrage:='AL' ;
      LastPaie:=X.ModePaie ; LastEche:=X.DateEche ; LastVal:=X.DateValeur ;
{$IFNDEF SPEC350}
      if O<>Nil then O.PutMvt('E_NUMTRAITECHQ',X.NumTraiteCHQ) ;
      if X.NumTraiteCHQ<>'' then LastTraiteCHQ:=X.NumTraiteCHQ ;
      RemonteValeurTraCHQ(X.NumTraiteCHQ,Lig) ;
{$ENDIF}
      END ;
   END ;
END ;

Procedure TFeSaisie.OuvreEche ( Lig : integer ; Cpte : String17 ; MR : String3 ; RempliAuto : boolean ; t : TLettre ) ;
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
   if ((Self.Action=taConsult) or (ModeLC)) then OMODR.MODR.Action:=taConsult else
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
      END else
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

Function TFeSaisie.AOuvrirEche ( Lig : integer ; Var Cpte : String17 ; Var MR : String3 ;
                                Var OuvreAuto,RempliAuto : boolean ; Var t : TLettre) : Boolean ;
Var CGen     : TGGeneral ;
    CAux     : TGTiers ;
    OMODR    : TMOD ;
    XD,XP    : Double ;
    OG       : TOBM ;
    St,StR   : String ;
BEGIN
Result:=False ; Cpte:='' ; MR:='' ; t:=NonEche ;
CGen:=GetGGeneral(GS,Lig) ; if CGen=Nil then exit ;
CAux:=GetGTiers(GS,Lig) ;
if CGen.Collectif then
   BEGIN
   if CAux=Nil then Exit ; if Not CAux.Lettrable then Exit ;
   Cpte:=CAux.Auxi ; MR:=CAux.ModeRegle ;
   if Ventilable(CGen,0) then t:=MonoEche else t:=MultiEche ;
   END else
   BEGIN
   t:=Lettrable(CGen) ; if t=NonEche then Exit ;
   if Ventilable(CGen,0) then t:=MonoEche ;
   Cpte:=CGen.General ; MR:=CGen.ModeRegle ;
   END ;
if Guide then
   BEGIN
   OG:=TOBM(TGUI[Lig-1]) ; if OG=Nil then Exit ;
   StR:=OG.GetMvt('EG_MODEREGLE') ; St:=OG.GetMvt('EG_ARRET') ;
   if (M.LeGuide='') or ((M.TypeGuide<>'DEC') and (M.TypeGuide<>'ENC')) then
      BEGIN
      if StR<>'' then BEGIN MR:=StR ; if St[7]='X' then OuvreAuto:=True else RempliAuto:=True ; END ;
      END else
      BEGIN
      {guides de règlement}
      OuvreAuto:=False ; RempliAuto:=True ;
      END ;
   END ;
if OuvreAuto then
   BEGIN
   OMODR:=TMOD(GS.Objects[SA_NumP,Lig]) ; if OMODR=Nil then Exit ;
   XD:=MontantLigne(GS,Lig,tsmDevise) ; XP:=MontantLigne(GS,Lig,tsmPivot) ; 
   if ((OMODR.MODR.TotalAPayerD=XD)  and (OMODR.MODR.TotalAPayerP=XP)) then Exit ;
   if OMODR.MODR.Aux=Cpte then BEGIN RecalculProrataEche(OMODR.MODR,XP,XD) ; Exit ; END ;
   END ;
Result:=True ;
END ;

{==========================================================================================}
{=================================== Contrôles ============================================}
{==========================================================================================}
Procedure TFeSaisie.TripoteYATP ;
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

function TFeSaisie.PasToucheLigne ( Lig : integer ) : boolean ;
Var OBM : TOBM ;
    TM  : TMOD ;
    k   : integer ;
    EtatLig,RefP,EtatMod,StEtatTva : String ;
    ExisteL,XEditeTva,ExportePCL : boolean ;
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
Result:=((EtatLig='PL') or (EtatLig='TL') or (RefP<>'') or (ExisteL) or (XEditeTva) or (ExportePCL)) ;
END ;

function TFeSaisie.ControleTVAOK : boolean ;
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

procedure TFeSaisie.ControleLignes ;
Var Lig : integer ;
BEGIN
for Lig:=1 to GS.RowCount-1 do if Not EstRempli(GS,Lig) then DefautLigne(Lig,True) ;
END ;

procedure TFeSaisie.AjusteLigne ( Lig : integer ; AvecM : boolean ) ;
Var Cpte              : String17 ;
    MR                : String3 ;
    t                 : TLettre ;
    Ouv,Remp          : boolean ;
BEGIN
AlimObjetMvt(Lig,AvecM) ; Remp:=False ; Ouv:=True ;
if AOuvrirEche(Lig,Cpte,MR,Ouv,Remp,t) then ;
if AOuvrirAnal(Lig,True) then ;
END ;

procedure TFeSaisie.ErreurSaisie ( Err : integer ) ;
BEGIN
if Err<100 then
   BEGIN
   HLigne.Execute(Err-1,caption,'') ;
   END else if Err<200 then
   BEGIN
   HPiece.Execute(Err-101,caption,'') ;
   END ;
END ;

Function TFeSaisie.PossibleRecupNum : Boolean ;
Var MM : String17 ;
    Facturier : String3 ;
    DD : TDateTime ;
BEGIN
Result:=True ;
if H_NUMEROPIECE_.Visible then Exit ;
if M.Simul<>'N' then Facturier:=SAJAL.CompteurSimul else Facturier:=SAJAL.CompteurNormal ;
DD:=StrToDate(E_DATECOMPTABLE.Text) ;
if GetNum(EcrGen,Facturier,MM,DD)=NumPieceInt+1 then Exit ;
Result:=False ;
END ;

Function TFeSaisie.PasModif : Boolean ;
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
Function TFeSaisie.TraiteDoublon ( Lig : integer ) : Boolean ;
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
SQL:='Select * from ECRITURE Where E_GENERAL="'+Gene+'" AND E_AUXILIAIRE="'+Aux+'" AND E_QUALIFPIECE="N" AND E_EXERCICE="'+O.GetMvt('E_EXERCICE')+'"' ;
Champ:='' ; Val:='' ;
if VH^.CPChampDoublon='RFI' then BEGIN Champ:='E_REFINTERNE' ; Val:=O.GetMvt('E_REFINTERNE') ; END else
if VH^.CPChampDoublon='LIB' then BEGIN Champ:='E_LIBELLE'    ; Val:=O.GetMvt('E_LIBELLE')    ; END else
if VH^.CPChampDoublon='RFX' then BEGIN Champ:='E_REFEXTERNE' ; Val:=O.GetMvt('E_REFEXTERNE') ; END else
if VH^.CPChampDoublon='RFL' then BEGIN Champ:='E_REFLIBRE'   ; Val:=O.GetMvt('E_REFLIBRE')   ; END ;
if ((Champ='') or (Val='')) then Exit ;
SQL:=SQL+' AND '+Champ+'="'+Val+'"' ;
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

Function TFeSaisie.LigneCorrecte ( Lig : integer ; Totale,Alim : boolean ) : Boolean ;
Var CGen  : TGGeneral ;
    CAux  : TGTiers ;
    Err,NumA,NumL : integer ;
    OMOD  : TMOD ;
    OBM   : TOBM ;
    OBA   : TOBM ;
    TotD,TotP{,TotE} : Double ;
    Sens,ii{,SensC}  : byte ;
    Col       : integer ;
BEGIN
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
{$IFDEF CEGID}
If Not Guide Then
  BEGIN
  if ((Err=0) and (ValD(GS,Lig)=0) and (ValC(GS,Lig)=0)) then
     if Not EstCptEcart(CGen.General) then
     BEGIN
     Err:=5 ; if QuelSens(E_NATUREPIECE.Value,CGen.NatureGene,CGen.Sens)=1 then Col:=SA_Debit else Col:=SA_Credit ;
     END ;
  END ;
{$ELSE}
if ((Err=0) and (ValD(GS,Lig)=0) and (ValC(GS,Lig)=0)) then
   if Not EstCptEcart(CGen.General) then
   BEGIN
   Err:=5 ; if QuelSens(E_NATUREPIECE.Value,CGen.NatureGene,CGen.Sens)=1 then Col:=SA_Debit else Col:=SA_Credit ;
   END ;
{$ENDIF}
{$IFDEF CEGID}
If Not Guide Then
  BEGIN
  if ((Err=0) and (ModeDevise) and (ValeurPivot(GS.Cells[SA_Debit,Lig],DEV.Taux,Dev.Quotite)=0)
              and (ValeurPivot(GS.Cells[SA_Credit,Lig],DEV.Taux,Dev.Quotite)=0)) then
      BEGIN
      Err:=6 ; if Sens=1 then Col:=SA_Debit else Col:=SA_Credit ;
      END ;
  END ;
{$ELSE}
if ((Err=0) and (ModeDevise) and (ValeurPivot(GS.Cells[SA_Debit,Lig],DEV.Taux,Dev.Quotite)=0)
            and (ValeurPivot(GS.Cells[SA_Credit,Lig],DEV.Taux,Dev.Quotite)=0)) then
    BEGIN
    Err:=6 ; if Sens=1 then Col:=SA_Debit else Col:=SA_Credit ;
    END ;
{$ENDIF}
if ((Err=0) and (MontantLigne(GS,Lig,tsmDevise)<>0) and
    (MontantLigne(GS,Lig,tsmPivot)=0) and (Not EstCptEcart(CGen.General))) then BEGIN Err:=6 ; if Sens=1 then Col:=SA_Debit else Col:=SA_Credit ; END ;
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
//            if Arrondi(TotE-MontantLigne(GS,Lig,tsmEuro),V_PGI.OkDecE)<>0 then BEGIN Err:=11+NumA ; Break ; END ;
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
     (Not ModeLC) and (VH^.CPChampDoublon<>'')) then
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
   if Alim
		 then AlimObjetMvt(Lig,TRUE) ;
   END ;
END ;

Function TFeSaisie.Equilibre : Boolean ;
BEGIN
Result:=ArrS(SI_TotDS-SI_TotCS)=0 ;
END ;

Function TFeSaisie.ControleRIB : Boolean ;
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

Procedure TFeSaisie.AjusteLigneEuro ( GS : THGrid ; k : integer ) ;
Var O : TOBM ;
    XD,XC{,ED,EC} : Double ;
BEGIN
O:=GetO(GS,k) ;
{if ModeOppose then
   BEGIN
   ED:=O.GetMvt('E_DEBITEURO') ; EC:=O.GetMvt('E_CREDITEURO') ;
   GS.Cells[SA_Debit,k]:=StrS(ED,DECDEV) ; GS.Cells[SA_Credit,k]:=StrS(EC,DECDEV) ;
   END else
   BEGIN}
   XD:=O.GetMvt('E_DEBITDEV') ; XC:=O.GetMvt('E_CREDITDEV') ;
   GS.Cells[SA_Debit,k]:=StrS(XD,DECDEV) ; GS.Cells[SA_Credit,k]:=StrS(XC,DECDEV) ;
//   END ;
FormatMontant(GS,SA_Debit,k,DECDEV) ; FormatMontant(GS,SA_Credit,k,DECDEV) ;
TraiteMontant(k,False) ;
AjusteLigne(k,False) ;
END ;

Function TFeSaisie.EquilEuroValide : Boolean ;
Var EcartP,EcartE : Double ;
    LaLig,k : integer ;
    LaRegle : String ;
BEGIN
Result:=True ;
if Not EuroOK then Exit ;
if PasModif then Exit ;
{$IFNDEF GCGC}
EcartP:=Arrondi(SI_TotDP-SI_TotCP,V_PGI.OkDecV) ; //EcartE:=Arrondi(SI_TotDE-SI_TotCE,V_PGI.OkDecE) ;
if ((EcartP=0) and (EcartE=0)) then Exit ;
if DatePiece<V_PGI.DateDebutEuro then BEGIN SoldeLaLigne(1) ; Exit ; END ;
if DEV.Code<>V_PGI.DevisePivot then BEGIN SoldeLaLigne(1) ; Exit ; END ;
{$ENDIF}
END ;

Function TFeSaisie.PieceCorrecte : Boolean ;
Var i,Err : integer ;
    OkokT : boolean ;
BEGIN
if PasModif then BEGIN Result:=True ; Exit ; END ;
{Result:=False ;} Err:=0 ; OkokT:=False ;
{$IFDEF SPEC350}
if ((OkBQE)) then
{$ELSE}
if ((OkBQE) or (OkJalEffet)) then
{$ENDIF}
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
   if ((OkBQE) and (GeneTreso<>'') and (GS.Cells[SA_Gen,i]=GeneTreso)) then OkokT:=True ;
{$IFNDEF SPEC350}
   if ((OkJalEffet) and (GeneTreso<>'') and (GS.Cells[SA_Gen,i]=GeneTreso)) then OkokT:=True ;
{$ENDIF}
   END ;
if ((Err=0) and (OkBQE) and (Not OkokT)) then Err:=126 ;
{$IFNDEF SPEC350}
If Not SuiviMP Then BEGIN if ((Err=0) and (OkJalEffet) and (Not OkokT)) then Err:=151 ; END ;
{$ENDIF}
Result:=(Err=0) ;
if not Result then
   BEGIN
   if ((ModeDevise) and (Err=103)) then ForcerMode(True,0) else ErreurSaisie(Err) ;
   END ;
END ;

{==========================================================================================}
{========================= Traitements de calcul liés aux lignes ==========================}
{==========================================================================================}
procedure TFeSaisie.PutScenar ( O : TOBM ; Lig : integer ) ;
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

procedure TFeSaisie.DefautLigne ( Lig : integer ; Init : boolean ) ;
Var O  : TOBM ;
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
//O.PutMvt('E_SAISIEEURO',CheckToString(ModeOppose)) ;
if ((OkScenario) and (Entete<>Nil) and (Init)) then PutScenar(O,Lig) ;
END ;

procedure TFeSaisie.DetruitLigne( Lig : integer ; Totale : boolean ) ;
Var R : integer ;
BEGIN
GS.CacheEdit ;
if Totale then R:=GS.Row else R:=1 ; GS.DeleteRow(Lig) ;
CalculDebitCredit ; NumeroteLignes(GS) ; NumeroteVentils ;
GS.SetFocus ; if R>=GS.RowCount then R:=GS.RowCount-1 ;
if R>1 then GS.Row:=R else GS.Row:=1 ;
GS.Col:=SA_Gen ; GereNewLigne(GS) ; GS.Invalidate ;
GS.MontreEdit ;
END ;

procedure TFeSaisie.TraiteMontant ( Lig : integer ; Calc : boolean ) ;
Var XC,XD{,SD,SC} : Double ;
    OBM         : TOBM ;
BEGIN
OBM:=GetO(GS,Lig) ; if OBM=Nil then Exit ;
XD:=ValD(GS,Lig) ; XC:=ValC(GS,Lig) ;
OBM.SetMontants(XD,XC,DEV,False) ;
if Calc then CalculDebitCredit ;
END ;

procedure TFeSaisie.AfficheConf ;
BEGIN
HConf.Visible:=(ModeConf>'0') or ((PieceConf) and (Action=taConsult)) ;
END ;

procedure TFeSaisie.CalculDebitCredit ;
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
SI_TotDD:=0 ; SI_TotCD:=0 ; SI_TotDE:=0 ; SI_TotCE:=0 ;
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
//       SI_TotDE:=SI_TotDE+OBM.GetMvt('E_DEBITEURO') ; SI_TotCE:=SI_TotCE+OBM.GetMvt('E_CREDITEURO') ;
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

procedure TFeSaisie.InfoLigne ( Lig : integer ; Var D,C,TD,TC : Double ) ;
Var Cpte : String17 ;
BEGIN
D:=0 ; C:=0 ; TD:=0 ; TC:=0 ;
Cpte:=GS.Cells[SA_Gen,Lig] ; if Cpte='' then Exit ;
D:=ValD(GS,Lig) ; C:=ValC(GS,Lig) ;
SommeSoldeCompte(SA_Gen,Cpte,TD,TC,FALSE,True) ;
END ;

procedure TFeSaisie.SommeSoldeCompte ( Col : integer ; Cpte : String ; Var TD,TC : Double ; Old,OkDev : Boolean) ;
Var i : integer ;
BEGIN
TD:=0 ; TC:=0 ; if Cpte='' then Exit ;
for i:=1 to GS.RowCount-2 do if GS.Cells[Col,i]=Cpte then
    BEGIN
    if OkDev then
       BEGIN
       TD:=TD+ValD(GS,i) ; TC:=TC+ValC(GS,i) ;
       END
       {
       else if ModeOppose then
       BEGIN
       if VH^.TenueEuro then
          BEGIN
          TD:=TD+PivotToEuro(Valeur(GS.Cells[SA_Debit,i]))-Ord(Old)*GetDouble(GS,i,'OLDDEBIT') ;
          TC:=TC+PivotToEuro(Valeur(GS.Cells[SA_Credit,i]))-Ord(Old)*GetDouble(GS,i,'OLDCREDIT');
          END else
          BEGIN
          TD:=TD+EuroToPivot(Valeur(GS.Cells[SA_Debit,i]))-Ord(Old)*GetDouble(GS,i,'OLDDEBIT') ;
          TC:=TC+EuroToPivot(Valeur(GS.Cells[SA_Credit,i]))-Ord(Old)*GetDouble(GS,i,'OLDCREDIT');
          END ;
       END
       }
       else BEGIN
       TD:=TD+ValeurPivot(GS.Cells[SA_Debit,i],DEV.Taux,DEV.Quotite)-Ord(Old)*GetDouble(GS,i,'OLDDEBIT') ;
       TC:=TC+ValeurPivot(GS.Cells[SA_Credit,i],DEV.Taux,DEV.Quotite)-Ord(Old)*GetDouble(GS,i,'OLDCREDIT');
       END ;
    END ;
END ;

procedure TFeSaisie.CalculSoldeCompte ( CpteG,CpteA : String ; DIG,CIG,DIA,CIA : Double ) ;
Var TDG,TCG,TDA,TCA : Double ;
BEGIN
if Action<>taConsult then SommeSoldeCompte(SA_Gen,CpteG,TDG,TCG,TRUE,FALSE) ;
AfficheLeSolde(SA_SoldeG,TDG+DIG,TCG+CIG) ;
if Action<>taConsult then SommeSoldeCompte(SA_Aux,CpteA,TDA,TCA,TRUE,FALSE) ;
AfficheLeSolde(SA_SoldeT,TDA+DIA,TCA+CIA) ;
END ;

procedure TFeSaisie.CaptureRatio ( Lig : integer ; Force : boolean ) ;
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

procedure TFeSaisie.InsereLigne(Lig : integer) ;
BEGIN
if PasModif then Exit ;
if Lig<GS.RowCount-1 then BEGIN GS.InsertRow(Lig) ; DefautLigne(Lig,True) ; END ;
NumeroteLignes(GS) ; NumeroteVentils ; AffichePied ;
END ;

{==========================================================================================}
{================================ Appels des comptes en saisie=== =========================}
{==========================================================================================}
procedure TFeSaisie.ChercheJAL ( Jal : String3 ; Zoom : boolean ) ;
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
if SAJAL.OkFerme then HPiece.Execute(30,caption,SAJAL.JOURNAL+'  '+SAJAL.LibelleJournal+')') ;
end;

procedure TFeSaisie.AppelAuto ( indice : integer ) ;
Var Cpte : String ;
BEGIN
if PasModif then Exit ; if GS.Col<>SA_Gen then Exit ;
Cpte:=TrouveAuto(SAJAL.COMPTEAUTOMAT,indice) ; if Cpte='' then Exit ;
GS.Cells[SA_Gen,GS.Row]:=Cpte ;
END ;

Function TFeSaisie.ChercheGen ( C,L : integer ; Force : boolean ) : byte ;
Var St   : String ;
    CGen,CGenAvant : TGGeneral ;
    CAux : TGTiers ;
    Idem,Changed,CollAvant : boolean ;
BEGIN
  Result:=0 ;
  Changed:=False ;
  CollAvant:=False ;
  if M.ANouveau
    then Cache.ZoomTable := tzGBilan
    else Cache.ZoomTable := tzGeneral ;
  St := uppercase(ConvertJoker(GS.Cells[C,L])) ;
  Cache.Text := St ;
  CGenAvant := GetGGeneral(GS,L) ;
  if CGenAvant<>Nil
    then CollAvant := CGenAvant.Collectif ;
  CAux := GetGTiers(GS,L) ;
  if ((CAux<>Nil) and (GS.Cells[SA_Aux,L]<>''))
    then QuelZoomTableG(Cache,CAux) ;
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
      if EstInterdit(SAJAL.COMPTEINTERDIT,Cache.Text,0)>0
        then BEGIN ErreurSaisie(4) ; Exit ; END ;
      if ((Not M.ANouveau) and (EstOuvreBil(Cache.Text)))
        then BEGIN ErreurSaisie(20) ; Exit ; END ;
      GS.Cells[C,L]   := Cache.Text ;
      Desalloue(GS,C,L) ;
      CGen            := TGGeneral.Create(Cache.Text) ;
      GS.Objects[C,L] := TObject(CGen) ;
      if Not Idem then
        BEGIN
        if ((Not CollAvant) or (Not CGen.Collectif)) then DesalloueEche(L) ;
        DesalloueAnal(L) ;
        END ;
      if ((M.ANouveau) and (EstChaPro(CGen.NatureGene))) then
        BEGIN
        DesalloueEche(L) ;
        DesalloueAnal(L) ;
        Desalloue(GS,C,L) ;
        GS.Cells[C,L]:='' ;
        ErreurSaisie(21) ; Exit ;
        END ;
      if ((CGen.NatureGene='BQE') and (E_NATUREPIECE.Value<>'ECC')) then if PasDeviseBanque(CGen.General,E_DEVISE.Value) then
        BEGIN
        DesalloueEche(L) ;
        DesalloueAnal(L) ;
        Desalloue(GS,C,L) ;
        GS.Cells[C,L]:='' ;
        ErreurSaisie(25) ;
        Exit ;
        END ;
      if CGen.Ferme then
        BEGIN
        if Arrondi(CGen.TotalDebit-CGen.TotalCredit,V_PGI.OkDecV)=0 then
          BEGIN
          DesalloueEche(L) ;
          DesalloueAnal(L) ;
          Desalloue(GS,C,L) ;
          GS.Cells[C,L]:='' ;
          ErreurSaisie(30) ;
          Exit ;
          END
        else
          BEGIN
          ErreurSaisie(17) ;
          END ;
        END ;
      if Not Idem then AlloueEcheAnal(C,L) ;
      Result:=1 ;
      AffecteConso(C,L) ;
      AffecteTva(C,L) ;
      If (CGen.NatureGene='TIC') Or (CGen.NatureGene='TID') Then
        If AffecteRIB(C,L,FALSE)
          Then BModifRIB.Enabled:=TRUE ;
      END
    else
      BEGIN
      CGen:=TGGeneral(GS.Objects[C,L]) ;
      Result:=2 ;
      END ;
    if ((Result>0) and (CGen<>Nil)) then
      if not CGen.Collectif then
        BEGIN
        Desalloue(GS,SA_Aux,L) ;
        GS.Cells[SA_Aux,L]:='' ;
        if Lettrable(CGen)=NonEche then DesalloueEche(L) ;
        if Not Ventilable(CGen,0) then DesalloueAnal(L) ;
        END ;
   AffichePied ;
   END ;
END ;

Procedure TFeSaisie.AffecteTva ( C,L : integer ) ;
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
         END ;
      END ;
   O.PutMvt('E_REGIMETVA',GeneRegTva) ;
   END else if isHT(GS,L,True) then
   BEGIN
   CGen:=GetGGeneral(GS,L) ; if CGen=Nil then Exit ;
   StE:=FlagEncais(ExigeEntete,CGen.TvaSurEncaissement) ;
   O.PutMvt('E_TVAENCAISSEMENT',StE) ; O.PutMvt('E_REGIMETVA',GeneRegTva) ;
   O.PutMvt('E_TVA',CGen.Tva)        ; O.PutMvt('E_TPF',CGen.Tpf) ;
   END ;
END ;

Procedure TFeSaisie.AffecteConso ( C,L : integer ) ;
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

Procedure TFeSaisie.AffecteLeRib(O : TOBM ; Cpt : String ; ForceRAZ : Boolean = FALSE) ;
Var St,SRib : String ;
BEGIN
If O=Nil Then Exit ; If Cpt='' Then Exit ;
sRIB:=O.GetMvt('E_RIB') ;
If (SRib<>'') And ForceRAZ Then O.PutMvt('E_RIB',St) ;
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

Function  TFeSaisie.AffecteRIB ( C,L : integer ; IsAux : Boolean ; ForceRAZ : Boolean = FALSE) : Boolean ;
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

Function TFeSaisie.ChercheAux ( C,L : integer ; Force : boolean ) : byte ;
Var St   : String ;
    CAux : TGTiers ;
    CGen,GenColl : TGGeneral ;
    Idem,Changed : boolean ;
    Err  : integer ;
BEGIN
  Result   := 0 ;
  Changed  := False ;
  Cache.ZoomTable:=tzTiers ;
  St:=uppercase(ConvertJoker(GS.Cells[C,L])) ;
  Cache.Text:=St ;
  CGen:=GetGGeneral(GS,L) ;
  if ((CGen<>Nil) and (GS.Cells[SA_Gen,L]<>''))
    then QuelZoomTableT(Cache,CGen) ;

  // ====> DEBUT IF
{$IFNDEF GCGC}
 {$IFDEF EAGLCLIENT}
  if GChercheCompte(Cache,Nil) then
  {$ELSE}
	if GChercheCompte(Cache,FicheTiers) then
  {$ENDIF}
{$ELSE}
  if GChercheCompte(Cache,Nil) then
{$ENDIF}
    BEGIN
    if GS.Objects[C,L]<>Nil then
      if TGTiers(GS.Objects[C,L]).Auxi<>St then Changed:=True ;
    Idem:=((St=Cache.Text) and (GS.Objects[C,L]<>Nil) and (Not Changed)) ;
    if ((Not Idem) or (Force)) then
      BEGIN
      GS.Cells[C,L]:=Cache.Text ;
      Desalloue(GS,C,L) ;
      if Not Idem then
        BEGIN
        DesalloueEche(L) ;
        if ((CGen=Nil) or (GS.Cells[SA_Gen,L]='')) then DesalloueAnal(L) else
          if CGen<>Nil then
          BEGIN
          if Not Ventilable(CGen,0) then DesalloueAnal(L) ;
          END ;
        END ;
      CAux:=TGTiers.Create(Cache.Text) ;
      GS.Objects[C,L]:=TObject(CAux) ;
      // Refuser tiers en devise si pas celle de saisie et pas sur devisepivot ni opposée
      if ((Not CAux.MultiDevise) and (DEV.Code<>CAux.Devise) and (DEV.Code<>V_PGI.DevisePivot) and
          (CAux.Devise<>'') and (E_NATUREPIECE.Value<>'ECC')) then
        BEGIN
        DesalloueEche(L) ;
        Desalloue(GS,C,L) ;
        GS.Cells[C,L]:='' ;
        ErreurSaisie(28) ;
        Exit ;
        END ;
      if CAux.Ferme then
        BEGIN
        if Arrondi(CAux.TotalDebit-CAux.TotalCredit,V_PGI.OkDecV)=0 then
          BEGIN
          DesalloueEche(L) ;
          Desalloue(GS,C,L) ;
          GS.Cells[C,L]:='' ;
          ErreurSaisie(30) ;
          Exit ;
          END
        else
          BEGIN
          ErreurSaisie(27) ;
          END ;
        END ;
      if Not Idem then AlloueEcheAnal(C,L) ;
      Result:=1 ;
      AffecteRIB(C,L,TRUE,(Not Idem)) ;
      AffecteConso(C,L) ;
      AffecteTva(C,L) ;
      END
    else
      BEGIN
      Result:=2 ;
      END ;
    if Result>0 then
      if GS.Cells[SA_Gen,L]='' then
        BEGIN
        GS.Cells[SA_Gen,L]:=GetGTiers(GS,L).Collectif ;
        Desalloue(GS,SA_Gen,L) ;
        Err:=0 ;
        GenColl:=TGGeneral.Create(GS.Cells[SA_Gen,L]) ;
        GS.Objects[SA_Gen,L]:=GenColl ;
        if EstInterdit(SAJAL.COMPTEINTERDIT,GS.Cells[SA_Gen,L],0)>0 then Err:=4 ;
        if EstConfidentiel(GenColl.Confidentiel) then Err:=24 ;
        if GenColl.Ferme then
          BEGIN
          if Arrondi(GenColl.TotalDebit-GenColl.TotalCredit,V_PGI.OkDecV)=0
             then Err:=30 else ErreurSaisie(17) ;
          END ;
        if Err>0 then
          BEGIN
          DesalloueEche(L) ;
          Desalloue(GS,C,L) ;
          Desalloue(GS,SA_Gen,L) ;
          GS.Cells[SA_Gen,L]:='' ;
          GS.Cells[SA_Aux,L]:='' ;
          ErreurSaisie(Err) ;
          END
        else
          BEGIN
          AlloueEcheAnal(SA_Gen,L) ;
          END ;
        END ;
    AffichePied ;
    END ;
    // ====> FIN IF
END ;

procedure TFeSaisie.ChercheMontant(Acol,Arow : longint) ;
BEGIN
CaptureRatio(ARow,False) ;
if ACol=SA_Debit then BEGIN if ValD(GS,ARow)<>0 then GS.Cells[SA_Credit,ARow]:='' else GS.Cells[ACol,ARow]:='' ; END
                 else BEGIN if ValC(GS,ARow)<>0 then GS.Cells[SA_Debit,ARow]:='' else GS.Cells[ACol,ARow]:='' ; END ;
FormatMontant(GS,ACol,ARow,DECDEV) ; TraiteMontant(ARow,True) ;
END ;

{==========================================================================================}
{================================= Conversions, Caluls ====================================}
{==========================================================================================}
Function  TFeSaisie.DateMvt ( Lig : integer ) : TDateTime ;
BEGIN
Result:=StrToDate(E_Datecomptable.Text) ;
END ;

Function  TFeSaisie.ArrS( X : Double ) : Double ;
BEGIN
Result:=Arrondi(X,DECDEV) ;
END ;

{==========================================================================================}
{========================= Affichages, Positionnements ====================================}
{==========================================================================================}
Procedure TFeSaisie.AutoSuivant( Suiv : boolean ) ;
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

Procedure TFeSaisie.SoldelaLigne ( Lig : integer ) ;
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
{
// Euro
Diff:=SI_TotDE-SI_TotCE ;
if Diff<>0 then
   BEGIN
   X:=-1*OBM.GetMvt('E_DEBITEURO') ; if X=0 then X:=OBM.GetMvt('E_CREDITEURO') ; Diff:=Diff+X ;
   if Diff>0 then OBM.PutMvt('E_CREDITEURO',Diff) else OBM.PutMvt('E_DEBITEURO',-Diff) ;
   Okok:=True ;
   END ;
}
if Okok then CalculDebitCredit ;
AjusteLigne(Lig,False) ;
END ;

Procedure TFeSaisie.PositionneDevise ( ReInit : boolean ) ;
Var OldIndex : integer ;
BEGIN
if ((Not E_DEVISE.Enabled) and (E_DEVISE.Value<>'') and (GS.RowCount>2)) then Exit ;
OldIndex:=-2 ;
//if ((ReInit) and (ModeOppose)) then OldIndex:=E_DEVISE.ItemIndex ;
if SAJAL.MultiDevise then
   BEGIN
   E_DEVISE_.Enabled:=True ;
   {RestoreDeviseEuro ;} E_DEVISE.Enabled:=True ;
   END else
   BEGIN
   E_DEVISE_.Value:=V_PGI.DevisePivot ; E_DEVISE_.Enabled:=False ;
//   AddEuroFranc ;
   END ;
if ((ReInit) and (OldIndex<>-2)) then BEGIN E_DEVISE.ItemIndex:=OldIndex ; E_DEVISEChange(Nil) ; END ; 
END ;

procedure TFeSaisie.AffichePied ;
Var CGen : TGGeneral ;
    CAux : TGTiers ;
    CpteG,CpteA : String ;
    DIG,CIG,DIA,CIA : double ;
BEGIN
G_LIBELLE.Font.Color:=clBlack ; G_LIBELLE.Caption:='' ; T_LIBELLE.Caption:='' ; CpteG:='' ; CpteA:='' ;
DIG:=0 ; CIG:=0 ; DIA:=0 ; CIA:=0 ;
if GS.Objects[SA_Gen,GS.Row]<>Nil then
   BEGIN
   CGen:=TGGeneral(GS.Objects[SA_Gen,GS.Row]) ;
   G_LIBELLE.Caption:=CGen.Libelle ; CpteG:=CGen.GENERAL ;
   DIG:=CGen.TotalDebit ; CIG:=CGen.TotalCredit ;
   LSA_SoldeG.Visible:=True ;
   END else LSA_SoldeG.Visible:=False ;
if GS.Objects[SA_Aux,GS.Row]<>Nil then
   BEGIN
   CAux:=TGTiers(GS.Objects[SA_Aux,GS.Row]) ;
   T_LIBELLE.Caption:=CAux.Libelle ; CpteA:=CAux.AUXI ;
   DIA:=CAux.TotalDebit ; CIA:=CAux.TotalCredit ;
   LSA_SoldeT.Visible:=True ;
   END else LSA_SoldeT.Visible:=False ;
CalculSoldeCompte(CpteG,CpteA,DIG,CIG,DIA,CIA) ;
END ;

procedure TFeSaisie.GereZoom ;
Var Okok : boolean ;
BEGIN
Okok:=False ;
if ((GS.Col=SA_Gen) and (GS.Cells[SA_Gen,GS.Row]<>'')) then Okok:=True ;
if ((GS.Col=SA_Aux) and (GS.Cells[SA_Aux,GS.Row]<>'')) then Okok:=True ;
BZoom.Enabled:=Okok ;
END ;

(*
{==========================================================================================}
{======================================== EURO ============================================}
{==========================================================================================}
procedure TFeSaisie.SetOpposeEuro ;
BEGIN
ModeOppose:=False ;
if ((E_DEVISE.Values.Count>=2) and (E_DEVISE.Values[E_DEVISE.Values.Count-1]=V_PGI.DevisePivot) and
    (E_DEVISE.ItemIndex=E_DEVISE.Values.Count-1)) then ModeOppose:=True ;
ColorOpposeEuro(GS,ModeOppose,ModeDevise) ;
END ;

procedure TFeSaisie.RestoreDeviseEuro ;
Var ii : integer ;
BEGIN
if E_DEVISE.DataType<>'' then Exit ;
ii:=E_DEVISE.ItemIndex ;
E_DEVISE.DataType:='TTDEVISE' ; E_DEVISE.Reload ;
AddEuroCombo ;
E_DEVISE.Value:=V_PGI.DevisePivot ;
if ii>0 then E_DEVISE.ItemIndex:=E_DEVISE.Values.Count-1 ;
END ;

procedure TFeSaisie.AddEuroFranc ;
BEGIN
E_DEVISE.DataType:='' ; E_DEVISE.Items.Clear ; E_DEVISE.Values.Clear ;
E_DEVISE.Values.Add(V_PGI.DevisePivot) ; E_DEVISE.Items.Add(RechDom('TTDEVISE',V_PGI.DevisePivot,False)) ;
AddEuroCombo ;
E_DEVISE.Value:=E_DEVISE.Values[0] ;
E_DEVISE.Enabled:=True ;
END ;

procedure TFeSaisie.AddEuroCombo ;
Var i,Nb : integer ;
BEGIN
Nb:=0 ;
for i:=0 to E_DEVISE.Values.Count-1 do if E_DEVISE.Values[i]=V_PGI.DevisePivot then Inc(Nb) ;
if Nb>=2 then Exit ;
if VH^.TenueEuro then E_DEVISE.Items.Add(RechDom('TTDEVISETOUTES',V_PGI.DeviseFongible,False))
                 else E_DEVISE.Items.Add(HDiv.Mess[13]) ;
E_DEVISE.Values.Add(V_PGI.DevisePivot) ;
END ;
*)
{==========================================================================================}
{================================ Methodes de la form =====================================}
{==========================================================================================}
procedure TFeSaisie.FormKeyPress(Sender: TObject; var Key: Char);
begin
if Not GS.SynEnabled then Key:=#0 else
   BEGIN
   if Key=#127 then Key:=#0 else
    if ((Key='=') and ((GS.Col=SA_Debit) or (GS.Col=SA_Credit) or (GS.Col=SA_Gen) or (GS.Col=SA_Aux))) then Key:=#0 else
     if Key=#17 then Key:=#0 ;
   END ;
OkSuite:=False ;
end;

Procedure TFeSaisie.FocusSurNatP ;
BEGIN
If E_JOURNAL.Focused And VH^.JalLookUp Then
  BEGIN
  If E_NATUREPIECE.CanFocus Then E_NATUREPIECE.SETFOCUS Else If E_DateComptable.CanFocus Then E_DateComptable.SetFocus ;
  END ;
(*NextControl(Self,TRUE) ;*)
END ;

procedure TFeSaisie.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var OkG,Vide : boolean ;
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
Case Key of
  VK_F3     : BEGIN
              if ((OkG) and (Vide) and (GS.Col=SA_Gen)) then BEGIN FocusSurNatP ; Key:=0 ; AutoSuivant(False) ; END else
               if Shift=[ssShift] then BEGIN FocusSurNatP ; Key:=0 ; BPrevClick(Nil) ; END ;
              END ;
  VK_F4     : BEGIN
              if ((OkG) and (Vide) and (GS.Col=SA_Gen)) then BEGIN Key:=0 ; AutoSuivant(True) ; END else
               if Shift=[ssShift] then BEGIN Key:=0 ; BNextClick(Nil) ; END ;
              END ;
  VK_F5     : if ((OkG) and (Vide)) then BEGIN Key:=0 ; ClickZoom ; END else
                if ((OkG) and (Shift=[ssAlt])) then BEGIN Key:=0 ; ClickConsult ; END ;
  VK_F6     : if ((OkG) and (Vide)) then BEGIN Key:=0 ; ClickSolde(False,True) ; END ;
  VK_F8     : if Vide then BEGIN Key:=0 ; ClickControleTva(False,0) ; END else
               if Shift=[ssAlt] then BEGIN Key:=0 ; ClickSwapPivot ; END else
                if Shift=[ssShift] then BEGIN Key:=0 ; {ClickSwapEuro ;} END ;
  VK_F9     : if Vide then BEGIN Key:=0 ; ClickGenereTva ; END ;
  VK_F10    : if Vide then BEGIN Key:=0 ; FocusSurNatP ; ClickValide ; END ;
  VK_ESCAPE : if Vide then ClickAbandon ;
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
{AT}     84 : if Shift=[ssAlt]  then BEGIN Key:=0 ; ClickEtabl ; END ;
{AV}     86 : if Shift=[ssAlt]  then BEGIN Key:=0 ; ClickModifTva ; END ;
{^Y}     89 : if Shift=[ssCtrl] then BEGIN Key:=0 ; ClickVidePiece(True) ; END ;
{=}     187 : if ((OkG) and (Vide)) then BEGIN Key:=0 ; ClickSolde(True,True) ; END ;
  END ;
end;

procedure TFeSaisie.CreateSaisie ;
BEGIN
GS.VidePile(True) ; GS.RowCount:=2 ; GS.TypeSais:=tsGene ; GS.ListeParam:='SSAISIE1' ;
SAJAL:=Nil ; DecDev:=V_PGI.OkdecV ; PieceModifiee:=False ; Revision:=False ; UnChange:=False ;
TS:=TList.Create ; TGUI:=TList.Create ; TPIECE:=TList.Create ; TECHEORIG:=TList.Create ;
TDELECR:=TList.Create ; TDELANA:=TList.Create ;
Scenario:=TOBM.Create(EcrSais,'',True) ; Entete:=TOBM.Create(EcrSais,'',True) ;
MemoComp:=TStringList.Create ; GeneRegTva:='' ; StCellCur:='' ;
ModifSerie:=TOBM.Create(EcrGen,'',True) ; ModifNext:=False ; GeneCharge:=False ;
DejaRentre:=False ; OkScenario:=False ; ModeConf:='0' ; PieceConf:=False ;
Volatile:=False ; {ModeOppose:=False ;} CEstGood:=False ; ModeLC:=False ;
ValideEcriture:=False ; GestionRIB:='...' ; NeedEdition:=False ; EtatSaisie:='' ;
YATP:=yaRien ; ListeTP:=TStringList.Create ; RentreDate:=False ; RegimeForce:=False ;
{#TVAENC}
SorteTva:=stvDivers ; OuiTvaLoc:=False ;
ExigeTva:=tvaMixte ; ExigeEntete:=tvaMixte ; ChoixExige:=False ;
BPopTva.ItemIndex:=Ord(ExigeEntete)+1 ;
ToutDebit:=True ; ToutEncais:=True ;
{Param Liste}
GS.ColLengths[SA_Gen]:=VH^.Cpta[fbGene].Lg ;
if VH^.CPCodeAuxiOnly then GS.ColLengths[SA_Aux]:=VH^.Cpta[fbAux].Lg else GS.ColLengths[SA_Aux]:=35 ;
GS.ColLengths[SA_RefI]:=35 ; GS.ColLengths[SA_Lib]:=35 ;
END ;

procedure TFeSaisie.PosLesSoldes ;
BEGIN
LSA_SoldeG.SetBounds(SA_SoldeG.Left,SA_SoldeG.Top,SA_SoldeG.Width,SA_SoldeG.Height) ;
LSA_SoldeT.SetBounds(SA_SoldeT.Left,SA_SoldeT.Top,SA_SoldeT.Width,SA_SoldeT.Height) ;
LSA_Solde.SetBounds(SA_Solde.Left,SA_Solde.Top,SA_Solde.Width,SA_Solde.Height) ;
LSA_TotalDebit.SetBounds(SA_TotalDebit.Left,SA_TotalDebit.Top,SA_TotalDebit.Width,SA_TotalDebit.Height) ;
LSA_TotalCredit.SetBounds(SA_TotalCredit.Left,SA_TotalCredit.Top,SA_TotalCredit.Width,SA_TotalCredit.Height) ;
SA_SoldeG.Visible:=False ; SA_SoldeT.Visible:=False ; SA_Solde.Visible:=False ;
SA_TotalDebit.Visible:=False ; SA_TotalCredit.Visible:=False ;
END ;

procedure TFeSaisie.FormCreate(Sender: TObject);
//Var PP : thPanel ;
begin
CreateSaisie ;
//AddEuroCombo ;
WMinX:=Width ; WMinY:=Height_Ecr ;
RegLoadToolbarPos(Self,'Saisie') ;
PosLesSoldes ;
{$IFDEF CCMP}
initSA_SaisieNormale ;
{$ENDIF}
end;

procedure TFeSaisie.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  CanClose := FermerSaisie ;
end;

procedure TFeSaisie.CloseSaisie ;
BEGIN
GS.VidePile(True) ; SAJAL.Free ; SAJAL:=Nil ;
VideListe(TS) ; TS.Free ; TS:=Nil ;
VideListe(TECHEORIG) ; TECHEORIG.Free ; TECHEORIG:=Nil ;
VideListe(TGUI) ; TGUI.Free ; TGUI:=Nil ;
VideListe(TPIECE) ; TPIECE.Free ; TPIECE:=Nil ;
VideListe(TDELECR) ; TDELECR.Free ; TDELECR:=Nil ;
VideListe(TDELANA) ; TDELANA.Free ; TDELANA:=Nil ;
MemoComp.Free ; MemoComp:=Nil ;
Scenario.Free ; Scenario:=Nil ;
Entete.Free ; Entete:=Nil ;
ListeTP.Clear ; ListeTP.Free ; ListeTP:=Nil ;
ModifSerie.Free ; ModifSerie:=Nil ;
PurgePopup(POPS) ; PurgePopup(POPZ) ;
{#LUM}
//TOBLumiere.Free ; TOBLumiere:=Nil ;
END ;

procedure TFeSaisie.CloseFen ;
var lBoDansPanel : Boolean ;
begin
  lBoDansPanel := IsInside(Self) ;
  Close ;
  if lBoDansPanel
    then CloseInsidePanel(Self) ;
end ;

procedure TFeSaisie.FormClose(Sender: TObject; var Action: TCloseAction);
begin
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
RestoreSAEff ;
{$ENDIF}
end;

procedure TFeSaisie.GereANouveau ;
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

procedure TFeSaisie.LectureSeule ;
BEGIN
BInsert.Enabled:=False ; BSDel.Enabled:=False ; BGuide.Enabled:=False ;
BGenereTVA.Enabled:=False ; BProrata.Enabled:=False ; BLibAuto.Enabled:=False ;
BSolde.Enabled:=False ; BDernPieces.Enabled:=False ; BCreerGuide.Enabled:=False ;
BMenuGuide.Enabled:=False ; BModifSerie.Enabled:=False ;
END ;

procedure TFeSaisie.RenseignePopTva ;
BEGIN
if Action=taCreat then Exit ;
if Not BPopTva.Visible then Exit ;
if ToutDebit then ExigeTva:=tvaDebit else
 if ToutEncais then ExigeTva:=tvaEncais else
    ExigeTva:=tvaMixte ;
ExigeEntete:=ExigeTva ;
BPopTva.ItemIndex:=Ord(ExigeEntete)+1 ;
END ;

procedure TFeSaisie.TraiteOrigMP ;
Var TOBL : TOB ;
    i    : integer ;
    St   : String ;
    T    : TStrings ;
    YY   : RMVT ;
BEGIN
if Not SuiviMP then Exit ;
T:=TStringList.Create ;
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
END ;

Function YAFormule(St : string) : Boolean ;
BEGIN
{$IFDEF CCMP}
Result:=((Pos('[',St)>0) And (Pos(']',St)>0)) Or (St<>'') ;
{$ELSE}
Result:=((Pos('[',St)>0) And (Pos(']',St)>0)) ;
{$ENDIF}
END ;

procedure TFeSaisie.TraiteFormuleMP(St : String ; O : TOBM ; Lig,Col : Integer ; FormuleEnDur : String = '') ;
Var St1,St2 : String ;
BEGIN
If FormuleEnDur='' Then St2:=O.GetMvt(St) Else St2:=FormuleEnDur ;
If St2='' Then Exit ;
If YAFormule(St2) Then
  BEGIN
  CurLig:=Lig ;
  St1:=GFormule(St2,Load_Sais,Nil,1) ;
  O.PutMvt(St,St1) ;
  If Col>0 Then GS.Cells[Col,Lig]:=St1 ;
  END ;
END ;

procedure TFeSaisie.AffecteRef(FormuleOrigine : Boolean) ;
Var O : TOBM ;
    i : Integer ;
{$IFDEF CCMP}
    CptG : String ;
{$ENDIF}
  Label 0 ;
BEGIN
{$IFDEF CCMP}
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
{$ENDIF}
0:for i:=1 to GS.RowCount-2 do
    BEGIN
    O:=GetO(GS,i) ; if O=Nil then Continue ;
    TraiteFormuleMP('E_REFINTERNE',O,i,SA_RefI) ; TraiteFormuleMP('E_LIBELLE',O,i,SA_Lib) ;
    TraiteFormuleMP('E_REFEXTERNE',O,i,0) ; TraiteFormuleMP('E_REFLIBRE',O,i,0) ;
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
procedure TFeSaisie.ShowSaisie ;
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
      			if M.Simul<>'N'
            	then E_JOURNAL.DataType:='ttJalSansEcart'
              else E_JOURNAL.DataType:='ttJalSaisie' ;
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
OkBloquePieceImport:=FALSE ;
If OkSynchro Then
  BEGIN
  OkBloquePieceImport:=M.BloquePieceImport ;
  If OkBloquePieceImport Then Action:=taConsult ;
  END ;
OkPasJal:=((VH^.JalAutorises<>'') and (Pos(';'+M.Jal+';',VH^.JalAutorises)<=0) and (Action=taModif)) ;
OkClo:=((Action=taModif) and (ControleDate(DateToStr(M.DateC)) in [2,3])) ;
if ((OkM) or (OkClo) or (OkPasJal) or (OkPasModif)) then Action:=taConsult ;
if ((Action=taCreat) and (M.Etabl='')) then M.Etabl:=VH^.EtablisDefaut ;
DefautEntete ; DefautPied ; NbLigEcr:=0 ; NbLigAna:=0 ; GeneTreso:='' ;
ChargeLignes ;
ChargeAnals ;
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
    If OkSynchro And OkBloquePieceImport then HPiece.Execute(52,caption,'') ;
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

procedure TFeSaisie.TraiteAffichePCL ;
BEGIN
if Not (ctxPCL in V_PGI.PGIContexte) then Exit ;
if VH^.CPPCLSansAna then BVentil.Visible:=False ;
END ;

procedure TFeSaisie.TraiteAfficheGescom ;
BEGIN
if Not (ctxGescom in V_PGI.PGIContexte) then Exit ;
{$IFDEF GCGC}
BGenereTVA.Visible:=False   ; BGenereTVA.Enabled:=False ;
BControleTVA.Visible:=False ; BControleTVA.Enabled:=False ;
BModifTVA.Visible:=False    ; BModifTVA.Enabled:=False ;
BComplement.Visible:=False  ; BComplement.Enabled:=False ;
BZoomJournal.Visible:=False ; BZoomJournal.Enabled:=False ;
BPieceGC.Visible:=False     ; BPieceGC.Enabled:=False ;
BScenario.Visible:=False    ; BScenario.Enabled:=False ;
BDernPieces.Visible:=False  ; BDernPieces.Enabled:=False ;
BMenuGuide.Visible:=False   ; BMenuGuide.Enabled:=False ;
BModifS.Visible:=False      ; BModifS.Enabled:=False ;
BZoomImmo.Visible:=False    ; BZoomImmo.Enabled:=False ;
BZoomTP.Visible:=False      ; BZoomTP.Enabled:=False ;
BSaisTaux.Visible:=False    ; BSaisTaux.Enabled:=False ;
//BSwapEuro.Visible:=False    ; BSwapEuro.Enabled:=False ;
BSwapPivot.Visible:=False   ; BSwapPivot.Enabled:=False ;
{$ENDIF}
END ;

procedure TFeSaisie.FormShow(Sender: TObject);
begin
LookLesDocks(Self) ;
TraiteAffichePCL ;
TraiteAfficheGescom ;
ShowSaisie ;
end;

{==========================================================================================}
{================================ Ecriture Fichier=========================================}
{==========================================================================================}
{$IFDEF EAGLCLIENT}
{$ELSE}
procedure TFeSaisie.IntoucheGene ;
Var Cpte  : String ;
    QGene : TQuery ;
    Lig   : integer ;
    CGen  : TGGeneral ;
BEGIN
if PasModif then Exit ; if M.Simul='N' then Exit ;
QGene:=FabricReqLire(fbGene) ;
for Lig:=1 to GS.RowCount-2 do if ((GS.Cells[SA_Gen,Lig]<>'') and (GS.Objects[SA_Gen,Lig]<>Nil)) then
    BEGIN
    CGen:=TGGeneral(GS.Objects[SA_Gen,Lig]) ; Cpte:=GS.Cells[SA_Gen,Lig] ;
    if ((CGen.TotalDebit<>0) or (CGen.TotalCredit<>0)) then Continue ;
    QGene.ParamByName('CPTE').AsString:=Cpte ;
    AttribParamsComp(QGene,CGen) ;
    QGene.Open ;
    if QGene.EOF then BEGIN V_PGI.IoError:=oeSaisie ; Break ; END ;
    QGene.Close ;
    END ;
QGene.Free ;
END ;

procedure TFeSaisie.IntoucheAux ;
Var Cpte  : String ;
    QAux  : TQuery ;
    Lig   : integer ;
    CAux  : TGTiers ;
BEGIN
if PasModif then Exit ; if M.Simul='N' then Exit ;
QAux:=FabricReqLire(fbAux) ;
for Lig:=1 to GS.RowCount-2 do if ((GS.Cells[SA_Aux,Lig]<>'') and (GS.Objects[SA_Aux,Lig]<>Nil)) then
    BEGIN
    CAux:=TGTiers(GS.Objects[SA_Aux,Lig]) ; Cpte:=GS.Cells[SA_Aux,Lig] ;
    if ((CAux.TotalDebit<>0) or (CAux.TotalCredit<>0)) then Continue ;
    QAux.ParamByName('CPTE').AsString:=Cpte ;
    AttribParamsComp(QAux,CAux) ;
    QAux.Open ;
    if QAux.EOF then BEGIN V_PGI.IoError:=oeSaisie ; Break ; END ;
    QAux.Close ;
    END ;
QAux.Free ;
END ;
{$ENDIF}

procedure TFeSaisie.MajCptesGeneNew ;
Var //Cpte  : String ;
    XD,XC : Double ;
    Lig   : integer ;
    FRM : TFRM ;
    ll : LongInt ;
BEGIN
if PasModif then Exit ;
if M.Simul<>'N' then Exit ;
for Lig:=1 to GS.RowCount-2 do if ((GS.Cells[SA_Gen,Lig]<>'') and (GS.Objects[SA_Gen,Lig]<>Nil)) then
    BEGIN
    Fillchar(FRM,SizeOf(FRM),#0) ;
    FRM.Cpt:=GS.Cells[SA_Gen,Lig] ; XD:=RecupDebit(GS,Lig,tsmPivot) ; XC:=RecupCredit(GS,Lig,tsmPivot) ;
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
    if Action<>taConsult then AttribParamsCompNew(FRM,GS.Objects[SA_Gen,Lig]) ;
    LL:=ExecReqMAJNew(fbGene,M.ANouveau,Action<>taConsult,FRM) ;
    If ll<>1 then V_PGI.IoError:=oeSaisie ;
    END ;
END ;

procedure TFeSaisie.MajCptesAuxNew ;
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
    Fillchar(FRM,SizeOf(FRM),#0) ;
    FRM.Cpt:=GS.Cells[SA_Aux,Lig] ; XD:=RecupDebit(GS,Lig,tsmPivot) ; XC:=RecupCredit(GS,Lig,tsmPivot) ;
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
    if Action<>taConsult then AttribParamsCompNew(FRM,GS.Objects[SA_Aux,Lig]) ;
    ll:=ExecReqMAJNew(fbAux,M.ANouveau,Action<>taConsult,FRM) ;
    If ll<>1 then V_PGI.IoError:=oeSaisie ;
    END ;
END ;

procedure TFeSaisie.MajCptesSectionNew ;
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
END ;

function TFeSaisie.QuelEnc ( Lig : Integer ) : String3 ;
BEGIN
Result:=SensEnc(ValD(GS,Lig),ValC(GS,Lig)) ;
END ;

Procedure TFeSaisie.GetCptsContreP ( Lig : Integer ; Var LeGene,LeAux,TL0 : String17 ) ;
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

Procedure TFeSaisie.RecupTronc ( Lig : Integer ; MO : TMOD ; Var OBM : TOBM);
Var CGen : TGGeneral ;
    CAux : TGTiers ;
    NatG,NatT : String3 ;
    OBA  : TOB ;
    TVA,TPF  : String3 ;
    Coll     : String ;
    LeGen,LeAux,TL0 : String17 ;
    i,NumAxe,NumV{,iEche} : integer ;
BEGIN
  if PasModif then exit ;
  CGen:=NIL ; CAux:=NIL ; NatG:='' ; NatT:='' ; OBM:=Nil ;
  if GS.Objects[SA_Gen,Lig]<>Nil then BEGIN CGen:=TGGeneral(GS.Objects[SA_Gen,Lig]) ; NatG:=CGen.NatureGene ; END ;
  if GS.Objects[SA_Aux,Lig]<>Nil then BEGIN CAux:=TGTiers(GS.Objects[SA_Aux,Lig]) ; NatT:=CAux.NatureAux ; END ;
  if GS.Objects[SA_Exo,Lig]<>Nil then OBM:=GetO(GS,Lig) ;
  if OBM = nil then Exit;

  {Teste CFONB}
  if Not SuiviMP then
    BEGIN
    if ((M.TypeGuide='ENC') or (M.TypeGuide='DEC')) and ((M.ExportCFONB) or (M.Bordereau)) and (OBM.LC.Count>0)
    	then OBM.PutMvt('E_CFONBOK','#') ;
    END
  else
    BEGIN
    if ((M.ExportCFONB) or (M.Bordereau)) and (OBM.GetMvt('E_TRACE')<>'')
    	then OBM.PutMvt('E_CFONBOK','#') ;
    END ;

  OBM.SetCotation(DatePiece) ;
  if (not SuiviMP) or (M.MPGUnique <> '') then OBM.SetMPACC ; // FQ 12438

  {$IFDEF EAGLCLIENT}
  // ????
  {$ELSE}
//  	TMemoField(TEcrGen.FindField('E_BLOCNOTE')).Assign(OBM.M) ; // EAGL : Memo a gérer
  {$ENDIF}

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
  if NatT<>'' then OBM.PutMvt('E_TYPEMVT',QuelTypeMvt(CAux.Auxi,NatT,E_NATUREPIECE.Value))
              else OBM.PutMvt('E_TYPEMVT',QuelTypeMvt(CGen.General,NatG,E_NATUREPIECE.Value)) ;
  if OBM.GetMvt('E_REGIMETVA')='' then OBM.PutMvt('E_REGIMETVA',GeneRegTVA) ;
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
     if OBM.GetMvt('E_TVA')='' then OBM.PutMvt('E_TVA',GeneTva) ;
     if OBM.GetMvt('E_TPF')='' then OBM.PutMvt('E_TPF',GeneTpf) ;
     OBM.PutMvt('E_BUDGET',CGen.Budgene) ;
     if ((Lettrable(CGen)=MonoEche) and (MO<>Nil)) then
        BEGIN
        OBM.PutMvt('E_NUMECHE',1) ;
        OBM.PutMvt('E_MODEPAIE',MO.MODR.TabEche[1].ModePaie) ;
        OBM.PutMvt('E_DATEECHEANCE',MO.MODR.TabEche[1].DateEche) ;
        OBM.PutMvt('E_DATEVALEUR',MO.MODR.TabEche[1].DateValeur) ;
        OBM.PutMvt('E_ECHE','X') ;
        OBM.PutMvt('E_ENCAISSEMENT',QuelEnc(Lig)) ;
        OBM.PutMvt('E_ORIGINEPAIEMENT',MO.MODR.TabEche[1].DateEche) ;
        {#TVAENC}
        if OuiTvaLoc then
           BEGIN
           Coll:=CGen.General ;
           if EstCollFact(Coll) then
              BEGIN
              if ((SorteTva<>stvDivers) and (UnSeulTiers)) then
                 BEGIN
                {Cycle facture}
                 for i:=1 to 4 do OBM.PutMvt('E_ECHEENC'+IntToStr(i),TabTvaEnc[i]) ;
                 OBM.PutMvt('E_ECHEDEBIT',TabTvaEnc[5]) ;
                 OBM.PutMvt('E_EMETTEURTVA','X') ;
                 END
              else if ((SorteTva=stvDivers) and (MO.MODR.ModifTva)) then
                     BEGIN
                     {Cycle bases HT sur échéances}
                     for i:=1 to 4 do OBM.PutMvt('E_ECHEENC'+IntToStr(i),MO.MODR.TabEche[1].TAV[i]) ;
                     OBM.PutMvt('E_ECHEDEBIT',MO.MODR.TabEche[1].TAV[5]) ;
                     END ;
              END ;
           END ;
        END ;
     END ;
  GetCptsContreP(Lig,LeGen,LeAux,TL0) ;
  If (TL0<>'') And EstSpecif('51196') Then OBM.PutMvt('E_TABLE0',TL0) ;
  OBM.PutMvt('E_TRACE','') ;
  OBM.PutMvt('E_CONTREPARTIEGEN',LeGen) ;
  OBM.PutMvt('E_CONTREPARTIEAUX',LeAux) ;
  OBM.PutMvt('E_DATEMODIF',NowFutur) ;
  OBM.PutMvt('E_CONFIDENTIEL',ModeConf) ;
  {Mise à jour contreparties analytiques}
  if OBM.Detail.Count>0 then
  	for NumAxe:=0 to OBM.Detail.Count-1 do
      for NumV:=0 to OBM.Detail[NumAxe].Detail.Count-1 do
        BEGIN
        OBA := OBM.Detail[NumAxe].Detail[NumV];
        OBA.PutValue('Y_CONTREPARTIEGEN',OBM.GetMvt('E_CONTREPARTIEGEN')) ;
        OBA.PutValue('Y_CONTREPARTIEAUX',OBM.GetMvt('E_CONTREPARTIEAUX')) ;
        END ;
END ;

Procedure TFeSaisie.RecupAnal(Lig : integer ; var OBA : TOB ; NumAxe,NumV : integer ) ;
//Var i : integer ;
BEGIN
	if PasModif then Exit ;
	// Attention NumV démarre à 0
	if NumV=0
  	then OBA.PutValue('Y_TYPEMVT','AE')
    else OBA.PutValue('Y_TYPEMVT','AL') ;

  OBA.PutValue('Y_TYPEANALYTIQUE','-') ;
//	TMemoField(TEcrAna.FindField('Y_BLOCNOTE')).Assign(TT.M[NumV]) ;
  OBA.PutValue('Y_DATEMODIF',				NowFutur) ;
  OBA.PutValue('Y_DATECOMPTABLE',		DatePiece) ;
  OBA.PutValue('Y_EXERCICE',				QuelExoDT(DatePiece)) ;
{$IFNDEF SPEC302}
	OBA.PutValue('Y_PERIODE',					GetPeriode(DatePiece)) ;
	OBA.PutValue('Y_SEMAINE',					NumSemaine(DatePiece)) ;
{$ENDIF}
  OBA.PutValue('Y_NATUREPIECE',E_NATUREPIECE.Value) ;
  OBA.PutValue('Y_ETABLISSEMENT',E_ETABLISSEMENT.Value) ;
  OBA.PutValue('Y_CONFIDENTIEL',ModeConf) ;
//  OBA.PutValue('Y_SAISIEEURO',CheckToString(ModeOppose)) ;
  if Not M.ANouveau then OBA.PutValue('Y_ECRANOUVEAU','N')
                    else OBA.PutValue('Y_ECRANOUVEAU','H') ;
  // SBO : Fiche 12085 : Refinterne et Libellé non reporté sur analytique en mode modif
  if action=taModif then
    begin
    OBA.PutValue('Y_REFINTERNE',  GS.Cells[SA_RefI,Lig] ) ;
    OBA.PutValue('Y_LIBELLE',     GS.Cells[SA_Lib,Lig]  ) ;
    end ;
END ;


Procedure TFeSaisie.RecupEche(Lig : integer ; R : T_ModeRegl ; NumEche : integer; Var OBM : TOBM  ) ;
Var Deb  : boolean ;
    i    : integer ;
    Coef : Double ;
    Coll : String ;
BEGIN
	if PasModif then Exit ; //Coef:=1.0 ;
  OBM.PutMvt('E_NUMECHE',NumEche) ;
  OBM.PutMvt('E_ECHE','X') ;
  OBM.PutMvt('E_NIVEAURELANCE',	R.TabEche[NumEche].NiveauRelance) ;
  OBM.PutMvt('E_MODEPAIE',			R.TabEche[NumEche].ModePaie) ;
  OBM.PutMvt('E_DATEECHEANCE',	R.TabEche[NumEche].DateEche) ;
  OBM.PutMvt('E_DATEVALEUR',		R.TabEche[NumEche].DateValeur) ;
  OBM.PutMvt('E_COUVERTURE',		R.TabEche[NumEche].Couverture) ;
  OBM.PutMvt('E_COUVERTUREDEV',	R.TabEche[NumEche].CouvertureDev) ;
  OBM.PutMvt('E_ETATLETTRAGE',	R.TabEche[NumEche].EtatLettrage) ;
  OBM.PutMvt('E_LETTRAGE',			R.TabEche[NumEche].CodeLettre) ;
  OBM.PutMvt('E_LETTRAGEDEV',		R.TabEche[NumEche].LettrageDev) ;
  OBM.PutMvt('E_DATEPAQUETMAX',	R.TabEche[NumEche].DatePaquetMax) ;
  OBM.PutMvt('E_DATEPAQUETMIN',	R.TabEche[NumEche].DatePaquetMin) ;
  OBM.PutMvt('E_DATERELANCE',		R.TabEche[NumEche].DateRelance) ;
  OBM.PutMvt('E_DEBIT',0) ;  		OBM.PutMvt('E_CREDIT',0) ;
  OBM.PutMvt('E_DEBITDEV',0) ;	OBM.PutMvt('E_CREDITDEV',0) ;
//  OBM.PutMvt('E_DEBITEURO',0) ;	OBM.PutMvt('E_CREDITEURO',0) ;
  Deb:=(ValD(GS,Lig)<>0) ;
  if Deb then
    BEGIN
    OBM.PutMvt('E_DEBIT',R.TabEche[NumEche].MontantP) ;
    OBM.PutMvt('E_DEBITDEV',R.TabEche[NumEche].MontantD) ;
    END
  else
    BEGIN
    OBM.PutMvt('E_CREDIT',R.TabEche[NumEche].MontantP) ;
    OBM.PutMvt('E_CREDITDEV',R.TabEche[NumEche].MontantD) ;
    END ;
  if ModeDevise then
    BEGIN
    if Deb	then OBM.PutMvt('E_DEBITDEV',R.TabEche[NumEche].MontantD)
	          else OBM.PutMvt('E_CREDITDEV',R.TabEche[NumEche].MontantD) ;
    END ;
  OBM.PutMvt('E_ENCAISSEMENT',QuelEnc(Lig)) ;
  OBM.PutMvt('E_ORIGINEPAIEMENT',R.TabEche[NumEche].DateEche) ;
{$IFDEF EAGLCLIENT}
  if (not SuiviMP) or (M.MPGUnique <> '') then SetMPACCDB(OBM) ; // FQ 12438
{$ENDIF}
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

Procedure TFeSaisie.GetAna(Lig : Integer) ;
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
      if revision then OBA.InsertDB(nil);
      END ;
END ;


Procedure TFeSaisie.EcartChange(Obj : TOB) ;
Var stPref : String;
BEGIN
	if PasModif then Exit ;
	if E_NATUREPIECE.Value<>'ECC' then Exit ;
  if Obj.NomTable = 'ECRITURE' then stPref := 'E'
  	else if Obj.NomTable = 'ANALYTIQ' then stPref := 'Y'
    	else Exit;
	Obj.PutValue(stPref+'_DEBITDEV',	0) ;
	Obj.PutValue(stPref+'_CREDITDEV',	0) ;
	Obj.PutValue(stPref+'_DEVISE',		E_DEVISE_.Value) ;
END ;

Procedure TFeSaisie.GetEcr(Lig : Integer) ;
Var M    : TMOD ;
    R    : T_ModeRegl ;
    i    : integer ;
    CGen : TGGeneral ;
    OBM  : TOBM;
BEGIN
	if PasModif then Exit ;
	M:=TMOD(GS.Objects[SA_NumP,Lig]) ;
  CGen:=GetGGeneral(GS,Lig) ;
  if CGen=Nil then Exit ;
	if ((M=Nil) or (Lettrable(CGen)=MonoEche)) then
  	BEGIN
		// Récupération et remplissage TOB Ecritures générales
		RecupTronc(Lig,M,OBM);
    EcartChange(OBM);
		// Finalisation des TOBs Analytiques
  	GetAna(Lig) ;
  	// Enregistrement Base
		OBM.InsertDBByNivel(False);
   	END
  else
  	BEGIN
  	R:=M.MODR ;
   	for i:=1 to R.NbEche do
	    BEGIN
			// Récupération et remplissage TOB Ecritures générales
      RecupTronc(Lig,M,OBM);
      RecupEche(Lig,R,i,OBM);
      EcartChange(OBM);
			// Finalisation des TOBs Analytiques
      if i=1 then	GetAna(Lig) ;
			// Enregistrement Base
      OBM.InsertDBByNivel(False);
    	END ;
  	END ;
END ;

procedure TFeSaisie.RecupFromGrid( Lig : integer ; MO : TMOD ; NumE : integer ) ;
Var OBM  : TOBM ;
    St   : String ;
    R    : T_MODEREGL ;
    ie,k,iEche : integer ;
    Sens : integer ;
    SQL  : String ;
    LeGen,LeAux,TL0 : String17 ;
    CGen : TGGeneral ;
    CAux : TGTiers ;
    NatT,NatG,StMvt : String3 ;
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
NatG:='' ; NatT:='' ; CGen:=Nil ; CAux:=Nil ;
if GS.Objects[SA_Gen,Lig]<>Nil then BEGIN CGen:=TGGeneral(GS.Objects[SA_Gen,Lig]) ; NatG:=CGen.NatureGene ; END ;
if GS.Objects[SA_Aux,Lig]<>Nil then BEGIN CAux:=TGTiers(GS.Objects[SA_Aux,Lig]) ; NatT:=CAux.NatureAux ; END ;
if NatT<>'' then BEGIN if CAux<>Nil then StMvt:=QuelTypeMvt(CAux.Auxi,NatT,E_NATUREPIECE.Value) ; END
            else BEGIN if CGen<>Nil then StMvt:=QuelTypeMvt(CGen.General,NatG,E_NATUREPIECE.Value) ; END ;
OBM.PutMvt('E_TYPEMVT','JLD') ; {Forcer la modif}
OBM.PutMvt('E_TYPEMVT',StMvt) ; OBM.PutMvt('E_VISION','DEM') ;
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

Procedure TFeSaisie.GetEcrGrid(Lig : Integer) ;
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

procedure TFeSaisie.ValideLignesGene ;
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
    END
	else
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

procedure TFeSaisie.ValideLeJournalNew ;
Var Per  : Byte ;
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
   AttribParamsNew(FRM,SI_TotDP,SI_TotCP,GeneTypeExo) ;
   END else
   BEGIN
   FRM.DE:=SI_TotDP ; FRM.CE:=SI_TotCP ;
   FRM.Deb:=0 ; FRM.Cre:=0 ;
   FRM.DS:=0 ; FRM.CS:=0 ;
   FRM.DP:=0 ; FRM.CP:=0 ;
   END ;
ll:=ExecReqMAJNew(fbJal,M.ANouveau,False,FRM) ;
If ll<>1 then V_PGI.IoError:=oeSaisie ;
{Dévalidation éventuelle période+jal}
if GeneTypeExo=teEncours then
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

procedure TFeSaisie.ValideLesComptes ;
BEGIN
  if PasModif then exit ;
  if M.Simul='N' then
    BEGIN
    MajCptesGeneNew ;
    MajCptesAuxNew ;
    MajCptesSectionNew ;
    END
  else
    BEGIN
{$IFDEF EAGLCLIENT}
{$ELSE}
    IntoucheGene ;
    IntoucheAux ;
{$ENDIF}
    END ;
END ;

procedure TFeSaisie.InverseSoldeNew ;
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
    Fillchar(FRM,SizeOf(FRM),#0) ;
    T:=T_SC(TS[i]) ;
    FRM.Cpt:=T.Cpte ; FRM.Deb:=T.Debit ; FRM.Cre:=T.Credit ; FRM.Axe:=T.Identi ;
    if FRM.Axe='G' then
       BEGIN
       AttribParamsNew(FRM,T.Debit,T.Credit,GeneTypeExo) ; ll:=ExecReqINVNew(fbGene,FRM) ;
       If ll<>1 then V_PGI.IoError:=oeSaisie ;
       END  ;
    END ;
// Tiers
for i:=0 to TS.Count-1 do
    BEGIN
    Fillchar(FRM,SizeOf(FRM),#0) ;
    T:=T_SC(TS[i]) ;
    FRM.Cpt:=T.Cpte ; FRM.Deb:=T.Debit ; FRM.Cre:=T.Credit ; FRM.Axe:=T.Identi ;
    if FRM.Axe='T' then
       BEGIN
       AttribParamsNew(FRM,T.Debit,T.Credit,GeneTypeExo) ; ll:=ExecReqINVNew(fbAux,FRM) ;
       If ll<>1 then V_PGI.IoError:=oeSaisie ;
       END ;
    END ;
// Sections
for i:=0 to TS.Count-1 do
    BEGIN
    Fillchar(FRM,SizeOf(FRM),#0) ;
    T:=T_SC(TS[i]) ;
    FRM.Cpt:=T.Cpte ; FRM.Deb:=T.Debit ; FRM.Cre:=T.Credit ; FRM.Axe:=T.Identi ;
    if ((FRM.Axe<>'G') and (FRM.Axe<>'T')) then
       BEGIN
       AttribParamsNew(FRM,T.Debit,T.Credit,GeneTypeExo) ; ll:=ExecReqINVNew(fbSect,FRM) ;
       If ll<>1 then V_PGI.IoError:=oeSaisie ;
       END ;
    END ;
// Journal
Fillchar(FRM,SizeOf(FRM),#0) ;
FRM.Cpt:=SAJAL.JOURNAL ; FRM.Deb:=SAJAL.OldDebit ; FRM.Cre:=SAJAL.OldCredit ;
AttribParamsNew(FRM,SAJAL.OldDebit,SAJAL.OldCredit,GeneTypeExo) ;
ll:=ExecReqINVNew(fbJal,FRM) ; If LL<>1 then V_PGI.IoError:=oeSaisie ;
END ;

procedure TFeSaisie.TvaTauxDirecteur ( Lig : integer ; Client,ToutDebit : boolean ; Regime : String3 ) ;
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


procedure TFeSaisie.GereLesImmos ( Lig : integer ) ;
Var O : TOBM ;
    CGen   : TGGeneral ;
    CAux   : TGTiers ;
    X      : Double ;
    LaLig  : integer ;
{$IFNDEF EAGLCLIENT}
{$IFNDEF CCMP}
{$IFNDEF GCGC}
{$IFNDEF IMP}
    RecImmo : TBaseImmo ;
{$ENDIF}
{$ENDIF}
{$ENDIF}
{$ENDIF}
    CodeImmo, LibAux : String ;
BEGIN
{$IFNDEF EAGLCLIENT}
{$IFNDEF IMP}
{$IFNDEF CCMP}
{$IFNDEF GCGC}
{$IFDEF SANSIMMO}
Exit ;
{$ELSE}
if ((Not VH^.OkModImmo) and (Not V_PGI.VersionDemo)) then Exit ;
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
O:=GetO(GS,Lig) ; if O=Nil then Exit ;
if O.GetMvt('E_NUMEROIMMO')=666 then Exit else O.PutMvt('E_NUMEROIMMO',666) ;
if O.GetMvt('E_IMMO')<>'' then Exit ;
if HPiece.Execute(46,Caption,'')<>mrYes then Exit ;
FillChar(RecImmo,Sizeof(RecImmo),#0) ;
X:=O.GetMvt('E_DEBIT')+O.GetMvt('E_CREDIT') ;
LaLig:=TrouveLigTiers(GS,0) ; LibAux:='' ;
if LaLig>0 then BEGIN CAux:=GetGTiers(GS,LaLig) ; if CAux<>Nil then LibAux:=CAux.Libelle ; END ;
RecImmo.CompteImmo:=CGen.General ; RecImmo.Fournisseur:=LibAux ;
RecImmo.Reference:=GS.Cells[SA_RefI,Lig] ;
RecImmo.Libelle:=GS.Cells[SA_Lib,Lig] ;
RecImmo.CodeEtab:=E_ETABLISSEMENT.Value ;
RecImmo.DateAchat:=DatePiece ;
RecImmo.MontantHT:=X ;
RecImmo.MontantTVA:=HT2TVA(X,GeneRegTVA,GeneSoumisTPF,O.GetMvt('E_TVA'),O.GetMvt('E_TPF'),True,DECDEV) ;
RecImmo.Quantite:=O.GetMvt('E_QTE1') ; if RecImmo.Quantite=0 then RecImmo.Quantite:=1 ;
Codeimmo:=CreationSaisieCompta(RecImmo) ;
O.PutMvt('E_IMMO',CodeImmo) ;
{$ENDIF}
{$ENDIF}
{$ENDIF}
{$ENDIF}
END ;

procedure TFeSaisie.GereTvaEncNonFact ( Lig : integer ) ;
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

procedure TFeSaisie.AfficheExigeTva ;
Var Nat : String3 ;
BEGIN
if Not OuiTvaLoc then BEGIN BPopTva.Visible:=False ; Exit ; END ;
Nat:=E_NATUREPIECE.Value ;
if ((Nat<>'FC') and (Nat<>'AC') and (Nat<>'OC') and
    (Nat<>'FF') and (Nat<>'AF') and (Nat<>'OF')) then BEGIN BPopTva.Visible:=False ; Exit ; END ;
if (((Nat='OC') or (Nat='FC') or (Nat='AC')) and (VH^.TvaEncSociete='TD')) then BEGIN BPopTva.Visible:=False ; Exit ; END ;
if (((Nat='FC') or (Nat='AC')) and (SorteTva<>stvVente)) then BEGIN BPopTva.Visible:=False ; Exit ; END ;
if (((Nat='FF') or (Nat='AF')) and (SorteTva<>stvAchat)) then BEGIN BPopTva.Visible:=False ; Exit ; END ;
BPopTva.Visible:=True ;
END ;

procedure TFeSaisie.GetSorteTva ;
BEGIN
SorteTva:=stvDivers ; if SAJAL=Nil then Exit ;
if ((SAJAL.NatureJal='VTE') and ((E_NATUREPIECE.Value='FC') or (E_NATUREPIECE.Value='AC'))) then SorteTva:=stvVente else
 if ((SAJAL.NatureJal='ACH') and ((E_NATUREPIECE.Value='FF') or (E_NATUREPIECE.Value='AF'))) then SorteTva:=stvAchat ;
END ;

procedure TFeSaisie.RenseigneRegime ( Lig : integer ; Recharge : boolean ) ;
Var CGen : TGGeneral ;
    CAux : TGTiers ;
BEGIN
if Lig<=0 then Exit ;
if GS.Cells[SA_Aux,Lig]<>'' then
   BEGIN
   CAux:=GetGTiers(GS,Lig) ;
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
      GeneSoumisTPF:=CGen.SoumisTPF ;
      END ;
   END ;
END ;

procedure TFeSaisie.AttribRegimeEtTVA ;
Var ia : integer ;
    CGen : TGGeneral ;
//    CAux : TGTiers ;
BEGIN
ia:=TrouveLigTiers(GS,0) ;
if ia>0 then RenseigneRegime(ia,False) ;
ia:=TrouveLigHT(GS,0,False) ; if ia<=0 then exit ;
if GS.Cells[SA_Gen,ia]<>'' then
   BEGIN
   CGen:=GetGGeneral(GS,ia) ;
   if CGen<>Nil then BEGIN GeneTVA:=CGen.TVA ; GeneTPF:=CGen.TPF ; END ;
   END ;
END ;

procedure TFeSaisie.DetruitAncien ;
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
      END
    else
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
    END
  else
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

procedure TFeSaisie.EnvoiCFONB ;
Var RR    : RMVT ;
    Q     : TQuery ;
    SQL   : String ;
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
    END
  else
    OkEnc:=isEncMP(M.smp) ;
  if HPiece.Execute(35,Caption,'')<>mrYes then Exit ;
  RR:=P_MV(TPIECE[TPIECE.Count-1]).R ;
  TEche:=TList.Create ;
  SQL := 'Select * from Ecriture where '
          + WhereEcriture(tsGene,RR,False)
          + ' AND E_AUXILIAIRE<>"" AND E_ECHE="X" AND E_NUMECHE>0 AND E_CFONBOK="#"'; // FQ 12393
  Q:=OpenSQL(SQL,True) ;
  While Not Q.EOF do
    BEGIN
    O := TOBM.Create(EcrGen,'',False) ;
    O.ChargeMvt(Q) ;
    TEche.Add(O) ;
    Q.Next ;
    END ;
  Ferme(Q) ;
{$IFNDEF IMP}
  {$IFNDEF GCGC}
    {$IFNDEF CCS3}
  If Not VH^.OldTeleTrans Then M.EnvoiTrans:='' ;
  if TEche.Count>0 then
  begin
    ExportCFONB(OkEnc,M.General,M.FormatCFONB,M.EnvoiTrans,TECHE,M.smp) ;
    // VL 30102003 FQ 12958
    SQL:='Update Ecriture Set E_CFONBOK="X" where '+WhereEcriture(tsGene,RR,False)+' AND E_AUXILIAIRE<>"" AND E_ECHE="X" AND E_NUMECHE>0 AND E_CFONBOK="#"' ;
    ExecuteSQL(SQL) ;
    // FIN VL
  end;
    {$ENDIF}
  {$ENDIF}
{$ENDIF}
  VideListe(TEche) ;
  TEche.Free ;
END ;

procedure TFeSaisie.EnvoiBordereau ;
Var RR    : RMVT ;
    Q     : TQuery ;
    SQL,SWhere   : String ;
    O     : TOBM ;
    TEche : TList ;
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
SQL:='Select * from Ecriture where '+WhereEcriture(tsGene,RR,False)+' AND E_AUXILIAIRE<>"" AND E_ECHE="X" AND E_NUMECHE>0 AND E_CFONBOK="#"' ;
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
   SQL:='Update Ecriture Set E_CFONBOK="X" where '+WhereEcriture(tsGene,RR,False)+' AND E_AUXILIAIRE<>"" AND E_ECHE="X" AND E_NUMECHE>0 AND E_CFONBOK="#"' ;
   ExecuteSQL(SQL) ;
   END ;
VideListe(TEche) ; TEche.Free ;
END ;

procedure TFeSaisie.LettrerReglement ;
Var RR  : RMVT ;
BEGIN
RR:=P_MV(TPIECE[TPIECE.Count-1]).R ;
// Modif Fiche 12017
  LettrageEncaDeca(TECHEORIG,RR,DEV,TOBNumChq) ;  // ajout tobNumChq SBO maj n°chq depuis CCMP
  FreeAndNil(TOBNumChq) ;                         // libération TobNumChq
// Fin Modif Fiche 12017
END ;

procedure TFeSaisie.TraiteMP ;
Var i : integer ;
    O : TOBM ;
BEGIN
M.Simul:='N' ;
for i:=1 to GS.RowCount-1 do
    BEGIN
    O:=GetO(GS,i) ; if O=Nil then Break ;
    O.PutMvt('E_QUALIFPIECE','N') ;
    O.PutMvt('E_NUMENCADECA',M.NumEncaDeca); // VL 17122003 FFF
    END ;
END ;

procedure TFeSaisie.ValideLaPiece ;
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

procedure TFeSaisie.ValideLeReste ;
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

{==========================================================================================}
{================================ Barre d'outils ==========================================}
{==========================================================================================}
procedure TFeSaisie.ForcerMode ( Cons : boolean ; Key : Word ) ;
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

procedure TFeSaisie.NormaliserEscomptes ;
Var i : integer ;
    TOBL : TOB ;
    YY   : RMVT ;
    T    : TStrings ;
    St   : String ; 
BEGIN
  if Not SuiviMP then Exit ;
  if TOBMPEsc=Nil then Exit ;
  if TOBMPEsc.Detail.Count<=0 then Exit ;
  for i:=0 to TOBMPEsc.Detail.Count-1 do
    BEGIN
    TOBL:=TOBMPEsc.Detail[i] ;
    if Not NormaliserPieceSimu(TOBL,True) then
      BEGIN
      V_PGI.IoError := oeUnknown ;
      Break ;
      END ;
    T:=TStringList.Create ;
    YY:=TOBToIdent(TOBL,True) ;
    St:=EncodeLC(YY) ;
    T.Add(St) ;
    TECHEORIG.Add(T) ;
    END ;
END ;

procedure TFeSaisie.SuiteEncaDeca ;
BEGIN
  GS.VidePile(True) ;
  if SuiviMP then
    if Transactions( NormaliserEscomptes, 1) <> oeOk
      then MessageAlerte(HDiv.Mess[25]) ;
  if Transactions( LettrerReglement, 1) <> oeOk
    then MessageAlerte(HDiv.Mess[19]) ;
  EnvoiCFONB ;
  EnvoiBordereau ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 25/09/2002
Modifié le ... :   /  /
Description .. : - LG - 25/09/2002 - l'appel de MajIOCom est conditionne
Suite ........ : par par le test OkSynchro
Mots clefs ... :
*****************************************************************}
procedure TFeSaisie.ClickValide ;
Var StE : String3 ;
    io,io1  : TIOErr ;
    TesterLig : boolean ;
BEGIN
if Not GS.SynEnabled then Exit ;
if ((M.TypeGuide='NOR') and (M.FromGuide)) then BEGIN ClickAbandon ; Exit ; END ;
If Guide And (M.TypeGuide<>'DEC') and (M.TypeGuide<>'ENC') And FlashGuide.Visible Then
  BEGIN
  HDiv.Execute(26,Caption,'') ; Exit ;
  END ;
{$IFDEF CEGID}
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
{$ENDIF}
if ((E_JOURNAL.Value<>'') and (GS.RowCount<=2) and (GS.CanFocus) and
    (Screen.ActiveControl<>GS) and (Screen.ActiveControl<>Valide97)) then BEGIN GS.SetFocus ; Exit ; END ;
{Contrôles avant validation}
if Action=taConsult then BEGIN CloseFen ; Exit ; END ;
if ((Not BValide.Enabled) and (CurM<=0)) then exit ;
{$IFDEF CEGID}
TraiteLigneAZero ;
{$ENDIF}
ValideCeQuePeux(GS.Row) ;
TesterLig:=((EstRempli(GS,GS.Row)) or (Guide)) ;
if ((TesterLig) and (Not LigneCorrecte(GS.Row,False,(Action=taCreat) or (Not Revision) or (UnChange)))) then Exit ;
if ((TesterLig) and (DEV.Code<>V_PGI.DevisePivot) and (Action<>taConsult) and (Not Revision) and (EstRempli(GS,GS.Row)))
   then CalculDebitCredit ;
if not PieceCorrecte then Exit ;
if not EquilEuroValide then Exit ;
if Not ControleRIB then Exit ;
If Not OkSynchro Then MajIoCom(GS) ;
if ((Action=taCreat) and (H_NUMEROPIECE_.Visible) and (GS.RowCount>2)) or (SuiviMP) then AttribNumeroDef ;
if GS.RowCount<>NbLOrig then Revision:=False ;
GS.Row:=1 ; GS.Col:=SA_Gen ; GS.SetFocus ; BValide.Enabled:=False ;
AttribRegimeEtTVA ;
StE:=QuelExoDT(StrToDate(E_DATECOMPTABLE.Text)) ;
if StE=VH^.Encours.Code then GeneTypeExo:=teEncours else
   if StE=VH^.Suivant.Code then GeneTypeExo:=teSuivant else GeneTypeExo:=tePrecedent ;
{Scenario avant validation}
if ((OkScenario) and (Scenario.GetMvt('SC_CONTROLETVA')='OUI')) then if Not ControleTVAOK then Exit ;
{Validation}
DatePiece:=StrToDate(E_DATECOMPTABLE.Text) ;
NowFutur:=NowH ; GS.Enabled:=False ;
io:=Transactions(ValideLaPiece,5) ;
io1:=Transactions(ValideLeReste,10) ;
If io1<>oeOK Then MessageAlerte(HDiv.Mess[22]+HDiv.Mess[23]) ;
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
PieceModifiee:=False ; ModeLC:=False ;
if ((Action=taCreat) and (M.LeGuide='') and (Not M.ANouveau)) then
   BEGIN
   if VH^.BouclerSaisieCreat then ReInitPiece(False) else
      BEGIN
      CloseFen ;
//      if Not IsInside(Self) then CloseFen else CloseInsidePanel(Self) ;
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
            END else CloseFen ;
         END else CloseFen ;
      END else If M.MSED.MultiSessionEncaDeca Then CloseFen
            Else if Not ModifNext then CloseFen ;
   END ;
END ;

procedure TFeSaisie.ClickVentil ;
Var OkL : Boolean ;
BEGIN
if GS.Row>=GS.RowCount-1 then Exit ;
if Not EstRempli(GS,GS.Row) then Exit ;
OkL:=LigneCorrecte(GS.Row,False,True) ;
if OkL then GereAnal(GS.Row,False,False) ;
END ;

procedure TFeSaisie.ClickEche ;
Var OkL : Boolean ;
BEGIN
if GS.Row>=GS.RowCount-1 then Exit ;
if Not EstRempli(GS,GS.Row) then Exit ;
OkL:=LigneCorrecte(GS.Row,False,True) ;
if OkL then BEGIN GereEche(GS.Row,False,False) ; Revision:=False ; END ;
END ;

procedure TFeSaisie.ClickDel ;
BEGIN
if PasModif then Exit ;
if ModeLC then Exit ;
if GS.RowCount<=3 then Exit ;
if ((GS.Row=GS.RowCount-1) and (Not EstRempli(GS,GS.Row))) then BEGIN Gs.SetFocus ; Exit ; END ;
if PasToucheLigne(GS.Row) then Exit ;
GS.SynEnabled:=False ;
if Guide then DetruitLigneGuide(GS.Row);
DetruitLigne(GS.Row,True) ; Revision:=False ;
ControleLignes ; GereOptionsGrid(GS.Row) ;
GS.SynEnabled:=True ;
END ;

procedure TFeSaisie.ClickInsert ;
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

procedure TFeSaisie.ClickAbandon ;
BEGIN
  if ModeForce=tmPivot then
    ClickSwapPivot
  else
    BEGIN
    ModeLC:=False ;
    CloseFen ;
    END ;
END ;

procedure TFeSaisie.ClickSolde (Egal,RowEx : boolean) ;
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
BValide.Enabled:=((H_NumeroPiece.Visible) and (TRUE)) ;
END ;

Function TFeSaisie.SortePiece : boolean ;
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

function TFeSaisie.ClickControleTVA ( Bloque : boolean ; LigHT : integer ) : boolean ;
{$IFNDEF GCGC}
{$IFNDEF CCS3}
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
{$ENDIF}
{$ENDIF}
BEGIN
{$IFNDEF CCS3}
{$IFNDEF GCGC}
Result:=(Not Bloque) ; if ((Not BControleTVA.Enabled) and (Not Bloque)) or (M.ANouveau) then Exit ;
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
if CTVA<>'' then BEGIN LigTVA:=TrouveLigCompte(GS,SA_Gen,0,CTVA) ; if LigTVA>0 then InfoLigne(LigTVA,SDA,SCA,STDA,STCA) ; END ;
CDA:=HT2TVA(SDH,GeneRegTVA,GeneSoumisTPF,CodeTva,CodeTpf,Achat,DECDEV) ;
CCA:=HT2TVA(SCH,GeneRegTVA,GeneSoumisTPF,CodeTva,CodeTpf,Achat,DECDEV) ;
CTDA:=HT2TVA(STDH,GeneRegTVA,GeneSoumisTPF,CodeTva,CodeTpf,Achat,DECDEV) ;
CTCA:=HT2TVA(STCH,GeneRegTVA,GeneSoumisTPF,CodeTva,CodeTpf,Achat,DECDEV) ;
// Calculs et recup de la ligne TPF
if CTPF<>'' then BEGIN LigTPF:=TrouveLigCompte(GS,SA_Gen,0,CTPF) ; if LigTPF>0 then InfoLigne(LigTPF,SDF,SCF,STDF,STCF) ; END ;
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
{$ENDIF}
END ;

procedure TFeSaisie.RegroupeTVA ;
Var LigTVA,LigNext,NbT,NbN : integer ;
    CpteTVA : String17 ;
    DTVA,CTVA,DNext,CNext : Double ;
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
    END ;
Until ((LigTVA<=0) or (LigTVA>=GS.RowCount-1) or (NbT>50)) ;
END ;

procedure TFeSaisie.ClickGenereTVA ;
Var Nb,i,iAux,iHT,iTva : integer ;
    RegTVA,TvaEnc : String3 ;
    SoumisTPF,Achat  : Boolean ;
    CAux       : TGTiers ;
    CGen       : TGGeneral ;
    OBM        : TOBM ;
    OkCpteTVA : boolean ;
    Cpte       : String17 ;
BEGIN
if PasModif then Exit ;
if ModeLC then Exit ;
if ((M.ANouveau) or (Not BGenereTVA.Enabled)) then Exit ;
OkMessAnal:=False ; CoherTVA:=True ; OkCpteTVA:=True ;
// Contrôles de présence 1 seul TTC et au moins 1 HT
if ((EstRempli(GS,GS.Row)) and (Not LigneCorrecte(GS.Row,False,True))) then exit ;
if ((SAJAL.NATUREJAL<>'ACH') and (SAJAL.NATUREJAL<>'VTE')) then BEGIN HTva.Execute(0,caption,'') ; Exit ; END ;
Achat:=SortePiece ;
iAux:=TrouveLigTiers(GS,0) ; if iAux<=0 then BEGIN HTVA.Execute(1,caption,'') ; Exit ; END ;
if TrouveLigTiers(GS,iAux)>0 then BEGIN HTVA.Execute(2,caption,'') ; Exit ; END ;
iHT:=TrouveLigHT(GS,0,TRUE) ; if iHT<=0 then BEGIN HTVA.Execute(3,caption,'') ; Exit ; END ;
if GS.Cells[SA_Aux,iAux]<>'' then
   BEGIN
   CAux:=GetGTiers(GS,iAux) ; RegTVA:=GeneRegTVA ; SoumisTpf:=CAux.SoumisTPF ;
   if Achat then TvaEnc:=CAux.TVA_Encaissement else TvaEnc:=VH^.TvaEncSociete ;
   Cpte:=GS.Cells[SA_Aux,iAux] ;
   END else
   BEGIN
   CGen:=GetGGeneral(GS,iAux) ; RegTVA:=GeneRegTVA ; SoumisTpf:=CGen.SoumisTPF ;
   if Achat then TvaEnc:=CGen.TVA_Encaissement else TvaEnc:=VH^.TvaEncSociete ;
   Cpte:=GS.Cells[SA_Gen,iAux] ;
   END ;
if RegTVA='' then BEGIN HTVA.Execute(6,caption,Cpte) ; Exit ; END ;
iHT:=0 ; Nb:=0 ;
// Remplissage des comptes HT avec infos de TVA/TPF associées
Repeat
 iHT:=TrouveLigHT(GS,iHT,True) ; Inc(Nb) ;
 if iHT>0 then AlimHTByTVA(iHT,RegTVA,TvaEnc,SoumisTpf,Achat) ;
Until ((iHT<=0) or (Nb>100)) ;
GS.CacheEdit ;
// Destruction des lignes TVA/TPF concernées déja présentes
Nb:=0 ;
Repeat
 iTva:=TrouveLigTVA(GS,0) ; Inc(Nb) ;
 if iTva>0 then DetruitLigne(iTva,False) ;
Until ((iTva<=0) or (Nb>100)) ;
Nb:=GS.RowCount-2 ;
// Génération des lignes TVA/TPF
for i:=1 to Nb do if isHT(GS,i,True) then
    BEGIN
    OBM:=GetO(GS,i) ; OBM.PutMvt('E_REGIMETVA',RegTVA) ;
    if Not InsereLigneTVA(i,RegTVA,SoumisTpf,Achat) then OkCpteTVA:=False ;
    END ;
// Regroupement selon scénario
if ((OkScenario) and (Scenario.GetMvt('SC_CONTROLETVA')='GTG')) then RegroupeTVA ;
// Fin et ré-init du traitement
CalculDebitCredit ; GS.MontreEdit ;
BValide.Enabled:=((H_NumeroPiece.Visible) and (Equilibre)) ;
if OkMessAnal then HTVA.Execute(4,caption,'')  ;
if Not OkCpteTVA then HTVA.Execute(7,caption,'') else
   if Not CoherTVA then HTVA.Execute(5,caption,'')  ;
Revision:=False ; OkMessAnal:=False ; CoherTVA:=True ;
END ;

procedure TFeSaisie.ClickZoom ;
Var b : byte ;
    C,R : longint ;
    A   : TActionFiche ;
BEGIN
if ((E_JOURNAL.Value='') or (Not GS.Enabled)) then Exit ;
if ((Action=taConsult) or (ModeForce<>tmNormal))
    then BEGIN GS.MouseToCell(GX,GY,C,R) ; A:=taConsult ; END
    else BEGIN R:=GS.Row ; C:=GS.Col ; A:=taModif ; END ;
if R<=0 then BEGIN R:=GS.Row ; C:=SA_Gen ; END ;
if R<1 then Exit ; if C<SA_Gen then Exit ;
if ((Action=taConsult) or (ModeLC)) and (GS.Cells[C,R]='') then Exit ;
if ModeForce<>tmNormal then Exit ;
if C=SA_Gen then
   BEGIN
   {$IFNDEF GCGC}
   if ((Not ExJaiLeDroitConcept(TConcept(ccGenModif),False)) or (ModeLC))
     then A:=taConsult ;
   b:=ChercheGen(C,R,False) ;
   if b=2 then
     begin
     FicheGene(Nil,'',GS.Cells[C,R],A,0) ;
     if Action<>taConsult then ChercheGen(C,R,True) ;
     end ;
   {$ELSE}
   Exit ;
   {$ENDIF}
   END else if C=SA_Aux then
   BEGIN
   if ((Not ExJaiLeDroitConcept(TConcept(ccAuxModif),False)) or (ModeLC))
     then A:=taConsult ;
   b:=ChercheAux(C,R,False) ;
   {$IFNDEF GCGC}
   if b=2 then
     BEGIN
     {$IFDEF EAGLCLIENT}
       FicheTiers(GS.Cells[C,R],A,1) ;
     {$ELSE}
       FicheTiers(Nil,'',GS.Cells[C,R],A,1) ;
     {$ENDIF}
     if Action<>taConsult then ChercheAux(C,R,True) ;
     END ;
   {$ELSE}
   if b=2 then
     AGLLanceFiche('GC','GCTIERS','',GS.Cells[C,R],ActionToString(taConsult)+';MONOFICHE;T_NATUREAUXI=CLI') ;
   {$ENDIF}
   END ;
END ;

procedure TFeSaisie.ClickConsult;
{$IFNDEF GCGC}
{$IFNDEF CCMP}
{$IFNDEF IMP}
var
    SG,LeExo : String;
{$ENDIF}
{$ENDIF}
{$ENDIF}
begin
{$IFNDEF GCGC}
{$IFNDEF CCMP}
{$IFNDEF IMP}
if ModeForce<>tmNormal then Exit ;
if ModeLC then Exit ;
if ((E_JOURNAL.Value='') or (Not GS.Enabled)) then Exit ;
SG:=GS.Cells[SA_Gen,GS.Row] ; if SG='' then Exit ;
if QuelExoDT(StrToDate(E_DATECOMPTABLE.Text))=VH^.Encours.Code then LeExo:='0' else LeExo:='1' ;
OperationsSurComptes(SG,LeExo,'',GS.Cells[SA_Aux,GS.Row],True) ;
{$ENDIF}
{$ENDIF}
{$ENDIF}
end;

procedure TFeSaisie.ClickSwapPivot ;
Var Cancel,Chg : boolean ;
BEGIN
if ((Not ModeDevise) and (ModeForce=tmNormal)) then Exit ;
if ModeLC then Exit ;
if Not BSwapPivot.Enabled then Exit ;
if ((EstRempli(GS,GS.Row)) and (Not LigneCorrecte(GS.Row,False,True))) then Exit ;
if Not EstRempli(GS,1) then Exit ;
if ModeForce=tmNormal then
   BEGIN
   ModeForce:=tmPivot ; OldDevise:=DEV.Code ; OldTauxDEV:=DEV.Taux ; E_DEVISE.Value:=V_PGI.DevisePivot ;
   DEV.Decimale:=V_PGI.OkDecV ; DECDEV:=V_PGI.OkDecV ; E_DEVISE.Enabled:=True ; E_Devise.Font.Color:=clRed ;
   AffecteGrid(GS,taConsult) ; GS.SetFocus ;
   PEntete.Enabled:=False ; Outils.Enabled:=False ; Valide97.Enabled:=False ;
   END else
   BEGIN
   E_DEVISE.Enabled:=False ; E_DEVISE.Value:=OldDevise ; DEV.Taux:=OldTauxDEV ; E_Devise.Font.Color:=clBlack ;
   ModeForce:=tmNormal ;
   PEntete.Enabled:=True ; Outils.Enabled:=True ;
   if Outils.CanFocus then Outils.SetFocus ; Valide97.Enabled:=True ;
   GS.Col:=SA_Gen ; GS.Row:=1 ;
   GS.CacheEdit ; AffecteGrid(GS,Action) ;
   GS.Col:=SA_Gen ; GS.Row:=2 ;
   GS.SetFocus ; GS.MontreEdit ;
   GS.Row:=1 ; Cancel:=False ; Chg:=False ; GSRowEnter(Nil,GS.Row,Cancel,Chg) ;
   END ;
ChangeAffGrid(GS,ModeForce,DECDEV) ;
CalculDebitCredit ;
END ;

(*
procedure TFeSaisie.ClickSwapEuro ;
//Var Cancel,Chg : boolean ;
BEGIN
{$IFNDEF GCGC}
if ModeForce<>tmNormal then Exit ;
if ModeLC then Exit ;
if Not BSwapEuro.Enabled then Exit ;
if Action<>taConsult then if ((EstRempli(GS,GS.Row)) and (Not LigneCorrecte(GS.Row,False,False))) then Exit ;
CorrigeEcartConvert(DEV,GS,taConsult) ;
{$ENDIF}
END ;
*)

procedure TFeSaisie.ClickModifRIB ;
Var O : TOBM ;
    IsAux : Boolean ;
BEGIN
if Not BModifRIB.Enabled then Exit ;
O:=GetO(GS,GS.Row) ; if O=Nil then Exit ;
IsAux:=O.GetMvt('E_AUXILIAIRE')<>'' ;
ModifRIBOBM(O,False,FALSE,'',IsAux) ;
END ;

procedure TFeSaisie.ClickComplement ;
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
RC.Attributs:=False ; RC.MemoComp:=Nil ; RC.Origine:=-1 ;
if Not SaisieComplement(O,EcrGen,Action,ModBN,RC) then Exit ;
if GS.CanFocus then GS.SetFocus ;
if MODBN then Revision:=False ;
NQ1:=O.GetMvt('E_QTE1') ; NQ2:=O.GetMvt('E_QTE2') ;
if ((NQ1<>OQ1) and (OQ1<>0)) then ProrateQteAnal(GS,Lig,OQ1,NQ1,1) ;
if ((NQ2<>OQ2) and (OQ2<>0)) then ProrateQteAnal(GS,Lig,OQ2,NQ2,2) ;
{$ENDIF}
end;

Function TFeSaisie.ChoixFromTenueSaisie ( LeMontantPIVOT,LeMontantEURO,LeMontantDEV : Double ) : Double ;
BEGIN
if DEV.Code<>V_PGI.DevisePivot then Result:=LeMontantDEV {else
   if ModeOppose then Result:=LeMontantEURO}
      else Result:=LeMontantPIVOT ;
END ;

{ // Ancien code pour mémoire, en attente validation nouvelle procédure
Procedure TFeSaisie.SetNumTraiteDOCREGLE ;
Var Q : TQuery ;
    i,j,NbPass : integer ;
    T        : TStrings ;
    RRMM     : RMVT ;
    MP,MD,ME,TotE,TotP,TotD : Double ;
    TOBEcr,TOBE,TOBDOC,TOBD : TOB ;
    sGen,sAux,sTra : String ;
BEGIN
if (M.smp<>smpEncTraEdtNC) And (M.smp<>smpDecBorEdtNC) then Exit ;
if ((Action<>taCreat) and (Not SuiviMP)) then Exit ;
if M.SorteLettre=tslAucun then Exit ;
if ModeLC then Exit ;
MP:=0;MD:=0;ME:=0;
//Vider toutes les lignes
ClickVidePiece(False) ;
TOBECR:=TOB.Create('',Nil,-1) ;
InitMove(TEcheOrig.Count*4,'') ;
//Lire les echéances d'origine
for i:=0 to TEcheOrig.Count-1 do
    BEGIN
    MoveCur(False) ;
    T:=TStrings(TEcheOrig[i]) ;
    for j:=0 to T.Count-1 do
        BEGIN
        RRMM:=DecodeLC(T[j]) ;
        Q:=OpenSQl('Select E_GENERAL,E_AUXILIAIRE,E_DEBIT,E_CREDIT,E_DEBITDEV,E_CREDITDEV,E_DEBITEURO,E_CREDITEURO from Ecriture Where '+WhereEcriture(tsGene,RRMM,True),True) ;
        if Not Q.EOF then
           BEGIN
           TOBE:=TOB.Create('',TOBECR,-1) ; TOBE.SelectDB('',Q) ;
           TOBE.AddChampSup('MMORIG',False) ; TOBE.PutValue('MMORIG',T[j]) ;
           END ;
        Ferme(Q) ;
        END ;
    END ;
//Lire le fichiers des échéances éditées
TOBDOC:=TOB.Create('',Nil,-1) ;
Q:=OpenSQL('Select * from DOCREGLE Where DR_USER="'+V_PGI.User+'" AND DR_EDITE="X" ORDER BY DR_NUMCHEQUE',True) ;
if Not Q.EOF then
   BEGIN
   ModeLC:=True ;
   While Not Q.EOF do
      BEGIN
      MoveCur(False) ;
      TOBD:=TOB.Create('',TOBDOC,-1) ;
      TOBD.AddChampSup('E_GENERAL',False)      ; TOBD.PutValue('E_GENERAL',Q.FindField('DR_GENERAL').AsString) ;
      TOBD.AddChampSup('E_AUXILIAIRE',False)   ; TOBD.PutValue('E_AUXILIAIRE',Q.FindField('DR_AUXILIAIRE').AsString) ;
      TOBD.AddChampSup('E_MONTANT',False)      ; TOBD.PutValue('E_MONTANT',Q.FindField('DR_MONTANT').AsFloat) ;
      TOBD.AddChampSup('E_MONTANTDEV',False)   ; TOBD.PutValue('E_MONTANTDEV',Q.FindField('DR_MONTANTDEV').AsFloat) ;
      TOBD.AddChampSup('E_MONTANTEURO',False)  ; TOBD.PutValue('E_MONTANTEURO',Q.FindField('DR_MONTANTEURO').AsFloat) ;
      TOBD.AddChampSup('E_NUMTRAITECHQ',False) ; TOBD.PutValue('E_NUMTRAITECHQ',GetNumLotTraChq(Q.FindField('DR_NUMCHEQUE').AsString)) ;
      TOBD.AddChampSup('MARQUE',False)         ; TOBD.PutValue('MARQUE','-') ;
      Q.Next ;
      END ;
   END ;
Ferme(Q) ;
//Faire le Mapping
for NbPass:=1 to 2 do
   BEGIN
   //Passage 1 pour 1
   // Pour chaques documents...
   for i:=0 to TOBDOC.Detail.Count-1 do
   	if V_PGI.ioError=OeOk then
       BEGIN
       MoveCur(False) ;
       TOBD:=TOBDOC.Detail[i] ;
       if TOBD.GetValue('MARQUE')='X' then Continue ;
       sGen:=TOBD.GetValue('E_GENERAL') ;
       sAux:=TOBD.GetValue('E_AUXILIAIRE') ;
       sTra:=TOBD.Getvalue('E_NUMTRAITECHQ') ;
       if NbPass=1 then
          BEGIN
          MP:=TOBD.GetValue('E_MONTANT') ;
          MD:=TOBD.GetValue('E_MONTANTDEV') ;
          ME:=TOBD.GetValue('E_MONTANTEURO') ;
          END ;
		   // Pour chaques écritures...
       for j:=TOBECR.Detail.Count-1 downto 0 do
           BEGIN
           TOBE:=TOBECR.Detail[j] ;
           // Vérif positionnement sur la bonne écriture
           if TOBE.GetValue('E_GENERAL')<>sGen then Continue ;
           if TOBE.GetValue('E_AUXILIAIRE')<>sAux then Continue ;
           // Au premier passage, le traitement n'a lieu que si Mono-Ech
           if NbPass=1 then
              BEGIN
              TotP:=TOBE.GetValue('E_DEBIT')+TOBE.GetValue('E_CREDIT') ;
              TotD:=TOBE.GetValue('E_DEBITDEV')+TOBE.GetValue('E_CREDITDEV') ;
              TotE:=TOBE.GetValue('E_DEBITEURO')+TOBE.GetValue('E_CREDITEURO') ;
              if Arrondi(MP-TotP,V_PGI.OkDecV)<>0 then Continue ;
              if Arrondi(MD-TotD,DEV.Decimale)<>0 then Continue ;
              if Arrondi(ME-TotE,V_PGI.OkDecE)<>0 then Continue ;
              END ;
					 // MAJ Ecriture : a lieu au premier passage si Mono-Ech,
           //								 au 2eme passage sinon
           RRMM:=DecodeLC(TOBE.GetValue('MMORIG')) ;
           if ExecuteSQL('UPDATE ECRITURE SET E_NUMTRAITECHQ="'+sTra+'" WHERE '+WhereEcriture(tsGene,RRMM,True))<>1 then
              BEGIN
              V_PGI.IoError:=oeUnknown ;
              Break ;
              END ;
           //Marques
           TOBE.Free ;
           // Si on est encore au premier passage c'est qu'on a trouvé l'ecriture
           // en Mono-Ech correspondant au document, on peut quitter la boucle sur
           // les écritures après avoir marquer document
           if NbPass=1 then BEGIN TOBD.PutValue('MARQUE','X') ; Break ; END ;
           END ;
       // Au 2ème passage, on marque auto le document après MAJ toutes les écritures
       if NbPass=2 then BEGIN TOBD.PutValue('MARQUE','X') ; END ;
       END ;
   END ;
//Supprimer les enregistrements temporaires
ExecuteSQL('DELETE from DOCREGLE Where DR_USER="'+V_PGI.USER+'"') ;
ExecuteSQL('DELETE from DOCFACT Where DF_USER="'+V_PGI.USER+'"') ;
FiniMove ;
TOBECR.Free ; TOBDOC.Free ;
END ;
}

Function TFeSaisie.EcheancesLC ( NbC : integer ) : boolean ;
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
      LeMontant:=ChoixFromTenueSaisie(Q.FindField('DR_MONTANT').AsFloat,Q.FindField('DR_MONTANTDEV').AsFloat,Q.FindField('DR_MONTANTDEV').AsFloat) ;
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
      // VL 301003 FQ 12958
      if M.ExportCFONB then
        OBM.PutMVT('E_CFONBOK', '#');
      // FIN VL

      //VT 06/02/2003 if M.SorteLettre<>tslCheque then OBM.PutMvt('E_CFONBOK','X') ;
      TM:=TMOD(GS.Objects[SA_NumP,Lig]) ; if TM<>Nil then EcheLettreCheque(Lig,TM.MODR) ;
      AlimGuideVentil(Lig) ;
{$IFDEF CCMP}
      AlimTableTemporaire(OBM,'',M.smp,'','') ;
{$ENDIF}
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
{$IFDEF CCMP}
         if (M.FormuleRefCCMP.Ref2 <> '') then GS.Cells[SA_RefI,Lig]:= M.FormuleRefCCMP.Ref2; // Référence génération
         if (M.FormuleRefCCMP.Lib2 <> '') then GS.Cells[SA_Lib,Lig] := M.FormuleRefCCMP.Lib2; // Libellé génération
{$ENDIF}
         ValideCeQuePeux(Lig) ;
         AlimObjetMvt(Lig,False) ;
         RemplirEcheAuto(Lig) ;
         RemplirAnalAuto(Lig) ;
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
{$IFDEF CCMP}
   if (M.FormuleRefCCMP.Ref2 <> '') then GS.Cells[SA_RefI,Lig]:= M.FormuleRefCCMP.Ref2; // Référence génération
   if (M.FormuleRefCCMP.Lib2 <> '') then GS.Cells[SA_Lib,Lig] := M.FormuleRefCCMP.Lib2; // Libellé génération
{$ENDIF}
   ValideCeQuePeux(Lig) ;
   AlimObjetMvt(Lig,False) ;
   RemplirEcheAuto(Lig) ;
   RemplirAnalAuto(Lig) ;
   OBM.PutMvt('E_MODEPAIE',LastPaie) ;
   OBM.PutMvt('E_DATEECHEANCE',OBM.GetMvt('E_DATECOMPTABLE')) ;
   TM:=TMOD(GS.Objects[SA_NumP,Lig]) ; if TM<>Nil then EcheLettreCheque(Lig,TM.MODR) ;
   END ;
{Ajustements}
GereNewLigne(GS) ;
CalculDebitCredit ;
GereOptionsGrid(GS.Row) ;
{Supprimer les enregistrements temporaires}
ExecuteSQL('DELETE from DOCREGLE Where DR_USER="'+V_PGI.USER+'"') ;
ExecuteSQL('DELETE from DOCFACT Where DF_USER="'+V_PGI.USER+'"') ;
FiniMove ;
{$IFDEF CCMP}
If SuiviMP Then AffecteRef(TRUE) ;
{$ENDIF}
Result:=True ;
END ;

procedure TFeSaisie.ClickCheque ;
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
  LL:=TList.Create ;
  Err:=0 ;
  CurMP:='' ;
  Premier:=True ;
  DevDocRegl:=OnLocale ;

  for i:=1 to GS.RowCount-1 do
    BEGIN
    O:=GetO(GS,i) ;
    if O=Nil then Break ;
    If i=1 Then
//      BEGIN
      If O.GetMvt('E_DEVISE')<>V_PGI.DevisePivot Then DevDocRegl:=OnDevise;
      {
      Else
        BEGIN
        If VH^.TenueEuro Then
          BEGIN
//          If O.GetMvt('E_SAISIEEURO')='X' Then DevDocRegl:=OnLocale Else DevDocRegl:=OnEuro ;
          If O.GetMvt('E_SAISIEEURO')='X' Then DevDocRegl:=OnEuro Else DevDocRegl:=OnLocale ;
          END Else
          BEGIN
          If O.GetMvt('E_SAISIEEURO')='X' Then DevDocRegl:=OnEuro Else DevDocRegl:=OnLocale ;
          END ;
        END ;
      END ;}
    MM:=TMOD(GS.Objects[SA_NumP,i]) ;
    if MM=Nil
      then Continue
      else if MM.ModR.NbEche>1 then
        BEGIN
        Err:=9 ;
        Break ;
        END ;
    If not M.TIDTIC Then
      if O.GetMvt('E_AUXILIAIRE')='' then Continue ;
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
    if ((SuiviMP) and (M.MPGUnique<>''))
      then CurMP:=M.MPGUnique
      else
        BEGIN
        if Premier
          then CurMP:=MM.ModR.TabEche[1].ModePaie
          else if MM.ModR.TabEche[1].ModePaie<>CurMP then
                 BEGIN
                 Err:=10 ;
                 Break ;
                 END ;
        END ;
    {Préparation OBM échéance}
    O.PutMvt('E_DATEECHEANCE',  MM.ModR.TabEche[1].DateEche ) ;
    O.PutMvt('E_MODEPAIE',      CurMP ) ;
    if (M.SorteLettre=tslCheque) Or (M.SorteLettre=tslVir) Or (M.SorteLettre=tslPre)
      then O.PutMvt('E_CONTREPARTIEGEN',M.General) ;
    LL.Add(O) ;
    Premier:=False ;
    END ;

  if Err<=0 then
    BEGIN
    OkPrintDialog:=TRUE ;
    If M.MSED.MultiSessionEncaDeca And M.MSED.SessionFaite
      Then OkPrintDialog:=FALSE ;
    If M.MSED.MultiSessionEncaDeca
      Then BAbandon.Enabled:=FALSE ;
    Err:=12 ;
    If M.TIDTIC
      Then NbC:=LanceDocReglTID(LL, M.SorteLettre, '', M.GroupeEncadeca, SuiviMP,
                                M.MSED, OkPrintDialog, DevDocRegl, M.smp, TOBParamEsc)
      Else NbC:=LanceDocRegl(   LL, M.SorteLettre, '', M.GroupeEncadeca, SuiviMP,
                                M.MSED, OkPrintDialog, DevDocRegl, M.smp, TOBParamEsc) ;
   // Attention bloc modifié 08/2003 suite pb maj E_NUMTRAITECHQ  // FICHE 12017
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
    if Err<>123456
      then HDiv.Execute(Err,'','') ;
    CEstGood:=False ;
    PieceModifiee:=False ;
    CloseFen ;
    END ;
  {$ENDIF}
{$ENDIF}

END ;

procedure TFeSaisie.ClickCherche ;
BEGIN
if GS.RowCount<=2 then Exit ;
FindFirst:=True ; FindSais.Execute ;
END ;

procedure TFeSaisie.ClickLibelleAuto ;
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
     StL:=Choisir(HDiv.Mess[1],'REFAUTO','RA_LIBELLE ||"(" || RA_JOURNAL || " " || RA_NATUREPIECE || ")"','RA_CODE','','')
   else
     StL:=Choisir(HDiv.Mess[1],'REFAUTO','RA_LIBELLE ||"(" || RA_JOURNAL || " " || RA_NATUREPIECE || ")"','RA_CODE',SWhere2,'') ;
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
Modifié le ... :   /  /    
Description .. : - LG - 18/03/2003 - blocage du rafraichissement de la grille
Mots clefs ... : 
*****************************************************************}
procedure TFeSaisie.ClickGuide ;
var StG         : String3 ;
begin
{$IFNDEF GCGC}
GS.BeginUpdate ;
try
  if ModeLC then Exit ;
  if PasModif then Exit ;
  if ((M.ANouveau) or (Not BGuide.Enabled)) then Exit ;
  if Guide then
    begin
    FinGuide ;
    Exit ;
    end ;
  if GS.RowCount>2 then Exit ;
  if EstRempli(GS,1) then Exit ;
  if ((EstRempli(GS,GS.Row)) and (Not LigneCorrecte(GS.Row,False,True))) then Exit ;
  StG := SelectGuide( E_JOURNAL.Value, E_NATUREPIECE.Value, E_DEVISE.Value,
                      E_ETABLISSEMENT.Value, (V_PGI.DateEntree>=DatePiece) ) ;
  if StG='' then Exit ;
  LanceGuide('NOR', StG) ;
finally
GS.EndUpdate ;
end;
  {$ENDIF}
END ;

procedure TFeSaisie.ClickProrata ;
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

procedure TFeSaisie.ClickDernPieces ;
BEGIN
  {$IFNDEF GCGC}
  VisuDernieresPieces(TPIECE) ;
  {$ENDIF}
END ;

procedure TFeSaisie.BCValideClick(Sender: TObject);
var Q     : TQuery ;
    i     : integer ;
    ValSt,CodeG : String ;
begin
  if FNomGuide.Text='' then
    begin
    HPiece.Execute(21,caption,'') ;
    Exit ;
    end ;
  if HPiece.Execute(22,caption,'')<>mrYes then
    begin
    BCAbandonClick(Nil) ;
    Exit ;
    end ;
  Q := OpenSQL('Select GU_GUIDE from GUIDE Where GU_TYPE="NOR" AND UPPER(GU_LIBELLE)="'+uppercase(FNomGuide.Text)+'"',True) ;
  if not Q.EOF then
    begin
    if HPiece.Execute(23,caption,'')<>mrYes then
      begin
      Ferme(Q) ;
      Exit ;
      end ;
    CodeG := Q.Fields[0].AsString ;
    EXECUTESQL('DELETE FROM GUIDE WHERE GU_TYPE="NOR" AND GU_GUIDE="'  + CodeG + '"') ;
    EXECUTESQL('DELETE FROM ECRGUI WHERE EG_TYPE="NOR" AND EG_GUIDE="' + CodeG + '"') ;
    EXECUTESQL('DELETE FROM ANAGUI WHERE AG_TYPE="NOR" AND AG_GUIDE="' + CodeG + '"') ;
    end ;
  Ferme(Q) ;
  Q := OpenSQL('Select MAX(GU_GUIDE) from GUIDE Where GU_TYPE="NOR"',True) ;
  if not Q.EOF then
    begin
    ValSt := Q.Fields[0].AsString ;
    if ValSt <> '' then
      begin
      i     := StrToInt( ValSt ) ;
      ValSt := IntToStr( i + 1 ) ;
      end
    else
      ValSt := '001' ;
    while Length(ValSt) < 3 do
      ValSt := '0' + ValSt ;
    end
  else
    ValSt := '001' ;

  Ferme(Q) ;
  Application.ProcessMessages ;
  BeginTrans ;
  CreerEnteteGuide(ValSt) ;	
  CreerLignesGuide(ValSt) ;	
  CommitTrans ;
  AvertirTable('ttGuideEcr') ;
  BCAbandonClick(Nil) ;
end;

procedure TFeSaisie.BCAbandonClick(Sender: TObject);
begin
PEntete.Enabled:=True ; Outils.Enabled:=True ; Valide97.Enabled:=True ;
GS.Enabled:=True ; PPied.Enabled:=True ;
GS.SetFocus ; PCreerGuide.Visible:=False ;
ModeCG:=False ; FNomGuide.Text:='' ;
end;

procedure TFeSaisie.ClickCreerGuide ;
var NatJ : String3 ;
begin
  if ModeLC then Exit ;
  if Not GS.Enabled then Exit ;
  if GS.RowCount<=2 then Exit ;
  if Not ExJaiLeDroitConcept(TConcept(ccSaisCreatGuide),True) then Exit ;
  if ((M.Anouveau) or (Not BCreerGuide.Enabled)) then Exit ;
  if ((EstRempli(GS,GS.Row)) and (Not LigneCorrecte(GS.Row,False,True))) then Exit ;
  if SAJAL=Nil then Exit ;

  NatJ:=SAJAL.NatureJal ;
// Modif SBO : On autorise la création pour les journaux de caisse
  if ((NATJ='ANO') or (NATJ='CLO') or (NATJ='ODA') or (NATJ='ANA') or (NATJ='ECC')) then
    begin
    HPiece.Execute(29,caption,'') ;
    Exit ;
    end ;
  PCreerGuide.Left    := GS.Left + (GS.Width-PCreerGuide.Width) div 2 ;
  PCreerGuide.Top     := GS.Top  + (GS.Height-PCreerGuide.Height) div 2 ;
  PCreerGuide.Visible :=True ;
  FNomGuide.SetFocus ;
  PEntete.Enabled   := False ;
  Outils.Enabled    := False ;
  Valide97.Enabled  := False ;
  GS.Enabled        := False ;
  PPied.Enabled     := False ;
  ModeCG            := True ;
END ;

procedure TFeSaisie.ClickVidePiece ( Parle : boolean ) ;
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

procedure TFeSaisie.ClickChancel ;
Var DD : TDateTime ;
    AA : TActionFiche ;
BEGIN
if ModeLC then Exit ;
if Action<>taCreat then Exit ;
if DEV.Code=V_PGI.DevisePivot then Exit ;
if ModeForce<>tmNormal then Exit ;
if Not BChancel.Enabled then Exit ;
if EstMonnaieIN(DEV.Code) then Exit ;
DD:=StrToDate(E_DATECOMPTABLE.Text) ;
AA:=taModif ; if Action=taConsult then AA:=taConsult ;
FicheChancel(DEV.Code,True,DD,AA,(DD>=V_PGI.DateDebutEuro)) ;	// EAGL : EN ATTENTE
if AA<>taConsult then DEV.Taux:=GetTaux(DEV.Code,DEV.DateTaux,DD) ;
END ;

procedure TFeSaisie.TabTvaDansModR ( OMODR : TMOD ) ;
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

procedure TFeSaisie.ClickModifTva ;
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

procedure TFeSaisie.ClickScenario ;
{$IFNDEF GCGC}
Var Jal,Nat : String3 ;
{$ENDIF}
BEGIN
  {$IFNDEF GCGC}
  Jal:='' ;
  Nat:='' ;
  if OkScenario then
    BEGIN
    Jal:=E_JOURNAL.Value ;
    Nat:=E_NATUREPIECE.Value ;
    END ;
  ParamScenario(Jal,Nat) ;
  {$ENDIF}
END ;

procedure TFeSaisie.ClickJournal ;
Var a : TActionFiche ;
begin
  {$IFNDEF GCGC}
  if Action = taConsult
    then a := taConsult
    else a := taModif ;
  if Not ExJaiLeDroitConcept(TConcept(ccJalModif),False) then a:=taConsult ;
  if ((E_JOURNAL.Value='') or (SAJAL=Nil)) then Exit ;
  if ModeLC then a:=taConsult ;
  FicheJournal(Nil,'',E_JOURNAL.Value,a,0) ;
  if a<>taConsult then
    ChercheJal(E_JOURNAL.Value,True) ;
  {$ENDIF}
end;

procedure TFeSaisie.ClickPieceGC ;
Var RefGC : String ;
    O     : TOBM ; 
begin
if Not BPieceGC.Enabled then Exit ;
O:=GetO(GS,GS.Row) ; if O=Nil then Exit ;
RefGC:=O.GetMvt('E_REFGESCOM') ; if RefGC='' then Exit ;
LanceZoomPieceGC(RefGC) ;
end;

procedure TFeSaisie.ClickImmo ;
{$IFDEF EAGLCLIENT}
{$ELSE}
Var O : TOBM ;
    CodeI : String ;
{$ENDIF}
begin
{$IFDEF EAGLCLIENT}
  PGIInfo('Non implémenté en mode CWAS', Caption ) ;
{$ELSE}
  {$IFNDEF CCMP}
    {$IFNDEF GCGC}
      {$IFNDEF IMP}
if Not BZoomImmo.Enabled then Exit ;
O:=GetO(GS,GS.Row) ; CodeI:=O.GetMvt('E_IMMO') ; if CodeI='' then Exit ;
if Not Presence('IMMO','I_IMMO',CodeI) then HDiv.Execute(20,Caption,'')
                                       else FicheImmobilisation(Nil,CodeI,taConsult,'') ;
      {$ENDIF}
    {$ENDIF}
  {$ENDIF}
{$ENDIF}
end;

procedure TFeSaisie.ClickTP ;
Var O : TOBM ;
    PieceTP : String ;
begin
  if Not BZoomTP.Enabled then Exit ;
  {$IFNDEF CCS3}
  O := GetO(GS,GS.Row) ;
  PieceTP := O.GetMvt('E_PIECETP') ;
  if PieceTP='' then Exit ;
  ZoomPieceTP(PieceTP) ;
  {$ENDIF}
end;

procedure TFeSaisie.BModifSerieClick(Sender: TObject);
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

procedure TFeSaisie.ClickDevise ;
BEGIN
	FicheDevise(E_DEVISE.Value,taConsult,False) ;
END ;

procedure TFeSaisie.ClickSaisTaux ;
BEGIN
if Not BSaisTaux.Enabled then Exit ;
if DEV.Code=V_PGI.DevisePivot then Exit ;
if DEV.Code='' then Exit ;
if Action<>taCreat then Exit ;
if ModeLC then Exit ;
if SaisieNewTaux2000(DEV,DatePiece) then BEGIN Volatile:=True ; M.TauxD:=DEV.Taux ; END ;
END ;

procedure TFeSaisie.ClickChoixRegime ;
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

procedure TFeSaisie.ClickEtabl ;
BEGIN
FicheEtablissement_AGL(taConsult) ;
END ;

procedure TFeSaisie.BVentilClick(Sender: TObject);
begin ClickVentil ; if GS.Enabled then GS.SetFocus ; end;

procedure TFeSaisie.BEcheClick(Sender: TObject);
begin ClickEche ; if GS.Enabled then GS.SetFocus ; end;

procedure TFeSaisie.BSDelClick(Sender: TObject);
begin ClickDel ; if GS.Enabled then GS.SetFocus ; end;

procedure TFeSaisie.BVidePieceClick(Sender: TObject);
begin ClickVidePiece(True) ; end;

procedure TFeSaisie.BInsertClick(Sender: TObject);
begin ClickInsert ; if GS.Enabled then GS.SetFocus ; end;

procedure TFeSaisie.BValideClick(Sender: TObject);
begin
if ModeForce<>tmNormal then Exit ;
if ((M.TypeGuide<>'NOR') or (Not M.FromGuide)) then ClickValide else ClickAbandon ;
end;

procedure TFeSaisie.BZoomDeviseClick(Sender: TObject);
begin ClickDevise ; end;

procedure TFeSaisie.BModifTvaClick(Sender: TObject);
begin ClickModifTva ; end;

procedure TFeSaisie.BSaisTauxClick(Sender: TObject);
begin ClickSaisTaux ; end;

procedure TFeSaisie.BChoixRegimeClick(Sender: TObject);
begin ClickChoixRegime ; end;

procedure TFeSaisie.BZoomEtablClick(Sender: TObject);
begin ClickEtabl ; end;

procedure TFeSaisie.BSoldeClick(Sender: TObject) ;
BEGIN ClickSolde(False,False) ; if GS.Enabled then GS.SetFocus ; end;

procedure TFeSaisie.BAbandonClick(Sender: TObject);
begin ClickAbandon ; end;

procedure TFeSaisie.BZoomClick(Sender: TObject);
begin
ClickZoom ; if GS.Enabled then GS.SetFocus ;
end;

procedure TFeSaisie.BScenarioClick(Sender: TObject);
begin ClickScenario ; end;

procedure TFeSaisie.BZoomJournalClick(Sender: TObject);
begin ClickJournal ; end;

procedure TFeSaisie.BPieceGCClick(Sender: TObject);
begin ClickPieceGC ; end;

procedure TFeSaisie.BZoomImmoClick(Sender: TObject);
begin ClickImmo ; end;

procedure TFeSaisie.BZoomTPClick(Sender: TObject);
begin ClickTP ; end;

procedure TFeSaisie.BGenereTvaClick(Sender: TObject);
begin ClickGenereTVA ; if GS.Enabled then GS.SetFocus ; end;

procedure TFeSaisie.BControleTVAClick(Sender: TObject);
begin ClickControleTVA(False,0) ; if GS.Enabled then GS.SetFocus ; end;

procedure TFeSaisie.BProrataClick(Sender: TObject);
begin ClickProrata ; if GS.Enabled then GS.SetFocus ; end;

procedure TFeSaisie.BSwapPivotClick(Sender: TObject);
begin ClickSwapPivot ; end;

procedure TFeSaisie.BModifRIBClick(Sender: TObject);
begin ClickModifRIB ; end;

{
procedure TFeSaisie.BSwapEuroClick(Sender: TObject);
begin ClickSwapEuro ; end;
}

procedure TFeSaisie.BComplementClick(Sender: TObject);
begin ClickComplement ; end;

procedure TFeSaisie.BChercherClick(Sender: TObject);
begin ClickCherche ; end;

procedure TFeSaisie.BChequeClick(Sender: TObject);
begin ClickCheque ; end;

procedure TFeSaisie.BGuideClick(Sender: TObject);
begin ClickGuide ; end;

procedure TFeSaisie.BLibAutoClick(Sender: TObject);
begin ClickLibelleAuto ; if GS.Enabled then GS.SetFocus ; end;

procedure TFeSaisie.BChancelClick(Sender: TObject);
begin ClickChancel ; end;

procedure TFeSaisie.BDernPiecesClick(Sender: TObject);
begin ClickDernPieces ; if GS.Enabled then GS.SetFocus ; end;

procedure TFeSaisie.BCreerGuideClick(Sender: TObject);
begin ClickCreerGuide ; if GS.Enabled then GS.SetFocus ; end;

{==========================================================================================}
{============================ Création Guide depuis saisie ================================}
{==========================================================================================}
function TFeSaisie.BonRegle ( Lig,NbEche : integer ; LeMode : String ; T : T_TabEche ) : Boolean ;
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

procedure TFeSaisie.CreerLignesGuide ( CodeG : String ) ;
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


procedure TFeSaisie.CreerEnteteGuide ( CodeG : String ) ;
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
//TGuide.FindField('GU_SAISIEEURO').AsString:=CheckToString(ModeOppose) ;
  TGuide.PutValue('GU_TRESORERIE', '-') ;
  TGuide.InsertDB(nil) ;
  TGuide.Free ;
END ;


{==========================================================================================}
{================================ Guide de saisie =========================================}
{==========================================================================================}
procedure TFeSaisie.EtudieSiGuide(ACol,ARow : longint);
Var St : String ;
BEGIN
if PasModif then Exit ; if M.ANouveau then Exit ;
St:=GS.Cells[ACol,ARow] ; if St='' then Exit ; if St[1]<>'*' then Exit ; if GS.Objects[ACol,ARow]<>Nil then Exit ;
System.Delete(St,1,1) ; if St<>'' then LanceGuide('NOR',St) ;
END ;

procedure TFeSaisie.AlimGuideTVA ( Lig : integer ; Q : TQuery ) ;
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

procedure TFeSaisie.AlimGuideVentil ( Lig : integer ) ;
var CGen : TGGeneral ;
begin
  if Action<>taCreat then Exit ;
  if ((M.TypeGuide<>'ENC') and (M.TypeGuide<>'DEC')) then Exit ;
  CGen := GetGGeneral(GS,Lig) ;
  if CGen = Nil then Exit ;
  if Ventilable(CGen,0)
    then RemplirAnalAuto(Lig) ;
end ;

procedure TFeSaisie.AlimGuideLettrage ( Lig : integer ; Q : TQuery ) ;
Var O  : TOBM ;
    T  : TStrings ;
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
{Mémoriser les échéances d'origine}
T:=TStringList.Create ; T.Assign(O.LC) ; TECHEORIG.Add(T) ;
END ;

{$IFDEF CEGID}
procedure TFeSaisie.ResaisirLaDatePourLeGuide(OkOk : Boolean) ;
Var DD : TDateTime ;
    OBA   : TOB ;
    T     : TList ;
    NumV,i,NumAxe : Integer ;
    OBM : TOBM ;
BEGIN
{$IFNDEF GCGC}
{$IFNDEF IMP}
  If Not MethodeGuideDate1 Then
    begin
    if OkOk then
      begin
      DD := V_PGI.DateEntree ;
     SaisirDateGuide(DD) ;
     E_DATECOMPTABLE.Text:=DateToStr(DD) ;
     for i:=1 to GS.RowCount-2 do
       begin
       OBM := GetO(GS,i) ;
       if OBM<>NIL
        then OBM.PutMvt('E_DATECOMPTABLE',DD) ;

       for NumAxe:=1 to Min(MaxAxe, OBM.Detail.Count) do
        if OBM.Detail[NumAxe-1].Detail.Count > 0 then
          begin
          for NumV:=0 to T.Count-1 do
            begin
            OBA := OBM.Detail[NumAxe-1].Detail[NumV] ;
          	OBA.PutValue('Y_DATECOMPTABLE', DD) ;
            end ;
          end ;
       end ;
    end ;
  end ;
{$ENDIF}
{$ENDIF}
END ;
{$ENDIF}

procedure TFeSaisie.LanceGuide ( TypeG,CodeG : String3 ) ;
Var OG : TOBM ;
    Premier : boolean ;
    St        : String ;
    Col,Lig   : integer ;
    Q,QE      : TQuery ;
BEGIN
if M.ANouveau then Exit ;
{$IFDEF CEGID}
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
{$ENDIF}
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
            {
            if ((E_DEVISE.Value=V_PGI.DevisePivot) and (QE.FindField('GU_SAISIEEURO').AsString='X')) then
               BEGIN
               E_DEVISE.ItemIndex:=E_DEVISE.Values.Count-1 ; E_DEVISEChange(Nil) ;
               END ;
            }
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
   ValideCeQuePeux(Lig) ;
   AlimObjetMvt(Lig,True) ;
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
   GS.SetFocus ; FlashGuide.Visible:=True ; FlashGuide.Flashing:=True ;
   if Col>0 then BEGIN GS.Col:=Col ; END ;
   END ;
if ((Action=taCreat) and ((M.TypeGuide='ENC') or (M.TypeGuide='DEC'))) then
   BEGIN
   BlocageEntete(Self,False) ;
   if TypeG=V_PGI.User then BEGIN LeTimer.Enabled:=True ; ModeLC:=True ; GridEna(GS,False) ; END ;
   END ;
{$IFDEF CEGID}
ResaisirLaDatePourLeGuide((Not M.FromGuide) And (TypeG='NOR') And (M.TypeGuide='')) ; // a partir de Alt G
{$ENDIF}
END ;

procedure TFeSaisie.RecalcGuide ;
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
       if ((EstFormule(St)) or (GS.Cells[SA_Lib,Lig]='')) then
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

function TFeSaisie.Load_TVATPF ( CGen : TGGeneral ; OKTVA : boolean ; TVA : String ) : Variant ;
BEGIN
Result:=0 ; if CGen=Nil then Exit ;
AttribRegimeEtTva ;
if OKTVA then
   BEGIN
   if TVA='' then TVA:=CGen.TVA ;
   Result:=Tva2Taux(GeneRegTVA,TVA,(SAJAL.NatureJAL='ACH')) ;
   END else Result:=Tpf2Taux(GeneRegTVA,CGen.TPF,(SAJAL.NatureJAL='ACH')) ;
END ;

function TFeSaisie.Load_Sais ( St : String ) : Variant ;
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

{$IFDEF CEGID}
Function TrouveNextLigneAZero(GS : thGrid) : Integer ;
Var i : Integer ;
BEGIN
Result:=0 ;
for i:=1 to GS.RowCount-2 do
  BEGIN
  If ((ValD(GS,i)=0) and (ValC(GS,i)=0)) then
    BEGIN
    if GS.RowCount>3 then BEGIN Result:=i ; Break ; END ;
    END ;
  END ;
END ;

procedure TFeSaisie.TraiteLigneAZero ;
Var LigneAZero : Boolean ;
    LigNext : Integer ;
    i : Integer ;
BEGIN
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
{$ENDIF}

procedure TFeSaisie.FinGuide ;
Var Cancel : boolean ;
    i : integer ;
BEGIN
{$IFDEF CEGID}
{$ENDIF}
if Guide then CalculDebitCredit ;
for i:=1 to GS.RowCount-2 do
    BEGIN
    Cancel:=False ;
    GSRowExit(Nil,i,Cancel,True) ;
    END ;
Guide:=False ; FlashGuide.Flashing:=False ; FlashGuide.Visible:=False ;
END ;

Function TFeSaisie.CommeUneBQE ( Lig : integer ) : Boolean ;
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

function TFeSaisie.ZoneObli ( Col,Lig : integer ; St : String ) : Boolean ;
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
{$IFDEF CEGID}
{$ELSE}
   if ((TrouveSens(GS,E_NATUREPIECE.Value,Lig,CommeUneBQE(Lig))=1) and (GS.Cells[SA_Credit,Lig]='') and
       (GS.Cells[SA_Debit,Lig]='') and (Not ArretZone(SA_Credit,St,False))) then Exit ;
{$ENDIF}
   END ;
if Col=SA_Credit then
   BEGIN
{$IFDEF CEGID}
{$ELSE}
   if ((TrouveSens(GS,E_NATUREPIECE.Value,Lig,CommeUneBQE(Lig))=2) and (GS.Cells[SA_Debit,Lig]='') and
       (GS.Cells[SA_Credit,Lig]='') and (Not ArretZone(SA_Debit,St,False))) then Exit ;
{$ENDIF}
   END ;
ZoneObli:=False ;
END ;

function TFeSaisie.ProchainArret ( Col,Lig : integer ) : integer ;
Var OBM : TOBM ;
    St  : String ;
BEGIN
Result:=-1 ; if Not Guide then Exit ; if Lig>TGUI.Count then Exit ;
OBM:=TOBM(TGUI[Lig-1]) ; if OBM=Nil then BEGIN FinGuide ; Result:=-1 ; Exit ; END ;
St:=OBM.GetMvt('EG_ARRET') ;
for Col:=SA_Gen to SA_Credit do
    BEGIN
    if ArretZone(Col,St,True) then BEGIN OBM.PutMvt('EG_ARRET',St) ; Result:=Col ; Exit ; END ;
    if ZoneObli(Col,Lig,St) then BEGIN Result:=Col ; Exit ; END ;
    END ;
END ;

procedure TFeSaisie.ValideCeQuePeux( Lig : longint ) ;
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
OBM.PutMvt('E_REFINTERNE',Copy(GS.Cells[SA_RefI,Lig],1,35)) ;
OBM.PutMvt('E_LIBELLE',Copy(GS.Cells[SA_Lib,Lig],1,35)) ;
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
procedure TFeSaisie.SuiteLigneColonne ( Var Col,Lig : longint ) ;
Var CN: integer ;

  procedure _SuiteLigneColonne ( Var Col,Lig : longint ) ;
   begin
    OkSuite:=False ;
    if Not Guide then Exit ; if Lig>NbLG then BEGIN FinGuide ; Exit ; END ;
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
Function TFeSaisie.FabricFromM : RMVT ;
Var X : RMVT ;
BEGIN
X.Axe:='' ; X.Etabl:=E_ETABLISSEMENT.Value ;
X.Jal:=E_JOURNAL.Value ; X.Exo:=QuelExoDT(StrToDate(E_DATECOMPTABLE.Text)) ;
X.CodeD:=E_DEVISE.Value ; X.Simul:=M.Simul ; X.Nature:=E_NATUREPIECE.Value ;
X.DateC:=StrToDate(E_DATECOMPTABLE.Text) ; X.DateTaux:=DEV.DateTaux ;
X.Num:=NumPieceInt ; X.TauxD:=DEV.Taux ; X.Valide:=False ; X.ANouveau:=M.ANouveau ;
Result:=X ;
END ;

procedure TFeSaisie.StockeLaPiece ;
Var X : RMVT ;
    P : P_MV ;
BEGIN
if ((Action<>taCreat) and (Not SuiviMP)) then Exit ; FillChar(X,Sizeof(X),#0) ;
X:=FabricFromM ; P:=P_MV.Create ; P.R:=X ; TPIECE.Add(P) ;
END ;

procedure TFeSaisie.NumeroteVentils ;
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

procedure TFeSaisie.GotoEntete ;
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

procedure TFeSaisie.GereComplements(Lig : integer) ;
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

procedure TFeSaisie.DecrementeNumero ;
Var Facturier : String3 ;
    DD : TDateTime ;
BEGIN
if M.Simul<>'N' then Facturier:=SAJAL.CompteurSimul else Facturier:=SAJAL.CompteurNormal ;
DD:=StrToDate(E_DATECOMPTABLE.Text) ;
SetDecNum(EcrGen,Facturier,DD) ;
END ;

Procedure TFeSaisie.EditionSaisie ;
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

Procedure TFeSaisie.DetruitPieceMP ;
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

Function TFeSaisie.FermerSaisie : boolean  ;
Var i : integer ;
    Okok : boolean ;
BEGIN
Result:=True ;
if ((Action<>taConsult) and (ModeForce<>tmNormal)) then BEGIN Result:=False ; Exit ; END ;
if ((NeedEdition) and (Action=taCreat) and (TPIECE.Count>0) and
    (Not M.MajDirecte) And (Not M.MSED.MultiSessionEncaDeca)) then EditionSaisie ;
if ((Action=taConsult) or (Not PieceModifiee) or (M.smp=smpEncTraEdtNC) Or (M.smp=smpDecBorEdtNC) Or (M.smp=smpDecChqEdtNC) Or (M.smp=smpDecVirEdtNC) Or (M.smp=smpDecVirInEdtNC) Or (M.smp=smpEncPreEdtNC) ) then
   BEGIN
   if SuiviMP then DetruitPieceMP ;
   Exit ;
   END ;
if ((Action=taCreat) and (M.TypeGuide='NOR') and (M.FromGuide)) then Exit ;
if Action=taModif then
   BEGIN
   if HPiece.Execute(3,caption,'')<>mrYes then Result:=False else if SuiviMP then DetruitPieceMP ;
   END else
   BEGIN
   if PossibleRecupNum then
      BEGIN
      Okok:=True ; i:=HPiece.Execute(3,caption,'') ;
      if i<>mrYes then BEGIN Result:=False ; Exit ; END ;
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

procedure TFeSaisie.AlimObjetMvt ( Lig : integer ; AvecM : boolean ) ;
Var OBM   : TOBM ;
    XD,XC,PD,PC,SD,SC,ED,EC : Double ;
    CGen  : TGGeneral ;
    LaDate      : TDateTime ;
    St{,ME} : String ;
BEGIN
if PasModif then Exit ;// ME:=CheckToString(ModeOppose) ;
OBM:=GetO(GS,Lig) ;
if OBM=Nil then BEGIN AlloueMvt(GS,EcrGen,Lig,True) ; OBM:=GetO(GS,Lig) ; END ;
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
OBM.PutMvt('E_CONTROLETVA','RIE')               ; //OBM.PutMvt('E_SAISIEEURO',ME) ;
OBM.PutMvt('E_LIBELLE',Copy(GS.Cells[SA_Lib,Lig],1,35)) ;
OBM.PutMvt('E_REFINTERNE',Copy(GS.Cells[SA_RefI,Lig],1,35)) ;
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
procedure TFeSaisie.RemplirEcheAuto ( Lig : integer ) ;
BEGIN
GereEche(Lig,False,True) ;
END ;

procedure TFeSaisie.RemplirAnalAuto ( Lig : integer ) ;
Var j,NumA,NumL,NumV 	: integer ;
    CGen 							: TGGeneral ;
    Ax   							: String ;
    DD,CD,DP,CP{,DE,CE} : Double ;
    OBM, OBA 	        : TOBM ;
BEGIN
	CGen:=GetGGeneral(GS,Lig) ;
	if DEV.Code<>V_PGI.DevisePivot then
    BEGIN
    DD:=ValD(GS,Lig) ; CD:=ValC(GS,Lig) ;
    if VH^.TenueEuro then
      BEGIN
      DP:=DeviseToEuro(DD,DEV.Taux,DEV.Quotite)  ; CP:=DeviseToEuro(CD,DEV.Taux,DEV.Quotite) ;
      //DE:=DeviseToPivot(DD,DEV.Taux,DEV.Quotite) ; CE:=DeviseToPivot(CD,DEV.Taux,DEV.Quotite) ;
      END
		else
      BEGIN
      DP:=DeviseToPivot(DD,DEV.Taux,DEV.Quotite) ; CP:=DeviseToPivot(CD,DEV.Taux,DEV.Quotite) ;
      //DE:=DeviseToEuro(DD,DEV.Taux,DEV.Quotite)  ; CE:=DeviseToEuro(CD,DEV.Taux,DEV.Quotite) ;
      END ;
   END
  else
    {if ModeOppose then
    BEGIN
    DE:=ValD(GS,Lig) ; CE:=ValC(GS,Lig) ;
    if VH^.TenueEuro
	    then BEGIN DP:=PivotToEuro(DE) ; CP:=PivotToEuro(CE) ; END
  	  else BEGIN DP:=EuroToPivot(DE) ; CP:=EuroToPivot(CE) ; END ;
    DD:=DP ; CD:=CP ;
    END
	else}
    BEGIN
    DD:=ValD(GS,Lig) ; CD:=ValC(GS,Lig) ;
    {if VH^.TenueEuro
	    then BEGIN DE:=EuroToPivot(DD) ; CE:=EuroToPivot(CD) ; END
  	  else BEGIN DE:=PivotToEuro(DD) ; CE:=PivotToEuro(CD) ; END ;}
    DP:=DD ; CP:=CD ;
    END ;

  // Récupération de la TOBM écriture 
  OBM	:= GetO(GS,Lig) ;
  if OBM=Nil then Exit ;
  if OBM.Detail.Count<MaxAxe then AlloueAxe(OBM);
  NumL := Lig ;
  NumV := 1 ;

	// Pour chaque Axe, génération s'une seule écriture analytique sur compte d'attente
  for NumA:=1 to MaxAxe do
  	if Ventilable(CGen,NumA) then
      BEGIN
      // On efface ventilation existante sur axe
      if OBM.Detail[NumA-1].Detail.Count>0
      	then OBM.Detail[NumA-1].ClearDetail ;
			// Création objet TOBM analytique
      Ax:='A'+Chr(48+NumA) ;
      OBA:=TOBM.Create(EcrAna,Ax,True) ;
      // Affectation des données
      InitCommunObjAnalNew(OBM,OBA) ;     // EAGL : OBA remplit depuis OBM ou Interface ???
      OBA.PutValue('Y_SECTION',				VH^.Cpta[AxeToFb(Ax)].Attente) ;
      OBA.PutValue('Y_POURCENTAGE',		100.0) ;
      OBA.PutValue('Y_DEBIT',					DP)      ;
      OBA.PutValue('Y_CREDIT',				CP) ;
      OBA.PutValue('Y_DEBITDEV',			DD)   ;
      OBA.PutValue('Y_CREDITDEV',			CD) ;
//      OBA.PutValue('Y_DEBITEURO',			DE)  ;
  //    OBA.PutValue('Y_CREDITEURO',		CE) ;
      OBA.PutValue('Y_NUMLIGNE',			NumL) ;
      OBA.PutValue('Y_NUMVENTIL',			NumV) ;
      OBA.PutValue('Y_AXE',						Ax) ;
      OkMessAnal:=True ;
      END ;
END ;

function TFeSaisie.InsereLigneTVA ( LigHT : integer ; RegTVA : String3 ; SoumisTpf,Achat : boolean ) : boolean ;
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
      if Lettrable(TpfGen)<>NonEche
      	then RemplirEcheAuto(LigF)
	      else if Ventilable(TpfGen,0) then RemplirAnalAuto(LigF) ;
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
      if Lettrable(TvaGen)<>NonEche
      	then RemplirEcheAuto(LigA)
      	else if Ventilable(TvaGen,0) then RemplirAnalAuto(LigA) ;
      OBM:=GetO(GS,LigA) ; OBM.PutMvt('E_REGIMETVA',RegTVA) ; OBM.PutMvt('E_TVA',CodeTVA) ; OBM.PutMvt('E_TPF',CodeTPF) ;
      END ;
   END else if CGen.TauxTva<>0 then Result:=False ;
END ;

procedure TFeSaisie.AlimHTByTVA ( Lig : integer ; RegTVA,TVAENC : String3 ; SoumisTpf,Achat : boolean ) ;
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
(*ocedure TFeSaisie.GetCellCanvas ( ACol,ARow : Longint; Canvas : TCanvas ; AState: TGridDrawState) ;
Var St : String ;
BEGIN
St:=GS.Cells[SA_Gen,ARow] ; if St='' then Exit ;
if ((St=VH^.EccEuroDebit) or (St=VH^.EccEuroCredit)) then Canvas.Font.Color:=clGray ;
END ;*)

procedure TFeSaisie.GSDblClick(Sender: TObject);
begin
if ((Action<>taConsult) and (ModeForce=tmDevise)) then ForcerMode(False,VK_RETURN) else ClickZoom ;
end;

procedure TFeSaisie.GSCellExit(Sender: TObject; var ACol,ARow : Longint ; var Cancel : Boolean );
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
if ((ACol=SA_Debit) or (ACol=SA_Credit)) then ChercheMontant(Acol,ARow) ;
if Guide then RecalcGuide ;
if ((Not Cancel) and (GS.Cells[ACol,ARow]<>StCellCur)) then UnChange:=True ; 
end;

procedure TFeSaisie.GereOptionsGrid ( Lig : integer ) ;
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
         G_LIBELLE.Caption:=HDiv.Mess[7] + ' (' + GetO(GS,Lig).GetMvt('E_REFPOINTAGE') + ')' ; // Modif BPY 08/08/2003
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

procedure TFeSaisie.GSCellEnter(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
Var CGen : TGGeneral ;
begin
if ARow<>GS.Row then GereOptionsGrid(GS.Row) ;
if ((Action<>taConsult) and (Not EstRempli(GS,GS.Row)) and (GS.Col<>SA_Gen) and (GS.Col<>SA_Aux))
   then BEGIN ACol:=SA_Gen ; ARow:=GS.Row ; Cancel:=True ; Exit ; END ;
if ((H_NUMEROPIECE_.Visible) and (GS.Row>1)) then
   BEGIN
   if Transactions(AttribNumeroDef,10)<>oeOK then MessageAlerte(HDiv.Mess[5]) ;
   BlocageEntete(Self,FALSE) ;
   END ;
BValide.Enabled:=((H_NumeroPiece.Visible) and (Equilibre)) ;
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

procedure TFeSaisie.GSRowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
begin
if ((Not EstRempli(GS,Ou)) and (Not PasModif)) then DefautLigne(Ou,True) ;
CurLig:=Ou ; AffichePied ;
GereEnabled(Ou) ; GereOptionsGrid(Ou) ;
end;

procedure TFeSaisie.GSRowExit(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
Var OkL  : Boolean ;
begin
if PasModif then Exit ;
if Ou>=GS.RowCount-1 then Exit ;
GridEna(GS,False) ;
OkL:=LigneCorrecte(Ou,False,(Action=taCreat) or (Not Revision) or (UnChange)) ;
if OkL then
   BEGIN
   GereComplements(Ou) ;
   GereEche(Ou,True,False) ;
   GereAnal(Ou,True,True) ;
   GereTvaEncNonFact(Ou) ;
   GereLesImmos(Ou) ;
   END else Cancel:=True ;
GridEna(GS,True) ;
end;

procedure TFeSaisie.GSEnter(Sender: TObject);
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
{$IFDEF CEGID}
If MethodeGuideDate1 Then
  BEGIN
  if (appeldatepourguide And (Not M.FromGuide) And (((M.TypeGuide='NOR') and (M.SaisieGuidee)) Or (CodeGuideFromSaisie<>''))) then
    BEGIN
    If M.SaisieGuidee Then LanceGuide(M.TypeGuide,M.LeGuide) Else LanceGuide('NOR',CodeGuideFromSaisie) ;
    END ;
  END ;
{$ENDIF}
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
if ((Action=taCreat) and (Not M.FromGuide) and (Not EstRempli(GS,1)) and
    (OkD) and (Not DejaRentre) and (Not M.SaisieGuidee) and (Not RentreDate)) then
   BEGIN
   if E_DATECOMPTABLE.CanFocus then E_DATECOMPTABLE.SetFocus ; Exit ;
   END ;
{$ENDIF}
{$IFDEF CEGID}
ResaisirLaDatePourLeGuide((Not M.FromGuide) And (M.TypeGuide='NOR') And M.SaisieGuidee And (Not RentreDate)) ; // a prtie de saisie guidée
{$ENDIF}
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

procedure TFeSaisie.GSKeyPress(Sender: TObject; var Key: Char);
begin
if Not GS.SynEnabled then Key:=#0 ;
end;

procedure TFeSaisie.GSExit(Sender: TObject);
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

procedure TFeSaisie.GSMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin GX:=X ; GY:=Y ; end;

procedure TFeSaisie.POPSPopup(Sender: TObject);
begin InitPopUp(Self) ; end;

procedure TFeSaisie.FindSaisFind(Sender: TObject);
Var bc : boolean ;
begin
Rechercher(GS,FindSais,FindFirst) ;
bc:=False ; GSRowEnter(Nil,GS.Row,bc,False) ;
end;

procedure TFeSaisie.GSSetEditText(Sender: TObject; ACol, ARow: Longint; const Value: string);
begin PieceModifiee:=True ; end;

//=============================================================================

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
	Result := FALSE ;
	if (Q.EOF) and (Q.Bof) then Exit ; //n°1825
	if Not QOK(Q,Simul) then
  	begin
    Result:=FALSE ;
    Exit ;
    end ;
	if Simul='' then
	  Q1:=OpenSQL('Select * from Ecriture where E_JOURNAL="'+Q.FindField('E_JOURNAL').AsString+'"'
            +' AND E_EXERCICE="'+QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime))+'"'
            +' AND E_DATECOMPTABLE="'+USDATETIME(Q.FindField('E_DATECOMPTABLE').AsDateTime)+'"'
            +' AND E_NUMEROPIECE='+Q.FindField('E_NUMEROPIECE').AsString,True)
  else
  	Q1:=OpenSQL('Select * from Ecriture where E_JOURNAL="'+Q.FindField('E_JOURNAL').AsString+'"'
            +' AND E_EXERCICE="'+QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime))+'"'
            +' AND E_DATECOMPTABLE="'+USDATETIME(Q.FindField('E_DATECOMPTABLE').AsDateTime)+'"'
            +' AND E_NUMEROPIECE='+Q.FindField('E_NUMEROPIECE').AsString
            +' AND E_QUALIFPIECE="'+Simul+'" ',True) ;
	Trouv := Not Q1.EOF ;
  if Trouv then
  	M := MvtToIdent(Q1,fbGene,False) ;
  Ferme(Q1) ;
	Result := Trouv ;
END ;



Function TrouveEtLanceSaisie(Q : TQuery ; TypeAction : TActionFiche ; Simul : String3 ; ModLess : Boolean = FALSE) : Boolean ;
Var M : RMVT ;
begin
	Result := TrouveSaisie(Q,M,Simul) ;
	if Result then
  	BEGIN
    {$IFDEF EAGLCLIENT}
    {$ELSE}
			{$IFDEF RECUPSISCOPGI}
   	  if ((M.ModeSaisieJal<>'-') and (M.ModeSaisieJal<>''))
    	  then LanceSaisieFolio(Q,TypeAction) else
			{$ELSE}
				{$IFNDEF GCGC}
   	    if ((M.ModeSaisieJal<>'-') and (M.ModeSaisieJal<>''))
          then LanceSaisieFolio(Q,TypeAction) else
				{$ENDIF}
			{$ENDIF}
	  {$ENDIF}
		LanceSaisie(Q,TypeAction,M,ModLess) ;
  END ;
END ;


Function TFeSaisie.Recharge(QM : RMVT ) : Boolean ;
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

procedure TFeSaisie.BPrevClick(Sender: TObject);
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

procedure TFeSaisie.BNextClick(Sender: TObject);
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

procedure TFeSaisie.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFeSaisie.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFeSaisie.BMenuZoomMouseEnter(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
{$ELSE}
PopZoom97(BMenuZoom,POPZ) ;
{$ENDIF}
end;

procedure TFeSaisie.BValideMouseEnter(Sender: TObject);
begin
Valide97.Tag:=1 ;
end;

procedure TFeSaisie.BValideMouseExit(Sender: TObject);
begin
Valide97.Tag:=0 ;
end;

procedure TFeSaisie.GereTagEnter(Sender: TObject);
Var B : TToolbarButton97 ;
begin
B:=TToolbarButton97(Sender) ; if B=Nil then Exit ;
if B.Parent=Outils then Outils.Tag:=1 ;
end;

procedure TFeSaisie.GereTagExit(Sender: TObject);
Var B : TToolbarButton97 ;
begin
B:=TToolbarButton97(Sender) ; if B=Nil then Exit ;
if B.Parent=Outils then Outils.Tag:=0 ;
end;

procedure TFeSaisie.LeTimerTimer(Sender: TObject);
begin
LeTimer.Enabled:=False ; ModeLC:=False ; GridEna(GS,True) ;
if ((GenereAuto) and (SuiviMP) and (M.smp in [smpEncTous,smpDecTous])) then BValideClick(Nil) else
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

procedure TFeSaisie.PEnteteEnter(Sender: TObject);
begin
if ModeLC then
   BEGIN
   if GS.CanFocus then GS.SetFocus else if Valide97.CanFocus then Valide97.SetFocus ;
   END ;
end;

procedure TFeSaisie.BPopTvaChange(Sender: TObject);
Var O : TOBM ;
    i : integer ;
    CGen : TGGeneral ;
    Ste  : String ;
    OMODR : TMOD ;
begin
if GeneCharge then Exit ;
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

procedure TFeSaisie.E_DEVISEExit(Sender: TObject);
begin
if csDestroying in ComponentState then Exit ;
{
if ((ModeOppose) and (Action=taCreat) and (Not VH^.TenueEuro)) then if DatePiece<V_PGI.DateDebuteuro then
   BEGIN
   HPiece.Execute(39,Caption,'') ;
   E_DEVISE.Value:=V_PGI.DevisePivot ;
   if E_DEVISE.CanFocus then E_DEVISE.SetFocus ;
   END ;
   }
end;

procedure TFeSaisie.GereAffSolde(Sender: TObject);
Var Nam : String ;
    C   : THLabel ;
begin
Nam:=THNumEdit(Sender).Name ; Nam:='L'+Nam ;
C:=THLabel(FindComponent(Nam)) ;
if C<>Nil then C.Caption:=THNumEdit(Sender).Text ;
end;

{================================ Tiers payeur ==================================}
Procedure TFeSaisie.DetruitPiecesTP ;
Var i : integer ;
    PieceTP : String ;
BEGIN
{$IFNDEF GCGC}
  {$IFNDEF CCS3}
  for i:=0 to ListeTP.Count-1 do
    begin
    PieceTP := ListeTP[i] ;
    SupprimePieceTP(PieceTP) ;
    END ;
  {$ENDIF}
{$ENDIF}
end ;

Procedure TFeSaisie.DetruitOldPayeurs ;
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

Procedure TFeSaisie.GereTiersPayeur ;
Var X : RMVT ;
BEGIN
{$IFNDEF GCGC}
  {$IFNDEF CCS3}
  if YATP=yaNL
    then DetruitPiecesTP ;
  FillChar(X,Sizeof(X),#0) ;
  X.Axe       := '' ;
  X.Etabl     := E_ETABLISSEMENT.Value ;
  X.Jal       := E_JOURNAL.Value ;
  X.Exo       := QuelExoDT(StrToDate(E_DATECOMPTABLE.Text)) ;
  X.CodeD     := E_DEVISE.Value ;
  X.Simul     := M.Simul ;
  X.Nature    := E_NATUREPIECE.Value ;
  X.DateC     := StrToDate(E_DATECOMPTABLE.Text) ;
  X.DateTaux  := DEV.DateTaux ;
  X.Num       := NumPieceInt ;
  X.TauxD     := DEV.Taux ;
  X.Valide    := False ;
  X.ANouveau  := False ;
//X.ModeOppose:=ModeOppose ;
  GenerePiecesPayeur(X) ;
  {$ENDIF}
{$ENDIF}
END ;

Procedure TFeSaisie.YaBienTP ( QEcr : TDataSet ) ;
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
    begin
    if VH^.JalVTP='' then Exit ;
    TestJ:=RechDom('TTJOURNAUX',VH^.JalVTP,False) ;
    if ((TestJ='') or (TestJ='Error')) then Exit ;
    end ;
  if ((Nat='FF') or (Nat='AF')) then
    begin
    if VH^.JalATP='' then Exit ;
    TestJ:=RechDom('TTJOURNAUX',VH^.JalATP,False) ;
    if ((TestJ='') or (TestJ='Error')) then Exit ;
    END ;
  TiersP := QEcr.FindField('E_TIERSPAYEUR').AsString ;
  if TiersP='' then Exit ;
  PieceTP := QEcr.FindField('E_PIECETP').AsString ;
  if PieceTP='' then Exit ;
  ListeTP.Add(PieceTP) ;
  YATP:=yaNL ;
  if ExisteLettrageSurTP(TiersP,PieceTP) then YATP:=yaL ;
  {$ENDIF}
{$ENDIF}
END ;

Function TFeSaisie.ExisteTP : boolean ;
Var i : integer ;
    OKP,LaQuestion,AvoirRbt : boolean ;
    Nat : String ;
    GAux : TGTiers ;
BEGIN
Result:=False ; OkP:=False ; LaQuestion:=False ; AvoirRbt:=False ;
if Action=taConsult then Exit ;
if Not VH^.OuiTP then Exit ;
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

// Détruit la ligne d'un guide lors d'un CTRL+SUPP (Fiche 10809)
procedure TFeSaisie.DetruitLigneGuide(Lig: integer);
var
  iLigne,iNumLigne : integer;
  OG,OG1 : TOBM;
begin
  for iLigne:=NbLG downto Lig do
    begin
    OG:=TOBM(TGUI[iLigne-1]);
    OG1:=TOBM(TGUI[iLigne-2]);
    if ((OG=Nil) or (OG1=Nil)) then Continue ;

    // Récupère le n° de ligne
    iNumLigne:=OG1.GetMvt('EG_NUMLIGNE');

    // Affecte le n° de ligne à la ligne suivante
    OG.PutValue('EG_NUMLIGNE',iNumLigne);
    end;

  // Supprime la ligne
  TGUI.Delete(Lig);
  TGUI.Capacity := TGUI.Count;
  NbLG := TGUI.Count;
end;

procedure TFeSaisie.BMenuZoomClick(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
PopZoom97(BMenuZoom,POPZ) ;
{$ELSE}
{$ENDIF}

end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 25/08/2003
Modifié le ... :   /  /
Description .. : Uniquement pour CCMP...
Suite ........ :
Suite ........ : Réécriture de SetNumTraiteDocRegle, en remprenant le
Suite ........ : numéro de traite/Chèque directement dans la tob mise à
Suite ........ : jour depuis DocRegl lors de l'édition.
Suite ........ : Evite de refaire requête et d'oublier des maj.
Mots clefs ... :
*****************************************************************}
procedure TFeSaisie.SetNumTraiteDOCREGLE ;
Var i     : integer ;
    RR    : RMVT ;
    TobD  : TOB ;
    sTra  : String ;
begin
  // Tests
  if (M.smp<>smpEncTraEdtNC) And (M.smp<>smpDecBorEdtNC) then Exit ;
  if ((Action<>taCreat) and (Not SuiviMP)) then Exit ;

  // Supprimer les enregistrements temporaires
  ExecuteSQL('DELETE from DOCREGLE Where DR_USER="'+V_PGI.USER+'"') ;
  ExecuteSQL('DELETE from DOCFACT Where DF_USER="'+V_PGI.USER+'"') ;
  FiniMove ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 25/08/2003
Modifié le ... : 13/11/2003
Description .. : Pour utilisation dans CCMP uniqement...
Suite ........ : 
Suite ........ : Stocke dans TobNumChq la liste des numéro de traite/chq
Suite ........ : pour chaque pièce d'origine, afin de pouvoir la récupérer
Suite ........ : lors de la maj pendant le lettrage
Suite ........ : 
Suite ........ : // Modif fiche 12017 : pb maj E_NUMTRAITECHQ
Mots clefs ... : 
*****************************************************************}
procedure TFeSaisie.StockNumChqDepuisEche(LL: TList);
Var i         : integer ;
    O         : TOBM ;
    TobDetail : TOB ;
begin

  FreeAndNil(TOBNumChq) ;

  TOBNumChq := TOB.Create('VNUMCHQ', nil, -1);

  For i := 0 to ( LL.Count - 1 ) do
    begin
    O     := TOBM( LL[i] ) ;
    if (O.FieldExists('NUMTRAITECHQ')) and (O.GetValue('NUMTRAITECHQ') <> '')
                                       and (O.GetValue('E_TRACE') <> '' ) then
      begin
      TobDetail := TOB.Create('VNUMCHQDETAIL', TOBNumChq, - 1) ;
      TobDetail.AddChampsupValeur('NUMTRAITECHQ', O.GetValue('NUMTRAITECHQ')) ;
      TobDetail.AddChampsupValeur('CODELC',       O.GetValue('E_TRACE')) ;
      end ;
    end ;

end;

end.





