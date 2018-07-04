{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 02/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
unit eSaisEff ;

interface

uses utobdebug,
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Grids, Hctrls, Mask, ExtCtrls, ComCtrls,
  Buttons, Hspliter, Ent1,
  HCompte,            // GChercheCompte
  HEnt1,
  HTB97,
  Math,               // Min
  CPGENERAUX_TOM,     // FicheGene
  CPTIERS_TOM,        // FicheTiers
  CPJOURNAL_TOM,      // FicheAGLJournal
  eSaisAnal,          // eSaisieAnal, eSaisieAnalEff
  CPCHANCELL_TOF,     // FicheChancel
  DEVISE_TOM,         // FicheDevise
  SUIVCPTA_TOM,       // ParamScenario
  CPMULSAISLETT_TOF,  // LettrerEnSaisie
{$IFDEF EAGLCLIENT}
  MainEagl,
{$ELSE}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  DB,
  Fe_main,
{$ENDIF}
{$IFDEf VER150}
  Variants,
{$ENDIF}
  ULibAnalytique,     // AlloueAxe
  hmsgbox,
  HQry,
  SaisTVA,
  EcheMono,
  SaisComp,
  SaisUtil,
  SaisComm,          // AttribParamsNew, RecupTotalPivot, AttribParamsCompNew, ExecReqMAJNew
  LettUtil,
  Choix,
  Echeance,
  About,
  SaisVisu,
  HStatus,
  Filtre,
  SaisTaux1,
  FichComm,
  HDebug,
  HLines,
  HSysMenu,
  HPop97,
  SaisEnc,
  SaisBase,
  ed_tools,
  SaisEcar,
  HPanel,
  UiUtil,
{$IFDEF V530}
  EdtEtat,
{$ELSE}
  {$IFDEF EAGLCLIENT}
    UtileAGL,   // LanceEtat 
  {$ELSE}
    EdtREtat,   // LanceEtat
  {$ENDIF}
{$ENDIF}
  ImgList,
  UtilSais,
  UtilPGI,
  UTob,
  lookup,
  UlibEcriture, // FQ 13246 : SBO 30/03/2005
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  ParamSoc,     // GetParamSocSecur
  Menus, TntButtons, TntStdCtrls, TntExtCtrls, TntGrids ;

procedure SaisieEffet(ModeSaisie : tmodeSaisieEff ; Const ModeSR : tModeSR = srRien) ;

type
  TFSaisieEff = class(TForm)
    DockTop         : TDock97;
    DockRight       : TDock97;
    DockLeft        : TDock97;
    DockBottom      : TDock97;
    GS              : THGrid;
    PEntete         : THPanel;
    E_JOURNAL       : THValComboBox;
    H_JOURNAL       : THLabel;
    H_DATECOMPTABLE : THLabel;
    E_ETABLISSEMENT : THValComboBox;
    H_ETABLISSEMENT : THLabel;
    H_NUMEROPIECE   : THLabel;
    E_DEVISE        : THValComboBox;
    H_DEVISE        : THLabel;
    H_EXERCICE      : THLabel;
    PPied           : THPanel;
    G_LIBELLE       : THLabel;
    T_LIBELLE       : THLabel;
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
    POPS            : TPopupMenu;
    POPZ            : TPopupMenu;
    BZoom: THBitBtn;
    BZoomJournal: THBitBtn;
    BZoomEtabl: THBitBtn;
    Cache           : THCpteEdit;
    PForce          : TPanel;
    Label2          : TLabel;
    Label3          : TLabel;
    HForce          : THNumEdit;
    Image1          : TImage;
    FindSais        : TFindDialog;
    BScenario: THBitBtn;
    H_SOLDET        : THLabel;
    SA_SOLDETR      : THNumEdit;
    PPiedTreso      : THPanel;
    E_CREDITTRESO   : THNumEdit;
    E_DEBITTRESO    : THNumEdit;
    E_LIBTRESO      : TEdit;
    E_REFTRESO      : TEdit;
    E_CPTTRESO      : THPanel;
    Bevel2          : TBevel;
    Bevel1          : TBevel;
    Bevel3          : TBevel;
    Bevel4          : TBevel;
    Bevel5          : TBevel;
    E_MODEPAIE      : THValComboBox;
    H_MODEPAIE      : THLabel;
    E_NUMREF        : THNumEdit;
    Bevel6          : TBevel;
    Bevel7          : TBevel;
    HE_MODEP        : THLabel;
    HE_DATEECHE     : THLabel;
    BSwapPivot: THBitBtn;
    BSwapEuro: THBitBtn;
    BModifSerie: THBitBtn;
    BZoomT: THBitBtn;
    POPZT           : TPopupMenu;
    HMTrad          : THSystemMenu;
    EURO            : TEdit;
    Valide97        : TToolbar97;
    BValide         : TToolbarButton97;
    BAbandon: TToolbarButton97;
    BAide           : TToolbarButton97;
    Outils          : TToolbar97;
    BEche           : TToolbarButton97;
    BVentil         : TToolbarButton97;
    BComplement     : TToolbarButton97;
    BChercher       : TToolbarButton97;
    BRef            : TToolbarButton97;
    BLig            : TToolbarButton97;
    ToolbarSep971   : TToolbarSep97;
    ToolbarSep972   : TToolbarSep97;
    E_RIB           : THLabel;
    BModifRIB: THBitBtn;
    PRefAuto        : TPanel;
    H_TitreRefAuto  : TLabel;
    PFenGuide       : TPanel;
    BRAValide: THBitBtn;
    BRAAbandon: THBitBtn;
    H_REF           : THLabel;
    E_REF           : TEdit;
    E_NUMDEP        : TMaskEdit;
    H_DATEECHEANCE  : THLabel;
    H_CONTREPARTIEGEN : THLabel;
    HConf             : TToolbarButton97;
    ISigneEuro        : TImage;
    LE_DEBITTRESO     : THLabel;
    LE_CREDITTRESO    : THLabel;
    LSA_SOLDEG        : THLabel;
    LSA_SOLDET        : THLabel;
    LSA_TOTALDEBIT    : THLabel;
    LSA_TOTALCREDIT   : THLabel;
    LSA_SOLDETR       : THLabel;
    E_CONTREPARTIEGEN : THCpteEdit;
    BParamListeSaisieEff : TToolbarButton97;
    BZoomMvtEff          : TToolbarButton97;
    H_NATUREPIECE        : THLabel;
    E_NATUREPIECE        : THValComboBox;
    H_TYPECTR            : THLabel;
    E_TYPECTR            : THValComboBox;
    BSaisTaux: THBitBtn;
    BPopTva              : TPopButton97;
    BComplementT: THBitBtn;
    BVentilCtr           : TToolbarButton97;
    BMenuZoom            : TToolbarButton97;
    BMenuZoomT           : TToolbarButton97;
    POPMVT: TPopupMenu;
    POPZOOMFACT: TMenuItem;
    BZOOMFACT2: TMenuItem;
    compte: THEdit;
    E_DATECOMPTABLE: THCritMaskEdit;
    E_DATEECHEANCE: THCritMaskEdit;
    procedure FormShow(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure E_JOURNALChange(Sender: TObject);
    procedure E_DEVISEChange(Sender: TObject);
    procedure E_NATUREPIECEChange(Sender: TObject);
    procedure E_DATECOMPTABLEExit(Sender: TObject);
    procedure H_JOURNALDblClick(Sender: TObject);
    procedure GSExit(Sender: TObject);
    procedure GSMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure GSDblClick(Sender: TObject);
    procedure GSCellExit(Sender: TObject; var ACol, ARow: Longint;  var Cancel: Boolean);
    procedure GSCellEnter(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
    procedure GSRowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
    procedure GSRowExit(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
    procedure GSEnter(Sender: TObject);
    procedure BValideClick(Sender: TObject);
    procedure BVentilClick(Sender: TObject);
    procedure BEcheClick(Sender: TObject);
    procedure BZoomClick(Sender: TObject);
    procedure POPSPopup(Sender: TObject);
    procedure BModifRIBClick(Sender: TObject);
    procedure BSwapPivotClick(Sender: TObject);
    procedure BSwapEuroClick(Sender: TObject);
    procedure BComplementClick(Sender: TObject);
    procedure BChercherClick(Sender: TObject);
    procedure FindSaisFind(Sender: TObject);
    procedure BZoomEtablClick(Sender: TObject);
    procedure BSaisTauxClick(Sender: TObject);
    procedure BZoomJournalClick(Sender: TObject);
    procedure BScenarioClick(Sender: TObject);
    procedure BModifSerieClick(Sender: TObject);
    procedure GSKeyPress(Sender: TObject; var Key: Char);
    procedure BAideClick(Sender: TObject);
    procedure BMenuZoomMouseEnter(Sender: TObject);
    procedure BValideMouseEnter(Sender: TObject);
    procedure BValideMouseExit(Sender: TObject);
    procedure GereTagEnter(Sender: TObject);
    procedure GereTagExit(Sender: TObject);
    procedure BPopTvaChange(Sender: TObject);
    procedure E_DEVISEExit(Sender: TObject);
    procedure GereAffSolde(Sender: TObject);
    procedure E_DATECOMPTABLEEnter(Sender: TObject);
    procedure InitBoutonValide ;
    procedure E_NUMDEPExit(Sender: TObject);
    procedure BAbandonClick(Sender: TObject);
    procedure BLigClick(Sender: TObject);
    procedure BRefClick(Sender: TObject);
    procedure H_MODEPAIEDblClick(Sender: TObject);
    procedure E_MODEPAIEExit(Sender: TObject);
    procedure BZoomTClick(Sender: TObject);
    procedure E_JOURNALClick(Sender: TObject);
    procedure BMenuZoomMouseExit(Sender: TObject);
    procedure BRAAbandonClick(Sender: TObject);
    procedure BRAValideClick(Sender: TObject);
    procedure E_DATEECHEANCEExit(Sender: TObject);
    procedure E_CONTREPARTIEGENExit(Sender: TObject);
    procedure E_CONTREPARTIEGENChange(Sender: TObject);
    procedure BMenuZoomTMouseEnter(Sender: TObject);
    procedure BParamListeSaisieEffClick(Sender: TObject);
    procedure E_NATUREPIECEExit(Sender: TObject);
    procedure E_TYPECTRChange(Sender: TObject);
    procedure E_REFTRESOExit(Sender: TObject);
    procedure E_LIBTRESOExit(Sender: TObject);
    procedure BComplementTClick(Sender: TObject);
    procedure BVentilCtrClick(Sender: TObject);
    procedure BZOOMFACT2Click(Sender: TObject);
    procedure POPZOOMFACTClick(Sender: TObject);
  private
    SAJAL                      : TSAJAL ;
    DEV                        : RDEVISE ;
    NumPieceInt,NbLigEcr,NbLigAna : Longint ;
    CpteAuto,OkScenario,PieceConf,ModeRA,ChoixExige,ModeOppose,NeedEdition : boolean ;
    CurAuto,CurLig             : integer ;
    SorteTva                   : TSorteTva ;
    ExigeTva,ExigeEntete       : TExigeTva ;
    TS,TDELECR,TDELANA         : TList ;
    LAnaEff                    : T5LL ;
    M                          : RMVT ;
    Scenario,Entete,ModifSerie : TOBM ;
    GX,GY,LAUX,LHT             : integer ;
    OBMT                       : Array[1..2] Of TOBM ; //TR
    SI_NumRef                  : LongInt ;
    LigneTreso                 : Boolean ;//TR
    ModeCG,RentreDate          : boolean ;
    GeneRegTVA,GeneTVA,GeneTPF,GestionRIB : String3 ;
//    GeneSoumisTPF     : boolean ;
    RegimeForce                : boolean ;
    DejaRentre                 : Boolean ;
    GeneTypeExo                : TTypeExo ;
    OkMessAnal,AutoCharged     : boolean ;
    ModeDevise,FindFirst       : Boolean ;
    DecDev                     : integer ;
    ModeForce                  : tmSais ;
    NowFutur,DatePiece,DateEcheDefaut : TDateTime ;
    ModeConf                   : String[1] ;
    OldDevise,DernVentiltype,EtatSaisie   : String ;
    MajEnCours                 : Boolean ;
    MemoComp                   : HTStrings ;
    WMinX,WMinY                : Integer ;
    ModeSaisie                 : tmodeSaisieEff ;
    Volatile                   : Boolean ;
    ModeSR                     : tModeSR ;
    RegulAFaire,ResteLettrageAFaire : Boolean ;
    FInfo                      : TInfoEcriture ; // FQ 13246 : SBO 30/03/2005
    FClosing                   : Boolean ; {JP 14/05/07 : FQ 19254}
    FTobModePaie               : TOB; {JP 14/05/07 : FQ 18324}
    FBoInsertSpecif            : Boolean ;       // SBO 01/07/2007 : enregistrement spécifique pour pb CWAS
    FNeFermePlus               : Boolean; {JP 05/07/07 : FQ 19022 : Si on répond non à la confirmation fermeture}

    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
// Scénario de saisie
    procedure ChargeScenario ;
    procedure LettrageEnSaisie ;
// Init et defaut
    Function  QuelleListe : String ;
    function  OkPourLigne : Boolean ;
    function  DateToJour ( DD : TDateTime ) : String ;
    procedure SwapNatureCompte ;
    procedure InitLesComposants ;
    procedure InitEnteteJal(Totale,Zoom,ReInit : boolean) ;
    procedure InitGrid ;
    procedure DefautEntete ;
    procedure DefautPied ;
    procedure AttribNumeroDef ;
    procedure AttribLeTitre ;
    function  EstLettrable(Lig : Integer) : TLettre ;
    procedure ReInitPiece ;
    Procedure InitPiedTreso ;
    procedure GereEnabled ( Lig : integer ) ;
    procedure PosLesSoldes ;

// Chargements
    procedure ChargeLignes ;
//    procedure RuptureEche( Lig : integer ; TDD,TCD,TDP,TCP,TDE,TCE : Double ) ;
//    procedure AlimNoyauEche ( Lig : integer ; Premier : boolean ; LDD,LCD,LDP,LCP,LDE,LCE : Double ; QEcr : TDataSet ) ;
// Tva Enc
    Function  TvaTauxDirecteur ( Lig : integer ; Client,ToutDebit : boolean ; Regime : String3 ) : Boolean ;
    Function  RecopieTVAEnc ( Lig : integer ; Client,ToutDebit : boolean ; Regime : String3 ) : Boolean ;
    Function  GereTvaEncNonFact ( Lig : integer ) : Boolean ;
    procedure GetSorteTVA ;
//    procedure RenseigneRegime ( Lig : integer ) ;
//    Procedure AffecteTva ( C,L : integer ) ;
    procedure AfficheExigeTva ;
    Procedure ModifOBMPourTVAENC(i : Integer)  ;
// Allocation et Désallocation
    procedure AlloueEcheAnal(C,L : LongInt) ;
    procedure AlloueEche(LaLig : LongInt) ;
    procedure DesalloueEche(LaLig : LongInt) ;
    procedure AlloueAnal(LaLig : LongInt) ;
    procedure DesalloueAnal(LaLig : LongInt) ;
    Procedure AlloueOBML(LaLig : LongInt) ;
    Procedure DesAlloueOBML(LaLig : LongInt) ;
    procedure DesalloueLesOBML ;
// Object Mouvement, Divers
    procedure AlimObjetMvt(Lig : integer ; AvecM : boolean ; Treso : Byte) ;
    procedure GereComplements(Lig : integer) ;
    function  FermerSaisie : boolean ;
    procedure DecrementeNumero ;
    procedure GotoEntete ;
    procedure NumeroteVentils ;
    procedure QuelNExo ;
    procedure BasculeGS(AccesConsult : Boolean) ;
    Function  IsAccessible(Ou : LongInt) : Boolean ;
    Procedure BloqueGrille(Bloc : boolean ) ;
    Procedure FocusSurNatP ;
// Divers
    procedure ValideCeQuePeux( Lig : longint ) ;
    Function  FabricFromM : RMVT ;
    procedure StockeLaPiece ;
    procedure LettrageEffet ;
    Procedure EditionSaisie ;
    Procedure InitNTP ;
//    procedure InitTCTR ;
    Function  MvtCommun(LOBM,LOBM1 : TList) : Boolean ;
    Function  VerifAutreLigne ( Col,Lig : integer ) : Boolean ;
// Barre Outils
    procedure ClickVentil ;
    procedure ClickEche ;
    procedure ClickZoom(DoPopup : boolean = false);
    procedure RecupAnaSel(O : TOBM) ;
    Function  RecupMontantSel(LOBM : TList ; Var M : Double ; Var RIB : String ; OkAna : Boolean) : Boolean ;
//    Function  PrevientSiSelection(LOBM : tlist) : Boolean ;
    procedure DupliqueLigne(LOBM : Tlist ; O : TOBM ; Lig : Integer) ;
    procedure ClickZoomFact ;
    procedure ClickZoomFactConsult ;
    procedure ClickAbandon ;
//    Function  TrouveLigneTreso : Integer ;
    procedure ClickSwapPivot ;
    procedure ClickModifRIB ;
    procedure ClickSwapEuro ;
    procedure ClickComplement ;
    procedure AlimOBMMemoireEnBatch(OSource,ODestination : TOBM) ;
    procedure ClickComplementT ;
    procedure ClickCherche ;
    procedure ClickRAZLigne ;
    procedure ClickEtabl ;
    procedure ClickScenario ;
    procedure ClickJournal ;
    procedure ForcerMode(Cons : boolean ; Key : word) ;
    procedure ClickSaisTaux ;
// Calculs lignes
//    procedure InfoLigne( Lig : integer ; Var D,C,TD,TC : Double ) ;
    procedure DetruitLigne( Lig : integer ; Totale : boolean) ;
    Procedure PlusMvt(i : Integer ; Var SI_TDS,SI_TCS,SI_TDP,SI_TCP,SI_TDD,SI_TCD,SI_TDE,SI_TCE : Double) ;
    procedure CalculDebitCredit ;
    procedure AfficheConf ;
    procedure CalculSoldeCompte ( CpteG,CpteA : String ; DIG,CIG,DIA,CIA : Double ) ;
    procedure SommeSoldeCompte ( Col : integer ; Cpte : String ; Var TD,TC : Double ; Old,OkDev : Boolean) ;
    procedure DefautLigne ( Lig : integer ; Init : boolean ) ;
    procedure TraiteMontant( Lig : integer ; Calc : boolean ) ;
// Analytiques
    Procedure RecupLesVentilsCtr(NumAxe,Lig : Integer ; LL : HTStringList) ;
    Procedure OuvreAnal ( Lig : integer ; Scena,CtrLigne,NewLigne : boolean; vNumAxe : integer ) ;
    Function  AOuvrirAnal ( Lig : integer ; Auto,CtrLigne,NewLigne : boolean ) : Boolean ;
    Procedure GereAnal ( Lig : integer ; Auto,Scena,CtrLigne,NewLigne : boolean ) ;
    Procedure RecupAnal(Lig : integer ; var OBA : TOB ; NumAxe,NumV : integer ) ;
// Echéances
    Function  AOuvrirEche ( Lig : integer ; Var Cpte : String17 ; Var MR : String3 ; Var OuvreAuto,RempliAuto : boolean ; Var t : TLettre ) : Boolean ;
    Procedure DebutOuvreEche(Lig : Integer ; Cpte : String17 ; MR : String3 ; OMODR : TMOD) ;
    Procedure FinOuvreEche (Lig : Integer ; OMODR : TMOD ; t : TLettre) ;
    Procedure AlimEcheVersObm(Lig : Integer ; ModePaie : String3 ; DateEche,DateValeur : TDateTime) ;
//    Procedure MajOBMViaEche(Lig : integer ; OBM : TOBM ; R : T_ModeRegl ; NumEche : integer ) ;
    Procedure OuvreEche ( Lig : integer ; Cpte : String17 ; MR : String3 ; RempliAuto : boolean ; t : TLettre ; LigTresoEnPied : Boolean ) ;
    Procedure GereEche ( Lig : integer ; OuvreAuto,RempliAuto : boolean ) ;
    Procedure RecupEche(Lig : integer ; R : T_ModeRegl ; NumEche : integer; Var OBM : TOBM  ) ;
// Contrôles
    Function  LigneCorrecte ( Lig : integer ; Totale,Alim,CtrlCptTreso : boolean ) : boolean ;
    Function  PieceCorrecte(OkMess : Boolean) : boolean ;
    Function  PasModif : boolean ;
    Function  Equilibre : boolean ;
    procedure ErreurSaisie ( Err : integer ) ;
    procedure AjusteLigne ( Lig : integer ) ;
    Function  PossibleRecupNum : Boolean ;
    procedure ControleLignes  ;
// Appels comptes
//    Procedure QuelZoomTable(C,L : Integer) ;
    Function  ChercheGen(C,L : integer ; Force : boolean ; DoPopup : boolean = false) : byte;
    Function  ChercheAux(C,L : integer ; Force : boolean ; DoPopup : boolean = false) : byte;
//    Function  ChercheGen ( C,L : integer ; Force : boolean ) : byte ;
//    Function  ChercheAux ( C,L : integer ; Force : boolean ) : byte ;
    Function  ChargeCompteEffetJAL : Boolean ;
    procedure ChercheJAL ( Jal : String3 ; Zoom : boolean ) ;
    procedure ChercheMontant(Acol,Arow : longint) ;
    procedure AppelAuto ( indice : integer ) ;
    Function  AffecteRIB ( C,L : integer ; IsAux : Boolean) : Boolean ;
    Procedure AffecteConso ( C,L : integer ) ;
    Function  VerifEtUpdateRib ( L : integer ) : Boolean ;
// MAJ Fichier
    Procedure GetEcr(Lig : Integer) ;
    Procedure GetAna(Lig : Integer) ;
    Procedure RecupTronc(Lig : Integer ; MO : TMOD ; Var OBM : TOBM);
    Procedure ClickValide ;
    procedure ValideLaPiece ;
    procedure ValideLeReste ;
    Procedure ValideLignesGene ;
    Procedure ValideLesComptes ;
    procedure ValideLeJournalNew ;
    procedure MajCptesGeneNew ;
    procedure MajCptesAuxNew  ;
    procedure MajCptesSectionNew ;
    function  QuelEnc ( Lig : Integer ) : String3 ;
// Affichages, Positionnements
    Procedure AutoSuivant( Suiv : boolean ) ;
    Procedure SoldelaLigne ( Lig : integer ) ;
    procedure AffichePied ;
    procedure GereNewLigneT ;
    procedure PositionneDevise(ReInit : boolean) ;
    procedure AvertirPbTaux(Code : String3 ; DateTaux,DateCpt : TDateTime) ;
    Function  PbTaux ( DEV : RDevise ; DateCpt : TDateTime ) : boolean ;
// Calcul sur ligne des trésorerie
    function  TrouveSensTreso ( GS : THGrid ; Nat : String3 ; Lig : integer ) : byte ;
    procedure ClickLibelleAuto ;
    procedure CalculRefAuto(Sender : TObject) ;
    procedure ChargeLibelleAuto ;
    procedure ModifNaturePiece ;
    procedure ModifModePaie ;
    procedure AfficheLigneTreso ;
    Procedure GereLigneTreso ( Lig : integer ) ;
    procedure CalculSoldeCompteT ;
    Procedure AlimLigneTreso(Lig : Integer) ;
    Procedure RecopieGS(i : Integer) ;
    Procedure LigneEnPlusTreso(Sens : Integer ; MSaisi,MPivot,MDevise : Double) ;
    Procedure MajGridPiedTreso ;
    procedure MajRefLibOBMT(Quoi : Integer) ;
    function  Load_Sais ( St : hString ) : Variant ;
// Gestion du combo mode paiement
(*
    Procedure ShowModePaie(ACol,ARow: Longint) ;
*)
// Conversions, Caluls
    Function  ArrS( X : Double ) : Double ;
    Function  DateMvt ( Lig : integer ) : TDateTime ;
// Gestion référence entete
    Procedure IncRef ;
    procedure CloseFen ;
// Grid
    procedure PostDrawCell(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
    Procedure ReinitGrid ;
    Function  PlusModePaie : String ;
    procedure ReinitNumCol(St : String) ;
    function  GetModePaie  (ModePaie : string) : TOB; {JP 14/05/07 : FQ 18324 : gestion du codeacceptation}
    function  GetCodeAccept(ModePaie : string) : string; {JP 14/05/07 : FQ 18324 : gestion du codeacceptation}
    function  TestQteAna : boolean ;
  public
    Action : TActionFiche ;
    TPIECE : TList ;
    SI_TotDS,SI_TotCS,SI_TotDP,SI_TotCP,SI_TotDD,SI_TotCD,SI_TotDE,SI_TotCE : Double ;
    SI_FormuleRef,SI_FormuleLib : String ;
    SC_TotDS,SC_TotCS,SC_TotDP,SC_TotCP,SC_TotDD,SC_TotCD,SC_TotDE,SC_TotCE : Double ;
    {JP 20/10/05 : FQ 16923 : On applique le même système que dans Saisie.Pas, à savoir
                   que l'on renseigne uniquement les lignes vides}
    procedure GereRefLib(Lig : Integer; var StRef, StLib : string; LigOk : Boolean);
    // DEV 3216 10/04/2006 SBO
    function  GetFormulePourCalc(Formule: hstring): Variant;
    function  GetFormatPourCalc : string ;
  end;

implementation

Uses
  {$IFDEF MODENT1}
  CPProcGen,
  CPProcMetier,
  CPVersion,
  {$ENDIF MODENT1}
  Formule,
  ParamDBG,
  {$IFDEF EAGLCLIENT}
  CPMULFACTEFF_TOF,  // PointeMvtCpt
  {$ELSE}
  MulFactEff,
  UtilSOC, {MarquerPublifi}
  {$ENDIF EAGLCLIENT}
  CPZOOMFACTEFF_TOF, // ZoomFactCliEff
  CPLETREGUL_TOF,    // RegulLettrageMPAuto
  ParamSaisEff,      // ParamSaisieListeEff
  LetBatch
  ;

{$R *.DFM}

Var VEFFET : tmodeSaisieEff ;
    VMODE  : tModeSR ;

procedure LanceSaisietr ( Action : TActionFiche ; Var M : RMVT ; ModeSaisie : tmodeSaisieEff ; Const ModeSR : tModeSR = srRien) ;
Var
  X  : TFSaisieEff ;
  OA : TActionFiche ;
  PP : THpanel ;
BEGIN
  OA:=Action ;
  Case Action of
    taCreat : BEGIN
                if PasCreerDate(V_PGI.DateEntree) then Exit ;
                if DepasseLimiteDemo then Exit ;
                if Blocage(['nrCloture'],True,'nrSaisieCreat') then Exit ;
              END ;
    taModif : BEGIN
                if RevisionActive(M.DateC) then Exit ;
                if Blocage(['nrCloture','nrBatch'],True,'nrSaisieModif') then Exit ;
              END ;
  END ;
  VEFFET:=ModeSaisie ;
  VMODE:=ModeSR ;
  X:=TFSaisieEff.Create(Application) ;
  X.ModeSR:=ModeSR ;
  X.ModeSaisie:=ModeSaisie ;
  X.NowFutur:=NowH ;
  X.Action:=Action ;
  X.M:=M ;
  PP:=FindInsidePanel ;
  if PP=Nil then BEGIN
    Try
      X.ShowModal ;
    Finally
      X.Free ;
      Case OA of
         taCreat : Bloqueur('nrSaisieCreat',False) ;
         taModif : Bloqueur('nrSaisieModif',False) ;
      END ;
    End ;
    Screen.Cursor:=crDefault ;
    END
  else BEGIN
    InitInside(X,PP) ;
    X.Show ;
  END ;
END ;

procedure SaisieEffet(ModeSaisie : tmodeSaisieEff ; Const ModeSR : tModeSR = srRien) ;
Var
  M : RMVT ;
BEGIN
  FillChar(M,Sizeof(M),#0) ;
  M.CodeD:=V_PGI.DevisePivot ;
  M.DateC:=V_PGI.DateEntree ;
  M.TauxD:=1 ;
  M.DateTaux:= M.DateC ;
  M.Valide:=False ;
  M.Etabl:=VH^.ETABLISDEFAUT ;
  M.ANouveau:=FALSE;
  M.Treso:=TRUE ;
  M.MajDirecte:=FALSE ;
  M.Simul:='N' ;
  M.Effet:=TRUE ;
  M.Nature:='OD' ;
  LanceSaisietr(taCreat,M,ModeSaisie,ModeSR) ;
END ;

procedure TFSaisieEff.ChargeLibelleAuto ;
Var
  Q : TQuery ;
BEGIN
  if E_JOURNAL.Value='' then Exit ;
  if E_NATUREPIECE.Value='' then Exit ;
  if PasModif then Exit ;
  Q:=OpenSQL('Select RA_FORMULEREF, RA_FORMULELIB from REFAUTO Where RA_JOURNAL="'+E_JOURNAL.Value+'" AND RA_NATUREPIECE="'+E_NATUREPIECE.Value+'"',True) ;
  if Not Q.EOF then BEGIN
    SI_FormuleRef:=Q.FindField('RA_FORMULEREF').AsString ;
    SI_FormuleLib:=Q.FindField('RA_FORMULELIB').AsString ;
    AutoCharged:=True ;
    If Not Q.Eof Then BEGIN
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
{procedure TFSaisieEff.AlimNoyauEche ( Lig : integer ; Premier : boolean ; LDD,LCD,LDP,LCP,LDE,LCE : Double ; QEcr : TDataSet ) ;
Var TM : TMOD ;
    QEcr : TQuery;
BEGIN
TM:=TMOD(GS.Objects[SA_NumP,Lig]) ; if TM=Nil then Exit ;
With TM.MODR do
     BEGIN
     if Premier then
        BEGIN
        Action:=taModif ; NbEche:=0 ; CodeDevise:=DEV.Code ; Symbole:=DEV.Symbole ;
        TauxDevise:=DEV.Taux ; Quotite:=DEV.Quotite ; Decimale:=DecDev ;
        DateFact:=StrToDate(E_DATECOMPTABLE.Text) ; DateBl:=DateFact ; DateFactExt:=DateFact ;
        Aux:=GS.Cells[SA_Aux,Lig] ; if Aux='' then Aux:=GS.Cells[SA_Gen,Lig] ;
        TotalAPayerP:=0 ; TotalAPayerD:=0 ; TotalAPayerE:=0 ;
        END ;
     TotalAPayerP:=TotalAPayerP+LDP+LCP ;
     TotalAPayerD:=TotalAPayerD+LDD+LCD ;
     TotalAPayerE:=TotalAPayerE+LDE+LCE ;
     Inc(NbEche) ;
     With TabEche[NbEche] do
        BEGIN
        MontantD:=LDD+LCD ; MontantP:=LDP+LCP ; MontantE:=LDE+LCE ;
        ReadOnly:=(Trim(QEcr.FindField('E_LETTRAGE').AsString)<>'') ;
        ModePaie:=QEcr.FindField('E_MODEPAIE').AsString ;
        DateEche:=QEcr.FindField('E_DATEECHEANCE').AsDateTime ;
        DateRelance:=QEcr.FindField('E_DATERELANCE').AsDateTime ;
        NiveauRelance:=QEcr.FindField('E_NIVEAURELANCE').AsInteger ;
        Couverture:=QEcr.FindField('E_COUVERTURE').AsFloat ;
        CouvertureDev:=QEcr.FindField('E_COUVERTUREDEV').AsFloat ;
        CouvertureEuro:=QEcr.FindField('E_COUVERTUREEURO').AsFloat ;
        LettrageDev:=QEcr.FindField('E_LETTRAGEDEV').AsString ;
        DatePaquetMax:=QEcr.FindField('E_DATEPAQUETMAX').AsDateTime ;
        DatePaquetMin:=QEcr.FindField('E_DATEPAQUETMIN').AsDateTime ;
        EtatLettrage:=QEcr.FindField('E_ETATLETTRAGE').AsString ;
        DateValeur:=QEcr.FindField('E_DATEVALEUR').AsDateTime ;
        END ;
     END ;
END ;}

{procedure TFSaisieEff.RuptureEche( Lig : integer ; TDD,TCD,TDP,TCP,TDE,TCE : Double ) ;
Var OBM : TOBM ;
    TM  : TMOD ;
    i   : integer ;
BEGIN
if ModeDevise then BEGIN GS.Cells[SA_Debit,Lig]:=StrS(TDD,DECDEV) ; GS.Cells[SA_Credit,Lig]:=StrS(TCD,DECDEV) ; END
              else BEGIN GS.Cells[SA_Debit,Lig]:=StrS(TDP,DECDEV) ; GS.Cells[SA_Credit,Lig]:=StrS(TCP,DECDEV) ; END ;
FormatMontant(GS,SA_Debit,Lig,DECDEV) ;
FormatMontant(GS,SA_Credit,Lig,DECDEV) ;
OBM:=GetO(GS,Lig) ;
if OBM<>Nil then BEGIN
  OBM.PutMvt('E_DEBIT',TDP)    ; OBM.PutMvt('E_CREDIT',TCP) ;
  OBM.PutMvt('E_DEBITDEV',TDD) ; OBM.PutMvt('E_CREDITDEV',TCD) ;
  OBM.PutMvt('E_DEBITEURO',TDE) ; OBM.PutMvt('E_CREDITEURO',TCE) ;
  OBM.HistoMontants ;
END ;
  TM:=TMOD(GS.Objects[SA_NumP,Lig]) ;
  if TM<>Nil then if TM.ModR.TotalAPayerP<>0 then BEGIN
    for i:=1 to TM.ModR.NbEche do
      TM.ModR.TabEche[i].Pourc := Arrondi(100.0*TM.ModR.TabEche[i].MontantP/TM.ModR.TotalAPayerP,ADecimP) ;
  END ;
END ;}

procedure TFSaisieEff.ChargeLignes ;
BEGIN
  CalculDebitCredit ;
  AffichePied ;
END ;

{==========================================================================================}
{=========================== Allocations et Désallocations ================================}
{==========================================================================================}
procedure TFSaisieEff.AlloueEcheAnal(C,L : LongInt) ;
Var
  CGen : TGGeneral ;
  CAux : TGTiers ;
BEGIN
  if C=SA_Gen then BEGIN
    CGen:=GetGGeneral(GS,L) ; if CGen=Nil then Exit ;
    if Lettrable(CGen)<>NonEche then AlloueEche(L) ;
    if Ventilable(CGen,0) then AlloueAnal(L) ;
    END
  else BEGIN
    CAux:=GetGTiers(GS,L) ; if CAux=Nil then Exit ;
    if CAux.Lettrable then AlloueEche(L) ;
  END ;
END ;

procedure TFSaisieEff.AlloueEche(LaLig : LongInt) ;
Var
  TTM : TMOD ;
begin
  if GS.Objects[SA_NumP,LaLig]<>Nil then Exit ;
  TTM:=TMOD.Create ;
  GS.Objects[SA_NumP,LaLig]:=TObject(TTM) ;
  TMOD(GS.Objects[SA_NumP,LaLig]).MODR.TabEche[1].DateEche:=DateMvt(LaLig) ;
  TMOD(GS.Objects[SA_NumP,LaLig]).MODR.TabEche[1].ModePaie:=E_MODEPAIE.VALUE ;
end ;

procedure TFSaisieEff.DesalloueEche(LaLig : LongInt) ;
Var
  OBM : TOBM ;
begin
  if GS.Objects[SA_NumP,LaLig]=Nil then Exit ;
  TMOD(GS.Objects[SA_NumP,LaLig]).Free ;
  GS.Objects[SA_NumP,LaLig]:=Nil ;
  OBM:=GetO(GS,LaLig) ;
  if OBM=Nil then Exit ;
  // Re-init des infos echéances
  OBM.PutMvt('E_LETTRAGE','') ;
  OBM.PutMvt('E_LETTRAGEDEV','-') ;
  OBM.PutMvt('E_COUVERTURE',0) ;
  OBM.PutMvt('E_COUVERTUREDEV',0) ;
  OBM.PutMvt('E_MODEPAIE','') ;
  OBM.PutMvt('E_NIVEAURELANCE',0) ;
  OBM.PutMvt('E_ETATLETTRAGE','RI') ;
  OBM.PutMvt('E_NUMECHE',0) ;
  OBM.PutMvt('E_ECHE','-') ;
  OBM.PutMvt('E_DATEECHEANCE',IDate1900) ;
  OBM.PutMvt('E_DATERELANCE',IDate1900) ;
end ;

function  TFSaisieEff.DateToJour ( DD : TDateTime ) : String ;
Var
  y,m,d : Word ;
  StD   : String ;
BEGIN
  DecodeDate(DD,y,m,d) ;
  StD:=IntToStr(d) ;
  if d<10 then StD:='0'+StD ;
  Result:=StD ;
END ;

procedure TFSaisieEff.AlloueAnal(LaLig : LongInt) ;
var OBM : TOBM ;
begin
	OBM := GetO(GS,LaLig) ;
  if OBM = nil then Exit;
	AlloueAxe(OBM) ;
end ;

procedure TFSaisieEff.DesalloueAnal(LaLig : LongInt) ;
Var
  OBM : TOBM ;
begin
  OBM:=GetO(GS,LaLig) ;
  if OBM.Detail.Count=0 then Exit ;
  OBM.ClearDetail;
end ;

procedure TFSaisieEff.AlloueOBML(LaLig : LongInt) ;
Var
  L : TList ;
begin
  if GS.Objects[SA_NumL,LaLig]<>Nil then Exit ;
  L:=TList.Create ;
  GS.Objects[SA_NumL,LaLig]:=TObject(L) ;
end ;

procedure TFSaisieEff.DesalloueOBML(LaLig : LongInt) ;
Var
  L : TList ;
begin
  L:=TList(GS.Objects[SA_NumL,LaLig]) ;
  if L=Nil then Exit ;
  VideListe(L) ;
  L.Free ;
  GS.Objects[SA_NumL,LaLig]:=Nil ;
end ;

procedure TFSaisieEff.DesalloueLesOBML ;
Var
  i : Integer ;
BEGIN
  For i:=1 To GS.RowCount-1 Do
    DesalloueOBML(i) ;
end;

{==========================================================================================}
{=========================== Initialisations et valeurs par défaut ========================}
{==========================================================================================}

procedure TFSaisieEff.SwapNatureCompte ;
BEGIN
  E_NATUREPIECE.Visible:=(ModeSaisie=OnBqe) ;
  H_NATUREPIECE.Visible:=(ModeSaisie=OnBqe) ;
  E_CONTREPARTIEGEN.Visible:=(ModeSaisie<>OnBqe) ;
  H_CONTREPARTIEGEN.Visible:=(ModeSaisie<>OnBqe) ;
END ;

procedure TFSaisieEff.InitLesComposants ;
begin
  Bref.Visible:=TRUE ;
  BLig.Enabled:=TRUE ;
  H_TYPECTR.Visible:=(ModeSaisie=OnBqe) ;
  E_TYPECTR.Visible:=(ModeSaisie=OnBqe) ;
  H_SOLDET.Visible:=TRUE ;
  LSA_SOLDETR.Visible:=TRUE ;
  PPiedTreso.Visible:=TRUE ;
  Case ModeSaisie Of
    OnEffet : BEGIN
                E_MODEPAIE.PLUS:='(MP_ENCAISSEMENT="MIX" OR MP_ENCAISSEMENT="ENC") AND MP_CATEGORIE="LCR" ' ;
                E_JOURNAL.DataType:='TTJALEFFET' ;
              END ;
    OnChq : BEGIN
              E_MODEPAIE.PLUS:='(MP_ENCAISSEMENT="MIX" OR MP_ENCAISSEMENT="ENC") AND MP_CATEGORIE="CHQ" ' ;
              E_JOURNAL.DataType:='TTJALEFFET' ;
            END ;
    OnCB : BEGIN
             E_MODEPAIE.PLUS:='(MP_ENCAISSEMENT="MIX" OR MP_ENCAISSEMENT="ENC") AND MP_CATEGORIE="CB" ' ;
             E_JOURNAL.DataType:='TTJALEFFET' ;
           END ;
    OnBqe : BEGIN
              Case ModeSR Of
                srCli : BEGIN
                          E_MODEPAIE.PLUS:='(MP_ENCAISSEMENT="MIX" OR MP_ENCAISSEMENT="ENC") ' ;
                          E_NATUREPIECE.PLUS:='AND (CO_CODE="RC" OR CO_CODE="OC" OR CO_CODE="OD")' ;
                        END ;
                srFou : BEGIN
                          E_MODEPAIE.PLUS:='(MP_ENCAISSEMENT="MIX" OR MP_ENCAISSEMENT="DEC") ' ;
                          E_NATUREPIECE.PLUS:='AND (CO_CODE="RF" OR CO_CODE="OF" OR CO_CODE="OD")' ;
                        END ;
              END ;
              E_JOURNAL.DataType:='TTJALBANQUE' ;
            END ;
  END ;
  SwapNatureCompte ;
  E_CONTREPARTIEGEN.Enabled:=(ModeSaisie<>OnBqe) ;
  H_CONTREPARTIEGEN.Enabled:=(ModeSaisie<>OnBqe) ;
  E_NATUREPIECE.Enabled:=(ModeSaisie=OnBqe) ;
  H_NATUREPIECE.Enabled:=(ModeSaisie=OnBqe) ;
end ;

procedure TFSaisieEff.GereEnabled ( Lig : integer ) ;
Var OKL,Remp : Boolean ;
    O   : TOBM ;
    CGen : TGGeneral ;
    CAux : TGTiers ;
    OkZoomMvt : Boolean ;
BEGIN
  Remp:=EstRempli(GS,Lig) ;
  O:=GetO(GS,Lig) ;
  OkL:=((Remp) and (O<>Nil)) ;
  BEche.Enabled:=(GS.Objects[SA_NumP,Lig]<>Nil) ;
  BVentil.Enabled:=(GS.Objects[SA_DateC,Lig]<>Nil) ;
  If (SAJAL<>NIL) Then BEGIN
    Case SAJAL.TRESO.TypCtr Of
      Ligne : BEGIN
                BVentilCtr.Visible:=TRUE ;
                If Lig+1<=GS.RowCount-1 Then BVentilCtr.Enabled:=(GS.Objects[SA_DateC,Lig+1]<>Nil)
                                        Else BVentilCtr.Enabled:=FALSE ;
              END ;
      Else BVentilCtr.Visible:=FALSE ;
    END ;
  END ;
  BVentilCtr.Enabled:=(GS.Objects[SA_DateC,Lig]<>Nil) ;
  BComplement.Enabled:=((OkL) and (Not M.ANouveau)) ;
  BSwapPivot.Enabled:=(GS.RowCount>2) and ((ModeDevise) or (ModeForce=tmPivot)) ;
  BSwapEuro.Enabled:=GS.RowCount>2 ;

  {Gestion du RIB}
//  OkRIB:=OkScenario and ((GestionRIB='MAN') or (GestionRIB='PRI')) ;
  BModifRIB.Enabled:=((OkL) and (BEche.Enabled)) ;
  OkZoomMvt:=FALSE ;
  CGen:=GetGGeneral(GS,Lig) ;
  if (CGen<>Nil) And (Lettrable(CGen)=MultiEche) Then
    OkZoomMvt:=TRUE
  Else BEGIN
    CAux:=GetGTiers(GS,Lig) ;
    if (CAux<>Nil) And CAux.Lettrable then OkZoomMvt:=TRUE ;
  END ;
  If SAJAL<>NIL Then
    If GS.Cells[SA_Gen,Lig]=SAJAL.TRESO.Cpt Then OkZoomMvt:=FALSE ;
  BZoomMvtEff.Enabled:=OkZoomMvt ;
END ;

function TFSaisieEff.EstLettrable(Lig : Integer) : TLettre ;
Var CGen : TGGeneral ;
    CAux : TGTiers ;
begin
  Result:=NonEche ;
  CGen:=GetGGeneral(GS,Lig) ;
  if CGen=Nil then exit ;
  Result:=Lettrable(CGen) ;
  If Result In [MonoEche,MultiEche] then Exit ;
  CAux:=GetGTiers(GS,Lig) ;
  if CAux=Nil then Exit ;
  If CAux.Lettrable then Result:=MultiEche ;
end ;

procedure TFSaisieEff.ReInitPiece ;
Var bc : boolean ;
BEGIN
  DesalloueLesOBML ;
  GS.VidePile(True) ;
  VideListe(TS) ;
  TS.Free ;
  TS:=Nil ;
  TS:=TList.Create ;
  if Action<>taCreat then Exit ;
  // Ré-init des infos comptes / auxi pour rechargement des totaux...
  if Assigned( FInfo )  then // FQ 13246 SBO 04/08/2005
    begin
    FInfo.Compte.Clear ;
    FInfo.Aux.Clear ; 
    end ;
  BValide.Enabled:=True ;
  DefautPied ;
  NowFutur:=NowH ;
  H_NUMEROPIECE.Visible:=False ;
  H_NUMEROPIECE_.Visible:=True ;
  SwapNatureCompte ;
  E_DATECOMPTABLE.Enabled:=True ;
  If ModeSaisie<>OnBqe Then BEGIN
    E_CONTREPARTIEGEN.Enabled:=TRUE ;
    END
  Else BEGIN
    E_NATUREPIECE.Enabled:=TRUE ;
  END ;
//  E_CONTREPARTIEGEN.Enabled:=TRUE ;
  E_JOURNAL.Enabled:=True ;
  InitEnteteJal(False,False,True) ;
  BParamListeSaisieEff.Enabled:=TRUE ;
  DateEcheDefaut:=0 ;
  DateEcheDefaut:=StrToDate(E_DATEECHEANCE.Text) ;
  ModifSerie.Free ;
  ModifSerie:=Nil ;
  ModifSerie:=TOBM.Create(EcrGen,'',True) ;
  OBMT[1].Free;
  OBMT[1]:=NIL ;
  OBMT[1]:=TOBM.Create(EcrGen,'',TRUE) ;
  OBMT[2].Free;
  OBMT[2]:=NIL ;
  OBMT[2]:=TOBM.Create(EcrGen,'',TRUE) ;
  AlimObjetMvt(0,FALSE,1) ;
  AlimObjetMvt(0,FALSE,2) ;
  OkMessAnal:=False ;
  DejaRentre:=False ;
  RegimeForce:=False ;
  GS.SetFocus ;
  GS.Col:=SA_GEN ;
  GS.Row:=1 ;
  PieceConf:=False ;
  RentreDate:=False ;
  VideListe(TDELECR) ;
  VideListe(TDELANA) ;
  ModeConf:='0' ;
  Volatile:=False ;
//  if Action=taCreat then GSEnter(Nil) ; GSRowEnter(Nil,1,bc,False) ;

  {Affichage}
  if ((VH^.CPDateObli) and (Action=taCreat) and (E_DATECOMPTABLE.CanFocus)) then
    E_DATECOMPTABLE.SetFocus
  else BEGIN
    if Action=taCreat then GSEnter(Nil) ;
    GSRowEnter(Nil,1,bc,False) ;
  END ;
END ;

Procedure TFSaisieEff.InitPiedTreso ;
begin
  E_CPTTRESO.Visible:=TRUE;
  E_REFTRESO.Visible:=TRUE;
  E_LIBTRESO.Visible:=TRUE ;
  LE_DEBITTRESO.Visible:=TRUE ;
  LE_CREDITTRESO.Visible:=TRUE ;
  PPiedTreso.Visible:=TRUE ;
  E_REFTRESO.Enabled:=FALSE ;
  E_LIBTRESO.Enabled:=FALSE ;
  E_REFTRESO.Color:=clBtnFace ;
  E_LIBTRESO.Color:=clBtnFace ;
  BLig.Enabled:=FALSE ;
  E_REFTRESO.Visible:=FALSE ;
  E_LIBTRESO.Visible:=FALSE ;
  BLig.Enabled:=TRUE ;
  Case SAJAL.TRESO.TypCtr Of
    PiedDC,PiedSolde : BEGIN
                         E_REFTRESO.Enabled:=TRUE ; E_LIBTRESO.Enabled:=TRUE ;
                         E_REFTRESO.Text:='' ; E_LIBTRESO.Text:='' ;
                         E_REFTRESO.Color:=clWhite ;
                         E_LIBTRESO.Color:=clWhite ;
                       END ;
    Ligne : BEGIN
//              E_CPTTRESO.Visible:=FALSE   ;
//              E_DEBITTRESO.Visible:=FALSE ; E_CREDITTRESO.Visible:=FALSE ; PPiedTreso.Visible:=FALSE ;
              E_REFTRESO.Visible:=FALSE ;
              E_LIBTRESO.Visible:=FALSE ;
              BLig.Enabled:=TRUE ;
            END ;
    AuChoix : BEGIN
                E_CPTTRESO.Visible:=FALSE   ;
                LE_DEBITTRESO.Visible:=FALSE ;
                LE_CREDITTRESO.Visible:=FALSE ;
                PPiedTreso.Visible:=FALSE ;
              END ;
  END ;
end ;

procedure TFSaisieEff.InitEnteteJal(Totale,Zoom,ReInit : boolean) ;
Var MM  : String17 ;
    DD : TDateTime ;
BEGIN
  if Action<>taCreat then Exit ;
  if SAJAL=Nil then Exit ;
  PositionneDevise(ReInit) ;
  If ModeSaisie=OnEffet Then BEGIN
    E_NaturePiece.DataType:='ttNaturePiece' ;
    E_NaturePiece.Value:='OD' ;
    END
  Else BEGIN
    E_NaturePiece.DataType:='ttNatPieceBanque' ;
    If MODESR=srCli Then E_NaturePiece.Value:='RC'
                    Else E_NaturePiece.Value:='RF' ;
  END ;
  if Not Zoom then BEGIN
    DD:=StrToDate(E_DATECOMPTABLE.Text) ;
    InitNTP ;
    NumPieceInt:=GetNum(EcrGen,SAJAL.COMPTEURNORMAL,MM,DD) ;
    E_CPTTRESO.Caption:=SAJAL.Treso.Cpt ;
    E_TYPECTR.VALUE:=SAJAL.TRESO.STypCtr ;
    If ((Not ReInit) and (ModeSaisie=OnBqe)) Then E_DEVISE.VALUE:=SAJAL.TRESO.DevBqe ;
    InitPiedTreso ;
    InitGrid ;
  END ;
  if NumPieceInt>0 then E_NUMEROPIECE.Caption:=IntToStr(NumPieceInt) ;
END ;

procedure TFSaisieEff.DefautEntete ;
BEGIN
  DateEcheDefaut:=0 ;
  E_DATECOMPTABLE.Text:=DateToStr(M.DateC) ;
  E_DATEECHEANCE.Text:=E_DATECOMPTABLE.Text ;
  DateEcheDefaut:=M.DateC ;
  SI_NumRef:=-1 ;
  E_NUMREF.Visible:=FALSE ;
  E_JOURNAL.Value:=M.JAL ;
  InitNTP ;
  E_ETABLISSEMENT.Value:=M.Etabl ;
  E_ETABLISSEMENT.Enabled:=VH^.EtablisCpta ;
  DatePiece:=M.DateC ;
  If ModeSaisie<>OnBqe Then H_SOLDET.Caption:=HDiv.Mess[19]  // Solde Contrepartie
                       Else H_SOLDET.Caption:=HDiv.Mess[18]; // Solde Trésorerie
  if ((ModeOppose) and (M.CodeD=V_PGI.DevisePivot)) then BEGIN
    E_DEVISE.ItemIndex:=E_DEVISE.Values.Count-1 ;
    E_DEVISEChange(Nil) ;
    END
  else BEGIN
    E_DEVISE.Value:=M.CodeD ;
  END ;

  Case Action of
    taCreat : BEGIN
                if (M.DateC<V_PGI.DateDebutEuro) then E_DEVISE.Enabled:=False ;
//                E_DEVISE.Enabled:=TRUE ;
                If ModeSaisie<>OnBqe Then E_DEVISE.Enabled:=TRUE ;
                E_NumeroPiece.Caption:='' ; GS.Enabled:=False ;
                H_NUMEROPIECE.Visible:=False ; H_NUMEROPIECE_.Visible:=True ;
                E_REF.Text:='' ; E_NUMDEP.Text:='' ;
              END ;
  END ;
// TR
  If OBMT[1]=NIL Then OBMT[1]:=TOBM.Create(EcrGen,'',TRUE) ;
  If OBMT[2]=NIL Then OBMT[2]:=TOBM.Create(EcrGen,'',TRUE) ;
END ;

procedure TFSaisieEff.DefautPied ;
BEGIN
  SI_TotDS:=0 ; SI_TotCS:=0 ; SI_TotDP:=0 ; SI_TotCP:=0 ;
  SI_TotDD:=0 ; SI_TotCD:=0 ; SI_TotDE:=0 ; SI_TotCE:=0 ;
  SC_TotDS:=0 ; SC_TotCS:=0 ; SC_TotDP:=0 ; SC_TotCP:=0 ;
  SC_TotDD:=0 ; SC_TotCD:=0 ; SC_TotDE:=0 ; SC_TotCE:=0 ;
  ZeroBlanc(PPied) ;
  ZeroBlanc(PPiedTreso) ;
END ;

procedure TFSaisieEff.InitGrid ;
BEGIN
  if Action<>taCreat then Exit ;
  GS.Enabled:=True ;
  DefautLigne(GS.RowCount-1,True) ;
  GS.Col:=SA_Gen ;
  GS.Row:=1 ;
END ;

procedure TFSaisieEff.AttribLeTitre ;
BEGIN
  Case ModeSaisie Of
    OnEffet : Caption:=HTitres.Mess[31] ; // Saisie des traites en retour d'acceptation
    OnChq   : Caption:=HTitres.Mess[32] ; // Saisie des chèques
    OnCB    : Caption:=HTitres.Mess[33] ; // Saisie des cartes bleues
    OnBqe   : Caption:=HTitres.Mess[30] ; // Saisie des écritures de règlement
  END ;
  UpdateCaption(Self) ;
END ;

procedure TFSaisieEff.AttribNumeroDef ;
Var Facturier : String3 ;
    Lig,NumAxe,NumV : integer ;
    O          : TOBM ;
    DD         : TDateTime ;
BEGIN
  if Action<>taCreat then Exit ;
  Facturier:=SAJAL.CompteurNormal ;
  DD:=StrToDate(E_DATECOMPTABLE.Text) ;
  SetIncNum(EcrGen,Facturier,NumPieceInt,DD) ;
  H_NUMEROPIECE.Visible:=True ;
  H_NUMEROPIECE_.Visible:=False ;
  E_NUMEROPIECE.Caption:=IntToStr(NumPieceInt) ;

  {Attribution aux objets du nouveau numéro}
	try
    for Lig:=1 to GS.RowCount-1 do BEGIN
      GS.Cells[SA_NumP,Lig]:=IntToStr(NumPieceInt) ;
      O:=GetO(GS,Lig) ;
      if O=Nil then Break ;
      O.PutMvt('E_NUMEROPIECE',NumPieceInt) ;
      for NumAxe:=1 to O.Detail.Count do
        for NumV:=0 to O.Detail[NumAxe-1].Detail.Count-1 do
          TOBM(O.Detail[NumAxe-1].Detail[NumV]).PutMvt('Y_NUMEROPIECE',NumPieceInt) ;
    end;
  except
 		on e:exception do MessageAlerte( 'Erreur lors de l''affectation du numéro définitif'+#10#13+e.Message );
  end;
end;

{==========================================================================================}
{================================== Scenario de saisie ====================================}
{==========================================================================================}
procedure TFSaisieEff.LettrageEnSaisie ;
{$IFNDEF GCGC}
Var X : RMVT ;
    Q : TQuery ;
    Nb  : integer ;
    St,StWhere : String ;
{$ENDIF}    
BEGIN
{$IFNDEF GCGC}
  If SAJAL=Nil Then Exit ;
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
  StWhere:='E_GENERAL<>"'+SAJAL.TRESO.CPT+'" ' ;
  St:='Select Count(*) from ECRITURE Where '+WhereEcriture(tsGene,X,False)+
      ' AND E_ECHE="X" AND E_NUMECHE>0 AND (E_ETATLETTRAGE="AL" OR E_ETATLETTRAGE="PL") '+
      ' AND '+StWhere ;
  Q:=OpenSQL(St,True) ;
  if Not Q.EOF then BEGIN
    Nb:=Q.Fields[0].AsInteger ;
    if Nb>0 then BEGIN
      // Voulez-vous lettrer vos échéances ?
      if HPiece.Execute(24,caption,'')=mrYes then LettrerEnSaisie(X,Nb,StWhere) ;
    END ;
  END ;
  Ferme(Q) ;
{$ENDIF}
END ;

procedure TFSaisieEff.ChargeScenario ;
Var Q        : TQuery ;
    STVA,SQL : String ;
    StComp   : String[10] ;
BEGIN
  Scenario.PutDefautDivers ;
  OkScenario:=False ;
  DejaRentre:=False ;
  Entete.Free ;
  Entete:=TOBM.Create(EcrGen,'',True) ;
  GestionRIB:='...' ;
  MemoComp.Free ;
  MemoComp:= HTStringList.Create ;
  if PasModif then Exit ;
  if M.ANouveau then Exit ;
  if ((E_JOURNAL.Value='') or (E_NATUREPIECE.Value='')) then Exit ;
  SQL:='Select * from SUIVCPTA Where SC_JOURNAL="'+E_JOURNAL.Value+'" AND SC_NATUREPIECE="'+E_NATUREPIECE.Value
      +'" AND (SC_USERGRP="" OR SC_USERGRP="'+V_PGI.Groupe+'" OR SC_USERGRP="...")' ;
  BeginTrans ;
  Q:=OpenSQL(SQL,True) ;
  if Not Q.EOF then BEGIN
    Scenario.ChargeMvt(Q) ;
    if Not TMemoField(Q.FindField('SC_ATTRCOMP')).IsNull then
{$IFDEF EAGLCLIENT}
      MemoComp.Text := (TMemoField(Q.FindField('SC_ATTRCOMP')).AsString) ; // A FAIRE   Est-ce OK ?
{$ELSE}
      MemoComp.Assign(TMemoField(Q.FindField('SC_ATTRCOMP'))) ;
{$ENDIF}
    OkScenario:=True ;
  END ;
  Ferme(Q) ;
  CommitTrans ;
  if OkScenario then BEGIN
    StComp:=Scenario2Comp(Scenario) ;
    Entete.PutMvt('E_ETAT',StComp) ;
    STVA:=Scenario.GetMvt('SC_CONTROLETVA') ;
    GestionRIB:=Scenario.GetMvt('SC_RIB') ;
    if Scenario.GetMvt('SC_BROUILLARD')='X' then BEGIN
      NeedEdition:=True ;
      EtatSaisie:='' ;
    END ;
    if Scenario.GetMvt('SC_DOCUMENT')<>'' then BEGIN
      NeedEdition:=True ;
      EtatSaisie:=Scenario.GetMvt('SC_DOCUMENT') ;
    END ;
    END
  else BEGIN
    GestionRIB:='...' ;
  END ;
END ;

{==========================================================================================}
{============================== Traitements liés à l'entête ===============================}
{==========================================================================================}

procedure TFSaisieEff.E_MODEPAIEExit(Sender: TObject);
begin
  OkPourLigne ;
  If (GS.RowCount>2) Then ModifModePaie ;
end;

procedure TFSaisieEff.E_DEVISEExit(Sender: TObject);
begin
  if csDestroying in ComponentState then Exit ;
  if ((ModeOppose) and (Action=taCreat) and (Not VH^.TenueEuro)) then if DatePiece<V_PGI.DateDebuteuro then BEGIN
    // Vous ne pouvez pas saisir en Euro sur une date antérieure à celle d'entrée en vigueur.
    HPiece.Execute(30,Caption,'') ;
    E_DEVISE.Value:=V_PGI.DevisePivot ;
    if E_DEVISE.CanFocus then E_DEVISE.SetFocus ;
    Exit ;
  END ;
  OkPourLigne ;
end;

function TFSaisieEff.OkPourLigne : boolean ;
Var
  DateCpt : TDateTime ;
begin
  Result := True;
  If Action<>taCreat Then Exit;
  if E_MODEPAIE.Text='' then begin
    if (Screen.ActiveControl.Name)='GS' then begin
      // Vous devez indiquer un mode de paiement.
      HDiv.Execute(8,caption,'') ;
      if E_MODEPAIE.CanFocus then E_MODEPAIE.SetFocus ;
      Screen.Cursor:=SyncrDefault ;
      Result := False;
    end;
    end
  Else if E_TYPECTR.Text='' Then begin
    If (Screen.ActiveControl.Name)='GS' Then BEGIN
      // Vous devez indiquer un type de contrepartie (Pensez aussi à mettre le journal à jour).
      HDiv.Execute(9,caption,'') ;
      if E_TYPECTR.CanFocus then E_TYPECTR.SetFocus ;
      Screen.Cursor:=SyncrDefault ;
      result:=false ;
    END ;
//    Test GG Pour Saisie sur ctr ventilable
{    END
  Else If (SAJAL<>NIL) AND SAJAL.Treso.Ventilable Then BEGIN
    // Le compte de contrepartie est ventilable. Vous devez utiliser la saisie courante pour saisir sur ce compte.
//    HDiv.Execute(20,caption,'') ;
    If ModeSaisie<>OnBqe Then BEGIN
      // Le compte de contrepartie est ventilable. Vous devez utiliser la saisie courante pour saisir sur ce compte.
      HDiv.Execute(20,caption,'') ;
      E_CONTREPARTIEGEN.SetFocus ;
      END
    Else BEGIN
      // Le compte de trésorerie est ventilable. Vous devez utiliser la saisie courante pour saisir sur ce compte.
      HDiv.Execute(10,caption,'') ;
      E_JOURNAL.SetFocus ;
    END ;
//    E_CONTREPARTIEGEN.SetFocus ;
    Screen.Cursor:=SyncrDefault ;
    result:=false ;}
    END
  Else If (SAJAL<>NIL) And (SAJAL.NATUREJAL='BQE') And (PasOkDevBqe(SAJAL.TRESO.DevBqe,E_DEVISE.Value)) Then BEGIN
    // Cette devise ne correspond pas à la devise du compte de banque. Vous devez la modifier.
    HDiv.Execute(11,caption,'') ;
    if E_DEVISE.CanFocus then E_DEVISE.SetFocus ;
    Screen.Cursor:=SyncrDefault ;
    result:=false ;
  END ;
  If Result Then
    if ((DEV.Code<>V_PGI.DevisePivot) and (Action=taCreat) and (Not Volatile)) then BEGIN
      DateCpt:=StrToDate(E_DATECOMPTABLE.Text) ;
      DEV.Taux:=GetTaux(DEV.Code,DEV.DateTaux,DateCpt) ;
      If ((DEV.DateTaux<>DateCpt) or (PbTaux(DEV,DateCpt))) Then AvertirPbTaux(DEV.Code,DEV.DateTaux,DateCpt) ;
    END ;

  If Result And (ModeSaisie<>OnBqe)Then BEGIN
    If E_CONTREPARTIEGEN.Text='' Then
      If (Screen.ActiveControl.Name)='GS' Then BEGIN
        // Vous devez indiquer un compte de contrepartie.
        HDiv.Execute(15,caption,'') ;
        if E_CONTREPARTIEGEN.CanFocus then E_CONTREPARTIEGEN.SetFocus ;
        Screen.Cursor:=SyncrDefault ;
        result:=false ;
      END ;
  END ;

  If Result And (ModeSaisie<>OnBqe) Then BEGIN
    If (SAJAL<>NIL) AND SAJAL.Treso.Collectif And (SAJAL.Treso.TypCtr in [PiedDC,PiedSolde]) Then BEGIN
      // Le compte de contrepartie est collectif. Vous ne pouvez pas utiliser un mode de contrepartie "Pied".
      HDiv.Execute(16,caption,'') ;
      if E_TYPECTR.CanFocus then E_TYPECTR.SetFocus ;
      Screen.Cursor:=SyncrDefault ;
      result:=false ;
    END ;
  END ;

  If Result And (ModeSaisie<>OnBqe) Then BEGIN
    If (Screen.ActiveControl.Name)='GS' Then BEGIN
      {JP 17/01/08 : FQ 18563 : Calcul du solde en fonction du compte et surtout de l'exercice}
      SAJAL.PutDate(E_DATECOMPTABLE.Text);

      If Not SAJAL.ChargeCompteTreso Then BEGIN
        // Le compte de contrepartie n'existe pas.
        HDiv.Execute(17,caption,'') ;
        if E_CONTREPARTIEGEN.CanFocus then E_CONTREPARTIEGEN.SetFocus ;
        Screen.Cursor:=SyncrDefault ;
        result:=false ;
      END ;
    END ;
  END ;
end ;

Procedure TFSaisieEff.InitNTP ;
BEGIN
  AutoCharged:=FALSE ;
  if GS.RowCount<=2 then BEGIN
    ChargeScenario ;
    ChargeLibelleAuto ;
  END ;
END ;

procedure TFSaisieEff.QuelNExo ;
Var EXO : String ;
BEGIN
  if Action=taCreat then EXO:=QuelExoDT(StrToDate(E_DATECOMPTABLE.Text)) else EXO:=M.EXO ;
  if EXO=VH^.Encours.Code then H_EXERCICE.Caption:='(N)' else
   if EXO=VH^.Suivant.Code then H_EXERCICE.Caption:='(N+1)' else
     H_EXERCICE.Caption:='(N-1)' ;
END ;

procedure TFSaisieEff.InitBoutonValide ;
begin
  If SAJAL=NIL Then Exit ;
  SAJAL.COMPTEAUTOMAT:='' ;
  SAJAL.NbAuto:=0 ;
  If SAJAL.TRESO.TypCtr=AuChoix Then BEGIN
    SAJAL.NbAuto:=1 ;
    SAJAL.COMPTEAUTOMAT:=SAJAL.TRESO.CPT+';' ;
  END ;
end ;

procedure TFSaisieEff.E_JOURNALChange(Sender: TObject);
Var
  Jal : String ;
begin
  Jal:=E_JOURNAL.Value ;
  if Jal='' then BEGIN SAJAL.Free ; SAJAL:=Nil ; Exit; END ;
  if SAJAL<>Nil then if SAJAL.JOURNAL=Jal then Exit;
  if ((VH^.JalAutorises<>'') and (Pos(';'+Jal+';',VH^.JalAutorises)<=0)) then BEGIN
   // Vous n'avez pas le droit de saisir sur ce journal. 
   HPiece.Execute(26,caption,'') ;
   if SAJAL<>Nil then BEGIN SAJAL.Free ; SAJAL:=Nil ; END ;
   E_JOURNAL.Value:='' ;
   if Action=taCreat then BEGIN E_JOURNAL.SetFocus ; GS.Enabled:=False ; END ;
   Exit ;
   END ;
if SAJAL<>Nil then if ((Jal=VH^.JalATP) or (Jal=VH^.JalVTP)) then
   BEGIN
   // Vous ne pouvez pas saisir sur les journaux dédiés aux tiers payeurs.
   HPiece.Execute(33,caption,'') ;
   if SAJAL<>Nil then BEGIN SAJAL.Free ; SAJAL:=Nil ; END ;
   E_JOURNAL.Value:='' ;
   if Action=taCreat then BEGIN E_JOURNAL.SetFocus ; GS.Enabled:=False ; END ;
   Exit ;
   END ;
ChercheJal(Jal,False) ;
if ((Action=taCreat) and (NumPieceInt=0)) then
   BEGIN
   // Vous n'avez pas défini de compteur de numérotation pour ce journal.
   HPiece.Execute(27,caption,'') ;
   if SAJAL<>Nil then BEGIN SAJAL.Free ; SAJAL:=Nil ; END ;
   E_JOURNAL.Value:='' ; E_JOURNAL.SetFocus ; GS.Enabled:=False ;
   Exit ;
   END ;
CalculSoldeCompteT ;
If (SAJAL<>NIL) Then E_CONTREPARTIEGEN.Text:=SAJAL.Treso.Cpt ;
AlimObjetMvt(0,FALSE,1) ; AlimObjetMvt(0,FALSE,2) ;
InitBoutonValide ;
END ;

procedure TFSaisieEff.AvertirPbTaux(Code : String3 ; DateTaux,DateCpt : TDateTime) ;
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
(*
if ((EstMonnaieIN(Code)) and (DateCpt>=V_PGI.DateDebutEuro)) then
   BEGIN
   // La parité fixe est mal renseignée. Voulez-vous la saisir ?
   ii:=HDiv.Execute(23,Caption,'') ;
   // ATTENTION : La parité est incorrecte. Voulez-vous la renseigner ?
   if ii<>mrYes then ii:=HDiv.Execute(24,caption,'') ;
   if ii=mrYes then
      BEGIN
      FicheDevise(Code,taModif,False) ;
      DEV.Taux:=GetTaux(DEV.Code,DEV.DateTaux,DateCpt) ;
      END ;
   END else
   *)
   BEGIN
   if OkTaux1 then ii:=HDiv.Execute(25,caption,'') // ATTENTION : Le taux en cours est de 1. Voulez-vous saisir ce taux dans la table de chancellerie ?
              else ii:=HDiv.Execute(2,caption,''); // Voulez-vous saisir ce taux dans la table de chancellerie ?
   if ii=mrYes then
      BEGIN
      FicheChancel(E_DEVISE.VALUE,True,DateCpt,taCreat,(DateCpt>=V_PGI.DateDebutEuro)) ;
      DEV.Taux:=GetTaux(DEV.Code,DEV.DateTaux,DateCpt) ;
      END ;
   END ;
end ;

Function TFSaisieEff.PbTaux ( DEV : RDevise ; DateCpt : TDateTime ) : boolean ;
Var Code : string3 ;
BEGIN
  Result:=False ;
  Code:=DEV.Code ;
  if ((Code=V_PGI.DevisePivot) or (Code=V_PGI.DeviseFongible)) then Exit ;
  if ((DateCpt<V_PGI.DateDebutEuro) or (Code<>V_PGI.DevisePivot) {(Not EstMonnaieIn(Code))}) then Result:=(DEV.Taux=1)
  //else if EstMonnaieIn(Code) then Result:=(DEV.Taux=V_PGI.TauxEuro) ;
END ;

procedure TFSaisieEff.E_DEVISEChange(Sender: TObject);
Var DateCpt : TDateTime ;
begin
if E_DEVISE.Value='' then BEGIN FillChar(DEV,Sizeof(DEV),#0) ; Exit ; END ;
DEV.Code:=E_DEVISE.Value ; GetInfosDevise(DEV) ;
if ModeOppose then DEV.Decimale:=V_PGI.OkDecE ;
DecDev:=DEV.Decimale ;
Case Action of
   taCreat   : BEGIN
               if ((DEV.Code=V_PGI.DevisePivot) {or (EstMonnaieIN(DEV.Code))}) then Volatile:=False ;
               DateCpt:=StrToDate(E_DATECOMPTABLE.Text) ;
               if Not Volatile then
                  BEGIN
                  DEV.Taux:=GetTaux(DEV.Code,DEV.DateTaux,DateCpt) ;
                  If Not M.Treso Then
                     if (ModeForce=tmNormal) and (E_DEVISE.Enabled) and ((DEV.DateTaux<>DateCpt) or (PbTaux(DEV,DateCpt))) then AvertirPbTaux(DEV.Code,DEV.DateTaux,DateCpt) ;
                  END ;
               BSaisTaux.Enabled:=((GS.RowCount<=2) and
                                   (E_DEVISE.Value<>V_PGI.DevisePivot) and
                                   (Action=taCreat)) ;
//               DEV.Taux:=GetTaux(DEV.Code,DEV.DateTaux,DateCpt) ;
               END ;
   END ;
ChangeFormatDevise(Self,DECDEV,DEV.Symbole) ;
ModeDevise:=EnDevise(DEV.Code) ;
BSwapPivot.Enabled:=((ModeDevise) or (ModeForce=tmPivot)) ;
end;

procedure TFSaisieEff.H_JOURNALDblClick(Sender: TObject);
begin ClickJournal ; end;

{procedure TFSaisieEff.InitTCTR ;
Begin
If SAJAL=NIL then Exit ;
SAJAL.TRESO.STypCtr:='LIG' ;
SAJAL.TRESO.TypCtr:=StrToTypCtr(SAJAL.TRESO.STypCtr) ;
InitBoutonValide ;
End ;}

procedure TFSaisieEff.E_DATECOMPTABLEEnter(Sender: TObject);
begin
RentreDate:=True ;
end;

procedure TFSaisieEff.E_DATECOMPTABLEExit(Sender: TObject);
Var Err,i : integer ;
    DD    : TDateTime ;
    StD   : String ;
begin
  if csDestroying in ComponentState then Exit ;
  Err:=ControleDate(E_DATECOMPTABLE.Text) ;
  if Err>0 then BEGIN
    // ?? 
    HPiece.Execute(15+Err,caption,'') ;
    E_DATECOMPTABLE.SetFocus ;
    Exit ;
  END ;
  if ((Action=taCreat) and (RevisionActive(StrToDate(E_DATECOMPTABLE.Text)))) then BEGIN
    E_DATECOMPTABLE.SetFocus ;
    Exit ;
  END ;
  if ((Action=taCreat) and (ModeOppose) and (Not VH^.TenueEuro) and (StrToDate(E_DATECOMPTABLE.Text)<V_PGI.DateDebutEuro)) then BEGIN
    // Vous ne pouvez pas saisir en Euro sur une date antérieure à celle d'entrée en vigueur.
    HPiece.Execute(30,Caption,'') ;
    E_DATECOMPTABLE.SetFocus ;
    Exit ;
  END ;
  QuelNExo;
  DatePiece:=StrToDate(E_DATECOMPTABLE.Text) ;
  OkPourLigne ;
  DD:=StrToDate(E_DATECOMPTABLE.Text) ;
  StD:=DateToJour(DD) ;
  for i:=1 to GS.RowCount-1 do
    GS.Cells[SA_DateC,i]:=StD ;

  {JP 17/01/08 : FQ 18563 : Calcul du solde en fonction du compte et surtout de l'exercice}
  SAJAL.PutDate(E_DATECOMPTABLE.Text);
  SAJAL.ChargeCompteTreso ;

  // Recalcul du pied pour solde des comptes à dates FQ13246 SBO 03/08/2005
  AffichePied ;
end;

{==========================================================================================}
{=================================== ANALYTIQUES ==========================================}
{==========================================================================================}
Function TFSaisieEff.AOuvrirAnal ( Lig : integer ; Auto,CtrLigne,NewLigne : boolean ) : Boolean ;
Var CGen  : TGGeneral ;
    OBM   : TOBM;
    XD : Double ;
BEGIN
  Result:=False ;
  CGen:=GetGGeneral(GS,Lig) ;
  if CGen=Nil then exit ;
  if not Ventilable(CGen,0) then Exit ;
  if Auto then BEGIN
    OBM:=GetO(GS,Lig);
    if OBM=Nil then Exit ;
    XD:=MontantLigne(GS,Lig,tsmDevise) ;
    if (GetTotalVentilDev(OBM)=XD) then Exit ;
    if GetVentilGeneral(OBM)=CGen.General then BEGIN
      RecalculProrataAnalNEW('Y',OBM,0,DEV) ;
      Exit ;
    END ;
  END ;
  Result:=True ;
END ;

Procedure TFSaisieEff.GereAnal ( Lig : integer ; Auto,Scena,CtrLigne,NewLigne : Boolean ) ;
var  bOuvreAnal : Boolean;
     OBM        : TOBM;
     lNumAxe    : integer ;
BEGIN
  if AOuvrirAnal(Lig,Auto,CtrLigne,NewLigne) then
    begin
    lNumAxe := 0 ;
    if Auto then
      begin
      OBM := GetO(GS, Lig);
      if okScenario
        then bOuvreAnal := CAOuvrirVentil( OBM, Scenario, FInfo, lNumAxe )
        else bOuvreAnal := CAOuvrirVentil( OBM, nil, FInfo, lNumAxe ) ;
      end
    else bOuvreAnal := True ;
    // Ouvre la fenêtre d'analytique
    if bOuvreAnal then
      OuvreAnal(Lig,Scena,CtrLigne,NewLigne, lNumAxe) ;
    end ;
END ;

Procedure MajLL(LL : HTStringList ; Section : String ; Taux : Double) ;
Var i : Integer ;
    OkAdd : Boolean ;
    St,Sect,Staux : String ;
BEGIN
OkAdd:=TRUE ;
For i:=0 To LL.Count-1 Do
  BEGIN
  St:=LL[i] ; Sect:='' ; STaux:='' ;
  If St<>'' Then Sect:=ReadTokenSt(St) ; If St<>'' Then STaux:=ReadTokenSt(St) ;
  If Sect=Section Then
    BEGIN
    OkAdd:=FALSE ;
    LL[i]:=Sect+';'+formatfloat('0000.0000',Taux)+';' ;
    Break ;
    END ;
  END ;
If OkAdd Then LL.Add(Section+';'+formatfloat('0000.0000',Taux)+';') ;
END ;

Procedure TFSaisieEff.RecupLesVentilsCtr(NumAxe,Lig : Integer ; LL : HTStringList) ;
Var i,j : Integer ;
    CGen  : TGGeneral ;
    OBM : TOBM;
    Section : String ;
    Taux : Double ;
BEGIN
// A FAIRE Voir si OK
  If SAJAl=Nil Then Exit ;
  For i:=Lig-1 DownTo 1 Do BEGIN
    OBM := GetO(GS,i);
    CGen:=GetGGeneral(GS,i) ;
    If CGen=NIL Then Continue ;
    If CGen.General=SAJAL.Treso.Cpt Then Break ;
    If Not CGen.Ventilable[NumAxe] Then Break ;
    for j:=0 to OBM.Detail[i-1].Detail.Count-1 do BEGIN
      Section := TOBM(OBM.Detail[i-1].Detail[j]).GetMvt('Y_SECTION') ;
      Taux	  := TOBM(OBM.Detail[i-1].Detail[j]).GetMvt('Y_POURCENTAGE')  ;
      MajLL(LL,Section,Taux) ;
    END ;
  END ;
END ;

Procedure TFSaisieEff.OuvreAnal ( Lig : integer ; Scena,CtrLigne,NewLigne : boolean; vNumAxe : integer ) ;
Var Ax            : String ;
    ArgAna 	  : ARG_ANA;
    OBM 	  : TOBM;
    i             : Integer ;
    LL            : T5LL ;
BEGIN
	// Mise en place des arguments de lancement de la saisie analytique
  ArgAna.GuideA	         := '' ;
  ArgAna.DernVentilType  := DernVentilType ;
  ArgAna.ControleBudget  := False;
  ArgAna.ModeConf	 := ModeConf ;
  ArgAna.CC		 := GetGGeneral(GS,Lig) ;
  ArgAna.DEV		 := DEV ;
  ArgAna.NumLigneDecal   := 0;
  ArgAna.QuelEcr         := EcrGen;
  if DernVentilType<>'' then
    ArgAna.DernVentilType:=DernVentilType ;
  OBM := GetO(GS,Lig);

{$IFDEF OLDVENTIL}
  AlloueAxe( OBM ) ;// SBO 25/01/2006
  ArgAna.VerifVentil := False;
  ArgAna.Action      := Action;
  vNumAxe := 0;
{$ELSE  OLDVENTIL}
{  if VentilationExiste( OBM )
      then ArgAna.Action := taModif
      else
}  ArgAna.Action := taCreat ;
  ArgAna.Verifventil	:= True ;
  if vNumAxe = 0 then
{$ENDIF OLDVENTIL}
    if ((OkScenario) and (Scena) and (Scenario.GetMvt('SC_OUVREANAL') <> 'X')) then
    begin
      Ax := Trim(Scenario.GetMvt('SC_NUMAXE'));
      if Length(Ax) < 2 then Exit;
      vNumAxe := Ord(Ax[2]) - 48;
    end;

  ArgAna.NumGeneAxe     := vNumAxe;
  ArgAna.VerifQte       := ((OkScenario) and (Action<>taConsult) and (Scenario.GetValue('SC_CONTROLEQTE')='X'));
  ArgAna.Info           := FInfo ;
  ArgAna.MessageCompta  := nil ;

  For i:=1 To maxAxe Do LL[i]:=NIL;
  If CtrLigne Then  BEGIN
    If ArgAna.CC<>NIL Then BEGIN
      If NewLigne Then BEGIN
        For i:=1 To maxAxe Do If TGGeneral(ArgAna.CC).Ventilable[i] Then BEGIN
          LL[i]:=HTStringList.Create ;
          RecupLesVentilsCtr(i,Lig,LL[i]) ;
        END ;
      END ;
    END ;
    eSaisieAnalEff(TOB(OBM), LL, ArgAna);
    END
Else BEGIN
  eSaisieAnalEff(TOB(OBM), LAnaEff, ArgAna);
END ;

For i:=1 To maxAxe Do BEGIN
  If LL[i]<>NIL Then BEGIN
    LL[i].Clear ;
    LL[i].Free ;
  END ;
  LL[i]:=NIL ;
END ;

if ArgAna.DernVentilType<>'' then DernVentilType:=ArgAna.DernVentilType ;

  if ArgAna.ModeConf<>ModeConf then BEGIN
    if ArgAna.ModeConf>ModeConf then ModeConf:=ArgAna.ModeConf ;
    CalculDebitCredit ;
  END ;
{
  if Action=taConsult then Exit ;

  if ModeDevise then BEGIN
    if ((XD=ArgAna.TotalDevise) and (XE=ArgAna.TotalEuro)) then Exit ; V:=ArgAna.TotalDevise ;
    END
  else BEGIN
    if ((XP=ArgAna.TotalEcriture) and (XE=ArgAna.TotalEuro)) then Exit ; V:=ArgAna.TotalEcriture ;
  END ;
  if OBA.Sens=1 then C:=SA_Debit else C:=SA_Credit ;
  GS.Cells[C,Lig]:=StrS(V,DECDEV) ;
  FormatMontant(GS,C,Lig,DECDEV) ;
  TraiteMontant(Lig,True) ;
}
END ;

{==========================================================================================}
{=================================== ECHEANCES ============================================}
{==========================================================================================}
Procedure TFSaisieEff.GereEche ( Lig : integer ; OuvreAuto,RempliAuto : Boolean ) ;
Var Cpte : String17 ;
    MR   : String3 ;
    t    : TLettre ;
    FromLigne : Boolean ;
BEGIN
  FromLigne:=((OuvreAuto) and (Not RempliAuto)) ;
  if AOuvrirEche(Lig,Cpte,MR,OuvreAuto,RempliAuto,t) then BEGIN
    if ((OkScenario) and (FromLigne)) then
      if Scenario.GetMvt('SC_OUVREECHE')<>'X' then RempliAuto:=True ;
    OuvreEche(Lig,Cpte,MR,RempliAuto,t,FALSE) ;
  END ;
END ;

Procedure TFSaisieEff.DebutOuvreEche(Lig : Integer ; Cpte : String17 ; MR : String3 ; OMODR : TMOD) ;
BEGIN
  With OMODR.MODR do BEGIN
    TotalAPayerP:=MontantLigne(GS,Lig,tsmPivot) ;
    TotalAPayerD:=MontantLigne(GS,Lig,tsmDevise) ;
    CodeDevise:=DEV.Code ;
    TauxDevise:=DEV.Taux ;
    Quotite:=DEV.Quotite ;
    Decimale:=DECDEV ;
    Symbole:=DEV.Symbole ;
    DateFact:=DateMvt(Lig) ;
    DateFactExt:=DateFact ;
    DateBL:=DateFact     ;
    Aux:=Cpte ;
    ModeInitial:=MR      ;
  END ;
END ;

Procedure TFSaisieEff.FinOuvreEche (Lig : Integer ; OMODR : TMOD ; t : TLettre) ;
var i    : Integer;
    CGen : TGGeneral ;
begin

  CGen:=GetGGeneral(GS,Lig) ;
  if CGen=Nil then Exit ;

  With OMODR.MODR do
    for i:=1 to MaxEche do
      With TabEche[i] do BEGIN
        Couverture:=0 ;
        CouvertureDev:=0 ;
        CodeLettre:='' ;
        LettrageDev:='-' ;
        DatePaquetMax:=DateMvt(Lig) ;
        DatePaquetMin:=DateMvt(Lig) ;
        DateRelance:=IDate1900 ;
        NiveauRelance:=0 ;

        // Cas spécial des comptes divers lettrables FQ 16136 SBO 09/08/2005
        if ( CGen.Lettrable and ( CGen.NatureGene = 'DIV' ) ) or ( t = MultiEche )
          then EtatLettrage:='AL'
          else EtatLettrage:='RI' ;

      END ;
END ;

Procedure TFSaisieEff.AlimEcheVersObm(Lig : Integer ; ModePaie : String3 ; DateEche,DateValeur : TDateTime) ;
Var OBM : TOBM ;
BEGIN
  OBM:=GetO(GS,Lig) ;
  If OBM<>NIL Then BEGIN
    OBM.PutMvt('E_MODEPAIE',ModePaie) ;
    OBM.PutMvt('E_DATEECHEANCE',DateEche) ;
    OBM.PutMvt('E_DATEVALEUR',DateValeur) ;
//    OBM.PutMvt('E_NUMTRAITECHQ',NumTraiteChq) ;
  END ;
END ;

Procedure TFSaisieEff.OuvreEche ( Lig : integer ; Cpte : String17 ; MR : String3 ; RempliAuto : boolean ; t : TLettre ; LigTresoEnPied : Boolean) ;
Var OMODR  : TMOD ;
    OkInit : Boolean ;
    X      : T_MONOECH ;
    O      : TOBM ;
BEGIN
  OMODR:=TMOD(GS.Objects[SA_NumP,Lig]) ;
  if OMODR=Nil then Exit ;
  O:=GetO(GS,Lig) ;
  With OMODR.MODR do BEGIN
    if Self.Action=taConsult then Action:=taConsult
    else BEGIN
      if Not ((TotalAPayerP=0) or (NbEche=0)) then Action:=taModif
                                              else Action:=taCreat;
    END ;
    OkInit:=(Action=taCreat) ;
    DebutOuvreEche(Lig,Cpte,MR,OMODR) ;
    if Action=taCreat then BEGIN
      If DateEcheDefaut=0 Then TabEche[1].DateEche:=DateMvt(Lig)
                          Else TabEche[1].DateEche:=DateEcheDefaut ;
      TabEche[1].DateValeur:=CalculDateValeur(DateMvt(Lig),t,SAJAL.Treso.CPT) ;
    END ;
    If SA_ModeP>=0 Then TabEche[1].ModePaie:=GS.CellValues[SA_MODEP,Lig] ;
    If SA_DateEche>=0 Then TabEche[1].DateEche:=StrtoDate(GS.Cells[SA_DAteEche,Lig]) ;
    If Sa_NumTraChq>=0 Then O.PutMvt('E_NUMTRAITECHQ',GS.Cells[SA_NumTraChq,Lig])
  END ;

  if RempliAuto then BEGIN
    With OMODR.MODR do BEGIN
      NbEche:=1 ;
      TabEche[1].Pourc:=100.0 ;
      TabEche[1].MontantP:=TotalAPayerP ;
      TabEche[1].MontantD:=TotalAPayerD ;
    END ;
    END
  else BEGIN
    Case t of
      MultiEche : With OMODR.MODR do BEGIN
                    X.DateEche:=TabEche[1].DateEche ;
                    X.ModePaie:=TabEche[1].ModePaie ;
                    X.DateValeur:=TabEche[1].DateValeur ;
                    X.DateMvt:=DateMvt(Lig) ;
                    X.Cat:='' ;
                    X.Treso:=True ;
                    X.OkInit:=OkInit ;
                    X.OkVal:=True ;
                    X.NumTraiteCHQ:='' ;
{$IFNDEF SPEC350}
                    if O<>Nil then X.NumTraiteCHQ:=O.GetMvt('E_NUMTRAITECHQ') ;
{$ENDIF}
                    if SaisirMonoEcheance(X,PlusModePaie) then BEGIN
                      TabEche[1].DateEche:=X.DateEche ;
                      TabEche[1].ModePaie:=X.ModePaie ;
                      TabEche[1].DateValeur:=X.DateValeur ;
                      NbEche:=1 ; TabEche[1].Pourc:=100.0 ;
                      TabEche[1].MontantP:=TotalAPayerP ;
                      TabEche[1].MontantD:=TotalAPayerD ;
                      DateEcheDefaut:=X.DateEche ;
{$IFNDEF SPEC350}
                      if O<>Nil then O.PutMvt('E_NUMTRAITECHQ',X.NumTraiteCHQ) ;
{$ENDIF}
                    END ;
                  END ;
      MonoEche  : With OMODR.MODR do BEGIN
                    X.DateEche:=TabEche[1].DateEche ;
                    X.ModePaie:=TabEche[1].ModePaie ;
                    X.DateValeur:=TabEche[1].DateValeur ;
                    X.Action:=Action ;
                    X.Cat:='' ;
                    X.Treso:=True ;
                    X.OkInit:=OkInit ;
                    X.DateMvt:=DateMvt(Lig) ;
                    X.NumTraiteCHQ:='' ;
                    if O<>Nil then X.NumTraiteCHQ:=O.GetMvt('E_NUMTRAITECHQ') ;
                    if LigTresoEnPied Or SaisirMonoEcheance(X,PlusModePaie) then BEGIN
                      TabEche[1].DateEche:=X.DateEche ;
                      TabEche[1].ModePaie:=X.ModePaie ;
                      TabEche[1].DateValeur:=X.DateValeur ;
                      NbEche:=1 ; TabEche[1].Pourc:=100.0 ;
                      TabEche[1].MontantP:=TotalAPayerP ;
                      TabEche[1].MontantD:=TotalAPayerD ;
                      DateEcheDefaut:=X.DateEche ;
                      if O<>Nil then O.PutMvt('E_NUMTRAITECHQ',X.NumTraiteCHQ) ;
                    END ;
                  END ;
    END ;
  END ;

  If OkInit Then FinOuvreEche(Lig,OMODR,t) ;
  AlimEcheVersObm(Lig,OMODR.MODR.TabEche[1].ModePaie,OMODR.MODR.TabEche[1].DateEche, OMODR.MODR.TabEche[1].DateValeur) ;

  With OMODR.MODR do BEGIN
    If SA_ModeP>=0 Then GS.CellValues[SA_MODEP,Lig]:=TabEche[1].ModePaie ;
    If SA_DateEche>=0 Then GS.Cells[SA_DAteEche,Lig]:=DateToStr(TabEche[1].DateEche) ;
    If Sa_NumTraChq>=0 Then GS.Cells[SA_NumTraChq,Lig]:=O.GetMvt('E_NUMTRAITECHQ') ;
  END ;
END ;

Function TFSaisieEff.AOuvrirEche ( Lig : integer ; Var Cpte : String17 ; Var MR : String3 ; Var OuvreAuto,RempliAuto : boolean ; Var t : TLettre) : Boolean ;
Var CGen     : TGGeneral ;
    CAux     : TGTiers ;
    OMODR    : TMOD ;
    XD,XP : Double ;
BEGIN
  Result:=False ;
  Cpte:='' ;
  MR:='' ;
  t:=NonEche ;
  CGen:=GetGGeneral(GS,Lig) ;
  if CGen=Nil then exit ;
  CAux:=GetGTiers(GS,Lig) ;
  if CGen.Collectif then BEGIN
    if CAux=Nil then Exit ;
    if Not CAux.Lettrable then Exit ;
    Cpte:=CAux.Auxi ;
    MR:=CAux.ModeRegle ;
    t:=MultiEche ;
    END
  else BEGIN
    t:=Lettrable(CGen) ;
    if t=NonEche then Exit ;
    Cpte:=CGen.General ;
    MR:=CGen.ModeRegle ;
    // Ajout test pour comptes divers lettrables... FQ 16136 SBO 09/08/2005
    if CGen.NatureGene = 'DIV' then
      begin
      OuvreAuto  := False ;
      RempliAuto := True ;
      MR         := GetParamSocSecur('SO_GCMODEREGLEDEFAUT','') ;
      end ;
  END ;

  if OuvreAuto then BEGIN
    OMODR:=TMOD(GS.Objects[SA_NumP,Lig]) ;
    if OMODR=Nil then Exit ;

    XD:=MontantLigne(GS,Lig,tsmDevise) ;
    XP:=MontantLigne(GS,Lig,tsmPivot) ;
    if ((OMODR.MODR.TotalAPayerD=XD) and (OMODR.ModR.TotalAPayerP=XP)) then Exit ;
    if OMODR.MODR.Aux=Cpte then BEGIN
      RecalculProrataEche(OMODR.MODR,XP,XD) ;
      Exit ;
    END ;
  END ;
  Result:=True ;
END ;

{==========================================================================================}
{=================================== Contrôles ============================================}
{==========================================================================================}
procedure TFSaisieEff.ControleLignes;
Var Lig : integer ;
BEGIN
for Lig:=1 to GS.RowCount-1 do if Not EstRempli(GS,Lig) then DefautLigne(Lig,True) ;
END ;

procedure TFSaisieEff.AjusteLigne ( Lig : integer ) ;
Var Cpte              : String17 ;
    MR                : String3 ;
    t                 : TLettre ;
    Ouv,Remp          : boolean ;
BEGIN
AlimObjetMvt(Lig,True,0) ; Remp:=False ; Ouv:=True ;
if AOuvrirEche(Lig,Cpte,MR,Ouv,Remp,t) then ;
if AOuvrirAnal(Lig,True,FALSE,FALSE) then ;
END ;

procedure TFSaisieEff.ErreurSaisie ( Err : integer ) ;
BEGIN
if Err<100 then
   BEGIN
   // ??
   HLigne.Execute(Err-1,caption,'') ;
   END else if Err<200 then
   BEGIN
   // ??
   HPiece.Execute(Err-101,caption,'') ;
   END ;
END ;

Function TFSaisieEff.PossibleRecupNum : Boolean ;
Var MM : String17 ;
    Facturier : String3 ;
    DD : TDateTime ;
BEGIN
Result:=True ;
if H_NUMEROPIECE_.Visible then Exit ;
DD:=StrToDate(E_DATECOMPTABLE.Text) ;
Facturier:=SAJAL.CompteurNormal ;
if GetNum(EcrGen,Facturier,MM,DD)=NumPieceInt+1 then Exit ;
Result:=False ;
END ;

Function TFSaisieEff.PasModif : Boolean ;
BEGIN
PasModif:=((Action=taConsult) or (ModeForce<>tmNormal)) ;
END ;

Function TFSaisieEff.LigneCorrecte ( Lig : integer ; Totale,Alim,CtrlCptTreso : boolean ) : Boolean ;
Var CGen  : TGGeneral ;
    CAux  : TGTiers ;
    Err,NumA,NumL : integer ;
    OMOD  : TMOD ;
    OBM   : TOBM ;
    OBA   : TOBM ;
    TotD,TotP : Double ;
    Sens,ii   : byte ;
    Col       : integer ;
    IsCptTreso : Boolean ;
    O : TOBM ;
    RIB : String ;
BEGIN
if PasModif then BEGIN Result:=True ; Exit ; END ;
  Err:=0 ;
  Sens:=1 ;
  if ValC(GS,Lig)<>0 then Sens:=2 ;
  CGen:=GetGGeneral(GS,Lig) ;
  Col:=-1 ;
  if ((Err=0) and (CGen=Nil)) then BEGIN Err:=1 ; Col:=SA_Gen ; END ;
  if ((Err=0) and (CGen.General<>GS.Cells[SA_Gen,Lig])) then BEGIN Err:=1 ; Col:=SA_Gen ; END ;

  CAux:=GetGTiers(GS,Lig) ;
  if ((Err=0) and (Not CGen.Collectif) and (CAux<>Nil)) then BEGIN Err:=2 ; Col:=SA_Gen ; END ;
  if ((Err=0) and (CAux<>Nil)) then if CAux.Auxi<>GS.Cells[SA_Aux,Lig] then BEGIN
    if GS.Cells[SA_Aux,Lig]<>'' then Err:=2
                                else Err:=3 ;
    Col:=SA_Aux ;
  END ;
  if ((Err=0) and (CGen.Collectif) and (CAux=Nil)) then BEGIN Err:=3 ; Col:=SA_Aux ; END ;
  if Err=0 then BEGIN
    ii:=EstInterdit(SAJAL.CompteInterdit,CGen.General,Sens) ;
    If CtrlCptTreso Then BEGIN
      IsCptTreso:=(Cgen<>NIL) And (CGen.General=SAJAL.TRESO.Cpt) ;
//      If (Lig Mod 2<>0) And IsCptTreso Then Err:=30 ;
      Case SAJAL.TRESO.TYPCTR Of
        Ligne : If (Lig Mod 2<>0) And IsCptTreso Then If ModeSaisie<>OnBqe Then Err:=30 Else Err:=22;
        PiedSolde,PiedDC : If isCptTreso Then Err:=22 ;
      END ;
    END ;
    Case ii of 1 : Err:=18 ; 2 : Err:=19 ; 3 : Err:=4 ; END ;
    if Err>0 then Col:=SA_Gen ;
  END ;

  if ((Err=0) and (Not M.ANouveau) and (EstOuvreBil(CGen.General))) then Err:=20 ;
  if ((Err=0) and (M.ANouveau) and (EstChaPro(CGen.NatureGene))) then Err:=21 ;

  if ((Err=0) and (ValD(GS,Lig)=0) and (ValC(GS,Lig)=0)) then BEGIN
    Err:=5 ;
    if QuelSens(E_NATUREPIECE.Value,CGen.NatureGene,CGen.Sens)=1 then Col:=SA_Debit
                                                                 else Col:=SA_Credit ;
  END ;

  if ((Err=0) and (ModeDevise) and (ValeurPivot(GS.Cells[SA_Debit,Lig],DEV.Taux,Dev.Quotite)=0) and
     (ValeurPivot(GS.Cells[SA_Credit,Lig],DEV.Taux,Dev.Quotite)=0)) then BEGIN Err:=6 ; if Sens=1 then Col:=SA_Debit else Col:=SA_Credit ; END ;

  if ((Err=0) and (MontantLigne(GS,Lig,tsmDevise)<>0) and (MontantLigne(GS,Lig,tsmPivot)=0)) then BEGIN Err:=6 ; if Sens=1 then Col:=SA_Debit else Col:=SA_Credit ; END ;

  if ((Err=0) And (CGen.Collectif) and (Not NatureGenAuxOk(CGen.NatureGene,CAux.NatureAux))) Then BEGIN Err:=7 ; Col:=SA_Aux ; END ;

  if ((Err=0) and (not NaturePieceCompteOk(E_NATUREPIECE.Value,CGen.NatureGene))) then BEGIN Err:=8 ; Col:=SA_Gen ; END ;

  if ((Err=0) and (Not VH^.MontantNegatif)) then BEGIN
    if ((ValD(GS,Lig)<0) or (ValC(GS,Lig)<0)) then BEGIN Err:=9 ; Col:=SA_Debit ; if Sens=2 then Col:=SA_Credit ; END ;
  END ;

  if Err=0 then BEGIN
    if ((Abs(MontantLigne(GS,Lig,tsmPivot))<VH^.GrpMontantMin) or (Abs(MontantLigne(GS,Lig,tsmPivot))>VH^.GrpMontantMax)) then BEGIN Err:=23 ; Col:=SA_Debit ; if Sens=2 then Col:=SA_Credit ; END ;
  END ;

  if ((Err=0) and (Totale)) then BEGIN
    if TMOD(GS.Objects[SA_NumP,Lig])<>Nil then BEGIN
      OMOD:=TMOD(GS.Objects[SA_NumP,Lig]) ;
      if TotDifferent(MontantLigne(GS,Lig,tsmDevise),OMOD.MODR.TotalAPayerD) then BEGIN Err:=10 ; Col:=SA_Debit ; if Sens=2 then Col:=SA_Credit ; END ;
    END ;

    if ((Err=0) and (GetO(GS,Lig).Detail.Count>0)) then BEGIN
      OBM := GetO(GS,Lig) ;
      if Err=0 then
        for NumA:=1 to Min(MaxAxe , OBM.Detail.Count) do begin
          if Ventilable(CGen,NumA) then BEGIN
            if OBM.Detail[NumA-1].Detail.Count=0 then Err:=11+NumA
            else BEGIN
              TotD:=0 ;
              TotP:=0 ;
              for NumL:=0 to OBM.Detail[NumA-1].Detail.Count-1 do BEGIN
                OBA := TOBM(OBM.Detail[NumA-1].Detail[NumL]);
                TotD:=TotD + GetMontantDev(OBA) ;
                TotP:=TotP + GetMontant(OBA) ;
              END ;
              if TotDifferent(TotD,MontantLigne(GS,Lig,tsmDevise)) then BEGIN Err:=11+NumA ; Break ; END ;
              if TotDifferent(TotP,MontantLigne(GS,Lig,tsmPivot)) then BEGIN Err:=11+NumA ; Break ; END ;
            END ;
            if Err<>0 then BEGIN Col:=SA_Debit ; if Sens=2 then Col:=SA_Credit ; END ;
          END ;
        END ;
    END ;
  end;

  If (Err=0) And (CGen.General<>SAJAL.TRESO.Cpt) And (ModeSaisie In [OnEffet,OnBqe]) Then BEGIN
    O:=GetO(GS,Lig) ;
    RIB:='---' ;
    If SA_RIB=-1 Then BEGIN If O<>NIL Then RIB:=O.GetMvt('E_RIB') ; END Else RIB:=GS.Cells[SA_RIB,Lig] ;
    If ModeSaisie = OnEffet Then If Trim(RIB)='' Then Err:=32 ;
  END ;

  // controles de ventilation Err=12..16
  Result:=(Err=0) ;
  if Not Result then BEGIN
    ErreurSaisie(Err) ;
    if Col>0 then GS.Col:=Col ;
    END
  else BEGIN
    if Alim then BEGIN
      AlimObjetMvt(Lig,TRUE,0) ;
    END ;
  END ;
END ;

Function TFSaisieEff.Equilibre : Boolean ;
BEGIN
Result:=ArrS(SI_TotDS-SI_TotCS)=0 ;
END ;

Function TFSaisieEff.PieceCorrecte(OkMess : Boolean) : Boolean ;
Var i,Err : integer ;
    OkBanque : Boolean ;
    CGen  : TGGeneral ;
BEGIN
if PasModif then BEGIN Result:=True ; exit ; END ;
Err:=0 ;
if ((EstRempli(GS,GS.RowCount-1)) or (GS.Objects[SA_Gen,GS.RowCount-1]<>Nil) or (GS.Objects[SA_Aux,GS.RowCount-1]<>Nil)) then Err:=22 ;
if Not Equilibre then BEGIN if ModeDevise then Err:=101 else Err:=103 ; END ;
if ((Err=0) and (GS.RowCount<3)) then Err:=102 ;
OkBanque:=FALSE ;
if Err=0 then for i:=1 to GS.RowCount-2 do
   BEGIN
   if Not LigneCorrecte(i,True,False,FALSE) then BEGIN Err:=1000 ; Break ; END ;
   CGen:=GetGGeneral(GS,i) ;
   If (CGen<>NIL) And (SAJAL<>NIL) And (CGen.General=SAJAL.TRESO.CPT) Then OkBanque:=TRUE ;
   END ;
Result:=(Err=0) ;
if not Result then
   BEGIN
   if ((ModeDevise) and (Err=103)) then ForcerMode(True,0) else ErreurSaisie(Err) ;
   END Else
   BEGIN
   If (ModeSaisie=OnBQe) And (SAJAL<>NIL) And (SAJAL.TRESO.TypCtr In [AuChoix]) And (Not OkBanque) Then
    if OkMess then
     BEGIN
     // Attention : Le RIB n'est pas renseigné.
     If HLigne.Execute(31,Caption,'')<>mrYes Then Result:=FALSE ;
     END ;
   END ;
END ;

{==========================================================================================}
{========================= Traitements de calcul liés aux lignes ==========================}
{==========================================================================================}
procedure TFSaisieEff.DefautLigne ( Lig : integer ; Init : boolean ) ;
Var St : String ;
    DD : TDateTime ;
    O  : TOBM ;
    i  : integer ;
    XX : double ;
BEGIN
GS.Cells[SA_NumP,Lig]:=IntToStr(NumPieceInt) ;
if Lig>1 then
   BEGIN
   GS.Cells[SA_DateC,Lig]:=GS.Cells[SA_DateC,Lig-1] ;
   If SA_ModeP>=0 Then GS.Cells[SA_ModeP,Lig]:=GS.Cells[SA_ModeP,Lig-1] ;
   If SA_DateEche>=0 Then GS.Cells[SA_DateEche,Lig]:=GS.Cells[SA_DateEche,Lig-1] ;
   END else
   BEGIN
   GS.Cells[SA_DateC,Lig]:=DateToJour(StrToDate(E_DATECOMPTABLE.Text)) ;
   If SA_ModeP>=0 Then GS.CellValues[SA_ModeP,Lig]:=E_MODEPAIE.Value ;
   If SA_DateEche>=0 Then GS.Cells[SA_DateEche,Lig]:=E_DATEECHEANCE.Text ;
   END ;
GS.Cells[SA_NumL,Lig]:=IntToStr(Lig) ;
GS.Cells[SA_Exo,Lig]:=QuelExoDT(StrToDate(E_DATECOMPTABLE.EditText)) ;
AlloueMvt(GS,EcrGen,Lig,Init) ; O:=GetO(GS,Lig) ; if O=Nil then Exit ;
if ((OkScenario) and (Entete<>Nil) and (Init)) then
   BEGIN
   St:=Entete.GetMvt('E_REFINTERNE') ; if St<>'' then BEGIN O.PutMvt('E_REFINTERNE',St) ; GS.Cells[SA_RefI,Lig]:=St ; END ;
   St:=Entete.GetMvt('E_LIBELLE') ; if St<>'' then BEGIN O.PutMvt('E_LIBELLE',St) ; GS.Cells[SA_Lib,Lig]:=St ; END ;
   St:=Entete.GetMvt('E_REFEXTERNE') ; if St<>'' then O.PutMvt('E_REFEXTERNE',St) ;
   DD:=Entete.GetMvt('E_DATEREFEXTERNE') ; if DD>Encodedate(1901,01,01) then O.PutMvt('E_DATEREFEXTERNE',DD) ;
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
END ;

procedure TFSaisieEff.DetruitLigne( Lig : integer ; Totale : boolean ) ;
Var R : integer ;
BEGIN
GS.CacheEdit ;
if Totale then R:=GS.Row else R:=1 ;
//GS.DeleteRow(Lig+1) ; GS.DeleteRow(Lig) ;
If SAJAL.TRESO.TypCtr=Ligne Then BEGIN
   GS.DeleteRow(Lig+1) ;
   GS.DeleteRow(Lig) ;
   END
else GS.DeleteRow(Lig) ;
CalculDebitCredit ; NumeroteLignes(GS) ; NumeroteVentils ;
GS.SetFocus ; if R>=GS.RowCount then R:=GS.RowCount-1 ;
if R>1 then GS.Row:=R else GS.Row:=1 ;
GS.Col:=SA_Gen ; GereNewLigneT ; GS.Invalidate ;
GS.MontreEdit ;
END ;

procedure TFSaisieEff.TraiteMontant ( Lig : integer ; Calc : boolean ) ;
Var XC,XD : Double ;
    OBM         : TOBM ;
BEGIN
OBM:=GetO(GS,Lig) ; if OBM=Nil then Exit ;
XD:=ValD(GS,Lig) ; XC:=ValC(GS,Lig) ;
OBM.SetMontants(XD,XC,DEV,False) ;
if Calc then CalculDebitCredit ;
END ;

Procedure TFSaisieEff.PlusMvt(i : Integer ; Var SI_TDS,SI_TCS,SI_TDP,SI_TCP,SI_TDD,SI_TCD,SI_TDE,SI_TCE : Double) ;
Var DD,CD : Double ;
    OBM : TOBM ;
BEGIN
  DD:=ValD(GS,i) ;
  CD:=ValC(GS,i) ;
  SI_TDS:=SI_TDS+DD ;
  SI_TCS:=SI_TCS+CD ;
  OBM:=GetO(GS,i) ;
  if OBM<>Nil then BEGIN
    SI_TDP:=SI_TDP+OBM.GetMvt('E_DEBIT')     ; SI_TCP:=SI_TCP+OBM.GetMvt('E_CREDIT') ;
    SI_TDD:=SI_TDD+OBM.GetMvt('E_DEBITDEV')  ; SI_TCD:=SI_TCD+OBM.GetMvt('E_CREDITDEV') ;
  END ;
END ;

procedure TFSaisieEff.AfficheConf ;
BEGIN
HConf.Visible:=(ModeConf>'0') or ((PieceConf) and (Action=taConsult)) ;
END ;

procedure TFSaisieEff.CalculDebitCredit ;
Var i  : integer ;
    CptGen : String ;
    CGen : TGGeneral ;
    CAux : TGTiers ;
//    OBA  : TObjAna ;
BEGIN
ModeConf:='0' ;
SI_TotDS:=0 ; SI_TotCS:=0 ; SI_TotDP:=0 ; SI_TotCP:=0 ;
SI_TotDD:=0 ; SI_TotCD:=0 ; SI_TotDE:=0 ; SI_TotCE:=0 ;
SC_TotDS:=0 ; SC_TotCS:=0 ; SC_TotDP:=0 ; SC_TotCP:=0 ;
SC_TotDD:=0 ; SC_TotCD:=0 ; SC_TotDE:=0 ; SC_TotCE:=0 ;
for i:=1 to GS.RowCount-1 do BEGIN
   CptGen:=Trim(GS.Cells[SA_Gen,i]) ;
   if CptGen<>'' then BEGIN
      PlusMvt(i,SI_TotDS,SI_TotCS,SI_TotDP,SI_TotCP,SI_TotDD,SI_TotCD,SI_TotDE,SI_TotCE) ;
      if (SAJAL.TRESO.TypCtr In [Ligne,AuChoix]) and (CptGen=SAJAL.TRESO.CPT) then PlusMvt(i,SC_TotDS,SC_TotCS,SC_TotDP,SC_TotCP,SC_TotDD,SC_TotCD,SC_TotDE,SC_TotCE) ;
      CGen:=GetGGeneral(GS,i) ; if ((CGen<>Nil) and (CGen.Confidentiel>ModeConf)) then ModeConf:=CGen.Confidentiel ;
      CAux:=GetGTiers(GS,i) ; if CAux<>Nil then if CAux.Confidentiel>ModeConf then ModeConf:=CAux.Confidentiel ;
//      OBA:=TObjAna(GS.Objects[SA_DateC,i]) ; if OBA<>Nil then if OBA.ModeConf>ModeConf then ModeConf:=OBA.ModeConf ;
      END ;
   END ;
If SAJAL<>NIL Then
  BEGIN
  Case SAJAL.TRESO.TypCtr Of
    Ligne,AuChoix : BEGIN
                    E_DEBITTRESO.VALUE:=SC_TotDS ; E_CREDITTRESO.Value:=SC_TotCS ;
                    END ;
 PiedSolde,PiedDC : BEGIN
                    SC_TotDS:=SI_TotCS ; SC_TotCS:=SI_TotDS ;
                    SC_TotDP:=SI_TotCP ; SC_TotCP:=SI_TotDP ;
                    SC_TotDD:=SI_TotCD ; SC_TotCD:=SI_TotDD ;
                    SC_TotDE:=SI_TotCE ; SC_TotCE:=SI_TotDE ;
                    If SAJAL.TRESO.TypCtr=PiedSolde Then
                       BEGIN
                       If SC_TotDS>SC_TotCS Then
                          BEGIN
                          E_DEBITTRESO.VALUE:=SC_TotDS-SC_TotCS ; E_CREDITTRESO.Value:=0 ;
                          END Else
                          BEGIN
                          E_CREDITTRESO.VALUE:=SC_TotCS-SC_TotDS ; E_DEBITTRESO.Value:=0 ;
                          END ;
                       END Else
                       BEGIN
                       E_DEBITTRESO.VALUE:=SC_TotDS ; E_CREDITTRESO.Value:=SC_TotCS ;
                       END ;
                    OBMT[1].PutMvt('E_DEBIT',SC_TotDP)     ; OBMT[2].PutMvt('E_CREDIT',SC_TotCP) ;
                    OBMT[1].PutMvt('E_DEBITDEV',SC_TotDD)  ; OBMT[2].PutMvt('E_CREDITDEV',SC_TotCD) ;
                    SI_TotDS:=SI_TotDS+SC_TotDS ; SI_TotCS:=SI_TotCS+SC_TotCS ;
                    SI_TotDP:=SI_TotDP+SC_TotDP ; SI_TotCP:=SI_TotCP+SC_TotCP ;
                    SI_TotDD:=SI_TotDD+SC_TotDD ; SI_TotCD:=SI_TotCD+SC_TotCD ;
                    SI_TotDE:=SI_TotDE+SC_TotDE ; SI_TotCE:=SI_TotCE+SC_TotCE ;
                    END ;
  END ;
  END ;
SA_TotalDebit.Value:=SI_TotDS ; SA_TotalCredit.Value:=SI_TotCS ;
AfficheConf ;
END ;

{procedure TFSaisieEff.InfoLigne ( Lig : integer ; Var D,C,TD,TC : Double ) ;
Var Cpte : String17 ;
BEGIN
D:=0 ; C:=0 ; TD:=0 ; TC:=0 ;
Cpte:=GS.Cells[SA_Gen,Lig] ; if Cpte='' then Exit ;
D:=ValD(GS,Lig) ; C:=ValC(GS,Lig) ;
SommeSoldeCompte(SA_Gen,Cpte,TD,TC,FALSE,True) ;
END ;}

procedure TFSaisieEff.SommeSoldeCompte ( Col : integer ; Cpte : String ; Var TD,TC : Double ; Old,OkDev : Boolean) ;
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

procedure TFSaisieEff.CalculSoldeCompte ( CpteG,CpteA : String ; DIG,CIG,DIA,CIA : Double ) ;
Var TDG : Double ;
    TCG : Double ;
    TDA : Double ;
    TCA : Double ;
begin
  TDG := 0 ; TCG := 0 ;
  TDA := 0 ; TCA := 0 ;
  // Affichage solde des généraux
  SommeSoldeCompte(SA_Gen,CpteG,TDG,TCG,TRUE,FALSE) ;
  AfficheLeSolde(SA_SoldeG,TDG+DIG,TCG+CIG) ;
  // Affichage solde des tiers
  SommeSoldeCompte(SA_Aux,CpteA,TDA,TCA,TRUE,FALSE) ;
  AfficheLeSolde(SA_SoldeT,TDA+DIA,TCA+CIA) ;
END ;

{==========================================================================================}
{================================ Appels des comptes en saisie=== =========================}
{==========================================================================================}
Function TFSaisieEff.ChargeCompteEffetJAL : Boolean ;
BEGIN
  Result:=FALSE ;
  If SAJAL=NIL Then Exit ;
  If E_CONTREPARTIEGEN.Text='' Then Exit ;
  If SAJAL.Journal='' Then Exit ;
  SAJAL.TRESO.Cpt:=E_CONTREPARTIEGEN.Text ;

  {JP 17/01/08 : FQ 18563 : Calcul du solde en fonction du compte et surtout de l'exercice}
  SAJAL.PutDate(E_DATECOMPTABLE.Text);
  SAJAL.ChargeCompteTreso ;

  SAJAL.TRESO.DevBqe:=E_DEVISE.Value ;
  InitBoutonValide ;
  Result:=TRUE ;
END ;

procedure TFSaisieEff.ChercheJAL ( Jal : String3 ; Zoom : boolean ) ;
BEGIN
  SAJAL.Free ;

  If ModeSaisie<>OnBqe Then BEGIN
    SAJAL:=TSAJAL.CreateEff(Jal) ;
    if SAJAL.Treso.STypCtr='' Then BEGIN
      SAJAL.TRESO.STypCtr:='LIG' ;
      SAJAL.TRESO.TypCtr:=Ligne ;
    END ;
    If SAJAL.TRESO.Cpt<>'' Then E_CONTREPARTIEGEN.Text:=SAJAL.TRESO.Cpt ;

    ChargeCompteEffetJal ;
    END
  Else SAJAL:=TSAJAL.Create(Jal,M.Treso) ;
  Case Action of
     taCreat   : InitEnteteJal(True,Zoom,False) ;
  END ;
end;

procedure TFSaisieEff.AppelAuto ( indice : integer ) ;
Var
  Cpte : String ;
BEGIN
  if PasModif then Exit ;
  if GS.Col<>SA_Gen then Exit ;
  Cpte:=TrouveAuto(SAJAL.COMPTEAUTOMAT,indice) ;
  if Cpte='' then Exit ;
  GS.Cells[SA_Gen,GS.Row]:=Cpte ;
END ;
(*
Procedure TFSaisieEff.QuelZoomTable(C,L : Integer) ;
Var CGen : TGGeneral ;
    CAux : TGTiers ;
    PasDaux : Boolean ;
    PasDeGene,SurCptTreso : Boolean ;
begin
  CAux:=GetGTiers(GS,L) ;
  CGen:=GetGGeneral(GS,L);
  SurCptTreso:=GS.Cells[SA_GEN,L]=SAJAL.Treso.Cpt ;
  PasDaux:=(CAux=NIL) Or (GS.Cells[SA_AUX,L]='') ;
  PasDeGene:=(CGen=NIL) Or (GS.Cells[SA_GEN,L]='') ;
  If C=SA_GEN Then BEGIN
    Cache.ZoomTable:=tzGeneral ;
    if (CAux<>Nil) and (GS.Cells[SA_Aux,L]<>'') And (Not PasDaux) then QuelZoomTableG(Cache,CAux)
    Else If E_NATUREPIECE.Value='RC' Then BEGIN
      If (SAJAL.TRESO.TypCtr<>AuChoix) And (Not SurCptTreso) Then Cache.ZoomTable:=tzGToutDebit ;
      END
    Else If E_NATUREPIECE.Value='RF' Then BEGIN
      If (SAJAL.TRESO.TypCtr<>AuChoix) And (Not SurCptTreso) Then Cache.ZoomTable:=tzGToutCredit ;
    END ;
    END
  Else BEGIN
    Cache.ZoomTable:=tzTiers ;
    if (CGen<>Nil) And (Not PasDeGene)then QuelZoomTableT(Cache,CGen) Else
    If E_NATUREPIECE.Value='RC' Then BEGIN
      If (SAJAL.TRESO.TypCtr<>AuChoix) And (Not SurCptTreso) Then Cache.ZoomTable:=tzTToutDebit ;
      END
    Else If E_NATUREPIECE.Value='RF' Then BEGIN
      If (SAJAL.TRESO.TypCtr<>AuChoix) And (Not SurCptTreso) Then Cache.ZoomTable:=tzTToutCredit ;
    END ;
  END ;
end ;
*)

function TFSaisieEff.ChercheGen(C,L : integer ; Force : boolean ; DoPopup : boolean) : byte;
var
    St : String;
    CGen,CGenAvant : TGGeneral;
    CAux : TGTiers;
    Idem,Changed,CollAvant,HavePopup,SurCptTreso,found : Boolean;
    q : tquery;
begin
    result := 0;

    // sauvegarde de la valeutr saisie
    St := uppercase(ConvertJoker(GS.Cells[C,L]));

    // compte collectif ??
    CGenAvant := GetGGeneral(GS,L);
    if (CGenAvant <> nil) then CollAvant := CGenAvant.Collectif
    else CollAvant := false;

// BPY le 18/11/2004 => Fiche n° 14948 : intitulé du llokup ne doit pas changé ... modification du plus a la place de la tablette ...
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

    //==============================
    SurCptTreso := (GS.Cells[SA_GEN,L] = SAJAL.Treso.Cpt);

    if ((SAJAL.TRESO.TypCtr <> AuChoix) and (not SurCptTreso)) then
      begin
      // Ajout des comptes lettrables divers... FQ 16136 SBO 09/08/2005
      if (E_NATUREPIECE.Value = 'RC')
        then compte.Plus := '((G_COLLECTIF="X" AND (G_NATUREGENE="COC" OR G_NATUREGENE="COD")) OR (G_NATUREGENE="TID") OR ((G_NATUREGENE="DIV") AND (G_LETTRABLE="X")))'
      else if (E_NATUREPIECE.Value = 'RF')
        then compte.Plus := '((G_COLLECTIF="X" AND (G_NATUREGENE="COF" OR G_NATUREGENE="COD")) OR (G_NATUREGENE="TIC") OR ((G_NATUREGENE="DIV") AND (G_LETTRABLE="X")))';
      end;


  compte.Plus := compte.Plus + ' AND G_FERME<>"X" ' ;
      
    //==============================
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
//            {JP 20/02/04 : Sur le premier Focus, en ouverture de fiche, le CellsEnter n'est pas exécuté}
//            if ((TGGeneral(GS.Objects[C,L]).General <> '') and (StCellCur = '')) then StCellCur := TGGeneral(GS.Objects[C,L]).General;
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
                DesalloueAnal(L);
                DesalloueOBML(L);
            end;

            // compte interdit ??
            if EstInterdit(SAJAL.COMPTEINTERDIT,compte.Text, 0) > 0 then
            begin
                DesalloueEche(L);
                DesalloueAnal(L);
                Desalloue(GS,C,L);
                DesalloueOBML(L);
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
                DesalloueOBML(L);
                GS.Cells[C,L] := '';
                ErreurSaisie(20);
                Exit;
            end;

            if (ModeSaisie in [OnEffet, OnChq, OnCB]) then if (CGen.NatureGene = 'COF') then
            begin
                DesalloueEche(L);
                DesalloueAnal(L);
                Desalloue(GS,C,L);
                DesalloueOBML(L);
                GS.Cells[C,L] := '';
                ErreurSaisie(33);
                Exit;
            end;

            // SBO 19/08/2004 FQ 13290 Saisie interdite sur les comptes fermés
            if CGen.Ferme then
            begin
                DesalloueEche(L);
                DesalloueAnal(L);
                Desalloue(GS,C,L);
                DesalloueOBML(L);
                GS.Cells[C,L] := '';
                ErreurSaisie(31);
                Exit;
            end;

            // recup l'analytique
            if (not Idem) then AlloueEcheAnal(C,L);

            Result := 1;

            // affecte les truc
            AffecteConso(C,L);
            if ((CGen.NatureGene = 'TIC') or (CGen.NatureGene = 'TID')) then if (AffecteRIB(C,L,FALSE)) then BModifRIB.Enabled := true;
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

        // Chargement TInfoECriture // FQ 13246 : SBO 30/03/2005
        if GS.Cells[C,L] <> '' then
          FInfo.LoadCompte( GS.Cells[C,L] ) ;

        AffichePied;
    end;
end;


(*
Function TFSaisieEff.ChercheGen ( C,L : integer ; Force : boolean ) : byte ;
Var St   : String ;
    CGen,CGenAvant : TGGeneral ;
    Idem,CollAvant,Changed : boolean ;
BEGIN
  Result:=0 ;
  Changed:=False ;
{  if M.ANouveau then Cache.ZoomTable:=tzGBilan
                 else Cache.ZoomTable:=tzGeneral ;}
  St:=uppercase(ConvertJoker(GS.Cells[C,L])) ;
  Cache.Text:=St ;
  CGenAvant:=GetGGeneral(GS,L) ;
  CollAvant:=FALSE ;
  if CGenAvant<>Nil then CollAvant:=CGenAvant.Collectif ;

  QuelZoomTable(SA_Gen,L) ;
{   CAux:=GetGTiers(GS,L) ;
  if CAux<>Nil then QuelZoomTableG(Cache,CAux) ;}
  if GChercheCompte(Cache,FicheGene) then BEGIN
    if GS.Objects[C,L]<>Nil then
      if TGGeneral(GS.Objects[C,L]).General<>St then Changed:=True ;
    Idem:=((St=Cache.Text) and (GS.Objects[C,L]<>Nil) and (Not Changed)) ;
    if ((Not Idem) or (Force)) then BEGIN
      if EstInterdit(SAJAL.COMPTEINTERDIT,Cache.Text,0)>0 then BEGIN ErreurSaisie(4) ; Exit ; END ;
      if ((Not M.ANouveau) and (EstOuvreBil(Cache.Text))) then BEGIN ErreurSaisie(20) ; Exit ; END ;
      GS.Cells[C,L]:=Cache.Text ;
      Desalloue(GS,C,L) ;
      CGen:=TGGeneral.Create(Cache.Text) ;
      GS.Objects[C,L]:=TObject(CGen) ;
      if Not Idem then BEGIN
        if ((Not CollAvant) or (Not CGen.Collectif)) then DesalloueEche(L) ;
        DesalloueAnal(L) ;
        DesalloueOBML(L) ;
      END ;
      If ModeSaisie In [OnEffet,OnChq,OnCB] Then BEGIN
        If CGen.NatureGene='COF' Then BEGIN
          DesalloueEche(L) ;
          DesalloueAnal(L) ;
          Desalloue(GS,C,L) ;
          DesalloueOBML(L) ;
          GS.Cells[C,L]:='' ;
          ErreurSaisie(33) ;
          Exit ;
        END ;
      END ;
      if CGen.Ferme then BEGIN
        if Arrondi(CGen.TotalDebit-CGen.TotalCredit,V_PGI.OkDecV)=0 then BEGIN
          DesalloueEche(L) ;
          DesalloueAnal(L) ;
          Desalloue(GS,C,L) ;
          DesalloueOBML(L) ;
          GS.Cells[C,L]:='' ;
          ErreurSaisie(31) ;
          Exit ;
          END
        else BEGIN
          ErreurSaisie(17) ;
        END ;
      END ;
      if Not Idem then AlloueEcheAnal(C,L) ;
      Result:=1 ;
      AffecteConso(C,L) ;
      If (CGen.NatureGene='TID') Or (CGen.NatureGene='TIC') Then
        If AffecteRIB(C,L,FALSE) Then BModifRIB.Enabled:=TRUE ;
      END
    else BEGIN
      CGen:=TGGeneral(GS.Objects[C,L]) ;
      Result:=2 ;
    END ;

    if ((Result>0) and (CGen<>Nil)) then if not CGen.Collectif then BEGIN
      Desalloue(GS,SA_Aux,L) ;
      GS.Cells[SA_Aux,L]:='' ;
      if Lettrable(CGen)=NonEche then DesalloueEche(L) ;
      if Not Ventilable(CGen,0) then DesalloueAnal(L) ;
      If Lettrable(CGen)=MultiEche Then BZoomMvtEff.Enabled:=TRUE ;
    END ;

    // Chargement TInfoECriture // FQ 13246 : SBO 30/03/2005
    if GS.Cells[C,L] <> '' then
      FInfo.LoadCompte( GS.Cells[C,L] ) ;

    AffichePied ;
  END ;
END ;
*)

Procedure TFSaisieEff.AffecteConso ( C,L : integer ) ;
Var O : TOBM ;
    CAux : TGTiers ;
    CGen : TGGeneral ;
    Code : String ;
BEGIN
  O:=GetO(GS,L) ; if O=Nil then Exit ;
  if C=SA_AUX then BEGIN
    CAux:=GetGTiers(GS,L) ;
    if CAux=Nil then Exit ;
    Code:=CAux.CodeConso ;
    END
  else BEGIN
    CGen:=GetGGeneral(GS,L) ;
    if CGen=Nil then Exit ;
    Code:=CGen.CodeConso ;
  END ;
  if ((C=SA_Aux) or (Code<>'')) then O.PutMvt('E_CONSO',Code) ;
END ;

Function TFSaisieEff.AffecteRIB ( C,L : integer ; IsAux : Boolean) : Boolean ;
Var O : TOBM ;
    Q : TQuery ;
    Cpt,St,SRIB : String ;
BEGIN
  Result:=FALSE ;
  if Action=taConsult then Exit ;
  If (ModeSaisie In [OnEffet,OnBqe])=FALSE Then Exit ;
  if (IsAux And (C<>SA_Aux)) Or ((Not IsAux) And (C<>SA_Gen)) then Exit ;
  Cpt:=GS.Cells[C,L] ;
  if Cpt='' then Exit ;
  O:=GetO(GS,L) ;
  if O=Nil then Exit ;
  sRIB:=O.GetMvt('E_RIB') ;
  Result:=TRUE ;
(*
if OkScenario then
   BEGIN
   If (GestionRIB='MAN') or (GestionRIB='PRI') Then Result:=TRUE ;
   if GestionRIB<>'PRI' then Exit ;
   END else
   BEGIN
   if Not VH^.AttribRIBAuto then Exit ;
   if sRIB<>'' then Exit ;
   END ;
*)
  Q:=OpenSQL('Select * from RIB Where R_AUXILIAIRE="'+Cpt+'" AND R_PRINCIPAL="X"',True) ;
  if Not Q.EOF then BEGIN
    {JP 09/05/07 : FQ 19919 : Mieux vaut travailler sur le codeIso2 du pays, car la fiche qualité concerne
                   le cas d'un code pays à FR.
     if (Q.FindField('R_PAYS').AsString <> 'FRA') then St:='*'+Q.FindField('R_CODEIBAN').AsString}
    if CodeIsoDuPays(Q.FindField('R_PAYS').AsString) <> CodeISOFR then
      St:='*'+Q.FindField('R_CODEIBAN').AsString
      
    else St:=EncodeRIB(Q.FindField('R_ETABBQ').AsString,
                       Q.FindField('R_GUICHET').AsString,
                       Q.FindField('R_NUMEROCOMPTE').AsString,
                       Q.FindField('R_CLERIB').AsString,
                       Q.FindField('R_DOMICILIATION').AsString) ;
    O.PutMvt('E_RIB',St) ;
    If IsAux Then O.PutMvt('E_AUXILIAIRE',Cpt)
             Else O.PutMvt('E_GENERAL',Cpt) ;
    If SA_Rib>-1 Then GS.Cells[SA_RIB,L]:=Trim(St) ;
    Result:=TRUE ;
    END
  else BEGIN
    if sRIB<>'' then BEGIN
      If SA_Rib>-1 Then GS.Cells[SA_RIB,L]:='' ;
      O.PutMvt('E_RIB','') ;
    END ;
  END ;
  Ferme(Q) ;
END ;

Function TFSaisieEff.VerifEtUpdateRib ( L : integer ) : Boolean ;
Var Cpt,Rib : String ;
BEGIN
  Result:=TRUE ;
  Rib:=GS.Cells[SA_Rib,L] ;
  If GS.Cells[SA_Aux,L]='' Then Cpt:=GS.Cells[SA_Gen,L]
                           Else Cpt:=GS.Cells[SA_Aux,L] ;
  If FormatRIBEcr(Cpt,TRUE,TRUE,Rib)=0 Then GS.Cells[SA_RIB,L]:=Trim(Rib)
                                       Else Result:=FALSE ;
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
function TFSaisieEff.ChercheAux(C,L : integer ; Force : boolean ; DoPopup : boolean) : byte;
var
    st : string;
    CAux : TGTiers;
    CGen,GenColl : TGGeneral;
    Idem,Changed,HavePopup,SurCptTreso,found : boolean;
    Err : integer;
    q : tquery;
begin
    result := 0;

    // sauvegarde de la valeutr saisie
    St := uppercase(ConvertJoker(GS.Cells[C,L]));

// BPY le 15/11/2004 => Fiche n° 14948 : intitulé du llokup ne doit pas changé ... modification du plus a la place de la tablette ...
    // set de la tablette
    compte.DataType := 'TZTTOUS';
    compte.Plus := '(T_NATUREAUXI<>"NCP" AND T_NATUREAUXI<>"CON" AND T_NATUREAUXI<>"PRO" AND T_NATUREAUXI<>"SUS")' ; // FQ 15838 : SBO 03/08/2005

    // ... ou en fonction du compte general !!
    CGen := GetGGeneral(GS,L);
    if ((CGen <> Nil) and (GS.Cells[SA_Gen,L] <> '')) then
    begin
        if (CGen.NatureGene = 'COC') then compte.Plus := '(T_NATUREAUXI="AUD" OR T_NATUREAUXI="CLI" OR T_NATUREAUXI="DIV")'
        else if (CGen.NatureGene = 'COF') then compte.Plus := '(T_NATUREAUXI="AUC" OR T_NATUREAUXI="FOU" OR T_NATUREAUXI="DIV" OR T_NATUREAUXI="SAL")'
        else if (CGen.NatureGene = 'COS') then compte.Plus := 'T_NATUREAUXI="SAL"';
    end;

    //==============================
    SurCptTreso := (GS.Cells[SA_GEN,L] = SAJAL.Treso.Cpt);

    if ((SAJAL.TRESO.TypCtr <> AuChoix) and (not SurCptTreso)) then
    begin
        if (E_NATUREPIECE.Value = 'RC') then compte.Plus := '(T_NATUREAUXI="AUD" OR T_NATUREAUXI="CLI" OR T_NATUREAUXI="DIV")'
        else if (E_NATUREPIECE.Value = 'RF') then compte.Plus := '(T_NATUREAUXI="AUC" OR T_NATUREAUXI="FOU" OR T_NATUREAUXI="DIV" OR T_NATUREAUXI="SAL")';
    end;

    compte.Plus := compte.Plus + ' AND T_FERME<>"X" ' ;
    
    //==============================
// Fin BPY

    // set des coordonnées
    compte.Top := GS.Top + GS.CellRect(C,L).Top;
    compte.Left := Gs.Left + GS.CellRect(C,L).Left;
    compte.Height := GS.CellRect(C,L).Bottom - GS.CellRect(C,L).top;
    compte.Width := GS.CellRect(C,L).Right - GS.CellRect(C,L).Left;

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
        result := 1;

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
            if ((Not CAux.MultiDevise) and (DEV.Code <> CAux.Devise) and (DEV.Code <> V_PGI.DevisePivot) and (CAux.Devise <> '')) then
            begin
                DesalloueEche(L);
                Desalloue(GS,C,L);
                DesalloueOBML(L);
                GS.Cells[C,L] := '';
                ErreurSaisie(28);
                result := 0;
                Exit;
            end;

            // ???
            if ((not CAux.Lettrable) and (SAJAL.TRESO.TypCtr <> AuChoix)) then
            begin
                DesalloueEche(L);
                Desalloue(GS, C, L);
                DesalloueOBML(L);
                GS.Cells[C, L] := '';
                ErreurSaisie(29);
                Exit;
            end;

            // ???
            if (ModeSaisie in [OnEffet, OnChq, OnCB]) then if (CAux.NatureAux = 'FOU') then
            begin
                DesalloueEche(L);
                Desalloue(GS, C, L);
                DesalloueOBML(L);
                GS.Cells[C, L] := '';
                ErreurSaisie(33);
                Exit;
            end;

            // SBO 19/08/2004 FQ 13290 Saisie interdite sur les comptes fermés
            if CAux.Ferme then
            begin
                DesalloueEche(L);
                Desalloue(GS,C,L);
                DesalloueOBML(L);
                GS.Cells[C,L] := '';
                ErreurSaisie(31);
                result := 0;
                Exit;
            end;

            // recup l'analytique
            if (not Idem) then AlloueEcheAnal(C,L);

            Result := 1;

            // affecte les truc
            if (AffecteRIB(C,L,true)) then BModifRIB.Enabled := true;
            if (CAux.Lettrable) then BZoomMvtEff.Enabled := true;
            AffecteConso(C,L);
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
                DesalloueOBML(L);
                GS.Cells[SA_Gen,L] := '';
                GS.Cells[SA_Aux,L] := '';
                ErreurSaisie(24);
            end
            else AlloueEcheAnal(SA_Gen,L);
        end;

        // Chargement TInfoECriture // FQ 13246 : SBO 30/03/2005
        if GS.Cells[C,L] <> '' then
          FInfo.LoadAux( GS.Cells[C,L] ) ;

        AffichePied;
    end;
end;

(*
Function TFSaisieEff.ChercheAux ( C,L : integer ; Force : boolean ) : byte ;
Var St   : String ;
    CAux : TGTiers ;
    CGen,GenColl : TGGeneral ;
    Idem,Changed : boolean ;
    Err  : integer ;
BEGIN
  Result:=0 ;
  Err:=0 ;
  Changed:=False ;
  Cache.ZoomTable:=tzTiers ;
  St:=uppercase(ConvertJoker(GS.Cells[C,L])) ;
  Cache.Text:=St ;
  CGen:=GetGGeneral(GS,L) ;
  if CGen<>Nil then QuelZoomTableT(Cache,CGen) ;
  QuelZoomTable(SA_Aux,L) ;
{$IFDEF EAGLCLIENT}
  if GChercheCompte(Cache,Nil) then
{$ELSE}
	if GChercheCompte(Cache,FicheTiers) then
{$ENDIF}
  begin
    if GS.Objects[C,L]<>Nil then
      if TGTiers(GS.Objects[C,L]).Auxi<>St then Changed:=True ;

    Idem:=((St=Cache.Text) and (GS.Objects[C,L]<>Nil) and (Not Changed)) ;

    if ((Not Idem) or (Force)) then BEGIN
      GS.Cells[C,L]:=Cache.Text ;
      Desalloue(GS,C,L) ;
      if Not Idem then BEGIN
        DesalloueEche(L) ;
        if ((CGen=Nil) or (GS.Cells[SA_Gen,L]='')) then DesalloueAnal(L)
        else if CGen<>Nil then BEGIN
          if Not Ventilable(CGen,0) then DesalloueAnal(L) ;
        END ;
      END ;

      CAux:=TGTiers.Create(Cache.Text) ;
      GS.Objects[C,L]:=TObject(CAux) ;
      if ((Not CAux.MultiDevise) and (DEV.Code<>CAux.Devise) and (CAux.Devise<>'') and (DEV.Code<>V_PGI.DevisePivot)) then BEGIN
        DesalloueEche(L) ;
        Desalloue(GS,C,L) ;
        GS.Cells[C,L]:='' ;
        DesalloueOBML(L) ;
        ErreurSaisie(28) ;
        Exit ;
      END ;
      if (Not CAux.Lettrable) and (SAJAL.TRESO.TypCtr<>AuChoix) then BEGIN
        DesalloueEche(L) ;
        Desalloue(GS,C,L) ;
        GS.Cells[C,L]:='' ;
        DesalloueOBML(L) ;
        ErreurSaisie(29) ;
        Exit ;
      END ;
      If ModeSaisie In [OnEffet,OnChq,OnCB] Then BEGIN
        If CAux.NatureAux='FOU' Then BEGIN
          DesalloueEche(L) ;
          Desalloue(GS,C,L) ;
          GS.Cells[C,L]:='' ;
          DesalloueOBML(L) ;
          ErreurSaisie(33) ;
          Exit ;
        END ;
      END ;
      if CAux.Ferme then BEGIN
        if Arrondi(CAux.TotalDebit-CAux.TotalCredit,V_PGI.OkDecV)=0 then BEGIN
          DesalloueEche(L) ;
          Desalloue(GS,C,L) ;
          GS.Cells[C,L]:='' ;
          DesalloueOBML(L) ;
          ErreurSaisie(31) ;
          Exit ;
          END
        else BEGIN
          ErreurSaisie(27) ;
        END ;
      END ;
      if Not Idem then AlloueEcheAnal(C,L) ;
      Result:=1 ;
      If AffecteRIB(C,L,TRUE) Then BModifRIB.Enabled:=TRUE ;
      If CAux.Lettrable Then BZoomMvtEff.Enabled:=TRUE ;
      AffecteConso(C,L) ;
      END
    else BEGIN
      Result:=2 ;
    END ;

    if Result>0 then if GS.Cells[SA_Gen,L]='' then BEGIN
      GS.Cells[SA_Gen,L]:=GetGTiers(GS,L).Collectif ;
      Desalloue(GS,SA_Gen,L) ;
      GenColl:=TGGeneral.Create(GS.Cells[SA_Gen,L]) ;
      GS.Objects[SA_Gen,L]:=GenColl ;
      if EstInterdit(SAJAL.COMPTEINTERDIT,GS.Cells[SA_Gen,L],0)>0 then Err:=4 ;
      if EstConfidentiel(GenColl.Confidentiel) then Err:=24 ;
      if Err>0 then BEGIN
        DesalloueEche(L) ;
        Desalloue(GS,C,L) ;
        Desalloue(GS,SA_Gen,L) ;
        DesalloueOBML(L) ;
        GS.Cells[SA_Gen,L]:='' ;
        GS.Cells[SA_Aux,L]:='' ;
        ErreurSaisie(24) ;
        END
      else BEGIN
        AlloueEcheAnal(SA_Gen,L) ;
      END ;
    END ;

    // Chargement TInfoECriture // FQ 13246 : SBO 30/03/2005
    if GS.Cells[C,L] <> '' then
      FInfo.LoadAux( GS.Cells[C,L] ) ;

    AffichePied ;
  END ;
END ;

*)

procedure TFSaisieEff.ChercheMontant(Acol,Arow : longint) ;
BEGIN
  if ACol=SA_Debit then BEGIN
    if ValD(GS,ARow)<>0 then GS.Cells[SA_Credit,ARow]:=''
                        else GS.Cells[ACol,ARow]:='' ;
    END
  else BEGIN
    if ValC(GS,ARow)<>0 then GS.Cells[SA_Debit,ARow]:=''
                        else GS.Cells[ACol,ARow]:='' ;
  END ;
  FormatMontant(GS,ACol,ARow,DECDEV) ;
  TraiteMontant(ARow,True) ;
END ;

{==========================================================================================}
{================================= Conversions, Caluls ====================================}
{==========================================================================================}
Function  TFSaisieEff.DateMvt ( Lig : integer ) : TDateTime ;
BEGIN
  Result:=StrToDate(E_Datecomptable.Text) ;
END ;

Function  TFSaisieEff.ArrS( X : Double ) : Double ;
BEGIN
  Result:=Arrondi(X,DECDEV) ;
END ;

{==========================================================================================}
{=========================== Calculs sur ligne de trésorerie ==============================}
{==========================================================================================}

procedure TFSaisieEff.E_NUMDEPExit(Sender: TObject);
Var ST,St1,Masque : String ;
    i : Integer ;
begin
St:=Trim(E_NUMDEP.Text) ; St1:=Trim(E_REF.Text) ;
If St='' then BEGIN SI_NUMREF:=-1 ; Exit ; END ;
Masque:='' ; For i:=1 To Length(St) Do Masque:=Masque+'0' ;
E_NUMREF.Masks.PositiveMask:=St1+' '+Masque ;
SI_NUMREF:=StrToInt(E_NUMDEP.Text) ;
E_NUMREF.VALUE:=SI_NUMREF ;
end;

Procedure TFSaisieEff.IncRef ;
Begin
  If SI_NUMREF>=0 Then BEGIN
    Inc(SI_NUMREF) ;
    E_NUMREF.Value:=SI_NUMREF ;
  END ;
End ;

//TR sur toutes ces procédures :
procedure TFSaisieEff.CalculRefAuto(Sender : TObject) ;
begin
If SAJAL=NIL Then Exit ;

  PRefAuto.Left:=GS.Left+(GS.Width-PRefAuto.Width) div 2 ;
  PRefAuto.Top:=GS.Top+(GS.Height-PRefAuto.Height) div 2 ;
  PRefAuto.Visible:=True ;
  E_REF.SetFocus ;
  PEntete.Enabled:=False ;
  Outils.Enabled:=False ;
  Valide97.Enabled:=False ;
  GS.Enabled:=False ;
  PPied.Enabled:=False ;
  ModeRA:=True ;
end;

procedure TFSaisieEff.ClickLibelleAuto ;
var
  StL, StRef, StLib: string;
  Q: TQuery;
  i, OldL: integer;
  OkOk: Boolean;
  Swhere1,SWhere2 : String ;
begin
  if not AutoCharged then
  begin
    // FQ 16095 SBO 11/10/2005
   SWhere1:='' ; SWhere2:='' ;
   If SAJAL<>NIL Then SWhere1:=' AND RA_JOURNAL="'+SAJAL.Journal+'" ' ;
   SWhere2:=' RA_NATUREPIECE="'+E_NATUREPIECE.Value+'" ' ;
   If SWhere1<>'' Then SWhere2:=SWhere2+SWhere1 ;
   Q:=OpenSQL('SELECT RA_FORMULEREF FROM REFAUTO WHERE '+SWhere2,True);
   if Q.EOF then
     StL:=Choisir(HDiv.Mess[1],'REFAUTO','RA_LIBELLE ||"(" || RA_JOURNAL || " " || RA_NATUREPIECE || ")"','RA_CODE','','', False, False, 7244310)
   else
     StL:=Choisir(HDiv.Mess[1],'REFAUTO','RA_LIBELLE ||"(" || RA_JOURNAL || " " || RA_NATUREPIECE || ")"','RA_CODE',SWhere2,'', False, False, 7244310);
   Ferme(Q) ;
    // Fin FQ 16095 SBO 11/10/2005
    if StL='' then BEGIN
      // Aucun libellé automatique n'a été choisi.
      HPiece.Execute(15,caption,'') ;
      Exit ;
    END ;
    Q:=OpenSQL('Select RA_FORMULEREF, RA_FORMULELIB from REFAUTO where RA_CODE="'+StL+'"',True) ;
    SI_FormuleRef:=Q.FindField('RA_FORMULEREF').AsString ;
    SI_FormuleLib:=Q.FindField('RA_FORMULELIB').AsString ;
    Ferme(Q) ;
  end;
  OldL := CurLig;

  OkOk := True;

  {JP 20/10/05 : FQ 16923 : Pour une mise à jour de la ligne de contrepartie, si elle existe déjà}
  if (GS.Cells[SA_Debit, CurLig] <> '') or (GS.Cells[SA_Credit, CurLig] <> '') then
    GS.OnRowExit(GS, CurLig, OkOk, True);

  BeginTrans;
  for i := 1 to GS.RowCount - 2 do begin
  {JP 20/10/05 : FQ 16923 : Mise en conformité avec Saisie.Pas de la gestion des libellés automatiques}
    OkOk := True;
    if (SAJAL.Treso.TypCtr = Ligne) and (GS.Cells[SA_Gen, i] = SAJAL.TRESO.Cpt) then
      OkOk := False
    else begin
      StRef := '';
      StLib := '';
    end;
    CurLig := i;
    GereRefLib(i, StRef, StLib, OkOk);
  end;
  CommitTrans;
  GS.EditorMode := True;
  CurLig := OldL;
end;

procedure TFSaisieEff.ModifNaturePiece ;
Var OBM, OBA : TOBM ;
    i,NumA,NumL : Integer ;
    St,NaturePieceAvant : String ;
    CGen  : TGGeneral ;
begin
  for i:=1 to GS.RowCount-2 do BEGIN
    St:=E_NATUREPIECE.Value ;
    OBM:=GetO(GS,i) ;
    if OBM<>NIL then BEGIN
      NaturePieceAvant:=OBM.GetMvt('E_NATUREPIECE') ;
      If (NaturePieceAvant<>St) Then BEGIN
        OBM.PutMvt('E_NATUREPIECE',St) ;
      END ;
    END ;

    // Analytique
    CGen:=GetGGeneral(GS,i);
    for NumA:=1 to Min(MaxAxe , OBM.Detail.Count) do begin
      if Ventilable(CGen,NumA) then BEGIN
        if OBM.Detail[NumA-1].Detail.Count>0 then BEGIN
          for NumL:=0 to OBM.Detail[NumA-1].Detail.Count-1 do BEGIN
            OBA := TOBM(OBM.Detail[NumA-1].Detail[NumL]);
            OBA.PutMvt('Y_NATUREPIECE',St) ;
          END ;
        END ;
      END ;
    end;
  END ;
end;

procedure TFSaisieEff.ModifModePaie ;
Var OBM : TOBM ;
    i : Integer ;
    St,ModePaieAvant : String ;
    M : TMOD ;
begin
  for i:=1 to GS.RowCount-2 do BEGIN
    OBM:=GetO(GS,i) ;
    St:=E_MODEPAIE.Value ;
    if OBM<>NIL then BEGIN
      If (EstLettrable(i) In [MonoEche,MultiEche]) Then BEGIN
        M:=TMOD(GS.Objects[SA_NumP,i]) ;
        If M<>NIL Then BEGIN
          ModePaieAvant:=M.MODR.TabEche[1].ModePaie ;
          If (ModePaieAvant<>St) Then BEGIN
            M.MODR.TabEche[1].ModePaie:=E_MODEPAIE.Value ;
            If SA_ModeP>=0 Then GS.CellValues[SA_ModeP,i]:=E_MODEPAIE.Value ;
          END ;
        END ;
      END ;
    END ;
  END ;
end;

procedure TFSaisieEff.AfficheLigneTreso ;
Var i,ll : Integer ;
begin
  If SAJAL=NIL Then Exit ;
  LigneTreso:=Not LigneTreso ;
  {JP 15/01/08 : FQ 19363 : Hauteur à 1 au lieur de 0. Cela change légèrement l'ergonomie,
                 mais c'est la seule solution que j'ai trouvé au positionnement de la ScrollBar.
                 En fait si la hauteur est à 0, la position théorique dans grille n'est pas la
                 même que la position réèlle car 0 ne rend pas réellement la ligne invisible
                 puisque le trait de séparation est grossi (Cf FQ 16790 pour eSaisieTr)}
  ll := 1;
  If LigneTreso then ll:=GS.RowHeights[0] ;
  for i:=1 to GS.RowCount-2 do BEGIN
    if GS.Cells[SA_Gen,i]=SAJAL.TRESO.CPT then GS.RowHeights[i]:=ll ;
  END ;
end;

procedure TFSaisieEff.E_CONTREPARTIEGENChange(Sender: TObject);
begin
If SAJAL<>NIL Then E_CPTTRESO.Caption:=SAJAL.Treso.Cpt ;
end;

procedure TFSaisieEff.E_CONTREPARTIEGENExit(Sender: TObject);
begin
If ModeSaisie<>OnBqe Then ChargeCompteEffetJal ;
OkPourLigne ;
If SAJAL<>NIL Then
  BEGIN
  E_CPTTRESO.Caption:=SAJAL.Treso.Cpt ;
  AlimObjetMvt(0,FALSE,1) ; AlimObjetMvt(0,FALSE,2) ;
  END ;
end;

procedure TFSaisieEff.CalculSoldeCompteT  ;
Var TDG,TCG : Double ;
BEGIN
  If SAJAL<>NIL Then BEGIN
    TDG:=SAJAL.TRESO.TotD+SC_TotDP ;
    TCG:=SAJAL.TRESO.TotC+SC_TotCP ;
    AfficheLeSolde(SA_SoldeTR,TDG,TCG) ;
  END ;
END ;

Procedure TFSaisieEff.RecopieGS(i : Integer) ;
Var OBM1,OBM2 : TOBM ;
    CGen : TGGeneral ;
    OMODR1,OMODR2 : TMOD ;
    Cpte : String17 ;
    Let : TLettre ;
    Okok : boolean ;
begin
  GS.Cells[SA_GEN,i+1]   :=SAJAL.TRESO.Cpt ;
  GS.Cells[SA_AUX,i+1]   :='' ;
  GS.Cells[SA_REFI,i+1]  :=GS.Cells[SA_REFI,i] ;
  GS.Cells[SA_LIB,i+1]   :=GS.Cells[SA_LIB,i] ;
  GS.Cells[SA_DEBIT,i+1] :=GS.Cells[SA_CREDIT,i] ;
  GS.Cells[SA_CREDIT,i+1]:=GS.Cells[SA_DEBIT,i] ;
  If SA_RIB>-1 Then GS.Cells[SA_RIB,i+1]:=GS.Cells[SA_RIB,i] ;
  If SA_REFE>-1 Then GS.Cells[SA_REFE,i+1]:=GS.Cells[SA_REFE,i] ;
  If SA_REFL>-1 Then GS.Cells[SA_REFL,i+1]:=GS.Cells[SA_REFL,i] ;
  If SA_NumTraChq>-1 Then GS.Cells[SA_NumTraChq,i+1]:=GS.Cells[SA_NumTraChq,i] ;
  If SA_DATEECHE>-1 Then GS.Cells[SA_DATEECHE,i+1]:=GS.Cells[SA_DATEECHE,i] ;
  If SA_MODEP>-1 Then GS.Cells[SA_MODEP,i+1]:=GS.Cells[SA_MODEP,i] ;
  // Suite tests avec comptes divers lettrables, le chargement des comptes doit être forcé FQ 16136 SBO 09/08/2005
  ChercheGen(SA_Gen,i+1,True) ;
  CGen:=GetGGeneral(GS,i+1) ;
  If (ModeSaisie<>OnBqe) And CGen.Collectif Then BEGIN
    GS.Cells[SA_AUX,i+1]:=GS.Cells[SA_AUX,i] ;
    If ChercheAux(SA_Aux,i+1,True)=0 then ;  // idem ci-dessus FQ 16136 SBO 09/08/2005
  END ;
  AlimObjetMvt(i+1,TRUE,0) ;
  OMODR1:=TMOD(GS.Objects[SA_NumP,i]) ;
  OMODR2:=TMOD(GS.Objects[SA_NumP,i+1]) ;
  Let:=EstLettrable(i+1) ;
  OkOk:=(Let=MonoEche) Or ((ModeSaisie<>OnBqe) And (Let=MultiEche)) ;

  // Compte de banque pointable
  If OkOk Then BEGIN
    DebutOuvreEche(i+1,Cpte,'',OMODR2) ;
    FinOuvreEche(i+1,OMODR2,Let) ;
    With OMODR2.MODR Do BEGIN
      NbEche:=1 ;
      TabEche[1].Pourc:=100.0 ;
      TabEche[1].MontantP:=TotalAPayerP ;
      TabEche[1].MontantD:=TotalAPayerD ;
      If OModR1<>NIL Then BEGIN
        TabEche[1].DateEche:=OMODR1.MODR.TabEche[1].DateEche ;
        TabEche[1].DateValeur:=OMODR1.MODR.TabEche[1].DateValeur ;
        TabEche[1].ModePaie:=OMODR1.MODR.TabEche[1].ModePaie ;
//        TabEche[1].NumTraiteChq:=OMODR1.MODR.TabEche[1].NumTraiteChq ;
        END
      Else BEGIN
        TabEche[1].DateEche:=DateMvt(i+1) ;
        TabEche[1].ModePaie:=E_MODEPAIE.Value ;
        TabEche[1].DateValeur:=CalculDateValeur(DateMvt(i+1),MonoEche,SAJAL.Treso.CPT) ;
      END ;
    END ;
  END ;
  OBM1:=GetO(GS,i) ;
  OBM2:=GetO(GS,i+1) ;
  If (OBM1=NIL) or (OBM2=NIL) then Exit ;
  If OModR2<>NIL Then AlimEcheVersObm(i+1,OMODR2.MODR.TabEche[1].ModePaie,OMODR2.MODR.TabEche[1].DateEche, OMODR2.MODR.TabEche[1].DateValeur) ;
  OBM2.PutMvt('E_REFEXTERNE',OBM1.GetMvt('E_REFEXTERNE')) ;
  OBM2.PutMvt('E_DATEREFEXTERNE',OBM1.GetMvt('E_DATEREFEXTERNE')) ;
  OBM2.PutMvt('E_AFFAIRE',OBM1.GetMvt('E_AFFAIRE')) ;
  OBM2.PutMvt('E_REFLIBRE',OBM1.GetMvt('E_REFLIBRE')) ;
  OBM2.PutMvt('E_QTE1',OBM1.GetMvt('E_QTE1')) ;
  OBM2.PutMvt('E_QTE2',OBM1.GetMvt('E_QTE2')) ;
  OBM2.PutMvt('E_QUALIFQTE1',OBM1.GetMvt('E_QUALIFQTE1')) ;
  OBM2.PutMvt('E_QUALIFQTE2',OBM1.GetMvt('E_QUALIFQTE2')) ;
  OBM2.PutMvt('E_RIB',OBM1.GetMvt('E_RIB')) ;
  OBM2.PutMvt('E_NUMTRAITECHQ',OBM1.GetMvt('E_NUMTRAITECHQ')) ;

end ;

Procedure TFSaisieEff.AlimLigneTreso(Lig : Integer) ;
BEGIN
  If Lig<1 then Exit ;
  // Nouvelle Ligne
  If Lig=GS.RowCount-2 then BEGIN
    NewLigne(GS) ;
    DefautLigne(Lig+1,True) ;
    RecopieGS(Lig) ;
    GS.Cells[SA_NUML,Lig+2]:=IntToStr(Lig+2) ;
    GereAnal(Lig+1,True,True,TRUE,TRUE) ;
    END
  else BEGIN
    RecopieGS(Lig) ;
    GereAnal(Lig+1,True,True,TRUE,FALSE) ;
  END ;
  If Not LigneTreso then
    {JP 15/01/08 : FQ 19363 : Hauteur à 1 au lieur de 0. Cela change légèrement l'ergonomie,
                   mais c'est la seule solution que j'ai trouvé au positionnement de la ScrollBar.
                   En fait si la hauteur est à 0, la position théorique dans grille n'est pas la
                   même que la position réèlle car 0 ne rend pas réellement la ligne invisible
                   puisque le trait de séparation est grossi (Cf FQ 16790 pour eSaisieTr)}
    GS.RowHeights[Lig + 1] := 1;
END ;

Procedure TFSaisieEff.GereLigneTreso ( Lig : integer ) ;
BEGIN
  Case SAJAL.TRESO.TypCtr Of
    Ligne  : BEGIN
             AlimLigneTReso(Lig) ;
             CalculDebitCredit ;
             END ;
    PiedDC,PiedSolde : BEGIN
                   END ;
    AuChoix : ;
  END ;
END ;

Procedure TFSaisieEff.LigneEnPlusTreso(Sens : Integer ; MSaisi,MPivot,MDevise : Double) ;
Var Lig : Integer ;
    OBM,OBM1 : TOBM ;
    LDD,LCD,LDP,LCP,LDS,LCS : Double ;
    CGen : TGGeneral ;
    {$IFNDEF OLDVENTIL}
    bOuvreAnal : boolean ;
    {$ENDIF OLDVENTIL}
    lNumAxe : integer ;
BEGIN
  If EstRempli(GS,GS.RowCount-1) Then GS.RowCount:=GS.RowCount+1 ;
  Lig:=GS.RowCount-1 ;
  DefautLigne(Lig,TRUE) ;
  OBM1:=GetO(GS,Lig) ;
  If OBM1<>NIL Then BEGIN
   // LG* 21/02/2002
    OBM1.Dupliquer ( OBMT[Sens] , true , true );
    OBM1.M.Assign(OBMT[Sens].M) ;
  END ;
  GS.Cells[SA_Gen,Lig]:=OBMT[Sens].GetMvt('E_GENERAL') ;
  If ChercheGen(SA_GEN,Lig,TRUE)=0 Then ;
  CGen:=GetGGeneral(GS,Lig) ;
  GS.Cells[SA_Aux,Lig]:='' ;
  GS.Cells[SA_RefI,Lig]:=OBMT[Sens].GetMvt('E_REFINTERNE') ;
  GS.Cells[SA_Lib,Lig]:=OBMT[Sens].GetMvt('E_LIBELLE') ;
  LDD:=0 ; LCD:=0 ; LDP:=0 ; LCP:=0 ; LDS:=0 ; LCS:=0 ;
  if Sens=1 then BEGIN LDS:=MSaisi ; LDD:=MDevise ; LDP:=MPivot ; END
            else BEGIN LCS:=MSaisi ; LCD:=MDevise ; LCP:=MPivot ; END ;
  GS.Cells[SA_Debit,Lig]:=StrS(LDS,DECDEV) ;
  GS.Cells[SA_Credit,Lig]:=StrS(LCS,DECDEV) ;
  FormatMontant(GS,SA_Debit,Lig,DECDEV) ;
  FormatMontant(GS,SA_Credit,Lig,DECDEV) ;
  AlimObjetMvt(Lig,FALSE,0) ;
  OBM:=GetO(GS,Lig) ;
  OBM.CompS:=True ;
  OBM.PutMvt('E_DEBIT',LDP) ;
  OBM.PutMvt('E_CREDIT',LCP) ;
  OBM.PutMvt('E_DEBITDEV',LDD) ;
  OBM.PutMvt('E_CREDITDEV',LCD) ;

  //If EstLettrable(Lig) in [Monoeche,MultiEche] Then OuvreEche(Lig,GS.Cells[SA_Gen,Lig],'',FALSE,MonoEche,TRUE) ;
  If ModeSaisie<>OnBqe Then BEGIN
    If EstLettrable(Lig) in [Monoeche,MultiEche] Then OuvreEche(Lig,GS.Cells[SA_Gen,Lig],'',FALSE,MonoEche,TRUE) ;
    END
  Else BEGIN
    If Lettrable(CGen)=Monoeche Then OuvreEche(Lig,GS.Cells[SA_Gen,Lig],'',FALSE,MonoEche,TRUE) ;
  END ;

  If Ventilable(CGen,0) Then
    begin
    {$IFDEF OLDVENTIL}
    lNumAxe := 0 ;
    {$ELSE}
    if okScenario
      then bOuvreAnal := CAOuvrirVentil( OBM, Scenario, FInfo, lNumAxe )
      else bOuvreAnal := CAOuvrirVentil( OBM, nil, FInfo, lNumAxe ) ;
    if bOuvreAnal then
    {$ENDIF OLDVENTIL}
      OuvreAnal( Lig, False, False, False, lNumAxe ) ;
    end ;

  GereNewLigneT ;
END ;

Procedure TFSaisieEff.MajGridPiedTreso ;
BEGIN
  If SAJAL=NIL Then Exit ;
  Case SAJAL.TRESO.TypCtr Of
    PiedDC :    BEGIN
                if SC_TotDS<>0 then LigneEnPlusTreso(1,SC_TotDS,SC_TotDP,SC_TotDD);
                if SC_TotCS<>0 then LigneEnPlusTreso(2,SC_TotCS,SC_TotCP,SC_TotCD);
                END ;
    PiedSolde : BEGIN
                if SC_TotDS>SC_TotCS then LigneEnPlusTreso(1,SC_TotDS-SC_TotCS,SC_TotDP-SC_TotCP,SC_TotDD-SC_TotCD);
                if SC_TotDS<SC_TotCS then LigneEnPlusTreso(2,SC_TotCS-SC_TotDS,SC_TotCP-SC_TotDP,SC_TotCD-SC_TotDD);
                END ;
  END ;
END ;
//TR Fin sur toutes ces procédures :

{==========================================================================================}
{========================= Affichages, Positionnements ====================================}
{==========================================================================================}
Procedure TFSaisieEff.AutoSuivant( Suiv : boolean ) ;
Var Cpte : String17 ;
BEGIN
If SAJAl=NIL Then Exit ;
if SAJAL.NbAuto<=0 then Exit ;
if ((GS.Cells[SA_Gen,GS.Row]<>'') and (Not CpteAuto)) then Exit ;
if GS.Cells[SA_Gen,GS.Row]='' then CpteAuto:=True ;
if Suiv then BEGIN if CurAuto>=SAJAL.NbAuto then CurAuto:=1 else Inc(CurAuto) ; END
        else BEGIN if CurAuto<=1 then CurAuto:=SAJAL.NbAuto else Dec(CurAuto) ; END ;
Cpte:=TrouveAuto(SAJAL.COMPTEAUTOMAT,CurAuto) ; if Cpte='' then Exit ;
GS.Cells[SA_Gen,GS.Row]:=Cpte ;
END ;

Procedure TFSaisieEff.SoldelaLigne ( Lig : integer ) ;
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
AjusteLigne(Lig) ;
END ;

Procedure TFSaisieEff.PositionneDevise ( ReInit : boolean ) ;
Var OldIndex : integer ;
    MD : Boolean ;
BEGIN
  OldIndex:=-2 ;
  if ((ReInit) and (ModeOppose)) then OldIndex:=E_DEVISE.ItemIndex ;
  MD:=(SAJAL<>NIL) And (SAJAL.MultiDevise) And (ModeSaisie In [OnChq,OnBqe]) ;
  if MD then E_DEVISE.Enabled:=True ;
  if ((ReInit) and (OldIndex<>-2)) then BEGIN
    E_DEVISE.ItemIndex:=OldIndex ;
    E_DEVISEChange(Nil) ;
  END ;
END ;

procedure TFSaisieEff.AffichePied ;
Var CpteG : String ;
    CpteA : String ;
    DIG   : double ;
    CIG   : double ;
    DIA   : double ;
    CIA   : double ;
    OModR : TMOD ;
    OBM   : TOBM ;
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
    FInfo.Compte.Solde( DIG, CIG, GeneTypeExo ) ;   // Calcul du solde en fonction de l'année de référence   // FQ 13246 : SBO 30/03/2005
    LSA_SoldeG.Visible:=True ;
    end
  else LSA_SoldeG.Visible:=False ;

  // Infos Tiers
  if FInfo.LoadAux( GS.Cells[SA_Aux,GS.Row] ) then
    begin
    T_LIBELLE.Caption := FInfo.Aux_GetValue('T_LIBELLE') ;
    CpteA := FInfo.Aux_GetValue('T_AUXILIAIRE') ;
    FInfo.Aux.Solde( DIA, CIA, GeneTypeExo ) ; // Calcul du solde en fonction de l'année de référence   // FQ 13246 : SBO 30/03/2005
    LSA_SoldeT.Visible:=True ;
    if (FInfo.Aux_GetValue('YTC_ACCELERATEUR')='X') then
      T_LIBELLE.Caption := T_LIBELLE.Caption + ' ' + TraduireMemoire('Accélérateur') + ' ' + VarToStr(FInfo.Aux_GetValue('YTC_SCHEMAGEN')) + ' (AUTOMATIQUE) ' ;
    end
  else LSA_SoldeT.Visible:=False ;

  // Affichage des soldes cumulés
  CalculSoldeCompte( CpteG, CpteA, DIG, CIG, DIA, CIA ) ;

  // Spécifique Saisie de trésorerie
  OMODR:=TMOD(GS.Objects[SA_NumP,GS.Row]) ;
  if (OMODR<>Nil) and (OMODR.MODR.NbECHE>0) then
    begin
    HE_MODEP.Caption:=RechDom('ttModePaie',OMODR.MODR.TabEche[1].ModePaie,FALSE) ;
    HE_DATEECHE.Caption:=DateToStr(OMODR.MODR.TabEche[1].DateEche) ;
    end
  else
    begin
    HE_MODEP.Caption:='' ;
    HE_DATEECHE.Caption:='' ;
    end ;
  CalculSoldeCompteT ;
  OBM := GetO(GS,GS.Row) ;
  if OBM<>Nil then
    begin
    if OBM.GetMvt('E_RIB')<>''
      then E_RIB.Caption:=HDiv.Mess[14]+' '+OBM.GetMvt('E_RIB') // RIB :
      else E_RIB.Caption:='' ;
    end
  else E_RIB.Caption:='' ;

end ;

procedure TFSaisieEff.GereNewLigneT ;
Var OBM : TOBM ;
BEGIN //TR
  If SAJAL.TRESO.TypCtr=Ligne Then BEGIN
    if EstRempli(GS,GS.RowCount-1) then BEGIN
      NewLigne(GS) ;
      IncRef ;
      END
    else if Not EstRempli(GS,GS.RowCount-2) then BEGIN
      OBM:=GetO(GS,GS.RowCount-2) ;
      If (OBM<>NIL) And (OBM.Status=DansTable) Then
                                               else GS.RowCount:=GS.RowCount-1 ;
    END ;
    END
  else BEGIN
    if EstRempli(GS,GS.RowCount-1) then BEGIN
      NewLigne(GS) ;
      IncRef ;
      END
    else if Not EstRempli(GS,GS.RowCount-2) then BEGIN
      OBM:=GetO(GS,GS.RowCount-2) ;
      If (OBM<>NIL) And (OBM.Status=DansTable) Then
                                               Else GS.RowCount:=GS.RowCount-1 ;
    END ;
  END ;
END ;

{==========================================================================================}
{================================ Methodes de la form =====================================}
{==========================================================================================}
procedure TFSaisieEff.FormKeyPress(Sender: TObject; var Key: Char);
begin
If MajEnCours Or (Not GS.SynEnabled) Then Key:=#0 Else
if Key=#127 then Key:=#0 else
if ((Key='=') and ((GS.Col=SA_Debit) or (GS.Col=SA_Credit) or (GS.Col=SA_Gen) or (GS.Col=SA_Aux))) then Key:=#0 ;
end;

Procedure TFSaisieEff.FocusSurNatP ;
begin
  if E_JOURNAL.Focused And VH^.JalLookUp then begin
    if E_CONTREPARTIEGEN.CanFocus then E_CONTREPARTIEGEN.SetFocus;
  end;
end;

procedure TFSaisieEff.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var OkG,Vide : boolean ;
    keys : TKeyboardState;
begin
{Panel de Swap devise}
If MajEnCours Or (Not GS.SynEnabled) Then BEGIN Key:=0 ; Exit ; END ;
if ModeRA then
   BEGIN
   Case Key of
      VK_ESCAPE : BRAAbandonClick(Nil) ;
      VK_F10    : BRAValideClick(Nil) ;
//      VK_F1     : BCAideClick(Nil) ;
      END ;
   Exit ;
   END ;
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
  VK_F3     : if ((OkG) and (Vide) and (GS.Col=SA_Gen)) then BEGIN FocusSurNatP ; Key:=0 ; AutoSuivant(False) ; END ;
  VK_F4     : if ((OkG) and (Vide) and (GS.Col=SA_Gen)) then BEGIN FocusSurNatP ;Key:=0 ; AutoSuivant(True) ; END ;
  VK_F5     :   // BPY le 05/10/2004 et le 29/10/2004 : gestion du lookup std a la place du thcompte
                if ((OkG) and ((Vide) or (Shift=[ssCtrl]))) then
                begin
                    Key := 0;
//                    if ((GetParamSoc('SO_CPCODEAUXIONLY',false) = 'X') and (Shift = [ssCtrl])) then Shift := [];
                    if ((GetParamSoc('SO_CPCODEAUXIONLY',false) = true) and (GS.Col = SA_Aux) and (Shift = [ssCtrl])) then
                    begin
                        GetKeyboardState(Keys);
                        keys[VK_CONTROL] := 0;
                        keys[VK_LCONTROL] := 0;
                        keys[VK_RCONTROL] := 0;
                        SetKeyboardState(Keys);
                    end;
                    ClickZoom(true);
                end;
                // Fin BPY
  VK_F6     : if ((OkG) and (Vide)) then
                 BEGIN
                 Key:=0 ; ClickZoomFact ;
                 END Else If Shift=[ssCtrl] Then
                 BEGIN
                 Key:=0 ; ClickZoomFactConsult ;
                 END ;
  VK_F8     : {if Vide then BEGIN Key:=0 ; ClickControleTva(False,0) ; END else}
              if Shift=[ssAlt] then BEGIN Key:=0 ; ClickSwapPivot ; END else
               if Shift=[ssShift] then BEGIN Key:=0 ; ClickSwapEuro ; END ;
  VK_F10    : if Vide then BEGIN Key:=0 ; FocusSurNatP ; ClickValide ; END ;
  VK_ESCAPE : if Vide then begin
                {JP 14/05/07 : FQ 19254 : le Key à zéro est essentiel si on ne veut pas que le
                               Echapp se propage et qu'il y ait une demande de sortie de l'application}
                Key := 0;
                ClickAbandon;
              end;
  VK_RETURN : if ((OkG) and (Vide)) then KEY:=VK_TAB ;
  VK_BACK   : if ((OkG) and (Shift=[ssCtrl])) then BEGIN Key:=0 ; VideZone(GS) ; END ;
  VK_DELETE : if Shift = [ssCtrl] then // FQ 17382 SBO 12/04/2006
                 begin {JP 20/10/05 : Ajout du ssCtrl}
                 Key := 0;
                 ClickRAZLigne;
                 end;
{^Home}  36 : if Shift=[ssCtrl] then BEGIN Key:=0 ; GotoEntete ; END ;
{A1.9}49..57: if Shift=[ssAlt] then BEGIN AppelAuto(Key-48) ; Key:=0 ; END ;
//{AA}     65 : if ((OkG) and (Shift=[ssAlt])) then BEGIN Key:=0 ; ClickVentil ; END ;
{^A}     65 : if ((OkG) and (Shift=[ssCtrl])) then BEGIN Key:=0 ; ClickVentil ; END ;
{AC}     67 : if Shift=[ssAlt] then BEGIN Key:=0 ; ClickComplement ; END ;
{AE}     69 : if ((OkG) and (Shift=[ssAlt])) then BEGIN Key:=0 ; ClickEche ; END ;
{^F}     70 : if Shift=[ssCtrl] then BEGIN Key:=0 ; ClickCherche ; END ;
{AH}     72 : if Shift=[ssAlt] then BEGIN Key:=0 ; END ;
{AI}     73 : if Shift=[ssAlt] then BEGIN Key:=0 ; END ;
{AJ}     74 : if Shift=[ssAlt] then BEGIN FocusSurNatP ;Key:=0 ; ClickJournal ; END ;
(* //TR A voir
{AL}     76 : if ((OkG) and (Shift=[ssAlt])) then BEGIN Key:=0 ; CalculRefAuto(Nil) ; END ;
*)
{^R}     82 : if Shift=[ssCtrl]  then BEGIN Key:=0 ; END else
{AR}             if Shift=[ssAlt]  then BEGIN Key:=0 ; ClickModifRIB ; END ;
{AS}     83 : if Shift=[ssAlt] then BEGIN FocusSurNatP ; Key:=0 ; ClickScenario ; END ;
{AT}     84 : if Shift=[ssAlt] then BEGIN Key:=0 ; ClickEtabl ; END ;
{AV}     86 : if Shift=[ssAlt]  then BEGIN Key:=0 ; END ;
{^D}     68 : if Shift=[ssCtrl] then BEGIN Key:=0 ; END ;
{^L}     76 : if Shift=[ssCtrl] then BEGIN Key:=0 ;  END Else
{AL}            if Shift=[ssAlt]  then BEGIN Key:=0 ; ClickLibelleAuto ; END ;
  END ;
end;

procedure TFSaisieEff.PosLesSoldes ;
BEGIN
  LE_DEBITTRESO.SetBounds(E_DEBITTRESO.Left,E_DEBITTRESO.Top,E_DEBITTRESO.Width,E_DEBITTRESO.Height);
  LE_CREDITTRESO.SetBounds(E_CREDITTRESO.Left,E_CREDITTRESO.Top,E_CREDITTRESO.Width,E_CREDITTRESO.Height);
  E_DEBITTRESO.Visible:=False;  LE_DEBITTRESO.Caption:='' ;
  E_CREDITTRESO.Visible:=False; LE_CREDITTRESO.Caption:='' ;

  LSA_SoldeG.SetBounds(SA_SoldeG.Left,SA_SoldeG.Top,SA_SoldeG.Width,SA_SoldeG.Height) ;
  LSA_SoldeT.SetBounds(SA_SoldeT.Left,SA_SoldeT.Top,SA_SoldeT.Width,SA_SoldeT.Height) ;
  SA_SOLDEG.Visible:=False; LSA_SOLDEG.Caption:='' ;
  SA_SOLDET.Visible:=False; LSA_SOLDET.Caption:='' ;

  LSA_TotalDebit.SetBounds(SA_TotalDebit.Left,SA_TotalDebit.Top,SA_TotalDebit.Width,SA_TotalDebit.Height) ;
  LSA_TotalCredit.SetBounds(SA_TotalCredit.Left,SA_TotalCredit.Top,SA_TotalCredit.Width,SA_TotalCredit.Height) ;
  SA_TotalDebit.Visible:=False;  LSA_TotalDebit.Caption:='' ;
  SA_TotalCredit.Visible:=False; LSA_TotalCredit.Caption:='' ;

  LSA_SoldeTR.SetBounds(SA_SoldeTR.Left,SA_SoldeTR.Top,SA_SoldeTR.Width,SA_SoldeTR.Height) ;
  SA_SOLDE.Visible:=False;
  SA_SOLDETR.Visible:=False;
  LSA_SOLDETR.Caption:='' ;
END ;

Procedure TFSaisieEff.ReinitGrid ;
Var i : Integer ;
    OBM : TOBM ;
    StDebit,StCredit : String ;
BEGIN
  StDebit:='E_DEBITDEV' ; StCredit:='E_CREDITDEV' ;
  For i:=1 To GS.RowCount-1 Do BEGIN
    OBM:=GetO(GS,i) ;
    If OBM<>NIL Then BEGIN
      If SA_Exo>0   Then GS.Cells[SA_Exo,i]:=OBM.GetMvt('E_EXERCICE') ;
      If SA_NumP>0  Then GS.Cells[SA_NumP,i]:=IntToStr(OBM.GetMvt('E_NUMEROPIECE')) ;
      If SA_DateC>0 Then GS.Cells[SA_DateC,i]:=DateToJour(OBM.GetMvt('E_DATECOMPTABLE')) ;
      If SA_NumL>0  Then GS.Cells[SA_NumL,i]:=IntToStr(i) ;
      If SA_Gen>0   Then GS.Cells[SA_Gen,i]:=OBM.GetMvt('E_GENERAL') ;
      If SA_Aux>0   Then GS.Cells[SA_Aux,i]:=OBM.GetMvt('E_AUXILIAIRE') ;
      If SA_RefI>0  Then GS.Cells[SA_RefI,i]:=OBM.GetMvt('E_REFINTERNE') ;
      If SA_Lib>0   Then GS.Cells[SA_Lib,i]:=OBM.GetMvt('E_LIBELLE') ;
      If SA_Debit>0 Then BEGIN
        GS.Cells[SA_Debit,i]:=StrS(OBM.GetMvt(StDebit),DECDEV) ; FormatMontant(GS,SA_Debit,i,DECDEV) ;
      END ;
      If SA_Credit>0 Then BEGIN
        GS.Cells[SA_Credit,i]:=StrS(OBM.GetMvt(StCredit),DECDEV) ; FormatMontant(GS,SA_Credit,i,DECDEV) ;
      END ;
      If SA_ModeP>0 Then BEGIN
        GS.CellValues[SA_ModeP,i]:=OBM.GetMvt('E_MODEPAIE') ;
        If GS.CellValues[SA_ModeP,i]='' Then GS.CellValues[SA_ModeP,i]:=E_MODEPAIE.Value ;
      END ;
      If SA_DateEche>0  Then GS.Cells[SA_DateEche,i]:=DateToStr(DateEcheDefaut) ;
      If SA_RIB>0       Then GS.Cells[SA_RIB,i]:=OBM.GetMvt('E_RIB') ;
      If SA_RefE>0      Then GS.Cells[SA_RefE,i]:=OBM.GetMvt('E_REFEXTERNE') ;
      If SA_RefL>0      Then GS.Cells[SA_RefL,i]:=OBM.GetMvt('E_REFLIBRE') ;
      If SA_NumTraChq>0 Then GS.Cells[SA_NumTraChq,i]:=OBM.GetMvt('E_NUMTRAITECHQ') ;
    END ;
  END ;
END ;

Function TFSaisieEff.PlusModePaie : String ;
BEGIN
  Result:='' ;
  Case ModeSaisie Of
    OnEffet : Result:='(MP_ENCAISSEMENT="MIX" OR MP_ENCAISSEMENT="ENC") AND MP_CATEGORIE="LCR" ' ;
    OnChq : Result:='(MP_ENCAISSEMENT="MIX" OR MP_ENCAISSEMENT="ENC") AND MP_CATEGORIE="CHQ" ' ;
    OnCB : Result:='(MP_ENCAISSEMENT="MIX" OR MP_ENCAISSEMENT="ENC") AND MP_CATEGORIE="CB" ' ;
    OnBqe : Case ModeSR Of
              srCli : Result:='(MP_ENCAISSEMENT="MIX" OR MP_ENCAISSEMENT="ENC") ' ;
              srFou : Result:='(MP_ENCAISSEMENT="MIX" OR MP_ENCAISSEMENT="DEC") ' ;
            END ;
  END ;
END ;

procedure TFSaisieEff.ReinitNumCol(St : String) ;
Var St1 : String ;
    i : Integer ;
    StPlus : String ;
BEGIN
  StPlus:=PlusModePaie ;
  i:=0 ;
  SA_ModeP:=-1 ; SA_DateEche:=-1 ; SA_RIB:=-1 ; SA_RefE:=-1 ; SA_RefL:=-1 ; SA_NumTraChq:=-1 ;
  While St<>'' Do BEGIN
    St1:=ReadTokenSt(St) ;
    If St1='E_EXERCICE' Then SA_Exo:=i Else
      If St1='E_NUMEROPIECE' Then SA_NumP:=i Else
        If St1='E_NUMLIGNE' Then SA_NumL:=i Else
          If St1='E_DATECOMPTABLE' Then SA_DateC:=i Else
            If St1='E_GENERAL' Then SA_Gen:=i Else
              If St1='E_AUXILIAIRE' Then SA_Aux:=i Else
                If St1='E_REFINTERNE' Then SA_RefI:=i Else
                  If St1='E_LIBELLE' Then SA_Lib:=i Else
                    If St1='E_DEBIT' Then SA_Debit:=i Else
                      If St1='E_CREDIT' Then SA_Credit:=i Else
                        If St1='E_MODEPAIE' Then BEGIN
                          SA_ModeP:=i ;
                          GS.ColFormats[SA_ModeP]:='CB=TTMODEPAIE|'+StPlus ; // Récupérer code par GS.CellValues[ACol,ARow]
                          END
                        Else If St1='E_DATEECHEANCE' Then BEGIN
                          SA_DateEche:=i ;
                          GS.ColFormats[SA_DateEche]:=ShortDateFormat ;
                          GS.ColTypes[SA_DateEche]:='D' ;
                          END
                        Else If St1='E_RIB' Then SA_RIB:=i Else
                          If St1='E_REFEXTERNE' Then SA_RefE:=i Else
                            If St1='E_REFLIBRE' Then SA_RefL:=i Else
                              If St1='E_NUMTRAITECHQ' Then SA_NumTraChq:=i ;
    Inc(i) ;
  END ;
  SauveSAEFF ;
  ReinitGrid ;
END ;

Function TFSaisieEff.QuelleListe : String ;
BEGIN
  Result:='CPSAISIEEFFA' ;
  Case VEFFET Of
    OnEffet : Result:='CPSAISIEEFFA' ;
    OnChq   : Result:='CPSAISIEEFFB' ;
    OnCB    : Result:='CPSAISIEEFFC' ;
    OnBqe   : If VMode=srCli Then Result:='CPSAISIEEFFD'
                             Else Result:='CPSAISIEEFFE' ;
  END ;
END ;

procedure TFSaisieEff.FormCreate(Sender: TObject);
Var i : Integer ;
begin
  WMinX:=Width ;
  WMinY:=Height_Tres ;
  DesalloueLesOBML ;
  GS.VidePile(True) ;
  GS.RowCount:=2 ;
  E_NUMREF.Visible:=FALSE ;
  GS.TypeSais:=tsTreso ;
  LigneTreso:=FALSE ;
//  GS.ListeParam:='CPSAISIEEFFA' ;
  GS.ListeParam:='' ;
  GS.ListeParam:=QuelleListe ;
  GS.FixedCols:=3 ;
  If (GS.Titres.Count>0) And (GS.Tag=1) Then ReinitNumCol(GS.Titres[0]) ;
  SA_PILESAI := 'EFF;' + SA_PILESAI ;
  SAJAL:=Nil ;
  DecDev:=V_PGI.OkdecV ;
  TS:=TList.Create ;
  TPIECE:=TList.Create ;
  TDELECR:=TList.Create ;
  TDELANA:=TList.Create ;
  For i:=1 To MaxAxe Do
    LAnaEff[i]:= HTStringList.Create ;
  Scenario:=TOBM.Create(EcrSais,'',True) ;
  Entete:=TOBM.Create(EcrSais,'',True) ;
  MemoComp:= HTStringList.Create ;
  FInfo := TInfoEcriture.Create ; // FQ 13246 : SBO 30/03/2005
  ModifSerie:=TOBM.Create(EcrGen,'',True) ;
  PieceConf:=False ;
  GestionRIB:='...' ;
  DejaRentre:=False ;
  OkScenario:=False ;
  ModeConf:='0' ;
  Volatile:=False ;
  NeedEdition:=False ;
  EtatSaisie:='' ;
  RentreDate:=False ;
  RegimeForce:=False ;
  SorteTva:=stvDivers ;
  ExigeTva:=tvaMixte ;
  ExigeEntete:=tvaMixte ;
  ChoixExige:=False ;
  RegLoadToolbarPos(Self,'SaisieEff') ;
  PosLesSoldes ;
{Param Liste}
  GS.ColLengths[SA_Gen]:=VH^.Cpta[fbGene].Lg ;
  if VH^.CPCodeAuxiOnly then GS.ColLengths[SA_Aux]:=VH^.Cpta[fbAux].Lg
                        else GS.ColLengths[SA_Aux]:=35 ;
  GS.ColLengths[SA_RefI]:=35 ;
  GS.ColLengths[SA_Lib]:=35 ;
  GS.PostDrawCell:=PostDrawCell ;
  FClosing := False; {JP 14/05/07 : FQ 19254}
  FTobModePaie := TOB.Create('_MP', nil, -1); {JP 14/05/07 : FQ 18324}
  // SBO 01/07/2007 : enregistrement spécifique pour pb CWAS
  FBoInsertSpecif := EstSpecif('51215') ;
end;

procedure TFSaisieEff.PostDrawCell(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
var R : TRect ;
    TT : TList ;
BEGIN
  if ACol<>SA_Credit then Exit ;
  if GS.Cells[SA_Credit,ARow]='' then Exit ;
  TT:=Tlist(GS.Objects[SA_NumL,ARow]  ) ;
  if TT=Nil then Exit ;
  if TT.Count<=0 then Exit ;
  Canvas.Brush.Color:= clBlue ;
  Canvas.Brush.Style:=bsSolid ;
  Canvas.Pen.Color:=clBlue ;
  Canvas.Pen.Mode:=pmCopy ;
  Canvas.Pen.Width:= 1 ;
  R:=GS.CellRect(ACol,ARow) ;
  Canvas.Rectangle(R.Left+1,R.Top+1,R.Left+5,R.Top+5);
END ;

procedure TFSaisieEff.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  CanClose:=FermerSaisie ;
  FClosing := True; {JP 14/05/07 : FQ 19254}
end;

Procedure TFSaisieEff.EditionSaisie ;
Var SWhere : String ;
    i      : integer ;
    RR     : RMVT ;
BEGIN
  if Not NeedEdition then Exit ;
  if EtatSaisie='' then Exit ;
  // Voulez-vous éditer le document interne des écritures créées ?
  if HDiv.Execute(22,Caption,'')<>mrYes then Exit ;
  NeedEdition:=False ;
  SWhere:='' ;
  for i:=0 to TPIECE.Count-1 do BEGIN
    RR:=P_MV(TPIECE[i]).R ;
    SWhere:=SWhere+'(E_JOURNAL="'+RR.Jal+'" AND E_NUMEROPIECE='+IntToStr(RR.Num)+' AND E_QUALIFPIECE="'+RR.Simul+'" AND E_EXERCICE="'+RR.Exo+'")' ;
    if i<TPIECE.Count-1 then SWhere:=SWhere+' OR ' ;
  END ;
  if TPIECE.Count>1 then SWhere:='('+SWhere+')' ;
  LanceEtat('E','SAI',EtatSaisie,True,False,False,Nil,SWhere,'',False) ;
END ;

procedure TFSaisieEff.FormClose(Sender: TObject; var Action: TCloseAction);
Var i : Integer ;
    St : String ;
begin
  DesalloueLesOBML ; GS.VidePile(True) ; St:='' ;
  Case ModeSaisie Of
    OnEffet : St:='EFC' ;
    OnChq   : St:='CHQ' ;
    OnCB    : St:='CBL' ;
  End ;
  If SAJAL<>NIL Then SauveJalEffet(St,SAJAL.Journal) ;
  SAJAL.Free ;
  VideListe(TS);      TS.Free;      TS:=Nil ;
  VideListe(TPIECE);  TPIECE.Free;  TPIECE:=Nil ;
  VideListe(TDELECR); TDELECR.Free; TDELECR:=Nil ;
  VideListe(TDELANA); TDELANA.Free; TDELANA:=Nil ;
  FreeAndNil(FInfo) ; // FQ 13246 : SBO 30/03/2005

  For i:=1 To MaxAxe Do BEGIN
    LAnaEff[i].Clear ;
    LAnaEff[i].Free ;
    LanaEff[i]:=Nil ;
  END ;
  Scenario.Free ;
  Entete.Free ;
  ModifSerie.Free ;
  MemoComp.Free ;
  MemoComp:=Nil ;
  OBMT[1].Free ;
  OBMT[2].Free ; //TR
  PurgePopup(POPS) ;
  PurgePopup(POPZ) ;
  PurgePopup(POPZT) ;
//  SA_DateC:=2 ;
//  SA_NumL:=3 ;

  if Assigned(FTobModePaie) then FreeAndNil(FTobModePaie);{JP 14/05/07 : FQ 18324}

  InitDefautColSaisie ;
  REadTokenSt( SA_PILESAI ) ;
  RegSaveToolbarPos(Self,'SaisieTr') ;
  if Parent is THPanel then BEGIN
    Case Self.Action of
      taCreat : Bloqueur('nrSaisieCreat',False) ;
      taModif : Bloqueur('nrSaisieModif',False) ;
    END ;
    Action:=caFree ;
  END ;
end;

procedure TFSaisieEff.FormShow(Sender: TObject);
Var bc : boolean ;
    JalSauve,Qui : String ;
    Ind : Integer ;
begin
  If ModeSaisie<>OnBqe Then E_NATUREPIECE.Value:='OD' ;
  LookLesDocks(Self) ;
  MajEnCours:=FALSE ;
  //SA_DateC:=3 ; SA_NumL:=2 ;
  if ((Action=taCreat) and (M.Etabl='')) then M.Etabl:=VH^.EtablisDefaut ;
  InitLesComposants ;
  E_REF.TEXT:='' ;
  E_NUMDEP.text:='' ;
  PieceConf:=False ;
  E_MODEPAIE.ItemIndex:=0 ;
  ModeCG:=False ;
  ModeRA:=FALSE ;
  DefautPied ;
  DefautEntete ;
  NbLigEcr:=0 ;
  NbLigAna:=0 ;
  If E_JOURNAL.Values.Count>0 Then BEGIN
    Qui:='' ;
    JalSauve:='' ;
    Ind:=0 ;
    Case ModeSaisie Of
      OnEffet : Qui:='EFC' ;
      OnChq : Qui:='CHQ' ;
      OnCB : Qui:='CBL' ;
    End ;
    If Qui<>'' Then JalSauve:=ChargeJalEffet(Qui) ;
    If JalSauve<>'' then BEGIN
      Ind:=E_JOURNAL.Values.IndexOf(JalSauve) ;
      If Ind<0 Then Ind:=0 ;
    END ;
    E_JOURNAL.Value:=E_JOURNAL.Values[Ind] ;
    E_JOURNAL.Refresh ;
    E_JOURNALChange(NIL) ;
  END ;
  ChargeLignes ;
  AffecteGrid(GS,Action) ;
  AfficheLeSolde(SA_Solde,0,0) ;
  AttribLeTitre ;
  QuelNExo ;
  GSRowEnter(Nil,GS.Row,bc,False) ;
end;

{==========================================================================================}
{================================ Ecriture Fichier=========================================}
{==========================================================================================}
procedure TFSaisieEff.MajCptesGeneNew ;
Var XD,XC : Double ;
    Lig   : integer ;
    FRM : TFRM ;
    ll : LongInt ;
BEGIN
if PasModif then Exit ;
for Lig:=1 to GS.RowCount-2 do
  if ((GS.Cells[SA_Gen,Lig]<>'') and (GS.Objects[SA_Gen,Lig]<>Nil)) then
    BEGIN
    Fillchar(FRM,SizeOf(FRM),#0) ;
    FRM.Cpt:=GS.Cells[SA_Gen,Lig] ;
    RecupTotalPivot( GS, Lig, XD, XC ) ;// SBO 10/07/2007 FQ 20910 maj solde des comptes erroné en devise
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
    if Action<>taConsult then AttribParamsComp(FRM,GS.Objects[SA_Gen,Lig]) ;
    LL:=ExecReqMAJ(fbGene,M.ANouveau,Action<>taConsult,FRM) ;
    If ll<>1 then V_PGI.IoError:=oeSaisie ;
    END ;
END ;

procedure TFSaisieEff.MajCptesAuxNew ;
Var XD,XC : Double ;
    Lig   : integer ;
    Trouv : Boolean ;
    FRM : TFRM ;
    ll : LongInt ;
BEGIN
if PasModif then Exit ;
Trouv:=False ;
for Lig:=1 to GS.RowCount-2 do if ((GS.Cells[SA_Aux,Lig]<>'') and (GS.Objects[SA_Aux,Lig]<>Nil))
    then BEGIN Trouv:=True ; Break ; END ;
if Not Trouv then Exit ;

for Lig:=1 to GS.RowCount-2 do
  if ((GS.Cells[SA_Aux,Lig]<>'') and (GS.Objects[SA_Aux,Lig]<>Nil)) then
    BEGIN
    Fillchar(FRM,SizeOf(FRM),#0) ;
    FRM.Cpt:=GS.Cells[SA_Aux,Lig] ;
    RecupTotalPivot( GS, Lig, XD, XC ) ;// SBO 10/07/2007 FQ 20910 maj solde des comptes erroné en devise
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
    if Action<>taConsult then AttribParamsComp(FRM,GS.Objects[SA_Aux,Lig]) ;
    ll:=ExecReqMAJ(fbAux,M.ANouveau,Action<>taConsult,FRM) ;
    If ll<>1 then V_PGI.IoError:=oeSaisie ;
    END ;
END ;

procedure TFSaisieEff.MajCptesSectionNew ;
Var Sect   : String ;
    NumAxe,NumV,Lig : integer ;
    OBM    : TOBM;
    OBA    : TOB ;
    XC,XD  : Double ;
    Trouv  : boolean ;
    FRM : TFRM ;
BEGIN
if PasModif then Exit ;
Trouv:=False ;
for Lig:=1 to GS.RowCount-2 do
  if GS.Cells[SA_Gen,Lig]<>'' then BEGIN
    OBM := GetO(GS,LIG);
    if OBM<>Nil then BEGIN
      Trouv:=VentilationExiste(OBM);
      Break ;
    END ;
  END ;
  if Not Trouv then Exit ;

  // Parcours des écritures
  for Lig:=1 to GS.RowCount-2 do
    if GS.Cells[SA_Gen,Lig]<>'' then BEGIN
      OBM := GetO(GS,Lig);
      if OBM=Nil then Continue ;

      // Parcours des axes
      for NumAxe:=1 to OBM.Detail.Count do
        for NumV:=0 to OBM.Detail[NumAxe-1].Detail.Count-1 do BEGIN
          OBA 	:= OBM.Detail[NumAxe-1].Detail[NumV];
          Sect	:= OBA.GetValue('Y_SECTION') ;
          XD		:= OBA.GetValue('Y_DEBIT') ;
          XC		:= OBA.GetValue('Y_CREDIT') ;
          Fillchar(FRM,SizeOf(FRM),#0) ;
          FRM.Cpt:= Sect ;
          FRM.Axe:= 'A'+Chr(48+NumAxe) ;
          if Not M.ANouveau then BEGIN
            FRM.NumD:=NumPieceInt ;
            FRM.DateD:=DateMvt(Lig) ;
            FRM.LigD:=StrToInt(GS.Cells[SA_NumL,Lig]) ;
            AttribParamsNew(FRM,XD,XC,GeneTypeExo) ;
            END
          else BEGIN
            FRM.Deb:=XD ;
            FRM.Cre:=XC ;
          END ;
          ExecReqMAJ(fbSect,M.ANouveau,False,FRM) ;
        END ;
    END ;
END ;

function TFSaisieEff.QuelEnc ( Lig : Integer ) : String3 ;
var CGen : TGGeneral ;
begin
  CGen:=GetGGeneral(GS,Lig) ;
  // Pour les comptes lettrables divers... FQ 16136 SBO 09/08/2005
  if ( CGen <> Nil ) and CGen.Lettrable and ( CGen.NatureGene = 'DIV' )
    then result := 'RIE'
    // Sinon récultat habituel
    else Result := SensEnc(ValD(GS,Lig),ValC(GS,Lig)) ;
end ;

Procedure TFSaisieEff.RecupTronc(Lig : Integer ; MO : TMOD ; Var OBM : TOBM);
Var CGen : TGGeneral ;
    CAux : TGTiers ;
    NatG,NatT : String3 ;
    OBA  : TOB ;
    TVA,TPF : String3 ;
    NumAxe,NumV : integer ;
BEGIN
  if PasModif then Exit ;
  CGen:=NIL ; CAux:=NIL ; NatG:='' ; NatT:='' ; OBM:=Nil ;
  if GS.Objects[SA_Gen,Lig]<>Nil then BEGIN CGen:=TGGeneral(GS.Objects[SA_Gen,Lig]) ; NatG:=CGen.NatureGene ; END ;
  if GS.Objects[SA_Aux,Lig]<>Nil then BEGIN CAux:=TGTiers(GS.Objects[SA_Aux,Lig]) ; NatT:=CAux.NatureAux ; END ;
  if GS.Objects[SA_Exo,Lig]<>Nil then OBM:=GetO(GS,Lig) ;
  if OBM = nil then Exit;

  If OBM<>Nil then BEGIN
    OBM.SetCotation(0) ;
    {JP 14/05/07 : FQ 18324 : gestion du codeacceptation}
    OBM.PutMvt('E_CODEACCEPT',GetCodeAccept(OBM.GetString('E_MODEPAIE')));
//    TMemoField(TEcrGen.FindField('E_BLOCNOTE')).Assign(OBM.M) ;  // A FAIRE
  END ;

  // GGGGG
  OBM.PutMvt('E_QUALIFPIECE',M.Simul);

  // Modif Lettrage des comptes divers  FQ 16136 SBO 09/08/2005
  CGetTypeMvt( OBM, FInfo ) ;

  if OBM.GetMvt('E_REGIMETVA') = '' then OBM.PutMvt('E_REGIMETVA' , GeneRegTVA);
  if Not M.ANouveau then BEGIN
    OBM.PutMvt('E_ECRANOUVEAU', 'N');
    OBM.PutMvt('E_VALIDE', '-');
    END
  else BEGIN
    OBM.PutMvt('E_ECRANOUVEAU', 'H');
    OBM.PutMvt('E_VALIDE', 'X');
  END ;
  OBM.PutMvt('E_BUDGET', '');
  OBM.PutMvt('E_ECHE', '-');
  if CAux<>Nil then OBM.PutMvt('E_REGIMETVA', CAux.RegimeTVA);
  if CGen<>Nil then BEGIN
    TVA:=CGen.Tva ; if TVA='' then TVA:=GeneTVA ;
    TPF:=CGen.Tpf ; if TPF='' then TPF:=GeneTPF ;
    if OBM.GetMvt('E_TVA') = '' then OBM.PutMvt('E_TVA', GeneTva);
    if OBM.GetMvt('E_TPF') = '' then OBM.PutMvt('E_TPF', GeneTpf);
    OBM.PutMvt('E_BUDGET', CGen.Budgene);
    if ((Lettrable(CGen)=MonoEche) and (MO<>Nil)) then BEGIN

      // récupération de la 1ère échéance
      RecupEche( Lig, MO.MODR , 1 , OBM ) ; // FQ 18736 SBO 07/09/2006 : pb d'affectation des zones de lettrage pour les comptes divers lettrables

      // Ajout Sinon bug LCD FQ 16584 SBO 06/08/2005
      if CGen.pointable then
        OBM.PutMvt('E_ETATLETTRAGE',	'RI') ;  // FQ 16761 SBO 30/09/2005

    END ;
  END ;

  if GS.Cells[SA_Gen,Lig]<>SAJAL.TRESO.Cpt then BEGIN
    OBM.PutMvt('E_CONTREPARTIEGEN', SAJAL.TRESO.Cpt);
    OBM.PutMvt('E_CONTREPARTIEAUX', '');
//    If SAJAL.TRESO.Collectif Then OBM.PutMvt('E_CONTREPARTIEAUX', GS.Cells[SA_Aux,Lig]);
    If (ModeSaisie<>OnBqe) And (SAJAL.TRESO.TypCtr=Ligne) And SAJAL.TRESO.Collectif Then OBM.PutMvt('E_CONTREPARTIEAUX', GS.Cells[SA_Aux,Lig]);
    END
  else BEGIN
    Case SAJAL.TRESO.TypCtr Of
      PiedDC,PiedSolde,AuChoix :
             BEGIN
               OBM.PutMvt('E_CONTREPARTIEGEN', GS.Cells[SA_Gen,1]);
               OBM.PutMvt('E_CONTREPARTIEAUX', GS.Cells[SA_Aux,1]);
             END ;
      Ligne : BEGIN
               OBM.PutMvt('E_CONTREPARTIEGEN', GS.Cells[SA_Gen,Lig-1]);
               OBM.PutMvt('E_CONTREPARTIEAUX', GS.Cells[SA_Aux,Lig-1]);
             END ;
    END ;
  END ;
  OBM.PutMvt('E_DATEMODIF', NowFutur);
  OBM.PutMvt('E_CONFIDENTIEL', ModeConf);

  {Mise à jour contreparties analytiques}
  if OBM.Detail.Count>0 then
  	for NumAxe:=0 to OBM.Detail.Count-1 do
      for NumV:=0 to OBM.Detail[NumAxe].Detail.Count-1 do BEGIN
        OBA := OBM.Detail[NumAxe].Detail[NumV];
        OBA.PutValue('Y_CONTREPARTIEGEN',OBM.GetMvt('E_CONTREPARTIEGEN')) ;
        OBA.PutValue('Y_CONTREPARTIEAUX',OBM.GetMvt('E_CONTREPARTIEAUX')) ;
      END ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 27/03/2002
Modifié le ... :   /  /    
Description .. : LG* nouvelle fonction de saisUtil suite au modif du tobm
Mots clefs ... : 
*****************************************************************}
Procedure TFSaisieEff.RecupAnal(Lig : integer ; var OBA : TOB ; NumAxe,NumV : integer ) ;
BEGIN
  // Attention NumV démarre à 0
  if PasModif then Exit ;
  //LG* 27/03/2002 nv fonction de SaisUtil
  if NumV=0 then OBA.PutValue('Y_TYPEMVT', 'AE')
            else OBA.PutValue('Y_TYPEMVT', 'AL');
  OBA.PutValue('Y_TYPEANALYTIQUE', '-');
//  TMemoField(TEcrAna.FindField('Y_BLOCNOTE')).Assign(TT.M[NumV]) ; // A FAIRE
  OBA.PutValue('Y_DATEMODIF', NowFutur);
  OBA.PutValue('Y_DATECOMPTABLE', DatePiece);
{$IFNDEF SPEC302}
  OBA.PutValue('Y_PERIODE', GetPeriode(DatePiece));
  OBA.PutValue('Y_SEMAINE', NumSemaine(DatePiece));
{$ENDIF}
  OBA.PutValue('Y_NATUREPIECE', E_NATUREPIECE.Value);
  OBA.PutValue('Y_ETABLISSEMENT', E_ETABLISSEMENT.Value);
  OBA.PutValue('Y_CONFIDENTIEL', ModeConf);
  OBA.PutValue('Y_ECRANOUVEAU', 'N');
  OBA.PutValue('Y_QUALIFPIECE', 'N'); // SBO 12/04/2006 : Y_QUALIFPIECE à N Pour l'analytique FQ 17456
END ;

Procedure TFSaisieEff.RecupEche(Lig : integer ; R : T_ModeRegl ; NumEche : integer; Var OBM : TOBM);
Var Deb  : boolean ;
    i    : Integer ;
    Coll : String ;
BEGIN
  if PasModif then Exit ;
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
  (* déjà fait dans AlimEcheVersObm suite à saisrMonoEche :
  OBM.PutMvt('E_MODEPAIE'),R.TabEche[NumEche].ModePaie ;
  OBM.PutMvt('E_DATEECHEANCE',R.TabEche[NumEche].DateEche ;
  OBM.PutMvt('E_DATEVALEUR').asDateTime:=R.TabEche[NumEche].DateValeur ;
  *)
  OBM.PutMvt('E_DEBIT',0) ;  		OBM.PutMvt('E_CREDIT',0) ;
  OBM.PutMvt('E_DEBITDEV',0) ;	OBM.PutMvt('E_CREDITDEV',0) ;
  Deb:=(ValD(GS,Lig)<>0) ;
  if Deb then BEGIN
    OBM.PutMvt('E_DEBIT',     R.TabEche[NumEche].MontantP);
    OBM.PutMvt('E_DEBITDEV',  R.TabEche[NumEche].MontantD);
    END
  else BEGIN
    OBM.PutMvt('E_CREDIT',     R.TabEche[NumEche].MontantP);
    OBM.PutMvt('E_CREDITDEV',  R.TabEche[NumEche].MontantD);
  END ;
  if ModeDevise then BEGIN
    if Deb then OBM.PutMvt('E_DEBITDEV',  R.TabEche[NumEche].MontantD)
           else OBM.PutMvt('E_CREDITDEV', R.TabEche[NumEche].MontantD);
  END ;
  OBM.PutMvt('E_ENCAISSEMENT',    QuelEnc(Lig));
  OBM.PutMvt('E_ORIGINEPAIEMENT', R.TabEche[NumEche].DateEche);
//SetMPACCDB(TEcrGen) ;
  {JP 14/05/07 : FQ 18324 : gestion du codeacceptation}
  OBM.PutMvt('E_CODEACCEPT', GetCodeAccept(R.TabEche[NumEche].ModePaie));

  {#TVAENC}
  if VH^.OuiTvaEnc then
  BEGIN
    Coll := OBM.GetMvt('E_GENERAL');
    if EstCollFact(Coll) then
    BEGIN
      if ((SorteTva=stvDivers) and (R.ModifTva)) then
      BEGIN
        for i:=1 to 4 do
            OBM.PutMvt('E_ECHEENC'+IntToStr(i), Arrondi(R.TabEche[NumEche].TAV[i],V_PGI.OkDecV));

        OBM.PutMvt('E_ECHEDEBIT', Arrondi(R.TabEche[NumEche].TAV[5],V_PGI.OkDecV));
      END ;
    END ;

   //YMOO Recopie des infos TVA sur les lignes de contrepartie
   If OBM.GetMvt('E_GENERAL') = E_CONTREPARTIEGEN.Text then
   begin
     for i:=1 to 4 do
          OBM.PutMvt('E_ECHEENC'+IntToStr(i), Arrondi(R.TabEche[NumEche].TAV[i],V_PGI.OkDecV));

     OBM.PutMvt('E_ECHEDEBIT', Arrondi(R.TabEche[NumEche].TAV[5],V_PGI.OkDecV));

     OBM.PutMvt('E_TVA', R.TabEche[NumEche].CodeTva);
   end;

  END ;
END ;

Procedure TFSaisieEff.GetAna(Lig : Integer) ;
Var i,j : integer ;
    OBM : TOBM ;
    OBA : TOB ;
BEGIN
  if PasModif then exit ;
  OBM := GetO(GS,Lig) ;

	if OBM = Nil then Exit;
  for i:=0 to OBM.Detail.Count-1 do
    for j:=0 to OBM.Detail[i].Detail.Count-1 do BEGIN
      OBA := OBM.Detail[i].Detail[j];
      RecupAnal(Lig,OBA,i,j) ;
      // SBO 01/07/2007 : enregistrement spécifique pour pb CWAS
      if FBoInsertSpecif
        then begin
             if not CTobInsertDB( OBA ) then
               V_PGI.IOError := oeSaisie ;
             end
        else OBA.InsertDB(nil);
    END ;
END ;

Procedure TFSaisieEff.GetEcr(Lig : Integer) ;
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
  if ((M=Nil) or (Lettrable(CGen)=MonoEche)) then BEGIN
    // Récupération et remplissage TOB Ecritures générales
    RecupTronc(Lig,M,OBM);
    // Finalisation des TOBs Analytiques
    GetAna(Lig) ;

    {JP 27/10/05 : Gestion éventuelle des écritures de Tréso}
    OnUpdateEcritureTOB(OBM, Action, []);

    {Enregistrement Base}
      if FBoInsertSpecif  // SBO 01/07/2007 : enregistrement spécifique pour pb CWAS
        then begin
             if not CTobInsertDB( OBM, True, False, False ) then
               V_PGI.IOError := oeSaisie ;
              end
        // ??? Pourquoi on ne fait pas du niveau 0 uniquement comme dans eSaisieTr.pas...sinon violation d'accès
        else OBM.InsertDBByNivel(False, 0, 0); // OBM.InsertDBByNivel(False) ;

  END
  else BEGIN
    R:=M.MODR ;
    for i:=1 to R.NbEche do BEGIN
      // Récupération et remplissage TOB Ecritures générales
      RecupTronc(Lig,M,OBM);
      RecupEche(Lig,R,i,OBM);

      // Finalisation des TOBs Analytiques
      if i = 1 then GetAna(Lig) ;

      {JP 27/10/05 : Gestion éventuelle des écritures de Tréso}
      OnUpdateEcritureTOB(OBM, Action, []);
      {Enregistrement Base}
      if FBoInsertSpecif  // SBO 01/07/2007 : enregistrement spécifique pour pb CWAS
        then begin
             if not CTobInsertDB( OBM, True, False, False ) then
               V_PGI.IOError := oeSaisie ;
             end
        // ??? Pourquoi on ne fait pas du niveau 0 uniquement comm dans eSaisieTr.pas...sinon violation d'accès
        else OBM.InsertDBByNivel(False, 0, 0); // OBM.InsertDBByNivel(False);
    END ;
  END ;
END ;

procedure TFSaisieEff.ValideLignesGene ;
Var i : integer ;
BEGIN
  if PasModif then Exit ;
  LAUX:=TrouveLigTiers(GS,0) ;
  if LAUX<0 then LAUX:=1 ;
  LHT:=TrouveLigHT(GS,0,False) ;
  if LHT<0 then BEGIN
    if LAUX<>2 then LHT:=2
               else LHT:=1 ;
  END ;
  InitMove(GS.RowCount-2,'') ;

  for i:=1 to GS.RowCount-2 do BEGIN
    MoveCur(FALSE);
    GetEcr(i) ;
    if V_PGI.IOError<>oeOK then Break ;
  END ;
  FiniMove ;
END ;

procedure TFSaisieEff.ValideLeJournalNew ;
Var Per  : Byte ;
    DD   : TDateTime ;
    FRM : TFRM ;
    ll : LongInt ;
BEGIN
if PasModif then Exit ;
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
ll:=ExecReqMAJ(fbJal,M.ANouveau,False,FRM) ;
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
// A FAIRE Voir UtilSOC.PAS
{$ELSE}
   MarquerPublifi(True) ;
{$ENDIF}
END ;

procedure TFSaisieEff.ValideLesComptes ;
BEGIN
  if PasModif then exit;

  MajCptesGeneNew;
  MajCptesAuxNew ;
  MajCptesSectionNew ;
END ;

procedure TFSaisieEff.ValideLaPiece ;
BEGIN
if Not GS.SynEnabled then Exit ;
GridEna(GS,False) ;
if V_PGI.IOError=oeOK then ValideLignesGene ;
END ;

procedure TFSaisieEff.ValideLeReste ;
BEGIN
V_PGI.IOError:=oeOK ;
if V_PGI.IOError=oeOK then ValideLesComptes ;
if V_PGI.IOError=oeOK then ValideLeJournalNew;
GridEna(GS,True) ;
END ;


{==========================================================================================}
{================================ Barre d'outils ==========================================}
{==========================================================================================}
procedure TFSaisieEff.ForcerMode ( Cons : boolean ; Key : Word ) ;
BEGIN
if ((Cons) and (Modeforce=tmDevise)) then Exit ;
if ((Not Cons) and (Modeforce=tmNormal)) then Exit ;
if Cons then
   BEGIN
   AffecteGrid(GS,taConsult) ; ModeForce:=tmDevise ; GS.SetFocus ;
   PEntete.Enabled:=False ; Outils.Enabled:=False ;
   PForce.Align:=alClient ; PForce.Parent:=PPied ; PForce.Visible:=True ;
   AfficheLeSolde(HForce,SI_TotDP,SI_TotCP) ;
   END else
   BEGIN
   if Key=VK_RETURN then SoldeLaligne(GS.Row) ;
   PEntete.Enabled:=True ; Outils.Enabled:=True ;
   if Outils.CanFocus then Outils.SetFocus ;
   GS.CacheEdit ; AffecteGrid(GS,Action) ;
   GS.Col:=SA_Gen ; if GS.Row=1 then GS.Row:=2 else GS.Row:=1 ;
   GS.SetFocus ; GS.MontreEdit ;
   ModeForce:=tmNormal ; PForce.Visible:=False ;
   END ;
END ;

Function TFSaisieEff.FabricFromM : RMVT ;
Var X : RMVT ;
BEGIN
X.Axe:='' ; X.Etabl:=E_ETABLISSEMENT.Value ;
X.Jal:=E_JOURNAL.Value ; X.Exo:=QuelExoDT(StrToDate(E_DATECOMPTABLE.Text)) ;
X.CodeD:=E_DEVISE.Value ; X.Simul:='N' ; X.Nature:=E_NATUREPIECE.Value ;
X.DateC:=StrToDate(E_DATECOMPTABLE.Text) ; X.DateTaux:=DEV.DateTaux ;
X.Num:=NumPieceInt ; X.TauxD:=DEV.Taux ; X.Valide:=False ; X.ANouveau:=M.ANouveau ;
Result:=X ;
END ;

procedure TFSaisieEff.StockeLaPiece ;
Var X : RMVT ;
    P : P_MV ;
BEGIN
  if Action<>taCreat then Exit ;
  FillChar(X,Sizeof(X),#0) ;
  X:=FabricFromM ;
  P:=P_MV.Create ;
  P.R:=X ;
  TPIECE.Add(P) ;
END ;

Procedure PrepareLettrageEffet(TLett : TList ; O : TOBM) ;
Var L : TL_Rappro ;
BEGIN
L:=TL_Rappro.Create ;
L.General:=O.GetMvt('E_GENERAL')     ; L.Auxiliaire:=O.GetMvt('E_AUXILIAIRE') ;
L.DateC:=O.GetMvt('E_DATECOMPTABLE') ; L.DateE:=O.GetMvt('E_DATEECHEANCE') ; L.DateR:=O.GetMvt('E_DATEREFEXTERNE') ;
L.RefI:=O.GetMvt('E_REFINTERNE')     ; L.RefL:=O.GetMvt('E_REFLIBRE') ;
L.RefE:=O.GetMvt('E_REFEXTERNE')     ; L.Lib:=O.GetMvt('E_LIBELLE') ;
L.Jal:=O.GetMvt('E_JOURNAL')         ; L.Numero:=O.GetMvt('E_NUMEROPIECE') ;
L.NumLigne:=O.GetMvt('E_NUMLIGNE')   ; L.NumEche:=O.GetMvt('E_NUMECHE') ;
L.CodeL:=O.GetMvt('E_LETTRAGE')      ; L.TauxDEV:=O.GetMvt('E_TAUXDEV') ;
L.CodeD:=O.GetMvt('E_DEVISE')        ; L.Debit:=O.GetMvt('E_DEBIT') ;
L.Credit:=O.GetMvt('E_CREDIT')       ; L.DebDev:=O.GetMvt('E_DEBITDEV') ;
L.CredDev:=O.GetMvt('E_CREDITDEV')   ; L.DebitCur:=O.GetMvt('E_DEBIT') ;
L.CreditCur:=O.GetMvt('E_CREDIT')    ; L.Nature:=O.GetMvt('E_NATUREPIECE') ;
L.Facture:=((L.Nature='FC') or (L.Nature='FF') or (L.Nature='AC') or (L.Nature='AF')) ;
L.Client:=((L.Nature='FC') or (L.Nature='AC') or (L.Nature='RC') or (L.Nature='OC')) ;
L.Solution:=0 ; L.Exo:=O.GetMvt('E_EXERCICE') ;
L.EditeEtatTva:=O.GetMvt('E_EDITEETATTVA')='X' ;
L.Decim:=V_PGI.OkDecV ;
TLett.Add(L) ;
END ;

Procedure TFSaisieEff.LettrageEffet ;
Var Lig,i : Integer ;
    LOBM,TLett : TList ;
    O,OO : TOBM ;
    St : String ;
    CodeL : String ;
    Total : Boolean ;
BEGIN
  ResteLettrageAFaire:=FALSE ;
  TLett:=TList.Create ;
  for Lig:=1 to GS.RowCount-2 do BEGIN
    LOBM:=TList(GS.Objects[SA_NumL,Lig]) ;
    O:=GetO(GS,Lig) ;
    If LOBM<>NIL Then BEGIN
      VideListe(TLett) ;
      For i:=0 To LOBM.Count-1 Do BEGIN
        OO:=LOBM[i] ;
        PrepareLettrageEffet(TLett,OO) ;
      END ;
      If LOBM.Count<=0 Then BEGIN
        ResteLettrageAFaire:=TRUE ;
        Continue ;
      END ;
      If O<>Nil Then BEGIN
        PrepareLettrageEffet(TLett,O) ;
        Total:=TRUE ;
        St:=LettrerUnPaquet(TLett,False,False,'','',FALSE,TRUE) ;
        If (St<>'') And (St[1]='X') Then BEGIN
          Total:=FALSE ;
          RegulAFaire:=TRUE ;
          AlimTableTemporaire(O,'REG',smpAucun,Copy(St,3,4),'') ;
          ResteLettrageAFaire:=TRUE ;
          END
        Else BEGIN
          If St<>'' Then BEGIN
            CodeL:=Copy(St,5,1) ;
            Total:=CodeL[1]='T' ;
          END ;
        END ;
        If Not Total Then ResteLettrageAFaire:=TRUE ;
      END ;
      END
    Else BEGIN
      If O<>NIL Then BEGIN
        If O.GetMvt('E_GENERAL')<>SAJAL.Treso.Cpt Then ResteLettrageAFaire:=TRUE ;
      END ;
    END ;
  END ;
  VideListe(TLett) ;
  TLett.Free ;
END ;

procedure TFSaisieEff.ClickValide ;
Var StE     : String3 ;
    io,io1  : TIOErr ;
    OBM     : TOBM ;
    i       : integer ;
BEGIN
if Not GS.SynEnabled then Exit ;
If SAJAL=NIL Then Exit ;
if ((SaJal.Journal<>'') and (GS.RowCount<=2) and (GS.CanFocus) and
    (Screen.ActiveControl<>GS) and (Screen.ActiveControl<>Valide97)) then
    BEGIN
    If OkPourLigne Then GS.SetFocus ; Exit ;
    END ;
{Contrôles avant validation}
if Action=taConsult then BEGIN CloseFen ; Exit ; END ;
if Not BValide.Enabled then Exit ;
ValideCeQuePeux(GS.Row) ;
If M.treso And (SAJal.Treso.TypCtr In [PiedDC,PiedSolde]) Then
   BEGIN
   if EstRempli(GS,GS.Row) Then
      If Not LigneCorrecte(GS.Row,TRUE,FALSE,TRUE) then Exit ;
   END ;
if not PieceCorrecte(TRUE) then Exit ;
if ((EstRempli(GS,GS.Row)) and (Not LigneCorrecte(GS.Row,False,True,FALSE))) then Exit ;
If M.Treso Then
   BEGIN
   If SAJAL<>NIL Then If (SAJAL.TRESO.TypCtr In [PiedDC,PiedSolde]) Then MajGridPiedTreso ;
   END ;
if not PieceCorrecte(FALSE) then Exit ;
MajIoCom(GS) ;
GS.Row:=1 ; GS.Col:=SA_Gen ; GS.SetFocus ; BValide.Enabled:=False ;
StE:=QuelExoDT(StrToDate(E_DATECOMPTABLE.Text)) ;
if StE=VH^.Encours.Code then GeneTypeExo:=teEncours else
   if StE=VH^.Suivant.Code then GeneTypeExo:=teSuivant else GeneTypeExo:=tePrecedent ;
{Scenario avant validation}
{Validation}
DatePiece:=StrToDate(E_DATECOMPTABLE.Text) ;
If Not M.MajDirecte Then NowFutur:=NowH ;
GS.Enabled:=False ;
io:=Transactions(ValideLaPiece,10) ;
io1:=Transactions(ValideLeReste,10) ;
// ATTENTION : Des conflits d'accès ont été détectés. #13
// La pièce est validée, mais vous devez demander à l'administrateur de lancer un recalcul du solde des comptes.
If io1<>oeOK Then MessageAlerte(HDiv.Mess[27]+HDiv.Mess[28]) ;
  //SG6 21/12/2004 FQ 14731 Vérification quantité saisie
  for i := 1 to GS.RowCount - 1 do
  begin
    OBM := GetO(GS, i) ;
    if ( OBM <> nil ) and ( OBM.GetValue('E_ANA')='X' )
                      and ( ( OBM.GetNumChamp('CHECKQTE') < 0 ) or
                            ( OBM.GetString('CHECKQTE')<>'X'  )
                          ) then
      if not CheckQuantite( OBM, False ) then
        begin
        PGIInfo('Certaines lignes de la pièce présentent des incohérences entre les quantités de la ligne d''écriture et les quantités de lignes d''écritures analytiques.');
        break ;
        end ;
  end;
GridEna(GS,True) ; GS.Enabled:=True ;
Case io of
    oeOK  : BEGIN
            StockeLaPiece ;
            RegulAFaire:=FALSE ;
            ResteLettrageAFaire:=FALSE ;
            Transactions(LettrageEffet,3) ;
            If EstSerie(S3) Then RegulAFaire:=FALSE ;
            If RegulAFaire Then RegulLettrageMPAuto (TRUE,FALSE) ;
            if ResteLettrageAFaire then LettrageEnSaisie ;
            END ;
oeSaisie  : MessageAlerte(HDiv.Mess[4]); // Validation impossible. Certaines écritures ont été modifiées par un autre utilisateur.
oeUnknown : MessageAlerte(HDiv.Mess[3]); // ATTENTION : Pièce non validée
   END ;
{Ré-initialisation}
if (Action=taCreat) then ReInitPiece ;
END ;

procedure TFSaisieEff.ClickVentil ;
Var OkL : Boolean ;
BEGIN
if GS.Row>=GS.RowCount-1 then Exit ;
if Not EstRempli(GS,GS.Row) then Exit ;
OkL:=LigneCorrecte(GS.Row,False,True,True) ;
if OkL then GereAnal(GS.Row,False,False,FALSE,FALSE) ;
END ;

procedure TFSaisieEff.ClickEche ;
Var OkL : Boolean ;
BEGIN
if GS.Row>=GS.RowCount-1 then Exit ;
if Not EstRempli(GS,GS.Row) then Exit ;
OkL:=LigneCorrecte(GS.Row,False,True,TRUE) ;
if OkL then
   BEGIN
   GereEche(GS.Row,False,False) ;
   If SAJAL.TRESO.TypCtr=Ligne Then AlimLigneTreso(GS.Row) ;
   END ;
END ;

procedure TFSaisieEff.ClickAbandon ;
BEGIN
  if ModeForce=tmPivot then
    ClickSwapPivot
  else BEGIN
    CloseFen ;
  END ;
END ;

procedure TFSaisieEff.CloseFen ;
begin
  {JP 14/05/07 : FQ 19254 : Reprise du code de la saisie}
  Close;

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

{Function TFSaisieEff.TrouveLigneTreso : Integer ;
Var i : Integer ;
BEGIN
Result:=0 ;
If SAJAL=NIL Then Exit ;
For i:=1 To GS.RowCount-2 Do
  BEGIN
  If GS.Cells[SA_GEN,i]=SAJAL.Treso.Cpt Then BEGIN Result:=i ; Exit ; END ;
  END ;
END ;}

procedure TFSaisieEff.ClickZoom(DoPopup : boolean = false);
Var b : byte ;
    C,R : longint ;
    A   : TActionFiche ;
BEGIN
if ((E_JOURNAL.Value='') or (Not GS.Enabled)) then Exit ;
If SAJAL=NIL Then Exit ;
if ((Action=taConsult) or (ModeForce<>tmNormal))
    then BEGIN GS.MouseToCell(GX,GY,C,R) ; A:=taConsult ; END
    else BEGIN R:=GS.Row ; C:=GS.Col ; A:=taModif ; END ;
if R<1 then Exit ; if C<SA_Gen then Exit ;
if ((Action=taConsult) and (GS.Cells[C,R]='')) then Exit ;
if ModeForce<>tmNormal then Exit ;
if C=SA_Gen then BEGIN
   {$IFNDEF GCGC}
   if (Not ExJaiLeDroitConcept(TConcept(ccGenModif),False)) then A:=taConsult ;
   b:=ChercheGen(C,R,False,DoPopup) ;
   if b=2 then BEGIN
     FicheGene(Nil,'',GS.Cells[C,R],A,0) ;
     if Action<>taConsult then ChercheGen(C,R,True) ;
   END ;
   {$ELSE}
   Exit ;
   {$ENDIF}
   END else if C=SA_Aux then BEGIN
   if (Not ExJaiLeDroitConcept(TConcept(ccAuxModif),False)) then A:=taConsult ;
   b:=ChercheAux(C,R,False) ;
   {$IFNDEF GCGC}
   if b=2 then BEGIN
     FicheTiers(Nil,'',GS.Cells[C,R],A,1) ;
     if Action<>taConsult then ChercheAux(C,R,True) ; END ;
   {$ELSE}
   if b=2 then AGLLanceFiche('GC','GCTIERS','',GS.Cells[C,R],ActionToString(taConsult)+';MONOFICHE;T_NATUREAUXI=CLI') ;
   {$ENDIF}
   END Else If C=SA_RIB Then
   BEGIN
   ClickModifRib ;
   END ;
END ;

Procedure TFSaisieEff.RecupAnaSel(O : TOBM) ;
Var QAna : tQuery ;
    Ind,i : Integer ;
    Axe : String ;
    Section : String ;
    Taux : Double ;
BEGIN
For i:= 1 To MaxAxe Do LAnaEff[i].Clear ;
QAna:=OpenSQL('Select Y_SECTION,Y_POURCENTAGE,Y_AXE from ANALYTIQ where Y_JOURNAL="'+O.GetMvt('E_JOURNAL')+'"'
          +' AND Y_EXERCICE="'+QuelExo(DateToStr(O.GetMvt('E_DATECOMPTABLE')))+'"'
          +' AND Y_DATECOMPTABLE="'+USDATETIME(O.GetMvt('E_DATECOMPTABLE'))+'"'
          +' AND Y_NUMEROPIECE='+IntToStr(O.GetMvt('E_NUMEROPIECE'))
          +' AND Y_NUMLIGNE='+IntToStr(O.GetMvt('E_NUMLIGNE'))
          +' AND Y_GENERAL="'+O.GetMvt('E_GENERAL')+'" '
//          +' AND Y_AUXILIAIRE="'+O.GetMvt('E_AUXILIAIRE')+'" '
          +' AND Y_QUALIFPIECE="N" ORDER BY Y_AXE,Y_NUMVENTIL',True) ;
While Not QAna.Eof Do BEGIN
  Axe:=QAna.FindField('Y_AXE').AsString ; Ind:=0 ;
  If (Length(Axe)>1) And (Axe[2] in ['1'..'5']) Then Ind:=StrToInt(Copy(Axe,2,1)) ;
  If Ind>0 Then
    BEGIN
    Section:=QAna.FindField('Y_SECTION').AsString ;
    Taux:=QAna.FindField('Y_POURCENTAGE').AsFloat ;
    MajLL(LAnaEff[Ind],Section,Taux) ;
    END ;
  QAna.Next ;
  END ;
Ferme(QAna) ;
END ;

Function TFSaisieEff.recupMontantSel(LOBM : TList ; Var M : Double; Var RIB : String ; OkAna : Boolean) : Boolean ;
Var O : TOBM ;
    D,C,Couv : Double ;
    i : integer ;
    LeRib : String ;
BEGIN
  RIB:='' ;
  Result:=FALSE ;
  D:=0 ; C:=0 ; Couv:=0 ; M:=0 ;
  If LOBM=NIL Then Exit ;
  If LOBM.Count<=0 Then Exit ;
  For i:=0 To LOBM.Count-1 Do BEGIN
    O:=TOBM(LOBM[i]) ;
    If OkAna Then RecupAnaSel(O) ;
    D:=Arrondi(D+O.GetMvt('E_DEBIT'),V_PGI.OkDecV) ;
    C:=Arrondi(C+O.GetMvt('E_CREDIT'),V_PGI.OkDecE) ;
    Couv:=Arrondi(Couv+O.GetMvt('E_COUVERTURE'),V_PGI.OkDecV) ;
    LeRIB:=O.GetMvt('E_RIB') ;
    If (Rib='') And (LeRib<>'') Then Rib:=LeRib ;
  END ;
  Result:=TRUE ;
  If ModeSr=srFou Then M:=Arrondi(C-D-Couv,V_PGI.OkDecV) 
  Else M:=Arrondi(D-C-Couv,V_PGI.OkDecV) ;

  If M<0 Then BEGIN
    Result:=FALSE ;
    Rib:='' ;
  END ;
END ;

Function TFSaisieEff.MvtCommun(LOBM,LOBM1 : TList) : Boolean ;
Var i, j : Integer ;
    O,O1 : TOBM ;
    LesMemes : Boolean ;
BEGIN
  Result:=FALSE ;
  For i:=0 To LOBM.Count-1 Do For j:=0 To LOBM1.Count-1 Do BEGIN
    O:=TOBM(LOBM[i]) ;
    O1:=TOBM(LOBM1[j]) ;
    LesMemes:=FALSE ;
    If (O.GetMvt('E_JOURNAL')=O1.GetMvt('E_JOURNAL')) And
       (O.GetMvt('E_DATECOMPTABLE')=O1.GetMvt('E_DATECOMPTABLE')) And
       (O.GetMvt('E_EXERCICE')=O1.GetMvt('E_EXERCICE')) And
       (O.GetMvt('E_NUMEROPIECE')=O1.GetMvt('E_NUMEROPIECE')) And
       (O.GetMvt('E_NUMLIGNE')=O1.GetMvt('E_NUMLIGNE')) And
       (O.GetMvt('E_NUMECHE')=O1.GetMvt('E_NUMECHE')) Then LesMemes:=TRUE ;
    If LesMemes Then BEGIN Result:=TRUE ; Exit ; END ;
  END ;
END ;

Function TFSaisieEff.VerifAutreLigne ( Col,Lig : integer ) : Boolean ;
Var i,j : integer ;
    TabL : Array[1..200] Of Integer ;
    LOBM,LOBM1 : TList ;
    Cpte : String ;
BEGIN
Result:=TRUE ; Exit ;
  If GS.Cells[SA_Aux,Lig]<>'' Then Cpte:=GS.Cells[SA_Aux,Lig]
                              Else Cpte:=GS.Cells[SA_Gen,Lig] ;
  if Cpte='' then Exit ;
  Fillchar(TabL,SizeOf(TabL),#0) ;
  j:=1 ;
  for i:=1 to GS.RowCount-2 do
    if (i<>Lig) And (GS.Cells[SA_AUX,i]=Cpte) And (GS.Objects[SA_NumL,i]<>NIL) then BEGIN
      LOBM:=TList(GS.Objects[SA_NumL,Lig]) ;
      If LOBM.Count>0 Then BEGIN
        TabL[j]:=i ;
        Inc(j) ;
      END ;
    END ;
  LOBM:=TList(GS.Objects[SA_NumL,Lig]) ;
  If j>1 Then
    For i:=1 To j-1 Do BEGIN
      LOBM1:=TList(GS.Objects[SA_NumL,TabL[i]]) ;
      If MvtCommun(LOBM,LOBM1) Then BEGIN
        // Parmi les éléments sélectionnés, certains ont déjà été pris en compte sur une autre ligne.
        HDiv.Execute(30,'','(Ligne '+IntToStr(TabL[i])+')') ;
        Result:=FALSE ;
        Break ;
      END ;
    END ;
END ;

procedure TFSaisieEff.ClickZoomFactConsult ;
Var LOBM : TList ;
BEGIN
  if ((E_JOURNAL.Value='') or (Not GS.Enabled)) then Exit ;
  If SAJAL=NIL Then Exit ;
  if ModeForce<>tmNormal then Exit ;
  LOBM:=TList(GS.Objects[SA_NumL,GS.Row]) ;
  If LOBM=NIL Then Exit ;
  IF LOBM.Count = 0 then Exit;
  ZoomFactCliEff(LOBM) ;
END ;

{Function TFSaisieEff.PrevientSiSelection(LOBM : tlist) : Boolean ;
BEGIN
Result:=TRUE ;
If LOBM=NIL Then Exit ;
If LOBM.Count<=0 Then Exit ;
// ATTENTION : vous avez déjà seléctionné des écritures pour cette ligne (Ctrl-F6 pour voir la sélection).
HDiv.execute(31,Caption,'') ;
// ATTENTION : La sélection va être annulée. Confirmez-vous ?
If HDiv.execute(32,Caption,'')<>mrYes Then BEGIN Result:=FALSE ; Exit ; END ;
VideListe(LOBM) ;
END ;}

Procedure O2O1(O,O1 : TOBM ; St : String ; LeType : Char) ;
BEGIN
Case LeType Of
  'S' : If O.GetMvt(St)='' Then O.PutMvt(St,O1.GetMvt(St)) ;
  'D' : If O.GetMvt(St)=iDate1900 Then O.PutMvt(St,O1.GetMvt(St)) ;
  'F' : If Arrondi(O.GetMvt(St),4)=0 Then O.PutMvt(St,O1.GetMvt(St)) ;
  END ;
END ;

procedure TFSaisieEff.DupliqueLigne(LOBM : Tlist ; O : TOBM ; Lig : Integer) ;
Var O1 : TOBM ;
    St : String ;
BEGIN
//Exit ;
O1:=TOBM(LOBM[0]) ;
O2O1(O,O1,'E_REFINTERNE','S') ;
O2O1(O,O1,'E_LIBELLE','S') ;
O2O1(O,O1,'E_DATEREFEXTERNE','D') ;
O2O1(O,O1,'E_AFFAIRE','S') ;
O2O1(O,O1,'E_LIBRETEXTE0','S') ;
O2O1(O,O1,'E_LIBRETEXTE1','S') ;
O2O1(O,O1,'E_LIBRETEXTE2','S') ;
O2O1(O,O1,'E_LIBRETEXTE3','S') ;
O2O1(O,O1,'E_LIBRETEXTE4','S') ;
O2O1(O,O1,'E_LIBRETEXTE5','S') ;
O2O1(O,O1,'E_LIBRETEXTE6','S') ;
O2O1(O,O1,'E_LIBRETEXTE7','S') ;
O2O1(O,O1,'E_LIBRETEXTE8','S') ;
O2O1(O,O1,'E_LIBRETEXTE9','S') ;
O2O1(O,O1,'E_TABLE0','S') ;
O2O1(O,O1,'E_TABLE1','S') ;
O2O1(O,O1,'E_TABLE2','S') ;
O2O1(O,O1,'E_TABLE3','S') ;
O2O1(O,O1,'E_LIBREDATE','D') ;
O2O1(O,O1,'E_LIBREBOOL0','S') ;
O2O1(O,O1,'E_LIBREBOOL1','S') ;
O2O1(O,O1,'E_LIBREMONTANT0','F') ;
O2O1(O,O1,'E_LIBREMONTANT1','F') ;
O2O1(O,O1,'E_LIBREMONTANT2','F') ;
O2O1(O,O1,'E_LIBREMONTANT3','F') ;
O2O1(O,O1,'E_CONSO','S') ;
St:=O.GetMvt('E_REFINTERNE') ; If St<>'' Then GS.Cells[SA_RefI,Lig]:=St ;
St:=O.GetMvt('E_LIBELLE') ;    If St<>'' Then GS.Cells[SA_LIB,Lig]:=St ;
If SA_REFL>-1 Then BEGIN St:=O.GetMvt('E_REFLIBRE') ;   If St<>'' Then GS.Cells[SA_RefL,Lig]:=St ; END ;
If SA_REFE>-1 Then BEGIN St:=O.GetMvt('E_REFEXTERNE') ; If St<>'' Then GS.Cells[SA_RefE,Lig]:=St ; END ;
END ;

procedure TFSaisieEff.ClickZoomFact ;
Var Lig,Col : longint ;
    Gene,Auxi : String ;
    LOBM,LOBM1 : TList ;
    OkZoom : Boolean ;     M : Double ;

    CGen   : TGGeneral ;
    CAux : TGTiers ;
    LeLib,LeRibOrigine : String ;
    NewFlux : Boolean ;
    O,O3 : TOBM ;
    OkAna : Boolean ;
    RienAuDepart : Boolean ;
    O2,OD : TOB;
    i,j : Integer;
BEGIN
  if ((E_JOURNAL.Value='') or (Not GS.Enabled)) then Exit ;
  If SAJAL=NIL Then Exit ;
  if ModeForce<>tmNormal then Exit ;
  Lig:=GS.Row ;
  Col:=GS.Col ;
  Gene:=GS.Cells[SA_Gen,Lig] ;
  Auxi:=GS.Cells[SA_Aux,Lig] ;
  AlloueOBML(Lig) ;
  LOBM:=TList(GS.Objects[SA_NumL,Lig]) ;
  RienAuDepart:=FALSE ;
  If LOBM<>NIl Then RienAuDepart:=LOBM.Count<=0 ;
//If LOBM=NIL Then BEGIN LOBM:=TList.Create ; GS.Objects[SA_NumL,Lig]:=TObject(LOBM) ; END ;
  LeLib:='' ;
  CGen:=GetGGeneral(GS,Lig) ;
  if CGen=Nil then Exit ;
  OkAna:=Ventilable(CGen,0) ;
  If CGen.Collectif Then BEGIN
    CAux:=GetGTiers(GS,Lig) ;
    if CAux=Nil then Exit ;
    LeLib:=CAux.Libelle ;
    If Not CAux.Lettrable Then Exit ;
    END
  Else BEGIN
    If Not CGen.Lettrable Then Exit ;
    LeLib:=CGen.Libelle ;
  END ;

  NewFlux:=(GS.Cells[SA_DEBIT,Lig]='') And (GS.Cells[SA_CREDIT,Lig]='') ;
//If Not PrevientSiSelection(LOBM) Then Exit ;

  // Mémorise le contenu des autres lignes
  O2 := TOB.Create('Pas de table',nil,-1);

  for i := 1 to GS.RowCount-1 do begin
    if (i = Lig) or ((Gene <> GS.Cells[SA_Gen,i]) and (Auxi <> GS.Cells[SA_Aux,i])) then Continue;
    LOBM1 := TList(GS.Objects[SA_NumL,i]);
    if LOBM1=Nil then Continue;

    for j := 0 to LOBM1.Count-1 do begin
      O3 := TOBM(LOBM1.Items[j]);

      // Création de la tob fille
      OD := TOB.Create('Pas de table',O2,-1);
      OD.AddChampSup('E_JOURNAL', True);
      OD.AddChampSup('E_EXERCICE', True);
      OD.AddChampSup('E_DATECOMPTABLE', True);
      OD.AddChampSup('E_NUMEROPIECE', True);
      OD.AddChampSup('E_NUMLIGNE', True);
      OD.AddChampSup('E_NUMECHE', True);
      OD.AddChampSup('E_QUALIFPIECE', True);

      // Remplissage de la tob fille
      OD.PutValue('E_JOURNAL', O3.GetValue('E_JOURNAL') );
      OD.PutValue('E_EXERCICE', O3.GetValue('E_EXERCICE') );
      OD.PutValue('E_DATECOMPTABLE', O3.GetValue('E_DATECOMPTABLE') );
      OD.PutValue('E_NUMEROPIECE', O3.GetValue('E_NUMEROPIECE') );
      OD.PutValue('E_NUMLIGNE', O3.GetValue('E_NUMLIGNE') );
      OD.PutValue('E_NUMECHE', O3.GetValue('E_NUMECHE') );
      OD.PutValue('E_QUALIFPIECE', O3.GetValue('E_QUALIFPIECE') );
      
    end;
  end;

  // Appel le mul pour la sélection des pièces
  OkZoom:=PointeMvtCpt(Gene,Auxi,LeLib,LOBM,DateEcheDefaut,NewFlux,ModeSaisie,ModeSR,O2) ;

  O2.Free;

  If OkZoom And (LOBM.Count>0) And RecupMontantSel(LOBM,M,LeRibOrigine,OkAna) Then BEGIN
    If VerifAutreLigne(Col,Lig) Then BEGIN
      GS.Cells[SA_Credit,Lig]:=StrS(M,DECDEV) ;
      TraiteMontant(Lig,True) ;
      If (ModeSaisie In [OnBqe,OnEffet]) And (LeRIBOrigine<>'') Then BEGIN
        If (SA_Rib>-1) Then GS.Cells[SA_RIB,Lig]:=Trim(LeRibOrigine) ;
        O:=GetO(GS,Lig) ;
        if O<>Nil then BEGIN
          O.PutMvt('E_RIB',LeRibOrigine) ;
          E_RIB.Caption:=HDiv.Mess[14]+' '+O.GetMvt('E_RIB') ; // RIB :
        END ;
      end;
      If RienAuDepart And (LOBM.Count=1) Then BEGIN
        O:=GetO(GS,Lig) ;
        DupliqueLigne(LOBM,O,Lig) ;
      END ;
    END ;
//  ChercheMontant(Col,Lig) ;
  END ;
 
END ;

procedure TFSaisieEff.ClickSwapPivot ;
Var Cancel,Chg : boolean ;
BEGIN
if ((Not ModeDevise) and (ModeForce=tmNormal)) then Exit ;
if Not BSwapPivot.Enabled then Exit ;
if ((EstRempli(GS,GS.Row)) and (Not LigneCorrecte(GS.Row,False,True,TRUE))) then Exit ;
if Not EstRempli(GS,1) then Exit ;
if ModeForce=tmNormal then
   BEGIN
   ModeForce:=tmPivot ; OldDevise:=DEV.Code ; E_DEVISE.Value:=V_PGI.DevisePivot ;
   DEV.Decimale:=V_PGI.OkDecV ; DECDEV:=V_PGI.OkDecV ; E_DEVISE.Enabled:=True ; E_Devise.Font.Color:=clRed ;
   AffecteGrid(GS,taConsult) ; GS.SetFocus ;
   PEntete.Enabled:=False ; Outils.Enabled:=False ; Valide97.Enabled:=False ;
   END else
   BEGIN
   E_DEVISE.Enabled:=False ; E_DEVISE.Value:=OldDevise ; E_Devise.Font.Color:=clBlack ;
   ModeForce:=tmNormal ;
   PEntete.Enabled:=True ; Outils.Enabled:=True ; Valide97.Enabled:=True ;
   if Outils.CanFocus then Outils.SetFocus ;
   GS.Col:=SA_Gen ; GS.Row:=1 ;
   GS.CacheEdit ; AffecteGrid(GS,Action) ;
   GS.Col:=SA_Gen ; GS.Row:=2 ;
   GS.SetFocus ; GS.MontreEdit ;
   GS.Row:=1 ; Cancel:=False ; Chg:=False ; GSRowEnter(Nil,GS.Row,Cancel,Chg) ;
   END ;
ChangeAffGrid(GS,ModeForce,DECDEV) ;
CalculDebitCredit ;
END ;

procedure TFSaisieEff.ClickSwapEuro ;
BEGIN
if ModeForce<>tmNormal then Exit ;
if Not BSwapEuro.Enabled then Exit ;
if ((EstRempli(GS,GS.Row)) and (Not LigneCorrecte(GS.Row,False,True,TRUE))) then Exit ;
CorrigeEcartConvert(DEV,GS,taConsult) ;
END ;

procedure TFSaisieEff.ClickModifRIB ;
Var O : TOBM ;
//    RibAvant : String ;
    IsAux : Boolean ;
BEGIN
//if Not BModifRIB.Enabled then Exit ;
O:=GetO(GS,GS.Row) ; if O=Nil then Exit ; //RibAvant:=O.GetMvt('E_RIB') ;
IsAux:=O.GetMvt('E_AUXILIAIRE')<>'' ;
ModifRIBOBM(O,False,TRUE,GS.Cells[SA_Aux,GS.Row],IsAux) ;

  If O.GetMvt('E_RIB')<>'' Then E_RIB.Caption:=HDiv.Mess[14]+' '+O.GetMvt('E_RIB') // RIB :
                           Else E_RIB.Caption:='' ;
  If SA_RIB>-1 Then BEGIN
    GS.Cells[SA_RIB,GS.Row]:=Trim(O.GetMvt('E_RIB')) ;
  END ;
END ;

procedure TFSaisieEff.ClickComplement ;
Var ModBN : boolean ;
    O     : TOBM ;
    Lig   : integer ;
    RC    : R_COMP ;
    OQ1,OQ2,NQ1,NQ2 : Double ;
begin
Lig:=GS.Row ;
if Not EstRempli(GS,Lig) then Exit ;
if Not LigneCorrecte(Lig,False,True,TRUE) then Exit ;
O:=GetO(GS,Lig) ; OQ1:=O.GetMvt('E_QTE1') ; OQ2:=O.GetMvt('E_QTE2') ;
RC.StLibre:='XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX' ;
RC.StComporte:='--XXXXXXXX' ;
RC.Conso:=True ; RC.DateC:=DatePiece ;
RC.Attributs:=False ; RC.MemoComp:=Nil ; RC.Origine:=-1 ;
if Not SaisieComplement(O,EcrGen,Action,ModBN,RC) then Exit ;
if GS.CanFocus then GS.SetFocus ;
NQ1:=O.GetMvt('E_QTE1') ; NQ2:=O.GetMvt('E_QTE2') ;
if ((NQ1<>OQ1) and (OQ1<>0)) then ProrateQteAnal(GS,Lig,OQ1,NQ1,1) ;
if ((NQ2<>OQ2) and (OQ2<>0)) then ProrateQteAnal(GS,Lig,OQ2,NQ2,2) ;
end;

procedure TFSaisieEff.AlimOBMMemoireEnBatch(OSource,ODestination : TOBM) ;
begin
If (OSource=NIL) Or (ODestination=NIL) Then Exit ;
ODestination.PutMvt('E_AFFAIRE',OSource.GetMvt('E_AFFAIRE')) ;
ODestination.PutMvt('E_REFEXTERNE',OSource.GetMvt('E_REFEXTERNE')) ;
ODestination.PutMvt('E_DATEREFEXTERNE',OSource.GetMvt('E_DATEREFEXTERNE')) ;
ODestination.PutMvt('E_REFLIBRE',OSource.GetMvt('E_REFLIBRE')) ;
ODestination.PutMvt('E_QTE1',OSource.GetMvt('E_QTE1')) ;
ODestination.PutMvt('E_QTE2',OSource.GetMvt('E_QTE2')) ;
ODestination.PutMvt('E_QUALIFQTE1',OSource.GetMvt('E_QUALIFQTE1')) ;
ODestination.PutMvt('E_QUALIFQTE2',OSource.GetMvt('E_QUALIFQTE2')) ;
ODestination.M.Assign(OSource.M) ;
end ;

procedure TFSaisieEff.ClickComplementT ;
Var LesDeux,Inutile : boolean ;
    O     : TOBM ;
    RC    : R_COMP ;
begin
  if ((E_JOURNAL.Value='') or (SAJAL=Nil)) then Exit ;
  LesDeux:=(OBMT[1]<>NIL) And (OBMT[2]<>NIL) ;
  O:=Nil ;
  if OBMT[1]<>NIL Then O:=OBMT[1] ;
  if OBMT[2]<>NIL Then O:=OBMT[2] ;
  If O=NIL Then Exit ;
  O.PutMvt('E_NUMLIGNE',GS.RowCount-1) ;
  RC.StComporte:='--XXXXXXXX' ;
  RC.StLibre:='XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX' ;
  RC.Conso:=True ;
  RC.Attributs:=False ;
  RC.MemoComp:=Nil ;
  RC.Origine:=-1 ;
  RC.DateC:=DatePiece ;
  if Not SaisieComplement(O,EcrGen,Action,Inutile,RC) then Exit ;
  if GS.CanFocus then GS.SetFocus ;
  If LesDeux Then AlimOBMMemoireEnBatch(OBMT[2],OBMT[1]) ;
end;

procedure TFSaisieEff.ClickCherche ;
BEGIN
if GS.RowCount<=2 then Exit ;
FindFirst:=True ; FindSais.Execute ;
END ;

procedure TFSaisieEff.ClickRAZLigne ;
Var Lig : longint ;
BEGIN
if ((E_JOURNAL.Value='') or (Not GS.Enabled)) then Exit ;
If SAJAL=NIL Then Exit ;
if ((Action<>taCreat) or (ModeForce<>tmNormal)) Then Exit ;
Lig:=GS.Row ; if Lig<=1 then Exit ;
GS.SynEnabled:=False ;
DetruitLigne(GS.Row,True) ;
ControleLignes ;
GS.SynEnabled:=True ;
END ;

procedure TFSaisieEff.ClickScenario ;
var lStLequel : string ;
begin
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
end ;

procedure TFSaisieEff.ClickJournal ;
Var
  a : TActionFiche ;
begin
{$IFNDEF GCGC}
  if ((E_JOURNAL.Value='') or (SAJAL=Nil)) then Exit ;
  a:=taConsult ;
    FicheJournal(Nil,'',E_JOURNAL.Value,a,0) ;
{$ENDIF}
end;

procedure TFSaisieEff.ClickEtabl ;
BEGIN
FicheEtablissement_AGL(taConsult) ;
END ;

procedure TFSaisieEff.BVentilClick(Sender: TObject);
begin ClickVentil ; if GS.Enabled then GS.SetFocus ; end;

procedure TFSaisieEff.BEcheClick(Sender: TObject);
begin ClickEche ; if GS.Enabled then GS.SetFocus ; end;

procedure TFSaisieEff.BLigClick(Sender: TObject);
begin
If GS.RowCount<=2 Then Exit ;
AfficheLigneTreso ; GS.SetFocus ;
end;

procedure TFSaisieEff.BValideClick(Sender: TObject);
begin
if ((M.TypeGuide<>'NOR') or (Not M.FromGuide)) then ClickValide else ClickAbandon ;
end;

procedure TFSaisieEff.BZoomEtablClick(Sender: TObject);
begin ClickEtabl ; end;

procedure TFSaisieEff.BAbandonClick(Sender: TObject);
begin ClickAbandon ; end;

procedure TFSaisieEff.BZoomClick(Sender: TObject);
begin
ClickZoom ; if GS.Enabled then GS.SetFocus ;
end;

procedure TFSaisieEff.BScenarioClick(Sender: TObject);
begin ClickScenario ; end;

procedure TFSaisieEff.BZoomJournalClick(Sender: TObject);
begin ClickJournal ; end;

procedure TFSaisieEff.BSwapPivotClick(Sender: TObject);
begin ClickSwapPivot ; end;

procedure TFSaisieEff.BSwapEuroClick(Sender: TObject);
begin ClickSwapEuro ; end;

procedure TFSaisieEff.BComplementClick(Sender: TObject);
begin ClickComplement ; end;

procedure TFSaisieEff.BModifRIBClick(Sender: TObject);
begin ClickModifRIB ; end;

procedure TFSaisieEff.BChercherClick(Sender: TObject);
begin ClickCherche ; end;

procedure TFSaisieEff.BRefClick(Sender: TObject);
begin
CalculRefAuto(Sender) ; If GS.Enabled Then GS.SetFocus ;
end;

procedure TFSaisieEff.ValideCeQuePeux( Lig : longint ) ;
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

{==========================================================================================}
{============================ Objet MOUVEMENT, Divers =====================================}
{==========================================================================================}
procedure TFSaisieEff.NumeroteVentils ;
Var OBM	 : TOBM;
		OBA	 : TOB ;
    NumL,NumAxe,NumV : integer ;
BEGIN
	for NumL:=1 to GS.RowCount-2 do BEGIN
    OBM := GetO(GS,NumL) ;
    if OBM=Nil then Continue ;
    for NumAxe:=0 to OBM.Detail.Count-1 do
    	if OBM.Detail[NumAxe].Detail.Count>0 then BEGIN
        for NumV:=0 to OBM.Detail[NumAxe].Detail.Count-1 do begin
          OBA := OBM.Detail[NumAxe].Detail[NumV];
        	OBA.PutValue('Y_NUMLIGNE',NumL) ;
        end;
      end;
  end;
end;

procedure TFSaisieEff.GotoEntete ;
Var i : integer ;
    T : TWinControl ;
BEGIN
if PasModif then Exit ; if Not PEntete.Enabled then Exit ;
for i:=0 to PEntete.ControlCount-1 do if PEntete.Controls[i] is TWinControl then
    BEGIN
    T:=TWinControl(PEntete.Controls[i]) ;
    if ((T.CanFocus) and (Copy(T.Name,1,2)='E_') and (T.ClassType<>THLabel)) then BEGIN T.SetFocus ; Break ; END ;
    END ;
END ;

procedure TFSaisieEff.GereComplements(Lig : integer) ;
Var O      : TOBM ;
    CGen   : TGGeneral ;
    ModBN  : boolean ;
    StComp,StLibre : String ;
    RC             : R_COMP ;
    NumRad         : integer ;
BEGIN
if PasModif then Exit ;
O:=GetO(GS,Lig) ; if O=Nil then Exit ; if O.CompS then Exit ;
CGen:=GetGGeneral(GS,Lig) ; if CGen=Nil then Exit ;
NumRad:=ComporteLigne(Scenario,CGen.General,StComp,StLibre) ; if NumRad<=0 then Exit ;
if ((Pos('X',StComp)<=0) and (Pos('X',StLibre)<=0)) then Exit ;
RC.StComporte:=StComp ; RC.StLibre:=StLibre ; RC.Conso:=False ; RC.DateC:=DatePiece ; 
RC.Attributs:=True ; RC.MemoComp:=MemoComp ; RC.Origine:=NumRad ;
O.CompS:=True ; if not SaisieComplement(O,EcrGen,Action,ModBN,RC) then Exit ;
if GS.CanFocus then GS.SetFocus ;
END ;


procedure TFSaisieEff.DecrementeNumero ;
Var Facturier : String3 ;
    DD : TDateTime ;
BEGIN
Facturier:=SAJAL.CompteurNormal ;
DD:=StrToDate(E_DATECOMPTABLE.Text) ;
SetDecNum(EcrGen,Facturier,DD) ;
END ;

Function TFSaisieEff.FermerSaisie : boolean  ;
Var i : integer ;
    Okok : boolean ;
BEGIN
  {JP 05/07/07 : FQ 19022 : Si on répond non à la confirmation fermeture}
  FNeFermePlus := False;

  Result:=True ;

if ((Action<>taConsult) and (ModeForce<>tmNormal)) then BEGIN Result:=False ; Exit ; END ;
if ((NeedEdition) and (Action=taCreat) and (TPIECE.Count>0)) then EditionSaisie ;
if ((Action=taCreat) and (M.TypeGuide='NOR') and (M.FromGuide)) then Exit ;
if Action=taModif then
   BEGIN
     // Confirmez-vous l'abandon de la saisie ?
     if HPiece.Execute(3,caption,'')<>mrYes then begin
       {JP 05/07/07 : FQ 19022 : Si on répond non à la confirmation fermeture}
       FNeFermePlus := True;

       Result:=False ;
     end;
   END else
   BEGIN
   if PossibleRecupNum  then
      BEGIN
      Okok:=True ;
      // Confirmez-vous l'abandon de la saisie ?
      i:=HPiece.Execute(3,caption,'') ;
      {JP 05/07/07 : FQ 19022 : Si on répond non à la confirmation fermeture}
      FNeFermePlus := i <> mrYes;

      if i<>mrYes then BEGIN Result:=False ; Exit ; END ;
      END else Okok:=False ;
   if ((Okok) and (PossibleRecupNum)) then {Revérifier, le numéro a pu être attribué entre temps!}
      BEGIN
      if (H_NUMEROPIECE_.Visible=False) then
         BEGIN
         if Transactions(DecrementeNumero,10)<>oeOk then BEGIN
           Result:=FALSE ;
           // ATTENTION : Problème de numérotation. Vérifiez votre compteur journal.
           MessageAlerte(HDiv.Mess[6]) ;
         END ;
         END ;
      END else
      BEGIN
      // L'abandon de la saisie est impossible. Vous devez valider la pièce.
      HPiece.Execute(4,caption,'') ; Result:=False ;
      END ;
   END ;
END ;

procedure TFSaisieEff.AlimObjetMvt ( Lig : integer ; AvecM : boolean ; Treso : Byte) ;
Var OBM   : TOBM ;
    XD,XC : Double ;
    CGen : TGGeneral ;
    St   : String ;
    LaDate : TDateTime ;
BEGIN
if PasModif then exit ; If SAJAL=NIL Then Exit ;
Case Treso Of
  0 : BEGIN
      OBM:=GetO(GS,Lig) ;
      if OBM=Nil then BEGIN AlloueMvt(GS,EcrGen,Lig,True) ; OBM:=GetO(GS,Lig) ; END ;
      OBM.PutMvt('E_GENERAL',GS.Cells[SA_Gen,Lig]) ;
      OBM.PutMvt('E_AUXILIAIRE',GS.Cells[SA_Aux,Lig]) ;
      OBM.PutMvt('E_EXERCICE',QuelExoDT(StrToDate((E_Datecomptable.Text)))) ;
      (*  GG ap priori ne sert à rien ( et source de bug en plus) :
      if EstLettrable(Lig)<>NonEche Then
         BEGIN
         OBM.PutMvt('E_DATEECHEANCE',IDate1900) ;
         OBM.PutMvt('E_MODEPAIE','') ;
         END else
         BEGIN
         OBM.PutMvt('E_DATEECHEANCE',IDate1900) ;
         OBM.PutMvt('E_MODEPAIE','') ;
         END ;
      *)
      END ;
  1,2 : BEGIN
        If OBMT[Treso]=NIL Then OBMT[Treso]:=TOBM.Create(EcrGen,'',TRUE) ;
        OBM:=OBMT[Treso] ;
        OBM.PutMvt('E_GENERAL',SAJAL.TRESO.CPT) ;
        OBM.PutMvt('E_AUXILIAIRE','') ;
        OBM.PutMvt('E_LIBELLE',E_LIBTRESO.TEXT) ;
        OBM.PutMvt('E_REFINTERNE',E_REFTRESO.TEXT) ;
        OBM.PutMvt('E_EXERCICE',QuelExoDT(StrToDate(E_Datecomptable.Text))) ;
        END Else Exit ;
  END ;
If Lig<=0 Then Exit ;
OBM:=GetO(GS,Lig) ;
if OBM=Nil then BEGIN AlloueMvt(GS,EcrGen,Lig,True) ; OBM:=GetO(GS,Lig) ; END ;
LaDate:=StrToDate(E_DateComptable.Text) ;
OBM.PutMvt('E_EXERCICE',QuelExoDT(LaDate)) ;
OBM.PutMvt('E_VISION',SAJAL.TRESO.STypCtr) ;
OBM.PutMvt('E_JOURNAL',SAJAL.JOURNAL)           ; OBM.PutMvt('E_DATECOMPTABLE',LaDate) ;
OBM.PutMvt('E_CODEACCEPT','ACC') ;
{$IFNDEF SPEC302}
OBM.PutMvt('E_PERIODE',GetPeriode(LaDate))      ; OBM.PutMvt('E_SEMAINE',NumSemaine(LaDate)) ;
{$ENDIF}
OBM.PutMvt('E_NUMEROPIECE',NumPieceInt)         ; OBM.PutMvt('E_NUMLIGNE',Lig) ;
OBM.PutMvt('E_GENERAL',GS.Cells[SA_Gen,Lig])    ; OBM.PutMvt('E_AUXILIAIRE',GS.Cells[SA_Aux,Lig]) ;
OBM.PutMvt('E_QUALIFPIECE',M.Simul)             ; OBM.PutMvt('E_DEVISE',DEV.Code) ;
OBM.PutMvt('E_TAUXDEV',DEV.Taux)                ; OBM.PutMvt('E_DATETAUXDEV',DEV.DateTaux) ;
OBM.PutMvt('E_NATUREPIECE',E_NATUREPIECE.Value) ; OBM.PutMvt('E_ETABLISSEMENT',E_ETABLISSEMENT.Value) ;
OBM.PutMvt('E_LIBELLE',Copy(GS.Cells[SA_Lib,Lig],1,35)) ; 
OBM.PutMvt('E_REFINTERNE',Copy(GS.Cells[SA_RefI,Lig],1,35)) ;
If SA_RIB>-1 Then OBM.PutMvt('E_RIB',Copy(GS.Cells[SA_Rib,Lig],1,35)) ;
If SA_REFL>-1 Then OBM.PutMvt('E_REFLIBRE',Copy(GS.Cells[SA_RefL,Lig],1,35)) ;
If SA_REFE>-1 Then OBM.PutMvt('E_REFEXTERNE',Copy(GS.Cells[SA_RefE,Lig],1,35)) ;
OBM.MajLesDates(Action) ;
if AvecM then
   BEGIN
   XD:=ValD(GS,Lig) ; XC:=ValC(GS,Lig) ;
   OBM.SetMontants(XD,XC,DEV,True) ;
   END ;
CGen:=GetGGeneral(GS,Lig) ;
if CGen<>Nil then
   BEGIN
   if Ventilable(CGen,0) Then OBM.PutMvt('E_ANA','X') Else OBM.PutMvt('E_ANA','-') ;
   St:=OBM.GetMvt('E_QUALIFQTE1') ; if ((St='') or (St='...')) then OBM.PutMvt('E_QUALIFQTE1',CGen.QQ1) ;
   St:=OBM.GetMvt('E_QUALIFQTE2') ; if ((St='') or (St='...')) then OBM.PutMvt('E_QUALIFQTE2',CGen.QQ2) ;
   END ;
END ;

Procedure TFSaisieEff.BloqueGrille(Bloc : boolean ) ;
BEGIN
  If Action<>taConsult Then BEGIN
    If Bloc Then BEGIN
      GridEna(GS,False);
      MajEnCours := True;
      GS.controlstyle := GS.controlstyle-[csClickEvents,csCaptureMouse];
      END
    Else BEGIN
      GS.controlstyle := GS.controlstyle+[csClickEvents,csCaptureMouse];
      MajEnCours := False;
      GridEna(GS,True);
    end;
  end;
END ;

{==========================================================================================}
{================================ Methodes du HGrid =======================================}
{==========================================================================================}
procedure TFSaisieEff.GSDblClick(Sender: TObject);
begin
if ((Action<>taConsult) and (ModeForce=tmDevise)) then ForcerMode(False,VK_RETURN) else ClickZoom ;
end;

procedure TFSaisieEff.GSCellExit(Sender: TObject; var ACol,ARow : Longint ; var Cancel : Boolean );
begin
if PasModif then exit ;
if GS.Row=ARow then
   BEGIN
   if ACol=SA_Gen then
      BEGIN
      (*
      If (ModeSaisie<>OnBqe) And (GS.ROW=ARow) And (GS.Cells[ACol,ARow]='') And (ARow>1) Then
         BEGIN
         If SAJAL.Treso.TypCtr=Ligne Then
            BEGIN
            If ARow>2 Then GS.Cells[ACol,ARow]:=GS.Cells[ACol,ARow-2] ;
            END Else GS.Cells[ACol,ARow]:=GS.Cells[ACol,ARow-1] ;
         END ;
      *)
      if ((GS.Cells[ACol,ARow]='') and ((GS.Cells[SA_Aux,ARow]='') or (GS.Objects[SA_Aux,ARow]=Nil))) then Exit ;
      if LeMeme(GS,ACol,ARow) then Exit ;
      if ChercheGen(ACol,ARow,False)<=0 then BEGIN Cancel:=True ; Exit ; END ;
      if Not Cancel then CpteAuto:=False ;
      END else if ACol=SA_Aux then
      BEGIN
      if ((GS.Cells[ACol,ARow]='') and (GS.Col=SA_Gen)) then Exit ;
      if LeMeme(GS,ACol,ARow) then Exit ;
      if ChercheAux(Acol,ARow,False)<=0 then BEGIN Cancel:=True ; Exit ; END ;
      END else
      if ACol=SA_REFI Then
         BEGIN
         if (GS.Cells[ACol,ARow]='') Then
            BEGIN
            if GS.ROW=ARow Then
               BEGIN
               if SI_NUMREF<0 Then
                  BEGIN
                  if ARow>1 then GS.Cells[ACol,ARow]:=GS.Cells[ACol,ARow-1] ;
                  END else GS.Cells[ACol,ARow]:=E_NUMREF.Text ;
               END ;
            END ;
         END else if ACol=SA_LIB Then
         BEGIN
         if GS.ROW=ARow Then
            BEGIN
            if GS.Cells[ACol,ARow]='' Then
               BEGIN
               if ARow>1 then GS.Cells[ACol,ARow]:=GS.Cells[ACol,ARow-1] ;
               END ;
            END ;
         END Else
      If ACol=SA_RIB Then If (ARow=GS.RowCount-2) And (GS.Col>ACol) Then
         BEGIN
         If Not VerifEtUpdateRib(ARow) Then GS.Col:=Acol ; 
         END ;
   END ;
if ((ACol=SA_Debit) or (ACol=SA_Credit)) then
  begin
  // DEV 3216 10/04/2006 SBO
  if (pos('+',GS.Cells[ACol, ARow])>0) or (pos('-',GS.Cells[ACol, ARow])>0) or (pos('/',GS.Cells[ACol, ARow])>0) or (pos('*',GS.Cells[ACol, ARow])>0) then
    begin
    // Opération sur ligne précédente ?
    //if (not VH^.MontantNegatif) and ( GS.Cells[ACol, ARow][1] in ['+','-','/','*'] ) then
    if (not VH^.MontantNegatif) and ( pos( Copy(GS.Cells[ACol, ARow], 1, 1), '+;-;/;*' ) > 0 ) then
      begin
      if ACol = SA_DEBIT
        then GS.Cells[ACol, ARow]:=GFormule('{[E_DEBIT]'+GS.Cells[ACol, ARow]+'}',  GetFormulePourCalc, nil, 1)
        else GS.Cells[ACol, ARow]:=GFormule('{[E_CREDIT]'+GS.Cells[ACol, ARow]+'}', GetFormulePourCalc, nil, 1);
      end
    else  // sinon on effectue l'opération simplement
      begin
      if Pos(GS.Cells[ACol, ARow][1],'/')>-1
          then GS.Cells[ACol, ARow]:='{' + GetFormatPourCalc + GS.Cells[ACol, ARow]+'}'; // FQ18371 et FQ 18080
      GS.Cells[ACol, ARow] := GFormule( '{' + GS.Cells[ACol, ARow] + '}', GetFormulePourCalc, nil, 1) ;
      end ;
   end;
   FormatMontant(GS, ACol, ARow, DEV.Decimale) ;
  // FIN DEV 3216 10/04/2006 SBO
  ChercheMontant(Acol,ARow) ;
  end ;
end;

Function TFSaisieEff.IsAccessible(Ou : LongInt) : Boolean ;
Var OBM : TOBM ;
    EtatLig : String3 ;
BEGIN
  OBM := GetO(Gs,Ou);
  Result := True;
  if OBM<>nil then begin
    EtatLig:=OBM.GetMvt('E_ETATLETTRAGE') ;
    If (EtatLig='TL') or (EtatLig='PL') then Result := False;
  end;
end;

procedure TFSaisieEff.BasculeGS(AccesConsult : Boolean) ;
BEGIN
  if ((ModeForce=tmNormal)) then BEGIN
    if AccesConsult then BEGIN
      GS.CacheEdit ;
      GS.Options:=GS.Options-[GoEditing,GoTabs,GoAlwaysShowEditor] ;
      GS.Options:=GS.Options+[GoRowSelect] ;
      G_LIBELLE.Font.Color:=clRed ;
      // Mouvement lettré non modifiable.
      G_LIBELLE.Caption:=HDiv.Mess[7] ;
      END
    else BEGIN
      GS.Options:=GS.Options-[GoRowSelect] ;
      GS.Options:=GS.Options+[GoEditing,GoTabs,GoAlwaysShowEditor] ;
      GS.MontreEdit ;
    END ;
  END ;
END ;

function TFSaisieEff.TrouveSensTreso ( GS : THGrid ; Nat : String3 ; Lig : integer ) : byte ;
Var CGen : TGGeneral ;
BEGIN
  If ModeSaisie=OnBqe Then Result:=TrouveSens(GS,Nat,Lig)
  Else BEGIN
    Result:=3 ;
    if GS.Cells[SA_Gen,Lig]='' then exit ;
    CGen:=GetGGeneral(GS,Lig) ;
    if CGen=Nil then Exit ;
    If (Cgen.natureGene='COC') Or (Cgen.natureGene='TID') Then Result:=2 ;
  END ;
END ;


procedure TFSaisieEff.GSCellEnter(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
Var CGen : TGGeneral ;
    OnSort : Boolean ;
    AccesConsult : Boolean ;
begin
If ARow<>GS.Row Then
   BEGIN
   AccesConsult:=Not IsAccessible(GS.Row) ;
   BasculeGS(AccesConsult) ;
   END ;
OnSort:=(GS.Col<>SA_Gen) and (GS.Col<>SA_Aux) And (GS.Col<>SA_DateC) ;
if ((Action<>taConsult) and (Not EstRempli(GS,GS.Row)) and OnSort)
   then BEGIN ACol:=SA_Gen ; ARow:=GS.Row ; Cancel:=True ; Exit ; END ;
if ((H_NUMEROPIECE_.Visible) and (GS.Row>1)) then
   BEGIN
   // ATTENTION : Le numéro de pièce n'est pas définitivement attribué. Abandonnez la saisie et ré-essayez de nouveau.
   if Transactions(AttribNumeroDef,10)<>oeOK then MessageAlerte(HDiv.Mess[5]) ;
   BlocageEntete(Self,TRUE) ;
   E_CONTREPARTIEGEN.Enabled:=FALSE ;
   If ModeSaisie<>OnBqe Then E_CONTREPARTIEGEN.Enabled:=FALSE ;
   END ;
BValide.Enabled:=((H_NumeroPiece.Visible) and (Equilibre)) ;
GereNewLigneT ;
If ((H_NUMEROPIECE_.Visible) and (GS.Row=1)) Then //TR
   BEGIN
   if SI_NUMREF>=0 Then
      BEGIN
      SI_NUMREF:=StrToInt(Trim(E_NUMDEP.Text)) ;
      E_NUMREF.Value:=SI_NUMREF ;
      END ;
   END ;
// TR
If (SAJAL.TRESO.TypCtr=Ligne) Then
   BEGIN
   CGen:=GetGGeneral(GS,GS.Row) ;
   if (GS.Cells[SA_Gen,GS.Row]=SAJAL.TRESO.Cpt) And (CGen<>NIL) And (GS.Row Mod 2=0) Then
      BEGIN
      if ARow<GS.Row then BEGIN ACol:=GS.Col{SA_Gen} ; ARow:=GS.Row+1 ; END else BEGIN ACol:=GS.col{SA_Credit} ; ARow:=GS.Row-1 ; END ;
      Cancel:=TRUE ; Exit ;
      END ;
   END ;

BZoom.Enabled := ((GS.Col=SA_Gen) and (GS.Cells[SA_Gen,GS.Row]<>'')) or ((GS.Col=SA_Aux) and (GS.Cells[SA_Aux,GS.Row]<>''));

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
         if ACol<=SA_Debit-1 then ACol:=SA_Credit else if ACol>=SA_Credit then ACol:=SA_Debit-1 ;
         END else
         BEGIN
         ACol:=SA_Credit ; ARow:=GS.Row ;
          END ;
      Cancel:=TRUE ;
      END else
      BEGIN
      if ((GS.Row=ARow) and (TrouveSensTreso(GS,E_NATUREPIECE.Value,GS.Row)=2) and (Acol<GS.Col) and (GS.Cells[SA_Debit,GS.Row]=''))
         then BEGIN ACol:=SA_Credit ; Cancel:=True ; END ;
      END ;
   END else if (GS.Col=SA_DateC) then
   BEGIN
   If (ACol=SA_GEN) And (ARow=GS.ROW) And (ARow>1) Then
      BEGIN
      ACol:=SA_CREDIT ; ARow:=ARow-1 ;
      END Else
      BEGIN
      ACol:=SA_GEN ; ARow:=GS.Row ;
      END ;
   {
   if GS.Row=ARow then
      BEGIN
      ACol:=SA_GEN ; ARow:=GS.Row ;
      END else
      BEGIN
      if ACol=SA_GEN then BEGIN ARow:=ARow+1 ; ACol:=SA_Gen ; END else BEGIN ACol:=SA_Debit ; ARow:=GS.Row ; END ;
      END ;
   }
   Cancel:=True ;
   END ;
end;

procedure TFSaisieEff.GSRowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
Var AccesConsult : Boolean ;
begin
GridEna(GS,FALSE) ;
AccesConsult:=Not IsAccessible(Ou) ;
BasculeGS(AccesConsult) ;
if ((Not EstRempli(GS,Ou)) and (Not PasModif)) then DefautLigne(Ou,True) ;
CurLig:=Ou ; AffichePied ; GereEnabled(Ou) ;
GridEna(GS,TRUE) ;
end;

{Procedure TFSaisieEff.MajOBMViaEche(Lig : integer ; OBM : TOBM ; R : T_ModeRegl ; NumEche : integer ) ;
// Pour contrer le bug suivant :
//   Modif sur une ligne d'un aux. Désalloueeche remet etatlettrage à"RI" sur l'OBM.
//   Ensuite,seul l'objet RMOD connait l'état lettrage et la maj directe table se base
//   sur l'état du champ e_etatlettrage de l'OBM. D'ou le besoin de remettre à jour l'OBM
//   sur les zones non saisissables concernant le lettrage
BEGIN
If OBM<>NIL Then
   BEGIN
   OBM.PutMvt('E_ETATLETTRAGE',R.TabEche[NumEche].EtatLettrage) ; OBM.PutMvt('E_NUMECHE',NumEche) ;
   If NumEche>0 Then OBM.PutMvt('E_ECHE','X') ;
   END ;
END ;}

procedure TFSaisieEff.AfficheExigeTva ;
Var Nat : String3 ;
BEGIN
  if Not VH^.OuiTvaEnc then BEGIN BPopTva.Visible:=False; Exit; END ;
  Nat:=E_NATUREPIECE.Value ;
  if ((Nat<>'FC') and (Nat<>'AC') and (Nat<>'OC') and (Nat<>'FF') and (Nat<>'AF') and (Nat<>'OF')) then BEGIN BPopTva.Visible:=False ; Exit ; END ;
  if (((Nat='OC') or (Nat='FC') or (Nat='AC')) and (VH^.TvaEncSociete='TD')) then BEGIN BPopTva.Visible:=False ; Exit ; END ;
  if (((Nat='FC') or (Nat='AC')) and (SorteTva<>stvVente)) then BEGIN BPopTva.Visible:=False ; Exit ; END ;
  if (((Nat='FF') or (Nat='AF')) and (SorteTva<>stvAchat)) then BEGIN BPopTva.Visible:=False ; Exit ; END ;
  BPopTva.Visible:=True ;
END ;

{Procedure TFSaisieEff.AffecteTva ( C,L : integer ) ;
Var O : TOBM ;
    CAux  : TGTiers ;
    CGen  : TGGeneral ;
    StE   : String ;
BEGIN
// #TVAENC
if Action=taConsult then Exit ;
O:=GetO(GS,L) ; if O=Nil then Exit ;
if isTiers(GS,L) then
   BEGIN
   RenseigneRegime(L) ;
   if GS.Cells[SA_Aux,L]<>'' then
      BEGIN
      CAux:=GetGTiers(GS,L) ; if CAux=Nil then Exit ;
      if Not ChoixExige then
         BEGIN
         if ((SorteTva=stvAchat) or (E_NATUREPIECE.Value='OF')) then ExigeTva:=Code2Exige(CAux.Tva_Encaissement) else
          if ((SorteTva=stvDivers) and (E_NATUREPIECE.Value='OC') and (VH^.TvaEncSociete='TE')) then ExigeTva:=tvaEncais
                                                                else ExigeTva:=tvaMixte ;
         ExigeEntete:=ExigeTva ;
         BPopTva.ItemIndex:=Ord(ExigeEntete)+1 ;
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
END ;}

{procedure TFSaisieEff.RenseigneRegime ( Lig : integer ) ;
Var CGen : TGGeneral ;
    CAux : TGTiers ;
BEGIN
if Lig<=0 then Exit ;
if GS.Cells[SA_Aux,Lig]<>'' then
   BEGIN
   CAux:=GetGTiers(GS,Lig) ;
   if CAux<>Nil then
      BEGIN
      if ((GeneRegTVA='') or (Not RegimeForce)) then GeneRegTVA:=CAux.RegimeTVA ;
      GeneSoumisTPF:=CAux.SoumisTPF ;
      END ;
   END else
   BEGIN
   CGen:=GetGGeneral(GS,Lig) ;
   if CGen<>Nil then
      BEGIN
      if ((GeneRegTVA='') or (Not RegimeForce)) then GeneRegTVA:=CGen.RegimeTVA ;
      GeneSoumisTPF:=CGen.SoumisTPF ;
      END ;
   END ;
END ;}

procedure TFSaisieEff.GetSorteTva ;
BEGIN
  SorteTva:=stvDivers ;
  if SAJAL=Nil then Exit ;
  if ((SAJAL.NatureJal='VTE') and ((E_NATUREPIECE.Value='FC') or (E_NATUREPIECE.Value='AC'))) then
    SorteTva:=stvVente
  else if ((SAJAL.NatureJal='ACH') and ((E_NATUREPIECE.Value='FF') or (E_NATUREPIECE.Value='AF'))) then
    SorteTva:=stvAchat ;
END ;

Function TFSaisieEff.TvaTauxDirecteur ( Lig : integer ; Client,ToutDebit : boolean ; Regime : String3 ) : Boolean ;
Var OMODR   : TMOD ;
    TTC,HT  : Double ;
    CodeTva : String3 ;
    Taux,XX : Double ;
    i       : integer ;
    O       : TOBM ;
BEGIN
{#TVAENC}
  Result:=FALSE ;
  if Not VH^.OuiTvaEnc then Exit ;
  OMODR:=TMOD(GS.Objects[SA_NumP,Lig]) ;
  if OMODR=Nil then Exit ;

  {Si modifié manuellement ou recalculé prorata --> pas de traitement}
  if ((OMODR.ModR.ModifTva) or (OMODR.ModR.TotalAPayerP=0)) then Exit ;
  O:=GetO(GS,Lig) ;
  if O=Nil then Exit ;

 { LOBM:=TList(GS.Objects[SA_NumL,Lig-1]) ;
  TTC:=TOBM(LOBM[0]).GetMvt('E_ECHEENC1'); ymoo}

  TTC:=O.GetMvt('E_CREDIT')-O.GetMvt('E_DEBIT') ;
  if Not Client then TTC:=-TTC ;
  CodeTva:=VH^.NumCodeBase[1] ;
  if CodeTva='' then Exit ;
  Taux:=Tva2Taux(Regime,CodeTva,Not Client) ;
  if Taux=-1 then Exit ;

  {Prorater le HT, calculé sur le taux directeur, sur les échéances}
  HT:=Arrondi(TTC/(1.0+Taux),V_PGI.OkDecV) ;
  for i:=1 to MaxEche do BEGIN
    FillChar(OMODR.ModR.TabEche[i].TAV,Sizeof(OMODR.ModR.TabEche[i].TAV),#0) ;
    XX:=Arrondi(HT*OMODR.ModR.TabEche[i].MontantP/OMODR.ModR.TotalAPayerP,V_PGI.OkDecV) ;
    if ToutDebit then BEGIN
      If XX<>OMODR.ModR.TabEche[i].TAV[5] Then Result:=TRUE ;
      OMODR.ModR.TabEche[i].TAV[5]:=XX ;
      END
    else BEGIN
      If XX<>OMODR.ModR.TabEche[i].TAV[1] Then Result:=TRUE ;
      OMODR.ModR.TabEche[i].TAV[1]:=XX ;
    END ;
  END ;

  {Eviter la ré-initialisation ultérieure, recalcul par prorata ou manuel}
  OMODR.ModR.ModifTva:=True ;
END ;

Function TFSaisieEff.GereTvaEncNonFact ( Lig : integer ) : Boolean ;
Var CGen : TGGeneral ;
    CAux : TGTiers ;
    Client : boolean ;
    RegEnc,StEnc,Regime : String3 ;
    Exi          : TExigeTva ;
BEGIN
{#TVAENC}
  Result:=FALSE ;
  if PasModif then Exit ;
  if Not VH^.OuiTvaEnc then Exit ;

  {Si facture --> géré ailleurs (Calculdébitcrédit)}
  if SorteTva<>stvDivers then Exit ;
  if Not isTiers(GS,Lig) then Exit ;
  Regime:='' ;
  if GS.Cells[SA_Aux,Lig]<>'' then BEGIN
    CAux:=GetGTiers(GS,Lig) ;
    if CAux=Nil then Exit ;
    if ((CAux.NatureAux='CLI') or (CAux.NatureAux='AUD')) then Client:=True else
      if ((CAux.NatureAux='FOU') or (CAux.NatureAux='AUC')) then Client:=False else Exit ;
     StEnc:=CAux.Tva_Encaissement ;
     Regime:=CAux.RegimeTva ;
     END
  else if GS.Cells[SA_Gen,Lig]<>'' then BEGIN
    CGen:=GetGGeneral(GS,Lig) ;
    if CGen=Nil then Exit ;
    if CGen.NatureGene='TID' then Client:=True else
      if CGen.NatureGene='TIC' then Client:=False else Exit ;
    StEnc:=CGen.Tva_Encaissement ;
    Regime:=CGen.RegimeTva ;
    END
  else Exit ;
  CGen:=GetGGeneral(GS,Lig) ;
  if CGen=Nil then Exit ;
  if Not EstCollFact(CGen.General) then Exit ;

  //YMOO Recopie des TVA encaiss sur les lignes de contrepartie
  RecopieTVAEnc(Lig+1,Client,False,Regime) ;

  {Si client --> param soc, sinon exige du fourn. Si choix explicite, forcer. Débit --> Aucun traitement}
  if Client then RegEnc:=VH^.TvaEncSociete
            else RegEnc:=StEnc ;
  if ChoixExige then Exi:=TExigeTva(BPopTva.ItemIndex-1)
                else Exi:=Code2Exige(RegEnc) ;

  Case Exi of
    tvaMixte  : Exit ;
    tvaEncais : Result:=TvaTauxDirecteur(Lig,Client,False,Regime) ;
    tvaDebit  : Result:=TvaTauxDirecteur(Lig,Client,True,Regime) ;
  END ;

END ;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 01/10/2007
Modifié le ... :   /  /
Description .. : Recopie des infos TVA depuis les lignes d'origine
Mots clefs ... : TVA ENCAISSEMENT
*****************************************************************}
Function TFSaisieEff.RecopieTVAEnc ( Lig : integer ; Client,ToutDebit : boolean ; Regime : String3 ) : Boolean ;
Var OMODR   : TMOD ;
    i, j, k : integer ;
    O       : TOBM ;
    LOBM    : TList ;
    EcheD, ValMax : Double ;
    Eche: array[1..4] of Double;
    TxTva, RegTva : String ;
BEGIN
  //Recopie sur les lignes de contrepartie
  Result:=FALSE ;
  if Not VH^.OuiTvaEnc then Exit ;
  OMODR:=TMOD(GS.Objects[SA_NumP,Lig]) ;
  if (OMODR=Nil) or (OMODR.ModR.ModifTva) then Exit ;

  O:=GetO(GS,Lig) ;
  if O=Nil then Exit ;

  VideListe(Lobm) ;
  Lobm.Free ;
  LOBM:=TList(GS.Objects[SA_NumL,Lig-1]) ;

  If not Assigned(Lobm) then Exit; {29/11/2007 YMO Saisie directe de montant}

  for i:=1 to MaxEche do               //échéances
  BEGIN
    FillChar(OMODR.ModR.TabEche[i].TAV,Sizeof(OMODR.ModR.TabEche[i].TAV),#0) ;

    TxTva:='';
    RegTva:='';
    For k:=1 to 4 do Eche[k]:=0.0;
    EcheD:=0;
    ValMax:=0;
    For j:=0 to LOBM.Count-1 do   //somme des lignes d'origine
    begin

    /////////////////////////////
        For k:=1 to 4 do    {ValMax est valable pour toutes les lignes}
        begin
          Eche[k] := Eche[k]+TOBM(LOBM[j]).GetMvt('E_ECHEENC'+inttostr(k));
          If TOBM(LOBM[j]).GetMvt('E_ECHEENC'+inttostr(k))>ValMax then
          begin
            TxTva:=TOBM(LOBM[j]).GetMvt('E_TVA');
            RegTva:=TOBM(LOBM[j]).GetMvt('E_REGIMETVA');
            ValMax:=TOBM(LOBM[j]).GetMvt('E_ECHEENC'+inttostr(k));
          end;
        end;
        EcheD := EcheD+TOBM(LOBM[j]).GetMvt('E_ECHEDEBIT') ;

        {Transfert des infos de TVA}
        for k:=1 to 4 do
          OMODR.ModR.TabEche[i].TAV[k]:=Eche[k] ;

        OMODR.ModR.TabEche[i].TAV[5]:=EcheD ;
        OMODR.ModR.TabEche[i].CodeTva:=txTva;

    end;
    /////////////////////////////



  END ;

  OMODR.ModR.ModifTva:=True ;
end;

procedure TFSaisieEff.GSRowExit(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
Var OkL  : Boolean ;
    OBM : TOBM ;
    i : Integer ;
begin
  if PasModif then Exit ;
  if Ou>=GS.RowCount-1 then Exit ;
  BloqueGrille(TRUE) ;
  OkL:=LigneCorrecte(Ou,False,True,TRUE) ;
  if OkL  then BEGIN
    if ((H_NUMEROPIECE_.Visible) and (ou=1)) then BEGIN
      // ATTENTION : Le numéro de pièce n'est pas définitivement attribué. Abandonnez la saisie et ré-essayez de nouveau.
      if Transactions(AttribNumeroDef,10)<>oeOK then MessageAlerte(HDiv.Mess[5]) ;
      BlocageEntete(Self,TRUE) ;
//      E_CONTREPARTIEGEN.Enabled:=FALSE ;
      If ModeSaisie<>OnBqe Then E_CONTREPARTIEGEN.Enabled:=FALSE ;
      OBM := GetO(GS,Ou) ;
      If OBM<>NIL Then OBM.PutMvt('E_NUMEROPIECE',NumPieceInt) ;
    END ;
    GereComplements(Ou) ;
    GereEche(Ou,True,False) ;
    GereAnal(Ou,True,True,FALSE,FALSE) ;
    // Dev 3946 : control cohérence qté analytique soumis à indicateur
    OBM := GetO(GS,Ou) ;
    if ( OBM <> nil ) and ( OBM.GetValue('E_ANA')='X' )
                      and TestQteAna
                      and ( ( OBM.GetNumChamp('CHECKQTE') < 0 )  or ( OBM.GetString('CHECKQTE')<>'X'  ) )
        then CheckQuantite( OBM ) ;
    // Fin Dev 3946

    For i:=1 To MaxAxe Do
      LAnaEff[i].Clear ;
    GereTvaEncNonFact(Ou) ;
    If SAJAL.TRESO.TypCtr=Ligne Then BEGIN
      GereLigneTreso(Ou) ;
    END ;
    BloqueGrille(FALSE) ;
    If SAJAL.TRESO.TypCtr In [PiedSolde,PiedDC] Then BEGIN
      If E_REFTRESO.TEXT='' Then BEGIN E_REFTRESO.TEXT:=GS.Cells[SA_REFI,Ou]; MajRefLibOBMT(0) ; END ;
      If E_LIBTRESO.TEXT='' Then BEGIN E_LIBTRESO.TEXT:=GS.Cells[SA_LIB,Ou];  MajRefLibOBMT(1) ; END ;
    END ;
    END
  else BEGIN
    Cancel:=True ;
    BloqueGrille(FALSE) ;
  END ;
end;

procedure TFSaisieEff.GSEnter(Sender: TObject);
Var StComp : String ;
    b,OkD  : boolean ;
    RC     : R_COMP ;
    StLibreEnt : String ;
begin
  {JP 15/01/08 : FQ 19363 : Pour afficher les scrollbars, bien que je pense que ce ne soit pas la source du problème}
  ShowScrollBar(GS.Handle , SB_BOTH, True);
  
  If SAJAL=NIL Then Exit ;
  If SAJAL.Journal<>E_JOURNAL.Value Then BEGIN
    // ATTENTION : Problème fichier. Veuillez revenir au menu, puis relancer la fonction.
    MessageAlerte(HDiv.Mess[13]) ; Exit ;
  END ;
  GereEnabled(GS.Row) ;
  if Action<>taCreat then Exit ;
  if PasModif then Exit ;
 if Not OkScenario then Exit ;
  if GS.RowCount>2 then Exit ;
  if DejaRentre then Exit ;
  if Entete=Nil then Exit ;
  if EstRempli(GS,1) then Exit ;

{$IFNDEF SPEC302}
  OkD:=False ;
  if VH^.CPDateObli then OkD:=True
  else BEGIN
    if ((OkScenario) and (Scenario<>Nil)) then
      if Scenario.GetMvt('SC_DATEOBLIGEE')='X' then OkD:=True ;
  END ;
  if ((Action=taCreat) and (Not EstRempli(GS,1)) and (OkD) and (Not DejaRentre) and (Not RentreDate)) then BEGIN
    // JLD 4394 if HPiece.Execute(41,Caption,'')<>mrYes then BEGIN E_DATECOMPTABLE.SetFocus ; Exit ; END ;
    if E_DATECOMPTABLE.CanFocus then E_DATECOMPTABLE.SetFocus ; Exit ;
  END ;
{$ENDIF}
  RentreDate:=True ;
  DejaRentre:=True ;
  StComp:=Entete.GetMvt('E_ETAT') ;
  StLibreEnt:=Scenario.GetMvt('SC_LIBREENTETE') ;
  if ((Pos('X',StComp)<=0) and (Pos('X',StLibreEnt)<=0)) then Exit ;
  RC.StComporte:=StComp ;
  RC.StLibre:=StLibreEnt ;
  RC.Conso:=False ;
  RC.Attributs:=True ;
  RC.MemoComp:=MemoComp ;
  RC.Origine:=0 ;
  RC.DateC:=DatePiece ;
  SaisieComplement(Entete,EcrGen,Action,b,RC) ;
  DejaRentre:=True ;
  DefautLigne(1,True) ;
  GereEnabled(1) ;
end;

procedure TFSaisieEff.GSExit(Sender: TObject);
Var b,bb : boolean ;
    C,R  : longint ;
begin
  bb:=FALSE ; b:=FALSE ;
  if PasModif then Exit ;
  if Valide97.Tag<>1 then Exit ;
  if Outils.Tag<>1 then Exit ;
  C:=GS.Col ; R:=GS.Row ;
  GSCellExit(GS,C,R,b) ;
  GSRowExit(GS,R,b,bb) ;
end;

procedure TFSaisieEff.GSMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin GX:=X ; GY:=Y ; end;

procedure TFSaisieEff.POPSPopup(Sender: TObject);
begin InitPopUp(Self) ; end;

procedure TFSaisieEff.FindSaisFind(Sender: TObject);
begin Rechercher(GS,FindSais,FindFirst) ; end;

procedure TFSaisieEff.H_MODEPAIEDblClick(Sender: TObject);
begin
FicheModePaie_AGL(E_MODEPAIE.VALUE) ;
end;

procedure TFSaisieEff.BModifSerieClick(Sender: TObject);
Var b : boolean ;
    St,StAvant : String ;
    Lig     : integer ;
    OBM         : TOBM ;
    DD : TDateTime ;
    RC         : R_COMP ;
begin
  if Action=taConsult then Exit ;
  RC.StComporte:='--XXXX----' ;
  RC.StLibre:='' ;
  RC.Conso:=True ;
  RC.Attributs:=False ;
  RC.MemoComp:=Nil ;
  RC.Origine:=-1 ;
  RC.DateC:=DatePiece ;

  {pas de modif série zones libres en saisie tréso}
  ModifSerie.Free ;
  ModifSerie:=TOBM.Create(EcrGen,'',True) ;
  if Not SaisieComplement(ModifSerie,EcrGen,taCreat,b,RC) then Exit ;
  // Confirmez-vous la modification de toutes les lignes de la pièce ?
  if HPiece.Execute(28,caption,'')<>mrYes then Exit ;
  BeginTrans ;
  for Lig:=1 to GS.RowCount-2 do BEGIN
    OBM:=GetO(GS,Lig);
    If OBM<>NIL Then BEGIN
      StAvant:=OBM.GetMvt('E_REFEXTERNE') ; St:=ModifSerie.GetMvt('E_REFEXTERNE') ;
      If St<>'' Then OBM.PutMvt('E_REFEXTERNE',St) ;

      StAvant:=OBM.GetMvt('E_REFLIBRE') ; St:=ModifSerie.GetMvt('E_REFLIBRE') ;
      If St<>'' Then OBM.PutMvt('E_REFLIBRE',St) ;

      StAvant:=OBM.GetMvt('E_AFFAIRE') ; St:=ModifSerie.GetMvt('E_AFFAIRE') ;
      If St<>'' Then OBM.PutMvt('E_AFFAIRE',St) ;

      StAvant:=OBM.GetMvt('E_CONSO') ; St:=ModifSerie.GetMvt('E_CONSO') ;
      If St<>'' Then OBM.PutMvt('E_CONSO',St) ;

      DD:=ModifSerie.GetMvt('E_DATEREFEXTERNE') ;
      if DD>Encodedate(1901,01,01) then OBM.PutMvt('E_DATEREFEXTERNE',DD) ;
    END ;
  END ;
  CommitTrans ;
end;

procedure TFSaisieEff.BZoomTClick(Sender: TObject);
Var A : TActionFiche ;
begin
  // attention, utiliser JaiLeDroit
  A:=taModif ; If SAJAL=NIL Then Exit ;

  if Not ExJaiLeDroitConcept(TConcept(ccGenModif),False) then a:=taConsult ;
  FicheGene(Nil,'',SAJAL.TRESO.CPT,A,0) ;
end;


procedure TFSaisieEff.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFSaisieEff.GSKeyPress(Sender: TObject; var Key: Char);
begin
If MajEnCours Or (Not GS.SynEnabled) Then BEGIN Key:=#0 ; Exit ; END ;
end;

procedure TFSaisieEff.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFSaisieEff.E_JOURNALClick(Sender: TObject);
Var Jal : String ;
begin
  Jal:=E_JOURNAL.Value ;
  if Jal='' then BEGIN SAJAL.Free ; SAJAL:=Nil ; Exit; END ;
  if SAJAL<>Nil then if SAJAL.JOURNAL=Jal then Exit;
  if ((VH^.JalAutorises<>'') and (Pos(';'+Jal+';',VH^.JalAutorises)<=0)) then BEGIN
    // Vous n'avez pas le droit de saisir sur ce journal.
    HPiece.Execute(26,caption,'') ;
    if SAJAL<>Nil then BEGIN SAJAL.Free ; SAJAL:=Nil ; END ;
    E_JOURNAL.Value:='' ;
    if Action=taCreat then BEGIN E_JOURNAL.SetFocus ; GS.Enabled:=False ; END ;
    Exit;
  END;
  ChercheJal(Jal,False) ;
  if ((Action=taCreat) and (NumPieceInt=0)) then BEGIN
    // Vous n'avez pas défini de compteur de numérotation pour ce journal.
    HPiece.Execute(27,caption,'') ;
    if SAJAL<>Nil then BEGIN SAJAL.Free ; SAJAL:=Nil ; END ;
    E_JOURNAL.Value:='' ; E_JOURNAL.SetFocus ; GS.Enabled:=False ;
    Exit;
  END ;
  CalculSoldeCompteT ;
  AlimObjetMvt(0,FALSE,1) ;
  AlimObjetMvt(0,FALSE,2) ;
  InitBoutonValide ;
END ;

procedure TFSaisieEff.BMenuZoomMouseEnter(Sender: TObject);
begin
PopZoom97(BMenuZoom,POPZ) ;
Outils.Tag:=1 ;
end;

procedure TFSaisieEff.BValideMouseEnter(Sender: TObject);
begin
Valide97.Tag:=1 ;
end;

procedure TFSaisieEff.BValideMouseExit(Sender: TObject);
begin
Valide97.Tag:=0 ;
end;

procedure TFSaisieEff.BMenuZoomMouseExit(Sender: TObject);
begin
Outils.Tag:=0 ;
end;

procedure TFSaisieEff.BRAAbandonClick(Sender: TObject);
begin
  PEntete.Enabled:=True ;
  Outils.Enabled:=True ;
  Valide97.Enabled:=True ;
  GS.Enabled:=True ;
  PPied.Enabled:=True ;
  GS.SetFocus ;
  PRefAuto.Visible:=False ;
  ModeRA:=False ;
end;

procedure TFSaisieEff.BRAValideClick(Sender: TObject);
Var OBM, OBA : TOBM ;
    i,NumA,NumL : Integer ;
    OkOk : Boolean ;
    St,RefAvant : String ;
    CGen  : TGGeneral ;
begin
  SI_NumRef:=-1 ;
  E_NumDepExit(Sender) ;
  BeginTrans ;
  for i:=1 to GS.RowCount-2 do BEGIN
    OkOk:=TRUE ; St:='' ;
    if (SAJAL.Treso.TypCtr=Ligne) and (GS.Cells[SA_Gen,i]=SAJAL.TRESO.Cpt) then OkOk:=FALSE ;
    if OkOk and (i>1) then IncRef ;
    if SI_NumRef<0 then St:=''
                   else St:=E_NUMREF.Text ;
    GS.Cells[SA_RefI,i]:=St ;
    OBM:=GetO(GS,i) ;
    if OBM<>NIL then BEGIN
      RefAvant:=OBM.GetMvt('E_REFINTERNE') ;
      OBM.PutMvt('E_REFINTERNE',St) ;
    END ;

    // Analytique
    CGen:=GetGGeneral(GS,i);
    for NumA:=1 to Min(MaxAxe , OBM.Detail.Count) do begin
      if Ventilable(CGen,NumA) then BEGIN
        if OBM.Detail[NumA-1].Detail.Count>0 then BEGIN
          for NumL:=0 to OBM.Detail[NumA-1].Detail.Count-1 do BEGIN
            OBA := TOBM(OBM.Detail[NumA-1].Detail[NumL]);
            OBA.PutMvt('Y_REFINTERNE',St) ;
          END ;
        END ;
      END ;
    end;
  END ;
  CommitTrans ;
  BRAAbandonClick(Nil) ;
end;

function TFSaisieEff.Load_Sais ( St : hString ) : Variant ;
Var V    : Variant ;
    Lig  : integer ;
    OBM  : TOBM ;
    CGen : TGGeneral ;
    CAux : TGTiers ;
BEGIN
  V:=#0 ;
  Result:=V ;
  St:=uppercase(Trim(St)) ;
  if St='' then Exit ;
  Lig:=QuelleLigne(St) ;
  if ((Lig=-1) and (CurLig>1)) then Lig:=CurLig-1
  else if Lig<=0 then Lig:=CurLig ;
  if Lig<=0 then Lig:=GS.Row ;
  OBM:=GetO(GS,Lig) ;
  CGen:=GetGGeneral(GS,Lig) ;
  CAux:=GetGTiers(GS,Lig) ;

  {Zones entête}
  if ((St='JOURNAL') or (St='E_JOURNAL')) then V:=SAJAL.JOURNAL else
    if ((St='DATECOMPTABLE') or (St='E_DATECOMPTABLE')) then V:=StrToDate(E_DATECOMPTABLE.Text) else
      if ((St='NATUREPIECE') or (St='E_NATUREPIECE')) then V:=E_NATUREPIECE.Value else
        if ((St='NUMEROPIECE') or (St='E_NUMEROPIECE')) then V:=NumPieceInt else
          if ((St='DEVISE') or (St='E_DEVISE')) then V:=E_DEVISE.Text else
            if ((St='ETABLISSEMENT') or (St='E_ETABLISSEMENT')) then V:=E_ETABLISSEMENT.Text else
              if St='E_GENERAL' then BEGIN if CGen<>Nil then V:=CGen.General ; END else
                if St='E_AUXILIAIRE' then BEGIN if CAux<>Nil then V:=CAux.auxi ; END else
{                  if St='TVA' then V:=Load_TVATPF(CGen,TRUE,'') else
                    if St='TVANOR' then V:=Load_TVATPF(CGen,TRUE,'NOR') else
                      if St='TVARED' then V:=Load_TVATPF(CGen,TRUE,'RED') else
                        if St='TPF' then V:=Load_TVATPF(CGen,False,'') else }
                          if St='SOLDE' then V:='SOLDE' else
{Zones Général}
  if Copy(St,1,2)='G_' then BEGIN
    if CGen=Nil then Exit ;
    System.Delete(St,1,2) ;
    if St='GENERAL' then V:=CGen.General else
    if St='LIBELLE' then V:=CGen.Libelle else
    if St='ABREGE' then V:=CGen.Abrege else
      V:=RechercheLente('G_'+St,CGen.General) ;
    END
{Zones Auxiliaire}
  else if Copy(St,1,2)='T_' then BEGIN
    if CAux=Nil then Exit ;
    System.Delete(St,1,2) ;
    if St='AUXILIAIRE' then V:=CAux.Auxi else
    if St='LIBELLE' then V:=CAux.Libelle else
    if St='ABREGE' then V:=CAux.Abrege else
      V:=RechercheLente('T_'+St,CAux.Auxi) ;
    END
{Zones Journal}
  else if Copy(St,1,2)='J_' then BEGIN
    if St='J_JOURNAL' then V:=SAJAL.Journal else
    if St='J_LIBELLE' then V:=SAJAL.LibelleJournal else
    if St='J_ABREGE' then V:=SAJAL.AbregeJournal else
      V:=RechercheLente('J_'+St,SAJAL.JOURNAL) ;
    END
{Comptes auto}
  else if Copy(St,1,4)='AUTO' then
    V:=TrouveAuto(SAJAL.COMPTEAUTOMAT,Ord(St[Length(St)])-48)
  else if St='INTITULE' then BEGIN
    if CAux<>Nil then V:=CAux.Libelle else
    if CGen<>Nil then V:=CGen.Libelle ;
    END
{Zones Ecriture}
  else BEGIN
    if OBM=Nil then Exit ;
    if Copy(St,1,2)='E_' then System.Delete(St,1,2) ;
    if St='REFERENCE' then St:='REFINTERNE' ;
    V:=OBM.GetMvt('E_'+St) ;
  END ;
  Load_Sais:=V ;
END ;

procedure TFSaisieEff.E_DATEECHEANCEExit(Sender: TObject);
Var OBM : TOBM ;
    i : Integer ;
    EcheAvant : tDateTime ;
    M : TMOD ;
begin
  DateEcheDefaut:=StrToDate(E_DATEECHEANCE.Text) ;
  If SAJAL<>NIL Then E_CPTTRESO.Caption:=SAJAL.Treso.Cpt ;
  for i:=1 to GS.RowCount-2 do BEGIN
    OBM:=GetO(GS,i) ;
    if OBM<>NIL then BEGIN
      If (EstLettrable(i) In [MonoEche,MultiEche]) Then BEGIN
        M:=TMOD(GS.Objects[SA_NumP,i]) ;
        If M<>NIL Then BEGIN
          EcheAvant:=M.MODR.TabEche[1].DateEche ;
          If (EcheAvant<>DateEcheDefaut) Then BEGIN
            M.MODR.TabEche[1].DateEche:=DateEcheDefaut ;
            If SA_DateEche>=0 Then GS.CellValues[SA_DateEche,i]:=E_DATEECHEANCE.Text ;
          END ;
        END ;
      END ;
    END ;
  END ;
end;

procedure TFSaisieEff.GereTagEnter(Sender: TObject);
Var B : TToolbarButton97 ;
begin
  B:=TToolbarButton97(Sender) ;
  if B=Nil then Exit ;
  if B.Parent=Outils then Outils.Tag:=1 ;
end;

procedure TFSaisieEff.GereTagExit(Sender: TObject);
Var B : TToolbarButton97 ;
begin
  B:=TToolbarButton97(Sender) ;
  if B=Nil then Exit ;
  if B.Parent=Outils then Outils.Tag:=0 ;
end;

procedure TFSaisieEff.BMenuZoomTMouseEnter(Sender: TObject);
begin
  PopZoom97(BMenuZoomT,POPZT) ;
end;

procedure TFSaisieEff.GereAffSolde(Sender: TObject);
Var Nam : String ;
    C   : THLabel ;
begin
  Nam:=THNumEdit(Sender).Name ; Nam:='L'+Nam ;
  C:=THLabel(FindComponent(Nam)) ;
  if C<>Nil then C.Caption:=THNumEdit(Sender).Text ;
end;

procedure TFSaisieEff.BParamListeSaisieEffClick(Sender: TObject);
Var St : String ;
BEGIN
  If EstRempli(GS,GS.Row) Then BEGIN
    // Vous devez vous positionner sur une ligne vide pour accéder au paramétrage de la grille de saisie.
    HDiv.Execute(29,Caption,'') ;
    Exit ;
  END ;
  Case ModeSaisie Of
    OnEffet : St:='CSF' ;
    OnChq : St:='CSG' ;
    OnCB : St:='CSH' ;
    OnBqe : If ModeSR=srCli Then St:='CSI'
                            Else St:='CSC';
  END ;
  ParamSaisieListeEff(St);
  AvertirCacheServer('LISTE');  // Rechargement de la liste en eAGL
  GS.ListeParam:='' ;
  GS.ListeParam:=QuelleListe ;
  HMTrad.ResizeGridColumns(GS) ;
  GS.Refresh ;
  If (GS.Titres.Count>0) And (GS.Tag=1) Then ReinitNumCol(GS.Titres[0]) ;
END ;

procedure TFSaisieEff.E_NATUREPIECEChange(Sender: TObject);
begin
  AutoCharged:=FALSE ;
  if GS.RowCount<=2 then BEGIN
    ChargeScenario ;
    ChargeLibelleAuto ;
  END ;
  GetSorteTva ;
  AfficheExigeTva ;
end;

procedure TFSaisieEff.E_NATUREPIECEExit(Sender: TObject);
begin
  OkPourLigne ;
  If (GS.RowCount>2) Then ModifNaturePiece ;
end;

procedure TFSaisieEff.E_TYPECTRChange(Sender: TObject);
Var OldTyp : String ;
begin
  If SAJAL=NIL then Exit ;
  OldTyp:=SAJAL.TRESO.STypCtr ;
  SAJAL.TRESO.STypCtr:=E_TYPECTR.Value ;
  SAJAL.TRESO.TypCtr:=StrToTypCtr(SAJAL.TRESO.STypCtr) ;
  If OldTyp<>SAJAL.TRESO.STypCtr then InitPiedTreso ;
  InitBoutonValide ;
  BComplementT.Enabled:=SAJAL.TRESO.TYPCTR In [PiedDC,PiedSolde] ;
end;

procedure TFSaisieEff.MajRefLibOBMT(Quoi : Integer) ;
Var St1,St2 : String ;
begin
  If SAJAL=NIL Then Exit ;
  If Quoi=0 Then BEGIN
    St1:='E_REFINTERNE' ;
    St2:=E_REFTRESO.Text ;
    END
  Else BEGIN
    St1:='E_LIBELLE'    ;
    St2:=E_LIBTRESO.Text ;
  END ;
  If SAJAL.TRESO.TypCtr In [PiedSolde,PiedDC] Then BEGIN
    if OBMT[1]<>NIL then OBMT[1].PutMvt(St1,St2) ;
    if OBMT[2]<>NIL then OBMT[2].PutMvt(St1,St2) ;
  END ;
end ;

procedure TFSaisieEff.E_REFTRESOExit(Sender: TObject);
begin
  MajRefLibOBMT(0) ;
end;

procedure TFSaisieEff.E_LIBTRESOExit(Sender: TObject);
begin
  MajRefLibOBMT(1) ;
  GS.SetFocus ;
end;

procedure TFSaisieEff.ClickSaisTaux ;
BEGIN
  if Not BSaisTaux.Enabled then Exit ;
  if DEV.Code=V_PGI.DevisePivot then Exit ;
  if DEV.Code='' then Exit ;
  if SaisieNewTaux2000(DEV,DatePiece) then BEGIN
    M.TauxD:=DEV.Taux ;
    Volatile:=True ;
  END ;
END ;

procedure TFSaisieEff.BSaisTauxClick(Sender: TObject);
begin
  ClickSaisTaux ;
end;

Procedure TFSaisieEff.ModifOBMPourTVAENC(i : Integer)  ;
Var O : TOBM ;
    CGen : TGGeneral ;
    Coll : String ;
    j : Integer ;
    OMODR : TMOD ;
BEGIN
  CGen:=Nil ;
  O:=GetO(GS,i) ;
  if GS.Objects[SA_Gen,i]<>Nil then CGen:=TGGeneral(GS.Objects[SA_Gen,i]) ;
  if VH^.OuiTvaEnc And (CGen<>NIL) And (O<>NIL)then BEGIN
    Coll:=CGen.General ;
    OMODR:=TMOD(GS.Objects[SA_NumP,i]) ;
    if OMODR=Nil then Exit ;
    if EstCollFact(Coll) then BEGIN
      if (SorteTva=stvDivers)  then BEGIN
        {Cycle bases HT sur échéances}
        for j:=1 to 4 do
          O.PutMvt('E_ECHEENC'+IntToStr(j),OMODR.MODR.TabEche[1].TAV[j]) ;
        O.PutMvt('E_ECHEDEBIT',OMODR.MODR.TabEche[1].TAV[5]) ;
      END ;
    END ;
  END ;
END ;


procedure TFSaisieEff.BPopTvaChange(Sender: TObject);
Var O : TOBM ;
    i : integer ;
    CGen : TGGeneral ;
    Ste,StEAvant  : String ;
    OMODR : TMOD ;
    BaseOntChange : Boolean ;
begin
  ChoixExige:=True ;
  GridEna(GS,False);
  ExigeEntete:=TExigeTva(BPopTva.ItemIndex-1) ;
  for i:=1 to GS.RowCount-2 do BEGIN
    if isTiers(GS,i) then BEGIN
      OMODR:=TMOD(GS.Objects[SA_NumP,i]) ;
      if OMODR=Nil then Continue ;
      OMODR.ModR.ModifTva:=False ;
      BaseOntChange:=GereTvaEncNonFact(i) ;
      If BaseOntChange then ModifOBMPourTvaEnc(i) ;
      END
    Else If isHT(GS,i,True) then BEGIN
      O:=GetO(GS,i) ;
      if O=Nil then Break ;
      CGen:=GetGGeneral(GS,i) ;
      if CGen=Nil then Break ;
      StEAvant:=O.GetMvt('E_TVAENCAISSEMENT') ;
      StE:=FlagEncais(ExigeEntete,CGen.TvaSurEncaissement) ;
      O.PutMvt('E_TVAENCAISSEMENT',StE) ;
    END ;
  END ;
  GridEna(GS,True) ;
end;


procedure TFSaisieEff.BComplementTClick(Sender: TObject);
begin
  ClickComplementT ;
end;

procedure TFSaisieEff.BVentilCtrClick(Sender: TObject);
Var OkL : Boolean ;
BEGIN
  Case SAJAL.TRESO.TypCtr Of
    Ligne : BEGIN
              if GS.Row>=GS.RowCount-1 then Exit ;
              if Not EstRempli(GS,GS.Row) then Exit ;
              OkL:=LigneCorrecte(GS.Row,False,True,TRUE) ;
              if OkL then GereAnal(GS.Row+1,False,False,FALSE,FALSE) ;
            END ;
    PiedDC,PiedSolde : BEGIN
                       END ;
  END ;
END ;

procedure TFSaisieEff.BZOOMFACT2Click(Sender: TObject);
begin
  ClickZoomFactConsult;
end;

procedure TFSaisieEff.POPZOOMFACTClick(Sender: TObject);
begin
  ClickZoomFact;
end;

{JP 20/10/05 : FQ 16923 : On applique le même système que dans Saisie.Pas, à savoir
               que l'on renseigne uniquement les lignes vides
{---------------------------------------------------------------------------------------}
procedure TFSaisieEff.GereRefLib(Lig : Integer; var StRef, StLib : string; LigOk : Boolean);
{---------------------------------------------------------------------------------------}
var
  StRefAvant,
  StLibAvant : string;
  OBM        : TOBM;
begin
  OBM := GetO(GS, Lig);

  if (OBM <> nil) and LigOk then begin
    StRefAvant := Trim(GS.Cells[SA_RefI, Lig]);
    StRef := GFormule(SI_FormuleRef, Load_Sais, nil, 1);
    StRef := Copy(StRef, 1, 35); {Modif SBO : si non tronqué --> plantage sql !}

    if (StRef <> '') then begin
      if (StRefAvant = '') then begin
        OBM.PutMvt('E_REFINTERNE', StRef);
        GS.Cells[SA_RefI, Lig] := StRef;
      end
      else
        {Remise à vide pour la ligne de contrepartie}
        StRef := '';
    end;

    StLibAvant := Trim(GS.Cells[SA_Lib, Lig]);
    StLib := GFormule(SI_FormuleLib, Load_Sais, nil, 1);
    StLib := Copy(StLib, 1, 35); {Modif SBO : si non tronqué --> plantage sql !}
    if (StLib <> '') and (StLib <> #0) then begin
      if (StLibAvant = '') then begin
        OBM.PutMvt('E_LIBELLE', StLib);
        GS.Cells[SA_Lib, Lig] := StLib;
      end
      else
        {Remise à vide pour la ligne de contrepartie}
        StLib := '';
    end;
  end

  {Ligne de treso en contrepartie ligne}
  else begin
    if (StRef <> '') then begin
      OBM.PutMvt('E_REFINTERNE', StRef);
      GS.Cells[SA_RefI, Lig] := StRef;
    end;

    if (StLib <> '') then begin
      OBM.PutMvt('E_LIBELLE', StLib); // Modif SBO : Fiche 12470
      GS.Cells[SA_Lib, Lig] := StLib;
    end;
  end;
end;




function TFSaisieEff.GetFormulePourCalc(Formule: hstring): Variant;
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
function TFSaisieEff.GetFormatPourCalc: string;
begin
  // FQ18371 et FQ18080
  result := '"#.###' ;
  if DEV.Decimale > 0 then
    result := result + ',' + Copy('000000000', 1, Dev.Decimale ) ;
  result := result + '"' ;
end;


{JP 14/05/07 : FQ 18324 : gestion du codeacceptation
{---------------------------------------------------------------------------------------}
function TFSaisieEff.GetModePaie(ModePaie : string) : TOB;
{---------------------------------------------------------------------------------------}
var
  lTob : Tob ;
  Q    : TQuery ;
begin
  Result := nil;
  if ModePaie = '' then Exit;

  Result := FTobModePaie.FindFirst(['MP_MODEPAIE'], [ModePaie], True);
  if not Assigned(Result) then begin
    Q := OpenSQL('SELECT MP_MODEPAIE, MP_CODEACCEPT, MP_CATEGORIE FROM MODEPAIE WHERE MP_MODEPAIE = "' + ModePaie + '"',True) ;
    try
      if not Q.EOF then begin
        lTob := Tob.Create('_MPF', FTobModePaie, -1);
        lTob.AddChampSupValeur('MP_MODEPAIE'  , Q.FindField('MP_MODEPAIE'  ).AsString);
        lTob.AddChampSupValeur('MP_CODEACCEPT', Q.FindField('MP_CODEACCEPT').AsString);
        lTob.AddChampSupValeur('MP_CATEGORIE' , Q.FindField('MP_CATEGORIE' ).AsString);
        Result := lTob;
      end;
    finally
      Ferme(Q);
    end;
  end;
end;


{JP 14/05/07 : FQ 18324 : gestion du codeacceptation
{---------------------------------------------------------------------------------------}
function TFSaisieEff.GetCodeAccept(ModePaie : string) : string;
{---------------------------------------------------------------------------------------}
var
  lTob : TOB;
begin
  Result := 'ACC';
  if ModePaie = '' then Exit;

  {Récupération des infos sur le mode de paiement}
  lTob := GetModePaie(ModePaie);
  if not Assigned(lTob) then Exit;
  {S'il s'agit d'une traite ...}
  if lTob.GetString('MP_CATEGORIE') = 'LCR' then
    {... Récupération du code acceptation}
    Result := lTob.GetString('MP_CODEACCEPT');
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
function TFSaisieEff.TestQteAna: boolean;
begin
  if ctxPCL in V_PGI.PGIContexte then
    result := False
  else if OkScenario then
    result := Scenario.GetValue('SC_CONTROLEQTE') = 'X'
  else
    result := GetParamSocSecur('SO_ZCTRLQTE', False) ;
end;

end.


