{***********UNITE*************************************************
Auteur  ...... : SG6 Passage en eAgl de la Saisie des réglements
Créé le ...... :   /  /
Modifié le ... : 16/11/2005
Description .. : JP 15/11/05 : FQ 16807 : alimentation des champs
Suite ........ : e_contrepartieXXX
Suite ........ :
Suite ........ : JP 15/11/05 : Gestion de l'analytique (cf. ##ANA##)
Suite ........ :
Suite ........ : JP 16/11/05 : Suppression de M.TRESO qui est toujours à
Suite ........ : True
Suite ........ :
Mots clefs ... :
*****************************************************************}
unit eSaisieTr;

interface

uses  
  Windows,
  Messages,
  SysUtils,
  Classes,
  Graphics,
  Controls,
  Forms,
  Dialogs,
  StdCtrls,
  Grids,
  Hctrls,
  Mask,
  ExtCtrls,
  ComCtrls,
  Buttons,
  Hspliter,
  Ent1,
  HCompte,
  {$IFDEF VER150}
  Variants,
  {$ENDIF}
  DB,
  HEnt1,
  HTB97,
  CPGeneraux_TOM,
  CPTiers_TOM,
  CPJournal_TOM,
  hmsgbox,
  HQry,
  Menus,
  {$IFDEF EAGLCLIENT}
   MainEagl,
  {$ELSE}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  Fe_Main,
  {$ENDIF EAGLCLIENT}
  CPMULSAISLETT_TOF,
  SaisTVA,
  EcheMono,
  eSaiSAnal,
  SaisComp,
  SaisUtil,
  SaisComm,  // RecupTotalPivot
  LettUtil,
  Choix,
  Echeance,
  About,
  SaisVisu,
  HStatus,
  Filtre,
  SaisTaux1,
  Devise_TOM,
  FichComm,
  SUIVCPTA_TOM,
  HDebug,
  HLines,
  ValPerio,
  HSysMenu,
  HPop97,
  SaisEnc,
  SaisBase,
  ed_tools,
  HPanel,
  UiUtil,
  {$IFDEF EAGLCLIENT}
  UtileAGL,
  {$ELSE}
  {$IFNDEF IMP}
  {$IFNDEF GCGC}
  EdtREtat,
  {$ELSE !GCGC}
  {$IFDEF CMPGIS35}
  EdtREtat,
  {$ENDIF CMPGIS35}
  {$ENDIF GCGC}
  {$ELSE !IMP}
  {$IFDEF CMPGIS35}
  EdtREtat,
  {$ENDIF CMPGIS35}
  {$ENDIF IMP}
  {$ENDIF EAGLCLIENT}
  ImgList,
  UtilSais,
  UtilPGI,
  CPChancell_TOF,
  UtilSOC,
  lookup,
  UTOB,
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  paramsoc,
  Constantes,  // ets_Rien, ets_Nouveau
  UlibEcriture, // FQ 13246 : SBO 30/03/2005
  ULibAnalytique, HImgList, TntButtons, TntStdCtrls, TntExtCtrls, TntGrids
  ;

procedure PrepareSaisTres;
procedure PrepareSaisTresEffet;
procedure PrepareSaisTresAgio(var M: RMVT);
procedure LanceSaisietr(Action: TActionFiche; var M: RMVT);
procedure LanceSaisietrAgio(Action: TActionFiche; var M: RMVT);

type
  TFSaisietr = class(TForm)
    DockTop: TDock97;
    DockRight: TDock97;
    DockLeft: TDock97;
    DockBottom: TDock97;
    GS: THGrid;
    PEntete: THPanel;
    E_JOURNAL: THValComboBox;
    H_JOURNAL: THLabel;
    H_DATECOMPTABLE: THLabel;
    E_NATUREPIECE: THValComboBox;
    H_NATUREPIECE: THLabel;
    E_ETABLISSEMENT: THValComboBox;
    H_ETABLISSEMENT: THLabel;
    H_NUMEROPIECE: THLabel;
    E_DEVISE: THValComboBox;
    H_DEVISE: THLabel;
    H_EXERCICE: THLabel;
    PPied: THPanel;
    G_LIBELLE: THLabel;
    H_SOLDE: THLabel;
    T_LIBELLE: THLabel;
    H_NUMEROPIECE_: THLabel;
    SA_SOLDEG: THNumEdit;
    SA_SOLDET: THNumEdit;
    SA_TOTALDEBIT: THNumEdit;
    SA_TOTALCREDIT: THNumEdit;
    SA_SOLDE: THNumEdit;
    E_NUMEROPIECE: THLabel;
    HLigne: THMsgBox;
    HPiece: THMsgBox;
    HTitres: THMsgBox;
    HDiv: THMsgBox;
    POPS: TPopupMenu;
    POPZ: TPopupMenu;
    BZoom: THBitBtn;
    BZoomJournal: THBitBtn;
    BZoomDevise: THBitBtn;
    BZoomEtabl: THBitBtn;
    PForce: TPanel;
    Label2: TLabel;
    Label3: TLabel;
    HForce: THNumEdit;
    Image1: TImage;
    FindSais: TFindDialog;
    BScenario: THBitBtn;
    BDernPieces: THBitBtn;
    BChancel: THBitBtn;
    H_TYPECTR: THLabel;
    E_TYPECTR: THValComboBox;
    H_SOLDET: THLabel;
    SA_SOLDETR: THNumEdit;
    PPiedTreso: THPanel;
    E_CREDITTRESO: THNumEdit;
    E_DEBITTRESO: THNumEdit;
    E_LIBTRESO: TEdit;
    E_REFTRESO: TEdit;
    E_CPTTRESO: THPanel;
    BMenuZoomT: TToolbarButton97;
    Bevel2: TBevel;
    Bevel1: TBevel;
    Bevel3: TBevel;
    Bevel4: TBevel;
    Bevel5: TBevel;
    E_MODEPAIE: THValComboBox;
    H_MODEPAIE: THLabel;
    E_NUMREF: THNumEdit;
    Bevel6: TBevel;
    Bevel7: TBevel;
    HE_MODEP: THLabel;
    HE_DATEECHE: THLabel;
    BSwapPivot: THBitBtn;
    BModifSerie: THBitBtn;
    BZoomT: THBitBtn;
    POPZT: TPopupMenu;
    BComplementT: THBitBtn;
    BSaisTaux: THBitBtn;
    HMTrad: THSystemMenu;
    EURO: TEdit;
    Valide97: TToolbar97;
    BValide: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BAide: TToolbarButton97;
    Outils: TToolbar97;
    BEche: TToolbarButton97;
    BVentil: TToolbarButton97;
    BComplement: TToolbarButton97;
    BChercher: TToolbarButton97;
    BRef: TToolbarButton97;
    BLig: TToolbarButton97;
    BMenuZoom: TToolbarButton97;
    BSolde: TToolbarButton97;
    ToolbarSep971: TToolbarSep97;
    ToolbarSep972: TToolbarSep97;
    E_RIB: THLabel;
    BModifRIB: THBitBtn;
    PRefAuto: TPanel;
    H_TitreRefAuto: TLabel;
    PFenGuide: TPanel;
    BRAValide: THBitBtn;
    BRAAbandon: THBitBtn;
    H_REF: THLabel;
    E_REF: TEdit;
    E_NUMDEP: TMaskEdit;
    H_DATEECHEANCE: THLabel;
    H_CONTREPARTIEGEN: THLabel;
    E_CONTREPARTIEGEN: THCpteEdit;
    HConf: TToolbarButton97;
    BPopTva: TPopButton97;
    TVAImages: THImageList;
    BModifTva: TToolbarButton97;
    ISigneEuro: TImage;
    LE_DEBITTRESO: THLabel;
    LE_CREDITTRESO: THLabel;
    LSA_SOLDEG: THLabel;
    LSA_SOLDET: THLabel;
    LSA_TOTALDEBIT: THLabel;
    LSA_TOTALCREDIT: THLabel;
    LSA_SOLDE: THLabel;
    LSA_SOLDETR: THLabel;
    BChoixRegime: THBitBtn;
    BModifs: THBitBtn;
    compte: THEdit;
    E_DATECOMPTABLE: THCritMaskEdit;
    E_DATEECHEANCE: THCritMaskEdit;
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure InitBoutonValide;
    procedure E_JOURNALChange(Sender: TObject);
    procedure E_DEVISEChange(Sender: TObject);
    procedure E_TYPECTRChange(Sender: TObject);
    procedure E_NATUREPIECEChange(Sender: TObject);
    procedure E_DATECOMPTABLEChange(Sender: TObject);
    procedure E_DATECOMPTABLEExit(Sender: TObject);
    procedure E_REFTRESOExit(Sender: TObject);
    procedure E_LIBTRESOExit(Sender: TObject);
    procedure H_JOURNALDblClick(Sender: TObject);
    procedure E_NUMDEPExit(Sender: TObject);
    procedure GSSetEditText(Sender: TObject; ACol, ARow: Longint; const Value: string);
    procedure GSExit(Sender: TObject);
    procedure GSMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure GSDblClick(Sender: TObject);
    procedure GSCellExit(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
    procedure GSCellEnter(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
    procedure GSRowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
    procedure GSRowExit(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
    procedure GSEnter(Sender: TObject);
    procedure BAbandonClick(Sender: TObject);
    procedure BValideClick(Sender: TObject);
    procedure BLigClick(Sender: TObject);
    procedure BVentilClick(Sender: TObject);
    procedure BEcheClick(Sender: TObject);
    procedure BZoomClick(Sender: TObject);
    procedure POPSPopup(Sender: TObject);
    procedure BSwapPivotClick(Sender: TObject);
    procedure BComplementClick(Sender: TObject);
    procedure E_ETABLISSEMENTChange(Sender: TObject);
    procedure BChercherClick(Sender: TObject);
    procedure FindSaisFind(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure BChancelClick(Sender: TObject);
    procedure BZoomDeviseClick(Sender: TObject);
    procedure BZoomEtablClick(Sender: TObject);
    procedure BZoomJournalClick(Sender: TObject);
    procedure BScenarioClick(Sender: TObject);
    procedure BRefClick(Sender: TObject);
    procedure H_MODEPAIEDblClick(Sender: TObject);
    procedure E_MODEPAIEExit(Sender: TObject);
    procedure E_NATUREPIECEExit(Sender: TObject);
    procedure E_TYPECTRExit(Sender: TObject);
    procedure BSoldeClick(Sender: TObject);
    procedure BModifSerieClick(Sender: TObject);
    procedure BZoomTClick(Sender: TObject);
    procedure BComplementTClick(Sender: TObject);
    procedure BSaisTauxClick(Sender: TObject);
    procedure E_DEVISEExit(Sender: TObject);
    procedure GSKeyPress(Sender: TObject; var Key: Char);
    procedure BAideClick(Sender: TObject);
    procedure E_JOURNALClick(Sender: TObject);
    procedure BMenuZoomMouseEnter(Sender: TObject);
    procedure BValideMouseEnter(Sender: TObject);
    procedure BValideMouseExit(Sender: TObject);
    procedure BMenuZoomMouseExit(Sender: TObject);
    procedure BModifRIBClick(Sender: TObject);
    procedure BRAAbandonClick(Sender: TObject);
    procedure BRAValideClick(Sender: TObject);
    procedure E_DATEECHEANCEExit(Sender: TObject);
    procedure E_CONTREPARTIEGENExit(Sender: TObject);
    procedure E_CONTREPARTIEGENChange(Sender: TObject);
    procedure BPopTvaChange(Sender: TObject);
    procedure GereTagEnter(Sender: TObject);
    procedure GereTagExit(Sender: TObject);
    procedure BModifTvaClick(Sender: TObject);
    procedure BMenuZoomTMouseEnter(Sender: TObject);
    procedure GereAffSolde(Sender: TObject);
    procedure BChoixRegimeClick(Sender: TObject);
    procedure E_DATECOMPTABLEEnter(Sender: TObject);
  private
    SAJAL: TSAJAL;
    DEV: RDEVISE;
    NumPieceInt, NbLigEcr, NbLigAna: Longint;
    CpteAuto, OkScenario, PieceConf, ModeRA, ChoixExige, NeedEdition: boolean;
    CurAuto, CurLig: integer;
    SorteTva: TSorteTva;
    ExigeTva, ExigeEntete: TExigeTva;
    TS, TDELECR, TDELANA: TList;
    M: RMVT;
    Scenario, Entete, ModifSerie: TOBM;
    GX, GY, LAUX, LHT: integer;
    OBMT: array[1..2] of TOBM; //TR
    SI_NumRef: LongInt;
    LigneTreso: Boolean; //TR
    PieceModifiee, ModeCG, RentreDate: boolean;
    GeneRegTVA, GeneTVA, GeneTPF, GestionRIB: String3;
    GeneSoumisTPF, RegimeForce: boolean;
    Revision, DejaRentre: Boolean;
    GeneTypeExo: TTypeExo;
    OkMessAnal, AutoCharged: boolean;
    ModeDevise, FindFirst: Boolean;
    NbLOrig, DecDev: integer;
    ModeForce: tmSais;
    NowFutur, DatePiece, DateEcheDefaut: TDateTime;
    ModeConf: string[1];
    OldDevise, DernVentiltype, EtatSaisie: string;
    Agio, Volatile: Boolean;
    RowEnterPourModeConsult: Boolean;
    MajEnCours: Boolean;
    MemoComp: HTStrings;
    WMinX, WMinY: Integer;
    tAppelLettrage: TBits;
    FInfo                      : TInfoEcriture ; // FQ 13246 : SBO 30/03/2005
    FBoInsertSpecif            : Boolean ;       // SBO 01/07/2007 : enregistrement spécifique pour pb CWAS
    FNeFermePlus               : Boolean; {JP 05/07/07 : FQ 19022 : Si on répond non à la confirmation fermeture}

    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    // Scénario de saisie
    procedure ChargeScenario;
    procedure RechargeMvtApresLettrage(Lig: Integer; OldGen, OldAux: string);
    procedure LettrageEnSaisieTreso(Lig: Integer; FromTouche: Boolean);
    procedure DeLettrageEnSaisieTreso(Lig: Integer);
    function DejaLettreDansPiece(Gen, Aux: string; i: Integer): Boolean;
    // Init et defaut
    function OkPourLigne: Boolean;
    function DateToJour(DD: TDateTime): string;
    procedure SwapNatureCompte;
    procedure InitLesComposants;
    procedure LectureSeule;
    procedure InitEnteteJal(Totale, Zoom, ReInit: boolean);
    procedure InitGrid;
    procedure DefautEntete;
    procedure DefautPied;
    procedure AttribNumeroDef;
    procedure AttribLeTitre;
    function EstLettrable(Lig: Integer): TLettre;
    procedure ReInitPiece;
    procedure InitPiedTreso;
    procedure GereEnabled(Lig: integer);
    procedure PosLesSoldes;
    // Euro
    procedure AddEuroFranc;
    procedure RestoreDeviseEuro;
    // Chargements
    procedure ChargeLignes;
    // Tva Enc
    function TvaTauxDirecteur(Lig: integer; Client, ToutDebit: boolean; Regime: String3): Boolean;
    function GereTvaEncNonFact(Lig: integer): Boolean;
    procedure GetSorteTVA;
    procedure RenseigneRegime(Lig: integer);
    procedure AffecteTva(C, L: integer);
    procedure AfficheExigeTva;
    procedure ModifOBMPourTVAENC(i: Integer);
    // Allocation et Désallocation
    procedure AlloueEcheAnal(C, L: LongInt);
    procedure AlloueEche(LaLig: LongInt);
    procedure DesalloueEche(LaLig: LongInt);
    procedure AlloueAnal(LaLig: LongInt);
    procedure DesalloueAnal(LaLig: LongInt);
    // Object Mouvement, Divers
    procedure AlimObjetMvt(Lig: integer; AvecM: boolean; Treso: Byte);
    procedure GereComplements(Lig: integer);
    function LettrageFaitEnSaisieTreso: Boolean;
    function OnPeutDetruire: Boolean;
    procedure DetruirePieceTresoEnCours;
    function FermerSaisie: boolean;
    procedure DecrementeNumero;
    procedure GotoEntete;
    procedure NumeroteVentils;
    procedure QuelNExo;
    procedure BasculeGS(AccesConsult: Boolean);
    function IsAccessible(Ou: LongInt): Boolean;
    procedure BloqueGrille(Bloc: boolean);
    function StatusOBMOk(OBM: TOBM): Boolean;
    procedure FocusSurNatP;
    // Divers
    procedure ValideCeQuePeux(Lig: longint);
    function FabricFromM: RMVT;
    procedure StockeLaPiece;
    procedure EditionSaisie;
    // Barre Outils
    procedure ClickVentil;
    procedure ClickEche;
    procedure ClickZoom(DoPopup: boolean = false);
    procedure ClickAbandon;
    procedure ClickSolde(Egal: boolean);
    procedure ClickSwapPivot;
    procedure ClickModifRIB;
    procedure ClickComplement;
    procedure ClickDelettrage;
    procedure Clicklettrage;
    procedure AlimOBMMemoireEnBatch(OSource, ODestination: TOBM);
    procedure ModifOBMTSerie(NomChamp: string; V: Variant);
    procedure ClickComplementT;
    procedure ClickCherche;
    procedure ClickRAZLigne;
    procedure ClickModifTva;
    procedure ClickChancel;
    procedure ClickEtabl;
    procedure ClickDevise;
    procedure ClickSaisTaux;
    procedure ClickChoixRegime;
    procedure ClickScenario;
    procedure ClickJournal;
    procedure ForcerMode(Cons: boolean; Key: word);
    // Calculs lignes
    procedure DetruitLigne(Lig: integer; Totale: boolean);
    function  SupprimeLigneEnBase( Lig : integer ) : boolean ;
    procedure PlusMvt(i: Integer; var SI_TDS, SI_TCS, SI_TDP, SI_TCP, SI_TDD, SI_TCD: Double);
    procedure CalculDebitCredit;
    procedure AfficheConf;
    procedure CalculSoldeCompte(CpteG, CpteA: string; DIG, CIG, DIA, CIA: Double);
    procedure SommeSoldeCompte(Col: integer; Cpte: string; var TD, TC: Double; Old, OkDev: Boolean);
    procedure DefautLigne(Lig: integer; Init: boolean);
    procedure TraiteMontant(Lig: integer; Calc: boolean);
    // Analytiques
    procedure OuvreAnal(Lig: integer; Scena: boolean; vNumAxe : integer );
    function  AOuvrirAnal(Lig: integer; Auto: boolean): Boolean;
    procedure GereAnal(Lig: integer; Auto, Scena: boolean);
    procedure RecupAnal(Lig: integer; var OBA: TOB; NumAxe, NumV: integer);
    // Echéances
    function AOuvrirEche(Lig: integer; var Cpte: String17; var MR: String3;
                         var OuvreAuto, RempliAuto: boolean; var t: TLettre): Boolean;
    procedure DebutOuvreEche(Lig: Integer; Cpte: String17; MR: String3; OMODR: TMOD);
    procedure FinOuvreEche(Lig: Integer; OMODR: TMOD; t: TLettre);
    procedure AlimEcheVersObm(Lig: Integer; ModePaie: String3; DateEche, DateValeur: TDateTime);
    procedure MajOBMViaEche(Lig: integer; OBM: TOBM; R: T_ModeRegl; NumEche: integer);
    procedure OuvreEche(Lig: integer; Cpte: String17; MR: String3; RempliAuto: boolean; t: TLettre;
      LigTresoEnPied: Boolean);
    procedure GereEche(Lig: integer; OuvreAuto, RempliAuto: boolean);
    procedure RecupEche(Lig: integer; R: T_ModeRegl; NumEche: integer; OBM: TOBM);
    // Contrôles
    function LigneCorrecte(Lig: integer; Totale, Alim, CtrlCptTreso: boolean): boolean;
    function MajIOComTreso: Boolean;
    function PieceCorrecte(OkMess: Boolean): boolean;
    function PasModif: boolean;
    function Equilibre: boolean;
    procedure ErreurSaisie(Err: integer; szApres: string = '');
    procedure AjusteLigne(Lig: integer);
    function PossibleRecupNum: Boolean;
    procedure ControleLignes;
    // Appels comptes
    function ChercheGen(C, L : integer; Force : boolean; DoPopup : boolean = false) : byte;
    function ChercheAux(C, L: integer; Force: boolean; DoPopup: boolean = false): byte;
    {JP 17/01/08 : ChargeOk : Pour savoir s'il faut charger le compte general}
    function ChargeCompteEffetJAL (ChargeOk : Boolean) : Boolean;
    procedure ChercheJAL(Jal: String3; Zoom: boolean);
    procedure ChercheMontant(Acol, Arow: longint);
    procedure AppelAuto(indice: integer);
    function AffecteRIB(C, L: integer; IsAux: Boolean): Boolean;
    procedure AffecteConso(C, L: integer);
    // MAJ Fichier
    procedure MajDirecteTable(Ou: LongInt; NatPiece: String3);
    procedure GetEcr(Lig: Integer);
    procedure GetEcrGrid(Lig: Integer);
    procedure GetAna(Lig: Integer);
    procedure RecupTronc(Lig: Integer; MO: TMOD; var OBM: TOBM);
    procedure RecupFromGrid(Lig: integer; MO: TMOD; NumE: integer);
    procedure ClickValide;
    procedure ValideLaPiece;
    procedure ValideLeReste;
    procedure ValideLignesGene;
    procedure ValideLignesAnal;
    procedure ValideLesComptes;
    procedure ValideLeJournalNew;
    procedure AttribRegimeEtTVA;
    procedure MajCptesGeneNew;
    procedure MajCptesAno;
    procedure MajCptesAuxNew;
    function QuelEnc(Lig: Integer): String3;
    procedure RetouchePiece;
    // Affichages, Positionnements
    procedure AutoSuivant(Suiv: boolean);
    procedure SoldelaLigne(Lig: integer);
    procedure AffichePied;
    procedure GereNewLigneT;
    procedure GereNewLigne1;
    procedure PositionneDevise(ReInit: boolean);
    procedure GereZoom;
    procedure AvertirPbTaux(Code: String3; DateTaux, DateCpt: TDateTime);
    function PbTaux(DEV: RDevise; DateCpt: TDateTime): boolean;
    // Calcul sur ligne des trésorerie
    function TrouveSensTreso(GS: THGrid; Nat: String3; Lig: integer): byte;
    procedure ClickLibelleAuto;
    procedure CalculRefAuto(Sender: TObject);
    procedure ChargeLibelleAuto;
    procedure ModifNaturePiece;
    procedure ModifModePaie;
    procedure AfficheLigneTreso;
    procedure GereLigneTreso(Lig: integer);
    procedure CalculSoldeCompteT;
    procedure AlimLigneTreso(Lig: Integer);
    procedure RecopieGS(i: Integer);
    function LigneEnPlusTreso(Sens: Integer; MSaisi, MPivot, MDevise: Double): boolean;
    function MajGridPiedTreso: boolean;
    procedure MajRefLibOBMT(Quoi: Integer);
    function Load_Sais(St: hstring): Variant;
    // Conversions, Caluls
    function ArrS(X: Double): Double;
    function DateMvt(Lig: integer): TDateTime;
    // Gestion référence entete
    procedure IncRef;
    procedure CloseFen;
    procedure SetTypeExo ; // FQ 16852 : SBO 30/03/2005
    function  TestQteAna : boolean ;
  public
    Action: TActionFiche;
    TPIECE: TList;
    SI_TotDS, SI_TotCS, SI_TotDP, SI_TotCP, SI_TotDD, SI_TotCD: Double;
    SI_FormuleRef, SI_FormuleLib: string;
    SC_TotDS, SC_TotCS, SC_TotDP, SC_TotCP, SC_TotDD, SC_TotCD: Double;
    {JP 20/10/05 : FQ 16923 : On applique le même système que dans Saisie.Pas, à savoir
                   que l'on renseigne uniquement les lignes vides}
    procedure GereRefLib(Lig : Integer; var StRef, StLib : string; LigOk : Boolean);
    {JP 14/11/05 : FQ 16807 : Gère la contrepartie (genéral et auxiliaire) sur le compte bancaire}
    procedure GetContrepartie(var Aux, Gen : string);
    // DEV 3216 10/04/2006 SBO
    function  GetFormulePourCalc(Formule: hstring): Variant;
    function  GetFormatPourCalc : string ;
  protected
    FClosing: Boolean ;
  end;

implementation

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  CPProcGen,
  CPVersion,
  {$ENDIF MODENT1}
  Formule
  {$IFNDEF EAGLCLIENT}
  , MenuOLG
  {$ENDIF EAGLCLIENT}
  ;

var
  VFTreso, VFEffet : Boolean;

  {$R *.DFM}

procedure PrepareSaisTres;
var
  M: RMVT;
begin
  FillChar(M, Sizeof(M), #0);
  M.Simul := V_PGI.User;
  M.CodeD := V_PGI.DevisePivot;
  M.DateC := V_PGI.DateEntree;
  M.TauxD := 1;
  M.DateTaux := M.DateC;
  M.Valide := False;
  M.Etabl := VH^.ETABLISDEFAUT;
  M.ANouveau := FALSE;
  M.Treso := TRUE;
  M.MajDirecte := TRUE;
  M.Effet := FALSE;
  LanceSaisietr(taCreat, M);
end;

procedure PrepareSaisTresEffet;
var
  M: RMVT;
begin
  FillChar(M, Sizeof(M), #0);
  M.Simul := V_PGI.User;
  M.CodeD := V_PGI.DevisePivot;
  M.DateC := V_PGI.DateEntree;
  M.TauxD := 1;
  M.DateTaux := M.DateC;
  M.Valide := False;
  M.Etabl := VH^.ETABLISDEFAUT;
  M.ANouveau := FALSE;
  M.Treso := TRUE;
  M.MajDirecte := TRUE;
  M.Effet := TRUE;
  LanceSaisietr(taCreat, M);
end;


procedure PrepareSaisEff;
var
  M: RMVT;
begin
  FillChar(M, Sizeof(M), #0);
  M.Simul := V_PGI.User;
  M.CodeD := V_PGI.DevisePivot;
  M.DateC := V_PGI.DateEntree;
  M.TauxD := 1;
  M.DateTaux := M.DateC;
  M.Valide := False;
  M.Etabl := VH^.ETABLISDEFAUT;
  M.ANouveau := FALSE;
  M.Treso := TRUE;
  M.MajDirecte := TRUE;
  M.Effet := TRUE;
  LanceSaisietr(taCreat, M);
end;


procedure PrepareSaisTresAgio(var M: RMVT);
begin
  M.Effet := FALSE;
  LanceSaisietrAgio(taCreat, M);
end;

procedure LanceSaisietrAgio(Action: TActionFiche; var M: RMVT);
var
  X: TFSaisietr;
  OA: TActionFiche;
begin
  OA := Action;
  case Action of
    taCreat:
      begin
        if PasCreerDate(M.DateC) then Exit;
        if DepasseLimiteDemo then Exit;
        if _Blocage(['nrCloture'], True, 'nrSaisieCreat') then Exit;
      end;
    taModif:
      begin
        if RevisionActive(M.DateC) then Exit;
        if _Blocage(['nrCloture', 'nrBatch'], True, 'nrSaisieModif') then Exit;
      end;
  end;
  //if ((Action=taCreat) and (PasCreerDate(V_PGI.DateEntree))) then Exit ;
  VFTreso := M.Treso;
  VFEffet := M.Effet;
  X := TFSaisietr.Create(Application);
  try
    X.NowFutur := NowH;
    X.Action := Action;
    X.M := M;
    X.Agio := TRUE;
    X.ShowModal;
    M := X.M;
  finally
    X.Free;
    case OA of
      taCreat: _Bloqueur('nrSaisieCreat', False);
      taModif: _Bloqueur('nrSaisieModif', False);
    end;
  end;
  Screen.Cursor := crDefault;
end;

procedure LanceSaisietr(Action: TActionFiche; var M: RMVT);
var
  X: TFSaisietr;
  OA: TActionFiche;
  PP: THpanel;
begin
  VFTreso := M.Treso;
  VFEffet := M.Effet;
  OA := Action;
  case Action of
    taCreat:
      begin
        if PasCreerDate(V_PGI.DateEntree) then Exit;
        if DepasseLimiteDemo then Exit;
        if _Blocage(['nrCloture'], True, 'nrSaisieCreat') then Exit;
      end;
    taModif:
      begin
        if RevisionActive(M.DateC) then Exit;
        if _Blocage(['nrCloture', 'nrBatch'], True, 'nrSaisieModif') then Exit;
      end;
  end;
  X := TFSaisietr.Create(Application);
  X.NowFutur := NowH;
  X.Action := Action;
  X.M := M;
  X.Agio := FALSE;
  PP := FindInsidePanel;
  if PP = nil then
  begin
    try
      X.ShowModal;
    finally
      X.Free;
      case OA of
        taCreat: _Bloqueur('nrSaisieCreat', False);
        taModif: _Bloqueur('nrSaisieModif', False);
      end;
    end;
    Screen.Cursor := crDefault;
  end else
  begin
    InitInside(X, PP);
    X.Show;
  end;
end;

procedure TFSaisietr.ChargeLibelleAuto;
var
  Q: TQuery;
begin
  if E_JOURNAL.Value = '' then Exit;
  if E_NATUREPIECE.Value = '' then Exit;
  if PasModif then Exit;
  Q := OpenSQL('Select RA_FORMULEREF, RA_FORMULELIB from REFAUTO Where RA_JOURNAL="' + E_JOURNAL.Value + '" AND RA_NATUREPIECE="' + E_NATUREPIECE.Value + '"', True);
  if not Q.EOF then
  begin
    SI_FormuleRef := Q.FindField('RA_FORMULEREF').AsString;
    SI_FormuleLib := Q.FindField('RA_FORMULELIB').AsString;
    AutoCharged := True;
    Q.Next ; // SBO 11/10/2005 report 10129 saisie pièce jamais fait !
    // N° 10129 le 13/06/2002
    If Not Q.Eof Then
      BEGIN
      SI_FormuleRef:='' ;
      SI_FormuleLib:='' ;
      AutoCharged:=FALSE ;
      END ;
  end;
  Ferme(Q);
end;

{==========================================================================================}
{================================== Chargements ===========================================}
{==========================================================================================}

procedure TFSaisietr.ChargeLignes;
begin
  CalculDebitCredit;
  if Action <> taCreat then
  begin
    GS.Row := 1;
    GS.Col := SA_Gen;
    GS.SetFocus;
  end;
  AffichePied;
  NbLOrig := GS.RowCount;
end;

{==========================================================================================}
{=========================== Allocations et Désallocations ================================}
{==========================================================================================}
procedure TFSaisietr.AlloueEcheAnal(C, L: LongInt);
var
  CGen: TGGeneral;
  CAux: TGTiers;
begin
  if C = SA_Gen then
  begin
    CGen := GetGGeneral(GS, L);
    if CGen = nil then Exit;
    if Lettrable(CGen) <> NonEche then AlloueEche(L);
    if Ventilable(CGen, 0) then AlloueAnal(L);
  end else
  begin
    CAux := GetGTiers(GS, L);
    if CAux = nil then Exit;
    if CAux.Lettrable then AlloueEche(L);
  end;
end;

procedure TFSaisietr.AlloueEche(LaLig: LongInt);
var //OBM : TOBM ;
  TTM: TMOD;
begin
  if GS.Objects[SA_NumP, LaLig] <> nil then Exit;
  TTM := TMOD.Create;
  GS.Objects[SA_NumP, LaLig] := TObject(TTM);
//  if M.Treso then
//  begin
    TMOD(GS.Objects[SA_NumP, LaLig]).MODR.TabEche[1].DateEche := DateMvt(LaLig);
    TMOD(GS.Objects[SA_NumP, LaLig]).MODR.TabEche[1].ModePaie := E_MODEPAIE.VALUE;
    //   OBM:=GetO(GS,LaLig) ;
//  end;
  Revision := False;
end;

procedure TFSaisietr.DesalloueEche(LaLig: LongInt);
var
  OBM: TOBM;
begin
  if GS.Objects[SA_NumP, LaLig] = nil then Exit;
  TMOD(GS.Objects[SA_NumP, LaLig]).Free;
  GS.Objects[SA_NumP, LaLig] := nil;
  OBM := GetO(GS, LaLig);
  if OBM = nil then Exit;
  // Re-init des infos echéances
  OBM.PutMvt('E_LETTRAGE', '');
  OBM.PutMvt('E_LETTRAGEDEV', '-');
  OBM.PutMvt('E_COUVERTURE', 0);
  OBM.PutMvt('E_COUVERTUREDEV', 0);
  OBM.PutMvt('E_MODEPAIE', '');
  OBM.PutMvt('E_NIVEAURELANCE', 0);
  OBM.PutMvt('E_ETATLETTRAGE', 'RI');
  OBM.PutMvt('E_NUMECHE', 0);
  OBM.PutMvt('E_ECHE', '-');
  OBM.PutMvt('E_DATEECHEANCE', IDate1900);
  OBM.PutMvt('E_DATERELANCE', IDate1900);
  Revision := False;
end;

function TFSaisietr.DateToJour(DD: TDateTime): string;
var
  y, m, d: Word;
  StD: string;
begin
  DecodeDate(DD, y, m, d);
  StD := IntToStr(d);
  if d < 10 then StD := '0' + StD;
  Result := StD;
end;

procedure TFSaisietr.AlloueAnal(LaLig: LongInt);
{var
  OBA: TObjAna;
begin
  if GS.Objects[SA_DateC, LaLig] <> nil then Exit;
  OBA := TObjAna.Create;
  GS.Objects[SA_DateC, LaLig] := TObject(OBA);
  Revision := False;}
var
  OBM : TOBM;
begin
  {##ANA##}
  OBM := GetO(GS,LaLig);
  if OBM = nil then Exit;
  AlloueAxe(OBM);
  Revision := False;
  {##ANA##}
end;

procedure TFSaisietr.DesalloueAnal(LaLig: LongInt);
var
  OBM: TOBM;
begin
  {if GS.Objects[SA_DateC, LaLig] = nil then Exit;
  TObjAna(GS.Objects[SA_DateC, LaLig]).Free;
  GS.Objects[SA_DateC, LaLig] := nil;
  Revision := False;
  OBM := GetO(GS, LaLig);
  if OBM = nil then Exit;
  // Re-init des infos Analytiques
  OBM.PutMvt('E_ANA', '-');}

  {##ANA##}
  OBM := GetO(GS,LaLig) ;
  if OBM.Detail.Count=0 then Exit ;
  OBM.ClearDetail;
  Revision:=False ;
  OBM.PutMvt('E_ANA', '-');
  {##ANA##}
end;

{==========================================================================================}
{=========================== Initialisations et valeurs par défaut ========================}
{==========================================================================================}

procedure TFSaisietr.SwapNatureCompte;
begin
  E_NATUREPIECE.Visible := not (M.Effet);
  H_NATUREPIECE.Visible := not (M.Effet);
  E_CONTREPARTIEGEN.Visible := (M.Effet);
  H_CONTREPARTIEGEN.Visible := (M.Effet);
end;

procedure TFSaisietr.InitLesComposants;
begin
  Bref.Visible := True;//(M.Treso);
  BLig.Enabled := True;//(M.Treso);
  H_TYPECTR.Visible := True;//(M.Treso);
  E_TYPECTR.Visible := True;//(M.Treso);
  H_SOLDET.Visible := True;//(M.Treso);
  LSA_SOLDETR.Visible := True;//(M.Treso);
  H_SOLDE.Visible := False;//not (M.Treso);
  LSA_SOLDE.Visible := False;//not (M.Treso);
  PPiedTreso.Visible := True;//(M.Treso);
  SwapNatureCompte;
  E_CONTREPARTIEGEN.Enabled := (M.Effet);
  H_CONTREPARTIEGEN.Enabled := (M.Effet);
  E_NATUREPIECE.Enabled := not (M.Effet);
  H_NATUREPIECE.Enabled := not (M.Effet);
end;

procedure TFSaisietr.GereEnabled(Lig: integer);
var
  OKL, Remp, OkRIB: Boolean;
  O: TOBM;
begin
  Remp := EstRempli(GS, Lig);
  O := GetO(GS, Lig);
  OkL := ((Remp) and (O <> nil));
  BEche.Enabled := (GS.Objects[SA_NumP, Lig] <> nil);
  {##ANA##}
  {if M.Treso then BVentil.Enabled := (GS.Objects[SA_DateC, Lig] <> nil) else
    BVentil.Enabled := ((GS.Objects[SA_DateC, Lig] <> nil) or (LigneEche(GS, Lig)));}
  BVentil.Enabled := ((O <> nil) and VentilationExiste(O));
  {##ANA##}

  BComplement.Enabled := ((OkL) and (not M.ANouveau));
  BChancel.Enabled := ((GS.RowCount <= 2) and (not EstRempli(GS, 1)) and (Action = taCreat));
  BSwapPivot.Enabled := (GS.RowCount > 2) and ((ModeDevise) or (ModeForce = tmPivot));
  BSaisTaux.Enabled := ((GS.RowCount <= 2) and (ModeDevise) and (Action = taCreat) {and (not EstMonnaieIN(DEV.Code))});
  {Gestion du RIB}
  OkRIB := OkScenario and ((GestionRIB = 'MAN') or (GestionRIB = 'PRI'));
  BModifRIB.Enabled := ((OkL) and (BEche.Enabled) and (OkRIB));
  //BRef.Enabled:=Not M.Effet ;
  BModifTva.Enabled := ((OkL) and (VH^.OuiTvaEnc));
  BChoixRegime.Enabled := (Action = taCreat);
end;

function TFSaisietr.EstLettrable(Lig: Integer): TLettre;
var
  CGen: TGGeneral;
  CAux: TGTiers;
begin
  Result := NonEche;
  CGen := GetGGeneral(GS, Lig);
  if CGen = nil then exit;
  Result := Lettrable(CGen);
  if Result in [MonoEche, MultiEche] then Exit;
  CAux := GetGTiers(GS, Lig);
  if CAux = nil then Exit;
  if CAux.Lettrable then Result := MultiEche;
end;

procedure TFSaisietr.ReInitPiece;
var
  bc: boolean;
begin
  GS.VidePile(True);
  tAppelLettrage.Size := 0;
  if Action <> taCreat then Exit;
  // Ré-init des infos comptes / auxi pour rechargement des totaux...
  if Assigned( FInfo )  then // FQ 13246 SBO 04/08/2005
    begin
    FInfo.Compte.Clear ;
    FInfo.Aux.Clear ; 
    end ;
  BValide.Enabled := True;
  DefautPied;
  {if M.TReso then} NowFutur := NowH;
  H_NUMEROPIECE.Visible := False;
  H_NUMEROPIECE_.Visible := True;
  SwapNatureCompte;
  E_DATECOMPTABLE.Enabled := True;
  if M.Effet then
  begin
    E_CONTREPARTIEGEN.Enabled := TRUE;
  end else
  begin
    E_NATUREPIECE.Enabled := TRUE;
  end;
  E_JOURNAL.Enabled := True;
  InitEnteteJal(False, False, True);
  DateEcheDefaut := 0;
  DateEcheDefaut := StrToDate(E_DATEECHEANCE.Text);
  ModifSerie.Free;
  ModifSerie := nil;
  ModifSerie := TOBM.Create(EcrGen, '', True);
//  if M.Treso then
//  begin
    E_TYPECTR.Enabled := TRUE;
    OBMT[1].Free;
    OBMT[1] := nil;
    OBMT[1] := TOBM.Create(EcrGen, '', TRUE);
    OBMT[2].Free;
    OBMT[2] := nil;
    OBMT[2] := TOBM.Create(EcrGen, '', TRUE);
    AlimObjetMvt(0, FALSE, 1);
    AlimObjetMvt(0, FALSE, 2);
//  end;
  PieceModifiee := False;
  OkMessAnal := False;
  DejaRentre := False;
  RegimeForce := False;
  GS.SetFocus;
  GS.Col := SA_GEN;
  GS.Row := 1;
  PieceConf := False;
  RentreDate := False;
  VideListe(TDELECR);
  VideListe(TDELANA);
  ModeConf := '0';
  Volatile := False;
  (*
  if Action=taCreat then GSEnter(Nil) ; GSRowEnter(Nil,1,bc,False) ;
  *)
  {Affichage}
  if ((VH^.CPDateObli) and (Action = taCreat) and (E_DATECOMPTABLE.CanFocus)) then E_DATECOMPTABLE.SetFocus else
    if ((OkScenario) and (Scenario <> nil) and (Action = taCreat) and (E_DATECOMPTABLE.CanFocus)) then E_DATECOMPTABLE.SetFocus else
  begin
    if Action = taCreat then GSEnter(nil);
    GSRowEnter(nil, 1, bc, False);
  end;
end;

procedure TFSaisietr.InitPiedTreso;
begin
  E_CPTTRESO.Visible := TRUE;
  E_REFTRESO.Visible := TRUE;
  E_LIBTRESO.Visible := TRUE;
  LE_DEBITTRESO.Visible := TRUE;
  LE_CREDITTRESO.Visible := TRUE;
  PPiedTreso.Visible := TRUE;
  E_REFTRESO.Enabled := FALSE;
  E_LIBTRESO.Enabled := FALSE;
  E_REFTRESO.Color := clBtnFace;
  E_LIBTRESO.Color := clBtnFace;
  BLig.Enabled := FALSE;
  case SAJAL.TRESO.TypCtr of
    PiedDC, PiedSolde:
      begin
        E_REFTRESO.Enabled := TRUE;
        E_LIBTRESO.Enabled := TRUE;
        E_REFTRESO.Text := '';
        E_LIBTRESO.Text := '';
        E_REFTRESO.Color := clWhite;
        E_LIBTRESO.Color := clWhite;
      end;
    Ligne:
      begin
        (*
        E_CPTTRESO.Visible:=FALSE   ;
        E_DEBITTRESO.Visible:=FALSE ; E_CREDITTRESO.Visible:=FALSE ; PPiedTreso.Visible:=FALSE ;
        *)
        E_REFTRESO.Visible := FALSE;
        E_LIBTRESO.Visible := FALSE;
        BLig.Enabled := TRUE;
      end;
    AuChoix:
      begin
        E_CPTTRESO.Visible := FALSE;
        LE_DEBITTRESO.Visible := FALSE;
        LE_CREDITTRESO.Visible := FALSE;
        PPiedTreso.Visible := FALSE;
      end;
  end;
end;

procedure TFSaisietr.InitEnteteJal(Totale, Zoom, ReInit: boolean);
var
  MM: String17;
  DD: TDateTime;
begin
  if Action <> taCreat then Exit;
  if SAJAL = nil then Exit;
  PositionneDevise(ReInit);
  if Totale then
  begin
//    if M.Treso then
  //  begin
      if M.Effet then
      begin
        E_NaturePiece.DataType := 'ttNaturePiece';
        E_NaturePiece.Value := 'OD';
      end else
      begin
        E_NaturePiece.DataType := 'ttNatPieceBanque';
        E_NaturePiece.Value := 'RC';
      end;
      if Agio then E_NaturePiece.Value := 'OD';
    {end else
    begin
      case CaseNatJal(SAJAL.NatureJal) of
        tzJVente:
          begin
            E_NaturePiece.DataType := 'ttNatPieceVente';
            E_NaturePiece.Value := 'FC';
          end;
        tzJAchat:
          begin
            E_NaturePiece.DataType := 'ttNatPieceAchat';
            E_NaturePiece.Value := 'FF';
          end;
        tzJBanque:
          begin
            E_NaturePiece.DataType := 'ttNatPieceBanque';
            E_NaturePiece.Value := 'RC';
          end;
        tzJOD:
          begin
            E_NaturePiece.DataType := 'ttNaturePiece';
            E_NaturePiece.Value := 'OD';
          end;
      end;
    end;}
  end;
  if not Zoom then
  begin
    DD := StrToDate(E_DATECOMPTABLE.Text);
//    if M.Treso then
  //  begin
      NumPieceInt := GetNum(EcrGen, SAJAL.COMPTEURNORMAL, MM, DD);
      E_CPTTRESO.Caption := SAJAL.Treso.Cpt;

      {JP 17/11/05 : FQ 16120 : Initialisation du compte de contrepartie}
      if M.Effet then
        E_CONTREPARTIEGEN.Text := SAJAL.TRESO.Cpt;

      E_TYPECTR.VALUE := SAJAL.TRESO.STypCtr;
      if ((not ReInit) and (not M.Effet)) then E_DEVISE.VALUE := SAJAL.TRESO.DevBqe;
      InitPiedTreso;
    {end else
    begin
      if M.Simul <> 'N' then NumPieceInt := GetNum(EcrGen, SAJAL.COMPTEURSIMUL, MM, DD)
      else NumPieceInt := GetNum(EcrGen, SAJAL.COMPTEURNORMAL, MM, DD);
    end;}
    InitGrid;
  end;
  if NumPieceInt > 0 then E_NUMEROPIECE.Caption := IntToStr(NumPieceInt);
end;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 29/07/2003
Modifié le ... :   /  /
Description .. : LG - 29/07/2003 - reenitialisation des colonnes de la grille
Mots clefs ... :
*****************************************************************}
procedure TFSaisietr.DefautEntete;
begin
  InitDefautColSaisie;
  // FQ 13041 // Modif SBO : pb d'affichage du numéro de ligne
//  if M.Treso then //TR
  //begin

    SA_DateC := 3;
    SA_NumL := 2;
//  end; // Fin Modif SBO : pb d'affichage du numéro de ligne
  DateEcheDefaut := 0;
  E_DATECOMPTABLE.Text := DateToStr(M.DateC);
  E_DATEECHEANCE.Text := E_DATECOMPTABLE.Text;
  DateEcheDefaut := M.DateC;
  SI_NumRef := -1;
  E_NUMREF.Visible := FALSE;
  E_JOURNAL.Value := M.JAL;
  if Agio then E_JOURNALChange(nil);
  if not Agio then E_NATUREPIECE.Value := M.Nature;
  E_ETABLISSEMENT.Value := M.Etabl;
  E_ETABLISSEMENT.Enabled := VH^.EtablisCpta;
  if Action = taCreat then PositionneEtabUser(E_ETABLISSEMENT);
  DatePiece := M.DateC;
  SetTypeExo ; // FQ 16852 : SBO 30/03/2005
  if M.Effet then H_SOLDET.Caption := HDiv.Mess[19] else H_SOLDET.Caption := HDiv.Mess[18];
  E_DEVISE.Value := M.CodeD;
  case Action of
    taCreat:
      begin
        if ((not Agio) or (M.DateC < V_PGI.DateDebutEuro)) then E_DEVISE.Enabled := False;
        if M.Effet then E_DEVISE.Enabled := TRUE;
        if not Agio then
        begin
          E_NumeroPiece.Caption := '';
          GS.Enabled := False;
        end;
        H_NUMEROPIECE.Visible := False;
        H_NUMEROPIECE_.Visible := True;
        E_REF.Text := '';
        E_NUMDEP.Text := '';
      end;
    taModif:
      begin
        E_JOURNAL.Enabled := False;
        E_DEVISE.Enabled := False;
        E_DATECOMPTABLE.Enabled := False;
        E_NATUREPIECE.Enabled := False;
        E_NumeroPiece.Caption := InttoStr(M.Num);
        NumPieceInt := M.Num;
        H_NUMEROPIECE.Visible := True;
        H_NUMEROPIECE_.Visible := False;
        E_MODEPAIE.Enabled := FALSE;
        E_TYPECTR.Enabled := FALSE;
        GS.Enabled := True;
        GS.SetFocus;
      end;
    taConsult:
      begin
        E_NumeroPiece.Caption := InttoStr(M.Num);
        NumPieceInt := M.Num;
        H_NUMEROPIECE.Visible := True;
        H_NUMEROPIECE_.Visible := False;
        GS.Enabled := True;
        PEntete.Enabled := False;
      end;
  end;
  // TR
  if OBMT[1] = nil then OBMT[1] := TOBM.Create(EcrGen, '', TRUE);
  if OBMT[2] = nil then OBMT[2] := TOBM.Create(EcrGen, '', TRUE);
end;

procedure TFSaisietr.DefautPied;
begin
  SI_TotDS := 0;
  SI_TotCS := 0;
  SI_TotDP := 0;
  SI_TotCP := 0;
  SI_TotDD := 0;
  SI_TotCD := 0;
  SC_TotDS := 0;
  SC_TotCS := 0;
  SC_TotDP := 0;
  SC_TotCP := 0;
  SC_TotDD := 0;
  SC_TotCD := 0;
  ZeroBlanc(PPied);
  ZeroBlanc(PPiedTreso);
end;

procedure TFSaisietr.InitGrid;
begin
  if Action <> taCreat then Exit;
  GS.Enabled := True;
  DefautLigne(GS.RowCount - 1, True);
  GS.Col := SA_Gen;
  GS.Row := 1;
  { GP 02/07/96 GS.SetFocus ; }
end;

procedure TFSaisietr.AttribLeTitre;
//var
//  i, j: integer;
//  C: Char;
begin
//  i := 1;
//  j := 1;

//  if M.Treso then
//  begin
    if M.Effet then
    begin
      // saisie des ecritures d'effet
      Caption := HTitres.Mess[31];
      HelpContext := 7245000;
    end
    else
    begin
      // saisie des ecritures de reglement
      Caption := HTitres.Mess[30];
      HelpContext := 7247000;
    end;
  {end
  else
  begin
    case Action of
      taConsult: i := 1;
      taModif: i := 2;
      taCreat: i := 3;
    end;
    C := M.Simul[1];
    case C of
      'N': j := 1;
      'S': j := 2;
      'P': j := 3;
      'U': j := 4;
    end;
    Caption := HTitres.Mess[10 * (i - 1) + (j - 1)];
  end;}
  Caption := TraduireMemoire(Caption); // FQ 22279 : Pb traduction
  UpdateCaption(Self);
end;

procedure TFSaisietr.AttribNumeroDef;
var
  Facturier: String3;
  Lig, NumAxe, NumV: integer;
  O : TOBM;
  {##ANA## OBA: TObjAna;
  T : TList;}
  DD: TDateTime;
begin
  if Action <> taCreat then Exit;
  //if M.Treso then
    Facturier := SAJAL.CompteurNormal;
  {else
  begin
    if M.Simul <> 'N' then Facturier := SAJAL.CompteurSimul else Facturier := SAJAL.CompteurNormal;
  end;}
  DD := StrToDate(E_DATECOMPTABLE.Text);
  SetIncNum(EcrGen, Facturier, NumPieceInt, DD);
  H_NUMEROPIECE.Visible := True;
  H_NUMEROPIECE_.Visible := False;
  E_NUMEROPIECE.Caption := IntToStr(NumPieceInt);
  {Attribution aux objets du nouveau numéro}
  for Lig := 1 to GS.RowCount - 1 do
  begin
    GS.Cells[SA_NumP, Lig] := IntToStr(NumPieceInt);
    O := GetO(GS, Lig);
    if O = nil then Break;
    O.PutMvt('E_NUMEROPIECE', NumPieceInt);

    {##ANA##}
    {OBA := TObjAna(GS.Objects[SA_DateC, Lig]);
    if OBA = nil then Continue;
    OBA.NumeroPiece := NumPieceInt;
    for NumAxe := 1 to MaxAxe do if OBA.AA[NumAxe] <> nil then
    begin
      T := OBA.AA[NumAxe].L;
      for NumV := 0 to T.Count - 1 do PutMvtA(OBA.AA[NumAxe], P_TV(T.Items[NumV]).F, 'Y_NUMEROPIECE', NumPieceInt);
    end;}
    {JP 15/11/05 : Nouveau code avec une Tob et sans objet Analytique}
    for NumAxe := 1 to O.Detail.Count do
      for NumV := 0 to O.Detail[NumAxe - 1].Detail.Count - 1 do
        TOBM(O.Detail[NumAxe - 1].Detail[NumV]).PutMvt('Y_NUMEROPIECE', NumPieceInt) ;

    {##ANA##}
  end;
end;

{==========================================================================================}
{================================== Scenario de saisie ====================================}
{==========================================================================================}
procedure TFSaisietr.RechargeMvtApresLettrage(Lig: Integer; OldGen, OldAux: string);
var
  OBM: TOBM;
  SQL, SQL2, SQL1, SQL3: string;
  M: TMOD;
  R: T_ModeRegl;
  CGen: TGGeneral;
  CAux: TGTiers;
  Q: TQuery;
  OkLet, OkCpt: Boolean;
  Eche : array[1..4] of Double;
  RegTva, CodeTva : String;
  k : integer;
  EcheDeb, ValMax : Double;

begin
  OkCpt := FALSE;
  OBM := GetO(GS, Lig);
  if OBM = nil then Exit;
  M := TMOD(GS.Objects[SA_NumP, Lig]);
  if M = nil then Exit;
  if M <> nil then R := M.MODR;
  CGen := GetGGeneral(GS, Lig);
  if CGen = nil then Exit;
  CAux := GetGTiers(GS, Lig);
  if CAux = nil then OkLet := CGen.Lettrable else OkLet := CAux.Lettrable;
  if not OkLet then Exit;
  if (CGen.General = OldGen) then
  begin
    if (CAux = nil) and (OldAux = '') then OkCpt := TRUE else
      if (CAux <> nil) and (CAux.Auxi = OldAux) then OkCpt := TRUE;
  end;
  if not OkCpt then Exit;
  SQL1 := 'UPDATE ECRITURE SET E_DATEMODIF="' + UsTime(NowFutur) + '"';
  SQL2 := ' Where  ' + WhereEcriture(tsGene, OBMToIdent(OBM, FALSE), FALSE);
  SQL3 := ' AND E_NUMLIGNE=' + IntToStr(Lig) + ' AND E_NUMECHE=' + IntToStr(1);
  SQL := SQL1 + SQL2 + SQL3;
  ExecuteSQL(SQL);
  SQL1 := 'SELECT * FROM ECRITURE Where  ' + WhereEcriture(tsGene, OBMToIdent(OBM, FALSE), FALSE);
  SQL2 := ' AND E_NUMLIGNE=' + IntToStr(Lig) + ' AND E_NUMECHE=' + IntToStr(1);
  SQL3 := ' ORDER BY E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE';
  SQL := SQL1 + SQL2 + SQL3;
  BeginTrans;
  Q := OpenSQL(SQL, TRUE);
  if not Q.Eof then
  begin
    OBM.ChargeMvt(Q);
    if M <> nil then
    begin
      with R.TabEche[1] do
      begin
        Couverture := OBM.GetMvt('E_COUVERTURE');
        CouvertureDev := OBM.GetMvt('E_COUVERTUREDEV');
        //         CouvertureEuro:=OBM.GetMvt('E_COUVERTUREEURO') ;
       //          CouvertureEuro:=0 ;
        CodeLettre := OBM.GetMvt('E_LETTRAGE');
        LettrageDev := OBM.GetMvt('E_LETTRAGEDEV');
        DatePaquetMax := OBM.GetMvt('E_DATEPAQUETMIN');
        DatePaquetMin := OBM.GetMvt('E_DATEPAQUETMAX');
        EtatLettrage := OBM.GetMvt('E_ETATLETTRAGE');
      end;
      M.MODR := R;
    end;
  end;
  Ferme(Q);
  CommitTrans;

  BeginTrans;
  ///////////////////////////////////////////////////////////////
  {09/10/2007 YMO Reprise des infos Tva sur encaissement, reprise des infos des lignes d'origine après lettrage}

  SQL:='SELECT * FROM ECRITURE'
  +' WHERE E_GENERAL="'+OBM.GetMvt('E_GENERAL')+'"'
  +' AND E_AUXILIAIRE="'+OBM.GetMvt('E_AUXILIAIRE')+'"'
  +' AND E_LETTRAGE="'+OBM.GetMvt('E_LETTRAGE')+'"'
  +' AND E_ETATLETTRAGE="'+OBM.GetMvt('E_ETATLETTRAGE')+'"'
  +' AND E_NUMEROPIECE<>"'+floattostr(OBM.GetMvt('E_NUMEROPIECE'))+'"'
  +' ORDER BY E_NUMEROPIECE, E_NUMLIGNE';

  Q := OpenSQL(SQL, TRUE);

  For k:=1 to 4 do Eche[k]:=0;
  EcheDeb:=0;
  RegTva:='';
  CodeTva:='';
  ValMax:=0;
  while not Q.EOF do
  begin
    For k:=1 to 4 do
    begin
        Eche[k] := Eche[k]+Q.FindField('E_ECHEENC'+inttostr(k)).AsFloat;
        If Q.FindField('E_ECHEENC'+inttostr(k)).AsFloat>ValMax then
        begin
          CodeTva:=Q.FindField('E_TVA').AsString;
          RegTva:=Q.FindField('E_REGIMETVA').AsString;
          ValMax:=Q.FindField('E_ECHEENC'+inttostr(k)).AsFloat;
        end;
    end;
    EcheDeb:=EcheDeb+Q.FindField('E_ECHEDEBIT').AsFloat;

    Q.Next ;
  end;
  Ferme(Q);

  SQL := 'UPDATE ECRITURE'  {07/11/2007 YMO Correctif format}
  + ' SET E_ECHEENC1=' + StrFPoint(Eche[1])
  + ', E_ECHEENC2=' + StrFPoint(Eche[2])
  + ', E_ECHEENC3=' + StrFPoint(Eche[3])
  + ', E_ECHEENC4=' + StrFPoint(Eche[4])
  + ', E_ECHEDEBIT='+ StrFPoint(EcheDeb)
  + ', E_TVA="'      + CodeTva + '"'
  + ', E_REGIMETVA="'+ RegTva + '"'
  + ' WHERE  ' + WhereEcriture(tsGene, OBMToIdent(OBM, FALSE), FALSE)
  + ' AND E_GENERAL="' + E_CONTREPARTIEGEN.Text + '" AND E_NUMECHE=' + IntToStr(1);

  ExecuteSQL(SQL);

  ///////////////////////////////////////////////////////////////

  CommitTrans;
end;

function TFSaisietr.DejaLettreDansPiece(Gen, Aux: string; i: Integer): Boolean;
var
  Lig: Integer;
  OBM: TOBM;
  Aux2, Gen2: string;
  M: TMod;
  R: T_ModeRegl;
begin
  Result := FALSE;
  for Lig := 1 to GS.RowCount - 2 do
  begin
    OBM := GetO(GS, Lig);
    if (OBM <> nil) and (i <> Lig) then
    begin
      Aux2 := GS.Cells[SA_Aux, Lig];
      Gen2 := GS.Cells[SA_Gen, Lig];
      if (Aux2 = Aux) and (Gen2 = Gen) and (OBM.GetMvt('E_TRESOLETTRE') = 'X') then
      begin
        M := TMOD(GS.Objects[SA_NumP, Lig]);
        if M <> nil then
        begin
          R := M.MODR;
          if (R.TabEche[1].EtatLettrage = 'PL') and (OBM.GetMvt('E_TRESOLETTRE') = 'X') then
          begin
            Result := TRUE;
            Break;
          end;
        end;
      end;
    end;
  end;
end;

procedure TFSaisietr.LettrageEnSaisieTreso(Lig: Integer; FromTouche: Boolean);
var
  X: RMVT;
  OBM: TOBM;
  M: TMod;
  R: T_ModeRegl;
  SansLesPartiel: Boolean;
  Lig1: Integer;
  OldGen, OldAux: string;
  bc: Boolean;
begin
  if Action = taConsult then Exit;
  if not Self.M.MajDirecte then Exit;
  if not Self.M.Treso then Exit;
  OBM := GetO(GS, Lig);
  if OBM = nil then Exit;
  M := TMOD(GS.Objects[SA_NumP, Lig]);
  if M = nil then Exit;
  R := M.MODR;
  if (R.TabEche[1].EtatLettrage = 'TL') or
    (R.TabEche[1].Couverture = R.TabEche[1].MontantP) then Exit;
  if (R.TabEche[1].EtatLettrage = 'RI') then Exit;
  OldGen := OBM.GetMvt('E_GENERAL');
  OldAux := OBM.GetMvt('E_AUXILIAIRE');
  SansLesPartiel := DejaLettreDansPiece(OldGen, OldAux, Lig);
  X.Axe := '';
  X.Etabl := E_ETABLISSEMENT.Value;
  X.Treso := Self.M.treso;
  X.Jal := SAJAL.Journal;
  X.Exo := QuelExoDT(StrToDate(E_DATECOMPTABLE.Text));
  X.CodeD := E_DEVISE.Value;
  X.Simul := Self.M.Simul;
  X.Nature := E_NATUREPIECE.Value;
  X.DateC := StrToDate(E_DATECOMPTABLE.Text);
  X.DateTaux := DEV.DateTaux;
  X.Num := NumPieceInt;
  X.TauxD := DEV.Taux;
  X.Valide := False;
  X.ANouveau := False;
  X.NumLigne := Lig;
  X.NumEche := 1;
  if HPiece.Execute(24, caption, '') = mrYes then
  begin
    LettrerEnSaisieTreso(X, GS.Cells[SA_Gen, Lig], GS.Cells[SA_Aux, Lig], SansLesPartiel);
    GridEna(GS, FALSE);
    GS.CacheEdit;
    for Lig1 := 1 to GS.RowCount - 2 do RechargeMvtApresLettrage(Lig1, OldGen, OldAux);
    GridEna(GS, TRUE);
    if FromTouche then GSRowEnter(nil, 1, bc, False);
    GS.MontreEdit;
  end;
  Screen.Cursor := SyncrDefault;
end;

procedure TFSaisietr.DeLettrageEnSaisieTreso(Lig: Integer);
var
  X: RMVT;
  OBM: TOBM;
  M: TMod;
  R: T_ModeRegl;
  CodeLettre: string;
  OldGen, OldAux: string;
  Lig1: Integer;
  bc: Boolean;
begin
  if Action = taConsult then Exit;
  if not Self.M.MajDirecte then Exit;
  if not Self.M.Treso then Exit;
  OBM := GetO(GS, Lig);
  if OBM = nil then Exit;
  M := TMOD(GS.Objects[SA_NumP, Lig]);
  if M = nil then Exit;
  R := M.MODR;
  if (R.TabEche[1].EtatLettrage = 'RI') or (R.TabEche[1].EtatLettrage = 'AL') then Exit;
  //SansLesPartiel:=DejaLettreDansPiece(OBM.GetMvt('E_GENERAL'),OBM.GetMvt('E_AUXILIAIRE'),Lig) ;
  X.Axe := '';
  X.Etabl := E_ETABLISSEMENT.Value;
  X.Treso := Self.M.treso;
  X.Jal := SaJal.Journal;
  X.Exo := QuelExoDT(StrToDate(E_DATECOMPTABLE.Text));
  X.CodeD := E_DEVISE.Value;
  X.Simul := Self.M.Simul;
  X.Nature := E_NATUREPIECE.Value;
  X.DateC := StrToDate(E_DATECOMPTABLE.Text);
  X.DateTaux := DEV.DateTaux;
  X.Num := NumPieceInt;
  X.TauxD := DEV.Taux;
  X.Valide := False;
  X.ANouveau := False;
  X.NumLigne := Lig;
  X.NumEche := 1;
  CodeLettre := R.TabEche[1].CodeLettre;
  OldGen := OBM.GetMvt('E_GENERAL');
  OldAux := OBM.GetMvt('E_AUXILIAIRE');
  if HPiece.Execute(29, caption, '') = mrYes then
  begin
    DeLettrerEnSaisieTreso(X, GS.Cells[SA_Gen, Lig], GS.Cells[SA_Aux, Lig], CodeLettre, FALSE);
    for Lig1 := 1 to GS.RowCount - 2 do RechargeMvtApresLettrage(Lig1, OldGen, OldAux);
    GSRowEnter(nil, 1, bc, False);
  end;
  Screen.Cursor := SyncrDefault;
end;

procedure TFSaisietr.ChargeScenario;
var
  Q: TQuery;
  STVA, SQL: string;
  StComp: string[10];
begin
  Scenario.PutDefautDivers;
  OkScenario := False;
  DejaRentre := False;
  Entete.Free;
  Entete := TOBM.Create(EcrGen, '', True);
  GestionRIB := '...';
  MemoComp.Free;
  MemoComp := HTStringList.Create;
  if PasModif then Exit;
  if M.ANouveau then Exit;
  if ((E_JOURNAL.Value = '') or (E_NATUREPIECE.Value = '')) then Exit;
  SQL := 'Select * from SUIVCPTA Where SC_JOURNAL="' + E_JOURNAL.Value + '" AND SC_NATUREPIECE="' + E_NATUREPIECE.Value
    + '" AND (SC_USERGRP="" OR SC_USERGRP="' + V_PGI.Groupe + '" OR SC_USERGRP="...")';
  BeginTrans;
  Q := OpenSQL(SQL, True);
  if not Q.EOF then
  begin
    Scenario.ChargeMvt(Q);
    if not TMemoField(Q.FindField('SC_ATTRCOMP')).IsNull then
{$IFDEF EAGLCLIENT}
   	 MemoComp.Text := (TMemoField(Q.FindField('SC_ATTRCOMP')).AsString) ;
{$ELSE}
   	 MemoComp.Assign(TMemoField(Q.FindField('SC_ATTRCOMP'))) ;
{$ENDIF}
    OkScenario := True;
  end;
  Ferme(Q);
  CommitTrans;
  if OkScenario then
  begin
    StComp := Scenario2Comp(Scenario);
    Entete.PutMvt('E_ETAT', StComp);
    STVA := Scenario.GetMvt('SC_CONTROLETVA');
    GestionRIB := Scenario.GetMvt('SC_RIB');
    if Scenario.GetMvt('SC_BROUILLARD') = 'X' then
    begin
      NeedEdition := True;
      EtatSaisie := '';
    end;
    if Scenario.GetMvt('SC_DOCUMENT') <> '' then
    begin
      NeedEdition := True;
      EtatSaisie := Scenario.GetMvt('SC_DOCUMENT');
    end;
  end else
  begin
    GestionRIB := '...';
  end;
  //if M.Treso then Scenario.PutMvt('SC_OUVREECHE','X') ;
end;

{==========================================================================================}
{============================== Traitements liés à l'entête ===============================}
{==========================================================================================}

procedure TFSaisietr.E_TYPECTRExit(Sender: TObject);
begin
  OkPourLigne;
end;

procedure TFSaisietr.E_NATUREPIECEExit(Sender: TObject);
begin
  OkPourLigne;
  if {M.Treso and} (GS.RowCount > 2) then ModifNaturePiece;
end;

procedure TFSaisietr.E_MODEPAIEExit(Sender: TObject);
begin
  if RowEnterPourModeConsult then Exit;
  OkPourLigne;
  if {M.Treso and} (GS.RowCount > 2) then ModifModePaie;
end;

procedure TFSaisietr.E_DEVISEExit(Sender: TObject);
begin
  if csDestroying in ComponentState then Exit;
  OkPourLigne;
end;

function TFSaisietr.OkPourLigne: boolean;
var
  DateCpt: TDateTime;
begin
  result := true;
  if Action <> taCreat then Exit;
//  if M.Treso then
//  begin
    if E_MODEPAIE.Text = '' then
    begin
      if (Screen.ActiveControl.Name) = 'GS' then
      begin
        HDiv.Execute(8, caption, '');
        E_MODEPAIE.SetFocus;
        Screen.Cursor := SyncrDefault;
        result := false;
      end;
    end else if E_TYPECTR.Text = '' then
    begin
      if (Screen.ActiveControl.Name) = 'GS' then
      begin
        HDiv.Execute(9, caption, '');
        E_TYPECTR.SetFocus;
        Screen.Cursor := SyncrDefault;
        result := false;
      end;
    end else if (SAJAL <> nil) and SAJAL.Treso.Ventilable then
    begin
      if M.Effet then
      begin
        HDiv.Execute(20, caption, '');
        E_CONTREPARTIEGEN.SetFocus;
      end else
      begin
        HDiv.Execute(10, caption, '');
        E_JOURNAL.SetFocus;
      end;
      Screen.Cursor := SyncrDefault;
      result := false;
    end else if (SAJAL <> nil) and (SAJAL.NATUREJAL = 'BQE') and (PasOkDevBqe(SAJAL.TRESO.DevBqe, E_DEVISE.Value)) then
    begin
      HDiv.Execute(11, caption, '');
      E_DEVISE.SetFocus;
      Screen.Cursor := SyncrDefault;
      result := false;
    end;
    if Result then
      if ((DEV.Code <> V_PGI.DevisePivot) and (Action = taCreat) and (not Volatile)) then
      begin
        DateCpt := StrToDate(E_DATECOMPTABLE.Text);
        DEV.Taux := GetTaux(DEV.Code, DEV.DateTaux, DateCpt);
        if ((DEV.DateTaux <> DateCpt) or (PbTaux(DEV, DateCpt))) then AvertirPbTaux(DEV.Code, DEV.DateTaux, DateCpt);
      end;
    if Result and M.Effet then
    begin
      if E_CONTREPARTIEGEN.Text = '' then
        if (Screen.ActiveControl.Name) = 'GS' then
        begin
          HDiv.Execute(15, caption, '');
          E_CONTREPARTIEGEN.SetFocus;
          Screen.Cursor := SyncrDefault;
          result := false;
        end;
    end;
    if Result and M.Effet then
    begin
      if (SAJAL <> nil) and SAJAL.Treso.Collectif and (SAJAL.Treso.TypCtr in [PiedDC, PiedSolde]) then
      begin
        HDiv.Execute(16, caption, '');
        E_TYPECTR.SetFocus;
        Screen.Cursor := SyncrDefault;
        result := false;
      end;
    end;
    if Result and M.Effet then
    begin
      if (Screen.ActiveControl.Name) = 'GS' then
      begin
        {JP 17/01/08 : FQ 18563 : Calcul du solde en fonction du compte et surtout de l'exercice}
        SAJAL.PutDate(E_DATECOMPTABLE.Text);

        if not SAJAL.ChargeCompteTreso then
        begin
          HDiv.Execute(17, caption, '');
          E_CONTREPARTIEGEN.SetFocus;
          Screen.Cursor := SyncrDefault;
          result := false;
        end;
      end;
    end;
  //end;
end;

procedure TFSaisietr.E_NATUREPIECEChange(Sender: TObject);
begin
  PieceModifiee := True;
  AutoCharged := FALSE;
  if GS.RowCount <= 2 then
  begin
    ChargeScenario;
    ChargeLibelleAuto;
  end;
  GetSorteTva;
  AfficheExigeTva;
end;

procedure TFSaisietr.QuelNExo;
var
  EXO: string;
begin
  if Action = taCreat then EXO := QuelExoDT(StrToDate(E_DATECOMPTABLE.Text)) else EXO := M.EXO;
  if EXO = VH^.Encours.Code then H_EXERCICE.Caption := '(N)' else
    if EXO = VH^.Suivant.Code then H_EXERCICE.Caption := '(N+1)' else
    H_EXERCICE.Caption := '(N-1)';
end;

procedure TFSaisietr.E_DATECOMPTABLEChange(Sender: TObject);
begin
  PieceModifiee := True;
end;

procedure TFSaisietr.E_ETABLISSEMENTChange(Sender: TObject);
begin
  PieceModifiee := True;
end;

procedure TFSaisietr.InitBoutonValide;
begin
  if SAJAL = nil then Exit;
//  if M.Treso then
  //begin
    BSolde.Enabled := SAJAL.TRESO.TypCtr = AuChoix;
    SAJAL.COMPTEAUTOMAT := '';
    SAJAL.NbAuto := 0;
    if SAJAL.TRESO.TypCtr = AuChoix then
    begin
      SAJAL.NbAuto := 1;
      SAJAL.COMPTEAUTOMAT := SAJAL.TRESO.CPT + ';';
    end;
  //end;
end;

procedure TFSaisietr.E_JOURNALChange(Sender: TObject);
var
  Jal: string;
label
  0;
begin
  Jal := E_JOURNAL.Value;
  if Jal = '' then
  begin
    SAJAL.Free;
    SAJAL := nil;
    goto 0;
  end;
  // 14611
  if SAJAL <> nil then
  begin
    if SAJAL.TRESO.Cpt <> '' then
    begin
      // Le compte du journal est fermé
      if ExisteSQL('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL="' + SAJAL.TRESO.Cpt + '" AND G_FERME="X"') then
      begin
        HLigne.Execute(30, '', '(' + SAJAL.TRESO.Cpt + ')'); // Le compte est fermé.
        E_Journal.ItemIndex := -1;
        E_Journal.SetFocus;
        exit;
      end;
    end;
    if SAJAL.JOURNAL = Jal then goto 0;
  end;

  if ((VH^.JalAutorises <> '') and (Pos(';' + Jal + ';', VH^.JalAutorises) <= 0)) then
  begin
    HPiece.Execute(26, caption, '');
    if SAJAL <> nil then
    begin
      SAJAL.Free;
      SAJAL := nil;
    end;
    E_JOURNAL.Value := '';
    if Action = taCreat then
    begin
      E_JOURNAL.SetFocus;
      GS.Enabled := False;
    end;
    goto 0;
  end;
  if SAJAL <> nil then if ((Jal = VH^.JalATP) or (Jal = VH^.JalVTP)) then
    begin
      HPiece.Execute(33, caption, '');
      if SAJAL <> nil then
      begin
        SAJAL.Free;
        SAJAL := nil;
      end;
      E_JOURNAL.Value := '';
      if Action = taCreat then
      begin
        E_JOURNAL.SetFocus;
        GS.Enabled := False;
      end;
      Exit;
    end;
  ChercheJal(Jal, False);
  if ((Action = taCreat) and (NumPieceInt = 0)) then
  begin
    HPiece.Execute(27, caption, '');
    if SAJAL <> nil then
    begin
      SAJAL.Free;
      SAJAL := nil;
    end;
    E_JOURNAL.Value := '';
    E_JOURNAL.SetFocus;
    GS.Enabled := False;
    goto 0;
  end;
  CalculSoldeCompteT;
//  if M.Treso then //TR
//  begin
    AlimObjetMvt(0, FALSE, 1);
    AlimObjetMvt(0, FALSE, 2);
//  end;
  InitBoutonValide;
  0:
end;

procedure TFSaisietr.AvertirPbTaux(Code: String3; DateTaux, DateCpt: TDateTime);
var
  ii: integer;
  OkTaux1: boolean;
begin
  if Action <> taCreat then Exit;
  OkTaux1 := (Arrondi(DEV.Taux - 1, ADecimP) = 0);
  if ((OkScenario) and (not OkTaux1)) then
  begin
    if Scenario.GetMvt('SC_ALERTEDEV') = 'AUC' then Exit else
      if Scenario.GetMvt('SC_ALERTEDEV') = 'MOI' then
    begin
      if GetPeriode(DateTaux) = GetPeriode(DateCpt) then Exit;
    end else
      if Scenario.GetMvt('SC_ALERTEDEV') = 'SEM' then
    begin
      if NumSemaine(DateTaux) = NumSemaine(DateCpt) then Exit;
    end;
  end;
  (*
  if ((EstMonnaieIN(Code)) and (DateCpt >= V_PGI.DateDebutEuro)) then
  begin
    ii := HDiv.Execute(23, Caption, '');
    if ii <> mrYes then ii := HDiv.Execute(24, caption, '');
    if ii = mrYes then
    begin
      FicheDevise(Code, taModif, False);
      DEV.Taux := GetTaux(DEV.Code, DEV.DateTaux, DateCpt);
    end;
  end else
  *)
  begin
    if OkTaux1 then ii := HDiv.Execute(25, caption, '') else ii := HDiv.Execute(2, caption, '');
    if ii = mrYes then
    begin
      FicheChancel(E_DEVISE.VALUE, True, DateCpt, taCreat, (DateCpt >= V_PGI.DateDebutEuro));
      DEV.Taux := GetTaux(DEV.Code, DEV.DateTaux, DateCpt);
    end;
  end;
end;

function TFSaisietr.PbTaux(DEV: RDevise; DateCpt: TDateTime): boolean;
var
  Code: string3;
begin
  Result := False;
  Code := DEV.Code;
  if ((Code = V_PGI.DevisePivot) or (Code = V_PGI.DeviseFongible)) then Exit;
  if ((DateCpt < V_PGI.DateDebutEuro) or (Code<>V_PGI.DevisePivot) {(Not EstMonnaieIn(Code))}) then Result := (DEV.Taux = 1)
//  else if EstMonnaieIn(Code) then Result := (DEV.Taux = V_PGI.TauxEuro);
end;

procedure TFSaisietr.E_DEVISEChange(Sender: TObject);
var
  DateCpt: TDateTime;
begin
  if E_DEVISE.Value = '' then
  begin
    FillChar(DEV, Sizeof(DEV), #0);
    Exit;
  end;
  DEV.Code := E_DEVISE.Value;
  GetInfosDevise(DEV);
  DecDev := DEV.Decimale;
  case Action of
    taCreat:
      begin
        if ((DEV.Code = V_PGI.DevisePivot) {or (EstMonnaieIN(DEV.Code))}) then Volatile := False;
        DateCpt := StrToDate(E_DATECOMPTABLE.Text);
        if not Volatile then
        begin
          DEV.Taux := GetTaux(DEV.Code, DEV.DateTaux, DateCpt);
          //if not M.Treso then
            //if (ModeForce = tmNormal) and (E_DEVISE.Enabled) and ((DEV.DateTaux <> DateCpt) or (PbTaux(DEV, DateCpt))) then AvertirPbTaux(DEV.Code, DEV.DateTaux, DateCpt);
        end;
        BSaisTaux.Enabled := ((GS.RowCount <= 2) and
          (E_DEVISE.Value <> V_PGI.DevisePivot) and
          (Action = taCreat));
      end;
    taModif:
      begin
        DEV.Taux := M.TauxD;
        DEV.DateTaux := M.DateTaux;
      end;
    taConsult: DEV.Taux := M.TauxD;
  end;
  ChangeFormatDevise(Self, DECDEV, DEV.Symbole);
  PieceModifiee := True;
  ModeDevise := EnDevise(DEV.Code);
  BSwapPivot.Enabled := ((ModeDevise) or (ModeForce = tmPivot));
end;

procedure TFSaisietr.H_JOURNALDblClick(Sender: TObject);
begin
  ClickJournal;
end;

procedure TFSaisietr.E_TYPECTRChange(Sender: TObject);
var
  OldTyp: string;
begin
  if RowEnterPourModeConsult then Exit;
  if SAJAL = nil then Exit;
  OldTyp := SAJAL.TRESO.STypCtr;
  SAJAL.TRESO.STypCtr := E_TYPECTR.Value;
  SAJAL.TRESO.TypCtr := StrToTypCtr(SAJAL.TRESO.STypCtr);
  if OldTyp <> SAJAL.TRESO.STypCtr then InitPiedTreso;
  InitBoutonValide;
  BComplementT.Enabled := SAJAL.TRESO.TYPCTR in [PiedDC, PiedSolde];
end;

procedure TFSaisietr.E_DATECOMPTABLEEnter(Sender: TObject);
begin
  RentreDate := True;
end;

procedure TFSaisietr.E_DATECOMPTABLEExit(Sender: TObject);
var
  Err, i: integer;
  DD: TDateTime;
  StD: string;
begin
  if csDestroying in ComponentState then Exit;
  Err := ControleDate(E_DATECOMPTABLE.Text);
  if Err > 0 then
  begin
    HPiece.Execute(15 + Err, caption, '');
    E_DATECOMPTABLE.SetFocus;
    Exit;
  end;
  if ((Action = taCreat) and (RevisionActive(StrToDate(E_DATECOMPTABLE.Text)))) then
  begin
    E_DATECOMPTABLE.SetFocus;
    Exit;
  end;
  QuelNExo;
  DatePiece := StrToDate(E_DATECOMPTABLE.Text);
  SetTypeExo ; // FQ 16852 : SBO 30/03/2005
  OkPourLigne;
  DD := StrToDate(E_DATECOMPTABLE.Text);
  StD := DateToJour(DD);
  for i := 1 to GS.RowCount - 1 do GS.Cells[SA_DateC, i] := StD;

  {JP 17/01/08 : FQ 18563 : Calcul du solde en fonction du compte et surtout de l'exercice}
  SAJAL.PutDate(E_DATECOMPTABLE.Text);
  SAJAL.ChargeCompteTreso;

  // Recalcul du pied pour solde des comptes à dates FQ13246 SBO 03/08/2005
  AffichePied ;
end;

{==========================================================================================}
{=================================== ANALYTIQUES ==========================================}
{==========================================================================================}
function TFSaisietr.AOuvrirAnal(Lig: integer; Auto: boolean): Boolean;
var
  CGen : TGGeneral;
  {##ANA## OBA: TObjAna;}
  OBM  : TOBM ;
  XD,
  XP,
  Deb  : Double;
  Sens : integer;
begin
  Result := False;
  CGen := GetGGeneral(GS, Lig);
  if CGen = nil then exit;
  if not Ventilable(CGen, 0) then Exit;
  {##ANA##}
  {if Auto then
  begin
    OBA := TObjAna(GS.Objects[SA_DateC, Lig]);
    if OBA = nil then Exit;
    Deb := ValD(GS, Lig);
    if Deb <> 0 then Sens := 1 else Sens := 2;
    XD := MontantLigne(GS, Lig, tsmDevise);
    XP := MontantLigne(GS, Lig, tsmPivot);
    if (OBA.TotalDevise = XD) then Exit;
    if OBA.General = CGen.General then
    begin
      RecalculProrataAnal('Y', OBA, XP, XD, 0, Sens);
      Exit;
    end;
  end;}

  {JP 15/11/05 : Nouveau code avec une Tob et sans objet Analytique}
  if Auto then begin
    OBM := GetO(GS,Lig);
    if OBM = nil then Exit;

    Deb := ValD(GS,Lig);
    if Deb <> 0 then Sens := 1
                else Sens := 2;

    XD := MontantLigne(GS, Lig, tsmDevise);
    XP := MontantLigne(GS, Lig, tsmPivot);

    if ((GetTotalVentilDev(OBM) = XD) and (GetTotalVentil(OBM) = XP) and (Sens = GetSensVentil(OBM))) then Exit;

    if GetVentilGeneral(OBM) = CGen.General then begin
      RecalculProrataAnalNEW('Y', OBM, 0, DEV);
      Exit;
    end;
  end;
  {##ANA##}

  // GP inspecter le détail du TList et faire la somme
  Result := True;
end;

procedure TFSaisietr.GereAnal(Lig: integer; Auto, Scena: Boolean);
var  bOuvreAnal : Boolean;
     OBM        : TOBM;
     lNumAxe    : integer ;
begin
  if AOuvrirAnal(Lig, Auto) then
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
      OuvreAnal(Lig, Scena,lNumAxe);
    end ;
end;

procedure TFSaisietr.OuvreAnal( Lig : integer ; Scena : boolean ; vNumAxe : integer );
var
  OBM: TOBM;
  XD, XP, VV: Double;
  C : integer;
  Ax: string;
  RecA: ARG_ANA;
begin
  //Remplissage du record pour la ventilation
  RecA.QuelEcr := EcrGen;
  RecA.GuideA := '';
  if DernVentilType <> '' then RecA.DernVentilType := DernVentilType;
  RecA.ControleBudget := ((OkScenario) and (Action <> taConsult) and (Scenario.GetMvt('SC_CONTROLEBUDGET') = 'X'));
  RecA.ModeConf := ModeConf;
  RecA.CC := GetGGeneral(GS, Lig);
  RecA.DEV := DEV;
  RecA.NumLigneDecal := 0;
  RecA.VerifQte    := ((OkScenario) and (Action<>taConsult) and (Scenario.GetValue('SC_CONTROLEQTE')='X'));

  //préparation tob des axes analytiques
  OBM := GetO(GS, Lig);

{$IFDEF OLDVENTIL}
  AlloueAxe( OBM ) ;// SBO 25/01/2006
  RecA.VerifVentil := False;
  RecA.Action      := Action;
  vNumAxe := 0;
{$ELSE  OLDVENTIL}
{  if VentilationExiste( OBM )
      then RecA.Action := taModif
      else
} RecA.Action := taCreat ;
  RecA.Verifventil	:= True ;
  if vNumAxe = 0 then
{$ENDIF OLDVENTIL}
    if ((OkScenario) and (Scena) and (Scenario.GetMvt('SC_OUVREANAL') <> 'X')) then
    begin
      Ax := Trim(Scenario.GetMvt('SC_NUMAXE'));
      if Length(Ax) < 2 then Exit;
      vNumAxe := Ord(Ax[2]) - 48;
    end;

  RecA.NumGeneAxe     := vNumAxe;
  RecA.Info           := FInfo ;
  RecA.MessageCompta  := nil ;

  //  AlimAnalDef(Lig, OBA);
  XD := MontantLigne(GS, Lig, tsmDevise);
  XP := MontantLigne(GS, Lig, tsmPivot);

  //on lance la fiche des axes analytiques
  eSaisieAnal(TOB(OBM), RecA);


  if RecA.DernVentilType <> '' then DernVentilType := RecA.DernVentilType;
  if RecA.ModeConf <> ModeConf then
  begin
    if RecA.ModeConf > ModeConf then ModeConf := RecA.ModeConf;
    CalculDebitCredit;
  end;
  if Action = taConsult then Exit;

  // Test montant écriture modifiée en saisie analytique
  if ((XD = GetMontantDev(OBM)) and (XP = GetMontant(OBM))) then Exit;
  VV := GetMontantDev(OBM);
  if GetSens(OBM) = 1
    then C := SA_Debit
  else C := SA_Credit;
  GS.Cells[C, Lig] := StrS(VV, DECDEV);
  FormatMontant(GS, C, Lig, DECDEV);
  TraiteMontant(Lig, True);
end;

{==========================================================================================}
{=================================== ECHEANCES ============================================}
{==========================================================================================}
procedure TFSaisietr.GereEche(Lig: integer; OuvreAuto, RempliAuto: Boolean);
var
  Cpte: String17;
  MR: String3;
  t: TLettre;
  FromLigne: Boolean;
begin
  FromLigne := ((OuvreAuto) and (not RempliAuto));
  if AOuvrirEche(Lig, Cpte, MR, OuvreAuto, RempliAuto, t) then
  begin
    {if Not M.Treso Then }if ((OkScenario) and (FromLigne)) then if Scenario.GetMvt('SC_OUVREECHE') <> 'X' then RempliAuto := True;
    OuvreEche(Lig, Cpte, MR, RempliAuto, t, FALSE);
  end;
end;

procedure TFSaisietr.DebutOuvreEche(Lig: Integer; Cpte: String17; MR: String3; OMODR: TMOD);
begin
  with OMODR.MODR do
  begin
    TotalAPayerP := MontantLigne(GS, Lig, tsmPivot);
    TotalAPayerD := MontantLigne(GS, Lig, tsmDevise);
    CodeDevise := DEV.Code;
    TauxDevise := DEV.Taux;
    Quotite := DEV.Quotite;
    Decimale := DECDEV;
    Symbole := DEV.Symbole;
    DateFact := DateMvt(Lig);
    DateFactExt := DateFact;
    DateBL := DateFact;
    Aux := Cpte;
    ModeInitial := MR;
  end;
end;

procedure TFSaisietr.FinOuvreEche(Lig: Integer; OMODR: TMOD; t: TLettre);
var i    : Integer;
    CGen : TGGeneral ;
begin

  CGen:=GetGGeneral(GS,Lig) ;
  if CGen=Nil then Exit ;

  with OMODR.MODR do for i := 1 to MaxEche do with TabEche[i] do
      begin
        Couverture := 0;
        CouvertureDev := 0;
        CodeLettre := '';
        LettrageDev := '-';
        DatePaquetMax := DateMvt(Lig);
        DatePaquetMin := DateMvt(Lig);
        DateRelance := IDate1900;
        NiveauRelance := 0;

        // FQ 19462 : pb état lettrage sur les compte BQE... SBO 28/06/2007....Annule les modif ?? : // FQ 18895 : pb d'état lettrage sur compte TIC / TID SBO 07/11/2006
        if ( CGen.Lettrable and ( CGen.NatureGene = 'DIV' ) ) or ( EstLettrable(Lig) = MultiEche )
          then EtatLettrage:='AL'
          else EtatLettrage:='RI' ;

      end;
end;

procedure TFSaisietr.AlimEcheVersObm(Lig: Integer; ModePaie: String3; DateEche, DateValeur: TDateTime);
var
  OBM: TOBM;
begin
  //if not M.Treso then Exit;
  OBM := GetO(GS, Lig);
  if OBM <> nil then
  begin
    OBM.PutMvt('E_MODEPAIE', ModePaie);
    OBM.PutMvt('E_DATEECHEANCE', DateEche);
    OBM.PutMvt('E_DATEVALEUR', DateValeur);
  end;
end;

procedure TFSaisietr.OuvreEche(Lig: integer; Cpte: String17; MR: String3; RempliAuto: boolean; t: TLettre; LigTresoEnPied: Boolean);
var
  OMODR   : TMOD;
  OkInit  : Boolean;
  X       : T_MONOECH;
  O       : TOBM;
  CGen    : TGGeneral ;
begin
  OMODR := TMOD(GS.Objects[SA_NumP, Lig]);
  if OMODR = nil then Exit;
  CGen:=GetGGeneral(GS,Lig) ;
  if CGen=Nil then Exit ;
  O := GetO(GS, Lig);
  with OMODR.MODR do
  begin
    if Self.Action = taConsult then Action := taConsult else
    begin
      if not ((TotalAPayerP = 0) or (NbEche = 0)) then Action := taModif else Action := taCreat;
    end;
    OkInit := (Action = taCreat);
    DebutOuvreEche(Lig, Cpte, MR, OMODR);
    if Action = taCreat then
    begin
      if DateEcheDefaut = 0 then TabEche[1].DateEche := DateMvt(Lig) else TabEche[1].DateEche := DateEcheDefaut;
      TabEche[1].DateValeur := CalculDateValeur(DateMvt(Lig), t, SAJAL.Treso.CPT);
    end;
  end;
  if RempliAuto then
  begin
    //if M.Treso then
      //begin
      with OMODR.MODR do
        begin
        NbEche := 1;
        TabEche[1].Pourc := 100.0;
        TabEche[1].MontantP := TotalAPayerP;
        TabEche[1].MontantD := TotalAPayerD;
        end;
      //end
   //else CalculModeRegle(OMODR.MODR, True);
  end
  else
   begin
    case t of
      MultiEche: if M.Treso then
        begin
          with OMODR.MODR do
          begin
            X.DateEche := TabEche[1].DateEche;
            X.ModePaie := TabEche[1].ModePaie;
            X.DateValeur := TabEche[1].DateValeur;
            X.DateMvt := DateMvt(Lig);
            X.Cat := '';
            X.Treso := True;
            X.OkInit := OkInit;
            X.OkVal := True;
            X.NumTraiteCHQ := '';
            {$IFNDEF SPEC350}
            if O <> nil then X.NumTraiteCHQ := O.GetMvt('E_NUMTRAITECHQ');
            {$ENDIF}
            if SaisirMonoEcheance(X) then
            begin
              TabEche[1].DateEche := X.DateEche;
              TabEche[1].ModePaie := X.ModePaie;
              TabEche[1].DateValeur := X.DateValeur;
              NbEche := 1;
              TabEche[1].Pourc := 100.0;
              TabEche[1].MontantP := TotalAPayerP;
              TabEche[1].MontantD := TotalAPayerD;
              DateEcheDefaut := X.DateEche;
              {$IFNDEF SPEC350}
              if O <> nil then O.PutMvt('E_NUMTRAITECHQ', X.NumTraiteCHQ);
              {$ENDIF}
            end;
          end;
        end else SaisirEcheance(OMODR.MODR);
      MonoEche: with OMODR.MODR do
        begin
          X.DateEche := TabEche[1].DateEche;
          X.ModePaie := TabEche[1].ModePaie;
          X.DateValeur := TabEche[1].DateValeur;
          X.Action := Action;
          X.Cat := '';
          X.Treso := True;
          X.OkInit := OkInit;
          X.DateMvt := DateMvt(Lig);
          X.NumTraiteCHQ := '';
          {$IFNDEF SPEC350}
          if O <> nil then X.NumTraiteCHQ := O.GetMvt('E_NUMTRAITECHQ');
          {$ENDIF}
          if LigTresoEnPied or SaisirMonoEcheance(X) then
          begin
            TabEche[1].DateEche := X.DateEche;
            TabEche[1].ModePaie := X.ModePaie;
            TabEche[1].DateValeur := X.DateValeur;
            NbEche := 1;
            TabEche[1].Pourc := 100.0;
            TabEche[1].MontantP := TotalAPayerP;
            TabEche[1].MontantD := TotalAPayerD;
            DateEcheDefaut := X.DateEche;
            {$IFNDEF SPEC350}
            if O <> nil then O.PutMvt('E_NUMTRAITECHQ', X.NumTraiteCHQ);
            {$ENDIF}
          end;
        end;
    end;
    Revision := False;
  end;
  if OkInit then FinOuvreEche(Lig, OMODR, t);
//  if M.Treso then
    AlimEcheVersObm(Lig, OMODR.MODR.TabEche[1].ModePaie, OMODR.MODR.TabEche[1].DateEche,
      OMODR.MODR.TabEche[1].DateValeur);
end;

function TFSaisietr.AOuvrirEche(Lig: integer; var Cpte: String17; var MR: String3;
  var OuvreAuto, RempliAuto: boolean; var t: TLettre): Boolean;
var
  CGen: TGGeneral;
  CAux: TGTiers;
  OMODR: TMOD;
  XD, XP: Double;
begin
  Result := False;
  Cpte := '';
  MR := '';
  t := NonEche;
  CGen := GetGGeneral(GS, Lig);
  if CGen = nil then exit;
  CAux := GetGTiers(GS, Lig);
  if CGen.Collectif then
  begin
    if CAux = nil then Exit;
    if not CAux.Lettrable then Exit;
    Cpte := CAux.Auxi;
    MR := CAux.ModeRegle;
    t := MultiEche;
  end else
  begin
    t := Lettrable(CGen);
    if t = NonEche then Exit;
    Cpte := CGen.General;
    MR := CGen.ModeRegle;
    // Ajout test pour comptes divers lettrables... FQ 16136 SBO 09/08/2005
    if CGen.NatureGene = 'DIV' then
      begin
      OuvreAuto  := False ;
      RempliAuto := True ;
      MR         := GetParamSoc('SO_GCMODEREGLEDEFAUT', True) ;
      end ;
  end;
  if OuvreAuto then
  begin
    OMODR := TMOD(GS.Objects[SA_NumP, Lig]);
    if OMODR = nil then Exit;
    XD := MontantLigne(GS, Lig, tsmDevise);
    XP := MontantLigne(GS, Lig, tsmPivot);
    if (OMODR.MODR.TotalAPayerD = XD) and (OMODR.ModR.TotalAPayerP = XP) then Exit;
    if OMODR.MODR.Aux = Cpte then
    begin
      RecalculProrataEche(OMODR.MODR, XP, XD);
      Exit;
    end;
  end;
  Result := True;
end;

{==========================================================================================}
{=================================== Contrôles ============================================}
{==========================================================================================}
procedure TFSaisietr.ControleLignes;
var
  Lig: integer;
begin
  for Lig := 1 to GS.RowCount - 1 do if not EstRempli(GS, Lig) then DefautLigne(Lig, True);
end;

procedure TFSaisietr.AjusteLigne(Lig: integer);
var
  Cpte: String17;
  MR: String3;
  t: TLettre;
  Ouv, Remp: boolean;
begin
  AlimObjetMvt(Lig, True, 0);
  Remp := False;
  Ouv := True;
  if AOuvrirEche(Lig, Cpte, MR, Ouv, Remp, t) then ;
  if AOuvrirAnal(Lig, True) then ;
end;

procedure TFSaisietr.ErreurSaisie(Err: integer; szApres: string = '');
begin
  if Err < 100 then
  begin
    HLigne.Execute(Err - 1, caption, szApres);
  end else if Err < 200 then
  begin
    HPiece.Execute(Err - 101, caption, szApres);
  end;
end;

function TFSaisietr.PossibleRecupNum: Boolean;
var
  MM: String17;
  Facturier: String3;
  DD: TDateTime;
begin
  Result := True;
  if H_NUMEROPIECE_.Visible then Exit;
  DD := StrToDate(E_DATECOMPTABLE.Text);
  //if M.Treso then
    Facturier := SAJAL.CompteurNormal;
  {else begin
    if M.Simul <> 'N' then Facturier := SAJAL.CompteurSimul else Facturier := SAJAL.CompteurNormal;
  end;}
  if GetNum(EcrGen, Facturier, MM, DD) = NumPieceInt + 1 then Exit;
  Result := False;
end;

function TFSaisietr.PasModif: Boolean;
begin
  PasModif := ((Action = taConsult) or (ModeForce <> tmNormal));
end;

function TFSaisietr.LigneCorrecte(Lig: integer; Totale, Alim, CtrlCptTreso: boolean): Boolean;
var
  CGen: TGGeneral;
  CAux: TGTiers;
  Err, NumA, NumL: integer;
  OMOD: TMOD;
  {##ANA##
  OBA: TObjAna;
  P: P_TV;}
  OBA  : TOBM ;
  OBM  : TOBM ;
  premier_axe : Integer;
  {##ANA##}
  TotD, TotP : Double;
  Sens, ii: byte;
  Col: integer;
  IsCptTreso: Boolean;
begin
  if PasModif then
  begin
    Result := True;
    Exit;
  end;
  Err := 0;
  Sens := 1;
  {Contrôle du compte général}
  if ValC(GS, Lig) <> 0 then Sens := 2;
  CGen := GetGGeneral(GS, Lig);
  Col := -1;
  if ((Err = 0) and (CGen = nil)) then
  begin
    Err := 1;
    Col := SA_Gen;
  end;
  if ((Err = 0) and (CGen.General <> GS.Cells[SA_Gen, Lig])) then
  begin
    Err := 1;
    Col := SA_Gen;
  end;

  {Contrôle du compte auxiliaire}
  CAux := GetGTiers(GS, Lig);
  if ((Err = 0) and (not CGen.Collectif) and (CAux <> nil)) then
  begin
    Err := 2;
    Col := SA_Gen;
  end;
  if ((Err = 0) and (CAux <> nil)) then if CAux.Auxi <> GS.Cells[SA_Aux, Lig] then
  begin
    if GS.Cells[SA_Aux, Lig] <> '' then Err := 2 else Err := 3;
    Col := SA_Aux;
  end;
  if ((Err = 0) and (CGen.Collectif) and (CAux = nil)) then
  begin
    Err := 3;
    Col := SA_Aux;
  end;

  {Contrôles divers sur la pièce}
  if Err = 0 then
  begin
    ii := EstInterdit(SAJAL.CompteInterdit, CGen.General, Sens);
    if {M.Treso and} CtrlCptTreso then
    begin
      IsCptTreso := (Cgen <> nil) and (CGen.General = SAJAL.TRESO.Cpt);
      case SAJAL.TRESO.TYPCTR of
        Ligne: if (Lig mod 2 <> 0) and IsCptTreso then if M.Effet then Err := 30 else Err := 22;
        PiedSolde, PiedDC: if isCptTreso then Err := 22;
      end;
    end;
    case ii of 1: Err := 18;
      2: Err := 19;
      3: Err := 4;
    end;
    if Err > 0 then Col := SA_Gen;
  end;

  if ((Err = 0) and (not M.ANouveau) and (EstOuvreBil(CGen.General))) then Err := 20;
  if ((Err = 0) and (M.ANouveau) and (EstChaPro(CGen.NatureGene))) then Err := 21;
  if ((Err = 0) and (ValD(GS, Lig) = 0) and (ValC(GS, Lig) = 0)) then
  begin
    Err := 5;
    if QuelSens(E_NATUREPIECE.Value, CGen.NatureGene, CGen.Sens) = 1 then Col := SA_Debit else Col := SA_Credit;
  end;


  {Contrôle de la devise}
  if ((Err = 0) and (ModeDevise) and (ValeurPivot(GS.Cells[SA_Debit, Lig], DEV.Taux, Dev.Quotite) = 0) and
    (ValeurPivot(GS.Cells[SA_Credit, Lig], DEV.Taux, Dev.Quotite) = 0)) then
  begin
    Err := 6;
    if Sens = 1 then Col := SA_Debit else Col := SA_Credit;
  end;

  {Contrôle du Montant}
  if ((Err = 0) and (MontantLigne(GS, Lig, tsmDevise) <> 0) and (MontantLigne(GS, Lig, tsmPivot) = 0)) then
  begin
    Err := 6;
    if Sens = 1 then Col := SA_Debit else Col := SA_Credit;
  end;

  {Contrôle de la nature des comptes}
  if ((Err = 0) and (CGen.Collectif) and (not NatureGenAuxOk(CGen.NatureGene, CAux.NatureAux))) then
  begin
    Err := 7;
    Col := SA_Aux;
  end;

  {Contrôle de la nature de pièce}
  if ((Err = 0) and (not NaturePieceCompteOk(E_NATUREPIECE.Value, CGen.NatureGene))) then
  begin
    Err := 8;
    Col := SA_Gen;
  end;

  {Contrôle des montants négatifs}
  if ((Err = 0) and (not VH^.MontantNegatif)) then
  begin
    if ((ValD(GS, Lig) < 0) or (ValC(GS, Lig) < 0)) then
    begin
      Err := 9;
      Col := SA_Debit;
      if Sens = 2 then Col := SA_Credit;
    end;
  end;
  if Err = 0 then
  begin
    if ((Abs(MontantLigne(GS, Lig, tsmPivot)) < VH^.GrpMontantMin) or (Abs(MontantLigne(GS, Lig, tsmPivot)) > VH^.GrpMontantMax))
      then
    begin
      Err := 23;
      Col := SA_Debit;
      if Sens = 2 then Col := SA_Credit;
    end;
  end;

  {Conrôle de la totalisté de la ligne, si demandée (en fin de saisie)}
  if ((Err = 0) and (Totale)) then
  begin
    if TMOD(GS.Objects[SA_NumP, Lig]) <> nil then
    begin
      OMOD := TMOD(GS.Objects[SA_NumP, Lig]);
      if TotDifferent(MontantLigne(GS, Lig, tsmDevise), OMOD.MODR.TotalAPayerD) then
      begin
        Err := 10;
        Col := SA_Debit;
        if Sens = 2 then Col := SA_Credit;
      end;
    end;

    {Gestion de l'analytique}
    {if ((Err = 0) and (TObjAna(GS.Objects[SA_DateC, Lig]) <> nil)) then
    begin
      OBA := TObjAna(GS.Objects[SA_DateC, Lig]);
      if TotDifferent(MontantLigne(GS, Lig, tsmDevise), OBA.TotalDevise) then
      begin
        Err := 11;
        Col := SA_Debit;
        if Sens = 2 then Col := SA_Credit;
      end;

      if Err = 0 then for NumA := 1 to MaxAxe do
        if Ventilable(CGen, NumA) then begin
            if OBA.AA[NumA] = nil then Err := 11 + NumA else if OBA.AA[NumA].L.Count <= 0 then Err := 11 + NumA else
            begin
              TotD := 0;
              TotP := 0; //TotE:=0 ;
              for NumL := 0 to OBA.AA[NumA].L.Count - 1 do
              begin
                //LG*
                P := P_TV(OBA.AA[NumA].L.Items[NumL]);
                if P = nil then continue;
                TotD := TotD + GetMvtA(P.F, 'Y_DEBITDEV') + GetMvtA(P.F, 'Y_CREDITDEV');
                TotP := TotP + GetMvtA(P.F, 'Y_DEBIT') + GetMvtA(P.F, 'Y_CREDIT');
              end;
              if TotDifferent(TotD, MontantLigne(GS, Lig, tsmDevise)) then
              begin
                Err := 11 + NumA;
                Break;
              end;
              if TotDifferent(TotP, MontantLigne(GS, Lig, tsmPivot)) then
              begin
                Err := 11 + NumA;
                Break;
              end;
            end;}

    if ((Err = 0) and (GetO(GS,Lig).Detail.Count > 0)) then begin
      OBM := GetO(GS,Lig) ;
      {Mode Croisaxe  - SG6}
      if VH^.AnaCroisaxe then begin
        if Ventilable(CGen,0) then begin  {FQ 16433 : SBO 04/08/2005}
          premier_axe := RecherchePremDerAxeVentil.premier_axe;
          if premier_axe <> 0 then
            if OBM.Detail[premier_axe - 1].Detail.Count = 0 then Err := 11 + premier_axe;
        end
      end

      {Mode classique de ventilation}
      else begin

        for NumA := 1 to MaxAxe do begin
          if Ventilable(CGen, NumA) then begin
            if OBM.Detail[NumA - 1].Detail.Count=0 then
              Err := 11 + NumA
            else begin
              TotD := 0;
              TotP := 0;
              for NumL := 0 to OBM.Detail[NumA - 1].Detail.Count - 1 do begin
                OBA  := TOBM(OBM.Detail[NumA-1].Detail[NumL]);
                TotD := TotD + GetMontantDev(OBA);
                TotP := TotP + GetMontant(OBA);
              end;

              if Arrondi(TotD - MontantLigne(GS, Lig, tsmDevise), DEV.Decimale) <> 0 then begin
                Err := 11 + NumA;
                Break;
              end;
              if Arrondi(TotP - MontantLigne(GS, Lig, tsmPivot ), V_PGI.OkDecV) <> 0 then begin
                Err := 11 + NumA;
                Break;
              end;
            end;

            if Err <> 0 then begin
              Col := SA_Debit;
              if Sens = 2 then Col := SA_Credit;
            end;

          end; {if Ventilable(CGen, NumA) then}
        end;{for NumA := 1 to}
      end;{if VH^.AnaCroisaxe else begin}
    end;{if ((Err = 0) and (GetO(GS,Lig).Detail.Count > 0))}
  end; {if ((Err = 0) and (Totale)) }

  // controles de ventilation Err=12..16
  Result := (Err = 0);
  if not Result then
  begin
    ErreurSaisie(Err);
    if Col > 0 then GS.Col := Col;
  end else
  begin
    if Alim then
    begin
      AlimObjetMvt(Lig, TRUE, 0);
    end;
  end;
end;

function TFSaisietr.Equilibre: Boolean;
begin
  Result := ArrS(SI_TotDS - SI_TotCS) = 0;
end;

function TFSaisietr.PieceCorrecte(OkMess: Boolean): Boolean;
var
  i, Err: integer;
  OkBanque: Boolean;
  CGen: TGGeneral;
begin
  if PasModif then
  begin
    Result := True;
    exit;
  end;
  Err := 0;
  if ((EstRempli(GS, GS.RowCount - 1)) or (GS.Objects[SA_Gen, GS.RowCount - 1] <> nil) or (GS.Objects[SA_Aux, GS.RowCount - 1] <> nil)) then Err := 22;
  if not Equilibre then
  begin
    if ModeDevise then Err := 101 else Err := 103;
  end;
  if ((Err = 0) and (GS.RowCount < 3)) then Err := 102;
  OkBanque := FALSE;
  if Err = 0 then for i := 1 to GS.RowCount - 2 do
    begin
      if not LigneCorrecte(i, True, False, FALSE) then
      begin
        Err := 1000;
        Break;
      end;
      CGen := GetGGeneral(GS, i);
      if (CGen <> nil) and (SAJAL <> nil) and (CGen.General = SAJAL.TRESO.CPT) then OkBanque := TRUE;
    end;
  Result := (Err = 0);
  if not Result then
  begin
    if ((ModeDevise) and (Err = 103)) then ForcerMode(True, 0) else ErreurSaisie(Err);
  end else
  begin
    if (not M.Effet) and (SAJAL <> nil) and (SAJAL.TRESO.TypCtr in [AuChoix]) and (not OkBanque) then
      if OkMess then
      begin
        if HLigne.Execute(31, Caption, '') <> mrYes then Result := FALSE;
      end;
  end;
end;

{==========================================================================================}
{========================= Traitements de calcul liés aux lignes ==========================}
{==========================================================================================}
procedure TFSaisietr.DefautLigne(Lig: integer; Init: boolean);
var
  St: string;
  DD: TDateTime;
  O: TOBM;
  i: integer;
  XX: double;
begin
  GS.Cells[SA_NumP, Lig] := IntToStr(NumPieceInt);
  if Lig > 1 then
  begin
    GS.Cells[SA_DateC, Lig] := GS.Cells[SA_DateC, Lig - 1];
    {if not M.Treso then
    begin
      GS.Cells[SA_RefI, Lig] := GS.Cells[SA_RefI, Lig - 1];
      GS.Cells[SA_Lib, Lig] := GS.Cells[SA_Lib, Lig - 1];
    end;}
  end else
  begin
    GS.Cells[SA_DateC, Lig] := DateToJour(StrToDate(E_DATECOMPTABLE.Text));
  end;
  GS.Cells[SA_NumL, Lig] := IntToStr(Lig);
  GS.Cells[SA_Exo, Lig] := QuelExoDT(StrToDate(E_DATECOMPTABLE.EditText));
  AlloueMvt(GS, EcrGen, Lig, Init);
  O := GetO(GS, Lig);
  if O = nil then Exit;
  if ((OkScenario) and (Entete <> nil) and (Init)) then
  begin
    St := Entete.GetMvt('E_REFINTERNE');
    if St <> '' then
    begin
      O.PutMvt('E_REFINTERNE', St);
      GS.Cells[SA_RefI, Lig] := St;
    end;
    St := Entete.GetMvt('E_LIBELLE');
    if St <> '' then
    begin
      O.PutMvt('E_LIBELLE', St);
      GS.Cells[SA_Lib, Lig] := St;
    end;
    St := Entete.GetMvt('E_REFEXTERNE');
    if St <> '' then O.PutMvt('E_REFEXTERNE', St);
    DD := Entete.GetMvt('E_DATEREFEXTERNE');
    if DD > Encodedate(1901, 01, 01) then O.PutMvt('E_DATEREFEXTERNE', DD);
    St := Entete.GetMvt('E_AFFAIRE');
    if St <> '' then O.PutMvt('E_AFFAIRE', St);
    St := Entete.GetMvt('E_REFLIBRE');
    if St <> '' then O.PutMvt('E_REFLIBRE', St);
    {Zones libres}
    for i := 0 to 9 do
    begin
      St := Entete.GetMvt('E_LIBRETEXTE' + IntToStr(i));
      if St <> '' then O.PutMvt('E_LIBRETEXTE' + IntToStr(i), St);
      if i <= 3 then
      begin
        XX := Entete.GetMvt('E_LIBREMONTANT' + IntToStr(i));
        if XX <> 0 then O.PutMvt('E_LIBREMONTANT' + IntToStr(i), XX);
        St := Entete.GetMvt('E_TABLE' + IntToStr(i));
        if St <> '' then O.PutMvt('E_TABLE' + IntToStr(i), St);
        if i <= 1 then
        begin
          St := Entete.GetMvt('E_LIBREBOOL' + IntToStr(i));
          if St <> '' then O.PutMvt('E_LIBREBOOL' + IntToStr(i), St);
        end;
      end;
    end;
    DD := Entete.GetMvt('E_LIBREDATE');
    if DD > Encodedate(1901, 01, 01) then O.PutMvt('E_LIBREDATE', DD);
  end;
end;

procedure TFSaisietr.DetruitLigne(Lig: integer; Totale: boolean);
var
  R: integer;
begin
  GS.CacheEdit;
  if Totale then R := GS.Row else R := 1;
  if SAJAL.TRESO.TypCtr = Ligne then
  begin
    if EstRempli( GS, Lig + 1 ) then
      GS.DeleteRow(Lig + 1);
    GS.DeleteRow(Lig);
  end
  else
    GS.DeleteRow(Lig);

  CalculDebitCredit;
  NumeroteLignes(GS);
  NumeroteVentils;
  GS.SetFocus;
  if R >= GS.RowCount then R := GS.RowCount - 1;
  if R > 1 then GS.Row := R else GS.Row := 1;
  GS.Col := SA_Gen;
  GereNewLigne1;
  GS.Invalidate;
  GS.MontreEdit;
end;

procedure TFSaisietr.TraiteMontant(Lig: integer; Calc: boolean);
var
  XC, XD: Double;
  OBM: TOBM;
begin
  OBM := GetO(GS, Lig);
  if OBM = nil then Exit;
  XD := ValD(GS, Lig);
  XC := ValC(GS, Lig);
  OBM.SetMontants(XD, XC, DEV, False);
  if Calc then CalculDebitCredit;
end;

procedure TFSaisietr.PlusMvt(i: Integer; var SI_TDS, SI_TCS, SI_TDP, SI_TCP, SI_TDD, SI_TCD: Double);
var
  DD, CD: Double;
  OBM: TOBM;
begin
  DD := ValD(GS, i);
  CD := ValC(GS, i);
  SI_TDS := SI_TDS + DD;
  SI_TCS := SI_TCS + CD;
  OBM := GetO(GS, i);
  if OBM <> nil then
  begin
    SI_TDP := SI_TDP + OBM.GetMvt('E_DEBIT');
    SI_TCP := SI_TCP + OBM.GetMvt('E_CREDIT');
    SI_TDD := SI_TDD + OBM.GetMvt('E_DEBITDEV');
    SI_TCD := SI_TCD + OBM.GetMvt('E_CREDITDEV');
  end;
end;

procedure TFSaisietr.AfficheConf;
begin
  HConf.Visible := (ModeConf > '0') or ((PieceConf) and (Action = taConsult));
end;

procedure TFSaisietr.CalculDebitCredit;
var
  i: integer;
  CptGen: string;
  CGen: TGGeneral;
  CAux: TGTiers;
  {##ANA## OBA: TObjAna;}
begin
  ModeConf := '0';
  //if M.Treso then
  //begin
    SI_TotDS := 0;
    SI_TotCS := 0;
    SI_TotDP := 0;
    SI_TotCP := 0;
    SI_TotDD := 0;
    SI_TotCD := 0;
    SC_TotDS := 0;
    SC_TotCS := 0;
    SC_TotDP := 0;
    SC_TotCP := 0;
    SC_TotDD := 0;
    SC_TotCD := 0;
    for i := 1 to GS.RowCount - 1 do
    begin
      CptGen := Trim(GS.Cells[SA_Gen, i]);
      if CptGen <> '' then
      begin
        PlusMvt(i, SI_TotDS, SI_TotCS, SI_TotDP, SI_TotCP, SI_TotDD, SI_TotCD);
        if (SAJAL.TRESO.TypCtr in [Ligne, AuChoix]) and (CptGen = SAJAL.TRESO.CPT) then PlusMvt(i, SC_TotDS, SC_TotCS, SC_TotDP, SC_TotCP, SC_TotDD, SC_TotCD);
        CGen := GetGGeneral(GS, i);
        if ((CGen <> nil) and (CGen.Confidentiel > ModeConf)) then ModeConf := CGen.Confidentiel;
        CAux := GetGTiers(GS, i);
        if CAux <> nil then if CAux.Confidentiel > ModeConf then ModeConf := CAux.Confidentiel;
        {##ANA##
        OBA := TObjAna(GS.Objects[SA_DateC, i]);
        if OBA <> nil then if OBA.ModeConf > ModeConf then ModeConf := OBA.ModeConf;
        ##ANA##}
      end;
    end;
    // TR
    if SAJAL <> nil then
      case SAJAL.TRESO.TypCtr of
        Ligne, AuChoix:
          begin
            E_DEBITTRESO.VALUE := SC_TotDS;
            E_CREDITTRESO.Value := SC_TotCS;
          end;
        PiedSolde, PiedDC:
          begin
            SC_TotDS := SI_TotCS;
            SC_TotCS := SI_TotDS;
            SC_TotDP := SI_TotCP;
            SC_TotCP := SI_TotDP;
            SC_TotDD := SI_TotCD;
            SC_TotCD := SI_TotDD;
            if SAJAL.TRESO.TypCtr = PiedSolde then
            begin
              if SC_TotDS > SC_TotCS then
              begin
                E_DEBITTRESO.VALUE := SC_TotDS - SC_TotCS;
                E_CREDITTRESO.Value := 0;
              end else
              begin
                E_CREDITTRESO.VALUE := SC_TotCS - SC_TotDS;
                E_DEBITTRESO.Value := 0;
              end;
            end else
            begin
              E_DEBITTRESO.VALUE := SC_TotDS;
              E_CREDITTRESO.Value := SC_TotCS;
            end;
            OBMT[1].PutMvt('E_DEBIT', SC_TotDP);
            OBMT[2].PutMvt('E_CREDIT', SC_TotCP);
            OBMT[1].PutMvt('E_DEBITDEV', SC_TotDD);
            OBMT[2].PutMvt('E_CREDITDEV', SC_TotCD);
            SI_TotDS := SI_TotDS + SC_TotDS;
            SI_TotCS := SI_TotCS + SC_TotCS;
            SI_TotDP := SI_TotDP + SC_TotDP;
            SI_TotCP := SI_TotCP + SC_TotCP;
            SI_TotDD := SI_TotDD + SC_TotDD;
            SI_TotCD := SI_TotCD + SC_TotCD;
          end;
      end;
    SA_TotalDebit.Value := SI_TotDS;
    SA_TotalCredit.Value := SI_TotCS;
  (*
  end else
  begin
    (*
    SI_TotDebit:=0 ; SI_TotCredit:=0 ;
    SI_TotDebitPivot:=0 ; SI_TotCreditPivot:=0 ;
    SI_TotDebitEuro:=0 ; SI_TotCreditEuro:=0 ;
    SI_DPT:=0 ; SI_CPT:=0 ; SI_DDT:=0 ; SI_CDT:=0 ; SI_DET:=0 ; SI_CET:=0 ;
    for i:=1 to GS.RowCount-1 do if ((Trim(GS.Cells[SA_Gen,i])<>'')) then
        BEGIN
        SI_TotDebit:=SI_TotDebit+ValD(GS,i) ; SI_TotCredit:=SI_TotCredit+ValC(GS,i) ;
        OBM:=GetO(GS,i) ;
        if OBM<>Nil then
           BEGIN
           SI_TotDebitPivot:=SI_TotDebitPivot+Double(OBM.GetMvt('E_DEBIT')) ;
           SI_TotCreditPivot:=SI_TotCreditPivot+Double(OBM.GetMvt('E_CREDIT')) ;
           SI_TotDebitEuro:=SI_TotDebitEuro+Double(OBM.GetMvt('E_DEBITEURO')) ;
           SI_TotCreditEuro:=SI_TotCreditEuro+Double(OBM.GetMvt('E_CREDITEURO')) ;
           CGen:=GetGGeneral(GS,i) ; if ((CGen<>Nil) and (CGen.Confidentiel>ModeConf)) then ModeConf:=CGen.Confidentiel ;
           CAux:=GetGTiers(GS,i) ; if CAux<>Nil then if CAux.Confidentiel>ModeConf then ModeConf:=CAux.Confidentiel ;
           OBA:=TObjAna(GS.Objects[SA_DateC,i]) ; if OBA=Nil then if OBA.ModeConf>ModeConf then ModeConf:=OBA.ModeConf ;
           END ;
        END ;
    SA_TotalDebit.Value:=SI_TotDebit ; SA_TotalCredit.Value:=SI_TotCredit ;
    AfficheLeSolde(SA_Solde,SI_TotDebit,SI_TotCredit) ;

  end;
  *)
  AfficheConf;
end;

procedure TFSaisietr.SommeSoldeCompte(Col: integer; Cpte: string; var TD, TC: Double; Old, OkDev: Boolean);
var
  i: integer;
begin
  TD := 0;
  TC := 0;
  if Cpte = '' then Exit;
  for i := 1 to GS.RowCount - 2 do if GS.Cells[Col, i] = Cpte then
    begin
      if OkDev then
      begin
        TD := TD + ValD(GS, i);
        TC := TC + ValC(GS, i);
      end else
      begin
        TD := TD + ValeurPivot(GS.Cells[SA_Debit, i], DEV.Taux, DEV.Quotite) - Ord(Old) * GetDouble(GS, i, 'OLDDEBIT');
        TC := TC + ValeurPivot(GS.Cells[SA_Credit, i], DEV.Taux, DEV.Quotite) - Ord(Old) * GetDouble(GS, i, 'OLDCREDIT');
      end;
    end;
end;

procedure TFSaisietr.CalculSoldeCompte(CpteG, CpteA: string; DIG, CIG, DIA, CIA: Double);
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
end;

{==========================================================================================}
{================================ Appels des comptes en saisie=== =========================}
{==========================================================================================}
{JP 17/01/08 : ChargeOk : Pour savoir s'il faut charger le compte general}
function TFSaisietr.ChargeCompteEffetJAL(ChargeOk : Boolean) : Boolean;
begin
  Result := FALSE;
  If SAJAL=NIL Then Exit ;
  If E_CONTREPARTIEGEN.Text='' Then Exit ;
  If SAJAL.Journal='' Then Exit ;

  if ChargeOk and FInfo.LoadCompte( E_CONTREPARTIEGEN.Text ) then
  begin
    if FInfo.GetString('G_VISAREVISION') = 'X' then
    begin
      E_CONTREPARTIEGEN.Text := '' ;
      exit ;
    end ;
    E_CONTREPARTIEGEN.Text := finfo.StCompte ;
    SAJAL.TRESO.Cpt := E_CONTREPARTIEGEN.Text;
  end ;
  (*
  SAJAL.TRESO.STypCtr:=JQ.FindField('J_TYPECONTREPARTIE').AsString ;
  SAJAL.TRESO.TypCtr:=StrToTypCtr(TRESO.STypCtr) ;
  *)

  {JP 17/01/08 : FQ 18563 : Calcul du solde en fonction du compte et surtout de l'exercice}
  SAJAL.PutDate(E_DATECOMPTABLE.Text);
  SAJAL.ChargeCompteTreso;
  CalculSoldeCompteT;
  
  SAJAL.TRESO.DevBqe := E_DEVISE.Value;
  InitBoutonValide;
  Result := TRUE;
end;

procedure TFSaisietr.ChercheJAL(Jal: String3; Zoom: boolean);
begin
  SAJAL.Free; //SAJAL:=TSAJAL.Create(Jal,M.Treso) ;
  if M.Effet then
  begin
    SAJAL := TSAJAL.CreateEff(Jal);
    if SAJAL.Treso.STypCtr = '' then
    begin
      SAJAL.TRESO.STypCtr := 'LIG';
      SAJAL.TRESO.TypCtr := Ligne;
    end;
  end else SAJAL := TSAJAL.Create(Jal, M.Treso);

  if Agio then
  begin
    SAJAL.TRESO.STypCtr := 'MAN';
    SAJAL.TRESO.TypCtr := AuChoix;
  end;
  case Action of
    taCreat: InitEnteteJal(True, Zoom, False);
    //   taModif   : InitModifJal(Action) ;
    //   taConsult : InitModifJal(Action) ;
  end;
  PieceModifiee := True;
  if M.Effet then
    ChargeCompteEffetJal(False);{JP 17/01/08}
end;

procedure TFSaisietr.AppelAuto(indice: integer);
var
  Cpte: string;
begin
  if PasModif then Exit;
  if GS.Col <> SA_Gen then Exit;
  Cpte := TrouveAuto(SAJAL.COMPTEAUTOMAT, indice);
  if Cpte = '' then Exit;
  GS.Cells[SA_Gen, GS.Row] := Cpte;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 05/10/2004
Modifié le ... : 06/10/2004
Description .. : - BPY le 05/10/2004 - Gestion des lookup sur compte tiers
Suite ........ : et generaux par l'intermediaire d'un THEdit et des lookup
Suite ........ : standard AGL plutot que par l'intermediaire d'un hcompte
Mots clefs ... :
*****************************************************************}
function TFSaisietr.ChercheGen(C, L: integer; Force: boolean; DoPopup: boolean): byte;
var
  St: string;
  CGen, CGenAvant: TGGeneral;
  CAux: TGTiers;
  Idem, Changed, CollAvant, HavePopup, SurCptTreso, found: Boolean;
  q: tquery;
begin
  result := 0;

  // sauvegarde de la valeutr saisie
  St := uppercase(ConvertJoker(GS.Cells[C, L]));

  // compte collectif ??
  CGenAvant := GetGGeneral(GS, L);
  if (CGenAvant <> nil) then CollAvant := CGenAvant.Collectif
  else CollAvant := false;

  // BPY le 18/11/2004 => Fiche n° 14948 : intitulé du llokup ne doit pas changé ... modification du plus a la place de la tablette ...
      // set de la tablette
  compte.DataType := 'TZGENERAL';

  // set de la clause plus en fonction de ??? ...
  if (M.ANouveau) then compte.Plus := 'G_NATUREGENE<>"CHA" AND G_NATUREGENE<>"PRO"'
  else compte.Plus := '';
  // ... ou du compte auxiliaire
  CAux := GetGTiers(GS, L);
  if ((CAux <> nil) and (GS.Cells[SA_Aux, L] <> '')) then
    begin
    if ((CAux.NatureAux = 'AUD') or (CAux.NatureAux = 'CLI')) then compte.Plus := 'G_COLLECTIF="X" AND (G_NATUREGENE="COC" OR G_NATUREGENE="COD")'
    else if ((CAux.NatureAux = 'AUC') or (CAux.NatureAux = 'FOU')) then compte.Plus := 'G_COLLECTIF="X" AND (G_NATUREGENE="COF" OR G_NATUREGENE="COD")'
    else if (CAux.NatureAux = 'SAL') then compte.Plus := 'G_COLLECTIF="X" AND G_NATUREGENE="COS"'
    else compte.Plus := 'G_COLLECTIF="X"';
    end
    // En mode "effet", on se base sur la contrepartie
  else if M.Effet then
      begin
      if FInfo.LoadCompte( E_CONTREPARTIEGEN.Text ) then
        begin
        if FInfo.GetString('G_NATUREGENE') = 'COF' then
          compte.Plus := compte.Plus + ' AND G_NATUREGENE<>"COC" '
        else if FInfo.GetString('G_NATUREGENE') = 'COF' then
          compte.Plus := compte.Plus + ' AND G_NATUREGENE<>"COF" ' ;
        end ;
      end
    // sinon test sur la nature de la pièce
    else if E_NATUREPIECE.Value <> '' then
      begin
      Case CaseNatP( E_NATUREPIECE.Value ) of
        1,2,3 : if compte.Plus<>''
                  then compte.Plus := compte.Plus + ' AND G_NATUREGENE<>"COF" '
                  else compte.Plus := ' G_NATUREGENE<>"COF" ' ;
        4,5,6 : if compte.Plus<>''
                  then compte.Plus := compte.Plus + ' AND G_NATUREGENE<>"COC" '
                  else compte.Plus := ' G_NATUREGENE<>"COC" ' ;
        end ;
      end ;

  //==============================
//  if (M.Treso) then
    SurCptTreso := (GS.Cells[SA_GEN, L] = SAJAL.Treso.Cpt);
//  else
  //  SurCptTreso := false;

//  if (M.Treso) then
  //begin
    if ((SAJAL.TRESO.TypCtr <> AuChoix) and (not SurCptTreso)) then
    begin
      // Ajout des comptes lettrables divers... FQ 16136 SBO 09/08/2005
      if (E_NATUREPIECE.Value = 'RC')
        then compte.Plus := '((G_COLLECTIF="X" AND (G_NATUREGENE="COC" OR G_NATUREGENE="COD")) OR (G_NATUREGENE="TID") OR ((G_NATUREGENE="DIV") AND (G_LETTRABLE="X")))'
      else if (E_NATUREPIECE.Value = 'RF')
        then compte.Plus := '((G_COLLECTIF="X" AND (G_NATUREGENE="COF" OR G_NATUREGENE="COD")) OR (G_NATUREGENE="TIC") OR ((G_NATUREGENE="DIV") AND (G_LETTRABLE="X")))';
    end;

  compte.Plus := compte.Plus + ' AND G_FERME<>"X" ' ;

//  end;
  //==============================
// Fin BPY

  {##ANA##
  // FQ 16836 : 11/10/2005 Exit les comptes ventilables pour cause non compatibilité depuis passage CWAS
  if compte.Plus <> ''
    then compte.Plus := compte.Plus + ' AND G_VENTILABLE<>"X" '
    else compte.Plus := ' G_VENTILABLE<>"X" ' ;
  ##ANA##}

  // set des coordonnées
  compte.Top := GS.Top + GS.CellRect(C, L).Top;
  compte.Left := Gs.Left + GS.CellRect(C, L).Left;
  compte.Height := GS.CellRect(C, L).Bottom - GS.CellRect(C, L).top;
  compte.Width := GS.CellRect(C, L).Right - GS.CellRect(C, L).Left;

  // recherche d'un possible compte acev le caractere de bourage !
  if GetKeyState(VK_CONTROL) < 0    // FQ 16060 : SBO 03/08/2005
    then Compte.text := st // pas de bourage en recherche sur libelle ...
    else compte.text := BourreLaDonc(st, fbGene);
  // recherche du code dans le lookup
  found := LookupValueExist(compte);
  // recherche du libellé unique
  if (not found) then
  begin
    if (not DoPopup) then
    begin
      q := OpenSQL('SELECT G_GENERAL FROM GENERAUX WHERE UPPER(G_LIBELLE) LIKE UPPER("' + st + '%") AND ' + compte.Plus, true);
      if (not (q.EOF) and (Q.RecordCount = 1)) then
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
      if (TGGeneral(GS.Objects[C, L]).General <> St) then Changed := True
      else changed := false;
    end
    else changed := false;
    // a t'il changer ??
    Idem := ((St = compte.Text) and (GS.Objects[C, L] <> nil) and (not Changed));

    if ((not Idem) or (Force)) then
    begin
      // set du nouveau compte !
      GS.Cells[C, L] := compte.Text;
      // vire l'ancien compte dans la grille
      Desalloue(GS, C, L);
      // recup du compte et mise en objet de la grille !
      CGen := TGGeneral.Create(compte.Text);
      GS.Objects[C, L] := TObject(CGen);

      // si ct pas la meme on desaloue !
      if (not Idem) then
      begin
        if ((not CollAvant) or (not CGen.Collectif)) then DesalloueEche(L);
        DesalloueAnal(L {,CGEN.Ventilable});
      end;

      // compte interdit ??
      if EstInterdit(SAJAL.COMPTEINTERDIT, compte.Text, 0) > 0 then
      begin
        DesalloueEche(L);
        DesalloueAnal(L);
        Desalloue(GS, C, L);
        GS.Cells[C, L] := '';
        ErreurSaisie(4);
        Exit;
      end;

      // compte de bilan ???
      if ((not M.ANouveau) and (EstOuvreBil(compte.Text))) then
      begin
        DesalloueEche(L);
        DesalloueAnal(L);
        Desalloue(GS, C, L);
        GS.Cells[C, L] := '';
        ErreurSaisie(20);
        Exit;
      end;

      // compte de charge
      if ((M.ANouveau) and (EstChaPro(CGen.NatureGene))) then
      begin
        DesalloueEche(L);
        DesalloueAnal(L);
        Desalloue(GS, C, L);
        GS.Cells[C, L] := '';
        ErreurSaisie(21);
        Exit;
      end;

      if ((CGen.NatureGene = 'BQE') and (CGen.General <> SAJAL.TRESO.Cpt)) then if (PasDeviseBanque(CGen.General, E_DEVISE.Value)) then
        begin
          DesalloueEche(L);
          DesalloueAnal(L);
          Desalloue(GS, C, L);
          GS.Cells[C, L] := '';
          ErreurSaisie(25);
          Exit;
        end;

      // SBO 19/08/2004 FQ 13290 Saisie interdite sur les comptes fermés
      if CGen.Ferme then
      begin
        DesalloueEche(L);
        DesalloueAnal(L);
        Desalloue(GS, C, L);
        GS.Cells[C, L] := '';
        ErreurSaisie(31, '(' + SAJAL.TRESO.Cpt + ')');
        Exit;
      end;

	    if FInfo.LoadCompte( GS.Cells[SA_Gen,L] ) then
	     begin
	      if FInfo.GetString('G_VISAREVISION') = 'X' then
	      begin
	          DesalloueEche(L);
	          DesalloueAnal(L);
	          Desalloue(GS,C,L);
	          GS.Cells[C,L] := '';
	          ErreurSaisie(33);
	          Exit;
	      end;
	     end ;

      // recup l'analytique
      if (not Idem) then AlloueEcheAnal(C, L);

      Result := 1;

      // affecte les truc
      AffecteConso(C, L);
      AffecteTva(C, L);
      if ((CGen.NatureGene = 'TIC') or (CGen.NatureGene = 'TID')) then if (AffecteRIB(C, L, FALSE)) then BModifRIB.Enabled := true;
    end
    else
    begin
      CGen := TGGeneral(GS.Objects[C, L]);
      if (HavePopup) and ((not HavePopup) and (DoPopup)) then result := 1
      else result := 2;
    end;

    if (CGen <> nil) then
      if (not CGen.Collectif) then
      begin
        Desalloue(GS, SA_Aux, L);
        GS.Cells[SA_Aux, L] := '';
        if (Lettrable(CGen) = NonEche) then DesalloueEche(L);
        if (not Ventilable(CGen, 0)) then DesalloueAnal(L);
      end;

    // Chargement TInfoECriture // FQ 13246 : SBO 30/03/2005
    if GS.Cells[C,L] <> '' then
      FInfo.LoadCompte( GS.Cells[C,L] ) ;

    AffichePied;
  end;
end;

procedure TFSaisietr.AffecteConso(C, L: integer);
var
  O: TOBM;
  CAux: TGTiers;
  CGen: TGGeneral;
  Code: string;
begin
  O := GetO(GS, L);
  if O = nil then Exit;
  if C = SA_AUX then
  begin
    CAux := GetGTiers(GS, L);
    if CAux = nil then Exit;
    Code := CAux.CodeConso;
  end else
  begin
    CGen := GetGGeneral(GS, L);
    if CGen = nil then Exit;
    Code := CGen.CodeConso;
  end;
  if ((C = SA_Aux) or (Code <> '')) then O.PutMvt('E_CONSO', Code);
end;

function TFSaisietr.AffecteRIB(C, L: integer; IsAux: Boolean): Boolean;
var
  O: TOBM;
  Q: TQuery;
  Cpt, St, SRIB: string;
begin
  Result := FALSE;
  if Action = taConsult then Exit;
  if (IsAux and (C <> SA_Aux)) or ((not IsAux) and (C <> SA_Gen)) then Exit;
  Cpt := GS.Cells[C, L];
  if Cpt = '' then Exit;
  O := GetO(GS, L);
  if O = nil then Exit;
  sRIB := O.GetMvt('E_RIB');
  if OkScenario then
  begin
    if (GestionRIB = 'MAN') or (GestionRIB = 'PRI') then Result := TRUE;
    if GestionRIB <> 'PRI' then Exit;
  end else
  begin
    if not VH^.AttribRIBAuto then Exit;
    if sRIB <> '' then Exit;
  end;
  St:='' ;
  if (VH^.PaysLocalisation=CodeISOES) then
   st:=GetRIBPrincipal(cpt)
   else
   begin
     Q:=OpenSQL('Select * from RIB Where R_AUXILIAIRE="'+Cpt+'" AND R_PRINCIPAL="X"',True) ;
     if Not Q.EOF then
        BEGIN
        IF (codeisodupays(Q.FindField('R_PAYS').AsString) <> 'FR') then
          St:='*'+Q.FindField('R_CODEIBAN').AsString
        else
          St:=EncodeRIB(Q.FindField('R_ETABBQ').AsString,
                        Q.FindField('R_GUICHET').AsString,
                        Q.FindField('R_NUMEROCOMPTE').AsString,
                        Q.FindField('R_CLERIB').AsString,
                        Q.FindField('R_DOMICILIATION').AsString) ;
      END ;
   Ferme(Q) ;
   End ;
if (VH^.PaysLocalisation<>CodeISOES) or (trim(St)<>'') then
   Begin
   O.PutMvt('E_RIB',St) ;
   If IsAux Then O.PutMvt('E_AUXILIAIRE',Cpt) Else O.PutMvt('E_GENERAL',Cpt) ;
   Result:=TRUE ;
   END else
   BEGIN
   if sRIB<>'' then O.PutMvt('E_RIB','') ;
   END ;
//Ferme(Q) ; //XVI 24/02/2005
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 05/10/2004
Modifié le ... : 06/10/2004
Description .. : - BPY le 05/10/2004 - Gestion des lookup sur compte tiers
Suite ........ : et generaux par l'intermediaire d'un THEdit et des lookup
Suite ........ : standard AGL plutot que par l'intermediaire d'un hcompte
Mots clefs ... :
*****************************************************************}
function TFSaisieTr.ChercheAux(C, L: integer; Force: boolean; DoPopup: boolean): byte;
var
  st: string;
  CAux: TGTiers;
  CGen, GenColl: TGGeneral;
  Idem, Changed, HavePopup, SurCptTreso, found: boolean;
  Err: integer;
  q: tquery;
begin
  result := 0;

  // sauvegarde de la valeutr saisie
  St := uppercase(ConvertJoker(GS.Cells[C, L]));

  // BPY le 15/11/2004 => Fiche n° 14948 : intitulé du llokup ne doit pas changé ... modification du plus a la place de la tablette ...
      // set de la tablette
  compte.DataType := 'TZTTOUS';
  compte.Plus := '(T_NATUREAUXI<>"NCP" AND T_NATUREAUXI<>"CON" AND T_NATUREAUXI<>"PRO" AND T_NATUREAUXI<>"SUS")' ;  // FQ 15838 : SBO 03/08/2005

  // ... ou en fonction du compte general !!
  CGen := GetGGeneral(GS, L);
  if ((CGen <> nil) and (GS.Cells[SA_Gen, L] <> '')) then
    begin
    if (CGen.NatureGene = 'COC') then compte.Plus := '(T_NATUREAUXI="AUD" OR T_NATUREAUXI="CLI" OR T_NATUREAUXI="DIV")'
    else if (CGen.NatureGene = 'COF') then compte.Plus := '(T_NATUREAUXI="AUC" OR T_NATUREAUXI="FOU" OR T_NATUREAUXI="DIV" OR T_NATUREAUXI="SAL")'
    else if (CGen.NatureGene = 'COS') then compte.Plus := 'T_NATUREAUXI="SAL"';
    end
    // En mode "effet", on se base sur la contrepartie
  else if M.Effet then
      begin
      if FInfo.LoadCompte( E_CONTREPARTIEGEN.Text ) then
        begin
        if FInfo.GetString('G_NATUREGENE') = 'COF' then
          compte.Plus := '(T_NATUREAUXI="AUC" OR T_NATUREAUXI="FOU" OR T_NATUREAUXI="DIV" OR T_NATUREAUXI="SAL")'
        else if FInfo.GetString('G_NATUREGENE') = 'COF' then
          compte.Plus := '(T_NATUREAUXI="AUD" OR T_NATUREAUXI="CLI" OR T_NATUREAUXI="DIV")'
        else if FInfo.GetString('G_NATUREGENE') = 'COF' then
          compte.Plus := 'T_NATUREAUXI="SAL"';
        end ;
      end
    // sinon "mode règlement" test sur la nature de la pièce
    else if E_NATUREPIECE.Value <> '' then
      begin
      Case CaseNatP( E_NATUREPIECE.Value ) of
        1,2,3 : compte.Plus := compte.Plus + ' AND T_NATUREAUXI<>"FOU" ' ;
        4,5,6 : compte.Plus := compte.Plus + ' AND T_NATUREAUXI<>"CLI" ' ;
        end ;
      end ;

  //==============================
//  if (M.Treso) then
    SurCptTreso := GS.Cells[SA_GEN, L] = SAJAL.Treso.Cpt;
//  else
  //  SurCptTreso := false;

//  if (M.Treso) then
//  begin
    if ((SAJAL.TRESO.TypCtr <> AuChoix) and (not SurCptTreso)) then
    begin
      if (E_NATUREPIECE.Value = 'RC') then compte.Plus := '(T_NATUREAUXI="AUD" OR T_NATUREAUXI="CLI" OR T_NATUREAUXI="DIV")'
      else if (E_NATUREPIECE.Value = 'RF') then compte.Plus := '(T_NATUREAUXI="AUC" OR T_NATUREAUXI="FOU" OR T_NATUREAUXI="DIV" OR T_NATUREAUXI="SAL")';
    end;
//  end;

    compte.Plus := compte.Plus + ' AND T_FERME<>"X" ' ;

  //==============================
// Fin BPY

  // set des coordonnées
  compte.Top := GS.Top + GS.CellRect(C, L).Top;
  compte.Left := Gs.Left + GS.CellRect(C, L).Left;
  compte.Height := GS.CellRect(C, L).Bottom - GS.CellRect(C, L).top;
  compte.Width := GS.CellRect(C, L).Right - GS.CellRect(C, L).Left;

  // recherche d'un possible compte acev le caractere de bourage !
  if GetKeyState(VK_CONTROL) < 0    // FQ 16060 : SBO 03/08/2005
    then Compte.text := st // pas de bourage en recherche sur libelle ...
    else compte.text := BourreLaDonc(st, fbAux);
  // recherche du code dans le lookup
  found := LookupValueExist(compte);
  // recherche du libellé unique
  if (not found) then
  begin
    if (not ((DoPopup) or (GetParamSoc('SO_CPCODEAUXIONLY', false)))) then
    begin
      q := OpenSQL('SELECT T_AUXILIAIRE FROM TIERS WHERE UPPER(T_LIBELLE) LIKE UPPER("' + st + '%") AND ' + compte.Plus, true);
      if (not (q.EOF) and (Q.RecordCount = 1)) then
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
    if (GS.Objects[C, L] <> nil) then
    begin
      if (TGTiers(GS.Objects[C, L]).Auxi <> st) then Changed := true
      else changed := false;
    end
    else changed := false;
    Idem := ((st = GS.Cells[C, L]) and (GS.Objects[C, L] <> nil) and (not Changed));

    // si oui  !
    if ((not Idem) or (Force)) then
    begin
      // set du nouveau compte
      GS.Cells[C, L] := compte.Text;
      // vire l'ancien compte dans la grille
      Desalloue(GS, C, L);
      // recup du compte et mise en objet de la grille !
      CAux := TGTiers.Create(GS.Cells[C, L]);
      GS.Objects[C, L] := TObject(CAux);

      // si c pas le meme on desaloue
      if (not Idem) then
      begin
        DesalloueEche(L);
        DesalloueAnal(L);
      end;

      // Refuser tiers en devise si pas celle de saisie et pas sur devisepivot ni opposée
      if ((not CAux.MultiDevise) and (DEV.Code <> CAux.Devise) and (DEV.Code <> V_PGI.DevisePivot) and (CAux.Devise <> '')) then
      begin
        DesalloueEche(L);
        Desalloue(GS, C, L);
        GS.Cells[C, L] := '';
        ErreurSaisie(28);
        result := 0;
        Exit;
      end;

      // ???
      if ((not CAux.Lettrable) and (SAJAL.TRESO.TypCtr <> AuChoix)) then
      begin
        DesalloueEche(L);
        Desalloue(GS, C, L);
        GS.Cells[C, L] := '';
        ErreurSaisie(29);
        Exit;
      end;

      // SBO 19/08/2004 FQ 13290 Saisie interdite sur les comptes fermés
      if CAux.Ferme then
      begin
        DesalloueEche(L);
        Desalloue(GS, C, L);
        GS.Cells[C, L] := '';
        ErreurSaisie(31, '(' + SAJAL.TRESO.Cpt + ')');
        result := 0;
        Exit;
      end;

      // recup l'analytique
      if (not Idem) then AlloueEcheAnal(C, L);

      Result := 1;

      // affecte les truc
      if (AffecteRIB(C, L, true)) then BModifRIB.Enabled := true;
      AffecteConso(C, L);
      AffecteTva(C, L);
    end
    else // si c'est le meme !
    begin
      if (HavePopup) and ((not HavePopup) and (DoPopup)) then result := 1
      else result := 2;
    end;

    // si on a trouvé un auxiliaire mais que l'on a pas de general
    if (GS.Cells[SA_Gen, L] = '') then
    begin
      // recup du collectif de l'auxiliaire
      GS.Cells[SA_Gen, L] := GetGTiers(GS, L).Collectif;
      Desalloue(GS, SA_Gen, L);

      Err := 0;
      GenColl := TGGeneral.Create(GS.Cells[SA_Gen, L]);
      GS.Objects[SA_Gen, L] := GenColl;

      // general interdit pour ce journal ?
      if (EstInterdit(SAJAL.COMPTEINTERDIT, GS.Cells[SA_Gen, L], 0) > 0) then Err := 4;
      // confidentialité ??
      if (EstConfidentiel(GenColl.Confidentiel)) then Err := 24;
      // SBO 19/08/2004 FQ 13290 Saisie interdite sur les comptes fermés
      if (GenColl.Ferme) then Err := 24;
      // FQ 16836 : 11/10/2005 Exit les comptes ventilables pour cause non compatibilité depuis passage CWAS
     {##ANA## if Ventilable(GenColl, 0) then Err := 24;}

      // si on a une erreur on bloque ..
      if (Err > 0) then
      begin
        DesalloueEche(L);
        Desalloue(GS, C, L);
        Desalloue(GS, SA_Gen, L);
        GS.Cells[SA_Gen, L] := '';
        GS.Cells[SA_Aux, L] := '';
        ErreurSaisie(24);
      end
      else AlloueEcheAnal(SA_Gen, L);
    end;

    // Chargement TInfoECriture // FQ 13246 : SBO 30/03/2005
    if GS.Cells[C,L] <> '' then
      FInfo.LoadAux( GS.Cells[C,L] ) ;

    AffichePied;
  end;
end;

procedure TFSaisietr.ChercheMontant(Acol, Arow: longint);
begin
  if ACol = SA_Debit then
  begin
    if ValD(GS, ARow) <> 0 then GS.Cells[SA_Credit, ARow] := '' else GS.Cells[ACol, ARow] := '';
  end
  else
  begin
    if ValC(GS, ARow) <> 0 then GS.Cells[SA_Debit, ARow] := '' else GS.Cells[ACol, ARow] := '';
  end;
  FormatMontant(GS, ACol, ARow, DECDEV);
  TraiteMontant(ARow, True);
end;

{==========================================================================================}
{================================= Conversions, Caluls ====================================}
{==========================================================================================}
function TFSaisietr.DateMvt(Lig: integer): TDateTime;
begin
  Result := StrToDate(E_Datecomptable.Text);
end;

function TFSaisietr.ArrS(X: Double): Double;
begin
  Result := Arrondi(X, DECDEV);
end;

{==========================================================================================}
{=========================== Calculs sur ligne de trésorerie ==============================}
{==========================================================================================}

procedure TFSaisietr.E_NUMDEPExit(Sender: TObject);
var
  ST, St1, Masque: string;
  i: Integer;
begin
  St := Trim(E_NUMDEP.Text);
  St1 := Trim(E_REF.Text);
  if St = '' then
  begin
    SI_NUMREF := -1;
    Exit;
  end;
  Masque := '';
  for i := 1 to Length(St) do Masque := Masque + '0';
  E_NUMREF.Masks.PositiveMask := St1 + ' ' + Masque;
  SI_NUMREF := StrToInt(E_NUMDEP.Text);
  E_NUMREF.VALUE := SI_NUMREF;
end;

procedure TFSaisietr.IncRef;
begin
  if SI_NUMREF >= 0 then
  begin
    Inc(SI_NUMREF);
    E_NUMREF.Value := SI_NUMREF;
  end;
end;

//TR sur toutes ces procédures :
procedure TFSaisietr.CalculRefAuto(Sender: TObject);
begin
  if SAJAL = nil then Exit;

  PRefAuto.Left := GS.Left + (GS.Width - PRefAuto.Width) div 2;
  PRefAuto.Top := GS.Top + (GS.Height - PRefAuto.Height) div 2;
  PRefAuto.Visible := True;
  E_REF.SetFocus;
  PEntete.Enabled := False;
  Outils.Enabled := False;
  Valide97.Enabled := False;
  GS.Enabled := False;
  PPied.Enabled := False;
  ModeRA := True;
  (*
  SI_NumRef:=-1 ; E_NumDepExit(Sender) ;
  BeginTrans ;
  for i:=1 to GS.RowCount-2 do
     BEGIN
     OkOk:=TRUE ; St:='' ;
     if (SAJAL.Treso.TypCtr=Ligne) and (GS.Cells[SA_Gen,i]=SAJAL.TRESO.Cpt) then OkOk:=FALSE ;
     if OkOk and (i>1) then IncRef ;
     if SI_NumRef<0 then St:='' else St:=E_NUMREF.Text ; GS.Cells[SA_RefI,i]:=St ;
     OBM:=GetO(GS,i) ;
     if OBM<>NIL then
        BEGIN
        RefAvant:=OBM.GetMvt('E_REFINTERNE') ;
        OBM.PutMvt('E_REFINTERNE',St) ;
        If (RefAvant<>St) And M.MajDirecte Then MajDirecteTable(i,'') ;
        END ;
     { A faire : Idem Sur analytique }
     END ;
  CommitTrans ;
  *)
end;

procedure TFSaisietr.ClickLibelleAuto;
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
    if StL = '' then
    begin
      HPiece.Execute(15, caption, '');
      Exit;
    end;
    Q := OpenSQL('Select RA_FORMULEREF, RA_FORMULELIB from REFAUTO where RA_CODE="' + StL + '"', True);
    SI_FormuleRef := Q.FindField('RA_FORMULEREF').AsString;
    SI_FormuleLib := Q.FindField('RA_FORMULELIB').AsString;
    Ferme(Q);
  end;

  OldL := CurLig;

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

procedure TFSaisietr.ModifNaturePiece;
var
  OBM: TOBM;
  i: Integer;
  St, NaturePieceAvant: string;
begin
  BeginTrans;
  for i := 1 to GS.RowCount - 2 do
  begin
    St := E_NATUREPIECE.Value;
    OBM := GetO(GS, i);
    if OBM <> nil then
    begin
      NaturePieceAvant := OBM.GetMvt('E_NATUREPIECE');
      if (NaturePieceAvant <> St) then
      begin
        OBM.PutMvt('E_NATUREPIECE', St);
        if M.MajDirecte and StatusOBMOk(OBM) then MajDirecteTable(i, NaturePieceAvant);
      end;
    end;
    { A faire : Idem Sur analytique }
  end;
  CommitTrans;
end;

procedure TFSaisietr.ModifModePaie;
var
  OBM: TOBM;
  i: Integer;
  St, ModePaieAvant: string;
  M: TMOD;
begin
  BeginTrans;
  for i := 1 to GS.RowCount - 2 do
  begin
    OBM := GetO(GS, i);
    St := E_MODEPAIE.Value;
    if OBM <> nil then
    begin
      if (EstLettrable(i) in [MonoEche, MultiEche]) then
      begin
        M := TMOD(GS.Objects[SA_NumP, i]);
        if M <> nil then
        begin
          ModePaieAvant := M.MODR.TabEche[1].ModePaie;
          if (ModePaieAvant <> St) then
          begin
            M.MODR.TabEche[1].ModePaie := E_MODEPAIE.Value;
            if Self.M.MajDirecte and StatusOBMOk(OBM) then MajDirecteTable(i, '');
          end;
        end;
      end;
    end;
  end;
  CommitTrans;
end;

procedure TFSaisietr.AfficheLigneTreso;
var
  i, ll: Integer;
begin
  if SAJAL = nil then Exit;
  LigneTreso := not LigneTreso;
  {JP 21/10/05 : FQ 16790 : Hauteur à 1 au lieur de 0. Cela change légèrement l'ergonomie,
                 mais c'est la seule solution que j'ai trouvé au positionnement de la ScrollBar.
                 En fait si la hauteur est à 0, la position théorique dans grille n'est pas la
                 même que la position réèlle car 0 ne rend pas réellement la ligne invisible
                 puisque le traite de séparation est grossi}
  ll := 1;
  if LigneTreso then ll := GS.RowHeights[0];
  for i := 1 to GS.RowCount - 2 do
  begin
    if GS.Cells[SA_Gen, i] = SAJAL.TRESO.CPT then GS.RowHeights[i] := ll;
  end;
end;

procedure TFSaisietr.MajRefLibOBMT(Quoi: Integer);
var
  St1, St2: string;
begin
  if SAJAL = nil then Exit;
  if Quoi = 0 then
  begin
    St1 := 'E_REFINTERNE';
    St2 := E_REFTRESO.Text;
  end
  else
  begin
    St1 := 'E_LIBELLE';
    St2 := E_LIBTRESO.Text;
  end;
  if SAJAL.TRESO.TypCtr in [PiedSolde, PiedDC] then
  begin
    if OBMT[1] <> nil then OBMT[1].PutMvt(St1, St2);
    if OBMT[2] <> nil then OBMT[2].PutMvt(St1, St2);
  end;
end;

procedure TFSaisietr.E_REFTRESOExit(Sender: TObject);
begin
  MajRefLibOBMT(0);
end;

procedure TFSaisietr.E_LIBTRESOExit(Sender: TObject);
begin
  MajRefLibOBMT(1);
  GS.SetFocus;
end;

procedure TFSaisietr.E_CONTREPARTIEGENChange(Sender: TObject);
begin
  if SAJAL <> nil then E_CPTTRESO.Caption := SAJAL.Treso.Cpt;
end;

procedure TFSaisietr.E_CONTREPARTIEGENExit(Sender: TObject);
begin
  if M.Effet {and M.Treso} then ChargeCompteEffetJal(True); {JP 17/01/08}
  OkPourLigne;
  if SAJAL <> nil then
  begin
    E_CPTTRESO.Caption := SAJAL.Treso.Cpt;
    //if M.Treso then //TR
    //begin
      AlimObjetMvt(0, FALSE, 1);
      AlimObjetMvt(0, FALSE, 2);
    //end;
  end;
end;

procedure TFSaisietr.CalculSoldeCompteT;
var
  TDG, TCG : Double;
begin
  //if not M.Treso then Exit;
  if SAJAL <> nil then
  begin
    TDG := SAJAL.TRESO.TotD + SC_TotDP;
    TCG := SAJAL.TRESO.TotC + SC_TotCP;
    AfficheLeSolde(SA_SoldeTR, TDG, TCG);
  end;
end;

procedure TFSaisietr.RecopieGS(i: Integer);
var
  OBM1, OBM2: TOBM;
  CGen: TGGeneral;
  //    CAux : TGTiers ;
  OMODR1, OMODR2: TMOD;
  Cpte: String17;
  Let: TLettre;
  Okok: boolean;
begin
  GS.Cells[SA_GEN, i + 1] := SAJAL.TRESO.Cpt;
  GS.Cells[SA_AUX, i + 1] := '';
  GS.Cells[SA_REFI, i + 1] := GS.Cells[SA_REFI, i];
  GS.Cells[SA_LIB, i + 1] := GS.Cells[SA_LIB, i];
  GS.Cells[SA_DEBIT, i + 1] := GS.Cells[SA_CREDIT, i];
  GS.Cells[SA_CREDIT, i + 1] := GS.Cells[SA_DEBIT, i];
  // Suite tests avec comptes divers lettrables, le chargement des comptes doit être forcé FQ 16136 SBO 09/08/2005
  ChercheGen(SA_Gen, i + 1, True);
  CGen := GetGGeneral(GS, i + 1);
  if M.Effet and CGen.Collectif then
  begin
    GS.Cells[SA_AUX, i + 1] := GS.Cells[SA_AUX, i];
    if ChercheAux(SA_Aux, i + 1, True) = 0 then ; // idem ci-dessus FQ 16136 SBO 09/08/2005
  end;
  AlimObjetMvt(i + 1, TRUE, 0);
  OMODR1 := TMOD(GS.Objects[SA_NumP, i]);
  OMODR2 := TMOD(GS.Objects[SA_NumP, i + 1]);
  Let := EstLettrable(i + 1);
  OkOk := (Let = MonoEche) or (M.Effet and (Let = MultiEche));
  if OkOk then { Compte de banque pointable }
  begin
    DebutOuvreEche(i + 1, Cpte, '', OMODR2);
    FinOuvreEche(i + 1, OMODR2, Let);
    with OMODR2.MODR do
    begin
      NbEche := 1;
      TabEche[1].Pourc := 100.0;
      TabEche[1].MontantP := TotalAPayerP;
      TabEche[1].MontantD := TotalAPayerD;
      if OModR1 <> nil then
      begin
        TabEche[1].DateEche := OMODR1.MODR.TabEche[1].DateEche;
        TabEche[1].DateValeur := OMODR1.MODR.TabEche[1].DateValeur;
        TabEche[1].ModePaie := OMODR1.MODR.TabEche[1].ModePaie;
      end else
      begin
        TabEche[1].DateEche := DateMvt(i + 1);
        TabEche[1].ModePaie := E_MODEPAIE.Value;
        TabEche[1].DateValeur := CalculDateValeur(DateMvt(i + 1), MonoEche, SAJAL.Treso.CPT);
      end;
    end;
  end;
  OBM1 := GetO(GS, i);
  OBM2 := GetO(GS, i + 1);
  if (OBM1 = nil) or (OBM2 = nil) then Exit;
  //if M.Treso then
    if OModR2 <> nil then
      AlimEcheVersObm(i + 1, OMODR2.MODR.TabEche[1].ModePaie, OMODR2.MODR.TabEche[1].DateEche,
        OMODR2.MODR.TabEche[1].DateValeur);
  OBM2.PutMvt('E_REFEXTERNE', OBM1.GetMvt('E_REFEXTERNE'));
  OBM2.PutMvt('E_DATEREFEXTERNE', OBM1.GetMvt('E_DATEREFEXTERNE'));
  OBM2.PutMvt('E_AFFAIRE', OBM1.GetMvt('E_AFFAIRE'));
  OBM2.PutMvt('E_REFLIBRE', OBM1.GetMvt('E_REFLIBRE'));
  OBM2.PutMvt('E_QTE1', OBM1.GetMvt('E_QTE1'));
  OBM2.PutMvt('E_QTE2', OBM1.GetMvt('E_QTE2'));
  OBM2.PutMvt('E_QUALIFQTE1', OBM1.GetMvt('E_QUALIFQTE1'));
  OBM2.PutMvt('E_QUALIFQTE2', OBM1.GetMvt('E_QUALIFQTE2'));
  if M.Effet then
    OBM2.PutMvt('E_RIB', OBM1.GetMvt('E_RIB'));
end;

procedure TFSaisietr.AlimLigneTreso(Lig: Integer);
begin
  if Lig < 1 then Exit;
  if Lig = GS.RowCount - 2 then { Nouvelle Ligne }
  begin
    NewLigne(GS);
    DefautLigne(Lig + 1, True);
    RecopieGS(Lig);
    GS.Cells[SA_NUML, Lig + 2] := IntToStr(Lig + 2);
  end else
  begin
    RecopieGS(Lig);
  end;
  if not LigneTreso then
    {JP 21/10/05 : FQ 16790 : Hauteur à 1 au lieur de 0. Cela change légèrement l'ergonomie,
                   mais c'est la seule solution que j'ai trouvé au positionnement de la ScrollBar.
                   En fait si la hauteur est à 0, la position théorique dans grille n'est pas la
                   même que la position réèlle car 0 ne rend pas réellement la ligne invisible
                   puisque le traite de séparation est grossi}
    GS.RowHeights[Lig + 1] := 1;
end;

procedure TFSaisietr.GereLigneTreso(Lig: integer);
begin
  case SAJAL.TRESO.TypCtr of
    Ligne:
      begin
        AlimLigneTReso(Lig);
        CalculDebitCredit;
      end;
    PiedDC, PiedSolde:
      begin
      end;
    AuChoix: ;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 12/04/2002
Modifié le ... :   /  /
Description .. : - LG - 12/04/2002 - correction de l'affectation du compte de
Suite ........ : contrepartie en pied d'ecriture
Mots clefs ... :
*****************************************************************}
function TFSaisietr.LigneEnPlusTreso(Sens: Integer; MSaisi, MPivot, MDevise: Double): boolean;
var
  Lig: Integer;
  OBM, OBM1: TOBM;
  LDD, LCD, LDP, LCP, LDS, LCS: Double;
  CGen: TGGeneral;
begin
  Result := False;
  if EstRempli(GS, GS.RowCount - 1) then GS.RowCount := GS.RowCount + 1;
  Lig := GS.RowCount - 1;
  DefautLigne(Lig, TRUE);
  OBM1 := GetO(GS, Lig);
  if OBM1 <> nil then
  begin
    // LG* 12/03/2002 OBM1 à la plce de OBM
    OBM1.Dupliquer(OBMT[Sens], false, true);
    OBM1.M.Assign(OBMT[Sens].M);
  end;
  GS.Cells[SA_Gen, Lig] := OBMT[Sens].GetMvt('E_GENERAL');
  if ChercheGen(SA_GEN, Lig, TRUE) = 0 then exit;
  CGen := GetGGeneral(GS, Lig);
  GS.Cells[SA_Aux, Lig] := '';
  GS.Cells[SA_RefI, Lig] := OBMT[Sens].GetMvt('E_REFINTERNE');
  GS.Cells[SA_Lib, Lig] := OBMT[Sens].GetMvt('E_LIBELLE');
  LDD := 0;
  LCD := 0;
  LDP := 0;
  LCP := 0;
  LDS := 0;
  LCS := 0;
  if Sens = 1 then
  begin
    LDS := MSaisi;
    LDD := MDevise;
    LDP := MPivot
  end
  else
  begin
    LCS := MSaisi;
    LCD := MDevise;
    LCP := MPivot
  end;
  GS.Cells[SA_Debit, Lig] := StrS(LDS, DECDEV);
  GS.Cells[SA_Credit, Lig] := StrS(LCS, DECDEV);
  FormatMontant(GS, SA_Debit, Lig, DECDEV);
  FormatMontant(GS, SA_Credit, Lig, DECDEV);
  AlimObjetMvt(Lig, FALSE, 0);
  OBM := GetO(GS, Lig);
  OBM.CompS := True;
  OBM.PutMvt('E_DEBIT', LDP);
  OBM.PutMvt('E_CREDIT', LCP);
  OBM.PutMvt('E_DEBITDEV', LDD);
  OBM.PutMvt('E_CREDITDEV', LCD);
  if M.Effet then
  begin
    if EstLettrable(Lig) in [Monoeche, MultiEche] then OuvreEche(Lig, GS.Cells[SA_Gen, Lig], '', FALSE, MonoEche, TRUE);
  end else
  begin
    if Lettrable(CGen) = Monoeche then OuvreEche(Lig, GS.Cells[SA_Gen, Lig], '', FALSE, MonoEche, TRUE);
  end;
  GereNewLigneT;
  Result := True;
end;

function TFSaisietr.MajGridPiedTreso: boolean;
begin
  Result := False;
  if SAJAL = nil then Exit;
  case SAJAL.TRESO.TypCtr of
    PiedDC:
      begin
        if SC_TotDS <> 0 then Result := LigneEnPlusTreso(1, SC_TotDS, SC_TotDP, SC_TotDD);
        if SC_TotCS <> 0 then Result := LigneEnPlusTreso(2, SC_TotCS, SC_TotCP, SC_TotCD);
      end;
    PiedSolde:
      begin
        if SC_TotDS > SC_TotCS then Result := LigneEnPlusTreso(1, SC_TotDS - SC_TotCS, SC_TotDP - SC_TotCP, SC_TotDD - SC_TotCD);
        if SC_TotDS < SC_TotCS then Result := LigneEnPlusTreso(2, SC_TotCS - SC_TotDS, SC_TotCP - SC_TotDP, SC_TotCD - SC_TotDD);
      end;
  end;
end;
//TR Fin sur toutes ces procédures :

{==========================================================================================}
{========================= Affichages, Positionnements ====================================}
{==========================================================================================}
procedure TFSaisietr.AutoSuivant(Suiv: boolean);
var
  Cpte: String17;
begin
  if SAJAl = nil then Exit;
  if SAJAL.NbAuto <= 0 then Exit;
  if ((GS.Cells[SA_Gen, GS.Row] <> '') and (not CpteAuto)) then Exit;
  if GS.Cells[SA_Gen, GS.Row] = '' then CpteAuto := True;
  if Suiv then
  begin
    if CurAuto >= SAJAL.NbAuto then CurAuto := 1 else Inc(CurAuto);
  end
  else
  begin
    if CurAuto <= 1 then CurAuto := SAJAL.NbAuto else Dec(CurAuto);
  end;
  Cpte := TrouveAuto(SAJAL.COMPTEAUTOMAT, CurAuto);
  if Cpte = '' then Exit;
  GS.Cells[SA_Gen, GS.Row] := Cpte;
end;

procedure TFSaisietr.SoldelaLigne(Lig: integer);
var
  Diff, X: Double;
  OBM: TOBM;
  Col: integer;
  Okok: boolean;
begin
  OBM := GetO(GS, Lig);
  if OBM = nil then Exit;
  Diff := SI_TotDS - SI_TotCS;
  Okok := False;
  X := -1 * ValD(GS, Lig);
  if X = 0 then X := ValC(GS, Lig);
  Diff := Diff + X;
  GS.Cells[SA_Debit, Lig] := '';
  GS.Cells[SA_Credit, Lig] := '';
  if Diff > 0 then
  begin
    Col := SA_Credit;
    GS.Cells[Col, Lig] := StrS(Diff, DECDEV);
  end
  else
  begin
    Col := SA_Debit;
    GS.Cells[Col, Lig] := StrS(-Diff, DECDEV);
  end;
  ChercheMontant(Col, Lig);
  // Pivot
  Diff := SI_TotDP - SI_TotCP;
  if Diff <> 0 then
  begin
    X := -1 * OBM.GetMvt('E_DEBIT');
    if X = 0 then X := OBM.GetMvt('E_CREDIT');
    Diff := Diff + X;
    if Diff > 0 then OBM.PutMvt('E_CREDIT', Diff) else OBM.PutMvt('E_DEBIT', -Diff);
    Okok := True;
  end;
  // Devise
  Diff := SI_TotDD - SI_TotCD;
  if Diff <> 0 then
  begin
    X := -1 * OBM.GetMvt('E_DEBITDEV');
    if X = 0 then X := OBM.GetMvt('E_CREDITDEV');
    Diff := Diff + X;
    if Diff > 0 then OBM.PutMvt('E_CREDITDEV', Diff) else OBM.PutMvt('E_DEBITDEV', -Diff);
    Okok := True;
  end;
  if Okok then CalculDebitCredit;
  AjusteLigne(Lig);
end;

procedure TFSaisietr.PositionneDevise(ReInit: boolean);
var
  OldIndex: integer;
begin
  OldIndex := -2;
  if (SAJAL <> nil) and (SAJAL.MultiDevise) then
  begin
    RestoreDeviseEuro;
    E_DEVISE.Enabled := True;
  end else
  begin
    AddEuroFranc;
  end;
  if ((ReInit) and (OldIndex <> -2)) then
  begin
    E_DEVISE.ItemIndex := OldIndex;
    E_DEVISEChange(nil);
  end;
end;

procedure TFSaisietr.AffichePied;
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
    G_LIBELLE.Caption := FInfo.GetString('G_LIBELLE') ;
    CpteG             := FInfo.GetString('G_GENERAL') ;
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
  //if M.treso then
  //begin
    OMODR := TMOD(GS.Objects[SA_NumP, GS.Row]);
    if (OMODR <> nil) and (OMODR.MODR.NbECHE > 0) then
      begin
      HE_MODEP.Caption := RechDom('ttModePaie', OMODR.MODR.TabEche[1].ModePaie, FALSE);
      HE_DATEECHE.Caption := DateToStr(OMODR.MODR.TabEche[1].DateEche);
      end
    else
      begin
      HE_MODEP.Caption := '';
      HE_DATEECHE.Caption := '';
      end;
    CalculSoldeCompteT;
    OBM := GetO(GS, GS.Row);
    if OBM <> nil then
      begin
      if OBM.GetMvt('E_RIB') <> ''
        then E_RIB.Caption := HDiv.Mess[14] + ' ' + OBM.GetMvt('E_RIB')
        else E_RIB.Caption := '';
      end
    else E_RIB.Caption := '';
  //end;

end;

procedure TFSaisietr.GereNewLigneT;
var
  OBM: TOBM;
begin //TR
  if SAJAL.TRESO.TypCtr = Ligne then
  begin
    if EstRempli(GS, GS.RowCount - 1) then
    begin
      NewLigne(GS);
      IncRef;
    end else if not EstRempli(GS, GS.RowCount - 2) then
    begin
      OBM := GetO(GS, GS.RowCount - 2);
      if (OBM <> nil) and (OBM.Status = DansTable) then else GS.RowCount := GS.RowCount - 1;
    end;
  end else
  begin
    if EstRempli(GS, GS.RowCount - 1) then
    begin
      NewLigne(GS);
      IncRef;
    end else if not EstRempli(GS, GS.RowCount - 2) then
    begin
      OBM := GetO(GS, GS.RowCount - 2);
      if (OBM <> nil) and (OBM.Status = DansTable) then else GS.RowCount := GS.RowCount - 1;
    end;
  end;
end;

procedure TFSaisietr.GereNewLigne1;
begin //TR
//  if M.Treso then
    GereNewLigneT
//  else
  //  GereNewLigne(GS);
end;

procedure TFSaisietr.GereZoom;
var
  Okok: boolean;
begin
  Okok := False;
  if ((GS.Col = SA_Gen) and (GS.Cells[SA_Gen, GS.Row] <> '')) then Okok := True;
  if ((GS.Col = SA_Aux) and (GS.Cells[SA_Aux, GS.Row] <> '')) then Okok := True;
  BZoom.Enabled := Okok;
end;

{==========================================================================================}
{======================================== EURO ============================================}
{==========================================================================================}


procedure TFSaisietr.RestoreDeviseEuro;
var
  ii: integer;
begin
  if E_DEVISE.DataType <> '' then Exit;
  ii := E_DEVISE.ItemIndex;
  E_DEVISE.DataType := 'TTDEVISE';
  E_DEVISE.Reload;
  E_DEVISE.Value := V_PGI.DevisePivot;
  if ii > 0 then E_DEVISE.ItemIndex := E_DEVISE.Values.Count - 1;
end;

procedure TFSaisietr.AddEuroFranc;
begin
  E_DEVISE.DataType := '';
  E_DEVISE.Items.Clear;
  E_DEVISE.Values.Clear;
  E_DEVISE.Values.Add(V_PGI.DevisePivot);
  E_DEVISE.Items.Add(RechDom('TTDEVISE', V_PGI.DevisePivot, False));
  E_DEVISE.Value := E_DEVISE.Values[0];
  E_DEVISE.Enabled := True;
end;

{==========================================================================================}
{================================ Methodes de la form =====================================}
{==========================================================================================}
procedure TFSaisietr.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if MajEnCours or (not GS.SynEnabled) then Key := #0 else
    if Key = #127 then Key := #0 else
    if ((Key = '=') and ((GS.Col = SA_Debit) or (GS.Col = SA_Credit) or (GS.Col = SA_Gen) or (GS.Col = SA_Aux))) then Key := #0;
end;

procedure TFSaisieTr.FocusSurNatP;
begin
  if E_JOURNAL.Focused and VH^.JalLookUp then
  begin
    if E_NATUREPIECE.CanFocus then E_NATUREPIECE.SETFOCUS else if E_CONTREPARTIEGEN.CanFocus then E_CONTREPARTIEGEN.SetFocus;
  end;
  (*NextControl(Self,TRUE) ;*)
end;

procedure TFSaisietr.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var
  OkG, Vide: boolean;
  keys: TKeyboardState;
begin
  {Panel de Swap devise}
  if MajEnCours or (not GS.SynEnabled) then
  begin
    Key := 0;
    Exit;
  end;
  if ModeRA then
  begin
    case Key of
      VK_ESCAPE: BRAAbandonClick(nil);
      VK_F10: BRAValideClick(nil);
      //      VK_F1     : BCAideClick(Nil) ;
    end;
    Exit;
  end;
  if ((Action <> taConsult) and (ModeForce <> tmNormal)) then
  begin
    case ModeForce of
      tmDevise:
        begin
          if Shift <> [] then Exit;
          if (Key in [VK_RETURN, VK_ESCAPE, VK_PRIOR, VK_NEXT, VK_END, VK_HOME, VK_UP, VK_DOWN]) = False then Exit;
          if Key in [VK_RETURN, VK_ESCAPE] then ForcerMode(False, Key);
          Exit;
        end;
      tmPivot: if ((Shift <> [ssAlt]) or (Key <> VK_F8)) then
        begin
          if (Key in [VK_RETURN, VK_ESCAPE, VK_PRIOR, VK_NEXT, VK_END, VK_HOME, VK_UP, VK_DOWN]) = False then Exit;
        end;
    end;
  end;
  {Saisie normale}
  OkG := (Screen.ActiveControl = GS);
  Vide := (Shift = []);
  case Key of
    VK_F3: if ((OkG) and (Vide) and (GS.Col = SA_Gen)) then
      begin
        FocusSurNatP;
        Key := 0;
        AutoSuivant(False);
      end;
    VK_F4: if ((OkG) and (Vide) and (GS.Col = SA_Gen)) then
      begin
        FocusSurNatP;
        Key := 0;
        AutoSuivant(True);
      end;
    VK_F5: // BPY le 05/10/2004 et le 29/10/2004 : gestion du lookup std a la place du thcompte
      if ((OkG) and ((Vide) or (Shift = [ssCtrl]))) then
      begin
        Key := 0;
        //                    if ((GetParamSoc('SO_CPCODEAUXIONLY',false) = 'X') and (Shift = [ssCtrl])) then Shift := [];
        if ((GetParamSoc('SO_CPCODEAUXIONLY', false) = true) and (GS.Col = SA_Aux) and (Shift = [ssCtrl])) then
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
    VK_F6: if ((OkG) and (Vide)) then
      begin
        Key := 0;
        ClickSolde(False);
      end;
    VK_F8: {if Vide then BEGIN Key:=0 ; ClickControleTva(False,0) ; END else}
      if Shift = [ssAlt] then
      begin
        Key := 0;
        ClickSwapPivot;
      end; //else
    VK_F10: if Vide then
      begin
        Key := 0;
        FocusSurNatP;
        ClickValide;
      end;
    VK_ESCAPE: if Vide then
      begin
        Key := 0;
        ClickAbandon;
      end;
    VK_RETURN: if ((OkG) and (Vide)) then KEY := VK_TAB;
    VK_BACK: if ((OkG) and (Shift = [ssCtrl])) then
      begin
        Key := 0;
        VideZone(GS);
      end;
    VK_DELETE: if Shift = [ssCtrl] then // FQ 17382 SBO 12/04/2006
                 begin {JP 20/10/05 : Ajout du ssCtrl}
                 Key := 0;
                 ClickRAZLigne;
                 end;
    {^Home}36: if Shift = [ssCtrl] then
      begin
        Key := 0;
        GotoEntete;
      end;
    {A1.9}49..57: if Shift = [ssAlt] then
      begin
        AppelAuto(Key - 48);
        Key := 0;
      end;
    //{AA}     65 : if ((OkG) and (Shift=[ssAlt])) then BEGIN Key:=0 ; ClickVentil ; END ;
        {^A}65: if ((OkG) and (Shift = [ssCtrl])) then
      begin
        Key := 0;
        ClickVentil;
      end;
    {AC}67: if Shift = [ssAlt] then
      begin
        Key := 0;
        ClickComplement;
      end;
    {AE}69: if ((OkG) and (Shift = [ssAlt])) then
      begin
        Key := 0;
        ClickEche;
      end;
    {^F}70: if Shift = [ssCtrl] then
      begin
        Key := 0;
        ClickCherche;
      end;
    {AH}72: if Shift = [ssAlt] then
      begin
        Key := 0;
        ClickChancel;
      end;
    {AI}73: if Shift = [ssAlt] then
      begin
        Key := 0;
        ClickDevise;
      end;
    {AJ}74: if Shift = [ssAlt] then
      begin
        FocusSurNatP;
        Key := 0;
        ClickJournal;
      end;
    (* //TR A voir
    {AL}     76 : if ((OkG) and (Shift=[ssAlt])) then BEGIN Key:=0 ; CalculRefAuto(Nil) ; END ;
    *)
        {^R}82: if Shift = [ssCtrl] then
      begin
        Key := 0;
        ClickChoixRegime;
      end else
        {AR}if Shift = [ssAlt] then
      begin
        Key := 0;
        ClickModifRIB;
      end;
    {AS}83: if Shift = [ssAlt] then
      begin
        FocusSurNatP;
        Key := 0;
        ClickScenario;
      end;
    {AT}84: if Shift = [ssAlt] then
      begin
        Key := 0;
        ClickEtabl;
      end;
    {AV}86: if Shift = [ssAlt] then
      begin
        Key := 0;
        ClickModifTva;
      end;
    {^D}68: if Shift = [ssCtrl] then
      begin
        Key := 0;
        ClickDelettrage;
      end;
    {^L}76: if Shift = [ssCtrl] then
      begin
        Key := 0;
        Clicklettrage;
      end else
        {AL}if Shift = [ssAlt] then
      begin
        Key := 0;
        ClickLibelleAuto;
      end;
  end;
end;

procedure TFSaisietr.PosLesSoldes;
begin
  LE_DEBITTRESO.SetBounds(E_DEBITTRESO.Left, E_DEBITTRESO.Top, E_DEBITTRESO.Width, E_DEBITTRESO.Height);
  LE_CREDITTRESO.SetBounds(E_CREDITTRESO.Left, E_CREDITTRESO.Top, E_CREDITTRESO.Width, E_CREDITTRESO.Height);
  E_DEBITTRESO.Visible := False;
  LE_DEBITTRESO.Caption := '';
  E_CREDITTRESO.Visible := False;
  LE_CREDITTRESO.Caption := '';

  LSA_SoldeG.SetBounds(SA_SoldeG.Left, SA_SoldeG.Top, SA_SoldeG.Width, SA_SoldeG.Height);
  LSA_SoldeT.SetBounds(SA_SoldeT.Left, SA_SoldeT.Top, SA_SoldeT.Width, SA_SoldeT.Height);
  SA_SOLDEG.Visible := False;
  LSA_SOLDEG.Caption := '';
  SA_SOLDET.Visible := False;
  LSA_SOLDET.Caption := '';

  LSA_TotalDebit.SetBounds(SA_TotalDebit.Left, SA_TotalDebit.Top, SA_TotalDebit.Width, SA_TotalDebit.Height);
  LSA_TotalCredit.SetBounds(SA_TotalCredit.Left, SA_TotalCredit.Top, SA_TotalCredit.Width, SA_TotalCredit.Height);
  SA_TotalDebit.Visible := False;
  LSA_TotalDebit.Caption := '';
  SA_TotalCredit.Visible := False;
  LSA_TotalCredit.Caption := '';

  LSA_Solde.SetBounds(SA_Solde.Left, SA_Solde.Top, SA_Solde.Width, SA_Solde.Height);
  LSA_SoldeTR.SetBounds(SA_SoldeTR.Left, SA_SoldeTR.Top, SA_SoldeTR.Width, SA_SoldeTR.Height);
  SA_SOLDE.Visible := False;
  LSA_SOLDE.Caption := '';
  SA_SOLDETR.Visible := False;
  LSA_SOLDETR.Caption := '';
end;

procedure TFSaisietr.FormCreate(Sender: TObject);
begin
  WMinX := Width;
  WMinY := Height_Tres;
  GS.VidePile(True);
  GS.RowCount := 2;
  E_NUMREF.Visible := FALSE;
  tAppelLettrage := Tbits.Create;
//  if VFTreso then //TR
  //begin
    GS.TypeSais := tsTreso;
    LigneTreso := FALSE;
    GS.ListeParam := 'SSAISIE2';
    GS.FixedCols := 3;
    if VFEffet then E_JOURNAL.DataType := 'ttJalEffet' else E_JOURNAL.DataType := 'ttJalBanque';
{  end else
  begin
    GS.TypeSais := tsGene;
    GS.ListeParam := 'SSAISIE1';
    GS.FixedCols := 4;
    E_JOURNAL.DataType := 'ttJalSaisie';
  end;}
  SAJAL := nil;
  DecDev := V_PGI.OkdecV;
  PieceModifiee := False;
  Revision := False;
  TS := TList.Create;
  TPIECE := TList.Create;
  TDELECR := TList.Create;
  TDELANA := TList.Create;
  Scenario := TOBM.Create(EcrSais, '', True);
  Entete := TOBM.Create(EcrSais, '', True);
  FInfo := TInfoEcriture.Create ; // FQ 13246 : SBO 30/03/2005
  MemoComp := HTStringList.Create;
  ModifSerie := TOBM.Create(EcrGen, '', True);
  PieceConf := False;
  GestionRIB := '...';
  DejaRentre := False;
  OkScenario := False;
  ModeConf := '0';
  Volatile := False;
  NeedEdition := False;
  EtatSaisie := '';
  RentreDate := False;
  RegimeForce := False;
  //  ChangeSQL(TEcrGen);
  //  ChangeSQL(TEcrAna);
  SorteTva := stvDivers;
  ExigeTva := tvaMixte;
  ExigeEntete := tvaMixte;
  ChoixExige := False;
  BPopTva.ItemIndex := Ord(ExigeEntete) + 1;
  RegLoadToolbarPos(Self, 'SaisieTr');
  PosLesSoldes;
  {Param Liste}
  GS.ColLengths[SA_Gen] := VH^.Cpta[fbGene].Lg;
  if VH^.CPCodeAuxiOnly then GS.ColLengths[SA_Aux] := VH^.Cpta[fbAux].Lg else GS.ColLengths[SA_Aux] := 35;
  GS.ColLengths[SA_RefI] := 35;
  GS.ColLengths[SA_Lib] := 35;
  // SBO 01/07/2007 : enregistrement spécifique pour pb CWAS
  FBoInsertSpecif := EstSpecif('51215') ;
end;

procedure TFSaisietr.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  CanClose := FermerSaisie;
  FClosing := CanClose ; 
end;

procedure TFSaisietr.EditionSaisie;
var
  SWhere: string;
  i: integer;
  RR: RMVT;
begin
  if not NeedEdition then Exit;
  if EtatSaisie = '' then Exit;
  if HDiv.Execute(22, Caption, '') <> mrYes then Exit;
  NeedEdition := False;
  SWhere := '';
  for i := 0 to TPIECE.Count - 1 do
  begin
    RR := P_MV(TPIECE[i]).R;
    SWhere := SWhere + '(E_JOURNAL="' + RR.Jal + '" AND E_NUMEROPIECE=' + IntToStr(RR.Num) + ' AND E_QUALIFPIECE="' + RR.Simul + '" AND E_EXERCICE="' + RR.Exo + '")';
    if i < TPIECE.Count - 1 then SWhere := SWhere + ' OR ';
  end;
  if TPIECE.Count > 1 then SWhere := '(' + SWhere + ')';
  LanceEtat('E', 'SAI', EtatSaisie, True, False, False, nil, SWhere, '', False);
end;

procedure TFSaisietr.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  tAppelLettrage.Free;
  GS.VidePile(True);
  SAJAL.Free;
  SAJAL := nil; // FQ 5683
  VideListe(TS);
  TS.Free;
  TS := nil;
  VideListe(TPIECE);
  TPIECE.Free;
  TPIECE := nil;
  VideListe(TDELECR);
  TDELECR.Free;
  TDELECR := nil;
  VideListe(TDELANA);
  TDELANA.Free;
  TDELANA := nil;
  FreeAndNil(FInfo) ; // FQ 13246 : SBO 30/03/2005 
  Scenario.Free;
  Entete.Free;
  ModifSerie.Free;
  MemoComp.Free;
  MemoComp := nil;
  OBMT[1].Free;
  OBMT[2].Free; //TR
  PurgePopup(POPS);
  PurgePopup(POPZ);
  PurgePopup(POPZT);
//  if M.Treso then //TR
//  begin
    SA_DateC := 2;
    SA_NumL := 3;
//  end;
  RegSaveToolbarPos(Self, 'SaisieTr');
  if Parent is THPanel then
  begin
    case Self.Action of
      taCreat: _Bloqueur('nrSaisieCreat', False);
      taModif: _Bloqueur('nrSaisieModif', False);
    end;
    Action := caFree;
  end;
end;

procedure TFSaisietr.LectureSeule;
begin
  BDernPieces.Enabled := False;
end;

procedure TFSaisietr.FormShow(Sender: TObject);
var
  OkM, bc: boolean;
begin
  LookLesDocks(Self);
  RowEnterPourModeConsult := FALSE;
  MajEnCours := FALSE;
  if ((Action = taCreat) and (M.Etabl = '')) then M.Etabl := VH^.EtablisDefaut;
  InitLesComposants;
  E_REF.TEXT := '';
  E_NUMDEP.text := '';
  PieceConf := False;

//  if M.Treso then
    E_MODEPAIE.ItemIndex := 0;
  {else
  if Action <> taCreat then
    E_JOURNAL.DataType := 'ttJournal'
  else begin
    if ((Action = taCreat) and (M.Simul = 'N') and (M.ANouveau)) then
      E_JOURNAL.DataType := 'ttJalANouveau'
    else
    begin
      if M.Simul <> 'N' then
        E_JOURNAL.DataType := 'ttJalSansEcart'
      else
        E_JOURNAL.DataType := 'ttJalSaisie';
    end;
  end;}

  If Action=taCreat Then
    BEGIN
    // GP le 20/08/2008 21496
    RetoucheHVCPoursaisie(E_JOURNAL) ;
    RetoucheHVCPoursaisie(E_ETABLISSEMENT) ;
    END ;
  ModeCG := False;
  ModeRA := FALSE;
  Revision := False;
  OkM := ((M.Valide) and (Action = taModif));
  if OkM then Action := taConsult;
  DefautPied;
  DefautEntete;
  NbLigEcr := 0;
  NbLigAna := 0;
  ChargeLignes;
  //if OkM then Action:=taConsult ;
  case Action of
    taConsult:
      begin
        LectureSeule;
      end;
    taCreat:
      begin
        //if not M.Treso then BDernPieces.Enabled := True;
      end;
  end;
  if Action <> taModif then AffecteGrid(GS, Action);
  PieceModifiee := False;
  AfficheLeSolde(SA_Solde, 0, 0);
  AttribLeTitre;
  QuelNExo;
  if OkM then HPiece.Execute(11, caption, '');
  if Action = taConsult then
  begin
    if ((M.NumLigVisu <> 0) and (M.NumLigVisu <= GS.RowCount - 2)) then
    begin
      GS.Row := M.NumLigVisu;
      GS.Cells[SA_NumL, GS.Row] := '**' + GS.Cells[SA_NumL, GS.Row] + '**';
    end else if M.General <> '' then GS.Row := TrouveLigCompte(GS, SA_Gen, 0, M.General);
  end;
  GSRowEnter(nil, GS.Row, bc, False);
end;

{==========================================================================================}
{================================ Ecriture Fichier=========================================}
{==========================================================================================}
procedure TFSaisietr.MajCptesGeneNew;
var
  XD, XC: Double;
  Lig: integer;
  FRM: TFRM;
  ll: LongInt;
begin
  if PasModif then Exit;
  //if not M.Treso then if M.Simul <> 'N' then Exit;
  for Lig := 1 to GS.RowCount - 2 do if ((GS.Cells[SA_Gen, Lig] <> '') and (GS.Objects[SA_Gen, Lig] <> nil)) then
    begin
      Fillchar(FRM, SizeOf(FRM), #0);
      FRM.Cpt := GS.Cells[SA_Gen, Lig];
      RecupTotalPivot( GS, Lig, XD, XC ) ;// SBO 09/08/2007 FQ 20910 maj solde des comptes erroné en devise
      if not M.ANouveau then
      begin
        FRM.NumD := NumPieceInt;
        FRM.DateD := DateMvt(Lig);
        FRM.LigD := StrToInt(GS.Cells[SA_NumL, Lig]);
        AttribParamsNew(FRM, XD, XC, GeneTypeExo);
      end else
      begin
        FRM.Deb := XD;
        FRM.Cre := XC;
      end;
      if Action <> taConsult then AttribParamsComp(FRM, GS.Objects[SA_Gen, Lig]);
      LL := ExecReqMAJ(fbGene, M.ANouveau, Action <> taConsult, FRM);
      if ll <> 1 then V_PGI.IoError := oeSaisie ;
    end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 02/06/2006
Modifié le ... :   /  /    
Description .. : - FB 10678 - LG - 02/06/2006 - gestion des devises ds les 
Suite ........ : ano dynamiques
Mots clefs ... : 
*****************************************************************}
procedure TFSaisietr.MajCptesAno;
var
  lTS        : TList ;
  lStNatGene : string ;
  i          : integer ;
begin

 if PasModif then Exit;

 lTS := TList.Create ;

 for i := 1 to GS.RowCount - 2 do if ((GS.Cells[SA_Gen, i] <> '') and (GS.Objects[SA_Gen, i] <> nil)) then
  begin
   lStNatGene := '' ;
   if FInfo.LoadCompte( GS.Cells[SA_Gen,i] ) then
    lStNatGene := FInfo.GetString('G_NATUREGENE') ;
   AjouteAno(lTS,GetO(GS, i),lStNatGene,false ) ;
  end ;

 if not ExecReqMAJAno(lTS) then
  begin
   V_PGI.IoError := oeSaisie ;
   exit ;
  end ; // if

 for i := 0 to lTS.Count - 1 do
  if assigned(lTS[i]) then Dispose(lTS[i]);
 lTS.Free ;
 
end ;


procedure TFSaisietr.MajCptesAuxNew;
var
  XD, XC: Double;
  Lig: integer;
  Trouv: Boolean;
  FRM: TFRM;
  ll: LongInt;
begin
  if PasModif then Exit;
  //if not M.Treso then if M.Simul <> 'N' then Exit;
  Trouv := False;
  for Lig := 1 to GS.RowCount - 2 do if ((GS.Cells[SA_Aux, Lig] <> '') and (GS.Objects[SA_Aux, Lig] <> nil))
    then
    begin
      Trouv := True;
      Break;
    end;
  if not Trouv then Exit;

  for Lig := 1 to GS.RowCount - 2 do if ((GS.Cells[SA_Aux, Lig] <> '') and (GS.Objects[SA_Aux, Lig] <> nil)) then
    begin
      Fillchar(FRM, SizeOf(FRM), #0);
      FRM.Cpt := GS.Cells[SA_Aux, Lig];
      RecupTotalPivot( GS, Lig, XD, XC ) ;// SBO 09/08/2007 FQ 20910 maj solde des comptes erroné en devise
      if not M.ANouveau then
      begin
        FRM.NumD := NumPieceInt;
        FRM.DateD := DateMvt(Lig);
        FRM.LigD := StrToInt(GS.Cells[SA_NumL, Lig]);
        AttribParamsNew(FRM, XD, XC, GeneTypeExo);
      end else
      begin
        FRM.Deb := XD;
        FRM.Cre := XC;
      end;
      if Action <> taConsult then AttribParamsComp(FRM, GS.Objects[SA_Aux, Lig]);
      ll := ExecReqMAJ(fbAux, M.ANouveau, Action <> taConsult, FRM);
      if ll <> 1 then V_PGI.IoError := oeSaisie;
    end;
end;

function TFSaisietr.QuelEnc(Lig: Integer): String3;
var CGen : TGGeneral ;
begin
  CGen:=GetGGeneral(GS,Lig) ;
  // Pour les comptes lettrables divers... FQ 16136 SBO 09/08/2005
  if ( CGen <> Nil ) and CGen.Lettrable and ( CGen.NatureGene = 'DIV' )
    then result := 'RIE'
    // Sinon récultat habituel
    else Result := SensEnc(ValD(GS,Lig),ValC(GS,Lig)) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 29/07/2003
Modifié le ... :   /  /
Description .. : LG - 29/07/2003 - FB 12574 - l'affeectation de
Suite ........ : e_encaissemetne etait incorrecte
Mots clefs ... :
*****************************************************************}
procedure TFSaisietr.RecupTronc(Lig: Integer; MO: TMOD; var OBM: TOBM);
var
  CGen: TGGeneral;
  CAux: TGTiers;
  NatG, NatT: String3;
  TVA, TPF: String3;
  {##ANA##
  T: TList;
  OBA: TObjAna;}
  OBA : TOB;
  {##ANA##}
  NumAxe, NumV: integer;
  Aux, Gen : string; {JP 14/11/05 : FQ 16807 : Nouvelle gestion de la contrepartie}
begin
  if PasModif then Exit;
  CGen := nil;
  CAux := nil;
  NatG := '';
  NatT := '';
  OBM := nil;
  if GS.Objects[SA_Gen, Lig] <> nil then
  begin
    CGen := TGGeneral(GS.Objects[SA_Gen, Lig]);
    NatG := CGen.NatureGene;
  end;
  if GS.Objects[SA_Aux, Lig] <> nil then
  begin
    CAux := TGTiers(GS.Objects[SA_Aux, Lig]);
    NatT := CAux.NatureAux;
  end;
  if GS.Objects[SA_Exo, Lig] <> nil then OBM := GetO(GS, Lig);
  if OBM <> nil then
  begin
    OBM.SetCotation(0);
    OBM.SetMPACC;
  end;
  // GGGGG
  OBM.PutMvt('E_QUALIFPIECE', M.Simul);

  // Modif Lettrage des comptes divers  FQ 16136 SBO 09/08/2005
  CGetTypeMvt( OBM, FInfo ) ;

  if OBM.GetMvt('E_REGIMETVA') = '' then OBM.PutMvt('E_REGIMETVA', GeneRegTVA);
  if not M.ANouveau then
  begin
    OBM.PutMvt('E_ECRANOUVEAU', 'N');
    OBM.PutMvt('E_VALIDE', '-');
  end else
  begin
    OBM.PutMvt('E_ECRANOUVEAU', 'H');
    OBM.PutMvt('E_VALIDE', 'X');
  end;
  OBM.PutMvt('E_BUDGET', '');
  OBM.PutMvt('E_ECHE', '-');
  if CAux <> nil then OBM.PutMvt('E_REGIMETVA', CAux.RegimeTVA);
  if CGen <> nil then
  begin
    TVA := CGen.Tva;
    if TVA = '' then TVA := GeneTVA;
    TPF := CGen.Tpf;
    if TPF = '' then TPF := GeneTPF;
    if OBM.GetMvt('E_TVA') = '' then OBM.PutMvt('E_TVA', GeneTva);
    if OBM.GetMvt('E_TPF') = '' then OBM.PutMvt('E_TPF', GeneTpf);
    OBM.PutMvt('E_BUDGET', CGen.Budgene);
    (*
    If M.Effet Then OkOk:=(EstLettrable(Lig) in [MonoEche,MultiEche]) And (CGen.General=SAJAL.Treso.Cpt) And (MO<>Nil)
               Else OkOk:=((Lettrable(CGen)=MonoEche) and (MO<>Nil)) ;
    *)
    if ((Lettrable(CGen) = MonoEche) and (MO <> nil)) then
    begin

      // récupération de la 1ère échéance
      RecupEche( Lig, MO.MODR , 1 , OBM ) ; // FQ 18736 SBO 07/09/2006 : pb d'affectation des zones de lettrage pour les comptes divers lettrables

      // Ajout Sinon bug LCD FQ 16584 SBO 06/08/2005
      if CGen.pointable then
        OBM.PutMvt('E_ETATLETTRAGE',	'RI') ;  // FQ 16761 SBO 30/09/2005

    end;
  end;
  //if M.Treso then
  //begin
    if GS.Cells[SA_Gen, Lig] <> SAJAL.TRESO.Cpt then
    begin
      OBM.PutMvt('E_CONTREPARTIEGEN', SAJAL.TRESO.Cpt);
      OBM.PutMvt('E_CONTREPARTIEAUX', '');
      if M.Effet and (SAJAL.TRESO.TypCtr = Ligne) and SAJAL.TRESO.Collectif then OBM.PutMvt('E_CONTREPARTIEAUX', GS.Cells[SA_Aux, Lig]);
    end else
    begin
      case SAJAL.TRESO.TypCtr of
        PiedDC, PiedSolde, AuChoix:
          begin
            {JP 14/11/05 : FQ 16807 : Nouvelle gestion de la contrepartie}
            GetContrepartie(Aux, Gen);
            OBM.PutMvt('E_CONTREPARTIEGEN', Gen);
            OBM.PutMvt('E_CONTREPARTIEAUX', Aux);
          end;
        Ligne:
          begin
            OBM.PutMvt('E_CONTREPARTIEGEN', GS.Cells[SA_Gen, Lig - 1]);
            OBM.PutMvt('E_CONTREPARTIEAUX', GS.Cells[SA_Aux, Lig - 1]);
          end;
      end;
    end;
  {end else
  begin
    if Lig = LAUX then
    begin
      OBM.PutMvt('E_CONTREPARTIEGEN', GS.Cells[SA_Gen, LHT]);
      OBM.PutMvt('E_CONTREPARTIEAUX', GS.Cells[SA_Aux, LHT]);
    end else
    begin
      OBM.PutMvt('E_CONTREPARTIEGEN', GS.Cells[SA_Gen, LAUX]);
      OBM.PutMvt('E_CONTREPARTIEAUX', GS.Cells[SA_Aux, LAUX]);
    end;
  end;}
  OBM.PutMvt('E_DATEMODIF', NowFutur);
  OBM.PutMvt('E_CONFIDENTIEL', ModeConf);
  {##ANA##
  OBA := TObjAna(GS.Objects[SA_DateC, Lig]);
  if OBA <> nil then
  begin
    for NumAxe := 1 to MaxAxe do
    begin
      if OBA.AA[NumAxe] <> nil then
      begin
        T := OBA.AA[NumAxe].L;
        for NumV := 0 to T.Count - 1 do
        begin
          PutMvtA(OBA.AA[NumAxe], P_TV(T.Items[NumV]).F, 'Y_CONTREPARTIEGEN', OBM.GetMvt('E_CONTREPARTIEGEN'));
          PutMvtA(OBA.AA[NumAxe], P_TV(T.Items[NumV]).F, 'Y_CONTREPARTIEAUX', OBM.GetMvt('E_CONTREPARTIEAUX'));
        end;
      end;
    end;
  end;}
  {Mise à jour contreparties analytiques}
  if OBM.Detail.Count > 0 then
    for NumAxe := 0 to OBM.Detail.Count - 1 do
      for NumV := 0 to OBM.Detail[NumAxe].Detail.Count - 1 do begin
        OBA := OBM.Detail[NumAxe].Detail[NumV];
        OBA.PutValue('Y_CONTREPARTIEGEN', OBM.GetMvt('E_CONTREPARTIEGEN'));
        OBA.PutValue('Y_CONTREPARTIEAUX', OBM.GetMvt('E_CONTREPARTIEAUX'));
      end ;
  {##ANA##}
end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 27/03/2002
Modifié le ... :   /  /
Description .. : LG* nouvelle fonction de saisUtil suite au modif du tobm
Mots clefs ... :
*****************************************************************}
procedure TFSaisietr.RecupAnal(Lig: integer; var OBA: TOB; NumAxe, NumV: integer);
begin
  // Attention NumV démarre à 0
  if PasModif then Exit;
  //LG* 27/03/2002 nv fonction de SaisUtil
  if NumV = 0 then OBA.PutValue('Y_TYPEMVT', 'AE')
  else OBA.PutValue('Y_TYPEMVT', 'AL');
  OBA.PutValue('Y_TYPEANALYTIQUE', '-');
  //  OBA.PutValue('Y_BLOCNOTE',TT.M[NumV]);
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
end;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 29/07/2003
Modifié le ... : 01/04/2004
Description .. : LG - 29/07/2003 - FB 12574 - l'affeectation de
Suite ........ : e_encaissement etait incorrecte
Suite ........ : -01/04/2004 - LG suppression de SetMPACCDB(TEcrGen)
Suite ........ : ;
Mots clefs ... :
*****************************************************************}
procedure TFSaisietr.RecupEche(Lig: integer; R: T_ModeRegl; NumEche: integer; OBM: TOBM);
var
  Deb: boolean;
  i: Integer;
  Coll: string;
begin
  if PasModif then exit;
  OBM.PutMvt('E_NUMECHE', NumEche);
  OBM.PutMvt('E_ECHE', 'X');
  OBM.PutMvt('E_NIVEAURELANCE', R.TabEche[NumEche].NiveauRelance);
  OBM.PutMvt('E_MODEPAIE', R.TabEche[NumEche].ModePaie);
  OBM.PutMvt('E_DATEECHEANCE', R.TabEche[NumEche].DateEche);
  OBM.PutMvt('E_DATEVALEUR', R.TabEche[NumEche].DateValeur);
  OBM.PutMvt('E_COUVERTURE', R.TabEche[NumEche].Couverture);
  OBM.PutMvt('E_COUVERTUREDEV', R.TabEche[NumEche].CouvertureDev);
  OBM.PutMvt('E_ETATLETTRAGE', R.TabEche[NumEche].EtatLettrage);
  OBM.PutMvt('E_LETTRAGE', R.TabEche[NumEche].CodeLettre);
  OBM.PutMvt('E_LETTRAGEDEV', R.TabEche[NumEche].LettrageDev);
  OBM.PutMvt('E_DATEPAQUETMAX', R.TabEche[NumEche].DatePaquetMax);
  OBM.PutMvt('E_DATEPAQUETMIN', R.TabEche[NumEche].DatePaquetMin);
  OBM.PutMvt('E_DATERELANCE', R.TabEche[NumEche].DateRelance);
  OBM.PutMvt('E_ENCAISSEMENT', QuelEnc(Lig));

  OBM.PutMvt('E_DEBIT', 0);
  OBM.PutMvt('E_CREDIT', 0);

  Deb := (ValD(GS, Lig) <> 0);
  if Deb then
  begin
    OBM.PutMvt('E_DEBIT', R.TabEche[NumEche].MontantP);
    OBM.PutMvt('E_DEBITDEV', R.TabEche[NumEche].MontantD);
  end else
  begin
    OBM.PutMvt('E_CREDIT', R.TabEche[NumEche].MontantP);
    OBM.PutMvt('E_CREDITDEV', R.TabEche[NumEche].MontantD);
  end;

  if ModeDevise then
  begin
    if Deb then OBM.PutMvt('E_DEBITDEV', R.TabEche[NumEche].MontantD)
    else OBM.PutMvt('E_CREDITDEV', R.TabEche[NumEche].MontantD);
  end;
  OBM.PutMvt('E_ORIGINEPAIEMENT', R.TabEche[NumEche].DateEche);
  OBM.PutMvt('E_CODEACCEPT', MPTOACC(OBM.GetMvt('E_MODEPAIE')));
  {#TVAENC}
  if VH^.OuiTvaEnc then
  begin
    Coll := OBM.GetMvt('E_GENERAL');
    if EstCollFact(Coll) then
    begin
      if ((SorteTva = stvDivers) and (R.ModifTva)) then
      begin
        for i := 1 to 4 do OBM.PutMvt('E_ECHEENC' + IntToStr(i), Arrondi(R.TabEche[NumEche].TAV[i], V_PGI.OkDecV));
        OBM.PutMvt('E_ECHEDEBIT', Arrondi(R.TabEche[NumEche].TAV[5], V_PGI.OkDecV));
      end;
    end;

  end;
end;

procedure TFSaisietr.GetAna(Lig: Integer);
var
  i, j: integer;
  OBM: TOBM;
  OBA: TOB;
begin
  if PasModif then exit;
  OBM := GetO(GS, Lig);
  if OBM = nil then Exit;
  for i := 0 to OBM.Detail.Count - 1 do
    for j := 0 to OBM.Detail[i].Detail.Count - 1 do
    begin
      OBA := OBM.Detail[i].Detail[j];
      RecupAnal(Lig, OBA, i, j);
      // SBO 01/07/2007 : enregistrement spécifique pour pb CWAS
      if FBoInsertSpecif
        then begin
             if not CTobInsertDB( OBA ) then
               V_PGI.IOError := oeSaisie ;
             end
        else OBA.InsertDB(nil);
    end;
end;

procedure TFSaisietr.GetEcr(Lig: Integer);
var
  M: TMOD;
  R: T_ModeRegl;
  i: integer;
  CGen: TGGeneral;
  OkOk: Boolean;
  OBM: TOBM;
begin
  if PasModif then exit;
  M := TMOD(GS.Objects[SA_NumP, Lig]);
  CGen := GetGGeneral(GS, Lig);
  if CGen = nil then Exit;
  if ((M = nil) or (Lettrable(CGen) = MonoEche)) then
  begin
    OkOk := TRUE;
    if Self.M.MAJDirecte then
    begin
      OBM := GetO(GS, Lig);
      if OBM <> nil then OkOk := OBM.Status = EnMemoire;
    end;
    if OkOk then
    begin
      RecupTronc(Lig, M, OBM);
      OnUpdateEcritureTOB(OBM, Action,[]);
      {JP 16/11/05 : On insère que les écritures comptables, l'analytique étant intégrée dans GetAna}
      if FBoInsertSpecif  // SBO 01/07/2007 : enregistrement spécifique pour pb CWAS
        then begin
             if not CTobInsertDB( OBM, True, False, False ) then
               V_PGI.IOError := oeSaisie ;
             end
        else OBM.InsertDBByNivel(False, 0, 0);
    end;
  end else
  begin
    R := M.MODR;
    for i := 1 to R.NbEche do
    begin
      OkOk := TRUE;
      if Self.M.MAJDirecte then
      begin
        OBM := GetO(GS, Lig);
        if OBM <> nil then OkOk := OBM.Status = EnMemoire;
      end;
      if OkOk then
      begin
        RecupTronc(Lig, M, OBM);
        RecupEche(Lig, R, i, OBM);
        OnUpdateEcritureTOB(OBM, Action,[]);
        {JP 16/11/05 : On insère que les écritures comptables, l'analytique étant intégrée dans GetAna}
        if FBoInsertSpecif  // SBO 01/07/2007 : enregistrement spécifique pour pb CWAS
          then begin
               if not CTobInsertDB( OBM, True, False, False ) then
                 V_PGI.IOError := oeSaisie ;
               end
          else OBM.InsertDBByNivel(False, 0, 0);
      end;
    end;
  end;
end;

procedure TFSaisietr.RecupFromGrid(Lig: integer; MO: TMOD; NumE: integer);
var
  OBM: TOBM;
  St: string;
  R: T_MODEREGL;
  ie, k: integer;
  Sens: integer;
  SQL: string;
  CGen: TGGeneral;
  CAux: TGTiers;
  NatT, NatG : String3;
  Stamp: TDateTime;
  XStamp: TDELMODIF;
begin
  OBM := GetO(GS, Lig); // if OBM.Etat=Inchange then Exit ;
  if ValD(GS, Lig) <> 0 then Sens := 1 else Sens := 2;
  if MO <> nil then
  begin
    R := MO.MODR;
    ie := NumE;
    if ie <= 0 then ie := 1;
    if Sens = 1 then
    begin
      OBM.PutMvt('E_DEBIT', R.TabEche[ie].MontantP);
      OBM.PutMvt('E_CREDIT', 0);
      OBM.PutMvt('E_DEBITDEV', R.TabEche[ie].MontantD);
      OBM.PutMvt('E_CREDITDEV', 0);
    end else
    begin
      OBM.PutMvt('E_CREDIT', R.TabEche[ie].MontantP);
      OBM.PutMvt('E_DEBIT', 0);
      OBM.PutMvt('E_CREDITDEV', R.TabEche[ie].MontantD);
      OBM.PutMvt('E_DEBITDEV', 0);
    end;
    OBM.PutMvt('E_NUMECHE', ie);
    OBM.PutMvt('E_MODEPAIE', R.TabEche[ie].ModePaie);
    OBM.PutMvt('E_DATEECHEANCE', R.TabEche[ie].DateEche);
    OBM.PutMvt('E_ECHE', 'X');
    OBM.PutMvt('E_ENCAISSEMENT', QuelEnc(Lig));
    OBM.PutMvt('E_ORIGINEPAIEMENT', R.TabEche[ie].DateEche);
  end else
  begin
    OBM.PutMvt('E_ECHE', '-');
  end;
  NatG := '';
  NatT := '';
  if GS.Objects[SA_Gen, Lig] <> nil then
  begin
    CGen := TGGeneral(GS.Objects[SA_Gen, Lig]);
    NatG := CGen.NatureGene;
  end;
  if GS.Objects[SA_Aux, Lig] <> nil then
  begin
    CAux := TGTiers(GS.Objects[SA_Aux, Lig]);
    NatT := CAux.NatureAux;
  end;

  // Modif Lettrage des comptes divers  FQ 16136 SBO 09/08/2005
  CGetTypeMvt( OBM, FInfo ) ;

  OBM.PutMvt('E_NATUREPIECE', E_NATUREPIECE.Value);
  OBM.PutMvt('E_ETABLISSEMENT', E_ETABLISSEMENT.Value);
  OBM.PutMvt('E_DATECOMPTABLE', DatePiece);
  {$IFNDEF SPEC302}
  OBM.PutMvt('E_PERIODE', GetPeriode(DatePiece));
  OBM.PutMvt('E_SEMAINE', NumSemaine(DatePiece));
  {$ENDIF}
  OBM.PutMvt('E_CONFIDENTIEL', ModeConf);
  if M.ANouveau then OBM.PutMvt('E_VALIDE', 'X');
  OnUpdateEcritureTOB(OBM, Action,[]);
  St := OBM.StPourUpdate;
  if St = '' then Exit;
  SQL := 'UPDATE ECRITURE SET ' + St + ', E_DATEMODIF="' + UsTime(NowFutur) + '"'
    + ' Where  ' + WhereEcriture(tsGene, M, False) + ' AND E_NUMLIGNE=' + IntToStr(Lig);
  if NumE > 0 then SQL := SQL + ' AND E_NUMECHE=' + IntToStr(NumE);
  {Réseau}
  Stamp := 0;
  for k := Lig - 1 to TDELECR.Count - 1 do
  begin
    XStamp := TDELMODIF(TDELECR[k]);
    if ((XStamp.NumLigne = Lig) and ((XStamp.NumEcheVent = NumE) or (NumE = 0))) then
    begin
      ;
      Stamp := XStamp.DateModification;
      Break;
    end;
  end;
  if Stamp > 0 then SQL := SQL + ' AND E_DATEMODIF="' + UsTime(Stamp) + '"';
  {Lancement}
  if ExecuteSQL(SQL) <= 0 then V_PGI.IOError := oeSaisie;
end;

procedure TFSaisietr.GetEcrGrid(Lig: Integer);
var
  M: TMOD;
  R: T_ModeRegl;
  i: integer;
  CGen: TGGeneral;
  t: TLettre;
begin
  if PasModif then exit;
  M := TMOD(GS.Objects[SA_NumP, Lig]);
  CGen := GetGGeneral(GS, Lig);
  if CGen = nil then Exit;
  t := Lettrable(CGen);
  if ((M = nil) or (t = MonoEche)) then RecupFromGrid(Lig, M, 0) else
  begin
    R := M.MODR;
    for i := 1 to R.NbEche do RecupFromGrid(Lig, M, i);
  end;
end;

procedure TFSaisietr.ValideLignesGene;
var
  i: integer;
begin
  if PasModif then Exit;
  LAUX := TrouveLigTiers(GS, 0);
  if LAUX < 0 then LAUX := 1;
  LHT := TrouveLigHT(GS, 0, False);
  if LHT < 0 then
  begin
    if LAUX <> 2 then LHT := 2 else LHT := 1;
  end;
  InitMove(GS.RowCount - 2, '');
  if Revision then
  begin
    for i := 1 to GS.RowCount - 2 do
    begin
      if MoveCur(FALSE) then ;
      GetEcrGrid(i);
      if V_PGI.IOError <> oeOK then Break;
    end;
  end else
  begin
    for i := 1 to GS.RowCount - 2 do
    begin
      if MoveCur(FALSE) then ;
      GetEcr(i);
      if V_PGI.IOError <> oeOK then Break;
    end;
  end;
  FiniMove;
end;

procedure TFSaisietr.ValideLeJournalNew;
var
  Per: Byte;
  DD: TDateTime;
  FRM: TFRM;
  ll: LongInt;
begin
  if PasModif then Exit;
  //if not M.Treso then if M.Simul <> 'N' then Exit;
  Fillchar(FRM, SizeOf(FRM), #0);
  FRM.CPT := SAJAL.Journal;
  FRM.NumD := NumPieceInt;
  DD := DateMvt(1);
  FRM.DateD := DD;
  if not M.ANouveau then
  begin
    AttribParamsNew(FRM, SI_TotDP, SI_TotCP, GeneTypeExo);
  end else
  begin
    FRM.DE := SI_TotDP;
    FRM.CE := SI_TotCP;
    FRM.Deb := 0;
    FRM.Cre := 0;
    FRM.DS := 0;
    FRM.CS := 0;
    FRM.DP := 0;
    FRM.CP := 0;
  end;
  ll := ExecReqMAJ(fbJal, M.ANouveau, False, FRM);
  if ll <> 1 then V_PGI.IoError := oeSaisie;
  {Dévalidation éventuelle période+jal}
  if GeneTypeExo = teEncours then
  begin
    Per := QuellePeriode(DD, VH^.Encours.Deb);
    if SAJAL.ValideEN[Per] = 'X' then ADevalider(E_JOURNAL.Value, DD);
  end else
  begin
    Per := QuellePeriode(DD, VH^.Suivant.Deb);
    if SAJAL.ValideEN1[Per] = 'X' then ADevalider(E_JOURNAL.Value, DD);
  end;
  MarquerPublifi(True);
end;

procedure TFSaisietr.ValideLignesAnal;
var
  i: integer;
begin
  if PasModif then exit;
  for i := 1 to GS.RowCount - 2 do
    GetAna(i);
end;

procedure TFSaisietr.ValideLesComptes;
var
//  OkOk         : Boolean ;
  lInCpt       : Integer ;
  lInCptAxe    : Integer ;
  lInCptVentil : Integer ;
  OBM          : TOBM ;
begin
  if PasModif then exit;

  MajCptesGeneNew;
  MajCptesAuxNew;
  MajCptesAno ;

  {##ANA##
  if not (VH^.AnaCroisaxe) then
  begin
    //MajCptesSectionNew;

  end
  else
  begin}

  //SG6 27.01.05 Gestion mode croisaxe
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
 {##ANA## end;}

end;

procedure TFSaisietr.AttribRegimeEtTVA;
var
  ia: integer;
  CGen: TGGeneral;
begin
  ia := TrouveLigTiers(GS, 0);
  RenseigneRegime(ia);
  ia := TrouveLigHT(GS, 0, False);
  if ia <= 0 then exit;
  if GS.Cells[SA_Gen, ia] <> '' then
  begin
    CGen := GetGGeneral(GS, ia);
    if CGen <> nil then
    begin
      GeneTVA := CGen.TVA;
      {$IFDEF ESP}
      if not IsTiersSoumisTPF(GS) then
        GeneTPF := ''
      else
        {$ENDIF ESP}
        GeneTPF := CGen.TPF;
    end;
  end;
end;

procedure TFSaisietr.ValideLaPiece;
begin
  if not GS.SynEnabled then Exit;
  GridEna(GS, False);
  if V_PGI.IOError = oeOK then
    ValideLignesGene;
  if V_PGI.IOError = oeOK then
    ValideLignesAnal;
end;

procedure TFSaisietr.ValideLeReste;
begin
  V_PGI.IOError := oeOK;
  if V_PGI.IOError = oeOK then
    ValideLesComptes;
  if V_PGI.IOError = oeOK then
  begin
    ValideLeJournalNew;
    if V_PGI.IOError = oeOK then
      if M.MajDirecte then RetouchePiece;
  end;
  GridEna(GS, True);
end;

{==========================================================================================}
{================================ Barre d'outils ==========================================}
{==========================================================================================}
procedure TFSaisietr.ForcerMode(Cons: boolean; Key: Word);
begin
  if ((Cons) and (Modeforce = tmDevise)) then Exit;
  if ((not Cons) and (Modeforce = tmNormal)) then Exit;
  if Cons then
  begin
    AffecteGrid(GS, taConsult);
    ModeForce := tmDevise;
    GS.SetFocus;
    PEntete.Enabled := False;
    Outils.Enabled := False;
    PForce.Align := alClient;
    PForce.Parent := PPied;
    PForce.Visible := True;
    AfficheLeSolde(HForce, SI_TotDP, SI_TotCP);
  end else
  begin
    if Key = VK_RETURN then SoldeLaligne(GS.Row);
    PEntete.Enabled := True;
    Outils.Enabled := True;
    Outils.SetFocus;
    GS.CacheEdit;
    AffecteGrid(GS, Action);
    GS.Col := SA_Gen;
    if GS.Row = 1 then GS.Row := 2 else GS.Row := 1;
    GS.SetFocus;
    GS.MontreEdit;
    ModeForce := tmNormal;
    PForce.Visible := False;
  end;
end;

procedure TFSaisietr.RetouchePiece;
var
  i: Integer;
  OMODR: TMOD;
  OBM: TOBM;
  R: T_ModeRegl;
  CGen: TGGeneral;
  CAux: TGTiers;
  sSQL, Gen, Aux, EtatLettrage, CodeLettre: string;
begin
  //  QEcr.Close;
  //  QECr.SQL.Clear;
    {JP 24/09/04 : On ne peut gérer ici E_TRESOSYNCHRO comme dans la saisie car E_QUALIFPIECE est à V_PGI.User pour
                   le lettrage ce qui fait que la fonction ULibTrSynchro.MajTresoEcritureTOB ne traite pas les écritures
                   issues de ces traitements. Au départ, Stéphane avait mis en place la fonction MajTresoEcritureTOB
                   dans GetEcr et RecupFromGrid mais c'était inutile. Comme par ailleurs, le traitement ne concernce
                   que des écritures de règlement, d'effets ou des collectifs TIC/TID, donc synchronisables on passe
                   en dur E_TRESOSYNCHRO à "CRE" (FQ TRESO 10104). On risque éventuellement d'écraser E_TRESOSYNCHRO
                   qui peut être à LET après le lettrage, mais il me parait plus imple de traiter le problème en Tréso.}
  // SBO 24/08/2005 : E_TRESOSYNCHRO à CRE pour les comptes qui ne sont pas divers lettrables
  sSQL := 'UPDATE ECRITURE SET E_TRESOLETTRE="-", E_QUALIFPIECE="N", E_TRESOSYNCHRO = "' + ets_Nouveau + '" WHERE ';
  sSQL := sSQL + 'E_JOURNAL="' + SaJal.Journal + '" AND E_EXERCICE="' + QuelExo(E_DateComptable.EditText) + '"';
  sSQL := sSQL + ' AND E_DATECOMPTABLE="' + UsDateTime(StrToDate(E_DATECOMPTABLE.TEXT)) + '" AND E_NUMEROPIECE=' + IntToStr(NumPieceInt);
  sSQL := sSQL + ' AND E_GENERAL NOT IN ( SELECT G_GENERAL FROM GENERAUX WHERE G_NATUREGENE="DIV" AND G_LETTRABLE="X")' ;
  ExecuteSQL(sSQL);
  // SBO 24/08/2005 : E_TRESOSYNCHRO à RIE Pour les comptes divers lettrables
  sSQL := 'UPDATE ECRITURE SET E_TRESOLETTRE="-", E_QUALIFPIECE="N", E_TRESOSYNCHRO = "' + ets_Rien + '" WHERE ';
  sSQL := sSQL + 'E_JOURNAL="' + SaJal.Journal + '" AND E_EXERCICE="' + QuelExo(E_DateComptable.EditText) + '" ';
  sSQL := sSQL + 'AND E_DATECOMPTABLE="' + UsDateTime(StrToDate(E_DATECOMPTABLE.TEXT)) + '" AND E_NUMEROPIECE=' + IntToStr(NumPieceInt);
  sSQL := sSQL + ' AND E_GENERAL IN ( SELECT G_GENERAL FROM GENERAUX WHERE G_NATUREGENE="DIV" AND G_LETTRABLE="X")' ;
  ExecuteSQL(sSQL);

  sSQL := 'UPDATE ECRITURE SET E_TRESOLETTRE="-" WHERE ';
  sSQL := sSQL + 'E_GENERAL=:GEN AND E_AUXILIAIRE=:AUX AND E_ETATLETTRAGE=:ETATL AND E_LETTRAGE=:CODE';
  for i := 1 to GS.RowCount - 2 do
  begin
    OBM := GetO(GS, i);
    if OBM <> nil then
    begin
      OMODR := TMOD(GS.Objects[SA_NumP, i]);
      if OMODR <> nil then
      begin
        R := OMODR.MODR;
        if Trim(R.TabEche[1].CodeLettre) <> '' then
        begin
          Gen := '';
          Aux := '';
          CodeLettre := '';
          EtatLettrage := 'AL';
          CGen := GetGGeneral(GS, i);
          CAux := GetGTiers(GS, i);
          if CGen <> nil then Gen := CGen.General;
          if CAux <> nil then Aux := CAux.Auxi;
          EtatLettrage := R.TabEche[1].EtatLettrage;
          CodeLettre := R.TabEche[1].CodeLettre;
          if (EtatLettrage = 'PL') or (EtatLettrage = 'TL') then
          begin
            sSQL := 'UPDATE ECRITURE SET E_TRESOLETTRE="-" WHERE ';
            sSQL := sSQL + 'E_GENERAL="' + Gen + '" AND E_AUXILIAIRE="' + Aux + '" AND E_ETATLETTRAGE="' + EtatLettrage + '" AND E_LETTRAGE="' + CodeLettre + '"';
            ExecuteSQL(sSQL);
          end;
        end;
      end;
    end;
  end;
end;

function TFSaisietr.FabricFromM: RMVT;
var
  X: RMVT;
begin
  X.Axe := '';
  X.Etabl := E_ETABLISSEMENT.Value;
  X.Jal := E_JOURNAL.Value;
  X.Exo := QuelExoDT(StrToDate(E_DATECOMPTABLE.Text));
  X.CodeD := E_DEVISE.Value;
  X.Simul := 'N';
  X.Nature := E_NATUREPIECE.Value;
  X.DateC := StrToDate(E_DATECOMPTABLE.Text);
  X.DateTaux := DEV.DateTaux;
  X.Num := NumPieceInt;
  X.TauxD := DEV.Taux;
  X.Valide := False;
  X.ANouveau := M.ANouveau;
  Result := X;
end;

procedure TFSaisietr.StockeLaPiece;
var
  X: RMVT;
  P: P_MV;
begin
  if Action <> taCreat then Exit;
  FillChar(X, Sizeof(X), #0);
  X := FabricFromM;
  P := P_MV.Create;
  P.R := X;
  TPIECE.Add(P);
end;

// GG COM
{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 25/09/2002
Modifié le ... :   /  /
Description .. : -25/09/2002- Ajout de la recopie de e_refrevision pour
Suite ........ : l'ensemble de la piece
Suite ........ : - 04/08/2004 - FB 12694 - recriture de la fct : on recher le
Suite ........ : prems e_refrevision diff de 0 et on le recopie sur l'ensemble
Suite ........ : de la peice
Mots clefs ... :
*****************************************************************}
function TFSaisietr.MajIOComTreso: Boolean;
var
  i: integer;
  O: TOBM;
  lInRefRevision: integer;
begin
  result := false;
  lInRefRevision := 0;
  for i := 1 to GS.RowCount - 2 do
  begin
    O := GetO(GS, i);
    if O = nil then Continue;
    if O.GetMvt('E_REFREVISION') <> 0 then
    begin
      lInRefRevision := O.GetValue('E_REFREVISION'); // on recupere le code revision
      result := true;
      break;
    end;
  end; // for
  if result then
    for i := 1 to GS.RowCount - 2 do
    begin
      O := GetO(GS, i);
      if O = nil then Continue;
      O.PutMvt('E_IO', 'X');
      // affectation du code revision s'il est different de 0
      if (lInRefRevision <> 0) and (O.GetValue('E_REFREVISION') = 0) then O.PutValue('E_REFREVISION', lInRefRevision);
      if M.MajDirecte and StatusOBMOk(O) then MajDirecteTable(i, '');
    end; // for
end;

procedure TFSaisietr.ClickValide;
var
  io, io1: TIOErr;
  OBM: TOBM;
  i : integer;
begin
  if not GS.SynEnabled then Exit;
  if SAJAL = nil then Exit;
  if ((SaJal.Journal <> '') and (GS.RowCount <= 2) and (GS.CanFocus) and
    (Screen.ActiveControl <> GS) and (Screen.ActiveControl <> Valide97)) then
  begin
    if OkPourLigne then GS.SetFocus;
    Exit;
  end;
  {Contrôles avant validation}
  if Action = taConsult then
  begin
    CloseFen;
    Exit;
  end;
  if not BValide.Enabled then Exit;

  {YMO FQ17246 17&25/01/06 Déclenchement des verif de modification sur les 2 lignes
  et lancement de la procédure MajDirecteTable qui fait l'update}
{  GS.BeginUpdate ;
  BeginTrans ;
  try
    try
      For i :=1 to GS.Rowcount-2 do
        begin
        if GS.Cells[SA_Gen, i] = SAJAL.TRESO.CPT then continue;//on ne prend pas les lignes de contrepartie
        if not LigneCorrecte(i, False, True, TRUE) then break ;
        if SAJAL.TRESO.TypCtr = Ligne then GereLigneTreso(i);
        MajDirecteTable(i, '');
        if SAJAL.TRESO.TypCtr = Ligne then MajDirecteTable(i + 1, '');
        end;
      CommitTrans;
    except
     On E : Exception do
      begin
       PGIInfo('Erreur' +#10#13+ E.Message ) ;
       Rollback ;
       exit ;
      end ;
    end ;
    finally
    GS.EndUpdate ;
  end ;
}
  // Finition de la ligne en cours de saisie
  ValideCeQuePeux(GS.Row);

  // En contreparite auto en solde, saisie au pied
  if (SAJal.Treso.TypCtr in [PiedDC, PiedSolde]) then
  begin
    if EstRempli(GS, GS.Row) then
    begin
      if not LigneCorrecte(GS.Row, TRUE, FALSE, TRUE) then Exit;
    end else
    begin
      OBM := GetO(GS, GS.Row);
      if (OBM <> nil) and (OBM.Status = DansTable) then Exit;
    end;
  end;

  // Vérification de la pièce
  if not PieceCorrecte(TRUE) then Exit; // ok mess = true >> affiche erreur sur contrepartie manuelle ...!?!?

  // Vérification de la ligne en cours de saisie
  if ((EstRempli(GS, GS.Row)) and
    (not LigneCorrecte(GS.Row, False, True, FALSE))) then Exit;

  // Mise en place de la ligne de contrepartie si paramétré en pied de pièce
  if SAJAL <> nil then if (SAJAL.TRESO.TypCtr in [PiedDC, PiedSolde]) then
    if not MajGridPiedTreso then exit;

  // Vérification de la pièce
  if not PieceCorrecte(FALSE) then Exit; // ok mess = false >> n'affiche pas les erreurs sur contrepartie manuelle ...!?!?

  MajIOComTreso;
  if GS.RowCount <> NbLOrig then Revision := False;

  GS.Row := 1;
  GS.Col := SA_Gen;
  GS.SetFocus;
  BValide.Enabled := False;
  SetTypeExo ; // FQ 16852 : SBO 30/03/2005
  AttribRegimeEtTVA;
  {Scenario avant validation}
  {Validation}
  DatePiece := StrToDate(E_DATECOMPTABLE.Text);
  if not M.MajDirecte then NowFutur := NowH;
  GS.Enabled := False;
  io  := Transactions(ValideLaPiece, 10);
  io1 := Transactions(ValideLeReste, 10);
  if io1 <> oeOK then MessageAlerte(HDiv.Mess[27] + HDiv.Mess[28]);
  GridEna(GS, True);
  GS.Enabled := True;
  case io of
    oeOK: StockeLaPiece;
    oeSaisie: MessageAlerte(HDiv.Mess[4]);
    oeUnknown: MessageAlerte(HDiv.Mess[3]);
  end;

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

  {Ré-initialisation}
  PieceModifiee := False;
  if (Action = taCreat) and (not Agio) then ReInitPiece else
  begin
    if Agio then
    begin
      M.Jal := E_JOURNAL.Value;
      M.DateC := StrToDate(E_DATECOMPTABLE.Text);
      M.Exo := QuelExoDT(StrToDate(E_DATECOMPTABLE.Text));
      M.Num := NumPieceInt;
      M.Simul := 'N';
      M.Nature := E_NATUREPIECE.Value;
    end;
    CloseFen;
  end;

  AffichePied; // YMO 27/12/2005 FQ 10420 Maj du solde de contrepartie après validation

  {JP 17/01/08 : FQ 18563 : Calcul du solde en fonction du compte et surtout de l'exercice}
  SAJAL.PutDate(E_DATECOMPTABLE.Text);
  //YMO 26/09/2006 FQ10420 Reinitialisation de l'enreg SAJAL et réaffichage
  SAJAL.ChargeCompteTreso ; // YMO
  CalCulSoldeCompteT; // YMO

end;

procedure TFSaisietr.ClickVentil;
var
  OkL: Boolean;
begin
  if GS.Row >= GS.RowCount - 1 then Exit;
  if not EstRempli(GS, GS.Row) then Exit;
  OkL := LigneCorrecte(GS.Row, False, True, TRUE);
  if OkL then GereAnal(GS.Row, False, False);
end;

procedure TFSaisietr.ClickEche;
var
  OkL: Boolean;
begin
  if GS.Row >= GS.RowCount - 1 then Exit;
  if not EstRempli(GS, GS.Row) then Exit;
  OkL := LigneCorrecte(GS.Row, False, True, TRUE);
  if OkL then
  begin
    GereEche(GS.Row, False, False);
    if {M.Treso and} M.MajDirecte then
    begin
      BeginTrans;
      MajDirecteTable(GS.Row, '');
      if SAJAL.TRESO.TypCtr = Ligne then
      begin
        AlimLigneTreso(GS.Row);
        MajDirecteTable(GS.Row + 1, '');
      end;
      CommitTrans;
    end;
    Revision := False;
  end;
end;

procedure TFSaisietr.ClickAbandon;
begin
  if ModeForce = tmPivot then ClickSwapPivot else
  begin
    //    if ((Action<>taConsult) and (PieceModifiee)) then BEGIN HPiece.Execute(4,caption,'') ; Exit ; END ;
    CloseFen;
  end;
end;

procedure TFSaisietr.CloseFen;
begin
  Close;
  {JP 05/07/07 : FQ 19022 : Si on a répondu non à la confirmation fermeture, on évite de vider le panel !}
  if FNeFermePlus then Exit;

  if FClosing and IsInside(Self) then
    begin
    CloseInsidePanel(Self) ;
    THPanel(Self.parent).InsideForm := nil;
    THPanel(Self.parent).VideToolBar;
//    THPanel(parent).CloseInside;
    end;
end;

procedure TFSaisietr.ClickSolde(Egal: boolean);
begin
  if {M.Treso and} (not BSolde.Enabled) then Exit;
  if PasModif then Exit;
  if not ExJaiLeDroitConcept(TConcept(ccSaisSolde), True) then Exit;
  if not EstRempli(GS, GS.Row) then Exit;
  {if not M.Treso then if Mvt pointé ou lettré then Exit (à rajouter si future compatibilité)}
  if ((Egal) and (GS.Col <> SA_Gen) and (GS.Col <> SA_Aux) and (GS.Col <> SA_Debit) and (GS.Col <> SA_Credit)) then Exit;
  GereNewLigne1;
  if ((GS.Col = SA_Debit) or (GS.Col = SA_Credit)) then
  begin
    FormatMontant(GS, SA_Debit, GS.Row, DECDEV);
    FormatMontant(GS, SA_Credit, GS.Row, DECDEV);
    TraiteMontant(GS.Row, True);
  end;
  SoldeLaLigne(GS.Row);
  //if not M.Treso then BValide.Enabled := ((H_NumeroPiece.Visible) and (TRUE));
end;


procedure TFSaisietr.ClickZoom(DoPopup: boolean);
var
  b    : byte;
  C, R : longint;
  A    : TActionFiche ;
begin
  if ((E_JOURNAL.Value = '') or (not GS.Enabled)) then Exit;
  if SAJAL = nil then Exit;
  if ((Action = taConsult) or (ModeForce <> tmNormal)) then
  begin
    GS.MouseToCell(GX, GY, C, R);
    A := taConsult;
  end
  else
  begin
    R := GS.Row;
    C := GS.Col;
    A := taModif;
  end;
  if R < 1 then Exit;
  if C < SA_Gen then Exit;
  if ((Action = taConsult) and (GS.Cells[C, R] = '')) then Exit;
  if ModeForce <> tmNormal then Exit;
  if C = SA_Gen then
  begin
    if not ExJaiLeDroitConcept(TConcept(ccGenModif), False) then A := taConsult;
    b := ChercheGen(C, R, False, DoPopup);
    if b = 2 then
    begin
      FicheGene(nil, '', GS.Cells[C, R], A, 0);
      if Action <> taConsult then ChercheGen(C, R, True);
    end;
  end else if C = SA_Aux then
  begin
    if not ExJaiLeDroitConcept(TConcept(ccAuxModif), False) then A := taConsult;
    b := ChercheAux(C, R, False);
    if b = 2 then
    begin
      FicheTiers(nil, '', GS.Cells[C, R], A, 1);
      if Action <> taConsult then ChercheAux(C, R, True);
    end;
  end;
end;

procedure TFSaisietr.ClickSwapPivot;
var
  Cancel, Chg: boolean;
begin
  if ((not ModeDevise) and (ModeForce = tmNormal)) then Exit;
  if not BSwapPivot.Enabled then Exit;
  if ((EstRempli(GS, GS.Row)) and (not LigneCorrecte(GS.Row, False, True, TRUE))) then Exit;
  if not EstRempli(GS, 1) then Exit;
  if ModeForce = tmNormal then
  begin
    ModeForce := tmPivot;
    OldDevise := DEV.Code;
    E_DEVISE.Value := V_PGI.DevisePivot;
    DEV.Decimale := V_PGI.OkDecV;
    DECDEV := V_PGI.OkDecV;
    E_DEVISE.Enabled := True;
    E_Devise.Font.Color := clRed;
    AffecteGrid(GS, taConsult);
    GS.SetFocus;
    PEntete.Enabled := False;
    Outils.Enabled := False;
    Valide97.Enabled := False;
  end else
  begin
    E_DEVISE.Enabled := False;
    E_DEVISE.Value := OldDevise;
    E_Devise.Font.Color := clBlack;
    ModeForce := tmNormal;
    PEntete.Enabled := True;
    Outils.Enabled := True;
    Valide97.Enabled := True;
    if Outils.CanFocus then Outils.SetFocus;
    GS.Col := SA_Gen;
    GS.Row := 1;
    GS.CacheEdit;
    AffecteGrid(GS, Action);
    GS.Col := SA_Gen;
    GS.Row := 2;
    GS.SetFocus;
    GS.MontreEdit;
    GS.Row := 1;
    Cancel := False;
    Chg := False;
    GSRowEnter(nil, GS.Row, Cancel, Chg);
  end;
  ChangeAffGrid(GS, ModeForce, DECDEV);
  CalculDebitCredit;
end;

procedure TFSaisietr.ClickModifRIB;
var
  O: TOBM;
  RibAvant: string;
  RibApres: string;
  IsAux: Boolean;
begin
  if not BModifRIB.Enabled then Exit;
  O := GetO(GS, GS.Row);
  if O = nil then Exit;
  RibAvant := O.GetMvt('E_RIB');
  IsAux := O.GetMvt('E_AUXILIAIRE') <> '';
  ModifRIBOBM(O, False, TRUE, GS.Cells[SA_Aux, GS.Row], IsAux);
  RibApres := O.GetMvt('E_RIB') ;
  if (RibAvant <> RibApres) and M.MajDirecte then
    begin
    if O.GetMvt('E_RIB') <> ''
      then E_RIB.Caption := HDiv.Mess[14] + ' ' + O.GetMvt('E_RIB')
      else E_RIB.Caption := '';
    // MAJ en base : ligne courante
    if O.Status <> EnMemoire then
      MajDirecteTable(GS.Row, '');
    // MAJ objet + base pour ligne de contrepartie si à la ligne et déjà générée
    if (SAJAL.TRESO.TypCtr = Ligne) and EstRempli( GS, GS.Row+1 ) then
      begin
      O := GetO(GS, GS.Row+1);
      if O <> nil then
        begin
        O.PutMvt('E_RIB', RibApres ) ;
        if O.Status <> EnMemoire then
          MajDirecteTable(GS.Row+1, '');
        end ;
      end ;
    end;
end;

procedure TFSaisietr.ClickComplement;
var
  ModBN: boolean;
  O: TOBM;
  Lig: integer;
  RC: R_COMP;
  OQ1, OQ2, NQ1, NQ2: Double;
begin
  Lig := GS.Row;
  if not EstRempli(GS, Lig) then Exit;
  if not LigneCorrecte(Lig, False, True, TRUE) then Exit;
  O := GetO(GS, Lig);
  OQ1 := O.GetMvt('E_QTE1');
  OQ2 := O.GetMvt('E_QTE2');
  RC.StComporte := '--XXXXXXXX';
  RC.StLibre := 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX';
  RC.Conso := True;
  RC.Attributs := False;
  RC.DateC := DatePiece;
  RC.MemoComp := nil;
  RC.Origine := -1;
  {08/06/06 : FQ 17556 : TOBCompl sert pour le CutOff, donc les comptes de charges / produits.
              Cela a donc peu de raison d'être ici}
  RC.TOBCompl := nil ;
  if not SaisieComplement(O, EcrGen, Action, ModBN, RC) then Exit;
  if GS.CanFocus then GS.SetFocus;
  if MODBN then Revision := False;
  NQ1 := O.GetMvt('E_QTE1');
  NQ2 := O.GetMvt('E_QTE2');
  if ((NQ1 <> OQ1) and (OQ1 <> 0)) then ProrateQteAnal(GS, Lig, OQ1, NQ1, 1);
  if ((NQ2 <> OQ2) and (OQ2 <> 0)) then ProrateQteAnal(GS, Lig, OQ2, NQ2, 2);
end;

procedure TFSaisietr.ClickDelettrage;
var
  Lig: integer;
begin
  Lig := GS.Row;
  if not EstRempli(GS, Lig) then Exit;
  if not LigneCorrecte(Lig, False, True, TRUE) then Exit;
  DeLettrageEnSaisieTreso(Lig);
end;

procedure TFSaisietr.Clicklettrage;
var
  Lig: integer;
begin
  Lig := GS.Row;
  if not (OkScenario and (Scenario.GetMvt('SC_LETTRAGESAISIE') = 'X')) then Exit;
  if not EstRempli(GS, Lig) then Exit;
  if not LigneCorrecte(Lig, False, True, TRUE) then Exit;
  LettrageEnSaisieTreso(Lig, TRUE);
end;

procedure TFSaisietr.AlimOBMMemoireEnBatch(OSource, ODestination: TOBM);
begin
  if (OSource = nil) or (ODestination = nil) then Exit;
  ODestination.PutMvt('E_AFFAIRE', OSource.GetMvt('E_AFFAIRE'));
  ODestination.PutMvt('E_REFEXTERNE', OSource.GetMvt('E_REFEXTERNE'));
  ODestination.PutMvt('E_DATEREFEXTERNE', OSource.GetMvt('E_DATEREFEXTERNE'));
  ODestination.PutMvt('E_REFLIBRE', OSource.GetMvt('E_REFLIBRE'));
  ODestination.PutMvt('E_QTE1', OSource.GetMvt('E_QTE1'));
  ODestination.PutMvt('E_QTE2', OSource.GetMvt('E_QTE2'));
  ODestination.PutMvt('E_QUALIFQTE1', OSource.GetMvt('E_QUALIFQTE1'));
  ODestination.PutMvt('E_QUALIFQTE2', OSource.GetMvt('E_QUALIFQTE2'));
  ODestination.M.Assign(OSource.M);
end;

procedure TFSaisietr.ClickComplementT;
var
  LesDeux, Inutile: boolean;
  O: TOBM;
  RC: R_COMP;
begin
  if ((E_JOURNAL.Value = '') or (SAJAL = nil)) then Exit;
  LesDeux := (OBMT[1] <> nil) and (OBMT[2] <> nil);
  O := nil;
  if OBMT[1] <> nil then O := OBMT[1];
  if OBMT[2] <> nil then O := OBMT[2];
  if O = nil then Exit;
  O.PutMvt('E_NUMLIGNE', GS.RowCount - 1);
  RC.StComporte := '--XXXXXXXX';
  RC.StLibre := 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX';
  RC.Conso := True;
  RC.Attributs := False;
  RC.MemoComp := nil;
  RC.Origine := -1;
  RC.DateC := DatePiece;
  RC.TOBCompl := nil ;
  if not SaisieComplement(O, EcrGen, Action, Inutile, RC) then Exit;
  if GS.CanFocus then GS.SetFocus;
  if LesDeux then AlimOBMMemoireEnBatch(OBMT[2], OBMT[1]);
end;

procedure TFSaisietr.ClickModifTva;
var
  Lig: integer;
  O: TOBM;
  OMODR: TMOD;
  Client: Boolean;
  CAux: TGTiers;
  CGen: TGGeneral;
  AA: TActionFiche;
  RTVA: Enr_Base;
begin
  {#TVAENC}
  if not VH^.OuiTvaEnc then Exit;
  if not BModifTva.Enabled then Exit;
  Lig := GS.Row;
  O := GetO(GS, Lig);
  if O = nil then Exit;
  AA := Action;
  // GP TvaEnc if PasToucheLigne(Lig) then AA:=taConsult ;
  if IsHT(GS, Lig, True) then
  begin
    if O.GetMvt('E_REGIMETVA') = '' then O.PutMvt('E_REGIMETVA', GeneRegTva);
    if (VH^.PaysLocalisation=CodeISOES) and
       (InfosTvaEnc(O,AA)) and (IsTiersSoumisTPF(GS)) then
      O.PutMvt('E_TPF',O.GetValue('E_TVA'))
    else
      InfosTvaEnc(O,AA) ;
    if M.MajDirecte and StatusOBMOk(O) then MajDirecteTable(Lig, '');
  end else if IsTiers(GS, Lig) then
  begin
    if SorteTva <> stvDivers then AA := taConsult;
    if GS.Cells[SA_Aux, Lig] <> '' then
    begin
      CAux := GetGTiers(GS, Lig);
      if CAux = nil then Exit;
      if ((CAux.NatureAux = 'CLI') or (CAux.NatureAux = 'AUD')) then Client := True else
        if ((CAux.NatureAux = 'FOU') or (CAux.NatureAux = 'AUC')) then Client := False else Exit;
    end else if GS.Cells[SA_Gen, Lig] <> '' then
    begin
      CGen := GetGGeneral(GS, Lig);
      if CGen = nil then Exit;
      if CGen.NatureGene = 'TID' then Client := True else
        if CGen.NatureGene = 'TIC' then Client := False else Exit;
    end else Exit;
    CGen := GetGGeneral(GS, Lig);
    if CGen = nil then Exit;
    if not EstCollFact(CGen.General) then Exit;
    OMODR := TMOD(GS.Objects[SA_NumP, Lig]);
    if OMODR = nil then Exit;
    RTVA.Regime := O.GetMvt('E_REGIMETVA');
    RTVA.Nature := E_NATUREPIECE.Value;
    RTVA.Client := Client;
    RTVA.Action := AA;
    RTVA.CodeTva := O.GetMvt('E_TVA');
    if SaisieBasesHT(OMODR, RTVA) and M.MajDirecte then
    begin
      ModifOBMPourTvaEnc(Lig);
      if StatusOBMOk(O) then MajDirecteTable(Lig, '');
    end;
  end else Exit;
  Revision := False;
end;

procedure TFSaisietr.ClickCherche;
begin
  if GS.RowCount <= 2 then Exit;
  FindFirst := True;
  FindSais.Execute;
end;

procedure TFSaisietr.ClickRAZLigne;
var
  Lig: longint;
  OBM: TOBM;
begin
  if ((E_JOURNAL.Value = '') or (not GS.Enabled)) then Exit;
  if SAJAL = nil then Exit;
  if ((Action <> taCreat) or (ModeForce <> tmNormal)) then Exit;
  Lig := GS.Row;
  if Lig <= 1 then Exit;
  if Self.M.MAJDirecte then
  begin
    OBM := GetO(GS, Lig);
    if OBM = nil then Exit else
      if OBM.Status <> EnMemoire then
        // Si non lettré, essai de suppression base...
        if OBM.GetString('E_LETTRAGE') <> '' then
          begin
          PGIInfo('Cette ligne est lettrée, elle ne peut plus être effacée') ;
          Exit ;
          end
        else
          begin
          // supression ligne saisie
          if not SupprimeLigneEnBase( GS.Row ) then
            Exit ;
          // supression ligne tréso si contrepartie à la ligne
          if SAJAL.TRESO.TypCtr = Ligne then
            if not SupprimeLigneEnBase( GS.Row + 1 ) then
              begin
              rollback;
              Exit ;
              end ;
          end ;
  end;
  GS.SynEnabled := False;
  DetruitLigne(GS.Row, True);
  Revision := False;
  ControleLignes;
  GS.SynEnabled := True;
end;

procedure TFSaisietr.ClickChancel;
var
  DD: TDateTime;
  AA: TActionFiche;
begin
  if DEV.Code = V_PGI.DevisePivot then Exit;
  if Action <> taCreat then Exit;
  if not BChancel.Enabled then Exit;
  DD := StrToDate(E_DATECOMPTABLE.Text);
  AA := taModif;
  if Action = taConsult then AA := taConsult;
  FicheChancel(DEV.Code, True, DD, AA, (DD >= V_PGI.DateDebutEuro));
  if AA <> taConsult then DEV.Taux := GetTaux(DEV.Code, DEV.DateTaux, DD);
end;

procedure TFSaisietr.ClickScenario;
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
end;

procedure TFSaisietr.ClickJournal;
var
  a: TActionFiche;
begin
  //if Action=taConsult then a:=taConsult else a:=taModif ;
  //if Not JaileDroitConcept(TConcept(ccJalModif),False) then a:=taConsult ;
  if ((E_JOURNAL.Value = '') or (SAJAL = nil)) then Exit;
  a := taConsult;
  FicheJournal(nil, '', E_JOURNAL.Value, a, 0);

  if a <> taConsult then ChercheJal(E_JOURNAL.Value, True);
end;

procedure TFSaisietr.ClickDevise;
begin
  FicheDevise(E_DEVISE.Value, taModif, False);
end;

procedure TFSaisieTr.ClickSaisTaux;
begin
  if not BSaisTaux.Enabled then Exit;
  if DEV.Code = V_PGI.DevisePivot then Exit;
  if DEV.Code = '' then Exit;
  if SaisieNewTaux2000(DEV, DatePiece) then
  begin
    M.TauxD := DEV.Taux;
    Volatile := True;
  end;
end;

procedure TFSaisieTr.ClickChoixRegime;
var
  NewReg, Titre: string;
  Lig: integer;
  O: TOBM;
begin
  if not BChoixRegime.Enabled then Exit;
  if Action = taConsult then Exit;
  if ModeForce <> tmNormal then Exit;
  Titre := HDiv.Mess[26] + ' ' + RechDom('TTREGIMETVA', GeneRegTVA, False) + ')';
  if GeneRegTva = '' then GeneRegTVA := VH^.RegimeDefaut;
  NewReg := Choisir(Titre, 'CHOIXCOD', 'CC_LIBELLE', 'CC_CODE', 'CC_TYPE="RTV"', '');
  if ((NewReg = '') or (NewReg = GeneRegTVA)) then Exit;
  if HPiece.Execute(32, Caption, '') <> mrYes then Exit;
  GeneRegTVA := NewReg;
  RegimeForce := True;
  for Lig := 1 to GS.RowCount - 2 do
  begin
    O := GetO(GS, Lig);
    if O = nil then Break;
    O.PutMvt('E_REGIMETVA', GeneRegTVA);
  end;
end;

procedure TFSaisietr.ClickEtabl;
begin
  FicheEtablissement_AGL(taConsult);
end;

procedure TFSaisietr.BVentilClick(Sender: TObject);
begin
  ClickVentil;
  if GS.Enabled then GS.SetFocus;
end;

procedure TFSaisietr.BEcheClick(Sender: TObject);
begin
  ClickEche;
  if GS.Enabled then GS.SetFocus;
end;

procedure TFSaisietr.BLigClick(Sender: TObject);
begin
  if GS.RowCount <= 2 then Exit;
  AfficheLigneTreso;
  GS.SetFocus;
end;

procedure TFSaisietr.BValideClick(Sender: TObject);
begin
  if ((M.TypeGuide <> 'NOR') or (not M.FromGuide)) then ClickValide else ClickAbandon;
end;

procedure TFSaisietr.BZoomDeviseClick(Sender: TObject);
begin
  ClickDevise;
end;

procedure TFSaisietr.BZoomEtablClick(Sender: TObject);
begin
  ClickEtabl;
end;

procedure TFSaisietr.BSoldeClick(Sender: TObject);
begin
  ClickSolde(False);
  if GS.Enabled then GS.SetFocus;
end;

procedure TFSaisietr.BAbandonClick(Sender: TObject);
begin
  ClickAbandon;
end;

procedure TFSaisietr.BZoomClick(Sender: TObject);
begin
  ClickZoom;
  if GS.Enabled then GS.SetFocus;
end;

procedure TFSaisietr.BScenarioClick(Sender: TObject);
begin
  ClickScenario;
end;

procedure TFSaisietr.BZoomJournalClick(Sender: TObject);
begin
  ClickJournal;
end;

procedure TFSaisietr.BSwapPivotClick(Sender: TObject);
begin
  ClickSwapPivot;
end;

procedure TFSaisietr.BComplementClick(Sender: TObject);
begin
  ClickComplement;
end;

procedure TFSaisietr.BModifTvaClick(Sender: TObject);
begin
  ClickModifTva;
end;

procedure TFSaisietr.BModifRIBClick(Sender: TObject);
begin
  ClickModifRIB;
end;

procedure TFSaisietr.BChercherClick(Sender: TObject);
begin
  ClickCherche;
end;

procedure TFSaisietr.BChancelClick(Sender: TObject);
begin
  ClickChancel;
end;

procedure TFSaisietr.BRefClick(Sender: TObject);
begin
  CalculRefAuto(Sender);
  if GS.Enabled then GS.SetFocus;
end;

procedure TFSaisietr.ValideCeQuePeux(Lig: longint);
var
  Cpte: string;
  OBM: TOBM;
begin
  OBM := GetO(GS, Lig);
  if OBM = nil then Exit;
  if not EstRempli( GS, Lig ) then
    begin
    DetruitLigne( Lig, True ) ;
    Exit ;
    end ;
  if ((GS.Cells[SA_Gen, Lig] <> '') and (GS.Objects[SA_Gen, Lig] = nil)) then
  begin
    Cpte := uppercase(GS.Cells[SA_Gen, Lig]);
    if Presence('GENERAUX', 'G_GENERAL', BourreLaDonc(Cpte, fbGene)) then
    begin
      ChercheGen(SA_Gen, Lig, False);
      OBM.PutMvt('E_GENERAL', GS.Cells[SA_Gen, Lig]);
    end;
  end;
  if ((GS.Cells[SA_Aux, Lig] <> '') and (GS.Objects[SA_Aux, Lig] = nil)) then
  begin
    Cpte := uppercase(GS.Cells[SA_Aux, Lig]);
    if Presence('TIERS', 'T_AUXILIAIRE', BourreLaDonc(Cpte, fbAux)) then
    begin
      ChercheAux(SA_Aux, Lig, False);
      OBM.PutMvt('E_AUXILIAIRE', GS.Cells[SA_Aux, Lig]);
    end;
  end;
  OBM.PutMvt('E_REFINTERNE', Copy(GS.Cells[SA_RefI, Lig], 1, 35));
  OBM.PutMvt('E_LIBELLE', Copy(GS.Cells[SA_Lib, Lig], 1, 35));
  if GS.Cells[SA_Debit, Lig] <> '' then ChercheMontant(SA_Debit, Lig);
  if GS.Cells[SA_Credit, Lig] <> '' then ChercheMontant(SA_Credit, Lig);
end;

{==========================================================================================}
{============================ Objet MOUVEMENT, Divers =====================================}
{==========================================================================================}
procedure TFSaisietr.NumeroteVentils;
{##ANA##
var
  OBA: TObjAna;
  NumL, NumAxe, NumV: integer;
  T: TList;
begin
  for NumL := 1 to GS.RowCount - 2 do
  begin
    OBA := TObjAna(GS.Objects[SA_DateC, NumL]);
    if OBA = nil then Continue;
    OBA.NumLigne := NumL;
    for NumAxe := 1 to MaxAxe do if OBA.AA[NumAxe] <> nil then
      begin
        T := OBA.AA[NumAxe].L;
        for NumV := 0 to T.Count - 1 do PutMvtA(OBA.AA[NumAxe], P_TV(T.Items[NumV]).F, 'Y_NUMLIGNE', NumL);
      end;
  end;
}
var
  OBM : TOBM;
  OBA : TOB;
  NumL,
  NumAxe,
  NumV : Integer;
begin
  for NumL := 1 to GS.RowCount - 2 do begin
    OBM := GetO(GS,NumL) ;
    if OBM = nil then Continue;

    for NumAxe := 0 to OBM.Detail.Count - 1 do
      if OBM.Detail[NumAxe].Detail.Count>0 then begin
        for NumV := 0 to OBM.Detail[NumAxe].Detail.Count - 1 do begin
          OBA := OBM.Detail[NumAxe].Detail[NumV];
          OBA.PutValue('Y_NUMLIGNE',NumL) ;
        end;
      end ;
  end ;
{##ANA##}
end;

procedure TFSaisietr.GotoEntete;
var
  i: integer;
  T: TWinControl;
begin
  if PasModif then Exit;
  if not PEntete.Enabled then Exit;
  for i := 0 to PEntete.ControlCount - 1 do
  begin
    T := TWinControl(PEntete.Controls[i]);
    if ((T.CanFocus) and (Copy(T.Name, 1, 2) = 'E_') and (T.ClassType <> THLabel)) then
    begin
      T.SetFocus;
      Break;
    end;
  end;
end;

procedure TFSaisietr.GereComplements(Lig: integer);
var
  O: TOBM;
  CGen: TGGeneral;
  ModBN: boolean;
  NumRad: integer;
  RC: R_COMP;
  StComp, StLibre: string;
begin
  if PasModif then Exit;
  O := GetO(GS, Lig);
  if O = nil then Exit;
  if O.CompS then Exit;
  CGen := GetGGeneral(GS, Lig);
  if CGen = nil then Exit;
  NumRad := ComporteLigne(Scenario, CGen.General, StComp, StLibre);
  if NumRad <= 0 then Exit;
  if ((Pos('X', StComp) <= 0) and (Pos('X', StLibre) <= 0)) then Exit;
  RC.StComporte := StComp;
  RC.StLibre := StLibre;
  RC.Conso := False;
  RC.DateC := DatePiece;
  RC.Attributs := True;
  RC.MemoComp := MemoComp;
  RC.Origine := NumRad;
  O.CompS := True;
  {08/06/06 : FQ 17556 : TOBCompl sert pour le CutOff, donc les comptes de charges / produits.
              Cela a donc peu de raison d'être ici}
  RC.TOBCompl := nil ;
  if not SaisieComplement(O, EcrGen, Action, ModBN, RC) then Exit;
  if GS.CanFocus then GS.SetFocus;
  if ModBN then Revision := False;
end;


function TFSaisietr.LettrageFaitEnSaisieTreso: Boolean;
var
  Lig: Integer;
  OBM: TOBM;
  M: TMod;
  R: T_ModeRegl;
begin
  Result := FALSE;
  for Lig := 1 to GS.RowCount - 2 do
  begin
    OBM := GetO(GS, Lig);
    if (OBM <> nil) then
    begin
      if (OBM.GetMvt('E_TRESOLETTRE') = 'X') then
      begin
        M := TMOD(GS.Objects[SA_NumP, Lig]);
        if M <> nil then
        begin
          R := M.MODR;
          if (R.TabEche[1].EtatLettrage = 'PL') or (R.TabEche[1].EtatLettrage = 'TL') then
          begin
            Result := TRUE;
            Break;
          end;
        end;
      end;
    end;
  end;
end;

procedure TFSaisietr.DecrementeNumero;
var
  Facturier: String3;
  DD: TDateTime;
begin
  //if M.Treso then
    Facturier := SAJAL.CompteurNormal;
  {else
  begin
    if M.Simul <> 'N' then Facturier := SAJAL.CompteurSimul else Facturier := SAJAL.CompteurNormal;
  end;}
  DD := StrToDate(E_DATECOMPTABLE.Text);
  SetDecNum(EcrGen, Facturier, DD);
end;

function TFSaisietr.OnPeutDetruire: Boolean;
var
  ii, jj, Lig: Integer;
  OBM: TOBM;
  sSQL: string;
  Q: TQuery;
begin
  sSQL := 'Select Count(E_numligne) FROM ECRITURE WHERE ';
  sSQL := sSQL + 'E_JOURNAL="' + SaJal.Journal + '" AND E_EXERCICE="' + QuelExo(E_DateComptable.EditText) + '"';
  sSQL := sSQL + 'AND E_DATECOMPTABLE="' + UsDateTime(StrToDate(E_DATECOMPTABLE.TEXT)) + '" AND E_NUMEROPIECE=' + IntToStr(NumPieceInt);

  Q := OpenSQL(sSQL, true);
  ii := 0;
  if not Q.eof then ii := Q.Fields[0].Asinteger;
  Ferme(Q);
  jj := 0;
  for Lig := 1 to GS.RowCount - 2 do
  begin
    OBM := GetO(GS, Lig);
    if (OBM <> nil) then inc(jj);
  end;
  Result := ii = jj;

end;

procedure TFSaisietr.DetruirePieceTresoEnCours;
var
  sSQL: string;
begin
  DecrementeNumero;
  sSQL := 'DELETE FROM ECRITURE WHERE ';
  sSQL := sSQL + 'E_JOURNAL="' + SaJal.Journal + '" AND E_EXERCICE="' + QuelExo(E_DateComptable.EditText) + '"';
  sSQL := sSQL + 'AND E_DATECOMPTABLE="' + UsDateTime(StrToDate(E_DATECOMPTABLE.TEXT)) + '" AND E_NUMEROPIECE=' + IntToStr(NumPieceInt);
  ExecuteSQL(sSQL);
end;

function TFSaisietr.FermerSaisie: boolean;
var
  i: integer;
  Okok: boolean;
begin
  {JP 05/07/07 : FQ 19022 : Si on répond non à la confirmation fermeture}
  FNeFermePlus := False;

  Result := True;
  if ((Action <> taConsult) and (ModeForce <> tmNormal)) then
  begin
    Result := False;
    Exit;
  end;
  if ((NeedEdition) and (Action = taCreat) and (TPIECE.Count > 0)) then EditionSaisie;
  if ((Action = taConsult) or (not PieceModifiee)) then Exit;
  if ((Action = taCreat) and (M.TypeGuide = 'NOR') and (M.FromGuide)) then Exit;
  if Action = taModif then
  begin
    if HPiece.Execute(3, caption, '') <> mrYes then begin
      {JP 05/07/07 : FQ 19022 : Si on répond non à la confirmation fermeture}
      FNeFermePlus := True;

      Result:=False ;
    end;
  end else
  begin
    (* GP le 08/07/97 N° 1835
    if PossibleRecupNum then
       BEGIN
       Okok:=True ; i:=HPiece.Execute(3,caption,'') ;
       if i<>mrYes then BEGIN Result:=False ; Exit ; END ;
       END else Okok:=False ;
    *)
    (* GP 17/11/97 : Re-modif : il faut pouvoir détruire une pièce transitoire de tréso non lettrées *)
    if PossibleRecupNum and (not LettrageFaitEnSaisieTreso) then
    begin
      Okok := True;
      i := HPiece.Execute(3, caption, '');
      if i <> mrYes then
      begin
        {JP 05/07/07 : FQ 19022 : Si on répond non à la confirmation fermeture}
        FNeFermePlus := True;
        Result := False;
        Exit;
      end;
    end else Okok := False;
    //   OkOk:=TRUE ;
    //   If OkOk And M.Treso And Self.M.MajDirecte And (H_NUMEROPIECE.Visible) Then OkOk:=FALSE ;
    (*
       if ((Okok) and (PossibleRecupNum)) then {Revérifier, le numéro a pu être attribué entre temps!}
          BEGIN
          if Not M.Treso And (H_NUMEROPIECE_.Visible=False) then
             BEGIN
             if Transactions(DecrementeNumero,10)<>oeOk then MessageAlerte(HDiv.Mess[6]) Else
                BEGIN
                END ;
             END ;
          END else
          BEGIN
          HPiece.Execute(4,caption,'') ; Result:=False ;
          END ;
    *)
    if ((Okok) and (PossibleRecupNum)) then {Revérifier, le numéro a pu être attribué entre temps!}
    begin
      if (H_NUMEROPIECE_.Visible = False) then
      begin
        if OnPeutDetruire then
        begin
          //repeat until (not DBSOC.inTransaction);
          if Transactions(DetruirePieceTresoEnCours, 10) <> oeOk then
          begin
            Result := FALSE;
            MessageAlerte(HDiv.Mess[6]);
          end;
        end else
        begin
          MessageAlerte(HDiv.Mess[12]);
          Result := FALSE;
        end;
      end;
    end else
    begin
      HPiece.Execute(4, caption, '');
      Result := False;
    end;
  end;
end;

procedure TFSaisietr.AlimObjetMvt(Lig: integer; AvecM: boolean; Treso: Byte);
var
  OBM: TOBM;
  XD, XC: Double;
  CGen: TGGeneral;
  St: string;
  LaDate: TDateTime;
begin
  if PasModif then exit;
  if SAJAL = nil then Exit;
  //if M.Treso then //TR
    case Treso of
      0:
        begin
          OBM := GetO(GS, Lig);
          if OBM = nil then
          begin
            AlloueMvt(GS, EcrGen, Lig, True);
            OBM := GetO(GS, Lig);
          end;
          OBM.PutMvt('E_GENERAL', GS.Cells[SA_Gen, Lig]);
          OBM.PutMvt('E_AUXILIAIRE', GS.Cells[SA_Aux, Lig]);
          OBM.PutMvt('E_EXERCICE', QuelExoDT(StrToDate((E_Datecomptable.Text))));
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
        end;
      1, 2:
        begin
          if OBMT[Treso] = nil then OBMT[Treso] := TOBM.Create(EcrGen, '', TRUE);
          OBM := OBMT[Treso];
          OBM.PutMvt('E_GENERAL', SAJAL.TRESO.CPT);
          OBM.PutMvt('E_AUXILIAIRE', '');
          OBM.PutMvt('E_LIBELLE', E_LIBTRESO.TEXT);
          OBM.PutMvt('E_REFINTERNE', E_REFTRESO.TEXT);
          OBM.PutMvt('E_EXERCICE', QuelExoDT(StrToDate(E_Datecomptable.Text)));
        end else Exit;
    end;
  if Lig <= 0 then Exit;
  OBM := GetO(GS, Lig);
  if OBM = nil then
  begin
    AlloueMvt(GS, EcrGen, Lig, True);
    OBM := GetO(GS, Lig);
  end;
  LaDate := StrToDate(E_DateComptable.Text);
  //if not M.Treso then OBM.PutMvt('E_EXERCICE', QuelExoDT(LaDate));
  {if M.Treso then} OBM.PutMvt('E_VISION', SAJAL.TRESO.STypCtr);
  OBM.PutMvt('E_JOURNAL', SAJAL.JOURNAL);
  OBM.PutMvt('E_DATECOMPTABLE', LaDate);
  {$IFNDEF SPEC302}
  OBM.PutMvt('E_PERIODE', GetPeriode(LaDate));
  OBM.PutMvt('E_SEMAINE', NumSemaine(LaDate));
  {$ENDIF}
  OBM.PutMvt('E_NUMEROPIECE', NumPieceInt);
  OBM.PutMvt('E_NUMLIGNE', Lig);
  OBM.PutMvt('E_GENERAL', GS.Cells[SA_Gen, Lig]);
  OBM.PutMvt('E_AUXILIAIRE', GS.Cells[SA_Aux, Lig]);
  OBM.PutMvt('E_QUALIFPIECE', M.Simul);
  OBM.PutMvt('E_DEVISE', DEV.Code);
  OBM.PutMvt('E_TAUXDEV', DEV.Taux);
  OBM.PutMvt('E_DATETAUXDEV', DEV.DateTaux);
  OBM.PutMvt('E_NATUREPIECE', E_NATUREPIECE.Value);
  OBM.PutMvt('E_ETABLISSEMENT', E_ETABLISSEMENT.Value);
  OBM.PutMvt('E_LIBELLE', Copy(GS.Cells[SA_Lib, Lig], 1, 35));
  OBM.PutMvt('E_REFINTERNE', Copy(GS.Cells[SA_RefI, Lig], 1, 35));
  OBM.MajLesDates(Action);
  if AvecM then
  begin
    XD := ValD(GS, Lig);
    XC := ValC(GS, Lig);
    OBM.SetMontants(XD, XC, DEV, True);
  end;
  CGen := GetGGeneral(GS, Lig);
  if CGen <> nil then
  begin
    if Ventilable(CGen, 0) then OBM.PutMvt('E_ANA', 'X') else OBM.PutMvt('E_ANA', '-');
    St := OBM.GetMvt('E_QUALIFQTE1');
    if ((St = '') or (St = '...')) then OBM.PutMvt('E_QUALIFQTE1', CGen.QQ1);
    St := OBM.GetMvt('E_QUALIFQTE2');
    if ((St = '') or (St = '...')) then OBM.PutMvt('E_QUALIFQTE2', CGen.QQ2);
  end;
end;

procedure TFSaisietr.BloqueGrille(Bloc: boolean);
begin
  if Action <> taConsult then
  begin
    if Bloc then
    begin
      GridEna(GS, FALSE);
      MajEnCours := TRUE;
      GS.controlstyle := GS.controlstyle - [csClickEvents, csCaptureMouse];
    end else
    begin
      GS.controlstyle := GS.controlstyle + [csClickEvents, csCaptureMouse];
      MajEnCours := FALSE;
      GridEna(GS, TRUE);
    end;
  end;
end;

{==========================================================================================}
{================================ Methodes du HGrid =======================================}
{==========================================================================================}
procedure TFSaisietr.GSDblClick(Sender: TObject);
begin
  if ((Action <> taConsult) and (ModeForce = tmDevise)) then ForcerMode(False, VK_RETURN) else ClickZoom;
end;

procedure TFSaisietr.GSCellExit(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
begin
  if PasModif then exit;
  if GS.Row = ARow then
  begin
    if ACol = SA_Gen then
    begin
      if M.Effet and (GS.ROW = ARow) and (GS.Cells[ACol, ARow] = '') and (ARow > 1) then
      begin
        if SAJAL.Treso.TypCtr = Ligne then
        begin
          if ARow > 2 then GS.Cells[ACol, ARow] := GS.Cells[ACol, ARow - 2];
        end else GS.Cells[ACol, ARow] := GS.Cells[ACol, ARow - 1];
      end;
      if ((GS.Cells[ACol, ARow] = '') and ((GS.Cells[SA_Aux, ARow] = '') or (GS.Objects[SA_Aux, ARow] = nil))) then Exit;
      if LeMeme(GS, ACol, ARow) then Exit;
      if ChercheGen(ACol, ARow, False) <= 0 then
      begin
        Cancel := True;
        Exit;
      end;
      if not Cancel then CpteAuto := False;
    end else if ACol = SA_Aux then
    begin
      if ((GS.Cells[ACol, ARow] = '') and (GS.Col = SA_Gen)) then Exit;
      if LeMeme(GS, ACol, ARow) then Exit;
      if ChercheAux(Acol, ARow, False) <= 0 then
      begin
        Cancel := True;
        Exit;
      end;
    end
    else {if M.Treso then}
      if ACol = SA_REFI then
      begin
        if (GS.Cells[ACol, ARow] = '') then
        begin
          if GS.ROW = ARow then
          begin
            if SI_NUMREF < 0 then
            begin
              if ARow > 1 then GS.Cells[ACol, ARow] := GS.Cells[ACol, ARow - 1];
            end else GS.Cells[ACol, ARow] := E_NUMREF.Text;
          end;
        end;
      end else if ACol = SA_LIB then
      begin
        if GS.ROW = ARow then
        begin
          if GS.Cells[ACol, ARow] = '' then
          begin
            if ARow > 1 then GS.Cells[ACol, ARow] := GS.Cells[ACol, ARow - 1];
          end;
        end;
      end;
  end;
  if ((ACol = SA_Debit) or (ACol = SA_Credit)) then
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
          then GS.Cells[ACol, ARow]:='{' + GetFormatPourCalc + GS.Cells[ACol, ARow]+'}'; // FQ18371 et FQ18080
        GS.Cells[ACol, ARow] := GFormule( '{' + GS.Cells[ACol, ARow] + '}', GetFormulePourCalc, nil, 1) ;
        end ;
     end;
     FormatMontant(GS, ACol, ARow, DEV.Decimale) ;
    // FIN DEV 3216 10/04/2006 SBO
    ChercheMontant(Acol, ARow);
    end ;
end;

function TFSaisietr.IsAccessible(Ou: LongInt): Boolean;
var
  OBM: TOBM;
  EtatLig: String3;
begin
  OBM := GetO(Gs, Ou);
  Result := TRUE;
  if OBM <> nil then
  begin
    EtatLig := OBM.GetMvt('E_ETATLETTRAGE');
    if (EtatLig = 'TL') or (EtatLig = 'PL') then Result := FALSE;
  end;
  if result and FInfo.LoadCompte( GS.Cells[SA_Gen,Ou] ) then // FQ18947 : ligne modifiable même si lettrée // SBO 07/11/2007
    result := ( FInfo.GetString('G_VISAREVISION') = '-' ) ;
end;

function TFSaisietr.StatusOBMOk(OBM: TOBM): Boolean;
begin
  Result := FALSE;
  if OBM = nil then Exit;
  Result := OBM.Status <> EnMemoire;
end;

procedure TFSaisietr.BasculeGS(AccesConsult: Boolean);
begin
  if ((ModeForce = tmNormal)) then
  begin
    if AccesConsult then
    begin
      GS.CacheEdit;
      GS.Options := GS.Options - [GoEditing, GoTabs, GoAlwaysShowEditor];
      GS.Options := GS.Options + [GoRowSelect];
      // SBO 19/08/2004 FQ 13290 Saisie interdite sur les comptes fermés
      G_LIBELLE.Font.Color := clRed;
      //      T_LIBELLE.Font.Color:=clRed ;
      G_LIBELLE.Caption := HDiv.Mess[7];
      //      T_LIBELLE.Caption:=' ( mouvement lettré, compte ou journal fermé ).
            // Fin SBO 19/08/2004 FQ 13290 Saisie interdite sur les comptes fermés
    end else
    begin
      GS.Options := GS.Options - [GoRowSelect];
      GS.Options := GS.Options + [GoEditing, GoTabs, GoAlwaysShowEditor];
      GS.MontreEdit;
    end;
  end;
end;

function TFSaisietr.TrouveSensTreso(GS: THGrid; Nat: String3; Lig: integer): byte;
var
  CGen: TGGeneral;
begin
  if not M.Effet then Result := TrouveSens(GS, Nat, Lig) else
  begin
    Result := 3;
    if GS.Cells[SA_Gen, Lig] = '' then exit;
    CGen := GetGGeneral(GS, Lig);
    if CGen = nil then Exit;
    if (Cgen.natureGene = 'COC') or (Cgen.natureGene = 'TID') then Result := 2;
  end;
end;


procedure TFSaisietr.GSCellEnter(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
var
  CGen: TGGeneral;
  OnSort: Boolean;
  AccesConsult: Boolean;
begin
  if ARow <> GS.Row then
  begin
    AccesConsult := not IsAccessible(GS.Row);
    {if M.Treso then} BasculeGS(AccesConsult);
  end;
  //if M.Treso then
    OnSort := (GS.Col <> SA_Gen) and (GS.Col <> SA_Aux) and (GS.Col <> SA_DateC);
  //else
    //OnSort := (GS.Col <> SA_Gen) and (GS.Col <> SA_Aux);
  if ((Action <> taConsult) and (not EstRempli(GS, GS.Row)) and OnSort)
    then
  begin
    ACol := SA_Gen;
    ARow := GS.Row;
    Cancel := True;
    Exit;
  end;
  if ((H_NUMEROPIECE_.Visible) and (GS.Row > 1)) then
  begin
    if Transactions(AttribNumeroDef, 10) <> oeOK then MessageAlerte(HDiv.Mess[5]);
    BlocageEntete(Self, TRUE);
    {if M.Treso then} E_TYPECTR.Enabled := FALSE;
    if M.Effet then E_CONTREPARTIEGEN.Enabled := FALSE;
  end;
  BValide.Enabled := ((H_NumeroPiece.Visible) and (Equilibre));
  GereNewLigne1;
  if {M.Treso and} ((H_NUMEROPIECE_.Visible) and (GS.Row = 1)) then //TR
  begin
    if SI_NUMREF >= 0 then
    begin
      SI_NUMREF := StrToInt(Trim(E_NUMDEP.Text));
      E_NUMREF.Value := SI_NUMREF;
    end;
  end;
  // TR
  if {M.Treso and} (SAJAL.TRESO.TypCtr = Ligne) then
  begin
    CGen := GetGGeneral(GS, GS.Row);
    if (GS.Cells[SA_Gen, GS.Row] = SAJAL.TRESO.Cpt) and (CGen <> nil) and (GS.Row mod 2 = 0) then
    begin
      if ARow < GS.Row then
      begin
        ACol := GS.Col {SA_Gen};
        ARow := GS.Row + 1;
      end else
      begin
        ACol := GS.col {SA_Credit};
        ARow := GS.Row - 1;
      end;
      Cancel := TRUE;
      Exit;
    end;
  end;
  GereZoom;
  if PasModif then exit;
  if GS.Col = SA_Aux then
  begin
    // Shunt de la zone aux si compte pas collectif
    CGen := GetGGeneral(GS, GS.Row);
    if ((CGen <> nil) and (GS.Cells[SA_Gen, GS.Row] <> '') and (not CGen.Collectif)) then
    begin
      if ACol = SA_Gen then ACol := SA_RefI else ACol := SA_Gen;
      ARow := GS.Row;
      Cancel := TRUE;
    end;
  end else if GS.Col = SA_Credit then
  begin
    if GS.Cells[SA_Debit, GS.Row] <> '' then
    begin
      if GS.Row = ARow then
      begin
        if ACol = SA_Debit then
        begin
          ARow := ARow + 1;
          ACol := SA_Gen;
        end else
        begin
          ACol := SA_Debit;
          ARow := GS.Row;
        end;
      end else
      begin
        ACol := SA_Debit;
        ARow := GS.Row;
      end;
      Cancel := True;
    end;
  end else if GS.Col = SA_Debit then
  begin
    if GS.Cells[SA_Credit, GS.Row] <> '' then
    begin
      if GS.Row = ARow then
      begin
        if ACol <= SA_Lib then ACol := SA_Credit else if ACol >= SA_Credit then ACol := SA_Lib;
      end else
      begin
        ACol := SA_Credit;
        ARow := GS.Row;
      end;
      Cancel := TRUE;
    end else
    begin
      if ((GS.Row = ARow) and (TrouveSensTreso(GS, E_NATUREPIECE.Value, GS.Row) = 2) and (Acol < GS.Col) and (GS.Cells[SA_Debit, GS.Row] = ''))
        then
      begin
        ACol := SA_Credit;
        Cancel := True;
      end;
    end;
  end else if (GS.Col = SA_DateC) {and (M.Treso)} then
  begin
    if (ACol = SA_GEN) and (ARow = GS.ROW) and (ARow > 1) then
    begin
      ACol := SA_CREDIT;
      ARow := ARow - 1;
    end else
    begin
      ACol := SA_GEN;
      ARow := GS.Row;
    end;
    {
    if GS.Row=ARow then
       BEGIN
       ACol:=SA_GEN ; ARow:=GS.Row ;
       END else
       BEGIN
       if ACol=SA_GEN then BEGIN ARow:=ARow+1 ; ACol:=SA_Gen ; END else BEGIN ACol:=SA_Debit ; ARow:=GS.Row ; END ;
       END ;
    }
    Cancel := True;
  end;
end;

procedure TFSaisietr.GSRowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
var
  AccesConsult: Boolean;
begin
  GridEna(GS, FALSE);
  AccesConsult := not IsAccessible(Ou);
  {if M.Treso then} BasculeGS(AccesConsult);
  if ((not EstRempli(GS, Ou)) and (not PasModif)) then DefautLigne(Ou, True);
  CurLig := Ou;
  AffichePied;
  GereEnabled(Ou);
  // Remise à zéro du test sur l'appel au lettrage...FQ 14818 SBO 02/11/2004
  if tAppelLettrage.Size < Ou then tAppelLettrage.Size := Ou;
  tAppelLettrage.Bits[Ou - 1] := False;
  GridEna(GS, TRUE);
end;

procedure TFSaisietr.MajOBMViaEche(Lig: integer; OBM: TOBM; R: T_ModeRegl; NumEche: integer);
{ Pour contrer le bug suivant :
   Modif sur une ligne d'un aux. Désalloueeche remet etatlettrage à"RI" sur l'OBM.
   Ensuite,seul l'objet RMOD connait l'état lettrage et la maj directe table se base
   sur l'état du champ e_etatlettrage de l'OBM. D'ou le besoin de remettre à jour l'OBM
   sur les zones non saisissables concernant le lettrage
}
begin
//  if not M.Treso then Exit;
  if OBM <> nil then
  begin
    OBM.PutMvt('E_ETATLETTRAGE', R.TabEche[NumEche].EtatLettrage);
    OBM.PutMvt('E_NUMECHE', NumEche);
    if NumEche > 0 then OBM.PutMvt('E_ECHE', 'X');
  end;
end;

procedure TFSaisietr.MajDirecteTable(Ou: LongInt; NatPiece: String3);
var
  OBM: TOBM;
  St, SQL, SQL2, SQL1, SQL3: string;
  M: TMOD;
  R: T_ModeRegl;
  i: integer;
  CGen: TGGeneral;
  MVT: RMVT;
label
  0;
begin
  OBM := GetO(GS, Ou);
  if OBM <> nil then
  begin
    //   M:=TMOD(GS.Objects[SA_NumP,Ou]) ;
    if OBM.Status = EnMemoire then
    begin
      GetEcr(Ou);
      OBM.Status := DansTable;
      OBM.MajEtat(Inchange);
    end else
    begin
      M := TMOD(GS.Objects[SA_NumP, Ou]);
      if M <> nil then R := M.MODR;
      CGen := GetGGeneral(GS, Ou);
      if CGen = nil then Exit;
      (*
      St:=OBM.StPourUpdate ; If St='' Then Goto 0 ;
      SQL1:='UPDATE ECRITURE SET '+St+', E_DATEMODIF="'+UsTime(NowFutur)+'"' ;
      If NatPiece='' Then MVT:=OBMToIdent(OBM,FALSE)
                     Else MVT:=OBMToIdentManuel(OBM,FALSE,NatPiece) ;
      SQL2:=' Where  '+WhereEcriture(tsGene,MVT,FALSE)  ;
      SQL3:=' AND E_NUMLIGNE='+IntToStr(Ou)+' AND E_NUMECHE=0' ;
      SQL:=SQL1+SQL2+SQL3 ;
      *)
      if M = nil then
      begin
        St := OBM.StPourUpdate;
        if St = '' then goto 0;
        SQL1 := 'UPDATE ECRITURE SET ' + St + ', E_DATEMODIF="' + UsTime(NowFutur) + '"';
        if NatPiece = '' then MVT := OBMToIdent(OBM, FALSE)
        else MVT := OBMToIdentManuel(OBM, FALSE, NatPiece);
        SQL2 := ' Where  ' + WhereEcriture(tsGene, MVT, FALSE);
        SQL3 := ' AND E_NUMLIGNE=' + IntToStr(Ou) (*+' AND E_NUMECHE=0'*);
        SQL := SQL1 + SQL2 + SQL3;
        ExecuteSQL(SQL);
      end else
      begin
        for i := 1 to R.NbEche do
        begin
          MajOBMViaEche(Ou, OBM, R, i);
          St := OBM.StPourUpdate;
          if St <> '' then
          begin
            SQL1 := 'UPDATE ECRITURE SET ' + St + ', E_DATEMODIF="' + UsTime(NowFutur) + '"';
            if NatPiece = '' then MVT := OBMToIdent(OBM, FALSE)
            else MVT := OBMToIdentManuel(OBM, FALSE, NatPiece);
            SQL2 := ' Where  ' + WhereEcriture(tsGene, MVT, FALSE);
            SQL3 := ' AND E_NUMLIGNE=' + IntToStr(Ou) + ' AND E_NUMECHE=' + IntToStr(i);
            SQL := SQL1 + SQL2 + SQL3;
            ExecuteSQL(SQL);
          end;
        end;
      end;
      0: OBM.MajEtat(Inchange);
    end;
  end;
end;

procedure TFSaisietr.AfficheExigeTva;
var
  Nat: String3;
begin
  if not VH^.OuiTvaEnc then
  begin
    BPopTva.Visible := False;
    Exit;
  end;
  Nat := E_NATUREPIECE.Value;
  if ((Nat <> 'FC') and (Nat <> 'AC') and (Nat <> 'OC') and
    (Nat <> 'FF') and (Nat <> 'AF') and (Nat <> 'OF')) then
  begin
    BPopTva.Visible := False;
    Exit;
  end;
  if (((Nat = 'OC') or (Nat = 'FC') or (Nat = 'AC')) and (VH^.TvaEncSociete = 'TD')) then
  begin
    BPopTva.Visible := False;
    Exit;
  end;
  if (((Nat = 'FC') or (Nat = 'AC')) and (SorteTva <> stvVente)) then
  begin
    BPopTva.Visible := False;
    Exit;
  end;
  if (((Nat = 'FF') or (Nat = 'AF')) and (SorteTva <> stvAchat)) then
  begin
    BPopTva.Visible := False;
    Exit;
  end;
  BPopTva.Visible := True;
end;

procedure TFSaisietr.AffecteTva(C, L : integer);
var
  O: TOBM;
  CAux: TGTiers;
  CGen: TGGeneral;
  StE: string;
begin
  {#TVAENC}
  if Action = taConsult then Exit;
  O := GetO(GS, L);
  if O = nil then Exit;
  if isTiers(GS, L) then
  begin
    RenseigneRegime(L);
    if GS.Cells[SA_Aux, L] <> '' then
    begin
      CAux := GetGTiers(GS, L);
      if CAux = nil then Exit;
      if not ChoixExige then
      begin
        if ((SorteTva = stvAchat) or (E_NATUREPIECE.Value = 'OF')) then ExigeTva := Code2Exige(CAux.Tva_Encaissement) else
          if ((SorteTva = stvDivers) and (E_NATUREPIECE.Value = 'OC') and (VH^.TvaEncSociete = 'TE')) then ExigeTva := tvaEncais
        else ExigeTva := tvaMixte;
        ExigeEntete := ExigeTva;
        BPopTva.ItemIndex := Ord(ExigeEntete) + 1;
      end;
    end else
    begin
      CGen := GetGGeneral(GS, L);
      if CGen = nil then Exit;
      if not ChoixExige then
      begin
        if ((SorteTva = stvAchat) or (E_NATUREPIECE.Value = 'OF')) then ExigeTva := Code2Exige(CGen.Tva_Encaissement) else
          if ((SorteTva = stvDivers) and (E_NATUREPIECE.Value = 'OC') and (VH^.TvaEncSociete = 'TE')) then ExigeTva := tvaEncais
        else ExigeTva := tvaMixte;
        ExigeEntete := ExigeTva;
        BPopTva.ItemIndex := Ord(ExigeEntete) + 1;
      end;
    end;
    O.PutMvt('E_REGIMETVA', GeneRegTva);
  end else if isHT(GS, L, True) then
  begin
    CGen := GetGGeneral(GS, L);
    if CGen = nil then Exit;
    StE := FlagEncais(ExigeEntete, CGen.TvaSurEncaissement);
    O.PutMvt('E_TVAENCAISSEMENT', StE);
    O.PutMvt('E_REGIMETVA', GeneRegTva);
    O.PutMvt('E_TVA', CGen.Tva);
    if not IsTiersSoumisTPF(GS) then
      O.PutMvt('E_TPF', '')
    else
      O.PutMvt('E_TPF', CGen.Tpf);
  end;
end;

procedure TFSaisietr.RenseigneRegime(Lig: integer);
var
  CGen: TGGeneral;
  CAux: TGTiers;
begin
  if Lig <= 0 then Exit;
  if GS.Cells[SA_Aux, Lig] <> '' then
  begin
    CAux := GetGTiers(GS, Lig);
    if CAux <> nil then
    begin
      if ((GeneRegTVA = '') or (not RegimeForce)) then GeneRegTVA := CAux.RegimeTVA;
      GeneSoumisTPF := CAux.SoumisTPF;
    end;
  end else
  begin
    CGen := GetGGeneral(GS, Lig);
    if CGen <> nil then
    begin
      if ((GeneRegTVA = '') or (not RegimeForce)) then GeneRegTVA := CGen.RegimeTVA;
      if (VH^.PaysLocalisation=CodeISOES) then
         GeneSoumisTPF:=IsTiersSoumisTPF(GS)
      else
         GeneSoumisTPF:=CGen.SoumisTPF ;
    end;
  end;
end;

procedure TFSaisietr.GetSorteTva;
begin
  SorteTva := stvDivers;
  if SAJAL = nil then Exit;
  if ((SAJAL.NatureJal = 'VTE') and ((E_NATUREPIECE.Value = 'FC') or (E_NATUREPIECE.Value = 'AC'))) then SorteTva := stvVente else
    if ((SAJAL.NatureJal = 'ACH') and ((E_NATUREPIECE.Value = 'FF') or (E_NATUREPIECE.Value = 'AF'))) then SorteTva := stvAchat;
end;

function TFSaisietr.TvaTauxDirecteur(Lig: integer; Client, ToutDebit: boolean; Regime: String3): Boolean;
var
  OMODR: TMOD;
  TTC, HT: Double;
  CodeTva: String3;
  Taux, XX: Double;
  i: integer;
  O: TOBM;
begin
  {#TVAENC}
  Result := FALSE;
  if not VH^.OuiTvaEnc then Exit;
  OMODR := TMOD(GS.Objects[SA_NumP, Lig]);
  if OMODR = nil then Exit;
  {Si modifié manuellement ou recalculé prorata --> pas de traitement}
  if ((OMODR.ModR.ModifTva) or (OMODR.ModR.TotalAPayerP = 0)) then Exit;
  O := GetO(GS, Lig);
  if O = nil then Exit;
  TTC := O.GetMvt('E_CREDIT') - O.GetMvt('E_DEBIT');
  if not Client then TTC := -TTC;
  CodeTva := VH^.NumCodeBase[1];
  if CodeTva = '' then Exit;
  Taux := Tva2Taux(Regime, CodeTva, not Client);
  if Taux = -1 then Exit;
  {Prorater le HT, calculé sur le taux directeur, sur les échéances}
  HT := Arrondi(TTC / (1.0 + Taux), V_PGI.OkDecV);
  for i := 1 to MaxEche do
  begin
    FillChar(OMODR.ModR.TabEche[i].TAV, Sizeof(OMODR.ModR.TabEche[i].TAV), #0);
    XX := Arrondi(HT * OMODR.ModR.TabEche[i].MontantP / OMODR.ModR.TotalAPayerP, V_PGI.OkDecV);
    if ToutDebit then
    begin
      if XX <> OMODR.ModR.TabEche[i].TAV[5] then Result := TRUE;
      OMODR.ModR.TabEche[i].TAV[5] := XX;
    end else
    begin
      if XX <> OMODR.ModR.TabEche[i].TAV[1] then Result := TRUE;
      OMODR.ModR.TabEche[i].TAV[1] := XX;
    end;
  end;
  {Eviter la ré-initialisation ultérieure, recalcul par prorata ou manuel}
  OMODR.ModR.ModifTva := True;
  Revision := False;
end;

function TFSaisietr.GereTvaEncNonFact(Lig: integer): Boolean;
var
  CGen: TGGeneral;
  CAux: TGTiers;
  Client: boolean;
  RegEnc, StEnc, Regime: String3;
  Exi: TExigeTva;
begin
  {#TVAENC}
  Result := FALSE;
  if PasModif then Exit;
  if not VH^.OuiTvaEnc then Exit;
  {Si facture --> géré ailleurs (Calculdébitcrédit)}
  if SorteTva <> stvDivers then Exit;
  if not isTiers(GS, Lig) then Exit;
  Regime := '';
  if GS.Cells[SA_Aux, Lig] <> '' then
  begin
    CAux := GetGTiers(GS, Lig);
    if CAux = nil then Exit;
    if ((CAux.NatureAux = 'CLI') or (CAux.NatureAux = 'AUD')) then Client := True else
      if ((CAux.NatureAux = 'FOU') or (CAux.NatureAux = 'AUC')) then Client := False else Exit;
    StEnc := CAux.Tva_Encaissement;
    Regime := CAux.RegimeTva;
  end else if GS.Cells[SA_Gen, Lig] <> '' then
  begin
    CGen := GetGGeneral(GS, Lig);
    if CGen = nil then Exit;
    if CGen.NatureGene = 'TID' then Client := True else
      if CGen.NatureGene = 'TIC' then Client := False else Exit;
    StEnc := CGen.Tva_Encaissement;
    Regime := CGen.RegimeTva;
  end else Exit;
  CGen := GetGGeneral(GS, Lig);
  if CGen = nil then Exit;
  if not EstCollFact(CGen.General) then Exit;
 
  {Si client --> param soc, sinon exige du fourn. Si choix explicite, forcer. Débit --> Aucun traitement}
  if Client then RegEnc := VH^.TvaEncSociete else RegEnc := StEnc;
  if ChoixExige then Exi := TExigeTva(BPopTva.ItemIndex - 1) else Exi := Code2Exige(RegEnc);
  case Exi of
    tvaMixte: Exit;
    tvaEncais: Result := TvaTauxDirecteur(Lig, Client, False, Regime);
    tvaDebit: Result := TvaTauxDirecteur(Lig, Client, True, Regime);
  end;
end;

procedure TFSaisietr.GSRowExit(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
var
  OkL, bc, AppelLettrage: Boolean;
  OBM: TOBM;
begin
  if PasModif then Exit;
  if Ou >= GS.RowCount - 1 then Exit;
  BloqueGrille(TRUE);
  OkL := LigneCorrecte(Ou, False, True, TRUE);

  // Pb double affectation des zones suite test comptes divers lettrables FQ 16136 SBO 10/08/2005
  OBM := GetO(GS,Ou) ;
  if ({(M.Treso) and} (OBM.GetNumChamp('_ETAT') <> - 1) and (TEtat( GetO(GS,Ou).GetValue('_Etat') ) = Inchange)) then
  //  if ((M.Treso) and {(SAJal.Treso.TypCtr=Ligne) AND}(GetO(GS, Ou).Etat = Inchange)) then
  begin
    if not OkL then Cancel := TRUE;
    BloqueGrille(FALSE);
    exit;
  end;
  if OkL then
  begin

    if ((H_NUMEROPIECE_.Visible) and (ou = 1)) then
    begin
      if Transactions(AttribNumeroDef, 10) <> oeOK then MessageAlerte(HDiv.Mess[5]);
      BlocageEntete(Self, TRUE);
      {if M.Treso then} E_TYPECTR.Enabled := FALSE;
      if M.Effet then E_CONTREPARTIEGEN.Enabled := FALSE;
      OBM := GetO(GS, Ou);
      if OBM <> nil then OBM.PutMvt('E_NUMEROPIECE', NumPieceInt);
    end;

    GereComplements(Ou);
    GereEche(Ou, True, False);
    GereAnal(Ou, True, True);
    GereTvaEncNonFact(Ou);

    // Dev 3946 : control cohérence qté analytique soumis à indicateur
    OBM := GetO(GS,Ou) ;
    if ( OBM <> nil ) and ( OBM.GetValue('E_ANA')='X' )
                      and TestQteAna
                      and ( ( OBM.GetNumChamp('CHECKQTE') < 0 )  or ( OBM.GetString('CHECKQTE')<>'X'  ) )
        then CheckQuantite( OBM ) ;
    // Fin Dev 3946

    if SAJAL.TRESO.TypCtr = Ligne then
    begin
      //      BloqueGrille(FALSE) ;
      GereLigneTreso(Ou);
      //      BloqueGrille(TRUE) ;
    end;

    if M.MajDirecte then
    begin
      BeginTrans;
      MajDirecteTable(ou, '');
      if SAJAL.TRESO.TypCtr = Ligne then MajDirecteTable(ou + 1, '');
      CommitTrans;
      AppelLettrage := OkScenario and (Scenario.GetMvt('SC_LETTRAGESAISIE') = 'X');
      if tAppelLettrage.Size < Ou then tAppelLettrage.Size := Ou;
      if AppelLettrage and (not tAppelLettrage.Bits[Ou - 1]) then LettrageEnSaisieTreso(ou, FALSE); // 13524
      tAppelLettrage.Bits[Ou - 1] := True; // 14151
      BloqueGrille(FALSE);
      if AppelLettrage then
      begin
        RowEnterPourModeConsult := TRUE;
        GSRowEnter(nil, GS.Row, bc, False);
        RowEnterPourModeConsult := FALSE;
      end;
    end else BloqueGrille(FALSE);
    if SAJAL.TRESO.TypCtr in [PiedSolde, PiedDC] then
    begin
      if E_REFTRESO.TEXT = '' then
      begin
        E_REFTRESO.TEXT := GS.Cells[SA_REFI, Ou];
        MajRefLibOBMT(0);
      end;
      if E_LIBTRESO.TEXT = '' then
      begin
        E_LIBTRESO.TEXT := GS.Cells[SA_LIB, Ou];
        MajRefLibOBMT(1);
      end;
    end;
  end else
  begin
    Cancel := True;
    BloqueGrille(FALSE);
  end;

end;

procedure TFSaisietr.GSEnter(Sender: TObject);
var
  StComp: string;
  b, OkD: boolean;
  RC: R_COMP;
  StLibreEnt: string;
begin
{Pour afficher les scrollbars, bien que je pense que ce ne soit pas la source du problème de la FQ 16790}
  ShowScrollBar(GS.Handle , SB_BOTH, True);
  if SAJAL = nil then Exit;
  if SAJAL.Journal <> E_JOURNAL.Value then
  begin
    MessageAlerte(HDiv.Mess[13]);
    Exit;
  end;
  GereEnabled(GS.Row);
  if Action <> taCreat then Exit;
  if PasModif then Exit;
  if not OkScenario then Exit;
  if GS.RowCount > 2 then Exit;
  if DejaRentre then Exit;
  if Entete = nil then Exit;
  if EstRempli(GS, 1) then Exit;
  {$IFNDEF SPEC302}
  OkD := False;
  if VH^.CPDateObli then OkD := True else
  begin
    if ((OkScenario) and (Scenario <> nil)) then OkD := (Scenario.GetMvt('SC_DATEOBLIGEE') = 'X');
  end;
  if ((Action = taCreat) and (not EstRempli(GS, 1)) and
    (OkD) and (not DejaRentre) and (not RentreDate)) then
  begin
    // JLD 4394 if HPiece.Execute(41,Caption,'')<>mrYes then BEGIN E_DATECOMPTABLE.SetFocus ; Exit ; END ;
    if E_DATECOMPTABLE.CanFocus then E_DATECOMPTABLE.SetFocus;
    Exit;
  end;
  {$ENDIF}
  RentreDate := True;
  DejaRentre := True;
  StComp := Entete.GetMvt('E_ETAT');
  StLibreEnt := Scenario.GetMvt('SC_LIBREENTETE');
  if ((Pos('X', StComp) <= 0) and (Pos('X', StLibreEnt) <= 0)) then Exit;
  RC.StComporte := StComp;
  RC.StLibre := StLibreEnt;
  RC.Conso := False;
  RC.Attributs := True;
  RC.MemoComp := MemoComp;
  RC.Origine := 0;
  RC.DateC := DatePiece;
  {08/06/06 : FQ 17556 : TOBCompl sert pour le CutOff, donc les comptes de charges / produits.
              Cela a donc peu de raison d'être ici}
  RC.TOBCompl := nil ;
  SaisieComplement(Entete, EcrGen, Action, b, RC);
  DejaRentre := True;
  DefautLigne(1, True);
  GereEnabled(1);
end;

procedure TFSaisietr.GSExit(Sender: TObject);
var
  b, bb: boolean;
  C, R: longint;
begin
  bb := FALSE;
  b := FALSE;
  if PasModif then Exit;
  if Valide97.Tag <> 1 then Exit;
  if Outils.Tag <> 1 then Exit;
  C := GS.Col;
  R := GS.Row;
  GSCellExit(GS, C, R, b);
  GSRowExit(GS, R, b, bb);
end;

procedure TFSaisietr.GSMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  GX := X;
  GY := Y;
end;

procedure TFSaisietr.POPSPopup(Sender: TObject);
begin
  InitPopUp(Self);
end;

procedure TFSaisietr.FindSaisFind(Sender: TObject);
begin
  Rechercher(GS, FindSais, FindFirst);
end;

procedure TFSaisietr.GSSetEditText(Sender: TObject; ACol, ARow: Longint; const Value: string);
begin
  PieceModifiee := True;
end;

procedure TFSaisietr.H_MODEPAIEDblClick(Sender: TObject);
begin
  FicheModePaie_AGL(E_MODEPAIE.VALUE);
end;

procedure TFSaisietr.ModifOBMTSerie(NomChamp: string; V: Variant);
begin
  if (SAJAL.TRESO.TYPCTR in [PiedDC, PiedSolde]) then
  begin
    if (OBMT[1] <> nil) then OBMT[1].PutMvt(NomChamp, V);
    if (OBMT[2] <> nil) then OBMT[2].PutMvt(NomChamp, V);
  end;
end;

procedure TFSaisietr.BModifSerieClick(Sender: TObject);
var
  b: boolean;
  St, StAvant: string;
  Lig: integer;
  OBM: TOBM;
  DD, DDAvant: TDateTime;
  RC: R_COMP;
  OkMajDirecte: boolean;
begin
  if Action = taConsult then Exit;
  RC.StComporte := '--XXXX----';
  RC.StLibre := '';
  RC.Conso := True;
  RC.Attributs := False;
  RC.MemoComp := nil;
  RC.Origine := -1;
  RC.DateC := DatePiece;
  {pas de modif série zones libres en saisie tréso}
  ModifSerie.Free;
  ModifSerie := TOBM.Create(EcrGen, '', True);
  {08/06/06 : FQ 17556 : TOBCompl sert pour le CutOff, donc les comptes de charges / produits.
              Cela a donc peu de raison d'être ici}
  RC.TOBCompl := nil ;
  if not SaisieComplement(ModifSerie, EcrGen, taCreat, b, RC) then Exit;
  if HPiece.Execute(28, caption, '') <> mrYes then Exit;
  BeginTrans;
  for Lig := 1 to GS.RowCount - 2 do
  begin
    OBM := GetO(GS, Lig);
    OkMajDirecte := FALSE;
    if OBM <> nil then
    begin
      StAvant := OBM.GetMvt('E_REFEXTERNE');
      St := ModifSerie.GetMvt('E_REFEXTERNE');
      if St <> '' then OBM.PutMvt('E_REFEXTERNE', St);
      if (St <> '') and (StAvant <> St) and M.MajDirecte then OkMajDirecte := TRUE;
      if (St <> '') then ModifOBMTSerie('E_REFEXTERNE', St);

      StAvant := OBM.GetMvt('E_REFLIBRE');
      St := ModifSerie.GetMvt('E_REFLIBRE');
      if St <> '' then OBM.PutMvt('E_REFLIBRE', St);
      if (St <> '') and (StAvant <> St) and M.MajDirecte then OkMajDirecte := TRUE;
      ModifOBMTSerie('E_REFLIBRE', St);

      StAvant := OBM.GetMvt('E_AFFAIRE');
      St := ModifSerie.GetMvt('E_AFFAIRE');
      if St <> '' then OBM.PutMvt('E_AFFAIRE', St);
      if (St <> '') and (StAvant <> St) and M.MajDirecte then OkMajDirecte := TRUE;
      ModifOBMTSerie('E_AFFAIRE', St);

      StAvant := OBM.GetMvt('E_CONSO');
      St := ModifSerie.GetMvt('E_CONSO');
      if St <> '' then OBM.PutMvt('E_CONSO', St);
      if (St <> '') and (StAvant <> St) and M.MajDirecte then OkMajDirecte := TRUE;
      ModifOBMTSerie('E_CONSO', St);

      DDAvant := OBM.GetMvt('E_DATEREFEXTERNE');
      DD := ModifSerie.GetMvt('E_DATEREFEXTERNE');
      if DD > Encodedate(1901, 01, 01) then OBM.PutMvt('E_DATEREFEXTERNE', DD);
      if (DD > Encodedate(1901, 01, 01)) and (DD <> DDAvant) and M.MajDirecte then OkMajDirecte := TRUE;
      ModifOBMTSerie('E_DATEREFEXTERNE', DD);

      if OkMajDirecte and StatusOBMOk(OBM) then MajDirecteTable(Lig, '');
    end;
  end;
  CommitTrans;
end;

procedure TFSaisietr.BZoomTClick(Sender: TObject);
var
  A: TActionFiche;
begin
  // attention, utiliser JaiLeDroit
  A := taModif;
  if SAJAL = nil then Exit;
  if not ExJaiLeDroitConcept(TConcept(ccGenModif), False) then A := taConsult; //FicheGene(Nil,'',SAJAL.TRESO.CPT,A,0) ;//XMG 20/04/04
  FicheGene(nil, '', SAJAL.TRESO.CPT, A, 0);
end;


procedure TFSaisietr.BComplementTClick(Sender: TObject);
begin
  ClickComplementT;
end;

procedure TFSaisietr.BSaisTauxClick(Sender: TObject);
begin
  ClickSaisTaux;
end;

procedure TFSaisietr.BChoixRegimeClick(Sender: TObject);
begin
  ClickChoixRegime;
end;

procedure TFSaisietr.WMGetMinMaxInfo(var MSG: Tmessage);
begin
  with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do
  begin
    X := WMinX;
    Y := WMinY;
  end;
end;

procedure TFSaisietr.GSKeyPress(Sender: TObject; var Key: Char);
begin
  if MajEnCours or (not GS.SynEnabled) then
  begin
    Key := #0;
    Exit;
  end;
end;

procedure TFSaisietr.BAideClick(Sender: TObject);
begin
  CallHelpTopic(Self);
end;

procedure TFSaisietr.E_JOURNALClick(Sender: TObject);
var
  Jal: string;
label
  0;
begin
  Jal := E_JOURNAL.Value;
  if Jal = '' then
  begin
    SAJAL.Free;
    SAJAL := nil;
    goto 0;
  end;
  if SAJAL <> nil then if SAJAL.JOURNAL = Jal then goto 0;
  if ((VH^.JalAutorises <> '') and (Pos(';' + Jal + ';', VH^.JalAutorises) <= 0)) then
  begin
    HPiece.Execute(26, caption, '');
    if SAJAL <> nil then
    begin
      SAJAL.Free;
      SAJAL := nil;
    end;
    E_JOURNAL.Value := '';
    if Action = taCreat then
    begin
      E_JOURNAL.SetFocus;
      GS.Enabled := False;
    end;
    goto 0;
  end;
  ChercheJal(Jal, False);
  if ((Action = taCreat) and (NumPieceInt = 0)) then
  begin
    HPiece.Execute(27, caption, '');
    if SAJAL <> nil then
    begin
      SAJAL.Free;
      SAJAL := nil;
    end;
    E_JOURNAL.Value := '';
    E_JOURNAL.SetFocus;
    GS.Enabled := False;
    goto 0;
  end;
  CalculSoldeCompteT;
//  if M.Treso then //TR
  //begin
    AlimObjetMvt(0, FALSE, 1);
    AlimObjetMvt(0, FALSE, 2);
//  end;
  InitBoutonValide;
  0:
end;

procedure TFSaisietr.BMenuZoomMouseEnter(Sender: TObject);
begin
  PopZoom97(BMenuZoom, POPZ);
  Outils.Tag := 1;
end;

procedure TFSaisietr.BValideMouseEnter(Sender: TObject);
begin
  Valide97.Tag := 1;
end;

procedure TFSaisietr.BValideMouseExit(Sender: TObject);
begin
  Valide97.Tag := 0;
end;

procedure TFSaisietr.BMenuZoomMouseExit(Sender: TObject);
begin
  Outils.Tag := 0;
end;


procedure TFSaisietr.BRAAbandonClick(Sender: TObject);
begin
  PEntete.Enabled := True;
  Outils.Enabled := True;
  Valide97.Enabled := True;
  GS.Enabled := True;
  PPied.Enabled := True;
  GS.SetFocus;
  PRefAuto.Visible := False;
  ModeRA := False;
end;

procedure TFSaisietr.BRAValideClick(Sender: TObject);
var
  OBM: TOBM;
  i: Integer;
  OkOk: Boolean;
  St, RefAvant: string;
begin
  SI_NumRef := -1;
  E_NumDepExit(Sender);
  BeginTrans;
  for i := 1 to GS.RowCount - 2 do
  begin
    OkOk := TRUE;
    St := '';
    if (SAJAL.Treso.TypCtr = Ligne) and (GS.Cells[SA_Gen, i] = SAJAL.TRESO.Cpt) then OkOk := FALSE;
    if OkOk and (i > 1) then IncRef;
    if SI_NumRef < 0 then St := '' else St := E_NUMREF.Text;
    GS.Cells[SA_RefI, i] := St;
    OBM := GetO(GS, i);
    if OBM <> nil then
    begin
      RefAvant := OBM.GetMvt('E_REFINTERNE');
      OBM.PutMvt('E_REFINTERNE', St);
      if (RefAvant <> St) and M.MajDirecte and StatusOBMOk(OBM) then MajDirecteTable(i, '');
    end;
    { A faire : Idem Sur analytique }
  end;
  CommitTrans;
  BRAAbandonClick(nil);
end;

function TFSaisietr.Load_Sais(St: hstring): Variant;
var
  V: Variant;
  Lig: integer;
  OBM: TOBM;
  CGen: TGGeneral;
  CAux: TGTiers;
begin
  V := #0;
  Result := V;
  St := uppercase(Trim(St));
  if St = '' then Exit;
  Lig := QuelleLigne(St);
  if ((Lig = -1) and (CurLig > 1)) then Lig := CurLig - 1 else if Lig <= 0 then Lig := CurLig;
  if Lig <= 0 then Lig := GS.Row;
  OBM := GetO(GS, Lig);
  CGen := GetGGeneral(GS, Lig);
  CAux := GetGTiers(GS, Lig);
  {Zones entête}
  if ((St = 'JOURNAL') or (St = 'E_JOURNAL')) then V := SAJAL.JOURNAL else
    if ((St = 'DATECOMPTABLE') or (St = 'E_DATECOMPTABLE')) then V := StrToDate(E_DATECOMPTABLE.Text) else
    if ((St = 'NATUREPIECE') or (St = 'E_NATUREPIECE')) then V := E_NATUREPIECE.Text else
    if ((St = 'NUMEROPIECE') or (St = 'E_NUMEROPIECE')) then V := NumPieceInt else
    if ((St = 'DEVISE') or (St = 'E_DEVISE')) then V := E_DEVISE.Text else
    if ((St = 'ETABLISSEMENT') or (St = 'E_ETABLISSEMENT')) then V := E_ETABLISSEMENT.Text else
    if St = 'E_GENERAL' then
  begin
    if CGen <> nil then V := CGen.General;
  end else
    if St = 'E_AUXILIAIRE' then
  begin
    if CAux <> nil then V := CAux.auxi;
  end else
    (*
    if St='TVA' then V:=Load_TVATPF(CGen,TRUE,'') else
    if St='TVANOR' then V:=Load_TVATPF(CGen,TRUE,'NOR') else
    if St='TVARED' then V:=Load_TVATPF(CGen,TRUE,'RED') else
    if St='TPF' then V:=Load_TVATPF(CGen,False,'') else
    *)
    if St = 'SOLDE' then V := 'SOLDE' else
    {Zones Général}
    if Copy(St, 1, 2) = 'G_' then
  begin
    if CGen = nil then Exit;
    System.Delete(St, 1, 2);
    if St = 'GENERAL' then V := CGen.General else
      if St = 'LIBELLE' then V := CGen.Libelle else
      if St = 'ABREGE' then V := CGen.Abrege else
      V := RechercheLente('G_' + St, CGen.General);
  end else
    {Zones Auxiliaire}
    if Copy(St, 1, 2) = 'T_' then
  begin
    if CAux = nil then Exit;
    System.Delete(St, 1, 2);
    if St = 'AUXILIAIRE' then V := CAux.Auxi else
      if St = 'LIBELLE' then V := CAux.Libelle else
      if St = 'ABREGE' then V := CAux.Abrege else
      V := RechercheLente('T_' + St, CAux.Auxi);
  end else
    {Zones Journal}
    if Copy(St, 1, 2) = 'J_' then
  begin
    if St = 'J_JOURNAL' then V := SAJAL.Journal else
      if St = 'J_LIBELLE' then V := SAJAL.LibelleJournal else
      if St = 'J_ABREGE' then V := SAJAL.AbregeJournal else
      V := RechercheLente('J_' + St, SAJAL.JOURNAL);
  end else
    {Comptes auto}
    if Copy(St, 1, 4) = 'AUTO' then V := TrouveAuto(SAJAL.COMPTEAUTOMAT, Ord(St[Length(St)]) - 48) else
    if St = 'INTITULE' then
  begin
    if CAux <> nil then V := CAux.Libelle else if CGen <> nil then V := CGen.Libelle;
  end else
    {Zones Ecriture}
  begin
    if OBM = nil then Exit;
    if Copy(St, 1, 2) = 'E_' then System.Delete(St, 1, 2);
    if St = 'REFERENCE' then St := 'REFINTERNE';
    V := OBM.GetMvt('E_' + St);
  end;
  Load_Sais := V;
end;

procedure TFSaisietr.E_DATEECHEANCEExit(Sender: TObject);
begin
  DateEcheDefaut := StrToDate(E_DATEECHEANCE.Text);
  if SAJAL <> nil then E_CPTTRESO.Caption := SAJAL.Treso.Cpt;
end;

procedure TFSaisietr.ModifOBMPourTVAENC(i: Integer);
var
  O: TOBM;
  CGen: TGGeneral;
  Coll: string;
  j: Integer;
  OMODR: TMOD;
begin
  CGen := nil;
  O := GetO(GS, i);
  if GS.Objects[SA_Gen, i] <> nil then CGen := TGGeneral(GS.Objects[SA_Gen, i]);
  if VH^.OuiTvaEnc and (CGen <> nil) and (O <> nil) then
  begin
    Coll := CGen.General;
    OMODR := TMOD(GS.Objects[SA_NumP, i]);
    if OMODR = nil then Exit;
    if EstCollFact(Coll) then
    begin
      if (SorteTva = stvDivers) then
      begin
        {Cycle bases HT sur échéances}
        for j := 1 to 4 do O.PutMvt('E_ECHEENC' + IntToStr(j), OMODR.MODR.TabEche[1].TAV[j]);
        O.PutMvt('E_ECHEDEBIT', OMODR.MODR.TabEche[1].TAV[5]);
      end;
    end;
  end;
end;

procedure TFSaisietr.BPopTvaChange(Sender: TObject);
var
  O: TOBM;
  i: integer;
  CGen: TGGeneral;
  Ste, StEAvant: string;
  OMODR: TMOD;
  OkMajDirecte, BaseOntChange: Boolean;
begin
  ChoixExige := True;
  GridEna(GS, False);
  O := nil;
  ExigeEntete := TExigeTva(BPopTva.ItemIndex - 1);
  for i := 1 to GS.RowCount - 2 do
  begin
    OkMajDirecte := FALSE;
    if isTiers(GS, i) then
    begin
      O := GetO(GS, i);
      OMODR := TMOD(GS.Objects[SA_NumP, i]);
      if OMODR = nil then Continue;
      OMODR.ModR.ModifTva := False;
      BaseOntChange := GereTvaEncNonFact(i);
      if BaseOntChange and M.MajDirecte then
      begin
        OkMajDirecte := TRUE;
        ModifOBMPourTvaEnc(i);
      end;
    end else if isHT(GS, i, True) then
    begin
      O := GetO(GS, i);
      if O = nil then Break;
      CGen := GetGGeneral(GS, i);
      if CGen = nil then Break;
      StEAvant := O.GetMvt('E_TVAENCAISSEMENT');
      StE := FlagEncais(ExigeEntete, CGen.TvaSurEncaissement);
      O.PutMvt('E_TVAENCAISSEMENT', StE);
      if (StEAvant <> StE) and M.MajDirecte then OkMajDirecte := TRUE;
    end;
    if OkMajDirecte and StatusOBMOk(O) then MajDirecteTable(i, '');
  end;
  GridEna(GS, True);
end;

procedure TFSaisietr.GereTagEnter(Sender: TObject);
var
  B: TToolbarButton97;
begin
  B := TToolbarButton97(Sender);
  if B = nil then Exit;
  if B.Parent = Outils then Outils.Tag := 1;
end;

procedure TFSaisietr.GereTagExit(Sender: TObject);
var
  B: TToolbarButton97;
begin
  B := TToolbarButton97(Sender);
  if B = nil then Exit;
  if B.Parent = Outils then Outils.Tag := 0;
end;

procedure TFSaisietr.BMenuZoomTMouseEnter(Sender: TObject);
begin
  PopZoom97(BMenuZoomT, POPZT);
end;

procedure TFSaisietr.GereAffSolde(Sender: TObject);
var
  Nam: string;
  C: THLabel;
begin
  Nam := THNumEdit(Sender).Name;
  Nam := 'L' + Nam;
  C := THLabel(FindComponent(Nam));
  if C <> nil then C.Caption := THNumEdit(Sender).Text;
end;

procedure TFSaisietr.SetTypeExo;
begin
  // FQ 16852 : SBO 30/03/2005
  GeneTypeExo := CGetTypeExo( StrToDate(E_DATECOMPTABLE.Text) ) ;
end;

{JP 20/10/05 : FQ 16923 : On applique le même système que dans Saisie.Pas, à savoir
               que l'on renseigne uniquement les lignes vides
{---------------------------------------------------------------------------------------}
procedure TFSaisietr.GereRefLib(Lig : Integer; var StRef, StLib : string; LigOk : Boolean);
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

    {Mise à jour évntuelle dans la table si le libellé ou la référence à changé}
    if ((StRefAvant = '') or (StLibAvant = '')) and M.MajDirecte and StatusOBMOk(OBM) then
      MajDirecteTable(Lig, '');
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

    {Mise à jour éventuelle dans la table si le libellé ou la référence à changé}
    if ((StRef <> '') or (StLib <> '')) and M.MajDirecte and StatusOBMOk(OBM) then
      MajDirecteTable(Lig, '');
  end;
end;

{JP 14/11/05 : FQ 16807 : Gère la contrepartie (genéral et auxiliaire) sur le compte bancaire :
               Plutôt que de reprendre systématiquement ce qui est défini sur la première ligne,
               on affine en fonction du type de pièce que l'on génère
{---------------------------------------------------------------------------------------}
procedure TFSaisietr.GetContrepartie(var Aux, Gen : string);
{---------------------------------------------------------------------------------------}
var
  n  : Integer;
  Ok : Boolean;
  s : string;
begin
  Aux := '';
  Gen := '';
  Ok  := False;
  s := E_NATUREPIECE.Value;

  for n := 1 to GS.RowCount - 1 do begin
    If EstRempli(GS, n) then // YMO 18/04/2006 FQ17588
    // On vérifie que le compte est bien renseigné ; dans le cas 'OD' par exemple,
    // on va chercher la nature du compte dans une cellule qui n'existe pas
    begin
      if (E_NATUREPIECE.Value = 'OC') or (E_NATUREPIECE.Value = 'RC') then {28/02/06 : OC au lieu de AC}
        Ok := TGGeneral(GS.Objects[SA_Gen, n]).NatureGene = 'COC' {Collectif client}
      else if (E_NATUREPIECE.Value = 'OF') or (E_NATUREPIECE.Value = 'RF') then {28/02/06 : OF au lieu de AF}
        Ok := TGGeneral(GS.Objects[SA_Gen, n]).NatureGene = 'COF' {Collectif Fournisseur}
      else if (E_NATUREPIECE.Value = 'OD') then
        Ok :=  (TGGeneral(GS.Objects[SA_Gen, n]).NatureGene = 'CHA') or {Compte de charge}
               (TGGeneral(GS.Objects[SA_Gen, n]).NatureGene = 'TID') or {Tiers débiteur}
               (TGGeneral(GS.Objects[SA_Gen, n]).NatureGene = 'TIC') or {Tiers créditeur}
               (TGGeneral(GS.Objects[SA_Gen, n]).NatureGene = 'PRO');   {Compte de produit}

      if Ok then begin
         Aux := GS.Cells[SA_Aux, n];
         Gen := GS.Cells[SA_Gen, n];
         Break;
      end;

    end;
  end;

  if Gen = '' then begin
    Gen := GS.Cells[SA_Gen, 1];
    Aux := GS.Cells[SA_Aux, 1];
  end;
end;


function TFSaisietr.GetFormulePourCalc(Formule: hstring): Variant;
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

function TFSaisietr.SupprimeLigneEnBase(Lig: integer): boolean;
var sSQL : string ;
begin
  sSQL := 'DELETE FROM ECRITURE WHERE '
          +     'E_JOURNAL="'       + SaJal.Journal + '" '
          + 'AND E_EXERCICE="'      + QuelExo(E_DateComptable.EditText) + '" '
          + 'AND E_DATECOMPTABLE="' + UsDateTime(StrToDate(E_DATECOMPTABLE.TEXT)) + '" '
          + 'AND E_NUMEROPIECE='    + IntToStr( NumPieceInt ) + ' '
          + 'AND E_NUMLIGNE='       + IntToStr( Lig ) ;

  result := ExecuteSQL( sSQL ) = 1 ;
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
function TFSaisietr.GetFormatPourCalc: string;
begin
  // FQ18371 et FQ18080
  result := '"#.###' ;
  if DEV.Decimale > 0 then
    result := result + ',' + Copy('000000000', 1, Dev.Decimale ) ;
  result := result + '"' ;
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
function TFSaisietr.TestQteAna: boolean;
begin
  if ctxPCL in V_PGI.PGIContexte then
    result := False
  else if OkScenario then
    result := Scenario.GetValue('SC_CONTROLEQTE') = 'X'
  else
    result := GetParamSocSecur('SO_ZCTRLQTE', False) ;
end;

end.

