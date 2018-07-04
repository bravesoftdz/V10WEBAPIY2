{$A+,B-,C+,D+,E-,F-,G+,H+,I+,J+,K-,L+,M-,N+,O-,P+,Q-,R-,S-,T-,U-,V+,W-,X+,Y+,Z1}
{$MINSTACKSIZE $00004000}
{$MAXSTACKSIZE $00100000}
{$IMAGEBASE $00400000}
{$APPTYPE GUI}
unit SaisBor ;

interface

uses
  {$IFDEF laurent}
  uWA , uHttp ,
  {$ENDIF}
  Windows,
  Messages,
  SysUtils,
  Classes,
  Graphics,
  Controls,
  Forms,
  StdCtrls,
  Grids,
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  DB,
  MenuOLG,
  {$ELSE}
  {$ENDIF}
  {$IFDEF VER150}
  Variants,
  {$ENDIF}
  ExtCtrls,
  ComCtrls,
  Buttons,
  Menus,
  HEnt1, // pour le TActionFiche
  Hctrls,
  HFLabel,
  HSysMenu,
  HTB97,
  HPanel,
  UiUtil,
  HRichEdt,
  HRichOLE,
  Ent1,
  ZlibAuto,
  SaisUtil,
  TZ,
  ZTypes,
  ZFolio,
  ZCompte,
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  UTOB,
  ULibEcriture,
  ZDevise, // saisie dans autre monnaie que pivot
{$IFDEF COMPTAAVECSERVANT}
  EntImo, FicheImo, 
{$ENDIF}
{$IFDEF AMORTISSEMENT}
  IMMO_TOM,
{$ENDIF}
 Dialogs, TntComCtrls, TntButtons, TntStdCtrls, TntExtCtrls, TntGrids
;



type RParFolio = record
     ParPeriode,             // '01/03/1999'
     ParCodeJal,             // 'ODE'
     ParNumFolio        : string ;  // '77'
     ParNumLigne        : Integer ; // 3
     ParRecupLettrage   : boolean ; // si true on ouvre la saisie de bordereau pour les ecritures de types L
     ParGuide           : string ;
     ParCreatCentral    : boolean ; // appel en creation depuis le journal centralisateur
     ParCentral         : boolean ; // appel en modif depuis le journal centralisateur
     ParEta             : string ;
     end ;


procedure SaisieFolio(Action : TActionFiche) ;
procedure ChargeSaisieFolio(Params : RParFolio; Action : TActionFiche) ;
procedure LanceSaisieFolio(Q : TQuery; Action : TActionFiche ; FromCentral : boolean = false ) ;
procedure LanceSaisieFolioL(Q : TQuery; Action : TActionFiche) ;
procedure LanceSaisieFolioGuide( M : RMVT ) ;
procedure LanceSaisieFolioOBM(O : TOB ; Action : TActionFiche) ;
{$IFDEF SCANGED}
procedure SaisBorMyAfterImport (Sender : TObject; FileGuid: string; var Cancel: Boolean) ;
{$ENDIF}


type
  TFSaisieFolio = class(TForm)
    GS              : THGrid;
    PEntete         : THPanel;
    E_JOURNAL       : THValComboBox;
    H_JOURNAL       : THLabel;
    H_DATECOMPTABLE : THLabel;
    H_NUMEROPIECE   : THLabel;
    H_DEVISE        : THLabel;
    PPied           : THPanel;
    G_LIBELLE       : THLabel;
    H_SOLDE         : THLabel;
    T_LIBELLE       : THLabel;
    SA_SOLDEG       : THNumEdit;
    SA_SOLDET       : THNumEdit;
    SA_FOLIODEBIT   : THNumEdit;
    SA_TOTALCREDIT  : THNumEdit;
    SA_SOLDE        : THNumEdit;
    POPS            : TPopupMenu;
    POPZ            : TPopupMenu;
    FindSais        : TFindDialog;
    Bevel2          : TBevel;
    Bevel3          : TBevel;
    Bevel4          : TBevel;
    HMTrad          : THSystemMenu;
    DockTop: TDock97;
    DockRight: TDock97;
    DockLeft: TDock97;
    Dock: TDock97;
    Valide97: TToolbar97;
    BValide: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BAide: TToolbarButton97;
    Outils: TToolbar97;
    BSolde: TToolbarButton97;
    BComplement: TToolbarButton97;
    BChercher: TToolbarButton97;
    BMenuZoom: TToolbarButton97;
    ToolbarSep971: TToolbarSep97;
    ToolbarSep972: TToolbarSep97;
    Sep97: TToolbarSep97;
    HConf: TToolbarButton97;
    ISigneEuro: TImage;
    LSA_SOLDEG: THLabel;
    LSA_SOLDET: THLabel;
    LSA_SOLDE: THLabel;
    LSA_TOTALCREDIT: THLabel;
    LSA_TOTALDEBIT: THLabel;
    E_DATECOMPTABLE: THValComboBox;
    E_NATUREPIECE: THValComboBox;
    SA_TOTALDEBIT: THNumEdit;
    Bevel1: TBevel;
    SA_FOLIOCREDIT: THNumEdit;
    Bevel5: TBevel;
    LSA_FOLIODEBIT: THLabel;
    LSA_FOLIOCREDIT: THLabel;
    HLabel1: THLabel;
    E_NUMEROPIECE: THValComboBox;
    HLabel2: THLabel;
    BEVELSOLDEPROG: TBevel;
    HLSOLDEPROG: THLabel;
    SA_SOLDEPROG: THNumEdit;
    LSA_SOLDEPROG: THLabel;
    BVentil: TToolbarButton97;
    BEche: TToolbarButton97;
    SELGUIDE: TEdit;
    BZoomJournal: THBitBtn;
    BZoom_FB19024: THBitBtn;
    FlashGuide: TFlashingLabel;
    BLibAuto: TToolbarButton97;
    BInsert: TToolbarButton97;
    BSDel: TToolbarButton97;
    BTools: TToolbarButton97;
    BZoomTiers: THBitBtn;
    BZoomDevise: THBitBtn;
    BZoomEtabl: THBitBtn;
    FLASHDEVISE: TFlashingLabel;
    BTGuide: TToolbarButton97;
    RichEdit: THRichEditOLE;
    BModifTva: TToolbarButton97;
    BSelectFolio: TToolbarButton97;
    BHideFocus: TButton;
    LabelInfos: THLabel;
    BZoomCpte: THBitBtn;
    BDevise: TToolbarButton97;
    H_EXERCICE: THLabel;
    BMonnaie: THBitBtn;
    BModifRIB: THBitBtn;
    LabelInfos2: THLabel;
    E_NUMEROPIECEC: THValComboBox;
    BZoomImmo: THBitBtn;
    BGuide: THBitBtn;
    H_ETABL: THLabel;
    E_ETABLISSEMENT: THValComboBox;
    BAutoSave: TToolbarButton97;
    BZoomEcrs: THBitBtn;
    HLSOLDEPROGCB: THLabel;
    BZoomJalCentral: THBitBtn;
    BLettrage: TToolbarButton97;
    BScan: TToolbarButton97;
    SD: TSaveDialog;
    BSSCANGED: THBitBtn;
    BAcc: TToolbarButton97;
    YTC_SCHEMAGEN: THLabel;
    BAccactif: THBitBtn;
    BAccTiers: THBitBtn;
    SA_SoldePerGen: THNumEdit;
    SA_SOLDEAUX: THNumEdit;
    SoldeD: TLabel;
    SoldeAuxper: TLabel;
    PQuantite: TPanel;
    GQte: TGroupBox;
    H_QTE1: THLabel;
    H_QTE2: THLabel;
    H_QUALIFQTE1: THLabel;
    H_QUALIFQTE2: THLabel;
    _QTE1: THNumEdit;
    _QTE2: THNumEdit;
    BPDF: THBitBtn;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure E_JOURNALChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure GereAffSolde(Sender: TObject);
    procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSElipsisClick(Sender: TObject);
    procedure GSEnter(Sender: TObject);
    procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSKeyPress(Sender: TObject; var Key: Char);
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure BValideClick(Sender: TObject);
    procedure E_DATECOMPTABLEChange(Sender: TObject);
    procedure E_NUMEROPIECEExit(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure BSDelClick(Sender: TObject);
    procedure BInsertClick(Sender: TObject);
    procedure BSoldeClick(Sender: TObject);
    procedure BVentilClick(Sender: TObject);
    procedure BEchClick(Sender: TObject);
    procedure BComplementClick(Sender: TObject);
    procedure BZoom_FB19024Click(Sender: TObject);
    procedure BZoomJournalClick(Sender: TObject);
    procedure BZoomTiersClick(Sender: TObject);
    procedure BMenuZoomMouseEnter(Sender: TObject);
    procedure POPSPopup(Sender: TObject);
    procedure GSMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure GSExit(Sender: TObject);
    procedure FindSaisFind(Sender: TObject);
    procedure BChercherClick(Sender: TObject);
    procedure BLibAutoClick(Sender: TObject);
    procedure BZoomDeviseClick(Sender: TObject);
    procedure BZoomEtablClick(Sender: TObject);
    procedure BTGuideClick(Sender: TObject);
    procedure BModifTvaClick(Sender: TObject);
    procedure BSelectFolioClick(Sender: TObject);
    procedure BAbandonClick(Sender: TObject);
    procedure BHideFocusEnter(Sender: TObject);
    procedure PFENMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure BZoomCpteClick(Sender: TObject);
    procedure BDeviseClick(Sender: TObject);
    procedure E_NUMEROPIECEKeyPress(Sender: TObject; var Key: Char);
    procedure BModifRIBClick(Sender: TObject);
    procedure BZoomImmoClick(Sender: TObject);
    procedure E_ETABLISSEMENTChange(Sender: TObject);
    procedure GSMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure BZoomEcrsClick(Sender: TObject);
    procedure BLettrageClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure E_JOURNALKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure PEnteteEnter(Sender: TObject);
    procedure GSKeyDown(Sender: TObject; var Key: Word;Shift: TShiftState);
    procedure BScanClick(Sender: TObject);
    procedure BAccClick(Sender: TObject);
    procedure BAccactifClick(Sender: TObject);
    procedure BAccTiersClick(Sender: TObject);
    procedure PQuantiteExit(Sender: TObject);
    procedure _QTE1Exit(Sender: TObject);
    procedure _QTE1KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure _QTE1Enter(Sender: TObject);
    procedure BPDFClick(Sender: TObject);
  private
    memJal               : RJal ;
    memPivot             : RDevise ;
    memDevise            : RDevise ;
    FParams              : RSDev ;
    memOptions           : ROpt ;
    FExoVisu             : TExoDate ;
    memParams            : RPar ;
    NumFolio             : LongInt ;
    bFirst, bExo         : Boolean ;
    ObjFolio             : TFolio ;
    GridX, GridY         : Integer ;
    bGuideRun            : Boolean ;
    bGuideStop           : Boolean ;
    GuideFirstRow        : LongInt ;
    GuideLastRow         : LongInt ;
    LastModePaie         : string ;
    bStopDevise          : Boolean ;
    CurJal               : string ;
    CurNumFolio          : string ;
    CurPeriode           : string ;
    bClosing             : Boolean ;
    bInit                : Boolean ;
    InitParams           : RParFolio ;
//    bModeOppose          : Boolean ;
    SoldeProg            : Double ;
    nbLignes             : LongInt ;
    nbLastSave           : LongInt ;
    MaxNum               : LongInt ;
    GuideCalcRow         : LongInt ;
    GuideCalcCol         : LongInt ;
    bModeRO              : Boolean ;
    FindFirst            : Boolean ;
    bDelDone             : Boolean ;
    WMinX, WMinY         : Integer ;
    CbFolio              : THValComboBox ;
    bUp                  : Boolean ;
    bReading, bWriting   : Boolean ;
    bWrited              : Boolean ;
    bCloseOnF10          : Boolean ;
    bEsc                 : Boolean ;
    SoldeTresor          : Double ;
    bJalKeyUp            : Boolean ;
    bGuideDoSolde        : Boolean ;
    bEnreg               : Boolean ;
    bFormuleOK           : Boolean ;
    bGuideJal            : Boolean ;
    //LG* - 27/11/2001 - Indicateur de changement de pièce pour le multi-devises
    FbNewLigne           : Boolean;               // true sur la premiere ligne d'une piece
    FStLastError         : string;                // Dernier message d'erreur delphi
    FZlibAuto            : TZLibAuto ;            // objet de gestion des libelles automatique
    FRechCompte          : TRechercheGrille ;     // objet de recherche des comptes
    FInfo                : TInfoEcriture ;
    FBoNatureComplete    : boolean ;              // boolean indiquant si l'on doit afficher le code nature ou le libelle ( reprend le param 'SO_ZNATCOMPL' )
    FInLigneCourante     : integer ;              // indicateur de ligne courante ( pour ne pas redeclencher le rowenter, les rowexit ... )
    FStatut              : TActionFiche ;         // statut de gestion de la grille on colsultation on ne traite pas certain evenement
    FListeMessage        : HTStringList ;         // objet de gestion des messages ( dans ULibEcriture )
    FCurrentRowEnter     : integer ;              // stocke la colonne courante dans le rowenter pour ne pas declencher 2 fois l'evenement
    FZScenario           : TZScenario ;           // objet de gestion des scenario ( dans ULibEcriture )
    FLastRef             : string ;               // stocke la derniere ref saisie du bordereau
    bNoEvent             : boolean ;              // bloque le cellexit apres un cellenter dans une piece avec un compte d'immo ( on ne peut pas modifier le jour des ces pieces )
    FLastNatureGuide     : string ;               // derniere nature associe au guide
    FBoReading           : Boolean ;
    FBoPasEcart          : boolean ;
    FBoLettrageEnCours   : boolean ;
    FTOBLettrage         : TOB ;
    FBoValEnCours        : boolean ;
    FBoGrilleModif       : boolean ;
    FBoPasAna            : boolean ;
    FParam               : RParFolio ;
    FInRow               : integer ;
    FBoClickSouris       : boolean ;
    FboFenLett           : boolean ;
    FBoAccSaisie         : boolean ;
    FStGenHT             : string ;
    FBoArretAcc          : boolean ;
//    FBoRestore           : boolean ;
    FBoFaireCumul        : boolean ;
    FBoCurseurDansGrille : boolean ;
    FRdSoldeDBq          : double ;
    FRdSoldeCBq          : double ;
    FInRowCur            : integer ;
    FRdValCur            : double ;
    FStLastTextGuide     : string ;
    FStLastResultGuide   : boolean ;
    FFirstRow            : integer ; // on stocke la premiere ligne d'une piece ds un bordereau pour faire l'equilibrage en devise dessus
    FBoCutEnCours        : boolean ;
    FTOBEnCours          : TZF ;
    FBoE_NUMEROPIECEExit : boolean ;
    FBoPFENMouseDown     : boolean ;
    FBoGetFolioClick     : boolean ;
    FBoLastJal           : string ; // code du derneir journal charger. la fct ChargeJal puet etre appelle plusieur fois de suite ( alimente entete, e_numeropieceexit )
    {$IFDEF SCANGED}
    FGuidID              : string ;
{$ENDIF}
    // Fonctions ressource
    function  GetMessageRC(MessageID : Integer) : string ;
    function  PrintMessageRC(MessageID : Integer; sCaption : string='' ; StAfter: string='') : Integer ;
    // Fonctions Message
    procedure WMGetMinMaxInfo(var Msg: TMessage); message WM_GETMINMAXINFO ;
    // Fonctions utilitaires
    procedure InitFolio ;
    procedure ChargeJal(CodeJal : string ; vBoAfterZoom : boolean = false) ;
    function  GetLastJal : Boolean ;
    procedure NumeroteFolio ;
    function  VoirSoldeTreso : Boolean ;
    procedure PosLesSoldes ;
    procedure InitPivot ;
    procedure InitEtabl ;
    procedure InitSaisie(bFolio : Boolean) ;
    procedure CloseSaisie ;
    procedure InitEntete ;
    procedure InitGrid ;
    procedure InfosPied(Row : LongInt ; bAvecTouteLesInfo : boolean = true) ;
    procedure SetNum(Row : LongInt) ;
    function  GetRowDate(Row : LongInt) : TDateTime ;
    function  GetRowFirst(Piece : LongInt) : LongInt ;
    procedure PutRowMontant(Col, Row : LongInt; Montant : Double) ;
    procedure EnableButtons( Ou : integer=-1) ;
    procedure NextNature(Row : LongInt) ;
    function  InitNature : string ;
    function  IncRefAuto(RefI : string): string ;
    procedure GetCompteAuto(Indice : Integer) ;
    Function  ChoixCpteAuto ( LesCptAuto : String ) : String ;
    procedure AlimenteEntete(ParPeriode, ParCodeJal, ParNumFolio : string) ;
    // Outils comptabilité
    // Fontions TZF
    function  SetRowTZF(Row : LongInt ; bCalculerContrepartie : boolean = true) : TZF ;
    function  GetRowTZF(Row : LongInt; bInit : Boolean=FALSE) : Boolean ;
    procedure SetRowLib(Row : LongInt) ;
    // Saisie complémentaires
    function  IsEch(Row : LongInt) : Boolean ;
    procedure GetEch(Row : LongInt; bAutoOpen, bAutoQuiet : Boolean) ;
    procedure ModifRefAnaPiece(Row: LongInt; NomChamp: string; Val: Variant; Piece : string) ;
    procedure ModifRefAna(Row: LongInt; NomChamp: string; Val: Variant) ;
    procedure GereAna(Row : LongInt ; bOpen : Boolean ) ;
//    procedure PutAna(Row : LongInt) ;
    procedure GetAna(Row : LongInt ) ;
    procedure PutRowProrata(Row : LongInt) ;
    procedure PutRowImmo(Row : LongInt) ;
    function  GetGuideColStop(Row : LongInt) : LongInt ;
    function  GetGuideValue(Col, Row : LongInt; bMove : Boolean=TRUE) : Boolean ;
    procedure GetGuideCalcRow(Row : LongInt) ;
    procedure PutRegime(Row : LongInt) ;
    procedure PutRegimePiece(Row : LongInt; Regime, Piece : string) ;
    function  IsJalBqe : Boolean ;
    function  IsJalAutorise(CodeJal : string) : Boolean ;
    procedure GetCPartie(Row : LongInt; Piece : string; var NumCompte, NumAux : string) ;
    procedure GereComplements(Row: integer);
    //Rib
    function  PutRib(Col, Row : LongInt; bAux : Boolean ; vBoForce : boolean = false) : Boolean ;
    procedure PutConso(Row : LongInt) ;
    // Calculs
    procedure CalculSoldeJournal ;
    procedure CalculSoldeCompte(NumCompte : string ; Col : Integer ; vTotDebit, vTotCredit : double) ;
    function  IsSoldeJournal : Boolean ;
    function  IsSoldeFolio : Boolean ;
    procedure SoldeProgressif(Row : LongInt ) ;
    function  SoldeLaLigne(Piece, DefRow : LongInt; var SP, SD : Double) : Boolean ;
    function  VerifNature(Piece, DefRow : LongInt ; vBoForce : boolean = false) : LongInt ;
    function  VerifTreso(Piece, DefRow : LongInt) : Boolean ;
    function  VerifEquilibre(Piece, DefRow : LongInt) : Boolean ;
    function  IsSoldePiece(Piece : LongInt) : Boolean ;
    function  SoldePiece(Piece : LongInt) : Double ;
    function  SoldePieceDevise(Piece : LongInt;  var CumulD, CumulC : Double) : Double ;
    procedure ChangeTauxPiece(Piece : LongInt; LDevise: RDevise) ;
    function  GetRowFromPiece(Row, Piece : LongInt) : LongInt ;
    function  GetAnySolde(Piece : LongInt; var SP,  SD : Double) : LongInt ;
    procedure SetPiece ;
    function  GetRowFormule(var Formule : hstring ) : Integer ;
    function  GetRowAbs(CurRow, ForRow: LongInt) : LongInt ;
    function  GetRowRel(CurRow, ForRow: LongInt) : LongInt ;
    function  GetFormule(Formule : hstring) : Variant ;
    function  GetFormulePourCalc(Formule: hstring): Variant;
    // Fonctions Grid
    procedure SetGridRO ;
    procedure SetGridOn(bMove : Boolean=TRUE; NumLigne : Integer=0; bFocus : Boolean=FALSE) ;
    procedure SetGridOptions(Row : LongInt) ;
    procedure CreateRow(Row : LongInt; bRowEcart : Boolean=FALSE) ;
    function  DeleteRow(Row : LongInt ; bPasLigneEcart : boolean=false) : Integer ;
    function  NextRow(Row : LongInt) : Boolean ;
    function  IsRowValid(Row : LongInt; var ACol : Integer; bTest : Boolean ; bAvecMontant : boolean=false) : Boolean ;
    procedure PostDrawCell(ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
    procedure SetGridGrise(ACol, ARow : Integer ; Canvas : TCanvas) ;
    procedure SetGridSep(ACol, ARow : Integer ;  Canvas : TCanvas ; bHaut : Boolean) ;
    function  IsPieceCanEdit(Piece: string;bAvecEcartdeConversion : boolean=true)  : Boolean ;
    function  IsRowCanEdit(Row : LongInt;bAvecEcartdeConversion : boolean=true) : integer ;
    procedure SetColVal(ACol : LongInt; NumPiece, Val : string) ;
    function  IsRowCollectif(Row : LongInt) : Boolean ;
    function  IsRowTiers(Row : LongInt) : Boolean ;
    function  IsRowHT(Row : LongInt) : Boolean ;
    function  IsRowBqe(Row : LongInt) : Boolean ;
    function  IsRowDecEnc(Row : LongInt) : string ;
    function  GetRowTiers(Piece, DefRow : LongInt) : LongInt ;
    function  IsJalLibre : Boolean ;
    // SQL
    procedure WriteFolio ;
    procedure WriteCumulCpte ;
    procedure WriteCumulAno ;
    function  ReadFolio(bRestore : Boolean) : Boolean ; // lecture du folio en base ( on utilise ZFolio defini dans ZFolio.pas )
//    function  RestoreFolio : Boolean ; // restore le folio, apres plantage, à partir du fichier de sauvegarde texte
    procedure FillComboFolio ;
//    function  BackupFolio ( vBoForce : boolean = false ) : Boolean ; // enregistre le folio dans un fichier texte chaque fois que le solde passe à zero
    // Click boutons
    function  ValideFolio(bVerbose : Boolean=TRUE; bCanClose : Boolean=FALSE; bOnlySave: Boolean=FALSE) : Boolean ;
    procedure SearchClick ;
    Procedure AbandonClick ( FromESC : boolean ) ;
    function  ValClick(bVerbose : Boolean=TRUE; bCanClose: Boolean=FALSE) : Boolean ;
//    procedure ByeClick ;
    procedure DelClick ;
    procedure InsClick ;
    procedure SoldeClick(Row : LongInt) ;
    procedure ParamTvaClick ;
    procedure ComplClick ;
    procedure RefLibAutoClick ;
    procedure EchClick ;
    procedure AnaClick ;
    procedure DevClick(bTestTaux : Boolean=FALSE) ;
    procedure RibClick ;
    procedure ImmoClick ;
    procedure GetFolioClick ;
    function  ExisteGuide : Boolean ;
    procedure GuiClick ;
    procedure CGuideRun ;
    procedure GuideStop ;
    // Fonctions Zoom
//    function  BeforeZoom(var ACol, ARow: LongInt) : Boolean ;
  {$IFDEF AMORTISSEMENT}
    procedure AfterZoom(ACol, ARow: LongInt) ;
  {$ENDIF}
    procedure ZoomCpte ;
//    procedure ZoomEcrs ;
    procedure ZoomJal ;
    procedure ZoomTiers ;
    procedure ZoomDevise ;
    procedure ZoomEtabl ;
    procedure ZoomNavCpte( Compte : string = '' ; Aux : string = '') ;
    // Fonction scénario
    procedure LettrageEnSaisie( Ou : integer = - 1 ; ALaLigne : boolean = false ) ;
    {Fonctions JLD}
    procedure GotoEntete ;
    procedure IncRefBOR( Increment :integer = 1 ) ;
    procedure DupZoneBOR ;

    procedure SetStatut ( Value : TActionFiche ); // affecte la statut de la grille de saisie
    procedure OnError (sender : TObject; Error : TRecError ) ; // fonction d'affichage des erreurs pour le FRechCompte
    function  GetCellNature(ARow : integer) : string ; // retourne le code nature de la colonne courante ( on affiche sois le code nature ou le libelle )
    function  GetCompte : TZCompte ; // retourne le l'objet compte du FRechCompte  (permet de ne pas modifier tous le code existant )
    function  GetTier : TZTier ; // idem pour les tiers
    function FinSaisieBor( vStCompte : string ) : boolean; // solde le folio sur le compte passe en parametre
    function SoldeAuto : boolean ; // solde automaitiquement le bordereau
    procedure AppelElipsisClick ( vBoAvecDeplacement : boolean = true ) ; // appel des fenetre lookup sur la nature, generaux,aux
    procedure DupPremierLignePiece ;
    procedure DupLibelleEtRef;
    procedure RecupLibelleCompte;
    procedure FinGuide;
    procedure AfficheTitreAvecCommeInfo( vInfo : string ='' ) ;
    procedure RenumAna ;
    procedure SetGrilleModif( Value : boolean ) ;
    function  AccelerateurDeSaisie ( Ou : integer = - 1 ) : integer;
    function  GetCompteAcc( Ou : integer ; E_AUXILIAIRE : string ) : string ;
    function  IsSoldeFolioPourAcc ( Ou : integer ) : Boolean ;
    procedure SoldeDuCompte( D, C : double ) ;
    procedure SoldeAux( D, C : double ) ;
    function  GetLib ( vRow, vRowRef : integer ) : string;
    procedure GereCutOff(ARow: integer);
    procedure SupprimeEcrCompl ( vTOB : TOB ) ;
    procedure AnnuleCutOffSiBesoin( vTOB : TOB ) ;
    procedure MajInfoCompte ( ARow,ACol : integer ) ;
    procedure FiniAux(ACol, ARow : integer );
    procedure SetNewLigne( Value : boolean ) ;
    procedure GereQuantite( Ecr : TZF ) ;
    function  EstAccelere(Ou : integer) : boolean;
    procedure AffecteDeviseSaisie ;
    {$IFDEF TT}
    procedure GestionException ( Sender : TObject ; E : Exception);
    {$ENDIF}
{$IFDEF SCANGED}
    procedure SetGuidId ( vGuidId : string ) ;
{$ENDIF}
    property Comptes : TZCompte read GetCompte ;
    property Tiers   : TZTier  read GetTier ;
    property Statut : TActionFiche read FStatut write SetStatut;
    property GrilleModif : boolean read FBoGrilleModif write SetGrilleModif ;
    property bNewLigne : boolean read FbNewLigne write SetNewLigne ;
  public
    FAction : TActionFiche ;
    procedure EffacerListeMemoire ;
{$IFDEF SCANGED}
    function GetInfoLigne : string ;
    property GuidID : string  read FGuidID write SetGuidID ;
{$ENDIF}
    property ValidationEnCours : boolean read FBoValEnCours ;
  end;

function GetFolioCourant : TFSaisieFolio ;
function CBordereauBloque : boolean ;


implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPProcGen,
  CPProcMetier,
  CPVersion,
  {$ENDIF MODENT1}
  ZCentral,
 {$IFDEF SCANGED}
  YNewDocument,
  UtilGed,
  UGedFiles,
  UGedViewer,
  AnnOutils,
  cbpPath ,
 {$ENDIF}
 Constantes,
 AGLInit,
 LettUtil,
 ULibAnalytique,
 ParamSoc,
 ULibWindows,
 Choix,
 SaisComm,
 UtilPGI,
 SaisEnc,
 Lookup, // pour le libelle automatique
 UtilSais,
 ZEch, // saisie mono echeance
 Formule, // pour le GFormule
 SaisComp,
 ZGuide,
 FichComm, // pour ouvrir la fenetre etablissement
 eSaisAnal,
 CPJournal_TOM,
 CPTiers_TOM,
 CPGeneraux_TOM,
 Echeance,
 hmsgbox, // pour le HShowMessage
 HStatus, // pour le FiniMove
 HCompte, // pour l'appel de l'analytique
 ZECRIMVT_TOF,
 UtilSoc,
 Devise_TOM,
 Lettrage,
{$IFNDEF CMPGIS35}
 UTOFConsEcr,
{$ENDIF}
 ULibExercice ;

const
      cInMaxEcr           = 15000 ;
      MAX_BOUCLE          = 4 ;

const NRC_NOJOURNAL      = '0;Comptabilité;Veuillez créer au moins un journal de type Bordereau ou Libre;W;O;O;O;' ;
      NRC_NOPERIODE      = '0;Comptabilité;Toutes les périodes sont clôturées entre le %s et le %s.'+#10#13+'Changez votre exercice de référence...;W;O;O;O;' ;
      NRC_NOJALAUTORISE  = '0;Comptabilité;Aucun journal autorisé pour cette saisie;W;O;O;O;' ;
      NRC_NOACTIVE       = '0;Comptabilité;Vous ne pouvez pas accéder à cette fonction;W;O;O;O;' ;
      NRC_NOETABL        = '0;Comptabilité;L''établissement par défaut n''est pas renseigné;W;O;O;O;' ;
      NRC_NOSAISIE       = '0;Comptabilité;La date d''entrée est incompatible avec la saisie.;W;O;O;O;' ;
      NRC_NOJOURNALCCMP  = '0;Comptabilité;Veuillez créer au moins un journal de type Bordereau ou Libre de nature Caisse ou Banque.;W;O;O;O;' ;


const SF_ETABL  : Integer = 0 ;
      SF_QUALIF : Integer = 1 ;
      SF_NEW    : Integer = 2 ;
      SF_EURO   : Integer = 3 ;
      SF_PIECE  : Integer = 4 ;
      SF_NUMO   : Integer = 5 ;
      SF_NUML   : Integer = 6 ;
      SF_NAT    : Integer = 7 ;
      SF_JOUR   : Integer = 8 ;
      SF_GEN    : Integer = 9 ;
      SF_AUX    : Integer = 10;
      SF_REFI   : Integer = 11;
      SF_LIB    : Integer = 12;
      SF_DEVISE : Integer = 13 ;
      SF_DEBIT  : Integer = 14;
      SF_CREDIT : Integer = 15;
      SF_FIRST  : Integer = 7 ;
      SF_LAST   : Integer = 15;



var
 gFolio : TFSaisieFolio ;
{$IFDEF TT}
 TheLog : TStringList;
 TheLastError : string ;


procedure AddEvenement( value : string );
begin
 if Assigned(TheLog) then
  begin
   TheLastError:=value ;
   TheLog.Add(Value) ;
   TheLog.SaveToFile('c:\SaisBor.txt') ;
  end ;
end;

procedure DelEvenement ;
begin
 TheLog.Clear ;
 TheLog.SaveToFile('c:\SaisBor.txt') ;
end;

procedure TFSaisieFolio.GestionException ( Sender : TObject ; E : Exception);
begin
 MessageAlerte( 'Erreur' + #13#10 + #13#10 + TheLastError + #13#10 + #13#10 +
               E.Message );

// AddEvenement( E.Message );
end;
{$ENDIF}

function GetFolioCourant : TFSaisieFolio ;
begin
 result := nil ;
 if gFolio <> nil then
  result := gFolio ;
end;

function CBordereauBloque : boolean ;
var
 lFolio : TFSaisieFolio ;
begin
 result := false ;
 lFolio := GetFolioCourant ;
 if lFolio <> nil then
  result := lFolio.ValidationEnCours ;
end ;


//=======================================================
//======= Point d'entrée dans la saisie bordereau =======
//=======================================================
{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 14/04/2004
Modifié le ... :   /  /
Description .. : - FB 11845 - VL - ds ccmp on ouvre que les journaux de
Suite ........ : type banques et caisse
Mots clefs ... :
*****************************************************************}
function CanOpenFolio(var Action : TActionFiche) : Boolean ;
var Q : TQuery ; sExo : string ;
begin
//Result:=TRUE ;
if (not (ctxPCL in V_PGI.PGIContexte)) then
  begin
  sExo:=QuelExoDt(V_PGI.DateEntree) ;
{  if (sExo=VH^.EnCours.Code) and ((V_PGI.DateEntree>VH^.Encours.Fin) or (V_PGI.DateEntree<VH^.Encours.Deb)) then sExo:='' ;
  if not ((sExo=VH^.EnCours.Code) or (sExo=VH^.Suivant.Code)) then
    begin HShowMessage(NRC_NOSAISIE, '', '') ; Result:=FALSE ; Exit ; end ;
  end ;  }
if (sExo=ctxExercice.EnCours.Code) and ((V_PGI.DateEntree>ctxExercice.Encours.Fin) or (V_PGI.DateEntree<ctxExercice.Encours.Deb)) then sExo:='' ;
  if not ((sExo=ctxExercice.EnCours.Code) or (sExo=ctxExercice.Suivant.Code)) then
    begin HShowMessage(NRC_NOSAISIE, '', '') ; Result:=FALSE ; Exit ; end ;
  end ;
if not VH^.ZACTIVEPFU then begin HShowMessage(NRC_NOACTIVE, '', '') ; Result:=FALSE ; Exit ; end ;
if VH^.EtablisDefaut='' then begin HShowMessage(NRC_NOETABL, '', '') ; Result:=FALSE ; Exit ; end ;
if not ExJaiLeDroitConcept(TConcept(ccSaisEcritures), TRUE) then begin Result:=FALSE ; Exit ; end ;
if SaisieFolioLancee then Action:=taConsult ;
if (Action<>taConsult) and (VH^.JalAutorises<>'') then
  begin
  Result:=FALSE ;
  {$IFDEF CCMP}
  Q:=OpenSQL('SELECT J_JOURNAL FROM JOURNAL WHERE J_FERME="-" AND (J_MODESAISIE="BOR" OR J_MODESAISIE="LIB") AND (J_NATUREJAL="CAI" OR J_NATUREJAL="BQE")', TRUE) ;
  {$ELSE}
  Q:=OpenSQL('SELECT J_JOURNAL FROM JOURNAL WHERE J_FERME="-" AND (J_MODESAISIE="BOR" OR J_MODESAISIE="LIB")', TRUE) ;
  {$ENDIF}
  while not Q.EOF do
    begin
    if (Pos(';'+Q.Fields[0].AsString+';', VH^.JalAutorises)>0) then
      begin Result:=TRUE ; Break ; end ;
    Q.Next ;
    end ;
  Ferme(Q) ;
  // Existe-t-il un journal de type Folio autorisé pour l'utilisateur courant ?
  if not Result then begin HShowMessage(NRC_NOJALAUTORISE, '', '') ; Exit ; end ;
  end else
  begin
 {$IFDEF CCMP}
  Result:=ExisteSQL('SELECT J_MODESAISIE FROM JOURNAL WHERE J_FERME="-" AND (J_MODESAISIE="BOR" OR J_MODESAISIE="LIB") AND (J_NATUREJAL="CAI" OR J_NATUREJAL="BQE")') ;
  // Existe-t-il un journal de type Folio ?
  if not Result then begin HShowMessage(NRC_NOJOURNALCCMP, '', '') ; Exit ; end;
 {$ELSE}
  Result:=ExisteSQL('SELECT J_MODESAISIE FROM JOURNAL WHERE J_FERME="-" AND (J_MODESAISIE="BOR" OR J_MODESAISIE="LIB")') ;
  // Existe-t-il un journal de type Folio ?
  if not Result then begin HShowMessage(NRC_NOJOURNAL, '', '') ; Exit ; end ;
 {$ENDIF}
  end ;
// Existe-t-il une période de saisie ?
//if (VH^.DateCloturePer>0) and (FinDeMois(VH^.Entree.Fin)<=VH^.DateCloturePer) then
//  begin HShowMessage(Format(NRC_NOPERIODE, [DateToStr(VH^.Entree.Deb), DateToStr(VH^.Entree.Fin)]), '', '') ; Result:=FALSE ; exit; end ;
// Existe-t-il une période de saisie au delà de la date de révision ?
if RevisionActive(FinDeMois(ctxExercice.Entree.Fin)) then begin Result:=FALSE ; exit; end;
Result:=TRUE ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 26/06/2006
Modifié le ... :   /  /
Description .. : - LG  - FB 18437 - 26/06/2006 - utilisation d'une variable
Suite ........ : globale pour creer le folio pour pouvoir appel la saisie
Suite ........ : bordereau facil d'un autre module
Mots clefs ... :
*****************************************************************}
procedure SaisieFolio(Action : TActionFiche) ;
var PP : THPanel ; lFolio : TFSaisieFolio ;
begin
if not CanOpenFolio(Action) then Exit ;
if _Blocage(['nrCloture'], TRUE, 'nrSaisieCreat') then Exit ;
lFolio:=TFSaisieFolio.Create(Application) ;
if gFolio = nil then
 gFolio:=lFolio ;
lFolio.bInit:=FALSE ;
lFolio.FAction:=Action ;
lFolio.bCloseOnF10:=TRUE ;
PP:=FindInsidePanel ;
if PP=nil then
   begin
   try
     lFolio.ShowModal ;
   finally
     lFolio.Free ;
   end ;
  Screen.Cursor:=SyncrDefault ;
  end else
  begin
  InitInside(lFolio, PP) ;
  lFolio.Show ;
  end ;  
end ;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 26/06/2006
Modifié le ... :   /  /
Description .. : - LG  - FB 18437 - 26/06/2006 - utilisation d'une variable
Suite ........ : globale pour creer le folio pour pouvoir appel la saisie
Suite ........ : bordereau facil d'un autre module
Mots clefs ... :
*****************************************************************}
procedure ChargeSaisieFolio(Params : RParFolio; Action : TActionFiche) ;
var PP : THPanel ; lFolio : TFSaisieFolio ;
begin
if not CanOpenFolio(Action) then Exit ;
lFolio:=TFSaisieFolio.Create(Application) ;
if gFolio = nil then
 gfolio := lFolio ;
lFolio.bInit:=true ;
lFolio.FAction:=Action ;
lFolio.bCloseOnF10:=TRUE ;
FillChar(lFolio.InitParams, Sizeof(lFolio.InitParams), #0) ;
lFolio.InitParams:=Params ;
PP:=FindInsidePanel ;
if PP=nil then
   begin
   try
     lFolio.ShowModal ;
   finally
     lFolio.Free ;
   end ;
  Screen.Cursor:=SyncrDefault ;
  end else
  begin
  InitInside(lFolio, PP) ;
  lFolio.Show ;
  end ;
end ;

Function QOK(Q : tQuery ; Simul : String3)  : Boolean ;
Var St : String ;
BEGIN
Result:=TRUE ;
If (ctxPCL in V_PGI.PGIContexte) Then Exit ;
St:='' ;
If Q.FindField('E_JOURNAL')=NIL Then St:='E_JOURNAL' ;
If Q.FindField('E_DATECOMPTABLE')=NIL Then St:='E_DATECOMPTABLE' ;
If Q.FindField('E_NUMEROPIECE')=NIL Then St:='E_NUMEROPIECE' ;
If Q.FindField('E_NUMLIGNE')=NIL Then St:='E_NUMLIGNE' ;
If Simul<>'' Then BEGIN If Q.FindField('E_QUALIFPIECE')=NIL Then St:='E_QUALIFPIECE' ; END ;
If St<>'' Then
  BEGIN
  Result:=FALSE ;
  HShowMessage('0;Opération impossible !;Le champ '+ST+' doit faire partie des colonnes affichées;E;O;O;O;','','') ;
  HShowMessage('0;Opération impossible !;Vous devez le rajouter dans le paramétrage de cette liste;E;O;O;O;','','') ;
  END ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 17/08/2004
Modifié le ... :   /  /
Description .. : - LG - 17/08/2004 - suppression de la fct DebutDeMois, ne
Suite ........ : fct pas avec les execercice decale
Mots clefs ... :
*****************************************************************}
procedure LanceSaisieFolio(Q : TQuery; Action : TActionFiche ; FromCentral : boolean = false ) ;
var Params : RParFolio ;
begin
FillChar(Params, Sizeof(Params), #0) ;
If Not QOK(Q,'') Then Exit ;
Params.ParPeriode:=DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime)  ;
Params.ParCodeJal:=Q.FindField('E_JOURNAL').AsString ;
Params.ParNumFolio:=Q.FindField('E_NUMEROPIECE').AsString ;
Params.ParCentral:=FromCentral ;
Params.ParNumLigne:=Q.FindField('E_NUMLIGNE').AsInteger ;
Params.ParRecupLettrage:=false ;
ChargeSaisieFolio(Params, Action) ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 17/08/2004
Modifié le ... :   /  /
Description .. : - LG - 17/08/2004 - suppression de la fct DebutDeMois, ne
Suite ........ : fct pas avec les execercice decale
Mots clefs ... :
*****************************************************************}
procedure LanceSaisieFolioL(Q : TQuery; Action : TActionFiche) ;
var Params : RParFolio ;
begin
FillChar(Params, Sizeof(Params), #0) ;
If Not QOK(Q,'') Then Exit ;
Params.ParPeriode:=DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime) ;
Params.ParCodeJal:=Q.FindField('E_JOURNAL').AsString ;
Params.ParNumFolio:=Q.FindField('E_NUMEROPIECE').AsString ;
Params.ParNumLigne:=Q.FindField('E_NUMLIGNE').AsInteger ;
Params.ParRecupLettrage:=true ;
ChargeSaisieFolio(Params, Action) ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : LG
Créé le ...... : 25/04/2006
Modifié le ... :   /  /
Description .. : - LG - FB 17926 - on recupere la date comptable envoyer
Suite ........ : par la saisie guider
Mots clefs ... :
*****************************************************************}
procedure LanceSaisieFolioGuide( M : RMVT ) ;
var Params : RParFolio ;
lAction : TActionFiche ;
begin
FillChar(Params, Sizeof(Params), #0) ;
Params.ParPeriode:= DateToStr(M.DateC) ;     //DateToStr(V_PGI.DateEntree) ;
Params.ParCodeJal:=M.Jal;
Params.ParNumFolio:='1' ;
Params.ParNumLigne:=1 ;
Params.ParGuide:=M.LeGuide ;
Params.ParRecupLettrage:=false ; lAction := taModif ;
ChargeSaisieFolio(Params, lAction) ;
end ;



{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 17/08/2004
Modifié le ... : 15/02/2005
Description .. : - LG - 17/08/2004 - suppression de la fct DebutDeMois, ne
Suite ........ : fct pas avec les execercice decale
Suite ........ : - SG6 - 15/02/2005 - changement des parametres de la fct
Suite ........ : O passe d'un TOBM vers TOB
Suite ........ :
Mots clefs ... :
*****************************************************************}
procedure LanceSaisieFolioOBM(O : TOB ; Action : TActionFiche) ;
var Params : RParFolio ;
begin
FillChar(Params, Sizeof(Params), #0) ;
Params.ParPeriode:=DateToStr(O.GetValue('E_DATECOMPTABLE')) ;
Params.ParCodeJal:=O.GetValue('E_JOURNAL') ;
Params.ParNumFolio:=IntToStr(O.GetValue('E_NUMEROPIECE')) ;
Params.ParNumLigne:=O.GetValue('E_NUMLIGNE') ;
ChargeSaisieFolio(Params, Action) ;
end ;

{$IFDEF SCANGED}
procedure SaisBorMyAfterImport (Sender : TObject; FileGuid: string; var Cancel: Boolean) ;
var LastError : String;
    lTobDoc, lTobDocGed : TOB;
    lDocGuid : string;
begin

 if  (ctxPCL in V_PGI.PGIContexte) and not (JaileDroitConceptBureau (187315)) then exit ;

 lTobDoc := Tob.Create('YDOCUMENTS', nil, -1) ;
 lTobDoc.LoadDb ; // Pas de clé = Charge une clé à 0
 lTobDocGed := Tob.Create('DPDOCUMENT', nil, -1) ;
 lTobDocGed.LoadDb ;

 try

  lTobDoc.PutValue('YDO_LIBELLEDOC'    , TFSaisieFolio(Sender).GetInfoLigne) ; // PAR EXEMPLE !!
  lTobDoc.PutValue('YDO_MOTSCLES'      , 'ECRITURE') ; // En majuscules, séparés par des ;
  lTobDoc.PutValue('YDO_ANNEE'         , FormatDateTime('yyyy', Date)) ;

  // TobDocGed : NODOSSIER
  lTobDocGed.PutValue('DPD_NODOSSIER', V_PGI.NoDossier) ;

  lDocGuid := InsertDocumentGed(lTobDoc, lTobDocGed, FileGuid, LastError) ;

 finally
  lTobDoc.Free ;
  lTobDocGed.Free ;
 end ;

 // Fichier refusé, suppression dans la GED
 if lDocGuid = '' then // consulter éventuellement LastError
  V_GedFiles.Erase(FileGuid)
   else
    TFSaisieFolio(Sender).GuidID := lDocGuid ; // pour traitement par l'intéressé
end ;

procedure AjouterFichierDansGed ;
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

 SaisBorMyAfterImport(GetFolioCourant,lStFileGuid,lBoCancel) ;

end ;


{$ENDIF}


//=======================================================
//================= Fonctions Ressource =================
//=======================================================
function TFSaisieFolio.GetMessageRC(MessageID : Integer) : string ;
begin
result:=CGetMessageRC(MessageID,FListeMessage);
end ;

function TFSaisieFolio.PrintMessageRC(MessageID : Integer; sCaption : string ; StAfter : string) : Integer ;
begin
result:=CPrintMessageRC(Caption,MessageID,FListeMessage) ;
end ;

//=======================================================
//================== Fonctions Message ==================
//=======================================================
procedure TFSaisieFolio.WMGetMinMaxInfo(var Msg: TMessage) ;
begin
with PMinMaxInfo(Msg.lparam)^.ptMinTrackSize do begin X:=WMinX ; Y:=WMinY ; end ;
end;

//=======================================================
//================ Evénements de la Form ================
//=======================================================
{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 04/11/2002
Modifié le ... : 25/03/2004
Description .. : - 04/11/2002 - creation de la stringlist de debug ici plus que
Suite ........ : dans la initialisation
Suite ........ : - 25/03/2004 - Ajout du boolean indiquant que l'on est
Suite ........ : rentrer en cliquant ou tabulation
Mots clefs ... :
*****************************************************************}
procedure TFSaisieFolio.FormCreate(Sender: TObject);
begin
{$IFDEF TT}
 TheLog:=TStringList.Create;
 TheLog.SaveToFile('c:\SaisBor.txt');
 Application.OnException := GestionException ;
 VH^.OkModGed:=true ;// activation du scan en saisie
 V_PGI.FSAV := false ;
 VH^.ZSAISIEECHE:=true ;
{$ENDIF}
FTOBLettrage:=TOB.Create('',nil,-1) ; FBoClickSouris := false ;
bUp:=FALSE ; bReading:=FALSE ; bWriting:=FALSE ;  
WMinX:=Width ; WMinY:=Height div 2 ;
ObjFolio:=nil ;
// Guide
bGuideRun:=FALSE ; bGuideStop:=FALSE ; bGuideDoSolde:=FALSE ;
GuideFirstRow:=-1 ; GuideLastRow:=-1 ; GuideCalcRow:=-1 ; GuideCalcCol:=-1 ;
bClosing:=FALSE ; bEsc:=FALSE ; bJalKeyUp:=FALSE ; bEnreg:=FALSE ;
// Paramètres courant
CurJal:='' ; CurPeriode:='' ; CurNumFolio:='' ; FBoLastJal := '' ; 
// Paramètres de saisie
//memParams.bAnalytique:=VH^.ZGEREANAL ;
//memParams.bAnaAuto:=VH^.ZSAISIEANAL ;
//memParams.bEcheance:=VH^.ZSAISIEECHE ;
memParams.bLibre:=FALSE ;
memParams.bDevise:=FALSE ; memParams.DevLibre:='' ;
// Paramètres DEBUG
memParams.bSoldeProg:=FALSE ;
memParams.bRealTime:=VH^.ZFOLIOTEMPSREEL ;
// Mode ?
CbFolio:=E_NUMEROPIECE ;
// Initialisations diverses
InitPivot ;
//InitEtabl ;
InitGrid ;
PosLesSoldes ;
// Mode opposé
//bModeOppose:=FALSE ;
if FAction=taCreat then FAction:=taModif ;
// Toolbar
RegLoadToolbarPos(Self, 'SaisBor') ;
//LG* 22/02/2002
FZlibAuto:=TZLibAuto.Create ; Statut:=taConsult ; FZScenario:=TZScenario.Create ;
FRdValCur:= - 999999 ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 26/06/2002
Modifié le ... : 21/08/2007
Description .. : -25/06/2002 - gestion de la nature abrege ou complete
Suite ........ : - init de la valeur contenant la ligne courante
Suite ........ : 
Suite ........ : - 19/07/2002 - nouvelle gestion des messages avec le
Suite ........ : CInitMessageBor
Suite ........ : - nouvelle objet de recherche des comptes FRechCompte
Suite ........ : 
Suite ........ : - 17/09/2002 -  test si l'on doit recreer le folio à partir du
Suite ........ : fichier de suavegarde locale on ne lance pas le test si on
Suite ........ : est en consultation
Suite ........ : - 03/09/2004 - FB 14230 - on bloque le numeroexit en 
Suite ........ : restoration
Suite ........ : - 06/04/2006 - FB 17694 - on acces du folio, on se 
Suite ........ : positionne sur la derniere ligne
Suite ........ : - 28/04/2006 - FB 17875 - en consultation depuis la 
Suite ........ : consultation des compte on se place sur la ligne 
Suite ........ : selectionnee
Suite ........ : - 20/07/2007 - FB 21128 - le code exo n'etait pas affecte
Suite ........ : - FB 21022 - LG - 21/08/2007 - calcul du total folio au 
Suite ........ : demarrage 
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.FormShow(Sender: TObject);
{$IFNDEF EAGLCLIENT}
var
 lQ     : TQuery ;
 lStMil : string ;
{$ENDIF}
begin
{$IFDEF CCMP}
  {JP 03/08/05 : FQ 15922 : Seulement si on est en création}
  if FAction = taCreat then
    E_JOURNAL.Plus := 'AND (J_NATUREJAL="CAI" OR J_NATUREJAL="BQE")';
{$ENDIF}
//LG* 08/01/2002 ClearLock  ;
{$IFDEF TT}
 V_PGI.Confidentiel:='0' ;
 VH^.OkModImmo := true ;
{$ENDIF}
 InitEtabl ;
 If FAction=taModif Then
  BEGIN
  // GP le 20/08/2008 21496
  RetoucheHVCPoursaisie(E_JOURNAL) ;
  RetoucheHVCPoursaisie(E_ETABLISSEMENT) ;
  END ;

 VH^.ZSAUVEFOLIOLOCAL := EstSpecif('51204') ;
 if V_PGI.OkDecP > 9 then
  V_PGI.OkDecP := 9 ;
 SoldeAuxper.caption := '' ;
 SoldeD.Caption      := '' ;
 FStLastTextGuide    := '' ;
 FStLastResultGuide  := false ;
 FBoLettrageEnCours  := InitParams.ParRecupLettrage ;
 FBoNatureComplete   := not GetParamSocSecur('SO_ZNATCOMPL',false) ;
// initialisation de l'objet de recherche ( avant le readFolio qui l'utilise )
 FInfo               := TInfoEcriture.Create ;
 FRechCompte         := TRechercheGrille.Create(FInfo) ;
 FRechCompte.OnError :=OnError ;
 FRechCompte.COL_GEN :=SF_GEN ;
 FRechCompte.COL_AUX :=SF_AUX ;
 FInLigneCourante    := -1 ;
 FListeMessage       :=HTStringList.Create ;

 CInitMessageBor(FListeMessage) ;

 if not FBoNatureComplete then
  begin
   GS.ColWidths[SF_NAT] := 35;
   GS.Cells[0,SF_NAT]   := TraduireMemoire('Nat');
  end ;
 BVentil.Visible     := VH^.ZGEREANAL ;
 BEche.Visible       := VH^.ZSAISIEECHE ;
 PQuantite.Visible   := false ;
{$IFDEF SCANGED}
 BScan.visible       := True ;
{$ENDIF}

 if not ExJaiLeDroitConcept(TConcept(ccSaisEcritures), FALSE) then FAction:=taConsult ;
 LookLesDocks(Self) ;
 InitSaisie(FALSE) ;
 bModeRO := (FAction = taConsult) ;

 if FAction = taConsult then
  begin
   E_NUMEROPIECEC.Height   := E_NUMEROPIECE.Height ;
   E_NUMEROPIECEC.Width    := E_NUMEROPIECE.Width ;
   E_NUMEROPIECEC.Visible  := TRUE ;
   E_NUMEROPIECE.Visible   := FALSE ;
   CbFolio                 := E_NUMEROPIECEC ;
  end
   else
    CbFolio := E_NUMEROPIECE ;

 if CbFolio.CanFocus and not bModeRO then
  CbFolio.SetFocus ;


 if bInit then
  begin
   // ouverture du bordereau depuis un autre module ( consultation des comptes ... )
   // on ouvre directement le bordereau avec les valeurs
   E_NUMEROPIECE.OnExit := nil ;
   if InitParams.ParCreatcentral then  // on demande la creation d'un nouveux folio depuis le journal centralisateur
    memOptions.NewObj := true
     else
      begin
       AlimenteEntete(InitParams.ParPeriode, InitParams.ParCodeJal, InitParams.ParNumFolio) ;
       if( not ReadFolio(FALSE) ) then
        begin
         CbFolio.SetFocus ;
         exit ;
        end ;
      end ;
   E_NUMEROPIECE.OnExit := E_NUMEROPIECEExit ;
   CalculSoldeJournal ;   // le calcul du solde du folio est fait trop souvent Fb 21022
   // Activer le Grid
   if (InitParams.ParNumLigne>0) then
    SetGridOn(TRUE, InitParams.ParNumLigne, TRUE)
     else
      SetGridOptions(GS.Row) ;

   EnableButtons ;

  end ;  // bInit

 //lg* 21/02/2002  force le redimensionnement de la grille
 HMTrad.ResizeGridColumns(GS);
 //LG* chargement de l'ensemble des libelles
 FZlibAuto.Load ;
 //CbFolio.SetFocus ; FB 13106 supprimer, je sais plus a quoi ca sert

 // on regarde le nombre d'ecriture ds la table DPTABCOMPTA, si > a 15000
 // on bloque le calcul des soldes des comptes
 FBoFaireCumul := false ;

 FRechCompte.Info.TypeExo := memOptions.TypeExo ;  // FB 21128

 {$IFNDEF EAGLCLIENT}

 case memOptions.TypeExo of
  teEncours   : lStMil:='N' ;
  teSuivant   : lStMil:='N+' ;
  tePrecedent : lStMil:='N-' ;
 end ;

 lQ := OpenSql('select dtc_nbecriture from DPTABCOMPTA ' +
              ' where dtc_nodossier="' + V_PGI.NoDossier + '" ' +
              ' and dtc_millesime="' +  lStMil + '" ' , true ) ;

 FBoFaireCumul := lQ.FindField('dtc_nbecriture').asInteger < cInMaxEcr ;
 ferme(lQ) ;
{$ENDIF}

if InitParams.ParGuide <> '' then
 begin
  E_NUMEROPIECEExit(nil) ;
  SELGUIDE.Text := InitParams.ParGuide ;
  CGuideRun ;
 end
  else
 //  if InitParams.ParCreatCentral then
 //   E_NUMEROPIECEExit(nil)
 //    else
      if InitParams.ParCentral then
       begin
        NextRow(nbLignes+1) ;
        GS.Row := nbLignes ;
       end ;
{$IFDEF SCANGED}
{
  CA - 20/02/2007 - Si la GED n'est pas sérialisée, on autorise quand même
                    l'insertion de documents en saisie
}
  if not VH^.OkModGed then
  begin
    if V_PGI.RunFromLanceur then
      InitializeGedFiles(V_PGI.DefaultSectionDbName, SaisBorMyAfterImport)
    else
      InitializeGedFiles(V_PGI.DbName, SaisBorMyAfterImport);
  end;
{$ENDIF}
end;

procedure TFSaisieFolio.FormClose(Sender: TObject; var Action: TCloseAction);
begin
{$IFDEF SCANGED}
  // Ajout CA - 20/02/2007 - Dans le cas où la GED n'est pas sérialisée
  if not VH^.OkModGed then FinalizeGedFiles;
{$ENDIF}

//LG* suppression de llock
if FBoE_NUMEROPIECEExit or FBoPFENMouseDown or FBoGetFolioClick then
 begin
  {$IFDEF TT}
   AddEvenement('FormClose exitNumero') ;
  {$ENDIF}
  exit ;
 end ;
{$IFDEF TT}
AddEvenement('FormClose');
{$ENDIF}

bClosing:=TRUE ;
try
CBloqueurJournal(false,E_JOURNAL.Value) ;
if FTOBLettrage<>nil then FTOBLettrage.Free ;
if FZlibAuto<>nil then begin FZlibAuto.Free ; FZlibAuto:=nil ; end ;
if FZScenario<>nil then begin FZScenario.Free ; FZScenario:=nil ; end ;
if assigned(FInfo) then begin FInfo.Free ; FInfo:=nil;  end ;
if assigned(FRechCompte) then begin FRechCompte.Free ; FRechCompte:=nil ; end ;
FListeMessage.Free ; FListeMessage:=nil;
CloseSaisie ;
RegSaveToolbarPos(Self, 'SaisBor') ;
//LG* 08/01/2002 suppression du blocage
_Bloqueur('nrSaisieCreat',False) ;
ClearLock ;
{$IFDEF TT}
  TheLog.Free;
  TheLog := nil ;
{$ENDIF}
Self.caption:='' ;
UpdateCaption(self) ;
V_PGI.IoError := oeOK ;
if ( Parent is THPanel ) then Action:=caFree ;
finally
 FTOBLettrage:=nil ; FZlibAuto:=nil ; FZScenario:=nil ; FRechCompte:=nil ; FListeMessage:=nil; FInfo:=nil ; gFolio:= nil ;
end ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 19/06/2002
Modifié le ... : 20/04/2004
Description .. : suppression d'une fuite memoire
Suite ........ : 
Suite ........ : - 11/07/2003 - FB 12513 - ajout d'un boolean
Suite ........ : FBoValeEnCours pour empecher la sortie avant la fin de
Suite ........ : l'enregistrement
Suite ........ : - 20/04/2004 - LG - deplacement du test sur le lettrage en
Suite ........ : saisie pour ne pas avoir deux fois la question si on sort par
Suite ........ : escape
Mots clefs ... :
*****************************************************************}
procedure TFSaisieFolio.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
var Ecr, Tmp : TZF ;
begin
if FBoE_NUMEROPIECEExit or FBoValEnCours or bReading or FBoPFENMouseDown  or FBoGetFolioClick then
 begin
  {$IFDEF TT}
  AddEvenement('FormCloseQuery exitNumero') ;
  {$ENDIF}
  CanClose := false ;
  exit ;
 end ;
{$IFDEF TT}
AddEvenement('FormCloseQuery') ;
{$ENDIF}

if FAction=taConsult then begin CanClose:=TRUE ; Exit ; end ;

{$IFNDEF TT1}
if FBoLettrageEnCours and ( not bModeRO ) then
 begin
   // LG 22/10/2003 - il faut absolument enregistrer la bordereau si on a commence un lettrage en saisie
   CanClose:=false ;
   PrintMessageRC(RC_ABANDONINTERDIT) ;
   exit ;
 end;
{$ENDIF}

if bEsc then begin CanClose:=TRUE ; Exit ; end ;

if bGuideRun then
   if (PrintMessageRC(RC_GUIDESTOP)=mrYes)
      then begin  GuideStop ; CanClose:=FALSE ; Exit ; end
      else begin CanClose:=FALSE ; Exit ; end ;

Tmp:=nil ;
if (ObjFolio<>nil) and (nbLignes>1) then
   begin
   Tmp:=TZF.Create('ECRITURE', nil, -1) ;
   SetRowTZF(GS.Row,false) ;
   Ecr:=ObjFolio.GetRow(GS.Row-GS.FixedRows) ;
   if Ecr<>Nil then Tmp.Assign(Ecr) ;
   end ;
CanClose:=True ; 
if ((NbLignes>1) and GrilleModif and (PrintMessageRC(RC_ABANDON)<>mrYes)) then
   begin
   CanClose:=FALSE ;
   if (GS.Row>nbLignes) then
       begin
       CreateRow(GS.Row) ;
       if Tmp<>nil then
          begin
          Ecr:=ObjFolio.GetRow(GS.Row-GS.FixedRows) ;
          Ecr.Assign(Tmp) ; GetRowTZF(GS.Row-GS.FixedRows) ;
          end ;
       GS.Invalidate ;
       end ;
   END ;
 if assigned(Tmp) then Tmp.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 24/09/2004
Modifié le ... : 28/10/2004
Description .. : - LG - 24/09/2004 - FB 13097 - on se replace ds l'entete 
Suite ........ : apres un f12
Suite ........ : - LG - 28/10/2004 - on ne retourne pas ds l'entete si la grille 
Suite ........ : est en modif
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.GotoEntete ;
BEGIN
 if GrilleModif then exit ;
 if Not PEntete.Enabled then Exit ;
 CurNumFolio              := '' ;
 InitFolio ;
 CbFolio.Enabled          := true ;
 E_JOURNAL.Enabled        := true ;
 E_DATECOMPTABLE.Enabled  := true ;
 if E_DATECOMPTABLE.CanFocus then E_DATECOMPTABLE.SetFocus ;
 FBoCurseurDansGrille     := false ;
 GrilleModif              := false ;
 FBoClickSouris           := false ;
 bESC                     := false ;
 EnableButtons ;
 {
if E_JOURNAL.CanFocus then E_JOURNAL.SetFocus else
 if E_DATECOMPTABLE.CanFocus then E_DATECOMPTABLE.SetFocus else
  if E_ETABLISSEMENT.CanFocus then E_ETABLISSEMENT.SetFocus ;
  }
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 09/08/2002
Modifié le ... : 03/12/2002
Description .. : - 09/08/2002 - pour la reference interne on duplique le 
Suite ........ : dernier numero saisie
Suite ........ : - 04/11/2002 - passe la grille en modif sur F7
Suite ........ : -07/11/2002 - la fct duplique la chose du haut ( plus 
Suite ........ : d'increment de la referece )
Suite ........ : - 03/12/2002 - fiche 10800 - apres un F7 on se place à la 
Suite ........ : fin du mot
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.DupZoneBOR ;
BEGIN
if bModeRO then Exit ;
if Not (GoEditing in GS.Options) then Exit ;
if ((GS.Col<>SF_GEN) and (GS.Col<>SF_AUX) and (GS.Col<>SF_REFI) and (GS.Col<>SF_LIB) and
    (GS.Col<>SF_DEBIT) and (GS.Col<>SF_CREDIT) and (GS.Col<>SF_NAT) and (GS.Col<>SF_JOUR)) then Exit ;
if GS.Row<=1 then Exit ;
GS.Cells[GS.Col,GS.Row]:=GS.Cells[GS.Col,GS.Row-1] ;
GS.InplaceEditor.SelStart:=Length(GS.Cells[GS.Col,GS.Row]) ; // se place à la fin du mot
GrilleModif:=true ;
Statut:=taModif ;
END ;

function _ValeurPremiereLignePiece( GS : THGrid ) : string ;
var
 i : integer ;
 lStPiece : string ;
begin
 i:=GS.Row ; lStPiece:=GS.Cells[SF_PIECE,GS.Row] ;//result:=GS.Cells[ACol,i] ; if i=1 then exit ;
 Dec(i) ;
 while (i>1) and (GS.Cells[SF_PIECE,i]=lStPiece) do Dec(i) ;
 result:=GS.Cells[GS.Col,i+1] ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 03/12/2002
Modifié le ... :   /  /
Description .. : fiche 10800 - reproduction de la cellule de la 1 ligne de la
Suite ........ : piece
Mots clefs ... :
*****************************************************************}
procedure TFSaisieFolio.DupPremierLignePiece;
begin
 if bModeRO then Exit ; if Not (GoEditing in GS.Options) then Exit ; if GS.Row<=1 then Exit ;
 if memParams.bLibre then
  BEGIN
   GS.Cells[GS.Col,GS.Row]:=GS.Cells[GS.Col,1] ; Statut:=taModif ;
   exit ;
  END;
 GS.Cells[GS.Col,GS.Row]:=_ValeurPremiereLignePiece(GS) ;
 GS.InplaceEditor.SelStart:=Length(GS.Cells[GS.Col,GS.Row]) ; // se place à la fin du mot
 GrilleModif:=true ;
 Statut:=taModif ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 03/12/2002
Modifié le ... :   /  /    
Description .. : - fiche 10799 - reproduction du libelle et le refrence de la
Suite ........ : ligne du dessus
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.DupLibelleEtRef ;
BEGIN
if bModeRO then Exit ;
if Not (GoEditing in GS.Options) then Exit ;
if GS.Row<=1 then Exit ;
GS.Cells[SF_REFI,GS.Row]:=GS.Cells[SF_REFI,GS.Row-1] ; GS.Cells[SF_LIB,GS.Row]:=GS.Cells[SF_LIB,GS.Row-1] ;
GrilleModif:=true ;
Statut:=taModif ;
END ;

procedure TFSaisieFolio.RecupLibelleCompte ;
begin
if bModeRO then Exit ;
if Not (GoEditing in GS.Options) then Exit ;
SetRowLib(GS.Row) ;
GS.InplaceEditor.SelStart:=Length(GS.Cells[GS.Col,GS.Row]) ; // se place à la fin du mot
Statut:=taModif ;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 04/11/2002
Modifié le ... :   /  /
Description .. : - 04/11/2002 - Passe la grille en modif
Mots clefs ... :
*****************************************************************}
procedure TFSaisieFolio.IncRefBOR( Increment : integer ) ;
Var sRef,sNum : String ;
    ii,ll     : integer ;
BEGIN
if bModeRO then Exit ;
if GS.Col<>SF_REFI then Exit ;
sRef:=GS.Cells[SF_RefI,GS.Row] ;
if sRef='' then begin if trim(FLastRef)<>'' then sRef:=FLastRef else exit ; end ;
sNum:='' ; ii:=Length(sRef) ; ll:=0 ;
While ((ii>=1) and (sRef[ii] in ['0'..'9'])) do BEGIN sNum:=sRef[ii]+sNum ; Dec(ii) ; inc(ll) ; END ;
if sNum='' then Exit ;
sNum:=IntToStr(ValeurI(sNum)+Increment) ; if ValeurI(sNum)<1 then sNum:='1' ;
sRef:=Copy(sRef,1,Length(sRef)-ll)+sNum ;
GS.Cells[SF_REFI,GS.Row]:=sRef ; Statut:=taModif ; GrilleModif:=true ; FLastRef := sRef ;
END ;


// Lorsque la fonction est associée à un bouton (ToolBar), c'est la
// propriété Enabled de ce bouton qui autorise ou non l'accès à cette fonction.
{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 26/06/2002
Modifié le ... : 19/10/2006
Description .. : - 25/06/2002 - pour les comptes auto, sur la premiere ligne
Suite ........ : on incremente le compte,sinon on repart du premier
Suite ........ : -02/10/2002 - test si l'ecriture est modifiable pour le
Suite ........ : changement de nature
Suite ........ : -09/10/2002 - utilisation du F12 pour passer dans l'entete à
Suite ........ : la place du crtl home
Suite ........ : -16/10/2002 - le F6 pour solde la ligne ne fct pas
Suite ........ : - 06/11/2002 - correction du ctrl+fin, on ne se deplace que
Suite ........ : si la grille est en consultation ( le rowenter va renvoyer une
Suite ........ : erreur )
Suite ........ : - on ne gere les comptes auto que sur la case des generaux
Suite ........ : - alt + Z pour agrandir la fenetre
Suite ........ : - correctoin du ctrl + fin
Suite ........ : - 13/11/2002 - changement de methode pour detecter que
Suite ........ : la touche de la touche Ctrl est enfonce
Suite ........ : - 28/11/2002 - on affiche plus la fenetre des comptes sur le
Suite ........ : F5
Suite ........ : - CA - 05/02/2003 - Pour ne pas saisir de racine dans un
Suite ........ : guide, on impose le déplacement avec TAB.
Suite ........ : -02/12/2002 - fiche 10905 - appel de la calculatrice sur le +
Suite ........ : dans les cases debit/credit
Suite ........ : -17/12/2002- suprression du alt Z integre dans l'agl
Suite ........ : -20/01/2002 - fiche 10791 - on ne relance pas le dernier
Suite ........ : guide tant que l'on est en mode guide
Suite ........ : - FB 11801 - on peut faire suppr unqiuement en nature
Suite ........ : abrege
Suite ........ : - FB 11851 - test precedent ne fct pas
Suite ........ : - FB 12623 - CTRL+G pour les guids (suite conflit avec
Suite ........ : scan)
Suite ........ : - LG 03/05/2004 - correction d'un plantage qd on fermait la
Suite ........ : fenetre a appuyant sur f10 en meme temps
Suite ........ : - LG - 19/10/2006 - FB 18878 - on interdit les pieces
Suite ........ : multidevise en echange vers S1
Mots clefs ... :
*****************************************************************}
procedure TFSaisieFolio.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var Vide : Boolean ;
    lStateCtrl : short;
begin
if csDestroying in ComponentState then Exit ;
if not GS.SynEnabled then begin Key:=0 ; Exit ; end ;
if bClosing or bReading then begin Key := 0 ; Exit ; end ;  Application.ProcessMessages ;
{$IFDEF TT}
AddEvenement('FormKeyDown ' + intToStr(Key) );
{$ENDIF}
FRechCompte.InKey:=Key ; FRechCompte.Shift:=Shift ;
if  ( (Screen.ActiveControl=_QTE1) or (Screen.ActiveControl=_QTE2) ) and (Key=VK_ESCAPE) then
 begin
  Key := 0 ;
  NextPrevControl(self) ;
  exit ;
 end ;

if Not (Screen.ActiveControl=GS) then
   BEGIN
    if (Key=VK_ESCAPE) then
    begin
      Key:=0 ;
      AbandonClick(True) ;
      exit;
    end
    else
     if Key = VK_RETURN then begin SendMessage(Self.Handle,WM_KEYDOWN,VK_TAB,0) ; exit ; end
      else
       Exit ; //or (NbLignes<=0))
   END ;
Vide:=(Shift=[]) ;
if (Vide) and (GS.Col=SF_NAT) and (Key in [VK_DELETE,VK_BACK]) and FBoNatureComplete  then begin Key:=0 ;Exit ; end ;
lStateCtrl:=GetKeyState(VK_CONTROL); // test si la touche Ctrl est enfonce
if (lStateCtrl<0) and (GS.Col=SF_GEN) and (Key in [96..105]) then GetCompteAuto(Key-96) ;
case Key of
  VK_RETURN : if (Vide) then BEGIN Key:=VK_TAB ; FRechCompte.InKey:=Key ; END ;
  VK_F5     : if Shift=[ssShift] then begin Key:=0 ; BZoomCpteClick(nil) ; exit ; end
              else
              if (not bModeRO) and ((GS.Col=SF_NAT) or (GS.Col=SF_GEN) or (GS.Col=SF_AUX)) then
               BEGIN
                 Key:=0 ;
                 AppelElipsisClick ;
               END ;
   VK_SPACE : if (Vide) and (not bModeRO) and (GS.Col=SF_NAT) and (IsRowCanEdit(GS.Row)<=0) then begin Key:=0 ; NextNature(GS.Row) ; end ;
  VK_INSERT : if (Vide) then begin Key:=0; InsClick ; end ;
  VK_DELETE : if  Shift=[ssCtrl] then begin Key:=0 ; DelClick ; end ;
  VK_ESCAPE : if Vide then begin Key:=0 ; AbandonClick(True) ; end ;
     VK_F6  : if Vide then
               begin
                Key:=0 ;
                if (not bGuideRun) and (trim(GS.Cells[SF_GEN, GS.Row])<>'') then SoldeClick(-1) // F6 solde la ligne
                 else
                  if (trim(MemJal.CContreP)<>'') and ( not bGuideRun ) then
                   FinSaisieBor(memJal.CContreP) // solde sur le compte de contrepartie
                    else
                     if (trim(memjal.CAutomat)<>'') and ( not bGuideRun ) then
                      FinSaisieBor(TrouveAuto(memjal.CAutomat,1)) ; // solde sur le 1 compte auto
               end ;
     VK_F7  : if Vide then BEGIN Key:=0 ; DupZoneBOR ; END
              else if Shift=[ssShift] then BEGIN Key:=0 ; DupPremierLignePiece ; END
              else if Shift=[ssAlt] then BEGIN Key:=0 ; DupLibelleEtRef ; END ;
     VK_F8  : if (Shift=[ssCtrl]) then
               begin
                Key := 0 ;
                {$IFDEF SCANGED}
                AjouterFichierDansGed() ;
                {$ENDIF}
               end
                else
                 SendMessage(Application.Handle,VK_F8,0,0) ;
     VK_F10 :  if Vide then
                 begin
                  Key:=0 ;
                  ValClick(FALSE) ;
                 end ;
    VK_F9 : InfosPied(GS.Row) ;
    VK_F12 :  begin Key:=0 ; if (not bGuideRun) then GotoEntete ; end ;
    VK_END : if (Shift=[ssCtrl]) and (Statut=taConsult) then
               begin // on ne place a la fin quand consultation
                Key:=0 ;
                GS.ElipsisButton := false ;
                SetGridOptions(nbLignes) ;
                GS.Row:=nbLignes ;
                GS.col:=SF_GEN ;
               end;
// CA - 05/02/2003
// Pour ne pas saisir de racine dans un guide, on impose le déplacement avec TAB.
    VK_UP : if bGuideRun and (GS.Col<>SF_AUX) then Key :=0;
    VK_DOWN : if bGuideRun then Key :=0;
// Fin CA - 05/03/2003
VK_F11      : begin Key:=0 ; PopS.PopUp(Mouse.CursorPos.X,Mouse.CursorPos.Y) ; end ;
VK_F4: begin
              if bModeRO then Exit ; //LG* 25/06/2002
              if GS.Row=FInLigneCourante then
               begin // sur la meme ligne on increment le compte auto
                Inc(memJal.CurAuto) ; if memJal.CurAuto>memJal.NbAuto then memJal.CurAuto:=1 ;
                GetCompteAuto(memJal.CurAuto) ; Key:=0 ;
               end // if
                else
                 begin
                  FInLigneCourante:=GS.Row ; GetCompteAuto(1) ; Key:=0 ;
                 end;
              end ;
VK_F3  : begin
              if bModeRO then Exit ; //LG* 25/06/2002
              if GS.Row=FInLigneCourante then
               begin // sur la meme ligne on increment le compte auto
                Dec(memJal.CurAuto) ; if memJal.CurAuto<1 then memJal.CurAuto:=memJal.NbAuto ;
                GetCompteAuto(memJal.CurAuto) ; Key:=0 ;
               end
                else
                 begin // sur une nouvelle ligne on prend le premier
                  FInLigneCourante:=GS.Row ; GetCompteAuto(1) ; Key:=0 ;
                 end;
              end ;
         48 : if (not bModeRO) and (Shift=[ssAlt]) then begin GetCompteAuto(Key-48) ; Key:=0 ; end ;
     49..57 : if (not bModeRO) and (Shift=[ssAlt]) then begin GetCompteAuto(Key-48) ; Key:=0 ; end ;
{CTRL+A} 65 : if Shift=[ssCtrl] then begin Key:=0 ; AnaClick ; end ;
{ALT+C}  67 : if Shift=[ssAlt] then begin Key:=0 ; ComplClick ; end ;
{CTRL+D} 68 : if Shift=[ssCtrl] then
                 begin
                 Key:=0 ;
                 if ( GetParamsocSecur('SO_CPLIENGAMME','') = 'S1' ) then
                  begin
                   if ( GS.Row = 1 ) then
                    BDeviseClick(Nil) ;
                  end
                   else
                    BDeviseClick(Nil) ;
                 end ;
{ALT+E}  69 : if Shift=[ssAlt] then begin Key:=0 ; EchClick ; end else
{CTRL+E}      if Shift=[ssCtrl] then
                 begin
                 Key:=0 ; DevClick(TRUE) ;
                 bStopDevise:=TRUE ;
                 PostMessage(GS.Handle, WM_KEYDOWN, VK_TAB, 0) ;
                 end ;
{CTRL+F} 70 : if Shift=[ssCtrl]       then begin Key:=0 ; SearchClick ; end ;
{ALT+G}  71 : if Shift=[ssAlt]        then begin Key:=0 ; GuiClick ; end
              else  { CA - 11/09/2003 }
                if Shift=[ssCtrl] then
                  if (not bGuideRun) then begin Key:=0 ; CGuideRun ; end ;
{ALT+I}  73 : if Shift=[ssAlt]        then begin Key:=0 ; ZoomDevise ; end ;
{ALT+J}  74 : if Shift=[ssAlt]        then begin Key:=0 ; ZoomJal ; end ;
{ALT+Q}  81 : if (Shift=[ssAlt]) and PQuantite.Visible and _QTE1.CanFocus then
               begin
                Key := 0 ;
                _QTE1.SetFocus ;
               end ;
{ALT+S}  83 : if Shift=[ssCtrl]       then begin Key:=0 ; RecupLibelleCompte ; end ;
{CTRL+L} 76 : if Shift=[ssCtrl]       then begin Key:=0 ; LettrageEnSaisie ; end
              else
              if Shift=[ssAlt]        then begin Key:=0 ; RefLibAutoClick ; end ;
{CTRL+O} 79 : if Shift=[ssCtrl]       then begin Key:=0 ; ImmoClick ; end ;
{ALT+R}  82 : if Shift=[ssAlt]        then begin Key:=0 ; RibClick ; end
                else if Shift=[ssCtrl]  then begin Key:=0 ; RecupLibelleCompte ; end;
{ALT+T}  84 : if Shift=[ssAlt]        then begin Key:=0 ; ZoomEtabl ; end ;
{CTRL -}109 : if Shift=[ssCtrl]       then begin Key:=0 ; IncRefBOR(-1) ; end ;

{CTRL +}107 : {if Vide and ( (GS.Col=SF_DEBIT) or (GS.Col=SF_CREDIT) ) then
               begin
                key:=0 ; AppelCalculatrice ;
               end
                else}
                 if Shift=[ssCtrl] then begin Key:=0 ; IncRefBOR ; end ;
         90 : if Shift=[ssCtrl] then begin Key:=0 ; FStGenHT:=GS.Cells[SF_GEN, GS.Row] ; end ;
{CTRL+W} 87 : if (Shift=[ssCtrl]) then
               begin
                Key:=0 ;
                AccelerateurDeSaisie ;
               end ; // necessecaire de remettre Key:=0, gere le cas ou l'on a ppuie plusieur de fois de suite sur ctrl+w en sur un auxi en milieu de grille ( sinon generation de ligne vide qd on tabule )
(*{^Z}    90 :  if Shift=[ssAlt] then
              BEGIN Key:=0 ; if WindowState=wsMaximized then WindowState:=wsNormal else WindowState:=wsMaximized END ; //LG**)
  end ;


end;

//=======================================================
//================ Evénements des contrôles =============
//=======================================================
{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 10/02/2003
Modifié le ... : 18/03/2004
Description .. : - LG - 10/02/2003 - FB 11854 - qd on change de journal,
Suite ........ : on remet à zero le dernier guide utilise
Suite ........ : - LG - 03/03/2004 - on recharge systematiquement le
Suite ........ : journal qd on choisis apres avoir annule le chagement de
Suite ........ : journal
Suite ........ : - LG - 08/03/2004 - on passe sytematiquement dans le
Suite ........ : traitement
Suite ........ : - LG - 18/03/2004 - réécriture suite au changement de 
Suite ........ : gestion
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.E_JOURNALChange(Sender: TObject);
var CodeJal : string ; i : Integer ;
begin
if bClosing then Exit ;
if not E_JOURNAL.Enabled then Exit ;

if (E_JOURNAL.Value='') then
  begin  Exit ; end ;
if CurJal=E_JOURNAL.Value then Exit ;
FParam.ParCodeJal:=CurJal ;
if not IsJalAutorise(E_JOURNAL.Value) then
  begin
  E_JOURNAL.Enabled:=FALSE ;
  PrintMessageRC(RC_NOJALAUTORISE, GetMessageRC(RC_SAISIEECR)) ;
  if bJalKeyUp then
    begin
    for i:=E_JOURNAL.ItemIndex-1 downto 0 do
      if IsJalAutorise(E_JOURNAL.Values[i]) then
        begin E_JOURNAL.Value:=E_JOURNAL.Values[i] ; Break ; end ;
    end else
    begin
    for i:=E_JOURNAL.ItemIndex+1 to E_JOURNAL.Items.Count-1 do
      if IsJalAutorise(E_JOURNAL.Values[i]) then
        begin E_JOURNAL.Value:=E_JOURNAL.Values[i] ; Break ; end ;
    end ;
    if not IsJalAutorise(E_JOURNAL.Value) then
      begin
      for i:=0 to E_JOURNAL.Items.Count-1 do
        if IsJalAutorise(E_JOURNAL.Values[i]) then
          begin E_JOURNAL.Value:=E_JOURNAL.Values[i] ; Break ; end ;
      end ;
  E_JOURNAL.Enabled:=TRUE ;
  E_JOURNAL.SetFocus ;
  Exit ;
  end ;
CodeJal:=E_JOURNAL.Value ;
if CodeJal='' then
   begin
   E_JOURNAL.Value:=CurJal ;
   //memJal.Code:='' ; CbFolio.Enabled:=FALSE ;
   Exit ;
   end ;
// Alimenter le record RJal
ChargeJal(CodeJal) ;
CalculSoldeJournal ;
// Mise à jour du combo Folio
FillComboFolio ;
// Attribuer un numéro de folio
NumeroteFolio ;
GS.Enabled:=FALSE ;
// Utiliser la bonne nature de pièce en fonction de la nature du journal
case CaseNatJal(memJal.Nature) of
   tzJVente       : E_NATUREPIECE.DataType:='ttNatPieceVente' ;
   tzJAchat       : E_NATUREPIECE.DataType:='ttNatPieceAchat' ;
   tzJBanque      : E_NATUREPIECE.DataType:='ttNatPieceBanque' ;
   tzJEcartChange : E_NATUREPIECE.DataType:='ttNatPieceEcartChange' ;
   tzJOD          : E_NATUREPIECE.DataType:='ttNaturePiece' ;
   end ;
CurJal:=E_JOURNAL.Value ;
FLastNatureGuide:='' ;
end;

procedure TFSaisieFolio.E_JOURNALKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
if Key=VK_UP then bJalKeyUp:=TRUE else bJalKeyUp:=FALSE ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 03/03/2004
Modifié le ... : 18/03/2004
Description .. : - 03/02/2004 - LG - changement de gestion du maxjour : on
Suite ........ : peut saisir des date d'exo diff du 1 et derneir jour du mois
Suite ........ : - LG - 03/03/2004 - on recharge systematiquement le 
Suite ........ : journal qd on choisis apres avoir annule le chagement de 
Suite ........ : journal- 
Suite ........ : - LG - 18/03/2004 - réécriture suite au changement de 
Suite ........ : gestion 
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.E_DATECOMPTABLEChange(Sender: TObject);
var Year, Month, Day : Word ;
begin
if bClosing then Exit ;
if not E_DATECOMPTABLE.Enabled then Exit ;
if (CurPeriode=E_DATECOMPTABLE.Value) and (CurPeriode<>'') then Exit ;
FParam.ParPeriode:=CurPeriode ;
// Renseigner les options
DecodeDate(StrToDate(E_DATECOMPTABLE.Value), Year, Month, Day) ;
memOptions.MaxJour:=CGetDaysPerMonth(E_DATECOMPTABLE.Value, FExoVisu ) ;
memOptions.Year:=Year ;
memOptions.Month:=Month ;
ChargeJal(memJal.Code) ;
CalculSoldeJournal ;
// Remplir la combo
FillComboFolio ;
// Attribuer un numéro de folio
NumeroteFolio ;
GS.Enabled:=FALSE ;
CurPeriode:=E_DATECOMPTABLE.Value ;
Tiers.RazSolde ; Comptes.RazSolde ;
end;

procedure TFSaisieFolio.AffecteDeviseSaisie ;
begin
if bClosing then Exit ;
FInfo.Devise.Load([V_PGI.DevisePivot]) ;
FInfo.Devise.AffecteTaux(EncodeDate(memOptions.Year, memOptions.Month, 1)) ;
//memSaisie.Code:=E_DEVISE.Value ;   
//if memSaisie.Code<>'' then GetInfosDevise(memSaisie) ;
//memSaisie.Taux:=GetTaux(memSaisie.Code,memSaisie.DateTaux,EncodeDate(memOptions.Year, memOptions.Month, 1)) ;
ChangeMask(SA_TOTALDEBIT,  FInfo.Devise.Dev.Decimale, FInfo.Devise.Dev.Symbole) ;
ChangeMask(SA_TOTALCREDIT, FInfo.Devise.Dev.Decimale, FInfo.Devise.Dev.Symbole) ;
ChangeMask(SA_FOLIODEBIT,  FInfo.Devise.Dev.Decimale, FInfo.Devise.Dev.Symbole) ;
ChangeMask(SA_FOLIOCREDIT, FInfo.Devise.Dev.Decimale, FInfo.Devise.Dev.Symbole) ;
ChangeMask(SA_SOLDE,       FInfo.Devise.Dev.Decimale, FInfo.Devise.Dev.Symbole) ;
ChangeMask(SA_SOLDEPROG,   FInfo.Devise.Dev.Decimale,  FInfo.Devise.Dev.Symbole) ;
ChangeMask(SA_SOLDEG,      FInfo.Devise.Dev.Decimale,  FInfo.Devise.Dev.Symbole) ;
ChangeMask(SA_SOLDET,      FInfo.Devise.Dev.Decimale,  FInfo.Devise.Dev.Symbole) ;
ChangeMask(SA_SoldePerGen, FInfo.Devise.Dev.Decimale,  FInfo.Devise.Dev.Symbole) ;
ChangeMask(SA_SoldeAux,    FInfo.Devise.Dev.Decimale,  FInfo.Devise.Dev.Symbole) ;
end;

procedure TFSaisieFolio.E_ETABLISSEMENTChange(Sender: TObject);
var i : LongInt ;
begin
if bClosing then Exit ;
HGBeginUpdate(GS);
try
 for i:=GS.FixedRows to nbLignes do GS.Cells[SF_ETABL, i]:=E_ETABLISSEMENT.Value ;
finally
 HGEndUpdate(GS);
end;
end;

procedure TFSaisieFolio.E_NUMEROPIECEKeyPress(Sender: TObject;
  var Key: Char);
begin
if (Ord(Key)>=32) and (not (Key in ['0'..'9'])) then Key:=#0 ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 20/11/2002
Modifié le ... : 24/05/2007
Description .. : -20/11/2002 - fiche 10515 - qd on change de journal, tab
Suite ........ : sur le numero
Suite ........ : piece passe dans la grille
Suite ........ : -23/01/2003 - qd on sur la grille ne pas relire le folio
Suite ........ : -01/04/2003 - FB 12209 - supression de la fiche 10515
Suite ........ : - 30/10/2003 - l'annulation du chg de numero ne fct pas
Suite ........ : avec un appel depuis la consultation des ecritures
Suite ........ : - 03/09/2004 - LG - FB 14230 - on bloque le enumeroexit 
Suite ........ : en restoration de folio
Suite ........ : - 24/05/2007 - LG - FB 19575 - on repasse le flag d'erreur a 
Suite ........ : faux qd on lit un new bordereau
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.E_NUMEROPIECEExit(Sender: TObject);
begin
if csDestroying in ComponentState then Exit ;
if bClosing then Exit ;
if FBoE_NUMEROPIECEExit then exit ;
{$IFDEF TT}
AddEvenement('*************** E_NUMEROPIECEExit entree ' );
{$ENDIF}
FBoE_NUMEROPIECEExit := true ;
try
if length(CbFolio.Text) > 10 then CbFolio.Text := '1' ;
//if not CbFolio.Enabled then Exit ;
if Valeur(CbFolio.Text)>0 then
 NumFolio:=Trunc(Valeur(CbFolio.Text))
  else
   begin
    CbFolio.Text:=IntToStr(NumFolio) ;
    if (not bClosing) then
     CbFolio.SetFocus ;
   end ;
if (CurNumFolio=CbFolio.Text) then exit ;
if Screen.ActiveControl=E_JOURNAL then Exit ;
if Screen.ActiveControl=E_DATECOMPTABLE then Exit ;
if not ValideFolio(TRUE, TRUE) then
begin
 if CurNumFolio<>'' then
  begin
  CbFolio.Text:=CurNumFolio ;
  NumFolio:=Trunc(Valeur(CbFolio.Text)) ;
 end ;
end ;
InitFolio ;
V_PGI.IoError := oeOK ; // FB 19975
ChargeJal(memJal.Code) ;
if not ReadFolio(FALSE) then
  begin
   if CbFolio.CanFocus then CbFolio.SetFocus ;
   Exit ;
  end ;
CalculSoldeJournal ;
// Activer le Grid
InfosPied(-1) ;
EnableButtons ;
SetGridOn ;
finally
 FBoE_NUMEROPIECEExit := false ;
 {$IFDEF TT}
 AddEvenement('$$$$$$$$$$$$$$$$$$$$$$E_NUMEROPIECEExit sortie ' );
 {$ENDIF}
end ;
end;

procedure TFSaisieFolio.BHideFocusEnter(Sender: TObject);
begin
// Activer le Grid
SetGridOn(FALSE) ;
end;

procedure TFSaisieFolio.GereAffSolde(Sender: TObject);
var Nam : string ; c : THLabel ;
begin
Nam:=THNumEdit(Sender).Name ; Nam:='L'+Nam ;
c:=THLabel(FindComponent(Nam)) ;
if c<>nil then c.Caption:=THNumEdit(Sender).Text ;
end;

//=======================================================
//================ Fonctions utilitaires ================
//=======================================================
procedure TFSaisieFolio.InitFolio ;
begin
CloseSaisie ;
InitPivot ;
InitGrid ;
PosLesSoldes ;
InitSaisie(TRUE) ;
end ;

procedure TFSaisieFolio.InitSaisie(bFolio : Boolean) ;
begin
nbLignes:=0 ; nbLastSave:=nbLignes ; SoldeProg:=0 ; MaxNum:=1 ;
memParams.bDevise:=FALSE ; memParams.DevLibre:='' ;
bExo:=FALSE ; bDelDone:=FALSE ;
if not bFolio then
 InitEntete ;
InfosPied(-1) ;
ZeroBlanc(PPied) ;
if (bFirst) and (not bFolio) and (bClosing) and (E_JOURNAL.Canfocus) then E_JOURNAL.SetFocus ;
ChangeMask(_QTE1,V_PGI.OkDecQ,'');
ChangeMask(_QTE2,V_PGI.OkDecQ,'');
end ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 09/01/2002
Modifié le ... :   /  /
Description .. : Modif du LockFolio
Mots clefs ... :
*****************************************************************}
procedure TFSaisieFolio.CloseSaisie ;
begin 
if ObjFolio<>nil then
   begin
   //LG* 08/01/2002 on enleve les verrrous
   ObjFolio.UnLockFolio;
   ObjFolio.Free ; ObjFolio:=nil ;
   end ;
PurgePopup(POPS) ; PurgePopup(POPZ) ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 17/01/2003
Modifié le ... : 18/09/2007
Description .. : - 17/01/2003 - plantait quand on rentrait sur un dossier vide
Suite ........ : - 29/09/2004 - LG - FB 14634 - bordereau bloque qd la
Suite ........ : pereiode est close
Suite ........ : - 05/10/2004 - LG  - le bocage du bordereau plantait en
Suite ........ : consultation des comptes
Suite ........ : - 05/01/2005 - LG - FB 14634 - si la periode de ref est 
Suite ........ : close, on passe en consult , et on ouvre la saisie avedc un 
Suite ........ : bordereau bloqué
Suite ........ : - 01/06/2006 - LG - FB 18216 - on reaafecte la date de 
Suite ........ : debut pour els excercices decalés
Suite ........ : - LG - 18/09/2007 - FB 20625 - correction pour les cloture 
Suite ........ : periodique sur les exercices decales
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.InitEntete ;
var bAllPer : Boolean ; MonoExo : Boolean ;
begin
{$IFDEF TT}
AddEvenement('InitEntete') ;
{$ENDIF}
// Bouton spécial
BHideFocus.Width:=0 ;
// Les boutons fonctionnels
BSolde.Enabled:=FALSE ;      BEche.Enabled:=FALSE ;      BVentil.Enabled:=FALSE ;
BComplement.Enabled:=FALSE ; BLibAuto.Enabled:=FALSE ;   BChercher.Enabled:=FALSE ;
BMenuZoom.Enabled:=FALSE ;   BInsert.Enabled:=FALSE ;    BSDel.Enabled:=FALSE ;
BTGuide.Enabled:=FALSE ;     BValide.Enabled:=FALSE ;    BDevise.Enabled:=FALSE ;
BModifTva.Enabled:=FALSE ;   BModifRIB.Enabled:=FALSE ;  BGuide.Enabled:=FALSE ;
BLettrage.Enabled:=FALSE ;
// Remplir la combo Période
if (bInit) and (FAction<>taConsult) then
  begin
  if QuelExoDate(StrToDate(InitParams.ParPeriode), StrToDate(InitParams.ParPeriode),MonoExo, FExoVisu) then
   memOptions.Exo:=FExoVisu.Code ;
  if (memOptions.Exo<>ctxExercice.Encours.Code) and (memOptions.Exo<>ctxExercice.Suivant.Code) then
    FAction:=taConsult ;
  if (VH^.DateCloturePer>0) and (DebutDeMois(StrToDate(InitParams.ParPeriode))<=VH^.DateCloturePer)
     and ( ctxExercice.QuelExoDate(VH^.DateCloturePer).Code = memOptions.Exo ) then // FB 20625
    FAction:=taConsult ;
  end else
  if FAction=taConsult then
    begin
    QuelExoDate(StrToDate(InitParams.ParPeriode), StrToDate(InitParams.ParPeriode),MonoExo, FExoVisu) ;
    memOptions.Exo:=FExoVisu.Code ;
    end else
    begin
    if (ctxPCL in V_PGI.PGIContexte) and ((ctxExercice.CPExoRef.Code=VH^.Encours.Code)
       or (ctxExercice.CPExoRef.Code=ctxExercice.Suivant.Code)) then begin memOptions.Exo:=ctxExercice.CPExoRef.Code ; FExoVisu:=ctxExercice.CPExoRef ; end ;
    if memOptions.Exo='' then begin memOptions.Exo:=ctxExercice.Entree.Code ; FExoVisu:=ctxExercice.Entree ; end ;
    end ;
if memOptions.Exo=ctxExercice.Encours.Code then memOptions.TypeExo:=teEncours else
   if memOptions.Exo=ctxExercice.Suivant.Code then memOptions.TypeExo:=teSuivant else
      memOptions.TypeExo:=tePrecedent ;
memOptions.QualifP:='N' ;
case memOptions.TypeExo of
  teEncours   : H_EXERCICE.Caption:='(N)' ;
  teSuivant   : H_EXERCICE.Caption:='(N+1)' ;
  tePrecedent : H_EXERCICE.Caption:='(N-1)' ;
  else          H_EXERCICE.Caption:='' ;
  end ;
bAllPer:=FALSE ;
if FAction=taConsult then bAllPer:=TRUE ;
ListePeriode(memOptions.Exo, E_DATECOMPTABLE.Items, E_DATECOMPTABLE.Values, TRUE, bAllPer) ;
if E_DATECOMPTABLE.Items.Count = 0 then
 begin
  ListePeriode(memOptions.Exo, E_DATECOMPTABLE.Items, E_DATECOMPTABLE.Values, TRUE, true) ;
  FAction:=taConsult ;
 end ;
// Positionnement par défaut
// Ancien code : E_DATECOMPTABLE.Value:=DateToStr(DebutDeMois(V_PGI.DateEntree)) ;
case memOptions.TypeExo of
  teEncours   : E_DATECOMPTABLE.ItemIndex:=E_DATECOMPTABLE.Values.IndexOf(DateToStr(DebutDeMois(V_PGI.DateEntree))) ;
  teSuivant   : E_DATECOMPTABLE.ItemIndex:=E_DATECOMPTABLE.Items.Count -1 ;
  tePrecedent : E_DATECOMPTABLE.ItemIndex:=E_DATECOMPTABLE.Items.Count -1 ;
  end ;
if E_DATECOMPTABLE.ItemIndex<0 then E_DATECOMPTABLE.ItemIndex:=E_DATECOMPTABLE.Items.IndexOf(E_DATECOMPTABLE.Items[0]) ;
E_DATECOMPTABLEChange(nil) ;
if (not bInit) and  (VH^.DateCloturePer>0) and (EncodeDate(memOptions.Year,memOptions.Month,memOptions.MaxJour) <=VH^.DateCloturePer) then
begin
 HShowMessage(Format(NRC_NOPERIODE, [DateToStr(FExoVisu.Deb), DateToStr(FExoVisu.Fin)]), '', '') ;
 FAction                 := taConsult ;
 bModeRO                 := true ;
 E_NUMEROPIECE.Enabled   := false ;
 E_NUMEROPIECEC.Enabled  := false ;
 E_JOURNAL.Enabled       := false ;
 E_DATECOMPTABLE.Enabled := false ;
 E_ETABLISSEMENT.Enabled := false ;
 exit ;
end ;
// Positionnement par défaut sur le dernier journal utilisé
if FAction<>taConsult then GetLastJal ;
// Positionnement par défaut sur la devise pivot
AffecteDeviseSaisie ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 20/09/2002
Modifié le ... : 03/11/2006
Description .. : -20/09/2002-chargement des guides correspondant aux
Suite ........ : journal
Suite ........ : - 26/05/2003 - on ne prends en compte que les ecritures de
Suite ........ : type N pour calculer le solde
Suite ........ : - LG - 06/09/2004 - FB 10812 - rajout ds l'etablissement ds 
Suite ........ : la clause select
Suite ........ : - LG - 20/10/2005 -  FB 12630 - ouvreture direct du lettrage
Suite ........ : - LG - 03/11/2006 - prise en compte des ecritrues des type 
Suite ........ : L
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.ChargeJal(CodeJal : string ; vBoAfterZoom : boolean = false ) ;
var DateMvt, DebPeriode, FinPeriode : TDateTime ; St : string ;
Q : TQuery ;
begin
{$IFDEF TT}
AddEvenement('ChargeJal') ;
{$ENDIF}
if (CodeJal = '' ) or ( not vBoAfterZoom and ( CodeJal = FBoLastJal ) )then Exit ;
FBoLastJal := CodeJal ;
FRechCompte.Info.LoadJournal(CodeJal,vBoAfterZoom) ;
//Q:=OpenSQL('SELECT J_JOURNAL, J_NATUREJAL, J_COMPTEINTERDIT, J_MULTIDEVISE, J_CONTREPARTIE,'
//          +' J_TOTDEBP, J_TOTCREP, J_TOTDEBE, J_TOTCREE, J_TOTDEBS, J_TOTCRES, J_MODESAISIE,J_FERME,'
//          +' J_VALIDEEN, J_VALIDEEN1, J_DATEDERNMVT, J_NUMDERNMVT, J_COMPTEAUTOMAT,J_INCREF,J_INCNUM,J_NATCOMPL,J_EQAUTO,J_NATDEFAUT,J_ACCELERATEUR,J_OUVRIRLETT '
//          +' FROM JOURNAL WHERE J_JOURNAL="'+CodeJal+'"', TRUE) ;
memJal.Code:=CodeJal ;
memJal.Nature:=FRechCompte.Info.Journal.GetString('J_NATUREJAL') ;  
memJal.CInterdit:=FRechCompte.Info.Journal.GetString('J_COMPTEINTERDIT') ;
memJal.CAutomat:=FRechCompte.Info.Journal.GetString('J_COMPTEAUTOMAT') ;
memJal.bDevise:=FRechCompte.Info.Journal.GetString('J_MULTIDEVISE')='X' ;
memJal.ValideN:=FRechCompte.Info.Journal.GetString('J_VALIDEEN') ;
memJal.ValideN1:=FRechCompte.Info.Journal.GetString('J_VALIDEEN1') ;
memJal.CContreP:=FRechCompte.Info.Journal.GetString('J_CONTREPARTIE') ;
memJal.ModeSaisie:=FRechCompte.Info.Journal.GetString('J_MODESAISIE') ;
memJal.Accelerateur:=FRechCompte.Info.Journal.GetString('J_ACCELERATEUR') ='X' ;
memJal.Lettrage:=FRechCompte.Info.Journal.GetString('J_OUVRIRLETT') ='X' ;
if (memJal.Nature='BQE') or (memJal.Nature='CAI') then
  begin
  memParams.bSoldeProg:=memJal.CContreP<>'' ;
  HLSOLDEPROG.Caption:=TraduireMemoire('Solde théorique') ;
  HLSOLDEPROGCB.Caption:='('+memJal.CContreP+')' ;
  memJal.CAutomat:=memJal.CContreP+';' ;
  end else memParams.bSoldeProg:=FALSE ;
memJal.NbAuto:=0 ; memJal.CurAuto:=1 ;
St:=memJal.CAutomat ; while ReadTokenSt(St)<>'' do Inc(memJal.NbAuto) ;
memParams.bLibre:=(memJal.ModeSaisie='LIB') ;
DebPeriode:=EncodeDate(memOptions.Year, memOptions.Month, 1) ;
FinPeriode:=EncodeDate(memOptions.Year, memOptions.Month, memOptions.MaxJour) ;
DateMvt:=FRechCompte.Info.Journal.GetValue('J_DATEDERNMVT') ;
if (DateMvt>=DebPeriode) and (DateMvt<=FinPeriode)
   then memJal.LastNum:=FRechCompte.Info.Journal.GetValue('J_NUMDERNMVT')
   else memJal.LastNum:=1 ;
memJal.IncRef:=(FRechCompte.Info.Journal.GetString('J_INCREF')='X') ;
memJal.IncNum:=(FRechCompte.Info.Journal.GetString('J_INCNUM')='X') ;
memJal.EqAuto:=(FRechCompte.Info.Journal.GetString('J_EQAUTO')='X') ;
memJal.NatCompl:=(FRechCompte.Info.Journal.GetString('J_NATCOMPL')='X') ;
memJal.NatDefaut:=FRechCompte.Info.Journal.GetString('J_NATDEFAUT') ;
memJal.Ferme:=FRechCompte.Info.Journal.GetString('J_FERME')='X' ;
if vBoAfterZoom then exit ;
memJal.TotFolDebit:=0 ; memJal.TotFolCredit:=0 ;
memJal.TotSaiDebit:=0 ; memJal.TotSaiCredit:=0 ;
memJal.TotVirDebit:=0 ; memJal.TotVirCredit:=0 ;
memJal.TotDebDebit:=0 ; memJal.TotDebCredit:=0 ;
// Totaux sur la période
Q:=OpenSQL('SELECT SUM(E_DEBIT), SUM(E_CREDIT) FROM ECRITURE'
          +' WHERE E_JOURNAL="'+CodeJal+'"'
          +' AND E_EXERCICE="'+memOptions.Exo+'"'
          +' AND E_DATECOMPTABLE>="'+USDateTime(EncodeDate(memOptions.Year, memOptions.Month, 1))+'"'
          +' AND E_DATECOMPTABLE<="'+USDateTime(EncodeDate(memOptions.Year, memOptions.Month, memOptions.MaxJour))+'"'
          +' AND ( E_QUALIFPIECE="N" OR E_QUALIFPIECE="C" OR E_QUALIFPIECE="L" ) '
          +' AND E_ETABLISSEMENT="'+E_ETABLISSEMENT.Value+'" ' ,
          TRUE) ;
if (Q.BOF) and (Q.EOF) then
   begin
   memJal.TotPerDebit:=0 ; memJal.TotPerCredit:=0 ;
   end else
   begin
   memJal.TotPerDebit:=Q.Fields[0].AsFloat ;
   memJal.TotPerCredit:=Q.Fields[1].AsFloat ;
   end ;
Ferme(Q) ;
//if (memParams.bLibre) and (ObjFolio<>nil) then ObjFolio.ReadEquilibre(memJal) ;
AfficheTitreAvecCommeInfo;
//if memParams.bLibre then Caption:=GetMessageRC(RC_SAISIEECR)+' '+GetMessageRC(RC_MODELIB)
//                    else Caption:=GetMessageRC(RC_SAISIEECR)+' '+GetMessageRC(RC_MODEBOR) ;
VoirSoldeTreso ;
bGuideJal:=ExisteSQL('SELECT GU_GUIDE FROM GUIDE WHERE GU_TYPE="NOR" AND GU_TRESORERIE<>"X" AND GU_JOURNAL="'+memJal.Code+'"') ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 16/09/2002
Modifié le ... : 27/04/2006
Description .. : ? - ? - Je cherche à positionner l'entête par rapport au
Suite ........ : dernier mouvement sinon sur le premier journal autorisé en
Suite ........ : saisie
Suite ........ : - LG - 16/09/2002 - on se replace sur le dernier journal
Suite ........ : utilise
Suite ........ : - 21/10/2002 - s'il n'existe pas de dernier numero, on pard
Suite ........ : de 1
Suite ........ : - 23/09/2004 - LG - FB 13070 - 13070  - controle que le 
Suite ........ : dernier journal utilise est autorisse pour cette utilisateur
Suite ........ : - 05/10/2004 - LG - FB 14555 - verif que je dernier journal 
Suite ........ : utulise est present ds la liste des journaux dispo. ( on 
Suite ........ : pouvais transformer celui en journal de type piece et revenir 
Suite ........ : en saisie bor )
Suite ........ : - 25/10/2004 - LG - la correction precedente non fct pas. le 
Suite ........ : dernier journal etait forcment supprimé
Suite ........ : - 03/11/2004 - LG - on controle que le denier journal saisie 
Suite ........ : ne soit pas saisie
Suite ........ : - 19/07/2005 - LG - FB 14992 - on ne tenait pas compte de 
Suite ........ : l'option d'incrementation des bordereau
Suite ........ : - 14/04/2006 - FB 17873 - LG - si le dernier num de folio est 
Suite ........ : a 1, on l'incremenete as
Suite ........ :  Positionner le numéro de folio  à + 1 et être placé dedans 
Suite ........ : en position de saisie (dans la grille). Attention le paramètre 
Suite ........ : journal (incrémenter le bordereau) ne doit pas entre en ligne 
Suite ........ : de compte. C'est +1 par rapport au dernier numéro utilisé 
Suite ........ : dans le couple journal/période. 
Mots clefs ... : 
*****************************************************************}
Function TFSaisieFolio.GetLastJal : boolean ;
var Q : TQuery ; LastDate : TDateTime ; LastNum, i : Integer ; LastCode : string ;
begin
Result:=FALSE ;
{$IFDEF TT}
AddEvenement('GetLastJal') ;
{$ENDIF}
if InitParams.ParCreatCentral then
 begin
  LastDate := StrToDate(InitParams.ParPeriode) ;
  LastCode := InitParams.ParCodeJal ;
 end
  else
   begin
    LastDate:=GetParamSocSecur('SO_ZLASTDATE',iDate1900) ;
    LastCode:=GetParamSocSecur('SO_ZLASTJAL','') ;
   end ;
if ( FRechCompte.Info.Journal.Load([LastCode])<>-1 ) and
   ( CIsValidJal(LastCode,memPivot.Code, FRechCompte.Info ) <> RC_PASERREUR ) then LastCode := '' ;
if E_JOURNAL.Values.indexOf(LastCode) = -1 then LastCode := '' ;
if FRechCompte.Info.Journal.Load([LastCode])<>-1 then
 begin
   if InitParams.ParCreatCentral then
    begin
     LastNum := StrToInt(InitParams.ParNumFolio) ;
     if ( FAction = taCreat ) or ( LastNum = 0 ) then Inc(LastNum)
    end
     else
      begin
       LastNum:=FRechCompte.Info.Journal.GetValue('J_NUMDERNMVT') ; //
       if LastNum=0 then
        Inc(LastNum)  // s'il n'existe pas de denier numero, on part de 1
         else
         if (FRechCompte.Info.Journal.GetValue('J_INCNUM') = 'X') then Inc(LastNum) ;
      end ;
   AlimenteEntete(DateToStr(LastDate), LastCode, IntToStr(LastNum)) ;
   CurPeriode:='' ; CurJal:='' ; CurNumFolio:='' ;
   exit ;
 end;

LastDate:=EncodeDate(1900, 1, 1) ; LastNum:=-1 ; LastCode:='' ;
Q:=OpenSQL('SELECT J_JOURNAL, J_DATEDERNMVT, J_NUMDERNMVT, J_MODESAISIE FROM JOURNAL WHERE J_MODESAISIE<>"-"', TRUE) ;
while not Q.EOF do
  begin
  if (Q.FindField('J_DATEDERNMVT').AsDateTime>LastDate) and
     (IsJalAutorise(Q.FindField('J_JOURNAL').AsString)) then
     begin
     LastDate:=Q.FindField('J_DATEDERNMVT').AsDateTime ;
     LastNum:=Q.FindField('J_NUMDERNMVT').AsInteger ;
     LastCode:=Q.FindField('J_JOURNAL').AsString ;
     end ;
  Q.Next ;
  end ;
Ferme(Q) ;
if ( FRechCompte.Info.Journal.Load([LastCode])<>-1 ) and
( CIsValidJal(LastCode,memPivot.Code, FRechCompte.Info ) <> RC_PASERREUR ) then LastCode := '' ;
if (LastNum<>-1) and (LastCode<>'') then
   begin
   Result:=True ;
   if memJal.IncNum then Inc(LastNum) ;
   AlimenteEntete(DateToStr(LastDate), LastCode, IntToStr(LastNum)) ;
   CurPeriode:='' ; CurJal:='' ; CurNumFolio:='' ;
   end else
   begin
   for i:=0 to E_JOURNAL.Items.Count-1 do
     if IsJalAutorise(E_JOURNAL.Values[i]) then
       begin
        E_JOURNAL.Value:=E_JOURNAL.Values[i] ;
        Break ;
       end ;
   end ;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 08/10/2002
Modifié le ... :   /  /
Description .. : - 08/10/2002 - les nouveaux folios n'avaient pas de
Suite ........ : numeros
Mots clefs ... :
*****************************************************************}
procedure TFSaisieFolio.NumeroteFolio ;
begin
if (FAction=taConsult) or ( not MemJal.IncNum ) then
 begin
  CbFolio.ItemIndex:=0 ;
  if not (CbFolio.Value='') then Exit ; // il existe une valeur on peut sortir
 end ;
while CbFolio.Values.IndexOf(IntToStr(NumFolio))>=0 do Inc(NumFolio) ;
CbFolio.Text:=IntToStr(NumFolio) ;
end ;

function TFSaisieFolio.VoirSoldeTreso : Boolean ;
begin
Result:=FALSE ;
if memParams.bSoldeProg then
   begin
   BEVELSOLDEPROG.Visible:=TRUE ; LSA_SOLDEPROG.Visible:=TRUE ;
   HLSOLDEPROG.Visible:=TRUE ;    HLSOLDEPROGCB.Visible:=TRUE ;
   Result:=TRUE ;
   end else
   begin
   BEVELSOLDEPROG.Visible:=FALSE ; LSA_SOLDEPROG.Visible:=FALSE ;
   HLSOLDEPROG.Visible:=FALSE ;    HLSOLDEPROGCB.Visible:=FALSE ;
   end ;
end ;

procedure TFSaisieFolio.PosLesSoldes ;
begin
LSA_SOLDEG.SetBounds(SA_SOLDEG.Left,SA_SOLDEG.Top,SA_SOLDEG.Width,SA_SOLDEG.Height) ;
LSA_SOLDET.SetBounds(SA_SOLDET.Left,SA_SOLDET.Top,SA_SOLDET.Width,SA_SOLDET.Height) ;
LSA_SOLDE.SetBounds(SA_SOLDE.Left,SA_SOLDE.Top,SA_SOLDE.Width,SA_SOLDE.Height) ;
LSA_SOLDEPROG.SetBounds(SA_SOLDEPROG.Left,SA_SOLDEPROG.Top,SA_SOLDEPROG.Width,SA_SOLDEPROG.Height) ;
LSA_TOTALDEBIT.SetBounds(SA_TOTALDEBIT.Left,SA_TOTALDEBIT.Top,SA_TOTALDEBIT.Width,SA_TOTALDEBIT.Height) ;
LSA_TOTALCREDIT.SetBounds(SA_TOTALCREDIT.Left,SA_TOTALCREDIT.Top,SA_TOTALCREDIT.Width,SA_TOTALCREDIT.Height) ;
LSA_FOLIODEBIT.SetBounds(SA_FOLIODEBIT.Left,SA_FOLIODEBIT.Top,SA_FOLIODEBIT.Width,SA_FOLIODEBIT.Height) ;
LSA_FOLIOCREDIT.SetBounds(SA_FOLIOCREDIT.Left,SA_FOLIOCREDIT.Top,SA_FOLIOCREDIT.Width,SA_FOLIOCREDIT.Height) ;
SA_SOLDEG.Visible:=FALSE ;     SA_SOLDET.Visible:=FALSE ;
SA_SOLDEPROG.Visible:=FALSE ;  SA_SOLDE.Visible:=FALSE ;
SA_TOTALDEBIT.Visible:=FALSE ; SA_TOTALCREDIT.Visible:=FALSE ;
SA_FOLIODEBIT.Visible:=FALSE ; SA_FOLIOCREDIT.Visible:=FALSE ;
LSA_SOLDEG.Visible:=FALSE ;    LSA_SOLDET.Visible:=FALSE ;
SA_SoldePerGen.Visible:=FALSE ; SA_SoldeAux.Visible:=FALSE ;
SoldeAuxPer.Visible:=false ; SA_SOLDEG.Visible:=FALSE ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 06/09/2002
Modifié le ... :   /  /
Description .. : - 06/10/2002 - la longeur saisisable dans la cellule des
Suite ........ : generaux passe à 30 car
Mots clefs ... :
*****************************************************************}
procedure TFSaisieFolio.InitGrid ;
begin
bFirst:=TRUE ;
// Avant le VidePile pour bien placer le curseur
GS.Row:=GS.FixedRows ; GS.Col:=GS.FixedCols ;
GS.VidePile(FALSE)  ;
GS.RowCount:=100 ;
GS.ElipsisButton:=FALSE ;
GS.Enabled:=FALSE ;
GS.ColWidths[SF_ETABL]:=-1 ;
GS.ColWidths[SF_QUALIF]:=-1 ;
GS.ColWidths[SF_EURO]:=-1 ;
//LG*
{$IFDEF TT}
//GS.ColWidths[SF_PIECE]:=20 ;
//GS.ColWidths[SF_NEW]:=20 ;
{$ELSE}
GS.ColWidths[SF_PIECE]:=-1 ;
GS.ColWidths[SF_NEW]:=-1 ;
{$ENDIF}
//GS.ColWidths[SF_NUML]:=0 ;
GS.ColWidths[SF_NUMO]:=0 ;
GS.ColAligns[SF_NUML]:=taCenter ;
GS.ColAligns[SF_NUMO]:=taCenter ;
GS.ColAligns[SF_JOUR]:=taCenter ;
GS.ColAligns[SF_DEBIT]:=taRightJustify ;
GS.ColAligns[SF_CREDIT]:=taRightJustify ;
GS.ColLengths[SF_JOUR]:=2 ;
GS.ColLengths[SF_GEN]:=30 ; //VH^.Cpta[fbGene].Lg ;
GS.ColLengths[SF_AUX]:=VH^.Cpta[fbAux].Lg ;
if VH^.CPCodeAuxiOnly then GS.ColLengths[SF_AUX]:=VH^.Cpta[fbAux].Lg else GS.ColLengths[SF_AUX]:=35 ;
GS.ColLengths[SF_REFI]:=35 ;
GS.ColLengths[SF_LIB]:=35 ;
GS.PostDrawCell:=PostDrawCell ;
FCurrentRowEnter:=-1 ;
GS.Refresh ;
end ;

procedure TFSaisieFolio.InitPivot ;
begin
if memPivot.Code<>'' then Exit ;
memPivot.Code:=V_PGI.DevisePivot ;
if memPivot.Code<>'' then GetInfosDevise(memPivot) ;
//memOppos.Code:=V_PGI.DeviseFongible ;
//if memOppos.Code<>'' then GetInfosDevise(memOppos) ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 05/01/2005
Modifié le ... :   /  /
Description .. : FB 15185 - 05/01/2005 - on fait tout le temps le
Suite ........ : positionneEtabUser
Mots clefs ... :
*****************************************************************}
procedure TFSaisieFolio.InitEtabl ;
begin
if InitParams.ParEta <> '' then
 E_ETABLISSEMENT.Value:= InitParams.ParEta
 else
  E_ETABLISSEMENT.Value:=VH^.EtablisDefaut ;
//if FAction=taCreat then
PositionneEtabUser(E_ETABLISSEMENT) ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 12/06/2002
Modifié le ... :   /  /
Description .. : -12/06/2002 - blocage du rafraichissement de la grille
Mots clefs ... :
*****************************************************************}
procedure TFSaisieFolio.SetNum(Row : LongInt) ;
var i : Integer ;
begin
HGBeginUpdate(GS) ;
try
 for i:=GS.FixedRows to nbLignes do GS.Cells[SF_NUML,i]:=IntToStr(i) ;
finally
 HGEndUpdate(GS);
end;
end ;

{procedure TFSaisieFolio.SetOrdre(Row : LongInt) ;
var OAvant, OApres, OApres2, Diff, iEspace, i : Integer ;
begin
// Détermination de l'ordre des lignes n-1, n+1 et n+2 : Insertion après
if nbLignes=1 then Row:=Row-1 ;
if nbLignes=1 then  OAvant:=-1 else  OAvant:=StrToInt(GS.Cells[SF_NUMO, Row-1]) ;
if Row+1>nbLignes then  OApres:=-1 else  OApres:=StrToInt(GS.Cells[SF_NUMO, Row+1]) ;
if Row+2>nbLignes then OApres2:=-1 else OApres2:=StrToInt(GS.Cells[SF_NUMO, Row+2]) ;
// La ligne est ajoutée à la fin de la liste : Pas de MAJ des lignes adjacentes
if OApres<0 then
  begin
  if OAvant<0 then GS.Cells[SF_NUMO, Row]:=IntToStr(MAGIC_NUMBER)
              else GS.Cells[SF_NUMO, Row]:=IntToStr(OAvant+MAGIC_NUMBER) ;
  end else
  begin
  if OAvant<0 then Diff:=OApres-1 else Diff:=OApres-OAvant-1 ;
  // Une seule possibilité !
  if Diff=1 then GS.Cells[SF_NUMO, Row]:=IntToStr(OApres-1) ;
  // Multiples possibilités !
  if Diff>1 then GS.Cells[SF_NUMO, Row]:=IntToStr(OApres-((Diff+1) div 2)) ;
  // Aucune possiblité, il faut décaler !
  if Diff=0 then
    begin
    if OApres2<0 then
      begin
      GS.Cells[SF_NUMO, Row]:=IntToStr(OAvant+MAGIC_NUMBER) ;
      GS.Cells[SF_NUMO, Row+1]:=IntToStr(OAvant+MAGIC_NUMBER+MAGIC_NUMBER) ;
      end else
      begin
      // Recherche d'un espace suffisant avant la fin de la liste
      iEspace:=-1 ;
      for i:=Row+2 to nbLignes do
        begin
        Diff:=StrToInt(GS.Cells[SF_NUMO, i])-StrToInt(GS.Cells[SF_NUMO, i-1])-1 ;
        if Diff>=1 then begin iEspace:=i ; Break ; end ;
        end ;
      // Espace insuffisant : modification de l'ordre jusqu'à la fin de la liste
      if iEspace<0 then
        begin
        for i:=Row to nbLignes do GS.Cells[SF_NUMO, i]:=IntToStr(OAvant+(MAGIC_NUMBER*(i-Row+1))) ;
        end else
        // Sinon réaliser l'optimisation
        begin
        GS.Cells[SF_NUMO, Row]:=IntToStr(OApres) ;
        for i:=Row+1 to iEspace-1 do GS.Cells[SF_NUMO, i]:=IntToStr(StrToInt(GS.Cells[SF_NUMO, i])+1) ;
        end ;
      end ;
    end ;
  end ;
end ; }


function TFSaisieFolio.GetRowDate(Row : LongInt) : TDateTime ;
begin
Result:=EncodeDate(memOptions.Year, memOptions.Month, StrToInt(GS.Cells[SF_JOUR, Row])) ;
end ;

function TFSaisieFolio.GetRowFirst(Piece : LongInt) : LongInt ;
var i : Integer ;
begin
Result:=-1 ; HGBeginUpdate(GS) ;
try
for i:=GS.FixedRows to nbLignes do
  begin
  if GS.Cells[SF_PIECE, i]<>IntToStr(Piece) then Continue ;
  Result:=i ; Break ;
  end ;
finally
 HGEndUpdate(GS) ;
end;
end ;

{function TFSaisieFolio.GetRowLast(Piece : LongInt) : LongInt ;
var i : Integer ;
begin
Result:=-1 ; HGBeginUpdate(GS) ;
try
for i:=GS.FixedRows to nbLignes do
  begin
  if GS.Cells[SF_PIECE, i]<>IntToStr(Piece) then Continue ;
  Result:=i ;
  end ;
finally
 HGEndUpdate(GS) ;
end;
end ;}

procedure TFSaisieFolio.PutRowMontant(Col, Row : LongInt; Montant : Double) ;
begin
if Montant<>0 then GS.Cells[Col, Row]:=StrFMontant(Montant,15,memPivot.Decimale,'',TRUE)
              else GS.Cells[Col, Row]:='' ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 05/04/2002
Modifié le ... : 07/06/2006
Description .. : - correction du mode libre - on ne peut plus valider qur la
Suite ........ : premiere ligne ( en pouvait enregistrer une ligne vide )
Suite ........ : -07/11/2002 - on ne test plus la validite de la ligne pour le
Suite ........ : bouton BInsert, fait ds la fct meme
Suite ........ : CA - 05/02/2003 - Pour éviter les problèmes en insertion sur
Suite ........ : lignes validée ou écart ( sol. temporaire, à corriger )
Suite ........ : - LG - 25/06/2003 - activation de la consultation et du 
Suite ........ : lettrage si on lettre en saisie
Suite ........ : - LG - 29/09/2003 - FB 12781 - on active pas le bp scan 
Suite ........ : sur l'entete du bor
Suite ........ : - LG - 02/10/2003 - FB 11907 - liberation du bouton insert
Suite ........ : sur une ligne vlidées
Suite ........ : - LG 23/10/2003 - oin avait aps acces au menu zoom qd le
Suite ........ : parma saisie en temlsp reeel etait coché
Suite ........ : - LG 01/07/2004 - FB 13589 - la fiche tiers est accesible 
Suite ........ : meme si le bordereau n'est aps equilibre
Suite ........ : - LG - 06/08/2004 - FB 13885 - permettre la validation d'un 
Suite ........ : folio libre a blanc
Suite ........ : - LG - 02/09/2004 - FB 14487 - on ne pouvais plus lettrer 
Suite ........ : les tic ou tid
Suite ........ : - LG - 11/10/2005 - FB 16720 - correction pour les comptes 
Suite ........ : divers lettrable
Suite ........ : - LG - 12/05/2006 - FB 17870 - J’ai systématiquement « 
Suite ........ : Voir le document numérisé » même si pas de document.  
Suite ........ : - LG - 07/06/2006 - FB 18263 - on desactive l'appel par le
Suite ........ : journal centralisateur qd on a fait un lettrage en saisie
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.EnableButtons(Ou : integer=-1) ;
var Ecr : TZF ; bEnabled, bExistG, bRowCanEdit : Boolean ;
    NumCompte : string ; ARow ,k,kt : Integer ;
    lBoLettrable : boolean ; NumAux : string ;
begin
if csDestroying in ComponentState then Exit ;
if bClosing then exit ;
{$IFDEF TT}
AddEvenement('EnableButtons') ;
try
{$ENDIF}
if Ou<>-1 then ARow:=Ou else ARow:=GS.Row ; lBoLettrable := false ;
if ObjFolio<>nil then Ecr:=ObjFolio.GetRow(ARow-GS.FixedRows) else Ecr:=nil ;
{$IFDEF SCANGED}
BScan.enabled:= ( InitParams.ParGuide = '' ) and ( True ) and BScan.visible and (not bGuideRun)
                and ( ( Ecr<> nil ) and ( ObjFolio.RechGuidId(Ecr)<> '' ) ) ;
BPDF.Enabled := VH^.OkModGed and ( InitParams.ParGuide = '' ) and (not bGuideRun) and not (ctxPCL in V_PGI.PGIContexte) ;
{$ELSE}
BScan.Visible := False;
BPDF.Enabled := false ;
{$ENDIF}
if bModeRO then
 bEnabled:=TRUE
  else
   if memParams.bLibre then
    bEnabled:=IsSoldeFolio
     else
      bEnabled:= (IsSoldeJournal and IsSoldeFolio ) ; //and (NbLignes>1)) ;   enlever a l'arrche a vérifier c truc

if (not bModeRO) then
 BValide.Enabled:=not FBoValEnCours and bEnabled and not bGuideRun and ( InitParams.ParGuide = '' ) ;

// Boutons
{$IFDEF CCS3}
BSelectFolio.Enabled:=False ; BSelectFolio.Visible:=False ;
{$ELSE}
BSelectFolio.Enabled:= ( not FBoLettrageEnCours ) and ( InitParams.ParGuide = '' ) and not FBoValEnCours and (bEnabled) and (not bGuideRun) and (FAction<>taConsult) ;
{$ENDIF}
BZoomJalCentral.Enabled:=BSelectFolio.Enabled ;
bRowCanEdit:=(IsRowCanEdit(ARow)<=0) ;
BSolde.Enabled:=( InitParams.ParGuide = '' ) and not FBoValEnCours and (not bGuideRun) and (not bModeRO) and (not bEnabled) and (bRowCanEdit) ;
BEche.Enabled:=( InitParams.ParGuide = '' ) and not FBoValEnCours and (VH^.ZSAISIEECHE(*memParams.bEcheance*)) and IsEch(ARow) ;
BVentil.Enabled:=FALSE ;

if Length(GS.Cells[SF_GEN, ARow]) = VH^.Cpta[fbGene].Lg then
 begin
  NumCompte:=GS.Cells[SF_GEN, ARow] ;k:=Comptes.GetCompte(NumCompte) ;
  if (k>-1) and (Comptes.GetValue('G_COLLECTIF', k)='X') then
  begin
   NumAux:=GS.Cells[SF_AUX, ARow] ; kt:=Tiers.GetCompte(NumAux) ;
   lBoLettrable:= (kt>-1) and (Tiers.GetValue('T_LETTRABLE', kt)='X') ;
  end
   else   // if (Comptes.GetValue('G_LETTRABLE', k)='-')
    //lBoLettrable:= (k>-1) and (Comptes.IsTiers(k)) ;
    lBoLettrable:= (k>-1) and ( Comptes.GetValue('G_LETTRABLE', k)='X' ) ;
 end ;

BLettrage.Enabled:= ( InitParams.ParGuide = '' ) and not FBoValEnCours and (not bGuideRun) and lBoLettrable and not bModeRO ;

if (NumCompte<>'') and ((GS.Cells[SF_DEBIT, ARow]<>'') or (GS.Cells[SF_CREDIT, ARow]<>'')) then
   begin
   k:=Comptes.GetCompte(NumCompte) ;
   if (k>=0) and (Comptes.IsVentilable(0, k)) then BVentil.Enabled:=TRUE ;
   end ;
bExistG:=(not bGuideRun)
          and ((nbLignes=1) or (bEnabled))
          //and (not memParams.bLibre)
          and (not bModeRO)
          and ((nbLignes=0) or ((nbLignes>0) and (ARow=nbLignes)))
          and (GS.Cells[SF_DEBIT,  ARow]='')
          and (GS.Cells[SF_CREDIT, ARow]='')
          and ( InitParams.ParGuide = '' ) ;
if bExistG then
 BTGuide.Enabled:=not FBoValEnCours and ExisteGuide
  else BTGuide.Enabled:=not FBoValEnCours and bExistG ;
BGuide.Enabled:=not FBoValEnCours and BTGuide.Enabled ;
// Bouton rechercher
BChercher.Enabled:=not FBoValEnCours and (not bGuideRun) and (nbLignes<>0) ;
// Boutons Ins/Del ligne
// CA - 05/02/2003 - Pour éviter les problèmes en insertion sur lignes validée ou écart ( sol. temporaire, à corriger )
// LG - 02/10/2003 - FB 11907 - debranche, on verra bien
BInsert.Enabled:=( InitParams.ParGuide = '' ) and not FBoValEnCours and (not bGuideRun) and (not bModeRO) and (nbLignes<>0) ; //and bRowCanEdit ;

BSDel.Enabled:=( InitParams.ParGuide = '' ) and not FBoValEnCours and (not bGuideRun) and (not bModeRO) and (nbLignes<>0) and bRowCanEdit ;
//Interdire de détruire la dernière ligne d'un folio existant
 // if (not memOptions.NewObj) and (nbLignes=1) then BSDel.Enabled:=FALSE ; {JLDBOR}
BMenuZoom.Enabled:=( InitParams.ParGuide = '' ) and not FBoValEnCours and (nbLignes<>0) ;
BComplement.Enabled:=( InitParams.ParGuide = '' ) and not FBoValEnCours and (nbLignes<>0) and (GS.Cells[SF_GEN, ARow]<>'') ;
BLibAuto.Enabled:=( InitParams.ParGuide = '' ) and not FBoValEnCours and (bRowCanEdit) and (not bModeRO) and (nbLignes<>0) ;
// Zoom
BMenuZoom.Enabled:=( InitParams.ParGuide = '' ) and not FBoValEnCours and (not bGuideRun) ;
if GS.Cells[SF_AUX, ARow]='' then BZoomTiers.Enabled:=FALSE
                               else BZoomTiers.Enabled:=not FBoValEnCours ;

if (GS.Cells[SF_GEN, ARow]='') and (GS.Cells[SF_AUX, ARow]='')
   then begin BZoom_FB19024.Enabled:=FALSE ; BZoomEcrs.Enabled:=FALSE ; BZoomCpte.Enabled:=FALSE ; end
   else
    begin
     BZoom_FB19024.Enabled:=not FBoValEnCours and ( InitParams.ParGuide = '' ) ;
     //BZoomEcrs.Enabled:=not FBoValEnCours and ( InitParams.ParGuide = '' ) ;
     BZoomCpte.Enabled:=not FBoValEnCours and ( InitParams.ParGuide = '' );
    end ;

if memParams.bRealTime then
  begin
     //BZoomEcrs.Enabled:=not FBoValEnCours and (not bGuideRun) and ( InitParams.ParGuide = '' ) ; //and (bEnabled) ;
     BZoomCpte.Enabled:=not FBoValEnCours and (not bGuideRun) and ( InitParams.ParGuide = '' ); //and (bEnabled) ;
     BZoomJournal.Enabled:=not FBoValEnCours and (not bGuideRun) and (bEnabled) and ( InitParams.ParGuide = '' ) ;
     //BZoomTiers.Enabled:=IsRowOk(ARow) and (GS.Cells[SF_AUX, ARow]<>'') and (not bGuideRun) and (bEnabled) ;
  end
   else
    begin
     BZoomCpte.Enabled:=FALSE ; // Consultation des comptes
    end ;

BModifTva.Enabled:=( InitParams.ParGuide = '' ) and not FBoValEnCours and ((bRowCanEdit) and (nbLignes<>0) and ((IsRowHT(ARow)) or (IsRowTiers(ARow)))) ;
// Saisie Devise
BDevise.Enabled:= not FBoValEnCours
                  and ((GS.Cells[SF_DEVISE, ARow]<>'') or  (((nbLignes=1) or ((bEnabled) and (nbLignes>1) and (ARow=nbLignes)))
                  and (GS.Cells[SF_DEBIT,  ARow]='')
                  and (GS.Cells[SF_CREDIT, ARow]='') ) ) ;
BMonnaie.Enabled:=not FBoValEnCours and BDevise.Enabled ;
// RIB
if (bRowCanEdit) and (Ecr<>nil) and (Ecr.GetValue('E_AUXILIAIRE')<>'') and ((not VH^.AttribRIBAuto) or (Ecr.GetValue('E_RIB')<>''))
  then BModifRIB.Enabled:=not FBoValEnCours else BModifRIB.Enabled:=FALSE ;
// Immos
if (Ecr<>nil) then
  BEGIN
{$IFDEF AMORTISSEMENT}
  BZoomImmo.Enabled:=( InitParams.ParGuide = '' ) and not FBoValEnCours and ((Ecr.GetValue('E_IMMO')<>'') and ((VH^.OkModImmo) or (V_PGI.VersionDemo))) ;
{$ELSE}
  BZoomImmo.Enabled:=FALSE ;
{$ENDIF}
  END ;
BAcc.Enabled:= ( InitParams.ParGuide = '' ) and not bModeRO and not FBoValEnCours and BLettrage.Enabled and memJal.Accelerateur and ( GS.Row = nbLignes)
              and ( (trim(GS.Cells[SF_DEBIT,  ARow])<>'') or (trim(GS.Cells[SF_CREDIT, ARow])<>'') ) and IsSoldefolio ;
{$IFDEF TT}
except
on E : Exception do
 begin
  ShowMessage('EnableButtons ' +  E.message );
 end ;
end ;
{$ENDIF}
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 26/06/2002
Modifié le ... : 09/10/2003
Description .. : - 25/06/2002 - gestion de l'affichage complet ou abrege de
Suite ........ : la nature de piece
Suite ........ : - 09/10/2003 - FB 11850 - on remet les boutons a jour suite
Suite ........ : au changement de nature, activation des guides
Mots clefs ... :
*****************************************************************}
procedure TFSaisieFolio.NextNature(Row : LongInt) ;
var i : Integer ;
begin
if FBoNatureComplete then
 begin
  i:=E_NATUREPIECE.Items.IndexOf(GS.Cells[SF_NAT, Row]) ; i:=i+1 ;
  if i>E_NATUREPIECE.Items.Count-1 then i:=0 ;
  GS.Cells[SF_NAT, Row]:=E_NATUREPIECE.Items[i] ;
 end
  else
   begin
    i:=E_NATUREPIECE.Values.IndexOf(GS.Cells[SF_NAT, Row]) ; i:=i+1 ;
    if i>E_NATUREPIECE.Values.Count-1 then i:=0 ;
    GS.Cells[SF_NAT, Row]:=E_NATUREPIECE.Values[i] ;
   end; // if
Statut:=taModif ;
GrilleModif:=true ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 25/06/2002
Modifié le ... : 02/09/2004
Description .. : - 25/06/2002 - gestion de l'affichage de la nature en clair 
Suite ........ : ou abrege
Suite ........ : - chg de la nature par defaut pour un journal de bq : OD
Suite ........ : 
Suite ........ : - 29/07/2002 - gestion de la nature de piece par defaut sur 
Suite ........ : un journal
Suite ........ : - 02/09/2004 - LG - FB 14247 - l'appel des scenario ne fct 
Suite ........ : pas pour un seul etablissement
Mots clefs ... :
*****************************************************************}
function TFSaisieFolio.InitNature : string ;
begin
if trim(MemJal.NatDefaut)<>'' then
 begin
  if FBoNatureComplete then result:=E_NATUREPIECE.Items[E_NATUREPIECE.Values.IndexOf(MemJal.NatDefaut)]
  else result:=MemJal.NatDefaut ;
 end
  else
   begin
    if FBoNatureComplete then
       case CaseNatJal(memJal.Nature) of
        tzJVente       : Result:=E_NATUREPIECE.Items[E_NATUREPIECE.Values.IndexOf('FC')] ;
        tzJAchat       : Result:=E_NATUREPIECE.Items[E_NATUREPIECE.Values.IndexOf('FF')] ;
        tzJBanque      : Result:=E_NATUREPIECE.Items[E_NATUREPIECE.Values.IndexOf('OD')] ;
        tzJEcartChange : Result:=E_NATUREPIECE.Items[E_NATUREPIECE.Values.IndexOf('OD')] ;
        tzJOD          : Result:=E_NATUREPIECE.Items[E_NATUREPIECE.Values.IndexOf('OD')] ;
        end
      else
       case CaseNatJal(memJal.Nature) of
         tzJVente       : Result:='FC' ;
         tzJAchat       : Result:='FF' ;
         tzJBanque      : Result:='OD' ;
         tzJEcartChange : Result:='OD' ;
         tzJOD          : Result:='OD' ;
        end ; // case
   end; // if
FZScenario.Load([MemJal.Code,Result,'N',E_ETABLISSEMENT.Value]);
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 22/07/2002
Modifié le ... : 20/09/2007
Description .. : Nouveau param -> on peut choisir de ne pas incremente le 
Suite ........ : numero
Suite ........ : 
Suite ........ : rem : appelle par le createrow
Suite ........ : - LG - 20/07/2005 - FB 10794 - stockage de la der ref 
Suite ........ : interne
Suite ........ : - LG - FB 21472 - 20/09/2007 - correction de l'incremente 
Suite ........ : de la reference
Mots clefs ... : 
*****************************************************************}
function TFSaisieFolio.IncRefAuto(RefI : string): string ;
var i, iStart, iEnd : Integer ; bFind, bNum : Boolean ; Val : LongInt ;
 lStZero : string ;
begin
Result:=RefI ; if not MemJal.IncRef then begin result:='' ; exit ; end ; bFind:=FALSE ; iStart:=-1 ; iEnd:=-1 ;
if trim(RefI) = '' then exit ; lStZero:='' ; i:=1 ;
while RefI[i] = '0' do
 begin
  lStZero := lStZero + RefI[i] ;
  inc(i) ;
 end;
Refi := copy(Refi,1,length(RefI)) ;
for i:=Length(RefI) downto 1 do
  begin
  bNum:=(RefI[i] in ['0'..'9']) ;
  if (bFind) and (not bNum) then begin iStart:=i ; Break ;       end ;
  if (not bFind) and (bNum) then begin iEnd:=i ;   bFind:=TRUE ; end ;
  end ;
if (iEnd>0) and (iStart<0) then iStart:=0 ;
if (iStart>=0) and (iEnd>0) then
  begin
  Result:='' ;
  if Length(Copy(RefI, iStart+1, iEnd-iStart))>9 then begin Result:=RefI ; Exit ; end ;
  Val:=ValeurI(Copy(RefI, iStart+1, iEnd-iStart))+1 ;
  Result:=Copy(RefI, 1, iStart) ;
  Result:=Result+IntToStr(Val) ;
  Result:=Result+Copy(RefI, iEnd+1, Length(RefI)-iEnd+1) ;
  if Length(Result)>35 then Result:=RefI ;
  end ;
 result := lStZero + result ;
 FLastRef := Result ;
end ;

Function TFSaisieFolio.ChoixCpteAuto ( LesCptAuto : String ) : String ;
Var sWhere,Cpts,Cpt : String ;
BEGIN
Result:='' ; sWhere:='' ;
if LesCptAuto='' then Exit ;
Cpts:=LesCptAuto ;
Repeat
 Cpt:=ReadTokenSt(Cpts) ;
 if Cpt<>'' then
    BEGIN
    if Length(Cpt)=VH^.Cpta[fbGene].Lg then sWhere:=sWhere+' OR G_GENERAL="'+Cpt+'" '
                                       else sWhere:=sWhere+' OR G_GENERAL LIKE "'+Cpt+'%"' ;
    END ;
Until ((Cpts='') or (Cpt='')) ;
if sWhere='' then Exit ;
Delete(sWhere,1,4) ;

Result:=Choisir(GetMessageRC(44),'GENERAUX','G_GENERAL||"  "||G_LIBELLE','G_GENERAL',sWhere,'G_GENERAL') ;

END ;

procedure TFSaisieFolio.GetCompteAuto(Indice : Integer) ;
var NumCompte : string ;
begin
if GS.Col<>SF_GEN then exit ;
if Indice<=0 then NumCompte:=ChoixCpteAuto(memJal.CAutomat)
             else NumCompte:=TrouveAuto(memJal.CAutomat, Indice) ;
if NumCompte='' then Exit ;
if EstInterdit(memJal.CInterdit, NumCompte, 0)>0 then Exit ;
GS.Cells[SF_GEN, GS.Row]:=NumCompte ; Statut:=taModif ; GrilleModif:=true ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 10/12/2002
Modifié le ... : 14/12/2005
Description .. : - 10/12/2002 - fiche  10938 - calcul du solde à la demande
Suite ........ : - 10/12/2002 - fiche 10938 - on réaffiche les libelles du 
Suite ........ : compte et du tiers
Suite ........ : -20/01/2003 - FB 10938 - le solde des aux apparaisait
Suite ........ : systematiquement
Suite ........ : - 08/10/2003 - on affiche systematiquement le solde
Suite ........ : - LG 24/08/2004 - FB 14291 - affichage de automaitque a 
Suite ........ : cote du libelle du compte d'accelerateur
Suite ........ : - LG - 14/09/2004 - le solde du compte est calcule en fct 
Suite ........ : de l'exercice de reference
Suite ........ : - LG - 04/10/2004 - FB 14750 - affichage du libelle du
Suite ........ : compte accelere
Suite ........ : - LG - 14/12/2005 - FB 17192 - on ne saisie pas les qt sur
Suite ........ : une ligne non modifiable
Mots clefs ... :
*****************************************************************}
procedure TFSaisieFolio.InfosPied(Row : LongInt ; bAvecTouteLesInfo : boolean = true) ;
var Ecr : TZF ; k : Integer ; DIG, CIG, DIA, CIA, MDevise : double ;
    NumCompte, NumAux : string ; RefRow : LongInt ;
    lStCompte : string ;
begin
if Row<0 then RefRow:=GS.Row else RefRow:=Row ;
G_LIBELLE.Font.Color:=clBlack ;
G_LIBELLE.Font.Style:=G_LIBELLE.Font.Style-[fsBold] ;
G_LIBELLE.Caption:='' ;  YTC_SCHEMAGEN.Caption := '' ;
NumCompte:=GS.Cells[SF_GEN, RefRow] ;
k:=Comptes.GetCompte(NumCompte) ;
if k>=0 then
   begin
   G_LIBELLE.Caption:=Comptes.GetValue('G_LIBELLE', k) ;
   Comptes.Solde(DIG,CIG,memOptions.TypeExo) ;
   LSA_SOLDEG.Visible:=true ;
   SoldeD.Visible:=true ;
//   if bAvecTouteLesInfo then //08/10/2003 suppression du param
   CalculSoldeCompte(Comptes.GetValue('G_GENERAL', k), SF_GEN, DIG, CIG) ;
   PQuantite.Visible :=  //(isRowCanEdit(GS.Row)=0) and (not EstAccelere(RefRow)) and  ( GetParamSocSecur('SO_CPPCLSAISIEQTE',false))  and
                          (not memJal.Lettrage) and (isRowCanEdit(GS.Row)=0) and ( GetParamSocSecur('SO_CPPCLSAISIEQTE',false))  and
                          not (bGuideRun) and   ( ( ( Comptes.GetValue('G_QUALIFQTE1', k) <> '' ) and ( Comptes.GetValue('G_QUALIFQTE1', k) <> 'AUC' ) ) or
                          ( ( Comptes.GetValue('G_QUALIFQTE2', k) <> '' ) and ( Comptes.GetValue('G_QUALIFQTE2', k) <> 'AUC' ) ) ) ;
   if PQuantite.Visible then
    begin
     Ecr:=ObjFolio.GetRow(RefRow-GS.FixedRows) ;
     if Ecr<>nil then
      GereQuantite(Ecr) ;
    end ; // if
   end
    else
     begin
      SoldeD.Visible:=FALSE ;
      LSA_SOLDEG.Visible:=FALSE ;
      PQuantite.Visible := false ;
     end ;
T_LIBELLE.Caption:='' ;
NumAux:=GS.Cells[SF_AUX, RefRow] ; k:=Tiers.GetCompte(NumAux) ;
if k>=0 then
   begin
   T_LIBELLE.Caption:=Tiers.GetValue('T_LIBELLE', k) ;
   DIA:=Tiers.GetValue('T_TOTALDEBIT', k) ;
   CIA:=Tiers.GetValue('T_TOTALCREDIT', k) ;
   LSA_SOLDET.Visible:=TRUE ;
   CalculSoldeCompte(Tiers.GetValue('T_AUXILIAIRE', k), SF_AUX, DIA, CIA) ;
   if (ObjFolio=nil) then exit ;
   Ecr:=ObjFolio.GetRow(RefRow-GS.FixedRows) ; if Ecr=nil then exit ;
   lStCompte := '' ;
   if memJal.Accelerateur and not (GS.Cells[SF_DEVISE, RefRow]<>'') then // on affiche pas les info sur l'accelelrateue de saisie pour un lettrage en devise
   begin

    if (tiers.getValue('YTC_ACCELERATEUR')='X') then
     begin
      lStCompte := VarToStr(Tiers.GetValue('YTC_SCHEMAGEN')) ;
      if lStCompte = '' then
       lStCompte := Ecr.GetValue('E_CONTREPARTIEGEN') ;
     end  // if
      else
       begin
        lStCompte:=Ecr.GetValue('E_CONTREPARTIEGEN') ;
        if lStCompte = '' then
         lStCompte := VarToStr(Tiers.GetValue('YTC_SCHEMAGEN')) ;
       end ;

    k:=Comptes.GetCompte(lStCompte) ;
    if k > - 1 then
     lStCompte := lStCompte + ' ' + Comptes.getValue('G_LIBELLE') + ' ' ;

    if (tiers.getValue('YTC_ACCELERATEUR')='X') then
     lStCompte := lStCompte + traduireMemoire(' (AUTOMATIQUE) ') ;

    if lStCompte <> '' then
      YTC_SCHEMAGEN.Caption:= TraduireMemoire('Accélérateur') + ' ' + lStCompte ;

   end ;
   end else begin LSA_SOLDET.Visible:=FALSE ; SoldeAuxper.Visible:=FALSE ; end ;
FLASHDEVISE.Caption:='' ; FLASHDEVISE.Visible:=FALSE ;
if (GS.Cells[SF_DEVISE, RefRow]<>'') and (ObjFolio<>nil) then
   begin
   Ecr:=ObjFolio.GetRow(RefRow-GS.FixedRows) ;
   if Ecr<>nil then
      begin
      if memDevise.Code<>GS.Cells[SF_DEVISE, RefRow] then
        begin
        memDevise.Code:=GS.Cells[SF_DEVISE, RefRow] ; GetInfosDevise(memDevise) ;
        memDevise.Taux:=GetTaux(memDevise.Code,memDevise.DateTaux,EncodeDate(memOptions.Year, memOptions.Month, 1)) ;
        end ;
      MDevise:=Ecr.GetValue('E_DEBITDEV')+Ecr.GetValue('E_CREDITDEV') ;
      FLASHDEVISE.Caption:=memDevise.Libelle+': '
                          +StrFMontant(MDevise,15,memDevise.Decimale,'',TRUE)
                          +' ('+memDevise.Symbole+') ' ;
      FLASHDEVISE.Visible:=TRUE ;
      end ;
   end ;
end ;



procedure TFSaisieFolio.SetRowLib(Row : LongInt) ;
var NumCompte : string ; k : Integer ;
begin
if GS.Cells[SF_LIB, Row]<>'' then Exit ;
NumCompte:=GS.Cells[SF_AUX, Row] ; k:=Tiers.GetCompte(NumCompte) ;
if k<0 then
   begin
   NumCompte:=GS.Cells[SF_GEN, Row] ; k:=Comptes.GetCompte(NumCompte) ;
   if k>=0 then GS.Cells[SF_LIB, Row]:=Comptes.GetValue('G_LIBELLE', k) ;
   end else GS.Cells[SF_LIB, Row]:=Tiers.GetValue('T_LIBELLE', k) ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 26/06/2002
Modifié le ... : 28/10/2004
Description .. : - 25/06/2002 - gestion de l'affichage complet ou abrege de
Suite ........ : la nature de piece
Suite ........ : - 28/10/2004 - LG - on ne recalcul pas e_contrepartiegen 
Suite ........ : et aux en mode guide, fait une seul fois a la fin
Mots clefs ... : 
*****************************************************************}
function TFSaisieFolio.SetRowTZF(Row : LongInt ; bCalculerContrepartie : boolean = true) : TZF ;
var Ecr : TZF ; k : Integer ; NumCompte, Gen, Aux : string ;
begin
if ObjFolio=nil then begin Result:=nil ; Exit ; end ;
Ecr:=ObjFolio.GetRow(Row-GS.FixedRows) ;
if Ecr=nil then begin Result:=nil ; Exit ; end ;
{$IFDEF TT}
AddEvenement('SetRowTZF Row:='+inttoStr(Row)) ;
{$ENDIF}
try //LG*
// Entête / Commun
Ecr.PutValue('E_EXERCICE',     memOptions.Exo) ;
Ecr.PutValue('E_JOURNAL',      memJal.Code) ;
Ecr.PutValue('E_NUMEROPIECE',  NumFolio) ;
Ecr.PutValue('E_MODESAISIE',   memJal.ModeSaisie) ;
Ecr.PutValue('E_PERIODE',      GetPeriode(GetRowDate(Row))) ;
Ecr.PutValue('E_SEMAINE',      NumSemaine(GetRowDate(Row))) ;
Ecr.PutValue('E_NUMGROUPEECR', GS.Cells[SF_PIECE, Row]) ;
if GS.Cells[SF_DEVISE, Row]='' then Ecr.PutValue('E_DEVISE', memPivot.Code) ;
if Ecr.GetValue('E_DEVISE')<>memPivot.Code then
   begin
    if memDevise.Code<>Ecr.GetValue('E_DEVISE') then
     begin
     memDevise.Code:=Ecr.GetValue('E_DEVISE') ;
     GetInfosDevise(memDevise) ;
     memDevise.Taux:=GetTaux(memDevise.Code,memDevise.DateTaux,EncodeDate(memOptions.Year, memOptions.Month, 1)) ;
     end ;
   Ecr.SetMontants(Ecr.GetValue('E_DEBITDEV'), Ecr.GetValue('E_CREDITDEV'),
                   memDevise, {bModeOppose, }FALSE) ;
   end else
   begin
    if FInfo.Devise.Dev.Taux=0 then
     FInfo.Devise.AffecteTaux(EncodeDate(memOptions.Year, memOptions.Month, 1)) ;
     //memSaisie.Taux:=GetTaux(memSaisie.Code,memSaisie.DateTaux,EncodeDate(memOptions.Year, memOptions.Month, 1)) ;
   Ecr.SetMontants(Valeur(GS.Cells[SF_DEBIT, Row]), Valeur(GS.Cells[SF_CREDIT, Row]),
                          FInfo.Devise.Dev,{ bModeOppose, }FALSE) ;
   end ;
if Ecr.GetValue('E_DEVISE')=''  then Ecr.PutValue('E_DEVISE',  FInfo.Devise.Dev.Code) ;
if Ecr.GetValue('E_TAUXDEV')=0  then Ecr.PutValue('E_TAUXDEV', FInfo.Devise.Dev.Taux) ;
if Ecr.GetValue('E_DATETAUXDEV')=iDate1900
   then Ecr.PutValue('E_DATETAUXDEV', FInfo.Devise.Dev.DateTaux) ;
Ecr.PutValue('E_CONTROLETVA',   'RIE') ;
Ecr.PutValue('E_ECRANOUVEAU',   'N') ;
// Corps
// GG COM
Ecr.PutValue('E_IO','X') ;
Ecr.PutValue('E_DATECOMPTABLE', GetRowDate(Row)) ;
if (GS.Cells[SF_NUML, Row]<>'') then
  begin
  Ecr.PutValue('E_NUMLIGNE', StrToInt(GS.Cells[SF_NUML, Row])) ;
  if Ecr.GetValue('E_ANA') = 'X' then
   ObjFolio.SetAnaOrdre(Row-GS.FixedRows, StrToInt(GS.Cells[SF_NUML, Row])) ;
  end ;
Ecr.PutValue('E_GENERAL',       GS.Cells[SF_GEN, Row]) ;
Ecr.PutValue('E_AUXILIAIRE',    GS.Cells[SF_AUX, Row]) ;
Ecr.PutValue('E_REFINTERNE',    GS.Cells[SF_REFI, Row]) ;
Ecr.PutValue('E_LIBELLE',       GS.Cells[SF_LIB, Row]) ;
Ecr.PutValue('E_NATUREPIECE',   GetCellNature(Row));
if GS.Cells[SF_AUX, Row]<>'' then
   begin
   NumCompte:=GS.Cells[SF_AUX, Row] ; k:=Tiers.GetCompte(NumCompte) ;
   if k>-1 then begin
   Ecr.PutValue('E_TYPEMVT', QuelTypeMvt(GS.Cells[SF_AUX, Row], Tiers.GetValue('T_NATUREAUXI', k), Ecr.GetValue('E_NATUREPIECE'))) ; end ;
   end else
   begin
   NumCompte:=GS.Cells[SF_GEN, Row] ; k:=Comptes.GetCompte(NumCompte) ;
   if k>-1 then begin
   Ecr.PutValue('E_TYPEMVT', QuelTypeMvt(GS.Cells[SF_GEN, Row], Comptes.GetValue('G_NATUREGENE', k), Ecr.GetValue('E_NATUREPIECE'))) ; end ;
   end ;
Ecr.PutValue('E_ETABLISSEMENT', GS.Cells[SF_ETABL, Row]) ;
// Colonnes corps non obligatoire
// N=Normal ; P=Prévision ; R=Révision ; S=Simulation ; U=Situation
Ecr.PutValue('E_QUALIFPIECE', GS.Cells[SF_QUALIF, Row]) ;
if bCalculerContrepartie and not bGuideRun then
 begin
  GetCPartie(Row, GS.Cells[SF_PIECE, Row], Gen, Aux) ;
  Ecr.PutValue('E_CONTREPARTIEGEN', Gen) ;
  Ecr.PutValue('E_CONTREPARTIEAUX', Aux) ;
  if (Gen<>'') or (Aux<>'') then
   begin
    NumCompte:=GS.Cells[SF_AUX, Row] ; k:=Tiers.GetCompte(NumCompte) ;
    if (k>-1) and memJal.Accelerateur and (tiers.getValue('YTC_ACCELERATEUR')='-') then
     begin
      Ecr.PutValue('CPTCEGID', Gen) ;
      Ecr.PutValue('YTCTIERS', tiers.getValue('T_TIERS') ) ;
     end ;
     CReporteVentil(Ecr,'E_CONTREPARTIEGEN',Gen,true) ;
     CReporteVentil(Ecr,'E_CONTREPARTIEAUX',Aux,true) ;
  //  ModifRefAnaPiece(Row-GS.FixedRows, 'CONTREPARTIEGEN',  Gen, GS.Cells[SF_PIECE, Row]) ;
  //  ModifRefAnaPiece(Row-GS.FixedRows, 'CONTREPARTIEAUX',  Aux, GS.Cells[SF_PIECE, Row]) ;
   end ; // if
 end ;

//if Pos('X',Ecr.GetValue('COMPTAG')) > 0 then
 CMAJTOBCompl(Ecr) ;

Ecr.SetCotation ;
Ecr.SetMPACC ;
Result:=Ecr ;
except
 On E : Exception do
  begin
   result:=nil ;
   MessageAlerte('La ligne est incorrecte...'+#10#13+E.Message);
  end;
end;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 26/06/2002
Modifié le ... : 22/01/2003
Description .. : - 25/06/2002 - gestion de l'affichage complet ou abrege de
Suite ........ : la nature de piece
Suite ........ : - 22/01/2003 - on affecte la TOB au TObject de la grille
Mots clefs ... :
*****************************************************************}
function TFSaisieFolio.GetRowTZF(Row : LongInt; bInit : Boolean=FALSE) : Boolean ;
var Ecr : TZF ; i : Integer ; NumCompte, NumAux : string ;
    Year, Month, Day : word ; EcrDebit, EcrCredit : double ;
begin
Result:=FALSE ;
if ObjFolio=nil then Exit ;
Ecr:=ObjFolio.GetRow(Row) ; if Ecr=nil then Exit ;
// Ajustement
Row:=Row+GS.FixedRows ;
// Entête / Commun
// Corps
if FBoNatureComplete then
 begin
  i:=E_NATUREPIECE.Values.IndexOf(Ecr.GetValue('E_NATUREPIECE')) ;
  if i<0 then i:=0 ;
  GS.Cells[SF_NAT, Row]:=E_NATUREPIECE.Items[i] ;
 end
  else
   begin
    GS.Cells[SF_NAT, Row]:=Ecr.GetValue('E_NATUREPIECE');
   end;
DecodeDate(Ecr.GetValue('E_DATECOMPTABLE'), Year, Month, Day) ;
GS.Cells[SF_JOUR, Row]:=IntToStr(Day) ;
//GS.Cells[SF_NUML, Row]:=IntToStr(Row) ;
(*if memParams.bOptim then *)
if Ecr.GetValue('E_NUMLIGNE') <> 0 then
GS.Cells[SF_NUML, Row]:=Ecr.GetValue('E_NUMLIGNE') ;
NumCompte:=Ecr.GetValue('E_GENERAL') ;
GS.Cells[SF_GEN, Row]:=NumCompte ;
// on charge touttes les info sur les comptes a la lecture du folio
// rem : le param memParams.bOptim est toujours faux
//Comptes.GetCompte(NumCompte) ;
if bInit then Ecr.PutValue('OLDGENERAL', Ecr.GetValue('E_GENERAL')) ;
NumAux:=Ecr.GetValue('E_AUXILIAIRE') ;
if length(trim(NumAux))>0 then  // en mode guide, le test precedent, cela remplissais la Cellule avec un #0,
  begin
  Tiers.GetCompte(NumAux) ;
  GS.Cells[SF_AUX, Row]:=NumAux ;
  if bInit then Ecr.PutValue('OLDAUXILIAIRE', Ecr.GetValue('E_AUXILIAIRE')) ;
  end ;
EcrDebit:=Arrondi(Ecr.GetValue('E_DEBIT'), memPivot.Decimale) ;
EcrCredit:=Arrondi(Ecr.GetValue('E_CREDIT'), memPivot.Decimale) ;
if EcrDebit<>0 then GS.Cells[SF_DEBIT, Row]:=StrFMontant(EcrDebit,15,memPivot.Decimale,'',TRUE)
               else GS.Cells[SF_DEBIT, Row]:='' ;
if EcrCredit<>0 then GS.Cells[SF_CREDIT, Row]:=StrFMontant(EcrCredit,15,memPivot.Decimale,'',TRUE)
                else GS.Cells[SF_CREDIT, Row]:='' ;
GS.Cells[SF_REFI, Row]:=Ecr.GetValue('E_REFINTERNE') ;
GS.Cells[SF_LIB, Row]:=Ecr.GetValue('E_LIBELLE') ;
GS.Cells[SF_ETABL,Row]:=Ecr.GetValue('E_ETABLISSEMENT') ;
GS.Cells[SF_QUALIF,Row]:=Ecr.GetValue('E_QUALIFPIECE') ;
// Devise
if Ecr.GetValue('E_DEVISE')<>memPivot.Code then
   GS.Cells[SF_DEVISE, Row]:=Ecr.GetValue('E_DEVISE') ;
GS.Objects[SA_Exo,Row]:=Ecr ;
_QTE1.Text := StrfMontant(Ecr.GetValue('E_QTE1'),15,V_PGI.OkDecP,'',True) ;
_QTE2.Text := StrfMontant(Ecr.GetValue('E_QTE2'),15,V_PGI.OkDecP,'',True) ;
Result:=TRUE ;
end ;


//=======================================================
//==================== Fonctions SQL ====================
//=======================================================
{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 25/09/2002
Modifié le ... : 30/04/2004
Description .. : -25/09/2002- Suppression de la fct du TZFolio et utisation
Suite ........ : de la fct de SaisUtil
Suite ........ : - 30/04/2004 - LG - affectation de V_PGI.IoError en cas
Suite ........ : d'erreur sur la suppression
Mots clefs ... :
*****************************************************************}
procedure TFSaisieFolio.WriteFolio ;
var  i: Integer ; ACol: LongInt ;
    FOldStatut : TActionFiche ;
    lEcr : TOB ;
    lStQualifPiece : string ;
begin
//LG* 11/01/2002
// En consultation, rien à faire ici !
if bModeRO then Exit ; FStLastError:='';
if ObjFolio = nil then
 begin
  messageAlerte('Erreur lors de l''enregistrement') ;
  V_PGI.IoError := oeSaisie ;
  exit ;
 end;
{$IFDEF TT}
AddEvenement('WriteFolio') ;
{$ENDIF}
try
// GG COM
MajIOCom(GS) ;
// Je positionne les colonnes  ACol pas utiliser
if not IsRowValid(nbLignes, ACol, FALSE) then
 begin
  DeleteRow(nbLignes) ;
  Statut:=taModif ;
 end ;
AfficheTitreAvecCommeInfo('Récupération des infos de la grille');
try
for i:=1 to nbLignes do
  begin
   FOldStatut:=Statut ;
   Statut:=taModif ;
   lEcr := SetRowTZF(i) ;
   if lEcr=nil then Continue ;
   Statut:=FOldStatut ;
   if GetRowDate(i)>memJal.LastDateSais then memJal.LastDateSais:=GetRowDate(i) ;
  end ;
//LG* on enregistre en base -> on supprime le verrous sur cette enregistrement
 if memOptions.NewObj then ObjFolio.UnLockFolio ;
AfficheTitreAvecCommeInfo('Suppression du folio en base') ;
FTOBLettrage.ClearDetail ;
if not ObjFolio.Del(memJal) then begin V_PGI.IoError:=oeSaisie ; exit; end ;
bWrited:=FALSE ;
if (not memParams.bLibre) and (nbLignes>1) then bWrited:=TRUE else
  if (memParams.bLibre) and (IsRowValid(1, ACol, FALSE)) then bWrited:=TRUE ;
if (V_PGI.IOError=oeOK)  then
  begin
  (* DEBUT Fiche 4767 : Ahhhh Grrrr !!! *)
  ObjFolio.GetRow(0).PutValue('E_CONTROLEUR', '') ;
  (* FIN Fiche 4767 *)
  AfficheTitreAvecCommeInfo('Enregistrement en base');
  lStQualifPiece := 'N' ;
  if FboFenLett then lStQualifPiece := 'L' ;
  ObjFolio.Write(memJal,lStQualifPiece ) ; //,FBoPasAna) ;
  {$IFNDEF EAGLCLIENT}
  if (V_PGI.IOError=oeOK) then MarquerPublifi(TRUE) ;
  {$ENDIF}
  end ;
finally
  AfficheTitreAvecCommeInfo ;
end;
except
 on E : Exception do
  begin
   MessageAlerte('Erreur lors de l''enregistrement' + #10#13 + E.Message ) ;
   V_PGI.IoError:=oeSaisie ;
  end ;
end;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 30/04/2004
Modifié le ... :   /  /
Description .. : - LG - 30/04/2004 - ajout de la gestion de erreurs
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.WriteCumulCpte ;
begin
// En consultation, rien à faire ici !
if ObjFolio = nil then
 begin
  MessageAlerte('Erreur lors de la mise à jour des comptes') ;
  V_PGI.IoError := oeSaisie ;
  exit ;
 end ;
if bModeRO then Exit ;
try
 if not ObjFolio.MajCumulsCpte(memJal) then
  PGIInfo('Erreur lors de la mise à jour des comptes !' +#10#13 + ObjFolio.StCompteEnErreur  ) ;
except
 on E : Exception do
  begin
   MessageAlerte('Erreur lors de la mise à jour des comptes' + #10#13 + E.Message ) ;
   V_PGI.IoError := oeSaisie ;
  end ;
end ;
end ;

procedure TFSaisieFolio.WriteCumulAno ;
begin
// En consultation, rien à faire ici !
if ObjFolio = nil then
 begin
  MessageAlerte('Erreur lors de la mise à jour des comptes') ;
  V_PGI.IoError := oeSaisie ;
  exit ;
 end ;
if bModeRO then Exit ;
 ObjFolio.MajCumulsAno ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 25/09/2002
Modifié le ... : 24/03/2005
Description .. : -25/09/2002- on test le solde du journal pour le mode libre
Suite ........ : pour tenir compte des ecarts de conversion euro
Suite ........ : - 24/03/2005 - LG - FB 15342 - ajout d'un param pour froce
Suite ........ : la sauvegarde des param
Mots clefs ... : 
*****************************************************************}
{function TFSaisieFolio.BackupFolio( vBoForce : boolean = false ) : Boolean ;
var bSave : Boolean ; i : integer ;
begin
Result:=FALSE ;
if FBoE_NUMEROPIECEExit or FBoPFENMouseDown or FBoGetFolioClick then exit ;

AddEvenement('BackupFolio') ;

if ( not VH^.ZSAUVEFOLIOLOCAL ) then begin Result:=TRUE ; Exit ; end ;
if FAction=taConsult then Exit ;
if bReading or bWriting then Exit ;
if bClosing then Exit ;
if bModeRO then Exit ;
if bGuideRun then Exit ;
if not vBoForce then
begin
if memParams.bLibre then bSave:=ISSoldeJournal    //IsSoldeFolio
                    else bSave:=(IsSoldeJournal and IsSoldeFolio) ;
if (not bSave) or (nbLignes=nbLastSave) then Exit ;
end ; // if
if ObjFolio<>nil then
 begin
  for i:=GS.FixedRows to NbLignes do
   if IsRowOk(i) then
    SetRowBad(i, FALSE)
     else
      SetRowBad(i, TRUE) ;
  Result:=ObjFolio.Backup;
end ;
nbLastSave:=nbLignes ;
end;
}

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 03/03/2004
Modifié le ... :   /  /
Description .. : - 03/02/2004 - LG - changement de gestion du maxjour : on
Suite ........ : peut saisir des date d'exo diff du 1 et derneir jour du mois
Mots clefs ... :
*****************************************************************}
procedure TFSaisieFolio.AlimenteEntete(ParPeriode, ParCodeJal, ParNumFolio : string) ;
var Year, Month, Day : Word ;
begin
{$IFDEF TT}
AddEvenement('AlimenteEntete') ;
{$ENDIF}
   E_DATECOMPTABLE.Enabled:=FALSE ; E_JOURNAL.Enabled:=FALSE ;
   CbFolio.Enabled:=FALSE ;
   E_DATECOMPTABLE.ItemIndex:=E_DATECOMPTABLE.Values.IndexOf(DateToStr(DebutDeMois(StrToDate(ParPeriode)))) ;
   if E_DATECOMPTABLE.ItemIndex<0 then
     begin
     E_DATECOMPTABLE.ItemIndex:=0 ;
     // Renseigner les options
     DecodeDate(StrToDate(E_DATECOMPTABLE.Value), Year, Month, Day) ;
     memOptions.MaxJour:=CGetDaysPerMonth(E_DATECOMPTABLE.Value, FExoVisu ) ;
     memOptions.Year:=Year ;
     memOptions.Month:=Month ;
     end else
     begin
     // Renseigner les options
     memOptions.MaxJour:=CGetDaysPerMonth(ParPeriode, FExoVisu ) ;
     DecodeDate(StrToDate(ParPeriode), Year, Month, Day) ;
     memOptions.Year:=Year ;
     memOptions.Month:=Month ;
     end ;
   E_JOURNAL.Value:=ParCodeJal ;
   // Alimenter le record RJal
   ChargeJal(ParCodeJal) ;
   CalculSoldeJournal ;
   // Mise à jour du combo Folio
   FillComboFolio ;
   if FAction=taConsult then CbFolio.ItemIndex:=CbFolio.Values.IndexOf(ParNumFolio) ;
   CbFolio.Text:=ParNumFolio ;
   NumFolio:=Trunc(Valeur(CbFolio.Text)) ;
   // Utiliser la bonne nature de pièce en fonction de la nature du journal
   case CaseNatJal(memJal.Nature) of
        tzJVente       : E_NATUREPIECE.DataType:='ttNatPieceVente' ;
        tzJAchat       : E_NATUREPIECE.DataType:='ttNatPieceAchat' ;
        tzJBanque      : E_NATUREPIECE.DataType:='ttNatPieceBanque' ;
        tzJEcartChange : E_NATUREPIECE.DataType:='ttNatPieceEcartChange' ;
        tzJOD          : E_NATUREPIECE.DataType:='ttNaturePiece' ;
        end ;
   CurPeriode:=E_DATECOMPTABLE.Value ;
   CurJal:=E_JOURNAL.Value ;
   CurNumFolio:=CbFolio.Text ;
   E_DATECOMPTABLE.Enabled:=TRUE ; E_JOURNAL.Enabled:=TRUE ; CbFolio.Enabled:=TRUE ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 09/01/2002
Modifié le ... : 08/10/2002
Description .. : Modif de la gestion des verrous
Suite ........ :
Suite ........ : - LG - 30/04/2002 - on supprime le verrou lorsque l'on
Suite ........ : recupere une sauvegarde de bordereau
Suite ........ : 
Suite ........ : - 08/10/2002 - La fct Restore est renomme en 
Suite ........ : FolioSauvegarde ( pour ne pas avoir 2 fct diff avec le meme 
Suite ........ : nom )
Mots clefs ... : 
*****************************************************************}
(*function TFSaisieFolio.RestoreFolio : Boolean ;
var Par1, Par2, Par3 : string ;
begin
{$IFDEF TT}
AddEvenement('RestoreFolio') ;
{$ENDIF}
Result:=FALSE ;
if InitParams.ParRecupLettrage or not VH^.ZSAUVEFOLIOLOCAL then exit ;
ObjFolio:=TFolio.Create('', memJal, memOptions, BAutoSave) ;
ObjFolio.Comptes:=Comptes ;
try
if ObjFolio.FolioSauvegarde(Par1, Par2, Par3) then
   begin
    Result:=TRUE ;
    FBoRestore:=true;
    AlimenteEntete(Par1, Par2, Par3) ;
    //LG* 30/04/2002 suppression du verrou
    ObjFolio.UnLockFolio(Par1,Par2,Par3) ;
    ObjFolio.Free ; ObjFolio:=nil ;
   end ;
except
 on E : Exception do
  begin
   Result:=false ;
   FBoRestore:=false;
   PGIInfo('Impossible de remonter le bordereau à partir de la sauvegarde !' + #10#13 + #10#13 + E.Message , Self.Caption ) ;
   exit ;
 end ;
end ;
end ;
*)

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 19/11/2003
Modifié le ... : 19/11/2003
Description .. : - 19/10/2003 - LG - apres la renumerotation du bordereau, 
Suite ........ : on controle la presence d'ecriture en base. le dbclick sur la 
Suite ........ : grille suprimais les ecritures
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.RenumAna ;
var
 bOldModeRO : boolean ;
begin
{$IFDEF TT}
AddEvenement('RenumAna') ;
{$ENDIF}
 bOldModeRO:=bModeRO ;
 bModeRO:=false ;
  try
   ObjFolio.RemuneroteAna ;
   WriteFolio ;
   if not existesql ('SELECT e_exercice FROM ECRITURE WHERE E_NUMEROPIECE='+IntToStr(NumFolio)
             +' AND E_JOURNAL="'+memJal.Code+'"'
             +' AND E_EXERCICE="'+memOptions.Exo+'"'
             +' AND E_DATECOMPTABLE>="'+USDateTime(EncodeDate(memOptions.Year, memOptions.Month, 1))+'"'
             +' AND E_DATECOMPTABLE<="'+USDateTime(EncodeDate(memOptions.Year, memOptions.Month, memOptions.MaxJour))+'"'
             +' AND E_QUALIFPIECE="N" ' ) then
   begin
    v_pgi.IOError:=oesaisie ;
    messagealerte('Problème lors de la mise à jour du folio.#10#13Veuillez recharger le bordereau (Saisie d''écritures - Mode bordereau).') ;
    exit ;
   end ;
   ReadFolio(false) ;
   finally
     bModeRO:=bOldModeRO ;
   end;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 09/01/2002
Modifié le ... : 03/09/2007
Description .. : Modif du lock du folio
Suite ........ : 
Suite ........ : - 12/06/2002 - bloque le rafraichissement de la grille
Suite ........ : 
Suite ........ : - 16/09/2002 - reecriture du code
Suite ........ : 
Suite ........ : -06/11/2002 - le verrouillage du folio ne fct pas
Suite ........ : - 17/06/2003 - en recup de folio ou d'un bordereau avec
Suite ........ : des ecritures de types L, on supprime les eventuels verrou
Suite ........ : - 09/10/2003 - on pouvait modifie un bordereau sur un
Suite ........ : journal fermé
Suite ........ : - 08/09/2004 - FB 14531 - blocage si le bordereau contient 
Suite ........ : un compte valide
Suite ........ : - 14/10/2004 - FB 14634 - ajout d'un message qd le 
Suite ........ : bordereau est cloture
Suite ........ : - 21/07/2005 - FB 14634 - on affichage un msg pour les 
Suite ........ : bordereaux validés qq soit le cas
Suite ........ : - 10/05/2005 - FB 14634 -  on a plus le deux message
Suite ........ : bordereau valide et bordereau bloque
Suite ........ : - LG - 29/05/2007 - FB 15745 - correction du calcul du
Suite ........ : solde de bq en consultation en saisie
Suite ........ : - LG - 05/07/2207 - FB 20918 - reecritue de l'affichage des
Suite ........ : msg de validation/cloture du bordereau
Suite ........ : - LG - 03/09/2007 - FB 15187 - bloque sur l'etablissemebt 
Suite ........ : autorisee
Mots clefs ... : 
*****************************************************************}
function TFSaisieFolio.ReadFolio(bRestore : Boolean) : Boolean ;
var k,i : Integer ;
Ecr : TZF ;
lInErreur : integer ;
begin
{$IFDEF TT}
AddEvenement('ReadFolio') ;
{$ENDIF}
result := false ;
if csDestroying in ComponentState then Exit ;
if V_PGI.IOError <> oeOk then exit ;
bReading:=TRUE ; SourisSablier ; Result:=TRUE ; HGBeginUpdate(GS) ;
try
 // on supprime les verrous sur le folio et on detruit le ObjFolio s'il existait
 InitFolio ;
 FParam.ParPeriode:=E_DATECOMPTABLE.Value ;
 FParam.ParCodeJal:=memJal.Code ;
 FParam.ParNumFolio:=IntToStr(NumFolio) ;
 if FBoLettrageEnCours and ( not InitParams.ParRecupLettrage ) then memOptions.QualifP:='L' else memOptions.QualifP:='N' ;
 ObjFolio           := TFolio.Create(IntToStr(NumFolio), memJal, memOptions ) ;
 ObjFolio.Comptes := Comptes ;
 ObjFolio.Tiers   := Tiers ;
 if bRestore or InitParams.ParRecupLettrage then ObjFolio.UnLockFolio;
 memOptions.NewObj  := true ;
 if InitParams.ParGuide = '' then
  memOptions.NewObj  := not ObjFolio.Read(bRestore, memJal.TotSavDebit, memJal.TotSavCredit,memParams.bLibre,FBoLettrageEnCours,InitParams.ParRecupLettrage) ;
 if V_PGI.IOError=oeLettrage then
  begin
   V_PGI.IOError:=oeOK ;
   InitFolio ;
   PrintMessageRC(RC_BORCONF) ;
   result := false ;
   exit ;
  end ; // if

 if FBoLettrageEnCours and ( not InitParams.ParRecupLettrage ) then
  begin
   HGBeginUpdate(GS) ;
   for i:=1 to GS.RowCount - 1 do GS.Cells[SF_QUALIF, i] := 'L' ;
   HGEndUpdate(GS) ;
  end ; // if

 // test si le bordereau est en consultation ou detruit
 bModeRO:= (FAction=taConsult) or memJal.Ferme or ( (not memOptions.NewObj) and (ObjFolio.GetRow(0).GetValue('E_CREERPAR')='DET') ) ;

 if (not memOptions.NewObj)  and
    ( ( ObjFolio.GetRow(0).GetValue('E_DATECOMPTABLE') < GetEncours.Deb ) or
    ( GetParamSocSecur('SO_DATECLOTUREPER',iDate1900,true) >= ObjFolio.GetRow(0).GetValue('E_DATECOMPTABLE') ) ) then
  PrintMessageRC(RC_PERIODECLOSE)
   else
    if (not memOptions.NewObj) and (ObjFolio.GetRow(0).GetValue('E_VALIDE')='X') then
     begin
      bModeRO:=TRUE ;
      PrintMessageRC(RC_BORCLOTURE) ;
     end; // if


 // GP Com
 if (not bModeRO) and (not memOptions.NewObj) and (ObjFolio.GetRow(0).GetValue('E_ETATREVISION')='X') then
  begin
   bModeRO:=TRUE ;
   PrintMessageRC(RC_BORREVISION) ;
  end ;

 // on verrouille le folio
 if not bModeRO and ( not OBjFolio.LockFolio ) then
  bModeRO:=TRUE ;

 // Ok, les paramètres courants peuvent être mis à jour
 if (not memOptions.NewObj) and (memParams.bLibre) and
   (ObjFolio.GetRow(0).GetValue('E_DEVISE')<>V_PGI.DevisePivot) then
  begin
   memParams.bDevise   :=TRUE ;
   memParams.DevLibre  :=ObjFolio.GetRow(0).GetValue('E_DEVISE') ;
  end ;

 if (memOptions.NewObj) then
  begin
   E_ETABLISSEMENT.Enabled:= ( E_ETABLISSEMENT.Items.Count > 1 ) ;
   if FAction=taCreat then PositionneEtabUser(E_ETABLISSEMENT) ;
  end
   else
    begin
     E_ETABLISSEMENT.Enabled:=FALSE ; // on ne modidife pas l'etablissement d'un folio existant
     E_ETABLISSEMENT.Value:=ObjFolio.GetRow(0).GetValue('E_ETABLISSEMENT') ;
     lInErreur := CIsValidEtabliss(ObjFolio.GetRow(0),FRechCompte.Info) ; 
     if lInErreur <> RC_PASERREUR then  // FB 15187
      begin
       PrintMessageRC(lInErreur) ;
       result := false ;
       exit ;
      end ;
    end ;

 CurPeriode    := E_DATECOMPTABLE.Value ;
 CurJal        := E_JOURNAL.Value ;
 CurNumFolio   := CbFolio.Text ;

 // Alimenter le grid
 k:=0 ; memJal.TotDebDebit:=0; memJal.TotDebCredit:=0 ; SoldeTresor:=0 ;
 if ObjFolio.Ecrs.Detail.Count+5>=GS.RowCount then GS.RowCount:=ObjFolio.Ecrs.Detail.Count+10 ;
 FBoPasAna:=false ; FRdSoldeCBq := 0 ; FRdSoldeDBq := 0 ;

 while GetRowTZF(k, TRUE) do
  begin
   nbLignes:=nbLignes+1 ;
   Ecr:=ObjFolio.GetRow(k) ; if Ecr=nil then Exit ;
   if not FBoPasAna then
    FBoPasAna := (nbLignes<>Ecr.GetValue('E_NUMLIGNE')) ;
   k:=k+1 ;
   // (PFU : !) Valeurs non présentes dans le TZF
   GS.Cells[SF_NEW,  k]:='N' ; // Ligne non créée en saisie
   // Folio au début de la saisie
   memJal.TotDebDebit:=memJal.TotDebDebit+Valeur(GS.Cells[SF_DEBIT, k]) ;
   memJal.TotDebCredit:=memJal.TotDebCredit+Valeur(GS.Cells[SF_CREDIT, k]) ;
   if memParams.bLibre then GS.Cells[SF_PIECE,k]:='1' ;
   if (memParams.bSoldeProg) and (GS.Cells[SF_GEN, k]=memJal.CContreP) then
    SoldeTresor:=SoldeTresor+Valeur(GS.Cells[SF_DEBIT, k])-Valeur(GS.Cells[SF_CREDIT, k]) ;
   if ( not FboFenLett ) and (memParams.bSoldeProg) and (GS.Cells[SF_GEN, k]<>memJal.CContreP) then  // FB 15745
    begin
     FRdSoldeCBq := FRdSoldeCBq + Valeur(GS.Cells[SF_DEBIT, k]) ;
     FRdSoldeDBq := FRdSoldeDBq + Valeur(GS.Cells[SF_CREDIT, k]) ;
    end ;
  end ; // while

 nbLastSave:=nbLignes ;
 if not memParams.bLibre then SetPiece ;

 if FBoPasAna then
  begin
    if transactions(RenumAna,1 )<>oeOk then
     MESSAGEALERTE('Mise à jour des numéros de ligne impossible !' ) ;
    FBoPasAna:=false ;
  end ;

 GrilleModif:=false ;

finally
 GS.Invalidate ; SourisNormale ; HGEndUpdate(GS) ; FStGenHT := '' ;
 bReading := false ;
end;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 26/05/2003
Modifié le ... :   /  /
Description .. : - 26/05/2003 - on ne prens en compte que les ecritures de
Suite ........ : type N pour calcul le numero de folio
Mots clefs ... :
*****************************************************************}
procedure TFSaisieFolio.FillComboFolio ;
begin
CbFolio.Items.Clear ; CbFolio.Values.Clear ; CbFolio.Text:='' ;
if memJal.Code='' then Exit ;
CRempliComboFolio ( CbFolio.Items, CbFolio.Values , memJal.Code,memOptions.Exo , strToDate(E_DATECOMPTABLE.Value) ) ;
NumFolio:=Trunc(Valeur(CbFolio.Text)) ;
if NumFolio=0 then NumFolio:=1 ;
end ;

//=======================================================
//================ Saisie complémentaires ===============
//=======================================================
function TFSaisieFolio.IsEch(Row : LongInt) : Boolean ;
var Ecr : TZF ; NumCompte, NumAux : string ; k : Integer ;
begin
Result:=FALSE ;
// Collectif + Auxiliaire lettrable
if ObjFolio=nil then Exit ;
Ecr:=ObjFolio.GetRow(Row-GS.FixedRows) ; if Ecr=nil then Exit ;
NumCompte:=Ecr.GetValue('E_GENERAL') ; if NumCompte='' then Exit ;
k:=Comptes.GetCompte(NumCompte) ;
if k<0 then Exit ;
// TIC/TID
if (Comptes.IsTiers(k)) and (Comptes.GetValue('G_LETTRABLE', k)='X') then
  begin Result:=TRUE ; Exit ; end ;
if Comptes.GetValue('G_COLLECTIF', k)='X' then
   begin
   NumAux:=Ecr.GetValue('E_AUXILIAIRE') ; if NumAux='' then Exit ;
   k:=Tiers.GetCompte(NumAux) ; if k<0 then Exit ;
   if Tiers.GetValue('T_LETTRABLE', k)='-' then Exit ;
   end else if (IsJalBqe) and (GS.Cells[SF_GEN, Row]=memJal.CContreP) then else Exit ;
Result:=TRUE ;
end ;

// bAutoOpen ouvre la boite de dialogue que si nécessaire
// bAutoQuiet est prioritaire et renseigne automatiquement sans ouvrir la BdD
{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 05/04/2002
Modifié le ... : 28/07/2005
Description .. : Correction : on affecte les param d'echeance et de lettrage
Suite ........ : si ceux ci ne sont pas definies :  initialisation de l'ecriture (
Suite ........ : sinon les champs de lettrage ne sont pas
Suite ........ : correctement renseignes )
Suite ........ : 
Suite ........ : - 17/09/2002 - correction de l'affectation des valeurs de
Suite ........ : pointage : sur les compte de banque ou de caisse et
Suite ........ : seulement s'il sont pointable
Suite ........ : 
Suite ........ : - 20/11/2002-fiche 10590 - remise à zero des info si le 
Suite ........ : compte n'est pas 
Suite ........ : lettrable
Suite ........ : - 01/10/2004 - reaffectfation du E_MODEPAIE si celui ci 
Suite ........ : n'est pas renseigner
Suite ........ : - 28/07/2005 - LG - FB 15998 - on ouvre pas la fenetre des 
Suite ........ : echenaces pour les comptes lettrables
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.GetEch(Row : LongInt; bAutoOpen, bAutoQuiet : Boolean) ;
var Ecr : TZF ; Ech : TDateTime ; ModeR : T_MODEREGL ; ZDev : RDEVISE ;
    Mode, NumCompte, NumAux, ModeAux : string ; k, kt : Integer ; bRes, bBanque : Boolean ;
    Action : TActionFiche ; nLigne : LongInt ;
    DateEcr : tDateTime ;
begin
ModeAux:='' ; bBanque:=FALSE ; kt:=-1 ;
// Collectif + Auxiliaire lettrable
if ObjFolio=nil then Exit ; Ecr:=ObjFolio.GetRow(Row-GS.FixedRows) ; if Ecr=nil then Exit ;
NumCompte:=Ecr.GetValue('E_GENERAL') ; if NumCompte='' then Exit ; k:=Comptes.GetCompte(NumCompte) ; if k<0 then Exit ;
if (Comptes.GetValue('G_COLLECTIF', k)='X') then
 begin
   NumAux:=Ecr.GetValue('E_AUXILIAIRE') ; if NumAux='' then Exit ;
   kt:=Tiers.GetCompte(NumAux) ; if kt<0 then Exit ;
   if Tiers.GetValue('T_LETTRABLE', kt)='-' then Exit ;
 end
  else
   if (Comptes.GetValue('G_POINTABLE', k)='X') and ( (Comptes.GetValue('G_NATUREGENE', k)='BQE') or (Comptes.GetValue('G_NATUREGENE', k)='CAI') ) then //(GS.Cells[SF_GEN, Row]=memJal.CContreP) then
    begin
     bBanque:=TRUE ;
     nLigne:=GetRowTiers(StrToInt(GS.Cells[SF_PIECE, Row]), Row) ;
     if nLigne>0 then
      begin
       ModeAux:=GS.Cells[SF_AUX, nLigne] ;
       kt:=Tiers.GetCompte(NumAux) ; if kt<0 then ModeAux:='' ;
      end ;
    end
     else
      //if not((Comptes.IsTiers(k)) and (Comptes.GetValue('G_LETTRABLE', k)='X')) then
     if (Comptes.GetValue('G_LETTRABLE', k)='-') then // LG 01/07/2005 le test precedent ne fct pas avec les divers lettrabel
       begin
        CSupprimerInfoLettrage(Ecr) ; Ecr.PutValue('E_MODEPAIE', '') ; exit ;
       end;
if Comptes.GetValue('G_NATUREGENE') = 'DIV' then bAutoQuiet := true ; // on force a false bAutoOpen pour les compte divers lettrable
Mode:=Ecr.GetValue('E_MODEPAIE') ; Ech:=Ecr.GetValue('E_DATEECHEANCE') ;
DateEcr:=Ecr.GetValue('E_DATECOMPTABLE') ;
if (not bAutoQuiet) and (bAutoOpen) and (Mode<>'') then Exit ;
// Valeurs par défaut
if (Ech=EncodeDate(1900, 1, 1)) or (Mode='') then  // LG 01/10/2004 - réffectation du e_modepaie si celui ce n'etait pas renseigner
 begin
  Ech:=GetRowDate(Row) ;
  ModeR.Action:=taCreat ;
  if (kt>=0) then ModeR.ModeInitial:=Tiers.GetValue('T_MODEREGLE', kt) ;
  ModeR.TotalAPayerP:=Ecr.GetValue('E_DEBIT')+Ecr.GetValue('E_CREDIT') ;
  ModeR.TotalAPayerD:=Ecr.GetValue('E_DEBITDEV')+Ecr.GetValue('E_CREDITDEV') ;
  if Ecr.GetValue('E_DEVISE')<>FInfo.Devise.Dev.Code then
   begin
    if memDevise.Code<>Ecr.GetValue('E_DEVISE') then
     begin
      memDevise.Code:=Ecr.GetValue('E_DEVISE') ; GetInfosDevise(memDevise) ;
     end ;
    memDevise.Taux:=GetTaux(memDevise.Code, memDevise.DateTaux, Ecr.GetValue('E_DATECOMPTABLE')) ;
    ZDev:=memDevise ;
   end
    else
     ZDev:=FInfo.Devise.Dev ;
  ModeR.CodeDevise:=ZDev.Code ; ModeR.Symbole:=ZDev.Symbole ;
  ModeR.Quotite:=ZDev.Quotite ; ModeR.TauxDevise:=ZDev.Taux ;
  ModeR.Decimale:=ZDev.Decimale ;
  ModeR.DateFact:=Ech ; ModeR.DateBL:=Ech ; ModeR.DateFactExt:=Ech ;
  if ModeAux<>'' then ModeR.Aux:=ModeAux
                 else ModeR.Aux:=Ecr.GetValue('E_AUXILIAIRE') ;
  if ModeR.Aux='' then ModeR.Aux:=Ecr.GetValue('E_GENERAL') ;
  CalculModeregle(ModeR, FALSE) ;
  Ech:=ModeR.TabEche[1].DateEche ;
  Mode:=ModeR.TabEche[1].ModePaie ;
 end ;
bRes:=FALSE ;
if ((bModeRO) or (IsRowCanEdit(Row)>0)) then Action:=taConsult else Action:=taModif ;
if not bAutoQuiet then
 bRes:=SaisieZEch(Mode, Ech, Action,DateEcr) ;
//LG* - 05/04/2002 - on passe dasn la boucle si la date echeance n'est pas definie - init de l'ecriture
if (bAutoQuiet) or (bRes) or ( Ecr.GetValue('E_DATEECHEANCE')=iDate1900) then
   begin
   if bBanque then
     begin
     Ecr.PutValue('E_MODEPAIE', Mode) ; Ecr.PutValue('E_DATEECHEANCE', Ech) ;
                                        Ecr.PutValue('E_ORIGINEPAIEMENT', Ech) ;
     Ecr.PutValue('E_ECHE', 'X') ;      Ecr.PutValue('E_NUMECHE', 1) ;
     Ecr.PutValue('E_ETATLETTRAGE' , 'RI') ;
     Ecr.PutValue('E_ENCAISSEMENT',  IsRowDecEnc(Row)) ;
     Ecr.PutValue('E_DATEVALEUR',    Ecr.GetValue('E_DATECOMPTABLE')) ;
     Ecr.PutValue('E_DATEPAQUETMIN', Ecr.GetValue('E_DATECOMPTABLE')) ;
     Ecr.PutValue('E_DATEPAQUETMAX', Ecr.GetValue('E_DATECOMPTABLE')) ;
     LastModePaie:=Mode ;
     end else
     begin
     Ecr.PutValue('E_MODEPAIE', Mode) ; Ecr.PutValue('E_DATEECHEANCE', Ech) ;
                                        Ecr.PutValue('E_ORIGINEPAIEMENT', Ech) ;
     Ecr.PutValue('E_ECHE', 'X') ;      Ecr.PutValue('E_NUMECHE', 1) ;
     if Ecr.GetValue('E_ETATLETTRAGE')='RI' then Ecr.PutValue('E_ETATLETTRAGE', 'AL') ;
     Ecr.PutValue('E_ENCAISSEMENT',  IsRowDecEnc(Row)) ;
     Ecr.PutValue('E_DATEVALEUR',    Ecr.GetValue('E_DATECOMPTABLE')) ;
     Ecr.PutValue('E_DATEPAQUETMIN', Ecr.GetValue('E_DATECOMPTABLE')) ;
     Ecr.PutValue('E_DATEPAQUETMAX', Ecr.GetValue('E_DATECOMPTABLE')) ;
     LastModePaie:=Mode ;
     end ;
   end ;
GS.Invalidate ;
end ;

procedure TFSaisieFolio.ModifRefAnaPiece(Row: LongInt; NomChamp: string; Val: Variant; Piece : string) ;
var Ecr: TZF ; i : Integer ;
begin
if ObjFolio = nil then exit ;
Ecr:=ObjFolio.GetRow(Row) ; if Ecr=nil then Exit ;
for i:=GS.FixedRows to nbLignes do
  begin
  if GS.Cells[SF_PIECE, i]<>Piece then Continue ;
  ModifRefAna(i-GS.FixedRows, NomChamp, Val) ;
  end ;
end ;

procedure TFSaisieFolio.ModifRefAna(Row: LongInt; NomChamp: string; Val: Variant) ;
var Ecr: TZF ; i, j : Integer ;
begin
if ObjFolio = nil then exit ;
Ecr:=ObjFolio.GetRow(Row) ; if Ecr=nil then exit ; if Ecr.GetValue('E_ANA')<>'X' then Exit ;
if (not VarIsNull(Val)) and (Ecr.GetValue('E_'+NomChamp)=Val) then Exit ;
if (Ecr<>nil) then
  begin
  for i:=0 to Ecr.Detail.Count-1 do
    for j:=0 to (Ecr.Detail[i].Detail).Count-1 do
      if VarIsNull(Val)
        then (Ecr.Detail[i].Detail[j]).PutValue('Y_'+NomChamp, Ecr.GetValue('E_'+NomChamp))
        else (Ecr.Detail[i].Detail[j]).PutValue('Y_'+NomChamp, Val)
  end ;
end ;



{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 25/02/2004
Modifié le ... : 13/07/2005
Description .. : - 25/02/2004 - LG - passage en eAgl
Suite ........ : - 13/07/2005 - LG - FB 15968 - l'appel de l'analytiq par 
Suite ........ : ctrl+a directment de la zone montant avant de valider la 
Suite ........ : ligne plantait
Mots clefs ... :
*****************************************************************}
//procedure TFSaisieFolio.GetAna(Row : LongInt; bAutoOpen : Boolean ; vOuiTvaEnc : boolean ) ;
procedure TFSaisieFolio.GetAna(Row : LongInt ) ;
var Ecr : TZF ; Action : TActionFiche ; Compte : TGGeneral ;
    RecA : ARG_ANA ; NumCompte : string ; k: Integer ;
begin
if ObjFolio=nil then Exit ;
Ecr:=ObjFolio.GetRow(Row-GS.FixedRows) ; if Ecr=nil then Exit ;
NumCompte:=Ecr.GetValue('E_GENERAL') ; if NumCompte='' then Exit ;
k:=Comptes.GetCompte(NumCompte) ; if k = - 1 then exit ;
//if Ecr.Detail.Count = 0 then
// for i:=1 to MAXAXE do TZF.Create('A'+IntToStr(i), Ecr, -1) ;
//if (bAutoOpen) and (ObjFolio.IsAnaFromRow(Row-GS.FixedRows)) then Exit ;
if ((bModeRO) or (IsRowCanEdit(Row)>0)) then Action:=taConsult else Action:=taModif ;
if bGuideRun then
   begin
   RecA.GuideA:=SELGUIDE.Text ;
   RecA.NumLigneDecal:=GetRowFromPiece(Row, StrToInt(GS.Cells[SF_PIECE, Row])) ;
   Action:=taCreat ;
   end ;
if Ecr.GetValue('E_DEVISE')<>FInfo.Devise.Dev.Code then
   begin
   if memDevise.Code<>Ecr.GetValue('E_DEVISE') then
     begin
     memDevise.Code:=Ecr.GetValue('E_DEVISE') ; GetInfosDevise(memDevise) ;
     end ;
   memDevise.Taux:=GetTaux(memDevise.Code, memDevise.DateTaux, Ecr.GetValue('E_DATECOMPTABLE')) ;
   RecA.DEV:=memDevise ;
   end else RecA.DEV:=FInfo.Devise.Dev ;
Compte:=  CZompteVersTGGeneral ( FInfo.Compte ) ; //TGGeneral.Create(Comptes.GetValue('G_GENERAL', k)) ;
// Paramètres d'ouvertures :
RecA.QuelEcr:=EcrGen ; RecA.CC:=Compte ;
RecA.Action:=Action ;
RecA.NumGeneAxe:=0 ;
RecA.VerifVentil:=TRUE ;
// Jamais de vérification de Qté en mode PCL // Dev 3946 param soc au lieu du scenario
RecA.VerifQte := GetParamSocSecur('SO_ZCTRLQTE', False) ;
RecA.Info           := FInfo ;
RecA.MessageCompta  := nil ;
eSaisieAnal(TOB(Ecr),RecA);
if not bModeRO then
 begin
  ArrondirAnaTOB(Ecr,V_PGI.OkDecV) ; // remplace le ventilerTob, seul cette fct etait utiliser ds le ventilerTOB
  ObjFolio.PostAnaCreate(Row-GS.FixedRows) ;
  ModifRefAna(Row-GS.FixedRows, 'CONTREPARTIEGEN',  Null) ; // Permet de forcer la mise à jour
  ModifRefAna(Row-GS.FixedRows, 'CONTREPARTIEAUX',  Null) ; // Permet de forcer la mise à jour
  {JP 09/11/07 : FQ 21812 : mise à jour de l'auxiliaire}
  ModifRefAna(Row-GS.FixedRows, 'AUXILIAIRE'     ,  Null) ;
 end;

Compte.Free ;
end ;


procedure TFSaisieFolio.PutRowProrata(Row : LongInt) ;
var Ecr : TZF ;
begin
if ObjFolio=nil then Exit ;
Ecr:=ObjFolio.GetRow(Row-GS.FixedRows) ; if Ecr=nil then Exit ;
// TODO : Echéance
if Ecr.GetValue('E_ANA')='X' then ProraterVentilsTOB(Ecr, FInfo.Devise.Dev.Decimale) ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 08/10/2002
Modifié le ... : 19/07/2005
Description .. : - 08/10/2002 - on ne pas pouvoir saisir une immo sans
Suite ........ : montant
Suite ........ : - LG - 19/07/2005 - FB 10054 - on ne saisie plus d'immo sur 
Suite ........ : une ligne crediteur
Suite ........ : - on ne saisie plus d'immo pour le Avoir fournisseur
Mots clefs ... : 
*****************************************************************}

procedure TFSaisieFolio.PutRowImmo(Row : LongInt) ;
{$IFDEF AMORTISSEMENT}
var Ecr : TZF ; TImmo : TOB ; k : Integer ; nLigne : LongInt ;
    NumCompte, NumAux, LibAux, CodeImmo : string ; Val : Double ;
{$ENDIF}
{$IFDEF COMPTAAVECSERVANT}
var RecImmo1 : TBaseImmo1 ;
{$ENDIF}
begin
{$IFDEF AMORTISSEMENT}
if ((not VH^.OkModImmo) and (not V_PGI.VersionDemo)) then Exit ;
if ObjFolio=nil then Exit ;
{$IFDEF TT}
AddEvenement('PutRowImmo') ;
{$ENDIF}
Ecr:=ObjFolio.GetRow(Row-GS.FixedRows) ; if Ecr=nil then Exit ;
if GS.Cells[SF_NEW, Row]<>'O' then Exit ;
if ((Valeur(GS.Cells[SF_DEBIT,Row])=0) and (Valeur(GS.Cells[SF_CREDIT,Row])=0)) then exit ; // on ne pas saisir une immmo sans montant
if bGuideRun then Exit ;
if Ecr.GetValue('E_DEVISE')<>V_PGI.DevisePivot then Exit ;
if memJal.Code='' then Exit ;
if memJal.Nature<>'ACH' then
   begin
   if ((memJal.Nature<>'OD') or
       (GetCellNature(Row)<>'FF')) then Exit ;
   end ;
if Valeur(GS.Cells[SF_DEBIT,Row]) = 0 then exit ;
if (GetCellNature(Row)='AF') then Exit ;
NumCompte:=Ecr.GetValue('E_GENERAL') ; if NumCompte='' then Exit ;
k:=Comptes.GetCompte(NumCompte) ;
if k<0 then Exit ;
if Comptes.GetValue('G_NATUREGENE', k)<>'IMO' then Exit ;
if Ecr.GetValue('E_NUMEROIMMO')=666 then Exit else Ecr.PutValue('E_NUMEROIMMO', 666) ;
if Ecr.GetValue('E_IMMO')<>'' then Exit ;
if (PrintMessageRC(RC_NEWIMMO)<>mrYes) then Exit ;
Val:=Ecr.GetValue('E_DEBIT')+Ecr.GetValue('E_CREDIT') ;
nLigne:=GetRowTiers(StrToInt(GS.Cells[SF_PIECE, Row]), Row) ; LibAux:='' ;
if nLigne>0 then
  begin
  NumAux:=GS.Cells[SF_AUX, nLigne] ; k:=Tiers.GetCompte(NumAux) ;
  if k>=0 then LibAux:=Tiers.GetValue('T_LIBELLE', k) ;
  end ;
TImmo := TOB.Create('IMMO',nil,-1);
TImmo.PutValue('I_DATEAMORT',GetRowDate(Row));
TImmo.PutValue('I_DATEPIECEA',GetRowDate(Row));
TImmo.PutValue('I_COMPTEIMMO',NumCompte);
TImmo.PutValue('I_MONTANTHT', Val);
TImmo.PutValue('I_TVARECUPERABLE',0);
TImmo.PutValue('I_TIERSA',LibAux);
TImmo.PutValue('I_REFINTERNEA',GS.Cells[SF_REFI, Row]);
TImmo.PutValue('I_VALEURACHAT',Val);
TImmo.PutValue('I_QUANTITE',Ecr.GetValue('E_QTE1'));
if (TImmo.GetValue('I_QUANTITE')=0) then TImmo.PutValue('I_QUANTITE',1);
TImmo.PutValue('I_ETABLISSEMENT',GS.Cells[SF_ETABL, Row]);
TImmo.PutValue('I_LIBELLE',GS.Cells[SF_LIB, Row]);
{$IFDEF COMPTAAVECSERVANT}
IF VHIM^.ServantPGI Then
  BEGIN
  RecImmo1.CompteImmo:=TImmo.GetValue('I_COMPTEIMMO');
  RecImmo1.Fournisseur:=NumAux ;
  RecImmo1.Reference:=TImmo.GetValue('I_REFINTERNEA');
  RecImmo1.Libelle:=TImmo.GetValue('I_LIBELLE');
  RecImmo1.CodeEtab:=TImmo.GetValue('I_ETABLISSEMENT');
  RecImmo1.DateAchat:=TImmo.GetValue('I_DATEPIECEA');
  RecImmo1.MontantHT:=TImmo.GetValue('I_MONTANTHT');
  RecImmo1.MontantTVA:=TImmo.GetValue('I_TVARECUPERABLE');
  RecImmo1.Quantite:=TImmo.GetValue('I_QUANTITE');
  Codeimmo:=SaisieExterne(taCreat,0,RecImmo1) ;
  END Else Codeimmo:=AMLanceFiche_FicheImmobilisationEnSaisie ( TImmo);
{$ELSE}
CodeImmo:=AMLanceFiche_FicheImmobilisationEnSaisie ( TImmo);
{$ENDIF}
TImmo.Free;
Ecr.PutValue('E_IMMO', CodeImmo) ;
{$ENDIF}
end ;

function TFSaisieFolio.GetGuideColStop(Row : LongInt) : LongInt ;
var Ecr : TZF ; Arret : string ;
begin
Result:=-1 ;
if ObjFolio=nil then Exit ;
Ecr:=ObjFolio.GetRow(Row-GS.FixedRows) ; if Ecr=nil then Exit ;
Arret:=Ecr.GetValue('EG_ARRET') ;
if (Arret[1]='X') then Result:=SF_GEN ;
if (Arret[2]='X') then Result:=SF_AUX ;
if (Arret[3]='X') then Result:=SF_REFI ;
if (Arret[4]='X') then Result:=SF_LIB ;
if (Arret[5]='X') then Result:=SF_DEBIT ;
if (Arret[6]='X') then Result:=SF_CREDIT ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 28/07/2003
Modifié le ... : 21/10/2004
Description .. : LG - 28/07/2003 - FB 12575 - on recalcul la ref quoi qu'il se 
Suite ........ : passe
Suite ........ : - FB 13598 - 21/10/2004 - LG - le champs solde se mets a 
Suite ........ : la fos au debit et au credit
Mots clefs ... : 
*****************************************************************}
function TFSaisieFolio.GetGuideValue(Col, Row : LongInt; bMove : Boolean) : Boolean ;
var Ecr : TZF ; bPost : Boolean ; Arret, sSauve, sVal : String ;
begin
Result:=FALSE ;
if ObjFolio=nil then Exit ;
Ecr:=ObjFolio.GetRow(Row-GS.FixedRows) ; if Ecr=nil then Exit ;
GuideCalcRow:=Row ;
GuideCalcCol:=Col ;
Arret:=Ecr.GetValue('EG_ARRET') ; bPost:=FALSE ;
if Col=SF_NAT then bPost:=TRUE ;
if Col=SF_JOUR then bPost:=TRUE ;
if Col=SF_GEN then
   begin
   if Ecr.GetValue('EG_GENERAL')<>#0 then
     GS.Cells[SF_GEN, Row]:=GFormule(Ecr.GetValue('EG_GENERAL'), GetFormule, nil, 1) ;
   if Arret[1]='-' then bPost:=TRUE ;
   end ;
if Col=SF_AUX then
   begin
   if Ecr.GetValue('EG_AUXILIAIRE')<>#0 then
     GS.Cells[SF_AUX, Row]:=GFormule(Ecr.GetValue('EG_AUXILIAIRE'), GetFormule, nil, 1) ;
   if Arret[2]='-' then bPost:=TRUE ;
   end ;
if Col=SF_REFI then
   begin
   if (Ecr.GetValue('EG_REFINTERNE')<>#0) and
      ((GS.Cells[SF_REFI, Row]='') ) then
     begin
     sSauve:=GS.Cells[SF_REFI, Row] ;
     GS.Cells[SF_REFI, Row]:=GFormule(Ecr.GetValue('EG_REFINTERNE'), GetFormule, nil, 1) ;
     if (GS.Cells[SF_REFI, Row]='') then GS.Cells[SF_REFI, Row]:=sSauve ;
    // Ecr.PutValue('INITGUIDE', 'X') ;  FB 12575 on recalcul la ref quoi qu'il se passe
     end ;
   if Arret[3]='-' then bPost:=TRUE ;
   end ;
if Col=SF_LIB then
   begin
   if (Ecr.GetValue('EG_LIBELLE')<>#0) and (GS.Cells[SF_LIB, Row]='') then
     begin
     bFormuleOK:=TRUE ;
     sVal:=GFormule(Ecr.GetValue('EG_LIBELLE'), GetFormule, nil, 1) ;
     if bFormuleOK then GS.Cells[SF_LIB, Row]:=sVal ;
     end ;
   if Arret[4]='-' then bPost:=TRUE ;
   end ;
if Col=SF_DEBIT then
   begin
    if UpperCase(Ecr.GetValue('EG_DEBITDEV'))='[SOLDE]' then
     begin
      sVal:=GFormule('[SOLDE]', GetFormule, nil, 1) ;
      PutRowMontant(SF_DEBIT, Row, Valeur(sVal)) ;
      if Valeur(sVal)=0 then
       begin
        GuideCalcCol:=SF_CREDIT ;
        sVal:=GFormule('[SOLDE]', GetFormule, nil, 1) ;
        PutRowMontant(SF_CREDIT, Row, Valeur(sVal)) ;
       end ; // if
      if Arret[5]='-' then bPost:=TRUE ; 
     end
      else
       if Ecr.GetValue('EG_DEBITDEV')<>#0 then
         begin
         bFormuleOK:=TRUE ;
         sVal:=GFormule(Ecr.GetValue('EG_DEBITDEV'), GetFormule, nil, 1) ;
         {$IFDEF TT}
          AddEvenement('Formule debit : ' + varToStr(Ecr.GetValue('EG_DEBITDEV')) ) ;
          AddEvenement('Valeur au debit : ' + sVal ) ;
         {$ENDIF}
         if bFormuleOK then PutRowMontant(SF_DEBIT, Row, Valeur(sVal)) ;
         if Arret[5]='-' then bPost:=TRUE ;
         end ;
   end ;
if Col=SF_CREDIT then
   begin
    if (trim(VarToStr(Ecr.GetValue('EG_CREDITDEV')))<>'') then
     begin
      bFormuleOK:=TRUE ;
      sVal:=GFormule(Ecr.GetValue('EG_CREDITDEV'), GetFormule, nil, 1) ;
      {$IFDEF TT}
      AddEvenement('Formule credit : ' + varToStr(Ecr.GetValue('EG_CREDITDEV')) ) ;
      AddEvenement('Valeur au credit : ' + sVal ) ;
     {$ENDIF}
      if bFormuleOK then PutRowMontant(SF_CREDIT, Row, Valeur(sVal)) ;
     end ;
    if Arret[6]='-' then bPost:=TRUE ;
   end ;
GuideCalcCol:=-1 ;
GuideCalcRow:=-1 ;
Result:=bPost and bMove;
end ;

procedure TFSaisieFolio.GetGuideCalcRow(Row : LongInt) ;
var i : Integer ;
begin
for i:=SF_FIRST to SF_LAST do
  begin GetGuideValue(i, Row, FALSE) ; SetRowTZF(Row) ; end ;
end ;


function TFSaisieFolio.IsJalBqe : Boolean ;
begin
Result:=FALSE ;
if (memJal.Nature='BQE') or (memJal.Nature='CAI') then Result:=TRUE ;
end ;

function TFSaisieFolio.IsJalAutorise(CodeJal : string) : Boolean ;
begin
Result:= not ( ((not bModeRO) and (VH^.JalAutorises<>'') and (Pos(';'+CodeJal+';',VH^.JalAutorises)<=0)) ) ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 21/03/2005
Modifié le ... :   /  /    
Description .. : - LG - 21/03/2005 - limitation des bornes pour la recherche 
Suite ........ : des compte de contreparties
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.GetCPartie(Row : LongInt; Piece : string; var NumCompte, NumAux : string) ;
var i : Integer ; bJalBqe : Boolean ; RowRef : LongInt ;
 lInMax,lInMin : integer ;
begin
if Statut=taconsult then exit ;
{$IFDEF TT}
AddEvenement('$$$$$$$         GetCPartie :'+IntToStr(Row));
{$ENDIF}
lInMax := Row + MAX_BOUCLE ;
if lInMax > nbLignes then lInMax := nbLignes ;
lInMin := Row - MAX_BOUCLE ;
if lInMin < 1 then lInMin := 1 ;
NumCompte:='' ; NumAux:='' ; bJalBqe:=IsJalBqe ;
HGBeginUpdate(GS) ;
try
if bJalBqe then
  begin
  {Pièce de banque}
  if GS.Cells[SF_GEN, Row]=memJal.CContreP then
    begin
    {Sur cpte de banque, contrep=première ligne non banque au dessus}
    for i:=Row-1 downto lInMin do
      begin
      if GS.Cells[SF_PIECE, i]<>Piece then Break ;
      if GS.Cells[SF_GEN, i]<>memJal.CContreP then begin NumCompte:=GS.Cells[SF_GEN, i] ; Break ; end ;
      end ;
    end else
    begin
    {Sur cpte non banque, contrep=première ligne banque au dessous}
    for i:=Row+1 to lInMax do
      begin
      if GS.Cells[SF_PIECE, i]<>Piece then Break ;
      if GS.Cells[SF_GEN, i]=memJal.CContreP then
        begin
        NumCompte:=GS.Cells[SF_GEN, i] ; NumAux:=GS.Cells[SF_AUX, i] ;
        Break ;
        end ;
      end ;
    end ;
  end else
  begin
  {Cas normal}
  if IsRowTiers(Row) then
    begin
    {Lecture avant pour trouver "HT"}
    for i:=Row+1 to lInMax do
      begin
      if GS.Cells[SF_PIECE, i]<>Piece then Break ;
      if IsRowHT(i) then
        begin
        NumCompte:=GS.Cells[SF_GEN, i] ;
        Break ;
        end ;
      end ;
    if NumCompte='' then
      begin
      {Lecture arrière pour trouver "HT"}
      for i:=Row-1 downto lInMin do
        begin
        if GS.Cells[SF_PIECE, i]<>Piece then Break ;
        if isRowHT(i) then
          begin
          NumCompte:=GS.Cells[SF_GEN, i] ;
          Break ;
          end;
        end ;
      end ;
    end else
    begin
    {Lecture arrière pour trouver "Tiers"}
    for i:=Row-1 downto lInMin do
      begin
      if GS.Cells[SF_PIECE, i]<>Piece then Break ;
      if IsRowTiers(i) then
        begin
        NumCompte:=GS.Cells[SF_GEN, i] ; NumAux:=GS.Cells[SF_AUX, i] ;
        Break ;
        end ;
      end ;
    if NumCompte='' then
      begin
      {Lecture avant pour trouver "Tiers"}
      for i:=Row+1 to lInMax do
        begin
        if GS.Cells[SF_PIECE, i]<>Piece then Break ;
        if IsRowTiers(i) then
          begin
          NumCompte:=GS.Cells[SF_GEN, i] ; NumAux:=GS.Cells[SF_AUX, i] ;
          Break ;
          end ;
        end ;
      end ;
    end ;
  end ;
{Cas particuliers}
if NumCompte<>'' then Exit ;
if ((bJalBqe) and (GS.Cells[SF_GEN, Row]<>memJal.CContreP)) then NumCompte:=memJal.CContreP else
 if not IsRowBqe(Row) then
   begin
   for i:=Row+1 to lInMax do
     begin
     if GS.Cells[SF_PIECE, i]<>Piece then Break ;
     if IsRowBqe(i) then begin NumCompte:=GS.Cells[SF_GEN, i] ; Break ; end ;
     end ;
   if NumCompte='' then
     for i:=Row-1 downto lInMin do
       begin
       if GS.Cells[SF_PIECE, i]<>Piece then Break ;
       if IsRowBqe(i) then begin NumCompte:=GS.Cells[SF_GEN, i] ; Break ; end ;
       end ;
   end else
   begin
   for i:=Row-1 downto lInMin do
     begin
     if GS.Cells[SF_PIECE, i]<>Piece then Break ;
     if IsRowTiers(i) then begin NumCompte:=GS.Cells[SF_GEN, i] ; NumAux:=GS.Cells[SF_AUX, i] ; Break ; end ;
     end ;
   if NumCompte='' then
     for i:=Row+1 to lInMax do
       begin
       if GS.Cells[SF_PIECE, i]<>Piece then Break ;
       if IsRowTiers(i) then begin NumCompte:=GS.Cells[SF_GEN, i] ; NumAux:=GS.Cells[SF_AUX, i] ; Break ; end ;
       end ;
   end ;
{Si rien trouvé, swaper les lignes 1 et 2}
if NumCompte<>'' then Exit ;
RowRef:=-1 ;
if (Row>1) and (GS.Cells[SF_PIECE, Row-1]=Piece) then RowRef:=Row-1 else
  if (Row+1<=nbLignes) and (GS.Cells[SF_PIECE, Row+1]=Piece) then RowRef:=Row+1 ;
if RowRef<>-1 then
  begin
  NumCompte:=GS.Cells[SF_GEN, RowRef] ; NumAux:=GS.Cells[SF_AUX, RowRef] ;
  end ;
finally
 HGEndUpdate(GS) ;
end;
end ;

procedure TFSaisieFolio.PutRegime(Row : LongInt) ;
var Ecr : TZF ; NumCompte : string ; k : Integer ;
begin
if ObjFolio=nil then Exit ;
Ecr:=ObjFolio.GetRow(Row-GS.FixedRows) ; if Ecr=nil then Exit ;
NumCompte:=GS.Cells[SF_AUX, Row] ;
if NumCompte='' then
  begin
  NumCompte:=GS.Cells[SF_GEN, Row] ; if NumCompte='' then Exit ;
  k:=Comptes.GetCompte(NumCompte) ; if k<0 then Exit ;
  if not Comptes.IsTiers(k) then Exit ;
  Ecr.PutValue('E_REGIMETVA', Comptes.GetValue('G_REGIMETVA', k)) ;
  end else
  begin
  k:=Tiers.GetCompte(NumCompte) ; if k<0 then Exit ;
  Ecr.PutValue('E_REGIMETVA', Tiers.GetValue('T_REGIMETVA', k)) ;
  end ;
if Ecr.GetValue('E_REGIMETVA')<>'' then
  PutRegimePiece(Row, Ecr.GetValue('E_REGIMETVA'), GS.Cells[SF_PIECE, Row]) ;
end ;

procedure TFSaisieFolio.PutRegimePiece(Row : LongInt; Regime, Piece : string) ;
var Ecr : TZF ; i : Integer ;
begin
HGBeginUpdate(GS) ;
if objfolio = nil then exit ;
try
for i:=Row-1 downto 1 do
  begin
  Ecr:=ObjFolio.GetRow(i-GS.FixedRows) ; if Ecr=nil then Break ;
  if GS.Cells[SF_PIECE, i]<>Piece then Break ;
  Ecr.PutValue('E_REGIMETVA', Regime) ;
  end ;
for i:=Row+1 to nbLignes do
  begin
  Ecr:=ObjFolio.GetRow(i-GS.FixedRows) ; if Ecr=nil then Break ;
  if GS.Cells[SF_PIECE, i]<>Piece then Break ;
  Ecr.PutValue('E_REGIMETVA', Regime) ;
  end ;
finally
 HGEndUpdate(GS) ;
end;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 05/09/2002
Modifié le ... : 16/08/2004
Description .. : 05/09/2002 - on ne tient plus compte de la colonne
Suite ........ : - 16/08/2004 - LG - FB 13056 - l'affectation du rib ne se
Suite ........ : faisait pas qd on cheangant de rib
Mots clefs ... : 
*****************************************************************}
function TFSaisieFolio.PutRib(Col, Row : LongInt; bAux : Boolean ; vBoForce : boolean = false) : Boolean ;
var Ecr : TZF ; sRib, sRibP, NumCompte : string ;
begin
Result:=FALSE ; if bModeRO then Exit ;
if objfolio = nil then exit ;
if bAux then NumCompte:=GS.Cells[SF_AUX, Row] else NumCompte:=GS.Cells[SF_GEN, Row] ;
if NumCompte='' then Exit ;
Ecr:=ObjFolio.GetRow(Row-GS.FixedRows) ;
if Ecr=nil then Exit ;
sRib:=Ecr.GetValue('E_RIB') ;
if not VH^.AttribRIBAuto then begin Result:=TRUE ; Exit ; end ;
if (sRib<>'') and (not vBoForce) then begin Result:=TRUE ; Exit ; end ;
sRibP:=GetRIBPrincipal(NumCompte) ; Ecr.PutValue('E_RIB', sRibP) ;
Result:=TRUE ;
end ;

procedure TFSaisieFolio.PutConso(Row : LongInt) ;
var Ecr : TZF ; NumCompte, Conso : string ; k : Integer ;
begin
if bModeRO then Exit ; if objfolio = nil then exit ;
Ecr:=ObjFolio.GetRow(Row-GS.FixedRows) ; if Ecr=nil then Exit ;
Conso:='' ; Ecr.PutValue('E_CONSO', Conso) ;
NumCompte:=GS.Cells[SF_AUX, Row] ; k:=Tiers.GetCompte(NumCompte) ;
if k<0 then
  begin
  NumCompte:=GS.Cells[SF_GEN, Row] ; k:=Comptes.GetCompte(NumCompte) ;
  if k>=0 then Conso:=Comptes.GetValue('G_CONSO', k) ;
  end else Conso:=Tiers.GetValue('T_CONSO', k) ;
Ecr.PutValue('E_CONSO', Conso) ;
end ;

//=======================================================
//================= Fonctions de calcul =================
//=======================================================
{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 25/06/2003
Modifié le ... : 20/09/2006
Description .. : - LG - 25/06/2003 - changement du mode de calcul du 
Suite ........ : solde. On utilise plus le solde des annees precedentes 
Suite ........ : poour calculer celui ci
Suite ........ : - 13/08/2003 - FB 12127 - le total folio etait incorrect ( 
Suite ........ : debrayer la modif plus ho )
Suite ........ : - 20/09/2006 - FB 18626 - calcul du solde du journal apres 
Suite ........ : un lettrage
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.CalculSoldeJournal ;
var i : Integer ; DeltaSavD, DeltaSavC : Double ;
begin
{$IFDEF TT}
  AddEvenement('CalculSoldeJournal') ;
{$ENDIF}
memJal.TotFolDebit:=0; memJal.TotFolCredit:=0 ; HGBeginUpdate(GS) ;
try
for i:=GS.FixedRows to nbLignes do
    begin
    memJal.TotFolDebit:=memJal.TotFolDebit+Valeur(GS.Cells[SF_DEBIT,i]) ;
    memJal.TotFolCredit:=memJal.TotFolCredit+Valeur(GS.Cells[SF_CREDIT,i]) ;
    end ;
DeltaSavD:=0 ; DeltaSavC:=0 ;
if memJal.TotSavDebit<>0  then DeltaSavD:=memJal.TotDebDebit-memJal.TotSavDebit ;
if memJal.TotSavCredit<>0 then DeltaSavC:=memJal.TotDebCredit-memJal.TotSavCredit ;
if InitParams.ParRecupLettrage or FBoLettrageEnCours then
begin
 memJal.TotSaiDebit:=memJal.TotFolDebit ;
 memJal.TotSaiCredit:=memJal.TotFolCredit ;
end
 else
  begin
   memJal.TotSaiDebit:=memJal.TotPerDebit+DeltaSavD-memJal.TotDebDebit+memJal.TotFolDebit-memJal.TotVirDebit ;
   memJal.TotSaiCredit:=memJal.TotPerCredit+DeltaSavC-memJal.TotDebCredit+memJal.TotFolCredit-memJal.TotVirCredit ;
  end;

SA_FOLIODEBIT.Value:=memJal.TotFolDebit ;
SA_FOLIOCREDIT.Value:=memJal.TotFolCredit ;

SA_TOTALDEBIT.Value:=memJal.TotSaiDebit ;
SA_TOTALCREDIT.Value:=memJal.TotSaiCredit ;

AfficheLeSolde(SA_SOLDE, SA_TOTALDEBIT.Value, SA_TOTALCREDIT.Value) ;
finally
 HGEndUpdate(GS) ;
end ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 13/09/2004
Modifié le ... : 03/11/2006
Description .. : - LG - 13/09/2004 - FB 10810 - on ne tient plus compte du
Suite ........ : compte courant pour calculer le solde
Suite ........ : - LG - 03/11/2006 - FB 16744 - coorection de l'affichage du
Suite ........ : solde apres l'ouverture d'un lettrage ou d'une consultation (
Suite ........ : on recup les ecritures de type L )
Mots clefs ... :
*****************************************************************}
procedure TFSaisieFolio.SoldeDuCompte( D, C : double ) ;
var
 lQ : TQuery ;
begin

 {$IFDEF TT}
  AddEvenement('TFSaisieFolio.SoldeDuCompte') ;
 {$ENDIF}

  if not FBoFaireCumul then exit ;

  if (FRechCompte.Info.Compte.GetValue('_SOLDED') = - 1) and (FRechCompte.Info.Compte.GetValue('_SOLDEC') = - 1) then
   begin
       lQ:=OpenSQL('SELECT SUM(E_DEBIT), SUM(E_CREDIT) FROM ECRITURE'
              +' WHERE E_GENERAL="'+GS.Cells[SF_GEN, GS.Row]+'"'
              +' AND E_EXERCICE="'+memOptions.Exo+'"'
              +' AND E_ETABLISSEMENT="'+E_ETABLISSEMENT.Value+'" '
              +' AND E_DATECOMPTABLE>="'+USDateTime(ctxExercice.CPExoRef.Deb)+'"'
              +' AND E_DATECOMPTABLE<="'+USDateTime(EncodeDate(memOptions.Year, memOptions.Month, memOptions.MaxJour))+'"'
              +' AND ( E_QUALIFPIECE="N" OR E_QUALIFPIECE="C" OR E_QUALIFPIECE="L" ) ',
              TRUE) ;
    FRechCompte.Info.Compte.PutValue('_SOLDED', lQ.Fields[0].AsFloat ) ;
    FRechCompte.Info.Compte.PutValue('_SOLDEC', lQ.Fields[1].AsFloat ) ;
    Ferme(lQ) ;
   end ;
 AfficheLeSolde(SA_SoldePerGen, FRechCompte.Info.Compte.GetValue('_SOLDED')+D, FRechCompte.Info.Compte.GetValue('_SOLDEC')+C) ;
 SoldeD.caption := '(M) ' + SA_SoldePerGen.Text ;
end ;


procedure TFSaisieFolio.CalculSoldeCompte(NumCompte : string ; Col : Integer;
          vTotDebit, vTotCredit : double) ;
var Ecr: TZF ; i : Integer ; TotCDebit, TotCCredit : Double ; D,C : double ;
begin
TotCDebit:=vTotDebit ; TotCCredit:=vTotCredit ; HGBeginUpdate(GS) ; if objfolio = nil then exit ;
D := 0 ; C := 0 ;
try

 for i:=GS.FixedRows to nbLignes do
  begin

   if GS.Cells[SF_NEW, i]='N' then
    begin

     Ecr:=ObjFolio.GetRow(i-GS.FixedRows) ; if Ecr=nil then Continue ;

     if (Col=SF_GEN) and (Ecr.GetValue('OLDGENERAL')=NumCompte) then
      begin
       TotCDebit  := TotCDebit-Ecr.GetValue('OLDDEBIT') ;
       TotCCredit := TotCCredit-Ecr.GetValue('OLDCREDIT') ;
      end ;

     if (Col=SF_AUX) and (Ecr.GetValue('OLDAUXILIAIRE')=NumCompte) then
      begin
       TotCDebit  := TotCDebit-Ecr.GetValue('OLDDEBIT') ;
       TotCCredit := TotCCredit-Ecr.GetValue('OLDCREDIT') ;
      end ;
      
    end ; // if

   if (GS.Cells[SF_NEW, i]='O') and (GS.Cells[Col, i]=NumCompte) then
    begin
     D := D + Valeur(GS.Cells[SF_DEBIT, i]) ;
     C := C + Valeur(GS.Cells[SF_CREDIT, i]) ;
    end ;

   if (GS.Cells[Col, i]=NumCompte) then
    begin
     TotCDebit:=TotCDebit+Valeur(GS.Cells[SF_DEBIT, i]) ;
     TotCCredit:=TotCCredit+Valeur(GS.Cells[SF_CREDIT, i]) ;
    end ;

  end ; // for
 
 if Col=SF_GEN then
  begin
   SoldeDuCompte(D,C) ; AfficheLeSolde(SA_SOLDEG, TotCDebit, TotCCredit) ;
  end ;

 if Col=SF_AUX then
  begin
   SoldeAux(D,C) ; AfficheLeSolde(SA_SOLDET, TotCDebit, TotCCredit) ;
  end ;
  
 finally
  HGEndUpdate(GS) ;
 end;
 
end ;


function TFSaisieFolio.IsSoldeJournal : Boolean ;
begin
Result:=Arrondi(memJal.TotSaiDebit-memJal.TotSaiCredit, memPivot.Decimale)=0 ;
end ;

function TFSaisieFolio.IsSoldeFolio : Boolean ;
begin
Result:=Arrondi(memJal.TotFolDebit-memJal.TotFolCredit, memPivot.Decimale)=0 ;
end ;

function TFSaisieFolio.IsSoldePiece(Piece : LongInt) : Boolean ;
begin
Result:=(SoldePiece(Piece)=0) ;
end ;

function TFSaisieFolio.SoldePiece(Piece : LongInt) : Double ;
var i, CurPiece : LongInt ; SoldeP, SoldeD, SoldeC, ValD, ValC : double ;
begin
SoldeP:=0 ; SoldeD:=0 ; SoldeC:=0 ; HGBeginUpdate(GS) ;
try
for i:=GS.FixedRows to nbLignes do
    begin
    CurPiece:=StrToInt(GS.Cells[SF_PIECE,i]) ; if CurPiece<>Piece then Continue ;
    ValD:=Valeur(GS.Cells[SF_DEBIT,i]) ; ValC:=Valeur(GS.Cells[SF_CREDIT,i]) ;
    SoldeD:=SoldeD+ValD ;                SoldeC:=SoldeC+ValC ;
    SoldeP:=Arrondi(SoldeD-SoldeC, memPivot.Decimale) ;
    end ;
Result:=SoldeP ;
finally
 HGEndUpdate(GS) ;
end;
end ;


function TFSaisieFolio.SoldePieceDevise(Piece : LongInt; var CumulD, CumulC : Double) : Double ;
var Ecr : TZF ; i, CurPiece : LongInt ; SoldeP, ValD, ValC : double ;
begin
SoldeP:=0 ; CumulD:=0 ; CumulC:=0 ;
for i:=GS.FixedRows to nbLignes do
    begin
    CurPiece:=StrToInt(GS.Cells[SF_PIECE,i]) ; if CurPiece<>Piece then Continue ;
    if ObjFolio=nil then Continue ;
    Ecr:=ObjFolio.GetRow(i-GS.FixedRows) ; if Ecr=nil then Continue ;
    ValD:=Ecr.GetValue('E_DEBITDEV') ;  ValC:=Ecr.GetValue('E_CREDITDEV') ;
    CumulD:=CumulD+ValD ; CumulC:=CumulC+ValC ;
    SoldeP:=Arrondi(CumulD-CumulC, memPivot.Decimale) ;
    end ;
Result:=SoldeP ;
end ;


procedure TFSaisieFolio.ChangeTauxPiece(Piece : LongInt; LDevise: RDevise) ;
var Ecr : TZF ; i, CurPiece : LongInt ;
begin
for i:=GS.FixedRows to nbLignes do
  begin
  CurPiece:=StrToInt(GS.Cells[SF_PIECE,i]) ; if CurPiece<>Piece then Continue ;
  if ObjFolio=nil then Continue ;
  SetRowTZF(i) ;
  Ecr:=ObjFolio.GetRow(i-GS.FixedRows) ; if Ecr=nil then Continue ;
  Ecr.SetMontants(Ecr.GetValue('E_DEBITDEV'), Ecr.GetValue('E_CREDITDEV'), LDevise, TRUE) ;
  GetRowTZF(i-GS.FixedRows) ;
  end ;
end ;


function TFSaisieFolio.GetAnySolde(Piece : LongInt; var SP, SD : Double) : LongInt ;
var Ecr : TZF ; i, CurPiece : LongInt ; VPD, VPC, VDD, VDC : Double ;
begin
SP:=0 ; SD:=0 ; VPD:=0 ; VPC:=0 ; VDD:=0 ; VDC:=0 ;
Result:=-1 ;
if ObjFolio=nil then Exit ;
for i:=GS.FixedRows to nbLignes do
    begin
    if not memParams.bLibre then
      begin
      CurPiece:=StrToInt(GS.Cells[SF_PIECE,i]) ;
      if CurPiece<>Piece then Continue ;
      end ;
    if ObjFolio=nil then Continue ;
    Ecr:=ObjFolio.GetRow(i-GS.FixedRows) ; if Ecr=nil then Continue ;
    VPD:=VPD+Ecr.GetValue('E_DEBIT') ;     VPC:=VPC+Ecr.GetValue('E_CREDIT') ;
    VDD:=VDD+Ecr.GetValue('E_DEBITDEV') ;  VDC:=VDC+Ecr.GetValue('E_CREDITDEV') ;
    SP:=VPD-VPC ;
    SD:=VDD-VDC ;
    end ;
end ;

// Solde théorique du compte de contrepartie

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 28/09/2004
Modifié le ... :   /  /    
Description .. : - LG - 28/09/2004 - FB 10162 - reecriture de la fct, ne fct 
Suite ........ : pas qd on supprimais le compte de contrepartie du folio
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.SoldeProgressif(Row : LongInt ) ;
var i : LongInt ; SoldeDPer, SoldeCPer,SoldeDTot, SoldeCTot : double ;
D,C : double ;
begin
if not VoirSoldeTreso then Exit ;
i:=Comptes.GetCompte(memJal.CContreP) ; if i<0 then Exit ;
Comptes.Solde(SoldeDPer,SoldeCPer,memOptions.TypeExo) ;
D:=0 ; C:=0 ; 
for i:=GS.FixedRows to nbLignes do
 if (GS.Cells[SF_GEN, i]<>memJal.CContreP) then
  begin
   C:=C+Valeur(GS.Cells[SF_DEBIT,  i]) ;
   D:=D+Valeur(GS.Cells[SF_CREDIT, i]) ;
  end ;
SoldeDTot := SoldeDPer - ( FRdSoldeDBq - D ) ;
SoldeCTot := SoldeCPer - ( FRdSoldeCBq - C ) ;
//SoldeProg:=Arrondi(SoldeTresor+(SoldeD-SoldeC), memPivot.Decimale) ;
AfficheLeSolde(SA_SOLDEPROG, SoldeDTot, SoldeCTot) ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 03/10/2002
Modifié le ... :   /  /
Description .. : -03/10/2002 - on force la suppression des lignes d'ecart de
Suite ........ : conversion ( on ne peut pas le faire dans la saisie )
Mots clefs ... :
*****************************************************************}
function TFSaisieFolio.VerifEquilibre(Piece, DefRow : LongInt) : Boolean ;
var EcartP, EcartD : Double ;
    RowEcart : LongInt ;
begin
{$IFDEF TT}
AddEvenement('VerifEquilibre DefRow:='+intToStr(DefRow) + ' : vérification des ecart de conversion' ) ;
{$ENDIF}
Result:=FALSE ;
RowEcart:=GetAnySolde(Piece, EcartP, EcartD) ;
EcartP:=Arrondi(EcartP, V_PGI.OkDecV) ;
if RowEcart>=0 then DeleteRow(RowEcart,true) ; // on force la suppression des lignes d'ecart de conversion ( on ne peut pas le faire dans la saisie )
//if ((EcartP=0) and (EcartO=0)) then Exit ;
if (GS.Cells[SF_DEVISE, DefRow]<>'') and (GS.Cells[SF_DEVISE, DefRow]<>V_PGI.DevisePivot) then
  begin Result:=SoldeLaLigne(Piece, DefRow, EcartP, EcartD) ; Exit ; end ;
end;

function TFSaisieFolio.SoldeLaLigne(Piece, DefRow : LongInt; var SP,SD : Double) : Boolean ;
var Ecr : TZF ; X : Double ;
begin
Result:=TRUE ;
if ObjFolio=nil then Exit ;
Ecr:=ObjFolio.GetRow(DefRow-GS.FixedRows) ; if Ecr=nil then Exit ;
// Pivot
X:=-1*Valeur(GS.Cells[SF_DEBIT, DefRow]) ;
if X=0 then X:=Valeur(GS.Cells[SF_CREDIT, DefRow]) ; SP:=SP+X ;
GS.Cells[SF_DEBIT, DefRow]:='' ; GS.Cells[SF_CREDIT, DefRow]:='' ;
if SP>0 then
  begin
  GS.Cells[SF_CREDIT, DefRow]:=StrFMontant(SP,15,memPivot.Decimale,'',TRUE) ;
  Ecr.PutValue('E_CREDIT', Valeur(GS.Cells[SF_CREDIT, DefRow])) ;
  end else
  begin
  GS.Cells[SF_DEBIT,  DefRow]:=StrFMontant(-SP,15,memPivot.Decimale,'',TRUE) ;
  Ecr.PutValue('E_DEBIT', Valeur(GS.Cells[SF_DEBIT, DefRow])) ;
  end ;
// Mode Euro ?
if (GS.Cells[SF_DEVISE, DefRow]<>'') then
   begin
   if memDevise.Code<>GS.Cells[SF_DEVISE, DefRow] then
     begin
     memDevise.Code:=GS.Cells[SF_DEVISE, DefRow] ; GetInfosDevise(memDevise) ;
     memDevise.Taux:=GetTaux(memDevise.Code, memDevise.DateTaux, Ecr.GetValue('E_DATECOMPTABLE')) ;
     end ;
   Ecr.SetMontants(Ecr.GetValue('E_DEBITDEV'), Ecr.GetValue('E_CREDITDEV'),
                   memDevise,FALSE) ;
   end else
     Ecr.SetMontants(Valeur(GS.Cells[SF_DEBIT, DefRow]), Valeur(GS.Cells[SF_CREDIT, DefRow]),
                     FInfo.Devise.Dev, FALSE) ;
// Devise
if SD<>0 then
   begin
   X:=-1*Ecr.GetValue('E_DEBITDEV') ;
   if X=0 then X:=Ecr.GetValue('E_CREDITDEV') ; SD:=SD+X ;
   if SD>0 then Ecr.PutValue('E_CREDITDEV', SD)
           else Ecr.PutValue('E_DEBITDEV', -SD) ;
   end ;
CalculSoldeJournal ;
PutRowProrata(DefRow) ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 13/06/2002
Modifié le ... : 02/08/2004
Description .. : - 13/06/2002 - suppression de warning
Suite ........ : - 02/08/2004 - on force le controle de la nature sur la ligne 
Suite ........ : courante lors de l'enregistrement
Mots clefs ... : 
*****************************************************************}
function TFSaisieFolio.VerifNature(Piece, DefRow : LongInt ; vBoForce : boolean = false) : LongInt ;
var i, k : LongInt ; NumCompte : string ; lStNat : string;
begin
result:=0 ;
for i:=GS.FixedRows to nbLignes do
  begin
  if (GS.Cells[SF_PIECE, i]<>IntToStr(Piece)) or ((i=DefRow) and not vBoForce ) then Continue ;
  NumCompte:=GS.Cells[SF_GEN, i] ;
  if NumCompte<>'' then
    begin
    k:=Comptes.GetCompte(NumCompte) ;
    if k<0 then Break ; lStNat:=GetCellNature(DefRow) ;
    if not NaturePieceCompteOK(lStNat,
                               Comptes.GetValue('G_NATUREGENE', k)) then
      begin
      Result:=i ; Break ;
      end ;
    end ;
  end ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 02/08/2004
Modifié le ... :   /  /
Description .. : - LG - 02/08/2004 - FB 13780 - ajout d'un controle supp, le
Suite ........ : compte de contrepartie doit etre mouvemente
Mots clefs ... : 
*****************************************************************}
function TFSaisieFolio.VerifTreso(Piece, DefRow : LongInt) : Boolean ;
var i, CurPiece : LongInt ;
begin
Result:=FALSE ;
if (not memParams.bSoldeProg) then begin Result:=TRUE ; Exit ; end ;
if (memParams.bLibre) then
  begin
  for i:=GS.FixedRows to nbLignes do
    begin
    if (GS.Cells[SF_GEN, i]=memJal.CContreP) and ((GS.Cells[SF_DEBIT, i]<>'') or (GS.Cells[SF_CREDIT, i]<>'')) then
     begin
      Result:=TRUE ;
      Break ;
     end ;
    end ;
  end
  else
  begin
  for i:=GS.FixedRows to nbLignes do
    begin
    CurPiece:=StrToInt(GS.Cells[SF_PIECE, i]) ;
    if CurPiece<>Piece then Continue ;
    if GS.Cells[SF_GEN, i]=memJal.CContreP then
     begin
      Result:=TRUE ;
      Break ;
     end ;
    end ;
  end ;
end ;


function TFSaisieFolio.GetRowFromPiece(Row, Piece : LongInt) : LongInt ;
var i : LongInt ;
begin
Result:=0 ;
for i:=GS.FixedRows to nbLignes do
    begin
    if Piece=StrToInt(GS.Cells[SF_PIECE,i]) then Result:=Result+1 ;
    if Row=i then break ;
    end ;
end ;


procedure TFSaisieFolio.SetPiece ;
var i, nPiece : LongInt ; SoldeD, SoldeC,  ValD, ValC : double ;
begin
HGBeginUpdate(GS) ;
try
SoldeD:=0 ; SoldeC:=0 ; nPiece:=1 ;
for i:=GS.FixedRows to nbLignes do
    begin
    if memParams.bLibre then begin GS.Cells[SF_PIECE, i]:=IntToStr(nPiece) ; Continue ; end ;
    GS.Cells[SF_PIECE, i]:=IntToStr(nPiece) ;
    ValD:=Valeur(GS.Cells[SF_DEBIT,i]) ; ValC:=Valeur(GS.Cells[SF_CREDIT,i]) ;
    SoldeD:=SoldeD+ValD ; SoldeC:=SoldeC+ValC ;
    SoldeProg:=Arrondi(SoldeD-SoldeC, memPivot.Decimale) ;
    if (SoldeProg=0) and ((ValD<>0) or (ValC<>0)) then nPiece:=nPiece+1 ;
    end ;
finally
 HGEndUpdate(GS) ;
end;
end ;

function TFSaisieFolio.GetRowFormule(var Formule : hstring ) : Integer ;
var Pos1, Pos2 : Integer ;
begin
Result:=0 ;
Pos2:=Pos(':L', Formule) ; if Pos2>0 then System.Delete(Formule, Pos2+1, 1) ;
Pos1:=Pos(':',  Formule) ; if (Pos1<=0) or (Pos1=Length(Formule)) then Exit ;
Result:=Round(Valeur(Copy(Formule, Pos1+1, 5))) ;
System.Delete(Formule, Pos1, 5) ;
end ;

function TFSaisieFolio.GetRowAbs(CurRow, ForRow: LongInt) : LongInt ;
var First: LongInt ;
begin
if GS.Cells[SF_PIECE, CurRow]='' then begin Result:=CurRow ; Exit ; end ;
First:=GetRowFirst(StrToInt(GS.Cells[SF_PIECE, CurRow])) ;
if First<0 then begin Result:=CurRow ; Exit ; end ;
Result:=First+(ForRow-1) ;
end ;

function TFSaisieFolio.GetRowRel(CurRow, ForRow: LongInt) : LongInt ;
var First: LongInt ;
begin
if GS.Cells[SF_PIECE, CurRow]='' then begin Result:=CurRow ; Exit ; end ;
First:=GetRowFirst(StrToInt(GS.Cells[SF_PIECE, CurRow])) ;
if First<0 then begin Result:=CurRow ; Exit ; end ;
Result:=First+((CurRow-1)+ForRow) ;
end ;

// Attention: GuideCalcRow est une variable globale de la classe
{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 03/04/2002
Modifié le ... : 28/10/2002
Description .. : LG* correction pour l'appel de la ligne precedente sur la
Suite ........ : premiere ligne ( dans les libelles automaitiques )
Suite ........ :
Suite ........ : - 28/10/2002 - gestion du champ e_tva des guides
Mots clefs ... :
*****************************************************************}
function TFSaisieFolio.GetFormule(Formule : hstring) : Variant ;
var Ecr, RefEcr : TZF ; RefRow, CurRow, ForRow, CalcRow, CalcCol: LongInt ;
    k : Integer ; NumCompte : string ;
    lIndex : integer;
    lRdTauxTva : double ;
    lRdCompteTva : string ;
    lStNatureAuxi : string ;
    lStRegimeTVA : string ;
begin
Result:=#0 ;
if ObjFolio=nil then Exit ;
if GuideCalcRow<0 then CalcRow:=GS.Row else CalcRow:=GuideCalcRow ;
if GuideCalcCol<0 then CalcCol:=GS.Col else CalcCol:=GuideCalcCol ;
Ecr:=ObjFolio.GetRow(CalcRow-GS.FixedRows) ; if Ecr=nil then Exit ;
Formule:=AnsiUpperCase(Trim(Formule)) ; if Formule='' then Exit ;
ForRow:=GetRowFormule(Formule) ;
if bGuideRun then
  begin
  CurRow:=Ecr.GetValue('EG_NUMLIGNE') ;
  if ForRow>0 then RefRow:=GuideFirstRow+(ForRow-1)
    else if ForRow<0 then RefRow:=GuideFirstRow+((CurRow-1)+ForRow)
      else RefRow:=CalcRow ;
  end else
  begin
  CurRow:=CalcRow ;
  if ForRow>0 then RefRow:=GetRowAbs(CurRow, ForRow)
    else if ForRow<0 then RefRow:=GetRowRel(CurRow, ForRow)
      else RefRow:=CurRow ;
  end ;
//LG* test de la ligne avant de recherche de tzf
lIndex:=RefRow-GS.FixedRows ; if lIndex < 0 then lIndex:=0 ;
RefEcr:=ObjFolio.GetRow(lIndex) ; if RefEcr=nil then Exit ;
if Pos('E_TVA', Formule)>0 then
   begin
    if Tiers.Load([GS.Cells[SF_AUX, lIndex+1]]) <> -1 then
     begin
      lStNatureAuxi := Tiers.getValue('T_NATUREAUXI') ;
      lStRegimeTVA  := Tiers.getValue('T_REGIMETVA') ;
    end ; // if
   Comptes.RecupInfoTVA( GS.Cells[SF_GEN, lIndex+1] ,
                         Valeur(GS.Cells[SF_DEBIT, lIndex+1] ) ,
                         GS.Cells[SF_NAT, lIndex+1] ,
                         memJal.Nature ,
                         lStNatureAuxi ,
                         lRdCompteTva,
                         lRdTauxTva ,
                         lStRegimeTVA  );
   Result:=lRdTauxTva;
   Exit ;
   end ;
if Pos('TVA', Formule)>0 then
   begin
   // Comptes.RecupInfoTVA(RefEcr.GetValue('E_GENERAL') , RefEcr.GetValue('E_DEBIT') ,lRdCompteTva, lRdTauxTva );
    Result:=lRdCompteTva;
    Exit ;
   end ;  
if Pos('E_', Formule)>0 then
   begin
   Result:=RefEcr.GetValue(Formule) ;
   Exit ;
   end ;
if Pos('G_', Formule)>0 then
   begin
   NumCompte:=RefEcr.GetValue('E_GENERAL') ;
   k:=Comptes.GetCompte(NumCompte) ;
   if k<0 then begin bFormuleOK:=FALSE ; Exit ; end ;
   // Attention si Comptes optimisé faire en SELECT Formule FROM GENERAUX
   Result:=Comptes.GetValue(Formule, k) ;
   Exit ;
   end ;
if (Pos('T_', Formule)>0)  then
   begin
   if (RefEcr.GetValue('E_AUXILIAIRE')<>'') then
     begin
     NumCompte:=RefEcr.GetValue('E_AUXILIAIRE') ;
     k:=Tiers.GetCompte(NumCompte) ;
     if k<0 then begin bFormuleOK:=FALSE ; Exit ; end ;
     Result:=Tiers.GetValue(Formule, k) ;
     bFormuleOK:=TRUE ;
     end else bFormuleOK:=FALSE ;
   Exit ;
   end ;
if (Pos('SOLDE', Formule)>0) then
   begin
   if (GS.Cells[SF_DEVISE, CalcRow]<>'') //or (GS.Cells[SF_EURO, CalcRow]='X')
     then begin bGuideDoSolde:=TRUE ; Exit ; end ;
   SoldeClick(CalcRow) ;
   Result:=GS.Cells[CalcCol, CalcRow] ;
   Exit ;
   end ;
end ;

function TFSaisieFolio.GetFormulePourCalc(Formule : hstring) : Variant ;
var
 lCalcRow : integer ;
 lEcr : TZF ;
begin
Result:=#0 ; 
Formule:=AnsiUpperCase(Trim(Formule)) ; if Formule='' then Exit ;
if ObjFolio=nil then Exit ;
if pos('E_DEBIT',Formule)>0 then lCalcRow:=GS.Row else lCalcRow:=GS.Row-1 ;
lEcr:=ObjFolio.GetRow(lCalcRow-GS.FixedRows) ; if lEcr=nil then Exit ;
if Pos('E_', Formule)>0 then
 Result:=lEcr.GetValue(Formule) ;
end ;


//=======================================================
//=================== Gestion du Grid ===================
//=======================================================

procedure TFSaisieFolio.SetGridSep(ACol, ARow : Integer ; Canvas : TCanvas; bHaut : Boolean) ;
var R : TRect ;
begin
Canvas.Brush.Color := clRed ;
Canvas.Brush.Style := bsSolid ;
Canvas.Pen.Color   := clRed ;
Canvas.Pen.Mode    := pmCopy ;
Canvas.Pen.Style   := psSolid ;
Canvas.Pen.Width   := 1 ;
R:=GS.CellRect(ACol, ARow) ;
if bHaut then begin Canvas.MoveTo(R.Left, R.Top) ; Canvas.LineTo(R.Right+1, R.Top) end
         else begin Canvas.MoveTo(R.Left, R.Bottom-1) ; Canvas.LineTo(R.Right+1, R.Bottom-1) end ;
end ;

procedure TFSaisieFolio.SetGridGrise(ACol, ARow : Integer ; Canvas : TCanvas) ;
var R : TRect ;
begin
Canvas.Brush.Color := GS.FixedColor ;
Canvas.Brush.Style := bsBDiagonal ;
Canvas.Pen.Color   := GS.FixedColor ;
Canvas.Pen.Mode    := pmCopy ;
Canvas.Pen.Style   := psClear ;
Canvas.Pen.Width   := 1 ;
R:=GS.CellRect(ACol, ARow) ;
Canvas.Rectangle(R.Left, R.Top, R.Right+1, R.Bottom+1) ;
end ;


procedure TFSaisieFolio.PostDrawCell(ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
var bGrise, bSep, bHaut : Boolean ;
    ECR     : TZF ;
    VV      : Variant ;
    ii       : integer ;
    OkR : boolean ;
        Rect    : TRect ;
    T1,T2,T3 : TPoint ;
begin
if (ARow<GS.FixedRows) or (ARow>nbLignes+1) then Exit ;
bGrise:=FALSE ; bSep:=FALSE ; bHaut:=FALSE ;
if (ACol=SF_DEBIT)  and (GS.Cells[SF_CREDIT, ARow]<>'') and (GS.Cells[SF_DEBIT, ARow]='') then bGrise:=TRUE ;
if (ACol=SF_CREDIT) and (GS.Cells[SF_DEBIT, ARow]<>'') and (GS.Cells[SF_CREDIT, ARow]='') then bGrise:=TRUE ;
if (ACol=SF_AUX) and (GS.Cells[SF_AUX, ARow]='') and (GS.Cells[SF_JOUR, ARow]<>'') then bGrise:=TRUE ;
if (ACol<>SF_NUML) and (GS.Cells[SF_PIECE, ARow]<>GS.Cells[SF_PIECE, ARow+1])
   and (GS.Row<>ARow) then bSep:=TRUE ;
if (ACol<>SF_NUML) and (GS.Cells[SF_PIECE, ARow]<>GS.Cells[SF_PIECE, ARow-1])
   and (GS.Row=ARow-1) then begin bSep:=TRUE ; bHaut:=TRUE ; end ;
if bGrise then SetGridGrise(ACol, ARow, Canvas);
if (not memParams.bLibre) and (bSep) then SetGridSep(ACol, ARow, Canvas, bHaut);
if ACol<>SF_GEN then Exit ;
if ObjFolio<>nil then Ecr:=ObjFolio.GetRow(ARow-GS.FixedRows) else Ecr:=nil ;
if Ecr<>Nil then
   BEGIN
   if Pos('D',Ecr.GetValue('COMPTAG')) > 0   then
    begin
     // CDessineTriangle(GS) ;
     Rect:=GS.CellRect(SF_NAT,ARow) ;
     Canvas.Brush.Color:= clBlue ; Canvas.Brush.Style:=bsSolid ; Canvas.Pen.Color:=clBlue ;
     Canvas.Pen.Mode:=pmCopy ; Canvas.Pen.Width:= 1 ;
     T1.X:=Rect.Right-5 ; T1.Y:=Rect.Top+1 ;
     T2.X:=T1.X+4       ; T2.Y:=T1.Y ;
     T3.X:=T2.X         ; T3.Y:=T2.Y+4 ;
     Canvas.Polygon([T1,T2,T3]) ;
    // exit ;
    end ;
    if Pos('P',Ecr.GetValue('COMPTAG')) > 0   then
    begin
     // CDessineTriangle(GS) ;
     Rect:=GS.CellRect(SF_NAT,ARow) ;
     Canvas.Brush.Color:= clPurple ; Canvas.Brush.Style:=bsSolid ; Canvas.Pen.Color:=clPurple ;
     Canvas.Pen.Mode:=pmCopy ; Canvas.Pen.Width:= 1 ;
     T1.X:=Rect.Left+5 ; T1.Y:=Rect.Top+1 ;
     T2.X:=T1.X-4       ; T2.Y:=T1.Y ;
     T3.X:=T2.X         ; T3.Y:=T2.Y-4 ;
     Canvas.Polygon([T1,T2,T3]) ;
    // exit ;
    end ;
   VV:=Ecr.GetValue('E_BLOCNOTE') ;
   if VarType(VV)<>VarNull then if trim(VV)<>'' then
      BEGIN
      Ecr.PutValue('COMPS','X') ;
      if RichEdit.Lines.Count>0 then
         BEGIN
         OkR:=False ; for ii:=0 to RichEdit.Lines.Count-1 do if RichEdit.Lines[ii]<>'' then OkR:=True ;
         if OkR then
            BEGIN
            // CDessineTriangle(GS) ;
            Rect:=GS.CellRect(SF_GEN,ARow) ;
            Canvas.Brush.Color:= clRed ; Canvas.Brush.Style:=bsSolid ; Canvas.Pen.Color:=clRed ;
            Canvas.Pen.Mode:=pmCopy ; Canvas.Pen.Width:= 1 ;
            T1.X:=Rect.Right-5 ; T1.Y:=Rect.Top+1 ;
            T2.X:=T1.X+4       ; T2.Y:=T1.Y ;
            T3.X:=T2.X         ; T3.Y:=T2.Y+4 ;
            Canvas.Polygon([T1,T2,T3]) ;
            END ;
         END ;
      END ;
   END ;
end ;

procedure TFSaisieFolio.SetGridOn(bMove : Boolean; NumLigne : Integer; bFocus : Boolean) ;
begin
if bMove then
  begin
  GS.Enabled:=TRUE ;
  if NumLigne>0 then GS.Row:=NumLigne else if bModeRO then GS.Row:=nbLignes else GS.Row:=GS.FixedRows+nbLignes ;
  end ;
if bModeRO then
  begin
  if bMove then if GS.Row-1<GS.FixedRows then GS.Row:=GS.FixedRows ;
  SetGridRO ;
  end ;
if (GS.Enabled) and (GS.Visible) and (not bClosing) then
  begin
  if bFocus then SetGridOptions(GS.Row) ;
  if GS.canFocus then
   GS.SetFocus ;
  end ;
end ;

procedure TFSaisieFolio.SetGridRO ;
begin
GS.CacheEdit ;
GS.Options:=GS.Options-[GoEditing, GoTabs, GoAlwaysShowEditor] ;
GS.Options:=GS.Options+[GoRowSelect] ;
GS.Invalidate ;
end ;

procedure TFSaisieFolio.SetGridOptions(Row : LongInt) ;
var Lab  : THLabel ;
    iErr : integer ;
    Okok : Boolean ;
begin
{$IFDEF TT}
AddEvenement('SetGridOptions Row:='+intToStr(Row)) ;
{$ENDIF}
Okok:=(GoRowSelect in GS.Options) ;
LabelInfos.Visible:=FALSE ;
LabelInfos2.Visible:=FALSE ;
Lab:=LabelInfos2 ;
if not bModeRO then
  begin
  iErr:=IsRowCanEdit(Row) ;
  if iErr<=0 then
    begin
    GS.CacheEdit ;
    GS.Options:=GS.Options-[GoRowSelect] ;
    GS.Options:=GS.Options+[GoEditing,GoTabs,GoAlwaysShowEditor] ;
    GS.MontreEdit ;
    if GS.Col = GS.ColCount - 1 then
     begin
      GS.Col := GS.Col - 1 ;
      GS.Col := GS.Col + 1 ;
     end
      else
       begin
        GS.Col := GS.Col + 1 ;
        GS.Col := GS.Col - 1 ;
       end;
    GS.Invalidate ;
    end else
    begin
    GS.CacheEdit ;
    GS.Options:=GS.Options-[GoEditing,GoTabs,GoAlwaysShowEditor] ;
    GS.Options:=GS.Options+[GoRowSelect] ;
    Lab.Caption:=GetMessageRC(iErr) ;
    Lab.Visible:=TRUE ;
    GS.Invalidate ;
    end ;
  end ;
if bModeRO then
  begin
  iErr:=IsRowCanEdit(Row) ;
  if iErr>0 then
     BEGIN
     Lab.Caption:=GetMessageRC(RC_ECRLETTREE) ;
     Lab.Visible:=TRUE ;
     END ;
  end ;
if Okok<>(GoRowSelect in GS.Options) then GS.Refresh ;
end ;

procedure TFSaisieFolio.SetColVal(ACol : LongInt; NumPiece, Val : string) ;
var i : Integer ; bVal : Boolean ;
begin
bVal:=FALSE ;
//if not IsPieceCanEdit(NumPiece) then Exit ;
for i:=GS.FixedRows to nbLignes do
    begin
    if GS.Cells[SF_PIECE,i]=NumPiece then
       begin GS.Cells[ACol,i]:=Val ; bVal:=TRUE ; Continue ; end ;
    if bVal then break ;
    end ;
end ;

function TFSaisieFolio.GetLib( vRow, vRowRef : integer ) : string ;
var
 lStLibCompte : string ;
 lStNumCompte : string ;
 k            : integer ;
begin

 lStLibCompte := '' ;

 if memJal.ModeSaisie <> 'LIB' then
  lStLibCompte := GS.Cells[SF_LIB, vRowRef] ;

 if lStLibCompte = '' then
  begin
   lStNumCompte := GS.Cells[SF_AUX, vRowRef] ;
   k         := Tiers.GetCompte(lStNumCompte) ;
   if k>=0 then
    lStLibCompte:=Tiers.GetValue('T_LIBELLE', k)
     else
      begin
       lStNumCompte:=GS.Cells[SF_GEN, vRowRef] ; k:=Comptes.GetCompte(lStNumCompte) ;
       if k>=0 then lStLibCompte:=Comptes.GetValue('G_LIBELLE', k) ;
      end ;
  end ;// if

 if memJal.ModeSaisie = 'LIB' then
  begin
   if (lStNumCompte<>'') and (GS.Cells[SF_LIB, vRowRef]<>lStLibCompte) then
    lStLibCompte:=GS.Cells[SF_LIB, vRowRef]
     else
      lStLibCompte := '' ;
  end ;

 result := lStLibCompte ;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 12/06/2002
Modifié le ... : 21/09/2005
Description .. : - 12/06/2002 - quand on cree une nouvelle ligne, on
Suite ........ : affecte directement son numero.
Suite ........ : -affectation du statut a creat
Suite ........ : - 17/12/2002 - fiche 10515 - correction de l'etat des buttons
Suite ........ : dès la premiere ligne
Suite ........ : -29/09/2003 - FB 12694 - on affecte le TOBM a la grille
Suite ........ : pour la gestion du tobm
Suite ........ : -25/03/2004 - On stocke le numero de ligne que l'on vient
Suite ........ : de creer pour pouvoir ce replacer lors du click ds la grille
Suite ........ : - 02/07/2004 - FB 13079 - on reprends systematiquement le
Suite ........ : libelle du compte de la premiere ligne
Suite ........ : - 02/09/2004 - LG - FB 14247 - l'appel des scenario ne fct
Suite ........ : pas pour un seul etablissement
Suite ........ : - LG - 13/09/2004 - FB 12778 - affectation du e_docid sur
Suite ........ : les nouvelles lignes
Suite ........ : - LG - 28/10/2004 - en mode libre, on ne reprneds pas
Suite ........ : systematiquement le libelle de la ligne du dessus
Suite ........ : - LG - 21/01/2005 - FB 12620 - on ne reprends pas la
Suite ........ : reference pour un bordereau en mode libre
Suite ........ : - LG - 20/07/2005 - FB 10794 - qd on passe au bordereau
Suite ........ : suivant on affecte la reference+1
Suite ........ : - LG - 21/09/2005 -  FB 13703 - modif de la gestion de la ref
Suite ........ : ( voir fiche de bug )
Mots clefs ... :
*****************************************************************}
procedure TFSaisieFolio.CreateRow(Row : LongInt; bRowEcart : Boolean(*=FALSE*)) ;
var Ecr, Prec : TZF ;
    bInsert : Boolean ; RowRef : LongInt ;

begin
Statut:=taCreat ;
{$IFDEF TT}
AddEvenement('CreateRow Row:='+intToStr(Row));
{$ENDIF}
if GS.RowCount-5<nbLignes then GS.RowCount:=GS.RowCount+10 ;
RowRef:=Row-1 ; bInsert:=FALSE ;
if Row<=nbLignes then
  begin
   bInsert:=TRUE ;
   GS.InsertRow(Row) ;
   RowRef:=Row+1 ;
  end ;
Inc(nbLignes) ;
// Etablissement par défaut
GS.Cells[SF_ETABL, Row]:=E_ETABLISSEMENT.Value ;
// Qualif Pièce par défaut
GS.Cells[SF_QUALIF, Row]:=memOptions.QualifP ;
// Ligne créée en saisie
GS.Cells[SF_NEW, Row]:='O' ;
// Numérotation de la ligne
if Row=nbLignes then // on est sur la derniere ligne
 GS.Cells[SF_NUML, Row]:=intToStr(RowRef+1)
  else
   SetNum(Row) ;
// Mise à jour du numéro de pièce virtuel
//GS.Cells[SF_PIECE, Row]:='1' ;
if memParams.bLibre then GS.Cells[SF_PIECE, Row]:='1'
 else
  if bInsert then
   GS.Cells[SF_PIECE, Row]:=GS.Cells[SF_PIECE, RowRef]
    else
     SetPiece ;
// Nature de pièce par défaut
//LG* 26/11/2001 correction de la nature pour les lignes en ecart
if nbLignes>1 then
begin
 if bRowEcart then GS.Cells[SF_NAT, Row]:=GS.Cells[SF_NAT, Row-1] // on recupere la reference precedente
              else GS.Cells[SF_NAT, Row]:=GS.Cells[SF_NAT, RowRef] ;
 FZScenario.Load([MemJal.Code,GetCellNature(Row),'N',E_ETABLISSEMENT.Value]) ; // chargement des scenarios
end
 else GS.Cells[SF_NAT, Row]:=InitNature ;
// Dernier jour du mois par défaut
if nbLignes>1 then GS.Cells[SF_JOUR,Row]:=GS.Cells[SF_JOUR, RowRef]
              else GS.Cells[SF_JOUR,Row]:=IntToStr(memOptions.MaxJour) ;
// Référence par défaut
if not memParams.bLibre then
begin
if ((FLastRef = '') and ( RowRef >0 )) then FLastRef := GS.Cells[SF_REFI, RowRef] ;
//if (nbLignes>1) then
  if (not bRowEcart) and (GS.Cells[SF_PIECE, Row]<>GS.Cells[SF_PIECE, RowRef]) then
   GS.Cells[SF_REFI, Row]:=IncRefAuto(FLastRef)
    else
     if not IsSoldeFolio then
      begin
       if ( RowRef >0 ) then
        GS.Cells[SF_REFI, Row]:= GS.Cells[SF_REFI, RowRef]
         else
          GS.Cells[SF_REFI, Row]:=FLastRef ;
      end ;
end ;
// Libellé par défaut
//LG* 22/11/2001  on ne recupère pas le libelle sur un chg de piece
if GS.Row=1 then bNewLigne:=true else bNewLigne:=GS.Cells[SF_PIECE, Row]<>GS.Cells[SF_PIECE, RowRef] ;
if bNewligne then memJal.CurAuto:=1 ;
if (nbLignes>1) and (not bGuideRun) and (not bNewLigne) and ( not IsSoldeFolio ) then
 GS.Cells[SF_LIB, Row]:= GetLib(Row,RowRef) ;
// Devise+Euro
if (GS.Cells[SF_PIECE, Row]=GS.Cells[SF_PIECE, RowRef]) or (memParams.bDevise) then
   begin
   // Devise par défaut
   GS.Cells[SF_DEVISE, Row]:=GS.Cells[SF_DEVISE, RowRef] ;
   // Euro par défaut
   GS.Cells[SF_EURO, Row]:=GS.Cells[SF_EURO, RowRef] ;
   end ;
// Création de l'objet correspondant
Ecr:=ObjFolio.CreateRow(nil, Row-GS.FixedRows) ;
GS.Objects[SA_Exo,Row]:=Ecr ;
if Ecr<>nil then
   begin
  // Ecr.PutValue('INSROW', 'X') ;
   Ecr.PutValue('E_DEVISE', GS.Cells[SF_DEVISE, Row]) ;
   if GS.Cells[SF_PIECE, Row]=GS.Cells[SF_PIECE, RowRef] then
     begin
     if (Row-GS.FixedRows-1)>=0 then Prec:=ObjFolio.GetRow(Row-GS.FixedRows-1) else Prec:=nil ;
     if Prec<>nil then Ecr.PutValue('E_REGIMETVA', Prec.GetValue('E_REGIMETVA')) ;
     end ;
   end ;
//BackupFolio ;
//LG* 21/02/2002 pour une nouvelle ligne, on se place sur le jour
if (nbLignes>1) and (not bGuideRun) and not bNewLigne and memParams.bLibre then
 GS.Col:=SF_JOUR
  else
   if (nbLignes>1) and (not bGuideRun) and not bNewLigne then
    GS.Col:=SF_GEN ;
if (nbLignes>1) and (bNewLigne) and (not bGuideRun) then
 GS.Col:=SF_JOUR ;
FInRow := RowRef + 1;
InfosPied(Row) ;
EnableButtons ;
end ;


procedure _DecS( var Value : string ) ;
Var sRef,sNum : String ;
    ii,ll     : integer ;
BEGIN
sRef:=Value ;
sNum:='' ; ii:=Length(sRef) ; ll:=0 ;
While ((ii>=1) and (sRef[ii] in ['0'..'9'])) do BEGIN sNum:=sRef[ii]+sNum ; Dec(ii) ; inc(ll) ; END ;
if sNum='' then Exit ;
sNum:=IntToStr(ValeurI(sNum)-1) ; if ValeurI(sNum)<1 then sNum:='1' ;
sRef:=Copy(sRef,1,Length(sRef)-ll)+sNum ;
Value:=sRef ;
END ;
{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 06/09/2002
Modifié le ... : 20/07/2005
Description .. : - 06/10/2002 - suppression du test de validite sur la
Suite ........ : premiere ligne
Suite ........ : - 23/09/2002 - on ne peut pas supprimer la premiere ligne
Suite ........ : s'il elle a été synchroniser avec S1 ( champs
Suite ........ : E_REFREVISION renseigne )
Suite ........ : -02/10/2002 - on ne pas pouvoir supprimer les ecart de
Suite ........ : conversion
Suite ........ : -03/10/2002- ajout d'un parametre bPasLigneEcart pour
Suite ........ : forcer la suppression des lignes d'ecart de conversion par la
Suite ........ : fct VerifEquilibre
Suite ........ : - 20/05/2005 - FB  10794 - qd on supprime une nouvelle 
Suite ........ : ligne on decremete la ref interne
Mots clefs ... : 
*****************************************************************}
function TFSaisieFolio.DeleteRow(Row : LongInt; bPasLigneEcart : boolean=false) : Integer ;
var PieceCur : string ; ACol : Integer ; bValid : Boolean ;
Ecr : TZF ;
//NumCompte : string ; i : integer ;
//D,C : double ;
begin
{$IFDEF TT}
AddEvenement('DeleteRow Row:='+intToStr(Row)) ;
{$ENDIF}
result := 0 ;
if csDestroying in ComponentState then Exit ;
if bClosing then exit ;
if IsRowCanEdit(Row) > 0 then exit ;
if memParams.bLibre then Statut:=taConsult ; HGBeginUpdate(GS) ;
try
if BGuideRun then Exit ;
if ObjFolio = nil then exit ;
Ecr:=ObjFolio.GetRow(Row-GS.FixedRows) ; if Ecr=nil then Exit ;
if (nbLignes=1) and (Ecr.GetValue('E_REFREVISION')<>0) then exit ;
bValid:=IsRowValid(Row, ACol, FALSE) ;
PieceCur:=GS.Cells[SF_PIECE, Row] ;
ObjFolio.DeleteRow(Row-GS.FixedRows) ;
GS.DeleteRow(Row) ;
if bNewLigne and memJal.IncRef then _DecS(FLastRef) ;
if GS.RowCount>nbLignes+10 then GS.RowCount:=GS.RowCount-1
                           else GS.RowCount:=GS.RowCount+10 ;
Dec(nbLignes) ;
SetNum(Row) ; CalculSoldeJournal ;
if (Row=GS.FixedRows) and (nbLignes=0) then
  begin
  if (memParams.bDevise) then
    begin
    memParams.bDevise:=FALSE ;
    memParams.bLibre:=TRUE ; memParams.DevLibre:='' ;
    end ;
  NextRow(Row) ;
  end ;
if (bValid) and (GS.Cells[SF_PIECE, Row]<>PieceCur) and (nbLignes>1)
   then if Row=1 then Result:=0 else Result:=-1 ;
EnableButtons ;
//BackupFolio ;
Statut:=taConsult ;
finally
 HGEndUpdate(GS) ;  
end;
end ;

function TFSaisieFolio.NextRow(Row : LongInt) : Boolean ;
begin
if Row>nbLignes+1 then begin Result:=FALSE ; Exit ; end ;
if Row>nbLignes then
 begin
  CreateRow(Row) ;
  SetGridOptions(Row) ;
 end ;
Result:=TRUE ;
end ;

// En saisie, on considère que la ligne est valide si le compte général est renseigné
{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 13/06/2002
Modifié le ... : 02/10/2002
Description .. : -13/06/2002 - empêche la saisie de deux montant ( par
Suite ........ : l'intermédiaire des guides )
Suite ........ : 
Suite ........ : - 16/09/2002 - modif pour permettre de saisir des ecritures
Suite ........ : sans montant
Suite ........ : 
Suite ........ : - 02/11/2002 - ajout d'un boolean bAvecMontant pour 
Suite ........ : forcer le test sur les montant pour determiner une ligne 
Suite ........ : valide
Mots clefs ... : 
*****************************************************************}
function TFSaisieFolio.IsRowValid(Row : LongInt; var ACol : Integer; bTest : Boolean; bAvecMontant : boolean=false) : Boolean ;
var Ecr : TZF ;
lRdDebit,lRdCredit : double ;
begin
Result:=TRUE ; bGuideStop:=FALSE ;
{$IFDEF TT}
AddEvenement('IsRowValid Row:='+intToStr(Row));
{$ENDIF}
if ObjFolio=nil then Exit ;
Ecr:=ObjFolio.GetRow(Row-GS.FixedRows) ;
if Ecr=nil then begin Result:=FALSE ; Exit ; end ;
//if GS.Row>(Row+1) then begin Result:=FALSE ; Exit ; end ;  // test supprimer -> ne marche pas lors de la generation des ecarts de conversion
//LG* 13/06/2002 empêche la saisie de deux montant ( par l'intermédiaire des guides )
lRdDebit:=Valeur(GS.Cells[SF_DEBIT,Row]) ; lRdCredit:=Valeur(GS.Cells[SF_CREDIT,Row]) ;
if (lRdDebit<>0) and (lRdCredit<>0) then begin Result:=FALSE ; ACol:=SF_GEN ; Exit ; end ;
if GS.Cells[SF_GEN, Row]='' then
   begin
      if (not bTest) and (bGuideRun) then
     begin
     bGuideStop:=TRUE ;
     if GetGuideColStop(Row)<0 then
       begin PrintMessageRC(RC_BADGUIDE) ; GuideStop ; GS.Invalidate ; end ;
     end ;
   Result:=FALSE ; ACol:=SF_GEN ; Exit ;
   end ;

if (GS.Cells[SF_AUX, Row]='') and (IsRowCollectif(Row)) then
   begin
   if (not bTest) and (bGuideRun) then
     begin
     bGuideStop:=TRUE ;
     if GetGuideColStop(Row)<0 then
       begin PrintMessageRC(RC_BADGUIDE) ; GuideStop ; GS.Invalidate ; end ;
     end ;
   Result:=FALSE ; ACol:=SF_AUX ; Exit ;
   end ;
if not bAvecMontant then exit ;
if (GS.Cells[SF_DEBIT, Row]='') and (GS.Cells[SF_CREDIT, Row]='') then
   begin
   if (not bTest) and (bGuideRun) then
     begin
     bGuideStop:=TRUE ;
     if GetGuideColStop(Row)<0 then
       begin PrintMessageRC(RC_BADGUIDE) ; GuideStop ; GS.Invalidate ; end ;
     end ;
   Result:=FALSE ; ACol:=SF_DEBIT ; Exit ;
   end ;

end ;



function TFSaisieFolio.IsPieceCanEdit(Piece: string;bAvecEcartdeConversion : boolean=true) : Boolean ;
var i : Integer ;
begin
Result:=TRUE ;
if IsJalLibre then Exit ;
for i:=GS.FixedRows to nbLignes do
  begin
  if GS.Cells[SF_PIECE, i]<>Piece then Continue ;
  if IsRowCanEdit(i,bAvecEcartdeConversion)>0 then begin Result:=FALSE ; Break ; end ;
  end ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 02/10/2002
Modifié le ... : 25/08/2005
Description .. : -02/10/2002 - les ecart de conversions sont traites comme 
Suite ........ : une ecriture lettrees
Suite ........ : - 12/08/2004 - FB 13560 - on ne modifie pas les ecritures 
Suite ........ : avec des aux /genrraux lettre
Suite ........ : - FB 16397 - LG  - 25/08/2005 - test de la revision 
Suite ........ : uniquement pour l'exercice en cours
Mots clefs ... : 
*****************************************************************}
function TFSaisieFolio.IsRowCanEdit(Row : LongInt; bAvecEcartdeConversion : boolean=true) : integer ;
var Ecr : TZF ;
k : integer ; NumCompte : string ;
begin
{$IFDEF TT}
AddEvenement('IsRowCanEdit Row:='+intToStr(Row)) ;
{$ENDIF}
Result:=0 ; if ObjFolio=nil then Exit ;
Ecr:=ObjFolio.GetRow(Row-GS.FixedRows) ; if Ecr=nil then Exit ; 
if (Ecr.GetValue('E_ETATLETTRAGE')='PL') or (Ecr.GetValue('E_ETATLETTRAGE')='TL') then Result:=RC_ECRLETTREE
 else if Ecr.GetValue('E_REFPOINTAGE')<>'' then Result:=RC_ECRPOINTEE
  else if Ecr.GetValue('E_IMMO')<>'' then Result:=RC_ECRIMMO
   else if (VH^.OuiTvaEnc) and (Ecr.GetValue('E_EDITEETATTVA')='#') then Result:=RC_ECRTVAENC
     else if (Ecr.GetValue('E_REFGESCOM') <> '') and (Ecr.GetValue('E_TYPEMVT') = 'TTC') then Result := RC_GCTTC
      {JP 26/06/07 : FQ TRESO 10491 : on vérouille les flux originaires de la Tréso}
      else if (Ecr.GetValue('E_QUALIFORIGINE') = QUALIFTRESO) then Result := RC_GCTTC;

NumCompte:=GS.Cells[SF_GEN, Row] ; k:=Comptes.GetCompte(NumCompte) ;
if (k>-1) and (Comptes.GetValue('G_FERME', k)='X') then  Result:=RC_CFERME ;
if (k>-1) and (GS.Cells[SF_NEW, Row]<>'O') and (Comptes.GetValue('G_VISAREVISION', k)='X') and (memOptions.TypeExo=teEncours) then  Result:=RC_COMPTEVISA ;
NumCompte:=GS.Cells[SF_AUX, Row] ; k:=Tiers.GetCompte(NumCompte) ;
if (k>-1) and (Tiers.GetValue('T_FERME', k)='X') then Result:=RC_CFERME ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 01/06/2004
Modifié le ... :   /  /
Description .. : - LG - 01/06/2004 - On ne lance pas la fct tant que le
Suite ........ : compte est pas complet
Mots clefs ... :
*****************************************************************}
function TFSaisieFolio.IsRowCollectif(Row : LongInt) : Boolean ;
var NumCompte : string ; k : Integer ;
begin
Result:=FALSE ;
NumCompte:=GS.Cells[SF_GEN, Row] ; if (NumCompte='') or (Length(NumCompte) <> VH^.Cpta[fbGene].Lg)  then Exit ;
k:=Comptes.GetCompte(NumCompte) ;
if k<>-1 then result:=Comptes.IsCollectif(k) ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 01/06/2004
Modifié le ... :   /  /
Description .. : - LG - 01/06/2004 - On ne lance pas la fct tant que le
Suite ........ : compte est pas complet
Mots clefs ... :
*****************************************************************}
function TFSaisieFolio.IsRowTiers(Row : LongInt) : Boolean ;
var NumCompte : string ; k : Integer ;
begin
Result:=FALSE ;
if GS.Cells[SF_AUX, Row]='' then
  begin
  if (NumCompte='') or (Length(NumCompte) <> VH^.Cpta[fbGene].Lg)  then Exit ;
  k:=Comptes.GetCompte(NumCompte) ; if k<0 then Exit ;
  if not Comptes.IsTiers(k) then Exit ;
  end ;
Result:=TRUE ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 01/06/2004
Modifié le ... :   /  /    
Description .. : - LG - 01/06/2004 - On ne lance pas la fct tant que le
Suite ........ : compte est pas complet
Mots clefs ... : 
*****************************************************************}
function TFSaisieFolio.IsRowHT(Row : LongInt) : Boolean ;
var NumCompte : string ; k :Integer ;
begin
Result:=FALSE ; NumCompte:=GS.Cells[SF_GEN, Row] ;
if (NumCompte='') or (Length(NumCompte) <> VH^.Cpta[fbGene].Lg)  then Exit ;
k:=Comptes.GetCompte(NumCompte) ; if k<0 then Exit ;
if not Comptes.IsHT(k) then Exit ;
Result:=TRUE ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 01/06/2004
Modifié le ... :   /  /
Description .. : - LG - 01/06/2004 - On ne lance pas la fct tant que le
Suite ........ : compte est pas complet
Mots clefs ... :
*****************************************************************}
function TFSaisieFolio.IsRowBqe(Row : LongInt) : Boolean ;
var NumCompte : string ; k :Integer ;
begin
Result:=FALSE ; NumCompte:=GS.Cells[SF_GEN, Row] ;
if (NumCompte='') or (Length(NumCompte) <> VH^.Cpta[fbGene].Lg)  then Exit ;
k:=Comptes.GetCompte(NumCompte) ; if k<0 then Exit ;
if not Comptes.IsBqe(k) then Exit ;
Result:=TRUE ;
end ;

function TFSaisieFolio.IsRowDecEnc(Row : LongInt) : string ;
var Solde : Double ;
begin
Result:='' ;
Solde:=Valeur(GS.Cells[SF_DEBIT, Row])-Valeur(GS.Cells[SF_CREDIT, Row]) ;
if Solde=0 then Exit ;
if Solde>0 then Result:='ENC' else Result:='DEC' ;
end ;

function TFSaisieFolio.GetRowTiers(Piece, DefRow : LongInt) : LongInt ;
var i: Integer ;
begin
Result:=-1 ;
for i:=DefRow-1 downto GS.FixedRows do
  begin
  if StrToInt(GS.Cells[SF_PIECE, i])<>Piece then Break ;
  if IsRowTiers(i) then begin Result:=i ; Exit ; end ;
  end ;
for i:=DefRow+1 to nbLignes do
  begin
  if StrToInt(GS.Cells[SF_PIECE, i])<>Piece then Break ;
  if IsRowTiers(i) then begin Result:=i ; Exit ; end ;
  end ;
end ;
{
function TFSaisieFolio.GetRowHT(Piece, DefRow : LongInt) : LongInt ;
var i: Integer ;
begin
Result:=-1 ;
for i:=DefRow-1 downto GS.FixedRows do
  begin
  if StrToInt(GS.Cells[SF_PIECE, i])<>Piece then Break ;
  if IsRowHT(i) then begin Result:=i ; Exit ; end ;
  end ;
for i:=DefRow+1 to nbLignes do
  begin
  if StrToInt(GS.Cells[SF_PIECE, i])<>Piece then Break ;
  if IsRowHT(i) then begin Result:=i ; Exit ; end ;
  end ;
end ; }

function TFSaisieFolio.IsJalLibre : Boolean ;
begin
Result:=(memParams.bLibre) or (memParams.bDevise) ;
end ;

//=======================================================
//================= Evénements du Grid ==================
//=======================================================
{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 19/07/2002
Modifié le ... :   /  /
Description .. : - 19/07/2002 - utilisation de la fonction
Suite ........ : FRechCompte.ElipsisClick
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.GSElipsisClick(Sender: TObject);
begin
if csDestroying in ComponentState then Exit ;
AppelElipsisClick ( false );
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 03/12/2002
Modifié le ... : 03/02/2006
Description .. : -03/12/2002 - fiche  10698 suppression du &#@ apres
Suite ........ : l'appel de GetCorrespType
Suite ........ : -02/07/2003 - qd on a choisit un compte ds le lookup, on
Suite ........ : passe a la case suivante
Suite ........ : - LG - 13/08/2004 - FB 13012 - afiichage des info sur le
Suite ........ : comptes
Suite ........ : - LG - 22/09/2004 - FB 14197 - l'ouverture du lookupList est
Suite ........ : bloque si l'auxi existe
Suite ........ : - LG - 28/10/2004 - qd on clique sur l'ellipsis et que le
Suite ........ : compte est complet, on ouvre la fenetre des comptes
Suite ........ : - LG - 13/05/2005 - FB 15667 - on rafraichi les info du
Suite ........ : compte qd on clique sur l'elipsis. fct MajInfoCompte
Suite ........ : - LG - 03/02/2005 - FB 17370 - mise ajour des info sur l'auxi
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.AppelElipsisClick     ( vBoAvecDeplacement : boolean = true ) ;
var Val, StTable, StCode, StWhere, StPrefixe, StLib : string ;
ACol,Arow : integer ;
lStAux,lStGen : string ;
k : integer ;
Ecr : TZF ;
begin
if csDestroying in ComponentState then Exit ;
if IsRowCanEdit(GS.Row)>0 then Exit ;
GrilleModif:=true ;
{$IFDEF TT}
AddEvenement('GSElipsisClick Row:='+intToStr(GS.Row)) ;
{$ENDIF}
if (Screen.ActiveControl<>GS) and (not bClosing) then GS.SetFocus ;
Statut:=taModif ;
//FRechCompte.StNature:=GetCellNature(GS.Row) ; FRechCompte.StJournal:=memJal.Code ; FRechCompte.StDevise:=FInfo.Devise.Dev.Code ;
ACol:=GS.Col ; ARow:=GS.Row ;
Ecr:=ObjFolio.GetRow(ARow-GS.FixedRows) ;
FRechCompte.Ecr := Ecr ;
lStAux := GS.Cells[SF_AUX ,GS.Row] ;
if FRechCompte.ElipsisClick(GS,vBoAvecDeplacement) then
 begin
  InfosPied(-1) ;
  GereCutOff(Arow) ;
  MajInfoCompte(ARow,ACol) ;
  lStGen := GS.Cells[SF_GEN ,GS.Row] ;
  k:=Comptes.GetCompte(lStGen) ;
  if (k>=0) and not (Comptes.GetValue('G_COLLECTIF', k)='X') then
      begin
      GS.Col:=SF_AUX+1 ;
      GS.ElipsisButton:=false ;
      SetGridOptions(ARow) ;  // FB 11819 - le rafraichissement de la grille etaient incorrecte // FB 13123 remis!
    end ;
  if trim(GS.Cells[SF_AUX, ARow])<>'' then
   FiniAux(ACol,ARow) ;
 end ;
// Nature de pièce
if GS.Col=SF_NAT then
   begin
    Val:=GS.Cells[GS.Col,GS.Row] ;
    {$IFDEF EAGCLIENT}
    GS.Cells[GS.Col,GS.Row]:='' ;
    {$ENDIF} 
    E_NATUREPIECE.Plus:='' ;
    GetCorrespType(E_NATUREPIECE.DataType,StTable,StCode,StWhere,StPrefixe,StLib) ;
    ChangeWhereTT(StWhere,'',false) ;
    if FBoNatureComplete then LookupList(GS,TraduireMemoire('Nature'),StTable,StLib,StCode,StWhere,'',FALSE,0)
    else LookupList(GS,TraduireMemoire('Nature'),StTable,StCode,StLib,StWhere,'',FALSE,0) ;
    if GS.Cells[GS.Col,GS.Row]='' then GS.Cells[GS.Col,GS.Row]:=Val ;
   end ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 13/04/2007
Modifié le ... : 20/07/2007
Description .. : - FB 19806 - LG - le changement d'auxi ne fct plus ( 601 -> 
Suite ........ : 471 l'anlytiq restait en base ) + optimisation de la gestion de 
Suite ........ : l'analytiq ds le bordereau
Suite ........ : - FB 21071 - LG - 20/07/2007 - on affiche les infos pied en 
Suite ........ : eAGl
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.MajInfoCompte ( ARow,ACol : integer ) ;
var
 Ecr : TZF ;
 lTOBAna : TZF ;
 lIndexGen : integer ;
 NumCompte : string ;
begin
 NumCompte:=GS.Cells[SF_GEN, ARow] ;
 lIndexGen:=Comptes.GetCompte(NumCompte) ;
 if lIndexGen = -1 then exit ;
 InfosPied(ARow,false) ;
 {$IFNDEF EAGLCLIENT}

 {$ENDIF}
 Ecr:=ObjFolio.GetRow(ARow-GS.FixedRows) ;
 if (lIndexGen>=0) and (Comptes.GetValue('G_CONFIDENTIEL')>Ecr.getValue('E_CONFIDENTIEL')) then
  Ecr.PutValue('E_CONFIDENTIEL',Comptes.GetValue('G_CONFIDENTIEL' ) ) ;
 if (lIndexGen>=0) and (Comptes.IsHT(lIndexGen)) then
   begin
   FStGenHT:=GS.Cells[SF_GEN, ARow] ;
   if Ecr<>nil then
     begin
     Ecr.PutValue('E_TVA', Comptes.GetValue('G_TVA', lIndexGen)) ;
     Ecr.PutValue('E_TPF', Comptes.GetValue('G_TPF', lIndexGen)) ;
     end ;
   end ;

  if (lIndexGen>=0) and (Comptes.IsTiers(lIndexGen)) then
    if not PutRib(ACol, ARow, FALSE) then PrintMessageRC(RC_RIB) ;

  if (Ecr<>nil) and (Ecr.GetValue('E_ANA')='X') and (Ecr.GetValue('E_GENERAL')<>'') and (Ecr.GetValue('E_GENERAL')<>GS.Cells[ACol, ARow]) then
    ObjFolio.DelAna(ARow-GS.FixedRows) ;
  lTOBAna := ObjFolio.GetRowAna(ARow-GS.FixedRows,0,0) ; // LG 13/04/20007
  if ( lTOBAna <> nil ) and ( lTOBAna.GetValue('Y_GENERAL') <> Ecr.GetString('E_GENERAL') ) then
   ObjFolio.DelAna(ARow-GS.FixedRows) ;

  //if not Cancel then
  PutRegime(ARow) ;
  if (Ecr<>nil) and (Ecr.GetValue('E_CONSO')='') then PutConso(ARow) ;

  if (Ecr.GetValue('E_GENERAL')<>'') and (Ecr.GetValue('E_GENERAL')<>GS.Cells[ACol, ARow]) then
   begin
    Ecr.PutValue('COMPTAG','-') ;
    SupprimeEcrCompl(Ecr) ;
   end ; // if

   Ecr.PutValue('E_GENERAL', GS.Cells[SF_GEN, ARow]) ;
  {$IFNDEF EAGLCLIENT}
  if not bGuideRun then EnableButtons ;
  {$ENDIF}

end ;



{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 28/11/2002
Modifié le ... :   /  /
Description .. : - 28/11/2002 - correction pour le click dans la grille des
Suite ........ : l'ouverture du bordereau
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.GSEnter(Sender: TObject);
begin
 {$IFDEF TT}
 AddEvenement('GSEnter');
 {$ENDIF}
if bFirst then
   begin
   if (FAction<>taConsult) and not ( CEstBloqueJournal(E_JOURNAL.Value,false) ) then
    CBloqueurJournal(true,E_JOURNAL.Value) ;
   FBoCurseurDansGrille := true ;
   if not bModeRO then
    begin
     if not NextRow(GS.Row) then
      begin
       GS.Row:=nbLignes+1 ;
       NextRow(GS.Row) ;
      end;
    end; // if
    // se replacer au début
   SoldeProgressif(GS.Row) ;
  // EnableButtons ;
  if (not bModeRO) and (GoEditing in GS.Options) then
     begin
        E_JOURNAL.Enabled       := false ;
        E_DATECOMPTABLE.Enabled := false;
        CbFolio.Enabled         := false ;
      if not FBoClickSouris then
       begin // on ne doit pas declencher d'evement windows qd on clique directement sur la grille
       // astuce pour donner le focus à la grille
        PostMessage(GS.Handle, WM_KEYDOWN, VK_TAB,  0) ;
        PostMessage(GS.Handle, WM_KEYDOWN, VK_LEFT, 0) ;
      end ;
     end ;

   end ;
bFirst:=FALSE ; // pour empecher cet evenement de ce declencher plusieur fois
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 05/04/2002
Modifié le ... : 08/12/2003
Description .. : // LG* - 05/04/2002 - correction bug : en mode guide on
Suite ........ : peut se deplace avec la souris -> pas de cellExit -> le
Suite ........ : compte general n'etait pas controle -> en mode guide on
Suite ........ : verifie systematiquement la validite du compte general
Suite ........ :
Suite ........ : - 12/06/2002 - desactive les evenements de la grille
Suite ........ : pendant le traitement
Suite ........ : 
Suite ........ : - 08/12/2003 - FB 12123 - remit le setgrigoption
Suite ........ : 
Suite ........ : - 03/10/2002 - modif pour la gestion des ecarts de
Suite ........ : conversion comme les ecritures lettrees
Suite ........ : -suppression de l'appel de la fct isRowSpecial
Suite ........ :
Suite ........ : - 20/01/2003 - FB 11819 - le rafraichissement de la grille
Suite ........ : etaient incorrecte
Mots clefs ... :
*****************************************************************}
procedure TFSaisieFolio.GSCellEnter(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
var lc,lr,sens,k,lColStop : Integer ;
    NumCompte : string ;
    SC : Byte ;
    Z,NatPiece  : string ;
    //lStAux,lStGen : string ;
begin
if csDestroying in ComponentState then Exit ;
if not (Screen.ActiveControl=GS) then exit ;
//CbFolio.Enabled         := false ;
{$IFDEF TT}
AddEvenement('GSCellEnter ACol:='+intToStr(ACol)+' ARow:='+intToStr(ARow));
{$ENDIF}
HGBeginUpdate(GS) ;
try
// Bouton Elipsis
GS.ElipsisButton:=FALSE ;
lc:=GS.Col ; lr:=GS.Row ; sens:=CGetGridSens(GS,ACol, ARow) ;
// En consultation, c'est fini !
if bModeRO then Exit ;
// Pas de saisie dans la colonne Devise
if lc=SF_DEVISE then begin ACol:=SF_DEVISE+sens ; ARow:=lr ; Cancel:=TRUE ; Exit ; end ;
// En saisie guidée ou saisie non libre, interdit de sortir de la pièce
//showmessage( GS.Cells[SF_PIECE,lr] +':'+GS.Cells[SF_PIECE, ARow] ) ;
 if ((bGuideRun) or (not memParams.bLibre)) and
   (GS.Cells[SF_PIECE,lr]<>GS.Cells[SF_PIECE, ARow]) and
   (GS.Cells[SF_PIECE, ARow]<>'') and
   (not IsSoldePiece(StrToInt(GS.Cells[SF_PIECE, ARow]))) then
   BEGIN
   // pour le mode bordereau : la piece n'est pas solde on reste dans celle ci. On rafraichi la ligne courante
   ACol:=SF_GEN ; SetGridOptions(ARow) ; Cancel:=TRUE ; Exit ;
   END ;

if lc = SF_REFI then
 GereCutOff( GS.Row ) ;
// Utiliser le sens par défaut du compte quand aucun montant saisit sur la ligne
NatPiece:=GetCellNature(GS.Row) ;
if (not bGuideRun) and (lc=SF_DEBIT) and (NatPiece<>'OD') and
   (GS.Cells[SF_DEBIT,lr]='') and (GS.Cells[SF_CREDIT,lr]='') then
   begin
   NumCompte:=GS.Cells[SF_GEN, GS.Row] ; k:=Comptes.GetCompte(NumCompte) ;
   // QuelSens : Dernier paramètre : 1=Débit ; 2=Crédit ; 3=Mixte
   if (k>=0) and (sens=1) then
    begin
//LG* 23/11/2001 correction du sens des comptes
     Z:=Comptes.GetValue('G_SENS', k) ; if Z='' then Z:='D' ;
     case Z[1] of 'D' : SC:=1 ; 'C' : SC:=2 ; else  SC:=3 ; end ;
     if ( QuelSens(NatPiece, Comptes.GetValue('G_NATUREGENE', k), SC)) = 2 then  // quelSens = 2 := credit
      begin
       ACol:=SF_CREDIT ; ARow:=lr ; Cancel:=TRUE ; Exit ;
      end;
    end ; // if k>0
   end ;
if (not bGuideRun) and (lc=SF_CREDIT) and
   (GS.Cells[SF_DEBIT,lr]='') and (GS.Cells[SF_CREDIT,lr]='') then
   begin
   NumCompte:=GS.Cells[SF_GEN, GS.Row] ; k:=Comptes.GetCompte(NumCompte) ;
   if (k>=0) and (sens=-1) and not (Comptes.GetValue('G_SENS', k)='D') then
      begin
      ACol:=SF_DEBIT ; ARow:=lr ; Cancel:=TRUE ; Exit ;
      end ;
   end ;
// Interdir la saisie d'un montant au débit si montant au crédit
if (lc=SF_DEBIT) and (GS.Cells[SF_CREDIT,lr]<>'') then
  begin
   PasseColSuivante(ACol,ARow,GS);
   Cancel:=TRUE ; Exit ;
  end ;
// Interdir la saisie d'un montant au crédit si montant au débit
if (lc=SF_CREDIT) and (GS.Cells[SF_DEBIT,lr]<>'') then
  begin
    PasseColSuivante(ACol,ARow,GS);
    Cancel:=TRUE ; Exit ;
  end ;
// Interdir la saisie d'un auxiliaire sur un compte non collectif
if (lc=SF_AUX) then
   begin
   NumCompte:=GS.Cells[SF_GEN, GS.Row] ; k:=Comptes.GetCompte(NumCompte) ;
//   sd := Comptes.GetValue('G_COLLECTIF', k) ;
   if (k>=0) and not (Comptes.GetValue('G_COLLECTIF', k)='X') then
      begin
      if (sens=1) then ACol:=SF_AUX+1 else ACol:=SF_AUX-1 ;
      SetGridOptions(ARow) ;  // FB 11819 - le rafraichissement de la grille etaient incorrecte // FB 13123 remis!
      Cancel:=TRUE ;
      ARow := GS.Row ; // on remplace la grille sur la ligne courante, pour ne pas rester sur une ligne qui a ete supprimé
      Exit ;
      end ;
   end ;
// Saisie en devise
if ((lc=SF_DEBIT) or (lc=SF_CREDIT)) and
   (GS.Cells[SF_DEVISE,lr]<>'') and
   (GS.Cells[SF_DEBIT,lr]='') and (GS.Cells[SF_CREDIT,lr]='') then
   begin
   if not bStopDevise then DevClick ;
   if sens=1 then PostMessage(GS.Handle, WM_KEYDOWN, VK_TAB, 0)
             else PostMessage(GS.Handle, WM_KEYDOWN, VK_LEFT, 0) ;
   Exit ;
   end ;
// Saisie en devise au débit
if (lc=SF_DEBIT) and
   ((GS.Cells[SF_DEVISE,lr]<>'') or (GS.Cells[SF_EURO,lr]='X')) and
   (GS.Cells[SF_DEBIT,lr]<>'') then
   begin
   if bStopDevise then bStopDevise:=FALSE
                  else begin DevClick ; bStopDevise:=TRUE ; end ;
   if sens=1 then PostMessage(GS.Handle, WM_KEYDOWN, VK_TAB, 0)
             else PostMessage(GS.Handle, WM_KEYDOWN, VK_LEFT, 0) ;
   Exit ;
   end ;
// Saisie en devise au crédit
if (lc=SF_CREDIT) and
   ((GS.Cells[SF_DEVISE,lr]<>'') or (GS.Cells[SF_EURO,lr]='X')) and
   (GS.Cells[SF_CREDIT,lr]<>'') then
   begin
   if bStopDevise then bStopDevise:=FALSE
                  else begin DevClick(GS.Cells[SF_EURO,lr]='X') ; bStopDevise:=TRUE ; end ;
   if sens=1 then PostMessage(GS.Handle, WM_KEYDOWN, VK_TAB, 0)
             else PostMessage(GS.Handle, WM_KEYDOWN, VK_LEFT, 0) ;
   Exit ;
   end ;
// Bouton Elipsis
{$IFNDEF EAGLCLIENT}
GS.ElipsisButton:=( (lc=SF_NAT) or (lc=SF_GEN) or (lc=SF_AUX) ) ;
{$ELSE}
GS.ElipsisButton:= lc=SF_NAT ;
{$ENDIF}
if (lc=SF_NAT) then GS.ElipsisHint:=GetMessageRC(RC_SELNATURE)
  else if (lc=SF_GEN) then GS.ElipsisHint:=GetMessageRC(RC_SELCOMPTE)
    else if (lc=SF_AUX) then GS.ElipsisHint:=GetMessageRC(RC_SELTIERS) ;
// Guide
if (bGuideRun) and (not Cancel) then
 BEGIN
  // NumCompte:=GS.Cells[SF_GEN, GS.Row] ; k:=Comptes.GetCompte(NumCompte) ;
 //  if (lColStop<>SF_GEN) and ((k<0) or (NumCompte<>GS.Cells[SF_GEN, GS.Row])) then
  // BEGIN ACol:=SF_GEN ; Cancel:=TRUE ; Exit ; END ;
   if GetGuideValue(lc, lr) then
     if (bGuideStop) then
       begin
        lColStop:=GetGuideColStop(lr) ;
        NumCompte:=GS.Cells[SF_GEN, GS.Row] ; k:=Comptes.GetCompte(NumCompte) ;
        if (lColStop<>SF_GEN) and ((k<0) or (NumCompte<>GS.Cells[SF_GEN, GS.Row])) then
        BEGIN ACol:=SF_GEN ; Cancel:=TRUE ; Exit ; END ;
        Cancel:=TRUE ; ACol:=lColStop ;
        bGuideStop:=FALSE ; GS.ElipsisButton:=FALSE ;
       end else
       begin
       if sens=1 then PostMessage(GS.Handle, WM_KEYDOWN, VK_TAB, 0)
                 else PostMessage(GS.Handle, WM_KEYDOWN, VK_LEFT, 0) ;
       end ;
 END ;
finally
 HGEndUpdate(GS) ;
end;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 13/09/2004
Modifié le ... : 16/04/2007
Description .. : - LG - 13/09/2004 - FB 10810 - on ne tient plus compte du
Suite ........ : compte courant pour calculer le solde
Suite ........ : - LG - 16/04/2007 - FB 16744 - on affiche les ecritures de 
Suite ........ : type L pour calculer corectment le solde du compte
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.SoldeAux ( D, C : double ) ;
var
 lQ : TQuery ;
begin

 {$IFDEF TT}
  AddEvenement('TFSaisieFolio.SoldeAux') ;
 {$ENDIF}

  if not FBoFaireCumul then exit ;

  if (Tiers.GetValue('_SOLDED') = - 1) and (Tiers.GetValue('_SOLDEC') = - 1) then
   begin
   lQ:=OpenSQL('SELECT SUM(E_DEBIT), SUM(E_CREDIT) FROM ECRITURE'
              +' WHERE E_AUXILIAIRE="'+GS.Cells[SF_AUX, GS.Row]+'"'
              +' AND E_EXERCICE="'+memOptions.Exo+'"'
              +' AND E_DATECOMPTABLE>="'+USDateTime(ctxExercice.CPExoRef.Deb)+'"'
              +' AND E_DATECOMPTABLE<="'+USDateTime(EncodeDate(memOptions.Year, memOptions.Month, memOptions.MaxJour))+'"'
              +' AND ( E_QUALIFPIECE="N" OR E_QUALIFPIECE="C" OR E_QUALIFPIECE="L" ) ' ,
              TRUE) ;
    Tiers.PutValue('_SOLDED', lQ.Fields[0].AsFloat ) ;
    Tiers.PutValue('_SOLDEC', lQ.Fields[1].AsFloat ) ;
    Ferme(lQ) ;
   end ;
 AfficheLeSolde(SA_SoldeAux, Tiers.GetValue('_SOLDED')+D, Tiers.GetValue('_SOLDEC')+C) ;
 SoldeAuxPer.Visible:=true ;
 SoldeAuxPer.caption := '(M) ' + SA_SoldeAux.Text ;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 23/03/2006
Modifié le ... :   /  /    
Description .. : - LG - FB 17681 -  23/03/2006 - la fenetre des echeances 
Suite ........ : ne s'ouvrait plus
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.FiniAux(ACol, ARow : integer ) ;
var
 NumAux : string ;
 kt : integer ;
 Ecr : TZF ;
begin
Ecr:=ObjFolio.GetRow(ARow-GS.FixedRows) ;
if (GS.Cells[SF_Aux, ARow]<>'') and (Ecr<>nil) and (Ecr.GetValue('E_AUXILIAIRE')<>GS.Cells[SF_Aux, ARow]) then
 begin
  Ecr.PutValue('E_MODEPAIE',     '') ;
  Ecr.PutValue('E_DATEECHEANCE', EncodeDate(1900, 1, 1)) ;
  Ecr.PutValue('E_RIB',          '') ;
  Ecr.PutValue('E_AUXILIAIRE',   GS.Cells[ACol, ARow]) ;
  NumAux:=GS.Cells[SF_AUX, ARow] ; kt:=Tiers.GetCompte(NumAux) ;
  if (kt>-1) and (Tiers.GetValue('T_CONFIDENTIEL')>Ecr.getValue('E_CONFIDENTIEL')) then
   Ecr.PutValue('E_CONFIDENTIEL',Tiers.GetValue('T_CONFIDENTIEL' ) ) ;
  if ( GS.Cells[SF_Debit, ARow] <> '' ) or ( GS.Cells[SF_Credit, ARow] <> '' ) then
   GetEch(ARow, FALSE, TRUE) ;
  if not PutRib(ACol, ARow, TRUE,true) then PrintMessageRC(RC_RIB) ;
 end
  else
  // Maj des infos provenant du compte auxiliaire
  if not PutRib(ACol, ARow, TRUE) then PrintMessageRC(RC_RIB) ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 25/07/2002
Modifié le ... : 06/01/2005
Description .. : - 25/07/2002 - correction du contrôle de gestion de nature
Suite ........ : et chargement des scenario correspondant
Suite ........ : -06/11/2002 - correction de l'affichage du message nature
Suite ........ : incorrecte
Suite ........ : - 03/12/2002 - fiche 10821 - on peut saisir dasn la zone
Suite ........ : nature
Suite ........ : - 05/12/2002 - on peut saisir dans la zone nature
Suite ........ : uniquement pour les natres abreges
Suite ........ : - si la ligne n'est pas encore cree, on sort directement. Dans
Suite ........ : un folio existant, on choisit un nouveau numero puis on
Suite ........ : clique dans la grille-> on ne fait pas de test sur le cellexit
Suite ........ : - 31/01/2003 - FB 11850 - test de la nature sur la premiere
Suite ........ : ligne
Suite ........ : - 03/06/2003 - FB 12457 - le rib n'etait pas affecter qd on
Suite ........ : saisisait l'auxi ds la case du general
Suite ........ : - 11/09/2003 - FB 12621 - calcul dans les zones débit et
Suite ........ : crédit en direct
Suite ........ : - on ne pouvait plus saisir de montant negatif juste a l'ajout
Suite ........ : de la gestion des calcul ds les cases debit/credit
Suite ........ : - LG -02/08/2004 - FB 12377 - ajout de la suppression de
Suite ........ : l'anaytique pour els compte collectifs ( le test sur le Cancle
Suite ........ : est mis a la fin du traitement des comptes )
Suite ........ : - 16/08/2004 - LG - FB 13056 - l'affectation du rib ne se
Suite ........ : faisait pas qd on cheangant de rib
Suite ........ : - 06/01/2005 - LG  - FB14776 - correction pour la maj des
Suite ........ : totaux du folio
Mots clefs ... :
*****************************************************************}
procedure TFSaisieFolio.GSCellExit(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
var Ecr : TZF ; NumCompte : string ; j,i : Integer ; RefRow : LongInt ;
lStCaseNat : string;   lStMess : string ; lStNat : string ;
lIndexAux,lIndexGen : integer ;
lSens : integer ;
lbOldModeRO ,lChg : boolean ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 02/09/2004
Modifié le ... : 09/07/2007
Description .. : - 02/09/2004 - LG - FB 14247 - l'appel des scenario ne fct
Suite ........ : pas pour un seul etablissement
Suite ........ : - 09/09/2004 - LG - suppression de l'appel des devise a la
Suite ........ : sortie du jour, ouvrait deux fois la saisie en devise
Suite ........ : - 15/10/2004 - LG - FB 14779 - le solde du journal n'etait
Suite ........ : pas reculculer qd on passait de debit a credit
Suite ........ : - 20/10/2004 - LG FB 14825 - corection sur le
Suite ........ : declenchement du calcul du solde
Suite ........ : - 20/07/2005 - LG - on controle la nature des que l'on sort
Suite ........ : de la case
Suite ........ : - LG - 20/10/2005 -  FB 12630 - ouvreture direct du lettrage
Suite ........ : - LG - 14/12/2005 - FB 17191 - affection de la tob ecriture 
Suite ........ : avant de lancer la saisie des qte
Suite ........ : - LG - 09/07/2007 - FB 20928 - on ne peut plus modifier el 
Suite ........ : jour qd une ligne de la piece est lettree
Mots clefs ... : 
*****************************************************************}
begin
if csDestroying in ComponentState then Exit ;
if GS.Cells[SF_NEW,ARow]='' then exit ;
// En consultation, rien à faire ici !
if bModeRO then Exit ;
if Statut=taConsult then
 begin
  lbOldModeRO := bModeRO ;
    try
     if ( (ACol=SF_DEBIT) or (ACol=SF_CREDIT) ) and (Valeur(GS.Cells[ACol, ARow])<>0) then
      begin
       bModeRO := true ;
       InfosPied(ARow) ;
       if PQuantite.Visible and _QTE1.CanFocus then
        begin
         Cancel := true ;
         _QTE1.SetFocus ;
        end ;
      end ;
    finally
     bModeRO:= lbOldModeRO ;
    end ;
  exit ;
 end ;
{$IFDEF TT}
//AddEvenement('GSCellExit ACol:='+intToStr(ACol)+' ARow:='+intToStr(ARow)) ;
{$ENDIF}
HGBeginUpdate(GS) ;
try
lStCaseNat:=GetCellNature(ARow) ; // recuperation de la nature suivant le type d'affichage choisit
// Nature de pièce
if (ACol=SF_NAT) then
   BEGIN
   if ( AnsiUpperCase(GS.Cells[ACol, ARow]) = '' ) then
    begin
     Cancel:=true ;
     PostMessage(GS.Handle , WM_KEYDOWN , VK_F5 , 0) ;
     exit ;
    end;

   if FBoNatureComplete then
    i:=E_NATUREPIECE.Items.IndexOf(GS.Cells[SF_NAT, ARow])
     else
      i:=E_NATUREPIECE.Values.IndexOf(GS.Cells[SF_NAT, ARow]) ;

   if ( i = -1 ) then
    begin
     Cancel:=true ;
     PostMessage(GS.Handle , WM_KEYDOWN , VK_F5 , 0) ;
     exit ;
    end;

   if not FBoNatureComplete then GS.Cells[ACol, ARow]:=AnsiUpperCase(GS.Cells[ACol, ARow]) ;
   if (not memParams.bLibre) and (GS.Cells[SF_PIECE, ARow]<>'') then
     begin
     RefRow:=VerifNature(StrToInt(GS.Cells[SF_PIECE, ARow]), ARow,FBoValEnCours) ;
     if RefRow>0 then
       begin
       lStMess:=GetMessageRC(RC_CNATURE2) ;
       PGIError(Format(lStMess, [GS.Cells[SF_GEN, RefRow], RefRow]) , Caption) ;
       Cancel:=TRUE ; Exit ;
       end ;
     SetColVal(SF_NAT, GS.Cells[SF_PIECE, ARow], GS.Cells[ACol, ARow]) ;
     end ; //else SetColVal(SF_NAT, GS.Cells[SF_PIECE, ARow], GS.Cells[ACol, ARow]) ;
   // fiche 10821 - controle de la nature

   if not FBoNatureComplete and ( AnsiUpperCase(RechDom(E_NATUREPIECE.DataType,GS.Cells[ACol, ARow],false)) = '' ) then
    begin
     Cancel:=true ;
     PostMessage(GS.Handle , WM_KEYDOWN , VK_F5 , 0) ;
     exit ;
    end;
   // Changement de nature ?
   if ObjFolio<>nil then
     begin
     if ARow>1 then lStNat:=GS.Cells[ACol, ARow-1] ;
     if (lStNat<>GS.Cells[ACol, ARow]) or (ARow=1) then // pour la premiere ligne on la nature est forcement diff de la ligne precedente !
        begin
        if memParams.bLibre then
        ModifRefAna(ARow-GS.FixedRows, 'NATUREPIECE',  lStCaseNat)
        else ModifRefAnaPiece(ARow-GS.FixedRows, 'NATUREPIECE',  lStCaseNat , GS.Cells[SF_PIECE, ARow]) ;
        FZScenario.Load([MemJal.Code,lStCaseNat,'N',E_ETABLISSEMENT.Value]);
        EnableButtons ; // rafraichissement des button en fct de la nouvelle nature
        end ;
     end ;
   END ;
// Validité du jour
if (ACol=SF_JOUR) then
   begin
   if bNoEvent then begin bNoEvent:=false ; exit ; end ;
   GS.Cells[ACol, ARow]:=StrFMontant(Abs(Valeur(GS.Cells[ACol, ARow])),15,0,'',TRUE) ;
   if (Valeur(GS.Cells[ACol, ARow])>memOptions.MaxJour) or (Valeur(GS.Cells[ACol, ARow])<=0) then
      GS.Cells[ACol, ARow]:=IntToStr(memOptions.MaxJour) ;
   if not CControleDateBor(GetRowDate(ARow),ctxExercice,false,memOptions.Exo) then
    begin
     Cancel := true ;
     exit ;
    end ;
   if RevisionActive(GetRowDate(ARow)) then begin Cancel:=TRUE ; Exit ; end ;
   Ecr:=ObjFolio.GetRow(ARow-GS.FixedRows) ;
   if ( not IsPieceCanEdit(GS.Cells[SF_PIECE, ARow]) ) and
      (Ecr<>nil) and ( Ecr.GetValue('e_datecomptable')<>GetRowDate(ARow) ) then
    begin
     LabelInfos2.Caption:=GetMessageRC(RC_PASTOUCHEDATE) ;
     LabelInfos2.Visible:=TRUE ;
     Cancel := true ;
     Exit ;
    end ;
   if not memParams.bLibre then SetColVal(SF_JOUR, GS.Cells[SF_PIECE, ARow], GS.Cells[ACol, ARow]) ;
   // Changement de jour ?
   if (Ecr<>nil) and (Ecr.GetValue('E_AUXILIAIRE')<>'') and
      (Ecr.GetValue('E_DATECOMPTABLE')<>GetRowDate(ARow)) then
        begin
        Ecr.PutValue('E_MODEPAIE',     '') ;
        Ecr.PutValue('E_DATEECHEANCE', EncodeDate(1900, 1, 1)) ;
        Ecr.PutValue('E_RIB',          '') ;
        Ecr.PutValue('E_AUXILIAIRE',   GS.Cells[ACol, ARow]) ;
        GetEch(ARow, FALSE, TRUE) ;
        end ;
   if (Ecr<>nil) and (Ecr.GetValue('E_DATECOMPTABLE')<>GetRowDate(ARow)) then
        begin
         Ecr.PutValue('COMPTAG','-') ;
         SupprimeEcrCompl(Ecr) ;
         if memParams.bLibre
          then      ModifRefAna(ARow-GS.FixedRows, 'DATECOMPTABLE',  GetRowDate(ARow))
          else ModifRefAnaPiece(ARow-GS.FixedRows, 'DATECOMPTABLE',  GetRowDate(ARow), GS.Cells[SF_PIECE, ARow]) ;
        end ;
//   if (GS.Cells[SF_DEVISE, ARow]<>'') then DevClick;
   end ;
// Validité du compte général
if (ACol=SF_GEN) then
   begin
    //FRechCompte.StNature:=lStCaseNat ; FRechCompte.StJournal:=memJal.Code ; FRechCompte.StDevise:=FInfo.Devise.Dev.Code ;
    if (GS.Cells[ACol, ARow]='') and (GS.Cells[SF_NEW, ARow]='O') then
      GS.Cells[SF_AUX, ARow]:='' // sur une nouvelle ligne avec un compte vide on peut sortir directement ( cas quand on fait fleche haut sur la derniere ligne
      else
       begin
        lSens:= CGetGridSens(GS,Acol,ARow ) ;
        SetRowTZF(ARow,false) ;
        Ecr:=ObjFolio.GetRow(ARow-GS.FixedRows) ;
        FRechCompte.Ecr := Ecr ;
        FRechCompte.CellExitGen(GS,ACol,ARow,Cancel) ;
        if cancel and ( lsens=-1) then
         begin
          cancel := false ;
          GS.Cells[SF_GEN, ARow]:='' ;
          exit ;
         end ;
       end;

    // controle du compte visé
    if ( not Cancel ) and (GS.Cells[SF_NEW, ARow]='O') then
     Cancel:= not CControleVisa(GS.Cells[ACol, ARow],FRechCompte.Info) ;
    // Maj des infos provenant du compte général
    MajInfoCompte ( ARow,ACol ) ;

     if trim(GS.Cells[SF_AUX, ARow])<>'' then
      FiniAux(ACol,ARow) ;

    if Cancel then exit ;
    //SoldeDuCompte ;
   end ; // Validité du compte général
// Validité du compte auxiliaire
if (ACol=SF_AUX) then
   begin
    //FRechCompte.StNature:=lStCaseNat ; FRechCompte.StJournal:=memJal.Code ; FRechCompte.StDevise:=FInfo.Devise.Dev.Code ;
    SetRowTZF(ARow,false) ;
    Ecr:=ObjFolio.GetRow(ARow-GS.FixedRows) ;
    FRechCompte.Ecr := Ecr ;
    FRechCompte.CellExitAux(GS,ACol,ARow,Cancel) ; lIndexAux:=FRechCompte.Info.Aux.InIndex ;

    if Cancel or (GS.Cells[SF_Aux, ARow]='') then
     begin
      GS.ElipsisButton := false ;
      exit ;
     end ;
   // Test MonoDevise du tiers
   if (GS.Cells[SF_DEVISE, ARow]<>'') and (GS.Cells[SF_DEVISE, ARow]<>V_PGI.DevisePivot) then
     if (lIndexAux>=0) and (Tiers.GetValue('T_MULTIDEVISE', lIndexAux)='-') and (Tiers.GetValue('T_DEVISE', lIndexAux)<>GS.Cells[SF_DEVISE, ARow]) then
       begin
       PrintMessageRC(RC_TMONODEVISE) ; GS.Cells[SF_AUX, ARow]:='' ;
       Cancel:=TRUE ; Exit ;
       end ;
   // Changement de compte ?
   Ecr:=ObjFolio.GetRow(ARow-GS.FixedRows) ;
   FiniAux(ACol,ARow) ;
   PutRegime(ARow) ; InfosPied(ARow,false) ;
   if (Ecr<>nil) and (Ecr.GetValue('E_CONSO')='') then PutConso(ARow) ;
   if not bGuideRun then EnableButtons ;
 end ;
// Utiliser le libellé du compte si pas de libellé
if (ACol=SF_LIB) and (GS.Cells[ACol, ARow]='') and (CGetGridSens(GS,ACol, ARow)=1) then
   begin
   NumCompte:=GS.Cells[SF_AUX, ARow] ; lIndexAux:=Tiers.GetCompte(NumCompte) ;
   if  lIndexAux>=0 then
     GS.Cells[ACol, ARow]:=Tiers.GetValue('T_LIBELLE', lIndexAux)
   else begin
     NumCompte:=GS.Cells[SF_GEN, ARow] ; lIndexGen:=Comptes.GetCompte(NumCompte) ;
     if (GS.Cells[ACol, ARow]='') and (lIndexGen>=0) then
       GS.Cells[ACol, ARow]:=Comptes.GetValue('G_LIBELLE', lIndexGen) ;
     end ;
   end ;
// Formatage du montant
if (ACol=SF_DEBIT) or (ACol=SF_CREDIT) then
   begin
   if (pos('+',GS.Cells[ACol, ARow])>0) or (pos('-',GS.Cells[ACol, ARow])>0) or (pos('/',GS.Cells[ACol, ARow])>0) or (pos('*',GS.Cells[ACol, ARow])>0) then
   begin
   //if (not VH^.MontantNegatif) and ( GS.Cells[ACol, ARow][1] in ['+','-','/','*'] ) then
   if (not VH^.MontantNegatif) and ( pos( Copy(GS.Cells[ACol, ARow], 1, 1), '+;-;/;*' ) > 0 ) then
     begin
       if ACol = SF_DEBIT then
         GS.Cells[ACol, ARow]:=GFormule('{[E_DEBIT]'+GS.Cells[ACol, ARow]+'}', GetFormulePourCalc, nil, 1)
       else GS.Cells[ACol, ARow]:=GFormule('{[E_CREDIT]'+GS.Cells[ACol, ARow]+'}', GetFormulePourCalc, nil, 1);
     end else
     { sinon on effectue l'opération simplement }
      begin
       if Pos(GS.Cells[ACol, ARow][1],'/')>-1 then
        GS.Cells[ACol, ARow]:='{"#.###,0"'+GS.Cells[ACol, ARow]+'}';
       GS.Cells[ACol, ARow]:=GFormule('{'+GS.Cells[ACol, ARow]+'}', GetFormulePourCalc, nil, 1)
      end ;
   end;
   FormatMontant(GS, ACol, ARow, FInfo.Devise.Dev.Decimale) ;
   if (not VH^.MontantNegatif) and (Valeur(GS.Cells[ACol, ARow])<0) then
     begin
     PrintMessageRC(RC_NONEGATIF) ; GS.Cells[ACol, ARow]:='' ;
     Cancel:=TRUE ; Exit ;
     end ;
   if ((Valeur(GS.Cells[ACol, ARow])<>0) and
      (Valeur(GS.Cells[ACol, ARow])<VH^.GrpMontantMin) or
      (Valeur(GS.Cells[ACol, ARow])>VH^.GrpMontantMax)) then
     begin
     PrintMessageRC(RC_NOGRPMONTANT) ; GS.Cells[ACol, ARow]:='' ;
     Cancel:=TRUE ; Exit ;
     end ;
   Ecr:=ObjFolio.GetRow(ARow-GS.FixedRows) ;
   if (Ecr<>nil) and ( (FInRowCur<>ARow) or
                       ( (FInRowCur=ARow) and (Valeur(GS.Cells[ACol, ARow])<>FRdValCur) ) ) then
    begin
     FInRowCur := ARow ;
     FRdValCur := Valeur(GS.Cells[ACol, ARow]) ;
     CalculSoldeJournal ;
    end ; // if
   if not bGuideRun then EnableButtons ;
   // Changement de compte ?
   Ecr:=ObjFolio.GetRow(ARow-GS.FixedRows) ;
   if (Ecr<>nil) and (Ecr.GetValue('E_ANA')='X') then
     begin
         if Ecr.GetValue('E_DEVISE')<>memPivot.Code then
         begin
         if (Ecr.GetValue('E_DEBITDEV') <>0) and (ACol=SF_CREDIT) then ObjFolio.DelAna(ARow-GS.FixedRows) ;
         if (Ecr.GetValue('E_CREDITDEV')<>0) and (ACol=SF_DEBIT)  then ObjFolio.DelAna(ARow-GS.FixedRows) ;
         end else
         begin
         if (Ecr.GetValue('E_DEBIT') <>0) and (ACol=SF_CREDIT) then ObjFolio.DelAna(ARow-GS.FixedRows) ;
         if (Ecr.GetValue('E_CREDIT')<>0) and (ACol=SF_DEBIT)  then ObjFolio.DelAna(ARow-GS.FixedRows) ;
         end ;
     end ;

    lbOldModeRO := bModeRO ;
    try
     if  (Valeur(GS.Cells[ACol, ARow])<>0) then
      begin
       bModeRO := true ;
       InfosPied(ARow) ;
       bModeRO:= lbOldModeRO ;
       if PQuantite.Visible and _QTE1.CanFocus and not FBoCutEnCours then
        begin
        {$IFDEF TT}
         AddEvenement('_QTE1.SetFocus'+intToStr(ACol)+' ARow:='+intToStr(ARow)) ;
        {$ENDIF}
         SetRowTZF(ARow,false) ;  lChg := true ;
         GSRowExit(nil,ARow,Cancel,lchg) ; // appel de l'analytiq
         GS.OnExit := nil ; // l'evement sera rebranche ds l'evenemetn PQuantite.Exit
         SetGridOn(FALSE) ;
         Cancel := true ;
         if _QTE1.CanFocus then
          _QTE1.SetFocus ;
        end ;
       {  else
          if memJal.Lettrage then
           begin
            EnableButtons(Arow) ;
            LettrageEnSaisie(ARow) ;
           end ;  }
      end ;
    finally
     bModeRO:= lbOldModeRO ;
    end ;

   end ;
if (ACol=SF_LIB) then
   begin
   // Changement de libellé ?
   Ecr:=ObjFolio.GetRow(ARow-GS.FixedRows) ;
   if (Ecr<>nil) and (Ecr.GetValue('E_ANA')='X') and
      (Ecr.GetValue('E_LIBELLE')<>GS.Cells[ACol, ARow]) then
        begin
        for i:=0 to Ecr.Detail.Count-1 do
          for j:=0 to (Ecr.Detail[i].Detail).Count-1 do
            (Ecr.Detail[i].Detail[j]).PutValue('Y_LIBELLE', GS.Cells[ACol, ARow]) ;
        end ;
   end ;
if (ACol=SF_REFI) then
   begin
   // Changement de libellé ?
   GS.ElipsisButton := false ;
   if trim(GS.Cells[ACol, ARow])<>'' then FLastRef := GS.Cells[ACol, ARow]
    else
      if bNewLigne and memJal.IncRef  then _DecS(FLastRef) ;

   Ecr:=ObjFolio.GetRow(ARow-GS.FixedRows) ;
   if (Ecr<>nil) and (Ecr.GetValue('E_ANA')='X') and
      (Ecr.GetValue('E_REFINTERNE')<>GS.Cells[ACol, ARow]) then
        begin
        for i:=0 to Ecr.Detail.Count-1 do
          for j:=0 to (Ecr.Detail[i].Detail).Count-1 do
            (Ecr.Detail[i].Detail[j]).PutValue('Y_REFINTERNE', GS.Cells[ACol, ARow]) ;
        end ;
   end ;
if (bGuideRun) and (not Cancel) and (not bGuideStop) then
 SetRowTZF(ARow,false) ;
finally
 HGEndUpdate(GS) ;
end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 20/12/2002
Modifié le ... : 11/08/2004
Description .. : - 20/12/2002 - on redeclencher des cellexits sur toutes les
Suite ........ : lignes pour les guides n'ayant pas de point d'arret
Suite ........ :
Suite ........ : - 28/03/2003 - LG - FB12208 - la fenetres des scenarios
Suite ........ : s'ouvrait pour toutes les lignes. On part du debit de guide
Suite ........ : - 31/03/2003 - LG - on bloque la generation des lignes
Suite ........ : d'ecart de conversion lors de la validation du guide, fait une 
Suite ........ : fois a la fin
Suite ........ : - FB 14100 - controle de tout le guide qd on a fins ou 
Suite ........ : abandonner celui ci
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.FinGuide ;
Var lCancel : boolean ;
    i : integer ;
    lTError : TRecError ;
BEGIN
{$IFDEF TT}
AddEvenement('FinGuide') ;
{$ENDIF}
if GuideFirstRow = - 1 then exit ;
FBoPasEcart:=true ;  HGBeginUpdate(GS) ;
lTError.RC_Error := RC_PASERREUR ; lTError.RC_Message := '' ;
i := GuideFirstRow ;
try
while i <= nbLignes  do
 begin
  lCancel:=False ; Statut:=taModif ;
  GSRowExit(Nil,i,lCancel,True) ;
  if lCancel then
    DeleteRow(i)
     else
      begin
       GetRowTZF(i-GS.FixedRows) ;
       inc(i) ;
      end ;
 end ;
finally
 FBoPasEcart:=false;
 HGEndUpdate(GS) ;
end;
END ;



{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 12/06/2002
Modifié le ... : 24/05/2007
Description .. : - 14/06/2002 - ajout de la gestion du statut
Suite ........ : - 08/10/2002 - gestion de la suppression de la derniere
Suite ........ : ligne
Suite ........ : - la saisie de l'analytique, info compl. ne ce fait pas sur les
Suite ........ : lignes sans montants
Suite ........ : -28/10/2002 - le message d'erreur ne s'affichaient pas
Suite ........ : -30/01/2003 - FB 11853 - on pouvait sorti d'un folio non
Suite ........ : equilibre
Suite ........ : - 07/05/2003 - FB 12167 - gestion des deplacements avec 
Suite ........ : la souris
Suite ........ : - 02/08/2004 - FB 13772 - correction du test sur 
Suite ........ : l'accelerateur de saisie, la variable ecr n'etait pas toujours
Suite ........ : definie.Permettait d'enregsitre un bordereau sans compte
Suite ........ : - 05/08/2004 - FB 13913 13914 - on controle l'integralite de
Suite ........ : laigne a la sortide de celle ci ( pb de validation qd on se
Suite ........ : deplca avec la souris, ou validation direct d'une ligne
Suite ........ : incorrecte )
Suite ........ : - 18/07/2005 - FB 12301 - en mode bordereau on equilibre
Suite ........ : les devise sur la premiere ligne
Mots clefs ... :
*************************~~****************************************}
procedure TFSaisieFolio.GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var
 RefRow : LongInt ;
 sens : integer ;
 Ecr : TOB ;
// NumAux : string ;
// kt : integer ;
 lTError : TRecError ;
 lInIndex : integer ;
 {$IFDEF LAURENT}
 lw : CWA ;
 {$ENDIF}
begin
{$IFDEF TT}
AddEvenement('GSRowExit Ou:='+intToStr(Ou)) ;
{$ENDIF}
if (csDestroying in ComponentState) then exit ; if (Chg) and (not (GoEditing in GS.Options)) then Exit ;
if FBoValEnCours then exit ;
//if not (Screen.ActiveControl=GS) then exit ;
// En consultation, rien à faire ici !
if bModeRO then
 begin
  EnableButtons ;
  Exit ;
 end;
if GS.Cells[SF_NEW,Ou]='' then exit ;
//if FboFenLett then exit ;
HGBeginUpdate(GS) ;
sens:=CGetGridSens(GS,GS.Col, Ou) ;

try

bStopDevise:=FALSE ;

if (Sens=1) and (( GS.Row <> Ou ) )  then
 begin // on descend
  if (GS.Cells[SF_GEN,Ou]='') then
   begin
    cancel:=true ;
    exit ;
   end;
 end
  else
   begin // on remonte
    if (GS.Cells[SF_GEN,Ou]='') then
     begin  // pas de generaux

      if (Ou=nbLignes) and (Screen.ActiveControl=GS)  and// on test le controle actif cas du click de souris dans la grille il ne faut pas supprime le nouvelle cellule
      (GS.Cells[SF_PIECE,GS.Row]=GS.Cells[SF_PIECE, Ou]) then
       begin  // sur une nouvelle ligne ou si l'on est à la fin du mode piece -> supprime la ligne
        deleteRow(Ou) ;
        exit ;
       end
        else
         if ((bGuideRun) or (not memParams.bLibre)) and
         (GS.Cells[SF_PIECE,GS.Row]<>GS.Cells[SF_PIECE, Ou]) and
         (GS.Cells[SF_PIECE, Ou]<>'') and
         (not IsSoldePiece(StrToInt(GS.Cells[SF_PIECE, Ou]))) then
         BEGIN // pour le mode bordereau : la piece n'est pas solde on reste dans celle ci. On rafraichi la ligne courante
          Cancel:=TRUE ; Exit ;
         END ;

        if (not (bGuideRun) and (not memParams.bLibre)) and
         (GS.Cells[SF_PIECE,GS.Row]<>GS.Cells[SF_PIECE, Ou]) and
         (GS.Cells[SF_PIECE, Ou]<>'') and
         ( IsSoldePiece(StrToInt(GS.Cells[SF_PIECE, Ou]))) then
         BEGIN // on a changer de piece avec la souris, mais la piece de depart est vide => on supprime la ligne( il n'y en a qu'une )
          deleteRow(Ou) ; Exit ;
         END ;

        if (bGuideRun) and (GS.Cells[SF_PIECE,GS.Row]<>GS.Cells[SF_PIECE, Ou]) then
         BEGIN // en mode guide on ne peut pas changer de piece
          Cancel:=TRUE ; Exit ;
         END ;

     end // pas de generaux
     else
      if (Valeur(GS.Cells[SF_DEBIT,Ou])=0) and (Valeur(GS.Cells[SF_CREDIT,Ou])=0) then
       begin // pas de debit credit saisie
        if bNewLigne or ( (Ou=nbLignes) and MemParams.bLibre ) then
         begin  // sur une nouvelle piece ou la dernier ligne d'un folio libre on supprime
          DeleteRow(Ou) ;
          exit ;
         end;
       end;
   end; // on remonte

  if Statut=taConsult then exit ;
  // sauvegarde de la reference pour la resortir
  if GS.Cells[SF_REFI,GS.Row]<>'' then FLastRef:=GS.Cells[SF_REFI,GS.Row] ;
  SetRowLib(Ou) ;
  Ecr:=SetRowTZF(Ou,false) ;
  if ((Valeur(GS.Cells[SF_DEBIT,Ou])<>0) xor (Valeur(GS.Cells[SF_CREDIT,Ou])<>0)) then
   BEGIN // fct non accesible sur les lignes sans montant
    if not memJal.Lettrage then GereComplements(Ou) ;
    PutRib(SF_Aux, Ou, TRUE) ; // LG 11/07/2005 - on force l'affectation du rib en sorti de ligne
    if (VH^.ZSAISIEECHE) then GetEch(Ou, TRUE, FALSE) else GetEch(Ou, FALSE, TRUE) ;
    GereAna(Ou,false) ;

    // Dev 3946 : control cohérence qté analytique soumis à indicateur
    if (not (ctxPCL in V_PGI.PGIContexte)) and GetParamSocSecur('SO_ZCTRLQTE', False) then
        if ( Ecr <> nil ) and ( Ecr.GetValue('E_ANA')='X' )
                          and ( ( Ecr.GetNumChamp('CHECKQTE') < 0 )  or
                                ( Ecr.GetString('CHECKQTE')<>'X'  ) )
          then CheckQuantite( Ecr ) ;
    // Fin Dev 3946
      if bUp then SendMessage(GS.Handle, WM_LBUTTONUP, 0, 0) ;
    PutRowProrata(Ou) ;
  //  PutRowImmo(Ou) ;
   END ;
  if (not memParams.bLibre) and (GS.Cells[SF_PIECE, Ou]<>'') and (IsSoldePiece(StrToInt(GS.Cells[SF_PIECE, Ou]))) then
   begin
     // Vérification du compte de contrepartie
     if not VerifTreso(StrToInt(GS.Cells[SF_PIECE, Ou]), Ou) then
       begin
        PrintMessageRC(RC_TRESO) ; GS.Col:=SF_GEN ; Cancel:=TRUE ; Exit ;
       end ;
     RefRow:=VerifNature(StrToInt(GS.Cells[SF_PIECE, Ou]), Ou) ;
     if (RefRow>0) and (GS.Col<>SF_GEN) then
       begin
        PGIError(Format(GetMessageRC(RC_CNATURE2), [GS.Cells[SF_GEN, RefRow], RefRow]), Caption) ;
        if (GS.Cells[SF_PIECE, Ou]<>GS.Cells[SF_PIECE, GS.Row]) then begin Cancel:=TRUE ; Exit ; end ;
       end ;
     // Vérification équilibre
     if FFirstRow = 0 then FFirstRow := Ou ;
     if not FBoPasEcart then VerifEquilibre(StrToInt(GS.Cells[SF_PIECE, Ou]), FFirstRow) ;
   end ;

 lTError.RC_Error := RC_PASERREUR ; lTError.RC_Message := '' ;
  if Ou <= nbLignes then
   begin
    {$IFDEF laurent}
     lW := cWA.Create ; //Showmessage( lw.aServer.SessionID ) ;
     LT := TOB.Create('$PARAM', nil, -1) ;
     LT.AddChampSupValeur('USERLOGIN' , V_PGI.UserLogin ) ;
     LT.AddChampSupValeur('INIFILE'   , HalSocIni ) ;
     LT.AddChampSupValeur('PASSWORD'  , V_PGI.Password ) ;
     LT.AddChampSupValeur('DOMAINNAME', '' ) ;
     LT.AddChampSupValeur('DATEENTREE', V_PGI.DateEntree ) ;
     LT.AddChampSupValeur('DOSSIER'   , V_PGI.NoDossier ) ;
     LT.AddChampSupValeur('BaseCommune', false);
     LT.AddChampSupValeur('ERROR', '');
     LT.AddChampSupValeur('APPLICATION', 'COMPTA');
     le := TOB.Create('ECRITURE',lt,-1) ;
     le.dupliquer(ecr,true,true) ;
     lw.Request('PGIComptaAPI.VALID','',lt,'','') ;
     lw.free ;
    {$ELSE}
    lTError := CIsValidLigneSaisie(Ecr ,FRechCompte.Info) ;
    {$ENDIF}
    if (lTError.RC_Error <> RC_PASERREUR) and (lTError.RC_Error <> RC_MONTANTINEXISTANT) and
    (lTError.RC_Error <> RC_CFERME) and (lTError.RC_Error <> RC_COMPTEVISA) then
     begin
      // FB 19575
      if (GuideFirstRow = - 1) and (GS.Row<>1) then OnError(nil,lTError) ; // indique que l'on est encore en saisie guide bGuideRun est daje a false ( voir fct GuideStop )
      V_PGI.IoError := oeSaisie ;
      Cancel        := true ;
      exit ;
     end ;
   end ;

 // LG 19/07/2005 - LG - on cree l'immo qd la ligne est correcte
 if ((Valeur(GS.Cells[SF_DEBIT,Ou])<>0) xor (Valeur(GS.Cells[SF_CREDIT,Ou])<>0)) then
  PutRowImmo(Ou) ;

 if EstAccelere(Ou) then
  begin
   BAcc.Enabled := true ;
   lInIndex := AccelerateurDeSaisie(Ou) ;
   if lInIndex <> -1 then
    GS.Row := lInIndex ;
  end
   else
    if ( not memJal.Accelerateur ) and ( memJal.Lettrage ) and ( Ou = nbLignes ) then
     begin
      EnableButtons(Ou) ;
      LettrageEnSaisie(Ou,true) ;
     end ;

 GereCutOff(Ou) ;


finally
 HGEndUpdate(GS) ;
end;

end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 12/06/2002
Modifié le ... : 11/06/2003
Description .. : - 12/06/2002 - suppression de fonction inutilise
Suite ........ : - 13/06/2002 - suppression de la fonction infopieds :
Suite ........ : affichage des info sur le comptes
Suite ........ : -20/09/2002 - les labels d'info sur le compte ne sont plus
Suite ........ : visible quand on change de ligne
Suite ........ : -26/09/2002- on test s'il existe un objet pour cette position
Suite ........ : de la grille. Permet de bloquer les deplacement à la souris
Suite ........ : -22/10/2002 - on consultation ( depuis le journal
Suite ........ : centralisateur) on ne fait rien ici
Suite ........ : -06/11/2002 - le test precedent ne marchait pas
Suite ........ : -28/11/2002- fiche 10766 - en mode guide en saisie libre on
Suite ........ : repesait en 
Suite ........ : consultation
Suite ........ : - 10/12/2002 - fiche 10938 - on réaffiche les libelles du
Suite ........ : compte et du tiers
Suite ........ : -20/12/2002 - pour les guides sans point d'arret on 
Suite ........ : redeclencher des cellexits sur toutes les lignes
Suite ........ : -20/01/2003 - FB 10956 - le solde theorique ne 
Suite ........ : s'incrémentait pas
Suite ........ : - FB 11819 - le rafraichissement de la grille etaient incorrecte
Suite ........ : - 02/04/2003 - on numerote les lignes apres la fin du mode 
Suite ........ : guide, fct FinGuide
Suite ........ : - 11/06/2003 - bloque les deplacements a la souris ds le
Suite ........ : mode guide ds un folio de type libre
Suite ........ : - LG - 24/05/2007 - FB 19575 on affiche pas l'erreru sur la 
Suite ........ : premeire ligne
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var Ecr : TZF ;
begin
if csDestroying in ComponentState then Exit ; InfosPied(Ou,false) ;
if ObjFolio=nil then exit ; if bModeRO then exit ;
{$IFDEF TT}
AddEvenement('GSRowEnter Ou:='+intToStr(Ou));
{$ENDIF}
if memParams.bLibre and bGuideRun and ( Ou < GuideFirstRow ) then
 begin
  cancel := true ;
  exit ;
 end ;
if FCurrentRowEnter=Ou then begin SetGridOptions(Ou) ; exit ; end ;
FCurrentRowEnter:=Ou ; SoldeProgressif(Ou) ;
//LG* - 27/11/2001 gestion du boolean indiquant l'entree dans une nouvelle piece
if GS.Row=1 then
 bNewLigne:=true
  else
   if GS.Cells[SF_PIECE, Ou]<>'' then
    bNewLigne:=GS.Cells[SF_PIECE, Ou]<>GS.Cells[SF_PIECE, Ou-1];
if bNewLigne then memJal.CurAuto:=1 ;
if (GS.Row>(nbLignes+1)) then  // test pour le controle fin
 begin
  FCurrentRowEnter:=-1 ; Cancel:=true ; exit ;
 end ; // on est au sur une ligne non valide, on se replace sur la ligne precedentes
 if ObjFolio = nil then exit ;
Ecr:=ObjFolio.GetRow(GS.Row-GS.FixedRows) ;
if (Ecr<>nil) and (Ecr.GetValue('GUIDE')<>'X') then
 begin
  GuideStop ;
 end ;
SetGridOptions(Ou) ;
if GS.Row=(nbLignes+1) then
 begin
  //if bGuideRun then FinGuide ;
  GuideStop ;
  Statut:=taCreat ;
  Cancel:=not NextRow(GS.Row) ;
  if Cancel then Statut:=taConsult ;
 end //if
  else
   if (GS.Row<nbLignes) and memParams.bLibre and not bGuideRun then Statut:=taConsult ;
// les labels ne sont plus visible quand on change de ligne
//LSA_SOLDEG.Visible:=FALSE ; LSA_SOLDET.Visible:=FALSE ;
if memParams.bLibre then GS.Cells[SF_PIECE, Ou]:='1'
 else
  if Ou=nbLignes then SetPiece ;
// deplacer apres la creation de la nouvelle ligne
EnableButtons ;
end;

procedure TFSaisieFolio.GSExit(Sender: TObject) ;
var Cancel, Chg : Boolean ; ACol, ARow, ATop : LongInt ;
begin
if csDestroying in ComponentState then Exit ;
{$IFDEF TT}
AddEvenement('GSExit');
{$ENDIF}
if bModeRO then Exit ;
if Comptes=Nil then Exit ;
if FboFenLett then exit ;
ACol:=GS.Col ; ARow:=GS.Row ; ATop:=GS.TopRow ; Cancel:=FALSE ; Chg:=FALSE ;
GSCellExit(GS, ACol, ARow, Cancel) ;
GSRowExit(GS, ARow, Cancel, Chg) ;
GS.TopRow:=ATop ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 03/12/2002
Modifié le ... : 06/08/2004
Description .. : - 03/12/2002 - fiche 10821 - on peut saisir dans la zone
Suite ........ : nature
Suite ........ : - 04/12/2002 - on peut saisir dasn la zone nature
Suite ........ : uniquement en code abrege
Suite ........ : - 18/03/2004 - LG - le vk_escape ne passe plus la grille en
Suite ........ : modification
Suite ........ : - 06/08/2004 - LG - FB 14067 - passait la grille en modif qd
Suite ........ : on tab sur un bordereau vide avec affichage des natures 
Suite ........ : completes
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.GSKeyPress(Sender: TObject; var Key: Char);
begin
{$IFDEF TT}
//AddEvenement('KeyPress :'+Key) ;
{$ENDIF}
if not GS.SynEnabled then Key:=#0 else
   begin
   if Key=#13 then Key:=#9 ;
   if (Key='+') and ((GS.Col=SF_GEN) or (GS.Col=SF_AUX)) then Key:=#0 ;
   if (Key='+') and ((GS.Col=SF_GEN) or (GS.Col=SF_AUX)) then Key:=#0 ;
   if FBoNatureComplete and (Key='+') and (GS.Col=SF_NAT) then Key:=#0 ;
   if (Key=' ') and (GS.Col=SF_NAT) then Key:=#0 ;
   if FBoNatureComplete and (Key<>' ') and (GS.Col=SF_NAT) then Key:=#0 ;
   end ;
//LG*
if ( Key <> #9 ) and (  Key <> #27 ) and ( Key <> #0 ) then
 begin
  Statut := taModif;
  GrilleModif:=true ;
 end ;
FBoClickSouris := false ;
{$IFDEF TT}
AfficheTitreAvecCommeInfo ;
{$ENDIF}
end;

procedure TFSaisieFolio.GSMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin GridX:=X ; GridY:=Y ; bUp:=TRUE ; end ;

procedure TFSaisieFolio.GSMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
bUp:=FALSE ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 07/08/2002
Modifié le ... : 27/08/2004
Description .. : Se declenche quand on clique sur le panel contenant la
Suite ........ : grille de saisie
Suite ........ : 
Suite ........ : - 28/11/2002 - on declenche le GSEnter pour forcer la
Suite ........ : creation de nouvelles lignes
Suite ........ : - 25/03/2004 - nouvelle gestion qd on rentre ds la grille
Suite ........ : - 27/08/2004 - FB 14295 - FB 14298 - empeche de rentrer
Suite ........ : ds la grille en validation
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.PFENMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var r : TRect ;
begin
FBoReading:=true;
{$IFDEF TT}
AddEvenement('$$$$$$$$$$$$$PFENMouseDown') ;
{$ENDIF}
if FBoValEnCours then exit ;
FBoPFENMouseDown := true ;
FBoClickSouris := true ;
try
if (not GS.Enabled) and (not bClosing) then
  begin
  r:=Rect(GS.Left, GS.Top, GS.Left+GS.Width, GS.Top+GS.Height) ;
  if PtInRect(r, Point(X, Y)) then
    begin
    (* PFU : DEBUT 3789 *)
    if Valeur(CbFolio.Text)>0
     then NumFolio:=Trunc(Valeur(CbFolio.Text))
     else begin
       CbFolio.Text:=IntToStr(NumFolio) ;
       if not bModeRO then CbFolio.SetFocus ; Exit ;
     end ;
    (* PFU : FIN 3789 *)
    if not ReadFolio(FALSE) then
      begin if not bModeRO then CbFolio.SetFocus ; Exit ; end ;
    CalculSoldeJournal ;
    // Activer le Grid
    //EnableButtons ;
    SetGridOn ;
  //  GSEnter(nil) ;
    if not bModeRO then
     GS.InplaceEditor.Setfocus ;
    if FInRow > 1 then GS.row := FInRow - 1 else GS.row := FInRow + 1 ; // pour donner le focus le focus a la grille qd tu clique dessus
    if not bModeRO then GS.row := FInRow ;

    end ;
  if bModeRO then PostMessage(GS.Handle, WM_LBUTTONUP, 0, 0) ;
  end ;
finally
FBoReading:=false; FBoPFENMouseDown := false ;
{$IFDEF TT}
AddEvenement('********************PFENMouseDown') ;
{$ENDIF}
end;
end;

//=======================================================
//=============== Evénements des boutons ================
//=======================================================
procedure TFSaisieFolio.BChercherClick(Sender: TObject);
begin SearchClick ; end ;

procedure TFSaisieFolio.BValideClick(Sender: TObject);
begin ValClick(FALSE) ; end ;

procedure TFSaisieFolio.BAbandonClick(Sender: TObject);
begin
AbandonClick(False) ;
end ;

procedure TFSaisieFolio.BSDelClick(Sender: TObject);
begin DelClick ; end ;

procedure TFSaisieFolio.BInsertClick(Sender: TObject);
begin InsClick ; end ;

procedure TFSaisieFolio.BSoldeClick(Sender: TObject);
begin SoldeClick(-1) ; end ;

procedure TFSaisieFolio.BModifTvaClick(Sender: TObject);
begin ParamTvaClick ; end ;

procedure TFSaisieFolio.BComplementClick(Sender: TObject);
begin
 ComplClick ;
end ;

procedure TFSaisieFolio.BLibAutoClick(Sender: TObject) ;
begin RefLibAutoClick ; end ;

procedure TFSaisieFolio.BEchClick(Sender: TObject) ;
begin EchClick ; end ;

procedure TFSaisieFolio.BVentilClick(Sender: TObject);
begin AnaClick ; end ;

procedure TFSaisieFolio.BTGuideClick(Sender: TObject);
begin GuiClick ; end ;

procedure TFSaisieFolio.SearchClick ;
begin FindFirst:=TRUE ; FindSais.Execute ; end ;


procedure TFSaisieFolio.BDeviseClick(Sender: TObject);
begin DevClick ; bStopDevise:=TRUE ; end ;

procedure TFSaisieFolio.BModifRIBClick(Sender: TObject);
begin RibClick ; end ;

procedure TFSaisieFolio.BZoomImmoClick(Sender: TObject);
begin ImmoClick ; end;

procedure TFSaisieFolio.FindSaisFind(Sender: TObject);
var bCancel : Boolean ;
begin
Rechercher(GS, FindSais, FindFirst) ;
//EnableButtons ;
//SetGridOptions(GS.Row) ;
GSRowEnter(nil, GS.Row, bCancel, FALSE);
end ;

procedure TFSaisieFolio.BSelectFolioClick(Sender: TObject);
begin GetFolioClick ; end ;

{procedure TFSaisieFolio.ByeClick ;
begin if not IsInside(Self) then Close ; end ;}

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 11/03/2004
Modifié le ... : 17/05/2006
Description .. : - LG - 11/03/2004 - reecriture pour l'eAGL
Suite ........ : - LG - 18/03/2004 - réécriture suite au changement de 
Suite ........ : gestion
Suite ........ : - LG - 20/042004 - on ne peu pas revenir ds l'entete si on a 
Suite ........ : appele le lettrage ou la consultation des comptes
Suite ........ : - LG - 30/04/2004 - on sort si on ets en train de valider le 
Suite ........ : bordereau
Suite ........ : - LG - FB 13780 - 13772 - reaffectation du bEsc a false, on
Suite ........ : pouvait 
Suite ........ : sortir sans verif du compte de contrepartie en mode libre
Suite ........ : - LG - FB 13769 - rafraichissement des bp suite a un echap
Suite ........ : - LG - 11/08/2004 - FB 13534 - suppression d'un mesasge 
Suite ........ : en double ( questionj pose ds le formclose query )
Suite ........ : - LG - 23/09/2004 - FB 13535 - gestion du echap sur une
Suite ........ : ligne lettre ( on sortait directement de la saisie )
Suite ........ : - LG - 01/12/2004 - FB 15047 - correction de la fermeture 
Suite ........ : avec l'option bouton an bas
Suite ........ : - LG - 17/05/2006 - FB 17715 - en testn des guides on 
Suite ........ : supprimer toutes les options dispo pour l'abandon de la 
Suite ........ : saisie
Mots clefs ... : 
*****************************************************************}
Procedure TFSaisieFolio.AbandonClick ( FromEsc : boolean ) ;
var
 lBoGrilleBloquer : boolean ;
begin

 if csDestroying in ComponentState then Exit ;
 if FBoValEnCours or ( bWriting or bReading or bClosing ) then exit ;

 {$IFDEF TT}
 AddEvenement('AbandonClick ' );
{$ENDIF}

 bESC               := True ;
 FStLastTextGuide   := '' ;
 FStLastResultGuide := false ;

 if ( InitParams.ParGuide = '' ) then
  begin

    if bGuideRun then
     begin
      if (PrintMessageRC(RC_GUIDESTOP)=mrYes) then GuideStop ;
      Exit ;
     end ;

    {$IFNDEF TT1}
    if FBoLettrageEnCours and ( not bModeRO ) then
    begin
      // LG 22/10/2003 - il faut absolument enregistrer la bordereau si on a commence un lettrage en saisie
      PrintMessageRC(RC_ABANDONINTERDIT) ;
      exit ;
    end;
    {$ENDIF}

    lBoGrilleBloquer := ( IsRowCanEdit(GS.Row) <> 0 ) ;
    FTOBEnCours:=nil ;

    if (not bModeRO) and ( not GrilleModif ) and  (GS.InplaceEditor.Focused or lBoGrilleBloquer ) and ( InitParams.ParGuide = '' ) then
     begin
      CurNumFolio := '' ;
      InitFolio ;
      CbFolio.Enabled   := true ;
      E_JOURNAL.Enabled       := true ;
      E_DATECOMPTABLE.Enabled := true ;
      if E_DATECOMPTABLE.CanFocus then E_DATECOMPTABLE.SetFocus ;
      GrilleModif := false ; FBoClickSouris := false ; bESC := false ; EnableButtons ;
      CBloqueurJournal(false,E_JOURNAL.Value) ;
      if bNewLigne and memJal.IncRef then _DecS(FLastRef) ;
      Exit ;
     end ;

    if (not bModeRO) and GrilleModif and ( GS.InplaceEditor.Focused or lBoGrilleBloquer ) and ( InitParams.ParGuide = '' )  then
     begin
      if ValideFolio(TRUE, TRUE) then
       begin
        CurNumFolio := '' ;
        InitFolio ;
        CbFolio.Enabled   := true ;
        E_JOURNAL.Enabled       := true ;
        E_DATECOMPTABLE.Enabled := true ;
        if E_DATECOMPTABLE.CanFocus then E_DATECOMPTABLE.SetFocus ;
        GrilleModif := false ; EnableButtons ; CBloqueurJournal(false,E_JOURNAL.Value) ;
        if bNewLigne and memJal.IncRef then _DecS(FLastRef) ;
       end ;// if

      bEsc := false ;
      exit ;
     end ;

    // quand on fait echap sur un bordereau en consultation avec le curseur sur la grille
    if bModeRO and FBoCurseurDansGrille and ( InitParams.ParGuide <> '' )  then
     begin
      CurNumFolio              := '' ;
      InitFolio ;
      CbFolio.Enabled          := true ;
      E_JOURNAL.Enabled        := true ;
      E_DATECOMPTABLE.Enabled  := true ;
      if E_DATECOMPTABLE.CanFocus then E_DATECOMPTABLE.SetFocus ;
      FBoCurseurDansGrille     := false ;
      GrilleModif              := false ;
      FBoClickSouris           := false ;
      bESC                     := false ;
      EnableButtons ; CBloqueurJournal(false,E_JOURNAL.Value) ;
      Exit ;
     end ;

    bESC := false ;

  end ;

 {$IFDEF EAGLCLIENT}
   Close ;
   if IsInside(Self) then  CloseInsidePanel(Self) ;
 {$ELSE}
   Close ;
   GS.SynEnabled := false ;
   if IsInside(Self) then
    begin
    CloseInsidePanel(Self) ;
    THPanel(Self.parent).InsideForm := nil;
    THPanel(Self.parent).VideToolBar;
    // if PrepareInside then FMenuG.AfficheSoc(True,TRUE) ; //CloseInsidePanel(Self) ;
   end ;
{$ENDIF}

end ;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 16/09/2002
Modifié le ... : 25/09/2006
Description .. : - 16/09/2002 - on stocke le journal et la periode pour la
Suite ........ : representer à l'utilisateur
Suite ........ : - le solde automatique ne fonctionne pas pour les comptes
Suite ........ : collectif
Suite ........ : - 21/01/2003 - FB 10958 - controle de la presence du 
Suite ........ : compte de contrepartie sur un journal de banque
Suite ........ : -18/02/2003 - FB 11899 - suppression du parametre 
Suite ........ : Boucler en creation ( ne fct pas )
Suite ........ : - 05/05/2004 - LG - test pour vérifier que l'on n'est pas deja 
Suite ........ : en validation
Suite ........ : - 27/03/2003 - FB 12233 - en mode libre on prvoque un cell
Suite ........ : exit lors de l'enregistrement pour rafraichir la tob
Suite ........ : - 28/03/2003 - FB 12208 - la fenetre des scenario s'ouvrait
Suite ........ : pour toutes les lignes
Suite ........ : - 09/06/2003 - FB 12216 - la validations d'un bordereau sur
Suite ........ : une ligne d'ecart de conversion ne fct pas
Suite ........ : - 11/07/2003 - FB 12513 - ajout d'un boolean 
Suite ........ : FBoValeEnCours pour empecher la sortie avant la fin de
Suite ........ : l'enregistrement
Suite ........ : - LG - 13/04/2004 - FB 13386 - on bloque la validation
Suite ........ : pendant que le fenetre s'ouvre
Suite ........ : - 05/05/2004 - on bloque la validation qd on est deja en 
Suite ........ : validation
Suite ........ : - 05/08/2004 - FB 13913 13914 - on controle l'integralite de 
Suite ........ : laigne a la sortide de celle ci ( pb de validation qd on se
Suite ........ : deplca avec la souris, ou validation direct d'une ligne 
Suite ........ : incorrecte )
Suite ........ : - 05/08/2004 - FB 13866 - le Shift + validation permettait
Suite ........ : d'enregsitrer des ligens incorrectes. Avant chaque validation
Suite ........ : on calcul le solde et on sort si le bordereau n'est pas
Suite ........ : equilibre
Suite ........ : - LG  - 12/08/2004 - FB 13885 - pour les journaux de types
Suite ........ : libres, on teste les compte de contrepartie des journaux de
Suite ........ : banque sur un folio vide !
Suite ........ : - 27/08/2004 - FB 14295 - FB 14298 - empeche de rentrer
Suite ........ : ds la ghrille en validation
Suite ........ : - 22/10/2004 - FB 13599 - LG - avant d'ouvrir un bordereau 
Suite ........ : avec des ecritures de type L, on regarde s'il n'est pas looke 
Suite ........ : par un autre utilisateur
Suite ........ : - 02/05/2005 - FB 15816 - on rafraisi les boutons lors de la 
Suite ........ : validation pour empecher un autre traitement pendant celle 
Suite ........ : ci
Suite ........ : - 27/06/2005 - FB 15642 - Suppression Bordereau ou 
Suite ........ : toutes écritures d'un Bordereau - contexte Synchronisation 
Suite ........ : avec Sx
Suite ........ : - 22/07/2005 - FB 16248 - test du param cpliengamme
Suite ........ : incorrect
Suite ........ : - 04/01/2006 - FB 17251 - le F10 sur un ligne non soldée
Suite ........ : ne fct pas
Suite ........ : - 23/03/2006 - FB 17252 - test sur la synchor s1 fait apres 
Suite ........ : le calcul du solde du journal
Suite ........ : - 19/07/2006 - FB 17623 - on interdit de supprimer un 
Suite ........ : bordereau ss montant en liaison eWs
Suite ........ : - 25/09/2006 - FB 18856 - on tient compte du mode de
Suite ........ : synchro
Mots clefs ... : 
*****************************************************************}
function TFSaisieFolio.ValClick(bVerbose : Boolean(*=TRUE*);
                                bCanClose: Boolean(*=FALSE*)): Boolean ;
var Cancel, Chg : Boolean ; ACol, ARow : LongInt ;
EcartP, EcartD : Double ;
lQ : TQuery ;
i : integer;
lTobEcr : Tob ;
lBoSynchro : boolean ;
begin
{$IFDEF TT}
AddEvenement('ValClick');
{$ENDIF}
//SG6 21/12/2004 FQ 14731
if InitParams.ParGuide <> '' then
 begin
  result := true ;
  AbandonClick(false) ;
  exit ;
 end ;
Result:=FALSE ; FStLastTextGuide := '' ;  FStLastResultGuide := false ;
if bClosing or bReading then Exit ;
if FboFenLett then exit ; if bGuideRun then exit ;
if Screen.ActiveControl<>GS then begin Result:=TRUE ; Exit ; end ;
if (Memjal.EqAuto)and (not IsSoldeFolio) and (GS.Cells[SF_GEN, GS.Row]='') then
 begin
  result:=SoldeAuto ;
  if not result then exit ;
 end ;
SourisSablier ;
if FBoValEnCours then exit ;
FBoValEnCours:= true ; EnableButtons ;
try
FTOBEnCours:=nil ;
//LG*1 06/12/2001 pour les journaux de type libre on recalcul les soldes dans toutes les
// devises et on controle la coherence solde et mode de saisie ( le bouton valide est actif même si le folio n'est
// pas equilibre )
// 30/01/2002 utilisation de la fonction abs() pour tester les ecarts : les ecart au credit sont negatifs
if memParams.bLibre then
 BEGIN
  //if not bGuideRun then SendMessage(GS.Handle,WM_KEYDOWN,VK_TAB,0) ;
  AfficheTitreAvecCommeInfo('Vérification des soldes');
  GetAnySolde(StrToInt(GS.Cells[SF_PIECE, 1]),EcartP, EcartD) ;
  if memParams.bDevise then BEGIN if (Arrondi(Abs(EcartD),memPivot.Decimale)>0) then exit; END
  else if (Arrondi(Abs(EcartP),memPivot.Decimale)>0) then
   begin
    AfficheTitreAvecCommeInfo('Bordereau non équilibré ! ');
    exit;
   end;
 END;// mode liébre
//if not BValide.Enabled then begin Result:=TRUE ; Exit ; end ;
ACol:=GS.Col ; ARow:=GS.Row ; Cancel:=FALSE ; Chg:=FALSE ;
if not bEsc then
 begin
  bEnreg:=TRUE ;
  GSCellExit(GS, ACol, ARow, Cancel) ; bEnreg:=FALSE ;
  if Cancel then begin PrintMessageRC(RC_BADROW) ; Exit ; end ;
  bEnreg:=TRUE ;
  FCurrentRowEnter:=-1 ;
  if (memParams.bLibre) and (IsSoldeFolio) and ( nbLignes > 1 ) then
   begin
    // Vérification du compte de contrepartie
    AfficheTitreAvecCommeInfo('vérification du compte de contrepartie');
    if not VerifTreso(1,ARow) then
     begin PrintMessageRC(RC_TRESO2) ; GS.Col:=SF_GEN ; bEnreg:=false ; Exit ; end ;
   end ;
  FBoValEnCours:= false ;
  Statut:=taModif ;
  GSRowExit(GS, ARow, Cancel, Chg) ;
  FBoValEnCours:= true ;
  bEnreg:=FALSE ;
  if Cancel then Exit ;
  //LG*1 06/12/2001 calcul du solde du bordereau
  if (memParams.bLibre) then VerifEquilibre(StrToInt(GS.Cells[SF_PIECE, 1]), 1) ;

 end ;
CalculSoldeJournal ;
if not IsSoldeFolio then begin nextRow(GS.Row) ; exit ; end ;
// ajout me fiche 19288 lBoSynchro := ( ( GetParamsocSecur('SO_CPLIENGAMME','') = 'S1' ) and GetParamsocSecur('SO_CPMODESYNCHRO',False) ) or ( GetParamsocSecur('SO_CPLIENGAMME','') = 'S5' );
lBoSynchro := (GetParamsocSecur('SO_CPMODESYNCHRO',False) ) and ( ( GetParamSocSecur('SO_CPLIENGAMME' , '' ) = 'S1' ) or ( GetParamsocSecur('SO_CPLIENGAMME','') = 'S5' ));
if lBoSynchro and (nbLignes=1) and IsSoldeFolio then
 begin
 // ajout me fiche 19288  PGIInfo('Vous ne pouvez pas supprimer ce bordereau dans un contexte Synchronisation avec ' + RechDom('CPLIENCOMPTABILITE', GetParamSocSecur('SO_CPLIENGAMME' , '' ), false ) , Self.Caption) ;
   PGIBox ('Le dossier est en liaison avec ' + RechDom('CPLIENCOMPTABILITE', GetParamSocSecur('SO_CPLIENGAMME' , '' ), false ) + ', la suppression est interdite.', Self.Caption);
   exit ;
 end ;
if ( memJal.TotFolDebit = 0 ) and ( memJal.TotFolCredit = 0 ) and lBoSynchro then
 begin
 // ajout me fiche 19288   PGIInfo('Vous ne pouvez pas supprimer ce bordereau dans un contexte Synchronisation avec ' + RechDom('CPLIENCOMPTABILITE', GetParamSocSecur('SO_CPLIENGAMME' , '' ), false ) ,Self.Caption) ;
   PGIBox ('Le dossier est en liaison avec ' + RechDom('CPLIENCOMPTABILITE', GetParamSocSecur('SO_CPLIENGAMME' , '' ), false ) + ', la suppression est interdite.', Self.Caption);
   exit ;
 end ;
Statut:=taConsult ;
if ValideFolio(bVerbose, bCanClose) then
  begin
  Result:=TRUE ;
  // Dev 3946 : control cohérence qté analytique soumis à indicateur
  if (not (ctxPCL in V_PGI.PGIContexte)) and GetParamSocSecur('SO_ZCTRLQTE', False) then
   for i:=0 to ObjFolio.Ecrs.Detail.count - 1 do
    begin
      lTobEcr := ObjFolio.GetRow( i ) ;
      if ( lTobEcr <> nil ) and ( lTobEcr.GetValue('E_ANA')='X' )
                            and ( ( lTobEcr.GetNumChamp('CHECKQTE') < 0 ) or
                                  ( lTobEcr.GetString('CHECKQTE')<>'X'  )
                                ) then
        if not CheckQuantite( lTobEcr, False ) then
          begin
          PGIInfo('Certaines lignes du bordereau présentent des incohérences de quantités saisies entre la ligne d''écriture et les lignes d''écritures analytiques.');
          break ;
          end ;
    end ;
  // Fin Dev 3946
  InitFolio ;
  end
else Exit ;
// Ajoute le folio dans la Combo
if (V_PGI.IoError=oeOk) and (CbFolio.Values.IndexOf(IntToStr(NumFolio))<0) then
   begin
   CbFolio.Items.Add(IntToStr(NumFolio)) ;
   CbFolio.Values.Add(IntToStr(NumFolio)) ;
   end ;
// on stocke le journal et la periode pour la representer à l'utilisateur
SetParamSoc('SO_ZLASTDATE', StrToDate(E_DATECOMPTABLE.Value)) ; SetParamSoc('SO_ZLASTJAL',E_JOURNAL.Value ) ;
// on recherche s'il reste des ecritures de types L , et l'on reouvre la saisie bordereau avec
{lQ := OpenSQL('select E_DATECOMPTABLE,E_JOURNAL,E_NUMEROPIECE,E_NUMLIGNE from ECRITURE where E_QUALIFPIECE="L" ' , true) ;
if not lQ.Eof and  not CEstBloqueBor( lQ.FindField('E_JOURNAL').AsString,
                        lQ.FindField('E_DATECOMPTABLE').AsDateTime,
                        lQ.FindField('E_NUMEROPIECE').AsInteger,
                        false) and not CEstBloqueLett(lQ.FindField('E_JOURNAL').AsString ,
                                                        lQ.FindField('E_DATECOMPTABLE').AsDateTime ,
                                                        lQ.FindField('E_NUMEROPIECE').AsInteger, false ) then
 begin
  FillChar(InitParams, Sizeof(InitParams), #0) ;
  InitParams.ParPeriode:=DateToStr((lQ.FindField('E_DATECOMPTABLE').AsDateTime)) ;
  InitParams.ParCodeJal:=lQ.FindField('E_JOURNAL').AsString ; InitParams.ParNumFolio:=lQ.FindField('E_NUMEROPIECE').AsString ;
  InitParams.ParNumLigne:=lQ.FindField('E_NUMLIGNE').AsInteger ; FBoLettrageEnCours:=true ; InitParams.ParRecupLettrage := true ;
  Ferme(lQ) ;
  AlimenteEntete(InitParams.ParPeriode, InitParams.ParCodeJal, InitParams.ParNumFolio) ;
  ReadFolio(FALSE) ;
  CalculSoldeJournal ;
  SetGridOn ;
  if GS.CanFocus then GS.SetFocus ;
  exit ;
 end
  else
   begin // sinon }
    CBloqueurJournal(false,E_JOURNAL.Value) ;
    InitParams.ParRecupLettrage:=false ;
    FBoLettrageEnCours:=false ;   // on baisse le flag indiquant des ecritures de types L
    Ferme(lQ) ;
//   end;
// Suivant ...
Inc(NumFolio) ;
while CbFolio.Values.IndexOf(IntToStr(NumFolio))>=0 do Inc(NumFolio) ;
CbFolio.Text:=IntToStr(NumFolio) ;
if not bClosing then
  begin
  memJal.TotDebDebit:=0 ;  memJal.TotDebCredit:=0 ;
  ChargeJal(memJal.Code);  CalculSoldeJournal ;
  SetGridOptions(GS.Row) ; FBoClickSouris := false ;
//  if CbFolio.CanFocus then CbFolio.SetFocus ;
  end ;

finally
 SourisNormale ; FBoValEnCours:= false ; EnableButtons ;
end;

 CbFolio.Enabled         := true ;
 E_JOURNAL.Enabled       := true ;
 E_DATECOMPTABLE.Enabled := true ;

 if not bClosing and CbFolio.CanFocus then CbFolio.SetFocus ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 08/04/2002
Modifié le ... : 21/07/2005
Description .. : - on n'essaye plus qu'enregistrer 2 fois le bordereau ( au lieu
Suite ........ : de 5 )
Suite ........ : -en cas d'erreur on reste sur le bordereau pour
Suite ........ : eventuellement le modifier
Suite ........ : - LG - 21/07/2005 - FB 12908 - suppression du document
Suite ........ : numeri en cas d'abandon
Mots clefs ... :
*****************************************************************}
function TFSaisieFolio.ValideFolio(bVerbose: Boolean(*=TRUE*);
                                   bCanClose: Boolean(*=FALSE*);
                                   bOnlySave: Boolean(*=FALSE*)) : Boolean ;
//var Start : LongInt ;
var myvalid,myValidCpte,myvalidAno : TIOErr ;
begin
Result:=FALSE ;
if (GS=nil) or ((GS.Enabled=FALSE) and (nbLignes<=0)) then begin Result:=TRUE ; Exit ; end ;
if nbLignes<=0 then Exit ;
// Gestion du message d'abandon
if not bOnlySave then
  begin
  if bVerbose then if PrintMessageRC(RC_ABANDON)=mrYes then
   begin
    Result:=TRUE ;
    Exit ;
   end
    else
     begin if bCanClose then Exit ; end ;
  end ;
//start := GetTickCount;
bWriting:=TRUE ;
CDeBlocageLettrage ;
myvalid:=Transactions(WriteFolio, 1) ;
if myvalid<>oeOk then BEGIN PGIInfo('Ce bordereau n''a pas été enregistré'+#10+#13+ V_PGI.LastSQLError,Caption); result:= false ; exit; END;
If (myvalid=oeOk) Then
  BEGIN
  FBoLettrageEnCours:=false ; InitParams.ParRecupLettrage:=false ;
  AfficheTitreAvecCommeInfo('Mise à jour des soldes');
//  myvalidCpte:=oeOK ;
//  if not FboFenLett then
  myvalidCpte:=Transactions(WriteCumulCpte, 2) ;
  If myvalidCpte<>oeOK Then
    BEGIN
    Transactions(ObjFolio.EcritLogSB,1) ;
    //MessageAlerte(GetMessageRC(48)) ;
    END
     else            
      begin
       myvalidAno := Transactions(WriteCumulAno,2) ;
       if myvalidAno<>oeOK Then
        begin
         Transactions(ObjFolio.EcritLogSB,1) ;
         if V_PGI.SAV then
          MessageAlerte(GetMessageRC(90)+#10#13+V_PGI.LastSQLError)
           else
            MessageAlerte(GetMessageRC(90)) ;
       end ;
      end ;
 end ;
if myvalid<>oeOk then
 begin
   PrintMessageRC(RC_BADWRITE) ; FiniMove ;
 end
  else
   if not bOnlySave then
    begin
    FRechCompte.Info.Compte.Clear ; FRechCompte.Info.Aux.Clear ;
  end ;
if not bOnlySave then
  CurNumFolio:='' ;
bWriting:=FALSE ;
//ShowMessage(Format('%f', [(GetTickCount-Start)/1000])) ;
E_ETABLISSEMENT.Enabled:=FALSE ;
Result:=TRUE ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 17/09/2002
Modifié le ... : 06/11/2006
Description .. : - LG - 17/09/2002 - la ligne courante est remise a zero pour
Suite ........ : redeclencher l'evement create row
Suite ........ : - LG - 06/12/2002 - fiche 10805 - message lors de la 
Suite ........ : suppression d'une
Suite ........ : ligne avec info complementaire
Suite ........ : - LG - 21/10/2005 - on ne lance pas la suppression sur une 
Suite ........ : ligne lettre ou pointe
Suite ........ : - LG - 06/11/2006 - FB 16744 - on remets à jour les info du
Suite ........ : compte
Mots clefs ... :
*****************************************************************}
procedure TFSaisieFolio.DelClick ;
var lr : LongInt ; Decal : Integer ;
Ecr : TZF ;
begin
{$IFDEF TT}
AddEvenement('DelClick') ;
{$ENDIF}
if BSDel.Enabled=FALSE then Exit ; if IsRowCanEdit(GS.Row) > 0 then exit ;
if ObjFolio=nil then Exit ; Ecr:=ObjFolio.GetRow(GS.Row-GS.FixedRows) ;
if Ecr=nil then Exit ;
if Ecr.GetValue('COMPS')='X' then
 begin
   if PGIAsk('Il existe des informations complémentaires, voulez-vous vraiment supprimer la ligne ?')<>mrYes then exit ;
 end;
{$IFDEF SCANGED}
ObjFolio.SupprimeLeDocGuid(Ecr) ;
{$ENDIF}
lr:=GS.Row ; Decal:=DeleteRow(lr) ;
if lr>nbLignes then GS.Row:=nbLignes else GS.Row:=lr+Decal ;
GS.Col:=SF_NAT ; // Se placer dans une colonne valide
if GS.Row<>GS.FixedRows then GS.ElipsisButton:=FALSE ;
SetGridOptions(GS.Row) ; InfosPied(GS.Row) ;
if not IsPieceCanEdit(GS.Cells[SF_PIECE, GS.Row]) then GS.Col:=SF_GEN ;
FCurrentRowEnter:=-1 ; EnableButtons ; bDelDone:=TRUE ; GS.Invalidate ;
GrilleModif:=true ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 07/11/2002
Modifié le ... : 30/07/2003
Description .. : -07/11/2002- enregistrement de la ligne courante et
Suite ........ : calcul du solde
Suite ........ : 
Suite ........ : -30/07/2003 - FB 12602 - insert par F11 ne fct pas
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.InsClick ;
var lr : LongInt ;
Cancel,Chg : boolean ;
ACol,ARow : integer ;
begin
if BInsert.Enabled=FALSE then Exit ;
{$IFDEF TT}
AddEvenement('InsClick') ;
{$ENDIF}
FCurrentRowEnter:=-1 ; ACol:=GS.Col ; ARow:=GS.Row ; Cancel:=false  ;
// on formate la cellule courante et on enregistre la ligne
GSCellExit(GS, ACol, ARow, Cancel) ; if Cancel then exit ; Chg:=false ;
GSRowExit(GS, GS.Row, Cancel, Chg) ; if Cancel then exit ; CalculSoldeJournal ;
lr:=GS.Row(*+1*) ; CreateRow(lr) ; GS.Row:=lr ;
SetGridOptions(GS.Row) ;
if not IsPieceCanEdit(GS.Cells[SF_PIECE, GS.Row]) then GS.Col:=SF_GEN ;
EnableButtons ;
GrilleModif:=true ;
GS.Invalidate ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 07/10/2002
Modifié le ... :   /  /
Description .. : -07/10/2002- on passe la grille en modif apres un f6 pour
Suite ........ : enregistrer la grille dans la rowexit
Mots clefs ... :
*****************************************************************}
procedure TFSaisieFolio.SoldeClick(Row : LongInt) ;
var lr : LongInt ; Solde, Debit, Credit : Double ; bSoldeDebit, bMove : Boolean ;
    ACol, ARow : LongInt ; Cancel : Boolean ;
begin
{$IFDEF TT}
AddEvenement('SoldeClick') ;
{$ENDIF}
if (not bGuideRun) and (not BSolde.Enabled) then Exit ;
if Row<0 then lr:=GS.Row else lr:=Row ;
if (GS.Cells[SF_DEVISE, lr]<>'') or (GS.Cells[SF_EURO, lr]='X') then Exit ;
ACol:=GS.Col ; ARow:=GS.Row ; Cancel:=FALSE ;
if (not bGuideRun) then begin GSCellExit(GS, ACol, ARow, Cancel) ; if Cancel then Exit ; end ;
if (GS.Cells[SF_GEN, lr]='') then Exit ;
Statut:=taCreat ;
bMove:=(not bGuideRun) and ((GS.Col=SF_DEBIT) or (GS.Col=SF_CREDIT)) ;
Solde:=Arrondi(Abs(memJal.TotFolDebit-memJal.TotFolCredit), memPivot.Decimale) ; //SA_SOLDE.Value ;
bSoldeDebit:=(memJal.TotFolDebit-memJal.TotFolCredit) > 0 ; //SA_SOLDE.Debit ;
Debit:=Valeur(GS.Cells[SF_DEBIT,lr]) ; Credit:=Valeur(GS.Cells[SF_CREDIT,lr]) ;
// Partie débit
if (Debit<>0) and (bSoldeDebit) then
   if (Debit-Solde)<0 then
      begin
      GS.Cells[SF_CREDIT,lr]:=StrFMontant(Abs(Debit-Solde),15,memPivot.Decimale,'',TRUE) ;
      GS.Cells[SF_DEBIT,lr]:='' ;
      if bMove then GS.Col:=SF_CREDIT ;
      end else
      begin
      GS.Cells[SF_DEBIT,lr]:=StrFMontant(Debit-Solde,15,memPivot.Decimale,'',TRUE) ;
      if bMove then GS.Col:=SF_DEBIT ;
      end ;
if (Debit<>0) and (not bSoldeDebit) then
   begin
   GS.Cells[SF_DEBIT,lr]:=StrFMontant(Debit+Solde,15,memPivot.Decimale,'',TRUE) ;
   if bMove then GS.Col:=SF_DEBIT ;
   end ;
// Partie crédit
if (Credit<>0) and (bSoldeDebit) then
   begin
   GS.Cells[SF_CREDIT,lr]:=StrFMontant(Credit+Solde,15,memPivot.Decimale,'',TRUE) ;
   if bMove then GS.Col:=SF_CREDIT ;
   end ;
if (Credit<>0) and (not bSoldeDebit) then
   if (Credit-Solde)<0 then
      begin
      GS.Cells[SF_DEBIT,lr]:=StrFMontant(Abs(Credit-Solde),15,memPivot.Decimale,'',TRUE) ;
      GS.Cells[SF_CREDIT,lr]:='' ;
      if bMove then GS.Col:=SF_DEBIT ;
      end else
      begin
      GS.Cells[SF_CREDIT,lr]:=StrFMontant(Credit-Solde,15,memPivot.Decimale,'',TRUE) ;
      if bMove then GS.Col:=SF_CREDIT ;
      end ;
if (Debit=0) and (Credit=0) then
   begin
   if bSoldeDebit then
      begin
      GS.Cells[SF_CREDIT,lr]:=StrFMontant(Solde,15,memPivot.Decimale,'',TRUE) ;
      if bMove then GS.Col:=SF_CREDIT ;
      end else
      begin
      GS.Cells[SF_DEBIT,lr]:=StrFMontant(Solde,15,memPivot.Decimale,'',TRUE) ;
      if bMove then GS.Col:=SF_DEBIT ;
      end ;
   end ;
GS.ElipsisButton:=FALSE ;
if (not bGuideRun) then SetRowLib(lr) ;
CalculSoldeJournal ;
if (not bGuideRun) and (Row<0) then
  begin
  if GS.Cells[SF_DEBIT,  lr]<>'' then GS.Col:=SF_DEBIT ;
  if GS.Cells[SF_CREDIT, lr]<>'' then GS.Col:=SF_CREDIT ;
  PostMessage(GS.Handle, WM_KEYDOWN, VK_TAB, 0) ;
  end ;
GrilleModif:=true ;
EnableButtons ;
end ;

procedure TFSaisieFolio.ParamTvaClick ;
var Ecr : TZF ; Obm : TObm ; NumCompte : string ; k : Integer ; Action : TActionFiche ;
begin
if not BModifTva.Enabled then Exit ;
if ObjFolio=nil then Exit ;
Ecr:=ObjFolio.GetRow(GS.Row-GS.FixedRows) ; if Ecr=nil      then Exit ;
NumCompte:=GS.Cells[SF_GEN, GS.Row] ;       if NumCompte='' then Exit ;
k:=Comptes.GetCompte(NumCompte) ;           if k<0          then Exit ;
if Comptes.IsHT(k) then
   begin
   // En attendant...
   if Ecr.GetValue('E_REGIMETVA')='' then Exit ;
   Obm:=ObjFolio.GetObmFromRow(GS.Row-GS.FixedRows, RichEdit) ;
   if Obm=nil then Exit ;
//   if O.GetMvt('E_REGIMETVA')='' then O.PutMvt('E_REGIMETVA',GeneRegTva) ;
   if ((bModeRO) or (IsRowCanEdit(GS.Row)>0)) then Action:=taConsult else Action:=taModif ;
   //InfosTvaEnc(Obm, Action) ; //XMG 20/04/04
   {$IFNDEF COMSX}
   {$IFDEF COMPTA}
   if (VH^.PaysLocalisation=CodeISOES) and (InfosTvaEnc(Obm, Action)) and (IsTiersSoumisTPF(GS)) then //XVI 24/02/2005
      Obm.PutMvt('E_TPF',obm.GetValue('E_TVA'))
   else
      InfosTvaEnc(Obm, Action) ;
   {$ENDIF}
   {$ELSE}
   InfosTvaEnc(Obm, Action) ;
   {$ENDIF}
   end
end ;



{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 26/09/2002
Modifié le ... : 12/04/2006
Description .. : -26/09/2002 - suppression d'une fuite memoire sur lobjet
Suite ........ : lAna
Suite ........ : - 03/03/2005 - LG -  gestion du cutoff
Suite ........ : - LG  - 22/07/2005 - FB 15876 - qd on re-saisie un cutoff le 
Suite ........ : carre ble n'apparaissait pas
Suite ........ : - LG - 24/08/2005 - FB 16410 - plantais a la validation
Suite ........ : - LG - 12/04/2006 - FB 15949 - on permet de modifier les 
Suite ........ : info compl sur les ecritrues lettre , pointe etc etc
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.ComplClick ;
var Ecr : TZF ; Obm : TObm ; lAna : TObjAna ; RC : R_COMP ; Action : TActionFiche ;
    Qt1, Qt2, aQt1, aQt2 : Double ; bModifBlocNote : Boolean ;
    RefInterne, RefExterne, RefLibre : string ;
    k : integer ; NumCompte : string ;
begin
{$IFDEF TT}
AddEvenement('ComplClick') ;
{$ENDIF}
if ObjFolio = nil then exit ;
Ecr:=SetRowTZF(GS.Row,false) ; if Ecr=nil then Exit ;
if ( InitParams.ParGuide <> '' ) then exit ;
if ( Ecr.GetValue('COMPTAG') = '0' ) then
 begin
  Ecr.PutValue('COMPTAG','1') ;
  Ecr.PutValue('OLDGENERAL', '') ;
  GereCutOff(GS.Row) ;
  exit ;
 end ;
Qt1:=Ecr.GetValue('E_QTE1') ; Qt2:=Ecr.GetValue('E_QTE2') ;
RefInterne:=Ecr.GetValue('E_REFINTERNE') ; RefExterne:=Ecr.GetValue('E_REFEXTERNE') ;
RefLibre:=Ecr.GetValue('E_REFLIBRE') ;
RC.StLibre:='XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX' ;
RC.StComporte:='--XXXXXXXX' ; // Pas de ModeLC
RC.Conso:=TRUE ; RC.DateC:=Ecr.GetValue('E_DATECOMPTABLE') ;
RC.Attributs:=FALSE ; RC.MemoComp:=nil ; RC.Origine:=-1 ;
RC.TOBCompl := CGetTOBCompl(Ecr) ; RC.AvecCalcul := false ;
NumCompte:=GS.Cells[SF_GEN, GS.Row] ; k:=Comptes.GetCompte(NumCompte) ;
if (k=-1) then exit ;
RC.CutOffPer := Comptes.GetValue('G_CUTOFFPERIODE') ;
RC.CutOffEchue := Comptes.GetValue('G_CUTOFFECHUE') ;
Obm:=ObjFolio.GetObmFromRow(GS.Row-GS.FixedRows, RichEdit) ; if Obm=nil then Exit ; lAna:=nil ;
try
//if ((bModeRO) or (IsRowCanEdit(GS.Row)>0)) then Action:=taConsult else Action:=taModif ;
if bModeRO then Action:=taConsult else Action:=taModif ;
if not SaisieComplement(Obm, EcrGen, Action, bModifBlocNote, RC, False, True) then
  Exit
else
begin
  // GCO - 30/05/2006 - 17837
  GrilleModif := True;
  Statut := TaModif;
end;

aQt1:=Ecr.GetValue('E_QTE1') ; aQt2:=Ecr.GetValue('E_QTE2') ;
if bModeRO then begin Obm.Free ; Exit ; end ;
if not ObjFolio.IsAnaFromRow(GS.Row-GS.FixedRows) then
 begin
  ObjFolio.SetObmFromRow(GS.Row-GS.FixedRows, RichEdit, Obm) ;
  if (RC.TOBCompl<>nil) and (RC.TOBCompl.GetValue('EC_CUTOFFDEB')<>iDate1900) then
   Ecr.PutValue('COMPTAG','XD') ;
  GereQuantite(Ecr) ; 
  AnnuleCutOffSiBesoin(Ecr) ;
  Exit ;
 end ;
if ((aQt1<>Qt1) and (Qt1<>0)) then ProrateQteAnalTOB( Ecr, Qt1, aQt1, '1') ;
if ((aQt2<>Qt2) and (Qt2<>0)) then ProrateQteAnalTOB( Ecr, Qt2, aQt2, '2') ;
ObjFolio.SetObmFromRow(GS.Row-GS.FixedRows, RichEdit, Obm) ;
if (RC.TOBCompl<>nil) and (RC.TOBCompl.GetValue('EC_CUTOFFDEB')<>iDate1900) then
 Ecr.PutValue('COMPTAG','XD') ;
GereQuantite(Ecr) ;
AnnuleCutOffSiBesoin(Ecr) ;
// Changement de référence(s) ?
ModifRefAna(GS.Row-GS.FixedRows, 'REFINTERNE',     Null) ;
ModifRefAna(GS.Row-GS.FixedRows, 'REFEXTERNE',     Null) ;
ModifRefAna(GS.Row-GS.FixedRows, 'REFLIBRE',       Null) ;
ModifRefAna(GS.Row-GS.FixedRows, 'DATEREFEXTERNE', Null) ;
ModifRefAna(GS.Row-GS.FixedRows, 'CONSO',          Null) ;
ModifRefAna(GS.Row-GS.FixedRows, 'AFFAIRE',        Null) ;
finally
 if Obm<>nil then Obm.Free ; FreeAndNil(lAna) ;

end;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 25/07/2002
Modifié le ... : 25/09/2002
Description .. : Affichage des info complementaire ne fonction des
Suite ........ : scenarios
Suite ........ :
Suite ........ : -25/09/2002-suppression d'une fuite memoiresur l'obj Ana
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.GereComplements(Row : integer) ;
Var O      : TOBM ;
    ModBN  : boolean ;
    StComp,StLibre : String ;
    RC             : R_COMP ;
    NumRad         : integer ;
    Ecr            : TZF ;
    Qt1, Qt2, aQt1, aQt2 : Double ;
    k : integer;
    bQuantite      : boolean;    
    stCompte : string;
 //   Ana : TObjAna ;
BEGIN
{$IFDEF TT}
AddEvenement('GereComplements') ;
{$ENDIF}
if ObjFolio=nil then Exit ;
Ecr:=ObjFolio.GetRow(Row-GS.FixedRows) ; if Ecr=nil then Exit ;
O:=ObjFolio.GetObmFromRow(Row-GS.FixedRows, RichEdit) ; if O=Nil then Exit ;
try
if O.CompS then exit ; // une saisie complementaire a deja ete efectue
Qt1:=Ecr.GetValue('E_QTE1') ; Qt2:=Ecr.GetValue('E_QTE2') ;
stCompte := Ecr.GetValue('E_GENERAL');
k:=Comptes.GetCompte(stCompte) ;
if k >= 0 then
begin
    bQuantite :=  ( GetParamSocSecur('SO_CPPCLSAISIEQTE',False))  and
                  ((( Comptes.GetValue('G_QUALIFQTE1', k) <> '' ) and ( Comptes.GetValue('G_QUALIFQTE1', k) <> 'AUC' )) or
                  (( Comptes.GetValue('G_QUALIFQTE2', k) <> '' ) and ( Comptes.GetValue('G_QUALIFQTE2', k) <> 'AUC' ))) ;
end else bQuantite := False;
NumRad:=ComporteLigne(TOBM(FZScenario.Item),Ecr.GetValue('E_GENERAL'),StComp,StLibre) ;
if ((NumRad<=0) and (not bQuantite)) then Exit ;
if ((bQuantite) and (NumRad<=0)) then
begin
  StLibre:='----------------------------------------' ;
  StComp:='------';
  if (( Comptes.GetValue('G_QUALIFQTE1', k) <> '' ) and ( Comptes.GetValue('G_QUALIFQTE1', k) <> 'AUC' )) then
    stComp := stComp+'X' else stComp := stComp+'-';
  if (( Comptes.GetValue('G_QUALIFQTE2', k) <> '' ) and ( Comptes.GetValue('G_QUALIFQTE2', k) <> 'AUC' )) then
    stComp := stComp+'X' else stComp := stComp+'-';
  stComp := stComp + '--' ;
end;
if ((Pos('X',StComp)<=0) and (Pos('X',StLibre)<=0)) then Exit ;
RC.StComporte:=StComp ; RC.StLibre:=StLibre ; RC.Conso:=False ; RC.DateC:=Ecr.GetValue('E_DATECOMPTABLE') ;
RC.Attributs:=FAction=taCreat ; if (not bQuantite) then RC.MemoComp:=FZScenario.Memo ; RC.Origine:=NumRad ;
O.CompS:=True ;
if Not SaisieComplement(O,EcrGen,FAction,ModBN,RC) then
 begin
  ObjFolio.SetObmFromRow(Row-GS.FixedRows, RichEdit, O) ;
  GereQuantite(Ecr) ;
  Exit ;
 end
 else
 begin
   // GCO - 30/05/2006 - 17837
   GrilleModif := True;
   Statut := TaModif;
 end;
if GS.CanFocus then GS.SetFocus ;
aQt1:=Ecr.GetValue('E_QTE1') ; aQt2:=Ecr.GetValue('E_QTE2') ;
if bModeRO then begin O.Free ; Exit ; end ;
//Ana:=ObjFolio.GetObjFromRow(Row-GS.FixedRows, RichEdit) ;
//if Ana=nil then begin ObjFolio.SetObmFromRow(Row-GS.FixedRows, RichEdit, O) ; Exit ; end ;
if not ObjFolio.IsAnaFromRow(Row-GS.FixedRows) then
 begin
  ObjFolio.SetObmFromRow(Row-GS.FixedRows, RichEdit, O) ;
  GereQuantite(Ecr) ;
  Exit ;
 end ;
if ((aQt1<>Qt1) and (Qt1<>0)) then ProrateQteAnalTOB( O, Qt1, aQt1, '1') ;
if ((aQt2<>Qt2) and (Qt2<>0)) then ProrateQteAnalTOB( O, Qt2, aQt2, '2') ;
ObjFolio.SetObmFromRow(Row-GS.FixedRows, RichEdit, O) ;
GereQuantite(Ecr) ;
// Changement de référence(s) ?
ModifRefAna(Row-GS.FixedRows, 'REFINTERNE',     Null) ;
ModifRefAna(Row-GS.FixedRows, 'REFEXTERNE',     Null) ;
ModifRefAna(Row-GS.FixedRows, 'REFLIBRE',       Null) ;
ModifRefAna(Row-GS.FixedRows, 'DATEREFEXTERNE', Null) ;
ModifRefAna(Row-GS.FixedRows, 'CONSO',          Null) ;
ModifRefAna(Row-GS.FixedRows, 'AFFAIRE',        Null) ;
finally
 if O<>nil then O.Free ;
end;
END ;


procedure TFSaisieFolio.SupprimeEcrCompl ( vTOB : TOB ) ;
begin
 if ( ObjFolio.RechGuidId(vTOB) = '' ) then
  CFreeTOBCompl(vTOB)
   else
    begin
     CPutValueTOBCompl(vTOB,'EC_CUTOFFDEB',iDate1900) ;
     CPutValueTOBCompl(vTOB,'EC_CUTOFFFIN',iDate1900) ;
    end ;
end ;


procedure TFSaisieFolio.AnnuleCutOffSiBesoin( vTOB : TOB ) ;
var
 lValue : variant ;

 procedure Supp ;
 begin
  vTOB.PutValue('COMPTAG','0') ;
  vTOB.PutValue('OLDGENERAL', vTOB.GetValue('E_GENERAL') ) ;
  SupprimeEcrCompl(vTOB) ;
 end ;

begin
 lValue := CGetValueTOBCompl(vTOB,'EC_CUTOFFDEB') ;
 Case VarType(lValue) of
  varEmpty,varNull : Supp ;
  varDate : if ( lValue = iDate1900 ) then Supp ;
 end ; // if
end ;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 08/06/2005
Modifié le ... : 08/06/2005
Description .. : Valeur possible de comptag
Suite ........ : '-' : cutoff a saisir
Suite ........ : 'X' : cutoff saisie ( mais la date peu etre egale a idate1900 )
Suite ........ : '1' : reforce la saisie du cutoff ( qd on fait alt+C sur un 
Suite ........ : compte de cutoff )
Suite ........ : '0' : cutoff saisie mais annulé ( ne sera plus reproposé )
Suite ........ : 'XD' : cutoff saisie et date <> iDate1900
Suite ........ : 
Suite ........ : - FB 15876 - LG - 08/06/2005 - Lorsque j'indique 
Suite ........ : 01/01/1900, il ne faut pas mettre le triangle bleu sur le jour 
Suite ........ : dans l'écriture
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.GereCutOff( ARow : integer) ;
var
 O               : TOBM ;
 Ecr             : TZF ;
 RC              : R_COMP ;
 Action          : TActionFiche ;
 bModifBlocNote  : Boolean ;
 NumCompte       : string ;
 k               : integer ;
 lTOBCompl       : TOB ;
begin
{$IFDEF TT}
AddEvenement('GereCutOff') ;
{$ENDIF}
NumCompte:=GS.Cells[SF_GEN, ARow] ; k:=Comptes.GetCompte(NumCompte) ;
if (k=-1) then exit ;
if(Comptes.GetValue('G_CUTOFF') <> 'X') then exit ;
if ObjFolio=nil then Exit ;
Ecr:=ObjFolio.GetRow(ARow-GS.FixedRows) ; if Ecr=nil then exit ;
O:=ObjFolio.GetObmFromRow(ARow-GS.FixedRows, RichEdit) ; if O=nil then Exit ;
try
if Ecr.GetValue('E_QUALIFORIGINE') = 'CUT' then exit ; // ecriture genere par la traitement de cutoff
if Ecr.GetValue('COMPTAG') = '0' then exit ;
if Ecr.GetValue('OLDGENERAL')<>NumCompte then Ecr.PutValue('COMPTAG','-') ;
if Pos('X',Ecr.GetValue('COMPTAG')) > 0 then exit ;
RC.StLibre:='---CUTXXXXXXXXXXXXXXXXXXXXXXXX' ;
RC.StComporte:='--XXXXXXXX' ; RC.CutOffPer := Comptes.GetValue('G_CUTOFFPERIODE') ;
RC.CutOffEchue := Comptes.GetValue('G_CUTOFFECHUE') ;
Action:=taModif ; RC.AvecCalcul := true ;
lTOBCompl := CGetTOBCompl(Ecr) ;
if lTOBCompl = nil then
 lTOBCompl := CCreateTOBCompl(Ecr,ObjFolio.EcrsCompl) ;
lTOBCompl.PutValue('EC_DATECOMPTABLE', GetRowDate(ARow) ) ;
CCalculDateCutOff( lTOBCompl, Comptes.GetValue('G_CUTOFFPERIODE') , Comptes.GetValue('G_CUTOFFECHUE')) ;
RC.TOBCompl := lTOBCompl ;
RC.Attributs:=FALSE ; RC.MemoComp:=nil ; RC.Origine:=-1 ;
if not SaisieComplement(O, EcrGen, Action, bModifBlocNote, RC, False, True) then
 begin
  SupprimeEcrCompl(Ecr) ;
  Ecr.PutValue('COMPTAG','0') ;
  Ecr.PutValue('OLDGENERAL', NumCompte) ;
  exit ;
 end
 else
 begin
   // GCO - 30/05/2006 - 17837
   GrilleModif := True;
   Statut := TaModif;
 end;
Ecr.PutValue('COMPTAG','X') ;
if RC.TOBCompl.GetValue('EC_CUTOFFDEB') <> iDate1900 then
 Ecr.PutValue('COMPTAG','XD') ;
Ecr.PutValue('OLDGENERAL', NumCompte) ;
Ecr.PutValue('E_BLOCNOTE', O.GetValue('E_BLOCNOTE') ) ;
finally
 if O<>nil then O.Free ; 
end;

end ;



{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 22/02/2002
Modifié le ... : 02/04/2002
Description .. : Modif de la gestion des libelles autos : possibilité d'associé
Suite ........ : plusieur libelle ou meme journal et nature de compte.
Suite ........ : rem : le calcul de la formule du guide est fait ici pour pouvoir
Suite ........ : beneficie de l'ensemble des ecritures
Mots clefs ... :
*****************************************************************}
procedure TFSaisieFolio.RefLibAutoClick ;
var R : TRect ;  Ecr : TZF ;
lTOB : TOB ; lStOldLib : string ;
begin
{$IFDEF TT}
AddEvenement('RefLibAutoClick') ;
{$ENDIF}
if bLibAuto.Enabled=FALSE then Exit ;
Ecr:=SetRowTZF(GS.Row) ; if Ecr=nil then Exit ;
R:=GS.CellRect(GS.Col, GS.Row) ; lStOldLib:=GS.Cells[SF_LIB, GS.Row] ;
SELGUIDE.Text:='' ; SELGUIDE.Left:=R.Left ; SELGUIDE.Top:=R.Bottom+10 ; lTOB:=nil ;
try
lTOB:=FZlibAuto.RechercheEnBase(SELGUIDE,GS.Cells[SF_LIB,GS.Row],memJal.Code,Ecr.GetValue('E_NATUREPIECE'),Ecr);
if assigned(lTOB) then
 BEGIN
  if GS.Cells[SF_REFI, GS.Row]='' then
   GS.Cells[SF_REFI, GS.Row]:=GFormule(lTOB.GetValue('RA_FORMULEREF'), GetFormule, nil, 1) ;
  // Permet d'utiliser Référence interne dans la formule du libellé
  Ecr.PutValue('E_REFINTERNE', GS.Cells[SF_REFI, GS.Row]) ;
  GS.Cells[SF_LIB, GS.Row]:=GFormule(lTOB.GetValue('RA_FORMULELIB'), GetFormule, nil, 1) ;
 END;
GrilleModif:=true ;
finally
 lTOB.Free ;
end;// try
end ;

procedure TFSaisieFolio.EchClick ;
var Ecr : TZF ;
begin
if Not BEche.Visible then Exit ;
if Not BEche.Enabled then Exit ;
Ecr:=SetRowTZF(GS.Row) ; if Ecr=nil then Exit ;
GetEch(GS.Row, FALSE, FALSE) ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 31/05/2006
Modifié le ... :   /  /
Description .. : - FB 18204 - LG - l'appel de l'analytiq avec les touches de
Suite ........ : racourci fct pour tous les compte
Mots clefs ... :
*****************************************************************}
procedure TFSaisieFolio.AnaClick ;
var Ecr : TZF ;
begin
if (GS.Cells[SF_GEN, GS.Row]='') or
   ((GS.Cells[SF_DEBIT, GS.Row]='') and (GS.Cells[SF_CREDIT, GS.Row]='')) then Exit ;
Ecr:=SetRowTZF(GS.Row) ; if Ecr=nil then Exit ;
GereAna(GS.Row,true) ;
PutRowProrata(GS.Row) ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 02/12/2002
Modifié le ... : 11/07/2007
Description .. : - 02/12/2002 - fiche 10833 - on tient compte de la nature
Suite ........ : de l'ecriture
Suite ........ : pour determiner le sens dans la fenetre des devises
Suite ........ : - 24/01/2003 - FB 11846 - En saisie devise, il etait possible
Suite ........ : de saisir sur un collectif sans auxiliaire.
Suite ........ :  )
Suite ........ : - 18/07/2005 - FB 12301 - en mode bordereau on equilibre
Suite ........ : les devise sur la premiere ligne
Suite ........ : - LG - 02/11/2006 - Suite demande GPAO, on permet de 
Suite ........ : saisir une devise par défaut en multi devises
Suite ........ : 
Suite ........ : 
Suite ........ : - on equilibre le bordereau en mode libre des qu'il est solde
Suite ........ : en devise
Suite ........ : - LG - 19/10/2006 - FB 18878 - on interdit les peices 
Suite ........ : multidevise ne ehcnage vers S1
Suite ........ : - LG - 22/02/2007 - FB 17607 - on equilibre sur la derneir 
Suite ........ : ligne en mode libre
Suite ........ : - LG - 11/07/2007 - Params est passe en globale en 
Suite ........ : FPAram pour garde els taux volatile
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.DevClick(bTestTaux : Boolean) ;
var Ecr, Tmp : TZF ; CumulD, CumulC, SoldeDevise, Montant : Double ;
    SensDef, NumCompte, NumAux, CodeDev : string ; k : Integer ;
    bMonoDev : Boolean ; Action : TActionFiche ;
    ZFDevise : TZFDevise ;
    Z,NatPiece : string ; Sc : integer ;
begin
{$IFDEF TT}
AddEvenement('DevClick') ;
{$ENDIF}
SensDef:='#' ;
if (not BDevise.Enabled)  then Exit ;
Ecr:=SetRowTZF(GS.Row) ;
if Ecr=nil then Exit ;
//LG* - 12/12/2001 - Interdit le changement de devise à l'intérieur d'une pièce
if  (Ecr.GetValue('E_DEVISE')=memPivot.Code) and not bNewLigne then exit;
if Ecr.GetValue('E_IMMO')<>'' then Exit ;
if (Ecr.GetValue('E_DEVISE')=memPivot.Code) and
   ((Ecr.GetValue('E_DEBIT')<>0) or (Ecr.GetValue('E_CREDIT')<>0)) then Exit ;
if (not memJal.bDevise)  then Exit ;
NumCompte:=GS.Cells[SF_GEN, GS.Row] ;
k:=Comptes.GetCompte(NumCompte) ; if k = - 1 then exit ;
if (k>=0) and (Comptes.IsVentilable(0, k)) then BVentil.Enabled:=TRUE ;
if k<>-1 then
 begin
  if Comptes.IsCollectif(k) and (trim(GS.Cells[SF_AUX, GS.Row])='') then exit ;
 end; // if
bMonoDev:=FALSE ;
Tmp:=TZF.Create('ECRITURE', nil, -1) ;
Tmp.Assign(Ecr) ;
SoldeDevise:=SoldePieceDevise(StrToInt(GS.Cells[SF_PIECE, GS.Row]), CumulD, CumulC) ;
if ((bModeRO) or (IsRowCanEdit(GS.Row)>0)) then Action:=taConsult
  else if (CumulD=0) and (CumulC=0) then Action:=taCreat
  else Action:=taModif ;
if (bGuideRun) and (Tmp.GetValue('E_DEBITDEV')=0) and (Tmp.GetValue('E_CREDITDEV')=0) then
  begin
  if GetGuideColStop(GS.Row)=SF_DEBIT  then SensDef:='D' ;
  if GetGuideColStop(GS.Row)=SF_CREDIT then SensDef:='C' ;
  if Tmp.GetValue('EG_DEBITDEV')<>#0 then
    begin
    Montant:=Valeur(GFormule(Tmp.GetValue('EG_DEBITDEV'), GetFormule, nil, 1)) ;
    Tmp.PutValue('E_DEBITDEV',  StrFMontant(Montant,15,memPivot.Decimale,'',TRUE)) ;
    end ;
  if Tmp.GetValue('EG_CREDITDEV')<>#0 then
    begin
    Montant:=Valeur(GFormule(Tmp.GetValue('EG_CREDITDEV'), GetFormule, nil, 1)) ;
    Tmp.PutValue('E_CREDITDEV',  StrFMontant(Montant,15,memPivot.Decimale,'',TRUE)) ;
    end ;
  end ;
if (SensDef='#') and (k<>-1) then
  begin
   SensDef:='D' ; NatPiece:=GetCellNature(GS.Row) ;
   if NatPiece<>'OD' then
    begin
     Z:=Comptes.GetValue('G_SENS', k) ; if Z='' then Z:='D' ;
     case Z[1] of 'D' : SC:=1 ; 'C' : SC:=2 ; else  SC:=3 ; end ;
     if ( QuelSens(NatPiece, Comptes.GetValue('G_NATUREGENE', k), SC)) <> 1 then sensDef:='C' else SensDef:='D' ;
    end; // if
  end ;
// Gestion de la saisie mono-devise sur un tiers
if (Ecr.GetValue('E_DEVISE')=V_PGI.DevisePivot) then
  begin
  if (GS.Cells[SF_AUX, GS.Row]<>'') then
    begin
    NumAux:=GS.Cells[SF_AUX, GS.Row] ; k:=Tiers.GetCompte(NumAux) ;
    CodeDev:=Tiers.GetValue('T_DEVISE', k) ;
    // LG - 02/11/2006 - Suite demande GPAO, on permet de saisir une devise par défaut en multi devises
    if (CodeDev=V_PGI.DevisePivot) and (k>=0) and (Tiers.GetValue('T_MULTIDEVISE', k)='-') then begin Tmp.Free ; Exit ; end ;
    //if (k>=0) and (Tiers.GetValue('T_MULTIDEVISE', k)='-') and (CodeDev<>'') then
     if (k>=0) and (CodeDev<>'') then
      begin
      Tmp.PutValue('E_DEVISE', CodeDev) ;
      bMonoDev:=(Tiers.GetValue('T_MULTIDEVISE', k)='-') ; //TRUE ;
      end ;
    end
     else
      begin
       if Comptes.IsBqe(k) then
        CodeDev:=GetDeviseBanque(NumCompte) ;
       if (CodeDev<>'') then
        begin
         Tmp.PutValue('E_DEVISE', CodeDev) ;
         bMonoDev:=TRUE ;
        end ;
      end ;
  end ;
if memParams.bDevise then
  begin
  bMonoDev:=TRUE ; Tmp.PutValue('E_DEVISE', memParams.DevLibre) ;
  end ;
FParams.Ecr:=Tmp ;
FParams.DPivot:=memPivot ;
FParams.TD:=CumulD ;
FParams.TC:=CumulC ;
FParams.Solde:=SoldeDevise ;
FParams.SensDef:=SensDef ;
FParams.bMonoDev:=bMonoDev ;
FParams.bDoSolde:=bGuideDoSolde ; bGuideDoSolde:=FALSE ;
FParams.Action:=Action ;
if bTestTaux then
  begin
  ZFDevise:=TZFDevise.Create(Application, FParams) ;
  memDevise.Code:=Ecr.GetValue('E_DEVISE') ;
  GetInfosDevise(memDevise) ;
  memDevise.Taux:=GetTaux(memDevise.Code, memDevise.DateTaux, Ecr.GetValue('E_DATECOMPTABLE')) ;
  if (memDevise.DateTaux<>Ecr.GetValue('E_DATECOMPTABLE')) or ZFDevise.PbTaux(memDevise, Ecr.GetValue('E_DATECOMPTABLE'))
    then ZFDevise.AvertirPbTaux(memDevise.Code, memDevise.DateTaux, Ecr.GetValue('E_DATECOMPTABLE'), TRUE) ;
  ZFDevise.Free ;
  memDevise.Taux:=GetTaux(memDevise.Code, memDevise.DateTaux, Ecr.GetValue('E_DATECOMPTABLE')) ;
  ChangeTauxPiece(StrToInt(GS.Cells[SF_PIECE, GS.Row]), memDevise) ;
  GetRowTZF(GS.Row-1) ;
  SoldeDevise:=SoldePieceDevise(StrToInt(GS.Cells[SF_PIECE, GS.Row]),  CumulD, CumulC) ;
  //LG*1 06/12/2001 en mode libre et en saisie en contrevaleur on equilibre au bordereau et non pas a la piece
  if (SoldeDevise=0) and (not memParams.bLibre) then VerifEquilibre(StrToInt(GS.Cells[SF_PIECE, GS.Row]), GS.Row) ;
  if (memParams.bLibre) or ( GetParamsocSecur('SO_CPLIENGAMME','') = 'S1' )  then
    begin
    memParams.bDevise:=TRUE ;
    //LG*1 26/11/2001 on ne repasse pas en mode bordereau apres une saisie en euro dans un bordereau en mode libre
    //memParams.bLibre:=FALSE ;
    memParams.DevLibre:=Tmp.GetValue('E_DEVISE') ;
    end ;
  CalculSoldeJournal ;
  end else
  begin
  if SaisieZDevise(FParams) then
     begin
     Ecr.Assign(Tmp) ;
     GetRowTZF(GS.Row-1) ;
     SoldeDevise:=SoldePieceDevise(StrToInt(GS.Cells[SF_PIECE, GS.Row]),  CumulD, CumulC) ;
     //LG*1 06/12/2001 en mode libre et en saisie en contrevaleur on equilibre au bordereau et non pas a la piece
     if (FFirstRow = 0) and (memParams.bLibre) then
      FFirstRow := nbLignes
      else
       FFirstRow := GS.Row ;

     if (SoldeDevise=0) then VerifEquilibre(StrToInt(GS.Cells[SF_PIECE, GS.Row]), FFirstRow) ;
     if (memParams.bLibre)  or ( GetParamsocSecur('SO_CPLIENGAMME','') = 'S1' ) then
       begin
       memParams.bDevise:=TRUE ;
       memParams.DevLibre:=Tmp.GetValue('E_DEVISE') ;
       end ;
     end ;
  end ;
bStopDevise:=true ;
Tmp.Free ;
GS.Invalidate ;

end ;

procedure TFSaisieFolio.RibClick ;
var Ecr : TZF ; Obm : TObm ;
    lBoAuxiliaire : Boolean;
begin
if not BModifRIB.Enabled then Exit ;
if ObjFolio=nil then Exit ;
Ecr:=ObjFolio.GetRow(GS.Row-GS.FixedRows) ; if Ecr=nil then Exit ;
Obm:=ObjFolio.GetObmFromRow(GS.Row-GS.FixedRows, RichEdit) ;
if Obm=nil then Exit ;

// GCO - 30/05/2006 - FQ 16233 - pb avec E_AUXILIAIRE qui vaut #0 donc <> ''
lBoAuxiliaire := Length(Trim(Ecr.GetString('E_AUXILIAIRE'))) > 0;

ModifRIBOBM(Obm, FALSE, FALSE, '', lBoAuxiliaire) ;
Ecr.PutValue('E_RIB',OBM.GetValue('E_RIB')) ;
if Obm<>nil then Obm.Free ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 06/09/2002
Modifié le ... :   /  /
Description .. : - 06/10/2002 - on relit le bordereau suite à la suppression
Suite ........ : de l'immo
Mots clefs ... :
*****************************************************************}
procedure TFSaisieFolio.ImmoClick ;
var Ecr : TZF ;
    CodeImmo : string ;
{$IFDEF COMPTAAVECSERVANT}
    i : Integer ;
{$ENDIF}
begin
  if not BZoomImmo.Enabled then Exit ;
  Ecr:=ObjFolio.GetRow(GS.Row-GS.FixedRows) ; if Ecr=nil then Exit ;
  CodeImmo:=Ecr.GetValue('E_IMMO') ;
  if CodeImmo='' then Exit ;
  { CA - 09/06/2006 - FQ 18340 }
  if Assigned(TheMulQ) then
    TheMulQ := nil;

{$IFDEF COMPTAAVECSERVANT}
  IF VHIM^.ServantPGI Then
  BEGIN

  For i:=1 To Length(CodeImmo) Do If CodeImmo[i] in ['0'..'9']=FALSE Then Exit ;
  if Not Presence('IMOREF','IRF_NUM',CodeImmo) then PrintMessageRC(RC_NOEXISTIMMO)
                                               else VoirFicheImo(taConsult,StrToInt(CodeImmo)) ;
  END Else
  BEGIN
{$IFDEF AMORTISSEMENT}
    if not Presence('IMMO', 'I_IMMO', CodeImmo) then PrintMessageRC(RC_NOEXISTIMMO)
    else
    begin
      AMLanceFiche_FicheImmobilisation(CodeImmo, taConsult, '') ;
      if not Presence('IMMO', 'I_IMMO', CodeImmo) then AfterZoom(GS.Col,GS.Row) ; // on relit le bordereau suite à la suppression de l'immo
    end;
{$ENDIF}
  END;
{$ENDIF}
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 27/05/2002
Modifié le ... : 10/01/2006
Description .. : on n'ouvre pas les guides spécifique à la saisie de trésorerie
Suite ........ : - LG - 03/08/2004 - FB 11850 - il manquait un test sur les 
Suite ........ : guides de treso
Suite ........ : - LG - 28/10/2004 - recrture de la fct avec un existesql + on 
Suite ........ : stocke le der guide pour ne pas relancer unitilment le 
Suite ........ : traitement
Suite ........ : - LG - 19/07/2005 - FB 14441 - lancasi une requete avec 
Suite ........ : un code auxi egale a #0
Suite ........ : GU_NATUREPIECE="
Mots clefs ... : 
*****************************************************************}
function TFSaisieFolio.ExisteGuide : Boolean ;
var Critere, Table, NumCompte, NumAux : string ;
   // Q : TQuery ;
begin
{$IFDEF TT}
AddEvenement('ExisteGuide ') ;
{$ENDIF}
Result:=FALSE ;
if bModeRO then Exit ;
if not bGuideJal then Exit ;
NumAux:=GS.Cells[SF_AUX, GS.Row] ; NumCompte:=GS.Cells[SF_GEN, GS.Row] ;
if length(trim(NumAux))<>0 then     
  begin
  NumAux:=BourreLaDonc(NumAux, fbAux) ;
  Critere:='GU_TYPE="NOR" AND GU_TRESORERIE<>"X" AND GU_JOURNAL="'+memJal.Code+'" AND EG_AUXILIAIRE="'+NumAux+'"' ;
  if E_ETABLISSEMENT.Value<>'' then Critere:=Critere+' AND (GU_ETABLISSEMENT="'+E_ETABLISSEMENT.Value+'" OR GU_ETABLISSEMENT="")' ;
  if GS.Cells[SF_NAT, GS.Row]<>'' then
    Critere:=Critere+' AND ( GU_NATUREPIECE="'+GetCellNature(GS.Row)+'" OR GU_NATUREPIECE="" ) ' ;
  if (IsJalLibre) and (GS.Row>1) then
    begin
    if (GS.Cells[SF_DEVISE, GS.Row]<>'') then Critere:=Critere+' AND GU_DEVISE="'+GS.Cells[SF_DEVISE, GS.Row]+'"'
                                         else Critere:=Critere+' AND GU_DEVISE="'+V_PGI.DevisePivot+'"'
    end ;
  Table:='GUIDE LEFT JOIN ECRGUI ON GU_GUIDE=EG_GUIDE' ;
  end else
  if NumCompte<>'' then
    begin
    NumCompte:=BourreLaDonc(NumCompte, fbGene) ;
    Critere:='GU_TYPE="NOR" AND GU_TRESORERIE<>"X" AND GU_JOURNAL="'+memJal.Code+'" AND EG_GENERAL="'+NumCompte+'"' ;
    if E_ETABLISSEMENT.Value<>'' then Critere:=Critere+' AND (GU_ETABLISSEMENT="'+E_ETABLISSEMENT.Value+'" OR GU_ETABLISSEMENT="")' ;
    if GS.Cells[SF_NAT, GS.Row]<>'' then
      Critere:=Critere+' AND ( GU_NATUREPIECE="'+GetCellNature(GS.Row)+'" OR GU_NATUREPIECE="" ) ' ;
    if (IsJalLibre) and (GS.Row>1) then
      begin
      if (GS.Cells[SF_DEVISE, GS.Row]<>'') then Critere:=Critere+' AND GU_DEVISE="'+GS.Cells[SF_DEVISE, GS.Row]+'"'
                                           else Critere:=Critere+' AND GU_DEVISE="'+V_PGI.DevisePivot+'"'
      end ;
    Table:='GUIDE LEFT JOIN ECRGUI ON GU_GUIDE=EG_GUIDE' ;
    end else
      begin
      Critere:='GU_TYPE="NOR" AND GU_TRESORERIE<>"X" AND GU_JOURNAL="'+memJal.Code+'"' ;
      if E_ETABLISSEMENT.Value<>'' then Critere:=Critere+' AND (GU_ETABLISSEMENT="'+E_ETABLISSEMENT.Value+'" OR GU_ETABLISSEMENT="")' ;
      if GS.Cells[SF_NAT, GS.Row]<>'' then
        Critere:=Critere+' AND ( GU_NATUREPIECE="'+GetCellNature(GS.Row)+'" OR GU_NATUREPIECE="" ) ' ;
      if (IsJalLibre) and (GS.Row>1) then
        begin
        if (GS.Cells[SF_DEVISE, GS.Row]<>'') then Critere:=Critere+' AND GU_DEVISE="'+GS.Cells[SF_DEVISE, GS.Row]+'"'
                                             else Critere:=Critere+' AND GU_DEVISE="'+V_PGI.DevisePivot+'"'
        end ;
      Table:='GUIDE' ;
      end ;
 if FStLastTextGuide <> 'SELECT GU_GUIDE, GU_ETABLISSEMENT FROM '+Table+' WHERE '+Critere  then
  begin
   FStLastTextGuide   := 'SELECT GU_GUIDE, GU_ETABLISSEMENT FROM '+Table+' WHERE '+Critere ;
   result             := ExisteSQL(FStLastTextGuide) ;
   FStLastResultGuide := result ;
  end
   else
    result := FStLastResultGuide ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 27/05/2002
Modifié le ... : 10/01/2006
Description .. : on n'ouvre pas les guides spécifique à la saisie de trésorerie
Suite ........ : -20/01/2002 - fiche 10791 - on n'ouvre pas les guides si la 
Suite ........ : nature n'est pas renseigne
Suite ........ : - 19/10/2004 - FB 11850 - rafraissisement de la case nature 
Suite ........ : avant d'appeller les guides
Suite ........ : - 10/01/2005 - FB 10772 - laisser la possibilite que la zone 
Suite ........ : nature ne comporte aucune valeur
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.GuiClick ;
var Q : TQuery ; R : TRect ; Critere, Table, NumCompte, NumAux : string ;
begin
if GS.Col = SF_NAT then PostMessage(GS.Handle, WM_KEYDOWN, VK_TAB, 0) ;
EnableButtons ;
if BTGuide.Enabled=FALSE then Exit ; if GS.Cells[SF_NAT,GS.Row]='' then exit ;
R:=GS.CellRect(GS.Col, GS.Row) ;
SELGUIDE.Text:='' ; SELGUIDE.Left:=R.Left ; SELGUIDE.Top:=R.Bottom+10 ;
NumAux:=GS.Cells[SF_AUX, GS.Row] ; NumCompte:=GS.Cells[SF_GEN, GS.Row] ;
if NumAux<>'' then
  begin
  NumAux:=BourreLaDonc(NumAux, fbAux) ;
  Critere:='GU_TYPE="NOR" AND GU_TRESORERIE<>"X" AND GU_JOURNAL="'+memJal.Code+'" AND EG_AUXILIAIRE="'+NumAux+'"' ;
  if E_ETABLISSEMENT.Value<>'' then Critere:=Critere+' AND (GU_ETABLISSEMENT="'+E_ETABLISSEMENT.Value+'" OR GU_ETABLISSEMENT="")' ;
  if GS.Cells[SF_NAT, GS.Row]<>'' then
    Critere:=Critere+' AND ( GU_NATUREPIECE="'+GetCellNature(GS.Row)+'" OR GU_NATUREPIECE="" )' ;
  if (IsJalLibre) and (GS.Row>1) then
    begin
    if (GS.Cells[SF_DEVISE, GS.Row]<>'') then Critere:=Critere+' AND GU_DEVISE="'+GS.Cells[SF_DEVISE, GS.Row]+'"'
                                         else Critere:=Critere+' AND GU_DEVISE="'+V_PGI.DevisePivot+'"'
    end ;
  Table:='GUIDE LEFT JOIN ECRGUI ON GU_GUIDE=EG_GUIDE' ;
  // Un seul guide ?
  Q:=OpenSQL('SELECT GU_GUIDE, GU_ETABLISSEMENT FROM '+Table+' WHERE '+Critere+' ORDER BY GU_ETABLISSEMENT DESC', TRUE) ;
  if not Q.EOF then SELGUIDE.Text:=Q.Fields[0].AsString ;
  Q.Next ; if not Q.EOF then SELGUIDE.Text:='' ;
  Ferme(Q) ;
  if SELGUIDE.Text='' then
    LookupList(SELGUIDE,TraduireMemoire('Guide'),Table,'GU_GUIDE','GU_LIBELLE',Critere,'GU_LIBELLE',TRUE,0) ;
  end else
  if NumCompte<>'' then
    begin
    NumCompte:=BourreLaDonc(NumCompte, fbGene) ;
    Critere:='GU_TYPE="NOR" AND GU_TRESORERIE<>"X" AND GU_JOURNAL="'+memJal.Code+'" AND EG_GENERAL="'+NumCompte+'"' ;
    if E_ETABLISSEMENT.Value<>'' then Critere:=Critere+' AND (GU_ETABLISSEMENT="'+E_ETABLISSEMENT.Value+'" OR GU_ETABLISSEMENT="")' ;
    if GS.Cells[SF_NAT, GS.Row]<>'' then
      Critere:=Critere+' AND ( GU_NATUREPIECE="'+GetCellNature(GS.Row)+'" OR GU_NATUREPIECE="" )' ;
    if (IsJalLibre) and (GS.Row>1) then
      begin
      if (GS.Cells[SF_DEVISE, GS.Row]<>'') then Critere:=Critere+' AND GU_DEVISE="'+GS.Cells[SF_DEVISE, GS.Row]+'"'
                                           else Critere:=Critere+' AND GU_DEVISE="'+V_PGI.DevisePivot+'"'
      end ;
    Table:='GUIDE LEFT JOIN ECRGUI ON GU_GUIDE=EG_GUIDE' ;
    // Un seul guide ?
    Q:=OpenSQL('SELECT GU_GUIDE, GU_ETABLISSEMENT FROM '+Table+' WHERE '+Critere+' ORDER BY GU_ETABLISSEMENT DESC', TRUE) ;
    if not Q.EOF then SELGUIDE.Text:=Q.Fields[0].AsString ;
    Q.Next ; if not Q.EOF then SELGUIDE.Text:='' ;
    Ferme(Q) ;
    if SELGUIDE.Text='' then
      LookupList(SELGUIDE,TraduireMemoire('Guide'),Table,'GU_GUIDE','GU_LIBELLE',Critere,'GU_LIBELLE',TRUE,0) ;
    end else
      begin
      Critere:='GU_TYPE="NOR" AND GU_TRESORERIE<>"X" AND GU_JOURNAL="'+memJal.Code+'"' ;
      if E_ETABLISSEMENT.Value<>'' then Critere:=Critere+' AND (GU_ETABLISSEMENT="'+E_ETABLISSEMENT.Value+'" OR GU_ETABLISSEMENT="")' ;
      if GS.Cells[SF_NAT, GS.Row]<>'' then
        Critere:=Critere+' AND ( GU_NATUREPIECE="'+GetCellNature(GS.Row)+'" OR GU_NATUREPIECE="" )' ;
      if (IsJalLibre) and (GS.Row>1) then
        begin
        if (GS.Cells[SF_DEVISE, GS.Row]<>'') then Critere:=Critere+' AND GU_DEVISE="'+GS.Cells[SF_DEVISE, GS.Row]+'"'
                                             else Critere:=Critere+' AND GU_DEVISE="'+V_PGI.DevisePivot+'"'
        end ;
      // Un seul guide ?
      Q:=OpenSQL('SELECT GU_GUIDE, GU_ETABLISSEMENT FROM GUIDE WHERE '+Critere+' ORDER BY GU_ETABLISSEMENT DESC', TRUE) ;
      if not Q.EOF then SELGUIDE.Text:=Q.Fields[0].AsString ;
      Q.Next ; if not Q.EOF then SELGUIDE.Text:='' ;
      Ferme(Q) ;
      if SELGUIDE.Text='' then
        LookupList(SELGUIDE,TraduireMemoire('Guide'),'GUIDE','GU_GUIDE','GU_LIBELLE',Critere,'GU_LIBELLE',TRUE,0) ;
      end ;
FLastNatureGuide:=GS.Cells[SF_NAT,GS.Row] ;
GrilleModif:=true ;
if SELGUIDE.Text<>'' then CGuideRun ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 13/06/2002
Modifié le ... : 10/11/2004
Description .. : - 13/06/2002 - on ne s'arrete plus sur le la case nature mais
Suite ........ : sur la case general
Suite ........ : - 22/01/2003 - FB 10791 - on controle la coherence du
Suite ........ : guide precedent avec la nature courante
Suite ........ : - 22/01/2003 - FB 10791 - ne pas propose de guide si la
Suite ........ : zone est vide
Suite ........ : - 10/11/2004 - FB 11613 - ajout du nom du guide sur la 
Suite ........ : flashlabel
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.CGuideRun ;
var  ObjGuide : TZGuide ; Ecr : TZF ; i : Integer ; Tmp: string ;
s , lStLib : string ;
begin
if ObjFolio=nil then Exit ;
if ( (FLastNatureGuide<>GS.Cells[SF_NAT,GS.Row]) and (InitParams.ParGuide='') ) or (trim(SELGUIDE.Text)='')  then exit ;
{$IFDEF TT}
AddEvenement('CGuideRun') ;
{$ENDIF}
HGBeginUpdate(GS) ;  AfficheTitreAvecCommeInfo('Création du guide en cours...');
ObjGuide:=TZGuide.Create(FRechCompte.Info) ; //('GuideEcr', SELGUIDE.Text) ;
try ;
ObjGuide.Load(['NOR',SELGUIDE.Text]) ;
GuideFirstRow:=GS.Row ;
for i:=0 to ObjGuide.Item.Detail.Count-1 do
    begin
    if i=0 then
     begin
     bGuideRun:=TRUE ; bGuideStop:=FALSE ;
     //GS.Cells[SF_REFI, GS.Row]:='' ;
     end ;
    if i>0 then CreateRow(GS.Row+i) ;
    Ecr:=ObjFolio.GetRow(GS.Row-GS.FixedRows+i) ; if Ecr=nil then break ;
    Ecr.PutValue('GUIDE',         'X') ;
    Ecr.PutValue('E_QUALIFORIGINE', 'GUI') ;
    Tmp:=ObjGuide.GetValue('GU_DEVISE') ; if Tmp = #0 then exit ;
    Ecr.PutValue('EG_DEVISE',     Tmp) ;
    Ecr.PutValue('E_DEVISE',      Tmp) ;
    if Tmp<>V_PGI.DevisePivot then GS.Cells[SF_DEVISE, GS.Row+i]:=Tmp ;
    GS.Cells[SF_EURO, GS.Row+i]:= '-' ;
    Ecr.PutValue('EG_NUMLIGNE',   ObjGuide.Item.Detail[i].GetValue('EG_NUMLIGNE')) ;
    Ecr.PutValue('EG_GENERAL',    ObjGuide.Item.Detail[i].GetValue('EG_GENERAL')) ;
    Ecr.PutValue('EG_AUXILIAIRE', ObjGuide.Item.Detail[i].GetValue('EG_AUXILIAIRE')) ;
    Ecr.PutValue('EG_REFINTERNE', ObjGuide.Item.Detail[i].GetValue('EG_REFINTERNE')) ;
    Ecr.PutValue('EG_LIBELLE',    ObjGuide.Item.Detail[i].GetValue('EG_LIBELLE')) ;
    Ecr.PutValue('EG_DEBITDEV',   ObjGuide.Item.Detail[i].GetValue('EG_DEBITDEV')) ;
    Ecr.PutValue('EG_CREDITDEV',  ObjGuide.Item.Detail[i].GetValue('EG_CREDITDEV')) ;
    Ecr.PutValue('EG_ARRET',      ObjGuide.Item.Detail[i].GetValue('EG_ARRET')) ;
    lStLib := ObjGuide.GetValue('GU_LIBELLE') ;
    GetGuideCalcRow(GS.Row+i) ;  
    if (Ecr.GetValue('EG_GENERAL')<>'') and (not FRechCompte.Info.LoadCompte(Ecr.GetValue('EG_GENERAL'))) then
     begin
      PGIInfo('Le compte ' + Ecr.GetValue('EG_GENERAL') + ' n''existe pas !',Caption);
      s:=Ecr.GetValue('EG_ARRET') ; s[1]:='X' ; Ecr.PutValue('EG_ARRET',s) ;
     end; // if
    if (Ecr.GetValue('EG_AUXILIAIRE')<>'') and (not FRechCompte.Info.LoadAux(Ecr.GetValue('EG_AUXILIAIRE'))) then
     begin
      PGIInfo('L''auxiliaire ' + Ecr.GetValue('EG_AUXILIAIRE') + ' n''existe pas !',Caption);
      s:=Ecr.GetValue('EG_ARRET') ; s[2]:='X' ; Ecr.PutValue('EG_ARRET',s) ;
     end; // if
    end ;
GuideLastRow:=GuideFirstRow+ObjGuide.Items.Detail.Count-1 ;
FlashGuide.Caption := TraduireMemoire('Guide : ' + lStLib ) ;
FlashGuide.Visible:=TRUE ; FlashGuide.Flashing:=TRUE ;
GS.Row:=GuideFirstRow ; GS.Col:=SF_GEN ;
FCurrentRowEnter:=-1 ;
EnableButtons ;
finally
 HGEndUpdate(GS) ; AfficheTitreAvecCommeInfo ;
 ObjGuide.Free ;
end;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 11/08/2004
Modifié le ... :   /  /    
Description .. : - FB 14100 - controle de tout le guide qd on a fins ou 
Suite ........ : abandonner celui ci
Mots clefs ... :
*****************************************************************}
procedure TFSaisieFolio.GuideStop ;
begin
{$IFDEF TT}
 AddEvenement('GuideStop') ;
{$ENDIF}
 bGuideRun:=FALSE ;
 FinGuide ;
 FlashGuide.Flashing:=FALSE ; FlashGuide.Visible:=FALSE ;
 GuideFirstRow:=-1 ; GuideLastRow:=-1 ; GuideCalcRow:=-1 ; GuideCalcCol:=-1 ;
 EnableButtons ; 
end ;


procedure TFSaisieFolio.GetFolioClick ;
var ParPeriode, ParCodeJal, ParNumFolio : string ; ARow: LongInt ;
begin
if not BSelectFolio.Enabled then Exit ;
{$IFNDEF CCS3}
FBoGetFolioClick := true ;
try
ARow:=GS.Row ;
if GetFromCentralisateur('N', ParPeriode, ParCodeJal, ParNumFolio, E_DATECOMPTABLE) then
  begin
  if not ValClick(TRUE, TRUE) then
    begin
    if (ARow>1) and (ARow>nbLignes) then
      begin CreateRow(ARow) ; GS.Row:=ARow ; GS.Invalidate ; end ;
    Exit ;
    end ;
  AlimenteEntete(ParPeriode, ParCodeJal, ParNumFolio) ;
  ReadFolio(FALSE) ;
  CalculSoldeJournal ;
  // Activer le Grid
  SetGridOn ;
  EnableButtons ;
//  GS.Enabled:=TRUE ; GS.Row:=GS.FixedRows+nbLignes ; GS.SetFocus ; GSEnter(nil) ;
  end ;
finally
FBoGetFolioClick := false ;
end ;
{$ENDIF}
end ;

//=======================================================
//============= Evénements des boutons Zoom==============
//=======================================================
procedure TFSaisieFolio.BMenuZoomMouseEnter(Sender: TObject) ;
begin
PopZoom97(BMenuZoom, POPZ) ;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 13/08/2004
Modifié le ... : 28/03/2006
Description .. : - FB 13637 - LG - 13/08/2004 - Ajout du debrayable de 
Suite ........ : l'accelerateur de saisie
Suite ........ : - FB 14295 - LG - Maintenant on n'a plus de bordereau 
Suite ........ : déséquilibré mais pendant la validation, on voit la 
Suite ........ : modification apparaître une fraction de seconde. (cette 
Suite ........ : modif n'est pas conservée)
Suite ........ : Pourrait-on aussi bloquer le clic droit?
Suite ........ : - FB 17715 - 28/03/2006 - on bloque le menu qd on test un 
Suite ........ : guide
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.POPSPopup(Sender: TObject) ;
var
 T : TMenuItem ;
 NumAux : string ;
 k : integer ;
begin
 if FBoValEnCours then exit ;
 if InitParams.ParGuide <> '' then exit ;
 EnableButtons ;
 InitPopUp(Self) ;
 if memJal.Accelerateur and not bModeRO then
  begin
   T:=CreerLigPop(BAccactif,PopS,True,'P') ;
   T.Checked := FBoArretAcc ;
   PopS.Items.Insert(0,T) ;
   NumAux:=GS.Cells[SF_AUX, GS.Row] ; k:=Tiers.GetCompte(NumAux) ;
   if (k>=0) and (tiers.getValue('YTC_ACCELERATEUR')='-') then
    begin
     T:=CreerLigPop(BAccTiers,PopS,True,'P') ;
     PopS.Items.Insert(1,T) ;
    end ; // if
  end ;
end ;

procedure TFSaisieFolio.BZoom_FB19024Click(Sender: TObject) ;
begin
if not BZoom_FB19024.Enabled then Exit ;
ZoomCpte ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 04/01/2006
Modifié le ... :   /  /
Description .. : - FB 17249 - LG - 04/01/2005 - mise a jour
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.BZoomEcrsClick(Sender: TObject);
//var ACol, ARow: LongInt ;
begin
if not BZoomEcrs.Enabled then Exit ;
BZoomCpteClick(nil) ;
//if BeforeZoom(ACol, ARow) then begin ZoomEcrs ; AfterZoom(ACol, ARow) ; end ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 23/12/2004
Modifié le ... : 04/01/2006
Description .. : - FB 14307 - 23/12/2004 - rafraisissement des comptes 
Suite ........ : apers l'appel du zoom sur les journaux
Suite ........ : - FB 14307 - 19/05/2005 - remise a zero des infos ur les 
Suite ........ : journaux
Suite ........ : - FB 17250- LG - 04/01/2005 - mise a jour
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.BZoomJournalClick(Sender: TObject) ;
begin
if not BZoomJournal.Enabled then Exit ;
//if BeforeZoom(ACol, ARow) then begin
ZoomJal ; //AfterZoom(ACol, ARow) ; end ;
ChargeJal(E_JOURNAL.Value,true) ;
//FRechCompte.Info.ClearJal ;
end ;
                                  
procedure TFSaisieFolio.BZoomTiersClick(Sender: TObject) ;
begin
if not BZoomTiers.Enabled then Exit ;
ZoomTiers ;
end ;

procedure TFSaisieFolio.BZoomDeviseClick(Sender: TObject) ;
begin ZoomDevise ; end ;

procedure TFSaisieFolio.BZoomEtablClick(Sender: TObject) ;
begin ZoomEtabl ; end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 20/11/2002
Modifié le ... : 28/05/2007
Description .. : -20/11/2002- on ne relisait pas le bordereau apres la
Suite ........ : consultation des ecritures
Suite ........ : - LG - 25/06/2003 - consultation des ecritures de type L
Suite ........ : - LG - 21/03/2005 - FB 15342 - suppression du deblocage
Suite ........ : du lettrage. Fait au moment de l'enregistrement
Suite ........ : - LG -13/09/2005 - FB 13151- on ne peut plus consulte un
Suite ........ : bordereau non equilibre, saus la der ligne
Suite ........ : - LG - 10/05/2006 - FB 17715 - on desactive la fct qd on
Suite ........ : test les guides
Suite ........ : - LG - 06/11/2006 - FB 16744 - on remest a jour les info du
Suite ........ : compte apres une consultation
Suite ........ : - LG - 16/04/2007 - FB 15709 - pour le MAJ+F5 sur la
Suite ........ : derniere ligne, on supprime celle ci et on la re affcihe par la
Suite ........ : suite
Suite ........ : - LG - 22/05/2007 - FB 15709 - on doit saisir un montant
Suite ........ : pour visualiser la 1rer ligne
Suite ........ : - LG - 28/05/2007 - on passais ds les deux test, cela ne fct 
Suite ........ : pas sur la 1ere ligne
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.BZoomCpteClick(Sender: TObject);
var ACol, ARow: LongInt ;
 i : integer ;
 Cancel,Chg : boolean ;
 lStJournal : string ;
 lDtDateComptable : TDateTime ;
 lInNumeroPiece : integer ;
 lEcrRef : TZF ;
 lEcr : TZF ;
 lTError : TRecError ;
begin
EnableButtons ; Application.ProcessMessages ;
if FboFenLett or FBoValEnCours or ( InitParams.ParGuide <> '' ) then exit ;
FboFenLett := true ;
lEcrRef:=nil ;
HGBeginUpdate(GS) ;
try
if GS.Row=nbLignes then
 begin
  SetRowTZF(nbLignes) ;
  lEcr:=ObjFolio.GetRow(GS.Row-GS.FixedRows) ;
  lTError := CIsValidLigneSaisie(lEcr ,FRechCompte.Info) ;
  if ( lTError.RC_Error <> RC_PASERREUR ) and ( nbLignes > 1 ) then
   begin
    lEcrRef:= TZF.Create('ECRITURE',nil,-1);
    lEcrRef.Dupliquer(lEcr,true,true);
    DeleteRow(nbLignes) ;
    GS.Row:=nbLignes ;
   end
    else
     if ( lTError.RC_Error = RC_MONTANTINEXISTANT ) and ( nbLignes = 1 ) then
      begin
       PGIInfo('Veuillez saisir un montant pour visualiser les écritures du compte sur la première ligne') ;
       exit ;
      end ;
 end ;
ACol:=GS.Col ; ARow:=GS.Row ; Cancel:=FALSE ; Chg:=FALSE ;
GSCellExit(GS, ACol, ARow, Cancel) ; if Cancel then exit ;
if lECrRef=nil then
begin

GSRowExit(GS, ARow, Cancel, Chg) ; if Cancel then exit ;
end ;
if GS.Cells[SF_PIECE, ARow] = '' then exit ;
if (not memParams.bLibre) and (not IsSoldePiece(StrToInt(GS.Cells[SF_PIECE, ARow]))) and ( ARow<>nbLignes) then
 begin
  PrintMessageRC(RC_NONSOLDER) ;
  exit ;
 end ;
//if (GS.Cells[SF_DEBIT,  ARow]='') and (GS.Cells[SF_CREDIT, ARow]='') then exit ;
BZoomCpte.Enabled := false ;
for i:=1 to GS.RowCount - 1 do GS.Cells[SF_QUALIF, i] := 'L' ;
 try
  lStJournal := E_JOURNAL.Value ;
  lDtDateComptable := StrToDate(CurPeriode) ;
  lInNumeroPiece := StrToInt(E_NUMEROPIECE.Text) ;
  FBoLettrageEnCours:=true ;
  if not ValideFolio(FALSE, FALSE, TRUE) then
   begin
    FBoLettrageEnCours:=false ;
    exit ;
   end ;
  FBoLettrageEnCours:=true ;
  if not CBlocageLettrage(lStJournal , lDtDateComptable , lInNumeroPiece ) then exit ;
  if lEcrRef <> nil then  // sur la derniere ligne du bordereau, on a recup les info ds une tob que l'on utilise pour afficher les info du compte
   ZoomNavCpte(lEcrRef.GetValue('e_general'), lEcrRef.GetValue('e_auxiliaire') )
    else
     ZoomNavCpte ;
 finally
//   CDeBlocageLettrageP(lStJournal , lDtDateComptable , lInNumeroPiece ) ;
   ReadFolio(FALSE) ;
   for i:=1 to GS.RowCount - 1 do
     GS.Cells[SF_QUALIF, i] := 'N' ;
   memOptions.QualifP:='N' ;
   ChargeJal(memJal.Code) ;
   SetGridOn ;
   GS.Col:=ACol ; GS.Row:=ARow ;
   SetGridOptions(GS.Row) ;
   EnableButtons ;
   InfosPied(GS.Row) ;
   if (ARow<nbLignes) and (not IsRowValid(nbLignes, ACol, FALSE))then
    DeleteRow(nbLignes) ;
   if lEcrRef<>nil then
    begin
     GS.Row:=nbLignes+1 ;
     CreateRow(GS.Row) ;
     lEcr:=ObjFolio.GetRow(GS.Row-1) ;
     if lEcr<>nil then
      begin
       GS.Col:=ACol;
       lEcr.Dupliquer(lEcrRef,true,true);
       GetRowTZF(GS.Row-1) ;
      end ;
    end ;
 end;
 finally
  FreeAndNil(lEcrRef) ;
  FboFenLett := false ;
  Application.ProcessMessages ;
  HGEndUpdate(GS) ;
 end ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 30/04/2002
Modifié le ... : 31/01/2003
Description .. : On ne relit plus le bordereau apres un zoom, ces fonctions
Suite ........ : ne
Suite ........ : peuvent plus modifie le bordereau
Suite ........ :
Suite ........ : - 31/01/2003 - FB 11901 - on relit le bordereau
Mots clefs ... :
*****************************************************************}
{$IFDEF AMORTISSEMENT}
procedure TFSaisieFolio.AfterZoom(ACol, ARow : LongInt ) ;
begin
//if bWrited then // Variable globale (de la classe) qui indique si le UPDATE a été fait !
if memParams.bRealTime then
  begin
  ReadFolio(FALSE) ;
  CalculSoldeJournal ;
  SetGridOn ;
  GS.Col:=ACol ; GS.Row:=ARow ;
  SetGridOptions(GS.Row) ;
  EnableButtons ;
  if (ARow<nbLignes) and (not IsRowValid(nbLignes, ACol, FALSE))then
   DeleteRow(nbLignes) ;
  end ;
end ;
{$ENDIF}
{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 20/02/2003
Modifié le ... : 01/06/2006
Description .. : - 20/02/2003 - FB 11878 - remise à zero comptes stockes
Suite ........ : pour prendre en compte les modifs
Suite ........ : - 01/06/2006 - FB 18189 - violation d'acces ds le traitement 
Suite ........ : d'ano dyna suite  la visu d'une compte
Mots clefs ... :
*****************************************************************}
procedure TFSaisieFolio.ZoomCpte ;
var Action : TActionFiche ;
begin
if (memJal.Code='') or (GS.Cells[SF_GEN, GS.Row]='') then Exit ;
if FAction = taCreat then
 Action := taModif
  else
   Action:=FAction ;
if not ExJaiLeDroitConcept(TConcept(ccGenModif), FALSE) then Action:=taConsult ;
FicheGene(nil, '', GS.Cells[SF_GEN, GS.Row], Action, 0) ;
FRechCompte.Info.Compte.Clear ;
//ObjFolio.Comptes := Comptes ;
//ObjFolio.Tiers   := Tiers ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 19/06/2006
Modifié le ... :   /  /    
Description .. : 
Suite ........ : - LG - 19/06/2006 - FB 18443 - on ne passe taCreat au
Suite ........ : fenetre de zoom
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.ZoomJal ;
var Action : TActionFiche ;
begin
if memJal.Code='' then Exit ;
if FAction = taCreat then
 Action := taModif
  else
   Action:=FAction ;
if not ExJaiLeDroitConcept(TConcept(ccJalModif), FALSE) then Action:=taConsult ;
FicheJournal(nil, '', memJal.Code, Action, 0) ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 02/06/2004
Modifié le ... : 05/07/2004
Description .. : - LG - 02/06/2004 - remise a zero de la liste des auxi charge
Suite ........ : - LG - 05/07/2004 - FB 13759 - rafraissichement de
Suite ........ : l'affichage apres l'appel de la fenetre des tiers. ( on a peu
Suite ........ : etre changer des truc )
Mots clefs ... :
*****************************************************************}
{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 02/06/2004
Modifié le ... : 19/06/2006
Description .. : - LG - 02/06/2004 - remise a zero de la liste des auxi charge
Suite ........ : - LG - 19/06/2006 - FB 18443 - on ne passe taCreat au 
Suite ........ : fenetre de zoom
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.ZoomTiers ;
var Action : TActionFiche ;
begin
if (memJal.Code='') or (GS.Cells[SF_AUX, GS.Row]='') then Exit ;
if FAction = taCreat then
 Action := taModif
  else
   Action:=FAction ;
if not ExJaiLeDroitConcept(TConcept(ccAuxModif), FALSE) then Action:=taConsult ;
FicheTiers(nil, '', GS.Cells[SF_AUX, GS.Row], Action, 1) ;
FRechCompte.Info.Aux.Clear ;
ObjFolio.Comptes := Comptes ;
ObjFolio.Tiers   := Tiers ;
InfosPied(GS.Row) ;
end ;

procedure TFSaisieFolio.ZoomDevise ;
begin
if GS.Cells[SF_DEVISE, GS.Row]<>''
   then FicheDevise(GS.Cells[SF_DEVISE, GS.Row], taConsult, FALSE)
   else FicheDevise(memPivot.Code, taConsult, FALSE) ;
end ;

procedure TFSaisieFolio.ZoomEtabl ;
begin
FicheEtablissement_AGL(taConsult) ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 29/04/2002
Modifié le ... : 16/04/2007
Description .. : on n'a plus acces au lettrage depuis la consultation des
Suite ........ : comptes
Suite ........ : - LG - 25/06/2003 - consultation des ecritures de type L
Suite ........ : - LG - 20/04/2004 - appel de la consultation des comptes
Suite ........ : qui etaient entre comentaire
Suite ........ : - LG - 16/04/2007 - FB 15709 - pour le MAJ+F5 sur la 
Suite ........ : derniere ligne, on supprime celle ci et on la re affcihe par la 
Suite ........ : suite. on a donc passer les info sur le compte et l'auxi a la 
Suite ........ : fct de zoom
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.ZoomNavCpte( Compte : string = '' ; Aux : string = '') ;
var NumCompte, NumAux, sExo : string ;
begin
if (memJal.Code='') or
  ((GS.Cells[SF_GEN, GS.Row]='') and (GS.Cells[SF_AUX, GS.Row]='')) then Exit ;
NumCompte:=Compte ; NumAux:=Aux ;
if NumCompte='' then
 begin
  NumCompte:=Trim(GS.Cells[SF_GEN, GS.Row]) ;
  if NumCompte<>'' then NumCompte:=BourreLaDonc(NumCompte, fbGene) ;
  NumAux:=Trim(GS.Cells[SF_AUX, GS.Row]) ;
  if NumAux<>'' then NumAux:=BourreLaDonc(NumAux, fbAux) ;
 end ;
// Exercice
// -1 = précédent
// 0 = courant
// 1 = suivant
if memOptions.TypeExo=teEncours then sExo:='0' else
  if memOptions.TypeExo=teSuivant then sExo:='1' else
    if memOptions.TypeExo=tePrecedent then sExo:='-1' ;
//LG* 29/04/2001
{$IFNDEF CMPGIS35}
OperationsSurComptes(NumCompte, sExo, '', NumAux,true) ;
{$ENDIF}
end ;

//=======================================================
//================= Scénario de saisie ==================
//=======================================================
procedure TFSaisieFolio.BLettrageClick(Sender: TObject);
begin LettrageEnSaisie ; end ;


{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 29/04/2002
Modifié le ... : 13/10/2005
Description .. : Suppression du verrou sur le folio avant d'entrer dans le
Suite ........ : lettrage
Suite ........ : - LG - 25/06/2003 - lettrage des ecritures de type L
Suite ........ : - LG - 10/07/2003 - FB 12484 - on ppouvait appeller le 
Suite ........ : lettrage en saisie avec des cractere alpha ds la case 
Suite ........ : debit/credit
Suite ........ : - LG - 02/10/2003 - on ouvre toutes les ecritures meme les 
Suite ........ : deja lettres
Suite ........ : - LG - 03/04/2004 - si le compte est pas lettrable, on sort
Suite ........ : - LG - 13/04/2004 - FB 13386 - on bloque la validation 
Suite ........ : pendant que le fenetre s'ouvre
Suite ........ : - LG - 04/03/2004 - ajout du blocage en lettrage en saisie
Suite ........ : - LG 06/09/2004 - FB 13633 - recalcul du solde en debut 
Suite ........ : de periode apres l'enregistrement
Suite ........ : - LG - 29/09/2004 - FB 14643 - recalcul du solde du journal 
Suite ........ : apres l'ouverture de la consultation
Suite ........ : - LG - 21/03/2005 - FB 15342 - suppression du deblocage 
Suite ........ : du lettrage. Fait au moment de l'enregistrement
Suite ........ : - LG - 24/03/2005 - FB 15342 - on force la sauvegarde du 
Suite ........ : folio avant d'appeller le lettrage
Suite ........ : - LG -13/09/2005 - FB 131515 - on ne peut plsu lettre un
Suite ........ : bordereau non equilibre, saus la der ligne
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.LettrageEnSaisie( Ou : integer = - 1 ; ALaLigne : boolean = false ) ;
var X: RMVT ;
    Ecr: TZF ;
    ACol, ARow: LongInt ;
    R  : RLETTR ;
    CodeL : String4 ;
    Cancel, Chg : Boolean ;
    i : integer ;
    lRdDebit,lRdCredit : double ;
    lStJournal : string ;
    lDtDateComptable : TDateTime ;
    lInNumeroPiece : integer ;
    lTError : TRecError ;
begin
 if FAction=taConsult then Exit ; if bGuideRun then Exit ; if ObjFolio=nil then Exit ;
 if bModeRO then exit ; Application.ProcessMessages ;
 if Ou=-1 then Ou:=GS.Row ;
 GrilleModif:=true ; Statut:=taModif ;
 {$IFDEF TT}
  AddEvenement('LettrageEnSaisie') ;
{$ENDIF}
 if (not BLettrage.Enabled) then exit ;
 if FboFenLett then exit ;
 FboFenLett := true ; 
 try
 lRdDebit:=Valeur(GS.Cells[SF_DEBIT,Ou]) ; lRdCredit:=Valeur(GS.Cells[SF_CREDIT,Ou]) ;
 if (lRdDebit=0) and (lRdCredit=0) then exit ;
 ACol:=GS.Col ; ARow:=Ou ; Cancel:=FALSE ;
 GSCellExit(GS, ACol, ARow, Cancel) ; if Cancel then exit ; Chg:=false ;
 if not ALaLigne then GSRowExit(GS, ARow, Cancel, Chg ); if Cancel then exit ;
 Ecr:=ObjFolio.GetRow(Ou-GS.FixedRows) ;
 if (Ecr<>nil) and ( Ecr.GetValue('LETTAG') = 'X' ) and ALaLigne then exit ;
 lTError.RC_Error := RC_PASERREUR ; lTError.RC_Message := '' ;
 if Ou <= nbLignes then
  lTError := CIsValidLigneSaisie(Ecr ,FRechCompte.Info) ;
 if lTError.RC_Error <> RC_PASERREUR then
  begin
   OnError(nil,lTError) ;
   FBoLettrageEnCours:=false ;
   exit ;
  end ;
 if (not memParams.bLibre) and (not IsSoldePiece(StrToInt(GS.Cells[SF_PIECE, Ou]))) and ( Ou<>nbLignes) then
  begin
   PrintMessageRC(RC_NONSOLDER) ;
   exit ;
  end ;

   HGBeginUpdate(GS) ;
   for i:=1 to GS.RowCount - 1 do GS.Cells[SF_QUALIF, i] := 'L' ;
   HGEndUpdate(GS) ;

 FBoLettrageEnCours:=true ;
 if not ValideFolio(FALSE, FALSE, TRUE) then
  begin
   FBoLettrageEnCours:=false ;
   {$IFDEF TT}
   AddEvenement('LettrageEnSaisie : sortie a l''enregistrement') ;
   {$ENDIF}
   exit ;
  end ;
 FBoLettrageEnCours:=true ;
// BackupFolio(true) ; // on force la sauvegarde du folio
 //LG* 29/04/2002 on enleve les verrrous avant d'appeller le lettrage
 //if not ObjFolio.UnLockFolio then exit ;
 {Remplissage du RMVT}
 FillChar(X, SizeOf(X), #0) ;
 X.Axe:='' ; X.Etabl:=Ecr.GetValue('E_ETABLISSEMENT') ;
 X.Jal:=Ecr.GetValue('E_JOURNAL') ;
 X.Exo:=QuelExoDT(Ecr.GetValue('E_DATECOMPTABLE')) ;
 X.CodeD:=Ecr.GetValue('E_DEVISE') ; X.Simul:=Ecr.GetValue('E_QUALIFPIECE') ;
 X.Nature:=Ecr.GetValue('E_NATUREPIECE') ; X.DateC:=Ecr.GetValue('E_DATECOMPTABLE') ;
 X.DateTaux:=Ecr.GetValue('E_TAUXDEV') ; X.Num:=Ecr.GetValue('E_NUMEROPIECE') ;
 if Ecr.GetValue('E_AUXILIAIRE')<>'' then X.NumLigne:=Ecr.GetValue('E_NUMLIGNE') ;
 X.TauxD:=Ecr.GetValue('E_TAUXDEV') ;
 X.Valide:=FALSE ; X.ANouveau:=FALSE ;
 {Remplissage du RLETT}
 FillChar(R,Sizeof(R),#0) ;
 CodeL:=Ecr.GetValue('E_LETTRAGE') ;
 R.TOBResult:=FTOBLettrage ;
 R.General:=Ecr.GetValue('E_GENERAL') ; R.Auxiliaire:=Ecr.GetValue('E_AUXILIAIRE') ;
 R.Appel:=tlSaisieCour ; R.GL:=Nil ;                                                 // LG 09/10/2003 supprime, on veut voi toute les ecritures
 R.CritMvt:=' AND E_GENERAL="'+R.General+'" AND E_AUXILIAIRE="'+R.Auxiliaire+'" ' ; //AND ( E_ETATLETTRAGE="TL" OR E_ETATLETTRAGE="PL" OR E_ETATLETTRAGE="AL" )' ; //AND E_ETATLETTRAGE<>"TL" AND E_ETATLETTRAGE<>"RI" ' ;
 R.CritDev:=Ecr.GetValue('E_DEVISE') ; R.DeviseMvt:=Ecr.GetValue('E_DEVISE') ;
 R.LettrageDevise:=(X.CodeD<>V_PGI.DevisePivot) ;
 R.Ident:=X ; R.Ident.NumLigne:=Ecr.GetValue('E_NUMLIGNE') ;
 R.Ident.NumEche:=Ecr.GetValue('E_NUMECHE') ; R.Distinguer:=True ;
 R.LettrageEnSaisie:=true ;
 if R.DeviseMvt=V_PGI.DevisePivot then
   R.CritDev:=V_PGI.DevisePivot ;
  lStJournal := Ecr.GetValue('E_JOURNAL') ;
  lDtDateComptable := Ecr.GetValue('E_DATECOMPTABLE') ;
  lInNumeroPiece := Ecr.GetValue('E_NUMEROPIECE') ;
  gs.OnEnter:= nil ;
  try
   if not CBlocageLettrage(lStJournal , lDtDateComptable , lInNumeroPiece ) then exit ;
   R.CritEtatLettrage:=' AND ( E_ETATLETTRAGE="TL" OR E_ETATLETTRAGE="PL" OR E_ETATLETTRAGE="AL" )' ;
   LettrageManuel(R,true,taModif) ;
   V_PGI.IOError := oeOk ;
   ReadFolio(FALSE) ;
  finally
//   CDeBlocageLettrageP(lStJournal , lDtDateComptable , lInNumeroPiece ) ;
   Ecr:=ObjFolio.GetRow(Ou-GS.FixedRows) ;
   if Ecr<>nil then
    Ecr.PutValue('LETTAG','X') ;
   for i:=1 to GS.RowCount - 1 do GS.Cells[SF_QUALIF, i] := 'N' ;
   memOptions.QualifP:='N' ;
   //ChargeJal(memJal.Code) ;
   //CalculSoldeJournal ;
  // BackupFolio(true) ; // on force la sauvegarde du folio
   SetGridOn ;
   if not ( CEstBloqueJournal(E_JOURNAL.Value,false) ) then
    CBloqueurJournal(true,E_JOURNAL.Value) ;
   if ((GS.Row - 1 ) = Ou) or (GS.Row > Ou)  then
    NextRow(GS.Row)
    else
     GS.Row := Ou ;
   SetGridOptions(GS.Row) ;
   EnableButtons ;
   CalculSoldeJournal ;
   GrilleModif:=true ;
   {$IFDEF TT}
    AddEvenement('Fini Lettrage en saisie $$$$$$$$$$$') ;
   {$ENDIF}
   gs.OnEnter:= GSEnter ;
  end;
 finally
  FboFenLett := false ;
 end ;
end ;

procedure TFSaisieFolio.BAideClick(Sender: TObject);
begin
{$IFDEF TT}
 DelEvenement ;
 exit ;
{$ENDIF}
 CallHelpTopic(Self) ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 30/01/2003
Modifié le ... : 10/06/2003
Description .. : LG - FB 11905 - on bloque la touche tabulation sur les
Suite ........ : lignes non editable
Suite ........ : - 09/06/2003 - FB 12216 - la validations d'un bordereau sur
Suite ........ : une ligne d'ecart de conversion ne fct pas
Suite ........ :
Mots clefs ... :
*****************************************************************}
procedure TFSaisieFolio.PEnteteEnter(Sender: TObject);
begin
if csDestroying in ComponentState then Exit ;
if not (GoEditing in GS.Options) then begin if GS.CanFocus then GS.SetFocus ; Exit ; end ;
if bClosing then Exit ;
if bModeRO then Exit ;
if Not bguideRun then Exit ;
if PrintMessageRC(RC_GUIDESTOP)=mrYes then GuideStop else GS.SetFocus ;
end;

procedure TFSaisieFolio.SetStatut ( Value : TActionFiche );
begin
// if not ( ( Value = taModif ) or ( Statut = taCreat ) ) then // le F5 place la grille en modification
  FStatut := Value;
 {$IFDEF TT}
 AfficheTitreAvecCommeInfo;
 {$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 26/06/2002
Modifié le ... : 25/09/2002
Description .. : - 25/06/2002 - gestion de l'affichage complet ou abrege de
Suite ........ : la nature de piece
Suite ........ : - test pour verifie que la case nature contient effectivement
Suite ........ : une valur
Mots clefs ... :
*****************************************************************}
function TFSaisieFolio.GetCellNature(ARow : integer) : string;
begin
result:='' ;
if FBoNatureComplete then
 begin
  if GS.Cells[SF_NAT, ARow]<>'' then
   result:=E_NATUREPIECE.Values[E_NATUREPIECE.Items.IndexOf(GS.Cells[SF_NAT, ARow])]
 end
  else
   result:=GS.Cells[SF_NAT, ARow] ;
end;


procedure TFSaisieFolio.OnError (sender : TObject; Error : TRecError ) ;
begin
 if trim(Error.RC_Message)<>'' then PGIInfo(Error.RC_Message, Caption )
  else
   if (Error.RC_Error<>RC_PASERREUR) then
    begin
     if Error.RC_Valeur <> '' then
      CPrintMessageReC(Caption,Error,FListeMessage)
       else
        PrintMessageRC(Error.RC_Error);
    end ;
end;


function  TFSaisieFolio.GetCompte : TZCompte ;
begin
if FRechCompte<>nil then
 result:=FRechCompte.Info.Compte
  else
   result:=nil;
end;

function TFSaisieFolio.GetTier : TZTier ;
begin
if (FRechCompte<>nil) and (FRechCompte.Info<>nil) then
 result:=FRechCompte.Info.Aux
  else
   result := nil ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 25/09/2002
Modifié le ... : 09/10/2002
Description .. : 25/09/2002 - solde le folio sur le compte passe en
Suite ........ : parametre
Suite ........ :
Suite ........ : -10/10/2002 - le compte ne doit pas collectif ou ventilable
Mots clefs ... :
*****************************************************************}
function TFSaisieFolio.FinSaisieBor( vStCompte : string ) : boolean;
var
 lIndex : integer ;
begin
result:=false ; if IsSoldeFolio then exit ;
lIndex:=FRechCompte.Info.Compte.GetCompte(vStCompte) ;
if (lIndex=-1) or FRechCompte.Info.Compte.IsCollectif(lIndex) or FRechCompte.Info.Compte.IsVentilable(0,lIndex) then exit ;
if (GS.Cells[SF_GEN, GS.Row]='') then GS.Cells[SF_GEN, GS.Row]:=vStCompte ;
SoldeClick(-1) ;
result:=true ;
end;

function TFSaisieFolio.SoldeAuto : boolean;
var
 lBoCancel : boolean ;
begin
result:= false ;
if trim(memJal.CContreP)<>'' then
 result:=FinSaisieBor(memJal.CContreP)
  else
   if (trim(memjal.CAutomat)<>'') and ( not bGuideRun ) then
    result:=FinSaisieBor(TrouveAuto(memjal.CAutomat,1)) ;
if not result then exit ; FCurrentRowEnter:=-1 ;
GSRowExit(Self, GS.Row,lBoCancel,true) ;
result:=not lBoCancel ;
end;



{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 04/10/2002
Modifié le ... :   /  /
Description .. : - 04/10/2002 - passe la grille en modif quand on appuie sur
Suite ........ : suppr
Mots clefs ... :
*****************************************************************}
procedure TFSaisieFolio.GSKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
// if ( Shift = [] ) and ( key = VK_DELETE ) then Statut:=taModif ;
if ( key = VK_DELETE ) then //or ( key = 16 ) then
 begin
  GrilleModif:=true ;
  Statut:=taModif ;
 end ;
end;

procedure TFSaisieFolio.AfficheTitreAvecCommeInfo( vInfo : string ='' );
begin
 if FListeMessage=nil then exit;
 if memParams.bLibre then Caption:=GetMessageRC(RC_SAISIEECR)+' '+GetMessageRC(RC_MODELIB)
 else Caption:=GetMessageRC(RC_SAISIEECR)+' '+GetMessageRC(RC_MODEBOR) ;
 {$IFDEF TT}
 case FStatut of
 taCreat   : Caption:='Creation';
 taConsult : Caption:='Consult';
 taModif   : Caption:='Modif';
 end; // case
 if GrilleModif then caption:=caption+' Modifiee ' else caption:=caption+' Pas modifiee ' ;
 {$ENDIF}
 if memJal.Accelerateur and not FBoArretAcc then
  Caption:=Caption+TraduireMemoire(' (accélérateur de saisie activé)') ;
 if trim(vInfo)<>'' then Caption:=Caption+' '+TraduireMemoire(vInfo) ;
 Caption:=traduireMemoire(Caption) ;
 UpdateCaption(Self) ;
end;


procedure TFSaisieFolio.BScanClick(Sender: TObject);
{$IFDEF SCANGED}
var
 Ecr : TZF ;
 lStGuid : string ;
{$ENDIF}
begin
{$IFDEF SCANGED}
 Ecr:=ObjFolio.GetRow(GS.Row-GS.FixedRows) ; if Ecr=nil then exit ;
lStGuid := ObjFolio.RechGuidId(Ecr) ; if lStGuid = '' then exit ;
if not ShowGedViewer(lStGuid,true) then
  ObjFolio.SupprimeLeDocGuid(Ecr) ;
{$ENDIF}
end;

{$IFDEF SCANGED}
procedure TFSaisieFolio.SetGuidId( vGuidId : string ) ;
var
 Ecr : TOB ;
begin
  GrilleModif:=true ; Statut:=taModif ;  if ObjFolio = nil then exit ;
  Ecr:=ObjFolio.GetRow(GS.Row-GS.FixedRows) ;  if Ecr=nil then exit ;
  if ( ObjFolio.RechGuidId(Ecr) <> '' ) and
  ( PGIAsk('Il existe une image ou un fichier associé , voulez-vous le remplacer ?')<>mrYes ) then exit ;
  ObjFolio.AjouteGuidId(Ecr,vGuidId ) ;
  if Pos('A',Ecr.GetValue('COMPTAG')) =  -1 then
   Ecr.PutValue('COMPTAG',Ecr.getValue('COMPTAG')+'A') ;
  EnableButtons ;
end;
{$ENDIF}

{$IFDEF SCANGED}
function TFSaisieFolio.GetInfoLigne : string;
var
 Ecr : TZF ;
begin
 result := '' ; if ObjFolio = nil then exit ;
 Ecr:=ObjFolio.GetRow(GS.Row-GS.FixedRows) ; if Ecr=nil then exit ;
 result:='Document scanné Période : '+E_DATECOMPTABLE.Text+' Journal : '+E_JOURNAL.Text+' Folio : '+E_NUMEROPIECE.Text ;
end;
{$ENDIF}

procedure TFSaisieFolio.SetGrilleModif(Value: boolean);
begin
 FBoGrilleModif := Value ; //EnableButtons ;
end;


function TFSaisieFolio.IsSoldeFolioPourAcc ( Ou : integer ) : Boolean ;
begin
 Result:=Arrondi(memJal.TotFolDebit-memJal.TotFolCredit-Valeur(GS.Cells[SF_DEBIT,Ou])+Valeur(GS.Cells[SF_CREDIT,Ou]), memPivot.Decimale)=0 ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 30/06/2004
Modifié le ... : 02/09/2004
Description .. : - 30/06/2004 - FB 13633 - gestion de la combinaison Ctrl
Suite ........ : +W. Sur un ligne sans montant, on recre la ligne
Suite ........ : - FB 13632 - 01/07/2004 - on reprend systematiquementle
Suite ........ : libelle du compte de la premiere ligne
Suite ........ : - FB 13631 - 10/08/2004 - controle que le folio est solde 
Suite ........ : avant de lancer l'aacelarateur de saisie
Suite ........ : - 25/08/2004 - LG FB 13633 - calcul du solde avant et 
Suite ........ : apres l'appel de l'accelerateur de saisie
Suite ........ : - 02/09/2004 - LG - FB 14364 - on ne recuperas pas le 
Suite ........ : regiem fiscal du tiers
Mots clefs ... : 
*****************************************************************}
function TFSaisieFolio.AccelerateurDeSaisie ( Ou : integer = - 1 ) : integer;
var
 EcrRef,Ecr              : TZF ;
 NumAux                  : string ;
 kt                      : integer ;
 lStGen                  : string ;
 lStCompteTVA            : string ;
 lRdTauxTva              : double ;
 ACol,ARow               : integer ;
 Cancel                  : boolean ;
 Chg                     : boolean ;
 C                       : double;
 lBoMvtDebit             : boolean ;
 lRdVal                  : double ;
 lRdTva                  : double ;
 Piece                   : string ;
 lStNatureAuxi           : string ;
 i                       : integer ;
 FOldStatut              : TActionFiche ;
 lStRegimeTVA            : string ;
begin
{$IFDEF TT}
 AddEvenement('AccelerateurDeSaisie');
{$ENDIF}
 result := - 1 ;
 if not memJal.Accelerateur then exit ;
 result := GS.Row ;  if Ou=-1 then Ou:=GS.Row ;
 if (Ou <> nbLignes)  then exit ; // on bloque si on est pas sur la derniere ligne
 if (GS.Cells[SF_DEVISE, Ou]<>'') then exit ; // l'accelerateur de saisei n'est pas actif pour une saisie en devise
 if ObjFolio = nil then exit ;
 if FBoValEnCours then exit ;
 CalculSoldeJournal ;
 if bGuideRun or not IsSoldeFolioPourAcc(Ou) then exit ;
 if FBoArretAcc then exit ;   FBoCutEnCours := true ;
 //if GS.Row <> nb    Lignes then exit ;
 GS.OnCellExit := nil ;
 GS.OnRowExit := nil ;

 FBoAccSaisie := true ;
 HGBeginUpdate(GS);
 try
 EcrRef:=ObjFolio.GetRow(Ou-GS.FixedRows) ; if EcrRef=nil then Exit ; EcrRef.PutValue('ACC','X') ; // sinon le rowexit boucle
 ACol:=GS.Col ; ARow:=Ou ; Cancel:=FALSE ;
 GSCellExit(GS, ACol, ARow, Cancel) ; if Cancel then exit ; Chg := false ;
 GSRowExit(GS, ARow, Cancel, Chg ); if Cancel then exit ;
 EcrRef:=ObjFolio.GetRow(Ou-GS.FixedRows) ;
 if EcrRef=nil then
  begin
   NextRow(GS.Row) ;
   exit ;
  end ; // if
 if ( EcrRef.GetValue('e_debit')=0 ) and ( EcrRef.GetValue('e_credit')=0 ) then exit ;

 for i:=1 to nbLignes do
  begin
   FOldStatut:=Statut ;
   Statut:=taModif ;
   if SetRowTZF(i,true)=nil then Continue ;
   Statut:=FOldStatut ;
  end ;

 EcrRef.PutValue('ACC','X') ;
 Piece:=GS.Cells[SF_PIECE, Ou] ;

 NumAux:=EcrRef.GetValue('E_AUXILIAIRE') ; kt:=Tiers.GetCompte(NumAux) ;
 if kt=-1 then exit ;
 lStGen := GetCompteAcc(Ou,NumAux) ;
 if lStGen='' then exit ;
 InitPivot ;

 lStNatureAuxi := '' ;
 lStRegimeTVA  := '' ;
 if Tiers.Load([GS.Cells[SF_AUX, ARow]]) <> -1 then
  begin
   if (tiers.getValue('YTC_ACCELERATEUR')='-') then
    begin
     EcrRef.PutValue('CPTCEGID',lStGen) ;
     EcrRef.PutValue('YTCTIERS', tiers.getValue('T_TIERS') ) ;
    end ;
   lStNatureAuxi := Tiers.getValue('T_NATUREAUXI') ;
   lStRegimeTVA  := Tiers.getValue('T_REGIMETVA') ;
  end ;
 FRechCompte.Info.Compte.RecupInfoTVA(lStGen,
                                      EcrRef.GetValue('E_DEBIT') ,
                                      EcrRef.GetValue('E_NATUREPIECE') ,
                                      memJal.Nature ,
                                      lStNatureAuxi ,
                                      lStCompteTVA,
                                      lRdTauxTva,
                                      lStRegimeTVA );

 lBoMvtDebit := EcrRef.GetValue('E_DEBIT') <> 0 ;
 if lBoMvtDebit then C := EcrRef.GetValue('E_DEBIT') else C := EcrRef.GetValue('E_CREDIT') ;

 if lStCompteTVA <> '' then
  begin
   lRdVal := Arrondi(C / ( 1 + ( lRdTauxTva / 100 ) ),FInfo.Devise.Dev.Decimale) ;
   lRdTva := Arrondi(C - lRdVal,FInfo.Devise.Dev.Decimale) ;
  end
   else
    begin
     lRdVal := Arrondi(C ,FInfo.Devise.Dev.Decimale) ;
     lRdTva := 0 ;
    end;

 GS.Row:=Ou+1 ;
 CreateRow(GS.Row) ;
 Ecr:=ObjFolio.GetRow(GS.Row-1) ;
 if Ecr=nil then exit ;
 CDupliquerTOBEcr(TOB(EcrRef),TOB(Ecr)) ;
 ecr.PutValue('e_general',lStGen) ;
 if lBoMvtDebit then
  ecr.SetMontants(0,lRdVal,FInfo.Devise.Dev,true)
   else
    ecr.SetMontants(lRdVal,0,FInfo.Devise.Dev,true) ;
 ecr.PutValue('e_libelle',EcrRef.GetValue('e_libelle')) ;
 ecr.PutValue('e_refinterne',EcrRef.GetValue('E_refinterne')) ;
 GetRowTZF(GS.Row-1) ;
 ACol:=GS.Col ; ARow:=GS.Row ; Cancel:=FALSE ;
 GSCellExit(GS, ACol, ARow, Cancel) ; if Cancel then exit ; Chg := false ;
 GSRowExit(GS, ARow, Cancel, Chg ); if Cancel then exit ;

 if lStCompteTVA <> '' then
  begin
   GS.Row:=GS.Row+1 ;
   CreateRow(GS.Row) ;
   Ecr:=ObjFolio.GetRow(GS.Row-GS.FixedRows) ;
   if Ecr=nil then exit ;
   CDupliquerTOBEcr(TOB(EcrRef),TOB(Ecr)) ;
   if lBoMvtDebit then
    ecr.SetMontants(0,lRdTva,FInfo.Devise.Dev,true)
     else
      ecr.SetMontants(lRdTva,0,FInfo.Devise.Dev,true) ;
   ecr.PutValue('e_general',lStCompteTVA);
   ecr.PutValue('e_libelle',EcrRef.GetValue('e_libelle')) ;
   ecr.PutValue('e_refinterne',EcrRef.GetValue('E_refinterne')) ;
   GetRowTZF(GS.Row-1) ;
   ACol:=GS.Col ; ARow:=GS.Row ; Cancel:=FALSE ;
   GSCellExit(GS, ACol, ARow, Cancel) ; if Cancel then exit ; Chg := false ;
   GSRowExit(GS, ARow, Cancel, Chg ); if Cancel then exit ;
  end ; // if

 NextRow(GS.Row+1) ;
 GS.Row := GS.Row + 1 ;

 finally
  result :=  GS.Row ;
  GS.OnCellExit := GSCellExit ;
  GS.OnRowExit := GSRowExit ;
  FBoAccSaisie:=false ;
  CalculSoldeJournal ;
  HGEndUpdate(GS);
 end ;

end;

procedure TFSaisieFolio.BAccClick(Sender: TObject);
begin
 AccelerateurDeSaisie ;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 01/07/2004
Modifié le ... : 07/09/2004
Description .. : - LG - 01/07/2004 - on pouvait utiliser un compte auxiliaire
Suite ........ : pour accelelreer la saisie
Suite ........ : - LG - 02/07/2004 - 13635 - la recherche du compte de
Suite ........ : contrepartie etait fausse
Suite ........ : - LG - 07/09/2004 - FB 13636 - la recupératyion du compte 
Suite ........ : de 
Suite ........ : contrepartie pôur la ligne en cours ne fct pas
Mots clefs ... : 
*****************************************************************}
function TFSaisieFolio.GetCompteAcc( Ou : integer ; E_AUXILIAIRE : string ) : string;
var
 i : integer ;
 lIndex : integer ;

 function _GetContre : string ;
 var
  lEcr : TZF ;
  i : integer ;
 begin
   for i:=Ou-1 downto 1 do
      begin
       if E_AUXILIAIRE = GS.Cells[SF_AUX, i]  then
        begin
         lEcr:=ObjFolio.GetRow(i-GS.FixedRows) ; if lEcr=nil then Exit ;
         result:=lEcr.GetValue('E_CONTREPARTIEGEN') ;
         Break ;
       end;
      end ; // for
 end ;


begin
 result := '' ;
 if (tiers.getValue('YTC_ACCELERATEUR')='X') then
  begin
   result:= VarToStr(Tiers.GetValue('YTC_SCHEMAGEN')) ;
   if result = '' then result := _GetContre ;
  end  // if
   else
    begin
     result := _GetContre ;
     if result = '' then
      result:= VarToStr(Tiers.GetValue('YTC_SCHEMAGEN')) ;
    end ;

 if result = '' then
  begin
   if result = '' then result := FStGenHT ;
   if result='' then
    for i:=Ou-1 downto 1 do
        begin
        if isRowHT(i) then
          begin
          result:=GS.Cells[SF_GEN, i] ;
          Break ;
          end;
        end ;
  end ;

 if result <> '' then
  begin
   lIndex:=FRechCompte.Info.Compte.GetCompte(result) ;
   if (lIndex=-1) or FRechCompte.Info.Compte.IsCollectif(lIndex)  then
    result := '' ;
  end ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 24/08/2004
Modifié le ... :   /  /
Description .. : - LG  - 24/08/2004 - FB 13637 - remise a jour du caption de 
Suite ........ :  la fenetre apres l'arret de l'accelerateur de saisie
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.BAccactifClick(Sender: TObject);
begin
 FBoArretAcc           := not FBoArretAcc ;
 PopS.Items[0].Checked := FBoArretAcc ;
 AfficheTitreAvecCommeInfo ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 25/08/2004
Modifié le ... :   /  /    
Description .. : - 25/08/2004 - LG - FB 13636 - possibilite d'activer
Suite ........ : l'accelerateur de saisie
Mots clefs ... :
*****************************************************************}
procedure TFSaisieFolio.BAccTiersClick(Sender: TObject);
begin
if not BZoomTiers.Enabled then Exit ;
if not ExJaiLeDroitConcept(TConcept(ccAuxModif), FALSE) then exit ;
if (memJal.Code='') or (GS.Cells[SF_AUX, GS.Row]='') then Exit ;
FicheTiersMZS(GS.Cells[SF_AUX, GS.Row],taModifEnSerie,7,'YTC_SCHEMAGEN="'+GetCompteAcc(GS.Row,GS.Cells[SF_AUX, GS.Row])+'";YTC_ACCELERATEUR="X"') ;
FRechCompte.Info.Aux.Clear ; InfosPied(GS.Row) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 11/07/2007
Modifié le ... :   /  /    
Description .. : - LG - Fb 17585 - on remts a faux l'indicateur de taux volatil
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.SetNewLigne( Value : boolean ) ;
begin
 FbNewLigne := Value ;
 if Value then
  begin
   FFirstRow := GS.Row ;
   FParams.volatile:=false;
  end ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 21/12/2005
Modifié le ... :   /  /    
Description .. : - LG  - 21/12/2005 - FB 17186 - on bloque les click de 
Suite ........ : souris qd on saisie des qt
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.PQuantiteExit(Sender: TObject);
var
 //Ecr : TZF ;
 bCancel : boolean ;
 ARow : integer ;
begin

 if csDestroying in ComponentState then Exit ;
 if ObjFolio = nil then exit ;

 {$IFDEF TT}
 AddEvenement('PQuantiteExit');
{$ENDIF}

// Ecr:=ObjFolio.GetRow(GS.Row-GS.FixedRows) ;
 if FTOBEnCours = nil then exit ;
 FTOBEnCours.PutValue('E_QTE1', Valeur(_QTE1.text) ) ;
 FTOBEnCours.PutValue('E_QTE2', Valeur(_QTE2.text) ) ;

 if FBoClickSouris then
  begin
   if _QTE1.canFocus then
    _QTE1.SetFocus ;
  // FBoClickSouris := false ;
   exit ;
  end ;

 PQuantite.visible := false ;
 GS.SetFocus ;
 if not FBoCutEnCours then
  GS.Row:=GS.Row+1 ;
 FBoCutEnCours := false ;
 if (GS.Row <= nbLignes)  then
  begin
   GS.Col:=SF_GEN ;
   GSRowEnter(nil, GS.Row, bCancel, FALSE);
   if bCancel then exit ;
   ARow:=GS.Row ;
   GSCellEnter(nil,SF_GEN,ARow,bCancel) ;
  end
   else
    NextRow(GS.Row) ;

 GS.OnExit := GSExit ;

end;

procedure TFSaisieFolio._QTE1Exit(Sender: TObject);
begin
 GereQuantite(nil) ; 
end;

procedure TFSaisieFolio.GereQuantite ( Ecr : TZF ) ;
var
 lStCaption1,lStCaption2 : string ;
 k : integer ; NumCompte : string ;
 lStQualifQte1,lStQualifQte2 : string ;
begin

 {$IFDEF TT}
 AddEvenement('GereQuantite');
{$ENDIF}

 if Ecr = nil then
  begin
   Ecr := FTOBEnCours ;
   if Ecr = nil then exit ;
  end
   else
    begin
     _QTE1.Text           := StrfMontant(Ecr.GetValue('E_QTE1'),15,V_PGI.OkDecQ,'',True) ;
     _QTE2.Text           := StrfMontant(Ecr.GetValue('E_QTE2'),15,V_PGI.OkDecQ,'',True) ;
    end ;
 _QTE1.Enabled := False;
 _QTE2.Enabled := False; 

 FTOBEnCours := Ecr ;

 if csDestroying in ComponentState then Exit ;
 if Ecr = nil then exit ;

 lStCaption1          := StrfMontant(  0,15,V_PGI.OkDecP,'',True) ;
 lStCaption2          := StrfMontant(  0,15,V_PGI.OkDecP,'',True) ;
 H_QUALIFQTE2.Visible := true ;
 H_QUALIFQTE1.Visible := true ;

 NumCompte:=Ecr.GetValue('E_GENERAL') ;
 if Valeur(_QTE1.Text) <> 0 then
  lStCaption1 := StrfMontant(  ( Ecr.GetValue('E_DEBIT')+Ecr.GetValue('E_CREDIT') ) /  Valeur(_QTE1.Text),15,V_PGI.OkDecQ,'',True) ;
 if Valeur(_QTE2.Text) <> 0 then
  lStCaption2 := StrfMontant(  ( Ecr.GetValue('E_DEBIT')+Ecr.GetValue('E_CREDIT') ) /  Valeur(_QTE2.Text),15,V_PGI.OkDecQ,'',True) ;
 if (Ecr.GetValue('E_QUALIFQTE1') <> '...') and (Ecr.GetValue('E_QUALIFQTE1') <> '') then
  lStQualifQte1 := RechDom('TTQUALUNITMESURE',Ecr.GetValue('E_QUALIFQTE1'),false) ;
 if (Ecr.GetValue('E_QUALIFQTE2') <> '...') and (Ecr.GetValue('E_QUALIFQTE2') <> '') then
  lStQualifQte2 := RechDom('TTQUALUNITMESURE',Ecr.GetValue('E_QUALIFQTE2'),false) ;

 k:=Comptes.GetCompte(NumCompte) ;
 if k>=0 then
  begin
   if ( lStQualifQte1 = '' ) or ( lStQualifQte1 = 'AUC' ) then
    begin
     lStQualifQte1 := RechDom('TTQUALUNITMESURE',Comptes.GetValue('G_QUALIFQTE1', k),false) ;
     Ecr.PutValue('E_QUALIFQTE1', Comptes.GetValue('G_QUALIFQTE1', k) ) ;
    end ;
   if Comptes.GetValue('G_QUALIFQTE1', k) <> '' then
    lStCaption1 := lStCaption1 + ' ' + FInfo.Devise.Dev.Symbole + ' / ' + lStQualifQte1 ;
   if ( lStQualifQte2 = '' ) or ( lStQualifQte2 = 'AUC' ) then
    begin
     lStQualifQte2 := RechDom('TTQUALUNITMESURE',Comptes.GetValue('G_QUALIFQTE2', k),false) ;
     Ecr.PutValue('E_QUALIFQTE2', Comptes.GetValue('G_QUALIFQTE2', k) ) ;
    end ;
   if Comptes.GetValue('G_QUALIFQTE2', k) <> '' then
    lStCaption2 := lStCaption2 + ' ' + FInfo.Devise.Dev.Symbole + ' / ' + lStQualifQte2 ;
   H_QUALIFQTE2.Visible := ( Valeur(_QTE2.Text) <> 0) or ( Comptes.GetValue('G_QUALIFQTE2', k) <> '' ) and ( Comptes.GetValue('G_QUALIFQTE2', k) <> 'AUC'  ) ;
  end ;

 H_QUALIFQTE1.Caption := TraduireMemoire('soit ')+ lStCaption1 ;
 H_QUALIFQTE2.Caption := TraduireMemoire('soit ')+ lStCaption2 ;
 _QTE2.Visible        := H_QUALIFQTE2.Visible ;
 H_QTE2.Visible       := H_QUALIFQTE2.Visible ;

end;

function TFSaisieFolio.EstAccelere ( Ou : integer ) : boolean ;
var
 NumAux : string ;
 Ecr : TZF ;
 kt : integer ;
begin
 result := false ;
 if Ou = - 1 then exit ;
 if Ou <> nbLignes then exit ;
 Ecr := ObjFolio.GetRow(Ou-GS.FixedRows) ; if Ecr=nil then exit ;
 if memJal.Accelerateur and (Ecr.GetValue('ACC')<>'X') then
  begin
   NumAux:=GS.Cells[SF_AUX, Ou] ;
   kt:=Tiers.GetCompte(NumAux) ;
   result :=  (kt<>-1) and (tiers.getValue('YTC_ACCELERATEUR')='X') ;
  end ; // if

end ;

procedure TFSaisieFolio._QTE1KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
 if Key = VK_RETURN then
  NextPrevControl(self) ;
end;

procedure TFSaisieFolio._QTE1Enter(Sender: TObject);
begin
 //FBoQt := true ;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 11/05/2006
Modifié le ... : 13/04/2007
Description .. : - FB 17434 - LG - gestion de la tva sur les compte anlytiq
Suite ........ : - FB 18204 - LG - l'appel de l'analytiq avec les touches de
Suite ........ : racourci fct pour tous les compte
Suite ........ : - FB 17434 - LG - 22/06/2006 - correction pour les
Suite ........ : scomptes
Suite ........ : 70+
Suite ........ : - FB 18771 - LG- Affectation du NEWLIGNE pour
Suite ........ : l'analytique. Controle
Suite ........ : de l'analytique en sortie
Suite ........ : - FB 18977 - LG - 13/10/2006 - on appelle la fct de
Suite ........ : validation de l'analytique et on ouvre la fenetre d'analytiq en
Suite ........ : cas d'erreur
Suite ........ : - 09/11/2006 - LG - FB 18902 - gestion de l'ouverture sur la
Suite ........ : saisie du montant
Suite ........ : - FB 19806 - LG - le changement d'auxi ne fct plus ( 601 -> 
Suite ........ : 471 l'anlytiq restait en base ) + optimisation de la gestion de 
Suite ........ : l'analytiq ds le bordereau
Mots clefs ... : 
*****************************************************************}
procedure TFSaisieFolio.GereAna(Row : LongInt ; bOpen : Boolean ) ;
var Ecr : TZF ; NumCompte : string ; k : Integer ;
 lInNumAxe : integer ;
 lBoAna : boolean ;
 lTError : TRecError ;
begin
if ObjFolio=nil then Exit ;
Ecr:=ObjFolio.GetRow(Row-GS.FixedRows) ; if Ecr=nil then Exit ;
NumCompte:=Ecr.GetValue('E_GENERAL') ; if NumCompte='' then Exit ;
k:=Comptes.GetCompte(NumCompte) ;
if k<0 then Exit ;
if not Comptes.IsVentilable(0, k) then
 begin
  ObjFolio.DelAna(Row-GS.FixedRows) ;
  Ecr.PutValue('E_ANA', '-') ;
  Exit ;
 end ;
lBoAna := ObjFolio.IsAnaFromRow(Row-GS.FixedRows) ;
Ecr.PutValue('E_ANA', 'X') ;
if bOpen and lBoAna then
 begin
  Ecr.PutValue('NEWVENTIL' , '-') ;
  GetAna(Row) ;
  exit ;
 end ;
if lBoAna then Exit ;
Ecr.PutValue('NEWVENTIL' , 'X') ;
if CAOuvrirVentil( TOB(Ecr), nil, Finfo, lInNumAxe ) then
 GetAna(Row)
  else
   begin
    ObjFolio.PostAnaCreate(Row-GS.FixedRows) ;
    ModifRefAna(Row-GS.FixedRows, 'CONTREPARTIEGEN',  Null) ; // Permet de forcer la mise à jour
    ModifRefAna(Row-GS.FixedRows, 'CONTREPARTIEAUX',  Null) ; // Permet de forcer la mise à jour
   end ;
   //PutAna(Row) ;

 lTError := CIsValidVentil (Ecr,FInfo) ;
 if lTError.RC_Error <> RC_PASERREUR then
  begin
   OnError(nil,lTError) ;
   if lTError.RC_Error = RC_YSECTIONINCONNUE then
    GereAna(Row,true) ;
  end ;


end ;


procedure TFSaisieFolio.EffacerListeMemoire ;
begin
 FRechCompte.Info.Compte.Clear ;
 FRechCompte.Info.Aux.Clear ;
end;


procedure TFSaisieFolio.BPDFClick(Sender: TObject);
{$IFDEF SCANGED}
//var
//Ecr : TOB ;
{$ENDIF}
begin
{$IFDEF SCANGED}
{Ecr:=ObjFolio.GetRow(GS.Row-GS.FixedRows) ;  if Ecr=nil then exit ;
if ( ObjFolio.RechGuidId(Ecr) <> '' ) and
( PGIAsk('Il existe un fichier associé , voulez-vous le supprimer ?')<>mrYes ) then exit ;  }
 AjouterFichierDansGed() ;
{$ENDIF}
end;

end.


