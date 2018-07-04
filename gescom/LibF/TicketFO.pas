{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 25/06/2001
Modifié le ... : 23/07/2001
Description .. : Saisie d'un ticket du Front Office
Mots clefs ... : FO
*****************************************************************}
unit TicketFO;
                      
interface

uses
  Forms, ExtCtrls, Menus, hmsgbox, Dialogs, HSysMenu, StdCtrls, Mask,
  Hctrls, ComCtrls, HRichEdt, HRichOLE, HTB97, Buttons, Grids, LCD_Lab,
  HFLabel, HPanel, Controls, Classes, UTOB, Messages, Graphics, Hent1,
  Windows, SysUtils, ShellAPI, M3FP, UIUtil, HStatus, Paramsoc, HDebug, LookUp,
  AglInit, Choix, ED_TOOLS,
{$IFDEF CHR}
  HRUtilsFO,HRFAct,FactChr,
{$ENDIF}                                 
  {$IFDEF EAGLCLIENT}
  MaineAgl,
  {$ELSE}
  Fe_Main, DBTables, UserChg, Doc_Parser, Courrier,
  {$ENDIF}
  EntGC, SaisUtil,
  FactUtil, FactCpta, FactArticle, FactTOB, FactTiers, FactPiece, FactAdresse,
  FactLotSerie, FactContreM, FactPieceContainer,
  {$IFDEF TAXEUS}
  Taxe,
  {$ENDIF}
  {$IFDEF STK}
   FactStock, stkMouvement, UTom,
  {$ENDIF STK}
  EcheanceFO, TickUtilFO, KB_Ecran, Af_Base, UFidelite,
  FactFormule, UtilFonctionCalcul, UtilGC, FActVariante,
  VentAna; // JTR - eQualité 12048

function CreerTicketFO(NaturePiece: string): boolean;
procedure AppelTicketFO(Parms: array of variant; nb: integer);
function SaisieTicketFO(CleDoc: R_CleDoc; Action: TActionFiche): boolean;
procedure AppelTransformeTicketFO(Parms: array of variant; nb: integer);
function TransformeTicketFO(CleDoc: R_CleDoc; NewNat: string): boolean;
procedure AppelDupliqueTicketFO(Parms: array of variant; nb: integer);
function DupliqueTicketFO(CleDoc: R_CleDoc; NewNat: string): boolean;
procedure PieceAjouteSousDetailFO(FTOBPiece: TOB);

// Gestion des panels et forms
type TOuSaisie = (saEntete, saLigne, saPied, saEcheance, saRemiseTicket);

  // Gestion des afficheurs
type TQuoiAffiche = (qaLibreTexte, qaLibreMontant, qaTotal, qaLigne, qaTimer);
  TOuAffiche = (ofInterne, ofClient, ofTimer);
  TSetOuAffiche = set of TOuAffiche;

type R_IdentCol = record
    ColName, ColTyp: string;
  end;
type
  TFTicketFO = class(TForm)
    PEntete: THPanel;
    GP_TIERS: THCritMaskEdit;
    HGP_TIERS: THLabel;
    HGP_DATEPIECE: THLabel;
    GP_DATEPIECE: THCritMaskEdit;
    HGP_REPRESENTANT: THLabel;
    HMTrad: THSystemMenu;
    FindLigne: TFindDialog;
    GP_DEVISE: THValComboBox;
    DockBottom: TDock97;
    GP_REPRESENTANT: THCritMaskEdit;
    GP_ETABLISSEMENT: THValComboBox;
    HTitres: THMsgBox;
    BZoomArticle: TBitBtn;
    BZoomTiers: TBitBtn;
    BZoomTarif: TBitBtn;
    POPZ: TPopupMenu;
    HPiece: THMsgBox;
    FTitrePiece: THLabel;
    GP_DEPOT: THValComboBox;
    BZoomCommercial: TBitBtn;
    BZoomPrecedente: TBitBtn;
    BZoomOrigine: TBitBtn;
    TDescriptif: TToolWindow97;
    Descriptif: THRichEditOLE;
    Outils97: TToolbar97;
    BMenuZoom: TToolbarButton97;
    BEche: TToolbarButton97;
    BInfos: TToolbarButton97;
    BChercher: TToolbarButton97;
    BSousTotal: TToolbarButton97;
    BActionsLignes: TToolbarButton97;
    Valide97: TToolbar97;
    BValider: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BAide: TToolbarButton97;
    ToolbarSep971: TToolbarSep97;
    Sep2: TToolbarSep97;
    BOffice: TToolbarButton97;
    BZoomSuivante: TBitBtn;
    BZoomEcriture: TBitBtn;
    PopD: TPopupMenu;
    PopV: TPopupMenu;
    VPiece: TMenuItem;
    VLigne: TMenuItem;
    BVentil: TToolbarButton97;
    PopL: TPopupMenu;
    TInsLigne: TMenuItem;
    TSupLigne: TMenuItem;
    TCommentEnt: TMenuItem;
    TCommentPied: TMenuItem;
    HErr: THMsgBox;
    POPD_S1: TMenuItem;
    POPD_S2: TMenuItem;
    POPD_S3: TMenuItem;
    GP_SAISIECONTRE: TCheckBox;
    GP_FACTUREHT: TCheckBox;
    GP_REGIMETAXE: THValComboBox;
    BDescriptif: TToolbarButton97;
    BZoomDevise: TBitBtn;
    PMain: TPanel;
    PnlCorps: THPanel;
    PPied: THPanel;
    LblPrixU: THLabel;
    LblA_LIBELLE: THLabel;
    T_PrixU: TFlashingLabel;
    A_QTESTOCK: TFlashingLabel;
    PI_NETAPAYERDEV: TLCDLabel;
    TPI_NETAPAYERDEV: THLabel;
    LIBELLETIERS: THLabel;
    GS: THGrid;
    fAffTXT: TLCDLabel;
    TimerGen: TTimer;
    TimerLCD: TTimer;
    GP_TOTALHTDEV: THNumEdit;
    GP_TOTALREMISEDEV: THNumEdit;
    HTOTALTAXESDEV: THNumEdit;
    GP_TOTALESCDEV: THNumEdit;
    GP_TOTALTTCDEV: THNumEdit;
    GP_ESCOMPTE: THNumEdit;
    GP_REMISEPIED: THNumEdit;
    GP_MODEREGLE: THValComboBox;
    GP_NUMEROPIECE: THLabel;
    BAcompte: TToolbarButton97;
    BDelete: TToolbarButton97;
    LIBELLEARTICLE: THLabel;
    GP_DATELIVRAISON: THCritMaskEdit;
    BPorcs: TToolbarButton97;
    BImprimer: TToolbarButton97;
    ToolbarButton972: TToolbarButton97;
    POPY: TPopupMenu;
    CpltEntete: TMenuItem;
    CpltLigne: TMenuItem;
    N1: TMenuItem;
    AdrLiv: TMenuItem;
    AdrFac: TMenuItem;
    N2: TMenuItem;
    MBtarif: TMenuItem;
    MBDatesLivr: TMenuItem;
    MBSoldeReliquat: TMenuItem;
    N3: TMenuItem;
    MBDetailNomen: TMenuItem;
    N4: TMenuItem;
    TLigneRetour: TMenuItem;
    PopT: TPopupMenu;
    MCCreerClient: TMenuItem;
    MCChoixClient: TMenuItem;
    BZoomStock: TBitBtn;
    BClient: TToolbarButton97;
    MCModifClient: TMenuItem;
    PImage: THPanel;
    ImageArticle: TImage;
    GP_TOTALQTEFACT: THNumEdit;
    LGP_TOTALQTEFACT: TLabel;
    TLGP_TOTALQTEFACT: TLabel;
    Librepiece: TMenuItem;
    BTicketAttente: TToolbarButton97;
    BActionsDiv: TToolbarButton97;
    PopM: TPopupMenu;
    MDChoixModele: TMenuItem;
    MDSituFlash: TMenuItem;
    N5: TMenuItem;
    MDReimpTicket: TMenuItem;
    MDOuvreTiroir: TMenuItem;
    MDStock: TMenuItem;
    MBChangeDepot: TMenuItem;
    N6: TMenuItem;
    MDDetaxe: TMenuItem;
    LDETAXE: TLabel;
    SOLDETIERS: TFlashingLabel;
    TSOLDETIERS: THLabel;
    BJustifTarif: TBitBtn;
    N7: TMenuItem;
    MCFidelite: TMenuItem;
    MCSoldeClient: TMenuItem;
    MDReimpBons: TMenuItem;
    PGrille: TPanel;
    PClavier2: TPanel;
    TRechArtImage: TMenuItem;
    MCContactClient: TMenuItem;
    PopX: TPopupMenu;
    MXExcepTaxe: TMenuItem;
    BTaxation: TToolbarButton97;
    MXTaxePiece: TMenuItem;
    MXTaxeLigne: TMenuItem;
    MDALivrer: TMenuItem;
    LALIVRER: TLabel;
    N8: TMenuItem;
    MCHistoClient: TMenuItem;
    MCArticleClient: TMenuItem;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure BChercherClick(Sender: TObject);
    procedure GSMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure GSElipsisClick(Sender: TObject);
    procedure GSEnter(Sender: TObject);
    procedure GSDblClick(Sender: TObject);
    procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure FindLigneFind(Sender: TObject);
    procedure BZoomArticleClick(Sender: TObject);
    procedure BMenuZoomMouseEnter(Sender: TObject);
    procedure BZoomTarifClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BZoomTiersClick(Sender: TObject);
    procedure PEnteteExit(Sender: TObject);
    procedure GP_DEVISEChange(Sender: TObject);
    procedure GP_DATEPIECEExit(Sender: TObject);
    procedure GP_TIERSElipsisClick(Sender: TObject);
    procedure GP_TIERSDblClick(Sender: TObject);
    procedure GP_TIERSExit(Sender: TObject);
    procedure GP_DEPOTChange(Sender: TObject);
    procedure GP_ESCOMPTEExit(Sender: TObject);
    procedure GP_REMISEPIEDExit(Sender: TObject);
    procedure CpltEnteteClick(Sender: TObject);
    procedure CpltLigneClick(Sender: TObject);
    procedure AfficheZonePied(Sender: TObject);
    procedure AdrLivClick(Sender: TObject);
    procedure AdrFacClick(Sender: TObject);
    procedure GP_REPRESENTANTElipsisClick(Sender: TObject);
    procedure BZoomCommercialClick(Sender: TObject);
    procedure GP_REPRESENTANTExit(Sender: TObject);
    procedure MBtarifClick(Sender: TObject);
    procedure BZoomPrecedenteClick(Sender: TObject);
    procedure BZoomOrigineClick(Sender: TObject);
    procedure MBSoldeReliquatClick(Sender: TObject);
    procedure TDescriptifClose(Sender: TObject);
    procedure BZoomEcritureClick(Sender: TObject);
    procedure BSousTotalClick(Sender: TObject);
    procedure BZoomSuivanteClick(Sender: TObject);
    procedure BEcheClick(Sender: TObject);
    procedure VPieceClick(Sender: TObject);
    procedure VLigneClick(Sender: TObject);
    procedure TInsLigneClick(Sender: TObject);
    procedure TSupLigneClick(Sender: TObject);
    procedure TCommentEntClick(Sender: TObject);
    procedure TCommentPiedClick(Sender: TObject);
    procedure PopEuro(Sender: TObject);
    procedure GP_ETABLISSEMENTChange(Sender: TObject);
    procedure BDescriptifClick(Sender: TObject);
    procedure BOfficeClick(Sender: TObject);
    procedure GP_TIERSEnter(Sender: TObject);
    procedure MBDetailNomenClick(Sender: TObject);
    procedure BZoomDeviseClick(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure BAcompteClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure GP_REGIMETAXEChange(Sender: TObject);
    procedure MBDatesLivrClick(Sender: TObject);
    procedure BPorcsClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure LibrepieceClick(Sender: TObject);
    procedure EnvoieToucheGrid;
    procedure FormResize(Sender: TObject);
    procedure PEnteteResize(Sender: TObject);
    procedure TimerGenTimer(Sender: TObject);
    procedure BAbandonClick(Sender: TObject);
    procedure TimerLCDTimer(Sender: TObject);
    procedure BZoomStockClick(Sender: TObject);
    procedure TLigneRetourClick(Sender: TObject);
    procedure MCCreerclientClick(Sender: TObject);
    procedure MCChoixClientClick(Sender: TObject);
    procedure MCModifClientClick(Sender: TObject);
    procedure GP_REPRESENTANTDblClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure BTicketAttenteClick(Sender: TObject);
    procedure GP_REPRESENTANTEnter(Sender: TObject);
    procedure MDChoixModeleClick(Sender: TObject);
    procedure MDSituFlashClick(Sender: TObject);
    procedure MDReimpTicketClick(Sender: TObject);
    procedure MDOuvreTiroirClick(Sender: TObject);
    procedure MDStockClick(Sender: TObject);
    procedure MBChangeDepotClick(Sender: TObject);
    procedure MDDetaxeClick(Sender: TObject);
    procedure BJustifTarifClick(Sender: TObject);
    procedure MCFideliteClick(Sender: TObject);
    procedure MCSoldeClientClick(Sender: TObject);
    procedure MDReimpBonsClick(Sender: TObject);
    procedure TRechArtImageClick(Sender: TObject);
    procedure MCContactClientClick(Sender: TObject);
    procedure MXExcepTaxeClick(Sender: TObject);
    procedure MXTaxePieceClick(Sender: TObject);
    procedure MXTaxeLigneClick(Sender: TObject);
    procedure MDALivrerClick(Sender: TObject);
    procedure MCArticleClientClick(Sender: TObject);
    procedure MCHistoClientClick(Sender: TObject);
    // eQualité 11973
    function InsertNewArtInGS(RefArt, Designation: string; Qte: Double) : boolean;
    procedure MultiSelectionArticle(TobMultiSel : TOB);
    // Fin eQualité 11973
  private
{$IFDEF GESCOM} // JTR 15/01/2004
    function CumulArtOk(Acol, Arow : integer; Cancel : boolean) : boolean;
    procedure RemiseDuTarif(Arow : integer); // JTR - Démarque obligatoire
{$ENDIF GESCOM}
  protected
    GereStatPiece: boolean;
    GX, GY: Integer;
    WMinX, WMinY: Integer;
    FindFirst, ForcerFerme, ReliquatTransfo, GereReliquat, GeneCharge, OkCpta, NeedVisa, GereLot, GereAcompte,
      GereSerie, GereAdresse,
      OkCptaStock,
      CommentLigne, ForceRupt, EstAvoir, QuestionRisque, SansRisque, DejaRentre, ValideEnCours,
      CataFourn, TotEnListe, OuvreAutoPort, PasToucheDepot: boolean;
    ChangeComplEntete: boolean;
    StCellCur, LesColonnes, CompAnalP, CompAnalL, GereEche, CommentEnt, CommentPied, CalcRupt,
      CliCur, VenteAchat, DimSaisie, TypeCom, OldRepresentant: string;
    NewLigneDim: Boolean; //AC, nouvelle ligne pour gestion d'un art GEN existant
    IdentCols: array[0..19] of R_IdentCol;
    {$IFNDEF FOS5}
    PPInfosLigne: TStrings;
    {$ENDIF}
    MontantVisa, RisqueTiers: Double;
    LeAcompte: T_GCAcompte;
    ANCIEN_TOBDimDetailCount: integer; // Sauvegarde du nb initial de dimensions
    MontantHT: Double; //AC
{$IFDEF CHR}
    TransfertFac   : Boolean;
{$ENDIF}
    VenteALivrer: boolean; // vente à livrer ou sur place
    InMaximise: Boolean; // maximisation en cours
    InTraiteBouton: Boolean; // traitement d'un bouton en cours
    DejaSaisie: Boolean; // une saisie a eu lieu dans la grille
    SurLigSuivante: Boolean; // pour se positionner automatiquement sur la prochaine ligne vierge
    RetourGS: boolean; // retour sur la grille en sortie du champ courant
    PnlBoutons: TClavierEcran; // pavé tactile
    PnlBtn2: TClavierEcran; // pavé tactile secondaire
    AFFEXTERNE: TAFF; // Afficheur externe
    QuoiSurLCD: TQuoiAffiche; // Type du message affiché actuellement sur le LCD
    TMOSimuleClav: Integer; // Temps d'attente pour FOSimuleClavier
    // Objets mémoire
    TOBOpCais, TOBOpCais_O: TOB;
    TOBArtFiTiers : TOB; // JTR Lien Opération Caisse avec Tiers - Tob contenant les art financiers + tiers associés
    TobTicketAttente : TOB;
    MultiSel_SilentMode : boolean; //eQualité 11973
    // Dimensionnement
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    // Actions liées au Grid
    procedure EtudieColsListe;
    procedure ZoneSuivanteOuOk(var ACol, ARow: Longint; var Cancel: boolean);
    function ZoneAccessible(ACol, ARow: Longint): boolean;
    procedure FormateZoneSaisie(ACol, ARow: Longint);
    function FormateZoneDivers(St: string; ACol: Longint): string;
    function GereElipsis(LaCol: integer): boolean;
    procedure AllerSurLigSuivante(ACol, ARow, OldCol, OldRow: Integer);
    // Initialisations
    procedure InitPasModif;
    procedure InitToutModif;
    procedure BloquePiece(ActionSaisie: TActionFiche; Bloc: boolean);
    procedure BlocageSaisie(Bloc: boolean);
    procedure InitEnteteDefaut(Totale: boolean);
    procedure InitPieceCreation; dynamic;
    procedure LoadLesTOB;
    procedure GereVivante;
    procedure EtudieReliquat;
    procedure ChargelaPiece;
    procedure InitEuro;
    procedure InitRIB;
    function FromAvoir: boolean;
    procedure ToutAllouer;
    procedure ToutLiberer;
    procedure ReInitPiece;
    function  ReInitUtilisateur: boolean;
    {$IFNDEF FOS5}
    procedure AppliqueTransfoDuplic;
    {$ENDIF}
    procedure ChargeFromNature;
    // Calculs, Euro
    procedure CalculeLaSaisie(ACol, ARow: integer; AffTout: boolean);
    procedure AfficheEuro;
    procedure RempliPopD;
    procedure ChoixEuro(Sender: TObject);
//{$IFNDEF FOS5}
    procedure CalculeLeDocument;
{$IFNDEF FOS5}
    procedure AfficheTaxes;
{$ENDIF}
    procedure ZoomOuChoixLib(ACol, ARow: integer);
    // Affichages
    procedure ShowDetail(ARow: integer);
    procedure GetTotalaAfficher(TOBL: TOB; var Montant: double; var NbDec: integer; var Sigle: string);
    procedure AffichePUBrut(TOBLig: TOB);
    procedure AfficheTotalSurLCD;
    procedure ZoomOuChoixArt(ACol, ARow: integer);
    procedure GereEnabled(ARow: integer);
    procedure EnabledPied;
    procedure EnabledGrid;
    procedure GereTiersEnabled;
    procedure GereCommercialEnabled;
    procedure GetCellCanvas(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure PostDrawCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    function GotoEntete: Boolean;
    procedure GotoGrid(PageDown, FromF10: boolean);
    procedure GotoPied;
    procedure FlagRepres;
    procedure FlagTiers;
    procedure FlagPU;
    procedure FlagDepot;
    // TOBS lignes
    function PieceModifiee: boolean;
    procedure OnGetVar(Sender: TObject; VarName: string; VarIndx: Integer; var Value: variant);
    procedure TOBLigneExiste(ARow: integer);
    // LIGNES
//    function InitLaLigne(ARow: integer; NewQte, NewPrix: double): T_ActionTarifArt; // eQualité 11973
    function InitLaLigne(ARow: integer; NewQte : double; NewPrix: double = 0): T_ActionTarifArt;
    procedure ComplementLigne(ARow: integer);
    procedure AfficheLaLigne(ARow: integer; AffLCD: Boolean = False);
    procedure InsertComment(Ent: Boolean);
    procedure RecalculeSousTotaux(TOBPiece: TOB);
    procedure ReAfficheSousTotal(ARow: integer);
    procedure AffichePorcs;
    procedure ClickPorcs;
    // ENTETE
    function ComplementEntete: boolean;
    // DEMARQUES
    procedure TraiteDemarque(var ACol, ARow: integer; var Cancel: boolean);
    function VerifDemarque(CodeDem: string): Boolean;
    // REPRESENTANTS
    procedure BalayeLignesRep;
    function IdentifierRepres(var ACol, ARow: integer): boolean;
    procedure ZoomRepres(Repres: string);
    function RepresExiste(Repres: string): boolean;
    procedure ZoomOuChoixRep(ACol, ARow: integer);
    procedure TraiteRepres(var ACol, ARow: integer; var Cancel: boolean);
    // COMPTABILITE
    procedure ZoomEcriture(DeStock: boolean);
    procedure GereAnal(ARow: integer);
    procedure ClickVentil(LePied: Boolean);
    // DEPOTS
    procedure TraiteDepot(var ACol, ARow: integer; var Cancel: boolean);
    procedure ZoomOuChoixDep(ACol, ARow: integer);
    function IdentifierDepot(var ACol, ARow: integer): boolean;
    function DepotExiste(Depot: string): boolean;
    procedure ChangeDepot(ARow: Integer);
    // DIMENSIONS
    procedure RemplirTOBDim(CodeArticle: string; Ligne: Integer);
    procedure AppelleDim(ARow: integer);
    procedure TraiteLesDim(var ARow: integer; NewArt: boolean);
    procedure AffichageDim;
    procedure MAJLigneDim(ARow: Integer);
    procedure ReconstruireTobDim(ARow: Integer);
    // ARTICLES
    procedure UpdateArtLigne(ARow: longint; FromMacro, Calc: boolean; NewQte, NewPrix: double);
    procedure TraiteLesLies(ARow: integer);
    procedure TraiteLesCons(ARow: integer);
    procedure TraiteLesNomens(ARow: integer);
    procedure TraiteLesMacros(ARow: integer);
    procedure TraiteLaCompta(ARow: integer);
    function IdentifierArticle(var ACol, ARow: integer; var Cancel: boolean; Click, FromMacro: Boolean): boolean;
    procedure TraiteArticle(var ACol, ARow: integer; var Cancel: boolean; FromMacro, Calc: boolean; NewQte: double; FromKB: boolean = FALSE);
    procedure CodesArtToCodesLigne(TOBArt: TOB; ARow: integer);
    function PlusEnStock(TOBArt: TOB): boolean;
    procedure ChargeTOBDispo(ARow: integer);
    procedure LoadLesArticles;
//    procedure ConstruireTobArt(TOBPiece: Tob);
    procedure LoadTOBCond(RefUnique: string);
    procedure GereArtsLies(ARow: integer);
{$IFDEF STK}
    procedure GereStock (ARow: integer; bAppel, bForceAffichage, bAuto : boolean; bParle : boolean = true);
    function  GereStockNomen (ARow: integer; bAppel, bForceAffichage, bAuto : boolean; bParle : boolean = true) : boolean;
    function  DuplicLeDetailSurAvoir : boolean;
{$ELSE STK}
    procedure GereLesLots(ARow: integer);
    procedure DetruitLot(ARow: integer);
    procedure GereLesSeries(ARow: integer);
    procedure DetruitSerie(ARow: integer);
{$ENDIF STK}
    procedure TraiteLeCata(ARow: integer);
    // QUANTITES / TARIFS / DIVERS
{$IFDEF CHR}
    Procedure TraiteFolio (ACol,ARow : integer; Var Cancel : boolean) ;
    Procedure GereChampsChr( Arow : integer;var cancel:boolean) ;
{$ENDIF}
//    function TraiteRupture(ARow: integer): boolean; // eQualité 11973
    function TraiteRupture(ARow: integer; bParle : boolean = True): boolean;
    function TraiteQte(ACol, ARow: integer; bParle : boolean = true): boolean; dynamic;
    procedure TraiteLibelle(ARow: integer);
    procedure TraitePrix(ARow: integer; var Cancel: Boolean);
    procedure TraitePrixNet(ACol, ARow: integer; var Cancel: Boolean);
    procedure TraiteRemise(ARow: integer; var Cancel: Boolean);
    procedure TraiteMontantRemise(ACol, ARow: integer; var Cancel: Boolean);
    procedure RetourneArticle(ARow: Integer);
{$IFDEF MODE}
    procedure RechercheArticleImage;
{$ENDIF}
    procedure TraiteMontantNetLigne(ACol, ARow: integer; var Cancel: Boolean);
    function AppliquePUBrut(TOBL: TOB; NewPU: Double; ACol, ARow: Integer): Boolean;
    procedure TraiteLesDivers(ACol, ARow: integer);
    procedure DeflagTarif(ARow: integer);
    function TesteMargeMini(ARow: integer): boolean;
    // TAXES
    procedure ChangeTVA;
    procedure ChangeRegime;
    procedure TrouveRegimePiece;
{$IFDEF TAXEUS}
    procedure ExceptionTaxe;
    procedure TaxationPiece;
    procedure TaxationLigne;
    procedure IncidenceTaxe;
    function IncidenceTaxeLigne(ARow: integer; ForceModele: boolean): boolean;
{$ENDIF}
    // TIERS, DEVISES
    procedure IncidenceTauxDate;
    procedure IncidenceTiers;
    function IncidenceTarif: boolean;
    procedure IncidenceRefs;
    function IdentifieTiers(Click: boolean): boolean;
    function TestNatureTiers(Nature, Tiers : string) : boolean;
    procedure ForceTiers(CodeTiers: string);
    procedure CreationTiers;
    procedure ModificationTiers;
    procedure VoirSoldeTiers;
    procedure AfficheSoldeTiers;
    procedure ChoixTiers;
    procedure MajTobAvecTiers;
    procedure AssigneTiers;
    procedure LectureCBFromKB;
    procedure CalculeRisqueTiers;
    function TesteRisqueTiers(Valid: Boolean): boolean;
    function ChargeTiers: boolean;
    function GetLeTauxDevise(CodeD: string; var DateTaux: TDateTime; DateP: TDateTime): Double;
    procedure TrouveDevisePiece;
    // AFFAIRES
    procedure MajDatesLivr;
    // ADRESSES
    procedure LoadTOBAdresses;
    procedure TicketALivrer;
    // Gestion du descriptif détaillé
    procedure GereDescriptif(ARow: Integer; Enter: Boolean);
    // Actions, outils
    procedure ClickDel(ARow: integer; AvecC, FromUser: boolean);
    procedure ClickInsert(ARow: integer);
    procedure ClickPrecedente;
    procedure ClickSuivante;
    procedure ClickOrigine;
    procedure SoldeReliquat;
    function ClickSousTotal: Boolean;
    procedure ClickDevise;
    function ExitDatePiece: Boolean;
    // Contrôles d'intégrités et saisie des échéances
    function FOPieceCorrecte(TOBPiece, TOBArticles: TOB): Integer;
    function ReglementDispo(ARow: integer; var Montant, Quantite: double; var Ok: boolean): boolean;
    function ArtFiClientOblig(TOBLig: TOB; ARow: Integer): Integer;
    function VerifClientOblig(IdMsg: Integer): Boolean;
    procedure RemiseSousTotalCorrecte(ARow: integer; TauxRemise, RemiseSousTotal, MontantNet: double; TypeRemise: string; var Cancel: boolean);
    procedure RemiseCorrecte(TauxRemise: double; TypeRemise: string; TypeUniquement: boolean; var Cancel: boolean);
    procedure AppelEcheancesFO;
    // Validations
    function SortDeLaLigne: boolean;
    procedure InverseAvoir;
    procedure ValideLaPiece;
    procedure ValideImpression;
    procedure ValideNumeroPiece;
    function ValideTablesLibres(FromBouton: boolean): boolean;
    procedure DelOrUpdateAncien;
    procedure ValideTiers;
    procedure ValideAdresses;
    procedure ValideLaNumerotation;
    procedure ValideLesNumerosLignes;
    // MODIF_DBR_DEBUT
    {$IFNDEF STK}
    procedure ValideLesLots;
    procedure ValideLesSeries;
    {$ENDIF STK}
    // MODIF_DBR_FIN
    procedure ValideLaCompta(OldEcr, OldStk: RMVT);
    procedure GereLesReliquats;
    function CreerReliquat: boolean;
    procedure DetruitLaPiece;
    procedure ClickValideTicket;
    // Afficheur
    procedure InvertirAfficheur;
    procedure Resizerafficheur;
    function TraiteTexteAff(ST: string): string;
    procedure LCDInVisible;
    // Ecran tactiles
    procedure OnChangeCalculette(AValue: Variant);
    procedure OnChangeCalculette2(AValue: Variant);
    function AllerAuControle(Vers: TOuSaisie; NoLig, NoCol: Integer; RechLigne, VerifAcces: Boolean): Boolean;
    procedure ToucheArticle(Concept, Code, Extra: string; Qte, Prix: Double);
    procedure ToucheCommentaire(Concept, Code, Extra: string; Qte, Prix: Double);
    procedure ToucheRemise(Concept, Code, Extra: string; Qte, Prix: Double);
    procedure ToucheVendeur(Concept, Code, Extra: string; Qte, Prix: Double);
    procedure ToucheModePaie(Concept, Code, Extra: string; Qte, Prix: Double);
    procedure ToucheModele(Concept, Code, Extra: string; Qte, Prix: Double);
    {$IFDEF CHR}
    Procedure ToucheRes ( Concept, Code, Extra : String ) ;
    {$ENDIF}
    procedure SimuleChoixLigneMenu(LigneMenu: TMenuItem; BtnEcheance, RechLigne: Boolean);
    procedure SimuleBoutonZoom(BZoom: TBitBtn; BtnEcheance, RechLigne: Boolean);
    procedure AllerSur(NoCol: Integer);
    procedure DispatchFonctionClavier(Concept, Code, Extra: string; Qte, Prix: Double);
    procedure SaisieClavierEcran(Concept, Code, Extra: string; Qte, Prix: Double);
    procedure SaisieCalculette(Val: Double);
    procedure BoutonCalculetteClick(Val: string);
    procedure ChargeChoixDesPaves(var PropPv2: RPropPave2);
    procedure TraiteFormuleQte (ARow : integer) ; // JTR - eQualité 11203 (formules de qté)
    procedure RappelFormuleQte(ARow : integer) ; // JTR - eQualité 11203 (formules de qté)

  public
    SaisieTypeAffaire: boolean;
    ////MontantHT : Double; //AC
    AnnulPiece: Boolean; // annulation du ticket courant
    AssignePiece: Boolean; // assignation du ticket courant à un tiers
    FormEcheance: TFEcheanceFO; // Forme de saisie des échéances
    DepuisEcheance: Boolean; // retour de la saisie des échéances
    DemarqueCliObl: Boolean; // il existe de code démarque qui rend le client obligatoire
    DemandeDetaxe: boolean; // demande de détaxe formulée
    ColCarac: array[0..41] of RColCarac; // Caractéristiques des colonnes
    //    de  0 à 19 => colonnes de la grille des lignes (GS)
    //    de 20 à 39 => colonnes de la grille des échéances (GSReg)
    //    de 40 à 41 => champs d'en-tête (40=Client, 41=Vendeur)
    TOBPiece, TOBEches, TOBBases, TOBTiers, TOBArticles, TOBPorcs: TOB;
    TOBPIECERG, TOBBASESRG: TOB;
    PieceContainer: TPieceContainer;
    TOBTarif, TOBAdresses, TOBConds, TOBComms,
    TOBCpta, TOBAnaP, TOBAnaS: TOB;
    TobLigneTarif, TobLigneTarif_O: Tob; { Pour mémoriser le détail du système tarifaire }
    TOBPiece_O, TOBBases_O, TOBEches_O, TOBCXV, TOBAffaire, TOBNomenclature, TOBN_O: TOB;
    TOBDim, TOBSerie_O, TOBDesLots, TOBSerie, TOBSerRel, TOBLOT_O, TOBAcomptes, TOBAcomptes_O, TOBPorcs_O: TOB;
    TOBCatalogu, TOBDispoContreM: TOB;
    {$IFDEF CHR}
    TobHRDossier,Tobregrpe : TOB;
    {$ENDIF}
    // Fidélité JD
    FideliteCli: Fidelite;
    TOBAFormule, TOBLigFormule, TobAConsignes : TOB;// JTR - eQualité 11023
    procedure ClickValide; dynamic;
    // Afficheur
    procedure AffichageLCD(Texte: string; Qte, Mnt: Double; NbDec: Integer; Sigle: string; Ou: TSetOuAffiche; Quoi: TQuoiAffiche; Attente: Boolean);
    // Ecran tactiles
    procedure InitFormEcheance;
    function IndiqueOuSaisie(EcheanceOk: boolean = False): TOuSaisie;
    procedure AffichePaveChoisi(ACol: Integer);
    // Détaxe
    function DetaxeActive: boolean;
    procedure TraiteDemandeDetaxe;
    procedure ValideDetaxe;
    function InsereLigneGEN(CodeGen: string; ARow: integer; Qte: Double): TOB;
    function NbDimensionsLigneGEN(ARow: integer): integer;
  end;

implementation
uses
  Ent1, TiersUtil, StockUtil, AdressePiece, UtilPGI, UtilArticle,
  FactComm, FactCalc, FactSpec, FactNomen, LigNomen,
  {$IFNDEF CCS3}
    {$IFNDEF STK}
      SaisieSerie_TOF,
    {$ENDIF !STK}
  {$ENDIF}
  {$IFNDEF STK}
    LigDispoLot,
  {$ENDIF !STK}
  UTofPiedPort, UTofGCPieceArtLie,
  DepotUtil,
  {$IFNDEF SANSCOMPTA}
  Devise_tom,
  cpChancell_tof,
  {$ENDIF}
  MC_Erreur, FOConsultCaisse, MFOSOLDECLIENT_TOF,
  {$IFNDEF GESCOM}
  FOGbRefund,
  {$ENDIF}
  FOChoixPave, FODefi, FOUtil,
  MFORECHARTIMAGE_TOF, MFODETAILCB_TOF,
  {$IFDEF TAXEUS}
  MBOTAXATIONLIGNE_TOF,
  {$ENDIF}
  yTarifs,
  FactTarifs
  {$IFDEF STK}
  ,DispoDetail,Dispo,StkNature, GCSTKNomenclature_TOF
  {$ENDIF STK}
  ;

// Gestion des Timers
const MaxInterval = 5000;
  MinInterval = 500;
  ToMaximise = 100;

const NbArtParRequete: integer = 200;

{$IFDEF GESCOM} // JTR 15/01/2004
var TypePiece, TypePieceFi : string;
{$ENDIF GESCOM}

  {$R *.DFM}

function CreerTicketFO(NaturePiece: string): boolean;
var CleDoc: R_CleDoc;
begin
  Result := True;
  if PasCreerDateGC(V_PGI.DateEntree) then Exit;
  FillChar(CleDoc, Sizeof(CleDoc), #0);
  CleDoc.NaturePiece := NaturePiece;
  CleDoc.DatePiece := V_PGI.DateEntree;
  CleDoc.Souche := '';
  CleDoc.NumeroPiece := 0;
  CleDoc.Indice := 0;
  SaisieTicketFO(CleDoc, taCreat);
end;

function SaisieTicketFO(CleDoc: R_CleDoc; Action: TActionFiche): boolean;
var
  X: TFTicketFO;
  PP: THPanel;
  ModalOk : boolean;
  {$IFNDEF EAGL}
  DrawXP : boolean;
  {$ENDIF}
begin
  {$IFNDEF EAGL}
    DrawXP:=V_PGI.DrawXP;
    V_PGI.DrawXP:=False;
  {$ENDIF}
  Result := True;
  SourisSablier;
  PP := FindInsidePanel;
  X := TFTicketFO.Create(Application);
  X.PieceContainer.CleDoc := CleDoc;
  X.PieceContainer.Action := Action;
  X.PieceContainer.NewNature := X.PieceContainer.CleDoc.NaturePiece;
  X.PieceContainer.TransfoPiece := False;
  X.PieceContainer.DuplicPiece := False;
  ModalOk := True; // pas de mode inside
  if (ModalOk) or (PP = nil) then
  begin
    try
      X.ShowModal;
    finally
      X.Free;
    end;
    SourisNormale;
  end else
  begin
    InitInside(X, PP);
    X.Show;
  end;
  {$IFNDEF EAGL}
    V_PGI.DrawXP:=DrawXP;
  {$ENDIF}
end;

procedure AppelTicketFO(Parms: array of variant; nb: integer);
var CleDoc: R_CleDoc;
  StA, StC, StM: string;
  i_ind: integer;
  Action: TActionFiche;
begin
  FillChar(CleDoc, Sizeof(CleDoc), #0);
  StA := string(Parms[0]);
  StM := string(Parms[1]);

  Action := taModif;
  i_ind := Pos('ACTION=', StM);
  if i_ind > 0 then
  begin
    Delete(StM, 1, i_ind + 6);
    StM := uppercase(ReadTokenSt(StM));
    if StM = 'CREATION' then
    begin
      Action := taCreat;
    end;
    if StM = 'MODIFICATION' then
    begin
      Action := taModif;
    end;
    if StM = 'CONSULTATION' then
    begin
      Action := taConsult;
    end;
  end;
  if Action = tacreat then
  begin
    CreerTicketFO(StC);
  end else
  begin
    StringToCleDoc(StA, CleDoc);
    if CleDoc.NaturePiece <> '' then SaisieTicketFO(CleDoc, Action);
  end
end;

function TransformeTicketFO(CleDoc: R_CleDoc; NewNat: string): boolean;
var
  X: TFTicketFO;
  PP: THPanel;
  {$IFNDEF EAGL}
  DrawXP: boolean;
  {$ENDIF}
begin
  {$IFNDEF EAGL}
    DrawXP:=V_PGI.DrawXP;
    V_PGI.DrawXP:=False;
  {$ENDIF}
  Result := True;
  SourisSablier;
  X := TFTicketFO.Create(Application);
  X.PieceContainer.CleDoc := CleDoc;
  X.PieceContainer.Action := taModif;
  X.PieceContainer.NewNature := NewNat;
  X.PieceContainer.TransfoPiece := True;
  X.PieceContainer.DuplicPiece := False;
  PP := FindInsidePanel;
  if ((PP = nil) or (ctxFO in V_PGI.PGIContexte)) then
  begin
    try
      X.ShowModal;
    finally
      X.Free;
    end;
    SourisNormale;
  end else
  begin
    InitInside(X, PP);
    X.Show;
  end;
  {$IFNDEF EAGL}
    V_PGI.DrawXP:=DrawXP;
  {$ENDIF}
end;

procedure AppelTransformeTicketFO(Parms: array of variant; nb: integer);
var CleDoc: R_CleDoc;
  StA, StM, NewNat: string;
begin
  FillChar(CleDoc, Sizeof(CleDoc), #0);
  StA := string(Parms[0]);
  StM := string(Parms[1]);
  StringToCleDoc(STA, CleDoc);
  NewNat := StM;
  if CleDoc.NaturePiece <> '' then TransformeTicketFO(CleDoc, NewNat);
end;

function DupliqueTicketFO(CleDoc: R_CleDoc; NewNat: string): boolean;
var
  X: TFTicketFO;
  PP: THPanel;
  {$IFNDEF EAGL}
  DrawXP: boolean;
  {$ENDIF}
begin
  {$IFNDEF EAGL}
  DrawXP:=V_PGI.DrawXP;
  V_PGI.DrawXP:=False;
  {$ENDIF}
  Result := True;
  SourisSablier;
  X := TFTicketFO.Create(Application);
  X.PieceContainer.CleDoc := CleDoc;
  X.PieceContainer.Action := taModif;
  X.PieceContainer.NewNature := NewNat;
  X.PieceContainer.TransfoPiece := False;
  X.PieceContainer.DuplicPiece := True;
  PP := FindInsidePanel;
  if ((PP = nil) or (ctxFO in V_PGI.PGIContexte)) then
  begin
    try
      X.ShowModal;
    finally
      X.Free;
    end;
    SourisNormale;
  end else
  begin
    InitInside(X, PP);
    X.Show;
  end;
  {$IFNDEF EAGL}
    V_PGI.DrawXP:=DrawXP;
  {$ENDIF}
end;

procedure AppelDupliqueTicketFO(Parms: array of variant; nb: integer);
var CleDoc: R_CleDoc;
  StA, StM, NewNat: string;
begin
  FillChar(CleDoc, Sizeof(CleDoc), #0);
  StA := string(Parms[0]);
  StM := string(Parms[1]);
  StringToCleDoc(STA, CleDoc);
  NewNat := StM;
  if CleDoc.NaturePiece <> '' then DupliqueTicketFO(CleDoc, NewNat);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/01/2002
Modifié le ... : 10/01/2002
Description .. : Annulation d'un ticket par génération d'une pièce inverse.
Mots clefs ... : FO
*****************************************************************}

function AnnuleTicketFO(CleDoc: R_CleDoc; NewNat: string): boolean;
var
  X: TFTicketFO;
  PP: THPanel;
  {$IFNDEF EAGL}
  DrawXP: boolean;
  {$ENDIF}
begin
  {$IFNDEF EAGL}
    DrawXP:=V_PGI.DrawXP;
    V_PGI.DrawXP:=False;
  {$ENDIF}
  Result := False;
  if not FOVerifTicketAnnulable(CleDoc) then Exit;
  SourisSablier;
  X := TFTicketFO.Create(Application);
  X.PieceContainer.CleDoc := CleDoc;
  X.PieceContainer.Action := taConsult;
  X.PieceContainer.NewNature := NewNat;
  X.PieceContainer.TransfoPiece := False;
  X.PieceContainer.DuplicPiece := True;
  X.AnnulPiece := True;
  PP := FindInsidePanel;
  if ((PP = nil) or (ctxFO in V_PGI.PGIContexte)) then
  begin
    try
      X.ShowModal;
    finally
      X.Free;
    end;
    SourisNormale;
  end else
  begin
    InitInside(X, PP);
    X.Show;
  end;
  Result := True;
  {$IFNDEF EAGL}
    V_PGI.DrawXP:=DrawXP;
  {$ENDIF}
end;

procedure AppelAnnuleTicketFO(Parms: array of variant; nb: integer);
var CleDoc: R_CleDoc;
  StA, StM, NewNat: string;
begin
  FillChar(CleDoc, Sizeof(CleDoc), #0);
  StA := string(Parms[0]);
  StM := string(Parms[1]);
  StringToCleDoc(STA, CleDoc);
  NewNat := StM;
  if NewNat = '' then NewNat := CleDoc.NaturePiece;
  if CleDoc.NaturePiece <> '' then AnnuleTicketFO(CleDoc, NewNat);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/01/2002
Modifié le ... : 10/01/2002
Description .. : Assignation d'un ticket à un client, c'est-à-dire changement
Suite ........ : du client d'un ticket.
Mots clefs ... : FO
*****************************************************************}

function AssigneTicketFO(CleDoc: R_CleDoc; NewNat: string): boolean;
var
  X: TFTicketFO;
  PP: THPanel;
  {$IFNDEF EAGL}
    DrawXP: boolean;
  {$ENDIF}
begin
  {$IFNDEF EAGL}
    DrawXP:=V_PGI.DrawXP;
    V_PGI.DrawXP:=False;
  {$ENDIF}
  Result := False;
  if VH_GC.TOBPCaisse.GetValue('GPK_CLISAISIE') <> 'X' then
  begin
    PGIBox('La gestion des clients n''est pas activée pour cette caisse !', 'Rattachement d''un ticket à un client');
    Exit;
  end;
  if FOComptaEstActive(CleDoc.NaturePiece) then
  begin
    PGIBox('Opération impossible car la comptabilité est alimentée depuis la caisse ', 'Rattachement d''un ticket à un client');
    Exit;
  end;
  Result := True;
  SourisSablier;
  X := TFTicketFO.Create(Application);
  X.PieceContainer.CleDoc := CleDoc;
  X.PieceContainer.Action := taConsult;
  X.PieceContainer.NewNature := NewNat;
  X.PieceContainer.TransfoPiece := False;
  X.PieceContainer.DuplicPiece := False;
  X.AssignePiece := True;
  PP := FindInsidePanel;
  if ((PP = nil) or (ctxFO in V_PGI.PGIContexte)) then
  begin
    try
      X.ShowModal;
    finally
      X.Free;
    end;
    SourisNormale;
  end else
  begin
    InitInside(X, PP);
    X.Show;
  end;
  {$IFNDEF EAGL}
    V_PGI.DrawXP:=DrawXP;
  {$ENDIF}
end;

procedure AppelAssigneTicketFO(Parms: array of variant; nb: integer);
var CleDoc: R_CleDoc;
  StA, StM, NewNat: string;
begin
  FillChar(CleDoc, Sizeof(CleDoc), #0);
  StA := string(Parms[0]);
  StM := string(Parms[1]);
  StringToCleDoc(STA, CleDoc);
  NewNat := StM;
  if NewNat = '' then NewNat := CleDoc.NaturePiece;
  if CleDoc.NaturePiece <> '' then AssigneTicketFO(CleDoc, NewNat);
end;

procedure TFTicketFO.FindLigneFind(Sender: TObject);
begin
  Rechercher(GS, FindLigne, FindFirst);
end;

procedure TFTicketFO.BChercherClick(Sender: TObject);
begin
  if GS.RowCount <= 2 then Exit;
  FindFirst := True;
  FindLigne.Execute;
end;

{==============================================================================================}
{======================================= Initialisations ======================================}
{==============================================================================================}

procedure TFTicketFO.ReInitPiece;
begin
  {$IFDEF CHR}
  TransfertFac:=False;
  {$ENDIF}
  ForcerFerme := False;
  GeneCharge := False;
  ToutLiberer;
  ToutAllouer;
  BlocageSaisie(False);
  TOBPiece.PutEcran(Self);
  CliCur := '';
  LibelleArticle.Caption := '';
  VenteALivrer := False;
  DejaSaisie := False;
  RetourGS := False;
  DemandeDetaxe := False;
  SurLigSuivante := False;
  DepuisEcheance := False;
  InTraiteBouton := False;
  GP_TIERS.Text := '';
  LibelleTiers.Caption := '';
  SOLDETIERS.Caption := '';
  TSOLDETIERS.Visible := False;
  DejaRentre := False;
  GP_ESCOMPTE.Value := 0;
  GP_REMISEPIED.Value := 0;
  GP_MODEREGLE.Value := '';
  InitEnteteDefaut(False);
  InitPieceCreation;
  {$IFDEF GESCOM}
  //IncidenceTiers;
  TypePiece := ''; //JTR 15/01/2004
  TypePieceFi := '';
  {$ENDIF}
  GS.ElipsisButton := False;
  GS.Row := GS.FixedRows;
  AffecteGrid(GS, PieceContainer.Action);
  if PieceContainer.Action = taConsult then GS.MultiSelect := False else GS.Col := SG_RefArt;
  T_PrixU.Caption := '';
  A_QTESTOCK.Caption := '';
  if (VH_GC.TOBPCaisse.GetValue('GPK_PREPOSGRID') = 'X') or (not GoToEntete) then
  begin
    if Screen.ActiveControl = GS then
      GSEnter(GS)
    else if GS.CanFocus then
      GS.SetFocus;
  end;
  {$IFDEF CHR}
  GSEnter(GS) ;
  {$ENDIF}
  if (FOAfficheImageArt) or (FOLogoCaisse <> '') then FOAffichePhotoArticle('', ImageArticle);
end;

function TFTicketFO.ReInitUtilisateur: boolean;
{***
var
  OldPass: string;
***}
begin
  Result := True;
{***
  OldPass := V_PGI.PassWord;
  V_PGI.PassWord := '';
  if not ChangeUser then
  begin
    Result := False;
    V_PGI.PassWord := OldPass;
    Close;
    Exit;
  end;
***}
end;

procedure TFTicketFO.InitPasModif;
begin
  TOBPiece.SetAllModifie(False);
  TOBAdresses.SetAllModifie(False);
  TOBBases.SetAllModifie(False);
  TOBEches.SetAllModifie(False);
  TOBAcomptes.SetAllModifie(False);
  TOBPorcs.SetAllModifie(False);
  TOBTiers.SetAllModifie(False);
  TOBAnaP.SetAllModifie(False);
  TOBAnaS.SetAllModifie(False);
  TOBDesLots.SetAllModifie(False);
  TOBSerie.SetAllModifie(False);
  TOBOpCais.SetAllModifie(False);
end;

procedure TFTicketFO.InitToutModif;
var NowFutur: TDateTime;
begin
  NowFutur := NowH;
  TOBPiece.SetAllModifie(True);
  TOBPiece.SetDateModif(NowFutur);
  TOBAdresses.SetAllModifie(True);
  TOBBases.SetAllModifie(True);
  TOBEches.SetAllModifie(True);
  TOBAcomptes.SetAllModifie(True);
  TOBPorcs.SetAllModifie(True);
  InvalideModifTiersPiece(TOBTiers);
  TOBAnaP.SetAllModifie(True);
  TOBAnaS.SetAllModifie(True);
  TOBDesLots.SetAllModifie(True);
  TOBSerie.SetAllModifie(True);
  TOBOpCais.SetAllModifie(True);
  TobLigneTarif.SetAllModifie(True);
end;

procedure PieceAjouteSousDetailFO(FTOBPiece: TOB);
var i: integer;
begin
  if FTOBPiece.Detail.Count > 0 then
  begin
    AddLesSupLigne(FTOBPiece.Detail[0], True);
    for i := 0 to FTOBPiece.Detail.Count - 1 do TOB.Create('', FTOBPiece.Detail[i], -1);
  end;
end;

procedure TFTicketFO.LoadLesTOB;
var Q: TQuery;
  sSQL: string;
begin
  // Lecture Pied
  Q := OpenSQL('SELECT * FROM PIECE WHERE ' + WherePiece(PieceContainer.CleDoc, ttdPiece, False), True);
  TOBPiece.SelectDB('', Q);
  Ferme(Q);
  // Lecture Lignes
  Q := OpenSQL('SELECT * FROM LIGNE WHERE ' + WherePiece(PieceContainer.CleDoc, ttdLigne, False) + ' ORDER BY GL_NUMLIGNE', True);
  TOBPiece.LoadDetailDB('LIGNE', '', '', Q, False, True);
  Ferme(Q);
  PieceAjouteSousDetail(TOBPiece);
  // Lecture bases
  Q := OpenSQL('SELECT * FROM PIEDBASE WHERE ' + WherePiece(PieceContainer.CleDoc, ttdPiedBase, False), True);
  TOBBases.LoadDetailDB('PIEDBASE', '', '', Q, False);
  Ferme(Q);
  // Lecture Echéances
  //Q:=OpenSQL('SELECT * FROM PIEDECHE WHERE '+WherePiece(CleDoc,ttdEche,False),True) ;
  sSQL := 'SELECT PIEDECHE.*,MP_TYPEMODEPAIE AS TYPEMODEPAIE FROM PIEDECHE LEFT JOIN MODEPAIE ON GPE_MODEPAIE=MP_MODEPAIE WHERE ' + WherePiece(PieceContainer.CleDoc, ttdEche, False);
  Q := OpenSQL(sSQL, True);
  TOBEches.LoadDetailDB('PIEDECHE', '', '', Q, False);
  Ferme(Q);
  // Lecture Ports
  Q := OpenSQL('SELECT * FROM PIEDPORT WHERE ' + WherePiece(PieceContainer.CleDoc, ttdPorc, False), True);
  TOBPorcs.LoadDetailDB('PIEDPORT', '', '', Q, False);
  Ferme(Q);
  // Lecture Adresses
  LoadTOBAdresses;
  // Lecture Nomenclatures
  LoadLesNomen(PieceContainer);
  {$IFNDEF STK}
  LoadLesLots(PieceContainer);
  {$ENDIF STK}
  {$IFNDEF CCS3}
  // Lecture Serie
  //LoadLesSeries(TOBPiece, TOBSerie, CleDoc);
  {$ENDIF}
  // Lecture ACompte
  LoadLesAcomptes(PieceContainer);
  // Lecture Analytiques
  LoadLesAna(PieceContainer);
  // Opérations de caisse
  FOLoadOperCaisse(TOBPiece, TOBEches, TOBOpCais);
  { Chargement des détails tarif par ligne de pièce }
//  LoadLesLignesTarifs(TobLigneTarif, PieceContainer.CleDoc);
end;


procedure TFTicketFO.LoadTOBAdresses;
begin
  if not GereAdresse then Exit;
  LoadLesAdresses(PieceContainer);
  if (TOBAdresses = nil) or (TOBAdresses.Detail.Count < 2) then
  begin
    // cas d'une pièce sans adresse
    if TOBAdresses <> nil then TOBAdresses.Free;
    CreerTOBAdresses;
  end;
end;

procedure TFTicketFO.TicketALivrer;
begin
  if PieceContainer.Action = taConsult then Exit;
  if not GereAdresse then Exit;
  if FOGetParamCaisse('GPK_CLISAISIE') = '-' then Exit;
  if (IndiqueOuSaisie(True) = saEcheance) or (IndiqueOuSaisie = saRemiseTicket) then Exit;

  if VenteALivrer then
  begin
    LALIVRER.Visible := False;
    VenteALivrer := False;
  end else
  begin
    if not VerifClientOblig(57) then Exit;
    if not ComplementEntete then Exit;

    LALIVRER.Visible := True;
    VenteALivrer := True;
  end;
end;

procedure TFTicketFO.LoadTOBCond(RefUnique: string);
var Q: TQuery;
begin
  if TOBConds.FindFirst(['GCO_ARTICLE'], [RefUnique], False) <> nil then Exit;
  Q := OpenSQL('SELECT * FROM CONDITIONNEMENT WHERE GCO_ARTICLE="' + RefUnique +
                '" AND GCO_TYPESFLUX LIKE "%' + VenteAchat + '%" ORDER BY GCO_NBARTICLE', True);
  if not Q.EOF then TOBConds.LoadDetailDB('CONDITIONNEMENT', '', '', Q, True, False);
  Ferme(Q);
end;

procedure TFTicketFO.LoadLesArticles;
var i, iArt, iRefCata: integer;
  TOBL, TOBArt, TOBC, TOBCata, LocAnaP, LocAnaS: TOB;
  RefUnique, RefCata, RefFour: string;
  NaturePieceG: string;
  OldCleDoc, NewCleDoc : R_CleDoc;
begin
  TOBCata := nil;
  iArt := 0;
  iRefCata := 0;
  NaturePieceG := TOBPiece.GetValue('GP_NATUREPIECEG');
//  ConstruireTobArt(TOBPiece);
  if TobTicketAttente <> nil then
  begin
    OldCleDoc := PieceContainer.CleDoc;
    NewCleDoc.NaturePiece := TobTicketAttente.GetString('GP_NATUREPIECEG');
    NewCleDoc.DatePiece := TobTicketAttente.GetDateTime('GP_DATEPIECE');
    NewCleDoc.Souche := TobTicketAttente.GetString('GP_SOUCHE');
    NewCleDoc.NumeroPiece := TobTicketAttente.GetInteger('GP_NUMERO');
    NewCleDoc.Indice := 0;
    PieceContainer.CleDoc := NewCledoc;
  end;
  ConstruireTobArt(PieceContainer);
  if TobTicketAttente <> nil then
  begin                                                          
    PieceContainer.CleDoc := OldCledoc;
    TobTicketAttente.DeleteDB;
    FreeAndNil(TobTicketAttente);
  end;
  InitMove(TOBPiece.Detail.Count, '');
  for i := 0 to TOBPiece.Detail.Count - 1 do
  begin
    TOBL := TOBPiece.Detail[i];
    MoveCur(False);
    if i = 0 then
    begin
      iRefCata := TOBL.GetNumChamp('GL_REFCATALOGUE');
      iArt := TOBL.GetNumChamp('GL_ARTICLE');
    end;
    InitLesSupLigne(TOBL);
    // Récupération article
    RefCata := '';
    RefFour := '';
    RefUnique := '';
    if (TOBL.GetValue('GL_ENCONTREMARQUE') = 'X') then
    begin
      RefCata := TOBL.GetValeur(iRefCata);
      RefFour := GetCodeFourDCM(TOBL);
      TOBCata := TOBCatalogu.FindFirst(['GCA_REFERENCE', 'GCA_TIERS'], [RefCata, RefFour], False);
      if TOBCata = nil then
        TOBCata := InitTOBCata(TOBCatalogu, RefCata, RefFour);
      LoadTOBDispoContreM(TOBL, TOBCata, False);
    end;
    RefUnique := TOBL.GetValeur(iArt);
    TobArt := TOBArticles.FindFirst(['GA_ARTICLE'], [RefUnique], False);
    if TobArt <> nil then
    begin
      if (VenteAchat = 'ACH') and ((CataFourn) or (TOBL.GetValue('GL_REFARTTIERS') <> '')) then AjouteCatalogueArt(TOBL, TOBArt);
      CodesLigneToCodesArt(TOBL, TOBArt);
      TOBL.PutValue('QTEORIG', TOBL.GetValue('GL_QTESTOCK'));
    end;
    if (RefUnique <> '') or ((RefCata <> '') and (RefFour <> '')) then
    begin
      // Chargement du commercial
      AjouteRepres(TOBL.GetValue('GL_REPRESENTANT'), TypeCom, TOBComms);
      // Compta
      TOBC := ChargeAjouteCompta(TOBCpta, TOBPiece, TOBL, TOBArt, TOBTiers, TOBCata, TobAffaire, True);
      if ((OkCpta) or (OkCptaStock)) then
      begin
        if (TOBC <> nil) and (TOBL.Detail.Count > 0) then
        begin
          if ((TOBL.Detail[0].Detail.Count <= 0) and (OkCpta)) then LocAnaP := TOBAnaP else LocAnaP := nil;
          if ((TOBL.Detail[1].Detail.Count <= 0) and (OkCptaStock)) then LocAnaS := TOBAnaS else LocAnaS := nil;
          PreVentileLigne(TOBC, LocAnaP, LocAnaS, TOBL);
        end;
      end;
    end;
    // Affichage
    AfficheLaLigne(i + 1);
  end;
  FiniMove;
  if (ctxMode in V_PGI.PGIContexte) and ((PieceContainer.Action = taConsult) or (PieceContainer.Action = taModif)) then AffichageDim;
end;

(*
procedure TFTicketFO.ConstruireTobArt(TOBPiece: Tob);
var i, NbArt, x, CountStArt, NbRequete, y: integer;
  StSelect, StSelectDepot, StSelectCon, StArticle, StCodeArticle, ListeDepot, RefGen, Statut: string;
  TabWhere, TabWhereDepot, TabWhereCon, TabWhereFormule : array of string;
  TOBDispo, TOBArt, TobDispoArt, TobAF: TOB;
  QArticle, QDepot, QCon, QFormule: TQuery;
  OkTab: Boolean;
  StSelectFormule : string; // JTR - eQualité 11023
begin
  StSelect := 'SELECT * FROM ARTICLE';
  StSelectDepot := 'SELECT * FROM DISPO';
  StSelectCon := 'SELECT * FROM CONDITIONNEMENT';
  StSelectFormule := 'SELECT * FROM ARTICLEQTE'; // JTR - eQualité 11203 (formules de qté)
  CountStArt := 0;
  NbRequete := 0;
  ListeDepot := '';
  SetLength(TabWhere, 1);
  SetLength(TabWhereDepot, 1);
  SetLength(TabWhereCon, 1);
  SetLength(TabWhereFormule, 1); // JTR - eQualité 11203 (formules de qté)
  if (TOBPiece.GetValue('GP_NATUREPIECEG') = 'TEM') or (TOBPiece.GetValue('GP_NATUREPIECEG') = 'TRE') or (TOBPiece.GetValue('GP_NATUREPIECEG') = 'TRV') then
    ListeDepot := '"' + TOBPiece.GetValue('GP_DEPOT') + '","' + TOBPiece.GetValue('GP_DEPOTDEST') + '"'
    else
    ListeDepot := '"' + TOBPiece.GetValue('GP_DEPOT') + '"';
  for i := 0 to TOBPiece.Detail.Count - 1 do
  begin
    StArticle := TOBPiece.Detail[i].GetValue('GL_ARTICLE');
    StCodeArticle := TOBPiece.Detail[i].GetValue('GL_CODEARTICLE');
    RefGen := TOBPiece.Detail[i].GetValue('GL_CODESDIM');
    Statut := TOBPiece.Detail[i].GetValue('GL_TYPEDIM');
    OkTab := False;
    if CountStArt >= NbArtParRequete then
    begin
      NbRequete := NbRequete + 1;
      SetLength(TabWhere, NbRequete + 1);
      SetLength(TabWhereDepot, NbRequete + 1);
      SetLength(TabWhereCon, NbRequete + 1);
      SetLength(TabWhereFormule, NbRequete + 1); // JTR - eQualité 11203 (formules de qté)
      CountStArt := 0;
    end;
    if (Statut = 'GEN') or (Statut = 'DIM') or (Statut = 'NOR') then
    begin
      if (StArticle = '') and (RefGen <> '') then
        RefGen := CodeArticleUnique2(RefGen, '');
      if StArticle <> '' then
        RefGen := StArticle;
      if TabWhere[NbRequete] = '' then
        TabWhere[NbRequete] := '"' + RefGen + '"'
        else
        TabWhere[NbRequete] := TabWhere[NbRequete] + ',"' + RefGen + '"';
      OkTab := True;
    end;
    if (Statut = 'DIM') or (Statut = 'NOR') then
    begin
      if TabWhereDepot[NbRequete] = '' then
        TabWhereDepot[NbRequete] := '"' + StArticle + '"'
        else
        TabWhereDepot[NbRequete] := TabWhereDepot[NbRequete] + ',"' + StArticle + '"';
      if TabWhereCon[NbRequete] = '' then
        TabWhereCon[NbRequete] := '"' + StArticle + '"'
        else
        TabWhereCon[NbRequete] := TabWhereCon[NbRequete] + ',"' + StArticle + '"';
      if TabWhereFormule[NbRequete] = '' then
        TabWhereformule[NbRequete] := '"' + StArticle + '"'
        else
        TabWhereFormule[NbRequete] := TabWhereFormule[NbRequete] + ',"' + StArticle + '"';
      OkTab := True;
    end;
    if OkTab then CountStArt := CountStArt + 1;
  end;
  if TabWhere[0] <> '' then
  begin
    TOBDispo := TOB.CREATE('Les Dispos', nil, -1);
    for y := Low(TabWhere) to High(TabWhere) do
    begin
      if TabWhere[y] <> '' then
      begin
        QArticle := OpenSQL(StSelect + ' WHERE GA_ARTICLE IN (' + TabWhere[y] + ')', True);
        if not QArticle.EOF then
          TOBArticles.LoadDetailDB('ARTICLE', '', '', QArticle, True);
        Ferme(QArticle);
      end;
      if TabWhereDepot[y] <> '' then
      begin
        if CtxMode in V_PGI.PGIContexte then
          QDepot := OpenSQL(StSelectDepot + ' WHERE GQ_ARTICLE IN (' + TabWhereDepot[y] + ') AND GQ_DEPOT IN (' + ListeDepot + ') AND GQ_CLOTURE="-"', True)
          else
          QDepot := OpenSQL(StSelectDepot + ' WHERE GQ_ARTICLE IN (' + TabWhereDepot[y] + ') AND GQ_CLOTURE="-"', True);
        if not QDepot.EOF then
          TOBDispo.LoadDetailDB('DISPO', '', '', QDepot, True);
        Ferme(QDepot);
      end;
      if TabWhereCon[y] <> '' then
      begin
        QCon := OpenSQL (StSelectCon + ' WHERE GCO_ARTICLE IN (' + TabWhereCon[y] +
                ') AND GCO_TYPESFLUX LIKE "%' + VenteAchat + '%" ORDER BY GCO_NBARTICLE', True);
        if not QCon.EOF then TOBConds.LoadDetailDB('CONDITIONNEMENT', '', '', QCon, True);
        Ferme(QCon);
      end;
      if TabWhereformule[y] <> '' then  // JTR - eQualité 11203 (formules de qté)
      begin
        QFormule := OpenSQL (StSelectFormule + ' WHERE GAF_ARTICLE IN (' + TabWhereFormule[y] + ')', True);
        if not QFormule.EOF then
          TobAFormule.LoadDetailDB('ARTICLEQTE', '', '', QFormule, True);
        Ferme(QFormule);
      end;
    end;
    for x := 0 to TOBArticles.detail.count - 1 do
    begin
      TOBArticles.detail[x].AddChampSup('UTILISE', False);
      TOBArticles.detail[x].PutValue('UTILISE', '-');
      TOBArticles.detail[x].AddChampSup('REFARTSAISIE', False);
      TOBArticles.detail[x].AddChampSup('REFARTBARRE', False);
      TOBArticles.detail[x].AddChampSup('REFARTTIERS', False);
    end;
    //Affecte les stocks aux articles sélectionnés
    for NbArt := 0 to TOBArticles.detail.Count - 1 do
    begin
      TOBArt := TOBArticles.detail[NbArt];
      TobDispoArt := TOBDispo.FindFirst(['GQ_ARTICLE'], [TOBArt.GetValue('GA_ARTICLE')], False);
      while TobDispoArt <> nil do
      begin
        TobDispoArt.Changeparent(TOBArt, -1);
        TobDispoArt := TOBDispo.FindNext(['GQ_ARTICLE'], [TOBArt.GetValue('GA_ARTICLE')], False);
      end;
    end;
    TOBDispo.Free;
  end;
  // JTR - eQualité 11203 (formules de qté)
  for x := 0 to TobAFormule.Detail.Count -1 do
  begin
    TobAF := TobAFormule.Detail[x];
    QFormule := OpenSql('SELECT * FROM ARTFORMULEVAR WHERE GAV_ARTICLE="'+
                        TobAF.GetValue('GAF_ARTICLE')+'" AND GAV_VENTEACHAT="'+VenteAchat+'"',True);
    if not QFormule.Eof then
      TobAF.LoadDetailDB('ARTICLEQTE','','',QFormule,False) ;
    Ferme(QFormule);
  end;
  // Fin JTR
end;
*)
{$IFNDEF FOS5}

procedure TFTicketFO.AppliqueTransfoDuplic;
var i: integer;
  TOBL, TOBA, TOBB: TOB;
  OldNat, ColPlus, ColMoins, EtatVC, RefPiece: string;
  MajLotZero, Galop, MajSerieZero, GaSer, LotSerieTrouve: boolean;
  bMajZeroQte : boolean;
  dQteStock : double;
  PEntSup : double;
  TobC : TOB;
  TobLT: TOB;
begin
  if ((not PieceContainer.TransfoPiece) and (not DuplicPiece)) then Exit;
  PieceContainer.Cledoc.NaturePiece := PieceContainer.NewNature;
  PieceContainer.Cledoc.Souche := GetSoucheG(PieceContainer.Cledoc.NaturePiece, TOBPiece.GetValue('GP_ETABLISSEMENT'), '');
  PieceContainer.Cledoc.NumeroPiece := GetNumSoucheG(PieceContainer.Cledoc.Souche);
  PieceContainer.Cledoc.Indice := 0;
  PieceContainer.Cledoc.DatePiece := V_PGI.DateEntree;
  GP_DATEPIECE.Text := DateToStr(PieceContainer.Cledoc.DatePiece);
  GP_NUMEROPIECE.Caption := HTitres.Mess[10];
  MajFromCleDoc(TOBPiece, PieceContainer.Cledoc);
  MajLotZero := False;
  TOBPiece.PutValue('GP_DEVENIRPIECE', '');
  TOBPiece.PutValue('GP_REFCOMPTABLE', '');
  TOBPiece.PutValue('GP_ETATVISA', 'NON');
  TOBPiece.PutValue('GP_EDITEE', '-');
  if DuplicPiece then
  begin
    TOBPiece.PutValue('GP_VIVANTE', 'X');
    TOBPiece.PutValue('GP_ACOMPTE', 0);
    TOBPiece.PutValue('GP_ACOMPTEDEV', 0);
    TOBPiece.PutValue('GP_ACOMPTECON', 0);
    ClearAffaire(TOBPiece);
    TOBAcomptes.ClearDetail;
  end else
  begin
    if ((GereLot) and (TOBPiece_O.GetValue('GP_ARTICLESLOT') <> 'X')) then MajLotZero := True;
    if ((GereSerie) and (TOBSerie_O.Detail.Count <= 0)) then MajSerieZero := True;
    GereEcheancesGC(PieceContainer, False);
    {Etudier modif dépôt O/N}
    OldNat := TOBPiece_O.GetValue('GP_NATUREPIECEG');
    ColPlus := GetInfoParPiece(OldNat, 'GPP_QTEPLUS');
    ColMoins := GetInfoParPiece(OldNat, 'GPP_QTEMOINS');
    if ((Pos('PHY', ColPlus) > 0) or (Pos('PHY', ColMoins) > 0)) then PasToucheDepot := True;
    if PasToucheDepot then GP_DEPOT.Enabled := False;
  end;
  for i := 0 to TOBPiece.Detail.Count - 1 do
  begin
    TOBL := TOBPiece.Detail[i];
    MajFromCleDoc(TOBL, PieceContainer.Cledoc);
    bMajZeroQte := false;
    dQteStock := Tobl.GetValue ('GL_QTERESTE');
    if DuplicPiece then
    begin
      TOBL.PutValue('GL_PIECEORIGINE', '');
      TOBL.PutValue('GL_PIECEPRECEDENTE', '');
      TOBL.PutValue('GL_QTERELIQUAT', 0);
      TOBL.PutValue('GL_SOLDERELIQUAT', '-');
      TOBL.PutValue('GL_QTERESTE', 0);
      TOBL.PutValue('GL_VALIDECOM', 'NON');
{$IFDEF V500}
      CommVersLigne(TOBPiece, TOBArticles, TOBComms, i + 1, True);
{$ELSE}
      CommVersLigne(PieceContainer, i + 1, True);
{$ENDIF}
      TOBL.PutValue('GL_VIVANTE', 'X');
      ClearAffaire(TOBL);
    end else
    begin
      RefPiece := EncodeRefPiece(TOBPiece_O.Detail[i]);
      TOBL.PutValue('GL_PIECEPRECEDENTE', RefPiece);
      if TOBL.GetValue('GL_PIECEORIGINE') = '' then TOBL.PutValue('GL_PIECEORIGINE', RefPiece);
      EtatVC := TOBL.GetValue('GL_VALIDECOM');
      if      (EtatVC = 'VAL') then TOBL.PutValue('GL_VALIDECOM', 'AFF')
{$IFDEF V500}
      else if (EtatVC = 'AFF') then CommVersLigne(TOBPiece, TOBArticles, TOBComms, i + 1, False)
{$ELSE}
      else if (EtatVC = 'AFF') then CommVersLigne(PieceContainer, i + 1, False)
{$ENDIF}
      ;
      {Lots}
      if MajLotZero then
      begin
        TOBA := FindTOBArtRow(TOBPiece, TOBArticles, i + 1);
        if TOBA <> nil then Galop := (TOBA.GetValue('GA_LOT') = 'X') else Galop := False;
        if Galop then
        begin
          if GereReliquat then TOBL.PutValue('GL_QTERELIQUAT', dQteStock); // TOBL.GetValue('GL_QTESTOCK'));
          TOBL.PutValue('GL_QTEFACT', 0);
          TOBL.PutValue('GL_QTESTOCK', 0);
          LotSerieTrouve := True;
          bMajZeroQte := True;
        end;
      end;
      {Série}
      if MajSerieZero then
      begin
        GaSer := (TOBL.GetValue('GL_NUMEROSERIE') = 'X');
        if not (GaSer) then
        begin
          TOBA := FindTOBArtRow(TOBPiece, TOBArticles, i + 1);
          if TOBA <> nil then GaSer := (TOBA.GetValue('GA_NUMEROSERIE') = 'X') else GaSer := False;
        end;
        if GaSer then
        begin
          TOBL.PutValue('GL_NUMEROSERIE', 'X');
          if GereReliquat then TOBL.PutValue('GL_QTERELIQUAT', dQteStock); // TOBL.GetValue('GL_QTESTOCK'));
          TOBL.PutValue('GL_QTEFACT', 0);
          TOBL.PutValue('GL_QTESTOCK', 0);
          LotSerieTrouve := True;
          bMajZeroQte := True;
        end;
      end;
      if (TobL.GetValue ('GL_CODECOND') <> '') and (not bMajZeroQte) then
      begin
        TobC := TobConds.FindFirst (['GCO_ARTICLE', 'GCO_CODECOND'],
                                    [TobL.GetValue ('GL_ARTICLE'), TobL.GetValue ('GL_CODECOND')],
                                    False);
        if Tobc <> nil then
        begin
          if TobC.GetValue ('GCO_NBARTICLE') <> 0 then
          begin
            PEntSup := TobL.GetValue ('GL_QTEFACT') / TobC.GetValue ('GCO_NBARTICLE');
            if PEntSup <> Trunc (PEntSup) then
            begin
              if GereReliquat then TOBL.PutValue('GL_QTERELIQUAT', dQteStock);
              TOBL.PutValue('GL_QTEFACT',0) ;
              TOBL.PutValue('GL_QTESTOCK',0) ;
              TOBL.PutValue('GL_QTERESTE',0) ;
            end;
          end;
        end;
      end;
    end;
    if ((DuplicPiece) or (GetInfoParPiece(PieceContainer.NewNature, 'GPP_RECALCULPRIX') = 'X')) then
    begin
      TOBA := FindTOBArtRow(TOBPiece, TOBArticles, i + 1);
      if TOBA <> nil then AffectePrixValo(TOBL, TOBA);
    end;
  end;
  for i := 0 to TOBBases.Detail.Count - 1 do
  begin
    TOBB := TOBBases.Detail[i];
    MajFromCleDoc(TOBB, PieceContainer.Cledoc);
  end;
  for i := 0 to TOBEches.Detail.Count - 1 do
  begin
    TOBB := TOBEches.Detail[i];
    MajFromCleDoc(TOBB, PieceContainer.Cledoc);
  end;
  ReliquatTransfo := ((PieceContainer.TransfoPiece) and (GereReliquat));
  if ((GereSerie) and (TOBSerie_O.Detail.Count > 0) and (ReliquatTransfo)) then InitialiseReliquatSerie (TobSerRel, TobSerie_O);
  if ((PieceContainer.TransfoPiece) and (LotSerieTrouve) and ((MajLotZero) or (MajSerieZero))) then
  begin
    PutValueDetail(TOBPiece, 'GP_RECALCULER', 'X');
    CalculeLaSaisie(-1, -1, True);
  end;
  for i := 0 to TobLigneTarif.Detail.Count - 1 do
  begin
    TobLT := TobLigneTarif.Detail[i];
    TobLT.PutValue('GL_NATUREPIECEG', CleDoc.NaturePiece);
    TobLT.PutValue('GL_SOUCHE', CleDoc.Souche);
    TobLT.PutValue('GL_NUMERO', CleDoc.NumeroPiece);
    TobLT.PutValue('GL_INDICEG', CleDoc.Indice);
    { Boucle sur les lignes }
    for j := 0 to TobLT.Detail.count - 1 do
      MajFromCleDoc(TobLT.Detail[j], CleDoc);
  end;
end;
{$ENDIF}

procedure TFTicketFO.EtudieReliquat;
var TOBT: TOB;
  TypeRel: string;
begin
  GereReliquat := (PieceContainer.Action = taModif) and (GetInfoParPiece(PieceContainer.NewNature, 'GPP_RELIQUAT') = 'X');
  {Exceptions tiers}
  if TOBTiers = nil then Exit;
  TOBT := TOBTiers.FindFirst(['GTP_NATUREPIECEG'], [PieceContainer.NewNature], False);
  if TOBT <> nil then
  begin
    TypeRel := TOBT.GetValue('GTP_RELIQUAT');
    if TypeRel = 'OUI' then GereReliquat := (PieceContainer.Action = taModif) else if TypeRel = 'NON' then GereReliquat := False;
  end;
end;

procedure TFTicketFO.GereVivante;
begin
  if PieceContainer.Action <> taModif then Exit;
  if PieceContainer.TransfoPiece then Exit;
  if PieceContainer.DuplicPiece then Exit;
  if TOBPiece.GetValue('GP_VIVANTE') = '-' then
  begin
    PieceContainer.Action := taConsult;
    HPiece.Execute(9, Caption, '');
  end else
  begin
    BDelete.Visible := True;
  end;
end;

procedure TFTicketFO.InitEuro;
begin
  if TOBPiece.GetValue('GP_SAISIECONTRE') = 'X' then PopD.Items[1].Checked := True else
    if TOBPiece.GetValue('GP_DEVISE') <> V_PGI.DevisePivot then PopD.Items[2].Checked := True else
    PopD.Items[0].Checked := True;
  AfficheEuro;
end;

procedure TFTicketFO.InitRIB;
begin
  if PieceContainer.Action = taCreat then Exit;
  if PieceContainer.DuplicPiece then Exit;
  TOBTiers.PutValue('RIB', TOBPiece.GetValue('GP_RIB'));
end;

function TFTicketFO.FromAvoir: boolean;
begin
  Result := False;
  if TOBPiece = nil then Exit;
  if PieceContainer.Action = taConsult then Exit;
  if not PieceContainer.DuplicPiece then Exit;
  Result := (GetInfoParPiece(TOBPiece.GetValue('GP_NATUREPIECEG'), 'GPP_ESTAVOIR') = 'X');
end;

procedure TFTicketFO.ChargeLaPiece;
var Ind: integer;
{$IFDEF GESCOM}
    TobTmp, TobA : TOB;
{$ENDIF GESCOM}
begin
  PieceContainer.ChargeLaPiece;
  // Opérations de caisse
  FOLoadOperCaisse(TOBPiece, TOBEches, TOBOpCais);  
//  LoadLesTOB;
  {$IFDEF TAXEUS}
  // prise en compte du modèle de taxe de l'établissement
  GetModeleTaxeEtab(TOBPiece.GetValue('GP_ETABLISSEMENT'));
  {$ENDIF}
  if (FromAvoir) or ((not PieceContainer.DuplicPiece) and (EstAvoir)) then InverseAvoir;
  if (LesColonnes = '') or (GP_FACTUREHT.Checked <> (TOBPiece.GetValue('GP_FACTUREHT') = 'X')) then
  begin
    GP_FACTUREHT.Checked := (TOBPiece.GetValue('GP_FACTUREHT') = 'X');
    EtudieColsListe;
  end;
  TOBPiece.PutEcran(Self);
  GP_NUMEROPIECE.Caption := FTitrePiece.Caption + IntToStr(TOBPiece.GetValue('GP_NUMERO'));
  Ind := TOBPiece.GetValue('GP_INDICEG');
  if Ind > 0 then GP_NUMEROPIECE.Caption := FTitrePiece.Caption + GP_NUMEROPIECE.Caption + '  ' + IntToStr(Ind);
  ChargeTiers;
  InitEuro;
  InitRIB;
  LoadLesArticles;
  {$IFDEF STK}
//  if (not PieceContainer.DuplicPiece) or DuplicLeDetailSurAvoir then
  LoadLesTobGQDServi (PieceContainer);
  {$ENDIF STK}
  TOBPiece_O.Dupliquer(TOBPiece, True, True);
  TOBBases_O.Dupliquer(TOBBases, True, True);
  TOBEches_O.Dupliquer(TOBEches, True, True);
  TOBPorcs_O.Dupliquer(TOBPorcs, True, True);
  TOBN_O.Dupliquer(TOBNomenclature, True, True);
  TOBLOT_O.Dupliquer(TOBdesLots, True, True);
  TOBSerie_O.Dupliquer(TOBSerie, True, True);
  TOBAcomptes_O.Dupliquer(TOBAcomptes, True, True);
  TOBOpCais_O.Dupliquer(TOBOpCais, True, True);
  VenteALivrer := (TOBPiece.GetValue('GP_NUMADRESSELIVR') > 0);
  LALIVRER.Visible := VenteALivrer;
{$IFDEF GESCOM}
  TobTmp := TobPiece.FindFirst(['GL_TYPEARTICLE'],['FI'],True);
  while TobTmp <> nil do
  begin
    TobA := TobArticles.FindFirst(['GA_ARTICLE'],[TobTmp.GetValue('GL_ARTICLE')], true);
    if TobA <> nil then
    begin
      TobTmp.PutValue('TYPEARTFI',TobA.GetValue('GA_TYPEARTFINAN'));
      TypePiece := TobTmp.GetValue('GL_TYPEARTICLE');
      TypePieceFi := TobA.GetValue('GA_TYPEARTFINAN');
    end;
    TobTmp := TobPiece.FindNext(['GL_TYPEARTICLE'],['FI'],True);
  end;
{$ENDIF GESCOM}
  TobLigneTarif_O.Dupliquer(TobLigneTarif, True, True);
end;

procedure TFTicketFO.InitEnteteDefaut(Totale: boolean);
var
  Texte, Caisse: string;
  NumZ: integer;
begin
  // on force la date de la piece et la date d'entrée à la date de la dernière ouverture de journée
  Caisse := FOCaisseCourante;
  NumZ := FOGetNumZCaisse(Caisse);
  V_PGI.DateEntree := FOGetDateOuv(Caisse, NumZ);

  PieceContainer.Cledoc.DatePiece := V_PGI.DateEntree;
  GP_DATEPIECE.Text := DateToStr(PieceContainer.Cledoc.DatePiece);
  GP_DATELIVRAISON.Text := DateToStr(PieceContainer.Cledoc.DatePiece);
  GP_ETABLISSEMENT.Value := VH^.EtablisDefaut;
  GP_DEPOT.Value := VH_GC.GCDepotDefaut;
  TrouveDevisePiece;
  GP_NUMEROPIECE.Caption := FTitrePiece.Caption + '';
  PieceContainer.Cledoc.NumeroPiece := 0;
  {$IFDEF GESCOM}
  GP_FACTUREHT.Checked := True;
  {$ELSE}
  GP_FACTUREHT.Checked := False;
  {$ENDIF}
  EtudieColsListe;
  GP_SAISIECONTRE.Checked := False;
  TrouveRegimePiece;
  ChangeTzTiersSaisie(PieceContainer.NewNature);
  TOBPiece.GetEcran(Self);
  TOBPiece.PutValue('GP_NATUREPIECEG', PieceContainer.NewNature);
  TOBPiece.PutValue('GP_VENTEACHAT', VenteAchat);
  {$IFDEF TAXEUS}
  // prise en compte du modèle de taxe de l'établissement
  TOBPiece.PutValue('GP_CODEMODELETAXE', GetModeleTaxeEtab(GP_ETABLISSEMENT.Value));
  {$ENDIF}
  {commercial}
  TypeCom := GetInfoParPiece(PieceContainer.NewNature, 'GPP_TYPECOMMERCIAL');
  if TypeCom = '' then TypeCom := 'VEN';
  {caisse}
  TOBPiece.PutValue('GP_CAISSE', Caisse);
  TOBPiece.PutValue('GP_NUMZCAISSE', NumZ);
  // Remise à blanc des libellés
  LibelleArticle.Caption := '';
  T_PrixU.Caption := '';
  A_QTESTOCK.Caption := '';
  LIBELLETIERS.Caption := '';
  SOLDETIERS.Caption := '';
  TSOLDETIERS.Visible := False;
  LDETAXE.Visible := False;
  LALIVRER.Visible := False;
  // Message par défaut sur l'afficheur
  InvertirAfficheur;
  Texte := Trim(VH_GC.TOBPCaisse.GetValue('GPK_AFFMESG'));
  if Texte = '' then Texte := 'Bienvenue chez ' + GP_ETABLISSEMENT.Text;
  Texte := Texte + ' ';
  AffichageLCD(Texte, 0, 0, 0, '', [ofInterne, ofClient], qaLibreTexte, False);
  if fAffTXT.Visible then
  begin
    if (not TimerLCD.Enabled) and (TimerLCD.Interval <> MaxInterval) then TimerLCD.Interval := MaxInterval;
    TimerLCD.Enabled := True;
  end;
  // 1ère page du pavé
  if PnlBoutons <> nil then PnlBoutons.PageCourante := 0;
end;

procedure TFTicketFO.InitPieceCreation;
begin
  PieceContainer.Cledoc.Souche := GetSoucheG(PieceContainer.Cledoc.NaturePiece, VH^.EtablisDefaut, '');
  PieceContainer.Cledoc.NumeroPiece := GetNumSoucheG(PieceContainer.Cledoc.Souche);
  GP_NUMEROPIECE.Caption := FTitrePiece.Caption + HTitres.Mess[10];
  InitTOBPiece(TOBPiece);
  TOBPiece.PutValue('GP_NUMERO', PieceContainer.Cledoc.NumeroPiece);
  TOBPiece.PutValue('GP_SOUCHE', PieceContainer.Cledoc.Souche);
  GotoEntete;
  //if GP_TIERS.CanFocus then GP_TIERS.SetFocus else GotoEntete ;
  {**if not GoToEntete then
     BEGIN
     if GS.CanFocus then GS.SetFocus ;
     GSEnter(GS) ;
     END ;**}
end;

procedure TFTicketFO.BlocageSaisie(Bloc: boolean);
begin
  if PieceContainer.Action <> taCreat then Exit;
  if ((Bloc) and (TOBArticles.Detail.Count <= 0)) then Exit;
  BloquePiece(taCreat, Bloc);
end;

procedure TFTicketFO.BloquePiece(ActionSaisie: TActionFiche; Bloc: boolean);
begin
  if ActionSaisie = taConsult then
  begin
    PEntete.Enabled := not Bloc;
    GP_ESCOMPTE.Enabled := not Bloc;
    GP_REMISEPIED.Enabled := not Bloc;
    if (not AnnulPiece) and (not AssignePiece) then BValider.Enabled := not Bloc;
    Descriptif.Enabled := not Bloc;
  end else if ActionSaisie = taCreat then
  begin
    GP_DEVISE.Enabled := not Bloc;
    {$IFNDEF FOS5}
    BDEVISE.Enabled := not Bloc;
    {$ENDIF}
  end else
  begin
    if not PieceContainer.DuplicPiece then GP_TIERS.Enabled := not Bloc;
    GP_DEVISE.Enabled := not Bloc;
    {$IFNDEF FOS5}
    BDEVISE.Enabled := not Bloc;
    {$ENDIF}
  end;
end;

procedure TFTicketFO.GotoGrid(PageDown, FromF10: boolean);
var
  Okok : boolean ;
begin
  if Screen.ActiveControl = GS then Exit ;
  if FromF10 then
    Okok := True
  else
    Okok:=((Screen.ActiveControl.Parent=Pentete) and PageDown) or ((Screen.ActiveControl.Parent=PPied) and not PageDown) ;
  if Okok and GS.CanFocus then GS.SetFocus;
end;

function TFTicketFO.GotoEntete: Boolean;
var Ind: integer;
begin
  Result := False;
  if (Screen.ActiveControl = GP_DATEPIECE) or (Screen.ActiveControl = GP_TIERS) or
    (Screen.ActiveControl = GP_ETABLISSEMENT) or (Screen.ActiveControl = GP_REPRESENTANT) then
  begin
    Result := True;
    Exit;
  end;
  for Ind := 0 to 3 do
  begin
    if (GP_DATEPIECE.TabOrder = Ind) and (GP_DATEPIECE.CanFocus) then
    begin
      GP_DATEPIECE.SetFocus;
      Result := True;
      Exit;
    end;
    if (GP_TIERS.TabOrder = Ind) and (GP_TIERS.CanFocus) then
    begin
      GP_TIERS.SetFocus;
      Result := True;
      Exit;
    end;
    if (GP_ETABLISSEMENT.TabOrder = Ind) and (GP_ETABLISSEMENT.CanFocus) then
    begin
      GP_ETABLISSEMENT.SetFocus;
      Result := True;
      Exit;
    end;
    if (GP_REPRESENTANT.TabOrder = Ind) and (GP_REPRESENTANT.CanFocus) then
    begin
      GP_REPRESENTANT.SetFocus;
      Result := True;
      Exit;
    end;
  end;
end;

procedure TFTicketFO.GotoPied;
begin
  if GP_ESCOMPTE.CanFocus then
  begin
    GP_ESCOMPTE.SetFocus;
    Exit;
  end;
  if GP_REMISEPIED.CanFocus then
  begin
    GP_REMISEPIED.SetFocus;
    Exit;
  end;
end;

procedure TFTicketFO.WMGetMinMaxInfo(var MSG: Tmessage);
begin
  with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do
  begin
    X := WMinX;
    Y := WMinY;
  end;
end;

procedure TFTicketFO.InitFormEcheance;
begin
  FormEcheance := nil;
end;

function TFTicketFO.IndiqueOuSaisie(EcheanceOk: boolean = False): TOuSaisie;
begin
  Result := saEntete;
  if FormEcheance <> nil then
  begin
    if FormEcheance.RemiseTicket.Visible then
      Result := saRemiseTicket
    else
      if (EcheanceOk) or (Screen.ActiveControl = FormEcheance.GSReg) then
        Result := saEcheance;
  end else
    if (Screen.ActiveControl = GS) or (Screen.ActiveControl = GS.ValCombo) then
      Result := saLigne;
end;

{==============================================================================================}
{=============================== Evènements de la Form ========================================}
{==============================================================================================}

procedure TFTicketFO.FormCreate(Sender: TObject);
var Err: TMC_Err;
  Dispositif, Port, Params: string;
begin
  InitLesCols;
  VenteALivrer := False;
  DejaSaisie := False;
  RetourGS := False;
  DemandeDetaxe := False;
  SurLigSuivante := False;
  DepuisEcheance := False;
  TMOSimuleClav := FOGetFromRegistry(REGSAISIETIC, REGTMOSIMUCLAV, 0);
  FillChar(IdentCols, Sizeof(IdentCols), #0);
  PieceContainer := TPieceContainer.Create;
  ToutAllouer;
  GeneCharge := False;
  DejaRentre := False;
  ValideEnCours := False;
  CataFourn := False;
  TotEnListe := False;
  GS.GetCellCanvas := GetCellCanvas;
  GS.PostDrawCell := PostDrawCell;
  GS.SquizzInvisibleCells := True;
  RempliPopD;
  WMinX := Width - 100;
  WMinY := Height - 10;
  InMaximise := FALSE;
  InTraiteBouton := FALSE;
  InitFormEcheance;
  // rend la barre de bouton invisible
  {$IFDEF GESCOM}
  DockBottom.Visible := ((V_PGI.SAV) or (FOGetParamCaisse('GPK_TOOLBAR') = 'X') or
    (FOGetParamCaisse('GPK_CLAVIERECRAN') <> 'X'));
  {$ELSE}
  DockBottom.Visible := ((V_PGI.SAV) or (FOGetParamCaisse('GPK_TOOLBAR') = 'X'));
  {$ENDIF}
  if not (IsInside(Self)) and (FOGetParamCaisse('GPK_TOOLBAR') = 'X') then FoShowEtiquette(Self, True);
  // gestion de l'afficheur externe
  Affexterne := nil;
  if FODonneParamAfficheur(Dispositif, Port, Params) then
  begin
    AFFExterne := TAFF.Create(Self, Dispositif);
    if not AffExterne.ChargePortetParams(Port, Params, Err) then FOMCErr(Err, 'Gestion de l''afficheur') else
      if not AffExterne.OuvrirAfficheur(Err) then FOMCErr(Err, 'Gestion de l''afficheur');
  end;
  // vérifie si la saisie du client est obligatoire pour des codes démarques
  DemarqueCliObl := FODemarqueClientOblig('');
end;

procedure TFTicketFO.FormDestroy(Sender: TObject);
begin
  if Assigned(AFFExterne) then
  begin
    AFFExterne.Free;
    AffExterne := nil;
  end;
  // Appel de la fonction de dépilage dans la liste des fiches
  AglDepileFiche;
end;

procedure TFTicketFO.PopEuro(Sender: TObject);
begin
  ChoixEuro(Sender);
end;

procedure TFTicketFO.ChoixEuro(Sender: TObject);
var T: TMenuItem;
begin
  T := TMenuItem(Sender);
  if T.Name = 'POPD_S1' then PopD.Items[0].Checked := True else
    if T.Name = 'POPD_S2' then PopD.Items[1].Checked := True else
    if T.Name = 'POPD_S3' then PopD.Items[2].Checked := True;
  AfficheEuro;
end;

procedure TFTicketFO.RempliPopD;
var S1, S2, S3: string;
begin
  S1 := RechDom('TTDEVISETOUTES', V_PGI.DevisePivot, False);
  if VH^.TenueEuro then S2 := RechDom('TTDEVISETOUTES', V_PGI.DeviseFongible, False)
  else S2 := HTitres.Mess[14];
  S3 := HTitres.Mess[15];
  POPD_S1.Caption := S1;
  POPD_S2.Caption := S2;
  POPD_S3.Caption := S3;
end;

procedure TFTicketFO.ChargeFromNature;
var PassP: string;
begin
  GS.ListeParam := GetInfoParPiece(PieceContainer.NewNature, 'GPP_LISTESAISIE');
  {Caractéristiques nature}
  VenteAchat := GetInfoParPiece(PieceContainer.NewNature, 'GPP_VENTEACHAT');
  GereStatPiece := GetInfoParPiece(PieceContainer.NewNature, 'GPP_AFFPIECETABLE') = 'X';
  {Comptabilité}
  PassP := GetInfoParPiece(PieceContainer.NewNature, 'GPP_TYPEECRCPTA');
  OkCpta := ((PassP <> '') and (PassP <> 'RIE'));
  CompAnalP := GetInfoParPiece(PieceContainer.NewNature, 'GPP_COMPANALPIED');
  CompAnalL := GetInfoParPiece(PieceContainer.NewNature, 'GPP_COMPANALLIGNE');
  OkCptaStock := IsComptaStock(PieceContainer.NewNature);
  {Visas}
  NeedVisa := (GetInfoParPiece(PieceContainer.NewNature, 'GPP_VISA') = 'X');
  MontantVisa := GetInfoParPiece(PieceContainer.NewNature, 'GPP_MONTANTVISA');
  {Commentaires}
  CommentEnt := GetInfoParPiece(PieceContainer.NewNature, 'GPP_COMMENTENT');
  CommentPied := GetInfoParPiece(PieceContainer.NewNature, 'GPP_COMMENTPIED');
  {Ruptures}
  CalcRupt := GetInfoParPiece(PieceContainer.NewNature, 'GPP_CALCRUPTURE');
  ForceRupt := (GetInfoParPiece(PieceContainer.NewNature, 'GPP_FORCERUPTURE') = 'X');
  {Mécanismes et automatismes}
  GereEche := GetInfoParPiece(PieceContainer.NewNature, 'GPP_GEREECHEANCE');
  GereAcompte := (GetInfoParPiece(PieceContainer.NewNature, 'GPP_ACOMPTE') = 'X');
  EstAvoir := (GetInfoParPiece(PieceContainer.NewNature, 'GPP_ESTAVOIR') = 'X');
  OuvreAutoPort := (GetInfoParPiece(PieceContainer.NewNature, 'GPP_OUVREAUTOPORT') = 'X');
  if not (ctxAffaire in V_PGI.PGIContexte) then // PA a voir pb plantage ?
    GereLot := (GetInfoParPiece(PieceContainer.NewNature, 'GPP_LOT') = 'X');
  {$IFDEF CCS3}
  GereSerie := False;
  {$ELSE}
  GereSerie := (GetInfoParPiece(PieceContainer.NewNature, 'GPP_NUMEROSERIE') = 'X');
  {$ENDIF}
  {Divers}
  //DimSaisie:=GetInfoParPiece(PieceContainer.NewNature,'GPP_DIMSAISIE') ;
  DimSaisie := 'TOU';
  {$IFDEF MODE}
  GereAdresse := (GetInfoParPiece(PieceContainer.NewNature, 'GPP_MAJINFOTIERS') = 'X');
  {$ELSE}
  GereAdresse := True;
  {$ENDIF}
  FTitrePiece.Caption := TraduireMemoire(GetInfoParPiece(PieceContainer.NewNature, 'GPP_LIBELLE')) + ' ' + HTitres.Mess[30] + '  ';
  {Commercial}
  TypeCom := GetInfoParPiece(PieceContainer.NewNature, 'GPP_TYPECOMMERCIAL');
  if TypeCom = '' then TypeCom := 'VEN';
  {Tiers}
  ChangeTzTiersSaisie(PieceContainer.NewNature);
end;

procedure TFTicketFO.EnabledGrid;
begin
  BVentil.Visible := ((VPiece.Visible) or (VLigne.Visible));
  BEche.Visible := (((GereEche = 'AUT') or (GereEche = 'DEM')) and (PieceContainer.Action = taConsult));
  BEche.Enabled := BEche.Visible;
  BTicketAttente.Visible := ((FOGetParamCaisse('GPK_GERETICKETATT') = 'X') and (PieceContainer.Action <> taConsult));
  BTicketAttente.Enabled := BTicketAttente.Visible;
  if ((not BEche.Visible) and (not BVentil.Visible)) then Sep2.Visible := False;
  BAcompte.Enabled := ((GereAcompte) and (PieceContainer.Action <> taConsult));
  BInfos.Enabled := True;
  BPorcs.Enabled := True;
  if PieceContainer.Action = taConsult then
  begin
    CommentLigne := False;
  end else
  begin
    TCommentEnt.Visible := (CommentEnt <> '');
    TCommentPied.Visible := (CommentPied <> '');
    BActionsLignes.Enabled := True;
    BSousTotal.Enabled := True;
    CommentLigne := True;
  end;
end;

procedure TFTicketFO.EnabledPied;
begin
  BZoomEcriture.Enabled := (TOBPiece.GetValue('GP_REFCOMPTABLE') <> '');
  BZoomSuivante.Enabled := ((PieceContainer.Action <> taCreat) and (TOBPiece.GetValue('GP_DEVENIRPIECE') <> ''));
  MBTarif.Enabled := ((GetInfoParPiece(PieceContainer.NewNature, 'GPP_CONDITIONTARIF') = 'X') and (PieceContainer.Action <> taConsult));
  MBDatesLivr.Enabled := (PieceContainer.Action <> taConsult);
  GereTiersEnabled;
  GereCommercialEnabled;
  if ((CompAnalP = '') or (CompAnalP = 'SAN')) then
  begin
    VPiece.Visible := False;
    VPiece.Enabled := False;
  end;
  if ((CompAnalL = '') or (CompAnalL = 'SAN')) then
  begin
    VLigne.Visible := False;
    VLigne.Enabled := False;
  end;
  BVentil.Enabled := False;
  BEche.Enabled := False;
  BInfos.Enabled := False;
  BPorcs.Enabled := False;
  if PieceContainer.Action = taConsult then
  begin
    BActionsLignes.Visible := False;
    BSousTotal.Visible := False;
  end else
  begin
    BActionsLignes.Enabled := False;
    BSousTotal.Enabled := False;
  end;
  MCFidelite.Visible := FideliteCli.LoadProgramme(FODonneClientDefaut, GP_ETABLISSEMENT.Value);
  MDALivrer.Visible := ((FOGetParamCaisse('GPK_CLISAISIE') = 'X') and (GereAdresse));

  {$IFDEF GESCOM}
  MCSoldeClient.Visible := (FOGetParamCaisse('GPK_LIAISONREG') = 'X');
  MDStock.Visible := False;
  {$ENDIF}

  {$IFDEF MODE}
  TRechArtImage.Visible := True;
  MCArticleClient.Visible := True;
  MCHistoClient.Visible := True;
  {$ELSE}
  TRechArtImage.Visible := False;
  MCArticleClient.Visible := False;
  MCHistoClient.Visible := False;
  {$ENDIF}

  {$IFDEF TAXEUS}
  BTaxation.Visible := True;
  {$ELSE}
  BTaxation.Visible := False;
  {$ENDIF}
end;

procedure TFTicketFO.FlagRepres;
begin
  if FOGetParamCaisse('GPK_VENDSAISIE') = 'X' then Exit;
  GP_REPRESENTANT.Visible := False;
  GP_REPRESENTANT.TabStop := False;
  HGP_REPRESENTANT.Visible := False;
end;

procedure TFTicketFO.FlagTiers;
begin
  if FOGetParamCaisse('GPK_CLISAISIE') = 'X' then Exit;
  GP_TIERS.Visible := False;
  GP_TIERS.TabStop := False;
  HGP_TIERS.Visible := False;
  LIBELLETIERS.Visible := False;
  TSOLDETIERS.Visible := False;
  SOLDETIERS.Visible := False;
end;

procedure TFTicketFO.FlagPU;
begin
  if FOGetParamCaisse('GPK_AFFPRIXTIC') = '002' then Exit;
  T_PrixU.Visible := False;
  LblPrixU.Visible := False;
end;

procedure TFTicketFO.FlagDepot;
begin
  if VH_GC.GCMultiDepots then Exit;
  GP_DEPOT.Visible := False;
  GP_DEPOT.TabStop := False;
  {$IFNDEF FOS5}
  HGP_DEPOT.Visible := False;
  {$ENDIF}
end;

procedure TFTicketFO.FormShow(Sender: TObject);
var
  ResizeGrid: Boolean;
  PropPv2: RPropPave2;
begin
  if InMaximise then Exit;
  GeneCharge := True;
  PasToucheDepot := False;
  ChangeComplEntete := False;
  LookLesDocks(Self);
  ChargeFromNature;
  ChargeChoixDesPaves(PropPv2);
  FlagRepres;
  FlagDepot;
  FlagTiers;
  FlagPU;
  case PieceContainer.Action of
    taCreat:
      begin
        InitEnteteDefaut(True);
        InitPieceCreation;
{$IFDEF GESCOM}
        TypePiece := ''; //JTR 15/01/2004
        TypePieceFi := '';
{$ENDIF GESCOM}
      end;
    taModif:
      begin
        ChargeLaPiece;
        GereVivante;
        EtudieReliquat;
        BloquePiece(PieceContainer.Action, True);
        if ((not PieceContainer.TransfoPiece) and (not PieceContainer.DuplicPiece) and (PieceContainer.Cledoc.Indice > 0)) then GP_DATEPIECE.Enabled := False;
      end;
    taConsult:
      begin
        ChargeLaPiece;
        BloquePiece(PieceContainer.Action, True);
        if AssignePiece then
        begin
          PEntete.Enabled := True;
          GP_DATEPIECE.Enabled := False;
          GP_REPRESENTANT.Enabled := False;
          GP_TIERS.Enabled := True;
        end;
        BImprimer.Visible := True;
      end;
  end;
  HMTrad.ResizeGridColumns(GS);
  AffecteGrid(GS, PieceContainer.Action);
  if PieceContainer.Action = taConsult then GS.MultiSelect := False;
  InitPasModif;
  EnabledPied;
  EnabledGrid;
  GeneCharge := False;
  ResizeGrid := False;
  // Suppression de la colonne N° de ligne
  if SG_NL <> -1 then
  begin
    GS.ColWidths[SG_NL] := 0;
    GS.ColLengths[SG_NL] := -1;
    ResizeGrid := True;
  end;
  // Suppression de la colonne Vendeur
  if (VH_GC.TOBPCaisse.GetValue('GPK_VENDSAISLIG') = '-') and (SG_Rep <> -1) then
  begin
    GS.ColWidths[SG_Rep] := -1;
    GS.ColLengths[SG_Rep] := -1;
    ResizeGrid := True;
    {$IFDEF GESCOM}
    SG_Rep := -1;
    {$ENDIF}
  end;
  // Suppression de la colonne prix unitaire
  if FOGetParamCaisse('GPK_AFFPRIXTIC') = '001' then
  begin
    if SG_Px <> -1 then
    begin
      GS.ColWidths[SG_Px] := -1;
      GS.ColLengths[SG_Px] := -1;
      ResizeGrid := True;
    end;
    if SG_PxNet <> -1 then
    begin
      GS.ColWidths[SG_PxNet] := -1;
      GS.ColLengths[SG_PxNet] := -1;
      ResizeGrid := True;
    end;
    {$IFDEF GESCOM}
    SG_Px := -1;
    SG_PxNet := -1;
    {$ENDIF}
  end;
  // Suppression des Remises
  if (VH_GC.TOBPCaisse.GetValue('GPK_REMSAISIE') = '-') or (VH_GC.TOBPCaisse.GetValue('GPK_REMAFFICH') = '-') then
  begin
    if SG_Rem <> -1 then
    begin
      GS.ColWidths[SG_Rem] := -1;
      GS.ColLengths[SG_Rem] := -1;
      if VH_GC.TOBPCaisse.GetValue('GPK_REMSAISIE') = '-' then SG_Rem := -1;
    end;
    if SG_RV <> -1 then
    begin
      GS.ColWidths[SG_RV] := -1;
      GS.ColLengths[SG_RV] := -1;
      if VH_GC.TOBPCaisse.GetValue('GPK_REMSAISIE') = '-' then SG_RV := -1;
    end;
    if SG_RL <> -1 then
    begin
      GS.ColWidths[SG_RL] := -1;
      GS.ColLengths[SG_RL] := -1;
      if VH_GC.TOBPCaisse.GetValue('GPK_REMSAISIE') = '-' then SG_RL := -1;
    end;
    ResizeGrid := True;
    {$IFDEF GESCOM}
    SG_Rem := -1;
    SG_RV := -1;
    SG_RL := -1;
    {$ENDIF}
  end;
  // Suppression de la colonne Démarque
  if (not FOGereDemarque(False)) and (SG_TypeRem <> -1) then
  begin
    GS.ColWidths[SG_TypeRem] := -1;
    GS.ColLengths[SG_TypeRem] := -1;
    SG_TypeRem := -1;
    ResizeGrid := True;
  end;
  if ResizeGrid then HMTrad.ResizeGridColumns(GS);
  // Ouverture du tiroir caisse
  MDOuvreTiroir.Visible := (FOExisteTiroir);
  // Détaxe
  MDDetaxe.Visible := DetaxeActive;
  // Ecran tactile
  if FOExisteClavierEcran then
  begin
    PImage.Visible := False;
    FOCreateClavierEcran(PnlBoutons, Self, Pmain);
    PnlBoutons.LanceBouton := SaisieClavierEcran;
    PnlBoutons.LanceCalculette := SaisieCalculette;
    PnlBoutons.BoutonCalculette := BoutonCalculetteClick;
    if PropPv2.Visible then
    begin
      PClavier2.Visible := True;
      PClavier2.Width := PropPv2.Largeur;
      if PropPv2.Align = taLeftJustify then
        PClavier2.Align := alLeft
      else
        PClavier2.Align := alRight;
      //HMTrad.ResizeGridColumns(GS);
      PnlBtn2 := TClavierEcran.Create(Self, 1);
      PnlBtn2.Parent := PClavier2;
      PnlBtn2.Align := alBottom;
      PnlBtn2.Height := PClavier2.Tag;
      PnlBtn2.colornb := 6;
      PnlBtn2.LanceBouton := SaisieClavierEcran;
      PnlBtn2.LanceCalculette := SaisieCalculette;
      PnlBtn2.BoutonCalculette := BoutonCalculetteClick;
      PnlBtn2.Caisse := FOAlphaCodeNumeric(FOCaisseCourante);
      PnlBtn2.NbrBtnWidth := PropPv2.NbrBtnWidth;
      PnlBtn2.NbrBtnHeight := PropPv2.NbrBtnHeight;
      PnlBtn2.ClcPosition := PropPv2.ClcPosition;
      PnlBtn2.ClcVisible := PropPv2.ClcVisible;
      PnlBoutons.OnChangeCalculette := OnChangeCalculette2;
      PnlBtn2.OnChangeCalculette := OnChangeCalculette;
    end;
  end else
  begin
    PImage.Visible := True;
    PImage.Align := alClient;
  end;
  // Afficheur client
  if FOGetParamCaisse('GPK_AFFICHEUR') <> 'X' then LCDInVisible;
  // Photo de l'article
  if (FOAfficheImageArt) or (FOLogoCaisse <> '') then
  begin
    ImageArticle.Visible := True;
    FOAffichePhotoArticle('', ImageArticle);
  end else ImageArticle.Visible := False;
  // Appel de la fonction d'empilage dans la liste des fiches
  AglEmpileFiche(Self);
  // Maximisation de la forme
  if not TimerGen.Enabled then
  begin
    TimerGen.Interval := ToMaximise;
    TimerGen.Enabled := True;
  end;
end;

procedure TFTicketFO.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var OkG, Vide: Boolean;
begin
  if Assigned(PnlBoutons) then PnlBoutons.GereKey(Key, Shift);
  if IndiqueOuSaisie = saEcheance then
  begin
    FormEcheance.FormKeyDown(Sender, Key, Shift);
    Exit;
  end;
  OkG := (Screen.ActiveControl = GS);
  Vide := (Shift = []);

  if key = VK_RECHERCHE then
  begin
    if ((OkG) and (Vide)) then
    begin
      Key := 0;
      GSElipsisClick(nil);
    end;
  end else
    if key = VK_VALIDE then
  begin
    if Vide then
    begin
      Key := 0;
      ClickValideTicket;
    end;
  end else
    case Key of
      VK_RETURN: if OkG then Key := VK_TAB else
        begin
          if Screen.ActiveControl is TMaskEdit then TMaskEdit(Screen.ActiveControl).SelLength := 0;
          Key := 0;
          NextControl(Self, False);
        end;
      VK_INSERT: if ((OkG) and (Vide)) then
        begin
          Key := 0;
          ClickInsert(GS.Row);
        end;
      VK_DELETE: if ((OkG) and (Shift = [ssCtrl])) then
        begin
          Key := 0;
          ClickDel(GS.Row, True, True);
        end;
      VK_HOME: if Shift = [ssCtrl] then
        begin
          Key := 0;
          GotoEntete;
        end;
      VK_END: if Shift = [ssCtrl] then
        begin
          Key := 0;
          GotoPied;
        end;
      VK_NEXT : if Shift = [ssCtrl] then
        begin
          Key := 0;
          GotoGrid(True, False) ;
        end;
      VK_PRIOR : if Shift = [ssCtrl] then
        begin
          Key := 0;
          GotoGrid(False, False) ;
        end;
      {Ctrl+D} 68: if Shift = [ssCtrl] then
        begin
          Key := 0;
          AppelleDim(GS.Row);
        end;
      {Alt+E} 69: if Shift = [ssAlt] then
        begin
          Key := 0;
          BEcheClick(nil);
        end;
      VK_ESCAPE: if Vide then
        begin
          Key := 0;
          DepuisEcheance := False;
          BAbandonClick(nil);
        end;
    end;
end;

procedure TFTicketFO.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
var Abandon: Boolean;
    NoErr: integer;
    PieceContainerOrig : TPieceContainer;
begin
  if DepuisEcheance then
  begin
    CanClose := False;
    DepuisEcheance := False;
    Exit;
  end;
  if PieceContainer.Action = taConsult then Exit;
  if ValideEnCours then
  begin
    CanClose := False;
    Exit;
  end;
  if ((not ForcerFerme) and (PieceModifiee) and (DejaSaisie)) then
  begin
{$IFDEF CHR}
      Abandon := True ;
      ReInitPiece ;
      if Abandon then FOIncrJoursCaisse(jfoNbTicAbandon) ;
{$ELSE}
    if not FOJaiLeDroit(24) then
    begin
      CanClose := False;
      Exit;
    end;
    if HPiece.Execute(6, Caption, '') = mrYes then
    begin
      Abandon := True;
      if FOGetParamCaisse('GPK_GERETICKETATT') = 'X' then
      begin
        PieceContainerOrig := TPieceContainer.Create;
        PieceContainerOrig.CreateTobs;
        PieceContainerOrig.Dupliquer(PieceContainer);
        FODepileTOBGrid(GS, TOBPiece, [SG_RefArt, SG_Lib]);
        NoErr := FOMisePieceEnAttente(PieceContainer);
        if NoErr = 0 then
          Abandon := False
        else if NoErr = 2 then
        begin
          PieceContainer.Dupliquer(PieceContainerOrig);
          PieceContainerOrig.FreeTobs;
          PieceContainerOrig.Free;
          CanClose := False;
          Exit;
        end;
        PieceContainerOrig.FreeTobs;
        PieceContainerOrig.Free;
      end;
      ReInitPiece;
      if Abandon then FOIncrJoursCaisse(jfoNbTicAbandon);
    end;
{$ENDIF} // FIN CHR
    CanClose := (not ReInitUtilisateur);
    //CanClose := False;
  end;
end;

procedure TFTicketFO.ToutAllouer;
begin
  GS.RowCount := NbRowsInit;
  StCellCur := '';
  // Pièce
  TOBPiece := TOB.Create('PIECE', nil, -1);
  TOBPiece_O := TOB.Create('', nil, -1);
  AddLesSupEntete(TOBPiece);
  AddLesSupEntete(TOBPiece_O);
  TOBBases := TOB.Create('BASES', nil, -1);
  TOBBases_O := TOB.Create('', nil, -1);
  TOBEches := TOB.Create('Les ECHEANCES', nil, -1);
  TOBEches_O := TOB.Create('', nil, -1);
  TOBPorcs := TOB.Create('PORCS', nil, -1);
  TOBPorcs_O := TOB.Create('', nil, -1);
  // Fiches
  TOBTiers := TOB.Create('TIERS', nil, -1);
  TOBTiers.AddChampSup('RIB', False);
  TOBArticles := TOB.Create('ARTICLES', nil, -1);
  TOBConds := TOB.Create('CONDS', nil, -1);
  TOBTarif := TOB.Create('TARIF', nil, -1);
  TOBComms := TOB.Create('COMMERCIAUX', nil, -1);
  TOBCatalogu := TOB.Create('LECATALOGUE', nil, -1);
  // Adresses
  TOBAdresses := CreerTobAdresses;
  // Divers
  TOBCXV := TOB.Create('', nil, -1);
  TOBNomenclature := TOB.Create('NOMENCLATURES', nil, -1);
  TOBN_O := TOB.Create('', nil, -1);
  TOBDim := TOB.Create('', nil, -1);
  TOBDesLots := TOB.Create('', nil, -1);
  TOBLOT_O := TOB.Create('', nil, -1);
  TOBSerie := TOB.Create('', nil, -1);
  TOBSerie_O := TOB.Create('', nil, -1);
  TOBSerRel := TOB.Create('', nil, -1);
  TOBAcomptes := TOB.Create('', nil, -1);
  TOBAcomptes_O := TOB.Create('', nil, -1);
  TOBDispoContreM := TOB.Create('', nil, -1);
  // Affaires
  TOBAffaire := TOB.Create('AFFAIRE', nil, -1);
  // Comptabilité
  TOBCPTA := TOB.Create('', nil, -1);
  TOBANAP := TOB.Create('', nil, -1);
  TOBANAS := TOB.Create('', nil, -1);
  // retenues de garantie
  TOBPieceRG := TOB.create('', nil, -1);
  // Bases de tva sur RG
  TOBBasesRG := TOB.create('', nil, -1);
  // Opérations de caisse
  TOBOpCais := TOB.Create('', nil, -1);
  TOBOpCais_O := TOB.Create('', nil, -1);
  // Fidelite JD
  FideliteCli := Fidelite.Create;
  TobLigneTarif := Tob.Create('_ENTETE_', nil, -1);
  TobLigneTarif_O := Tob.Create('', nil, -1);
  PutLesSupEnteteLigneTarif(TobPiece, TobLigneTarif, TobLigneTarif_O);
  PutTobPieceInTobLigneTarif(TobPiece, TobLigneTarif);
  TOBArtFiTiers := TOB.Create('',nil,-1);
  TOBAFormule := TOB.Create('Article Formule Qte', nil,-1);
  TOBLigFormule := TOB.Create('Formule ligne', nil,-1);
  TobAConsignes :=TOB.Create('Article Consignes', nil, -1);
  { Création & Mise à jour du container de tob }
  PieceContainer.CreateTobs;
  PieceContainer.InitTobs(Self);
  MultiSel_SilentMode := False; // eQualité 11973
end;

procedure TFTicketFO.ToutLiberer;
begin
  GS.SynEnabled := False;
  GS.Row := GS.FixedRows;
  GS.Col := GS.FixedCols;
  GS.SynEnabled := True;
  GS.VidePile(False);
  PurgePop(POPZ);
  PieceContainer.FreeTobs;
  if Assigned(TOBEches_O) then FreeAndNil(TOBEches_O);
  if Assigned(TOBPIECERG) then FreeAndNil(TOBPIECERG);
  if Assigned(TOBBASESRG) then FreeAndNil(TOBBASESRG);
  if Assigned(TOBOpCais) then FreeAndNil(TOBOpCais);
  if Assigned(TOBOpCais_O) then FreeAndNil(TOBOpCais_O);
  if Assigned(FideliteCli) then FreeAndNil(FideliteCli);
  {$IFDEF CHR}
  if Assigned(Tobregrpe) then FreeAndNil(Tobregrpe);
  {$ENDIF} // FIN CHR
  if Assigned(TobLigneTarif_O) then FreeAndNil(TobLigneTarif_O);
  if Assigned(TOBArtFiTiers) then FreeAndNil(TOBArtFiTiers); // JTR Lien Opération Caisse avec Tiers
  if TobTicketAttente <> nil then FreeAndNil(TobTicketAttente);  
end;

procedure TFTicketFO.FormClose(Sender: TObject; var Action: TCloseAction);
var Err: TMC_Err;
begin
  if TimerLCD.Enabled then TimerLCD.Enabled := False;
  Application.ProcessMessages;
  if Assigned(AFFExterne) then
  begin
    if AffExterne.isConnected then
    begin
      Affexterne.initialise(Err);
      if not AffExterne.FermeAfficheur(Err) then FOMCErr(Err, 'Gestion de l''afficheur');
    end;
  end;
  {$IFNDEF FOS5}
  PPInfosLigne.Clear;
  PPInfosLigne.Free;
  PPInfosLigne := nil;
  {$ENDIF}
  ToutLiberer;
  PieceContainer.Free;
  if IsInside(Self) then Action := caFree;
end;

{==============================================================================================}
{=============================== Actions liées au Grid ========================================}
{==============================================================================================}

procedure TFTicketFO.EtudieColsListe;
var i: integer;
  Nam, St: string;
begin
  LesColonnes := GS.Titres[0];
  // type de prix unitaire affiché (brut ou net)
  if FOGetParamCaisse('GPK_AFFPRIXTIC') = '002' then // prix net
  begin
    LesColonnes := FindEtReplace(LesColonnes, 'GL_PUHTDEV', 'GL_PUHTNETDEV', False);
    GS.Titres[0] := LesColonnes;
  end;
  if FOGetParamCaisse('GPK_AFFPRIXTIC') = '003' then // prix brut
  begin
    LesColonnes := FindEtReplace(LesColonnes, 'GL_PUHTNETDEV', 'GL_PUHTDEV', False);
    GS.Titres[0] := LesColonnes;
  end;
  // prix et montant HT ou TTC
  if not GP_FACTUREHT.Checked then
  begin
    LesColonnes := FindEtReplace(LesColonnes, 'GL_PUHTDEV', 'GL_PUTTCDEV', False);
    LesColonnes := FindEtReplace(LesColonnes, 'GL_PUHTNETDEV', 'GL_PUTTCNETDEV', False);
    LesColonnes := FindEtReplace(LesColonnes, 'GL_TOTALHTDEV', 'GL_TOTALTTCDEV', False);
    LesColonnes := FindEtReplace(LesColonnes, 'GL_MONTANTHTDEV', 'GL_MONTANTTTCDEV', False);
  end;
  EtudieColsGrid(GS, True);
  St := LesColonnes;
  for i := 0 to GS.ColCount - 1 do
  begin
    Nam := ReadTokenSt(St);
    IdentCols[i].ColName := Nam;
    IdentCols[i].ColTyp := ChampToType(Nam);
    if IdentCols[i].ColTyp = 'DATE' then
    begin
      GS.ColTypes[i] := 'D';
      GS.ColFormats[i] := ShortdateFormat;
    end;
    //   if Nam='GL_TYPEREMISE' then GS.ColFormats[i]:='CB=GCTYPEREMISE' ;
    if Nam = 'GL_TOTALHTDEV' then TotEnListe := True;
  end;
end;

function TFTicketFO.ZoneAccessible(ACol, ARow: Longint): boolean;
var TOBA, TOBL: TOB;
  CodeA, TypeDim, PiecePrec: string;
  IndiceLot, IndiceSerie: integer;
  RemA, VerifColLen: boolean;
begin
  Result := True;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  TOBA := FindTOBArtRow(TOBPiece, TOBArticles, ARow);
  if TOBA <> nil then CodeA := TOBA.GetValue('GA_ARTICLE') else CodeA := '';
  {$IFDEF CCS3}
  IndiceLot := 0;
  IndiceSerie := 0;
  {$ELSE}
  IndiceLot := TOBL.GetValue('GL_INDICELOT');
  IndiceSerie := TOBL.GetValue('GL_INDICESERIE');
  {$ENDIF}
  TypeDim := TOBL.GetValue('GL_TYPEDIM');
  PiecePrec := TOBL.GetValue('GL_PIECEPRECEDENTE');
  {$IFDEF CHR}
  if (SG_Montant > 0) and (Acol = SG_Montant) Then
  begin
    result := False;
    exit;
  end;
  if (SG_LibRegrpe > 0) and (Acol = SG_LibRegrpe) Then
  begin
    result := False;
    exit;
  end;
  if (SG_DateProd > 0) and (Acol = SG_DateProd) Then
  begin
    result := False;
    exit;
  end;
  {$ENDIF} // FIN CHR
  // Si la longueur de la colonne est inférieure à 0 la colonne n'est pas saisissable
  // Dans le cas où on gère les remise mais on ne veut pas les afficher, on considére la colonne % de remise comme disponible.
  VerifColLen := True;
  if (ACol = SG_Rem) and (FOGetParamCaisse('GPK_REMSAISIE') = 'X') and (FOGetParamCaisse('GPK_REMAFFICH') = '-') then VerifColLen := False;
  if (VerifColLen) and (GS.ColLengths[ACol] < 0) then
  begin
    Result := False;
    Exit;
  end;
  // Modification de la référence saisie
  if (ACol = SG_RefArt) and ((CodeA <> '') or (TOBL.GetValue('GL_LIBELLE') <> '')) then
  begin
    // Réglement lié
    if FOChampSupValeurNonVide(TOBL, 'GOC_NUMPIECELIEN') then
    begin
      Result := False;
      Exit;
    end;
    if not FOJaiLeDroit(26, False, False) then
    begin
      Result := False;
      Exit;
    end;
  end;
  // Modification du libellé
  if (ACol = SG_Lib) and ((CodeA <> '') or (TOBL.GetValue('GL_LIBELLE') <> '')) then
  begin
    if not FOJaiLeDroit(28, False, False) then
    begin
      Result := False;
      Exit;
    end;
  end;

  if TypeDim = 'GEN' then
  begin
    if DimSaisie = 'DIM' then
    begin
      Result := False;
    end
    else
    begin
      Result := (ACol = SG_Lib); {if Not Result then ACol:=SG_Lib ;}
    end;
    Exit;
  end;

  if TypeDim = 'DIM' then
  begin
    if DimSaisie = 'GEN' then
    begin
      Result := False;
      Exit;
    end;
  end;

  if ((GereLot) and (IndiceLot > 0) and (ACol = SG_Dep)) then
  begin
    Result := False;
    Exit;
  end;
  if ((GereSerie) and (IndiceSerie > 0) and (ACol = SG_Dep)) then
  begin
    Result := False;
    Exit;
  end;
  if ((PasToucheDepot) and (ACol = SG_Dep)) then
  begin
    Result := False;
    Exit;
  end;

  if CodeA = '' then
  begin
    Result := ((ACol = SG_RefArt) or (ACol = SG_Lib)); {if Not Result then ACol:=GS.FixedCols ;}
    if (TOBL.GetValue('GL_TYPELIGNE') = 'TOT') and (VH_GC.TOBPCaisse.GetValue('GPK_REMSAISIE') = 'X') then
    begin
      if (ACol = SG_Rem) or (ACol = SG_RL) or (ACol = SG_Montant) then Result := True;
      if (ACol = SG_TypeRem) and (FOGereDemarque(False)) then Result := (TOBL.GetValue('GL_TOTREMLIGNEDEV') <> 0);
    end;
  end else
  begin
    // Remises
 //   RemA:=((TOBA.GetValue('GA_REMISELIGNE')='X') and (VH_GC.TOBPCaisse.GetValue('GPK_REMSAISIE')='X')) ;
    RemA := ((TOBA.GetValue('GA_REMISELIGNE') = 'X') and (VH_GC.TOBPCaisse.GetValue('GPK_REMSAISIE') = 'X') and
      (TOBL.GetValue('GL_REMISABLELIGNE') = 'X'));
    if not RemA then
    begin
      if ACol = SG_Rem then Result := False;
      //      if ((IdentCols[ACol].ColName='GL_TOTREMLIGNEDEV') or (IdentCols[ACol].ColName='GL_MONTANTHTDEV') or
      //          (IdentCols[ACol].ColName='GL_MONTANTTTCDEV')) then Result:=False ;
      if ACol = SG_Montant then
      begin
        if (FOGetParamCaisse('GPK_AFFPRIXTIC') = '001') and
          (TOBL.FieldExists('MODIFPU')) then
          Result := (TOBL.GetValue('MODIFPU') = 'X')
        else Result := False;
      end;
      if IdentCols[ACol].ColName = 'GL_TOTREMLIGNEDEV' then Result := False;
      if IdentCols[ACol].ColName = 'GL_TYPEREMISE' then if TOBL.GetValue('GL_TOTREMLIGNEDEV') = 0 then Result := False;
      if not Result then Exit;
    end;
    if TOBL.FieldExists('MODIFPU') then RemA := (TOBL.GetValue('MODIFPU') = 'X')
    else RemA := (TOBA.GetValue('GA_REMISELIGNE') = 'X');

    if not RemA then
    begin
      //      if (IdentCols[ACol].ColName='GL_PUHTNETDEV') or (IdentCols[ACol].ColName='GL_PUTTCNETDEV') then Result:=False ;
      //{$IFDEF GESCOM} - NA 02/12/03 fiche 11131
      //if (ACol = SG_PxNet) then Result := False;
      //{$ELSE}
      if (ACol = SG_Px) or (ACol = SG_PxNet) then Result := False;
      //{$ENDIF}
      if not Result then Exit;
    end;
    //   if IdentCols[ACol].ColName='GL_TYPEREMISE' then if TOBL.GetValue('GL_TOTREMLIGNEDEV')=0 then Result:=False ;
    if ACol = SG_TypeRem then if TOBL.GetValue('GL_TOTREMLIGNEDEV') = 0 then Result := False;

    // On ne peut pas changer la quantité sur une pièce transformée qui gère les reliquats
    if ((not PieceContainer.TransfoPiece) and (not PieceContainer.DuplicPiece) and (PieceContainer.Action = taModif) and
      (GereReliquat) and ((ACol = SG_QF) or (ACol = SG_QS))) then
    begin
      if PiecePrec <> '' then
      begin
        Result := False;
        Exit;
      end;
    end;

    // Modification de la quantité
    if ((ACol = SG_QF) or (ACol = SG_QS)) then
    begin
      // On ne peut pas changer la quantité en transformation sur une ligne qui avait des lots
      if ((PieceContainer.TransfoPiece) and (not PieceContainer.DuplicPiece) and (PieceContainer.Action = taModif) and (GereLot)) then
      begin
        if ((PiecePrec <> '') and (IndiceLot > 0) and (TOBPiece_O.GetValue('GP_ARTICLESLOT') = 'X')) then
        begin
          Result := False;
          Exit;
        end;
      end;
      // Réglement lié
      if FOChampSupValeurNonVide(TOBL, 'GOC_NUMPIECELIEN') then
      begin
        Result := False;
        Exit;
      end;
      if not FOJaiLeDroit(27, False, False) then
      begin
        Result := False;
        Exit;
      end;
    end;

    if (ACol = SG_TypeRem) and ((not FOGereDemarque(False)) or
      ((not FOJaiLeDroit(29, False, False)) and (TOBL.GetValue('GL_TYPEREMISE') <> ''))) then
    begin
      Result := False;
      Exit;
    end;
    if (ACol = SG_Rep) and ((VH_GC.TOBPCaisse.GetValue('GPK_VENDSAISLIG') = '-') or
      ((not FOJaiLeDroit(36, False, False)) and (TOBL.GetValue('GL_REPRESENTANT') <> ''))) then
//    begin
      Result := False;
//      Exit;
//    end;
  end;
end;

procedure TFTicketFO.ZoneSuivanteOuOk(var ACol, ARow: Longint; var Cancel: boolean);
var Sens, ii: integer;
  OldEna, ChgLig: boolean;
begin
  OldEna := GS.SynEnabled;
  GS.SynEnabled := False;
  Sens := -1;
  ChgLig := (GS.Row <> ARow);
  if GS.Row > ARow then Sens := 1 else if ((GS.Row = ARow) and (ACol < GS.Col)) then Sens := 1;
  ACol := GS.Col;
  ARow := GS.Row;
  ii := 0;
  while not ZoneAccessible(ACol, ARow) do
  begin
    Cancel := True;
    inc(ii);
    if ii > 500 then Break;
    //if Sens=1 then
    if (Sens = 1) or ((ACol = GS.FixedCols) and (ARow = 1)) then
    begin
      Sens := 1;
      if ((ACol = GS.ColCount - 1) and (ARow = GS.RowCount - 1)) then Break;
      if ChgLig then
      begin
        ACol := GS.FixedCols - 1;
        ChgLig := False;
      end;
      if ACol < GS.ColCount - 1 then Inc(ACol) else
      begin
        Inc(ARow);
        ACol := GS.FixedCols;
      end;
    end else
    begin
      if ((ACol = GS.FixedCols) and (ARow = 1)) then Break;
      //if ChgLig then BEGIN ACol:=GS.ColCount ; ChgLig:=False ; END ;
      if ChgLig then
      begin
        if ARow <= GS.FixedRows then ACol := GS.ColCount else ACol := GS.FixedCols - 1;
        ChgLig := False;
      end;
      if ACol > GS.FixedCols then Dec(ACol) else
      begin
        Dec(ARow);
        ACol := GS.ColCount - 1;
      end;
    end;
  end;
  GS.SynEnabled := OldEna;
  if PieceContainer.Action <> taConsult then CreerTOBLignes(GS, TOBPiece, TOBTiers, TOBAffaire, ARow);
end;

function TFTicketFO.FormateZoneDivers(St: string; ACol: Longint): string;
var ColName, Typ, StC: string;
begin
  Result := St;
  StC := St;
  ColName := IdentCols[ACol].ColName;
  if ColName = '' then Exit;
  Typ := IdentCols[ACol].ColTyp;
  if (Typ = 'INTEGER') or (Typ = 'SMALLINT') then StC := InttoStr(ValeurI(St)) else
    if (Typ = 'DOUBLE') or (Typ = 'RATE') then StC := StrF00(Valeur(St), V_PGI.OkDecV) else
    if Typ = 'DATE' then
  begin
    if St <> '' then if not IsValidDate(St) then StC := GP_DATEPIECE.Text;
  end else
    if Typ = 'BOOLEAN' then
  begin
    if Uppercase(St) <> 'X' then St := '-';
    StC := Uppercase(St);
  end else
    if ((Typ = 'COMBO') and (Copy(GS.ColFormats[ACol], 1, 3) <> 'CB=')) then StC := Uppercase(St)
  else StC := St;
  Result := StC;
end;

procedure TFTicketFO.FormateZoneSaisie(ACol, ARow: Longint);
var St, StC: string;
begin
  St := GS.Cells[ACol, ARow];
  StC := St;
  if ((ACol = SG_RefArt) or (ACol = SG_Aff) or (ACol = SG_Rep) or (ACol = SG_Dep)) then StC := uppercase(Trim(St)) else
    if ACol = SG_Px then StC := StrF00(Valeur(St), PieceContainer.Dev.Decimale) else
    {$IFNDEF FOS5}
    if ACol = SG_Rem then StC := StrF00(Valeur(St), ADecimP) else
    {$ENDIF}
    if ((ACol = SG_QF) or (ACol = SG_QS)) then StC := StrF00(Valeur(St), V_PGI.OkDecQ) else
    StC := FormateZoneDivers(St, ACol);
  GS.Cells[ACol, ARow] := StC;
end;

{==============================================================================================}
{=============================== Evènements de le Grid ========================================}
{==============================================================================================}

procedure TFTicketFO.PostDrawCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);

  procedure TireLigne(DebX, DebY, FinX, FinY: integer);
  begin
    GS.Canvas.MoveTo(DebX, DebY);
    GS.Canvas.LineTo(FinX, FinY);
  end;

var ARect: TRect;
  Fix: boolean;
begin
  if GS.RowHeights[ARow] <= 0 then Exit;
  if ARow > GS.TopRow + GS.VisibleRowCount - 1 then Exit;
  ARect := GS.CellRect(ACol, ARow);
  Fix := ((ACol < GS.FixedCols) or (ARow < GS.FixedRows));
  GS.Canvas.Pen.Style := psSolid;
  GS.Canvas.Pen.Color := clgray;
  GS.Canvas.Brush.Style := bsClear;
  TireLigne(ARect.Left, ARect.Bottom - 1, ARect.Right, ARect.Bottom - 1);
  if Fix then
  begin
    GS.Canvas.Brush.Style := bsSolid;
    GS.Canvas.Pen.Color := clBlack;
    TireLigne(ARect.Left, ARect.Bottom - 1, ARect.Right, ARect.Bottom - 1);
    GS.Canvas.Pen.Color := clWhite;
    GS.Canvas.Brush.Style := bsClear;
    if ARow < GS.FixedRows then TireLigne(ARect.Left, ARect.Top, ARect.Right, ARect.Top);
    if ACol < GS.FixedCols then TireLigne(ARect.Left + 1, ARect.Top, ARect.Right, ARect.Top);
  end;
end;

procedure TFTicketFO.GetCellCanvas(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
var TOBL: TOB;
begin
  if ACol < GS.FixedCols then Exit;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  {Ligne non imprimable}
  if TOBL.GetValue('GL_NONIMPRIMABLE') = 'X' then if ((PieceContainer.Action <> taConsult) or (ARow <> GS.Row)) then Canvas.Font.Color := clBlue;
  {Lignes Multi-Dim}
  if TOBL.GetValue('GL_TYPEDIM') = 'DIM' then AfficheDetailDim(Canvas, DimSaisie, ACol) else
    if TOBL.GetValue('GL_TYPEDIM') = 'GEN' then AfficheGenDim(Canvas, DimSaisie);
  {Lignes de sous-total}
  if TOBL.GetValue('GL_TYPELIGNE') = 'TOT' then Canvas.Font.Style := Canvas.Font.Style + [fsBold, fsItalic];
  {Reliquats}
  if ((PieceContainer.Action <> taCreat) and (GereReliquat)) then
  begin
    if IdentCols[ACol].ColName = 'GL_QTERELIQUAT' then
    begin
      if TOBL.GetValue('GL_SOLDERELIQUAT') = 'X' then Canvas.Font.Style := Canvas.Font.Style + [fsStrikeOut];
    end;
  end;
end;

procedure TFTicketFO.GSMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  GX := X;
  GY := Y;
end;

procedure TFTicketFO.GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  // Pour éviter de créer des lignes vides
  if Ou >= GS.RowCount - 1 then GS.RowCount := GS.RowCount + NbRowsPlus;
  if PieceContainer.Action <> taConsult then CreerTOBLignes(GS, TOBPiece, TOBTiers, TOBAffaire, Ou);
  GereEnabled(Ou);
  ShowDetail(Ou);
  GereDescriptif(Ou, True);
end;

procedure TFTicketFO.GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var Ok: boolean;
begin
  if csDestroying in ComponentState then Exit;
  {$IFDEF GESCOM}
  if not (FOExisteClavierEcran) then GS.Refresh;
  {$ENDIF}
  Ok := True;
  if Ok then GS.Refresh;
  if PieceContainer.Action = taConsult then Exit;
  TOBLigneExiste(ou);
  {$IFDEF CHR}
  GereChampsChr(ou, cancel);
  {$ENDIF} // FIN CHR
  DeflagTarif(Ou);
{$IFDEF V500}
  CommVersLigne(TOBPiece, TOBArticles, TOBComms, Ou, True);
{$ELSE}
  CommVersLigne(PieceContainer, Ou, True);
{$ENDIF}
  DepileTOBLignes(GS, TOBPiece, Ou, GS.Row);
  GereDescriptif(Ou, False);
  GereAnal(Ou);
{$IFDEF STK}
  GereStock (ou, false, false, false);
{$ELSE STK}
  GereLesLots(Ou);
  GereLesSeries(Ou);
{$ENDIF STK}
  GereArtsLies(Ou);
  if PieceContainer.Action = taCreat then BlocageSaisie(True);
  if TesteMargeMini(Ou) then Cancel := True;
end;

procedure TFTicketFO.GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var TOBL: TOB;
  OldCol, OldRow: Integer;
begin
  if PieceContainer.Action = taConsult then Exit;
  OldCol := ACol;
  OldRow := ARow;
  ZoneSuivanteOuOk(ACol, ARow, Cancel);
  if not Cancel then
  begin
    TOBLigneExiste(ARow);
    //GS.ElipsisButton := (GS.Col = SG_RefArt) or (GS.Col = SG_Aff) or (GS.Col = SG_Rep) or (GS.Col = SG_Dep) or
      //(GS.Col = SG_TypeRem);
// MODIF CHR
    GS.ElipsisButton:=(GS.Col=SG_RefArt) or (GS.Col=SG_Aff) or (GS.Col=SG_Rep) or (GS.Col=SG_Dep) or
      (GS.Col=SG_TypeRem) or (GS.Col=SG_Regrpe) ;
    if ((CommentLigne) and (GS.Col = SG_Lib)) then GS.ElipsisButton := true;
    if GS.Col = SG_RefArt then GS.ElipsisHint := HTitres.Mess[0]
    else if GS.Col = SG_Aff then GS.ElipsisHint := HTitres.Mess[3]
    else if GS.Col = SG_Rep then GS.ElipsisHint := HTitres.Mess[31]
    else if GS.Col = SG_TypeRem then GS.ElipsisHint := HTitres.Mess[32]
    else if GS.Col = SG_Dep then GS.ElipsisHint := HTitres.Mess[12]
    else if GS.Col = SG_Lib then GS.ElipsisHint := HTitres.Mess[16]
    else if ((GS.Col = SG_QF) or (GS.Col = SG_QS)) then
    begin
      TOBL := GetTOBLigne(TOBPiece, ARow);
      if TOBL <> nil then
      begin
        if TobConds.FindFirst (['GCO_ARTICLE'], [TobL.GetValue ('GL_ARTICLE')], False) <> nil then
        begin
          GS.ElipsisButton := True;
          GS.ElipsisHint := HTitres.Mess[8];
        end;
        // JTR - eQualité 11203 (formules de qté)
        if (TOBL.FieldExists('GL_INDICEFORMULE')) and (TOBL.GetValue('GL_INDICEFORMULE')>0) then
        begin
          GS.ElipsisButton:=True ;
          GS.ElipsisHint:=TraduireMemoire('Formule Calcul Quantité') ;
        end ;
      end;
    end;
    AllerSurLigSuivante(ACol, ARow, OldCol, OldRow);
  end;
end;

procedure TFTicketFO.GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var
  {$IFDEF CHR}
    ffo:boolean;
  {$ENDIF}
  TobLigne, TobArticle: tob;
begin
  {$IFDEF CHR}
  ffo := True;
  {$ENDIF}
  if csDestroying in ComponentState then Exit;
  if PieceContainer.Action = taConsult then Exit;
  if GS.Cells[ACol, ARow] <> '' then DejaSaisie := True;
  if GS.Cells[ACol, ARow] = StCellCur then Exit;
  TOBLigneExiste(ARow);
  FormateZoneSaisie(ACol, ARow);
  if ACol = SG_RefArt then
  begin
    TraiteArticle(ACol, ARow, Cancel, False, False, 1);
{$IFDEF GESCOM}  // JTR - Ne pas cumuler des articles avec articles financiers dans un même ticket
    if not CumulArtOk(Acol,Arow,Cancel) then exit;
    RemiseDuTarif(Arow); // JTR - Démarque obligatoire
{$ENDIF GESCOM}
  end
  else if ACol = SG_Rep then TraiteRepres(ACol, ARow, Cancel)
  else if ACol = SG_Dep then TraiteDepot(ACol, ARow, Cancel)
  else if ((ACol = SG_QF) or (ACol = SG_QS)) then
  begin
    if not TraiteQte(ACol, ARow) then Cancel := True;
  end
  else if ACol = SG_Lib then TraiteLibelle(ARow)
  else if ACol = SG_PxNet then TraitePrixNet(ACol, ARow, Cancel)
  else if ACol = SG_Px then TraitePrix(ARow, Cancel)
  else if ACol = SG_Rem then TraiteRemise(ARow, Cancel)
  else if ((ACol = SG_RV) or (ACol = SG_RL)) then TraiteMontantRemise(ACol, ARow, Cancel)
  else if ACol = SG_Montant then TraiteMontantNetLigne(ACol, ARow, Cancel)
  else if ACol = SG_TypeRem then TraiteDemarque(ACol, ARow, Cancel)
{$IFDEF CHR}
  else if ACol=SG_Regrpe then HRTraiteRegroupe(Self, TobPiece, TobRegrpe, nil, ACol, ARow, Cancel, ffo)
  else if ACol=SG_Folio then TraiteFolio(ACol, ARow, Cancel)
{$ENDIF}
  else TraiteLesDivers(ACol, ARow);

  { On doit relancer la recherche tarif si
    * Piece de saisie et non de transformation
    * système non verrouillé
    * quantité <> 0
    * dépot
    * article
    * circuit
    * quantité
    * date de livraison
  }
  if (GetParamSoc('SO_PREFSYSTTARIF')) and (not PieceContainer.TransfoPiece) then
  begin
    TobLigne := GetTobLigne(TobPiece, ARow);
    if ((TobLigne <> nil)
      and (TobLigne.GetValue('GL_QTESTOCK') <> 0)
      and ((ACol = SG_Dep) //Champ dépot
      or (ACol = SG_Aff) //Champ affaire
      //or (ACol = SG_RefArt) //Champ article
      or (ACol = SG_QF) //Champ quantité facturée
      or (ACol = SG_QS) //Champ quantité de stock
      or (Acol = SG_DateLiv) //Champ date de livraison
      )
      ) then
    begin
      TobArticle := FindTobArtRow(TOBPiece, TOBArticles, ARow);
      TobLigne.SetBoolean ('RECALCULTARIF', TobLigne.GetBoolean ('GL_BLOQUETARIF'));
      RechercheTarifsDepuisPiece('LIGNE', PieceContainer.Action, TobTiers, TobArticle, TobPiece, TobLigne, TobLigneTarif, TobTarif, PieceContainer.DEV);
      AfficheLaLigne(ARow);
      TobPiece.PutValue('GP_TARIFSGROUPES', 'X');
    end;
  end;

  if not Cancel then
  begin
    CalculeLaSaisie(ACol, ARow, False);
    StCellCur := GS.Cells[ACol, ARow];
  end;
end;

procedure TFTicketFO.GSElipsisClick(Sender: TObject);
var TOBL: TOB;
  StCode: string;
  PEntSup : double;
  TobC : TOB;
begin
  TOBLigneExiste(GS.Row);
  // Articles
  if GS.Col = SG_RefArt then ZoomOuChoixArt(GS.Col, GS.Row) else
    if GS.Col = SG_Rep then ZoomOuChoixRep(GS.Col, GS.Row) else
    if GS.Col = SG_Dep then ZoomOuChoixDep(GS.Col, GS.Row) else
    if GS.Col = SG_Lib then ZoomOuChoixLib(GS.Col, GS.Row) else
    {$IFDEF CHR}
    if GS.Col = SG_Regrpe then HRZoomOuChoixRegroupe(Self,TobPiece, GS.Col, GS.Row, True) else
    {$ENDIF} // FIN CHR
    // Conditionnements
    if ((GS.Col = SG_QF) or (GS.Col = SG_QS)) then
  begin
    TOBL := GetTOBLigne(TOBPiece, GS.Row);
    if TOBL <> nil then
    begin
      if (TOBL.FieldExists('GL_INDICEFORMULE')) and (TOBL.GetValue('GL_INDICEFORMULE')>0) then
      begin
        RappelFormuleQte (GS.Row)  ;
      end else
      begin
        StCode := GetCondRecherche(GS, TOBL.GetValue('GL_CODECOND'), TOBL.GetValue('GL_ARTICLE'), TOBConds, VenteAchat);
        if stCode <> '' then
        begin
          TobC := TobConds.FindFirst (['GCO_ARTICLE', 'GCO_CODECOND'],
                                      [TobL.GetValue ('GL_ARTICLE'), stcode], False);
          if TobC <> nil then
          begin
            TOBL.PutValue('GL_CODECOND', TOBC.GetValue('GCO_CODECOND'));
            if (TobL.GetValue ('GL_QTEFACT') <> 0) and (TobC.GetValue ('GCO_NBARTICLE') <> 0) then
            begin
              PEntSup := TobL.GetValue ('GL_QTEFACT') / TobC.GetValue ('GCO_NBARTICLE');
              if PEntSup <> Trunc (PEntSup) then PEntSup := Trunc(PEntSup) + 1;
              TobL.PutValue ('GL_QTEFACT', PentSup * TobC.GetValue ('GCO_NBARTICLE'));
            end else
              TobL.PutValue ('GL_QTEFACT', TOBC.GetValue('GCO_NBARTICLE'));
            if PieceContainer.Action = taCreat then
            begin
              TOBL.PutValue('GL_QTESTOCK', TobL.GetValue ('GL_QTEFACT'));
            end;
            TOBL.PutValue('GL_QTERESTE', TobL.GetValue ('GL_QTEFACT'));

            GS.Cells[GS.Col, GS.Row] := StrF00(TobL.GetValue ('GL_QTEFACT'), V_PGI.OkDecQ);
            TraiteQte(GS.Col, GS.Row);
            TOBPiece.PutValue('GP_RECALCULER', 'X');
            CalculeLaSaisie(GS.Col, GS.Row, False);
          end;
        end;
      end;
    end;
  end else
    if GS.Col = SG_TypeRem then
      FOGetDemarqueSaisi(GS, HTitres.Mess[32], tlLocate, FideliteCli.Fidelite)
      else
    ;
  // Champ suivant
  EnvoieToucheGrid;
end;

procedure TFTicketFO.AllerSurLigSuivante(ACol, ARow, OldCol, OldRow: Integer);
var NoLig, ColPrix, NewCol, NewRow: Integer;
  ChercheLigne: Boolean;
begin
  if FOGetParamCaisse('GPK_LIGNESUIVAUTO') <> 'X' then
  begin
    StCellCur := GS.Cells[GS.Col, GS.Row];
    Exit;
  end;
  NewRow := GS.Row;
  NewCol := GS.Col;
  if SurLigSuivante then
  begin
    SurLigSuivante := False;
    if SG_PxNet <> -1 then ColPrix := SG_PxNet else ColPrix := SG_Px;
    if FOGetParamCaisse('GPK_AFFPRIXTIC') = '001' then ColPrix := SG_Montant;
    ChercheLigne := True;
    if (ColPrix <> -1) and (GS.Cells[SG_RefArt, ARow] <> '') and (Valeur(GS.Cells[ColPrix, ARow]) = 0) then
    begin
      // on se positionne dans la colonne prix car il n'est pas renseigné
      NewRow := ARow;
      NewCol := ColPrix;
      // pour passer à la ligne suivante après la saisie du prix
      if ZoneAccessible(NewCol, NewRow) then
      begin
        SurLigSuivante := True;
        ChercheLigne := False;
      end;
    end;
    if ChercheLigne then
    begin
      // on se positionne sur la 1ère ligne vierge qui suit
      for NoLig := ARow to GS.RowCount - 1 do
      begin
        if GetCodeArtUnique(TOBPiece, NoLig) = '' then
        begin
          if (SG_Lib = -1) or (GS.Cells[SG_Lib, NoLig] = '') then Break;
        end;
      end;
      NewRow := NoLig;
      NewCol := SG_RefArt;
    end;
  end else
  begin
    AffichePaveChoisi(GS.Col);
  end;
  // mémorisation de la valeur de la cellule avant modification
  if (GS.Row = NewRow) and (GS.Col = NewCol) then
  begin
    // on reste dans la cellule prévue initialement
    StCellCur := GS.Cells[GS.Col, GS.Row];
  end else
  begin
    // on se positionne dans une autre cellule que celle prévue initialement
    // les événements OnRow..., OnCol... et OnCell... vont à nouveau se déclencher
    // il faut donc conserver la valeur de la cellule d'où l'on vient.
    StCellCur := GS.Cells[OldCol, OldRow];
    GS.Row := NewRow;
    GS.Col := NewCol;
  end;
end;

function TFTicketFO.GereElipsis(LaCol: integer): boolean;
begin
  Result := False;
  if LaCol = SG_RefArt then
  begin
    if FOJaiLeDroit(38) then Result := FOGetArticleRecherche(GS, HTitres.Mess[1], PieceContainer.Cledoc.NaturePiece);
  end else
    if LaCol = SG_TypeRem then Result := FOGetDemarqueSaisi(GS, HTitres.Mess[32], tlLocate, FideliteCli.Fidelite) else
    if LaCol = SG_Rep then Result := GetRepresentantSaisi(GS, HTitres.Mess[11], TOBTiers.GetValue('T_ZONECOM'), tlLocate, TOBPiece) else
    if LaCol = SG_Dep then Result := GetDepotSaisi(GS, HTitres.Mess[12], tlLocate) else
    if LaCol = SG_Lib then Result := GetLibellesAutos(GS, HTitres.Mess[16]) else
    ;
end;

procedure TFTicketFO.GSDblClick(Sender: TObject);
begin
  TOBLigneExiste(GS.Row);
  if GS.Col = SG_RefArt then ZoomOuChoixArt(GS.Col, GS.Row) else
    if GS.Col = SG_Rep then ZoomOuChoixRep(GS.Col, GS.Row) else
    if GS.Col = SG_Dep then ZoomOuChoixDep(GS.Col, GS.Row) else
    if GS.Col = SG_Lib then ZoomOuChoixLib(GS.Col, GS.Row) else
    if GS.Col = SG_TypeRem then FOGetDemarqueSaisi(GS, HTitres.Mess[32], tlLocate, FideliteCli.Fidelite) else
    ;
  // Champ suivant
  EnvoieToucheGrid;
end;

procedure TFTicketFO.GSEnter(Sender: TObject);
var bc, Cancel: boolean;
  ACol, ARow: integer;
begin
  if (GP_TIERS.Text = '') and ((FOGetParamCaisse('GPK_CLISAISIE') = '-') or (FOGetParamCaisse('GPK_CLIREPRISE') <> '-')) then
    ForceTiers(FODonneClientDefaut);
  AjouteRepres(GP_REPRESENTANT.Text, TypeCom, TOBComms);
  if ((GP_TIERS.Text = '') or (GP_TIERS.Text <> TOBTiers.GetValue('T_TIERS'))) then
  begin
    HPiece.Execute(5, Caption, '');
    if GP_TIERS.CanFocus then GP_TIERS.SetFocus else GotoEntete;
    Exit;
  end;
  if GP_DEVISE.Value = '' then
  begin
    HPiece.Execute(32, Caption, '');
    TrouveDevisePiece;
    if GP_DEVISE.CanFocus then GP_DEVISE.SetFocus else GotoEntete;
    Exit;
  end;
  bc := False;
  Cancel := False;
  ACol := GS.Col;
  ARow := GS.Row;
  GSRowEnter(GS, GS.Row, bc, False);
  GSCellEnter(GS, ACol, ARow, Cancel);
  EnabledGrid;
  DejaRentre := True;
end;

{==============================================================================================}
{=============================== Actions sur le Entête ========================================}
{==============================================================================================}
{$IFNDEF FOS5}

procedure TFTicketFO.AfficheTaxes;
begin
  LTotalTaxesDEV.Caption := HTotalTaxesDEV.Text;
end;
{$ENDIF}

procedure TFTicketFO.PEnteteExit(Sender: TObject);
begin
  if csDestroying in ComponentState then Exit;
  TOBPiece.GetEcran(Self, PEntete);
  AjouteRepres(GP_REPRESENTANT.Text, TypeCom, TOBComms);
end;

procedure TFTicketFO.GP_REPRESENTANTEnter(Sender: TObject);
begin
  OldRepresentant := '';
  if FORepresentantExiste(GP_REPRESENTANT.Text, TypeCom, TOBComms) then
    OldRepresentant := GP_REPRESENTANT.Text;
  AffichePaveChoisi(41);
end;

procedure TFTicketFO.GP_REPRESENTANTExit(Sender: TObject);
var VdOk, VdExiste: boolean;
begin
  if csDestroying in ComponentState then Exit;
  VdOk := True;
  if GP_REPRESENTANT.Text = '' then GP_REPRESENTANT.Text := FOPreAffecteVendeur(TOBTiers);
  if (OldRepresentant <> '') and (GP_REPRESENTANT.Text <> OldRepresentant) and (not FOJaiLeDroit(36)) then
  begin
    GP_REPRESENTANT.Text := OldRepresentant;
    if GP_REPRESENTANT.CanFocus then GP_REPRESENTANT.SetFocus;
    Exit;
  end;
  if (GP_REPRESENTANT.Text <> '') or (FOGetParamCaisse('GPK_VENDOBLIG') = 'X') then
  begin
    if not RepresExiste(GP_REPRESENTANT.Text) then // JTR - eQualité 11295
    begin
      if CommercialFerme(GP_REPRESENTANT.Text) then
      begin
        GP_REPRESENTANT.Text := OldRepresentant;
        if GP_REPRESENTANT.CanFocus then GP_REPRESENTANT.SetFocus;
        exit;
      end;
    end;
    VdExiste := FORepresentantExiste(GP_REPRESENTANT.Text, TypeCom, TOBComms);
    if (VdExiste) or
      (GetRepresentantEntete(GP_REPRESENTANT, HTitres.Mess[31], TOBTiers.GetValue('T_ZONECOM'), TOBPiece)) then
    begin
      if GP_REPRESENTANT.Text <> TOBPiece.GetValue('GP_REPRESENTANT') then
      begin
        AjouteRepres(GP_REPRESENTANT.Text, TypeCom, TOBComms);
        BalayeLignesRep;
      end;
    end;
    if FORepresentantExiste(GP_REPRESENTANT.Text, TypeCom, TOBComms) then
    begin
      GereCommercialEnabled;
      TOBPiece.PutValue('GP_REPRESENTANT', GP_REPRESENTANT.Text);
    end else
    begin
      if GP_REPRESENTANT.CanFocus then GP_REPRESENTANT.SetFocus;
      VdOk := False;
    end;
  end;
  if (VdOk) and (RetourGS) then
  begin
    RetourGS := False;
    if FormEcheance = nil then
    begin
      if GS.CanFocus then GS.SetFocus;
    end else
    begin
      if FormEcheance.GSReg.CanFocus then FormEcheance.GSReg.SetFocus;
    end;
  end;
end;

procedure TFTicketFO.GP_REPRESENTANTElipsisClick(Sender: TObject);
var VdOk: Boolean;
begin
  if (OldRepresentant <> '') and (GP_REPRESENTANT.Text <> OldRepresentant) and (not FOJaiLeDroit(36)) then
  begin
    GP_REPRESENTANT.Text := OldRepresentant;
    if GP_REPRESENTANT.CanFocus then GP_REPRESENTANT.SetFocus;
    Exit;
  end;
  VdOk := FORepresentantExiste(GP_REPRESENTANT.Text, TypeCom, TOBComms);
  {if (VdOk) or
    (GetRepresentantEntete(GP_REPRESENTANT, HTitres.Mess[31], TOBTiers.GetValue('T_ZONECOM'), TOBPiece)) then}
  if GetRepresentantEntete(GP_REPRESENTANT, HTitres.Mess[31], TOBTiers.GetValue('T_ZONECOM'), TOBPiece) then
  begin
    if GP_REPRESENTANT.Text <> OldRepresentant then
    begin
      if not RepresExiste(GP_REPRESENTANT.Text) then // JTR - eQualité 11295
      begin
        if CommercialFerme(GP_REPRESENTANT.Text) then
        begin
          GP_REPRESENTANT.Text := OldRepresentant;
          exit;
        end;
      end;
      AjouteRepres(GP_REPRESENTANT.Text, TypeCom, TOBComms);
      BalayeLignesRep;
      OldRepresentant := GP_REPRESENTANT.Text;
      if not VdOk then FOSimuleClavier(VK_TAB);
    end;
  end;
  GereCommercialEnabled;
  TOBPiece.PutValue('GP_REPRESENTANT', GP_REPRESENTANT.Text);
end;

procedure TFTicketFO.GP_REPRESENTANTDblClick(Sender: TObject);
begin
  GP_REPRESENTANTElipsisClick(Sender);
end;

procedure TFTicketFO.GP_DATEPIECEExit(Sender: TObject);
begin
  if csDestroying in ComponentState then Exit;
  ExitDatePiece;
end;

function TFTicketFO.ExitDatePiece: Boolean;
var Err: integer;
  DD: TDateTime;
begin
  Result := False;
  Err := ControleDate(GP_DATEPIECE.Text);
  if Err > 0 then
  begin
    HPiece.Execute(17 + Err, caption, '');
    if GP_DATEPIECE.CanFocus then GP_DATEPIECE.SetFocus else GotoEntete;
    Exit;
  end;
  DD := StrToDate(GP_DATEPIECE.Text);
  if DD < V_PGI.DateDebutEuro then
  begin
    HPiece.Execute(25, caption, ' ' + DateToStr(V_PGI.DateDebutEuro));
    if GP_DATEPIECE.CanFocus then GP_DATEPIECE.SetFocus else GotoEntete;
    Exit;
  end;
  {if ((DD <= VH_GC.GCDateClotureStock) and (VH_GC.GCDateClotureStock > 100)) then
  begin
    HPiece.Execute(28, caption, '');
    if GP_DATEPIECE.CanFocus then GP_DATEPIECE.SetFocus else GotoEntete;
    Exit;
  end; } //JSI le 09/10/04 
  PieceContainer.Cledoc.DatePiece := StrToDate(GP_DATEPIECE.Text);
  PutValueDetail(TOBPiece, 'GP_DATEPIECE', PieceContainer.Cledoc.DatePiece);
  IncidenceTauxDate;
  Result := True;
end;

procedure TFTicketFO.GP_TIERSElipsisClick(Sender: TObject);
begin
  if IdentifieTiers(True) then
  begin
    CliCur := GP_TIERS.Text;
    FOSimuleClavier(VK_TAB);
  end;
end;

procedure TFTicketFO.GP_TIERSDblClick(Sender: TObject);
begin
  if IdentifieTiers(True) then
  begin
    CliCur := GP_TIERS.Text;
    FOSimuleClavier(VK_TAB);
  end;
end;

procedure TFTicketFO.GP_TIERSExit(Sender: TObject);
var
  ClientVide: boolean;
begin
  if csDestroying in ComponentState then Exit;
  GereTiersEnabled;
  if GeneCharge then Exit;
  if PieceContainer.Action = taConsult then Exit;
//  if CliCur <> GP_TIERS.Text then
  ClientVide := (CliCur = '');
  if TOBTiers.GetValue('T_TIERS') <> GP_TIERS.Text then
  begin
    if IdentifieTiers(False) then CliCur := GP_TIERS.Text;
  end;
  // si on vient de renseigner le code client et si le vendeur est saisi, on passe dans la grille
  if (ClientVide) and (GP_REPRESENTANT.Text <> '') then RetourGS := True;

  if RetourGS then
  begin
    RetourGS := False;
    if FormEcheance = nil then
    begin
      if GS.CanFocus then GS.SetFocus;
    end else
    begin
      if FormEcheance.GSReg.CanFocus then FormEcheance.GSReg.SetFocus;
    end;
  end;
end;

procedure TFTicketFO.GP_TIERSEnter(Sender: TObject);
begin
  CliCur := GP_TIERS.Text;
  AffichePaveChoisi(40);
end;

procedure TFTicketFO.GP_DEPOTChange(Sender: TObject);
var i, IndiceLot, NbSel: integer;
  IndiceSerie: integer;
  TOBL: TOB;
  Dep, OldDep: string;
  Okok: boolean;
begin
  if GeneCharge then Exit;
  if GP_DEPOT.Value = '' then Exit;
  OldDep := TOBPiece.GetValue('GP_DEPOT');
  NbSel := GS.NbSelected;
  Okok := False;
  if ((OldDep <> GP_DEPOT.Value) and (TOBArticles.Detail.Count > 0)) then
  begin
    if HPiece.Execute(7, Caption, '') = mrYes then
    begin
      for i := 0 to TOBPiece.Detail.Count - 1 do
      begin
        if ((NbSel = 0) or (GS.IsSelected(i + 1))) then
        begin
          TOBL := TOBPiece.Detail[i];
          {$IFNDEF CCS3}
          IndiceLot := TOBL.GetValue('GL_INDICELOT');
          if IndiceLot > 0 then Continue;
          IndiceSerie := TOBL.GetValue('GL_INDICESERIE');
          if IndiceSerie > 0 then Continue;
          {$ENDIF}
          Dep := TOBL.GetValue('GL_DEPOT');
          if Dep = OldDep then TOBL.PutValue('GL_DEPOT', GP_DEPOT.Value);
        end;
      end;
      GS.ClearSelected;
      Okok := True;
    end;
  end;
  if ((NbSel > 0) and (Okok)) then
  begin
    TOBPiece.PutValue('GP_DEPOT', OldDep);
    GP_DEPOT.Value := OldDep;
  end;
end;

procedure TFTicketFO.GP_ETABLISSEMENTChange(Sender: TObject);
var Etab, ListeDepot: string;
begin
  if ((PieceContainer.Action <> taCreat) and (not PieceContainer.DuplicPiece)) then Exit;
  Etab := GP_ETABLISSEMENT.Value;
  if Etab = '' then Exit;
  if Etab = TOBPiece.GetValue('GP_ETABLISSEMENT') then Exit;
  PutValueDetail(TOBPiece, 'GP_ETABLISSEMENT', Etab);
  if not VH_GC.GCMultiDepots then GP_DEPOT.Value := Etab;
  PieceContainer.Cledoc.Souche := GetSoucheG(PieceContainer.Cledoc.NaturePiece, Etab, '');
  PieceContainer.Cledoc.NumeroPiece := GetNumSoucheG(PieceContainer.Cledoc.Souche);
  PutValueDetail(TOBPiece, 'GP_SOUCHE', PieceContainer.Cledoc.Souche);
  PutValueDetail(TOBPiece, 'GP_NUMERO', PieceContainer.Cledoc.NumeroPiece);
  // Si on est en contexte mode on ne prend que les dépôts sont gérés sur site
  if (ctxMode in V_PGI.PGIContexte) then
  begin
    GP_DEPOT.Plus := 'GDE_SURSITE="X"';
  end;
  // Appel de la fonction renvoyant la liste des dépôts liés à l'établissement
  if VH_GC.GCMultiDepots then
  begin
    ListeDepot := ListeDepotParEtablissement(Etab);
    if ListeDepot <> '' then
    begin
      if GP_DEPOT.Plus <> '' then GP_DEPOT.Plus := GP_DEPOT.Plus + ' AND';
      GP_DEPOT.Plus := GP_DEPOT.Plus + ' GDE_DEPOT in (' + ListeDepot + ')';
      GP_DEPOT.Value := Copy(ListeDepot, 2, 3); // Le dépôt par défaut est le premier de la liste
    end
    else GP_DEPOT.Value := VH_GC.GCDepotDefaut; // Le dépôt par défaut est celui des paramètres sociétés
  end;
  {$IFDEF TAXEUS}
  // prise en compte du modèle de taxe de l'établissement
  PutValueDetail(TOBPiece, 'GP_CODEMODELETAXE', GetModeleTaxeEtab(Etab));
  {$ENDIF}
end;

procedure TFTicketFO.GP_REGIMETAXEChange(Sender: TObject);
var Reg: string;
begin
  if GeneCharge then Exit;
  if ((PieceContainer.Action <> taCreat) and (not PieceContainer.DuplicPiece) and (not SaisieTypeAffaire)) then Exit;

  Reg := GP_REGIMETAXE.Value;
  if Reg = '' then Exit;
  if (Reg = TOBPiece.GetValue('GP_REGIMETAXE')) and not (ChangeComplEntete) then Exit;
  {$IFDEF TAXEUS}
  ChangeRegimeTaxePiece(Reg, TOBPiece);
  {$ELSE}
  PutValueDetail(TOBPiece, 'GP_REGIMETAXE', Reg);
  PutValueDetail(TOBPiece, 'GP_RECALCULER', 'X');
  {$ENDIF}
  if ChangeComplEntete then CalculeLaSaisie(-1, -1, False);
end;

function TFTicketFO.GetLeTauxDevise(CodeD: string; var DateTaux: TDateTime; DateP: TDateTime): Double;
var Taux: Double;
  AvecT, ChoixT, Taux1: Boolean;
  ii: integer;
begin
  AvecT := False;
  ChoixT := False;
  Taux1 := False;
  case PieceContainer.Action of
    taCreat: AvecT := True;
    taModif:
      begin
        if PieceContainer.DuplicPiece then AvecT := True else
          if ((PieceContainer.TransfoPiece) and (GetInfoParPiece(PieceContainer.NewNature, 'GPP_TYPEECRCPTA') = 'NOR')) then AvecT := True;
      end;
  end;
  if AvecT then
  begin
    Taux := GetTaux(CodeD, DateTaux, DateP);
    if ((CodeD <> V_PGI.DevisePivot) and (not EstMonnaieIN(CodeD))) then
    begin
      if VH_GC.GCAlerteDevise = 'JOU' then ChoixT := (DateTaux <> DateP) else
        if VH_GC.GCAlerteDevise = 'SEM' then ChoixT := (NumSemaine(DateTaux) <> NumSemaine(DateP)) else
        if VH_GC.GCAlerteDevise = 'MOI' then ChoixT := (GetPeriode(DateTaux) <> GetPeriode(DateP));
      if Taux = 1 then
      begin
        Taux1 := True;
        ChoixT := True;
      end;
      if ChoixT then
      begin
        if Taux1 then ii := HPiece.Execute(16, caption, '') else ii := HPiece.Execute(15, caption, '');
        if ii = mrYes then
        begin
          {$IFNDEF SANSCOMPTA}
          FicheChancel(CodeD, True, DateP, taCreat, (DateP >= V_PGI.DateDebutEuro));
          {$ENDIF}
          Taux := GetTaux(CodeD, DateTaux, DateP);
        end;
      end;
    end;
  end else
  begin
    Taux := TOBPiece.GetValue('GP_TAUXDEV');
  end;
  Result := Taux;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  TrouveDevisePiece : détermine la devise de saisie de la pièce
///////////////////////////////////////////////////////////////////////////////////////

procedure TFTicketFO.TrouveDevisePiece;
var Stg: string;
begin
  Stg := FOGetParamCaisse('DEVISECAISSE'); // devise de la caisse
  if Stg = '' then Stg := V_PGI.DevisePivot; // devise de tenue du dossier
  GP_DEVISE.Value := Stg;
end;

procedure TFTicketFO.GP_DEVISEChange(Sender: TObject);
begin
  if GP_DEVISE.Value = '' then
  begin
    FillChar(PieceContainer.DEV, Sizeof(PieceContainer.Dev), #0);
    Exit;
  end;
  PieceContainer.DEV.Code := GP_DEVISE.Value;
  GetInfosDevise(PieceContainer.DEV);
  PieceContainer.DEV.Taux := GetLeTauxDevise(PieceContainer.DEV.Code, PieceContainer.DEV.DateTaux, PieceContainer.Cledoc.DatePiece);
  PutValueDetail(TOBPiece, 'GP_DEVISE', PieceContainer.DEV.Code);
  PutValueDetail(TOBPiece, 'GP_TAUXDEV', PieceContainer.DEV.Taux);
  PutValueDetail(TOBPiece, 'GP_DATETAUXDEV', PieceContainer.DEV.DateTaux);
  if PieceContainer.Dev.Code = FODonneCodeEuro then PieceContainer.DEV.Symbole := SIGLEEURO;
  ChangeFormatDevise(Self, PieceContainer.DEV.Decimale, PieceContainer.DEV.Symbole);
  if PieceContainer.DEV.Code <> V_PGI.DevisePivot then
  begin
    PopD.Items[2].Checked := True;
    GP_DEVISE.Enabled := True;
  end else
  begin
    if PopD.Items[2].Checked then PopD.Items[0].Checked := True;
    GP_DEVISE.Enabled := False;
  end;
  if not GeneCharge then
  begin
    AfficheEuro;
    ValideLaCotation(PieceContainer);
  end;
end;

{==============================================================================================}
{=============================== Actions sur le Pied ==========================================}
{==============================================================================================}

procedure TFTicketFO.GP_ESCOMPTEExit(Sender: TObject);
begin
  if csDestroying in ComponentState then Exit;
  if GP_ESCOMPTE.Value = TOBPiece.GetValue('GP_ESCOMPTE') then Exit;
  PutValueDetail(TOBPiece, 'GP_ESCOMPTE', GP_ESCOMPTE.Value);
  PutValueDetail(TOBPiece, 'GP_RECALCULER', 'X');
  CalculeLaSaisie(-1, -1, TotEnListe);
end;

procedure TFTicketFO.GP_REMISEPIEDExit(Sender: TObject);
begin
  if csDestroying in ComponentState then Exit;
  if GP_REMISEPIED.Value = TOBPiece.GetValue('GP_REMISEPIED') then Exit;
  PutValueDetail(TOBPiece, 'GP_REMISEPIED', GP_REMISEPIED.Value);
  PutValueDetail(TOBPiece, 'GP_RECALCULER', 'X');
  CalculeLaSaisie(-1, -1, TotEnListe);
end;

{==============================================================================================}
{=============================== Manipulation des LIGNES ======================================}
{==============================================================================================}

procedure TFTicketFO.AfficheLaLigne(ARow: integer; AffLCD: Boolean = False);
var TOBL: TOB;
  i, QTE: integer;
  MontantHT, PX: Double;
  Sigle: string;
  NbDec: Integer;
begin
  MontantHT := 0;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL <> nil then
  begin
    TOBL.PutLigneGrid(GS, ARow, False, False, LesColonnes);
    if TOBL.GetValue('GL_TYPEDIM') = 'GEN' then
    begin
      if SG_RefArt <> -1 then GS.Cells[SG_RefArt, ARow] := TOBL.GetValue('GL_CODESDIM');
      Qte := TOBL.GetValue('GL_QTEFACT');
      if GP_FACTUREHT.Checked then Px := TOBL.GetValue('GL_PUHTDEV') else Px := TOBL.GetValue('GL_PUTTCDEV');
      MontantHT := MontantHT + (Px * Qte);
      if SG_Total <> -1 then GS.Cells[SG_Total, Arow] := FloatToStr(MontantHT);
    end;
  end;
  for i := 1 to GS.ColCount - 1 do FormateZoneSaisie(i, ARow);
  if (TOBL <> nil) and (AffLCD) then
  begin
    GetTotalaAfficher(TOBL, MontantHT, NbDec, Sigle);
    AffichageLCD(TOBL.GetValue('GL_LIBELLE'), TOBL.GetValue('GL_QTEFACT'), MontantHT, NbDec, Sigle, [ofInterne, ofClient], qaLigne, False);
  end;
end;

procedure TFTicketFO.TOBLigneExiste(ARow: integer);
var
  TOBL: TOB;
begin
  if PieceContainer.Action = taConsult then Exit;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then CreerTOBLignes(GS, TOBPiece, TOBTiers, TOBAffaire, ARow);
end;

function TFTicketFO.InitLaLigne(ARow: integer; NewQte : double; NewPrix: double = 0): T_ActionTarifArt;
var TOBL, TOBA, TOBCATA, TOBArtREf: TOB;
  sCode: string;
begin
  Result := ataOk;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  TOBA := FindTOBArtRow(TOBPiece, TOBArticles, ARow);
  TOBCATA := nil;
  TOBArtREf := nil;
  Result := PreAffecteLigne(PieceContainer, TOBL, TOBA, TOBCata, TOBArtRef, ARow, NewQte);
  {$IFDEF CHR}
  TOBL.PutValue('GL_DATEPRODUCTION', V_PGI.DateEntree) ;
  {$ENDIF}
  case Result of
    ataOk: ;
    ataAlerte: HPiece.Execute(0, Caption, '');
    ataCancel:
      begin
        HPiece.Execute(1, Caption, '');
        GS.Cells[SG_RefArt, ARow] := '';
        VideCodesLigne(TOBPiece, ARow);
        InitialiseTOBLigne(TOBPiece, TOBTiers, TOBAffaire, ARow);
      end;
  end;
  // Prix Unitaire forcé
  if NewPrix <> 0 then
  begin
    if TOBPiece.GetValue('GP_FACTUREHT') = 'X' then TOBL.PutValue('GL_PUHTDEV', NewPrix)
    else TOBL.PutValue('GL_PUTTCDEV', NewPrix);
  end;
  // Prix Unitaire modifiable
  if ((TOBPiece.GetValue('GP_FACTUREHT') = 'X') and (TOBL.GetValue('GL_PUHTDEV') = 0)) or
    ((TOBPiece.GetValue('GP_FACTUREHT') = '-') and (TOBL.GetValue('GL_PUTTCDEV') = 0)) then
  begin
    // pas de prix unitaire
    if TOBL.GetValue('GL_TYPEARTICLE') = 'FI' then
      sCode := 'X' // possibilité de définir le prix pour une opération de caisse
    else sCode := TOBA.GetValue('GA_REMISELIGNE'); // article remisable, prix unitaire modifiable
  end else
  begin
    if SG_PxNet <> -1 then
      sCode := TOBA.GetValue('GA_REMISELIGNE') // article remisable, prix net modifiable
    else sCode := '-'; // prix brut non modifiable
  end;
  TOBL.AddChampSupValeur('MODIFPU', sCode);
  // Choix du vendeur de la ligne
  FOPreAffecteVendeurLigne(TOBPiece, TOBL, TOBTiers, Arow);
  AfficheLaLigne(ARow);
  GereDescriptif(ARow, True);
end;

procedure TFTicketFO.ComplementLigne(ARow: integer);
var stArg: string;
begin
  TheTOB := GetTOBLigne(TOBPiece, ARow);
  if TheTOB = nil then exit;
  if not FOJaiLeDroit(7) then Exit;
  stArg := ActionToString(PieceContainer.Action);
  if (ctxAffaire in V_PGI.PGIContexte) or (SaisieTypeAffaire) then
    AGLLanceFiche('AFF', 'AFCOMPLLIGNE', '', '', stArg)
  else
    AGLLanceFiche('GC', 'GCCOMPLLIGNE', '', '', stArg);
  if ((TheTob <> nil) and (PieceContainer.Action <> taConsult)) then
  begin
    TheTOB.PutLigneGrid(GS, ARow, False, False, LesColonnes);
    if (ctxAffaire in V_PGI.PGIContexte) or (SaisieTypeAffaire) then
    begin
      CalculeLaSaisie(-1, ARow, False);
    end;
    ShowDetail(ARow);
  end;
  TheTob := nil;
end;

function TFTicketFO.PieceModifiee: boolean;
begin
  Result := False;
  if (PieceContainer.Action = taConsult) and (not AssignePiece) then Exit;
  Result := ((TOBPiece.IsOneModifie) or (TOBAdresses.IsOneModifie) or (TOBAdresses.IsOneModifie) or
    (PieceContainer.TransfoPiece) or (PieceContainer.DuplicPiece));
end;

{==============================================================================================}
{=============================== Manipulation de l'en-tête ====================================}
{==============================================================================================}

function TFTicketFO.ComplementEntete: boolean;
var
  stComp, ModeRegle_O, Ressource_O, sTiersLiv, sTiersFac: string;
  Ind, iNumTiersLiv, iNumTiersFac: integer;
  TOBL: TOB;
begin
  Result := False;
  if not FOJaiLeDroit(3) then Exit;
  ModeRegle_O := TobPiece.GetValue('GP_MODEREGLE');
  Ressource_O := TobPiece.GetValue('GP_RESSOURCE');
  TheTob := TobPiece;
  TobPiece.data := TobAdresses;
  if (ctxAffaire in V_PGI.PGIContexte) or (SaisieTypeAffaire) then
  begin
    stComp := '';
    if SaisieTypeAffaire then stComp := ';ENTETEAFFAIREPIECE';
    if (TobTiers.GetValue('T_PARTICULIER') = 'X') then stComp := stComp + ';PARTICULIER';
    AGLLanceFiche('AFF', 'AFCOMPLPIECE', '', '', ActionToString(PieceContainer.Action) + stComp);
  end else
  begin
    AGLLanceFiche('GC', 'GCCOMPLPIECE', '', '', ActionToString(PieceContainer.Action));
  end;
  if ((TheTob <> nil) and (PieceContainer.Action <> taConsult)) then
  begin
    ChangeComplEntete := True;
    TOBPiece.PutEcran(Self, PEntete);
    sTiersLiv := TobPiece.GetValue('GP_TIERSLIVRE');
    sTiersFac := TobPiece.GetValue('GP_TIERSFACTURE');
    iNumTiersLiv := 0;
    iNumTiersFac := 0;
    for Ind := 0 to TobPiece.Detail.Count - 1 do
    begin
      TOBL := TobPiece.Detail[Ind];
      if Ind = 0 then
      begin
        iNumTiersLiv := TOBL.GetNumChamp('GL_TIERSLIVRE');
        iNumTiersFac := TOBL.GetNumChamp('GL_TIERSFACTURE');
      end;
      TOBL.PutValeur(iNumTiersLiv, sTiersLiv);
      TOBL.PutValeur(iNumTiersFac, sTiersFac);
    end;
    // Verif si le mode de règlement a été modifié dans l'entête (affaire uniquement)
    if ctxAffaire in V_PGI.PGIContexte then
    begin
      if (ModeRegle_O <> TobPiece.GetValue('GP_MODEREGLE')) and (TOBEches.Detail.Count > 0) then
        TOBEches.ClearDetail;
    end;
    if Ressource_O <> TobPiece.GetValue('GP_RESSOURCE') then PieceVersLigneRessource(TOBPiece, nil, Ressource_O);

    {$IFDEF TAXEUS}
    ChangeAdresseTaxePiece(TOBPiece, TOBTiers, TOBAdresses, VenteALivrer);
    if TOBPiece.GetValue('GP_RECALCULER') = 'X' then CalculeLaSaisie(-1, -1, False);
    {$ENDIF}

    ChangeComplEntete := False;
    Result := True;
  end;
  TheTob := nil;
  TobPiece.data := nil;
end;

{==============================================================================================}
{=============================== Manipulation des tiers =======================================}
{==============================================================================================}

function TFTicketFO.IdentifieTiers(Click: boolean): boolean;
var
  sOldCli: string;
begin
  Result := True;
  if (not Click) and (GP_TIERS.Text <> '') and (GP_TIERS.Text = TOBTiers.GetValue('T_TIERS')) and
    (GP_TIERS.Text <> FODonneClientDefaut) then
  begin
    if not TestNatureTiers(PieceContainer.TCPiece.GetValue('GP_NATUREPIECEG'),PieceContainer.TCTiers.GetValue('T_NATUREAUXI')) then
      Result := False;
    Exit;
  end;

  sOldCli := GP_TIERS.Text;
  if sOldCli = '' then sOldCli := FODonneClientDefaut;
  if GP_TIERS.Text = FODonneClientDefaut then GP_TIERS.Text := '';
  // vérifie si le code saisi est un code client correct
  if (not Click) and (GP_TIERS.Text <> '') then
    Result := ExisteSQL('SELECT T_TIERS, T_NATUREAUXI FROM TIERS WHERE T_TIERS="' + GP_TIERS.Text + '"')
  else
    Result := False;
  // recherche des clients
  if not Result then
  begin
    if (GP_TIERS.Text = '') or (FOGetParamCaisse('GPK_CLISAISIENOM') <> 'X') or (IsNumeric(GP_TIERS.Text)) then
      Result := GetTiers(GP_TIERS)
    else
      Result := FOGetTiersMulLibelle(GP_TIERS);
  end;
  if Result then Result := ChargeTiers;
  if not Result then
  begin
    // JTR Annule choix tiers
//    GP_TIERS.Text := sOldCli;
    GP_TIERS.Text := '';
    if GP_TIERS.CanFocus then GP_TIERS.SetFocus;
    CliCur := TOBTiers.GetValue('T_TIERS');
  end;
  MajTobAvecTiers;
  if (Result) and (AssignePiece) then AssigneTiers;
  GereTiersEnabled;
end;

procedure TFTicketFO.ForceTiers(CodeTiers: string);
var
  Ok: boolean;
begin
  if (PieceContainer.Action = taConsult) and (not AssignePiece) then Exit;
  if (CodeTiers <> '') and ((GP_TIERS.Text = '') or (GP_TIERS.Text = FODonneClientDefaut)) then
  begin
    GP_TIERS.Text := CodeTiers;
    Ok := ChargeTiers;
    GereTiersEnabled;
    if Ok then
    begin
      MajTobAvecTiers;
      if (GP_TIERS.Focused) and (GS.CanFocus) then GS.SetFocus;
    end else
    begin
      if GP_TIERS.CanFocus then GP_TIERS.SetFocus;
    end;
  end;
end;

procedure TFTicketFO.CreationTiers;
var Stg: string;
begin
  if (PieceContainer.Action = taConsult) and (not AssignePiece) then Exit;
  if not FOJaiLeDroit(104) then Exit;
  //if Assigned(V_PGI.DispatchTT) then V_PGI.DispatchTT(8, taCreat, '', '', '') ;
  {$IFDEF MODE}
  Stg := AGLLanceFiche('MFO', 'GCTIERSFO', '', '', ActionToString(taCreat) + ';MONOFICHE;T_NATUREAUXI=CLI');
  {$ELSE}
  Stg := AGLLanceFiche('GC', 'GCTIERS', '', '', ActionToString(taCreat) + ';MONOFICHE;T_NATUREAUXI=CLI');
  {$ENDIF}
  if (Stg <> '') and (GP_TIERS.CanFocus) then ForceTiers(Stg);
end;

procedure TFTicketFO.ModificationTiers;
begin
  if (PieceContainer.Action = taConsult) and (not AssignePiece) then Exit;
  if (GP_TIERS.Text = '') or (GP_TIERS.Text = FODonneClientDefaut) then
    ChoixTiers
  else
  begin
    if not FOJaiLeDroit(35) then Exit;
    {$IFDEF MODE}
    AGLLanceFiche('MFO', 'GCTIERSFO', '', TOBTiers.GetValue('T_AUXILIAIRE'), ActionToString(taModif) + ';MONOFICHE;T_NATUREAUXI=CLI');
    {$ELSE}
    AGLLanceFiche('GC', 'GCTIERS', '', TOBTiers.GetValue('T_AUXILIAIRE'), ActionToString(taModif) + ';MONOFICHE;T_NATUREAUXI=CLI');
    {$ENDIF}
  end;
end;

procedure TFTicketFO.VoirSoldeTiers;
var
  Stg: string;
begin
  {$IFDEF GESCOM}
  if FOGetParamCaisse('GPK_LIAISONREG') <> 'X' then Exit;
  {$ENDIF}
  if (GP_TIERS.Text <> '') and (GP_TIERS.Text <> FODonneClientDefaut) then
  begin
    Stg := 'T_TIERS=' + GP_TIERS.Text + ';T_LIBELLE=' + LIBELLETIERS.Caption;
    AGLLanceFiche('MFO', 'MFOSOLDECLIENT', '', '', Stg);
  end;
end;

procedure TFTicketFO.AfficheSoldeTiers;
var
  AfficheOk: boolean;
  Symbole: string;
  SoldeCli: double;
begin
  {$IFDEF GESCOM}
  AfficheOk := (FOGetParamCaisse('GPK_LIAISONREG') = 'X');
  {$ELSE}
  AfficheOk := True;
  {$ENDIF}
  if GP_TIERS.Text = FODonneClientDefaut then AfficheOk := False;
  if AfficheOk then
  begin
    // affichage du solde du client
    SoldeCli := FOCalculSoldeClient(GP_TIERS.Text, False);
    if SoldeCli < 0 then
    begin
      SOLDETIERS.Flashing := True;
      SOLDETIERS.Font.Color := clRed;
    end else
    begin
      SOLDETIERS.Flashing := False;
      SOLDETIERS.Font.Color := clBlack;
    end;
    if V_PGI.DevisePivot = CODEEURO then Symbole := SIGLEEURO else Symbole := V_PGI.SymbolePivot;
    SOLDETIERS.Caption := StrfMontant(SoldeCli, 12, V_PGI.OkDecV, Symbole, True);
    TSOLDETIERS.Visible := True;
    if FormEcheance <> nil then FormEcheance.AfficheSoldeClient;
  end else
  begin
    SOLDETIERS.Caption := '';
    TSOLDETIERS.Visible := False;
  end;
end;

procedure TFTicketFO.ChoixTiers;
var
  Ok: Boolean;
  sOldCli: string;
begin
  if (PieceContainer.Action = taConsult) and (not AssignePiece) then Exit;
  sOldCli := GP_TIERS.Text;
  if sOldCli = '' then sOldCli := FODonneClientDefaut;
  if GP_TIERS.Text = FODonneClientDefaut then GP_TIERS.Text := '';
  Ok := ((GetTiers(GP_TIERS)) and (ChargeTiers));
  if not Ok then
  begin
    // JTR Annule choix tiers
//    GP_TIERS.Text := sOldCli;
    GP_TIERS.Text := '';
    if GP_TIERS.CanFocus then GP_TIERS.SetFocus;
    CliCur := TOBTiers.GetValue('T_TIERS');
  end;
  MajTobAvecTiers;
  if (Ok) and (AssignePiece) then AssigneTiers;
  GereTiersEnabled;
end;

procedure TFTicketFO.MajTobAvecTiers;
begin
  PutValueDetail(TOBPiece, 'GP_TIERS', GP_TIERS.Text);
  FOPutValueDetail(TOBEches, 'GPE_TIERS', GP_TIERS.Text);
  if FormEcheance <> nil then
  begin
    PutValueDetail(FormEcheance.TOBPiec, 'GP_TIERS', GP_TIERS.Text);
    FOPutValueDetail(FormEcheance.TOBEche, 'GPE_TIERS', GP_TIERS.Text);
  end;
end;

procedure TFTicketFO.AssigneTiers;
var Stg: string;
begin
  if not AssignePiece then Exit;
  PutValueDetail(TOBPiece, 'GP_TIERS', GP_TIERS.Text);
  Stg := TOBPIECE.GetValue('GP_TIERSLIVRE');
  PutValueDetail(TOBPiece, 'GP_TIERSLIVRE', Stg);
  Stg := TOBPIECE.GetValue('GP_TIERSFACTURE');
  PutValueDetail(TOBPiece, 'GP_TIERSFACTURE', Stg);
  Stg := TOBPIECE.GetValue('GP_TIERSPAYEUR');
  PutValueDetail(TOBPiece, 'GP_TIERSPAYEUR', Stg);
  FOPutValueDetail(TOBEches, 'GPE_TIERS', GP_TIERS.Text);
  FOPutValueDetail(TOBOpCais, 'GOC_TIERS', GP_TIERS.Text);
end;

procedure TFTicketFO.LectureCBFromKB;
var
  Stg, sTiers: string;
  Ind: integer;
begin
  if PieceContainer.Action = taConsult then Exit;

  // Lecture d'une carte depuis la piste du clavier
  Stg := FOLanceFicheLectureCBFromKB('', '', '');
  TOBEches.AddChampSupValeur('LECTURECBFROMKB', Stg);
  for Ind := 1 to 5 do sTiers := Trim(ReadTokenST(Stg));
  if sTiers <> '' then ForceTiers(sTiers);
end;

function TFTicketFO.TesteMargeMini(ARow: integer): boolean;
var TOBL, TOBA: TOB;
  RefU, QualM: string;
  PxValo, PuNet, MC: double;
  {$IFNDEF FOS5}
  MargeMini: double;
  OkUs, PasOkM: Boolean;
  ii: integer;
  {$ENDIF}
begin
  Result := False;
  if ARow <= 0 then Exit;
  if PieceContainer.Action = taConsult then Exit;
  if EstAvoir then Exit;
  if VenteAchat <> 'VEN' then Exit;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  RefU := TOBL.GetValue('GL_ARTICLE');
  if RefU = '' then Exit;
  TOBA := FindTOBArtRow(TOBPiece, TOBArticles, ARow);
  if TOBA = nil then Exit;
  QualM := TOBA.GetValue('GA_QUALIFMARGE');
  if QualM = '' then Exit;
  {$IFNDEF FOS5}
  MargeMini := TOBA.GetValue('GA_MARGEMINI');
  {$ENDIF}
  MC := 0;
  if VH_GC.GCMargeArticle = 'PMA' then PxValo := TOBL.GetValue('GL_PMAP') else
    if VH_GC.GCMargeArticle = 'PMR' then PxValo := TOBL.GetValue('GL_PMRP') else
    if VH_GC.GCMargeArticle = 'DPA' then PxValo := TOBL.GetValue('GL_DPA') else
    if VH_GC.GCMargeArticle = 'DPR' then PxValo := TOBL.GetValue('GL_DPR') else Exit;
  PuNet := TOBL.GetValue('GL_PUHTNET');
  {$IFNDEF FOS5}
  PasOkM := False;
  {$ENDIF}
  if QualM = 'MO' then
  begin
    MC := PuNet - PxValo;
    {$IFNDEF FOS5}
    PasOkM := (MC < MargeMini);
    {$ENDIF}
  end else if QualM = 'CO' then
  begin
    if TOBPiece.GetValue('GP_FACTUREHT') = '-' then PuNet := TOBL.GetValue('GL_PUTTCNET');
    if PxValo > 0 then MC := (PuNet / PxValo) else MC := 0;
    {$IFNDEF FOS5}
    if MC <> 0 then PasOkM := (MC < MargeMini);
    {$ENDIF}
  end else if QualM = 'PC' then
  begin
    if PuNet > 0 then MC := (PuNet - PxValo) / PuNet else MC := 0;
    {$IFNDEF FOS5}
    if MC <> 0 then PasOkM := (MC < MargeMini);
    {$ENDIF}
  end;
  TOBL.PutValue('MARGE', MC);
  {$IFNDEF FOS5}
  if PasOkM then
  begin
    if ((V_PGI.Superviseur) or (not VH_GC.GCBloqueMarge)) then
    begin
      HPiece.Execute(17, caption, '');
    end else
    begin
      ii := HPiece.Execute(23, caption, '');
      if ii <> mrYes then Result := True else
      begin
        {$IFDEF EAGLCLIENT}
        {AFAIREEAGL}
        {$ELSE}
        OkUs := ChangeUser;
        {$ENDIF}
        if ((not OkUs) or (not V_PGI.Superviseur)) then Result := True;
      end;
    end;
  end;
  {$ENDIF}
end;

function TFTicketFO.TesteRisqueTiers(Valid: boolean): boolean;
var TotR, CreditA, PlaF: Double;
  OkUs: boolean;
  InfoR: string;
begin
  Result := False;
  if SansRisque then Exit;
  if TOBTiers.GetValue('T_AUXILIAIRE') = '' then Exit;
  if ((QuestionRisque) and (not Valid)) then Exit;
  TotR := RisqueTiers + TOBPiece.GetValue('GP_TOTALTTC');
  Plaf := TOBTiers.GetValue('T_CREDITPLAFOND');
  if TOTR >= Plaf then
  begin
    QuestionRisque := True;
    InfoR := #13 + #10 + '(' + HTitres.Mess[21] + ' ' + Strf00(Plaf, V_PGI.OkDecV) + ' / ' + HTitres.Mess[22] + ' ' + Strf00(TOTR, V_PGI.OkDecV) + ')';
    if V_PGI.Superviseur then HPiece.Execute(13, caption, InfoR) else
    begin
      if HPiece.Execute(24, caption, '') = mrYes then
      begin
        {$IFDEF EAGLCLIENT}
        {AFAIREEAGL}
        OkUs := False;
        {$ELSE}
        OkUs := ChangeUser;
        {$ENDIF}
        if ((not OkUs) or (not V_PGI.Superviseur)) then Result := True;
      end else Result := True;
    end;
  end else
  begin
    CreditA := TOBTiers.GetValue('T_CREDITACCORDE');
    InfoR := #13 + #10 + '(' + HTitres.Mess[21] + ' ' + Strf00(CreditA, V_PGI.OkDecV) + ' / ' + HTitres.Mess[22] + ' ' + Strf00(TOTR, V_PGI.OkDecV) + ')';
    if ((CreditA > 0) and (TOTR > CreditA)) then
    begin
      HPiece.Execute(14, caption, '');
      QuestionRisque := True;
    end;
  end;
end;

procedure TFTicketFO.CalculeRisqueTiers;
var NaturePieceG: string;
begin
  RisqueTiers := 0;
  SansRisque := True;
  QuestionRisque := False;
  if PieceContainer.Action = taConsult then Exit;
  if EstAvoir then Exit;
  NaturePieceG := TOBPiece.GetValue('GP_NATUREPIECEG');
  if VenteAchat <> 'VEN' then Exit;
  if GetInfoParPiece(NaturePieceG, 'GPP_ENCOURS') <> 'X' then Exit;
  if TOBTiers.GetValue('T_CREDITPLAFOND') = 0 then Exit;
  RisqueTiers := RisqueTiersGC(TOBTiers) + RisqueTiersCPTA(TOBTiers, V_PGI.DateEntree);
  if ((PieceContainer.Action = taModif) and (not PieceContainer.DuplicPiece)) then RisqueTiers := RisqueTiers - TOBPiece.GetValue('GP_TOTALTTC');
  SansRisque := False;
end;

// JTR - Test nature tiers
function TFTicketFO.TestNatureTiers(Nature, Tiers : string) : boolean;
begin
  if (Nature = '') or (Tiers = '') or (pos(Tiers,GetInfoParpiece(Nature,'GPP_NATURETIERS')) = 0) then
    Result := False
    else
    Result := True;
end;

function TFTicketFO.ChargeTiers: boolean;
var
  tr: T_RechTiers;
  ChgTiers, Ok: boolean;
  OldT: string;
  TOBT: TOB;
  CodeLivr, CodeFact, CodePay, AuxiLivr, AuxiFact, AuxiPay: string;
  ProfilCalcSolde : String; // JTR Calcul solde tiers
begin
  Result := True;
  OldT := TOBTiers.GetValue('T_TIERS');
  ChgTiers := (GP_TIERS.Text <> OldT);
  tr := RemplirTOBTiers(TOBTiers, GP_TIERS.Text, PieceContainer.NewNature, True);
  case tr of
    trtOk: begin
             if not TestNatureTiers(TOBPiece.GetValue('GP_NATUREPIECEG'),TOBTiers.GetValue('T_NATUREAUXI')) then
             begin
               HPiece.Execute(5, Caption, '');
               TOBTiers.InitValeurs;
               tr := trtIncorrect;
             end else
              CalculeRisqueTiers;
           end;
    trtIncorrect: HPiece.Execute(8, Caption, '');
    trtFerme: HPiece.Execute(3, Caption, '');
  end;
  if tr <> trtOk then
  begin
    GP_TIERS.Text := TOBTiers.GetValue('T_TIERS');
    Result := False;
  end else
  begin
    if ChgTiers then
    begin
      CataFourn := False;
      if VenteAchat = 'ACH' then
        CataFourn := ExisteSQL('SELECT GCA_TIERS FROM CATALOGU WHERE GCA_TIERS="' + GP_TIERS.Text + '"');
    end;
  end;
  if GP_TIERS.Text = '' then
    Exit;
  LibelleTiers.Caption := FOMakeNomClient(TOBTiers);
  // JTR Calcul solde tiers
  ProfilCalcSolde := GetParamSoc('SO_TIERSDIVERS');
  if (ProfilCalcSolde = '') or ((ProfilCalcSolde <>'') and (PieceContainer.TCTiers.GetValue('T_PROFIL') <> ProfilCalcSolde)) then
  begin
    SOLDETIERS.Caption := '';
    TSOLDETIERS.Visible := True;
    SOLDETIERS.Visible := True;
    AfficheSoldeTiers;
  end else
  begin
    TSOLDETIERS.Visible := False;
    SOLDETIERS.Visible := False;
  end;
  // Fin JTR
  {Exceptions tiers}
  TOBT := TOBTiers.FindFirst(['GTP_NATUREPIECEG'], [PieceContainer.NewNature], False);
  if TOBT <> nil then
  begin
    NeedVisa := (TOBT.GetValue('GTP_VISA') = 'X');
    MontantVisa := TOBT.GetValue('GTP_MONTANTVISA');
  end;
  PieceContainer.InitTobs(Self);
  {Incidences tiers}
  if ((PieceContainer.Action = taCreat) or (PieceContainer.DuplicPiece)) then
  begin
    IncidenceTiers;
    if ((ChgTiers) and (not GeneCharge)) then
    begin
      IncidenceRefs;
      if not IncidenceTarif then
      begin
        PutValueDetail(TOBPiece, 'GP_RECALCULER', 'X');
        CalculeLaSaisie(-1, -1, TotEnListe);
      end;
    end;
  end;
  if (AssignePiece) and (not GeneCharge) then
  begin
    //TiersVersPiece(TOBTiers,TOBPiece) ;
    CodeLivr := TOBTiers.GetValue('T_TIERS');
    AuxiLivr := TOBTiers.GetValue('T_AUXILIAIRE');
    AuxiFact := TOBTiers.GetValue('T_FACTURE');
    AuxiPay := TOBTiers.GetValue('T_PAYEUR');
    if ((AuxiFact = '') or (AuxiFact = AuxiLivr)) then
    begin
      AuxiFact := AuxiLivr;
      CodeFact := CodeLivr;
    end else
    begin
      CodeFact := TiersAuxiliaire(AuxiFact, True);
    end;
    if ((AuxiPay = '') or (AuxiPay = AuxiFact)) then
    begin
      AuxiPay := AuxiFact;
      CodePay := CodeFact;
    end else
    begin
      CodePay := TiersAuxiliaire(AuxiPay, True);
    end;
    TOBPiece.PutValue('GP_TIERSLIVRE', CodeLivr);
    TOBPiece.PutValue('GP_TIERSFACTURE', CodeFact);
    TOBPiece.PutValue('GP_TIERSPAYEUR', CodePay);
    if GereAdresse then
      TiersVersAdresses(PieceContainer);
  end;
  GP_REPRESENTANT.Enabled := True;
  // Fidélité
  with FideliteCli do
  begin
    if not ChargeCarte(GP_TIERS.Text, FODonneClientDefaut, GP_ETABLISSEMENT.Value, '', '', True) then NewCarte;
    if PieceContainer.Action = taCreat then
      if Fidelite then PGIInfo(MessageFidelite);
    CumulPiece(TOBPiece);
    MCFidelite.Enabled := (NumeroCarteInterne <> '');
  end;
  Ok := (GP_TIERS.Text <> FODonneClientDefaut);
  MCSoldeClient.Enabled := Ok;
  MCContactClient.Enabled := Ok;
  MCArticleClient.Enabled := Ok;
  MCHistoClient.Enabled := Ok;
end;

procedure TFTicketFO.IncidenceTauxDate;
var OldTaux: Double;
begin
  if GeneCharge then Exit;
  if PieceContainer.Action = taConsult then Exit;
  if ((PieceContainer.Action = taModif) and (not PieceContainer.DuplicPiece) and (not PieceContainer.TransfoPiece)) then Exit;
  OldTaux := PieceContainer.DEV.Taux;
  GP_DEVISEChange(nil);
  if Arrondi(OldTaux - PieceContainer.DEV.Taux, 9) <> 0 then
  begin
    ChangeTauxAcomptes(PieceContainer);
    PutValueDetail(TOBPiece, 'GP_RECALCULER', 'X');
    CalculeLaSaisie(-1, -1, False);
    ChangeTauxEches(PieceContainer);
  end;
end;

procedure TFTicketFO.IncidenceRefs;
var TOBL: TOB;
  i, iArt: integer;
  RefA: string;
begin
  if PieceContainer.Action = taConsult then Exit;
  if TOBPiece.Detail.Count <= 0 then Exit;
  iArt := ChampToNum('GL_CODEARTICLE');
  for i := 0 to TOBPiece.Detail.Count - 1 do
  begin
    TOBL := TOBPiece.Detail[i];
    if TOBL = nil then Break;
    RefA := TOBL.GetValeur(iArt);
    if RefA = '' then Continue;
    if TOBL.GetValue('GL_REFARTSAISIE') = TOBL.GetValue('GL_REFARTTIERS') then
    begin
      TOBL.PutValue('GL_REFARTSAISIE', RefA);
      TOBL.PutValue('GL_REFARTTIERS', '');
      GS.Cells[SG_RefArt, i + 1] := RefA;
      AfficheLaLigne(i + 1);
    end;
  end;
end;

function TFTicketFO.IncidenceTarif: boolean;
var St: string;
  TOBL, TOBA: TOB;
  i: integer;
begin
  Result := False;
  if TOBPiece.Detail.Count <= 0 then Exit;
  if TOBArticles.Detail.Count <= 0 then Exit;
  if GetInfoParPiece(PieceContainer.NewNature, 'GPP_CONDITIONTARIF') <> 'X' then Exit;
  {$IFDEF FOS5}
  St := '1'; // Recalcul des prix de base avec conservation des remises
  {$ELSE}
  St := AGLLanceFiche('GC', 'GCOPTIONTARIF', '', '', '');
  {$ENDIF}
  if ((St = '') or (St = '0')) then Exit;
  if (not GetParamSoc('SO_PREFSYSTTARIF')) or (PGIAsk('Confirmez vous la mise à jour des tarifs sur toutes les lignes de la pièce ?', 'Système tarifaire') =     mrYes) then
  begin
    for i := 0 to TOBPiece.Detail.Count - 1 do
    begin
      TOBA := FindTOBArtRow(TOBPiece, TOBArticles, i + 1);
      if TOBA = nil then Continue;
      TOBL := TOBPiece.Detail[i];
      if ((GS.NbSelected = 0) or (GS.IsSelected(i + 1))) then
      begin
	      TarifVersLigne(TobA, TobTiers, TobL, nil, TobPiece, TobTarif, True, (St = '2'), PieceContainer.DEV);
        TOBL.putvalue('GL_RECALCULER', 'X');
        AfficheLaLigne(i + 1);
      end;
    end;
  end;
  TOBPiece.PutValue('GP_RECALCULER', 'X');
  CalculeLaSaisie(-1, -1, False);
  GS.ClearSelected;
  Result := True;
end;

procedure TFTicketFO.AfficheEuro;
var ModeOppose: boolean;
begin
  ModeOppose := False;
  if VH^.TenueEuro then
  begin
    if PopD.Items[0].Checked then ModeOppose := False else
      if PopD.Items[1].Checked then ModeOppose := True else
      if PopD.Items[2].Checked then ModeOppose := False;
  end else
  begin
    if PopD.Items[0].Checked then ModeOppose := False else
      if PopD.Items[1].Checked then ModeOppose := True else
      if PopD.Items[2].Checked then ModeOppose := False;
  end;
  GP_SAISIECONTRE.Checked := ModeOppose;
  GP_DEVISE.Enabled := PopD.Items[2].Checked;
  if ModeOppose then
  begin
    TOBPiece.PutValue('GP_SAISIECONTRE', 'X');
    PieceContainer.DEV.Decimale := V_PGI.OkDecE;
    if not VH^.TenueEuro then PieceContainer.DEV.Symbole := 'E';
    ChangeFormatDevise(Self, PieceContainer.DEV.Decimale, PieceContainer.DEV.Symbole);
  end else
  begin
    TOBPiece.PutValue('GP_SAISIECONTRE', '-');
    if PieceContainer.DEV.Code = V_PGI.DevisePivot then
    begin
      PieceContainer.DEV.Decimale := V_PGI.OkDecV;
      PieceContainer.DEV.Symbole := V_PGI.SymbolePivot;
      if PieceContainer.DEV.Code = FODonneCodeEuro then PieceContainer.DEV.Symbole := SIGLEEURO;
      ChangeFormatDevise(Self, PieceContainer.DEV.Decimale, PieceContainer.DEV.Symbole);
    end;
  end;
end;

procedure TFTicketFO.IncidenceTiers;
var CodeTiers, TypePrix: string;
  EnHt, OkE: boolean;
  {$IFNDEF FOS5}
  CodeD, TypeRel: string;
  {$ENDIF}
begin
  if GeneCharge then Exit;
  if ((PieceContainer.Action <> taCreat) and (not PieceContainer.DuplicPiece)) then Exit;
  CodeTiers := TOBTiers.GetValue('T_TIERS');
  GP_ESCOMPTE.Value := TOBTiers.GetValue('T_ESCOMPTE');
  PutValueDetail(TOBPiece, 'GP_ESCOMPTE', GP_ESCOMPTE.Value);
  if ctxScot in V_PGI.PGIContexte then
  begin
    GP_REMISEPIED.Value := TOBTiers.GetValue('T_REMISE');
    PutValueDetail(TOBPiece, 'GP_REMISEPIED', GP_REMISEPIED.Value);
  end;
  GP_MODEREGLE.Value := TOBTiers.GetValue('T_MODEREGLE');
  {$IFNDEF FOS5}
  GP_REGIMETAXE.Value := TOBTiers.GetValue('T_REGIMETVA');
  {$ENDIF}
  if GP_REPRESENTANT.Text = '' then
  begin
    GP_REPRESENTANT.Text := FOPreAffecteVendeur(TOBTiers);
    if (GP_REPRESENTANT.Text <> '') and (not FORepresentantExiste(GP_REPRESENTANT.Text, TypeCom, TOBComms)) then GP_REPRESENTANT.Text := '';
  end;
  GereCommercialEnabled;
  {**if ((TOBPiece.Detail.Count<=0) and (PieceContainer.Action=taCreat)) then
     BEGIN
     TypePrix:=VH_GC.TOBPCaisse.GetValue('GPK_APPELPRIXTIC') ;
     if TypePrix='HT' then GP_FACTUREHT.Checked:=True else
      if TypePrix='TTC' then GP_FACTUREHT.Checked:=False else
       GP_FACTUREHT.Checked:=(TOBTiers.GetValue('T_FACTUREHT')='X') ; **}
  OkE := ((TOBPiece.Detail.Count <= 0) and (PieceContainer.Action = taCreat));
  if PieceContainer.Action = taCreat then
  begin
    TypePrix := VH_GC.TOBPCaisse.GetValue('GPK_APPELPRIXTIC');
    if TypePrix = 'HT' then EnHt := True else
      if TypePrix = 'TTC' then EnHt := False else
      EnHt := (TOBTiers.GetValue('T_FACTUREHT') = 'X');
    if EnHt <> GP_FACTUREHT.Checked then
    begin
      if (TOBPiece.Detail.Count <= 0) or (not DejaSaisie) then
      begin
        GP_FACTUREHT.Checked := EnHt;
        OkE := True;
      end else
      begin
        if GP_FACTUREHT.Checked then
          HPiece.Execute(54, Caption, '')
        else
          HPiece.Execute(53, Caption, '');
      end;
    end;
    {$IFNDEF FOS5}
    if TOBPiece.Detail.Count <= 0 then
    begin
      CodeD := TOBTiers.GetValue('T_DEVISE');
      if CodeD <> '' then GP_DEVISE.Value := CodeD;
    end;
    {$ENDIF}
  end;
  if OkE then EtudieColsListe;

  TOBPiece.PutValue('GP_TVAENCAISSEMENT', PositionneExige(TOBTiers));
  TiersVersPiece(TOBTiers, TOBPiece);
  if GereAdresse then TiersVersAdresses(PieceContainer);
  {$IFDEF TAXEUS}
  TaxeVersPiece(TOBPiece, TOBTiers);
  IncidenceTaxe;
  {$ENDIF}
  if ((PieceContainer.Action = taCreat) and (not PieceContainer.DuplicPiece)) then
  begin
    {$IFDEF FOS5}
    FOAffecteFactureHT(TOBPiece);
    {$ELSE}
    OkE := (TOBTiers.GetValue('T_EURODEFAUT') = 'X');
    if OkE then
    begin
      if VH^.TenueEuro then PopD.Items[0].Checked := True else PopD.Items[1].Checked := True;
      GP_DEVISE.Value := V_PGI.DevisePivot;
    end else if GP_DEVISE.Value = V_PGI.DevisePivot then
    begin
      if VH^.TenueEuro then PopD.Items[1].Checked := True else PopD.Items[0].Checked := True;
      if not GeneCharge then AfficheEuro;
    end;
    {$ENDIF}
  end;
  {Ecran}
  TOBPiece.GetEcran(Self);
end;

{==============================================================================================}
{=============================== Manipulation des divers ======================================}
{==============================================================================================}

procedure TFTicketFO.TraiteLesDivers(ACol, ARow: integer);
var TOBL: TOB;
  VV: Variant;
  ColName, St, Typ: string;
  j: integer;
begin
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  ColName := IdentCols[ACol].ColName;
  if ColName = '' then Exit;
  j := TOBL.GetNumChamp(ColName);
  if j < 0 then Exit;
  Typ := IdentCols[ACol].ColTyp;
  if Typ = 'COMBO' then St := GS.CellValues[ACol, ARow] else St := GS.Cells[ACol, ARow];
  if (Typ = 'INTEGER') or (Typ = 'SMALLINT') then Vv := ValeurI(St) else
    if (Typ = 'DOUBLE') or (Typ = 'RATE') then Vv := Valeur(St) else
    if Typ = 'DATE' then Vv := StrToDate(St) else
    if Typ = 'BOOLEAN' then
  begin
    if Uppercase(St) <> 'X' then St := '-';
    Vv := Uppercase(St);
  end
  else Vv := St;
  TOBL.PutValeur(j, Vv);
end;

procedure TFTicketFO.TraiteLibelle(ARow: integer);
var TOBL: TOB;
begin
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  TOBL.PutValue('GL_LIBELLE', GS.Cells[SG_Lib, ARow]);
end;

procedure TFTicketFO.TraitePrix(ARow: integer; var Cancel: Boolean);
var TOBL: TOB;
  PUDEV: double;
  sFonctionnalite: string;
begin
  PUDEV := Valeur(GS.Cells[SG_Px, ARow]);
  if not FOPrixUnitaireAutorise(PUDEV) then
  begin
    HErr.Execute(28, Caption, '');
    Cancel := True;
    Exit;
  end;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  if GP_FACTUREHT.Checked then TOBL.PutValue('GL_PUHTDEV', PUDEV)
  else TOBL.PutValue('GL_PUTTCDEV', PUDEV);
  if GetParamSoc('SO_PREFSYSTTARIF') then
  begin
    sFonctionnalite := RechercheFonctionnalite(TobL.GetValue('GL_NATUREPIECEG'));
    AnnulationTobLigneTarif(TobL, TobLigneTarif, sFonctionnalite, '17'); //Annulation du prix net
    MajTobLigneTarifFromTobLigne(TobL, TobLigneTarif, sFonctionnalite, '14', 'PRIXBRUT', PUDEV);
  end;

  TOBL.PutValue('GL_RECALCULER', 'X');
  TOBPIece.putvalue('GP_RECALCULER', 'X');
end;

procedure TFTicketFO.TraitePrixNet(ACol, ARow: integer; var Cancel: Boolean);
var TOBL: TOB;
  QteF, PUDEV, PQ, MontantBrutL, MontantNetL, MtRemL, RemL: Double;
begin
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  QteF := TOBL.GetValue('GL_QTEFACT');
  if QteF = 0 then Exit;
  if not FOPrixUnitaireAutorise(Valeur(GS.Cells[ACol, ARow])) then
  begin
    HErr.Execute(28, Caption, '');
    Cancel := True;
    Exit;
  end;
  // cas des articles sans prix
  if AppliquePUBrut(TOBL, Valeur(GS.Cells[ACol, ARow]), ACol, ARow) then
  begin
    PUDEV := Valeur(GS.Cells[ACol, ARow]);
    if GP_FACTUREHT.Checked then
      TOBL.PutValue('GL_PUHTDEV', PUDEV)
    else TOBL.PutValue('GL_PUTTCDEV', PUDEV);
    // Affichage du prix unitaire de l'article
    AffichePUBrut(TOBL);
  end;
  // Calcul du montant brut
  if GP_FACTUREHT.Checked then PUDEV := TOBL.GetValue('GL_PUHTDEV') else PUDEV := TOBL.GetValue('GL_PUTTCDEV'); //if PUDEV=0 then Exit ;
  PQ := TOBL.GetValue('GL_PRIXPOURQTE');
  if PQ <= 0 then
  begin
    PQ := 1.0;
    TOBL.PutValue('GL_PRIXPOURQTE', PQ);
  end;
  MontantBrutL := Arrondi(PUDEV * QteF / PQ, 6);
  // Calcul du montant net
  MontantNetL := Arrondi((Valeur(GS.Cells[ACol, ARow]) * QteF) / PQ, 6);
  if MontantBrutL <> 0 then
  begin
    // Calcul du montant de la remise avec application de l'arrondi
    MtRemL := MontantBrutL - MontantNetL;
    if MtRemL <> 0 then MtRemL := FOArrondiRemise(MtRemL, MontantBrutL);
    // Calcul du pourcentage de remise
    RemL := Arrondi(((MtRemL / MontantBrutL)) * 100.0, 6);
  end else
  begin
    MtRemL := 0;
    RemL := 0;
  end;
  RemiseCorrecte(RemL, TOBL.GetValue('GL_TYPEREMISE'), False, Cancel);
  if Cancel then Exit;
  TOBL.PutValue('GL_VALEURREMDEV', MtRemL);
  TOBL.PutValue('GL_REMISELIGNE', RemL);
  if RemL = 0 then
  begin
    TOBL.PutValue('GL_VALEURREMDEV', 0);
    TOBL.PutValue('GL_TOTREMLIGNEDEV', 0);
    TOBL.PutValue('GL_TYPEREMISE', '');
    if SG_RV > 0 then GS.Cells[SG_RV, ARow] := '';
    if SG_RL > 0 then GS.Cells[SG_RL, ARow] := '';
  end;
  TOBL.PutValue('GL_RECALCULER', 'X');
  TOBPIece.putvalue('GP_RECALCULER', 'X');
end;

procedure TFTicketFO.TraiteRemise(ARow: integer; var Cancel: Boolean);
var TOBL: TOB;
  RemL: Double;
  QteF, PUDEV, PQ, MontantBrutL, MtRemL: Double;
begin
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  RemL := Valeur(GS.Cells[SG_Rem, ARow]);
  if TOBL.GetValue('GL_TYPELIGNE') = 'TOT' then
  begin
    RemiseSousTotalCorrecte(ARow, RemL, 0, 0, '', Cancel);
    GS.Cells[SG_Rem, ARow] := '';
    Exit;
  end;
  if FODonneArrondiRemise = '' then
  begin
    RemiseCorrecte(RemL, TOBL.GetValue('GL_TYPEREMISE'), False, Cancel);
    if Cancel then Exit;
    TOBL.PutValue('GL_REMISELIGNE', RemL);
    TOBL.PutValue('GL_VALEURREMDEV', 0);
  end else
  begin
    QteF := TOBL.GetValue('GL_QTEFACT');
    if QteF = 0 then Exit;
    // Calcul du montant brut
    if GP_FACTUREHT.Checked then PUDEV := TOBL.GetValue('GL_PUHTDEV') else PUDEV := TOBL.GetValue('GL_PUTTCDEV');
    if PUDEV = 0 then Exit;
    PQ := TOBL.GetValue('GL_PRIXPOURQTE');
    if PQ <= 0 then
    begin
      PQ := 1.0;
      TOBL.PutValue('GL_PRIXPOURQTE', PQ);
    end;
    MontantBrutL := Arrondi(PUDEV * QteF / PQ, 6);
    // Calcul du montant de la remise avec application de l'arrondi
    MtRemL := Arrondi(MontantBrutL * (RemL / 100), 6);
    if MtRemL <> 0 then MtRemL := FOArrondiRemise(MtRemL, MontantBrutL);
    TOBL.PutValue('GL_VALEURREMDEV', MtRemL);
    // Calcul et vérification du pourcentage de remise
    if MontantBrutL <> 0 then RemL := Arrondi(100.0 * MtRemL / MontantBrutL, 6) else RemL := 0;
    RemiseCorrecte(RemL, TOBL.GetValue('GL_TYPEREMISE'), False, Cancel);
    if Cancel then Exit;
    TOBL.PutValue('GL_REMISELIGNE', RemL);
  end;
  if RemL = 0 then TOBL.PutValue('GL_TYPEREMISE', '');
  TOBL.PutValue('GL_RECALCULER', 'X');
  TOBPIece.putvalue('GP_RECALCULER', 'X');
end;

procedure TFTicketFO.TraiteMontantRemise(ACol, ARow: integer; var Cancel: Boolean);
var TOBL: TOB;
  MontantL, RemL: Double;
  QteF, PUDEV, PQ, MtRemL: Double;
begin
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  if TOBL.GetValue('GL_TYPELIGNE') = 'TOT' then
  begin
    MtRemL := Valeur(GS.Cells[ACol, ARow]);
    RemiseSousTotalCorrecte(ARow, 0, MtRemL, 0, '', Cancel);
    Exit;
  end;
  // Calcul du montant brut
  QteF := TOBL.GetValue('GL_QTEFACT');
  if QteF = 0 then Exit;
  if GP_FACTUREHT.Checked then PUDEV := TOBL.GetValue('GL_PUHTDEV') else PUDEV := TOBL.GetValue('GL_PUTTCDEV'); //if PUDEV=0 then Exit ;
  PQ := TOBL.GetValue('GL_PRIXPOURQTE');
  if PQ <= 0 then
  begin
    PQ := 1.0;
    TOBL.PutValue('GL_PRIXPOURQTE', PQ);
  end;
  MontantL := Arrondi(PUDEV * QteF / PQ, 6);
  // Application de l'arrondi
  MtRemL := Valeur(GS.Cells[ACol, ARow]);
  if MtRemL <> 0 then MtRemL := FOArrondiRemise(MtRemL, MontantL);
  TOBL.PutValue('GL_VALEURREMDEV', MtRemL);
  // Calcul et vérification du pourcentage de remise
  if MontantL <> 0 then RemL := Arrondi(100.0 * MtRemL / MontantL, 6) else RemL := 0;
  RemiseCorrecte(RemL, TOBL.GetValue('GL_TYPEREMISE'), False, Cancel);
  if Cancel then Exit;
  TOBL.PutValue('GL_REMISELIGNE', RemL);
  if RemL = 0 then
  begin
    TOBL.PutValue('GL_VALEURREMDEV', 0);
    TOBL.PutValue('GL_TOTREMLIGNEDEV', 0);
    TOBL.PutValue('GL_TYPEREMISE', '');
    if SG_RV > 0 then GS.Cells[SG_RV, ARow] := '';
    if SG_RL > 0 then GS.Cells[SG_RL, ARow] := '';
  end;
  TOBL.PutValue('GL_RECALCULER', 'X');
  TOBPIece.putvalue('GP_RECALCULER', 'X');
end;

procedure TFTicketFO.TraiteMontantNetLigne(ACol, ARow: integer; var Cancel: Boolean);
var TOBL: TOB;
  QteF, PUDEV, PQ, MontantBrutL, MontantNetL, MtRemL, RemL: Double;
begin
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  if TOBL.GetValue('GL_TYPELIGNE') = 'TOT' then
  begin
    MontantNetL := Valeur(GS.Cells[ACol, ARow]);
    RemiseSousTotalCorrecte(ARow, 0, 0, MontantNetL, '', Cancel);
    Exit;
  end;
  QteF := TOBL.GetValue('GL_QTEFACT');
  if QteF = 0 then Exit;
  // Calcul du montant brut
  if GP_FACTUREHT.Checked then PUDEV := TOBL.GetValue('GL_PUHTDEV') else PUDEV := TOBL.GetValue('GL_PUTTCDEV');
  if (PUDEV = 0) and (not (TOBL.FieldExists('MODIFPU')) or (TOBL.GetValue('MODIFPU') = '-')) then Exit;
  PQ := TOBL.GetValue('GL_PRIXPOURQTE');
  if PQ <= 0 then
  begin
    PQ := 1.0;
    TOBL.PutValue('GL_PRIXPOURQTE', PQ);
  end;
  // cas des articles sans prix
  if AppliquePUBrut(TOBL, Arrondi(Valeur(GS.Cells[ACol, ARow]) * PQ / QteF, 6), ACol, ARow) then
  begin
    MontantBrutL := Valeur(GS.Cells[ACol, ARow]);
    PUDEV := Arrondi(MontantBrutL * PQ / QteF, 6);
    if GP_FACTUREHT.Checked then
      TOBL.PutValue('GL_PUHTDEV', PUDEV)
    else TOBL.PutValue('GL_PUTTCDEV', PUDEV);
    // Affichage du prix unitaire de l'article
    AffichePUBrut(TOBL);
  end else
  begin
    MontantBrutL := Arrondi(PUDEV * QteF / PQ, 6);
  end;
  // Calcul du montant de la remise avec application de l'arrondi
  MtRemL := MontantBrutL - Valeur(GS.Cells[ACol, ARow]);
  if MtRemL <> 0 then MtRemL := FOArrondiRemise(MtRemL, MontantBrutL);
  // Calcul du montant net
  MontantNetL := MontantBrutL - MtRemL;
  // Calcul du pourcentage de remise
  if MontantBrutL <> 0 then RemL := Arrondi((1.0 - (MontantNetL / MontantBrutL)) * 100.0, 6) else RemL := 0;
  RemiseCorrecte(RemL, TOBL.GetValue('GL_TYPEREMISE'), False, Cancel);
  if Cancel then Exit;
  TOBL.PutValue('GL_VALEURREMDEV', MtRemL);
  TOBL.PutValue('GL_REMISELIGNE', RemL);
  if RemL = 0 then
  begin
    TOBL.PutValue('GL_VALEURREMDEV', 0);
    TOBL.PutValue('GL_TOTREMLIGNEDEV', 0);
    TOBL.PutValue('GL_TYPEREMISE', '');
    if SG_RV > 0 then GS.Cells[SG_RV, ARow] := '';
    if SG_RL > 0 then GS.Cells[SG_RL, ARow] := '';
  end;
  TOBL.PutValue('GL_RECALCULER', 'X');
  TOBPIece.putvalue('GP_RECALCULER', 'X');
end;
{$IFDEF CHR}
Procedure TFTicketFO.GereChampsChr( Arow : integer;var cancel:boolean) ;
begin
if ((GS.Cells[SG_RefArt,Arow] <> '')
     and (GS.Cells[SG_Folio,Arow] = '0')) then
  begin
  GS.Cells[SG_Folio, Arow]:='1';
  GetTOBLigne(TOBPiece,Arow).PutValue('GL_NOFOLIO', 1);
  end;
end;

Procedure TFTicketFO.TraiteFolio(ACol,ARow : integer; Var cancel: boolean ) ;
var numf:Integer;
  TOBL : TOB;
Begin
  NumF:=0;
  TOBL:=GetTOBLigne(TOBPiece,ARow) ; if TOBL=Nil then Exit ;
  NumF:=StrToInt(GS.Cells[Sg_Folio,AROW]) ;

  If (NumF > 3) Then
  begin
     Cancel := True;
  End
  else
  Begin
     TOBL.PutValue('GL_NOFOLIO',NumF) ;
     Cancel := False;
  End;
End ;

{$ENDIF} // FIN AJOUT CHR
function TFTicketFO.AppliquePUBrut(TOBL: TOB; NewPU: Double; ACol, ARow: Integer): Boolean;
var TOBA: TOB;
  Rep: Word;
  PUDEV: Double;
begin
  Rep := mrNo;
  if GP_FACTUREHT.Checked then PUDEV := TOBL.GetValue('GL_PUHTDEV')
  else PUDEV := TOBL.GetValue('GL_PUTTCDEV');
  if (PUDEV <> 0) and (TOBL.GetValue('GL_PUTTCBASE') = 0) then
  begin
    // cas d'un article sans prix
    TOBA := FindTOBArtRow(TOBPiece, TOBArticles, ARow);
    if TOBA.GetValue('GA_REMISELIGNE') = '-' then Rep := mrYes else
      if (ACol = SG_Montant) and (FOGetParamCaisse('GPK_AFFPRIXTIC') <> '001') then Rep := mrNo else
      if PUDEV <> NewPU then Rep := HPiece.Execute(47, Caption, '');
  end;
  if Rep = mrYes then PUDEV := 0;
  Result := (PUDEV = 0);
end;

{==============================================================================================}
{=============================== Manipulation des Qtés ========================================}
{==============================================================================================}

procedure TFTicketFO.DeflagTarif(ARow: integer);
var TOBL: TOB;
begin
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  TOBL.SetBoolean ('RECALCULTARIF', TobL.GetBoolean ('GL_BLOQUETARIF'));
end;

//function TFTicketFO.TraiteRupture(ARow: integer): boolean;
function TFTicketFO.TraiteRupture(ARow: integer; bParle : boolean = True): boolean; // eQualité 11973
var TOBL, TOBA, TOBD, TOBLiee: TOB;
  RefU, Depot, OldNat, ColPlus, ColMoins, ArtSub, sComp: string;
  QteSais, QteDisp, RatioVA, QteAdd, RatioAdd: Double;
  AddQ, AnnuArt: Boolean;
begin
  Result := True;
  if PieceContainer.Action = taConsult then Exit;
  if ((CalcRupt = '') or (CalcRupt = 'AUC')) then Exit;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  RefU := TOBL.GetValue('GL_ARTICLE');
  if RefU = '' then Exit;
  QteSais := TOBL.GetValue('GL_QTESTOCK');
  if QteSais <= 0 then Exit;
  if TOBL.GetValue('GL_TENUESTOCK') <> 'X' then Exit;
  Depot := TOBL.GetValue('GL_DEPOT');
  if Depot = '' then Exit;
  TOBA := FindTOBArtRow(TOBPiece, TOBArticles, ARow);
  if TOBA = nil then Exit;
  TOBD := TOBA.FindFirst(['GQ_DEPOT'], [Depot], False);
  RatioAdd := 1;
  RatioVA := GetRatio(TOBL, nil, trsStock);
  if RatioVA = 0 then RatioVA := 1;
  if TOBD = nil then
  begin
    QteDisp := 0;
    QteAdd := 0;
  end else
  begin
    QteDisp := TOBD.GetValue('GQ_PHYSIQUE');
    AddQ := False;
    QteAdd := 0;
    if CalcRupt = 'DIS' then QteDisp := QteDisp - TOBD.GetValue('GQ_RESERVECLI') - TOBD.GetValue('GQ_PREPACLI');
    // Cas particulier rappel de pièce
    if ((PieceContainer.Action = taModif) and (not PieceContainer.DuplicPiece)) then
    begin
      OldNat := TOBPiece_O.GetValue('GP_NATUREPIECEG');
      ColPlus := GetInfoParPiece(OldNat, 'GPP_QTEPLUS');
      ColMoins := GetInfoParPiece(OldNat, 'GPP_QTEMOINS');
      if ((Pos('PHY', ColPlus) > 0) or (Pos('PHY', ColMoins) > 0)) then AddQ := True;
      if ((CalcRupt = 'DIS') or (not PieceContainer.TransfoPiece)) then
      begin
        if ((Pos('RC', ColPlus) > 0) or (Pos('RC', ColMoins) > 0)) then AddQ := True;
        if ((Pos('PRE', ColPlus) > 0) or (Pos('PRE', ColMoins) > 0)) then AddQ := True;
      end;
      if AddQ then
      begin
        if PieceContainer.TransfoPiece then
        begin
          TOBLiee := GetTOBPrec(TOBL, TOBPiece_O);
          if TOBLiee <> nil then
          begin
            QteAdd := TOBLiee.GetValue('GL_QTESTOCK');
            RatioAdd := GetRatio(TOBLiee, nil, trsStock);
            if RatioAdd = 0 then RatioAdd := 1;
          end;
        end else
        begin
          if TOBL.FieldExists('QTEORIG') then QteAdd := TOBL.GetValue('QTEORIG');
          RatioAdd := RatioVA;
        end;
      end;
    end;
  end;
  QteSais := (QteSais / RatioVA) - (QteAdd / RatioAdd);
  if QteSais <= QteDisp then Exit;
  ArtSub := Trim(Copy(TOBA.GetValue('GA_SUBSTITUTION'), 1, 18));
  if ArtSub <> '' then sComp := HTitres.Mess[17] + ' ' + ArtSub + ')' else sComp := '';
  AnnuArt := True;
  if ForceRupt then
  begin
    if HPiece.Execute(10, caption, sComp) = mrYes then AnnuArt := False;
  end else
  begin
    HPiece.Execute(11, caption, sComp);
  end;
  if AnnuArt then Result := False;
end;

function TFTicketFO.TraiteQte(ACol, ARow: integer; bParle : boolean = true): Boolean;
var TOBL, TOBA, TOBLiee: TOB;
  NewQte, Qte, PCB, X, OldQte, Ecart: Double;
  Neg, OkCond: boolean;
  RefUnique: string;
begin
  Qte := Valeur(GS.Cells[ACol, ARow]);
  if not FOQuantiteAutorise(Qte) then
  begin
    HErr.Execute(27, Caption, '');
    Result := False;
    Exit;
  end;
  if (Qte < 0) and (not FOJaiLeDroit(8)) then
  begin
    Result := False;
    Exit;
  end;
  Result := True;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  RefUnique := TOBL.GetValue('GL_ARTICLE');
  // Conditionnement
  OkCond := AdapteCond(TOBL, TOBConds, Qte, GS, VenteAchat);
  if OkCond then
  begin
    GS.Cells[ACol, ARow] := StrF00(Qte, 10);
    FormateZoneSaisie(ACol, ARow);
  end else
  begin
    // Prix par combien
    PCB := TOBL.GetValue('GL_PCB');
    if ((PCB <> 0) and (Qte <> 0) and (VenteAchat = 'VEN')) then
    begin
      Neg := (Qte < 0);
      Qte := Abs(Qte);
      X := Arrondi(Qte / PCB, 9);
      if Arrondi(X - Trunc(X), 6) <> 0 then
      begin
        X := Arrondi(X + 0.5, 0);
        X := X * PCB;
        if Neg then X := -X;
        Qte := X;
        GS.Cells[ACol, ARow] := StrF00(Qte, 10);
        FormateZoneSaisie(ACol, ARow);
        TOBPIece.putvalue('GP_RECALCULER', 'X');
        TOBL.PutValue('GL_RECALCULER', 'X');
      end;
    end;
  end;
  if ACol = SG_QF then
  begin
    TOBL.PutValue('GL_QTEFACT', Valeur(GS.Cells[ACol, ARow]));
    if SG_QS < 0 then
    begin
      TOBL.PutValue('GL_QTESTOCK', Valeur(GS.Cells[ACol, ARow]));
      // MODIF LM 16/02/2004 Bug Qtereste à 1 !!!!!!!!!!!!!
      TOBL.PutValue('GL_QTERESTE', Valeur(GS.Cells[ACol, ARow]));
    end;
    if (TOBL.GetValue('RECALCULTARIF') = 'X') and (not GetParamSoc('SO_PREFSYSTTARIF')) then
    begin
      TOBA := FindTOBArtRow(TOBPiece, TOBArticles, ARow);
      TarifVersLigne(TobA, TobTiers, TobL, nil, TobPiece, TobTarif, True, True, PieceContainer.DEV);
      AfficheLaLigne(ARow, True);
    end;
    TOBPIece.putvalue('GP_RECALCULER', 'X');
    TOBL.PutValue('GL_RECALCULER', 'X');
  end else
  begin
    TOBL.PutValue('GL_QTESTOCK', Valeur(GS.Cells[ACol, ARow]));
    if SG_QF < 0 then
    begin
      TOBL.PutValue('GL_QTEFACT', Valeur(GS.Cells[ACol, ARow]));
      // MODIF LM 16/02/2004 Bug Qtereste à 1 !!!!!!!!!!!!!
      TOBL.PutValue('GL_QTERESTE', Valeur(GS.Cells[ACol, ARow]));
    end;
    TOBPIece.putvalue('GP_RECALCULER', 'X');
    TOBL.PutValue('GL_RECALCULER', 'X');
  end;
  if ((ACol = SG_QF) or (SG_QF < 0)) then TOBL.PutValue('RECALCULTARIF', '-');
  if ReliquatTransfo then
  begin
    TOBLiee := GetTOBPrec(TOBL, TOBPiece_O);
    if TOBLiee <> nil then
    begin
      NewQte := TOBL.GetValue('GL_QTESTOCK');
      OldQte := TOBLiee.GetValue('GL_QTESTOCK');
      if OldQte > NewQte then Ecart := OldQte - NewQte else Ecart := 0;
      TOBL.PutValue('GL_QTERELIQUAT', Ecart);
      TOBPIece.putvalue('GP_RECALCULER', 'X');
      TOBL.PutValue('GL_RECALCULER', 'X');
    end;
  end;
  if not TraiteRupture(ARow) then
  begin
    if SG_QF >= 0 then GS.Cells[SG_QF, ARow] := '';
    if SG_QS >= 0 then GS.Cells[SG_QS, ARow] := '';
    Result := False;
  end;
  if Result then
  begin
    TOBL.PutValue('QTECHANGE', 'X');
//    FOMajQte(TOBL, (TOBPiece.GetValue('GP_FACTUREHT') = 'X'));
  end;
end;

procedure TFTicketFO.RetourneArticle(ARow: Integer);
var Qte, NewQte: Double;
begin
  if GetCodeArtUnique(TOBPiece, ARow) <> '' then
  begin
    // cas d'une ligne existante : on inverse la quantité
    if SG_QS > 0 then
    begin
      Qte := Valeur(GS.Cells[SG_QS, ARow]);
      NewQte := (Qte * -1);
      GS.Cells[SG_QS, ARow] := StrF00(NewQte, V_PGI.OkDecQ);
      if TraiteQte(SG_QS, ARow) then CalculeLaSaisie(SG_QS, ARow, False)
      else GS.Cells[SG_QS, ARow] := StrF00(Qte, V_PGI.OkDecQ);
    end;
    if SG_QF > 0 then
    begin
      Qte := Valeur(GS.Cells[SG_QF, ARow]);
      NewQte := (Qte * -1);
      GS.Cells[SG_QF, ARow] := StrF00(NewQte, V_PGI.OkDecQ);
      if TraiteQte(SG_QF, ARow) then CalculeLaSaisie(SG_QF, ARow, False)
      else GS.Cells[SG_QF, ARow] := StrF00(Qte, V_PGI.OkDecQ);
    end;
  end;
end;

{$IFDEF MODE}
procedure TFTicketFO.RechercheArticleImage;
var
  Stg: string;
begin
  if PieceContainer.Action = taConsult then Exit;
  if SG_RefArt = -1 then Exit;
  if IndiqueOuSaisie <> saLigne then Exit;
  // Positionnement dans le Grid
  if not AllerAuControle(saLigne, -1, SG_RefArt, False, True) then Exit;
  // Le client doit être renseiné
  if TOBTiers.GetValue('T_TIERS') = '' then
  begin
    if GP_TIERS.CanFocus then GP_TIERS.SetFocus else GotoEntete;
    Exit;
  end;
  Stg := FOLanceFicheRechArtImage('', '', '');
  if Stg <> '' then ToucheArticle('ART', Stg, '', 0, 0);
end;
{$ENDIF}

{==============================================================================================}
{====================================== DIMENSIONS ============================================}
{==============================================================================================}
// modif 02/08/2001

procedure TFTicketFO.RemplirTOBDim(CodeArticle: string; Ligne: Integer);
var iLigne, jDim: integer;
  TOBL, TOBD: TOB;
  Q: TQuery;
  QteDejaSaisi, ReliqDejaSaisi: Double;
begin
  TOBDim.ClearDetail;
  ANCIEN_TOBDimDetailCount := 0;
  Q := OpenSQL('Select GA_Article from Article where GA_CodeArticle="' + CodeArticle + '" AND GA_STATUTART="DIM" order by GA_ARTICLE', True);
  while not Q.EOF do
  begin
    TOBD := TOB.Create('', TOBDim, -1);
    QteDejaSaisi := CalcQteDejaSaisie(TOBPiece, Q.FindField('GA_ARTICLE').AsString);
    ReliqDejaSaisi := CalcQteDejaSaisie(TOBPiece_O, Q.FindField('GA_ARTICLE').AsString);
    InitDim(TOBD, Q.FindField('GA_ARTICLE').AsString, QteDejaSaisi, ReliqDejaSaisi);
    Q.Next;
  end;
  ferme(Q);
  if not NewLigneDim then
  begin
    for jDim := 0 to TOBDim.Detail.Count - 1 do
    begin
      TOBD := TOBDim.Detail[jDim];
      iLigne := Ligne;
      TOBL := TOBPiece.Detail[iLigne];
      while TOBL.GetValue('GL_CODEARTICLE') = CodeArticle do
      begin
        if TOBL.GetValue('GL_ARTICLE') = TOBD.GetValue('GA_ARTICLE') then
        begin
          QteDejaSaisi := CalcQteDejaSaisie(TOBPiece, TOBD.GetValue('GA_ARTICLE'));
          ReliqDejaSaisi := CalcQteDejaSaisie(TOBPiece_O, TOBD.GetValue('GA_ARTICLE'));
          LigneVersDim(TOBL, TOBD, QteDejaSaisi, ReliqDejaSaisi);
          ANCIEN_TOBDimDetailCount := ANCIEN_TOBDimDetailCount + 1;
          Break;
        end else
        begin
          iLigne := iLigne + 1;
          if iLigne > TOBPiece.Detail.Count - 1 then break;
          TOBL := TOBPiece.Detail[iLigne];
        end;
      end;
    end;
  end;
  if not TOBDim.FieldExists('GP_DEVISE') then TOBDim.AddChampSup('GP_DEVISE', False);
  if not TOBDim.FieldExists('GP_TAUXDEV') then TOBDim.AddChampSup('GP_TAUXDEV', False);
  if not TOBDim.FieldExists('GP_SAISIECONTRE') then TOBDim.AddChampSup('GP_SAISIECONTRE', False);
  if not TOBDim.FieldExists('GP_FACTUREHT') then TOBDim.AddChampSup('GP_FACTUREHT', False);
  if not TOBDim.FieldExists('GP_VENTEACHAT') then TOBDim.AddChampSup('GP_VENTEACHAT', False);
  if not TOBDim.FieldExists('GP_NATUREPIECEG') then TOBDim.AddChampSup('GP_NATUREPIECEG', False);
  if not TOBDim.FieldExists('GP_DATEPIECE') then TOBDim.AddChampSup('GP_DATEPIECE', False); // recup tarif
  if not TOBDim.FieldExists('GP_TIERS') then TOBDim.AddChampSup('GP_TIERS', False); // recup tarif achat
  if not TOBDim.FieldExists('GP_TARIFTIERS') then TOBDim.AddChampSup('GP_TARIFTIERS', False); // recup tarif tiers
  TOBDim.PutValue('GP_DEVISE', TOBPiece.GetValue('GP_DEVISE'));
  TOBDim.PutValue('GP_TAUXDEV', TOBPiece.GetValue('GP_TAUXDEV'));
  TOBDim.PutValue('GP_SAISIECONTRE', TOBPiece.GetValue('GP_SAISIECONTRE'));
  TOBDim.PutValue('GP_FACTUREHT', TOBPiece.GetValue('GP_FACTUREHT'));
  TOBDim.PutValue('GP_VENTEACHAT', TOBPiece.GetValue('GP_VENTEACHAT'));
  TOBDim.PutValue('GP_NATUREPIECEG', TOBPiece.GetValue('GP_NATUREPIECEG'));
  TOBDim.PutValue('GP_DATEPIECE', TOBPiece.GetValue('GP_DATEPIECE'));
  TOBDim.PutValue('GP_TIERS', TOBPiece.GetValue('GP_TIERS'));
  TOBDim.PutValue('GP_TARIFTIERS', TOBPiece.GetValue('GP_TARIFTIERS'));
  TheTOB := TOBDim;
end;

procedure TFTicketFO.AppelleDim(ARow: integer);
var TypeDim, CodeArticle, CodeDepot: string;
  TOBL, TOBGEN: TOB;
  LigGen: integer;
begin
  NewLigneDim := False;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  TypeDim := TOBL.GetValue('GL_TYPEDIM');
  if ((TypeDim <> 'GEN') and (TypeDim <> 'DIM')) then Exit;
  if TypeDim = 'DIM' then
  begin
    CodeArticle := TOBL.GetValue('GL_CODEARTICLE');
    LigGen := -1;
  end
  else
  begin
    CodeArticle := TOBL.GetValue('GL_CODESDIM');
    LigGen := ARow;
  end;
  if CodeArticle = '' then Exit;
  RemplirTOBDim(CodeArticle, ARow);
  //if SelectMultiDimensions(CodeArticle,GS,TobPiece.GetValue('GP_ETABLISSEMENT'),PieceContainer.Action) then TOBDim:=TheTOB else TOBDim.ClearDetail ; //AC
  if (PieceContainer.Cledoc.NaturePiece = 'TRE') or (PieceContainer.Cledoc.NaturePiece = 'TRV') then
    CodeDepot := TOBPiece.GetValue('GP_DEPOTDEST')
  else CodeDepot := TOBPiece.GetValue('GP_ETABLISSEMENT');
  if SelectMultiDimensions(CodeArticle, GS, CodeDepot, PieceContainer.Action) then TOBDim := TheTOB else TOBDim.ClearDetail; //AC+JD
  TheTOB := nil;
  if LigGen < 0 then
  begin
    TOBGEN := TOBPiece.FindFirst(['GL_CODESDIM', 'GL_TYPEDIM'], [CodeArticle, 'GEN'], False);
    if TOBGen = nil then Exit;
    LigGen := TOBGEN.GetValue('GL_NUMLIGNE');
  end;
  TraiteLesDim(LigGen, False);
end;

//modif 02/08/2001

procedure TFTicketFO.TraiteLesDim(var ARow: integer; NewArt: boolean);
var RefUnique, CodeArticle, LeCom, TypeDim, LibDim, Grille, CodeDim, CodeArrondi, ArtLie: string;
  TOBL, TOBD, TOBA, TOBArt: TOB;
  i, NbD, LaLig, ACol, k, NbDimInsert, j, NbSup, Indice, ifam: integer;
  Cancel, Premier, OkPxUnique, RemA, ModifLigne: boolean;
  Qte, TotQte, Prix, PrixUnique, TotPrix, Remise, TotRem: Double;
  Marge, MontantHT, MontantTTC, TotalHT, TotalTTC, QteReliq, TotQteReliq: Double; //MAJ Ligne GEN
  TOBItem, TobTemp, TobTempArt, TobTempArtDim, TobDispo, TobDispoArt, TobCond: TOB;
  QArticle, QDispo, QCond: TQuery;
  RefGen, OldRefPrec, StQuelDepot: string;
  FamilleTaxe: array[1..5] of string;
begin
  TobTempArtDim := nil;
  TobDispo := nil;
  TobTempArt := nil;
  RemA := False;
  TotQteReliq := 0;
  PrixUnique := 0;
  QteReliq := 0;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  if (TOBDim.Detail.Count <= 0) and (TOBDim.FieldExists('ANNULE')) then
  begin
    TOBDim.DelChampSup('ANNULE', False);
    exit;
  end; //saisie dim annulé
  // Evite d'exectuter n fois les requêtes
  RefGen := TOBL.GetValue('GL_CODEARTICLE');
  RefUnique := TOBL.GetValue('GL_ARTICLE');
  if RefGen = '' then RefGen := TOBL.GetValue('GL_CODESDIM');
  if RefUnique = '' then RefUnique := CodeArticleUnique2(RefGen, '');
  TOBArt := FindTOBArtSais(TOBArticles, RefUnique);
  if ((TOBArt = nil) or (TOBArt.GetValue('GA_STATUTART') = 'GEN')) then
  begin
    if ExisteSQL('SELECT GAL_ARTICLE FROM ARTICLELIE WHERE GAL_TYPELIENART="LIE" AND GAL_ARTICLE="' + RefGen + '"') then
      ArtLie := 'X' else ArtLie := '-';
    QArticle := OpenSQL('Select * from Article where GA_CODEARTICLE="' + RefGen + '" AND GA_STATUTART ="DIM"', True);
    //JD : Chargement de tous les dépôts en multi-dépôts sinon chargement du dépôt courant
    // et si transfert, chargement du dépôt récepteur en plus du dépôt émetteur
    StQuelDepot := CreerQuelDepot(TOBPiece);
    if StQuelDepot <> '' then StQuelDepot := ' AND GQ_DEPOT IN (' + StQuelDepot + ')';
    QDispo := OpenSQL('Select * from Dispo where GQ_ARTICLE like "' + Copy(RefUnique, 1, 18) + '%" ' +
      'AND GQ_CLOTURE="-"' + StQuelDepot, True);
    QCond := OpenSQL('SELECT * FROM CONDITIONNEMENT WHERE GCO_ARTICLE="' + Copy(RefUnique, 1, 18) + '%" ORDER BY GCO_NBARTICLE', True);
    if not QArticle.EOF then
    begin
      TobTempArt := TOB.Create('LesArticles', nil, -1);
      TobTempArt.LoadDetailDB('ARTICLE', '', '', QArticle, False);
    end;
    if not QDispo.EOF then
    begin
      TobDispo := TOB.Create('', nil, -1);
      TobDispo.LoadDetailDB('DISPO', '', '', QDispo, False);
    end;
    if not QCond.EOF then TOBConds.LoadDetailDB('CONDITIONNEMENT', '', '', QCond, True, False);
    for i := 0 to TOBDim.Detail.Count - 1 do
    begin
      TOBItem := TOBDim.Detail[i];
      if TOBItem = nil then exit;
      RefUnique := TOBItem.GetValue('GA_ARTICLE');
      if TobTempArt <> nil then TobTempArtDim := TobTempArt.FindFirst(['GA_ARTICLE'], [RefUnique], False);
      if (TobTempArtDim <> nil) and (TOBArticles.FindFirst(['GA_ARTICLE'], [RefUnique], False) = nil) then
      begin
        TobTempArtDim.AddChampSup('REFARTBARRE', False);
        TobTempArtDim.AddChampSup('REFARTTIERS', False);
        TobTempArtDim.AddChampSup('REFARTSAISIE', False);
        TobTempArtDim.AddChampSup('UTILISE', False);
        TobTempArtDim.PutValue('UTILISE', '-');
        TobTempArtDim.Changeparent(TOBArticles, -1);
        //JD : Charge le dispo du dépôt émetteur et récépteur s'il existe
        //if TobDispo<>nil then TobDispoArt:=TobDispo.FindFirst(['GQ_ARTICLE'],[RefUnique],False) ;
        if TobTempArtDim.FindFirst(['GQ_ARTICLE'], [RefUnique], False) = nil then
        begin
          if TobDispo <> nil then
          begin
            TobDispoArt := TobDispo.FindFirst(['GQ_ARTICLE'], [RefUnique], False);
            while TobDispoArt <> nil do
            begin
              TobDispoArt.Changeparent(TobTempArtDim, -1);
              TobDispoArt := TobDispo.FindNext(['GQ_ARTICLE'], [RefUnique], False);
            end;
          end;
        end;
      end;
    end;
    Ferme(QArticle);
    Ferme(QDispo);
    Ferme(QCond);
    TobTempArt.Free;
    TobDispo.Free;
  end;

  ////////////////////////////////////////////////////
  if NewArt then
  begin
    TOBA := FindTOBArtRow(TOBPiece, TOBArticles, ARow);
    if TOBA.GetValue('GA_STATUTART') <> 'GEN' then Exit;
    CodeArticle := TOBL.GetValue('GL_CODEARTICLE');
    RemA := (TOBA.GetValue('GA_REMISELIGNE') = 'X');
  end else
  begin
    TypeDim := TOBL.GetValue('GL_TYPEDIM');
    if ((TypeDim <> 'GEN') and (TypeDim <> 'DIM')) then Exit;
    if TypeDim = 'GEN' then
    begin
      CodeArticle := TOBL.GetValue('GL_CODESDIM');
      TOBA := TOBArticles.FindFirst(['GA_ARTICLE'], [CodeArticleUnique2(CodeArticle, '')], False);
      if TOBA <> nil then RemA := (TOBA.GetValue('GA_REMISELIGNE') = 'X');
    end;
    if TypeDim = 'DIM' then CodeArticle := TOBL.GetValue('GL_CODEARTICLE');
  end;
  if CodeArticle = '' then Exit;
  if PieceContainer.TransfoPiece then
  begin
    if TOBDim.Detail.Count = 0 then MAJLigneDim(ARow) //Génération de pièce en aveugle
    else ReconstruireTobDim(ARow);
  end else
  begin
    if (ANCIEN_TOBDimDetailCount <> 0) then NbSup := SupDimAZero(GS, CodeArticle, TOBDim, TOBPiece, ARow) else NbSup := 0; // Supprime si QTE=0 dans THDim
    if TOBDim.Detail.Count = 0 then
    begin
      CreerTOBLignes(GS, TOBPiece, TOBTiers, TOBAffaire, ARow);
      exit;
    end;
  end;
  NbD := TOBDim.Detail.Count - (ANCIEN_TOBDimDetailCount - NbSup); //ac
  NbDimInsert := NbD;
  GS.RowCount := GS.RowCount + NbD;
  CreerTOBLignes(GS, TOBPiece, TOBTiers, TOBAffaire, TOBPiece.Detail.Count + NbD);
  TotQte := 0;
  Premier := True;
  OkPxUnique := True;
  TotPrix := 0;
  Remise := 0; //TotRem:=0 ;
  for ifam := 1 to 5 do FamilleTaxe[ifam] := '.';
  for i := 0 to TOBDim.Detail.Count - 1 do
  begin
    TOBD := TOBDim.Detail[i];
    {Infos depuis TOBDim}
    RefUnique := TOBD.GetValue('GA_ARTICLE');
    CodeArticle := Trim(Copy(RefUnique, 1, 18));
    Qte := Arrondi(TOBD.GetValue('GL_QTEFACT'), 6); //if Qte=0 then Continue ;
    CodeArrondi := TOBD.GetValue('GL_CODEARRONDI');
    if RemA then Remise := Valeur(TOBD.GetValue('GL_REMISELIGNE')); //Si remiseligne autorisé sur l'article
    //trf
    if TOBPiece.GetValue('GP_NATUREPIECEG') = 'TEM' then Prix := TOBD.GetValue('GL_PUHTDEV')
      //fin trf
    else if GP_FACTUREHT.Checked then Prix := TOBD.GetValue('GL_PUHTDEV') else Prix := TOBD.GetValue('GL_PUTTCDEV');
    for J := ARow to ARow + TOBDim.Detail.Count - 1 do
    begin
      TOBL := TOBPiece.Detail[j];
      if TOBL.GetValue('GL_ARTICLE') = RefUnique then Break
      else TOBL := nil;
    end;
    if TOBL <> nil then
    begin
      {Modification ligne}
      LaLig := j + 1;
      if SG_QF >= 0 then
      begin
        GS.Cells[SG_QF, LaLig] := FloatToStr(Qte);
        TraiteQte(SG_QF, LaLig);
      end;
      if SG_QS >= 0 then
      begin
        GS.Cells[SG_QS, LaLig] := FloatToStr(Qte);
        TraiteQte(SG_QS, LaLig);
      end;
    end else
    begin
      {Insertion nouvelle ligne}
      NbD := i;
      Inc(NbD);
      LaLig := ARow + NbD;
      GS.Cells[SG_RefArt, LaLig] := RefUnique;
      ACol := SG_RefArt;
      Cancel := False;
      if (NbDimInsert > 0) and (ARow < TobPiece.Detail.Count) then
      begin
        InsertTOBLigne(TOBPiece, Lalig);
        TOBPiece.Detail[TobPiece.Detail.count - 1].Free;
        DeleteTobLigneTarif(TobPiece, TobPiece.Detail.count - 1);
        GS.DeleteRow(TobPiece.Detail.count + 1);
        NbDimInsert := NbDimInsert - 1;
        GS.RowCount := GS.RowCount + 1;
      end;
      TraiteArticle(ACol, LaLig, Cancel, True, False, Qte);
      if Cancel then Continue;
      GS.Cells[SG_RefArt, LaLig] := CodeArticle;
    end;
    TOBL := GetTOBLigne(TOBPiece, LaLig);
    TOBL.PutValue('GL_TYPEDIM', 'DIM');
    TOBL.PutValue('GL_QTEFACT', Qte);
    TOBL.PutValue('GL_QTESTOCK', Qte);
    TOBL.PutValue('GL_CODEARRONDI', CodeArrondi);
    TOBL.PutValue('ARTSLIES', ArtLie);
    if RemA then TOBL.PutValue('GL_REMISELIGNE', Remise); //Si remiseligne autorisé sur l'article
    if GP_FACTUREHT.Checked then TOBL.PutValue('GL_PUHTDEV', Prix) else TOBL.PutValue('GL_PUTTCDEV', Prix);
    AfficheLaLigne(LaLig);
    TotQte := TotQte + Qte;
    TotQteReliq := TotQteReliq + QteReliq;
    TotPrix := Arrondi(TotPrix + Qte * Prix, PieceContainer.DEV.Decimale); //TotRem:=TotRem+Remise;
    if Premier then PrixUnique := Prix else if Prix <> PrixUnique then OkPxUnique := False;
    for ifam := 1 to 5 do
    begin
      if (FamilleTaxe[ifam] <> TOBL.GetValue('GL_FAMILLETAXE' + IntToStr(ifam))) and (FamilleTaxe[ifam] <> '.') then FamilleTaxe[ifam] := ''
      else FamilleTaxe[ifam] := TOBL.GetValue('GL_FAMILLETAXE' + IntToStr(ifam));
    end;
    Premier := False;
  end;
  TOBPiece.puTvalue('GP_RECALCULER', 'X');
  CalculeLaSaisie(-1, -1, False);
  {Article générique passe en commentaire}
  TOBL := GetTOBLigne(TOBPiece, ARow);
  LeCom := TOBL.GetValue('GL_LIBELLE');
  OldRefPrec := TOBL.GetValue('GL_PIECEPRECEDENTE');
  ClickDel(ARow, True, False);
  ClickInsert(ARow);
  TOBL := GetTOBLigne(TOBPiece, ARow);
  TOBL.PutValue('GL_LIBELLE', LeCom);
  TOBL.PutValue('GL_PIECEPRECEDENTE', OldRefPrec);
  TOBL.PutValue('GL_TYPEDIM', 'GEN');
  TOBL.PutValue('GL_TYPELIGNE', 'COM');
  TOBL.PutValue('GL_QTEFACT', TotQte);
  TOBL.PutValue('GL_QTESTOCK', TotQte);
  TOBL.PutValue('GL_CODESDIM', CodeArticle);
  TOBL.PutValue('GL_CODEARRONDI', CodeArrondi);
  // pour InfoLignes
  TOBL.PutValue('GL_DATELIVRAISON', TobPiece.GetValue('GP_DATELIVRAISON'));
  TOBL.PutValue('GL_REFARTSAISIE', CodeArticle);
  //////////////////
  if not OkPxUnique then
  begin
    if ((ctxMode in V_PGI.PGIContexte) and (TotQte <> 0)) then PrixUnique := Arrondi(TotPrix / TotQte, PieceContainer.DEV.Decimale);
  end;
  if GP_FACTUREHT.Checked then TOBL.PutValue('GL_PUHTDEV', PrixUnique) else TOBL.PutValue('GL_PUTTCDEV', PrixUnique);
  TobTemp := TOBArticles.FindFirst(['GA_ARTICLE'], [RefUnique], False);
  if TobTemp <> nil then
  begin
    TOBL.PutValue('GL_QUALIFQTEVTE', TobTemp.GetValue('GA_QUALIFUNITEVTE'));
    if TobTemp.FieldExists('GCA_QUALIFUNITEACH') then TOBL.PutValue('GL_QUALIFQTEACH', TobTemp.GetValue('GCA_QUALIFUNITEACH'));
    TOBL.PutValue('GL_QUALIFQTESTO', TobTemp.GetValue('GA_QUALIFUNITESTO'));
    TOBL.PutValue('GL_FOURNISSEUR', TobTemp.GetValue('GA_FOURNPRINC'));
  end;
  for ifam := 1 to 5 do TOBL.PutValue('GL_FAMILLETAXE' + IntToStr(ifam), FamilleTaxe[ifam]);
  {Rebalayage des lignes dimensionnées}
  GS.RowHeights[TOBPiece.Detail.Count + 1] := GS.DefaultRowHeight;
  AffichageDim;
  TOBDim.ClearDetail;
end;

//modif 02/08/2001

procedure TFTicketFO.AffichageDim; // Pour gérer l'affichage par rapport à DimSaisie
var TOBA, TOBL: Tob;
  Grille, CodeDim, LibDim: string;
  k, i: integer;
begin
  GS.RowCount := GS.RowCount + TOBPiece.Detail.count;
  for i := 0 to TOBPiece.Detail.Count - 1 do
  begin
    if DimSaisie = 'GEN' then GS.RowHeights[i + 1] := GS.DefaultRowHeight;
    TOBL := TOBPiece.Detail[i]; //if TOBL.GetValue('GL_CODEARTICLE')<>CodeArticle then Continue ;
    if TOBL.GetValue('GL_TYPEDIM') = 'DIM' then
    begin
      if DimSaisie = 'GEN' then GS.RowHeights[i + 1] := 0;
      TOBA := FindTOBArtRow(TOBPiece, TOBArticles, i + 1);
      if TOBA <> nil then for k := 1 to 5 do
        begin
          Grille := TOBA.GetValue('GA_GRILLEDIM' + IntToStr(k));
          CodeDim := TOBA.GetValue('GA_CODEDIM' + IntToStr(k));
          if ((Grille <> '') and (CodeDim <> '')) then
          begin
            LibDim := RechDom('GCGRILLEDIM' + IntToStr(k), Grille, True) + ' ' + GCGetCodeDim(Grille, CodeDim, k);
            LibDim := Copy(TOBL.GetValue('GL_LIBELLE') + '  ' + LibDim, 1, 70);
            TOBL.PutValue('GL_LIBELLE', LibDim);
          end;
        end;
    end else if DimSaisie = 'DIM' then GS.RowHeights[i + 1] := 0;
    AfficheLaLigne(i + 1);
  end;
end;

procedure TFTicketFO.MAJLigneDim(ARow: Integer);
var i, ColQuantite: Integer;
  TOBL: TOB;
begin
  i := ARow;
  ColQuantite := -1;
  if SG_QF <> -1 then ColQuantite := SG_QF else if SG_QS <> -1 then ColQuantite := SG_QS;
  if ColQuantite = -1 then Exit;
  TOBL := TOBPiece.Detail[i - 1];
  if TOBL.GetValue('GL_TYPEDIM') = 'GEN' then
  begin
    Inc(i);
    TOBL := TOBPiece.Detail[i - 1];
  end;
  while TOBL.GetValue('GL_TYPEDIM') = 'DIM' do
  begin
    GS.Cells[ColQuantite, i] := '0';
    TraiteQte(ColQuantite, i);
    inc(i);
    if i > TOBPiece.Detail.Count then Break;
    TOBL := TOBPiece.Detail[i - 1];
  end;
  TOBPiece.putValue('GP_RECALCULER', 'X');
  CalculeLaSaisie(-1, -1, False);
end;

procedure TFTicketFO.ReconstruireTobDim(ARow: Integer);
var i, j: Integer;
  TOBL, TOBD: TOB;
  CodeArticle: string;
  QteDejaSaisi, ReliqDejaSaisi: Double;
begin
  i := Arow;
  TOBL := TOBPiece.Detail[i - 1];
  if TOBL.GetValue('GL_TYPEDIM') = 'GEN' then
  begin
    CodeArticle := TOBL.GetValue('GL_CODESDIM');
    i := i + 1;
    TOBL := TOBPiece.Detail[i - 1];
  end;
  while TOBL.GetValue('GL_CODEARTICLE') = CodeArticle do
  begin
    for j := 0 to TOBDim.Detail.count - 1 do
    begin
      if TOBL.GetValue('GL_ARTICLE') = TOBDim.Detail[j].GetValue('GA_ARTICLE') then break else
      begin
        if (TOBL.GetValue('GL_ARTICLE') < TOBDim.Detail[j].GetValue('GA_ARTICLE')) or (j = TOBDim.Detail.count - 1) then
        begin
          if j = TOBDim.Detail.count - 1 then TOBD := TOB.Create('', TOBDim, TOBDim.Detail.count)
          else TOBD := TOB.Create('', TOBDim, j);
          TOBD.Dupliquer(TOBDim.Detail[TOBDim.Detail.Count - 1], false, false);
          QteDejaSaisi := CalcQteDejaSaisie(TOBPiece, TOBL.GetValue('GL_ARTICLE'));
          ReliqDejaSaisi := CalcQteDejaSaisie(TOBPiece_O, TOBL.GetValue('GL_ARTICLE'));

          if not TOBD.FieldExists('GA_ARTICLE') then TOBD.AddChampSup('GA_ARTICLE', False);
          if not TOBD.FieldExists('GL_QTEFACT') then TOBD.AddChampSup('GL_QTEFACT', False);
          if not TOBD.FieldExists('GL_PUHTDEV') then TOBD.AddChampSup('GL_PUHTDEV', False);
          if not TOBD.FieldExists('GL_PUTTCDEV') then TOBD.AddChampSup('GL_PUTTCDEV', False);
          if not TOBD.FieldExists('GL_REMISELIGNE') then TOBD.AddChampSup('GL_REMISELIGNE', False);
          if not TOBD.FieldExists('GL_CODEARRONDI') then TOBD.AddChampSup('GL_CODEARRONDI', False);
          if not TOBD.FieldExists('_QteDejaSaisi') then TOBD.AddChampSup('_QteDejaSaisi', False);
          if not TOBD.FieldExists('_ReliqDejaSaisi') then TOBD.AddChampSup('_ReliqDejaSaisi', False);

          TOBD.PutValue('GA_ARTICLE', TOBL.GetValue('GL_ARTICLE'));
          TOBD.PutValue('GL_QTEFACT', 0);
          TOBD.PutValue('GL_PUHTDEV', TOBL.GetValue('GL_PUHTDEV'));
          TOBD.PutValue('GL_PUTTCDEV', TOBL.GetValue('GL_PUTTCDEV'));
          TOBD.PutValue('GL_REMISELIGNE', TOBL.GetValue('GL_REMISELIGNE'));
          TOBD.PutValue('GL_CODEARRONDI', TOBL.GetValue('GL_CODEARRONDI'));
          TOBD.PutValue('_QteDejaSaisi', QteDejaSaisi);
          TOBD.PutValue('_ReliqDejaSaisi', ReliqDejaSaisi);
          break;
        end;
      end;
    end;
    Inc(i);
    if i > TOBPiece.Detail.Count then break;
    TOBL := TOBPiece.Detail[i - 1];
  end;
end;

{==============================================================================================}
{=============================== Manipulation des articles ====================================}
{==============================================================================================}

procedure TFTicketFO.CodesArtToCodesLigne(TOBArt: TOB; ARow: integer);
var TOBL: TOB;
  RefSais: string;
begin
  TOBL := GetTOBLigne(TOBPiece, ARow);
  TOBL.InitValeurs;
  RefSais := Trim(Copy(GS.Cells[SG_RefArt, ARow], 1, 18));
  TOBArt.PutValue('REFARTSAISIE', RefSais);
  TOBL.PutValue('GL_ARTICLE', TOBArt.GetValue('GA_ARTICLE'));
  TOBL.PutValue('GL_REFARTSAISIE', TOBArt.GetValue('REFARTSAISIE'));
  TOBL.PutValue('GL_CODEARTICLE', TOBArt.GetValue('GA_CODEARTICLE'));
  TOBL.PutValue('GL_REFARTBARRE', TOBArt.GetValue('REFARTBARRE'));
  TOBL.PutValue('GL_REFARTTIERS', TOBArt.GetValue('REFARTTIERS'));
  TOBL.PutValue('GL_TYPEREF', 'ART');
  TOBL.PutValue('GL_TYPEARTICLE', TOBArt.GetValue('GA_TYPEARTICLE'));
{$IFDEF GESCOM}
  // JTR - Gestion reste dû
  if not TOBL.FieldExists('TYPEARTFI') then TOBL.AddChampSup('TYPEARTFI',False);
  TOBL.PutValue('TYPEARTFI',TOBArt.GetValue('GA_TYPEARTFINAN'));
{$ENDIF GESCOM}
end;

function TFTicketFO.PlusEnStock(TOBArt: TOB): boolean;
var Depot: string;
  QQ: TQuery;
  QteDisp: Double;
begin
  Result := True;
  QteDisp := 0;
  if TOBArt.GetValue('GA_TENUESTOCK') = '-' then Exit;
  Depot := GP_DEPOT.Value;
  if Depot = '' then Depot := VH_GC.GCDepotDefaut;
  QQ := OpenSQL('SELECT GQ_PHYSIQUE-QG_RESERVECLI FROM DISPO WHERE GQ_ARTICLE="' + TOBArt.GetValue('GA_ARTICLE') + '" AND GQ_DEPOT="' + Depot + '"', True);
  if not QQ.EOF then QteDisp := QQ.Fields[0].AsFloat;
  Ferme(QQ);
  if QteDisp <= 0 then Exit;
  Result := False;
end;

function TFTicketFO.IdentifierArticle(var ACol, ARow: integer; var Cancel: boolean; Click, FromMacro: Boolean): boolean;
var RefSais, StatutArt, RefUnique, CodeArticle: string;
    TOBArt: TOB;
    QQ, QArt: TQuery;
    MultiDim, Okok: Boolean;
    RechArt, EtatRech: T_RechArt;
    TobMultiSel: TOB; // eQualité 11973
begin
  RefSais := GS.Cells[SG_RefArt, ARow];
  if not Click then
    if RefSais = '' then
    begin
      Result := True;
      Exit;
    end;
  Result := False;
  RefUnique := '';
  CodeArticle := '';
  StatutArt := '';
  MultiDim := False;
  TOBArt := FindTOBArtSais(TOBArticles, RefSais);
  if TOBArt <> nil then
  begin
    RechArt := traOk;
  end else
  begin
    TOBArt := CreerTOBArt(TOBArticles);
    if FromMacro then
    begin
      QArt := OpenSQL('Select * from ARTICLE Where GA_ARTICLE="' + RefSais + '"', True);
      if not QArt.EOF then
      begin
        TOBArt.SelectDB('', QArt);
        RechArt := traOk;
      end
      else RechArt := traErreur;
      Ferme(QArt);
    end else
    begin
      RechArt := TrouverArticleSQL(PieceContainer.Cledoc.NaturePiece, RefSais, '', '', PieceContainer.Cledoc.DatePiece, TOBArt, TOBTiers,GS);
    end;
    FOAjoutArticleCompl(TOBArt);
  end;
  case RechArt of
    traErreur:
      begin
        if not FromMAcro then ErreurCodeArticle(RefSais);
        CodeArticle := '';
        RefUnique := '';
        StatutArt := '';
      end;
    traSus:
      begin
        CodeArticle := '';
        RefUnique := '';
        StatutArt := '';
      end;
    traOk:
      begin
        CodeArticle := TOBArt.GetValue('GA_CODEARTICLE');
        RefUnique := TOBArt.GetValue('GA_ARTICLE');
        StatutArt := TOBArt.GetValue('GA_STATUTART');
      end;
    traOkGene:
      begin
        CodeArticle := TOBArt.GetValue('GA_CODEARTICLE');
        RefUnique := TOBArt.GetValue('GA_ARTICLE');
        StatutArt := 'UNI';
      end;
    traGrille:
      begin
        CodeArticle := TOBArt.GetValue('GA_CODEARTICLE');
        RefUnique := TOBArt.GetValue('GA_ARTICLE');
        StatutArt := TOBArt.GetValue('GA_STATUTART');
        GS.Cells[SG_RefArt, ARow] := CodeArticle;
      end;
    traAucun:
      begin
        GS.Col := SG_RefArt;
        GS.Row := ARow;
        if GereElipsis(SG_RefArt) then
        begin
          // eQualité 11973
          if (TheTob <> nil) and (TheTob.Detail.Count > 1) then
          begin
            TobMultiSel := TheTob;
            TheTob := nil;
            MultiSelectionArticle(TobMultiSel);
            TobMultiSel.Free;
            Result := False;
            exit;
          end else
          begin
            RefUnique := GS.Cells[SG_RefArt, ARow];
            TOBArt := FindTOBArtSais(TOBArticles, RefUnique);
            if TOBArt <> nil then
            begin
              CodeArticle := TOBArt.GetValue('GA_CODEARTICLE');
              StatutArt := TOBArt.GetValue('GA_STATUTART');
            end else
           // Fin eQualité 11973
            begin
              QQ := OpenSQL('Select * from ARTICLE Where GA_ARTICLE="' + RefUnique + '" ', True);
              if not QQ.EOF then
              begin
                TOBArt := CreerTOBArt(TOBArticles);
                TOBArt.SelectDB('', QQ);
                CodeArticle := QQ.FindField('GA_CODEARTICLE').AsString;
                StatutArt := QQ.FindField('GA_STATUTART').AsString;
                FOAjoutArticleCompl(TOBArt);
              end;
              Ferme(QQ);
            end;
            EtatRech := EtudieEtat(TOBArt, PieceContainer.Cledoc.NaturePiece);
            if EtatRech = traGrille then StatutArt := 'GEN' else StatutArt := 'UNI';
            GS.Cells[SG_RefArt, ARow] := CodeArticle;
          end;
        end;
      end;
  end;
  if ((StatutArt = '') or (CodeArticle = '') or (RefUnique = '')) then
  begin
    Cancel := True;
    GS.Col := ACol;
    GS.Row := ARow;
    TOBArt.Free;
    Exit;
  end;
  if StatutArt = 'GEN' then
  begin
    {$IFDEF FOS5}
    RefUnique := SelectUneDimension(RefUnique);
    {$ELSE}
    if (ctxMode in V_PGI.PGIContexte) then
    begin
      RemplirTOBDim(CodeArticle, ARow);
      if SelectMultiDimensions(CodeArticle, GS, TobPiece.GetValue('GP_DEPOT')) then //AC
      begin
        MultiDim := True;
        TOBDim := TheTOB;
        TheTOB := nil;
      end else
      begin
        RefUnique := '';
        TOBDim.ClearDetail;
        TheTOB := nil;
      end;
    end else
    begin
      RefUnique := SelectUneDimension(RefUnique);
    end;
    {$ENDIF}
    if RefUnique = '' then
    begin
      Cancel := True;
      GS.Col := ACol;
      GS.Row := ARow;
      TOBArt.Free;
      Exit;
    end;
  end;
  Okok := (StatutArt = 'UNI') or ((StatutArt = 'GEN') and (MultiDim)) or (StatutArt = 'DIM');
  if not Okok then
  begin
    TOBArt.Free;
    TOBArt := CreerTOBArt(TOBArticles);
    TOBArt.SelectDB('"' + RefUnique + '"', nil);
    FOAjoutArticleCompl(TOBArt);
  end;
  CodesArtToCodesLigne(TOBArt, ARow);
  ChargerTobArt(TOBArt,TobAFormule,TobAConsignes,VenteAchat,'',QArt) ; // JTR - eQualité 11203 (formules de qté)
  Result := True;
end;

(*procedure TFTicketFO.ChargeTOBDispo(ARow: integer);
var TOBA: TOB;
begin
  TOBA := FindTOBArtRow(TOBPiece, TOBArticles, ARow);
  if TOBA = nil then Exit;
  LoadTOBDispo(TOBA, False, CreerQuelDepot(TOBPiece));
  if (CtxMode in V_PGI.PgiContexte) and (TOBA.GetValue('GA_ARTICLE') <> 'UNI') then exit
  else LoadTOBCond(TOBA.GetValue('GA_ARTICLE'));
end;
*)

procedure TFTicketFO.ChargeTOBDispo(ARow: integer);
var TOBA: TOB;
    stQuelDepot : string;
begin
  TOBA := FindTOBArtRow(TOBPiece, TOBArticles, ARow);
  if TOBA = nil then Exit;
  stQuelDepot := CreerQuelDepot(TOBPiece);
  LoadTOBDispo(TOBA, False, stQuelDepot,TOBPiece);
  if (CtxMode in V_PGI.PgiContexte) and (TOBA.GetValue('GA_ARTICLE') <> 'UNI') then exit
  else LoadTOBCond(TOBA.GetValue('GA_ARTICLE'));
end;

procedure TFTicketFO.TraiteLeCata(ARow: integer);
var TOBL, TOBA: TOB;
begin
  if VenteAchat <> 'ACH' then Exit;
  if not CataFourn then Exit;
  TOBA := FindTOBArtRow(TOBPiece, TOBArticles, ARow);
  if TOBA = nil then Exit;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  if not TOBA.FieldExists('DPANEW') then AjouteCatalogueArt(TOBL, TOBA);
end;

procedure TFTicketFO.UpdateArtLigne(ARow: integer; FromMacro, Calc: boolean; NewQte, NewPrix: double);
var TobL : TOB;
begin
  ChargeTOBDispo(ARow);
  {$IFDEF STK}
  TobL := GetTOBLigne(TOBPiece, ARow);
  CreerTobGQDServiLigne (TobL, PieceContainer.TCGQDServi);
  {$ENDIF STK}
  InitLaLigne(ARow, NewQte, NewPrix);
  {$IFDEF TAXEUS}
  IncidenceTaxeLigne(ARow, True);
  {$ENDIF}
  if not FromMacro then
  begin
    TraiteLesDim(ARow, True);
    TraiteLesCons(ARow);
    TraiteLesNomens(ARow);
    TraiteLesMacros(ARow);
    TraiteLesLies(ARow);
    TraiteFormuleQte(ARow); // JTR - eQualité 11203 (formules de qté)
  end;
  TraiteLaCompta(ARow);
  TraiteLeCata(ARow);
  if ((not FromMacro) or (Calc)) then
  begin
    TOBPiece.putvalue('GP_RECALCULER', 'X');
    CalculeLaSaisie(SG_RefArt, ARow, False);
  end;
  ShowDetail(ARow);
  if GP_DEVISE.Enabled then GP_DEVISE.Enabled := False;
  {$IFNDEF FOS5}
  if BDEVISE.Enabled then BDEVISE.Enabled := False;
  {$ENDIF}
end;

procedure TFTicketFO.TraiteLaCompta(ARow: integer);
var RefUnique, RefCata, RefFour: string;
  TOBL, TOBA, TOBC, TOBCata: TOB;
begin
  TOBCata := nil;
  TOBA := nil;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  if TOBL.GetValue('GL_TYPEREF') = 'CAT' then
  begin
    GetCodeCataUnique(TOBPiece, ARow, RefCata, RefFour);
    if (RefCata = '') or (RefFour = '') then Exit;
  end else
    RefUnique := GetCodeArtUnique(TOBPiece, ARow);
  if RefUnique = '' then Exit;
  if ((OkCpta) or (OkCptaStock)) then
  begin
    if TOBL.GetValue('GL_TYPEREF') = 'CAT' then
    begin
      TOBCata := FindTOBCataRow(TOBPiece, TOBCataLogu, ARow);
      if TOBCata = nil then Exit;
    end else
    begin
      TOBA := FindTOBArtRow(TOBPiece, TOBArticles, ARow);
      if TOBA = nil then Exit;
    end;
    TOBC := ChargeAjouteCompta(TOBCpta, TOBPiece, TOBL, TOBA, TOBTiers, TOBCata, TobAffaire, True);
  end else
  begin
    TOBC := nil;
  end;
  PreVentileLigne(TOBC, TOBAnaP, TOBAnaS, TOBL);
end;

procedure TFTicketFO.TraiteLesNomens(ARow: integer);
var RefUnique, TypeArt, TypeNomenc, Depot: string;
  TOBL, TOBNomen, TOBLN: TOB;
  IndiceNomen: integer;
begin
  RefUnique := GetCodeArtUnique(TOBPiece, ARow);
  if RefUnique = '' then Exit;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  TypeArt := TOBL.GetValue('GL_TYPEARTICLE');
  if TypeArt <> 'NOM' then Exit;
  TypeNomenc := TOBL.GetValue('GL_TYPENOMENC');
  if TypeNomenc <> 'ASS' then Exit;
  Depot := TOBL.GetValue('GL_DEPOT');
  if Depot = '' then Exit;
  TOBNomen := ChoixNomenclature(RefUnique, Depot, TobPiece,TOBArticles, False);
  TOBLN := TOB.Create('', TOBNomenclature, -1);
  TOBLN.AddChampSup('UTILISE', False);
  TOBLN.PutValue('UTILISE', '-');
  if TOBNomen <> nil then
  begin
    NomenligVersLignomen(TOBL, TOBNomen, TOBLN, 1, 0);
  end else
  begin
    TOB.Create('LIGNENOMEN', TOBLN, -1);
    Entree_LigneNomen(TOBLN, TobArticles, TOBL, False, 1, 0);
  end;
  RenseigneValoNomen(TOBL, TOBLN);
  TOBNomen.Free;
  IndiceNomen := TOBNomenclature.Detail.Count;
  TOBL.PutValue('GL_INDICENOMEN', IndiceNomen);
end;

{$IFNDEF STK}
procedure TFTicketFO.DetruitLot(ARow: integer);
var TOBA, TOBL, TOBDepot, TOBNoeud: TOB;
  IndiceLot: integer;
  Depot: string;
begin
  if PieceContainer.Action = taConsult then Exit;
  if not GereLot then Exit;
  TOBA := FindTOBArtRow(TOBPiece, TOBArticles, ARow);
  if TOBA = nil then Exit;
  if TOBA.GetValue('GA_LOT') <> 'X' then Exit;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  IndiceLot := TOBL.GetValue('GL_INDICELOT');
  if IndiceLot <= 0 then Exit;
  Depot := TOBL.GetValue('GL_DEPOT');
  TOBDepot := TOBA.FindFirst(['GQ_DEPOT'], [Depot], False);
  if TOBDepot = nil then Exit;
  TOBNoeud := TOBDesLots.Detail[IndiceLot - 1];
  SupprimeLot(TobDepot, TobNoeud, VenteAchat);
end;

procedure TFTicketFO.GereLesLots(ARow: integer);
var TOBA, TOBDepot, TOBL, TOBNoeud: TOB;
  Depot, RefUnique: string;
  IndiceLot: integer;
  Qte, NewQte: Double;
  Valid, OkSerie: boolean;
begin
  if PieceContainer.Action = taConsult then Exit;
  if not GereLot then Exit;
  TOBA := FindTOBArtRow(TOBPiece, TOBArticles, ARow);
  if TOBA = nil then Exit;
  if TOBA.GetValue('GA_LOT') <> 'X' then Exit;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  if not TOBL.FieldExists('QTECHANGE') then
  begin
    TOBL.AddChampSup('QTECHANGE', False);
    TOBL.PutValue('QTECHANGE', 'X');
  end;
  if TOBL.GetValue('QTECHANGE') <> 'X' then Exit;
  Depot := TOBL.GetValue('GL_DEPOT');
  RefUnique := TOBL.GetValue('GL_ARTICLE');
  Qte := TOBL.GetValue('GL_QTEFACT');
  TOBDepot := TOBA.FindFirst(['GQ_DEPOT'], [Depot], False);
  if TOBDepot = nil then
  begin
    TOBDepot := TOB.Create('DISPO', TOBA, -1);
    TOBDepot.PutValue('GQ_ARTICLE', RefUnique);
    TOBDepot.PutValue('GQ_DEPOT', Depot);
    TOBDepot.PutValue('GQ_CLOTURE', '-');
  end;
  IndiceLot := TOBL.GetValue('GL_INDICELOT');
  if IndiceLot > 0 then TOBNoeud := TOBDesLots.Detail[IndiceLot - 1] else
  begin
    TOBNoeud := CreerNoeudLot(TOBPiece, TOBDesLots);
    IndiceLot := TOBDesLots.Detail.Count;
    TOBL.PutValue('GL_INDICELOT', IndiceLot);
  end;
  OkSerie := (GereSerie) and (TOBA.GetValue('GA_NUMEROSERIE') = 'X');
  GS.CacheEdit;
  GS.SynEnabled := False;
  Valid := Entree_LigDispolot(TOBDepot, TOBNoeud, TOBL, TobSerie, OkSerie, PieceContainer.Action);
  GS.MontreEdit;
  GS.SynEnabled := True;
  TOBL.PutValue('QTECHANGE', '-');
  if not Valid then
  begin
    NewQte := Arrondi(TOBNoeud.GetValue('QUANTITE'), 6);
    if NewQte <> Qte then
    begin
      if SG_QS > 0 then
      begin
        GS.Cells[SG_QS, ARow] := StrF00(NewQte, V_PGI.OkDecQ);
        TraiteQte(SG_QS, ARow);
      end;
      if SG_QF > 0 then
      begin
        GS.Cells[SG_QF, ARow] := StrF00(NewQte, V_PGI.OkDecQ);
        TraiteQte(SG_QF, ARow);
      end;
    end;
    TOBPIece.putvalue('GP_RECALCULER', 'X');
    TOBL.PutValue('GL_RECALCULER', 'X');
    CalculeLaSaisie(SG_QF, ARow, False);
  end;
end;

procedure TFTicketFO.DetruitSerie(ARow: integer);
var TOBL, TOBLS: TOB;
  IndiceSerie, i_ind: integer;
begin
  {$IFDEF CCS3}
  Exit;
  {$ELSE}
  if PieceContainer.Action = taConsult then Exit;
  if not GereSerie then Exit;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  if TOBL.GetValue('GL_NUMEROSERIE') <> 'X' then Exit;
  IndiceSerie := TOBL.GetValue('GL_INDICESERIE');
  if IndiceSerie <= 0 then Exit;
  TOBL.PutValue('GL_INDICESERIE', 0);
  for i_ind := TOBSerie.Detail[IndiceSerie - 1].Detail.Count - 1 downto 0 do
  begin
    TOBLS := TOBSerie.Detail[IndiceSerie - 1].Detail[i_ind];
    if i_ind = 0 then TOBLS.PutValue('GLS_IDSERIE', '')
    else TOBLS.Free;
  end;
  {$ENDIF}
end;

procedure TFTicketFO.GereLesSeries(ARow: integer);
var TOBA, TOBL, TOBM, TOBN, TOBPlat: TOB;
  RefUnique, QualQte: string;
  IndiceSerie, IndiceNomen, i_ind: integer;
  Qte, NewQte, UnitePiece, QteN, NbCompoSerie, QteRel: Double;
  Valid: boolean;
begin
  {$IFDEF CCS3}
  Exit;
  {$ELSE}
  if PieceContainer.Action = taConsult then Exit;
  if not GereSerie then Exit;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  if TOBL.GetValue('GL_TYPELIGNE') <> 'ART' then Exit;
  if not TOBL.FieldExists('QTECHANGE') then
  begin
    TOBL.AddChampSup('QTECHANGE', False);
    TOBL.PutValue('QTECHANGE', 'X');
  end;
  if TOBL.GetValue('QTECHANGE') <> 'X' then Exit;
  IndiceNomen := 0;
  NbCompoSerie := 0;
  Qte := TOBL.GetValue('GL_QTESTOCK');
  QteRel := TOBL.GetValue('GL_QTERELIQUAT');
  if ((TOBL.GetValue('GL_TYPEARTICLE') = 'NOM') and (TOBL.GetValue('GL_TYPENOMENC') = 'ASS')) then
  begin // Nomenclatures
    IndiceNomen := TOBL.GetValue('GL_INDICENOMEN');
    if IndiceNomen <= 0 then exit;
    if TOBNomenclature = nil then Exit;
    if TOBNomenclature.Detail.Count - 1 < IndiceNomen - 1 then Exit;
    TOBN := TOBNomenclature.Detail[IndiceNomen - 1];
    TOBPlat := TOB.Create('', nil, -1);
    NomenAPlat(TOBN, TOBPlat, 1);
    NbCompoSerie := 0;
    for i_ind := TOBPlat.Detail.Count - 1 downto 0 do
    begin
      RefUnique := TOBPlat.Detail[i_ind].GetValue('GLN_ARTICLE');
      TOBA := TOBArticles.FindFirst(['GA_ARTICLE'], [RefUnique], False);
      if TOBA = nil then
      begin
        TOBPlat.Detail[i_ind].Free;
        Continue;
      end;
      if TOBA.GetValue('GA_NUMEROSERIE') <> 'X' then
      begin
        TOBPlat.Detail[i_ind].Free;
        Continue;
      end;
      QteN := TOBPlat.Detail[i_ind].GetValue('GLN_QTE');
      NbCompoSerie := NbCompoSerie + QteN;
      TOBPlat.Detail[i_ind].PutValue('GLN_QTE', Qte * QteN);
      if not TOBPlat.Detail[i_ind].FieldExists('QTERELIQUAT') then
        TOBPlat.Detail[i_ind].AddChampSup('QTERELIQUAT', False);
      TOBPlat.Detail[i_ind].PutValue('QTERELIQUAT', QteRel * QteN);
    end;
    if (TOBPlat.Detail.Count = 0) or (NbCompoSerie = 0) then
    begin
      if TOBPlat <> nil then TOBPlat.Free;
      Exit;
    end else
    begin
      TOBL.PutValue('GL_NUMEROSERIE', 'X');
    end;
  end else
  begin
    //   TOBA:=FindTOBArtRow(TOBPiece,TOBArticles,ARow) ; if TOBA=Nil then Exit ;
    //   if TOBA.GetValue('GA_NUMEROSERIE')<>'X' then Exit ;
    if TOBL.GetValue('GL_NUMEROSERIE') <> 'X' then Exit;
    TOBPlat := nil;
  end;
  TOBL.PutValue('QTECHANGE', '-');
  GS.CacheEdit;
  GS.SynEnabled := False;
  IndiceSerie := TOBL.GetValue('GL_INDICESERIE');
  if IndiceSerie <= 0 then
  begin
    IndiceSerie := TOBSerie.Detail.Count + 1;
    TOBL.PutValue('GL_INDICESERIE', IndiceSerie);
  end;
  Valid := Entree_SaisiSerie(TOBL, TOBPlat, TOBSerie, TOBSerie_O, TOBSerRel, PieceContainer.Action);
  GS.MontreEdit;
  GS.SynEnabled := True;
  if TOBPlat <> nil then TOBPlat.Free;
  if not Valid then
  begin
    if (IndiceSerie > TOBSerie.Detail.Count) then
    begin
      TOBL.PutValue('GL_INDICESERIE', 0);
      NewQte := 0;
    end else
    begin
      if IndiceNomen <= 0 then NewQte := TOBSerie.Detail[IndiceSerie - 1].Detail.Count
      else NewQte := TOBSerie.Detail[IndiceSerie - 1].Detail.Count / NbCompoSerie;
      if VenteAchat = 'ACH' then QualQte := TOBL.GetValue('GLN_QUALIFQTEACH')
      else if (VenteAchat = 'AUT') or (VenteAchat = 'TRF') then QualQte := TOBL.GetValue('GL_QUALIFQTESTO')
      else QualQte := TOBL.GetValue('GL_QUALIFQTEVTE');
      TOBM := VH_GC.TOBMEA.FindFirst(['GME_QUALIFMESURE', 'GME_MESURE'], ['PIE', QualQte], False);
      UnitePiece := 0;
      if TOBM <> nil then UnitePiece := TOBM.GetValue('GME_QUOTITE');
      if (UnitePiece = 0) or (IndiceNomen > 0) then UnitePiece := 1.0;
      NewQte := Arrondi(NewQte / UnitePiece, 6);
    end;
    if NewQte <> Qte then
    begin
      if SG_QS > 0 then
      begin
        GS.Cells[SG_QS, ARow] := StrF00(NewQte, V_PGI.OkDecQ);
        TraiteQte(SG_QS, ARow);
      end;
      if SG_QF > 0 then
      begin
        GS.Cells[SG_QF, ARow] := StrF00(NewQte, V_PGI.OkDecQ);
        TraiteQte(SG_QF, ARow);
      end;
    end;
    TOBPIece.putvalue('GP_RECALCULER', 'X');
    CalculeLaSaisie(SG_QF, ARow, False);
    TOBL.PutValue('QTECHANGE', '-');
  end;
  {$ENDIF}
end;
{$ENDIF STK}

{$IFDEF STK}
procedure TFTicketFO.GereStock (ARow: integer; bAppel, bForceAffichage, bAuto : boolean; bParle : boolean = True);
var TobL, TobA, TobGQDServi, TobDispo : TOB;
    dNewQte, dQte : double;
    bTraiteQte : boolean;
    Article, Depot, NaturePiece : string;
begin
  TobL := GetTOBLigne (TobPiece, ARow);
  if TobL = nil then Exit;

  Article := TobL.GetString('GL_ARTICLE');
  Depot := TobL.GetString('GL_DEPOT');
  NaturePiece := Tobl.GetString ('GL_NATUREPIECEG');

  if (TOBL.GetString ('GL_TYPELIGNE') <> 'ART') or       // pas une ligne article
     (not TobL.GetBoolean ('GL_TENUESTOCK') and not (TOBL.GetString ('GL_TYPEARTICLE')='NOM')) or
     (GetInfoParPiece (NaturePiece, 'GPP_STKQUALIFMVT') = '') then Exit; // pas de mouvement de stock

  TobGQDServi := Tob(LongInt (TobL.GetValue ('TobGQDServi')));

  if EstNomenclatureASS(TOBL) then
  begin
    GereStockNomen(ARow,bAppel,bForceAffichage,bAuto,bParle);
    exit;
  end;
  TOBA := TOBArticles.FindFirst(['GA_ARTICLE'], [Article], False);
  TobDispo := TobArticles.FindFirst (['GQ_ARTICLE', 'GQ_DEPOT'], [Article, Depot], true);
  if TobDispo = nil then  //JS 29/07/04
  begin
   TobDispo:=TOB.Create('DISPO',TOBA,-1) ;
   TobDispo.PutValue('GQ_ARTICLE',Article) ;
   TobDispo.PutValue('GQ_DEPOT',Depot) ;
   TobDispo.PutValue('GQ_CLOTURE','-') ;
   DispoChampSupp(TobDispo);
   TobDispo.AddChampSupValeur('NEW_ENREG','X');
  end;
  if (GetStkTypeMvtFromGPP (NaturePiece) = 'ATT') and
     (not bForceAffichage) and
     ((not PieceContainer.GereSerie) or (not TOBL.GetBoolean('GL_NUMEROSERIE'))) and
     ((not PieceContainer.GereLot) or (not TOBA.GetBoolean ('GA_LOT'))) then
  begin
    SertLeStock (TobL, PieceContainer);
  end else
  begin
    if (GetFieldFromGSN ('GSN_STKTYPEMVT', GetInfoParPiece (Tobl.GetString ('GL_NATUREPIECEG'), 'GPP_STKQUALIFMVT')) = 'PHY') or
       (bForceAffichage) or
       (((TobA.GetBoolean ('GA_LOT')) and         (PieceContainer.GereLot)) or
        ((TobA.GetBoolean ('GA_NUMEROSERIE')) and (PieceContainer.GereSerie)) or (TobL.GetDouble ('GL_QTEFACT') < QuelleQteDetailServi (TobL))) then
    begin
      if not TOBL.FieldExists('QTECHANGE') then
      begin
        TOBL.AddChampSup('QTECHANGE', False);
        TOBL.PutValue('QTECHANGE', 'X');
      end;
      if (TOBL.GetValue('QTECHANGE') <> 'X') and (not bAppel) then Exit;
      dQte := TobL.GetDouble ('GL_QTEFACT');

      if (TobL.GetBoolean ('QTECHANGE')) and (TobA.GetBoolean ('GA_LOT')) and
         (TobA.GetBoolean ('GA_NUMEROSERIE')) then bForceAffichage := true;
      bTraiteQte := False;
      while not bTraiteQte do
      begin
(*
        if not AvecEcran and Assigned(PieceContainer.TCGQDSaisie) then
        { cas hors écran en passe par "GerePiece.pas" }
          AffectServiExt(TobGQDServi, PieceContainer.TCGQDSaisie)
        else if AvecEcran or ((dQte = 0) and (TobGQDServi.Detail.Count > 0)) then
        { cas normal : on passe par la Dfm de "Facture.pas" } { ou } { on passe par "GerePiece.pas" et on met à 0 les GL_QTEFACT et le dispo prévu }
          Entree_SaisieDetail (TOBL, TobGQDServi, TobDispo, PieceContainer.Action, dQte, bForceAffichage, bAuto, PieceContainer);
*)
//        if (dQte = 0) and (TobGQDServi.Detail.Count > 0) then
          Entree_SaisieDetail (TOBL, TobGQDServi, TobDispo, PieceContainer.Action, dQte, bForceAffichage, bAuto, PieceContainer);
        if PieceContainer.Action <> taConsult then
        begin
          TobL.PutValue ('QTECHANGE', '-');
          if (not bForceAffichage) or
             ((GetStkTypeMvtFromGPP (TobL.GetString ('GL_NATUREPIECEG')) <> 'RES') and
              (GetStkTypeMvtFromGPP (TobL.GetString ('GL_NATUREPIECEG')) <> 'ATT'))  then // on a cliquer sur "détail du stock"
          begin
            dNewQte := QuelleQteDetailServi (TobL);
          end else
          begin
            if QuelleQteDetailServi (TobL) > dQte then
              dNewQte := QuelleQteDetailServi (TobL)
            else
              dNewQte := dQte;
          end;
          if (dNewQte <> dQte) or (not bTraiteQte) then
          begin
            if TobA.GetValue('GA_TYPEARTICLE') <> 'CNS' then
            begin
              if SG_QS > 0 then
              begin
                GS.Cells[SG_QS, ARow] := StrF00(dNewQte, V_PGI.OkDecQ);
                bTraiteQte := TraiteQte(SG_QS, ARow, bParle);
                if bTraiteQte then
                  if Valeur (GS.Cells [SG_QS, ARow]) = dNewQte then TobL.SetBoolean ('QTECHANGE', false);
              end;
              if SG_QF > 0 then
              begin
                GS.Cells[SG_QF, ARow] := StrF00(dNewQte, V_PGI.OkDecQ);
                bTraiteQte := TraiteQte(SG_QF, ARow, bParle);
                if bTraiteQte then
                  if Valeur (GS.Cells [SG_QF, ARow]) = dNewQte then TobL.SetBoolean ('QTECHANGE', false);
              end;
            end else bTraiteQte := true;
          end;
        end else bTraiteQte := true;
      end;
      TOBPiece.PutValue('GP_RECALCULER', 'X');
      TOBL.PutValue('GL_RECALCULER', 'X');
      CalculeLaSaisie(SG_QF, ARow, False);
    end;
  end;
end;

function TFTicketFO.GereStockNomen (ARow: integer; bAppel, bForceAffichage, bAuto : boolean; bParle : boolean = true) : boolean;
var TobL, TobGQDServi : TOB;
    dNewQte, dQte : double;
    Article, Depot, NaturePiece : string;
begin
  result := true;
  TobL := GetTOBLigne (TobPiece, ARow);
  if TobL = nil then Exit;
  Article := TobL.GetString('GL_ARTICLE');
  Depot := TobL.GetString('GL_DEPOT');
  NaturePiece := Tobl.GetString ('GL_NATUREPIECEG');

  if (GetFieldFromGSN ('GSN_STKTYPEMVT', GetInfoParPiece (NaturePiece, 'GPP_STKQUALIFMVT')) = 'PHY') or
      // ((TobPlat.FindFirst(['GLN_ARTICLE'],[Article],false) <> nil) and ((PieceContainer.GereLot) or (PieceContainer.GereSerie))) then
      ((PieceContainer.GereLot) or (PieceContainer.GereSerie)) then
  begin
    if not TOBL.FieldExists('QTECHANGE') then
    begin
      TOBL.AddChampSup('QTECHANGE', False);
      TOBL.SetString('QTECHANGE', 'X');
    end;
    if (TOBL.GetValue('QTECHANGE') <> 'X') and (not bAppel) then Exit;
    dQte := TobL.GetDouble ('GL_QTEFACT');
    TobL.SetString ('QTECHANGE', '-');
    TobGQDServi := Tob(LongInt(TobL.GetValue ('TobGQDServi')));
    if not Entree_SaisieDetailNomenclature(PieceContainer,ARow,bForceAffichage) then
    begin
      result := false;
      dNewQte := TobGQDServi.GetDouble('QTEPREC');
      if (dNewQte <> dQte) then
      begin
        if SG_QS > 0 then
        begin
          GS.Cells[SG_QS, ARow] := StrF00(dNewQte, V_PGI.OkDecQ);
          TraiteQte(SG_QS, ARow, bParle);
          TobL.SetBoolean ('QTECHANGE', false);
        end;
        if SG_QF > 0 then
        begin
          GS.Cells[SG_QF, ARow] := StrF00(dNewQte, V_PGI.OkDecQ);
          TraiteQte(SG_QF, ARow, bParle);
          TobL.SetBoolean ('QTECHANGE', false);
        end;
      end;

    end else TobGQDServi.SetDouble('QTEPREC', dQte);
    TOBPiece.SetString('GP_RECALCULER', 'X');
    TOBL.SetString('GL_RECALCULER', 'X');
    CalculeLaSaisie(SG_QF, ARow, False);
  end;
end;

function TFTicketFO.DuplicLeDetailSurAvoir : boolean;
var OldNature : string;

begin
  result := false;
  if PieceContainer.DuplicPiece and GerePhy(PieceContainer.NewNature) then
  begin
    OldNature := TOBPiece_O.GetString('GP_NATUREPIECEG');
    result := GerePhy(OldNature) and
             (GetInfoParPiece(PieceContainer.NewNature,'GPP_SENSPIECE') <> GetInfoParPiece(OldNature,'GPP_SENSPIECE'));
   end;
end;
{$ENDIF STK}

procedure TFTicketFO.GereArtsLies(ARow: integer);
var RefUnique, CodeArticle, LesRefsLies, StLoc, QteRef: string;
  TOBL: TOB;
  Qte: Double;
  Cancel: boolean;
  NbLies, LaLig, ACol: integer;
begin
  if not VH_GC.GCArticlesLies then Exit;
  if PieceContainer.Action = taConsult then Exit;
  if ((PieceContainer.TransfoPiece) or (PieceContainer.DuplicPiece)) then Exit;
  if VenteAchat <> 'VEN' then Exit;
  RefUnique := GetCodeArtUnique(TOBPiece, ARow);
  if RefUnique = '' then Exit;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  if TOBL.GetValue('ARTSLIES') <> 'X' then Exit;
  TOBL.PutValue('ARTSLIES', '-');
  Qte := TOBL.GetValue('GL_QTEFACT');
  CodeArticle := Trim(Copy(RefUnique, 1, 18));
  LesRefsLies := SelectArticlesLies(CodeArticle);
  if LesRefsLies = '' then Exit;
  StLoc := LesRefsLies;
  NbLies := 0;
  {Comptage pour creer les lignes et les tob à l'avance}
  repeat
    RefUnique := ReadTokenSt(StLoc);
    ReadTokenSt(StLoc); //pour QteRef
    if RefUnique <> '' then Inc(NbLies);
  until ((RefUnique = '') or (StLoc = ''));
  GS.RowCount := GS.RowCount + NbLies;
  CreerTOBLignes(GS, TOBPiece, TOBTiers, TOBAffaire, TOBPiece.Detail.Count + NbLies);
  {Insertion et création des lignes liées}
  NbLies := 0;
  repeat
    RefUnique := ReadTokenSt(LesRefsLies);
    QteRef := ReadTokenSt(LesRefsLies);
    if RefUnique <> '' then
    begin
      CodeArticle := Trim(Copy(RefUnique, 1, 18));
      LaLig := ARow + 1 + NbLies;
      GS.Cells[SG_RefArt, LaLig] := RefUnique;
      ACol := SG_RefArt;
      Cancel := False;
      if QteRef <> '-' then TraiteArticle(ACol, LaLig, Cancel, True, True, Qte)
      else TraiteArticle(ACol, LaLig, Cancel, True, True, 1);
      if Cancel then Continue;
      GS.Cells[SG_RefArt, LaLig] := CodeArticle;
      AfficheLaLigne(LaLig);
      Inc(NbLies);
    end;
  until ((RefUnique = '') or (LesRefsLies = ''));
end;

procedure TFTicketFO.TraiteLesLies(ARow: integer);
var RefUnique, CodeArticle: string;
  TOBL: TOB;
  OkL: boolean;
begin
  if not VH_GC.GCArticlesLies then Exit;
  if PieceContainer.Action = taConsult then Exit;
  if ((PieceContainer.TransfoPiece) or (PieceContainer.DuplicPiece)) then Exit;
  RefUnique := GetCodeArtUnique(TOBPiece, ARow);
  if RefUnique = '' then Exit;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  if TOBL.GetValue('ARTSLIES') = '-' then Exit;
  CodeArticle := TOBL.GetValue('GL_CODEARTICLE');
  if CodeArticle = '' then Exit;
  OkL := ExisteSQL('SELECT GAL_ARTICLE FROM ARTICLELIE WHERE GAL_TYPELIENART="LIE" AND GAL_ARTICLE="' + CodeArticle + '"');
  if OkL then TOBL.PutValue('ARTSLIES', 'X') else TOBL.PutValue('ARTSLIES', '-');
end;

procedure TFTicketFO.TraiteLesMacros(ARow: integer);
var RefUnique, CodeArticle, TypeArt, TypeNomenc, Depot, LeCom: string;
  TOBL, TOBNomen, TOBLN, TOBPlat, TOBpl: TOB;
  i, ACol, LaLig, DecL: integer;
  Cancel: boolean;
  Qte: Double;
  LeNumOrdre, Cpt, Cpt1 : integer; // JTR - Lien nomenclature macro
begin
  RefUnique := GetCodeArtUnique(TOBPiece, ARow);
  if RefUnique = '' then Exit;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  TypeArt := TOBL.GetValue('GL_TYPEARTICLE');
  if TypeArt <> 'NOM' then Exit;
  TypeNomenc := TOBL.GetValue('GL_TYPENOMENC');
  if TypeNomenc <> 'MAC' then Exit;
  Depot := TOBL.GetValue('GL_DEPOT');
  if Depot = '' then Exit;
  TOBNomen := ChoixNomenclature(RefUnique, Depot, TobPiece,TOBArticles, False);
  TOBLN := TOB.Create('', TOBNomenclature, -1);
  TOBLN.AddChampSup('UTILISE', False);
  TOBLN.PutValue('UTILISE', '-');
  if TOBNomen <> nil then
  begin
    NomenligVersLignomen(TOBL, TOBNomen, TOBLN, 1, 0);
    TOBPlat := TOB.Create('', nil, -1);
    NomenAPlat(TOBLN, TOBPlat, 1);
    DecL := 0;
    GS.RowCount := GS.RowCount + TOBPlat.Detail.Count;
    CreerTOBLignes(GS, TOBPiece, TOBTiers, TOBAffaire, TOBPiece.Detail.Count + TOBPlat.Detail.Count);
    for i := 0 to TOBPlat.Detail.Count - 1 do
    begin
      TOBpl := TOBPlat.Detail[i];
      RefUnique := TOBpl.GetValue('GLN_ARTICLE');
      CodeArticle := TOBpl.GetValue('GLN_CODEARTICLE');
      Qte := TOBpl.GetValue('GLN_QTE');
      LaLig := ARow + 1 + i - DecL;
      GS.Cells[SG_RefArt, LaLig] := RefUnique;
      ACol := SG_RefArt;
      Cancel := False;
      TraiteArticle(ACol, LaLig, Cancel, True, True, Qte);
      if Cancel then
      begin
        Inc(DecL);
        Continue;
      end;
      GS.Cells[SG_RefArt, LaLig] := CodeArticle;
    end;
    SurLigSuivante := False;
    TOBL := GetTOBLigne(TOBPiece, ARow);
    LeCom := TOBL.GetValue('GL_LIBELLE');
    ClickDel(ARow, True, False);
    ClickInsert(ARow);
    TOBL := GetTOBLigne(TOBPiece, ARow);
    TOBL.PutValue('GL_LIBELLE', LeCom);
    // JTR - Lien nomenclature macro
    Cpt1 := 0;
    LeNumOrdre := TOBL.GetValue('GL_NUMORDRE');
    for Cpt := Arow to TOBPiece.Detail.Count do
    begin
      Cpt1 := Cpt1 + 1;
      if Cpt1 > TOBPlat.Detail.Count then break;
      TOBPiece.Detail[Cpt].PutValue('GL_LIGNELIEE',LeNumOrdre);
      TOBPiece.Detail[Cpt].PutValue('GL_TYPENOMENC','MAC');
    end;
    TOBL.PutValue('GL_LIGNELIEE', LeNumOrdre);
    TOBL.PutValue('GL_TYPENOMENC', 'MAC');
    // Fin JTR
    TOBPIece.putvalue('GP_RECALCULER', 'X');
    CalculeLaSaisie(-1, -1, True);
    // pour passer à la ligne suivante
    if not InTraiteBouton then
    begin
      SurLigSuivante := True;
      GSCellEnter(GS, ACol, ARow, Cancel); // ClickDel appelle CellExit mais ClickInsert n'appelle pas CellEnter
    end;
    DejaSaisie := True;
    TOBPlat.Free;
  end else                 
  begin
    VideCodesLigne(TOBPiece, ARow);
    InitialiseTOBLigne(TOBPiece, TOBTiers, TOBAffaire, ARow);
  end;
//  if TobLN <> nil then FreeAndNil(TOBLN);
  if TobNomenclature.detail[0] <> nil then TobNomenclature.detail[0].free;
  if TobNomen <> nil then FreeAndNil(TOBNomen);
end;

procedure TFTicketFO.TraiteLesCons(ARow: integer);
var RefUnique: string;
  TOBTOF, TOBC, TOBFille, TOBL: TOB;
  Qte: double;
  CodeCond: string;
begin
  RefUnique := GetCodeArtUnique(TOBPiece, ARow);
  if RefUnique = '' then Exit;
  TOBC := TOBConds.FindFirst(['GCO_ARTICLE', 'GCO_CHANGEUNITE'], [RefUnique, 'X'], False);
  if TOBC = nil then Exit;
  TOBTOF := TOB.Create('', nil, -1);
  TOBTOF.AddChampSup('QTEARTICLE', False);
  TOBTOF.AddChampSup('GCO_CODECOND', False);
  Qte := GetChampLigne(TOBPiece, 'GL_QTESTOCK', ARow);
  TOBTOF.PutValue('QTEARTICLE', Qte);
  CodeCond := GetChampLigne(TOBPiece, 'GL_CODECOND', ARow);
  TOBTOF.PutValue('GCO_CODECOND', CodeCond);
  while TOBC <> nil do
  begin
    TOBFille := TOB.Create('CONDITIONNEMENT', TOBTOF, -1);
    TOBFille.Dupliquer(TOBC, False, True);
    TOBC := TOBConds.FindNext(['GCO_ARTICLE', 'GCO_CHANGEUNITE'], [RefUnique, 'X'], False);
  end;
  TheTOB := TOBTOF;
  // MODIF_DBR_DEBUT
  //AGLLanceFiche('GC', 'GCSAISCOND', '', '', '');
  AGLLanceFiche('GC', 'GCSAISCOND', '', '', VenteAchat);
  // MODIF_DBR_FIN
  if TheTob <> nil then
  begin
    Qte := TheTOB.GetValue('QTEARTICLE');
    CodeCond := TheTOB.GetValue('GCO_CODECOND');
    TOBL := GetTOBLigne(TOBPiece, ARow);
    TOBL.PutValue('GL_CODECOND', CodeCond);
    if SG_QS > 0 then
    begin
      GS.Cells[SG_QS, ARow] := StrF00(Qte, V_PGI.OkDecQ);
      TraiteQte(SG_QS, ARow);
    end;
    if SG_QF > 0 then
    begin
      GS.Cells[SG_QF, ARow] := StrF00(Qte, V_PGI.OkDecQ);
      TraiteQte(SG_QF, ARow);
    end;
  end;
  TOBTOF.Free;
  TheTOB := nil;
end;

procedure TFTicketFO.TraiteArticle(var ACol, ARow: integer; var Cancel: boolean; FromMacro, Calc: boolean; NewQte: double; FromKB: boolean = FALSE);
var OkArt, bMacro, OkLiaison: Boolean;
  NewCol, NewRow: integer;
  TypeL, TypeD: string;
  NoErr: Integer;
  NewPrix: double;
  TobL : TOB;
begin
  OkLiaison := False;
  NewPrix := 0;
  TypeD := GetChampLigne(TOBPiece, 'GL_TYPEDIM', ARow);
  if GS.Cells[ACol, ARow] = '' then
  begin
    TypeL := GetChampLigne(TOBPiece, 'GL_TYPELIGNE', ARow);
    if ((TypeL <> 'TOT') and (TypeD <> 'GEN')) then
    begin
      if not FOJaiLeDroit(25) then
      begin
        AfficheLaLigne(ARow);
        Cancel := True;
        Exit;
      end;
      VideCodesLigne(TOBPiece, ARow);
      InitialiseTOBLigne(TOBPiece, TOBTiers, TOBAffaire, ARow, True);
      SupprimeLigneTarif(TobPiece, TobLigneTarif, ARow);
    end;
    Exit;
  end;
  {$IFDEF STK}
  TobL := GetTOBLigne(TOBPiece, ARow);
  if (Tobl.GetValue('GL_REFARTSAISIE')<>'') and
     (GS.Cells[ACol, ARow] <> Tobl.GetValue('GL_REFARTSAISIE')) and
     (TobL.FieldExists ('TobGQDServi')) and (EstLigneStock (TobL)) then
  begin
    RendServi (TobL, PieceContainer);
  end;
  {$ENDIF STK}
  if TypeD = 'GEN' then
  begin
    if GS.Cells[ACol, ARow] = GetChampLigne(TOBPiece, 'GL_CODESDIM', ARow) then Exit;
  end;
  NewCol := GS.Col;
  NewRow := GS.Row;
  if FromKB then bMacro := True else bMacro := FromMacro;
  OkArt := IdentifierArticle(ACol, ARow, Cancel, False, bMacro);
  if ((OkArt) and (not Cancel)) then
  begin
    GS.Col := NewCol;
    GS.Row := NewRow;
    NoErr := 0;
    case FOArticleAutorise(TOBPiece, TOBArticles, PieceContainer.Cledoc.NaturePiece, ARow) of
      afoFerme: NoErr := 2; // Article fermé
      afoInCompatible: NoErr := 45; // Article incompatible avec les autres lignes
      afoInterdit: NoErr := 46; // Article interdit à la vente
      afoConfidentiel: NoErr := -1; // Article confidentiel
    end;
    if (NoErr = 0) and (ArtFiClientOblig(nil, ARow) > 0) then NoErr := 51; // Client incompatible avec l'opération de caisse
    if NoErr = 0 then
    begin
      if not ReglementDispo(ARow, NewPrix, NewQte, OkLiaison) then
      begin
        Cancel := True;
        Exit;
      end;
    end;
    if NoErr <> 0 then
    begin
      Cancel := True;
      if NoErr > 0 then HPiece.Execute(NoErr, Caption, '');
      GS.Cells[ACol, ARow] := '';
      VideCodesLigne(TOBPiece, ARow);
      InitialiseTOBLigne(TOBPiece, TOBTiers, TOBAffaire, ARow);
      SupprimeLigneTarif(TobPiece, TobLigneTarif, ARow);
    end else
    begin
      if not (FromMacro) and not (FromKb) and not (OkLiaison) then
      begin
        // recherche du signe de la quantité en fonction de l'article
        NewQte := FOSigneQteArticle(TOBPiece, TOBArticles, ARow, NewQte);
        // recherche de la quantité saisie sur la calculette du pavé tactile
        if PnlBoutons <> nil then
          NewQte := PnlBoutons.ChercheValeurCalculette(NewQte, True);
      end;
      UpdateArtLigne(ARow, FromMacro, Calc, NewQte, NewPrix);
      if not TraiteRupture(ARow) then
      begin
        if SG_QF >= 0 then
        begin
          GS.Cells[SG_QF, ARow] := '';
          TraiteQte(SG_QF, ARow);
        end;
        if SG_QS >= 0 then
        begin
          GS.Cells[SG_QS, ARow] := '';
          TraiteQte(SG_QS, ARow);
        end;
      end;
      TOBPiece.putvalue('GP_RECALCULER', 'X');
      // pour passer à la ligne suivante
      if (((SG_RefArt <> -1) and (GS.Cells[SG_RefArt, GS.Row] <> '')) or
        ((SG_Lib <> -1) and (GS.Cells[SG_Lib, GS.Row] <> ''))) {and
      (not InTraiteBouton)}then SurLigSuivante := True;
    end;
  end;
end;

{==============================================================================================}
{======================================== Calculs =============================================}
{==============================================================================================}

procedure TFTicketFO.RecalculeSousTotaux;
var i: integer;
begin
  CalculeSousTotauxPiece(TOBPiece);
  for i := 0 to TOBPiece.Detail.Count - 1 do AfficheLaLigne(i + 1);
end;

procedure TFTicketFO.ReAfficheSousTotal(ARow: integer);
var Ind: integer;
  TOBL: TOB;
begin
  for Ind := ARow - 1 downto 1 do
  begin
    TOBL := GetTOBLigne(TOBPiece, Ind);
    if TOBL = nil then Break;
    if TOBL.GetValue('GL_TYPELIGNE') = 'TOT' then Break;
    AfficheLaLigne(Ind);
  end;
end;

procedure TFTicketFO.AffichePorcs;
var X: Double;
  i: integer;
begin
  X := 0;
  if TOBPorcs <> nil then
  begin
    for i := 0 to TOBPorcs.Detail.Count - 1 do X := X + TOBPorcs.Detail[i].GetValue('GPT_TOTALHTDEV');
  end;
  {$IFNDEF FOS5}
  TOTALPORTSDEV.Value := Arrondi(X, 6);
  {$ENDIF}
end;

procedure TFTicketFO.CalculeLeDocument;
var PieceContainerLoc: TPieceContainer;
begin
  if (TOBPIece.Getvalue('GP_RECALCULER') = 'X') then
  begin
    PieceContainerLoc := TPieceContainer.Create;
    PieceContainerLoc.CreateTobs;
    PieceContainerLoc.Dupliquer(PieceContainer);
    if BeforeCalcul(Self, PieceContainerLoc, PieceContainer.DEV) then
      CalculFacture(PieceContainerLoc);
    AfterCalcul(Self, PieceContainerLoc, PieceContainer.DEV);
    if TobPiece <> nil then TobPiece.Dupliquer(PieceContainerLoc.TCPiece, true, true);
    if TobTiers <> nil then TobTiers.Dupliquer(PieceContainerLoc.TCTiers, true, true);
    if TobPiece_O <> nil then TobPiece_O.Dupliquer(PieceContainerLoc.TCPiece_O,true,true);
    if TOBBases <> nil then TOBBases.Dupliquer(PieceContainerLoc.TCBases,true,true);
    if TOBBases_O <> nil then TOBBases_O.Dupliquer(PieceContainerLoc.TCBases_O,true,true);
    if TobEches <> nil then TobEches.Dupliquer(PieceContainerLoc.TCEches,true,true);
    if TobArticles <> nil then TobArticles.Dupliquer(PieceContainerLoc.TCArticles,true,true);
    if TobPorcs <> nil then TobPorcs.Dupliquer(PieceContainerLoc.TCPorcs,true,true);
    if TobPorcs_O <> nil then TobPorcs_O.Dupliquer(PieceContainerLoc.TCPorcs_O,true,true);
    if TobConds <> nil then TobConds.Dupliquer(PieceContainerLoc.TCConds,true,true);
    if TobTarif <> nil then TobTarif.Dupliquer(PieceContainerLoc.TCTarif,true,true);
    if TobComms <> nil then TobComms.Dupliquer(PieceContainerLoc.TCComms,true,true);
    if TobCatalogu <> nil then TobCatalogu.Dupliquer(PieceContainerLoc.TCCatalogu,true,true);
    if TobCXV <> nil then TobCXV.Dupliquer(PieceContainerLoc.TCCXV,true,true);
    if TobNomenclature <> nil then TobNomenclature.Dupliquer(PieceContainerLoc.TCNomenclature,true,true);
    if TobN_O <> nil then TobN_O.Dupliquer(PieceContainerLoc.TCN_O,true,true);
    if TobDim <> nil then TobDim.Dupliquer(PieceContainerLoc.TCDim,true,true);
    if TobDesLots <> nil then TobDesLots.Dupliquer(PieceContainerLoc.TCDesLots,true,true);
    if TobLot_O <> nil then TobLot_O.Dupliquer(PieceContainerLoc.TCLot_O,true,true);
    if TobSerie <> nil then TobSerie.Dupliquer(PieceContainerLoc.TCSerie,true,true);
    if TobSerie_O <> nil then TobSerie_O.Dupliquer(PieceContainerLoc.TCSerie_O,true,true);
    if TobSerRel <> nil then TobSerRel.Dupliquer(PieceContainerLoc.TCSerRel,true,true);
    if TobAcomptes <> nil then TobAcomptes.Dupliquer(PieceContainerLoc.TCAcomptes,true,true);
    if TobAcomptes_O <> nil then TobAcomptes_O.Dupliquer(PieceContainerLoc.TCAcomptes_O,true,true);
    if TobDispoContreM <> nil then TobDispoContreM.Dupliquer(PieceContainerLoc.TCDispoContreM,true,true);
    if TobAffaire <> nil then TobAffaire.Dupliquer(PieceContainerLoc.TCAffaire,true,true);
    if TobCpta <> nil then TobCpta.Dupliquer(PieceContainerLoc.TCCpta,true,true);
    if TobAnaP <> nil then TobAnaP.Dupliquer(PieceContainerLoc.TCAnaP,true,true);
    if TobAnaS <> nil then TobAnaS.Dupliquer(PieceContainerLoc.TCAnaS,true,true);
{$IFDEF BTP}
    if TobPieceRG <> nil then TobPieceRG.Dupliquer(PieceContainerLoc.TCPieceRG,true,true);
    if TobBasesRG <> nil then TobBasesRG.Dupliquer(PieceContainerLoc.TCBasesRG,true,true);
{$ENDIF}
    PieceContainerLoc.FreeTobs;
    PieceContainerLoc.Destroy;
  end;
end;

procedure TFTicketFO.CalculeLaSaisie(ACol, ARow: integer; AffTout: boolean);
var i: integer;
  Okok: boolean;
begin
{$IFDEF FOS5}
  if ((ACol <> SG_RefArt) and (ACol <> SG_QF) and (ACol <> SG_QS) and (ACol <> SG_Px) and
    (ACol <> SG_Rem) and (ACol <> SG_RV) and (ACol <> SG_RL) and (ACol <> SG_PxNet) and (ACol <> SG_Montant) and
    (ACol <> -1)) then Exit;
  if TOBPiece.GetValue('GP_RECALCULER') = 'X' then
  begin
    CalculeLeDocument;
    TOBPiece.PutEcran(Self, PPied);
    if ARow > 0 then AfficheLaLigne(ARow, True);
{$ELSE}
  if ((ACol <> SG_RefArt) and (ACol <> SG_QF) and (ACol <> SG_QS) and (ACol <> SG_Px) and
    (ACol <> SG_Rem) and (ACol <> SG_RV) and (ACol <> SG_RL) and (ACol <> -1)) then Exit;
  if TOBPiece.GetValue('GP_RECALCULER') = 'X' then
  begin
    CalculeLeDocument;
    TOBPiece.PutEcran(Self, PPied);
    AfficheTaxes;
    AffichePorcs;
    if ARow > 0 then AfficheLaLigne(ARow);
{$ENDIF}
    RecalculeSousTotaux(TOBPiece);
    if AffTout then
    begin
      for i := 0 to TOBPiece.Detail.Count - 1 do AfficheLaLigne(i + 1);
    end;
    if NeedVisa then
    begin
      Okok := True;
      if Okok then
      begin
        if Abs(TOBPiece.GetValue('GP_TOTALHT')) >= MontantVisa then
          TOBPiece.PutValue('GP_ETATVISA', 'ATT')
          else
          TOBPiece.PutValue('GP_ETATVISA', 'NON');
      end;
    end;
    TesteRisqueTiers(False);
  end;
  TOBPiece.putvalue('GP_RECALCULER', '-');
end;

{==============================================================================================}
{======================================== Affichages ==========================================}
{==============================================================================================}

procedure TFTicketFO.AfficheZonePied(Sender: TObject);
var Nam, Nam2: string;
  CC: THLabel;
  rTaxes: Double;
  {$IFDEF FOS5}
  St: string;
  LCD: TLCDLabel;
  {$ELSE}
  rPay: Double;
  {$ENDIF}
begin
  if Sender = nil then Exit;
  Nam := THNumEdit(Sender).Name;
  Nam := 'L' + Nam;
  CC := THLabel(FindComponent(Nam));
  if CC <> nil then CC.Caption := THNumEdit(Sender).Text;
  if Nam = 'LGP_TOTALTTCDEV' then
  begin
    Nam := 'LNETAPAYERDEV';
    CC := THLabel(FindComponent(Nam));
    if CC <> nil then CC.Caption := THNumEdit(Sender).Text;
    {*Nam:='PI_NETAPAYERDEV' ; LCD:=TLCDLabel(FindComponent(Nam)) ;
    if LCD<>Nil then LCD.Caption:=THNumEdit(Sender).Text ;*}
    if QuoiSurLCD = qaTotal then AfficheTotalSurLCD;
    // Affichage des taxes
    rTaxes := TOBPiece.GetValue('GP_TOTALTTCDEV') - TOBPiece.GetValue('GP_TOTALHTDEV');
    Nam2 := 'LTOTALTAXESDEV';
    CC := THLabel(FindComponent(Nam2));
    HTotalTaxesDEV.Value := rTaxes;
    if CC <> nil then CC.Caption := HTotalTaxesDEV.Text;
  end;
  {$IFNDEF FOS5}
  if ((Nam = 'LGP_TOTALTTCDEV') or (Nam = 'LGP_ACOMPTEDEV') or (Nam = 'LNETAPAYERDEV')) then
  begin
    // Affichage du net à payer
    rPay := TOBPiece.GetValue('GP_TOTALTTCDEV') - TOBPiece.GetValue('GP_ACOMPTEDEV');
    Nam2 := 'LNETAPAYERDEV';
    CC := THLabel(FindComponent(Nam2));
    NetAPayerDEV.Value := rPay;
    if CC <> nil then CC.Caption := NetAPayerDEV.Text;
  end;
  {$ENDIF}
  if ((GP_FACTUREHT.Checked) and (THNumEdit(Sender).Name = 'GP_TOTALHTDEV')) or
    ((not GP_FACTUREHT.Checked) and (THNumEdit(Sender).Name = 'GP_TOTALTTCDEV')) then
  begin
    Nam := 'PI_NETAPAYERDEV';
    LCD := TLCDLabel(FindComponent(Nam));
    St := THNumEdit(Sender).Text;
    while length(St) < LCD.NoOfChars do st := ' ' + st;
    if LCD <> nil then LCD.Caption := St;
  end;
end;

procedure TFTicketFO.GereTiersEnabled;
begin
  if FOGetParamCaisse('GPK_CLISAISIE') = 'X' then
  begin
    BZoomTiers.Enabled := ((TOBTiers.GetValue('T_TIERS') <> '') and (GP_TIERS.Text = TOBTiers.GetValue('T_TIERS')));
    BClient.Enabled := ((PieceContainer.Action <> taConsult) or (AssignePiece));
  end else
  begin
    BZoomTiers.Enabled := False;
    BClient.Visible := False;
  end;
end;

procedure TFTicketFO.GereCommercialEnabled;
var CodeVend: string;
  TOBL: TOB;
begin
  CodeVend := GP_REPRESENTANT.Text;
  if (IndiqueOuSaisie = saLigne) then
  begin
    TOBL := GetTOBLigne(TOBPiece, GS.Row);
    if TOBL <> nil then CodeVend := TOBL.GetValue('GL_REPRESENTANT');
  end;
  BZoomCommercial.Enabled := (CodeVend <> '');
end;

procedure TFTicketFO.GereEnabled(ARow: integer);
var TOBL: TOB;
  OkArt: boolean;
begin
  TOBL := GetTOBLigne(TOBPiece, ARow);
  BZoomArticle.Enabled := (GetCodeArtUnique(TOBPiece, ARow) <> '');
  if TOBL = nil then
  begin
    BZoomTarif.Enabled := False;
    BJustifTarif.Enabled := False;
    BZoomPrecedente.Enabled := False;
    BZoomOrigine.Enabled := False;
    MBSoldeReliquat.Enabled := False;
    MBDetailNomen.Enabled := False;
    MBChangeDepot.Enabled := False;
    VLigne.Enabled := False;
    BZoomStock.Enabled := False;
    TLigneRetour.Enabled := False;
    {$IFDEF TAXEUS}
    MXTaxeLigne.Enabled := False;
    MXExcepTaxe.Enabled := False;
    {$ENDIF}
  end else
  begin
    OkArt := TOBL.GetValue('GL_ARTICLE') <> '';
    BJustifTarif.Enabled := ((OkArt) and (TOBL.GetValue('GL_TARIF') > 0));
    BZoomTarif.Enabled := OkArt;
    BZoomPrecedente.Enabled := ((OkArt) and (TOBL.GetValue('GL_PIECEPRECEDENTE') <> ''));
    BZoomOrigine.Enabled := ((OkArt) and (TOBL.GetValue('GL_PIECEORIGINE') <> ''));
    MBSoldeReliquat.Enabled := ((OkArt) and (ReliquatTransfo) and (TOBL.GetValue('GL_PIECEPRECEDENTE') <> ''));
    MBDetailNomen.Enabled := ((OkArt) and (TOBL.GetValue('GL_INDICENOMEN') > 0));
    MBChangeDepot.Enabled := OkArt;
    VLigne.Enabled := OkArt;
    BSousTotal.Enabled := ((ARow > 1) and (TOBL.GetValue('GL_TYPELIGNE') <> 'TOT'));
    BZoomStock.Enabled := ((OkArt) and (TOBL.GetValue('GL_TENUESTOCK') = 'X'));
    TLigneRetour.Enabled := OkArt;
    {$IFDEF TAXEUS}
    MXTaxeLigne.Enabled := OkArt;
    MXExcepTaxe.Enabled := OkArt;
    {$ENDIF}
  end;
  BVentil.Enabled := ((VPiece.Enabled) or (VLigne.Enabled));
  Librepiece.Visible := GereStatPiece;
  MBChangeDepot.Visible := ((VH_GC.GCMultiDepots) and (PieceContainer.Action <> taConsult));
end;

procedure TFTicketFO.ZoomOuChoixLib(ACol, ARow: integer);
var TOBL: TOB;
  Contenu: TStrings;
  LigCom: string;
  i_ind, ANewRow: integer;
begin
  if ACol <> SG_Lib then Exit;
  if PieceContainer.Action = taConsult then Exit;
  if not CommentLigne then Exit;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  if GetParamSoc('SO_GCCOMMENTAIRE') then
  begin
    Contenu := GetCommentaireLigne;
    if Contenu <> nil then
    begin
      if GS.RowCount < TOBPiece.Detail.Count + Contenu.Count then GS.RowCount := GS.RowCount + Contenu.Count;
      for i_ind := 0 to Contenu.Count - 1 do
      begin
        LigCom := Contenu[i_ind];
        ANewRow := ARow + i_ind;
        if ANewRow > TOBPiece.Detail.Count then CreerTOBLignes(GS, TOBPiece, TOBTiers, TOBAffaire, ANewRow)
        else if i_ind > 0 then ClickInsert(ANewRow);
        TOBL := GetTOBLigne(TOBPiece, ANewRow);
        if TOBL = nil then Continue;
        GS.Cells[ACol, ANewRow] := LigCom;
        TOBL.PutValue('GL_LIBELLE', GS.Cells[ACol, ANewRow]);
      end;
      contenu.Free;
    end;
  end else if GereElipsis(ACol) then TOBL.PutValue('GL_LIBELLE', GS.Cells[ACol, ARow]);
end;

procedure TFTicketFO.ZoomOuChoixArt(ACol, ARow: integer);
var RefUnique: string;
  Cancel, OkArt, OkLiaison: boolean;
  TOBA, TOBL: TOB;
  NoErr: Integer;
  NewPrix, NewQte: double;
begin
  OkLiaison := False;
  NewPrix := 0;
  NewQte := 1;
  if ACol <> SG_RefArt then Exit;
  RefUnique := GetCodeArtUnique(TOBPiece, ARow);
  if RefUnique <> '' then
  begin
    TOBA := FindTOBArtRow(TOBPiece, TOBArticles, ARow);
    if FOJaiLeDroit(12) then
      ZoomArticle(RefUnique, TobA.GetValue('GA_TYPEARTICLE'), taConsult);
  end else if PieceContainer.Action <> taConsult then
  begin
    TOBL := GetTOBLigne(TOBPiece, ARow);
    if TOBL.GetValue('GL_TYPEDIM') = 'GEN' then
    begin
      AppelleDim(ARow);
    end else
    begin
      Cancel := False;
      OkArt := IdentifierArticle(ACol, ARow, Cancel, True, False);
      if ((OkArt) and (not Cancel)) then
      begin
{$IFDEF GESCOM}  // JTR - Ne pas cumuler des articles avec articles financiers dans un même ticket
        if not CumulArtOk(Acol,Arow,Cancel) then
          exit;
{$ENDIF GESCOM}
        NoErr := 0;
        case FOArticleAutorise(TOBPiece, TOBArticles, PieceContainer.Cledoc.NaturePiece, ARow) of
          afoFerme: NoErr := 2; // Article fermé
          afoInCompatible: NoErr := 45; // Article incompatible avec les autres lignes
          afoInterdit: NoErr := 46; // Article interdit à la vente
          afoConfidentiel: NoErr := -1; // Article confidentiel
        end;
        if (NoErr = 0) and (ArtFiClientOblig(nil, ARow) > 0) then
          NoErr := 51; // Client incompatible avec l'opération de caisse
        if NoErr = 0 then
        begin
          if not ReglementDispo(ARow, NewPrix, NewQte, OkLiaison) then
          begin
            Cancel := True;
            Exit;
          end;
        end;
        if NoErr <> 0 then
        begin
          if NoErr > 0 then HPiece.Execute(NoErr, Caption, '');
          VideCodesLigne(TOBPiece, ARow);
          InitialiseTOBLigne(TOBPiece, TOBTiers, TOBAffaire, ARow);
        end else
        begin
          if not OkLiaison then
          begin
            // recherche du signe de la quantité en fonction de l'article
            NewQte := FOSigneQteArticle(TOBPiece, TOBArticles, ARow, NewQte);
            // recherche de la quantité saisie sur la calculette du pavé tactile
            if PnlBoutons <> nil then
              NewQte := PnlBoutons.ChercheValeurCalculette(NewQte, True);
          end;
          InTraiteBouton := True; // pour obtenir un traitement identique à la saisie depuis le clavier
          UpdateArtLigne(ARow, False, False, NewQte, NewPrix);
{$IFDEF GESCOM}
          RemiseDuTarif(Arow); // JTR - Démarque obligatoire
{$ENDIF GESCOM} 
          // pour passer à la ligne suivante
          InTraiteBouton := False;
          StCellCur := GS.Cells[ACol, ARow];
          SurLigSuivante := True;
        end;
      end;
    end;
  end;
end;

procedure TFTicketFO.ShowDetail(ARow: integer);
var TOBArt, TOBLig: TOB;
  RefUnique, sLib: string;
  Montant: Double;
  NbDec: Integer;
  Sigle: string;
begin
  if ARow <= TOBPiece.Detail.Count then
  begin
    TOBLig := GetTOBLigne(TOBPiece, ARow);
    RefUnique := TOBLig.GetValue('GL_ARTICLE');
    sLib := TOBLig.GetValue('GL_LIBELLE');
  end else
  begin
    TOBLig := nil;
    RefUnique := '';
    sLib := '';
  end;
  //if DejaRentre then
  if (DejaSaisie) or (RefUnique <> '') or (sLib <> '') then
  begin
    TimerLCD.Enabled := False;
    if (TOBLig <> nil) and ((RefUnique <> '') or (sLib <> '')) then
    begin
      // Afficheur client
      GetTotalaAfficher(TOBLig, Montant, NbDec, Sigle);
      AffichageLCD(TOBLig.GetValue('GL_LIBELLE'), TOBLig.GetValue('GL_QTEFACT'), Montant, NbDec, Sigle, [ofInterne, ofClient], qaLigne, False);
      // Affichage du prix unitaire et du stock de l'article
      AffichePUBrut(TOBLig);
      TOBArt := TOBArticles.FindFirst(['GA_ARTICLE'], [RefUnique], False);
      A_QTESTOCK.Caption := FOMontreStockDispo(TOBArt, TOBLig);
      // Affichage de l'informations sur l'article
      LIBELLEARTICLE.Caption := FOMontreInfoArticle(TOBArt);
      // Affichage de la photo de l'article
      if FOAfficheImageArt then FOAffichePhotoArticle(RefUnique, ImageArticle);
    end else
    begin
      AfficheTotalSurLCD;
    end;
  end;
end;

procedure TFTicketFO.GetTotalaAfficher(TOBL: TOB; var Montant: double; var NbDec: integer; var Sigle: string);
begin
  if GP_FACTUREHT.Checked then
    Montant := TOBL.GetValue('GL_TOTALHTDEV')
  else
    Montant := TOBL.GetValue('GL_TOTALTTCDEV');
  Sigle := PieceContainer.DEV.Symbole;
  NbDec := PieceContainer.DEV.Decimale;
end;

procedure TFTicketFO.AffichePUBrut(TOBLig: TOB);
var Montant: Double;
begin
  if TOBLig = nil then Exit;
  // Affichage du prix unitaire de l'article
  if GP_FACTUREHT.Checked then Montant := TOBLig.GetValue('GL_PUHTDEV')
  else Montant := TOBLig.GetValue('GL_PUTTCDEV');
  T_PrixU.Caption := StrfMontant(Montant, 12, PieceContainer.DEV.Decimale, '', True);
end;

procedure TFTicketFO.AfficheTotalSurLCD;
var Montant: Double;
  Sigle: string;
  NbDec: Integer;
begin
  if TimerLCD.Enabled then Exit;
  if GP_FACTUREHT.Checked then
    Montant := TOBPiece.GetValue('GP_TOTALHTDEV')
  else
    Montant := TOBPiece.GetValue('GP_TOTALTTCDEV');
  Sigle := PieceContainer.DEV.Symbole;
  NbDec := PieceContainer.DEV.Decimale;
  AffichageLCD('Total', 0, Montant, NbDec, Sigle, [ofInterne, ofClient], qaTotal, False);
end;

{==============================================================================================}
{========================================= DEMARQUES ==========================================}
{==============================================================================================}

function TFTicketFO.DetaxeActive: boolean;
{$IFNDEF GESCOM}
var ModeDetaxe: string;
  {$ENDIF}
begin
  Result := False;
  {$IFNDEF GESCOM}
  if PieceContainer.Action <> taCreat then Exit;
  ModeDetaxe := FOGetParamCaisse('GPK_MODEDETAXE');
  if (ModeDetaxe = '') or (ModeDetaxe = '000') then Exit;
  Result := True;
  {$ENDIF}
end;

procedure TFTicketFO.TraiteDemandeDetaxe;
begin
  {$IFNDEF GESCOM}
  if not DetaxeActive then Exit;
  if not FOJaiLeDroit(6) then Exit;
  if FOGetParamCaisse('GPK_CLIOBLIGDETAXE') = 'X' then
  begin
    if GP_TIERS.Text = FOGetParamCaisse('GPK_TIERS') then
    begin
      if HPiece.Execute(52, Caption, '') = mrYes then MCChoixClientClick(nil);
    end;
    if GP_TIERS.Text = FOGetParamCaisse('GPK_TIERS') then Exit;
  end;
  DemandeDetaxe := (not DemandeDetaxe);
  LDETAXE.Visible := DemandeDetaxe;
  {$ENDIF}
end;

procedure TFTicketFO.ValideDetaxe;
begin
  {$IFNDEF GESCOM}
  if not DetaxeActive then Exit;
  if not DemandeDetaxe then Exit;
  if TOBPiece.GetValue('GP_TOTALTTCDEV') <= 0 then Exit;
  if FOGetParamCaisse('GPK_MODEDETAXE') = '001' then
  begin
    FOLanceGlobalRefund(TOBPiece, TOBTiers);
  end;
  {$ENDIF}
end;

{==============================================================================================}
{========================================= DEMARQUES ==========================================}
{==============================================================================================}

procedure TFTicketFO.TraiteDemarque(var ACol, ARow: integer; var Cancel: boolean);
var TOBL: TOB;
  sDem, NewDem: string;
  NewCol, NewRow: integer;
  Ok: boolean;
begin
  if ACol <> SG_TypeRem then Exit;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  if (TOBL.GetValue('GL_ARTICLE') = '') and (TOBL.GetValue('GL_TYPELIGNE') <> 'TOT') then Exit;
  sDem := GS.Cells[ACol, ARow];
  if ((sDem <> '') and (sDem = TOBL.GetValue('GL_TYPEREMISE'))) then Exit;
  NewCol := GS.Col;
  NewRow := GS.Row;
  GS.SynEnabled := False;
  GS.Col := ACol;
  GS.Row := ARow;
  NewDem := GS.Cells[ACol, ARow];
  //if ((LookupValueExist(GS)) or (FOGetDemarqueSaisi(GS,HTitres.Mess[32],tlWhere))) then
  ///if (FOTabletteValueExist('GCTYPEREMISE',NewDem)) or (FOGetDemarqueSaisi(GS,HTitres.Mess[32],tlWhere)) then
  Ok := False;
  if VerifDemarque(NewDem) then
  begin
    if TOBL.GetValue('GL_TYPELIGNE') = 'TOT' then
    begin
      RemiseSousTotalCorrecte(ARow, 0, 0, 0, NewDem, Cancel);
      GS.Cells[ACol, ARow] := '';
      if not Cancel then
      begin
        ReAfficheSousTotal(ARow);
        Ok := True;
      end;
    end else
    begin
      RemiseCorrecte(TOBL.GetValue('GL_REMISELIGNE'), NewDem, True, Cancel);
      if not Cancel then
      begin
        TOBL.PutValue('GL_TYPEREMISE', NewDem);
        GS.Cells[ACol, ARow] := NewDem;
        Ok := True;
      end;
    end;
  end;
  if Ok then
  begin
    GS.Col := NewCol;
    GS.Row := NewRow;
  end else
  begin
    TOBL.PutValue('GL_TYPEREMISE', '');
    GS.Cells[ACol, ARow] := '';
    Cancel := True;
  end;
  GS.SynEnabled := True;
end;

function TFTicketFO.VerifDemarque(CodeDem: string): Boolean;
begin
  Result := False;
  if (FOTabletteValueExist('GCTYPEREMISE', CodeDem, 'GTR_FERME<>"X"')) or
    (FOGetDemarqueSaisi(GS, HTitres.Mess[32], tlWhere, FideliteCli.Fidelite)) then
  begin
    // période d'utilisation du motif de démarque
    if not FODemarquePeriode(CodeDem, PieceContainer.Cledoc.DatePiece) then
    begin
      HPiece.Execute(55, Caption, '');
      Exit;
    end;
    // client obligatoire
    Result := True;
    if (DemarqueCliObl) and (FOGetParamCaisse('GPK_CLISAISIE') = 'X') then
    begin
      if FODemarqueClientOblig(CodeDem) then
      begin
        if GP_TIERS.Text = FOGetParamCaisse('GPK_TIERS') then
        begin
          if HPiece.Execute(48, Caption, '') = mrYes then MCChoixClientClick(nil);
        end;
        Result := (GP_TIERS.Text <> FOGetParamCaisse('GPK_TIERS'));
      end;
    end;
  end;
end;

{==============================================================================================}
{============================================ DEPOTS ==========================================}
{==============================================================================================}

procedure TFTicketFO.TraiteDepot(var ACol, ARow: integer; var Cancel: boolean);
var TOBL: TOB;
  sDep, NewDep: string;
  NewCol, NewRow: integer;
begin
  if ACol <> SG_Dep then Exit;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  if TOBL.GetValue('GL_ARTICLE') = '' then Exit;
  sDep := GS.Cells[ACol, ARow];
  if ((sDep <> '') and (sDep = TOBL.GetValue('GL_DEPOT'))) then Exit;
  NewCol := GS.Col;
  NewRow := GS.Row;
  GS.SynEnabled := False;
  GS.Col := ACol;
  GS.Row := ARow;
  if ((DepotExiste(sDep)) or (GetDepotSaisi(GS, HTitres.Mess[12], tlWhere))) then
  begin
    NewDep := GS.Cells[ACol, ARow];
    TOBL.PutValue('GL_DEPOT', NewDep);
    GS.Col := NewCol;
    GS.Row := NewRow;
  end else
  begin
    TOBL.PutValue('GL_DEPOT', '');
    Cancel := True;
  end;
  GS.SynEnabled := True;
end;

procedure TFTicketFO.ZoomOuChoixDep(ACol, ARow: integer);
var Depot, RefUnique, DepL: string;
  TOBL: TOB;
begin
  if ACol <> SG_Dep then Exit;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  RefUnique := TOBL.GetValue('GL_ARTICLE');
  if RefUnique = '' then Exit;
  Depot := GS.Cells[ACol, ARow];
  DepL := TOBL.GetValue('GL_DEPOT');
  if ((Depot <> '') and (Depot = DepL) and (DepotExiste(Depot))) then
  begin
    TOBL.PutValue('GL_DEPOT', Depot);
  end else if PieceContainer.Action <> taConsult then
  begin
    if IdentifierDepot(ACol, ARow) then
    begin
      StCellCur := GS.Cells[ACol, ARow];
      TOBL.PutValue('GL_DEPOT', GS.Cells[ACol, ARow]);
    end else
    begin
      TOBL.PutValue('GL_DEPOT', '');
    end;
  end;
end;

function TFTicketFO.IdentifierDepot(var ACol, ARow: integer): boolean;
var RefUnique, Depot: string;
  TOBL: TOB;
  OkD: boolean;
begin
  Result := False;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  RefUnique := TOBL.GetValue('GL_ARTICLE');
  if RefUnique = '' then Exit;
  Depot := GS.Cells[ACol, ARow];
  if Depot <> '' then OkD := DepotExiste(Depot) else OkD := False;
  if OkD then Result := True else Result := GetDepotSaisi(GS, HTitres.Mess[12], tlWhere);
end;

function TFTicketFO.DepotExiste(Depot: string): boolean;
var StD: string;
begin
  StD := RechDom('GCDEPOT', Depot, False);
  Result := ((StD <> '') and (StD <> 'Error'));
end;

procedure TFTicketFO.ChangeDepot(ARow: Integer);
var sDep, sNewDep, sPlus, sTitre: string;
  TOBL: TOB;
begin
  if PieceContainer.Action = taConsult then Exit;
  if not VH_GC.GCMultiDepots then Exit;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  if TOBL.GetValue('GL_ARTICLE') = '' then Exit;
  if not FOJaiLeDroit(10) then Exit;
  sDep := TOBL.GetValue('GL_DEPOT');
  sPlus := '';
  {$IFDEF MODE}
  sPlus := ListeDepotParEtablissement(TOBL.GetValue('GL_ETABLISSEMENT'));
  if sPlus <> '' then sPlus := ' GDE_DEPOT IN (' + sPlus + ') AND ';
  sPlus := sPlus + 'GDE_SURSITE="X"';
  {$ENDIF}
  sTitre := HTitres.Mess[12] + ' : ' + RechDom('GCDEPOT', sDep, False);
  sNewDep := FOChoisirPave(sTitre, 'GCDEPOT', sPlus);
  if (sNewDep <> '') and (sNewDep <> sDep) then
  begin
    if SG_Dep <> -1 then GS.Cells[SG_Dep, ARow] := sNewDep;
    TOBL.PutValue('GL_DEPOT', sNewDep);
  end;
end;

procedure TFTicketFO.MBChangeDepotClick(Sender: TObject);
begin
  ChangeDepot(GS.Row);
end;

{==============================================================================================}
{======================================== REPRESENTANTS =======================================}
{==============================================================================================}

procedure TFTicketFO.TraiteRepres(var ACol, ARow: integer; var Cancel: boolean);
var TOBL: TOB;
  sRep, NewRep: string;
  NewCol, NewRow: integer;
begin
  if ACol <> SG_Rep then Exit;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  if TOBL.GetValue('GL_ARTICLE') = '' then Exit;
  sRep := GS.Cells[ACol, ARow];
  if sRep = '' then
  begin
    TOBL.PutValue('GL_REPRESENTANT', '');
    Exit;
  end;
  if sRep = TOBL.GetValue('GL_REPRESENTANT') then Exit;
  if not RepresExiste(sRep) then  // JTR - eQualité 11295
  begin
    if CommercialFerme(sRep) then
    begin
      TOBL.PutValue('GL_REPRESENTANT', '');
      GS.Cells[ACol, ARow] := '';
      Cancel := True;
      Exit;
    end;
  end;
  NewCol := GS.Col;
  NewRow := GS.Row;
  GS.SynEnabled := False;
  GS.Col := ACol;
  GS.Row := ARow;
  if (FORepresentantExiste(GS.Cells[ACol, ARow], TypeCom, TOBComms)) or
    (GetRepresentantSaisi(GS, HTitres.Mess[31], TOBTiers.GetValue('T_ZONECOM'), tlWhere, TOBPiece)) then
  begin
    NewRep := GS.Cells[ACol, ARow];
    TOBL.PutValue('GL_REPRESENTANT', NewRep);
    GS.Col := NewCol;
    GS.Row := NewRow;
    AjouteRepres(NewRep, TypeCom, TOBComms);
  end else
  begin
    TOBL.PutValue('GL_REPRESENTANT', '');
    GS.Cells[ACol, ARow] := '';
    Cancel := True;
  end;
  GereCommercialEnabled;
  GS.SynEnabled := True;
end;

function TFTicketFO.RepresExiste(Repres: string): boolean;
begin
  Result := (TOBComms.FindFirst(['GCL_COMMERCIAL'], [Repres], False) <> nil);
end;

procedure TFTicketFO.ZoomOuChoixRep(ACol, ARow: integer);
var Repres, RefUnique, RepL: string;
  TOBL: TOB;
begin
  if ACol <> SG_Rep then Exit;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  RefUnique := TOBL.GetValue('GL_ARTICLE');
  if RefUnique = '' then Exit;
  Repres := GS.Cells[ACol, ARow];
  RepL := TOBL.GetValue('GL_REPRESENTANT');
  if (not V_PGI.LookUpLocate) and (Repres <> '') and (Repres = RepL) and (RepresExiste(Repres)) then
  begin
    TOBL.PutValue('GL_REPRESENTANT', Repres);
    ZoomRepres(Repres);
  end else if PieceContainer.Action <> taConsult then
  begin
    if IdentifierRepres(ACol, ARow) then
    begin
      StCellCur := GS.Cells[ACol, ARow];
      if not RepresExiste(StCellCur) then // JTR - eQualité 11295
      begin
        if CommercialFerme(StCellCur) then
        begin
          TOBL.PutValue('GL_REPRESENTANT', '');
          GS.Cells[ACol, ARow] := '';
          exit;
        end;
      end;
      TOBL.PutValue('GL_REPRESENTANT', GS.Cells[ACol, ARow]);
      AjouteRepres(GS.Cells[ACol, ARow], TypeCom, TOBComms);
    end else
    begin
      TOBL.PutValue('GL_REPRESENTANT', '');
    end;
  end;
end;

function TFTicketFO.IdentifierRepres(var ACol, ARow: integer): boolean;
var RefUnique, Repres: string;
  TOBL, TOBC: TOB;
begin
  Result := False;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  RefUnique := TOBL.GetValue('GL_ARTICLE');
  if RefUnique = '' then Exit;
  Repres := GS.Cells[ACol, ARow];
  if Repres <> '' then
    TOBC := TOBComms.FindFirst(['GCL_COMMERCIAL'], [Repres], False)
    else
    TOBC := nil;
  //if TOBC <> nil then Result := True else Result := GetRepresentantSaisi(GS, HTitres.Mess[11], TOBTiers.GetValue('T_ZONECOM'), tlWhere, TOBPiece);
  if GetRepresentantSaisi(GS, HTitres.Mess[11], TOBTiers.GetValue('T_ZONECOM'), tlWhere, TOBPiece) then
  begin
    Result := True;
  end else
  if TOBC <> nil then
  begin
    GS.Cells[ACol, ARow] := Repres;
    Result := True;
  end;
end;

procedure TFTicketFO.BalayeLignesRep;
var i: integer;
  TOBL: TOB;
  OldRep: string;
begin
  if PieceContainer.Action = taConsult then Exit;
  if GeneCharge then Exit;
  {$IFDEF FOS5}
  for i := 1 to GS.RowCount - 1 do
  begin
    TOBL := GetTOBLigne(TOBPiece, i);
    if TOBL = nil then Break;
    OldRep := TOBL.GetValue('GL_REPRESENTANT');
    if (OldRep <> GP_REPRESENTANT.Text) and (OldRep = TOBPiece.GetValue('GP_REPRESENTANT')) and (FOEstUneLigneVente(TOBL)) then
    begin
      TOBL.PutValue('GL_REPRESENTANT', GP_REPRESENTANT.Text);
      TOBL.PutValue('GL_VALIDECOM', 'NON');
{$IFDEF V500}
      CommVersLigne(TOBPiece, TOBArticles, TOBComms, i, True);
{$ELSE}
      CommVersLigne(PieceContainer, i, True);
{$ENDIF}
      AfficheLaLigne(i);
    end;
  end;
  {$ELSE}
  if GS.NbSelected <= 0 then Exit;
  for i := 1 to GS.RowCount - 1 do
  begin
    TOBL := GetTOBLigne(TOBPiece, i);
    if TOBL = nil then Break;
    if GS.isSelected(i) then
    begin
      OldRep := TOBL.GetValue('GL_REPRESENTANT');
      if OldRep <> GP_REPRESENTANT.Text then
      begin
        TOBL.PutValue('GL_REPRESENTANT', GP_REPRESENTANT.Text);
        TOBL.PutValue('GL_VALIDECOM', 'NON');
{$IFDEF V500}
        CommVersLigne(TOBPiece, TOBArticles, TOBComms, i, True);
{$ELSE}
        CommVersLigne(PieceContainer, i, True);
{$ENDIF}
      end;
    end;
  end;
  GS.ClearSelected;
  {$ENDIF}
  ShowDetail(GS.Row);
end;

procedure TFTicketFO.ZoomRepres(Repres: string);
var Crits: string;
begin
  if not FOJaiLeDroit(15) then Exit;
  Crits := ';GCL_TYPECOMMERCIAL="' + TypeCom + '"';
  {$IFDEF GESCOM}
  AGLLanceFiche('GC', 'GCCOMMERCIAL', '', Repres, 'ACTION=CONSULTATION' + Crits);
  {$ELSE}
  AGLLanceFiche('MFO', 'FOCOMMERCIAL', '', Repres, 'ACTION=CONSULTATION;MONOFICHE' + Crits);
  {$ENDIF}
end;

{==============================================================================================}
{======================================== COMPTABILITE ========================================}
{==============================================================================================}

procedure TFTicketFO.ZoomEcriture(DeStock: boolean);
begin
  {$IFNDEF SANSCOMPTA}
  LanceZoomEcriture(TOBPiece, DeStock);
  {$ENDIF}
end;

procedure TFTicketFO.GereAnal(ARow: integer);
{$IFNDEF FOS5}
var TOBL: TOB;
  RefU, OuvreA: string;
  {$ENDIF}
begin
  {$IFNDEF FOS5}
  if PieceContainer.Action = taConsult then Exit;
  if CompAnalL <> 'AUT' then Exit;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  RefU := TOBL.GetValue('GL_ARTICLE');
  if RefU = '' then Exit;
  if TOBL.FieldExists('OUVREANAL') then OuvreA := TOBL.GetValue('OUVREANAL') else OuvreA := '';
  if OuvreA = 'X' then Exit;
  {$IFDEF EAGLCLIENT}
  {AFAIREEAGL}
  {$ELSE}
  {$IFNDEF PAUL}
  YYVentilAna(TOBL, TOBL.Detail[0], PieceContainer.Action);
  {$ENDIF}
  {$ENDIF}
  TOBL.PutValue('OUVREANAL', 'X');
  {$ENDIF}
end;

procedure TFTicketFO.VPieceClick(Sender: TObject);
begin
  ClickVentil(True);
end;

procedure TFTicketFO.VLigneClick(Sender: TObject);
begin
  ClickVentil(False);
end;

procedure TFTicketFO.ClickVentil(LePied: Boolean);
var TOBL, TOBAL: TOB;
  i, iArt: integer;
begin
  if LePied then
  begin
    YYVentilAna(TOBPiece, TOBAnaP, PieceContainer.Action, 0);  // JTR - eQualité 12048
    iArt := ChampToNum('GL_ARTICLE');
    for i := 0 to TOBPiece.Detail.Count - 1 do
    begin
      TOBL := TOBPiece.Detail[i];
      if TOBL.GetValeur(iArt) = '' then Continue;
      if TOBL.FieldExists('OUVREANAL') then if TOBL.GetValue('OUVREANAL') = 'X' then Continue;
      TOBAL := TOBL.Detail[0];
      if TOBAL.Detail.Count > 0 then Continue;
      PreVentileLigne(nil, TOBAnaP, TOBAnaS, TOBL);
    end;
  end else
  begin
    TOBL := GetTOBLigne(TOBPiece, GS.Row);
    if TOBL = nil then Exit;
    YYVentilAna(TOBL, TOBL.Detail[0], PieceContainer.Action, 0); // JTR - eQualité 12048
  end;
end;

{==============================================================================================}
{===================================== Actions, boutons =======================================}
{==============================================================================================}

procedure TFTicketFO.BEcheClick(Sender: TObject);
begin
  if not BEche.Enabled then Exit;
  if not BEche.Visible then Exit;
  if GereEcheancesFO(Self, PieceContainer.Cledoc, TOBPiece, TOBEches, TobAcomptes, PieceContainer.Action) then GP_MODEREGLE.Value := TOBPiece.GetValue('GP_MODEREGLE');
end;

function TFTicketFO.ClickSousTotal: Boolean;
var ARow: integer;
  TOBL: TOB;
begin
  Result := False;
  if PieceContainer.Action = taConsult then Exit;
  if not BSousTotal.Enabled then Exit;
  if TOBPiece.Detail.Count <= 0 then Exit;
  ARow := GS.Row;
  if ARow <= 1 then Exit;
  TOBL := GetTOBLigne(TOBPiece, ARow - 1);
  if TOBL.GetValue('GL_TYPELIGNE') = 'TOT' then Exit;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL.GetValue('GL_TYPELIGNE') = 'TOT' then Exit;
  ClickInsert(ARow);
  TOBL := GetTOBLigne(TOBPiece, ARow);
  TOBL.PutValue('GL_TYPELIGNE', 'TOT');
  RecalculeSousTotaux(TOBPiece);
  Result := True;
end;

procedure TFTicketFO.BImprimerClick(Sender: TObject);
begin
  FOImprimerLaPiece(TOBPiece, TOBTiers, TOBEches, TOBComms, PieceContainer.Cledoc, False, False);
end;

procedure TFTicketFO.BSousTotalClick(Sender: TObject);
begin
  ClickSousTotal;
end;

procedure TFTicketFO.MBtarifClick(Sender: TObject);
begin
  IncidenceTarif;
end;

procedure TFTicketFO.MajDatesLivr;
var NbSel, NumQ: integer;
  i: integer;
  TOBL: TOB;
  DDLiv: TDateTime;
begin
  NbSel := GS.NbSelected;
  if NbSel > 0 then NumQ := 31 else NumQ := 30;
  if HPiece.Execute(NumQ, Caption, '') <> mrYes then Exit;
  DDLiv := StrToDate(GP_DATELIVRAISON.Text);
  for i := 0 to TOBPiece.Detail.Count - 1 do
  begin
    if ((NbSel = 0) or (GS.IsSelected(i + 1))) then
    begin
      TOBL := TOBPiece.Detail[i];
      if TOBL <> nil then TOBL.PutValue('GL_DATELIVRAISON', DDLiv);
    end;
  end;
  GS.ClearSelected;
end;

procedure TFTicketFO.MBDatesLivrClick(Sender: TObject);
begin
  MajDatesLivr;
end;

procedure TFTicketFO.SoldeReliquat;
var TOBL: TOB;
begin
  if not MBSoldeReliquat.Enabled then Exit;
  if not ReliquatTransfo then Exit;
  TOBL := GetTOBLigne(TOBPiece, GS.Row);
  if TOBL = nil then Exit;
  if TOBL.GetValue('GL_SOLDERELIQUAT') = '-' then TOBL.PutValue('GL_SOLDERELIQUAT', 'X')
  else TOBL.PutValue('GL_SOLDERELIQUAT', '-');
  GS.InvalidateRow(GS.Row);
end;

procedure TFTicketFO.ClickPorcs;
begin
  AppelPiedPort(PieceContainer.Action, TOBPiece, TOBPorcs);
  if PieceContainer.Action <> taConsult then
  begin
    TOBPiece.putvalue('GP_RECALCULER', 'X');
    CalculeLaSaisie(-1, -1, False);
  end;
  AffichePorcs;
end;

procedure TFTicketFO.BPorcsClick(Sender: TObject);
begin
  ClickPorcs;
end;

procedure TFTicketFO.BAcompteClick(Sender: TObject);
var TOBAcc: TOB;
begin
  TobAcc := Tob.create('Les acomptes', nil, -1);
  Tob.create('', TobAcc, -1);
  TobAcc.detail[0].Dupliquer(TobTiers, False, TRUE, TRUE);
  Tob.create('', TobAcc.detail[0], -1);
  TobAcc.detail[0].detail[0].Dupliquer(TobPiece, False, TRUE, TRUE);
  TheTob := TobAcc;
  TOBAcomptes.changeparent(TobAcc.detail[0].detail[0], -1);
  AGLLanceFiche('GC', 'GCACOMPTES', '', '', ActionToString(PieceContainer.Action));
  TOBAcomptes.changeparent(nil, -1);
  TobAcc.free;
  AcomptesVersPiece(PieceContainer);
end;

procedure TFTicketFO.MBSoldeReliquatClick(Sender: TObject);
begin
  SoldeReliquat;
end;

procedure TFTicketFO.BZoomPrecedenteClick(Sender: TObject);
begin
  ClickPrecedente;
end;

procedure TFTicketFO.BZoomSuivanteClick(Sender: TObject);
begin
  ClickSuivante;
end;

procedure TFTicketFO.BZoomOrigineClick(Sender: TObject);
begin
  ClickOrigine;
end;

procedure TFTicketFO.DetruitLaPiece;
var OldEcr, OldStk: RMVT;
  sSql: string;
begin
  if not FlagDetruitArriere(TOBPiece_O) then V_PGI.IoError := oeUnknown;
  if not DetruitAncien(PieceContainer, True) then
    V_PGI.IoError := oeUnknown;
  if (TOBOpCais_O <> nil) and (TOBOpCais_O.Detail.Count > 0) then
  begin
    sSql := 'DELETE FROM OPERCAISSE WHERE GOC_NATUREPIECEG="' + TOBPiece_O.GetValue('GP_NATUREPIECEG') + '"'
      + ' AND GOC_SOUCHE="' + TOBPiece_O.GetValue('GP_SOUCHE') + '"'
      + ' AND GOC_NUMERO=' + IntToStr(TOBPiece_O.GetValue('GP_NUMERO'))
      + ' AND GOC_INDICEG=' + IntToStr(TOBPiece_O.GetValue('GP_INDICEG'));
    if ExecuteSQL(sSql) <= 0 then V_PGI.IoError := oeUnknown;
  end;
  if V_PGI.IoError = oeOk then InverseStock(PieceContainer);
  if V_PGI.IoError = oeOk then DetruitCompta(TOBPiece_O, OldEcr, OldStk);
  if V_PGI.IoError = oeOk then ValideLesArticles(PieceContainer, TobPiece_O);
end;

procedure TFTicketFO.BDeleteClick(Sender: TObject);
begin
  if HPiece.Execute(29, caption, '') <> mrYes then Exit;
  if Transactions(DetruitLaPiece, 1) <> oeOk then
  begin
    MessageAlerte(HTitres.Mess[20]);
  end else
  begin
    ForcerFerme := True;
    Close
  end;
end;

procedure TFTicketFO.ClickSuivante;
var StSuiv: string;
  CD: R_CleDoc;
begin
  StSuiv := TOBPiece.GetValue('GP_DEVENIRPIECE');
  if StSuiv = '' then Exit;
  DecodeRefPiece(StSuiv, CD);
  if not ExistePiece(CD) then ShowPiecePasHisto(CD, True) else SaisieTicketFO(CD, taConsult);
end;

procedure TFTicketFO.ClickPrecedente;
var TOBL: TOB;
  StPrec: string;
  CD: R_CleDoc;
begin
  TOBL := GetTOBLigne(TOBPiece, GS.Row);
  if TOBL = nil then Exit;
  StPrec := TOBL.GetValue('GL_PIECEPRECEDENTE');
  if StPrec = '' then Exit;
  DecodeRefPiece(StPrec, CD);
  if not ExistePiece(CD) then ShowPiecePasHisto(CD, True) else SaisieTicketFO(CD, taConsult);
end;

procedure TFTicketFO.ClickOrigine;
var TOBL: TOB;
  StOri: string;
  CD: R_CleDoc;
begin
  TOBL := GetTOBLigne(TOBPiece, GS.Row);
  if TOBL = nil then Exit;
  StOri := TOBL.GetValue('GL_PIECEORIGINE');
  if StOri = '' then Exit;
  DecodeRefPiece(StOri, CD);
  if not ExistePiece(CD) then ShowPiecePasHisto(CD, False) else SaisieTicketFO(CD, taConsult);
end;

procedure TFTicketFO.ClickDel(ARow: integer; AvecC, FromUser: Boolean);
var RefUnique, CodeArticle, TypeDim: string;
  bc, cancel: boolean;
  ACol, LaLig: integer;
  TOBL: TOB;
begin
  if PieceContainer.Action = taConsult then Exit;
  if ((ARow < 1) or (ARow > TOBPiece.Detail.Count)) then Exit;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  if (FromUser) and (not FOJaiLeDroit(25)) then Exit;
  RefUnique := TOBL.Getvalue('GL_ARTICLE');
  CodeArticle := TOBL.GetValue('GL_CODESDIM');
  TypeDim := TOBL.GetValue('GL_TYPEDIM');
  // MODIF_DBR_DEBUT
  {$IFNDEF STK}
  DetruitLot(ARow);
  {$ENDIF STK}
  // MODIF_DBR_FIN
  GS.CacheEdit;
  GS.SynEnabled := False;
  LaLig := GS.TopRow;
  GS.DeleteRow(ARow);
  if (FromUser) and (LaLig > GS.FixedRows) then
    GS.TopRow := (LaLig - GS.FixedRows)
    else
    GS.TopRow := GS.FixedRows;
{  if ((TOBPiece.Detail.Count > 1) and (ARow <> TOBPiece.Detail.Count)) then
    TOBPiece.Detail[ARow - 1].Free
    else
    TOBPiece.Detail[ARow - 1].InitValeurs;
}
  if ((TOBPiece.Detail.Count > 1) and (ARow <> TOBPiece.Detail.Count)) then
  begin
    DeleteTobLigneTarif(TobPiece, ARow - 1);
    TOBPiece.Detail[ARow - 1].Free;
  end else
    TOBPiece.Detail[ARow - 1].InitValeurs;

  if GS.RowCount < NbRowsInit then GS.RowCount := GS.RowCount + NbRowsPlus;
  GS.MontreEdit;
  GS.SynEnabled := True;
  NumeroteLignesGC(GS, TOBPiece);
  bc := False;
  Cancel := False;
  ACol := GS.Col;
  GSRowEnter(GS, ARow, bc, False);
  GSCellEnter(GS, ACol, ARow, Cancel);
  if Cancel then
  begin
    GS.Col := ACol;
    GS.Row := ARow;
  end;
  ShowDetail(ARow);
  if (FromUser) and (RefUnique <> '') then
  begin
    FOIncrJoursCaisse(jfoNbLigAnnul, FOFabriqueCommentairePiece(TOBPiece, 'Annulation de l''article ' + RefUnique + ' du'));
  end;
  if (AvecC) and (RefUnique <> '') then
  begin
    PutValueDetail(TOBPiece, 'GP_RECALCULER', 'X');
    CalculeLaSaisie(-1, -1, False);
  end;
  if ((CodeArticle <> '') and (TypeDim = 'GEN') and (FromUser)) then
  begin
    repeat
      TOBL := TOBPiece.FindFirst(['GL_CODEARTICLE', 'GL_TYPEDIM'], [CodeArticle, 'DIM'], False);
      if TOBL <> nil then
      begin
        LaLig := TOBL.GetValue('GL_NUMLIGNE');
        ClickDel(LaLig, False, False);
      end;
    until TOBL = nil;
    PutValueDetail(TOBPiece, 'GP_RECALCULER', 'X');
    CalculeLaSaisie(-1, -1, False);
  end;
end;

procedure TFTicketFO.InsertComment(Ent: Boolean);
var PosDep, Num: integer;
  StComm, StC, LeCom, StLoc: string;
  TOBL: TOB;
begin
  if Ent then Num := 1 else Num := 2;
  if not FOJaiLeDroit(Num) then Exit;
  if Ent then
  begin
    StLoc := CommentEnt;
    PosDep := 1;
  end
  else
  begin
    StLoc := CommentPied;
    PosDep := TOBPiece.Detail.Count + 1;
  end;
  StComm := '';
  repeat
    StC := ReadTokenSt(StLoc);
    if StC <> '' then StComm := StC + ';' + StComm;
  until ((StLoc = '') or (StC = ''));
  if not Ent then
  begin
    GS.RowCount := GS.RowCount + 1;
    CreerTOBLignes(GS, TOBPiece, TOBTiers, TOBAffaire, TOBPiece.Detail.Count + 2);
  end;
  repeat
    StC := ReadTokenSt(StComm);
    if StC <> '' then
    begin
      if Ent then LeCom := RechDom('GCCOMMENTAIREENT', StC, False) else LeCom := RechDom('GCCOMMENTAIREPIED', StC, False);
      ClickInsert(PosDep);
      TOBL := GetTOBLigne(TOBPiece, PosDep);
      TOBL.PutValue('GL_LIBELLE', LeCom);
      AfficheLaLigne(PosDep, True);
    end;
  until ((StC = '') or (StComm = ''));
end;

procedure TFTicketFO.TCommentEntClick(Sender: TObject);
begin
  InsertComment(True);
end;

procedure TFTicketFO.TCommentPiedClick(Sender: TObject);
begin
  InsertComment(False);
end;

procedure TFTicketFO.TInsLigneClick(Sender: TObject);
begin
  ClickInsert(GS.Row);
end;

procedure TFTicketFO.TSupLigneClick(Sender: TObject);
begin
  ClickDel(GS.Row, True, True);
end;

procedure TFTicketFO.TLigneRetourClick(Sender: TObject);
begin
  RetourneArticle(GS.Row);
end;

procedure TFTicketFO.TRechArtImageClick(Sender: TObject);
begin
  {$IFDEF MODE}
  RechercheArticleImage;
  {$ENDIF}
end;

procedure TFTicketFO.MCCreerclientClick(Sender: TObject);
begin
  CreationTiers;
end;

procedure TFTicketFO.MCChoixClientClick(Sender: TObject);
begin
  ChoixTiers;
end;

procedure TFTicketFO.MCModifClientClick(Sender: TObject);
begin
  ModificationTiers;
end;

procedure TFTicketFO.MCSoldeClientClick(Sender: TObject);
begin
  VoirSoldeTiers;
end;

procedure TFTicketFO.MCArticleClientClick(Sender: TObject);
var
  Stg: string;
begin
  if not (MCArticleClient.Visible) or not (MCArticleClient.Enabled) then Exit;

  Stg := 'SELECT GP_NATUREPIECEG FROM PIECE WHERE GP_NATUREPIECEG="'+ PieceContainer.NewNature +'"'
    + ' AND GP_TIERS="'+ GP_TIERS.Text +'"';
  if ExisteSQL(Stg) then
  begin
    Stg := 'GL_TIERS='+ GP_TIERS.Text +';GL_NATUREPIECEG='+ PieceContainer.NewNature
      +';AF_LIB='+ LIBELLETIERS.Caption +';AF_CLIENT='+ HGP_TIERS.Caption;

    AGLLanceFiche('MBO', 'ARTCLI_MUL', Stg, '', 'ACTION=CONSULTATION') ;
  end else
    PGIError('Pas de vente pour ce client');
end;

procedure TFTicketFO.MCHistoClientClick(Sender: TObject);
var
  Stg: string;
begin
  if not (MCHistoClient.Visible) or not (MCHistoClient.Enabled) then Exit;
  Stg := 'SELECT GP_NATUREPIECEG FROM PIECE WHERE GP_NATUREPIECEG="'+ PieceContainer.NewNature +'"'
    + ' AND GP_TIERS="'+ GP_TIERS.Text +'"';
  if ExisteSQL(Stg) then
  begin
    Stg := 'GP_TIERS='+ GP_TIERS.Text +';GP_NATUREPIECEG='+ PieceContainer.NewNature
      +';AF_LIBELLE='+ LIBELLETIERS.Caption +';AF_CLIENT='+ HGP_TIERS.Caption;

    AGLLanceFiche('GC', 'GCVTECLI_MUL', Stg, '', 'ACTION=CONSULTATION') ;
  end else
    PGIError('Pas de vente pour ce client');
end;

procedure TFTicketFO.MCFideliteClick(Sender: TObject);
var Stg, StRetour : string;
begin
  if (not MCFidelite.Enabled) or (not MCFidelite.Visible) then Exit;

  Stg := FideliteCli.NumeroCarteInterne;
  if Stg <> '' then
  begin
    if not SortDeLaLigne then Exit;
    StRetour := AGLLanceFiche('MBO', 'FIDELITEINFO', '', '', FideliteCli.PrepareArgumentsInfoCarte(TOBTiers, TOBPiece));
    if StRetour = 'X' then
    begin
      FideliteCli.CumulFidelite := 0; //Force le recalcul du cumul
      FideliteCli.Fidelite;
    end;
  end;
end;

procedure TFTicketFO.MCContactClientClick(Sender: TObject);
var
  Stg: string;
begin
  if (not MCContactClient.Enabled) or (not MCContactClient.Visible) then Exit;

  if ((PieceContainer.Action = taConsult) and (not AssignePiece)) or (not FOJaiLeDroit(35, False, False)) then
  begin
    if not FOJaiLeDroit(13) then Exit;
    Stg := ActionToString(taConsult);
  end else
  begin
    Stg := ActionToString(taModif);
  end;

  Stg := Stg + ';TYPE=T;' + 'TYPE2=' + TobTiers.GetValue('T_NATUREAUXI')
    + ';PART=' + TobTiers.GetValue('T_PARTICULIER') + ';TITRE=' + TobTiers.GetValue('T_LIBELLE')
    + ';TIERS=' + TobTiers.GetValue('T_TIERS') + ';ALLCONTACT';
  AGLLanceFiche('YY', 'YYCONTACT', 'T;'+TobTiers.GetValue('T_AUXILIAIRE'), '', Stg);
end;

procedure TFTicketFO.MDChoixModeleClick(Sender: TObject);
begin
  ToucheModele('', '', '', 0, 0);
end;

procedure TFTicketFO.MDSituFlashClick(Sender: TObject);
begin
  if FOJaiLeDroit(103) then
    FOConsultationCaisse(FOCaisseCourante, '', '', nil, True); // Situation Flash
end;

procedure TFTicketFO.MDDetaxeClick(Sender: TObject);
begin
  TraiteDemandeDetaxe;
end;

procedure TFTicketFO.MDALivrerClick(Sender: TObject);
begin
  TicketALivrer;
end;

procedure TFTicketFO.MDOuvreTiroirClick(Sender: TObject);
begin
  if (FOExisteTiroir) and (FOJaiLeDroit(23)) then FOOuvretiroir(True); // Ouverture du tiroir caisse
end;

procedure TFTicketFO.MDReimpTicketClick(Sender: TObject);
begin
  FOReImpTicket(False); // Réimpression du dernier ticket
end;

procedure TFTicketFO.MDReimpBonsClick(Sender: TObject);
begin
  FOReImprimeBons; // Réimpression des bons du dernier ticket
end;

procedure TFTicketFO.ClickInsert(ARow: integer);
begin
  if PieceContainer.Action = taConsult then Exit;
  if ((ARow < 1) or (ARow > TOBPiece.Detail.Count)) then Exit;
  GS.CacheEdit;
  GS.SynEnabled := False;
  InsertTOBLigne(TOBPiece, ARow);
  GS.InsertRow(ARow);
  InitialiseTOBLigne(TOBPiece, TOBTiers, TOBAffaire, ARow);
  GS.Row := ARow;
  GS.MontreEdit;
  GS.SynEnabled := True;
  NumeroteLignesGC(GS, TOBPiece);
  ShowDetail(ARow);
end;

procedure TFTicketFO.BZoomArticleClick(Sender: TObject);
var RefUnique: string;
  TOBA: TOB;
begin
  if not BZoomArticle.Enabled then Exit;
  if not FOJaiLeDroit(12) then Exit;
  RefUnique := GetCodeArtUnique(TOBPiece, GS.Row);
  TOBA := FindTOBArtRow(TOBPiece, TOBArticles, GS.Row);
  ZoomArticle(RefUnique, TobA.GetValue('GA_TYPEARTICLE'), taConsult);
end;

procedure TFTicketFO.BZoomStockClick(Sender: TObject);
var
  RefUnique, sDepot: string;
  TOBL: TOB;
  {$IFDEF GESCOM}
  RefFour, Client, Four: string;
  Contrem: boolean;
  {$ELSE}
  sSql: string;
  {$ENDIF}
begin
  if not BZoomStock.Enabled then Exit;
  RefUnique := GetCodeArtUnique(TOBPiece, GS.Row);
  if RefUnique = '' then Exit;
  TOBL := GetTOBLigne(TOBPiece, GS.Row);
  if TOBL = nil then Exit;
  if not FOJaiLeDroit(18) then Exit;
  {$IFDEF GESCOM}
  if VenteAchat = 'ACH' then
  begin
    Four := TOBPiece.GetValue('GP_TIERS');
    Client := GetChampLigne(TOBPiece, 'GL_FOURNISSEUR', GS.Row);
  end else
  begin
    Client := TOBPiece.GetValue('GP_TIERS');
    Four := GetChampLigne(TOBPiece, 'GL_FOURNISSEUR', GS.Row);
  end;
  RefFour := GetChampLigne(TOBPiece, 'GL_REFCATALOGUE', GS.Row);
  sDepot := GetChampLigne(TOBPiece, 'GL_DEPOT', GS.Row);
  Contrem := (GetChampLigne(TOBPiece, 'GL_ENCONTREMARQUE', GS.Row) = 'X');
  ZoomDispo(RefUnique, RefFour, Client, Four, sDepot, Contrem);
  {$ELSE}
  sDepot := TOBL.GetValue('GL_DEPOT');
  sSql := 'SELECT * FROM DISPO WHERE GQ_ARTICLE="' + RefUnique + '" AND GQ_DEPOT="' + sDepot + '" '
    + 'AND GQ_CLOTURE="-" AND GQ_DATECLOTURE="01/01/1900"';
  if FOExisteSQL(sSql) then AGLLanceFiche('MFO', 'FODISPO', '', RefUnique + ';' + sDepot + ';-;01/01/1900', 'ACTION=CONSULTATION')
  else HPiece.Execute(44, Caption, '');
  {$ENDIF}
end;

procedure TFTicketFO.BMenuZoomMouseEnter(Sender: TObject);
begin
  if csDestroying in ComponentState then Exit;
  PopZoom97(BMenuZoom, POPZ);
end;

procedure ConsultationStock(Distant: boolean);
{$IFNDEF GESCOM}
var Num: integer;
  {$ENDIF}
begin
  {$IFNDEF GESCOM}
  if Distant then Num := 20 else Num := 19;
  if not FOJaiLeDroit(Num) then Exit;
  if Distant then // Consultation du stock à distance
    {$IFNDEF EAGLCLIENT}
    AGLLanceFiche('MBO', 'DISPODIST_MUL', '', '', 'MULDISTANT')
      {$ENDIF}
  else // Consultation du stock local
    AGLLanceFiche('MBO', 'DISPODIST_MUL', '', '', 'MULLOCAL');
  {$ENDIF}
end;

procedure TFTicketFO.MDStockClick(Sender: TObject);
begin
  ConsultationStock(False); // Consultation du stock local
end;

procedure TFTicketFO.ClickDevise;
begin
  if GP_DEVISE.Value = '' then Exit;
  {$IFNDEF SANSCOMPTA}
  if not FOJaiLeDroit(16) then Exit;
  FicheDevise(GP_DEVISE.Value, taConsult, False);
  {$ENDIF}
end;

procedure TFTicketFO.BZoomDeviseClick(Sender: TObject);
begin
  ClickDevise;
end;

procedure TFTicketFO.BJustifTarifClick(Sender: TObject);
var TOBL: TOB;
  sArg, sCodeArt: string;
  dMont: Double;
begin
  if not BJustifTarif.Enabled then Exit;
  if not FOJaiLeDroit(14) then Exit;
  TOBL := GetTOBLigne(TOBPiece, GS.Row);
  if TOBL = nil then Exit;
  sCodeArt := TOBL.GetValue('GL_CODEARTICLE');
  if sCodeArt = '' then Exit;
  if GP_FACTUREHT.Checked then dMont := TOBL.GetValue('GL_PUHTDEV')
  else dMont := TOBL.GetValue('GL_PUTTCDEV');
  sArg := 'CODEARTICLE=' + CodeArticleUnique2(sCodeArt, '')
    + ';TARIFARTICLE=' + TOBL.GetValue('GL_TARIFARTICLE')
    + ';PRIXVENTEART=' + FloatToStr(dMont)
    + ';DEPOT=' + GP_ETABLISSEMENT.Text
    + ';DATE=' + GP_DATEPIECE.Text;
  AGLLanceFiche('MBO', 'JUSTIFPXVENTE', '', '', sArg);
end;

procedure TFTicketFO.BZoomTarifClick(Sender: TObject);
var TOBL: TOB;
  sArg, sCodeArt: string;
  dMont: Double;
begin
  if not BZoomTarif.Enabled then Exit;
  if not FOJaiLeDroit(14) then Exit;
  TOBL := GetTOBLigne(TOBPiece, GS.Row);
  if TOBL = nil then Exit;
  sCodeArt := TOBL.GetValue('GL_CODEARTICLE');
  if sCodeArt = '' then Exit;
  if GP_FACTUREHT.Checked then dMont := TOBL.GetValue('GL_PUHTDEV')
  else dMont := TOBL.GetValue('GL_PUTTCDEV');
  sArg := 'CODEARTICLE=' + CodeArticleUnique2(sCodeArt, '')
    + ';TARIFARTICLE=' + TOBL.GetValue('GL_TARIFARTICLE')
    + ';PRIXVENTE=' + FloatToStr(dMont)
    + ';ETABLISSEMENT=' + GP_ETABLISSEMENT.Value
    + ';DATE=' + GP_DATEPIECE.Text;
  AGLLanceFiche('MBO', 'TARFPXVENTE', '', '', sArg);
end;

procedure TFTicketFO.BZoomCommercialClick(Sender: TObject);
var CodeVend: string;
  TOBL: TOB;
begin
  CodeVend := GP_REPRESENTANT.Text;
  if (IndiqueOuSaisie = saLigne) then
  begin
    TOBL := GetTOBLigne(TOBPiece, GS.Row);
    if TOBL <> nil then CodeVend := TOBL.GetValue('GL_REPRESENTANT');
  end;
  if CodeVend <> '' then ZoomRepres(CodeVend);
end;

procedure TFTicketFO.BZoomTiersClick(Sender: TObject);
begin
  if not FOJaiLeDroit(13) then Exit;
  if ctxAffaire in V_PGI.PGIContexte then
    V_PGI.DispatchTT(8, taConsult, TOBTiers.GetValue('T_TIERS'), '', '')
      {$IFDEF MODE}
  else AGLLanceFiche('MFO', 'GCTIERSFO', '', TOBTiers.GetValue('T_AUXILIAIRE'), 'ACTION=CONSULTATION');
  {$ELSE}
  else AGLLanceFiche('GC', 'GCTIERS', '', TOBTiers.GetValue('T_AUXILIAIRE'), 'ACTION=CONSULTATION');
  {$ENDIF}
end;

procedure TFTicketFO.BZoomEcritureClick(Sender: TObject);
begin
  ZoomEcriture(False);
end;

procedure TFTicketFO.CpltEnteteClick(Sender: TObject);
begin
  ComplementEntete;
end;

procedure TFTicketFO.MBDetailNomenClick(Sender: TObject);
var TOBL, TOBNomen: TOB;
  IndiceNomen: integer;
begin
  TOBL := GetTOBLigne(TOBPiece, GS.Row);
  if TOBL = nil then Exit;
  IndiceNomen := TOBL.GetValue('GL_INDICENOMEN');
  if IndiceNomen <= 0 then Exit;
  TOBNomen := TOBNomenclature.Detail[IndiceNomen - 1];
  if TOBNomen = nil then Exit;
  if not FOJaiLeDroit(17) then Exit;
  Entree_LigneNomen(TOBNomen, TobArticles, TOBL, False, 1, 0);
  RenseigneValoNomen(TOBL, TOBNomen);
end;

procedure TFTicketFO.CpltLigneClick(Sender: TObject);
begin
  ComplementLigne(GS.Row);
end;

procedure TFTicketFO.AdrLivClick(Sender: TObject);
begin
  EntreeAdressePiece(TobPiece, TobAdresses, 0);
end;

procedure TFTicketFO.AdrFacClick(Sender: TObject);
begin
  EntreeAdressePiece(TobPiece, TobAdresses, 1);
end;

procedure TFTicketFO.OnGetVar(Sender: TObject; VarName: string; VarIndx: Integer; var Value: variant);
begin
  if VarName = 'NBL' then Value := TOBPiece.Detail.Count else
    if VarName = 'GP_BLOCNOTE' then
  begin
    BlobToFile(VarName, TOBPiece.GetValue(VarName));
    Value := '';
  end else //PCS
    if Pos('T_', VarName) > 0 then Value := TOBTiers.GetValue(VarName) else // Paul
    if Pos('GP_', VarName) > 0 then Value := TOBPiece.GetValue(VarName) else
  begin
    if VarName = 'GL_BLOCNOTE' then
    begin
      BlobToFile(VarName, TOBPiece.Detail[VarIndx - 1].GetValue(VarName));
      Value := '';
    end else //PCS
      Value := TOBPiece.Detail[VarIndx - 1].GetValue(VarName);
  end;
end;

procedure TFTicketFO.BOfficeClick(Sender: TObject);
{$IFNDEF EAGLCLIENT}
var stDocWord: string;
  {$ENDIF}
begin
  {$IFDEF EAGLCLIENT}
  {AFAIREEAGL}
  {$ELSE}
  stDocWord := 'C:\DEVISCLIENT.DOC';
  DeleteFile(stDocWord);
  ConvertDocFile('C:\DEVIS.DOC', stDocWord, nil, nil, nil, OnGetVar, nil, nil);
  ShellExecute(0, PCHAR('open'), PChar(stDocWord), nil, nil, SW_RESTORE); //PCS
  {$ENDIF}
end;

procedure TFTicketFO.BAideClick(Sender: TObject);
begin
  CallHelpTopic(Self);
end;

procedure TFTicketFO.LibrepieceClick(Sender: TObject);
begin
  ValideTablesLibres(True);
end;

procedure TFTicketFO.BTicketAttenteClick(Sender: TObject);
begin
  if PieceContainer.Action = taConsult then Exit;
  if FOGetParamCaisse('GPK_GERETICKETATT') <> 'X' then Exit;
  if not FOJaiLeDroit(11) then Exit;
  // pré-positionnement dans la grille
  if IndiqueOuSaisie = saEntete then if GS.CanFocus then GS.SetFocus;
  if IndiqueOuSaisie <> saLigne then Exit;
  if TobTicketAttente <> nil then FreeAndNil(TobTicketAttente);
  if FOReprisePieceEnAttente(TOBPiece, GS, DejaSaisie, TobTicketAttente, TobLigFormule) then
  begin
  if not DejaSaisie then
    begin
      // affichage de l'en-tête
      TOBPiece.PutEcran(Self);
      if not ChargeTiers then
        if GP_TIERS.CanFocus then GP_TIERS.SetFocus;
      DejaSaisie := True;
    end;
    // recalcul de la pièce
    LoadLesArticles;
    DispoChampSupp(TOBArticles);
    CalculeLaSaisie(-1, -1, True);
  end;
end;

procedure TFTicketFO.EnvoieToucheGrid;
var Key: Word;
  TOBL: TOB;
begin
  TOBL := GetTOBLigne(TOBPiece, GS.Row);
  if TOBL = nil then Exit;
  Key := 0;
  if GS.Col = SG_RefArt then
  begin
    if FOGetParamCaisse('GPK_LIGNESUIVAUTO') = 'X' then
    begin
      if TOBL.GetValue('GL_ARTICLE') <> '' then
      begin
        if TOBL.GetValue('GL_PUHTDEV') <> 0 then
          Key := VK_DOWN
        else
          Key := VK_TAB;
      end else
        if TOBL.GetValue('GL_LIBELLE') <> '' then Key := VK_TAB;
    end else
      Key := VK_TAB;
  end else
    if GS.Col = SG_Rep then
  begin
    if TOBL.GetValue('GL_REPRESENTANT') <> '' then
      Key := VK_TAB;
  end;
  if Key > 0 then FOSimuleClavier(Key);
end;

{==============================================================================================}
{========================================== TAXATION ==========================================}
{==============================================================================================}

procedure TFTicketFO.ChangeTVA;
var NewCode: string;
  i, NbSel: integer;
  Okok: Boolean;
  TOBL: TOB;
begin
  if PieceContainer.Action = taConsult then Exit;
  if not GP_FACTUREHT.Checked then Exit;
  if TOBPiece.Detail.Count <= 0 then Exit;
  if not FOJaiLeDroit(5) then Exit;
  if FOExisteClavierEcran then
    NewCode := FOChoisirPave(HTitres.Mess[28], 'GCFAMILLETAXE1', '', clBtnFace, 6, True, 0, '')
  else
    NewCode := Choisir(HTitres.Mess[28], 'CHOIXCOD', 'CC_LIBELLE', 'CC_CODE', 'CC_TYPE="TX1"', 'CC_CODE');
  if NewCode = '' then Exit;
  if HPiece.Execute(38, Caption, '') <> mrYes then Exit;
  NbSel := GS.NbSelected;
  Okok := False;
  for i := 0 to TOBPiece.Detail.Count - 1 do
  begin
    if ((NbSel = 0) or (GS.IsSelected(i + 1))) then
    begin
      TOBL := TOBPiece.Detail[i];
      if TOBL.GetValue('GL_ARTICLE') = '' then Continue;
      if TOBL.GetValue('GL_FAMILLETAXE1') <> NewCode then
      begin
        TOBL.PutValue('GL_FAMILLETAXE1', NewCode);
        TOBL.putValue('GL_RECALCULER', 'X');
        Okok := True;
      end;
    end;
  end;
  if Okok then
  begin
    PutValueDetail(TOBPiece, 'GP_RECALCULER', 'X'); // positionne le recalcul du document
    CalculeLaSaisie(-1, -1, True);
  end;
end;

procedure TFTicketFO.ChangeRegime;
var NewCode: string;
begin
  if PieceContainer.Action = taConsult then Exit;
  if not FOJaiLeDroit(4) then Exit;
  if (TOBPiece.GetValue('GP_TOTALTTCDEV') = 0) and (TOBPiece.GetValue('GP_TOTALQTEFACT') = 0) then
  begin
    NewCode := FOChoisirPave(HTitres.Mess[33], 'TTREGIMETVA', '', clBtnFace, 6, True, 0, GP_REGIMETAXE.Value);
    if (NewCode <> '') and (NewCode <> GP_REGIMETAXE.Value) then
    begin
      ChangeComplEntete := True;
      TOBPiece.PutValue('GP_REGIMETAXE', NewCode);
      TOBPiece.PutEcran(Self, PEntete);
      GP_REGIMETAXEChange(nil);
      ChangeComplEntete := False;
    end;
  end;
  {*****
  if FOExisteClavierEcran then
     NewCode := FOChoisirPave(HTitres.Mess[33], 'TTREGIMETVA', '', clBtnFace, 6, True, 0, GP_REGIMETAXE.Value)
  else
     NewCode := Choisir(HTitres.Mess[33],'CHOIXCOD','CC_LIBELLE','CC_CODE','CC_TYPE="RTV"','CC_CODE') ;
  if (NewCode = '') or (NewCode = GP_REGIMETAXE.Value) then Exit ;
  if TOBPiece.Detail.Count <= 0 then
     BEGIN
     ChangeComplEntete := True;
     TOBPiece.PutValue('GP_REGIMETAXE', NewCode) ;
     TOBPiece.PutEcran(Self, PEntete) ;
     GP_REGIMETAXEChange(Nil) ;
     ChangeComplEntete := False;
     END else
     BEGIN
     if HPiece.Execute(39, Caption, '') <> mrYes then Exit ;
     GP_REGIMETAXE.Value := NewCode ;
     PutValueDetail(TOBPiece, 'GP_RECALCULER', 'X') ;
     CalculeLaSaisie(-1, -1, True) ;
     END ;
  *****}
end;

///////////////////////////////////////////////////////////////////////////////////////
//  TrouveRegimePiece : détermine le régime de taxe de la pièce
///////////////////////////////////////////////////////////////////////////////////////

procedure TFTicketFO.TrouveRegimePiece;
var Stg: string;
begin
  Stg := FOGetParamCaisse('REGIMETAXECAISSE'); // régime de la caisse
  if Stg = '' then Stg := VH^.RegimeDefaut; // régime par défaut du dossier
  GP_REGIMETAXE.Value := Stg;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ExceptionTaxe : saisie d'une exception de taxation
///////////////////////////////////////////////////////////////////////////////////////
{$IFDEF TAXEUS}
procedure TFTicketFO.ExceptionTaxe;
var
  OldCode, NewCode, sModele, Stg: string;
  TOBL: TOB;
  NoExec, NoModele: integer;
begin
  if PieceContainer.Action = taConsult then Exit;
  if TOBPiece.Detail.Count < 1 then Exit;

  NoExec := TOBPiece.Detail[0].GetNumChamp('GL_EXCEPTIONTAXE');
  NoModele := TOBPiece.Detail[0].GetNumChamp('GL_CODEMODELETAXE');

  TOBL := GetTobLigne(TobPiece, GS.Row);
  if TOBL <> nil then
  begin
    OldCode := TOBL.GetValeur(NoExec);
    sModele := TOBL.GetValeur(NoModele);
  end else
  begin
    if TOBPiece.FieldExists('GP_EXCEPTIONTAXE') then
      OldCode := TobPiece.GetValue('GP_EXCEPTIONTAXE')
    else
      OldCode := '';
    sModele := TOBL.GetValue('GP_CODEMODELETAXE');
  end;

  Stg := 'GXA_CODEMODELETAXE="'+ sModele +'"';
  NewCode := FOChoisirPave(HTitres.Mess[36], 'GCEXCEPTIONTAXE', Stg, clBtnFace, 6, True, 0, OldCode);
  if (NewCode <> '') and (NewCode <> OldCode) then
  begin
    Stg := TraduireMemoire('Exception de taxation') +' : '+ RechDom('GCEXCEPTIONTAXE', NewCode, False);
    if (sModele = TOBPiece.GetValue('GP_CODEMODELETAXE')) and
      (PGIAsk('Voulez-vous appliquer cette exception de taxation sur toutes les lignes de la pièce ?', Stg) = mrYes) then
    begin
      ChangeExceptionTaxePiece(NewCode, TOBPiece, TOBTiers, TOBAdresses, TOBArticles, VenteALivrer);
    end else
    if TOBL <> nil then
    begin
      TOBL.PutValeur(NoExec, NewCode);
      if IncidenceTaxeLigne(GS.Row, False) then TOBPiece.PutValue('GP_RECALCULER', 'X');
    end;
    if TOBPiece.GetValue('GP_RECALCULER') = 'X' then CalculeLaSaisie(-1, -1, False);
  end;
end;
{$ENDIF}

///////////////////////////////////////////////////////////////////////////////////////
//  TaxationPiece : saisie des paramètres de taxation de la pièce
///////////////////////////////////////////////////////////////////////////////////////
{$IFDEF TAXEUS}
procedure TFTicketFO.TaxationPiece;
var
  OldCode, NewCode: string;
begin
  if PieceContainer.Action = taConsult then Exit;
  OldCode := TOBPiece.GetValue('GP_CODEMODELETAXE');

  // Choix du modèle de taxe
  NewCode := FOChoisirPave(HTitres.Mess[37], 'GCMODELETAXE', '', clBtnFace, 6, True, 0, OldCode);
  if (NewCode <> '') and (NewCode <> OldCode) then
  begin
    ChangeModeleTaxePiece(NewCode, TOBPiece, TOBTiers, TOBAdresses, TOBArticles, VenteALivrer);
    if TOBPiece.GetValue('GP_RECALCULER') = 'X' then CalculeLaSaisie(-1, -1, False);
  end;
end;
{$ENDIF}

///////////////////////////////////////////////////////////////////////////////////////
//  TaxationLigne : saisie des paramètres de taxation de la ligne
///////////////////////////////////////////////////////////////////////////////////////
{$IFDEF TAXEUS}
procedure TFTicketFO.TaxationLigne;
var
  Stg, sArgs: string;
  TOBL: TOB;
begin
  TOBL := GetTOBLigne(TOBPiece, GS.Row);
  if TOBL = nil then Exit;

  sArgs := ActionToString(PieceContainer.Action);
  TheTOB := TOBL;
  Stg := FOLanceTaxationLigne('', '', sArgs);
  TheTOB := nil;
  if Stg = 'OK' then
  begin
    IncidenceTaxeLigne(GS.Row, False);
    TOBPiece.PutValue('GP_RECALCULER', 'X');
    CalculeLaSaisie(-1, -1, False);
  end;
end;
{$ENDIF}

procedure TFTicketFO.MXExcepTaxeClick(Sender: TObject);
begin
{$IFDEF TAXEUS}
  ExceptionTaxe;
{$ENDIF}
end;

procedure TFTicketFO.MXTaxePieceClick(Sender: TObject);
begin
{$IFDEF TAXEUS}
  TaxationPiece;
{$ENDIF}
end;

procedure TFTicketFO.MXTaxeLigneClick(Sender: TObject);
begin
{$IFDEF TAXEUS}
  TaxationLigne;
{$ENDIF}
end;

///////////////////////////////////////////////////////////////////////////////////////
//  IncidenceTaxeLigne : incidence de la taxation d'une ligne
///////////////////////////////////////////////////////////////////////////////////////
{$IFDEF TAXEUS}
function TFTicketFO.IncidenceTaxeLigne(ARow: integer; ForceModele: boolean): boolean;
var
  TOBL, TOBA: TOB;
begin
  Result := False;
  TOBL := GetTobLigne(TobPiece, ARow);
  if TOBL = nil then Exit;
  TOBA := FindTOBArtRow(TOBPiece, TOBArticles, ARow);
  if TOBA = nil then Exit;

  TaxeVersLigne(TOBPiece, TOBTiers, TOBAdresses, TOBL, TOBA, VenteALivrer, ForceModele);
  TOBL.PutValue('GL_RECALCULER', 'X');
  //AfficheLaLigne(ARow);
  Result := True;
end;
{$ENDIF}

///////////////////////////////////////////////////////////////////////////////////////
//  IncidenceTaxe : incidence de la taxation sur l'ensemble des lignes
///////////////////////////////////////////////////////////////////////////////////////
{$IFDEF TAXEUS}
procedure TFTicketFO.IncidenceTaxe;
var
  Ind: integer;
  Ok: boolean;
begin
  Ok := False;
  for Ind := 0 to TOBPiece.Detail.Count -1 do
    if IncidenceTaxeLigne(Ind +1, True) then Ok := True;

  if Ok then
  begin
    TOBPiece.PutValue('GP_RECALCULER', 'X');
    CalculeLaSaisie(-1, -1, False);
  end;
end;
{$ENDIF}

{=======================================================================================}
{===============================Gestion descriptif détail===============================}
{=======================================================================================}

procedure TFTicketFO.TDescriptifClose(Sender: TObject);
begin
  GereDescriptif(GS.Row, False);
  BDescriptif.Down := False;
end;

procedure TFTicketFO.BDescriptifClick(Sender: TObject);
begin
  GereDescriptif(GS.Row, True);
  TDescriptif.Visible := BDescriptif.Down;
end;

procedure TFTicketFO.GereDescriptif(ARow: Integer; Enter: Boolean);
var TOBL: TOB;
  st: string;
begin
  if not (BDescriptif.Down) then Exit;
  TOBL := GetTobLigne(TobPiece, ARow);
  if TOBL = nil then Exit;
  if not FOJaiLeDroit(9) then
  begin
    BDescriptif.Down := False;
    Exit;
  end;
  if Enter then
  begin
    StringToRich(Descriptif, TOBL.GetValue('GL_BLOCNOTE'));
    TDescriptif.Caption := Format('Descriptif ligne N° %d', [ARow]);
  end
  else
  begin
    st := Descriptif.Text;
    if (Length(st) <> 0) and (st <> #$D#$A) then
    begin
      TOBL.PutValue('GL_BLOCNOTE', RichToString(Descriptif));
    end
    else
    begin
      if (TOBL.GetValue('GL_BLOCNOTE') <> '') then TOBL.PutValue('GL_BLOCNOTE', '');
    end;
  end;
end;

{==============================================================================================}
{============================== Contrôles d'intégrité  ========================================}
{==============================================================================================}
///////////////////////////////////////////////////////////////////////////////////////
//  FOPieceCorrecte : Contrôle d'intégrité spécifique au Front Office
///////////////////////////////////////////////////////////////////////////////////////

function TFTicketFO.FOPieceCorrecte(TOBPiece, TOBArticles: TOB): Integer;
const erPasArticle = 2;
  erVendeur = 21;
  erMotif = 22;
  erVdIncorrect = 23;
  erClient = 24;
  erCliInterdit = 25;
  erMixeArt = 26;
  erRegleDejaLie = 29;
  erRegleLie = 30;
  erFidelite = 31;
var NoLig, NoCol, NoErr: Integer;
  sVendeur, sClient: string;
  PQ, CA: Double;
  TOBL: TOB;
  OkArt: boolean;
begin
  Result := 0;
  // Vérification du client
  //if TOBPiece.GetValue('GP_TIERS') <> TOBTiers.GetValue('T_TIERS') then
  sClient := TOBPiece.GetValue('GP_TIERS');
  if ((sClient = '') or (sClient <> TOBTiers.GetValue('T_TIERS')) or
     (RemplirTOBTiers(TOBTiers, sClient, PieceContainer.NewNature, True) <> trtOk)) then
  begin
    Result := erClient;
    if GP_TIERS.CanFocus then GP_TIERS.SetFocus else GotoEntete;
    Exit;
  end;
  // Vendeur obligatoire sur l'entête
  if VH_GC.TOBPCaisse.GetValue('GPK_VENDSAISIE') = 'X' then
  begin
    sVendeur := TOBPiece.GetValue('GP_REPRESENTANT');
    if sVendeur = '' then
    begin
      if VH_GC.TOBPCaisse.GetValue('GPK_VENDOBLIG') = 'X' then Result := erVendeur;
    end else
    begin
      if not FORepresentantExiste(sVendeur, TypeCom, TOBComms) then Result := erVdIncorrect;
    end;
    if Result <> 0 then
    begin
      if GP_REPRESENTANT.CanFocus then GP_REPRESENTANT.SetFocus else GotoEntete;
      Exit;
    end;
  end;
  // verification des lignes
  CA := 0;
  NoLig := 0;
  OkArt := False;
  while (NoLig < TOBPiece.Detail.Count) and (Result = 0) do
  begin
    TOBL := TOBPiece.Detail[NoLig];
    if TOBL.GetValue('GL_ARTICLE') <> '' then OkArt := True;
    if FOEstUneLigneVente(TOBL) then
    begin
      // Vendeur obligatoire sur les lignes
      if VH_GC.TOBPCaisse.GetValue('GPK_VENDSAISLIG') = 'X' then
      begin
        sVendeur := TOBL.GetValue('GL_REPRESENTANT');
        if sVendeur = '' then
        begin
          if VH_GC.TOBPCaisse.GetValue('GPK_VENDOBLIG') = 'X' then Result := erVendeur;
        end else
        begin
          if not FORepresentantExiste(sVendeur, TypeCom, TOBComms) then Result := erVdIncorrect;
        end;
      end;
      // Motif de remise obligatoire
      if FOGereDemarque(True) then
      begin
{$IFDEF GESCOM}
        // JTR - Démarque obligatoire. Test si la remise présente est celle du tarif
        if (TOBL.FieldExists('REMISEDUTARIF')) and   
           (TOBL.GetValue('GL_TOTREMLIGNEDEV') <> 0) and (TOBL.GetValue('GL_TYPEREMISE') = '') and
           (TOBL.GetValue('REMISEDUTARIF') <> TOBL.GetValue('GL_REMISELIGNE')) then Result := erMotif;
{$ELSE GESCOM}
        if (TOBL.GetValue('GL_TOTREMLIGNEDEV') <> 0) and (TOBL.GetValue('GL_TYPEREMISE') = '') then Result := erMotif;
{$ENDIF GESCOM}
      end;
      // Client Obligatoire
      if (DemarqueCliObl) and (TOBL.GetValue('GL_TYPEREMISE') <> '') then
      begin
        if FODemarqueClientOblig(TOBL.GetValue('GL_TYPEREMISE')) then
        begin
          if not VerifClientOblig(48) then Result := erClient;
        end;
      end;
      if TOBL.GetValue('GL_TYPEARTICLE') = 'FI' then
      begin
        // Client obligatoire ou interdit
        Result := ArtFiClientOblig(TOBL, NoLig);
      end else
      begin
        // Calcul CA (hors opération de caisse)
        if GP_FACTUREHT.Checked then CA := CA + TOBL.GetValue('GL_MONTANTHT')
        else CA := CA + TOBL.GetValue('GL_MONTANTTTC');
      end;
    end;
    // Prix pour quantité
    PQ := TOBL.GetValue('GL_PRIXPOURQTE');
    if PQ <= 0 then
    begin
      PQ := 1.0;
      TOBL.PutValue('GL_PRIXPOURQTE', PQ);
    end;
    Inc(NoLig);
  end;
  // Ticket sans article
  if (Result = 0) and not (OkArt) then
  begin
    Result := erPasArticle;
    NoLig := GS.FixedRows;
  end;
  // Pour ne pas mixer certains opérations financières avec des articles
  if (Result = 0) and not (FOVerifMixageArticle(TOBPiece, TOBArticles, NoLig)) then
  begin
    Result := erMixeArt;
    Inc(NoLig, GS.FixedRows);
  end;
  // Vérification de la cohérence des règlements liés
  if Result = 0 then
  begin
    NoErr := FOVerifReglementLie(TOBPiece, TOBEches, TOBArticles, AnnulPiece, NoLig);
    if NoErr <> 0 then
    begin
      if NoErr = 2 then Result := erRegleLie
      else Result := erRegleDejaLie;
      Inc(NoLig, GS.FixedRows);
    end;
  end;
  // Client obligatoire si le CA dépasse le seuil
  if (Result = 0) and (FOSeuilCliOblig(CA)) then
  begin
    if not VerifClientOblig(50) then Result := erClient;
  end;
  // Fidelité
  if (Result = 0) then Result := FideliteCli.FideliteCorrecte(TOBPiece);
  if Result <> 0 then
  begin
    if Result in [erClient, erCliInterdit] then
    begin
      GoToEntete;
    end else
    begin
      case Result of
        erVendeur,
          erVdIncorrect: NoCol := SG_Rep;
        erMotif: NoCol := SG_TypeRem;
        erMixeArt: NoCol := SG_RefArt;
        erRegleLie: NoCol := SG_Lib;
        erFidelite:
          begin
            NoLig := 1;
            NoCol := SG_TypeRem;
          end;
      else NoCol := SG_RefArt;
      end;
      AllerAuControle(saLigne, NoLig, NoCol, False, True);
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ReglementDispo : Recherche des règlements disponibles
///////////////////////////////////////////////////////////////////////////////////////

function TFTicketFO.ReglementDispo(ARow: integer; var Montant, Quantite: double; var Ok: boolean): boolean;
var
  TypeOpe: string;
  TOBL, TOBA: TOB;
begin
  Result := True;
  Montant := 0;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if (TOBL = nil) or (TOBL.GetValue('GL_TYPEARTICLE') <> 'FI') then
  begin
    if TOBL.FieldExists('GOC_NUMPIECELIEN') then
      TOBL.DelChampSup('GOC_NUMPIECELIEN', False);
    Exit;
  end;
  TOBA := FindTOBArtRow(TOBPiece, TOBArticles, ARow);
  TypeOpe := TOBA.GetValue('GA_TYPEARTFINAN');
  Result := FORechercheReglementDispo(TOBL, TOBPiece, TOBEches, TypeOpe, Montant, Ok);
  if Ok then
    Quantite := FOSigneQteArticle(TOBPiece, TOBArticles, ARow, Quantite);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ArtFiClientOblig : vérifie si le code du client est obligatoire pour l'article financier
///////////////////////////////////////////////////////////////////////////////////////

function TFTicketFO.ArtFiClientOblig(TOBLig: TOB; ARow: Integer): Integer;
const erClient = 24;
  erCliInterdit = 25;
var Stg: string;
  TOBArt: TOB;
begin
  Result := 0;
  if TOBLig = nil then TOBLig := GetTOBLigne(TOBPiece, ARow);
  if TOBLig = nil then Exit;
  if TOBLig.GetValue('GL_TYPEARTICLE') <> 'FI' then Exit;
  TOBArt := FindTOBArtRow(TOBPiece, TOBArticles, ARow);
  // Client obligatoire ou interdit
  Stg := FOOpCaisseCliOblig(TOBLig, TOBArt);
  if (Stg = 'OBL') and not (VerifClientOblig(49)) then
    Result := erClient;
  if (Stg = 'INT') and (GP_TIERS.Text <> FOGetParamCaisse('GPK_TIERS')) then
    Result := erCliInterdit;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  VerifClientOblig : vérifie si le code du client est obligatoire
///////////////////////////////////////////////////////////////////////////////////////

function TFTicketFO.VerifClientOblig(IdMsg: Integer): Boolean;
begin
  Result := True;
  if FOGetParamCaisse('GPK_CLISAISIE') = '-' then Exit;
  if GP_TIERS.Text = FOGetParamCaisse('GPK_TIERS') then
  begin
    if HPiece.Execute(IdMsg, Caption, '') = mrYes then MCChoixClientClick(nil);
  end;
  Result := ((GP_TIERS.Text <> '') and (GP_TIERS.Text <> FOGetParamCaisse('GPK_TIERS')));
end;

///////////////////////////////////////////////////////////////////////////////////////
//  RemiseSousTotalCorrecte : Contrôle et répartition de la remise accordée sur un sous-total
///////////////////////////////////////////////////////////////////////////////////////

procedure TFTicketFO.RemiseSousTotalCorrecte(ARow: integer; TauxRemise, RemiseSousTotal, MontantNet: double; TypeRemise: string; var Cancel: boolean);
var MaxRem, NoErr: integer;
begin
  NoErr := FOAppliqueRemiseSousTotal(TOBPiece, ARow, TauxRemise, RemiseSousTotal, MontantNet, TypeRemise, PieceContainer.DEV, MaxRem);
  if NoErr <> 0 then Cancel := True;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  RemiseCorrecte : Contrôle de la remise accordée sur une ligne
///////////////////////////////////////////////////////////////////////////////////////

procedure TFTicketFO.RemiseCorrecte(TauxRemise: double; TypeRemise: string; TypeUniquement: boolean; var Cancel: boolean);
var MaxRem, NoErr: integer;
  CodeDemValide: string;
begin
  CodeDemValide := '';
  NoErr := FOVerifMaxRemise(TauxRemise, TypeRemise, TypeUniquement, MaxRem, CodeDemValide);
  if NoErr <> 0 then Cancel := True;
end;

{==============================================================================================}
{============================== Saisie des écéhances ==========================================}
{==============================================================================================}
///////////////////////////////////////////////////////////////////////////////////////
//  AppelEcheancesFO : appel de la saisie des échéances
///////////////////////////////////////////////////////////////////////////////////////

procedure TFTicketFO.AppelEcheancesFO;
var ResGC: integer;
begin
  // Tests et actions préalables
  if not BValider.Enabled then Exit;
  if PieceContainer.Action = taConsult then Exit;
  if not SortDeLaLigne then Exit;
  TOBPiece.GetEcran(Self, PEntete);
  if not PieceModifiee then Exit;
  FODepileTOBGrid(GS, TOBPiece, [SG_RefArt, SG_Lib]);
  DepileTOBLignes(GS, TOBPiece, GS.Row, 1);
  if TesteRisqueTiers(True) then Exit;
  if TesteMargeMini(GS.Row) then Exit;
  // Contrôle intégrité
  ResGC := GCPieceCorrecte(PieceContainer);
  if ResGC > 0 then
  begin
    HErr.Execute(ResGC, Caption, '');
    Exit;
  end;
  // Contrôle intégrité spécifique au Front Office
  ResGC := FOPieceCorrecte(TOBPiece, TOBArticles);
  if ResGC > 0 then
  begin
    HErr.Execute(ResGC, Caption, '');
    Exit;
  end;
  if (Not FideliteCli.FideliteDansPiece) and (FideliteCli.Fidelite) then
  begin
    if HPiece.Execute(56, '', FideliteCli.MessageFidelite) = mrYes then Exit;
  end;
  // Tables libres
  if not ValideTablesLibres(False) then Exit;
  // Saisie des échéances
  AllerAuControle(saEcheance, 1, 1, False, False);
  if GereEcheancesFO(Self, PieceContainer.Cledoc, TOBPiece, TOBEches, TobAcomptes, PieceContainer.Action) then GP_MODEREGLE.Value := TOBPiece.GetValue('GP_MODEREGLE');
end;

{=======================================================================================}
{======================================= VALIDATIONS ===================================}
{=======================================================================================}

function TFTicketFO.CreerReliquat: boolean;
var TOBRel, BasesRel, NomenRel, EchesRel, AnaPRel, AnaSRel, AcomptesRel: TOB;
  PieceContainerRel: TPieceContainer;
  i: integer;
  TOBL, TOBLOld, TOBLRel, PorcsRel: TOB;
  NewInd: integer;
  Ecart: double;
  NumL: integer;
  OldEcr, OldStk: RMVT;
  NowFutur: TDateTime;
begin
  Result := False;
  PieceContainerRel := TPieceContainer.Create;
  try
    TOBRel := TOB.Create('PIECE', nil, -1);
    TOBRel.Dupliquer(TOBPiece_O, False, True);
    PorcsRel := TOB.Create('PORCS', nil, -1);
    PorcsRel.Dupliquer(TOBPorcs_O, False, True);
    NomenRel := TOB.Create('', nil, -1);
    NomenRel.Dupliquer(TOBN_O, True, True);
    AcomptesRel := TOB.Create('', nil, -1);
    ReliquatAcomptes(TOBAcomptes, TOBAcomptes_O, AcomptesRel);
    PieceContainerRel.TCPiece := TobRel;
    PieceContainerRel.TCPorcs := PorcsRel;
    PieceContainerRel.TCNomenclature := NomenRel;
    PieceContainerRel.TCAcomptes := AcomptesRel;
    PieceContainerRel.TCTiers := TobTiers;
    PieceContainerRel.TCArticles := TOBArticles;
    FillChar(OldEcr, SizeOf(OldEcr), #0);
    FillChar(OldStk, SizeOf(OldStk), #0);
    PieceContainerRel.Dev := PieceContainer.Dev;
    NumL := 0;
    NowFutur := NowH;
    FillChar(OldEcr, SizeOf(OldEcr), #0);
    for i := 0 to TOBPiece.Detail.Count - 1 do
    begin
      TOBL := TOBPiece.Detail[i];
      if TOBL.GetValue('GL_SOLDERELIQUAT') = 'X' then Continue;
      TOBLOld := GetTOBPrec(TOBL, TOBPiece_O);
      if TOBLOld = nil then Continue;
      Ecart := Arrondi(TOBLOld.GetValue('GL_QTESTOCK') - TOBL.GetValue('GL_QTESTOCK'), 6);
      if Ecart > 0 then
      begin
        TOBLRel := TOB.Create('LIGNE', TOBRel, -1);
        AddLesSupLigne(TOBLRel, True);
        TOBLRel.Dupliquer(TOBLOld, True, True);
        Inc(NumL);
        TOBLRel.PutValue('GL_QTESTOCK', Ecart);
        TOBLRel.PutValue('GL_QTEFACT', Ecart);
        TOBLRel.PutValue('GL_QTERELIQUAT', 0);
        TOBLRel.PutValue('GL_QTERESTE', 0);
        TOBLRel.PutValue('GL_VALIDECOM', 'AFF');
        TOBLRel.PutValue('GL_PIECEPRECEDENTE', EncodeRefPiece(TOBLOld));
        TOBLOld.PutValue('GL_QTERESTE', Ecart);
      end;
    end;
    if NumL <= 0 then
    begin
      TOBRel.Free;
      PorcsRel.Free;
      NomenRel.Free;
      AcomptesRel.Free;
      Exit;
    end;
    NumeroteLignesGC(nil, TOBRel);
    NewInd := TOBPiece_O.GetValue('GP_INDICEG') + 1;
    TOBRel.PutValue('GP_DEVENIRPIECE', '');
    PutValueDetail(TOBRel, 'GP_INDICEG', NewInd);
    PutValueDetail(TOBRel, 'GP_VIVANTE', 'X');
    PutValueDetail(TOBRel, 'GP_RECALCULER', 'X');
    // Création des TOB utiles à la nouvelle pièce
    BasesRel := TOB.Create('Les BASES', nil, -1);
    EchesRel := TOB.Create('Les ECHEANCES', nil, -1);
    EchesRel.Dupliquer(TOBEches, True, True);
    AnaPRel := TOB.Create('', nil, -1);
    AnaPRel.Dupliquer(TOBAnaP, True, True);
    AnaSRel := TOB.Create('', nil, -1);
    AnaSRel.Dupliquer(TOBAnaS, True, True);
    // Calculs
    CalculFacture(PieceContainer);
    ValideLesLignes(PieceContainer, True, False);
    GereEcheancesFO(Self, PieceContainer.Cledoc, TOBRel, EchesRel, AcomptesRel, taModif);
    // Tout modif
    TOBRel.SetAllModifie(True);
    TOBRel.SetDateModif(NowFutur);
    BasesRel.SetAllModifie(True);
    NomenRel.SetAllModifie(True);
    AcomptesRel.SetAllModifie(True);
    // Analytiques
    ValideAnalytiques(TOBRel, AnaPRel, AnaSRel);
    // Comptabilité
    if not PassationComptable(PieceContainerRel, PieceContainerRel.DEV.Decimale, OldEcr, OldStk, True)
      then V_PGI.IoError := oeLettrage;
    // Enregistrements
    ValideLesAcomptes(PieceContainerRel);
    if not TOBRel.InsertDBByNivel(False) then V_PGI.IoError := oeUnknown;
    if V_PGI.IoError = oeOk then if not BasesRel.InsertDB(nil) then V_PGI.IoError := oeUnknown;
    if V_PGI.IoError = oeOk then
    begin
      for i := NomenRel.Detail.Count - 1 downto 0 do
        if NomenRel.Detail[i].GetValue('UTILISE') <> 'X' then NomenRel.Detail[i].Free;
      if not NomenRel.InsertDB(nil) then V_PGI.IoError := oeUnknown;
    end;
    if V_PGI.IoError = oeOk then AnaPRel.InsertDB(nil);
    if V_PGI.IoError = oeOk then AnaSRel.InsertDBTable(nil);
    // Pièce origine
    if V_PGI.IoError = oeOk then TOBPiece_O.UpdateDB(False);
    // Frees
    TOBRel.Free;
    TOBRel := nil;
    BasesRel.Free;
    NomenRel.Free;
    EchesRel.Free;
    EchesRel := nil;
    AnaPRel.Free;
    AnaSRel.Free;
    AcomptesRel.Free;
    AcomptesRel := nil;
  finally
    PieceContainerRel.Free;
  end;
  Result := True;
end;

procedure TFTicketFO.GereLesReliquats;
begin
  if PieceContainer.Action <> taModif then Exit;
  if not PieceContainer.TransfoPiece then Exit;
  if not ReliquatTransfo then Exit;
  if CreerReliquat then ;
end;

procedure TFTicketFO.ValideLaCompta(OldEcr, OldStk: RMVT);
var ActiveCompta: boolean;
    TobOpCaisse : TOB;
begin
  ActiveCompta := FOComptaEstActive(PieceContainer.Cledoc.NaturePiece);
  if (ActiveCompta <> True) then exit;
{$IFDEF GESCOM}
  if IsArtFiPce(TobPiece) then
  begin
    TobOpCaisse := TOB.Create('PIECE', nil, -1);
    TobOpCaisse.Dupliquer(PieceContainer.TCPiece,true,true,true);
  end else
    TobOpCaisse := nil;
{$ENDIF GESCOM}
  if not PassationComptable(PieceContainer, PieceContainer.DEV.Decimale, OldEcr, OldStk, True, TobOpCaisse) then
    V_PGI.IoError := oeLettrage;
  if TobOpCaisse <> nil then FreeAndNil(TobOpCaisse);
end;

// MODIF_DBR_DEBUT
{$IFNDEF STK}
procedure TFTicketFO.ValideLesLots;
var i, IndiceLot: integer;
  TOBDL, TOBL: TOB;
begin
  if not GereLot then Exit;
  if TOBDesLots.Detail.Count <= 0 then Exit;
  TOBDesLots.Detail[0].AddChampSup('UTILISE', True);
  for i := 0 to TOBPiece.Detail.Count - 1 do
  begin
    IndiceLot := GetChampLigne(TOBPiece, 'GL_INDICELOT', i + 1);
    if IndiceLot > 0 then
    begin
      TOBL := TOBPiece.Detail[i];
      TOBDL := TOBDesLots.Detail[IndiceLot - 1];
      TOBDL.PutValue('UTILISE', 'X');
      UpdateNoeudLot(TOBL, TOBDL);
    end;
  end;
  for i := TOBDesLots.Detail.Count - 1 downto 0 do
  begin
    TOBDL := TOBDesLots.Detail[i];
    if TOBDL.GetValue('UTILISE') <> 'X' then TOBDL.Free;
  end;
  if TOBDesLots.Detail.Count > 0 then TOBPiece.PutValue('GP_ARTICLESLOT', 'X');
  if not TOBDesLots.InsertDB(nil) then V_PGI.IoError := oeUnknown;
end;

procedure TFTicketFO.ValideLesSeries;
var i, IndiceSerie: integer;
  TOBSerLig, TOBL: TOB;
begin
  {$IFDEF CCS3}
  Exit;
  {$ELSE}
  if not GereSerie then Exit;
  if TOBSerie.Detail.Count <= 0 then Exit;
  TOBSerie.Detail[0].AddChampSup('UTILISE', True);
  for i := 0 to TOBPiece.Detail.Count - 1 do
  begin
    IndiceSerie := GetChampLigne(TOBPiece, 'GL_INDICESERIE', i + 1);
    if IndiceSerie > 0 then
    begin
      TOBL := TOBPiece.Detail[i];
      TOBSerLig := TOBSerie.Detail[IndiceSerie - 1];
      TOBSerLig.PutValue('UTILISE', 'X');
      UpdateLesSeries(TOBL, TOBSerLig);
      MAJDispoSerie(TOBL, TOBSerLig, False);
    end;
  end;
  for i := TOBSerie.Detail.Count - 1 downto 0 do
  begin
    TOBSerLig := TOBSerie.Detail[i];
    if TOBSerLig.GetValue('UTILISE') <> 'X' then TOBSerLig.Free;
  end;
  //if TOBSerie.Detail.Count>0 then TOBPiece.PutValue('GP_ARTICLESLOT','X') ;
  if not TOBSerie.InsertDBTable(nil) then V_PGI.IoError := oeUnknown;
  {$ENDIF}
end;
{$ENDIF STK}
// MODIF_DBR_FIN

procedure TFTicketFO.ValideTiers;
begin
  if PieceContainer.Action <> taCreat then if ((not PieceContainer.DuplicPiece) and (not PieceContainer.TransfoPiece) and (not AssignePiece)) then Exit;
  // pas de mise à jour du tiers par defaut de la caisse
  if GP_TIERS.Text = FODonneClientDefaut then Exit;
  ValideleTiers(PieceContainer);
end;

procedure TFTicketFO.ValideAdresses;
begin
  // pas de mise à jour des adresses pour le tiers par defaut de la caisse
  if (not GereAdresse) or (not VenteALivrer) or (GP_TIERS.Text = FODonneClientDefaut) then
  begin
    TOBPiece.PutValue('GP_NUMADRESSELIVR', 0);
    TOBPiece.PutValue('GP_NUMADRESSEFACT', 0);
  end else
  begin
    if V_PGI.IoError = oeOk then ValideLesAdresses(PieceContainer);
  end;
end;

procedure TFTicketFO.ValideLaNumerotation;
var ind, NewNum: integer;
  NaturePieceG: string;
begin
  if ((PieceContainer.Action = taCreat) or (PieceContainer.TransfoPiece) or (PieceContainer.DuplicPiece)) then
  begin
    NaturePieceG := TOBPiece.GetValue('GP_NATUREPIECEG');
    if not SetDefinitiveNumber(PieceContainer, PieceContainer.Cledoc.NumeroPiece) then
      V_PGI.IoError := oePointage;
    //if Not FOSetNumeroDefinitif(TOBPiece,TOBBases,TOBEches,TOBNomenclature, TOBAcomptes) then V_PGI.IoError:=oePointage ;
    //if Not SetNumeroDefinitif(TOBPiece,TOBBases,TOBEches,TOBNomenclature,TOBAcomptes) then V_PGI.IoError:=oePointage ;
    if GetInfoParPiece(NaturePieceG, 'GPP_ACTIONFINI') = 'ENR' then TOBPiece.PutValue('GP_VIVANTE', '-');
    NewNum := TOBPiece.GetValue('GP_NUMERO');
    GP_NUMEROPIECE.Caption := FTitrePiece.Caption + IntToStr(NewNum);
    PieceContainer.Cledoc.NumeroPiece := NewNum;
    Ind := TOBPiece.GetValue('GP_INDICEG');
    if Ind > 0 then GP_NUMEROPIECE.Caption := FTitrePiece.Caption + GP_NUMEROPIECE.Caption + ' ' + IntToStr(Ind);
  end;
end;

procedure TFTicketFO.ValideLesNumerosLignes;
var
  Ind, InN, InS, InDt, InNo, InI, NoErr: integer;
  CD: R_CleDoc;
  TOBL: TOB;
begin
  CD := TOB2CleDoc(TOBPiece);
  if TOBPiece.Detail.Count <= 0 then
  begin
    LogAGL(DateTimeToStr(Now) + ' Ticket n°' + IntToStr(CD.NumeroPiece) + ' vide');
    V_PGI.IoError := oeUnknown;
    Exit;
  end;
  InN := 0;
  InS := 0;
  InNo := 0;
  InI := 0;
  InDt := 0;
  for Ind := 0 to TOBPiece.Detail.Count - 1 do
  begin
    NoErr := 0;
    TOBL := TOBPiece.Detail[Ind];
    if Ind = 0 then
    begin
      InN := TOBL.GetNumChamp('GL_NATUREPIECEG');
      InS := TOBL.GetNumChamp('GL_SOUCHE');
      InNo := TOBL.GetNumChamp('GL_NUMERO');
      InI := TOBL.GetNumChamp('GL_INDICEG');
      InDt := TOBL.GetNumChamp('GL_DATEPIECE');
    end;
    if TOBL.GetValeur(InN) <> CD.NaturePiece then
    begin
      TOBL.PutValeur(InN, CD.NaturePiece);
      NoErr := 1;
    end;
    if TOBL.GetValeur(InS) <> CD.Souche then
    begin
      TOBL.PutValeur(InS, CD.Souche);
      NoErr := 2;
    end;
    if TOBL.GetValeur(InNo) <> CD.NumeroPiece then
    begin
      TOBL.PutValeur(InNo, CD.NumeroPiece);
      NoErr := 3;
    end;
    if TOBL.GetValeur(InI) <> CD.Indice then
    begin
      TOBL.PutValeur(InI, CD.Indice);
      NoErr := 4;
    end;
    if TOBL.GetValeur(InDt) <> CD.DatePiece then
    begin
      TOBL.PutValeur(InDt, CD.DatePiece);
      NoErr := 7;
    end;
    if NoErr > 0 then
      LogAGL(DateTimeToStr(Now) + ' ERREUR ValideLesNumerosLignes n°' + IntToStr(NoErr)
        + ' sur Ticket n°' + IntToStr(CD.NumeroPiece) + ' Ligne n°' + IntToStr(Ind + 1));
    TOBL.SetAllModifie(True);
  end;
  NumeroteLignesGC(nil, TOBPiece);
end;

procedure TFTicketFO.DelOrUpdateAncien;
var PasHisto: Boolean;
  NowFutur: TdateTime;
  sSql: string;
begin
  NowFutur := NowH;
  TOBPiece_O.SetDateModif(NowFutur);
  TOBPiece_O.CleWithDateModif := True;
  PasHisto := (GetInfoParPiece(TOBPiece_O.GetValue('GP_NATUREPIECEG'), 'GPP_HISTORIQUE') = '-');
  if ((not PieceContainer.TransfoPiece) or (PasHisto)) then
  begin
    if not DetruitAncien(PieceContainer) then
      V_PGI.IoError := oeSaisie;
    if (V_PGI.IoError = oeOk) and (TOBOpCais_O <> nil) and (TOBOpCais_O.Detail.Count > 0) then
    begin
      sSql := 'DELETE FROM OPERCAISSE WHERE GOC_NATUREPIECEG="' + TOBPiece_O.GetValue('GP_NATUREPIECEG') + '"'
        + ' AND GOC_SOUCHE="' + TOBPiece_O.GetValue('GP_SOUCHE') + '"'
        + ' AND GOC_NUMERO=' + IntToStr(TOBPiece_O.GetValue('GP_NUMERO'))
        + ' AND GOC_INDICEG=' + IntToStr(TOBPiece_O.GetValue('GP_INDICEG'));
      if ExecuteSQL(sSql) <= 0 then V_PGI.IoError := oeUnknown;
    end;
  end else
  begin
    if not UpdateAncien(PieceContainer) then V_PGI.IoError := oeSaisie;
  end;
end;

function TFTicketFO.SortDeLaLigne: boolean;
var ACol, ARow: integer;
  Cancel: boolean;
begin
  Result := False;
  if GP_ESCOMPTE.Focused then
  begin
    if GP_REMISEPIED.CanFocus then GP_REMISEPIED.SetFocus;
    GP_ESCOMPTEExit(nil);
  end;
  if GP_REMISEPIED.Focused then
  begin
    if GP_ESCOMPTE.CanFocus then GP_ESCOMPTE.SetFocus;
    GP_REMISEPIEDExit(nil);
  end;
  if GP_DATEPIECE.Focused then
  begin
    if not ExitDatePiece then Exit;
    if GP_ETABLISSEMENT.CanFocus then GP_ETABLISSEMENT.SetFocus;
  end;
  if not DejaRentre then
  begin
    Result := True;
    Exit;
  end;
  if Screen.ActiveControl <> GS then
  begin
    Result := True;
    Exit;
  end;
  Result := False;
  ACol := GS.Col;
  ARow := GS.Row;
  Cancel := False;
  GSCellExit(nil, ACol, ARow, Cancel);
  if Cancel then Exit;
  GSRowExit(nil, ARow, Cancel, False);
  if Cancel then Exit;
  Result := True;
end;

procedure TFTicketFO.InverseAvoir;
begin
  InverseLesPieces(TOBPiece, 'PIECE');
  InverseLesPieces(TOBBases, 'PIEDBASE');
  InverseLesPieces(TOBEches, 'PIEDECHE');
  InverseLesPieces(TOBPorcs, 'PIEDPORT'); //mcd 07/06/02 port non pris en compte
end;

procedure TFTicketFO.ValideLaPiece;
var OldEcr, OldStk: RMVT;
  Caisse: string;
  NumZ: Integer;
  {$IFDEF CHR}
  Dossier,Dosres:String;
  i: Integer;
  TOBW:TOB;
  {$ENDIF}
begin
  if (PieceContainer.Action = taConsult) and (not AnnulPiece) and (not AssignePiece) then
  begin
    Close;
    Exit;
  end;
  InitMove(8, '');
  PieceContainer.InitTobs(Self);
  VenteALivrer := ((GetInfoParPiece(TOBPiece.GetValue('GP_NATUREPIECEG'),'GPP_MAJINFOTIERS') ='X') and
                   (TOBPiece.GetValue('GP_NUMADRESSELIVR') > 0));
  // MODIF_DBR_DEBUT
  {$IFNDEF STK}
  ReaffecteDispoDiff(TOBArticles); //JD 09/10/2003
  {$ENDIF STK}
  // MODIF_DBR_FIN
  {Traitement annulation de la pièce}
  if AnnulPiece then
  begin
    PieceContainer.Cledoc.DatePiece := V_PGI.DateEntree;
    PutValueDetail(TOBPiece, 'GP_DATEPIECE', PieceContainer.Cledoc.DatePiece);
    FOPutValueDetail(TOBEches, 'GPE_DATEPIECE', PieceContainer.Cledoc.DatePiece);
    PutValueDetail(TOBPiece, 'GP_DATECREATION', Date);
    TOBPiece.PutValue('GP_HEURECREATION', NowH);
    Caisse := FOCaisseCourante;
    NumZ := FOGetNumZCaisse(FOCaisseCourante);
    TOBPiece.PutValue('GP_CAISSE', Caisse);
    PutValueDetail(TOBPiece, 'GP_CAISSE', Caisse);
    FOPutValueDetail(TOBEches, 'GPE_CAISSE', Caisse);
    TOBPiece.PutValue('GP_NUMZCAISSE', NumZ);
    FOPutValueDetail(TOBEches, 'GPE_NUMZCAISSE', NumZ);
    EstAvoir := True;
    FOMarqueChqDifUtilise(TOBEches);
    FOInsereLigneCommentaire(TOBPiece, 'Annule le', False);
    TOBPiece.PutValue('GP_TICKETANNULE', 'X');
  end;
  if PieceContainer.DuplicPiece then
  begin
    PutValueDetail(TOBPiece, 'GP_RECALCULER', 'X');
    CalculeLaSaisie(-1, -1, False);
  end;
  if EstAvoir then InverseAvoir;
  MoveCur(False);
  {Traitement pièce d'origine}
  if (AssignePiece) or ((PieceContainer.Action = taModif) and (not PieceContainer.DuplicPiece)) then
  begin
    if V_PGI.IoError = oeOk then DetruitCompta(TOBPiece_O, OldEcr, OldStk);
    if V_PGI.IoError = oeOk then DelOrUpdateAncien;
    if V_PGI.IoError = oeOk then InverseStock(PieceContainer);
  end;
  {Mise à jour de compteurs de journée}
  if AnnulPiece then FOIncrJoursCaisse(jfoNbTicAnnul, FOFabriqueCommentairePiece(TOBPiece, 'Annulation du'));
  if AssignePiece then FOIncrJoursCaisse(jfoNbRatcli, FOFabriqueCommentairePiece(TOBPiece, ''));
  MoveCur(False);
  {Enregistrement nouvelle pièce}
  if V_PGI.IoError = oeOk then
  begin
    InitToutModif;
    ValideLaNumerotation;
    ValideLaCotation(PieceContainer);
    ValideLaPeriode(TOBPiece);
    if GetInfoParPiece(PieceContainer.NewNature, 'GPP_IMPIMMEDIATE') = 'X' then TOBPiece.PutValue('GP_EDITEE', 'X');
  end;
  MoveCur(False);
  if V_PGI.IoError = oeOk then ValideLesLignes(PieceContainer, False, False);
  if V_PGI.IoError = oeOk then ValideAdresses;
  if V_PGI.IoError = oeOk then GereLesReliquats;
  MoveCur(False);
//if V_PGI.IoError = oeOk then ValideLesArticles(PieceContainer, TobPiece);
  if V_PGI.IoError = oeOk then ValideLesNumerosLignes;
{$IFDEF STK}
  if EstAvoir then
    STK_ValideLesArticles(apsSuppression, PieceContainer)
    else
    STK_ValideLesArticles(apsCreation, PieceContainer);
{$ELSE STK}
  if V_PGI.IoError = oeOk then ValideLesArticles(PieceContainer, TobPiece);
{$ENDIF STK}
  if (V_PGI.IoError = oeOk) and ((TOBAnaP.detail.count > 0) or (TOBAnaS.detail.count > 0)) then
    ValideAnalytiques(TOBPiece, TOBAnaP, TOBAnaS);
  if V_PGI.IoError = oeOk then ValideTiers;
  MoveCur(False);
  if V_PGI.IoError = oeOk then ValideLaCompta(OldEcr, OldStk);
  // MODIF_DBR_DEBUT
  {$IFNDEF STK}
  if V_PGI.IoError = oeOk then ValideLesLots;
  if V_PGI.IoError = oeOk then ValideLesSeries;
  {$ENDIF STK}
  if V_PGI.IoError = oeOk then ValideLesFormules(TOBPiece,TOBLigFormule); // JTR - eQualité 11203 (formules de qté)
  // MODIF_DBR_FIN
//  if V_PGI.IoError = oeOk then ValideLesAcomptes(PieceContainer);
  if (V_PGI.IoError = oeOk) and (PieceContainer.TCAcomptes.detail.count > 0) then ValideLesAcomptes(PieceContainer);
//  if V_PGI.IoError = oeOk then ValideLesPorcs(PieceContainer);
  if (V_PGI.IoError = oeOk) and (PieceContainer.TCPorcs.detail.count > 0) then ValideLesPorcs(PieceContainer);
  MoveCur(False);
  {$IFNDEF CHR}
  if (PieceContainer.Action = taCreat) or (AnnulPiece) then
  begin
    if V_PGI.IoError = oeOk then
      FOValideLesOperCaisse(TOBPiece, TOBArticles, TOBEches, TOBPiece_O, AnnulPiece);
  end else
  if AssignePiece then
  begin
    if V_PGI.IoError = oeOk then
    begin
      if not TOBOpCais.InsertDB(nil) then V_PGI.IoError := oeUnknown;
    end;
  end;
  {$ENDIF}
  MoveCur(False);
  {$IFDEF CHR}
  if V_PGI.IoError=oeOk then
  begin
    Dossier:='FFO'+(formatdatetime('yyyymmdd',V_PGI.DateEntree));
    Dosres:=CreerDossierFO(Dossier,'',GetParamSoc('SO_HRRESSFFO'),'1','300','0',Tobhrdossier,False);
    TobPiece.Putvalue('GP_HRDOSSIER', Dosres);
    EcheVersPiece (TOBEches,TOBPiece) ;
    for i:=0 to TOBEches.Detail.Count-1 do
    Begin
      TOBW:=TOBEches.Detail[i];
      TOBW.PutValue('GPE_NOFOLIO', TOBPiece.Detail[0].Getvalue('GL_NOFOLIO')) ;
    end;
  end;
  {$ENDIF}
  if V_PGI.IoError = oeOk then
  begin
    if not TOBPiece.InsertDBByNivel(False) then V_PGI.IoError := oeUnknown;
  end;
  MoveCur(False);
  if V_PGI.IoError = oeOk then
  begin
    if not TOBBases.InsertDB(nil) then V_PGI.IoError := oeUnknown;
  end;
  if V_PGI.IoError = oeOk then
  begin
    if not TOBEches.InsertDB(nil) then V_PGI.IoError := oeUnknown;
  end;
  if V_PGI.IoError = oeOk then
  begin
    if not TOBAnaP.InsertDB(nil) then V_PGI.IoError := oeUnknown;
  end;
  if V_PGI.IoError = oeOk then
  begin
    if not TOBAnaS.InsertDB(nil) then V_PGI.IoError := oeUnknown;
  end;
  if V_PGI.IoError = oeOk then ValideLesNomen(TOBNomenclature);
  {Traitement de la pièce annulée}
  if AnnulPiece then
  begin
    FOMarqueChqDifUtilise(TOBEches_O, True);
    FOMarqueTicketAnnule(TOBPiece_O, TOBPiece);
  end;
  // Fidelité JD
  if V_PGI.IoError = oeOk then
  begin
    if AssignePiece then FideliteCli.AnnuleFidelite(TOBPiece_O.GetValue('GP_TIERS'));
    FideliteCli.ValideFidelite();
  end;
  FiniMove;
end;

procedure TFTicketFO.ValideImpression;
var Editee: string;
  OuvrirTiroir: boolean;
begin
  // Fidélité
  TOBPiece.AddChampSupValeur('CUMULPIECE', FideliteCli.CumulPiece(TOBPiece));
  FideliteCli.CumulFidelite := 0; //Force le recalcul du cumul
  TOBPiece.AddChampSupValeur('CUMULFIDELITE', FideliteCli.CumulFidelite);
  // Pour ne pas imprimer la mention Duplicata
  if (PieceContainer.Action = taCreat) or (AnnulPiece) then
  begin
    Editee := TOBPiece.GetValue('GP_EDITEE');
    TOBPiece.PutValue('GP_EDITEE', '-');
  end;
  {$IFNDEF CHR}
  OuvrirTiroir := not (AnnulPiece);
  // Impression de la pièce
  FOImprimerLaPiece(TOBPiece, TOBTiers, TOBEches, TOBComms, PieceContainer.Cledoc, True, False, OuvrirTiroir);
  // Impression des bons d'achat et des chèques
  if PieceContainer.Action = taCreat then
  begin
    FOImpressionBons(TOBPiece, TOBArticles);
    FOImpressionAvoirs(TOBEches);
  end;
  {$ELSE}
  ImprimerLapieceChr(TOBPiece,TOBTiers,TobEches,PieceContainer.Cledoc,-1);
  {$ENDIF}
  if (PieceContainer.Action = taCreat) or (AnnulPiece) then TOBPiece.PutValue('GP_EDITEE', Editee);
end;

procedure TFTicketFO.ValideNumeroPiece;
var NewNum: integer;
begin
  if ((PieceContainer.Action = taCreat) or (PieceContainer.TransfoPiece) or (PieceContainer.DuplicPiece)) then
  begin
    NewNum := SetNumberAttribution(TOBPiece.GetValue('GP_SOUCHE'), 1);
    if NewNum > 0 then PieceContainer.Cledoc.NumeroPiece := NewNum;
  end;
end;

function TFTicketFO.ValideTablesLibres(FromBouton: boolean): boolean;
begin
  Result := True;
  {$IFDEF GESCOM}
  Exit;
  {$ELSE}
  if (AnnulPiece) or (AssignePiece) then Exit;
  Result := SaisieTablesLibres(TOBPiece, FromBouton);
  {$ENDIF}
end;

procedure TFTicketFO.ClickValideTicket;
begin
  if not BValider.Enabled then Exit;
  {$IFDEF CHR}
  if TransfertFac=True then Exit;
  {$ENDIF}
  if (PieceContainer.Action = taConsult) and (not AnnulPiece) and (not AssignePiece) then Exit;
  if not SortDeLaLigne then Exit;
  TOBPiece.GetEcran(Self, PEntete);
  if (TOBPiece.GetValue('GP_TOTALTTCDEV') = 0) or (AnnulPiece) or (AssignePiece) then ClickValide
  else AppelEcheancesFO;
end;

procedure TFTicketFO.ClickValide;
var io: TIOErr;
  ResGC, NoMsg: integer;
begin
  // Tests et actions préalables
  if not BValider.Enabled then Exit;
  if (PieceContainer.Action = taConsult) and (not AnnulPiece) and (not AssignePiece) then Exit;
  {$IFDEF CHR}
  if TransfertFac = True
  then
  begin
    ReInitPiece ;
    exit;
  end;
  {$ENDIF}
  if not SortDeLaLigne then Exit;
  TOBPiece.GetEcran(Self, PEntete);
  if (not PieceModifiee) and (not AnnulPiece) then Exit;
  FODepileTOBGrid(GS, TOBPiece, [SG_RefArt, SG_Lib]);
  // Stat
  if not ValideTablesLibres(False) then Exit;
  DepileTOBLignes(GS, TOBPiece, GS.Row, 1);
  if TesteRisqueTiers(True) then Exit;
  if TesteMargeMini(GS.Row) then Exit;
  // Contrôle intégrité
  ResGC := GCPieceCorrecte(PieceContainer);
  if ResGC > 0 then
  begin
    HErr.Execute(ResGC, Caption, '');
    Exit;
  end;
  // Contrôle intégrité spécifique au Front Office
  ResGC := FOPieceCorrecte(TOBPiece, TOBArticles);
  if ResGC > 0 then
  begin
    HErr.Execute(ResGC, Caption, '');
    Exit;
  end;
  // Appels automatiques en fin de saisie
  if OuvreAutoPort then ClickPorcs;
  if CompAnalP = 'AUT' then ClickVentil(True);
  // Confirmation pour annuler la pièce
  if AnnulPiece then
  begin
    if (TOBPiece.FieldExists('GP_TICKETANNULE')) and (TOBPiece.GetValue('GP_TICKETANNULE') = 'X') then
    begin
      PGIError('Opération impossible, ce ticket a déjà été annulé.', 'Annulation de ticket');
      BValider.Enabled := False;
      Close;
      Exit;
    end else
    begin
      if (PGIAsk('Confirmez-vous l''annulation de ce ticket ?', 'Annulation de ticket') <> mrYes) then Exit;
    end;
  end;
  // Enregistrement de la saisie
  BValider.Enabled := False;
  ValideEnCours := True;
  NoMsg := 0;
  V_PGI.TransacErrorMessage := '';
  io := Transactions(ValideNumeroPiece, 5);
  if io = oeOk then
  begin
    if AnnulPiece then
      FOAnnulReglementLie(TOBPiece, TOBArticles, TOBEches, TOBPiece_O)
    else
      if not FOAffecteNoBonReglementLie(TOBPiece, TOBArticles, TOBEches) then io := oeUnknown;
  end;
  if io = oeOk then io := Transactions(ValideLaPiece, 5);
  BValider.Enabled := True;
  case io of
    oeOk: ForcerFerme := True;
    oeUnknown: NoMsg := 5;
    oeLettrage: NoMsg := 13;
    oeSaisie: NoMsg := 6;
    oePointage: NoMsg := 23;
    oeStock: NoMsg := 27;
  end;
  if io <> oeOk then
  begin
    if NoMsg <= 0 then NoMsg := 5;
    MessageAlerte(HTitres.Mess[NoMsg]);
    ValideEnCours := False;
    LogAGL(DateTimeToStr(Now) + ' ERREUR n°' + IntToStr(Ord(io)) + ' sur Ticket n°' + IntToStr(TOBPiece.GetValue('GP_NUMERO'))
      + ' Caisse n°' + TOBPiece.GetValue('GP_CAISSE') + ' => ' + HTitres.Mess[NoMsg] + ' ' + V_PGI.TransacErrorMessage);
    Exit;
  end;
  if DemandeDetaxe then
  begin
    if PieceContainer.Action = taCreat then
    begin
      io := Transactions(ValideDetaxe, 1);
      if io <> oeOk then MessageAlerte(HTitres.Mess[35]);
    end;
  end;
  if GetInfoParPiece(PieceContainer.NewNature, 'GPP_IMPIMMEDIATE') = 'X' then
  begin
    if ((PieceContainer.Action = taCreat) or (PieceContainer.DuplicPiece) or (PieceContainer.TransfoPiece)) then
    begin
      io := Transactions(ValideImpression, 1);
      if io <> oeOk then MessageAlerte(HTitres.Mess[19]);
    end;
  end;
  //VH_GC.TOBPCaisse.PutValue('LASTNUMTIC', TOBPiece.GetValue('GP_NUMERO'));
  FOMajChampSupValeur(VH_GC.TOBPCaisse, 'LASTNUMTIC', TOBPiece.GetValue('GP_NUMERO'));
  FOMajChampSupValeur(VH_GC.TOBPCaisse, 'LASTVENDEUR', TOBPiece.GetValue('GP_REPRESENTANT'));
  ValideEnCours := False;
  if PieceContainer.Action <> taCreat then
  begin
    {$IFNDEF FOS5}
    if ((DuplicPiece) or (PieceContainer.TransfoPiece)) then MontreNumero(TOBPiece);
    {$ENDIF}
    Close;
  end else
  begin
    {$IFNDEF FOS5}
    MontreNumero(TOBPiece);
    {$ENDIF}
    ReInitPiece;
    ReInitUtilisateur;
  end;
end;

procedure TFTicketFO.BValiderClick(Sender: TObject);
begin
  ClickValideTicket;
end;

procedure TFTicketFO.BAbandonClick(Sender: TObject);
begin
  if IsInside(Self) then
  begin
    Close;
    THPanel(Parent).CloseInside;
  end else ModalResult := mrCancel;
end;

{==============================================================================================}
{======================================= AFFICHEUR ============================================}
{==============================================================================================}
///////////////////////////////////////////////////////////////////////////////////////
//  TimerLCDTimer
///////////////////////////////////////////////////////////////////////////////////////

procedure TFTicketFO.TimerLCDTimer(Sender: TObject);
var Stg: string;
begin
  if not TimerLCD.Enabled then exit;
  Stg := FODecaleGauche(fAffTXT.Caption, 1, fAffTXT.NoOfChars);
  AffichageLCD(Stg, 0, 0, 0, '', [ofInterne], qaLibreTexte, False);
  if TimerLCD.Interval <> MinInterval then TimerLCD.Interval := MinInterval;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  PEnteteResize
///////////////////////////////////////////////////////////////////////////////////////

procedure TFTicketFO.PEnteteResize(Sender: TObject);
begin
  ResizerAfficheur;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  InvertirAfficheur : place l'afficheur en mode inversé
///////////////////////////////////////////////////////////////////////////////////////

procedure TFTicketFO.InvertirAfficheur;
begin
  if not fAffTXT.Visible then exit;
  fAffTXT.Caption := '';
  FAffTXT.UpSideDown := (VH_GC.TOBPCaisse.GetValue('GPK_AFFINVERSE') = 'X');
  fAffTXT.Invalidate;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  Resizerafficheur : ajuste la taille de l'afficheur
///////////////////////////////////////////////////////////////////////////////////////

procedure TFTicketFO.Resizerafficheur;
begin
  if not fAffTXT.Visible then exit;
  fAffTXT.Width := PEntete.Width - 2 * fAffTXT.left;
  fAffTXT.Left := (PEntete.Width - fAffTXT.Width) div 2;
  fAffTXT.top := 4;
  fAffTXT.CalcCharSize;
  fAffTXT.Invalidate;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  TraiteTexteAff : convertit un texte pour l'afficheur client
///////////////////////////////////////////////////////////////////////////////////////

function TFTicketFO.TraiteTexteAff(ST: string): string;
var i: integer;
begin
  result := AnsiUppercase(traduirememoire(St));
  for i := 1 to length(Result) do
    case Result[i] of
      'á', 'à', 'â', 'ä', 'Á', 'À', 'Â', 'Ä': Result[i] := 'A';
      'é', 'è', 'ê', 'ë', 'É', 'È', 'Ê', 'Ë': Result[i] := 'E';
      'í', 'ì', 'î', 'ï', 'Í', 'Ì', 'Î', 'Ï': Result[i] := 'I';
      'ó', 'ò', 'ô', 'ö', 'Ó', 'Ò', 'Ô', 'Ö': Result[i] := 'O';
      'ú', 'ù', 'û', 'ü', 'Ú', 'Ù', 'Û', 'Ü': Result[i] := 'U';
      '£': Result[i] := 'L';
      '¥': Result[i] := 'Y';
      /////    '€'                              : Result[i]:='E' ;
    end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  AffichageLCD : affiche un texte, une quantité ou un montant sur l'afficheur client
///////////////////////////////////////////////////////////////////////////////////////

procedure TFTicketFO.AffichageLCD(Texte: string; Qte, Mnt: Double; NbDec: Integer; Sigle: string; Ou: TSetOuAffiche; Quoi: TQuoiAffiche; Attente: Boolean);
var Libelle, Valeur: string;
  Taille, Pos, Nb: Integer;
  Err: TMC_Err;
begin
  if not (fAffTXT.Visible) and not (Assigned(AFFExterne)) then Exit;
  if fAffTXT.Visible then
  begin
    if fAffTXT.Tag <> 0 then Delay(2500); //If tag<>0, le dernier message doit être affiché plus temps
    fAffTXT.Tag := ord(Attente);
  end;
  QuoiSurLCD := Quoi;
  Valeur := '';
  Libelle := '';
  if not (ofTimer in ou) then Libelle := TraiteTexteAff(Texte);
  if Qte <> 0 then
  begin
    Taille := fAffTXT.NoOfChars;
    if Length(Libelle) > Taille then
    begin
      // on tronque le texte si on ne dispose pas d'assez de place pour le montant
      Pos := Taille;
      Nb := Length(Libelle) - Taille;
      Delete(Libelle, Pos, Nb);
    end else
    begin
      // on complete par des espaces pour se placer sur la 2ème ligne
      while Length(Libelle) < Taille do Libelle := Libelle + ' ';
    end;
    Libelle := Libelle + StrfMontant(Qte, fAffTXT.noOfChars, V_PGI.OkDecQ, '', True) + ' X '
  end;
  if Mnt <> 0 then Valeur := StrfMontant(Mnt, fAffTXT.noOfChars, NbDec, Sigle, True);
  if Valeur <> '' then
  begin
    Taille := fAffTXT.NoOfChars * fAffTXT.TextLines;
    // on tronque le texte si on ne dispose pas d'assez de place pour le montant
    if Length(Valeur) > (Taille - Length(Libelle) + 1) then
    begin
      Pos := Taille - Length(Valeur);
      Nb := Length(Libelle) - Length(Valeur) + 1;
      Delete(Libelle, Pos, Nb);
    end;
    // cadrage du montant à gauche
    while Length(Valeur) < (Taille - length(Libelle)) do Valeur := ' ' + Valeur;
  end;
  // Afficheur interne
  if fAffTXT.Visible then
  begin
    if fAffTXT.Caption <> Libelle + Valeur then fAffTXT.Caption := Libelle + Valeur;
  end;
  // Afficheur externe
  if (Assigned(AFFExterne)) and (ofClient in Ou) then
  begin
    if (AffExterne.IsConnected) and (not AffExterne.Affiche(Texte, Valeur, Err)) then FOMCErr(Err, Caption);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  LCDInVisible : rend invisible l'afficheur client
///////////////////////////////////////////////////////////////////////////////////////

procedure TFTicketFO.LCDInVisible;
var Ind: Integer;
begin
  Ind := fAffTXT.Top + fAffTXT.Height;
  PEntete.Height := PEntete.Height - Ind;
  GP_NUMEROPIECE.Top := GP_NUMEROPIECE.Top - Ind;
  HGP_DATEPIECE.Top := HGP_DATEPIECE.Top - Ind;
  GP_DATEPIECE.Top := GP_DATEPIECE.Top - Ind;
  HGP_TIERS.Top := HGP_TIERS.Top - Ind;
  GP_TIERS.Top := GP_TIERS.Top - Ind;
  HGP_REPRESENTANT.Top := HGP_REPRESENTANT.Top - Ind;
  GP_REPRESENTANT.Top := GP_REPRESENTANT.Top - Ind;
  fAffTXT.Visible := False;
  TimerLCD.Enabled := False;
end;

{==============================================================================================}
{=================================== MAXIMISATION DE LA FORME =================================}
{==============================================================================================}
///////////////////////////////////////////////////////////////////////////////////////
//  TimerGenTimer
///////////////////////////////////////////////////////////////////////////////////////

procedure TFTicketFO.TimerGenTimer(Sender: TObject);
begin
  if not TimerGen.Enabled then exit;
  if TimerGen.interval = ToMaximise then
  begin
    InMaximise := True;
    TimerGen.Interval := MaxInterval;
    BorderStyle := bsNone;
    Caption := '';
    BorderIcons := [];
    Visible := False;
    WindowState := wsMaximized;
    Visible := True;
    InMaximise := False;
    TimerGen.Enabled := False;
    // positionnement des barres de boutons
    Valide97.DragHandleStyle := dhNone;
    Valide97.BorderStyle := bsNone;
    Valide97.DockPos := DockBottom.Width - Valide97.Width;
    Outils97.DragHandleStyle := dhNone;
    Outils97.BorderStyle := bsNone;
    Outils97.DockPos := 0;
    if AssignePiece then
    begin
      if GP_TIERS.CanFocus then GP_TIERS.SetFocus;
      ChoixTiers;
    end else
    begin
      // pré-positionnement dans la grille
      if (VH_GC.TOBPCaisse.GetValue('GPK_PREPOSGRID') = 'X') or (not GoToEntete) then
        if GS.CanFocus then GS.SetFocus;
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FormResize
///////////////////////////////////////////////////////////////////////////////////////

procedure TFTicketFO.FormResize(Sender: TObject);
var Position: Integer;
begin
  PMain.Tag := PMain.Height - PnlCorps.Height;
  // Redimensionnent du LCD client
  ResizerAfficheur;
  // Redimensionnent du LCD total général
  PI_NETAPAYERDEV.left := pPied.width - PI_NETAPAYERDEV.Width - 4;
  TPI_NETAPAYERDEV.Left := PI_NETAPAYERDEV.left - TPI_NETAPAYERDEV.width - 8;
  // Affichage de la photo de l'article
  if (FOAfficheImageArt) or (FOLogoCaisse <> '') then
  begin
    Position := (ImageArticle.Left * 2) + ImageArticle.Width;
    LIBELLETIERS.Left := Position;
    LblPrixU.Left := Position;
    LblA_LIBELLE.Left := Position;
    LIBELLEARTICLE.Left := Position;
    TSOLDETIERS.Left := Position;
  end;
  // Redimensionnent des libellés du tiers et de l'article
  LIBELLETIERS.Width := PI_NETAPAYERDEV.Left - LIBELLETIERS.Left - 6;
  LIBELLEARTICLE.Width := PI_NETAPAYERDEV.Left - LIBELLEARTICLE.Left - 6;
  // Repositionnement de la quantité totale et du total en contre-valeur
  TLGP_TOTALQTEFACT.Left := PI_NETAPAYERDEV.Left;
  LGP_TOTALQTEFACT.Left := TLGP_TOTALQTEFACT.Left + TLGP_TOTALQTEFACT.Width + 6;
  // Repositionnement des mentions détaxe et à livrer
  LDETAXE.Left := PI_NETAPAYERDEV.Left + PI_NETAPAYERDEV.Width - LDETAXE.Width - 16;
  LALIVRER.Left := TPI_NETAPAYERDEV.Left;
  // Redimensionnent des boutons tactiles
  if assigned(PnlBoutons) then PnlBoutons.ResizeClavier(nil);
  // Redimensionnent de l'image de base d'écran
  {**
  if not FOExisteClavierEcran then
     BEGIN
     // Visualisation des titres et de 12 lignes dans la grille
     Taille:=PMain.Height-PPied.Height-((GS.DefaultRowHeight+1)*13)-3 ;
     if Taille < 1 then Taille:=1 ;
     PImage.Height:=Taille ;
    END ;
   **}
   HMTrad.ResizeGridColumns(GS);
end;

{==============================================================================================}
{======================================= GESTION TACTILE ======================================}
{==============================================================================================}

///////////////////////////////////////////////////////////////////////////////////////
//  OnChangeCalculette2: répercute le changement de valeur de la calculette du 2ème pavé sur le 1er
///////////////////////////////////////////////////////////////////////////////////////
procedure TFTicketFO.OnChangeCalculette(AValue: Variant);
var
  Stg: string;
begin
  if (PnlBoutons <> nil) and (PnlBoutons.Calculette.LCD.Valeur <> AValue) then
  begin
    if VarIsNull(AValue) then
    begin
      Stg := '';
    end else
    begin
      Stg := VarAsType(AValue, varString);
      if Stg = #0 then Stg := '';
    end;
    if Stg = '' then
      PnlBoutons.Calculette.BtnsCalculetteUP
    else
      PnlBoutons.Calculette.LCD.Valeur := AValue;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnChangeCalculette2 : répercute le changement de valeur de la calculette du 1er pavé sur le 2ème
///////////////////////////////////////////////////////////////////////////////////////
procedure TFTicketFO.OnChangeCalculette2(AValue: Variant);
var
  Stg: string;
begin
  if (PnlBtn2 <> nil) and (PnlBtn2.Calculette.LCD.Valeur <> AValue) then
  begin
    if VarIsNull(AValue) then
    begin
      Stg := '';
    end else
    begin
      Stg := VarAsType(AValue, varString);
      if Stg = #0 then Stg := '';
    end;
    if Stg = '' then
      PnlBtn2.Calculette.BtnsCalculetteUP
    else
      PnlBtn2.Calculette.LCD.Valeur := AValue;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  AllerAuControle :
///////////////////////////////////////////////////////////////////////////////////////

function TFTicketFO.AllerAuControle(Vers: TOuSaisie; NoLig, NoCol: Integer; RechLigne, VerifAcces: Boolean): Boolean;
var ACol, ARow, OldCol, OldRow: Integer;
  Cancel: Boolean;
  Ou: TOuSaisie;
begin
  Result := False;
  Ou := IndiqueOuSaisie;
  OldCol := GS.Col;
  OldRow := GS.Row;
  // Appelle interdit depuis la saisie des échéances
  if (Ou = saEcheance) and (Vers <> saEcheance) then Exit;
  // Recherche de la ligne sur laquelle on veut se positionner
  if NoLig < 0 then NoLig := GS.Row;
  if (Vers = saLigne) and (RechLigne) then
  begin
    for NoLig := GS.Row downto 0 do
    begin
      if GetCodeArtUnique(TOBPiece, NoLig) <> '' then Break;
      if (NoCol in [SG_Rem, SG_TypeRem]) and
        (GetChampLigne(TOBPiece, 'GL_TYPELIGNE', NoLig) = 'TOT') then Break;
    end;
    if NoLig < 0 then Exit;
  end;
  // Vérification si le champ à atteindre est saisissable
  if (Vers = saLigne) and (VerifAcces) and not (ZoneAccessible(NoCol, NoLig)) then Exit;
  // On quitte un contrôle de l'entête ou du pied pour aller dans le Grid
  if (Ou in [saEntete, saPied]) and (Vers = saLigne) then
  begin
    // On quitte le contrôle
    if (Screen.ActiveControl is THValComboBox) then
    begin
      if Assigned(THValComboBox(Screen.ActiveControl).OnExit) then THValComboBox(Screen.ActiveControl).OnExit(Self);
    end else
      if (Screen.ActiveControl is THCritMaskEdit) then
    begin
      if Assigned(THCritMaskEdit(Screen.ActiveControl).OnExit) then THCritMaskEdit(Screen.ActiveControl).OnExit(self);
    end else
      if (Screen.ActiveControl is TEdit) then
    begin
      if Assigned(TEdit(Screen.ActiveControl).OnExit) then TEdit(Screen.ActiveControl).OnExit(Self);
    end;
    // On entre dans le Grid (le SetFocus génére un GSEnter)
    if GS.CanFocus then GS.SetFocus;
  end;
  // Si On quitte une cellule du Grid on desactive le combo
  if (Ou = saLigne) and ((Vers <> saLigne) or (NoLig <> OldRow) or (NoCol <> OldCol)) then
  begin
    ACol := GS.Col;
    ARow := GS.Row;
    if GS.ValCombo <> nil then GS.ValCombo.Enabled := False;
    ///GS.ShowCombo(ACol, ARow) ;
  end;
  // positionnement sur la cellule à atteindre
  if Vers = saLigne then
  begin
    GS.Col := NoCol;
    GS.Row := NoLig;
  end;
  // On quitte une cellule du Grid
  if Ou = saLigne then
  begin
    Cancel := False;
    if (Vers <> saLigne) or (NoLig <> OldRow) or (NoCol <> OldCol) then
    begin
      ACol := OldCol;
      ARow := OldRow;
      {**
      GS.ValCombo.Enabled := False ;
      GS.ShowCombo(ACol, ARow) ;
      **}
      GSCellExit(Self, ACol, ARow, Cancel);
      if Cancel then
      begin
        GS.Col := OldCol;
        GS.Row := OldRow;
        Exit;
      end;
    end;
    if (Vers <> saLigne) or (NoLig <> OldRow) then
    begin
      ARow := OldRow;
      GSRowExit(Self, ARow, Cancel, False);
      if Cancel then
      begin
        GS.Col := OldCol;
        GS.Row := OldRow;
        Exit;
      end;
    end;
  end;
  // On entre dans la cellule du Grid
  if Vers = saLigne then
  begin
    Cancel := False;
    // Le champ est-il saisissable ?
    {**
    if (VerifAcces) and not (ZoneAccessible(NoCol, NoLig)) then
       BEGIN
       GS.Col := OldCol ; GS.Row := OldRow ;
       Exit ;
       END ;
    **}
    if NoLig <> OldRow then
    begin
      ARow := NoLig;
      GSRowEnter(Self, ARow, Cancel, False);
      if Cancel then
      begin
        GS.Col := OldCol;
        GS.Row := OldRow;
        Exit;
      end;
    end;
    if (NoLig <> OldRow) or (NoCol <> OldCol) then
    begin
      ACol := OldCol;
      ARow := OldRow;
      GSCellEnter(Self, ACol, ARow, Cancel);
      if Cancel then
      begin
        GS.Col := OldCol;
        GS.Row := OldRow;
        Exit;
      end;
      if GS.ValCombo <> nil then
      begin
        GS.ValCombo.Enabled := TRUE;
        GS.ShowCombo(ACol, ARow);
      end;
    end;
  end;
  Result := True;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ToucheArticle : gestion d'une touche "article"
///////////////////////////////////////////////////////////////////////////////////////

procedure TFTicketFO.ToucheArticle(Concept, Code, Extra: string; Qte, Prix: Double);
var ACol, ARow, ColPrix: Integer;
    Cancel: Boolean;
    TobTmp, TobTmp1 : TOB; // JTR Lien Opération Caisse avec Tiers
begin
  if PieceContainer.Action = taConsult then Exit;
  if (SG_RefArt = -1) or (Code = '') then Exit;
  // JTR Lien Opération Caisse avec Tiers
  if (GP_TIERS.Text = '') and (Concept = 'OCT') then
  begin
    TobTmp := FOGetPTobArtFiTiers;
    if TobTmp <> nil then
    begin
      TobTmp1 := TobTmp.FindFirst(['GA_CODEARTICLE'],[Code],True);
      GP_TIERS.Text := TobTmp1.GetValue('GA_URLWEB');
      ChargeTiers;
    end;
  end;
  // Fin - JTR Lien Opération Caisse avec Tiers

  // Positionnement dans le Grid
  if not AllerAuControle(saLigne, -1, SG_RefArt, False, True) then Exit;
  ACol := GS.Col;
  ARow := GS.Row;
  // Le client doit être renseiné
  if TOBTiers.GetValue('T_TIERS') = '' then
  begin
    if GP_TIERS.CanFocus then GP_TIERS.SetFocus else GotoEntete;
    Exit;
  end;
  // Insertion du code article : choix entre renseigner le code article ou insérer une ligne
  if (GetCodeArtUnique(TOBPiece, ARow) <> '') or (VarToStr(GetChampLigne(TOBPiece, 'GL_LIBELLE', ARow)) <> '') then
    ClickInsert(ARow);
  GS.Cells[SG_RefArt, ARow] := Code;
  TraiteArticle(ACol, ARow, Cancel, False, True, 1, (Concept = 'ART'));
  if Cancel then Exit;
{$IFDEF GESCOM} // JTR - Ne pas cumuler des articles avec articles financiers dans un même ticket
  if not CumulArtOk(Acol,Arow,Cancel) then exit;
  RemiseDuTarif(Arow); // JTR - Démarque obligatoire
{$ENDIF GESCOM}
  StCellCur := GS.Cells[ACol, ARow];
  SurLigSuivante := False;
  // Insertion de la quantité
  if (SG_QF <> -1) and (Qte <> 0) and (Qte <> Valeur(GS.Cells[SG_QF, ARow])) then
  begin
    if not AllerAuControle(saLigne, -1, SG_QF, False, True) then Exit;
    GS.Cells[SG_QF, ARow] := FloatToStr(Qte);
  end;
  // Insertion du prix
  if SG_PxNet <> -1 then ColPrix := SG_PxNet else ColPrix := SG_Px;
  if (ColPrix <> -1) and (Prix <> 0) and (Prix <> Valeur(GS.Cells[ColPrix, ARow])) then
  begin
    //if not AllerAuControle(saLigne, -1, ColPrix, False, True) then Exit;
    //GS.Cells[ColPrix, ARow] := FloatToStr(Prix);
    if AllerAuControle(saLigne, -1, ColPrix, False, True) then
      GS.Cells[ColPrix, ARow] := FloatToStr(Prix);
  end;
  // Repositionnement sur la cellule article
  if ((SG_QF <> -1) and (Qte <> 0)) or ((ColPrix <> -1) and (Prix <> 0)) then
  begin
    if not AllerAuControle(saLigne, -1, SG_RefArt, False, True) then Exit;
  end;
  Beep;
  // Aller sur la ligne suivante
  SurLigSuivante := True;
  FOSimuleClavier(VK_TAB, False, TMOSimuleClav);
end;
{$IFDEF CHR}
///////////////////////////////////////////////////////////////////////////////////////
//  ToucheRes : gestion d'une touche "Ressource"
///////////////////////////////////////////////////////////////////////////////////////
Procedure TFTicketFO.ToucheRes ( Concept, Code, Extra : String  ) ;
Var ACol, ARow : Integer ;
    Cancel     : Boolean ;
    Ressource:string;
    QDos:TQuery;
BEGIN
  if PieceContainer.Action = taConsult then Exit ;
  if TOBPiece.GetValue('GP_TOTALTTCDEV') = 0 then Exit ;
  if (SG_RefArt = -1 ) or (Code = '') then Exit ;
  QDos:=OpenSQL('Select HDR_RESSOURCE from HRDOSRES where HDR_DOSRES="'+Code+'"',True) ;
  if (not QDos.eof) then
  begin
    Ressource:=QDos.FindField('HDR_RESSOURCE') .AsString;
  end;
  ferme(QDos);
  if HShowMessage('0;Confirmez-vous le transfert ? ; Les lignes vont être transférées sur la ressource '+Ressource+';Q;YN;N;N;','','')=mrYes then
  begin
    TransfertRessource (Code,TOBPIECE,TOBECHES);
  end;
END ;
{$ENDIF}
///////////////////////////////////////////////////////////////////////////////////////
//  ToucheCommentaire : gestion d'une touche "commentaire"
///////////////////////////////////////////////////////////////////////////////////////

procedure TFTicketFO.ToucheCommentaire(Concept, Code, Extra: string; Qte, Prix: Double);
var NoLig: Integer;
  QQ: TQuery;
  SQL, Libelle, Chaine: string;
  Contenu: TStrings;
  LigCom: string;
  i_ind, ANewRow, ARow: integer;
  TOBL: TOB;
begin
  if PieceContainer.Action = taConsult then Exit;
  if Screen.ActiveControl <> GS then Exit;
  if not CommentLigne then Exit;
  if (SG_Lib = -1) or (Code = '') then Exit;
  // Recherche du Libellé du commentaire
  if GetParamSoc('SO_GCCOMMENTAIRE') then
  begin
    ARow := GS.Row;
    Contenu := GetCommentaireLigne(Code);
    if Contenu <> nil then
    begin
      if GS.RowCount < TOBPiece.Detail.Count + Contenu.Count then GS.RowCount := GS.RowCount + Contenu.Count;
      for i_ind := 0 to Contenu.Count - 1 do
      begin
        LigCom := Contenu[i_ind];
        ANewRow := ARow + i_ind;
        if ANewRow > TOBPiece.Detail.Count then CreerTOBLignes(GS, TOBPiece, TOBTiers, TOBAffaire, ANewRow)
        else if i_ind > 0 then ClickInsert(ANewRow);
        TOBL := GetTOBLigne(TOBPiece, ANewRow);
        if TOBL = nil then Continue;
        GS.Cells[SG_Lib, ANewRow] := copy(LigCom, 1, 70);
        TOBL.PutValue('GL_LIBELLE', GS.Cells[SG_Lib, ANewRow]);
      end;
      Contenu.Free;
    end;
  end else
  begin
    SQL := KBMakeSQL(Concept, Code, Chaine, Libelle);
    if SQL = '' then Exit;
    Chaine := '';
    QQ := OpenSQL(SQL, True);
    if not QQ.Eof then Chaine := QQ.FindField(Libelle).AsString;
    Ferme(QQ);
    if Chaine = '' then Exit;
    // Positionnement dans le Grid
    if not AllerAuControle(saLigne, -1, SG_Lib, False, True) then Exit;
    NoLig := GS.Row;
    // Insertion du commentaire : choix entre renseigner le code article ou insérer une ligne
    if (GetCodeArtUnique(TOBPiece, NoLig) <> '') or (GetChampLigne(TOBPiece, 'GL_LIBELLE', NoLig) <> '') then ClickInsert(NoLig);
    GS.Cells[SG_Lib, NoLig] := Chaine;
  end;
  // Aller sur la ligne suivante
  SurLigSuivante := True;
  FOSimuleClavier(VK_DOWN, False, TMOSimuleClav);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ToucheRemise : gestion d'une touche "remise"
///////////////////////////////////////////////////////////////////////////////////////

procedure TFTicketFO.ToucheRemise(Concept, Code, Extra: string; Qte, Prix: Double);
var CurLig, CurCol, Ind, ACol, ARow: Integer;
  Ou: TOuSaisie;
  TOBA: TOB;
  Cancel: Boolean;
begin
  Cancel := False;
  if PieceContainer.Action = taConsult then Exit;
  Ou := IndiqueOuSaisie;
  if Ou = saRemiseTicket then
  begin
    // Insertion du pourcentage de remise
    if Prix <> 0 then
    begin
      FormEcheance.GP_REMISEPOURPIED.Value := Prix;
      FormEcheance.GP_REMISEPOURPIEDExit(nil);
    end;
    // Insertion du code type de remise
    if (Code <> '') and (FOTabletteValueExist('GCTYPEREMISE', Code, 'GTR_FERME<>"X"')) then
    begin
      FormEcheance.GP_TYPEREMISE.Value := Code;
    end;
    FormEcheance.RemTicketCtrl.SetFocus;
  end else
    if Ou = saLigne then
  begin
    //////if Screen.ActiveControl <> GS then Exit ;
    CurLig := GS.Row;
    CurCol := GS.Col;
    // Insertion du pourcentage de remise
    if (SG_Rem <> -1) and (Prix <> 0) then
    begin
      //      if (FOGetParamCaisse('GPK_REMAFFICH') = 'X') and not (AllerAuControle(saLigne, -1, SG_Rem, True, True)) then
      if not AllerAuControle(saLigne, -1, SG_Rem, True, True) then
      begin
        for Ind := GS.Row downto 0 do if GetCodeArtUnique(TOBPiece, Ind) <> '' then Break;
        if Ind >= 0 then TOBA := FindTOBArtRow(TOBPiece, TOBArticles, Ind)
        else TOBA := nil;
        if (TOBA <> nil) and (TOBA.GetValue('GA_REMISELIGNE') = '-') then HPiece.Execute(41, Caption, '');
        Exit;
      end;
      GS.Cells[SG_Rem, GS.Row] := FloatToStr(Prix);
      ACol := SG_Rem;
      ARow := GS.Row;
      GSCellExit(Self, ACol, ARow, Cancel);
      if not Cancel then StCellCur := GS.Cells[ACol, ARow];
    end;
    // Insertion du code type de remise
    if (not Cancel) and (SG_TypeRem <> -1) and (Code <> '') and (FOTabletteValueExist('GCTYPEREMISE', Code, 'GTR_FERME<>"X"')) then
    begin
      if not AllerAuControle(saLigne, -1, SG_TypeRem, True, True) then Exit;
      GS.CellValues[SG_TypeRem, GS.Row] := Code;
    end;
    if (not Cancel) and ((CurLig <> GS.Row) or (CurCol <> GS.Col)) then
    begin
      AllerAuControle(saLigne, CurLig, CurCol, False, False);
    end;
  end;
  if Ou = saEcheance then FormEcheance.SaisieRemiseTicket(Prix, Code);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ToucheVendeur : gestion d'une touche "vendeur"
///////////////////////////////////////////////////////////////////////////////////////

procedure TFTicketFO.ToucheVendeur(Concept, Code, Extra: string; Qte, Prix: Double);
var CurLig, CurCol: Integer;
  NiveauLigne: Boolean;
begin
  if PieceContainer.Action = taConsult then Exit;
  if Code = '' then Exit;
  NiveauLigne := False;
  if (IndiqueOuSaisie = saLigne) and (SG_Rep <> -1) and (VH_GC.TOBPCaisse.GetValue('GPK_VENDSAISLIG') = 'X') then
  begin
    // si on utilise une touche vendeur sur la 1ère ligne et si cette dernière est vide, le choix du vendeur s'applique à l'en-tête
    if (GS.Row > GS.FixedRows) or (GetCodeArtUnique(TOBPiece, GS.Row) <> '') then NiveauLigne := True;
  end;
  if NiveauLigne then
  begin
    // Insertion du code vendeur sur la ligne
    CurLig := GS.Row;
    CurCol := GS.Col;
    if not AllerAuControle(saLigne, -1, SG_Rep, True, True) then Exit;
    GS.Cells[SG_Rep, GS.Row] := Code;
    if (CurLig <> GS.Row) or (CurCol <> GS.Col) then AllerAuControle(saLigne, CurLig, CurCol, False, False);
    //////END else if (GP_REPRESENTANT.Enabled) and (VH_GC.TOBPCaisse.GetValue('GPK_VENDSAISIE') = 'X') then
  end else if VH_GC.TOBPCaisse.GetValue('GPK_VENDSAISIE') = 'X' then
  begin
    // Insertion du code vendeur du ticket
    OldRepresentant := '';
    if FORepresentantExiste(GP_REPRESENTANT.Text, TypeCom, TOBComms) then
      OldRepresentant := GP_REPRESENTANT.Text;
    GP_REPRESENTANT.Text := Code;
    if Screen.ActiveControl = GP_REPRESENTANT then FOSimuleClavier(VK_TAB) else GP_REPRESENTANTExit(Self);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ToucheModePaie : gestion d'une touche "mode de paiement"
///////////////////////////////////////////////////////////////////////////////////////

procedure TFTicketFO.ToucheModePaie(Concept, Code, Extra: string; Qte, Prix: Double);
var ACol, ARow, OldCol, OldRow: Integer;
  Cancel: Boolean;
begin
  if PieceContainer.Action = taConsult then Exit;
  if TOBPiece.GetValue('GP_TOTALTTCDEV') = 0 then Exit;
  // Pour quitter le Grid de saisie des lignes du ticket
  AllerAuControle(saEcheance, 1, 1, False, False);
  // Incrustation de l'écran de saisie des échéances si besoin
  if FormEcheance = nil then AppelEcheancesFO;
  // Ajout d'un mode de paiement
  if FormEcheance = nil then Exit;
  // Ajout d'un mode de paiement
  if FormEcheance.GSReg.CanFocus then
  begin
    FormEcheance.GSReg.SetFocus;
    FormEcheance.GSReg.Col := FormEcheance.SG_Mode;
    FormEcheance.GSReg.Cells[FormEcheance.SG_Mode, FormEcheance.GSReg.Row] := Code;
    // Insertion du montant de l'échéance
    if (FormEcheance.SG_Mont <> -1) and (Prix <> 0) then
    begin
      ACol := FormEcheance.GSReg.Col;
      ARow := FormEcheance.GSReg.Row;
      OldCol := ACol;
      OldRow := ARow;
      Cancel := False;
      FormEcheance.GSReg.Col := FormEcheance.SG_Mont;
      FormEcheance.GSReg.OnCellExit(Self, ACol, ARow, Cancel);
      if Cancel then
      begin
        GS.Col := OldCol;
        GS.Row := OldRow;
        Exit;
      end;
      FormEcheance.GSReg.OnCellEnter(Self, ACol, ARow, Cancel);
      FormEcheance.GSReg.Cells[FormEcheance.SG_Mont, FormEcheance.GSReg.Row] := FloatToStr(Prix);
    end;
  end;
  FormEcheance.EnvoieToucheGrid;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ToucheModele : gestion d'une touche "modele d'impression"
///////////////////////////////////////////////////////////////////////////////////////

procedure TFTicketFO.ToucheModele(Concept, Code, Extra: string; Qte, Prix: Double);
var Format, Stg: string;
  EnPlus: Boolean;
  NbCopies: Integer;
begin
  if not FOJaiLeDroit(22) then Exit;
  EnPlus := False;
  Format := ReadTokenst(Extra);
  if FOChoixModeleImpression(Format, Code, EnPlus, NbCopies) then
  begin
    if not TOBPiece.FieldExists('IMPFORMAT') then TOBPiece.AddChampSup('IMPFORMAT', False);
    TOBPiece.PutValue('IMPFORMAT', Format);
    if not TOBPiece.FieldExists('IMPMODELE') then TOBPiece.AddChampSup('IMPMODELE', False);
    TOBPiece.PutValue('IMPMODELE', Code);
    if not TOBPiece.FieldExists('NBEXEMPLAIRES') then TOBPiece.AddChampSup('NBEXEMPLAIRES', False);
    TOBPiece.PutValue('NBEXEMPLAIRES', NbCopies);
    if not TOBPiece.FieldExists('PLUSMODIMP') then TOBPiece.AddChampSup('PLUSMODIMP', False);
    if EnPlus then Stg := 'X' else Stg := '-';
    TOBPiece.PutValue('PLUSMODIMP', Stg);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  SimuleChoixLigneMenu : simule le choix d'une ligne de menu
///////////////////////////////////////////////////////////////////////////////////////

procedure TFTicketFO.SimuleChoixLigneMenu(LigneMenu: TMenuItem; BtnEcheance, RechLigne: Boolean);
var NoLig, CurLig: Integer;
begin
  if (BtnEcheance) and (IndiqueOuSaisie <> saEcheance) then Exit;
  if not (BtnEcheance) and (IndiqueOuSaisie = saEcheance) then Exit;
  if Assigned(LigneMenu.OnClick) then
  begin
    // Choix de la ligne sur laquelle appliquer le zoom
    CurLig := GS.Row;
    if RechLigne then
    begin
      if IndiqueOuSaisie <> saLigne then Exit;
      for NoLig := GS.Row downto 0 do if GetCodeArtUnique(TOBPiece, NoLig) <> '' then Break;
      if NoLig < 0 then Exit;
      if GS.Row <> NoLig then
      begin
        GereEnabled(NoLig);
        GS.Row := NoLig;
      end;
    end;
    if LigneMenu.Enabled then LigneMenu.OnClick(Self);
    if GS.Row <> CurLig then
    begin
      GereEnabled(CurLig);
      GS.Row := CurLig;
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  SimuleBoutonZoom : simule la frappe d'un bouton "Zoom"
///////////////////////////////////////////////////////////////////////////////////////

procedure TFTicketFO.SimuleBoutonZoom(BZoom: TBitBtn; BtnEcheance, RechLigne: Boolean);
var NoLig, CurLig: Integer;
begin
  /////if (BtnEcheance) and (IndiqueOuSaisie <> saEcheance) then Exit ;
  if not (BtnEcheance) and (IndiqueOuSaisie = saEcheance) then Exit;
  if Assigned(BZoom.OnClick) then
  begin
    // Choix de la ligne sur laquelle appliquer le zoom
    CurLig := GS.Row;
    if RechLigne then
    begin
      if IndiqueOuSaisie <> saLigne then Exit;
      for NoLig := GS.Row downto 0 do if GetCodeArtUnique(TOBPiece, NoLig) <> '' then Break;
      if NoLig < 0 then Exit;
      if GS.Row <> NoLig then
      begin
        GereEnabled(NoLig);
        GS.Row := NoLig;
      end;
    end;
    if BZoom.Enabled then BZoom.OnClick(Self);
    if GS.Row <> CurLig then
    begin
      GereEnabled(CurLig);
      GS.Row := CurLig;
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  AllerSur : donne le focus à une colonne du Grid
///////////////////////////////////////////////////////////////////////////////////////

procedure TFTicketFO.AllerSur(NoCol: Integer);
begin
  if PieceContainer.Action = taConsult then Exit;
  if IndiqueOuSaisie <> saLigne then Exit;
  if (NoCol < GS.FixedCols) or (NoCol >= GS.ColCount) then Exit;
  AllerAuControle(saLigne, -1, NoCol, True, True);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  DispatchFonctionClavier : gestion d'une touche de fonction
///////////////////////////////////////////////////////////////////////////////////////

procedure TFTicketFO.DispatchFonctionClavier(Concept, Code, Extra: string; Qte, Prix: Double);
var NoLig, CurLig: Integer;
  Cancel: Boolean;
  ACol, ARow: Integer;
begin
  if Code = '' then Exit;
  if IndiqueOuSaisie = saEcheance then
    FormEcheance.TOBEcheLigneExiste
  else
    TOBLigneExiste(GS.Row);
  if Code = '100' then // Allez sur l'en-tête
  begin
    if IndiqueOuSaisie = saEcheance then
    begin
      Cancel := False;
      ACol := FormEcheance.GSReg.Col;
      ARow := FormEcheance.GSReg.Row;
      FormEcheance.GSRegCellExit(FormEcheance.GSReg, ACol, ARow, Cancel);
      if Cancel then Exit;
      FormEcheance.GSRegRowExit(FormEcheance.GSReg, ARow, Cancel, False);
      if Cancel then Exit;
    end;
    GotoEntete;
  end else
    if Code = '101' then // Commentaire d'en-tête
  begin
    if PieceContainer.Action = taConsult then Exit;
    if IndiqueOuSaisie <> saLigne then Exit;
    InsertComment(True);
  end else
    if Code = '102' then // Complément d'en-tête
  begin
    CpltEnteteClick(Self);
  end else
    if Code = '103' then // Zones libres de la pièce
  begin
    LibrepieceClick(Self);
  end else
    if Code = '104' then // Saisir le régime de taxe
  begin
    ChangeRegime;
  end else
    if Code = '105' then // Saisir le code de taxe
  begin
    ChangeTVA;
  end else
    if Code = '106' then // Demande de détaxe
  begin
    if IndiqueOuSaisie = saEcheance then
      FormEcheance.TraiteDetaxe
    else
      TraiteDemandeDetaxe;
  end else
    if Code = '107' then // Saisir l'exception de taxation
  begin
{$IFDEF TAXEUS}
    ExceptionTaxe;
{$ENDIF}
  end else
    if Code = '110' then // Allez sur la ligne
  begin
    if FormEcheance <> nil then
    begin
      if FormEcheance.GSReg.CanFocus then FormEcheance.GSReg.SetFocus;
    end else
    begin
      if GS.CanFocus then GS.SetFocus;
    end;
  end else
    if Code = '111' then // Complément de la ligne
  begin
    if IndiqueOuSaisie <> saLigne then Exit;
    // Choix de la ligne sur laquelle saisir le complément
    for NoLig := GS.Row downto 0 do if GetCodeArtUnique(TOBPiece, NoLig) <> '' then Break;
    if NoLig < 0 then Exit;
    ComplementLigne(NoLig);
  end else
    if Code = '112' then // Insérer sous-total
  begin
    if PieceContainer.Action = taConsult then Exit;
    if IndiqueOuSaisie <> saLigne then Exit;
    if ClickSousTotal then FOSimuleClavier(VK_DOWN);
  end else
    if Code = '113' then // Rechercher dans la pièce
  begin
    if IndiqueOuSaisie <> saLigne then Exit;
    BChercherClick(Self);
  end else
    if Code = '114' then // Ligne de retour
  begin
    if PieceContainer.Action = taConsult then Exit;
    if IndiqueOuSaisie <> saLigne then Exit;
    // Choix de la ligne sur laquelle appliquer le retour
    for NoLig := GS.Row downto 0 do if GetCodeArtUnique(TOBPiece, NoLig) <> '' then Break;
    if NoLig < 0 then Exit;
    RetourneArticle(NoLig);
  end else
    if Code = '115' then // Descriptif détaillé de l'article
  begin
    if IndiqueOuSaisie <> saLigne then Exit;
    for NoLig := GS.Row downto 0 do if GetCodeArtUnique(TOBPiece, NoLig) <> '' then Break;
    if NoLig < 0 then Exit;
    BDescriptif.Down := not (BDescriptif.Down);
    GereDescriptif(NoLig, True);
    TDescriptif.Visible := BDescriptif.Down;
  end else
    if Code = '116' then // Recherche article par image
  begin
  {$IFDEF MODE}
    RechercheArticleImage;
  {$ENDIF}
  end else
    if Code = '117' then // Modification du dépôt de la ligne
  begin
    if PieceContainer.Action = taConsult then Exit;
    if IndiqueOuSaisie <> saLigne then Exit;
    // Choix de la ligne
    for NoLig := GS.Row downto 0 do if GetCodeArtUnique(TOBPiece, NoLig) <> '' then Break;
    if NoLig < 0 then Exit;
    ChangeDepot(NoLig);
  end else
    if Code = '120' then // Allez sur le pied
  begin
    if PieceContainer.Action = taConsult then Exit;
    if IndiqueOuSaisie <> saLigne then Exit;
    GotoPied;
  end else
    if Code = '121' then // Commentaire de pied
  begin
    if PieceContainer.Action = taConsult then Exit;
    if IndiqueOuSaisie <> saLigne then Exit;
    InsertComment(False);
  end else
    if Code = '130' then // Reprise d'un ticket en attente
  begin
    BTicketAttenteClick(nil);
  end else
    if Code = '200' then // Zoom Article
  begin
    SimuleBoutonZoom(BZoomArticle, False, True);
  end else
    if Code = '201' then // Zoom Tiers
  begin
    SimuleBoutonZoom(BZoomTiers, True, False);
  end else
    if Code = '202' then // Justifier le tarif Zoom Tarif
  begin
    SimuleBoutonZoom(BJustifTarif, False, True);
  end else
    if Code = '203' then // Zoom Vendeur
  begin
    if IndiqueOuSaisie = saLigne then SimuleBoutonZoom(BZoomCommercial, False, True)
    else SimuleBoutonZoom(BZoomCommercial, False, False);
  end else
    if Code = '204' then // Zoom Tarif
  begin
    SimuleBoutonZoom(BZoomTarif, False, True);
  end else
    if Code = '207' then // Zoom Ecritures comptables
  begin
    SimuleBoutonZoom(BZoomEcriture, False, False);
  end else
    if Code = '208' then // Zoom Devise
  begin
    if IndiqueOuSaisie = saEcheance then
    begin
      for NoLig := FormEcheance.GSReg.Row downto 0 do if FormEcheance.GSReg.Cells[FormEcheance.SG_Mode, NoLig] <> '' then Break;
      if NoLig < 0 then Exit;
      CurLig := FormEcheance.GSReg.Row;
      FormEcheance.GSReg.Row := NoLig;
      FormEcheance.ZoomDevise;
      if FormEcheance.GSReg.Row <> CurLig then FormEcheance.GSReg.Row := CurLig;
    end else SimuleBoutonZoom(BZoomDevise, False, False);
  end else
    if Code = '209' then // Zoom Mode de paiement
  begin
    if IndiqueOuSaisie = saEcheance then
    begin
      for NoLig := FormEcheance.GSReg.Row downto 0 do if FormEcheance.GSReg.Cells[FormEcheance.SG_Mode, NoLig] <> '' then Break;
      if NoLig < 0 then Exit;
      CurLig := FormEcheance.GSReg.Row;
      FormEcheance.GSReg.Row := NoLig;
      FormEcheance.ZoomModePaie;
      if FormEcheance.GSReg.Row <> CurLig then FormEcheance.GSReg.Row := CurLig;
    end;
  end else
    if Code = '210' then // Détail échéance
  begin
    if IndiqueOuSaisie = saEcheance then
    begin
      for NoLig := FormEcheance.GSReg.Row downto 0 do if FormEcheance.GSReg.Cells[FormEcheance.SG_Mode, NoLig] <> '' then Break;
      if NoLig < 0 then Exit;
      CurLig := FormEcheance.GSReg.Row;
      FormEcheance.GSReg.Row := NoLig;
      FormEcheance.DetailLigneClick(Self);
      if FormEcheance.GSReg.Row <> CurLig then FormEcheance.GSReg.Row := CurLig;
    end;
  end else
    if Code = '211' then // Détail nomenclature
  begin
    if PieceContainer.Action = taConsult then Exit;
    if IndiqueOuSaisie <> saLigne then Exit;
    SimuleChoixLigneMenu(MBDetailNomen, False, True);
  end else
    if Code = '212' then // Création d'un client
  begin
    CreationTiers;
  end else
    if Code = '213' then // Modification du client
  begin
    ModificationTiers;
  end else
    if Code = '214' then // Recherche du client
  begin
    ChoixTiers;
  end else
    if Code = '215' then // Zoom Stock de l'article
  begin
    SimuleBoutonZoom(BZoomStock, False, True);
  end else
    if Code = '216' then // Consultation du stock local
  begin
    ConsultationStock(False);
  end else
    if Code = '217' then // Consultation du stock à distance
  begin
    ConsultationStock(True);
  end else
    if Code = '220' then // Saisir le client
  begin
    if GP_TIERS.CanFocus then
    begin
      if (IndiqueOuSaisie = saLigne) or (IndiqueOuSaisie = saEcheance) then
        RetourGS := True;
      GP_TIERS.SetFocus;
    end;
  end else
    if Code = '221' then // Saisir le vendeur
  begin
    if GP_REPRESENTANT.CanFocus then
    begin
      if (IndiqueOuSaisie = saLigne) or (IndiqueOuSaisie = saEcheance) then
        RetourGS := True;
      GP_REPRESENTANT.SetFocus;
    end;
  end else
    if Code = '222' then // Solde du client
  begin
    VoirSoldeTiers;
  end else
    if Code = '223' then // Fidelité du client
  begin
    MCFideliteClick(nil);
  end else
    if Code = '224' then // Contacts du client
  begin
    MCContactClientClick(nil);
  end else
    if Code = '225' then // Historique des achats
  begin
    MCHistoClientClick(nil);
  end else
    if Code = '226' then // Liste des articles
  begin
    MCArticleClientClick(nil);
  end else
    if Code = '250' then // Solder la ligne d'échéance
  begin
    if PieceContainer.Action = taConsult then Exit;
    if IndiqueOuSaisie = saEcheance then FormEcheance.BSoldeClick(Self);
  end else
    if Code = '251' then // Remise globale sur le ticket
  begin
    if PieceContainer.Action = taConsult then Exit;
    if IndiqueOuSaisie = saEcheance then FormEcheance.SaisieRemiseTicket(0, '');
  end else
    if Code = '300' then // Allez sur la quantité
  begin
    if SG_QF <> -1 then AllerSur(SG_QF) else AllerSur(SG_QS);
  end else
    if Code = '301' then // Allez sur le prix
  begin
    if SG_PxNet <> -1 then AllerSur(SG_PxNet) else AllerSur(SG_Px);
  end else
    if Code = '302' then // Allez sur le % de remise
  begin
    if IndiqueOuSaisie = saRemiseTicket then FormEcheance.AllerSurChpRemise(rtPourcent)
    else AllerSur(SG_Rem);
  end else
    if Code = '303' then // Allez sur la remise
  begin
    if IndiqueOuSaisie = saRemiseTicket then FormEcheance.AllerSurChpRemise(rtMontant) else
      if SG_RV <> -1 then AllerSur(SG_RV) else
      if SG_RL <> -1 then AllerSur(SG_RL);
  end else
    if Code = '304' then // Allez sur le montant
  begin
    if IndiqueOuSaisie = saRemiseTicket then FormEcheance.AllerSurChpRemise(rtTotal)
    else AllerSur(SG_Montant);
  end else
    if Code = '305' then // Allez sur le code
  begin
    AllerSur(SG_RefArt);
  end else
    if Code = '306' then // Allez sur le libellé
  begin
    AllerSur(SG_Lib);
  end else
    if Code = '390' then // Zoom Pièce précédente
  begin
    SimuleBoutonZoom(BZoomPrecedente, False, True);
  end else
    if Code = '391' then // Zoom Pièce suivante
  begin
    SimuleBoutonZoom(BZoomSuivante, False, False);
  end else
    if Code = '392' then // Zoom Pièce d'origine
  begin
    SimuleBoutonZoom(BZoomOrigine, False, True);
  end else
    if Code = '500' then // Ligne précédente
  begin
    FOSimuleClavier(VK_UP);
  end else
    if Code = '501' then // Ligne suivante
  begin
    FOSimuleClavier(VK_DOWN);
  end else
    if Code = '502' then // Insertion d'une ligne
  begin
    if PieceContainer.Action = taConsult then Exit;
    if IndiqueOuSaisie = saLigne then ClickInsert(GS.Row) else
      if IndiqueOuSaisie = saEcheance then FormEcheance.bNewligneClick(Self);
  end else
    if Code = '503' then // Suppression d'une ligne
  begin
    if PieceContainer.Action = taConsult then Exit;
    if IndiqueOuSaisie = saLigne then
    begin
      // Choix de la ligne à supprimer
      for NoLig := GS.Row downto 0 do
      begin
        if (GetCodeArtUnique(TOBPiece, NoLig) <> '') or
          (GetChampLigne(TobPiece, 'GL_LIBELLE', NoLig) <> '') then Break;
      end;
      if NoLig > 0 then ClickDel(NoLig, True, True);
      if (NoLig >= GS.FixedRows) and (NoLig < GS.Row) then GS.Row := NoLig;
    end else
      if IndiqueOuSaisie = saEcheance then FormEcheance.SupprimeLigne; //FormEcheance.bDelLigneClick(Self) ;
  end else
    if Code = '510' then // Champ précédent
  begin
    FOSimuleClavier(VK_TAB, True);
  end else
    if Code = '511' then // Champ suivant
  begin
    if (IndiqueOuSaisie = saRemiseTicket) and (FormEcheance.RemTicketCtrl.CanFocus) then FormEcheance.RemTicketCtrl.SetFocus;
    FOSimuleClavier(VK_TAB);
  end else
    if Code = '512' then // Validation du champ
  begin
    FOSimuleClavier(VK_RETURN);
  end else
    if Code = '513' then // DoubleClick sur le champ
  begin
    if (Screen.ActiveControl is THValComboBox) then
    begin
      if Assigned(THValComboBox(Screen.ActiveControl).OnDblClick) then THValComboBox(Screen.ActiveControl).OnDblClick(Self);
    end else
      if (Screen.ActiveControl is THCritMaskEdit) then
    begin
      if Assigned(THCritMaskEdit(Screen.ActiveControl).OnDblClick) then THCritMaskEdit(Screen.ActiveControl).OnDblClick(self);
    end else
      if (Screen.ActiveControl is THGrid) then
    begin
      if Assigned(THGrid(Screen.ActiveControl).OnDblClick) then THGrid(Screen.ActiveControl).OnDblClick(Self);
    end;
  end else
    if Code = '540' then // Clavier
  begin
    FOLanceClavierVisuel;
  end else
    if Code = '550' then // Recherche
  begin
    if Screen.ActiveControl is THValComboBox then FOSimuleClavier(VK_F4)
    else FOSimuleClavier(VK_RECHERCHE);
  end else
    if Code = '551' then // Validation de la pièce
  begin
    FOSimuleClavier(VK_VALIDE);
  end else
    if Code = '552' then // Abandon de la pièce
  begin
    if IndiqueOuSaisie = saEcheance then FormEcheance.ClickAbandon
    else BAbandonClick(nil);
  end else
    if Code = '553' then // Impression de la pièce
  begin
    if FOPieceUneLigne(TOBPiece) then FOImprimerLaPiece(TOBPiece, TOBTiers, TOBEches, TOBComms, PieceContainer.CleDoc, False, False);
  end else
    if Code = '554' then // Aide
  begin
    FOSimuleClavier(VK_F1);
  end else
    if Code = '555' then // Envoyer un message
  begin
    {$IFNDEF EAGLCLIENT}
    SendCourrier('', 0);
    {$ENDIF}
  end else
    if Code = '556' then // Aperçu avant Impression de la pièce
  begin
    if FOPieceUneLigne(TOBPiece) then FOImprimerLaPiece(TOBPiece, TOBTiers, TOBEches, TOBComms, PieceContainer.CleDoc, False, True);
  end else
    if Code = '557' then // Choix du modèle d'impression de la pièce
  begin
    ToucheModele('', '', '', 0, 0);
  end else
    if Code = '558' then // Situation Flash
  begin
    MDSituFlashClick(Self);
  end else
    if Code = '559' then // Réimpression du dernier ticket
  begin
    FOReImpTicket(False);
  end else
    if Code = '560' then // Ouverture du tiroir caisse
  begin
    MDOuvreTiroirClick(Self);
  end else
    if Code = '561' then // Réimpression du dernier ticket avec choix du modèle
  begin
    FOReImpTicket(True);
  end else
    if Code = '562' then // Réimpression des bons du dernier ticket
  begin
    FOReImprimeBons;
  end else
    if Code = '600' then // Lecture d'une carte depuis la piste du clavier
  begin
    LectureCBFromKB;
  end else
    if Code = '610' then // Vente à livrer
  begin
    TicketALivrer;
  end else
    ;
end;

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

procedure TFTicketFO.SaisieClavierEcran(Concept, Code, Extra: string; Qte, Prix: Double);
begin
  if InTraiteBouton then Exit;
  InTraiteBouton := True;
  if IndiqueOuSaisie = saEcheance then
    FormEcheance.TOBEcheLigneExiste
  else
    TOBLigneExiste(GS.Row);
  if Concept = 'ART' then // saisie d'un article
  begin
    ToucheArticle(Concept, Code, Extra, Qte, Prix);
  end else
    if Concept = 'OCT' then // saisie d'une opération de caisse
  begin
    ToucheArticle(Concept, Code, Extra, Qte, Prix);
  end else
    if Concept = 'COM' then // saisie d'un commentaire
  begin
    ToucheCommentaire(Concept, Code, Extra, Qte, Prix);
  end else
    if Concept = 'REM' then // saisie d'une remise ligne
  begin
    ToucheRemise(Concept, Code, Extra, Qte, Prix);
  end else
    if Concept = 'VEN' then // changement de vendeur
  begin
    ToucheVendeur(Concept, Code, Extra, Qte, Prix);
  end else
    if Concept = 'REG' then // saisie d'un mode de paiement
  begin
    ToucheModePaie(Concept, Code, Extra, Qte, Prix);
  end else
    if Concept = 'FON' then // executer une fonction
  begin
    DispatchFonctionClavier(Concept, Code, Extra, Qte, Prix);
  end else
    if Concept = 'MOD' then // choix d'un modèle d'impression
  begin
    ToucheModele(Concept, Code, Extra, Qte, Prix);
  end else
{$IFDEF CHR}
if Concept = 'RES' then       // choix d'une ressource
   begin
    ToucheRes(Concept, Code, Extra) ;
    TransfertFac:=True;
    if (Code = '') then
    begin
      if IndiqueOuSaisie <> saEcheance then
        DispatchFonctionClavier(Concept, '551', Extra, Qte, Prix) else
      begin
        FormEcheance.FAPayer.Value:=0;
      end;
    end else
    begin
      if TOBPiece.GetValue('GP_TOTALTTCDEV') = 0 then
      begin
        DejaSaisie:=True;// pour ne pas sortir
      end;
      if IndiqueOuSaisie <> saEcheance then
        DispatchFonctionClavier(Concept, '552', Extra, Qte, Prix) else
      begin
        DispatchFonctionClavier(Concept, '551', Extra, Qte, Prix);
      end;
    end;
   end else
{$ENDIF}
  ;
  /////////////if PnlBoutons.PageCourante <> 0 then PnlBoutons.AlleraPage('FST') ;
  InTraiteBouton := False;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  SaisieCalculette : prend en compte un chiffre saisi sur la calculette
///////////////////////////////////////////////////////////////////////////////////////

procedure TFTicketFO.SaisieCalculette(Val: Double);
var ACol, ARow: integer;
  CC: TWinControl;
  Texte, sDate: string;
  Ou: TOuSaisie;
  Key: Word;
begin
  if PieceContainer.Action = taConsult then Exit;
  if InTraiteBouton then Exit;
  InTraiteBouton := True;
  Texte := FloatToStr(Val);
  Key := VK_TAB;
  Ou := IndiqueOuSaisie;
  case Ou of
    saRemiseTicket:
      begin
        if FormEcheance.RemTicketCtrl.CanFocus then
        begin
          FormEcheance.RemTicketCtrl.SetFocus;
          CC := FormEcheance.RemTicketCtrl;
          if CC is TCustomEdit then TCustomEdit(CC).Text := Texte;
          if CC is THValComboBox then THValComboBox(CC).Value := Texte;
        end;
      end;
    saEcheance:
      begin
        ACol := FormEcheance.GSReg.Col;
        ARow := FormEcheance.GSReg.Row;
        if ACol = FormEcheance.SG_Ech then
        begin
          // on attend une date avec des séparateurs
          sDate := '';
          if (Length(Texte) = 3) or (Length(Texte) = 5) then Texte := '0' + Texte;
          if Length(Texte) = 4 then Texte := Texte + FODonneAn(0);
          if Length(Texte) >= 6 then
          begin
            sDate := Copy(Texte, 1, 2) + DateSeparator + Copy(Texte, 3, 2) + DateSeparator;
            if Length(Texte) >= 8 then sDate := sDate + Copy(Texte, 5, 4)
            else sDate := sDate + Copy(Texte, 5, 2);
          end;
          if IsValidDate(sDate) then FormEcheance.GSReg.Cells[ACol, ARow] := sDate;
        end else FormEcheance.GSReg.Cells[ACol, ARow] := Texte;
      end;
    saLigne:
      begin
        ACol := Gs.Col;
        ARow := Gs.Row;
        Gs.Cells[ACol, ARow] := Texte;
        if ACol = SG_RefArt then Key := VK_DOWN;
      end;
  else
    begin
      CC := Screen.ActiveControl;
      if CC is TCustomEdit then TCustomEdit(CC).Text := Texte;
      if CC is THValComboBox then THValComboBox(CC).Value := Texte;
    end;
  end;
  FOSimuleClavier(Key);
  InTraiteBouton := False;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BoutonCalculetteClick : interprète les boutons de la calculette
///////////////////////////////////////////////////////////////////////////////////////

procedure TFTicketFO.BoutonCalculetteClick(Val: string);
begin
  if InTraiteBouton then Exit;
  InTraiteBouton := True;
  if Val = 'ENTER' then
  begin
    if (AnnulPiece) or (AssignePiece) then ClickValideTicket else
      case IndiqueOuSaisie(True) of
        saRemiseTicket: FormEcheance.BValiderRemiseClick(nil);
        saEcheance: FormEcheance.ClickValide;
      else ClickValideTicket;
      end;
  end else
    if Val = 'CLEAR' then
  begin
    case IndiqueOuSaisie(True) of
      saRemiseTicket: FormEcheance.BCancelRemiseClick(nil);
      saEcheance:
        begin
          FormEcheance.ClickAbandon;
          DepuisEcheance := False;
        end;
    else BAbandonClick(nil);
    end;
  end;
  InTraiteBouton := False;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  AffichePaveChoisi : affichage de la page du pavé affectée au champ
///////////////////////////////////////////////////////////////////////////////////////

procedure TFTicketFO.AffichePaveChoisi(ACol: Integer);
begin
  if PnlBoutons = nil then Exit;
  if (ACol < Low(ColCarac)) or (ACol > High(ColCarac)) then Exit;
  AffichePaveContextuel(PnlBoutons, ColCarac[ACol].NoPage);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ChargeChoixDesPaves
///////////////////////////////////////////////////////////////////////////////////////

procedure TFTicketFO.ChargeChoixDesPaves(var PropPv2: RPropPave2);
begin
  ChargePaveContextuel([PEntete, PPied], [GS], [fAffTXT, PI_NETAPAYERDEV], ColCarac, [GP_TIERS, GP_REPRESENTANT, GP_DATEPIECE], PropPv2);
end;

{*****************************************************************
Auteur  ...... : JTR
Créé le ...... : 01/09/2004
Modifié le ... :
Description .. : Gestion des formules de quantité
Mots clefs ... : FO;ENCAISSEMENT DE CREDIT
*****************************************************************}
Procedure TFTicketFO.TraiteFormuleQte (ARow : integer) ;
Var FoncQte : R_FonctCal ;
    Qte : double ;
    Formule,Flux : string;
    TOBA,TOBADet,TOBL : TOB ;
    ind : integer;
begin
  TOBA:=FindTOBArtRow(TOBPiece,TOBArticles,ARow) ;
  if TOBA=Nil then exit;
  TOBADet:=TOBAFormule.FindFirst(['GAF_ARTICLE'],[TOBA.GEtValue('GA_ARTICLE')],True) ;
  if TOBADet=nil then exit;
  if VenteAchat='VEN' then
    Flux:='VTE'
  else if VenteAchat='ACH' then
    Flux:='ACH'
  else exit;
  Formule:=TOBADet.GetValue('GAF_FORMULEQTE'+Flux);
  If Formule='' then exit;
  if TOBADet.Detail.Count=0 then
  begin
    FoncQte:=EvaluationFormule (Formule,Formule);
  end else
  begin
    FoncQte.Formule:=Formule;
    FoncQte.Affichage:=Formule;
    for ind:=0 to TOBADet.Detail.Count-1 do
    begin
      FoncQte.VarLibelle[ind]:=TOBADet.Detail[ind].GetValue('GAV_LIBELLEVAR');
      FoncQte.VarValeur[ind]:=TOBADet.Detail[ind].GetValue('GAV_VALEURVAR');
      FoncQte.Modifiable[ind]:=(TOBADet.Detail[ind].GetValue('GAV_ENABLEVAR')='X');
      FoncQte.Affichable[ind]:=(TOBADet.Detail[ind].GetValue('GAV_VISIBLEVAR')='X');
    end;
    FoncQte:=EvaluationFormule (ActionToString(PieceContainer.Action),FoncQte);
  end;
  if FoncQte.Annule then exit;
  Qte:=FoncQte.Resultat ;
  if SG_QS>0 then
  begin
    GS.Cells[SG_QS,ARow]:=StrF00(Qte,V_PGI.OkDecQ) ;
    TraiteQte(SG_QS,ARow) ;
  end;
  if SG_QF>0 then
  begin
    GS.Cells[SG_QF,ARow]:=StrF00(Qte,V_PGI.OkDecQ) ;
    TraiteQte(SG_QF,ARow) ;
  end;
  TOBL:=GetTOBLigne(TOBPiece,ARow) ;
  if TOBL=Nil then Exit ;
  InsertFormuleToTOB(TOBL,TOBLigFormule,FoncQte);
end;

{*****************************************************************
Auteur  ...... : JTR
Créé le ...... : 01/09/2004
Modifié le ... :
Description .. : Rappel des formules de quantité
Mots clefs ... : FO;ENCAISSEMENT DE CREDIT
*****************************************************************}
Procedure TFTicketFO.RappelFormuleQte (ARow : integer) ;
Var FoncQte : R_FonctCal ;
    Qte : double ;
    Formule,Flux : string;
    TOBDet,TOBL : TOB ;                                 
    ind,IndiceFormule : integer;
begin
  TOBL:=GetTOBLigne(TOBPiece,ARow) ; if TOBL=Nil then Exit ;
  if TOBL.FieldExists('GL_INDICEFORMULE') then
    IndiceFormule:=TOBL.GetValue('GL_INDICEFORMULE')
    else
    IndiceFormule:=0;
  if IndiceFormule=0 then exit;
  if TOBLigFormule.Detail.Count<IndiceFormule then exit;
  TOBDet:=TOBLigFormule.Detail[IndiceFormule-1];
  if VenteAchat='VEN' then
    Flux:='VTE'
  else if VenteAchat='ACH' then
    Flux:='ACH'
  else
    exit;
  Formule:=TOBDet.GetValue('GLF_FORMULE');
  If Formule='' then exit;
  if TOBDet.Detail.Count=0 then
  begin
    FoncQte:=EvaluationFormule (Formule,Formule);
  end else
  begin
    FoncQte.Formule:=Formule;
    FoncQte.Affichage:=Formule;
    for ind:=0 to TOBDet.Detail.Count-1 do
    begin
      FoncQte.VarLibelle[ind]:=TOBDet.Detail[ind].GetValue('GLF_LIBELLEVAR');
      FoncQte.VarValeur[ind]:=TOBDet.Detail[ind].GetValue('GLF_VALEURVAR');
      FoncQte.Modifiable[ind]:=(TOBDet.Detail[ind].GetValue('GLF_ENABLEVAR')='X');
      FoncQte.Affichable[ind]:=(TOBDet.Detail[ind].GetValue('GLF_VISIBLEVAR')='X');
    end;
    FoncQte:=EvaluationFormule (ActionToString(PieceContainer.Action),FoncQte);
  end;
  if FoncQte.Annule then exit;
  Qte:=FoncQte.Resultat ;
  if SG_QS>0 then
  begin
    GS.Cells[SG_QS,ARow]:=StrF00(Qte,V_PGI.OkDecQ) ;
    TraiteQte(SG_QS,ARow) ;
  end;
  if SG_QF>0 then
  begin
    GS.Cells[SG_QF,ARow]:=StrF00(Qte,V_PGI.OkDecQ) ;
    TraiteQte(SG_QF,ARow) ;
  end;
  UpdateFormuleToTOB(TOBL,TOBLigFormule,FoncQte);
end;

procedure InitFacGescom();
begin
  RegisterAglProc('AppelTicketFO', False, 2, AppelTicketFO);
  RegisterAglProc('AppelTransformeTicketFO', False, 2, AppelTransformeTicketFO);
  RegisterAglProc('AppelDupliqueTicketFO', False, 2, AppelDupliqueTicketFO);
  RegisterAglProc('AppelAnnuleTicketFO', False, 2, AppelAnnuleTicketFO);
  RegisterAglProc('AppelAssigneTicketFO', False, 2, AppelAssigneTicketFO);
end;

{$IFDEF GESCOM}
{*****************************************************************
Auteur  ...... : JTR
Créé le ...... : 17/02/2004
Modifié le ... :
Description .. : Empêche le cumul des articles financier avec des marchandises
Suite ........ : sauf les encaissement de crédit.
Mots clefs ... : FO;ENCAISSEMENT DE CREDIT
*****************************************************************}
function TFTicketFO.CumulArtOk(Acol, Arow : integer; Cancel : boolean) : boolean;
var TypeArticle, TypeFi : string;
    TobA : TOB;
    Ok : boolean;
begin
  Result := True;
  if Cancel then
    Result := False
  else if GS.Cells[ACol, ARow] <> '' then
  begin
    TobA := TOBArticles.FindFirst(['GA_ARTICLE'], [GetChampLigne(TOBPiece, 'GL_ARTICLE', ARow)], False);
    if TobA <> nil then
      TypeFi := TobA.GetValue('GA_TYPEARTFINAN')
      else
      TypeFi := '';
    if (TypePiece = '') or (TobPiece.Detail.count = 1) then
    begin
      TypePiece := GetChampLigne(TOBPiece, 'GL_TYPEARTICLE', ARow);
      TypePieceFi := TypeFi;
    end else
    begin
      Ok := True;
      TypeArticle := GetChampLigne(TOBPiece, 'GL_TYPEARTICLE', ARow);
      if (TypePiece='FI') and (TypePieceFi<>'ECR') then //La pièce est une pièce financière autre que ECR
      begin
        if (TypeArticle<>'FI') or ((TypeArticle='FI') and (TypeFi='ECR')) then
         Ok := False
         else
         Ok := True;
      end else
      if (TypePiece='FI') and (TypePieceFi='ECR') then
      begin
        if (TypeArticle='FI') and (TypeFi<>'ECR') then
          Ok := False
          else
          Ok := True;
      end else
      if TypePiece <>'FI' then
      begin
        if (TypeArticle='FI') and (TypeFI='ECR') then
          Ok := True
        else if (TypeArticle='FI') and (TypeFI<>'ECR') then
          Ok := False;
      end;
      if not Ok then
      begin
        if not MultiSel_SilentMode then // eQualité 11973
          PGIBox('Impossible de cumuler ce type d''article avec le(s) article(s) déjà présent(s)', 'Saisie d''un ticket');
        Result := False;
        ClickDel(ARow, True, True);
        if not MultiSel_SilentMode then // eQualité 11973
          StCellCur := GS.Cells[ACol, ARow];
      end ;
    end;
  end;
end;

{*****************************************************************
Auteur  ...... : JTR - Démarque obligatoire
Créé le ...... : 17/02/2004
Modifié le ... :
Description .. : En cas de paramètrage de démarque obligatoire en cas de remise
Suite ........ : mémorise la remise du tarif s'il y en a une pour, lors de la validation,
Suite ........ : la comparer à la remise de la ligne.
Suite ........ : Si la remise ligne est différente, il faut saisir la démarque.
Mots clefs ... : FO;REMISE;DEMARQUE
*****************************************************************}
procedure TFTicketFO.RemiseDuTarif(Arow : integer);
var TobL : TOB;
begin
  if TobPiece = nil then exit;
  if TobPiece.Detail.count = 0 then exit;
  if Arow < 0 then exit;
  TobL := TobPiece.detail[Arow-1];
  if not TobL.FieldExists('REMISEDUTARIF') then Tobl.AddChampSup('REMISEDUTARIF',False);
  TobL.PutValue('REMISEDUTARIF',TobL.GetValue('GL_REMISELIGNE'));
end;
{$ENDIF GESCOM}

// eQualité 11973
function TFTicketFO.InsertNewArtInGS(RefArt, Designation: string; Qte: Double) : boolean;
var ACol, ARow, ARowGEN: Integer;
    Cancel, bc: Boolean;
    Etat, CodeGen : string;
    TOBArt, TOBLGEN, TOBL: TOB;
begin
  Result := true;
  Cancel := False;
  bc := False;
  ACol := SG_RefArt;
  TOBLGEN := nil;
  if (TOBPiece.Detail.Count = 1) and (GS.Cells[SG_RefArt, 1] = '') then
    ARow := 1
  else if (GS.Cells[SG_RefArt, TOBPiece.Detail.Count] = '') then
    ARow := TOBPiece.Detail.Count
  else
    ARow := TOBPiece.Detail.Count + 1;
  TOBArt := FindTOBArtSais(TOBArticles, RefArt);
  if (TOBArt <> nil) then
  begin
    if not ArticleAutorise(TOBPiece, TOBArticles, PieceContainer.CleDoc.NaturePiece, ARow, RefArt) then
    begin
      Result := false;
      exit;
    end;
    if not CumulArtOk(Acol,Arow,Cancel) then // eQualité 11973
    begin
      Result := false;
      exit;
    end;
    CodeGen := TOBArt.GetValue('GA_CODEARTICLE');
    Etat := TOBArt.GetValue('GA_STATUTART');
    if Etat = 'UNI' then Etat := 'NOR';
    if (Etat = 'DIM') then
    begin
      TOBLGEN := TOBPiece.FindFirst(['GL_CODESDIM', 'GL_TYPEDIM'], [CodeGen, 'GEN'], False);
      if TOBLGEN = nil then
      begin
        if ARow = 1 then
          GSEnter(GS)
        else if ARow = TOBPiece.Detail.Count + 1 then
        begin
          GSRowEnter(GS, ARow, bc, False);
          GS.Col := ACol;
          GS.Row := ARow;
          GSCellEnter(GS, ACol, ARow, Cancel);
        end;
        GS.Cells[SG_RefArt, ARow] := CodeGen;
        TOBLGEN := InsereLigneGEN(CodeGen, ARow, 0);
        if TOBLGEN = nil then
          exit;
        if DimSaisie = 'DIM' then
          GS.RowHeights[ARow] := 0;
        ARow := ARow + 1;
      end else
        ARow := TOBLGEN.GetValue('GL_NUMLIGNE') + NbDimensionsLigneGEN(TOBLGEN.GetValue('GL_NUMLIGNE')) + 1;
    end;
    if ARow = 1 then
      GSEnter(GS)
    else if ARow = TOBPiece.Detail.Count + 1 then
    begin
      GSRowEnter(GS, ARow, bc, False);
      GS.Col := ACol;
      GS.Row := ARow;
      GSCellEnter(GS, ACol, ARow, Cancel);
    end else
      ClickInsert(ARow);
    if (Etat <> 'DIM') then
      GS.Cells[SG_RefArt, ARow] := CodeGen;
    GS.Col := ACol;
    GS.Row := ARow;
    if Etat = 'GEN' then
    begin
      StCellCur := '';
      GSCellExit(nil, ACol, ARow, Cancel);
      if Cancel then
      begin
        GS.Cells[SG_RefArt, ARow] := '';
        DepileTOBLignes(GS, TOBPiece, ARow, 1);
        Result := false;
      end;
      exit;
    end;
    CodesArtToCodesLigne(TOBArt, ARow);
    ChargeTOBDispo(ARow);
    TOBL := GetTOBLigne(TOBPiece, ARow);
    TOBL.PutValue('GL_REFARTSAISIE', TOBL.GetValue('GL_CODEARTICLE'));
    InitLaLigne(ARow, Qte);
    if TOBL <> nil then
    begin
{$IFDEF STK}
      CreerTobGQDServiLigne (TOBL, PieceContainer.TCGQDServi);
{$ENDIF STK}
      TOBL.PutValue('GL_TYPEDIM', Etat);
      TOBL.PutValue('GL_LIBELLE', Designation);
      GS.Cells[SG_Lib, ARow] := Designation;
      TOBL.PutValue('GL_REFARTSAISIE', TOBL.GetValue('GL_CODEARTICLE'));
      TOBL.PutValue('GL_FOURNISSEUR', TOBArt.GetValue('GA_FOURNPRINC'));
    end;
    TraiteLesCons(ARow);
    TraiteLesNomens(ARow);
    TraiteLesMacros(ARow);
    if GS.Cells[SG_RefArt, ARow] = '' then
      exit;   // JSI 29/09/04 pour ClickDel
    TraiteLesLies(ARow);
    TraiteLaCompta(ARow);
    TraiteLeCata(ARow);
    TraiteRupture(ARow,False);
    TraiteFormuleQte(ARow); // JTR - eQualité 11203 (formules de qté)
{$IFDEF STK}
     CreerTobGQDServiLigneNomenclature(TOBL,TobNomenclature,TobArticles, PieceContainer);
{$ENDIF}
    if (PieceContainer.ConsTraite) and (GetChampLigne(TOBPiece,'GL_TYPEARTICLE',Arow) = 'CNS') then
    begin
      if SG_QF >= 0 then
        TraiteQte(SG_QF, ARow);
      if SG_QS >= 0 then
        TraiteQte(SG_QS, ARow);
    end;
    TOBPiece.PutValue('GP_RECALCULER', 'X');
    CalculeLaSaisie(SG_RefArt, ARow, False);
    if (Etat = 'DIM') then
    begin
      if DimSaisie = 'GEN' then
        GS.RowHeights[ARow] := 0;
      AfficheLaLigne(ARow);
      if TOBLGEN <> nil then
        ARowGEN := TOBLGEN.getvalue('GL_NUMLIGNE')
        else
        ARowGEN := 0;
      if ARowGEN <> 0 then
      begin
        TOBPiece.PutValue('GP_RECALCULER', 'X');
        CalculeLaSaisie(SG_RefArt, ARowGEN, False);
      end;
    end;
  end;
end;

procedure TFTicketFO.MultiSelectionArticle(TobMultiSel : TOB);
var RefUnique, LibRef, Msg : string;
    ARow, iInd, QteInit : integer;                                                               
    Cancel,InfoLotSerie, InfoRupture : Boolean;
    TOBArt, TOBL : TOB;
    QQ : TQuery;                                     
begin
  Cancel := False;
  ARow := GS.Row;
  GS.Cells[SG_RefArt, ARow] := '';
  MultiSel_SilentMode := True;
  InfoLotSerie := False;
  InfoRupture := False;
  DepileTOBLignes(GS, TOBPiece, ARow, ARow);
  InitMoveProgressForm(nil,Caption,TraduireMemoire('Insertion de l''article : '),TobMultiSel.Detail.Count,True,True);
  for iInd := 0 to TobMultiSel.Detail.Count -1 do
  begin
    RefUnique := TobMultiSel.Detail[iInd].GetValue('ARTICLE');
    if RefUnique <> '' then
    begin
      TOBArt := FindTOBArtSais(TOBArticles, RefUnique);
      if TOBArt = nil then
      begin
        QQ := OpenSQL('SELECT * FROM ARTICLE WHERE GA_ARTICLE="' + RefUnique + '" ', True);
        if not QQ.EOF then
        begin
          TOBArt := CreerTOBArt(TOBArticles);
          ChargerTobArt(TOBArt,TobAFormule,TobAConsignes,PieceContainer.VenteAchat,'',QQ) ;
        end;
        Ferme(QQ);
      end;
      if TobArt = nil then continue;
      LibRef := TOBArt.GetValue('GA_LIBELLE');
      if not MoveCurProgressForm(TOBArt.GetValue('GA_CODEARTICLE') + ' ' + LibRef) then
        break;
      QteInit := 1;
      if not InsertNewArtInGS(RefUnique, LibRef, QteInit) then
        continue;
      GSRowExit(nil, GS.Row, Cancel, False);
      if Cancel then
        break;
      TOBL := GetTOBLigne(TOBPiece, GS.Row);
      if not InfoRupture then
        InfoRupture := (TOBL.GetValue('RUPTURE') = 'X');
    end;
  end;
  FiniMoveProgressForm;
  if InfoRupture then Msg := 'Les quantités signalées en rouge indiquent une rupture de stock.';
  if InfoLotSerie then
  begin
    if Msg <> '' then Msg := Msg + '#13';
    Msg := Msg + 'Certaines lignes n''ont pas de quantité. Une gestion particulière vous oblige à la saisir manuellement.';
  end;
  if Msg <> '' then PGIInfo(Msg,Caption);
  StCellCur := GS.Cells[GS.Col, GS.Row];
  CalculeLaSaisie(-1, -1, False);
  MultiSel_SilentMode := False;
end;                                                                                          

function TFTicketFO.InsereLigneGEN(CodeGen: string; ARow: integer; Qte: Double): TOB;
var TOBArt, TOBL: TOB;
  WhereNat, SelectFourniss: string;
  Q: TQuery;
  i: integer;
begin
  WhereNat := '';
  SelectFourniss := '';
  TOBArt := TOBArticles.FindFirst(['GA_CODEARTICLE', 'GA_STATUTART'], [CodeGen, 'GEN'], False);
  if TOBArt = nil then
  begin
    if (PieceContainer.VenteAchat = 'ACH') and (GetParamsoc('SO_MONOFOURNISS') = True)
        and (GetInfoParPiece(PieceContainer.NewNature, 'GPP_ARTFOURPRIN') = 'X') then
      SelectFourniss := GP_TIERS.Text;
    WhereNat := FabricWhereNatArt(PieceContainer.CleDoc.NaturePiece, '', SelectFourniss);
    if WhereNat <> '' then
      WhereNat := ' AND ' + WhereNat;
    Q := OpenSQL('SELECT * FROM ARTICLE WHERE GA_CODEARTICLE="' + CodeGen + '" AND GA_STATUTART="GEN"' + WhereNat, True);
    try
      if not Q.EOF then
      begin
        TOBArt := CreerTOBArt(TOBArticles);
        TOBArt.SelectDB('', Q);
      end else
        MessageAlerte(HTitres.Mess[38]);
    finally Ferme(Q);
    end;
  end;
  if TOBArt <> nil then
  begin
    CodesArtToCodesLigne(TOBArt, ARow);
    InitLaLigne(ARow, Qte);
    TOBL := GetTOBLigne(TOBPiece, ARow);
    if TOBL <> nil then
    begin
      TOBL.PutValue('GL_ARTICLE', '');
      TOBL.PutValue('GL_VALIDECOM', 'AFF');
      TOBL.PutValue('GL_FOURNISSEUR', GP_TIERS.Text);
      TOBL.PutValue('GL_CODEARTICLE', '');
      TOBL.PutValue('GL_TYPEREF', '');
      SetLigneCommentaire (TobPiece,TOBL,Arow);
      TOBL.PutValue('GL_TYPEARTICLE', '');
      TOBL.PutValue('GL_TENUESTOCK', '-');
      TOBL.PutValue('GL_REMISABLEPIED', '-');
      TOBL.PutValue('GL_TYPEDIM', 'GEN');
      TOBL.PutValue('GL_CODESDIM', CodeGen);
      TOBL.PutValue('GL_REFARTSAISIE', CodeGen);
      for i := 1 to 9 do TOBL.PutValue('GL_LIBREART' + IntToStr(i), '');
      TOBL.PutValue('GL_LIBREARTA', '');
      if (PieceContainer.CleDoc.NaturePiece = 'FCF') then TOBL.PutValue('GL_FOURNISSEUR', TOBArt.GetValue('GA_FOURNPRINC'));
    end;
  end
  else TOBL := nil;
  result := TOBL;
end;

function TFTicketFO.NbDimensionsLigneGEN(ARow: integer): integer;
var TOBL: TOB;
  Ligne, NbLignes: integer;
  CodeArticle: string;
  Continue: Boolean;
begin
  Result := 0;
  NbLignes := 0;
  Ligne := ARow;
  Continue := True;
  TOBL := GetTOBLigne(TOBPiece, Ligne);
  if TOBL = nil then
    exit;
  if TOBL.GetValue('GL_TYPEDIM') <> 'GEN' then
    exit;
  CodeArticle := TOBL.GetValue('GL_CODESDIM');
  while Continue do
  begin
    TOBL := GetTOBLigne(TOBPiece, Ligne + 1);
    if TOBL <> nil then
    begin
      if (TOBL.GetValue('GL_CODEARTICLE') = CodeArticle) and (TOBL.GetValue('GL_TYPEDIM') = 'DIM') then
      begin
        Ligne := Ligne + 1;
        NbLignes := NbLignes + 1;
      end else
        Continue := False;
    end else
      Continue := False;
  end;
  Result := NbLignes;
end;
// Fin eQualité 11973

initialization
  InitFacGescom();

end.
