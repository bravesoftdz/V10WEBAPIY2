unit afplanning;
         
interface

uses
  {$IFDEF VER150} Variants, {$ENDIF}
  Windows,
  Messages,
  SysUtils,
  Classes,                          
  Graphics,                    
  Controls,
  Forms,                             
  hDrawXp,

{$IFNDEF EAGLCLIENT}
{$IFNDEF DBXPRESS}dbtables {BDE},
{$ELSE}uDbxDataSet,                                            
{$ENDIF}
  db,
  FE_Main,
  Mul,               
{$ELSE}
  MaineAGL,
  emul,
{$ENDIF}
  StdCtrls,
  Grids,
  HTB97,
  ExtCtrls,
  HPanel,
  HPlanning,
  uTob,
  ComCtrls,
  Mask,
  HSysMenu,
  Menus,
  Hctrls,
  AFPlanningCst,
  Hent1, Buttons, TntButtons, TntStdCtrls, TntExtCtrls;

type

  THintWindowPlanning = class(THintWindow)
  public
    constructor Create(AOwner: TComponent); override;
  end;

  TPlanningGA = class;

  TAF_Planning = class(TForm)
    PPlanningJour: THPanel;
    POPUPGENERAL: TPopupMenu;
    MnClient: TMenuItem;
    MnAffaire: TMenuItem;
    PA_SELECTION: TPanel;
    LA_SELECTION: TLabel;
    MnTache: TMenuItem;
    MnPlanning: TMenuItem;
    MnEcole: TMenuItem;
    PA_COMMENTAIRE: TPanel;
    LA_Commentaire: TLabel;
    MnInterventions: TMenuItem;
    LA_Quantite: TLabel;
    LA_BASE: TLabel;
    MnListeInterv: TMenuItem;
    PA_Quick: TPanel;
    CB_RAP: THCheckbox;
    ED_TACHE: THCritMaskEdit;
    ED_AFF1: THCritMaskEdit;
    ED_AFF2: THCritMaskEdit;
    ED_AFF3: THCritMaskEdit;
    ED_AVE: THCritMaskEdit;
    BCHERCHEAFF: THBitBtn;
    ED_AFF: THCritMaskEdit;
    ED_TIERS: THCritMaskEdit;
    ED_AFF0: THCritMaskEdit;
    Dock972: TDock97;
    TB_NAVIGATION: TToolbar97;
    BFirst: TToolbarButton97;
    BPrev: TToolbarButton97;
    BNext: TToolbarButton97;
    BLast: TToolbarButton97;
    BCalendrier: TToolbarButton97;
    DateEdit: THCritMaskEdit;
    TB_SELECTION: TToolbar97;
    BTNonProductive: TToolbarButton97;
    BT_LIGNES: TToolbarButton97;
    BT_EFFACELIGNES: TToolbarButton97;
    TB_ZOOM: TToolbar97;
    BEXCEL: TToolbarButton97;
    BINTERVENANT: TToolbarButton97;
    BAFFAIRE: TToolbarButton97;
    TB_ACTIONS: TToolbar97;
    BT_REFRESH: TToolbarButton97;
    BPARAM: TToolbarButton97;
    Bprint: TToolbarButton97;
    Bannuler: TToolbarButton97;
    BAide: TToolbarButton97;
    HMTrad: THSystemMenu;
    BCHERCHEART: THBitBtn;

    procedure BannulerClick(Sender: TObject);
    procedure BPrevClick(Sender: TObject);
    procedure BNextClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure BLastClick(Sender: TObject);
    procedure BFirstClick(Sender: TObject);
    procedure BCalendrierClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BprintClick(Sender: TObject);
    procedure BParamClick(Sender: TObject);
    procedure BExcelClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    function GetNumModele: Integer;
    procedure MnAffaireClick(Sender: TObject);
    procedure MnClientClick(Sender: TObject);
    procedure BLignesClick(Sender: TObject);
    procedure BT_EFFACELIGNESClick(Sender: TObject);
    procedure MnTacheClick(Sender: TObject);
    procedure MnFactureClick(Sender: TObject);
    procedure MnPlanningClick(Sender: TObject);
    procedure BINTERVENANTClick(Sender: TObject);
    procedure BTNonProductiveClick(Sender: TObject);
    procedure MnEcoleClick(Sender: TObject);
    procedure BT_REFRESHClick(Sender: TObject);
    procedure MnInterventionsClick(Sender: TObject);
    procedure MnListeIntervClick(Sender: TObject);
    procedure ED_AVEExit(Sender: TObject);
    procedure ED_AFF1Change(Sender: TObject);
    procedure ED_AFF2Change(Sender: TObject);
    procedure ED_AFF3Change(Sender: TObject);
    procedure BCHERCHEAFFClick(Sender: TObject);
    procedure BCHERCHEARTClick(Sender: TObject);
    procedure CB_RAPClick(Sender: TObject);

  private
    fDtStart: TDateTime;
    fDtDebut: TDateTime;
    fDtFin: TDateTime;
    fStSqlWhere: string; // ajout du filtre retourné par le mul
    fStWhereEtatLg: string; // ajout du filtre sur l etat des lignes retourné par le mul
    fStAbsences: string; // indique si on gere les absences ou non
    fStDemiJournee: string; // indique si le planning est affiché en demi-journée
    fStArgument: string; //passer l'action de la fiche

    fInNumPlanning: Integer; // numero du planning sélectionné
    FPlanningJour: TPlanningGA;
    fStFontName: string;

    fCurrentPlanning: TPlanningGA;
    fTOBModelePlanning: TOB; // tob virtuelle des modeles

    function GetPlanningJour: TPlanningGA;
    function LoadTobParam(T: TOB): integer;
    function LoadModeles: boolean;
    procedure InitNonProductive;
    procedure LoadAffaireSel;
    function LoadArticlesSel : Boolean;
    function SelectArt(pInNumTob : Integer; pStNumTache : String) : Boolean;
    procedure SelectToutesTaches; // selection de toutes les taches
    procedure SelectTaches; // selection des taches selon le param RAP

  public
    property DtStart: TDateTime read fDtStart write fDtStart;
    property DtDebut: TDateTime read fDtDebut write fDtDebut;
    property DtFin: TDateTime read fDtFin write fDtFin;

    property PlanningJour: TPlanningGA read GetPlanningJour;

    property InNumPlanning: Integer read fInNumPlanning write fInNumPlanning;
    property StSqlWhere: string read fStSqlWhere write fStSqlWhere; // ajout du filtre retourné par le mul
    property StWhereEtatLg: string read fStWhereEtatLg write fStWhereEtatLg;
    property StAbsences: string read fStAbsences write fStAbsences;
    property StDemiJournee: string read fStDemiJournee write fStDemiJournee;
    property StArgument: string read fStArgument write fStArgument;
  end;

  TPlanningGA = class(TObject)
  private
    FComposantPlanning: THPlanning;
    FFichePlanning: TAF_Planning; // récupération du pointeur de la fiche qui contient le planning
    // pour réafficher le label la_sélection
    fTobsPlanningTobItems: TOB;
    fTobsPlanningTobEtats: TOB;
    fTobsPlanningTobRes: TOB;
    fTobsPlanningTobCols: TOB;
    fTobsPlanningTobRows: TOB;
    fTobsPlanningTobEvents: TOB;
    fTobAffAdm : Tob;

    fTobEtatDefaut: Tob; // etat par defaut defini en parametrage
    fTobEtatXml: Tob; // table etats
    fTOBModelePlanning: Tob;

    fDtDebut: TDateTime;
    fDtFin: TDateTime;
    fDtStart: TDateTime; // date de l'item sélectionné récupéré pour ce positionner dessus lors d'un zoom
    fInNumModele: Integer;
    fStTypeModele: string;
    fStNumPlanning: string;
    fNbEtat: Integer;
    fStWhere: string;
    fBoGestionAbs: Boolean;
    fBoDemiJournee: Boolean;
    fStArgument: string;
    fStAction: string;
    fStStatutDefaut : string; // statut par defaut des etats

    fColonne1: RecordColonne;
    fColonne2: RecordColonne;
    fColonne3: RecordColonne;

    fTobCouranteMere: Tob; // on utilise une tob mere pour conserve le bon nom de tob lors de la duplication
    fTobCourante: Tob; // fille de fTobCouranteMere
    fTobEtats: TOB; // chargement des etats popup

    fStAffaire: string;
    fStTiers: string;
    fStFonction: string;
    fStRes: string;
    fStTache: string;
    fStCodeArt: string;

    fTobSelection: Tob;
    fTobArtSel:Tob; // liste des articles de la tob sélectionnée
    fStAffaireSel: string;
    fStTiersSel: string;
    fStResSel: string;
    fStTacheSel: string;
    fStArticleSel: string;
    fStCodeArtSel: string;
    fStTypeArtSel: string;
    fStLibelleArtSel: string;
    fStIdentLigne: string;
    fStAffNonProductive: string; // Journée non productive
    fStlibNonProductive: string;

    // C.B 08/07/2005 valeurs de fInModeCreation
    // cInCreationDirecte, cInCreationSansSelection, cInCreationAvecSelection,
    // cInCreationNonProductive, cInCreationSelectionEtRes
    fInModeCreation: Integer;
    fAction: TActionFiche;
 
    fBoMonPlanning : Boolean; // permet de savoir si monplanning pour lancer l'dition avec un filtre

    //procedure ClearTob;
    procedure CreateTob;
    function ModeCreationDirecte(pBoLoadRessource: Boolean): string;
    procedure Display(pDtStart: TDateTime);

    procedure InitPlanning; virtual;
    procedure InitParamUtil;
    procedure DecodeCodeModele(pStCodeModele: string);
    procedure SelectPlanning; virtual; abstract;
    function LoadTobEtat(T: TOB): integer;
    procedure FormatTobRes; virtual; abstract;
    procedure LoadPopUpMenu(pBoEtatPourMonPlanning: Boolean);
    procedure ModifyItem(Sender: TObject; Item: TOB; var Cancel: boolean); virtual;
    function UpdateAfPlanning(ToItem : Tob; var pStRetour : String) : Boolean;
    function UpdateItem(FromItem, ToItem: TOB; vStAction: string): Boolean; virtual; abstract;
    procedure CopyItem(Sender: TObject; Item: TOB; var Cancel: boolean); virtual; abstract;
    procedure DeleteItem(Sender: TObject; Item: TOB; var Cancel: boolean);
    procedure InitItem(Sender: TObject; var Item: TOB; var Cancel: boolean);
    procedure CreateItem(Sender: TObject; Item: TOB; var Cancel: boolean); virtual;
    //function MoveItem(FromItem, ToItem : TOB; pBoBreak : Boolean) : Boolean;
    procedure MoveItem(Sender: TObject; Item: TOB; var Cancel: boolean);
    procedure DoubleClick(Sender: TObject);
    procedure AvertirApplication(Sender: TObject; FromItem: TOB; ToItem: TOB; Actions: THPlanningAction);
    procedure BeforeChange(const Item: TOB; const LaRessource: string; const LaDateDeDebut, LaDateDeFin: TDateTime; var Cancel: Boolean);
    procedure Link(Sender: TObject; TobSource, TobDestination: TOB; Option: THPlanningOptionLink; var Cancel: Boolean);

    function GetTiers(pStAffaire: string): string;
    function ChangeEtats(ZCode: Integer): string;
    procedure OnPopup(Item: tob; ZCode: integer; var Redraw: boolean);

    procedure LoadItemsPlanning(pStWhere: string); virtual;
    function MakeRequetePaie(pStWhere : String) : String;
    procedure AjoutTacheAbsence(pInIndex: Integer; pStRes, pStResLib: string); virtual; abstract;
    procedure AjoutNonProductive(pInIndex: Integer; pStRes, pStResLib: string); virtual; abstract;
    procedure LoadAbsencesPlanning(pStWhere: string); virtual; abstract;
    function LoadResPlanning(pStWhere: string): Boolean; virtual; abstract;
    procedure LoadAffNonProductive;
    function LoadEtatPlanning: Boolean;
    //procedure LoadAdresses;
    //procedure LoadTiers;
    procedure LoadArgument;

    procedure DecodeCle(pStCle: string); virtual; abstract;
    procedure EncodeCle(pStCle1, pStCle2, pStCle3: string; var pStCle: string); virtual;
    function ChangeLigneItemOk(Item: Tob): boolean; virtual; abstract;
    function CreateTacheRessource(Item: Tob; pBoForce: Boolean): boolean; virtual; abstract;
    procedure UpdateData(FromItem, ToItem: Tob);
    procedure DoubleClickSpec(ACol, ARow: INTEGER; TypeCellule: TPlanningTypeCellule; T: TOB = nil);
    procedure BParamClick; virtual; abstract;
    procedure BExcelClick(pStFileName: string; pBoOuvrir: Boolean);
    function CreationDirecte(pItem: Tob): string;
    procedure LoadLibelleRessource(pItem: Tob);
    procedure LoadComplementItem(pItem: Tob);
    procedure ClearAffaireSelect;

    function IsBloque(Item: Tob): Boolean;
    procedure ChangeOptions(pBoValeur: Boolean);
    procedure FormatChampLibelle(pItem: Tob);
    procedure RafraichirItem(Item: Tob; pStRetour, pStMode: string; pBoLoadTache, pBoLoadCodePostal : Boolean);
    procedure RafraichirItemPartiel(Item: Tob; pStRetour, pStMode: string);
    procedure RafraichirAssocies(Item: Tob);
    procedure IconeCEGID(i: Integer);
    procedure AffichageCommentaire(pItem: Tob);
    procedure AffichageQuantites(pItem: Tob);
    function FormatLibTache(pStRtt, pStValider : String) : String;
    procedure CreateAbsence(pItem, pTobAbs : Tob; pStRes, pStICode : String; pIn, pInNumTache, pInNumligne : Integer);
    //procedure CreateAbsence(pItem: Tob; pStRes, pStICode, pStDebJ, pStFinJ, pStAffaire, pStLibTache: string;
    //  pIn, pInNumTache, pInNumligne: Integer; pDtDeb, pDtFin: TDateTime);
    procedure LoadAbsencesPlanningRessource(pStWhere: string);
    function AffaireTermine(pStAffaire : String) : Boolean;
    function AfficherTiers : Boolean; virtual; abstract; // affichage du tiers dans l'item
    function AfficherRess  : Boolean; virtual; abstract; // affichage de la ressource dans l'item
    function ModeAffichage : String; virtual; abstract; // indique le contenu de l'item
                                                        // utilisé dans les item école
                                                        // et dans les items affaire
                                                        // ARTICLE
                                                        // AFFAIRE
                                                        // RESSOURCE

    function GetStatutItem(pStStatutTache : String) : String; // statut de l'item à la création

  public
    property DtFin: TDateTime read fDtFin;
    property DtDebut: TDateTime read fDtDebut;

    constructor Create(pFichePlanning: TAF_Planning);
    destructor Destroy; override;
    function Load(pDtDebut, pDtFin, pDtStart: TDateTime; pStWhere, pStCodeModele, pStWhereEtatLigne, pStAbscences, pStDemiJournee, pStArgument: string; pTOBModelePlanning: Tob):
      Boolean;
    function DisplayHint(pItem: Tob; pStHint: string = ''; pTobTache: Tob = nil): string;
  end;

  // ACCES PAR AFFAIRE / RESSOURCE
  // PLANNING DES INTERVENTIONS
  // PAS TIERS, PAS RESSOURCE DANS ITEM
  // MODE AFFICHAGE : ARTICLE
  TPlanningGA1 = class(TPlanningGA)
  private
    procedure SelectPlanning; override;
    procedure LoadItemsPlanning(pStWhere: string); override;
    function LoadResPlanning(pStWhere: string): Boolean; override;
    procedure AjoutTacheAbsence(pInIndex: Integer; pStRes, pStResLib: string); override;
    procedure AjoutNonProductive(pInIndex: Integer; pStRes, pStResLib: string); override;
    procedure LoadAbsencesPlanning(pStWhere: string); override;
    procedure FormatTobRes; override;
    procedure DecodeCle(pStCle: string); override;
    function ChangeLigneItemOk(Item: Tob): boolean; override;
    function CreateTacheRessource(Item: Tob; pBoForce: Boolean): boolean; override;
    function UpdateItem(FromItem, ToItem: TOB; vStAction: string): Boolean; override;
    procedure InitPlanning; override;
    procedure CreateItem(Sender: TObject; Item: TOB; var Cancel: boolean); override;
    procedure CopyItem(Sender: TObject; Item: TOB; var Cancel: boolean); override;
    procedure ModifyItem(Sender: TObject; Item: TOB; var Cancel: boolean); override;
    procedure BParamClick; override;
    function AfficherTiers : Boolean; override;
    function AfficherRess  : Boolean; override;
    function ModeAffichage : String; override;
  end;

  // ACCES PAR RESSOURCE / AFFAIRE
  // PLANNING DES INTERVENTIONS
  // PAS TIERS, PAS RESSOURCE DANS ITEM
  // MODE AFFICHAGE : ARTICLE
  TPlanningGA2 = class(TPlanningGA)
  private
    procedure SelectPlanning; override;
    procedure LoadItemsPlanning(pStWhere: string); override;
    function LoadResPlanning(pStWhere: string): Boolean; override;
    procedure AjoutTacheAbsence(pInIndex: Integer; pStRes, pStResLib: string); override;
    procedure AjoutNonProductive(pInIndex: Integer; pStRes, pStResLib: string); override;
    procedure LoadAbsencesPlanning(pStWhere: string); override;
    procedure FormatTobRes; override;
    procedure DecodeCle(pStCle: string); override;
    function ChangeLigneItemOk(Item: Tob): boolean; override;
    function CreateTacheRessource(Item: Tob; pBoForce: Boolean): boolean; override;
    function UpdateItem(FromItem, ToItem: TOB; vStAction: string): Boolean; override;
    procedure InitPlanning; override;
    procedure CreateItem(Sender: TObject; Item: TOB; var Cancel: boolean); override;
    procedure CopyItem(Sender: TObject; Item: TOB; var Cancel: boolean); override;
    procedure ModifyItem(Sender: TObject; Item: TOB; var Cancel: boolean); override;
    procedure BParamClick; override;
    function AfficherTiers : Boolean; override;
    function AfficherRess  : Boolean; override;
    function ModeAffichage : String; override;
  end;

  // ACCES PAR AFFAIRE / TACHE
  // PLANNING DES RESSOURCES
  // PAS TIERS, RESSOURCE DANS ITEM
  // MODE AFFICHAGE RESSOURCE
  TPlanningGA3 = class(TPlanningGA)
  private
    procedure SelectPlanning; override;
    procedure LoadItemsPlanning(pStWhere: string); override;
    function LoadResPlanning(pStWhere: string): Boolean; override;
    procedure AjoutTacheAbsence(pInIndex: Integer; pStRes, pStResLib: string); override;
    procedure AjoutNonProductive(pInIndex: Integer; pStRes, pStResLib: string); override;
    procedure LoadAbsencesPlanning(pStWhere: string); override;
    procedure FormatTobRes; override;
    procedure DecodeCle(pStCle: string); override;
    function ChangeLigneItemOk(Item: Tob): boolean; override;
    function CreateTacheRessource(Item: Tob; pBoForce: Boolean): boolean; override;
    function UpdateItem(FromItem, ToItem: TOB; vStAction: string): Boolean; override;
    procedure InitPlanning; override;
    procedure CreateItem(Sender: TObject; Item: TOB; var Cancel: boolean); override;
    procedure CopyItem(Sender: TObject; Item: TOB; var Cancel: boolean); override;
    procedure ModifyItem(Sender: TObject; Item: TOB; var Cancel: boolean); override;
    procedure BParamClick; override;
    function AfficherTiers : Boolean; override;
    function AfficherRess  : Boolean; override;
    function ModeAffichage : String; override;
  end;

  // ACCES PAR RESSOURCE
  // PLANNING DES AFFAIRES
  // PAS TIERS, PAS RESSOURCE DANS ITEM
  // MODE AFFICHAGE AFFAIRE
  TPlanningGA4 = class(TPlanningGA)
  private
    procedure SelectPlanning; override;
    procedure LoadItemsPlanning(pStWhere: string); override;
    function LoadResPlanning(pStWhere: string): Boolean; override;
    procedure AjoutTacheAbsence(pInIndex: Integer; pStRes, pStResLib: string); override;
    procedure AjoutNonProductive(pInIndex: Integer; pStRes, pStResLib: string); override;
    procedure LoadAbsencesPlanning(pStWhere: string); override;
    procedure FormatTobRes; override;
    procedure DecodeCle(pStCle: string); override;
    function ChangeLigneItemOk(Item: Tob): boolean; override;
    function CreateTacheRessource(Item: Tob; pBoForce: Boolean): boolean; override;
    function UpdateItem(FromItem, ToItem: TOB; vStAction: string): Boolean; override;
    procedure InitPlanning; override;
    procedure CreateItem(Sender: TObject; Item: TOB; var Cancel: boolean); override;
    procedure CopyItem(Sender: TObject; Item: TOB; var Cancel: boolean); override;
    procedure ModifyItem(Sender: TObject; Item: TOB; var Cancel: boolean); override;
    procedure BParamClick; override;
    function AfficherTiers : Boolean; override;
    function AfficherRess  : Boolean; override;
    function ModeAffichage : String; override;
  end;

  // ACCES PAR RESSOURCE
  // PLANNING DES INTERVENTIONS
  // TIERS, PAS RESSOURCE DANS ITEM
  // MODE AFFICHAGE ARTICLE
  TPlanningGA5 = class(TPlanningGA)
  private
    procedure SelectPlanning; override;
    procedure LoadItemsPlanning(pStWhere: string); override;
    function LoadResPlanning(pStWhere: string): Boolean; override;
    procedure LoadAbsencesPlanning(pStWhere: string); override;
    procedure AjoutTacheAbsence(pInIndex: Integer; pStRes, pStResLib: string); override;
    procedure AjoutNonProductive(pInIndex: Integer; pStRes, pStResLib: string); override;
    procedure FormatTobRes; override;
    procedure DecodeCle(pStCle: string); override;
    function ChangeLigneItemOk(Item: Tob): boolean; override;
    function CreateTacheRessource(Item: Tob; pBoForce: Boolean): boolean; override;
    function UpdateItem(FromItem, ToItem: TOB; vStAction: string): Boolean; override;
    procedure InitPlanning; override;
    procedure CreateItem(Sender: TObject; Item: TOB; var Cancel: boolean); override;
    procedure CopyItem(Sender: TObject; Item: TOB; var Cancel: boolean); override;
    procedure ModifyItem(Sender: TObject; Item: TOB; var Cancel: boolean); override;
    procedure BParamClick; override;
    function AfficherTiers : Boolean; override;
    function AfficherRess  : Boolean; override;
    function ModeAffichage : String; override;
  end;

  // ACCES PAR AFFAIRE
  // PLANNING DES INTERVENTIONS
  // PAS TIERS, RESSOURCE DANS ITEM
  // MODE AFFICHAGE ARTICLE
  TPlanningGA6 = class(TPlanningGA)
  private
    procedure SelectPlanning; override;
    procedure LoadItemsPlanning(pStWhere: string); override;
    function LoadResPlanning(pStWhere: string): Boolean; override;
    procedure AjoutTacheAbsence(pInIndex: Integer; pStRes, pStResLib: string); override;
    procedure AjoutNonProductive(pInIndex: Integer; pStRes, pStResLib: string); override;
    procedure LoadAbsencesPlanning(pStWhere: string); override;
    procedure FormatTobRes; override;
    procedure DecodeCle(pStCle: string); override;
    function ChangeLigneItemOk(Item: Tob): boolean; override;
    function CreateTacheRessource(Item: Tob; pBoForce: Boolean): boolean; override;
    function UpdateItem(FromItem, ToItem: TOB; vStAction: string): Boolean; override;
    procedure InitPlanning; override;
    procedure CreateItem(Sender: TObject; Item: TOB; var Cancel: boolean); override;
    procedure CopyItem(Sender: TObject; Item: TOB; var Cancel: boolean); override;
    procedure ModifyItem(Sender: TObject; Item: TOB; var Cancel: boolean); override;
    procedure BParamClick; override;
    function AfficherTiers : Boolean; override;
    function AfficherRess  : Boolean; override;
    function ModeAffichage : String; override;
  end;

  // ACCES PAR RESSOURCE
  // MON PLANNING
  // TIERS, PAS RESSOURCE DANS ITEM 
  // MODE AFFICHAGE ARTICLE
  TPlanningGA7 = class(TPlanningGA)
  private
    procedure SelectPlanning; override;
    procedure LoadItemsPlanning(pStWhere: string); override;
    function LoadResPlanning(pStWhere: string): Boolean; override;
    procedure LoadAbsencesPlanning(pStWhere: string); override;
    procedure AjoutTacheAbsence(pInIndex: Integer; pStRes, pStResLib: string); override;
    procedure AjoutNonProductive(pInIndex: Integer; pStRes, pStResLib: string); override;
    procedure FormatTobRes; override;
    procedure DecodeCle(pStCle: string); override;
    function ChangeLigneItemOk(Item: Tob): boolean; override;
    function CreateTacheRessource(Item: Tob; pBoForce: Boolean): boolean; override;
    function UpdateItem(FromItem, ToItem: TOB; vStAction: string): Boolean; override;
    procedure InitPlanning; override;
    procedure CreateItem(Sender: TObject; Item: TOB; var Cancel: boolean); override;
    procedure CopyItem(Sender: TObject; Item: TOB; var Cancel: boolean); override;
    procedure ModifyItem(Sender: TObject; Item: TOB; var Cancel: boolean); override;
    procedure BParamClick; override;
    function AfficherTiers : Boolean; override;
    function AfficherRess  : Boolean; override;
    function ModeAffichage : String; override;
  end;

procedure ExecPlanning(pStNumPlanning, pStDateDebut, pStDateFin, pStDateStart, pStSqlWhere, pStWhereEtatLigne, pStAbsences, pStDemiJournee, pStArgument: string);

implementation

// M3FP, HColor, affaireUtil, UTof, AffEcheanceUtil, Buttons, Dialogs,
uses //UTOFAFTACHE_MUL,
     UTomAFPlanning,
     uTofBTTaches,
     Utofaffaire_mul,
     utilmenuaff,
     //UTofAFLIGPLANNING_Mul,
     HMsgBox,
     HStatus,
     //UtilPLanning,
     paramsoc,
     dicobtp,
     AglInit,
     Paramdat,
     UtilTaches,
     UtilRessource,
     UAFO_Ressource,
     ConfidentAffaire,
     EntGC,
     CalcOleGenericAff,
     //uTofAFPlanningParam,
     rtfCounter,
     //uTofAfExpMsExcel,
     GereTobInterne,
     //UtofAfListeTache_Pla,
     AffaireUtil,
     lookup;
     //uTofAfPlanningArt,
     //UtilPlanningEcran;

{$R *.DFM}

const
  // libellés des messages  
  TexteMessage: array[1..81] of string = (
    {1}'Erreur lors de la mise à jour de l''intervention.',
    {2}'L''article doit être renseigné.',
    {3}'L''état doit être renseigné.',
    {4}'L''article doit être renseigné.',
    {5}'On ne peut pas changer l''affaire d''une intervention.',
    {6}'Erreur au chargement de l''affaire.',
    {7}'Ce code ressource n''est pas prévu pour cette tâche, voulez l''ajouter à la tâche ?', //mcd 14/03/05 mis au masculin pour Ok GI
    {8}'Vous ne pouvez pas déplacer une instance d''une tâche vers une autre tâche.',
    {9}'On ne peut pas déplacer une ressource sur une autre tâche.',
    {10}'Pour les plannings par ressource, il faut affecter les ressources pour pouvoir afficher le planning.',
    {11}'L''état par défaut défini dans les paramètres n''existe pas !',
    {12}'Aucune action n''est possible sur une absence !',
    {13}'Vous allez planifier plus de jours que prévu. #13#10 Voulez-vous continuer ?',
    {14}'Voulez-vous vraiment planifier plus de jours que prévu ?',
    {15}'Aucune intervention sélectionnée.',
    {16}'Le paramètre société "Statut par défaut" du planning doit être défini pour pouvoir utiliser le planning.',
    {17}'Voulez-vous créer une intervention pour une autre affaire que celle de la ligne à planifier sélectionnée ?',
    {18}'Cette intervention est bloquée par %s.',
    {19}'Voulez-vous créer une intervention pour une autre ressource que celle de la ligne à planifier sélectionnée ?',
    {20}'Attention, Les samedis et dimanches ne sont pas visibles, mais ils sont planifiés.',
    {21}'Vous devez sélectionner ou créer une affaire avec des tâches non facturables.#13#10 Voulez-vous continuer ?',
    {22}'Vous devez vous positionner sur une ligne NON FACTURABLES.',
    {23}'La ressource %s ne fait pas partie de votre équipe.',
    {24}'La ressource ne fait pas partie de votre équipe.',
    {25}'Pour les plannings par tâches, il faut créer les tâches pour pouvoir afficher le planning.',
    {26}'Revoir les paramètres des plannings.',
    {27}'Voulez-vous consulter le fichier excel qui va être généré ?',
    {28}'Voulez-vous également déplacer tout ce qui est prévu pour cette intervention ?',
    {29}'Aucun modèle de planning n''a été défini',
    {30}'Attention !',
    {31}'Erreur de données lors du chargement de la tâche.',
    {32}'Suppression interdite.',
    {33}'L''état %s n''est pas paramétré.',
    {34}'Il y a d''autres ressources planifiées pour cette intervention, #13#10 pensez à également les déplacer.',
    {35}'Aucune affaire ne correspond la ressource sélectionnée.',
    {36}'L''état de cette intervention ne permet aucune modification.',
    {37}'Planning par affaires',
    {38}'Planning des interventions par ressources',
    {39}'Planning par affaires par articles',
    {40}'Planning par affaires par ressources',
    {41}'Planning des affaires',
    {42}'Tâches',
    {43}'Lignes d''articles',
    {44}'Commande/Affaire école',
    {45}'La date que vous sélectionnée n''est pas dans l''interval actuellement affiché.',
    {46}'Journée non facturable : %s',
    {47}'Ressource : %s  %s     ',
    {48}'Nombre d''intervenants inscrits : %s',
    {49}'Article',
    {50}'Libellé',
    {51}'Ressource',
    {52}'Client',
    {53}'Dernière modif',
    {54}'Affaire',
    {55}' Statut : %s',
    {56}'Affaire non facturable',
    {57}'Le concept saisie planning manager vous autorise à faire cette modification.#13#10Voulez vous vraiment la faire ?',
    {58}'Voulez-vous vraiment supprimer cette intervention ?',
    {59}'Planning',
    {60}'Tâche',
    {61}'Article',
    {62}'Intervention',
    {63}'NON FACTURABLES',
    {64}'Nombre d''intervenants inscrits : %s',
    {65}'Affaire : %s      Tiers :  %s     Article :  %s  %s     ',
    {66}'ABSENCE',
    {67}'Tiers/Affaire',
    {68}'Vous ne pouvez plus planifier d''intervention pour une affaire terminée.',
    {69}'Vous ne pouvez plus planifier d''intervention pour une tâche terminée.',
    {70}'Vous ne pouvez pas créer, déplacer, réduire ou étirer une intervention à cheval sur 2 semaines.',
    {71}'Vous ne pouvez pas créer, déplacer, réduire ou étirer une intervention à cheval sur 2 mois.',
    {72} 'ABS VALIDEE', //43 // attention, a remplacer également dans utilplanning
    {73} 'RTT', //44
    {74} 'RTT VALIDE', //45
    {75} 'ABSENCES',
    {76} 'Vous ne pouvez pas planifier des tâches en heures dans ce planning, l''unité minimum est la demi-journée.',
    {77} 'Cette intervention est dans un statut qui n''est pas modifiable.',
    {78} 'Toutes les taches sont planifiées. Voulez-vous les voir ?',
    {79} 'Toutes les lignes sont planifiées. Voulez-vous les voir ?',
    {80} 'Cette affaire ne contient pas de tâches planifiables.',
    {81} 'Cette affaire ne contient pas de lignes planifiables.'
    );

{***********A.G.L.***********************************************
Auteur  ...... : C. BOUET
Créé le ...... : 22/02/2005
Modifié le ... :   /  /
Description .. : Point d'entrée de l'unité
Mots clefs ... :
*****************************************************************}
procedure ExecPlanning(pStNumPlanning, pStDateDebut, pStDateFin, pStDateStart, pStSqlWhere, pStWhereEtatLigne, pStAbsences, pStDemiJournee, pStArgument: string);
begin
  with TAF_Planning.Create(Application) do
  try
    //iDate1900
    if pStDateStart = '2' then
      DtStart := iDate1900
    else
      DtStart := StrToDate(pStDateStart);

    if pStDateDebut = dateToStr(iDate1900) then
    begin

      // C.B 28/05/2005
      // changement de paramsoc pour le chargement
      DtDebut := PlusDate(now, -getParamSocSecur('SO_AFPLANAFFICHS1', '00:00:00'), 'S');
      DtFin := PlusDate(now, getParamSocSecur('SO_AFPLANAFFICHS2', '00:00:00'), 'S');

      {if (getParamSocSecur('SO_AFPLANAFFICH') = 'PER') then
      begin
        DtDebut := DebutDeMois(PlusMois(now, getParamSocSecur('SO_AFPLANAFFICHM1') * (-1)));
        DtFin := FinDeMois(PlusMois(now, getParamSocSecur('SO_AFPLANAFFICHM2')));
      end
      else
      begin
        DtDebut := DebutAnnee(now);
        DtFin := FinAnnee(now);
      end;
      }
    end
    else
    begin
      DtDebut := StrToDate(pStDateDebut);
      DtFin := StrToDate(pStDateFin);
    end;

    if pStSqlWhere <> '' then
      StSqlWhere := ' AND ' + pStSqlWhere;

    InNumPlanning := ValeurI(pStNumPlanning) - cInMenuPlanning;

    if pStWhereEtatLigne <> '' then
      StWhereEtatLg := pStWhereEtatLigne;
    StAbsences := pStAbsences;
    StDemiJournee := pStDemiJournee;
    StArgument := pStArgument;

    case InNumPlanning of
      6:
        begin
          Caption := TraduitGA(TexteMessage[37]); //Planning par affaires
          POPUPGENERAL.Items[3].Caption := TraduitGA(TexteMessage[38]); //Planning des interventions par ressources
        end;
      1:
        begin
          Caption := TexteMessage[40]; //Planning par affaires par ressources
          POPUPGENERAL.Items[3].Caption := TraduitGA(TexteMessage[38]); //Planning des interventions par ressources
        end;
      2, 5, 7:
        begin
          Caption := TraduitGA(TexteMessage[38]); //Planning des interventions par ressources
          POPUPGENERAL.Items[3].Caption := TraduitGA(TexteMessage[39]); //Planning par affaires par articles
        end;
      3:
        begin
          Caption := TraduitGA(TexteMessage[39]); //Planning par affaires par articles
          POPUPGENERAL.Items[3].Caption := TraduitGA(TexteMessage[38]); //Planning des interventions par ressources
        end;
      4:
        begin
          Caption := TraduitGA(TexteMessage[41]); //Planning des affaires
          POPUPGENERAL.Items[3].Destroy;
        end;
    end;

    // C.B 04/05/05
    // test de l'existance des ressources avant l'ouverture du planning pour les plannings par ressource
    // soit les plannings 1,2,4
    if (InNumPlanning = 1) or (InNumPlanning = 2) or (InNumPlanning = 4) then
    begin
      if ExisteRessources(pStSqlWhere) then
        ShowModal
      else
        PGIBoxAF(TexteMessage[10], TexteMessage[59]);
    end

      // C.B 26/09/2005
      // on le teste meme si les taches sont obligatoires
      // test de l'existence des taches pour le planning 3 si la créeation n'est pas obligatoire
    else
      if (InNumPlanning = cInPlanning3) then //and (getParamSocSecur('SO_AFMODEPLANNING') <> 'ART') then
    begin
      if existeTache(pStSqlWhere) then
        ShowModal
      else
        PGIBoxAF(TexteMessage[25], TexteMessage[59]); //Pour les plannings par tâches, il faut créer les tâches pour pouvoir afficher le planning.',
    end
    else
      ShowModal;
  finally
    Free;
  end;
end;

/////////////////////////////////////////
//             THintWindowPlanning
/////////////////////////////////////////
constructor THintWindowPlanning.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  Color := $80FFFF;
  Canvas.Font := Screen.HintFont;
  Canvas.Font.Name := 'Arial';
  Canvas.Font.Size := 8;
  Canvas.Brush.Style := bsClear;
end;

{***********A.G.L.***********************************************
Auteur  ...... : C. BOUET
Créé le ...... : 22/02/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TAF_Planning.FormCreate(Sender: TObject);
begin
  application.HintHidePause := 15000;
  fStFontName := screen.HintFont.name;
  HintWindowClass := THintWindowPlanning;
  //
  {if ModeAutomatique then
  begin
    BT_LIGNES.Hint := TexteMessage[43];
  end
  else
  begin
    BT_LIGNES.Hint := TexteMessage[42];
  end;}
  
end;

procedure TAF_Planning.FormShow(Sender: TObject);
begin
  La_Selection.caption := '';

  // C.B 08/09/2005
  // affichage du commentaire
  La_Commentaire.caption := '';
  La_Quantite.caption := '';
 
  // C.B 15/05/2006
  // libellé de la base connectée dans le planning
  if VH_GC.GCIfDefCEGID then            
		LA_BASE.caption := v_PGI.NomSociete
  else
  	LA_BASE.caption := '';

  // C.B 09/08/2005
  // on renomme ce libellé si GC + GA
  //if (CtxGcAFF in V_PGI.PGIContexte) and getParamSocSecur('SO_AFECOLE') then
  // 20/10/2005 changement que pour CEGID
  if VH_GC.GCIfDefCEGID then
    MnAffaire.Caption := TraduitGA(TexteMessage[44]); //Commande/Affaire école

  // on met le zoom tache enabled à false
  // pour ne pas réouvrir une seconde fois le planning d'origine
  if pos('ZOOMTACHE', fStArgument) > 0 then
    POPUPGENERAL.Items[2].Enabled := False;

  // suppression du zoom pour ne pas réouvrir une seconde fois le planning d'origine
  if pos('ZOOMPLANNING', fStArgument) > 0 then
    POPUPGENERAL.Items[3].Enabled := False;

  // C.B 05/09/2005
  // suppression des affaires ecoles
  if ctxAffaire in V_PGI.PGIContexte then
  begin
    if fInNumPlanning = 4 then
      // planning des affaires, pas de zoom sur autre planning
			POPUPGENERAL.Items[5].Destroy
    else                            
      POPUPGENERAL.Items[6].Destroy; // affaire ecole
  end;
   
  //if assigned(fTobModelePlanning) then fTobModelePlanning.cleardetail;
  if not LoadModeles then
  begin
    HShowMessage('', texteMessage[30], TexteMessage[29]);
  end
  else
  begin
    InitNonProductive;
    //Chargement et affichage des données
    if PlanningJour.load(fDtDebut, fDtFin, fDtStart, fStSqlWhere, 'P' + IntToStr(GetNumModele) + 'J', fStWhereEtatLg, fStAbsences, fStDemiJournee, fStArgument, fTOBModelePlanning) then
      fCurrentPlanning := PlanningJour;
  end;

  // seul le superviseur peut modifier les paramètres dépuis le planning
  BPARAM.visible := v_pgi.superviseur;

  // formatage des champs de la fiche
  // utilisé avec les option type.action = create + false, donne accès au zone en création
  //                         type.action = consult + false, donne accès au zone en consultation
  if (fCurrentPlanning.fAction = TaConsult) then
  begin
    ChargeCleAffaire(THEDIT(ED_AFF0), THEDIT(ED_AFF1),
                     THEDIT(ED_AFF2), THEDIT(ED_AFF3),
                     THEDIT(ED_AVE),
                     Nil, TaConsult, ED_AFF.text, False);
    PA_Quick.Visible := false; 
  end
  else
    ChargeCleAffaire(THEDIT(ED_AFF0), THEDIT(ED_AFF1),
                     THEDIT(ED_AFF2), THEDIT(ED_AFF3),
                     THEDIT(ED_AVE),          
                     Nil, TaCreat, ED_AFF.text, False);
                                
end;                                   

{***********A.G.L.***********************************************
Auteur  ...... : C. BOUET
Créé le ...... : 22/02/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TAF_Planning.GetNumModele: Integer;
begin
  result := 1;
  case fInNumPlanning of
    1, 6: result := 1;
    2, 5, 7: result := 2;
    3: result := 3;
    4: result := 4;
  end;
end;

procedure TAF_Planning.InitNonProductive;
begin
  case fInNumPlanning of
    2, 5, 7:
      begin
        BTNonProductive.Visible := true;
        BT_LIGNES.GroupIndex := BTNonProductive.GroupIndex;
      end;

    4:
      begin
        BTNonProductive.Visible := false;
        BT_LIGNES.Visible := false;
        BT_EFFACELIGNES.Visible := false;
      end;
  else
    begin
      BTNonProductive.Visible := false;
      BT_LIGNES.GroupIndex := 0;
    end;
  end;
end;

procedure TAF_Planning.LoadAffaireSel;
var
  vBoNewSel : Boolean;

begin
  vBoNewSel := False;

  // selectionAffaire
  {if RechercheAffairePlanning(ED_AFF, ED_AFF0, ED_AFF1, ED_AFF2, ED_AFF3, ED_AVE, ED_TIERS) then
      begin
      fCurrentPlanning.fStAffaireSel := ED_AFF.text;
      fCurrentPlanning.fStTiersSel := ED_TIERS.text;
      vBoNewSel := True;
      end
   else if SelectionAffairePlanning(ED_TIERS, ED_AFF, ED_AFF0, ED_AFF1, ED_AFF2,
                                   ED_AFF3, ED_AVE, False) then
      begin
      fCurrentPlanning.fStAffaireSel :=  ED_AFF.text;
      fCurrentPlanning.fStTiersSel := ED_TIERS.text;
      vBoNewSel := True;
      end;}

  if vBoNewSel then
    SelectTaches

  else if fCurrentPlanning.fStAffaireSel = '' then
    BT_EFFACELIGNESClick(self);
end;                                

// selection des taches en fonction de rap ou non
procedure TAF_Planning.SelectTaches;
var
  vSt : String;

begin
  if LoadArticlesSel then
      // au chargement on se positionne sur le premier article
      SelectArt(0,'')

    // si toutes les taches ont été recherchées et qu'on a rien
    else if (CB_RAP.state <> cbChecked) then
    begin
      {if ModeAutomatique then
        vSt := TexteMessage[81]
      else
        vSt := TexteMessage[80];}
      PGIBoxAF(vSt, TexteMessage[59]);
    end

    // si on n'avait pas déjà chargé toutes les tâches
    else
      SelectToutesTaches;
end;

function TAF_Planning.LoadArticlesSel : boolean;
var
  vSt     : String;

begin

  if fCurrentPlanning <> nil then
  begin
    vSt := 'SELECT ATA_NUMEROTACHE, ATA_CODEARTICLE, ATA_LIBELLETACHE1, ';
    vSt := vSt + 'ATA_ARTICLE, ATA_TYPEARTICLE, ATA_QTEINITPLA, ATA_QTEPLANIFPLA, ATA_QTEINITPLA - ATA_QTEPLANIFPLA AS QTERESTE, ATA_QTEFACTPLA FROM TACHE';
    vSt := vSt + ' WHERE ATA_AFFAIRE="' + fCurrentPlanning.fStAffaireSel + '"';
    if (CB_RAP.state = cbChecked) then
      vSt := vSt + ' AND ATA_QTEPLANIFPLA < ATA_QTEINITPLA ';

    vSt := vSt + ' ORDER BY ATA_CODEARTICLE';

    DetruitMaTobInterne('SELECTART');
    fCurrentPlanning.fTobArtSel := TOB.create ('#TACHE', MaTobInterne ('SELECTART'), -1);
    fCurrentPlanning.fTobArtSel.LoadDetailFromSQL(vSt);
    result := fCurrentPlanning.fTobArtSel.detail.count > 0;
  end
  else
    result := false;   
end;

function TAF_Planning.SelectArt(pInNumTob : Integer; pStNumTache : String) : Boolean;
var
  vTob : Tob;
begin
  result := True;
  if (pInNumTob = -1) then
  begin
    if (fCurrentPlanning.fStTache <> '') then
      vTob := fCurrentPlanning.fTobArtSel.FindFirst(['ATA_NUMEROTACHE'], [StrToInt(pStNumTache)], False)
    else
      vTob := fCurrentPlanning.fTobArtSel.detail[0];
  end
  else
    vTob := fCurrentPlanning.fTobArtSel.detail[pInNumTob];

  if vTob = nil then
    result := false
  else
  begin
    ED_TACHE.Text := vTob.GetString('ATA_CODEARTICLE');
    fCurrentPlanning.fStArticleSel := vTob.GetString('ATA_ARTICLE');
    fCurrentPlanning.fStCodeArtSel := vTob.GetString('ATA_CODEARTICLE');
    fCurrentPlanning.fStTacheSel   := vTob.GetString('ATA_NUMEROTACHE');
    fCurrentPlanning.fStTypeArtSel := vTob.GetString('ATA_TYPEARTICLE');
    fCurrentPlanning.fStLibelleArtSel := vTob.GetString('ATA_LIBELLETACHE1');
    fCurrentPlanning.fStResSel     := '';
    La_Selection.Caption := fCurrentPlanning.ModeCreationDirecte(True);
    ED_AVE.OnExit := nil;
    if BCHERCHEART.canFocus then
      BCHERCHEART.SetFocus;
    ED_AVE.OnExit := ED_AVEExit;
  end;
end;

procedure TAF_Planning.SelectToutesTaches;
var
  vSt : String;

begin
  {if ModeAutomatique then vSt := TexteMessage[79] else vSt := TexteMessage[78];}
  
  begin
    if PGIAskAF(vSt, TexteMessage[59]) = MrYes then
    begin
      // on décoche rap et on sélectionne
      CB_RAP.checked := false;
      if LoadArticlesSel then
        SelectArt(-1, fCurrentPlanning.fStTache)
    end
    else
      BT_EFFACELIGNESClick(self);
  end;
end;

procedure TAF_Planning.BannulerClick(Sender: TObject);
begin
  Close;
end;

procedure TAF_Planning.BPrevClick(Sender: TObject);
begin
  case fCurrentPlanning.FComposantPlanning.CumulInterval of
    pciSemaine: fCurrentPlanning.Display(PlusDate(fCurrentPlanning.FComposantPlanning.DateOfStart, -1, 'S'));
    pciMois: fCurrentPlanning.Display(PlusDate(fCurrentPlanning.FComposantPlanning.DateOfStart, -1, 'M'));
    pciTrimestre: fCurrentPlanning.Display(PlusDate(fCurrentPlanning.FComposantPlanning.DateOfStart, -3, 'M'));
  end;
end;

procedure TAF_Planning.BNextClick(Sender: TObject);
begin
  case fCurrentPlanning.FComposantPlanning.CumulInterval of
    pciSemaine: fCurrentPlanning.Display(PlusDate(fCurrentPlanning.FComposantPlanning.DateOfStart, 1, 'S'));
    pciMois: fCurrentPlanning.Display(PlusDate(fCurrentPlanning.FComposantPlanning.DateOfStart, 1, 'M'));
    pciTrimestre: fCurrentPlanning.Display(PlusDate(fCurrentPlanning.FComposantPlanning.DateOfStart, 3, 'M'));
  end;
end;

procedure TAF_Planning.BLastClick(Sender: TObject);
begin
  fCurrentPlanning.Display(PlusDate(fCurrentPlanning.DtFin, -1, 'S'));
end;

procedure TAF_Planning.BFirstClick(Sender: TObject);
begin
  fCurrentPlanning.Display(fCurrentPlanning.DtDebut);
end;

procedure TAF_Planning.BCalendrierClick(Sender: TObject);
var
  Key: Char;
  vDtDateEnCours: TDateTime;
begin
  DateEdit.enabled := True;
  vDtDateEnCours := fCurrentPlanning.FComposantPlanning.DateOfStart;
  DateEdit.text := DateTimeToStr(vDtDateEnCours);
  Key := '*';
  PARAMDATE(Self, DateEdit, Key);
  fDtStart := StrToDateTime(DateEdit.text);

  // pas dans la sélection
  if (fCurrentPlanning.FComposantPlanning.dateofstart < fCurrentPlanning.FComposantPlanning.IntervalDebut) or
     (fCurrentPlanning.FComposantPlanning.dateofstart > fCurrentPlanning.FComposantPlanning.IntervalFin) then
    PgiboxAf(texteMessage[45], TexteMessage[59])

  // on se positionne sur la bonne date
  else if (vDtDateEnCours <> fDtStart) then
    fCurrentPlanning.Display(fDtStart);
  DateEdit.enabled := False;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 16/04/2002
Modifié le ... :
Description .. : chargement des modeles de planning
Suite ........ :
Mots clefs ... :
*****************************************************************}
function TAF_Planning.LoadModeles: boolean;
var
  Q: TQuery;
  vListe: TStringList;
  vStream: TStream;
  vTOBModelePLanningXML: TOB; // table des modeles (blob en xml)

begin

  result := true;

  //creation du modele
  if not assigned(fTOBModelePlanning) then
    fTOBModelePlanning := TOB.Create('les modeles', nil, -1);

  vTOBModelePLanningXML := TOB.Create('#AFPLANNINGPARAM', nil, -1);
  vListe := TStringList.Create;

  // on charge TOUS les champs
  Q := nil;
  try
    Q := OpenSQL('Select APP_PARAMS from AFPLANNINGPARAM WHERE APP_TYPEAFPARAM="PLA"', True,-1,'',true);
    if not Q.Eof then
      vTOBModelePLanningXML.LoadDetailDB('AFPLANNINGPARAM', '', '', Q, False, True);

    // chargement de la liste
    vListe.Text := vTOBModelePLanningXML.Detail[GetNumModele - 1].GetValue('APP_PARAMS');

    // transfert dans une stream
    vStream := TStringStream.Create(vListe.Text);

    // recuperation dans une tob virtuelle fTOBModelePlanning
    TOBLoadFromXMLStream(vStream, LoadTobParam);
    vStream.Free;

    // TobsPlanningTobEtats.SaveToXmlFile('c:\toto.xml',false,true);
  finally
    Ferme(Q);
    vListe.Free;
    vTOBModelePLanningXML.Free;
  end;
end;

function TAF_Planning.LoadTobParam(T: TOB): integer;
var
  NewTob: Tob;
begin
  NewTOB := TOB.Create('param', fTOBModelePlanning, -1);
  try
    // récuperer les données xml dans newTob
    NewTob.Dupliquer(T.detail[0], True, True);
    result := 0;
  finally
    FreeAndNil(T);
  end;
end;

procedure TAF_Planning.BprintClick(Sender: TObject);
begin
  fCurrentPlanning.FComposantPlanning.TypeEtat := 'E';
  fCurrentPlanning.FComposantPlanning.NatureEtat := 'APL';
  fCurrentPlanning.FComposantPlanning.CodeEtat := 'APL';
  fCurrentPlanning.FComposantPlanning.Print(caption);
end;

procedure TAF_Planning.BParamClick(Sender: TObject);
begin

  fCurrentPlanning.BParamClick;
  if assigned(fTobModelePlanning) then
    fTobModelePlanning.cleardetail;
  LoadModeles;
  fCurrentPlanning.InitParamUtil;

end;

procedure TAF_Planning.BExcelClick(Sender: TObject);
var
  vStRetour: string;
  vStFileName: string;
  Tmp: string;
  Champ: string;
  Valeur: string;

begin

  {vStRetour := AFLanceFicheAFExpMsExcel('DATEDEB:' + DateToStr(fCurrentPlanning.fDtDebut) + ';DATEFIN:' + DateToStr(fCurrentPlanning.fDtFin));}

  if vStRetour <> '' then
  begin
    Tmp := (Trim(ReadTokenSt(vStRetour)));
    while (Tmp <> '') do
    begin
      DecodeArgument(Tmp, Champ, valeur);
      if Champ = 'FILE' then
        vStFileName := valeur

      else
        if Champ = 'DATEDEB' then
        fCurrentPlanning.FComposantPlanning.WorkAreaStart := strtodate(valeur)

      else
        if Champ = 'DATEFIN' then
        fCurrentPlanning.FComposantPlanning.WorkAreaEnd := strtodate(valeur);

      Tmp := (Trim(ReadTokenSt(vStRetour)));
    end;

    if (PGIAskAF(TexteMessage[27], '') = mrYes) then
      fCurrentPlanning.BExcelClick(vStFileName, True)
    else
      fCurrentPlanning.BExcelClick(vStFileName, false);
  end;
end;

// Saisie des intervenants de la tâche
procedure TAF_Planning.BINTERVENANTClick(Sender: TObject);
var
  vItem: TOB;
begin
  if (fCurrentPlanning.fStAffaire <> '') and (fCurrentPlanning.fStTache <> '') then
  begin
    vItem := fCurrentPlanning.FComposantPlanning.GetCurItem;

    if vItem <> nil then
    begin
      if getParamSocSecur('SO_AFECOLE', False) and (vItem.GetString('ATA_TYPEPLANIF') = 'ECO') then
      begin
        MnAffaireClick(nil); //AB-200511- Inscription sur la commande école
        Exit;
      end;
      SaisieIntervenantsTache(fCurrentPlanning.fStAffaire, fCurrentPlanning.fStTache, vItem.GetString('ATA_IDENTLIGNE'), vItem.GetString('APL_LIBELLEPLA'), ActionToString(fCurrentPlanning.fAction))
    end
    else
      PGIBoxAF(TexteMessage[15], TexteMessage[59]);
  end
  else
    PGIBoxAF(TexteMessage[15], TexteMessage[59]);
end;

procedure TAF_Planning.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (Key = VK_ESCAPE) then
    Close
  else
    if (Key = VK_F5) then
    begin
      if ((Self.ActiveControl = ED_AFF1) or
         (Self.ActiveControl = ED_AFF2) or
         (Self.ActiveControl = ED_AFF3) or
         (Self.ActiveControl = ED_AVE) or
         (Self.ActiveControl = CB_RAP)) then
        begin
          if fCurrentPlanning.fStAffaireSel = '' then
            LoadAffaireSel
          else 
            BCHERCHEARTClick(self);
        end                            
        else
          BT_REFRESHClick(self);
    end;          
end;

procedure TAF_Planning.MnAffaireClick(Sender: TObject);
var
  vStArg: string;
  vItem: TOB;
  vActionEcole: TActionFiche;

begin

  vStArg := '';
  if fCurrentPlanning.fStIdentligne <> '' then
    vStArg := 'IDENTLIGNE=' + fCurrentPlanning.fStIdentligne + ';'; //AB-200511- Identifiant de la commande d'origine

  if fCurrentPlanning.fStAffaire <> '' then
  begin
    vItem := fCurrentPlanning.FComposantPlanning.GetCurItem;
    if getParamSocSecur('SO_AFECOLE', False) and (vItem.GetString('ATA_TYPEPLANIF') = 'ECO') //AB-200512-
    and ((vItem.GetString('APL_ETATLIGNE') = 'FAC') or (vItem.GetString('APL_ETATLIGNE') = 'TER')) then
      vActionEcole := taConsult
    else
      vActionEcole := taModif;

    {AiguillageClickAffaire(fCurrentPlanning.fStAffaire, vStArg, taconsult, vActionEcole, false);}
    fCurrentPlanning.RafraichirItem(vItem, '', 'ARTICLE', True, False);

    // C.B 06/12/2005
    // on ne peut pas seulement raffraichir l'item car il faut rafraichir les materiels associés
    // pour le même jour, la même affaire et le même article
    fCurrentPlanning.RafraichirAssocies(vItem); 
  end
  else
    PGIBoxAF(TexteMessage[15], TexteMessage[59]); //Aucune intervention sélectionnée
end;

procedure TAF_Planning.MnClientClick(Sender: TObject);
begin
  if fCurrentPlanning.fStTiers <> '' then
    V_PGI.DispatchTT(8, taConsult, fCurrentPlanning.fStTiers, '', 'MONOFICHE')
  else
    PGIBoxAF(TexteMessage[15], TexteMessage[59]); //Aucune intervention sélectionnée
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... :
Modifié le ... : 05/04/2005
Description .. : On ouvre toujours la fiche tache en consultation
                 Car la supression d'une ressource peut entrainer la supression du planning
                 dans lequel on est, et il n'y aurait pas de rafraichissement des items
Mots clefs ... :
*****************************************************************}
procedure TAF_Planning.MnTacheClick(Sender: TObject);
var
  vStUpdate : String;
	vItem     : Tob;

begin

  // C.B 07/07/2005
  // On ne fait rien sur les absences, pas de zoom
  if fCurrentPlanning.fStAffaire = traduireMemoire(TexteMessage[75]) then //'ABSENCES'
    PGIBoxAF(TexteMessage[12], TexteMessage[59])   //On ne peux rien faire sur une Absence.

  // C.B 15/05/2005
  // ajout du zoom sur la tache et pas seulement sur l'ensemble des tâches
  else if fCurrentPlanning.fStAffaire <> '' then
  begin
    vItem := fCurrentPlanning.FComposantPlanning.GetCurItem;
  	{if (not JaiLeDroitRessource(vItem.GetString('APL_RESSOURCE'))) then
      vStUpdate := ';ACTION=CONSULTATION'
    else
      vStUpdate := ';ACTION=MODIFICATION';}
//C.B 06/11/2006
//      vStUpdate := ';UPDATE_LIGHT';                    
      vStUpdate := vStUpdate + ';SUPPRESSION_INTERDITE';

    AFLanceFicheBTTaches('ATA_AFFAIRE:' + fCurrentPlanning.fStAffaire +
                         ';ATA_NUMEROTACHE:' + fCurrentPlanning.fStTache +
                         ';ZOOMTACHE' + vStUpdate)
	end
  else
    PGIBoxAF(TexteMessage[15], TexteMessage[59]); // Aucune intervention sélectionnée
end;

procedure TAF_Planning.MnPlanningClick(Sender: TObject);
var
  vSt: string;
  vTob: Tob;
  i: Integer;
  vStAction: string;
  vStMenu : string;
  vStTexte : string;
  vInLg :  Integer;

begin

  if (fCurrentPlanning.fAction = TaConsult) then
    vStAction := 'ACTION=CONSULTATION'
  else
    vStAction := 'ACTION=CREATION';
 
  // ouvrir le planning des ressources  copy
  vInLg := length(POPUPGENERAL.Items[3].Caption) - length(TexteMessage[38]);
  vStMenu 	:= copy(POPUPGENERAL.Items[3].Caption, 1 + vInLg, length(POPUPGENERAL.Items[3].Caption)-vInLg);
  vStTexte 	:= copy(TraduitGA(TexteMessage[38]), 1, length(TexteMessage[38]));
  if (vStMenu = vStTexte) then // Planning des interventions par ressources
  begin            
    // C.B 07/07/2005
    // pas de zoom sur les absences
    if fCurrentPlanning.fStAffaire = traduireMemoire(TexteMessage[75]) then //'ABSENCES'
      PGIBoxAF(TexteMessage[12], TexteMessage[59]) //On ne peux rien faire sur une Absence

    else
      if fCurrentPlanning.fStAffaire <> '' then
    begin
      vTob := TOB.Create('#AFPLANNING', nil, -1);
      try
        vSt := 'SELECT DISTINCT APL_RESSOURCE FROM AFPLANNING WHERE APL_AFFAIRE = "' + fCurrentPlanning.fStAffaire + '"';
        vST := vSt + ' AND APL_DATEDEBPLA >= "' + UsDateTime(fDtDebut) + '"';
        vST := vSt + ' AND APL_DATEFINPLA <= "' + UsDateTime(fDtFin) + '"';
        vTob.LoadDetailFromSQL(vSt);

        // Construction du where passé à l'execPlanning contenant
        // toutes les ressources de l'affaire sélectionnée
        vSt := '';
        if vTob.Detail.Count > 0 then
        begin
          for i := 0 to vTob.detail.count - 1 do
            if vSt = '' then
              vst := 'APL_RESSOURCE = "' + vTob.detail[i].GetValue('APL_RESSOURCE') + '" '
            else
              vSt := vSt + ' OR APL_RESSOURCE ="' + vTob.detail[i].GetValue('APL_RESSOURCE') + '" ';

          vSt := '(' + vst + ')';

          // on lance un second planning dans le même mode que le premier planning (demi-journée ou non)
          ExecPlanning(inttostr(cInMenuPlanning + 5), DateToStr(fDtDebut), DateToStr(fDtFin), DateToStr(fCurrentPlanning.fDtStart), vSt, '', 'X', StDemiJournee, 'ZOOMPLANNING;' + vStAction);
        end;

      finally
        FreeAndNil(vTob);
      end;
    end
    else
      PGIBoxAF(TexteMessage[15], TexteMessage[59]); //Aucune intervention sélectionnée
  end

  else
  begin
    // pas de zoom sur les absences
    if fCurrentPlanning.fStAffaire = traduireMemoire(TexteMessage[75]) then //'ABSENCES'
      PGIBoxAF(TexteMessage[12], TexteMessage[59]) //On ne peux rien faire sur une Absence

    else if fCurrentPlanning.fStRes <> '' then
    begin
      vSt := '(ATA_AFFAIRE ="' + fCurrentPlanning.fStAffaire + '")';
      // on lance un second planning dans le même mode que le premier planning (demi-journée ou non)
      ExecPlanning(inttostr(cInMenuPlanning + 3), DateToStr(fDtDebut), DateToStr(fDtFin), DateToStr(fCurrentPlanning.fDtStart), vSt, '', '-', StDemiJournee, 'ZOOMPLANNING;' + vStAction);
    end
    
    else
      PGIBoxAF(TexteMessage[15], TexteMessage[59]); //Aucune intervention sélectionnée
  end;
end;
 
procedure TAF_Planning.MnInterventionsClick(Sender: TObject);
begin
  // C.B 10/10/2006
	{if fCurrentPlanning.fBoMonPlanning and (fCurrentPlanning.fStRes <> '') then
  	AFLanceFiche_DetailPlanning('RESSOURCE:'+ fCurrentPlanning.fStRes + ';NOFILTRE')
  else if fCurrentPlanning.fStAffaire <> '' then
    AFLanceFiche_DetailPlanning('AFFAIRE:' + fCurrentPlanning.fStAffaire + ';NOFILTRE')
  else
		AFLanceFiche_DetailPlanning('');}
end;

procedure TAF_Planning.MnFactureClick(Sender: TObject);
begin
  //                                         
end;

procedure TAF_Planning.BLignesClick(Sender: TObject);
var
  vStRetour: string;
  Champ: string;
  Valeur: string;
  vStTmp: string;
  vSt: string;

begin

  vSt := '';
  if fStSqlWhere <> '' then
    vSt := 'WHERE:' + fStSqlWhere + ';';

  // C.B 27/09/2005
  // on ne passe pas les criteres si des critères ressource sont concernés
  if pos('ARS_', vSt) > 0 then
    vSt := '';

  // C.B 17/11/2005
  // les ressources peuvent également venir de la table APL
  if pos('APL_RESSOURCE', vSt) > 0 then
    vSt := '';

  // C.B 1710/2006
  // suppression du filtre si on a une affaire sélectionnée
  if (fCurrentPlanning.fStAffaire <> '') and
     (fCurrentPlanning.fStAffaire <> trim(getParamSocSecur('SO_AFPLANAFFAIRE', ''))) then
	  vSt := vSt + 'DANS_PLANNING;NOFILTRE;AFFAIRE:' + fCurrentPlanning.fStAffaire
  else
	  vSt := vSt + 'DANS_PLANNING';

  //vStRetour := AFLanceFicheAFTache_Mul('', vSt);
  if vStRetour = '' then
  begin
    fCurrentPlanning.ClearAffaireSelect;
    La_Selection.Caption := '';
  end
  else
  begin
    vStTmp := (Trim(ReadTokenSt(vStRetour)));
    while (vStTmp <> '') do
    begin
      DecodeArgument(vStTmp, Champ, Valeur);
      if Champ = 'AFFAIRE' then
        fCurrentPlanning.fStAffaireSel := Valeur
      else
        if Champ = 'TIERS' then
        fCurrentPlanning.fStTiersSel := Valeur
      else
        if Champ = 'NUMTACHE' then
        fCurrentPlanning.fStTacheSel := Valeur
      else
        if Champ = 'CODEARTICLE' then
        fCurrentPlanning.fStCodeArtSel := Valeur
      else
        if Champ = 'TYPEARTICLE' then
        fCurrentPlanning.fStTypeArtSel := Valeur
      else
        if Champ = 'ARTICLE' then
        fCurrentPlanning.fStArticleSel := Valeur
      else
        if Champ = 'LIBELLEARTICLE' then
        fCurrentPlanning.fStLibelleArtSel := Valeur;

      {if Champ = 'AFFAIRE' then fCurrentPlanning.fTobSelection.GetString('APL_AFFAIRE') := Valeur
      else if Champ = 'TIERS' then fCurrentPlanning.fTobSelection.GetString('APL_TIERS'):= Valeur
      else if Champ = 'NUMTACHE' then fCurrentPlanning.fTobSelection.GetString('APL_NUMEROTACHE') := Valeur
      else if Champ = 'CODEARTICLE' then fCurrentPlanning.fTobSelection.GetString('APL_CODEARTICLE') := Valeur
      else if Champ = 'TYPEARTICLE' then fCurrentPlanning.fTobSelection.GetString('APL_TYPEARTICLE') := Valeur
      else if Champ = 'ARTICLE' then fCurrentPlanning.fStArticleSel := Valeur
      }

      vStTmp := (Trim(ReadTokenSt(vStRetour)));
    end;

    La_Selection.Caption := fCurrentPlanning.ModeCreationDirecte(True);
  end;
end;

procedure TAF_Planning.BT_EFFACELIGNESClick(Sender: TObject);
begin
  fCurrentPlanning.fStAffaireSel := '';
  fCurrentPlanning.fStTacheSel := '';
  fCurrentPlanning.fStCodeArtSel := '';     
  ED_AFF1.text := '';
  ED_AFF2.text := '';
  ED_AFF3.text := '';
  ED_AVE.text := '';
  ED_TACHE.text := '';
  ED_TIERS.text := '';
  DetruitMaTobInterne('SELECTART');
  fCurrentPlanning.fTobArtSel := nil;
  
  La_Selection.Caption := '';
  fCurrentPlanning.fInModeCreation := cInCreationSansSelection;
  BT_Lignes.down := False;
  BTNonProductive.down := False;
  ED_AVE.OnExit := nil;
  if ED_AFF1.canFocus then
    ED_AFF1.setFocus;
  ED_AVE.OnExit := ED_AVEExit;
end;
    
// Journée non productive
procedure TAF_Planning.BTNonProductiveClick(Sender: TObject);
var
  vStRetour: string;
  Champ: string;
  Valeur: string;
  vStTmp, vStArg: string;
  vStLibTache: string;
begin
  vStRetour := '';
  vStLibTache := '';
  try
    if trim(getParamSocSecur('SO_AFPLANAFFAIRE', '')) = '' then
    begin
      if PGIAskAF(TraduitGA(TexteMessage[21]), '') <> mrYes then
        Exit;
      vStArg := 'AFF_STATUTAFFAIRE=AFF;';
      if getParamSocSecur('SO_AFTIERSPLANNING', '') <> '' then
        vStArg := vStArg + 'AFF_TIERS=' + getParamSocSecur('SO_AFTIERSPLANNING', '') + ';';
      vStRetour := AFLanceFiche_AffaireRECH('', vStArg);
      if vStRetour = '' then
        Exit;
      SetParamsoc('SO_AFPLANAFFAIRE', ReadTokenSt(vStRetour))
    end;
    //vStRetour := AFLanceFicheAFTache_Mul('', 'AFFAIRE:' + getParamSocSecur('SO_AFPLANAFFAIRE', '') + ';NONPRODUCTIVE');
    if vStRetour = '' then
      Exit;

    vStTmp := (Trim(ReadTokenSt(vStRetour)));
    while (vStTmp <> '') do
    begin
      DecodeArgument(vStTmp, Champ, Valeur);
      if Champ = 'AFFAIRE' then
        fCurrentPlanning.fStAffaireSel := Valeur
      else
        if Champ = 'TIERS' then
        fCurrentPlanning.fStTiersSel := Valeur
      else
        if Champ = 'NUMTACHE' then
        fCurrentPlanning.fStTacheSel := Valeur
      else
        if Champ = 'CODEARTICLE' then
        fCurrentPlanning.fStCodeArtSel := Valeur
      else
        if Champ = 'TYPEARTICLE' then
        fCurrentPlanning.fStTypeArtSel := Valeur
      else
        if Champ = 'ARTICLE' then
        fCurrentPlanning.fStArticleSel := Valeur
      else
        if Champ = 'LIBELLETACHE1' then
        vStLibTache := Valeur;

      //      if Champ = 'AFFAIRE' then fCurrentPlanning.fTobSelection.GetString('APL_AFFAIRE') := Valeur;
      vStTmp := (Trim(ReadTokenSt(vStRetour)));
    end;

    La_Selection.Caption := Format(traduitGA(TexteMessage[46]), [vStLibTache]); //Journée non facturable : %s
    fCurrentPlanning.fInModeCreation := cInCreationNonProductive;
  finally
    if vStLibTache = '' then
    begin
      BTNonProductive.down := false;
      if fCurrentPlanning.fInModeCreation = cInCreationDirecte then
        BT_LIGNES.down := true
      else
      begin
        fCurrentPlanning.ClearAffaireSelect;
        La_Selection.Caption := '';
        fCurrentPlanning.fInModeCreation := cInCreationSansSelection;
      end;
    end;
  end;
end;

procedure TAF_Planning.MnEcoleClick(Sender: TObject);
begin
  AFLanceFiche_MulAffaire('AFF_STATUTAFFAIRE=AFF', 'STATUT=AFF;PASMULTI;TOUTEAFFAIRE;ECOLE');
end;

procedure TAF_Planning.BT_REFRESHClick(Sender: TObject);
var
  vDtStart: TDateTime;
begin

	if fCurrentPlanning.FComposantPlanning.GetCurItem <> nil then
    vDtStart := fCurrentPlanning.FComposantPlanning.GetCurItem.GetValue('APL_DATEDEBPLA')
  else
    vDtStart := fCurrentPlanning.FComposantPlanning.GetDateOfCol(fCurrentPlanning.FComposantPlanning.col);
  fCurrentPlanning.FComposantPlanning.Activate := False;
  fCurrentPlanning.LoadItemsPlanning(fStSqlWhere + fStWhereEtatLg);
                                                
  // C.B 25/04/2006
  // pour les absences, on ne passe pas le where de état ligne
  if fStAbsences = 'X' then
    fCurrentPlanning.LoadAbsencesPlanning(fStSqlWhere);
  fCurrentPlanning.FComposantPlanning.DateOfStart := vDtStart;
  fCurrentPlanning.FComposantPlanning.Activate := True;
end;

procedure TAF_Planning.MnListeIntervClick(Sender: TObject);
var
  vItem   : Tob;
  vStDeb  : String;
  vStFin  : String;

  vInNumSem : Integer;
  vInYear   : Integer;


begin
	// liste planning
	{if fCurrentPlanning.fBoMonPlanning and (fCurrentPlanning.fStRes <> '') then
  	AFLanceFiche_ListeTachePla('RESSOURCE='+ fCurrentPlanning.fStRes)
  else if fCurrentPlanning.fStAffaire <> '' then
		AFLanceFiche_ListeTachePla('RESSOURCE='+ fCurrentPlanning.fStRes)
  else
		AFLanceFiche_ListeTachePla('');}

  // C.B 28/11/2006
  // demande dans la FQ n° 13087
  // On ne met plus le critère affaire mais que le critère ressource
  // + ';' + 'AFFAIRE=' + fCurrentPlanning.fStAffaire
  // debut de semaine sélectionnée + semaine suivante
  vItem := fCurrentPlanning.FComposantPlanning.GetCurItem;

  vInNumSem := NumSemaine(vItem.GetValue('APL_DATEDEBPLA'), vInYear);
  vStDeb := DateToStr(PremierJourSemaine(vInNumSem, vInYear));
  vStFin := DateToStr(PremierJourSemaine(vInNumSem + 2, vInYear) -1); 

	//AFLanceFiche_ListeTachePla('RESSOURCE=' + vItem.GetString('APL_RESSOURCE') + ';' + 'DATEDEB=' + vStDeb + ';' + 'DATEFIN=' + vStFin);
  
end;

//Recherche Affaire
procedure TAF_Planning.ED_AVEExit(Sender: TObject);
begin
  LoadAffaireSel;
end;

procedure TAF_Planning.ED_AFF1Change(Sender: TObject);
begin
  GoSuiteCodeAffaire(ED_AFF1, ED_AFF2, 1);
end;

procedure TAF_Planning.ED_AFF2Change(Sender: TObject);
begin
  GoSuiteCodeAffaire(ED_AFF2, ED_AFF3, 2);
end;

procedure TAF_Planning.ED_AFF3Change(Sender: TObject);
begin
  GoSuiteCodeAffaire(ED_AFF3, ED_AVE, 3);
end;

procedure TAF_Planning.BCHERCHEAFFClick(Sender: TObject);
begin
  LoadAffaireSel;
end;

procedure TAF_Planning.BCHERCHEARTClick(Sender: TObject);
var
  vStRet  : String;
begin
  //vStRet := AFLanceFicheAFPLANNINGART ('', 'NUMTACHE=' + fCurrentplanning.fStTacheSel);
  if vStRet <> '' then
    SelectArt(StrToInt(vStRet), '');
end;

procedure TAF_Planning.CB_RAPClick(Sender: TObject);
begin           
  SelectTaches;
end;

{------------------------------------------------------------------------------}
{ OBJET TPlanningGA                                                             }
{------------------------------------------------------------------------------}

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 05/03/2002
Modifié le ... : 22/02/2005
Description .. : creation, chargement du modele ...
Mots clefs ... :
*****************************************************************}

constructor TPlanningGA.Create(pFichePlanning: TAF_Planning);
begin
  // Création du planning
  FComposantPlanning := THPLanning.Create(pFichePlanning);
  FFichePlanning := pFichePlanning;
  FComposantPlanning.Parent := FFichePlanning.PPlanningJour;

  CreateTob;

  // liaison des evenements
  FComposantPlanning.OnModifyItem := ModifyItem;
  FComposantPlanning.OnDeleteItem := DeleteItem;
  FComposantPlanning.OnCreateItem := CreateItem;
  FComposantPlanning.OnInitItem := InitItem;
  FComposantPlanning.OnMoveItem := MoveItem;
  FComposantPlanning.OnCopyItem := CopyItem;

  FComposantPlanning.OnDblClick := DoubleClick;
  FComposantPlanning.OnDblClickSpec := DoubleClickSpec;

  FComposantPlanning.OnOptionPopup := OnPopup;
  FComposantPlanning.OnAvertirApplication := AvertirApplication;

  FComposantPlanning.OnLink := Link;
  FComposantPlanning.OnBeforeChange := BeforeChange;
end;

destructor TPlanningGA.destroy;
begin
  // Suppression du planning
  FreeAndNil(FComposantPlanning);

  if assigned(fTobEtatXml) then
    FreeAndNil(FTobEtatXml);
  //if assigned(fTobAdresses) then FreeAndNil(fTobAdresses);
  //if fTobTiers <> nil then FreeAndNil(fTobTiers) ;

  if assigned(fTobsPlanningTobItems) then
    FreeAndNil(fTobsPlanningTobItems);
  if assigned(fTobsPlanningTobEtats) then
    FreeAndNil(fTobsPlanningTobEtats);
  if assigned(fTobsPlanningTobRes) then
    FreeAndNil(fTobsPlanningTobRes);
  if assigned(fTobsPlanningTobCols) then
    FreeAndNil(fTobsPlanningTobCols);
  if assigned(fTobsPlanningTobRows) then
    FreeAndNil(fTobsPlanningTobRows);
  if assigned(fTobsPlanningTobEvents) then
    FreeAndNil(fTobsPlanningTobEvents);
  if assigned(fTobEtats) then
    FreeAndNil(fTobEtats);
  if assigned(fTobAffAdm) then
    FreeAndNil(fTobAffAdm);
end;

procedure TPlanningGA.CreateTob;
begin
  // Création du planning
  fTobsPlanningTobItems := TOB.Create('Les_items', nil, -1);
  fTobsPlanningTobRes := TOB.Create('Les_ressources', nil, -1);
  fTobsPlanningTobEtats := TOB.Create('Les_etats', nil, -1);
  fTobsPlanningTobCols := TOB.Create('Les_colonnes', nil, -1);
  fTobsPlanningTobRows := TOB.Create('Les_lignes', nil, -1);
  fTobsPlanningTobEvents := TOB.Create('Les_evenements', nil, -1);
  fTobEtatXml := TOB.Create('#AFPLANNINGPARAM', nil, -1);
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 09/05/2005
Modifié le ... :   /  /
Description .. : passage en mode creation directe ou creation avec selection d'article
                 retourne la chaine de sélection à afficher
Mots clefs ... :
*****************************************************************}

function TPlanningGA.ModeCreationDirecte(pBoLoadRessource: Boolean): string;
var
  vStAff: string;
  vStResLib: string;
  vStLibArt : String;

begin
  vStAff := CodeAffaireAffiche(fStAffaireSel, ' ');

  // C.B 28/08/2006
  // tronquer l'affichage du libellé
  vStLibArt := copy(fStLibelleArtSel, 1, 35);
  result := format(TraduitGa(texteMessage[65]), [vStAff, fStTiersSel, fStCodeArtSel, vStLibArt]); // 'Affaire : %s      Tiers :  %s     Article :  %s  %s     '

  {    if UniqueRes(fCurrentPlanning.fTobSelection.GetString('APL_AFFAIRE'),
        fCurrentPlanning.fStTacheSel,
        fCurrentPlanning.fStResSel, vStResLib) then
  }
  if pBoLoadRessource then
  begin
    {if UniqueRes(fStAffaireSel, fStTacheSel, fStResSel, vStResLib) and
      getParamSocSecur('SO_AFAFFECTRES', False) then
    begin
      result := result + format(TraduitGA(TexteMessage[47]), [fStResSel, vStResLib]);
      fInModeCreation := cInCreationDirecte;
    end
    else
      fInModeCreation := cInCreationAvecSelection;}
  end
  else
    fInModeCreation := cInCreationAvecSelection;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 28/02/2002
Modifié le ... :
Description .. :
Suite ........ :
Mots clefs ... :
*****************************************************************}

function TPlanningGA.Load(pDtDebut, pDtFin, pDtStart: TDateTime; pStWhere, pStCodeModele, pStWhereEtatLigne, pStAbscences, pStDemiJournee, pStArgument: string; pTOBModelePlanning: Tob): Boolean;
begin
 
  result := true;
  fTobModelePlanning := pTobModelePLanning;
  DecodeCodeModele(pStCodeModele);
  fStWhere := pStWhere;
  fBoGestionAbs := pStAbscences = 'X';
  fBoDemiJournee := pStDemiJournee = 'X';
  fStArgument := pStArgument;
  fDtDebut := pDtDebut;
  fDtFin := pDtFin;
  fInModeCreation := cInCreationSansSelection;
 
  // surbooking autorisé, affichage de tous les articles
  FFichePlanning.CB_RAP.onclick := nil;
  FFichePlanning.CB_RAP.Checked := not GetParamsocSecur('SO_AFSURBOOKING', False);
  FFichePlanning.CB_RAP.OnClick := FFichePlanning.CB_RAPClick;
                                                      
  // C.B 26/09/2006
  // statut par defaut de l'item
  fStStatutDefaut := getparamsocSecur('SO_AFPLANNINGETAT', '');

  // test si l'état par défaut est défini
  if fStStatutDefaut = '' then
    PGIBoxAF(TexteMessage[16], TexteMessage[59]) //Le paramètre société "Statut par défaut" du planning doit être défini pour pouvoir utiliser le planning.
  else
  begin
    // chargement de l'action
    LoadArgument;

    // chargement des adresses d'intervention
    //LoadAdresses;

    // chargement des tiers
    //LoadTiers;

    // Charge les paramètres utilisateurs
    InitParamUtil;

    // selection du type de planning (ex tache par affaire et par ressource)
    SelectPlanning;

    // initialisation des paramatres generaux du planning (ex : champs, couleurs)
    InitPlanning;

    //C.B 21/09/2006
    //fTobAffAdm := LoadAffaireAdministrative;

    // on charge le planning, meme s'il n'y a pas de lignes de planning
    // loadEtat avant LoadItemsPlanning car DisplayHint dans LoadItemsPlanning
    // utilise la tobEtat
    if LoadEtatPlanning then
    begin
      LoadItemsPlanning(pStWhere + pStWhereEtatLigne);
      LoadAffNonProductive; // Journée non productive
      if LoadResPlanning(pStWhere) then
      begin
		  	// C.B 25/04/2006
  			// pour les absences, on ne passe pas le where de état ligne
        if pStAbscences = 'X' then       
          LoadAbsencesPlanning(pStWhere);
        Display(pDtStart);
      end
      else
        // pas de ressource
        result := false;
    end
    else
      result := false;

    FFichePlanning.ED_AVE.OnExit := nil;
    if FFichePlanning.ED_AFF1.canFocus then
      FFichePlanning.ED_AFF1.setFocus;
    FFichePlanning.ED_AVE.OnExit := FFichePlanning.ED_AVEExit;

    if fAction = TaConsult then
      FFichePlanning.Pa_Quick.Visible := False;
 
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 22/02/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TPlanningGA.Display(pDtStart: TDateTime);
begin
  FComposantPlanning.Activate := False;
  if pDtStart = iDate1900 then
    pDtStart := DebutDeMois(now);

  FComposantPlanning.IntervalDebut := fDtDebut;
  FComposantPlanning.IntervalFin := fDtFin;
  FComposantPlanning.DateOfStart := pDtStart;

  // Les Tobs
  FComposantPlanning.TobItems := fTobsPlanningTobItems;
  FComposantPlanning.TobEtats := fTobsPlanningTobEtats;
  FComposantPlanning.TobRes := fTobsPlanningTobRes;

  // Activation du planning
  FComposantPlanning.Activate := True;

  // chargement de l'etat pas defaut
  fTobEtatDefaut := fTobsPlanningTobEtats.FindFirst(['E_CODE'], [getParamSocSecur('SO_AFPLANNINGETAT', '')], true);
  if fTobEtatDefaut = nil then
    PGIBoxAF(TexteMessage[11], TexteMessage[59]); //L'état par défaut défini dans les paramètres n''existe pas !
  
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 15/03/2002
Modifié le ... : 25/04/2005
Description .. : charge les etats dans le popupmenu
Suite ........ : on ne charge pas les etats n'autorisant pas de modifications
Mots clefs ... :
*****************************************************************}
procedure TPlanningGA.LoadPopUpMenu(pBoEtatPourMonPlanning: Boolean);
var
  vQR: TQuery;
  vSt: string;
  vTobEtat: TOB;
  i, j: Integer;
  vStEtats: string;

begin

  //C.B 26/10/2004 GESTION DES ABSENCES
  vST := 'SELECT APP_CODEPARAM, APP_LIBELLEPARAM FROM AFPLANNINGPARAM ';
  vST := vST + ' WHERE APP_TYPEAFPARAM="STA" AND APP_CODEPARAM <> "ABS"';
  vST := vST + ' AND APP_CODEPARAM <> "DEC" AND APP_MODIFIABLE = "X"';
  vQr := nil;
  vTobEtat := Tob.Create('#AFPLANNINGPARAM', nil, -1);

  try
    vQR := OpenSql(vSt, True,-1,'',true);
    if not vQR.Eof then
      vTobEtat.LoadDetailDB('AFPLANNINGPARAM', '', '', vQR, False, True);
    FComposantPlanning.AddOptionPopup(1, '-');

    fNbEtat := vTobEtat.detail.count;
    j := 0;
    for i := 0 to vTobEtat.detail.count - 1 do
    begin
      // C.B 23/05/2005
      // si on est dans mon planning, on a certains etats autorisés
      if pBoEtatPourMonPlanning then
      begin
        // C.B 22/08/2005
        // FAC est supprimer de la tablette et on gere la valeur TOUS
        if (pos(vTobEtat.detail[i].GetString('APP_CODEPARAM'), 'FAC') > 0) then
        {begin
          if SaisiePlanningManager then
            FComposantPlanning.AddOptionPopup(i - j + cInPopUp, vTobEtat.Detail[i].GetValue('APP_LIBELLEPARAM'))
        end}
        else
        begin
          //C.B 11/09/2006
          // la notion de aucun n'existe pas, on a du mettre tous
          // si tous, on considère que c'est aucun pour sic
          // pour le cas standard, tous veut toujours dire tous
          // mais on ajoutera un état dans la table pour gérer le
          // aucun en V8 
          vStEtats := getParamSocSecur('SO_AFETATRES','');
          if VH_GC.GCIfDefCEGID then
          begin
          	if (pos(vTobEtat.detail[i].GetString('APP_CODEPARAM'), vStEtats) > 0) then
              FComposantPlanning.AddOptionPopup(i - j + cInPopUp, vTobEtat.Detail[i].GetValue('APP_LIBELLEPARAM'));
          end
          else if (pos(vTobEtat.detail[i].GetString('APP_CODEPARAM'), vStEtats) > 0) or (vStEtats = '') then
            FComposantPlanning.AddOptionPopup(i - j + cInPopUp, vTobEtat.Detail[i].GetValue('APP_LIBELLEPARAM'));
        end;
      end

        // C.B 19/07/2005
        // ajout du concept permettant de modifier les états bloqués
      else
        {if (EtatModifAutorisee(vTobEtat.detail[i].GetString('APP_CODEPARAM'), True, True) and
        (fAction <> TaConsult)) then
        FComposantPlanning.AddOptionPopup(i - j + cInPopUp, vTobEtat.Detail[i].GetValue('APP_LIBELLEPARAM'))
      else
      begin
        fNbEtat := fNbEtat - 1;
        j := j + 1;
      end;}
    end;
 
    // C.B 11/07/2006
    // ajout d'un menu cegid feuille de présence reçu
    // C.B 11/09/2006
    // restriction sur autoplanification
    {if VH_GC.GCIfDefCEGID and autoplanification then
    begin
      fNbEtat := fNbEtat + 1;
			FComposantPlanning.AddOptionPopup(vTobEtat.detail.count - j + cInPopUp, '-');
      FComposantPlanning.AddOptionPopup(vTobEtat.detail.count - j + 1 + cInPopUp, 'Feuille de présence reçue'); // traduction inutile, spécif cegid
    end;}
     
  finally
    Ferme(vQR);
    vTobEtat.Free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 28/02/2002
Modifié le ... : copie la tob retournée par TOBLoadFromXMLStream
Description .. : dans la tob de travail
Suite ........ :
Mots clefs ... :
*****************************************************************}

function TPlanningGA.LoadTobEtat(T: TOB): integer;
begin
  with TOB.Create('fille_param', fTobsPlanningTobEtats, -1) do
  try
    Dupliquer(T.detail[0], True, True);
  finally
    FreeAndNil(T);
    result := 0;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 22/02/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TPlanningGA.InitPlanning;
begin
  if fBoDemiJournee then
  begin
    FComposantPlanning.Interval := piDemiJour;
    FComposantPlanning.CumulInterval := pciSemaine;

    FComposantPlanning.DebutAM := getParamSocSecur('SO_AFAMDEBUT','00:00:00');
    FComposantPlanning.FinAM := getParamSocSecur('SO_AFAMFIN','00:00:00');
    FComposantPlanning.DebutAP := getParamSocSecur('SO_AFPMDEBUT','00:00:00');
    FComposantPlanning.FinAP := getParamSocSecur('SO_AFPMFIN','00:00:00');

    FComposantPlanning.ChampDateDebut := 'APL_HEUREDEB_PLA';
    FComposantPlanning.ChampDateFin := 'APL_HEUREFIN_PLA';
                                                  
    //FComposantPlanning.DateFormatAP := 'dd mmm #13#10 ddd "AP"';
    //FComposantPlanning.DateFormatAM := 'dd mmm #13#10 ddd "MA"';
  end
  else
    if fStTypeModele = 'J' then
  begin
    FComposantPlanning.Interval := piJour;
    FComposantPlanning.CumulInterval := pciSemaine;
    FComposantPlanning.ChampDateDebut := 'APL_DATEDEBPLA';
    FComposantPlanning.ChampDateFin := 'APL_DATEFINPLA';
  end;

  // affichage des lignes de date
  FComposantPlanning.ActiveLigneGroupeDate := true;
  FComposantPlanning.ActiveLigneDate := true;

  // Agrandissement
  FComposantPlanning.Align := alClient;

  // Surbooking et week end
  FComposantPlanning.SurBooking := True;

  // C.B 12/10/2006         
  // ajout de la possibilité de cacher les week-ends                                
  if getparamsocSecur('SO_AFPLANNINGWE', True) then
  begin
    FComposantPlanning.ActiveSaturday := True;
    FComposantPlanning.ActiveSunday := True;
    FComposantPlanning.VisibleSaturday := True;
    FComposantPlanning.VisibleSunday := True;
  end
  else
  begin
    FComposantPlanning.VisibleSaturday := False;
    FComposantPlanning.VisibleSunday := False; 
  end;

  FComposantPlanning.GestionJoursFeriesActive := True;

  //Les autorisations
  if fAction = TaConsult then
    changeOptions(False)
  else
    changeOptions(True);

  //Le Zoom
  FComposantPlanning.Zoom := False;

  // Les champs des états
  FComposantPlanning.EtatChampBackGroundColor := 'E_COULEURFOND';
  FComposantPlanning.EtatChampCode := 'E_CODE';
  FComposantPlanning.EtatChampFontColor := 'E_COULEURFONTE';
  FComposantPlanning.EtatChampFontName := 'E_NOMFONTE';
  FComposantPlanning.EtatChampFontSize := 'E_TAILLEFONTE';
  FComposantPlanning.EtatChampFontStyle := 'E_STYLEFONTE';
  //FComposantPlanning.EtatChampIcone := 'E_ICONE' ;
  FComposantPlanning.EtatChampLibelle := 'E_LIBELLE';

  // Les champs Item
  FComposantPlanning.ChampLineID := 'I_CODE';
  FComposantPlanning.ChampEtat := 'APL_ETATLIGNE';
// définissait une couleur de fond pour l'item, mais est remlpacé par la couleur de l'item  
//  FComposantPlanning.ChampColor := 'I_COLOR';
  FComposantPlanning.ChampHint := 'I_HINT';

  // C.B 13/05/2005 spécifique cegid
  if VH_GC.GCIfDefCEGID then
    FComposantPlanning.ChampIcone := 'I_ICONE';

  FComposantPlanning.ChampLibelle := 'CHAMPLIBELLE';

  // Les champs des ressources
  FComposantPlanning.ResChampColor := '';
  FComposantPlanning.ResChampFixedColor := '';
  FComposantPlanning.ResChampID := 'CLE';
  FComposantPlanning.ResChampReadOnly := '';

  // colonnes fixes
  if fColonne3.StField <> '' then
  begin
    FComposantPlanning.TokenFieldColFixed := fColonne1.StField + ';' + fColonne2.StField + ';' + fColonne3.StField;
    FComposantPlanning.TokenAlignColFixed := fColonne1.StAlign + ';' + fColonne2.StAlign + ';' + fColonne3.StAlign;
    FComposantPlanning.TokenFieldColEntete := fColonne1.StEntete + ';' + fColonne2.StEntete + ';' + fColonne3.StEntete;
    FComposantPlanning.TokenSizeColFixed := inttostr(FComposantPlanning.ColSizeEntete) + ';' + inttostr(FComposantPlanning.ColSizeEntete) + ';' + inttostr(FComposantPlanning.ColSizeEntete);
  end
  else
    if FColonne2.StField <> '' then
  begin
    FComposantPlanning.TokenFieldColFixed := fColonne1.StField + ';' + fColonne2.StField;
    FComposantPlanning.TokenAlignColFixed := fColonne1.StAlign + ';' + fColonne2.StAlign;
    FComposantPlanning.TokenFieldColEntete := fColonne1.StEntete + ';' + fColonne2.StEntete;
    FComposantPlanning.TokenSizeColFixed := inttostr(FComposantPlanning.ColSizeEntete) + ';' + inttostr(FComposantPlanning.ColSizeEntete);
  end
  else
  begin
    FComposantPlanning.TokenFieldColFixed := fColonne1.StField;
    FComposantPlanning.TokenAlignColFixed := fColonne1.StAlign;
    FComposantPlanning.TokenFieldColEntete := fColonne1.StEntete;
    FComposantPlanning.TokenSizeColFixed := inttostr(FComposantPlanning.ColSizeEntete);
  end;

  FComposantPlanning.DisplayEnteteColFixed := true;

  // evenements
  FComposantPlanning.GestionEvenements := False;

  // C.B 16/08/2005
  // permet de prendre en comàte les jours fériés définis dans la base
  FComposantPlanning.JourFerieCalcule := False;

  //C.B 29/05/2006
  // couleur de fond dégradé
  FComposantPlanning.ColorBackGroundBasStart := FMenuGStartColor;
	FComposantPlanning.ColorBackGroundBasEnd   := FMenuGEndColor;

end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 16/04/2002
Modifié le ... :
Description .. : chargement des paramètes du planning
Suite ........ :
Mots clefs ... :
*****************************************************************}

procedure TPlanningGA.InitParamUtil;
var
  vStFormeGraph: string;
  vTob: Tob;
  vQR: TQuery;
  vSt: string;
  
begin

  with FComposantPlanning do
  begin
    // Changement date d'un item sur planning possible O/N
    MoveHorizontal := true;

    // Chargement de la forme graphique
    if fTOBModelePlanning.detail.count = 0 then
    begin
      PGIBoxAF(TexteMessage[26], TexteMessage[59]); //Revoir les paramètres des plannings
      exit; 
    end
    else
      vStFormeGraph := fTOBModelePlanning.detail[0].Getvalue('FORMEGRAPHIQUE');

    if (vStFormeGraph = 'PGF') then
      FormeGraphique := pgFleche
    else
      if (vStFormeGraph = 'PGB') then
      FormeGraphique := pgBerceau
    else
      if (vStFormeGraph = 'PGL') then
      FormeGraphique := pgLosange
    else
      if (vStFormeGraph = 'PGE') then
      FormeGraphique := pgEtoile
    else
      if (vStFormeGraph = 'PGA') then
      FormeGraphique := pgRectangle
    else
      if (vStFormeGraph = 'PGI') then
      FormeGraphique := pgBisot
    else
      FormeGraphique := pgFleche;

    //Mise à jour du format de la date pour le planning Jour
    if fStTypeModele = 'J' then
    begin

      // format jour
      vSt := 'SELECT CO_ABREGE FROM COMMUN WHERE CO_TYPE = "HFD" and CO_CODE = "';
      vSt := vSt + fTOBModelePlanning.detail[0].Getvalue('FORMATDATE') + '"';
      vTob := Tob.Create('tablette_dates', nil, -1);
      vQR := OpenSql(vSt, True,-1,'',true);
      try
        if not vQR.Eof then
        begin
          vTob.LoadDetailDB('COMMUN', '', '', vQR, False, True);
          DateFormat := vTob.Detail[0].Getvalue('CO_ABREGE');
        end;
      finally
        Ferme(vQR);
        FreeAndNil(vTob);
      end;

      // format demi-journée
      if fTOBModelePlanning.detail[0].GetNumChamp('FORMATDATEDJ') <> -1 then
      begin
        vSt := 'SELECT CO_ABREGE FROM COMMUN WHERE CO_TYPE = "AFJ" and CO_CODE = "';
        vSt := vSt + fTOBModelePlanning.detail[0].Getvalue('FORMATDATEDJ') + '"';
        vTob := Tob.Create('tablette_dates', nil, -1);
        vQR := OpenSql(vSt, True,-1,'',true);
        try
          if not vQR.Eof then
          begin
            vTob.LoadDetailDB('COMMUN', '', '', vQR, False, True);
            DateFormatAP := vTob.Detail[0].Getvalue('CO_ABREGE') + 'ddd "AP"';
            DateFormatAM := vTob.Detail[0].Getvalue('CO_ABREGE') + 'ddd "MA"';
          end;
        finally
          Ferme(vQR);
          FreeAndNil(vTob);
        end;
			end
      else
      begin
        DateFormatAP := 'dd mmm #13#10 ddd "AP"';
        DateFormatAM := 'dd mmm #13#10 ddd "MA"';
      end;
    end;

    // MAJ couleur
    ColorSelection := fTOBModelePlanning.detail[0].Getvalue('LSELECTION');
    ColorBackground := fTOBModelePlanning.detail[0].Getvalue('LFOND');
    ColorOfSaturday := fTOBModelePlanning.detail[0].Getvalue('LSAMEDI');
    ColorOfSunday := fTOBModelePlanning.detail[0].Getvalue('LDIMANCHE');
    ColorJoursFeries := fTOBModelePlanning.detail[0].Getvalue('LJOURSFERIES');

    // C.B 19/11/2004 ne fonctionne pas
    // MAJ de la police des colonnes
    //FontNameColRowFixed := fTOBModelePlanning.detail[0].Getvalue('LFONTNAME');

    // Encadrement
    FrameOn := True;
    MultiLine := True;

    //Lignes du planning
    RowFieldID := 'R_CODE';

    //P.RowFieldReadOnly := 'R_RO';
    RowFieldColor := 'R_COLOR';
    MoveCur(FALSE);

    // Activation bulles d'aide
    ShowHint := True;
    MouseAlready := True;
    DisplayIcon := True;

    // MAJ Taille des colonnes
    ColSizeData := fTOBModelePlanning.detail[0].GetInteger('LARGEURCOL');
    ColSizeEntete := fTOBModelePlanning.detail[0].GetInteger('LARGEURCOLENTETE');

    if fTOBModelePlanning.detail[0].GetNumChamp('HAUTEURLIG') <> -1 then
      RowSizeData := fTOBModelePlanning.detail[0].GetINteger('HAUTEURLIG')
		else
      RowSizeData := 40;

    // en dur pour l'instant
    RowSizeEntete := 30;
                             
    if RowSizeData = 0 then
      RowSizeData := 40;

    // C.B 25/11/2005
    // était limité pour pb de raffraichissement ok maintenant
    //if int(ColSizeEntete) > 110 then ColSizeEntete := 110;
  end;
end;

procedure TPlanningGA.DecodeCodeModele(pStCodeModele: string);
begin

  //peut etre un parametrage
  fStTypeModele := copy(pStCodeModele, 3, 1);

  if (pStCodeModele = 'P1J') then
    fInNumModele := 0
  else
    if (pStCodeModele = 'P2J') then
    fInNumModele := 1
  else
    if (pStCodeModele = 'P3J') then
    fInNumModele := 2
  else
    if (pStCodeModele = 'P4J') then
    fInNumModele := 3;

  fStNumPlanning := copy(pStCodeModele, 2, 1);

end;

procedure TPlanningGA.DoubleClick(Sender: TObject);
var
  vBool: Boolean;
begin
  vBool := false;
  modifyItem(sender, FComposantPlanning.GetCurItem, vBool);
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 15/04/2002
Modifié le ... :
Description .. : Gestion des modifications des items
Suite ........ :
Mots clefs ... :
*****************************************************************}

procedure TPlanningGA.AvertirApplication(Sender: TObject; FromItem: TOB; ToItem: TOB; Actions: THPlanningAction);
var
  i: Integer;
  vSt: string;
  vTob: Tob;
  //vTobItem      : Tob;
  vTobItemDest: Tob;
  //vBoCancel     : Boolean;

begin
  case Actions of
    paClickRight:
      if (fromItem <> nil) and (FromItem.GetValue('APL_ETATLIGNE') <> 'ABS') then
      begin
        for i := 0 to fNbEtat - 1 do
          FComposantPlanning.EnableOptionPopup(i + cInPopUp, true);
        // C.B 25/08/2006
        // il y a un trait de plus pour feuille de présence recu
        if VH_GC.GCIfDefCEGID then
				  FComposantPlanning.EnableOptionPopup(fNbEtat + cInPopUp, true);
      end
      else
      begin
        for i := 0 to fNbEtat - 1 do
          FComposantPlanning.EnableOptionPopup(i + cInPopUp, false);
        // C.B 25/08/2006
        // il y a un trait de plus pour feuille de présence recu
        if VH_GC.GCIfDefCEGID then
				  FComposantPlanning.EnableOptionPopup(fNbEtat + cInPopUp, false);
      end;
    paClickLeft:
      begin
        if assigned(fromItem) then
        begin
          fDtStart := FromItem.GetValue('APL_DATEDEBPLA');
          fStAffaire := FromItem.GetString('APL_AFFAIRE');
          fStTiers := FromItem.GetString('APL_TIERS');
          fStRes := FromItem.GetString('APL_RESSOURCE');
          fStTache := FromItem.GetString('APL_NUMEROTACHE');
          // C.B 22/09/2005 affichage du commentaire de la tache en plus !
          // C.B 08/09/2005
          // affichage du commentaire
          AffichageCommentaire(FromItem);
          // C.B 12/12/2005
          // on va calculer les quantités des la tache pour chaque item
          AffichageQuantites(FromItem);
        end
        else
        begin
          fStAffaire := '';
          fStTiers := '';
          fStRes := '';
          fStTache := '';
          fDtStart := iDate1900;
          // C.B 08/09/2005
          // affichage du commentaire
          FFichePlanning.LA_COMMENTAIRE.Caption := '';
          FFichePlanning.LA_QUANTITE.Caption := '';
        end;
      end;

    PaMove:
      begin
        // message pour déplacer tout ce qui est prevu ce jour pour cette tache
        if (getParamSocSecur('SO_AFPLANAFFAIRE','') <> FromItem.GetString('APL_AFFAIRE')) then
        begin
          vSt := 'SELECT APL_AFFAIRE FROM AFPLANNING ';
          vSt := vSt + ' WHERE APL_AFFAIRE = "' + FromItem.GetString('APL_AFFAIRE') + '"';
          vSt := vSt + ' AND APL_NUMEROTACHE = ' + FromItem.GetString('APL_NUMEROTACHE');
          vSt := vSt + ' AND APL_TYPEPLA = "PLA"';
          vSt := vSt + ' AND APL_NUMEROLIGNE <> ' + FromItem.GetString('APL_NUMEROLIGNE');

          if not fBoDemiJournee then
          begin
            vSt := vSt + ' AND APL_DATEDEBPLA = "' + UsDateTime(FromItem.GetValue('APL_DATEDEBPLA')) + '"';
            vSt := vSt + ' AND APL_DATEFINPLA = "' + UsDateTime(FromItem.GetValue('APL_DATEFINPLA')) + '"';
          end
          else
          begin
            vSt := vSt + ' AND APL_HEUREDEB_PLA = "' + UsTime(FromItem.GetValue('APL_HEUREDEB_PLA')) + '"';
            vSt := vSt + ' AND APL_HEUREFIN_PLA = "' + UsTime(FromItem.GetValue('APL_HEUREFIN_PLA')) + '"';
          end;
          vTobItemDest := nil;
          vTob := TOB.Create('#AFPLANNING', nil, -1);
          vTobItemDest := TOB.Create('#AFPLANNING', nil, -1);
          try
            vTob.LoadDetailFromSQL(vSt);
            if (vTob.Detail.Count > 0) then
              PgiBoxAf(TexteMessage[34], TexteMessage[59]); //Il y a d''autres ressources planifiées pour cette intervention, #13#10 pensez à également les déplacer.

            //---------------------------------------------------------------
            // C.B 23/05/2005
            // test si on doit faire suivre les autres items pour cette ressource
            // pour cette date et pour cette affaire
            //---------------------------------------------------------------
            // pas au point, regler le rafraichissement + les tests des dates
            //---------------------------------------------------------------
            {if (vTob.Detail.Count > 0) and
               (PGIAskAF(TexteMessage[28], '') = mrYes) then
            begin
              for i := 0 to vTob.Detail.Count - 1 do
              begin
                vTobItem := fTobsPlanningTobItems.findFirst(['APL_AFFAIRE','APL_NUMEROTACHE'], [ToItem.GetValue('APL_AFFAIRE'), ToItem.GetValue('APL_NUMEROTACHE')], False);
                vTobItem.putValue('APL_DATEDEBPLA', vTobItem.GetValue('APL_DATEDEBPLA') + (ToItem.GetValue('APL_DATEDEBPLA') - FromItem.GetValue('APL_DATEDEBPLA')));
                vTobItem.putValue('APL_DATEFINPLA', vTobItem.GetValue('APL_DATEFINPLA') + (ToItem.GetValue('APL_DATEFINPLA') - FromItem.GetValue('APL_DATEFINPLA')));
                vTobItem.putValue('APL_HEUREDEB_PLA', vTobItem.GetValue('APL_HEUREDEB_PLA') + (ToItem.GetValue('APL_HEUREDEB_PLA') - FromItem.GetValue('APL_HEUREDEB_PLA')));
                vTobItem.putValue('APL_HEUREFIN_PLA', vTobItem.GetValue('APL_HEUREFIN_PLA') + (ToItem.GetValue('APL_HEUREFIN_PLA') - FromItem.GetValue('APL_HEUREFIN_PLA')));
                MoveItem(nil, vTobItem, vBoCancel);
                //FComposantPlanning.InvalidateItem(vTobItem);
              end;
            end;
            }
          finally
            FreeAndNil(vTob);
            FreeAndNil(vTobItemDest);
          end;
        end;
      end;
  end;
end;

procedure TPlanningGA.Link(Sender: TObject; TobSource, TobDestination: TOB; Option: THPlanningOptionLink; var Cancel: Boolean);
var
  vStAffaire: string;
begin
  if TobSource.GetValue('APL_ETATLIGNE') = 'ABS' then
    cancel := True
  else
  begin
    cancel := False;
    case Option of
      polExtend:
        begin
          {if not JaiLeDroitRessource(TobSource.getValue('APL_RESSOURCE')) then
          begin
            PgiBoxAf(TexteMessage[24], TexteMessage[59]); //La ressource ne fait pas partie de votre équipe
            Cancel := True;
          end

          else if IsBloque(TobSource) then
          begin
            PGIBox(format(traduitGA(TexteMessage[18]), [TobSource.GetValue('APL_UTILISATEUR')]), TraduitGA(TexteMessage[59])); //Cette intervention est bloquée par %s.
            Cancel := True;
          end

          // C.B 06/04/2006
          // test si on est a cheval sur 2 semaines ou 2 mois
          else if GetParamSocSecur('SO_AFPARSEMAINE', True) and
             EstAChevalSemaine(TobDestination, fBoDemiJournee) then
          begin
            PGIBoxAF(TexteMessage[70], TexteMessage[59]); // Vous ne pouvez pas créer, déplacer, réduire ou étirer une intervention à cheval sur 2 semaines.
            cancel := True;
          end

          // C.B 06/04/2006
          // test si on est a cheval sur 2 semaines ou 2 mois
          else if GetParamSocSecur('SO_AFPARMOIS', True) and
             EstAChevalMois(TobDestination, fBoDemiJournee) then
          begin
            PGIBoxAF(TexteMessage[71], TexteMessage[59]); // Vous ne pouvez pas créer, déplacer, réduire ou étirer une intervention à cheval sur 2 semaines.
            cancel := True;
          end

          else if not AutoriserPlanif(TobDestination, fBoDemiJournee, False, vStAffaire, 'EXTEND') then
          begin
            Cancel := True;
          end

          else if not UpdateItem(TobSource, TobDestination, 'EXTEND') then
            cancel := True

          // C.B 31/10/2006
          // si on n'affiche pas les samedis et les dimanches et qu'on
          // est à cheval sur 2 semaines, on averti que les samedis et les
          // dimanches sont comptés
          else if (not GetParamSocSecur('SO_AFPLANNINGWE', true)) and
               EstAChevalSemaine(TobDestination, fBoDemiJournee) then
          begin
            PGIBoxAF(TexteMessage[20], TexteMessage[59]); // Attention, Les samedis et dimanches ne sont pas visibles, mais ils sont planifiés.
          end;}
        end;

      polReduce:
        begin
          {if not JaiLeDroitRessource(TobSource.getValue('APL_RESSOURCE')) then
          begin
            PgiBoxAf(TexteMessage[24], TexteMessage[59]); //La ressource ne fait pas partie de votre équipe
            Cancel := True;
          end

          else if IsBloque(TobSource) then
          begin
            PGIBox(format(traduitGA(TexteMessage[18]), [TobSource.GetValue('APL_UTILISATEUR')]), TraduitGA(TexteMessage[59])); //Cette intervention est bloquée par %s.
            Cancel := True;
          end

          // C.B 06/04/2006
          // test si on est a cheval sur 2 semaines ou 2 mois
          else if GetParamSocSecur('SO_AFPARSEMAINE', True) and
             EstAChevalSemaine(TobDestination, fBoDemiJournee) then
          begin
            PGIBoxAF(TexteMessage[70], TexteMessage[59]); // Vous ne pouvez pas créer, déplacer, réduire ou étirer une intervention à cheval sur 2 semaines.
            cancel := True;
          end

          // C.B 06/04/2006
          // test si on est a cheval sur 2 semaines ou 2 mois
          else if GetParamSocSecur('SO_AFPARMOIS', True) and
             EstAChevalMois(TobDestination, fBoDemiJournee) then
          begin
            PGIBoxAF(TexteMessage[71], TexteMessage[59]); // Vous ne pouvez pas créer, déplacer, réduire ou étirer une intervention à cheval sur 2 semaines.
            cancel := True;
          end

          else if not AutoriserPlanif(TobDestination, fBoDemiJournee, False, vStAffaire, 'REDUCE') then
          begin
            Cancel := True;
          end

          else if not UpdateItem(TobSource, TobDestination, 'REDUCE') then
            Cancel := True

          // C.B 31/10/2006
          // si on n'affiche pas les samedis et les dimanches et qu'on
          // est à cheval sur 2 semaines, on averti que les samedis et les
          // dimanches sont comptés
          else if (not GetParamSocSecur('SO_AFPLANNINGWE', true)) and
               EstAChevalSemaine(TobDestination, fBoDemiJournee) then
          begin
            PGIBoxAF(TexteMessage[20], TexteMessage[59]); // Attention, Les samedis et dimanches ne sont pas visibles, mais ils sont planifiés.
          end;}
        end;

      {      polMove :
              begin
                if not MoveItem(TobSource, TobDestination, False) then
                  cancel := True;
              end;
      }
    end;
    AffichageQuantites(TobSource);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 26/10/2004
Modifié le ... :
Description .. : GESTION DES ABSENCES
Suite ........ :
Mots clefs ... :
*****************************************************************}
procedure TPlanningGA.BeforeChange(const Item: TOB; const LaRessource: string; const LaDateDeDebut, LaDateDeFin: TDateTime; var Cancel: Boolean);
begin
  cancel := Item.GetValue('APL_ETATLIGNE') = 'ABS';
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 28/03/2002
Modifié le ... :
Description .. : valeur de création d'un item
Suite ........ :
Mots clefs ... :
*****************************************************************}
procedure TPlanningGA.InitItem(Sender: TObject; var Item: TOB; var Cancel: boolean);
begin
  cancel := false;
  if Item = nil then
    Item := TOB.Create('AFPLANNING', nil, -1);

  Item.putvalue('APL_TYPEPLA', 'PLA');

  // ajout des champs qui ne sont pas dans la table AFPLANNING
  if not Item.FieldExists('ATA_LIBELLETACHE1') then
  begin
    Item.AddChampSup('ATA_LIBELLETACHE1', true);
    Item.AddChampSup('ATA_QTEINITIALE', true);
    Item.AddChampSup('ARS_LIBELLE', true);
    Item.AddChampSup('ARS_LIBELLE2', true);
    Item.AddChampSup('T_LIBELLE', true);

    //C.B 21/09/2006
    Item.AddChampSup('T_PRENOM', true);
    Item.AddChampSup('T_ADRESSE1', true);
		Item.AddChampSup('T_ADRESSE2', true);
		Item.AddChampSup('T_ADRESSE3', true);
		Item.AddChampSup('T_CODEPOSTAL', true);
		Item.AddChampSup('T_VILLE', true);
    Item.AddChampSup('T_TELEPHONE', true);

    Item.AddChampSup('ATA_IDENTLIGNE', true);
    Item.AddChampSup('ATA_STATUTPLA', true);
    Item.AddChampSup('ATA_TYPEPLANIF', true);
    Item.AddChampSup('ATA_DESCRIPTIF', true);
    Item.AddChampSup('ATA_TERMINE', true);    
    Item.AddChampSup('ATA_NBPARTINSCRIT', true);
    Item.AddChampSup('ATA_NBPARTPREVU', true);
    Item.AddChampSup('ATA_TYPEADRESSE', true);
    Item.AddChampSup('ATA_NUMEROADRESSE', true);
    Item.AddChampSup('AFF_LIBELLE', true);
  end;
  Item.AddChampSup('I_HINT', False);

  // C.B 13/05/2005 spécifique cegid
  if VH_GC.GCIfDefCEGID then
    Item.AddChampSupValeur('I_ICONE', cInIconeNORMALE, False);

  //Item.AddChampSupValeur('I_COLOR', '', False);
  Item.AddChampSup('I_CODE', False);
  Item.AddChampSupValeur('CHAMPLIBELLE', '', False);
  Item.AddChampSupValeur('CODEPOSTAL', '', False);
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 28/03/2002
Modifié le ... :
Description .. : création d'une ligne de planning et affichage à
Suite ........ : l'écran
...............: création directe ou avec saisie des données
Mots clefs ... :
*****************************************************************}
procedure TPlanningGA.CreateItem(Sender: TObject; Item: TOB; var Cancel: boolean);
var
  vStRetour: string;
  //vStCodePostal : string;
  vStAction: string;
  vStResLib: string;
  vSt: string;
begin
  cancel := false;

  if (fAction = TaConsult) then
    vStAction := 'ACTION=CONSULTATION'
  else
    vStAction := 'ACTION=CREATION';

  if fBoDemiJournee then
    vStAction := vStAction + ';DEMIJOURNEE';

  if (fStAffaire <> '') then
  begin
    fTobCourante.PutValue('APL_AFFAIRE', fStAffaire);
    fTobCourante.PutValue('APL_TIERS', fStTiers);
    Item.PutValue('APL_AFFAIRE', fStAffaire);
    Item.PutValue('APL_TIERS', fStTiers);
  end;

  if fStRes <> '' then
  begin
    fTobCourante.PutValue('APL_RESSOURCE', fStRes);
    Item.PutValue('APL_RESSOURCE', fStRes);
  end;

  // C.B 06/04/2006
  // test si on est a cheval sur 2 semaines ou 2 mois
  {if GetParamSocSecur('SO_AFPARSEMAINE', True) and
     EstAChevalSemaine(Item, fBoDemiJournee) then
  begin
    PGIBoxAF(TexteMessage[70], TexteMessage[59]); // Vous ne pouvez pas créer, déplacer, réduire ou étirer une intervention à cheval sur 2 semaines.
    cancel := True;
  end

  // C.B 06/04/2006
  // test si on est a cheval sur 2 semaines ou 2 mois
  else if GetParamSocSecur('SO_AFPARMOIS', True) and
     EstAChevalMois(Item, fBoDemiJournee) then
  begin
    PGIBoxAF(TexteMessage[71], TexteMessage[59]); // Vous ne pouvez pas créer, déplacer, réduire ou étirer une intervention à cheval sur 2 semaines.
    cancel := True;
  end

  // C.B 23/02/2006
  // tester si l'affaire n'est pas terminée
  else if (not cancel) and (fStAffaire <> '') and AffaireTermine(fStAffaire) then
  begin
    PGIBoxAF(TexteMessage[68], TexteMessage[59]); // 'Vous ne pouvez plus planifier d''interventions pour une affaire terminée.'
    Cancel := True;
  end

  else if (not cancel) and (fStTache <> '') and (item.GetString('ATA_TERMINE') = 'X') then
  begin
    PGIBoxAF(TexteMessage[69], TexteMessage[59]); // 'Vous ne pouvez plus planifier d''interventions pour une tâche terminée.'
    Cancel := True;
  end

  // creation directe
  else if (not cancel) and (fInModeCreation = cInCreationDirecte) or
    ((fInModeCreation = cInCreationNonProductive) and (fStRes <> '')) or
    (fInModeCreation = cInCreationSelectionEtRes) then
  begin

    // C.B 03/10/2005 on vérifie qu'on a le droit de planifier cette ressource
    if not JaiLeDroitRessource(fStRes) then
    begin
      PGIBoxAF(TexteMessage[24], TexteMessage[59]); //La ressource ne fait pas partie de votre équipe
      cancel := True;
    end
    else
    begin
      vStRetour := CreationDirecte(Item);
      if vStRetour <> '' then
      begin
        // ne pas faire le decode clé pour les non productives ...
        if (fInModeCreation = cInCreationDirecte) then
          DecodeCle(item.getvalue('I_CODE'));
        DisplayHint(Item, vStRetour);
        FComposantPlanning.InvalidateItem(Item);
      end
      else
        cancel := True;
    end;
  end

    // AVEC SELECTION MAIS SANS RESSOURCE
  else if (not cancel) and (fInModeCreation = cInCreationAvecSelection) and
    AutoriserPlanif(Item, fBoDemiJournee, False, vSt, 'CREATE') then
  begin
    // choix d'une autre affaire que l'affaire sélectionnée
    // on passe l'affaire mais on ne passe pas l'article
    // on vide la sélection

    // si la tache a été sélectionnée, mais qu'il y a plusieurs ressource  ou aucune ressource
    // passer la tache en paramètre
    //C.B vStAction en premier pour gerer le code article connu en création
    if (fStCodeArtSel <> '') then
      vStRetour := AFLanceFicheAFPlanning('', vStAction + ';' + 'CODEARTICLE=' + fStCodeArtSel + ';' + 'TACHE=' + fStTacheSel)
    else
      vStRetour := AFLanceFicheAFPlanning('', vStAction);

    if (vStRetour = 'CANCEL') or (vStRetour = '') then
      cancel := True;
  end

    // SANS SELECTION
  else if (not cancel) and (fInModeCreation = cInCreationSansSelection) and
    AutoriserPlanif(Item, fBoDemiJournee, False, vSt, 'CREATE') then
  begin
    // dans le planning 3, on est sans sélection mais on connait le codeArticle
    //C.B vStAction en premier pour gerer le code article connu en création
    if fStCodeArt <> '' then
      vStRetour := AFLanceFicheAFPlanning('', vStAction + ';' + 'CODEARTICLE=' + fStCodeArt + ';' + 'TACHE=' + fStTache)
    else
      vStRetour := AFLanceFicheAFPlanning('', vStAction);

    if (vStRetour = 'CANCEL') or (vStRetour = '') then
      cancel := True;
  end

    // ANNULATION
  else
    Cancel := True;}

  // RAFRAICHISSEMENT SAUF POUR CREATION DIRECTE
  if (not cancel) and ((fInModeCreation = cInCreationAvecSelection) or
    (fInModeCreation = cInCreationSansSelection)) then
  begin
    // C.B 26/09/2006
    Item.PutValue('APL_ETATLIGNE', GetStatutItem(Item.GetString('ATA_STATUTPLA')));
    if fStAffaire = '' then fStAffaire := fTobCourante.GetString('APL_AFFAIRE');
    if fStTiers = '' then fStTiers := fTobCourante.GetString('APL_TIERS');
    if fStRes = '' then fStRes := fTobCourante.GetString('APL_RESSOURCE');
    if fStTache = '' then fStTache := fTobCourante.GetString('APL_NUMEROTACHE');

    Item.dupliquer(fTobCourante, True, True, True);
    Item.loadDB(False);
    LoadComplementItem(Item);
    LoadLibelleRessource(Item);
    DisplayHint(Item, vStRetour);
    FComposantPlanning.InvalidateItem(Item);
  end;

  // POUR TOUS
  if not cancel then
  begin
    // creation des ressources ajoutées
    CreateTacheRessource(Item, True);

    // on passe de avec sélection en créationDirecte
    {if (fInModeCreation = cInCreationAvecSelection) and
      UniqueRes(fStAffaireSel, fStTacheSel, fStResSel, vStResLib) and
      getParamSocSecur('SO_AFAFFECTRES', False) then
    begin
      FFichePlanning.LA_SELECTION.Caption := FFichePlanning.LA_SELECTION.Caption + TraduitGa('Ressource : ') + fStResSel + '  ' + vStResLib + '     ';
      fInModeCreation := cInCreationDirecte;
    end;}

    // C.B 08/09/2005
    // affichage du commentaire
    AffichageCommentaire(Item);
    AffichageQuantites(Item);

    // C.B 31/10/2006
    // si on n'affiche pas les samedis et les dimanches et qu'on
    // est à cheval sur 2 semaines, on averti que les samedis et les
    // dimanches sont comptés
    {if (not GetParamSocSecur('SO_AFPLANNINGWE', true)) and
       EstAChevalSemaine(Item, fBoDemiJournee) then
    begin
      PGIBoxAF(TexteMessage[20], TexteMessage[59]); // Attention, Les samedis et dimanches ne sont pas visibles, mais ils sont planifiés.
    end;}
    
  end;

  if (fInModeCreation = cInCreationSelectionEtRes) then
  begin
    fInModeCreation := cInCreationAvecSelection;
    fStRes := '';
    fStResSel := '';
  end;

  DetruitMaTobInterne('AFFAIRECOURANTE');
end;
                        
// recherche du libelle de la ressource pour l'afficher
{procedure TPlanningGA.LoadLibelleRessource(pItem: Tob);
var
  vQr : Tquery;
begin
  if FComposantPlanning.ChampLibelle = 'ARS_LIBELLE' then
  begin
    vQr := nil;
    try
      vQr := OpenSQL('SELECT ARS_LIBELLE FROM RESSOURCE WHERE ARS_RESSOURCE ="' + pItem.GetValue('APL_RESSOURCE') + '"', True);
      if not vQr.Eof then
        pItem.putValue('ARS_LIBELLE', vQr.Fields[0].AsString)
      else
        pItem.putValue('ARS_LIBELLE', '');
    finally
      Ferme(vQr);
    end;
  end;
end;
}

procedure TPlanningGA.LoadLibelleRessource(pItem: Tob);
var
  vQr: Tquery;
begin
  vQr := nil;
  try
    vQr := OpenSQL('SELECT ARS_LIBELLE, ARS_LIBELLE2 FROM RESSOURCE WHERE ARS_RESSOURCE ="' + pItem.GetValue('APL_RESSOURCE') + '"', True,-1,'',true);
    if not vQr.Eof then
    begin
      pItem.putValue('ARS_LIBELLE', vQr.Fields[0].AsString);
      pItem.putValue('ARS_LIBELLE2', vQr.Fields[1].AsString); 
      //      pItem.putValue('CHAMPLIBELLE', vQr.Fields[0].AsString);
    end
    else
    begin
      pItem.putValue('ARS_LIBELLE', '');
      pItem.putValue('ARS_LIBELLE2', '');
      //      pItem.putValue('CHAMPLIBELLE', '');
    end;
  finally
    Ferme(vQr);
  end;
end;

procedure TPlanningGA.LoadComplementItem(pItem: Tob);
var
  vQr: Tquery;
  vSt: string;

begin
  vQr := nil;
  vSt := 'SELECT ATA_LIBELLETACHE1, ATA_QTEINITIALE, ATA_IDENTLIGNE, T_LIBELLE, ';
  //C.B 21/09/2006 
  vSt := vSt + ' T_PRENOM, T_ADRESSE1, T_ADRESSE2, T_ADRESSE3, T_CODEPOSTAL, T_VILLE, T_TELEPHONE, ';
  vSt := vSt + ' ATA_NBPARTINSCRIT, ATA_NBPARTPREVU, ATA_TYPEADRESSE, ATA_NUMEROADRESSE, ';
  vSt := vSt + ' ATA_STATUTPLA ';
  vSt := vSt + ' FROM TACHE, TIERS WHERE ATA_AFFAIRE = "' + pItem.GetString('APL_AFFAIRE') + '"';
  vSt := vSt + ' AND ATA_NUMEROTACHE = ' + pItem.GetString('APL_NUMEROTACHE');
  vSt := vSt + ' AND ATA_TIERS = T_TIERS';

  try
    vQr := OpenSQL(vSt, True,-1,'',true);
    if not vQr.Eof then
    begin
      pItem.putValue('ATA_LIBELLETACHE1', vQr.Fields[0].AsString);
      pItem.putValue('ATA_QTEINITIALE', vQr.Fields[1].AsString);
      pItem.putValue('ATA_IDENTLIGNE', vQr.Fields[2].AsString);
      //C.B 21/09/2006
      pItem.putValue('T_LIBELLE', vQr.Fields[3].AsString);
      pItem.putValue('T_PRENOM', vQr.Fields[4].AsString);
      pItem.putValue('T_ADRESSE1', vQr.Fields[5].AsString);
      pItem.putValue('T_ADRESSE2', vQr.Fields[6].AsString);
      pItem.putValue('T_ADRESSE3', vQr.Fields[7].AsString);
      pItem.putValue('T_CODEPOSTAL', vQr.Fields[8].AsString);
      pItem.putValue('T_VILLE', vQr.Fields[9].AsString);
      pItem.putValue('T_TELEPHONE', vQr.Fields[10].AsString);
      
      pItem.putValue('ATA_NBPARTINSCRIT', vQr.Fields[11].AsInteger);
      pItem.putValue('ATA_NBPARTPREVU', vQr.Fields[12].AsInteger);
      pItem.putValue('ATA_TYPEADRESSE', vQr.Fields[13].AsString);
      pItem.putValue('ATA_NUMEROADRESSE', vQr.Fields[14].AsInteger);
      pItem.putValue('ATA_STATUTPLA', vQr.Fields[15].AsString);
    end
    else
    begin
      pItem.putValue('ATA_LIBELLETACHE1', '');
      pItem.putValue('ATA_QTEINITIALE', '');
      pItem.putValue('ATA_IDENTLIGNE', '');

      //C.B 21/09/2006
      pItem.putValue('T_PRENOM', '');
      pItem.putValue('T_ADRESSE1', '');
      pItem.putValue('T_ADRESSE2', '');
      pItem.putValue('T_ADRESSE3', '');
      pItem.putValue('T_CODEPOSTAL', '');
      pItem.putValue('T_VILLE', '');
      pItem.putValue('T_TELEPHONE', '');
 
      pItem.putValue('ATA_NBPARTINSCRIT', 0);
      pItem.putValue('ATA_NBPARTPREVU', 0);
      pItem.putValue('ATA_TYPEADRESSE', '');
      pItem.putValue('ATA_NUMEROADRESSE', 0);
      pItem.putValue('ATA_STATUTPLA', '');
    end;
  finally
    Ferme(vQr);
  end;
end;

function TPlanningGA.CreationDirecte(pItem: Tob): string;
var
  vTOBAffaires: Tob;
  vTOBArticles: Tob;
  vTobTache: Tob;
  vSt: string;
  //  vStCodePostal: string;
  vAFOAssistants: TAFO_Ressources;
  vBoOk: Boolean;
 vStRetour : string;
 //redraw : Boolean;

begin
  result := '';
  vSt := 'SELECT ATA_FONCTION, ATA_TYPEARTICLE, ATA_ARTICLE,';
  vSt := vSt + ' ATA_CODEARTICLE, ATA_AFFAIRE0, ATA_AFFAIRE1,';
  vSt := vSt + ' ATA_AFFAIRE2, ATA_AFFAIRE3, ATA_AVENANT,';
  vSt := vSt + ' ATA_UNITETEMPS, ATA_DEVISE, ATA_FONCTION,';
  vSt := vSt + ' ATA_LIBRETACHE1, ATA_LIBRETACHE2, ATA_LIBRETACHE3,';
  vSt := vSt + ' ATA_LIBRETACHE4, ATA_LIBRETACHE5, ATA_LIBRETACHE6,';
  vSt := vSt + ' ATA_LIBRETACHE7, ATA_LIBRETACHE8, ATA_LIBRETACHE9, ATA_LIBRETACHEA,';
  vSt := vSt + ' ATA_BOOLLIBRE1, ATA_BOOLLIBRE2, ATA_BOOLLIBRE3,';
  vSt := vSt + ' ATA_DATELIBRE1, ATA_DATELIBRE2, ATA_DATELIBRE3,';
  vSt := vSt + ' ATA_CHARLIBRE1, ATA_CHARLIBRE2, ATA_CHARLIBRE3,';
  vSt := vSt + ' ATA_VALLIBRE1, ATA_VALLIBRE2, ATA_VALLIBRE3, T_LIBELLE, ATA_QTEINITIALE, ';
  vSt := vSt + ' ATA_ACTIVITEREPRIS, ATA_LIBELLETACHE1, ATA_IDENTLIGNE, ATA_DESCRIPTIF, ATA_TERMINE, ATA_TYPEPLANIF, ';
  vSt := vSt + ' ATA_NBPARTINSCRIT, ATA_NBPARTPREVU, ATA_TYPEADRESSE, ATA_NUMEROADRESSE, ';

  //C.B 21/09/2006
  //optimisation
  // on charge les données du tiers directement dans le loadItem
	vSt := vSt + ' T_PRENOM, T_ADRESSE1, T_ADRESSE2, T_ADRESSE3, T_CODEPOSTAL, T_VILLE, T_TELEPHONE, ';
  vSt := vSt + ' ATA_STATUTPLA ';
  
  vSt := vSt + ' FROM TACHE, TIERS ';
  vSt := vSt + ' WHERE ATA_AFFAIRE = "' + fStAffaireSel + '"';
  vSt := vSt + ' AND ATA_NUMEROTACHE = ' + fStTacheSel;
  vSt := vSt + ' AND ATA_TIERS = T_TIERS';

  vTobtache := TOB.Create('#AFTACHE', nil, -1);
  vTobtache.LoadDetailFromSQL(vSt);

  if vTobtache.detail.count = 0 then
    PGIBoxAF(TexteMessage[31], TexteMessage[59])

  // C.B 15/06/2006
	//Vous ne pouvez pas planifier des tâches en heures dans ce planning, l''unité minimum est la demi-journée.
  else if vTobtache.detail[0].GetString('ATA_UNITETEMPS') = 'H' then
    PGIBoxAF(TexteMessage[76], TexteMessage[59])
         
  else
  try
    pItem.putValue('APL_TYPEPLA', 'PLA');
    pItem.putValue('APL_AFFAIRE', fStAffaireSel);
    //pItem.putValue('APL_NUMEROLIGNE', GetNumLignePlanning(fStAffaireSel));
    pItem.putValue('APL_TYPELIGNEPLA', 'TAC');
    pItem.putValue('APL_NUMEROTACHE', fStTacheSel);
    if fInModeCreation = cInCreationNonProductive then // Journée non productive - affaire et tâche unique
      pItem.putValue('APL_RESSOURCE', fStRes)
    else
      pItem.putValue('APL_RESSOURCE', fStResSel);

    loadLibelleRessource(pItem);
    pItem.putValue('APL_FONCTION', vTobTache.detail[0].GetValue('ATA_FONCTION'));
    pItem.putValue('APL_ARTICLE', vTobTache.detail[0].GetValue('ATA_ARTICLE'));
    pItem.putValue('APL_TYPEARTICLE', vTobTache.detail[0].GetValue('ATA_TYPEARTICLE'));
    pItem.putValue('APL_CODEARTICLE', vTobTache.detail[0].GetValue('ATA_CODEARTICLE'));
    pItem.putValue('APL_CREATEUR', v_pgi.user);
    pItem.putValue('APL_DATECREATION', now);
    pItem.putValue('APL_UTILISATEUR', v_pgi.user);
    pItem.putValue('APL_DATEMODIF', now);
    pItem.putValue('APL_TIERS', fStTiersSel);

    // C.B 26/09/2006
    // suppression du spécifique cegid
    pItem.putValue('APL_ETATLIGNE', GetStatutItem(vTobTache.detail[0].GetValue('ATA_STATUTPLA')));
    pItem.putValue('APL_MOTIFETAT', '');
    pItem.putValue('APL_LIGNEGENEREE', '-');
    pItem.putValue('APL_ACTIVITEGENERE', '-');
    pItem.putValue('APL_FOLIO', '0');
    pItem.putValue('APL_AFFAIRE0', vTobTache.detail[0].GetValue('ATA_AFFAIRE0'));
    pItem.putValue('APL_AFFAIRE1', vTobTache.detail[0].GetValue('ATA_AFFAIRE1'));
    pItem.putValue('APL_AFFAIRE2', vTobTache.detail[0].GetValue('ATA_AFFAIRE2'));
    pItem.putValue('APL_AFFAIRE3', vTobTache.detail[0].GetValue('ATA_AFFAIRE3'));
    pItem.putValue('APL_AVENANT', vTobTache.detail[0].GetValue('ATA_AVENANT'));
    pItem.putValue('APL_ACTIVITEEFFECT', '-');
    pItem.putValue('APL_UNITETEMPS', vTobTache.detail[0].GetValue('ATA_UNITETEMPS'));

    //ModifQte(nil, pItem, fBoDemiJournee);

      // C.B 14/06/2005
      // on recharge les valeurs libres sauf pour cegid ou on ne récupere rien
      //C.B 02/02/2006
      // on charge certaines valeurs pour cegid, mais pas toutes ...
      if VH_GC.GCIfDefCEGID then
      begin
        pItem.putValue('APL_LIBRETACHE1', '-'); // RI/FP Reçu :  à forcer à « - »
        pItem.putValue('APL_LIBRETACHE2', '-'); // Doublons :  à forcer à « - »
        pItem.putValue('APL_LIBRETACHE3', vTobTache.detail[0].GetValue('ATA_LIBRETACHE3'));
        pItem.putValue('APL_LIBRETACHE4', vTobTache.detail[0].GetValue('ATA_LIBRETACHE4'));
        pItem.putValue('APL_LIBRETACHE5', vTobTache.detail[0].GetValue('ATA_LIBRETACHE5'));
        pItem.putValue('APL_LIBRETACHE6', vTobTache.detail[0].GetValue('ATA_LIBRETACHE6'));
        pItem.putValue('APL_LIBRETACHE7', vTobTache.detail[0].GetValue('ATA_LIBRETACHE7'));
        pItem.putValue('APL_LIBRETACHE8', vTobTache.detail[0].GetValue('ATA_LIBRETACHE8'));
        pItem.putValue('APL_LIBRETACHE9', vTobTache.detail[0].GetValue('ATA_LIBRETACHE9'));
        pItem.putValue('APL_LIBRETACHEA', vTobTache.detail[0].GetValue('ATA_LIBRETACHEA'));
        pItem.putValue('APL_BOOLLIBRE1', vTobTache.detail[0].GetValue('ATA_BOOLLIBRE1'));
        pItem.putValue('APL_BOOLLIBRE2', vTobTache.detail[0].GetValue('ATA_BOOLLIBRE2'));
        pItem.putValue('APL_BOOLLIBRE3', vTobTache.detail[0].GetValue('ATA_BOOLLIBRE3'));
        pItem.putValue('APL_DATELIBRE1', vTobTache.detail[0].GetValue('ATA_DATELIBRE1'));
        pItem.putValue('APL_DATELIBRE2', vTobTache.detail[0].GetValue('ATA_DATELIBRE2'));
        pItem.putValue('APL_DATELIBRE3', vTobTache.detail[0].GetValue('ATA_DATELIBRE3'));
        pItem.putValue('APL_CHARLIBRE1', ''); // n° de RI
        pItem.putValue('APL_CHARLIBRE2', ''); //utilisateur qui a mis le n° de RI
        pItem.putValue('APL_CHARLIBRE3', vTobTache.detail[0].GetValue('ATA_CHARLIBRE3')); // C.B 09/02/06 compétence qui est redescendue de la tache
        pItem.putValue('APL_VALLIBRE1', vTobTache.detail[0].GetValue('ATA_VALLIBRE1'));
        pItem.putValue('APL_VALLIBRE2', vTobTache.detail[0].GetValue('ATA_VALLIBRE2'));
        pItem.putValue('APL_VALLIBRE3', vTobTache.detail[0].GetValue('ATA_VALLIBRE3'));
      end
      else
      begin
        pItem.putValue('APL_LIBRETACHE1', vTobTache.detail[0].GetValue('ATA_LIBRETACHE1'));
        pItem.putValue('APL_LIBRETACHE2', vTobTache.detail[0].GetValue('ATA_LIBRETACHE2'));
        pItem.putValue('APL_LIBRETACHE3', vTobTache.detail[0].GetValue('ATA_LIBRETACHE3'));
        pItem.putValue('APL_LIBRETACHE4', vTobTache.detail[0].GetValue('ATA_LIBRETACHE4'));
        pItem.putValue('APL_LIBRETACHE5', vTobTache.detail[0].GetValue('ATA_LIBRETACHE5'));
        pItem.putValue('APL_LIBRETACHE6', vTobTache.detail[0].GetValue('ATA_LIBRETACHE6'));
        pItem.putValue('APL_LIBRETACHE7', vTobTache.detail[0].GetValue('ATA_LIBRETACHE7'));
        pItem.putValue('APL_LIBRETACHE8', vTobTache.detail[0].GetValue('ATA_LIBRETACHE8'));
        pItem.putValue('APL_LIBRETACHE9', vTobTache.detail[0].GetValue('ATA_LIBRETACHE9'));
        pItem.putValue('APL_LIBRETACHEA', vTobTache.detail[0].GetValue('ATA_LIBRETACHEA'));
        pItem.putValue('APL_BOOLLIBRE1', vTobTache.detail[0].GetValue('ATA_BOOLLIBRE1'));
        pItem.putValue('APL_BOOLLIBRE2', vTobTache.detail[0].GetValue('ATA_BOOLLIBRE2'));
        pItem.putValue('APL_BOOLLIBRE3', vTobTache.detail[0].GetValue('ATA_BOOLLIBRE3'));
        pItem.putValue('APL_DATELIBRE1', vTobTache.detail[0].GetValue('ATA_DATELIBRE1'));
        pItem.putValue('APL_DATELIBRE2', vTobTache.detail[0].GetValue('ATA_DATELIBRE2'));
        pItem.putValue('APL_DATELIBRE3', vTobTache.detail[0].GetValue('ATA_DATELIBRE3'));
        pItem.putValue('APL_CHARLIBRE1', vTobTache.detail[0].GetValue('ATA_CHARLIBRE1'));
        pItem.putValue('APL_CHARLIBRE2', vTobTache.detail[0].GetValue('ATA_CHARLIBRE2'));
        pItem.putValue('APL_CHARLIBRE3', vTobTache.detail[0].GetValue('ATA_CHARLIBRE3'));
        pItem.putValue('APL_VALLIBRE1', vTobTache.detail[0].GetValue('ATA_VALLIBRE1'));
        pItem.putValue('APL_VALLIBRE2', vTobTache.detail[0].GetValue('ATA_VALLIBRE2'));
        pItem.putValue('APL_VALLIBRE3', vTobTache.detail[0].GetValue('ATA_VALLIBRE3'));
      end;

    pItem.putValue('APL_DEVISE', vTobTache.detail[0].GetValue('ATA_DEVISE'));
    pItem.putValue('APL_ACTIVITEREPRIS', vTobTache.detail[0].GetValue('ATA_ACTIVITEREPRIS'));
    pItem.putValue('APL_BLOQUE', '-');
    pItem.putValue('APL_ESTFACTURE', '-');
    pItem.putValue('APL_LIBELLEPLA', vTobTache.detail[0].GetValue('ATA_LIBELLETACHE1'));
    pItem.putValue('ATA_DESCRIPTIF', vTobTache.detail[0].GetValue('ATA_DESCRIPTIF'));
    pItem.putValue('ATA_TERMINE', vTobTache.detail[0].GetValue('ATA_TERMINE'));
    pItem.putValue('ATA_IDENTLIGNE', vTobTache.detail[0].GetValue('ATA_IDENTLIGNE'));
    pItem.putValue('ATA_LIBELLETACHE1', vTobTache.detail[0].GetValue('ATA_LIBELLETACHE1'));
    pItem.putValue('ATA_QTEINITIALE', vTobTache.detail[0].GetValue('ATA_QTEINITIALE'));
    pItem.putValue('ATA_TYPEPLANIF', vTobTache.detail[0].GetValue('ATA_TYPEPLANIF'));
    pItem.putValue('ATA_TYPEADRESSE', vTobTache.detail[0].GetValue('ATA_TYPEADRESSE'));
    pItem.putValue('ATA_NUMEROADRESSE', vTobTache.detail[0].GetValue('ATA_NUMEROADRESSE'));

    //C.B 21/09/2006
    pItem.putValue('T_LIBELLE', vTobTache.detail[0].GetValue('T_LIBELLE'));
    pItem.putValue('T_PRENOM', vTobTache.detail[0].GetValue('T_PRENOM'));
    pItem.putValue('T_ADRESSE1', vTobTache.detail[0].GetValue('T_ADRESSE1'));
    pItem.putValue('T_ADRESSE2', vTobTache.detail[0].GetValue('T_ADRESSE2'));
    pItem.putValue('T_ADRESSE3', vTobTache.detail[0].GetValue('T_ADRESSE3'));
    pItem.putValue('T_CODEPOSTAL', vTobTache.detail[0].GetValue('T_CODEPOSTAL'));
    pItem.putValue('T_VILLE', vTobTache.detail[0].GetValue('T_VILLE'));
    pItem.putValue('T_TELEPHONE', vTobTache.detail[0].GetValue('T_TELEPHONE'));

    {if AutoriserPlanif(pItem, fBoDemiJournee, False, vSt, 'CREATE') then
    begin
      vBoOk := True;
      if ((fInModeCreation = cInCreationDirecte) or
        (fInModeCreation = cInCreationSelectionEtRes)) and
        PlusDeJours(nil, pItem, fBoDemiJournee) then
      begin
        vBoOk := (PGIAskAF(TexteMessage[13], '') = mrYes);
        if vBoOk then vBoOk := (PGIAskAF(TexteMessage[14], '') = mrYes);
        if vBoOk and VH_GC.GCIfDefCEGID then RevalorisationCegid(nil, pItem, 'CREATE');
      end;

      if vBoOk then
      begin
        vAFOAssistants := TAFO_Ressources.Create;
        vTOBAffaires := TOB.Create('Les Affaires', nil, -1);
        vTOBArticles := TOB.Create('les Articles', nil, -1);
        try
          ValorisationPlanning(pItem, 'APL', vAFOAssistants, vTOBAffaires, vTobArticles, True, True);

          pItem.InsertDB(nil);
          UpdatePlanifieTache(fStAffaireSel, fStTacheSel, vTobTache.detail[0].GetValue('ATA_UNITETEMPS'));
          result := DisplayHint(pItem, '', vTobTache.detail[0]);
          RafraichirItemPartiel(pItem, vStRetour, 'ARTICLE');
        finally
          vTOBAffaires.Free;
          vTobArticles.Free;
          vAFOAssistants.Free;
        end;
      end;
    end;}

  finally
    vTobTache.Free;
  end;
end;

procedure TPlanningGA.ClearAffaireSelect;
begin
  fStAffaireSel := '';
  fStTacheSel := '';
  fStCodeArtSel := '';
  fStTypeArtSel := '';
  fStArticleSel := '';
  fInModeCreation := cInCreationSansSelection;
end;

procedure TPlanningGA.MoveItem(Sender: TObject; Item: TOB; var Cancel: boolean);
var
  vSt	: string;
  vStResSource 	: String;

begin

  // C.B 05/06/2006
  // correction du déplacement des ressources
  // C.B 28/11/2006
  // attention, cas particulier dans l'accès par affaire, concervé les ressources
	vStResSource := Item.getValue('APL_RESSOURCE');
  DecodeCle(Item.GetValue('I_CODE'));
  if fStRes = '' then
    Item.PutValue('APL_RESSOURCE', vStResSource)
  else
    Item.PutValue('APL_RESSOURCE', fStRes);

  {if not JaiLeDroitRessource(vStResSource) then
  begin
    PgiBoxAf(Format(traduitGA(TexteMessage[23]), [vStResSource]), TexteMessage[59]); //La ressource %s ne fait pas partie de votre équipe
    cancel := true;
  end

  else if not JaiLeDroitRessource(Item.getValue('APL_RESSOURCE')) then
  begin
    PgiBoxAf(Format(traduitGA(TexteMessage[23]), [Item.getValue('APL_RESSOURCE')]), TexteMessage[59]); //La ressource %s ne fait pas partie de votre équipe
    cancel := true;
  end

  else if IsBloque(Item) then
  begin
    PGIBox(format(traduitGA(TexteMessage[18]), [Item.GetValue('APL_UTILISATEUR')]), TraduitGA(TexteMessage[59])); //Cette intervention est bloquée par %s.
    cancel := true;
  end

  // C.B 06/04/2006
  // test si on est a cheval sur 2 semaines ou 2 mois
  else if GetParamSocSecur('SO_AFPARSEMAINE', True) and
     EstAChevalSemaine(Item, fBoDemiJournee) then
  begin
    PGIBoxAF(TexteMessage[70], TexteMessage[59]); // Vous ne pouvez pas créer, déplacer, réduire ou étirer une intervention à cheval sur 2 semaines.
    cancel := True;
  end

  // C.B 06/04/2006
  // test si on est a cheval sur 2 semaines ou 2 mois
  else if GetParamSocSecur('SO_AFPARMOIS', True) and
     EstAChevalMois(Item, fBoDemiJournee) then
  begin
    PGIBoxAF(TexteMessage[71], TexteMessage[59]); // Vous ne pouvez pas créer, déplacer, réduire ou étirer une intervention à cheval sur 2 semaines.
    cancel := True;
  end

  else
    cancel := not AutoriserPlanif(Item, fBoDemiJournee, False, vSt, 'MOVE');

  if (not cancel) and (not UpdateItem(nil, Item, 'MOVE')) then
    cancel := true;

  // C.B 31/10/2006
  // si on n'affiche pas les samedis et les dimanches et qu'on
  // est à cheval sur 2 semaines, on averti que les samedis et les
  // dimanches sont comptés
  if (not cancel) and
     (not GetParamSocSecur('SO_AFPLANNINGWE', true)) and
     EstAChevalSemaine(Item, fBoDemiJournee) then
  begin
    PGIBoxAF(TexteMessage[20], TexteMessage[59]); // Attention, Les samedis et dimanches ne sont pas visibles, mais ils sont planifiés.
  end;}

end;

procedure TPlanningGA.ModifyItem(Sender: TObject; Item: TOB; var Cancel: boolean);
begin
  if item <> nil then
  begin
    DetruitMaTobInterne('AFFAIRECOURANTE');
    fTobCouranteMere := MaTobInterne('AFFAIRECOURANTE');
    fTobCourante := TOB.Create('AFFAIRECOURANTE', fTobCouranteMere, -1);
    fTobCourante.Dupliquer(Item, TRUE, TRUE, TRUE);

    if fAction = taConsult then
      fStAction := 'ACTION=CONSULTATION'
    else
      {if (not EtatModifAutorisee(Item.GetValue('APL_ETATLIGNE'), False, False)) or (IsBloque(Item)) then
      fStAction := 'ACTION=CONSULTATION'
      else}
         begin
         fStAction := 'ACTION=MODIFICATION';
         // C.B 08/09/2005
         // affichage du commentaire
         AffichageCommentaire(Item);
         AffichageQuantites(Item);
         end;
    if fBoDemiJournee then
      fStAction := fStAction + ';DEMIJOURNEE';
  end;
end;

function TPlanningGA.UpdateAfPlanning(ToItem : Tob; var pStRetour : String) : Boolean;
var
  vStQr : String;

begin
  {result := false;
  vStQR := 'UPDATE AFPLANNING SET APL_DATEDEBPLA = "' + UsDateTime(ToItem.GetValue('APL_DATEDEBPLA'));
  vStQR := vStQR + '", APL_DATEFINPLA = "' + UsDateTime(ToItem.GetValue('APL_DATEFINPLA'));
  vStQR := vStQR + '", APL_HEUREDEB_PLA = "' + UsTime(ToItem.GetValue('APL_HEUREDEB_PLA'));
  vStQR := vStQR + '", APL_HEUREFIN_PLA = "' + UsTime(ToItem.GetValue('APL_HEUREFIN_PLA'));
  vStQR := vStQR + '", APL_DATEDEBREAL = "' + UsDateTime(ToItem.GetValue('APL_DATEDEBREAL'));
  vStQR := vStQR + '", APL_DATEFINREAL = "' + UsDateTime(ToItem.GetValue('APL_DATEFINREAL'));
  vStQR := vStQR + '", APL_HEUREDEB_REAL = "' + UsTime(ToItem.GetValue('APL_HEUREDEB_REAL'));
  vStQR := vStQR + '", APL_HEUREFIN_REAL = "' + UsTime(ToItem.GetValue('APL_HEUREFIN_REAL'));

  vStQR := vStQR + '", APL_DATEMODIF = "' + usTime(now) ;
  vStQR := vStQR + '", APL_UTILISATEUR = "' + V_PGI.User;

  vStQR := vStQR + '", APL_RESSOURCE = "' + fStRes;
  vStQR := vStQR + '", APL_ETATLIGNE = "' + ToItem.GetValue('APL_ETATLIGNE');
  vStQR := vStQR + '", APL_FONCTION = "' + ToItem.GetValue('APL_FONCTION');
  vStQR := vStQR + '", APL_QTEPLANIFIEE = ' + variantToSql(ToItem.GetValue('APL_QTEPLANIFIEE'));
  vStQR := vStQR + ', APL_QTEREALISE = ' + variantToSql(ToItem.GetValue('APL_QTEREALISE'));

  vStQR := vStQR + ', APL_QTEPLANIFUREF = ' + variantToSql(AFConversionUnite(ToItem.GetValue('APL_UNITETEMPS'), getParamSocSecur('SO_AFMESUREACTIVITE',''),
    valeur(ToItem.GetValue('APL_QTEPLANIFIEE'))));
  vStQR := vStQR + ', APL_QTEREALUREF = ' + variantToSql(AFConversionUnite(ToItem.GetValue('APL_UNITETEMPS'), getParamSocSecur('SO_AFMESUREACTIVITE',''),
    valeur(ToItem.GetValue('APL_QTEREALISE'))));

  vStQR := vStQR + ', APL_PUPR = ' + variantToSql(ToItem.GetValue('APL_PUPR'));
  vStQR := vStQR + ', APL_PUPRDEV = ' + variantToSql(ToItem.GetValue('APL_PUPRDEV'));
  vStQR := vStQR + ', APL_INITPTPR = ' + variantToSql(ToItem.GetValue('APL_INITPTPR'));
  vStQR := vStQR + ', APL_INITPTPRDEV = ' + variantToSql(ToItem.GetValue('APL_INITPTPRDEV'));
  vStQR := vStQR + ', APL_REALPTPR = ' + variantToSql(ToItem.GetValue('APL_REALPTPR'));
  vStQR := vStQR + ', APL_REALPTPRDEV = ' + variantToSql(ToItem.GetValue('APL_REALPTPRDEV'));

  vStQR := vStQR + ', APL_PUVENTEHT = ' + variantToSql(ToItem.GetValue('APL_PUVENTEHT'));
  vStQR := vStQR + ', APL_PUVENTEDEVHT = ' + variantToSql(ToItem.GetValue('APL_PUVENTEDEVHT'));
  vStQR := vStQR + ', APL_INITPTVENTEHT = ' + variantToSql(ToItem.GetValue('APL_INITPTVENTEHT'));
  vStQR := vStQR + ', APL_INITPTVTDEVHT = ' + variantToSql(ToItem.GetValue('APL_INITPTVTDEVHT'));
  vStQR := vStQR + ', APL_REALPTVENTEHT = ' + variantToSql(ToItem.GetValue('APL_REALPTVENTEHT'));
  vStQR := vStQR + ', APL_REALPTVTDEVHT = ' + variantToSql(ToItem.GetValue('APL_REALPTVTDEVHT'));

  vStQR := vStQR + ' WHERE APL_AFFAIRE = "' + ToItem.GetValue('APL_AFFAIRE');
  vStQR := vStQR + '" AND APL_NUMEROLIGNE = ' + varAsType(ToItem.GetValue('APL_NUMEROLIGNE'), varString);
  vStQR := vStQR + ' AND APL_TYPEPLA = "PLA"';

  try
    if (ExecuteSql(vStQR) = 1) then
    begin
      result := true;
      RafraichirItemPartiel(ToItem, pStRetour, 'ARTICLE');
      UpdatePlanifieTache(ToItem.GetString('APL_AFFAIRE'), ToItem.GetString('APL_NUMEROTACHE'), ToItem.GetString('APL_UNITETEMPS'));
    end;
  except
    PGIBoxAF(TexteMessage[1], TexteMessage[59]);
    result := false;
  end;}
  
end;

procedure TPlanningGA.DeleteItem(Sender: TObject; Item: TOB; var Cancel: boolean);
var
  vStAffaire: string;
  vStNumTache: string;
  vStUnitTemps: string;
begin
 {if item <> nil then
  begin
    if Item.GetValue('APL_ETATLIGNE') = 'ABS' then
    begin
      cancel := True;
      PGIBoxAF(TexteMessage[12], TexteMessage[59]); //On ne peux rien faire sur une Absence.
    end
    else
    begin

      if not JaiLeDroitRessource(Item.getValue('APL_RESSOURCE')) then
      begin
        PgiBoxAf(TexteMessage[24], TexteMessage[59]); //La ressource ne fait pas partie de votre équipe
        cancel := true;
      end

      else
        if (not EtatModifAutorisee(Item.GetValue('APL_ETATLIGNE'), false, True)) or (IsBloque(Item)) then
      begin
        PgiBoxAf(TexteMessage[32], TexteMessage[59]); //Suppression interdite
        cancel := true;
      end

      else
      begin
        // C.B 13/09/2005
        // demander une confirmation avant suppression
        if (PGIAskAF(TexteMessage[58], '') = mrYes) then //Voulez-vous vraiment supprimer cette intervention ?
        begin
          cancel := false;
          vStAffaire := Item.GetString('APL_AFFAIRE');
          vStNumTache := Item.GetString('APL_NUMEROTACHE');
          vStUnitTemps := Item.GetString('APL_UNITETEMPS');
          Item.DeleteDB(False);
          //C.B 05/05/2006
          //C.B 22/08/2006
          if VH_GC.GCIfDefCEGID then
            RevalorisationCegid(nil, Item, 'DELETE');
          UpdatePlanifieTache(vStAffaire, vStNumTache, vStUnitTemps);
          FFichePlanning.LA_QUANTITE.Caption := '';
        end
        else
          cancel := true;
      end;
    end;
  end;}
end;

function TPlanningGA.GetTiers(pStAffaire: string): string;
var
  vQR: TQuery;
begin
  vQr := nil;
  try
    vQR := OpenSql('SELECT AFF_TIERS FROM AFFAIRE WHERE AFF_AFFAIRE = "' + pStAffaire + '"', True,-1,'',true);
    if not vQR.Eof then
      result := vQR.FindField('AFF_TIERS').AsString;
  finally
    if vQR <> nil then
      ferme(vQR);
  end;
end;

procedure TPlanningGA.DoubleClickSpec(ACol, ARow: INTEGER; TypeCellule: TPlanningTypeCellule; T: TOB = nil);
var
  vSt: string;
  vQr: TQuery;
  vStAff0:String;
  vStAff1:String;
  vStAff2:String;
  vStAff3:String;
  vStAvenant:String;

begin

  if (TypeCellule = ptcRessource) then
  begin
    if (T <> nil) and (T.FieldExists('CLE')) then
    begin
      DecodeCle(T.GetValue('CLE'));
                                                             
      if (fStAffaire <> '') and (fAction <> TaConsult) then
      begin
        // chargement de l'affaire
        fStAffaireSel := fStAffaire;
        fStTiersSel := fStTiers;
        CodeAffaireDecoupe(fStAffaire, vStAff0, vStAff1, vStAff2, vStAff3, vStAvenant, taModif, false);
        FFichePlanning.ED_AFF0.text := vStAff0;
        FFichePlanning.ED_AFF1.text := vStAff1;
        FFichePlanning.ED_AFF2.text := vStAff2;
        FFichePlanning.ED_AFF3.text := vStAff3;
        FFichePlanning.ED_AVE.text := vStAvenant;
        FFichePlanning.ED_TIERS.text := fStTiers;

        if FFichePlanning.LoadArticlesSel then
        begin
          // si tache connu
          if fStCodeArt <> '' then
          begin

            // affaire sans taches, on efface la sélection
            // ou on recherche toutes les taches déjà planifiées
            if not FFichePlanning.SelectArt(-1, fStTache) then
              FFichePlanning.SelectToutesTaches

            // tache connue, on la sélectionne
            else
              FFichePlanning.SelectArt(-1, fStTache);
          end
          else
            FFichePlanning.SelectArt(0, '');
        end

        // affaire sans taches, on efface la sélection
        // ou on recherche toutes les taches déjà planifiées
        else              
          FFichePlanning.SelectToutesTaches;
      end                    

      else if fStRes <> '' then
      begin
        vSt := 'SELECT ARS_CALENSPECIF, ARS_LIBELLE, ARS_STANDCALEN ';
        vSt := vSt + ' FROM RESSOURCE ';
        vSt := vSt + ' WHERE ARS_RESSOURCE = "' + fStRes + '"';

        vQr := OpenSQL(vSt, True,-1,'',true);
        try
          if not vQr.Eof then
            if vQR.FindField('ARS_CALENSPECIF').AsString = 'X' then
            begin
              AglLanceFiche('AFF', 'HORAIRESTD', '', '',
                'TYPE=RES;CODE=' + fStRes + ';LIBELLE=' + vQR.FindField('ARS_LIBELLE').AsString +
                ';STANDARD=' + vQR.FindField('ARS_STANDCALEN').AsString);
            end
            else
              AglLanceFiche('AFF', 'HORAIRESTD', '', '', 'TYPE:STD;STANDARD:' + vQR.FindField('ARS_STANDCALEN').asSTring + ';');
        finally
          ferme(vQr);
        end;
      end;
    end;
  end;
end;

procedure TAF_Planning.FormDestroy(Sender: TObject);
begin
  application.HintHidePause := 10000;
  HintWindowClass := THintWindow;
  try
    if assigned(FPlanningJour) then
      FreeAndNil(FPlanningJour);
    if assigned(fTobModelePlanning) then
      fTobModelePlanning.Free;
  finally
    fTobModelePlanning := nil;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 15/03/2002
Modifié le ... :
Description .. : retourne le code de l'etat de la tablette AFTETAT
Suite ........ :
Mots clefs ... :
*****************************************************************}
function TPlanningGA.ChangeEtats(ZCode: Integer): string;
var
  vQR: TQuery;
  vSt: string;
  vStEtats: string;
  vStEtat: string;

begin

  {result := '';
  if (fTobEtats = nil) then
  begin
    //C.B 26/10/2004 GESTION DES ABSENCES
    vSt := 'SELECT APP_CODEPARAM, APP_LIBELLEPARAM FROM AFPLANNINGPARAM ';
    vSt := vSt + ' WHERE APP_TYPEAFPARAM="STA" AND APP_CODEPARAM <> "ABS" ';
    vSt := vSt + ' AND APP_CODEPARAM <> "DEC" AND APP_MODIFIABLE = "X"';

    //C.B 22/11/2005 SUPPRESSION DE LA LISTE DES ETATS CACHES
    // si on n'est pas en saisie manager
    if (not SaisiePlanningManager) then
    begin
      vStEtats := getParamSocSecur('SO_AFETATINTERDIT','');
      while vStEtats <> '' do
      begin
        vStEtat := Trim(ReadTokenSt(vStEtats));
        vSt := vSt + ' AND APP_CODEPARAM <> "' + vStEtat + '"';
      end;
    end;

    vQr := nil;

    fTobEtats := Tob.Create('#AFPLANNINGEPARAM', nil, -1);
    try
      vQR := OpenSql(vSt, True);
      if not vQR.Eof then
        fTobEtats.LoadDetailDB('AFPLANNINGPARAM', '', '', vQR, False, True);
    finally
      Ferme(vQR);
    end;
  end;
  result := fTobEtats.Detail[ZCode - cInPopUp].GetValue('APP_CODEPARAM');}
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 15/04/2002
Modifié le ... :
Description .. : Gestion des etats
Suite ........ :
Mots clefs ... :
*****************************************************************}
procedure TPlanningGA.OnPopup(Item: tob; ZCode: integer; var Redraw: boolean);
var
  vBoStop : Boolean;

begin

  {if item <> nil then
  begin

    // C.B 11/07/2006
    // ajout d'un menu cegid feuille de présence reçu
    // C.B 25/08/2006
    // changement du test
    // a ne pas déplacer, toujours avant changeEtats car pas le même nb d'états
   	vBoStop := False;
    if VH_GC.GCIfDefCEGID and (ZCode = (fNbEtat + cInPopUp)) then
	  begin
  	  item.putValue('APL_ETATLIGNE', 'TER');
			item.putValue('APL_DATELIBRE1', Date);
			item.putValue('APL_CHARLIBRE2', V_PGI.UserName);
		  item.putValue('APL_BOOLLIBRE1', 'X');
    	Item.InsertOrUpdateDB;
			redraw := True;
  	  DisplayHint(Item, '');
    	vBoStop := True;
    end
    //C.B 07/09/2006
    // si pas ri recu et pas facturé, on remet les valeurs de RI
    // et on passe dans le traitement habituel
    else if VH_GC.GCIfDefCEGID and (ChangeEtats(ZCode) <> 'FAC') then
    begin
      item.putValue('APL_DATELIBRE1', iDate1900);
			item.putValue('APL_CHARLIBRE2', '');
		  item.putValue('APL_BOOLLIBRE1', '-');
    	Item.InsertOrUpdateDB;
 			redraw := True;
  	  DisplayHint(Item, '');
    	vBoStop := False;
    end;

    if Item.GetValue('APL_ETATLIGNE') = 'ABS' then
    begin
      redraw := false;
      PGIBoxAF(TexteMessage[12], TexteMessage[59]); //On ne peux rien faire sur une Absence.
    end

    else if not vBoStop then
    begin
      //mcd 14/03/2005
      if not JaiLeDroitRessource(Item.getValue('APL_RESSOURCE')) then
      begin
        PgiBoxAf(TexteMessage[24], TexteMessage[59]); //La ressource ne fait pas partie de votre équipe
        redraw := False;
      end
      else
      begin
        // C.B 17/07/2005
        // ajout de tests sur changement d'état
        if EtatModifAutorisee(Item.GetValue('APL_ETATLIGNE'), false, True) then
        begin
          if (not IsBloque(Item)) then
          begin
            if ((ChangeEtats(ZCode) = 'FAC') and (PGIAskAF(TexteMessage[57], '') = mrYes)) or
              (ChangeEtats(ZCode) <> 'FAC') then
            begin
              Item.PutValue('APL_ETATLIGNE', ChangeEtats(ZCode));
              redraw := UpdateItem(nil, Item, 'POPUP');
            end
          end
          else
            PGIBoxAF(TexteMessage[18], TexteMessage[59]); //Cette intervention est bloquée par %s
        end
          // pour ne pas avoir 2 messages, car il y en a deja un dans SaisiePlanningManager
        else
          if not SaisiePlanningManager then
          PGIBoxAF(TexteMessage[36], TexteMessage[59]); //L''état de cette intervention ne permet aucune modification
      end;
    end;

    // à enlever dans les versions suivantes
    if redraw then
      Item.loadDB(False);
  end;}
end;

procedure TPlanningGA.LoadAffNonProductive; // Journée non productive
begin
  fStAffNonProductive := '';
  fStlibNonProductive := TexteMessage[63]; //NON FACTURABLES

  if (trim(getParamSocSecur('SO_AFPLANAFFAIRE','')) <> '') and
    ExisteSql('SELECT AFF_AFFAIRE FROM AFFAIRE WHERE AFF_AFFAIRE = "' + getParamSocSecur('SO_AFPLANAFFAIRE','') + '"') then
    fStAffNonProductive := getParamSocSecur('SO_AFPLANAFFAIRE','')
  else
    SetParamsoc('SO_AFPLANAFFAIRE', '');
end;

procedure TPlanningGA.LoadItemsPlanning(pStWhere: string);
var
  S: string;
  Q: TQuery;

begin

  {S := ' SELECT APL_TYPEPLA, APL_TYPELIGNEPLA, APL_AFFAIRE, APL_AFFAIRE1, APL_AFFAIRE2, APL_AFFAIRE3, ';
  S := S + ' APL_AFFAIRE0, APL_AVENANT, APL_FONCTION, APL_RESSOURCE, ';
  S := S + ' APL_TIERS, APL_NUMEROTACHE, APL_NUMEROLIGNE, APL_LIBELLEPLA, APL_UNITETEMPS, ';
  S := S + ' APL_QTEPLANIFIEE, APL_QTEREALISE, APL_DEVISE, ';
  S := S + ' APL_QTEPLANIFUREF, APL_QTEREALUREF, APL_ETATLIGNE, ';
  S := S + ' APL_UTILISATEUR, APL_BLOQUE, ';
  S := S + ' APL_DATEDEBPLA, APL_DATEFINPLA, APL_HEUREDEB_PLA, APL_HEUREFIN_PLA, ';
  S := S + ' APL_DATEDEBREAL, APL_DATEFINREAL, APL_HEUREDEB_REAL, APL_HEUREFIN_REAL, ';
  S := S + ' APL_PUPR, APL_PUPRDEV, APL_PUVENTEHT, APL_PUVENTEDEVHT, APL_ARTICLE, ';
  S := S + ' APL_CODEARTICLE, APL_TYPEARTICLE, ';
  S := S + ' APL_INITPTPR, APL_INITPTPRDEV, APL_INITPTVENTEHT, APL_INITPTVTDEVHT, ';
  S := S + ' APL_REALPTPR, APL_REALPTPRDEV, APL_REALPTVENTEHT, APL_REALPTVTDEVHT, ';
  S := S + ' APL_TYPEADRESSE, APL_NUMEROADRESSE, ATA_TYPEADRESSE, ATA_NUMEROADRESSE, ';
  S := S + ' ATA_LIBELLETACHE1, ATA_QTEINITIALE, ATA_IDENTLIGNE, ';
  S := S + ' ATA_NBPARTINSCRIT, ATA_NBPARTPREVU, ATA_STATUTPLA,';
  S := S + ' ATA_TYPEPLANIF, ARS_LIBELLE, ARS_LIBELLE2, T_LIBELLE, APL_BOOLLIBRE1, APL_BOOLLIBRE2, APL_BOOLLIBRE3, ';
  S := S + ' APL_LIBRETACHE1, APL_LIBRETACHE2, APL_LIBRETACHE3, APL_LIBRETACHE4, APL_LIBRETACHE5, ';
  S := S + ' APL_LIBRETACHE6, APL_LIBRETACHE7, APL_LIBRETACHE8, APL_LIBRETACHE9, APL_LIBRETACHEA, ';
  S := S + ' APL_CHARLIBRE1, APL_CHARLIBRE2, APL_CHARLIBRE3, APL_VALLIBRE1, APL_VALLIBRE2, APL_VALLIBRE3, ';
  S := S + ' APL_DATELIBRE1, APL_DATELIBRE2, APL_DATELIBRE3, APL_DESCRIPTIF, ATA_DESCRIPTIF, ATA_TERMINE, AFF_LIBELLE, ';

  //C.B 21/09/2006
  //optimisation
  // on charge les données du tiers directement dans le loadItem
	S := S + ' T_PRENOM, T_ADRESSE1, T_ADRESSE2, T_ADRESSE3, T_CODEPOSTAL, T_VILLE, T_TELEPHONE ';

  S := S + ' FROM AFFAIRE, TACHE, AFPLANNING, RESSOURCE, TIERS, PIECE ';
  S := S + ' WHERE APL_TYPEPLA = "PLA"';
  S := S + ' AND APL_AFFAIRE = AFF_AFFAIRE';
  S := S + ' AND APL_AFFAIRE = ATA_AFFAIRE';
  S := S + ' AND APL_NUMEROTACHE = ATA_NUMEROTACHE ';
  S := S + ' AND ATA_TYPEARTICLE = "PRE"';
  S := S + ' AND AFF_AFFAIRE = GP_AFFAIRE';
  S := S + ' AND GP_NATUREPIECEG = AFF_NATUREPIECEG';

  // C.B 04/10/2005
  // on affiche le planning des clients fermés
  //S := S + ' AND ARS_FERME = "-"';

  S := S + ' AND APL_TIERS = T_TIERS';
  S := S + ' AND APL_RESSOURCE = ARS_RESSOURCE';

  // limitation du planning
  pStWhere := pStWhere + ' AND APL_DATEDEBPLA <= "' + usDateTime(fDtFin) + '"';
  pStWhere := pStWhere + ' AND APL_DATEFINPLA >= "' + usDateTime(fDtDebut) + '"';

  if pos('ATR_NUMEROTACHE', pStWhere) <> 0 then
    ReplaceSubStr(pStWhere, 'ATR_NUMEROTACHE', 'ATA_NUMEROTACHE');

  S := S + pStWhere;

  fTobsPlanningTobItems.ClearDetail;
  Q := nil;
  try
    Q := OpenSql(S, True);
    if not Q.Eof then
    begin
      fTobsPlanningTobItems.LoadDetailDB('AFPLANNING', '', '', Q, False, True);
      fTobsPlanningTobItems.detail[0].AddChampSup('I_HINT', true);

      // C.B 13/05/2005 spécifique cegid
      if VH_GC.GCIfDefCEGID then
        fTobsPlanningTobItems.detail[0].AddChampSupValeur('I_ICONE', cInIconeNORMALE, true);

      fTobsPlanningTobItems.detail[0].AddChampSup('I_CODE', true);
      fTobsPlanningTobItems.detail[0].AddChampSupValeur('CHAMPLIBELLE', '', true);
      fTobsPlanningTobItems.detail[0].AddChampSupValeur('CODEPOSTAL', '', true);
    end;
  finally
    Ferme(Q);
  end;}
end;

// C.B 19/05/2006
// chargement des absences des bases ccmx et cegid en dur
function TPlanningGA.MakeRequetePaie(pStWhere : String) : String;
begin
  // C.B 19/01/2006
  // on ramène PMA_TYPEABS = "RTT" si RTT
  // et PCN_VALIDRESP = "VAL" si validé
  {Result := ' SELECT ARS_RESSOURCE, ARS_LIBELLE, ARS_LIBELLE2, PCN_DATEDEBUTABS,PCN_DATEFINABS '; //PCN_TYPECONGE
  Result := Result + ',PCN_DEBUTDJ, PCN_FINDJ, PMA_TYPEABS, PCN_VALIDRESP ';

  // C.B 25/08/2006
  // ajout du test des noms de société pour tester cegid sur d'autres bases
  if VH_GC.GCIfDefCEGID and ((v_PGI.DBName = 'CCMX') or (v_PGI.DBName = 'CEGID')) then
    result := result + ' FROM RESSOURCE, CEGID.dbo.ABSENCESALARIE '
  else
    Result := Result + ' FROM RESSOURCE, ABSENCESALARIE ';

  Result := Result + ' LEFT JOIN MOTIFABSENCE ON ##PMA_PREDEFINI## PMA_MOTIFABSENCE = PCN_TYPECONGE';
  Result := Result + ' WHERE ARS_SALARIE = PCN_SALARIE';
  Result := Result + MakeWherePaie;
  Result := Result + pStWhere;
  Result := Result + ' AND PCN_DATEDEBUTABS <= "' + UsDateTime(fDtFin) + '"';
  Result := Result + ' AND PCN_DATEFINABS >= "' + UsDateTime(fDtDebut) + '"';

  // C.B 25/08/2006
  // ajout du test des noms de société pour tester cegid sur d'autres bases
  if VH_GC.GCIfDefCEGID and ((v_PGI.DBName = 'CCMX') or (v_PGI.DBName = 'CEGID')) then
  begin
    Result := Result + ' UNION SELECT ARS_RESSOURCE, ARS_LIBELLE, ARS_LIBELLE2, PCN_DATEDEBUTABS,PCN_DATEFINABS ';
    Result := Result + ',PCN_DEBUTDJ, PCN_FINDJ, PMA_TYPEABS, PCN_VALIDRESP ';
    result := result + ' FROM RESSOURCE, CCMX.dbo.ABSENCESALARIE ';
    Result := Result + ' LEFT JOIN MOTIFABSENCE ON ##PMA_PREDEFINI## PMA_MOTIFABSENCE = PCN_TYPECONGE';
    Result := Result + ' WHERE ARS_SALARIE = PCN_SALARIE';
    Result := Result + MakeWherePaie;
    Result := Result + pStWhere;
    Result := Result + ' AND PCN_DATEDEBUTABS <= "' + UsDateTime(fDtFin) + '"';
    Result := Result + ' AND PCN_DATEFINABS >= "' + UsDateTime(fDtDebut) + '"';
    Result := Result + ' ORDER BY PCN_DATEDEBUTABS,PCN_DATEFINABS';
	end
  else
    Result := Result + ' ORDER BY PCN_DATEDEBUTABS,PCN_DATEFINABS';}

end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... :
Modifié le ... :
Description .. : chargement des etats des items
Mots clefs ... :
*****************************************************************}
function TPlanningGA.LoadEtatPlanning: Boolean;
var
  S: string;
  vQR: TQuery;
  vListe: TStringList;
  vStream: TStream;
  i: Integer;

begin

  result := True;
  S := 'SELECT APP_CODEPARAM, APP_LIBELLEPARAM, APP_MODIFIABLE, APP_PARAMS FROM AFPLANNINGPARAM WHERE APP_TYPEAFPARAM="STA"';
  fTobEtatXml.ClearDetail;
  vListe := TStringList.Create;
  vQr := nil;
  try
    vQR := OpenSql(S, True,-1,'',true);
    if not vQR.Eof then
      fTobEtatXml.LoadDetailDB('AFPLANNINGPARAM', '', '', vQR, False, True);

    fTobsPlanningTobEtats.ClearDetail;
    for i := 0 to fTobEtatXml.Detail.Count - 1 do
    begin
      // chargement de la liste
      vListe.Text := fTobEtatXml.Detail[i].GetValue('APP_PARAMS');
      // transfert dans une stream
      if vListe.Text = '' then
      begin
        PGIBox(format(traduitGA(TexteMessage[33]), [fTobEtatXml.Detail[i].GetValue('APP_LIBELLEPARAM')]), TraduitGA(TexteMessage[59])); //L''état %s n''est pas paramétré.
        result := False;
      end
      else
      begin
        vStream := TStringStream.Create(vListe.Text);
        try
          // recuperation dans une tob virtuelle TobsPlanningTobEtats
          TOBLoadFromXMLStream(vStream, LoadTobEtat);
          // ajout du code et du libelle
          fTobsPlanningTobEtats.detail[i].AddChampSupValeur('E_CODE', fTobEtatXml.Detail[i].GetValue('APP_CODEPARAM'), false);
          fTobsPlanningTobEtats.detail[i].AddChampSupValeur('E_LIBELLE', fTobEtatXml.Detail[i].GetValue('APP_LIBELLEPARAM'), false);
          fTobsPlanningTobEtats.detail[i].AddChampSupValeur('E_MODIFIABLE', fTobEtatXml.Detail[i].GetValue('APP_MODIFIABLE'), false);
        finally
          vStream.Free;
        end;
      end;
    end;

  finally
    Ferme(vQR);
    vListe.Free;
  end;
end;

procedure TPlanningGA.LoadArgument;
var
  Tmp: string;
  champ: string;
  valeur: string;

begin
  // traitement des arguments
  fBoMonPlanning := False;
  Tmp := (Trim(ReadTokenSt(fStArgument)));
  while (Tmp <> '') do
  begin
    if Tmp <> '' then
    begin
      DecodeArgument(Tmp, Champ, valeur);
      if Champ = 'ACTION' then
      begin
        if valeur = 'MODIFICATION' then
          fAction := taModIf
        else
          if valeur = 'CONSULTATION' then
        begin
          fAction := taConsult;
        end;
      end

      // C.B 13/07/2006
      else if champ = 'MONPLANNING' then
        fBoMonPlanning := True
              
      else if champ = 'BUFFER' then
      begin
        fTobSelection := MaTobInterne('AFFAIRECOURANTE').detail[0];

        // en attendant de passer par la tob, on affecte toutes les variables
        fStAffaireSel := fTobSelection.GetString('APL_AFFAIRE');
        fStTiersSel := fTobSelection.GetString('APL_TIERS');
        fStTacheSel := fTobSelection.GetString('APL_NUMEROTACHE');

        // C.B supprimé le 16/08/2005
        // a priori ne sert pas
        // controler que cela n'implique pas des bugs
        fStArticleSel := fTobSelection.GetString('APL_ARTICLE');
        fStCodeArtSel := fTobSelection.GetString('APL_CODEARTICLE');
        fStTypeArtSel := fTobSelection.GetString('APL_TYPEARTICLE');
        fStLibelleArtSel := fTobSelection.GetString('APL_LIBELLEARTICLE');

        if fTobSelection.FieldExists('IDENTLIGNE') then
          fStIdentLigne := fTobSelection.GetString('IDENTLIGNE') //AB-200511- Identifiant de la commande d'origine
        else
          fStIdentLigne := '';
        FFichePlanning.LA_SELECTION.Caption := ModeCreationDirecte(False);

        // C.B 29/09/2005
        // cacher le bouton ligne d'affaire quand on vient déja des lignes
        FFichePlanning.BT_LIGNES.visible := False;
      end;
    end;
    Tmp := (Trim(ReadTokenSt(fStArgument)));
  end;
end;

procedure TPlanningGA.EncodeCle(pStCle1, pStCle2, pStCle3: string; var pStCle: string);
begin
  if pStCle3 <> '' then
    pStCle := pStCle1 + cStSep + pStCle2 + cStSep + pStCle3
  else
    if pStCle2 <> '' then
    pStCle := pStCle1 + cStSep + pStCle2
  else
    pStCle := pStCle1;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 04/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TPlanningGA.DisplayHint(pItem: Tob; pStHint: string; pTobTache: Tob): string;
var
  vSt: string;
  vTobEtat		: Tob;
	vStAffaire  : String;

begin

  if assigned(pItem) then
  begin
    if pStHint <> '' then
      vSt := pStHint
    else
    begin
      if fInModeCreation = cInCreationNonProductive then
        vSt := ' ' + TraduitGA(TexteMessage[56]) + #10#13 //Affaire non facturable

      // on est en creation directe apres une creation
      else
        if (fStAffaireSel <> '') and (pTobTache <> nil) then
      begin
                                                                   
        if GetParamsocSecur('SO_AFAFFICHAFFAIRE', False) then
          vStAffaire := pItem.GetString('AFF_LIBELLE') + #13#10 + '               ' + CodeAffaireAffiche(fStAffaireSel, ' ')
        else
          vStAffaire := CodeAffaireAffiche(fStAffaireSel, ' ');

        vSt := ' ' + TraduitGA(TexteMessage[54]) + ' : ' + vStAffaire + #10#13; // Affaire
        vSt := vSt + ' ' + TexteMessage[49] + ' : ' + pItem.GetString('APL_CODEARTICLE') + #10#13; // article
        vSt := vSt + ' ' + TexteMessage[50] + ' : ' + pItem.GetString('APL_LIBELLEPLA') + #10#13; // libellé
        vSt := vSt + ' ' + TraduitGA(TexteMessage[51]) + ' : ' + pItem.GetString('ARS_LIBELLE') + #10#13; // ressource
        vSt := vSt + ' ' + TexteMessage[52] + ' : ' + fStTiersSel + ' ' + pTobTache.GetValue('T_LIBELLE') + #10#13; // client
        vSt := vSt + ' ' + TexteMessage[53] + ' : ' + v_pgi.user; // Dernière modif

        // si affaire école, on affiche dans le hint le nb d'intervenants
        if getParamSocSecur('SO_AFECOLE', False) and (pTobTache.GetString('ATA_TYPEPLANIF') = 'ECO') then
          vSt := vSt + #10#13 + ' ' + format(traduitGA(TexteMessage[64]), [pTobTache.GetString('ATA_NBPARTINSCRIT') + '/' + pTobTache.GetString('ATA_NBPARTPREVU')]); //Nombre d''intervenants inscrits : %s
      end
      else
      begin
   
        if GetParamsocSecur('SO_AFAFFICHAFFAIRE', False) then
          vStAffaire := pItem.GetString('AFF_LIBELLE') + #13#10 + '               ' + CodeAffaireAffiche(pItem.GetString('APL_AFFAIRE'), ' ')
        else
          vStAffaire := CodeAffaireAffiche(pItem.GetString('APL_AFFAIRE'), ' ');

        vSt := ' ' + traduitGA(TexteMessage[54]) + ' : ' + vStAffaire + #10#13; // Affaire
        if pItem.GetNumChamp('ATA_LIBELLETACHE1') > 0 then
        begin
          vSt := vSt + ' ' + TexteMessage[49] + ' : ' + pItem.GetString('APL_CODEARTICLE') + #10#13; // article
          vSt := vSt + ' ' + TexteMessage[50] + ' : ' + pItem.GetString('ATA_LIBELLETACHE1') + #10#13; // libellé
          vSt := vSt + ' ' + traduitGA(TexteMessage[51]) + ' : ' + pItem.GetString('ARS_LIBELLE') + #10#13; // ressource
          vSt := vSt + ' ' + TexteMessage[52] + ' : ' + pItem.GetString('APL_TIERS') + ' ' + pItem.GetString('T_LIBELLE') + #10#13; // client
          vSt := vSt + ' ' + TexteMessage[53] + ' : ' + pItem.GetString('APL_UTILISATEUR'); // 'Dernière modif'
        end;

        if getParamSocSecur('SO_AFECOLE', False) and
          (pItem.GetNumChamp('ATA_LIBELLETACHE1') > 0) and
          (pItem.GetString('ATA_TYPEPLANIF') = 'ECO') then
        begin
          vSt := vSt + #10#13 + ' ' + format(traduitGA(TexteMessage[48]), [pItem.GetString('ATA_NBPARTINSCRIT') + '/' + pItem.GetString('ATA_NBPARTPREVU')]); //Nombre d''intervenants inscrits : %s
        end;
      end;

      // affichage de l'état dans le hint
      vTobEtat := fTobEtatXml.findfirst(['APP_CODEPARAM'], [pItem.GetString('APL_ETATLIGNE')], true);
      if vTobEtat <> nil then
        vSt := vSt + #10#13 + format(traduitGA(texteMessage[55]), [vTobEtat.GetString('APP_LIBELLEPARAM')]); // Statut : %s
    end;

    pItem.PutValue('I_HINT', vSt);
    result := vSt;
  end;
end;

function TPlanningGA.IsBloque(Item: Tob): Boolean;
begin
  {Result := (Item.GetString('APL_BLOQUE') = 'X') and
    (Item.GetString('APL_UTILISATEUR') <> V_PGI.User);

  // C.B 19/08/2005
  // les droits saisi manager planning donne accès aux interventions bloquées
  if result and SaisiePlanningManager then
    result := False;}
end;

procedure TPlanningGA.ChangeOptions(pBoValeur: Boolean);
begin
  FComposantPlanning.DisplayOptionSuppression := pBoValeur;
  FComposantPlanning.DisplayOptionModification := pBoValeur;
  FComposantPlanning.DisplayOptionCreation := pBoValeur;
  FComposantPlanning.DisplayOptionDeplacement := pBoValeur;
  FComposantPlanning.DisplayOptionCopie := pBoValeur;
  FComposantPlanning.DisplayOptionReduire := pBoValeur;
  FComposantPlanning.DisplayOptionEtirer := pBoValeur;

  // en attendant de les gérer
  FComposantPlanning.DisplayOptionLiaison := False;
  FComposantPlanning.DisplayOptionLier := False;
  FComposantPlanning.DisplayOptionSuppressionLiaison := False;
end;

procedure TPlanningGA.FormatChampLibelle(pItem: Tob);
var
  vSt           : String;
  vStAccompagne : String;
 
begin

  {vStAccompagne := '';

  if VH_GC.GCIfDefCEGID and
    (pItem.GetValue('APL_BOOLLIBRE2') = 'X') then
    vStAccompagne := '@ ';

  if pItem.GetString('CODEPOSTAL') = '' then
    pItem.PutValue('CODEPOSTAL',
      GetCodePostal(pItem, fTobAffAdm,
      pItem.GetString('APL_AFFAIRE'),
      pItem.GetString('APL_TIERS'),
      pItem.GetString('ATA_TYPEADRESSE'),
      pItem.GetString('ATA_NUMEROADRESSE'),
      pItem.GetString('APL_TYPEADRESSE'),
      pItem.GetString('APL_NUMEROADRESSE'), True, True));

  // ECOLE
  if pItem.GetString('ATA_TYPEPLANIF') = 'ECO' then
  begin
    if (ModeAffichage = 'ARTICLE') or (ModeAffichage = 'AFFAIRE') then
    begin
      pItem.putValue('CHAMPLIBELLE', vStAccompagne + pItem.GetValue('APL_LIBELLEPLA')
          + #13#10 + pItem.GetValue('ARS_LIBELLE') + ' ' + pItem.GetValue('ARS_LIBELLE2')
          + #13#10 + pItem.GetString('ATA_NBPARTINSCRIT') + '/' + pItem.GetString('ATA_NBPARTPREVU'))
    end
    else if ModeAffichage = 'RESSOURCE' then
      pItem.putValue('CHAMPLIBELLE', vStAccompagne + pItem.GetValue('ARS_LIBELLE') + ' ' + pItem.GetValue('ARS_LIBELLE2'));
  end

  // INTERVENTIONS
  else
  begin

    // CAS PARTICULIER AFFAIRE
    if ModeAffichage = 'AFFAIRE' then
    begin
        pItem.putValue('CHAMPLIBELLE', vStAccompagne + pItem.GetValue('APL_AFFAIRE')
          + #13#10 + pItem.GetValue('T_LIBELLE')
          + #13#10 + pItem.GetValue('CODEPOSTAL'));
    end
    // TOUS LES AUTRES PLANNINGS
    else
    begin
      vSt := vStAccompagne + pItem.GetValue('APL_LIBELLEPLA');
      if AfficherRess then
        vSt := vSt + #13#10 + pItem.GetValue('ARS_LIBELLE') + ' ' + pItem.GetValue('ARS_LIBELLE2');
      if AfficherTiers then
        vSt := vSt + #13#10 + pItem.GetValue('T_LIBELLE');
      vSt := vSt + #13#10 + pItem.GetValue('CODEPOSTAL');
      pItem.putValue('CHAMPLIBELLE', vSt);
    end;
  end;}
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 12/05/2005
Modifié le ... :   /  /
Description .. : rafraichi les éléments affichés graphiquement dans l'item
                 zprès création ou modification
Mots clefs ... :
*****************************************************************}
procedure TPlanningGA.RafraichirItem(Item: Tob; pStRetour, pStMode: string; pBoLoadTache, pBoLoadCodePostal : Boolean);
begin
  {Item.loadDB(False);

  if pBoLoadCodePostal then
    Item.PutValue('CODEPOSTAL', GetCodePostal(Item, fTobAffAdm, Item.GetString('APL_AFFAIRE'),
      Item.GetString('APL_TIERS'),
      Item.GetString('ATA_TYPEADRESSE'),
      Item.GetString('ATA_NUMEROADRESSE'),
      Item.GetString('APL_TYPEADRESSE'),
      Item.GetString('APL_NUMEROADRESSE'), True, True));

  if pBoLoadTache then
    LoadComplementItem(Item);

  FormatChampLibelle(Item);
  DisplayHint(Item, pStRetour);
  FComposantPlanning.InvalidateItem(Item);
  AffichageCommentaire(Item);
  AffichageQuantites(Item);}
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 06/12/2005
Modifié le ... :
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TPlanningGA.RafraichirAssocies(Item: Tob);
var
  vItem: Tob;
begin
  vItem := fTobsPlanningTobItems.FindFirst(['APL_DATEDEBPLA', 'APL_AFFAIRE', 'APL_NUMEROTACHE'],
    [Item.GetValue('APL_DATEDEBPLA'), Item.GetValue('APL_AFFAIRE'), Item.GetValue('APL_NUMEROTACHE')], true);
  RafraichirItem(vItem, '', 'ARTICLE', True, False);

  while vItem <> nil do
  begin
    vItem := fTobsPlanningTobItems.FindNext(['APL_DATEDEBPLA', 'APL_AFFAIRE', 'APL_NUMEROTACHE'],
      [Item.GetValue('APL_DATEDEBPLA'), Item.GetValue('APL_AFFAIRE'), Item.GetValue('APL_NUMEROTACHE')], true);
    if vItem <> nil then
      RafraichirItem(vItem, '', 'ARTICLE', True, False);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 21/11/2005
Modifié le ... : rafraichir sans rechargement base de donnée de l'item
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TPlanningGA.RafraichirItemPartiel(Item: Tob; pStRetour, pStMode: string);
begin
  FormatChampLibelle(Item);
  DisplayHint(Item, pStRetour);
  FComposantPlanning.InvalidateItem(Item);
end;

// C.B 13/05/2005 spécifique cegid
procedure TPlanningGA.IconeCEGID(i: Integer);
begin
  if VH_GC.GCIfDefCEGID then
  begin
    if fTobsPlanningTobItems.detail[i].GetString('APL_BOOLLIBRE3') = 'X' then
      fTobsPlanningTobItems.detail[i].putValue('I_ICONE', cInIconeINTERNE)
    else
      if fTobsPlanningTobItems.detail[i].GetValue('ATA_TYPEPLANIF') = 'ECO' then
      fTobsPlanningTobItems.detail[i].putValue('I_ICONE', cInIconeECOLE)
    else
      fTobsPlanningTobItems.detail[i].putValue('I_ICONE', cInIconeNORMALE);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 21/10/2002
Modifié le ... :
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TPlanningGA.UpdateData(FromItem, ToItem: Tob);
var
  vTOBAffaires: Tob;
  vTOBArticles: Tob;
  vAFOAssistants: TAFO_Ressources;
  vArticle: RecordArticle;

begin

  {ModifQte(FromItem, ToItem, fBoDemiJournee);

  vAFOAssistants := TAFO_Ressources.Create;
  vTOBAffaires := TOB.Create('#AFFAIRE', nil, -1);
  vTOBArticles := TOB.Create('#ARTICLE', nil, -1);

  try
    // apl_codeArticle n'est pas toujours initialisé (selon comment on click dans le planning)
    // si il n'est pas renseigné, on va le relire ...
    if ToItem.Getvalue('APL_CODEARTICLE') = '' then
    begin
      GetArticle(ToItem.Getvalue('APL_ARTICLE'), vArticle);
      ToItem.PutValue('APL_CODEARTICLE', vArticle.StCodeArticle);
      ToItem.PutValue('APL_TYPEARTICLE', vArticle.StTypeArticle);
    end;

    ValorisationPlanning (ToItem, 'APL', vAFOAssistants, vTOBAffaires, vTobArticles, true, true);
  finally
    vTOBAffaires.Free;
    vTobArticles.Free;
    vAFOAssistants.Free;
  end;}
end;

procedure TPlanningGA.BExcelClick(pStFileName: string; pBoOuvrir: Boolean);
begin
  FComposantPlanning.ExportToExcel(pBoOuvrir, pStFileName, pBoOuvrir)
end;

procedure TPlanningGA.AffichageCommentaire(pItem: Tob);
begin
  if trim(GetRTFStringText(pItem.GetValue('ATA_DESCRIPTIF'))) <> '' then
  begin
    if trim(GetRTFStringText(pItem.GetValue('APL_DESCRIPTIF'))) <> '' then
      FFichePlanning.LA_COMMENTAIRE.Caption := Trim(GetRTFStringText(pItem.GetValue('ATA_DESCRIPTIF'))) + ' / ' + Trim(GetRTFStringText(pItem.GetValue('APL_DESCRIPTIF')))
    else
      FFichePlanning.LA_COMMENTAIRE.Caption := Trim(GetRTFStringText(pItem.GetValue('ATA_DESCRIPTIF')));
  end
  else
    FFichePlanning.LA_COMMENTAIRE.Caption := Trim(GetRTFStringText(pItem.GetValue('APL_DESCRIPTIF')));
end;

procedure TPlanningGA.AffichageQuantites(pItem: Tob);
begin
  //if pItem <> nil then FFichePlanning.LA_QUANTITE.Caption := LoadQteHint(pItem);
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 20/01/2006
Modifié le ... :
Description .. : on ramène PMA_TYPEABS = "RTT" si RTT
                 et PCN_VALIDRESP = "VAL" si validé
Mots clefs ... :
*****************************************************************}
function TPlanningGA.FormatLibTache(pStRtt, pStValider : String) : String;
begin
  if (pStRtt = 'RTT') then
  begin
    if pStValider = 'VAL' then
      result := TraduireMemoire(TexteMessage[74]) //RTT VALIDE
    else
      result := TraduireMemoire(TexteMessage[73]); //RTT
  end
  else
  begin
    if pStValider = 'VAL' then
      result := TraduireMemoire(TexteMessage[72]) //ABSENCE VALIDEE
    else
      result := TraduireMemoire(TexteMessage[66]); //ABSENCE
  end;

  // C.B 11/07/2006
  // absence sur 3 lignes
  result := result + #13#10 + '' + #13#10 + '';
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 04/10/2005
Modifié le ... :
Description .. : Création d'une absence ajoutée dans la liste des items
Mots clefs ... :
*****************************************************************}
procedure TPlanningGA.CreateAbsence(pItem, pTobAbs : Tob; pStRes, pStICode : String; pIn, pInNumTache, pInNumligne : Integer);
var
  vStDebJ     : String;
  vStFinJ     : String;
  vDtDeb      : TDateTime;
  vDtFin      : TDateTime;
  vStRtt      : String;
  vStValider  : String;
  vRdQte      : Double;

begin

  vStDebJ     := pTobAbs.GetString('PCN_DEBUTDJ');
  vStFinJ     := pTobAbs.GetString('PCN_FINDJ');
  vDtDeb      := pTobAbs.GetValue('PCN_DATEDEBUTABS');
  vDtFin      := pTobAbs.GetValue('PCN_DATEFINABS');
  vStRtt      := pTobAbs.GetString('PMA_TYPEABS');
  vStValider  := pTobAbs.GetString('PCN_VALIDRESP');

  if not pItem.FieldExists('ATA_LIBELLETACHE1') then
  begin
    pItem.AddChampSupValeur('ATA_LIBELLETACHE1', '', False);
    pItem.AddChampSupValeur('ATA_QTEINITIALE', '', False);
    pItem.AddChampSupValeur('ARS_LIBELLE', '', False);
    pItem.AddChampSupValeur('ARS_LIBELLE2', '', False);
    pItem.AddChampSupValeur('T_LIBELLE', '', False);
    pItem.AddChampSupValeur('ATA_IDENTLIGNE', '', False);
    pItem.AddChampSupValeur('ATA_TYPEPLANIF', '', False);
    pItem.AddChampSupValeur('ATA_DESCRIPTIF', '', False);
    pItem.AddChampSupValeur('ATA_NBPARTINSCRIT', '0', False);
    pItem.AddChampSupValeur('ATA_NBPARTPREVU', '0', False);
    pItem.AddChampSupValeur('ATA_TYPEADRESSE', '', False);
    pItem.AddChampSupValeur('ATA_NUMEROADRESSE', '', False);
    pItem.AddChampSupValeur('ATA_STATUTPLA', '', False);

    // C.B 13/05/2005 spécifique cegid
    if VH_GC.GCIfDefCEGID then pItem.AddChampSupValeur('I_ICONE', cInIconeNORMALE, False);
    pItem.AddChampSupValeur ('I_HINT', '', False);
    pItem.AddChampSupValeur ('I_COLOR', '', False);
    pItem.AddChampSupValeur ('I_CODE', pStICode, False);
    pItem.AddChampSupValeur('CHAMPLIBELLE', traduireMemoire(texteMessage[66]), False); //'ABSENCE'
    pItem.AddChampSupValeur('CODEPOSTAL', '', False);
  end;

  pItem.putValue('APL_TYPEPLA', 'PLA');
  pItem.putValue('APL_AFFAIRE', '');
  pItem.putValue('APL_AFFAIRE0', '');
  pItem.putValue('APL_AFFAIRE1', '');
  pItem.putValue('APL_AFFAIRE2', '');
  pItem.putValue('APL_AFFAIRE3', '');
  pItem.putValue('APL_AVENANT', '');
  pItem.putValue('APL_FONCTION', '');
  pItem.putValue('APL_RESSOURCE', pStRes);
// C.B 06/06/2006
// optimisation
// loadLibelleRessource(pItem);
  pItem.putValue('APL_TIERS', '');
  pItem.putValue('APL_NUMEROTACHE', pInNumTache - pIn);
  pItem.putValue('APL_NUMEROLIGNE', intToStr(pInNumligne + pIn));
  pItem.putValue('APL_LIBELLEPLA', traduireMemoire(texteMessage[66])); //'ABSENCE'
  pItem.putValue('APL_UNITETEMPS', 'J');
  pItem.putValue('APL_QTEPLANIFIEE', '1');
  pItem.putValue('APL_QTEREALISE', '0');
  pItem.putValue('APL_DEVISE', GetParamSocSecur('SO_DEVISEPRINC', 'EUR'));

  // C.B 06/06/2006
  // optimisation
  // conversionUnite fait une requete
  vRdQte := ConversionUnite('J', getparamsoc('SO_AFMESUREACTIVITE'), 1);
  pItem.putValue('APL_QTEPLANIFUREF', vRdQte);
  pItem.putValue('APL_QTEREALUREF', vRdQte);
  pItem.putValue('APL_ETATLIGNE', 'ABS');
  pItem.putValue('APL_DATEDEBPLA', vDtDeb);
  pItem.putValue('APL_DATEFINPLA', vDtFin);
  pItem.putValue('APL_TYPELIGNEPLA', 'TAC');
  pItem.putValue('APL_TYPEARTICLE', '');
  pItem.putValue('APL_ARTICLE', '');
  pItem.putValue('APL_CODEARTICLE', '');
  pItem.putValue('APL_MOTIFETAT', '');
  pItem.putValue('APL_ACTIVITEGENERE', '-');
  pItem.putValue('APL_FOLIO', '0');
  pItem.putValue('APL_ACTIVITEEFFECT', '-');
  pItem.putValue('APL_DATEDEBREAL', vDtDeb);
  pItem.putValue('APL_DATEFINREAL', vDtFin);
  pItem.putValue('APL_LIGNEGENEREE', '-');
  pItem.putValue('APL_PUPR', '0');
  pItem.putValue('APL_PUPRDEV', '0');
  pItem.putValue('APL_PUVENTEHT', '0');
  pItem.putValue('APL_PUVENTEDEVHT', '0');
  pItem.putValue('APL_INITPTPR', '0');
  pItem.putValue('APL_INITPTPRDEV', '0');
  pItem.putValue('APL_INITPTVENTEHT', '0');
  pItem.putValue('APL_INITPTVTDEVHT', '0');
  pItem.putValue('APL_REALPTPR', '0');
  pItem.putValue('APL_REALPTPRDEV', '0');
  pItem.putValue('APL_REALPTVENTEHT', '0');
  pItem.putValue('APL_REALPTVTDEVHT', '0');
  pItem.putValue('APL_ACTIVITEREPRIS', '');
  pItem.putValue('APL_CREATEUR', v_pgi.user);
  pItem.putValue('APL_DATECREATION', now);
  pItem.putValue('APL_UTILISATEUR', v_pgi.user);
  pItem.putValue('APL_DATEMODIF', now);
  pItem.putValue('APL_LIBRETACHE1', '');
  pItem.putValue('APL_LIBRETACHE2', '');
  pItem.putValue('APL_LIBRETACHE3', '');
  pItem.putValue('APL_LIBRETACHE4', '');
  pItem.putValue('APL_LIBRETACHE5', '');
  pItem.putValue('APL_LIBRETACHE6', '');
  pItem.putValue('APL_LIBRETACHE7', '');
  pItem.putValue('APL_LIBRETACHE8', '');
  pItem.putValue('APL_LIBRETACHE9', '');
  pItem.putValue('APL_LIBRETACHEA', '');
  pItem.putValue('APL_BOOLLIBRE1', '');
  pItem.putValue('APL_BOOLLIBRE2', '');
  pItem.putValue('APL_BOOLLIBRE3', '');
  pItem.putValue('APL_DATELIBRE1', now);
  pItem.putValue('APL_DATELIBRE2', now);
  pItem.putValue('APL_DATELIBRE3', now);
  pItem.putValue('APL_CHARLIBRE1', '');
  pItem.putValue('APL_CHARLIBRE2', '');
  pItem.putValue('APL_CHARLIBRE3', '');
  pItem.putValue('APL_VALLIBRE1', '0');
  pItem.putValue('APL_VALLIBRE2', '0');
  pItem.putValue('APL_VALLIBRE3', '0');
  pItem.putValue('APL_ACTIVITEREPRIS', '-');
  pItem.putValue('APL_BLOQUE', '-');
  pItem.putValue('APL_ESTFACTURE', '-');
  pItem.putValue('APL_LIBELLEPLA', '');
  pItem.putValue('ATA_DESCRIPTIF', '');
  pItem.putValue('ATA_IDENTLIGNE', '');
  pItem.putValue('ATA_QTEINITIALE', '');
  pItem.putValue('ATA_TYPEPLANIF', '');
  pItem.putValue('ATA_TYPEADRESSE', '');
  pItem.putValue('ATA_NUMEROADRESSE', '');
  pItem.putValue('ATA_NBPARTINSCRIT', '0');
  pItem.putValue('ATA_NBPARTPREVU', '0');
  pItem.putValue('APL_TYPEADRESSE', '');
  pItem.putValue('APL_NUMEROADRESSE', '0');
  pItem.putValue('APL_STATUTPLA', '');
  pItem.putValue('T_LIBELLE', '');

  if fBoDemiJournee then
  begin
    // matin
    if (vStDebJ = 'MAT') and
       (vStFinJ = 'MAT') then
    begin
      pItem.putValue('APL_HEUREDEB_PLA', vDtDeb + getparamsoc('SO_AFAMDEBUT'));
      pItem.putValue('APL_HEUREFIN_PLA', vDtFin + getparamsoc('SO_AFAMDEBUT'));
    end
    // apres midi
    else if (vStDebJ = 'PAM') and
            (vStFinJ = 'PAM') then
    begin
      pItem.putValue('APL_HEUREDEB_PLA', vDtDeb + getparamsoc('SO_AFPMDEBUT'));
      pItem.putValue('APL_HEUREFIN_PLA', vDtFin + getparamsoc('SO_AFPMDEBUT'));
    end
    // apres-midi / matin
    else if (vStDebJ = 'PAM') and
            (vStFinJ = 'MAT') then
    begin
      pItem.putValue('APL_HEUREDEB_PLA', vDtDeb + getparamsoc('SO_AFPMDEBUT'));
      pItem.putValue('APL_HEUREFIN_PLA', vDtFin + getparamsoc('SO_AFAMDEBUT'));
    end
    // journée
    else
    begin
      pItem.putValue('APL_HEUREDEB_PLA', vDtDeb + getparamsoc('SO_AFAMDEBUT'));
      pItem.putValue('APL_HEUREFIN_PLA', vDtFin + getparamsoc('SO_AFPMDEBUT'));
    end;
  end;

  pItem.putValue('APL_ARTICLE', '0');
  pItem.putValue('APL_PUPR', '0');
  pItem.putValue('APL_PUPRDEV', '0');
  pItem.putValue('APL_PUVENTEHT', '0');
  pItem.putValue('APL_PUVENTEDEVHT', '0');
  pItem.putValue('APL_INITPTPR', '0');
  pItem.putValue('APL_INITPTPRDEV', '0');
  pItem.putValue('APL_INITPTVENTEHT', '0');
  pItem.putValue('APL_INITPTVTDEVHT', '0');
  pItem.putValue('APL_REALPTPR', '0');
  pItem.putValue('APL_REALPTPRDEV', '0');
  pItem.putValue('APL_REALPTVENTEHT', '0');
  pItem.putValue('APL_REALPTVTDEVHT', '0');

  // libellé absence
  pItem.putValue('CHAMPLIBELLE', FormatLibTache(vStRtt, vStValider));
  pItem.putValue('ATA_LIBELLETACHE1', traduireMemoire(texteMessage[66])); //'ABSENCE'
  DisplayHint(pItem, traduireMemoire(texteMessage[66])); // 'ABSENCE'
end;

procedure TPlanningGA.LoadAbsencesPlanningRessource(pStWhere: string);
var
  vSt: string;
  i: Integer;
  vTobItem: Tob;
  vInNumTache: Integer;
  vInNumLigne: Integer;
  vTobAbsence : Tob;
  vTobAbsences : Tob;
  vTobRes : Tob;
  vTob : Tob;

begin

  {vTobAbsences := TOB.Create('Mes Absences', nil, -1);
  vTobRes := TOB.Create('#RESSOURCE', nil, -1);
  try
    vInNumTache := -1;
    vInNumLigne := 1000;

    if pos('APL_RESSOURCE', pStWhere) <> 0 then
      ReplaceSubStr(pStWhere, 'APL_RESSOURCE', 'ARS_RESSOURCE');

    vSt := MakeRequetePaie(pStWhere);
    vTobAbsences.LoadDetailFromSql(vSt);

    // Selection des ressources
    vSt := 'SELECT ARS_RESSOURCE, ARS_LIBELLE FROM RESSOURCE WHERE ARS_RESSOURCE <> "" ';
    vSt := vSt + pStWhere;

    vTobRes.LoadDetailFromSql(vSt);

    // parcours des ressources et création des jours d'absences
    for i := 0 to vTobRes.detail.count - 1 do
    begin
      vTob := fTobsPlanningTobRes.FindFirst(['ATR_RESSOURCE'], [vTobRes.detail[i].GetValue('ARS_RESSOURCE')], true);
      if vTob <> nil then
      begin
			  // ajout d'un item absence
        vTobAbsence := vTobAbsences.FindFirst(['ARS_RESSOURCE'], [vTobRes.detail[i].GetValue('ARS_RESSOURCE')], true);
        while vTobAbsence <> nil do
        begin
          vTobItem := Tob.create('AFPLANNING', fTobsPlanningTobItems, -1);
          CreateAbsence(vTobItem, vTobAbsence,
                        vTobRes.detail[i].GetValue('ARS_RESSOURCE') ,
                        vTobRes.detail[i].GetValue('ARS_RESSOURCE'),
                        i, vInNumTache, vInNumligne);
          vTobAbsence := vTobAbsences.FindNext(['ARS_RESSOURCE'], [vTobRes.detail[i].GetValue('ARS_RESSOURCE')], true);
          DisplayHint(vTobItem, TexteMessage[66]); //'ABSENCE'
        end;
      end;
    end;
  finally
    formatTobRes;
    FreeAndNil(vTobRes);
    FreeAndNil(vTobAbsences);
  end;}
  
end;

function TPlanningGA.AffaireTermine(pStAffaire : String) : Boolean;
begin
  result := existeSql('SELECT * FROM AFFAIRE WHERE AFF_AFFAIRE = "' + pStAffaire + '" AND AFF_ETATAFFAIRE = "CLO"');
end;

function TPlanningGA.GetStatutItem(pStStatutTache : String) : String;
begin
  if pStStatutTache <> '' then
      result := pStStatutTache
  else if fStStatutDefaut <> '' then
    result := fStStatutDefaut
  else
    result := '';
end;

{------------------------------------------------------------------------------}
{ TAF_Planning                                                                 }
{------------------------------------------------------------------------------}
function TAF_Planning.GetPlanningJour: TPlanningGA;
begin
  if assigned(FPlanningJour) then
    result := FPlanningJour
  else
  begin
    case fInNumPlanning of
      cInPlanning1: FPlanningJour := TPlanningGA1.Create(Self);
      cInPlanning2: FPlanningJour := TPlanningGA2.Create(Self);
      cInPlanning3: FPlanningJour := TPlanningGA3.Create(Self);
      cInPlanning4: FPlanningJour := TPlanningGA4.Create(Self);
      cInPlanning5: FPlanningJour := TPlanningGA5.Create(Self);
      cInPlanning6: FPlanningJour := TPlanningGA6.Create(Self);
      cInPlanning7: FPlanningJour := TPlanningGA7.Create(Self);
    end;
    result := FPlanningJour;
  end;
end;

{------------------------------------------------------------------------------}
{ OBJET TPlanningGA1 : Taches / Affaire, fonction, ressource                   }
{------------------------------------------------------------------------------}
{ CLE : AFFAIRE / FONCTION / RESSOURCE                                         }
{------------------------------------------------------------------------------}

procedure TPlanningGA1.SelectPlanning;
begin

  fColonne1.StField := 'CHAMP1';
  fColonne2.StField := 'CHAMP2';

  fColonne1.StSize := '70';
  fColonne3.StSize := '50';

  fColonne1.StAlign := 'L';
  fColonne2.StAlign := 'L';
  fColonne3.StAlign := 'L';

  fColonne1.StEntete := TraduitGA(TexteMessage[67]);

  if getParamSocSecur('SO_AFFONCTION', False) then
  begin
    fColonne3.StField := 'CHAMP3';
    fColonne3.StSize := '50';
    fColonne2.StEntete := TraduitGA('Fonction');
    fColonne3.StEntete := TraduitGA('Ressource');
  end
  else
  begin
    fColonne3.StField := '';
    fColonne3.StSize := '0';
    fColonne2.StEntete := TraduitGA('Ressource');
    fColonne3.StEntete := '';
  end;
end;

// chargement de toutes les lignes de planning et de toutes les ressources
procedure TPlanningGA1.LoadItemsPlanning(pStWhere: string);
var
  i: Integer;
  vStCle: string;

begin

  inherited;

  {for i := 0 to fTobsPlanningTobItems.detail.count - 1 do
  begin
    DisplayHint(fTobsPlanningTobItems.detail[i]);

    if getParamSocSecur('SO_AFFONCTION', False) then
      vStCle := fTobsPlanningTobItems.detail[i].GetValue('APL_AFFAIRE') + cStSep +
        fTobsPlanningTobItems.detail[i].GetValue('APL_FONCTION') + cStSep +
        fTobsPlanningTobItems.detail[i].GetValue('APL_RESSOURCE')
    else
      vStCle := fTobsPlanningTobItems.detail[i].GetValue('APL_AFFAIRE') + cStSep +
        fTobsPlanningTobItems.detail[i].GetValue('APL_RESSOURCE');

    fTobsPlanningTobItems.detail[i].putValue('I_CODE', vStCle);

    fTobsPlanningTobItems.detail[i].putValue('CODEPOSTAL',
      GetCodePostal(fTobsPlanningTobItems.detail[i],
      fTobAffAdm,
      fTobsPlanningTobItems.detail[i].GetString('APL_AFFAIRE'),
      fTobsPlanningTobItems.detail[i].GetString('APL_TIERS'),
      fTobsPlanningTobItems.detail[i].GetString('ATA_TYPEADRESSE'),
      fTobsPlanningTobItems.detail[i].GetString('ATA_NUMEROADRESSE'),
      fTobsPlanningTobItems.detail[i].GetString('APL_TYPEADRESSE'),
      fTobsPlanningTobItems.detail[i].GetString('APL_NUMEROADRESSE'), True, True));

    // on n'affiche ni la ressource ni le client dans l'item
    FormatChampLibelle(fTobsPlanningTobItems.detail[i]);
    IconeCEGID(i);
  end;}
  
end;

function TPlanningGA1.LoadResPlanning(pStWhere: string): Boolean;
var
  Q: TQuery;
  S: string;
  vStWhere: string;

begin

  {if getParamSocSecur('SO_AFFONCTION', False) then
  begin
    S := ' SELECT DISTINCT AFF_LIBELLE, ATR_AFFAIRE, ATR_TIERS, ATR_FONCTION, ATR_RESSOURCE, ';
    S := S + ' AFO_LIBELLE AS CHAMP2, ARS_LIBELLE || " " || ARS_LIBELLE2 AS CHAMP3 ';
    S := S + ' FROM AFFAIRE, PIECE, RESSOURCE, TACHERESSOURCE LEFT OUTER JOIN FONCTION ';
    S := S + ' ON (ATR_FONCTION = AFO_FONCTION) ';
    S := S + ' WHERE AFF_AFFAIRE = ATR_AFFAIRE AND ATR_RESSOURCE = ARS_RESSOURCE ';
    S := S + ' AND AFF_AFFAIRE = GP_AFFAIRE ';
    S := S + ' AND GP_NATUREPIECEG = AFF_NATUREPIECEG';
  end
  else
  begin
    S := ' SELECT DISTINCT AFF_LIBELLE, ATR_AFFAIRE, ATR_TIERS, ATR_FONCTION, ATR_RESSOURCE, ';
    S := S + ' ARS_LIBELLE || " " || ARS_LIBELLE2 AS CHAMP2';
    S := S + ' FROM AFFAIRE, PIECE, RESSOURCE, TACHERESSOURCE ';
    S := S + ' WHERE AFF_AFFAIRE = ATR_AFFAIRE AND ATR_RESSOURCE = ARS_RESSOURCE ';
    S := S + ' AND AFF_AFFAIRE = GP_AFFAIRE ';
    S := S + ' AND GP_NATUREPIECEG = AFF_NATUREPIECEG';
  end;

  // C.B 04/10/2005
  // on affiche le planning des clients fermés
  //S := S + ' AND ARS_FERME = "-"';
  // C.B 11/08/2003
  // tous les plannings doivent etre vu, meme ceux des ressources inactives
  //      S := S + ' AND ATR_STATUTRES <> "INA"';

  vStWhere := pStWhere;
  if pos('APL_AFFAIRE', vStWhere) <> 0 then
    ReplaceSubStr(vStWhere, 'APL_AFFAIRE', 'ATR_AFFAIRE');

  if pos('APL_RESSOURCE', vStWhere) <> 0 then
    ReplaceSubStr(vStWhere, 'APL_RESSOURCE', 'ATR_RESSOURCE');

  if vStWhere <> '' then
    S := S + vStWhere;

  if getParamSocSecur('SO_AFFONCTION', False) then
    S := S + ' ORDER BY ATR_AFFAIRE, ATR_FONCTION, ATR_RESSOURCE'
  else
    S := S + ' ORDER BY ATR_AFFAIRE, ATR_RESSOURCE';

  Q := nil;
  try
    Q := OpenSql(S, True);
    if not Q.Eof then
    begin
      fTobsPlanningTobRes.ClearDetail;
      fTobsPlanningTobRes.LoadDetailDB('les_Ressources', '', '', Q, False, True);
      fTobsPlanningTobRes.detail[0].addchampSup('CLE', true);
      fTobsPlanningTobRes.detail[0].addchampSup('CHAMP1', true);
      formatTobRes;
      result := true;
    end
    else
    begin
      result := false;
    end;

  finally
    if Q <> nil then
      Ferme(Q);
  end;}
  
end;

procedure TPlanningGA1.AjoutTacheAbsence(pInIndex: Integer; pStRes, pStResLib: string);
begin

end;

procedure TPlanningGA1.AjoutNonProductive(pInIndex: Integer; pStRes, pStResLib: string);
begin
  // pas de tâches non productives dans ce planning
end;

procedure TPlanningGA1.LoadAbsencesPlanning(pStWhere: string);
begin
  // pad d'abscences
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 08/03/2002
Modifié le ... :
Description .. : formate l'affichage de la tob des ressources
Suite ........ :
Mots clefs ... :
*****************************************************************}

procedure TPlanningGA1.formatTobRes;
var
  i						: Integer;
  vStCle			: String;
  vStAffaire 	: String;

begin

  for i := 0 to fTobsPlanningTobRes.detail.count - 1 do
  begin
    // creation de la cle
    vStCle := fTobsPlanningTobRes.detail[i].GetValue('ATR_AFFAIRE');

    if getParamSocSecur('SO_AFFONCTION', False) then
    begin
      if fTobsPlanningTobRes.detail[i].GetValue('ATR_FONCTION') <> NULL then
        vStCle := vStCle + cStSep + fTobsPlanningTobRes.detail[i].GetValue('ATR_FONCTION')
      else
        vStCle := vStCle + cStSep;
    end;

    if fTobsPlanningTobRes.detail[i].GetValue('ATR_RESSOURCE') <> NULL then
      vStCle := vStCle + cStSep + fTobsPlanningTobRes.detail[i].GetValue('ATR_RESSOURCE')
    else
      vStCle := vStCle + cStSep;

    fTobsPlanningTobRes.detail[i].PutValue('CLE', vStCle);
 
    // C.B 18/05/2006
    if GetParamsocSecur('SO_AFAFFICHAFFAIRE', False) then
      vStAffaire := CodeAffaireAffiche(fTobsPlanningTobRes.detail[i].GetValue('ATR_AFFAIRE'))
    else
      vStAffaire := fTobsPlanningTobRes.detail[i].GetValue('AFF_LIBELLE');

    fTobsPlanningTobRes.detail[i].PutValue('CHAMP1', fTobsPlanningTobRes.detail[i].GetValue('ATR_TIERS') + #13#10 +
        RechDom('GCTIERSCLI', fTobsPlanningTobRes.detail[i].GetValue('ATR_TIERS'), False) + #13#10 +
        vStAffaire);

    // supression des null des ressources
    if getParamSocSecur('SO_AFFONCTION', False) then
      if (fTobsPlanningTobRes.detail[i].GetValue('CHAMP3') = NULL) then
        fTobsPlanningTobRes.detail[i].putvalue('CHAMP3', '')
			// attendre V7
		  else if not (pos(#13#10, fTobsPlanningTobRes.detail[i].GetValue('CHAMP3')) > 0) then
				fTobsPlanningTobRes.detail[i].PutValue('CHAMP3', fTobsPlanningTobRes.detail[i].GetValue('CHAMP3') + #13#10 + fTobsPlanningTobRes.detail[i].GetValue('ATR_RESSOURCE'));

    // supression des null des fonctions
    // ou des ressources si pas de fonction
    if (fTobsPlanningTobRes.detail[i].GetValue('CHAMP2') = NULL) then
      fTobsPlanningTobRes.detail[i].putvalue('CHAMP2', '')
		// attendre V7                                                            
    else if (not GetparamSoc('SO_AFFONCTION')) and
    			  (not (pos(#13#10, fTobsPlanningTobRes.detail[i].GetValue('CHAMP2')) > 0)) then
  		fTobsPlanningTobRes.detail[i].PutValue('CHAMP2', fTobsPlanningTobRes.detail[i].GetValue('CHAMP2') + #13#10 + fTobsPlanningTobRes.detail[i].GetValue('ATR_RESSOURCE'));

  end;

  // construction de l'arborescence
  // on n'affiche pas 2 fois la meme affaire
  for i := fTobsPlanningTobRes.detail.count - 1 downto 1 do
  begin
    if fTobsPlanningTobRes.detail[i].GetValue('CHAMP1') = fTobsPlanningTobRes.detail[i - 1].GetValue('CHAMP1') then
      fTobsPlanningTobRes.detail[i].PutValue('CHAMP1', '');
  end;

end;

// affaire, fonction, ressource
procedure TPlanningGA1.DecodeCle(pStCle: string);
var
  vSt: string;
begin
  vSt := pStCle;
  fStAffaire := (Trim(ReadTokenPipe(vSt, cStSep)));

  if getParamSocSecur('SO_AFFONCTION', False) then
    fStFonction := (Trim(ReadTokenPipe(vSt, cStSep)));

  fStRes := (Trim(ReadTokenPipe(vSt, cStSep)));
  fStTiers := GetTiers(fStAffaire);

end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 28/03/2002
Modifié le ... :
Description .. : création d'une ligne de planning et affichage à
Suite ........ : l'écran
Mots clefs ... :
*****************************************************************}
procedure TPlanningGA1.CreateItem(Sender: TObject; Item: TOB; var Cancel: boolean);
begin

  cancel := False;
  DecodeCle(item.getvalue('I_CODE'));

  // création directe = sélection d'une ligne effectuée
  if fInModeCreation = cInCreationDirecte then
  begin
    // sélection possible d'affaire puis de ressource
    if (fStRes <> fStResSel) then
    begin
      if (PGIAskAF(TraduitGA(TexteMessage[19]), '') = mrYes) then
      begin
        fInModeCreation := cInCreationSansSelection;
        FFichePlanning.La_Selection.caption := '';
        FFichePlanning.BT_Lignes.down := False;
      end
      else
        cancel := True;
    end;
  end

  // creation sans selection de ressource
  else
    if (fInModeCreation = cInCreationAvecSelection) then
  begin
    if (fStAffaireSel <> '') and (fStAffaire <> fStAffaireSel) then
    begin
      Item.PutValue('APL_AFFAIRE', fStAffaire);
      Item.PutValue('APL_TIERS', fStTiers);
      Item.PutValue('APL_RESSOURCE', fStRes);
      if getParamSocSecur('SO_AFFONCTION', False) then
        Item.PutValue('APL_FONCTION', fStFonction);
    end;
  end;

  if not cancel then
  begin
    fTobCouranteMere := MaTobInterne('AFFAIRECOURANTE');
    fTobCourante := TOB.Create('AFFAIRECOURANTE', fTobCouranteMere, -1);
    fTobCourante.Dupliquer(Item, TRUE, TRUE, TRUE);

    inherited;
    FormatChampLibelle(Item);
  end;
end;

procedure TPlanningGA1.CopyItem(Sender: TObject; Item: TOB; var Cancel: boolean);
var
  vStRetour   : string;
  vStAffaire  : string;
  vSt         : string;

begin

  // C.B 26/09/2006
  // suppression du spécifique cegid
  {if IsStatutModifiableTob(Item.GetValue('APL_ETATLIGNE'), fTobsPlanningTobEtats) then
    Item.putValue('APL_ETATLIGNE', GetStatutItem(Item.GetString('ATA_STATUTPLA')));

  // C.B 30/10/2006
  // correction du déplacement des ressources
  DecodeCle(Item.GetValue('I_CODE'));
  Item.PutValue('APL_RESSOURCE', fStRes);

  //mcd 14/03/2005
  if not JaiLeDroitRessource(Item.getValue('APL_RESSOURCE')) then
  begin
    PgiBoxAf(TexteMessage[24], TexteMessage[59]);
    cancel := true;
  end

  // test si on duplique sur le même emplacement
  else if not AutoriserPlanif(Item, fBoDemiJournee, False, vStAffaire, 'DUPLICATION') then
    cancel := true

  // C.B 31/10/2006
  // test si on est a cheval sur 2 semaines ou 2 mois
  else if GetParamSocSecur('SO_AFPARSEMAINE', True) and
     EstAChevalSemaine(Item, fBoDemiJournee) then
  begin
    PGIBoxAF(TexteMessage[70], TexteMessage[59]); // Vous ne pouvez pas créer, déplacer, réduire ou étirer une intervention à cheval sur 2 semaines.
    cancel := True;
  end

  // C.B 06/04/2006
  // test si on est a cheval sur 2 semaines ou 2 mois
  else if GetParamSocSecur('SO_AFPARMOIS', True) and
     EstAChevalMois(Item, fBoDemiJournee) then
  begin
    PGIBoxAF(TexteMessage[71], TexteMessage[59]); // Vous ne pouvez pas créer, déplacer, réduire ou étirer une intervention à cheval sur 2 semaines.
    cancel := True;
  end

  // Recherche de l'existant dans la table tacheessource
  else if ChangeLigneItemOk(Item) and CreateTacheRessource(Item, true) then
  begin
    decodeCle(Item.GetValue('I_CODE'));
    Item.PutValue('APL_AFFAIRE', fStAffaire);
    Item.PutValue('APL_RESSOURCE', fStRes);
    if getParamSocSecur('SO_AFFONCTION', False) then
      Item.PutValue('APL_FONCTION', fStFonction);

    fTobCouranteMere := MaTobInterne('AFFAIRECOURANTE');
    fTobCourante := TOB.Create('AFFAIRECOURANTE', fTobCouranteMere, -1);
    fTobCourante.Dupliquer(Item, TRUE, TRUE, TRUE);

    if fBoDemiJournee then
      vSt := ';DEMIJOURNEE'
    else
      vSt := '';
    vStRetour := AFLanceFicheAFPlanning('', 'ACTION=CREATION;DUPLIQUER' + vSt);

    if (vStRetour = 'CANCEL') or (vStRetour = '') then
      cancel := true
    else
    begin

      // C.B 31/10/2006
      // si on n'affiche pas les samedis et les dimanches et qu'on
      // est à cheval sur 2 semaines, on averti que les samedis et les
      // dimanches sont comptés
      if (not GetParamSocSecur('SO_AFPLANNINGWE', true)) and
         EstAChevalSemaine(Item, fBoDemiJournee) then
      begin
        PGIBoxAF(TexteMessage[20], TexteMessage[59]); // Attention, Les samedis et dimanches ne sont pas visibles, mais ils sont planifiés.
      end;

      Item.dupliquer(fTobCourante, True, True, True);
      RafraichirItem(Item, vStRetour, 'ARTICLE', False, False);
      cancel := false;
    end;

    DetruitMaTobInterne('AFFAIRECOURANTE');
  end
  else
    cancel := true;}
    
end;

procedure TPlanningGA1.ModifyItem(Sender: TObject; Item: TOB; var Cancel: boolean);
var
  vStItem: string;
  vStCle: string;

begin
  inherited;
  if item <> nil then
  begin
    vStItem := AFLanceFicheAFPlanning(varAsType(Item.GetValue('APL_AFFAIRE'), varString) + ';' + varAsType(Item.GetValue('APL_NUMEROLIGNE'), varString), fStAction);

    if (vStItem = 'CANCEL') or (vStItem = '') then
      cancel := true
    else
    begin                                        
      // si changement de cle
      if getParamSocSecur('SO_AFFONCTION', False) then
        EncodeCle(Item.GetValue('APL_AFFAIRE'), Item.GetValue('APL_FONCTION'), Item.GetValue('APL_RESSOURCE'), vStCle)
      else
        EncodeCle(Item.GetValue('APL_AFFAIRE'), Item.GetValue('APL_RESSOURCE'), '', vStCle);

      Item.putvalue('I_CODE', vStCle);
      RafraichirItem(Item, vStItem, 'ARTICLE', False, False);
      cancel := false;
    end;
    DetruitMaTobInterne('AFFAIRECOURANTE');
  end;
end;

procedure TPlanningGA1.BParamClick;
begin
  //AFLanceFiche_PlanningParam('P1J');
end;

function TPlanningGA1.AfficherTiers : Boolean;
begin
  result := False;
end;

function TPlanningGA1.AfficherRess  : Boolean;
begin
  result := False;
end;

function TPlanningGA1.ModeAffichage : String;
begin
 result := 'ARTICLE';
end;

function TPlanningGA1.ChangeLigneItemOk(Item: TOB): boolean;
begin
  result := true;
  decodeCle(Item.GetValue('I_CODE'));

  if (Item.GetValue('APL_AFFAIRE') = '') then
  begin
    PGIBoxAF(TexteMessage[6], TexteMessage[59]);
    result := false;
  end
  else
    if (fStAffaire <> Item.GetValue('APL_AFFAIRE')) then
  begin
    PGIBoxAF(TexteMessage[5], TexteMessage[59]);
    result := false;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 12/04/2002
Modifié le ... :
Description .. : Mise a jour des items déplacés dans le planning
Suite ........ :
Mots clefs ... :
*****************************************************************}

function TPlanningGA1.UpdateItem(FromItem, ToItem: TOB; vStAction: string): Boolean;
var
  vBoOk: Boolean;
  vStRetour: string;

begin

  {vBoOk := True;
  if (vStAction <> 'POPUP') and
     (vStAction <> 'MOVE') then
  begin
    if PlusDeJours(FromItem, ToItem, fBoDemiJournee) then
	  begin
  	  vBoOk := (PGIAskAF(TexteMessage[13], '') = mrYes);
    	if vBoOk then vBoOk := (PGIAskAF(TexteMessage[14], '') = mrYes);
			if vBoOk and VH_GC.GCIfDefCEGID then
  	  	RevalorisationCegid(FromItem, ToItem, vStAction);
	  end
    // dans le cas de reduce, on execute systématiquement la revalorisation
    // car on est pas dans le cas plus de jours
    else if (vStAction = 'REDUCE') and VH_GC.GCIfDefCEGID then
    	RevalorisationCegid(FromItem, ToItem, vStAction);
	end;

  if not vBoOk then
    result := false
  else
  begin
    if ChangeLigneItemOk(ToItem) and CreateTacheRessource(ToItem, True) then
    begin
      decodeCle(ToItem.GetValue('I_CODE'));
      ToItem.PutValue('APL_RESSOURCE', fStRes);
      ToItem.PutValue('APL_UTILISATEUR', V_PGI.User);

      DecodeCle(ToItem.getvalue('I_CODE'));
      UpdateData(FromItem, ToItem);

      // C.B 30/11/2006 fonction commune
      result := UpdateAfPlanning(ToItem, vStRetour);
    end
    else
      result := false;
  end;}
end;

{
      vStQR := 'UPDATE AFPLANNING SET APL_DATEDEBPLA = "' + UsDateTime(ToItem.GetValue('APL_DATEDEBPLA'));
      vStQR := vStQR + '", APL_DATEFINPLA = "' + UsDateTime(ToItem.GetValue('APL_DATEFINPLA'));
      vStQR := vStQR + '", APL_HEUREDEB_PLA = "' + UsTime(ToItem.GetValue('APL_HEUREDEB_PLA'));
      vStQR := vStQR + '", APL_HEUREFIN_PLA = "' + UsTime(ToItem.GetValue('APL_HEUREFIN_PLA'));
      vStQR := vStQR + '", APL_DATEDEBREAL = "' + UsDateTime(ToItem.GetValue('APL_DATEDEBREAL'));
      vStQR := vStQR + '", APL_DATEFINREAL = "' + UsDateTime(ToItem.GetValue('APL_DATEFINREAL'));
      vStQR := vStQR + '", APL_HEUREDEB_REAL = "' + UsTime(ToItem.GetValue('APL_HEUREDEB_REAL'));
      vStQR := vStQR + '", APL_HEUREFIN_REAL = "' + UsTime(ToItem.GetValue('APL_HEUREFIN_REAL'));

      vStQR := vStQR + '", APL_DATEMODIF = "' + usTime(now) ;
      vStQR := vStQR + '", APL_UTILISATEUR = "' + V_PGI.User;

      vStQR := vStQR + '", APL_RESSOURCE = "' + fStRes;
      vStQR := vStQR + '", APL_ETATLIGNE = "' + ToItem.GetValue('APL_ETATLIGNE');
      vStQR := vStQR + '", APL_FONCTION = "' + ToItem.GetValue('APL_FONCTION');
      vStQR := vStQR + '", APL_QTEPLANIFIEE = ' + variantToSql(ToItem.GetValue('APL_QTEPLANIFIEE'));
      vStQR := vStQR + ', APL_QTEREALISE = ' + variantToSql(ToItem.GetValue('APL_QTEREALISE'));

      vStQR := vStQR + ', APL_QTEPLANIFUREF = ' + variantToSql(AFConversionUnite(ToItem.GetValue('APL_UNITETEMPS'), getParamSocSecur('SO_AFMESUREACTIVITE',''),
        valeur(ToItem.GetValue('APL_QTEPLANIFIEE'))));
      vStQR := vStQR + ', APL_QTEREALUREF = ' + variantToSql(AFConversionUnite(ToItem.GetValue('APL_UNITETEMPS'), getParamSocSecur('SO_AFMESUREACTIVITE',''),
        valeur(ToItem.GetValue('APL_QTEREALISE'))));

      vStQR := vStQR + ', APL_PUPR = ' + variantToSql(ToItem.GetValue('APL_PUPR'));
      vStQR := vStQR + ', APL_PUPRDEV = ' + variantToSql(ToItem.GetValue('APL_PUPRDEV'));
      vStQR := vStQR + ', APL_INITPTPR = ' + variantToSql(ToItem.GetValue('APL_INITPTPR'));
      vStQR := vStQR + ', APL_INITPTPRDEV = ' + variantToSql(ToItem.GetValue('APL_INITPTPRDEV'));
      vStQR := vStQR + ', APL_REALPTPR = ' + variantToSql(ToItem.GetValue('APL_REALPTPR'));
      vStQR := vStQR + ', APL_REALPTPRDEV = ' + variantToSql(ToItem.GetValue('APL_REALPTPRDEV'));

      vStQR := vStQR + ', APL_PUVENTEHT = ' + variantToSql(ToItem.GetValue('APL_PUVENTEHT'));
      vStQR := vStQR + ', APL_PUVENTEDEVHT = ' + variantToSql(ToItem.GetValue('APL_PUVENTEDEVHT'));
      vStQR := vStQR + ', APL_INITPTVENTEHT = ' + variantToSql(ToItem.GetValue('APL_INITPTVENTEHT'));
      vStQR := vStQR + ', APL_INITPTVTDEVHT = ' + variantToSql(ToItem.GetValue('APL_INITPTVTDEVHT'));
      vStQR := vStQR + ', APL_REALPTVENTEHT = ' + variantToSql(ToItem.GetValue('APL_REALPTVENTEHT'));
      vStQR := vStQR + ', APL_REALPTVTDEVHT = ' + variantToSql(ToItem.GetValue('APL_REALPTVTDEVHT'));

      vStQR := vStQR + ' WHERE APL_AFFAIRE = "' + ToItem.GetValue('APL_AFFAIRE');
      vStQR := vStQR + '" AND APL_NUMEROLIGNE = ' + varAsType(ToItem.GetValue('APL_NUMEROLIGNE'), varString);
      vStQR := vStQR + ' AND APL_TYPEPLA = "PLA"';

      try
        if (ExecuteSql(vStQR) = 1) then
        begin
          result := true;
          RafraichirItemPartiel(ToItem, vStRetour, 'ARTICLE');
          // mise à jour de la quantité planifiée de la tache
          UpdatePlanifieTache(ToItem.GetString('APL_AFFAIRE'), ToItem.GetString('APL_NUMEROTACHE'), ToItem.GetString('APL_UNITETEMPS'));
        end;
      except
        PGIBoxAF(TexteMessage[1], TexteMessage[59]);
        result := false;
      end;}

procedure TPlanningGA1.InitPlanning;
begin
  inherited;
  // les changements d'états sont disponibles en fonction des autorisations
  LoadPopUpMenu(False);
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 01/08/2002
Modifié le ... :   /  /
Description .. : Test si la Ressource existe dans la table tacheRessource
Mots clefs ... : Si non, on propose la creation
*****************************************************************}

function TPlanningGA1.CreateTacheRessource(Item: Tob; pBoForce: Boolean): boolean;
var
  vTob: Tob;
  vStAff0: string;
  vStAff1: string;
  vStAff2: string;
  vStAff3: string;
  vStAvenant: string;

begin

  result := false;
  decodeCle(Item.GetValue('I_CODE'));

  if not ExisteTacheRessource(iTem.GetString('APL_AFFAIRE'), intToStr(item.getValue('APL_NUMEROTACHE')), fStRes) then
  begin
    if pBoForce or (PGIAskAF(TraduitGA(TexteMessage[7]), '') = mrYes) then
    begin

      vTob := Tob.Create('TACHERESSOURCE', nil, -1);
      try
        vTob.PutValue('ATR_NUMEROTACHE', intToStr(item.getValue('APL_NUMEROTACHE')));
        vTob.PutValue('ATR_TIERS', item.GetValue('APL_TIERS'));
        vTob.PutValue('ATR_AFFAIRE', iTem.GetString('APL_AFFAIRE'));
        CodeAffaireDecoupe(iTem.GetString('APL_AFFAIRE'), vStAff0, vStAff1, vStAff2, vStAff3, vStAvenant, taModif, false);
        vTob.PutValue('ATR_AFFAIRE0', vStAff0);
        vTob.PutValue('ATR_AFFAIRE1', vStAff1);
        vTob.PutValue('ATR_AFFAIRE2', vStAff2);
        vTob.PutValue('ATR_AFFAIRE3', vStAff3);
        vTob.PutValue('ATR_AVENANT', vStAvenant);
        vTob.PutValue('ATR_RESSOURCE', fStRes);
        vTob.PutValue('ATR_FONCTION', fStFonction);
        vTob.PutValue('ATR_STATUTRES', 'ACT');
        vTob.PutValue('ATR_DEVISE', V_PGI.DevisePivot);

        if not vTob.InsertDB(nil) then
          PGIBoxAF(TexteMessage[1], TexteMessage[59])
        else
          result := true;
      finally
        vTob.Free;
      end;
    end;
  end
  else
    result := true;
end;

{------------------------------------------------------------------------------}
{ OBJET TPlanningGA2 : Taches / ressource, Affaire                             }
{------------------------------------------------------------------------------}
{ CLE : RESSOURCE / AFFAIRE                                                    }
{------------------------------------------------------------------------------}
procedure TPlanningGA2.SelectPlanning;
begin

  fColonne1.StField := 'CHAMP1';
  fColonne2.StField := 'CHAMP2';
  fColonne3.StField := '';

  fColonne1.StSize := '50';
  fColonne2.StSize := '50';
  fColonne3.StSize := '0';

  fColonne1.StAlign := 'L';
  fColonne2.StAlign := 'L';
  fColonne3.StAlign := '';

  fColonne1.StEntete := TraduitGA('Ressource');
  fColonne2.StEntete := TraduitGA(TexteMessage[67]); //Tiers/Affaire
  fColonne3.StEntete := '';
end;
 
// chargement de toutes les lignes de planning et de toutes les ressources
procedure TPlanningGA2.LoadItemsPlanning(pStWhere: string);
var
  i: Integer;
  vStCle: string;
begin

  inherited;

  {for i := 0 to fTobsPlanningTobItems.detail.count - 1 do
  begin
    DisplayHint(fTobsPlanningTobItems.detail[i]);
    vStCle := fTobsPlanningTobItems.detail[i].GetValue('APL_RESSOURCE') + cStSep +
      fTobsPlanningTobItems.detail[i].GetValue('APL_AFFAIRE');
    fTobsPlanningTobItems.detail[i].putValue('I_CODE', vStCle);

    fTobsPlanningTobItems.detail[i].putValue('CODEPOSTAL',
    	GetCodePostal(fTobsPlanningTobItems.detail[i],
      fTobAffAdm,
      fTobsPlanningTobItems.detail[i].GetValue('APL_AFFAIRE'),
      fTobsPlanningTobItems.detail[i].GetValue('APL_TIERS'),
      fTobsPlanningTobItems.detail[i].GetString('ATA_TYPEADRESSE'),
      fTobsPlanningTobItems.detail[i].GetString('ATA_NUMEROADRESSE'),
      fTobsPlanningTobItems.detail[i].GetString('APL_TYPEADRESSE'),
      fTobsPlanningTobItems.detail[i].GetString('APL_NUMEROADRESSE'), True, True));

    FormatChampLibelle(fTobsPlanningTobItems.detail[i]);
    IconeCEGID(i);
  end;}
end;

function TPlanningGA2.LoadResPlanning(pStWhere: string): Boolean;
var
  Q: TQuery;
  S: string;
  vStWhere: string;
  vRessource: string;
  i_ind: Integer;
  vTobRes: TOB;
  vTob: TOB;

begin

  // C.B 18/05/2006
  {if GetParamsocSecur('SO_AFAFFICHAFFAIRE', False) then
  begin
    S := ' SELECT DISTINCT AFF_LIBELLE, ATR_AFFAIRE, ATR_RESSOURCE, ';
    S := S + ' ARS_LIBELLE || " " || ARS_LIBELLE2 AS CHAMP1, ATR_TIERS ';
    S := S + ' FROM RESSOURCE, TACHERESSOURCE, AFFAIRE ';
    S := S + ' WHERE ATR_RESSOURCE = ARS_RESSOURCE ';
    S := S + ' AND ATR_AFFAIRE = AFF_AFFAIRE ';
  end
  else
  begin
    S := ' SELECT DISTINCT "" AS AFF_LIBELLE, ATR_AFFAIRE, ATR_RESSOURCE, ';
    S := S + ' ARS_LIBELLE || " " || ARS_LIBELLE2 AS CHAMP1, ATR_TIERS ';
    S := S + ' FROM RESSOURCE, TACHERESSOURCE ';
    S := S + ' WHERE ATR_RESSOURCE = ARS_RESSOURCE ';
	end;

  vStWhere := pStWhere;

  if pos('APL_AFFAIRE', vStWhere) <> 0 then
    ReplaceSubStr(vStWhere, 'APL_AFFAIRE', 'ATR_AFFAIRE');

  if pos('APL_RESSOURCE', vStWhere) <> 0 then
    ReplaceSubStr(vStWhere, 'APL_RESSOURCE', 'ATR_RESSOURCE');

  if vStWhere <> '' then
    S := S + vStWhere;

  if fStAffNonProductive <> '' then // Journée non productive
    S := S + ' AND ATR_AFFAIRE <> "' + fStAffNonProductive + '"';

  S := S + ' GROUP BY ';

  if GetParamsocSecur('SO_AFAFFICHAFFAIRE', False) then
    S := S + ' AFF_LIBELLE, ';

	S := S + ' ATR_AFFAIRE, ATR_RESSOURCE, ARS_LIBELLE || " " || ARS_LIBELLE2, ATR_TIERS';
  S := S + ' ORDER BY ARS_LIBELLE || " " || ARS_LIBELLE2';

  Q := nil;
  try
    Q := OpenSql(S, True);
    if not Q.Eof then
    begin
      fTobsPlanningTobRes.ClearDetail;
      fTobsPlanningTobRes.LoadDetailDB('les_Ressources', '', '', Q, False, True);
      fTobsPlanningTobRes.detail[0].addchampSup('CLE', true);
      fTobsPlanningTobRes.detail[0].addchampSup('CHAMP2', true);
      result := true;
    end
    else
    begin
      result := false;
    end;

  finally
    if Q <> nil then
      Ferme(Q);
  end;

  // Journée non-productive
  if (not fBoGestionAbs) and (fStAffNonProductive <> '') and
    (fTobsPlanningTobRes <> nil) and (fTobsPlanningTobRes.detail.count > 0) then
  begin
    i_ind := fTobsPlanningTobRes.detail.count - 1;
    vTobRes := fTobsPlanningTobRes.detail[i_ind];
    vRessource := vTobRes.GetString('ATR_RESSOURCE');
    for i_ind := fTobsPlanningTobRes.detail.count - 1 downto 1 do
    begin
      if vRessource <> fTobsPlanningTobRes.detail[i_ind].GetString('ATR_RESSOURCE') then
      begin
        vTob := Tob.create('les_Ressources', fTobsPlanningTobRes, i_ind + 1);
        vTob.Dupliquer(vTobRes, false, true, false);
        vTob.PutValue('ATR_AFFAIRE', fStAffNonProductive);
        vTob.PutValue('ATR_TIERS', '');
      end;
      vTobRes := fTobsPlanningTobRes.detail[i_ind];
      vRessource := vTobRes.GetString('ATR_RESSOURCE');
    end;
    vTob := Tob.create('les_Ressources', fTobsPlanningTobRes, 0);
    vTob.Dupliquer(vTobRes, false, true, false);
    vTob.PutValue('ATR_AFFAIRE', fStAffNonProductive);
    vTob.PutValue('ATR_TIERS', '');
  end;

  // C.B 02/11/2004
  // si on gere les absences
  // on ne doit formater qu'une seule fois
  if not fBoGestionAbs then
    formatTobRes;}
    
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 02/11/2004
Modifié le ... :   /  /
Description .. : création d'une tache virtuelle pour les absences
Mots clefs ... :
*****************************************************************}
procedure TPlanningGA2.AjoutTacheAbsence(pInIndex: Integer; pStRes, pStResLib: string);
var
  vTob: Tob;
  i: integer;

begin

  vTob := Tob.create('les_Ressources', fTobsPlanningTobRes, pInIndex);
  if pInIndex = 0 then
    i := 1
  else
    i := 0;
  vTob.Dupliquer(fTobsPlanningTobRes.detail[i], false, false, false);
  vTob.PutValue('ATR_RESSOURCE', pStRes);
  vTob.PutValue('CHAMP1', pStResLib);
  vTob.PutValue('ATR_AFFAIRE', traduireMemoire(TexteMessage[75])); //'ABSENCES'
  vTob.PutValue('ATR_TIERS', '');
end;

// Journée non productive
procedure TPlanningGA2.AjoutNonProductive(pInIndex: Integer; pStRes, pStResLib: string);
var
  vTob: Tob;
  i: integer;

begin
  vTob := Tob.create('les_Ressources', fTobsPlanningTobRes, pInIndex);
  if pInIndex = 0 then
    i := 1
  else
    i := 0;
  vTob.Dupliquer(fTobsPlanningTobRes.detail[i], false, false, false);
  vTob.PutValue('ATR_AFFAIRE', fStAffNonProductive);
  vTob.PutValue('ATR_RESSOURCE', pStRes);
  vTob.PutValue('CHAMP1', pStResLib);
  vTob.PutValue('ATR_TIERS', '');
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 08/03/2002
Modifié le ... :
Description .. : formate l'affichage de la tob des ressources
Suite ........ :
Mots clefs ... :
*****************************************************************}
procedure TPlanningGA2.formatTobRes;
var
  i						: Integer;
  vStCle			: String;
  vStAffaire	: String;

begin

  for i := 0 to fTobsPlanningTobRes.detail.count - 1 do
  begin
    // creation de la cle
    vStCle := fTobsPlanningTobRes.detail[i].GetValue('ATR_RESSOURCE');

    if fTobsPlanningTobRes.detail[i].GetValue('ATR_AFFAIRE') <> NULL then
      vStCle := vStCle + cStSep + fTobsPlanningTobRes.detail[i].GetValue('ATR_AFFAIRE')
    else
      vStCle := vStCle + cStSep;

    fTobsPlanningTobRes.detail[i].PutValue('CLE', vStCle);

    // formatage du code a affaire
    // C.B 18/05/2006
    if GetParamsocSecur('SO_AFAFFICHAFFAIRE', False) then
      vStAffaire := fTobsPlanningTobRes.detail[i].GetValue('AFF_LIBELLE')
    else
      vStAffaire := CodeAffaireAffiche(fTobsPlanningTobRes.detail[i].GetValue('ATR_AFFAIRE'));

    if (fStAffNonProductive <> '') and (fTobsPlanningTobRes.detail[i].GetValue('ATR_AFFAIRE') = fStAffNonProductive) then
      fTobsPlanningTobRes.detail[i].PutValue('CHAMP2', fStLibNonProductive) // Journée non productive
    else
      if fBoGestionAbs and (fTobsPlanningTobRes.detail[i].GetValue('ATR_AFFAIRE') = traduireMemoire(TexteMessage[75])) then //'ABSENCES'
      fTobsPlanningTobRes.detail[i].PutValue('CHAMP2', fTobsPlanningTobRes.detail[i].GetValue('ATR_AFFAIRE'))
    else
      fTobsPlanningTobRes.detail[i].PutValue('CHAMP2', fTobsPlanningTobRes.detail[i].GetValue('ATR_TIERS') + #13#10 +
        RechDom('GCTIERSCLI', fTobsPlanningTobRes.detail[i].GetValue('ATR_TIERS'), False) + #13#10 +
        vStAffaire);

    // supression des null des affaires
    if (fTobsPlanningTobRes.detail[i].GetValue('CHAMP2') = NULL) then
      fTobsPlanningTobRes.detail[i].putvalue('CHAMP2', '');

		// attendre V7
		if not (pos(#13#10, fTobsPlanningTobRes.detail[i].GetValue('CHAMP1')) > 0) then
  		fTobsPlanningTobRes.detail[i].PutValue('CHAMP1', fTobsPlanningTobRes.detail[i].GetValue('CHAMP1') + #13#10 + fTobsPlanningTobRes.detail[i].GetValue('ATR_RESSOURCE'));

  end;

  // construction de l'arborescence
  // on n'affiche pas 2 fois la meme ressource
  for i := fTobsPlanningTobRes.detail.count - 1 downto 1 do
  begin
    if fTobsPlanningTobRes.detail[i].GetValue('CHAMP1') = fTobsPlanningTobRes.detail[i - 1].GetValue('CHAMP1') then
      fTobsPlanningTobRes.detail[i].PutValue('CHAMP1', '');
  end;

end;

// affaire, fonction, ressource
procedure TPlanningGA2.DecodeCle(pStCle: string);
var
  vSt: string;
begin
  vSt := pStCle;
  fStRes := (Trim(ReadTokenPipe(vSt, cStSep)));
  fStAffaire := (Trim(ReadTokenPipe(vSt, cStSep)));
  fStTiers := GetTiers(fStAffaire);
  fStTache := '';
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 28/03/2002
Modifié le ... :
Description .. : création d'une ligne de planning et affichage à
Suite ........ : l'écran
Mots clefs ... :
*****************************************************************}
procedure TPlanningGA2.CreateItem(Sender: TObject; Item: TOB; var Cancel: boolean);
begin

  DecodeCle(item.getvalue('I_CODE'));
  if fStAffaire = traduireMemoire(TexteMessage[75]) then //'ABSENCES'
  begin
    PGIBoxAF(TexteMessage[12], TexteMessage[59]); //On ne peux rien faire sur une Absence.
    cancel := True;
  end
  else
  begin
    cancel := False;

    // création directe = sélection d'une ligne effectuée
    if fInModeCreation = cInCreationDirecte then
    begin
      // sélection possible d'affaire puis de ressource
      if (fStAffaireSel <> '') and (fStAffaire <> fStAffaireSel) then
        FFichePlanning.BT_EFFACELIGNESClick(self);
    end

    else if (fInModeCreation = cInCreationNonProductive) then // Journée non productive
    begin
      if (fStAffaireSel <> '') and (fStAffaire <> fStAffaireSel) then
      begin
        PGIBoxAF(TexteMessage[22], TexteMessage[59]);
        cancel := True;
      end;
    end

    // creation sans selection de ressource
    else if (fInModeCreation = cInCreationAvecSelection) then
    begin
      if (fStAffaireSel <> '') and (fStAffaire <> fStAffaireSel) then
      begin
        Item.PutValue('APL_AFFAIRE', fStAffaire);
        Item.PutValue('APL_TIERS', fStTiers);
        Item.PutValue('APL_RESSOURCE', fStRes);
      end;
    end;

    // C.B 25/03/2005
    // sur une ligne de tache, on force cette tache
    //vStRetour := AFLanceFicheAFPlanning ('', 'CODEARTICLE=' + fStCodeArtSel + ';' + 'TACHE=' + fStTacheSel + ';' + vStAction)
    fStTacheSel := fStTache;

    if not cancel then
    begin
      fTobCouranteMere := MaTobInterne('AFFAIRECOURANTE');
      fTobCourante := TOB.Create('AFFAIRECOURANTE', fTobCouranteMere, -1);
      fTobCourante.Dupliquer(Item, TRUE, TRUE, TRUE);

      inherited;
    end;
    FormatChampLibelle(Item);
  end;
end;

procedure TPlanningGA2.CopyItem(Sender: TObject; Item: TOB; var Cancel: boolean);
var
  vStRetour   : string;
  vStAffaire  : string;
  vSt         : string;

begin

  // C.B 26/09/2006
  // suppression du spécifique cegid
  {if IsStatutModifiableTob(Item.GetValue('APL_ETATLIGNE'), fTobsPlanningTobEtats) then
    Item.putValue('APL_ETATLIGNE', GetStatutItem(Item.GetString('ATA_STATUTPLA')));

  // C.B 30/10/2006
  // correction du déplacement des ressources
  DecodeCle(Item.GetValue('I_CODE'));
  Item.PutValue('APL_RESSOURCE', fStRes);

  //mcd 14/03/2005
  if not JaiLeDroitRessource(Item.getValue('APL_RESSOURCE')) then
  begin
    PgiBoxAF(TexteMessage[24], TexteMessage[59]);
    cancel := true;
  end

  // test si on duplique sur le même emplacement
  else if not AutoriserPlanif(Item, fBoDemiJournee, False, vStAffaire, 'DUPLICATION') then
    cancel := true

  // C.B 31/10/2006
  // test si on est a cheval sur 2 semaines ou 2 mois
  else if GetParamSocSecur('SO_AFPARSEMAINE', True) and
     EstAChevalSemaine(Item, fBoDemiJournee) then
  begin
    PGIBoxAF(TexteMessage[70], TexteMessage[59]); // Vous ne pouvez pas créer, déplacer, réduire ou étirer une intervention à cheval sur 2 semaines.
    cancel := True;
  end

  // C.B 06/04/2006
  // test si on est a cheval sur 2 semaines ou 2 mois
  else if GetParamSocSecur('SO_AFPARMOIS', True) and
     EstAChevalMois(Item, fBoDemiJournee) then
  begin
    PGIBoxAF(TexteMessage[71], TexteMessage[59]); // Vous ne pouvez pas créer, déplacer, réduire ou étirer une intervention à cheval sur 2 semaines.
    cancel := True;
  end

  // Recherche de l'existant dans la table tacheessource
  else if ChangeLigneItemOk(Item) and CreateTacheRessource(Item, True) then
  begin
    DecodeCle(Item.GetValue('I_CODE'));
    Item.PutValue('APL_AFFAIRE', fStAffaire);
    Item.PutValue('APL_TIERS', fStTiers);
    Item.PutValue('APL_RESSOURCE', fStRes);

    fTobCouranteMere := MaTobInterne('AFFAIRECOURANTE');
    fTobCourante := TOB.Create('AFFAIRECOURANTE', fTobCouranteMere, -1);
    fTobCourante.Dupliquer(Item, TRUE, TRUE, TRUE);

    if fBoDemiJournee then
      vSt := vSt + ';DEMIJOURNEE'
    else
      vSt := '';
    vStRetour := AFLanceFicheAFPlanning('', 'ACTION=CREATION;DUPLIQUER' + vSt);

    if (vStRetour = 'CANCEL') or (vStRetour = '') then
      cancel := true
    else
    begin

      // C.B 31/10/2006
      // si on n'affiche pas les samedis et les dimanches et qu'on
      // est à cheval sur 2 semaines, on averti que les samedis et les
      // dimanches sont comptés
      if (not GetParamSocSecur('SO_AFPLANNINGWE', true)) and
         EstAChevalSemaine(Item, fBoDemiJournee) then
      begin
        PGIBoxAF(TexteMessage[20], TexteMessage[59]); // Attention, Les samedis et dimanches ne sont pas visibles, mais ils sont planifiés.
      end;
          
      Item.dupliquer(fTobCourante, True, True, True);
      RafraichirItem(Item, vStRetour, 'ARTICLE', False, False);
      cancel := false;
    end;

    DetruitMaTobInterne('AFFAIRECOURANTE');
  end
  else
    cancel := true;}
    
end;

procedure TPlanningGA2.ModifyItem(Sender: TObject; Item: TOB; var Cancel: boolean);
var
  vStItem: string;
  vStCle: string;

begin

  if item <> nil then
  begin
    if Item.GetValue('APL_ETATLIGNE') = 'ABS' then
    begin
      cancel := True;
      PGIBoxAF(TexteMessage[12], TexteMessage[59]); //On ne peux rien faire sur une Absence.
    end
    else
    begin
      inherited;
      vStItem := AFLanceFicheAFPlanning(varAsType(Item.GetValue('APL_AFFAIRE'), varString) + ';' + varAsType(Item.GetValue('APL_NUMEROLIGNE'), varString), fStAction);
      if (vStItem = 'CANCEL') or (vStItem = '') then
        cancel := true
      else
      begin
        // si changement de cle
        EncodeCle(Item.GetValue('APL_RESSOURCE'), Item.GetValue('APL_AFFAIRE'), '', vStCle);
        Item.putvalue('I_CODE', vStCle);
        RafraichirItem(Item, vStItem, 'ARTICLE', False, False);
        cancel := false;
      end;
      DetruitMaTobInterne('AFFAIRECOURANTE');
    end;
  end;
end;

procedure TPlanningGA2.BParamClick;
begin
  //AFLanceFiche_PlanningParam('P2J');
end;

function TPlanningGA2.AfficherTiers : Boolean;
begin
  result := False;
end;

function TPlanningGA2.AfficherRess  : Boolean;
begin
  result := False;
end;

function TPlanningGA2.ModeAffichage : String;
begin
 result := 'ARTICLE';
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 12/04/2002
Modifié le ... :
Description .. : Mise a jour des items déplacés dans le planning
Suite ........ :
Mots clefs ... :
*****************************************************************}
function TPlanningGA2.UpdateItem(FromItem, ToItem: TOB; vStAction: string): Boolean;
var
  vBoOk: Boolean;
  vStRetour: string;

begin

  {vBoOk := True;
  if (vStAction <> 'POPUP') and
     (vStAction <> 'MOVE') then
  begin
    if PlusDeJours(FromItem, ToItem, fBoDemiJournee) then
	  begin
  	  vBoOk := (PGIAskAF(TexteMessage[13], '') = mrYes);
    	if vBoOk then vBoOk := (PGIAskAF(TexteMessage[14], '') = mrYes);
			if vBoOk and VH_GC.GCIfDefCEGID then
  	  	RevalorisationCegid(FromItem, ToItem, vStAction);
	  end
    // dans le cas de reduce, on execute systématiquement la revalorisation
    // car on est pas dans le cas plus de jours
    else if (vStAction = 'REDUCE') and VH_GC.GCIfDefCEGID then
    	RevalorisationCegid(FromItem, ToItem, vStAction);
	end;

  if not vBoOk then
    result := false
  else
  begin
    if ChangeLigneItemOk(ToItem) and CreateTacheRessource(ToItem, true) then
    begin
      DecodeCle(ToItem.GetValue('I_CODE'));
      ToItem.PutValue('APL_RESSOURCE', fStRes);
      ToItem.PutValue('APL_UTILISATEUR', V_PGI.User);

      UpdateData(FromItem, ToItem);

      // C.B 30/11/2006 fonction commune
      result := UpdateAfPlanning(ToItem, vStRetour);
    end
    else
      result := false;
  end;}
end;

function TPlanningGA2.ChangeLigneItemOk(Item: TOB): boolean;
begin
  result := true;
  DecodeCle(Item.GetValue('I_CODE'));
  if (Item.GetValue('APL_AFFAIRE') = '') then
  begin
    PGIBoxAF(TexteMessage[6], TexteMessage[59]);
    result := false;
  end
  else
    if (fStAffaire <> Item.GetValue('APL_AFFAIRE')) then
  begin
    PGIBoxAF(TexteMessage[5], TexteMessage[59]);
    result := false;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 01/08/2002
Modifié le ... :   /  /
Description .. : Test si la Ressource existe dans la table tacheRessource
Mots clefs ... : Si non, on propose la creation
*****************************************************************}
function TPlanningGA2.CreateTacheRessource(Item: Tob; pBoForce: Boolean): boolean;
var
  vTob: Tob;
  vStAff0: string;
  vStAff1: string;
  vStAff2: string;
  vStAff3: string;
  vStAvenant: string;

begin

  result := false;
  DecodeCle(Item.GetValue('I_CODE'));

  if not ExisteTacheRessource(iTem.GetString('APL_AFFAIRE'), item.getString('APL_NUMEROTACHE'), item.getString('APL_RESSOURCE')) then
  begin
    if pBoForce or (PGIAskAF(TraduitGA(TexteMessage[7]), '') = mrYes) then
    begin

      vTob := Tob.Create('TACHERESSOURCE', nil, -1);
      try
        vTob.PutValue('ATR_NUMEROTACHE', intToStr(item.getValue('APL_NUMEROTACHE')));
        vTob.PutValue('ATR_TIERS', item.GetValue('APL_TIERS'));
        vTob.PutValue('ATR_AFFAIRE', iTem.GetValue('APL_AFFAIRE'));
        CodeAffaireDecoupe(iTem.GetValue('APL_AFFAIRE'), vStAff0, vStAff1, vStAff2, vStAff3, vStAvenant, taModif, false);
        vTob.PutValue('ATR_AFFAIRE0', vStAff0);
        vTob.PutValue('ATR_AFFAIRE1', vStAff1);
        vTob.PutValue('ATR_AFFAIRE2', vStAff2);
        vTob.PutValue('ATR_AFFAIRE3', vStAff3);
        vTob.PutValue('ATR_AVENANT', vStAvenant);
        vTob.PutValue('ATR_RESSOURCE', fStRes);
        vTob.PutValue('ATR_FONCTION', item.GetValue('APL_FONCTION'));
        vTob.PutValue('ATR_STATUTRES', 'ACT');
        vTob.PutValue('ATR_DEVISE', V_PGI.DevisePivot);

        if not vTob.InsertDB(nil) then
          PGIBoxAF(TexteMessage[1], TexteMessage[59])
        else
          result := true;
      finally
        vTob.Free;
      end;
    end;
  end
  else
    result := true;
end;

procedure TPlanningGA2.InitPlanning;
begin
  inherited;
  // les changements d'états sont disponibles en fonction des autorisations
  LoadPopUpMenu(False);
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : attention : cet appel n'est possible, que si des
                 éléments de la table afplanning dans le where
Mots clefs ... :
*****************************************************************}
procedure TPlanningGA2.LoadAbsencesPlanning(pStWhere: string);
var
  vSt 					: string;
  i 						: Integer;
  vInNumTache 	: Integer;
  vInNumLigne 	: Integer;
  vTobItem 			: Tob;
  vTobAbsence 	: Tob;
  vTobAbsences 	: Tob;
  vTob 					: Tob;
  vTobRes 			: Tob;

begin

  vTobAbsences := TOB.Create('Mes Absences', nil, -1);
  vTobRes := TOB.Create('#RESSOURCE', nil, -1);

  try
    vInNumTache := -1;
    vInNumLigne := 1000;

    vSt := MakeRequetePaie(pStWhere);
    vTobAbsences.LoadDetailFromSql(vSt);

    // Selection des ressources
    //C.B 02/06/2006
    //ajout du libelle2
    vSt := 'SELECT ARS_RESSOURCE, ARS_LIBELLE || " " || ARS_LIBELLE2 AS ARS_LIBELLE FROM RESSOURCE WHERE ARS_RESSOURCE <> ""';
    vSt := vSt + fStWhere;
    vSt := vSt + ' ORDER BY ARS_LIBELLE';

    vTobRes.LoadDetailFromSql(vSt);

    // parcours des ressources et création des jours d'absences
    for i := 0 to vTobRes.detail.count - 1 do
    begin
      vTob := fTobsPlanningTobRes.FindFirst(['ATR_RESSOURCE'], [vTobRes.detail[i].GetValue('ARS_RESSOURCE')], true);
      if vTob <> nil then
      begin
        AjoutTacheAbsence(vTob.GetIndex,
                          vTobRes.detail[i].GetValue('ARS_RESSOURCE'),
                          vTobRes.detail[i].GetValue('ARS_LIBELLE'));

        AjoutNonProductive(vTob.GetIndex,
                           vTobRes.detail[i].GetValue('ARS_RESSOURCE'),
                           vTobRes.detail[i].GetValue('ARS_LIBELLE'));
                                                              
        vTobAbsence := vTobAbsences.FindFirst(['ARS_RESSOURCE'], [vTobRes.detail[i].GetValue('ARS_RESSOURCE')], true);
        while vTobAbsence <> nil do
        begin
          // ajout d'un item absence
          vTobItem := Tob.create('AFPLANNING', fTobsPlanningTobItems, -1);
          CreateAbsence(vTobItem, vTobAbsence,
                        vTobRes.detail[i].GetValue('ARS_RESSOURCE'),
                        vTobRes.detail[i].GetValue('ARS_RESSOURCE') + cStSep + traduireMemoire(TexteMessage[75]), //'ABSENCES'
                        i, vInNumTache, vInNumligne);                                                                          
          vTobAbsence := vTobAbsences.FindNext(['ARS_RESSOURCE'], [vTobRes.detail[i].GetValue('ARS_RESSOURCE')], true);
	        DisplayHint(vTobItem, TexteMessage[66]); //'ABSENCE'
        end;
      end;
    end;
  finally
    formatTobRes;
    FreeAndNil(vTobRes);
    FreeAndNil(vTobAbsences);
  end;
end;

{------------------------------------------------------------------------------}
{ OBJET TPlanningGA3 : Ressource / Affaire, Tache                              }
{------------------------------------------------------------------------------}
{ CLE : AFFAIRE / TACHE                                                        }
{------------------------------------------------------------------------------}
procedure TPlanningGA3.SelectPlanning;
begin

  {fColonne1.StField := 'CHAMP1';
  fColonne2.StField := 'CHAMP2';
  fColonne3.StField := '';

  fColonne1.StSize := '50';
  fColonne2.StSize := '50';
  fColonne3.StSize := '0';

  fColonne1.StAlign := 'L';
  fColonne2.StAlign := 'L';
  fColonne3.StAlign := '';

  fColonne1.StEntete := TraduitGA(TexteMessage[67]); //Tiers/Affaire

  if ModeAutomatique then
    fColonne2.StEntete := TraduitGA(TexteMessage[61]) //Article
  else
    fColonne2.StEntete := TraduitGA(TexteMessage[60]); //Tâche

  fColonne3.StEntete := '';}

end;

// chargement de toutes les lignes de planning et de toutes les ressources
procedure TPlanningGA3.LoadItemsPlanning(pStWhere: string);
var
  i: Integer;
  vStCle: string;
begin
  inherited;
  {for i := 0 to fTobsPlanningTobItems.detail.count - 1 do
  begin
    DisplayHint(fTobsPlanningTobItems.detail[i]);
    vStCle := fTobsPlanningTobItems.detail[i].GetValue('APL_AFFAIRE') + cStSep +
      intToStr(fTobsPlanningTobItems.detail[i].GetValue('APL_NUMEROTACHE'));
    fTobsPlanningTobItems.detail[i].putValue('I_CODE', vStCle);

    fTobsPlanningTobItems.detail[i].putValue('CODEPOSTAL',
      GetCodePostal(fTobsPlanningTobItems.detail[i],
      fTobAffAdm,
      fTobsPlanningTobItems.detail[i].GetValue('APL_AFFAIRE'),
      fTobsPlanningTobItems.detail[i].GetValue('APL_TIERS'),
      fTobsPlanningTobItems.detail[i].GetString('ATA_TYPEADRESSE'),
      fTobsPlanningTobItems.detail[i].GetString('ATA_NUMEROADRESSE'),
      fTobsPlanningTobItems.detail[i].GetString('APL_TYPEADRESSE'),
      fTobsPlanningTobItems.detail[i].GetString('APL_NUMEROADRESSE'), True, True));


    FormatChampLibelle(fTobsPlanningTobItems.detail[i]);
    IconeCEGID(i);
  end;}
end;

function TPlanningGA3.LoadResPlanning(pStWhere: string): Boolean;
var
  Q: TQuery;
  S: string;
  vStWhere: string;

begin

  {S := ' SELECT DISTINCT AFF_LIBELLE, ATA_AFFAIRE , ATA_TIERS, ';
  S := S + ' ATA_CODEARTICLE, ATA_LIBELLETACHE1 AS CHAMP2, ';
  S := S + ' ATA_NUMEROTACHE, ATA_IDENTLIGNE, ATA_STATUTPLA ';
  S := S + ' FROM AFFAIRE, PIECE, TACHE ';
  S := S + ' WHERE AFF_AFFAIRE = ATA_AFFAIRE';
  S := S + ' AND ATA_TYPEARTICLE = "PRE"';
  S := S + ' AND AFF_AFFAIRE = GP_AFFAIRE ';
  S := S + ' AND GP_NATUREPIECEG = AFF_NATUREPIECEG';

  if pos('ATR_NUMEROTACHE', pStWhere) <> 0 then
    ReplaceSubStr(pStWhere, 'ATR_NUMEROTACHE', 'ATA_NUMEROTACHE');

  vStWhere := pStWhere;
  if vStWhere <> '' then
    S := S + vStWhere;

  S := S + ' GROUP BY AFF_LIBELLE, ATA_AFFAIRE, ATA_LIBELLETACHE1, ATA_IDENTLIGNE, ATA_NUMEROTACHE, ATA_TIERS, ATA_CODEARTICLE, ATA_STATUTPLA ';
  // tri par lignes de commandes
  S := S + ' ORDER BY ATA_AFFAIRE, ATA_IDENTLIGNE';

  Q := nil;
  try
    Q := OpenSql(S, True);
    if not Q.Eof then
    begin
      fTobsPlanningTobRes.ClearDetail;
      fTobsPlanningTobRes.LoadDetailDB('les_Ressources', '', '', Q, False, True);
      fTobsPlanningTobRes.detail[0].addchampSup('CLE', true);
      fTobsPlanningTobRes.detail[0].addchampSup('CHAMP1', true);
      formatTobRes;
      result := true;
    end
    else
    begin
      result := false;
    end;

  finally
    if Q <> nil then
      Ferme(Q);
  end;}
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 08/03/2002
Modifié le ... :
Description .. : formate l'affichage de la tob des ressources
Suite ........ :
Mots clefs ... :
*****************************************************************}
procedure TPlanningGA3.AjoutTacheAbsence(pInIndex: Integer; pStRes, pStResLib: string);
begin
  // pas d'abscences dans ce planning
end;

procedure TPlanningGA3.AjoutNonProductive(pInIndex: Integer; pStRes, pStResLib: string);
begin
  // pas de tâches non productives dans ce planning
end;

procedure TPlanningGA3.LoadAbsencesPlanning(pStWhere: string);
begin
  // pas d'abscences dans ce planning
end;

procedure TPlanningGA3.formatTobRes;
var
  i						: Integer;
  vStCle			: String;
  vStAffaire 	: String;

begin

  for i := 0 to fTobsPlanningTobRes.detail.count - 1 do
  begin
    // creation de la cle
    vStCle := fTobsPlanningTobRes.detail[i].GetValue('ATA_AFFAIRE');
    vStCle := vStCle + cStSep + intToStr(fTobsPlanningTobRes.detail[i].GetValue('ATA_NUMEROTACHE'));
    fTobsPlanningTobRes.detail[i].PutValue('CLE', vStCle);

    // C.B 18/05/2006
    if GetParamsocSecur('SO_AFAFFICHAFFAIRE', False) then
      vStAffaire := fTobsPlanningTobRes.detail[i].GetValue('AFF_LIBELLE')
    else
      vStAffaire := CodeAffaireAffiche(fTobsPlanningTobRes.detail[i].GetValue('ATA_AFFAIRE'));

    fTobsPlanningTobRes.detail[i].PutValue('CHAMP1', fTobsPlanningTobRes.detail[i].GetValue('ATA_TIERS') + #13#10 +
      RechDom('GCTIERSCLI', fTobsPlanningTobRes.detail[i].GetValue('ATA_TIERS'), False) + #13#10 +
      vStAffaire);

    fTobsPlanningTobRes.detail[i].PutValue('CHAMP2', fTobsPlanningTobRes.detail[i].GetValue('ATA_CODEARTICLE') + #13#10 +
      fTobsPlanningTobRes.detail[i].GetValue('CHAMP2'));

  end;

  // on n'affiche pas 2 fois la meme affaire
  for i := fTobsPlanningTobRes.detail.count - 1 downto 1 do
  begin
    if fTobsPlanningTobRes.detail[i].GetValue('CHAMP1') = fTobsPlanningTobRes.detail[i - 1].GetValue('CHAMP1') then
      fTobsPlanningTobRes.detail[i].PutValue('CHAMP1', '');
  end;
end;

// affaire, fonction, ressource
procedure TPlanningGA3.DecodeCle(pStCle: string);
var
  vSt: string;
begin
  vSt := pStCle;
  fStAffaire := (Trim(ReadTokenPipe(vSt, cStSep)));
  fStTache := (Trim(ReadTokenPipe(vSt, cStSep)));
  fStTiers := GetTiers(fStAffaire);
  fStRes := '';
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 28/03/2002
Modifié le ... :
Description .. : création d'une ligne de planning et affichage à
Suite ........ : l'écran
Mots clefs ... :
*****************************************************************}
procedure TPlanningGA3.CreateItem(Sender: TObject; Item: TOB; var Cancel: boolean);
var
  vTob: Tob;

begin

  cancel := False;
  DecodeCle(item.getvalue('I_CODE'));

  // création directe = sélection d'une ligne effectuée
  if fInModeCreation = cInCreationDirecte then
  begin
    // sélection possible d'affaire puis de ressource
    if (fStAffaireSel <> '') and (fStAffaire <> fStAffaireSel) then
      FFichePlanning.BT_EFFACELIGNESClick(sender)
  end

    // creation sans selection de ressource
  else
    if (fInModeCreation = cInCreationAvecSelection) then
  begin
    if (fStAffaire <> '') and
      (fStAffaireSel <> '') and
      (fStAffaire <> fStAffaireSel) then
    begin
      Item.PutValue('APL_AFFAIRE', fStAffaire);
      Item.PutValue('APL_TIERS', fStTiers);
      Item.PutValue('APL_NUMEROTACHE', fStTache);
    end;
  end

  // planning 3, si la tache est connu, on se positionne sur l'enregitrement
  else
  begin
    vTob := fTobsPlanningTobRes.FindFirst(['ATA_AFFAIRE', 'ATA_NUMEROTACHE'], [fStAffaire, fStTache], true);
    if assigned(vTob) then
    begin
      fStAffaire := vTob.GetString('ATA_AFFAIRE');
      fStTiers := vTob.GetString('ATA_TIERS');
      fStTache := vTob.GetString('ATA_NUMEROTACHE');
      fStCodeArt := vTob.GetString('ATA_CODEARTICLE');
    end;
  end;

  if not cancel then
  begin
    fTobCouranteMere := MaTobInterne('AFFAIRECOURANTE');
    fTobCourante := TOB.Create('AFFAIRECOURANTE', fTobCouranteMere, -1);
    fTobCourante.Dupliquer(Item, TRUE, TRUE, TRUE);
    inherited;
    FormatChampLibelle(Item);
  end;
end;

procedure TPlanningGA3.CopyItem(Sender: TObject; Item: TOB; var Cancel: boolean);
var
  vStRetour   : string;
  vStAffaire  : string;
  vSt         : string;

begin

  // C.B 26/09/2006
  // suppression du spécifique cegid
  {if IsStatutModifiableTob(Item.GetValue('APL_ETATLIGNE'), fTobsPlanningTobEtats) then
    Item.putValue('APL_ETATLIGNE', GetStatutItem(Item.GetString('ATA_STATUTPLA')));

  //mcd 14/03/2005
  if not JaiLeDroitRessource(Item.getValue('APL_RESSOURCE')) then
  begin
    PgiBoxAf(TexteMessage[24], TexteMessage[59]); //La ressource ne fait pas partie de votre équipe
    cancel := true;
  end

  // test si on duplique sur le même emplacement
  else if not AutoriserPlanif(Item, fBoDemiJournee, False, vStAffaire, 'DUPLICATION') then
    cancel := true

  // C.B 31/10/2006
  // test si on est a cheval sur 2 semaines ou 2 mois
  else if GetParamSocSecur('SO_AFPARSEMAINE', True) and
     EstAChevalSemaine(Item, fBoDemiJournee) then
  begin
    PGIBoxAF(TexteMessage[70], TexteMessage[59]); // Vous ne pouvez pas créer, déplacer, réduire ou étirer une intervention à cheval sur 2 semaines.
    cancel := True;
  end

  // C.B 06/04/2006
  // test si on est a cheval sur 2 semaines ou 2 mois
  else if GetParamSocSecur('SO_AFPARMOIS', True) and
     EstAChevalMois(Item, fBoDemiJournee) then
  begin
    PGIBoxAF(TexteMessage[71], TexteMessage[59]); // Vous ne pouvez pas créer, déplacer, réduire ou étirer une intervention à cheval sur 2 semaines.
    cancel := True;
  end

  // Recherche de l'existant dans la table tacheessource
  else if ChangeLigneItemOk(Item) and CreateTacheRessource(Item, True) then
  begin

    DecodeCle(Item.GetValue('I_CODE'));
    Item.PutValue('APL_AFFAIRE', fStAffaire);
    Item.PutValue('APL_TIERS', fStTiers);
    Item.PutValue('APL_NUMEROTACHE', fStTache);

    fTobCouranteMere := MaTobInterne('AFFAIRECOURANTE');
    fTobCourante := TOB.Create('AFFAIRECOURANTE', fTobCouranteMere, -1);
    fTobCourante.Dupliquer(Item, TRUE, TRUE, TRUE);

    if fBoDemiJournee then
      vSt := ';DEMIJOURNEE'
    else
      vSt := '';
    vStRetour := AFLanceFicheAFPlanning('', 'ACTION=CREATION;DUPLIQUER' + vSt);
    if (vStRetour = 'CANCEL') or (vStRetour = '') then
      cancel := true
    else
    begin

      // C.B 31/10/2006
      // si on n'affiche pas les samedis et les dimanches et qu'on
      // est à cheval sur 2 semaines, on averti que les samedis et les
      // dimanches sont comptés
      if (not GetParamSocSecur('SO_AFPLANNINGWE', true)) and
         EstAChevalSemaine(Item, fBoDemiJournee) then
      begin
        PGIBoxAF(TexteMessage[20], TexteMessage[59]); // Attention, Les samedis et dimanches ne sont pas visibles, mais ils sont planifiés.
      end;

      Item.dupliquer(fTobCourante, True, True, True);
      RafraichirItem(Item, vStRetour, 'ARTICLE', False, True);
      cancel := false;
    end;
    DetruitMaTobInterne('AFFAIRECOURANTE');
  end
  else
    cancel := true;}
end;

procedure TPlanningGA3.ModifyItem(Sender: TObject; Item: TOB; var Cancel: boolean);
var
  vStItem: string;
  vStCle: string;

begin
  inherited;
  if item <> nil then
  begin
    vStItem := AFLanceFicheAFPlanning(varAsType(Item.GetValue('APL_AFFAIRE'), varString) + ';' + varAsType(Item.GetValue('APL_NUMEROLIGNE'), varString), fStAction);

    if (vStItem = 'CANCEL') or (vStItem = '') then
      cancel := true
    else
    begin
      // si changement de cle
      EncodeCle(Item.GetValue('APL_AFFAIRE'), Item.GetValue('APL_NUMEROTACHE'), '', vStCle);
      Item.putvalue('I_CODE', vStCle);
      RafraichirItem(Item, vStItem, 'RESSOURCE', False, True);
      cancel := false;
    end;
    DetruitMaTobInterne('AFFAIRECOURANTE');
  end;
end;

procedure TPlanningGA3.BParamClick;
begin
  //AFLanceFiche_PlanningParam('P3J');
end;

function TPlanningGA3.AfficherTiers : Boolean;
begin
  result := False;
end;

function TPlanningGA3.AfficherRess  : Boolean;
begin
  result := True;
end;

function TPlanningGA3.ModeAffichage : String;
begin
 result := 'RESSOURCE';
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 12/04/2002
Modifié le ... :
Description .. : Mise a jour des items déplacés dans le planning
Suite ........ :
Mots clefs ... :
*****************************************************************}
function TPlanningGA3.UpdateItem(FromItem, ToItem: TOB; vStAction: string): Boolean;
var
  vBoOk: Boolean;
  vStRetour: string;

begin

  {vBoOk := True;
  if (vStAction <> 'POPUP') and
     (vStAction <> 'MOVE') then
  begin
    if PlusDeJours(FromItem, ToItem, fBoDemiJournee) then
	  begin
  	  vBoOk := (PGIAskAF(TexteMessage[13], '') = mrYes);
    	if vBoOk then vBoOk := (PGIAskAF(TexteMessage[14], '') = mrYes);
			if vBoOk and VH_GC.GCIfDefCEGID then
  	  	RevalorisationCegid(FromItem, ToItem, vStAction);
	  end
    // dans le cas de reduce, on execute systématiquement la revalorisation
    // car on est pas dans le cas plus de jours
    else if (vStAction = 'REDUCE') and VH_GC.GCIfDefCEGID then
    	RevalorisationCegid(FromItem, ToItem, vStAction);
	end;

  if not vBoOk then
    result := false
  else
  begin
    if ChangeLigneItemOk(ToItem) and CreateTacheRessource(ToItem, True) then
    begin
      ToItem.PutValue('APL_UTILISATEUR', V_PGI.User);
      UpdateData(FromItem, ToItem);

      // C.B 30/11/2006 fonction commune
      result := UpdateAfPlanning(ToItem, vStRetour);
    end
    else
      result := false;
  end;}

end;

procedure TPlanningGA3.InitPlanning;
begin
  inherited;
  // les changements d'états sont disponibles en fonction des autorisations
  LoadPopUpMenu(False);
end;

function TPlanningGA3.ChangeLigneItemOk(Item: TOB): boolean;
begin
  result := true;
  DecodeCle(Item.GetValue('I_CODE'));
  if (Item.GetValue('APL_AFFAIRE') = '') then
  begin
    PGIBoxAF(TexteMessage[6], TexteMessage[59]); //Erreur au chargement de l''affaire.
    result := false;
  end
  else
    if (fStAffaire <> Item.GetValue('APL_AFFAIRE')) then
  begin
    PGIBoxAF(TexteMessage[5], TexteMessage[59]); //On ne peut pas changer l''affaire d''une intervention.
    result := false;
  end
  else
    if (fStTache <> Item.GetValue('APL_NUMEROTACHE')) then
  begin
    PGIBoxAF(TexteMessage[9], TexteMessage[59]); //On ne peut pas déplacer une ressource sur une autre tâche.
    result := false;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 01/08/2002
Modifié le ... :   /  /
Description .. : Test si la Ressource existe dans la table tacheRessource
Mots clefs ... : Si non, on propose la creation
*****************************************************************}
function TPlanningGA3.CreateTacheRessource(Item: Tob; pBoForce: Boolean): boolean;
var
  vTob: Tob;
  vStAff0: string;
  vStAff1: string;
  vStAff2: string;
  vStAff3: string;
  vStAvenant: string;

begin

  result := false;
  DecodeCle(Item.GetValue('I_CODE'));

  if not ExisteTacheRessource(iTem.GetString('APL_AFFAIRE'), fStTache, item.getValue('APL_RESSOURCE')) then
  begin
    if pBoForce or (PGIAskAF(TraduitGA(TexteMessage[7]), '') = mrYes) then //Ce code ressource n est pas prévu pour cette tâche, voulez l ajouter à la tâche ?
    begin

      vTob := Tob.Create('TACHERESSOURCE', nil, -1);
      try
        vTob.PutValue('ATR_NUMEROTACHE', fStTache);
        vTob.PutValue('ATR_TIERS', item.GetValue('APL_TIERS'));
        vTob.PutValue('ATR_AFFAIRE', iTem.GetValue('APL_AFFAIRE'));
        CodeAffaireDecoupe(iTem.GetValue('APL_AFFAIRE'), vStAff0, vStAff1, vStAff2, vStAff3, vStAvenant, taModif, false);
        vTob.PutValue('ATR_AFFAIRE0', vStAff0);
        vTob.PutValue('ATR_AFFAIRE1', vStAff1);
        vTob.PutValue('ATR_AFFAIRE2', vStAff2);
        vTob.PutValue('ATR_AFFAIRE3', vStAff3);
        vTob.PutValue('ATR_AVENANT', vStAvenant);
        vTob.PutValue('ATR_RESSOURCE', item.getValue('APL_RESSOURCE'));
        vTob.PutValue('ATR_FONCTION', item.GetValue('APL_FONCTION'));
        vTob.PutValue('ATR_STATUTRES', 'ACT');
        vTob.PutValue('ATR_DEVISE', V_PGI.DevisePivot);

        if not vTob.InsertDB(nil) then
          PGIBoxAF(TexteMessage[1], TexteMessage[59]) //Erreur lors de la mise à jour de l''intervention.
        else
          result := true;
      finally
        vTob.Free;
      end;
    end;
  end
  else
    result := true;
end;

{------------------------------------------------------------------------------}
{ OBJET TPlanningGA4 : Affaire / ressource, tache                              }
{------------------------------------------------------------------------------}
{ CLE : RESSOURCE                                                              }
{------------------------------------------------------------------------------}
{ C.B 07/07/2005 suppression du regroupement par tâches qui n'apportait rien   }
{------------------------------------------------------------------------------}
procedure TPlanningGA4.SelectPlanning;
begin

  fColonne1.StField := 'CHAMP1';
  fColonne2.StField := '';
  fColonne3.StField := '';

  fColonne1.StSize := '50';
  fColonne2.StSize := '0';
  fColonne3.StSize := '0';

  fColonne1.StAlign := 'L';
  fColonne2.StAlign := '';
  fColonne3.StAlign := '';

  fColonne1.StEntete := TraduitGA(TexteMessage[51]); //Ressource

  fColonne2.StEntete := '';
  fColonne3.StEntete := '';
end;

// chargement de toutes les lignes de planning et de toutes les ressources
procedure TPlanningGA4.LoadItemsPlanning(pStWhere: string);
var
  i: Integer;
  vStCle: string;

begin
  inherited;
  {for i := 0 to fTobsPlanningTobItems.detail.count - 1 do
  begin
    fTobsPlanningTobItems.detail[i].AddChampSupValeur('AFFICHE_AFFAIRE', fTobsPlanningTobItems.detail[i].GetValue('APL_TIERS') + #13#10 +
        RechDom('GCTIERSCLI', fTobsPlanningTobItems.detail[i].GetValue('APL_TIERS'), False) + #13#10 +
        fTobsPlanningTobItems.detail[i].GetValue('APL_AFFAIRE'), False);

    DisplayHint(fTobsPlanningTobItems.detail[i]);
    vStCle := fTobsPlanningTobItems.detail[i].GetValue('APL_RESSOURCE');

    fTobsPlanningTobItems.detail[i].putValue('I_CODE', vStCle);

    fTobsPlanningTobItems.detail[i].putValue('CODEPOSTAL',
      GetCodePostal(fTobsPlanningTobItems.detail[i],
      fTobAffAdm,
      fTobsPlanningTobItems.detail[i].GetValue('APL_AFFAIRE'),
      fTobsPlanningTobItems.detail[i].GetValue('APL_TIERS'),
      fTobsPlanningTobItems.detail[i].GetString('ATA_TYPEADRESSE'),
      fTobsPlanningTobItems.detail[i].GetString('ATA_NUMEROADRESSE'),
      fTobsPlanningTobItems.detail[i].GetValue('APL_TYPEADRESSE'),
      fTobsPlanningTobItems.detail[i].GetValue('APL_NUMEROADRESSE'), True, True));

    FormatChampLibelle(fTobsPlanningTobItems.detail[i]);
    IconeCEGID(i);
  end;}
end;

function TPlanningGA4.LoadResPlanning(pStWhere: string): Boolean;
var
  Q: TQuery;
  S: string;
  vStWhere: string;
begin

  {S := ' SELECT DISTINCT ATR_RESSOURCE, ';
  S := S + ' ARS_LIBELLE || " " || ARS_LIBELLE2 AS CHAMP1 ';
  S := S + ' FROM RESSOURCE, TACHERESSOURCE ';
  S := S + ' WHERE ATR_RESSOURCE = ARS_RESSOURCE ';

  // C.B 04/10/2005
  // on affiche le planning des clients fermés
  //S := S + ' AND ARS_FERME = "-"';

  vStWhere := pStWhere;
  if pos('APL_AFFAIRE', vStWhere) <> 0 then
    ReplaceSubStr(vStWhere, 'APL_AFFAIRE', 'ATR_AFFAIRE');

  if pos('APL_RESSOURCE', vStWhere) <> 0 then
    ReplaceSubStr(vStWhere, 'APL_RESSOURCE', 'ATR_RESSOURCE');

  if vStWhere <> '' then
    S := S + vStWhere;
  if fStAffNonProductive <> '' then // Journée non productive
    S := S + ' AND ATR_AFFAIRE <> "' + fStAffNonProductive + '"';

  S := S + ' GROUP BY ATR_RESSOURCE, ARS_LIBELLE || " " || ARS_LIBELLE2';
  S := S + ' ORDER BY ARS_LIBELLE || " " || ARS_LIBELLE2';

  Q := nil;
  try
    Q := OpenSql(S, True);
    if not Q.Eof then
    begin
      fTobsPlanningTobRes.ClearDetail;
      fTobsPlanningTobRes.LoadDetailDB('les_Ressources', '', '', Q, False, True);
      fTobsPlanningTobRes.detail[0].addchampSup('CLE', true);
      formatTobRes;
      result := true;
    end
    else
    begin
      result := false;
    end;

  finally
    if Q <> nil then
      Ferme(Q);
  end;}
end;

procedure TPlanningGA4.AjoutTacheAbsence(pInIndex: Integer; pStRes, pStResLib: string);
begin
  // pas d'abscences dans ce planning
end;

procedure TPlanningGA4.AjoutNonProductive(pInIndex: Integer; pStRes, pStResLib: string);
begin
  // pas de tâches non productives dans ce planning
end;

procedure TPlanningGA4.LoadAbsencesPlanning(pStWhere: string);
begin
  // pas d'abscences dans ce planning
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 08/03/2002
Modifié le ... :
Description .. : formate l'affichage de la tob des ressources
Suite ........ :
Mots clefs ... :
*****************************************************************}
procedure TPlanningGA4.formatTobRes;
var
  i: integer;
  vStCle: string;

begin

  for i := 0 to fTobsPlanningTobRes.detail.count - 1 do
  begin
    // creation de la cle
    vStCle := fTobsPlanningTobRes.detail[i].GetValue('ATR_RESSOURCE');
    //vStCle := vStCle + cStSep + intToStr(fTobsPlanningTobRes.detail[i].GetValue('ATR_NUMEROTACHE'));
    fTobsPlanningTobRes.detail[i].PutValue('CLE', vStCle);
		// attendre V7
		if not (pos(#13#10, fTobsPlanningTobRes.detail[i].GetValue('CHAMP1')) > 0) then
  		fTobsPlanningTobRes.detail[i].PutValue('CHAMP1', fTobsPlanningTobRes.detail[i].GetValue('CHAMP1') + #13#10 + fTobsPlanningTobRes.detail[i].GetValue('ATR_RESSOURCE'));
  end;
end;

// affaire, fonction, ressource
procedure TPlanningGA4.DecodeCle(pStCle: string);
var
  vSt: string;
begin
  vSt := pStCle;
  fStRes := (Trim(ReadTokenPipe(vSt, cStSep)));
  fStTache := '';
  fStAffaire := '';
  fStTiers := '';
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 28/03/2002
Modifié le ... :
Description .. : création d'une ligne de planning et affichage à
Suite ........ : l'écran
Mots clefs ... :
*****************************************************************}
procedure TPlanningGA4.CreateItem(Sender: TObject; Item: TOB; var Cancel: boolean);
begin
  cancel := False;
end;

procedure TPlanningGA4.CopyItem(Sender: TObject; Item: TOB; var Cancel: boolean);
begin
  cancel := True;
end;

procedure TPlanningGA4.ModifyItem(Sender: TObject; Item: TOB; var Cancel: boolean);
begin
  cancel := false;
end;

procedure TPlanningGA4.BParamClick;
begin
  //AFLanceFiche_PlanningParam('P4J');
end;

function TPlanningGA4.AfficherTiers : Boolean;
begin
  result := False;
end;

function TPlanningGA4.AfficherRess  : Boolean;
begin
  result := False;
end;

function TPlanningGA4.ModeAffichage : String;
begin
 result := 'AFFAIRE';
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 12/04/2002
Modifié le ... :
Description .. : Mise a jour des items déplacés dans le planning
Suite ........ :
Mots clefs ... :
*****************************************************************}
function TPlanningGA4.UpdateItem(FromItem, ToItem: TOB; vStAction: string): Boolean;
begin
  result := false;
end;

procedure TPlanningGA4.InitPlanning;
begin
  // forcé en consultation
  fAction := TaConsult;
  inherited;

  // les changements d'états sont disponibles en fonction des autorisations
  LoadPopUpMenu(False);
end;

function TPlanningGA4.ChangeLigneItemOk(Item: TOB): boolean;
begin
  result := true;
  DecodeCle(Item.GetValue('I_CODE'));
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 01/08/2002
Modifié le ... :   /  /
Description .. : Test si la Ressource existe dans la table tacheRessource
Mots clefs ... : Si non, on propose la creation
*****************************************************************}
function TPlanningGA4.CreateTacheRessource(Item: Tob; pBoForce: Boolean): boolean;
begin
  result := true;
end;

{------------------------------------------------------------------------------}
{ OBJET TPlanningGA5 : Taches / ressource, Affaire                             }
{------------------------------------------------------------------------------}
{ CLE : RESSOURCE / AFFAIRE                                                    }
{------------------------------------------------------------------------------}
procedure TPlanningGA5.SelectPlanning;
begin

  fColonne1.StField := 'CHAMP1';
  fColonne2.StField := '';
  fColonne3.StField := '';

  fColonne1.StSize := '50';
  fColonne2.StSize := '50';
  fColonne3.StSize := '0';

  fColonne1.StAlign := 'L';
  fColonne2.StAlign := '';
  fColonne3.StAlign := '';

  fColonne1.StEntete := TraduitGA(TexteMessage[51]); //Ressource
  fColonne2.StEntete := '';
  fColonne3.StEntete := '';
end;

// chargement de toutes les lignes de planning et de toutes les ressources
procedure TPlanningGA5.LoadItemsPlanning(pStWhere: string);
var
  i: Integer;
  vStCle: string;
begin
  inherited;
  {for i := 0 to fTobsPlanningTobItems.detail.count - 1 do
  begin
    DisplayHint(fTobsPlanningTobItems.detail[i]);
    vStCle := fTobsPlanningTobItems.detail[i].GetValue('APL_RESSOURCE');
    fTobsPlanningTobItems.detail[i].putValue('I_CODE', vStCle);

    fTobsPlanningTobItems.detail[i].putValue('CODEPOSTAL',
      GetCodePostal(fTobsPlanningTobItems.detail[i],
      fTobAffAdm,
      fTobsPlanningTobItems.detail[i].GetValue('APL_AFFAIRE'),
      fTobsPlanningTobItems.detail[i].GetValue('APL_TIERS'),
      fTobsPlanningTobItems.detail[i].GetString('ATA_TYPEADRESSE'),
      fTobsPlanningTobItems.detail[i].GetString('ATA_NUMEROADRESSE'),
      fTobsPlanningTobItems.detail[i].GetValue('APL_TYPEADRESSE'),
      fTobsPlanningTobItems.detail[i].GetValue('APL_NUMEROADRESSE'), True, True));

    FormatChampLibelle(fTobsPlanningTobItems.detail[i]);
    IconeCEGID(i);
  end;}
end;

function TPlanningGA5.LoadResPlanning(pStWhere: string): Boolean;
var
  Q: TQuery;
  S: string;
  vStWhere: string;
begin

  {S := ' SELECT DISTINCT ARS_RESSOURCE AS ATR_RESSOURCE, ';
  S := S + ' ARS_LIBELLE || " " || ARS_LIBELLE2 AS CHAMP1 ';
  S := S + ' FROM RESSOURCE ';
  S := S + ' WHERE ARS_RESSOURCE <> "" ';

  vStWhere := pStWhere;

  if pos('APL_RESSOURCE', vStWhere) <> 0 then
    ReplaceSubStr(vStWhere, 'APL_RESSOURCE', 'ARS_RESSOURCE');

  if vStWhere <> '' then
    S := S + vStWhere;

  S := S + ' GROUP BY ARS_RESSOURCE, ARS_LIBELLE || " " || ARS_LIBELLE2';
  S := S + ' ORDER BY ARS_LIBELLE || " " || ARS_LIBELLE2';

  Q := nil;
  try
    Q := OpenSql(S, True);
    if not Q.Eof then
    begin
      fTobsPlanningTobRes.ClearDetail;
      fTobsPlanningTobRes.LoadDetailDB('les_Ressources', '', '', Q, False, True);
      fTobsPlanningTobRes.detail[0].addchampSup('CLE', true);
      // C.B 02/11/2004
      // si on gere les absences
      // on ne doit formater qu'une seule fois
      if not fBoGestionAbs then
        formatTobRes;
      result := true;
    end
    else
    begin
      result := false;
    end;
  finally
    if Q <> nil then
      Ferme(Q);
  end;}

end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... :
Modifié le ... : 05/01/2006
Description .. : recherche des abscences
                 optimisation de la requete
Mots clefs ... :
*****************************************************************}
procedure TPlanningGA5.LoadAbsencesPlanning(pStWhere: string);
begin
  LoadAbsencesPlanningRessource(pStWhere);
end;

procedure TPlanningGA5.AjoutTacheAbsence(pInIndex: Integer; pStRes, pStResLib: string);
begin
  // pas de lignes a ajouter pour les abscences dans ce planning
end;

procedure TPlanningGA5.AjoutNonProductive(pInIndex: Integer; pStRes, pStResLib: string);
begin
  // pas de lignes à ajouter pour les tâches non productives dans ce planning
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 08/03/2002
Modifié le ... :
Description .. : formate l'affichage de la tob des ressources
Suite ........ :
Mots clefs ... :
*****************************************************************}
procedure TPlanningGA5.formatTobRes;
var
  i: integer;
  vStCle: string;
begin
  for i := 0 to fTobsPlanningTobRes.detail.count - 1 do
  begin
    // creation de la cle
    vStCle := fTobsPlanningTobRes.detail[i].GetValue('ATR_RESSOURCE');
    fTobsPlanningTobRes.detail[i].PutValue('CLE', vStCle);
    // attendre V7
		if not (pos(#13#10, fTobsPlanningTobRes.detail[i].GetValue('CHAMP1')) > 0) then
  		fTobsPlanningTobRes.detail[i].PutValue('CHAMP1', fTobsPlanningTobRes.detail[i].GetValue('CHAMP1') + #13#10 + fTobsPlanningTobRes.detail[i].GetValue('ATR_RESSOURCE'));
  end;
end;

// affaire, fonction, ressource
procedure TPlanningGA5.DecodeCle(pStCle: string);
var
  vSt: string;
begin
  vSt := pStCle;
  fStRes := (Trim(ReadTokenPipe(vSt, cStSep)));
  fStAffaire := '';
  fStTiers := '';
  fStTache := '';
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 28/03/2002
Modifié le ... :
Description .. : création d'une ligne de planning et affichage à
Suite ........ : l'écran
Mots clefs ... :
*****************************************************************}

procedure TPlanningGA5.CreateItem(Sender: TObject; Item: TOB; var Cancel: boolean);
begin
  cancel := False;
  DecodeCle(item.getvalue('I_CODE'));

  // création directe = sélection d'une ligne effectuée
  if fInModeCreation = cInCreationDirecte then
  begin
    // sélection possible ressource
    if (fStRes <> fStResSel) then
    begin
      if (PGIAskAF(TraduitGA(TexteMessage[19]), '') = mrYes) then //Voulez-vous créer une intervention pour une autre ressource que celle de la ligne à planifier sélectionnée ?
      begin
        fInModeCreation := cInCreationSansSelection;
        FFichePlanning.La_Selection.caption := '';
        FFichePlanning.BT_Lignes.down := False;
      end
      else
        cancel := True;
    end;
  end

    // creation sans selection de ressource car plusieurs ressource pour cette tache
    // mais comme sur une ressource, on force une création directe ...
  else
    if (fInModeCreation = cInCreationAvecSelection) then
  begin
    fStAffaire := fStAffaireSel;
    fStTiers := fStTiersSel;
    fStResSel := fStRes;
    fInModeCreation := cInCreationSelectionEtRes;
  end

  else
    if (fInModeCreation = cInCreationNonProductive) then
  begin
    fStAffaire := fStAffaireSel;
    fStTiers := fStTiersSel;
  end;

  if not cancel then
  begin
    fTobCouranteMere := MaTobInterne('AFFAIRECOURANTE');
    fTobCourante := TOB.Create('AFFAIRECOURANTE', fTobCouranteMere, -1);
    fTobCourante.Dupliquer(Item, TRUE, TRUE, TRUE);

    inherited;
    FormatChampLibelle(Item);
  end;
end;

procedure TPlanningGA5.CopyItem(Sender: TObject; Item: TOB; var Cancel: boolean);
var
  vStRetour   : string;
  vStAffaire  : string;
  vSt         : string;

begin

  // C.B 26/09/2006
  // suppression du spécifique cegid
  {if IsStatutModifiableTob(Item.GetValue('APL_ETATLIGNE'), fTobsPlanningTobEtats) then
    Item.putValue('APL_ETATLIGNE', GetStatutItem(Item.GetString('ATA_STATUTPLA')));

  // C.B 30/10/2006
  // correction du déplacement des ressources
  DecodeCle(Item.GetValue('I_CODE'));
  Item.PutValue('APL_RESSOURCE', fStRes);

  //mcd 14/03/2005
  if not JaiLeDroitRessource(Item.getValue('APL_RESSOURCE')) then
  begin
    PgiBoxAf(TexteMessage[24], TexteMessage[59]); //La ressource ne fait pas partie de votre équipe
    cancel := true;
  end
  else
  begin
    // Recherche de l'existant dans la table tacheressource
    fStAffaire := Item.GetValue('APL_AFFAIRE');

    // test si on duplique sur le même emplacement
    if not AutoriserPlanif(Item, fBoDemiJournee, False, vStAffaire, 'DUPLICATION') then
      cancel := true

    // C.B 31/10/2006
    // test si on est a cheval sur 2 semaines ou 2 mois
    else if GetParamSocSecur('SO_AFPARSEMAINE', True) and
       EstAChevalSemaine(Item, fBoDemiJournee) then
    begin
      PGIBoxAF(TexteMessage[70], TexteMessage[59]); // Vous ne pouvez pas créer, déplacer, réduire ou étirer une intervention à cheval sur 2 semaines.
      cancel := True;
    end

    // C.B 06/04/2006
    // test si on est a cheval sur 2 semaines ou 2 mois
    else if GetParamSocSecur('SO_AFPARMOIS', True) and
       EstAChevalMois(Item, fBoDemiJournee) then
    begin
      PGIBoxAF(TexteMessage[71], TexteMessage[59]); // Vous ne pouvez pas créer, déplacer, réduire ou étirer une intervention à cheval sur 2 semaines.
      cancel := True;
    end

    else if CreateTacheRessource(Item, True) then
    begin
      DecodeCle(Item.GetValue('I_CODE'));
      Item.PutValue('APL_RESSOURCE', fStRes);

      fTobCouranteMere := MaTobInterne('AFFAIRECOURANTE');
      fTobCourante := TOB.Create('AFFAIRECOURANTE', fTobCouranteMere, -1);
      fTobCourante.Dupliquer(Item, TRUE, TRUE, TRUE);

      if fBoDemiJournee then
        vSt := ';DEMIJOURNEE'
      else
        vSt := '';
      vStRetour := AFLanceFicheAFPlanning('', 'ACTION=CREATION;DUPLIQUER' + vSt);

      if (vStRetour = 'CANCEL') or (vStRetour = '') then
        cancel := true
      else
      begin

        // C.B 31/10/2006
        // si on n'affiche pas les samedis et les dimanches et qu'on
        // est à cheval sur 2 semaines, on averti que les samedis et les
        // dimanches sont comptés
        if (not GetParamSocSecur('SO_AFPLANNINGWE', true)) and
           EstAChevalSemaine(Item, fBoDemiJournee) then
        begin
          PGIBoxAF(TexteMessage[20], TexteMessage[59]); // Attention, Les samedis et dimanches ne sont pas visibles, mais ils sont planifiés.
        end;

        Item.dupliquer(fTobCourante, True, True, True);
        RafraichirItem(Item, vStRetour, 'ARTICLE', False, True);
        cancel := false;
      end;
      DetruitMaTobInterne('AFFAIRECOURANTE');
    end

      // test si on duplique sur le même emplacement
    else
      if not AutoriserPlanif(Item, fBoDemiJournee, False, vStAffaire, 'DUPLICATION') then
      cancel := true

    else
      cancel := true;
  end;}
  
end;

procedure TPlanningGA5.ModifyItem(Sender: TObject; Item: TOB; var Cancel: boolean);
var
  vStItem: string;
  vStCle: string;

begin
  if item <> nil then
  begin
    if Item.GetValue('APL_ETATLIGNE') = 'ABS' then
    begin
      cancel := True;
      PGIBoxAF(TexteMessage[12], TexteMessage[59]); //On ne peux rien faire sur une Absence.
    end
    else
    begin
      inherited;
      vStItem := AFLanceFicheAFPlanning(varAsType(Item.GetValue('APL_AFFAIRE'), varString) + ';' + varAsType(Item.GetValue('APL_NUMEROLIGNE'), varString), fStAction);

      if (vStItem = 'CANCEL') or (vStItem = '') then
        cancel := true
      else
      begin
        // si changement de cle
        EncodeCle(Item.GetValue('APL_RESSOURCE'), Item.GetValue('APL_AFFAIRE'), '', vStCle);
        Item.putvalue('I_CODE', vStCle);
        RafraichirItem(Item, vStItem, 'ARTICLE', False, True);
        cancel := false;
      end;
      DetruitMaTobInterne('AFFAIRECOURANTE');
    end;
  end;
end;

procedure TPlanningGA5.BParamClick;
begin
  //AFLanceFiche_PlanningParam('P2J');
end;

function TPlanningGA5.AfficherTiers : Boolean;
begin
  result := True;
end;

function TPlanningGA5.AfficherRess  : Boolean;
begin
  result := False;
end;

function TPlanningGA5.ModeAffichage : String;
begin
 result := 'ARTICLE';
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 12/04/2002
Modifié le ... :
Description .. : Mise a jour des items déplacés dans le planning
Suite ........ :
Mots clefs ... :
*****************************************************************}
function TPlanningGA5.UpdateItem(FromItem, ToItem: TOB; vStAction: string): Boolean;
var
  vBoOk: Boolean;
  vStRetour: string;

begin

  {vBoOk := True;
  if (vStAction <> 'POPUP') and
     (vStAction <> 'MOVE') then
  begin
    if PlusDeJours(FromItem, ToItem, fBoDemiJournee) then
	  begin
  	  vBoOk := (PGIAskAF(TexteMessage[13], '') = mrYes);
    	if vBoOk then vBoOk := (PGIAskAF(TexteMessage[14], '') = mrYes);
			if vBoOk and VH_GC.GCIfDefCEGID then
  	  	RevalorisationCegid(FromItem, ToItem, vStAction);
	  end
    // dans le cas de reduce, on execute systématiquement la revalorisation
    // car on est pas dans le cas plus de jours
    else if (vStAction = 'REDUCE') and VH_GC.GCIfDefCEGID then
    	RevalorisationCegid(FromItem, ToItem, vStAction);
	end;

  if not vBoOk then
    result := false
  else
  begin
    // en copy, seul ToItem est renseigné
    if not assigned(FromItem) and Assigned(ToItem) then
      fStAffaire := ToItem.GetValue('APL_AFFAIRE')
    else
      fStAffaire := FromItem.GetValue('APL_AFFAIRE');

    if CreateTacheRessource(ToItem, True) then
    begin
      DecodeCle(ToItem.GetValue('I_CODE'));
      ToItem.PutValue('APL_RESSOURCE', fStRes);
      ToItem.PutValue('APL_UTILISATEUR', V_PGI.User);

      UpdateData(FromItem, ToItem);

      // C.B 30/11/2006 fonction commune
      result := UpdateAfPlanning(ToItem, vStRetour);
    end
    else
      result := false;
  end;}
end;

// on supprime ce test car fStAffaire n'est pas forcement renseigné
// à ce moment la dans cet écran
function TPlanningGA5.ChangeLigneItemOk(Item: TOB): boolean;
begin
  result := true;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 01/08/2002
Modifié le ... :   /  /
Description .. : Test si la Ressource existe dans la table tacheRessource
Mots clefs ... : Si non, on propose la creation
*****************************************************************}
function TPlanningGA5.CreateTacheRessource(Item: Tob; pBoForce: Boolean): boolean;
var
  vTob: Tob;
  vStAff0: string;
  vStAff1: string;
  vStAff2: string;
  vStAff3: string;
  vStAvenant: string;

begin

  result := false;
  if (fStAffaire <> '') and (not ExisteTacheRessource(iTem.GetString('APL_AFFAIRE'), intToStr(item.getValue('APL_NUMEROTACHE')), fStRes)) then
  begin
    if pBoForce or (PGIAskAF(TraduitGA(TexteMessage[7]), '') = mrYes) then
    begin

      vTob := Tob.Create('TACHERESSOURCE', nil, -1);
      try
        vTob.PutValue('ATR_NUMEROTACHE', intToStr(item.getValue('APL_NUMEROTACHE')));
        vTob.PutValue('ATR_TIERS', item.GetValue('APL_TIERS'));
        vTob.PutValue('ATR_AFFAIRE', iTem.GetValue('APL_AFFAIRE'));
        CodeAffaireDecoupe(iTem.GetValue('APL_AFFAIRE'), vStAff0, vStAff1, vStAff2, vStAff3, vStAvenant, taModif, false);
        vTob.PutValue('ATR_AFFAIRE0', vStAff0);
        vTob.PutValue('ATR_AFFAIRE1', vStAff1);
        vTob.PutValue('ATR_AFFAIRE2', vStAff2);
        vTob.PutValue('ATR_AFFAIRE3', vStAff3);
        vTob.PutValue('ATR_AVENANT', vStAvenant);
        vTob.PutValue('ATR_RESSOURCE', fStRes);
        vTob.PutValue('ATR_FONCTION', item.GetValue('APL_FONCTION'));
        vTob.PutValue('ATR_STATUTRES', 'ACT');
        vTob.PutValue('ATR_DEVISE', V_PGI.DevisePivot);

        if not vTob.InsertDB(nil) then
          PGIBoxAF(TexteMessage[1], TexteMessage[59]) //Erreur lors de la mise à jour de l''intervention.
        else
          result := true;
      finally
        vTob.Free;
      end;
    end;
  end
  else
    result := true;
end;

procedure TPlanningGA5.InitPlanning;
begin
  inherited;
  // les changements d'états sont disponibles en fonction des autorisations
  LoadPopUpMenu(False);
end;

{------------------------------------------------------------------------------}
{ OBJET TPlanningGA6 : Taches / Affaire                                        }
{------------------------------------------------------------------------------}
{ CLE : AFFAIRE                                                                }
{------------------------------------------------------------------------------}
procedure TPlanningGA6.SelectPlanning;
begin

  fColonne1.StField := 'CHAMP1';
  fColonne1.StSize := '70';
  fColonne1.StAlign := 'L';

  fColonne2.StField := '';
  fColonne3.StField := '';

  fColonne3.StSize := '0';
  fColonne2.StSize := '0';

  fColonne2.StEntete := '';
  fColonne3.StEntete := '';

  fColonne1.StEntete := TraduitGA(TexteMessage[67]) //Tiers/Affaire
end;

// chargement de toutes les lignes de planning et de toutes les ressources
procedure TPlanningGA6.LoadItemsPlanning(pStWhere: string);
var
  i: Integer;
  vStCle: string;

begin
  inherited;

  {for i := 0 to fTobsPlanningTobItems.detail.count - 1 do
  begin
    DisplayHint(fTobsPlanningTobItems.detail[i]);
    vStCle := fTobsPlanningTobItems.detail[i].GetValue('APL_AFFAIRE');
    fTobsPlanningTobItems.detail[i].putValue('I_CODE', vStCle);

    fTobsPlanningTobItems.detail[i].putValue('CODEPOSTAL',
      GetCodePostal(fTobsPlanningTobItems.detail[i],
      fTobAffAdm,
      fTobsPlanningTobItems.detail[i].GetValue('APL_AFFAIRE'),
      fTobsPlanningTobItems.detail[i].GetValue('APL_TIERS'),
      fTobsPlanningTobItems.detail[i].GetString('ATA_TYPEADRESSE'),
      fTobsPlanningTobItems.detail[i].GetString('ATA_NUMEROADRESSE'),
      fTobsPlanningTobItems.detail[i].GetValue('APL_TYPEADRESSE'),
      fTobsPlanningTobItems.detail[i].GetValue('APL_NUMEROADRESSE'), True, True));

    FormatChampLibelle(fTobsPlanningTobItems.detail[i]);
    IconeCEGID(i);
  end;}

end;

function TPlanningGA6.LoadResPlanning(pStWhere: string): Boolean;
var
  Q: TQuery;
  S: string;
  vStWhere: string;

begin

  {S := ' SELECT DISTINCT AFF_LIBELLE, AFF_AFFAIRE, AFF_TIERS ';
  S := S + ' FROM AFFAIRE, PIECE ';
  S := S + ' WHERE AFF_AFFAIRE = GP_AFFAIRE ';
  S := S + ' AND GP_NATUREPIECEG = AFF_NATUREPIECEG';

  vStWhere := pStWhere;
  if pos('APL_AFFAIRE', vStWhere) <> 0 then
    ReplaceSubStr(vStWhere, 'APL_AFFAIRE', 'AFF_AFFAIRE');

  if vStWhere <> '' then
    S := S + vStWhere;
  S := S + ' ORDER BY AFF_AFFAIRE';

  Q := OpenSql(S, True);
  try
    if not Q.Eof then
    begin
      fTobsPlanningTobRes.ClearDetail;
      fTobsPlanningTobRes.LoadDetailDB('les_Ressources', '', '', Q, False, True);
      fTobsPlanningTobRes.detail[0].addchampSup('CLE', true);
      fTobsPlanningTobRes.detail[0].addchampSup('CHAMP1', true);
      formatTobRes;
      result := true;
    end
    else
      result := false;

  finally
    Ferme(Q);
  end;}
  
end;

procedure TPlanningGA6.AjoutTacheAbsence(pInIndex: Integer; pStRes, pStResLib: string);
begin
  // pas d'abscences dans ce planning
end;

procedure TPlanningGA6.AjoutNonProductive(pInIndex: Integer; pStRes, pStResLib: string);
begin
  // pas de tâches non productives dans ce planning
end;

procedure TPlanningGA6.LoadAbsencesPlanning(pStWhere: string);
begin
  // pas d'abscences dans ce planning
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 08/03/2002
Modifié le ... :
Description .. : formate l'affichage de la tob des ressources
Suite ........ :
Mots clefs ... :
*****************************************************************}
procedure TPlanningGA6.formatTobRes;
var
  i						: Integer;
  vStAffaire 	: String;

begin

  for i := 0 to fTobsPlanningTobRes.detail.count - 1 do
  begin
    // creation de la cle
    fTobsPlanningTobRes.detail[i].PutValue('CLE', fTobsPlanningTobRes.detail[i].GetValue('AFF_AFFAIRE'));

    // C.B 18/05/2006
    if GetParamsocSecur('SO_AFAFFICHAFFAIRE', False) then
	    vStAffaire := fTobsPlanningTobRes.detail[i].GetValue('AFF_LIBELLE')
    else
      vStAffaire := CodeAffaireAffiche(fTobsPlanningTobRes.detail[i].GetValue('AFF_AFFAIRE'));

    fTobsPlanningTobRes.detail[i].PutValue('CHAMP1', fTobsPlanningTobRes.detail[i].GetValue('AFF_TIERS') + #13#10 +
        RechDom('GCTIERSCLI', fTobsPlanningTobRes.detail[i].GetValue('AFF_TIERS'), False) + #13#10 +
        vStAffaire);
  end;

  // construction de l'arborescence
  // on n'affiche pas 2 fois la meme affaire
  for i := fTobsPlanningTobRes.detail.count - 1 downto 1 do
  begin
    if fTobsPlanningTobRes.detail[i].GetValue('CHAMP1') = fTobsPlanningTobRes.detail[i - 1].GetValue('CHAMP1') then
      fTobsPlanningTobRes.detail[i].PutValue('CHAMP1', '');
  end;
end;

// affaire
procedure TPlanningGA6.DecodeCle(pStCle: string);
var
  vSt: string;
begin
  vSt := pStCle;
  fStAffaire := (Trim(ReadTokenPipe(vSt, cStSep)));
  fStTiers := GetTiers(fStAffaire);
  fStTache := '';
  fStRes := '';
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 28/03/2002
Modifié le ... :
Description .. : création d'une ligne de planning et affichage à
Suite ........ : l'écran
Mots clefs ... :
*****************************************************************}
procedure TPlanningGA6.CreateItem(Sender: TObject; Item: TOB; var Cancel: boolean);
begin

  cancel := False;
  DecodeCle(item.getvalue('I_CODE'));
          
  // création directe = sélection d'une ligne effectuée
  if fInModeCreation = cInCreationDirecte then
  begin
    // sélection possible d'affaire
    if (fStAffaireSel <> '') and (fStAffaire <> fStAffaireSel) then
    begin
      // demande de crozier le 26/04/2005 de ne pas faire le test
      FFichePlanning.BT_EFFACELIGNESClick(self);
    end;
  end

  // creation sans selection de ressource
  else
    if (fInModeCreation = cInCreationAvecSelection) then
  begin
    if (fStAffaireSel <> '') and (fStAffaire <> fStAffaireSel) then
      Item.PutValue('APL_AFFAIRE', fStAffaire);
    Item.PutValue('APL_TIERS', fStTiers);
  end;                                 

  if not cancel then
  begin
    fTobCouranteMere := MaTobInterne('AFFAIRECOURANTE');
    fTobCourante := TOB.Create('AFFAIRECOURANTE', fTobCouranteMere, -1);
    fTobCourante.Dupliquer(Item, TRUE, TRUE, TRUE);
    inherited;
    FormatChampLibelle(Item);
  end;
end;

procedure TPlanningGA6.CopyItem(Sender: TObject; Item: TOB; var Cancel: boolean);
var
  vStRetour   : string;
  vStAffaire  : string;
  vSt         : string;

begin

  // C.B 26/09/2006
  // suppression du spécifique cegid
  {if IsStatutModifiableTob(Item.GetValue('APL_ETATLIGNE'), fTobsPlanningTobEtats) then
    Item.putValue('APL_ETATLIGNE', GetStatutItem(Item.GetString('ATA_STATUTPLA')));

  if not JaiLeDroitRessource(Item.getValue('APL_RESSOURCE')) then
  begin //mcd 14/03/2005
    PgiBoxAf(TexteMessage[24], TexteMessage[59]); //La ressource ne fait pas partie de votre équipe
    cancel := true;
  end

  // C.B 31/10/2006
  // test si on est a cheval sur 2 semaines ou 2 mois
  else if GetParamSocSecur('SO_AFPARSEMAINE', True) and
     EstAChevalSemaine(Item, fBoDemiJournee) then
  begin
    PGIBoxAF(TexteMessage[70], TexteMessage[59]); // Vous ne pouvez pas créer, déplacer, réduire ou étirer une intervention à cheval sur 2 semaines.
    cancel := True;
  end

  // C.B 06/04/2006
  // test si on est a cheval sur 2 semaines ou 2 mois
  else if GetParamSocSecur('SO_AFPARMOIS', True) and
     EstAChevalMois(Item, fBoDemiJournee) then
  begin
    PGIBoxAF(TexteMessage[71], TexteMessage[59]); // Vous ne pouvez pas créer, déplacer, réduire ou étirer une intervention à cheval sur 2 semaines.
    cancel := True;
  end

  // Recherche de l'existant dans la table tacheessource
  // test si on duplique sur le même emplacement
  else if ChangeLigneItemOk(Item) and
          CreateTacheRessource(Item, True) and
          AutoriserPlanif(Item, fBoDemiJournee, False, vStAffaire, 'DUPLICATION') then
  begin
    decodeCle(Item.GetValue('I_CODE'));
    Item.PutValue('APL_AFFAIRE', fStAffaire);
    fTobCouranteMere := MaTobInterne('AFFAIRECOURANTE');
    fTobCourante := TOB.Create('AFFAIRECOURANTE', fTobCouranteMere, -1);
    fTobCourante.Dupliquer(Item, TRUE, TRUE, TRUE);

    if fBoDemiJournee then
      vSt := ';DEMIJOURNEE'
    else
      vSt := '';
    vStRetour := AFLanceFicheAFPlanning('', 'ACTION=CREATION;DUPLIQUER' + vSt);

    if (vStRetour = 'CANCEL') or (vStRetour = '') then
      cancel := true
    else
    begin

      // C.B 31/10/2006
      // si on n'affiche pas les samedis et les dimanches et qu'on
      // est à cheval sur 2 semaines, on averti que les samedis et les
      // dimanches sont comptés
      if (not GetParamSocSecur('SO_AFPLANNINGWE', true)) and
         EstAChevalSemaine(Item, fBoDemiJournee) then
      begin
        PGIBoxAF(TexteMessage[20], TexteMessage[59]); // Attention, Les samedis et dimanches ne sont pas visibles, mais ils sont planifiés.
      end;

      Item.dupliquer(fTobCourante, True, True, True);
      RafraichirItem(Item, vStRetour, 'ARTICLE', False, True);
      cancel := false;
    end;
    DetruitMaTobInterne('AFFAIRECOURANTE');
  end
  else
    cancel := true;}
end;

procedure TPlanningGA6.ModifyItem(Sender: TObject; Item: TOB; var Cancel: boolean);
var
  vStItem: string;
  vStCle: string;

begin
  inherited;
  {if item <> nil then
  begin
    vStItem := AFLanceFicheAFPlanning(varAsType(Item.GetValue('APL_AFFAIRE'), varString) + ';' + varAsType(Item.GetValue('APL_NUMEROLIGNE'), varString), fStAction);

    if (vStItem = 'CANCEL') or (vStItem = '') then
      cancel := true
    else
    begin
      EncodeCle(Item.GetValue('APL_AFFAIRE'), '', '', vStCle);
      Item.putvalue('I_CODE', vStCle);
      RafraichirItem(Item, vStItem, 'ARTICLE', False, True);
      cancel := false;
    end;
    DetruitMaTobInterne('AFFAIRECOURANTE');
  end;}
end;

procedure TPlanningGA6.BParamClick;
begin
  //AFLanceFiche_PlanningParam('P1J');
end;

function TPlanningGA6.AfficherTiers : Boolean;
begin
  result := False;
end;

function TPlanningGA6.AfficherRess  : Boolean;
begin
  result := True;
end;

function TPlanningGA6.ModeAffichage : String;
begin
  result := 'ARTICLE';
end;

function TPlanningGA6.ChangeLigneItemOk(Item: TOB): boolean;
begin
  result := true;
  decodeCle(Item.GetValue('I_CODE'));

  if (Item.GetValue('APL_AFFAIRE') = '') then
  begin
    PGIBoxAF(TexteMessage[6], TexteMessage[59]); //Erreur au chargement de l''affaire.
    result := false;
  end
  else
    if (fStAffaire <> Item.GetValue('APL_AFFAIRE')) then
  begin
    PGIBoxAF(TexteMessage[5], TexteMessage[59]); //On ne peut pas changer l''affaire d''une intervention.
    result := false;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 12/04/2002
Modifié le ... :
Description .. : Mise a jour des items déplacés dans le planning
Suite ........ :
Mots clefs ... :
*****************************************************************}
function TPlanningGA6.UpdateItem(FromItem, ToItem: TOB; vStAction: string): Boolean;
var
  vBoOk: Boolean;
  vStRetour: string;

begin
 
  {vBoOk := True;
  if (vStAction <> 'POPUP') and
     (vStAction <> 'MOVE') then
  begin
    if PlusDeJours(FromItem, ToItem, fBoDemiJournee) then
	  begin
  	  vBoOk := (PGIAskAF(TexteMessage[13], '') = mrYes);
    	if vBoOk then vBoOk := (PGIAskAF(TexteMessage[14], '') = mrYes);
			if vBoOk and VH_GC.GCIfDefCEGID then
  	  	RevalorisationCegid(FromItem, ToItem, vStAction);
	  end
    // dans le cas de reduce, on execute systématiquement la revalorisation
    // car on est pas dans le cas plus de jours
    else if (vStAction = 'REDUCE') and VH_GC.GCIfDefCEGID then
    	RevalorisationCegid(FromItem, ToItem, vStAction);
	end;

  if not vBoOk then
    result := false
  else
  begin
    if ChangeLigneItemOk(ToItem) and CreateTacheRessource(ToItem, True) then
    begin
      decodeCle(ToItem.GetValue('I_CODE'));
      ToItem.PutValue('APL_UTILISATEUR', V_PGI.User);

      UpdateData(FromItem, ToItem);

      // C.B 30/11/2006 fonction commune
      result := UpdateAfPlanning(ToItem, vStRetour);
    end
    else
      result := false;
  end;}

end;


procedure TPlanningGA6.InitPlanning;
begin
  inherited;
  // les changements d'états sont disponibles en fonction des autorisations
  LoadPopUpMenu(False);
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 01/08/2002
Modifié le ... :   /  /
Description .. : Test si la Ressource existe dans la table tacheRessource
Mots clefs ... : Si non, on propose la creation
*****************************************************************}
function TPlanningGA6.CreateTacheRessource(Item: Tob; pBoForce: Boolean): boolean;
var
  vStAff0: string;
  vStAff1: string;
  vStAff2: string;
  vStAff3: string;
  vStAvenant: string;
begin
  result := false;
  decodeCle(Item.GetValue('I_CODE'));
  if not ExisteTacheRessource(iTem.GetString('APL_AFFAIRE'), intToStr(item.getValue('APL_NUMEROTACHE')), item.getValue('APL_RESSOURCE')) then
  begin
    if pBoForce or (PGIAskAF(TraduitGA(TexteMessage[7]), '') = mrYes) then  //Ce code ressource n est pas prévu pour cette tâche, voulez l''ajouter à la tâche ?', //mcd 14/03/05 mis au masculin pour Ok GI
    begin
      with Tob.Create('TACHERESSOURCE', nil, -1) do
      try
        PutValue('ATR_NUMEROTACHE', intToStr(item.getValue('APL_NUMEROTACHE')));
        PutValue('ATR_TIERS', item.GetValue('APL_TIERS'));
        PutValue('ATR_AFFAIRE', iTem.GetValue('APL_AFFAIRE'));
        CodeAffaireDecoupe(iTem.GetValue('APL_AFFAIRE'), vStAff0, vStAff1, vStAff2, vStAff3, vStAvenant, taModif, false);
        PutValue('ATR_AFFAIRE0', vStAff0);
        PutValue('ATR_AFFAIRE1', vStAff1);
        PutValue('ATR_AFFAIRE2', vStAff2);
        PutValue('ATR_AFFAIRE3', vStAff3);
        PutValue('ATR_AVENANT', vStAvenant);
        PutValue('ATR_RESSOURCE', item.getValue('APL_RESSOURCE'));
        PutValue('ATR_FONCTION', item.getValue('APL_FONCTION'));
        PutValue('ATR_STATUTRES', 'ACT');
        PutValue('ATR_DEVISE', V_PGI.DevisePivot);

        if not InsertDB(nil) then
          PGIBoxAF(TexteMessage[1], TexteMessage[59]) //Erreur lors de la mise à jour de l''intervention.
        else
          result := true;
      finally
        Free;
      end;
    end;
  end
  else
    result := true;
end;

{------------------------------------------------------------------------------}
{ OBJET TPlanningGA7 : Taches / ressource                                      }
{------------------------------------------------------------------------------}
{ CLE : RESSOURCE                                                              }
{------------------------------------------------------------------------------}
procedure TPlanningGA7.SelectPlanning;
begin

  fColonne1.StField := 'CHAMP1';
  fColonne2.StField := '';
  fColonne3.StField := '';

  fColonne1.StSize := '50';
  fColonne2.StSize := '50';
  fColonne3.StSize := '0';

  fColonne1.StAlign := 'L';
  fColonne2.StAlign := '';
  fColonne3.StAlign := '';

  fColonne1.StEntete := TraduitGA('Ressource');
  fColonne2.StEntete := '';
  fColonne3.StEntete := '';
end;

// chargement de toutes les lignes de planning et de toutes les ressources
procedure TPlanningGA7.LoadItemsPlanning(pStWhere: string);
var
  i: Integer;
  vStCle: string;

begin

  inherited;
  {for i := 0 to fTobsPlanningTobItems.detail.count - 1 do
  begin
    DisplayHint(fTobsPlanningTobItems.detail[i]);
    vStCle := fTobsPlanningTobItems.detail[i].GetValue('APL_RESSOURCE');
    fTobsPlanningTobItems.detail[i].putValue('I_CODE', vStCle);

    fTobsPlanningTobItems.detail[i].putValue('CODEPOSTAL',
      GetCodePostal(fTobsPlanningTobItems.detail[i],
      fTobAffAdm,
      fTobsPlanningTobItems.detail[i].GetValue('APL_AFFAIRE'),
      fTobsPlanningTobItems.detail[i].GetValue('APL_TIERS'),
      fTobsPlanningTobItems.detail[i].GetValue('ATA_TYPEADRESSE'),
      fTobsPlanningTobItems.detail[i].GetValue('ATA_NUMEROADRESSE'),
      fTobsPlanningTobItems.detail[i].GetValue('APL_TYPEADRESSE'),
      fTobsPlanningTobItems.detail[i].GetValue('APL_NUMEROADRESSE'), True, True));

    FormatChampLibelle(fTobsPlanningTobItems.detail[i]);
    IconeCEGID(i);
  end;}
end;

function TPlanningGA7.LoadResPlanning(pStWhere: string): Boolean;
var
  Q: TQuery;
  S: string;
  vStWhere: string;

begin

  {S := ' SELECT DISTINCT ARS_RESSOURCE AS ATR_RESSOURCE, ';
  S := S + ' ARS_LIBELLE || " " || ARS_LIBELLE2 AS CHAMP1 ';
  S := S + ' FROM RESSOURCE ';
  S := S + ' WHERE ARS_RESSOURCE <> "" ';

  vStWhere := pStWhere;

  if pos('APL_RESSOURCE', vStWhere) <> 0 then
    ReplaceSubStr(vStWhere, 'APL_RESSOURCE', 'ARS_RESSOURCE');

  if vStWhere <> '' then
    S := S + vStWhere;

  S := S + ' GROUP BY ARS_RESSOURCE, ARS_LIBELLE || " " || ARS_LIBELLE2';
  S := S + ' ORDER BY ARS_LIBELLE || " " || ARS_LIBELLE2';

  Q := nil;
  try
    Q := OpenSql(S, True);
    if not Q.Eof then
    begin
      fTobsPlanningTobRes.ClearDetail;
      fTobsPlanningTobRes.LoadDetailDB('les_Ressources', '', '', Q, False, True);
      fTobsPlanningTobRes.detail[0].addchampSup('CLE', true);
      // C.B 02/11/2004
      // si on gere les absences
      // on ne doit formater qu'une seule fois
      if not fBoGestionAbs then
        formatTobRes;
      result := true;
    end
    else
    begin
      result := false;
    end;
  finally
    if Q <> nil then
      Ferme(Q);
  end;}

end;

procedure TPlanningGA7.LoadAbsencesPlanning(pStWhere: string);
begin
  LoadAbsencesPlanningRessource(pStWhere);
end;

procedure TPlanningGA7.AjoutTacheAbsence(pInIndex: Integer; pStRes, pStResLib: string);
begin
  // pas de lignes a ajouter pour les abscences dans ce planning
end;

procedure TPlanningGA7.AjoutNonProductive(pInIndex: Integer; pStRes, pStResLib: string);
begin
  // pas de lignes à ajouter pour les tâches non productives dans ce planning
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 08/03/2002
Modifié le ... :
Description .. : formate l'affichage de la tob des ressources
Suite ........ :
Mots clefs ... :
*****************************************************************}
procedure TPlanningGA7.formatTobRes;
var
  i: integer;
  vStCle: string;
begin
  for i := 0 to fTobsPlanningTobRes.detail.count - 1 do
  begin
    // creation de la cle
    vStCle := fTobsPlanningTobRes.detail[i].GetValue('ATR_RESSOURCE');
    fTobsPlanningTobRes.detail[i].PutValue('CLE', vStCle);
		// attendre V7
		if not (pos(#13#10, fTobsPlanningTobRes.detail[i].GetValue('CHAMP1')) > 0) then
    	fTobsPlanningTobRes.detail[i].PutValue('CHAMP1', fTobsPlanningTobRes.detail[i].GetValue('CHAMP1') + #13#10 + fTobsPlanningTobRes.detail[i].GetValue('ATR_RESSOURCE'));
  end;
end;

// affaire, fonction, ressource
procedure TPlanningGA7.DecodeCle(pStCle: string);
var
  vSt: string;
begin
  vSt := pStCle;
  fStRes := (Trim(ReadTokenPipe(vSt, cStSep)));
  fStAffaire := '';
  fStTiers := '';
  fStTache := '';
end;

procedure TPlanningGA7.CreateItem(Sender: TObject; Item: TOB; var Cancel: boolean);
begin
  //
end;                            

procedure TPlanningGA7.CopyItem(Sender: TObject; Item: TOB; var Cancel: boolean);
begin
  cancel := True;
end;

//C.B 19/05/2006
// ajout de la consultation du planning dans les plannings
procedure TPlanningGA7.ModifyItem(Sender: TObject; Item: TOB; var Cancel: boolean);
var
  vStItem: string;
  vStCle: string;

begin
  inherited;
  if item <> nil then
  begin
    vStItem := AFLanceFicheAFPlanning (varAsType(Item.GetValue('APL_AFFAIRE'), varString) + ';' + varAsType(Item.GetValue('APL_NUMEROLIGNE'), varString), 'ACTION=CONSULTATION');
    if (vStItem = 'CANCEL') or (vStItem = '') then
      cancel := true
    else
    begin
      EncodeCle(Item.GetValue('APL_AFFAIRE'), '', '', vStCle);
      Item.putvalue('I_CODE', vStCle);
      RafraichirItem(Item, vStItem, 'ARTICLE', False, True);
      cancel := false;
    end;
    DetruitMaTobInterne ('AFFAIRECOURANTE');
  end;
end;

procedure TPlanningGA7.BParamClick;
begin
  //AFLanceFiche_PlanningParam('P2J');
end;

function TPlanningGA7.AfficherTiers : Boolean;
begin
  result := True;
end;

function TPlanningGA7.AfficherRess  : Boolean;
begin
  result := False;
end;

function TPlanningGA7.ModeAffichage : String;
begin
 result := 'ARTICLE';
end;
                                 
{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 12/04/2002
Modifié le ... :
Description .. : Mise a jour des items déplacés dans le planning
Suite ........ :
Mots clefs ... :
*****************************************************************}
function TPlanningGA7.UpdateItem(FromItem, ToItem: TOB; vStAction: string): Boolean;
var
  vBoOk: Boolean;
  vStRetour: string;

begin

  {vBoOk := True;
  if (vStAction <> 'POPUP') and
     (vStAction <> 'MOVE') then
  begin
    if PlusDeJours(FromItem, ToItem, fBoDemiJournee) then
	  begin
  	  vBoOk := (PGIAskAF(TexteMessage[13], '') = mrYes);
    	if vBoOk then vBoOk := (PGIAskAF(TexteMessage[14], '') = mrYes);
			if vBoOk and VH_GC.GCIfDefCEGID then
  	  	RevalorisationCegid(FromItem, ToItem, vStAction);
	  end
    // dans le cas de reduce, on execute systématiquement la revalorisation
    // car on est pas dans le cas plus de jours
    else if (vStAction = 'REDUCE') and VH_GC.GCIfDefCEGID then
    	RevalorisationCegid(FromItem, ToItem, vStAction);
	end;

  if not vBoOk then
    result := false
  else
  begin
    // en copy, seul ToItem est renseigné
    if not assigned(FromItem) and Assigned(ToItem) then
      fStAffaire := ToItem.GetValue('APL_AFFAIRE')
    else
      fStAffaire := FromItem.GetValue('APL_AFFAIRE');

    //if ChangeLigneItemOk(ToItem) and CreateTacheRessource(ToItem, True) then
    if CreateTacheRessource(ToItem, True) then
    begin
      DecodeCle(ToItem.GetValue('I_CODE'));
      ToItem.PutValue('APL_RESSOURCE', fStRes);
      ToItem.PutValue('APL_UTILISATEUR', V_PGI.User);

      UpdateData(FromItem, ToItem);

      // C.B 30/11/2006 fonction commune
      result := UpdateAfPlanning(ToItem, vStRetour);
    end
    else
      result := false;
  end;}
end;

function TPlanningGA7.ChangeLigneItemOk(Item: TOB): boolean;
begin
  result := true;
end;

function TPlanningGA7.CreateTacheRessource(Item: Tob; pBoForce: Boolean): boolean;
begin
  result := True;
end;

procedure TPlanningGA7.InitPlanning;
begin
  // forcé en consultation
  fAction := TaConsult;
  inherited;
  // les changements d'états sont disponibles en fonction des autorisations
  LoadPopUpMenu(True);
end;

end.
